#!/usr/bin/env python3
# coding: utf-8
"""
Автор: technoconsulx
AI Mail Processor - Умный обработчик входящей почты с поддержкой AI

Возможности:
- Поддержка форматов: txt, pdf, docx, doc, xlsx, eml, изображения (jpg/png/webp)
- AI-модель: Ollama с поддержкой анализа текста и изображений
- IMAP/SMTP интеграция
- Автоматический поиск контекста через DuckDuckGo
"""

import asyncio
import base64
import email
import json
import logging
import os
import re
import sys
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.mime.text import MIMEText
import imaplib
import requests
from aiosmtplib import SMTP
from bs4 import BeautifulSoup
import tempfile

# Try imports for extractors
try:
    import PyPDF2
    from PyPDF2 import PdfReader
except ImportError:
    PyPDF2 = None

try:
    import fitz  # PyMuPDF - для извлечения текста и изображений из PDF
except ImportError:
    fitz = None

try:
    from pdf2image import convert_from_path  # для конвертации PDF в изображения
except ImportError:
    convert_from_path = None

try:
    import docx as python_docx
except ImportError:
    python_docx = None

try:
    import textract
except ImportError:
    textract = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    from PIL import Image
    import io
except ImportError:
    Image = None
    io = None


class Config:
    """Конфигурация приложения"""
    
    # IMAP settings
    IMAP_HOST = os.getenv("IMAP_HOST", "mail.example.com")
    IMAP_USER = os.getenv("IMAP_USER", "ai@example.com")
    IMAP_PASS = os.getenv("IMAP_PASS", "password")
    
    # SMTP settings
    SMTP_HOST = os.getenv("SMTP_HOST", "mail.example.com")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "25"))
    SMTP_USER = os.getenv("SMTP_USER", "ai@example.com")
    SMTP_PASS = os.getenv("SMTP_PASS", "password")
    
    # Processing settings
    CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL", "30"))  # seconds
    MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024  # 10 MB
    
    # Directories
    TELEGRAMS_DIR = os.getenv("TELEGRAMS_DIR", "telegrams")
    ERROR_DIR = os.getenv("ERROR_DIR", "ErrorTelegrams")
    PROCESSED_DIR = os.getenv("PROCESSED_DIR", "ProcessedTelegrams")
    
    # AI settings
    OLLAMA_URL = os.getenv("OLLAMA_URL", "http://localhost:11434/api/generate")
    OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "gemma3:12b")
    OLLAMA_TIMEOUT = int(os.getenv("OLLAMA_TIMEOUT", "9999"))  # seconds
    
    # Search settings
    DUCKDUCKGO_MAX = int(os.getenv("DUCKDUCKGO_MAX", "3"))  # number of search results
    
    # PDF processing settings
    PDF_EXTRACT_IMAGES = os.getenv("PDF_EXTRACT_IMAGES", "true").lower() == "true"
    PDF_EXTRACT_TEXT = os.getenv("PDF_EXTRACT_TEXT", "true").lower() == "true"
    PDF_CONVERT_TO_IMAGES = os.getenv("PDF_CONVERT_TO_IMAGES", "false").lower() == "true"


class MailAIProcessor:
    """Основной класс обработчика почты с AI"""
    
    def __init__(self, config: Config):
        self.config = config
        self.setup_logging()
        self.create_directories()
        
    def setup_logging(self):
        """Настройка логирования"""
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(levelname)-7s | %(name)-12s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        self.logger = logging.getLogger("MailAIProcessor")
        
    def create_directories(self):
        """Создание необходимых директорий"""
        os.makedirs(self.config.TELEGRAMS_DIR, exist_ok=True)
        os.makedirs(self.config.ERROR_DIR, exist_ok=True)
        os.makedirs(self.config.PROCESSED_DIR, exist_ok=True)
        self.logger.info("Directories created/verified")

    @staticmethod
    def safe_filename(name: str) -> str:
        """Создание безопасного имени файла"""
        return re.sub(r"[^\w\-. ]", "_", name)

    def save_bytes_to_file(self, dirname: str, filename: str, data: bytes) -> str:
        """Сохранение данных в файл с проверкой размера"""
        if len(data) > self.config.MAX_ATTACHMENT_SIZE:
            self.logger.warning(f"Skipping {filename}: exceeds size limit")
            return None
        
        fname = self.safe_filename(filename)
        path = os.path.join(dirname, fname)
        base, ext = os.path.splitext(path)
        counter = 1
        while os.path.exists(path):
            path = f"{base}_{counter}{ext}"
            counter += 1
        
        try:
            with open(path, "wb") as f:
                f.write(data)
            return path
        except Exception as e:
            self.logger.error(f"Failed to save file {filename}: {e}")
            return None

    def extract_pdf_text_pypdf2(self, pdf_path: str) -> str:
        """Извлечение текста из PDF с помощью PyPDF2"""
        try:
            with open(pdf_path, "rb") as f:
                reader = PdfReader(f)
                pages = []
                for page in reader.pages:
                    try:
                        page_text = page.extract_text() or ""
                        pages.append(page_text)
                    except Exception as e:
                        self.logger.warning(f"PyPDF2 page extraction failed: {e}")
                        pages.append("")
                return "\n".join(pages)
        except Exception as e:
            self.logger.warning(f"PyPDF2 extraction failed for {pdf_path}: {e}")
            return ""

    def extract_pdf_text_images_pymupdf(self, pdf_path: str) -> tuple:
        """
        Извлечение текста и изображений из PDF с помощью PyMuPDF
        Возвращает (текст, список путей к извлеченным изображениям)
        """
        text = ""
        extracted_images = []
        
        try:
            doc = fitz.open(pdf_path)
            
            # Извлечение текста
            if self.config.PDF_EXTRACT_TEXT:
                text_pages = []
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    page_text = page.get_text()
                    text_pages.append(f"--- Page {page_num + 1} ---\n{page_text}")
                text = "\n\n".join(text_pages)
            
            # Извлечение изображений
            if self.config.PDF_EXTRACT_IMAGES:
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    image_list = page.get_images()
                    
                    for img_index, img in enumerate(image_list):
                        try:
                            xref = img[0]
                            base_image = doc.extract_image(xref)
                            image_bytes = base_image["image"]
                            image_ext = base_image["ext"]
                            
                            # Сохраняем изображение
                            img_filename = f"pdf_image_p{page_num+1}_i{img_index+1}.{image_ext}"
                            img_path = self.save_bytes_to_file(
                                self.config.TELEGRAMS_DIR, img_filename, image_bytes
                            )
                            if img_path:
                                extracted_images.append(img_path)
                                self.logger.info(f"Extracted image from PDF: {img_path}")
                                
                        except Exception as e:
                            self.logger.warning(f"Failed to extract image from PDF page {page_num}: {e}")
            
            doc.close()
            return text, extracted_images
            
        except Exception as e:
            self.logger.error(f"PyMuPDF extraction failed for {pdf_path}: {e}")
            return "", []

    def convert_pdf_to_images(self, pdf_path: str) -> list:
        """Конвертация PDF страниц в изображения"""
        if not convert_from_path:
            self.logger.warning("pdf2image not available for PDF conversion")
            return []
        
        try:
            images = convert_from_path(pdf_path, dpi=150)
            image_paths = []
            
            for i, image in enumerate(images):
                # Сохраняем как временный файл
                with tempfile.NamedTemporaryFile(
                    suffix=f"_page_{i+1}.jpg", 
                    delete=False, 
                    dir=self.config.TELEGRAMS_DIR
                ) as tmp_file:
                    image.save(tmp_file.name, "JPEG", quality=85)
                    image_paths.append(tmp_file.name)
                    self.logger.info(f"Converted PDF page {i+1} to image: {tmp_file.name}")
            
            return image_paths
            
        except Exception as e:
            self.logger.error(f"PDF to image conversion failed for {pdf_path}: {e}")
            return []

    def process_pdf_comprehensive(self, pdf_path: str) -> tuple:
        """
        Комплексная обработка PDF:
        - Извлечение текста (PyMuPDF + PyPDF2 как fallback)
        - Извлечение встроенных изображений
        - Опциональная конвертация страниц в изображения
        """
        all_text = ""
        all_images = []
        
        # 1. Извлечение текста и изображений с помощью PyMuPDF (если доступен)
        if fitz:
            text, images = self.extract_pdf_text_images_pymupdf(pdf_path)
            all_text += text
            all_images.extend(images)
        else:
            self.logger.warning("PyMuPDF not available, using PyPDF2 for text extraction only")
        
        # 2. Fallback: извлечение текста с помощью PyPDF2
        if PyPDF2 and (not all_text.strip() or self.config.PDF_EXTRACT_TEXT):
            pdf2_text = self.extract_pdf_text_pypdf2(pdf_path)
            if pdf2_text.strip():
                all_text += f"\n\n--- PyPDF2 Extraction ---\n{pdf2_text}"
        
        # 3. Конвертация PDF в изображения (если включено)
        if self.config.PDF_CONVERT_TO_IMAGES:
            converted_images = self.convert_pdf_to_images(pdf_path)
            all_images.extend(converted_images)
        
        return all_text.strip(), all_images

    def extract_attachments_and_text(self, msg) -> tuple:
        """Извлечение вложений и текста из email сообщения"""
        attachments = []
        text_parts = []
        image_paths = []
        
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                disp = str(part.get("Content-Disposition") or "")
                
                if ctype == "text/plain" and "attachment" not in disp:
                    try:
                        raw = part.get_payload(decode=True)
                        if raw:
                            charset = part.get_content_charset() or "utf-8"
                            text = raw.decode(charset, errors="ignore")
                            text_parts.append(text)
                    except Exception as e:
                        self.logger.warning(f"Failed to decode text/plain part: {e}")
                
                filename = part.get_filename()
                if filename:
                    try:
                        payload = part.get_payload(decode=True) or b""
                        path = self.save_bytes_to_file(
                            self.config.TELEGRAMS_DIR, filename, payload
                        )
                        if path:
                            attachments.append(path)
                            self.logger.info(f"Saved attachment: {path}")
                    except Exception as e:
                        self.logger.error(f"Failed to save attachment {filename}: {e}")
        else:
            try:
                raw = msg.get_payload(decode=True)
                if raw:
                    charset = msg.get_content_charset() or "utf-8"
                    try:
                        text_parts.append(raw.decode(charset, errors="ignore"))
                    except Exception:
                        text_parts.append(raw.decode("utf-8", errors="ignore"))
            except Exception as e:
                self.logger.warning(f"Failed to extract singlepart body: {e}")

        # Извлечение текста из поддерживаемых вложений
        extracted_texts = []
        extracted_images = []
        
        for path in attachments:
            if not path:
                continue
            lower = path.lower()
            try:
                if lower.endswith(".pdf"):
                    # Комплексная обработка PDF
                    pdf_text, pdf_images = self.process_pdf_comprehensive(path)
                    if pdf_text:
                        extracted_texts.append(pdf_text)
                        self.logger.info(f"Extracted text from PDF: {path} (length: {len(pdf_text)})")
                    if pdf_images:
                        extracted_images.extend(pdf_images)
                        self.logger.info(f"Extracted {len(pdf_images)} images from PDF: {path}")
                
                elif lower.endswith(".docx") and python_docx:
                    try:
                        doc = python_docx.Document(path)
                        texts = [p.text for p in doc.paragraphs if p.text]
                        extracted_texts.append("\n".join(texts))
                        self.logger.info(f"Extracted text from DOCX: {path}")
                    except Exception as e:
                        self.logger.warning(f"DOCX extraction failed for {path}: {e}")
                
                elif lower.endswith(".doc"):
                    if textract:
                        try:
                            raw = textract.process(path)
                            extracted_texts.append(raw.decode("utf-8", errors="ignore"))
                            self.logger.info(f"Extracted text from DOC via textract: {path}")
                        except Exception as e:
                            self.logger.warning(f"textract failed for {path}: {e}")
                    else:
                        self.logger.warning(f"Skipping .doc extraction (textract not available): {path}")
                
                elif lower.endswith(".txt"):
                    try:
                        with open(path, "r", encoding="utf-8", errors="ignore") as f:
                            extracted_texts.append(f.read())
                        self.logger.info(f"Read txt file: {path}")
                    except Exception as e:
                        self.logger.warning(f"TXT read failed: {e}")
                
                elif lower.endswith(".eml"):
                    try:
                        with open(path, "rb") as f:
                            em = BytesParser(policy=policy.default).parsebytes(f.read())
                        parts = []
                        if em.is_multipart():
                            for p in em.walk():
                                if p.get_content_type() == "text/plain":
                                    payload = p.get_payload(decode=True)
                                    if payload:
                                        parts.append(payload.decode(
                                            p.get_content_charset() or "utf-8", errors="ignore"
                                        ))
                        else:
                            payload = em.get_payload(decode=True)
                            if payload:
                                parts.append(payload.decode(
                                    em.get_content_charset() or "utf-8", errors="ignore"
                                ))
                        extracted_texts.append("\n".join(parts))
                        self.logger.info(f"Extracted text from EML: {path}")
                    except Exception as e:
                        self.logger.warning(f"EML extraction failed: {e}")
                
                elif lower.endswith(".xlsx") and openpyxl:
                    try:
                        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                        sheets_text = []
                        for sheet in wb.worksheets:
                            rows = []
                            for row in sheet.iter_rows(values_only=True):
                                cells = [str(c) for c in row if c is not None]
                                if cells:
                                    rows.append("\t".join(cells))
                            if rows:
                                sheets_text.append("\n".join(rows))
                        extracted_texts.append("\n\n".join(sheets_text))
                        self.logger.info(f"Extracted text from XLSX: {path}")
                    except Exception as e:
                        self.logger.warning(f"XLSX extraction failed for {path}: {e}")
                
                # Проверяем, является ли файл изображением
                elif lower.endswith((".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif")):
                    image_paths.append(path)
                    self.logger.info(f"Found image attachment: {path}")
                    
                else:
                    self.logger.debug(f"No extractor for {path}")
                    
            except Exception as e:
                self.logger.warning(f"Error extracting from {path}: {e}")

        # Добавляем извлеченные изображения из PDF к общему списку
        image_paths.extend(extracted_images)
        
        combined_text = "\n\n".join([t for t in (text_parts + extracted_texts) if t])
        return attachments, combined_text, image_paths

    def duckduckgo_search(self, query: str, max_results: int = None) -> list:
        """Поиск через DuckDuckGo для дополнительного контекста"""
        if max_results is None:
            max_results = self.config.DUCKDUCKGO_MAX
            
        try:
            params = {"q": query}
            headers = {"User-Agent": "Mozilla/5.0 (compatible; bot/1.0)"}
            r = requests.get("https://duckduckgo.com/html/", params=params, headers=headers, timeout=15)
            r.raise_for_status()
            
            soup = BeautifulSoup(r.text, "lxml")
            results = []
            
            for a in soup.select("a.result__a")[:max_results]:
                title = a.get_text(strip=True)
                href = a.get("href")
                parent = a.find_parent("div", class_="result")
                snippet = ""
                if parent:
                    s = parent.select_one(".result__snippet")
                    if s:
                        snippet = s.get_text(strip=True)
                results.append((title, href, snippet))
            
            if not results:
                for rdiv in soup.select(".result")[:max_results]:
                    a = rdiv.find("a", href=True)
                    if a:
                        title = a.get_text(strip=True)
                        href = a["href"]
                        s = rdiv.find(class_="result__snippet")
                        snippet = s.get_text(strip=True) if s else ""
                        results.append((title, href, snippet))
            
            return results[:max_results]
        except Exception as e:
            self.logger.warning(f"DuckDuckGo search failed: {e}")
            return []

    def call_ollama(self, prompt: str, image_paths: list = None, extra_ctx: str = None):
        """Вызов Ollama AI модели"""
        images_b64 = []
        if image_paths:
            for p in image_paths:
                try:
                    with open(p, "rb") as f:
                        images_b64.append(base64.b64encode(f.read()).decode("utf-8"))
                except Exception as e:
                    self.logger.warning(f"Failed to encode image {p}: {e}")
                    
        system_prompt = """Ты - дружелюбный AI-ассистент, который отвечает на входящие письма. 
Отвечай естественно и по делу, как живой человек в деловой переписке. 
Будь вежливым, но не слишком формальным. Отвечай на вопросы прямо и помогай решать проблемы."""
        
        full_prompt = (extra_ctx + "\n\n" if extra_ctx else "") + (prompt or "")
        payload = {
            "model": self.config.OLLAMA_MODEL, 
            "prompt": full_prompt,
            "stream": False
        }
        
        if images_b64:
            payload["images"] = images_b64
        
        try:
            self.logger.info(f"Calling Ollama model={self.config.OLLAMA_MODEL} images={len(images_b64)}")
            resp = requests.post(self.config.OLLAMA_URL, json=payload, timeout=self.config.OLLAMA_TIMEOUT)
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Ollama request failed: {e}")
            return None

    async def send_reply(self, to_addr: str, subject: str, body: str):
        """Отправка ответа через SMTP"""
        try:
            msg = MIMEText(body, "plain", "utf-8")
            msg["Subject"] = f"Re: {subject or 'Без темы'}"
            msg["From"] = self.config.SMTP_USER
            msg["To"] = to_addr
            msg["Reply-To"] = self.config.SMTP_USER

            smtp = SMTP(
                hostname=self.config.SMTP_HOST,
                port=self.config.SMTP_PORT,
                use_tls=False,
                start_tls=False,
                tls_context=None
            )

            await smtp.connect()
            await smtp.login(self.config.SMTP_USER, self.config.SMTP_PASS)
            await smtp.send_message(msg)
            await smtp.quit()
            
            self.logger.info(f"Reply sent to {to_addr}")
            
        except Exception as e:
            self.logger.error(f"send_reply failed: {e}")
            self.logger.error(f"SMTP details: host={self.config.SMTP_HOST}, port={self.config.SMTP_PORT}, user={self.config.SMTP_USER}, to={to_addr}")
            raise

    def fetch_unseen_uids(self, mail) -> list:
        """Получение UID непрочитанных сообщений"""
        typ, data = mail.search(None, "UNSEEN")
        if typ != "OK":
            self.logger.warning("IMAP search failed")
            return []
        return data[0].split()

    def fetch_full_message(self, mail, uid_bytes):
        """Получение полного сообщения по UID"""
        typ, data = mail.fetch(uid_bytes, "(RFC822)")
        if typ != "OK":
            raise RuntimeError("IMAP fetch failed")
        return data[0][1]

    async def handle_message(self, mail, uid_bytes):
        """Обработка одного сообщения"""
        raw = self.fetch_full_message(mail, uid_bytes)
        msg = BytesParser(policy=policy.default).parsebytes(raw)
        
        subject = msg.get("subject", "(no subject)")
        from_header = msg.get("from", "")
        
        from_addr = ""
        if from_header:
            email_match = re.search(r'<([^>]+)>', from_header)
            if email_match:
                from_addr = email_match.group(1)
            else:
                from_addr = from_header.strip()
        
        if not from_addr:
            self.logger.error(f"Cannot extract email from: {from_header}")
            return
        
        uid_str = uid_bytes.decode() if isinstance(uid_bytes, bytes) else str(uid_bytes)
        
        self.logger.info(f"Handling msg uid={uid_str} subject={subject} from={from_addr}")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        raw_path = os.path.join(self.config.TELEGRAMS_DIR, f"raw_{timestamp}_{uid_str}.eml")
        
        try:
            with open(raw_path, "wb") as f:
                f.write(raw)
        except Exception as e:
            self.logger.warning(f"Failed to save raw email: {e}")
        
        attachments, combined_text, image_paths = self.extract_attachments_and_text(msg)
        self.logger.info(f"Found {len(attachments)} attachments; extracted_text_len={len(combined_text)}; images={len(image_paths)}")
        
        extra_ctx = ""
        if not combined_text or len(combined_text.strip()) < 80:
            q_parts = [subject or "", from_addr or ""]
            if attachments:
                q_parts.append(os.path.basename(attachments[0]))
            query = " ".join([p for p in q_parts if p]).strip()
            
            if query:
                dd = self.duckduckgo_search(query, max_results=self.config.DUCKDUCKGO_MAX)
                if dd:
                    lines = ["DuckDuckGo quick search results:"]
                    for title, url, snip in dd:
                        lines.append(f"- {title} | {url}\n  {snip}")
                    extra_ctx = "\n".join(lines)
                    self.logger.info(f"Added DuckDuckGo context ({len(dd)} results) for uid={uid_str}")
        
        filenames = [os.path.basename(p) for p in attachments if p]
        prompt_parts = []
        
        if combined_text:
            prompt_parts.append("Письмо / извлечённый текст:\n" + combined_text)
        if filenames:
            prompt_parts.append("Вложенные файлы: " + ", ".join(filenames))
        
        prompt = "\n\n".join(prompt_parts) if prompt_parts else "(No text provided)"
        
        response = self.call_ollama(prompt, image_paths=image_paths, extra_ctx=extra_ctx)
        
        response_path = os.path.join(self.config.TELEGRAMS_DIR, f"response_{timestamp}_{uid_str}.txt")
        try:
            with open(response_path, "w", encoding="utf-8") as f:
                f.write(json.dumps(response, ensure_ascii=False, indent=2) if isinstance(response, dict) else str(response))
            self.logger.info(f"Saved response to {response_path}")
        except Exception as e:
            self.logger.error(f"Failed to save response: {e}")
        
        body = ""
        if isinstance(response, dict):
            for k in ("response", "generated_text", "text", "output", "result"):
                if k in response and isinstance(response[k], str) and response[k].strip():
                    body = response[k]
                    break
            if not body:
                body = json.dumps(response, ensure_ascii=False, indent=2)
        else:
            body = str(response)
        
        if not body:
            body = "(No response from model)"
        
        try:
            await self.send_reply(from_addr, subject, body)
            try:
                processed_path = os.path.join(self.config.PROCESSED_DIR, os.path.basename(raw_path))
                os.replace(raw_path, processed_path)
            except Exception:
                pass
        except Exception as e:
            self.logger.error(f"Failed to send reply for uid={uid_str}: {e}")
            err_path = os.path.join(self.config.ERROR_DIR, f"error_{timestamp}_{uid_str}.eml")
            try:
                with open(err_path, "wb") as f:
                    f.write(raw)
                self.logger.error(f"Moved problematic email to {err_path}")
            except Exception as e2:
                self.logger.error(f"Also failed to save error eml: {e2}")

    async def main_loop(self):
        """Главный цикл обработки"""
        self.logger.info("Starting Mail AI Processor main loop")
        
        # Проверяем доступность библиотек для PDF
        if not fitz:
            self.logger.warning("PyMuPDF (fitz) not available - limited PDF image extraction")
        if not convert_from_path:
            self.logger.warning("pdf2image not available - PDF to image conversion disabled")
        
        while True:
            try:
                self.logger.info("Connecting to IMAP...")
                mail = imaplib.IMAP4_SSL(self.config.IMAP_HOST)
                mail.login(self.config.IMAP_USER, self.config.IMAP_PASS)
                mail.select("inbox")
                
                uids = self.fetch_unseen_uids(mail)
                self.logger.info(f"Found {len(uids)} new messages")
                
                for uid in uids:
                    try:
                        await self.handle_message(mail, uid)
                    except Exception as e:
                        self.logger.error(f"Error handling uid {uid}: {e}")
                
                mail.logout()
                self.logger.info("IMAP logout")
                
            except Exception as e:
                self.logger.error(f"Main loop error: {e}")
            
            self.logger.info(f"Sleeping {self.config.CHECK_INTERVAL} seconds until next check")
            await asyncio.sleep(self.config.CHECK_INTERVAL)


async def main():
    """Главная функция"""
    config = Config()
    processor = MailAIProcessor(config)
    
    try:
        await processor.main_loop()
    except KeyboardInterrupt:
        processor.logger.info("Interrupted by user, exiting")
    except Exception as e:
        processor.logger.exception(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())


