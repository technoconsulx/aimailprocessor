# aimailprocessor
Обрабатывает ваши письма по электронной почте. - Поддержка форматов: txt, pdf, docx, doc, xlsx, eml, изображения (jpg/png/webp) - AI-модель: Ollama с поддержкой анализа текста и изображений - IMAP/SMTP интеграция - Автоматический поиск контекста через DuckDuckGo

Изменяя systemprompt и настройки модели под ваши требования, вы можете создать AI ассистента или агента техподдержки

Установка Linux:
1. sudo apt update && sudo apt upgrade -y
2. sudo apt update
sudo apt install -y fonts-dejavu-core fonts-liberation fonts-freefont-ttf python3-dev build-essential libssl-dev libffi-dev poppler-utils libxml2-dev libxslt1-dev antiword python3-pip
3. python3 -m venv ~/mailai_venv
4. source ~/mailai_venv/bin/activate
5. git clone https://github.com/technoconsulx/aimailprocessor.git
6. cd aimailprocessor
7. pip install --upgrade pip
8. pip install -r requirements.txt
9. curl -fsSL https://ollama.com/install.sh | sh
10. ollama serve
11. ollama pull gemma3:12b
12. Измените конфигурацию почты в коде или env.

