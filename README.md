# aimailprocessor
Обрабатывает ваши письма по электронной почте. - Поддержка форматов: txt, pdf, docx, doc, xlsx, eml, изображения (jpg/png/webp) - AI-модель: Ollama с поддержкой анализа текста и изображений - IMAP/SMTP интеграция - Автоматический поиск контекста через DuckDuckGo

Изменяя systemprompt и нстройки модели под ваши требования, вы можете создать AI ассистента или агента техподдержки

Установка Linux:
1. Подготовка системы
Установка зависимостей
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3 python3-pip python3-venv git wget build-essential libpoppler-dev libxml2-dev libxslt1-dev zlib1g-dev
Создание виртуального окружения
python3 -m venv ~/mailai_venv
source ~/mailai_venv/bin/activate

2. Установка Python-пакетов
Клонирование репозитория
git clone https://github.com/technoconsulx/aimailprocessor.git
cd ai-mail-processor
Установка основных зависимостей
pip install --upgrade pip
pip install -r requirements.txt
3. Установка Ollama
curl -fsSL https://ollama.com/install.sh | sh
ollama serve
ollama pull gemma3:12b
4. Измените конфигурацию почты в коде или env.
