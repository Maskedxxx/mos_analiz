# Мос Мониторинг — Project Instructions

## API ключ
source .env

## Python venv
# Создание (один раз):
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

# Активация:
source .venv/bin/activate

## Запуск
source .venv/bin/activate && source .env
python run_audit.py --doc-type <type> --target <file>
python run_audit.py --list-types
