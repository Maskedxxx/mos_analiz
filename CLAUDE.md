# Мос Мониторинг — Project Instructions

## OpenAI API Key
# Ключ хранится локально в переменной окружения, не коммитится
# export OPENAI_API_KEY="ваш_ключ"

## Python venv
# Пока используем общий semantic_venv:
source /Users/mask/Documents/ПРОЕКТЫ_2024/СОЮЗ_СНАБ_workRepo/knowledge_map_release_v2/ai-neuro/semantic_venv/bin/activate

# Собственный .venv/ будет создан позже под прод:
# python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

## Запуск
source /Users/mask/Documents/ПРОЕКТЫ_2024/СОЮЗ_СНАБ_workRepo/knowledge_map_release_v2/ai-neuro/semantic_venv/bin/activate
export OPENAI_API_KEY="$OPENAI_API_KEY"
python run_audit.py --doc-type <type> --target <file>
python run_audit.py --list-types
