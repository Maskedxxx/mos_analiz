# kpsc_validation

Сервис валидации КПСЦ (Excel) по набору правил.

Пайплайн:
1) Парсинг нужных листов Excel в `parser_outputs/*.json`
2) Запуск валидаторов (`validation_scripts/`) и сбор `validation_report.xlsx`

## Вход
`POST /jobs` (multipart):
- `file` — `.xlsx` КПСЦ
- `rules_file` — опционально `.json` с правилами (если не задан, используется встроенный `kpsc_validation/validation_rules.json`)

## Выход (артефакты в run‑папке)
- `parser_outputs/` (директория)
- `validation_outputs/` (директория)
- `validation_report.xlsx`
- `artifacts.zip`

## Правила
- По умолчанию: `kpsc_validation/validation_rules.json`
- Для override: передать `rules_file`, сервис установит `VALIDATION_RULES_PATH` для валидаторов.

