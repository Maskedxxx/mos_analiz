# drivers_excel

Сервис проверки драйверов (Excel → LLM → Excel‑отчёт).

## Вход
- `.xlsx` файл с драйверами (multipart поле `file`).

## Выход (артефакты в run‑папке)
- `01_parsed.json`
- `02_llm_analysis.json`
- `03_report.xlsx`
- `artifacts.zip`

## API
- `GET /health`
- `POST /jobs`
- `GET /jobs/{job_id}`
- `GET /jobs/{job_id}/artifacts`
- `GET /jobs/{job_id}/artifacts/{name}`

## Правила/настройки
- Промпт проверки: `rules/driver_check_prompt.txt`
- Дефолтные пороги/параметры модели: `config/config.yaml` (секреты не хранятся)

