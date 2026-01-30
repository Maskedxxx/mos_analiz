# word_audit

Сервис LLM-аудита Word-документов (приказы/положения) по ТЗ и шаблону.

## Вход
`POST /jobs` (multipart):
- `job_type`: `order_comp_ppu | order_ppu | reg_comp_ppu | reg_ppu | package`

Для `job_type != package`:
- `target_doc` (docx)
- `template_doc` (docx)
- `tz_doc` (docx)

Для `job_type=package`:
- `order_comp_doc`, `reg_comp_doc`, `order_ppu_doc`, `reg_ppu_doc` (docx)

## Выход (артефакты в run‑папке)
- `violations.json`
- `report.xlsx`
- `debug_parsed.json`
- `artifacts.zip`

## Правила/настройки
- Правила проверки берутся из загруженного `tz_doc` (ТЗ).
- Маппинг SCOPE/COMPARE хранится в `rules/ruleset_mapping.json`.

