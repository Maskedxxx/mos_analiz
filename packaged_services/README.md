# Packaged Services (Docker)

Этот каталог содержит 3 автономных сервиса (по одному на каждую задачу), упакованных в Docker и запускаемых через `docker compose`.

## Быстрый старт

1) Создайте файл `packaged_services/.env` (можно с копированием из примера):
```bash
cp packaged_services/.env.example packaged_services/.env
```
и задайте `OPENAI_API_KEY`.

2) Запуск:
```bash
cd packaged_services
docker compose up --build
```

## Сервисы и порты

- `drivers_excel` (проверка драйверов Excel): http://localhost:8000
- `word_audit` (аудит Word документов): http://localhost:8001
- `kpsc_validation` (валидация КПСЦ Excel): http://localhost:8002

У каждого сервиса есть:
- `GET /health`
- `POST /jobs` (создать задачу)
- `GET /jobs/{job_id}` (статус)
- `GET /jobs/{job_id}/artifacts` (список файлов)
- `GET /jobs/{job_id}/artifacts/{name}` (скачать файл)

Артефакты каждого запуска сохраняются на хосте в `packaged_services/runs/`.

---

## Примеры запросов (curl)

### 1) Drivers Excel

```bash
JOB_JSON=$(curl -s -F "file=@/path/to/driver.xlsx" http://localhost:8000/jobs)
echo "$JOB_JSON"
```

Дальше:
```bash
JOB_ID=<подставьте job_id>
curl -s http://localhost:8000/jobs/$JOB_ID | jq .
curl -s http://localhost:8000/jobs/$JOB_ID/artifacts | jq .
curl -L -o 03_report.xlsx http://localhost:8000/jobs/$JOB_ID/artifacts/03_report.xlsx
curl -L -o artifacts.zip http://localhost:8000/jobs/$JOB_ID/artifacts/artifacts.zip
```

### 2) Word audit (одиночная проверка)

```bash
curl -s -F "job_type=order_comp_ppu" \
  -F "target_doc=@/path/to/order_comp_ppu_version_3_10.docx" \
  -F "template_doc=@/path/to/order_comp_ppu_template.docx" \
  -F "tz_doc=@/path/to/order_comp_ppu_tz_3_10.docx" \
  http://localhost:8001/jobs
```

### 2b) Word audit (package: 4 документа)

```bash
curl -s -F "job_type=package" \
  -F "order_comp_doc=@/path/to/order_comp_ppu_version_3_10.docx" \
  -F "reg_comp_doc=@/path/to/reg_comp_ppu_version_3_10.docx" \
  -F "order_ppu_doc=@/path/to/order_ppu_version_3_10.docx" \
  -F "reg_ppu_doc=@/path/to/reg_ppu_version_3_10.docx" \
  http://localhost:8001/jobs
```

### 3) КПСЦ validation

```bash
curl -s -F "file=@/path/to/kpsc.xlsx" http://localhost:8002/jobs
```

С override правил:
```bash
curl -s -F "file=@/path/to/kpsc.xlsx" -F "rules_file=@/path/to/validation_rules.json" http://localhost:8002/jobs
```
