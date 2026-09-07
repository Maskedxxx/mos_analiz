# МосАнализ — проверка документов по правилам с помощью языковой модели

Сервис принимает документ (DOCX, PPTX, PDF-скан или XLSX-форму), набор правил для его типа и
возвращает отчёт в Excel: по каждому правилу — пройдено или нарушение, с цитатой из документа и
обоснованием. Языковая модель подключается по OpenAI-совместимому API (vLLM, llama.cpp, Ollama —
любой сервер с `/v1/chat/completions`). Типы документов и правила описываются JSON-файлами,
без кода.

Для стенда заказчика (МосМониторинг) сервис проверяет документы предприятий-участников
нацпроекта «Производительность труда» по методикам РЦК. Заведён 31 тип документов.

## Как это работает

```
файл + тип документа
  │
  ├─ гейт формата: допустимые расширения — из parser_by_ext в config.json типа
  ├─ парсер формата → полный текст документа            src/format_parsers/ (docx, pptx, pdf→OCR)
  ├─ сборка промпта: имя файла + текст + карта секций + правила   src/llm/multi_rule.py
  ├─ вызов модели, два прохода: базовые правила и методические
  ├─ разбор JSON-вердиктов → список нарушений
  └─ Excel-отчёт (все правила, статус ОК / FAIL)             src/audit/excel_reporter.py
       + артефакты прогона в logs_result/<тип>/session_<время>/
```

Второй путь — **спецдвижки** для табличных форм (карточка проекта, формы «в цифрах»,
план-график, КПСЦ, листы присутствия): парсер XLSX → Python-валидаторы → тот же Excel-отчёт.
Модель они вызывают точечно или не вызывают вовсе. Подробно — `docs/ARCHITECTURE.md`.

## Требования

| Что | Версия | Зачем |
|---|---|---|
| Python | 3.12 | бэкенд и CLI |
| Node.js | 20+ (проверено на 22) | сборка веб-интерфейса |
| LibreOffice (`soffice`), poppler-utils | любые актуальные | конвертация DOCX/PPTX → PDF → PNG для OCR |
| LLM-сервис | OpenAI-совместимый `/v1` | обязателен |
| OCR-сервис (VLM, OpenAI-совместимый `/v1`) и layout-детектор | | нужны только для PDF-сканов и картинок в шапке DOCX |

Без OCR и layout-сервисов работают все типы, у которых в `parser_by_ext` нет `.pdf`.

## Быстрый старт (5 минут)

```bash
git clone <repo> mos_analiz && cd mos_analiz
python3 -m venv .venv && .venv/bin/pip install -r requirements-dev.txt
cp .env.example .env            # заполнить: AUDIT_* (логин/пароль/токен), LLM_BASE_URL, LLM_MODEL
.venv/bin/python main.py --list-types
.venv/bin/python main.py --doc-type akt_nachala \
    --target doc_configs/akt_nachala/template/template.docx --session-dir /tmp/akt
ls /tmp/akt/audit_result.xlsx    # отчёт
```

Веб-интерфейс:

```bash
cd ui-demo && npm ci && cd ..
bash scripts/restart_all.sh      # бэкенд :8081 + фронт :5174 в screen-сессиях, затем healthcheck
```

Открыть `http://localhost:5174/`, войти логином и паролем из `.env`.

## Запуск в Docker (весь стенд одной командой)

Нужен Docker 24+ с Compose v2. Репозиторий генерации должен лежать рядом: `../mos_generated`
(compose собирает его образ сервисом `gen`).

```bash
cp .env.example .env             # AUDIT_*, LLM_BASE_URL, LLM_MODEL, при необходимости OCR_*/LAYOUT_*
docker compose up -d --build     # api :8081, gen :8090, ui :5174 (nginx: /api → api, /gen → gen)
docker compose ps                # все три — healthy
curl -s http://localhost:5174/api/health
```

Что монтируется томами (правится без пересборки): `doc_configs/` (типы и правила — после правки
`docker compose restart api`), `logs_result/`, `uploads/`, у генерации — `templates/`,
`logs_generated/` и стор значений по организациям. Адреса модели и OCR внутри контейнера — те же
`LLM_BASE_URL`/`OCR_BASE_URL` из `.env`: если модель на хосте, укажите `http://host.docker.internal:<порт>/v1/`
(Docker Desktop) или адрес хоста в сети. Порты наружу меняются переменными `API_PORT`, `GEN_PORT`, `UI_PORT`.
Остановить: `docker compose down` (тома с данными остаются).

## Конфигурация — переменные окружения

Все значения — в `.env` в корне проекта (образец `.env.example`), приоритет: переменная
процесса > `.env` > значение по умолчанию в коде.

| Переменная | Что | По умолчанию |
|---|---|---|
| `AUDIT_LOGIN`, `AUDIT_PASSWORD`, `AUDIT_TOKEN` | вход в веб-интерфейс и токен сессии | нет — без них сервер не стартует |
| `LLM_BASE_URL` | адрес модели до `/v1/` | `http://127.0.0.1:11437/v1/` |
| `LLM_MODEL` | имя модели, как отдаёт `GET /v1/models` | `Qwen3.6-35B-A3B` |
| `LLM_API_KEY` | ключ; для локальных серверов любое непустое | `none` |
| `LLM_MAX_TOKENS`, `LLM_SEED`, `OPENAI_TIMEOUT_SEC` | лимит ответа, seed, таймаут запроса | 4096, 42, 300 |
| `OCR_BASE_URL`, `OCR_API_KEY`, `LAYOUT_BASE_URL` | OCR-модель и layout-детектор для PDF | `127.0.0.1:11438`, `none`, `127.0.0.1:11439` |
| `CORS_ORIGINS` | адреса фронтенда через запятую | `http://localhost:5174,http://127.0.0.1:5174` |
| `API_HOST`, `API_PORT`, `UI_PORT`, `NODE_VERSION`, `LOG_DIR` | параметры скриптов запуска | `127.0.0.1`, 8081, 5174, —, `/tmp` |

Код конфигурации: `config/llm.py` (модель) и `config/parsers.py` (парсеры, OCR). Остальные
параметры парсеров (DPI, пороги детекции) — там же, с описанием каждого поля.

## Свой тип документа

Тип = папка в `doc_configs/<имя>/` с `config.json`, картой секций `sections.json` и правилами
`rules_multi.json` (плюс необязательные `rules_methodology.json`, `rules_custom.json`).
Код писать не нужно: тип подхватывается сканом каталога. Пошагово — **`docs/ADD_DOC_TYPE.md`**.

## Структура репозитория

```
config/            конфигурация рантайма: llm.py, parsers.py (pydantic-settings, .env)
doc_configs/       типы документов: по папке на тип (config.json, sections.json, rules_*.json)
docs/              ARCHITECTURE.md, ADD_DOC_TYPE.md
main.py            CLI и точка входа uvicorn (main:app)
scripts/           запуск: restart_backend.sh, restart_frontend.sh, restart_all.sh, healthcheck.sh
src/api/           веб-бэкенд FastAPI: авторизация, очередь аудитов, SSE-прогресс, правила, отчёты
src/audit/         generic-движок, модели результата, Excel-отчёт, логгер сессии
src/engines.py     реестр спецдвижков (единственное место) и detect_engine
src/format_parsers/  парсеры форматов: docx, pptx, pdf (OCR + layout)
src/doc_type_parsers/    парсеры конкретных XLSX-форм
src/doc_type_validators/ спецдвижки: валидаторы по данным парсеров, сквозная сверка комплекта
src/llm/           клиент к модели (единственное место создания), multi_rule — сборка промпта и разбор ответа
tests/             pytest: конфиги, парсеры, реестр, API, спецдвижки (фикстуры — tests/data/README.md)
ui-demo/           веб-интерфейс React + Vite (сайдбар «Аудит» / «Генерация»)
logs_result/       артефакты прогонов (не в git)
uploads/           загруженные через веб файлы (не в git)
```

Самый тяжёлый файл репозитория — `doc_configs/presentation_eu/template/template.pptx` (23 МБ),
образец презентации для smoke-тестов парсера; нужен.

## Использование

**CLI** — `python main.py --doc-type <тип> --target <файл> [--session-dir <папка>] [--out-xlsx <путь>]`.
Ещё: `--list-types`, `--parse-only` (только парсинг), `--print-prompts` (показать промпты без
вызова модели), `--rule-filter <индекс>` (одно правило), `--fail-on-violations` (exit 1 при нарушениях).

**HTTP API** (все, кроме `health` и `login`, требуют cookie или `Authorization: Bearer <AUDIT_TOKEN>`):

| Метод и путь | Что |
|---|---|
| `GET /api/health` | состояние сервера и внешних сервисов (LLM, OCR) с их адресами |
| `POST /api/login` `{login, password}` | cookie сессии |
| `GET /api/types` | типы документов с допустимыми расширениями |
| `GET /api/types/{тип}/rules`, `POST …/rules`, `POST …/rules/draft` | правила типа; добавление кастомного; черновик формулировки через модель |
| `POST /api/audit` (multipart: `file`, `doc_type`) | запуск, возвращает `session_id` |
| `POST /api/audit/cross` (multipart: `files[]`) | сквозная сверка комплекта документов |
| `GET /api/audit/{id}/events` | SSE-прогресс |
| `GET /api/audit/{id}/result` | 202 пока идёт, затем `{violations, rules_checked, duration_sec}` |
| `GET /api/audit/{id}/download` | Excel-отчёт |

**Артефакты прогона** — `logs_result/<тип>/session_<время>/`: `multi_rule_user_prompt.txt`
(что получила модель), `multi_rule_response_parsed.json` (вердикт по каждому правилу, включая
пройденные), `multi_rule_methodology_*` (второй слой), `final_results.json` (только нарушения),
`audit_result.xlsx`, `pipeline.log`.

**Отчёт Excel** — 8 колонок: №, Проверка, Тип проверки (базовая / методическая), Статус (ОК / FAIL),
Источник (МР/МУ), Целевой документ (цитата), Различие, Обоснование.

## Тесты

```bash
.venv/bin/pytest -q --ignore=tests/test_kpsc_parsers.py    # ~3 мин, часть тестов ходит в модель
```

Тесты спецдвижков ждут образцы в `tests/data/<тип>/sample.xlsx` (в git нет — документы с
персональными данными; без них пропускаются). Откуда взять — `tests/data/README.md`.
`tests/test_kpsc_parsers.py` требует каталог `test_docs/` с реальными КПСЦ.

## После правок

Python не подхватывает изменения на лету — после правки кода или правил перезапустить бэкенд:
`bash scripts/restart_backend.sh`. Фронт после правок в `ui-demo/`: `bash scripts/restart_frontend.sh`.
