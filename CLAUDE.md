# МосАнализ — ИИ-проверка документов (рефактор)

Сервис проверяет организационно-распорядительные документы предприятий-участников
нацпроекта на соответствие методикам заказчика (Мостратегия). Находит нарушения,
объясняет каждое и выгружает отчёт в Excel.

## Куда смотреть

| Что нужно | Где |
|---|---|
| Как запустить, конфигурация, API, структура | `README.md` |
| Как устроен пайплайн, узлы, решения (ADR) | `docs/ARCHITECTURE.md` |
| Как добавить тип документа / спецдвижок | `docs/ADD_DOC_TYPE.md` |
| Состояние проекта, история работ, открытые долги | `docs/internal/ДОВОДКА_ПРАВИЛ_состояние.md` — **§0 читать первым** |
| Задачи по обоим сервисам | `~/projects/ТРЕКЕР_МосМониторинг.md` (стенд) |
| Соответствие типов каталогу заказчика | `~/Downloads/МАППИНГ_типов_МосМониторинг.md` (у владельца на Маке) |
| Методики заказчика (первоисточник правил) | `docs/методички/` — 18 PDF (не в git) |
| Расхождения экспертной разметки с файлами | `docs/internal/MISMARK_реестр.md` |
| Механики обоих сервисов | скил `mosmonitoring-mechanics` |

## Жёсткие правила

1. **После правок кода или правил перезапустить бэкенд** — Python не подхватывает изменения
   на лету, uvicorn без `--reload`. Команда: `bash scripts/restart_backend.sh`
   (`~/wf_restart_backend.sh` на стенде — симлинк на неё). Фронт: `bash scripts/restart_frontend.sh`.
   Исключение — прогоны бенчмарка: они каждый раз запускают `main.py` заново.
2. **Реестр спецдвижков `SPECIAL_ENGINE_RUNNERS` — только в `src/engines.py`.** CLI и веб-бэкенд
   импортируют его оттуда. Новый движок = строка в реестре + поле `engine` в config.json типа;
   `tests/test_registry.py` проверяет, что реестр и конфиги согласованы. (До 07.09.2026 реестр
   был продублирован в `main.py` и `server.py` — так терялись 6 типов.)
3. **Адреса модели и OCR, логин/пароль — только через `.env`** (`config/llm.py`, `config/parsers.py`,
   `.env.example`). В коде и в `doc_configs/*/config.json` адресов и имён модели быть не должно;
   поля `model`/`llm_base_url` в конфиге типа — осознанный override, по умолчанию не нужны.
4. **Не коммитить и не пушить без слова владельца.** Идентичность коммитов:
   `Maskedxxx <aangers07@gmail.com>`, без постфикса Claude.
5. **Факты — только чтением кода и артефактов**, не по памяти и не по документации:
   документы устаревают, проверять текущее состояние инструментом.
6. **Кастомные правила пользователя** живут в `doc_configs/<тип>/rules_custom.json`
   с индексами от 200. Базовые `rules_multi.json` НЕ трогать — движок мёрджит их на каждом прогоне.
7. **Клиент к модели создаётся только в `src/llm/client.py`** (`make_llm_client`). Параметры
   вызова в `multi_rule.py` (temperature 0.7, top_p, top_k, ретраи со сменой seed) — калиброваны,
   не менять без бенчмарка.
8. **Тесты после правок:** `.venv/bin/pytest -q --ignore=tests/test_kpsc_parsers.py` (~3 мин).
   Фикстуры спецдвижков — `tests/data/` (в git нет, см. `tests/data/README.md`).

## Как устроено

**Два движка проверки:**

| Движок | Признак в `config.json` | Как работает | Типов |
|---|---|---|---|
| generic-LLM | есть `parser_by_ext`, нет `engine` | текст документа + правила → языковая модель → вердикты | 21 |
| special-runner | есть `engine` | Python-валидатор по таблице Excel, без модели | 10 |

Всего **31 тип** в `doc_configs/`. Типы подхватываются сканом каталога — отдельного реестра типов нет.

**Допустимые форматы файла** выводятся из конфига типа: у generic — ключи `parser_by_ext`,
у спецдвижков — всегда `.xlsx`. Зашитых списков расширений быть не должно: так уже ломались
презентации (не принимали `.pptx`) и текстовые типы (принимали `.doc` без парсера и падали).

**Парсинг:** PDF — всегда через OCR постранично (VLM `OCR_BASE_URL` + layout `LAYOUT_BASE_URL`);
DOCX — python-docx плюс OCR только для картинок шапки; XLSX — openpyxl.

**Артефакты прогона** — каталог `logs_result/<тип>/session_<время>/`: промпты, сырой и разобранный
ответ модели, вердикт по каждому правилу (включая пройденные), `final_results.json` только
с нарушениями, `audit_result.xlsx`.

## Стек

Python 3.12, FastAPI + uvicorn, pydantic-settings, openpyxl, python-docx, PyMuPDF; фронт —
React + Vite + Tailwind в `ui-demo/` (Node 22). Языковая модель — Qwen через OpenAI-совместимый
API (на стенде llama.cpp/GGUF: имя модели в запросе игнорируется, отвечает загруженная).
Рассуждения модели отключены через `chat_template_kwargs: {"enable_thinking": false}`.

## Дерево проекта

```
config/                   llm.py, parsers.py — конфигурация из .env поверх дефолтов
doc_configs/<тип>/        config.json, rules_multi.json, rules_methodology.json,
                          sections.json, rules_custom.json, validation_rules.json
docs/                     ARCHITECTURE.md, ADD_DOC_TYPE.md, методички/ (PDF, не в git)
scripts/                  restart_backend.sh, restart_frontend.sh, restart_all.sh, healthcheck.sh
src/engines.py            реестр спецдвижков + detect_engine
src/audit/                движок generic-проверки, модели результата, Excel-отчёт
src/api/server.py         веб-бэкенд: авторизация, SSE-прогресс, правила, эндпоинты аудита
src/llm/                  client.py (клиент к модели), multi_rule.py (промпт, разбор ответа)
src/doc_type_validators/  спецдвижки (карточки, формы, КПСЦ, листы присутствия, сквозная сверка)
src/doc_type_parsers/     парсеры под конкретные таблицы
src/format_parsers/       docx, pptx, pdf (OCR + layout)
main.py                   CLI: python main.py --doc-type <тип> --target <файл>
tests/                    pytest; tests/data/ — фикстуры спецдвижков (не в git)
ui-demo/                  React-фронт: сайдбар «Аудит» / «Генерация»
logs_result/, uploads/    артефакты прогонов и загрузки (не в git)
test_docs/                образцы документов (не в git)
```

## Запуск

```bash
bash scripts/restart_all.sh       # бэкенд :8081 + фронт :5174 (screen `backend`, `frontend`) + healthcheck
docker compose up -d --build      # то же в Docker: api, gen (../mos_generated), ui с nginx-прокси
python main.py --doc-type protokol_vypolneniya --target <файл>   # проверка из CLI
```

Домен: `https://mosaudit.ru.tuna.am` (tuna → :5174) — объединённый интерфейс, слева переключатель
«Аудит» / «Генерация». Генерация ходит на `:8090` через прокси `/gen` с rewrite
(префикс обязателен: оба бэкенда отдают `/api/types/*`).

## Доступы

Стенд `ms-llm-dev` — команда `dev` (SSH `nemovm@172.16.10.64`, корпоративный OpenVPN).
Вход в интерфейс: `guest` / пароль в `.env` стенда (у владельца).
Репозиторий: `git@github.com:Maskedxxx/mos_analiz.git`, разработка в ветке `dev`
(`main` — старый прод, заморожен).
