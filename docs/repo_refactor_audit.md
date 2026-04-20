# Аудит репозитория и целевая структура рефакторинга

Дата: 2026-04-20

## Цель

Привести репозиторий к форме, удобной для агентной разработки:

- без глубокой вложенности;
- без избыточного дробления на микрофайлы;
- с единым жизненным циклом аудита;
- с явными границами между web, orchestration, extraction, validation и reporting;
- с минимальным количеством "особых путей" исполнения.

## Ключевые выводы

- Репозиторий функционально богатый, но архитектурно раздвоен.
- В коде одновременно живут три разных мира:
  - стандартный `AuditEngine`;
  - `multi_rule`-аудит для большинства типов документов;
  - спецдвижки `drivers`, `kpsc`, `kartochka_proekta`, `plan_grafik`.
- Конфигурационная модель сильная, но кодовая модель исполнения слишком неоднородная.
- Для агентов главный тормоз сейчас не размер проекта, а количество архитектурных исключений.

## Findings

### Critical

1. В коде зашиты чувствительные значения и дефолтные секреты.
   [api_server.py](/home/nemovm/projects/mos_analiz/api_server.py:68) содержит логин, пароль и токен по умолчанию.
   [healthcheck.sh](/home/nemovm/projects/mos_analiz/healthcheck.sh:7) содержит `sudo`-пароль в открытом виде.

2. Веб-контур не поддерживает multi-file сценарий, хотя движок его поддерживает.
   [api_server.py](/home/nemovm/projects/mos_analiz/api_server.py:172) принимает только один `file`.
   [ui-demo/src/api.js](/home/nemovm/projects/mos_analiz/ui-demo/src/api.js:30) тоже отправляет только один файл.
   При этом [audit_engine/engine.py](/home/nemovm/projects/mos_analiz/audit_engine/engine.py:229) умеет `secondary_path`.
   Итог: один из штатных сценариев (`prikaz_vyhod`) полноценно работает только через CLI.

3. Для всех `multi_rule` типов движок всё равно парсит шаблон, хотя сам `multi_rule` использует только `target_doc`.
   [audit_engine/engine.py](/home/nemovm/projects/mos_analiz/audit_engine/engine.py:235) всегда вызывает `_get_template`.
   Реальный `multi_rule` запуск использует только `parsed=target_doc`: [audit_engine/engine.py](/home/nemovm/projects/mos_analiz/audit_engine/engine.py:295).
   Это лишняя стоимость по времени, OCR и логам.

4. Часть валидаторов `kartochka_proekta` обходит локальную LLM-инфраструктуру и жёстко использует облачную модель.
   [validate_6_justification.py](/home/nemovm/projects/mos_analiz/audit_engine/kartochka_proekta/validation_scripts/validate_6_justification.py:45) создаёт `OpenAI(api_key=api_key)`.
   [validate_6_justification.py](/home/nemovm/projects/mos_analiz/audit_engine/kartochka_proekta/validation_scripts/validate_6_justification.py:78) хардкодит `gpt-4.1-mini`.
   То же самое в [validate_3_flow_name.py](/home/nemovm/projects/mos_analiz/audit_engine/kartochka_proekta/validation_scripts/validate_3_flow_name.py:45) и [validate_3_flow_name.py](/home/nemovm/projects/mos_analiz/audit_engine/kartochka_proekta/validation_scripts/validate_3_flow_name.py:75).
   Это ломает единый operational contract проекта.

### High

5. Протокол прогресса сервера и UI рассинхронизирован.
   Сервер шлёт `queue`: [api_server.py](/home/nemovm/projects/mos_analiz/api_server.py:425).
   UI ожидает `queue`: [ui-demo/src/ScreenProgress.jsx](/home/nemovm/projects/mos_analiz/ui-demo/src/ScreenProgress.jsx:35).
   Но transport-слой на него не подписывается: [ui-demo/src/api.js](/home/nemovm/projects/mos_analiz/ui-demo/src/api.js:55).
   Аналогично UI не использует `checking_rules` и `checking_rules_done`, хотя движок их эмитит.

6. API хранит состояние аудитов только в памяти и сериализует все прогоны глобальной блокировкой.
   In-memory store: [api_server.py](/home/nemovm/projects/mos_analiz/api_server.py:75).
   Глобальная блокировка: [api_server.py](/home/nemovm/projects/mos_analiz/api_server.py:54).
   Реальная сериализация запуска: [api_server.py](/home/nemovm/projects/mos_analiz/api_server.py:430).
   Это плохо и для масштабирования, и для надёжности после рестартов.

7. Спецдвижки `kpsc` и `kartochka_proekta` оркестрируют валидаторы через subprocess на каждое правило.
   [audit_engine/kpsc/run_validations.py](/home/nemovm/projects/mos_analiz/audit_engine/kpsc/run_validations.py:128) и [audit_engine/kartochka_proekta/run_validations.py](/home/nemovm/projects/mos_analiz/audit_engine/kartochka_proekta/run_validations.py:94).
   Это делает пайплайн тяжёлым, трудно трассируемым и плохо управляемым агентами.

8. Документация и infra-конфиг дрейфуют относительно реального кода.
   В [QUICKSTART.md](/home/nemovm/projects/mos_analiz/QUICKSTART.md:21) OCR стартует на `8000`.
   В [docker-compose.yml](/home/nemovm/projects/mos_analiz/docker-compose.yml:19) OCR уже на `8010`.
   Боевые `config.json` массово указывают на удалённый `172.16.10.35`, а локальные инструкции описывают другой режим.

### Medium

9. Инфраструктурные пути захардкожены под одну машину.
   [start_server.sh](/home/nemovm/projects/mos_analiz/start_server.sh:10) и [docker-compose.yml](/home/nemovm/projects/mos_analiz/docker-compose.yml:10) завязаны на `/home/masked/...`.
   Это снижает переносимость и усложняет запуск агентами в другой среде.

10. `multi_rule.py` содержит уже нетривиальную ручную санацию сломанного JSON-ответа LLM.
   [audit_engine/multi_rule.py](/home/nemovm/projects/mos_analiz/audit_engine/multi_rule.py:200) и далее.
   Это симптом того, что уровень контракта с LLM и нормализации ответа не выделен в отдельный устойчивый слой.

11. В проекте слишком много execution-patterns для одного домена.
   Сейчас есть:
   - standard multi-rule;
   - standard legacy per-rule;
   - xlsx special pipelines;
   - UI/API orchestration с отдельной логикой очереди.
   Для агента это означает высокий cost на построение mental model.

12. Структура папок исторически наращивалась по типам задач, а не по стабильным архитектурным слоям.
   В итоге логика одного сценария размазана между:
   - `audit_engine/*`;
   - `doc_configs/*`;
   - `api_server.py`;
   - `ui-demo/*`;
   - спецмодулями внутри `audit_engine/kpsc` и `audit_engine/kartochka_proekta`.

## Техдолг по слоям

### Web/API

- Нет устойчивого job-store.
- Нет единой схемы событий прогресса.
- Нет поддержки multi-file upload.
- Весь сервер сосредоточен в одном большом файле `api_server.py`.

### Audit orchestration

- `AuditEngine` знает слишком много сразу:
  - конфиги;
  - парсеры;
  - кэш шаблона;
  - режимы правил;
  - логи;
  - Excel.
- `legacy` и `multi_rule` смешаны в одном orchestrator.
- Шаблонный контур не отделён от target-only контура.

### Extraction/parsing

- Есть несколько хороших пайплайнов, но они лежат в разных подсистемах и собраны разными паттернами.
- Общий контракт "документ -> parsed scopes" есть по факту, но не оформлен как единый интерфейс.

### Validation

- `multi_rule` и non-LLM проверки живут рядом.
- `kpsc` и `kartochka_proekta` реализованы другим паттерном: parser scripts + validation scripts + subprocess orchestration.
- Валидации часто хардкодят свои модели и клиента вместо использования общего слоя.

### Infrastructure

- Секреты в коде.
- Разный режим запуска в docs, shell и compose.
- Абсолютные пути и machine-specific volumes.

## Что не надо делать в рефакторинге

- Не раскладывать проект на десятки микромодулей по 50 строк.
- Не углублять вложенность больше чем на 2 уровня под корневым пакетом.
- Не размазывать один жизненный цикл документа по 6 разным папкам.
- Не переносить `doc_type`-логику из конфигов в код без необходимости.

## Целевая структура

Нужна одна корневая кодовая папка и один конфигурационный каталог типов документов.

```text
src/mosaudit/
  api/
    app.py
    routes.py
    auth.py
    jobs.py
    schemas.py

  core/
    settings.py
    models.py
    events.py
    registry.py
    logging.py

  audit/
    service.py
    standard.py
    special.py
    progress.py

  extract/
    office.py
    ocr.py
    paddle.py
    xlsx.py

  validate/
    standard.py
    non_llm.py
    multi_rule.py
    kpsc.py
    kartochka.py
    plan_grafik.py
    drivers.py

  report/
    excel.py
    artifacts.py

doc_specs/
  akt_nachala/
  cheklist_eu/
  ...

ui/
  src/
  package.json

tests/
docs/
scripts/
```

## Что именно переедет

- `api_server.py` распадается на `api/app.py`, `api/routes.py`, `api/jobs.py`, `api/auth.py`.
- `audit_engine/models.py`, `logger.py`, часть registry-логики переходят в `core/`.
- `AuditEngine` распадается на:
  - `audit/service.py` как фасад;
  - `audit/standard.py` для стандартного document pipeline;
  - `audit/special.py` для адаптеров спецдвижков.
- `vision_parser`, `ocr_parser`, `paddle_parser`, `docx_parser`, `pptx_parser` собираются под одним слоем `extract/`.
- `multi_rule.py` и `non_llm_checks/*` переезжают в `validate/`.
- `kpsc`, `kartochka_proekta`, `plan_grafik`, `drivers` становятся адаптерами в `validate/*.py` и `extract/xlsx.py`, а не отдельными мини-фреймворками внутри `audit_engine/`.

## Архитектурный принцип после рефакторинга

Один документ любого типа должен проходить один и тот же верхнеуровневый контракт:

1. `DocumentSpec` описывает тип.
2. `Extractor` строит `ParsedDocument`.
3. `Validator` возвращает `ValidationResult`.
4. `ReportWriter` сохраняет Excel и артефакты.
5. `AuditService` связывает это в один job.

Это важнее, чем сохранять текущие папки.

## Целевой runtime contract

- Один входной объект задания аудита.
- Один store состояния job.
- Один формат progress events.
- Один формат результата.
- Один клиент LLM.
- Один слой конфигурации моделей.

## План рефакторинга

### Phase 0. Stabilize

- Убрать секреты из кода.
- Зафиксировать текущий baseline тестами smoke + golden runs.
- Отделить machine-specific `.env` и локальные compose overrides.

### Phase 1. API contract

- Вынести API в пакет `src/mosaudit/api`.
- Ввести `JobStore` и `ProgressEvent`.
- Починить event protocol UI/API.
- Добавить поддержку multi-file upload.

### Phase 2. Unified audit service

- Ввести `AuditService`.
- Вынести `multi_rule` и `legacy` в отдельные runner-ы.
- Не парсить шаблон в `multi_rule`-режиме.
- Оставить один фасад запуска для CLI и API.

### Phase 3. Validation cleanup

- Перевести спецвалидаторы на общий LLM client/settings.
- Убрать хардкод моделей и прямой `OpenAI(api_key=...)` из rule scripts.
- Свести subprocess orchestration к минимуму.

### Phase 4. Repository cleanup

- Переехать в `src/` layout.
- Переименовать `doc_configs` в `doc_specs` или `document_types`.
- Перенести shell-утилиты в `scripts/`.
- Упростить корень репозитория.

## Рекомендуемый конечный вид корня репозитория

```text
.
├── src/
├── doc_specs/
├── ui/
├── tests/
├── docs/
├── scripts/
├── docker-compose.yml
├── pyproject.toml
└── README.md
```

## Приоритет внедрения

1. Безопасность и конфигурация окружения.
2. API/job-store/progress contract.
3. Единый audit service.
4. Уплощение структуры пакетов.
5. Нормализация спецдвижков.

## Критерий успеха

Репозиторий можно считать приведённым в порядок для агентной разработки, если:

- новый агент понимает основной lifecycle по 4-5 файлам, а не по 20+;
- любой `doc_type` запускается через один фасад;
- web и CLI используют один и тот же orchestration path;
- LLM/infra настройки задаются централизованно;
- для добавления нового типа документа не нужно изобретать новый execution pattern.
