# Наш стенд (внутреннее — коллегам не передаётся)

## Доступ
- Стенд `ms-llm-dev`: команда `dev` = SSH `nemovm@172.16.10.64` через корпоративный OpenVPN.
  При обрыве: `dev -K`, повтор; таймаут — попросить владельца перезапустить OpenVPN.
  Кириллица в аргументах ssh ломается — файлы передавать `tar | ssh 'tar xzf -'` или base64.
- Репозитории: `~/projects/mos_analiz_refactor` (GitHub `Maskedxxx/mos_analiz`, ветка `dev`; `main` — старый прод),
  `~/projects/mos_generated` (GitHub `Maskedxxx/mos_generated`, ветка `dev`). Push по SSH-ключу
  `~/.ssh/id_ed25519_github`. Коммиты `Maskedxxx <aangers07@gmail.com>` без постфикса, только по слову владельца.
- `~/projects/mos_analiz` — старый прод (`:8080`, `:5173`), заморожен, не трогать.

## Сервисы
| Что | Порт | Запуск | Screen |
|---|---|---|---|
| проверка, бэкенд | 8081 | `scripts/restart_backend.sh` | `backend` |
| фронт (vite preview, прокси `/api`, `/gen`) | 5174 | `scripts/restart_frontend.sh` | `frontend` |
| генерация | 8090 | `mos_generated/scripts/restart.sh` | `mosgen` |
| туннели tuna | — | `mosaudit.ru.tuna.am` → 5174, `mosgen.ru.tuna.am` → 8090 (fallback) | `mosaudit-tuna`, `mosgen-tuna` |

`~/wf_restart_backend.sh`, `~/wf_rebuild_frontend.sh` — симлинки на `scripts/`. `.env` стенда лежит
в корне репо проверки (логин `admin`/`guest`, пароль там). Docker на стенде не установлен (нужен sudo).

## Модели (хост 172.16.10.35, стенд Влада, SSH туда нет)
`:11437` LLM Qwen3.6-35B-A3B (llama.cpp, GGUF), `:11436` qwen-fast, `:11438` OCR PaddleOCR-VL, `:11439` layout Heron.
Адреса — в `.env` (`LLM_BASE_URL`, `OCR_BASE_URL`, `LAYOUT_BASE_URL`).

## Где состояние и история
- Журнал проверки: `mos_analiz_refactor/docs/internal/ДОВОДКА_ПРАВИЛ_состояние.md` — §0 читать первым.
- Журнал генерации: `mos_generated/docs/internal/СОСТОЯНИЕ_ГЕНЕРАЦИЯ.md` — §0.
- Трекер задач: `~/projects/ТРЕКЕР_МосМониторинг.md` (копия `~/Downloads/` на Маке), раздел «Где что лежит».
- Маппинг типов ↔ каталог заказчика: `~/Downloads/МАППИНГ_типов_МосМониторинг.md` (Мак).
- Бенчмарк с экспертной разметкой: `~/projects/_benchmark/` (dataset.json, results/, metrics_*.json);
  исходные документы бенчмарка — `~/Downloads/Для Максима-1/` на Маке.
- Образцы: `test_docs/` (kpsc, драйверы, карточки, план-график, pdf), `tests/data/<тип>/sample.xlsx`
  (архив `~/Downloads/ПЕРЕДАЧА_фикстуры_тестов.zip`).
- Методички заказчика: `docs/методички/` (18 PDF, не в git).

## Договорённости с владельцем
- Факты — только чтением кода и артефактов, не по памяти.
- Один вариант решения, не список на выбор. Таблицы рамками ┌─┬─┐.
- Перед правкой файла — копия; не упрощать в ущерб качеству; все изменения согласовывать.
- Vision-документы читать целиком, без кропов.
