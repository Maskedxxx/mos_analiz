# ui-demo — Веб-интерфейс аудита документов

React-приложение для ИИ-аудита документов. Подключено к бэкенду (api_server.py).

## Стек

React 19 + Vite 7 + Tailwind CSS 4 (`@import "tailwindcss"`) + lucide-react (иконки)

## Структура

```
ui-demo/
├── src/
│   ├── main.jsx            # Точка входа React
│   ├── App.jsx             # Оркестратор: 4 экрана + хедер (~90 строк)
│   ├── ScreenLogin.jsx     # Форма логин/пароль
│   ├── ScreenUpload.jsx    # Выбор типа + drag-drop файла
│   ├── ScreenProgress.jsx  # Stepper 5 шагов + SSE + прогресс-бар
│   ├── ScreenResults.jsx   # Сводка + замечания + скачать Excel
│   ├── api.js              # HTTP + SSE клиент (5 функций)
│   ├── App.css             # Кастомные анимации (fade-in, slide-up, pulse)
│   └── index.css           # Tailwind import
├── index.html              # HTML-шаблон
├── vite.config.js          # Vite + React plugin + proxy /api → :8080
└── postcss.config.js       # Tailwind v4 через @tailwindcss/postcss
```

## Экраны (flow)

```
LOGIN → UPLOAD → PROGRESS → RESULTS
                  (SSE)     (отчёт)
```

- **LOGIN:** Форма логин/пароль, POST /api/login, cookie auth_token
- **UPLOAD:** Dropdown типов (GET /api/types) + drag-drop файла + кнопка «Выполнить аудит» (POST /api/audit)
- **PROGRESS:** Вертикальный stepper 5 шагов (конвертация, OCR target, OCR шаблон, правила N/M, отчёт) + SSE-подписка + таймер
- **RESULTS:** 3 метрики + раскрывающиеся карточки замечаний + кнопки «Новый аудит» / «Скачать Excel»

## API-клиент (api.js)

- `login(login, password)` → POST /api/login
- `fetchDocTypes()` → GET /api/types
- `startAudit(file, docType)` → POST /api/audit
- `subscribeToProgress(sessionId, onEvent)` → EventSource /api/audit/{id}/events
- `getDownloadUrl(sessionId)` → строка URL /api/audit/{id}/download

## Запуск

```bash
# Бэкенд (из корня проекта)
bash start_server.sh

# Фронтенд
cd ui-demo && npm install && npm run dev
# → http://localhost:5173 (proxy → :8080)

# Логин: admin / admin (по умолчанию)
```
