# ui-demo — Веб-интерфейс демо

React-приложение для демонстрации ИИ-аудита. Полностью на mock-данных, бэкенд не подключён.

## Стек

React 19 + Vite 7 + Tailwind CSS 4 (`@import "tailwindcss"`) + lucide-react (иконки)

## Структура

```
ui-demo/
├── src/
│   ├── main.jsx      # Точка входа React
│   ├── App.jsx       # Всё приложение (~900 строк): данные + 3 экрана + логика
│   ├── App.css       # Кастомные анимации (fade-in, slide-up, pulse)
│   └── index.css     # Tailwind import
├── index.html        # HTML-шаблон
├── vite.config.js    # Vite + React plugin
└── postcss.config.js # Tailwind v4 через @tailwindcss/postcss
```

## Экраны (flow)

```
LANDING → PROCESSING → RESULTS
(2.1с)     (12.4с)     (отчёт)
```

- **LANDING:** Приветствие + 5 критериев + drag-n-drop загрузка файла
- **PROCESSING:** Последовательный анализ 5 критериев, живой таймер, нарастающий счётчик 0→20
- **RESULTS:** Сводка (3 метрики + bar-chart + severity) → фильтр-табы → 20 карточек ошибок → CSV/JSON экспорт

## Запуск

```bash
npm install   # Первый раз
npm run dev   # Dev-сервер (http://localhost:5173)
npm run build # Production → dist/
```

## Для подключения бэкенда (будущее)

Заменить mock-таймеры на реальные API-вызовы:
- Загрузка: `POST /upload` вместо имитации 2.1с
- Анализ: SSE/WebSocket вместо цепочки setTimeout
- Результаты: данные из API вместо `DEMO_VIOLATIONS`
