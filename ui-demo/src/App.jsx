import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  FileText, Upload, CheckCircle, AlertCircle,
  Download, RefreshCw, X, Loader2,
  Clock, Palette, Calculator, LayoutList,
  MessageSquare, ClipboardCheck, BarChart3, Filter
} from 'lucide-react';
import './App.css';

// --- MOCK-ДАННЫЕ (демо-сценарий: 1 документ, 20 ошибок, 5 критериев) ---

// 5 категорий ИИ-анализа
const ANALYSIS_CATEGORIES = [
  { id: 'formatting',    name: 'Единообразие форматирования',            shortName: 'Форматирование',     icon: Palette,        color: 'purple' },
  { id: 'accuracy',      name: 'Корректность чисел, дат, реквизитов',    shortName: 'Реквизиты',          icon: Calculator,     color: 'red' },
  { id: 'structure',     name: 'Соответствие структуре',                 shortName: 'Структура',          icon: LayoutList,     color: 'blue' },
  { id: 'clarity',       name: 'Чёткость формулировок',                  shortName: 'Формулировки',       icon: MessageSquare,  color: 'amber' },
  { id: 'completeness',  name: 'Наличие всех необходимых элементов',     shortName: 'Элементы',           icon: ClipboardCheck, color: 'green' },
];

// Тайминги обработки по этапам (сумма = 12400ms)
const STAGES_TIMING = [
  { categoryId: 'formatting',   duration: 2800, errorsFound: 5 },
  { categoryId: 'accuracy',     duration: 3200, errorsFound: 5 },
  { categoryId: 'structure',    duration: 2600, errorsFound: 4 },
  { categoryId: 'clarity',      duration: 2000, errorsFound: 3 },
  { categoryId: 'completeness', duration: 1800, errorsFound: 3 },
];

// Метаданные демо-документа
const DEMO_DOCUMENT = {
  filename: 'Приказ_о_создании_ИЦ_ООО_Ромашка.docx',
  fileSize: '245 KB',
  uploadTime: 2.1,
  analysisTime: 12.4,
  totalViolations: 20,
};

// Цвета категорий → tailwind-классы
const CATEGORY_COLORS = {
  purple: { bg: 'bg-purple-100', text: 'text-purple-700', border: 'border-purple-300', bar: 'bg-purple-500' },
  red:    { bg: 'bg-red-100',    text: 'text-red-700',    border: 'border-red-300',    bar: 'bg-red-500' },
  blue:   { bg: 'bg-blue-100',   text: 'text-blue-700',   border: 'border-blue-300',   bar: 'bg-blue-500' },
  amber:  { bg: 'bg-amber-100',  text: 'text-amber-700',  border: 'border-amber-300',  bar: 'bg-amber-500' },
  green:  { bg: 'bg-green-100',  text: 'text-green-700',  border: 'border-green-300',  bar: 'bg-green-500' },
};

// Цвета severity
const SEVERITY_CONFIG = {
  critical: { label: 'КРИТИЧ.',  labelFull: 'критических',  bg: 'bg-red-100',    text: 'text-red-700',    border: 'border-red-300',   dot: 'bg-red-500' },
  major:    { label: 'СУЩЕСТВ.', labelFull: 'существенных', bg: 'bg-orange-100', text: 'text-orange-700', border: 'border-orange-300', dot: 'bg-orange-500' },
  minor:    { label: 'НЕЗНАЧИТ.',labelFull: 'незначительных',bg: 'bg-yellow-100', text: 'text-yellow-700', border: 'border-yellow-300', dot: 'bg-yellow-500' },
};

// 20 ошибок демо-документа (5+5+4+3+3), severity: 4 critical, 9 major, 7 minor
const DEMO_VIOLATIONS = [
  // --- Форматирование (5) ---
  {
    id: 1, category: 'formatting', severity: 'major',
    rule: 'Шрифт основного текста',
    location: 'стр. 1–4, основной текст',
    target: 'Шрифт: Arial, 11 пт',
    expected: 'Times New Roman, 14 пт (ГОСТ Р 7.0.97-2016)',
    diff: 'Используется шрифт Arial 11 пт вместо требуемого Times New Roman 14 пт. Весь основной текст документа оформлен некорректным шрифтом.'
  },
  {
    id: 2, category: 'formatting', severity: 'minor',
    rule: 'Межстрочный интервал',
    location: 'стр. 1–4, все абзацы',
    target: 'Интервал: одинарный (1.0)',
    expected: 'Полуторный интервал (1.5)',
    diff: 'Межстрочный интервал одинарный, требуется полуторный по стандарту оформления организационно-распорядительных документов.'
  },
  {
    id: 3, category: 'formatting', severity: 'minor',
    rule: 'Выравнивание заголовка',
    location: 'стр. 1, заголовок «ПРИКАЗ»',
    target: 'Выравнивание: по левому краю',
    expected: 'По центру страницы',
    diff: 'Слово «ПРИКАЗ» выровнено по левому краю. Заголовки организационно-распорядительных документов располагаются по центру.'
  },
  {
    id: 4, category: 'formatting', severity: 'minor',
    rule: 'Нумерация на титульной странице',
    location: 'стр. 1, нижний колонтитул',
    target: 'Номер страницы: «1»',
    expected: 'Без номера на первой странице',
    diff: 'На титульной (первой) странице приказа проставлен номер. Первая страница не нумеруется.'
  },
  {
    id: 5, category: 'formatting', severity: 'major',
    rule: 'Отступ красной строки',
    location: 'стр. 2–4, абзацы основного текста',
    target: 'Отступ: 0 мм (отсутствует)',
    expected: '12.5 мм (1.25 см)',
    diff: 'Абзацы основного текста не имеют отступа первой строки. Стандартный отступ красной строки — 1.25 см.'
  },

  // --- Реквизиты (5) ---
  {
    id: 6, category: 'accuracy', severity: 'critical',
    rule: 'Дата приказа',
    location: 'стр. 1, шапка',
    target: '«27. г.»',
    expected: 'Полная дата: «27.01.2026 г.» или «27 января 2026 г.»',
    diff: 'Дата приказа неполная — отсутствует месяц и год. Документ с неполной датой не имеет юридической силы.'
  },
  {
    id: 7, category: 'accuracy', severity: 'critical',
    rule: 'Номер приказа',
    location: 'стр. 1, шапка',
    target: '«№» (пустой)',
    expected: '«№ 15-БП» или аналогичный',
    diff: 'После символа «№» отсутствует номер приказа. Документ без регистрационного номера не подлежит учёту и не имеет юридической силы.'
  },
  {
    id: 8, category: 'accuracy', severity: 'major',
    rule: 'Наименование организации',
    location: 'стр. 1 (шапка) ↔ стр. 3 (Приложение 1, п.1.1.1)',
    target: 'Шапка: «ООО «Ромашка»»\nп.1.1.1: «ООО «Своло»»',
    expected: 'Единое наименование: «ООО «Ромашка»»',
    diff: 'В п.1.1.1 приложения указано «ООО «Своло»», что не совпадает с наименованием из шапки «ООО «Ромашка»». Возможно, текст скопирован из другого документа.'
  },
  {
    id: 9, category: 'accuracy', severity: 'major',
    rule: 'Сумма в п.5.2',
    location: 'стр. 3, п.5.2 «Бюджет»',
    target: 'Итого: 1 260 000 ₽ (сумма строк: 1 350 000 ₽)',
    expected: 'Итого: 1 350 000 ₽',
    diff: 'Итоговая сумма бюджета (1 260 000 ₽) не соответствует сумме отдельных статей (1 350 000 ₽). Расхождение: 90 000 ₽.'
  },
  {
    id: 10, category: 'accuracy', severity: 'major',
    rule: 'ИНН организации',
    location: 'стр. 1, реквизиты',
    target: 'ИНН: 77234561 (9 цифр)',
    expected: '10 цифр для юр. лица / 12 для ИП',
    diff: 'ИНН содержит 9 цифр вместо 10. Некорректный ИНН делает невозможной идентификацию организации в реестрах.'
  },

  // --- Структура (4) ---
  {
    id: 11, category: 'structure', severity: 'major',
    rule: 'Отсутствие раздела «Ответственность»',
    location: 'Приложение 1, Положение об ИЦ',
    target: 'Разделы: 1–5 (без раздела «Ответственность»)',
    expected: 'Раздел 6 «Ответственность» обязателен',
    diff: 'В Положении об информационном центре отсутствует раздел «Ответственность», обязательный для организационных положений.'
  },
  {
    id: 12, category: 'structure', severity: 'major',
    rule: 'Порядок разделов',
    location: 'Приложение 1, разделы 4–5',
    target: 'Раздел 4: «Отчётность» → Раздел 5: «Функции»',
    expected: 'Раздел 4: «Функции» → Раздел 5: «Отчётность»',
    diff: 'Разделы 4 и 5 переставлены местами. «Функции» должны предшествовать «Отчётности» согласно типовой структуре положений.'
  },
  {
    id: 13, category: 'structure', severity: 'critical',
    rule: 'Ссылка на несуществующее приложение',
    location: 'стр. 2, п.3',
    target: '«...согласно Приложению 3»',
    expected: 'Приложение 3 отсутствует (документ содержит только Прил. 1 и 2)',
    diff: 'В тексте приказа есть ссылка на Приложение 3, которое отсутствует в документе. Это создаёт юридическую неопределённость.'
  },
  {
    id: 14, category: 'structure', severity: 'minor',
    rule: 'Лист ознакомления',
    location: 'последняя страница',
    target: 'Заголовок: «Лист ознакомления»',
    expected: '«Лист ознакомления с Приказом № ___ от ___»',
    diff: 'В листе ознакомления не указаны номер и дата приказа, к которому он относится.'
  },

  // --- Формулировки (3) ---
  {
    id: 15, category: 'clarity', severity: 'major',
    rule: 'Неконкретный срок',
    location: 'стр. 2, п.2',
    target: '«...выполнить в кратчайшие сроки»',
    expected: '«...выполнить до 15.02.2026» (конкретная дата)',
    diff: 'Формулировка «в кратчайшие сроки» не является юридически определённой. Необходимо указать конкретную дату исполнения.'
  },
  {
    id: 16, category: 'clarity', severity: 'critical',
    rule: 'Незаполненное поле ответственного',
    location: 'стр. 2, п.4',
    target: '«Ответственный: ____________________»',
    expected: '«Ответственный: Иванов И.И., начальник отдела»',
    diff: 'Поле ответственного лица не заполнено (содержит метку заполнения). Приказ без указания ответственного лица неисполним.'
  },
  {
    id: 17, category: 'clarity', severity: 'minor',
    rule: 'Канцелярская избыточность',
    location: 'стр. 2, п.1',
    target: '«В целях обеспечения надлежащего исполнения мероприятий по реализации задач в рамках...»',
    expected: '«Для выполнения задач по...»',
    diff: 'Избыточная канцелярская конструкция затрудняет понимание. Рекомендуется упростить формулировку.'
  },

  // --- Элементы (3) ---
  {
    id: 18, category: 'completeness', severity: 'major',
    rule: 'Подпись руководителя',
    location: 'стр. 2, блок подписи',
    target: 'Подпись: отсутствует',
    expected: 'Подпись + расшифровка: «Генеральный директор ___ /И.И. Фамилия/»',
    diff: 'Отсутствует подпись руководителя. Без подписи уполномоченного лица приказ не вступает в силу.'
  },
  {
    id: 19, category: 'completeness', severity: 'minor',
    rule: 'Место издания',
    location: 'стр. 1, шапка',
    target: 'Город: не указан',
    expected: '«г. Москва» (или иной город)',
    diff: 'Не указано место издания приказа (город). Реквизит обязателен по ГОСТ Р 7.0.97-2016.'
  },
  {
    id: 20, category: 'completeness', severity: 'minor',
    rule: 'Приложение без ссылки на приказ',
    location: 'стр. 3, заголовок Приложения 1',
    target: '«Приложение 1»',
    expected: '«Приложение 1 к Приказу № ___ от ___»',
    diff: 'В заголовке приложения нет ссылки на номер и дату приказа. Приложение должно однозначно идентифицировать основной документ.'
  },
];


// --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

// Генерация CSV с BOM и разделителем ; для русского Excel
const generateCSV = () => {
  const BOM = '\uFEFF';
  const header = 'ID;Категория;Важность;Правило;Расположение;В документе;Ожидается;Пояснение';
  const rows = DEMO_VIOLATIONS.map(v => {
    const cat = ANALYSIS_CATEGORIES.find(c => c.id === v.category);
    return [
      v.id,
      cat?.name || v.category,
      SEVERITY_CONFIG[v.severity]?.label || v.severity,
      v.rule,
      v.location,
      `"${v.target.replace(/"/g, '""')}"`,
      `"${v.expected.replace(/"/g, '""')}"`,
      `"${v.diff.replace(/"/g, '""')}"`
    ].join(';');
  });
  return BOM + header + '\n' + rows.join('\n');
};

// Генерация JSON-отчёта
const generateJSON = () => {
  const report = {
    document: DEMO_DOCUMENT,
    categories: ANALYSIS_CATEGORIES.map(c => ({ id: c.id, name: c.name })),
    violations: DEMO_VIOLATIONS,
    summary: {
      total: DEMO_DOCUMENT.totalViolations,
      bySeverity: {
        critical: DEMO_VIOLATIONS.filter(v => v.severity === 'critical').length,
        major: DEMO_VIOLATIONS.filter(v => v.severity === 'major').length,
        minor: DEMO_VIOLATIONS.filter(v => v.severity === 'minor').length,
      },
      byCategory: ANALYSIS_CATEGORIES.map(c => ({
        id: c.id,
        name: c.name,
        count: DEMO_VIOLATIONS.filter(v => v.category === c.id).length,
      })),
    },
    generatedAt: new Date().toISOString(),
  };
  return JSON.stringify(report, null, 2);
};

// Скачивание файла
const downloadFile = (content, filename, mimeType) => {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};


// --- КОМПОНЕНТЫ ---

// Хедер приложения
const Header = () => (
  <header className="bg-white border-b border-gray-200 py-4 px-6 flex items-center justify-between sticky top-0 z-50">
    <div className="flex items-center gap-3">
      <div className="w-9 h-9 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold text-sm">AI</div>
      <span className="font-bold text-xl text-gray-800 tracking-tight">ИИ-АУДИТ ДОКУМЕНТОВ</span>
    </div>
    <div className="text-sm text-gray-500">МосМониторинг</div>
  </header>
);


// =============================================
// ЭКРАН 1: LANDING (приветствие + загрузка)
// =============================================
const ScreenLanding = ({ onFileAccepted }) => {
  const [uploadPhase, setUploadPhase] = useState('idle'); // idle | uploading | done
  const [uploadTimer, setUploadTimer] = useState(0);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);
  const timerRef = useRef(null);

  // Запуск имитации загрузки
  const startUpload = useCallback((file) => {
    setUploadPhase('uploading');
    setUploadTimer(0);

    const startTime = Date.now();
    const uploadDuration = DEMO_DOCUMENT.uploadTime * 1000; // 2100ms

    timerRef.current = setInterval(() => {
      const elapsed = Date.now() - startTime;
      const progress = Math.min(elapsed / uploadDuration, 1);
      setUploadTimer(+(progress * DEMO_DOCUMENT.uploadTime).toFixed(1));

      if (elapsed >= uploadDuration) {
        clearInterval(timerRef.current);
        setUploadTimer(DEMO_DOCUMENT.uploadTime);
        setUploadPhase('done');

        // Автопереход через 0.5с
        setTimeout(() => {
          onFileAccepted();
        }, 500);
      }
    }, 50);
  }, [onFileAccepted]);

  useEffect(() => {
    return () => {
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, []);

  // Обработчики drag-and-drop
  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const handleDragLeave = () => setIsDragging(false);
  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (uploadPhase !== 'idle') return;
    if (e.dataTransfer.files?.[0]) {
      startUpload(e.dataTransfer.files[0]);
    }
  };
  const handleFileSelect = (e) => {
    if (e.target.files?.[0]) {
      startUpload(e.target.files[0]);
    }
  };

  return (
    <div className="max-w-4xl mx-auto py-12 px-4 animate-fade-in">
      {/* Приветствие */}
      <div className="text-center mb-10">
        <h1 className="text-4xl font-bold text-gray-900 mb-3">
          ИИ-ассистент аудита документов
        </h1>
        <p className="text-lg text-gray-500 max-w-2xl mx-auto">
          Загрузите документ — система проверит его по 5 ключевым критериям за секунды
        </p>
      </div>

      {/* Ряд критериев */}
      <div className="flex flex-wrap justify-center gap-4 mb-12">
        {ANALYSIS_CATEGORIES.map((cat) => {
          const Icon = cat.icon;
          const colors = CATEGORY_COLORS[cat.color];
          return (
            <div
              key={cat.id}
              className={`flex items-center gap-2 px-4 py-2 rounded-full ${colors.bg} ${colors.text} text-sm font-medium`}
            >
              <Icon className="w-4 h-4" />
              {cat.shortName}
            </div>
          );
        })}
      </div>

      {/* Зона загрузки */}
      <div className="max-w-2xl mx-auto">
        {uploadPhase === 'idle' && (
          <div
            className={`
              border-2 border-dashed rounded-2xl bg-gray-50 cursor-pointer
              flex flex-col items-center justify-center text-center h-64
              transition-all duration-200
              ${isDragging
                ? 'border-blue-500 bg-blue-50 scale-[1.02]'
                : 'border-gray-300 hover:border-blue-400 hover:bg-blue-50'}
            `}
            onClick={() => fileInputRef.current?.click()}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <Upload className="w-12 h-12 text-gray-400 mb-4" />
            <span className="text-lg font-medium text-gray-700">Перетащите документ сюда</span>
            <span className="text-gray-400 text-sm mt-2">или нажмите для выбора (.docx)</span>
            <input
              ref={fileInputRef}
              type="file"
              accept=".docx"
              className="hidden"
              onChange={handleFileSelect}
            />
          </div>
        )}

        {uploadPhase === 'uploading' && (
          <div className="border-2 border-blue-200 bg-blue-50 rounded-2xl flex flex-col items-center justify-center h-64 animate-fade-in">
            <FileText className="w-12 h-12 text-blue-600 mb-3" />
            <span className="font-bold text-gray-800 text-lg mb-1">{DEMO_DOCUMENT.filename}</span>
            <span className="text-sm text-gray-500 mb-4">{DEMO_DOCUMENT.fileSize}</span>

            {/* Прогресс-бар загрузки */}
            <div className="w-64 bg-blue-200 rounded-full h-2 mb-3 overflow-hidden">
              <div
                className="bg-blue-600 h-full rounded-full transition-all duration-100"
                style={{ width: `${(uploadTimer / DEMO_DOCUMENT.uploadTime) * 100}%` }}
              />
            </div>
            <span className="text-blue-600 font-medium text-sm">
              Загрузка... {uploadTimer}с / {DEMO_DOCUMENT.uploadTime}с
            </span>
          </div>
        )}

        {uploadPhase === 'done' && (
          <div className="border-2 border-green-300 bg-green-50 rounded-2xl flex flex-col items-center justify-center h-64 animate-fade-in">
            <CheckCircle className="w-14 h-14 text-green-600 mb-3" />
            <span className="font-bold text-green-800 text-lg">Файл загружен за {DEMO_DOCUMENT.uploadTime} сек</span>
            <span className="text-green-600 text-sm mt-2">Переход к анализу...</span>
          </div>
        )}
      </div>
    </div>
  );
};


// =============================================
// ЭКРАН 2: PROCESSING (ИИ-анализ по 5 критериям)
// =============================================
const ScreenProcessing = ({ onComplete }) => {
  const [elapsedTime, setElapsedTime] = useState(0);
  const [currentStageIndex, setCurrentStageIndex] = useState(0);
  const [stageStatuses, setStageStatuses] = useState(
    // 'waiting' | 'analyzing' | 'done'
    STAGES_TIMING.map(() => 'waiting')
  );
  const [foundErrors, setFoundErrors] = useState(0);

  const totalDuration = DEMO_DOCUMENT.analysisTime * 1000; // 12400ms
  const timerRef = useRef(null);
  const stageTimeoutsRef = useRef([]);

  useEffect(() => {
    // Живой таймер каждые 50ms
    const startTime = Date.now();
    timerRef.current = setInterval(() => {
      const elapsed = Date.now() - startTime;
      setElapsedTime(Math.min(elapsed / 1000, DEMO_DOCUMENT.analysisTime));
    }, 50);

    // Цепочка этапов
    let accumulatedDelay = 0;
    STAGES_TIMING.forEach((stage, idx) => {
      // Начало этапа
      const startTimeout = setTimeout(() => {
        setCurrentStageIndex(idx);
        setStageStatuses(prev => {
          const next = [...prev];
          next[idx] = 'analyzing';
          return next;
        });
      }, accumulatedDelay);
      stageTimeoutsRef.current.push(startTimeout);

      accumulatedDelay += stage.duration;

      // Завершение этапа
      const endTimeout = setTimeout(() => {
        setStageStatuses(prev => {
          const next = [...prev];
          next[idx] = 'done';
          return next;
        });
        // Нарастающий счётчик ошибок
        setFoundErrors(prev => prev + stage.errorsFound);
      }, accumulatedDelay);
      stageTimeoutsRef.current.push(endTimeout);
    });

    // Завершение: пауза 0.8с и переход
    const completeTimeout = setTimeout(() => {
      onComplete();
    }, accumulatedDelay + 800);
    stageTimeoutsRef.current.push(completeTimeout);

    return () => {
      clearInterval(timerRef.current);
      stageTimeoutsRef.current.forEach(t => clearTimeout(t));
    };
  }, [onComplete]);

  // Прогресс общий (0-100%)
  const progressPercent = Math.min((elapsedTime / DEMO_DOCUMENT.analysisTime) * 100, 100);

  return (
    <div className="max-w-3xl mx-auto py-12 px-4 animate-fade-in">
      {/* Заголовок */}
      <div className="text-center mb-8">
        <h2 className="text-2xl font-bold text-gray-900 mb-1">ИИ-анализ документа</h2>
        <p className="text-gray-500 font-mono text-sm">{DEMO_DOCUMENT.filename}</p>
      </div>

      {/* Общий прогресс */}
      <div className="bg-white rounded-2xl border border-gray-200 shadow-sm p-6 mb-6">
        {/* Прогресс-бар */}
        <div className="bg-gray-200 rounded-full h-3 w-full mb-4 overflow-hidden">
          <div
            className="bg-blue-600 h-full rounded-full transition-all duration-100"
            style={{ width: `${progressPercent}%` }}
          />
        </div>

        {/* Таймер и счётчик */}
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2 text-gray-600">
            <Clock className="w-4 h-4" />
            <span className="font-mono text-sm">
              {elapsedTime.toFixed(1)}с / {DEMO_DOCUMENT.analysisTime}с
            </span>
          </div>
          <div className="flex items-center gap-2">
            <AlertCircle className="w-5 h-5 text-red-500" />
            <span className="text-lg font-bold text-red-600 animate-count-pulse" key={foundErrors}>
              Найдено проблем: {foundErrors}
            </span>
          </div>
        </div>
      </div>

      {/* Список критериев */}
      <div className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden">
        {STAGES_TIMING.map((stage, idx) => {
          const cat = ANALYSIS_CATEGORIES.find(c => c.id === stage.categoryId);
          const Icon = cat.icon;
          const status = stageStatuses[idx];
          const colors = CATEGORY_COLORS[cat.color];

          return (
            <div
              key={cat.id}
              className={`flex items-center px-6 py-4 border-b border-gray-100 last:border-b-0 transition-colors duration-300 ${
                status === 'analyzing' ? 'bg-blue-50' : ''
              }`}
            >
              {/* Иконка категории */}
              <div className={`w-10 h-10 rounded-lg flex items-center justify-center mr-4 ${colors.bg}`}>
                <Icon className={`w-5 h-5 ${colors.text}`} />
              </div>

              {/* Название */}
              <div className="flex-1">
                <span className={`font-medium ${status === 'analyzing' ? 'text-gray-900' : status === 'done' ? 'text-gray-700' : 'text-gray-400'}`}>
                  {cat.name}
                </span>
              </div>

              {/* Статус */}
              <div className="ml-4">
                {status === 'waiting' && (
                  <span className="text-sm text-gray-400">Ожидание</span>
                )}
                {status === 'analyzing' && (
                  <span className="flex items-center gap-2 text-sm text-blue-600 font-medium">
                    <div className="w-2 h-2 bg-blue-600 rounded-full animate-pulse-dot" />
                    Анализ...
                  </span>
                )}
                {status === 'done' && stage.errorsFound > 0 && (
                  <span className={`inline-flex items-center px-3 py-1 rounded-full text-sm font-bold ${
                    stage.errorsFound >= 5 ? 'bg-red-100 text-red-700' : 'bg-orange-100 text-orange-700'
                  }`}>
                    {stage.errorsFound} {stage.errorsFound === 1 ? 'проблема' : stage.errorsFound < 5 ? 'проблемы' : 'проблем'}
                  </span>
                )}
                {status === 'done' && stage.errorsFound === 0 && (
                  <span className="flex items-center gap-1 text-sm text-green-600 font-medium">
                    <CheckCircle className="w-4 h-4" /> OK
                  </span>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};


// =============================================
// ЭКРАН 3: RESULTS (сводка + детали + скачивание)
// =============================================
const ScreenResults = ({ onReset }) => {
  const [activeFilter, setActiveFilter] = useState('all');

  // Фильтрация нарушений
  const filteredViolations = activeFilter === 'all'
    ? DEMO_VIOLATIONS
    : DEMO_VIOLATIONS.filter(v => v.category === activeFilter);

  // Подсчёты severity
  const severityCounts = {
    critical: DEMO_VIOLATIONS.filter(v => v.severity === 'critical').length,
    major: DEMO_VIOLATIONS.filter(v => v.severity === 'major').length,
    minor: DEMO_VIOLATIONS.filter(v => v.severity === 'minor').length,
  };

  // Подсчёт по категориям
  const categoryCounts = ANALYSIS_CATEGORIES.map(c => ({
    ...c,
    count: DEMO_VIOLATIONS.filter(v => v.category === c.id).length,
  }));

  const maxCount = Math.max(...categoryCounts.map(c => c.count));

  return (
    <div className="max-w-5xl mx-auto py-10 px-4 animate-fade-in">
      {/* Заголовок */}
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Результаты проверки</h1>
          <p className="text-gray-500 mt-1">
            Документ: <span className="font-mono text-gray-700">{DEMO_DOCUMENT.filename}</span>
          </p>
        </div>
        <div className="text-right">
          <div className="text-sm text-gray-400">Дата проверки</div>
          <div className="font-medium">
            {new Date().toLocaleDateString('ru-RU')}{' '}
            {new Date().toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' })}
          </div>
        </div>
      </div>

      {/* Секция A: Сводная статистика */}

      {/* 3 карточки-метрики */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <div className="bg-white p-5 rounded-xl border-l-4 border-red-500 shadow-sm flex items-center gap-4">
          <AlertCircle className="w-8 h-8 text-red-500 shrink-0" />
          <div>
            <div className="text-3xl font-bold text-red-600">{DEMO_DOCUMENT.totalViolations}</div>
            <div className="text-sm text-gray-500 font-medium">проблем найдено</div>
          </div>
        </div>
        <div className="bg-white p-5 rounded-xl border-l-4 border-blue-500 shadow-sm flex items-center gap-4">
          <CheckCircle className="w-8 h-8 text-blue-500 shrink-0" />
          <div>
            <div className="text-3xl font-bold text-blue-600">5</div>
            <div className="text-sm text-gray-500 font-medium">критериев проверено</div>
          </div>
        </div>
        <div className="bg-white p-5 rounded-xl border-l-4 border-gray-400 shadow-sm flex items-center gap-4">
          <Clock className="w-8 h-8 text-gray-400 shrink-0" />
          <div>
            <div className="text-3xl font-bold text-gray-700">{DEMO_DOCUMENT.analysisTime} сек</div>
            <div className="text-sm text-gray-500 font-medium">время анализа</div>
          </div>
        </div>
      </div>

      {/* Bar-chart по категориям */}
      <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-4">
        <div className="flex items-center gap-2 mb-4">
          <BarChart3 className="w-5 h-5 text-gray-500" />
          <h3 className="font-bold text-gray-700">Распределение по категориям</h3>
        </div>
        <div className="space-y-3">
          {categoryCounts.map(cat => {
            const colors = CATEGORY_COLORS[cat.color];
            const widthPercent = maxCount > 0 ? (cat.count / maxCount) * 100 : 0;
            return (
              <div key={cat.id} className="flex items-center gap-3">
                <span className="text-sm text-gray-600 w-48 shrink-0 text-right">{cat.shortName}</span>
                <div className="flex-1 bg-gray-100 rounded-full h-6 overflow-hidden">
                  <div
                    className={`${colors.bar} h-full rounded-full transition-all duration-700`}
                    style={{ width: `${widthPercent}%` }}
                  />
                </div>
                <span className="text-sm font-bold text-gray-700 w-6 text-right">{cat.count}</span>
              </div>
            );
          })}
        </div>
      </div>

      {/* Severity-строка */}
      <div className="flex flex-wrap gap-3 mb-8">
        {Object.entries(severityCounts).map(([key, count]) => {
          const cfg = SEVERITY_CONFIG[key];
          return (
            <div key={key} className={`flex items-center gap-2 px-4 py-2 rounded-lg ${cfg.bg} ${cfg.border} border`}>
              <div className={`w-2.5 h-2.5 rounded-full ${cfg.dot}`} />
              <span className={`text-sm font-bold ${cfg.text}`}>{count} {cfg.labelFull}</span>
            </div>
          );
        })}
      </div>

      {/* Секция B: Детальный список */}

      {/* Фильтр-табы */}
      <div className="flex items-center gap-2 mb-6 overflow-x-auto pb-2">
        <Filter className="w-4 h-4 text-gray-400 shrink-0" />
        <button
          onClick={() => setActiveFilter('all')}
          className={`px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition-colors ${
            activeFilter === 'all'
              ? 'bg-gray-800 text-white'
              : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
          }`}
        >
          Все ({DEMO_VIOLATIONS.length})
        </button>
        {categoryCounts.map(cat => {
          const isActive = activeFilter === cat.id;
          const colors = CATEGORY_COLORS[cat.color];
          return (
            <button
              key={cat.id}
              onClick={() => setActiveFilter(cat.id)}
              className={`px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition-colors ${
                isActive
                  ? `${colors.bar} text-white`
                  : `bg-gray-100 text-gray-600 hover:bg-gray-200`
              }`}
            >
              {cat.shortName} ({cat.count})
            </button>
          );
        })}
      </div>

      {/* Карточки ошибок */}
      <div className="space-y-4 mb-10">
        {filteredViolations.map((v, idx) => {
          const cat = ANALYSIS_CATEGORIES.find(c => c.id === v.category);
          const Icon = cat.icon;
          const catColors = CATEGORY_COLORS[cat.color];
          const sevCfg = SEVERITY_CONFIG[v.severity];

          return (
            <div
              key={v.id}
              className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm hover:shadow-md transition-shadow animate-slide-up"
              style={{ animationDelay: `${idx * 50}ms` }}
            >
              {/* Шапка карточки */}
              <div className="px-6 py-3 border-b border-gray-100 flex items-center gap-3 flex-wrap">
                <span className="text-sm font-mono text-gray-400">#{v.id}</span>
                <span className={`inline-flex items-center px-2.5 py-0.5 rounded text-xs font-bold ${sevCfg.bg} ${sevCfg.text}`}>
                  {sevCfg.label}
                </span>
                <span className="font-bold text-gray-900">{v.rule}</span>
                <span className={`ml-auto inline-flex items-center gap-1.5 text-xs font-medium ${catColors.text}`}>
                  <Icon className="w-3.5 h-3.5" />
                  {cat.shortName}
                </span>
              </div>

              {/* Расположение */}
              <div className="px-6 pt-3 pb-1">
                <span className="text-xs text-gray-400 font-medium uppercase">Расположение: </span>
                <span className="text-sm text-gray-600">{v.location}</span>
              </div>

              {/* Два столбца: в документе / ожидается */}
              <div className="px-6 py-3 grid md:grid-cols-2 gap-4">
                <div>
                  <div className="text-xs font-bold text-gray-400 uppercase mb-1.5">В документе</div>
                  <div className="bg-red-50 p-3 rounded-lg border border-red-100 text-sm text-red-800 font-mono whitespace-pre-wrap">
                    {v.target}
                  </div>
                </div>
                <div>
                  <div className="text-xs font-bold text-green-500 uppercase mb-1.5">Ожидается</div>
                  <div className="bg-green-50 p-3 rounded-lg border border-green-100 text-sm text-green-800 font-mono whitespace-pre-wrap">
                    {v.expected}
                  </div>
                </div>
              </div>

              {/* Пояснение */}
              <div className="px-6 pb-4">
                <div className="text-xs font-bold text-gray-400 uppercase mb-1.5">Пояснение</div>
                <div className="bg-gray-50 p-3 rounded-lg border border-gray-100 text-sm text-gray-700">
                  {v.diff}
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {/* Секция C: Кнопки действий */}
      <div className="flex flex-col md:flex-row gap-4 justify-between border-t border-gray-200 pt-8">
        <button
          onClick={onReset}
          className="flex items-center justify-center gap-2 px-6 py-3 bg-gray-100 text-gray-700 font-bold rounded-lg hover:bg-gray-200 transition-colors"
        >
          <RefreshCw className="w-4 h-4" /> Проверить другой документ
        </button>

        <div className="flex gap-4">
          <button
            onClick={() => {
              const json = generateJSON();
              downloadFile(json, 'audit_report.json', 'application/json');
            }}
            className="flex items-center justify-center gap-2 px-6 py-3 border border-gray-300 text-gray-700 font-bold rounded-lg hover:bg-gray-50 transition-colors"
          >
            <Download className="w-4 h-4" /> Экспорт JSON
          </button>
          <button
            onClick={() => {
              const csv = generateCSV();
              downloadFile(csv, 'audit_report.csv', 'text/csv;charset=utf-8');
            }}
            className="flex items-center justify-center gap-2 px-6 py-3 bg-green-600 text-white font-bold rounded-lg hover:bg-green-700 transition-colors shadow-sm"
          >
            <Download className="w-4 h-4" /> Скачать отчёт
          </button>
        </div>
      </div>
    </div>
  );
};


// --- ГЛАВНЫЙ КОМПОНЕНТ ---

function App() {
  const [screen, setScreen] = useState('landing'); // landing | processing | results

  const handleFileAccepted = useCallback(() => {
    setScreen('processing');
  }, []);

  const handleProcessComplete = useCallback(() => {
    setScreen('results');
  }, []);

  const handleReset = useCallback(() => {
    setScreen('landing');
  }, []);

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-900 pb-20">
      <Header />

      <main>
        {screen === 'landing' && (
          <ScreenLanding onFileAccepted={handleFileAccepted} />
        )}

        {screen === 'processing' && (
          <ScreenProcessing onComplete={handleProcessComplete} />
        )}

        {screen === 'results' && (
          <ScreenResults onReset={handleReset} />
        )}
      </main>
    </div>
  );
}

export default App;
