import { useState, useEffect, useRef } from 'react';
import {
  FileText, ScanLine, LayoutTemplate, ClipboardCheck, FileSpreadsheet,
  Loader2, CheckCircle, Clock, AlertCircle, XCircle,
} from 'lucide-react';
import { subscribeToProgress, fetchAuditResult } from './api';

const STEPS = [
  { id: 'convert',    label: 'Конвертация документа',    icon: FileText },
  { id: 'ocr_target', label: 'OCR целевого документа',   icon: ScanLine },
  { id: 'ocr_tpl',    label: 'Загрузка шаблона',         icon: LayoutTemplate },
  { id: 'rules',      label: 'Проверка правил',          icon: ClipboardCheck },
  { id: 'report',     label: 'Формирование отчёта',      icon: FileSpreadsheet },
];

export default function ScreenProgress({ sessionId, filename, onComplete, onError }) {
  // 'waiting' | 'active' | 'done' | 'error'
  const [stepStatuses, setStepStatuses] = useState(STEPS.map(() => 'waiting'));
  const [rulesProgress, setRulesProgress] = useState({ current: 0, total: 0 });
  const [elapsed, setElapsed] = useState(0);
  const [errorMsg, setErrorMsg] = useState('');
  const [errorTech, setErrorTech] = useState('');   // исходный текст исключения — под «Подробности»
  const [queuePosition, setQueuePosition] = useState(0);
  const sourceRef = useRef(null);
  const startRef = useRef(Date.now());
  const timerRef = useRef(null);
  // Метка последней активности потока (событие/ping) — для сторожевого таймера.
  const lastActivityRef = useRef(Date.now());
  const watchdogRef = useRef(null);

  useEffect(() => {
    let cancelled = false;
    timerRef.current = setInterval(() => {
      setElapsed((Date.now() - startRef.current) / 1000);
    }, 100);

    // Показать результат завершённого аудита и остановить таймеры.
    const finishWith = (data) => {
      if (cancelled) return;
      clearInterval(timerRef.current);
      setStepStatuses(STEPS.map(() => 'done'));
      setTimeout(() => { if (!cancelled) onComplete(data); }, 300);
    };

    // Обработчик события прогресса. Любое событие (в т.ч. ping) продлевает «жизнь» потока.
    const handleEvent = (type, data) => {
      lastActivityRef.current = Date.now();
      if (type === 'ping') return;
      // Обработка очереди — вне setStepStatuses, т.к. не меняет шаги
      if (type === 'queue') {
        setQueuePosition(data.position || 0);
        return;
      }

      setStepStatuses((prev) => {
        const next = [...prev];

        if (type === 'audit_start') {
          next[0] = 'active';
        }
        if (type === 'parsing_target') {
          next[0] = 'done';
          next[1] = 'active';
        }
        if (type === 'parsing_target_done') {
          next[1] = 'done';
          next[2] = 'done';
        }
        if (type === 'checking_rules') {
          next[3] = 'active';
          setRulesProgress({ current: 0, total: data.total || 0 });
        }
        if (type === 'checking_rules_done') {
          next[3] = 'done';
          next[4] = 'active';
        }
        if (type === 'complete') {
          next[4] = 'done';
          clearInterval(timerRef.current);

          setTimeout(() => onComplete(data), 600);
        }
        if (type === 'error') {
          // Помечаем текущий active-шаг как error
          for (let i = 0; i < next.length; i++) {
            if (next[i] === 'active') {
              next[i] = 'error';
              break;
            }
          }
          clearInterval(timerRef.current);

          setErrorMsg(data.message || 'Неизвестная ошибка');
          setErrorTech(data.technical || '');
        }

        return next;
      });
    };

    // Сторожевой таймер: если 60 с нет ни событий, ни ping — спросить /result напрямую.
    // Так ловится «тихий» обрыв через прокси, при котором onerror у EventSource не срабатывает
    // (аудит устойчивости, находки 4.4b / 6.1). 202 — аудит ещё идёт, продлеваем ожидание;
    // 200 — показать результат; 404/500 — соединение потеряно / аудит прерван.
    const checkSilence = async () => {
      if (cancelled || Date.now() - lastActivityRef.current < 60000) return;
      try {
        const { status, data } = await fetchAuditResult(sessionId);
        if (cancelled) return;
        if (status === 200) {
          finishWith(data);
        } else if (status === 202) {
          lastActivityRef.current = Date.now();
        } else {
          clearInterval(timerRef.current);
          setErrorMsg(data.detail || 'Соединение с сервером потеряно, аудит прерван.');
          setErrorTech(data.technical || '');
        }
      } catch {
        // Сеть недоступна — не мигаем ошибкой сразу, ждём следующий цикл.
      }
    };
    watchdogRef.current = setInterval(checkSilence, 10000);

    // Немедленная проба результата. Нужна при восстановлении после перезагрузки страницы (F5):
    // если аудит уже завершился, показываем результат сразу; если сессия потеряна — ошибку;
    // при обычном старте /result ещё выполняется (202) — подписываемся на прогресс.
    fetchAuditResult(sessionId).then(({ status, data }) => {
      if (cancelled) return;
      if (status === 200) {
        finishWith(data);
        return;
      }
      if (status === 404) {
        clearInterval(timerRef.current);
        setErrorMsg('Аудит не найден: сессия недоступна. Запустите проверку заново.');
        return;
      }
      if (status === 500) {
        clearInterval(timerRef.current);
        setErrorMsg(data.detail || 'Аудит завершился с ошибкой.');
        setErrorTech(data.technical || '');
        return;
      }
      // 202 или прочее — подписываемся на прогресс.
      lastActivityRef.current = Date.now();
      sourceRef.current = subscribeToProgress(sessionId, handleEvent);
    }).catch(() => {
      // Первичный опрос не удался (сеть) — всё равно подписываемся, обрыв поймает watchdog.
      if (!cancelled) sourceRef.current = subscribeToProgress(sessionId, handleEvent);
    });

    return () => {
      cancelled = true;
      clearInterval(timerRef.current);
      clearInterval(watchdogRef.current);
      sourceRef.current?.close();
    };
  }, [sessionId]);

  // Прогресс 0-100
  const doneCount = stepStatuses.filter((s) => s === 'done').length;
  const activeCount = stepStatuses.filter((s) => s === 'active').length;
  const progress = Math.round(((doneCount + activeCount * 0.5) / STEPS.length) * 100);

  return (
    <div className="animate-fade-in max-w-xl mx-auto px-4 py-12 space-y-8">
      {/* Баннер очереди */}
      {queuePosition > 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex items-center gap-3 animate-fade-in">
          <Clock className="w-5 h-5 text-amber-600 shrink-0" />
          <div>
            <p className="font-medium text-amber-800">
              Идёт проверка другого документа
            </p>
            <p className="text-sm text-amber-600">
              Вы {queuePosition}-{queuePosition === 1 ? 'й' : 'й'} в очереди. Проверка начнётся автоматически.
            </p>
          </div>
        </div>
      )}

      {/* Заголовок */}
      <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold text-gray-900">Анализ документа</h2>
        <p className="text-gray-500 text-sm truncate">{filename}</p>
      </div>

      {/* Stepper */}
      <div className="space-y-0">
        {STEPS.map((step, i) => {
          const status = stepStatuses[i];
          const Icon = step.icon;
          const isLast = i === STEPS.length - 1;

          return (
            <div key={step.id} className="flex items-start gap-4 relative">
              {/* Линия */}
              {!isLast && (
                <div
                  className={`absolute left-5 top-11 w-0.5 h-8 ${
                    status === 'done'
                      ? 'bg-green-400'
                      : status === 'error'
                        ? 'bg-red-300'
                        : 'bg-gray-200'
                  } transition-colors`}
                />
              )}

              {/* Кружок */}
              <div
                className={`w-10 h-10 rounded-full flex items-center justify-center shrink-0 transition-colors ${
                  status === 'done'
                    ? 'bg-green-100'
                    : status === 'active'
                      ? 'bg-blue-100'
                      : status === 'error'
                        ? 'bg-red-100'
                        : 'bg-gray-100'
                }`}
              >
                {status === 'done' ? (
                  <CheckCircle className="w-5 h-5 text-green-600" />
                ) : status === 'active' ? (
                  <Loader2 className="w-5 h-5 text-blue-600 animate-spin" />
                ) : status === 'error' ? (
                  <XCircle className="w-5 h-5 text-red-600" />
                ) : (
                  <Icon className="w-5 h-5 text-gray-400" />
                )}
              </div>

              {/* Текст */}
              <div className="pt-2 pb-8">
                <p className={`font-medium ${
                  status === 'done'
                    ? 'text-green-700'
                    : status === 'active'
                      ? 'text-blue-700'
                      : status === 'error'
                        ? 'text-red-700'
                        : 'text-gray-400'
                }`}>
                  {step.label}
                  {step.id === 'rules' && rulesProgress.total > 0 && (
                    <span className="text-sm font-normal ml-2">
                      {rulesProgress.total} правил
                    </span>
                  )}
                </p>
                <p className="text-xs text-gray-400 mt-0.5">
                  {status === 'done' && 'Завершено'}
                  {status === 'active' && 'Выполняется...'}
                  {status === 'error' && 'Ошибка'}
                  {status === 'waiting' && 'Ожидание'}
                </p>
              </div>
            </div>
          );
        })}
      </div>

      {/* Прогресс-бар */}
      <div className="space-y-2">
        <div className="w-full bg-gray-200 rounded-full h-2">
          <div
            className={`h-2 rounded-full transition-all duration-500 ${
              errorMsg ? 'bg-red-500' : 'bg-blue-600'
            }`}
            style={{ width: `${progress}%` }}
          />
        </div>
        <div className="flex justify-between text-xs text-gray-400">
          <span className="flex items-center gap-1">
            <Clock className="w-3 h-3" />
            {elapsed.toFixed(1)} сек
          </span>
          <span>{progress}%</span>
        </div>
      </div>

      {/* Ошибка */}
      {errorMsg && (
        <div className="bg-red-50 border border-red-200 rounded-xl p-5 space-y-3">
          <div className="flex items-start gap-3">
            <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
            <div>
              <p className="font-bold text-red-800">Ошибка при обработке документа</p>
              <p className="text-sm text-red-600 mt-1">{errorMsg}</p>
              {errorTech && (
                <details className="mt-2 text-xs text-red-500">
                  <summary className="cursor-pointer select-none">Подробности</summary>
                  <pre className="mt-1 whitespace-pre-wrap break-all font-mono">{errorTech}</pre>
                </details>
              )}
            </div>
          </div>
          <button
            onClick={() => onError(errorMsg)}
            className="w-full py-2.5 bg-red-600 text-white font-medium rounded-lg hover:bg-red-700 transition-colors cursor-pointer"
          >
            Вернуться к загрузке
          </button>
        </div>
      )}
    </div>
  );
}
