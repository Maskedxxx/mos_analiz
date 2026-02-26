import { useState } from 'react';
import {
  AlertCircle, CheckCircle, Clock, Download, RefreshCw, ChevronDown, ChevronUp,
} from 'lucide-react';
import { getDownloadUrl } from './api';

export default function ScreenResults({ result, sessionId, filename, onReset }) {
  const [expandedIdx, setExpandedIdx] = useState(null);

  const violations = result.violations || [];
  const rulesChecked = result.rules_checked || 0;
  const duration = result.duration_sec || 0;
  const failedRules = [...new Set(violations.map((v) => v.rule_index))].length;

  return (
    <div className="animate-fade-in max-w-3xl mx-auto px-4 py-12 space-y-8">
      {/* Заголовок */}
      <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold text-gray-900">Результат проверки</h2>
        <p className="text-gray-500 text-sm truncate">{filename}</p>
      </div>

      {/* 3 метрики */}
      <div className="grid grid-cols-3 gap-4">
        <MetricCard
          label="Проверено правил"
          value={rulesChecked}
          color="blue"
          icon={<CheckCircle className="w-5 h-5" />}
        />
        <MetricCard
          label="Замечания"
          value={violations.length}
          color={violations.length > 0 ? 'red' : 'green'}
          icon={<AlertCircle className="w-5 h-5" />}
        />
        <MetricCard
          label="Время"
          value={`${duration.toFixed(1)}с`}
          color="gray"
          icon={<Clock className="w-5 h-5" />}
        />
      </div>

      {/* Сводка */}
      {violations.length === 0 ? (
        <div className="bg-green-50 border border-green-200 rounded-xl p-6 text-center">
          <CheckCircle className="w-10 h-10 text-green-500 mx-auto mb-2" />
          <p className="text-green-800 font-medium">Все проверки пройдены!</p>
          <p className="text-green-600 text-sm mt-1">Замечаний не обнаружено</p>
        </div>
      ) : (
        <div className="space-y-3">
          <h3 className="text-lg font-semibold text-gray-900">
            Замечания ({violations.length})
          </h3>

          {violations.map((v, idx) => {
            const isExpanded = expandedIdx === idx;
            return (
              <div
                key={idx}
                className="animate-slide-up bg-white border border-gray-200 rounded-xl overflow-hidden"
                style={{ animationDelay: `${idx * 40}ms` }}
              >
                {/* Шапка карточки */}
                <button
                  onClick={() => setExpandedIdx(isExpanded ? null : idx)}
                  className="w-full px-5 py-4 flex items-center gap-3 text-left hover:bg-gray-50 transition-colors cursor-pointer"
                >
                  <span className="text-xs font-mono bg-red-100 text-red-700 px-2 py-0.5 rounded-md shrink-0">
                    #{v.rule_index}
                  </span>
                  <span className="text-gray-900 font-medium flex-1 truncate">
                    {v.rule_title}
                  </span>
                  {isExpanded ? (
                    <ChevronUp className="w-4 h-4 text-gray-400 shrink-0" />
                  ) : (
                    <ChevronDown className="w-4 h-4 text-gray-400 shrink-0" />
                  )}
                </button>

                {/* Тело карточки */}
                {isExpanded && (
                  <div className="px-5 pb-4 space-y-3 border-t border-gray-100 pt-3">
                    {v['Целевой документ'] && (
                      <div>
                        <p className="text-xs text-gray-400 uppercase mb-1">В документе</p>
                        <p className="text-sm text-gray-700 bg-red-50 rounded-lg px-3 py-2 whitespace-pre-wrap">
                          {v['Целевой документ']}
                        </p>
                      </div>
                    )}
                    {v['Различие'] && (
                      <div>
                        <p className="text-xs text-gray-400 uppercase mb-1">Замечание</p>
                        <p className="text-sm text-gray-700 bg-gray-50 rounded-lg px-3 py-2 whitespace-pre-wrap">
                          {v['Различие']}
                        </p>
                      </div>
                    )}
                    {/* Дополнительные поля */}
                    {Object.entries(v)
                      .filter(([k]) => !['rule_index', 'rule_title', 'Целевой документ', 'Различие'].includes(k))
                      .map(([k, val]) => (
                        <div key={k}>
                          <p className="text-xs text-gray-400 uppercase mb-1">{k}</p>
                          <p className="text-sm text-gray-700 bg-gray-50 rounded-lg px-3 py-2 whitespace-pre-wrap">
                            {typeof val === 'object' ? JSON.stringify(val, null, 2) : String(val)}
                          </p>
                        </div>
                      ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Кнопки */}
      <div className="flex gap-3">
        <button
          onClick={onReset}
          className="flex-1 py-3 border border-gray-300 text-gray-700 hover:bg-gray-100 font-medium rounded-xl flex items-center justify-center gap-2 transition-colors cursor-pointer"
        >
          <RefreshCw className="w-5 h-5" />
          Новый аудит
        </button>
        <a
          href={getDownloadUrl(sessionId)}
          download
          className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-xl flex items-center justify-center gap-2 transition-colors"
        >
          <Download className="w-5 h-5" />
          Скачать Excel
        </a>
      </div>
    </div>
  );
}

function MetricCard({ label, value, color, icon }) {
  const colors = {
    blue:  'border-l-blue-500 text-blue-600',
    red:   'border-l-red-500 text-red-600',
    green: 'border-l-green-500 text-green-600',
    gray:  'border-l-gray-400 text-gray-600',
  };

  return (
    <div className={`bg-white rounded-xl border border-gray-200 border-l-4 ${colors[color]} px-4 py-4`}>
      <div className="flex items-center gap-2 mb-1 opacity-70">{icon}<span className="text-xs uppercase">{label}</span></div>
      <p className="text-2xl font-bold">{value}</p>
    </div>
  );
}
