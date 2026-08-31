/**
 * Левое меню платформы: переключение режимов «Аудит» / «Генерация».
 * На узких экранах превращается в горизонтальные пилюли над контентом.
 */
import { ClipboardCheck, FileSignature } from 'lucide-react';

// Высота шапки (py-4 + логотип 36px + бордер) — на неё липнет сайдбар
const HEADER_H = '69px';

const ITEMS = [
  { key: 'audit', label: 'Аудит', hint: 'Проверка документов', Icon: ClipboardCheck },
  { key: 'generation', label: 'Генерация', hint: 'Создание документов', Icon: FileSignature },
];

export default function Sidebar({ mode, onSelect }) {
  return (
    <aside
      className="shrink-0 bg-white border-b border-gray-200 md:border-b-0 md:border-r md:w-56 md:sticky md:self-start md:h-[calc(100vh-69px)]"
      style={{ top: HEADER_H }}
    >
      <nav className="flex md:flex-col gap-1 p-3">
        {ITEMS.map((item) => {
          const { key, label, hint } = item;
          const active = mode === key;
          return (
            <button
              key={key}
              type="button"
              onClick={() => onSelect(key)}
              className={`relative flex-1 md:flex-none flex items-center gap-3 rounded-xl px-3 py-2.5 text-left transition-colors cursor-pointer ${
                active ? 'bg-blue-50 text-blue-700' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              {active && <span className="absolute left-0 top-2 bottom-2 w-1 rounded-r-full bg-blue-600" />}
              <item.Icon className={`w-5 h-5 shrink-0 ${active ? 'text-blue-600' : 'text-gray-400'}`} />
              <span className="min-w-0">
                <span className={`block text-sm ${active ? 'font-semibold' : 'font-medium'}`}>{label}</span>
                <span className="hidden md:block text-xs text-gray-400 truncate">{hint}</span>
              </span>
            </button>
          );
        })}
      </nav>
    </aside>
  );
}
