/**
 * Экран генерации документов.
 * Порт формы mos_generated/static/index.html: выбор типа → схема полей →
 * валидация обязательных → нативное скачивание .docx через скрытый iframe.
 * Плюс автозаполнение полей ранее введёнными значениями по ООО.
 */
import { useState, useEffect, useRef, useCallback } from 'react';
import { ChevronDown, Download, AlertCircle } from 'lucide-react';
import {
  fetchGenTypes, fetchGenSchema, fetchOrgs, fetchOrgSuggest, getGenDownloadUrl,
} from './api';

// Сортировка типов по коду документа (0.1, 0.2 … 3.10) — как в исходной форме генерации
const codeKey = (c) => {
  const p = (c || '99.99').split('.').map(Number);
  return p[0] * 1000 + (p[1] || 0);
};

// Поле-якорь автозаполнения: по нему бэкенд ищет ранее введённые значения
const ORG_KEY = 'org_full';

export default function ScreenGeneration() {
  const [nodes, setNodes] = useState([]);
  const [current, setCurrent] = useState('');
  const [menuOpen, setMenuOpen] = useState(false);
  const [schema, setSchema] = useState(null);
  const [schemaLoading, setSchemaLoading] = useState(false);
  const [values, setValues] = useState({});
  const [orgs, setOrgs] = useState([]);
  const [suggest, setSuggest] = useState({});   // {field_key: [values]} по выбранному ООО
  const [status, setStatus] = useState(null);   // {kind:'ok'|'err', text}
  const [loadError, setLoadError] = useState('');

  // Выбор типа: ставим тип и обнуляем форму (значения, подсказки, статус)
  const selectType = useCallback((docType) => {
    setCurrent(docType);
    setSchema(null);
    setValues({});
    setSuggest({});
    setStatus(null);
    setSchemaLoading(true);
  }, []);

  const ddRef = useRef(null);
  const orgRef = useRef(null);

  // Список типов при монтировании; первый тип выбирается автоматически
  useEffect(() => {
    fetchGenTypes()
      .then((list) => {
        const sorted = list.slice().sort((a, b) => codeKey(a.code) - codeKey(b.code));
        setNodes(sorted);
        if (sorted.length) selectType(sorted[0].doc_type);
      })
      .catch((e) => setLoadError(e.message));
  }, [selectType]);

  // Известные ООО — подсказки для поля org_full
  useEffect(() => {
    fetchOrgs()
      .then((o) => setOrgs(Array.isArray(o) ? o : []))
      .catch(() => {});
  }, []);

  // Смена типа: схема тянется в эффекте, сброс формы делает selectType (см. ниже)
  useEffect(() => {
    if (!current) return undefined;
    let cancelled = false;
    fetchGenSchema(current)
      .then((s) => { if (!cancelled) { setSchema(s); setLoadError(''); } })
      .catch((e) => { if (!cancelled) setLoadError(e.message); })
      .finally(() => { if (!cancelled) setSchemaLoading(false); });
    return () => { cancelled = true; };
  }, [current]);

  // Клик вне выпадающего списка — закрыть
  useEffect(() => {
    const onDocClick = (e) => { if (!ddRef.current?.contains(e.target)) setMenuOpen(false); };
    document.addEventListener('click', onDocClick);
    return () => document.removeEventListener('click', onDocClick);
  }, []);

  // Подтянуть значения по указанному ООО: наполнить подсказки и заполнить ПУСТЫЕ поля
  const applyOrgSuggestions = useCallback(async (orgVal) => {
    if (!orgVal || !orgVal.trim()) return;
    let sug = {};
    try { sug = await fetchOrgSuggest(orgVal); } catch { return; }
    setSuggest(sug || {});
    setValues((prev) => {
      const next = { ...prev };
      Object.entries(sug || {}).forEach(([key, vals]) => {
        if (key === ORG_KEY || !vals || !vals.length) return;
        if (!String(next[key] || '').trim()) next[key] = vals[0];
      });
      return next;
    });
  }, []);

  // Нативный change поля ООО: ловит и выбор из выпадающих подсказок, не только ручной ввод
  useEffect(() => {
    const el = orgRef.current;
    if (!el) return undefined;
    const handler = () => applyOrgSuggestions(el.value);
    el.addEventListener('change', handler);
    return () => el.removeEventListener('change', handler);
  }, [schema, applyOrgSuggestions]);

  const fields = schema?.fields || [];
  const currentNode = nodes.find((n) => n.doc_type === current);
  const setValue = (key, v) => setValues((p) => ({ ...p, [key]: v }));

  // Генерация: проверяем обязательные, затем СКАЧИВАЕМ через fetch и показываем «Готово»
  // только после реального ответа. Скрытый iframe не годился: при 400/500/502 ответ прокси
  // невидим, и статус «Готово» ставился при отсутствии файла (находка 5.7b). Blob-скачивание
  // надёжно сохраняет, если якорь добавлен в DOM до click и объект-URL освобождён с задержкой.
  const generate = async () => {
    if (!current) return;
    const params = { session_id: 'web' };
    const missing = [];
    fields.forEach((f) => {
      const v = String(values[f.key] || '').trim();
      if (v) params[f.key] = v;
      else if (f.required) missing.push(f.label || f.key);
    });
    if (missing.length) {
      setStatus({ kind: 'err', text: 'Заполните обязательные поля: ' + missing.join(', ') });
      return;
    }
    setStatus({ kind: 'ok', text: 'Формируется документ…' });
    try {
      const res = await fetch(getGenDownloadUrl(current, params), { credentials: 'include' });
      if (!res.ok) {
        let detail = 'Сервис генерации недоступен, попробуйте позже';
        try { const j = await res.json(); if (j && j.detail) detail = j.detail; } catch { /* ответ не JSON */ }
        setStatus({ kind: 'err', text: detail });
        return;
      }
      const blob = await res.blob();
      // Имя файла из Content-Disposition, иначе — по типу документа.
      const cd = res.headers.get('Content-Disposition') || '';
      const m = /filename\*?=(?:UTF-8''|")?([^";]+)/i.exec(cd);
      const name = m ? decodeURIComponent(m[1].replace(/"/g, '')) : `${current}.docx`;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);   // якорь в DOM — иначе часть браузеров не сохраняет
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);   // освобождаем URL после сохранения
      setStatus({ kind: 'ok', text: 'Готово — документ скачан ✓' });
    } catch {
      setStatus({ kind: 'err', text: 'Сервер недоступен, попробуйте позже' });
    }
  };

  return (
    <div className="animate-fade-in max-w-4xl mx-auto px-4 py-12 space-y-6">
      <div className="space-y-2">
        <h2 className="text-2xl font-bold text-gray-900">Генерация документов</h2>
        <p className="text-gray-500">Создание документов по утверждённым шаблонам</p>
      </div>

      {loadError && (
        <div className="flex items-center gap-2 bg-red-50 text-red-700 rounded-xl px-4 py-3 text-sm">
          <AlertCircle className="w-4 h-4 shrink-0" />
          {loadError}
        </div>
      )}

      {/* ── 1. Тип документа ── */}
      <div className="bg-white border border-gray-200 rounded-2xl p-6">
        <h3 className="font-semibold text-gray-900">1. Выберите тип документа</h3>
        <p className="text-sm text-gray-500 mt-1 mb-4">ИИ создаст готовый .docx по утверждённому шаблону</p>

        <div className="relative max-w-xl" ref={ddRef}>
          <button
            type="button"
            onClick={(e) => { e.stopPropagation(); setMenuOpen((v) => !v); }}
            className={`w-full flex items-center justify-between gap-3 bg-white border rounded-xl px-4 py-3 text-left transition cursor-pointer ${
              menuOpen ? 'border-blue-500 ring-2 ring-blue-500/20' : 'border-gray-200 hover:border-gray-300'
            }`}
          >
            <span className="min-w-0">
              {currentNode ? (
                <>
                  <span className="block font-semibold text-gray-900 truncate">
                    <span className="text-blue-600 font-bold mr-2">{currentNode.code || ''}</span>
                    {currentNode.title}
                  </span>
                  <span className="block text-xs text-gray-400 font-mono mt-0.5">{currentNode.doc_type}</span>
                </>
              ) : (
                <span className="text-gray-400">Загрузка…</span>
              )}
            </span>
            <ChevronDown className={`w-5 h-5 text-gray-400 shrink-0 transition-transform ${menuOpen ? 'rotate-180' : ''}`} />
          </button>

          {menuOpen && (
            <div className="absolute z-30 left-0 right-0 mt-1.5 bg-white border border-gray-200 rounded-xl shadow-xl max-h-80 overflow-auto p-1.5">
              {nodes.map((n) => (
                <div
                  key={n.doc_type}
                  onClick={() => { selectType(n.doc_type); setMenuOpen(false); }}
                  className={`px-3 py-2 rounded-lg cursor-pointer ${n.doc_type === current ? 'bg-blue-50' : 'hover:bg-gray-50'}`}
                >
                  <p className="text-sm font-semibold text-gray-900">
                    <span className="text-blue-600 font-bold mr-2">{n.code || ''}</span>
                    {n.title}
                  </p>
                  <p className="text-xs text-gray-400 font-mono">{n.doc_type}</p>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* ── 2. Поля документа ── */}
      {current && (
        <div className="bg-white border border-gray-200 rounded-2xl p-6">
          <h3 className="font-semibold text-gray-900">2. Заполните значения</h3>
          <p className="text-sm text-gray-500 mt-1 mb-5">
            {schemaLoading ? 'Загрузка формы…' : (schema?.title || '')}
          </p>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-5 gap-y-4">
            {fields.map((f) => (
              <div key={f.key}>
                <label className="flex items-center justify-between gap-2 text-sm font-medium text-gray-700 mb-1.5">
                  <span className="min-w-0">
                    {f.label}
                    {f.required && <span className="text-red-600"> *</span>}
                  </span>
                  {f.source && (
                    <span className="shrink-0 text-[11px] font-normal text-gray-500 bg-gray-50 border border-gray-200 rounded-md px-1.5 py-0.5">
                      {f.source === 'own' ? 'своё' : f.source}
                    </span>
                  )}
                </label>
                <input
                  ref={f.key === ORG_KEY ? orgRef : undefined}
                  value={values[f.key] || ''}
                  onChange={(e) => setValue(f.key, e.target.value)}
                  onBlur={f.key === ORG_KEY ? (e) => applyOrgSuggestions(e.target.value) : undefined}
                  list={`dl_${f.key}`}
                  placeholder={f.hint || ''}
                  className="w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm text-gray-900 placeholder:text-gray-400 placeholder:italic focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition"
                />
                <datalist id={`dl_${f.key}`}>
                  {(f.key === ORG_KEY ? orgs : (suggest[f.key] || [])).map((v, i) => (
                    <option key={`${f.key}_${i}`} value={String(v)} />
                  ))}
                </datalist>
              </div>
            ))}
          </div>

          <div className="mt-6 flex items-center gap-4 flex-wrap">
            <button
              onClick={generate}
              disabled={schemaLoading || !fields.length}
              className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 text-white font-medium rounded-xl flex items-center gap-2 transition-colors cursor-pointer disabled:cursor-not-allowed"
            >
              <Download className="w-5 h-5" />
              Сгенерировать документ
            </button>
            {status && (
              <span className={`text-sm ${status.kind === 'ok' ? 'text-emerald-600' : 'text-red-600'}`}>
                {status.text}
              </span>
            )}
          </div>
        </div>
      )}

    </div>
  );
}
