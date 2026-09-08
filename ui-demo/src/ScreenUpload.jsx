import { useState, useEffect, useRef } from 'react';
import {
  Upload, FileText, X, Play, Loader2, AlertCircle, ChevronDown,
  ListChecks, Search, MapPin, Plus, Pencil, Trash2, ChevronLeft, Sparkles, Check,
} from 'lucide-react';
import { fetchDocTypes, startAudit, startCrossAudit, fetchRules, draftRule, createRule, updateRule, deleteRule } from './api';

// Чип секции с кастомным тултипом (появляется через 0.5 с): описание + границы блока (start/end)
function SectionChip({ section, selected, onToggle }) {
  const [show, setShow] = useState(false);
  const timerRef = useRef(null);
  const enter = () => { timerRef.current = setTimeout(() => setShow(true), 500); };
  const leave = () => { clearTimeout(timerRef.current); setShow(false); };
  const hasHint = section.description || section.start || section.end;
  return (
    <div className="relative" onMouseEnter={enter} onMouseLeave={leave}>
      <button
        type="button"
        onClick={onToggle}
        className={`text-sm rounded-full px-3 py-1.5 border transition-colors cursor-pointer ${
          selected ? 'bg-blue-50 border-blue-300 text-blue-700' : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
        }`}
      >
        {section.name.replace(/_/g, ' ')}
      </button>
      {show && hasHint && (
        <div className="absolute z-50 bottom-full left-0 mb-2 w-64 rounded-xl bg-gray-900 text-white text-xs leading-relaxed px-3 py-2.5 shadow-xl">
          {section.description && <p className="mb-1.5">{section.description}</p>}
          {(section.start || section.end) && (
            <p className="text-gray-300">
              <span className="text-gray-500">Начало:</span> {section.start || '—'}<br />
              <span className="text-gray-500">Конец:</span> {section.end || '—'}
            </p>
          )}
          <span className="absolute top-full left-4 border-4 border-transparent border-t-gray-900" />
        </div>
      )}
    </div>
  );
}

const CROSS_TYPE = 'crosscheck_2_4_0_6_0_5';
const CROSS_SLOTS = [
  { role: 'kartochka', label: '2.4 Карточка проекта', hint: '.xlsx' },
  { role: 'protokol', label: '0.6 Протокол выполнения', hint: '.docx / .pdf' },
  { role: 'tirazh', label: '0.5 Приказ о тираже', hint: '.docx / .pdf' },
];

// Слот загрузки одного из трёх документов сквозной сверки.
function CrossSlot({ slot, file, onPick }) {
  const ref = useRef(null);
  return (
    <div className="border border-gray-200 rounded-xl px-4 py-3 flex items-center justify-between bg-white">
      <div className="min-w-0 mr-3">
        <p className="text-sm font-medium text-gray-900">{slot.label}</p>
        {file
          ? <p className="text-xs text-gray-500 truncate">{file.name}</p>
          : <p className="text-xs text-gray-400">{slot.hint}</p>}
      </div>
      <div className="flex items-center gap-2 shrink-0">
        {file && (
          <button onClick={() => onPick(null)} className="p-1 hover:bg-gray-100 rounded-lg cursor-pointer">
            <X className="w-4 h-4 text-gray-400" />
          </button>
        )}
        <button onClick={() => ref.current?.click()}
          className="text-sm px-3 py-1.5 rounded-lg border border-gray-200 hover:border-gray-300 text-gray-600 cursor-pointer">
          {file ? 'Заменить' : 'Выбрать'}
        </button>
        <input ref={ref} type="file" accept=".xlsx,.docx,.pdf"
          onChange={(e) => e.target.files[0] && onPick(e.target.files[0])} className="hidden" />
      </div>
    </div>
  );
}

export default function ScreenUpload({ onAuditStarted, initialType = '' }) {
  const [docTypes, setDocTypes] = useState([]);
  const [selectedType, setSelectedType] = useState(initialType); // после ошибки возвращаемся с тем же типом
  const [file, setFile] = useState(null);
  const [crossFiles, setCrossFiles] = useState({ kartochka: null, protokol: null, tirazh: null });
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const fileInputRef = useRef(null);

  // Панель правил
  const [editable, setEditable] = useState(false);
  const [sections, setSections] = useState([]);
  const [rules, setRules] = useState([]);
  const [rulesLoading, setRulesLoading] = useState(false);
  const [rulesError, setRulesError] = useState('');
  const [pendingDelete, setPendingDelete] = useState(null);

  // Редактор правила (drill-in)
  const [panelView, setPanelView] = useState('list'); // 'list' | 'editor'
  const [editIndex, setEditIndex] = useState(null);
  const [fTitle, setFTitle] = useState('');
  const [fRaw, setFRaw] = useState('');
  const [fSections, setFSections] = useState([]);
  const [fCheck, setFCheck] = useState('');
  const [fExclusions, setFExclusions] = useState('');
  const [hasDraft, setHasDraft] = useState(false);
  const [drafting, setDrafting] = useState(false);
  const [saving, setSaving] = useState(false);
  const [editorError, setEditorError] = useState('');

  // Загрузка типов при монтировании
  useEffect(() => {
    fetchDocTypes()
      .then((types) => setDocTypes(types))
      .catch((err) => setError(err.message));
  }, []);

  // Автозагрузка правил при выборе типа
  useEffect(() => {
    setPanelView('list');
    setPendingDelete(null);
    if (!selectedType) {
      setRules([]); setSections([]); setEditable(false); setRulesError('');
      return;
    }
    setRulesLoading(true); setRulesError('');
    fetchRules(selectedType)
      .then((d) => { setEditable(d.editable); setSections(d.sections || []); setRules(d.rules || []); })
      .catch((e) => setRulesError(e.message))
      .finally(() => setRulesLoading(false));
  }, [selectedType]);

  const reloadRules = async () => {
    const d = await fetchRules(selectedType);
    setEditable(d.editable); setSections(d.sections || []); setRules(d.rules || []);
  };

  const handleFile = (f) => {
    if (f) { setFile(f); setError(''); }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    handleFile(e.dataTransfer.files[0]);
  };

  const handleStart = async () => {
    setLoading(true);
    setError('');
    try {
      if (selectedType === CROSS_TYPE) {
        const files = [crossFiles.kartochka, crossFiles.protokol, crossFiles.tirazh].filter(Boolean);
        const { session_id } = await startCrossAudit(files);
        onAuditStarted(session_id, files.map((f) => f.name).join(', '), selectedType);
      } else {
        const { session_id } = await startAudit(file, selectedType);
        onAuditStarted(session_id, file.name, selectedType);
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  // ── Редактор ──
  const openNew = () => {
    setEditIndex(null);
    setFTitle(''); setFRaw(''); setFSections([]);
    setFCheck(''); setFExclusions(''); setHasDraft(false);
    setEditorError(''); setPanelView('editor');
  };

  const openEdit = (r) => {
    setEditIndex(r.index);
    setFTitle(r.title || '');
    setFRaw(r.check || '');
    setFSections(r.target_sections || []);
    setFCheck(r.check || ''); setFExclusions(r.exclusions || ''); setHasDraft(true);
    setEditorError(''); setPanelView('editor');
  };

  const doFormulate = async () => {
    setDrafting(true); setEditorError('');
    try {
      const d = await draftRule(selectedType, { title: fTitle, raw_check: fRaw, sections: fSections });
      setFCheck(d.check || ''); setFExclusions(d.exclusions || ''); setHasDraft(true);
    } catch (e) {
      setEditorError(e.message);
    } finally {
      setDrafting(false);
    }
  };

  const doSave = async () => {
    setSaving(true); setEditorError('');
    const payload = { title: fTitle.trim(), check: fCheck.trim(), target_sections: fSections, exclusions: fExclusions.trim() };
    try {
      if (editIndex == null) await createRule(selectedType, payload);
      else await updateRule(selectedType, editIndex, payload);
      await reloadRules();
      setPanelView('list');
    } catch (e) {
      setEditorError(e.message);
    } finally {
      setSaving(false);
    }
  };

  const doDelete = async (index) => {
    try {
      await deleteRule(selectedType, index);
      setPendingDelete(null);
      await reloadRules();
    } catch (e) {
      setRulesError(e.message);
    }
  };

  // Мультивыбор секций («где проверять»)
  const toggleSection = (name) => {
    setFSections((prev) => (prev.includes(name) ? prev.filter((s) => s !== name) : [...prev, name]));
  };
  const allSelected = sections.length > 0 && fSections.length === sections.length;
  const toggleAll = () => {
    setFSections(allSelected ? [] : sections.map((s) => s.name));
  };

  const isCross = selectedType === CROSS_TYPE;
  // Допустимые форматы выбранного типа (из /api/types) — для accept и подсказки (4.3a)
  const allowedExt = (docTypes.find((t) => t.doc_type === selectedType) || {}).allowed_extensions || [];
  const acceptList = allowedExt.length ? allowedExt.join(',') : '.docx,.pptx,.xlsx,.pdf';
  const extHint = allowedExt.length ? allowedExt.join(', ') : '.docx, .pptx, .xlsx, .pdf';
  const canStart = !loading && selectedType && (isCross
    ? (crossFiles.kartochka && crossFiles.protokol && crossFiles.tirazh)
    : file);

  return (
    <div className="animate-fade-in max-w-5xl mx-auto px-4 py-12">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-start">

        {/* ── ЛЕВО: загрузка документа ── */}
        <div className="space-y-8">
          <div className="text-center lg:text-left space-y-2">
            <h2 className="text-2xl font-bold text-gray-900">Загрузка документа</h2>
            <p className="text-gray-500">Выберите тип и загрузите файл для проверки</p>
          </div>

          {/* Dropdown — тип документа */}
          <div className="relative">
            <select
              value={selectedType}
              onChange={(e) => setSelectedType(e.target.value)}
              className="w-full appearance-none bg-white border border-gray-200 rounded-xl px-4 py-3.5 pr-10 text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition cursor-pointer"
            >
              <option value="" disabled>Выберите тип документа...</option>
              {docTypes.map((t) => (
                // Тип с повреждённой конфигурацией (broken) — виден, но недоступен для выбора
                <option key={t.doc_type} value={t.doc_type} disabled={!!t.broken}>
                  {t.broken ? `${t.doc_title} — недоступен: ${t.broken_reason}` : t.doc_title}
                </option>
              ))}
            </select>
            <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400 pointer-events-none" />
          </div>

          {/* Drag-drop зона / выбранный файл */}
          {isCross ? (
            <div className="space-y-3">
              {CROSS_SLOTS.map((slot) => (
                <CrossSlot key={slot.role} slot={slot} file={crossFiles[slot.role]}
                  onPick={(f) => setCrossFiles((p) => ({ ...p, [slot.role]: f }))} />
              ))}
            </div>
          ) : !file ? (
            <div
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`
                border-2 border-dashed rounded-2xl p-12 text-center cursor-pointer transition-colors
                ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-gray-400 bg-white'}
              `}
            >
              <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
              <p className="text-gray-600 font-medium">Перетащите файл сюда</p>
              <p className="text-gray-400 text-sm mt-1">или нажмите для выбора</p>
              <p className="text-gray-400 text-xs mt-3">{extHint}</p>
              <input
                ref={fileInputRef}
                type="file"
                accept={acceptList}
                onChange={(e) => handleFile(e.target.files[0])}
                className="hidden"
              />
            </div>
          ) : (
            <div className="bg-blue-50 border border-blue-200 rounded-xl px-5 py-4 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <FileText className="w-6 h-6 text-blue-600" />
                <div>
                  <p className="text-gray-900 font-medium">{file.name}</p>
                  <p className="text-gray-500 text-sm">{(file.size / 1024).toFixed(0)} КБ</p>
                </div>
              </div>
              <button
                onClick={() => setFile(null)}
                className="p-1.5 hover:bg-blue-100 rounded-lg transition-colors cursor-pointer"
              >
                <X className="w-5 h-5 text-gray-500" />
              </button>
            </div>
          )}

          {/* Ошибка */}
          {error && (
            <div className="flex items-center gap-2 bg-red-50 text-red-700 rounded-xl px-4 py-3 text-sm">
              <AlertCircle className="w-4 h-4 shrink-0" />
              {error}
            </div>
          )}

          {/* Кнопка запуска */}
          <button
            onClick={handleStart}
            disabled={!canStart}
            className="w-full py-3.5 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 text-white font-medium rounded-xl flex items-center justify-center gap-2 transition-colors cursor-pointer disabled:cursor-not-allowed"
          >
            {loading ? (
              <Loader2 className="w-5 h-5 animate-spin" />
            ) : (
              <>
                <Play className="w-5 h-5" />
                Выполнить аудит
              </>
            )}
          </button>
        </div>

        {/* ── ПРАВО: правила проверки ── */}
        <div className="space-y-6 lg:border-l lg:border-gray-100 lg:pl-8">
          {panelView === 'list' ? (
            <>
              <div className="flex items-start justify-between gap-3">
                <div className="text-center lg:text-left space-y-1">
                  <div className="flex items-center gap-2 justify-center lg:justify-start">
                    <ListChecks className="w-5 h-5 text-blue-600" />
                    <h2 className="text-2xl font-bold text-gray-900">Правила проверки</h2>
                    {selectedType && !rulesLoading && !rulesError && (
                      <span className="text-sm text-gray-400 font-medium">{rules.length}</span>
                    )}
                  </div>
                  <p className="text-gray-500 text-sm">Что и где проверяет система</p>
                </div>
                {selectedType && editable && (
                  <button
                    onClick={openNew}
                    title="Добавить правило"
                    className="shrink-0 w-9 h-9 rounded-full border border-gray-200 text-gray-500 hover:text-blue-600 hover:border-blue-200 flex items-center justify-center transition-colors cursor-pointer"
                  >
                    <Plus className="w-5 h-5" />
                  </button>
                )}
              </div>

              {!selectedType ? (
                <div className="text-sm text-gray-400 border border-dashed border-gray-200 rounded-2xl p-8 text-center">
                  Выберите тип документа слева — покажем правила проверки
                </div>
              ) : rulesLoading ? (
                <div className="flex items-center justify-center py-10 text-gray-300">
                  <Loader2 className="w-6 h-6 animate-spin" />
                </div>
              ) : rulesError ? (
                <div className="flex items-center gap-2 bg-red-50 text-red-700 rounded-xl px-4 py-3 text-sm">
                  <AlertCircle className="w-4 h-4 shrink-0" />
                  {rulesError}
                </div>
              ) : !editable ? (
                <div className="text-sm text-gray-400 border border-dashed border-gray-200 rounded-2xl p-8 text-center">
                  Правила этого типа заданы алгоритмом проверки. Добавление правил через интерфейс недоступно.
                </div>
              ) : (
                <div className="space-y-3 max-h-[60vh] overflow-y-auto pr-1">
                  {rules.map((r) => (
                    <div key={r.index} className="group border border-gray-200 rounded-xl p-4 space-y-2.5 hover:border-gray-300 transition-colors">
                      <div className="flex items-start justify-between gap-2">
                        <p className="font-medium text-gray-900 leading-snug">{r.title}</p>
                        <div className="flex items-center gap-1.5 shrink-0">
                          {r.custom && (
                            <span className="text-[11px] text-blue-500 bg-blue-50 rounded-full px-2 py-0.5">Своё</span>
                          )}
                          {r.custom && pendingDelete !== r.index && (
                            <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button onClick={() => openEdit(r)} className="p-1 text-gray-400 hover:text-blue-600 cursor-pointer">
                                <Pencil className="w-4 h-4" />
                              </button>
                              <button onClick={() => setPendingDelete(r.index)} className="p-1 text-gray-400 hover:text-red-600 cursor-pointer">
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          )}
                        </div>
                      </div>

                      {pendingDelete === r.index ? (
                        <div className="flex items-center gap-3 text-sm">
                          <span className="text-gray-500">Удалить правило?</span>
                          <button onClick={() => doDelete(r.index)} className="text-red-600 font-medium cursor-pointer">Удалить</button>
                          <button onClick={() => setPendingDelete(null)} className="text-gray-400 cursor-pointer">Отмена</button>
                        </div>
                      ) : (
                        <>
                          <div className="flex items-start gap-2 text-sm text-gray-600">
                            <Search className="w-4 h-4 shrink-0 mt-0.5 text-gray-300" />
                            <div><span className="text-gray-400">Проверяет · </span>{r.check}</div>
                          </div>
                          {r.sections && r.sections.length > 0 && (
                            <div className="flex items-start gap-2 text-sm text-gray-600">
                              <MapPin className="w-4 h-4 shrink-0 mt-0.5 text-gray-300" />
                              <div><span className="text-gray-400">Где · </span>{r.sections.map((s) => s.description || s.name).join('; ')}</div>
                            </div>
                          )}
                        </>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </>
          ) : (
            /* ── Редактор правила ── */
            <>
              <button onClick={() => setPanelView('list')} className="flex items-center gap-1 text-sm text-gray-500 hover:text-gray-700 cursor-pointer">
                <ChevronLeft className="w-4 h-4" /> Правила
              </button>
              <h2 className="text-xl font-bold text-gray-900">{editIndex == null ? 'Новое правило' : 'Изменить правило'}</h2>

              <div className="space-y-5">
                <div className="space-y-1.5">
                  <label className="text-xs uppercase tracking-wide text-gray-400 font-medium">Заголовок</label>
                  <input
                    value={fTitle}
                    onChange={(e) => setFTitle(e.target.value)}
                    placeholder="Напр.: Проверка печати"
                    className="w-full bg-white border border-gray-200 rounded-xl px-4 py-2.5 text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="text-xs uppercase tracking-wide text-gray-400 font-medium">Что проверять</label>
                  <textarea
                    value={fRaw}
                    onChange={(e) => setFRaw(e.target.value)}
                    rows={3}
                    placeholder="Своими словами: что должно быть в документе"
                    className="w-full bg-white border border-gray-200 rounded-xl px-4 py-2.5 text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition resize-none"
                  />
                </div>

                <div className="space-y-2">
                  <label className="text-xs uppercase tracking-wide text-gray-400 font-medium">Где проверять</label>
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={toggleAll}
                      className={`text-sm rounded-full px-3 py-1.5 border transition-colors cursor-pointer ${
                        allSelected ? 'bg-blue-600 border-blue-600 text-white' : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
                      }`}
                    >
                      Весь документ
                    </button>
                    {sections.map((s) => (
                      <SectionChip
                        key={s.name}
                        section={s}
                        selected={fSections.includes(s.name)}
                        onToggle={() => toggleSection(s.name)}
                      />
                    ))}
                  </div>
                </div>

                <button
                  onClick={doFormulate}
                  disabled={!fTitle || !fRaw || fSections.length === 0 || drafting}
                  className="flex items-center gap-2 text-sm text-blue-600 hover:text-blue-700 font-medium disabled:text-gray-300 cursor-pointer disabled:cursor-not-allowed"
                >
                  {drafting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />}
                  {hasDraft ? 'Переформулировать' : 'Сформулировать'}
                </button>

                {hasDraft && (
                  <div className="space-y-4 border-t border-gray-100 pt-4">
                    <div className="space-y-1.5">
                      <label className="text-xs uppercase tracking-wide text-gray-400 font-medium">Как поймёт модель</label>
                      <textarea
                        value={fCheck}
                        onChange={(e) => setFCheck(e.target.value)}
                        rows={4}
                        className="w-full bg-gray-50 border border-gray-200 rounded-xl px-4 py-2.5 text-sm text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition resize-none"
                      />
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-xs uppercase tracking-wide text-gray-400 font-medium">Исключения (необязательно)</label>
                      <input
                        value={fExclusions}
                        onChange={(e) => setFExclusions(e.target.value)}
                        placeholder="Что не считать нарушением"
                        className="w-full bg-gray-50 border border-gray-200 rounded-xl px-4 py-2.5 text-sm text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition"
                      />
                    </div>
                  </div>
                )}

                {editorError && (
                  <div className="flex items-center gap-2 bg-red-50 text-red-700 rounded-xl px-4 py-3 text-sm">
                    <AlertCircle className="w-4 h-4 shrink-0" />
                    {editorError}
                  </div>
                )}

                <button
                  onClick={doSave}
                  disabled={!fTitle || !fCheck || fSections.length === 0 || saving}
                  className="w-full py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-200 text-white font-medium rounded-xl flex items-center justify-center gap-2 transition-colors cursor-pointer disabled:cursor-not-allowed"
                >
                  {saving ? <Loader2 className="w-5 h-5 animate-spin" /> : <><Check className="w-5 h-5" /> Сохранить</>}
                </button>
              </div>
            </>
          )}
        </div>

      </div>
    </div>
  );
}
