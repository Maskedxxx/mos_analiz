import { useState, useEffect, useRef } from 'react';
import { Upload, FileText, X, Play, Loader2, AlertCircle, ChevronDown } from 'lucide-react';
import { fetchDocTypes, startAudit } from './api';

export default function ScreenUpload({ onAuditStarted }) {
  const [docTypes, setDocTypes] = useState([]);
  const [selectedType, setSelectedType] = useState('');
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const fileInputRef = useRef(null);

  // Загрузка типов при монтировании
  useEffect(() => {
    fetchDocTypes()
      .then((types) => setDocTypes(types))
      .catch((err) => setError(err.message));
  }, []);

  const handleFile = (f) => {
    if (f) {
      setFile(f);
      setError('');
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const f = e.dataTransfer.files[0];
    handleFile(f);
  };

  const handleStart = async () => {
    setLoading(true);
    setError('');
    try {
      const { session_id } = await startAudit(file, selectedType);
      onAuditStarted(session_id, file.name);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const canStart = file && selectedType && !loading;

  return (
    <div className="animate-fade-in max-w-2xl mx-auto px-4 py-12 space-y-8">
      <div className="text-center space-y-2">
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
            <option key={t.doc_type} value={t.doc_type}>
              {t.doc_title}
            </option>
          ))}
        </select>
        <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400 pointer-events-none" />
      </div>

      {/* Drag-drop зона */}
      {!file ? (
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
          <p className="text-gray-400 text-xs mt-3">.docx, .pptx, .xlsx</p>
          <input
            ref={fileInputRef}
            type="file"
            accept=".docx,.pptx,.xlsx"
            onChange={(e) => handleFile(e.target.files[0])}
            className="hidden"
          />
        </div>
      ) : (
        /* Выбранный файл */
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
  );
}
