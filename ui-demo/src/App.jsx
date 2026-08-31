import { useState, useCallback, useEffect } from 'react';
import { LogOut } from 'lucide-react';
import ScreenLogin from './ScreenLogin';
import { fetchDocTypes } from './api';
import Sidebar from './Sidebar';
import ScreenUpload from './ScreenUpload';
import ScreenProgress from './ScreenProgress';
import ScreenResults from './ScreenResults';
import ScreenGeneration from './ScreenGeneration';
import './App.css';

function Header({ onLogout }) {
  return (
    <header className="bg-white border-b border-gray-200 py-4 px-6 flex items-center justify-between sticky top-0 z-50">
      <div className="flex items-center gap-3">
        <div className="w-9 h-9 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold text-sm">
          AI
        </div>
        <span className="font-bold text-xl text-gray-800 tracking-tight">
          МосМониторинг <span className="font-normal text-gray-400">· Платформа</span>
        </span>
      </div>
      {onLogout && (
        <button
          onClick={onLogout}
          className="flex items-center gap-2 text-sm text-gray-500 hover:text-gray-700 transition-colors cursor-pointer"
        >
          <LogOut className="w-4 h-4" />
          Выйти
        </button>
      )}
    </header>
  );
}

export default function App() {
  const [screen, setScreen] = useState('checking'); // checking | login | upload | progress | results
  const [mode, setMode] = useState('audit');        // audit | generation — пункт сайдбара
  const [genVisited, setGenVisited] = useState(false); // генерация монтируется при первом заходе
  const [sessionId, setSessionId] = useState(null);
  const [filename, setFilename] = useState('');
  const [result, setResult] = useState(null);

  // Проба сессии при монтировании: жива ли кука
  useEffect(() => {
    fetchDocTypes()
      .then(() => setScreen('upload'))
      .catch(() => setScreen('login'));
  }, []);

  const handleLoginSuccess = useCallback(() => {
    setScreen('upload');
  }, []);

  const handleAuditStarted = useCallback((sid, fname) => {
    setSessionId(sid);
    setFilename(fname);
    setScreen('progress');
  }, []);

  const handleComplete = useCallback((data) => {
    setResult(data);
    setScreen('results');
  }, []);

  const handleError = useCallback(() => {
    setScreen('upload');
  }, []);

  const handleReset = useCallback(() => {
    setSessionId(null);
    setFilename('');
    setResult(null);
    setScreen('upload');
  }, []);

  const handleLogout = useCallback(() => {
    setSessionId(null);
    setFilename('');
    setResult(null);
    setScreen('login');
    setMode('audit');
  }, []);

  // Переключение режима: генерацию монтируем один раз и дальше только прячем
  const handleModeSelect = useCallback((next) => {
    setMode(next);
    if (next === 'generation') setGenVisited(true);
  }, []);

  // Пока идёт проба сессии — лёгкий лоадер, не мигаем логином
  if (screen === 'checking') {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50 text-gray-500">
        Загрузка…
      </div>
    );
  }

  // Экран логина — без хедера и сайдбара
  if (screen === 'login') {
    return <ScreenLogin onLoginSuccess={handleLoginSuccess} />;
  }

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-900">
      <Header onLogout={handleLogout} />
      <div className="flex flex-col md:flex-row">
        <Sidebar mode={mode} onSelect={handleModeSelect} />
        <main className="flex-1 min-w-0 pb-20">
          {/* Аудит остаётся смонтированным при уходе на генерацию: не рвём SSE и не теряем прогресс */}
          <div className={mode === 'audit' ? '' : 'hidden'}>
            {screen === 'upload' && (
              <ScreenUpload onAuditStarted={handleAuditStarted} />
            )}
            {screen === 'progress' && (
              <ScreenProgress
                sessionId={sessionId}
                filename={filename}
                onComplete={handleComplete}
                onError={handleError}
              />
            )}
            {screen === 'results' && result && (
              <ScreenResults
                result={result}
                sessionId={sessionId}
                filename={filename}
                onReset={handleReset}
              />
            )}
          </div>
          {genVisited && (
            <div className={mode === 'generation' ? '' : 'hidden'}>
              <ScreenGeneration />
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
