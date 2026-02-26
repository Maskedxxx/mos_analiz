import { useState, useCallback } from 'react';
import { LogOut } from 'lucide-react';
import ScreenLogin from './ScreenLogin';
import ScreenUpload from './ScreenUpload';
import ScreenProgress from './ScreenProgress';
import ScreenResults from './ScreenResults';
import './App.css';

function Header({ onLogout }) {
  return (
    <header className="bg-white border-b border-gray-200 py-4 px-6 flex items-center justify-between sticky top-0 z-50">
      <div className="flex items-center gap-3">
        <div className="w-9 h-9 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold text-sm">
          AI
        </div>
        <span className="font-bold text-xl text-gray-800 tracking-tight">
          ИИ-АУДИТ ДОКУМЕНТОВ
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
  const [screen, setScreen] = useState('login'); // login | upload | progress | results
  const [sessionId, setSessionId] = useState(null);
  const [filename, setFilename] = useState('');
  const [result, setResult] = useState(null);

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
  }, []);

  // Экран логина — без хедера
  if (screen === 'login') {
    return <ScreenLogin onLoginSuccess={handleLoginSuccess} />;
  }

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-900 pb-20">
      <Header onLogout={handleLogout} />
      <main>
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
      </main>
    </div>
  );
}
