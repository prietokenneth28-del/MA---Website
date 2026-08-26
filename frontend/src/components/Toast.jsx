import React, { useEffect } from 'react';
import { X, CheckCircle, AlertCircle, Info } from 'lucide-react';

export const Toast = ({ message, type = 'info', onClose }) => {
  useEffect(() => {
    const timer = setTimeout(() => onClose(), 4000);
    return () => clearTimeout(timer);
  }, [onClose]);

  const bgColors = {
    success: 'bg-emerald-600',
    error: 'bg-rose-600',
    info: 'bg-blue-600',
  };

  const icons = {
    success: <CheckCircle className="w-5 h-5" />,
    error: <AlertCircle className="w-5 h-5" />,
    info: <Info className="w-5 h-5" />,
  };

  return (
    <div className={`fixed top-4 right-4 ${bgColors[type] || 'bg-slate-800'} text-white px-5 py-3.5 rounded-xl shadow-2xl z-50 flex items-center space-x-3 animate-fade-in-down border border-white/10 backdrop-blur-md`}>
      <span>{icons[type]}</span>
      <span className="text-sm font-medium">{message}</span>
      <button onClick={onClose} className="text-white/80 hover:text-white transition-colors p-1">
        <X className="w-4 h-4" />
      </button>
    </div>
  );
};

export default Toast;

