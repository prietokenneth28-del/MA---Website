import React from 'react';
import { LayoutDashboard, UserPlus, Search, Image as ImageIcon, X, LogOut, ShieldCheck, FileEdit } from 'lucide-react';

export const Sidebar = ({ activeTab, setActiveTab, isMobileMenuOpen, setIsMobileMenuOpen, userProfile, onLogout, isDemoMode }) => {
  const menuItems = [
    { id: 'dashboard', label: 'Panel Principal', icon: LayoutDashboard },
    { id: 'digitacion', label: 'Digitación de Datos', icon: FileEdit },
    { id: 'registro', label: 'Alta de Personal', icon: UserPlus },
    { id: 'faltantes', label: 'Auditoría Faltantes', icon: Search },
    { id: 'imagenes', label: 'Mejorar Imágenes', icon: ImageIcon },
  ];

  return (
    <>
      {/* Mobile Overlay */}
      {isMobileMenuOpen && (
        <div 
          className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-40" 
          onClick={() => setIsMobileMenuOpen(false)} 
        />
      )}

      {/* Sidebar Container */}
      <aside className={`fixed md:static inset-y-0 left-0 w-72 bg-slate-900 text-white z-50 transform transition-transform duration-300 ease-in-out flex flex-col shadow-2xl border-r border-slate-800 flex-shrink-0 ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'} md:translate-x-0`}>
        <div className="p-6 flex items-center justify-between border-b border-slate-800">
          <div className="flex items-center space-x-3">
            <div className="w-9 h-9 bg-yellow-500 rounded-xl flex items-center justify-center text-slate-900 font-bold shadow-md shadow-yellow-500/20">
              <LayoutDashboard className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-white tracking-tight">
                Sis<span className="text-yellow-400">Operaciones</span>
              </h1>
              <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider block">Máquinas Amarillas</span>
            </div>
          </div>
          <button className="md:hidden text-slate-400 hover:text-white p-1" onClick={() => setIsMobileMenuOpen(false)}>
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Status Indicator */}
        {isDemoMode && (
          <div className="mx-4 mt-4 px-3 py-1.5 bg-yellow-500/10 border border-yellow-500/30 rounded-lg flex items-center gap-2">
            <div className="w-2 h-2 rounded-full bg-yellow-400 animate-pulse" />
            <span className="text-xs text-yellow-300 font-medium">Modo Simulación / Demo</span>
          </div>
        )}

        {/* User Profile Snippet */}
        {userProfile && (
          <div className="p-4 mx-4 mt-4 bg-slate-800/80 rounded-2xl flex items-center space-x-3 border border-slate-700/60 shadow-inner">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-yellow-500 to-amber-400 flex items-center justify-center font-bold text-slate-900 shrink-0 shadow-md">
              {userProfile.avatar || 'AD'}
            </div>
            <div className="overflow-hidden flex-1">
              <p className="text-sm font-semibold truncate text-white">{userProfile.name}</p>
              <p className="text-xs text-slate-400 truncate">{userProfile.email}</p>
              <button 
                onClick={onLogout} 
                className="text-xs text-rose-400 hover:text-rose-300 font-medium transition-colors mt-1 flex items-center gap-1"
              >
                <LogOut className="w-3 h-3" /> Cerrar sesión
              </button>
            </div>
          </div>
        )}

        {/* Navigation items */}
        <nav className="flex-1 p-4 space-y-1.5 overflow-y-auto mt-2">
          {menuItems.map(item => {
            const Icon = item.icon;
            const isActive = activeTab === item.id;
            return (
              <button
                key={item.id}
                onClick={() => { setActiveTab(item.id); setIsMobileMenuOpen(false); }}
                className={`w-full flex items-center space-x-3 px-4 py-3.5 rounded-xl transition-all duration-200 ${
                  isActive 
                    ? 'bg-yellow-500 text-slate-950 font-bold shadow-lg shadow-yellow-500/20' 
                    : 'text-slate-400 hover:bg-slate-800/80 hover:text-white font-medium'
                }`}
              >
                <Icon className={`w-5 h-5 ${isActive ? 'text-slate-950' : 'text-slate-400'}`} />
                <span className="text-sm">{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className="p-4 border-t border-slate-800/80 text-center">
          <p className="text-[11px] text-slate-500 flex items-center justify-center gap-1">
            <ShieldCheck className="w-3.5 h-3.5 text-emerald-500" /> Microsoft Graph Direct Sync
          </p>
        </div>
      </aside>
    </>
  );
};

export default Sidebar;