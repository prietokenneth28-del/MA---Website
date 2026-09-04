import React, { useState } from 'react';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { loginRequest, isConfiguredAzure } from './authConfig';

import Sidebar from './components/Sidebar';
import Toast from './components/Toast';
import PersonalRegistration from './components/PersonalRegistration';
import MissingAudit from './components/MissingAudit';
import ImageEnhancer from './components/ImageEnhancer';
import DataEntryWorkspace from './components/DataEntryWorkspace';

import { 
  LayoutDashboard, 
  UserPlus, 
  Search, 
  Image as ImageIcon, 
  Menu, 
  ShieldCheck, 
  Info,
  CheckCircle2,
  FileEdit
} from 'lucide-react';

export const App = () => {
  const { instance, accounts } = useMsal();
  const [activeTab, setActiveTab] = useState('dashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [toast, setToast] = useState(null);

  // Estado para autenticación en Modo Simulación
  const [simulatedAuth, setSimulatedAuth] = useState(false);
  const [simulatedUser, setSimulatedUser] = useState(null);
  const [accessToken, setAccessToken] = useState(null);

  const showToast = (message, type = 'info') => setToast({ message, type });

  const hasConfiguredAzure = isConfiguredAzure();

  // Flujo de Login con Microsoft Entra ID (MSAL)
  const handleMicrosoftLogin = async () => {
    if (!hasConfiguredAzure) {
      // Si no hay Client ID de Azure, iniciamos en Modo Simulación
      showToast('Iniciando en Modo Simulación (Azure Client ID no detectado)...', 'info');
      setTimeout(() => {
        setSimulatedAuth(true);
        setSimulatedUser({
          name: 'Administrador de Operaciones',
          email: 'admin@empresa.com',
          avatar: 'AD'
        });
        showToast('Sesión iniciada correctamente (Modo Simulación).', 'success');
      }, 800);
      return;
    }

    try {
      showToast('Conectando con Microsoft Entra ID...', 'info');
      const response = await instance.loginPopup(loginRequest);
      setAccessToken(response.accessToken);
      showToast('Sesión iniciada con Microsoft Graph.', 'success');
    } catch (error) {
      console.error("Error en autenticación MSAL:", error);
      showToast('No se pudo iniciar sesión con Microsoft.', 'error');
    }
  };

  const handleLogout = () => {
    if (hasConfiguredAzure && accounts.length > 0) {
      instance.logoutPopup();
    } else {
      setSimulatedAuth(false);
      setSimulatedUser(null);
    }
    setAccessToken(null);
    showToast('Sesión cerrada.', 'info');
  };

  // Determinar si el usuario está autenticado (MSAL o Simulado)
  const isAuthenticated = (accounts.length > 0) || simulatedAuth;
  const currentUser = accounts.length > 0 ? {
    name: accounts[0].name || 'Usuario OneDrive',
    email: accounts[0].username || '',
    avatar: (accounts[0].name || 'U').substring(0, 2).toUpperCase()
  } : simulatedUser;

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-100 p-4">
        <div className="max-w-md w-full bg-white rounded-3xl shadow-2xl overflow-hidden p-8 border border-slate-200/80">
          <div className="text-center mb-8">
            <div className="w-16 h-16 bg-yellow-500 rounded-2xl mx-auto mb-4 flex items-center justify-center shadow-lg shadow-yellow-500/30 transform rotate-3">
              <LayoutDashboard className="w-8 h-8 text-slate-950" />
            </div>
            <h1 className="text-3xl font-black text-slate-900 mb-1 tracking-tight">
              Sis<span className="text-yellow-500">Operaciones</span>
            </h1>
            <p className="text-slate-500 text-sm font-medium">Gestión en la nube de reportes y máquinas amarillas</p>
          </div>

          <div className="space-y-6">
            <div className="bg-blue-50 border border-blue-200/60 p-4 rounded-2xl flex items-start space-x-3">
              <Info className="w-5 h-5 text-blue-600 shrink-0 mt-0.5" />
              <p className="text-xs text-blue-900 leading-relaxed">
                Esta aplicación requiere acceso a su cuenta de Microsoft para sincronizar con <strong>OneDrive</strong> mediante Graph API.
              </p>
            </div>

            <button 
              onClick={handleMicrosoftLogin}
              className="w-full flex justify-center items-center space-x-3 bg-slate-900 hover:bg-slate-800 text-white p-4 rounded-2xl font-bold transition-all shadow-xl hover:shadow-2xl transform hover:-translate-y-0.5 cursor-pointer"
            >
              <svg className="w-5 h-5" viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg">
                <path d="M0 0h10v10H0zm11 0h10v10H11zM0 11h10v10H0zm11 0h10v10H11z" fill="currentColor"/>
              </svg>
              <span>Iniciar sesión con Microsoft</span>
            </button>

            {!hasConfiguredAzure && (
              <p className="text-center text-[11px] text-slate-400">
                (Al no tener Client ID en .env, iniciará automáticamente en Modo Simulación para pruebas localmente).
              </p>
            )}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-slate-50 font-sans selection:bg-yellow-200">
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}

      <Sidebar 
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        isMobileMenuOpen={isMobileMenuOpen}
        setIsMobileMenuOpen={setIsMobileMenuOpen}
        userProfile={currentUser}
        onLogout={handleLogout}
        isDemoMode={!hasConfiguredAzure}
      />

      <main className="flex-1 flex flex-col overflow-hidden bg-slate-50">
        {/* Top Header con Botón de Menú para abrir Sidebar */}
        <header className="bg-white border-b border-slate-200 p-3.5 px-4 flex items-center justify-between shadow-xs shrink-0">
          <div className="flex items-center space-x-3">
            <button 
              onClick={() => setIsMobileMenuOpen(true)} 
              className="text-slate-700 hover:text-slate-900 hover:bg-slate-100 p-2 rounded-xl transition-colors cursor-pointer"
              title="Abrir menú"
            >
              <Menu className="w-6 h-6" />
            </button>
            <div className="w-7 h-7 bg-yellow-500 rounded-lg text-slate-950 flex items-center justify-center font-bold">
              <LayoutDashboard className="w-4 h-4" />
            </div>
            <h1 className="font-bold text-slate-900 text-base tracking-tight">Sis<span className="text-yellow-500">Operaciones</span></h1>
          </div>
          <div className="text-xs font-semibold text-slate-500 uppercase tracking-wider hidden sm:block">
            {activeTab === 'dashboard' && 'Panel Principal'}
            {activeTab === 'digitacion' && 'Digitación de Datos'}
            {activeTab === 'registro' && 'Alta de Personal'}
            {activeTab === 'faltantes' && 'Auditoría de Faltantes'}
            {activeTab === 'imagenes' && 'Mejora de Imágenes'}
          </div>
        </header>

        {/* Content Body */}
        <div className={`flex-1 overflow-y-auto ${activeTab === 'digitacion' ? 'p-0' : 'p-4 md:p-8'} flex flex-col`}>
          <div className={`${activeTab === 'digitacion' ? 'w-full h-full flex flex-col' : 'max-w-5xl mx-auto space-y-6 w-full'}`}>

            {/* TAB: Dashboard */}
            {activeTab === 'dashboard' && (
              <div className="space-y-6 animate-fade-in-up">
                <div className="bg-gradient-to-r from-slate-900 to-slate-800 text-white p-8 rounded-3xl shadow-xl relative overflow-hidden">
                  <div className="relative z-10">
                    <h2 className="text-3xl font-black tracking-tight">Bienvenido al Panel de Operaciones</h2>
                    <p className="text-slate-300 text-sm mt-2 max-w-xl">
                      Gestiona la creación de estructuras, auditoría, mejora de imágenes y digitación de reportes diarios.
                    </p>
                  </div>
                  <div className="absolute right-[-20px] bottom-[-20px] opacity-10 pointer-events-none">
                    <LayoutDashboard className="w-64 h-64 text-white" />
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
                  <div 
                    onClick={() => setActiveTab('digitacion')}
                    className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-all cursor-pointer group hover:-translate-y-1"
                  >
                    <div className="w-12 h-12 bg-emerald-50 text-emerald-600 rounded-xl flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                      <FileEdit className="w-6 h-6" />
                    </div>
                    <h3 className="font-bold text-slate-900 text-lg">Digitación</h3>
                    <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                      Visualiza reportes y vales simultáneamente para transcribir datos al sistema.
                    </p>
                  </div>

                  <div 
                    onClick={() => setActiveTab('registro')}
                    className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-all cursor-pointer group hover:-translate-y-1"
                  >
                    <div className="w-12 h-12 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                      <UserPlus className="w-6 h-6" />
                    </div>
                    <h3 className="font-bold text-slate-900 text-lg">Alta Personal</h3>
                    <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                      Crea la estructura de carpetas anual en OneDrive.
                    </p>
                  </div>

                  <div 
                    onClick={() => setActiveTab('faltantes')}
                    className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-all cursor-pointer group hover:-translate-y-1"
                  >
                    <div className="w-12 h-12 bg-rose-50 text-rose-600 rounded-xl flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                      <Search className="w-6 h-6" />
                    </div>
                    <h3 className="font-bold text-slate-900 text-lg">Faltantes</h3>
                    <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                      Auditoría de días sin reportes subidos.
                    </p>
                  </div>

                  <div 
                    onClick={() => setActiveTab('imagenes')}
                    className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm hover:shadow-md transition-all cursor-pointer group hover:-translate-y-1"
                  >
                    <div className="w-12 h-12 bg-purple-50 text-purple-600 rounded-xl flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                      <ImageIcon className="w-6 h-6" />
                    </div>
                    <h3 className="font-bold text-slate-900 text-lg">Mejora Imag.</h3>
                    <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                      Filtros interactivos de claridad y contraste para fotos oscuras.
                    </p>
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'registro' && <PersonalRegistration showToast={showToast} accessToken={accessToken} />}
            {activeTab === 'faltantes' && <MissingAudit showToast={showToast} accessToken={accessToken} />}
            {activeTab === 'imagenes' && <ImageEnhancer showToast={showToast} />}
            
            {/* NUEVO TAB: Digitación */}
            {activeTab === 'digitacion' && (
              <DataEntryWorkspace showToast={showToast} accessToken={accessToken} />
            )}

          </div>
        </div>
      </main>
    </div>
  );
};

export default App;
