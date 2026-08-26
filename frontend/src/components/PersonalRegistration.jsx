import React, { useState } from 'react';
import { UserPlus, Copy, Send, CheckCircle2, Loader2, FolderPlus, FileSpreadsheet } from 'lucide-react';
import { createFolderStructureOneDrive } from '../services/graphService';

const PROJECT_CONFIG = {
  CONDUCTOR: {
    label: "Conductor",
    projects: ['1. AGUAS BOGOTA', '2. QUORA', '3. GRUAS AGUAS BOGOTA', '4. TERMOZIPA', '5. NUEVOS']
  },
  OPERARIO: {
    label: "Operario",
    projects: ['AES SANTA MARIA', 'AGUAS BOGOTA', 'GRUPO CEMEX TUNJUELITO', 'PLANTA MONDOÑEDO', 'TERMOZIPA', 'VARIOS CONTRATOS', 'Z. NUEVOS']
  }
};

export const PersonalRegistration = ({ showToast, accessToken }) => {
  const [role, setRole] = useState('CONDUCTOR');
  const [project, setProject] = useState(PROJECT_CONFIG.CONDUCTOR.projects[0]);
  const [fullName, setFullName] = useState('');
  const [idNumber, setIdNumber] = useState('');
  const [phone, setPhone] = useState('');
  const [loading, setLoading] = useState(false);
  const [credentials, setCredentials] = useState(null);

  const handleRoleChange = (newRole) => {
    setRole(newRole);
    setProject(PROJECT_CONFIG[newRole].projects[0]);
    setCredentials(null);
  };

  // Algoritmo idéntico al script de Python para generar usuario y contraseña
  const generateCredentials = (name) => {
    const parts = name.trim().toUpperCase().split(/\s+/);
    if (parts.length === 0 || !parts[0]) return { user: '', pass: '', firstName: '' };

    const firstName = parts[0];
    let firstLastName = '';

    if (parts.length >= 4) {
      firstLastName = parts[2];
    } else if (parts.length >= 2) {
      firstLastName = parts[1];
    }

    const firstLetter = firstName[0] || '';
    const user = `${firstLetter}${firstLastName}`;
    const pass = `${firstLetter}${firstLastName}.123*`;
    const formattedFirstName = firstName.charAt(0).toUpperCase() + firstName.slice(1).toLowerCase();

    return { user, pass, firstName: formattedFirstName };
  };

  const copyToClipboard = (text, label) => {
    navigator.clipboard.writeText(text);
    showToast(`¡${label} copiado al portapapeles!`, 'success');
  };

  const openWhatsApp = (firstName, phoneNumber, user, pass) => {
    const cleanPhone = phoneNumber.replace(/\+/g, '').trim();
    const phoneWithCountry = cleanPhone.startsWith('57') ? cleanPhone : `57${cleanPhone}`;
    
    const message = `Tus credenciales de acceso para diligenciar los preoperacionales digitales son:\n\n` +
      `👤 *Usuario:* ${user}\n` +
      `🔑 *Contraseña:* ${pass}\n\n` +
      `Por favor guarda esta información.`;

    const encodedMessage = encodeURIComponent(message);
    const whatsappUrl = `https://web.whatsapp.com/send?phone=${phoneWithCountry}&text=${encodedMessage}`;
    
    window.open(whatsappUrl, '_blank');
    showToast(`WhatsApp Web abierto para ${firstName}`, 'success');
  };

  const handleSubmit = async (e) => {
    e.preventDefault();

    if (!fullName || !idNumber || !phone) {
      showToast('Por favor diligencia todos los campos obligatorios.', 'error');
      return;
    }

    const creds = generateCredentials(fullName);
    setCredentials(creds);
    setLoading(true);

    try {
      showToast('Creando estructura de carpetas en OneDrive...', 'info');

      // 1. Crear carpetas de meses y días en OneDrive vía Graph API
      await createFolderStructureOneDrive(accessToken, role, project, fullName.trim().toUpperCase());

      // 2. Copiar automáticamente datos al portapapeles
      const fullCopyData = `Nombre: ${fullName}\nID: ${idNumber}\nCelular: ${phone}\nUsuario: ${creds.user}\nContraseña: ${creds.pass}`;
      navigator.clipboard.writeText(fullCopyData);

      showToast(`¡Alta completada para ${fullName}! Carpetas creadas y credenciales copiadas.`, 'success');

      // 3. Abrir WhatsApp Web con mensaje listo
      openWhatsApp(creds.firstName, phone, creds.user, creds.pass);

    } catch (error) {
      console.error(error);
      showToast('Ocurrió un error al procesar el alta en OneDrive.', 'error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-6 animate-fade-in-up">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 md:p-8">
        <div className="mb-6 border-b border-slate-100 pb-4 flex items-center justify-between">
          <div>
            <h2 className="text-2xl font-bold text-slate-900 flex items-center gap-2.5">
              <UserPlus className="w-7 h-7 text-yellow-500" /> Alta de Personal
            </h2>
            <p className="text-slate-500 text-sm mt-1">Crea la estructura de carpetas anual en OneDrive y genera credenciales.</p>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1.5">Rol de Personal</label>
              <div className="grid grid-cols-2 gap-3">
                <button
                  type="button"
                  onClick={() => handleRoleChange('CONDUCTOR')}
                  className={`py-3 px-4 rounded-xl text-sm font-bold border transition-all ${
                    role === 'CONDUCTOR' 
                      ? 'bg-slate-900 text-white border-slate-900 shadow-md' 
                      : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  Conductor
                </button>
                <button
                  type="button"
                  onClick={() => handleRoleChange('OPERARIO')}
                  className={`py-3 px-4 rounded-xl text-sm font-bold border transition-all ${
                    role === 'OPERARIO' 
                      ? 'bg-slate-900 text-white border-slate-900 shadow-md' 
                      : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  Operario
                </button>
              </div>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1.5">Proyecto Asignado</label>
              <select 
                value={project}
                onChange={(e) => setProject(e.target.value)}
                className="w-full border border-slate-300 rounded-xl p-3 bg-slate-50 text-slate-800 font-medium focus:ring-2 focus:ring-yellow-500 outline-none transition-shadow"
              >
                {PROJECT_CONFIG[role].projects.map(proj => (
                  <option key={proj} value={proj}>{proj}</option>
                ))}
              </select>
            </div>
          </div>

          <div>
            <label className="block text-sm font-semibold text-slate-700 mb-1.5">Nombre Completo</label>
            <input 
              type="text" 
              value={fullName}
              onChange={(e) => setFullName(e.target.value)}
              placeholder="Ej. JUAN CARLOS PEREZ GOMEZ" 
              className="w-full border border-slate-300 rounded-xl p-3 focus:ring-2 focus:ring-yellow-500 outline-none transition-shadow uppercase"
            />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1.5">Identificación / Cédula</label>
              <input 
                type="text" 
                value={idNumber}
                onChange={(e) => setIdNumber(e.target.value)}
                placeholder="Número de cédula" 
                className="w-full border border-slate-300 rounded-xl p-3 focus:ring-2 focus:ring-yellow-500 outline-none transition-shadow"
              />
            </div>
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1.5">Celular (WhatsApp)</label>
              <input 
                type="text" 
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                placeholder="3000000000" 
                className="w-full border border-slate-300 rounded-xl p-3 focus:ring-2 focus:ring-yellow-500 outline-none transition-shadow"
              />
            </div>
          </div>

          <div className="pt-2">
            <button 
              type="submit" 
              disabled={loading}
              className="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold py-4 rounded-xl transition-all shadow-lg flex justify-center items-center gap-2 hover:shadow-xl disabled:opacity-50"
            >
              {loading ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin text-yellow-400" />
                  <span>Creando Estructura en OneDrive...</span>
                </>
              ) : (
                <>
                  <FolderPlus className="w-5 h-5 text-yellow-400" />
                  <span>Registrar y Crear Carpetas en OneDrive</span>
                </>
              )}
            </button>
          </div>
        </form>

        {credentials && (
          <div className="mt-8 bg-slate-900 text-white p-6 rounded-2xl border border-slate-800 space-y-4 animate-fade-in-up">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="font-bold text-yellow-400 flex items-center gap-2">
                <CheckCircle2 className="w-5 h-5" /> Credenciales Generadas
              </h3>
              <span className="text-xs bg-yellow-500/20 text-yellow-300 px-2.5 py-1 rounded-full border border-yellow-500/30">Auto Generado</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
              <div className="bg-slate-800/80 p-3.5 rounded-xl border border-slate-700/60 flex items-center justify-between">
                <div>
                  <span className="text-xs text-slate-400 block">Usuario</span>
                  <span className="font-mono font-bold text-white text-base">{credentials.user}</span>
                </div>
                <button 
                  onClick={() => copyToClipboard(credentials.user, 'Usuario')} 
                  className="text-slate-400 hover:text-white p-2"
                >
                  <Copy className="w-4 h-4" />
                </button>
              </div>

              <div className="bg-slate-800/80 p-3.5 rounded-xl border border-slate-700/60 flex items-center justify-between">
                <div>
                  <span className="text-xs text-slate-400 block">Contraseña</span>
                  <span className="font-mono font-bold text-white text-base">{credentials.pass}</span>
                </div>
                <button 
                  onClick={() => copyToClipboard(credentials.pass, 'Contraseña')} 
                  className="text-slate-400 hover:text-white p-2"
                >
                  <Copy className="w-4 h-4" />
                </button>
              </div>
            </div>

            <div className="pt-2 flex flex-col sm:flex-row gap-3">
              <button 
                type="button"
                onClick={() => openWhatsApp(credentials.firstName, phone, credentials.user, credentials.pass)}
                className="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-3 px-4 rounded-xl transition-all shadow-md flex items-center justify-center gap-2 text-sm"
              >
                <Send className="w-4 h-4" /> Abrir WhatsApp con Mensaje
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default PersonalRegistration;

