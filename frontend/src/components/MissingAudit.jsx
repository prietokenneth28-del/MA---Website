import React, { useState } from 'react';
import { Search, FileSpreadsheet, Send, AlertTriangle, CheckCircle, Loader2 } from 'lucide-react';
import { scanMissingReportsOneDrive } from '../services/graphService';
import { downloadAuditExcel } from '../services/excelService';

const PROJECT_FOLDERS = {
  conductores: [
    '1. AGUAS BOGOTA',
    '2. QUORA',
    '3. GRUAS AGUAS BOGOTA',
    '4. TERMOZIPA',
    '5. NUEVOS'
  ],
  operadores: [
    'AES SANTA MARIA',
    'AGUAS BOGOTA',
    'GRUPO CEMEX TUNJUELITO',
    'PLANTA MONDOÑEDO',
    'TERMOZIPA',
    'VARIOS CONTRATOS',
    'Z. NUEVOS'
  ]
};

export const MissingAudit = ({ showToast, accessToken }) => {
  const [type, setType] = useState('conductores');
  const [startDate, setStartDate] = useState('2026-08-01');
  const [endDate, setEndDate] = useState('2026-08-08');
  const [scanning, setScanning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [auditResults, setAuditResults] = useState(null);

  const handleScan = async (e) => {
    e.preventDefault();

    if (new Date(startDate) > new Date(endDate)) {
      showToast('La fecha inicial no puede ser mayor que la fecha final.', 'error');
      return;
    }

    setScanning(true);
    setProgress(0);
    setAuditResults(null);

    try {
      showToast(`Escaneando carpetas en OneDrive para ${type}...`, 'info');

      const projectsToScan = PROJECT_FOLDERS[type];
      const results = await scanMissingReportsOneDrive(
        accessToken, 
        type, 
        projectsToScan, 
        startDate, 
        endDate, 
        (pct) => setProgress(pct)
      );

      setAuditResults(results);
      
      if (results.length === 0) {
        showToast('¡Excelente! Todos los reportes están al día en las fechas seleccionadas.', 'success');
      } else {
        showToast(`Auditoría completa. Se encontraron ${results.length} registros con faltantes.`, 'info');
      }
    } catch (error) {
      console.error("Error en auditoría:", error);
      showToast('Ocurrió un problema durante el escaneo en la nube.', 'error');
    } finally {
      setScanning(false);
    }
  };

  const handleExportExcel = () => {
    if (!auditResults || auditResults.length === 0) return;
    downloadAuditExcel(auditResults, `reporte_faltantes_${type}_${startDate}_al_${endDate}.xlsx`);
    showToast('Archivo Excel descargado correctamente.', 'success');
  };

  const openWhatsAppMessage = (msg) => {
    const encoded = encodeURIComponent(msg);
    window.open(`https://web.whatsapp.com/send?text=${encoded}`, '_blank');
  };

  return (
    <div className="space-y-6 animate-fade-in-up">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 md:p-8">
        <div className="mb-6 border-b border-slate-100 pb-4">
          <h2 className="text-2xl font-bold text-slate-900 flex items-center gap-2.5">
            <Search className="w-7 h-7 text-rose-500" /> Auditoría de Faltantes
          </h2>
          <p className="text-slate-500 text-sm mt-1">Escanea OneDrive buscando días sin reportes y genera un Excel con mensajes de alerta.</p>
        </div>

        <form onSubmit={handleScan} className="space-y-5">
          <div>
            <label className="block text-sm font-semibold text-slate-700 mb-1.5">Categoría a Auditar</label>
            <div className="grid grid-cols-2 gap-3 max-w-md">
              <button
                type="button"
                onClick={() => setType('conductores')}
                className={`py-3 px-4 rounded-xl text-sm font-bold border transition-all ${
                  type === 'conductores' 
                    ? 'bg-rose-600 text-white border-rose-600 shadow-md' 
                    : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Conductores
              </button>
              <button
                type="button"
                onClick={() => setType('operadores')}
                className={`py-3 px-4 rounded-xl text-sm font-bold border transition-all ${
                  type === 'operadores' 
                    ? 'bg-rose-600 text-white border-rose-600 shadow-md' 
                    : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Operadores
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1.5">Fecha Inicial</label>
              <input 
                type="date" 
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="w-full border border-slate-300 rounded-xl p-3 focus:ring-2 focus:ring-rose-500 outline-none transition-shadow font-medium"
              />
            </div>
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1.5">Fecha Final</label>
              <input 
                type="date" 
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="w-full border border-slate-300 rounded-xl p-3 focus:ring-2 focus:ring-rose-500 outline-none transition-shadow font-medium"
              />
            </div>
          </div>

          <div className="pt-2">
            <button 
              type="submit" 
              disabled={scanning}
              className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-4 rounded-xl transition-all shadow-lg flex justify-center items-center gap-2 hover:shadow-xl disabled:opacity-50"
            >
              {scanning ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin text-white" />
                  <span>Escaneando Nube ({progress}%)...</span>
                </>
              ) : (
                <>
                  <Search className="w-5 h-5" />
                  <span>Iniciar Escaneo en OneDrive</span>
                </>
              )}
            </button>
          </div>
        </form>

        {/* Barra de Progreso */}
        {scanning && (
          <div className="mt-6 space-y-2">
            <div className="flex justify-between text-xs font-semibold text-slate-600">
              <span>Procesando proyectos y días...</span>
              <span>{progress}%</span>
            </div>
            <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden">
              <div 
                className="bg-rose-500 h-full transition-all duration-300 rounded-full" 
                style={{ width: `${progress}%` }} 
              />
            </div>
          </div>
        )}

        {/* Resultados de la Auditoría */}
        {auditResults && (
          <div className="mt-8 space-y-6 animate-fade-in-up">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div>
                <h3 className="text-xl font-bold text-slate-900 flex items-center gap-2">
                  Resultados del Escaneo ({auditResults.length})
                </h3>
                <p className="text-xs text-slate-500">Rango: {startDate} al {endDate}</p>
              </div>

              {auditResults.length > 0 && (
                <button 
                  onClick={handleExportExcel}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2.5 rounded-xl shadow-md transition-all flex items-center gap-2 text-sm"
                >
                  <FileSpreadsheet className="w-4 h-4" /> Exportar Reporte Excel
                </button>
              )}
            </div>

            {auditResults.length === 0 ? (
              <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 p-6 rounded-2xl flex items-center gap-3 text-sm">
                <CheckCircle className="w-6 h-6 text-emerald-600 shrink-0" />
                <span><strong>¡Todo al día!</strong> No se registraron días faltantes para las fechas seleccionadas.</span>
              </div>
            ) : (
              <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-xs">
                <table className="w-full text-left text-sm border-collapse">
                  <thead className="bg-slate-900 text-white font-semibold text-xs uppercase">
                    <tr>
                      <th className="p-3.5">Proyecto</th>
                      <th className="p-3.5">Nombre Conductor / Operario</th>
                      <th className="p-3.5">Días Faltantes</th>
                      <th className="p-3.5 text-center">Cant.</th>
                      <th className="p-3.5 text-center">Acciones</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {auditResults.map((item, idx) => {
                      const isHighAlert = item['Cantidad Faltantes'] >= 3;
                      return (
                        <tr key={idx} className={isHighAlert ? 'bg-rose-50/60 hover:bg-rose-100/50' : 'hover:bg-slate-50'}>
                          <td className="p-3.5 font-semibold text-slate-800">{item.Proyecto}</td>
                          <td className="p-3.5 text-slate-900 font-bold">{item['Nombre Conductor']}</td>
                          <td className="p-3.5 text-slate-600 font-mono text-xs">{item['Días Faltantes']}</td>
                          <td className="p-3.5 text-center">
                            <span className={`px-2.5 py-1 rounded-full font-bold text-xs ${
                              isHighAlert ? 'bg-rose-600 text-white' : 'bg-slate-200 text-slate-800'
                            }`}>
                              {item['Cantidad Faltantes']}
                            </span>
                          </td>
                          <td className="p-3.5 text-center">
                            <button 
                              onClick={() => openWhatsAppMessage(item['Mensaje Generado'])}
                              className="p-2 text-emerald-600 hover:bg-emerald-50 rounded-lg transition-colors"
                              title="Abrir mensaje de WhatsApp"
                            >
                              <Send className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default MissingAudit;

