import React, { useState, useRef } from 'react';
import { Image as ImageIcon, Upload, Download, Sparkles, Sliders, CheckCircle2 } from 'lucide-react';
import { enhanceImageCanvas } from '../services/imageService';

export const ImageEnhancer = ({ showToast }) => {
  const [selectedFile, setSelectedFile] = useState(null);
  const [originalUrl, setOriginalUrl] = useState(null);
  const [enhancedUrl, setEnhancedUrl] = useState(null);
  const [brightness, setBrightness] = useState(1.15);
  const [contrast, setContrast] = useState(1.3);
  const [sharpen, setSharpen] = useState(true);
  const [processing, setProcessing] = useState(false);
  const imgRef = useRef(null);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      showToast('Por favor selecciona un archivo de imagen válido (.jpg, .png)', 'error');
      return;
    }

    setSelectedFile(file);
    const url = URL.createObjectURL(file);
    setOriginalUrl(url);
    setEnhancedUrl(null);
  };

  const handleProcessImage = () => {
    if (!imgRef.current) return;
    setProcessing(true);

    setTimeout(() => {
      try {
        const resultDataUrl = enhanceImageCanvas(imgRef.current, {
          brightness: parseFloat(brightness),
          contrast: parseFloat(contrast),
          sharpen: sharpen
        });

        setEnhancedUrl(resultDataUrl);
        showToast('¡Imagen procesada con éxito con filtros de claridad y enfoque!', 'success');
      } catch (error) {
        console.error("Error al procesar imagen en canvas:", error);
        showToast('Ocurrió un error al procesar la imagen.', 'error');
      } finally {
        setProcessing(false);
      }
    }, 100);
  };

  const handleDownload = () => {
    if (!enhancedUrl) return;
    const a = document.createElement('a');
    a.href = enhancedUrl;
    a.download = `mejora_${selectedFile ? selectedFile.name : 'reporte.jpg'}`;
    a.click();
    showToast('Imagen mejorada guardada.', 'success');
  };

  return (
    <div className="space-y-6 animate-fade-in-up">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 md:p-8">
        <div className="mb-6 border-b border-slate-100 pb-4">
          <h2 className="text-2xl font-bold text-slate-900 flex items-center gap-2.5">
            <ImageIcon className="w-7 h-7 text-purple-600" /> Mejora de Imágenes
          </h2>
          <p className="text-slate-500 text-sm mt-1">Aplica filtros de claridad, ecualización y enfoque a las fotos ilegibles de horómetros y preoperacionales.</p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Panel de Controles */}
          <div className="space-y-6 bg-slate-50 p-6 rounded-2xl border border-slate-200/80">
            <h3 className="font-bold text-slate-900 flex items-center gap-2 text-base">
              <Sliders className="w-5 h-5 text-purple-600" /> Ajustes de Filtro
            </h3>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">Seleccionar Imagen</label>
              <label className="border-2 border-dashed border-slate-300 hover:border-purple-500 rounded-xl p-4 bg-white flex flex-col items-center justify-center cursor-pointer transition-colors text-center">
                <Upload className="w-8 h-8 text-slate-400 mb-2" />
                <span className="text-xs font-bold text-slate-700">Subir foto de reporte / horómetro</span>
                <span className="text-[11px] text-slate-400 mt-1">JPG, JPEG o PNG</span>
                <input type="file" accept="image/*" onChange={handleFileChange} className="hidden" />
              </label>
            </div>

            {originalUrl && (
              <div className="space-y-4 pt-2">
                <div>
                  <div className="flex justify-between text-xs font-semibold text-slate-700 mb-1">
                    <span>Brillo / Claridad</span>
                    <span>{Math.round(brightness * 100)}%</span>
                  </div>
                  <input 
                    type="range" 
                    min="0.8" 
                    max="1.8" 
                    step="0.05" 
                    value={brightness}
                    onChange={(e) => setBrightness(e.target.value)}
                    className="w-full accent-purple-600"
                  />
                </div>

                <div>
                  <div className="flex justify-between text-xs font-semibold text-slate-700 mb-1">
                    <span>Contraste</span>
                    <span>{Math.round(contrast * 100)}%</span>
                  </div>
                  <input 
                    type="range" 
                    min="0.8" 
                    max="2.0" 
                    step="0.05" 
                    value={contrast}
                    onChange={(e) => setContrast(e.target.value)}
                    className="w-full accent-purple-600"
                  />
                </div>

                <label className="flex items-center space-x-3 cursor-pointer select-none">
                  <input 
                    type="checkbox" 
                    checked={sharpen}
                    onChange={(e) => setSharpen(e.target.checked)}
                    className="w-4 h-4 text-purple-600 rounded border-slate-300 focus:ring-purple-500" 
                  />
                  <span className="text-xs font-semibold text-slate-700">Aplicar Máscara de Enfoque (Sharpen)</span>
                </label>

                <button 
                  onClick={handleProcessImage}
                  disabled={processing}
                  className="w-full bg-purple-600 hover:bg-purple-700 text-white font-bold py-3.5 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 text-sm"
                >
                  <Sparkles className="w-4 h-4 text-yellow-300" />
                  <span>{processing ? 'Procesando en Canvas...' : 'Aplicar Mejora de Imagen'}</span>
                </button>
              </div>
            )}
          </div>

          {/* Panel de Previsualización */}
          <div className="lg:col-span-2 space-y-6">
            {!originalUrl ? (
              <div className="h-80 border-2 border-dashed border-slate-200 rounded-2xl flex flex-col items-center justify-center p-8 text-center bg-slate-50/50">
                <ImageIcon className="w-12 h-12 text-slate-300 mb-3" />
                <p className="font-semibold text-slate-600 text-sm">No se ha cargado ninguna imagen</p>
                <p className="text-xs text-slate-400 mt-1 max-w-sm">Selecciona una imagen desde el panel izquierdo para aplicar los filtros de claridad y enfoque.</p>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* Original */}
                  <div className="space-y-2">
                    <span className="text-xs font-bold text-slate-500 uppercase tracking-wider block">Foto Original</span>
                    <div className="border border-slate-200 rounded-xl overflow-hidden bg-black/5 flex items-center justify-center max-h-80">
                      <img 
                        ref={imgRef}
                        src={originalUrl} 
                        alt="Original" 
                        className="max-h-72 object-contain"
                      />
                    </div>
                  </div>

                  {/* Mejorada */}
                  <div className="space-y-2">
                    <span className="text-xs font-bold text-purple-600 uppercase tracking-wider block flex items-center gap-1">
                      <Sparkles className="w-3.5 h-3.5" /> Foto Mejorada
                    </span>
                    <div className="border border-purple-200 rounded-xl overflow-hidden bg-black/5 flex items-center justify-center max-h-80 min-h-[18rem]">
                      {enhancedUrl ? (
                        <img 
                          src={enhancedUrl} 
                          alt="Mejorada" 
                          className="max-h-72 object-contain"
                        />
                      ) : (
                        <div className="text-center p-4 text-slate-400 text-xs">
                          Haz clic en "Aplicar Mejora de Imagen" para visualizar el resultado.
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {enhancedUrl && (
                  <div className="pt-2 flex justify-end">
                    <button 
                      onClick={handleDownload}
                      className="bg-slate-900 hover:bg-slate-800 text-white font-bold py-3.5 px-6 rounded-xl shadow-lg transition-all flex items-center gap-2 text-sm"
                    >
                      <Download className="w-4 h-4 text-yellow-400" /> Guardar Imagen Procesada
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default ImageEnhancer;

