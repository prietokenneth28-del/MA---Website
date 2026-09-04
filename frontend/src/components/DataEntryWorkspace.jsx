import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { getGraphClient, downloadFileFromOneDrive, uploadFileToOneDrive } from '../services/graphService';
import { appendPersonToExcelBuffer, createFreshExcelBuffer } from '../services/excelService';
import { 
  FolderSearch, 
  Image as ImageIcon, 
  X, 
  ZoomIn, 
  ZoomOut, 
  Save, 
  Loader2,
  RefreshCw,
  LayoutGrid,
  ChevronRight,
  Maximize,
  FileEdit,
  RotateCw,
  RotateCcw,
  Download,
  ChevronUp,
  ChevronDown
} from 'lucide-react';

const MESES = {
  1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
  5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
  9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
};

const BASE_PATHS = {
  CONDUCTORES: "1. MAQUINAS AMARILLAS/CONDUCTORES",
  OPERADORES: "1. MAQUINAS AMARILLAS/OPERADORES"
};

const EXCEL_PATHS = {
  CONDUCTOR: "1. MAQUINAS AMARILLAS/CONDUCTORES/REPORTES/LISTADO DE CONDUCTORES.xlsx",
  OPERARIO: "1. MAQUINAS AMARILLAS/OPERADORES/1. REPORTES/LISTADO DE OPERARIOS.xlsx"
};

const TIPOS_PAGO = ['CONTADO', 'CREDITO', 'BONO', 'TRANSFERENCIA', 'OTRO'];

// --- Subcomponente: Visor de Imagen con Zoom, Pan, Rotación y Arrastre a Windows ---
const ImagePanZoom = ({ src, alt, onRemove }) => {
  const [scale, setScale] = useState(1);
  const [rotation, setRotation] = useState(0);
  const [isDragging, setIsDragging] = useState(false);
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  const containerRef = useRef(null);

  // Estado para exportar/arrastrar la imagen hacia Windows Desktop o carpetas de Windows
  const [isFetchingBlob, setIsFetchingBlob] = useState(false);
  const blobRef = useRef(null);

  // Pre-descarga el Blob de la imagen para tenerlo listo al arrastrar a Windows
  useEffect(() => {
    if (src && !blobRef.current && !isFetchingBlob) {
      setIsFetchingBlob(true);
      fetch(src)
        .then(res => res.blob())
        .then(blob => { blobRef.current = blob; })
        .catch(err => console.error('[ImagePanZoom] No se pudo pre-cargar la imagen:', err))
        .finally(() => setIsFetchingBlob(false));
    }
  }, [src]);

  useEffect(() => {
    blobRef.current = null;
  }, [src]);

  // Manejador de arrastre directo hacia el Escritorio o Explorador de Windows
  const handleImageDragStart = (e) => {
    const fileName = (alt && /\.(jpg|jpeg|png)$/i.test(alt)) ? alt : `${alt || 'reporte'}.jpg`;
    const mimeType = (blobRef.current && blobRef.current.type) || 'image/jpeg';
    
    if (src) {
      try {
        // Formato nativo de Chromium para guardar archivos al arrastrar fuera del navegador hacia Windows
        e.dataTransfer.setData('DownloadURL', `${mimeType}:${fileName}:${src}`);
      } catch (err) {
        console.warn('[ImagePanZoom] DownloadURL no soportado:', err);
      }
      e.dataTransfer.setData('text/plain', src);
      e.dataTransfer.setData('text/uri-list', src);
    }

    if (blobRef.current) {
      try {
        const file = new File([blobRef.current], fileName, { type: mimeType });
        e.dataTransfer.items.add(file);
      } catch (err) {
        console.warn('[ImagePanZoom] No se pudo adjuntar el objeto File:', err);
      }
    }
    e.dataTransfer.effectAllowed = 'copy';
  };

  // Descarga directa a la PC mediante clic
  const handleDownloadToPC = async () => {
    if (!src) return;
    try {
      const response = await fetch(src);
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      const fileName = (alt && /\.(jpg|jpeg|png)$/i.test(alt)) ? alt : `${alt || 'reporte'}.jpg`;
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('[ImagePanZoom] Error al descargar imagen:', err);
      window.open(src, '_blank');
    }
  };

  const handleWheel = (e) => {
    if (e.ctrlKey) {
      e.preventDefault();
      const scaleAdjust = e.deltaY * -0.01;
      setScale(Math.min(Math.max(0.5, scale + scaleAdjust), 4));
    }
  };

  const handleMouseDown = (e) => {
    if (e.button !== 0) return;
    setIsDragging(true);
    setDragStart({ x: e.clientX - position.x, y: e.clientY - position.y });
  };

  const handleMouseMove = (e) => {
    if (isDragging) {
      setPosition({
        x: e.clientX - dragStart.x,
        y: e.clientY - dragStart.y
      });
    }
  };

  const handleMouseUp = () => setIsDragging(false);

  useEffect(() => {
    const container = containerRef.current;
    if (container) {
      container.addEventListener('wheel', handleWheel, { passive: false });
      return () => container.removeEventListener('wheel', handleWheel);
    }
  }, [scale]);

  return (
    <div 
      className="relative w-full h-full bg-slate-900 overflow-hidden border border-slate-700 rounded-lg group"
      ref={containerRef}
      onMouseDown={handleMouseDown}
      onMouseMove={handleMouseMove}
      onMouseUp={handleMouseUp}
      onMouseLeave={handleMouseUp}
      style={{ cursor: isDragging ? 'grabbing' : 'grab' }}
    >
      {/* Barra de herramientas flotante */}
      <div className="absolute top-2 right-2 z-10 flex space-x-1 opacity-0 group-hover:opacity-100 transition-opacity bg-black/60 p-1 rounded-lg backdrop-blur-xs">
        <button onClick={() => setScale(s => Math.min(s + 0.2, 4))} title="Zoom In" className="p-1 text-white hover:bg-white/20 rounded"><ZoomIn size={16}/></button>
        <button onClick={() => setScale(s => Math.max(s - 0.2, 0.5))} title="Zoom Out" className="p-1 text-white hover:bg-white/20 rounded"><ZoomOut size={16}/></button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={() => setRotation(r => (r - 90 + 360) % 360)} title="Rotar Izquierda" className="p-1 text-white hover:bg-white/20 rounded"><RotateCcw size={16}/></button>
        <button onClick={() => setRotation(r => (r + 90) % 360)} title="Rotar Derecha" className="p-1 text-white hover:bg-white/20 rounded"><RotateCw size={16}/></button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={handleDownloadToPC} title="Descargar archivo a Windows (PC)" className="p-1 text-emerald-400 hover:bg-emerald-500/20 rounded flex items-center gap-1">
          <Download size={16}/>
        </button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={() => { setScale(1); setPosition({x:0, y:0}); setRotation(0); }} title="Restablecer vista" className="p-1 text-white hover:bg-white/20 rounded"><Maximize size={16}/></button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={onRemove} title="Cerrar imagen" className="p-1 text-rose-400 hover:bg-rose-500/20 rounded"><X size={16}/></button>
      </div>
      
      <div 
        className="w-full h-full flex items-center justify-center origin-center"
        style={{
          transform: `translate(${position.x}px, ${position.y}px) scale(${scale}) rotate(${rotation}deg)`,
          transition: isDragging ? 'none' : 'transform 0.1s ease-out'
        }}
      >
        <img 
          src={src} 
          alt={alt} 
          draggable={true}
          onDragStart={handleImageDragStart}
          className="max-w-full max-h-full object-contain cursor-grab active:cursor-grabbing"
          referrerPolicy="no-referrer"
          onError={(e) => {
            console.error(`[ImagePanZoom] Falló la carga de "${alt}". URL:`, src);
          }}
        />
      </div>
      
      {/* Indicador de ayuda */}
      <div className="absolute bottom-2 left-2 text-[10px] text-white/60 bg-black/60 backdrop-blur-xs px-2.5 py-1 rounded pointer-events-none flex items-center gap-2">
        <span>Ctrl + Scroll (Zoom) | Arrastrar (Paneo) | Arrastra la foto hacia Windows / Escritorio para guardar</span>
        {isFetchingBlob && <Loader2 className="w-3 h-3 animate-spin text-yellow-400" />}
      </div>
    </div>
  );
};


const DataEntryWorkspace = ({ showToast, accessToken }) => {
  // --- Estados de Búsqueda ---
  const [tipo, setTipo] = useState('CONDUCTORES');
  const [proyecto, setProyecto] = useState('');
  const [persona, setPersona] = useState('');
  const [fecha, setFecha] = useState(new Date().toISOString().split('T')[0]);
  
  // Listas cargadas dinámicamente
  const [proyectosList, setProyectosList] = useState([]);
  const [personasList, setPersonasList] = useState([]); 

  // --- Estados de Imágenes y Workspace ---
  const [availableImages, setAvailableImages] = useState([]);
  const [isLoadingImages, setIsLoadingImages] = useState(false);
  const [selectedImages, setSelectedImages] = useState([]);

  // Loading separado para las listas de Proyectos/Trabajadores
  const [isLoadingLists, setIsLoadingLists] = useState(false);

  // Estado para expandir/colapsar el panel inferior de Digitación de Datos
  const [isFormExpanded, setIsFormExpanded] = useState(false);
  
  // --- Estados del Formulario ---
  const [formData, setFormData] = useState({
    consecutivoRegistro: '',
    horometroInicial: '',
    horometroFinal: '',
    tipoPago: '',
    vales: [{ consecutivo: '', pesoEntrada: '', pesoSalida: '' }]
  });

  // Estado de guardado del reporte
  const [isSaving, setIsSaving] = useState(false);

  // 1. Cargar proyectos reales cuando cambia el tipo de personal o el token de acceso
  useEffect(() => {
    const fetchProjects = async () => {
      if (!accessToken) {
        setProyectosList(
          tipo === 'CONDUCTORES'
            ? ['1. AGUAS BOGOTA', '2. QUORA', '3. GRUAS AGUAS BOGOTA', '4. TERMOZIPA', '6. NUEVOS']
            : ['AES SANTA MARIA', 'AGUAS BOGOTA', 'GRUPO CEMEX TUNJUELITO', 'PLANTA MONDOÑEDO', 'TERMOZIPA', 'VARIOS CONTRATOS', 'Z. NUEVOS']
        );
        return;
      }

      setIsLoadingLists(true);
      try {
        const client = getGraphClient(accessToken);
        const basePath = BASE_PATHS[tipo];
        const response = await client.api(`/me/drive/root:/${basePath}:/children`).get();
        
        const projects = (response.value || [])
          .filter(item => item.folder)
          .map(item => item.name);
        
        setProyectosList(projects);
      } catch (error) {
        console.error("Error cargando proyectos desde OneDrive:", error);
        showToast("Error al cargar los proyectos desde OneDrive", "error");
      } finally {
        setIsLoadingLists(false);
      }
    };

    fetchProjects();
  }, [tipo, accessToken]);

  // 2. Cargar conductores/operadores reales cuando cambia el proyecto seleccionado
  useEffect(() => {
    const fetchWorkers = async () => {
      if (!proyecto) {
        setPersonasList([]);
        return;
      }

      if (!accessToken) {
        const cleanedProj = proyecto.replace(/^\d+\.\s*/, '');
        if (cleanedProj.includes('AGUAS BOGOTA')) {
          setPersonasList(['CARLOS RODRIGUEZ', 'BRAYAN CASTRO']);
        } else if (cleanedProj.includes('QUORA')) {
          setPersonasList(['JOSE GOMEZ']);
        } else {
          setPersonasList(['OPERARIO PRUEBA']);
        }
        return;
      }

      setIsLoadingLists(true);
      try {
        const client = getGraphClient(accessToken);
        const basePath = BASE_PATHS[tipo];
        const projectPath = `${basePath}/${proyecto}`;
        const response = await client.api(`/me/drive/root:/${projectPath}:/children`).get();
        
        const workers = (response.value || [])
          .filter(item => item.folder)
          .map(item => item.name);
        
        setPersonasList(workers);
      } catch (error) {
        console.error("Error cargando trabajadores desde OneDrive:", error);
        showToast("Error al cargar la lista de trabajadores", "error");
        setPersonasList([]);
      } finally {
        setIsLoadingLists(false);
      }
    };

    fetchWorkers();
    setPersona('');
  }, [proyecto, accessToken, tipo]);

  // --- Funciones MS Graph para buscar imágenes ---
  const fetchImagesFromOneDrive = async () => {
    if (!proyecto || !persona || !fecha) {
      showToast("Selecciona Proyecto, Trabajador y Fecha", "warning");
      return;
    }

    setIsLoadingImages(true);
    setAvailableImages([]);
    setSelectedImages([]);

    const [year, month, day] = fecha.split('-');
    const monthName = MESES[parseInt(month, 10)];
    const monthFolderName = `${month}_${monthName}`;
    const basePath = BASE_PATHS[tipo];
    const folderPath = `${basePath}/${proyecto}/${persona}/${monthFolderName}/${day}`;

    if (!accessToken) {
      setTimeout(() => {
        setAvailableImages([
          { id: 'img1', name: 'Vale_Botadero.jpg', url: 'https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?auto=format&fit=crop&w=800&q=80' },
          { id: 'img2', name: 'Registro_Operacion.jpg', url: 'https://images.unsplash.com/photo-1554415707-6e8cfc93fe23?auto=format&fit=crop&w=800&q=80' },
          { id: 'img3', name: 'Vale_Adicional.jpg', url: 'https://images.unsplash.com/photo-1506784983877-45594efa4cbe?auto=format&fit=crop&w=800&q=80' }
        ]);
        setIsLoadingImages(false);
        showToast("Imágenes de prueba cargadas (Modo Demo)", "success");
      }, 1000);
      return;
    }

    try {
      const client = getGraphClient(accessToken);
      const response = await client.api(`/me/drive/root:/${folderPath}:/children`).get();

      const rawItems = (response.value || [])
        .filter(item => item.file && /\.(jpg|jpeg|png)$/i.test(item.name));

      const images = rawItems.map(item => ({
        id: item.id,
        name: item.name,
        url: item['@microsoft.graph.downloadUrl'] || null
      }));

      setAvailableImages(images);
      if (images.length === 0) showToast("No se encontraron imágenes en esta fecha", "info");
    } catch (error) {
      console.error("Error obteniendo imágenes:", error);
      if (error.statusCode === 404) {
         showToast("La carpeta para esta fecha no existe en OneDrive", "warning");
      } else {
         showToast("Error al conectar con OneDrive", "error");
      }
    } finally {
      setIsLoadingImages(false);
    }
  };

  // UI Selection
  const toggleImageSelection = (img) => {
    if (selectedImages.find(s => s.id === img.id)) {
      setSelectedImages(selectedImages.filter(s => s.id !== img.id));
    } else {
      if (selectedImages.length >= 3) {
        showToast("Máximo 3 imágenes simultáneas en el visualizador", "warning");
        return;
      }
      setSelectedImages([...selectedImages, img]);
    }
  };

  // Manejo de Vales
  const handleValeChange = (index, field, value) => {
    const newVales = [...formData.vales];
    newVales[index][field] = value;
    setFormData({...formData, vales: newVales});
  };

  const addVale = () => {
    setFormData({
      ...formData,
      vales: [...formData.vales, { consecutivo: '', pesoEntrada: '', pesoSalida: '' }]
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!proyecto || !persona || !fecha) {
      showToast("Selecciona Proyecto, Trabajador y Fecha para guardar", "warning");
      return;
    }

    setIsSaving(true);
    try {
      showToast("Guardando reporte...", "info");
      await new Promise(r => setTimeout(r, 1200));
      showToast("Reporte guardado con éxito", "success");
    } catch (err) {
      console.error(err);
      showToast("Error al guardar reporte", "error");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="flex flex-col h-full bg-slate-50 min-h-[500px]">
      
      {/* BARRA SUPERIOR DE BÚSQUEDA */}
      <div className="bg-white border-b border-slate-200 p-4 shadow-sm z-10 flex-shrink-0">
        <div className="flex flex-wrap items-end gap-4">
          <div className="flex-1 min-w-[150px]">
            <label className="block text-xs font-semibold text-slate-500 mb-1">Tipo de Personal</label>
            <select 
              value={tipo} 
              onChange={(e) => { setTipo(e.target.value); setProyecto(''); setPersona(''); }}
              className="w-full text-sm border-slate-300 rounded-lg focus:ring-yellow-500 focus:border-yellow-500"
            >
              <option value="CONDUCTORES">Conductores</option>
              <option value="OPERADORES">Operadores</option>
            </select>
          </div>
          <div className="flex-1 min-w-[150px]">
            <label className="block text-xs font-semibold text-slate-500 mb-1 flex items-center gap-1">
              Proyecto
              {isLoadingLists && <Loader2 className="w-3 h-3 animate-spin text-slate-400" />}
            </label>
            <select 
              value={proyecto} 
              onChange={(e) => setProyecto(e.target.value)} 
              disabled={isLoadingLists}
              className="w-full text-sm border-slate-300 rounded-lg focus:ring-yellow-500 focus:border-yellow-500 disabled:bg-slate-100"
            >
              <option value="">{isLoadingLists ? 'Cargando...' : 'Seleccione...'}</option>
              {proyectosList.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div className="flex-1 min-w-[150px]">
            <label className="block text-xs font-semibold text-slate-500 mb-1 flex items-center gap-1">
              Trabajador
              {isLoadingLists && proyecto && <Loader2 className="w-3 h-3 animate-spin text-slate-400" />}
            </label>
            <select 
              value={persona} 
              onChange={(e) => setPersona(e.target.value)} 
              disabled={!proyecto || isLoadingLists} 
              className="w-full text-sm border-slate-300 rounded-lg focus:ring-yellow-500 focus:border-yellow-500 disabled:bg-slate-100"
            >
              <option value="">{isLoadingLists ? 'Cargando...' : 'Seleccione...'}</option>
              {personasList.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div className="flex-1 min-w-[150px]">
            <label className="block text-xs font-semibold text-slate-500 mb-1">Fecha</label>
            <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className="w-full text-sm border-slate-300 rounded-lg focus:ring-yellow-500 focus:border-yellow-500" />
          </div>
          <button 
            onClick={fetchImagesFromOneDrive}
            disabled={isLoadingImages}
            className="bg-slate-900 text-white px-5 py-2.5 rounded-lg text-sm font-bold hover:bg-slate-800 transition-colors flex items-center space-x-2 disabled:opacity-50"
          >
            {isLoadingImages ? <Loader2 className="w-4 h-4 animate-spin text-yellow-500" /> : <FolderSearch className="w-4 h-4" />}
            <span>Buscar Archivos</span>
          </button>
        </div>
      </div>

      {/* ÁREA PRINCIPAL DE VISUALIZACIÓN (Abarca 100% de la pantalla) */}
      <div className="flex-1 flex overflow-hidden min-h-0">
        
        {/* PANEL IZQUIERDO: Lista de imágenes encontradas */}
        <div className="w-64 bg-white border-r border-slate-200 flex flex-col overflow-hidden hidden md:flex flex-shrink-0">
          <div className="p-3 bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-600 uppercase flex items-center justify-between">
            <span>Archivos en Carpeta</span>
            <span className="bg-slate-200 px-2 py-0.5 rounded-full text-slate-700">{availableImages.length}</span>
          </div>
          <div className="flex-1 overflow-y-auto p-2 space-y-2">
            {availableImages.map((img) => {
              const isSelected = selectedImages.some(s => s.id === img.id);
              return (
                <div 
                  key={img.id}
                  onClick={() => toggleImageSelection(img)}
                  className={`relative group cursor-pointer border-2 rounded-lg overflow-hidden transition-all ${isSelected ? 'border-yellow-500 shadow-md' : 'border-transparent hover:border-slate-300'}`}
                >
                  {img.url ? (
                    <img 
                      src={img.url} 
                      alt={img.name} 
                      className="w-full h-24 object-cover bg-slate-100"
                      referrerPolicy="no-referrer"
                    />
                  ) : (
                    <div className="w-full h-24 bg-slate-100 flex items-center justify-center">
                      <ImageIcon className="w-6 h-6 text-slate-300" />
                    </div>
                  )}
                  <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 flex items-center justify-center transition-opacity">
                    <span className="text-white text-xs font-bold">{isSelected ? 'Quitar' : 'Visualizar'}</span>
                  </div>
                  {isSelected && (
                    <div className="absolute top-1 right-1 w-5 h-5 bg-yellow-500 rounded-full border-2 border-white flex items-center justify-center">
                      <ChevronRight size={12} className="text-white"/>
                    </div>
                  )}
                  <div className="absolute bottom-0 left-0 right-0 bg-gradient-to-t from-black/80 to-transparent p-1 text-[10px] text-white truncate">
                    {img.name}
                  </div>
                </div>
              )
            })}
            {availableImages.length === 0 && !isLoadingImages && (
              <div className="text-center p-4 text-slate-400 text-sm flex flex-col items-center justify-center h-full">
                <ImageIcon className="w-8 h-8 mb-2 opacity-50 text-slate-300" />
                <p>Realiza una búsqueda para ver imágenes</p>
              </div>
            )}
          </div>
        </div>

        {/* WORKSPACE CENTRAL: Visualizador principal que abarca todo el ancho libre */}
        <div className="flex-1 bg-slate-800 p-2 flex flex-col min-w-0 overflow-hidden">
          {selectedImages.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center text-slate-400">
              <LayoutGrid className="w-16 h-16 mb-4 opacity-30" />
              <p className="text-lg font-semibold text-slate-300">Visualizador de Vales y Reportes</p>
              <p className="text-sm">Selecciona las imágenes del panel izquierdo para abrirlas aquí.</p>
              <p className="text-xs text-slate-500 mt-2 bg-slate-900/60 px-3 py-1.5 rounded-lg border border-slate-700">
                Tip: Puedes arrastrar cualquier fotografía directamente desde aquí hacia el Escritorio o carpetas de Windows.
              </p>
            </div>
          ) : (
            <div className={`flex-1 grid gap-2 overflow-hidden ${
              selectedImages.length === 1 ? 'grid-cols-1' : 
              selectedImages.length === 2 ? 'grid-cols-2' : 'grid-cols-2 lg:grid-cols-3'
            }`}>
              {selectedImages.map((img) => (
                <ImagePanZoom 
                  key={img.id} 
                  src={img.url} 
                  alt={img.name} 
                  onRemove={() => toggleImageSelection(img)} 
                />
              ))}
            </div>
          )}
        </div>

      </div>

      {/* SECCIÓN INFERIOR: FORMULARIO DE DIGITACIÓN DE DATOS (DESPLEGABLE EN LA PARTE DE ABAJO) */}
      <div className="border-t border-slate-200 bg-white flex-shrink-0 shadow-lg z-20">
        
        {/* Cabecera para Expandir / Reducir */}
        <div 
          onClick={() => setIsFormExpanded(!isFormExpanded)}
          className="bg-slate-900 text-white px-5 py-3 flex items-center justify-between cursor-pointer hover:bg-slate-800 transition-colors select-none"
        >
          <div className="flex items-center space-x-3">
            <FileEdit className="w-5 h-5 text-yellow-400" />
            <span className="font-bold text-sm tracking-wide">Formulario de Digitación de Datos</span>
            <span className="text-[10px] bg-slate-800 text-yellow-300 px-2 py-0.5 rounded-full border border-slate-700">Opcional</span>
          </div>
          <div className="flex items-center space-x-2 text-xs text-slate-300 font-medium">
            <span>{isFormExpanded ? 'Reducir Formulario' : 'Abrir Formulario de Digitación'}</span>
            {isFormExpanded ? <ChevronDown size={18} className="text-yellow-400" /> : <ChevronUp size={18} className="text-yellow-400" />}
          </div>
        </div>

        {/* Cuerpo del Formulario en la parte de abajo */}
        {isFormExpanded && (
          <div className="p-5 max-h-80 overflow-y-auto bg-slate-50 border-t border-slate-200 animate-fade-in-up">
            <form id="dataEntryForm" onSubmit={handleSubmit} className="space-y-4">
              
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Consecutivo RDO</label>
                  <input 
                    type="text" 
                    value={formData.consecutivoRegistro} 
                    onChange={e => setFormData({...formData, consecutivoRegistro: e.target.value})}
                    placeholder="Ej. 177570"
                    className="w-full text-sm border-slate-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 bg-slate-50" 
                  />
                </div>
                
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Tipo de Pago</label>
                  <select
                    value={formData.tipoPago}
                    onChange={e => setFormData({...formData, tipoPago: e.target.value})}
                    className="w-full text-sm border-slate-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 bg-slate-50"
                  >
                    <option value="">Seleccione...</option>
                    {TIPOS_PAGO.map(tp => <option key={tp} value={tp}>{tp}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Horómetro Inicial</label>
                  <input 
                    type="number" 
                    value={formData.horometroInicial} 
                    onChange={e => setFormData({...formData, horometroInicial: e.target.value})}
                    className="w-full text-sm border-slate-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 bg-slate-50" 
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Horómetro Final</label>
                  <input 
                    type="number" 
                    value={formData.horometroFinal} 
                    onChange={e => setFormData({...formData, horometroFinal: e.target.value})}
                    className="w-full text-sm border-slate-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 bg-slate-50" 
                  />
                </div>
              </div>

              {/* Sección Vales Botadero */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs space-y-3">
                <div className="flex items-center justify-between border-b pb-2">
                  <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider">Vales de Botadero</h4>
                  <button type="button" onClick={addVale} className="text-xs text-blue-600 font-bold hover:text-blue-800">+ Agregar Vale</button>
                </div>
                
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                  {formData.vales.map((vale, index) => (
                    <div key={index} className="p-3 bg-slate-50 rounded-lg border border-slate-200 relative group">
                      {index > 0 && (
                        <button 
                          type="button"
                          onClick={() => {
                            const newVales = formData.vales.filter((_, i) => i !== index);
                            setFormData({...formData, vales: newVales});
                          }}
                          className="absolute -top-2 -right-2 bg-rose-500 text-white rounded-full p-1 opacity-0 group-hover:opacity-100 transition-opacity"
                        >
                          <X size={12} />
                        </button>
                      )}
                      <span className="absolute top-1 left-2 text-[10px] font-bold text-slate-400">#{index + 1}</span>
                      
                      <div className="mt-2 space-y-2">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Consecutivo Vale</label>
                          <input 
                            type="text" 
                            value={vale.consecutivo}
                            onChange={e => handleValeChange(index, 'consecutivo', e.target.value)}
                            placeholder="Ej. 433"
                            className="w-full text-xs border-slate-300 rounded py-1 px-2 focus:ring-blue-500 focus:border-blue-500" 
                          />
                        </div>
                        <div className="grid grid-cols-2 gap-2">
                          <div>
                            <label className="block text-[11px] font-semibold text-slate-600 mb-1">Peso Entrada</label>
                            <input 
                              type="number" 
                              value={vale.pesoEntrada}
                              onChange={e => handleValeChange(index, 'pesoEntrada', e.target.value)}
                              placeholder="kg"
                              className="w-full text-xs border-slate-300 rounded py-1 px-2 focus:ring-blue-500 focus:border-blue-500" 
                            />
                          </div>
                          <div>
                            <label className="block text-[11px] font-semibold text-slate-600 mb-1">Peso Salida</label>
                            <input 
                              type="number" 
                              value={vale.pesoSalida}
                              onChange={e => handleValeChange(index, 'pesoSalida', e.target.value)}
                              placeholder="kg"
                              className="w-full text-xs border-slate-300 rounded py-1 px-2 focus:ring-blue-500 focus:border-blue-500" 
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex justify-end pt-2">
                <button 
                  type="submit" 
                  disabled={isSaving}
                  className="bg-yellow-500 hover:bg-yellow-600 text-slate-900 px-6 py-2.5 rounded-xl font-black transition-colors shadow-md disabled:opacity-60 flex items-center gap-2 text-sm"
                >
                  {isSaving ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
                  <span>{isSaving ? 'Guardando...' : 'Guardar Reporte'}</span>
                </button>
              </div>

            </form>
          </div>
        )}
      </div>

    </div>
  );
};

export default DataEntryWorkspace;