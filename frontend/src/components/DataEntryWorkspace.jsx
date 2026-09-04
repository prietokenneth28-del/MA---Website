import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { getGraphClient } from '../services/graphService';
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
  RotateCcw
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

// --- Tipos de Pago disponibles para el registro diario ---
const TIPOS_PAGO = ['PESO', 'VIAJE', 'DISPONIBILIDAD', 'DISPONIBLE', 'NO PROGRAMADO'];

// --- Configuración del Excel maestro (acumulativo) en OneDrive ---
// Todos los registros de digitación (Conductores y Operadores) se guardan como filas
// nuevas en este único archivo, dentro de la carpeta raíz de Máquinas Amarillas.
const MASTER_EXCEL_PATH = "1. MAQUINAS AMARILLAS/Registros_Digitacion.xlsx";
const encodedMasterPath = MASTER_EXCEL_PATH.split('/').map(encodeURIComponent).join('/');
const GRAPH_ENDPOINT = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedMasterPath}:/content`;

const MASTER_HEADERS = [
  'Fecha Registro', 'Tipo', 'Proyecto', 'Trabajador', 'Fecha Reporte',
  'Consecutivo RDO', 'Horometro Inicial', 'Horometro Final', 'Tipo Pago',
  'Consecutivo Vale', 'Peso Entrada', 'Peso Salida'
];

// Construye una fila (arreglo) por cada vale del reporte, replicando los campos
// generales del registro diario. Si no hay vales con datos, igual se genera una
// fila única para no perder el registro del día.
const buildMasterRows = ({ tipo, proyecto, persona, fecha, formData }) => {
  const timestamp = new Date().toISOString();
  const baseFields = [
    timestamp, tipo, proyecto, persona, fecha,
    formData.consecutivoRegistro, formData.horometroInicial, formData.horometroFinal, formData.tipoPago
  ];

  const vales = formData.vales && formData.vales.length > 0
    ? formData.vales
    : [{ consecutivo: '', pesoEntrada: '', pesoSalida: '' }];

  return vales.map(v => [...baseFields, v.consecutivo, v.pesoEntrada, v.pesoSalida]);
};

// Descarga el Excel maestro existente desde OneDrive y lo agrega las filas nuevas.
// Si el archivo aún no existe (404), crea uno nuevo con encabezados.
const appendRowsToOneDriveMaster = async (accessToken, newRows) => {
  const headers = { Authorization: `Bearer ${accessToken}` };

  const getResponse = await fetch(GRAPH_ENDPOINT, { headers });

  let sheetRows;
  if (getResponse.status === 404) {
    // El archivo no existe todavía: se crea con encabezados + las filas nuevas
    sheetRows = [MASTER_HEADERS, ...newRows];
  } else if (getResponse.ok) {
    const buffer = await getResponse.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const existingSheet = workbook.Sheets[firstSheetName];
    const existingRows = XLSX.utils.sheet_to_json(existingSheet, { header: 1, defval: '' });
    // Si por algún motivo el archivo existente no tiene encabezados, los añadimos
    const hasHeaders = existingRows.length > 0 && existingRows[0][0] === MASTER_HEADERS[0];
    sheetRows = hasHeaders
      ? [...existingRows, ...newRows]
      : [MASTER_HEADERS, ...existingRows, ...newRows];
  } else {
    const errorText = await getResponse.text().catch(() => '');
    throw new Error(`No se pudo leer el Excel maestro (HTTP ${getResponse.status}): ${errorText}`);
  }

  const newSheet = XLSX.utils.aoa_to_sheet(sheetRows);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Registros');
  const outBuffer = XLSX.write(newWorkbook, { type: 'array', bookType: 'xlsx' });

  const putResponse = await fetch(GRAPH_ENDPOINT, {
    method: 'PUT',
    headers: {
      ...headers,
      'Content-Type': 'application/octet-stream'
    },
    body: outBuffer
  });

  if (!putResponse.ok) {
    const errorText = await putResponse.text().catch(() => '');
    throw new Error(`No se pudo guardar el Excel maestro (HTTP ${putResponse.status}): ${errorText}`);
  }
};

// Modo demo (sin accessToken): genera y descarga un Excel local con las filas del registro,
// para que el botón siga siendo funcional en pruebas locales sin OneDrive.
const downloadRowsLocally = (newRows, fileLabel) => {
  const sheet = XLSX.utils.aoa_to_sheet([MASTER_HEADERS, ...newRows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Registros');
  XLSX.writeFile(workbook, `Registro_${fileLabel}.xlsx`);
};

// --- Subcomponente: Visor de Imagen con Zoom y Pan ---
const ImagePanZoom = ({ src, alt, onRemove }) => {
  const [scale, setScale] = useState(1);
  const [rotation, setRotation] = useState(0);
  const [isDragging, setIsDragging] = useState(false);
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  const containerRef = useRef(null);

  // --- Estado para exportar la imagen (Drag & Drop hacia Sinco) ---
  // Mientras se mantiene presionada la tecla Shift, la imagen se vuelve arrastrable
  // (draggable nativo del navegador) en lugar de usar el pan del contenedor.
  const [isShiftHeld, setIsShiftHeld] = useState(false);
  const [isFetchingBlob, setIsFetchingBlob] = useState(false);
  const blobRef = useRef(null);

  useEffect(() => {
    const handleKeyDown = (e) => { if (e.key === 'Shift') setIsShiftHeld(true); };
    const handleKeyUp = (e) => { if (e.key === 'Shift') setIsShiftHeld(false); };
    window.addEventListener('keydown', handleKeyDown);
    window.addEventListener('keyup', handleKeyUp);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('keyup', handleKeyUp);
    };
  }, []);

  // Pre-descarga el Blob de la imagen en cuanto se activa el modo arrastre (Shift),
  // para tenerlo listo de forma síncrona cuando el usuario suelte el archivo sobre Sinco.
  useEffect(() => {
    if (isShiftHeld && src && !blobRef.current && !isFetchingBlob) {
      setIsFetchingBlob(true);
      fetch(src)
        .then(res => res.blob())
        .then(blob => { blobRef.current = blob; })
        .catch(err => console.error('[ImagePanZoom] No se pudo pre-cargar la imagen para arrastrar:', err))
        .finally(() => setIsFetchingBlob(false));
    }
  }, [isShiftHeld, src]);

  // Si cambia la imagen mostrada, invalidamos el blob cacheado de la anterior
  useEffect(() => {
    blobRef.current = null;
  }, [src]);

  const handleImageDragStart = (e) => {
    if (!blobRef.current) {
      // Todavía no terminó de descargarse el archivo; evitamos iniciar un drag vacío
      e.preventDefault();
      return;
    }
    const fileName = (alt && /\.(jpg|jpeg|png)$/i.test(alt)) ? alt : `${alt || 'imagen'}.jpg`;
    const file = new File([blobRef.current], fileName, { type: blobRef.current.type || 'image/jpeg' });
    e.dataTransfer.items.add(file);
    e.dataTransfer.effectAllowed = 'copy';
  };

  const handleWheel = (e) => {
    if (e.ctrlKey) {
      e.preventDefault();
      const scaleAdjust = e.deltaY * -0.01;
      setScale(Math.min(Math.max(0.5, scale + scaleAdjust), 4));
    }
  };

  const handleMouseDown = (e) => {
    // Con Shift presionado dejamos que el navegador gestione el drag nativo de la imagen
    // en vez de iniciar el paneo del contenedor.
    if (isShiftHeld) return;
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
      <div className="absolute top-2 right-2 z-10 flex space-x-1 opacity-0 group-hover:opacity-100 transition-opacity bg-black/50 p-1 rounded-lg">
        <button onClick={() => setScale(s => Math.min(s + 0.2, 4))} className="p-1 text-white hover:bg-white/20 rounded"><ZoomIn size={16}/></button>
        <button onClick={() => setScale(s => Math.max(s - 0.2, 0.5))} className="p-1 text-white hover:bg-white/20 rounded"><ZoomOut size={16}/></button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={() => setRotation(r => (r - 90 + 360) % 360)} className="p-1 text-white hover:bg-white/20 rounded"><RotateCcw size={16}/></button>
        <button onClick={() => setRotation(r => (r + 90) % 360)} className="p-1 text-white hover:bg-white/20 rounded"><RotateCw size={16}/></button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={() => { setScale(1); setPosition({x:0, y:0}); setRotation(0); }} className="p-1 text-white hover:bg-white/20 rounded"><Maximize size={16}/></button>
        <div className="w-px bg-white/20 mx-1"></div>
        <button onClick={onRemove} className="p-1 text-rose-400 hover:bg-rose-500/20 rounded"><X size={16}/></button>
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
          draggable={isShiftHeld}
          onDragStart={handleImageDragStart}
          className={`max-w-full max-h-full object-contain ${isShiftHeld ? 'cursor-grab' : 'pointer-events-none'}`}
          referrerPolicy="no-referrer"
          onError={(e) => {
            console.error(
              `[ImagePanZoom] Falló la carga de "${alt}". ` +
              `Revisa la pestaña Network/Console para ver el código de error (CORS, 401, CSP, etc). URL:`,
              src
            );
          }}
        />
      </div>
      
      {!src && (
        <div className="absolute inset-0 flex flex-col items-center justify-center text-white/60 text-xs gap-2">
          <ImageIcon className="w-8 h-8 opacity-40" />
          <span>Sin URL de imagen (revisa la consola)</span>
        </div>
      )}
      
      {/* Indicador de ayuda */}
      <div className="absolute bottom-2 left-2 text-[10px] text-white/50 bg-black/40 px-2 py-1 rounded pointer-events-none">
        Ctrl + Scroll para Zoom | Arrastrar para mover | Shift + Arrastrar para subir a Sinco
        {isFetchingBlob && ' (preparando archivo...)'}
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
  const [selectedImages, setSelectedImages] = useState([]); // Arreglo de objetos {id, url, name}

  // Loading separado para las listas de Proyectos/Trabajadores (evita pisar el spinner de "Buscar Archivos")
  const [isLoadingLists, setIsLoadingLists] = useState(false);
  
  // --- Estados del Formulario ---
  const [formData, setFormData] = useState({
    consecutivoRegistro: '',
    horometroInicial: '',
    horometroFinal: '',
    tipoPago: '',
    vales: [{ consecutivo: '', pesoEntrada: '', pesoSalida: '' }] // Soporte para múltiples vales
  });

  // Estado de guardado del reporte (Excel maestro en OneDrive)
  const [isSaving, setIsSaving] = useState(false);

  // 1. Cargar proyectos reales cuando cambia el tipo de personal o el token de acceso
  useEffect(() => {
    const fetchProjects = async () => {
      if (!accessToken) {
        // Fallback de demostración
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
        const basePath = BASE_PATHS[tipo]; // e.g. "1. MAQUINAS AMARILLAS/CONDUCTORES"
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

  // 2. Cargar conductores/operadores reales cuando cambia el proyecto seleccionado o el token
  useEffect(() => {
    const fetchWorkers = async () => {
      if (!proyecto) {
        setPersonasList([]);
        return;
      }

      if (!accessToken) {
        // Fallback de demostración
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

  // --- Funciones MS Graph ---
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
      // MODO DEMO: Simular carga de imágenes usando mock URLs
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

    // MODO PRODUCCIÓN: Consultar Graph API
    try {
      const client = getGraphClient(accessToken);
      const response = await client.api(`/me/drive/root:/${folderPath}:/children`)
        .get();

      // Filtrar por extensiones válidas de imágenes
      const rawItems = (response.value || [])
        .filter(item => item.file && /\.(jpg|jpeg|png)$/i.test(item.name));

      // Diagnóstico: detectar items sin downloadUrl (Graph a veces no lo incluye
      // si el $select se reescribe o si faltan permisos de Files.Read)
      const missingUrl = rawItems.filter(item => !item['@microsoft.graph.downloadUrl']);
      if (missingUrl.length > 0) {
        console.warn(
          '[DataEntryWorkspace] Estos archivos no trajeron @microsoft.graph.downloadUrl:',
          missingUrl.map(i => i.name)
        );
      }

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

  // --- Funciones UI ---
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

  const handleValeChange = (index, field, value) => {
    const newVales = [...formData.vales];
    newVales[index][field] = value;
    setFormData({ ...formData, vales: newVales });
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
      showToast("Selecciona Proyecto, Trabajador y Fecha antes de guardar", "warning");
      return;
    }
    if (!formData.consecutivoRegistro) {
      showToast("Ingresa el Consecutivo RDO antes de guardar", "warning");
      return;
    }
    if (!formData.tipoPago) {
      showToast("Selecciona el Tipo de Pago antes de guardar", "warning");
      return;
    }

    const newRows = buildMasterRows({ tipo, proyecto, persona, fecha, formData });

    setIsSaving(true);
    try {
      if (accessToken) {
        await appendRowsToOneDriveMaster(accessToken, newRows);
        showToast("Reporte guardado en el Excel maestro de OneDrive", "success");
      } else {
        // MODO DEMO: no hay sesión real de OneDrive, se descarga localmente
        downloadRowsLocally(newRows, `${persona}_${fecha}`);
        showToast("Sin conexión a OneDrive: se descargó el Excel localmente (Modo Demo)", "info");
      }

      // Se limpia el formulario para digitar el siguiente registro,
      // manteniendo la búsqueda (Proyecto/Trabajador/Fecha) seleccionada.
      setFormData({
        consecutivoRegistro: '',
        horometroInicial: '',
        horometroFinal: '',
        tipoPago: '',
        vales: [{ consecutivo: '', pesoEntrada: '', pesoSalida: '' }]
      });
    } catch (error) {
      console.error("Error guardando el reporte en el Excel maestro:", error);
      showToast("Error al guardar el reporte en OneDrive", "error");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="flex flex-col h-full bg-slate-50 min-h-[500px]">
      {/* HEADER DE BÚSQUEDA */}
      <div className="bg-white border-b border-slate-200 p-4 shadow-sm z-10 flex-shrink-0">
        <div className="flex flex-wrap items-end gap-4">
          <div className="flex-1 min-w-[150px]">
            <label className="block text-xs font-semibold text-slate-500 mb-1">Tipo</label>
            <select value={tipo} onChange={(e) => setTipo(e.target.value)} className="w-full text-sm border-slate-300 rounded-lg focus:ring-yellow-500 focus:border-yellow-500">
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

      {/* ÁREA PRINCIPAL */}
      <div className="flex-1 flex overflow-hidden min-h-0">
        
        {/* PANEL IZQUIERDO: Miniaturas de imágenes encontradas */}
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
                      onError={(e) => {
                        console.error(
                          `[DataEntryWorkspace] Falló la carga de imagen "${img.name}". ` +
                          `Revisa la pestaña Network/Console para ver el código de error (CORS, 401, CSP, etc). URL:`,
                          img.url
                        );
                        e.target.dataset.broken = 'true';
                      }}
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

        {/* WORKSPACE CENTRAL: Visualizador de imágenes divididas */}
        <div className="flex-1 bg-slate-800 p-2 flex flex-col min-w-0 overflow-hidden">
          {selectedImages.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center text-slate-400">
              <LayoutGrid className="w-16 h-16 mb-4 opacity-30" />
              <p className="text-lg">Selecciona imágenes del panel izquierdo</p>
              <p className="text-sm">Puedes abrir hasta 3 al mismo tiempo.</p>
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

        {/* PANEL DERECHO: Formulario de Digitación */}
        <div className="w-80 bg-white border-l border-slate-200 flex flex-col overflow-hidden flex-shrink-0 shadow-lg z-10">
          <div className="p-4 bg-slate-900 text-white flex items-center space-x-2 flex-shrink-0">
            <FileEdit className="w-5 h-5 text-yellow-400" />
            <h3 className="font-bold">Digitación de Datos</h3>
          </div>
          
          <div className="flex-1 overflow-y-auto p-4 space-y-6">
            <form id="dataEntryForm" onSubmit={handleSubmit} className="space-y-5">
              
              {/* Sección Registro Diario */}
              <div className="space-y-3">
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider border-b pb-1">Registro Diario de Operación</h4>
                
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

                <div className="grid grid-cols-2 gap-3">
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
              </div>

              {/* Sección Vales Botadero */}
              <div className="space-y-3">
                <div className="flex items-center justify-between border-b pb-1">
                  <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Vales de Botadero</h4>
                  <button type="button" onClick={addVale} className="text-xs text-blue-600 font-bold hover:text-blue-800">+ Agregar</button>
                </div>
                
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

            </form>
          </div>
          
          {/* FOOTER DEL FORMULARIO */}
          <div className="p-4 border-t border-slate-200 bg-slate-50 flex-shrink-0">
            <button 
              form="dataEntryForm"
              type="submit" 
              disabled={isSaving}
              className="w-full flex justify-center items-center space-x-2 bg-yellow-500 hover:bg-yellow-600 text-slate-900 py-3 rounded-xl font-black transition-colors shadow-md disabled:opacity-60"
            >
              {isSaving ? <Loader2 className="w-5 h-5 animate-spin" /> : <Save className="w-5 h-5" />}
              <span>{isSaving ? 'Guardando...' : 'Guardar Reporte'}</span>
            </button>
          </div>
        </div>

      </div>
    </div>
  );
};

export default DataEntryWorkspace;