import { Client } from "@microsoft/microsoft-graph-client";
import { loginRequest } from "../authConfig";

/**
 * Obtiene un Access Token válido de MSAL de forma silenciosa (o con popup de respaldo si expiró).
 */
export const getOrRefreshToken = async (msalInstance) => {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  if (!accounts || accounts.length === 0) return null;

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0]
    });
    return response.accessToken;
  } catch (error) {
    console.warn("Renovación silenciosa de token falló, solicitando mediante popup:", error);
    try {
      const response = await msalInstance.acquireTokenPopup(loginRequest);
      return response.accessToken;
    } catch (err) {
      console.error("No se pudo obtener el token de Microsoft Graph:", err);
      return null;
    }
  }
};

// Inicializa el cliente de Microsoft Graph usando el token de acceso obtenido por MSAL
export const getGraphClient = (accessToken) => {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
};

const MESES = {
  1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
  5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
  9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
};

const BASE_PATHS = {
  CONDUCTOR: "MAQUINAS AMARILLAS/CONDUCTORES",
  OPERARIO: "MAQUINAS AMARILLAS/OPERADORES"
};

/**
 * Crea la estructura anual de carpetas (Proyecto -> Nombre -> Meses -> Días) en OneDrive.
 */
export const createFolderStructureOneDrive = async (accessToken, role, project, personName, year = new Date().getFullYear()) => {
  if (!accessToken) {
    // Simulación en caso de modo demo sin token real
    await new Promise(res => setTimeout(res, 1200));
    return { success: true, mode: "demo" };
  }

  const client = getGraphClient(accessToken);
  const basePath = BASE_PATHS[role] || BASE_PATHS.CONDUCTOR;
  const targetFolderPath = `${basePath}/${project}/${personName}`;

  try {
    // 1. Crear carpeta del trabajador si no existe
    await client.api(`/me/drive/root:/${targetFolderPath}:`).patch({
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail"
    });

    // 2. Crear carpetas de meses
    for (let monthNum = 1; monthNum <= 12; monthNum++) {
      const monthName = MESES[monthNum];
      const monthFolderName = `${String(monthNum).padStart(2, '0')}_${monthName}`;
      const daysInMonth = new Date(year, monthNum, 0).getDate();

      const monthPath = `${targetFolderPath}/${monthFolderName}`;
      
      // Crear carpeta del mes
      await client.api(`/me/drive/root:/${monthPath}:`).patch({
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail"
      });

      // Crear subcarpetas para cada día del mes
      for (let day = 1; day <= daysInMonth; day++) {
        const dayFolderName = String(day).padStart(2, '0');
        const dayPath = `${monthPath}/${dayFolderName}`;
        
        await client.api(`/me/drive/root:/${dayPath}:`).patch({
          folder: {},
          "@microsoft.graph.conflictBehavior": "skip"
        });
      }
    }

    return { success: true, mode: "graph" };
  } catch (error) {
    console.error("Error al crear carpetas en OneDrive:", error);
    throw error;
  }
};

/**
 * Escanea OneDrive de forma ultra-optimizada buscando días sin reportes para una categoría y rango de fechas.
 * Consulta las carpetas de meses en 1 sola llamada batch por trabajador en lugar de peticiones individuales por día.
 */
export const scanMissingReportsOneDrive = async (accessToken, type, projects, startDate, endDate, onProgress) => {
  if (!accessToken) {
    // Simulación para modo demo
    for (let i = 1; i <= 5; i++) {
      if (onProgress) onProgress(i * 20);
      await new Promise(res => setTimeout(res, 300));
    }
    
    return [
      {
        Proyecto: projects[0] || "AGUAS BOGOTA",
        "Nombre Conductor": "CARLOS RODRIGUEZ",
        "Días Faltantes": "02/08, 04/08",
        "Cantidad Faltantes": 2,
        "Mensaje Generado": "Buenos días Carlos, ¿podrías enviarme los reportes de los días 02 y 04 de agosto?"
      },
      {
        Proyecto: projects[1] || "QUORA",
        "Nombre Conductor": "JOSE GOMEZ",
        "Días Faltantes": "01/08, 03/08, 05/08",
        "Cantidad Faltantes": 3,
        "Mensaje Generado": "Buenos días Jose, ¿podrías enviarme los reportes de los días 01, 03 y 05 de agosto?"
      }
    ];
  }

  const client = getGraphClient(accessToken);
  const roleKey = type === "operadores" ? "OPERARIO" : "CONDUCTOR";
  const basePath = BASE_PATHS[roleKey];
  const missingData = [];

  // Parsear fechas de inicio y fin (YYYY-MM-DD)
  const [sYear, sMonth, sDay] = startDate.split('-').map(Number);
  const [eYear, eMonth, eDay] = endDate.split('-').map(Number);

  const startDt = new Date(sYear, sMonth - 1, sDay);
  const endDt = new Date(eYear, eMonth - 1, eDay);

  try {
    let totalProjects = projects.length;
    let completedProjects = 0;

    for (const project of projects) {
      const projectPath = `${basePath}/${project}`;
      
      try {
        // Listar trabajadores dentro del proyecto
        const response = await client.api(`/me/drive/root:/${projectPath}:/children`).get();
        const workers = (response.value || []).filter(item => item.folder);

        for (const worker of workers) {
          const workerName = worker.name;
          const firstName = workerName.split(' ')[0];
          const missingDates = [];

          // Agrupar días por mes para consultar 1 sola vez la carpeta del mes por trabajador
          let curr = new Date(startDt);
          const monthMap = new Map(); // monthKey -> Array de objetos { dayNum, fullDate }

          while (curr <= endDt) {
            const m = curr.getMonth() + 1;
            const d = curr.getDate();
            const monthKey = `${String(m).padStart(2, '0')}_${MESES[m]}`;

            if (!monthMap.has(monthKey)) {
              monthMap.set(monthKey, []);
            }
            monthMap.get(monthKey).push({ dayNum: d, fullDate: new Date(curr) });

            curr.setDate(curr.getDate() + 1);
          }

          // Consultar cada mes escaneando sus subcarpetas de días
          for (const [monthFolderName, daysToCheck] of monthMap.entries()) {
            const monthPath = `${projectPath}/${workerName}/${monthFolderName}`;

            try {
              // Obtener todos los días creados en este mes en 1 sola consulta
              const monthChildrenResp = await client.api(`/me/drive/root:/${monthPath}:/children`).get();
              const dayFolders = monthChildrenResp.value || [];

              // Crear un mapa de día -> childCount (número de archivos subidos dentro del día)
              const dayFolderMap = new Map();
              for (const df of dayFolders) {
                if (df.folder) {
                  const dayNumStr = df.name;
                  dayFolderMap.set(parseInt(dayNumStr, 10), df.folder.childCount || 0);
                }
              }

              for (const { dayNum, fullDate } of daysToCheck) {
                const filesCount = dayFolderMap.get(dayNum);
                
                // Si la carpeta del día no existe o su número de archivos es 0, cuenta como faltante
                if (filesCount === undefined || filesCount === 0) {
                  missingDates.push(fullDate);
                }
              }
            } catch (errMonth) {
              // Si la carpeta del mes no existe en OneDrive, todos los días de ese mes son faltantes
              for (const { fullDate } of daysToCheck) {
                missingDates.push(fullDate);
              }
            }
          }

          if (missingDates.length > 0) {
            const diasColumna = missingDates.map(dt => `${String(dt.getDate()).padStart(2, '0')}/${String(dt.getMonth() + 1).padStart(2, '0')}`).join(', ');
            const diasMensaje = missingDates.map(dt => String(dt.getDate()).padStart(2, '0')).join(', ');
            const mesNombre = MESES[missingDates[0].getMonth() + 1].toLowerCase();

            missingData.push({
              Proyecto: project,
              "Nombre Conductor": workerName,
              "Días Faltantes": diasColumna,
              "Cantidad Faltantes": missingDates.length,
              "Mensaje Generado": `Buenos días ${firstName}, ¿podrías enviarme los reportes de los días ${diasMensaje} de ${mesNombre}?`
            });
          }
        }
      } catch (errProj) {
        console.warn(`No se encontró o no se pudo leer la carpeta del proyecto: ${projectPath}`, errProj);
      }

      completedProjects++;
      if (onProgress) {
        onProgress(Math.round((completedProjects / totalProjects) * 100));
      }
    }

    return missingData;
  } catch (error) {
    console.error("Error en escaneo de auditoría en OneDrive:", error);
    throw error;
  }
};
