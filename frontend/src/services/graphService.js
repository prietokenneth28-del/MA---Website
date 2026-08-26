import { Client } from "@microsoft/microsoft-graph-client";

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
      "@microsoft.graph.conflictBehavior": "skip"
    });

    // 2. Crear carpetas de meses y días
    for (let monthNum = 1; monthNum <= 12; monthNum++) {
      const monthName = MESES[monthNum];
      const monthFolderName = `${String(monthNum).padStart(2, '0')}_${monthName}`;
      const daysInMonth = new Date(year, monthNum, 0).getDate();

      const monthPath = `${targetFolderPath}/${monthFolderName}`;
      
      // Crear carpeta del mes
      await client.api(`/me/drive/root:/${monthPath}:`).patch({
        folder: {},
        "@microsoft.graph.conflictBehavior": "skip"
      });

      // Crear subcarpetas para cada día del mes (en lotes o peticiones)
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
 * Escanea OneDrive buscando fotos en las carpetas de días para un rango dado de fechas.
 */
export const scanMissingReportsOneDrive = async (accessToken, type, projects, startDate, endDate, onProgress) => {
  if (!accessToken) {
    // Simulación para modo demo
    for (let i = 1; i <= 5; i++) {
      if (onProgress) onProgress(i * 20);
      await new Promise(res => setTimeout(res, 400));
    }
    
    // Retorna datos de muestra representativos
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

  const start = new Date(startDate);
  const end = new Date(endDate);

  try {
    let totalProjects = projects.length;
    let completedProjects = 0;

    for (const project of projects) {
      const projectPath = `${basePath}/${project}`;
      
      try {
        // Listar trabajadores en el proyecto
        const response = await client.api(`/me/drive/root:/${projectPath}:/children`).get();
        const workers = response.value.filter(item => item.folder);

        for (const worker of workers) {
          const workerName = worker.name;
          const firstName = workerName.split(' ')[0];
          const missingDates = [];

          let currDate = new Date(start);
          while (currDate <= end) {
            const m = currDate.getMonth() + 1;
            const d = currDate.getDate();
            const monthFolderName = `${String(m).padStart(2, '0')}_${MESES[m]}`;
            const dayFolderName = String(d).padStart(2, '0');

            const dayPath = `${projectPath}/${workerName}/${monthFolderName}/${dayFolderName}`;

            try {
              const dayFilesResponse = await client.api(`/me/drive/root:/${dayPath}:/children`).get();
              const hasImages = dayFilesResponse.value.some(file => 
                /\.(jpg|jpeg|png)$/i.test(file.name)
              );

              if (!hasImages) {
                missingDates.push(new Date(currDate));
              }
            } catch (err) {
              // Carpeta del día no existe
              missingDates.push(new Date(currDate));
            }

            currDate.setDate(currDate.getDate() + 1);
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
      } catch (err) {
        console.warn(`No se pudo acceder a la carpeta del proyecto: ${projectPath}`, err);
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

