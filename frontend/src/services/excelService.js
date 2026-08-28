import * as XLSX from 'xlsx';

/**
 * Genera un archivo Excel (.xlsx) a partir de los datos de auditoría de faltantes y fuerza su descarga en el navegador.
 */
export const downloadAuditExcel = (data, filename = 'reporte_faltantes.xlsx') => {
  const worksheet = XLSX.utils.json_to_sheet(data);

  const colWidths = [
    { wch: 24 }, // Proyecto
    { wch: 30 }, // Nombre Conductor / Operario
    { wch: 24 }, // Días Faltantes
    { wch: 18 }, // Cantidad Faltantes
    { wch: 60 }  // Mensaje Generado
  ];
  worksheet['!cols'] = colWidths;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Faltantes');

  XLSX.writeFile(workbook, filename);
};

/**
 * Crea un buffer de Excel fresco con encabezados y el primer registro.
 */
export const createFreshExcelBuffer = (newRowData) => {
  const headers = ['PROYECTO', 'NOMBRE', 'IDENTIFICACION', 'CELULAR', 'USUARIO', 'CONTRASEÑA'];
  const data = [
    headers,
    [
      newRowData.proyecto,
      newRowData.nombre,
      newRowData.identificacion,
      newRowData.celular,
      newRowData.usuario,
      newRowData.contraseña
    ]
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Listado');
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
};

/**
 * Añade una fila con los datos de un nuevo trabajador/conductor a un array de buffer de Excel.
 */
export const appendPersonToExcelBuffer = (arrayBuffer, newRowData) => {
  try {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    jsonData.push([
      newRowData.proyecto,
      newRowData.nombre,
      newRowData.identificacion,
      newRowData.celular,
      newRowData.usuario,
      newRowData.contraseña
    ]);

    const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
    workbook.Sheets[firstSheetName] = newWorksheet;

    return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  } catch (error) {
    console.error("Error procesando Excel en memoria:", error);
    throw error;
  }
};
