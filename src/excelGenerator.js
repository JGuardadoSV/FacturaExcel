import ExcelJS from 'exceljs';
import {
  departamentos,
  municipios,
  actividadesEconomicas,
  tiposDocumento,
  condicionesPago,
  tiposProducto
} from './catalogData.js';

export async function generateExcel(emisorData) {
  const workbook = new ExcelJS.Workbook();

  workbook.creator = 'Sistema FacturaExcel';
  workbook.created = new Date();
  workbook.modified = new Date();

  // Crear todas las hojas en orden
  await createDashboardSheet(workbook, emisorData);
  await createEmisorSheet(workbook, emisorData);
  await createCorrelativosSheet(workbook);
  await createCatalogSheets(workbook);
  await createClientesSheet(workbook);
  await createProductosSheet(workbook);
  await createFacturaConsumidorFinalSheet(workbook);
  await createFacturaCreditoFiscalSheet(workbook);
  await createFacturaSujetoExcluidoSheet(workbook);
  await createNotasCreditoSheet(workbook);
  await createRegistroFacturasSheet(workbook);
  await createFacturasAnuladasSheet(workbook);

  return workbook;
}

// ============================================================================
// HOJA 1: DASHBOARD
// ============================================================================
async function createDashboardSheet(workbook, emisorData) {
  const sheet = workbook.addWorksheet('Dashboard', {
    properties: { tabColor: { argb: 'FF4472C4' } }
  });

  sheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 35 },
    { width: 35 }
  ];

  // Título principal
  sheet.mergeCells('B2:D2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = 'SISTEMA DE FACTURACIÓN ELECTRÓNICA';
  titleCell.font = { size: 20, bold: true, color: { argb: 'FF4472C4' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getRow(2).height = 40;

  // Información del emisor
  sheet.mergeCells('B4:D4');
  const emisorTitle = sheet.getCell('B4');
  emisorTitle.value = 'INFORMACIÓN DEL EMISOR';
  emisorTitle.font = { size: 14, bold: true, color: { argb: 'FF2E75B6' } };
  emisorTitle.alignment = { horizontal: 'center' };
  emisorTitle.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9E2F3' }
  };

  const emisorInfo = [
    ['Nombre o Razón Social:', emisorData.razonSocial],
    ['NIT:', emisorData.nit],
    ['NRC:', emisorData.nrc],
    ['Actividad Económica:', emisorData.actividadEconomica],
    ['Dirección:', emisorData.direccion],
    ['Teléfono:', emisorData.telefono],
    ['Correo:', emisorData.correo]
  ];

  let row = 6;
  emisorInfo.forEach(([label, value]) => {
    sheet.getCell(`B${row}`).value = label;
    sheet.getCell(`B${row}`).font = { bold: true };
    sheet.mergeCells(`C${row}:D${row}`);
    sheet.getCell(`C${row}`).value = value || 'N/A';
    row++;
  });

  // Menú de navegación
  row += 2;
  sheet.mergeCells(`B${row}:D${row}`);
  const menuTitle = sheet.getCell(`B${row}`);
  menuTitle.value = 'MENÚ DE NAVEGACIÓN';
  menuTitle.font = { size: 14, bold: true, color: { argb: 'FF2E75B6' } };
  menuTitle.alignment = { horizontal: 'center' };
  menuTitle.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9E2F3' }
  };

  row += 2;

  const menuItems = [
    { title: '📋 CATÁLOGOS', items: [
      { name: 'Ver Catálogos', sheet: 'Departamentos' },
      { name: 'Gestionar Clientes', sheet: 'Clientes' },
      { name: 'Gestionar Productos', sheet: 'Productos' }
    ]},
    { title: '📄 FACTURACIÓN', items: [
      { name: 'Factura Consumidor Final', sheet: 'Factura_ConsumidorFinal' },
      { name: 'Crédito Fiscal', sheet: 'Factura_CreditoFiscal' },
      { name: 'Sujeto Excluido', sheet: 'Factura_SujetoExcluido' },
      { name: 'Notas de Crédito', sheet: 'Notas_Credito' }
    ]},
    { title: '📊 REPORTES', items: [
      { name: 'Registro de Facturas', sheet: 'Registro_Facturas' },
      { name: 'Facturas Anuladas', sheet: 'Facturas_Anuladas' }
    ]}
  ];

  menuItems.forEach(section => {
    sheet.mergeCells(`B${row}:D${row}`);
    const sectionCell = sheet.getCell(`B${row}`);
    sectionCell.value = section.title;
    sectionCell.font = { size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    sectionCell.alignment = { horizontal: 'center' };
    sectionCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF70AD47' }
    };
    row++;

    section.items.forEach(item => {
      sheet.mergeCells(`B${row}:D${row}`);
      const itemCell = sheet.getCell(`B${row}`);
      itemCell.value = {
        text: `   → ${item.name}`,
        hyperlink: `#'${item.sheet}'!A1`,
        tooltip: `Ir a ${item.name}`
      };
      itemCell.font = { size: 11, color: { argb: 'FF0563C1' }, underline: true };
      itemCell.alignment = { horizontal: 'left', vertical: 'middle' };
      itemCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2EFDA' }
      };
      itemCell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
      sheet.getRow(row).height = 25;
      row++;
    });
    row++;
  });

  // Instrucciones
  row += 2;
  sheet.mergeCells(`B${row}:D${row}`);
  const instrCell = sheet.getCell(`B${row}`);
  instrCell.value = '💡 Haga clic en los enlaces para navegar entre las hojas';
  instrCell.font = { size: 10, italic: true, color: { argb: 'FF7F7F7F' } };
  instrCell.alignment = { horizontal: 'center' };
}

// ============================================================================
// HOJA 2: DATOS EMISOR
// ============================================================================
async function createEmisorSheet(workbook, emisorData) {
  const sheet = workbook.addWorksheet('Emisor', {
    properties: { tabColor: { argb: 'FFED7D31' } }
  });

  sheet.columns = [
    { width: 5 },
    { width: 30 },
    { width: 50 }
  ];

  // Título
  sheet.mergeCells('B2:C2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = 'DATOS DEL EMISOR';
  titleCell.font = { size: 16, bold: true };
  titleCell.alignment = { horizontal: 'center' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFED7D31' }
  };

  const fields = [
    ['Razón Social:', emisorData.razonSocial],
    ['Nombre Comercial:', emisorData.nombreComercial],
    ['NIT:', emisorData.nit],
    ['NRC:', emisorData.nrc],
    ['Actividad Económica:', emisorData.actividadEconomica],
    ['Dirección:', emisorData.direccion],
    ['Departamento:', emisorData.departamento],
    ['Municipio:', emisorData.municipio],
    ['Teléfono:', emisorData.telefono],
    ['Correo Electrónico:', emisorData.correo]
  ];

  let row = 4;
  fields.forEach(([label, value]) => {
    sheet.getCell(`B${row}`).value = label;
    sheet.getCell(`B${row}`).font = { bold: true };
    sheet.getCell(`C${row}`).value = value || '';
    sheet.getCell(`C${row}`).border = {
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    };
    row++;
  });

  // Enlace de regreso
  row += 2;
  sheet.getCell(`B${row}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`B${row}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: CORRELATIVOS / CONFIGURACIÓN
// ============================================================================
async function createCorrelativosSheet(workbook) {
  const sheet = workbook.addWorksheet('Correlativos', {
    properties: { tabColor: { argb: 'FFFFC000' } }
  });

  sheet.columns = [
    { width: 5 },
    { width: 35 },
    { width: 20 },
    { width: 50 }
  ];

  // Título
  sheet.mergeCells('B2:D2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = 'CONTROL DE CORRELATIVOS';
  titleCell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFC000' }
  };
  sheet.getRow(2).height = 35;

  // Instrucciones
  sheet.mergeCells('B4:D4');
  const instrCell = sheet.getCell('B4');
  instrCell.value = 'Configure aquí los números correlativos para cada tipo de factura';
  instrCell.font = { italic: true, size: 10 };
  instrCell.alignment = { horizontal: 'center' };

  // Headers
  const headerRow = 6;
  sheet.getCell(`B${headerRow}`).value = 'Tipo de Documento';
  sheet.getCell(`C${headerRow}`).value = 'Próximo Correlativo';
  sheet.getCell(`D${headerRow}`).value = 'Descripción';

  ['B', 'C', 'D'].forEach(col => {
    const cell = sheet.getCell(`${col}${headerRow}`);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // Datos de correlativos
  const correlativos = [
    { tipo: 'Consumidor Final', correlativo: 1, desc: 'Facturas para consumidor final (CF)' },
    { tipo: 'Crédito Fiscal', correlativo: 1, desc: 'Comprobante de Crédito Fiscal (CCF)' },
    { tipo: 'Sujeto Excluido', correlativo: 1, desc: 'Facturas de Sujeto Excluido (FSE)' },
    { tipo: 'Notas de Crédito', correlativo: 1, desc: 'Notas de crédito para anulaciones' }
  ];

  let row = 7;
  correlativos.forEach((item) => {
    sheet.getCell(`B${row}`).value = item.tipo;
    sheet.getCell(`B${row}`).font = { bold: true };
    sheet.getCell(`B${row}`).border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };

    sheet.getCell(`C${row}`).value = item.correlativo;
    sheet.getCell(`C${row}`).numFmt = '0';
    sheet.getCell(`C${row}`).alignment = { horizontal: 'center' };
    sheet.getCell(`C${row}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }  // Amarillo para indicar que es editable
    };
    sheet.getCell(`C${row}`).border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };

    sheet.getCell(`D${row}`).value = item.desc;
    sheet.getCell(`D${row}`).font = { italic: true, size: 9 };
    sheet.getCell(`D${row}`).border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };

    row++;
  });

  // Nota importante
  row += 2;
  sheet.mergeCells(`B${row}:D${row}`);
  const noteCell = sheet.getCell(`B${row}`);
  noteCell.value = '⚠️ IMPORTANTE: Actualice manualmente el correlativo después de emitir cada factura';
  noteCell.font = { bold: true, size: 10, color: { argb: 'FFFF0000' } };
  noteCell.alignment = { horizontal: 'center' };
  noteCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFEB9C' }
  };

  row++;
  sheet.mergeCells(`B${row}:D${row}`);
  const note2Cell = sheet.getCell(`B${row}`);
  note2Cell.value = 'Las hojas de facturación leerán automáticamente el próximo número disponible de esta tabla';
  note2Cell.font = { italic: true, size: 9 };
  note2Cell.alignment = { horizontal: 'center' };

  // Enlace de regreso
  row += 3;
  sheet.getCell(`B${row}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`B${row}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJAS DE CATÁLOGOS
// ============================================================================
async function createCatalogSheets(workbook) {
  // Catálogo de Departamentos
  const deptSheet = workbook.addWorksheet('Departamentos', {
    properties: { tabColor: { argb: 'FFF4B084' } }
  });

  deptSheet.columns = [
    { header: 'Código', key: 'codigo', width: 12 },
    { header: 'Departamento', key: 'nombre', width: 30 }
  ];

  styleHeaderRow(deptSheet.getRow(1));
  departamentos.forEach(dept => deptSheet.addRow(dept));

  // Catálogo de Municipios
  const muniSheet = workbook.addWorksheet('Municipios', {
    properties: { tabColor: { argb: 'FFF4B084' } }
  });

  muniSheet.columns = [
    { header: 'Código', key: 'codigo', width: 12 },
    { header: 'Municipio', key: 'nombre', width: 35 },
    { header: 'Departamento', key: 'departamento', width: 15 }
  ];

  styleHeaderRow(muniSheet.getRow(1));
  municipios.forEach(muni => muniSheet.addRow(muni));

  // Catálogo de Actividades Económicas
  const actSheet = workbook.addWorksheet('Actividades_Economicas', {
    properties: { tabColor: { argb: 'FFF4B084' } }
  });

  actSheet.columns = [
    { header: 'Código', key: 'codigo', width: 12 },
    { header: 'Descripción', key: 'descripcion', width: 60 }
  ];

  styleHeaderRow(actSheet.getRow(1));
  actividadesEconomicas.forEach(act => actSheet.addRow(act));

  // Agregar enlaces de navegación a cada catálogo
  [deptSheet, muniSheet, actSheet].forEach(sheet => {
    const lastRow = sheet.rowCount + 2;
    sheet.getCell(`A${lastRow}`).value = {
      text: '← Volver al Dashboard',
      hyperlink: '#Dashboard!A1'
    };
    sheet.getCell(`A${lastRow}`).font = { color: { argb: 'FF0563C1' }, underline: true };
  });
}

// ============================================================================
// HOJA: CLIENTES
// ============================================================================
async function createClientesSheet(workbook) {
  const sheet = workbook.addWorksheet('Clientes', {
    properties: { tabColor: { argb: 'FF9DC3E6' } }
  });

  sheet.columns = [
    { header: 'Código Cliente', key: 'codigo', width: 15 },
    { header: 'Tipo Doc.', key: 'tipoDoc', width: 12 },
    { header: 'Número Documento', key: 'numDoc', width: 20 },
    { header: 'NIT', key: 'nit', width: 20 },
    { header: 'NRC', key: 'nrc', width: 15 },
    { header: 'Nombre / Razón Social', key: 'nombre', width: 40 },
    { header: 'Nombre Comercial', key: 'nombreComercial', width: 35 },
    { header: 'Actividad Económica', key: 'actividadEconomica', width: 25 },
    { header: 'Dirección', key: 'direccion', width: 45 },
    { header: 'Departamento', key: 'departamento', width: 20 },
    { header: 'Municipio', key: 'municipio', width: 25 },
    { header: 'Teléfono', key: 'telefono', width: 15 },
    { header: 'Correo', key: 'correo', width: 30 },
    { header: 'Gran Contribuyente', key: 'granContribuyente', width: 18 }
  ];

  styleHeaderRow(sheet.getRow(1));

  // Agregar validaciones de datos
  const tiposDocList = tiposDocumento.map(t => t.nombre).join(',');
  const deptList = departamentos.map(d => d.nombre).join(',');

  // Aplicar validaciones a las filas 2-1000
  for (let i = 2; i <= 1000; i++) {
    // Tipo de documento
    sheet.getCell(`B${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [`"${tiposDocList}"`]
    };

    // Departamento
    sheet.getCell(`J${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [`"${deptList}"`]
    };

    // Gran Contribuyente (Sí/No)
    sheet.getCell(`N${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Sí,No"']
    };
  }

  // Ejemplos de clientes
  const ejemplosClientes = [
    {
      codigo: 'CLI-001',
      tipoDoc: 'NIT - Número de Identificación Tributaria',
      numDoc: '',
      nit: '0614-123456-123-4',
      nrc: '12345-6',
      nombre: 'Empresa Ejemplo S.A. de C.V.',
      nombreComercial: 'Ejemplo',
      actividadEconomica: 'Comercio al por Mayor',
      direccion: 'Col. Escalón, San Salvador',
      departamento: 'San Salvador',
      municipio: 'San Salvador',
      telefono: '2222-2222',
      correo: 'info@ejemplo.com',
      granContribuyente: 'No'
    },
    {
      codigo: 'CLI-002',
      tipoDoc: 'NIT - Número de Identificación Tributaria',
      numDoc: '',
      nit: '0614-987654-321-8',
      nrc: '98765-4',
      nombre: 'Corporación ABC S.A. de C.V.',
      nombreComercial: 'ABC Corp',
      actividadEconomica: 'Servicios Profesionales',
      direccion: 'Col. San Benito, San Salvador',
      departamento: 'San Salvador',
      municipio: 'San Salvador',
      telefono: '2333-3333',
      correo: 'contacto@abc.com',
      granContribuyente: 'Sí'
    },
    {
      codigo: 'CLI-003',
      tipoDoc: 'DUI - Documento Único de Identidad',
      numDoc: '01234567-8',
      nit: '0614-010180-101-0',
      nrc: '',
      nombre: 'Juan Carlos Pérez López',
      nombreComercial: 'Ferretería El Tornillo',
      actividadEconomica: 'Comercio al por Menor',
      direccion: 'Santa Ana',
      departamento: 'Santa Ana',
      municipio: 'Santa Ana',
      telefono: '2440-1234',
      correo: 'jperez@email.com',
      granContribuyente: 'No'
    }
  ];

  ejemplosClientes.forEach(cliente => sheet.addRow(cliente));

  // Enlace de regreso
  const lastRow = sheet.rowCount + 2;
  sheet.getCell(`A${lastRow}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`A${lastRow}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: PRODUCTOS
// ============================================================================
async function createProductosSheet(workbook) {
  const sheet = workbook.addWorksheet('Productos', {
    properties: { tabColor: { argb: 'FFA9D08E' } }
  });

  sheet.columns = [
    { header: 'Código', key: 'codigo', width: 15 },
    { header: 'Descripción', key: 'descripcion', width: 45 },
    { header: 'Tipo Producto', key: 'tipo', width: 25 },
    { header: 'Precio Unitario', key: 'precio', width: 18 },
    { header: 'Unidad Medida', key: 'unidad', width: 18 }
  ];

  styleHeaderRow(sheet.getRow(1));

  // Validación tipo de producto
  const tiposList = tiposProducto.map(t => t.nombre).join(',');
  for (let i = 2; i <= 1000; i++) {
    sheet.getCell(`C${i}`).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: [`"${tiposList}"`]
    };
  }

  // Ejemplos de productos
  const ejemplosProductos = [
    { codigo: 'PROD-001', descripcion: 'Laptop Dell Latitude 7420', tipo: 'Gravado (13% IVA)', precio: 850.00, unidad: 'Unidad' },
    { codigo: 'PROD-002', descripcion: 'Medicamento Paracetamol 500mg', tipo: 'Exento (0% IVA)', precio: 2.50, unidad: 'Caja' },
    { codigo: 'PROD-003', descripcion: 'Libro de Contabilidad', tipo: 'Exento (0% IVA)', precio: 15.00, unidad: 'Unidad' },
    { codigo: 'PROD-004', descripcion: 'Mouse Inalámbrico Logitech', tipo: 'Gravado (13% IVA)', precio: 25.00, unidad: 'Unidad' },
    { codigo: 'PROD-005', descripcion: 'Escritorio de Oficina 1.20m', tipo: 'Gravado (13% IVA)', precio: 150.00, unidad: 'Unidad' },
    { codigo: 'PROD-006', descripcion: 'Resma de Papel Carta', tipo: 'Gravado (13% IVA)', precio: 4.50, unidad: 'Resma' },
    { codigo: 'SERV-001', descripcion: 'Consultoría Técnica', tipo: 'Gravado (13% IVA)', precio: 100.00, unidad: 'Hora' },
    { codigo: 'SERV-002', descripcion: 'Capacitación Empresarial', tipo: 'Gravado (13% IVA)', precio: 250.00, unidad: 'Sesión' }
  ];

  ejemplosProductos.forEach(prod => sheet.addRow(prod));

  // Formato de moneda
  for (let i = 2; i <= 1000; i++) {
    sheet.getCell(`D${i}`).numFmt = '$#,##0.00';
  }

  // Enlace de regreso
  const lastRow = sheet.rowCount + 2;
  sheet.getCell(`A${lastRow}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`A${lastRow}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: FACTURA CONSUMIDOR FINAL
// ============================================================================
async function createFacturaConsumidorFinalSheet(workbook) {
  const sheet = workbook.addWorksheet('Factura_ConsumidorFinal', {
    properties: { tabColor: { argb: 'FF92D050' } }
  });

  sheet.columns = [
    { width: 5 },   // #
    { width: 15 },  // Código producto
    { width: 35 },  // Descripción
    { width: 10 },  // Cantidad
    { width: 15 },  // Precio
    { width: 15 },  // Subtotal
    { width: 15 },  // IVA
    { width: 15 }   // Total
  ];

  // Título
  sheet.mergeCells('B2:I2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = 'FACTURA - CONSUMIDOR FINAL';
  titleCell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF70AD47' }
  };
  sheet.getRow(2).height = 30;

  // Datos de la factura
  sheet.getCell('B4').value = 'No. Factura:';
  sheet.getCell('B4').font = { bold: true };
  sheet.getCell('C4').value = { formula: 'Correlativos!C7' };  // Lee el correlativo de Consumidor Final
  sheet.getCell('C4').numFmt = '0';
  sheet.getCell('C4').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFCCFFCC' }  // Verde claro para indicar que es automático
  };
  sheet.getCell('C4').border = { bottom: { style: 'thin' } };

  sheet.getCell('F4').value = 'Fecha:';
  sheet.getCell('F4').font = { bold: true };
  sheet.getCell('G4').value = new Date();
  sheet.getCell('G4').numFmt = 'dd/mm/yyyy';
  sheet.getCell('G4').border = { bottom: { style: 'thin' } };

  sheet.getCell('B5').value = 'Cliente (opcional):';
  sheet.getCell('B5').font = { bold: true, size: 10 };
  sheet.mergeCells('C5:I5');
  sheet.getCell('C5').value = '';
  sheet.getCell('C5').border = { bottom: { style: 'thin' } };
  sheet.getCell('C5').note = {
    texts: [{ text: 'Puede dejar "CONSUMIDOR FINAL" o ingresar nombre del cliente' }]
  };

  // Tabla de productos
  const headerRow = 7;
  const headers = ['#', 'Código Producto', 'Descripción', 'Cantidad', 'Precio Unit.', 'Subtotal', 'IVA 13%', 'Total'];
  let col = 'B';
  headers.forEach(header => {
    sheet.getCell(`${col}${headerRow}`).value = header;
    sheet.getCell(`${col}${headerRow}`).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getCell(`${col}${headerRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF70AD47' }
    };
    sheet.getCell(`${col}${headerRow}`).alignment = { horizontal: 'center' };
    col = String.fromCharCode(col.charCodeAt(0) + 1);
  });

  // Filas para productos (8-27 = 20 líneas)
  for (let i = 8; i <= 27; i++) {
    sheet.getCell(`B${i}`).value = i - 7;
    sheet.getCell(`B${i}`).alignment = { horizontal: 'center' };

    // Lista desplegable de productos en columna C (Código Producto)
    sheet.getCell(`C${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['Productos!$A$2:$A$1000'],
      showErrorMessage: true,
      errorTitle: 'Código inválido',
      error: 'Seleccione un código de la lista de productos'
    };

    // VLOOKUP para descripción (columna D)
    sheet.getCell(`D${i}`).value = {
      formula: `IF(C${i}="","",IFERROR(VLOOKUP(C${i},Productos!$A$2:$E$1000,2,FALSE),""))`
    };

    // Columna E = Cantidad (INPUT MANUAL del usuario, sin fórmula)
    // Solo formato numérico
    sheet.getCell(`E${i}`).numFmt = '0';

    // VLOOKUP para precio unitario (columna F)
    sheet.getCell(`F${i}`).value = {
      formula: `IF(C${i}="","",IFERROR(VLOOKUP(C${i},Productos!$A$2:$E$1000,4,FALSE),""))`
    };
    sheet.getCell(`F${i}`).numFmt = '$#,##0.00';

    // Fórmula protegida para Subtotal (Cantidad * Precio) - columna G
    sheet.getCell(`G${i}`).value = {
      formula: `IF(AND(ISNUMBER(E${i}),ISNUMBER(F${i})),E${i}*F${i},"")`
    };
    sheet.getCell(`G${i}`).numFmt = '$#,##0.00';

    // Fórmula protegida para IVA (Subtotal * 13%) - columna H
    sheet.getCell(`H${i}`).value = {
      formula: `IF(ISNUMBER(G${i}),G${i}*0.13,"")`
    };
    sheet.getCell(`H${i}`).numFmt = '$#,##0.00';

    // Fórmula protegida para Total (En Consumidor Final = Subtotal, el IVA es solo informativo) - columna I
    sheet.getCell(`I${i}`).value = {
      formula: `IF(ISNUMBER(G${i}),G${i},"")`
    };
    sheet.getCell(`I${i}`).numFmt = '$#,##0.00';
  }

  // Ejemplo pre-llenado en la primera fila (fila 8)
  sheet.getCell('C8').value = 'PROD-001';  // Esto activará el VLOOKUP
  sheet.getCell('E8').value = 2;  // Cantidad de ejemplo

  // Totales
  const totRow = 29;
  sheet.mergeCells(`B${totRow}:F${totRow}`);
  sheet.getCell(`B${totRow}`).value = 'TOTALES:';
  sheet.getCell(`B${totRow}`).font = { bold: true, size: 12 };
  sheet.getCell(`B${totRow}`).alignment = { horizontal: 'right' };

  sheet.getCell(`G${totRow}`).value = { formula: `SUMIF(G8:G27,"<>",G8:G27)` };
  sheet.getCell(`G${totRow}`).numFmt = '$#,##0.00';
  sheet.getCell(`G${totRow}`).font = { bold: true };
  sheet.getCell(`G${totRow}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9E2F3' }
  };

  sheet.getCell(`H${totRow}`).value = { formula: `SUMIF(H8:H27,"<>",H8:H27)` };
  sheet.getCell(`H${totRow}`).numFmt = '$#,##0.00';
  sheet.getCell(`H${totRow}`).font = { bold: true };
  sheet.getCell(`H${totRow}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9E2F3' }
  };

  sheet.getCell(`I${totRow}`).value = { formula: `SUMIF(I8:I27,"<>",I8:I27)` };
  sheet.getCell(`I${totRow}`).numFmt = '$#,##0.00';
  sheet.getCell(`I${totRow}`).font = { bold: true, size: 12 };
  sheet.getCell(`I${totRow}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF92D050' }
  };

  // Instrucciones
  sheet.mergeCells(`B${totRow + 2}:I${totRow + 2}`);
  sheet.getCell(`B${totRow + 2}`).value = '⚠️ CÓMO USAR: 1) Vaya a hoja "Productos" y registre sus productos con códigos como PROD-001, PROD-002, etc. 2) Regrese aquí y haga clic en columna "Código Producto"';
  sheet.getCell(`B${totRow + 2}`).font = { italic: true, size: 9, bold: true, color: { argb: 'FFFF0000' } };
  sheet.getCell(`B${totRow + 2}`).alignment = { horizontal: 'center' };
  sheet.getCell(`B${totRow + 2}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFEB9C' }
  };

  sheet.mergeCells(`B${totRow + 3}:I${totRow + 3}`);
  sheet.getCell(`B${totRow + 3}`).value = '✅ EJEMPLO: Fila 1 ya tiene "PROD-001" seleccionado (producto de ejemplo). Descripción y precio se auto-completaron. Solo ingrese cantidad. Puede escribir el código directamente o seleccionar de la lista (▼).';
  sheet.getCell(`B${totRow + 3}`).font = { italic: true, size: 9 };
  sheet.getCell(`B${totRow + 3}`).alignment = { horizontal: 'center' };

  // Enlace de regreso
  sheet.getCell(`B${totRow + 5}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`B${totRow + 4}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: FACTURA CRÉDITO FISCAL
// ============================================================================
async function createFacturaCreditoFiscalSheet(workbook) {
  const sheet = workbook.addWorksheet('Factura_CreditoFiscal', {
    properties: { tabColor: { argb: 'FF5B9BD5' } }
  });

  sheet.columns = [
    { width: 5 },   // #
    { width: 15 },  // Código producto
    { width: 35 },  // Descripción
    { width: 10 },  // Cantidad
    { width: 15 },  // Precio
    { width: 15 },  // Subtotal
    { width: 15 },  // IVA
    { width: 15 },  // Retención
    { width: 15 }   // Total
  ];

  // Título
  sheet.mergeCells('B2:J2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = 'FACTURA - CRÉDITO FISCAL (CCF)';
  titleCell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF5B9BD5' }
  };
  sheet.getRow(2).height = 30;

  // Datos de la factura
  sheet.getCell('B4').value = 'No. Factura:';
  sheet.getCell('B4').font = { bold: true };
  sheet.getCell('C4').value = { formula: 'Correlativos!C8' };  // Lee el correlativo de Crédito Fiscal
  sheet.getCell('C4').numFmt = '0';
  sheet.getCell('C4').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFCCFFCC' }  // Verde claro
  };
  sheet.getCell('C4').border = { bottom: { style: 'thin' } };

  sheet.getCell('F4').value = 'Fecha:';
  sheet.getCell('F4').font = { bold: true };
  sheet.getCell('G4').value = new Date();
  sheet.getCell('G4').numFmt = 'dd/mm/yyyy';
  sheet.getCell('G4').border = { bottom: { style: 'thin' } };

  // Datos del cliente CON LISTA DESPLEGABLE
  sheet.getCell('B5').value = 'Código Cliente:';
  sheet.getCell('B5').font = { bold: true };
  sheet.getCell('C5').border = { bottom: { style: 'thin' } };

  // Lista desplegable de clientes
  sheet.getCell('C5').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['Clientes!$A$2:$A$1000'],
    showErrorMessage: true,
    errorTitle: 'Código inválido',
    error: 'Seleccione un código de cliente válido'
  };

  // VLOOKUP para nombre del cliente
  sheet.getCell('B6').value = 'Cliente:';
  sheet.getCell('B6').font = { bold: true };
  sheet.mergeCells('C6:I6');
  sheet.getCell('C6').value = {
    formula: 'IF(C5="","",IFERROR(VLOOKUP(C5,Clientes!$A$2:$N$1000,6,FALSE),""))'
  };
  sheet.getCell('C6').border = { bottom: { style: 'thin' } };

  // VLOOKUP para NIT
  sheet.getCell('B7').value = 'NIT:';
  sheet.getCell('B7').font = { bold: true };
  sheet.getCell('C7').value = {
    formula: 'IF(C5="","",IFERROR(VLOOKUP(C5,Clientes!$A$2:$N$1000,4,FALSE),""))'
  };
  sheet.getCell('C7').border = { bottom: { style: 'thin' } };

  // VLOOKUP para NRC
  sheet.getCell('E7').value = 'NRC:';
  sheet.getCell('E7').font = { bold: true };
  sheet.getCell('F7').value = {
    formula: 'IF(C5="","",IFERROR(VLOOKUP(C5,Clientes!$A$2:$N$1000,5,FALSE),""))'
  };
  sheet.getCell('F7').border = { bottom: { style: 'thin' } };

  // VLOOKUP para Gran Contribuyente
  sheet.getCell('H7').value = 'Gran Contribuyente:';
  sheet.getCell('H7').font = { bold: true, size: 9 };
  sheet.getCell('J7').value = {
    formula: 'IF(C5="","",IFERROR(VLOOKUP(C5,Clientes!$A$2:$N$1000,14,FALSE),""))'
  };
  sheet.getCell('J7').border = { bottom: { style: 'thin' } };

  // Tabla de productos
  const headerRow = 9;
  const headers = ['#', 'Código Producto', 'Descripción', 'Cantidad', 'Precio Unit.', 'Subtotal', 'IVA 13%', 'Retención 1%', 'Total'];
  let col = 'B';
  headers.forEach(header => {
    sheet.getCell(`${col}${headerRow}`).value = header;
    sheet.getCell(`${col}${headerRow}`).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    sheet.getCell(`${col}${headerRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF5B9BD5' }
    };
    sheet.getCell(`${col}${headerRow}`).alignment = { horizontal: 'center', wrapText: true };
    col = String.fromCharCode(col.charCodeAt(0) + 1);
  });
  sheet.getRow(headerRow).height = 30;

  // Filas para productos (10-29 = 20 líneas)
  for (let i = 10; i <= 29; i++) {
    sheet.getCell(`B${i}`).value = i - 9;
    sheet.getCell(`B${i}`).alignment = { horizontal: 'center' };

    // Lista desplegable de productos - columna C
    sheet.getCell(`C${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['Productos!$A$2:$A$1000']
    };

    // VLOOKUP para descripción - columna D
    sheet.getCell(`D${i}`).value = {
      formula: `IF(C${i}="","",IFERROR(VLOOKUP(C${i},Productos!$A$2:$E$1000,2,FALSE),""))`
    };

    // Columna E = Cantidad (INPUT MANUAL)
    sheet.getCell(`E${i}`).numFmt = '0';

    // VLOOKUP para precio - columna F
    sheet.getCell(`F${i}`).value = {
      formula: `IF(C${i}="","",IFERROR(VLOOKUP(C${i},Productos!$A$2:$E$1000,4,FALSE),""))`
    };
    sheet.getCell(`F${i}`).numFmt = '$#,##0.00';

    // Fórmula protegida para Subtotal - columna G (Cantidad*Precio = E*F)
    sheet.getCell(`G${i}`).value = {
      formula: `IF(AND(ISNUMBER(E${i}),ISNUMBER(F${i})),E${i}*F${i},"")`
    };
    sheet.getCell(`G${i}`).numFmt = '$#,##0.00';

    // Fórmula protegida para IVA - columna H
    sheet.getCell(`H${i}`).value = {
      formula: `IF(ISNUMBER(G${i}),G${i}*0.13,"")`
    };
    sheet.getCell(`H${i}`).numFmt = '$#,##0.00';

    // Retención 1% SOLO si es gran contribuyente (lee de J7) - columna I
    sheet.getCell(`I${i}`).value = {
      formula: `IF(AND(ISNUMBER(H${i}),UPPER($J$7)="SÍ"),H${i}*0.01,IF(ISNUMBER(H${i}),0,""))`
    };
    sheet.getCell(`I${i}`).numFmt = '$#,##0.00';

    // Total protegido - columna J
    sheet.getCell(`J${i}`).value = {
      formula: `IF(ISNUMBER(G${i}),G${i}+H${i}-I${i},"")`
    };
    sheet.getCell(`J${i}`).numFmt = '$#,##0.00';
  }

  // Ejemplo pre-llenado en la primera fila (fila 10)
  sheet.getCell('C5').value = 'CLI-001';  // Código de cliente de ejemplo
  sheet.getCell('C10').value = 'PROD-001';  // Código de producto
  sheet.getCell('E10').value = 3;  // Cantidad de ejemplo

  // Totales
  const totRow = 31;
  sheet.mergeCells(`B${totRow}:F${totRow}`);
  sheet.getCell(`B${totRow}`).value = 'TOTALES:';
  sheet.getCell(`B${totRow}`).font = { bold: true, size: 12 };
  sheet.getCell(`B${totRow}`).alignment = { horizontal: 'right' };

  ['G', 'H', 'I', 'J'].forEach(col => {
    sheet.getCell(`${col}${totRow}`).value = {
      formula: `SUMIF(${col}10:${col}29,"<>",${col}10:${col}29)`
    };
    sheet.getCell(`${col}${totRow}`).numFmt = '$#,##0.00';
    sheet.getCell(`${col}${totRow}`).font = { bold: true };
    sheet.getCell(`${col}${totRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9E2F3' }
    };
  });

  sheet.getCell(`J${totRow}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF5B9BD5' }
  };

  // Instrucciones
  sheet.mergeCells(`B${totRow + 2}:J${totRow + 2}`);
  sheet.getCell(`B${totRow + 2}`).value = '⚠️ CÓMO USAR: 1) Registre clientes en hoja "Clientes" (CLI-001, CLI-002...) y productos en "Productos" (PROD-001...). 2) Haga clic en "Código Cliente" o "Código Producto" y seleccione/escriba el código.';
  sheet.getCell(`B${totRow + 2}`).font = { italic: true, size: 9, bold: true, color: { argb: 'FFFF0000' } };
  sheet.getCell(`B${totRow + 2}`).alignment = { horizontal: 'center' };
  sheet.getCell(`B${totRow + 2}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFEB9C' }
  };

  sheet.mergeCells(`B${totRow + 3}:J${totRow + 3}`);
  sheet.getCell(`B${totRow + 3}`).value = '✅ EJEMPLO: Fila 1 ya tiene CLI-001 (cliente) y PROD-001 (producto). Datos se auto-completaron. Retención 1% se calcula si cliente es Gran Contribuyente. Solo ingrese cantidad.';
  sheet.getCell(`B${totRow + 3}`).font = { italic: true, size: 9 };
  sheet.getCell(`B${totRow + 3}`).alignment = { horizontal: 'center' };

  // Enlace de regreso
  sheet.getCell(`B${totRow + 5}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`B${totRow + 5}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: FACTURA SUJETO EXCLUIDO
// ============================================================================
async function createFacturaSujetoExcluidoSheet(workbook) {
  const sheet = workbook.addWorksheet('Factura_SujetoExcluido', {
    properties: { tabColor: { argb: 'FFC65911' } }
  });

  sheet.columns = [
    { width: 5 },   // #
    { width: 15 },  // Código producto
    { width: 35 },  // Descripción
    { width: 10 },  // Cantidad
    { width: 15 },  // Precio
    { width: 15 }   // Total
  ];

  // Título
  sheet.mergeCells('B2:G2');
  const titleCell = sheet.getCell('B2');
  titleCell.value = 'FACTURA - SUJETO EXCLUIDO (FSE)';
  titleCell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC65911' }
  };
  sheet.getRow(2).height = 30;

  // Datos de la factura
  sheet.getCell('B4').value = 'No. Factura:';
  sheet.getCell('B4').font = { bold: true };
  sheet.getCell('C4').value = { formula: 'Correlativos!C9' };  // Lee el correlativo de Sujeto Excluido
  sheet.getCell('C4').numFmt = '0';
  sheet.getCell('C4').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFCCFFCC' }  // Verde claro
  };
  sheet.getCell('C4').border = { bottom: { style: 'thin' } };

  sheet.getCell('E4').value = 'Fecha:';
  sheet.getCell('E4').font = { bold: true };
  sheet.getCell('F4').value = new Date();
  sheet.getCell('F4').numFmt = 'dd/mm/yyyy';
  sheet.getCell('F4').border = { bottom: { style: 'thin' } };

  // Campo de cliente con dropdown opcional
  sheet.getCell('B5').value = 'Código Cliente (opcional):';
  sheet.getCell('B5').font = { bold: true, size: 10 };
  sheet.getCell('C5').border = { bottom: { style: 'thin' } };

  // Lista desplegable de clientes (opcional)
  sheet.getCell('C5').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['Clientes!$A$2:$A$1000']
  };

  // VLOOKUP para nombre del cliente
  sheet.getCell('B6').value = 'Cliente:';
  sheet.getCell('B6').font = { bold: true };
  sheet.mergeCells('C6:G6');
  sheet.getCell('C6').value = {
    formula: 'IF(C5="","",IFERROR(VLOOKUP(C5,Clientes!$A$2:$N$1000,6,FALSE),""))'
  };
  sheet.getCell('C6').border = { bottom: { style: 'thin' } };

  // Documento
  sheet.getCell('B7').value = 'Documento:';
  sheet.getCell('B7').font = { bold: true };
  sheet.getCell('C7').value = {
    formula: 'IF(C5="","",IFERROR(VLOOKUP(C5,Clientes!$A$2:$N$1000,3,FALSE),""))'
  };
  sheet.getCell('C7').border = { bottom: { style: 'thin' } };

  // Tabla de productos
  const headerRow = 9;
  const headers = ['#', 'Código Producto', 'Descripción', 'Cantidad', 'Precio Unit.', 'Total'];
  let col = 'B';
  headers.forEach(header => {
    sheet.getCell(`${col}${headerRow}`).value = header;
    sheet.getCell(`${col}${headerRow}`).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getCell(`${col}${headerRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFC65911' }
    };
    sheet.getCell(`${col}${headerRow}`).alignment = { horizontal: 'center' };
    col = String.fromCharCode(col.charCodeAt(0) + 1);
  });

  // Filas para productos
  for (let i = 10; i <= 29; i++) {
    sheet.getCell(`B${i}`).value = i - 9;
    sheet.getCell(`B${i}`).alignment = { horizontal: 'center' };

    // Lista desplegable de productos - columna C
    sheet.getCell(`C${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['Productos!$A$2:$A$1000']
    };

    // VLOOKUP para descripción - columna D
    sheet.getCell(`D${i}`).value = {
      formula: `IF(C${i}="","",IFERROR(VLOOKUP(C${i},Productos!$A$2:$E$1000,2,FALSE),""))`
    };

    // Columna E = Cantidad (INPUT MANUAL)
    sheet.getCell(`E${i}`).numFmt = '0';

    // VLOOKUP para precio - columna F
    sheet.getCell(`F${i}`).value = {
      formula: `IF(C${i}="","",IFERROR(VLOOKUP(C${i},Productos!$A$2:$E$1000,4,FALSE),""))`
    };
    sheet.getCell(`F${i}`).numFmt = '$#,##0.00';

    // Total protegido - columna G (Cantidad * Precio = E * F)
    sheet.getCell(`G${i}`).value = {
      formula: `IF(AND(ISNUMBER(E${i}),ISNUMBER(F${i})),E${i}*F${i},"")`
    };
    sheet.getCell(`G${i}`).numFmt = '$#,##0.00';
  }

  // Ejemplo pre-llenado en la primera fila (fila 10)
  sheet.getCell('C10').value = 'PROD-003';  // Producto exento de ejemplo
  sheet.getCell('E10').value = 5;  // Cantidad de ejemplo

  // Totales
  const totRow = 31;
  sheet.mergeCells(`B${totRow}:F${totRow}`);
  sheet.getCell(`B${totRow}`).value = 'TOTAL A PAGAR:';
  sheet.getCell(`B${totRow}`).font = { bold: true, size: 12 };
  sheet.getCell(`B${totRow}`).alignment = { horizontal: 'right' };

  sheet.getCell(`G${totRow}`).value = { formula: `SUMIF(G10:G29,"<>",G10:G29)` };
  sheet.getCell(`G${totRow}`).numFmt = '$#,##0.00';
  sheet.getCell(`G${totRow}`).font = { bold: true, size: 12 };
  sheet.getCell(`G${totRow}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC65911' }
  };

  // Nota
  sheet.mergeCells(`B${totRow + 2}:G${totRow + 2}`);
  sheet.getCell(`B${totRow + 2}`).value = '⚠️ CÓMO USAR: 1) Registre clientes en "Clientes" (CLI-001...) y productos en "Productos" (PROD-001...). 2) Haga clic en "Código Cliente" o "Código Producto" y escriba/seleccione el código.';
  sheet.getCell(`B${totRow + 2}`).font = { italic: true, size: 9, bold: true, color: { argb: 'FFFF0000' } };
  sheet.getCell(`B${totRow + 2}`).alignment = { horizontal: 'center' };
  sheet.getCell(`B${totRow + 2}`).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFEB9C' }
  };

  sheet.mergeCells(`B${totRow + 3}:G${totRow + 3}`);
  sheet.getCell(`B${totRow + 3}`).value = '✅ EJEMPLO: Fila 1 ya tiene PROD-003 seleccionado. Descripción y precio se auto-completaron. Esta factura NO incluye IVA (sujeto excluido). Solo ingrese cantidad.';
  sheet.getCell(`B${totRow + 3}`).font = { italic: true, size: 9 };
  sheet.getCell(`B${totRow + 3}`).alignment = { horizontal: 'center' };

  // Enlace de regreso
  sheet.getCell(`B${totRow + 5}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`B${totRow + 5}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: NOTAS DE CRÉDITO
// ============================================================================
async function createNotasCreditoSheet(workbook) {
  const sheet = workbook.addWorksheet('Notas_Credito', {
    properties: { tabColor: { argb: 'FFE7E6E6' } }
  });

  sheet.columns = [
    { header: 'No. Nota', key: 'numero', width: 15 },
    { header: 'Fecha', key: 'fecha', width: 15 },
    { header: 'Tipo Documento', key: 'tipoDoc', width: 25 },
    { header: 'No. Factura Original', key: 'facturaOriginal', width: 20 },
    { header: 'Cliente', key: 'cliente', width: 35 },
    { header: 'Motivo', key: 'motivo', width: 40 },
    { header: 'Subtotal', key: 'subtotal', width: 15 },
    { header: 'IVA', key: 'iva', width: 15 },
    { header: 'Total', key: 'total', width: 15 }
  ];

  styleHeaderRow(sheet.getRow(1));

  // Validación tipo de documento
  for (let i = 2; i <= 1000; i++) {
    sheet.getCell(`C${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Consumidor Final,Crédito Fiscal,Sujeto Excluido"']
    };

    // Formato fecha
    sheet.getCell(`B${i}`).numFmt = 'dd/mm/yyyy';

    // Formato moneda
    sheet.getCell(`G${i}`).numFmt = '$#,##0.00';
    sheet.getCell(`H${i}`).numFmt = '$#,##0.00';
    sheet.getCell(`I${i}`).numFmt = '$#,##0.00';
  }

  // Enlace de regreso
  const lastRow = 3;
  sheet.getCell(`A${lastRow}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`A${lastRow}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: REGISTRO DE FACTURAS
// ============================================================================
async function createRegistroFacturasSheet(workbook) {
  const sheet = workbook.addWorksheet('Registro_Facturas', {
    properties: { tabColor: { argb: 'FF8EAADB' } }
  });

  sheet.columns = [
    { header: 'No. Factura', key: 'numero', width: 18 },
    { header: 'Fecha', key: 'fecha', width: 15 },
    { header: 'Tipo', key: 'tipo', width: 25 },
    { header: 'Cliente', key: 'cliente', width: 35 },
    { header: 'NIT/DUI', key: 'documento', width: 20 },
    { header: 'Subtotal', key: 'subtotal', width: 15 },
    { header: 'IVA', key: 'iva', width: 15 },
    { header: 'Retención 1%', key: 'retencion', width: 15 },
    { header: 'Total', key: 'total', width: 15 },
    { header: 'Estado', key: 'estado', width: 15 }
  ];

  styleHeaderRow(sheet.getRow(1));

  // Validaciones
  for (let i = 2; i <= 5000; i++) {
    // Tipo de factura
    sheet.getCell(`C${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Consumidor Final,Crédito Fiscal,Sujeto Excluido"']
    };

    // Estado
    sheet.getCell(`J${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Vigente,Anulada"']
    };

    // Formato fecha
    sheet.getCell(`B${i}`).numFmt = 'dd/mm/yyyy';

    // Formato moneda
    ['F', 'G', 'H', 'I'].forEach(col => {
      sheet.getCell(`${col}${i}`).numFmt = '$#,##0.00';
    });
  }

  // Sección de totales
  const totalRow = 2;
  sheet.getCell(`L${totalRow}`).value = 'RESUMEN';
  sheet.getCell(`L${totalRow}`).font = { bold: true, size: 12 };

  sheet.getCell(`L${totalRow + 1}`).value = 'Total Ventas:';
  sheet.getCell(`L${totalRow + 1}`).font = { bold: true };
  sheet.getCell(`M${totalRow + 1}`).value = { formula: 'SUMIF(I:I,"<>",I:I)' };
  sheet.getCell(`M${totalRow + 1}`).numFmt = '$#,##0.00';
  sheet.getCell(`M${totalRow + 1}`).font = { bold: true };

  sheet.getCell(`L${totalRow + 2}`).value = 'Total IVA:';
  sheet.getCell(`L${totalRow + 2}`).font = { bold: true };
  sheet.getCell(`M${totalRow + 2}`).value = { formula: 'SUMIF(G:G,"<>",G:G)' };
  sheet.getCell(`M${totalRow + 2}`).numFmt = '$#,##0.00';

  sheet.getCell(`L${totalRow + 3}`).value = 'Total Retención:';
  sheet.getCell(`L${totalRow + 3}`).font = { bold: true };
  sheet.getCell(`M${totalRow + 3}`).value = { formula: 'SUMIF(H:H,"<>",H:H)' };
  sheet.getCell(`M${totalRow + 3}`).numFmt = '$#,##0.00';

  // Enlace de regreso
  sheet.getCell(`A${totalRow + 6}`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`A${totalRow + 6}`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// HOJA: FACTURAS ANULADAS
// ============================================================================
async function createFacturasAnuladasSheet(workbook) {
  const sheet = workbook.addWorksheet('Facturas_Anuladas', {
    properties: { tabColor: { argb: 'FFFF0000' } }
  });

  sheet.columns = [
    { header: 'No. Factura', key: 'numero', width: 18 },
    { header: 'Fecha Emisión', key: 'fechaEmision', width: 15 },
    { header: 'Fecha Anulación', key: 'fechaAnulacion', width: 18 },
    { header: 'Tipo', key: 'tipo', width: 25 },
    { header: 'Cliente', key: 'cliente', width: 35 },
    { header: 'Monto Original', key: 'monto', width: 18 },
    { header: 'Motivo Anulación', key: 'motivo', width: 45 },
    { header: 'Anulado Por', key: 'anuladoPor', width: 25 }
  ];

  styleHeaderRow(sheet.getRow(1));

  // Formato rojo para el header
  sheet.getRow(1).eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF0000' }
    };
  });

  // Validaciones y formatos
  for (let i = 2; i <= 1000; i++) {
    sheet.getCell(`D${i}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Consumidor Final,Crédito Fiscal,Sujeto Excluido"']
    };

    sheet.getCell(`B${i}`).numFmt = 'dd/mm/yyyy';
    sheet.getCell(`C${i}`).numFmt = 'dd/mm/yyyy';
    sheet.getCell(`F${i}`).numFmt = '$#,##0.00';
  }

  // Enlace de regreso
  sheet.getCell(`A3`).value = {
    text: '← Volver al Dashboard',
    hyperlink: '#Dashboard!A1'
  };
  sheet.getCell(`A3`).font = { color: { argb: 'FF0563C1' }, underline: true };
}

// ============================================================================
// FUNCIONES AUXILIARES
// ============================================================================
function styleHeaderRow(row) {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' }
  };
  row.alignment = { vertical: 'middle', horizontal: 'center' };
  row.height = 25;

  row.eachCell(cell => {
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
    };
  });
}
