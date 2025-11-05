// Catálogos base de El Salvador para sistema de facturación

export const departamentos = [
  { codigo: '01', nombre: 'Ahuachapán' },
  { codigo: '02', nombre: 'Santa Ana' },
  { codigo: '03', nombre: 'Sonsonate' },
  { codigo: '04', nombre: 'Chalatenango' },
  { codigo: '05', nombre: 'La Libertad' },
  { codigo: '06', nombre: 'San Salvador' },
  { codigo: '07', nombre: 'Cuscatlán' },
  { codigo: '08', nombre: 'La Paz' },
  { codigo: '09', nombre: 'Cabañas' },
  { codigo: '10', nombre: 'San Vicente' },
  { codigo: '11', nombre: 'Usulután' },
  { codigo: '12', nombre: 'San Miguel' },
  { codigo: '13', nombre: 'Morazán' },
  { codigo: '14', nombre: 'La Unión' }
];

export const municipios = [
  // Ahuachapán
  { codigo: '01-01', nombre: 'Ahuachapán', departamento: '01' },
  { codigo: '01-02', nombre: 'Apaneca', departamento: '01' },
  { codigo: '01-03', nombre: 'Atiquizaya', departamento: '01' },
  { codigo: '01-04', nombre: 'Concepción de Ataco', departamento: '01' },

  // Santa Ana
  { codigo: '02-01', nombre: 'Santa Ana', departamento: '02' },
  { codigo: '02-02', nombre: 'Candelaria de la Frontera', departamento: '02' },
  { codigo: '02-03', nombre: 'Chalchuapa', departamento: '02' },
  { codigo: '02-04', nombre: 'Coatepeque', departamento: '02' },
  { codigo: '02-05', nombre: 'Metapán', departamento: '02' },

  // Sonsonate
  { codigo: '03-01', nombre: 'Sonsonate', departamento: '03' },
  { codigo: '03-02', nombre: 'Acajutla', departamento: '03' },
  { codigo: '03-03', nombre: 'Armenia', departamento: '03' },
  { codigo: '03-04', nombre: 'Izalco', departamento: '03' },
  { codigo: '03-05', nombre: 'Nahuizalco', departamento: '03' },

  // La Libertad
  { codigo: '05-01', nombre: 'Santa Tecla', departamento: '05' },
  { codigo: '05-02', nombre: 'Antiguo Cuscatlán', departamento: '05' },
  { codigo: '05-03', nombre: 'Colón', departamento: '05' },
  { codigo: '05-04', nombre: 'La Libertad', departamento: '05' },
  { codigo: '05-05', nombre: 'San Juan Opico', departamento: '05' },

  // San Salvador
  { codigo: '06-01', nombre: 'San Salvador', departamento: '06' },
  { codigo: '06-02', nombre: 'Apopa', departamento: '06' },
  { codigo: '06-03', nombre: 'Ayutuxtepeque', departamento: '06' },
  { codigo: '06-04', nombre: 'Cuscatancingo', departamento: '06' },
  { codigo: '06-05', nombre: 'Delgado', departamento: '06' },
  { codigo: '06-06', nombre: 'Ilopango', departamento: '06' },
  { codigo: '06-07', nombre: 'Mejicanos', departamento: '06' },
  { codigo: '06-08', nombre: 'Nejapa', departamento: '06' },
  { codigo: '06-09', nombre: 'San Marcos', departamento: '06' },
  { codigo: '06-10', nombre: 'San Martín', departamento: '06' },
  { codigo: '06-11', nombre: 'Soyapango', departamento: '06' },
  { codigo: '06-12', nombre: 'Tonacatepeque', departamento: '06' },

  // San Miguel
  { codigo: '12-01', nombre: 'San Miguel', departamento: '12' },
  { codigo: '12-02', nombre: 'Chinameca', departamento: '12' },
  { codigo: '12-03', nombre: 'Ciudad Barrios', departamento: '12' },
  { codigo: '12-04', nombre: 'Moncagua', departamento: '12' }
];

export const actividadesEconomicas = [
  { codigo: '10005', descripcion: 'Productos Farmacéuticos' },
  { codigo: '10410', descripcion: 'Comercio al por Mayor' },
  { codigo: '46900', descripcion: 'Comercio al por Mayor No Especializado' },
  { codigo: '47110', descripcion: 'Comercio al por Menor en Almacenes No Especializados' },
  { codigo: '47190', descripcion: 'Otro tipo de Comercio al por Menor en Almacenes No Especializados' },
  { codigo: '47300', descripcion: 'Comercio al por Menor de Combustible' },
  { codigo: '47410', descripcion: 'Comercio al por Menor de Computadoras y Equipo Periférico' },
  { codigo: '47510', descripcion: 'Comercio al por Menor de Productos Textiles' },
  { codigo: '47610', descripcion: 'Comercio al por Menor de Libros' },
  { codigo: '47730', descripcion: 'Comercio al por Menor de Productos Farmacéuticos' },
  { codigo: '52100', descripcion: 'Almacenamiento y Depósito' },
  { codigo: '55101', descripcion: 'Hoteles' },
  { codigo: '56101', descripcion: 'Restaurantes y Servicios de Comida' },
  { codigo: '56210', descripcion: 'Suministro de Comidas por Encargo' },
  { codigo: '58190', descripcion: 'Otras Actividades de Edición' },
  { codigo: '62010', descripcion: 'Actividades de Programación Informática' },
  { codigo: '62020', descripcion: 'Actividades de Consultoría Informática' },
  { codigo: '63110', descripcion: 'Procesamiento de Datos, Hospedaje y Actividades Conexas' },
  { codigo: '68100', descripcion: 'Actividades Inmobiliarias' },
  { codigo: '69200', descripcion: 'Actividades de Contabilidad, Teneduría de Libros y Auditoría' },
  { codigo: '70200', descripcion: 'Actividades de Consultoría de Gestión' },
  { codigo: '71100', descripcion: 'Actividades de Arquitectura e Ingeniería' },
  { codigo: '73110', descripcion: 'Agencias de Publicidad' },
  { codigo: '77110', descripcion: 'Alquiler de Automóviles' },
  { codigo: '78100', descripcion: 'Actividades de Agencias de Empleo' },
  { codigo: '80100', descripcion: 'Actividades de Seguridad Privada' },
  { codigo: '82110', descripcion: 'Servicios Administrativos de Oficina' },
  { codigo: '85410', descripcion: 'Enseñanza de Educación Superior' },
  { codigo: '86210', descripcion: 'Actividades de Práctica Médica' },
  { codigo: '86220', descripcion: 'Actividades de Práctica Odontológica' },
  { codigo: '93110', descripcion: 'Gestión de Instalaciones Deportivas' },
  { codigo: '95110', descripcion: 'Reparación de Computadoras y Equipo Periférico' },
  { codigo: '96020', descripcion: 'Peluquería y Otros Tratamientos de Belleza' }
];

export const tiposDocumento = [
  { codigo: '13', nombre: 'DUI - Documento Único de Identidad' },
  { codigo: '36', nombre: 'NIT - Número de Identificación Tributaria' },
  { codigo: '37', nombre: 'Otro' },
  { codigo: '02', nombre: 'Carnet de Residente' },
  { codigo: '03', nombre: 'Pasaporte' }
];

export const condicionesPago = [
  { codigo: '1', nombre: 'Contado' },
  { codigo: '2', nombre: 'Crédito' },
  { codigo: '3', nombre: 'Otro' }
];

export const tiposProducto = [
  { codigo: 'G', nombre: 'Gravado (13% IVA)' },
  { codigo: 'E', nombre: 'Exento (0% IVA)' },
  { codigo: 'N', nombre: 'No Sujeto (Sin IVA)' }
];
