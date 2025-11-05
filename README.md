# 📊 FacturaExcel - Sistema de Facturación Electrónica

Aplicación web para generar archivos Excel pre-configurados con hojas de facturación electrónica para El Salvador.

## 🌟 Características

### Hojas Incluidas en el Excel Generado:

1. **Dashboard** - Hoja principal con navegación a todas las demás hojas
2. **Emisor** - Datos del emisor (ya pre-cargados)
3. **Catálogos:**
   - Departamentos de El Salvador
   - Municipios
   - Actividades Económicas
4. **Clientes** - Gestión de clientes con campos para:
   - Gran contribuyente (para retención 1% IVA)
   - Datos tributarios (NIT, NRC)
   - Información de contacto
5. **Productos** - Catálogo de productos/servicios con:
   - Productos Gravados (13% IVA)
   - Productos Exentos (0% IVA)
   - Productos No Sujetos
6. **Hojas de Facturación:**
   - Factura Consumidor Final
   - Crédito Fiscal (CCF) con retención automática para grandes contribuyentes
   - Sujeto Excluido (FSE)
7. **Notas de Crédito** - Registro de notas de crédito
8. **Registro de Facturas** - Todas las facturas emitidas en un solo lugar
9. **Facturas Anuladas** - Control de facturas anuladas

### ✨ Funcionalidades Especiales:

- ✅ Fórmulas automáticas para cálculo de IVA
- ✅ Retención automática del 1% para grandes contribuyentes
- ✅ Validaciones de datos (listas desplegables)
- ✅ Formato profesional con colores diferenciados
- ✅ Navegación fácil entre hojas mediante hipervínculos
- ✅ Catálogos pre-cargados de El Salvador

## 🚀 Instalación

### Prerrequisitos

- Node.js (v14 o superior)
- npm o yarn

### Pasos de Instalación

1. Clonar el repositorio:
```bash
git clone <url-del-repositorio>
cd FacturaExcel
```

2. Instalar dependencias:
```bash
npm install
```

3. Iniciar el servidor:
```bash
npm start
```

4. Abrir navegador en:
```
http://localhost:3000
```

## 📖 Uso

1. **Completar el Formulario:**
   - Ingrese los datos del emisor (su empresa)
   - Complete los campos obligatorios marcados con *
   - Los campos opcionales ayudan a personalizar más el Excel

2. **Generar Excel:**
   - Haga clic en "Generar y Descargar Excel"
   - El archivo se descargará automáticamente
   - El archivo está listo para usar inmediatamente

3. **Usar el Excel:**
   - Abra el archivo descargado
   - Navegue desde el Dashboard
   - Complete los catálogos de clientes y productos
   - Empiece a facturar usando las hojas correspondientes

## 📁 Estructura del Proyecto

```
FacturaExcel/
├── public/
│   └── index.html          # Formulario web
├── src/
│   ├── catalogData.js      # Catálogos de El Salvador
│   └── excelGenerator.js   # Generador de hojas Excel
├── server.js               # Servidor Express
├── package.json
└── README.md
```

## 🛠️ Tecnologías Utilizadas

- **Backend:** Node.js + Express
- **Excel:** ExcelJS
- **Frontend:** HTML5 + CSS3 + JavaScript (Vanilla)

## 📊 Tipos de Productos

El sistema maneja tres tipos de productos según la legislación salvadoreña:

1. **Gravado (G):** 13% IVA
2. **Exento (E):** 0% IVA
3. **No Sujeto (N):** Sin IVA

## 👥 Grandes Contribuyentes

Cuando se factura a un gran contribuyente (Crédito Fiscal), el sistema:
- Calcula automáticamente la retención del 1% sobre el IVA
- Resta la retención del total a cobrar
- Muestra claramente el monto retenido

## 🔧 Configuración

### Cambiar Puerto del Servidor

Edite `server.js` o use variable de entorno:
```bash
PORT=4000 npm start
```

### Personalizar Catálogos

Los catálogos base se encuentran en `src/catalogData.js`. Puede modificarlos según necesite.

## 📝 API Endpoints

### POST `/api/generar-excel`
Genera y descarga el archivo Excel.

**Body (JSON):**
```json
{
  "razonSocial": "Empresa S.A. de C.V.",
  "nombreComercial": "MiEmpresa",
  "nit": "0614-123456-123-4",
  "nrc": "12345-6",
  "actividadEconomica": "Comercio al por Mayor",
  "direccion": "San Salvador",
  "departamento": "San Salvador",
  "municipio": "San Salvador",
  "telefono": "2222-2222",
  "correo": "info@empresa.com"
}
```

**Response:**
Archivo Excel (.xlsx) listo para descargar

### GET `/api/health`
Verifica el estado del servidor.

**Response:**
```json
{
  "status": "OK",
  "timestamp": "2025-01-01T00:00:00.000Z",
  "service": "FacturaExcel Generator"
}
```

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Cree una rama para su feature (`git checkout -b feature/NuevaCaracteristica`)
3. Commit sus cambios (`git commit -m 'Agregar nueva característica'`)
4. Push a la rama (`git push origin feature/NuevaCaracteristica`)
5. Abra un Pull Request

## 📄 Licencia

ISC

## 👨‍💻 Autor

Sistema desarrollado para facilitar la facturación electrónica en El Salvador.

## 📞 Soporte

Para reportar problemas o solicitar características, por favor abra un issue en el repositorio.

---

**Nota:** Este sistema genera archivos Excel para uso local. No transmite datos ni genera DTE (Documentos Tributarios Electrónicos) para el Ministerio de Hacienda. Es una herramienta de organización y control interno.
