import express from 'express';
import { generateExcel } from './src/excelGenerator.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

// Ruta principal - Servir formulario
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Ruta para generar y descargar el Excel
app.post('/api/generar-excel', async (req, res) => {
  try {
    const emisorData = req.body;

    // Validaciones básicas
    if (!emisorData.razonSocial || !emisorData.nit) {
      return res.status(400).json({
        error: 'Datos incompletos',
        mensaje: 'La Razón Social y NIT son obligatorios'
      });
    }

    console.log('Generando Excel para:', emisorData.razonSocial);

    // Generar el workbook
    const workbook = await generateExcel(emisorData);

    // Configurar headers para descarga
    const fileName = `Facturacion_${emisorData.razonSocial.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.xlsx`;

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${fileName}"`
    );

    // Escribir el Excel directamente al response
    await workbook.xlsx.write(res);

    console.log('Excel generado exitosamente:', fileName);
    res.end();

  } catch (error) {
    console.error('Error al generar Excel:', error);
    res.status(500).json({
      error: 'Error del servidor',
      mensaje: 'No se pudo generar el archivo Excel',
      detalle: error.message
    });
  }
});

// Ruta de salud (health check)
app.get('/api/health', (req, res) => {
  res.json({
    status: 'OK',
    timestamp: new Date().toISOString(),
    service: 'FacturaExcel Generator'
  });
});

// Iniciar servidor
app.listen(PORT, () => {
  console.log('='.repeat(60));
  console.log('🚀 SISTEMA DE FACTURACIÓN ELECTRÓNICA - ACTIVO');
  console.log('='.repeat(60));
  console.log(`📡 Servidor corriendo en: http://localhost:${PORT}`);
  console.log(`📄 Abra su navegador en: http://localhost:${PORT}`);
  console.log('='.repeat(60));
  console.log('Presione Ctrl+C para detener el servidor');
  console.log('='.repeat(60));
});

// Manejo de errores no capturados
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  process.exit(1);
});
