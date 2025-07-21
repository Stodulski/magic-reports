import express from 'express';
import generarYEnviarReporteExcel from './reporte.js';

const app = express();
const PORT = process.env.PORT || 3000;

app.get('/api/ejecutar-reporte', async (req, res) => {
  try {
    res.status(200).send('✅ Reporte enviado correctamente por email');
    await generarYEnviarReporteExcel();
  } catch (error) {
    res.status(500).send('❌ Error al enviar el reporte: ' + error.message);
  }
});

app.get('/', (req, res) => {
  res.send('🧾 Magic Store Reporte Backend está activo');
});

app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
});
