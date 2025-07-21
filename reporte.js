import axios from 'axios';
import nodemailer from 'nodemailer';
import ExcelJS from 'exceljs';

export default async function generarYEnviarReporteExcel() {
  try {
    const now = new Date();
    const day = now.getDay();
    const daysToLastSaturday = (day + 1) % 7;
    const start = new Date(now);
    start.setDate(now.getDate() - daysToLastSaturday);
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(end.getDate() + 6);
    end.setHours(23, 59, 59, 999);

    const isoStart = start.toISOString();
    const isoEnd = end.toISOString();

    const res = await axios.get('https://tu-tienda.com/wp-json/wc/v3/orders', {
      auth: {
        username: process.env.WC_KEY,
        password: process.env.WC_SECRET
      },
      params: {
        status: ['completed', 'processing'],
        after: isoStart,
        before: isoEnd,
        per_page: 100
      }
    });

    const orders = res.data;
    const porPersonaje = {};
    const porProducto = [];
    const pedidos = [];
    let totalPedidos = 0;
    let totalMonto = 0;

    for (const order of orders) {
      totalPedidos++;
      totalMonto += parseFloat(order.total);

      const fuente = order.meta_data?.find(m => m.key === '_ck_attribution_data')?.value?.source ?? 'No disponible';

      pedidos.push([
        order.id,
        parseFloat(order.total).toLocaleString('es-AR', { minimumFractionDigits: 2 }),
        order.billing.city,
        fuente
      ]);

      for (const item of order.line_items) {
        const name = item.name;
        const sku = item.sku || '';
        const qty = item.quantity;

        porProducto.push([name, sku, qty]);

        const personaje = item.meta_data?.find(m => m.key === 'pa_personajes')?.value;
        if (personaje) {
          porPersonaje[personaje] = (porPersonaje[personaje] || 0) + qty;
        }
      }
    }

    const workbook = new ExcelJS.Workbook();

    const hoja1 = workbook.addWorksheet('ventas_personajes');
    hoja1.addRow(['Personaje', 'Cantidad vendida']);
    for (const [nombre, cantidad] of Object.entries(porPersonaje)) {
      hoja1.addRow([nombre, cantidad]);
    }

    const hoja2 = workbook.addWorksheet('ventas_productos');
    hoja2.addRow(['Producto', 'SKU', 'Cantidad vendida']);
    porProducto.forEach(p => hoja2.addRow(p));

    const hoja3 = workbook.addWorksheet('pedidos');
    hoja3.addRow(['Pedido ID', 'Valor', 'Localidad', 'Fuente']);
    pedidos.forEach(p => hoja3.addRow(p));
    hoja3.addRow([]);
    hoja3.addRow(['Total de pedidos', totalPedidos]);
    hoja3.addRow(['Monto total vendido', totalMonto.toLocaleString('es-AR', { minimumFractionDigits: 2 })]);

    const buffer = await workbook.xlsx.writeBuffer();

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_FROM,
        pass: process.env.EMAIL_PASS
      }
    });

    await transporter.sendMail({
      from: `"Magic Store" <${process.env.EMAIL_FROM}>`,
      to: process.env.EMAIL_TO,
      subject: 'Reporte semanal de ventas - Magic Store',
      text: 'Hola! Adjunto el reporte semanal en formato Excel.',
      attachments: [{
        filename: 'reporte_magicstore.xlsx',
        content: buffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }]
    });

    console.log('✅ Reporte enviado con éxito');
  } catch (err) {
    console.error('❌ Error al generar o enviar el Excel:', err.message);
  }
}