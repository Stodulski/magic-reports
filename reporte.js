import axios from 'axios';
import nodemailer from 'nodemailer';
import ExcelJS from 'exceljs';

async function getAllOrders() {
  const now = new Date();
  const day = now.getDay();
  // Calcula cuántos días restar para llegar al último sábado (6)
  const daysToLastSaturday = (day + 1) % 7;

  // Último sábado a las 00:00 (hora local)
  const start = new Date(now);
  start.setDate(now.getDate() - daysToLastSaturday);
  start.setHours(0, 0, 0, 0);
  const isoAfter = start.toISOString();

  // Momento actual
  const isoBefore = now.toISOString();

  const paramsBase = {
    status: ["completed", "processing", "pending"],
    after: isoAfter,   // último sábado 00:00
    before: isoBefore, // ahora
    per_page: 100,
  };

  let allOrders = [];
  let page = 1;
  let totalPages = 1;

  try {
    do {
      const response = await axios.get("https://magicstore.com.ar/wp-json/wc/v3/orders", {
        params: { ...paramsBase, page },
        auth: {
          username: process.env.WC_KEY,
          password: process.env.WC_SECRET,
        },
      });

      const orders = response.data;
      totalPages = parseInt(response.headers['x-wp-totalpages'], 10);
      allOrders = allOrders.concat(orders);
      page++;
    } while (page <= totalPages);
    const totalPedidos = allOrders.length;
    const totalProductos = allOrders.reduce((sum, order) => {
      return sum + order.line_items.reduce((prodSum, item) => prodSum + item.quantity, 0);
    }, 0);

    const promedio = totalProductos / totalPedidos;
    const ingresoTotal = allOrders.reduce((sum, order) => sum + parseFloat(order.total), 0);
    const ticketPromedio = totalPedidos > 0 ? (ingresoTotal / totalPedidos).toFixed(2) : 0;

    return [allOrders, promedio, ticketPromedio, ingresoTotal, totalPedidos];
  } catch (error) {
    console.error("❌ Error obteniendo pedidos:", error.response?.data || error.message);
    return [];
  }
}

async function getProductAttributes(productId) {
  try {
    const response = await axios.get(`https://magicstore.com.ar/wp-json/wc/v3/products/${productId}`, {
      auth: {
        username: process.env.WC_KEY,
        password: process.env.WC_SECRET,
      },
    });

    const attributes = response.data.attributes;
    const personajesAttr = attributes.find((attr) => attr.name.toLowerCase() === "personajes");
    return personajesAttr ? personajesAttr.options : [];
  } catch (error) {
    console.error("❌ Error obteniendo atributos del producto:", error.response?.data || error.message);
    return [];
  }
}

async function countSalesByCharacter() {
  try {

  } catch (error) {

  }
  const orders = await getAllOrders();
  const characterSales = {};
  const productosPorSKU = {};
  const orderData = [];

  if (orders[0].length === 0) {
    console.log("❌ No se encontraron pedidos para el rango de fechas.");
    return;
  }

  for (const order of orders[0]) {
    const fuente = order.meta_data.find(
      (item) => item.key === '_wc_order_attribution_utm_source'
    );
    const orderId = order.number
    const localidad = order?.shipping.city || "Sin localidad"
    const orderResult = {
      orderId,
      fuente: fuente?.value,
      localidad,
      total: order.total
    }
    orderData.push(orderResult)

    for (const item of order.line_items) {
      let personajes = [];

      if (item.product_id) {
        personajes = await getProductAttributes(item.product_id);
      }

      personajes.forEach((character) => {
        characterSales[character] = (characterSales[character] || 0) + item.quantity;
      });

      const sku = item.sku || ''; // Si no hay SKU, queda vacío
      const nombre = item.name;
      const cantidad = item.quantity;

      if (!productosPorSKU[sku]) {
        productosPorSKU[sku] = {
          sku: sku,
          nombre: nombre,
          cantidad: 0,
        };
      }

      productosPorSKU[sku].cantidad += cantidad;

    }
  }


  return [characterSales, productosPorSKU, orderData, orders[1], orders[2], orders[3], orders[4]]
}

export default async function generarYEnviarReporteExcel() {
  try {
    const result = await countSalesByCharacter()
    console.log(result[1])

    const workbook = new ExcelJS.Workbook();

    const hoja1 = workbook.addWorksheet('ventas_personajes');
    hoja1.columns = [
      { header: 'Nombre', key: 'nombre', width: 40 },
      { header: 'Cantidad Vendida', key: 'cantidad', width: 20 },
    ];
    for (const [nombre, cantidad] of Object.entries(result[0])) {
      hoja1.addRow([nombre, cantidad]);
    }

    const hoja2 = workbook.addWorksheet('ventas_productos');
    hoja2.columns = [
      { header: 'SKU', key: 'sku', width: 20 },
      { header: 'Nombre', key: 'nombre', width: 50 },
      { header: 'Cantidad Vendida', key: 'cantidad', width: 20 },
    ];


    for (const producto of Object.values(result[1])) {
      hoja2.addRow({
        sku: producto.sku,
        nombre: producto.nombre,
        cantidad: producto.cantidad,
      });
    }


    const hoja3 = workbook.addWorksheet('pedidos');
    hoja3.columns = [
      { header: 'ID', key: 'id', width: 20 },
      { header: 'Localidad', key: 'localidad', width: 100 },
      { header: 'Valor total', key: 'total', width: 50 },
      { header: 'Fuente', key: 'fuente', width: 40 },
    ];

    for (const pedido of Object.values(result[2])) {
      hoja3.addRow({
        id: pedido.orderId,
        fuente: pedido.fuente,
        localidad: pedido.localidad,
        total: pedido.total
      });
    }

    const hoja4 = workbook.addWorksheet('General');
    hoja4.columns = [
      { header: 'Valor Total', key: 'total', width: 30 },
      { header: 'Total Pedidos', key: 'pedidos', width: 30 },
      { header: 'Productos Promedio', key: 'productos', width: 30 },
      { header: 'Ticket Promedio', key: 'ticket', width: 30 },
    ];

    hoja4.addRow({
      total: result[5],
      pedidos: result[6],
      productos: result[3],
      ticket: result[4]
    });




    const buffer = await workbook.xlsx.writeBuffer();

    // 📧 SMTP Gmail explícito
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: process.env.EMAIL_FROM,
        pass: process.env.EMAIL_PASS
      }
    });

    const now = new Date();
    const day = now.getDay();
    // Días a retroceder para llegar al último sábado (6)
    const daysToLastSaturday = (day + 1) % 7;

    // Último sábado a las 00:00
    const start = new Date(now);
    start.setDate(now.getDate() - daysToLastSaturday);
    start.setHours(0, 0, 0, 0);

    // Viernes siguiente a las 23:59:59.999
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);

    // Formateo DD/MM/YYYY en locale argentino
    const formattedStart = start.toLocaleDateString('es-AR');  // ej. "19/07/2025"
    const formattedEnd = end.toLocaleDateString('es-AR');    // ej. "25/07/2025"
console.log(formattedEnd)
    // Armo y envío el mail
    await transporter.sendMail({
      from: `"Magic Store" <${process.env.EMAIL_FROM}>`,
      to: process.env.EMAIL_TO,
      subject: `Reporte semanal de ventas - Magic Store (${formattedStart} - ${formattedEnd})`,
      text: `¡Hola! Adjunto el reporte semanal en formato Excel.\nPeríodo: ${formattedStart} - ${formattedEnd}`,
      attachments: [{
        filename: `${formattedStart.replace(/\//g, '-')}_${formattedEnd.replace(/\//g, '-')}.xlsx`,
        content: buffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }]
    });
    console.log('✅ Reporte enviado con éxito');
  } catch (error) {
    console.error('❌ Error al generar o enviar el Excel:', error.message);
  }
}


