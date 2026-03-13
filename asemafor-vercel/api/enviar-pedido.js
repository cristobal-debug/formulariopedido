const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Método no permitido' });

  try {
    const { cliente, pedido, fecha } = req.body;
    const partes = cliente.split(' | ');
    const solicitante = partes[0] || cliente;
    const proyecto = partes[1] || '';

    // ── Generar Excel ──────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Pedido');

    const ahora = new Date();
    const mes = ahora.toLocaleString('es-CL', { month: 'long', year: 'numeric' }).toUpperCase();

    // Título
    sheet.mergeCells('A1:E1');
    const titulo = sheet.getCell('A1');
    titulo.value = `PEDIDO ${mes}`;
    titulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2e7d32' } };
    titulo.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 14 };
    titulo.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 30;

    // Info cliente
    sheet.mergeCells('A2:E2');
    const info = sheet.getCell('A2');
    info.value = `Solicitante: ${solicitante}   |   Proyecto: ${proyecto}   |   Fecha: ${fecha}`;
    info.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf5f5f5' } };
    info.font = { size: 10 };
    info.alignment = { horizontal: 'left', vertical: 'middle' };
    sheet.getRow(2).height = 20;

    sheet.getRow(3).height = 8;

    // Encabezados
    const headerRow = sheet.getRow(4);
    ['N°', 'SKU', 'Producto', 'Cantidad', 'Unidad'].forEach((col, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = col;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1b5e20' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });
    headerRow.height = 22;

    // Filas
    pedido.forEach((item, idx) => {
      const row = sheet.getRow(5 + idx);
      const bgColor = idx % 2 === 0 ? 'FFFFFFFF' : 'FFf1f8e9';
      [idx + 1, item.sku, item.nombre, item.cantidad, item.unidad].forEach((val, i) => {
        const cell = row.getCell(i + 1);
        cell.value = val;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.alignment = { vertical: 'middle', horizontal: i === 2 ? 'left' : 'center' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFcccccc' } },
          bottom: { style: 'thin', color: { argb: 'FFcccccc' } },
          left: { style: 'thin', color: { argb: 'FFcccccc' } },
          right: { style: 'thin', color: { argb: 'FFcccccc' } }
        };
      });
      row.height = 18;
    });

    sheet.getColumn(1).width = 6;
    sheet.getColumn(2).width = 18;
    sheet.getColumn(3).width = 38;
    sheet.getColumn(4).width = 10;
    sheet.getColumn(5).width = 10;

    const buffer = await workbook.xlsx.writeBuffer();
    const attachment = Buffer.from(buffer).toString('base64');

    // ── Enviar correo ──────────────────────────────────────────
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: 'asemaforutfsm@gmail.com',
        pass: 'ftypbkrikmvoaflv'
      }
    });

    const asunto = `Pedido ${mes} — ${solicitante} | ${proyecto}`;

    let tablaHtml = `
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
        <tr style="background:#2e7d32;color:white;">
          <th>N°</th><th>SKU</th><th>Producto</th><th>Cantidad</th><th>Unidad</th>
        </tr>`;
    pedido.forEach((item, idx) => {
      const bg = idx % 2 === 0 ? '#ffffff' : '#f1f8e9';
      tablaHtml += `<tr style="background:${bg}">
        <td align="center">${idx + 1}</td>
        <td>${item.sku}</td>
        <td>${item.nombre}</td>
        <td align="center">${item.cantidad}</td>
        <td align="center">${item.unidad}</td>
      </tr>`;
    });
    tablaHtml += '</table>';

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:700px;">
        <h2 style="color:#2e7d32;">Nuevo Pedido — Asemafor</h2>
        <p><b>Solicitante:</b> ${solicitante}<br>
           <b>Proyecto:</b> ${proyecto}<br>
           <b>Fecha:</b> ${fecha}</p>
        ${tablaHtml}
        <p style="color:#888;font-size:11px;margin-top:20px;">Se adjunta archivo Excel con el detalle del pedido.</p>
      </div>`;

    await transporter.sendMail({
      from: `"Formulario Asemafor" <asemaforutfsm@gmail.com>`,
      to: 'kcontreras@asemafor.cl, aleal@asemafor.cl, dguerrero@asemafor.cl, cristobal@jeldrez.com',
      subject: asunto,
      html: htmlBody,
      attachments: [{
        filename: `Pedido_${solicitante.replace(/\s+/g, '_')}_${mes}.xlsx`,
        content: attachment,
        encoding: 'base64',
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }]
    });

    return res.status(200).json({ ok: true });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok: false, error: err.message });
  }
}
