function doGet() {
  return HtmlService.createHtmlOutputFromFile('orden')
    .setTitle('Órdenes — Merchi (Ejecutivos)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================
   CONFIGURACIÓN
   ============================ */

// Spreadsheet con clientes (formulario público)
const SHEET_ID_CLIENTES = '1o9q9cwgx4f1iV87yrkdI4a5JR8El_coZa24VsKWMQhs';
const TAB_CLIENTES      = 'Datos Cotizaciones';

// Spreadsheet de Órdenes (este formulario)
const SHEET_ID_ORDENES  = '1GUPTg4H4V2OIUGrX4LkUzA2mJgjdN7USlzbcPJLiheM';
const TAB_ORDENES       = 'Ordenes';

// Carpeta para guardar adjuntos (productos) y OC
const LOGOS_FOLDER_ID   = '12KE8T2uFFQFlXfjwtYV_yboIyP_nT8Jn'; // productos
const OC_FOLDER_ID      = '1VKxNC3X1AbZdJRJgQnustt6dN0t6YAof'; // OC

/* ===== Email ===== */
const RECIPIENTS_DEFAULT = ['ordenes.merchi@gmail.com', 'pablo@merchi.cl'];
const CC_DEFAULT         = ['sergio@merchi.cl'];
const SEND_TO_EXEC       = false;
const SEND_TO_CLIENT     = false;

const BRAND = {
  LOGO_URL   : 'https://merchi.cl/wp-content/uploads/2025/10/Logo-Merchi-Azul.png',
  COLOR      : '#0b53d0',
  COLOR_SOFT : '#f1f5ff',
  FOOTER     : 'Merchi - Regalos corporativos '
};


/* ============================
   CLIENTES (AUTOCOMPLETAR)
   ============================ */
function getClientesLite() {
  const sheet = SpreadsheetApp.openById(SHEET_ID_CLIENTES).getSheetByName(TAB_CLIENTES);
  if (!sheet) throw new Error('No existe la hoja "'+TAB_CLIENTES+'" en el archivo de CLIENTES.');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(h => (h||'').toString().trim());
  const idx = {
    id:      header.indexOf('ID_Registro'),
    empresa: header.indexOf('Empresa'),
    giro:    header.indexOf('Giro'),
    rut:     header.indexOf('RUT_CL')
  };

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id    = idx.id      >= 0 ? (row[idx.id]||'').toString().trim()      : '';
    const razon = idx.empresa >= 0 ? (row[idx.empresa]||'').toString().trim() : '';
    if (!id && !razon) continue;
    out.push({
      id: id,
      razon: razon,
      rut:  idx.rut  >= 0 ? (row[idx.rut]||'').toString().trim()  : '',
      giro: idx.giro >= 0 ? (row[idx.giro]||'').toString().trim() : ''
    });
  }
  return out;
}

function getClienteByIdOrRazon_(clienteId, razonSocial){
  const sh = SpreadsheetApp.openById(SHEET_ID_CLIENTES).getSheetByName(TAB_CLIENTES);
  if (!sh) return null;
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0].map(h => (h||'').toString().trim());
  const idx = {
    id:      header.indexOf('ID_Registro'),
    empresa: header.indexOf('Empresa'),
    correo:  header.indexOf('Correo')
  };
  for (let i=1;i<values.length;i++){
    const row = values[i];
    const id    = idx.id>=0 ? (row[idx.id]||'').toString().trim() : '';
    const emp   = idx.empresa>=0 ? (row[idx.empresa]||'').toString().trim() : '';
    const mail  = idx.correo>=0 ? (row[idx.correo]||'').toString().trim() : '';
    if (clienteId && id && id.toUpperCase() === (clienteId||'').toUpperCase()) return { id, empresa: emp, correo: mail };
    if (razonSocial && emp && emp.toLowerCase() === (razonSocial||'').toLowerCase()) return { id, empresa: emp, correo: mail };
  }
  return null;
}


/* ============================
   ÓRDENES
   ============================ */
/**
 * data = {
 *   cliente_id, razon_social, cot_num, observaciones,
 *   precio_final, costo_transporte,
 *   oc_num, oc_fecha, oc_url,
 *   pago_condicion, pago_detalle,
 *   transporte_modo, transporte_pref, transporte_detalle,
 *   items: [ { ... } ]
 * }
 */
function saveOrden(data) {
  try {
    if (!data || !Array.isArray(data.items) || data.items.length===0)
      return { ok:false, message:'Debes enviar al menos un producto.' };

    const ss = SpreadsheetApp.openById(SHEET_ID_ORDENES);
    const sheet = ss.getSheetByName(TAB_ORDENES) || ss.insertSheet(TAB_ORDENES);

    // Encabezado si vacío (agregamos Cot_Num y mantenemos columnas nuevas)
    if (sheet.getLastRow() === 0){
      sheet.appendRow([
        'Fecha','Orden_ID','Item_N','Cliente_ID','Razon_Social','Cot_Num',
        'Producto','Color','Cantidad','Metodo_Impresion','Tiempo_Produccion',
        'Fecha_Entrega','Observaciones',
        'Adjunto_URL',
        'Precio_Final_Orden','Costo_Transporte','Venta_Sin_Transporte',
        'OC_Num','OC_Fecha','OC_URL',
        'Pago_Condicion','Pago_Detalle',
        'Transporte_Modo','Transporte_Preferencia','Transporte_Detalle',
        'Adjuntos_JSON',
        'Colores_Serigrafia_Tampografia',
        'Grabado_Nombres',
        'Grabado_Archivo_URL',
        'Notas_Producto'
      ]);
      sheet.getRange(1,1,1,30).setFontWeight('bold');
      sheet.getRange(1,1,1,30).setWrap(true);
    }

    const clean = v => (v||'').toString().trim();

    // Fecha sin hora
    const now = new Date();
    const fechaSolo = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    const ordenId   = generarIdCorto_();
    const clienteId = clean(data.cliente_id);
    const razon     = clean(data.razon_social);
    const cotNum    = clean(data.cot_num);
    const obs       = clean(data.observaciones);

    const precioFinal = parseInt(clean(data.precio_final)||'0',10) || 0;
    const costoTrans  = parseInt(clean(data.costo_transporte)||'0',10) || 0;
    const ventaSinTransporte = Math.max(0, precioFinal - costoTrans);

    const ocNum   = clean(data.oc_num);
    const ocFecha = clean(data.oc_fecha);
    const ocUrl   = clean(data.oc_url);

    // Pago & Transporte
    const pagoCond = clean(data.pago_condicion);
    const pagoDet  = clean(data.pago_detalle);
    const transModo= clean(data.transporte_modo);
    const transPref= clean(data.transporte_pref);
    const transDet = clean(data.transporte_detalle);

    if(!razon)         return { ok:false, message:'Ingresa o selecciona Razón Social.' };
    if(precioFinal<=0) return { ok:false, message:'Precio final inválido.' };
    if(costoTrans<0)   return { ok:false, message:'Costo de transporte inválido.' };

    const rows = [];
    const itemsForEmail = [];

    data.items.forEach((it, idx)=>{
      const producto = clean(it.producto);
      const color    = clean(it.color);
      const cantidad = parseInt(clean(it.cantidad)||'0',10) || 0;
      const metodo   = clean(it.metodo);
      const tiempo   = clean(it.tiempo);
      const fechaStr = clean(it.fecha);

      const fileUrls = Array.isArray(it.file_urls) ? it.file_urls.map(clean).filter(Boolean) : [];
      const firstUrl = fileUrls.length ? fileUrls[0] : '';

      const colores        = clean(it.colores);
      const grabadoNombres = !!it.grabado_nombres;
      const grabadoUrl     = clean(it.grabado_file_url);
      const notas          = clean(it.notas);

      if (!producto) throw new Error('Producto faltante en ítem ' + (idx+1));
      if (!color)    throw new Error('Color faltante en ítem ' + (idx+1));
      if (cantidad<1)throw new Error('Cantidad inválida en ítem ' + (idx+1));
      if (!metodo)   throw new Error('Método de impresión faltante en ítem ' + (idx+1));
      if (!fechaStr) throw new Error('Fecha de entrega faltante en ítem ' + (idx+1));

      const mLow = metodo.toLowerCase();
      if ((mLow==='serigrafía' || mLow==='serigrafia' || mLow==='tampografía' || mLow==='tampografia') && !(+colores >= 1)){
        throw new Error('Debes indicar cantidad de colores en ítem ' + (idx+1));
      }
      if (mLow==='grabado láser' || mLow==='grabado laser'){
        if (grabadoNombres && !grabadoUrl) {
          throw new Error('Sube el archivo de nombres personalizados en ítem ' + (idx+1));
        }
      }

      const tiempoNum = parseTiempoNumero_(tiempo);
      const tiempoEmail = tiempoNum !== '' ? (tiempoNum + ' días hábiles') : '—';

      const parts = fechaStr.split('-').map(Number);
      const fechaEnt= new Date(parts[0], parts[1]-1, parts[2]);

      rows.push([
        fechaSolo, ordenId, idx+1, clienteId, razon, cotNum,
        producto, color, cantidad, metodo, tiempoNum,
        fechaEnt, obs,
        firstUrl,
        precioFinal, costoTrans, ventaSinTransporte,
        ocNum,
        ocFecha ? new Date(partsFromYMD_(ocFecha).y, partsFromYMD_(ocFecha).m-1, partsFromYMD_(ocFecha).d) : '',
        ocUrl,
        pagoCond, pagoDet,
        transModo, transPref, transDet,
        JSON.stringify(fileUrls),
        colores || '',
        grabadoNombres ? 'SI' : 'NO',
        grabadoUrl || '',
        notas || ''
      ]);

      itemsForEmail.push({
        n: idx+1,
        producto, color, cantidad, metodo,
        tiempo: tiempoEmail,
        fecha: Utilities.formatDate(fechaEnt, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        fileUrls,
        colores,
        grabadoNombres,
        grabadoUrl,
        notas
      });
    });

    // Escribir filas
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);

    // Formatos
    sheet.getRange(2,1,sheet.getMaxRows(),1).setNumberFormat('dd/MM/yyyy'); // Fecha registro
    sheet.getRange(2,12,sheet.getMaxRows(),1).setNumberFormat('dd/MM/yyyy'); // Fecha entrega
    sheet.getRange(2,19,sheet.getMaxRows(),1).setNumberFormat('dd/MM/yyyy'); // Fecha OC
    sheet.getRange(2,15,sheet.getMaxRows(),3).setNumberFormat('$ #,##0');    // CLP
    sheet.getRange(2,11,sheet.getMaxRows(),1).setNumberFormat('0');          // Tiempo_Produccion (número)

    /* ===== Envío de correo ===== */
    const clienteInfo = getClienteByIdOrRazon_(clienteId, razon);
    const clienteMail = (clienteInfo && clienteInfo.correo) ? clienteInfo.correo : '';

    const recipients = new Set();
    (RECIPIENTS_DEFAULT||[]).forEach(m => m && recipients.add(m));
    if (SEND_TO_EXEC) {
      const execMail = Session.getActiveUser().getEmail();
      if (execMail) recipients.add(execMail);
    }
    if (SEND_TO_CLIENT && clienteMail) recipients.add(clienteMail);

    const to = Array.from(recipients).join(',');
    const cc = Array.isArray(CC_DEFAULT) ? CC_DEFAULT.filter(Boolean).join(',') : '';

    if (to) {
      const asunto    = 'Nueva Orden ' + ordenId + ' — ' + razon;
      const emailHtml = buildEmailHtml_(
        ordenId, razon, clienteId, cotNum, obs, itemsForEmail,
        precioFinal, costoTrans, ventaSinTransporte,
        ocNum, ocFecha, ocUrl,
        pagoCond, pagoDet, transModo, transPref, transDet
      );
      const textoPlano= buildPlainTextSafe_(
        ordenId, razon, clienteId, cotNum, obs, itemsForEmail,
        precioFinal, costoTrans, ventaSinTransporte,
        ocNum, ocFecha, ocUrl,
        pagoCond, pagoDet, transModo, transPref, transDet
      );

      const mailOpts = {
        cc: cc || '',
        subject: asunto,
        htmlBody: emailHtml,
        name: 'Merchi — Órdenes',
        body: textoPlano
      };
      MailApp.sendEmail(Object.assign({ to }, mailOpts));
    }

    return { ok:true, message:'Orden '+ordenId+' guardada y correo enviado.' };

  } catch (err) {
    return { ok:false, message:'Error al guardar: ' + err.message };
  }
}


/* ============================
   Uploads a Drive (con subcarpetas por cliente)
   ============================ */
function uploadLogo(fileObj){
  return uploadToFolder_(LOGOS_FOLDER_ID, fileObj);
}
function uploadOC(fileObj){
  return uploadToFolder_((OC_FOLDER_ID || LOGOS_FOLDER_ID), fileObj);
}

function uploadToFolder_(folderId, fileObj){
  try{
    if(!fileObj || !fileObj.data) return { ok:false, message:'Sin datos de archivo.' };

    var targetFolder = getOrCreateSubFolder_(folderId, (fileObj.folderName || ''));

    var match = /^data:([^;]+);base64,(.+)$/.exec(fileObj.data);
    if(!match) return { ok:false, message:'Formato de archivo inválido.' };

    var mime  = match[1];
    var bytes = Utilities.base64Decode(match[2]);
    var safeName = (fileObj.name || ('archivo_'+Date.now())).replace(/[^\w\-. ]+/g,'_');
    var blob = Utilities.newBlob(bytes, mime, safeName);
    var file = targetFolder.createFile(blob);

    return { ok:true, url: 'https://drive.google.com/uc?export=view&id=' + file.getId() };
  }catch(err){
    return { ok:false, message: 'Upload fallo: ' + err.message };
  }
}

// Crea/obtiene subcarpeta por cliente (Razón Social)
function getOrCreateSubFolder_(parentFolderId, rawName){
  var parent = DriveApp.getFolderById(parentFolderId);
  var name = (rawName || '').toString().trim();
  if (!name) return parent;

  // Sanear nombre de carpeta
  name = name.replace(/[\\/:*?"<>|#%]+/g, ' ').replace(/\s+/g, ' ').trim();

  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}


/* ============================
   Email Helpers
   ============================ */
function fmtCLP_(n) {
  n = parseInt(n, 10) || 0;
  return '$ ' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
function rowKV_(k, v){
  return '<tr><td style="padding:4px 8px;color:#555">'+k+':</td><td style="padding:4px 8px">'+v+'</td></tr>';
}
function buildEmailHtml_(ordenId, razon, clienteId, cotNum, obs, items, precioFinal, costoTrans, ventaSin, ocNum, ocFecha, ocUrl, pagoCond, pagoDet, transModo, transPref, transDet){
  const color = BRAND.COLOR, logo = BRAND.LOGO_URL;

  const rows = items.map(it=>{
    const extras = [];
    const mLow = (it.metodo||'').toLowerCase();
    if ((mLow==='serigrafía'||mLow==='serigrafia'||mLow==='tampografía'||mLow==='tampografia') && it.colores){
      extras.push(`Colores: ${it.colores}`);
    }
    if (mLow==='grabado láser'||mLow==='grabado laser'){
      extras.push(`Nombres pers.: ${it.grabadoNombres ? 'Sí' : 'No'}`);
      if (it.grabadoNombres && it.grabadoUrl) extras.push(`Archivo nombres: ${it.grabadoUrl}`);
    }
    if (it.notas) extras.push(`Notas: ${escapeHtml_(it.notas)}`);

    return `
      <tr>
        <td style="padding:8px;border:1px solid #eee;text-align:center">${it.n}</td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.producto)}</td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.color)}</td>
        <td style="padding:8px;border:1px solid #eee;text-align:right">${it.cantidad}</td>
        <td style="padding:8px;border:1px solid #eee">
          ${escapeHtml_(it.metodo)}
          ${extras.length ? `<div style="color:#555;font-size:12px;margin-top:4px">${extras.map(escapeHtml_).join(' • ')}</div>` : ''}
        </td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.tiempo)}</td>
        <td style="padding:8px;border:1px solid #eee">${it.fecha}</td>
      </tr>`;
  }).join('');

  const urls = [];
  if (ocUrl) urls.push('OC: ' + ocUrl);
  items.forEach(it => (it.fileUrls||[]).forEach((u,idx)=> urls.push(`Producto #${it.n} (${idx+1}): ${u}`)));
  items.forEach(it => { if (it.grabadoUrl) urls.push(`Nombres #${it.n}: ${it.grabadoUrl}`); });
  const urlsHtml = urls.length
    ? '<pre style="font-family:monospace;white-space:pre-wrap;word-break:break-all;background:#fafafa;border:1px solid #eee;border-radius:8px;padding:8px;margin-top:6px">'+escapeHtml_(urls.join('\n'))+'</pre>'
    : '';

  const cotHtml = cotNum ? rowKV_('Nº Cotización', escapeHtml_(cotNum)) : '';
  const ocHtml  = (ocNum || ocFecha)
    ? `${ocNum ? rowKV_('Nº Orden de compra', escapeHtml_(ocNum)) : ''}${ocFecha ? rowKV_('Fecha OC', escapeHtml_(formatDateCL_(ocFecha))) : ''}`
    : '';

  const pagoTabla = `
    <table cellspacing="0" cellpadding="0" style="margin:6px 0 10px 0">
      ${rowKV_('Condición de pago', escapeHtml_(pagoCond || '—'))}
      ${(pagoCond==='Personalizado' && pagoDet) ? rowKV_('Detalle pago', escapeHtml_(pagoDet)) : ''}
    </table>`;
  const transTabla = `
    <table cellspacing="0" cellpadding="0" style="margin:6px 0 10px 0">
      ${rowKV_('Transporte', escapeHtml_(transModo || '—'))}
      ${rowKV_('Preferencia', escapeHtml_(transPref || '—'))}
      ${transDet ? rowKV_('Detalle', escapeHtml_(transDet)) : ''}
    </table>`;

  return `
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111">
    <div style="text-align:center;margin-bottom:12px"><img src="${logo}" alt="Merchi" style="max-width:160px;height:auto"/></div>
    <div style="height:6px;background:${color};border-radius:4px;margin-bottom:10px"></div>

    <h2 style="margin:0 0 8px 0;color:${color}">Nueva Orden ${ordenId}</h2>
    <p style="margin:0 0 12px 0">Razón Social: <strong>${escapeHtml_(razon)}</strong>${clienteId ? ` <span style="color:#777">[ID: ${escapeHtml_(clienteId)}]</span>` : ``}</p>
    ${cotHtml ? `<table cellspacing="0" cellpadding="0">${cotHtml}</table>` : ''}
    ${obs ? `<p style="margin:0 0 12px 0"><strong>Observaciones:</strong><br>${escapeHtml_(obs)}</p>` : ``}
    ${ocHtml ? `<table cellspacing="0" cellpadding="0" style="margin:6px 0 10px 0">${ocHtml}</table>` : ''}

    <h3 style="margin:12px 0 4px 0;color:#333">Pago & Transporte</h3>
    ${pagoTabla}
    ${transTabla}

    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin:10px 0">
      <thead>
        <tr style="background:${BRAND.COLOR_SOFT}">
          <th style="padding:8px;border:1px solid #eee;text-align:center">#</th>
          <th style="padding:8px;border:1px solid #eee">Producto</th>
          <th style="padding:8px;border:1px solid #eee">Color</th>
          <th style="padding:8px;border:1px solid #eee;text-align:right">Cantidad</th>
          <th style="padding:8px;border:1px solid #eee">Impresión</th>
          <th style="padding:8px;border:1px solid #eee">Tiempo</th>
          <th style="padding:8px;border:1px solid #eee">Entrega</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>

    <table cellspacing="0" cellpadding="0" style="margin:10px 0 0 0">
      ${rowKV_('Precio final de la orden', fmtCLP_(precioFinal))}
      ${rowKV_('Costo de transporte', fmtCLP_(costoTrans))}
      ${rowKV_('<strong>Venta sin transporte</strong>', '<strong>'+fmtCLP_(ventaSin)+'</strong>')}
    </table>

    ${urlsHtml}
    <div style="height:1px;background:#eee;margin:14px 0"></div>
    <p style="color:#666;font-size:12px">${escapeHtml_(BRAND.FOOTER)}</p>
  </div>`;
}

function buildPlainTextSafe_(ordenId, razon, clienteId, cotNum, obs, items, precio, transp, ventaSin, ocNum, ocFecha, ocUrl, pagoCond, pagoDet, transModo, transPref, transDet){
  let t = '';
  t += 'Nueva Orden ' + ordenId + '\n';
  t += 'Razón Social: ' + (razon||'') + (clienteId ? (' [ID: ' + clienteId + ']') : '') + '\n';
  if (cotNum) t += 'Nº Cotización: ' + cotNum + '\n';
  if (obs) t += 'Observaciones: ' + obs + '\n';
  if (ocNum)   t += 'Nº Orden de compra: ' + ocNum + '\n';
  if (ocFecha) t += 'Fecha OC: ' + formatDateCL_(ocFecha) + '\n';

  t += '\nPago & Transporte:\n';
  t += '- Condición de pago: ' + (pagoCond || '—') + '\n';
  if (pagoCond === 'Personalizado' && pagoDet) t += '  Detalle: ' + pagoDet + '\n';
  t += '- Transporte: ' + (transModo || '—') + '\n';
  t += '  Preferencia: ' + (transPref || '—') + (transDet ? (' | Detalle: ' + transDet) : '') + '\n';

  t += '\nÍtems:\n';
  items.forEach(it=>{
    const extras = [];
    const mLow = (it.metodo||'').toLowerCase();
    if ((mLow==='serigrafía'||mLow==='serigrafia'||mLow==='tampografía'||mLow==='tampografia') && it.colores){
      extras.push('Colores: '+it.colores);
    }
    if (mLow==='grabado láser'||mLow==='grabado laser'){
      extras.push('Nombres pers.: ' + (it.grabadoNombres ? 'Sí' : 'No'));
      if (it.grabadoNombres && it.grabadoUrl) extras.push('Archivo nombres: ' + it.grabadoUrl);
    }
    if (it.notas) extras.push('Notas: '+it.notas);

    t += '- #' + it.n + ' ' + (it.producto||'') + ' | Color ' + (it.color||'') + ' | Cant ' + (it.cantidad||'') +
         ' | Impresión ' + (it.metodo||'') +
         (extras.length ? (' ['+extras.join(' | ')+']') : '') +
         ' | Tiempo ' + (it.tiempo||'') + ' | Entrega ' + (it.fecha||'') + '\n';
  });

  t += '\nTotales:\n';
  t += 'Precio final de la orden: ' + fmtCLP_(precio) + '\n';
  t += 'Costo de transporte: ' + fmtCLP_(transp) + '\n';
  t += 'Venta sin transporte: ' + fmtCLP_(ventaSin) + '\n';

  const urls = [];
  if (ocUrl) urls.push('OC: ' + ocUrl);
  items.forEach(it=> (it.fileUrls||[]).forEach((u,idx)=> urls.push(`Producto #${it.n} (${idx+1}): ${u}`)));
  items.forEach(it=> { if (it.grabadoUrl) urls.push(`Nombres #${it.n}: ${it.grabadoUrl}`); });
  if (urls.length){ t += '\nURLs:\n' + urls.join('\n') + '\n'; }
  return t;
}


/* ============================
   Utils
   ============================ */
function generarIdCorto_() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let out = '';
  for (let i = 0; i < 5; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}
function escapeHtml_(s){
  return (s||'').toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function partsFromYMD_(ymd){
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd||'');
  if (!m) return {y:0,m:0,d:0};
  return { y: +m[1], m:+m[2], d:+m[3] };
}
function formatDateCL_(ymd){
  const p = partsFromYMD_(ymd);
  if (!p.y) return '';
  const d = new Date(p.y, p.m-1, p.d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}
function parseTiempoNumero_(s){
  s = (s||'').toString();
  var m = s.match(/(\d+)/);
  return m ? parseInt(m[1],10) : '';
}