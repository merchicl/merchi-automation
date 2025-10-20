/***** CONFIG *****/
const SHEET_NAME = 'Bocetos & Aprobaciones'; // hoja BD central (uso programático)
const ROOT_PEDIDOS_FOLDER_ID = '1naPzQUprFBgpoUD0EGt3CPhiWQ8pz1aS'; // Carpeta raíz en Drive

// IDs y pestañas de origen (tus IDs reales)
const SSID_COTIZADOR = '1o9q9cwgx4f1iV87yrkdI4a5JR8El_coZa24VsKWMQhs';
const HOJA_COTIZADOR = 'Datos Cotizaciones';
const SSID_ORDENES   = '1GUPTg4H4V2OIUGrX4LkUzA2mJgjdN7USlzbcPJLiheM';
const HOJA_ORDENES   = 'Ordenes';

// URL del WebApp /exec (PEGA AQUÍ TU URL DE DESPLIEGUE)
const WEBAPP_EXEC_URL = 'https://script.google.com/macros/s/AKfycbxcH-uqx5PCS2dHfu6LvoJHEQSAINcZ3Kbr3MsgqO_FRGH1rLwR1gLme1OnwRnNFSi7/exec';

// Correo de Producción
const CORREO_PRODUCCION = 'ordenes.merchi@gmail.com';

/***** ESTADOS *****/
const ESTADOS = {
  BORRADOR: 'Borrador',
  ENVIADO: 'Enviado',
  REV_EJEC: 'En revisión ejecutivo',
  ENV_CLIENTE: 'Enviado a cliente',
  CAMBIOS: 'Cambios',
  APROBADO: 'Aprobado',
  APROB_INT: 'Aprobado (interno)',
  CERRADO: 'Cerrado'
};

/***** EJECUTIVOS *****/
const EXECUTIVOS = [
  { nombre: 'Sergio Flores', email: 'sergio@merchi.cl' },
  { nombre: 'Pablo Ramírez', email: 'pablo@merchi.cl' },
];
const DEFAULT_EXECUTIVO_FALLBACK = 'pablo@merchi.cl';

function _normalize_(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim(); }
function resolveEjecutivoEmail_(value){
  if (!value) return DEFAULT_EXECUTIVO_FALLBACK;
  const v = String(value).trim();
  if (/@/.test(v)) return v;
  const nv = _normalize_(v);
  const found = EXECUTIVOS.find(e => _normalize_(e.nombre) === nv);
  return found ? found.email : DEFAULT_EXECUTIVO_FALLBACK;
}
function ejecutivoNameByEmail_(email){
  const e = EXECUTIVOS.find(x => x.email.toLowerCase() === String(email||'').toLowerCase());
  return e ? e.nombre : email || 'Ejecutivo';
}

/***** MENÚ *****/
function onOpen(){
  try {
    SpreadsheetApp.getUi()
      .createMenu('Merchi')
      .addItem('Abrir panel de Producción', 'openProdPanel_')
      .addItem('Diagnóstico de pedido', 'uiDiagnosticoLink_')
      .addToUi();
  } catch(_) {}
}

/***** BD HELPERS *****/
function getSheet_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('No existe la hoja "'+SHEET_NAME+'". Crea la pestaña y pega los encabezados.');
  return sh;
}
function ensureHeaders_(){
  const need = [
    'Fecha_Creación','Nro_Pedido','ID_Cliente','Empresa','Contacto','Correo',
    'Monto_Total','Cot_Num',
    'Pago_Condicion','Pago_Detalle','Transporte_Modo','Transporte_Preferencia','Transporte_Detalle',
    'OC_Num','OC_Fecha','OC_URL',
    'Items_JSON',
    'Estado','Versión_Actual',
    'Carpeta_Bocetos_FolderId','Carpeta_Aprobaciones_FolderId',
    'PDFs_Vigentes_URLs',
    'Token_Acceso','Link_Portal_Cliente',
    'Fecha_Envío_Cliente','Fecha_Respuesta_Cliente',
    'Respuesta','Comentarios_Cliente','Archivo_Cliente_URL',
    'Historial_JSON','Ejecutivo','Notas_Internas'
  ];
  const sh = getSheet_();
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,need.length).setValues([need]);
    return;
  }
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  need.forEach(h=>{
    if (header.indexOf(h)===-1){
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn(), 1, 1).setValue(h);
    }
  });
}
function getHeaderIndex_(header, name){ return header.indexOf(name); }
function findRowByPedido_(nroPedido){
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  const header = data[0] || [];
  const idx = (n)=>header.indexOf(n);
  for (let i=1;i<data.length;i++){
    if (String(data[i][idx('Nro_Pedido')]) === String(nroPedido)){
      return { row: i+1, header, values: data[i] };
    }
  }
  return null;
}

/***** DRIVE *****/
function getOrCreatePedidoFolders_(nroPedido){
  const root = DriveApp.getFolderById(ROOT_PEDIDOS_FOLDER_ID);
  const name = String(nroPedido).trim();
  const it = root.getFoldersByName(name);
  const pedidoFolder = it.hasNext() ? it.next() : root.createFolder(name);
  const bocIt = pedidoFolder.getFoldersByName('Bocetos');
  const bocFolder = bocIt.hasNext() ? bocIt.next() : pedidoFolder.createFolder('Bocetos');
  const aprIt = pedidoFolder.getFoldersByName('Aprobaciones');
  const aprFolder = aprIt.hasNext() ? aprIt.next() : pedidoFolder.createFolder('Aprobaciones');
  return { pedidoFolder, bocFolder, aprFolder };
}

/***** TOKEN & LINK *****/
function randToken_(len=28){
  const chars='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t=''; for (let i=0;i<len;i++) t += chars.charAt(Math.floor(Math.random()*chars.length));
  return t;
}
function makePortalLink_(token, nro){ return WEBAPP_EXEC_URL + '?t=' + encodeURIComponent(token) + '&p=' + encodeURIComponent(nro); }

/***** BRIDGE: lee TODOS los ítems del pedido desde Órdenes + enriquece con Cotizador *****/
function getDatosPedido_(nroPedido){
  const shO = SpreadsheetApp.openById(SSID_ORDENES).getSheetByName(HOJA_ORDENES);
  if (!shO) throw new Error('No encuentro la hoja "'+HOJA_ORDENES+'" en Órdenes.');
  const oVals = shO.getDataRange().getValues();
  const hO = oVals[0] || [];
  const ixO = (n)=>hO.indexOf(n);

  // Requeridas mínimas
  const MUST = ['Orden_ID','Cliente_ID','Razon_Social','Producto','Color','Cantidad','Metodo_Impresion'];
  const missMust = MUST.filter(c => ixO(c) === -1);
  if (missMust.length) throw new Error('Faltan columnas en Órdenes (obligatorias): ' + missMust.join(', '));

  // Opcionales (no rompen si faltan)
  const OPT = [
    'Tiempo_Produccion','Fecha_Entrega','Adjunto_URL','Adjuntos_JSON',
    'Colores_Serigrafia_Tampografia','Grabado_Nombres','Grabado_Archivo_URL','Notas_Producto',
    'Precio_Final_Orden','Venta_Sin_Transporte',
    'Pago_Condicion','Pago_Detalle','Transporte_Modo','Transporte_Preferencia','Transporte_Detalle',
    'OC_Num','OC_Fecha','OC_URL','Cot_Num','Ejecutivo'
  ];
  const I = {};
  MUST.concat(OPT).forEach(c => I[c] = ixO(c)); // -1 si no existe

  const items = [];
  let firstRow = null;

  for (let i=1;i<oVals.length;i++){
    if (String(oVals[i][I['Orden_ID']]) === String(nroPedido)){
      if (!firstRow) firstRow = oVals[i];

      // Adjuntos múltiples
      const fileUrls = [];
      const one = (I['Adjunto_URL']!==-1 ? (oVals[i][I['Adjunto_URL']]||'') : '').toString().trim();
      if (one) fileUrls.push(one);
      const many = (I['Adjuntos_JSON']!==-1 ? (oVals[i][I['Adjuntos_JSON']]||'') : '').toString().trim();
      if (many){
        try{
          const arr = JSON.parse(many);
          if (Array.isArray(arr)) arr.forEach(u=> u && fileUrls.push(String(u)));
        }catch(_){}
      }

      const tiempoNum = (I['Tiempo_Produccion']!==-1 ? oVals[i][I['Tiempo_Produccion']] : '');
      const fechaEnt  = (I['Fecha_Entrega']!==-1 && oVals[i][I['Fecha_Entrega']])
          ? Utilities.formatDate(new Date(oVals[i][I['Fecha_Entrega']]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '—';

      items.push({
        n: items.length+1,
        producto: oVals[i][I['Producto']] || '',
        color:    oVals[i][I['Color']] || '',
        cantidad: oVals[i][I['Cantidad']] || '',
        metodo:   oVals[i][I['Metodo_Impresion']] || '',
        tiempo:   (tiempoNum !== '' ? (tiempoNum + ' días hábiles') : '—'),
        fecha:    fechaEnt,
        colores:  (I['Colores_Serigrafia_Tampografia']!==-1 ? (oVals[i][I['Colores_Serigrafia_Tampografia']]||'') : ''),
        grabadoNombres: (I['Grabado_Nombres']!==-1 ? (String(oVals[i][I['Grabado_Nombres']]||'').toUpperCase()==='SI') : false),
        grabadoUrl: (I['Grabado_Archivo_URL']!==-1 ? (oVals[i][I['Grabado_Archivo_URL']]||'').toString().trim() : ''),
        notas: (I['Notas_Producto']!==-1 ? (oVals[i][I['Notas_Producto']]||'').toString().trim() : ''),
        fileUrls
      });
    }
  }
  if (!firstRow) throw new Error('Orden_ID no encontrado en Órdenes ('+nroPedido+').');

  // Cabecera del pedido + totales/condiciones (todas opcionales salvo las MUST)
  const clienteId = firstRow[I['Cliente_ID']] || '';
  const razonOrden = (firstRow[I['Razon_Social']]||'').toString().trim();
  const montoTotal = (I['Precio_Final_Orden']!==-1 && firstRow[I['Precio_Final_Orden']]!=='')
      ? firstRow[I['Precio_Final_Orden']]
      : (I['Venta_Sin_Transporte']!==-1 ? firstRow[I['Venta_Sin_Transporte']] : '');

  const pagoCond   = (I['Pago_Condicion']!==-1 ? firstRow[I['Pago_Condicion']] : '');
  const pagoDet    = (I['Pago_Detalle']!==-1 ? firstRow[I['Pago_Detalle']] : '');
  const transModo  = (I['Transporte_Modo']!==-1 ? firstRow[I['Transporte_Modo']] : '');
  const transPref  = (I['Transporte_Preferencia']!==-1 ? firstRow[I['Transporte_Preferencia']] : '');
  const transDet   = (I['Transporte_Detalle']!==-1 ? firstRow[I['Transporte_Detalle']] : '');
  const ocNum      = (I['OC_Num']!==-1 ? firstRow[I['OC_Num']] : '');
  const ocFecha    = (I['OC_Fecha']!==-1 ? firstRow[I['OC_Fecha']] : '');
  const ocUrl      = (I['OC_URL']!==-1 ? firstRow[I['OC_URL']] : '');
  const cotNum     = (I['Cot_Num']!==-1 ? firstRow[I['Cot_Num']] : '');
  const ejecutivo  = (I['Ejecutivo']!==-1 ? firstRow[I['Ejecutivo']] : '');

  // COTIZADOR
  const shC = SpreadsheetApp.openById(SSID_COTIZADOR).getSheetByName(HOJA_COTIZADOR);
  if (!shC) throw new Error('No encuentro la hoja "'+HOJA_COTIZADOR+'" en Cotizador.');
  const cVals = shC.getDataRange().getValues(); const ixC=(n)=>cVals[0].indexOf(n);
  const needC=['ID_Registro','Empresa','Contacto_Nombre','Correo'];
  const missC = needC.filter(n=>ixC(n)===-1);
  if (missC.length) throw new Error('Faltan columnas en Cotizador: '+missC.join(', '));

  let Empresa='', Contacto='', Correo='';
  for (let i=1;i<cVals.length;i++){
    if (String(cVals[i][ixC('ID_Registro')])===String(clienteId)){
      Empresa = cVals[i][ixC('Empresa')] || '';
      Contacto= cVals[i][ixC('Contacto_Nombre')] || '';
      Correo  = cVals[i][ixC('Correo')] || '';
      break;
    }
  }
  if (razonOrden) Empresa = razonOrden; // prioriza razón desde Órdenes

  return {
    ID_Cliente: clienteId,
    Empresa, Contacto, Correo,
    Monto_Total: montoTotal,
    Cot_Num: cotNum,
    Pago_Condicion: pagoCond,
    Pago_Detalle: pagoDet,
    Transporte_Modo: transModo,
    Transporte_Preferencia: transPref,
    Transporte_Detalle: transDet,
    OC_Num: ocNum,
    OC_Fecha: ocFecha,
    OC_URL: ocUrl,
    Items: items,
    Ejecutivo: ejecutivo // puede venir vacío si la columna no existe
  };
}

/***** CREA/ACTUALIZA FILA (multi-ítem) *****/
function ensurePedidoRow_(nroPedido){
  ensureHeaders_();
  const sh = getSheet_();
  let rowObj = findRowByPedido_(nroPedido);
  const datos = getDatosPedido_(nroPedido);
  const { bocFolder, aprFolder } = getOrCreatePedidoFolders_(nroPedido);

  let row, header, values;
  if (rowObj){ row=rowObj.row; header=rowObj.header; values=rowObj.values; }
  else {
    const last = Math.max(1, sh.getLastRow());
    if (sh.getLastRow()===0) ensureHeaders_();
    sh.insertRowAfter(last);
    row = last+1;
    header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    values = new Array(header.length).fill('');
  }
  const idx=(n)=>header.indexOf(n);
  const safe = (v)=> (v==null?'':v);

  values[idx('Fecha_Creación')] = values[idx('Fecha_Creación')] || new Date();
  values[idx('Nro_Pedido')]     = nroPedido;
  values[idx('ID_Cliente')]     = datos.ID_Cliente;
  values[idx('Empresa')]        = datos.Empresa;
  values[idx('Contacto')]       = datos.Contacto;
  values[idx('Correo')]         = datos.Correo;
  values[idx('Monto_Total')]    = datos.Monto_Total;
  values[idx('Cot_Num')]        = datos.Cot_Num;

  values[idx('Pago_Condicion')] = safe(datos.Pago_Condicion);
  values[idx('Pago_Detalle')]   = safe(datos.Pago_Detalle);
  values[idx('Transporte_Modo')]        = safe(datos.Transporte_Modo);
  values[idx('Transporte_Preferencia')] = safe(datos.Transporte_Preferencia);
  values[idx('Transporte_Detalle')]     = safe(datos.Transporte_Detalle);

  values[idx('OC_Num')]   = safe(datos.OC_Num);
  values[idx('OC_Fecha')] = safe(datos.OC_Fecha)? new Date(datos.OC_Fecha):'';
  values[idx('OC_URL')]   = safe(datos.OC_URL);

  values[idx('Items_JSON')] = JSON.stringify(datos.Items);
  values[idx('Estado')] = values[idx('Estado')] || ESTADOS.ENVIADO;
  values[idx('Versión_Actual')] = values[idx('Versión_Actual')] || 1;

  values[idx('Carpeta_Bocetos_FolderId')]    = bocFolder.getId();
  values[idx('Carpeta_Aprobaciones_FolderId')] = aprFolder.getId();

  // Ejecutivo
  const ixEj = idx('Ejecutivo');
  const actualEj = (ixEj!==-1 ? values[ixEj] : '');
  const resolvedEj = resolveEjecutivoEmail_( actualEj || datos.Ejecutivo );
  if (ixEj!==-1) values[ixEj] = resolvedEj;

  // Token + Link
  if (!values[idx('Token_Acceso')]) values[idx('Token_Acceso')] = randToken_();
  const linkPortal = makePortalLink_(values[idx('Token_Acceso')], nroPedido);
  values[idx('Link_Portal_Cliente')] = linkPortal;

  // PDFs vigentes
  const version = Number(values[idx('Versión_Actual')])||1;
  let vFolder = null;
  const it = bocFolder.getFoldersByName('v'+version);
  if (it.hasNext()) vFolder = it.next();
  const scanFolder = vFolder || bocFolder;
  const pdfUrls=[];
  const files = scanFolder.getFiles();
  while(files.hasNext()){
    const f=files.next();
    if (f.getMimeType()===MimeType.PDF) pdfUrls.push('https://drive.google.com/uc?export=download&id='+f.getId());
  }
  values[idx('PDFs_Vigentes_URLs')] = pdfUrls.join(' | ');

  // Historial
  const now = new Date().toISOString();
  const hist = values[idx('Historial_JSON')] ? JSON.parse(values[idx('Historial_JSON')]) : [];
  if (!hist.length) hist.push({ ts: now, ev: 'INIT', version, pdfs: pdfUrls });
  values[idx('Historial_JSON')] = JSON.stringify(hist);

  sh.getRange(row,1,1,header.length).setValues([values]);
  return { header, values, row, linkPortal };
}

/***** EMAIL HELPERS (estilo Órdenes) *****/
function fmtCLP_(n){ n=parseInt(n,10)||0; return '$ ' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function escapeHtml_(s){
  return (s||'').toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function rowKV_(k, v){ return '<tr><td style="padding:4px 8px;color:#555">'+k+':</td><td style="padding:4px 8px">'+v+'</td></tr>'; }

function buildItemsTableHtml_(items){
  const rows = items.map(it=>{
    const extras=[];
    const mLow = (it.metodo||'').toLowerCase();
    if ((/serigraf(í|i)a|tampograf(í|i)a/.test(mLow)) && it.colores) extras.push('Colores: '+escapeHtml_(it.colores));
    if (/grabado (láser|laser)/.test(mLow)){
      extras.push('Nombres pers.: ' + (it.grabadoNombres ? 'Sí':'No'));
      if (it.grabadoNombres && it.grabadoUrl) extras.push('Archivo nombres: '+escapeHtml_(it.grabadoUrl));
    }
    if (it.notas) extras.push('Notas: '+escapeHtml_(it.notas));

    return `
      <tr>
        <td style="padding:8px;border:1px solid #eee;text-align:center">${it.n}</td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.producto)}</td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.color)}</td>
        <td style="padding:8px;border:1px solid #eee;text-align:right">${it.cantidad}</td>
        <td style="padding:8px;border:1px solid #eee">
          ${escapeHtml_(it.metodo)}
          ${extras.length ? `<div style="color:#555;font-size:12px;margin-top:4px">${extras.join(' • ')}</div>` : ''}
        </td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.tiempo||'—')}</td>
        <td style="padding:8px;border:1px solid #eee">${escapeHtml_(it.fecha||'—')}</td>
      </tr>`;
  }).join('');

  return `
    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin:10px 0">
      <thead>
        <tr style="background:#f1f5ff">
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
    </table>`;
}

function buildUrlsBlock_(pdfUrls, items, ocUrl){
  const list = [];
  if (ocUrl) list.push('OC: ' + ocUrl);
  (pdfUrls||[]).forEach((u,i)=> list.push(`Boceto PDF #${i+1}: ${u}`));
  (items||[]).forEach(it=>{
    (it.fileUrls||[]).forEach((u,idx)=> list.push(`Arte Producto #${it.n} (${idx+1}): ${u}`));
    if (it.grabadoUrl) list.push(`Nombres #${it.n}: ${it.grabadoUrl}`);
  });
  if (!list.length) return '';
  return '<pre style="font-family:monospace;white-space:pre-wrap;word-break:break-all;background:#fafafa;border:1px solid #eee;border-radius:8px;padding:8px;margin-top:6px">'
         + escapeHtml_(list.join('\n')) + '</pre>';
}

/***** EMAILS (ricos y completos) *****/
function sendToEjecutivo_(values, header){
  const idx=(n)=>header.indexOf(n);
  const nro = values[idx('Nro_Pedido')];
  const linkPortal = values[idx('Link_Portal_Cliente')];
  const token = values[idx('Token_Acceso')];
  const ejEmail = resolveEjecutivoEmail_(values[idx('Ejecutivo')]);
  const ejName  = ejecutivoNameByEmail_(ejEmail);

  const items = JSON.parse(values[idx('Items_JSON')]||'[]');
  const cotNum= values[idx('Cot_Num')] || '';
  const ocNum = values[idx('OC_Num')] || '';
  const ocF   = values[idx('OC_Fecha')] ? Utilities.formatDate(new Date(values[idx('OC_Fecha')]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  const ocUrl = values[idx('OC_URL')] || '';
  const pagoC = values[idx('Pago_Condicion')] || '';
  const pagoD = values[idx('Pago_Detalle')] || '';
  const tModo = values[idx('Transporte_Modo')] || '';
  const tPref = values[idx('Transporte_Preferencia')] || '';
  const tDet  = values[idx('Transporte_Detalle')] || '';
  const monto = values[idx('Monto_Total')] || '';
  const pdfs = String(values[idx('PDFs_Vigentes_URLs')]||'').split('|').map(s=>s.trim()).filter(Boolean);

  const actionApprove = WEBAPP_EXEC_URL + '?a=approve_internal&t=' + encodeURIComponent(token) + '&p=' + encodeURIComponent(nro);
  const actionClient  = WEBAPP_EXEC_URL + '?a=send_to_client&t=' + encodeURIComponent(token) + '&p=' + encodeURIComponent(nro);

  const tabla = buildItemsTableHtml_(items);
  const urls  = buildUrlsBlock_(pdfs, items, ocUrl);
  const pagoTabla = `
    <table cellspacing="0" cellpadding="0" style="margin:6px 0 10px 0">
      ${rowKV_('Condición de pago', escapeHtml_(pagoC || '—'))}
      ${(pagoC==='Personalizado' && pagoD) ? rowKV_('Detalle pago', escapeHtml_(pagoD)) : ''}
    </table>`;
  const transTabla = `
    <table cellspacing="0" cellpadding="0" style="margin:6px 0 10px 0">
      ${rowKV_('Transporte', escapeHtml_(tModo || '—'))}
      ${rowKV_('Preferencia', escapeHtml_(tPref || '—'))}
      ${tDet ? rowKV_('Detalle', escapeHtml_(tDet)) : ''}
    </table>`;
  const totales = `
    <table cellspacing="0" cellpadding="0" style="margin:10px 0 0 0">
      ${rowKV_('Monto total estimado', fmtCLP_(monto))}
      ${ocNum ? rowKV_('Nº Orden de compra', escapeHtml_(ocNum)) : ''}
      ${ocF ? rowKV_('Fecha OC', escapeHtml_(ocF)) : ''}
      ${cotNum ? rowKV_('Nº Cotización', escapeHtml_(cotNum)) : ''}
    </table>`;

  const html = `
    <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111">
      <h2 style="margin:0 0 8px 0;color:#0b53d0">Revisión ejecutiva — Pedido ${nro} (v${values[idx('Versión_Actual')]})</h2>
      <p style="margin:0 0 10px">Empresa: <b>${escapeHtml_(values[idx('Empresa')]||'')}</b></p>
      ${pagoTabla}${transTabla}
      ${tabla}
      ${totales}
      ${urls}
      <p style="margin:14px 0 6px 0"><a href="${linkPortal}" target="_blank">🔗 Abrir portal (PDFs)</a></p>
      <p><a href="${actionApprove}" target="_blank">✅ Aprobar interno</a> &nbsp; | &nbsp; <a href="${actionClient}" target="_blank">✉️ Enviar a cliente</a></p>
    </div>`;

  MailApp.sendEmail({ to: ejEmail, subject: `Revisión ejecutiva — Pedido ${nro} v${values[idx('Versión_Actual')]}`, htmlBody: html });
}

function sendToCliente_(values, header){
  const idx=(n)=>header.indexOf(n);
  const correo = values[idx('Correo')];
  if (!correo) return;

  const nro = values[idx('Nro_Pedido')];
  const linkPortal = values[idx('Link_Portal_Cliente')];
  const ejEmail = resolveEjecutivoEmail_(values[idx('Ejecutivo')]);

  const items = JSON.parse(values[idx('Items_JSON')]||'[]');
  const tabla = buildItemsTableHtml_(items);
  const pdfs  = String(values[idx('PDFs_Vigentes_URLs')]||'').split('|').map(s=>s.trim()).filter(Boolean);
  const urls  = buildUrlsBlock_(pdfs, items, values[idx('OC_URL')]||'');

  const htmlCli = `
    <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111">
      <p>Hola,</p>
      <p>Te compartimos los bocetos para revisión del pedido <b>${nro}</b>.</p>
      ${tabla}
      <p>Revisa y aprueba/solicita cambios aquí: <a href="${linkPortal}" target="_blank">${linkPortal}</a></p>
      ${urls}
      <p>Saludos,<br>Merchi</p>
    </div>`;

  MailApp.sendEmail({ to: correo, subject: `Revisión de bocetos — Pedido ${nro}`, htmlBody: htmlCli });
  MailApp.sendEmail({ to: ejEmail, subject: `Cliente notificado — Pedido ${nro}`, htmlBody: `<p>Enviado a <b>${escapeHtml_(correo)}</b>. Portal: <a href="${linkPortal}" target="_blank">${linkPortal}</a></p>` });
}

/***** doGet / doPost *****/
function doGet(e){
  try{
    const r = (e.parameter.r||'').trim();
    const a = (e.parameter.a||'').trim();
    const t = (e.parameter.t||'').trim();
    const p = (e.parameter.p||'').trim();

    if (r === 'prod'){
      const tpl = HtmlService.createTemplateFromFile('prod');
      tpl.execUrl = WEBAPP_EXEC_URL;
      return tpl.evaluate().setTitle('Merchi · Producción').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (a){
      if (!t || !p) return HtmlService.createHtmlOutput('Faltan parámetros (?t & ?p).');
      const sh = getSheet_(); const data = sh.getDataRange().getValues(); const header=data[0]; const idx=(n)=>header.indexOf(n);
      let rowIndex=-1, values=null;
      for (let i=1;i<data.length;i++){
        if (String(data[i][idx('Nro_Pedido')])===p && String(data[i][idx('Token_Acceso')])===t){ rowIndex=i; values=data[i].slice(); break; }
      }
      if (rowIndex===-1) return HtmlService.createHtmlOutput('No autorizado.');

      const hist = values[idx('Historial_JSON')]?JSON.parse(values[idx('Historial_JSON')]):[];
      const ejEmail = resolveEjecutivoEmail_(values[idx('Ejecutivo')]);

      if (a === 'approve_internal'){
        values[idx('Estado')] = ESTADOS.APROB_INT;
        hist.push({ ts:new Date().toISOString(), ev:'EJECUTIVO_APRUEBA_INTERNO' });
        values[idx('Historial_JSON')] = JSON.stringify(hist);
        sh.getRange(rowIndex+1,1,1,header.length).setValues([values]);

        const items = JSON.parse(values[idx('Items_JSON')]||'[]');
        const tabla = buildItemsTableHtml_(items);
        const linkPortal = values[idx('Link_Portal_Cliente')];
        MailApp.sendEmail({
          to: CORREO_PRODUCCION,
          cc: ejEmail,
          subject: `APROBADO INTERNO — Pedido ${values[idx('Nro_Pedido')]}`,
          htmlBody: `<p>El Ejecutivo ha aprobado internamente el pedido <b>${values[idx('Nro_Pedido')]}</b> (${escapeHtml_(values[idx('Empresa')]||'')}).</p>${tabla}<p><a href="${linkPortal}" target="_blank">Portal</a></p>`
        });

        return HtmlService.createHtmlOutput('✅ Aprobación interna registrada. Producción notificada por correo.');
      }

      if (a === 'send_to_client'){
        values[idx('Estado')] = ESTADOS.ENV_CLIENTE;
        hist.push({ ts:new Date().toISOString(), ev:'EJECUTIVO_ENVIA_CLIENTE' });
        values[idx('Historial_JSON')] = JSON.stringify(hist);
        sh.getRange(rowIndex+1,1,1,header.length).setValues([values]);
        sendToCliente_(values, header);
        return HtmlService.createHtmlOutput('✉️ Enviado al cliente. Ejecutivo copiado.');
      }

      return HtmlService.createHtmlOutput('Acción no reconocida.');
    }

    if (!t || !p) return HtmlService.createHtmlOutput('Link inválido: faltan parámetros (?t o ?p).');
    const sh = getSheet_(); const data = sh.getDataRange().getValues(); const header=data[0]; const idx=(n)=>header.indexOf(n);
    let rowValues=null;
    for (let i=1;i<data.length;i++){
      if (String(data[i][idx('Nro_Pedido')])===p && String(data[i][idx('Token_Acceso')])===t){ rowValues=data[i]; break; }
    }
    if (!rowValues) return HtmlService.createHtmlOutput('Acceso no autorizado o pedido/token no coinciden.');

    const tpl = HtmlService.createTemplateFromFile('client');
    tpl.pedido = {
      nro: rowValues[idx('Nro_Pedido')],
      empresa: rowValues[idx('Empresa')],
      contacto: rowValues[idx('Contacto')],
      correo: rowValues[idx('Correo')],
      monto: rowValues[idx('Monto_Total')],
      cotnum: rowValues[idx('Cot_Num')],
      pago: { cond: rowValues[idx('Pago_Condicion')], det: rowValues[idx('Pago_Detalle')] },
      transporte: { modo: rowValues[idx('Transporte_Modo')], pref: rowValues[idx('Transporte_Preferencia')], det: rowValues[idx('Transporte_Detalle')] },
      oc: { num: rowValues[idx('OC_Num')], fecha: rowValues[idx('OC_Fecha')] ? Utilities.formatDate(new Date(rowValues[idx('OC_Fecha')]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '', url: rowValues[idx('OC_URL')] },
      estado: rowValues[idx('Estado')],
      version: rowValues[idx('Versión_Actual')],
      pdfs: String(rowValues[idx('PDFs_Vigentes_URLs')]||'').split('|').map(s=>s.trim()).filter(Boolean),
      token: t,
      items: JSON.parse(rowValues[idx('Items_JSON')]||'[]')
    };
    tpl.postUrl = WEBAPP_EXEC_URL;
    return tpl.evaluate().setTitle('Aprobación de bocetos | Merchi').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  }catch(err){
    return HtmlService.createHtmlOutput('Error inesperado: '+err.message);
  }
}

function doPost(e){
  try{
    const kind = (e.parameter.kind||'').trim(); // 'prod_publish' | 'client_reply'
    if (kind === 'prod_publish') return doPostProd_(e);
    if (kind === 'client_reply') return doPostClient_(e);
    return HtmlService.createHtmlOutput('Solicitud desconocida.');
  }catch(err){
    return HtmlService.createHtmlOutput('Error: '+err.message);
  }
}

function doPostProd_(e){
  const nroPedido = (e.parameter.pedido||'').trim();
  const enviarA   = (e.parameter.enviarA||'ejecutivo').trim(); // 'ejecutivo' | 'cliente'
  const ejecutivoSel = (e.parameter.ejecutivoSel || '').trim();
  if (!nroPedido) return HtmlService.createHtmlOutput('Falta Nro_Pedido.');

  const row = ensurePedidoRow_(nroPedido);
  const header=row.header; let values=row.values; const idx=(n)=>header.indexOf(n);

  if (ejecutivoSel){
    const ixEj = idx('Ejecutivo');
    if (ixEj !== -1) values[ixEj] = resolveEjecutivoEmail_(ejecutivoSel);
  }

  const bocFolder = DriveApp.getFolderById(values[idx('Carpeta_Bocetos_FolderId')]);
  const uploadBlobs=[];
  if (e.files){
    Object.keys(e.files).forEach(k=>{
      const f=e.files[k];
      if (f && f.length !== 0){
        const blob=Utilities.newBlob(f.data, f.mimeType||MimeType.PDF, f.filename||('boceto_'+Date.now()+'.pdf'));
        uploadBlobs.push(blob);
      }
    });
  }

  if (uploadBlobs.length){
    let version = Number(values[idx('Versión_Actual')]) || 1;
    const nextVersion = version + 1;
    const itNext=bocFolder.getFoldersByName('v'+nextVersion);
    const vFolder = itNext.hasNext()? itNext.next() : bocFolder.createFolder('v'+nextVersion);
    uploadBlobs.forEach(b=>vFolder.createFile(b));

    let urls=[]; const files=vFolder.getFiles();
    while(files.hasNext()){ const f=files.next(); if (f.getMimeType()===MimeType.PDF) urls.push('https://drive.google.com/uc?export=download&id='+f.getId()); }
    values[idx('Versión_Actual')] = nextVersion;
    values[idx('PDFs_Vigentes_URLs')] = urls.join(' | ');
    values[idx('Estado')] = ESTADOS.REV_EJEC;
    const hist = values[idx('Historial_JSON')]?JSON.parse(values[idx('Historial_JSON')]):[];
    hist.push({ ts:new Date().toISOString(), ev:'PROD_PUBLICA', version: nextVersion, pdfs: urls });
    values[idx('Historial_JSON')] = JSON.stringify(hist);
  }

  getSheet_().getRange(row.row,1,1,header.length).setValues([values]);

  if (enviarA === 'cliente'){
    sendToCliente_(values, header);
    return HtmlService.createHtmlOutput('<p>✉️ Publicado y enviado al cliente. Ejecutivo copiado.</p><p><a href="'+values[idx('Link_Portal_Cliente')]+'" target="_blank">Abrir portal</a></p>');
  } else {
    sendToEjecutivo_(values, header);
    return HtmlService.createHtmlOutput('<p>✅ Publicado y notificado al Ejecutivo.</p><p><a href="'+values[idx('Link_Portal_Cliente')]+'" target="_blank">Abrir portal</a></p>');
  }
}

function doPostClient_(e){
  const token = (e.parameter.token||'').trim();
  const pedidoNro = (e.parameter.pedido||'').trim();
  const accion = (e.parameter.accion||'').trim(); // aprobar | cambios
  const comentarios = (e.parameter.comentarios||'').trim();
  if (!token || !pedidoNro || !accion) return HtmlService.createHtmlOutput('Solicitud incompleta.');

  const sh = getSheet_(); const data = sh.getDataRange().getValues(); const header=data[0]; const idx=(n)=>header.indexOf(n);
  let row=-1, values=null;
  for (let i=1;i<data.length;i++){
    if (String(data[i][idx('Nro_Pedido')])===pedidoNro && String(data[i][idx('Token_Acceso')])===token){ row=i+1; values=data[i]; break; }
  }
  if (row===-1) return HtmlService.createHtmlOutput('No autorizado.');

  const aprFolder = DriveApp.getFolderById(values[idx('Carpeta_Aprobaciones_FolderId')]);

  let fileUrl='';
  if (e.files && e.files.adjunto && e.files.adjunto.length !== 0){
    const f=e.files.adjunto;
    const blob = Utilities.newBlob(f.data, f.mimeType||MimeType.PDF, f.filename||('adjunto_'+Date.now()+'.pdf'));
    const saved = aprFolder.createFile(blob);
    fileUrl = 'https://drive.google.com/uc?export=download&id='+saved.getId();
  }

  if (accion==='aprobar'){ values[idx('Estado')] = ESTADOS.APROBADO; values[idx('Respuesta')]='Aprobado'; }
  else { values[idx('Estado')] = ESTADOS.CAMBIOS; values[idx('Respuesta')]='Cambios'; }
  values[idx('Comentarios_Cliente')] = comentarios;
  values[idx('Archivo_Cliente_URL')] = fileUrl || values[idx('Archivo_Cliente_URL')];
  values[idx('Fecha_Respuesta_Cliente')] = new Date();

  const hist = values[idx('Historial_JSON')]?JSON.parse(values[idx('Historial_JSON')]):[];
  hist.push({ ts:new Date().toISOString(), ev:(accion==='aprobar'?'APROBADO':'CAMBIOS'), comentarios, adjunto:fileUrl });
  values[idx('Historial_JSON')] = JSON.stringify(hist);

  sh.getRange(row,1,1,header.length).setValues([values]);

  const items = JSON.parse(values[idx('Items_JSON')]||'[]');
  const tabla = buildItemsTableHtml_(items);
  const ejEmail = resolveEjecutivoEmail_(values[idx('Ejecutivo')]);

  const empresa = values[idx('Empresa')];
  const correo = values[idx('Correo')];
  const asunto = `Pedido ${pedidoNro} - ${accion==='aprobar'?'APROBADO':'CAMBIOS SOLICITADOS'} (${empresa})`;
  const cuerpo = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#111">
      <p>Pedido: <b>${escapeHtml_(pedidoNro)}</b><br>
      Empresa: <b>${escapeHtml_(empresa||'')}</b><br>
      Acción: <b>${accion==='aprobar'?'Aprobado':'Cambios'}</b><br>
      Comentarios cliente: ${escapeHtml_(comentarios||'(sin comentarios)')}<br>
      ${fileUrl ? ('Adjunto cliente: '+escapeHtml_(fileUrl)) : ''}<br>
      Fecha: ${escapeHtml_(new Date().toString())}</p>
      ${tabla}
    </div>`;

  try{
    MailApp.sendEmail({ to: CORREO_PRODUCCION, cc: ejEmail, subject: asunto, htmlBody: cuerpo });
    if (correo){
      MailApp.sendEmail({
        to: correo,
        subject: `Hemos recibido tu ${accion==='aprobar'?'aprobación':'solicitud de cambios'} — Merchi`,
        htmlBody: `¡Gracias! Hemos registrado tu respuesta para el pedido <b>${pedidoNro}</b>.<br><br>Pronto te contactaremos.`
      });
    }
  }catch(_){}

  return HtmlService.createHtmlOutput('<p>✅ ¡Tu respuesta fue registrada! Ya puedes cerrar esta ventana.</p>');
}

/***** Atajos de menú *****/
function openProdPanel_(){
  const html = HtmlService.createHtmlOutput(`<p>Panel de Producción: <a href="${WEBAPP_EXEC_URL}?r=prod" target="_blank">${WEBAPP_EXEC_URL}?r=prod</a></p>`);
  SpreadsheetApp.getUi().showModalDialog(html, 'Merchi · Producción');
}
function uiDiagnosticoLink_(){
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Diagnóstico', 'Orden_ID a revisar:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton()!==ui.Button.OK) return;
  const nro = resp.getResponseText().trim();
  const row = ensurePedidoRow_(nro);
  const header=row.header, values=row.values, idx=(n)=>header.indexOf(n);
  ui.alert('OK', `Link: ${values[idx('Link_Portal_Cliente')]}\nToken: ${values[idx('Token_Acceso')]}`, ui.ButtonSet.OK);
}
