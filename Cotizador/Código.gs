function doGet() {
  return HtmlService.createHtmlOutputFromFile('formulario')
    .setTitle('Registro de Clientes | Merchi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Orden esperado de columnas (fila 1):
 *  1  Fecha
 *  2  ID_Registro
 *  3  Tipo_Doc
 *  4  Empresa
 *  5  Giro
 *  6  ID_Trib_Tipo
 *  7  RUT_CL
 *  8  ID_Trib_Valor
 *  9  País_Fiscal
 * 10  Fact_Calle
 * 11  Fact_Dpto
 * 12  Fact_Comuna
 * 13  Fact_Región
 * 14  Contacto_Nombre
 * 15  Contacto_Cargo
 * 16  Tel_Completo
 * 17  Correo
 * 18  Entrega_Calle
 * 19  Entrega_Dpto
 * 20  Entrega_Comuna
 * 21  Entrega_Región
 * 22  Entrega_Igual_Fact   ("Sí"/"No")
 */
function saveLead(data) {
  const SHEET_ID  = '1o9q9cwgx4f1iV87yrkdI4a5JR8El_coZa24VsKWMQhs';
  const SHEET_TAB = 'Datos Cotizaciones';
  const TZ        = 'America/Santiago';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_TAB);
    if (!sheet) throw new Error('No existe la pestaña "' + SHEET_TAB + '"');

    // === Generar ID corto ===
    const id = generarIdCorto_();

    // === Helpers ===
    const clean = v => (v || '').toString().trim();
    const normPhone = v => (v || '').toString().trim().replace(/[\s\-.]/g, '');
    const truthy = v => (v === true || v === 'true' || v === 'on' || v === '1');

    // === Fecha solo día (medianoche) ===
    const now = new Date();
    const fechaSolo = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    // === Leer payload ===
    const tipoDoc        = clean(data.tipo_doc);
    const empresa        = clean(data.empresa);
    const giro           = clean(data.giro);

    const idTipo         = clean(data.id_tipo) || 'RUT_CL'; // RUT_CL | EXTRANJERO
    const rutCL          = normalizeRut_(clean(data.rut));
    const paisFiscal     = clean(data.pais_fiscal);
    const idTribValor    = clean(data.id_trib_valor);

    // Facturación
    const factCalle      = clean(data.dirT_calle);
    const factDpto       = clean(data.dirT_dpto);
    const factComuna     = clean(data.dirT_comuna);
    const factRegion     = clean(data.dirT_region);

    // Contacto
    const contactoNombre = clean(data.contacto_nombre);
    const contactoCargo  = clean(data.contacto_cargo);
    const telCompleto    = normPhone(data.telefono_full);
    const correo         = clean(data.correo).toLowerCase();

    // Entrega
    const entregaIgual   = truthy(data.dirT_same);
    let   entCalle       = clean(data.direccionEntrega);
    let   entDpto        = clean(data.dpto);
    let   entComuna      = clean(data.comuna);
    let   entRegion      = clean(data.region);

    if (entregaIgual) {
      entCalle  = factCalle;
      entDpto   = factDpto;
      entComuna = factComuna;
      entRegion = factRegion;
    }

    // === Validaciones mínimas ===
    if (!tipoDoc)                  return { ok:false, message:'Selecciona el tipo de documento.' };
    if (!empresa)                  return { ok:false, message:'Ingresa la razón social.' };
    if (!giro)                     return { ok:false, message:'Ingresa el giro.' };

    if (idTipo === 'RUT_CL') {
      if (!isValidRut_(rutCL))     return { ok:false, message:'RUT chileno inválido.' };
    } else if (idTipo === 'EXTRANJERO') {
      if (!paisFiscal)             return { ok:false, message:'Selecciona el país fiscal.' };
      if (!/^[A-Z0-9\-\.\/]{3,30}$/i.test(idTribValor)) {
        return { ok:false, message:'ID/DNI extranjero inválido.' };
      }
    } else {
      return { ok:false, message:'Tipo de identificación inválido.' };
    }

    if (!factRegion || !factComuna)
      return { ok:false, message:'Completa región y comuna de facturación.' };

    if (!entRegion || !entComuna)
      return { ok:false, message:'Completa región y comuna de entrega.' };

    if (!contactoNombre)           return { ok:false, message:'Ingresa el nombre de contacto.' };
    if (!contactoCargo)            return { ok:false, message:'Ingresa el cargo del contacto.' };
    if (!isValidEmail_(correo))    return { ok:false, message:'Correo inválido.' };
    if (!isValidIntlPhoneFull_(telCompleto))
      return { ok:false, message:'Teléfono inválido. Usa formato +[código][número].' };

    // === Persistir fila ===
    const row = [
      fechaSolo,              // 1  Fecha
      id,                     // 2  ID_Registro
      tipoDoc,                // 3  Tipo_Doc
      empresa,                // 4  Empresa
      giro,                   // 5  Giro
      idTipo,                 // 6  ID_Trib_Tipo
      (idTipo === 'RUT_CL' ? rutCL : ''),              // 7  RUT_CL (normalizado)
      (idTipo === 'EXTRANJERO' ? idTribValor : ''),    // 8  ID_Trib_Valor
      (idTipo === 'EXTRANJERO' ? paisFiscal : ''),     // 9  País_Fiscal
      factCalle,              // 10 Fact_Calle
      factDpto,               // 11 Fact_Dpto
      factComuna,             // 12 Fact_Comuna
      factRegion,             // 13 Fact_Región
      contactoNombre,         // 14 Contacto_Nombre
      contactoCargo,          // 15 Contacto_Cargo
      telCompleto,            // 16 Tel_Completo
      correo,                 // 17 Correo
      entCalle,               // 18 Entrega_Calle
      entDpto,                // 19 Entrega_Dpto
      entComuna,              // 20 Entrega_Comuna
      entRegion,              // 21 Entrega_Región
      (entregaIgual ? 'Sí' : 'No') // 22 Entrega_Igual_Fact
    ];

    sheet.appendRow(row);
    sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('dd/MM/yyyy');

    // === Enviar correo ===
    const fechaStr = Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm");
    const payload = {
      id, tipoDoc, empresa, giro, idTipo,
      rutCL, idTribValor, paisFiscal,
      factCalle, factDpto, factComuna, factRegion,
      contactoNombre, contactoCargo, telCompleto, correo,
      entCalle, entDpto, entComuna, entRegion, entregaIgual
    };

    const subj = `Nuevo registro de cotización (${id}) — ${empresa}`;
    const { html, text } = buildEmailBodies_(payload, fechaStr);

    let emailSent = false, emailNote = '';
    try {
      const to = 'sergio@merchi.cl';
      const aliases = GmailApp.getAliases();
      const desiredFrom = 'ordenes.merchi@gmail.com';
      const opts = {
        name: 'Merchi — Órdenes',
        htmlBody: html,
        replyTo: correo // responder al cliente
      };
      if (aliases && aliases.indexOf(desiredFrom) !== -1) {
        opts.from = desiredFrom; // usa alias si está autorizado
      } else {
        emailNote = 'Alias no autorizado: se envió desde la cuenta actual con reply-to del cliente.';
      }
      GmailApp.sendEmail(to, subj, text, opts);
      emailSent = true;
    } catch (e) {
      emailNote = 'No se pudo enviar el correo: ' + e.message;
    }

    const baseMsg = 'Registro completado con éxito.';
    const msg = emailSent ? (baseMsg + ' Correo enviado.') : (baseMsg + (emailNote ? ' ' + emailNote : ''));
    return { ok:true, id:id, message: msg };

  } catch (err) {
    return { ok:false, message:'Error al guardar: ' + err.message };
  }
}

/* === Generador de ID corto === */
function generarIdCorto_() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let out = '';
  for (let i = 0; i < 5; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

/* === Normaliza RUT: sin puntos/espacios y sin punto final === */
function normalizeRut_(rut) {
  rut = (rut || '').toString().trim().toUpperCase();
  rut = rut.replace(/\./g, '');    // elimina puntos
  rut = rut.replace(/\s+/g, '');   // elimina espacios
  rut = rut.replace(/\.+$/, '');   // elimina punto final si lo hay
  return rut;
}

/* === Email HTML/Text === */
function buildEmailBodies_(p, fechaStr) {
  const idStr   = (p.idTipo === 'RUT_CL') ? (p.rutCL || '') : (p.idTribValor || '');
  const idLabel = (p.idTipo === 'RUT_CL') ? 'RUT' : 'ID extranjero';

  const block = (label, value) => `
    <tr>
      <td style="padding:8px 10px;border:1px solid #eee;background:#fafafa;width:220px"><strong>${label}</strong></td>
      <td style="padding:8px 10px;border:1px solid #eee">${value || '-'}</td>
    </tr>`;

  const html = `
  <div style="font-family:Arial,system-ui,sans-serif;font-size:14px;color:#111">
    <h2 style="margin:0 0 6px">Nuevo registro de cotización</h2>
    <p style="margin:0 0 12px;color:#555">Fecha: ${fechaStr} — ID: <strong>${p.id}</strong></p>
    <table style="border-collapse:collapse;min-width:560px">
      ${block('Documento', p.tipoDoc)}
      ${block('Razón social', p.empresa)}
      ${block('Giro', p.giro)}
      ${block(idLabel, idStr)}
      ${p.idTipo === 'EXTRANJERO' ? block('País fiscal', p.paisFiscal) : ''}
      ${block('Facturación', [p.factCalle, p.factComuna, p.factRegion].filter(Boolean).join(', '))}
      ${block('Entrega', [p.entCalle, p.entComuna, p.entRegion].filter(Boolean).join(', '))}
      ${block('Entrega igual a facturación', p.entregaIgual ? 'Sí' : 'No')}
      ${block('Contacto', p.contactoNombre)} 
      ${block('Cargo', p.contactoCargo)}
      ${block('Teléfono', p.telCompleto)}
      ${block('Correo', p.correo)}
    </table>
  </div>`.replace(/\n\s+/g, '\n');

  const text = [
    `Nuevo registro de cotización`,
    `Fecha: ${fechaStr}`,
    `ID: ${p.id}`,
    `Documento: ${p.tipoDoc}`,
    `Razón social: ${p.empresa}`,
    `Giro: ${p.giro}`,
    `${idLabel}: ${idStr}`,
    p.idTipo === 'EXTRANJERO' ? `País fiscal: ${p.paisFiscal}` : '',
    `Facturación: ${[p.factCalle, p.factComuna, p.factRegion].filter(Boolean).join(', ')}`,
    `Entrega: ${[p.entCalle, p.entComuna, p.entRegion].filter(Boolean).join(', ')}`,
    `Entrega igual a facturación: ${p.entregaIgual ? 'Sí' : 'No'}`,
    `Contacto: ${p.contactoNombre}`,
    `Cargo: ${p.contactoCargo}`,
    `Teléfono: ${p.telCompleto}`,
    `Correo: ${p.correo}`
  ].filter(Boolean).join('\n');

  return { html, text };
}

/* === Validadores === */
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email || '');
}
function isValidIntlPhoneFull_(v) {
  if (!v) return false;
  v = v.toString().trim().replace(/[\s\-.]/g, '');
  return /^\+\d{8,15}$/.test(v); // E.164 relajado
}
function isValidRut_(rut) {
  rut = (rut || '').replace(/\./g, '').toUpperCase();
  if (!/^\d{1,9}-[\dK]$/.test(rut)) return false;
  const [num, dv] = rut.split('-');
  let s = 0, m = 2;
  for (let i = num.length - 1; i >= 0; i--) {
    s += parseInt(num[i], 10) * m;
    m = (m === 7) ? 2 : m + 1;
  }
  const r = 11 - (s % 11);
  const dvCalc = (r === 11) ? '0' : (r === 10 ? 'K' : String(r));
  return dvCalc === dv;
}


