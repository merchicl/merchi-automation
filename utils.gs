/** Utils compartidos — Merchi (Apps Script) **/
function sanitizeRut(rut){ if(!rut) return ''; rut = String(rut).trim().replace(/[.\s]+$/g,'').replace(/\u200B/g,''); return rut.replace(/k$/,'K'); }
function normalizeId(s){ return (s||'').toString().replace(/\u200B/g,'').trim(); }
