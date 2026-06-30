// ══════════════════════════════════════════════════════════════
//  APPS SCRIPT — Auditoría Interna de Cocina · Temple Colombia
//  Pega este código en: script.google.com → Nuevo proyecto
//  Luego: Implementar → Nueva implementación → Aplicación web
//    · Ejecutar como: Yo
//    · Quién tiene acceso: Cualquier usuario
// ══════════════════════════════════════════════════════════════

const SECRET_TOKEN = 'temple2026';
const SHEET_ID     = '1RZdgYL5HCUJ8pVnSo1wglixwR-Inaydkg0BO0FWDIWM';
const SHEET_NAME   = 'Auditorías';

const HEADERS = [
  'ID','PDV','Fecha','Hora','Auditor','Manager','Jefe Cocina',
  'Coordinador Zona','Persona que Recibe','Tipo PDV','Modalidad',
  'Puntaje Total','% Cumplimiento','Estado',
  'Hallazgos','Retroalimentación','Novedades',
  'Ítems (JSON)','Guardado En'
];

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, HEADERS.length);
    hr.setBackground('#1b4332');
    hr.setFontColor('white');
    hr.setFontWeight('bold');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(5, 130);
    sheet.setColumnWidth(15, 250);
    sheet.setColumnWidth(16, 250);
    sheet.setColumnWidth(17, 250);
  }
  return sheet;
}

function doPost(e) {
  try {
    if ((e.parameter.token || '') !== SECRET_TOKEN)
      return ok({error: 'Token inválido'});

    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    // Evitar duplicados
    const ids = sheet.getDataRange().getValues().slice(1).map(r => String(r[0]));
    if (ids.includes(String(data.id)))
      return ok({status: 'exists'});

    const pctNum = typeof data.pct === 'number' ? data.pct : parseFloat(data.pct) || 0;

    sheet.appendRow([
      data.id,
      data.pdv            || '',
      data.fecha          || '',
      data.hora           || '',
      data.auditor        || '',
      data.admin          || '',
      data.jefeCocina     || '',
      data.coordinadorZona|| '',
      data.receptor       || '',
      data.tipoPdv        || '',
      data.modalidad      || '',
      Number(data.puntajeTotal) || 0,
      pctNum.toFixed(1) + '%',
      data.estado         || '',
      data.hallazgo       || '',
      data.retroalimentacion || '',
      data.novedades      || '',
      JSON.stringify(data.items || []),
      data.savedAt        || new Date().toISOString()
    ]);

    return ok({status: 'saved'});
  } catch(err) {
    return ok({error: err.message});
  }
}

function doGet(e) {
  try {
    if ((e.parameter.token || '') !== SECRET_TOKEN)
      return ok({error: 'Token inválido'});

    const sheet = getSheet();
    const rows  = sheet.getDataRange().getValues().slice(1); // sin encabezado

    const records = rows
      .filter(r => r[0] && String(r[0]) !== 'ID')
      .map(r => {
        let items = [];
        try { items = JSON.parse(String(r[17] || '[]')); } catch(_) {}
        const pctStr = String(r[12]).replace('%','');
        return {
          id:                String(r[0]),
          pdv:               String(r[1]),
          fecha:             String(r[2]),
          hora:              String(r[3]),
          auditor:           String(r[4]),
          admin:             String(r[5]),
          jefeCocina:        String(r[6]),
          coordinadorZona:   String(r[7]),
          receptor:          String(r[8]),
          tipoPdv:           String(r[9]),
          modalidad:         String(r[10]),
          puntajeTotal:      Number(r[11]) || 0,
          pct:               parseFloat(pctStr) || 0,
          estado:            String(r[13]),
          hallazgo:          String(r[14]),
          retroalimentacion: String(r[15]),
          novedades:         String(r[16]),
          items:             items,
          fotos:             [],
          savedAt:           String(r[18]),
          _remoto:           true
        };
      });

    return ok({records, total: records.length});
  } catch(err) {
    return ok({error: err.message});
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
