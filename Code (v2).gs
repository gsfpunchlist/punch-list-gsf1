// ============================================================
//  PUNCH LIST GSF1 — Google Apps Script Backend v2.0
//  Proyecto: GSF1 CCPP - TSK
//  Autor: Jorge Perez
//  Fecha: 2026-04-10
// ============================================================
//  CONFIGURACION — editar antes de desplegar
// ============================================================

const CONFIG = {
  SHEET_ID:       '10ua5g0SerKUp6uPjTvK9xdz0AiWvOYbmI-AlgDbTqy0',
  FOLDER_ID:      '191HZC1Z73lJijuG0S-y9h2vOhO5kuE4s',
  ADMIN_PASSWORD: 'GSF1admin2026!',
  SHEET_NAME:     'PunchList',
  DASH_NAME:      'Dashboard'
};

// ── COLUMNAS ESPERADAS (orden canónico v2) ───────────────────
const COLS = [
  'ID','Fecha','KKS/Tag','Sistema KKS','Descripción Sistema',
  'Ubicación','Área','Descripción','Categoría','Reportado por',
  'Foto URL','Estatus','Cerrado por','Fecha cierre','Timestamp'
];

// ── KKS SYSTEM LOOKUP — GSF-1 CCPP (TSKI-002777) ────────────
const KKS = {
  'AE':'110-150 kV Systems','AEA':'Breakers','AEB':'Disconnectors',
  'AEC':'Current Transformers','AED':'Voltage Transformers','AET':'Main Transformers','AEW':'Lightning Protection kV',
  'AN':'<1 kV Systems','ANE':'LV Switchgear 480/208 Vac','ANF':'LV Switchgear UPS','ANK':'DC Switchgear 125/110 Vdc','ANQ':'DC Switchgear 24 Vdc',
  'AP':'Control Consoles','APA':'SCADA Panel',
  'AQ':'Measuring & Metering','AQA':'Tariff Metering Panel',
  'AR':'Protection Equipment','ARA':'Protection & Control Panel',
  'AS':'Decentralized Panels','ASQ':'Metering Junction Boxes','ASR':'Protection Junction Boxes',
  'AT':'Transformer Equipment','ATA':'Auxiliary Transformer',
  'AY':'Communication Equipment','AYA':'Telephone (PABX)','AYB':'Control Console Telephone','AYE':'Fire Alarm System','AYG':'Telecommunication Panel',
  'BA':'Power Transmission','BAA':'Generator Busbars','BAC':'Generation Circuit Breaker','BAT':'Generator Transformers','BAW':'Earthing & Lightning Protection','BAY':'Control & Protection',
  'BB':'MV Distribution — Normal','BBA':'MV Switchgear','BBT':'Auxiliary Transformer','BBY':'Control & Protection',
  'BD':'MV Emergency Distribution','BDA':'MV Switchgear','BDT':'MV Diesel Generator','BDY':'Control & Protection',
  'BF':'LV Main Distribution — Normal','BFA':'LV Main Dist. Board','BFB':'LV Main Dist. Board','BFC':'LV Main Dist. Board','BFD':'LV Main Dist. Board','BFT':'LV Main Dist. Transformers','BFY':'Control & Protection',
  'BJ':'LV Subdistribution — Normal','BJA':'MCC','BJT':'LV Aux. Power Transformers','BJY':'Control & Protection',
  'BL':'LV Subdistribution — General','BLA':'LV Subdist. Board','BLB':'LV Subdist. Board','BLC':'LV Subdist. Board','BLD':'LV Subdist. Board','BLE':'LV Subdist. Board','BLT':'LV Aux. Power Transformers','BLY':'Control & Protection',
  'BM':'LV Distribution — Diesel Emergency','BMA':'LV Emergency Dist. Board','BMB':'LV Emergency Dist. Board','BME':'LV Emergency Dist. Board','BMY':'Control & Protection',
  'BN':'LV Subdistribution — Diesel Emergency','BNA':'LV Emergency Subdist.','BNB':'LV Emergency Subdist.','BNC':'LV Emergency Subdist.','BND':'LV Emergency Subdist.',
  'BR':'LV Distribution — UPS','BRA':'UPS Dist. Board','BRB':'UPS Dist. Board','BRU':'UPS Inverter','BRT':'Isolation Transformer',
  'BT':'Battery Systems','BTA':'DC Batteries','BTB':'DC Batteries','BTL':'Battery Charger',
  'BU':'DC Distribution — Normal','BUA':'DC Dist. Board','BUB':'DC Dist. Board',
  'BY':'Control & Protection','BYA':'Generator & Transformer Protection','BYB':'Mains Coupling Relay Cabinet',
  'CA':'Protective Interlocks','CAA':'BPS Cabinets',
  'CB':'Functional Group Control','CBA':'ST Generator Control Panel','CBP':'Synchronization Cabinets',
  'CC':'Binary Signal Conditioning','CCA':'Binary Signal Cabinets',
  'CD':'Drive Control Interface','CDA':'Drive Control Cabinets',
  'CE':'Annunciation','CEA':'Annunciation Cabinets','CEJ':'Fault Recording',
  'CF':'Measurement & Recording','CFA':'Measurement Cabinets',
  'CH':'Protection','CHA':'Generator & Transformer Protection Cabinet',
  'CJ':'Unit Coordination (DCS)','CJA':'Unit Control System','CJF':'Boiler Control System','CJJ':'I&C — Steam Turbine','CJP':'I&C — Gas Turbine','CJU':'I&C — Main Machinery',
  'CK':'Process Computer','CKA':'Process Computer System',
  'CR':'Process Control System','CRJ':'Automation (Fail-Safe)','CRK':'Automation (High Availability)','CRR':'Terminal Bus','CRS':'Plant Bus','CRT':'Field Bus','CRU':'HMI Operation & Monitoring',
  'CW':'Control Rooms','CWA':'Main Control Consoles',
  'CX':'Local Control Stations','CXA':'Local Control Stations',
  'CY':'Communication & Information','CYA':'Telephone','CYE':'Fire Alarm','CYG':'Remote Control','CYQ':'Gas Detection',
  'EK':'Gaseous Fuel Supply','EKA':'Receiving Equipment / Pipeline','EKB':'Scrubber','EKC':'Heating System','EKD':'Reducing Station','EKE':'Cleaning System','EKF':'Storage','EKG':'Fuel Piping','EKH':'Main Pressure Boosting','EKN':'Purging System','EKY':'Control & Protection',
  'EG':'Liquid Fuel Supply','EGA':'Receiving Equipment','EGB':'Tank Farm','EGC':'Pump System','EGD':'Piping System','EGY':'Control & Protection',
  'GA':'Raw Water Supply','GAA':'Extraction & Cleaning','GAC':'Piping & Channel','GAD':'Storage','GAF':'Pump System',
  'GB':'Water Treatment — Carbonate Hardness','GBB':'Filtering','GBN':'Chemicals Supply','GBY':'Control & Protection',
  'GC':'Water Treatment — Demineralization','GCB':'Filtering','GCF':'Ion Exchange / RO','GCN':'Chemicals Supply','GCY':'Control & Protection',
  'HN':'Exhaust Gas System','HNA':'Exhaust Gas System','HNE':'Smokestack / Chimney',
  'LA':'Feedwater System','LAA':'Feedwater Tank / Deaerator','LAB':'Feedwater Piping','LAC':'Feedwater Pump','LAD':'HP Feedwater Heating','LAE':'HP Desuperheating Spray','LAF':'IP Desuperheating Spray','LAV':'Lube Oil Supply',
  'LB':'Steam System','LBA':'Main Steam Piping','LBB':'Hot Reheat Steam','LBC':'Cold Reheat Steam','LBD':'Extraction Piping','LBG':'Auxiliary Steam Piping','LBH':'Start-up / Shutdown Steam',
  'LC':'Condensate System','LCA':'Main Condensate Piping','LCB':'Main Condensate Pump','LCE':'Condensate Desuperheat Spray','LCH':'HP Heater Drains','LCJ':'LP Heater Drains','LCY':'Control & Protection',
  'MA':'Steam Turbine Plant','MAA':'HP Turbine','MAB':'IP Turbine','MAC':'LP Turbine','MAD':'Bearings','MAG':'Condensing System','MAJ':'Air Removal','MAK':'Turning Gear','MAL':'Drains & Vents','MAN':'Turbine Bypass Station','MAP':'LP Turbine Bypass','MAV':'Lube Oil System','MAX':'Non-Electric Control Equipment','MAY':'Electric Control & Protection',
  'MB':'Gas Turbine Plant','MBA':'Turbine & Compressor Rotor','MBD':'Bearings','MBH':'Cooling & Sealing Gas','MBJ':'Start-up Unit','MBK':'Transmission Gear','MBL':'Intake Air System','MBP':'Fuel Supply (Gaseous)','MBR':'Exhaust Gas System','MBV':'Lube Oil System','MBX':'Non-Electric Control Equipment','MBY':'Electric Control & Protection',
  'MJ':'Diesel Engine Plant','MJA':'Engine','MJV':'Lube Oil','MJY':'Control & Protection',
  'MK':'Generator Plant','MKA':'Generator (Stator/Rotor/Cooling)','MKB':'Generator Exciter','MKC':'Generator Exciter','MKD':'Bearings','MKF':'Stator/Rotor Liquid Cooling','MKG':'H2 Cooling System','MKH':'H2/N2/CO2 Cooling','MKX':'Fluid Supply — Control Equipment','MKY':'Control & Protection',
  'ML':'Electro-Motive Plant','MLA':'Motor Frame / Motor-Generator',
  'PA':'Main Cooling Water','PAA':'Extraction & Cleaning','PAB':'CW Piping & Culvert','PAC':'CW Pump','PAD':'Recirculating Cooling','PAE':'Cooling Tower Pump','PAH':'Condenser Cleaning',
  'PB':'Cooling Water Treatment','PBN':'Chemicals Supply',
  'PG':'Closed Cooling Water','PGB':'Closed CW (Combined Cycle)','PGD':'Closed CW (Gas Turbine)',
  'QC':'Central Chemicals Supply','QCA':'O2 Scavenger Injection','QCB':'Neutralizer Injection','QCC':'Phosphate Injection','QCD':'Corrosion Inhibitor','QCE':'Other Chemicals','QCF':'Ammonia Injection',
  'QE':'Service Compressed Air','QEA':'Service Air Generation','QEB':'Service Air Distribution',
  'QF':'Instrument Air','QFA':'Compressed Air Generation','QFB':'Instrument Air Distribution','QFC':'Service Air (Power Island)','QFD':'HP Compressed Air',
  'QH':'Auxiliary Steam Generating','QHA':'Pressure System','QHY':'Control & Protection',
  'QJ':'Central Gas Supply','QJA':'H2 Storage & Distribution','QJB':'CO2 Storage & Distribution','QJN':'N2 Storage & Distribution',
  'QU':'Sampling Systems','QUA':'Sampling Drains','QUB':'Steam Sampling','QUC':'Water Sampling',
  'SA':'HVAC — Conventional Area','SAA':'Buildings Ventilation','SAK':'Air Conditioning','SAM':'Turbine Hall Ventilation','SAU':'Buildings Heating',
  'SG':'Fire Protection','SGA':'Fire Water System','SGC':'Spray Deluge','SGE':'Sprinkler Systems','SGF':'Foam Fire-Fighting','SGJ':'CO2 Fire-Fighting','SGY':'Control & Protection',
  'SM':'Cranes & Hoisting','SMA':'Crane System',
  'ST':'Workshop & Stores','STA':'Workshop Equipment','STE':'Stores & Filling Station','STG':'Laboratory Equipment',
  'UA':'Structures — Grid & Distribution','UAA':'Switchyard Structure','UAG':'Structure for Transformers',
  'UB':'Structures — Power Transmission','UBA':'Switchgear Building',
  'UM':'Structures — Main Machine Set','UMA':'Steam Turbine Building',
  'UL':'Structures — Steam Cycle','ULA':'Feedwater Pump House',
  'US':'Area / Building','USG':'Fire Pump House','UST':'Workshop','USV':'Laboratory Building',
  'UT':'Structures — Auxiliary Systems','UTF':'Compressed Air Building',
  'UY':'General Service Structures','UYA':'Office & Staff Amenities'
};

const CATEGORIES = [
  'Defecto de instalación',
  'Documentación pendiente',
  'NCR (Non Conformance Report)',
  'Punch de comisionamiento',
  'Observación de seguridad',
  'Pendiente de materiales',
  'Prueba de aceptación',
  'Otros'
];

// ════════════════════════════════════════════════════════════
//  ROUTER
// ════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'submit':          return submitItem(data);
      case 'close':           return closeItem(data);
      case 'reopen':          return reopenItem(data);
      case 'getAll':          return getAllItems(data);
      case 'verifyAdmin':     return verifyAdmin(data);
      case 'updateCategory':  return updateCategory(data);
      case 'getConfig':       return getConfig();
      default:                return jsonResponse({ error: 'Acción desconocida' });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doGet() {
  return jsonResponse({ status: 'ok', version: '2.0', project: 'GSF1 CCPP - TSK' });
}

// ════════════════════════════════════════════════════════════
//  SUBMIT ITEM
// ════════════════════════════════════════════════════════════
function submitItem(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet    = getSheet();
    const colMap   = getColMap(sheet);
    const lastRow  = sheet.getLastRow();
    const id       = 'PL-' + String(lastRow).padStart(3, '0');
    const kksCode  = extractKKS(data.kks);
    const kksDesc  = kksCode ? (KKS[kksCode] || kksCode) : '';
    let   photoUrl = '';

    // Upload photo to Drive
    if (data.photo && data.photo.length > 100) {
      try {
        const folder   = DriveApp.getFolderById(CONFIG.FOLDER_ID);
        const b64      = data.photo.split(',')[1];
        const mime     = data.photo.split(';')[0].split(':')[1] || 'image/jpeg';
        const ext      = mime.includes('png') ? 'png' : 'jpg';
        const name     = id + '_' + data.fecha + '_' + sanitize(data.reportedBy) + '.' + ext;
        const blob     = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
        const file     = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl       = file.getUrl();
      } catch (pe) { Logger.log('Photo error: ' + pe.message); }
    }

    // Build row aligned to column map
    const row = buildRow(colMap, {
      'ID':                   id,
      'Fecha':                data.fecha        || '',
      'KKS/Tag':              data.kks          || '',
      'Sistema KKS':          kksCode           || '',
      'Descripción Sistema':  kksDesc           || '',
      'Ubicación':            data.ubicacion    || '',
      'Área':                 data.area         || '',
      'Descripción':          data.desc         || '',
      'Categoría':            data.categoria    || '',
      'Reportado por':        data.reportedBy   || '',
      'Foto URL':             photoUrl,
      'Estatus':              'Abierto',
      'Cerrado por':          '',
      'Fecha cierre':         '',
      'Timestamp':            new Date().toISOString()
    });

    sheet.appendRow(row);
    const newRow = sheet.getLastRow();
    const statusCol = colMap['Estatus'] || 12;
    sheet.getRange(newRow, statusCol)
      .setBackground('#FCEBEB').setFontColor('#791F1F').setFontWeight('bold');

    refreshDashboard(sheet);
    return jsonResponse({ success: true, id, photoUrl, kksCode, kksDesc });

  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════
//  CLOSE / REOPEN
// ════════════════════════════════════════════════════════════
function closeItem(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD) return jsonResponse({ error: 'Contraseña incorrecta' });
  const sheet  = getSheet();
  const colMap = getColMap(sheet);
  const values = sheet.getDataRange().getValues();
  const iID     = (colMap['ID']          || 1) - 1;
  const iStatus = (colMap['Estatus']     || 12) - 1;
  const iCloser = (colMap['Cerrado por'] || 13) - 1;
  const iDate   = (colMap['Fecha cierre']|| 14) - 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i][iID] === data.id) {
      const r = i + 1;
      sheet.getRange(r, iStatus + 1).setValue('Cerrado').setBackground('#C0DD97').setFontColor('#27500A').setFontWeight('bold');
      sheet.getRange(r, iCloser + 1).setValue(data.closedBy || 'Admin');
      sheet.getRange(r, iDate   + 1).setValue(new Date().toLocaleDateString('es-DO'));
      refreshDashboard(sheet);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'Ítem no encontrado' });
}

function reopenItem(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD) return jsonResponse({ error: 'Contraseña incorrecta' });
  const sheet  = getSheet();
  const colMap = getColMap(sheet);
  const values = sheet.getDataRange().getValues();
  const iID     = (colMap['ID']          || 1) - 1;
  const iStatus = (colMap['Estatus']     || 12) - 1;
  const iCloser = (colMap['Cerrado por'] || 13) - 1;
  const iDate   = (colMap['Fecha cierre']|| 14) - 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i][iID] === data.id) {
      const r = i + 1;
      sheet.getRange(r, iStatus + 1).setValue('Abierto').setBackground('#FCEBEB').setFontColor('#791F1F').setFontWeight('bold');
      sheet.getRange(r, iCloser + 1).setValue('');
      sheet.getRange(r, iDate   + 1).setValue('');
      refreshDashboard(sheet);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'Ítem no encontrado' });
}

// ════════════════════════════════════════════════════════════
//  UPDATE CATEGORY (admin only)
// ════════════════════════════════════════════════════════════
function updateCategory(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD) return jsonResponse({ error: 'Contraseña incorrecta' });
  const sheet  = getSheet();
  const colMap = getColMap(sheet);
  const values = sheet.getDataRange().getValues();
  const iID  = (colMap['ID']        || 1) - 1;
  const iCat = (colMap['Categoría'] || 9) - 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i][iID] === data.id) {
      sheet.getRange(i + 1, iCat + 1).setValue(data.categoria);
      refreshDashboard(sheet);
      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'Ítem no encontrado' });
}

// ════════════════════════════════════════════════════════════
//  GET ALL (admin)
// ════════════════════════════════════════════════════════════
function getAllItems(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD) return jsonResponse({ error: 'No autorizado' });
  const sheet  = getSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return jsonResponse({ items: [] });
  const headers = values[0];
  const items   = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return jsonResponse({ items });
}

function verifyAdmin(data) {
  return jsonResponse({ valid: data.password === CONFIG.ADMIN_PASSWORD });
}

function getConfig() {
  return jsonResponse({ categories: CATEGORIES });
}

// ════════════════════════════════════════════════════════════
//  MIGRATE & BACKFILL (run once manually from editor)
// ════════════════════════════════════════════════════════════
function migrateAndRefresh() {
  const sheet   = getSheet();          // ensures headers exist & adds missing cols
  const colMap  = getColMap(sheet);
  const values  = sheet.getDataRange().getValues();

  const iKKS   = (colMap['KKS/Tag']           || 3)  - 1;
  const iSys   = (colMap['Sistema KKS']        || -1);
  const iDesc  = (colMap['Descripción Sistema']|| -1);
  const iStat  = (colMap['Estatus']            || 12) - 1;

  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    const raw = String(values[i][iKKS] || '');
    const hasSys = iSys > 0 && values[i][iSys - 1];
    if (raw && !hasSys) {
      const code = extractKKS(raw);
      const desc = code ? (KKS[code] || '') : '';
      if (code && iSys  > 0) sheet.getRange(i + 1, iSys).setValue(code);
      if (desc && iDesc > 0) sheet.getRange(i + 1, iDesc).setValue(desc);
      if (code) updated++;
    }
  }
  Logger.log('Filas actualizadas: ' + updated);
  refreshDashboard(sheet);
  Logger.log('Dashboard creado/actualizado correctamente.');
}

// ════════════════════════════════════════════════════════════
//  DASHBOARD / PIVOT TABLE
// ════════════════════════════════════════════════════════════
function refreshDashboard(sheet) {
  if (!sheet) sheet = getSheet();
  const ss      = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) return;

  // Get or create Dashboard tab
  let dash = ss.getSheetByName(CONFIG.DASH_NAME);
  if (!dash) {
    dash = ss.insertSheet(CONFIG.DASH_NAME);
    // Move Dashboard to second position
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(2);
  }
  dash.clearContents();
  dash.clearFormats();

  // Column map of source data
  const hdr    = allData[0];
  const col    = {};
  hdr.forEach((h, i) => { col[String(h).trim()] = i; });

  const iKKS   = col['KKS/Tag']            !== undefined ? col['KKS/Tag']            : 2;
  const iSys   = col['Sistema KKS']         !== undefined ? col['Sistema KKS']         : -1;
  const iSDesc = col['Descripción Sistema'] !== undefined ? col['Descripción Sistema'] : -1;
  const iArea  = col['Área']                !== undefined ? col['Área']                : 4;
  const iCat   = col['Categoría']           !== undefined ? col['Categoría']           : -1;
  const iStat  = col['Estatus']             !== undefined ? col['Estatus']             : 9;
  const iRep   = col['Reportado por']       !== undefined ? col['Reportado por']       : 6;

  const rows   = allData.slice(1);
  const total  = rows.length;
  const open   = rows.filter(r => r[iStat] === 'Abierto').length;
  const closed = rows.filter(r => r[iStat] === 'Cerrado').length;

  const G = '#0F6E56'; const W = '#FFFFFF';
  const RL = '#FCEBEB'; const GL = '#EAF3DE'; const GR = '#F5F5F5';

  // ── HEADER ──────────────────────────────────
  dash.getRange('A1:H1').merge()
    .setValue('PUNCH LIST DASHBOARD — GSF1 CCPP - TSK')
    .setBackground(G).setFontColor(W).setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  dash.getRange('A2:H2').merge()
    .setValue('Actualizado: ' + new Date().toLocaleString('es-DO'))
    .setFontColor('#888888').setFontSize(9).setHorizontalAlignment('right');

  // ── SUMMARY CARDS ────────────────────────────
  const cards = [['TOTAL', total, '#1a1a1a', GR], ['ABIERTOS', open, '#791F1F', RL], ['CERRADOS', closed, '#27500A', GL]];
  cards.forEach(([lbl, val, fc, bg], i) => {
    const c = i * 2 + 1;
    dash.getRange(4, c, 1, 2).merge().setValue(lbl).setBackground(bg).setFontColor(fc)
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
    dash.getRange(5, c, 1, 2).merge().setValue(val).setBackground(bg).setFontColor(fc)
      .setFontWeight('bold').setFontSize(26).setHorizontalAlignment('center');
  });

  let row = 7;

  // ── SECTION HELPER ───────────────────────────
  function sectionHeader(title, cols) {
    dash.getRange(row, 1, 1, cols).merge()
      .setValue(title).setBackground(G).setFontColor(W).setFontWeight('bold').setFontSize(10);
    row++;
  }
  function tableHeader(labels) {
    labels.forEach((lbl, i) => {
      dash.getRange(row, i + 1).setValue(lbl)
        .setBackground('#333333').setFontColor(W).setFontWeight('bold').setFontSize(9);
    });
    row++;
  }
  function tableRow(values, highlights) {
    const bg = row % 2 === 0 ? W : GR;
    values.forEach((v, i) => {
      const cell = dash.getRange(row, i + 1);
      cell.setValue(v).setFontSize(9);
      const h = highlights && highlights[i];
      if (h) cell.setBackground(h.bg).setFontColor(h.fc).setFontWeight('bold');
      else   cell.setBackground(bg);
    });
    row++;
  }

  // ── BY KKS SYSTEM ────────────────────────────
  sectionHeader('POR SISTEMA KKS', 5);
  tableHeader(['Código', 'Sistema', 'Abiertos', 'Cerrados', 'Total']);

  const byKKS = {};
  rows.forEach(r => {
    let code = iSys >= 0 ? String(r[iSys] || '').trim() : '';
    let desc = iSDesc >= 0 ? String(r[iSDesc] || '').trim() : '';
    if (!code) {
      code = extractKKS(String(r[iKKS] || ''));
      desc = code ? (KKS[code] || '') : '';
    }
    if (!code) { code = '(Sin KKS)'; desc = ''; }
    if (!byKKS[code]) byKKS[code] = { desc, open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') byKKS[code].open++;
    else byKKS[code].closed++;
  });

  Object.entries(byKKS)
    .sort((a, b) => (b[1].open + b[1].closed) - (a[1].open + a[1].closed))
    .forEach(([code, v]) => {
      const t = v.open + v.closed;
      tableRow(
        [code, v.desc, v.open, v.closed, t],
        [null, null,
          v.open   > 0 ? { bg: RL, fc: '#791F1F' } : null,
          v.closed > 0 ? { bg: GL, fc: '#27500A' } : null,
          null]
      );
    });

  row++;

  // ── BY AREA ──────────────────────────────────
  sectionHeader('POR ÁREA / DISCIPLINA', 5);
  tableHeader(['Área', 'Abiertos', 'Cerrados', 'Total', '% Cierre']);

  const byArea = {};
  rows.forEach(r => {
    const a = String(r[iArea] || '(Sin área)').trim() || '(Sin área)';
    if (!byArea[a]) byArea[a] = { open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') byArea[a].open++;
    else byArea[a].closed++;
  });

  Object.entries(byArea)
    .sort((a, b) => (b[1].open + b[1].closed) - (a[1].open + a[1].closed))
    .forEach(([area, v]) => {
      const t   = v.open + v.closed;
      const pct = t > 0 ? Math.round((v.closed / t) * 100) + '%' : '0%';
      tableRow(
        [area, v.open, v.closed, t, pct],
        [null,
          v.open   > 0 ? { bg: RL, fc: '#791F1F' } : null,
          v.closed > 0 ? { bg: GL, fc: '#27500A' } : null,
          null, null]
      );
    });

  row++;

  // ── BY CATEGORY ──────────────────────────────
  sectionHeader('POR CATEGORÍA', 5);
  tableHeader(['Categoría', 'Abiertos', 'Cerrados', 'Total', '% Cierre']);

  const byCat = {};
  rows.forEach(r => {
    const c = (iCat >= 0 ? String(r[iCat] || '').trim() : '') || '(Sin categoría)';
    if (!byCat[c]) byCat[c] = { open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') byCat[c].open++;
    else byCat[c].closed++;
  });

  Object.entries(byCat)
    .sort((a, b) => (b[1].open + b[1].closed) - (a[1].open + a[1].closed))
    .forEach(([cat, v]) => {
      const t   = v.open + v.closed;
      const pct = t > 0 ? Math.round((v.closed / t) * 100) + '%' : '0%';
      tableRow(
        [cat, v.open, v.closed, t, pct],
        [null,
          v.open   > 0 ? { bg: RL, fc: '#791F1F' } : null,
          v.closed > 0 ? { bg: GL, fc: '#27500A' } : null,
          null, null]
      );
    });

  row++;

  // ── BY REPORTER ──────────────────────────────
  sectionHeader('POR INSPECTOR / USUARIO', 4);
  tableHeader(['Inspector', 'Abiertos', 'Cerrados', 'Total']);

  const byRep = {};
  rows.forEach(r => {
    const rep = String(r[iRep] || '(Desconocido)').trim() || '(Desconocido)';
    if (!byRep[rep]) byRep[rep] = { open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') byRep[rep].open++;
    else byRep[rep].closed++;
  });

  Object.entries(byRep)
    .sort((a, b) => (b[1].open + b[1].closed) - (a[1].open + a[1].closed))
    .forEach(([rep, v]) => {
      const t = v.open + v.closed;
      tableRow(
        [rep, v.open, v.closed, t],
        [null,
          v.open   > 0 ? { bg: RL, fc: '#791F1F' } : null,
          v.closed > 0 ? { bg: GL, fc: '#27500A' } : null,
          null]
      );
    });

  // ── COLUMN WIDTHS ─────────────────────────────
  dash.setColumnWidth(1, 90);
  dash.setColumnWidth(2, 230);
  [3,4,5,6,7,8].forEach(c => dash.setColumnWidth(c, 80));
}

// ════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ════════════════════════════════════════════════════════════

// Returns the PunchList sheet — creates with correct headers if missing,
// adds any missing columns if it already exists
function getSheet() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    // Brand new sheet
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(COLS);
    const hdr = sheet.getRange(1, 1, 1, COLS.length);
    hdr.setBackground('#0F6E56').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(8, 280);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(11, 200);
    return sheet;
  }

  // Sheet exists — check for missing columns and add them
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  COLS.forEach(col => {
    if (!existing.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol)
        .setValue(col)
        .setBackground('#0F6E56').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
    }
  });
  return sheet;
}

// Returns {columnName: 1-based-index} map from current header row
function getColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

// Build a row array aligned to the actual sheet columns
function buildRow(colMap, data) {
  const totalCols = Math.max(...Object.values(colMap));
  const row       = new Array(totalCols).fill('');
  Object.entries(data).forEach(([key, val]) => {
    if (colMap[key]) row[colMap[key] - 1] = val;
  });
  return row;
}

// Extract 2-3 letter KKS system code from a tag string
function extractKKS(tag) {
  if (!tag) return '';
  const m = tag.trim().toUpperCase().match(/[A-Z]{2,3}/);
  if (!m) return '';
  const c3 = m[0].substring(0, 3);
  const c2 = m[0].substring(0, 2);
  if (KKS[c3]) return c3;
  if (KKS[c2]) return c2;
  return c3; // return even if unknown, for display
}

function sanitize(str) {
  return (str || 'Usuario').replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ_\-]/g, '_').substring(0, 30);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
