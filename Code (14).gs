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
      case 'submit':           return submitItem(data);
      case 'close':            return closeItem(data);
      case 'reopen':           return reopenItem(data);
      case 'getAll':           return getAllItems(data);
      case 'verifyAdmin':      return verifyAdmin(data);
      case 'updateCategory':   return updateCategory(data);
      case 'getConfig':        return getConfig();
      case 'refreshDashboard': return refreshDashboardAction(data);
      case 'getReport':         return getReport(data);
      case 'backupPhotos':      return backupPhotos(data);
      default:                  return jsonResponse({ error: 'Acción desconocida' });
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
    const statusCol = colMap['Estatus'] !== undefined ? colMap['Estatus'] + 1 : 12;
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
  const iID     = colMap['ID']           !== undefined ? colMap['ID']           : 0;
  const iStatus = colMap['Estatus']      !== undefined ? colMap['Estatus']      : 9;
  const iCloser = colMap['Cerrado por']  !== undefined ? colMap['Cerrado por']  : 10;
  const iDate   = colMap['Fecha cierre'] !== undefined ? colMap['Fecha cierre'] : 11;

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
  const iID     = colMap['ID']           !== undefined ? colMap['ID']           : 0;
  const iStatus = colMap['Estatus']      !== undefined ? colMap['Estatus']      : 9;
  const iCloser = colMap['Cerrado por']  !== undefined ? colMap['Cerrado por']  : 10;
  const iDate   = colMap['Fecha cierre'] !== undefined ? colMap['Fecha cierre'] : 11;

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
  const iID  = colMap['ID']        !== undefined ? colMap['ID']        : 0;
  const iCat = colMap['Categoría'] !== undefined ? colMap['Categoría'] : 8;

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
//  REFRESH DASHBOARD (callable from Make.com via HTTP POST)
// ════════════════════════════════════════════════════════════
function refreshDashboardAction(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD)
    return jsonResponse({ error: 'No autorizado' });
  try {
    refreshDashboard(getSheet());
    return jsonResponse({ success: true, message: 'Dashboard actualizado correctamente', ts: new Date().toISOString() });
  } catch(e) {
    return jsonResponse({ success: false, error: e.message });
  }
}

// ════════════════════════════════════════════════════════════
//  WEEKLY REPORT (#2) — callable from Make every Friday
// ════════════════════════════════════════════════════════════
function getReport(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD)
    return jsonResponse({ error: 'No autorizado' });

  var sheet  = getSheet();
  var colMap = getColMap(sheet);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return jsonResponse({ error: 'Sin datos' });

  var iID    = colMap['ID']            !== undefined ? colMap['ID']            : 0;
  var iArea  = colMap['Área']          !== undefined ? colMap['Área']          : 4;
  var iDesc  = colMap['Descripción']   !== undefined ? colMap['Descripción']   : 5;
  var iRep   = colMap['Reportado por'] !== undefined ? colMap['Reportado por'] : 6;
  var iStat  = colMap['Estatus']       !== undefined ? colMap['Estatus']       : 9;
  var iFecha = colMap['Fecha']         !== undefined ? colMap['Fecha']         : 1;
  var iKKS   = colMap['KKS/Tag']       !== undefined ? colMap['KKS/Tag']       : 2;
  var iSys   = colMap['Sistema KKS']   !== undefined ? colMap['Sistema KKS']   : -1;

  var rows     = values.slice(1);
  var total    = rows.length;
  var abiertos = rows.filter(function(r){ return r[iStat] === 'Abierto'; });
  var cerrados = rows.filter(function(r){ return r[iStat] === 'Cerrado'; });

  // Group by area
  var byArea = {};
  rows.forEach(function(r) {
    var a = String(r[iArea] || 'Sin área').trim();
    if (!byArea[a]) byArea[a] = { open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') byArea[a].open++;
    else byArea[a].closed++;
  });

  var fecha = new Date().toLocaleDateString('es-DO', { weekday:'long', year:'numeric', month:'long', day:'numeric' });

  // Build area rows HTML
  var areaRows = '';
  var aIdx = 0;
  Object.keys(byArea).forEach(function(area) {
    var v  = byArea[area];
    var t  = v.open + v.closed;
    var pct = t > 0 ? Math.round((v.closed/t)*100) + '%' : '0%';
    var bg = aIdx % 2 === 0 ? '#ffffff' : '#f5f5f5';
    areaRows += '<tr style="background:' + bg + '">' +
      '<td style="padding:8px;border-bottom:1px solid #eee">' + area + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;text-align:center;color:' + (v.open>0?'#791F1F':'#333') + ';font-weight:' + (v.open>0?'bold':'normal') + '">' + v.open + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;text-align:center;color:' + (v.closed>0?'#27500A':'#333') + ';font-weight:' + (v.closed>0?'bold':'normal') + '">' + v.closed + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;text-align:center">' + t + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;text-align:center">' + pct + '</td>' +
      '</tr>';
    aIdx++;
  });

  // Build open items rows HTML
  var itemRows = '';
  abiertos.forEach(function(r, i) {
    var bg  = i % 2 === 0 ? '#ffffff' : '#fef5f5';
    var sys = iSys >= 0 ? (r[iSys] || '') : '';
    var desc = String(r[iDesc] || '');
    var descShort = desc.length > 80 ? desc.substring(0, 80) + '...' : desc;
    itemRows += '<tr style="background:' + bg + '">' +
      '<td style="padding:6px;font-family:monospace;white-space:nowrap">' + r[iID] + '</td>' +
      '<td style="padding:6px;white-space:nowrap">' + r[iFecha] + '</td>' +
      '<td style="padding:6px;white-space:nowrap">' + r[iArea] + '</td>' +
      '<td style="padding:6px;white-space:nowrap;font-family:monospace;font-size:11px">' + (sys || r[iKKS] || '—') + '</td>' +
      '<td style="padding:6px">' + descShort + '</td>' +
      '<td style="padding:6px;white-space:nowrap">' + r[iRep] + '</td>' +
      '</tr>';
  });

  var pct_global = total > 0 ? Math.round((cerrados.length / total) * 100) : 0;
  var dashUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/edit#gid=427250963';
  var appUrl  = 'https://gsfpunchlist.github.io/punch-list-gsf1/';

  var html =
    '<table width="100%" cellpadding="0" cellspacing="0" style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto">' +

    // Header
    '<tr><td style="background-color:#0F6E56;padding:24px 28px;border-radius:6px 6px 0 0">' +
    '<p style="margin:0;font-size:11px;color:#9FE1CB;font-weight:bold;letter-spacing:2px">REPORTE SEMANAL</p>' +
    '<p style="margin:4px 0 2px;font-size:20px;font-weight:bold;color:#ffffff">Punch List GSF1 CCPP - TSK</p>' +
    '<p style="margin:0;font-size:12px;color:#9FE1CB">' + fecha + '</p>' +
    '</td></tr>' +

    // Stats row
    '<tr><td style="background-color:#f8f8f8;padding:20px 28px;border:1px solid #e5e5e5">' +
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td width="32%" style="background:#ffffff;border:2px solid #e5e5e5;border-radius:6px;padding:16px;text-align:center">' +
    '<p style="margin:0;font-size:32px;font-weight:bold;color:#1a1a1a">' + total + '</p>' +
    '<p style="margin:4px 0 0;font-size:10px;color:#888;font-weight:bold;letter-spacing:1px">TOTAL</p></td>' +
    '<td width="4%"></td>' +
    '<td width="32%" style="background:#FFF0F0;border:2px solid #FFAAAA;border-radius:6px;padding:16px;text-align:center">' +
    '<p style="margin:0;font-size:32px;font-weight:bold;color:#CC0000">' + abiertos.length + '</p>' +
    '<p style="margin:4px 0 0;font-size:10px;color:#CC0000;font-weight:bold;letter-spacing:1px">ABIERTOS</p></td>' +
    '<td width="4%"></td>' +
    '<td width="32%" style="background:#F0FFF0;border:2px solid #88CC88;border-radius:6px;padding:16px;text-align:center">' +
    '<p style="margin:0;font-size:32px;font-weight:bold;color:#006600">' + cerrados.length + '</p>' +
    '<p style="margin:4px 0 0;font-size:10px;color:#006600;font-weight:bold;letter-spacing:1px">CERRADOS</p></td>' +
    '</tr></table>' +
    '</td></tr>' +

    // Area table
    '<tr><td style="padding:20px 28px;border:1px solid #e5e5e5;border-top:none;background:#ffffff">' +
    '<p style="margin:0 0 12px;font-size:12px;font-weight:bold;color:#0F6E56;text-transform:uppercase;letter-spacing:1px">Por Area / Disciplina</p>' +
    '<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px">' +
    '<tr style="background:#0F6E56;color:#ffffff">' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px">AREA</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px;text-align:center">ABIERTOS</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px;text-align:center">CERRADOS</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px;text-align:center">TOTAL</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px;text-align:center">% CIERRE</td>' +
    '</tr>' + areaRows +
    '</table>' +

    // Items section
    '<p style="margin:20px 0 12px;font-size:12px;font-weight:bold;color:#CC0000;text-transform:uppercase;letter-spacing:1px">Items Abiertos (' + abiertos.length + ')</p>' +
    '<table width="100%" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:12px">' +
    '<tr style="background:#7F1D1D;color:#ffffff">' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px">ID</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px">FECHA</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px">AREA</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px">DESCRIPCION</td>' +
    '<td style="padding:8px 10px;font-weight:bold;font-size:11px">INSPECTOR</td>' +
    '</tr>' + itemRows +
    '</table>' +
    '</td></tr>' +

    // CTA
    '<tr><td style="padding:24px 28px;text-align:center;background:#f8f8f8;border:1px solid #e5e5e5;border-top:none">' +
    '<a href="' + dashUrl + '" style="background-color:#0F6E56;color:#ffffff;text-decoration:none;padding:12px 30px;border-radius:5px;font-size:14px;font-weight:bold;display:inline-block;margin-right:12px">Ver Dashboard &rarr;</a>' +
    '<a href="' + appUrl + '" style="background-color:#ffffff;color:#0F6E56;text-decoration:none;padding:11px 20px;border-radius:5px;font-size:14px;font-weight:bold;display:inline-block;border:2px solid #0F6E56">Abrir App</a>' +
    '</td></tr>' +

    // Footer
    '<tr><td style="padding:14px 28px;text-align:center;background:#eeeeee;border-radius:0 0 6px 6px">' +
    '<p style="margin:0;font-size:11px;color:#888888">Punch List GSF1 CCPP &bull; TSK &bull; Reporte autom&aacute;tico cada viernes 4:00 PM</p>' +
    '</td></tr>' +

    '</table>';


    return jsonResponse({
    success:  true,
    total:    total,
    abiertos: abiertos.length,
    cerrados: cerrados.length,
    fecha:    fecha,
    html:     html,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID
  });
}

// ════════════════════════════════════════════════════════════
//  BACKUP PHOTOS TO SECOND FOLDER (#5)
// ════════════════════════════════════════════════════════════
function backupPhotos(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD)
    return jsonResponse({ error: 'No autorizado' });

  const BACKUP_FOLDER_NAME = 'Backup_Fotos_PunchList_GSF1';

  try {
    const sourceFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    // Get or create backup folder inside the same parent
    let backupFolder;
    const parents = sourceFolder.getParents();
    const parent  = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const existing = parent.getFoldersByName(BACKUP_FOLDER_NAME);
    backupFolder    = existing.hasNext() ? existing.next() : parent.createFolder(BACKUP_FOLDER_NAME);

    // Get all files already in backup
    const backupFiles = {};
    const bIter = backupFolder.getFiles();
    while (bIter.hasNext()) {
      const f = bIter.next();
      backupFiles[f.getName()] = true;
    }

    // Copy new files from source to backup
    let copied = 0; let skipped = 0;
    const srcIter = sourceFolder.getFiles();
    while (srcIter.hasNext()) {
      const file = srcIter.next();
      if (!backupFiles[file.getName()]) {
        file.makeCopy(file.getName(), backupFolder);
        copied++;
      } else { skipped++; }
    }

    return jsonResponse({
      success: true,
      copied, skipped,
      backupFolder: backupFolder.getUrl(),
      message: 'Backup completado: ' + copied + ' archivos nuevos copiados, ' + skipped + ' ya existian.'
    });

  } catch(e) {
    return jsonResponse({ success: false, error: e.message });
  }
}

// ════════════════════════════════════════════════════════════
//  MIGRATE & BACKFILL// ════════════════════════════════════════════════════════════
//  MIGRATE & BACKFILL (run once manually from editor)
// ════════════════════════════════════════════════════════════
function migrateAndRefresh() {
  const sheet   = getSheet();          // ensures headers exist & adds missing cols
  const colMap  = getColMap(sheet);
  const values  = sheet.getDataRange().getValues();

  // 0-based indices
  const iKKS   = colMap['KKS/Tag']            !== undefined ? colMap['KKS/Tag']            : 2;
  const iSys   = colMap['Sistema KKS']         !== undefined ? colMap['Sistema KKS']         : -1;
  const iDesc  = colMap['Descripción Sistema'] !== undefined ? colMap['Descripción Sistema'] : -1;
  const iStat  = colMap['Estatus']             !== undefined ? colMap['Estatus']             : 9;

  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    const raw = String(values[i][iKKS] || '');
    const hasSys = iSys > 0 && values[i][iSys - 1];
    if (raw && !hasSys) {
      const code = extractKKS(raw);
      const desc = code ? (KKS[code] || '') : '';
      if (code && iSys  >= 0) sheet.getRange(i + 1, iSys + 1).setValue(code);
      if (desc && iDesc >= 0) sheet.getRange(i + 1, iDesc + 1).setValue(desc);
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

  // All indices are 0-based (direct array access)
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
  let addedCols = false;
  COLS.forEach(col => {
    if (!existing.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol)
        .setValue(col)
        .setBackground('#0F6E56').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
      addedCols = true;
      Logger.log('Added column: ' + col + ' at position ' + newCol);
    }
  });
  if (addedCols) sheet.autoResizeColumns(1, sheet.getLastColumn());
  return sheet;
}

// Returns {columnName: 1-based-index} map from current header row
// Returns {headerName: 0-based-index} for direct array access: r[col["KKS/Tag"]]
function getColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

// Build a row array aligned to the actual sheet columns
// colMap is 0-based; row array length = max index + 1
function buildRow(colMap, data) {
  const totalCols = Math.max(...Object.values(colMap)) + 1;
  const row       = new Array(totalCols).fill('');
  Object.entries(data).forEach(([key, val]) => {
    if (colMap[key] !== undefined) row[colMap[key]] = val;
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
