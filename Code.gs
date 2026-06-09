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
  'Foto URL','Estatus','Cerrado por','Fecha cierre','Comentario Cierre','Foto Cierre','Timestamp'
];

// ── KKS SYSTEM LOOKUP — GSF-1 CCPP (TSKI-002777) ────────────
const KKS = {
  // A — Grid & Distribution
  'AE':'110-150 kV Systems','AEA':'Breakers','AEB':'Disconnectors','AEC':'Current transformers',
  'AED':'Voltage transformers','AET':'Main transformers','AEW':'Lightning protection kV','AN':'<1 kV systems',
  'ANE':'Low voltage switchgear 480/208 Vac','ANF':'Low voltage switchgear UPS','ANK':'DC switchgear 125/110 Vdc','ANQ':'DC switchgear 24 Vdc',
  'AP':'Control consoles','APA':'SCADA panel','AQ':'Measuring and metering equipment','AQA':'Tariff metering panel',
  'AR':'Protection equipment','ARA':'Protection and control panel','AS':'Decentralized panels and cabinets','ASQ':'Metering junction boxes',
  'ASR':'Protection junction boxes','AT':'Transformer equipment','ATA':'Auxiliary transformer','AY':'Communication equipment',
  'AYA':'Telephone system (PABX)','AYB':'Control console telephone system','AYE':'Fire alarm system','AYG':'Telecommunication panel',
  // B — Power Transmission
  'BA':'Power transmission','BAA':'Generator busbars','BAC':'Generation circuit breaker','BAT':'Generator transformers incl. cooling',
  'BAY':'Control and protection equipment','BAW':'Earthing and lightning protection','BB':'MV distribution boards — normal system','BBA':'MV switchgear',
  'BBT':'Auxiliary transformer','BBY':'Control and protection equipment','BD':'MV emergency distribution boards (diesel)','BDA':'MV switchgear (emergency)',
  'BDT':'MV diesel generator','BDY':'Control and protection equipment','BF':'LV main distribution boards — normal','BFA':'LV main distribution board A',
  'BFB':'LV main distribution board B','BFC':'LV main distribution board C','BFD':'LV main distribution board D','BFT':'LV main distribution transformers',
  'BFY':'Control and protection equipment','BJ':'LV subdistribution boards — normal','BJA':'MCC','BJT':'LV auxiliary power transformers',
  'BJD':'ST MCC (Motor Control Centre — Steam Turbine)',
  'BJY':'Control and protection equipment','BL':'LV subdistribution boards — general purpose','BLA':'LV subdistribution board A','BLB':'LV subdistribution board B',
  'BLC':'LV subdistribution board C','BLD':'LV subdistribution board D','BLE':'LV subdistribution board E','BLF':'LV subdistribution board F',
  'BLG':'LV subdistribution board G','BLH':'LV subdistribution board H','BLJ':'LV subdistribution board J','BLK':'LV subdistribution board K',
  'BLT':'LV auxiliary power transformers','BLY':'Control and protection equipment','BM':'LV distribution boards — diesel emergency','BMA':'LV emergency distribution board A',
  'BMB':'LV emergency distribution board B','BME':'LV emergency distribution board E','BMY':'Control and protection equipment','BN':'LV subdistribution boards — diesel emergency',
  'BNA':'LV emergency subdistribution board A',
  'BNT':'Emergency Lighting Transformer','BNB':'LV emergency subdistribution board B','BNC':'LV emergency subdistribution board C','BND':'LV emergency subdistribution board D',
  'BNE':'LV emergency subdistribution board E','BNF':'LV emergency subdistribution board F','BNG':'LV emergency subdistribution board G','BNH':'LV emergency subdistribution board H',
  'BNJ':'LV emergency subdistribution board J','BNK':'LV emergency subdistribution board K','BR':'LV distribution boards — UPS','BRA':'UPS distribution board A',
  'BRB':'UPS distribution board B','BRU':'UPS inverter','BRT':'Isolation transformer','BT':'Battery systems',
  'BTA':'DC system batteries A','BTB':'DC system batteries B','BTL':'Battery charger','BU':'DC distribution boards — normal',
  'BUA':'DC distribution board A','BUB':'DC distribution board B','BY':'Control and protection equipment','BYA':'Generator and transformer protection cabinets',
  'BYB':'Mains coupling relay cabinet',
  // C — I&C
  'CA':'Protective interlocks','CAA':'BPS cabinets','CB':'Functional group control','CBA':'ST generator and ST control panel',
  'CBP':'Synchronization cabinets','CC':'Binary signal conditioning','CCA':'Binary signal conditioning cabinets','CD':'Drive control interface',
  'CDA':'Drive control cabinets','CE':'Annunciation','CEA':'Annunciation system cabinets','CEJ':'Fault recording',
  'CF':'Measurement and recording','CFA':'Measurement cabinets','CFQ':'Recording cabinets (meters, recorders)','CH':'Protection',
  'CHA':'Generator and transformer protection cabinet','CHE':'Protection equipment','CJ':'Unit coordination level','CJA':'Unit control system',
  'CJD':'Start-up and setpoint control','CJF':'Boiler control system','CJJ':'I&C cabinets for steam turbine set','CJP':'I&C cabinets for gas turbine set',
  'CJU':'I&C cabinets for main machinery','CK':'Process computer system','CKA':'Process computer system','CKJ':'Access control computer',
  'CKN':'Process computer system','CP':'Separate automation system','CR':'Process Control System','CRL':'Communications Rack / Fiber Optic Distribution Rack','CRM':'EDH Network Rack','CRJ':'Automation system (fail-safe)',
  'CRK':'Automation system (high availability)','CRR':'Communication (terminal bus)','CRS':'Communication (plant bus)','CRT':'Communication (field bus)',
  'CRU':'Operation and monitoring','CRV':'Engineering','CRX':'Process optimization','CSA':'BOP Controller and RIO (Remote I/O)','CW':'Control rooms',
  'CWA':'Main control consoles','CX':'Local control stations','CXA':'Local control stations','CY':'Communication and information system',
  'CYA':'Telephone system','CYB':'Control console telephone system','CYD':'Alarm system (optical)','CYE':'Fire alarm system',
  'CYF':'Clock system','CYG':'Remote control system','CYQ':'Gas detection system','CYS':'CCTV System','CYW':'On Site Monitor / Workstation',
  // E — Fuel Supply
  'EG':'Liquid fuel supply','EGA':'Receiving equipment incl. pipeline','EGB':'Tank farm','EGC':'Pump system',
  'EGD':'Piping system','EGE':'Mechanical cleaning','EGF':'Temporary storage system','EGG':'Preheating',
  'EGR':'Residues removal system','EGT':'Heating medium system','EGV':'Lubricant supply system','EGX':'Fluid supply for control equipment',
  'EGY':'Control and protection equipment','EK':'Gaseous fuel supply','EKA':'Receiving equipment incl. pipeline','EKB':'Scrubber system',
  'EKC':'Heating system','EKD':'Reducing station','EKE':'Mechanical cleaning and scrubbing system','EKF':'Storage system',
  'EKG':'Piping system','EKH':'Main pressure boosting system','EKN':'Purging system','EKR':'Residues removal system',
  'EKT':'Heating medium system','EKU':'Billing meter station','EKY':'Control and protection system',
  // G — Water Supply
  'GA':'Raw water supply','GAA':'Extraction and mechanical cleaning','GAC':'Piping and channel system','GAD':'Storage system',
  'GAF':'Pump system','GB':'Treatment system (carbonate hardness removal)','GBB':'Filtering and mechanical cleaning system','GBC':'Aeration and gas injection',
  'GBD':'Precipitation system','GBK':'Piping, storage and pump system','GBL':'Storage system','GBN':'Chemicals supply system',
  'GBR':'Flushing water and residues removal','GBS':'Sludge thickening system','GBY':'Control and protection equipment','GC':'Treatment system (demineralization)',
  'GCB':'Filtering and mechanical cleaning system','GCC':'Aeration and gas injection','GCF':'Ion exchange and reverse osmosis system','GCK':'Piping, storage and pump system',
  'GCL':'Storage system','GCN':'Chemicals supply system','GCP':'Regeneration and flushing equipment','GCR':'Flushing water and residues removal',
  'GCY':'Control and protection system','GH':'Distribution systems (not drinking water)','GHA':'Distribution systems','GHC':'Distribution after demineralization',
  'GHD':'Distribution system after treatment','GK':'Potable water supply','GKA':'Receiving point','GKB':'Storage and distribution system',
  'GKC':'Potable water supply system','GM':'Process drainage system','GMA':'BOP oily effluent system','GMB':'Non-oily effluent system',
  'GMC':'Sanitary water drains','GMD':'GT washing water discharge','GME':'Final effluent water treatment plant','GMF':'Drains to sink',
  'GMM':'Oil/water separator','GMN':'Chemicals supply system','GT':'Water recovery from wastewater','GTA':'Backwashing recovery system',
  'GTB':'Water recovery treatment','GTC':'Clarified water from static thickener','GTD':'Clarified water from centrifugal decanter','GU':'Rainwater collection and drainage',
  'GUA':'Non-oily rainwater collection and drainage','GUB':'Oily rainwater collection and drainage',
  // H — Heat Generation
  'HA':'Pressure system — feed water and steam sections','HAA':'LP part-flow feed heating system','HAB':'HP part-flow feed heating system','HAC':'Economizer system',
  'HAD':'Evaporator system','HAG':'Circulation system','HAH':'HP superheater systems','HAJ':'Reheat system',
  'HAK':'Secondary reheat system','HAN':'Pressure system drainage and venting','HAY':'Control and protection equipment','HB':'Support structure, enclosure, steam generator interior',
  'HBA':'Frame incl. foundations','HBB':'Enclosures','HBD':'Platforms and stairways','HBK':'Steam generator interior',
  'HJ':'Ignition firing equipment','HJA':'Ignition burners','HJG':'Gas pressure reduction and distribution','HJL':'Combustion air supply system',
  'HJM':'Atomizer medium supply (steam)','HJN':'Atomizer medium supply (air)','HJP':'Cooling supply system (steam)','HJQ':'Cooling supply system (air)',
  'HJR':'Purging medium supply (steam)','HJS':'Purging medium supply (air)','HJT':'Heating medium supply (steam)','HJU':'Heating medium supply (hot water)',
  'HJV':'Lubricant supply system','HJW':'Sealing fluid supply system','HJX':'Fluid supply for control equipment','HJY':'Control and protection equipment',
  'HL':'Ducting system air','HLA':'Ducting system','HLB':'Forced draught fan system','HLC':'Air preheating (not flue gas heated)',
  'HLD':'Air preheating (flue gas heated)','HLE':'Ducting system oxygen','HLF':'Oxygen fan system','HLG':'Oxygen preheating (not flue gas heated)',
  'HLH':'Oxygen preheating flue gas heated','HLU':'Air pressure relief system','HLV':'Lubricant supply system','HLW':'Sealing fluid supply system',
  'HLX':'Fluid supply for control equipment','HLY':'Control and protection equipment','HN':'Flue gas exhaust','HNA':'Ducting system',
  'HNE':'Smokestack system (chimney)',
  // L — Steam/Water Cycle
  'LA':'Feedwater system','LAA':'Storage and deaeration (feedwater tank)','LAB':'Feedwater piping system','LAC':'Feedwater pump system',
  'LAD':'HP feedwater heating system','LAE':'HP desuperheating spray system','LAF':'IP desuperheating spray system','LAN':'N2 injection in LA system',
  'LAV':'Lubricant supply system','LB':'Steam system','LBA':'Main steam piping system','LBB':'Hot reheat steam system',
  'LBC':'Cold reheat steam system','LBD':'Extraction piping system','LBE':'Back-pressure piping system','LBG':'Auxiliary steam piping system',
  'LBH':'Start-up and shutdown steam system','LBQ':'Extraction steam piping for HP feedwater heating','LBS':'Extraction steam piping for LP feedwater heating','LC':'Condensate system',
  'LCA':'Main condensate piping system','LCB':'Main condensate pump system','LCE':'Condensate desuperheating spray system','LCH':'HP heater drains system',
  'LCJ':'LP heater drains system','LCL':'Steam generator drain system','LCM':'Clean drains system','LCN':'N2 injection in LC system',
  'LCP':'Standby condensate system','LCQ':'Steam generator blowdown system','LCR':'Standby condensate distribution system','LCS':'Reheater drains system',
  'LCT':'Auxiliary steam condensate system','LCW':'Sealing and cooling drains system','LCY':'Control and Protection System','LF':'Common installations for steam/water/gas cycles',
  'LFN':'Proportioning system for feedwater and condensate',
  // M — Main Machinery
  'MA':'Steam turbine plant','MAA':'HP Turbine','MAB':'IP Turbine','MAC':'LP Turbine',
  'MAD':'Bearings','MAG':'Condensing system','MAJ':'Air removal system','MAK':'Transmission gear incl. turning gear',
  'MAL':'Drains and vents systems','MAM':'Leak off steam system','MAN':'Turbine bypass station','MAP':'LP turbine bypass system',
  'MAV':'Lubricant supply system','MAW':'Sealing, heating and cooling steam system','MAX':'Non-electric control and protection equipment','MAY':'Electric control and protection equipment',
  'MB':'Gas Turbine Plant','MBA':'Turbine and compressor rotor with common casing','MBD':'Bearings','MBH':'Cooling and sealing gas system',
  'MBJ':'Start-up unit','MBK':'Transmission gear incl. turning gear and barring gear','MBL':'Intake air and cold gas system','MBM':'Combustion chamber',
  'MBP':'Fuel supply system (gaseous)','MBR':'Exhaust gas system (open cycle)','MBV':'Lubricant supply system','MBX':'Non-electric control and protection equipment',
  'MBY':'Electrical control and protection equipment','MJ':'Diesel engine plant','MJA':'Engine','MJB':'Turbocharger and blower',
  'MJG':'Liquid cooling system','MJH':'Air intercooling system','MJN':'Fuel systems','MJP':'Start-up unit',
  'MJQ':'Air intake system','MJR':'Exhaust gas system','MJV':'Lubricant supply system','MJY':'Control and protection equipment',
  'MK':'Generator Plant','MKA':'Generator incl. stator, rotor and cooling','MKB':'Generator exciter set','MKC':'Generator exciter set',
  'MKD':'Bearings','MKF':'Stator/rotor liquid cooling system','MKG':'Stator/rotor H2 cooling system','MKH':'Stator/rotor H2/N2/CO2 cooling system',
  'MKJ':'Stator/rotor hydrogen air cooling system','MKW':'Sealing fluid supply system','MKX':'Fluid supply for control equipment','MKY':'Control and protection equipment',
  'ML':'Electro-motive plant (incl. motor generator)','MLA':'Motor/generator frame incl. stator and rotor','MLD':'Bearings','MLF':'Stator/rotor liquid cooling system',
  // P — Cooling Water
  'PA':'Circulating (main cooling) water system','PAA':'Extraction and mechanical cleaning for direct cooling','PAB':'Circulating water piping and culvert system','PAC':'Circulating water pump system',
  'PAD':'Recirculating cooling system and outfall','PAE':'Cooling tower pump system','PAH':'Condenser cleaning system','PB':'Circulating water treatment system',
  'PBN':'Chemicals supply system','PG':'Closed Cooling Water System','PGB':'Closed cooling water system (Combined Cycle)','PGD':'Closed cooling water system (Gas Turbine)',
  // Q — Auxiliary Systems
  'QC':'Central chemicals supply','QCA':'Oxygen scavenger injection system','QCB':'Neutralizer injection system','QCC':'Phosphate injection system',
  'QCD':'Corrosion inhibitor injection system','QCE':'Other chemicals injection system','QCF':'Ammonia injection system','QCG':'Injection systems common elements',
  'QE':'Service compressed air supply','QEA':'Service air generation','QEB':'Service air distribution incl. power island/BOP','QF':'Control air supply',
  'QFA':'Compressed air generation system','QFB':'Instrument air distribution incl. power island/BOP','QFC':'Service air distribution within power island/BOP','QFD':'HP compressed air system',
  'QH':'Auxiliary steam generating system','QHA':'Pressure system','QHH':'Main firing system','QHY':'Control and protection equipment',
  'QJ':'Central gas supply','QJA':'Hydrogen storage and distribution system','QJB':'Carbon dioxide storage and distribution system','QJN':'Nitrogen gas storage and distribution system',
  'QL':'Feedwater, steam, condensate cycle of auxiliary steam system','QLA':'Feedwater system','QLB':'Steam system','QLC':'Condensate system',
  'QU':'Sampling systems','QUA':'Sampling drains','QUB':'Steam sampling system','QUC':'Water sampling system',
  'QUD':'Flue gas sampling system',
  // S — Ancillary Systems
  'SA':'HVAC systems for conventional area','SAA':'Buildings ventilation','SAB':'HVAC Panel and Transformer','SAK':'Buildings air conditioning','SAM':'Turbine hall ventilation',
  'SAU':'Buildings heating','SG':'Fire protection systems','SGA':'Fire water system — conventional area','SGC':'Spray deluge systems',
  'SGE':'Sprinkler systems','SGF':'Foam fire-fighting systems','SGJ':'CO2 fire-fighting systems','SGL':'Powder fire-fighting systems',
  'SGY':'Control and protection equipment','SM':'Cranes, stationary hoist and conveying appliances','SMA':'Crane system','SR':'Workshop, stores, laboratory inside controlled area',
  'SRA':'Hot workshop equipment','SRC':'Maintenance areas in controlled area','SRG':'Hot laboratory equipment','SRH':'Health physics laboratory equipment',
  'SRP':'Staff amenities in controlled area','SRY':'Control and protection equipment','ST':'Workshop, stores, laboratory outside controlled area','STA':'Workshop equipment outside controlled area',
  'STC':'Maintenance areas outside controlled area','STE':'Stores and filling station equipment','STG':'Laboratory equipment','STP':'Staff amenities',
  'STY':'Control and protection equipment',
  // AA — Mechanical Equipment Type Identifiers (component codes)
  'AB':'Isolating elements, air locks','AC':'Heat exchangers, heat transfer surfaces',
  'AF':'Continuous conveyors and feeders','AG':'Generator units',
  'AH':'Heating, cooling and air conditioning units','AJ':'Size reduction equipment',
  'AK':'Compacting and packaging equipment','AM':'Mixers and agitators',
  'AV':'Combustion equipment','AW':'Stationary tooling and treatment equipment',
  'AX':'Test and monitoring equipment','AZ':'Injection quills',
  'BE':'Shafts (erection and maintenance)','BP':'Flow restrictors, limiters and orifices',
  'BQ':'Hangers, supports, racks, piping penetrations','BS':'Silencers','BZ':'Miscellaneous',
  // C — Measuring Circuit Type Identifiers
  'CG':'Distance, length, position, direction of rotation','CL':'Level (also dividing level)',
  'CM':'Moisture and humidity','CS':'Velocity, speed, frequency and acceleration',
  'CT':'Temperature','CU':'Combined and other variables','CV':'Viscosity',
  // G — Electrical, I&C Equipment Type Identifiers
  'GD':'Instrumentation junction box','GE':'Instrumentation junction box',
  'GF':'Instrumentation junction box','GG':'Junction box for thermocouples',
  'GJ':'Processing and storage equipment for automation','GL':'Limiting equipment',
  'GN':'Network equipment','GP':'Subdistribution/junction boxes for lighting',
  'GQ':'Subdistribution/junction boxes for power sockets','GR':'DC generating equipment and batteries',
  'GW':'Cabinet power supply equipment','GX':'Actuating equipment for electrical variables',
  // U — Civil Structures
  'UA':'Structures for grid and distribution systems','UAA':'Switchyard structure','UAD':'Structures for grid and distribution (Busbar)','UAG':'Structure for transformers',
  'UAH':'Structure for supports and equipment (Insulator)','UAJ':'Structures for grid and distribution (Circuit Breaker)','UAK':'Structures for grid and distribution (Pantograph Disconnector)','UAL':'Structures for grid and distribution (Earthing Disconnector)',
  'UAM':'Structures for grid and distribution (Busbar Disconnector)','UAN':'Structures for grid and distribution (Current Transformer)','UAP':'Structures for grid and distribution (Capacitive Voltage Transformer)','UAQ':'Structures for grid and distribution (Surge Arrester)',
  'UB':'Structures for power transmission and auxiliary power supply','UBA':'Switchgear building','UBB':'Structure for power transmission','UBR':'Structure for power transmission',
  'UL':'Structures for Steam cycle','ULA':'Feedwater pump house','ULC':'Structure for condensate system','ULF':'Structure for steam-cycle',
  'UM':'Structures for main machine set','UMA':'Steam turbine Building','US':'Area / Building','USG':'Fire pump house',
  'UST':'Workshop','USU':'Storage Building','USV':'Laboratory Building','USY':'Bridge Structure',
  'UT':'Structures for auxiliary systems','UTF':'Compressed air system Building','UY':'General service structures','UYA':'Office and staff amenities building',
  'UYE':'Gate house',
};

const CATEGORIES = [
  'Categoría A — Seguridad: riesgo inmediato para personas, equipos o ambiente. Rechazo automático del sistema. Solución inmediata requerida.',
  'Categoría B — Operación: afecta la operación pero no la seguridad. Permite comisionamiento en frío. Solución antes del Hot Commissioning.',
  'Categoría C — Funcional menor: afecta condición funcional menor, no compromete seguridad. Permite actividades de comisionamiento en caliente.',
  'Categoría D — Menor: no incluido en categorías anteriores. No impide entrega del sistema al propietario. Resolución acordada en reunión especial.'
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
      case 'refreshDashboard':   return refreshDashboardAction(data);
      case 'protectSheet':       return protectSheetAction(data);
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
  lock.waitLock(15000);
  try {
    const sheet   = getSheet();
    const colMap  = getColMap(sheet);

    // Atomic ID counter using PropertiesService — prevents duplicates with concurrent users
    const props   = PropertiesService.getScriptProperties();
    const current = parseInt(props.getProperty('PL_COUNTER') || '0', 10);
    const next    = current + 1;
    props.setProperty('PL_COUNTER', String(next));
    const id      = 'PL-' + String(next).padStart(3, '0');
    const kksCode  = extractKKS(data.kks);
    const kksDesc  = kksCode ? (KKS[kksCode] || kksCode) : '';
    let   photoUrl = '';

    // Upload photo to Drive — organizado por Anno/Mes/Dia
    if (data.photo && data.photo.length > 100) {
      try {
        var rootFolder  = DriveApp.getFolderById(CONFIG.FOLDER_ID);
        var photoDate   = data.fecha ? new Date(data.fecha + 'T12:00:00') : new Date();
        var year        = photoDate.getFullYear().toString();
        var monthNames  = ['01-Enero','02-Febrero','03-Marzo','04-Abril','05-Mayo','06-Junio','07-Julio','08-Agosto','09-Septiembre','10-Octubre','11-Noviembre','12-Diciembre'];
        var month       = monthNames[photoDate.getMonth()];
        var day         = String(photoDate.getDate()).padStart(2, '0');
        var yIter = rootFolder.getFoldersByName(year);
        var yearFolder = yIter.hasNext() ? yIter.next() : rootFolder.createFolder(year);
        var mIter = yearFolder.getFoldersByName(month);
        var monthFolder = mIter.hasNext() ? mIter.next() : yearFolder.createFolder(month);
        var dIter = monthFolder.getFoldersByName(day);
        var dayFolder = dIter.hasNext() ? dIter.next() : monthFolder.createFolder(day);
        const b64      = data.photo.split(',')[1];
        const mime     = data.photo.split(';')[0].split(':')[1] || 'image/jpeg';
        const ext      = mime.includes('png') ? 'png' : 'jpg';
        const name     = id + '_' + data.fecha + '_' + sanitize(data.reportedBy) + '.' + ext;
        const blob     = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
        const file     = dayFolder.createFile(blob);
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
  const iID      = colMap['ID']                !== undefined ? colMap['ID']                : 0;
  const iStatus  = colMap['Estatus']           !== undefined ? colMap['Estatus']           : 9;
  const iCloser  = colMap['Cerrado por']       !== undefined ? colMap['Cerrado por']       : 10;
  const iDate    = colMap['Fecha cierre']      !== undefined ? colMap['Fecha cierre']      : 11;
  const iComment = colMap['Comentario Cierre'] !== undefined ? colMap['Comentario Cierre'] : -1;
  const iClosePh = colMap['Foto Cierre']       !== undefined ? colMap['Foto Cierre']       : -1;



  for (let i = 1; i < values.length; i++) {
    if (values[i][iID] === data.id) {
      const r = i + 1;

      // Upload closing photo if provided
      var closePhotoUrl = '';
      if (data.closePhoto && data.closePhoto.length > 100) {
        try {
          var rootFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
          var today      = new Date();
          var year       = today.getFullYear().toString();
          var monthNames = ['01-Enero','02-Febrero','03-Marzo','04-Abril','05-Mayo','06-Junio',
                            '07-Julio','08-Agosto','09-Septiembre','10-Octubre','11-Noviembre','12-Diciembre'];
          var month      = monthNames[today.getMonth()];
          var day        = String(today.getDate()).padStart(2, '0');

          var yIter = rootFolder.getFoldersByName(year);
          var yearFolder = yIter.hasNext() ? yIter.next() : rootFolder.createFolder(year);
          var mIter = yearFolder.getFoldersByName(month);
          var monthFolder = mIter.hasNext() ? mIter.next() : yearFolder.createFolder(month);
          var dIter = monthFolder.getFoldersByName(day);
          var dayFolder = dIter.hasNext() ? dIter.next() : monthFolder.createFolder(day);

          var b64  = data.closePhoto.split(',')[1];
          var mime = data.closePhoto.split(';')[0].split(':')[1] || 'image/jpeg';
          var ext  = mime.includes('png') ? 'png' : 'jpg';
          var name = 'CLOSE_' + data.id + '_' + day + '-' + month + '_' + sanitize(data.closedBy || 'Admin') + '.' + ext;
          var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
          var file = dayFolder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          closePhotoUrl = file.getUrl();
        } catch(e) { Logger.log('Close photo error: ' + e.message); }
      }

      sheet.getRange(r, iStatus  + 1).setValue('Cerrado').setBackground('#C0DD97').setFontColor('#27500A').setFontWeight('bold');
      sheet.getRange(r, iCloser  + 1).setValue(data.closedBy || 'Admin');
      sheet.getRange(r, iDate    + 1).setValue(new Date().toLocaleDateString('es-DO'));
      if (iComment >= 0) sheet.getRange(r, iComment + 1).setValue(data.closeComment || '');
      if (iClosePh >= 0) sheet.getRange(r, iClosePh + 1).setValue(closePhotoUrl);
      refreshDashboard(sheet);
      return jsonResponse({ success: true, closePhotoUrl: closePhotoUrl });
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
      sheet.getRange(r, iCloser  + 1).setValue('');
      sheet.getRange(r, iDate    + 1).setValue('');
      var iCommentR = colMap['Comentario Cierre'] !== undefined ? colMap['Comentario Cierre'] : -1;
      var iClosePhR = colMap['Foto Cierre']       !== undefined ? colMap['Foto Cierre']       : -1;
      if (iCommentR >= 0) sheet.getRange(r, iCommentR + 1).setValue('');
      if (iClosePhR >= 0) sheet.getRange(r, iClosePhR + 1).setValue('');
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
//  PROTECT SHEET ACTION
// ════════════════════════════════════════════════════════════
// ── Run this function ONCE manually from the Apps Script editor ──
function protectAndAnnotateSheet() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log('Sheet not found'); return; }

  var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMap    = {};
  headers.forEach(function(h, i) { if (h) colMap[String(h).trim()] = i; });

  var statusCol = colMap['Estatus'];
  var closerCol = colMap['Cerrado por'];
  var dateCol   = colMap['Fecha cierre'];

  // Add warning note on header
  if (statusCol !== undefined) {
    sheet.getRange(1, statusCol + 1).setNote(
      '⚠ IMPORTANTE: El cierre de items debe realizarse UNICAMENTE desde el Panel Admin de la app. ' +
      'Editar esta columna directamente NO registrara la fecha de cierre ni el responsable. ' +
      'App: https://gsfpunchlist.github.io/punch-list-gsf1/'
    );
  }

  // Remove existing protections
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach(function(p) { p.remove(); });

  // Protect columns with warning
  var lastRow = Math.max(sheet.getLastRow(), 2);
  [statusCol, closerCol, dateCol].forEach(function(col) {
    if (col !== undefined) {
      var range = sheet.getRange(2, col + 1, lastRow - 1, 1);
      var prot  = range.protect();
      prot.setDescription('Solo editable via Apps Script — Panel Admin de la app');
      prot.setWarningOnly(true);
    }
  });

  Logger.log('Sheet protected and annotated successfully. Columns protected: Estatus, Cerrado por, Fecha cierre.');
}

function protectSheetAction(data) {
  if (data.password !== CONFIG.ADMIN_PASSWORD)
    return jsonResponse({ error: 'No autorizado' });
  try {
    protectAndAnnotateSheet();
    return jsonResponse({ success: true, message: 'Sheet protegido correctamente.' });
  } catch(e) {
    return jsonResponse({ success: false, error: e.message });
  }
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
  var iSys   = colMap['Sistema KKS']         !== undefined ? colMap['Sistema KKS']         : -1;
  var iSDesc = colMap['Descripcion Sistema'] !== undefined ? colMap['Descripcion Sistema'] : (colMap['Descripción Sistema'] !== undefined ? colMap['Descripción Sistema'] : -1);
  var iCat   = colMap['Categoria']          !== undefined ? colMap['Categoria']          : (colMap['Categoría'] !== undefined ? colMap['Categoría'] : -1);

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

  // ─ KKS rows
  var kksRows = '';
  var kksData = {};
  rows.forEach(function(r) {
    var code = iSys >= 0 ? String(r[iSys] || '').trim() : '';
    var desc = iSDesc >= 0 ? String(r[iSDesc] || '').trim() : '';
    if (!code) { code = extractKKS(String(r[iKKS] || '')); desc = code ? (KKS[code] || '') : ''; }
    if (!code) { code = '(Sin KKS)'; desc = ''; }
    if (!kksData[code]) kksData[code] = { desc: desc, open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') kksData[code].open++; else kksData[code].closed++;
  });
  var kksIdx = 0;
  Object.keys(kksData).sort(function(a,b){ return (kksData[b].open+kksData[b].closed)-(kksData[a].open+kksData[a].closed); }).forEach(function(code) {
    var v = kksData[code]; var t = v.open + v.closed;
    var bg = kksIdx % 2 === 0 ? '#ffffff' : '#f5f5f5';
    kksRows += '<tr style="background-color:' + bg + '">' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;color:#333;border-bottom:1px solid #eee">' + code + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;color:#555;border-bottom:1px solid #eee">' + v.desc + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.open>0?'#CC0000':'#333') + '">' + v.open + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.closed>0?'#1A6E1A':'#333') + '">' + v.closed + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;border-bottom:1px solid #eee;color:#333">' + t + '</td></tr>';
    kksIdx++;
  });

  // ─ Area rows
  var areaRows = '';
  var aIdx = 0;
  Object.keys(byArea).sort(function(a,b){ return (byArea[b].open+byArea[b].closed)-(byArea[a].open+byArea[a].closed); }).forEach(function(area) {
    var v = byArea[area]; var t = v.open + v.closed;
    var pct = t > 0 ? Math.round((v.closed/t)*100) + '%' : '0%';
    var bg = aIdx % 2 === 0 ? '#ffffff' : '#f5f5f5';
    areaRows += '<tr style="background-color:' + bg + '">' +
      '<td style="padding:7px 12px;font-size:12px;color:#333;border-bottom:1px solid #eee">' + area + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.open>0?'#CC0000':'#333') + '">' + v.open + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.closed>0?'#1A6E1A':'#333') + '">' + v.closed + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;border-bottom:1px solid #eee;color:#333">' + t + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:#0F6E56">' + pct + '</td></tr>';
    aIdx++;
  });

  // ─ Category rows
  var catRows = '';
  var catData = {};
  rows.forEach(function(r) {
    var c = (iCat >= 0 ? String(r[iCat] || '').trim() : '') || '(Sin categoria)';
    if (!catData[c]) catData[c] = { open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') catData[c].open++; else catData[c].closed++;
  });
  var catIdx = 0;
  Object.keys(catData).sort().forEach(function(cat) {
    var v = catData[cat]; var t = v.open + v.closed;
    var pct = t > 0 ? Math.round((v.closed/t)*100) + '%' : '0%';
    var bg = catIdx % 2 === 0 ? '#ffffff' : '#f5f5f5';
    catRows += '<tr style="background-color:' + bg + '">' +
      '<td style="padding:7px 12px;font-size:12px;color:#333;border-bottom:1px solid #eee">' + cat + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.open>0?'#CC0000':'#333') + '">' + v.open + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.closed>0?'#1A6E1A':'#333') + '">' + v.closed + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;border-bottom:1px solid #eee;color:#333">' + t + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:#0F6E56">' + pct + '</td></tr>';
    catIdx++;
  });

  // ─ Reporter rows
  var repRows = '';
  var repData = {};
  rows.forEach(function(r) {
    var rep = String(r[iRep] || '(Desconocido)').trim() || '(Desconocido)';
    if (!repData[rep]) repData[rep] = { open: 0, closed: 0 };
    if (r[iStat] === 'Abierto') repData[rep].open++; else repData[rep].closed++;
  });
  var repIdx = 0;
  Object.keys(repData).sort(function(a,b){ return (repData[b].open+repData[b].closed)-(repData[a].open+repData[a].closed); }).forEach(function(rep) {
    var v = repData[rep]; var t = v.open + v.closed;
    var bg = repIdx % 2 === 0 ? '#ffffff' : '#f5f5f5';
    repRows += '<tr style="background-color:' + bg + '">' +
      '<td style="padding:7px 12px;font-size:12px;color:#333;border-bottom:1px solid #eee">' + rep + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.open>0?'#CC0000':'#333') + '">' + v.open + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;font-weight:bold;text-align:right;border-bottom:1px solid #eee;color:' + (v.closed>0?'#1A6E1A':'#333') + '">' + v.closed + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;border-bottom:1px solid #eee;color:#333">' + t + '</td></tr>';
    repIdx++;
  });

    var pct_global = total > 0 ? Math.round((cerrados.length / total) * 100) : 0;
  var dashUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/edit#gid=427250963';
  var appUrl  = 'https://gsfpunchlist.github.io/punch-list-gsf1/';

  // Simple area summary for email
  var areaSummary = '';
  Object.keys(byArea).sort(function(a,b){ return (byArea[b].open+byArea[b].closed)-(byArea[a].open+byArea[a].closed); }).forEach(function(area) {
    var v = byArea[area]; var t = v.open + v.closed;
    var pct = t > 0 ? Math.round((v.closed/t)*100) : 0;
    areaSummary +=
      '<tr>' +
      '<td style="padding:6px 12px;font-size:13px;color:#333333;border-bottom:1px solid #eeeeee">' + area + '</td>' +
      '<td style="padding:6px 12px;font-size:13px;font-weight:bold;text-align:center;border-bottom:1px solid #eeeeee;color:' + (v.open>0?'#CC0000':'#333333') + '">' + v.open + '</td>' +
      '<td style="padding:6px 12px;font-size:13px;font-weight:bold;text-align:center;border-bottom:1px solid #eeeeee;color:' + (v.closed>0?'#1A6E1A':'#333333') + '">' + v.closed + '</td>' +
      '<td style="padding:6px 12px;font-size:13px;text-align:center;border-bottom:1px solid #eeeeee;color:#0F6E56;font-weight:bold">' + pct + '%</td>' +
      '</tr>';
  });

  var html =
    '<table width="580" cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,sans-serif;border:1px solid #dddddd">' +
    '<tr><td style="background-color:#0F6E56;padding:20px 24px">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td><div style="font-size:16px;font-weight:bold;color:#ffffff">Reporte Semanal</div>' +
    '<div style="font-size:13px;color:#9FE1CB;margin-top:3px">Punch List GSF1 CCPP &bull; TSK &bull; ' + fecha + '</div></td>' +
    '<td align="right"><div style="background-color:#085041;padding:10px 16px;text-align:center">' +
    '<div style="font-size:24px;font-weight:bold;color:#ffffff">' + pct_global + '%</div>' +
    '<div style="font-size:10px;color:#9FE1CB">CIERRE</div></div></td>' +
    '</tr></table></td></tr>' +
    '<tr><td style="padding:0">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td width="33%" style="padding:18px;text-align:center;background-color:#f8f8f8;border-right:1px solid #dddddd">' +
    '<div style="font-size:11px;font-weight:bold;color:#888888;letter-spacing:1px">TOTAL</div>' +
    '<div style="font-size:40px;font-weight:bold;color:#1a1a1a;line-height:1.1">' + total + '</div></td>' +
    '<td width="33%" style="padding:18px;text-align:center;background-color:#FFF0F0;border-right:1px solid #ffcccc">' +
    '<div style="font-size:11px;font-weight:bold;color:#CC0000;letter-spacing:1px">ABIERTOS</div>' +
    '<div style="font-size:40px;font-weight:bold;color:#CC0000;line-height:1.1">' + abiertos.length + '</div></td>' +
    '<td width="34%" style="padding:18px;text-align:center;background-color:#F0FFF0">' +
    '<div style="font-size:11px;font-weight:bold;color:#1A6E1A;letter-spacing:1px">CERRADOS</div>' +
    '<div style="font-size:40px;font-weight:bold;color:#1A6E1A;line-height:1.1">' + cerrados.length + '</div></td>' +
    '</tr></table></td></tr>' +
    '<tr><td style="padding:0">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr style="background-color:#333333">' +
    '<td style="padding:8px 12px;font-size:11px;font-weight:bold;color:#ffffff">AREA</td>' +
    '<td style="padding:8px 12px;font-size:11px;font-weight:bold;color:#ffffff;text-align:center">ABIERTOS</td>' +
    '<td style="padding:8px 12px;font-size:11px;font-weight:bold;color:#ffffff;text-align:center">CERRADOS</td>' +
    '<td style="padding:8px 12px;font-size:11px;font-weight:bold;color:#ffffff;text-align:center">% CIERRE</td>' +
    '</tr>' + areaSummary +
    '</table></td></tr>' +
    '<tr><td style="padding:24px;text-align:center;background-color:#f8f8f8;border-top:2px solid #0F6E56">' +
    '<div style="font-size:13px;color:#555555;margin-bottom:14px">El Dashboard completo con KKS, Categor&iacute;as e Inspectores est&aacute; disponible en Google Sheets:</div>' +
    '<a href="' + dashUrl + '" style="background-color:#0F6E56;color:#ffffff;text-decoration:none;padding:12px 32px;font-size:14px;font-weight:bold">Ver Dashboard Completo &rarr;</a>' +
    '</td></tr>' +
    '<tr><td style="padding:10px 24px;text-align:center;background-color:#eeeeee">' +
    '<div style="font-size:10px;color:#999999">Punch List GSF1 CCPP &bull; TSK &bull; Reporte autom&aacute;tico cada viernes 4:00 PM</div>' +
    '</td></tr>' +
    '</table>'

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

// ── SYNC COUNTER — ejecutar UNA VEZ después del deployment ──
function syncCounter() {
  var sheet = getSheet();
  var count = Math.max(sheet.getLastRow() - 1, 0);
  PropertiesService.getScriptProperties().setProperty('PL_COUNTER', String(count));
  Logger.log('Contador sincronizado: ' + count + ' | Próximo ID: PL-' + String(count+1).padStart(3,'0'));
}

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
  const m = tag.trim().toUpperCase().match(/[A-Z]{3,}/);
  if (!m) return '';
  const c3 = m[0].substring(0, 3);
  if (KKS[c3]) return c3;
  return '';
}

function sanitize(str) {
  return (str || 'Usuario').replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ_\-]/g, '_').substring(0, 30);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
