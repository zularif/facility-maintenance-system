// ============================================================
// FACTORY MAINTENANCE INSPECTION SYSTEM v4.0
// Code.gs — Backend Server-Side Functions
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1AXQYycsUyPDT8Z1B7o1KG4khp3JuOFu2pwYqaAWoZs4',
  ADMIN_PIN: '1234',
  APP_NAME: 'Factory Maintenance Platform',
  VERSION: '1.0'
};
// ── CACHED SPREADSHEET ACCESSOR ──────────────────────────
// GAS execution is single-threaded per request; caching ss
// avoids repeated openById overhead within one execution.
var _ss = null;
function getSSv() {
  if (!_ss) _ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return _ss;
}
// Fast yyyy-MM-dd formatter — ~100x faster than Utilities.formatDate.
// Safe for Date objects returned by Sheets (already in script timezone).
function fmtD(d) {
  var y = d.getFullYear(), m = d.getMonth() + 1, dd = d.getDate();
  return y + '-' + (m < 10 ? '0' : '') + m + '-' + (dd < 10 ? '0' : '') + dd;
}



// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  const page = e.parameter.page || 'index';
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle(CONFIG.APP_NAME + ' — Admin')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle(CONFIG.APP_NAME + ' — Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'chart') {
    return HtmlService.createHtmlOutputFromFile('Chart')
      .setTitle(CONFIG.APP_NAME + ' — Parameter Chart')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'issues') {
    return HtmlService.createHtmlOutputFromFile('IssueTracker')
      .setTitle(CONFIG.APP_NAME + ' — Issue Tracker')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'projects') {
    return HtmlService.createHtmlOutputFromFile('ProjectManagement')
      .setTitle(CONFIG.APP_NAME + ' — Project Management')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'emergency') {
    return HtmlService.createHtmlOutputFromFile('EmergencyResponse')
      .setTitle('Emergency Response — ' + CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'report') {
    var tmpl = HtmlService.createTemplateFromFile('Report');
    tmpl.preAreaID  = e.parameter.areaID  || '';
    tmpl.preEquipID = e.parameter.equipID || '';
    return tmpl.evaluate()
      .setTitle(CONFIG.APP_NAME + ' — Report')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var indexTmpl = HtmlService.createTemplateFromFile('Index');
  indexTmpl.openAreaID  = e.parameter.openArea  || '';
  indexTmpl.openPlantID = e.parameter.plantID   || '';
  indexTmpl.openFEId    = e.parameter.feId      || '';
  indexTmpl.openEquipID = e.parameter.equipID   || '';
  return indexTmpl.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ============================================================
// SPREADSHEET SETUP
// ============================================================

function setupSpreadsheet() {
  const ss = getSSv();
  const sheets = {
    Plants: ['PlantID', 'PlantName', 'Location', 'Status', 'CreatedDate'],
    Areas: ['AreaID', 'PlantID', 'AreaName', 'Description', 'Status', 'CreatedDate'],
    Equipment: ['EquipID', 'AreaID', 'EquipmentName', 'QRCode', 'AssetTag', 'InspectionCycle', 'Category', 'Status', 'CreatedDate'],
    Templates: ['TemplateID', 'EquipID', 'ParameterName', 'Unit', 'MinValue', 'MaxValue', 'SortOrder', 'Type', 'Options', 'Required', 'StatusRule'],
    InspectionHeader: ['InspectionID', 'EquipID', 'TechnicianID', 'InspectionDate', 'Status', 'Notes', 'CreatedDate'],
    InspectionDetails: ['DetailID', 'InspectionID', 'ParameterName', 'Value', 'Status', 'CreatedDate'],
    Technicians: ['TechID', 'Name', 'PIN', 'Role', 'Status', 'CreatedDate'],
    Signatures: ['SigID', 'InspectionID', 'SignatureData', 'SignedBy', 'SignedDate']
  };

  for (const [name, headers] of Object.entries(sheets)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (existing[0] !== headers[0]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1a56db')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }

  // Seed default technicians if empty
  const techSheet = ss.getSheetByName('Technicians');
  if (techSheet.getLastRow() <= 1) {
    const now = new Date().toISOString();
    techSheet.getRange(2, 1, 3, 6).setValues([
      ['T001', 'Admin User', '1234', 'Admin', 'Active', now],
      ['T002', 'Ahmad Rizal', '1111', 'Tech', 'Active', now],
      ['T003', 'Siti Aminah', '2222', 'Tech', 'Active', now]
    ]);
  }

  Logger.log('Setup complete.');
}
// ============================================================
// ADD CATEGORY COLUMN TO EXISTING EQUIPMENT SHEET
// Run addCategoryColumn() once if Equipment sheet already exists
// without the Category column (inserted at column 7, before Status)
// ============================================================

function addCategoryColumn() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('Equipment');
  if (!sheet) {
    Logger.log('ERROR: Equipment sheet not found. Run setupSpreadsheet() first.');
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Current headers: ' + headers.join(', '));

  // Check if Category already exists
  var catIdx = headers.indexOf('Category');
  if (catIdx >= 0) {
    Logger.log('Category column already exists at column ' + (catIdx + 1) + '. Nothing to do.');
    return;
  }

  // Check where Status is — Category goes before Status
  var statusIdx = headers.indexOf('Status');
  if (statusIdx < 0) {
    Logger.log('ERROR: Status column not found in Equipment sheet.');
    return;
  }

  // Insert a new column before Status (statusIdx is 0-based, sheet cols are 1-based)
  var insertAtCol = statusIdx + 1; // 1-based column number to insert before
  sheet.insertColumnBefore(insertAtCol);

  // Set the header for the new column
  sheet.getRange(1, insertAtCol).setValue('Category');
  sheet.getRange(1, insertAtCol)
    .setBackground('#1a56db')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  Logger.log('Category column inserted at column ' + insertAtCol);
  Logger.log('Now run migrateCategory() to populate the values.');
  Logger.log('addCategoryColumn() complete.');
}



// ============================================================
// AUTH FUNCTIONS
// ============================================================

function getTechnicians() {
  try {
    // Server-side cache — technicians change rarely; 5-min TTL
    var cache = CacheService.getScriptCache();
    var cached = cache.get('fmp_technicians');
    if (cached) return JSON.parse(cached);

    const ss = getSSv();
    const sheet = ss.getSheetByName('Technicians');
    const data = sheet.getDataRange().getValues();
    const result = data.slice(1)
      .filter(row => row[4] === 'Active')
      .map(row => ({
        techID: row[0],
        name: row[1],
        role: row[3]
      }));
    try { cache.put('fmp_technicians', JSON.stringify(result), 300); } catch(ce) {}
    return result;
  } catch (e) {
    return { error: e.message };
  }
}

function invalidateTechCache() {
  try { CacheService.getScriptCache().remove('fmp_technicians'); } catch(e) {}
}

function verifyLogin(techID, pin) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Technicians');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === techID && String(data[i][2]) === String(pin) && data[i][4] === 'Active') {
        return {
          success: true,
          techID: data[i][0],
          name: data[i][1],
          role: data[i][3]
        };
      }
    }
    return { success: false, error: 'Invalid PIN' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// PIN-first login — looks up technician by PIN alone (PINs are unique per person)
function verifyLoginByPin(pin) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Technicians');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(pin) && data[i][4] === 'Active') {
        return {
          success: true,
          techID: data[i][0],
          name: data[i][1],
          role: data[i][3]
        };
      }
    }
    return { success: false, error: 'PIN not recognised' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// PLANTS
// ============================================================

function getPlants() {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Plants');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).filter(r => r[0]).map(row => ({
      plantID: row[0],
      plantName: row[1],
      location: row[2],
      status: row[3],
      createdDate: row[4]
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function savePlant(plant) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Plants');
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    if (plant.plantID) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === plant.plantID) {
          sheet.getRange(i + 1, 1, 1, 5).setValues([[
            plant.plantID, plant.plantName, plant.location, plant.status || 'Active', data[i][4]
          ]]);
          return { success: true };
        }
      }
    }
    // New plant
    const newID = generateID('P', sheet);
    sheet.appendRow([newID, plant.plantName, plant.location, plant.status || 'Active', now]);
    return { success: true, plantID: newID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deletePlant(plantID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Plants');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === plantID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Plant not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// AREAS
// ============================================================

function getAreas(plantID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Areas');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    let rows = data.slice(1).filter(r => r[0]);
    if (plantID) rows = rows.filter(r => r[1] === plantID);
    return rows.map(row => ({
      areaID: row[0],
      plantID: row[1],
      areaName: row[2],
      description: row[3],
      status: row[4],
      createdDate: row[5]
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function saveArea(area) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Areas');
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    if (area.areaID) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === area.areaID) {
          sheet.getRange(i + 1, 1, 1, 6).setValues([[
            area.areaID, area.plantID, area.areaName, area.description || '', area.status || 'Active', data[i][5]
          ]]);
          return { success: true };
        }
      }
    }
    const newID = generateID('A', sheet);
    sheet.appendRow([newID, area.plantID, area.areaName, area.description || '', area.status || 'Active', now]);
    return { success: true, areaID: newID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteArea(areaID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Areas');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === areaID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Area not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// EQUIPMENT
// ============================================================

function getEquipment(areaID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Equipment');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    let rows = data.slice(1).filter(r => r[0]);
    if (areaID) rows = rows.filter(r => r[1] === areaID);
    return rows.map(row => ({
      equipID: row[0],
      areaID: row[1],
      equipmentName: row[2],
      qrCode: row[3],
      assetTag: row[4],
      inspectionCycle: row[5],
      category: row[6],
      status: row[7],
      createdDate: row[8]
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function saveEquipment(equip) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Equipment');
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    if (equip.equipID) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === equip.equipID) {
          sheet.getRange(i + 1, 1, 1, 9).setValues([[
            equip.equipID, equip.areaID, equip.equipmentName,
            equip.qrCode || 'QR-' + equip.equipID,
            equip.assetTag || 'AT-' + equip.equipID,
            equip.inspectionCycle || 'Daily',
            equip.category || '',
            equip.status || 'Active', data[i][8]
          ]]);
          return { success: true };
        }
      }
    }
    const newID = generateID('EQ', sheet);
    sheet.appendRow([
      newID, equip.areaID, equip.equipmentName,
      equip.qrCode || 'QR-' + newID,
      equip.assetTag || 'AT-' + newID,
      equip.inspectionCycle || 'Daily',
      equip.category || '',
      equip.status || 'Active', now
    ]);
    return { success: true, equipID: newID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteEquipment(equipID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Equipment');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === equipID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Equipment not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// TEMPLATES / PARAMETERS
// ============================================================

function getTemplates(equipID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Templates');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    let rows = data.slice(1).filter(r => r[0]);
    if (equipID) rows = rows.filter(r => r[1] === equipID);
    return rows
      .map(row => ({
        templateID: row[0],
        equipID: row[1],
        parameterName: row[2],
        unit: row[3],
        minValue: row[4],
        maxValue: row[5],
        sortOrder: row[6],
        type: row[7],
        options: row[8],
        required: row[9],
        statusRule: row[10]
      }))
      .sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0));
  } catch (e) {
    return { error: e.message };
  }
}

function saveTemplate(tmpl) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Templates');
    const data = sheet.getDataRange().getValues();

    if (tmpl.templateID) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === tmpl.templateID) {
          sheet.getRange(i + 1, 1, 1, 11).setValues([[
  tmpl.templateID,
  tmpl.equipID,
  tmpl.parameterName,
  tmpl.unit || '',
  tmpl.minValue || '',
  tmpl.maxValue || '',
  tmpl.sortOrder || 1,
  tmpl.type || 'numeric',
  tmpl.options || '',
  tmpl.required !== undefined ? tmpl.required : true,
  tmpl.statusRule || ''
]]);
          return { success: true };
        }
      }
    }
    const newID = generateID('TM', sheet);
   sheet.appendRow([
  newID,
  tmpl.equipID,
  tmpl.parameterName,
  tmpl.unit || '',
  tmpl.minValue || '',
  tmpl.maxValue || '',
  tmpl.sortOrder || 1,
  tmpl.type || 'numeric',
  tmpl.options || '',
  tmpl.required !== undefined ? tmpl.required : true,
  tmpl.statusRule || ''
]);
    return { success: true, templateID: newID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteTemplate(templateID) {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Templates');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === templateID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Template not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// INSPECTIONS
// ============================================================

function getInspectionsByEquip(equipID) {
  try {
    const ss = getSSv();
    const hSheet = ss.getSheetByName('InspectionHeader');
    const dSheet = ss.getSheetByName('InspectionDetails');
    const hData = hSheet.getDataRange().getValues();
    const dData = dSheet.getDataRange().getValues();

    const headers = hData.slice(1)
      .filter(r => r[0] && r[1] === equipID)
      .map(row => ({
        inspectionID: row[0],
        equipID: row[1],
        technicianID: row[2],
        inspectionDate: row[3],
        status: row[4],
        notes: row[5],
        createdDate: row[6]
      }));

    return headers;
  } catch (e) {
    return { error: e.message };
  }
}

function getLastInspection(equipID) {
  try {
    const inspections = getInspectionsByEquip(equipID);
    if (!inspections || inspections.length === 0) return null;
    const completed = inspections.filter(i => i.status === 'Completed');
    if (completed.length === 0) return null;
    return completed.sort((a, b) => new Date(b.inspectionDate) - new Date(a.inspectionDate))[0];
  } catch (e) {
    return null;
  }
}

function getEquipmentWithStatus(areaID, dateStr) {
  try {
    const ss = getSSv();
    // Inline equipment load to avoid second sheet open via getEquipment()
    const eqSheet = ss.getSheetByName('Equipment');
    const eqData = eqSheet.getDataRange().getValues();
    const eqHeaders = eqData[0];
    const catIdx = eqHeaders.indexOf('Category');
    const equipment = eqData.slice(1).filter(r => r[0] && (!areaID || r[1] === areaID)).map(row => ({
      equipID: row[0], areaID: row[1], equipmentName: row[2],
      qrCode: row[3], assetTag: row[4], inspectionCycle: row[5],
      category: catIdx >= 0 ? (row[catIdx] || '') : '',
      status: row[7], createdDate: row[8]
    }));
    if (!equipment.length) return [];

    const targetDate = dateStr ? new Date(dateStr) : new Date();
    const targetDateStr = dateStr ? dateStr : Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const hSheet = ss.getSheetByName('InspectionHeader');
    const hData = hSheet.getDataRange().getValues();
    const inspMap = {};
    for (let i = 1; i < hData.length; i++) {
      const row = hData[i];
      if (!row[0]) continue;
      const eid = row[1];
      const iDate = row[3] ? (row[3] instanceof Date ? fmtD(row[3]) : String(row[3]).split('T')[0]) : '';
      if (!inspMap[eid]) inspMap[eid] = [];
      inspMap[eid].push({ inspectionID: row[0], status: row[4], date: iDate, createdDate: row[6] });
    }

    return equipment.map(eq => {
      const insp = inspMap[eq.equipID] || [];
      const todayInsp = insp.filter(i => i.date === targetDateStr && i.status === 'Completed');
      const lastInsp = insp.filter(i => i.status === 'Completed').sort((a, b) => new Date(b.date) - new Date(a.date))[0];

      const cycleDays = getCycleDays(eq.inspectionCycle);
      let equipStatus = 'Overdue'; // never inspected = overdue
      if (todayInsp.length > 0) {
        equipStatus = 'Completed Today';
      } else if (lastInsp) {
        equipStatus = 'Pending'; // inspected before
        const daysDiff = Math.floor((targetDate - new Date(lastInsp.date + 'T00:00:00')) / 86400000);
        if (daysDiff > cycleDays) equipStatus = 'Overdue';
      }

      return {
        ...eq,
        equipStatus,
        lastInspection: lastInsp ? lastInsp.date : null,
        lastInspectionFull: lastInsp ? lastInsp.createdDate : null
      };
    });
  } catch (e) {
    return { error: e.message };
  }
}


// ============================================================
// CHECKLIST — single call returns all equipment for a plant
// with status, area name, category — for the Checklist screen
// ============================================================

function getChecklistByPlant(plantID, dateStr) {
  try {
    var ss = getSSv();

    // 1. Load all areas for this plant
    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    var areaMap   = {};
    for (var a = 1; a < areaData.length; a++) {
      var ar = areaData[a];
      if (!ar[0]) continue;
      if (ar[1] === plantID) {
        areaMap[ar[0]] = { areaID: ar[0], areaName: ar[2], status: ar[4] };
      }
    }
    var areaIDs = Object.keys(areaMap);
    if (areaIDs.length === 0) return [];

    // 2. Load all equipment for those areas
    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var eqHeaders = eqData[0];
    var catColIdx = eqHeaders.indexOf('Category');
    var equipment = [];
    for (var e = 1; e < eqData.length; e++) {
      var row = eqData[e];
      if (!row[0]) continue;
      if (areaIDs.indexOf(row[1]) < 0) continue;
      if (row[7] !== 'Active') continue; // skip inactive
      equipment.push({
        equipID:        row[0],
        areaID:         row[1],
        equipmentName:  row[2],
        qrCode:         row[3],
        assetTag:       row[4],
        inspectionCycle: row[5],
        category:       catColIdx >= 0 ? (row[catColIdx] || '') : '',
        status:         row[7],
        areaName:       areaMap[row[1]] ? areaMap[row[1]].areaName : ''
      });
    }
    if (equipment.length === 0) return [];

    // 3. Load inspection headers — build status map
    var targetDate    = dateStr ? new Date(dateStr) : new Date();
    var targetDateStr = dateStr ? dateStr.split('T')[0] : Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    var inspMap = {};
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0]) continue;
      var eid   = hr[1];
      var iDate = hr[3] ? (hr[3] instanceof Date ? fmtD(hr[3]) : String(hr[3]).split('T')[0]) : '';
      if (!inspMap[eid]) inspMap[eid] = [];
      inspMap[eid].push({ status: hr[4], date: iDate });
    }

    // 4. Compute status for each equipment
    return equipment.map(function(eq) {
      var insp      = inspMap[eq.equipID] || [];
      var todayDone = insp.filter(function(i) { return i.date === targetDateStr && i.status === 'Completed'; });
      var lastDone  = insp.filter(function(i) { return i.status === 'Completed'; })
                         .sort(function(a,b) { return new Date(b.date) - new Date(a.date); })[0];

      var cycleDays = { 'Daily':1,'Weekly':7,'Bi-Weekly':14,'Semi-Monthly':15,'Monthly':30,'Quarterly':90,'Semi-Annual':180,'Annual':365 };
      var cycleLen = cycleDays[eq.inspectionCycle] || 1;
      var equipStatus = 'Pending';
      if (todayDone.length > 0) {
        equipStatus = 'Completed Today';
      } else if (!lastDone) {
        // Never inspected at all — overdue
        equipStatus = 'Overdue';
      } else {
        // days since last inspection
        var days = Math.floor((targetDate - new Date(lastDone.date + 'T00:00:00')) / 86400000);
        // Overdue if missed more than one full cycle
        // Daily: days >= 2 (missed yesterday), Weekly: days >= 8, etc.
        if (days > cycleLen) equipStatus = 'Overdue';
      }

      return {
        equipID:         eq.equipID,
        equipmentName:   eq.equipmentName,
        assetTag:        eq.assetTag,
        inspectionCycle: eq.inspectionCycle,
        category:        eq.category,
        areaID:          eq.areaID,
        areaName:        eq.areaName,
        equipStatus:     equipStatus,
        lastInspection:  lastDone ? lastDone.date : null
      };
    });
  } catch(e) {
    return { error: e.message };
  }
}

// ============================================================
// BATCH FUNCTIONS — Single calls replacing multiple round trips
// ============================================================

// Returns all areas for a plant WITH their stats computed in one call.
// Replaces: getAreas() + N x getAreaStats() calls
function getAreasWithStats(plantID) {
  try {
    var ss = getSSv();
    var today = new Date();
    // Use local date string (sheet timezone) not UTC to avoid off-by-one day issues
    var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // 1. Load all areas for this plant
    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    var areas = [];
    var areaIDs = [];
    for (var a = 1; a < areaData.length; a++) {
      var ar = areaData[a];
      if (!ar[0] || ar[1] !== plantID) continue;
      areas.push({ areaID: ar[0], plantID: ar[1], areaName: ar[2], description: ar[3], status: ar[4] });
      areaIDs.push(ar[0]);
    }
    if (areas.length === 0) return [];

    // 2. Load ALL equipment for these areas in ONE read
    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var eqHeaders = eqData[0];
    var catIdx = eqHeaders.indexOf('Category');
    var equipByArea = {};   // areaID -> [equipIDs]
    var equipIDs = {};      // equipID -> true
    areaIDs.forEach(function(id) { equipByArea[id] = []; });
    for (var e = 1; e < eqData.length; e++) {
      var row = eqData[e];
      if (!row[0] || areaIDs.indexOf(row[1]) < 0) continue;
      equipByArea[row[1]].push(row[0]);
      equipIDs[row[0]] = true;
    }

    // 3. Load ALL inspection headers in ONE read
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    var cycleDays = { 'Daily':1,'Weekly':7,'Bi-Weekly':14,'Semi-Monthly':15,'Monthly':30,'Quarterly':90,'Semi-Annual':180,'Annual':365 };

    // Build: for each equipID -> { completedToday: bool, lastDate: str, cycle: str }
    var inspMap = {};
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0] || !equipIDs[hr[1]]) continue;
      var eid   = hr[1];
      var iDate = '';
      if (hr[3]) {
        iDate = hr[3] instanceof Date
          ? fmtD(hr[3])
          : String(hr[3]).split('T')[0];
      }
      if (!inspMap[eid]) inspMap[eid] = { completedToday: false, lastDate: null };
      if (hr[4] === 'Completed') {
        if (iDate === todayStr) inspMap[eid].completedToday = true;
        if (!inspMap[eid].lastDate || iDate > inspMap[eid].lastDate) inspMap[eid].lastDate = iDate;
      }
    }

    // 4. Load equipment cycle for overdue calculation
    var eqCycleMap = {};
    for (var e2 = 1; e2 < eqData.length; e2++) {
      var er = eqData[e2];
      if (er[0] && equipIDs[er[0]]) eqCycleMap[er[0]] = er[5] || 'Daily';
    }

    // 5. Compute stats per area
    return areas.map(function(area) {
      var aEquips = equipByArea[area.areaID] || [];
      var total = aEquips.length;
      var completed = 0;
      aEquips.forEach(function(eid) {
        if (inspMap[eid] && inspMap[eid].completedToday) completed++;
      });
      var pending   = total - completed;
      var pct       = total > 0 ? Math.round((completed / total) * 100) : 0;
      return {
        areaID:        area.areaID,
        plantID:       area.plantID,
        areaName:      area.areaName,
        description:   area.description,
        status:        area.status,
        totalEquipment: total,
        completedToday: completed,
        pendingToday:   pending,
        progressPct:    pct
      };
    });
  } catch(e) {
    return { error: e.message };
  }
}

// Returns dashboard stats + per-plant stats all in one call.
// Replaces: getDashboardStats() + N x getPlantStats() calls
function getDashboardFull() {
  try {
    var ss = getSSv();
    var today = new Date();
    var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Load plants
    var plantSheet = ss.getSheetByName('Plants');
    var plantData  = plantSheet.getDataRange().getValues();
    var plants = [];
    var plantIDs = [];
    for (var p = 1; p < plantData.length; p++) {
      var pr = plantData[p];
      if (!pr[0]) continue;
      plants.push({ plantID: pr[0], plantName: pr[1], location: pr[2], status: pr[3] });
      plantIDs.push(pr[0]);
    }

    // Load areas
    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    var areaToPlant = {};  // areaID -> plantID
    var areasByPlant = {}; // plantID -> [areaIDs]
    plantIDs.forEach(function(id) { areasByPlant[id] = []; });
    var totalAreas = 0;
    for (var a = 1; a < areaData.length; a++) {
      var ar = areaData[a];
      if (!ar[0]) continue;
      areaToPlant[ar[0]] = ar[1];
      if (areasByPlant[ar[1]]) { areasByPlant[ar[1]].push(ar[0]); totalAreas++; }
    }

    // Load equipment
    var eqSheet  = ss.getSheetByName('Equipment');
    var eqData   = eqSheet.getDataRange().getValues();
    var eqToPlant = {};   // equipID -> plantID
    var eqByPlant = {};   // plantID -> count
    plantIDs.forEach(function(id) { eqByPlant[id] = 0; });
    var totalEquip = 0;
    for (var e = 1; e < eqData.length; e++) {
      var er = eqData[e];
      if (!er[0]) continue;
      var pid = areaToPlant[er[1]];
      if (pid) {
        eqToPlant[er[0]] = pid;
        eqByPlant[pid] = (eqByPlant[pid] || 0) + 1;
        totalEquip++;
      }
    }

    // Load inspection headers
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    var completedByPlant = {};
    plantIDs.forEach(function(id) { completedByPlant[id] = 0; });
    var totalCompletedToday = 0;
    var seenToday = {};
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0]) continue;
      var iDate = '';
      if (hr[3]) {
        iDate = hr[3] instanceof Date
          ? fmtD(hr[3])
          : String(hr[3]).split('T')[0];
      }
      if (iDate === todayStr && hr[4] === 'Completed' && !seenToday[hr[1]]) {
        seenToday[hr[1]] = true;
        var epid = eqToPlant[hr[1]];
        if (epid) { completedByPlant[epid] = (completedByPlant[epid] || 0) + 1; }
        totalCompletedToday++;
      }
    }

    // Build per-plant stats
    var plantStats = {};
    plantIDs.forEach(function(pid) {
      var eq   = eqByPlant[pid] || 0;
      var done = completedByPlant[pid] || 0;
      plantStats[pid] = {
        totalAreas:     (areasByPlant[pid] || []).length,
        totalEquipment: eq,
        completedToday: done,
        pendingToday:   Math.max(0, eq - done)
      };
    });

    return {
      plants: plants,
      dashboard: {
        totalPlants:    plants.filter(function(p) { return p.status === 'Active'; }).length,
        totalAreas:     totalAreas,
        totalEquipment: totalEquip,
        pendingToday:   Math.max(0, totalEquip - totalCompletedToday)
      },
      plantStats: plantStats
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ============================================================
// MY TASKS TODAY — flat list of pending/overdue equipment
// across all plants (or a single plant if specified)
// ============================================================
function getMyTasksToday(plantID) {
  try {
    var ss = getSSv();
    var today = new Date();
    var todayStr = fmtD(today);

    var plantSheet = ss.getSheetByName('Plants');
    var plantData  = plantSheet.getDataRange().getValues();
    var plantNames = {};
    var plantsList = [];
    for (var p = 1; p < plantData.length; p++) {
      var pr = plantData[p];
      if (!pr[0]) continue;
      plantNames[pr[0]] = pr[1];
      plantsList.push({ plantID: pr[0], plantName: pr[1] });
    }

    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    var areaToPlant = {};
    var areaNames   = {};
    for (var a = 1; a < areaData.length; a++) {
      var ar = areaData[a];
      if (!ar[0]) continue;
      areaToPlant[ar[0]] = ar[1];
      areaNames[ar[0]]   = ar[2];
    }

    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var eqHeaders = eqData[0];
    var catIdx = eqHeaders.indexOf('Category');
    var equipment = [];
    for (var e = 1; e < eqData.length; e++) {
      var er = eqData[e];
      if (!er[0]) continue;
      var aid = er[1];
      var pid = areaToPlant[aid];
      if (!pid) continue;
      if (plantID && pid !== plantID) continue;
      equipment.push({
        equipID: er[0], areaID: aid, equipmentName: er[2],
        assetTag: er[4], inspectionCycle: er[5],
        category: catIdx >= 0 ? (er[catIdx] || '') : '',
        areaName: areaNames[aid] || '', plantID: pid, plantName: plantNames[pid] || ''
      });
    }

    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    var lastInsp = {}; // equipID -> latest completed date string
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0] || hr[4] !== 'Completed') continue;
      var eid = hr[1];
      var iDate = hr[3] ? (hr[3] instanceof Date ? fmtD(hr[3]) : String(hr[3]).split('T')[0]) : '';
      if (!iDate) continue;
      if (!lastInsp[eid] || iDate > lastInsp[eid]) lastInsp[eid] = iDate;
    }

    var tasks = [];
    var overdueCount = 0, pendingCount = 0;
    equipment.forEach(function(eq) {
      var last = lastInsp[eq.equipID];
      if (last === todayStr) return; // already done today

      var cycleDays = getCycleDays(eq.inspectionCycle);
      var equipStatus, daysOverdue = null;
      if (!last) {
        equipStatus = 'Overdue'; // never inspected
      } else {
        var daysDiff = Math.floor((today - new Date(last + 'T00:00:00')) / 86400000);
        if (daysDiff > cycleDays) {
          equipStatus = 'Overdue';
          daysOverdue = daysDiff - cycleDays;
        } else {
          equipStatus = 'Pending';
        }
      }

      if (equipStatus === 'Overdue') overdueCount++; else pendingCount++;
      tasks.push({
        equipID: eq.equipID, equipmentName: eq.equipmentName, assetTag: eq.assetTag,
        inspectionCycle: eq.inspectionCycle, category: eq.category,
        areaID: eq.areaID, areaName: eq.areaName, plantID: eq.plantID, plantName: eq.plantName,
        equipStatus: equipStatus, lastInspection: last || null, daysOverdue: daysOverdue
      });
    });

    // Overdue first (longest overdue first), then Pending alphabetically
    tasks.sort(function(x, y) {
      if (x.equipStatus !== y.equipStatus) return x.equipStatus === 'Overdue' ? -1 : 1;
      if (x.equipStatus === 'Overdue') {
        var dx = x.daysOverdue === null ? 9999 : x.daysOverdue;
        var dy = y.daysOverdue === null ? 9999 : y.daysOverdue;
        return dy - dx;
      }
      return x.equipmentName.localeCompare(y.equipmentName);
    });

    return {
      tasks:  tasks,
      counts: { overdue: overdueCount, pending: pendingCount, total: tasks.length },
      plants: plantsList
    };
  } catch(e) {
    return { error: e.message };
  }
}

function getCycleDays(cycle) {
  const map = {
    'Daily': 1, 'Weekly': 7, 'Bi-Weekly': 14,
    'Semi-Monthly': 15, 'Monthly': 30, 'Quarterly': 90,
    'Semi-Annual': 180, 'Annual': 365
  };
  return map[cycle] || 1;
}


// ============================================================
// REPORT DATA — For equipment inspection report
// Returns structured data matching the report layout:
//   Equipment name, parameters as rows, dates as columns
//   Each cell: value, status (Normal/OutOfRange), technician
// ============================================================

function getEquipmentReport(equipID, dateFrom, dateTo) {
  try {
    var ss = getSSv();

    // Parse date range
    var from = new Date(dateFrom);
    var to   = new Date(dateTo);
    // Build array of date strings in range
    var dates = [];
    var cur = new Date(from);
    while (cur <= to) {
      dates.push(Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      cur.setDate(cur.getDate() + 1);
    }

    // Load equipment info
    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var eqHeaders = eqData[0];
    var catIdx = eqHeaders.indexOf('Category');
    var equipInfo = null;
    var areaID = null;
    for (var e = 1; e < eqData.length; e++) {
      if (eqData[e][0] === equipID) {
        equipInfo = {
          equipID:  eqData[e][0],
          areaID:   eqData[e][1],
          name:     eqData[e][2],
          assetTag: eqData[e][4],
          cycle:    eqData[e][5],
          category: catIdx >= 0 ? eqData[e][catIdx] : ''
        };
        areaID = eqData[e][1];
        break;
      }
    }
    if (!equipInfo) return { error: 'Equipment not found' };

    // Load area + plant info
    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    var areaName = ''; var plantID = '';
    for (var a = 1; a < areaData.length; a++) {
      if (areaData[a][0] === areaID) {
        areaName = areaData[a][2]; plantID = areaData[a][1]; break;
      }
    }
    var plantSheet = ss.getSheetByName('Plants');
    var plantData  = plantSheet.getDataRange().getValues();
    var plantName = '';
    for (var p = 1; p < plantData.length; p++) {
      if (plantData[p][0] === plantID) { plantName = plantData[p][1]; break; }
    }

    // Load templates (parameter definitions)
    var tmplSheet = ss.getSheetByName('Templates');
    var tmplData  = tmplSheet.getDataRange().getValues();
    var params = []; // ordered list of params
    for (var t = 1; t < tmplData.length; t++) {
      var tr = tmplData[t];
      if (!tr[0] || tr[1] !== equipID) continue;
      params.push({
        templateID:    tr[0],
        parameterName: tr[2],
        unit:          tr[3],
        minValue:      tr[4],
        maxValue:      tr[5],
        sortOrder:     tr[6],
        type:          tr[7]
      });
    }
    params.sort(function(a,b){ return (a.sortOrder||0)-(b.sortOrder||0); });

    // Load technicians for name lookup
    var techSheet = ss.getSheetByName('Technicians');
    var techData  = techSheet.getDataRange().getValues();
    var techMap   = {};
    for (var tc = 1; tc < techData.length; tc++) {
      if (techData[tc][0]) techMap[techData[tc][0]] = techData[tc][1];
    }

    // Load inspection headers for this equipment in date range
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    // inspByDate: dateStr -> { inspectionID, technicianID, techName, status }
    var inspByDate = {};
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0] || hr[1] !== equipID || hr[4] !== 'Completed') continue;
      var iDate = hr[3] ? (hr[3] instanceof Date ? fmtD(hr[3]) : String(hr[3]).split('T')[0]) : '';
      if (dates.indexOf(iDate) < 0) continue;
      // Keep most recent inspection per date
      if (!inspByDate[iDate] || hr[0] > inspByDate[iDate].inspectionID) {
        inspByDate[iDate] = {
          inspectionID: hr[0],
          technicianID: hr[2],
          techName:     techMap[hr[2]] || hr[2],
          notes:        hr[5] || ''
        };
      }
    }

    // Load inspection details for all relevant inspectionIDs
    var dSheet = ss.getSheetByName('InspectionDetails');
    var dData  = dSheet.getDataRange().getValues();
    // detailMap: inspectionID -> { paramName -> { value, status } }
    var inspIDs = {};
    Object.keys(inspByDate).forEach(function(d) { inspIDs[inspByDate[d].inspectionID] = true; });
    var detailMap = {};
    for (var d = 1; d < dData.length; d++) {
      var dr = dData[d];
      if (!dr[0] || !inspIDs[dr[1]]) continue;
      if (!detailMap[dr[1]]) detailMap[dr[1]] = {};
      detailMap[dr[1]][dr[2]] = { value: dr[3], status: dr[4] };
    }

    // Build parameter rows
    // Each row: { paramLabel, range, cells: [ {value, status} per date ] }
    var rows = params.map(function(param) {
      var rangeLabel = '';
      if (param.type === 'numeric' && param.minValue !== '' && param.maxValue !== '') {
        rangeLabel = '(' + param.minValue + '-' + param.maxValue;
        if (param.unit) rangeLabel += ' ' + param.unit;
        rangeLabel += ')';
      } else if (param.unit) {
        rangeLabel = '(' + param.unit + ')';
      }

      var cells = dates.map(function(dateStr) {
        var insp = inspByDate[dateStr];
        if (!insp) return { value: null, status: null }; // no inspection that day
        var detail = detailMap[insp.inspectionID] ? detailMap[insp.inspectionID][param.parameterName] : null;
        if (!detail) return { value: '-', status: 'Normal' };
        return { value: detail.value, status: detail.status };
      });

      return {
        parameterName: param.parameterName,
        rangeLabel:    rangeLabel,
        unit:          param.unit,
        minValue:      param.minValue,
        maxValue:      param.maxValue,
        type:          param.type,
        cells:         cells
      };
    });

    // Build footer rows: Technician, Status, Out of Spec count per date
    var footerTech    = dates.map(function(d) { return inspByDate[d] ? inspByDate[d].techName : null; });
    var footerStatus  = dates.map(function(d) {
      if (!inspByDate[d]) return null;
      var insp = inspByDate[d];
      var details = detailMap[insp.inspectionID] || {};
      var hasOOR = Object.values(details).some(function(v){
  return v.status === 'OutOfRange' ||
         v.status === 'OutOfSpec';
});
      return hasOOR ? 'Attention' : 'Normal';
    });
    var footerOutSpec = dates.map(function(d) {
      if (!inspByDate[d]) return null;
      var insp = inspByDate[d];
      var details = detailMap[insp.inspectionID] || {};
      return Object.values(details).filter(function(v){
  return v.status === 'OutOfRange' ||
         v.status === 'OutOfSpec';
}).length;
    });

    // Load photos for all inspections in range (safe — never breaks report)
    var footerPhotos = dates.map(function(d){ return []; });
    try {
      var inspIdsForPhotos = Object.values(inspByDate).map(function(i){ return i.inspectionID; });
      var photosByInsp = getPhotosByInspectionIds(inspIdsForPhotos);
      footerPhotos = dates.map(function(d) {
        if (!inspByDate[d]) return [];
        return photosByInsp[String(inspByDate[d].inspectionID)] || [];
      });
    } catch(pe) { /* photos failed silently */ }

    return {
      equipInfo: equipInfo,
      areaName:  areaName,
      plantName: plantName,
      dates:     dates,
      params:    rows,
      footer: {
        tech:     footerTech,
        status:   footerStatus,
        outOfSpec: footerOutSpec,
        photos:   footerPhotos,
        notes:    dates.map(function(d){ return inspByDate[d] ? inspByDate[d].notes : null; })
      }
    };
  } catch(e) {
    return { error: e.message };
  }
}

// Get all equipment for an area — for the report screen selector
function getEquipmentForReport(areaID) {
  try {
    var ss = getSSv();
    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var eqHeaders = eqData[0];
    var catIdx = eqHeaders.indexOf('Category');
    var result = [];
    for (var e = 1; e < eqData.length; e++) {
      var row = eqData[e];
      if (!row[0] || row[7] !== 'Active') continue;
      if (areaID && row[1] !== areaID) continue;
      result.push({
        equipID:  row[0],
        areaID:   row[1],
        name:     row[2],
        assetTag: row[4],
        cycle:    row[5],
        category: catIdx >= 0 ? row[catIdx] : ''
      });
    }
    return result;
  } catch(e) {
    return { error: e.message };
  }
}

// ============================================================
// DASHBOARD REPORT FUNCTIONS
// ============================================================

function getDashboardReport(year) {
  try {
    var ss = getSSv();
    var tz  = Session.getScriptTimeZone();
    var yr  = parseInt(year) || new Date().getFullYear();
    var today    = new Date();
    var todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');

    // ── 1. Plants ─────────────────────────────────────────────
    var plantMap = {};
    var plantSheet = ss.getSheetByName('Plants');
    var plantData  = plantSheet.getDataRange().getValues();
    for (var p = 1; p < plantData.length; p++) {
      if (plantData[p][0]) plantMap[plantData[p][0]] = plantData[p][1];
    }

    // ── 2. Areas ──────────────────────────────────────────────
    var areas = [];
    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    for (var a = 1; a < areaData.length; a++) {
      var ar = areaData[a];
      if (!ar[0]) continue;
      areas.push({
        areaID:    ar[0],
        plantID:   ar[1],
        areaName:  ar[2],
        plantName: plantMap[ar[1]] || ar[1],
        status:    ar[4]
      });
    }

    // ── 3. Equipment → Area map ───────────────────────────────
    // equipID -> areaID, and areaID -> cycle
    var equipToArea = {};
    var areaCycle   = {}; // areaID -> cycle
    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var eqHeaders = eqData[0];
    var statusColIdx = eqHeaders.indexOf('Status');
    if (statusColIdx < 0) statusColIdx = 7;

    for (var e = 1; e < eqData.length; e++) {
      var er = eqData[e];
      if (!er[0]) continue;
      equipToArea[er[0]] = er[1]; // equipID -> areaID
      if (!areaCycle[er[1]]) areaCycle[er[1]] = er[5] || 'Daily';
    }

    // ── 4. Inspections for this year ──────────────────────────
    // Build: inspByArea[areaID][dateStr] = true
    var inspByArea = {};
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();

    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0] || hr[4] !== 'Completed') continue;
      var equipID = hr[1];
      var aID     = equipToArea[equipID];
      if (!aID) continue;

      var iDate = hr[3] instanceof Date
        ? fmtD(hr[3])
        : String(hr[3]).split('T')[0];
      if (!iDate || iDate.substring(0,4) !== String(yr)) continue;

      if (!inspByArea[aID]) inspByArea[aID] = {};
      inspByArea[aID][iDate] = true;
    }

    // ── 5. Compute per-area per-month stats ───────────────────
    var result = areas.map(function(area) {
      var cycle  = areaCycle[area.areaID] || 'Daily';
      var aInsp  = inspByArea[area.areaID] || {};
      var months = [];

      for (var m = 0; m < 12; m++) {
        var monthStart = yr + '-' + pad2(m + 1) + '-01';

        // Future month — no data yet
        if (monthStart > todayStr) {
          months.push({ status: 'upcoming', completed: 0, total: 0, missedDays: [] });
          continue;
        }

        var requiredDays  = getRequiredDays(yr, m, cycle);
        var completedDays = [];
        var missedDays    = [];

        requiredDays.forEach(function(dayStr) {
          if (dayStr > todayStr) return; // skip future days within current month
          if (aInsp[dayStr]) completedDays.push(parseInt(dayStr.substring(8)));
          else               missedDays.push(parseInt(dayStr.substring(8)));
        });

        var completed = completedDays.length;
        var total     = missedDays.length + completed;
        var status    = total === 0       ? 'upcoming'
                      : completed === total ? 'completed'
                      : completed === 0     ? 'none'
                      : 'partial';

        months.push({
          status:        status,
          completed:     completed,
          total:         total,
          missedDays:    missedDays,
          completedDays: completedDays
        });
      }

      return {
        areaID:    area.areaID,
        areaName:  area.areaName,
        plantID:   area.plantID,
        plantName: area.plantName,
        cycle:     cycle,
        months:    months
      };
    });

    return { year: yr, areas: result };
  } catch(e) {
    return { error: e.message + ' | ' + e.stack };
  }
}

function getRequiredDays(year, month, cycle) {
  var days = [];
  var daysInMonth = new Date(year, month + 1, 0).getDate();

  if (cycle === 'Daily') {
    for (var d = 1; d <= daysInMonth; d++) {
      days.push(year + '-' + pad2(month+1) + '-' + pad2(d));
    }
  } else if (cycle === 'Weekly') {
    // Every 7 days starting from day 1
    for (var d = 1; d <= daysInMonth; d += 7) {
      days.push(year + '-' + pad2(month+1) + '-' + pad2(d));
    }
  } else if (cycle === 'Bi-Weekly') {
    for (var d = 1; d <= daysInMonth; d += 14) {
      days.push(year + '-' + pad2(month+1) + '-' + pad2(d));
    }
  } else if (cycle === 'Semi-Monthly') {
    days.push(year + '-' + pad2(month+1) + '-01');
    days.push(year + '-' + pad2(month+1) + '-15');
  } else {
    // Monthly, Quarterly, Semi-Annual, Annual — once per period, use day 1
    days.push(year + '-' + pad2(month+1) + '-01');
  }
  return days;
}

function pad2(n) { return n < 10 ? '0' + n : String(n); }

// Get day-by-day detail for a specific area+month (for modal)
function getAreaMonthDetail(areaID, year, month) {
  try {
    var ss = getSSv();
    var tz  = Session.getScriptTimeZone();
    var yr  = parseInt(year);
    var mo  = parseInt(month); // 0-based

    // Get equipment for area
    var eqSheet = ss.getSheetByName('Equipment');
    var eqData  = eqSheet.getDataRange().getValues();
    var equips  = [];
    var cycle   = 'Daily';
    for (var e = 1; e < eqData.length; e++) {
      if (eqData[e][1] === areaID && eqData[e][7] === 'Active') {
        equips.push(eqData[e][0]);
        cycle = eqData[e][5] || 'Daily';
      }
    }

    // Get inspections for this month
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    var inspByDay = {}; // 'YYYY-MM-DD' -> [techNames]

    // Get tech names
    var techSheet = ss.getSheetByName('Technicians');
    var techData  = techSheet.getDataRange().getValues();
    var techMap   = {};
    for (var t = 1; t < techData.length; t++) {
      if (techData[t][0]) techMap[techData[t][0]] = techData[t][1];
    }

    var monthPrefix = yr + '-' + pad2(mo + 1);
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0] || hr[4] !== 'Completed') continue;
      if (equips.indexOf(hr[1]) < 0) continue;
      var iDate = hr[3] instanceof Date
        ? fmtD(hr[3])
        : String(hr[3]).split('T')[0];
      if (!iDate || iDate.substring(0,7) !== monthPrefix) continue;
      if (!inspByDay[iDate]) inspByDay[iDate] = [];
      var tech = techMap[hr[2]] || hr[2] || '';
      if (tech && inspByDay[iDate].indexOf(tech) < 0) inspByDay[iDate].push(tech);
    }

    var today    = new Date();
    var todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
    var daysInMonth = new Date(yr, mo + 1, 0).getDate();
    var requiredDays = getRequiredDays(yr, mo, cycle);
    var reqSet = {};
    requiredDays.forEach(function(d){ reqSet[d] = true; });

    var days = [];
    for (var d = 1; d <= daysInMonth; d++) {
      var dateStr = yr + '-' + pad2(mo + 1) + '-' + pad2(d);
      var isRequired = !!reqSet[dateStr];
      var isFuture   = dateStr > todayStr;
      var techs      = inspByDay[dateStr] || [];
      var done       = techs.length > 0;
      days.push({
        day: d, dateStr: dateStr,
        required: isRequired,
        future: isFuture,
        done: done,
        techs: techs
      });
    }

    var completedCount = requiredDays.filter(function(d){
      return inspByDay[d] && inspByDay[d].length > 0;
    }).length;

    return {
      areaID: areaID, year: yr, month: mo, cycle: cycle,
      daysInMonth: daysInMonth,
      requiredCount: requiredDays.length,
      completedCount: completedCount,
      days: days
    };
  } catch(e) {
    return { error: e.message };
  }
}
function saveInspection(data) {
  try {
    const ss = getSSv();
    const hSheet = ss.getSheetByName('InspectionHeader');
    const dSheet = ss.getSheetByName('InspectionDetails');
    const now = new Date().toISOString();

    let inspectionID = data.inspectionID;

    if (inspectionID) {
      // Update existing
      const hData = hSheet.getDataRange().getValues();
      for (let i = 1; i < hData.length; i++) {
        if (hData[i][0] === inspectionID) {
          hSheet.getRange(i + 1, 1, 1, 7).setValues([[
            inspectionID, data.equipID, data.technicianID,
            data.inspectionDate, data.status || 'Completed',
            data.notes || '', now
          ]]);
          break;
        }
      }
      // Clear old details
      const dData = dSheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let i = dData.length - 1; i >= 1; i--) {
        if (dData[i][1] === inspectionID) rowsToDelete.push(i + 1);
      }
      rowsToDelete.forEach(r => dSheet.deleteRow(r));
    } else {
      inspectionID = generateID('INS', hSheet);
      hSheet.appendRow([
        inspectionID, data.equipID, data.technicianID,
        data.inspectionDate, data.status || 'Completed',
        data.notes || '', now
      ]);
    }

    // Save details
    if (data.details && data.details.length > 0) {
      const detailRows = data.details.map((d, idx) => [
        generateDetailID(inspectionID, idx),
        inspectionID, d.parameterName, d.value, d.status || 'Normal', now
      ]);
      dSheet.getRange(dSheet.getLastRow() + 1, 1, detailRows.length, 6).setValues(detailRows);
    }

    // Save signature if present
    if (data.signature) {
      const sigSheet = ss.getSheetByName('Signatures');
      const sigID = generateID('SIG', sigSheet);
      sigSheet.appendRow([sigID, inspectionID, data.signature, data.technicianID, now]);
    }

    // Out-of-spec email alert (category-routed, silent if no mapping)
    if ((data.status || 'Completed') === 'Completed' && data.details && data.details.length) {
      try {
        var oosDetails = data.details.filter(function(d) {
          return d.status === 'OutOfRange' || d.status === 'OutOfSpec';
        });
        if (oosDetails.length) sendOutOfSpecAlert(ss, data, oosDetails);
      } catch(alertErr) {
        // Never fail the save because of alert issues
      }
    }

    return { success: true, inspectionID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getInspectionDetails(inspectionID) {
  try {
    const ss = getSSv();
    const dSheet = ss.getSheetByName('InspectionDetails');
    const dData = dSheet.getDataRange().getValues();
    return dData.slice(1)
      .filter(r => r[0] && r[1] === inspectionID)
      .map(row => ({
        detailID: row[0],
        inspectionID: row[1],
        parameterName: row[2],
        value: row[3],
        status: row[4],
        createdDate: row[5]
      }));
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// DASHBOARD STATS
// ============================================================

function getDashboardStats() {
  try {
    const ss = getSSv();
    const plants = getPlants();
    const areas = getAreas();
    const equipment = getEquipment();

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const hSheet = ss.getSheetByName('InspectionHeader');
    const hData = hSheet.getDataRange().getValues();

    const todayCompleted = new Set();
    for (let i = 1; i < hData.length; i++) {
      const row = hData[i];
      if (!row[0]) continue;
      const iDate = row[3] ? (row[3] instanceof Date ? fmtD(row[3]) : String(row[3]).split('T')[0]) : '';
      if (iDate === today && row[4] === 'Completed') {
        todayCompleted.add(row[1]);
      }
    }

    const totalEquip = Array.isArray(equipment) ? equipment.length : 0;
    const pendingToday = totalEquip - todayCompleted.size;

    return {
      totalPlants: Array.isArray(plants) ? plants.filter(p => p.status === 'Active').length : 0,
      totalAreas: Array.isArray(areas) ? areas.length : 0,
      totalEquipment: totalEquip,
      pendingToday: Math.max(0, pendingToday)
    };
  } catch (e) {
    return { error: e.message };
  }
}

function getPlantStats(plantID) {
  try {
    const areas = getAreas(plantID);
    if (!Array.isArray(areas)) return { error: 'Failed to load areas' };

    const ss = getSSv();
    const eSheet = ss.getSheetByName('Equipment');
    const hSheet = ss.getSheetByName('InspectionHeader');
    const eData = eSheet.getDataRange().getValues();
    const hData = hSheet.getDataRange().getValues();

    const areaIDs = areas.map(a => a.areaID);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const equipByArea = {};
    let totalEquip = 0;
    for (let i = 1; i < eData.length; i++) {
      const row = eData[i];
      if (!row[0]) continue;
      if (areaIDs.includes(row[1])) {
        if (!equipByArea[row[1]]) equipByArea[row[1]] = 0;
        equipByArea[row[1]]++;
        totalEquip++;
      }
    }

    const completedToday = new Set();
    for (let i = 1; i < hData.length; i++) {
      const row = hData[i];
      if (!row[0]) continue;
      const iDate = row[3] ? (row[3] instanceof Date ? fmtD(row[3]) : String(row[3]).split('T')[0]) : '';
      if (iDate === today && row[4] === 'Completed') {
        completedToday.add(row[1]);
      }
    }

    // Filter by plant's equipment
    const allEquipIDs = new Set();
    for (let i = 1; i < eData.length; i++) {
      if (areaIDs.includes(eData[i][1])) allEquipIDs.add(eData[i][0]);
    }
    const completedCount = [...completedToday].filter(id => allEquipIDs.has(id)).length;

    return {
      totalAreas: areas.length,
      totalEquipment: totalEquip,
      completedToday: completedCount,
      pendingToday: Math.max(0, totalEquip - completedCount)
    };
  } catch (e) {
    return { error: e.message };
  }
}

function getAreaStats(areaID) {
  try {
    const ss = getSSv();
    const equipment = getEquipment(areaID);
    if (!Array.isArray(equipment)) return { error: 'Failed' };

    const hSheet = ss.getSheetByName('InspectionHeader');
    const hData = hSheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const completedToday = new Set();
    for (let i = 1; i < hData.length; i++) {
      const row = hData[i];
      if (!row[0]) continue;
      const iDate = row[3] ? (row[3] instanceof Date ? fmtD(row[3]) : String(row[3]).split('T')[0]) : '';
      if (iDate === today && row[4] === 'Completed') completedToday.add(row[1]);
    }

    const equipIDs = equipment.map(e => e.equipID);
    const completed = equipIDs.filter(id => completedToday.has(id)).length;

    return {
      totalEquipment: equipment.length,
      completedToday: completed,
      pendingToday: equipment.length - completed,
      progressPct: equipment.length > 0 ? Math.round((completed / equipment.length) * 100) : 0
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// TECHNICIAN MANAGEMENT (Admin)
// ============================================================

function getAllTechnicians() {
  try {
    const ss = getSSv();
    const sheet = ss.getSheetByName('Technicians');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).filter(r => r[0]).map(row => ({
      techID: row[0],
      name: row[1],
      pin: row[2],
      role: row[3],
      status: row[4],
      createdDate: row[5]
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function saveTechnician(tech) {
  try {
    invalidateTechCache();
    const ss = getSSv();
    const sheet = ss.getSheetByName('Technicians');
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    if (tech.techID) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === tech.techID) {
          sheet.getRange(i + 1, 1, 1, 6).setValues([[
            tech.techID, tech.name, tech.pin, tech.role || 'Tech', tech.status || 'Active', data[i][5]
          ]]);
          return { success: true };
        }
      }
    }
    const newID = generateID('T', sheet);
    sheet.appendRow([newID, tech.name, tech.pin, tech.role || 'Tech', tech.status || 'Active', now]);
    return { success: true, techID: newID };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteTechnician(techID) {
  try {
    invalidateTechCache();
    const ss = getSSv();
    const sheet = ss.getSheetByName('Technicians');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === techID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Technician not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// REPORT DATA
// ============================================================

function getReportData(year) {
  try {
    const ss = getSSv();
    const plants = getPlants();
    const areas = getAreas();
    const equipment = getEquipment();
    const hSheet = ss.getSheetByName('InspectionHeader');
    const hData = hSheet.getDataRange().getValues();

    const inspByEquip = {};
    for (let i = 1; i < hData.length; i++) {
      const row = hData[i];
      if (!row[0]) continue;
      const iDate = row[3] ? (row[3] instanceof Date ? row[3] : new Date(row[3])) : null;
      if (!iDate || iDate.getFullYear() !== parseInt(year)) continue;
      const eid = row[1];
      const month = iDate.getMonth(); // 0-indexed
      if (!inspByEquip[eid]) inspByEquip[eid] = {};
      if (row[4] === 'Completed') inspByEquip[eid][month] = true;
    }

    const areaMap = {};
    areas.forEach(a => areaMap[a.areaID] = a);
    const plantMap = {};
    plants.forEach(p => plantMap[p.plantID] = p);

    return equipment.map(eq => {
      const area = areaMap[eq.areaID] || {};
      const plant = plantMap[area.plantID] || {};
      const months = {};
      for (let m = 0; m < 12; m++) {
        const now = new Date();
        if (new Date(year, m, 1) > now) {
          months[m] = 'upcoming';
        } else {
          months[m] = inspByEquip[eq.equipID] && inspByEquip[eq.equipID][m] ? 'completed' : 'missed';
        }
      }
      return {
        equipID: eq.equipID,
        equipmentName: eq.equipmentName,
        areaName: area.areaName || '',
        plantName: plant.plantName || '',
        plantID: area.plantID || '',
        cycle: eq.inspectionCycle,
        months
      };
    });
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function generateID(prefix, sheet) {
  const data = sheet.getDataRange().getValues();
  const existing = data.slice(1)
    .map(r => r[0])
    .filter(id => String(id).startsWith(prefix))
    .map(id => parseInt(String(id).replace(prefix, '')) || 0);
  const next = existing.length > 0 ? Math.max(...existing) + 1 : 1;
  return prefix + String(next).padStart(3, '0');
}

function generateDetailID(inspectionID, idx) {
  return 'D-' + inspectionID + '-' + (idx + 1);
}

function verifyAdminPin(pin) {
  return String(pin) === String(CONFIG.ADMIN_PIN);
}




// ============================================================
// QR DEEP LINK BATCH — one call returns everything needed
// when a QR code is scanned: plants, area, equipment list with
// status, templates for the target equipment, and technicians.
// Replaces 4 separate round trips.
// ============================================================
function getQRDeepLink(areaID, equipID, dateStr) {
  try {
    var result = {};

    // 1. Plants (for nav context)
    result.plants = getPlants();
    if (result.plants && result.plants.error) return { error: result.plants.error };

    // 2. Areas with stats for the plant that owns this area
    var ss = getSSv();
    var areaSheet = ss.getSheetByName('Areas');
    var areaData  = areaSheet.getDataRange().getValues();
    var plantID   = '';
    for (var a = 1; a < areaData.length; a++) {
      if (areaData[a][0] === areaID) { plantID = areaData[a][1]; break; }
    }
    if (!plantID) return { error: 'Area not found: ' + areaID };
    result.plantID = plantID;
    result.areas   = getAreasWithStats(plantID);
    if (result.areas && result.areas.error) return { error: result.areas.error };

    // 3. Equipment with status for the area
    result.equipWithStatus = getEquipmentWithStatus(areaID, dateStr);
    if (result.equipWithStatus && result.equipWithStatus.error) return { error: result.equipWithStatus.error };

    // 4. Templates for the target equipment (if specified)
    result.templates = equipID ? getTemplates(equipID) : [];

    // 5. Technicians (cached)
    result.technicians = getTechnicians();

    return result;
  } catch(e) {
    return { error: e.message };
  }
}

// ============================================================
// PARAMETER CHART DATA
// Returns numeric readings for one parameter over a date range
// ============================================================
function getParameterChart(equipID, parameterName, dateFrom, dateTo) {
  try {
    var ss   = getSSv();
    var tz   = Session.getScriptTimeZone();
    var from = new Date(dateFrom);
    var to   = new Date(dateTo);

    // Build date array
    var dates = [];
    var cur = new Date(from);
    while (cur <= to) {
      dates.push(Utilities.formatDate(cur, tz, 'yyyy-MM-dd'));
      cur.setDate(cur.getDate() + 1);
    }

    // Get template info for this parameter (min/max/unit)
    var tmplSheet = ss.getSheetByName('Templates');
    var tmplData  = tmplSheet.getDataRange().getValues();
    var unit = '', minVal = null, maxVal = null;
    for (var t = 1; t < tmplData.length; t++) {
      if (tmplData[t][1] === equipID && tmplData[t][2] === parameterName && tmplData[t][7] === 'numeric') {
        unit   = tmplData[t][3] || '';
        minVal = tmplData[t][4] !== '' ? Number(tmplData[t][4]) : null;
        maxVal = tmplData[t][5] !== '' ? Number(tmplData[t][5]) : null;
        break;
      }
    }

    // Load inspection headers for this equipment in range
    var hSheet = ss.getSheetByName('InspectionHeader');
    var hData  = hSheet.getDataRange().getValues();
    var inspByDate = {};
    for (var h = 1; h < hData.length; h++) {
      var hr = hData[h];
      if (!hr[0] || hr[1] !== equipID || hr[4] !== 'Completed') continue;
      var iDate = hr[3] instanceof Date
        ? fmtD(hr[3])
        : String(hr[3]).split('T')[0];
      if (dates.indexOf(iDate) < 0) continue;
      if (!inspByDate[iDate] || hr[0] > inspByDate[iDate]) inspByDate[iDate] = hr[0];
    }

    // Load details for matched inspections
    var inspIDs = Object.values(inspByDate);
    var dSheet  = ss.getSheetByName('InspectionDetails');
    var dData   = dSheet.getDataRange().getValues();
    var valueMap = {}; // inspectionID -> value
    for (var d = 1; d < dData.length; d++) {
      var dr = dData[d];
      if (!dr[0]) continue;
      if (inspIDs.indexOf(dr[1]) >= 0 && dr[2] === parameterName) {
        var num = parseFloat(dr[3]);
        if (!isNaN(num)) valueMap[dr[1]] = num;
      }
    }

    // Build result arrays aligned to dates
    var resultDates  = [];
    var resultValues = [];
    dates.forEach(function(d) {
      var inspID = inspByDate[d];
      if (inspID && valueMap[inspID] !== undefined) {
        resultDates.push(d);
        resultValues.push(valueMap[inspID]);
      }
    });

    return {
      dates:     resultDates,
      values:    resultValues,
      unit:      unit,
      minValue:  minVal,
      maxValue:  maxVal,
      parameterName: parameterName
    };
  } catch(e) {
    return { error: e.message };
  }
}



// ============================================================
// PROJECT MANAGEMENT
// Projects sheet: ProjectID | PlantID | ProjectName | Category | Description |
//   Owner | Status | Priority | StartDate | DueDate | Progress | Notes | CreatedDate | UpdatedDate
// ProjectTasks sheet: TaskID | ProjectID | TaskName | Owner | Status | Priority |
//   StartDate | DueDate | Progress | Notes | CreatedDate | UpdatedDate
// "Progress" on Projects is a manual fallback used only when the project has
// zero tasks; once it has tasks, progress is the average of its tasks' progress.
// ============================================================

var PM_CATEGORIES = ['Capital Project', 'Preventive Maintenance', 'Upgrade / Retrofit', 'Compliance', 'General'];
var PM_STATUSES   = ['Backlog', 'Planned', 'In Progress', 'On Hold', 'Completed'];
var PM_PRIORITIES = ['Low', 'Medium', 'High'];

function ensureProjectsSheet() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('Projects');
  if (!sheet) {
    sheet = ss.insertSheet('Projects');
    var headers = ['ProjectID','PlantID','ProjectName','Category','Description','Owner','Status','Priority','StartDate','DueDate','Progress','Notes','CreatedDate','UpdatedDate'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureProjectTasksSheet() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('ProjectTasks');
  if (!sheet) {
    sheet = ss.insertSheet('ProjectTasks');
    var headers = ['TaskID','ProjectID','TaskName','Owner','Status','Priority','StartDate','DueDate','Progress','Notes','SortOrder','CreatedDate','UpdatedDate'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getProjectMeta() {
  var plants = getPlants();
  return {
    plants: (plants && !plants.error) ? plants : [],
    categories: PM_CATEGORIES,
    statuses: PM_STATUSES,
    priorities: PM_PRIORITIES
  };
}

// Returns all projects (optionally filtered by plant) with computed progress + taskCount
function getProjects(plantID) {
  try {
    var ss = getSSv();
    var pSheet = ensureProjectsSheet();
    var pData = pSheet.getDataRange().getValues();

    var plants = getPlants();
    var plantNames = {};
    (plants || []).forEach(function(p) { plantNames[p.plantID] = p.plantName; });

    // Pre-aggregate task progress per project
    var tSheet = ensureProjectTasksSheet();
    var tData = tSheet.getDataRange().getValues();
    var agg = {}; // projectID -> { sum, count }
    for (var t = 1; t < tData.length; t++) {
      var tr = tData[t];
      if (!tr[0]) continue;
      var pid = tr[1];
      if (!agg[pid]) agg[pid] = { sum: 0, count: 0 };
      agg[pid].sum += Number(tr[8]) || 0;
      agg[pid].count++;
    }

    if (pData.length <= 1) return [];

    var rows = pData.slice(1).filter(function(r) { return r[0]; });
    if (plantID) rows = rows.filter(function(r) { return r[1] === plantID; });

    return rows.map(function(r) {
      var projectID = r[0];
      var a = agg[projectID];
      var taskCount = a ? a.count : 0;
      var progress = taskCount > 0 ? Math.round(a.sum / taskCount) : (Number(r[10]) || 0);

      return {
        projectID: projectID, plantID: r[1], plantName: plantNames[r[1]] || r[1],
        projectName: r[2], category: r[3], description: r[4] || '',
        owner: r[5] || '', status: r[6] || 'Backlog', priority: r[7] || 'Medium',
        startDate: r[8] instanceof Date ? fmtD(r[8]) : (r[8] || ''),
        dueDate:   r[9] instanceof Date ? fmtD(r[9])  : (r[9]  || ''),
        progress: progress, manualProgress: Number(r[10]) || 0,
        notes: r[11] || '', taskCount: taskCount,
        createdDate: r[12], updatedDate: r[13]
      };
    }).sort(function(a, b) { return (a.dueDate || '9999').localeCompare(b.dueDate || '9999'); });
  } catch(e) {
    return { error: e.message };
  }
}

function saveProject(project) {
  try {
    var sheet = ensureProjectsSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();

    var row = [
      project.projectID || '', project.plantID || '', project.projectName || '',
      project.category || '', project.description || '', project.owner || '',
      project.status || 'Backlog', project.priority || 'Medium',
      project.startDate || '', project.dueDate || '',
      (project.progress !== undefined && project.progress !== null) ? project.progress : 0,
      project.notes || '', '', now
    ];

    if (project.projectID) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === project.projectID) {
          row[12] = data[i][12]; // preserve CreatedDate
          sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
          return { success: true, projectID: project.projectID };
        }
      }
    }

    var newID = generateID('PRJ', sheet);
    row[0] = newID;
    row[12] = now;
    sheet.appendRow(row);
    return { success: true, projectID: newID };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function deleteProject(projectID) {
  try {
    var pSheet = ensureProjectsSheet();
    var pData = pSheet.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      if (pData[i][0] === projectID) { pSheet.deleteRow(i + 1); break; }
    }
    // Cascade delete tasks
    var tSheet = ensureProjectTasksSheet();
    var tData = tSheet.getDataRange().getValues();
    for (var j = tData.length - 1; j >= 1; j--) {
      if (tData[j][1] === projectID) tSheet.deleteRow(j + 1);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function updateProjectStatus(projectID, status) {
  try {
    var sheet = ensureProjectsSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectID) {
        sheet.getRange(i + 1, 7).setValue(status);  // Status column
        sheet.getRange(i + 1, 14).setValue(now);    // UpdatedDate
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Tasks for one project (used when expanding a project row in the table)
function getProjectTasks(projectID) {
  try {
    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).filter(function(r) { return r[0] && r[1] === projectID; }).map(function(r) {
      return {
        taskID: r[0], projectID: r[1], taskName: r[2], owner: r[3] || '',
        status: r[4] || 'Backlog', priority: r[5] || 'Medium',
        startDate: r[6] instanceof Date ? fmtD(r[6]) : (r[6] || ''),
        dueDate:   r[7] instanceof Date ? fmtD(r[7])  : (r[7]  || ''),
        progress: Number(r[8]) || 0, notes: r[9] || '',
        sortOrder: Number(r[10]) || 0,
        createdDate: r[11], updatedDate: r[12]
      };
    }).sort(function(a, b) {
      if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
      return (a.dueDate || '9999').localeCompare(b.dueDate || '9999');
    });
  } catch(e) {
    return { error: e.message };
  }
}

// All tasks across all projects, enriched with project + plant context.
// Used by Kanban, Gantt, and Calendar views.
function getAllProjectTasks(plantID) {
  try {
    var projects = getProjects(plantID);
    if (projects && projects.error) return projects;
    var projectMap = {};
    projects.forEach(function(p) { projectMap[p.projectID] = p; });

    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { tasks: [], soloProjects: projects.filter(function(p){ return p.taskCount === 0; }) };

    var tasks = data.slice(1).filter(function(r) {
      return r[0] && projectMap[r[1]]; // only tasks whose project matches the plant filter
    }).map(function(r) {
      var proj = projectMap[r[1]];
      return {
        taskID: r[0], projectID: r[1], projectName: proj.projectName,
        plantID: proj.plantID, plantName: proj.plantName, category: proj.category,
        taskName: r[2], owner: r[3] || '', status: r[4] || 'Backlog', priority: r[5] || 'Medium',
        startDate: r[6] instanceof Date ? fmtD(r[6]) : (r[6] || ''),
        dueDate:   r[7] instanceof Date ? fmtD(r[7])  : (r[7]  || ''),
        progress: Number(r[8]) || 0, notes: r[9] || '',
        sortOrder: Number(r[10]) || 0
      };
    }).sort(function(a, b) {
      if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
      return (a.dueDate || '9999').localeCompare(b.dueDate || '9999');
    });

    var soloProjects = projects.filter(function(p) { return p.taskCount === 0; });
    return { tasks: tasks, soloProjects: soloProjects };
  } catch(e) {
    return { error: e.message };
  }
}

function saveProjectTask(task) {
  try {
    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();

    var row = [
      task.taskID || '', task.projectID || '', task.taskName || '', task.owner || '',
      task.status || 'Backlog', task.priority || 'Medium',
      task.startDate || '', task.dueDate || '',
      (task.progress !== undefined && task.progress !== null) ? task.progress : 0,
      task.notes || '', 0, '', now
    ];

    if (task.taskID) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === task.taskID) {
          row[10] = data[i][10]; // preserve SortOrder
          row[11] = data[i][11]; // preserve CreatedDate
          sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
          return { success: true, taskID: task.taskID };
        }
      }
    }

    // New task — append to the end of its project's task order
    var maxOrder = 0;
    for (var j = 1; j < data.length; j++) {
      if (data[j][1] === task.projectID) {
        var so = Number(data[j][10]) || 0;
        if (so > maxOrder) maxOrder = so;
      }
    }
    var newID = generateID('TSK', sheet);
    row[0] = newID;
    row[10] = maxOrder + 1;
    row[11] = now;
    sheet.appendRow(row);
    return { success: true, taskID: newID };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Persists a new task order within a project (drag-to-reorder in the table view)
function reorderProjectTasks(projectID, taskIDOrder) {
  try {
    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    taskIDOrder.forEach(function(taskID, idx) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === taskID && data[i][1] === projectID) {
          sheet.getRange(i + 1, 11).setValue(idx + 1); // SortOrder
          sheet.getRange(i + 1, 13).setValue(now);     // UpdatedDate
          break;
        }
      }
    });
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Updates a task's start/due dates — used by Gantt drag-to-reschedule
function updateProjectTaskDates(taskID, startDate, dueDate) {
  try {
    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskID) {
        sheet.getRange(i + 1, 7).setValue(startDate); // StartDate
        sheet.getRange(i + 1, 8).setValue(dueDate);   // DueDate
        sheet.getRange(i + 1, 13).setValue(now);      // UpdatedDate
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Updates a solo project's (no tasks) start/due dates — used by Gantt drag-to-reschedule
function updateProjectDates(projectID, startDate, dueDate) {
  try {
    var sheet = ensureProjectsSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === projectID) {
        sheet.getRange(i + 1, 9).setValue(startDate);  // StartDate
        sheet.getRange(i + 1, 10).setValue(dueDate);   // DueDate
        sheet.getRange(i + 1, 14).setValue(now);       // UpdatedDate
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function deleteProjectTask(taskID) {
  try {
    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Quick status update — used by Kanban drag-and-drop
function updateProjectTaskStatus(taskID, status) {
  try {
    var sheet = ensureProjectTasksSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskID) {
        sheet.getRange(i + 1, 5).setValue(status);  // Status column
        if (status === 'Completed') sheet.getRange(i + 1, 9).setValue(100); // Progress -> 100
        sheet.getRange(i + 1, 13).setValue(now);    // UpdatedDate
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}


// ============================================================
// EMERGENCY RESPONSE — SOP DOCUMENT LOADER
// Reads SOP files from Google Drive folders via Apps Script
// (avoids CORS issues with direct browser Drive API calls)
// ============================================================

var EMERGENCY_FOLDERS = {
  firefighting: '1Q_NJPFJa3Z-dNPFvr2gm58PxciI0DKQG',
  acmv:         '1SXTJnlnD31LXN3-ayIxrq_NRvLNjKT3E',
  electrical:   '1WBNQ1tdQPGG0KbSvTcLp7mQbiK9BPfn5',
  water:        '1DNFFvfq0knGrgQVFCIiH-q_qa8HQ6_E3',
  lift:         '1j57eDkXPTT-y47vslKBuAcxGMDNjVZ6O'
};

// Returns list of files in a category folder
function getEmergencyFileList(categoryID) {
  try {
    var folderID = EMERGENCY_FOLDERS[categoryID];
    if (!folderID) return { success: false, error: 'Unknown category: ' + categoryID };
    var folder = DriveApp.getFolderById(folderID);
    var files = folder.getFiles();
    var result = [];
    while (files.hasNext()) {
      var file = files.next();
      var mime = file.getMimeType();
      if (mime === 'application/pdf' ||
          mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
          mime === 'text/plain' ||
          mime === 'application/msword') {
        result.push({ id: file.getId(), name: file.getName(), mimeType: mime });
      }
    }
    return { success: true, files: result };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Returns base64-encoded content of a file for client-side extraction
function getEmergencyFileContent(fileID) {
  try {
    var file = DriveApp.getFileById(fileID);
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    var b64 = Utilities.base64Encode(bytes);
    return { success: true, base64: b64, mimeType: blob.getContentType(), name: file.getName() };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// OUT-OF-SPEC EMAIL ALERTS
// AlertSettings sheet: Category | Email | Active | UpdatedDate
// One row per category. saveInspection sends an email when any
// parameter is OutOfRange/OutOfSpec and the equipment's category
// has an active mapping. No mapping = no email (silent skip).
// ============================================================

function ensureAlertSettingsSheet() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('AlertSettings');
  if (!sheet) {
    sheet = ss.insertSheet('AlertSettings');
    sheet.getRange(1, 1, 1, 4).setValues([['Category', 'Email', 'Active', 'UpdatedDate']]);
    sheet.getRange(1, 1, 1, 4)
      .setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAlertSettings() {
  try {
    var sheet = ensureAlertSettingsSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).filter(function(r) { return r[0]; }).map(function(r) {
      return { category: r[0], email: r[1] || '', active: r[2] !== false && r[2] !== 'FALSE' && r[2] !== 'No' };
    });
  } catch(e) {
    return { error: e.message };
  }
}

function saveAlertSetting(setting) {
  try {
    if (!setting || !setting.category) return { success: false, error: 'Category is required' };
    var sheet = ensureAlertSettingsSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    var active = setting.active ? 'Yes' : 'No';
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === setting.category) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[setting.category, setting.email || '', active, now]]);
        return { success: true };
      }
    }
    sheet.appendRow([setting.category, setting.email || '', active, now]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function deleteAlertSetting(category) {
  try {
    var sheet = ensureAlertSettingsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === category) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Sends one email per inspection listing all out-of-spec parameters.
// Silently does nothing if the equipment's category has no active mapping.
function sendOutOfSpecAlert(ss, data, oosDetails) {
  // Equipment lookup
  var eqSheet = ss.getSheetByName('Equipment');
  var eqData  = eqSheet.getDataRange().getValues();
  var eqHeaders = eqData[0];
  var catIdx = eqHeaders.indexOf('Category');
  var equip = null;
  for (var e = 1; e < eqData.length; e++) {
    if (eqData[e][0] === data.equipID) {
      equip = {
        equipmentName: eqData[e][2], assetTag: eqData[e][4],
        areaID: eqData[e][1], category: catIdx >= 0 ? (eqData[e][catIdx] || '') : ''
      };
      break;
    }
  }
  if (!equip || !equip.category) return; // no category = no routing possible

  // Find active alert mapping for this category
  var settings = getAlertSettings();
  if (!settings || settings.error) return;
  var match = null;
  for (var s = 0; s < settings.length; s++) {
    if (settings[s].category === equip.category && settings[s].active && settings[s].email) { match = settings[s]; break; }
  }
  if (!match) return; // no mapping = skip silently

  // Area / Plant lookup
  var areaSheet = ss.getSheetByName('Areas');
  var areaData  = areaSheet.getDataRange().getValues();
  var areaName = '', plantID = '', plantName = '';
  for (var a = 1; a < areaData.length; a++) {
    if (areaData[a][0] === equip.areaID) { areaName = areaData[a][2]; plantID = areaData[a][1]; break; }
  }
  if (plantID) {
    var plantSheet = ss.getSheetByName('Plants');
    var plantData  = plantSheet.getDataRange().getValues();
    for (var p = 1; p < plantData.length; p++) {
      if (plantData[p][0] === plantID) { plantName = plantData[p][1]; break; }
    }
  }

  // Technician lookup
  var techName = data.technicianID || '';
  var techSheet = ss.getSheetByName('Technicians');
  var techData  = techSheet.getDataRange().getValues();
  for (var t = 1; t < techData.length; t++) {
    if (techData[t][0] === data.technicianID) { techName = techData[t][1]; break; }
  }

  // Template specs for context (unit, min, max)
  var specMap = {};
  var templates = getTemplates(data.equipID);
  if (templates && !templates.error) {
    templates.forEach(function(tpl) { specMap[tpl.parameterName] = tpl; });
  }

  // Build email body
  var lines = oosDetails.map(function(d) {
    var spec = specMap[d.parameterName];
    var range = spec ? (spec.minValue + ' - ' + spec.maxValue + (spec.unit ? ' ' + spec.unit : '')) : 'n/a';
    return '  - ' + d.parameterName + ': ' + d.value + (spec && spec.unit ? ' ' + spec.unit : '') + ' (spec: ' + range + ')';
  });

  var subject = '[Alert] Out-of-spec reading - ' + equip.equipmentName + ' (' + equip.category + ')';
  var body =
    'Out-of-spec inspection alert\n\n' +
    'Equipment: ' + equip.equipmentName + ' (' + (equip.assetTag || data.equipID) + ')\n' +
    'Category: ' + equip.category + '\n' +
    'Location: ' + plantName + (areaName ? ' \u203A ' + areaName : '') + '\n' +
    'Technician: ' + techName + '\n' +
    'Date: ' + (data.inspectionDate || '') + '\n\n' +
    'Out-of-spec parameters:\n' + lines.join('\n') + '\n\n' +
    (data.notes ? ('Notes: ' + data.notes + '\n\n') : '') +
    'This is an automated alert from the Factory Maintenance Platform.';

  try {
    MailApp.sendEmail({ to: match.email, subject: subject, body: body });
  } catch(mailErr) {
    // Swallow — email failure must never break the inspection save
  }
}


// ============================================================
// ISSUE TRACKER (CAPA log)
// Issues sheet: IssueID | PlantID | System | Location | Issue |
//   DateReported | Status | RootCause | Correction | TargetDate |
//   CompletedDate | PreventiveAction | PIC | Comment | CreatedDate | UpdatedDate
// IssuePhotos sheet: PhotoID | IssueID | Label | FileURL | FileName | UploadedDate | ThumbURL
// ============================================================

var ISSUE_SYSTEMS = ['HV Switchgear', 'LV Switchgear', 'Transformer', 'Chiller', 'Cooling Tower', 'Pump', 'AHU', 'AHU Panel', 'FCU Panel', 'Solar PV', 'Hygrometer', 'CDA System', 'General'];
var ISSUE_STATUSES = ['Open', 'In Progress', 'Closed'];

function ensureIssuesSheet() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('Issues');
  if (!sheet) {
    sheet = ss.insertSheet('Issues');
    var headers = ['IssueID','PlantID','System','Location','Issue','DateReported','Status',
                    'RootCause','Correction','TargetDate','CompletedDate','PreventiveAction','PIC','Comment','CreatedDate','UpdatedDate'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureIssuePhotosSheet() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('IssuePhotos');
  if (!sheet) {
    sheet = ss.insertSheet('IssuePhotos');
    sheet.getRange(1, 1, 1, 7).setValues([['PhotoID','IssueID','Label','FileURL','FileName','UploadedDate','ThumbURL']]);
    sheet.getRange(1, 1, 1, 7)
      .setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getIssueMeta() {
  // Returns dropdown options for the Issue Tracker UI
  var plants = getPlants();
  return {
    plants: (plants && !plants.error) ? plants : [],
    systems: ISSUE_SYSTEMS,
    statuses: ISSUE_STATUSES
  };
}

function getIssues(plantID) {
  try {
    var ss = getSSv();
    var sheet = ensureIssuesSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var plants = getPlants();
    var plantNames = {};
    (plants || []).forEach(function(p) { plantNames[p.plantID] = p.plantName; });

    var tz = Session.getScriptTimeZone();
    var today = new Date();

    var rows = data.slice(1).filter(function(r) { return r[0]; });
    if (plantID) rows = rows.filter(function(r) { return r[1] === plantID; });

    return rows.map(function(r) {
      var dateReported = r[5] instanceof Date ? fmtD(r[5]) : (r[5] || '');
      var targetDate   = r[9]  instanceof Date ? fmtD(r[9])  : (r[9]  || '');
      var completedDate = r[10] instanceof Date ? fmtD(r[10]) : (r[10] || '');
      var status = r[6] || 'Open';

      // Aging: days open (if not Closed) or days-to-resolve (if Closed)
      var aging = null;
      if (dateReported) {
        var startD = new Date(dateReported + 'T00:00:00');
        var endD = (status === 'Closed' && completedDate) ? new Date(completedDate + 'T00:00:00') : today;
        aging = Math.floor((endD - startD) / 86400000);
        if (aging < 0) aging = 0;
      }

      // Overdue flag: open/in-progress and past target date
      var overdue = false;
      if (status !== 'Closed' && targetDate) {
        overdue = new Date(targetDate + 'T00:00:00') < new Date(fmtD(today) + 'T00:00:00');
      }

      return {
        issueID: r[0], plantID: r[1], plantName: plantNames[r[1]] || r[1],
        system: r[2], location: r[3], issue: r[4],
        dateReported: dateReported, status: status,
        rootCause: r[7] || '', correction: r[8] || '',
        targetDate: targetDate, completedDate: completedDate,
        preventiveAction: r[11] || '', pic: r[12] || '', comment: r[13] || '',
        createdDate: r[14], updatedDate: r[15],
        aging: aging, overdue: overdue
      };
    }).sort(function(a, b) {
      // Open/In Progress first (by aging desc), then Closed (by completed date desc)
      var rank = { 'Open': 0, 'In Progress': 1, 'Closed': 2 };
      if (rank[a.status] !== rank[b.status]) return rank[a.status] - rank[b.status];
      return (b.aging || 0) - (a.aging || 0);
    });
  } catch(e) {
    return { error: e.message };
  }
}

function saveIssue(issue) {
  try {
    var sheet = ensureIssuesSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();

    var row = [
      issue.issueID || '', issue.plantID || '', issue.system || '', issue.location || '',
      issue.issue || '', issue.dateReported || '', issue.status || 'Open',
      issue.rootCause || '', issue.correction || '',
      issue.targetDate || '', issue.completedDate || '',
      issue.preventiveAction || '', issue.pic || '', issue.comment || '',
      '', now
    ];

    if (issue.issueID) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === issue.issueID) {
          row[14] = data[i][14]; // preserve original CreatedDate
          sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
          return { success: true, issueID: issue.issueID };
        }
      }
    }

    // New issue
    var newID = generateID('ISS', sheet);
    row[0] = newID;
    row[14] = now; // CreatedDate
    sheet.appendRow(row);
    return { success: true, issueID: newID };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function deleteIssue(issueID) {
  try {
    var sheet = ensureIssuesSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === issueID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── AUTO-ARCHIVE CLOSED ISSUES ────────────────────────────
// Runs daily via a time-based trigger (set up manually in Apps Script editor):
//   Triggers → Add Trigger → autoArchiveClosedIssues → Time-driven → Day timer
//
// Moves any issue where:
//   - Status === 'Closed'
//   - UpdatedDate is more than 14 days ago
// ...into Issues_Archive (and its photos into IssuePhotos_Archive).
// Data is preserved in the archive sheets; active sheets stay clean.
function autoArchiveClosedIssues() {
  try {
    var ss = getSSv();
    var issueSheet = ensureIssuesSheet();
    var photoSheet = ensureIssuePhotosSheet();

    // Ensure archive sheets exist with same headers + ArchivedDate column
    var archiveIssueSheet = ss.getSheetByName('Issues_Archive');
    if (!archiveIssueSheet) {
      archiveIssueSheet = ss.insertSheet('Issues_Archive');
      var issueHeaders = issueSheet.getRange(1, 1, 1, issueSheet.getLastColumn()).getValues()[0];
      issueHeaders.push('ArchivedDate');
      archiveIssueSheet.getRange(1, 1, 1, issueHeaders.length).setValues([issueHeaders]);
      archiveIssueSheet.getRange(1, 1, 1, issueHeaders.length)
        .setBackground('#7c3aed').setFontColor('#ffffff').setFontWeight('bold');
      archiveIssueSheet.setFrozenRows(1);
    }

    var archivePhotoSheet = ss.getSheetByName('IssuePhotos_Archive');
    if (!archivePhotoSheet) {
      archivePhotoSheet = ss.insertSheet('IssuePhotos_Archive');
      var photoHeaders = photoSheet.getRange(1, 1, 1, photoSheet.getLastColumn()).getValues()[0];
      photoHeaders.push('ArchivedDate');
      archivePhotoSheet.getRange(1, 1, 1, photoHeaders.length).setValues([photoHeaders]);
      archivePhotoSheet.getRange(1, 1, 1, photoHeaders.length)
        .setBackground('#7c3aed').setFontColor('#ffffff').setFontWeight('bold');
      archivePhotoSheet.setFrozenRows(1);
    }

    var issueData = issueSheet.getDataRange().getValues();
    var now = new Date();
    var nowStr = fmtD(now);
    var cutoff = 14 * 24 * 60 * 60 * 1000; // 14 days in ms

    // Find rows to archive (collect info first, delete bottom-up after)
    var toArchive = [];
    for (var i = 1; i < issueData.length; i++) {
      var status = issueData[i][6];       // Status column
      var updatedDate = issueData[i][15]; // UpdatedDate column
      if (status !== 'Closed') continue;
      if (!updatedDate) continue;
      var closedAt = updatedDate instanceof Date ? updatedDate : new Date(updatedDate);
      if (isNaN(closedAt.getTime())) continue;
      if ((now - closedAt) > cutoff) {
        var archiveRow = issueData[i].slice(); // copy full row
        archiveRow.push(nowStr);               // append ArchivedDate
        toArchive.push({ rowIndex: i + 1, issueID: issueData[i][0], archiveRow: archiveRow });
      }
    }

    if (!toArchive.length) {
      Logger.log('autoArchiveClosedIssues: nothing to archive.');
      return;
    }

    // Archive issue rows
    toArchive.forEach(function(item) {
      archiveIssueSheet.appendRow(item.archiveRow);
    });

    // Archive + remove matching photo rows
    var issueIDSet = {};
    toArchive.forEach(function(item) { issueIDSet[item.issueID] = true; });
    var photoData = photoSheet.getDataRange().getValues();
    var photoRowsToDelete = [];
    for (var p = 1; p < photoData.length; p++) {
      if (issueIDSet[photoData[p][1]]) {
        var photoArchiveRow = photoData[p].slice();
        photoArchiveRow.push(nowStr);
        archivePhotoSheet.appendRow(photoArchiveRow);
        photoRowsToDelete.push(p + 1);
      }
    }
    // Delete photo rows bottom-up
    for (var pd = photoRowsToDelete.length - 1; pd >= 0; pd--) {
      photoSheet.deleteRow(photoRowsToDelete[pd]);
    }

    // Delete original issue rows bottom-up
    toArchive.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
    toArchive.forEach(function(item) { issueSheet.deleteRow(item.rowIndex); });

    Logger.log('autoArchiveClosedIssues: archived ' + toArchive.length + ' issue(s).');
  } catch(e) {
    Logger.log('autoArchiveClosedIssues ERROR: ' + e.message);
  }
}

// Quick status update — used by Kanban drag-and-drop
function updateIssueStatus(issueID, status) {
  try {
    var sheet = ensureIssuesSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === issueID) {
        sheet.getRange(i + 1, 7).setValue(status); // Status column
        // Auto-set CompletedDate when moved to Closed (if not already set)
        if (status === 'Closed' && !data[i][10]) {
          sheet.getRange(i + 1, 11).setValue(fmtD(new Date()));
        }
        sheet.getRange(i + 1, 16).setValue(now); // UpdatedDate
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── Issue Photos (before/after) — reuses Drive folder pattern ──
function saveIssuePhotos(issueID, photos) {
  try {
    if (!photos || !photos.length) return { success: true, urls: [] };
    var folder = getPhotoFolder();
    var sheet  = ensureIssuePhotosSheet();
    var now    = new Date();
    var urls   = [];

    var subFolderName = 'ISSUE_' + issueID;
    var subFolders = folder.getFoldersByName(subFolderName);
    var subFolder  = subFolders.hasNext() ? subFolders.next() : folder.createFolder(subFolderName);

    photos.forEach(function(photo) {
      var bytes   = Utilities.base64Decode(photo.base64);
      var blob    = Utilities.newBlob(bytes, photo.mimeType, photo.name);
      var file    = subFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var fileId  = file.getId();
      var fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
      var thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w200';
      var photoID = 'PH-' + issueID + '-' + (sheet.getLastRow());
      sheet.appendRow([photoID, issueID, photo.label || 'before', fileUrl, photo.name, now.toISOString(), thumbUrl]);
      urls.push({ photoID: photoID, fileUrl: fileUrl, thumbUrl: thumbUrl, fileName: photo.name, label: photo.label || 'before' });
    });

    return { success: true, urls: urls };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getIssuePhotos(issueID) {
  try {
    var sheet = ensureIssuePhotosSheet();
    var data  = sheet.getDataRange().getValues();
    var result = { before: [], after: [] };
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === issueID) {
        var label = (data[i][2] === 'after') ? 'after' : 'before';
        result[label].push({ photoID: data[i][0], fileUrl: data[i][3], fileName: data[i][4], thumbUrl: data[i][6] || data[i][3] });
      }
    }
    return result;
  } catch(e) {
    return { before: [], after: [] };
  }
}

function deleteIssuePhoto(photoID) {
  try {
    var sheet = ensureIssuePhotosSheet();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === photoID) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// PHOTO ATTACHMENTS
// Photos are saved to Google Drive under a folder per inspection.
// URLs stored in Photos sheet: PhotoID|InspectionID|EquipID|FileURL|FileName|UploadedDate
// ============================================================

var PHOTO_FOLDER_NAME = 'FMP_Inspection_Photos';

function getPhotoFolder() {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function ensurePhotosSheet() {
  var ss = getSSv();
  var sheet = ss.getSheetByName('Photos');
  if (!sheet) {
    sheet = ss.insertSheet('Photos');
    sheet.getRange(1, 1, 1, 7).setValues([['PhotoID','InspectionID','EquipID','FileURL','FileName','UploadedDate','ThumbURL']]);
    sheet.getRange(1, 1, 1, 7)
      .setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function savePhotos(inspectionID, equipID, photos) {
  // photos = [{name, base64, mimeType}, ...]
  try {
    if (!photos || !photos.length) return { success: true, urls: [] };
    var folder = getPhotoFolder();
    var sheet  = ensurePhotosSheet();
    var now    = new Date();
    var urls   = [];

    // Create subfolder per inspection for organisation
    var subFolderName = inspectionID + '_' + equipID;
    var subFolders = folder.getFoldersByName(subFolderName);
    var subFolder  = subFolders.hasNext() ? subFolders.next() : folder.createFolder(subFolderName);

    photos.forEach(function(photo) {
      var bytes    = Utilities.base64Decode(photo.base64);
      var blob     = Utilities.newBlob(bytes, photo.mimeType, photo.name);
      var file     = subFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var fileId   = file.getId();
      var fileUrl  = 'https://drive.google.com/uc?export=view&id=' + fileId;
      var thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w120';
      var photoID  = 'PH-' + inspectionID + '-' + (sheet.getLastRow());
      sheet.appendRow([photoID, inspectionID, equipID, fileUrl, photo.name, now.toISOString(), thumbUrl]);
      urls.push({ photoID: photoID, fileUrl: fileUrl, thumbUrl: thumbUrl, fileName: photo.name });
    });

    return { success: true, urls: urls };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getPhotosByInspection(inspectionID) {
  try {
    var sheet = ensurePhotosSheet();
    var data  = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === inspectionID) {
        result.push({ photoID: data[i][0], fileUrl: data[i][3], fileName: data[i][4], thumbUrl: data[i][6] || data[i][3] });
      }
    }
    return result;
  } catch(e) {
    return [];
  }
}

// Returns photos for multiple inspectionIDs in one call (for report)
function getPhotosByInspectionIds(inspectionIDs) {
  try {
    if (!inspectionIDs || !inspectionIDs.length) return {};
    var ss = getSSv();
    var sheet = ss.getSheetByName('Photos');
    if (!sheet) return {}; // Photos sheet not created yet — skip silently
    var data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};
    var result = {};
    for (var i = 1; i < data.length; i++) {
      var iid = String(data[i][1]);
      if (inspectionIDs.indexOf(iid) >= 0) {
        if (!result[iid]) result[iid] = [];
        result[iid].push({ photoID: data[i][0], fileUrl: data[i][3], fileName: data[i][4], thumbUrl: data[i][6] || data[i][3] });
      }
    }
    return result;
  } catch(e) {
    return {}; // Never break the report if photos fail
  }
}

// ============================================================
// FIRE EXTINGUISHER TRACKING SYSTEM — Backend Functions
// Merged into FMP v4.0
// ============================================================

const CONFIG_FE = {
  SPREADSHEET_ID: '1forV8LDB5EZ9UgRBbRiZczi4WeXYjw0GbUWzoHxm8zo'
};

function getFESheet(name) {
  return SpreadsheetApp.openById(CONFIG_FE.SPREADSHEET_ID).getSheetByName(name);
}

function getFireExtinguishers() {
  var sheet = getFESheet('FireExtinguishers');
  var data  = sheet.getDataRange().getValues();
  data.shift();

  var history = getFESheet('InspectionHistory').getDataRange().getValues();
  if (history.length > 1) history.shift();
  else history = [];

  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var month = Number(Utilities.formatDate(now, tz, 'M'));
  var year  = Number(Utilities.formatDate(now, tz, 'yyyy'));

  var inspectedThisMonth = {};
  history.filter(function(r) {
    return Number(r[2]) === month && Number(r[3]) === year;
  }).forEach(function(r) { inspectedThisMonth[r[4]] = true; });

  var tz  = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var todayStrDisplay = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy');

  return data.map(function(row) {
    var expiryRaw = row[7] || '';
    var expiryISO = '';     // yyyy-MM-dd  — used for date comparison
    var expiryDisplay = ''; // dd-MM-yyyy  — shown in UI

    if (expiryRaw instanceof Date) {
      expiryISO     = Utilities.formatDate(expiryRaw, tz, 'yyyy-MM-dd');
      expiryDisplay = Utilities.formatDate(expiryRaw, tz, 'dd-MM-yyyy');
    } else if (expiryRaw) {
      // Handle both yyyy-MM-dd and dd-MM-yyyy stored values
      var raw = String(expiryRaw).split('T')[0];
      if (raw.indexOf('-') >= 0) {
        var parts = raw.split('-');
        if (parts[0].length === 4) {
          // stored as yyyy-MM-dd
          expiryISO     = raw;
          expiryDisplay = parts[2] + '-' + parts[1] + '-' + parts[0];
        } else {
          // stored as dd-MM-yyyy
          expiryDisplay = raw;
          expiryISO     = parts[2] + '-' + parts[1] + '-' + parts[0];
        }
      }
    }

    var daysToExpiry = null;
    var isExpired    = false;
    var expiringSoon = false;
    if (expiryISO) {
      var expDate  = new Date(expiryISO  + 'T00:00:00');
      var today    = new Date(todayStr   + 'T00:00:00');
      daysToExpiry = Math.floor((expDate - today) / 86400000);
      isExpired    = daysToExpiry < 0;
      expiringSoon = daysToExpiry >= 0 && daysToExpiry <= 30;
    }

    return {
      feId:         row[0],
      building:     row[1],
      floor:        row[2],
      location:     row[3],
      x:            Number(row[4]) || 120,
      y:            Number(row[5]) || 120,
      type:         row[6],
      expiryDate:   expiryDisplay,  // DD-MM-YYYY for display
      expiryISO:    expiryISO,      // YYYY-MM-DD for date input field
      daysToExpiry: daysToExpiry,
      isExpired:    isExpired,
      expiringSoon: expiringSoon,
      inspected:    !!inspectedThisMonth[row[0]]
    };
  });
}

function getFEDashboardStats() {
  var feList  = getFireExtinguishers();
  var total   = feList.length;
  var completed = feList.filter(function(x) { return x.inspected; }).length;

  var history = getFESheet('InspectionHistory').getDataRange().getValues();
  if (history.length <= 1) return { total: total, completed: completed, pending: total - completed, failed: 0 };
  history.shift();

  var tz    = Session.getScriptTimeZone();
  var now   = new Date();
  var month = Number(Utilities.formatDate(now, tz, 'M'));
  var year  = Number(Utilities.formatDate(now, tz, 'yyyy'));

  var thisMonth = history.filter(function(r) {
    return Number(r[2]) === month && Number(r[3]) === year;
  });
  var failed = thisMonth.filter(function(r) { return r[17] === 'FAIL'; }).length;

  return { total: total, completed: completed, pending: total - completed, failed: failed };
}

function saveFEInspection(feId, physicalCondition, rivetPin, pressureGauge,
    dischargeHose, handleCondition, labelCondition, visibility, remarks, technicianName) {
  var feList = getFireExtinguishers();
  var fe = null;
  for (var i = 0; i < feList.length; i++) { if (feList[i].feId === feId) { fe = feList[i]; break; } }
  if (!fe) return { success: false, error: 'FE not found' };

  var history = getFESheet('InspectionHistory');
  var now   = new Date();
  var tz    = Session.getScriptTimeZone();
  var month = Number(Utilities.formatDate(now, tz, 'M'));
  var year  = Number(Utilities.formatDate(now, tz, 'yyyy'));
  var inspectionId = 'INS' + now.getTime();

  var checks = [physicalCondition, rivetPin, pressureGauge, dischargeHose, handleCondition, labelCondition, visibility];
  var status = 'PASS';
  var failVals = ['Not Good', 'Missing', 'Damaged', 'Low Pressure', 'Over Pressure'];
  var attVals  = ['Faded', 'Obstructed'];
  for (var j = 0; j < checks.length; j++) {
    if (failVals.indexOf(checks[j]) >= 0) { status = 'FAIL'; break; }
    if (attVals.indexOf(checks[j]) >= 0 && status !== 'FAIL') status = 'ATTENTION';
  }

  history.appendRow([
    inspectionId, now, month, year,
    fe.feId, fe.building, fe.floor, fe.location, fe.type,
    physicalCondition, rivetPin, pressureGauge, dischargeHose,
    handleCondition, labelCondition, visibility, remarks, status,
    technicianName || ''
  ]);
  return { success: true };
}

function updateFEPosition(feId, x, y) {
  var sheet = getFESheet('FireExtinguishers');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === feId) {
      sheet.getRange(i + 1, 5).setValue(x);
      sheet.getRange(i + 1, 6).setValue(y);
      return { success: true };
    }
  }
  return { success: false };
}

function updateFEDetails(feId, building, floor, location, newFeId, expiryDate) {
  var sheet = getFESheet('FireExtinguishers');
  var data  = sheet.getDataRange().getValues();

  // Ensure column 8 header exists (ExpiryDate)
  var headers = data[0];
  if (!headers[7] || headers[7] !== 'ExpiryDate') {
    sheet.getRange(1, 8).setValue('ExpiryDate');
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(feId).trim()) {
      if (newFeId && newFeId !== feId) sheet.getRange(i + 1, 1).setValue(newFeId);
      sheet.getRange(i + 1, 2).setValue(building);
      sheet.getRange(i + 1, 3).setValue(floor);
      sheet.getRange(i + 1, 4).setValue(location);
      if (expiryDate !== undefined) {
        if (expiryDate) {
          // Convert from yyyy-MM-dd (browser input) to dd-MM-yyyy (sheet storage)
          var ep = expiryDate.split('-');
          var displayVal = ep.length === 3 ? ep[2] + '-' + ep[1] + '-' + ep[0] : expiryDate;
          sheet.getRange(i + 1, 8).setValue(displayVal);
        } else {
          sheet.getRange(i + 1, 8).setValue('');
        }
      }
      return { success: true };
    }
  }
  return { success: false };
}
function addFireExtinguisherCustom(feId, building, floor, location, x, y, type) {
  getFESheet('FireExtinguishers').appendRow([feId, building, floor, location, x || 50, y || 50, type]);
  return { success: true };
}

function getFEInspectionHistory(feId) {
  var sheet = getFESheet('InspectionHistory');
  var data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  data.shift();
  return data.filter(function(x) { return x[4] === feId; })
    .reverse().slice(0, 10)
    .map(function(x) {
      return { date: x[1], inspector: x[18], remarks: x[16], status: x[17] };
    });
}

function generateFEMonthlyPDF() {
  var sheet = getFESheet('InspectionHistory');
  var data  = sheet.getDataRange().getValues();
  data.shift();

  var tz    = Session.getScriptTimeZone();
  var now   = new Date();
  var month = Number(Utilities.formatDate(now, tz, 'M'));
  var year  = Number(Utilities.formatDate(now, tz, 'yyyy'));

  var records = data.filter(function(r) {
    return Number(r[2]) === month && Number(r[3]) === year;
  });

  var logoBase64 = getFELogoBase64();

  var html = '<html><head><style>' +
    'body{font-family:Arial;padding:20px;}' +
    'h2{text-align:center;}' +
    'table{width:100%;border-collapse:collapse;font-size:9px;}' +
    'th{background:#f3f4f6;}' +
    'th,td{border:1px solid #000;padding:4px;}' +
    '</style></head><body>' +
    '<div style="text-align:center;margin-bottom:20px">' +
    '<img src="data:image/jpeg;base64,' + logoBase64 + '" style="height:90px;margin-bottom:10px">' +
    '<h2 style="margin:0">Fire Fighting Equipment Inspection Checklist</h2>' +
    '<p>Month: ' + month + '/' + year + '</p></div>' +
    '<table><tr><th>FE ID</th><th>Building</th><th>Floor</th><th>Location</th><th>Type</th>' +
    '<th>Physical</th><th>Pin</th><th>Gauge</th><th>Hose</th><th>Handle</th><th>Label</th><th>Visibility</th><th>Remarks</th><th>Status</th></tr>';

  records.forEach(function(r) {
    html += '<tr><td>' + r[4] + '</td><td>' + r[5] + '</td><td>' + r[6] + '</td><td>' + r[7] + '</td><td>' + r[8] + '</td>' +
      '<td>' + r[9] + '</td><td>' + r[10] + '</td><td>' + r[11] + '</td><td>' + r[12] + '</td>' +
      '<td>' + r[13] + '</td><td>' + r[14] + '</td><td>' + r[15] + '</td>' +
      '<td>' + (r[16] || '') + '</td><td>' + r[17] + '</td></tr>';
  });

  html += '</table><br><br>' +
    '<table style="width:100%;border:none"><tr>' +
    '<td style="border:none;width:50%">Prepared By<br><br><br>____________________<br>Person in-charge</td>' +
    '<td style="border:none;width:50%">Verified By<br><br><br>____________________<br>Supervisor</td>' +
    '</tr></table></body></html>';

  var blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  var file = DriveApp.createFile(blob);
  return file.getUrl();
}

function getFELogoBase64() {
  var file  = DriveApp.getFileById('1vNs1dMaBupAwwOjnoN6X1WaZ8I0Pia4C');
  var bytes = file.getBlob().getBytes();
  return Utilities.base64Encode(bytes);
}

// ── FE INSPECTION HISTORY ─────────────────────────────────
function getFEInspectionHistory(feId) {
  var sheet = getFESheet('InspectionHistory');
  var data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  data.shift();
  return data.filter(function(r) { return r[4] === feId; })
    .reverse().slice(0, 10)
    .map(function(r) {
      return {
        inspectionId: r[0],
        date:         r[1],
        month:        r[2],
        year:         r[3],
        status:       r[17],
        technician:   r[18] || '',
        remarks:      r[16] || '',
        physicalCondition: r[9],
        rivetPin:          r[10],
        pressureGauge:     r[11],
        dischargeHose:     r[12],
        handleCondition:   r[13],
        labelCondition:    r[14],
        visibility:        r[15]
      };
    });
}

// ── HISTORICAL FE REPORT (by month/year) ──────────────────
function getFEHistoricalReport(month, year, plant) {


  Logger.log('getFEHistoricalReport START');
  Logger.log('month=' + month);
  Logger.log('year=' + year);


  try {
    var tz   = Session.getScriptTimeZone();
    var mo   = parseInt(month);
    var yr   = parseInt(year);
    var feList = getFireExtinguishers();

    

    var sheet   = getFESheet('InspectionHistory');
    var data    = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { month: mo, year: yr, total: feList.length, completed: 0, pending: feList.length, failed: 0, records: [] };
    }
    data.shift();

   var records = data.filter(function(r) {

  var monthMatch = Number(r[2]) === mo;
  var yearMatch  = Number(r[3]) === yr;

  var plantMatch =
    plant === 'ALL' ||
    String(r[5]).toUpperCase() === plant;

  return monthMatch &&
         yearMatch &&
         plantMatch;
});

    var inspectedIDs = {};
    records.forEach(function(r) { inspectedIDs[r[4]] = r[17]; });

    var uniqueCompleted = Object.keys(inspectedIDs).length;
    var failed = records.filter(function(r) { return r[17] === 'FAIL'; }).length;

    var result = {
  month: mo,
  year: yr,
  total: feList.length,
  completed: uniqueCompleted,
  pending: feList.length - uniqueCompleted,
  failed: failed,
  records: records.map(function(r) {
    return {
  feId: String(r[4] || ''),
  building: String(r[5] || ''),
  floor: String(r[6] || ''),
  location: String(r[7] || ''),
  type: String(r[8] || ''),

  physicalCondition: String(r[9] || ''),
  rivetPin: String(r[10] || ''),
  pressureGauge: String(r[11] || ''),
  dischargeHose: String(r[12] || ''),
  handleCondition: String(r[13] || ''),
  labelCondition: String(r[14] || ''),
  visibility: String(r[15] || ''),

  remarks: String(r[16] || ''),
  status: String(r[17] || ''),
  technician: String(r[18] || ''),

  date: r[1]
    ? Utilities.formatDate(
        new Date(r[1]),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      )
    : ''
};
  })
};

Logger.log(JSON.stringify(result));

return result;

}catch(e) {

  return {
    error:
      'MESSAGE=' + e.message +
      '\nLINE=' + (e.lineNumber || '?') +
      '\nSTACK=' + (e.stack || 'No Stack')
  };
}
}

// ── FE OVERDUE — check per FE ─────────────────────────────
function getFEOverdueStatus() {
  try {
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();
    var today = Number(Utilities.formatDate(now, tz, 'd'));
    var month = Number(Utilities.formatDate(now, tz, 'M'));
    var year  = Number(Utilities.formatDate(now, tz, 'yyyy'));

    var feList = getFireExtinguishers();

    // Days remaining in month
    var daysInMonth  = new Date(year, month, 0).getDate();
    var daysLeft     = daysInMonth - today;

    // Mark overdue: not inspected this month
    var overdue = feList.filter(function(fe) { return !fe.inspected; });

    return {
      today:       today,
      daysLeft:    daysLeft,
      daysInMonth: daysInMonth,
      overdue:     overdue.length,
      overdueList: overdue
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ── FE FAILED LIST ────────────────────────────────────────
function getFEFailedList() {
  try {
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();
    var month = Number(Utilities.formatDate(now, tz, 'M'));
    var year  = Number(Utilities.formatDate(now, tz, 'yyyy'));

    var sheet = getFESheet('InspectionHistory');
    var data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    data.shift();

    // Get latest inspection per FE this month
    var latestByFE = {};
    data.filter(function(r) {
      return Number(r[2]) === month && Number(r[3]) === year;
    }).forEach(function(r) {
      var feId = r[4];
      if (!latestByFE[feId] || r[0] > latestByFE[feId].id) {
        latestByFE[feId] = {
          id: r[0], feId: feId, building: r[5], floor: r[6],
          location: r[7], type: r[8], status: r[17],
          technician: r[18] || '', date: r[1], remarks: r[16] || ''
        };
      }
    });

    return Object.values(latestByFE).filter(function(r) {
      return r.status === 'FAIL' || r.status === 'ATTENTION';
    });
  } catch(e) {
    return [];
  }
}

// ── QR INSPECTION LINKS ───────────────────────────────────
function generateFEInspectionLinks() {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG_FE.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Inspection_Links') || ss.insertSheet('Inspection_Links');
    sheet.clear();
    sheet.getRange(1,1,1,3).setValues([['FE ID','Location','Inspection Link']]);

    var feList  = getFireExtinguishers();
    var baseUrl = ScriptApp.getService().getUrl();

    var rows = feList.map(function(fe) {
      return [fe.feId, fe.location, baseUrl + '?feId=' + encodeURIComponent(fe.feId)];
    });

    if (rows.length) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
      sheet.autoResizeColumns(1, 3);
    }
    return { success: true, count: rows.length, sheetUrl: ss.getUrl() };
  } catch(e) {
    return { error: e.message };
  }
}

// ── FE SETTINGS — get/save map URLs ───────────────────────
function getFESettings() {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG_FE.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { mapUrls: {}, feTypes: ['ABC','CO2'] };
    var data = sheet.getDataRange().getValues();
    var settings = { mapUrls: {}, feTypes: ['ABC','CO2'] };
    data.forEach(function(row) {
      if (row[0] === 'MAP_URL') settings.mapUrls[row[1]] = row[2];
      if (row[0] === 'FE_TYPE' && row[1]) {
        if (!settings.feTypes) settings.feTypes = [];
        if (settings.feTypes.indexOf(row[1]) < 0) settings.feTypes.push(row[1]);
      }
    });
    return settings;
  } catch(e) {
    return { mapUrls: {}, feTypes: ['ABC','CO2'] };
  }
}

function saveFESetting(key, subKey, value) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG_FE.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
    var data  = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key && data[i][1] === subKey) {
        sheet.getRange(i+1, 3).setValue(value);
        return { success: true };
      }
    }
    sheet.appendRow([key, subKey, value]);
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

// ── AUTO-DETECT FE INFO FROM ID ───────────────────────────
function detectFEInfo(feId) {
  var id  = String(feId).toUpperCase();
  var building = 'P1';
  var floor    = 'Ground';

  if (id.indexOf('CO2II') >= 0 || id.indexOf('DPII') >= 0 || id.indexOf('P2') >= 0) building = 'P2';

  if      (id.indexOf('-MF-') >= 0 || id.indexOf('MEZZ') >= 0) floor = 'Mezzanine';
  else if (id.indexOf('-FF-') >= 0 || id.indexOf('-1F-') >= 0) floor = '1st Floor';
  else if (id.indexOf('-SF-') >= 0 || id.indexOf('-2F-') >= 0) floor = '2nd Floor';
  else if (id.indexOf('-RF-') >= 0 || id.indexOf('ROOF') >= 0) floor = 'Roof';
  else floor = 'Ground';

  return { building: building, floor: floor };
}
