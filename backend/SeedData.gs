// ============================================================
// SeedData.gs — Plant 1 Seed Data
// 10 Areas, 89 Equipment
// Run seedData() once from Apps Script editor
// ============================================================

function seedData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const now = new Date().toISOString();

  // ── PLANT 1 ──────────────────────────────────────────────
  const plantSheet = ss.getSheetByName('Plants');
  const existingPlants = plantSheet.getDataRange().getValues().map(r => r[0]);
  if (!existingPlants.includes('P1')) {
    plantSheet.appendRow(['P1', 'Plant 1', 'Main Industrial Area', 'Active', now]);
    Logger.log('Plant 1 added.');
  } else {
    Logger.log('Plant 1 already exists, skipping.');
  }

  // ── AREAS ─────────────────────────────────────────────────
  const areaSheet = ss.getSheetByName('Areas');
  const existingAreas = areaSheet.getDataRange().getValues().map(r => r[0]);

  const areas = [
    ['A001', 'P1', 'Consumer Room HT 33kV P1', 'High Voltage Consumer Room', 'Active', now],
    ['A002', 'P1', 'Transformer Room 33kV P1', 'Main Transformer Room', 'Active', now],
    ['A003', 'P1', 'MSB LV Room 420V P1', 'Main Switch Board Low Voltage', 'Active', now],
    ['A004', 'P1', 'Production Hygrometer P1', 'Temperature and Humidity Monitoring', 'Active', now],
    ['A005', 'P1', 'Solar Panel P1', 'Solar PV System', 'Active', now],
    ['A006', 'P1', 'Office FCU P1', 'Office Fan Coil Units', 'Active', now],
    ['A007', 'P1', 'AHU Room P1', 'Air Handling Units', 'Active', now],
    ['A008', 'P1', 'Chiller Room P1', 'Chiller Plant Room', 'Active', now],
    ['A009', 'P1', 'L11 Testing Bay P1', 'L11 Testing Bay Sub-Boards', 'Active', now],
    ['A010', 'P1', 'Production Line P1', 'Production Line Sub-Boards', 'Active', now]
  ];

  areas.forEach(area => {
    if (!existingAreas.includes(area[0])) {
      areaSheet.appendRow(area);
    }
  });
  Logger.log('Areas for Plant 1 seeded.');

  // ── EQUIPMENT ─────────────────────────────────────────────
  const equipSheet = ss.getSheetByName('Equipment');
  const existingEquip = equipSheet.getDataRange().getValues().map(r => r[0]);

  const equipment = [
    // A001 - Consumer Room HT 33kV (13)
    ['EQ001', 'A001', 'Feeder 33kV Incomer 1 (1S5)', 'QR-EQ001', 'AT-EQ001', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ002', 'A001', 'CLR Feeder 1', 'QR-EQ002', 'AT-EQ002', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ003', 'A001', 'Feeder From CLR-1', 'QR-EQ003', 'AT-EQ003', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ004', 'A001', 'Feeder TX-M1 Phase 1', 'QR-EQ004', 'AT-EQ004', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ005', 'A001', 'Feeder TX-P1 Phase 1', 'QR-EQ005', 'AT-EQ005', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ006', 'A001', 'Feeder Outgoing 1 to Phase 2', 'QR-EQ006', 'AT-EQ006', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ007', 'A001', 'Feeder TX-P3 Phase 1', 'QR-EQ007', 'AT-EQ007', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ008', 'A001', 'Feeder 33kV Incomer 2 (2S5)', 'QR-EQ008', 'AT-EQ008', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ009', 'A001', 'CLR Feeder 2', 'QR-EQ009', 'AT-EQ009', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ010', 'A001', 'Feeder From CLR-2', 'QR-EQ010', 'AT-EQ010', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ011', 'A001', 'Feeder TX-M2 Phase 1', 'QR-EQ011', 'AT-EQ011', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ012', 'A001', 'Feeder TX-P2 Phase 1', 'QR-EQ012', 'AT-EQ012', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ013', 'A001', 'Feeder Outgoing 2 to Phase 2', 'QR-EQ013', 'AT-EQ013', 'Daily', 'HV Switchgear', 'Active', now],

    // A002 - Transformer Room (5)
    ['EQ014', 'A002', 'TX-P3', 'QR-EQ014', 'AT-EQ014', 'Daily', 'Transformer', 'Active', now],
    ['EQ015', 'A002', 'TX-M1', 'QR-EQ015', 'AT-EQ015', 'Daily', 'Transformer', 'Active', now],
    ['EQ016', 'A002', 'TX-M2', 'QR-EQ016', 'AT-EQ016', 'Daily', 'Transformer', 'Active', now],
    ['EQ017', 'A002', 'TX-P1', 'QR-EQ017', 'AT-EQ017', 'Daily', 'Transformer', 'Active', now],
    ['EQ018', 'A002', 'TX-P2', 'QR-EQ018', 'AT-EQ018', 'Daily', 'Transformer', 'Active', now],

    // A003 - MSB LV Room (9)
    ['EQ019', 'A003', 'MSB-M1', 'QR-EQ019', 'AT-EQ019', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ020', 'A003', 'MSB-M2', 'QR-EQ020', 'AT-EQ020', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ021', 'A003', 'MSB-P1', 'QR-EQ021', 'AT-EQ021', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ022', 'A003', 'MSB-P2', 'QR-EQ022', 'AT-EQ022', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ023', 'A003', 'ESB-M', 'QR-EQ023', 'AT-EQ023', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ024', 'A003', 'ESB-P1', 'QR-EQ024', 'AT-EQ024', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ025', 'A003', 'ESB-P2', 'QR-EQ025', 'AT-EQ025', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ026', 'A003', 'SSB-SR', 'QR-EQ026', 'AT-EQ026', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ027', 'A003', 'MSB-P3', 'QR-EQ027', 'AT-EQ027', 'Daily', 'LV Switchgear', 'Active', now],

    // A004 - Production Hygrometer (5)
    ['EQ028', 'A004', 'L10 Production Line', 'QR-EQ028', 'AT-EQ028', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ029', 'A004', 'L11 Test Room', 'QR-EQ029', 'AT-EQ029', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ030', 'A004', 'MPA Warehouse', 'QR-EQ030', 'AT-EQ030', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ031', 'A004', 'HV Room', 'QR-EQ031', 'AT-EQ031', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ032', 'A004', 'IQC Room', 'QR-EQ032', 'AT-EQ032', 'Daily', 'Hygrometer', 'Active', now],

    // A005 - Solar Panel (4)
    ['EQ033', 'A005', 'MSB-M1 (Incoming 1, Inverter 1~5)', 'QR-EQ033', 'AT-EQ033', 'Daily', 'Solar PV', 'Active', now],
    ['EQ034', 'A005', 'MSB-M2 (Incoming 2, Inverter 6~10)', 'QR-EQ034', 'AT-EQ034', 'Daily', 'Solar PV', 'Active', now],
    ['EQ035', 'A005', 'MSB-P1 (Incoming 1, Inverter 11~15)', 'QR-EQ035', 'AT-EQ035', 'Daily', 'Solar PV', 'Active', now],
    ['EQ036', 'A005', 'MSB-P2 (Incoming 2, Inverter 16~20)', 'QR-EQ036', 'AT-EQ036', 'Daily', 'Solar PV', 'Active', now],

    // A006 - Office FCU (4)
    ['EQ037', 'A006', 'CP-GF-OF-1A', 'QR-EQ037', 'AT-EQ037', 'Weekly', 'FCU Panel', 'Active', now],
    ['EQ038', 'A006', 'CP-GF-OF-1B', 'QR-EQ038', 'AT-EQ038', 'Weekly', 'FCU Panel', 'Active', now],
    ['EQ039', 'A006', 'CP-1F-OF-1A', 'QR-EQ039', 'AT-EQ039', 'Weekly', 'FCU Panel', 'Active', now],
    ['EQ040', 'A006', 'CP-1F-OF-1B', 'QR-EQ040', 'AT-EQ040', 'Weekly', 'FCU Panel', 'Active', now],

    // A007 - AHU Room (12)
    ['EQ041', 'A007', 'AHU 1', 'QR-EQ041', 'AT-EQ041', 'Daily', 'AHU', 'Active', now],
    ['EQ042', 'A007', 'AHU 2', 'QR-EQ042', 'AT-EQ042', 'Daily', 'AHU', 'Active', now],
    ['EQ043', 'A007', 'AHU 3', 'QR-EQ043', 'AT-EQ043', 'Daily', 'AHU', 'Active', now],
    ['EQ044', 'A007', 'AHU 4', 'QR-EQ044', 'AT-EQ044', 'Daily', 'AHU', 'Active', now],
    ['EQ045', 'A007', 'AHU 5', 'QR-EQ045', 'AT-EQ045', 'Daily', 'AHU', 'Active', now],
    ['EQ046', 'A007', 'CP-AHU 6', 'QR-EQ046', 'AT-EQ046', 'Daily', 'AHU Panel', 'Active', now],
    ['EQ047', 'A007', 'CP-AHU 6A MV', 'QR-EQ047', 'AT-EQ047', 'Daily', 'AHU Panel', 'Active', now],
    ['EQ048', 'A007', 'CP-AHU 7', 'QR-EQ048', 'AT-EQ048', 'Daily', 'AHU Panel', 'Active', now],
    ['EQ049', 'A007', 'CP-MIS & TES', 'QR-EQ049', 'AT-EQ049', 'Daily', 'AHU Panel', 'Active', now],
    ['EQ050', 'A007', 'CP-IF-CDA 7 MV', 'QR-EQ050', 'AT-EQ050', 'Daily', 'AHU Panel', 'Active', now],
    ['EQ051', 'A007', 'CDA 1 System', 'QR-EQ051', 'AT-EQ051', 'Daily', 'CDA System', 'Active', now],
    ['EQ052', 'A007', 'CDA 2 System', 'QR-EQ052', 'AT-EQ052', 'Daily', 'CDA System', 'Active', now],

    // A008 - Chiller Room (8)
    ['EQ053', 'A008', 'Chiller Incoming 1', 'QR-EQ053', 'AT-EQ053', 'Daily', 'Chiller', 'Active', now],
    ['EQ054', 'A008', 'Chiller 1', 'QR-EQ054', 'AT-EQ054', 'Daily', 'Chiller', 'Active', now],
    ['EQ055', 'A008', 'Cooling Tower Fan 1A', 'QR-EQ055', 'AT-EQ055', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ056', 'A008', 'Cooling Tower Fan 1B', 'QR-EQ056', 'AT-EQ056', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ057', 'A008', 'Chilled Water Pump 1', 'QR-EQ057', 'AT-EQ057', 'Daily', 'Pump', 'Active', now],
    ['EQ058', 'A008', 'Chilled Water Pump 2', 'QR-EQ058', 'AT-EQ058', 'Daily', 'Pump', 'Active', now],
    ['EQ059', 'A008', 'Chilled Water Pump 3', 'QR-EQ059', 'AT-EQ059', 'Daily', 'Pump', 'Active', now],
    ['EQ060', 'A008', 'Chiller Incoming 2', 'QR-EQ060', 'AT-EQ060', 'Daily', 'Chiller', 'Active', now],

    // A009 - L11 Testing Bay (12)
    ['EQ061', 'A009', 'SSB-L11-B1A', 'QR-EQ061', 'AT-EQ061', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ062', 'A009', 'SSB-L11-B1B', 'QR-EQ062', 'AT-EQ062', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ063', 'A009', 'SSB-L11-B2A', 'QR-EQ063', 'AT-EQ063', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ064', 'A009', 'SSB-L11-B2B', 'QR-EQ064', 'AT-EQ064', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ065', 'A009', 'SSB-L11-B3A', 'QR-EQ065', 'AT-EQ065', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ066', 'A009', 'SSB-L11-B3B', 'QR-EQ066', 'AT-EQ066', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ067', 'A009', 'SSB-L11-B4A', 'QR-EQ067', 'AT-EQ067', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ068', 'A009', 'SSB-L11-B4B', 'QR-EQ068', 'AT-EQ068', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ069', 'A009', 'SSB-L11-B5A', 'QR-EQ069', 'AT-EQ069', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ070', 'A009', 'SSB-L11-B6A', 'QR-EQ070', 'AT-EQ070', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ071', 'A009', 'SSB-L11-B7A', 'QR-EQ071', 'AT-EQ071', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ072', 'A009', 'SSB-L11-B11A', 'QR-EQ072', 'AT-EQ072', 'Daily', 'LV Switchgear', 'Active', now],

    // A010 - Production Line (17)
    ['EQ073', 'A010', 'SSB-LP1 (Lighting)', 'QR-EQ073', 'AT-EQ073', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ074', 'A010', 'SSB-LP1 (Power)', 'QR-EQ074', 'AT-EQ074', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ075', 'A010', 'SSB-LP1E (Lighting)', 'QR-EQ075', 'AT-EQ075', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ076', 'A010', 'SSB-LP1E (Power)', 'QR-EQ076', 'AT-EQ076', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ077', 'A010', 'SSB-LP2 (Lighting)', 'QR-EQ077', 'AT-EQ077', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ078', 'A010', 'SSB-LP2 (Power)', 'QR-EQ078', 'AT-EQ078', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ079', 'A010', 'SSB-LP1 (Lighting) B', 'QR-EQ079', 'AT-EQ079', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ080', 'A010', 'SSB-LP2E (Lighting)', 'QR-EQ080', 'AT-EQ080', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ081', 'A010', 'SSB-LP2E (Power)', 'QR-EQ081', 'AT-EQ081', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ082', 'A010', 'SSB-LP3 (Lighting)', 'QR-EQ082', 'AT-EQ082', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ083', 'A010', 'SSB-LP1 (Lighting) C', 'QR-EQ083', 'AT-EQ083', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ084', 'A010', 'SSB-LP3 (Power)', 'QR-EQ084', 'AT-EQ084', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ085', 'A010', 'SSB-LP3E (Lighting)', 'QR-EQ085', 'AT-EQ085', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ086', 'A010', 'SSB-LP3E (Power)', 'QR-EQ086', 'AT-EQ086', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ087', 'A010', 'SSB-P4E', 'QR-EQ087', 'AT-EQ087', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ088', 'A010', 'SSB-L10', 'QR-EQ088', 'AT-EQ088', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ089', 'A010', 'DB-P1', 'QR-EQ089', 'AT-EQ089', 'Daily', 'LV Switchgear', 'Active', now]
  ];

  equipment.forEach(eq => {
    if (!existingEquip.includes(eq[0])) {
      equipSheet.appendRow(eq);
    }
  });

  Logger.log('Plant 1 equipment seeded: ' + equipment.length + ' items.');
  Logger.log('seedData() complete. Plant 1 ready.');
}
