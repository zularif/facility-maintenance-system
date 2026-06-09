// ============================================================
// SeedDataP2.gs — Plant 2 Seed Data
// 8 Areas, 83 Equipment
// Run seedDataP2() once. Does NOT modify Plant 1 data.
// ============================================================

function seedDataP2() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const now = new Date().toISOString();

  // ── PLANT 2 ──────────────────────────────────────────────
  const plantSheet = ss.getSheetByName('Plants');
  const existingPlants = plantSheet.getDataRange().getValues().map(r => r[0]);
  if (!existingPlants.includes('P2')) {
    plantSheet.appendRow(['P2', 'Plant 2', 'Secondary Industrial Area', 'Active', now]);
    Logger.log('Plant 2 added.');
  } else {
    Logger.log('Plant 2 already exists, skipping.');
  }

  // ── AREAS ─────────────────────────────────────────────────
  const areaSheet = ss.getSheetByName('Areas');
  const existingAreas = areaSheet.getDataRange().getValues().map(r => r[0]);

  const areas = [
    ['A011', 'P2', 'Consumer Room HT P2', 'High Voltage Consumer Room P2', 'Active', now],
    ['A012', 'P2', 'Transformer Room P2', 'Transformer Room Plant 2', 'Active', now],
    ['A013', 'P2', 'MSB LV Room P2', 'Main Switch Board LV Room P2', 'Active', now],
    ['A014', 'P2', 'Production Hygrometer P2', 'Temperature and Humidity Monitoring P2', 'Active', now],
    ['A015', 'P2', 'Solar Panel P2', 'Solar PV System P2', 'Active', now],
    ['A016', 'P2', 'Chiller Room P2', 'Chiller Plant Room P2', 'Active', now],
    ['A017', 'P2', 'Production Line L1 P2', 'Production Line DIP/GF P2', 'Active', now],
    ['A018', 'P2', 'Production Line L2 P2', 'Production Line SMT/SF P2', 'Active', now]
  ];

  areas.forEach(area => {
    if (!existingAreas.includes(area[0])) {
      areaSheet.appendRow(area);
    }
  });
  Logger.log('Areas for Plant 2 seeded.');

  // ── EQUIPMENT ─────────────────────────────────────────────
  const equipSheet = ss.getSheetByName('Equipment');
  const existingEquip = equipSheet.getDataRange().getValues().map(r => r[0]);

  const equipment = [
    // A011 - Consumer Room HT 33kV P2 (9)
    ['EQ090', 'A011', 'Feeder 33kV Incomer 1', 'QR-EQ090', 'AT-EQ090', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ091', 'A011', 'Transformer Feeder 2M1', 'QR-EQ091', 'AT-EQ091', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ092', 'A011', 'Transformer Feeder 2M3', 'QR-EQ092', 'AT-EQ092', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ093', 'A011', 'Transformer Feeder 1C5', 'QR-EQ093', 'AT-EQ093', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ094', 'A011', 'Transformer Feeder 2P2', 'QR-EQ094', 'AT-EQ094', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ095', 'A011', 'Feeder 33kV Incomer 2', 'QR-EQ095', 'AT-EQ095', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ096', 'A011', 'Transformer Feeder 2M2', 'QR-EQ096', 'AT-EQ096', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ097', 'A011', 'Transformer Feeder 2M4', 'QR-EQ097', 'AT-EQ097', 'Daily', 'HV Switchgear', 'Active', now],
    ['EQ098', 'A011', 'Transformer Feeder 2P1', 'QR-EQ098', 'AT-EQ098', 'Daily', 'HV Switchgear', 'Active', now],

    // A012 - Transformer Room P2 (7)
    ['EQ099', 'A012', 'TX-2M1', 'QR-EQ099', 'AT-EQ099', 'Daily', 'Transformer', 'Active', now],
    ['EQ100', 'A012', 'TX-2M3', 'QR-EQ100', 'AT-EQ100', 'Daily', 'Transformer', 'Active', now],
    ['EQ101', 'A012', 'TX-2P2', 'QR-EQ101', 'AT-EQ101', 'Daily', 'Transformer', 'Active', now],
    ['EQ102', 'A012', 'TX-2P1', 'QR-EQ102', 'AT-EQ102', 'Daily', 'Transformer', 'Active', now],
    ['EQ103', 'A012', 'TX-OLIVIA 1C5', 'QR-EQ103', 'AT-EQ103', 'Daily', 'Transformer', 'Active', now],
    ['EQ104', 'A012', 'TX-2M4', 'QR-EQ104', 'AT-EQ104', 'Daily', 'Transformer', 'Active', now],
    ['EQ105', 'A012', 'TX-2M2', 'QR-EQ105', 'AT-EQ105', 'Daily', 'Transformer', 'Active', now],

    // A013 - MSB LV Room P2 (6)
    ['EQ106', 'A013', 'MSBII-M1', 'QR-EQ106', 'AT-EQ106', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ107', 'A013', 'MSBII-M2', 'QR-EQ107', 'AT-EQ107', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ108', 'A013', 'MSBII-M3', 'QR-EQ108', 'AT-EQ108', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ109', 'A013', 'MSBII-M4', 'QR-EQ109', 'AT-EQ109', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ110', 'A013', 'MSBII-P1', 'QR-EQ110', 'AT-EQ110', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ111', 'A013', 'MSBII-P2', 'QR-EQ111', 'AT-EQ111', 'Daily', 'LV Switchgear', 'Active', now],

    // A014 - Production Hygrometer P2 (13)
    ['EQ112', 'A014', 'Packaging (Temp)', 'QR-EQ112', 'AT-EQ112', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ113', 'A014', 'Packaging (Humidity)', 'QR-EQ113', 'AT-EQ113', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ114', 'A014', 'DIP 1', 'QR-EQ114', 'AT-EQ114', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ115', 'A014', 'DIP 2', 'QR-EQ115', 'AT-EQ115', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ116', 'A014', 'DIP 3', 'QR-EQ116', 'AT-EQ116', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ117', 'A014', 'DIP 4', 'QR-EQ117', 'AT-EQ117', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ118', 'A014', 'DIP 5', 'QR-EQ118', 'AT-EQ118', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ119', 'A014', 'DIP 6', 'QR-EQ119', 'AT-EQ119', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ120', 'A014', 'SMT 1', 'QR-EQ120', 'AT-EQ120', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ121', 'A014', 'SMT 2', 'QR-EQ121', 'AT-EQ121', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ122', 'A014', 'SMT 3', 'QR-EQ122', 'AT-EQ122', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ123', 'A014', 'SMT 4', 'QR-EQ123', 'AT-EQ123', 'Daily', 'Hygrometer', 'Active', now],
    ['EQ124', 'A014', 'SMT 5 & 6', 'QR-EQ124', 'AT-EQ124', 'Daily', 'Hygrometer', 'Active', now],

    // A015 - Solar Panel P2 (2)
    ['EQ125', 'A015', 'Solar PV MSB 1', 'QR-EQ125', 'AT-EQ125', 'Daily', 'Solar PV', 'Active', now],
    ['EQ126', 'A015', 'Solar PV MSB 2', 'QR-EQ126', 'AT-EQ126', 'Daily', 'Solar PV', 'Active', now],

    // A016 - Chiller Room P2 (20)
    ['EQ127', 'A016', 'Chiller Incoming 1', 'QR-EQ127', 'AT-EQ127', 'Daily', 'Chiller', 'Active', now],
    ['EQ128', 'A016', 'Chiller 3', 'QR-EQ128', 'AT-EQ128', 'Daily', 'Chiller', 'Active', now],
    ['EQ129', 'A016', 'Chiller 1', 'QR-EQ129', 'AT-EQ129', 'Daily', 'Chiller', 'Active', now],
    ['EQ130', 'A016', 'Cooling Tower 3A', 'QR-EQ130', 'AT-EQ130', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ131', 'A016', 'Cooling Tower 3B', 'QR-EQ131', 'AT-EQ131', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ132', 'A016', 'Cooling Tower 1A', 'QR-EQ132', 'AT-EQ132', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ133', 'A016', 'Cooling Tower 1B', 'QR-EQ133', 'AT-EQ133', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ134', 'A016', 'Cooling Tower 2A', 'QR-EQ134', 'AT-EQ134', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ135', 'A016', 'Cooling Tower 2B', 'QR-EQ135', 'AT-EQ135', 'Daily', 'Cooling Tower', 'Active', now],
    ['EQ136', 'A016', 'Chiller 2', 'QR-EQ136', 'AT-EQ136', 'Daily', 'Pump', 'Active', now],
    ['EQ137', 'A016', 'Chilled Water Pump 1', 'QR-EQ137', 'AT-EQ137', 'Daily', 'Pump', 'Active', now],
    ['EQ138', 'A016', 'Process Cooling Water Pump 1', 'QR-EQ138', 'AT-EQ138', 'Daily', 'Pump', 'Active', now],
    ['EQ139', 'A016', 'Chiller Incoming 2', 'QR-EQ139', 'AT-EQ139', 'Daily', 'Chiller', 'Active', now],
    ['EQ140', 'A016', 'Chiller 2 (second)', 'QR-EQ140', 'AT-EQ140', 'Daily', 'Chiller', 'Active', now],
    ['EQ141', 'A016', 'Chilled Water Pump 2', 'QR-EQ141', 'AT-EQ141', 'Daily', 'Pump', 'Active', now],
    ['EQ142', 'A016', 'Chilled Water Pump 3', 'QR-EQ142', 'AT-EQ142', 'Daily', 'Pump', 'Active', now],
    ['EQ143', 'A016', 'Chilled Water Pump 4', 'QR-EQ143', 'AT-EQ143', 'Daily', 'Pump', 'Active', now],
    ['EQ144', 'A016', 'Condenser Water Pump 2', 'QR-EQ144', 'AT-EQ144', 'Daily', 'Pump', 'Active', now],
    ['EQ145', 'A016', 'Condenser Water Pump 3', 'QR-EQ145', 'AT-EQ145', 'Daily', 'Pump', 'Active', now],
    ['EQ146', 'A016', 'Condenser Water Pump 4', 'QR-EQ146', 'AT-EQ146', 'Daily', 'Pump', 'Active', now],

    // A017 - Production Line P2 DIP/GF (7)
    ['EQ147', 'A017', 'SSBII-FF-DIP-LO1', 'QR-EQ147', 'AT-EQ147', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ148', 'A017', 'SSBII-FF-DIP-LO2', 'QR-EQ148', 'AT-EQ148', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ149', 'A017', 'SSBII-FF-DIP-LO3', 'QR-EQ149', 'AT-EQ149', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ150', 'A017', 'SSBII-FF-DIP-LO4', 'QR-EQ150', 'AT-EQ150', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ151', 'A017', 'SSBII-FF-DIP-LO5', 'QR-EQ151', 'AT-EQ151', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ152', 'A017', 'SSBII-FF-DIP-LO6', 'QR-EQ152', 'AT-EQ152', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ153', 'A017', 'SSBII-GF-PACK', 'QR-EQ153', 'AT-EQ153', 'Daily', 'LV Switchgear', 'Active', now],

    // A018 - Production Line P2 SMT/SF (19)
    ['EQ154', 'A018', 'SSBII-SF-SMT-LO1', 'QR-EQ154', 'AT-EQ154', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ155', 'A018', 'SSBII-SF-SMT-LO2', 'QR-EQ155', 'AT-EQ155', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ156', 'A018', 'SSBII-SF-SMT-LO3', 'QR-EQ156', 'AT-EQ156', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ157', 'A018', 'SSBII-SF-SMT-LO4', 'QR-EQ157', 'AT-EQ157', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ158', 'A018', 'SSBII-SF-SMT-LO5', 'QR-EQ158', 'AT-EQ158', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ159', 'A018', 'SSBII-SF-SMT-LO6', 'QR-EQ159', 'AT-EQ159', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ160', 'A018', 'Packaging', 'QR-EQ160', 'AT-EQ160', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ161', 'A018', 'DIP 1', 'QR-EQ161', 'AT-EQ161', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ162', 'A018', 'DIP 2', 'QR-EQ162', 'AT-EQ162', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ163', 'A018', 'DIP 3', 'QR-EQ163', 'AT-EQ163', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ164', 'A018', 'DIP 4', 'QR-EQ164', 'AT-EQ164', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ165', 'A018', 'DIP 5', 'QR-EQ165', 'AT-EQ165', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ166', 'A018', 'DIP 6', 'QR-EQ166', 'AT-EQ166', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ167', 'A018', 'SMT 1', 'QR-EQ167', 'AT-EQ167', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ168', 'A018', 'SMT 2', 'QR-EQ168', 'AT-EQ168', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ169', 'A018', 'SMT 3', 'QR-EQ169', 'AT-EQ169', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ170', 'A018', 'SMT 4', 'QR-EQ170', 'AT-EQ170', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ171', 'A018', 'SMT 5', 'QR-EQ171', 'AT-EQ171', 'Daily', 'LV Switchgear', 'Active', now],
    ['EQ172', 'A018', 'SMT 6', 'QR-EQ172', 'AT-EQ172', 'Daily', 'LV Switchgear', 'Active', now]
  ];

  equipment.forEach(eq => {
    if (!existingEquip.includes(eq[0])) {
      equipSheet.appendRow(eq);
    }
  });

  Logger.log('Plant 2 equipment seeded: ' + equipment.length + ' items.');
  Logger.log('seedDataP2() complete. Plant 2 ready.');
}
