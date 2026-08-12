// ==========================================
// KONFIGURASI DATABASE SPREADSHEET
// ==========================================
const SPREADSHEET_ID = '1MEq3tn_6jhCrSTQB1fdq1ai13SXyPRccKaxcwPSUXLM';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ==========================================
// HANDLER ROUTING UNTUK VERCEL & WEB APP (GET & POST)
// ==========================================
function doGet(e) {
  // Jika dipanggil oleh Vercel via fetch dengan parameter action
  if (e && e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    
    if (action === 'getKaryawan') {
      return ContentService.createTextOutput(getDataKaryawan())
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'getUnit') {
      return ContentService.createTextOutput(getDataUnit())
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Tampilan Default jika dibuka langsung di Google Apps Script
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SE Dashboard - Monitoring Karyawan & Unit')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    let result = { success: false, message: "Aksi tidak dikenal" };

    if (action === 'saveKaryawan') {
      result = saveKaryawan(contents.data);
    } else if (action === 'deleteKaryawan') {
      result = deleteKaryawan(contents.id);
    } else if (action === 'importKaryawanBulk') {
      result = importKaryawanBulk(contents.data);
    } else if (action === 'saveUnit') {
      result = saveUnit(contents.data);
    } else if (action === 'deleteUnit') {
      result = deleteUnit(contents.id);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// SETUP INITIAL DATABASE
// ==========================================
function setupDatabase() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Data Karyawan');
  if (!sheet) {
    sheet = ss.insertSheet('Data Karyawan');
    const headers = [
      "ID", "NRP", "Nama", "Perusahaan", "Jabatan", "TglLahir", 
      "ExpSIMPER_BIB", "ExpSIMPER_TIA", "ExpSIM_B2", "TglMCU", "ExpMCU", 
      "TglLOTOTO", "ExpDangerTag", "TglAwareness", "ExpKetinggian", "TglSIO", "ExpSIO"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f8fafc");
    sheet.setFrozenRows(1);
  }

  let unitSheet = ss.getSheetByName('Data Unit');
  if (!unitSheet) {
    unitSheet = ss.insertSheet('Data Unit');
    const unitHeaders = ["ID", "NoUnit", "Model", "ExpKom_BIB", "ExpKom_TIA", "ExpKom_TMA"];
    unitSheet.getRange(1, 1, 1, unitHeaders.length).setValues([unitHeaders]);
    unitSheet.getRange(1, 1, 1, unitHeaders.length).setFontWeight("bold").setBackground("#f8fafc");
    unitSheet.setFrozenRows(1);
  }
  return "Setup Database Selesai!";
}

// ==========================================
// FUNGSI CRUD KARYAWAN
// ==========================================
function getDataKaryawan() {
  const sheet = getSpreadsheet().getSheetByName('Data Karyawan');
  if (!sheet) return JSON.stringify([]);
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return JSON.stringify([]);
  
  const data = sheet.getRange(1, 1, lastRow, 17).getValues();
  const headers = data[0];
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // Skip jika NRP kosong
    
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) {
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        val = `${y}-${m}-${d}`;
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return JSON.stringify(result); 
}

function saveKaryawan(formData) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Data Karyawan');
    const data = sheet.getDataRange().getValues();
    const id = formData.ID || Utilities.getUuid();
    
    const rowData = [
      id, formData.NRP || "", formData.Nama || "", formData.Perusahaan || "", formData.Jabatan || "", formData.TglLahir || "",
      formData.ExpSIMPER_BIB || "", formData.ExpSIMPER_TIA || "", formData.ExpSIM_B2 || "", formData.TglMCU || "", formData.ExpMCU || "",
      formData.TglLOTOTO || "", formData.ExpDangerTag || "", formData.TglAwareness || "", formData.ExpKetinggian || "", formData.TglSIO || "", formData.ExpSIO || ""
    ];

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == formData.ID) { rowIndex = i + 1; break; }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, message: "Data berhasil diperbarui!" };
    } else {
      sheet.appendRow(rowData);
      return { success: true, message: "Data baru berhasil ditambahkan!" };
    }
  } catch (err) { return { success: false, error: err.message }; }
}

function deleteKaryawan(id) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Data Karyawan');
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Data berhasil dihapus!" };
      }
    }
    return { success: false, message: "ID tidak ditemukan." };
  } catch (err) { return { success: false, message: err.message }; }
}

function importKaryawanBulk(jsonData) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Data Karyawan');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let updatedCount = 0; let newCount = 0;

    jsonData.forEach(item => {
      if(!item.NRP) return; 
      const nrpString = String(item.NRP).trim();
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) { if (String(data[i][1]).trim() === nrpString) { rowIndex = i + 1; break; } }

      let rowData = new Array(headers.length).fill("");
      if (rowIndex > -1) { rowData = data[rowIndex - 1]; } else { rowData[0] = Utilities.getUuid(); }

      headers.forEach((header, index) => {
        if (index === 0) return;
        if (item[header] !== undefined) { rowData[index] = item[header]; }
      });

      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
        updatedCount++;
      } else {
        sheet.appendRow(rowData);
        data.push(rowData);
        newCount++;
      }
    });

    return { success: true, message: `Import Selesai! ${newCount} Data Baru, ${updatedCount} Data Diperbarui.` };
  } catch (err) { return { success: false, error: err.message }; }
}

// ==========================================
// FUNGSI CRUD UNIT
// ==========================================
function getDataUnit() {
  const sheet = getSpreadsheet().getSheetByName('Data Unit');
  if (!sheet) return JSON.stringify([]);
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return JSON.stringify([]);
  
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];
  const result = [];
  
  for(let i=1; i<data.length; i++) {
    const row = data[i];
    if(!row[1]) continue;
    let obj = {};
    for(let j=0; j<headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) {
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        val = `${y}-${m}-${d}`;
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return JSON.stringify(result);
}

function saveUnit(formData) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Data Unit');
    const data = sheet.getDataRange().getValues();
    const id = formData.ID || Utilities.getUuid();
    const noUnit = (formData.NoUnit || "").trim().toUpperCase();

    const rowData = [ id, noUnit, formData.Model || "", formData.ExpKom_BIB || "", formData.ExpKom_TIA || "", formData.ExpKom_TMA || "" ];
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) { if (data[i][0] == formData.ID || data[i][1] == noUnit) { rowIndex = i + 1; break; } }

    if (rowIndex > -1) { sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]); return { success: true, message: "Data Unit diperbarui!" }; } 
    else { sheet.appendRow(rowData); return { success: true, message: "Data Unit baru ditambahkan!" }; }
  } catch (err) { return { success: false, error: err.message }; }
}

function deleteUnit(id) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Data Unit');
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true, message: "Unit dihapus!" }; } }
    return { success: false, message: "ID Unit tidak ditemukan." };
  } catch (err) { return { success: false, message: err.message }; }
}
