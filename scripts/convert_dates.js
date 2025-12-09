const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

(async () => {
  const EXCEL_FILE = path.join(__dirname, '..', 'saved_texts.xlsx');
  if (!fs.existsSync(EXCEL_FILE)) {
    console.error('Excel file not found:', EXCEL_FILE);
    process.exit(1);
  }

  const backup = EXCEL_FILE.replace(/\.xlsx$/i, `.backup.${Date.now()}.xlsx`);
  fs.copyFileSync(EXCEL_FILE, backup);
  console.log('Backup created at', backup);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const worksheet = workbook.getWorksheet('User Details');
  if (!worksheet) {
    console.error('Worksheet "User Details" not found');
    process.exit(1);
  }

  // find Date column index
  let dateCol = null;
  worksheet.getRow(1).eachCell((cell, colNumber) => {
    if (cell && String(cell.value).trim() === 'Date') dateCol = colNumber;
  });

  if (!dateCol) {
    console.error('Date column not found');
    process.exit(1);
  }

  let changed = 0;
  for (let r = 2; r <= worksheet.actualRowCount; r++) {
    const row = worksheet.getRow(r);
    const cell = row.getCell(dateCol);
    const raw = cell.value;
    if (!raw) continue;
    if (raw instanceof Date) {
      // normalize existing Date cells to date-only (midnight)
      const dateOnly = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
      cell.value = dateOnly;
      cell.numFmt = 'dd-mm-yyyy';
      changed++;
      continue;
    }
    // try parse
    let parsed = null;
    const s = String(raw).trim();
    const iso = Date.parse(s);
    if (!isNaN(iso)) parsed = new Date(iso);
    else {
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (m) {
        let day = parseInt(m[1],10);
        let month = parseInt(m[2],10)-1;
        let year = parseInt(m[3],10);
        if (year < 100) year += 2000;
        const dt = new Date(year, month, day);
        if (!isNaN(dt)) parsed = dt;
      }
    }
    if (parsed) {
      // ensure time is zeroed (date-only)
      parsed.setHours(0,0,0,0);
      cell.value = parsed;
      cell.numFmt = 'dd-mm-yyyy';
      changed++;
    }
  }

  await workbook.xlsx.writeFile(EXCEL_FILE);
  console.log(`Conversion complete. ${changed} cells updated.`);
})();
