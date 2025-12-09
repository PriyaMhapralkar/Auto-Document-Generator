const Excel = require('exceljs');
const path = require('path');
const fs = require('fs');

(async () => {
  try {
    const EXCEL_FILE_PATH = path.join(__dirname, '..', 'saved_texts.xlsx');

    if (!fs.existsSync(EXCEL_FILE_PATH)) {
      console.log('No Excel file found at', EXCEL_FILE_PATH);
      process.exit(0);
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupPath = path.join(__dirname, '..', `saved_texts.backup.${timestamp}.xlsx`);

    // Create backup
    fs.copyFileSync(EXCEL_FILE_PATH, backupPath);
    console.log('✓ Backup created at', backupPath);

    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);

    const worksheet = workbook.getWorksheet('User Details') || workbook.worksheets[0];

    if (!worksheet) {
      console.log('No worksheets found in the workbook. Nothing to do.');
      process.exit(0);
    }

    
    let colToRemove = null;
    const headerRow = worksheet.getRow(1);
    
    headerRow.eachCell((cell, colNumber) => {
      const val = cell && cell.value ? String(cell.value).trim() : '';
      console.log(`Column ${colNumber}: "${val}"`);
      if (val === 'General Text Input') {
        colToRemove = colNumber;
      }
    });

    if (!colToRemove) {
      console.log('✓ No "General Text Input" column found. Headers are correct.');
      
      // Print current headers for verification
      console.log('\nCurrent headers in Excel:');
      headerRow.eachCell((cell, colNumber) => {
        const val = cell && cell.value ? String(cell.value).trim() : '';
        console.log(`  ${colNumber}. ${val}`);
      });
      
      process.exit(0);
    }

    console.log(`✓ Found "General Text Input" at column ${colToRemove}. Removing...`);

    // Remove the column
    worksheet.spliceColumns(colToRemove, 1);

    await workbook.xlsx.writeFile(EXCEL_FILE_PATH);
    console.log('✓ Excel file updated successfully.');
    
    // Print new headers for verification
    console.log('\nNew headers in Excel:');
    const newHeaderRow = worksheet.getRow(1);
    newHeaderRow.eachCell((cell, colNumber) => {
      const val = cell && cell.value ? String(cell.value).trim() : '';
      console.log(`  ${colNumber}. ${val}`);
    });
    
    process.exit(0);
  } catch (err) {
    console.error('❌ Error:', err.message);
    process.exit(1);
  }
})();
