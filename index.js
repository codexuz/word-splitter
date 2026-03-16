const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const XLSX = require('xlsx');

const WORDS_PER_FILE = 30;
const OUTPUT_DIR = path.join(__dirname, 'output');

function selectExcelFile() {
  const psScript = [
    'Add-Type -AssemblyName System.Windows.Forms',
    '$dialog = New-Object System.Windows.Forms.OpenFileDialog',
    "$dialog.Title = 'Select an Excel file'",
    "$dialog.Filter = 'Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls'",
    "$dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')",
    "if ($dialog.ShowDialog() -eq 'OK') { Write-Output $dialog.FileName } else { Write-Output '' }"
  ].join('; ');

  const result = execSync(
    `powershell -NoProfile -ExecutionPolicy Bypass -Command "${psScript}"`,
    { encoding: 'utf-8' }
  ).trim();

  return result;
}

function readRowsFromExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const allRows = [];
  let headers = null;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || row.every(cell => cell == null || String(cell).trim() === '')) continue;

      if (!headers) {
        headers = row.map(cell => String(cell).trim());
        continue;
      }

      allRows.push(row);
    }
  }

  return { headers, rows: allRows };
}

function escapeCsvField(value) {
  const str = value == null ? '' : String(value);
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function main() {
  console.log('Opening file picker...');
  const excelPath = selectExcelFile();

  if (!excelPath) {
    console.log('No file selected. Exiting.');
    process.exit(0);
  }

  console.log(`Selected: ${excelPath}`);

  const { headers, rows } = readRowsFromExcel(excelPath);

  if (!headers || rows.length === 0) {
    console.log('No data found in the Excel file.');
    process.exit(0);
  }

  console.log(`Columns: ${headers.join(', ')}`);
  console.log(`Total rows: ${rows.length}`);
  console.log(`Files to create: ${Math.ceil(rows.length / WORDS_PER_FILE)}`);

  // Create output directory
  if (fs.existsSync(OUTPUT_DIR)) {
    fs.rmSync(OUTPUT_DIR, { recursive: true });
  }
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  // Split into chunks of 30 rows and write CSV files
  let fileIndex = 1;
  for (let i = 0; i < rows.length; i += WORDS_PER_FILE) {
    const chunk = rows.slice(i, i + WORDS_PER_FILE);
    const csvHeader = headers.map(escapeCsvField).join(',');
    const csvRows = chunk.map(row =>
      headers.map((_, colIdx) => escapeCsvField(row[colIdx])).join(',')
    );
    const csvContent = [csvHeader, ...csvRows].join('\n');

    const fileName = `words_part_${String(fileIndex).padStart(3, '0')}.csv`;
    const outPath = path.join(OUTPUT_DIR, fileName);
    fs.writeFileSync(outPath, csvContent, 'utf-8');

    console.log(`Created ${fileName} (${chunk.length} rows)`);
    fileIndex++;
  }

  console.log(`\nDone! ${fileIndex - 1} CSV files created in ./output/`);
}

main();
