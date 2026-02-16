/* global alert, console, document, Excel, FileReader, Office, TextDecoder */



let fileInput = null;
let ColumnKeyName = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    fileInput = document.getElementById("fileInput");
    const run = document.getElementById("run");

    fileInput.addEventListener("change", () => {
      if (fileInput.files && fileInput.files.length > 0) {
        console.log("File selezionato:", fileInput.files[0].name);
        document.getElementById("fileName").textContent =
          "Selected file: " + fileInput.files[0].name;
      }
    });

    run.addEventListener("click", UpgradeExcel);
  }
});

//it's the main function of the program
async function UpgradeExcel() {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
    console.error("Nessun file selezionato");
    alert("Seleziona prima un file CSV");
    return;
  }

  // disabilita run durante l'esecuzione
  document.getElementById("run").disabled = true;

  try {
    const CSVFile = fileInput.files[0];
    const ReaderResult = await ReadFile(CSVFile);
    const FormattedCSV = processCSV(ReaderResult);

    await Excel.run(async (context) => {
      let currWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = currWorksheet.getUsedRange();
      range.load("values");
      await context.sync();

      let ExcelData = range.values;
      ExcelData = findInExcel(FormattedCSV, ExcelData);

      let writingRange = currWorksheet.getRangeByIndexes(0, 0, ExcelData.length, ExcelData[0].length);
      writingRange.values = ExcelData;

      await formatColumns(ExcelData, currWorksheet, context);

      writingRange.format.autofitColumns();
      writingRange.format.autofitRows();

      await context.sync();

      console.log("Importazione completata con successo!");
      alert("Dati aggiornati con successo!");
    });
  } catch (error) {
    console.error("Errore:", error);
    alert("Errore durante l'importazione: " + error.message);
  } finally {
    // Hide loading
    document.getElementById("loading").style.display = "none";
    document.getElementById("run").disabled = false;
  }
}

//read from file csv and return an arraybuffer with its content
function ReadFile(file) {
  return new Promise((resolve, reject) => {
    const Reader = new FileReader();
    Reader.onload = () => resolve(Reader.result);
    Reader.onerror = reject;
    Reader.readAsArrayBuffer(file);
  });
}

//return an array with the first column's data, this because the unique values are often in the first column
function TakeKeyColumn(currValue) {
  if (!currValue || currValue.length === 0) return [];

  let keyColumnIndex = -1;

  for (let c = 0; c < currValue[0].length; c++) {
    if (String(currValue[0][c]).toLowerCase() === "matricola") {
      keyColumnIndex = c;
      break;
    }
  }

  if (keyColumnIndex === -1) {
    for (let c = 0; c < currValue[0].length; c++) {
      let seen = new Set();
      let unique = true;

      for (let r = 1; r < currValue.length; r++) {
        if (seen.has(currValue[r][c])) {
          unique = false;
          break;
        }
        seen.add(currValue[r][c]);
      }

      if (unique) {
        keyColumnIndex = c;
        break;
      }
    }
  }

  if (keyColumnIndex === -1) return [];

  return keyColumnIndex;
}



//split the CSV, it works with , and ; in order to create a matrix 2D
function processCSV(ReaderResult) {
  let resultText = new TextDecoder().decode(ReaderResult).trim();

  return resultText.split(/\r?\n/).map((rows) => {
    let row = formatForSplitting(rows);
    let separator = row.includes(";") ? ";" : ",";
    return row.split(separator).map(cellFormat);
  });
}

//changes the format of decimals found in the row necessary to have a correct splitting 
function formatForSplitting(rows) {
  return rows.replace(/"(\d+),(\d+)"/g, "$1.$2");
}

//sets the cell format date, number or string
function cellFormat(cell) {
  cell = cell.trim();

  // control for the date, it can handle separator / and - and different setting for numbers 
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(cell)) {
    return cell; // is mantained as a string
  }

  let str = cell.toString();
  let num = Number(str);
  return isNaN(num) ? str : num;
}

// search the content of CSV first column into the first column of file excel and write in it
// decimals are summed and strings are overwritten by the one in cvs
function findInExcel(FormattedCSV, ExcelData) {
  let keyindexCSV = TakeKeyColumn(FormattedCSV);
  let keyindexExcel = TakeKeyColumn(ExcelData);

  let keyColumnCSV = FormattedCSV.slice(1).map(row => row[keyindexCSV]);
  let keyColumnExcel = ExcelData.slice(1).map(row => row[keyindexExcel]);

  for (let csv = 0; csv < keyColumnCSV.length; csv++) {
    let excelRowIndex = keyColumnExcel.indexOf(keyColumnCSV[csv]);

    if (excelRowIndex !== -1) {
      let indexRow = excelRowIndex + 1;

      for (let col = 0; col < ExcelData[0].length; col++) {
        let csvValue = FormattedCSV[csv + 1][col];
        let columnName = ExcelData[0][col];
        if (columnName.toLowerCase().includes("mese")) {
          ExcelData[indexRow][col] = csvValue;  // ✅ Sovrascrivi solo questa cella
          continue;  // Opzionale: salta il resto del loop per questa colonna
        }
        if (col === keyindexExcel) continue;

        if (typeof csvValue === 'number' && typeof ExcelData[indexRow][col] === 'number') {
          ExcelData[indexRow][col] += csvValue;
        } else if (ExcelData[indexRow][col] != csvValue) {
          ExcelData[indexRow][col] = csvValue;
        }
      }
    }
  }

  return ExcelData;
}


//set format of the column for dates mmm-yyyy and for numbers with . for the thousands and , for the decimal part, negative number are -#### and written in red, with separators as explained before
async function formatColumns(ExcelData, worksheet, context) {
  for (let c = 0; c < ExcelData[0].length; c++) {
    let columnName = ExcelData[0][c].toString().toLowerCase();

    let hasDate = false;
    let hasDecimal = false;

    for (let r = 1; r < ExcelData.length; r++) {
      let cellValue = ExcelData[r][c];

      //date check
      if (typeof cellValue === 'string' && /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(cellValue)) {
        hasDate = true;
        break;
      }

      //decimal check
      if (typeof cellValue === 'number' && !Number.isInteger(cellValue)) {
        hasDecimal = true;
      }
    }

    let column = worksheet.getRangeByIndexes(1, c, ExcelData.length - 1, 1);

    if (hasDate) {
      column.numberFormat = [["mmm-yyyy"]];
    } else if (hasDecimal) {
      column.numberFormat = [["#,##0.00;[Red]-#,##0.00"]];
    }
  }
  await context.sync();
}
