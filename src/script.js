// Exporta un Arreglo de objetos (JSON) a un Excel
const exportWorksheet = (jsonObject, myFile = "timelog.xlsx") => {
  let myWorkSheet = XLSX.utils.json_to_sheet(jsonObject);
  let myWorkBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(myWorkBook, myWorkSheet, "myWorkSheet");
  XLSX.writeFile(myWorkBook, `${myFile}.xlsx`);
};

// Codifica los caracteres especiales a utf8
const convert_to_utf8 = (code = "") => {
  return decodeURIComponent(escape(code));
};

// Convierte las líneas de un CSV en un Array
const CSVtoArray = (text) => {
  let re_valid =
    /^\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*(?:,\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*)*$/;
  let re_value =
    /(?!\s*$)\s*(?:'([^'\\]*(?:\\[\S\s][^'\\]*)*)'|"([^"\\]*(?:\\[\S\s][^"\\]*)*)"|([^,'"\s\\]*(?:\s+[^,'"\s\\]+)*))\s*(?:,|$)/g;
  if (!re_valid.test(text)) return null;
  let a = [];
  text.replace(re_value, (m0, m1, m2, m3) => {
    if (m1 !== undefined) a.push(m1.replace(/\\'/g, "'"));
    else if (m2 !== undefined) a.push(m2.replace(/\\"/g, '"'));
    else if (m3 !== undefined) a.push(m3);
    return "";
  });
  if (/,\s*$/.test(text)) a.push("");
  return a;
};

// Divide el CSV en línea y luego las formatea
const parseCSV = (text) => {
  let lines = text.replace(/\r/g, "").split("\n");
  return lines.map((line) => {
    line_converted = convert_to_utf8(line);
    line_converted = line_converted.replaceAll(`""`, "");
    let values = CSVtoArray(line_converted);
    return values;
  });
};

// Lee el archivo
const readFile = (evt) => {
  const month = document.getElementById("month").value;
  const year = document.getElementById("year").value;
  const firstName = document.getElementById("firstName").value;
  const lastName = document.getElementById("lastName").value;
  const analyst = document.getElementById("analyst").value;
  const environment = document.getElementById("environment").value;
  const internal = document.getElementById("internal").value;
  const formValues = {
    month,
    year,
    firstName,
    lastName,
    analyst,
    environment,
    internal,
  };
  console.log(formValues);
  let file = evt.target.files[0];
  let reader = new FileReader();
  reader.onload = (e) => {
    let lines = parseCSV(e.target.result);
    let newLines = [];
    let newLineHeader = [];
    newLineHeader[0] = "Project";
    newLineHeader[1] = "Issue Type";
    newLineHeader[2] = "Key";
    newLineHeader[3] = "Summary";
    newLineHeader[4] = "Analista";
    newLineHeader[5] = "Enviroment";
    newLineHeader[6] = "Interno";
    newLineHeader[7] = "Date Started";
    newLineHeader[8] = "Display Name";
    newLineHeader[9] = "Time Spent (h)";
    newLineHeader[10] = "Work Description";
    newLines.push(newLineHeader);
    for (let i = 1; i < lines.length; i++) {
      if (lines[i].length) {
        let newLine = [];
        newLine[0] = lines[i][8];
        newLine[1] = lines[i][6];
        newLine[2] = lines[i][3];
        newLine[3] = lines[i][9];
        newLine[4] = formValues.analyst;
        newLine[5] = formValues.environment;
        newLine[6] = formValues.internal;
        newLine[7] = lines[i][4];
        newLine[8] = lines[i][1];
        newLine[9] = lines[i][0] / 60;
        newLine[10] = lines[i][7];
        newLines.push(newLine);
      }
    }
    const timelog = [];
    for (let i = 1; i < newLines.length; i++) {
      let obj = {};
      for (let j = 0; j < newLines[0].length; j++) {
        obj = {
          ...obj,
          ...{ [newLines[0][j]]: newLines[i][j] },
        };
      }
      timelog.push(obj);
    }

    let fileName = `${formValues.month}_${formValues.year}_${formValues.firstName}_${formValues.lastName}`;
    let jsonDataObject = eval(timelog);
    exportWorksheet(jsonDataObject, fileName);
  };
  reader.readAsBinaryString(file);
};

document.getElementById("file").addEventListener("change", readFile, false);
