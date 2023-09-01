$(document).ready(function () {
  $("#exportWorksheet").click(function () {
    var josnData = $("#josnData").val();
    var jsonDataObject = eval(josnData);
    exportWorksheet(jsonDataObject);
  });

  $("#exportWorksheetPlus").click(function () {
    var josnData = $("#josnData").val();
    var jsonDataObject = eval(josnData);
    exportWSPlus(jsonDataObject);
  });
});

function exportWorksheet(jsonObject) {
  var myFile = "myFile.xlsx";
  var myWorkSheet = XLSX.utils.json_to_sheet(jsonObject);
  var myWorkBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(myWorkBook, myWorkSheet, "myWorkSheet");
  XLSX.writeFile(myWorkBook, myFile);
}

function exportWSPlus(jsonObject) {
  var myFile = "myFilePlus.xlsx";
  var myWorkSheet = XLSX.utils.json_to_sheet(jsonObject);
  XLSX.utils.sheet_add_aoa(myWorkSheet, [["Your Mesage Goes Here"]], {
    origin: 0,
  });
  var merges = (myWorkSheet["!merges"] = [{ s: "A1", e: "D1" }]);
  var myWorkBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(myWorkBook, myWorkSheet, "myWorkSheet");
  XLSX.writeFile(myWorkBook, myFile);
}

const convert_to_utf8 = (code = "") => {
  return decodeURIComponent(escape(code));
};

// Return array of string values, or NULL if CSV string not well formed.
function CSVtoArray1(text) {
  var re_valid =
    /^\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*(?:,\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*)*$/;
  var re_value =
    /(?!\s*$)\s*(?:'([^'\\]*(?:\\[\S\s][^'\\]*)*)'|"([^"\\]*(?:\\[\S\s][^"\\]*)*)"|([^,'"\s\\]*(?:\s+[^,'"\s\\]+)*))\s*(?:,|$)/g;
  // Return NULL if input string is not well formed CSV string.
  if (!re_valid.test(text)) return null;
  var a = []; // Initialize array to receive values.
  text.replace(
    re_value, // "Walk" the string using replace with callback.
    function (m0, m1, m2, m3) {
      // Remove backslash from \' in single quoted values.
      if (m1 !== undefined) a.push(m1.replace(/\\'/g, "'"));
      // Remove backslash from \" in double quoted values.
      else if (m2 !== undefined) a.push(m2.replace(/\\"/g, '"'));
      else if (m3 !== undefined) a.push(m3);
      return ""; // Return empty string.
    }
  );
  // Handle special case of empty last value.
  if (/,\s*$/.test(text)) a.push("");
  return a;
}

function parseCSV(text) {
  // Obtenemos las lineas del texto
  let lines = text.replace(/\r/g, "").split("\n");
  return lines.map((line) => {
    // Por cada linea obtenemos los valores
    line_converted = convert_to_utf8(line);
    line_converted = line_converted.replaceAll(`""`, "");
    let values = CSVtoArray1(line_converted);
    return values;
  });
}

/* function reverseMatrix(matrix){
  let output = [];
  // Por cada fila
  matrix.forEach((values, row) => {
    // Vemos los valores y su posicion
    values.forEach((value, col) => {
      // Si la posición aún no fue creada
      if (output[col] === undefined) output[col] = [];
      output[col][row] = value;
    });
  });
  return output;
}
 */
function readFile(evt) {
  let file = evt.target.files[0];
  let reader = new FileReader();
  reader.onload = (e) => {
    // Cuando el archivo se terminó de cargar
    let lines = parseCSV(e.target.result);
    console.log(lines);
    /* let output = reverseMatrix(lines);
    console.log(output); */
  };
  // Leemos el contenido del archivo seleccionado
  reader.readAsBinaryString(file);
}

document.getElementById("file").addEventListener("change", readFile, false);

console.log(convert_to_utf8("Se corrigiÃ³")); // Se corrigió
console.log(convert_to_utf8("Se corrigiÃ³")); // Se corrigió
// let stringTest = `30,"Luis Solano",2ee36de2-65bb-6b56-b5a2-986328861417,38294,2023-08-24,2023-W34,"Desarrollo","Se quitó la frase ""Reporte de"" del dashboard principal.","Equipos APLI","FIX: Quitar ""Reporte de"" en el dashboard principal"`
// console.log(stringTest.replaceAll(`""`, ''))
// console.log(exportAsExcelFile())
