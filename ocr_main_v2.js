const fs = require("fs");
const pdf = require("pdf-parse");
const path = require("path");
const ExcelJS = require("exceljs");
// Ruta del archivo PDF
const baseFolderPath = `C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\`;

async function extractDateFromText(text) {
  const lines = text.split("\n");
  let date = null;

  // Buscar la fecha después de encontrar '\n*'
  let startIndex = lines.findIndex((line) => line.includes("\n*"));
  if (startIndex !== -1) {
    for (let i = startIndex + 1; i < lines.length; i++) {
      const line = lines[i].trim();
      const possibleDate = line.match(/\d{2}\.\d{2}\.\d{4}/);

      if (possibleDate) {
        // Separar la fecha en día, mes y año
        const [day, month, year] = possibleDate[0].split(".");
        // Crear un objeto Date con el formato adecuado (año, mes - 1, día)
        date = new Date(year, month - 1, day);
        break;
      }
    }
  }

  // Si no se encontró la fecha después de '\n*', buscar la primera fecha en cualquier parte del texto
  if (!date) {
    const possibleDate = text.match(/\d{2}\.\d{2}\.\d{4}/);

    if (possibleDate) {
      const [day, month, year] = possibleDate[0].split(".");
      date = new Date(year, month - 1, day);
    }
  }

  return date;
}
async function createDirectoriesIfNeeded(date) {
  const year = date.getFullYear();
  const monthNames = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
  ];
  const month = monthNames[date.getMonth()];

  const yearFolderPath = path.join(baseFolderPath, year.toString());
  const monthFolderPath = path.join(yearFolderPath, month);
  //console.log("Month Folder Path:", monthFolderPath);

  // Crear la carpeta del año si no existe
  if (!fs.existsSync(yearFolderPath)) {
    fs.mkdirSync(yearFolderPath);
  }

  // Crear la carpeta del mes si no existe
  if (!fs.existsSync(monthFolderPath)) {
    fs.mkdirSync(monthFolderPath);
  }

  return monthFolderPath;
}
async function movePDFToFolder(pdfPath, folderPath, pdfnumero, extension) {
  const newPDFPath = path.join(folderPath, `${pdfnumero}${extension}`);

  // Mover el PDF al nuevo directorio
  fs.renameSync(pdfPath, newPDFPath);

  console.log(`Documento PDF movido a: ${newPDFPath}`);
}
async function extractDocumentNumberFromText(text) {
  const documentNumberRegex = /4300[\d\.-]+/;
  const match = text.match(documentNumberRegex);

  if (match) {
    const clearnumber = match[0].replace(/[^\d]/g, ""); // Eliminar caracteres no numéricos
    return clearnumber;
  } else {
    return null;
  }
}

async function extractImagesAndText(folderPath) {
  try {
    const extension = ".pdf";
    const files = fs.readdirSync(folderPath);
    // Crear un nuevo libro de Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Datos");
    const worksheet1 = workbook.addWorksheet("Archivos no procesados");
    // Agregar encabezados a las columnas
    worksheet.addRow([
      "Nombre de Archivo",
      "Fecha",
      "Número de Documento",
      "Ruta de almacenamiento",
    ]);
    worksheet1.addRow(["Nombre de Archivo"]);
    // Iterar sobre cada archivo
    for (const file of files) {
      // Verificar si el archivo es un PDF
      if (file.endsWith(".pdf")) {
        // Construir la ruta completa del archivo PDF
        const pdfPath = path.join(folderPath, file);
        // Leer el contenido del archivo PDF
        const dataBuffer = fs.readFileSync(pdfPath);
        const data = await pdf(dataBuffer);
        const text = data.text;
        // Extraer el número de documento del texto
        const documentNumber = await extractDocumentNumberFromText(text);
        var pdfnumero = documentNumber;
        // Si no se pudo extraer el número de documento, agregar el nombre del archivo al libro de Excel de archivos no procesados
        if (!documentNumber) {
          worksheet1.addRow([file]);
          console.log(`Archivo no procesado: ${pdfPath}`);
        } else {
          // Extraer la fecha del texto
          const date = await extractDateFromText(text);

          // Crear las carpetas necesarias
          const monthFolderPath = await createDirectoriesIfNeeded(date);

          // Mover el PDF al directorio creado
          await movePDFToFolder(pdfPath, monthFolderPath, pdfnumero, extension);

          // Agregar una fila con los datos al libro de Excel
          worksheet.addRow([file, date, documentNumber, pdfPath]);
          console.log(`PDF procesado: ${pdfPath}`);
          console.log(`Fecha extraída: ${date}`);
          console.log(`Número de documento extraído: ${documentNumber}`);
        }
      }
    }
    // Guardar el libro de Excel en la ruta especificada
    const excelFilePath = path.join(folderPath, "datos.xlsx");
    await workbook.xlsx.writeFile(excelFilePath);
  } catch (error) {
    console.error("Error al leer el PDF:", error);
  }
}
// Llamar a la función con la ruta de la carpeta que contiene los PDFs
const folderPath = `C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\DOCUMENTOS OCR\\`;
extractImagesAndText(folderPath);
