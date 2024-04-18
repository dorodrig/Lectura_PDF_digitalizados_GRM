const fs = require("fs");
const pdf = require("pdf-parse");
const path = require("path");
const ExcelJS = require("exceljs");
// Ruta del archivo PDF
const baseFolderPath = `C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\PROCESADO\\`;

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
  return  newPDFPath;
}

async function moverPDFtoFolderNotprocess ( pdfPath,file_name_v3, extension){
  let pathnotprocces=`C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\NO PROCESADO\\`;
  let newpathnotprocces= path.join(pathnotprocces, `${file_name_v3}${extension}`);
// Mover el PDF al nuevo directorio
fs.renameSync(pdfPath, newpathnotprocces); 
return  newpathnotprocces;
}
async function extractDocumentNumberFromText(text) {
  const documentNumberRegex = /4300[\d\.-]+/g; // Modificamos la regex para que encuentre todos los números de documento
  const matches = text.match(documentNumberRegex);

  if (matches) {
    const clearNumbers = matches.map(match => match.replace(/[^\d]/g, ""));
    const uniqueNumbers = clearNumbers.filter((number, index) => clearNumbers.indexOf(number) === index);
    console.log(uniqueNumbers);
    return uniqueNumbers;
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
    const worksheet2 = workbook.addWorksheet("Mas de un comprobante"); // Nueva hoja para archivos con más de un comprobante
    
    // Agregar encabezados a las columnas
    worksheet.addRow([
      "Nombre de Archivo",
      "Fecha",
      "Número de Documento",
      "Ruta de almacenamiento",
    ]);
    worksheet1.addRow(["Nombre de Archivo","Ruta del archivo"]);
    worksheet2.addRow(["Nombre de Archivo","Ruta del archivo"]);
    // Iterar sobre cada archivo
    for (const file of files) {
      let file_name = file;
      let file_name_v2 = file_name.split('.');
      let file_name_v3 = file_name_v2[0];      
      // Verificar si el archivo es un PDF
      if (file.endsWith(".pdf")) { 
        //console.log(file.)
        // Construir la ruta completa del archivo PDF
        const pdfPath = path.join(folderPath, file);
        // Leer el contenido del archivo PDF
        const dataBuffer = fs.readFileSync(pdfPath);
        const data = await pdf(dataBuffer);
        const text = data.text;
        // Extraer el número de documento del texto
        const documentNumbers = await extractDocumentNumberFromText(text);
        //console.log(documentNumbers)
        var pdfnumero = documentNumbers && documentNumbers.length > 0 ? documentNumbers[0] : null; // Tomamos solo el primer número de documento
        // Si no se pudo extraer el número de documento, agregar el nombre del archivo al libro de Excel de archivos no procesados
        if (!documentNumbers || documentNumbers.length === 0) {
          const pathnoproccessfile= await moverPDFtoFolderNotprocess(pdfPath,file_name_v3, extension);
          worksheet1.addRow([file,pathnoproccessfile]);
          console.log(`Archivo no procesado: ${pdfPath}`);
        } else {
           // Extraer la fecha del texto
           const date = await extractDateFromText(text);

           // Crear las carpetas necesarias
           const monthFolderPath = await createDirectoriesIfNeeded(date);
 
           // Mover el PDF al directorio creado
          const newpath = await movePDFToFolder(pdfPath, monthFolderPath, pdfnumero, extension);
          // Si hay más de un número de documento, agregar el nombre del archivo a la hoja correspondiente
          if (documentNumbers.length > 1) {
            worksheet2.addRow([file,newpath]);
            console.log(`Archivo con más de un comprobante: ${pdfPath}`);
          }

         

          // Agregar una fila con los datos al libro de Excel
          console.log(`Documento PDF movido a: ${newpath}`);
          worksheet.addRow([file, date, documentNumbers, newpath]);
          console.log(`PDF procesado: ${pdfPath}`);
          console.log(`Fecha extraída: ${date}`);
          console.log(`Número de documento extraído: ${documentNumbers}`);
        }
      }
    }
    // Guardar el libro de Excel en la ruta especificada
    let ruta_archivo_xlsx= "C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF";
    const excelFilePath = path.join(ruta_archivo_xlsx, "datos.xlsx");
    await workbook.xlsx.writeFile(excelFilePath);
  } catch (error) {
    console.error("Error al leer el PDF:", error);
  }
}
// Llamar a la función con la ruta de la carpeta que contiene los PDFs
const folderPath = `C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\DOCUMENTOS OCR\\`;
extractImagesAndText(folderPath);
