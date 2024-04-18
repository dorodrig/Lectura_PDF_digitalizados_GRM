const fs = require('fs');
const pdf = require('pdf-parse');


const pdfnumero = 'Doc1'
//const pdfnumero = '4300060040'
const extension = '.pdf'
//`C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\DOCUMENTOS OCR\\`;
const pdfPath = `C:\\Users\\David.Rodriguez\\OneDrive - GRM Colombia S.A.S\\Escritorio\\OCR PDF\\DOCUMENTOS OCR\\${pdfnumero}${extension}`;
async function extractImagesAndText() {
    try {
        const dataBuffer = fs.readFileSync(pdfPath);        
        const data = await pdf(dataBuffer);   
        var data2 = JSON.stringify(data,null,2);
        // Accede a las imágenes y al texto extraído
        const images = data.numrender; // Número de imágenes en el PDF
        const text = data.text; // Texto extraído del PDF
        console.log(`datos: ${data2}`);
        console.log(`Número de imágenes en el PDF: ${images}`);
        console.log(`Texto extraído:${text}`);
    } catch (error) {
        console.error('Error al leer el PDF:', error);
    }
}

extractImagesAndText();
