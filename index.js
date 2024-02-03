const fs = require('fs');
require('dotenv').config()
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const docxConverter = require('docx-pdf');


//generate temp docx file
function replacePlaceholders(templatePath, data) {
  try {
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip);

    // Set the data for replacing placeholders
    doc.setData(data);

    // Render current document
    doc.render();

    const generatedDoc = doc.getZip().generate({ type: 'nodebuffer' });
    return generatedDoc;
  } catch (error) {
    throw new Error(`Error replacing placeholders: ${error.message}`);
  }
}


//docx to pdf convertion function
function convertToPdf(docxBuffer, outputPath) {
    return new Promise((resolve, reject) => {
      fs.writeFileSync('temp.docx', docxBuffer);
  
      docxConverter('./temp.docx', outputPath, function (err, result) {
        fs.unlinkSync('temp.docx'); // Remove temporary DOCX file
        if (err) {
          reject(new Error(`Error converting to PDF: ${err.message}`));
        } else {
          console.log(`PDF generated successfully at: ${outputPath}`);
          resolve();
        }
      });
    });
  }

//inputs
const sampleData = {
  firstName: process.env.FIRSTNAME,
  lastName: process.env.LASTNAME,
  age:process.env.AGE,
  companyName: process.env.COMPANY,
};

const templatePath = './Documents/wordDoc.docx';   //will take the word file having dynamic placeholders
const outputPath = './Documents/Generated/document.pdf';     //path to save the pdf file

try {
  const docxBuffer = replacePlaceholders(templatePath, sampleData);
  convertToPdf(docxBuffer, outputPath);
} catch (err) {
  console.error(err.message);
}
