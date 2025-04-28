const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const libre = require('libreoffice-convert');
const { promisify } = require('util');

const Docxtemplater = require('docxtemplater');
const {convert} = require('docx2pdf-converter');
const { tmpdir } = require('os');

const app = express();

const convertAsync = promisify(libre.convert);

app.get('/api/generate-pdf', async (req, res) => {
    try {
        const msgText = req.query.msgText || 'Default Text';

        // Read DOCX template as binary
        const templatePath = path.resolve(__dirname, 'data', 'cl-template.docx');
        const content = await fs.promises.readFile(templatePath);

        // Load content into zip
        const zip = new PizZip(content);
        const docXml = zip.file('word/document.xml').asText();

        // Replace placeholders
        const updatedXml = docXml
            .replace(/Today-date/g, TodayDateFormatted())
            .replace(/\[CL-TEXT\]/g, msgText);

        zip.file('word/document.xml', updatedXml);

        // Generate new DOCX
        const updatedDocx = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });

        // Convert DOCX buffer to PDF
        const pdfBuf = await convertAsync(updatedDocx, '.pdf', undefined);

        // Send the PDF file as a download
        res.set({
            'Content-Type': 'application/pdf',
            'Content-Disposition': 'attachment; filename="generated.pdf"',
        });
        res.send(pdfBuf);
    } catch (error) {
        console.error('Error generating PDF:', error);
        res.status(500).send('Failed to generate PDF.');
    }
});

app.get('/api/generate-docx', async (req, res) => {
    try {
        const msgText = req.query.msgText || 'Default Text';

        // Read DOCX template
        const templatePath = path.resolve(__dirname, 'data', 'cl-template.docx');
        const content = await fs.promises.readFile(templatePath, 'binary');

        // Load DOCX into PizZip
        const zip = new PizZip(content);

        // Read original document.xml
        const documentXml = zip.file('word/document.xml').asText();

        // Replace placeholders
        const updatedXml = documentXml
            .replace(/Today-date/g, TodayDateFormatted())
            .replace(/\[CL-TEXT\]/g, msgText);

        // Overwrite document.xml
        zip.file('word/document.xml', updatedXml);

        // Create Docxtemplater instance with modified zip
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Generate final modified DOCX buffer
        const bufferDocx = doc.getZip().generate({ type: 'nodebuffer' });

        // Send DOCX file
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename="generated.docx"',
        });
        res.send(bufferDocx);
    } catch (error) {
        console.error('Error generating DOCX:', error);
        res.status(500).send('Failed to generate DOCX.');
    }
});
app.get('/api/generate-pdff', async (req, res) => {
    try {
        const msgText = req.query.msgText || 'Default Text';

        // Read DOCX template
        const templatePath = path.resolve(__dirname, 'data', 'cl-template.docx');
        const content = await fs.promises.readFile(templatePath, 'binary');

        // Load DOCX into PizZip
        const zip = new PizZip(content);

        // Read original document.xml
        const documentXml = zip.file('word/document.xml').asText();

        // Replace placeholders
        const updatedXml = documentXml
            .replace(/Today-date/g, TodayDateFormatted())
            .replace(/\[CL-TEXT\]/g, msgText);

        // Overwrite document.xml
        zip.file('word/document.xml', updatedXml);

        // Create Docxtemplater instance with modified zip
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Generate final modified DOCX buffer
        const bufferDocx = doc.getZip().generate({ type: 'nodebuffer' });

        const tmpFilePath = path.resolve(__dirname, 'data', 'generated.docx');
        const tmpFilePath2 = path.resolve(__dirname, 'data', 'generated.pdf');

       // const tmpFilePath = path.join(tmpdir(), 'generated.docx');
       //x const tmpFilePath2 = path.join(tmpdir(), 'generated.pdf');
        fs.writeFileSync(tmpFilePath, bufferDocx);


        // Convert DOCX buffer to PDF

        await convert(tmpFilePath,tmpFilePath2);

        var pdfBuffer = fs.readFileSync(tmpFilePath2);

                // Send the generated PDF as a response
                res.set({
                    'Content-Type': 'application/pdf',
                    'Content-Disposition': 'attachment; filename="generated.pdf"',
                });
                res.send(pdfBuffer);

    } catch (error) {
        console.error('Error generating DOCX:', error);
        res.status(500).send('Failed to generate DOCX.');
    }
});



function TodayDateFormatted(){
    // Prepare today's date
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    const month = monthNames[today.getMonth()];
    const year = today.getFullYear();

// Format date as "13 August, 2024"
    return `${day} ${month}, ${year}`;
}

module.exports = app;