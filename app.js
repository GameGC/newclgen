const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const libre = require('libreoffice-convert');
const { promisify } = require('util');

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

        // Create today's date in "dd MMMM, yyyy" format
        const date = new Date();
        const dd = String(date.getDate()).padStart(2, '0');
        const monthNames = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ];
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear();
        const todayStr = `${dd} ${month}, ${year}`;

        // Replace placeholders
        const updatedXml = docXml
            .replace(/Today-date/g, todayStr)
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

const serverless = require('serverless-http');
module.exports = serverless(app);