// api/files.js

var express = require('express');
var router = express.Router();
const { put, list } = require('@vercel/blob');
const fetch = require('node-fetch');

// ---------------------------
// Upload endpoint: POST /api/files/upload?filename=yourfile.docx
// ---------------------------
router.post('/upload', async (req, res) => {
  const { filename } = req.query;

  if (!filename) {
    return res.status(400).json({ message: 'Missing filename in query' });
  }

  try {
    const fileBuffer = await readRequestBody(req); // Read raw file

    if (!fileBuffer || fileBuffer.length === 0) {
      return res.status(400).json({ message: 'Empty file' });
    }

    const blob = await put(filename, fileBuffer, { access: 'public' });

    res.status(200).json({
      message: 'File uploaded successfully!',
      url: blob.url,
      pathname: blob.pathname,
    });

  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ message: error.message });
  }
});

// ---------------------------
// List endpoint: GET /api/files/list
// ---------------------------
router.get('/list', async (req, res) => {
  try {
    const { blobs } = await list();

    const files = blobs.map(blob => ({
      filename: blob.pathname.replace('/', ''),
      url: blob.url,
      size: blob.size,
      uploadedAt: blob.uploadedAt,
    }));

    res.status(200).json({ files });

  } catch (error) {
    console.error('List error:', error);
    res.status(500).json({ message: error.message });
  }
});

// ---------------------------
// Get file endpoint: GET /api/files/file/:filename
// ---------------------------
router.get('/file/:filename', async (req, res) => {
  const { filename } = req.params;

  if (!filename) {
    return res.status(400).json({ message: 'Filename is required' });
  }

  try {
    const blobUrl = `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}.vercel-blob.vercel-storage.com/${filename}`;

    const response = await fetch(blobUrl);
    if (!response.ok) {
      return res.status(404).json({ message: 'File not found' });
    }

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', response.headers.get('content-type'));

    response.body.pipe(res);

  } catch (error) {
    console.error('Get file error:', error);
    res.status(500).json({ message: error.message });
  }
});

// ---------------------------
// Utility to read raw stream
// ---------------------------
function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', (err) => reject(err));
  });
}

module.exports = router;
