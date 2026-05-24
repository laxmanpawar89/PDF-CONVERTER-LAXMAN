const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs/promises');
const { spawn } = require('child_process');

const app = express();

const upload = multer({ dest: path.join(__dirname, 'uploads') });

const PORT = process.env.PORT || 3001;

async function ensureDir(p) {
  await fs.mkdir(p, { recursive: true });
}

function runLibreOfficeConvert(inputPath, outputDir) {
  return new Promise((resolve, reject) => {
    // LibreOffice writes output as: <basename>.pdf in outputDir
    const child = spawn(
      'soffice',
      [
        '--headless',
        '--norestore',
        '--convert-to',
        'pdf',
        '--outdir',
        outputDir,
        inputPath,
      ],
      { windowsHide: true }
    );

    let stderr = '';
    child.stderr.on('data', (d) => (stderr += d.toString()));
    child.on('error', reject);

    child.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`LibreOffice conversion failed (code ${code}). ${stderr}`));
    });
  });
}

app.get('/health', (_req, res) => res.json({ ok: true }));

app.post('/convert-docx-to-pdf', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });


    const uploadDir = path.join(__dirname, 'uploads');
    const outDir = path.join(__dirname, 'outputs');
    await ensureDir(uploadDir);
    await ensureDir(outDir);

    const inputPath = req.file.path;
    const baseName = path.parse(req.file.originalname).name;
    const expectedPdf = path.join(outDir, `${baseName}.pdf`);

    // Clean old outputs with same name (best-effort)
    try {
      await fs.unlink(expectedPdf);
    } catch (_) { }

    await runLibreOfficeConvert(inputPath, outDir);

    // Wait a moment for file system settle (best-effort)
    let attempts = 0;
    while (attempts < 10) {
      try {
        await fs.stat(expectedPdf);
        break;
      } catch {
        attempts++;
        await new Promise((r) => setTimeout(r, 150));
      }
    }

    await fs.stat(expectedPdf);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}.pdf"`);

    await fs.readFile(expectedPdf).then((buf) => {
      res.status(200).send(buf);
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err?.message || String(err) });
  }
});

async function mergePdfsBuffers(pdfBuffers) {
  const { PDFDocument } = require('pdf-lib');

  const mergedPdf = await PDFDocument.create();

  for (const bytes of pdfBuffers) {
    const pdf = await PDFDocument.load(bytes);
    const pageIndices = pdf.getPageIndices();
    const copiedPages = await mergedPdf.copyPages(pdf, pageIndices);
    copiedPages.forEach((p) => mergedPdf.addPage(p));
  }

  return mergedPdf.save();
}

app.post('/merge-pdfs', upload.array('files'), async (req, res) => {
  try {
    if (!req.files || req.files.length < 2) {
      return res.status(400).json({ error: 'Upload at least 2 PDF files' });
    }

    // Only accept PDFs (best-effort validation)
    const pdfFiles = req.files.filter((f) => (f.mimetype || '').toLowerCase() === 'application/pdf');
    if (pdfFiles.length < 2) {
      return res.status(400).json({ error: 'All uploaded files must be PDFs' });
    }

    const pdfBuffers = [];
    for (const f of pdfFiles) {
      const buf = await fs.readFile(f.path);
      pdfBuffers.push(buf);
    }

    const mergedBytes = await mergePdfsBuffers(pdfBuffers);

    res.setHeader('Content-Type', 'application/pdf');
    const baseName = path.parse(pdfFiles[0].originalname || 'merged.pdf').name;
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}-merged.pdf"`);
    res.status(200).send(Buffer.from(mergedBytes));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err?.message || String(err) });
  }
});

app.listen(PORT, () => {
  console.log(`DOCX->PDF backend listening on http://localhost:${PORT}`);
});



