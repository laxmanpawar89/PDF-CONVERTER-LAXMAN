/*
  PDF -> JPG (fully in-browser via PDF.js canvas rendering)
  PDF -> DOCX (best-effort text extraction via PDF.js textContent; Word generated via docx library)
*/

const pdfInput = document.getElementById('pdfInput');
const docxInput = document.getElementById('docxInput');
const convertJpgBtn = document.getElementById('convertJpgBtn');
const convertDocxBtn = document.getElementById('convertDocxBtn');
const convertPdfFromDocxBtn = document.getElementById('convertPdfFromDocxBtn');

const pdfMergeInput = document.getElementById('pdfMergeInput');
const jpgMergeInput = document.getElementById('jpgMergeInput');
const mergePdfBtn = document.getElementById('mergePdfBtn');
const mergeJpgBtn = document.getElementById('mergeJpgBtn');

const clearBtn = document.getElementById('clearBtn');
const statusEl = document.getElementById('status');
const outputEl = document.getElementById('output');
const qualityEl = document.getElementById('quality');
const scaleEl = document.getElementById('scale');


function setStatus(msg) {
  statusEl.textContent = msg;
}

function getQuality() {
  const q = Number(qualityEl.value);
  if (!Number.isFinite(q)) return 0.92;
  return Math.min(1, Math.max(0.1, q));
}

function getScale() {
  const s = Number(scaleEl.value);
  if (!Number.isFinite(s)) return 1.5;
  return Math.min(3, Math.max(0.5, s));
}

function clearOutput() {
  outputEl.innerHTML = '';
}

function requirePDFjs() {
  const pdfjsLib = window['pdfjsLib'];
  const pdfjsGlobal = window['pdfjsDist'] || window['pdfjs'];
  const getDocument = pdfjsLib?.getDocument || pdfjsGlobal?.getDocument;

  if (typeof getDocument !== 'function') {
    throw new Error(
      'PDF.js failed to load: cannot find getDocument. If you are offline or the CDN is blocked, this will not work.'
    );
  }

  return getDocument;
}

async function convertPdfToJpg(file) {
  clearOutput();

  const arrayBuffer = await file.arrayBuffer();
  const getDocument = requirePDFjs();
  const pdf = await getDocument({ data: arrayBuffer }).promise;

  const numPages = pdf.numPages;
  const quality = getQuality();
  const scale = getScale();

  setStatus(`Loaded "${file.name}". Pages: ${numPages}. Converting to JPG...`);
  convertJpgBtn.disabled = true;
  convertDocxBtn.disabled = true;
  clearBtn.disabled = false;

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    setStatus(`Rendering page ${pageNum}/${numPages} ...`);

    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d', { willReadFrequently: false });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({ canvasContext: context, viewport }).promise;

    const dataUrl = canvas.toDataURL('image/jpeg', quality);

    const card = document.createElement('div');
    card.className = 'card';

    const img = document.createElement('img');
    img.src = dataUrl;
    img.alt = `Page ${pageNum} as JPG`;

    const meta = document.createElement('div');
    meta.className = 'meta';

    const title = document.createElement('div');
    title.className = 'title';
    title.textContent = `Page ${pageNum}`;

    const dl = document.createElement('div');
    dl.className = 'dl';

    const download = document.createElement('a');
    download.className = 'download';

    const baseName = file.name.replace(/\.pdf$/i, '');
    download.href = dataUrl;
    download.download = `${baseName}-page-${pageNum}.jpg`;
    download.textContent = 'Download JPG';

    dl.appendChild(download);

    meta.appendChild(title);
    meta.appendChild(dl);

    card.appendChild(img);
    card.appendChild(meta);

    outputEl.appendChild(card);
  }

  setStatus(`Done. Converted ${numPages} page(s) to JPG.`);
  convertJpgBtn.disabled = false;
  convertDocxBtn.disabled = false;
}

async function convertPdfToDocx(file) {
  clearOutput();

  const arrayBuffer = await file.arrayBuffer();
  const getDocument = requirePDFjs();
  const pdf = await getDocument({ data: arrayBuffer }).promise;

  const numPages = pdf.numPages;

  // docx is expected from CDN: window.docx
  const docx = window['docx'];
  if (!docx || !docx.Document || !docx.Packer || !docx.Paragraph || !docx.TextRun) {
    throw new Error('DOCX generator failed to load. Ensure docx CDN is reachable.');
  }

  const { Document, Packer, Paragraph, TextRun } = docx;

  setStatus(`Loaded "${file.name}". Pages: ${numPages}. Converting to DOCX (text-based)...`);
  convertJpgBtn.disabled = true;
  convertDocxBtn.disabled = true;
  clearBtn.disabled = false;

  const paras = [];

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    setStatus(`Extracting text from page ${pageNum}/${numPages} ...`);

    const page = await pdf.getPage(pageNum);

    const textContent = await page.getTextContent();
    const strings = (textContent.items || []).map((it) => it.str).filter(Boolean);

    if (strings.length === 0) {
      paras.push(new Paragraph({ children: [new TextRun({ text: `Page ${pageNum}` })] }));
      continue;
    }

    const text = strings.join(' ');
    paras.push(
      new Paragraph({
        children: [new TextRun({ text: `Page ${pageNum}`, bold: true })],
      })
    );
    paras.push(new Paragraph({ children: [new TextRun({ text })] }));
  }

  const baseName = file.name.replace(/\.pdf$/i, '');

  const doc = new Document({
    sections: [{
      properties: {},
      children: paras,
    }],
  });

  const blob = await Packer.toBlob(doc);

  const url = URL.createObjectURL(blob);

  const card = document.createElement('div');
  card.className = 'card';

  const meta = document.createElement('div');
  meta.className = 'meta';

  const title = document.createElement('div');
  title.className = 'title';
  title.textContent = 'DOCX Download';

  const dl = document.createElement('div');
  dl.className = 'dl';

  const download = document.createElement('a');
  download.className = 'download';
  download.href = url;
  download.download = `${baseName}.docx`;
  download.textContent = 'Download DOCX';

  dl.appendChild(download);

  meta.appendChild(title);
  meta.appendChild(dl);

  card.appendChild(meta);

  outputEl.appendChild(card);

  setStatus(`Done. Generated DOCX from extracted text.`);

  convertJpgBtn.disabled = false;
  convertDocxBtn.disabled = false;
}


pdfInput.addEventListener('change', () => {
  const file = pdfInput.files && pdfInput.files[0];
  clearOutput();

  if (!file) {
    convertJpgBtn.disabled = true;
    convertDocxBtn.disabled = true;
    clearBtn.disabled = true;
    setStatus('');
    return;
  }

  if (file.type !== 'application/pdf') {
    setStatus('Please choose a valid PDF file.');
    convertJpgBtn.disabled = true;
    convertDocxBtn.disabled = true;
    clearBtn.disabled = true;
    return;
  }

  convertJpgBtn.disabled = false;
  convertDocxBtn.disabled = false;
  clearBtn.disabled = false;
  setStatus(`Selected: "${file.name}". Ready.`);
});

convertJpgBtn.addEventListener('click', async () => {
  const file = pdfInput.files && pdfInput.files[0];
  if (!file) return;

  try {
    await convertPdfToJpg(file);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    convertJpgBtn.disabled = false;
    convertDocxBtn.disabled = false;
  }
});

convertDocxBtn.addEventListener('click', async () => {
  const file = pdfInput.files && pdfInput.files[0];
  if (!file) return;

  try {
    await convertPdfToDocx(file);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    convertJpgBtn.disabled = false;
    convertDocxBtn.disabled = false;
  }
});

clearBtn.addEventListener('click', () => {
  clearOutput();
  setStatus('Cleared. Upload a new PDF to convert.');
});

async function convertDocxToPdf(file) {
  clearOutput();

  const baseName = file.name.replace(/\.docx$/i, '');

  const formData = new FormData();
  formData.append('file', file);

  setStatus('Uploading DOCX and converting to PDF (server-side)...');
  convertJpgBtn.disabled = true;
  convertDocxBtn.disabled = true;
  convertPdfFromDocxBtn.disabled = true;
  clearBtn.disabled = false;

  // backend is expected at localhost:3001
  const resp = await fetch('http://localhost:3001/convert-docx-to-pdf', {
    method: 'POST',
    body: formData,
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    throw new Error(`Server conversion failed (${resp.status}). ${errText}`);
  }

  const blob = await resp.blob();
  const url = URL.createObjectURL(blob);

  const card = document.createElement('div');
  card.className = 'card';

  const meta = document.createElement('div');
  meta.className = 'meta';

  const title = document.createElement('div');
  title.className = 'title';
  title.textContent = 'PDF Download';

  const dl = document.createElement('div');
  dl.className = 'dl';

  const download = document.createElement('a');
  download.className = 'download';
  download.href = url;
  download.download = `${baseName}.pdf`;
  download.textContent = 'Download PDF';

  dl.appendChild(download);
  meta.appendChild(title);
  meta.appendChild(dl);
  card.appendChild(meta);
  outputEl.appendChild(card);

  setStatus('Done. DOCX converted to PDF.');
  convertJpgBtn.disabled = false;
  convertDocxBtn.disabled = false;
  convertPdfFromDocxBtn.disabled = false;
}

function updateButtonStates() {
  const pdfFile = pdfInput.files && pdfInput.files[0];
  const docxFile = docxInput.files && docxInput.files[0];

  convertJpgBtn.disabled = !(pdfFile && pdfFile.type === 'application/pdf');
  convertDocxBtn.disabled = !(pdfFile && pdfFile.type === 'application/pdf');
  convertPdfFromDocxBtn.disabled = !docxFile;

  const pdfMergeFiles = pdfMergeInput?.files ? Array.from(pdfMergeInput.files) : [];
  const jpgMergeFiles = jpgMergeInput?.files ? Array.from(jpgMergeInput.files) : [];

  const pdfOk = pdfMergeFiles.length >= 2 && pdfMergeFiles.every((f) => f.type === 'application/pdf');
  const jpgOk = jpgMergeFiles.length >= 2 && jpgMergeFiles.every((f) => f.type === 'image/jpeg');

  mergePdfBtn.disabled = !pdfOk;
  mergeJpgBtn.disabled = !jpgOk;
}


pdfInput.addEventListener('change', () => {
  clearOutput();
  if (!pdfInput.files || !pdfInput.files[0]) {
    convertJpgBtn.disabled = true;
    convertDocxBtn.disabled = true;
    setStatus('Upload a PDF to begin.');
    clearBtn.disabled = true;
    return;
  }

  const file = pdfInput.files[0];
  if (file.type !== 'application/pdf') {
    setStatus('Please choose a valid PDF file.');
    convertJpgBtn.disabled = true;
    convertDocxBtn.disabled = true;
    clearBtn.disabled = true;
    return;
  }

  clearBtn.disabled = false;
  setStatus(`Selected: "${file.name}". Ready.`);
  updateButtonStates();
});

docxInput.addEventListener('change', () => {
  clearOutput();
  if (!docxInput.files || !docxInput.files[0]) {
    convertPdfFromDocxBtn.disabled = true;
    setStatus('Upload a DOCX to convert to PDF.');
    clearBtn.disabled = true;
    return;
  }

  const file = docxInput.files[0];
  if (!file.type.includes('officedocument.wordprocessingml.document')) {
    setStatus('Please choose a valid DOCX file.');
    convertPdfFromDocxBtn.disabled = true;
    clearBtn.disabled = true;
    return;
  }

  clearBtn.disabled = false;
  setStatus(`Selected: "${file.name}". Ready.`);
  updateButtonStates();
});

convertJpgBtn.addEventListener('click', async () => {
  const file = pdfInput.files && pdfInput.files[0];
  if (!file) return;

  try {
    await convertPdfToJpg(file);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    updateButtonStates();
  }
});

convertDocxBtn.addEventListener('click', async () => {
  const file = pdfInput.files && pdfInput.files[0];
  if (!file) return;

  try {
    await convertPdfToDocx(file);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    updateButtonStates();
  }
});

convertPdfFromDocxBtn.addEventListener('click', async () => {
  const file = docxInput.files && docxInput.files[0];
  if (!file) return;

  try {
    await convertDocxToPdf(file);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    updateButtonStates();
  }
});

clearBtn.addEventListener('click', () => {
  clearOutput();
  setStatus('Cleared. Upload files to convert.');
  clearBtn.disabled = true;
  convertJpgBtn.disabled = true;
  convertDocxBtn.disabled = true;
  convertPdfFromDocxBtn.disabled = true;
  // Note: we intentionally do not clear the <input> values.
});

// Merge handlers
pdfMergeInput.addEventListener('change', () => {
  clearOutput();
  updateButtonStates();
  if (!pdfMergeInput.files || pdfMergeInput.files.length === 0) {
    setStatus('Upload PDFs to merge.');
  } else {
    setStatus(`Selected ${pdfMergeInput.files.length} PDF(s) to merge.`);
  }
});

jpgMergeInput.addEventListener('change', () => {
  clearOutput();
  updateButtonStates();
  if (!jpgMergeInput.files || jpgMergeInput.files.length === 0) {
    setStatus('Upload JPGs to merge.');
  } else {
    setStatus(`Selected ${jpgMergeInput.files.length} JPG(s) to merge.`);
  }
});

function toJpegDownloadCard({ dataUrl, filename, titleText }) {
  const card = document.createElement('div');
  card.className = 'card';

  const img = document.createElement('img');
  img.src = dataUrl;
  img.alt = filename;

  const meta = document.createElement('div');
  meta.className = 'meta';

  const title = document.createElement('div');
  title.className = 'title';
  title.textContent = titleText;

  const dl = document.createElement('div');
  dl.className = 'dl';

  const download = document.createElement('a');
  download.className = 'download';
  download.href = dataUrl;
  download.download = filename;
  download.textContent = 'Download JPG';

  dl.appendChild(download);
  meta.appendChild(title);
  meta.appendChild(dl);
  card.appendChild(img);
  card.appendChild(meta);
  outputEl.appendChild(card);
}

async function mergeJpgStacked(files) {
  clearOutput();

  const quality = getQuality(); // existing JPG Quality setting

  const imgs = files.map((f, idx) => ({ file: f, idx, img: new Image() }));

  const loadImg = (imgObj) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        imgObj.img.onload = () => resolve();
        imgObj.img.onerror = reject;
        imgObj.img.src = String(reader.result);
      };
      reader.onerror = reject;
      reader.readAsDataURL(imgObj.file);
    });
  };

  setStatus(`Loading ${imgs.length} image(s) ...`);
  convertJpgBtn.disabled = true;
  convertDocxBtn.disabled = true;
  mergePdfBtn.disabled = true;
  mergeJpgBtn.disabled = true;
  clearBtn.disabled = false;

  for (const it of imgs) await loadImg(it);

  const maxW = Math.max(...imgs.map((it) => it.img.naturalWidth || it.img.width));
  const widths = maxW;

  let totalH = 0;
  const scaled = imgs.map((it) => {
    const w = it.img.naturalWidth || it.img.width;
    const h = it.img.naturalHeight || it.img.height;
    const scale = widths / w;
    const sh = Math.round(h * scale);
    totalH += sh;
    return { img: it.img, sh };
  });

  setStatus('Rendering stacked JPG...');
  const canvas = document.createElement('canvas');
  canvas.width = widths;
  canvas.height = totalH;

  const ctx = canvas.getContext('2d');
  let y = 0;
  for (const it of scaled) {
    const w = widths;
    const h = it.sh;
    ctx.drawImage(it.img, 0, y, w, h);
    y += h;
  }

  const dataUrl = canvas.toDataURL('image/jpeg', quality);
  const baseName = (files[0].name || 'merged').replace(/\.(jpe?g)$/i, '');
  const filename = `${baseName}-merged.jpg`;

  toJpegDownloadCard({
    dataUrl,
    filename,
    titleText: `Merged JPG (stacked, ${files.length} file(s))`,
  });

  setStatus('Done. Merged JPG generated.');
  convertJpgBtn.disabled = false;
  convertDocxBtn.disabled = false;
  updateButtonStates();
}

async function mergePdfs(files) {
  clearOutput();

  setStatus(`Uploading ${files.length} PDF(s) and merging...`);
  convertJpgBtn.disabled = true;
  convertDocxBtn.disabled = true;
  mergePdfBtn.disabled = true;
  mergeJpgBtn.disabled = true;
  clearBtn.disabled = false;

  const formData = new FormData();
  for (const f of files) formData.append('files', f, f.name);

// NOTE: server.js currently does NOT implement /merge-pdfs.
// Keep this call, but it will fail unless the backend route exists.
const resp = await fetch('http://localhost:3001/merge-pdfs', {
    method: 'POST',
    body: formData,
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    throw new Error(`Server merge failed (${resp.status}). ${errText}`);
  }

  const blob = await resp.blob();
  const url = URL.createObjectURL(blob);

  const baseName = (files[0].name || 'merged').replace(/\.pdf$/i, '');
  const filename = `${baseName}-merged.pdf`;

  const card = document.createElement('div');
  card.className = 'card';

  const meta = document.createElement('div');
  meta.className = 'meta';

  const title = document.createElement('div');
  title.className = 'title';
  title.textContent = `Merged PDF (${files.length} file(s))`;

  const dl = document.createElement('div');
  dl.className = 'dl';

  const download = document.createElement('a');
  download.className = 'download';
  download.href = url;
  download.download = filename;
  download.textContent = 'Download PDF';

  dl.appendChild(download);
  meta.appendChild(title);
  meta.appendChild(dl);
  card.appendChild(meta);
  outputEl.appendChild(card);

  setStatus('Done. Merged PDF generated.');
  convertJpgBtn.disabled = false;
  convertDocxBtn.disabled = false;
  updateButtonStates();
}

mergeJpgBtn.addEventListener('click', async () => {
  try {
    const files = Array.from(jpgMergeInput.files || []);
    if (files.length < 2) return;
    await mergeJpgStacked(files);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    updateButtonStates();
  }
});

mergePdfBtn.addEventListener('click', async () => {
  try {
    const files = Array.from(pdfMergeInput.files || []);
    if (files.length < 2) return;
    await mergePdfs(files);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err?.message || String(err)}`);
    updateButtonStates();
  }
});

// Initial state
convertJpgBtn.disabled = true;
convertDocxBtn.disabled = true;
convertPdfFromDocxBtn.disabled = true;
mergePdfBtn.disabled = true;
mergeJpgBtn.disabled = true;
clearBtn.disabled = true;
setStatus('Upload a PDF or DOCX to begin.');




