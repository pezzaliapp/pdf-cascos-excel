import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

const excelInput = document.getElementById('excelFile');
const pdf1Input = document.getElementById('pdfFile1');
const pdf2Input = document.getElementById('pdfFile2');
const analyzeBtn = document.getElementById('analyzeBtn');
const generateBtn = document.getElementById('generateBtn');
const downloadJsonBtn = document.getElementById('downloadJsonBtn');
const statusEl = document.getElementById('status');
const previewBody = document.querySelector('#previewTable tbody');

let excelRows = [];
let pdfPages = [];
let matchedRows = [];

function setStatus(msg) {
  statusEl.textContent = msg;
}

function normalize(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9\.\- ]/gi, '')
    .trim();
}

function findKey(row, candidates) {
  const keys = Object.keys(row);
  for (const cand of candidates) {
    const found = keys.find(k => normalize(k).includes(normalize(cand)));
    if (found) return found;
  }
  return null;
}

function mapExcelRows(rows) {
  return rows.map((row, idx) => {
    const codeKey = findKey(row, ['codice', 'codice prodotto', 'ref', 'rif']);
    const nameKey = findKey(row, ['nome', 'nome prodotto', 'prodotto', 'modello', 'descrizione']);
    const priceKey = findKey(row, ['prezzo', 'prezzo vendita', 'listino']);
    const categoryKey = findKey(row, ['categoria', 'famiglia', 'gruppo']);

    return {
      originalIndex: idx + 2,
      codice: row[codeKey] ?? '',
      nome: row[nameKey] ?? '',
      prezzo: row[priceKey] ?? '',
      categoria: row[categoryKey] ?? '',
      raw: row
    };
  }).filter(r => r.codice || r.nome);
}

async function readExcel(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  return mapExcelRows(json);
}

async function renderPageToImage(page, scale = 1.2) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas.toDataURL('image/jpeg', 0.82);
}

async function readPdf(file, label) {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const text = textContent.items.map(item => item.str).join(' ');
    const image = await renderPageToImage(page);
    pages.push({
      source: file.name,
      sourceLabel: label,
      pageNumber: i,
      text,
      image
    });
  }
  return pages;
}

function scoreMatch(row, page) {
  const pageText = normalize(page.text);
  const code = normalize(row.codice);
  const nome = normalize(row.nome);
  let score = 0;

  if (code && pageText.includes(code)) score += 100;
  if (nome && nome.length > 2 && pageText.includes(nome)) score += 70;

  const nameParts = nome.split(' ').filter(p => p.length > 2);
  for (const part of nameParts) {
    if (pageText.includes(part)) score += 8;
  }

  if (/131\d+|139\d+/.test(code) && pageText.includes(code.replace(/[^0-9]/g, ''))) score += 20;
  return score;
}

function matchRows(rows, pages) {
  return rows.map(row => {
    let best = null;
    let bestScore = -1;
    for (const page of pages) {
      const score = scoreMatch(row, page);
      if (score > bestScore) {
        bestScore = score;
        best = page;
      }
    }

    let matchType = 'N/D';
    if (bestScore >= 100) matchType = 'Exact';
    else if (bestScore >= 30) matchType = 'Approx';

    return {
      ...row,
      fonteVisiva: best ? `${best.source} p.${best.pageNumber}` : '',
      matchType,
      matchedPage: best,
      description: row.nome ? `Prodotto ${row.nome}` : `Prodotto ${row.codice}`
    };
  });
}

function renderPreview(rows) {
  previewBody.innerHTML = '';
  rows.slice(0, 200).forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.codice}</td>
      <td>${row.nome}</td>
      <td>${row.prezzo}</td>
      <td>${row.categoria}</td>
      <td>${row.fonteVisiva || 'N/D'}</td>
      <td><span class="badge ${row.matchType === 'Exact' ? 'exact' : row.matchType === 'Approx' ? 'approx' : 'none'}">${row.matchType}</span></td>
    `;
    previewBody.appendChild(tr);
  });
}

function groupByCategory(rows) {
  const map = new Map();
  rows.forEach(r => {
    const key = r.categoria || 'Senza categoria';
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(r);
  });
  return [...map.entries()];
}

async function generatePdf(rows) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 12;
  let y = 18;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text('Listino Figurativo', margin, y);
  y += 8;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.text(`Generato il ${new Date().toLocaleString('it-IT')}`, margin, y);
  y += 10;

  const grouped = groupByCategory(rows);

  for (const [category, items] of grouped) {
    if (y > pageH - 20) { doc.addPage(); y = 18; }
    doc.setFillColor(185, 28, 28);
    doc.rect(margin, y - 5, pageW - margin * 2, 8, 'F');
    doc.setTextColor(255,255,255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.text(category, margin + 3, y);
    doc.setTextColor(0,0,0);
    y += 8;

    for (const item of items) {
      const cardHeight = item.matchedPage && item.matchType !== 'N/D' ? 58 : 26;
      if (y + cardHeight > pageH - 10) { doc.addPage(); y = 18; }

      doc.setDrawColor(220,220,220);
      doc.roundedRect(margin, y, pageW - margin * 2, cardHeight, 3, 3);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.text(`${item.codice} - ${item.nome}`.slice(0, 95), margin + 4, y + 7);

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      const leftX = margin + 4;
      let ty = y + 13;
      doc.text(`Prezzo: ${item.prezzo || 'N/D'}`, leftX, ty);
      ty += 5;
      doc.text(`Categoria: ${item.categoria || 'N/D'}`, leftX, ty);
      ty += 5;
      doc.text(`Match: ${item.matchType}`, leftX, ty);
      ty += 5;
      const fonte = `Fonte: ${item.fonteVisiva || 'N/D'}`;
      const splitFonte = doc.splitTextToSize(fonte, 90);
      doc.text(splitFonte, leftX, ty);

      if (item.matchedPage && item.matchType !== 'N/D') {
        try {
          doc.addImage(item.matchedPage.image, 'JPEG', pageW - 74, y + 4, 58, 46);
        } catch (e) {
          // ignore image add failure
        }
      }
      y += cardHeight + 6;
    }
  }

  doc.save('listino_figurativo_generato.pdf');
}

analyzeBtn.addEventListener('click', async () => {
  const excelFile = excelInput.files[0];
  const pdf1 = pdf1Input.files[0];
  const pdf2 = pdf2Input.files[0];

  if (!excelFile || !pdf1 || !pdf2) {
    setStatus('Carica tutti e 3 i file prima di procedere.');
    return;
  }

  try {
    analyzeBtn.disabled = true;
    generateBtn.disabled = true;
    downloadJsonBtn.disabled = true;
    setStatus('Lettura Excel...');
    excelRows = await readExcel(excelFile);

    setStatus(`Excel letto: ${excelRows.length} righe.\nLettura PDF 1...`);
    const pages1 = await readPdf(pdf1, 'PDF 1');

    setStatus(`Excel letto: ${excelRows.length} righe.\nPDF 1 letto: ${pages1.length} pagine.\nLettura PDF 2...`);
    const pages2 = await readPdf(pdf2, 'PDF 2');

    pdfPages = [...pages1, ...pages2];
    matchedRows = matchRows(excelRows, pdfPages);
    renderPreview(matchedRows);

    const exact = matchedRows.filter(r => r.matchType === 'Exact').length;
    const approx = matchedRows.filter(r => r.matchType === 'Approx').length;
    const none = matchedRows.filter(r => r.matchType === 'N/D').length;

    setStatus(
      `Analisi completata.\n` +
      `Righe Excel: ${matchedRows.length}\n` +
      `Pagine PDF analizzate: ${pdfPages.length}\n` +
      `Match Exact: ${exact}\n` +
      `Match Approx: ${approx}\n` +
      `Senza match: ${none}`
    );

    generateBtn.disabled = false;
    downloadJsonBtn.disabled = false;
  } catch (err) {
    console.error(err);
    setStatus(`Errore: ${err.message}`);
  } finally {
    analyzeBtn.disabled = false;
  }
});

generateBtn.addEventListener('click', async () => {
  if (!matchedRows.length) {
    setStatus('Prima esegui l\'analisi.');
    return;
  }
  setStatus('Generazione PDF in corso...');
  try {
    await generatePdf(matchedRows);
    setStatus('PDF generato con successo.');
  } catch (err) {
    console.error(err);
    setStatus(`Errore durante la generazione PDF: ${err.message}`);
  }
});

downloadJsonBtn.addEventListener('click', () => {
  if (!matchedRows.length) return;
  const clean = matchedRows.map(r => ({
    codice: r.codice,
    nome: r.nome,
    prezzo: r.prezzo,
    categoria: r.categoria,
    fonteVisiva: r.fonteVisiva,
    matchType: r.matchType
  }));
  const blob = new Blob([JSON.stringify(clean, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'match_listino.json';
  a.click();
  URL.revokeObjectURL(url);
});
