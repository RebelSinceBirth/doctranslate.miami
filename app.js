// ── Translation counter (countapi.xyz — free, no backend needed) ──
const COUNTER_NS  = 'doctranslate-miami';
const COUNTER_KEY = 'translations';

async function fetchCounter() {
  try {
    const res  = await fetch(`https://api.countapi.xyz/get/${COUNTER_NS}/${COUNTER_KEY}`);
    const data = await res.json();
    if (data && typeof data.value === 'number') showCounter(data.value);
  } catch (_) { /* silently skip if API unavailable */ }
}

async function incrementCounter() {
  try {
    const res  = await fetch(`https://api.countapi.xyz/hit/${COUNTER_NS}/${COUNTER_KEY}`);
    const data = await res.json();
    if (data && typeof data.value === 'number') showCounter(data.value);
  } catch (_) { /* silently skip */ }
}

function showCounter(value) {
  const wrap = document.getElementById('counterWrap');
  const num  = document.getElementById('counterNum');
  if (!wrap || !num) return;
  num.textContent = value.toLocaleString(); // formats 1234 → 1,234
  wrap.style.display = 'flex';
}

// Load count on page load
fetchCounter();

// ── Font toggle ──
let proMode = false;
function toggleFont() {
  proMode = !proMode;
  document.body.classList.toggle('pro-mode', proMode);
  document.getElementById('fontToggleIcon').textContent  = proMode ? '✦' : '🎨';
  document.getElementById('fontToggleLabel').textContent = proMode ? 'REPS Style' : 'Pro Style';
}

// ── PDF.js worker init ──
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// ── State ──
let stream = null;
let capturedDataUrl = null;
let capturedTextDirect = null; // set when DOCX is loaded — bypasses OCR
let facingMode = 'environment';
let orientMode = 'landscape'; // 'landscape' | 'portrait'
const CHUNK_SIZE = 4500; // Google Translate handles up to 5000 chars per request

// ── Orientation toggle ──
function toggleOrientation() {
  orientMode = orientMode === 'landscape' ? 'portrait' : 'landscape';
  const btn = document.getElementById('orientBtn');
  const wrap = document.getElementById('cameraWrap');
  if (orientMode === 'portrait') {
    btn.textContent = '📐 Portrait';
    btn.classList.add('portrait-active');
    wrap.style.aspectRatio = '9/16';
    wrap.style.maxHeight = '70vh';
  } else {
    btn.textContent = '📐 Landscape';
    btn.classList.remove('portrait-active');
    wrap.style.aspectRatio = '16/9';
    wrap.style.maxHeight = '';
  }
  // Restart camera if running to apply new constraints
  if (stream) { stopCamera(); startCamera(); }
}

// ── Tab switching ──
function switchTab(tab) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + tab).classList.add('active');
  document.getElementById('tab-' + tab).classList.add('active');
  if (tab !== 'camera') stopCamera();
}

// ── Camera ──
async function startCamera() {
  hideError();
  try {
    const isPortrait = orientMode === 'portrait';
    stream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode,
        width:       { ideal: 1920, min: 1280 }, // width drives sharpness — no height so camera won't zoom/crop
        aspectRatio: { ideal: isPortrait ? 9/16 : 16/9 },
      },
      audio: false
    });
    const video = document.getElementById('videoEl');
    video.srcObject = stream;
    video.style.display = 'block';
    document.getElementById('camPlaceholder').style.display = 'none';
    document.getElementById('startCamBtn').style.display = 'none';
    document.getElementById('snapBtn').disabled = false;
    document.getElementById('stopCamBtn').style.display = 'inline-flex';
    document.getElementById('flipBtn').style.display = 'inline-flex';
  } catch (e) {
    showError('Camera access denied or unavailable. Please allow camera access and try again.');
  }
}

function stopCamera() {
  if (stream) { stream.getTracks().forEach(t => t.stop()); stream = null; }
  const video = document.getElementById('videoEl');
  video.style.display = 'none';
  video.srcObject = null;
  document.getElementById('camPlaceholder').style.display = 'block';
  document.getElementById('startCamBtn').style.display = 'inline-flex';
  document.getElementById('snapBtn').disabled = true;
  document.getElementById('stopCamBtn').style.display = 'none';
  document.getElementById('flipBtn').style.display = 'none';
}

async function flipCamera() {
  facingMode = facingMode === 'environment' ? 'user' : 'environment';
  stopCamera();
  await startCamera();
}

function capturePhoto() {
  const video = document.getElementById('videoEl');
  const canvas = document.getElementById('captureCanvas');
  canvas.width  = video.videoWidth  || 1280;
  canvas.height = video.videoHeight || 720;
  canvas.getContext('2d').drawImage(video, 0, 0);
  capturedDataUrl = preprocessImage(canvas);
  showPreview(capturedDataUrl, '✓ Photo Captured — Ready to Process');
  document.getElementById('processBtn').disabled = false;
}

// ── Upload ──
function onFileSelected(e) { const f = e.target.files[0]; if (f) loadFile(f); }
function onDragOver(e)  { e.preventDefault(); document.getElementById('uploadArea').classList.add('dragover'); }
function onDragLeave()  { document.getElementById('uploadArea').classList.remove('dragover'); }
function onDrop(e)      { e.preventDefault(); onDragLeave(); const f = e.dataTransfer.files[0]; if (f) loadFile(f); }

function loadFile(file) {
  capturedDataUrl = null;
  capturedTextDirect = null;
  hideError();

  const isPDF  = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
  const isDOCX = file.name.toLowerCase().endsWith('.docx') ||
                 file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  if (isPDF)       { loadPDF(file);  return; }
  if (isDOCX)      { loadDOCX(file); return; }
  if (!file.type.startsWith('image/')) { showError('Unsupported file type. Use an image, PDF, or DOCX.'); return; }

  // Image path
  const reader = new FileReader();
  reader.onload = ev => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.getElementById('captureCanvas');
      canvas.width  = img.naturalWidth;
      canvas.height = img.naturalHeight;
      canvas.getContext('2d').drawImage(img, 0, 0);
      capturedDataUrl = preprocessImage(canvas);
      showPreview(capturedDataUrl, `✓ Loaded: ${file.name}`);
      document.getElementById('processBtn').disabled = false;
    };
    img.src = ev.target.result;
  };
  reader.readAsDataURL(file);
}

// ── PDF loader — renders all pages to a single stitched canvas ──
async function loadPDF(file) {
  showProgress('Loading PDF…', 5);
  try {
    if (typeof pdfjsLib === 'undefined') throw new Error('PDF engine not loaded yet — try again in a moment.');
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const ab  = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
    const totalPages = pdf.numPages;
    const renderPages = Math.min(totalPages, 8); // cap at 8 pages

    const rendered = [];
    for (let p = 1; p <= renderPages; p++) {
      updateProgress(`Rendering page ${p} of ${renderPages}…`, 5 + Math.round((p / renderPages) * 35));
      const page = await pdf.getPage(p);
      const vp   = page.getViewport({ scale: 1.8 }); // 1.8x = good OCR quality without huge memory
      const c    = document.createElement('canvas');
      c.width    = vp.width;
      c.height   = vp.height;
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      rendered.push(c);
    }

    // Stitch pages vertically
    const totalW = Math.max(...rendered.map(c => c.width));
    const totalH = rendered.reduce((h, c) => h + c.height + 20, 0); // 20px gap between pages
    const combined = document.getElementById('captureCanvas');
    combined.width  = totalW;
    combined.height = totalH;
    const ctx = combined.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, totalW, totalH);
    let y = 0;
    for (const c of rendered) { ctx.drawImage(c, 0, y); y += c.height + 20; }

    capturedDataUrl = preprocessImage(combined);
    const pageNote = totalPages > renderPages ? ` (first ${renderPages} of ${totalPages} pages)` : ` · ${totalPages} page${totalPages > 1 ? 's' : ''}`;
    showPreview(capturedDataUrl, `✓ PDF Loaded: ${file.name}${pageNote}`);
    document.getElementById('processBtn').disabled = false;
    hideProgress();
  } catch (e) {
    hideProgress();
    showError('Could not load PDF: ' + e.message);
  }
}

// ── DOCX loader — uses HTML output to preserve headings, lists, tables ──
async function loadDOCX(file) {
  showProgress('Reading Word document…', 10);
  try {
    if (typeof mammoth === 'undefined') throw new Error('DOCX engine not loaded yet — try again.');
    const ab     = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer: ab });
    capturedTextDirect = htmlToStructuredText(result.value).trim();
    if (!capturedTextDirect) throw new Error('No text found in document.');

    document.getElementById('previewImg').src = '';
    document.getElementById('previewImg').style.display = 'none';
    document.getElementById('previewLabel').textContent = `✓ Word Doc Loaded: ${file.name} · Structure preserved — Ready`;
    document.getElementById('previewWrap').style.display = 'block';
    document.getElementById('processBtn').disabled = false;
    hideProgress();
  } catch (e) {
    hideProgress();
    showError('Could not read Word document: ' + e.message);
  }
}

// ── HTML → structured plain text (preserves headings, lists, tables, spacing) ──
function htmlToStructuredText(html) {
  const div = document.createElement('div');
  div.innerHTML = html;

  function walk(node) {
    if (node.nodeType === 3) return node.textContent; // raw text node
    const tag = (node.tagName || '').toLowerCase();
    const kids = [...node.childNodes].map(walk).join('');

    switch(tag) {
      case 'h1': return '\n\n' + kids.toUpperCase() + '\n' + '='.repeat(Math.min(kids.length, 40)) + '\n';
      case 'h2': return '\n\n' + kids.toUpperCase() + '\n' + '-'.repeat(Math.min(kids.length, 40)) + '\n';
      case 'h3': case 'h4': case 'h5': case 'h6':
                  return '\n\n' + kids.toUpperCase() + '\n';
      case 'p':   return kids.trim() ? kids.trim() + '\n\n' : '';
      case 'br':  return '\n';
      case 'li':  return '  • ' + kids.trim() + '\n';
      case 'ul': case 'ol': return '\n' + kids + '\n';
      case 'table': return '\n' + kids + '\n';
      case 'tr':  return kids.trimEnd() + '\n';
      case 'td': case 'th': return kids.trim() + '\t';
      case 'strong': case 'b': return kids.toUpperCase();
      case 'em': case 'i': return kids;
      default:    return kids;
    }
  }

  return walk(div)
    .replace(/\t+$/gm, '')         // trim trailing tabs
    .replace(/\n{3,}/g, '\n\n')    // max 2 consecutive blank lines
    .trim();
}

// ── OCR post-processing — fix common misreads ──
function cleanOCRText(text) {
  const lines = text.split('\n').map(line => {

    // 1. Zero-gram artifacts: [o]], [o], [oJ], [oJ]], o], [0] → 0g
    line = line.replace(/\[o[Jj]?\]?\]?/g, '0g');
    line = line.replace(/\bo\]\]?(?=\s|$)/g, '0g');
    line = line.replace(/\[0\](?=\s|$)/g, '0g');
    line = line.replace(/\[O[Jj]?\]?\]?/g, '0g'); // uppercase O variant

    // 2. Cent symbol in gram measurements: 5¢g → 5g, 5¢ → 5g
    line = line.replace(/(\d+)¢g?\b/g, '$1g');

    // 3. Lowercase-L misread as 1 before digits: l4g → 14g
    line = line.replace(/\bl(\d)/g, '1$1');

    // 4. Common OCR word errors
    line = line.replace(/\bIltem\b/gi, 'Item');
    line = line.replace(/\bItern\b/gi, 'Item');
    line = line.replace(/\bcallfel\b/gi, 'cal');   // "callfel" → "cal"
    line = line.replace(/\bcallfei\b/gi, 'cal');   // alternate garble
    line = line.replace(/«\s*/g, '');              // stray « guillemet noise

    // 5. Strip leading bracket+digit section header artifacts: "[1 CROISSANT" → "CROISSANT"
    line = line.replace(/^\s*\[?\d\s+(?=[A-Z]{2})/g, '');

    // 6. Context-aware g→9 fix for macro/nutrition table rows
    // If a line has 3+ numeric tokens, treat it as a data row and fix trailing 9 → g
    const tokens = line.trim().split(/\s+/);
    const numericLike = tokens.filter(t => /^\d{1,4}[g9]?$/.test(t));
    if (tokens.length >= 3 && numericLike.length >= 3) {
      line = line.replace(/\b(\d{1,4})9\b/g, '$1g');        // 9 at end → g (no existing g)
      line = line.replace(/\b(\d{1,2})9g\b/g, '$1g');       // X9g → Xg  (e.g. 69g→6g, 59g→5g)
      line = line.replace(/\b(\d{2,3})9g\b/g, '$1g');       // XX9g → XXg (e.g. 129g→12g)
    }

    // 7. Dollar sign / price fixes
    line = line.replace(/\bS(\d{1,4}(?:[.,]\d{2})?)\b/g, '\$$1');   // S24.99 → $24.99, S100 → $100
    line = line.replace(/\b5(\d{2,3}(?:\.\d{2})?)\b(?=\s|$)/g, (m, n) => {
      // Only treat leading 5 as $ when it looks like a price (e.g. 524.99 → $24.99)
      return '\$' + n;
    });
    line = line.replace(/\$\s+(\d)/g, '\$$1');                        // $ 24 → $24 (space after $)
    line = line.replace(/(\$\d+)\s(\d{2})(?=\s|$)/g, '$1.$2');       // $24 99 → $24.99
    line = line.replace(/(\d{1,3})\s(\d{3})(?=\s|$)/g, '$1$2');      // 1 000 → 1000 (spaced thousands)

    // 8. Letter O / zero confusion inside numbers
    line = line.replace(/(\d)[Oo](\d)/g, (_, a, b) => a + '0' + b);  // 1O0 → 100, 2o5 → 205

    return line;
  });

  // 9. Sequential number gap repair — fixes "Day 1" when it should be "Day 11"
  // Detects numbered list/table rows and repairs gaps in the sequence
  const numberedRows = [];
  lines.forEach((line, i) => {
    const m = line.trimStart().match(/^(\d{1,3})(\s+\S)/);
    if (m) numberedRows.push({ i, num: parseInt(m[1]) });
  });
  for (let r = 1; r < numberedRows.length - 1; r++) {
    const prev = numberedRows[r - 1].num;
    const curr = numberedRows[r];
    const next = numberedRows[r + 1] ? numberedRows[r + 1].num : null;
    // If sequence goes 10 → 1 → 12, the 1 should be 11
    if (next !== null && curr.num < prev && next === prev + 2) {
      lines[curr.i] = lines[curr.i].replace(/^(\s*)(\d+)/, (_, sp, n) => sp + String(prev + 1));
    }
  }

  // Remove obvious decorative/noise-only lines
  return lines
    .filter(line => {
      const t = line.trim();
      if (!t) return true; // keep blank lines for spacing
      if (/^[-=_—~•·\s]{3,}$/.test(t)) return false;  // pure dashes/symbols
      if (/^[A-Za-z]{0,2}[\s]*[-—=_]{3,}/.test(t)) return false; // "L ———" style noise
      if (t.length <= 2 && !/^\d+$/.test(t)) return false; // very short noise
      // Remove lines that are >60% non-alphanumeric (decorative artifacts)
      const alphaNum = (t.match(/[a-zA-Z0-9]/g) || []).length;
      if (t.length > 4 && alphaNum / t.length < 0.3) return false;
      return true;
    })
    .join('\n')
    .replace(/  +/g, ' ')
    .trim();
}

// ════════════════════════════════════════════════════════
//  OCR CLEANUP PIPELINE
//  Pass 1: cleanOCRText   — fix character-level errors
//  Pass 2: reformatOCRText — rebuild professional structure
// ════════════════════════════════════════════════════════

function processOCRText(raw) {
  return reformatOCRText(cleanOCRText(raw));
}

// ── Pass 2: structural formatter ──
function reformatOCRText(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  if (!lines.length) return text;

  const out = [];
  let prevType = null;

  for (let i = 0; i < lines.length; i++) {
    const line  = lines[i];
    const next  = lines[i + 1] || '';
    const type  = classifyOCRLine(line);

    // ── Blank line before headers (unless document start) ──
    if (type === 'header' && out.length > 0 && out[out.length - 1] !== '') {
      out.push('');
    }

    if (type === 'header') {
      out.push(line);
      // Underline: === for first/top header, --- for sub-headers
      if (!next.startsWith('-') && !next.startsWith('=')) {
        out.push((prevType === null ? '=' : '-').repeat(Math.max(line.length, 20)));
      }
      out.push('');

    } else if (type === 'table-header') {
      if (out.length > 0 && out[out.length - 1] !== '') out.push('');
      out.push(line);
      out.push('-'.repeat(Math.max(line.length + 4, 30)));

    } else if (type === 'table-row') {
      out.push(line);

    } else if (type === 'total') {
      // Separator before TOTAL
      const last = out[out.length - 1] || '';
      if (last !== '' && !last.startsWith('-') && !last.startsWith('=')) {
        out.push('-'.repeat(30));
      }
      out.push(line);
      out.push('');

    } else if (type === 'list') {
      out.push(line);

    } else {
      // Body text — join wrapped lines when a sentence looks split
      if (prevType === 'body' && out.length > 0) {
        const last = out[out.length - 1];
        const startsLower  = /^[a-z]/.test(line);
        const prevNoEnd    = !/[.!?:;]$/.test(last);
        const lineIsShort  = line.length < 60;
        if (last && prevNoEnd && (startsLower || lineIsShort)) {
          out[out.length - 1] = last + ' ' + line;
          continue;
        }
      }
      out.push(line);
    }

    prevType = type;
  }

  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

// ── Line classifier ──
function classifyOCRLine(line) {
  const words    = line.trim().split(/\s+/);
  const nums     = words.filter(w => /^\d+[a-z%]?$/i.test(w)).length;
  const letters  = (line.match(/[a-zA-Z]/g) || []).length;
  const caps     = (line.match(/[A-Z]/g)    || []).length;
  const capsRatio = letters > 0 ? caps / letters : 0;

  // TOTAL rows always get a separator
  if (/^\s*TOTAL\b/i.test(line)) return 'total';

  // Header: short line, mostly uppercase, few/no numbers
  if (line.length <= 55 && capsRatio >= 0.65 && nums <= 1 && words.length <= 9) {
    return 'header';
  }

  // Table column header row — lots of short alphabetic tokens, no numbers
  const shortAlpha = words.filter(w => /^[A-Za-z]{1,10}$/.test(w) && w.length <= 10).length;
  if (nums === 0 && shortAlpha >= 3 && words.length <= 10 && line.length <= 60) {
    return 'table-header';
  }

  // Table data row: a text label followed by 2+ numeric values
  if (nums >= 2 && words.length >= 3) return 'table-row';

  // List item
  if (/^[•\-\*·]/.test(line) || /^\d+[\.\)]/.test(line)) return 'list';

  return 'body';
}

// ── Image preprocessing — grayscale + contrast boost for better OCR ──
function preprocessImage(canvas) {
  const ctx = canvas.getContext('2d');
  const w = canvas.width, h = canvas.height;
  const imgData = ctx.getImageData(0, 0, w, h);
  const d = imgData.data;

  for (let i = 0; i < d.length; i += 4) {
    // Convert to grayscale using luminance weights
    const gray = 0.299 * d[i] + 0.587 * d[i+1] + 0.114 * d[i+2];
    // Contrast boost: stretch towards black/white
    const contrast = 1.5;
    const boosted = Math.min(255, Math.max(0, contrast * (gray - 128) + 128));
    d[i] = d[i+1] = d[i+2] = boosted;
  }
  ctx.putImageData(imgData, 0, 0);
  return canvas.toDataURL('image/png');
}

// ── Preview ──
function showPreview(dataUrl, label) {
  document.getElementById('previewImg').src = dataUrl;
  document.getElementById('previewLabel').textContent = label;
  document.getElementById('previewWrap').style.display = 'block';
}

// ── Main: OCR + Translate ──
async function processDocument() {
  if (!capturedDataUrl && !capturedTextDirect) { showError('Please capture or upload a file first.'); return; }
  hideError();
  document.getElementById('processBtn').disabled = true;
  document.getElementById('resultsWrap').style.display = 'none';

  // DOCX path — skip OCR entirely
  if (capturedTextDirect) {
    showProgress('Translating…', 75);
    try {
      const extractedText = processOCRText(capturedTextDirect);
      document.getElementById('ocrText').value = extractedText;
      const langCode = document.getElementById('langSelect').value;
      const langName = document.getElementById('langSelect').selectedOptions[0].text.replace(/^\S+\s/, '');
      const translated = await translateText(extractedText, langCode);
      updateProgress('Complete!', 100);
      document.getElementById('translatedText').value = translated;
      document.getElementById('translatedLabel').textContent = `Translated → ${langName}`;
      document.getElementById('resultsWrap').style.display = 'block';
      document.getElementById('resultsWrap').scrollIntoView({ behavior: 'smooth' });
      setTimeout(hideProgress, 1000);
      incrementCounter(); // ✦ count this translation
    } catch (err) {
      hideProgress();
      showError('Translation failed: ' + err.message);
    } finally {
      document.getElementById('processBtn').disabled = false;
    }
    return;
  }

  showProgress('Initializing OCR Engine…', 0);

  try {
    const result = await Tesseract.recognize(capturedDataUrl, 'eng', {
      logger: m => {
        if (m.status === 'recognizing text')                  updateProgress(`Recognizing Text… ${Math.round(m.progress*70)}%`, Math.round(m.progress*70));
        else if (m.status === 'loading tesseract core')       updateProgress('Loading OCR Engine…', 5);
        else if (m.status === 'initializing tesseract')       updateProgress('Initializing…', 10);
        else if (m.status === 'loading language traineddata') updateProgress('Loading Language Data…', 20);
      },
    });
    const rawText = result.data.text.trim();
    if (!rawText) {
      hideProgress();
      showError('No text detected. Try a clearer photo with better lighting.');
      document.getElementById('processBtn').disabled = false;
      return;
    }
    const extractedText = processOCRText(rawText);
    document.getElementById('ocrText').value = extractedText;
    updateProgress('Translating…', 75);

    const langCode = document.getElementById('langSelect').value;
    const langName = document.getElementById('langSelect').selectedOptions[0].text.replace(/^\S+\s/, '');
    const translated = await translateText(extractedText, langCode);
    updateProgress('Complete!', 100);

    document.getElementById('translatedText').value = translated;
    document.getElementById('translatedLabel').textContent = `Translated → ${langName}`;
    document.getElementById('resultsWrap').style.display = 'block';
    document.getElementById('resultsWrap').scrollIntoView({ behavior: 'smooth' });
    setTimeout(hideProgress, 1000);
    incrementCounter(); // ✦ count this translation
  } catch (err) {
    hideProgress();
    showError('Something went wrong: ' + err.message);
  } finally {
    document.getElementById('processBtn').disabled = false;
  }
}

// ── Language code normalizer for Google Translate ──
const LANG_CODE_MAP = {
  'es-419': 'es',
  'es-MX':  'es',
  'es-CO':  'es',
  'es-AR':  'es',
};
function apiLangCode(code) {
  return LANG_CODE_MAP[code] ?? code;
}

// ── Translation — Google Translate (free, no API key) ──
async function translateText(text, targetLang) {
  const code = apiLangCode(targetLang);
  const chunks = chunkText(text, CHUNK_SIZE);
  const results = [];
  for (let i = 0; i < chunks.length; i++) {
    const label = chunks.length > 1 ? ` Part ${i+1}/${chunks.length}` : '';
    updateProgress(`Translating${label}…`, 75 + Math.round((i / chunks.length) * 22));
    results.push(await translateChunk(chunks[i], 'en', code));
  }
  // Join with newline so paragraph/line structure is preserved
  return results.join('\n');
}

async function translateChunk(text, from, to) {
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${from}&tl=${to}&dt=t&q=${encodeURIComponent(text)}`;
  let res;
  try {
    res = await fetch(url);
  } catch (netErr) {
    throw new Error('Network error reaching translation service. Check your connection.');
  }
  if (!res.ok) throw new Error(`Translation service returned ${res.status}`);

  const data = await res.json();

  // Google response: [ [ ["translated seg","original seg",...], ... ], null, "detected_lang", ... ]
  if (data && Array.isArray(data[0])) {
    const translated = data[0]
      .filter(seg => Array.isArray(seg) && typeof seg[0] === 'string')
      .map(seg => seg[0])
      .join('');
    if (translated.trim()) return translated;
  }
  throw new Error('Empty response from translation service — try again.');
}

function chunkText(text, size) {
  // Split at paragraph boundaries first to preserve document structure
  const blocks = text.split(/\n\s*\n/);
  const chunks = [];
  let current = '';

  for (const block of blocks) {
    const candidate = current ? current + '\n\n' + block : block;
    if (candidate.length <= size) {
      current = candidate;
    } else {
      if (current) chunks.push(current);
      if (block.length <= size) {
        current = block;
      } else {
        // Block too large — split at line breaks
        const lines = block.split('\n');
        current = '';
        for (const line of lines) {
          const lc = current ? current + '\n' + line : line;
          if (lc.length <= size) { current = lc; }
          else { if (current) chunks.push(current); current = line; }
        }
      }
    }
  }
  if (current) chunks.push(current);
  return chunks.length ? chunks : [text.slice(0, size)];
}

// ── Progress ──
function showProgress(msg, pct) { document.getElementById('progressWrap').style.display = 'block'; updateProgress(msg, pct); }
function updateProgress(msg, pct) {
  document.getElementById('progressStatus').textContent = msg;
  document.getElementById('progressPct').textContent = pct + '%';
  document.getElementById('progressFill').style.width = pct + '%';
}
function hideProgress() { document.getElementById('progressWrap').style.display = 'none'; }

// ── Error ──
function showError(msg) { const e = document.getElementById('errorMsg'); e.textContent = '⚠ ' + msg; e.style.display = 'block'; }
function hideError()    { document.getElementById('errorMsg').style.display = 'none'; }

// ── Copy ──
async function copyText(id, btn) {
  try {
    await navigator.clipboard.writeText(document.getElementById(id).value);
    btn.textContent = '✓ Copied'; btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 2000);
  } catch { document.getElementById(id).select(); document.execCommand('copy'); }
}

// ── Helpers ──
function getResults() {
  return {
    orig:  document.getElementById('ocrText').value,
    trans: document.getElementById('translatedText').value,
    lang:  document.getElementById('langSelect').selectedOptions[0].text.replace(/^\S+\s/, ''),
  };
}

// ── Download TXT ──
function downloadTXT() {
  const { orig, trans, lang } = getResults();
  const content = `DocTranslate - OG Bar Beast\n${'='.repeat(40)}\n\nORIGINAL TEXT:\n${orig}\n\n${'='.repeat(40)}\n\nTRANSLATED (${lang}):\n${trans}\n\n${'='.repeat(40)}\nrepsapp.ai - contact@ogbarbeast.com`;
  const a = document.createElement('a');
  const bom = '﻿';
  a.href = URL.createObjectURL(new Blob([bom + content], { type: 'text/plain;charset=utf-8' }));
  a.download = 'doctranslate_result.txt';
  a.click(); URL.revokeObjectURL(a.href);
}

// ── Download PDF ──
function downloadPDF() {
  if (typeof window.jspdf === 'undefined') { alert('PDF library still loading — try again in a moment.'); return; }
  const { jsPDF } = window.jspdf;
  const { orig, trans, lang } = getResults();
  const doc   = new jsPDF({ unit: 'mm', format: 'a4' });
  const margin    = 18;
  const pageW     = doc.internal.pageSize.getWidth();
  const pageH     = doc.internal.pageSize.getHeight();
  const usableW   = pageW - margin * 2;
  let y = margin;

  function writeLine(text, size, bold) {
    doc.setFontSize(size);
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    const lines = doc.splitTextToSize(text, usableW);
    lines.forEach(line => {
      if (y + size * 0.45 > pageH - margin) { doc.addPage(); y = margin; }
      doc.text(line, margin, y);
      y += size * 0.45;
    });
    y += 3;
  }

  // Header
  doc.setFillColor(8, 8, 15);
  doc.rect(0, 0, pageW, 22, 'F');
  doc.setTextColor(255, 255, 255);
  writeLine('DOCTRANSLATE  ·  OG BAR BEAST  ·  MIAMI', 13, true);
  doc.setTextColor(0, 212, 255);
  writeLine('doctranslate.miami', 9, false);
  doc.setTextColor(30, 30, 30);
  y += 4;

  writeLine('ORIGINAL TEXT', 12, true);
  doc.setDrawColor(0, 212, 255); doc.setLineWidth(0.4);
  doc.line(margin, y, pageW - margin, y); y += 5;
  writeLine(orig, 10, false);
  y += 6;

  writeLine(`TRANSLATED — ${lang.toUpperCase()}`, 12, true);
  doc.setDrawColor(255, 26, 109); doc.line(margin, y, pageW - margin, y); y += 5;
  writeLine(trans, 10, false);

  // Footer on last page
  y = pageH - 10;
  doc.setFontSize(7); doc.setTextColor(120, 120, 120);
  doc.text('Generated by DocTranslate · repsapp.ai · contact@ogbarbeast.com', margin, y);

  doc.save('doctranslate_result.pdf');
}

// ── Download DOCX ──
async function downloadDOCX() {
  if (typeof docx === 'undefined') { alert('Word library still loading — try again in a moment.'); return; }
  const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } = docx;
  const { orig, trans, lang } = getResults();

  const makeParas = (text) => text.split('\n').map(line =>
    new Paragraph({ children: [new TextRun({ text: line || ' ', size: 22 })] })
  );

  const wordDoc = new Document({
    sections: [{
      children: [
        new Paragraph({
          children: [new TextRun({ text: 'DocTranslate — OG Bar Beast', bold: true, size: 32, color: 'FF1A6D' })],
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
          children: [new TextRun({ text: 'doctranslate.miami  ·  contact@ogbarbeast.com', size: 18, color: '888888' })],
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ children: [new TextRun('')] }),
        new Paragraph({ children: [new TextRun({ text: 'ORIGINAL TEXT', bold: true, size: 26, color: '00B4D8' })] }),
        new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '00B4D8' } }, children: [] }),
        ...makeParas(orig),
        new Paragraph({ children: [new TextRun('')] }),
        new Paragraph({ children: [new TextRun({ text: `TRANSLATED — ${lang.toUpperCase()}`, bold: true, size: 26, color: 'FF1A6D' })] }),
        new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'FF1A6D' } }, children: [] }),
        ...makeParas(trans),
      ],
    }],
  });

  const blob = await Packer.toBlob(wordDoc);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'doctranslate_result.docx';
  a.click(); URL.revokeObjectURL(a.href);
}

// ── Reset ──
function resetAll() {
  capturedDataUrl = null; capturedTextDirect = null; stopCamera();
  document.getElementById('previewImg').style.display = '';
  document.getElementById('previewWrap').style.display = 'none';
  document.getElementById('previewImg').src = '';
  document.getElementById('ocrText').value = '';
  document.getElementById('translatedText').value = '';
  document.getElementById('resultsWrap').style.display = 'none';
  document.getElementById('processBtn').disabled = true;
  document.getElementById('fileInput').value = '';
  hideError(); hideProgress();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
