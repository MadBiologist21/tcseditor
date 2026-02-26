/**
 * Glass Line Cutter - Precision text extraction from .docx
 *
 * Single source of truth: currentLineIndex
 * Pill position is ALWAYS derived from DOM position of active line
 * Scrolling NEVER changes the active line index
 * Text formatting (bold, italic, underline) is preserved
 *
 * PARSING APPROACH:
 * We parse word/document.xml using direct regex on the raw string — NOT DOMParser.
 * DOMParser is a browser XML parser and may silently strip leading/trailing whitespace
 * from <w:t> elements that lack xml:space="preserve", depending on the browser and
 * how the DOCX was authored. By reading the raw XML string ourselves we capture
 * every character exactly as stored, including all spaces between runs.
 */

let lines = [];           // Array of line objects: { html: string, text: string }
let currentLineIndex = 0;
let multiLineCount = 1;

// DOM references
const editor = document.getElementById("editor");
const glassColumn = document.getElementById("glassColumn");
const focusPill = document.getElementById("focusPill");
const cutBtn = document.getElementById("cutBtn");
const fileInput = document.getElementById("fileInput");
const status = document.getElementById("status");
const lineCounter = document.getElementById("lineCounter");

/**
 * Unescape XML entities to their real characters.
 * Required when reading text from raw XML string (regex capture groups give us
 * the raw entity text like "&amp;" — we must convert it back to "&" before use).
 */
function xmlUnescape(str) {
  return str
    .replace(/&amp;/g,  '&')
    .replace(/&lt;/g,   '<')
    .replace(/&gt;/g,   '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

/**
 * Escape a plain text string for safe injection into innerHTML.
 */
function htmlEscape(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Parse word/document.xml raw string into an array of line objects.
 * Each line: { text: string, html: string }
 *
 * WHY RAW STRING PARSING:
 * The browser's DOMParser XML parser normalises whitespace in text nodes.
 * Specifically, <w:t> elements without xml:space="preserve" can have their
 * leading/trailing spaces silently dropped. This causes "for her belief" to
 * render as "herbelief" because the space that belongs to the start of the
 * underlined run is stripped before JavaScript ever sees it.
 * Reading the raw XML string with regex captures the content between <w:t>
 * and </w:t> verbatim — spaces and all — bypassing the parser entirely.
 */
function parseDocxXml(rawXml) {
  const result = [];

  // ── Step 1: remove all w:del blocks (deleted tracked-change content) ────────
  // We do this on the full string before splitting into paragraphs so that
  // a <w:del> spanning a paragraph boundary is also handled correctly.
  const cleanXml = rawXml.replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, '');

  // ── Step 2: split into paragraphs ──────────────────────────────────────────
  // Match every <w:p ...>...</w:p> block.
  const paraRe = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let paraMatch;

  while ((paraMatch = paraRe.exec(cleanXml)) !== null) {
    const paraXml = paraMatch[0];
    let lineText = '';
    let lineHtml  = '';

    // ── Step 3: extract each run within the paragraph ────────────────────────
    const runRe = /<w:r[ >][\s\S]*?<\/w:r>/g;
    let runMatch;

    while ((runMatch = runRe.exec(paraXml)) !== null) {
      const runXml = runMatch[0];

      // ── Step 4: read run properties (formatting) ─────────────────────────
      const rprMatch = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(runXml);
      const rpr = rprMatch ? rprMatch[1] : '';

      // Bold: <w:b/> or <w:b> present, but NOT <w:bCs>, and NOT val="0"/"false"
      const boldTag = /<w:b(?!C)(?:\s[^>]*)?\/>|<w:b(?!C)>/.test(rpr);
      const boldOff = /<w:b(?!C)[^>]+w:val="(?:0|false)"/.test(rpr);
      const isBold  = boldTag && !boldOff;

      // Italic: same pattern for w:i / w:iCs
      const italicTag = /<w:i(?!C)(?:\s[^>]*)?\/>|<w:i(?!C)>/.test(rpr);
      const italicOff = /<w:i(?!C)[^>]+w:val="(?:0|false)"/.test(rpr);
      const isItalic  = italicTag && !italicOff;

      // Underline: w:u must be present with a val that is NOT "none"/"0"/"false"
      const uMatch     = /<w:u\s+w:val="([^"]*)"/.exec(rpr);
      const isUnderline = uMatch !== null &&
                          !['none', '0', 'false'].includes(uMatch[1]);

      // ── Step 5: extract text from all <w:t> elements in this run ─────────
      // We use a regex on the raw string so we capture the content exactly as
      // stored, including any leading/trailing spaces.
      // Then we XML-unescape (e.g. &amp; → &) before HTML-escaping.
      let runText = '';
      const wtRe = /<w:t(?:[^>]*)>([\s\S]*?)<\/w:t>/g;
      let wtMatch;
      while ((wtMatch = wtRe.exec(runXml)) !== null) {
        runText += xmlUnescape(wtMatch[1]);
      }

      // Handle tab elements
      const tabCount = (runXml.match(/<w:tab\/>/g) || []).length;
      runText += '\t'.repeat(tabCount);

      // Handle soft line breaks (w:br without type="page")
      let softBreakHtml = '';
      const brRe = /<w:br(?:\s+w:type="([^"]*)")?\/>/g;
      let brMatch;
      while ((brMatch = brRe.exec(runXml)) !== null) {
        const brType = brMatch[1];
        if (!brType || brType === 'textWrapping') {
          runText      += '\n';
          softBreakHtml = '<br>';
        }
      }

      if (!runText) continue;

      // ── Step 6: build HTML for this run ──────────────────────────────────
      // HTML-escape the text (it is now plain text after xmlUnescape above).
      let runHtml = htmlEscape(runText);

      // Apply formatting tags around the full run text.
      // Spaces inside inline tags (e.g. <u> belief</u>) render correctly in
      // all browsers, so we wrap the entire content including surrounding spaces.
      if (isUnderline) runHtml = `<u>${runHtml}</u>`;
      if (isItalic)    runHtml = `<em>${runHtml}</em>`;
      if (isBold)      runHtml = `<strong>${runHtml}</strong>`;

      if (softBreakHtml) runHtml += softBreakHtml;

      lineText += runText;
      lineHtml  += runHtml;
    }

    // Trim plain text at paragraph boundaries (removes leading/trailing newlines
    // that may have come from soft breaks at the very start/end of a paragraph).
    // We do NOT trim lineHtml because trimming the raw HTML string would eat
    // spaces that sit immediately before or after inline formatting tags.
    lineText = lineText.trim();
    if (lineText.length > 0) {
      result.push({ text: lineText, html: lineHtml });
    }
  }

  return result;
}

/**
 * Load and process DOCX file with formatting preserved
 */
fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  try {
    const buffer = await file.arrayBuffer();
    const zip    = await JSZip.loadAsync(buffer);

    // Get the raw XML string — do NOT pass through DOMParser
    const rawXml = await zip.file('word/document.xml').async('string');

    lines = parseDocxXml(rawXml);

    console.log("Parsed lines:", lines.slice(0, 5));

    currentLineIndex = 0;
    multiLineCount   = 1;

    render();
    updateLineCounter();

    requestAnimationFrame(() => { scrollToCurrentLine(); });

    showStatus(`Loaded ${lines.length} lines`);
  } catch (err) {
    showStatus("Error loading file");
    console.error(err);
  }
});

/**
 * Render all lines as DOM elements with formatting preserved
 */
function render() {
  if (lines.length === 0) {
    editor.innerHTML = `
      <div style="text-align:center;color:#888;padding:100px 0;">
        No content found in document
      </div>`;
    return;
  }

  editor.innerHTML = lines
    .map((line, i) => `<div class="line" data-index="${i}"><span>${line.html}</span></div>`)
    .join('');
}

/**
 * Position the glass column so the focus pill aligns with the current line.
 */
function positionPill() {
  if (lines.length === 0) return;

  const lineEl = editor.querySelector(`.line[data-index="${currentLineIndex}"]`);
  if (!lineEl) return;

  const frame     = document.getElementById("frame");
  const frameRect = frame.getBoundingClientRect();
  const lineRect  = lineEl.getBoundingClientRect();

  glassColumn.style.top = `${lineRect.top - frameRect.top}px`;

  const lineHeight    = lineEl.offsetHeight;
  focusPill.style.height  = `${lineHeight * multiLineCount}px`;
  glassColumn.style.height = `${lineHeight * (3 + Math.max(0, multiLineCount - 1))}px`;
}

/**
 * Scroll editor to centre the current line, then reposition the pill.
 */
function scrollToCurrentLine() {
  if (lines.length === 0) return;

  const lineEl = editor.querySelector(`.line[data-index="${currentLineIndex}"]`);
  if (!lineEl) return;

  glassColumn.style.transition = '';           // restore smooth transition for navigation
  lineEl.scrollIntoView({ block: 'center', behavior: 'instant' });

  // Double rAF: first frame commits the scroll, second reads settled layout
  requestAnimationFrame(() => { requestAnimationFrame(positionPill); });
}

// Reposition pill on scroll (no transition — avoids lag fighting rapid events)
editor.addEventListener("scroll", () => {
  glassColumn.style.transition = 'none';
  positionPill();
}, { passive: true });

/**
 * CUT — copy current line(s) to clipboard and advance
 */
cutBtn.addEventListener("click", async () => {
  if (lines.length === 0)            { showStatus("No document loaded"); return; }
  if (currentLineIndex >= lines.length) { showStatus("All lines processed!"); return; }

  const endIndex    = Math.min(currentLineIndex + multiLineCount, lines.length);
  const linesToCopy = lines.slice(currentLineIndex, endIndex);
  const textToCopy  = linesToCopy.map(l => l.text).join('\n');
  const htmlToCopy  = linesToCopy.map(l => l.html).join('<br>');

  const lineElements = [];
  for (let i = currentLineIndex; i < endIndex; i++) {
    const el = editor.querySelector(`.line[data-index="${i}"]`);
    if (el) lineElements.push(el);
  }

  try {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/plain': new Blob([textToCopy], { type: 'text/plain' }),
        'text/html':  new Blob([htmlToCopy], { type: 'text/html'  }),
      })
    ]);

    const msg = multiLineCount === 1
      ? `Copied line ${currentLineIndex + 1}: "${truncate(linesToCopy[0].text, 40)}"`
      : `Copied lines ${currentLineIndex + 1}–${endIndex} (${multiLineCount} lines)`;
    showStatus(msg);

    lineElements.forEach(el => el.classList.add('cut-highlight'));

    currentLineIndex = endIndex;
    multiLineCount   = 1;
    updateLineCounter();

    if (currentLineIndex >= lines.length) { showStatus("All lines processed!"); return; }
    scrollToCurrentLine();

  } catch (err) {
    showStatus("Failed to copy to clipboard");
    console.error(err);
  }
});

/**
 * Update line counter display
 */
function updateLineCounter() {
  const cur = lines.length > 0 ? Math.min(currentLineIndex + 1, lines.length) : 0;
  if (multiLineCount > 1) {
    const end = Math.min(currentLineIndex + multiLineCount, lines.length);
    lineCounter.textContent = `Lines: ${cur}–${end} / ${lines.length} (${multiLineCount} selected)`;
  } else {
    lineCounter.textContent = `Line: ${cur} / ${lines.length}`;
  }
}

function showStatus(message) {
  status.textContent = message;
  status.classList.add("show");
  setTimeout(() => status.classList.remove("show"), 2500);
}

function truncate(str, max) {
  return str.length <= max ? str : str.substring(0, max) + "...";
}

/**
 * Keyboard shortcuts
 * Enter / Space → cut
 * Ctrl+↓ → expand selection
 * Ctrl+↑ → shrink selection
 */
document.addEventListener("keydown", (e) => {
  if (e.ctrlKey && e.key === "ArrowDown") {
    e.preventDefault();
    if (multiLineCount < lines.length - currentLineIndex) {
      multiLineCount++;
      updateLineCounter();
      positionPill();
      showStatus(`Selected ${multiLineCount} lines`);
    }
    return;
  }
  if (e.ctrlKey && e.key === "ArrowUp") {
    e.preventDefault();
    if (multiLineCount > 1) {
      multiLineCount--;
      updateLineCounter();
      positionPill();
      showStatus(`Selected ${multiLineCount} lines`);
    }
    return;
  }
  if ((e.key === "Enter" || e.key === " ") && document.activeElement !== fileInput) {
    e.preventDefault();
    cutBtn.click();
  }
});

// Click a line to jump to it
editor.addEventListener("click", (e) => {
  const lineEl = e.target.closest(".line");
  if (lineEl) {
    const index = parseInt(lineEl.dataset.index, 10);
    if (!isNaN(index) && index >= 0 && index < lines.length) {
      currentLineIndex = index;
      multiLineCount   = 1;
      updateLineCounter();
      positionPill();
    }
  }
});

window.addEventListener("resize", positionPill);