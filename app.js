const photosInput = document.getElementById('photos');
const excelInput = document.getElementById('excel');
const processBtn = document.getElementById('processBtn');
const step1Status = document.getElementById('step1Status');
const preview = document.getElementById('preview');
const resultWrap = document.getElementById('resultWrap');
const confirmBtn = document.getElementById('confirmBtn');

const step2 = document.getElementById('step2');
const softwareFileInput = document.getElementById('softwareFile');
const softwarePreview = document.getElementById('softwarePreview');
const step2Status = document.getElementById('step2Status');
const downloadTemplateBtn = document.getElementById('downloadTemplate');

let mergedRows = [];

processBtn.addEventListener('click', async () => {
  const excelFile = excelInput.files?.[0];
  const photoFiles = Array.from(photosInput.files ?? []);

  if (!excelFile) {
    setStatus(step1Status, 'Please upload the original Excel file first.', true);
    return;
  }

  try {
    setStatus(step1Status, 'Reading Excel file...');
    const rows = await parseExcel(excelFile);

    setStatus(step1Status, 'Extracting text from photo(s) with OCR...');
    const ocrNotes = await parsePhotosWithOCR(photoFiles);

    mergedRows = attachOCR(rows, ocrNotes);
    renderTable(mergedRows);
    resultWrap.classList.remove('hidden');

    setStatus(
      step1Status,
      `Done. Loaded ${rows.length} row(s) from Excel and processed ${photoFiles.length} photo(s).`
    );
  } catch (error) {
    console.error(error);
    setStatus(step1Status, `Processing failed: ${error.message}`, true);
  }
});

confirmBtn.addEventListener('click', () => {
  if (!mergedRows.length) {
    setStatus(step1Status, 'No auto-filled data found. Please run Step 1 first.', true);
    return;
  }
  step2.classList.remove('hidden');
  setStatus(step2Status, 'Step 2 unlocked. Upload software format file.');
  step2.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

softwareFileInput.addEventListener('change', async () => {
  const file = softwareFileInput.files?.[0];
  if (!file) {
    return;
  }

  try {
    let parsed;
    if (file.name.endsWith('.json')) {
      parsed = JSON.parse(await file.text());
    } else {
      parsed = csvToObjects(await file.text());
    }

    softwarePreview.textContent = JSON.stringify(parsed, null, 2);
    setStatus(step2Status, `Uploaded ${file.name}. File is ready to be shared with software.`);
  } catch (error) {
    console.error(error);
    setStatus(step2Status, `Could not parse file: ${error.message}`, true);
  }
});

downloadTemplateBtn.addEventListener('click', () => {
  const template = {
    source: 'field-book-digitizer',
    version: 1,
    observations: mergedRows,
  };
  const blob = new Blob([JSON.stringify(template, null, 2)], { type: 'application/json' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'software-upload-template.json';
  link.click();
  URL.revokeObjectURL(link.href);
});

async function parseExcel(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    defval: '',
    raw: false,
  });
}

async function parsePhotosWithOCR(photoFiles) {
  if (!photoFiles.length) {
    return [];
  }

  const notes = [];

  for (const file of photoFiles) {
    const { data } = await Tesseract.recognize(file, 'eng');
    notes.push({ fileName: file.name, text: data.text.trim() });
  }

  return notes;
}

function attachOCR(rows, ocrNotes) {
  const noteText = ocrNotes.map((n) => `${n.fileName}: ${n.text}`).join('\n');

  return rows.map((row) => ({
    ...row,
    OCR_Notes: noteText || 'No photo OCR content provided',
  }));
}

function renderTable(rows) {
  if (!rows.length) {
    preview.innerHTML = '<p class="muted">No rows found in spreadsheet.</p>';
    return;
  }

  const columns = [...new Set(rows.flatMap((row) => Object.keys(row)))];
  const head = `<tr>${columns.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}</tr>`;
  const body = rows
    .map(
      (row) =>
        `<tr>${columns
          .map((c) => `<td>${escapeHtml(String(row[c] ?? ''))}</td>`)
          .join('')}</tr>`
    )
    .join('');

  preview.innerHTML = `<div class="table-wrap"><table>${head}${body}</table></div>`;
}

function csvToObjects(csv) {
  const lines = csv.trim().split(/\r?\n/);
  if (!lines.length) {
    return [];
  }

  const headers = lines[0].split(',').map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const values = line.split(',').map((v) => v.trim());
    return headers.reduce((acc, header, idx) => {
      acc[header] = values[idx] ?? '';
      return acc;
    }, {});
  });
}

function setStatus(node, message, isError = false) {
  node.textContent = message;
  node.style.color = isError ? '#b91c1c' : '#111827';
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
