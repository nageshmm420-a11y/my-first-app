const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { createWorker, PSM } = require('tesseract.js');
const JSZip = require('jszip');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 3000;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';

app.use(express.static(__dirname));

app.post('/api/fill', upload.fields([{ name: 'excel', maxCount: 1 }, { name: 'fieldBookFiles', maxCount: 30 }]), async (req, res) => {
  const excelFile = req.files?.excel?.[0];
  const fieldBookFiles = req.files?.fieldBookFiles || [];

  if (!excelFile) {
    return res.status(400).json({ error: 'Original Excel file is required for Step 1.' });
  }

  try {
    const rows = parseExcelBuffer(excelFile.buffer);
    if (!rows.length) {
      return res.status(400).json({ error: 'No data rows found in original Excel first worksheet.' });
    }

    const ocrResult = await parseFieldBookFilesWithOCR(fieldBookFiles);
    const mergeResult = attachOCR(rows, ocrResult);
    const filledExcelBuffer = buildWorkbookFromRows(mergeResult.rows, 'Filled_Original');

    return res.json({
      filledRows: mergeResult.rows,
      detectedProjects: ocrResult.detectedProjects,
      mappedColumns: mergeResult.mappedColumns,
      summary: {
        appliedMatchCount: mergeResult.appliedMatchCount,
        unmatchedCount: mergeResult.unmatchedCount,
        totalPairs: ocrResult.totalPairs,
      },
      filledExcel: {
        fileName: 'filled-original.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        base64: filledExcelBuffer.toString('base64'),
      },
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: `Step 1 failed: ${error.message}` });
  }
});

app.post('/api/convert', upload.fields([{ name: 'softwareModel', maxCount: 1 }]), async (req, res) => {
  const softwareModelFile = req.files?.softwareModel?.[0];
  const filledRowsRaw = req.body?.filledRows;

  if (!softwareModelFile) {
    return res.status(400).json({ error: 'Software Model Sheet is required for Step 2.' });
  }

  try {
    const filledRows = JSON.parse(filledRowsRaw || '[]');
    if (!Array.isArray(filledRows) || !filledRows.length) {
      return res.status(400).json({ error: 'Filled rows are required. Please complete Step 1 first.' });
    }

    const templateResult = mapToSoftwareTemplate(filledRows, softwareModelFile.buffer);
    const softwareTemplateBuffer = buildWorkbookFromRows(templateResult.rows, 'Software_Template');
    const projectSplit = await buildProjectTypeFiles(templateResult.rows);

    return res.json({
      softwareTemplateRows: templateResult.rows,
      mappedColumns: templateResult.mappedColumns,
      projectPreview: projectSplit.preview,
      softwareTemplate: {
        fileName: 'software-template.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        base64: softwareTemplateBuffer.toString('base64'),
      },
      projectZip: projectSplit.zip,
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: `Step 2 failed: ${error.message}` });
  }
});

function parseExcelBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
}

function buildWorkbookFromRows(rows, sheetName) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{}]);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

async function parseFieldBookFilesWithOCR(files) {
  if (!files.length) {
    return { notes: [], totalPairs: 0, detectedProjects: [] };
  }

  const ocrInputs = [];
  for (const file of files) {
    if (isPdf(file)) {
      const pdfImages = await convertPdfToImageBuffers(file.buffer, file.originalname);
      ocrInputs.push(...pdfImages);
    } else {
      ocrInputs.push({ name: file.originalname, buffer: file.buffer });
    }
  }

  return parseImagesWithOCR(ocrInputs);
}

async function parseImagesWithOCR(inputs) {
  if (!inputs.length) {
    return { notes: [], totalPairs: 0, detectedProjects: [] };
  }

  const worker = await createWorker('eng');
  await worker.setParameters({
    tessedit_pageseg_mode: PSM.AUTO,
    preserve_interword_spaces: '1',
    user_defined_dpi: '300',
  });

  const notes = [];
  const detectedProjects = [];
  let totalPairs = 0;

  try {
    for (const input of inputs) {
      const preprocessed = await preprocessImageBuffer(input.buffer);
      const { data } = await worker.recognize(preprocessed);
      const rawText = data.text.trim();
      const cleanedText = await cleanupOCRText(rawText);
      const pairs = extractKeyValuePairs(cleanedText);
      const projectType = detectProjectType(cleanedText);

      totalPairs += pairs.length;
      notes.push({ fileName: input.name, text: cleanedText, pairs, projectType });

      if (projectType) {
        detectedProjects.push({ page: input.name, projectType });
      }
    }
  } finally {
    await worker.terminate();
  }

  return { notes, totalPairs, detectedProjects };
}

async function preprocessImageBuffer(buffer) {
  let sharp;
  try {
    sharp = require('sharp');
  } catch (_error) {
    throw new Error('Image preprocessing requires `sharp` dependency.');
  }

  // grayscale + contrast normalization + noise reduction + basic deskew(auto rotate by EXIF)
  return sharp(buffer)
    .rotate()
    .grayscale()
    .normalize()
    .median(1)
    .sharpen()
    .png()
    .toBuffer();
}

async function cleanupOCRText(text) {
  if (!text) {
    return '';
  }

  if (!OPENAI_API_KEY) {
    return text.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  }

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        temperature: 0,
        messages: [
          {
            role: 'system',
            content:
              'You clean OCR text from field books. Fix obvious spelling noise and preserve key-value structure. Return plain text only.',
          },
          { role: 'user', content: text },
        ],
      }),
    });

    if (!response.ok) {
      return text;
    }

    const payload = await response.json();
    return payload.choices?.[0]?.message?.content?.trim() || text;
  } catch (_error) {
    return text;
  }
}

function detectProjectType(text) {
  if (!text) {
    return null;
  }

  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const pattern = /(project|trial|program|nursery|experiment)\s*[:\-]?\s*(.+)/i;

  for (const line of lines) {
    const match = line.match(pattern);
    if (match) {
      return match[2].trim();
    }
  }

  return null;
}

async function convertPdfToImageBuffers(pdfBuffer, originalName) {
  let pdfjsLib;
  let createCanvas;

  try {
    pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
    ({ createCanvas } = await import('canvas'));
  } catch (error) {
    throw new Error('PDF support requires dependencies `pdfjs-dist` and `canvas` to be installed.');
  }

  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(pdfBuffer) });
  const pdf = await loadingTask.promise;
  const images = [];

  for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
    const page = await pdf.getPage(pageIndex);
    const viewport = page.getViewport({ scale: 2.0 });
    const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
    const context = canvas.getContext('2d');

    await page.render({ canvasContext: context, viewport }).promise;
    images.push({ name: `${stripExtension(originalName)}_page_${pageIndex}.png`, buffer: canvas.toBuffer('image/png') });
  }

  return images;
}

function stripExtension(fileName) {
  return String(fileName || '').replace(/\.[^.]+$/, '');
}

function isPdf(file) {
  return file.mimetype === 'application/pdf' || /\.pdf$/i.test(file.originalname || '');
}

async function buildProjectTypeFiles(rows) {
  const projectColumn = findProjectTypeColumn(rows);
  if (!projectColumn) {
    throw new Error('Project Type column not found in software template output.');
  }

  const groups = new Map();
  rows.forEach((row) => {
    const rawProject = String(row[projectColumn] ?? '').trim() || 'unassigned_project';
    if (!groups.has(rawProject)) {
      groups.set(rawProject, []);
    }
    groups.get(rawProject).push(row);
  });

  const preview = [];
  const zip = new JSZip();

  for (const [projectName, projectRows] of groups.entries()) {
    const safeName = sanitizeProjectFileName(projectName);
    const fileName = `${safeName}.xlsx`;
    const fileBuffer = buildWorkbookFromRows(projectRows, 'Software_Template');
    preview.push({ projectName, recordCount: projectRows.length });
    zip.file(fileName, fileBuffer);
  }

  const zipBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  return {
    preview,
    zip: { fileName: 'project_files.zip', mimeType: 'application/zip', base64: zipBuffer.toString('base64') },
  };
}

function findProjectTypeColumn(rows) {
  if (!rows.length) {
    return null;
  }
  const columns = [...new Set(rows.flatMap((row) => Object.keys(row)))];
  return columns.find((column) => normalizeKey(column) === 'projecttype') || null;
}

function sanitizeProjectFileName(name) {
  const cleaned = String(name || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\s_-]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return cleaned || 'unassigned_project';
}

function mapToSoftwareTemplate(filledRows, softwareTemplateBuffer) {
  const templateWorkbook = XLSX.read(softwareTemplateBuffer, { type: 'buffer' });
  const templateSheetName = templateWorkbook.SheetNames[0];
  const templateSheet = templateWorkbook.Sheets[templateSheetName];
  const templateMatrix = XLSX.utils.sheet_to_json(templateSheet, { header: 1, blankrows: false, defval: '' });

  const templateHeaders = (templateMatrix[0] || []).map((header) => String(header).trim()).filter(Boolean);
  if (!templateHeaders.length) {
    throw new Error('Software model sheet must have headers in the first row.');
  }

  const sourceColumns = [...new Set(filledRows.flatMap((row) => Object.keys(row)))];
  const sourceMatcher = buildColumnMatcher(sourceColumns);
  const mappedColumns = [];

  const rows = filledRows.map((row) => {
    const mapped = {};

    templateHeaders.forEach((header) => {
      const result = findBestColumnMatchDetailed(header, sourceMatcher);
      mapped[header] = result.column ? row[result.column] ?? '' : '';
      mappedColumns.push({ targetColumn: header, sourceColumn: result.column || null, score: result.score });
    });

    return mapped;
  });

  return { rows, mappedColumns: uniqMappedColumns(mappedColumns) };
}

function attachOCR(rows, ocrResult) {
  const columns = [...new Set(rows.flatMap((row) => Object.keys(row)))];
  const columnMatcher = buildColumnMatcher(columns);

  const allPairs = ocrResult.notes.flatMap((note) =>
    note.pairs.map((pair) => ({ ...pair, fileName: note.fileName, projectType: note.projectType }))
  );

  let appliedMatchCount = 0;
  const unmatchedPairs = [];
  const matchedValues = new Map();
  const mappedColumns = [];

  allPairs.forEach((pair) => {
    const result = findBestColumnMatchDetailed(pair.key, columnMatcher);
    if (!result.column) {
      unmatchedPairs.push(pair);
      return;
    }

    matchedValues.set(result.column, pair.value);
    mappedColumns.push({ sourceKey: pair.key, targetColumn: result.column, score: result.score });
    appliedMatchCount += 1;
  });

  const noteText = unmatchedPairs.map((pair) => `${pair.fileName} | ${pair.key}: ${pair.value}`).join('\n');
  const firstProject = ocrResult.notes.find((n) => n.projectType)?.projectType || '';

  const transformedRows = rows.map((row) => {
    const output = { ...row };
    matchedValues.forEach((value, column) => {
      if (!String(output[column] ?? '').trim()) {
        output[column] = value;
      }
    });

    if (!String(output.Project_Type || output['Project Type'] || '').trim() && firstProject) {
      if (Object.prototype.hasOwnProperty.call(output, 'Project Type')) {
        output['Project Type'] = firstProject;
      } else {
        output.Project_Type = firstProject;
      }
    }

    output.OCR_Notes = noteText || 'All OCR fields matched existing Excel columns.';
    return output;
  });

  return {
    rows: transformedRows,
    mappedColumns: uniqMappedColumns(mappedColumns),
    appliedMatchCount,
    unmatchedCount: unmatchedPairs.length,
  };
}

function uniqMappedColumns(items) {
  const seen = new Set();
  const out = [];
  items.forEach((item) => {
    const key = `${item.sourceKey || item.targetColumn}->${item.targetColumn || item.sourceColumn}`;
    if (!seen.has(key)) {
      seen.add(key);
      out.push(item);
    }
  });
  return out;
}

function extractKeyValuePairs(text) {
  if (!text) {
    return [];
  }

  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const pairs = [];

  for (const line of lines) {
    const keyValueMatch = line.match(/^([\w\s()./#-]{2,60})\s*[:=\-]\s*(.+)$/i);
    if (keyValueMatch) {
      pairs.push({ key: keyValueMatch[1].trim(), value: keyValueMatch[2].trim() });
      continue;
    }

    const tabularMatch = line.match(/^([\w\s()./#-]{2,60})\s{2,}(.+)$/i);
    if (tabularMatch) {
      pairs.push({ key: tabularMatch[1].trim(), value: tabularMatch[2].trim() });
    }
  }

  return pairs.filter((pair) => pair.key && pair.value);
}

function buildColumnMatcher(columns) {
  return columns.map((column) => {
    const key = normalizeKey(column);
    const words = String(column).toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
    return { column, key, aliases: new Set([key, ...words]) };
  });
}

function findBestColumnMatchDetailed(inputKey, columnMatcher) {
  const normalizedInput = normalizeKey(inputKey);
  if (!normalizedInput) {
    return { column: null, score: 0 };
  }

  let best = null;
  let bestScore = 0;

  columnMatcher.forEach((item) => {
    let score = 0;

    if (item.aliases.has(normalizedInput)) {
      score = 100;
    } else if (item.key.includes(normalizedInput) || normalizedInput.includes(item.key)) {
      score = 90;
    } else {
      const overlap = longestCommonSubsequence(item.key, normalizedInput);
      const overlapScore = Math.round((overlap / Math.max(item.key.length, normalizedInput.length, 1)) * 100);
      const editScore = 100 - levenshteinDistance(item.key, normalizedInput) * 5;
      score = Math.max(overlapScore, editScore);
    }

    if (score > bestScore) {
      bestScore = score;
      best = item.column;
    }
  });

  return { column: bestScore >= 55 ? best : null, score: bestScore };
}

function longestCommonSubsequence(a, b) {
  const dp = Array.from({ length: a.length + 1 }, () => Array(b.length + 1).fill(0));
  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
    }
  }
  return dp[a.length][b.length];
}

function levenshteinDistance(a, b) {
  const dp = Array.from({ length: a.length + 1 }, (_, i) => Array.from({ length: b.length + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0)));
  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[a.length][b.length];
}

function normalizeKey(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
