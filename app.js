const fieldBookFilesInput = document.getElementById('fieldBookFiles');
const originalExcelInput = document.getElementById('originalExcel');
const fillBtn = document.getElementById('fillBtn');
const step1Status = document.getElementById('step1Status');
const step1Summary = document.getElementById('step1Summary');
const step1PreviewSection = document.getElementById('step1PreviewSection');
const filledPreview = document.getElementById('filledPreview');
const detectedProjectsPreview = document.getElementById('detectedProjectsPreview');
const mappedColumnsStep1Preview = document.getElementById('mappedColumnsStep1Preview');
const downloadFilledExcelBtn = document.getElementById('downloadFilledExcel');
const continueBtn = document.getElementById('continueBtn');

const step2Section = document.getElementById('step2Section');
const softwareModelInput = document.getElementById('softwareModel');
const convertBtn = document.getElementById('convertBtn');
const step2Status = document.getElementById('step2Status');
const step2PreviewSection = document.getElementById('step2PreviewSection');
const softwarePreview = document.getElementById('softwarePreview');
const mappedColumnsStep2Preview = document.getElementById('mappedColumnsStep2Preview');
const projectSummaryPreview = document.getElementById('projectSummaryPreview');
const downloadSoftwareTemplateBtn = document.getElementById('downloadSoftwareTemplate');
const downloadProjectZipBtn = document.getElementById('downloadProjectZip');

let filledRows = [];
let filledExcelFile = null;
let softwareTemplateFile = null;
let projectZipFile = null;

fillBtn.addEventListener('click', async () => {
  const excelFile = originalExcelInput.files?.[0];
  const files = Array.from(fieldBookFilesInput.files ?? []);

  if (!excelFile) {
    setStatus(step1Status, 'Please upload the original Excel file.', true);
    return;
  }

  try {
    const formData = new FormData();
    formData.append('excel', excelFile);
    files.forEach((file) => formData.append('fieldBookFiles', file));

    setStatus(step1Status, 'Running OCR and filling original Excel...');

    const response = await fetch('/api/fill', { method: 'POST', body: formData });
    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.error || 'Step 1 failed.');
    }

    filledRows = payload.filledRows;
    filledExcelFile = payload.filledExcel;

    renderTable(filledPreview, filledRows, 'No filled rows generated.');
    renderDetectedProjects(payload.detectedProjects || []);
    renderMappedColumns(mappedColumnsStep1Preview, payload.mappedColumns || [], 'sourceKey', 'targetColumn');

    step1Summary.textContent = `${payload.summary.appliedMatchCount} OCR values matched columns, ${payload.summary.unmatchedCount} unmatched, ${payload.summary.totalPairs} extracted.`;
    step1PreviewSection.classList.remove('hidden');
    step2Section.classList.add('hidden');
    step2PreviewSection.classList.add('hidden');

    setStatus(step1Status, 'Step 1 completed successfully.');
  } catch (error) {
    setStatus(step1Status, `Step 1 error: ${error.message}`, true);
  }
});

continueBtn.addEventListener('click', () => {
  if (!filledRows.length) {
    setStatus(step1Status, 'Run Step 1 first before continuing.', true);
    return;
  }

  step2Section.classList.remove('hidden');
  setStatus(step2Status, 'Upload software model and click Convert and Split.');
  step2Section.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

convertBtn.addEventListener('click', async () => {
  const softwareModel = softwareModelInput.files?.[0];
  if (!softwareModel) {
    setStatus(step2Status, 'Please upload Software Model Sheet.', true);
    return;
  }

  if (!filledRows.length) {
    setStatus(step2Status, 'Filled rows missing. Complete Step 1 first.', true);
    return;
  }

  try {
    const formData = new FormData();
    formData.append('softwareModel', softwareModel);
    formData.append('filledRows', JSON.stringify(filledRows));

    setStatus(step2Status, 'Converting to software template and splitting by Project Type...');

    const response = await fetch('/api/convert', { method: 'POST', body: formData });
    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.error || 'Step 2 failed.');
    }

    softwareTemplateFile = payload.softwareTemplate;
    projectZipFile = payload.projectZip;

    renderMappedColumns(mappedColumnsStep2Preview, payload.mappedColumns || [], 'targetColumn', 'sourceColumn');
    renderTable(softwarePreview, payload.softwareTemplateRows, 'No software template rows generated.');
    renderProjectSummary(payload.projectPreview || []);
    step2PreviewSection.classList.remove('hidden');

    setStatus(step2Status, `Step 2 completed. Generated ${payload.projectPreview.length} project group(s).`);
  } catch (error) {
    setStatus(step2Status, `Step 2 error: ${error.message}`, true);
  }
});

downloadFilledExcelBtn.addEventListener('click', () => {
  if (!filledExcelFile) {
    setStatus(step1Status, 'No filled Excel file available.', true);
    return;
  }
  downloadBase64File(filledExcelFile.fileName, filledExcelFile.base64, filledExcelFile.mimeType);
});

downloadSoftwareTemplateBtn.addEventListener('click', () => {
  if (!softwareTemplateFile) {
    setStatus(step2Status, 'No software template file available.', true);
    return;
  }
  downloadBase64File(softwareTemplateFile.fileName, softwareTemplateFile.base64, softwareTemplateFile.mimeType);
});

downloadProjectZipBtn.addEventListener('click', () => {
  if (!projectZipFile) {
    setStatus(step2Status, 'No project ZIP file available.', true);
    return;
  }
  downloadBase64File(projectZipFile.fileName, projectZipFile.base64, projectZipFile.mimeType);
});

function renderDetectedProjects(items) {
  if (!items.length) {
    detectedProjectsPreview.innerHTML = '<p class="muted">No project type detected from OCR pages.</p>';
    return;
  }

  const rows = items.map((item) => ({ Page: item.page, 'Detected Project Type': item.projectType }));
  renderTable(detectedProjectsPreview, rows, 'No project type detected from OCR pages.');
}

function renderMappedColumns(container, items, leftKey, rightKey) {
  if (!items.length) {
    container.innerHTML = '<p class="muted">No column mappings available.</p>';
    return;
  }

  const rows = items.map((item) => ({
    Left: item[leftKey] ?? '',
    Right: item[rightKey] ?? '',
    Score: item.score ?? '',
  }));
  renderTable(container, rows, 'No column mappings available.');
}

function renderProjectSummary(items) {
  if (!items.length) {
    projectSummaryPreview.innerHTML = '<p class="muted">No project split information available.</p>';
    return;
  }

  const rows = items.map((item) => ({ 'Project Name': item.projectName, 'Number of Records': item.recordCount }));
  renderTable(projectSummaryPreview, rows, 'No project split information available.');
}

function downloadBase64File(fileName, base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  const blob = new Blob([bytes], { type: mimeType });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(link.href);
}

function renderTable(container, rows, emptyMessage = 'No rows available.') {
  if (!rows?.length) {
    container.innerHTML = `<p class="muted">${escapeHtml(emptyMessage)}</p>`;
    return;
  }

  const columns = [...new Set(rows.flatMap((row) => Object.keys(row)))];
  const head = `<tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('')}</tr>`;
  const body = rows
    .map(
      (row) =>
        `<tr>${columns
          .map((column) => `<td>${escapeHtml(String(row[column] ?? ''))}</td>`)
          .join('')}</tr>`
    )
    .join('');

  container.innerHTML = `<div class="table-wrap"><table>${head}${body}</table></div>`;
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
