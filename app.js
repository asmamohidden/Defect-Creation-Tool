const FIELDS = [
  ["defectId", "Defect ID"],
  ["assetTag", "Asset Tag"],
  ["environment", "Environment"],
  ["status", "Status"],
  ["severity", "Severity"],
  ["priority", "Priority"],
  ["summary", "Summary"],
  ["description", "Description"],
  ["stepsToReproduce", "Steps to Reproduce"],
  ["expectedResult", "Expected Result"],
  ["actualResult", "Actual Result"],
  ["investigationPerformed", "Investigation Performed"],
  ["evidence", "Evidence"],
  ["businessImpact", "Business Impact"],
  ["poc", "POC"],
  ["storyId", "Story ID"],
  ["release", "Release"],
  ["createdBy", "Created By"],
  ["createdDate", "Created Date"],
  ["assignedTo", "Assigned To"],
  ["comments", "Comments"],
];

const STORAGE_KEY = "defect-tracker-log";
const form = document.getElementById("defect-form");
const banner = document.getElementById("banner");
const logBody = document.getElementById("log-body");
const emptyLog = document.getElementById("empty-log");

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function nextDefectId() {
  const stamp = todayIso().replaceAll("-", "");
  const seq = String(Date.now()).slice(-4);
  return `DEF-${stamp}-${seq}`;
}

function showBanner(type, message) {
  banner.className = `banner show ${type}`;
  banner.textContent = message;
}

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function readDefect() {
  const data = {};
  for (const [id] of FIELDS) {
    data[id] = form.elements[id].value.trim();
  }
  return data;
}

function loadLog() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveLog(items) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
}

function renderLog() {
  const items = loadLog();
  logBody.innerHTML = "";
  emptyLog.hidden = items.length > 0;
  for (const item of items.slice().reverse()) {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${item.defectId}</td>
      <td>${item.summary}</td>
      <td>${item.severity}</td>
      <td>${item.priority}</td>
      <td>${item.status}</td>
      <td>${item.assignedTo}</td>
    `;
    logBody.appendChild(row);
  }
}

function excelCell(value) {
  return `<Cell><Data ss:Type="String">${escapeXml(value)}</Data></Cell>`;
}

function buildWorkbook(defect, log) {
  const detailRows = FIELDS.map(
    ([id, label]) => `<Row>${excelCell(label)}${excelCell(defect[id])}</Row>`
  ).join("");

  const headerRow = `<Row>${FIELDS.map(([, label]) => excelCell(label)).join("")}</Row>`;
  const logRows = log
    .map((item) => `<Row>${FIELDS.map(([id]) => excelCell(item[id])).join("")}</Row>`)
    .join("");

  return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="header">
   <Font ss:Bold="1"/>
   <Interior ss:Color="#1F4E79" ss:Pattern="Solid"/>
   <Font ss:Color="#FFFFFF" ss:Bold="1"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Defect">
  <Table>
   <Row>${excelCell("Field")}${excelCell("Value")}</Row>
   ${detailRows}
  </Table>
 </Worksheet>
 <Worksheet ss:Name="Defect Log">
  <Table>
   ${headerRow}
   ${logRows}
  </Table>
 </Worksheet>
</Workbook>`;
}

function downloadExcel(defect, log) {
  const xml = buildWorkbook(defect, log);
  const blob = new Blob([xml], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${defect.defectId || "defect"}.xls`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function resetFormKeepIdentity() {
  const createdBy = form.elements.createdBy.value;
  form.reset();
  form.elements.defectId.value = nextDefectId();
  form.elements.createdDate.value = todayIso();
  form.elements.createdBy.value = createdBy;
  form.elements.status.value = "New";
  form.elements.severity.value = "Medium";
  form.elements.priority.value = "P3";
  form.elements.environment.value = "QA";
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  const defect = readDefect();

  if (!defect.defectId || !defect.summary || !defect.createdBy) {
    showBanner("error", "Defect ID, Summary, and Created By are required.");
    return;
  }

  const log = loadLog();
  log.push(defect);
  saveLog(log);
  downloadExcel(defect, log);
  renderLog();
  resetFormKeepIdentity();
  showBanner(
    "ok",
    `${defect.defectId} created. Excel file downloaded. Open it in Excel to review the defect and the log sheet.`
  );
});

document.getElementById("reset-form").addEventListener("click", () => {
  resetFormKeepIdentity();
  showBanner("ok", "Form cleared. A new Defect ID was generated.");
});

form.elements.defectId.value = nextDefectId();
form.elements.createdDate.value = todayIso();
renderLog();
