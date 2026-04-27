const DAY_BATCH_SIZE = 10;
const language = document.documentElement.lang === "hy" ? "hy" : "en";
const copy = getCopy(language);
const sampleRows = createSampleRows();

const state = {
  rows: [],
  groupedSchools: [],
  dayBatchSize: DAY_BATCH_SIZE
};

const excelFileInput = document.getElementById("excelFile");
const useMockDataButton = document.getElementById("useMockData");
const downloadAllPdfsButton = document.getElementById("downloadAllPdfs");
const schoolListsContainer = document.getElementById("schoolLists");
const statusMessage = document.getElementById("statusMessage");
const studentCount = document.getElementById("studentCount");
const schoolCount = document.getElementById("schoolCount");
const pdfCount = document.getElementById("pdfCount");

excelFileInput.addEventListener("change", handleFileSelect);
useMockDataButton.addEventListener("click", () => {
  buildSchoolLists(sampleRows, copy.sampleSourceLabel);
});
downloadAllPdfsButton.addEventListener("click", () => {
  downloadAllPdfs();
});

async function handleFileSelect(event) {
  const [file] = event.target.files || [];

  if (!file) {
    return;
  }

  setStatus(copy.readingFile(file.name));

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    buildSchoolLists(rawRows, file.name);
  } catch (error) {
    console.error(error);
    resetOutput();
    setStatus(copy.invalidFile);
  }
}

function buildSchoolLists(rawRows, sourceLabel) {
  const normalizedRows = rawRows
    .map(normalizeRow)
    .filter((row) => row.studentid && row.studentname && row.school);

  if (!normalizedRows.length) {
    resetOutput();
    setStatus(copy.noValidRows);
    return;
  }

  const groupedMap = normalizedRows.reduce((accumulator, row) => {
    const schoolName = row.school.trim();

    if (!accumulator.has(schoolName)) {
      accumulator.set(schoolName, []);
    }

    accumulator.get(schoolName).push(row);
    return accumulator;
  }, new Map());

  state.rows = normalizedRows;
  state.groupedSchools = Array.from(groupedMap.entries())
    .map(([school, rows]) => createSchoolGroup(school, rows))
    .sort((left, right) => left.school.localeCompare(right.school));

  renderSchoolCards();
  updateStats();
  downloadAllPdfsButton.disabled = false;
  setStatus(copy.loadedRows(normalizedRows.length, sourceLabel, state.dayBatchSize));
}

function normalizeRow(row) {
  return {
    studentid: readValue(row, ["studentid", "student_id", "student id", "id"]),
    studentname: readValue(row, ["studentname", "student_name", "student name", "name"]),
    school: readValue(row, ["school", "schoolname", "school_name", "school name"]),
    dayLabel: normalizeDayLabel(readValue(row, ["day", "assignedday", "assigned_day", "daylabel"]))
  };
}

function readValue(row, candidateKeys) {
  const entries = Object.entries(row);

  for (const [key, value] of entries) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, "");

    if (candidateKeys.some((candidate) => normalizedKey === candidate.replace(/[^a-z0-9]/g, ""))) {
      return String(value).trim();
    }
  }

  return "";
}

function renderSchoolCards() {
  schoolListsContainer.classList.remove("empty-state");
  schoolListsContainer.innerHTML = "";

  state.groupedSchools.forEach((group) => {
    const card = document.createElement("article");
    card.className = "school-card";

    const tableRows = group.rows
      .map(
        (row) => `
          <tr>
            <td>${escapeHtml(row.dayLabel)}</td>
            <td>${escapeHtml(row.studentid)}</td>
            <td>${escapeHtml(row.studentname)}</td>
          </tr>
        `
      )
      .join("");

    card.innerHTML = `
      <div class="school-card-header">
        <div>
          <h3>${escapeHtml(group.school)}</h3>
          <div class="school-meta">
            <span class="school-tag">${group.rows.length} ${copy.studentsTag}</span>
            <span class="school-tag">${group.dayCount} ${copy.daysTag}</span>
          </div>
        </div>
        <button class="download-button" type="button">${copy.downloadPdfButton}</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>${copy.dayColumn}</th>
              <th>${copy.studentIdColumn}</th>
              <th>${copy.studentNameColumn}</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>
    `;

    card.querySelector("button").addEventListener("click", () => {
      downloadSchoolPdf(group);
    });

    schoolListsContainer.appendChild(card);
  });
}

function updateStats() {
  studentCount.textContent = String(state.rows.length);
  schoolCount.textContent = String(state.groupedSchools.length);
  pdfCount.textContent = String(state.groupedSchools.length);
}

function resetOutput() {
  state.rows = [];
  state.groupedSchools = [];
  schoolListsContainer.className = "school-lists empty-state";
  schoolListsContainer.textContent = copy.emptyState;
  downloadAllPdfsButton.disabled = true;
  updateStats();
}

async function downloadAllPdfs() {
  for (const group of state.groupedSchools) {
    await downloadSchoolPdf(group);
    await wait(250);
  }
}

async function downloadSchoolPdf(group) {
  const { jsPDF } = window.jspdf;
  const documentPdf = new jsPDF({
    orientation: "portrait",
    unit: "pt",
    format: "letter"
  });
  const useUnicodeFont = needsUnicodePdf(group);

  if (useUnicodeFont) {
    await ensurePdfFont(documentPdf);
  }

  documentPdf.setFillColor(11, 110, 79);
  documentPdf.rect(0, 0, documentPdf.internal.pageSize.getWidth(), 74, "F");
  documentPdf.setTextColor(248, 246, 240);
  documentPdf.setFont(useUnicodeFont ? embeddedPdfFont.family : "helvetica", "normal");
  documentPdf.setFontSize(24);
  documentPdf.text(group.school, 40, 45);

  documentPdf.setTextColor(93, 103, 103);
  documentPdf.setFont(useUnicodeFont ? embeddedPdfFont.family : "helvetica", "normal");
  documentPdf.setFontSize(11);
  documentPdf.text(copy.pdfSummary(group.rows.length, group.dayCount), 40, 98);

  documentPdf.autoTable({
    startY: 118,
    head: [[copy.dayColumn, copy.studentIdColumn, copy.studentNameColumn]],
    body: group.rows.map((row) => [row.dayLabel, row.studentid, row.studentname]),
    theme: "grid",
    headStyles: {
      font: useUnicodeFont ? embeddedPdfFont.family : "helvetica",
      fontStyle: "normal",
      fillColor: [23, 33, 33],
      textColor: [248, 246, 240]
    },
    styles: {
      font: useUnicodeFont ? embeddedPdfFont.family : "helvetica",
      fontStyle: "normal",
      fontSize: 10,
      cellPadding: 8
    },
    alternateRowStyles: {
      fillColor: [247, 242, 234]
    }
  });

  documentPdf.save(`${toFilename(group.school)}-students.pdf`);
}

function toFilename(value) {
  const normalized = String(value)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");

  return normalized || `school-${simpleHash(value)}`;
}

function createSchoolGroup(school, rows) {
  const hasExplicitDays = rows.every((row) => row.dayLabel);
  const assignedRows = hasExplicitDays
    ? rows.map((row) => {
        const dayNumber = extractDayNumber(row.dayLabel);

        return {
          ...row,
          dayNumber,
          dayLabel: formatDayLabel(dayNumber)
        };
      })
    : shuffleRows(rows.map((row) => ({ ...row }))).map((row, index) => {
        const dayNumber = Math.floor(index / state.dayBatchSize) + 1;

        return {
          ...row,
          dayNumber,
          dayLabel: formatDayLabel(dayNumber)
        };
      });

  assignedRows.sort((left, right) => {
    if (left.dayNumber !== right.dayNumber) {
      return left.dayNumber - right.dayNumber;
    }

    return left.studentname.localeCompare(right.studentname);
  });

  return {
    school,
    rows: assignedRows,
    dayCount: Math.max(...assignedRows.map((row) => row.dayNumber))
  };
}

function shuffleRows(rows) {
  const shuffled = [...rows];

  for (let index = shuffled.length - 1; index > 0; index -= 1) {
    const randomIndex = Math.floor(Math.random() * (index + 1));
    const currentRow = shuffled[index];
    shuffled[index] = shuffled[randomIndex];
    shuffled[randomIndex] = currentRow;
  }

  return shuffled;
}

function createSampleRows() {
  const schools = [
    { name: "North Ridge Academy", prefix: "NRA", start: 1 },
    { name: "South Valley School", prefix: "SVS", start: 51 }
  ];

  return schools.flatMap(({ name, prefix, start }) => {
    const schoolRows = Array.from({ length: 50 }, (_, index) => {
      const studentNumber = String(start + index).padStart(3, "0");

      return {
        studentid: `${prefix}-${studentNumber}`,
        studentname: `Student ${studentNumber}`,
        school: name
      };
    });

    return shuffleRows(schoolRows).map((row, index) => ({
      ...row,
      dayLabel: formatDayLabel(Math.floor(index / DAY_BATCH_SIZE) + 1)
    }));
  });
}

function needsUnicodePdf(group) {
  if (language === "hy") {
    return true;
  }

  if (containsArmenianText(group.school)) {
    return true;
  }

  return group.rows.some((row) => containsArmenianText(row.studentname) || containsArmenianText(row.dayLabel));
}

function containsArmenianText(value) {
  return /[\u0531-\u058F]/.test(String(value));
}

async function ensurePdfFont(documentPdf) {
  if (!embeddedPdfFont) {
    throw new Error("Embedded Armenian PDF font is missing.");
  }

  if (!documentPdf.getFontList()[embeddedPdfFont.family]) {
    documentPdf.addFileToVFS(embeddedPdfFont.fileName, embeddedPdfFont.base64);
    documentPdf.addFont(embeddedPdfFont.fileName, embeddedPdfFont.family, "normal");
  }
}

function normalizeDayLabel(value) {
  if (!value) {
    return "";
  }

  const dayNumber = extractDayNumber(value);
  return dayNumber ? formatDayLabel(dayNumber) : "";
}

function extractDayNumber(value) {
  const match = String(value).match(/\d+/);
  return match ? Number(match[0]) : Number.NaN;
}

function wait(milliseconds) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

function simpleHash(value) {
  return Array.from(String(value)).reduce((hash, character) => {
    return (hash * 31 + character.charCodeAt(0)) >>> 0;
  }, 7);
}

function formatDayLabel(dayNumber) {
  return copy.dayLabel(dayNumber);
}

function setStatus(message) {
  statusMessage.textContent = message;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getCopy(activeLanguage) {
  const copyByLanguage = {
    en: {
      sampleSourceLabel: "Built-in sample data",
      invalidFile: "Could not read the Excel file. Confirm the file is a valid .xlsx or .xls.",
      noValidRows:
        "No valid rows found. Your file needs the columns studentid, studentname, and school.",
      emptyState: "Upload a file to generate school rosters with day assignments.",
      downloadPdfButton: "Download PDF",
      studentsTag: "students",
      daysTag: "days",
      dayColumn: "Day",
      studentIdColumn: "Student ID",
      studentNameColumn: "Student Name",
      dayLabel: (dayNumber) => `Day ${dayNumber}`,
      readingFile: (fileName) => `Reading ${fileName}...`,
      loadedRows: (rowCount, sourceLabel, batchSize) =>
        `Loaded ${rowCount} students from ${sourceLabel}. Day assignments were created automatically in batches of ${batchSize}.`,
      pdfSummary: (rowCount, dayCount) =>
        `Generated roster • ${rowCount} students • ${dayCount} days`
    },
    hy: {
      sampleSourceLabel: "Ներկառուցված նմուշային տվյալներ",
      invalidFile:
        "Չհաջողվեց կարդալ Excel ֆայլը։ Համոզվեք, որ ֆայլը վավեր `.xlsx` կամ `.xls` է։",
      noValidRows:
        "Վավեր տողեր չեն գտնվել։ Ձեր ֆայլը պետք է ունենա studentid, studentname և school սյունակները։",
      emptyState:
        "Վերբեռնեք ֆայլ՝ օրերի բաժանումով դպրոցական ցուցակներ ստեղծելու համար։",
      downloadPdfButton: "Ներբեռնել PDF",
      studentsTag: "աշակերտ",
      daysTag: "օր",
      dayColumn: "Օր",
      studentIdColumn: "Աշակերտի ID",
      studentNameColumn: "Աշակերտի անուն",
      dayLabel: (dayNumber) => `Օր ${dayNumber}`,
      readingFile: (fileName) => `Ընթերցվում է ${fileName} ֆայլը...`,
      loadedRows: (rowCount, sourceLabel, batchSize) =>
        `Բեռնվել է ${rowCount} աշակերտ ${sourceLabel} աղբյուրից։ Օրերի բաժանումը ստեղծվել է ինքնաշխատ՝ յուրաքանչյուր ${batchSize} աշակերտից մեկ խմբով։`,
      pdfSummary: (rowCount, dayCount) =>
        `Ստեղծված ցուցակ • ${rowCount} աշակերտ • ${dayCount} օր`
    },
    en: {
      sampleSourceLabel: "Built-in sample data",
      invalidFile: "Could not read the Excel file. Confirm the file is a valid .xlsx or .xls.",
      noValidRows:
        "No valid rows found. Your file needs the columns studentid, studentname, and school.",
      emptyState: "Upload a file to generate school rosters with day assignments.",
      downloadPdfButton: "Download PDF",
      studentsTag: "students",
      daysTag: "days",
      dayColumn: "Day",
      studentIdColumn: "Student ID",
      studentNameColumn: "Student Name",
      versionLabel: (version) => `Version ${version}`,
      dayLabel: (dayNumber) => `Day ${dayNumber}`,
      readingFile: (fileName) => `Reading ${fileName}...`,
      loadedRows: (rowCount, sourceLabel, batchSize) =>
        `Loaded ${rowCount} students from ${sourceLabel}. Day assignments were created automatically in batches of ${batchSize}.`,
      pdfSummary: (rowCount, dayCount) =>
        `Generated roster • ${rowCount} students • ${dayCount} days`
    }
  };

  return copyByLanguage[activeLanguage];
}
