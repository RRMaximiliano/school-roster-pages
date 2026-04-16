const DAY_BATCH_SIZE = 10;
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
  buildSchoolLists(sampleRows, "Built-in sample data");
});
downloadAllPdfsButton.addEventListener("click", downloadAllPdfs);

async function handleFileSelect(event) {
  const [file] = event.target.files || [];

  if (!file) {
    return;
  }

  setStatus(`Reading ${file.name}...`);

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
    setStatus("Could not read the Excel file. Confirm the file is a valid .xlsx or .xls.");
  }
}

function buildSchoolLists(rawRows, sourceLabel) {
  const normalizedRows = rawRows
    .map(normalizeRow)
    .filter((row) => row.studentid && row.studentname && row.school);

  if (!normalizedRows.length) {
    resetOutput();
    setStatus(
      "No valid rows found. Your file needs the columns studentid, studentname, and school."
    );
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
  setStatus(
    `Loaded ${normalizedRows.length} students from ${sourceLabel}. Day assignments were created automatically in batches of ${state.dayBatchSize}.`
  );
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
        (row, index) => `
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
            <span class="school-tag">${group.rows.length} students</span>
            <span class="school-tag">${group.dayCount} days</span>
          </div>
        </div>
        <button class="download-button" type="button">Download PDF</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Day</th>
              <th>Student ID</th>
              <th>Student Name</th>
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
  schoolListsContainer.textContent = "Upload a file to generate school rosters with day assignments.";
  downloadAllPdfsButton.disabled = true;
  updateStats();
}

function downloadAllPdfs() {
  state.groupedSchools.forEach((group, index) => {
    window.setTimeout(() => downloadSchoolPdf(group), index * 250);
  });
}

function downloadSchoolPdf(group) {
  const { jsPDF } = window.jspdf;
  const documentPdf = new jsPDF({
    orientation: "portrait",
    unit: "pt",
    format: "letter"
  });

  documentPdf.setFillColor(11, 110, 79);
  documentPdf.rect(0, 0, documentPdf.internal.pageSize.getWidth(), 74, "F");
  documentPdf.setTextColor(248, 246, 240);
  documentPdf.setFont("helvetica", "bold");
  documentPdf.setFontSize(24);
  documentPdf.text(group.school, 40, 45);

  documentPdf.setTextColor(93, 103, 103);
  documentPdf.setFont("helvetica", "normal");
  documentPdf.setFontSize(11);
  documentPdf.text(
    `Generated roster • ${group.rows.length} students • ${group.dayCount} days`,
    40,
    98
  );

  documentPdf.autoTable({
    startY: 118,
    head: [["Day", "Student ID", "Student Name"]],
    body: group.rows.map((row) => [row.dayLabel, row.studentid, row.studentname]),
    theme: "grid",
    headStyles: {
      fillColor: [23, 33, 33],
      textColor: [248, 246, 240]
    },
    styles: {
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
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}

function createSchoolGroup(school, rows) {
  const hasExplicitDays = rows.every((row) => row.dayLabel);
  const assignedRows = hasExplicitDays
    ? rows.map((row) => ({
        ...row,
        dayNumber: extractDayNumber(row.dayLabel)
      }))
    : shuffleRows(rows.map((row) => ({ ...row }))).map((row, index) => {
        const dayNumber = Math.floor(index / state.dayBatchSize) + 1;

        return {
          ...row,
          dayNumber,
          dayLabel: `Day ${dayNumber}`
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
      dayLabel: `Day ${Math.floor(index / DAY_BATCH_SIZE) + 1}`
    }));
  });
}

function normalizeDayLabel(value) {
  if (!value) {
    return "";
  }

  const dayNumber = extractDayNumber(value);
  return dayNumber ? `Day ${dayNumber}` : "";
}

function extractDayNumber(value) {
  const match = String(value).match(/\d+/);
  return match ? Number(match[0]) : Number.NaN;
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
