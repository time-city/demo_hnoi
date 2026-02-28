/* ============================================================
   AI Learning Performance Analytics – Application Logic
   ============================================================ */

// ─── Mock Data (fallback) ───────────────────────────────────
const defaultStudents = [
  {
    name: "Nguyen Van A", totalScore: 2, rank: 18, difference: -1.8, trend: "Declining", risk: "High",
    questions: [
      { id: "Q1", score: 0, classRate: 90 }, { id: "Q2", score: 0, classRate: 85 },
      { id: "Q3", score: 1, classRate: 60 }, { id: "Q4", score: 1, classRate: 65 },
      { id: "Q5", score: 0, classRate: 40 },
    ],
  },
  {
    name: "Tran Thi B", totalScore: 5, rank: 1, difference: 1.2, trend: "Improving", risk: "Low",
    questions: [
      { id: "Q1", score: 1, classRate: 90 }, { id: "Q2", score: 1, classRate: 85 },
      { id: "Q3", score: 1, classRate: 60 }, { id: "Q4", score: 1, classRate: 65 },
      { id: "Q5", score: 1, classRate: 40 },
    ],
  },
  {
    name: "Le Van C", totalScore: 3, rank: 12, difference: -0.8, trend: "Stable", risk: "Medium",
    questions: [
      { id: "Q1", score: 1, classRate: 90 }, { id: "Q2", score: 0, classRate: 85 },
      { id: "Q3", score: 1, classRate: 60 }, { id: "Q4", score: 0, classRate: 65 },
      { id: "Q5", score: 1, classRate: 40 },
    ],
  },
  {
    name: "Pham Thi D", totalScore: 4, rank: 5, difference: 0.2, trend: "Improving", risk: "Low",
    questions: [
      { id: "Q1", score: 1, classRate: 90 }, { id: "Q2", score: 1, classRate: 85 },
      { id: "Q3", score: 0, classRate: 60 }, { id: "Q4", score: 1, classRate: 65 },
      { id: "Q5", score: 1, classRate: 40 },
    ],
  },
  {
    name: "Hoang Van E", totalScore: 1, rank: 20, difference: -2.8, trend: "Declining", risk: "High",
    questions: [
      { id: "Q1", score: 0, classRate: 90 }, { id: "Q2", score: 0, classRate: 85 },
      { id: "Q3", score: 0, classRate: 60 }, { id: "Q4", score: 1, classRate: 65 },
      { id: "Q5", score: 0, classRate: 40 },
    ],
  },
  {
    name: "Vo Thi F", totalScore: 4, rank: 6, difference: 0.2, trend: "Stable", risk: "Low",
    questions: [
      { id: "Q1", score: 1, classRate: 90 }, { id: "Q2", score: 1, classRate: 85 },
      { id: "Q3", score: 1, classRate: 60 }, { id: "Q4", score: 1, classRate: 65 },
      { id: "Q5", score: 0, classRate: 40 },
    ],
  },
  {
    name: "Dang Van G", totalScore: 3, rank: 10, difference: -0.8, trend: "Declining", risk: "Medium",
    questions: [
      { id: "Q1", score: 1, classRate: 90 }, { id: "Q2", score: 1, classRate: 85 },
      { id: "Q3", score: 0, classRate: 60 }, { id: "Q4", score: 0, classRate: 65 },
      { id: "Q5", score: 1, classRate: 40 },
    ],
  },
];

// Active data – will be replaced by uploaded file data or fallback to mock
let students = [...defaultStudents];

// ─── DOM References ─────────────────────────────────────────
const pages = {
  upload: document.getElementById("page-upload"),
  dashboard: document.getElementById("page-dashboard"),
  detail: document.getElementById("page-detail"),
};

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const btnUpload = document.getElementById("btn-upload");
const fileInfo = document.getElementById("file-info");
const processingOverlay = document.getElementById("processing-overlay");
const progressFill = document.getElementById("progress-fill");
const headerNav = document.getElementById("header-nav");
const logoHome = document.getElementById("logo-home");
const btnTemplate = document.getElementById("btn-template");

let selectedFile = null;
let chartsCreated = false;
let questionChartInstance = null;
let distributionChartInstance = null;
let currentStudentIndex = null;

// ─── Navigation ─────────────────────────────────────────────
function showPage(name) {
  Object.values(pages).forEach((p) => p.classList.remove("active"));
  pages[name].classList.add("active");
  window.scrollTo({ top: 0, behavior: "smooth" });
  updateNav(name);
}

function updateNav(current) {
  headerNav.innerHTML = "";
  if (current === "upload") return;
  const items = [{ label: "Dashboard", page: "dashboard" }];
  items.forEach((item) => {
    const btn = document.createElement("button");
    btn.className = "nav-pill" + (current === item.page ? " active" : "");
    btn.textContent = item.label;
    btn.addEventListener("click", () => showPage(item.page));
    headerNav.appendChild(btn);
  });
}

logoHome.addEventListener("click", () => showPage("upload"));

// ─── Template Download (generates a real .xlsx file) ────────
btnTemplate.addEventListener("click", () => {
  const templateData = [
    ["Name", "Q1", "Q2", "Q3", "Q4", "Q5"],
    ["Nguyen Van A", 0, 0, 1, 1, 0],
    ["Tran Thi B", 1, 1, 1, 1, 1],
    ["Le Van C", 1, 0, 1, 0, 1],
  ];
  const ws = XLSX.utils.aoa_to_sheet(templateData);

  // Set column widths
  ws["!cols"] = [
    { wch: 20 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Students");
  XLSX.writeFile(wb, "student_template.xlsx");
});

// ─── Upload Logic ───────────────────────────────────────────
dropZone.addEventListener("click", () => fileInput.click());

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener("change", () => {
  if (fileInput.files.length) handleFile(fileInput.files[0]);
});

function handleFile(file) {
  const validExts = [".xlsx", ".xls", ".csv"];
  const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();

  if (!validExts.includes(ext)) {
    fileInfo.textContent = "❌ Invalid file type. Please upload .xlsx, .xls, or .csv";
    fileInfo.className = "file-info error";
    selectedFile = null;
    return;
  }

  selectedFile = file;
  fileInfo.textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  fileInfo.className = "file-info success";
}

btnUpload.addEventListener("click", () => {
  if (!selectedFile) {
    // Fallback to demo data
    fileInfo.textContent = "📄 No file selected — using demo data…";
    fileInfo.className = "file-info success";
    students = [...defaultStudents];
    startProcessing();
    return;
  }
  // Read the real file
  parseUploadedFile(selectedFile);
});

// ─── Real File Parsing with SheetJS ─────────────────────────
function parseUploadedFile(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: 0 });

      if (!json.length) {
        fileInfo.textContent = "❌ File is empty. Please check the template.";
        fileInfo.className = "file-info error";
        return;
      }

      // Detect question columns (Q1, Q2, ... Qn)
      const headers = Object.keys(json[0]);
      const nameCol = headers.find(
        (h) => h.toLowerCase() === "name" || h.toLowerCase() === "student" || h.toLowerCase() === "tên"
      );
      const questionCols = headers.filter((h) => /^q\d+$/i.test(h)).sort(
        (a, b) => parseInt(a.replace(/\D/g, "")) - parseInt(b.replace(/\D/g, ""))
      );

      if (!nameCol) {
        fileInfo.textContent = '❌ Could not find a "Name" column. Check your header row.';
        fileInfo.className = "file-info error";
        return;
      }

      if (questionCols.length === 0) {
        fileInfo.textContent = '❌ Could not find question columns (Q1, Q2, …). Check your header row.';
        fileInfo.className = "file-info error";
        return;
      }

      // Parse students
      const parsed = json.map((row) => {
        const qs = questionCols.map((qCol) => ({
          id: qCol.toUpperCase(),
          score: Number(row[qCol]) >= 1 ? 1 : 0,
          classRate: 0, // will compute below
        }));
        const total = qs.reduce((s, q) => s + q.score, 0);
        return {
          name: String(row[nameCol]),
          totalScore: total,
          rank: 0,
          difference: 0,
          trend: "Stable",
          risk: "Low",
          questions: qs,
        };
      });

      // Compute class-wide stats
      const totalStudents = parsed.length;
      const avg = parsed.reduce((s, st) => s + st.totalScore, 0) / totalStudents;

      // Correct rate per question
      questionCols.forEach((_, qi) => {
        const correctCount = parsed.filter((s) => s.questions[qi].score === 1).length;
        const rate = Math.round((correctCount / totalStudents) * 100);
        parsed.forEach((s) => (s.questions[qi].classRate = rate));
      });

      // Rank
      const sortedByScore = [...parsed].sort((a, b) => b.totalScore - a.totalScore);
      sortedByScore.forEach((st, i) => (st.rank = i + 1));

      // Difference & Risk/Trend
      parsed.forEach((st) => {
        st.difference = parseFloat((st.totalScore - avg).toFixed(1));
        if (st.totalScore <= avg - 1.5) { st.risk = "High"; st.trend = "Declining"; }
        else if (st.totalScore < avg) { st.risk = "Medium"; st.trend = "Stable"; }
        else if (st.totalScore >= avg + 1) { st.risk = "Low"; st.trend = "Improving"; }
        else { st.risk = "Low"; st.trend = "Stable"; }
      });

      students = parsed;
      fileInfo.textContent = `✅ Loaded ${students.length} students with ${questionCols.length} questions from file.`;
      fileInfo.className = "file-info success";
      startProcessing();
    } catch (err) {
      console.error(err);
      fileInfo.textContent = "❌ Error reading file: " + err.message;
      fileInfo.className = "file-info error";
    }
  };
  reader.readAsArrayBuffer(file);
}

// ─── Processing Animation ───────────────────────────────────
function startProcessing() {
  processingOverlay.classList.add("active");
  progressFill.style.width = "0%";
  let progress = 0;
  const interval = setInterval(() => {
    progress += Math.random() * 18 + 4;
    if (progress > 100) progress = 100;
    progressFill.style.width = progress + "%";
    if (progress >= 100) {
      clearInterval(interval);
      setTimeout(() => {
        processingOverlay.classList.remove("active");
        chartsCreated = false; // reset charts for new data
        showPage("dashboard");
        renderDashboard();
      }, 400);
    }
  }, 200);
}

// ─── Dashboard ──────────────────────────────────────────────
function renderDashboard() {
  const total = students.length;
  const avg = (students.reduce((s, st) => s + st.totalScore, 0) / total).toFixed(1);
  const high = Math.max(...students.map((s) => s.totalScore));
  const below = students.filter((s) => s.totalScore < parseFloat(avg)).length;

  animateValue("val-total", 0, total, 600);
  animateValue("val-avg", 0, parseFloat(avg), 600, true);
  animateValue("val-high", 0, high, 600);
  animateValue("val-below", 0, below, 600);

  renderStudentTable();
  renderQuestionAnalysis();

  // Destroy existing charts before recreating
  if (questionChartInstance) { questionChartInstance.destroy(); questionChartInstance = null; }
  if (distributionChartInstance) { distributionChartInstance.destroy(); distributionChartInstance = null; }
  renderCharts();
  chartsCreated = true;
}

function renderQuestionAnalysis() {
  const tbody = document.getElementById("question-analysis-tbody");
  tbody.innerHTML = "";
  const total = students.length;
  const numQ = students[0].questions.length;

  for (let qi = 0; qi < numQ; qi++) {
    const qId = students[0].questions[qi].id;
    const correctCount = students.filter((s) => s.questions[qi] && s.questions[qi].score === 1).length;
    const wrongCount = total - correctCount;
    const rate = Math.round((correctCount / total) * 100);

    let difficulty, diffClass;
    if (rate >= 80) { difficulty = "Easy"; diffClass = "comment-correct"; }
    else if (rate >= 50) { difficulty = "Medium"; diffClass = "comment-difficult"; }
    else { difficulty = "Hard"; diffClass = "comment-wrong"; }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${qId}</strong></td>
      <td>${correctCount}</td>
      <td>${wrongCount}</td>
      <td>${rate}%</td>
      <td><span class="${diffClass}">${difficulty}</span></td>
    `;
    tbody.appendChild(tr);
  }
}

function animateValue(id, start, end, duration, isFloat = false) {
  const el = document.getElementById(id);
  const range = end - start;
  const startTime = performance.now();
  function step(now) {
    const elapsed = now - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    const current = start + range * eased;
    el.textContent = isFloat ? current.toFixed(1) : Math.round(current);
    if (progress < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

function renderStudentTable() {
  const tbody = document.getElementById("student-tbody");
  tbody.innerHTML = "";

  const sorted = [...students].sort((a, b) => a.rank - b.rank);

  sorted.forEach((st, i) => {
    const riskClass =
      st.risk === "High" ? "risk-high" : st.risk === "Medium" ? "risk-medium" : "risk-low";

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><strong>${st.name}</strong></td>
      <td>${st.totalScore}/${st.questions.length}</td>
      <td>#${st.rank}</td>
      <td><span class="risk-badge ${riskClass}">${st.risk}</span></td>
      <td><button class="btn btn-sm btn-outline" data-index="${students.indexOf(st)}">View Report</button></td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll("[data-index]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = parseInt(btn.getAttribute("data-index"));
      showStudentDetail(idx);
    });
  });
}

function renderCharts() {
  // Detect number of questions dynamically
  const numQ = students[0].questions.length;
  const questionLabels = students[0].questions.map((q) => q.id);
  const correctCounts = questionLabels.map((_, qi) => {
    const correct = students.filter((s) => s.questions[qi] && s.questions[qi].score === 1).length;
    return ((correct / students.length) * 100).toFixed(1);
  });

  const barColors = [
    "rgba(46,134,193,0.7)", "rgba(93,173,226,0.7)", "rgba(72,201,176,0.7)",
    "rgba(118,215,196,0.7)", "rgba(142,68,173,0.5)", "rgba(52,152,219,0.6)",
    "rgba(26,188,156,0.6)", "rgba(155,89,182,0.6)", "rgba(46,204,113,0.6)",
    "rgba(241,196,15,0.6)",
  ];

  const ctxQ = document.getElementById("chart-questions").getContext("2d");
  questionChartInstance = new Chart(ctxQ, {
    type: "bar",
    data: {
      labels: questionLabels,
      datasets: [{
        label: "% Correct",
        data: correctCounts,
        backgroundColor: questionLabels.map((_, i) => barColors[i % barColors.length]),
        borderRadius: 8,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => ctx.parsed.y + "% Correct" } } },
      scales: {
        y: { beginAtZero: true, max: 100, ticks: { callback: (v) => v + "%" }, grid: { color: "rgba(0,0,0,0.04)" } },
        x: { grid: { display: false } },
      },
    },
  });

  // Distribution
  const maxScore = students[0].questions.length;
  const scoreBuckets = Array(maxScore + 1).fill(0);
  students.forEach((s) => { if (s.totalScore >= 0 && s.totalScore <= maxScore) scoreBuckets[s.totalScore]++; });

  const ctxD = document.getElementById("chart-distribution").getContext("2d");
  distributionChartInstance = new Chart(ctxD, {
    type: "bar",
    data: {
      labels: scoreBuckets.map((_, i) => String(i)),
      datasets: [{
        label: "Students",
        data: scoreBuckets,
        backgroundColor: "rgba(72,201,176,0.65)",
        borderRadius: 8,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, ticks: { stepSize: 1 }, grid: { color: "rgba(0,0,0,0.04)" } },
        x: { title: { display: true, text: "Score" }, grid: { display: false } },
      },
    },
  });
}

// ─── Student Detail ─────────────────────────────────────────
function showStudentDetail(index) {
  currentStudentIndex = index;
  const st = students[index];
  showPage("detail");

  document.getElementById("student-avatar").textContent = st.name.charAt(0);
  document.getElementById("detail-name").textContent = st.name;
  document.getElementById("detail-score").textContent = st.totalScore;
  document.getElementById("detail-rank").textContent = st.rank;

  const diffEl = document.getElementById("detail-diff");
  const diffPill = document.getElementById("detail-diff-pill");
  diffEl.textContent = (st.difference > 0 ? "+" : "") + st.difference.toFixed(1);
  diffPill.className = "pill " + (st.difference >= 0 ? "pill-diff-pos" : "pill-diff-neg");

  const trendEl = document.getElementById("detail-trend");
  const trendPill = document.getElementById("detail-trend-pill");
  trendEl.textContent = st.trend;
  trendPill.className = "pill pill-" + st.trend.toLowerCase();

  // Question table
  const qtbody = document.getElementById("question-tbody");
  qtbody.innerHTML = "";
  st.questions.forEach((q) => {
    let comment, commentClass;
    if (q.score === 1) { comment = "✅ Correct"; commentClass = "comment-correct"; }
    else if (q.classRate < 50) { comment = "⚠️ Difficult Question"; commentClass = "comment-difficult"; }
    else { comment = "❌ Wrong"; commentClass = "comment-wrong"; }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${q.id}</strong></td>
      <td>${q.score === 1 ? '<span class="comment-correct">1</span>' : '<span class="comment-wrong">0</span>'}</td>
      <td>${q.classRate}%</td>
      <td><span class="${commentClass}">${comment}</span></td>
    `;
    qtbody.appendChild(tr);
  });

  // AI Evaluation
  document.getElementById("ai-evaluation").innerHTML = generateAIEvaluation(st);

  // Re-trigger fade-in
  document.querySelectorAll("#page-detail .fade-in").forEach((el) => {
    el.style.animation = "none";
    el.offsetHeight;
    el.style.animation = "";
  });
}

function generateAIEvaluation(st) {
  const wrongQs = st.questions.filter((q) => q.score === 0).map((q) => q.id);
  const correctQs = st.questions.filter((q) => q.score === 1).map((q) => q.id);
  const hardWrong = st.questions.filter((q) => q.score === 0 && q.classRate < 50);
  const easyWrong = st.questions.filter((q) => q.score === 0 && q.classRate >= 70);

  let text = `<strong>${st.name}</strong> scored <strong>${st.totalScore}/${st.questions.length}</strong> (Rank #${st.rank}), which is `;
  text += st.difference >= 0
    ? `<strong>${Math.abs(st.difference).toFixed(1)} points above</strong> the class average. `
    : `<strong>${Math.abs(st.difference).toFixed(1)} points below</strong> the class average. `;

  if (correctQs.length === st.questions.length) {
    text += `Exceptional performance — answered all questions correctly. Recommend peer tutoring or advanced enrichment.`;
  } else {
    text += `Answered <strong>${correctQs.join(", ")}</strong> correctly`;
    if (wrongQs.length) text += ` but missed <strong>${wrongQs.join(", ")}</strong>`;
    text += ". ";

    if (easyWrong.length)
      text += `Notably, ${easyWrong.map((q) => q.id).join(" and ")} had high class correct rates (${easyWrong.map((q) => q.classRate + "%").join(", ")}), indicating gaps in fundamentals. `;
    if (hardWrong.length)
      text += `${hardWrong.map((q) => q.id).join(" and ")} had lower class-wide scores, so the material itself is challenging. `;

    if (st.trend === "Declining")
      text += `<br><br>⚠️ Performance trend is <strong>declining</strong>. Immediate intervention recommended.`;
    else if (st.trend === "Improving")
      text += `<br><br>📈 Student is <strong>improving</strong>. Continue current strategies.`;
    else
      text += `<br><br>📊 Performance is <strong>stable</strong>. Consider new study techniques.`;
  }
  return text;
}

// ─── Real PDF Download with jsPDF ───────────────────────────
document.getElementById("btn-download").addEventListener("click", () => {
  if (currentStudentIndex === null) return;
  const st = students[currentStudentIndex];
  generatePDF(st);
});

function generatePDF(st) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 16;
  let y = 20;

  // ── Helper functions
  const setFont = (style, size) => { doc.setFontSize(size); doc.setFont("helvetica", style); };
  const addLine = (text, x, yPos, opts = {}) => {
    const maxW = opts.maxWidth || pageW - margin * 2;
    const lines = doc.splitTextToSize(text, maxW);
    doc.text(lines, x, yPos);
    return yPos + lines.length * (opts.lineH || doc.getFontSize() * 0.4);
  };
  const checkPage = (needed) => { if (y + needed > 280) { doc.addPage(); y = 20; } };

  // ── Title
  doc.setFillColor(46, 134, 193);
  doc.rect(0, 0, pageW, 38, "F");
  doc.setTextColor(255, 255, 255);
  setFont("bold", 18);
  doc.text("Student Performance Report", margin, 16);
  setFont("normal", 10);
  doc.text("AI Learning Performance Analytics", margin, 24);
  doc.text("Generated: " + new Date().toLocaleDateString("en-GB"), margin, 30);
  y = 48;

  // ── Student Info
  doc.setTextColor(44, 62, 80);
  setFont("bold", 14);
  doc.text(st.name, margin, y);
  y += 8;

  setFont("normal", 10);
  doc.setTextColor(100, 100, 100);
  const avg = (students.reduce((s, s2) => s + s2.totalScore, 0) / students.length).toFixed(1);
  const infoLines = [
    `Score: ${st.totalScore}/${st.questions.length}    |    Rank: #${st.rank}    |    Class Average: ${avg}`,
    `Difference from Average: ${st.difference > 0 ? "+" : ""}${st.difference.toFixed(1)}    |    Trend: ${st.trend}    |    Risk: ${st.risk}`,
  ];
  infoLines.forEach((line) => { doc.text(line, margin, y); y += 5; });
  y += 6;

  // ── Divider
  doc.setDrawColor(200, 200, 200);
  doc.line(margin, y, pageW - margin, y);
  y += 8;

  // ── Question Analysis Table
  doc.setTextColor(44, 62, 80);
  setFont("bold", 12);
  doc.text("Question Analysis", margin, y);
  y += 8;

  // Table header
  const cols = [margin, margin + 30, margin + 60, margin + 90, margin + 120];
  const colHeaders = ["Question", "Score", "Class %", "Comment"];

  doc.setFillColor(244, 246, 247);
  doc.rect(margin, y - 4, pageW - margin * 2, 8, "F");
  setFont("bold", 9);
  doc.setTextColor(80, 80, 80);
  colHeaders.forEach((h, i) => doc.text(h, cols[i], y));
  y += 8;

  setFont("normal", 9);
  doc.setTextColor(44, 62, 80);
  st.questions.forEach((q) => {
    checkPage(8);
    let comment = q.score === 1 ? "Correct" : q.classRate < 50 ? "Difficult" : "Wrong";
    doc.text(q.id, cols[0], y);
    doc.text(String(q.score), cols[1], y);
    doc.text(q.classRate + "%", cols[2], y);

    if (q.score === 1) doc.setTextColor(39, 174, 96);
    else if (q.classRate < 50) doc.setTextColor(243, 156, 18);
    else doc.setTextColor(231, 76, 60);
    doc.text(comment, cols[3], y);
    doc.setTextColor(44, 62, 80);
    y += 7;
  });
  y += 6;

  // ── AI Evaluation
  checkPage(30);
  doc.setDrawColor(200, 200, 200);
  doc.line(margin, y, pageW - margin, y);
  y += 8;
  setFont("bold", 12);
  doc.text("AI Evaluation", margin, y);
  y += 8;
  setFont("normal", 9);
  doc.setTextColor(60, 60, 60);

  // Strip HTML from AI text
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = generateAIEvaluation(st);
  const plainText = tempDiv.textContent || tempDiv.innerText || "";
  y = addLine(plainText, margin, y, { lineH: 4.5 });
  y += 6;

  // ── Improvement Roadmap
  checkPage(50);
  doc.setTextColor(44, 62, 80);
  doc.setDrawColor(200, 200, 200);
  doc.line(margin, y, pageW - margin, y);
  y += 8;
  setFont("bold", 12);
  doc.text("Improvement Roadmap", margin, y);
  y += 10;

  const roadmap = [
    { week: "Week 1", title: "Foundation Review", desc: "Revisit core concepts where the student scored below class average." },
    { week: "Week 2", title: "Accuracy Training", desc: "Targeted exercises on frequently-missed question types." },
    { week: "Week 3", title: "Practice Tests", desc: "Simulate real exam conditions with timed practice tests." },
    { week: "Week 4", title: "Reassessment", desc: "Final evaluation to measure improvement and adjust learning path." },
  ];

  roadmap.forEach((item, i) => {
    checkPage(18);
    // Timeline dot
    doc.setFillColor(46, 134, 193);
    doc.circle(margin + 3, y, 2, "F");
    if (i < roadmap.length - 1) {
      doc.setDrawColor(46, 134, 193);
      doc.line(margin + 3, y + 2, margin + 3, y + 14);
    }

    setFont("bold", 10);
    doc.setTextColor(44, 62, 80);
    doc.text(`${item.week} – ${item.title}`, margin + 10, y + 1);
    setFont("normal", 9);
    doc.setTextColor(100, 100, 100);
    y = addLine(item.desc, margin + 10, y + 6, { maxWidth: pageW - margin * 2 - 10, lineH: 4 });
    y += 6;
  });

  // ── Footer
  const totalPages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(160, 160, 160);
    doc.text("AI Learning Performance Analytics — Confidential", margin, 290);
    doc.text(`Page ${i} of ${totalPages}`, pageW - margin - 20, 290);
  }

  // Save
  const safeName = st.name.replace(/\s+/g, "_");
  doc.save(`Report_${safeName}.pdf`);
}

// ─── Back Button ────────────────────────────────────────────
document.getElementById("btn-back").addEventListener("click", () => showPage("dashboard"));

// ─── Init ───────────────────────────────────────────────────
showPage("upload");
