/* ============================================================
   AI Learning Performance Analytics – Application Logic
   C1–C13 Standards-Based Analysis System
   ============================================================ */

// ─── STANDARDS ──────────────────────────────────────────────
const STANDARDS = ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13"];
const MAX_PER_STD = 4;
let TOTAL_MAX = STANDARDS.length * MAX_PER_STD; // 52

const SUGGESTIONS = [
  "Tăng cường luyện tập các bài tập suy luận",
  "Củng cố kiến thức nền tảng",
  "Cung cấp học liệu trực quan và ví dụ minh họa",
  "Áp dụng học nhóm và hỗ trợ đồng đẳng",
  "Giao bài tập có mục tiêu rõ ràng về chủ đề này",
  "Tích hợp bài tập giải quyết vấn đề thực tế",
  "Hướng dẫn từng bước theo trình tự logic",
  "Thực hiện kiểm tra đánh giá thường xuyên hơn",
  "Bổ sung bài tập đọc hiểu và phân tích văn bản",
  "Rèn tư duy phân tích qua các tình huống thực tế",
  "Sử dụng trò chơi học tập để tăng hứng thú",
  "Tổ chức ôn tập trước các kỳ kiểm tra",
  "Cung cấp tài liệu học tập theo khung đỡ (scaffolding)",
];

// ─── MOCK DATA (8+ students, realistic C1–C13 values 0–4) ──
const rawStudents = [
  { id: "S001", name: "Nguyen Van A", class: "10A1", scores: [4, 3, 2, 4, 3, 1, 4, 2, 3, 4, 3, 2, 4] },
  { id: "S002", name: "Tran Thi B", class: "10A1", scores: [1, 0, 1, 2, 1, 0, 1, 2, 0, 1, 1, 0, 1] },
  { id: "S003", name: "Le Van C", class: "10A1", scores: [3, 4, 3, 2, 4, 3, 2, 4, 3, 3, 4, 3, 2] },
  { id: "S004", name: "Pham Thi D", class: "10A2", scores: [2, 2, 3, 1, 2, 3, 2, 1, 2, 2, 3, 2, 2] },
  { id: "S005", name: "Hoang Van E", class: "10A2", scores: [0, 1, 0, 1, 0, 1, 0, 1, 1, 0, 0, 1, 0] },
  { id: "S006", name: "Vo Thi F", class: "10A2", scores: [4, 4, 3, 4, 4, 3, 4, 4, 3, 4, 4, 3, 4] },
  { id: "S007", name: "Dang Van G", class: "10A1", scores: [2, 1, 2, 3, 1, 2, 3, 2, 1, 2, 1, 2, 3] },
  { id: "S008", name: "Bui Thi H", class: "10A2", scores: [3, 2, 1, 0, 2, 1, 3, 0, 2, 1, 2, 1, 3] },
  { id: "S009", name: "Do Van I", class: "10A1", scores: [1, 1, 0, 2, 1, 0, 1, 1, 0, 2, 1, 0, 1] },
  { id: "S010", name: "Ngo Thi K", class: "10A2", scores: [4, 3, 4, 3, 4, 4, 3, 4, 4, 3, 4, 3, 4] },
];

// ─── MOCK SHEET 4 DATA (mảng nhiều lớp, khớp API thực tế) ───────────
const mockSheet4Data = [
  // ── Lớp 1: 10A1 ──────────────────────────────────────────
  {
    Sheets: 4,
    class_analysis: {
      class_name: "10A1-Mock",
      total_students: 30,
      avg_score: 6.2,
      max_score: 9.5,
      min_score: 2.0,
      below_avg_count: 8,
      below_avg_percent: 27,
      top5: [
        { id: "HS001", name: "Nguyễn Văn A", total_score: 9.5 },
        { id: "HS002", name: "Trần Thị B",   total_score: 9.0 },
        { id: "HS003", name: "Lê Văn C",     total_score: 8.8 },
        { id: "HS004", name: "Phạm Thị D",   total_score: 8.5 },
        { id: "HS005", name: "Hoàng Văn E",  total_score: 8.2 }
      ],
      bottom5: [
        { id: "HS030", name: "Vũ Thị K",    total_score: 2.0 },
        { id: "HS029", name: "Đặng Văn L",  total_score: 2.5 },
        { id: "HS028", name: "Bùi Thị M",   total_score: 3.0 },
        { id: "HS027", name: "Trịnh Văn N", total_score: 3.2 },
        { id: "HS026", name: "Lý Thị O",    total_score: 3.5 }
      ],
      hardest_question: { question: "q11", avg_score: 1.867 },
      easiest_question: { question: "q7",  avg_score: 3.933 },
      question_stats: [
        { question: "q11", avg_score: 1.867 },
        { question: "q13", avg_score: 2.067 },
        { question: "q8",  avg_score: 2.333 },
        { question: "q10", avg_score: 2.4   },
        { question: "q5",  avg_score: 2.6   },
        { question: "q3",  avg_score: 2.733 },
        { question: "q7",  avg_score: 3.933 }
      ]
    },
    class_histogram: {
      class_name: "10A1-Mock",
      ranges: ["0-2","2-4","4-6","6-8","8-10"],
      counts: [3, 5, 10, 8, 4],
      total_students: 30
    },
    class_commentary: {
      class_name: "10A1-Mock",
      level: "Trung bình",
      summary: "Lớp 10A1 ở mức TRUNG BÌNH, một số học sinh cần hỗ trợ thêm ở các câu khó.",
      insights: [
        "Điểm trung bình lớp: 6.2/10.",
        "Tỉ lệ học sinh dưới trung bình: 27% (8/30).",
        "Câu khó nhất: q11 (ĐTB: 1.867).",
        "Câu dễ nhất: q7 (ĐTB: 3.933)."
      ],
      focus_questions: [
        { question: "q11", avg_score: 1.867 },
        { question: "q13", avg_score: 2.067 },
        { question: "q8",  avg_score: 2.333 }
      ],
      actions: [
        "Ôn lại nền tảng theo chuyên đề; học sinh dưới TB làm bài theo mức độ dễ → trung bình trước.",
        "Kèm nhóm bottom 5 theo checklist lỗi sai; kiểm tra ngắn 10–15 phút mỗi buổi.",
        "Ưu tiên chữa và luyện tập trọng tâm các câu: q11, q13, q8."
      ]
    },
    sheet_class_report: {
      class_name: "10A1-Mock",
      summary: { level: "Trung bình", avg_score: 6.2, max_score: 9.5, min_score: 2.0, below_avg_percent: 27, total_students: 30 },
      insights: ["Điểm trung bình lớp: 6.2/10.", "Tỉ lệ học sinh dưới trung bình: 27% (8/30)."],
      histogram: { ranges: ["0-2","2-4","4-6","6-8","8-10"], counts: [3,5,10,8,4] },
      top_students:    [{ id: "HS001", name: "Nguyễn Văn A", total_score: 9.5 }, { id: "HS002", name: "Trần Thị B", total_score: 9.0 }],
      bottom_students: [{ id: "HS030", name: "Vũ Thị K",    total_score: 2.0 }, { id: "HS029", name: "Đặng Văn L", total_score: 2.5 }],
      hardest_question: { question: "q11", avg_score: 1.867 },
      easiest_question: { question: "q7",  avg_score: 3.933 },
      focus_questions: [{ question: "q11", avg_score: 1.867 }, { question: "q13", avg_score: 2.067 }, { question: "q8", avg_score: 2.333 }],
      action_plan: [
        "Ôn lại nền tảng theo chuyên đề; học sinh dưới TB làm bài theo mức độ dễ → trung bình trước.",
        "Kèm nhóm bottom 5 theo checklist lỗi sai; kiểm tra ngắn 10–15 phút mỗi buổi.",
        "Ưu tiên chữa và luyện tập trọng tâm các câu: q11, q13, q8."
      ]
    }
  },
  // ── Lớp 2: 10A2 ──────────────────────────────────────────
  {
    Sheets: 4,
    class_analysis: {
      class_name: "10A2-Mock",
      total_students: 28,
      avg_score: 7.8,
      max_score: 10.0,
      min_score: 4.5,
      below_avg_count: 4,
      below_avg_percent: 14,
      top5: [
        { id: "HS101", name: "Đinh Thị Lan",   total_score: 10.0 },
        { id: "HS102", name: "Ngô Văn Hùng",   total_score: 9.8  },
        { id: "HS103", name: "Phan Thị Mai",    total_score: 9.5  },
        { id: "HS104", name: "Trương Văn Đức",  total_score: 9.2  },
        { id: "HS105", name: "Lưu Thị Hoa",    total_score: 9.0  }
      ],
      bottom5: [
        { id: "HS128", name: "Cao Văn Bình",  total_score: 4.5 },
        { id: "HS127", name: "Mai Thị Cúc",   total_score: 5.0 },
        { id: "HS126", name: "Tô Văn Dũng",   total_score: 5.5 },
        { id: "HS125", name: "Kiều Thị Em",   total_score: 5.8 },
        { id: "HS124", name: "Lê Văn Phúc",   total_score: 6.0 }
      ],
      hardest_question: { question: "q9",  avg_score: 2.5  },
      easiest_question: { question: "q1",  avg_score: 4.2  },
      question_stats: [
        { question: "q9",  avg_score: 2.5  },
        { question: "q12", avg_score: 2.8  },
        { question: "q6",  avg_score: 3.1  },
        { question: "q4",  avg_score: 3.4  },
        { question: "q2",  avg_score: 3.7  },
        { question: "q1",  avg_score: 4.2  }
      ]
    },
    class_histogram: {
      class_name: "10A2-Mock",
      ranges: ["0-2","2-4","4-6","6-8","8-10"],
      counts: [0, 1, 3, 12, 12],
      total_students: 28
    },
    class_commentary: {
      class_name: "10A2-Mock",
      level: "Khá",
      summary: "Lớp 10A2 đạt mức KHÁ, đa số học sinh nắm vững kiến thức cơ bản.",
      insights: [
        "Điểm trung bình lớp: 7.8/10.",
        "Tỉ lệ học sinh dưới trung bình: 14% (4/28).",
        "Câu khó nhất: q9 (ĐTB: 2.5).",
        "Câu dễ nhất: q1 (ĐTB: 4.2)."
      ],
      focus_questions: [
        { question: "q9",  avg_score: 2.5 },
        { question: "q12", avg_score: 2.8 },
        { question: "q6",  avg_score: 3.1 }
      ],
      actions: [
        "Tiếp tục duy trì chất lượng; tổ chức học nhóm để học sinh khá hỗ trợ học sinh yếu.",
        "Tập trung ôn luyện câu q9 và q12 theo hướng thực hành nhiều hơn.",
        "Khuyến khích học sinh top đạt mức Xuất sắc trong kỳ tới."
      ]
    },
    sheet_class_report: {
      class_name: "10A2-Mock",
      summary: { level: "Khá", avg_score: 7.8, max_score: 10.0, min_score: 4.5, below_avg_percent: 14, total_students: 28 },
      insights: ["Điểm trung bình lớp: 7.8/10.", "Tỉ lệ học sinh dưới trung bình: 14% (4/28)."],
      histogram: { ranges: ["0-2","2-4","4-6","6-8","8-10"], counts: [0,1,3,12,12] },
      top_students:    [{ id: "HS101", name: "Đinh Thị Lan",  total_score: 10.0 }, { id: "HS102", name: "Ngô Văn Hùng", total_score: 9.8 }],
      bottom_students: [{ id: "HS128", name: "Cao Văn Bình", total_score: 4.5  }, { id: "HS127", name: "Mai Thị Cúc",  total_score: 5.0 }],
      hardest_question: { question: "q9", avg_score: 2.5 },
      easiest_question: { question: "q1", avg_score: 4.2 },
      focus_questions: [{ question: "q9", avg_score: 2.5 }, { question: "q12", avg_score: 2.8 }, { question: "q6", avg_score: 3.1 }],
      action_plan: [
        "Tiếp tục duy trì chất lượng; tổ chức học nhóm để học sinh khá hỗ trợ học sinh yếu.",
        "Tập trung ôn luyện câu q9 và q12 theo hướng thực hành nhiều hơn.",
        "Khuyến khích học sinh top đạt mức Xuất sắc trong kỳ tới."
      ]
    }
  },
  // ── Lớp 3: 10A3 ──────────────────────────────────────────
  {
    Sheets: 4,
    class_analysis: {
      class_name: "10A3-Mock",
      total_students: 32,
      avg_score: 4.1,
      max_score: 7.5,
      min_score: 0.5,
      below_avg_count: 18,
      below_avg_percent: 56,
      top5: [
        { id: "HS201", name: "Vương Thị Thu",   total_score: 7.5 },
        { id: "HS202", name: "Hà Văn Sơn",      total_score: 7.2 },
        { id: "HS203", name: "Đỗ Thị Ngọc",     total_score: 7.0 },
        { id: "HS204", name: "Chu Văn Tài",      total_score: 6.8 },
        { id: "HS205", name: "Lương Thị Uyên",  total_score: 6.5 }
      ],
      bottom5: [
        { id: "HS232", name: "Quách Văn Xây",  total_score: 0.5 },
        { id: "HS231", name: "Tiêu Thị Yên",   total_score: 1.0 },
        { id: "HS230", name: "Mạc Văn Zân",    total_score: 1.5 },
        { id: "HS229", name: "Âu Thị Ân",      total_score: 2.0 },
        { id: "HS228", name: "Biên Văn Bảo",   total_score: 2.3 }
      ],
      hardest_question: { question: "q13", avg_score: 0.8  },
      easiest_question: { question: "q2",  avg_score: 2.9  },
      question_stats: [
        { question: "q13", avg_score: 0.8  },
        { question: "q11", avg_score: 1.0  },
        { question: "q10", avg_score: 1.2  },
        { question: "q7",  avg_score: 1.5  },
        { question: "q4",  avg_score: 1.9  },
        { question: "q2",  avg_score: 2.9  }
      ]
    },
    class_histogram: {
      class_name: "10A3-Mock",
      ranges: ["0-2","2-4","4-6","6-8","8-10"],
      counts: [8, 10, 9, 5, 0],
      total_students: 32
    },
    class_commentary: {
      class_name: "10A3-Mock",
      level: "Yếu",
      summary: "Lớp 10A3 ở mức YẾU, hơn nửa học sinh dưới trung bình, cần can thiệp khẩn cấp.",
      insights: [
        "Điểm trung bình lớp: 4.1/10.",
        "Tỉ lệ học sinh dưới trung bình: 56% (18/32).",
        "Câu khó nhất: q13 (ĐTB: 0.8).",
        "Câu dễ nhất: q2 (ĐTB: 2.9)."
      ],
      focus_questions: [
        { question: "q13", avg_score: 0.8 },
        { question: "q11", avg_score: 1.0 },
        { question: "q10", avg_score: 1.2 }
      ],
      actions: [
        "Xây dựng kế hoạch phụ đạo khẩn cấp cho 18 học sinh dưới TB; ưu tiên củng cố kiến thức nền.",
        "Chia nhóm học tập: nhóm giỏi kèm nhóm yếu theo từng chủ đề cụ thể.",
        "Tăng tần suất kiểm tra mini (10 phút) để theo dõi tiến độ từng tuần.",
        "Liên hệ phụ huynh học sinh bottom 5 để phối hợp hỗ trợ tại nhà."
      ]
    },
    sheet_class_report: {
      class_name: "10A3-Mock",
      summary: { level: "Yếu", avg_score: 4.1, max_score: 7.5, min_score: 0.5, below_avg_percent: 56, total_students: 32 },
      insights: ["Điểm trung bình lớp: 4.1/10.", "Tỉ lệ học sinh dưới trung bình: 56% (18/32)."],
      histogram: { ranges: ["0-2","2-4","4-6","6-8","8-10"], counts: [8,10,9,5,0] },
      top_students:    [{ id: "HS201", name: "Vương Thị Thu", total_score: 7.5 }, { id: "HS202", name: "Hà Văn Sơn", total_score: 7.2 }],
      bottom_students: [{ id: "HS232", name: "Quách Văn Xây", total_score: 0.5 }, { id: "HS231", name: "Tiêu Thị Yên", total_score: 1.0 }],
      hardest_question: { question: "q13", avg_score: 0.8 },
      easiest_question: { question: "q2",  avg_score: 2.9 },
      focus_questions: [{ question: "q13", avg_score: 0.8 }, { question: "q11", avg_score: 1.0 }, { question: "q10", avg_score: 1.2 }],
      action_plan: [
        "Xây dựng kế hoạch phụ đạo khẩn cấp cho 18 học sinh dưới TB; ưu tiên củng cố kiến thức nền.",
        "Chia nhóm học tập: nhóm giỏi kèm nhóm yếu theo từng chủ đề cụ thể.",
        "Tăng tần suất kiểm tra mini (10 phút) để theo dõi tiến độ từng tuần.",
        "Liên hệ phụ huynh học sinh bottom 5 để phối hợp hỗ trợ tại nhà."
      ]
    }
  }
];

// ─── API RESPONSE DATA (from webhook) ──────────────────────
let apiSheet1Data = [];
let apiSheet2Data = [];
let apiSheet3Data = [];
let apiSheet4Data = [];

// ─── COMPUTED DATA ──────────────────────────────────────────
let students = [];

function classify(score) {
  if (score <= 1) return "fail";
  if (score <= 3) return "pass";
  return "excellent";
}

function computeStudents(raw) {
  return raw.map(r => {
    const totalScore = r.scores.reduce((a, b) => a + b, 0);
    const score10 = parseFloat(((totalScore / TOTAL_MAX) * 10).toFixed(2));
    const classifications = r.scores.map(classify);
    const excellentCount = classifications.filter(c => c === "excellent").length;
    const passCount = classifications.filter(c => c === "pass").length;
    const failCount = classifications.filter(c => c === "fail").length;
    const strongStds = STANDARDS.filter((_, i) => r.scores[i] === 4);
    const weakStds = STANDARDS.filter((_, i) => r.scores[i] <= 1);
    const improveStds = STANDARDS.filter((_, i) => r.scores[i] >= 2 && r.scores[i] <= 3);
    const passExcPct = ((passCount + excellentCount) / STANDARDS.length) * 100;
    let risk, riskClass;
    if (passExcPct >= 70) { risk = "Thấp"; riskClass = "risk-low"; }
    else if (passExcPct >= 40) { risk = "Trung bình"; riskClass = "risk-medium"; }
    else { risk = "Cao"; riskClass = "risk-high"; }

    let evaluation;
    if (risk === "Thấp") evaluation = "Học sinh thể hiện năng lực tốt ở hầu hết các chuẩn.";
    else if (risk === "Trung bình") evaluation = "Học sinh có kết quả ở mức trung bình, cần củng cố thêm một số lĩnh vực.";
    else evaluation = "Học sinh gặp khó khăn ở nhiều chuẩn, cần can thiệp ngay.";

    return {
      id: r.id, name: r.name, class: r.class, scores: r.scores,
      totalScore, score10, classifications,
      excellentCount, passCount, failCount,
      strongStds, weakStds, improveStds,
      risk, riskClass, evaluation
    };
  });
}

function computeStandardStats(studs) {
  const n = studs.length;
  return STANDARDS.map((std, i) => {
    const failCount = studs.filter(s => s.scores[i] <= 1).length;
    const passCount = studs.filter(s => s.scores[i] >= 2 && s.scores[i] <= 3).length;
    const excCount = studs.filter(s => s.scores[i] === 4).length;
    const failPct = parseFloat(((failCount / n) * 100).toFixed(1));
    const passPct = parseFloat(((passCount / n) * 100).toFixed(1));
    const excPct = parseFloat(((excCount / n) * 100).toFixed(1));
    const needImprovement = failCount + passCount;
    return { std, failPct, passPct, excPct, needImprovement, failCount, suggestion: SUGGESTIONS[i] || "Review and reinforce this area" };
  });
}

// ─── RISK BADGE HELPER ──────────────────────────────────────
function getRiskBadgeClass(risk) {
  if (!risk) return "risk-low";
  const r = String(risk).toLowerCase();
  if (r === "high" || r === "cao") return "risk-high";
  if (r === "medium" || r === "trung bình" || r === "trung binh") return "risk-medium";
  return "risk-low";
}

// ─── API: POST parsed data & receive response ────────────────
async function postDataToAPI(parsedData) {
  showSpinner(true);
  try {
    const res = await fetch("http://localhost:5678/webhook/demo", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(parsedData)
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);

    const raw = await res.text();
    console.log("📥 API raw response:", raw);

    let json;
    try { json = JSON.parse(raw); } catch(e) { throw new Error("Phản hồi không phải JSON hợp lệ: " + raw.slice(0, 200)); }

    // Chuẩn hóa: chấp nhận array hoặc object có key "data"/"result"/"output"
    let arr;
    if (Array.isArray(json)) {
      arr = json;
    } else if (json && Array.isArray(json.data)) {
      arr = json.data;
    } else if (json && Array.isArray(json.result)) {
      arr = json.result;
    } else if (json && Array.isArray(json.output)) {
      arr = json.output;
    } else if (json && typeof json === "object") {
      // Có thể là object đơn lẻ (1 sheet) hoặc object chứa sheets
      // Thử bọc thành array
      arr = [json];
    } else {
      throw new Error("Định dạng phản hồi không nhận dạng được: " + JSON.stringify(json).slice(0, 200));
    }

    console.log("✅ API parsed array:", arr);
    processAPIResponse(arr);
    return true;
  } catch (err) {
    console.error("API error:", err);
    fileInfo.textContent = `❌ Gửi dữ liệu thất bại: ${err.message}. Vui lòng thử lại.`;
    fileInfo.className = "file-info error";
    return false;
  } finally {
    showSpinner(false);
  }
}

function showSpinner(visible) {
  let spinner = document.getElementById("api-spinner");
  if (!spinner) {
    spinner = document.createElement("div");
    spinner.id = "api-spinner";
    spinner.style.cssText = "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.45);display:flex;align-items:center;justify-content:center;z-index:9999;";
    spinner.innerHTML = `<div style="background:#fff;border-radius:12px;padding:36px 48px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,0.18);">
      <div style="width:48px;height:48px;border:5px solid #e3e8f0;border-top-color:#2e86c1;border-radius:50%;animation:spin 0.8s linear infinite;margin:0 auto 18px;"></div>
      <p style="margin:0;font-size:15px;color:#2c3e50;font-weight:600;">Đang phân tích dữ liệu…</p>
    </div>
    <style>@keyframes spin{to{transform:rotate(360deg)}}</style>`;
    document.body.appendChild(spinner);
  }
  spinner.style.display = visible ? "flex" : "none";
}

function processAPIResponse(responseArray) {
  apiSheet1Data = [];
  apiSheet2Data = [];
  apiSheet3Data = [];
  // Giữ lại apiSheet4Data nếu đã có (mock/pre-set), chỉ reset nếu API thực sự trả Sheet4
  const prevSheet4 = apiSheet4Data.slice();
  apiSheet4Data = [];

  responseArray.forEach(item => {
    const sheetNum = item["Sheets"] ?? item["sheets"] ?? item["sheet"];
    const data = item["data"] ?? [];

    if (sheetNum == 1) apiSheet1Data = Array.isArray(data) ? data : [];
    else if (sheetNum == 2) apiSheet2Data = Array.isArray(data) ? data : [];
    else if (sheetNum == 3) apiSheet3Data = Array.isArray(data) ? data : [];
    else if (sheetNum == 4) {
      // Sheet 4: item.data là mảng các lớp — push từng lớp
      const classes = Array.isArray(item.data) ? item.data : [item];
      classes.forEach(cls => apiSheet4Data.push(cls));
      console.log("✅ Sheet 4 – số lớp thêm vào:", classes.length);
    }
  });

  // Fallback: nếu API không trả Sheet 4, dùng mock data
  if (!apiSheet4Data.length && prevSheet4.length) {
    apiSheet4Data = prevSheet4;
    console.log("⚠️ Dùng lại Sheet 4 cũ:", apiSheet4Data);
  }

  console.log("API Sheet 1:", apiSheet1Data);
  console.log("API Sheet 2:", apiSheet2Data);
  console.log("API Sheet 3:", apiSheet3Data);
  console.log("API Sheet 4 (Phân tích Lớp):", apiSheet4Data);
  if (apiSheet4Data.length) {
    console.group("📊 SHEET 4 – CLASS ANALYTICS");
    console.log("Số lớp:", apiSheet4Data.length);
    console.log("Các trường:", Object.keys(apiSheet4Data[0]));
    console.table(apiSheet4Data);
    console.groupEnd();
  }
}

// ─── DOM ────────────────────────────────────────────────────
const pages = { upload: document.getElementById("page-upload"), dashboard: document.getElementById("page-dashboard") };
const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const btnUpload = document.getElementById("btn-upload");
const fileInfo = document.getElementById("file-info");
const overlay = document.getElementById("processing-overlay");
const progressFill = document.getElementById("progress-fill");
const logoHome = document.getElementById("logo-home");

let selectedFile = null;
let failChartInst = null, distChartInst = null;
let sortCol = null, sortAsc = true;
let _lastParsedData = [];

// ─── NAVIGATION ─────────────────────────────────────────────
function showPage(name) {
  Object.values(pages).forEach(p => p.classList.remove("active"));
  pages[name].classList.add("active");
  window.scrollTo({ top: 0, behavior: "smooth" });
}
logoHome.addEventListener("click", () => showPage("upload"));

// ─── BACK BUTTON ────────────────────────────────────────────
document.getElementById("btn-back-upload").addEventListener("click", () => {
  // Reset state
  selectedFile = null;
  _lastParsedData = [];
  apiSheet1Data = [];
  apiSheet2Data = [];
  apiSheet3Data = [];
  apiSheet4Data = [];
  students = [];
  fileInfo.textContent = "";
  fileInfo.className = "file-info";
  fileInput.value = "";
  if (failChartInst) { failChartInst.destroy(); failChartInst = null; }
  if (distChartInst) { distChartInst.destroy(); distChartInst = null; }
  // Reset active tab back to first tab
  document.querySelectorAll(".sidebar-btn[data-tab]").forEach((b, i) => {
    b.classList.toggle("active", i === 0);
  });
  document.querySelectorAll(".tab-panel").forEach((p, i) => {
    p.classList.toggle("active", i === 0);
  });
  showPage("upload");
});

// Tab switching (only buttons with data-tab, not export buttons)
document.querySelectorAll(".sidebar-btn[data-tab]").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".sidebar-btn[data-tab]").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.tab).classList.add("active");
  });
});

// ─── UPLOAD ─────────────────────────────────────────────────
dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("drag-over"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", e => { e.preventDefault(); dropZone.classList.remove("drag-over"); if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]); });
fileInput.addEventListener("change", () => { if (fileInput.files.length) handleFile(fileInput.files[0]); });

function handleFile(file) {
  const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
  if (![".xlsx", ".xls", ".csv"].includes(ext)) {
    fileInfo.textContent = "❌ Định dạng file không hợp lệ. Vui lòng tải lên file .xlsx, .xls hoặc .csv";
    fileInfo.className = "file-info error"; selectedFile = null; return;
  }
  selectedFile = file;
  fileInfo.textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  fileInfo.className = "file-info success";
}

btnUpload.addEventListener("click", () => {
  if (!selectedFile) {
    fileInfo.textContent = "📄 Đang dùng dữ liệu mẫu…";
    fileInfo.className = "file-info success";
    students = computeStudents(rawStudents);
    _lastParsedData = rawStudents;
    // Inject mock Sheet 4 để hiển thị tab Phân Tích Lớp ngay cả khi API chưa trả Sheet 4
    apiSheet4Data = mockSheet4Data;
    postDataToAPI(rawStudents).then(ok => { if (ok) startProcessing(); });
    return;
  }
  parseFile(selectedFile);
});

// ─── TEMPLATE DOWNLOAD ─────────────────────────────────────
document.getElementById("btn-template").addEventListener("click", () => {
  const hdr = ["Student_ID", "Student_Name", "Class", ...STANDARDS];
  const rows = rawStudents.slice(0, 3).map(s => [s.id, s.name, s.class, ...s.scores]);
  const ws = XLSX.utils.aoa_to_sheet([hdr, ...rows]);
  ws["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 8 }, ...STANDARDS.map(() => ({ wch: 5 }))];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Students");
  XLSX.writeFile(wb, "student_template.xlsx");
});

// ─── FILE PARSE ─────────────────────────────────────────────
function parseFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: 0 });

      // ─── LOG: Dữ liệu thô từ file (raw JSON) ───
      console.group("📂 FILE INPUT DATA");
      console.log("📋 Sheet name:", wb.SheetNames[0]);
      console.log("📊 Số dòng dữ liệu:", json.length);
      console.log("🔑 Các cột (headers):", json.length ? Object.keys(json[0]) : []);
      console.log("📄 Raw JSON data:");
      console.table(json);
      console.log("📄 Raw JSON (full object):", JSON.stringify(json, null, 2));
      console.groupEnd();

      if (!json.length) { fileInfo.textContent = "❌ File không có dữ liệu."; fileInfo.className = "file-info error"; return; }

      // ─── Normalize key: chuẩn hóa \r\n → \n trong tất cả header (để hỗ trợ file Windows) ───
      const normalizeKey = k => String(k).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      const normalizedJson = json.map(row => {
        const newRow = {};
        Object.keys(row).forEach(k => { newRow[normalizeKey(k)] = row[k]; });
        return newRow;
      });
      const hdrs = Object.keys(normalizedJson[0]);

      // ─── Nhận diện cột ID, Tên, Lớp ───
      // Dùng toLowerCase() để tránh lỗi /i flag với Unicode tiếng Việt
      let idCol   = hdrs.find(h => { const l = h.toLowerCase(); return /m[aã]\s*hs|ms\s*hs|mshs|student.?id/.test(l) || /^id$/.test(l); });
      let nameCol = hdrs.find(h => { const l = h.toLowerCase(); return /h[oọ]\s*(v[aà])?\s*t[eê]n|student.?name|^name$/.test(l); });
      let classCol = hdrs.find(h => { const l = h.toLowerCase(); return /^l[oơ]p$|^class$/.test(l); });
      // "#" là cột STT, không phải ID — chỉ dùng làm fallback nếu không có idCol
      const sttCol = hdrs.find(h => h.trim() === "#");

      // ─── Nhận diện cột điểm câu hỏi: "Câu N..." (có thể multiline "Câu 1\n(Max:...") ───
      // Dùng toLowerCase() vì /i flag trong JS không case-fold ký tự Unicode như â/Â
      let stdCols = hdrs
        .filter(h => {
          const hl = h.toLowerCase();
          // Match: "câu N", "cau N" (có/không dấu), kể cả có \n ngay sau số
          return /^c[aâ]u\s*\d+/.test(hl) || /^c\d+$/.test(hl);
        })
        .sort((a, b) => {
          // Lấy số đầu tiên trong key để sort đúng thứ tự câu
          const na = parseInt(a.match(/\d+/)?.[0]) || 0;
          const nb = parseInt(b.match(/\d+/)?.[0]) || 0;
          return na - nb;
        });

      // Fallback: nếu không khớp pattern nào, lấy cột số trừ id/name/class/tổng/điểm
      if (stdCols.length === 0) {
        const skipPattern = /tổng|điểm quy|total|score|max/i;
        stdCols = hdrs.filter(h => h !== idCol && h !== nameCol && h !== classCol && !skipPattern.test(h));
      }

      // ─── LOG: Kết quả nhận diện cột ───
      console.group("🔍 COLUMN DETECTION");
      console.log("All headers:", hdrs);
      console.log("ID column:", idCol || "(auto-generate)");
      console.log("Name column:", nameCol || "(auto-generate)");
      console.log("Class column:", classCol || "(not found)");
      console.log("Standard columns:", stdCols);
      console.groupEnd();

      if (stdCols.length === 0) {
        fileInfo.textContent = "❌ Không tìm thấy cột dữ liệu điểm nào."; fileInfo.className = "file-info error"; return;
      }

      // ─── Lọc bỏ dòng không có dữ liệu số ───
      const dataRows = normalizedJson.filter(r =>
        stdCols.some(c => {
          const val = r[c];
          return val !== undefined && val !== null && val !== "" && !isNaN(Number(val));
        })
      );

      console.group("🗑️ FILTERED ROWS");
      console.log("Tổng dòng trong file:", json.length);
      console.log("Dòng bị bỏ:", json.length - dataRows.length);
      console.log("Dòng giữ lại:", dataRows.length);
      console.groupEnd();

      // ─── Update STANDARDS dynamically ───
      STANDARDS.length = 0;
      stdCols.forEach((c, idx) => {
        // Key dạng "Câu 1\n(Max: 4)\n..." → lấy số câu đầu tiên → "C1"
        const num = String(c).match(/\d+/);
        STANDARDS.push(num ? "C" + num[0] : `C${idx + 1}`);
      });
      TOTAL_MAX = STANDARDS.length * MAX_PER_STD;

      // ─── Build parsed array (dùng để tính toán nội bộ) ───
      const parsed = dataRows.map((r, i) => ({
        id:    idCol   && r[idCol]   ? String(r[idCol])
             : sttCol && r[sttCol]  ? String(r[sttCol])   // fallback: dùng # STT làm id
             : `S${String(i + 1).padStart(3, "0")}`,
        name:  nameCol && r[nameCol] ? String(r[nameCol]) : `Student ${i + 1}`,
        class: classCol && r[classCol] ? String(r[classCol]) : "—",
        scores: stdCols.map(c => Math.min(4, Math.max(0, parseInt(r[c]) || 0)))
      }));

      // ─── Build raw payload giữ nguyên key gốc (sau normalize) từ file để gửi API ───
      const rawPayload = dataRows.map(r => {
        const obj = {};
        hdrs.forEach(h => { obj[h] = r[h]; });
        return obj;
      });

      console.group("✅ PARSED STUDENT DATA");
      console.log("Số học sinh:", parsed.length);
      console.log("Số tiêu chuẩn:", stdCols.length);
      console.table(parsed);
      console.log("Raw payload gửi API:", JSON.stringify(rawPayload, null, 2));
      console.groupEnd();

      students = computeStudents(parsed);

      console.group("🧠 COMPUTED STUDENT DATA");
      console.table(students.map(s => ({
        id: s.id, name: s.name, class: s.class,
        totalScore: s.totalScore, score10: s.score10,
        excellent: s.excellentCount, pass: s.passCount, fail: s.failCount,
        risk: s.risk
      })));
      console.groupEnd();

      fileInfo.textContent = `✅ Đã tải ${students.length} học sinh, ${stdCols.length} chuẩn.`;
      fileInfo.className = "file-info success";
      _lastParsedData = rawPayload;
      postDataToAPI(rawPayload).then(ok => { if (ok) startProcessing(); });
    } catch (err) { fileInfo.textContent = "❌ Lỗi: " + err.message; fileInfo.className = "file-info error"; }
  };
  reader.readAsArrayBuffer(file);
}

// ─── PROCESSING ─────────────────────────────────────────────
function startProcessing() {
  overlay.classList.add("active");
  progressFill.style.width = "0%";
  let p = 0;
  const iv = setInterval(() => {
    p += Math.random() * 16 + 5;
    if (p > 100) p = 100;
    progressFill.style.width = p + "%";
    if (p >= 100) { clearInterval(iv); setTimeout(() => { overlay.classList.remove("active"); showPage("dashboard"); renderAll(); }, 350); }
  }, 180);
}

// ─── RENDER ALL ─────────────────────────────────────────────
function renderAll() {
  if (apiSheet1Data.length || apiSheet2Data.length || apiSheet3Data.length || apiSheet4Data.length) {
    renderDashboardSummary();
    renderSheet1Tab();
    renderSheet2Tab();
    renderSheet3Tab();
    renderSheet4Tab();
  } else {
    renderStandardsTab();
    renderStudentsTab();
    renderWeakTab();
  }
}

// ─── DASHBOARD SUMMARY (from API Sheet 2 & Sheet 3) ─────────
function renderDashboardSummary() {
  const totalStudents = apiSheet2Data.length;
  // Score (10) là key thực tế từ API — "Tổng điểm" bị null
  const scores = apiSheet2Data.map(r => parseFloat(
    r["Score (10)"] ?? r["Điểm thang 10"] ?? r["diem_10"] ?? r["Tổng điểm"] ?? 0
  )).filter(v => !isNaN(v) && v > 0);
  const avgScore = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : "—";
  const highest = scores.length ? Math.max(...scores) : "—";
  const totalWeak = apiSheet3Data.length;

  // Inject into standards tab summary cards (reuse existing card slot)
  const el = document.getElementById("std-summary-cards");
  if (el) {
    el.innerHTML = `
      <div class="summary-card"><div class="card-icon blue">🎓</div><div class="card-data"><span class="card-value">${totalStudents}</span><span class="card-label">Tổng Học Sinh</span></div></div>
      <div class="summary-card"><div class="card-icon green">📈</div><div class="card-data"><span class="card-value">${avgScore}</span><span class="card-label">Điểm Trung Bình</span></div></div>
      <div class="summary-card"><div class="card-icon purple">⭐</div><div class="card-data"><span class="card-value">${highest}</span><span class="card-label">Điểm Cao Nhất</span></div></div>
      <div class="summary-card"><div class="card-icon red">⚠️</div><div class="card-data"><span class="card-value">${totalWeak}</span><span class="card-label">Học Sinh Yếu</span></div></div>
    `;
  }
}

// ──────────────────────────────────────────────────────────────
// SHEET 1 – Standard Analysis (from API)
// ──────────────────────────────────────────────────────────────
function renderSheet1Tab() {
  // Update table header
  const table = document.getElementById("standards-table");
  if (table) {
    table.querySelector("thead tr").innerHTML = `
      <th>Mã Chuẩn</th>
      <th>Câu Hỏi</th>
      <th>Tỉ Lệ Không Đạt</th>
      <th>Tỉ Lệ Đạt</th>
      <th>Tỉ Lệ Xuất Sắc</th>
      <th>Tỉ Lệ Cần Cải Thiện</th>
      <th>Gợi Ý</th>
    `;
  }

  const tbody = document.getElementById("standards-tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  apiSheet1Data.forEach(row => {
    const failRate = parseFloat(row["tỉ_lệ_fail"] || row["ti_le_fail"] || row["Tỉ lệ Fail"] || 0);
    const isDanger = failRate >= 40;
    const tr = document.createElement("tr");
    if (isDanger) tr.classList.add("row-danger");
    tr.innerHTML = `
      <td><strong>${row["standard_code"] || row["Standard Code"] || "—"}</strong></td>
      <td>${row["câu_hỏi"] || row["cau_hoi"] || row["Câu hỏi"] || "—"}</td>
      <td class="${isDanger ? "cell-danger" : ""}">${failRate}%</td>
      <td>${row["tỉ_lệ_pass"] || row["ti_le_pass"] || row["Tỉ lệ Pass"] || "—"}%</td>
      <td>${row["tỉ_lệ_excellent"] || row["ti_le_excellent"] || row["Tỉ lệ Excellent"] || "—"}%</td>
      <td>${row["tỉ_lệ_cần_cải_thiện"] || row["ti_le_can_cai_thien"] || row["Tỉ lệ Cần Cải Thiện"] || "—"}%</td>
      <td class="eval-text">${row["suggestion"] || row["Suggestion"] || "—"}</td>
    `;
    tbody.appendChild(tr);
  });

  // Hide chart section or leave as is (charts are already rendered by renderStandardsTab fallback)
  if (failChartInst) { failChartInst.destroy(); failChartInst = null; }
  if (distChartInst) { distChartInst.destroy(); distChartInst = null; }

  // Draw simplified chart from API data
  const ctx1 = document.getElementById("chart-fail-rate");
  if (ctx1) {
    const labels = apiSheet1Data.map(r => r["standard_code"] || r["Standard Code"] || "");
    const failRates = apiSheet1Data.map(r => parseFloat(r["tỉ_lệ_fail"] || r["ti_le_fail"] || 0));
    failChartInst = new Chart(ctx1.getContext("2d"), {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label: "% Không đạt",
          data: failRates,
          backgroundColor: failRates.map(v => v >= 40 ? "rgba(231,76,60,0.75)" : "rgba(46,134,193,0.65)"),
          borderRadius: 6, borderSkipped: false,
        }]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        scales: {
          y: { beginAtZero: true, max: 100, ticks: { callback: v => v + "%" }, grid: { color: "rgba(0,0,0,0.04)" } },
          x: { grid: { display: false } }
        }
      }
    });
  }

  const ctx2 = document.getElementById("chart-score-dist");
  if (ctx2) {
    const excRates = apiSheet1Data.map(r => parseFloat(r["tỉ_lệ_excellent"] || r["ti_le_excellent"] || 0));
    const passRates = apiSheet1Data.map(r => parseFloat(r["tỉ_lệ_pass"] || r["ti_le_pass"] || 0));
    const failRates2 = apiSheet1Data.map(r => parseFloat(r["tỉ_lệ_fail"] || r["ti_le_fail"] || 0));
    const labels = apiSheet1Data.map(r => r["standard_code"] || r["Standard Code"] || "");
    distChartInst = new Chart(ctx2.getContext("2d"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          { label: "Xuất sắc", data: excRates, backgroundColor: "rgba(46,134,193,0.7)", borderRadius: 4, borderSkipped: false },
          { label: "Đạt", data: passRates, backgroundColor: "rgba(72,201,176,0.6)", borderRadius: 4, borderSkipped: false },
          { label: "Không đạt", data: failRates2, backgroundColor: "rgba(231,76,60,0.6)", borderRadius: 4, borderSkipped: false },
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: true } },
        scales: {
          y: { beginAtZero: true, max: 100, ticks: { callback: v => v + "%" }, grid: { color: "rgba(0,0,0,0.04)" } },
          x: { grid: { display: false } }
        }
      }
    });
  }
}

// ──────────────────────────────────────────────────────────────
// SHEET 2 – Student Analytics (from API)
// ──────────────────────────────────────────────────────────────
function renderSheet2Tab() {
  const n = apiSheet2Data.length;
  const scores = apiSheet2Data.map(r => parseFloat(
    r["Score (10)"] ?? r["Điểm thang 10"] ?? r["diem_10"] ?? 0
  )).filter(v => !isNaN(v) && v > 0);
  const avgScore = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : "—";
  const highRisk  = apiSheet2Data.filter(r => String(r["Risk Level"] || "").toUpperCase() === "HIGH").length;
  const lowRisk   = apiSheet2Data.filter(r => String(r["Risk Level"] || "").toUpperCase() === "LOW").length;

  const stuSummary = document.getElementById("stu-summary-cards");
  if (stuSummary) {
    stuSummary.innerHTML = `
      <div class="summary-card"><div class="card-icon blue">🎓</div><div class="card-data"><span class="card-value">${n}</span><span class="card-label">Tổng Học Sinh</span></div></div>
      <div class="summary-card"><div class="card-icon green">📈</div><div class="card-data"><span class="card-value">${avgScore}</span><span class="card-label">Điểm Trung Bình</span></div></div>
      <div class="summary-card"><div class="card-icon red">⚠️</div><div class="card-data"><span class="card-value">${highRisk}</span><span class="card-label">Nguy Cơ Cao</span></div></div>
      <div class="summary-card"><div class="card-icon purple">✅</div><div class="card-data"><span class="card-value">${lowRisk}</span><span class="card-label">Nguy Cơ Thấp</span></div></div>
    `;
  }

  // Update table header for Sheet 2
  const table = document.getElementById("students-table");
  if (table) {
    table.querySelector("thead tr").innerHTML = `
      <th>Mã HS</th>
      <th>Họ và Tên</th>
      <th>Tổng Điểm</th>
      <th>Điểm Thang 10</th>
      <th>Số Xuất Sắc</th>
      <th>Số Câu Đạt</th>
      <th>Số Câu Không Đạt</th>
      <th>Mức Nguy Cơ</th>
      <th>Đánh Giá Học Sinh</th>
      <th>Báo Cáo</th>
    `;
  }

  renderSheet2Table(apiSheet2Data);

  const searchInput = document.getElementById("student-search");
  if (searchInput) {
    searchInput.oninput = () => {
      const q = searchInput.value.toLowerCase();
      const filtered = apiSheet2Data.filter(r => {
        const name = String(r["Họ và tên"] || r["ho_va_ten"] || r["name"] || "").toLowerCase();
        const id = String(r["Mã HS"] || r["ma_hs"] || r["id"] || "").toLowerCase();
        return name.includes(q) || id.includes(q);
      });
      renderSheet2Table(filtered);
    };
  }
}

function renderSheet2Table(data) {
  const tbody = document.getElementById("students-tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  data.forEach((r, idx) => {
    const risk  = String(r["Risk Level"] || r["risk_level"] || "LOW");
    const riskClass = getRiskBadgeClass(risk);
    const name  = r["Họ và tên"] || r["ho_va_ten"] || r["name"] || "—";
    const maHS  = r["Mã HS"]     || r["ma_hs"]     || r["id"]   || "";
    const score10 = r["Score (10)"] ?? r["Điểm thang 10"] ?? r["diem_10"] ?? "—";
    const excel = r["Excellent"]  ?? r["Số Excellent"] ?? r["so_excellent"] ?? "—";
    const pass  = r["Pass"]       ?? r["Số câu Pass"]  ?? r["so_cau_pass"]  ?? "—";
    const fail  = r["Fail"]       ?? r["Số câu Fail"]  ?? r["so_cau_fail"]  ?? "—";
    // Chuẩn mạnh/yếu có thể là array hoặc string
    const manhArr = r["Chuẩn mạnh"] || r["chuan_manh"] || [];
    const yeuArr  = r["Chuẩn yếu"]  || r["chuan_yeu"]  || [];
    const manhStr = Array.isArray(manhArr) ? manhArr.join(", ") : String(manhArr);
    const yeuStr  = Array.isArray(yeuArr)  ? yeuArr.join(", ")  : String(yeuArr);
    const tr = document.createElement("tr");
    // Lưu index gốc trong apiSheet2Data để tìm đúng row khi bấm PDF
    tr.dataset.rowIdx = String(apiSheet2Data.indexOf(r));
    tr.innerHTML = `
      <td>${maHS || "—"}</td>
      <td><strong>${name}</strong></td>
      <td>${r["Tổng điểm"] ?? "—"}</td>
      <td><strong>${score10}</strong></td>
      <td>${excel}</td>
      <td>${pass}</td>
      <td>${fail}</td>
      <td><span class="risk-badge ${riskClass}">${risk}</span></td>
      <td class="eval-text">${r["Đánh giá học sinh"] || r["danh_gia_hoc_sinh"] || "—"}</td>
      <td><button class="btn btn-sm btn-outline btn-pdf" data-rowidx="${apiSheet2Data.indexOf(r)}">📄 PDF</button></td>
    `;
    tbody.appendChild(tr);
  });

  // Attach PDF handlers — dùng rowIdx để tránh nhầm khi Mã HS trùng "000"
  tbody.querySelectorAll(".btn-pdf").forEach(btn => {
    btn.addEventListener("click", () => {
      const idx = parseInt(btn.dataset.rowidx);
      const studentRow = apiSheet2Data[idx];
      if (studentRow) generateStudentPDFFromAPI(studentRow);
    });
  });
}

// ──────────────────────────────────────────────────────────────
// SHEET 3 – Weak Students Report (from API)
// ──────────────────────────────────────────────────────────────
function renderSheet3Tab() {
  const weak = apiSheet3Data;
  const highCount = weak.filter(r => String(r["Risk Level"] || "").toUpperCase() === "HIGH").length;
  const failCounts = weak.map(r => parseFloat(r["Số chuẩn Fail"] || r["so_chuan_fail"] || 0)).filter(v => !isNaN(v));
  const avgFail = failCounts.length ? (failCounts.reduce((a, b) => a + b, 0) / failCounts.length).toFixed(1) : "—";
  const totalScores = weak.map(r => parseFloat(r["Tổng điểm"] || 0)).filter(v => !isNaN(v) && v > 0);
  const avgWeak = totalScores.length ? (totalScores.reduce((a, b) => a + b, 0) / totalScores.length).toFixed(1) : "—";

  const weakSummary = document.getElementById("weak-summary-cards");
  if (weakSummary) {
    weakSummary.innerHTML = `
      <div class="summary-card"><div class="card-icon red">⚠️</div><div class="card-data"><span class="card-value">${weak.length}</span><span class="card-label">Học Sinh Yếu</span></div></div>
      <div class="summary-card"><div class="card-icon orange">🔥</div><div class="card-data"><span class="card-value">${highCount}</span><span class="card-label">Nguy Cơ Cao</span></div></div>
      <div class="summary-card"><div class="card-icon purple">📋</div><div class="card-data"><span class="card-value">${avgFail}</span><span class="card-label">TB Chuẩn Không Đạt</span></div></div>
    `;
  }

  // Update table header for Sheet 3
  const table = document.getElementById("weak-table");
  if (table) {
    table.querySelector("thead tr").innerHTML = `
      <th>STT</th>
      <th>Họ và Tên</th>
      <th>Điểm /10</th>
      <th>Số Chuẩn Không Đạt</th>
      <th>Chuẩn Yếu</th>
      <th>Mức Nguy Cơ</th>
      <th>Gợi Ý Can Thiệp</th>
    `;
  }

  const tbody = document.getElementById("weak-tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  // Tạo map tên → Score(10) từ Sheet2 để hiển thị tên & điểm trong Sheet3
  // Sheet3 chỉ có Mã HS (luôn "000"), dùng index tương ứng với Sheet2 học sinh yếu
  const weakStudentsFromSheet2 = apiSheet2Data.filter(r =>
    String(r["Risk Level"] || "").toUpperCase() === "HIGH" ||
    String(r["Risk Level"] || "").toUpperCase() === "MEDIUM"
  );

  weak.forEach((r, i) => {
    const risk = String(r["Risk Level"] || "").toUpperCase() || "MEDIUM";
    const riskClass = getRiskBadgeClass(risk);
    // Lấy tên từ Sheet2 theo index (Sheet3 không có tên)
    const s2Match = weakStudentsFromSheet2[i] || apiSheet2Data[i] || {};
    const name    = s2Match["Họ và tên"] || s2Match["name"] || "—";
    const score10 = s2Match["Score (10)"] ?? s2Match["Điểm thang 10"] ?? "—";
    // Chuẩn yếu có thể là array
    const yeuArr  = r["Chuẩn yếu"] || r["chuan_yeu"] || [];
    const yeuStr  = Array.isArray(yeuArr) ? yeuArr.join(", ") : String(yeuArr || "—");
    const tr = document.createElement("tr");
    if (risk === "HIGH") tr.classList.add("row-danger");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><strong>${name}</strong></td>
      <td><strong>${score10}</strong></td>
      <td>${r["Số chuẩn Fail"] || r["so_chuan_fail"] || "—"}</td>
      <td class="std-list weak">${yeuStr}</td>
      <td><span class="risk-badge ${riskClass}">${risk}</span></td>
      <td class="eval-text">${r["Gợi ý can thiệp"] || r["goi_y_can_thiep"] || "—"}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ──────────────────────────────────────────────────────────────
// TAB 1 – Standard Analysis
// ──────────────────────────────────────────────────────────────
function renderStandardsTab() {
  const stats = computeStandardStats(students);
  const n = students.length;
  const avgTotal = (students.reduce((s, st) => s + st.totalScore, 0) / n).toFixed(1);
  const highFail = stats.filter(s => s.failPct > 40).length;
  const avgExc = (stats.reduce((s, st) => s + st.excPct, 0) / stats.length).toFixed(1);

  // Summary cards
  document.getElementById("std-summary-cards").innerHTML = `
    <div class="summary-card"><div class="card-icon blue">📋</div><div class="card-data"><span class="card-value">${n}</span><span class="card-label">Tổng Học Sinh</span></div></div>
    <div class="summary-card"><div class="card-icon green">📊</div><div class="card-data"><span class="card-value">${avgTotal}</span><span class="card-label">Điểm TB /52</span></div></div>
    <div class="summary-card"><div class="card-icon red">🚨</div><div class="card-data"><span class="card-value">${highFail}</span><span class="card-label">Chuẩn >40% Không Đạt</span></div></div>
    <div class="summary-card"><div class="card-icon purple">⭐</div><div class="card-data"><span class="card-value">${avgExc}%</span><span class="card-label">Tỉ Lệ Xuất Sắc TB</span></div></div>
  `;

  // Table
  const tbody = document.getElementById("standards-tbody");
  tbody.innerHTML = "";
  stats.forEach(s => {
    const danger = s.failPct > 40;
    const tr = document.createElement("tr");
    if (danger) tr.classList.add("row-danger");
    tr.innerHTML = `
      <td><strong>${s.std}</strong></td>
      <td class="${danger ? 'cell-danger' : ''}">${s.failPct}%</td>
      <td>${s.passPct}%</td>
      <td>${s.excPct}%</td>
      <td>${s.needImprovement} / ${n}</td>
      <td class="eval-text">${s.suggestion}</td>
    `;
    tbody.appendChild(tr);
  });

  // Charts
  if (failChartInst) failChartInst.destroy();
  if (distChartInst) distChartInst.destroy();

  const ctx1 = document.getElementById("chart-fail-rate").getContext("2d");
  failChartInst = new Chart(ctx1, {
    type: "bar",
    data: {
      labels: stats.map(s => s.std),
      datasets: [{
        label: "% Không đạt",
        data: stats.map(s => s.failPct),
        backgroundColor: stats.map(s => s.failPct > 40 ? "rgba(231,76,60,0.75)" : "rgba(46,134,193,0.65)"),
        borderRadius: 6, borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        annotation: undefined,
      },
      scales: {
        y: { beginAtZero: true, max: 100, ticks: { callback: v => v + "%" }, grid: { color: "rgba(0,0,0,0.04)" } },
        x: { grid: { display: false } }
      }
    }
  });

  // Score distribution (overall)
  const buckets = [0, 0, 0, 0, 0]; // 0,1,2,3,4
  students.forEach(st => st.scores.forEach(sc => buckets[sc]++));
  const ctx2 = document.getElementById("chart-score-dist").getContext("2d");
  distChartInst = new Chart(ctx2, {
    type: "bar",
    data: {
      labels: ["0 (Không đạt)", "1 (Không đạt)", "2 (Đạt)", "3 (Đạt)", "4 (Xuất sắc)"],
      datasets: [{
        label: "Số lượng",
        data: buckets,
        backgroundColor: ["rgba(231,76,60,0.6)", "rgba(231,76,60,0.4)", "rgba(243,156,18,0.5)", "rgba(72,201,176,0.6)", "rgba(46,134,193,0.7)"],
        borderRadius: 6, borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, ticks: { stepSize: 5 }, grid: { color: "rgba(0,0,0,0.04)" } },
        x: { grid: { display: false } }
      }
    }
  });
}

// ──────────────────────────────────────────────────────────────
// TAB 2 – Student Analytics
// ──────────────────────────────────────────────────────────────
function renderStudentsTab() {
  const n = students.length;
  const avgScore10 = (students.reduce((s, st) => s + st.score10, 0) / n).toFixed(2);
  const highRisk = students.filter(s => s.risk === "High").length;
  const lowRisk = students.filter(s => s.risk === "Low").length;

  document.getElementById("stu-summary-cards").innerHTML = `
    <div class="summary-card"><div class="card-icon blue">🎓</div><div class="card-data"><span class="card-value">${n}</span><span class="card-label">Tổng Học Sinh</span></div></div>
    <div class="summary-card"><div class="card-icon green">📈</div><div class="card-data"><span class="card-value">${avgScore10}</span><span class="card-label">Điểm TB /10</span></div></div>
    <div class="summary-card"><div class="card-icon red">⚠️</div><div class="card-data"><span class="card-value">${highRisk}</span><span class="card-label">Nguy Cơ Cao</span></div></div>
    <div class="summary-card"><div class="card-icon purple">✅</div><div class="card-data"><span class="card-value">${lowRisk}</span><span class="card-label">Nguy Cơ Thấp</span></div></div>
  `;

  renderStudentsTable(students);

  // Search
  const searchInput = document.getElementById("student-search");
  searchInput.oninput = () => {
    const q = searchInput.value.toLowerCase();
    const filtered = students.filter(s => s.name.toLowerCase().includes(q) || s.id.toLowerCase().includes(q));
    renderStudentsTable(filtered);
  };

  // Sorting
  document.querySelectorAll("#students-table th.sortable").forEach(th => {
    th.onclick = () => {
      const col = th.dataset.col;
      if (sortCol === col) sortAsc = !sortAsc; else { sortCol = col; sortAsc = true; }
      const sorted = [...students].sort((a, b) => {
        let va, vb;
        if (col === "id") { va = a.id; vb = b.id; }
        else if (col === "name") { va = a.name; vb = b.name; }
        else if (col === "total") { va = a.totalScore; vb = b.totalScore; }
        else if (col === "score10") { va = a.score10; vb = b.score10; }
        else if (col === "risk") { const order = { High: 0, Medium: 1, Low: 2 }; va = order[a.risk]; vb = order[b.risk]; }
        if (typeof va === "string") return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
        return sortAsc ? va - vb : vb - va;
      });
      renderStudentsTable(sorted);
    };
  });
}

function renderStudentsTable(data) {
  const tbody = document.getElementById("students-tbody");
  tbody.innerHTML = "";
  data.forEach(s => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${s.id}</td>
      <td><strong>${s.name}</strong></td>
      <td>${s.totalScore}</td>
      <td>${s.score10}</td>
      <td>${s.excellentCount}</td>
      <td>${s.passCount}</td>
      <td>${s.failCount}</td>
      <td class="std-list strong">${s.strongStds.join(", ") || "—"}</td>
      <td class="std-list weak">${s.weakStds.join(", ") || "—"}</td>
      <td class="std-list improve">${s.improveStds.join(", ") || "—"}</td>
      <td><span class="risk-badge ${s.riskClass}">${s.risk}</span></td>
      <td class="eval-text">${s.evaluation}</td>
      <td><button class="btn btn-sm btn-outline btn-pdf" data-sid="${s.id}">📄 PDF</button></td>
    `;
    tbody.appendChild(tr);
  });

  // Attach PDF handlers
  tbody.querySelectorAll(".btn-pdf").forEach(btn => {
    btn.addEventListener("click", () => {
      const st = students.find(s => s.id === btn.dataset.sid);
      if (st) generateStudentPDF(st);
    });
  });
}

// ──────────────────────────────────────────────────────────────
// TAB 3 – Weak Students
// ──────────────────────────────────────────────────────────────
function renderWeakTab() {
  const weak = students.filter(s => s.score10 < 5 || s.failCount >= 4);
  const highCount = weak.filter(s => s.risk === "High").length;

  document.getElementById("weak-summary-cards").innerHTML = `
    <div class="summary-card"><div class="card-icon red">⚠️</div><div class="card-data"><span class="card-value">${weak.length}</span><span class="card-label">Học Sinh Yếu</span></div></div>
    <div class="summary-card"><div class="card-icon orange">🔥</div><div class="card-data"><span class="card-value">${highCount}</span><span class="card-label">Nguy Cơ Cao</span></div></div>
    <div class="summary-card"><div class="card-icon blue">📊</div><div class="card-data"><span class="card-value">${weak.length > 0 ? (weak.reduce((s, st) => s + st.score10, 0) / weak.length).toFixed(2) : "—"}</span><span class="card-label">Điểm TB /10</span></div></div>
    <div class="summary-card"><div class="card-icon purple">📋</div><div class="card-data"><span class="card-value">${weak.length > 0 ? (weak.reduce((s, st) => s + st.failCount, 0) / weak.length).toFixed(1) : "—"}</span><span class="card-label">TB Chuẩn Không Đạt</span></div></div>
  `;

  const tbody = document.getElementById("weak-tbody");
  tbody.innerHTML = "";

  weak.forEach(s => {
    let intervention;
    if (s.risk === "High") intervention = "Provide foundational review sessions and one-on-one tutoring.";
    else if (s.failCount >= 6) intervention = "Assign guided practice tasks focusing on core standards.";
    else intervention = "Reinforce weak areas with additional exercises and peer study groups.";

    const tr = document.createElement("tr");
    if (s.risk === "High") tr.classList.add("row-danger");
    tr.innerHTML = `
      <td>${s.id}</td>
      <td><strong>${s.name}</strong></td>
      <td>${s.totalScore}</td>
      <td>${s.score10}</td>
      <td class="std-list weak">${s.weakStds.join(", ") || "—"}</td>
      <td>${s.failCount}</td>
      <td><span class="risk-badge ${s.riskClass}">${s.risk}</span></td>
      <td class="eval-text">${intervention}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ══════════════════════════════════════════════════════════════
// SHEET 4 – Class Analytics (from API)
// ══════════════════════════════════════════════════════════════
// Cấu trúc thực tế từ API – MỖI KEY LÀ OBJECT ĐƠN (không phải array):
// {
//   Sheets: 4,
//   class_analysis:    { class_name, total_students, avg_score, max_score, min_score,
//                        below_avg_count, below_avg_percent, top5[], bottom5[],
//                        hardest_question, easiest_question, question_stats[] }
//   class_histogram:   { class_name, bins[], ranges[], counts[], total_students }
//   class_commentary:  { class_name, level, summary, insights[], focus_questions[], actions[] }
//   sheet_class_report:{ class_name, summary{}, insights[], histogram{}, top_students[],
//                        bottom_students[], hardest_question, easiest_question,
//                        focus_questions[], action_plan[] }
// }

// ── Global helpers cho Sheet4 ─────────────────────────────
function scoreColor(v) {
  const n = parseFloat(v) || 0;
  const pct = (n / 10) * 100;
  if (pct < 40) return { color: "#C0392B", bg: "#FADBD8" };
  if (pct < 65) return { color: "#7E5109", bg: "#FEF9E7" };
  return { color: "#1D6A39", bg: "#D5F5E3" };
}
function levelBadge(lv) {
  const map = {
    "Xuất sắc": {color:"#1D6A39",bg:"#D5F5E3"}, "Tốt": {color:"#1D6A39",bg:"#D5F5E3"},
    "Khá":      {color:"#7E5109",bg:"#FEF9E7"}, "Trung bình": {color:"#7E5109",bg:"#FEF9E7"},
    "Yếu":      {color:"#C0392B",bg:"#FADBD8"}, "Kém": {color:"#C0392B",bg:"#FADBD8"},
  };
  const c = map[lv] || {color:"#555",bg:"#eee"};
  return `<span style="padding:3px 14px;border-radius:12px;font-weight:700;font-size:13px;color:${c.color};background:${c.bg};">${lv||"—"}</span>`;
}

function renderSheet4Tab() {
  const container = document.getElementById("tab-class");
  if (!container) return;

  const pageTitle = container.querySelector(".page-title");
  container.innerHTML = "";
  if (pageTitle) container.appendChild(pageTitle);

  if (!apiSheet4Data.length) {
    container.insertAdjacentHTML("beforeend", `
      <div class="summary-cards fade-in" id="class-summary-cards">
        <p style="color:#888;padding:16px;">⏳ Chưa có dữ liệu phân tích lớp từ API.</p>
      </div>`);
    return;
  }

  // ── Thanh tìm kiếm + bộ đếm ───────────────────────────────
  container.insertAdjacentHTML("beforeend", `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:20px;flex-wrap:wrap;">
      <div style="position:relative;flex:1;min-width:220px;max-width:400px;">
        <span style="position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:15px;pointer-events:none;">🔍</span>
        <input id="sheet4-search" type="text" placeholder="Tìm kiếm theo tên lớp…"
          style="width:100%;box-sizing:border-box;padding:10px 14px 10px 38px;border:1.5px solid #d0d7de;border-radius:8px;font-size:14px;outline:none;transition:border .2s;"
          oninput="filterSheet4Classes()"
          onfocus="this.style.borderColor='#2E86C1'"
          onblur="this.style.borderColor='#d0d7de'">
      </div>
      <span id="sheet4-count" style="font-size:13px;color:#666;white-space:nowrap;">
        ${apiSheet4Data.length} lớp
      </span>
      <button onclick="document.getElementById('sheet4-search').value='';filterSheet4Classes();"
        style="padding:8px 16px;border:1.5px solid #d0d7de;border-radius:8px;background:#fff;font-size:13px;color:#555;cursor:pointer;">
        ✕ Xóa
      </button>
    </div>
    <div id="sheet4-classes-wrapper"></div>
  `);

  // Render tất cả lớp vào wrapper
  renderSheet4Classes(apiSheet4Data);
}

// ── Lọc lớp theo từ khóa ──────────────────────────────────
function filterSheet4Classes() {
  const keyword = (document.getElementById("sheet4-search")?.value || "").trim().toLowerCase();
  const filtered = keyword
    ? apiSheet4Data.filter(item => {
        const name = (item.class_name || "").toLowerCase();
        return name.includes(keyword);
      })
    : apiSheet4Data;
  const countEl = document.getElementById("sheet4-count");
  if (countEl) countEl.textContent = `${filtered.length} / ${apiSheet4Data.length} lớp`;
  renderSheet4Classes(filtered);
}

// ── Render danh sách lớp ──────────────────────────────────
// ── Biến lưu lớp đang chọn ───────────────────────────────
let _sheet4ActiveIdx = 0;

function renderSheet4Classes(dataArr) {
  const wrapper = document.getElementById("sheet4-classes-wrapper");
  if (!wrapper) return;
  wrapper.innerHTML = "";

  if (!dataArr.length) {
    wrapper.innerHTML = `<div style="padding:32px;text-align:center;color:#aaa;font-size:15px;">🔍 Không tìm thấy lớp phù hợp.</div>`;
    return;
  }

  // ── Hàng nút lớp ─────────────────────────────────────────
  const btnBar = document.createElement("div");
  btnBar.id = "sheet4-btn-bar";
  btnBar.style.cssText = "display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px;";

  dataArr.forEach((s4, idx) => {
    const smr   = s4.summary || {};
    const name  = s4.class_name || `Lớp ${idx+1}`;
    const level = smr.level || "";
    const levelColors = {
      "Xuất sắc":"#1D6A39","Tốt":"#1D6A39","Khá":"#7E5109",
      "Trung bình":"#7E5109","Yếu":"#C0392B","Kém":"#C0392B"
    };
    const lc = levelColors[level] || "#2E86C1";

    const btn = document.createElement("button");
    btn.id = `sheet4-btn-${idx}`;
    btn.textContent = name;
    btn.style.cssText = `padding:8px 16px;border-radius:20px;border:2px solid ${lc};background:#fff;color:${lc};font-size:13px;font-weight:600;cursor:pointer;transition:all .18s;`;
    btn.onmouseenter = () => { if (btn.dataset.active !== "1") { btn.style.background = lc + "18"; } };
    btn.onmouseleave = () => { if (btn.dataset.active !== "1") { btn.style.background = "#fff"; } };
    btn.onclick = () => showSheet4Class(dataArr, idx);
    btnBar.appendChild(btn);
  });
  wrapper.appendChild(btnBar);

  // ── Vùng chi tiết lớp ────────────────────────────────────
  const detailPanel = document.createElement("div");
  detailPanel.id = "sheet4-detail-panel";
  wrapper.appendChild(detailPanel);

  // Hiện lớp đầu tiên mặc định
  showSheet4Class(dataArr, 0);
}

// ── Hiển thị chi tiết 1 lớp ──────────────────────────────
function showSheet4Class(dataArr, idx) {
  _sheet4ActiveIdx = idx;

  // Cập nhật trạng thái nút
  dataArr.forEach((_, i) => {
    const btn = document.getElementById(`sheet4-btn-${i}`);
    if (!btn) return;
    const smr   = dataArr[i].summary || {};
    const level = smr.level || "";
    const lc = { "Xuất sắc":"#1D6A39","Tốt":"#1D6A39","Khá":"#7E5109","Trung bình":"#7E5109","Yếu":"#C0392B","Kém":"#C0392B" }[level] || "#2E86C1";
    if (i === idx) {
      btn.dataset.active = "1";
      btn.style.background = lc;
      btn.style.color = "#fff";
      btn.style.boxShadow = `0 3px 10px ${lc}55`;
    } else {
      btn.dataset.active = "0";
      btn.style.background = "#fff";
      btn.style.color = lc;
      btn.style.boxShadow = "none";
    }
  });

  const s4      = dataArr[idx];
  const smr     = s4.summary   || {};
  const hist    = s4.histogram || {};
  const className     = s4.class_name            || `Lớp ${idx+1}`;
  const totalStudents = smr.total_students       ?? 0;
  const avgScore      = smr.avg_score            ?? "—";
  const belowPct      = smr.below_avg_percent    ?? "—";
  const level         = smr.level                || "—";
  const maxScore      = smr.max_score            ?? "—";
  const minScore      = smr.min_score            ?? "—";
  const hardestQ      = s4.hardest_question      || {};
  const easiestQ      = s4.easiest_question      || {};
  const insights      = s4.insights              || [];
  const actions       = s4.action_plan           || [];

  const panel = document.getElementById("sheet4-detail-panel");
  if (!panel) return;

  panel.innerHTML = `
    <div class="fade-in" style="border:1.5px solid #e0e7ef;border-radius:14px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.07);">
      <div style="padding:16px 22px;background:linear-gradient(135deg,#1A5276,#2E86C1);display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
        <span style="font-size:20px;">🏫</span>
        <span style="font-size:17px;font-weight:700;color:#fff;">${className}</span>
        ${levelBadge(level)}
        <span style="margin-left:auto;color:#AED6F1;font-size:13px;">
          👥 ${totalStudents} HS &nbsp;|&nbsp; ĐTB: <b style="color:#fff;">${avgScore}/10</b> &nbsp;|&nbsp;
          Dưới TB: <b style="color:${parseFloat(belowPct)>30?"#F1948A":"#A9DFBF"};">${belowPct}%</b>
        </span>
      </div>
      <div style="padding:20px;">
        ${buildClassBody(s4, hist, className, totalStudents, avgScore, belowPct, level, maxScore, minScore, "—", hardestQ, easiestQ, "", insights, actions)}
      </div>
    </div>
  `;
}

// ── Toggle collapse lớp ──────────────────────────────────
// ── Build nội dung chi tiết 1 lớp ────────────────────────
// s4 = object lớp trực tiếp từ API; hist = s4.histogram; các tham số còn lại đã extract sẵn
function buildClassBody(s4, hist, className, totalStudents, avgScore, belowPct, level, maxScore, minScore, belowCount, hardestQ, easiestQ, summary, insights, actions) {
  let html = "";

  // Nhận xét tổng thể
  if (summary || insights.length) {
    const insightList = insights.map(t => `<li style="margin-bottom:5px;font-size:13.5px;color:#2c3e50;">${t}</li>`).join("");
    html += `
      <div style="margin-bottom:16px;background:#EBF5FB;border-left:4px solid #2E86C1;border-radius:0 8px 8px 0;padding:14px 16px;">
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
          <span style="font-size:13px;font-weight:600;color:#1a5276;text-transform:uppercase;letter-spacing:.5px;">💬 Nhận Xét Tổng Thể</span>
          ${levelBadge(level)}
        </div>
        ${summary ? `<p style="margin:0 0 10px;font-size:14px;color:#2c3e50;font-weight:600;">${summary}</p>` : ""}
        ${insightList ? `<ul style="margin:0;padding-left:20px;">${insightList}</ul>` : ""}
      </div>`;
  }

  // Thống kê + câu khó/dễ
  html += `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;">
      <div class="table-card" style="padding:14px;">
        <div style="font-size:12px;font-weight:600;color:#5d6d7e;margin-bottom:10px;text-transform:uppercase;">📋 Thống Kê Lớp</div>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tbody>
            <tr><td style="padding:5px 8px;color:#555;">Điểm TB</td><td style="padding:5px 8px;font-weight:700;text-align:right;">${avgScore}/10</td></tr>
            <tr style="background:#f8f9fa;"><td style="padding:5px 8px;color:#555;">Điểm cao nhất</td><td style="padding:5px 8px;font-weight:700;text-align:right;color:#1D6A39;">${maxScore}</td></tr>
            <tr><td style="padding:5px 8px;color:#555;">Điểm thấp nhất</td><td style="padding:5px 8px;font-weight:700;text-align:right;color:#C0392B;">${minScore}</td></tr>
            <tr style="background:#f8f9fa;"><td style="padding:5px 8px;color:#555;">Dưới TB</td><td style="padding:5px 8px;font-weight:700;text-align:right;">${belowCount} HS (${belowPct}%)</td></tr>
            <tr><td style="padding:5px 8px;color:#555;">Xếp loại lớp</td><td style="padding:5px 8px;text-align:right;">${levelBadge(level)}</td></tr>
          </tbody>
        </table>
      </div>
      <div class="table-card" style="padding:14px;">
        <div style="font-size:12px;font-weight:600;color:#5d6d7e;margin-bottom:10px;text-transform:uppercase;">❓ Câu Hỏi Đặc Biệt</div>
        <div style="background:#FADBD8;border-radius:8px;padding:10px 14px;margin-bottom:10px;">
          <div style="font-size:11px;color:#C0392B;font-weight:600;margin-bottom:3px;">🔴 KHÓ NHẤT</div>
          <div style="font-size:15px;font-weight:700;color:#C0392B;">${hardestQ.question || "—"}</div>
          <div style="font-size:12px;color:#666;">ĐTB: ${hardestQ.avg_score ?? "—"}</div>
        </div>
        <div style="background:#D5F5E3;border-radius:8px;padding:10px 14px;">
          <div style="font-size:11px;color:#1D6A39;font-weight:600;margin-bottom:3px;">🟢 DỄ NHẤT</div>
          <div style="font-size:15px;font-weight:700;color:#1D6A39;">${easiestQ.question || "—"}</div>
          <div style="font-size:12px;color:#666;">ĐTB: ${easiestQ.avg_score ?? "—"}</div>
        </div>
      </div>
    </div>`;

  // Histogram
  const histRanges = hist.ranges || [];
  const histCounts = hist.counts || [];
  if (histRanges.length && histCounts.length) {
    const histTotal = histCounts.reduce((s, v) => s + (parseInt(v)||0), 0) || 1;
    const barColors = ["#E74C3C","#E67E22","#F1C40F","#2ECC71","#2E86C1"];
    const bars = histRanges.map((range, i) => {
      const cnt = parseInt(histCounts[i] || 0);
      const pct = Math.round((cnt / histTotal) * 100);
      return `
        <div style="margin-bottom:7px;display:flex;align-items:center;gap:8px;">
          <span style="width:44px;font-size:12px;font-weight:600;color:#555;">${range}</span>
          <div style="flex:1;background:#eee;border-radius:5px;height:16px;overflow:hidden;">
            <div style="width:${pct}%;background:${barColors[i]||"#95A5A6"};height:100%;border-radius:5px;transition:width .4s;"></div>
          </div>
          <span style="width:76px;font-size:12px;color:#333;text-align:right;">${cnt} HS (${pct}%)</span>
        </div>`;
    }).join("");
    html += `
      <div class="table-card" style="padding:14px;margin-bottom:16px;">
        <div style="font-size:12px;font-weight:600;color:#5d6d7e;margin-bottom:12px;text-transform:uppercase;">📊 Phân Bổ Điểm – ${className}</div>
        ${bars}
      </div>`;
  }

  // Top5 / Bottom5
  const top5    = s4.top_students    || [];
  const bottom5 = s4.bottom_students || [];
  if (top5.length || bottom5.length) {
    function studentRows(arr) {
      return arr.map((s, i) => {
        const sc = parseFloat(s.total_score) || 0;
        const c  = scoreColor(sc);
        return `<tr style="${i%2===1?"background:#f8f9fa;":""}">
          <td style="padding:5px 8px;font-size:12px;color:#888;text-align:center;">${i+1}</td>
          <td style="padding:5px 8px;font-weight:600;">${s.name||"—"}</td>
          <td style="padding:5px 8px;font-size:12px;color:#888;">${s.id||"—"}</td>
          <td style="padding:5px 8px;text-align:center;font-weight:700;color:${c.color};background:${c.bg};">${s.total_score??""}</td>
        </tr>`;
      }).join("") || `<tr><td colspan="4" style="padding:8px;color:#999;text-align:center;">—</td></tr>`;
    }
    html += `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;">
        <div class="table-card" style="padding:14px;">
          <div style="font-size:12px;font-weight:600;color:#1D6A39;margin-bottom:8px;text-transform:uppercase;">🏆 Top 5 Cao Nhất</div>
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead><tr style="background:#D5F5E3;"><th style="padding:4px 8px;font-size:11px;color:#1D6A39;">#</th><th style="padding:4px 8px;font-size:11px;color:#1D6A39;text-align:left;">Họ tên</th><th style="padding:4px 8px;font-size:11px;color:#1D6A39;">Mã HS</th><th style="padding:4px 8px;font-size:11px;color:#1D6A39;">Điểm</th></tr></thead>
            <tbody>${studentRows(top5)}</tbody>
          </table>
        </div>
        <div class="table-card" style="padding:14px;">
          <div style="font-size:12px;font-weight:600;color:#C0392B;margin-bottom:8px;text-transform:uppercase;">⚠️ Bottom 5 Thấp Nhất</div>
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead><tr style="background:#FADBD8;"><th style="padding:4px 8px;font-size:11px;color:#C0392B;">#</th><th style="padding:4px 8px;font-size:11px;color:#C0392B;text-align:left;">Họ tên</th><th style="padding:4px 8px;font-size:11px;color:#C0392B;">Mã HS</th><th style="padding:4px 8px;font-size:11px;color:#C0392B;">Điểm</th></tr></thead>
            <tbody>${studentRows(bottom5)}</tbody>
          </table>
        </div>
      </div>`;
  }

  // Question stats — API mới không có question_stats riêng, bỏ qua block này
  // (focus_questions sẽ hiển thị phía dưới)
  const qStats = [];
  if (qStats.length) {
    const maxQ = Math.max(...qStats.map(q => parseFloat(q.avg_score)||0), 1);
    const qRows = qStats.map((q, i) => {
      const c   = scoreColor(q.avg_score);
      const bar = Math.round((parseFloat(q.avg_score)||0) / maxQ * 100);
      return `<tr style="${i%2===1?"background:#f8f9fa;":""}">
        <td style="padding:7px 12px;font-weight:700;color:#2E86C1;width:70px;white-space:nowrap;">${q.question}</td>
        <td style="padding:7px 12px;">
          <div style="background:#e9ecef;border-radius:6px;height:14px;overflow:hidden;width:100%;">
            <div style="width:${bar}%;background:${c.color};height:100%;border-radius:6px;transition:width .4s;"></div>
          </div>
        </td>
        <td style="padding:7px 12px;text-align:center;width:90px;">
          <span style="display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:700;color:${c.color};background:${c.bg};">${q.avg_score}</span>
        </td>
      </tr>`;
    }).join("");
    html += `
      <div class="table-card" style="padding:14px;margin-bottom:16px;">
        <div style="font-size:12px;font-weight:600;color:#5d6d7e;margin-bottom:10px;text-transform:uppercase;">📝 Thống Kê Từng Câu Hỏi</div>
        <table style="width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed;">
          <colgroup><col style="width:70px;"><col><col style="width:90px;"></colgroup>
          <thead><tr style="background:#1A5276;">
            <th style="padding:7px 12px;color:#fff;text-align:left;">Câu</th>
            <th style="padding:7px 12px;color:#fff;text-align:left;">Phân bổ điểm</th>
            <th style="padding:7px 12px;color:#fff;text-align:center;">ĐTB</th>
          </tr></thead>
          <tbody>${qRows}</tbody>
        </table>
      </div>`;
  }

  // Focus questions + Action plan
  const focusQ = s4.focus_questions || [];
  if (focusQ.length || actions.length) {
    const fqRows = focusQ.map((q, qi) => {
      const c = scoreColor(q.avg_score);
      return `<tr style="${qi%2===1?"background:#f8f9fa;":""}">
        <td style="padding:5px 10px;font-weight:700;color:#C0392B;">${q.question??""}</td>
        <td style="padding:5px 10px;text-align:center;"><span style="padding:2px 10px;border-radius:10px;font-weight:700;font-size:12px;color:${c.color};background:${c.bg};">${q.avg_score??""}</span></td>
      </tr>`;
    }).join("") || `<tr><td colspan="2" style="padding:8px;color:#999;">Không có</td></tr>`;
    const apItems = actions.map((a, ai) =>
      `<li style="margin-bottom:8px;padding:9px 14px;background:${ai%2===0?"#EBF5FB":"#F0FFF4"};border-left:3px solid ${ai%2===0?"#2E86C1":"#27AE60"};border-radius:4px;font-size:13px;color:#2c3e50;line-height:1.5;">${a}</li>`
    ).join("") || `<li style="color:#999;">Không có kế hoạch</li>`;
    html += `
      <div class="table-card" style="padding:14px;margin-bottom:4px;">
        <h3 style="margin:0 0 14px;font-size:15px;color:#1a5276;">🎯 Câu Cần Tập Trung & Kế Hoạch Hành Động</h3>
        <div style="display:grid;grid-template-columns:1fr 2fr;gap:16px;">
          <div style="background:#fafbfc;border:1px solid #e8ecf0;border-radius:10px;padding:12px;">
            <div style="font-size:11px;font-weight:600;color:#C0392B;margin-bottom:8px;text-transform:uppercase;">🔴 Câu Cần Tập Trung</div>
            <table style="width:100%;border-collapse:collapse;">
              <thead><tr style="background:#FADBD8;"><th style="padding:4px 10px;font-size:12px;text-align:left;color:#922B21;">Câu</th><th style="padding:4px 10px;font-size:12px;text-align:center;color:#922B21;">ĐTB</th></tr></thead>
              <tbody>${fqRows}</tbody>
            </table>
          </div>
          <div style="background:#fafbfc;border:1px solid #e8ecf0;border-radius:10px;padding:12px;">
            <div style="font-size:11px;font-weight:600;color:#1a5276;margin-bottom:8px;text-transform:uppercase;">📌 Kế Hoạch Hành Động</div>
            <ol style="margin:0;padding-left:18px;">${apItems}</ol>
          </div>
        </div>
      </div>`;
  }

  return html;
}

// ══════════════════════════════════════════════════════════════
// EXPORT: Excel with 4 Sheets
// ══════════════════════════════════════════════════════════════
document.getElementById("btn-export-excel").addEventListener("click", exportExcel);

function exportExcel() {
  // If API data is available, export from API response
  if (apiSheet1Data.length || apiSheet2Data.length || apiSheet3Data.length || apiSheet4Data.length) {
    exportExcelFromAPI();
    return;
  }
  exportExcelFromComputed();
}

function exportExcelFromAPI() {
  const wb = XLSX.utils.book_new();

  // ── Shared style definitions ──────────────────────────────────────
  const S = {
    hdr: {                          // Header row: trắng bold trên nền xanh
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11, name: "Calibri" },
      fill: { patternType: "solid", fgColor: { rgb: "1A5276" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: {
        bottom: { style: "medium", color: { rgb: "117A65" } },
        right:  { style: "thin",   color: { rgb: "ABEBC6" } }
      }
    },
    riskHigh: {                     // Nguy cơ Cao: đỏ
      font: { bold: true, color: { rgb: "922B21" }, sz: 10, name: "Calibri" },
      fill: { patternType: "solid", fgColor: { rgb: "FADBD8" } },
      alignment: { horizontal: "center", vertical: "center" }
    },
    riskMid: {                      // Nguy cơ Trung bình: cam
      font: { bold: true, color: { rgb: "784212" }, sz: 10, name: "Calibri" },
      fill: { patternType: "solid", fgColor: { rgb: "FDEBD0" } },
      alignment: { horizontal: "center", vertical: "center" }
    },
    riskLow: {                      // Nguy cơ Thấp: xanh lá
      font: { bold: false, color: { rgb: "1D6A39" }, sz: 10, name: "Calibri" },
      fill: { patternType: "solid", fgColor: { rgb: "D5F5E3" } },
      alignment: { horizontal: "center", vertical: "center" }
    },
    fail: {                         // Không đạt: đỏ nhạt
      font: { bold: true, color: { rgb: "C0392B" }, sz: 10 },
      fill: { patternType: "solid", fgColor: { rgb: "FADBD8" } },
      alignment: { horizontal: "center", vertical: "center" }
    },
    pass: {                         // Đạt: vàng nhạt
      font: { color: { rgb: "7E5109" }, sz: 10 },
      fill: { patternType: "solid", fgColor: { rgb: "FEF9E7" } },
      alignment: { horizontal: "center", vertical: "center" }
    },
    exc: {                          // Xuất sắc: xanh lá
      font: { bold: true, color: { rgb: "1D6A39" }, sz: 10 },
      fill: { patternType: "solid", fgColor: { rgb: "D5F5E3" } },
      alignment: { horizontal: "center", vertical: "center" }
    },
    altRow: {                       // Hàng xen kẽ: nền xám rất nhạt
      fill: { patternType: "solid", fgColor: { rgb: "F2F3F4" } },
      font: { sz: 10, name: "Calibri" },
      alignment: { vertical: "center", wrapText: true }
    },
    normal: {
      font: { sz: 10, name: "Calibri" },
      alignment: { vertical: "center", wrapText: true }
    },
    numCenter: {
      font: { sz: 10, name: "Calibri" },
      alignment: { horizontal: "center", vertical: "center" }
    }
  };

  // ── Helpers ───────────────────────────────────────────────────────
  // Trả về style risk dựa theo giá trị chuỗi
  function riskStyle(val) {
    const v = String(val || "").toUpperCase();
    if (v === "HIGH"   || v.includes("CAO"))   return S.riskHigh;
    if (v === "MEDIUM" || v.includes("TRUNG")) return S.riskMid;
    return S.riskLow;  // LOW hoặc bất kỳ
  }

  // Tô màu header cho ws
  function applyHeader(ws, keys) {
    keys.forEach((_, ci) => {
      const addr = XLSX.utils.encode_cell({ r: 0, c: ci });
      if (ws[addr]) ws[addr].s = S.hdr;
    });
  }

  // Tô màu toàn bộ row dữ liệu (nền xen kẽ) + base font
  function applyRowBase(ws, ri, colCount, alt) {
    for (let ci = 0; ci < colCount; ci++) {
      const addr = XLSX.utils.encode_cell({ r: ri, c: ci });
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };
      ws[addr].s = alt ? { ...S.altRow } : { ...S.normal };
    }
  }

  // Set style cho 1 ô cụ thể
  function setCell(ws, ri, ci, style) {
    const addr = XLSX.utils.encode_cell({ r: ri, c: ci });
    if (!ws[addr]) ws[addr] = { v: "", t: "s" };
    ws[addr].s = style;
  }

  // Freeze row đầu (header)
  function freezeTop(ws) {
    ws["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft" };
  }

  // ── SHEET 1 — Phân tích Chuẩn ────────────────────────────────────
  if (apiSheet1Data.length) {
    const keys = Object.keys(apiSheet1Data[0]);
    const ws1 = XLSX.utils.json_to_sheet(apiSheet1Data);

    // Column widths thông minh
    const colW = keys.map(k => {
      const kl = k.toLowerCase();
      if (kl.includes("goi_y") || kl.includes("gợi") || kl.includes("suggestion")) return { wch: 48 };
      if (kl.includes("chuan") || kl.includes("chuẩn") || kl.includes("standard")) return { wch: 14 };
      if (kl.includes("ti_le") || kl.includes("tỉ") || kl.includes("pct") || kl.includes("%")) return { wch: 14 };
      return { wch: 16 };
    });
    ws1["!cols"] = colW;
    ws1["!rows"] = [{ hpt: 30 }]; // Header row cao hơn

    applyHeader(ws1, keys);
    freezeTop(ws1);

    apiSheet1Data.forEach((row, ri) => {
      const rowIdx = ri + 1;
      applyRowBase(ws1, rowIdx, keys.length, ri % 2 === 1);

      keys.forEach((k, ci) => {
        const val = row[k];
        const kl = k.toLowerCase();

        // Tô màu % không đạt
        if (kl.includes("fail") || kl.includes("khong_dat") || kl.includes("không đạt") || kl.includes("ti_le_fail")) {
          const pct = parseFloat(val) || 0;
          if (pct > 40)      setCell(ws1, rowIdx, ci, S.fail);
          else if (pct > 20) setCell(ws1, rowIdx, ci, S.pass);
          else               setCell(ws1, rowIdx, ci, S.exc);
        }
        // Tô màu % xuất sắc / pass
        if (kl.includes("pass") || kl.includes("dat") || kl.includes("xuat_sac") || kl.includes("excellent")) {
          const pct = parseFloat(val) || 0;
          if (pct >= 70) setCell(ws1, rowIdx, ci, S.exc);
        }
      });
    });

    XLSX.utils.book_append_sheet(wb, ws1, "Phan Tich Chuan");
  }

  // ── SHEET 2 — Phân tích Học Sinh ─────────────────────────────────
  if (apiSheet2Data.length) {
    // Flatten array fields (Chuẩn mạnh/yếu là array) + null Tổng điểm
    const s2Flat = apiSheet2Data.map(r => {
      const out = { ...r };
      if (Array.isArray(out["Chuẩn mạnh"])) out["Chuẩn mạnh"] = out["Chuẩn mạnh"].join(", ");
      if (Array.isArray(out["Chuẩn yếu"]))  out["Chuẩn yếu"]  = out["Chuẩn yếu"].join(", ");
      if (out["Tổng điểm"] === null) out["Tổng điểm"] = "";
      return out;
    });
    const keys = Object.keys(s2Flat[0]);
    const ws2 = XLSX.utils.json_to_sheet(s2Flat);

    ws2["!cols"] = keys.map(k => {
      const kl = k.toLowerCase();
      if (kl.includes("họ") || kl.includes("ten") || kl.includes("name"))     return { wch: 22 };
      if (kl.includes("đánh giá") || kl.includes("evaluation")) return { wch: 44 };
      if (kl.includes("chuẩn") || kl.includes("chuan"))  return { wch: 28 };
      if (kl.includes("lớp") || kl.includes("class"))  return { wch: 12 };
      if (kl.includes("mã") || kl.includes("id"))      return { wch: 10 };
      return { wch: 14 };
    });
    ws2["!rows"] = [{ hpt: 30 }];

    applyHeader(ws2, keys);
    freezeTop(ws2);

    // Tìm index cột Risk và Score(10) — khớp chính xác key "Score (10)"
    const riskColIdx  = keys.findIndex(k => /^risk/i.test(k));
    const diem10Idx   = keys.findIndex(k => k === "Score (10)" || /thang.?10|diem.?10/i.test(k));
    const failColIdx  = keys.indexOf("Fail");
    const excColIdx   = keys.indexOf("Excellent");

    s2Flat.forEach((row, ri) => {
      const rowIdx = ri + 1;
      applyRowBase(ws2, rowIdx, keys.length, ri % 2 === 1);

      // Cột điểm thang 10
      if (diem10Idx >= 0) {
        const d = parseFloat(row[keys[diem10Idx]]) || 0;
        if (d < 5)      setCell(ws2, rowIdx, diem10Idx, { ...S.fail, alignment: { horizontal: "center" } });
        else if (d < 7) setCell(ws2, rowIdx, diem10Idx, { ...S.pass, alignment: { horizontal: "center" } });
        else            setCell(ws2, rowIdx, diem10Idx, { ...S.exc,  alignment: { horizontal: "center" } });
      }
      // Cột Risk
      if (riskColIdx >= 0) setCell(ws2, rowIdx, riskColIdx, riskStyle(row[keys[riskColIdx]]));
      // Cột Fail
      if (failColIdx >= 0) {
        const n = parseInt(row["Fail"]) || 0;
        if (n >= 4)     setCell(ws2, rowIdx, failColIdx, { ...S.fail, alignment: { horizontal: "center" } });
        else if (n > 0) setCell(ws2, rowIdx, failColIdx, { ...S.pass, alignment: { horizontal: "center" } });
        else            setCell(ws2, rowIdx, failColIdx, { ...S.exc,  alignment: { horizontal: "center" } });
      }
      // Cột Excellent
      if (excColIdx >= 0) {
        const n = parseInt(row["Excellent"]) || 0;
        if (n >= 8) setCell(ws2, rowIdx, excColIdx, { ...S.exc, alignment: { horizontal: "center" } });
      }
    });

    XLSX.utils.book_append_sheet(wb, ws2, "Phan Tich Hoc Sinh");
  }

  // ── SHEET 3 — Học Sinh Yếu ───────────────────────────────────────
  if (apiSheet3Data.length) {
    // Flatten Chuẩn yếu array + null Tổng điểm
    const s3Flat = apiSheet3Data.map(r => {
      const out = { ...r };
      if (Array.isArray(out["Chuẩn yếu"])) out["Chuẩn yếu"] = out["Chuẩn yếu"].join(", ");
      if (out["Tổng điểm"] === null) out["Tổng điểm"] = "";
      return out;
    });
    const keys = Object.keys(s3Flat[0]);
    const ws3 = XLSX.utils.json_to_sheet(s3Flat);

    ws3["!cols"] = keys.map(k => {
      const kl = k.toLowerCase();
      if (kl.includes("gợi ý") || kl.includes("can_thiep")) return { wch: 52 };
      if (kl.includes("chuẩn") || kl.includes("chuan")) return { wch: 30 };
      return { wch: 14 };
    });
    ws3["!rows"] = [{ hpt: 30 }];

    applyHeader(ws3, keys);
    freezeTop(ws3);

    const riskColIdx3 = keys.findIndex(k => /^risk/i.test(k));

    s3Flat.forEach((row, ri) => {
      const rowIdx = ri + 1;
      const riskVal = riskColIdx3 >= 0 ? row[keys[riskColIdx3]] : "";
      const isHigh = String(riskVal).toUpperCase() === "HIGH";

      if (isHigh) {
        for (let ci = 0; ci < keys.length; ci++) {
          setCell(ws3, rowIdx, ci, {
            font: { sz: 10, name: "Calibri", color: { rgb: "7B241C" } },
            fill: { patternType: "solid", fgColor: { rgb: "FADBD8" } },
            alignment: { vertical: "center", wrapText: true }
          });
        }
      } else {
        applyRowBase(ws3, rowIdx, keys.length, ri % 2 === 1);
      }
      if (riskColIdx3 >= 0) setCell(ws3, rowIdx, riskColIdx3, riskStyle(riskVal));
    });

    XLSX.utils.book_append_sheet(wb, ws3, "Hoc Sinh Yeu");
  }

  // ── SHEET 4 — Phân tích Lớp (mỗi lớp 1 dòng, cấu trúc API thực tế) ──
  if (apiSheet4Data.length) {
    // Mỗi phần tử trong apiSheet4Data = 1 lớp: { class_name, summary, histogram, insights, top_students, bottom_students, hardest_question, easiest_question, focus_questions, action_plan }
    const s4Flat = apiSheet4Data.map(cls => {
      const smr  = cls.summary           || {};
      const hist = cls.histogram         || {};
      const hardQ = cls.hardest_question || {};
      const easyQ = cls.easiest_question || {};

      const histRanges = hist.ranges || [];
      const histCounts = hist.counts || [];
      const histStr    = histRanges.map((r, i) => `${r}: ${histCounts[i] ?? 0}`).join(" | ") || "—";

      const focusQ = cls.focus_questions || [];
      const fqStr  = focusQ.map(q => `${q.question}(ĐTB:${q.avg_score})`).join(", ") || "—";

      const ap    = cls.action_plan || [];
      const apStr = ap.map((a, i) => `${i+1}. ${a}`).join("\n") || "—";

      const top5str    = (cls.top_students    || []).map(s => `${s.name}(${s.total_score})`).join(", ") || "—";
      const bottom5str = (cls.bottom_students || []).map(s => `${s.name}(${s.total_score})`).join(", ") || "—";

      return {
        "Lớp":                cls.class_name           || "—",
        "Sĩ số":              smr.total_students        ?? "",
        "Điểm TB":            smr.avg_score             ?? "",
        "Điểm Max":           smr.max_score             ?? "",
        "Điểm Min":           smr.min_score             ?? "",
        "Dưới TB (%)":        smr.below_avg_percent     ?? "",
        "Xếp loại lớp":       smr.level                 || "—",
        "Câu khó nhất":       hardQ.question            || "—",
        "ĐTB câu khó":        hardQ.avg_score           ?? "",
        "Câu dễ nhất":        easyQ.question            || "—",
        "ĐTB câu dễ":         easyQ.avg_score           ?? "",
        "Phân Bổ Điểm":       histStr,
        "Câu Cần Tập Trung":  fqStr,
        "Top 5 HS":           top5str,
        "Bottom 5 HS":        bottom5str,
        "Kế Hoạch Hành Động": apStr,
      };
    });

    const keys = Object.keys(s4Flat[0]);
    const ws4 = XLSX.utils.json_to_sheet(s4Flat);

    ws4["!cols"] = [
      { wch: 16 }, // Lớp
      { wch: 8  }, // Sĩ số
      { wch: 10 }, // Điểm TB
      { wch: 10 }, // Điểm Max
      { wch: 10 }, // Điểm Min
      { wch: 12 }, // Dưới TB %
      { wch: 14 }, // Xếp loại
      { wch: 14 }, // Câu khó nhất
      { wch: 12 }, // ĐTB câu khó
      { wch: 14 }, // Câu dễ nhất
      { wch: 12 }, // ĐTB câu dễ
      { wch: 38 }, // Phân bổ
      { wch: 30 }, // Câu cần tập trung
      { wch: 36 }, // Top 5
      { wch: 36 }, // Bottom 5
      { wch: 64 }, // Kế hoạch
    ];
    ws4["!rows"] = [{ hpt: 30 }];

    applyHeader(ws4, keys);
    freezeTop(ws4);

    const scoreIdx4 = keys.indexOf("Điểm TB");
    const belowIdx4 = keys.indexOf("Dưới TB (%)");
    const levelIdx4 = keys.indexOf("Xếp loại lớp");
    const apIdx4    = keys.indexOf("Kế Hoạch Hành Động");

    s4Flat.forEach((row, ri) => {
      const rowIdx = ri + 1;
      applyRowBase(ws4, rowIdx, keys.length, ri % 2 === 1);

      // Điểm TB (thang 10)
      if (scoreIdx4 >= 0) {
        const d = parseFloat(row["Điểm TB"]) || 0;
        if (d > 0) {
          if (d < 4)        setCell(ws4, rowIdx, scoreIdx4, { ...S.fail, alignment: { horizontal: "center" } });
          else if (d < 6.5) setCell(ws4, rowIdx, scoreIdx4, { ...S.pass, alignment: { horizontal: "center" } });
          else              setCell(ws4, rowIdx, scoreIdx4, { ...S.exc,  alignment: { horizontal: "center" } });
        }
      }
      // Dưới TB %
      if (belowIdx4 >= 0) {
        const pct = parseFloat(row["Dưới TB (%)"]) || 0;
        if (pct >= 40)      setCell(ws4, rowIdx, belowIdx4, { ...S.fail, alignment: { horizontal: "center" } });
        else if (pct >= 20) setCell(ws4, rowIdx, belowIdx4, { ...S.pass, alignment: { horizontal: "center" } });
        else                setCell(ws4, rowIdx, belowIdx4, { ...S.exc,  alignment: { horizontal: "center" } });
      }
      // Xếp loại lớp
      if (levelIdx4 >= 0) {
        const lv = row["Xếp loại lớp"] || "";
        const lvMap = { "Xuất sắc": S.exc, "Tốt": S.exc, "Khá": S.pass, "Trung bình": S.pass, "Yếu": S.fail, "Kém": S.fail };
        const st = lvMap[lv];
        if (st) setCell(ws4, rowIdx, levelIdx4, { ...st, alignment: { horizontal: "center" } });
      }
      // Kế hoạch hành động: wrap text
      if (apIdx4 >= 0) {
        const addr = XLSX.utils.encode_cell({ r: rowIdx, c: apIdx4 });
        if (!ws4[addr]) ws4[addr] = { v: "", t: "s" };
        ws4[addr].s = { ...(ws4[addr].s || {}), alignment: { wrapText: true, vertical: "top" } };
      }
    });

    XLSX.utils.book_append_sheet(wb, ws4, "Phan Tich Lop");
  }

  XLSX.writeFile(wb, "AI_Learning_Analytics_Report.xlsx");
}

function exportExcelFromComputed() {
  const wb = XLSX.utils.book_new();
  const stats = computeStandardStats(students);

  // ── Style Definitions ──
  const headerStyle = {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
    fill: { fgColor: { rgb: "2E86C1" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: { bottom: { style: "thin", color: { rgb: "1B4F72" } } }
  };
  const riskStyles = {
    "Cao": { fill: { fgColor: { rgb: "FADBD8" } }, font: { color: { rgb: "C0392B" }, bold: true } },
    "Trung bình": { fill: { fgColor: { rgb: "FEF9E7" } }, font: { color: { rgb: "B7950B" } } },
    "Thấp": { fill: { fgColor: { rgb: "D5F5E3" } }, font: { color: { rgb: "1E8449" } } },
  };
  const scoreStyles = {
    fail: { fill: { fgColor: { rgb: "FADBD8" } }, font: { color: { rgb: "C0392B" }, bold: true } },
    pass: { fill: { fgColor: { rgb: "FEF9E7" } }, font: { color: { rgb: "B7950B" } } },
    excellent: { fill: { fgColor: { rgb: "D5F5E3" } }, font: { color: { rgb: "1E8449" }, bold: true } },
  };
  const dangerRow = { fill: { fgColor: { rgb: "F5B7B1" } }, font: { color: { rgb: "922B21" }, bold: true } };
  const normalCell = { alignment: { vertical: "center" } };

  // Helper: apply style to all cells in a row
  function styleRow(ws, rowIdx, colCount, style) {
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };
      ws[addr].s = { ...normalCell, ...style };
    }
  }

  // Helper: apply header style
  function styleHeader(ws, colCount) {
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[addr]) ws[addr].s = headerStyle;
    }
  }

  // ═══════════════════════════════════════
  // Sheet 1 — Standard Analysis (colored)
  // ═══════════════════════════════════════
  const s1Hdr = ["Chuẩn", "% Không Đạt (0–1)", "% Đạt (2–3)", "% Xuất Sắc (4)", "Cần Cải Thiện", "Gợi Ý"];
  const s1Data = [s1Hdr];
  stats.forEach(s => s1Data.push([s.std, s.failPct + "%", s.passPct + "%", s.excPct + "%", s.needImprovement + " / " + students.length, s.suggestion]));
  const ws1 = XLSX.utils.aoa_to_sheet(s1Data);
  ws1["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 18 }, { wch: 44 }];
  styleHeader(ws1, s1Hdr.length);
  stats.forEach((s, i) => {
    const rowIdx = i + 1;
    if (s.failPct > 40) {
      styleRow(ws1, rowIdx, s1Hdr.length, dangerRow);
    } else if (s.failPct > 20) {
      styleRow(ws1, rowIdx, s1Hdr.length, riskStyles["Trung bình"]);
    } else {
      styleRow(ws1, rowIdx, s1Hdr.length, riskStyles["Thấp"]);
    }
  });
  XLSX.utils.book_append_sheet(wb, ws1, "Phân Tích Chuẩn");

  // ═══════════════════════════════════════
  // Sheet 2 — Student Analytics (colored)
  // ═══════════════════════════════════════
  const s2Hdr = ["Mã HS", "Họ và Tên", "Lớp", ...STANDARDS, "Tổng Điểm", "Điểm /10", "Xuất Sắc", "Đạt", "Không Đạt", "Chuẩn Mạnh", "Chuẩn Yếu", "Cần Cải Thiện", "Mức Nguy Cơ", "Đánh Giá"];
  const s2Data = [s2Hdr];
  students.forEach(s => {
    s2Data.push([s.id, s.name, s.class, ...s.scores, s.totalScore, s.score10, s.excellentCount, s.passCount, s.failCount, s.strongStds.join(", "), s.weakStds.join(", "), s.improveStds.join(", "), s.risk, s.evaluation]);
  });
  const ws2 = XLSX.utils.aoa_to_sheet(s2Data);
  ws2["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 8 }, ...STANDARDS.map(() => ({ wch: 5 })), { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 22 }, { wch: 22 }, { wch: 22 }, { wch: 12 }, { wch: 50 }];
  styleHeader(ws2, s2Hdr.length);

  students.forEach((s, i) => {
    const rowIdx = i + 1;
    // Row background by risk
    const rowStyle = riskStyles[s.risk] || {};
    styleRow(ws2, rowIdx, s2Hdr.length, rowStyle);

    // Individual score cells (columns 3 to 3+STANDARDS.length-1) colored by classification
    s.scores.forEach((score, si) => {
      const colIdx = 3 + si;
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      if (ws2[addr]) {
        const cls = classify(score);
        ws2[addr].s = { ...normalCell, ...scoreStyles[cls], alignment: { horizontal: "center" } };
      }
    });

    // Risk cell bold colored
    const riskColIdx = s2Hdr.indexOf("Mức Nguy Cơ");
    const riskAddr = XLSX.utils.encode_cell({ r: rowIdx, c: riskColIdx });
    if (ws2[riskAddr]) {
      const rs = riskStyles[s.risk];
      ws2[riskAddr].s = { font: { bold: true, color: rs.font.color, sz: 11 }, fill: rs.fill, alignment: { horizontal: "center" } };
    }
  });
  XLSX.utils.book_append_sheet(wb, ws2, "Phân Tích Học Sinh");

  // ═══════════════════════════════════════
  // Sheet 3 — Weak Students (colored)
  // ═══════════════════════════════════════
  const weak = students.filter(s => s.score10 < 5 || s.failCount >= 4);
  const s3Hdr = ["Mã HS", "Họ và Tên", "Tổng Điểm", "Điểm /10", "Chuẩn Yếu", "Số Chuẩn Không Đạt", "Mức Nguy Cơ", "Gợi Ý Can Thiệp"];
  const s3Data = [s3Hdr];
  weak.forEach(s => {
    let intervention;
    if (s.risk === "Cao") intervention = "Tổ chức ôn tập nền tảng và kèm cặp một-một.";
    else if (s.failCount >= 6) intervention = "Giao bài tập có hướng dẫn tập trung vào các chuẩn cốt lõi.";
    else intervention = "Củng cố các điểm yếu bằng bài tập bổ sung và học nhóm.";
    s3Data.push([s.id, s.name, s.totalScore, s.score10, s.weakStds.join(", "), s.failCount, s.risk, intervention]);
  });
  const ws3 = XLSX.utils.aoa_to_sheet(s3Data);
  ws3["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 10 }, { wch: 10 }, { wch: 26 }, { wch: 18 }, { wch: 14 }, { wch: 50 }];
  styleHeader(ws3, s3Hdr.length);

  weak.forEach((s, i) => {
    const rowIdx = i + 1;
    if (s.risk === "Cao") {
      styleRow(ws3, rowIdx, s3Hdr.length, dangerRow);
    } else {
      styleRow(ws3, rowIdx, s3Hdr.length, riskStyles[s.risk] || riskStyles["Trung bình"]);
    }
  });
  XLSX.utils.book_append_sheet(wb, ws3, "Học Sinh Yếu");

  XLSX.writeFile(wb, "AI_Learning_Analytics_Report.xlsx");
} // end exportExcelFromComputed

// ══════════════════════════════════════════════════════════════
// HELPER: Bỏ dấu tiếng Việt cho PDF (jsPDF helvetica ko hỗ trợ Unicode)
// ══════════════════════════════════════════════════════════════
function stripVN(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")   // bỏ combining diacritics
    .replace(/đ/g, "d").replace(/Đ/g, "D")
    .replace(/\u0111/g, "d").replace(/\u0110/g, "D");
}

// ══════════════════════════════════════════════════════════════
// EXPORT: Individual Student PDF
// ══════════════════════════════════════════════════════════════
function generateStudentPDF(st) {
  // If API data is available, find matching row and use API-based generator
  if (apiSheet2Data.length) {
    const row = apiSheet2Data.find(r =>
      String(r["Mã HS"] || r["ma_hs"] || r["id"] || "") === String(st.id)
    );
    if (row) { generateStudentPDFFromAPI(row); return; }
  }
  generateStudentPDFLegacy(st);
}

// ── PDF từ dữ liệu API (Sheet 1 & Sheet 2) ──────────────────
function generateStudentPDFFromAPI(studentRow) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const W = doc.internal.pageSize.getWidth();
  const M = 16;
  let y = 0;

  const v = (s) => stripVN(s);
  const setF = (style, size) => { doc.setFontSize(size); doc.setFont("helvetica", style); };
  const wrap = (text, x, yy, maxW, lh) => {
    const lines = doc.splitTextToSize(v(text), maxW || W - M * 2);
    doc.text(lines, x, yy);
    return yy + lines.length * (lh || doc.getFontSize() * 0.4);
  };
  const check = (need) => { if (y + need > 278) { doc.addPage(); y = 20; } };

  const name     = studentRow["Họ và tên"]   || studentRow["name"]  || "—";
  const maHS     = studentRow["Mã HS"]        || studentRow["id"]    || "—";
  const lop      = studentRow["Lớp"]          || studentRow["class"] || "—";
  // Điểm: API trả "Score (10)", "Tổng điểm" có thể null
  const tongDiem = studentRow["Tổng điểm"] ?? studentRow["Tổng điểm thô"] ?? studentRow["tong_diem"] ?? "—";
  const diem10   = studentRow["Score (10)"] ?? studentRow["Điểm thang 10"] ?? studentRow["Điểm quy đổi thang 10"] ?? studentRow["diem_10"] ?? "—";
  // Số câu: API trả "Excellent", "Pass", "Fail" (không có prefix "Số")
  const soExc    = studentRow["Excellent"]    ?? studentRow["Số Excellent"]    ?? studentRow["Số câu Xuất sắc"] ?? "0";
  const soPass   = studentRow["Pass"]         ?? studentRow["Số câu Pass"]     ?? studentRow["Số câu Đạt"]    ?? "0";
  const soFail   = studentRow["Fail"]         ?? studentRow["Số câu Fail"]     ?? studentRow["Số câu Không đạt"] ?? "0";
  const risk     = studentRow["Risk Level"]   || studentRow["risk_level"] || "—";
  const danhGia  = studentRow["Đánh giá học sinh"] || studentRow["danh_gia"] || "";
  // Chuẩn mạnh/yếu từ API là array
  const manhRaw  = studentRow["Chuẩn mạnh"] || studentRow["chuan_manh"] || [];
  const yeuRaw   = studentRow["Chuẩn yếu"]  || studentRow["chuan_yeu"]  || [];
  const manhStr  = Array.isArray(manhRaw) ? manhRaw.join(", ") : String(manhRaw);
  const yeuStr   = Array.isArray(yeuRaw)  ? yeuRaw.join(", ")  : String(yeuRaw);

  // ── Tìm dòng raw có điểm từng câu từ _lastParsedData
  // Thử khớp theo: Mã HS → Họ và tên → Lớp → fallback dòng đầu
  const rawRow = (_lastParsedData || []).find(r => {
    // Khớp Mã HS nếu có
    const rid = String(r["Mã HS"] || r["ma_hs"] || r["id"] || r["Student_ID"] || "");
    if (maHS !== "—" && rid && rid === String(maHS)) return true;
    // Khớp theo Họ và tên (trim để bỏ khoảng trắng thừa)
    const rname = String(r["Họ và tên"] || r["ho_va_ten"] || r["name"] || "").trim();
    if (rname && rname === String(name).trim()) return true;
    return false;
  }) || (_lastParsedData && _lastParsedData[0]) || {};

  // Merge: ưu tiên rawRow cho cột điểm từng câu (key dài), studentRow cho tổng hợp
  const mergedRow = Object.assign({}, rawRow, studentRow);

  // ── Header Banner
  doc.setFillColor(46, 134, 193);
  doc.rect(0, 0, W, 42, "F");
  doc.setTextColor(255, 255, 255);
  setF("bold", 18);
  doc.text("Bao Cao Ket Qua Hoc Sinh", M, 16);
  setF("normal", 10);
  doc.text("He Thong Phan Tich Hieu Suat Hoc Tap AI", M, 24);
  doc.text("Ngay tao: " + new Date().toLocaleDateString("vi-VN", { day: "2-digit", month: "2-digit", year: "numeric" }), M, 32);
  y = 52;

  // ── Student Info Block
  doc.setTextColor(44, 62, 80);
  setF("bold", 15);
  doc.text(v(name), M, y); y += 8;
  setF("normal", 10);
  doc.setTextColor(100);
  doc.text(`Ma HS: ${v(maHS)}    |    Lop: ${v(lop)}    |    Tong diem: ${tongDiem !== null && tongDiem !== "—" ? v(tongDiem) : "N/A"}    |    Diem thang 10: ${v(diem10)}`, M, y); y += 5;
  doc.text(`Muc nguy co: ${v(risk)}    |    Xuat sac: ${v(soExc)}    |    Dat: ${v(soPass)}    |    Khong dat: ${v(soFail)}`, M, y); y += 5;
  if (manhStr) { y = wrap(`Chuan manh: ${manhStr}`, M, y, W - M * 2, 4.5); }
  if (yeuStr)  { y = wrap(`Chuan yeu: ${yeuStr}`,  M, y, W - M * 2, 4.5); }
  y += 5;

  // ── Divider
  doc.setDrawColor(200); doc.line(M, y, W - M, y); y += 8;

  // ── Standards Breakdown từ Sheet 1, sắp xếp câu số tăng dần
  doc.setTextColor(44, 62, 80);
  setF("bold", 12);
  doc.text("Chi Tiet Tung Cau", M, y); y += 8;

  const sheet1Rows = [...(apiSheet1Data || [])].sort((a, b) => {
    const na = parseInt(String(a["câu_hỏi"] || a["cau_hoi"] || a["Câu hỏi"] || a["standard_code"] || "").match(/\d+/)?.[0] || 0);
    const nb = parseInt(String(b["câu_hỏi"] || b["cau_hoi"] || b["Câu hỏi"] || b["standard_code"] || "").match(/\d+/)?.[0] || 0);
    return na - nb;
  });

  const allKeys = Object.keys(mergedRow);
  const cauKeys = allKeys
    .filter(k => {
      const kl = k.toLowerCase();
      // Match "câu N", "cau N" (có/không dấu), "cN", "c N" — key có thể là multiline
      return /^c[aâ]u\s*\d+/.test(kl) || /^c\d+$/.test(kl);
    })
    .sort((a, b) => {
      const na = parseInt(a.match(/\d+/)?.[0] || 0);
      const nb = parseInt(b.match(/\d+/)?.[0] || 0);
      return na - nb;
    });
  console.log("[PDF DEBUG] rawRow found:", Object.keys(rawRow).length, "keys");
  console.log("[PDF DEBUG] cauKeys:", cauKeys.length, cauKeys.slice(0,3));
  console.log("[PDF DEBUG] mergedRow sample keys:", Object.keys(mergedRow).slice(0,5));

  // Helper: trích mã chuẩn từ key dạng "Câu 1\n (Max:4)\n Chuẩn:\n 3Np.05 ..."
  function extractStdCodeFromKey(key) {
    const norm = String(key).toLowerCase();
    // Sau "chuẩn:" hoặc "chuan:" — lấy token alphanum đầu tiên
    const m = norm.match(/chu[aâ]n\s*[:\-]?\s*([a-z0-9][a-z0-9\._\-]+)/i)
           || String(key).match(/chu[aâ]n\s*[:\-]?\s*([A-Z0-9][A-Za-z0-9\._\-]+)/);
    return m ? String(key).substring(key.toLowerCase().indexOf(m[1]), key.toLowerCase().indexOf(m[1]) + m[1].length) : null;
  }

  // Helper: trích số câu từ key (dùng lowercase để tránh vấn đề Unicode với /i flag)
  function extractCauNum(key) {
    const kl = String(key).toLowerCase();
    const m = kl.match(/^c[aâ]u\s*(\d+)/) || kl.match(/^c(\d+)$/);
    return m ? parseInt(m[1]) : 0;
  }

  // Zebra stripe
  // Layout: Cau(12) | Diem(16) | Ket qua(24) | %K.Dat(22) | MaChuan(28) | GoiY(rest)
  const colX     = [M, M + 12, M + 28, M + 52, M + 76, M + 108];
  const goiYMaxW = W - M - colX[5] - 2;
  const thdrs = ["Cau", "Diem", "Ket qua", "% K.Dat", "Ma Chuan", "Goi Y"];
  doc.setFillColor(234, 240, 246);
  doc.rect(M, y - 4, W - M * 2, 8, "F");
  setF("bold", 8);
  doc.setTextColor(44, 62, 80);
  thdrs.forEach((h, i) => doc.text(h, colX[i], y));
  y += 7;

  setF("normal", 7.5);

  cauKeys.forEach((key, idx) => {
    const score    = parseFloat(mergedRow[key]) || 0;
    const maxScore = 4;
    const cauNum   = extractCauNum(key) || (idx + 1);

    // Match với Sheet 1 theo số câu
    const s1Row = sheet1Rows.find(r => {
      const cauHoi = String(r["câu_hỏi"] || r["cau_hoi"] || r["Câu hỏi"] || r["standard_code"] || "");
      const num = parseInt(cauHoi.match(/\d+/)?.[0] || 0);
      return num === cauNum;
    }) || sheet1Rows[idx];

    const failPct = s1Row ? parseFloat(s1Row["tỉ_lệ_fail"] || s1Row["ti_le_fail"] || 0) : 0;
    // Ưu tiên: Sheet1 standard_code → trích từ key Excel → fallback C{n}
    const stdCode = (s1Row && (s1Row["standard_code"] || s1Row["Standard Code"]))
      || extractStdCodeFromKey(key)
      || `C${cauNum}`;
    const suggest = v(String(s1Row ? (s1Row["suggestion"] || s1Row["Suggestion"] || "") : ""));

    // Tính số dòng Goi Y cần
    const goiYLines = doc.splitTextToSize(suggest, goiYMaxW);
    const rowH = Math.max(6, goiYLines.length * 4.2 + 1);
    check(rowH + 2);

    let ketQua, r, g, b;
    if (score >= 4)      { ketQua = "Xuat sac";  r=39;  g=174; b=96;  }
    else if (score >= 2) { ketQua = "Dat";        r=243; g=156; b=18;  }
    else                 { ketQua = "Khong dat";  r=231; g=76;  b=60;  }

    // Zebra stripe
    if (idx % 2 === 1) {
      doc.setFillColor(249, 249, 249);
      doc.rect(M, y - 4, W - M * 2, rowH, "F");
    }

    doc.setTextColor(44, 62, 80);
    doc.text(String(cauNum), colX[0], y);
    doc.text(`${score}/${maxScore}`, colX[1], y);

    doc.setTextColor(r, g, b);
    setF("bold", 7.5);
    doc.text(ketQua, colX[2], y);
    setF("normal", 7.5);

    doc.setTextColor(failPct >= 40 ? 231 : 80, failPct >= 40 ? 76 : 80, failPct >= 40 ? 60 : 80);
    doc.text(failPct + "%", colX[3], y);

    doc.setTextColor(80);
    // stdCode có thể dài như "3Np.05", hiển thị tối đa 18 ký tự
    doc.text(v(String(stdCode)).substring(0, 18), colX[4], y);

    doc.setTextColor(60);
    doc.text(goiYLines, colX[5], y);
    y += rowH;
  });

  y += 4;

  // ── Đánh giá học sinh (từ Sheet 2)
  if (danhGia) {
    check(30);
    doc.setDrawColor(200); doc.line(M, y, W - M, y); y += 8;
    doc.setTextColor(44, 62, 80);
    setF("bold", 12);
    doc.text("Danh Gia Hoc Sinh", M, y); y += 7;
    setF("normal", 9);
    doc.setTextColor(60);
    y = wrap(danhGia, M, y, W - M * 2, 4.5);
    y += 4;
  }

  // ── Thông tin can thiệp từ Sheet 3 (nếu có)
  const weakRow = (apiSheet3Data || []).find(r =>
    String(r["Mã HS"] || r["ma_hs"] || "") === String(maHS)
  );
  if (weakRow) {
    check(30);
    doc.setDrawColor(200); doc.line(M, y, W - M, y); y += 8;
    doc.setTextColor(44, 62, 80);
    setF("bold", 12);
    doc.text("Goi Y Can Thiep", M, y); y += 7;
    setF("normal", 9);
    doc.setTextColor(60);
    const goiY    = weakRow["Gợi ý can thiệp"] || weakRow["goi_y_can_thiep"] || "";
    const chuanYeu = weakRow["Chuẩn yếu"] || weakRow["chuan_yeu"] || "";
    if (chuanYeu) { y = wrap("Chuan yeu: " + v(chuanYeu), M, y, W - M * 2, 4.5); y += 3; }
    if (goiY)     { y = wrap(goiY, M, y, W - M * 2, 4.5); y += 4; }
  }

  // ── Footer
  const totalPages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    doc.setFontSize(7);
    doc.setTextColor(160);
    doc.text("He Thong Phan Tich Hieu Suat Hoc Tap AI - Bao Mat", M, 290);
    doc.text(`Trang ${p} / ${totalPages}`, W - M - 18, 290);
  }

  doc.save(`BaoCao_${v(maHS)}_${v(name).replace(/\s+/g, "_")}.pdf`);
}

// ── PDF legacy (fallback khi không có dữ liệu API) ───────────
function generateStudentPDFLegacy(st) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const W = doc.internal.pageSize.getWidth();
  const M = 16;
  let y = 0;

  const v = (s) => stripVN(s);
  const setF = (style, size) => { doc.setFontSize(size); doc.setFont("helvetica", style); };
  const wrap = (text, x, yy, maxW, lh) => { const lines = doc.splitTextToSize(v(text), maxW || W - M * 2); doc.text(lines, x, yy); return yy + lines.length * (lh || doc.getFontSize() * 0.4); };
  const check = (need) => { if (y + need > 280) { doc.addPage(); y = 20; } };

  // ── Header Banner
  doc.setFillColor(46, 134, 193);
  doc.rect(0, 0, W, 42, "F");
  doc.setTextColor(255, 255, 255);
  setF("bold", 18);
  doc.text("Bao Cao Ket Qua Hoc Sinh", M, 16);
  setF("normal", 10);
  doc.text("He Thong Phan Tich Hieu Suat Hoc Tap AI", M, 24);
  doc.text("Ngay tao: " + new Date().toLocaleDateString("vi-VN", { day: "2-digit", month: "2-digit", year: "numeric" }), M, 32);
  y = 52;

  // ── Student Info Block
  doc.setTextColor(44, 62, 80);
  setF("bold", 15);
  doc.text(v(st.name), M, y);
  y += 8;
  setF("normal", 10);
  doc.setTextColor(100);
  const avg10 = (students.reduce((s, s2) => s + s2.score10, 0) / students.length).toFixed(2);
  doc.text(`Ma HS: ${v(st.id)}    |    Lop: ${v(st.class)}    |    Tong diem: ${st.totalScore}/52 (${st.score10}/10)`, M, y); y += 5;
  doc.text(`Muc nguy co: ${v(st.risk)}    |    Diem TB lop: ${avg10}/10    |    Xuat sac: ${st.excellentCount}  Dat: ${st.passCount}  Khong dat: ${st.failCount}`, M, y); y += 5;
  doc.text(`Chuan manh: ${v(st.strongStds.join(", ") || "—")}    |    Chuan yeu: ${v(st.weakStds.join(", ") || "—")}    |    Can cai thien: ${v(st.improveStds.join(", ") || "—")}`, M, y);
  y += 10;

  // ── Divider
  doc.setDrawColor(200); doc.line(M, y, W - M, y); y += 8;

  // ── Standards Detail Table
  doc.setTextColor(44, 62, 80);
  setF("bold", 12);
  doc.text("Chi Tiet Tung Chuan", M, y); y += 8;

  const cols = [M, M + 22, M + 38, M + 62, M + 90];
  const hdrs = ["Chuan", "Diem", "Ket qua", "% K.Dat lop", "Nhan xet"];
  doc.setFillColor(244, 246, 247);
  doc.rect(M, y - 4, W - M * 2, 8, "F");
  setF("bold", 8);
  doc.setTextColor(80);
  hdrs.forEach((h, i) => doc.text(h, cols[i], y));
  y += 7;

  const stats = computeStandardStats(students);
  setF("normal", 8);
  st.scores.forEach((score, i) => {
    check(7);
    const cls = classify(score);
    const clsLabel = cls === "fail" ? "Khong dat" : cls === "pass" ? "Dat" : "Xuat sac";
    const failPct = stats[i] ? stats[i].failPct : 0;
    let comment = "";
    if (score === 4) comment = "Thanh thao tot";
    else if (score <= 1 && failPct > 40) comment = "Chuan kho (ca lop)";
    else if (score <= 1) comment = "Can chu y ngay";
    else comment = "Can cung co them";

    doc.setTextColor(44, 62, 80);
    doc.text(STANDARDS[i] || "C" + (i + 1), cols[0], y);
    doc.text(String(score) + " / 4", cols[1], y);

    if (cls === "fail") doc.setTextColor(231, 76, 60);
    else if (cls === "pass") doc.setTextColor(243, 156, 18);
    else doc.setTextColor(39, 174, 96);
    doc.text(clsLabel, cols[2], y);

    doc.setTextColor(100);
    doc.text(failPct + "%", cols[3], y);
    doc.text(comment, cols[4], y);
    y += 6;
  });
  y += 6;

  // ── Evaluation
  check(30);
  doc.setDrawColor(200); doc.line(M, y, W - M, y); y += 8;
  doc.setTextColor(44, 62, 80);
  setF("bold", 12);
  doc.text("Danh Gia Hoc Sinh", M, y); y += 8;
  setF("normal", 9);
  doc.setTextColor(60);

  let evalText = v(st.evaluation) + " ";
  if (st.weakStds.length > 0) evalText += `Hoc sinh con yeu o: ${st.weakStds.join(", ")}. `;
  if (st.strongStds.length > 0) evalText += `The hien tot o: ${st.strongStds.join(", ")}. `;
  if (st.risk === "Cao") evalText += "Can can thiep ngay voi cac buoi on tap nen tang.";
  else if (st.risk === "Trung binh") evalText += "Nen tang cuong luyen tap va bai tap bo sung.";
  else evalText += "Tiep tuc chien luoc hoc tap hien tai va phat trien nang cao.";
  y = wrap(evalText, M, y, W - M * 2, 4.5);
  y += 6;

  // ── Roadmap
  check(55);
  doc.setDrawColor(200); doc.line(M, y, W - M, y); y += 8;
  doc.setTextColor(44, 62, 80);
  setF("bold", 12);
  doc.text("Lo Trinh Cai Thien", M, y); y += 10;

  const roadmap = [
    { wk: "Tuan 1", title: "On Tap Nen Tang", desc: "Xem lai cac kien thuc co ban o chuan yeu: " + (st.weakStds.join(", ") || "on tap tong quat") + "." },
    { wk: "Tuan 2", title: "Luyen Tap Co Muc Tieu", desc: "Bai tap tap trung vao cac chuan can cai thien: " + (st.improveStds.join(", ") || "luyen tap toan dien") + "." },
    { wk: "Tuan 3", title: "Kiem Tra Thu", desc: "Mo phong dieu kien kiem tra thuc te de xac dinh cac diem con thieu hut." },
    { wk: "Tuan 4", title: "Danh Gia Lai", desc: "Kiem tra cuoi de do luong tien bo va dieu chinh lo trinh hoc tap." },
  ];

  roadmap.forEach((item, idx) => {
    check(18);
    doc.setFillColor(46, 134, 193);
    doc.circle(M + 3, y, 2, "F");
    if (idx < roadmap.length - 1) { doc.setDrawColor(46, 134, 193); doc.line(M + 3, y + 2, M + 3, y + 14); }
    setF("bold", 10);
    doc.setTextColor(44, 62, 80);
    doc.text(`${item.wk} - ${item.title}`, M + 10, y + 1);
    setF("normal", 8);
    doc.setTextColor(100);
    y = wrap(item.desc, M + 10, y + 6, W - M * 2 - 10, 3.8);
    y += 6;
  });

  // ── Footer
  const pages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    doc.setFontSize(7);
    doc.setTextColor(160);
    doc.text("He Thong Phan Tich Hieu Suat Hoc Tap AI - Bao Mat", M, 290);
    doc.text(`Trang ${p} / ${pages}`, W - M - 18, 290);
  }

  doc.save(`BaoCao_${v(st.id)}_${v(st.name).replace(/\s+/g, "_")}.pdf`);
}

// ══════════════════════════════════════════════════════════════
// EXPORT: Download All Student PDFs
// ══════════════════════════════════════════════════════════════
document.getElementById("btn-export-all-pdf").addEventListener("click", () => {
  const list = apiSheet2Data.length ? apiSheet2Data : students;
  if (!list.length) return;
  let idx = 0;
  const total = list.length;

  function next() {
    if (idx >= total) {
      alert(`✅ Đã tạo ${total} báo cáo PDF.`);
      return;
    }
    if (apiSheet2Data.length) {
      generateStudentPDFFromAPI(list[idx]);
    } else {
      generateStudentPDFLegacy(list[idx]);
    }
    idx++;
    setTimeout(next, 350);
  }
  next();
});

// ─── INIT ───────────────────────────────────────────────────
showPage("upload");

