// Hanger Automation Flow Demo - Client Logic with Real Backend Integration
document.addEventListener("DOMContentLoaded", () => {
  // Elements
  const btnLoadPreset = document.getElementById("btnLoadPreset");
  const btnRunPipeline = document.getElementById("btnRunPipeline");
  const actionHint = document.getElementById("actionHint");

  // Real File inputs
  const realInputOrder = document.getElementById("realInputOrder");
  const realInputSOF = document.getElementById("realInputSOF");
  const realInputRule = document.getElementById("realInputRule");

  // Dropzones
  const dropzoneOrder = document.getElementById("dropzoneOrder");
  const dropzoneSOF = document.getElementById("dropzoneSOF");
  const dropzoneRule = document.getElementById("dropzoneRule");
  const orderFileName = document.getElementById("orderFileName");
  const orderFileMeta = document.getElementById("orderFileMeta");
  const sofFileName = document.getElementById("sofFileName");
  const sofFileMeta = document.getElementById("sofFileMeta");
  const ruleFileName = document.getElementById("ruleFileName");
  const ruleFileMeta = document.getElementById("ruleFileMeta");

  // Progress bar
  const execProgressWrap = document.getElementById("execProgressWrap");
  const progressStepText = document.getElementById("progressStepText");
  const progressTimeText = document.getElementById("progressTimeText");
  const progressFill = document.getElementById("progressFill");

  // Flow steps
  const flowSteps = document.querySelectorAll(".flow-step");
  const stepBadge = document.getElementById("stepBadge");
  const stepTitle = document.getElementById("stepTitle");
  const stepGoal = document.getElementById("stepGoal");
  const stepLogic = document.getElementById("stepLogic");
  const stepCodePreview = document.getElementById("stepCodePreview");

  // Rule Tester
  const ruleSearchInput = document.getElementById("ruleSearchInput");
  const btnQuickTest = document.getElementById("btnQuickTest");
  const presetTags = document.querySelectorAll(".preset-tag");
  const resCategory = document.getElementById("resCategory");
  const resSheet = document.getElementById("resSheet");
  const resExplanation = document.getElementById("resExplanation");
  const resSpecialNote = document.getElementById("resSpecialNote");
  const toggleRulesTable = document.getElementById("toggleRulesTable");
  const rulesTableContainer = document.querySelector(".rules-table-collapsible");
  const rulesTableBody = document.getElementById("rulesTableBody");

  // Helper: Trạng thái NO để blank luôn
  const cleanVal = (v) => (!v || String(v).trim().toUpperCase() === "NO" ? "" : String(v).trim());

  // Results & Table
  const resultsTableBody = document.getElementById("resultsTableBody");
  const filterBtns = document.querySelectorAll(".filter-btn");
  const statTotal = document.getElementById("statTotal");
  const statInputFile = document.getElementById("statInputFile");
  const statMatched = document.getElementById("statMatched");
  const statMatchedSub = document.getElementById("statMatchedSub");
  const statReview = document.getElementById("statReview");
  const statReviewSub = document.getElementById("statReviewSub");
  const countAll = document.getElementById("countAll");
  const countMatched = document.getElementById("countMatched");
  const countReview = document.getElementById("countReview");
  const btnDownloadCleanExcel = document.getElementById("btnDownloadCleanExcel");
  const btnDownloadRealExcel = document.getElementById("btnDownloadRealExcel");
  const btnExportJSON = document.getElementById("btnExportJSON");
  const btnExportCSV = document.getElementById("btnExportCSV");

  // Modal
  const modalBackdrop = document.getElementById("modalBackdrop");
  const btnCloseModal = document.getElementById("btnCloseModal");
  const btnCloseModalBtn = document.getElementById("btnCloseModalBtn");
  const modalStatusBadge = document.getElementById("modalStatusBadge");
  const modalRowTitle = document.getElementById("modalRowTitle");
  const modalBody = document.getElementById("modalBody");

  // State
  let userOrderFile = null;
  let userSOFFiles = [];
  let userRuleFile = null;
  let usePreset = false; // Mặc định không có file nào được nạp trước
  let currentFilter = "all";
  let displayedRows = [];
  let lastCleanDownloadUrl = "";
  let lastCleanFileName = "";

  // Step definitions
  const STEP_DETAILS = {
    1: {
      badge: "Trạm 1: Tiếp Nhận & Khớp File",
      title: "Kiểm Tra File Đơn Hàng & Liên Kết SOForm Theo Account",
      goal: "Worker n8n nhận file đơn hàng (.xlsx / .docx / .pdf) và nhóm file SOForm. Mỗi dòng đơn hàng có cột Account (ví dụ H040M, K413M...). Worker tìm chính xác file SOF có tên chứa Account đó. Nếu không khớp bất kỳ file SOF nào, dòng được chuyển thẳng vào REVIEW để nhân viên kiểm duyệt, không lấy SOF của Account khác thay thế.",
      logic: "Đơn hàng Word/PDF được chuyển thành bảng Excel chuẩn trước khi đối soát: chỉ ghi giá trị file có ghi, cột nào thiếu để trống. Không sửa file gốc.",
      code: `// Quy tắc nhận diện Account SOF:
function matchSOF(account, uploadedFiles) {
  const matched = uploadedFiles.find(f => f.name.includes(account));
  if (!matched) {
    return { status: "REVIEW", reason: \`No SOF found for \${account}\` };
  }
  return { status: "VALID", file: matched.name };
}`
    },
    2: {
      badge: "Trạm 2: Phân Loại Nhóm Sản Phẩm",
      title: "Đối Chiếu Mô Tả Sản Phẩm Với File Rule_cho_Hanger",
      goal: "Từ trường Product Description (ví dụ: 'BODYSUIT PANT SET', '6PK CREW SOCK', 'FLEECE PANT'), hệ thống tra cứu theo từ điển quy tắc Rule_cho_Hanger.docx để xác định Product Category. Mỗi Category là tên của một Sheet trong file SOForm.",
      logic: "Khách hàng có quy ước đặc thù: Romper 1 món xếp TOPS, Romper 2PK xếp COVERALLS. Mô tả có đuôi SET hoặc dấu '+' nối nhiều món luôn xếp SETS. Không áp dụng suy đoán thông thường ngoài tài liệu rule.",
      code: `// Trích xuất nhóm theo từ điển Rule_cho_Hanger:
if (desc.endsWith("SET") || desc.includes(" SET ")) return "SETS";
if (desc.includes("ROMPER") && desc.includes("2PK")) return "COVERALLS";
if (desc.includes("ROMPER")) return "TOPS";
if (/PANT|SHORT|LEGGING|JOGGER/.test(desc)) return "BOTTOMS";
if (/TEE|TOP|POLO|SHIRT|HOODY/.test(desc)) return "TOPS";`
    },
    3: {
      badge: "Trạm 3: Động Cơ Rule Cố Định (Deterministic)",
      title: "Quét Bộ Quy Tắc Đã Phê Duyệt Trước (hanger_rules.json)",
      goal: "Hệ thống kiểm tra xem mã Ref# và Style của đơn hàng đã nằm trong danh mục các Rule cố định đã được kiểm duyệt hay chưa. Nếu đã có rule chuẩn, áp dụng trực tiếp 100% chính xác mà không cần gọi đến mô hình AI.",
      logic: "Quy tắc đã duyệt luôn chạy trước. Tiết kiệm thời gian xử lý và đảm bảo tính nhất quán tuyệt đối cho các mã hàng quen thuộc được tái đặt hàng thường xuyên.",
      code: `// Kiểm tra rule cứng đã phê duyệt:
const rule = HANGER_RULES[account]?.find(r => 
  r.match.ref_number === row.ref_number && 
  (r.match.style === row.style || r.match.style_any?.includes(row.style))
);
if (rule) {
  return { status: "MATCHED", result: rule.result, source: rule.source };
}`
    },
    4: {
      badge: "Trạm 4: Tra Cứu SOF & Kiểm Chứng Tọa Độ Ô",
      title: "Đọc Tọa Độ Ô Trong Sheet SOF & Xác Minh Chứng Cứ",
      goal: "Đối với các dòng cần tra cứu từ SOF (thủ công hoặc hỗ trợ bởi DeepSeek AI), hệ thống mở đúng Sheet đã tìm ở Trạm 2, định vị dải Size (Newborn, Infant, Toddler, 4-6x, 7-16...) và kiểu đóng gói (Hang hay Flat).",
      logic: "Trích xuất từ 1 đến tối đa 3 vùng ô (ví dụ: A26:C26, B37:B37) chứa cả mã hanger, màu sắc và sizer. AI chỉ được chấp nhận nếu trích dẫn chính xác tọa độ và nguyên văn câu chữ trong SOF.",
      code: `// Xác minh bằng chứng trích dẫn:
if (row.hang_flat === "Hang") {
  if (!hanger_code || hanger_code === "NO") {
    return { status: "REVIEW", reason: "Hang requires valid hanger code & color" };
  }
}
// Kiểm tra tọa độ ô nguồn có tồn tại trong Sheet:
validateCellRanges(sourceCells, targetSheet);`
    },
    5: {
      badge: "Trạm 5: Phân Định Trạng Thái & Xuất Bảng Chuẩn",
      title: "Phân Loại MATCHED / REVIEW & Xuất Bảng Màu Tổng Hợp Hanger",
      goal: "Chỉ các dòng đạt tiêu chuẩn MATCHED (đầy đủ mã hanger, màu sắc, bằng chứng tọa độ ô) mới được ghi vào đơn hàng. Các dòng thiếu chứng cứ hoặc có bất đồng (Hang/Flat mâu thuẫn) được đưa sang sheet HANGER REVIEW.",
      logic: "Bảng kết quả được bố trí theo định dạng TỔNG HỢP HANGER: Khối màu Cyan (cột Đơn Hàng), Khối màu Vàng (cột Thuộc Tính SOF), Khối màu Xám (cột Audit).",
      code: `// Kết quả xuất bảng TỔNG HỢP HANGER:
Columns: [
  ...ORDER_COLUMNS,    // Cyan Fill
  ...SOF_COLUMNS,      // Yellow Fill (Hanger code, color, sizer, sticker)
  ...AUDIT_COLUMNS     // Grey Fill (Status, Source File, Source Sheet, Cells)
]`
    }
  };

  // 1. Initialize Rules Table
  function renderRulesTable() {
    if (!window.DEMO_RULES || !rulesTableBody) return;
    rulesTableBody.innerHTML = window.DEMO_RULES.map(rule => `
      <tr>
        <td>${rule.category}</td>
        <td>${rule.description.replace(/\n/g, "<br>")}</td>
      </tr>
    `).join("");
  }

  // 2. Real Upload Dropzone Setup
  function setupDropzones() {
    // Dropzone 1: Order
    dropzoneOrder.addEventListener("click", () => realInputOrder.click());
    realInputOrder.addEventListener("change", (e) => {
      if (e.target.files && e.target.files[0]) {
        userOrderFile = e.target.files[0];
        usePreset = false;
        dropzoneOrder.classList.add("loaded");
        orderFileName.textContent = userOrderFile.name;
        orderFileMeta.textContent = `Tệp thật: ${(userOrderFile.size / 1024).toFixed(1)} KB • Sẵn sàng xử lý`;
        updateActionHint();
      }
    });

    // Dropzone 2: SOF
    dropzoneSOF.addEventListener("click", () => realInputSOF.click());
    realInputSOF.addEventListener("change", (e) => {
      if (e.target.files && e.target.files.length > 0) {
        userSOFFiles = Array.from(e.target.files);
        usePreset = false;
        dropzoneSOF.classList.add("loaded");
        if (userSOFFiles.length === 1) {
          sofFileName.textContent = userSOFFiles[0].name;
          sofFileMeta.textContent = `1 tệp SOForm thật (${(userSOFFiles[0].size / 1024).toFixed(1)} KB)`;
        } else {
          sofFileName.textContent = `Đã chọn ${userSOFFiles.length} file SOForm`;
          sofFileMeta.textContent = userSOFFiles.map(f => f.name).join(", ").slice(0, 50) + "...";
        }
        updateActionHint();
      }
    });

    // Dropzone 3: Rule
    dropzoneRule.addEventListener("click", () => realInputRule.click());
    realInputRule.addEventListener("change", (e) => {
      if (e.target.files && e.target.files[0]) {
        userRuleFile = e.target.files[0];
        usePreset = false;
        dropzoneRule.classList.add("loaded");
        ruleFileName.textContent = userRuleFile.name;
        ruleFileMeta.textContent = `Quy tắc tùy chỉnh: ${(userRuleFile.size / 1024).toFixed(1)} KB`;
      }
    });

    // Drag and drop events
    [dropzoneOrder, dropzoneSOF, dropzoneRule].forEach(dz => {
      dz.addEventListener("dragover", (e) => {
        e.preventDefault();
        dz.classList.add("dragover");
      });
      dz.addEventListener("dragleave", () => dz.classList.remove("dragover"));
      dz.addEventListener("drop", (e) => {
        e.preventDefault();
        dz.classList.remove("dragover");
      });
    });

    dropzoneOrder.addEventListener("drop", (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        userOrderFile = e.dataTransfer.files[0];
        usePreset = false;
        dropzoneOrder.classList.add("loaded");
        orderFileName.textContent = userOrderFile.name;
        orderFileMeta.textContent = `Tệp thật: ${(userOrderFile.size / 1024).toFixed(1)} KB`;
        updateActionHint();
      }
    });

    dropzoneSOF.addEventListener("drop", (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        userSOFFiles = Array.from(e.dataTransfer.files);
        usePreset = false;
        dropzoneSOF.classList.add("loaded");
        sofFileName.textContent = userSOFFiles[0].name;
        sofFileMeta.textContent = `Tệp SOF thật: ${(userSOFFiles[0].size / 1024).toFixed(1)} KB`;
        updateActionHint();
      }
    });

    dropzoneRule.addEventListener("drop", (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        userRuleFile = e.dataTransfer.files[0];
        usePreset = false;
        dropzoneRule.classList.add("loaded");
        ruleFileName.textContent = userRuleFile.name;
        ruleFileMeta.textContent = `Tệp Rule thật: ${(userRuleFile.size / 1024).toFixed(1)} KB`;
      }
    });
  }

  // 3. Preset Loader (Nạp Mẫu Dự Án)
  btnLoadPreset.addEventListener("click", () => {
    loadProjectPresets();
  });

  function loadProjectPresets() {
    usePreset = true;
    userOrderFile = null;
    userSOFFiles = [];
    userRuleFile = null;

    dropzoneOrder.classList.add("loaded");
    orderFileName.textContent = "VLK VLH LPO 09.09.26.XLS";
    orderFileMeta.textContent = "Đơn hàng đầy đủ • Xử lý toàn bộ file";

    dropzoneSOF.classList.add("loaded");
    sofFileName.textContent = "H040M Stock Replenishment SO Form 8.31.26.xlsx";
    sofFileMeta.textContent = "Khớp mã Account H040M • 14.6 MB";

    dropzoneRule.classList.add("loaded");
    ruleFileName.textContent = "Rule_cho_Hanger.docx";
    ruleFileMeta.textContent = "16 nhóm Category chuẩn mực & hanger_rules.json";

    actionHint.innerHTML = "✅ <strong>Đã nạp 3 file mẫu chuẩn!</strong> Nhấn nút xanh để chạy thử ngay.";
  }

  function resetDropzonesToEmpty() {
    usePreset = false;
    userOrderFile = null;
    userSOFFiles = [];
    userRuleFile = null;

    // Reset Order Dropzone (Không có file nào - để trống)
    dropzoneOrder.classList.remove("loaded");
    orderFileName.textContent = "Chọn hoặc kéo thả file đơn hàng";
    orderFileMeta.textContent = "Nhấp vào đây để duyệt file (.xlsx, .xls, .docx, .pdf)";
    if (realInputOrder) realInputOrder.value = "";

    // Reset SOF Dropzone (Không có file nào - để trống hoàn toàn)
    dropzoneSOF.classList.remove("loaded");
    sofFileName.textContent = "Chọn hoặc kéo thả file SOForm";
    sofFileMeta.textContent = "Nhấp vào đây để duyệt file (.xlsx - tối đa 20 file)";
    if (realInputSOF) realInputSOF.value = "";

    // Rule Dropzone (Tích hợp sẵn rule chuẩn)
    dropzoneRule.classList.add("loaded");
    ruleFileName.textContent = "Rule_cho_Hanger.docx (Hệ thống)";
    ruleFileMeta.textContent = "17 nhóm Category chuẩn & hanger_rules.json (Đã tích hợp sẵn)";
    if (realInputRule) realInputRule.value = "";

    updateActionHint();
  }

  function updateActionHint() {
    if (userOrderFile && userSOFFiles && userSOFFiles.length > 0) {
      actionHint.innerHTML = "✅ <strong>Đã chọn đủ 2 tệp!</strong> Nhấn nút xanh để chạy tra mã Hanger.";
    } else if (userOrderFile) {
      actionHint.innerHTML = "📄 <strong>Đã chọn File Đơn Hàng.</strong> Vui lòng chọn tiếp File SOForm (Tệp 2).";
    } else if (userSOFFiles && userSOFFiles.length > 0) {
      actionHint.innerHTML = "📊 <strong>Đã chọn File SOForm.</strong> Vui lòng chọn thêm File Đơn Hàng (Tệp 1).";
    } else {
      actionHint.innerHTML = "📌 <strong>Mặc định chưa chọn file:</strong> Hãy chọn File Đơn Hàng & SOForm ở trên để bắt đầu tra mã.";
    }
  }

  // 4. Stepper Diagram Click
  flowSteps.forEach(step => {
    step.addEventListener("click", () => {
      setActiveStep(step.getAttribute("data-step"));
    });
  });

  function setActiveStep(stepNum) {
    flowSteps.forEach(s => s.classList.remove("active"));
    const activeEl = document.querySelector(`.flow-step[data-step="${stepNum}"]`);
    if (activeEl) activeEl.classList.add("active");

    const data = STEP_DETAILS[stepNum];
    if (data) {
      stepBadge.textContent = data.badge;
      stepTitle.textContent = data.title;
      stepGoal.textContent = data.goal;
      stepLogic.innerHTML = `<strong>Quy tắc an toàn:</strong> ${data.logic}`;
      stepCodePreview.innerHTML = `<code>${data.code}</code>`;
    }
  }

  // 5. Keyword Tester
  function testRuleKeyword(query) {
    const q = (query || "").trim().toUpperCase();
    if (!q) return;

    let matchedCat = "OTHER";
    let matchedSheet = "OTHER";
    let explanation = "Không tìm thấy từ khóa đặc thù trong quy tắc mẫu. Dòng này sẽ được chuyển sang REVIEW.";
    let specialNote = "💡 Cần đối soát thêm với tài liệu kỹ thuật của khách hàng.";

    if (q.endsWith("SET") || q.includes(" SET ") || (q.includes("+") && !q.includes("+ ATTACHMENT"))) {
      matchedCat = "SETS";
      matchedSheet = "SETS";
      explanation = "Bộ nhiều món (track set, short set, pant set, 2PC/3PC set). Mô tả kết thúc bằng 'SET' hoặc nối bằng '+' luôn xếp vào SETS.";
      specialNote = "💡 Ví dụ: BODYSUIT PANT SET, 3PC PANT SET, FLEECE PANT SET.";
    } else if (q.includes("2PK") && (q.includes("ROMPER") || q.includes("COVERALL"))) {
      matchedCat = "COVERALLS";
      matchedSheet = "COVERALLS";
      explanation = "Coverall hoặc Bodysuit/Romper đóng gói nhiều cái (multi-pack, 2PK, 3PC creepers) được xếp vào nhóm COVERALLS.";
      specialNote = "💡 Điểm lưu ý cốt lõi: Romper 2PK xếp COVERALLS, nhưng Romper 1 món lại xếp TOPS!";
    } else if (q.includes("ROMPER") || q.includes("BODYSUIT") || /TEE|TOP|POLO|SHIRT|HOODY|HOODIE|SWEATSHIRT/.test(q)) {
      matchedCat = "TOPS";
      matchedSheet = "TOPS";
      explanation = "Áo 1 món: tee, top, polo, shirt, hoodie, sweatshirt. Romper 1 món (1pc) và bodysuit 1 món cũng thuộc nhóm TOPS.";
      specialNote = "💡 Mô tả có '+ ATTACHMENT' vẫn xếp theo món chính (ví dụ: ROMPER + ATTACHMENT vẫn thuộc TOPS).";
    } else if (/PANT|SHORT|LEGGING|JOGGER|SKIRT|SKORT/.test(q)) {
      matchedCat = "BOTTOMS";
      matchedSheet = "BOTTOMS";
      explanation = "Quần, short, legging, jogger, váy 1 món; đồ bơi bé trai (boardshorts).";
      specialNote = "💡 Thuộc sheet BOTTOMS trong SOForm.";
    } else if (/SOCK|CREW|ANKLE|HOSIERY/.test(q)) {
      matchedCat = "HOSIERY";
      matchedSheet = "HOSIERY";
      explanation = "Tất/vớ: socks, crew, quarter, no show, ankle (thường đóng gói 6PK, 8PK); bib, blanket.";
      specialNote = "💡 Hầu hết sản phẩm HOSIERY đóng gói dạng Flat (gấp phẳng), hanger code ghi NO.";
    } else if (/JACKET|COAT|VEST|OUTERWEAR/.test(q)) {
      matchedCat = "OUTERWEAR";
      matchedSheet = "OUTERWEAR";
      explanation = "Áo khoác: jacket, coat, vest, outerwear.";
      specialNote = "💡 Thuộc sheet OUTERWEAR trong file SOForm.";
    } else if (/BOX SET/.test(q)) {
      matchedCat = "BOX SETS";
      matchedSheet = "BOX SETS";
      explanation = "Set đóng hộp (bộ quà tặng đóng sẵn trong hộp).";
      specialNote = "💡 Thường đi theo quy định đóng hộp quà tặng.";
    } else if (/UNDERWEAR|BRIEF|BOXER|UNDERSHIRT/.test(q)) {
      matchedCat = "UNDERWEAR";
      matchedSheet = "UNDERWEAR";
      explanation = "Đồ lót: underwear, brief, boxer, undershirt.";
      specialNote = "💡 Thuộc sheet UNDERWEAR trong file SOForm.";
    } else if (/BEANIE|HAT|CAP|GLOVES|MITTENS|SCARF|COLD WEATHER/.test(q)) {
      matchedCat = "COLD WEATHER";
      matchedSheet = "COLD WEATHER";
      explanation = "Mũ len, nón, găng tay, khăn quàng cổ (Beanie, hat, cap, gloves, mittens, scarf).";
      specialNote = "💡 Tên sheet trong SOF chính xác là 'COLD WEATHER'.";
    } else if (/BACKPACK|LUNCH TOTE|SLING|CROSSBODY|DUFFLE|SHOE STRINGS/.test(q)) {
      matchedCat = "ACCESSORIES / BAGS";
      matchedSheet = "ACCESSORIES (hoặc BAGS)";
      explanation = "Túi, balo, phụ kiện. Riêng account K413M xếp vào sheet ACCESSORIES; các account khác xếp vào BAGS.";
      specialNote = "💡 Chú ý account K413M có sheet ACCESSORIES riêng biệt.";
    } else if (/BAG/.test(q)) {
      matchedCat = "BAGS";
      matchedSheet = "BAGS";
      explanation = "Túi, balo: bag, backpack.";
      specialNote = "💡 Thuộc sheet BAGS trong file SOForm.";
    } else if (/SWIMWEAR|SWIMSUIT|BOARDSHORT/.test(q)) {
      matchedCat = "GIRLS SWIMWEAR / 1 PC / 2-3 PC SWIMWEAR";
      matchedSheet = "GIRLS SWIMWEAR";
      explanation = "Đồ bơi: H040M có sheet 'GIRLS SWIMWEAR'; K816M & N748M chia thành '1 PC SWIMWEAR' và '2-3 PC SWIMWEAR'.";
      specialNote = "💡 Boys swimsuits/boardshorts 1 món vẫn xếp BOTTOMS theo quy ước.";
    } else if (/SWEATER|SWEATER-YARN/.test(q)) {
      matchedCat = "SWEATER-YARN";
      matchedSheet = "SWEATER-YARN";
      explanation = "Áo len, sweater dệt sợi (SOF H040M).";
      specialNote = "💡 Sheet SWEATER-YARN trong file SOForm.";
    } else if (/HUGGIES/.test(q)) {
      matchedCat = "HUGGIES";
      matchedSheet = "HUGGIES";
      explanation = "Hàng thương hiệu Huggies (SOF H040M).";
      specialNote = "💡 Thuộc sheet HUGGIES riêng biệt.";
    } else if (/DRESS/.test(q)) {
      matchedCat = "DRESS (chưa có sheet)";
      matchedSheet = "CẦN HỎI KHÁCH";
      explanation = "Dress, dress + attachment. Các file SOF hiện tại chưa có sheet riêng cho váy liền.";
      specialNote = "⚠️ Bắt buộc chuyển sang REVIEW để xác nhận với khách trước khi xử lý, không được tự ý gán vào TOPS!";
    }

    resCategory.textContent = matchedCat;
    resSheet.textContent = matchedSheet;
    resExplanation.textContent = explanation;
    resSpecialNote.innerHTML = specialNote;
  }

  btnQuickTest.addEventListener("click", () => testRuleKeyword(ruleSearchInput.value));
  ruleSearchInput.addEventListener("keyup", (e) => {
    if (e.key === "Enter") testRuleKeyword(ruleSearchInput.value);
  });
  presetTags.forEach(tag => {
    tag.addEventListener("click", () => {
      const q = tag.getAttribute("data-query");
      ruleSearchInput.value = q;
      testRuleKeyword(q);
    });
  });
  toggleRulesTable.addEventListener("click", () => rulesTableContainer.classList.toggle("open"));

  // 6. REAL PIPELINE EXECUTION (Gửi file thật lên Backend Python)
  btnRunPipeline.addEventListener("click", () => {
    runRealPipeline();
  });

  async function runRealPipeline() {
    if (!usePreset && !userOrderFile) {
      alert("Vui lòng chọn File Đơn Hàng (Tệp 1) để chạy thực tế, hoặc nhấn 'Nạp 3 Tệp Mẫu Dự Án' ở góc trên để chạy thử!");
      return;
    }
    if (!usePreset && (!userSOFFiles || userSOFFiles.length === 0)) {
      alert("Vui lòng chọn ít nhất 1 File SOForm (Tệp 2) để hệ thống có tài liệu đối soát!");
      return;
    }

    btnRunPipeline.disabled = true;
    btnRunPipeline.innerHTML = `
      <span class="spinner" style="display:inline-block; animation:spin 1s linear infinite; margin-right:8px;">⚙️</span>
      Đang gọi Python Worker Backend...
    `;

    execProgressWrap.style.display = "block";
    progressFill.style.width = "20%";
    progressStepText.textContent = "Trạm 1: Tiếp nhận và nạp tệp vào bộ nhớ đệm...";
    progressTimeText.textContent = "0.2s";

    const formData = new FormData();
    formData.append("use_preset", usePreset ? "true" : "false");

    if (userOrderFile) {
      formData.append("order_file", userOrderFile);
    }
    if (userSOFFiles && userSOFFiles.length > 0) {
      userSOFFiles.forEach(f => formData.append("sof_files", f));
    }
    if (userRuleFile) {
      formData.append("rule_file", userRuleFile);
    }

    let startTime = performance.now();
    let timer = setInterval(() => {
      let elapsed = ((performance.now() - startTime) / 1000).toFixed(1);
      progressTimeText.textContent = `${elapsed}s`;
    }, 100);

    // Step animation
    setTimeout(() => {
      setActiveStep(2);
      progressFill.style.width = "45%";
      progressStepText.textContent = "Trạm 2: Quét mô tả hàng & Phân loại Category...";
    }, 300);

    setTimeout(() => {
      setActiveStep(3);
      progressFill.style.width = "65%";
      progressStepText.textContent = "Trạm 3: Đối chiếu Rule cứng & Tra cứu SOForm...";
    }, 600);

    setTimeout(() => {
      setActiveStep(4);
      progressFill.style.width = "85%";
      progressStepText.textContent = "Trạm 4: Trích dẫn tọa độ ô & Đóng gói dữ liệu kết quả...";
    }, 900);

    try {
      const response = await fetch("/api/process", {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        let errText = await response.text();
        if (errText.includes("<title>502") || errText.includes("Bad gateway") || response.status === 502) {
          throw new Error("Máy chủ vừa khởi động lại hoặc đường truyền mạng bị gián đoạn (502 Bad Gateway). Vui lòng bấm 'Chạy lại Pipeline' ngay để thử lại.");
        }
        if (errText.includes("<html") || errText.includes("<!DOCTYPE")) {
          throw new Error(`Lỗi kết nối từ cổng mạng Cloudflare (Mã HTTP ${response.status}). Vui lòng bấm chạy lại.`);
        }
        throw new Error(errText || "Lỗi khởi tạo xử lý file từ server");
      }

      const initData = await response.json();
      const jobId = initData.job_id;

      if (!jobId) {
        throw new Error("Không nhận được mã tiến trình xử lý từ server");
      }

      // Poll status every 800ms (an toàn 100% không lo Cloudflare timeout)
      let jobResult = null;
      while (true) {
        await new Promise(r => setTimeout(r, 800));
        let pollRes;
        try {
          pollRes = await fetch(`/api/job-status?job_id=${encodeURIComponent(jobId)}`);
        } catch (pollErr) {
          console.warn("Poll retry:", pollErr);
          continue;
        }

        if (!pollRes.ok) continue;

        let jobState;
        try {
          jobState = await pollRes.json();
        } catch (e) {
          continue;
        }

        if (jobState.status === "COMPLETED") {
          jobResult = jobState.result;
          break;
        } else if (jobState.status === "ERROR") {
          throw new Error(jobState.error || "Lỗi xử lý file từ server");
        }
      }

      clearInterval(timer);
      const res = jobResult;

      // Finish step
      setActiveStep(5);
      progressFill.style.width = "100%";
      const llmBadge = (res.llm_info && res.llm_info.enabled) ? ` (Động cơ: Rules + DeepSeek ${res.llm_info.model || 'Flash'})` : " (Động cơ: Rules Engine)";
      progressStepText.textContent = `✅ Hoàn thành xuất sắc sau ${res.duration_seconds}s!${llmBadge}`;
      let finalElapsed = ((performance.now() - startTime) / 1000).toFixed(2);
      progressTimeText.textContent = `${res.duration_seconds || finalElapsed}s`;

      // Update KPI Statistics
      statTotal.textContent = Number(res.total_order_rows).toLocaleString();
      statInputFile.textContent = `Tệp: ${res.input_file}`;
      statMatched.textContent = Number(res.matched).toLocaleString();
      statReview.textContent = Number(res.review).toLocaleString();

      const ruleCount = res.rule_matched_count || 0;
      const llmCount = res.llm_matched_count || 0;
      statMatchedSub.textContent = `${ruleCount} dòng qua Rule cứng • ${llmCount} dòng qua DeepSeek LLM`;
      statReviewSub.textContent = `Bảo vệ an toàn chuyền may (chặn rủi ro đoán mò)`;

      // Real Excel Download button (Kèm Audit)
      if (res.result_download_url) {
        btnDownloadRealExcel.style.display = "inline-flex";
        btnDownloadRealExcel.href = res.result_download_url;
        btnDownloadRealExcel.download = res.result_file_name || "TONG_HOP_HANGER_KETQUA.xlsx";
        btnDownloadRealExcel.innerHTML = `
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          Tải Excel Đầy Đủ (Kèm Audit)
        `;
      }

      if (res.clean_download_url) {
        lastCleanDownloadUrl = res.clean_download_url;
        lastCleanFileName = res.clean_file_name || "TONG_HOP_HANGER_XUONG_GON.xlsx";
      }

      // Render Table Rows with real results or fallback to samples if empty
      const rowsToRender = (res.display_rows && res.display_rows.length > 0) ? res.display_rows : (window.DEMO_ROWS || []);
      renderResultsTable(rowsToRender);

      // Scroll to result
      setTimeout(() => {
        document.getElementById("resultsSection").scrollIntoView({ behavior: "smooth" });
      }, 400);

    } catch (err) {
      clearInterval(timer);
      console.error(err);
      progressStepText.textContent = `❌ Lỗi: ${err.message}`;
      alert(`Không thể xử lý file: ${err.message}`);
    } finally {
      btnRunPipeline.disabled = false;
      btnRunPipeline.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polygon points="5 3 19 12 5 21 5 3"/></svg>
        ⚡ Chạy Lại Pipeline (Chạy Thật)
      `;
    }
  }

  // 7. Results Table Handling
  function renderResultsTable(rows) {
    displayedRows = rows;
    updateCounts(rows);
    applyFilter(currentFilter);
  }

  function updateCounts(rows) {
    countAll.textContent = rows.length;
    countMatched.textContent = rows.filter(r => r.status === "MATCHED").length;
    countReview.textContent = rows.filter(r => r.status === "REVIEW").length;
  }

  function applyFilter(filter) {
    currentFilter = filter;
    filterBtns.forEach(btn => {
      btn.classList.toggle("active", btn.getAttribute("data-filter") === filter);
    });

    let filtered = displayedRows;
    if (filter === "MATCHED") filtered = displayedRows.filter(r => r.status === "MATCHED");
    else if (filter === "REVIEW") filtered = displayedRows.filter(r => r.status === "REVIEW");
    else if (filter === "Hang") filtered = displayedRows.filter(r => r.order_hang_flat === "Hang");
    else if (filter === "Flat") filtered = displayedRows.filter(r => r.order_hang_flat === "Flat");

    renderTableRows(filtered);
  }

  filterBtns.forEach(btn => {
    btn.addEventListener("click", () => applyFilter(btn.getAttribute("data-filter")));
  });

  function renderTableRows(rows) {
    if (!resultsTableBody) return;
    if (rows.length === 0) {
      resultsTableBody.innerHTML = `
        <tr>
          <td colspan="16" style="text-align:center; padding: 40px; color: var(--text-dim);">
            Chưa có dòng nào được xử lý hoặc không có dòng phù hợp bộ lọc.
          </td>
        </tr>
      `;
      return;
    }

    resultsTableBody.innerHTML = rows.map((r, index) => {
      const isMatched = r.status === "MATCHED";
      const statusClass = isMatched ? "matched" : "review";
      const isHang = r.order_hang_flat === "Hang";

      return `
        <tr>
          <!-- Cyan Order Columns -->
          <td class="cell-cyan code-cell">${r.po_number || "-"}</td>
          <td class="cell-cyan code-cell">${r.account || "-"}</td>
          <td class="cell-cyan code-cell">${r.ref_number || "-"} / ${r.style || "-"}</td>
          <td class="cell-cyan">${r.product_description || "-"}</td>
          <td class="cell-cyan"><span class="pill">${r.product_category || "-"}</span></td>
          <td class="cell-cyan" style="font-size:0.75rem;">${r.size_configuration || "-"}</td>
          <td class="cell-cyan">
            <span style="font-weight:700; color: ${isHang ? '#38bdf8' : '#a78bfa'};">
              ${r.order_hang_flat || "-"}
            </span>
          </td>

          <!-- Yellow SOF Columns (NO để Blank luôn) -->
          <td class="cell-yellow code-cell">${cleanVal(r.hanger_code)}</td>
          <td class="cell-yellow">${cleanVal(r.hanger_color)}</td>
          <td class="cell-yellow" style="font-size:0.75rem;">${cleanVal(r.color_sizer)}</td>
          <td class="cell-yellow">${cleanVal(r.sticker_hanger)}</td>
          <td class="cell-yellow">${cleanVal(r.size_sticker_hanger)}</td>
          <td class="cell-yellow" style="color:var(--text-dim);">${cleanVal(r.ncc)}</td>

          <!-- Grey Audit Columns -->
          <td class="cell-grey">
            <span class="status-badge ${statusClass}">${r.status}</span>
          </td>
          <td class="cell-grey code-cell">${r.source_sheet || "-"}</td>
          <td class="cell-grey code-cell" style="color:var(--accent-cyan);">${r.source_cells || "-"}</td>
          <td class="cell-grey">
            <button class="btn-inspect" data-index="${index}">Soi Bằng Chứng</button>
          </td>
        </tr>
      `;
    }).join("");

    document.querySelectorAll(".btn-inspect").forEach(btn => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.getAttribute("data-index"), 10);
        showEvidenceModal(rows[idx]);
      });
    });
  }

  // 8. Evidence Inspector Modal
  function showEvidenceModal(row) {
    if (!row) return;

    modalStatusBadge.textContent = row.status;
    modalStatusBadge.className = `modal-badge ${row.status.toLowerCase()}`;
    modalRowTitle.textContent = `Dòng #${row.row_number || 1} • PO ${row.po_number || "-"} • Style ${row.style || "-"}`;

    const isMatched = row.status === "MATCHED";

    modalBody.innerHTML = `
      <div class="modal-section">
        <div class="modal-section-title">1. Thông Tin Đơn Hàng Gốc</div>
        <div class="modal-grid-2">
          <div class="info-item">
            <div class="info-label">Account / PO#</div>
            <div class="info-val">${row.account || "-"} / ${row.po_number || "-"}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Ref# & Style</div>
            <div class="info-val">${row.ref_number || "-"} — ${row.style || "-"}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Mô tả sản phẩm</div>
            <div class="info-val" style="color:#7dd3fc;">${row.product_description || "-"}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Size Configuration & Hang/Flat</div>
            <div class="info-val">${row.size_configuration || "-"} (${row.order_hang_flat || "-"})</div>
          </div>
        </div>
      </div>

      <div class="modal-section">
        <div class="modal-section-title">2. Bằng Chứng Trích Xuất Từ SOForm (Ship Order Form)</div>
        <div class="evidence-box">
          <div class="evidence-header">
            <span>Tệp SOF: <strong>${row.source_file || "H040M Stock Replenishment SO Form"}</strong></span>
            <span>Sheet: <strong>${row.source_sheet || "N/A"}</strong></span>
          </div>
          <div class="evidence-content">
            <p><strong>Tọa độ ô trích dẫn:</strong> <code style="color:#38bdf8;">${row.source_cells || "Không tìm thấy ô phù hợp"}</code></p>
            <p><strong>Phương pháp khớp:</strong> <code>${row.match_method || "DETERMINISTIC_RULES"}</code></p>
            <p><strong>Độ tin cậy (Confidence):</strong> <strong>${row.confidence ? (row.confidence * 100).toFixed(0) : "100"}%</strong></p>
            ${row.validation_note ? `<p style="color:#fde047; margin-top:8px;"><strong>Ghi chú kiểm tra:</strong> ${row.validation_note}</p>` : ""}
          </div>
        </div>
      </div>

      <div class="modal-section">
        <div class="modal-section-title">3. Kết Quả Thuộc Tính Hanger Được Ghi Nhận (Trạng Thái NO Để Blank)</div>
        <div class="modal-grid-2">
          <div class="info-item">
            <div class="info-label">Mã Hanger (Hanger Code)</div>
            <div class="info-val" style="color:#fde047;">${cleanVal(row.hanger_code) || "<span style='color:var(--text-dim);'>[BLANK]</span>"}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Màu Hanger (Hanger Color)</div>
            <div class="info-val">${cleanVal(row.hanger_color) || "<span style='color:var(--text-dim);'>[BLANK]</span>"}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Color Sizer / Size Clip</div>
            <div class="info-val">${cleanVal(row.color_sizer) || "<span style='color:var(--text-dim);'>[BLANK]</span>"}</div>
          </div>
          <div class="info-item">
            <div class="info-label">Sticker Hanger</div>
            <div class="info-val">${cleanVal(row.sticker_hanger) || "<span style='color:var(--text-dim);'>[BLANK]</span>"}</div>
          </div>
        </div>
      </div>

      <div class="modal-section">
        <div class="modal-section-title" style="display:flex; justify-content:space-between; align-items:center;">
          <span>4. Định Dạng 12 Cột Chuẩn (Chuẩn Quy Ước SKILL.md — NO Để Blank)</span>
          <button id="btnCopy12Cols" class="btn btn-sm btn-outline" style="padding:2px 8px; font-size:0.75rem;">📋 Sao chép</button>
        </div>
        <div style="background:#070a12; border:1px solid #1e293b; border-radius:6px; padding:12px; font-family:var(--font-mono); font-size:0.78rem; color:#38bdf8; overflow-x:auto; white-space:nowrap;" id="skill12ColText">
${row.account || ""} | ${row.ref_number || ""} | ${row.style || ""} | ${row.product_category || ""} | ${row.order_hang_flat || ""} | ${isMatched ? cleanVal(row.hanger_code) : ""} | ${isMatched ? cleanVal(row.hanger_color) : ""} | ${isMatched ? cleanVal(row.color_sizer) : ""} | ${isMatched ? cleanVal(row.sticker_hanger) : ""} | ${isMatched ? cleanVal(row.size_sticker_hanger) : ""} | ${isMatched ? (row.source_sheet || "") : ""} | ${isMatched ? (row.source_cells || "") : ""}
        </div>
      </div>

      <div class="modal-section">
        <div class="modal-section-title">5. Đánh Giá An Toàn Chuyền May (Zero-Guessing Principle)</div>
        <div style="font-size:0.88rem; color:${isMatched ? '#34d399' : '#fbbf24'}; background:${isMatched ? 'rgba(16,185,129,0.1)' : 'rgba(245,158,11,0.1)'}; padding:14px 16px; border-radius:8px; line-height:1.6;">
          ${isMatched 
            ? "✅ <strong>ĐÃ KIỂM ĐỊNH AN TOÀN (MATCHED):</strong> Mọi giá trị hanger và màu sắc đều được đọc nguyên văn từ ô nguồn của SOForm. Dòng này đủ điều kiện ghi thẳng vào bảng đơn hàng."
            : `⚠️ <strong>CHUYỂN VÀO HANGER REVIEW:</strong> ${row.validation_note || "Cần nhân viên phụ trách đối chiếu lại với tài liệu khách hàng."}<br><span style="font-size:0.8rem; color:#fde68a;">📌 Theo nguyên tắc SKILL.md: Một dòng đoán sai đi tới tận xưởng may sẽ gây thiệt hại lớn. Để REVIEW chỉ tốn vài phút kiểm tra thủ công.</span>`}
        </div>
      </div>
    `;

    modalBackdrop.classList.add("open");

    // Copy 12 col listener
    const btnCopy = document.getElementById("btnCopy12Cols");
    if (btnCopy) {
      btnCopy.addEventListener("click", () => {
        const text = document.getElementById("skill12ColText").textContent.trim();
        navigator.clipboard.writeText(text).then(() => {
          btnCopy.textContent = "✅ Đã chép!";
          setTimeout(() => btnCopy.textContent = "📋 Sao chép", 2000);
        });
      });
    }
  }

  function closeModal() {
    modalBackdrop.classList.remove("open");
  }

  btnCloseModal.addEventListener("click", closeModal);
  btnCloseModalBtn.addEventListener("click", closeModal);
  modalBackdrop.addEventListener("click", (e) => {
    if (e.target === modalBackdrop) closeModal();
  });

  // 9. Export Clean Excel (Bỏ Cột Audit), JSON & CSV
  btnDownloadCleanExcel.addEventListener("click", async () => {
    // Nếu vừa chạy pipeline thật và có link file clean đã build sẵn:
    if (lastCleanDownloadUrl) {
      const a = document.createElement("a");
      a.href = lastCleanDownloadUrl;
      a.download = lastCleanFileName || "TONG_HOP_HANGER_XUONG_GON.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      return;
    }

    // Nếu tải trực tiếp từ các dòng đang hiển thị (kể cả mẫu thử hoặc đã filter):
    const originalHtml = btnDownloadCleanExcel.innerHTML;
    try {
      btnDownloadCleanExcel.disabled = true;
      btnDownloadCleanExcel.innerHTML = `
        <svg class="spin-slow" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 2a10 10 0 0 1 10 10"/></svg>
        Đang xuất Excel gọn...
      `;

      const rowsToSend = (displayedRows && displayedRows.length > 0) ? displayedRows : (window.DEMO_ROWS || []);
      const response = await fetch("/api/export-clean", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rows: rowsToSend })
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(errText || "Lỗi tạo file Excel");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "TONG_HOP_HANGER_XUONG_GON.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      alert("Không thể xuất file Excel gọn: " + err.message);
    } finally {
      btnDownloadCleanExcel.disabled = false;
      btnDownloadCleanExcel.innerHTML = originalHtml;
    }
  });

  btnExportJSON.addEventListener("click", () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(displayedRows || [], null, 2));
    const dlAnchor = document.createElement("a");
    dlAnchor.setAttribute("href", dataStr);
    dlAnchor.setAttribute("download", "hanger_audit_simulation.json");
    dlAnchor.click();
  });

  btnExportCSV.addEventListener("click", () => {
    if (!displayedRows || displayedRows.length === 0) return;
    const headers = [
      "PO#", "Account", "Ref#", "Style", "Description", "Category", "Size", "Hang/Flat",
      "Hanger Code", "Hanger Color", "Color Sizer", "Sticker Hanger", "SIZE Sticker", "NCC",
      "Status", "SOF Sheet", "Source Cells"
    ];

    const csvRows = [headers.join(",")];
    displayedRows.forEach(r => {
      const rowVals = [
        `"${r.po_number || ""}"`,
        `"${r.account || ""}"`,
        `"${r.ref_number || ""}"`,
        `"${r.style || ""}"`,
        `"${(r.product_description || "").replace(/"/g, '""')}"`,
        `"${r.product_category || ""}"`,
        `"${(r.size_configuration || "").replace(/"/g, '""')}"`,
        `"${r.order_hang_flat || ""}"`,
        `"${cleanVal(r.hanger_code)}"`,
        `"${cleanVal(r.hanger_color)}"`,
        `"${cleanVal(r.color_sizer)}"`,
        `"${cleanVal(r.sticker_hanger)}"`,
        `"${cleanVal(r.size_sticker_hanger)}"`,
        `"${cleanVal(r.ncc)}"`,
        `"${r.status || ""}"`,
        `"${r.source_sheet || ""}"`,
        `"${r.source_cells || ""}"`
      ];
      csvRows.push(rowVals.join(","));
    });

    const csvStr = "data:text/csv;charset=utf-8,\uFEFF" + encodeURIComponent(csvRows.join("\n"));
    const dlAnchor = document.createElement("a");
    dlAnchor.setAttribute("href", csvStr);
    dlAnchor.setAttribute("download", "TONG_HOP_HANGER_KETQUA.csv");
    dlAnchor.click();
  });

  // Initialization
  renderRulesTable();
  setupDropzones();
  resetDropzonesToEmpty(); // Mặc định 2 file này sẽ không có file nào!
  renderResultsTable(window.DEMO_ROWS || []);
  testRuleKeyword("BODYSUIT PANT SET");
});
