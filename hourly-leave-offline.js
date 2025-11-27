// hourly-leave-offline.js
// 離線特休試算：每月工時 → 各期特休（6m／1y／2y／3y…）→ 不休假獎金
// 版本：依台灣勞基法比例 + 你提供的 xlsx 切期邏輯
// - 滿 6 個月：前 6 筆，基準 1044，3 日
// - 滿 1 年：前 12 筆，基準 2088，7 日（整年）
// - 滿 2 年：第 13–24 筆，基準 2088，10 日
// - 滿 3 年：第 25–36 筆，基準 2088，14 日
// - 滿 4 年：第 37–48 筆，基準 2088，14 日
// - 滿 5 年：第 49–60 筆，基準 2088，15 日
// - 滿 6 年：第 61–72 筆，基準 2088，15 日

// DOM 元件
const sid       = document.getElementById("student-id");
const unit      = document.getElementById("unit");
const wageInput = document.getElementById("wage");
const hy        = document.getElementById("hire-year");
const hm        = document.getElementById("hire-month");
const hd        = document.getElementById("hire-day");
const tbody     = document.getElementById("tbody");
const periodEl    = document.getElementById("period");
const monthsEl    = document.getElementById("months");
const totalHEl    = document.getElementById("totalH");
const seniorityEl = document.getElementById("seniority");
const resultBox   = document.getElementById("result");

// 每月工時資料：[{ year, month, hours }]
let data = [];

// 排序方向：desc = 新→舊、asc = 舊→新（預設新→舊）
let sortDirection = "desc";

// ------------------------------
// 通用常數（對應 Excel 參數表）
// ------------------------------
const WORK_DAYS_PER_WEEK = 5;   // 一週工時天數
const HOURS_PER_DAY      = 8;   // 每日工時
const WEEKS_PER_YEAR     = 52;  // 一年週數
const EXTRA_DAYS         = 1;   // 365 天多出來的 1 天

// 一年平均正常工時（=2088）
const YEAR_BASE_HOURS = (WORK_DAYS_PER_WEEK * HOURS_PER_DAY) * WEEKS_PER_YEAR
                      + EXTRA_DAYS * HOURS_PER_DAY;

// 半年平均正常工時（=1044）
const HALF_YEAR_BASE_HOURS = YEAR_BASE_HOURS / 2;

// ------------------------------
// 排序 data（依 sortDirection）
// ------------------------------
function sortData() {
  data.sort((a, b) => {
    const ka = (a.year || 0) * 100 + (a.month || 0);
    const kb = (b.year || 0) * 100 + (b.month || 0);
    return sortDirection === "desc" ? kb - ka : ka - kb;
  });
}

const now       = new Date();
const thisYear  = now.getFullYear();
const thisMonth = now.getMonth() + 1;

// ------------------------------
// 下拉式選單初始化（年 / 月 / 日）
// ------------------------------

// 年份：從今年往前 15 年
for (let y = thisYear; y >= thisYear - 15; y--) {
  hy.add(new Option(y, y));
}

// 月份：01 ~ 12
for (let m = 1; m <= 12; m++) {
  hm.add(new Option(String(m).padStart(2, "0"), m));
}

function fillDayOptions(year, month) {
  hd.innerHTML = "";
  const daysInMonth = new Date(year, month, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    hd.add(new Option(d, d));
  }
}

hy.addEventListener("change", () => {
  const y = parseInt(hy.value);
  const m = parseInt(hm.value);
  if (!isNaN(y) && !isNaN(m)) fillDayOptions(y, m);
});

hm.addEventListener("change", () => {
  const y = parseInt(hy.value);
  const m = parseInt(hm.value);
  if (!isNaN(y) && !isNaN(m)) fillDayOptions(y, m);
});

// 預設到職日：今年 1 月 1 日
hy.value = thisYear;
hm.value = 1;
fillDayOptions(thisYear, 1);
hd.value = 1;

// ------------------------------
// 一鍵補到本月（含本月）
// ------------------------------
document.getElementById("fill").onclick = () => {
  const startY = parseInt(hy.value);
  const startM = parseInt(hm.value);
  if (!startY || !startM) {
    alert("請先填寫「到職日」才能一鍵補齊！");
    return;
  }

  const startIndex  = startY * 12 + (startM - 1);
  const targetIndex = thisYear * 12 + (thisMonth - 1);
  if (startIndex > targetIndex) {
    alert("到職年月在未來，無法補到本月。");
    return;
  }

  const diff       = targetIndex - startIndex + 1;
  const MAX_MONTHS = 600;
  const count      = Math.min(diff, MAX_MONTHS);

  for (let i = 0; i < count; i++) {
    const idx = startIndex + i;
    const y   = Math.floor(idx / 12);
    const m   = (idx % 12) + 1;
    if (!data.some(r => r.year === y && r.month === m)) {
      data.push({ year: y, month: m, hours: null });
    }
  }

  render();

  setTimeout(() => {
    const firstInput = tbody.querySelector('input[type="number"]');
    if (firstInput) firstInput.focus();
  }, 0);
};

// ------------------------------
// 新增空白欄位（讀取最新一筆年月 + 1）
// ------------------------------
document.getElementById("add-row").onclick = () => {
  let y, m;

  if (data.length > 0) {
    // 以「年月最大」那一筆為基準 +1 月
    let latest = null;
    for (const r of data) {
      if (!r.year || !r.month) continue;
      if (!latest) {
        latest = { year: r.year, month: r.month };
        continue;
      }
      if (r.year > latest.year || (r.year === latest.year && r.month > latest.month)) {
        latest = { year: r.year, month: r.month };
      }
    }
    if (latest) {
      y = latest.year;
      m = latest.month + 1;
      if (m > 12) {
        m = 1;
        y += 1;
      }
    } else {
      y = thisYear;
      m = thisMonth;
    }
  } else {
    // 沒有資料時，用本月
    y = thisYear;
    m = thisMonth;
  }

  data.push({ year: y, month: m, hours: null });
  render();

  setTimeout(() => {
    const rows = Array.from(tbody.querySelectorAll("tr"));
    let targetInput = null;

    for (const row of rows) {
      const ySel = row.querySelector("td:nth-child(1) select");
      const mSel = row.querySelector("td:nth-child(2) select");
      if (!ySel || !mSel) continue;

      const rowY = parseInt(ySel.value, 10);
      const rowM = parseInt(mSel.value, 10);

      if (rowY === y && rowM === m) {
        const input = row.querySelector('input[type="number"]');
        if (input) {
          targetInput = input;
        }
        break;
      }
    }

    if (targetInput) {
      targetInput.focus();
      targetInput.select();
      targetInput.scrollIntoView({ block: "center", behavior: "smooth" });
    }
  }, 0);
};

// ------------------------------
// 刪除一列
// ------------------------------
function deleteRow(index) {
  if (index < 0 || index >= data.length) return;
  if (!confirm("確定要刪除此列？")) return;
  data.splice(index, 1);
  render();
}

// ------------------------------
// 檢查欄位空值與範圍（0～160）
// ------------------------------
function validateHours() {
  let hasEmpty = false;
  let hasInvalid = false;

  const rows = tbody.querySelectorAll("tr");

  data.forEach((r, idx) => {
    const rowEl = rows[idx];
    if (!rowEl) return;
    const input = rowEl.querySelector('input[type="number"]');
    if (!input) return;

    const v = input.value.trim();

    if (v === "") {
      input.classList.add("error");
      r.hours = null;
      hasEmpty = true;
      return;
    }

    const num = Number(v);
    r.hours = num;

    if (!Number.isFinite(num) || num < 0 || num > 160) {
      input.classList.add("error");
      hasInvalid = true;
      return;
    }

    input.classList.remove("error");
  });

  if (hasEmpty) {
    resultBox.innerHTML = "有欄位未填寫，如要列入年資，請填寫 0～160 之間";
    return { ok: false, code: "EMPTY" };
  }

  if (hasInvalid) {
    resultBox.innerHTML = "有欄位超出 0～160 範圍，請修正後再計算";
    return { ok: false, code: "INVALID" };
  }

  return { ok: true };
}

// ------------------------------
// 畫面渲染
// ------------------------------
function render() {
  sortData();

  tbody.innerHTML = "";
  data.forEach((r, idx) => {
    const tr = document.createElement("tr");

    // 年
    const tdY = document.createElement("td");
    const sY  = document.createElement("select");
    sY.className = "highlight-input";
    for (let y = thisYear; y >= thisYear - 15; y--) {
      sY.add(new Option(y, y));
    }
    if (r.year) sY.value = r.year;
    sY.onchange = () => { r.year = +sY.value; render(); };
    tdY.appendChild(sY);
    tr.appendChild(tdY);

    // 月
    const tdM = document.createElement("td");
    const sM  = document.createElement("select");
    sM.className = "highlight-input";
    for (let m = 1; m <= 12; m++) {
      const label = String(m).padStart(2, "0");
      sM.add(new Option(label, m));
    }
    if (r.month) sM.value = r.month;
    sM.onchange = () => { r.month = +sM.value; render(); };
    tdM.appendChild(sM);
    tr.appendChild(tdM);

    // 工時
    const tdH = document.createElement("td");
    const inp = document.createElement("input");
    inp.type      = "number";
    inp.step      = "0.1";
    inp.min       = "0";
    inp.max       = "160";
    inp.className = "highlight-input";

    if (r.hours !== null && Number.isFinite(r.hours)) {
      inp.value = r.hours;
    }

    inp.oninput = () => {
      const v = inp.value.trim();
      if (v === "") {
        r.hours = null;
        inp.classList.add("error");
        return;
      }
      const num = Number(v);
      r.hours = num;
      if (!Number.isFinite(num) || num < 0 || num > 160) {
        inp.classList.add("error");
      } else {
        inp.classList.remove("error");
      }
    };

    inp.addEventListener("focus", () => inp.select());

    // Enter / Tab 跳到下一列工時欄位
    inp.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || (e.key === "Tab" && !e.shiftKey)) {
        e.preventDefault();
        const allInputs = Array.from(
          tbody.querySelectorAll('input[type="number"]')
        );
        const idx = allInputs.indexOf(inp);
        if (idx >= 0 && idx < allInputs.length - 1) {
          const next = allInputs[idx + 1];
          next.focus();
          next.select();
        }
      }
    });

    tdH.appendChild(inp);
    tr.appendChild(tdH);

    // 操作：刪除
    const tdOp  = document.createElement("td");
    const btnDel = document.createElement("button");
    btnDel.textContent = "刪除";
    btnDel.className = "btn-delete";
    btnDel.onclick = () => deleteRow(idx);
    tdOp.appendChild(btnDel);
    tr.appendChild(tdOp);

    tbody.appendChild(tr);
  });
}

// ------------------------------
// 算式核心：0.5 小時無條件進位
// ------------------------------
function ceilToHalfHour(value) {
  const v = Number(value);
  if (!Number.isFinite(v) || v <= 0) return 0;
  // 無條件進位到 0.5 小時
  return Math.ceil(v * 2 - 1e-9) / 2;
}

// 某一段期間（多個月份）的特休
function calcSegmentHours(slice, baseHours, days) {
  const segHours = slice.reduce((sum, r) => sum + (Number(r.hours) || 0), 0);
  const raw = (segHours / baseHours) * HOURS_PER_DAY * days;
  const rounded = ceilToHalfHour(raw);
  return { segHours, raw, rounded };
}

// ------------------------------
// 將 data → Summary（給畫面用）
// ------------------------------
function computeLeaveSummary() {
  const valid = data
    .filter(r => Number.isFinite(r.hours) && r.hours >= 0 && r.hours <= 160 && r.year && r.month)
    .sort((a, b) => a.year - b.year || a.month - b.month);

  const N = valid.length;
  if (N === 0) {
    return { ok: false, error: "NO_DATA" };
  }

  const totalHours = valid.reduce((s, r) => s + r.hours, 0);

  const first = valid[0];
  const last  = valid[valid.length - 1];

  const periodText =
    `${first.year}-${String(first.month).padStart(2, "0")} ~ ` +
    `${last.year}-${String(last.month).padStart(2, "0")}`;

  const seniorityMonths = N;
  const seniorityYears  = Math.floor(seniorityMonths / 12);
  const remainMonths    = seniorityMonths % 12;
  const seniorityText   = `${seniorityYears} 年 ${remainMonths} 個月`;

  const H_PER_DAY_LOCAL = HOURS_PER_DAY;

  const result = {
    ok: true,
    periodText,
    seniorityText,
    totalHours,
    validMonths: N,
    totalRounded: 0,
    totalRaw: 0,
    segments: []
  };

  // 定義各期（indexFrom = 0 起算）
  // NOTE: indexFrom 是「從第幾筆開始」，length 是「抓幾個月」
  const segmentsDef = [
    { label: "滿 6 個月", indexFrom: 0,  length: 6,  base: HALF_YEAR_BASE_HOURS, days: 3  }, // 第 1～6 筆
    { label: "滿 1 年",   indexFrom: 0,  length: 12, base: YEAR_BASE_HOURS,      days: 7  }, // 第 1～12 筆（整年 7 日）
    { label: "滿 2 年",   indexFrom: 12, length: 12, base: YEAR_BASE_HOURS,      days: 10 }, // 第 13～24 筆
    { label: "滿 3 年",   indexFrom: 24, length: 12, base: YEAR_BASE_HOURS,      days: 14 }, // 第 25～36 筆
    { label: "滿 4 年",   indexFrom: 36, length: 12, base: YEAR_BASE_HOURS,      days: 14 }, // 第 37～48 筆
    { label: "滿 5 年",   indexFrom: 48, length: 12, base: YEAR_BASE_HOURS,      days: 15 }, // 第 49～60 筆
    { label: "滿 6 年",   indexFrom: 60, length: 12, base: YEAR_BASE_HOURS,      days: 15 }, // 第 61～72 筆
  ];

  const segments = [];
  let totalRaw = 0;
  let totalR   = 0;

  for (const def of segmentsDef) {
    const { indexFrom, length, base, days, label } = def;
    const indexTo = indexFrom + length - 1;

    // 如果有效月份數不足該期，就跳過
    if (N <= indexTo) continue;

    const slice = valid.slice(indexFrom, indexFrom + length);
    const { segHours, raw, rounded } = calcSegmentHours(slice, base, days);

    const firstSeg = slice[0];
    const lastSeg  = slice[slice.length - 1];
    const rangeLabel =
      `${firstSeg.year}-${String(firstSeg.month).padStart(2, "0")} ~ ` +
      `${lastSeg.year}-${String(lastSeg.month).padStart(2, "0")}`;

    segments.push({
      label,
      rangeLabel,
      hours: segHours,
      raw,
      rounded
    });

    totalRaw += raw;
    totalR   += rounded;
  }

  result.segments     = segments;
  result.totalRaw     = totalRaw;
  result.totalRounded = totalR;

  return result;
}

// ------------------------------
// 開始計算（畫面輸出）
// ------------------------------
document.getElementById("calc").onclick = () => {
  const ok = validateHours();
  if (!ok.ok) return;

  const summary = computeLeaveSummary();
  if (!summary.ok) {
    alert("請至少輸入一筆工時");
    return;
  }

  periodEl.textContent    = summary.periodText;
  monthsEl.textContent    = summary.validMonths;
  totalHEl.textContent    = summary.totalHours.toFixed(1);
  seniorityEl.textContent = summary.seniorityText;

  resultBox.innerHTML = "";

  if (summary.validMonths < 6) {
    resultBox.textContent = "有效月份未滿 6 個月，尚無特休。";
    return;
  }

  const wage = parseFloat(wageInput.value) || 0;

  const totalLine = document.createElement("div");
  totalLine.style.fontWeight = "600";
  totalLine.textContent =
    `累計總特休：${summary.totalRounded.toFixed(2)} 小時（真實：${summary.totalRaw.toFixed(2)} 小時）`;
  resultBox.appendChild(totalLine);

  if (wage > 0) {
    const bonusRounded = summary.totalRounded * wage;
    const bonusRaw     = summary.totalRaw * wage;
    const bonusLine = document.createElement("div");
    bonusLine.style.color = "#005AB5";
    bonusLine.textContent =
      `累計總不休假獎金：${bonusRounded.toFixed(0)} 元（真實：${bonusRaw.toFixed(0)} 元）`;
    resultBox.appendChild(bonusLine);
  }

  summary.segments.forEach(seg => {
    const wrap = document.createElement("div");
    wrap.style.marginTop = "8px";

    const line1 = document.createElement("div");
    line1.textContent = `${seg.rangeLabel}｜${seg.label}`;
    wrap.appendChild(line1);

    const line2 = document.createElement("div");
    line2.style.color = "#b91c1c";
    line2.textContent =
      `本期特休：${seg.rounded.toFixed(2)} 小時（真實：${seg.raw.toFixed(2)} 小時）`;
    wrap.appendChild(line2);

    const wage = parseFloat(wageInput.value) || 0;
    if (wage > 0) {
      const bonusRounded = seg.rounded * wage;
      const bonusRaw     = seg.raw * wage;
      const line3 = document.createElement("div");
      line3.style.color = "#374151";
      line3.textContent =
        `本期不休假獎金：${bonusRounded.toFixed(0)} 元（真實：${bonusRaw.toFixed(0)} 元）`;
      wrap.appendChild(line3);
    }

    resultBox.appendChild(wrap);
  });
};

// ------------------------------
// 全部清除
// ------------------------------
document.getElementById("clear").onclick = () => {
  if (confirm("確定清除所有資料？")) {
    data = [];
    render();
    resultBox.innerHTML = "尚未計算";
    periodEl.textContent    = "—";
    monthsEl.textContent    = "—";
    totalHEl.textContent    = "—";
    seniorityEl.textContent = "—";
    sid.value   = "";
    unit.value  = "";
    wageInput.value = "";
  }
};

// ------------------------------
// 匯出 Excel
// ------------------------------
document.getElementById("export").onclick = () => {
  const ok = validateHours();
  if (!ok.ok) return;

  if (data.length === 0) {
    alert("沒有資料可匯出");
    return;
  }

  const summary = computeLeaveSummary();
  if (!summary.ok) {
    alert("請至少輸入一筆工時");
    return;
  }

  const studentId = (sid.value || "").trim();
  const unitText  = (unit.value || "").trim();
  const wage      = parseFloat(wageInput.value) || 0;

  const validSorted = data
    .filter(r => Number.isFinite(r.hours) && r.hours >= 0 && r.hours <= 160 && r.year && r.month)
    .sort((a, b) => a.year - b.year || a.month - b.month);

  const sheet = [];

  sheet.push(["累計特休表"]);
  sheet.push([`學號：${studentId || "（未填）"}`]);
  sheet.push([`時薪：${wage > 0 ? wage.toFixed(0) : "—"}`]);
  sheet.push([`服務單位：${unitText || "—"}`]);
  sheet.push([]);
  sheet.push(["試算結果"]);
  sheet.push([`任職區間：${summary.periodText}`]);
  sheet.push([`有效月份數：${summary.validMonths}`]);
  sheet.push([`累計總工時：${summary.totalHours.toFixed(1)} 小時`]);
  sheet.push([`年資：${summary.seniorityText}`]);
  sheet.push([]);

  if (summary.validMonths < 6) {
    sheet.push(["有效月份未滿 6 個月，尚無特休"]);
  } else {
    sheet.push([
      `累計總特休：${summary.totalRounded.toFixed(2)} 小時（真實：${summary.totalRaw.toFixed(2)} 小時）`
    ]);

    if (wage > 0) {
      const bonusRounded = summary.totalRounded * wage;
      const bonusRaw     = summary.totalRaw * wage;
      sheet.push([
        `累計總不休假獎金：${bonusRounded.toFixed(0)} 元（真實：${bonusRaw.toFixed(0)} 元）`
      ]);
    }
  }

  sheet.push([]);
  sheet.push(["每月工時明細"]);
  sheet.push(["年","月","工時（小時）"]);
  validSorted.forEach(r => {
    sheet.push([r.year, r.month, r.hours != null ? r.hours.toFixed(1) : ""]);
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheet), "特休試算");

  let id = studentId || "未填學號";
  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `時薪特休試算_${id}_${today}.xlsx`);
};

// ------------------------------
// 標題列點擊 → 切換排序方向（新→舊 / 舊→新）
// ------------------------------
const headerRow = document.querySelector(".months-table thead tr");
if (headerRow) {
  headerRow.style.cursor = "pointer";
  headerRow.addEventListener("click", () => {
    sortDirection = (sortDirection === "desc") ? "asc" : "desc";
    render();
  });
}

// 初次渲染
render();
