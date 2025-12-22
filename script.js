/* =========================================================
 1) CẤU HÌNH / GOOGLE SHEET / SUBMISSION
========================================================= */
const MAX_CARDS_PER_ORDER = 10;
const MAX_BILLS_PER_CARD = 10;

const GOOGLE_SHEET_WEBAPP_URL =
  "https://script.google.com/macros/s/AKfycbwv81WLcBso7hSH_a-Yfz7WFQX1KNLJp0RN-HyHfGTYJ14qFz2fAiyDBVjK-CzZWaNG/exec";
const GOOGLE_SHEET_SECRET = "THỬ";

// chặn double submit
let IS_SUBMITTING = false;

// submissionId để Apps Script dedupe
function makeSubmissionId() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}_${Math.random().toString(16).slice(2)}`;
}
let CURRENT_SUBMISSION_ID = makeSubmissionId(); // ✅ dùng let để reset xong tạo id mới

/* =========================================================
 2) DATA (STAFF / POS)
========================================================= */
const STAFF_BY_OFFICE = {
  ThaiHa: ["Cường", "Thái", "Thịnh", "Linh", "Trang", "Vượng", "Hoàng anh", "Huy"],
  NguyenXien: ["An", "Kiên", "Trang anh", "Phú", "Trung", "Nam", "Hiệp", "Dương", "Đức anh", "Vinh"],
};
const ALL_STAFF = [...STAFF_BY_OFFICE.ThaiHa, ...STAFF_BY_OFFICE.NguyenXien];

// POS -> HKD -> Máy
const POS_DATA = {
  BV: {
    "THU TRANG 92A": ["1077", "8244", "1076", "8243"],
    "HONG QUAN": ["1732", "9318", "1731", "9317"],
    "XUAN HUNG": ["1864", "9426", "1865", "9427"],
  },
  AB: { "NGOC QUYNH JK M1": ["47"], "THIEN PHONG 83 M1": ["51"] },
  MB: { "MANH THANG - 1": ["T1"], "LƯƠNG TUYẾT LAN 1": ["L1"], "LƯƠNG TUYẾT LAN 2": ["L2"] },
  MBV: { "DUC MANH 1": ["DM1"], "LONG HA 1": ["LH1"] },
  VP: {
    "THÊM NT 70": ["NT"],
    "ANH VN 93": ["AVN"],
    "LINH SANG 10": ["724"],
    "COFFEE 8": ["960"],
    "AN TUONG 10": ["749"],
    "NGO 3": ["707"],
    "LINH SANG 1": ["715"],
  },
};

/* =========================================================
 3) HELPERS (DOM / FORMAT / PARSE)
========================================================= */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function digitsOnly(str) {
  return String(str || "").replace(/[^\d]/g, "");
}

function parseCurrencyVND(str) {
  const d = digitsOnly(str);
  return d ? Math.round(Number(d)) : 0;
}
function formatVND(n) {
  return Number(n || 0).toLocaleString("vi-VN");
}

/**
 * Format currency input theo VND:
 * - allowEmpty=true: xóa sạch => giữ rỗng
 * - allowEmpty=false: xóa sạch => về "0"
 */
function formatCurrencyInput(el, allowEmpty = false) {
  const rawText = String(el.value ?? "");
  const onlyDigits = rawText.replace(/[^\d]/g, "");

  if (allowEmpty && onlyDigits.length === 0) {
    el.value = "";
    return;
  }

  // giữ caret tương đối theo số lượng digit phía trước
  const start = el.selectionStart ?? rawText.length;
  const before = rawText.slice(0, start);
  const digitsBefore = before.replace(/[^\d]/g, "").length;

  const raw = parseCurrencyVND(rawText);
  el.value = raw > 0 ? formatVND(raw) : "0";

  let pos = 0,
    seen = 0;
  while (pos < el.value.length && seen < digitsBefore) {
    if (/\d/.test(el.value[pos])) seen++;
    pos++;
  }
  el.setSelectionRange(pos, pos);
}

/**
 * Parse % theo kiểu VN:
 * - cho phép 1.2 hoặc 1,2 (normalize , -> .)
 * - nếu nhập chữ/ký tự lạ => invalid
 */
function parsePercentVN(raw) {
  const s0 = String(raw || "").trim();
  if (!s0) return { valid: false, value: 0, normalized: "", isEmpty: true };

  const normalized = s0.replace(/\s+/g, "").replace(/,/g, ".");
  const ok = /^\d+(\.\d+)?$/.test(normalized);
  if (!ok) return { valid: false, value: 0, normalized, isEmpty: false };

  const n = Number(normalized);
  if (!Number.isFinite(n)) return { valid: false, value: 0, normalized, isEmpty: false };

  return { valid: true, value: n, normalized, isEmpty: false };
}

/**
 * Chia tiền theo tỉ lệ bill, đảm bảo tổng share = total
 */
function allocateByBill(total, items, getWeight) {
  const list = items.filter((x) => getWeight(x) > 0);
  const sumW = list.reduce((s, x) => s + getWeight(x), 0);

  const map = new Map();
  items.forEach((x) => map.set(x.cardId, 0));
  if (total <= 0 || sumW <= 0) return map;

  let used = 0;
  list.forEach((x, i) => {
    const share =
      i === list.length - 1
        ? Math.max(0, total - used)
        : Math.max(0, Math.round((total * getWeight(x)) / sumW));
    map.set(x.cardId, share);
    used += share;
  });
  return map;
}

/* =========================================================
 4) OFFICE -> STAFF + STAFF FINAL + RETURN FEE AUTO
========================================================= */
function setupOfficeStaffLogic() {
  const form = $("#mainForm");
  const officeEl = $("#office");
  const staffEl = $("#staff");
  const shipEl = $("#staffShip");

  const contactField = $("#customerDetailField");
  const contactInput = $("#customerDetail");

  // hidden staffFinal (đẩy lên Sheet)
  let staffFinalEl = $("#staffFinal");
  if (!staffFinalEl) {
    staffFinalEl = document.createElement("input");
    staffFinalEl.type = "hidden";
    staffFinalEl.id = "staffFinal";
    staffFinalEl.name = "staffFinal";
    form?.appendChild(staffFinalEl);
  }

  // staff ship: all staff + Không
  shipEl.innerHTML = [
    `<option value="Không">Không</option>`,
    ...ALL_STAFF.map((n) => `<option value="${n}">${n}</option>`),
  ].join("");
  if (!shipEl.value) shipEl.value = "Không";

  function showContact() {
    contactField?.classList.remove("hidden");
    if (contactInput) contactInput.required = true;
  }
  function hideContact() {
    contactField?.classList.add("hidden");
    if (contactInput) {
      contactInput.required = false;
      contactInput.value = "";
    }
  }
  function resetStaff() {
    staffEl.disabled = true;
    staffEl.innerHTML = `<option value="">-- Vui lòng chọn Văn phòng trước --</option>`;
    staffEl.value = "";
  }
  function syncStaffFinal() {
    const staffVal = staffEl?.value || "";
    staffFinalEl.value = staffVal === "Khách văn phòng" ? (contactInput?.value || "").trim() : staffVal;
  }

  officeEl?.addEventListener("change", () => {
    const office = officeEl.value;
    hideContact();
    resetStaff();

    if (!office || !STAFF_BY_OFFICE[office]) {
      syncStaffFinal();
      updateReturnFeePercentAuto();
      recalcAll();
      return;
    }

    const list = [...STAFF_BY_OFFICE[office], "Khách văn phòng"];
    staffEl.disabled = false;
    staffEl.innerHTML = [
      `<option value="">-- Chọn Nhân viên --</option>`,
      ...list.map((n) => `<option value="${n}">${n}</option>`),
    ].join("");

    syncStaffFinal();
    updateReturnFeePercentAuto();
    recalcAll();
  });

  staffEl?.addEventListener("change", () => {
    if (staffEl.value === "Khách văn phòng") showContact();
    else hideContact();

    syncStaffFinal();
    updateReturnFeePercentAuto();
    recalcAll();
  });

  contactInput?.addEventListener("input", () => {
    syncStaffFinal();
    updateReturnFeePercentAuto();
  });

  hideContact();
  resetStaff();
  syncStaffFinal();
}

function updateReturnFeePercentAuto() {
  const staffVal = String($("#staff")?.value || "");
  const returnEl = $("#returnFeePercentAll");
  if (!returnEl) return;

  returnEl.readOnly = true;
  returnEl.tabIndex = -1;

  if (!staffVal) {
    returnEl.value = "";
    return;
  }

  // Khách VP: phí chuyển về = % thu khách (chỉ khi % hợp lệ)
  if (staffVal === "Khách văn phòng") {
    const pInfo = parsePercentVN($("#feePercentAll")?.value || "");
    returnEl.value = pInfo.valid ? pInfo.normalized : "";
  } else {
    returnEl.value = "1.45";
  }
}

/* =========================================================
 5) BILL: POS -> HKD -> MÁY + ADD/REMOVE
========================================================= */
function setSelectOptions(selectEl, values, placeholder) {
  if (!selectEl) return;
  selectEl.innerHTML = [`<option value="">${placeholder}</option>`]
    .concat(values.map((v) => `<option value="${String(v)}">${String(v)}</option>`))
    .join("");
}

function initBillRow(rowEl) {
  const posSel = $(".pos-select", rowEl);
  const hkdSel = $(".hkd-select", rowEl);
  const machineSel = $(".machine-select", rowEl);
  if (!posSel || !hkdSel || !machineSel) return;

  posSel.innerHTML = [
    `<option value="">-- Chọn POS --</option>`,
    ...Object.keys(POS_DATA).map((k) => `<option value="${k}">${k}</option>`),
  ].join("");

  setSelectOptions(hkdSel, [], "-- Chọn HKD --");
  setSelectOptions(machineSel, [], "-- Chọn Máy POS --");
  machineSel.classList.remove("auto-locked");
}

function onPosChange(rowEl) {
  const posSel = $(".pos-select", rowEl);
  const hkdSel = $(".hkd-select", rowEl);
  const machineSel = $(".machine-select", rowEl);
  if (!posSel || !hkdSel || !machineSel) return;

  const pos = posSel.value;
  const hkds = pos && POS_DATA[pos] ? Object.keys(POS_DATA[pos]) : [];

  setSelectOptions(hkdSel, hkds, "-- Chọn HKD --");
  setSelectOptions(machineSel, [], "-- Chọn Máy POS --");
  machineSel.classList.remove("auto-locked");
}

function onHKDChange(rowEl) {
  const posSel = $(".pos-select", rowEl);
  const hkdSel = $(".hkd-select", rowEl);
  const machineSel = $(".machine-select", rowEl);
  if (!posSel || !hkdSel || !machineSel) return;

  const pos = posSel.value;
  const hkd = hkdSel.value;
  const machines = pos && hkd && POS_DATA[pos] && POS_DATA[pos][hkd] ? POS_DATA[pos][hkd] : [];

  // BV: cho chọn máy
  if (pos === "BV") {
    setSelectOptions(machineSel, machines, "-- Chọn Máy POS --");
    machineSel.classList.remove("auto-locked");
    machineSel.value = "";
    return;
  }

  // POS khác: auto lock theo máy đầu tiên
  if (machines.length > 0) {
    setSelectOptions(machineSel, machines, "-- Máy POS --");
    machineSel.value = String(machines[0]);
    machineSel.classList.add("auto-locked");
  } else {
    setSelectOptions(machineSel, [], "-- Chọn Máy POS --");
    machineSel.classList.remove("auto-locked");
  }
}

function getBillCount(cardId) {
  const wrap = $(`#billDetails_${cardId}`);
  return wrap ? $$(".bill-row", wrap).length : 0;
}

function billRowMarkup(cardId, billIndex) {
  return `
    <div class="bill-row" data-bill-index="${billIndex}">
      <input type="text" class="bill-label" name="billLabel_${cardId}_${billIndex}" value="Bill ${billIndex}" readonly />

      <input type="text" class="bill-amount card-input currency-input"
             name="billAmount_${cardId}_${billIndex}"
             value="0" inputmode="numeric" placeholder="Số tiền" required />

      <select class="bill-pos card-input pos-select" name="billPOS_${cardId}_${billIndex}" required>
        <option value="">-- Chọn POS --</option>
      </select>

      <select class="bill-hkd card-input hkd-select" name="billHKD_${cardId}_${billIndex}" required>
        <option value="">-- Chọn HKD --</option>
      </select>

      <select class="bill-machine card-input machine-select" name="billMachine_${cardId}_${billIndex}" required>
        <option value="">-- Chọn Máy POS --</option>
      </select>

      <input type="text" class="bill-batch card-input integer-only" name="billBatch_${cardId}_${billIndex}" placeholder="Số lô" />
      <input type="text" class="bill-invoice card-input integer-only" name="billInvoice_${cardId}_${billIndex}" placeholder="Số hóa đơn" />

      <button type="button" class="remove-bill-btn" title="Xóa bill">Xóa BILL</button>
    </div>
  `;
}

function addBillRow(cardId) {
  const wrapper = $(`#billDetails_${cardId}`);
  if (!wrapper) return;

  const current = getBillCount(cardId);
  if (current >= MAX_BILLS_PER_CARD) {
    alert(`Mỗi thẻ tối đa ${MAX_BILLS_PER_CARD} bill.`);
    return;
  }

  const nextIndex = current + 1;
  wrapper.insertAdjacentHTML("beforeend", billRowMarkup(cardId, nextIndex));

  const row = wrapper.querySelector(`.bill-row[data-bill-index="${nextIndex}"]`);
  if (row) {
    initBillRow(row);
    $$(".currency-input", row).forEach((el) => formatCurrencyInput(el, true));
  }

  recalcAll();
}

function renumberBills(cardId) {
  const wrapper = $(`#billDetails_${cardId}`);
  if (!wrapper) return;

  $$(".bill-row", wrapper).forEach((row, i) => {
    const idx = i + 1;
    row.dataset.billIndex = String(idx);

    const label = $(".bill-label", row);
    const amount = $(".bill-amount", row);
    const pos = $(".pos-select", row);
    const hkd = $(".hkd-select", row);
    const machine = $(".machine-select", row);
    const batch = $(".bill-batch", row);
    const invoice = $(".bill-invoice", row);

    if (label) {
      label.value = `Bill ${idx}`;
      label.name = `billLabel_${cardId}_${idx}`;
    }
    if (amount) amount.name = `billAmount_${cardId}_${idx}`;
    if (pos) pos.name = `billPOS_${cardId}_${idx}`;
    if (hkd) hkd.name = `billHKD_${cardId}_${idx}`;
    if (machine) machine.name = `billMachine_${cardId}_${idx}`;
    if (batch) batch.name = `billBatch_${cardId}_${idx}`;
    if (invoice) invoice.name = `billInvoice_${cardId}_${idx}`;
  });
}

/* =========================================================
 6) CARD: ADD/REMOVE + SERVICE TOGGLE
========================================================= */
function getCardEls() {
  return $$(".card-item");
}
function getAllCardIds() {
  return getCardEls()
    .map((s) => String(s.dataset.cardId || ""))
    .filter(Boolean);
}
function getCardCount() {
  return getCardEls().length;
}

function updateCardLimitUI() {
  const btn = $("#addCardBtn");
  if (!btn) return;
  btn.disabled = getCardCount() >= MAX_CARDS_PER_ORDER;
  btn.title = btn.disabled ? `Tối đa ${MAX_CARDS_PER_ORDER} thẻ/đơn` : "";
}

function normalizeService(raw) {
  const s = String(raw || "").trim().toUpperCase();
  if (!s) return "";
  if (s.includes("DAO") && s.includes("RUT")) return "DAO_RUT";
  if (s.includes("RUT")) return "RUT";
  if (s.includes("DAO")) return "DAO";
  return s;
}

/* =========================================================
 6.1) “PHÍ ÂM / BACK KHÁCH”: CHÈN/XÓA THEO DỊCH VỤ
   - RÚT: ẩn
   - ĐÁO / ĐÁO+RÚT / MIX: hiện
========================================================= */
function ensureDifferenceRow(cardId) {
  const exist = document.getElementById(`differenceAmount_${cardId}`);
  if (exist) return exist;

  const cardEl = document.querySelector(`.card-item[data-card-id="${cardId}"]`);
  const metrics = cardEl?.querySelector(".plain-metrics");
  if (!cardEl || !metrics) return null;

  const row = document.createElement("div");
  row.className = "metric-row";
  row.innerHTML = `
    <span class="metric-label">Phí âm / Back khách:</span>
    <span class="metric-value"><span id="differenceAmount_${cardId}">0</span> VNĐ</span>
    <span class="metric-note">(tiền âm thẻ âm, tiền dương thẻ dư/ rút)</span>
  `;

  // chèn dưới “Tổng tiền làm thẻ”
  const firstRow = metrics.querySelector(".metric-row");
  if (firstRow && firstRow.nextSibling) metrics.insertBefore(row, firstRow.nextSibling);
  else metrics.appendChild(row);

  return document.getElementById(`differenceAmount_${cardId}`);
}

function removeDifferenceRow(cardId) {
  const valueEl = document.getElementById(`differenceAmount_${cardId}`);
  const row = valueEl?.closest(".metric-row");
  if (row) row.remove();
}

/**
 * Toggle fieldset:
 * - chưa chọn => hidden + disabled
 * - đã chọn => show + enable
 *
 * ✅ đổi RÚT -> ĐÁO/ĐÁO+RÚT: chèn “Phí âm” NGAY lập tức
 */
function toggleServiceDetails(cardId, opts = { recalc: true }) {
  const service = normalizeService($(`#serviceType_${cardId}`)?.value || "");
  const fs = $(`#serviceDetails_${cardId}`);
  if (!fs) return;

  const show = !!service;
  fs.classList.toggle("hidden", !show);
  fs.disabled = !show;

  // instant insert/remove
  if (service === "RUT") removeDifferenceRow(cardId);
  if (service === "DAO" || service === "DAO_RUT") ensureDifferenceRow(cardId);

  if (opts.recalc) recalcAll();
}

/* ===== reindex helpers (khi xóa thẻ) ===== */
function replaceCardToken(str, oldId, newId) {
  if (!str) return str;
  str = str.replaceAll(`_${oldId}_`, `_${newId}_`);
  str = str.replace(new RegExp(`_${oldId}$`), `_${newId}`);
  return str;
}

function updateCardIdentifiers(cardEl, oldId, newId) {
  cardEl.dataset.cardId = String(newId);

  const h2 = cardEl.querySelector("h2");
  if (h2) h2.textContent = `2. Thông tin Thẻ #${newId}`;

  $$("[data-card-id]", cardEl).forEach((el) => (el.dataset.cardId = String(newId)));
  $$("[id]", cardEl).forEach((el) => (el.id = replaceCardToken(el.id, oldId, newId)));
  $$("[name]", cardEl).forEach((el) => (el.name = replaceCardToken(el.name, oldId, newId)));
  $$("label[for]", cardEl).forEach((lb) =>
    lb.setAttribute("for", replaceCardToken(lb.getAttribute("for"), oldId, newId))
  );
}

function resetCardValues(cardEl) {
  const cardId = String(cardEl.dataset.cardId);

  // reset input
  $$("input", cardEl).forEach((inp) => {
    if (inp.classList.contains("bill-label")) return;
    if (inp.classList.contains("currency-input")) inp.value = "0";
    else inp.value = "";
  });

  // reset service
  const serviceSel = $(".service-selector", cardEl);
  if (serviceSel) serviceSel.value = "";

  // hide fieldset
  const fs = $(`#serviceDetails_${cardId}`, cardEl) || $(".service-details-container", cardEl);
  if (fs) {
    fs.classList.add("hidden");
    fs.disabled = true;
  }

  // bỏ row phí âm nếu có
  removeDifferenceRow(cardId);

  // reset bills -> 1 dòng
  const wrapper = $(`#billDetails_${cardId}`, cardEl);
  if (wrapper) {
    wrapper.innerHTML = billRowMarkup(cardId, 1);
    $$(".bill-row", wrapper).forEach(initBillRow);
  }

  // reset UI total bill
  const t = $(`#totalBillAmount_${cardId}`, cardEl);
  if (t) t.textContent = "0";

  // format currency inputs (transfer/bill allow empty)
  $$(".currency-input", cardEl).forEach((el) => {
    const allowEmpty =
      el.id === "actualFeeReceived" ||
      el.id === "shipFee" ||
      el.id === "feeFixedAll" ||
      String(el.id || "").startsWith("transferAmount_") ||
      el.classList.contains("bill-amount");
    formatCurrencyInput(el, allowEmpty);
  });
}

function addNewCard() {
  const container = $("#cardContainer");
  const template = $(`.card-item[data-card-id="1"]`);
  if (!container || !template) return;

  if (getCardCount() >= MAX_CARDS_PER_ORDER) {
    alert(`Mỗi đơn tối đa ${MAX_CARDS_PER_ORDER} thẻ.`);
    updateCardLimitUI();
    return;
  }

  const newId = String(getCardCount() + 1);
  const clone = template.cloneNode(true);

  updateCardIdentifiers(clone, "1", newId);
  resetCardValues(clone);

  container.appendChild(clone);
  updateCardLimitUI();
  recalcAll();
}

function reindexCards() {
  const cards = getCardEls();

  // pass 1: đổi qua TMP để tránh trùng ID
  cards.forEach((cardEl, i) => {
    const oldId = String(cardEl.dataset.cardId);
    updateCardIdentifiers(cardEl, oldId, `TMP${i + 1}`);
  });

  // pass 2: TMP -> 1..n
  getCardEls().forEach((cardEl, i) => {
    const tmpId = String(cardEl.dataset.cardId);
    updateCardIdentifiers(cardEl, tmpId, String(i + 1));
  });

  // init lại bill/select + toggle UI
  getAllCardIds().forEach((cardId) => {
    $$(`#billDetails_${cardId} .bill-row`).forEach(initBillRow);
    renumberBills(cardId);
    toggleServiceDetails(cardId, { recalc: false });
  });

  updateCardLimitUI();
  recalcAll();
}

/* =========================================================
 7) PHÍ (% -> CỨNG): UI + BASE FEE
========================================================= */
function showFeeFixedGroup(show) {
  const g = $("#feeFixedGroup");
  if (!g) return;
  g.classList.toggle("hidden", !show);
}

function getFeeFixedInput() {
  return parseCurrencyVND($("#feeFixedAll")?.value || "");
}

/**
 * ✅ Rule mới theo yêu cầu:
 * - % hợp lệ: ẨN phí cứng (không auto)
 * - % không hợp lệ (nhập chữ/ký tự): HIỆN phí cứng + bắt buộc nhập tay
 * - % rỗng: ẨN phí cứng
 */
function syncFeeFixedFromPercent(totalBillAll) {
  // totalBillAll giữ signature (không dùng vì không auto)
  const percentInfo = parsePercentVN($("#feePercentAll")?.value || "");
  const fixedEl = $("#feeFixedAll");
  if (!fixedEl) return percentInfo;

  const shouldShowFixed = !percentInfo.isEmpty && !percentInfo.valid;
  showFeeFixedGroup(shouldShowFixed);

  fixedEl.required = shouldShowFixed;

  // ✅ tuyệt đối KHÔNG auto tính phí cứng theo %
  return percentInfo;
}

/**
 * ✅ Base fee tổng:
 * - % hợp lệ => tính theo %
 * - % không hợp lệ (có nhập) => dùng phí cứng user nhập
 * - % rỗng => 0
 */
function getBaseFeeTotal(totalBillAll, percentInfo) {
  if (percentInfo.valid) {
    return Math.round((totalBillAll * percentInfo.value) / 100);
  }
  if (!percentInfo.isEmpty) {
    return getFeeFixedInput();
  }
  return 0;
}

/* =========================================================
 8) TÍNH TOÁN + UPDATE UI
========================================================= */
function getShipFee() {
  return parseCurrencyVND($("#shipFee")?.value || "");
}
function getActualFeeReceived() {
  return parseCurrencyVND($("#actualFeeReceived")?.value || "");
}

function getCardBillTotal(cardId) {
  const wrap = $(`#billDetails_${cardId}`);
  if (!wrap) return 0;
  let sum = 0;
  $$(".bill-amount", wrap).forEach((inp) => (sum += parseCurrencyVND(inp.value)));
  return sum;
}

function lockTransferInput(transferEl, lock) {
  if (!transferEl) return;
  if (lock) {
    transferEl.readOnly = true;
    transferEl.tabIndex = -1;
    transferEl.classList.add("auto-locked");
  } else {
    transferEl.readOnly = false;
    transferEl.tabIndex = 0;
    transferEl.classList.remove("auto-locked");
  }
}

function updateCardMetricsUI(cardId, totalBill, transfer, service) {
  $(`#totalBillAmount_${cardId}`)?.replaceChildren(document.createTextNode(formatVND(totalBill)));

  // RÚT: bỏ hẳn “Phí âm”
  if (service === "RUT") {
    removeDifferenceRow(cardId);
    return;
  }

  const diffSpan = ensureDifferenceRow(cardId);
  if (diffSpan) diffSpan.replaceChildren(document.createTextNode(formatVND(totalBill - transfer)));
}

/**
 * recalcAll:
 * - Tính tổng bill
 * - Tính baseFeeTotal theo (% hợp lệ) hoặc (phí cứng nếu % invalid)
 * - Nếu có thẻ RÚT: auto tiền chuyển = bill - feeShare + actualShare
 * - Tính totalTransferAll
 * - MIX dịch vụ: tính giống ĐÁO+RÚT (theo yêu cầu)
 *
 * ⚠️ NOTE theo logic hiện tại:
 * - RÚT_ONLY: totalCollect = baseFeeTotal + shipFee (chưa cộng ship trước đây -> đã cộng)
 * - totalPay = totalTransferAll
 */
function recalcAll() {
  const cardIds = getAllCardIds();

  // 1) meta
  let totalBillAll = 0;
  const meta = cardIds.map((cardId) => {
    const totalBill = getCardBillTotal(cardId);
    totalBillAll += totalBill;
    const service = normalizeService($(`#serviceType_${cardId}`)?.value || "");
    return { cardId, totalBill, service };
  });

  // 2) sync show/hide phí cứng theo rule mới
  const percentInfo = syncFeeFixedFromPercent(totalBillAll);
  const baseFeeTotal = getBaseFeeTotal(totalBillAll, percentInfo);

  // 3) chia baseFee cho thẻ có service (để RÚT auto tiền chuyển)
  const cardsWithService = meta.filter((c) => !!c.service);
  const feeShareMap = allocateByBill(baseFeeTotal, cardsWithService, (c) => c.totalBill);

  // 4) chia actualFee cho thẻ RÚT
  const actualFeeReceived = getActualFeeReceived();
  const withdrawCards = meta.filter((c) => c.service === "RUT");
  const actualShareMap = allocateByBill(actualFeeReceived, withdrawCards, (c) => c.totalBill);

  // 5) tính transfer per-card
  let totalTransferAll = 0;
  const serviceSet = new Set();

  meta.forEach((m) => {
    const { cardId, totalBill, service } = m;
    if (service) serviceSet.add(service);

    const transferEl = $(`#transferAmount_${cardId}`);
    const feeShare = feeShareMap.get(cardId) || 0;
    const actualShare = actualShareMap.get(cardId) || 0;

    // RÚT: auto Transfer = Bill - FeeShare + ActualShare
    if (service === "RUT" && transferEl) {
      const autoTransfer = Math.max(0, totalBill - feeShare + actualShare);
      lockTransferInput(transferEl, true);
      transferEl.value = autoTransfer > 0 ? formatVND(autoTransfer) : "0";
    } else {
      lockTransferInput(transferEl, false);
    }

    const transfer = parseCurrencyVND(transferEl?.value || "");
    totalTransferAll += transfer;

    updateCardMetricsUI(cardId, totalBill, transfer, service);
  });

  // 6) mode
  const onlyService = serviceSet.size === 1 ? [...serviceSet][0] : "";
  let mode = "";
  if (serviceSet.size === 1 && onlyService === "RUT") mode = "RUT_ONLY";
  else if (serviceSet.size === 1 && onlyService === "DAO") mode = "DAO_ONLY";
  else if (serviceSet.size === 1 && onlyService === "DAO_RUT") mode = "DAO_RUT_ONLY";
  else if (serviceSet.size >= 2) mode = "MIX";

  // 7) thẻ âm (RÚT_ONLY => 0)
  const totalDiff = totalBillAll - totalTransferAll;
  const negativeCardValue = mode === "RUT_ONLY" ? 0 : totalDiff < 0 ? Math.abs(totalDiff) : 0;

  const negEl = $("#negativeCardFee");
  if (negEl) {
    negEl.readOnly = true;
    negEl.tabIndex = -1;
    negEl.value = formatVND(negativeCardValue);
  }

  // 8) tổng thu/trả
  const shipFee = getShipFee();
  let totalCollect = 0;
  let totalPay = 0;
  let forcePaidStatus = false;

  function applyResult(result) {
    if (result >= 0) {
      totalCollect = result;
      totalPay = 0;
    } else {
      totalCollect = 0;
      totalPay = Math.abs(result);
      forcePaidStatus = true; // âm => đã thu
    }
  }

  if (!mode) {
    totalCollect = 0;
    totalPay = 0;
  } else if (mode === "DAO_ONLY") {
    // ĐÁO: phí = baseFee + ship + thẻ âm (có thể âm => chuyển sang trả khách)
    applyResult(baseFeeTotal + shipFee + negativeCardValue);
  } else if (mode === "DAO_RUT_ONLY") {
    // ĐÁO+RÚT: phí = baseFee + ship + thẻ âm + thực thu (có thể âm)
    applyResult(baseFeeTotal + shipFee + negativeCardValue + actualFeeReceived);
  } else if (mode === "RUT_ONLY") {
    // RÚT: trả = tổng tiền chuyển; thu = baseFee + ship (✅ đã cộng ship)
    totalPay = totalTransferAll;
    totalCollect = baseFeeTotal + shipFee;
  } else {
    // MIX: tính giống ĐÁO+RÚT (theo yêu cầu)
    applyResult(baseFeeTotal + shipFee + negativeCardValue + actualFeeReceived);
  }

  // 9) update UI tổng
  $("#totalBillAll").textContent = `${formatVND(totalBillAll)} VNĐ`;
  $("#totalFeeCollectedAll").textContent = `${formatVND(totalCollect)} VNĐ`;
  $("#totalCustomerPayment").textContent = `${formatVND(totalPay)} VNĐ`;

  // fee status: âm => đã thu luôn; còn lại theo actualFeeReceived
  const payStatus = $("#feePaymentStatus");
  if (payStatus) {
    if (forcePaidStatus) payStatus.value = "da_thu";
    else payStatus.value = actualFeeReceived > 0 ? "da_thu" : "chua_thu";
  }

  updateReturnFeePercentAuto();
}

/* =========================================================
 9) GOOGLE SHEET: PAYLOAD + SUBMIT
========================================================= */
function cardTypeToText(v) {
  switch (String(v || "")) {
    case "V":
      return "Visa";
    case "M":
      return "MasterCard";
    case "J":
      return "JCB";
    case "n":
      return "Napas";
    default:
      return String(v || "");
  }
}
function serviceToText(v) {
  switch (String(v || "")) {
    case "DAO":
      return "ĐÁO";
    case "RUT":
      return "RÚT";
    case "DAO_RUT":
      return "ĐÁO+RÚT";
    default:
      return String(v || "");
  }
}
function feeStatusToText(v) {
  return String(v || "") === "da_thu" ? "Đã thu" : "Chưa thu";
}

function collectPayloadForSheet() {
  const order = {
    office: $("#office")?.value || "",
    date: $("#date")?.value || "",
    staffFinal: $("#staffFinal")?.value || "",
    staffShip: $("#staffShip")?.value || "",
  };

  const percentInfo = parsePercentVN($("#feePercentAll")?.value || "");
  const summary = {
    feePercentAll: percentInfo.valid ? percentInfo.value : 0, // nếu % invalid => 0
    feePercentRaw: $("#feePercentAll")?.value || "", // lưu raw để trace
    feeFixedAll: getFeeFixedInput(),

    returnFeePercentAll: Number($("#returnFeePercentAll")?.value || 0),
    shipFee: parseCurrencyVND($("#shipFee")?.value || 0),
    negativeCardFee: parseCurrencyVND($("#negativeCardFee")?.value || 0),

    totalBillAll: parseCurrencyVND($("#totalBillAll")?.textContent || 0),
    totalFeeCollectedAll: parseCurrencyVND($("#totalFeeCollectedAll")?.textContent || 0),
    totalCustomerPayment: parseCurrencyVND($("#totalCustomerPayment")?.textContent || 0),
    actualFeeReceived: parseCurrencyVND($("#actualFeeReceived")?.value || 0),

    feePaymentStatus: $("#feePaymentStatus")?.value || "chua_thu",
    feePaymentStatusText: feeStatusToText($("#feePaymentStatus")?.value || "chua_thu"),
  };

  const cards = getCardEls().map((cardEl) => {
    const cardId = String(cardEl.dataset.cardId);

    const bills = $$(".bill-row", cardEl).map((row, idx) => ({
      isFirstOfCard: idx === 0,
      amount: parseCurrencyVND($(".bill-amount", row)?.value || 0),
      pos: $(".pos-select", row)?.value || "",
      hkd: $(".hkd-select", row)?.value || "",
      machine: $(".machine-select", row)?.value || "",
      batch: $(".bill-batch", row)?.value || "",
      invoice: $(".bill-invoice", row)?.value || "",
    }));

    const cardType = $(`#cardType_${cardId}`)?.value || "";
    const serviceType = normalizeService($(`#serviceType_${cardId}`)?.value || "");

    return {
      cardName: $(`#cardName_${cardId}`)?.value || "",
      cardNumber: $(`#cardNumber_${cardId}`)?.value || "",
      cardType: cardTypeToText(cardType),
      cardBank: $(`#cardBank_${cardId}`)?.value || "",
      serviceType: serviceToText(serviceType),

      transferAmount: parseCurrencyVND($(`#transferAmount_${cardId}`)?.value || 0),
      totalBill: parseCurrencyVND($(`#totalBillAmount_${cardId}`)?.textContent || 0),

      bills,
    };
  });

  return {
    secret: GOOGLE_SHEET_SECRET,
    submissionId: CURRENT_SUBMISSION_ID,
    order,
    summary,
    cards,
  };
}

async function sendToGoogleSheet(payload) {
  await fetch(GOOGLE_SHEET_WEBAPP_URL, {
    method: "POST",
    mode: "no-cors",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

/* =========================================================
 10) RESET FORM (KHÔNG RELOAD)
========================================================= */
function setTodayForDateInput() {
  const dateInput = $("#date");
  if (!dateInput) return;

  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  dateInput.value = `${yyyy}-${mm}-${dd}`;
}

/**
 * Reset “sạch”:
 * - Xóa các thẻ thêm, chỉ giữ lại thẻ #1
 * - Reset thẻ #1 về trạng thái ban đầu
 * - Reset mục 1 + mục 3
 * - Recalc lại
 * - Tạo submissionId mới cho lần gửi tiếp theo
 */
function resetFormAfterSubmit() {
  const form = $("#mainForm");
  if (!form) return;

  // 1) reset native inputs về default
  form.reset();

  // 2) giữ lại 1 thẻ
  const cards = getCardEls();
  cards.slice(1).forEach((el) => el.remove());

  // 3) reset thẻ #1
  const firstCard = getCardEls()[0];
  if (firstCard) {
    const oldId = String(firstCard.dataset.cardId || "1");
    if (oldId !== "1") updateCardIdentifiers(firstCard, oldId, "1");
    resetCardValues(firstCard);
    toggleServiceDetails("1", { recalc: false });
  }

  // 4) reset office/staff về trạng thái ban đầu
  const officeEl = $("#office");
  if (officeEl) {
    officeEl.value = "";
    officeEl.dispatchEvent(new Event("change", { bubbles: true }));
  }

  // 5) set lại ngày hôm nay
  setTodayForDateInput();

  // 6) reset fee inputs
  const feePercent = $("#feePercentAll");
  if (feePercent) feePercent.value = "";

  showFeeFixedGroup(false);
  const fixedEl = $("#feeFixedAll");
  if (fixedEl) {
    fixedEl.required = false;
    fixedEl.value = "";
  }

  const shipEl = $("#shipFee");
  if (shipEl) {
    shipEl.value = "";
    formatCurrencyInput(shipEl, true);
  }

  const actualEl = $("#actualFeeReceived");
  if (actualEl) {
    actualEl.value = "";
    formatCurrencyInput(actualEl, true);
  }

  // negative luôn readonly (recalcAll sẽ set)
  const neg = $("#negativeCardFee");
  if (neg) {
    neg.readOnly = true;
    neg.tabIndex = -1;
  }

  // status về chưa thu
  const payStatus = $("#feePaymentStatus");
  if (payStatus) payStatus.value = "chua_thu";

  // 7) init lại bill rows selects
  $$(".bill-row").forEach(initBillRow);

  // 8) update UI + recalc
  updateCardLimitUI();
  recalcAll();

  // 9) submission id mới
  CURRENT_SUBMISSION_ID = makeSubmissionId();

  // 10) mở khóa submit cho lần tiếp theo
  IS_SUBMITTING = false;
}

/* =========================================================
 11) INIT + EVENTS
========================================================= */
document.addEventListener("DOMContentLoaded", () => {
  // set date today
  setTodayForDateInput();

  setupOfficeStaffLogic();
  updateReturnFeePercentAuto();

  // init existing bill rows
  $$(".bill-row").forEach(initBillRow);

  // format currency inputs
  $$(".currency-input").forEach((el) => {
    const allowEmpty =
      el.id === "actualFeeReceived" ||
      el.id === "shipFee" ||
      el.id === "feeFixedAll" ||
      String(el.id || "").startsWith("transferAmount_") ||
      el.classList.contains("bill-amount");
    formatCurrencyInput(el, allowEmpty);
  });

  // thẻ âm readonly
  const neg = $("#negativeCardFee");
  if (neg) {
    neg.readOnly = true;
    neg.tabIndex = -1;
  }

  // toggle card #1
  toggleServiceDetails("1", { recalc: false });

  updateCardLimitUI();
  recalcAll();
});

document.addEventListener("click", (e) => {
  const t = e.target;

  // add card
  if (t?.id === "addCardBtn") {
    addNewCard();
    return;
  }

  // remove card (giữ tối thiểu 1)
  if (t?.classList?.contains("remove-card-btn")) {
    if (getCardCount() <= 1) {
      alert("Mỗi đơn phải có ít nhất 1 thẻ. Không thể xóa thêm.");
      return;
    }
    const cardId = t.dataset.cardId;
    $(`.card-item[data-card-id="${cardId}"]`)?.remove();
    reindexCards();
    return;
  }

  // add bill
  if (t?.classList?.contains("add-bill-btn")) {
    addBillRow(t.dataset.cardId);
    return;
  }

  // remove bill (giữ tối thiểu 1)
  if (t?.classList?.contains("remove-bill-btn")) {
    const row = t.closest(".bill-row");
    const card = t.closest(".card-item");
    const cardId = card?.dataset.cardId;
    if (!row || !cardId) return;

    if (getBillCount(cardId) <= 1) {
      alert("Mỗi thẻ cần tối thiểu 1 bill.");
      return;
    }

    row.remove();
    renumberBills(cardId);
    recalcAll();
  }
});

document.addEventListener("change", (e) => {
  const t = e.target;

  // service selector
  if (t?.classList?.contains("service-selector")) {
    const cardId = t.closest(".card-item")?.dataset.cardId;
    if (cardId) toggleServiceDetails(cardId);
    return;
  }

  // POS/HKD select in bill row
  const row = t?.closest?.(".bill-row");
  if (row) {
    if (t.classList.contains("pos-select")) onPosChange(row);
    if (t.classList.contains("hkd-select")) onHKDChange(row);
  }
});

document.addEventListener("input", (e) => {
  const t = e.target;

  // digits-only: 4 số thẻ
  if (String(t?.id || "").startsWith("cardNumber_")) {
    t.value = digitsOnly(t.value).slice(0, 4);
  }
  // digits-only: lô/hóa đơn
  if (t?.classList?.contains("integer-only")) {
    t.value = digitsOnly(t.value);
  }

  // currency format
  if (t?.classList?.contains("currency-input")) {
    const allowEmpty =
      t.id === "actualFeeReceived" ||
      t.id === "shipFee" ||
      t.id === "feeFixedAll" ||
      String(t.id || "").startsWith("transferAmount_") ||
      t.classList.contains("bill-amount");
    formatCurrencyInput(t, allowEmpty);
  }

  // recalc triggers
  if (t?.id === "feePercentAll") {
    recalcAll();
    return;
  }
  if (t?.id === "feeFixedAll") {
    recalcAll();
    return;
  }
  if (t?.id === "shipFee" || t?.id === "actualFeeReceived") {
    recalcAll();
    return;
  }

  if (t?.classList?.contains("bill-amount") || String(t?.id || "").startsWith("transferAmount_")) {
    recalcAll();
  }
});

// Submit: validate + gửi Google Sheet + reset form
$("#mainForm")?.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (IS_SUBMITTING) return;
  IS_SUBMITTING = true;

  // sync staffFinal
  const staffFinal = $("#staffFinal");
  const staffEl = $("#staff");
  const contactEl = $("#customerDetail");

  if (staffFinal) {
    staffFinal.value =
      staffEl?.value === "Khách văn phòng" ? (contactEl?.value || "").trim() : staffEl?.value || "";
  }

  if (staffEl?.value === "Khách văn phòng" && !(contactEl?.value || "").trim()) {
    alert("Vui lòng nhập Liên hệ cho Khách văn phòng.");
    contactEl?.focus();
    IS_SUBMITTING = false;
    return;
  }

  // ✅ Nếu % không hợp lệ => bắt buộc nhập phí cứng
  const percentInfo = parsePercentVN($("#feePercentAll")?.value || "");
  const fixed = getFeeFixedInput();
  if (!percentInfo.isEmpty && !percentInfo.valid && fixed <= 0) {
    alert("Phí thu khách (%) không hợp lệ. Vui lòng nhập Phí cứng (VNĐ).");
    $("#feeFixedAll")?.focus();
    IS_SUBMITTING = false;
    return;
  }

  // validate service + numbers
  for (const cardId of getAllCardIds()) {
    const service = normalizeService($(`#serviceType_${cardId}`)?.value || "");
    if (!service) {
      alert("Vui lòng chọn Dịch vụ cho tất cả các thẻ.");
      $(`#serviceType_${cardId}`)?.focus();
      IS_SUBMITTING = false;
      return;
    }

    const v = String($(`#cardNumber_${cardId}`)?.value || "").trim();
    if (!/^\d{4}$/.test(v)) {
      alert("4 số đuôi thẻ phải là số nguyên đúng 4 chữ số.");
      $(`#cardNumber_${cardId}`)?.focus();
      IS_SUBMITTING = false;
      return;
    }

    const rows = $$(`#billDetails_${cardId} .bill-row`);
    for (const row of rows) {
      const batch = String($(".bill-batch", row)?.value || "").trim();
      const inv = String($(".bill-invoice", row)?.value || "").trim();

      if (batch && !/^\d+$/.test(batch)) {
        alert("Số lô phải là số nguyên.");
        $(".bill-batch", row)?.focus();
        IS_SUBMITTING = false;
        return;
      }
      if (inv && !/^\d+$/.test(inv)) {
        alert("Số hóa đơn phải là số nguyên.");
        $(".bill-invoice", row)?.focus();
        IS_SUBMITTING = false;
        return;
      }
    }
  }

  // đảm bảo tổng mới nhất
  recalcAll();

  const payload = collectPayloadForSheet();

  try {
    await sendToGoogleSheet(payload);
    alert("Đã gửi đơn thành công. Em Hằng xin cảm ơn anh chị ạ!");
    resetFormAfterSubmit(); // ✅ reset form thay vì reload
  } catch (err) {
    console.error(err);
    alert("Gửi đơn thất bại. Liên hệ em Hằng báo lỗi.");
    IS_SUBMITTING = false;
  }
});
