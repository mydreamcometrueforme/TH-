/* =========================================================
   LIMITS
========================================================= */
const MAX_CARDS_PER_ORDER = 10;
const MAX_BILLS_PER_CARD = 10;

/* =========================================================
   GOOGLE SHEET WEB APP
========================================================= */
const GOOGLE_SHEET_WEBAPP_URL =
  "https://script.google.com/macros/s/AKfycbyeJqGboIntGGg4l28EyzRl4zHdQkftp6lbns14czS83Z24Ym5uC8iUztaGnE2LfOtS/exec";
const GOOGLE_SHEET_SECRET = "THỬ"; // ⚠️ phải trùng SECRET trong Apps Script

/* =========================================================
   SUBMIT GUARD (CHẶN GỬI NHIỀU LẦN)
========================================================= */
let IS_SUBMITTING = false;

function makeSubmissionId() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}_${Math.random().toString(16).slice(2)}`;
}
// ✅ 1 đơn / 1 ID cố định (reload trang sẽ có ID mới)
const CURRENT_SUBMISSION_ID = makeSubmissionId();

/* =========================================================
   DATA
========================================================= */
const STAFF_BY_OFFICE = {
  ThaiHa: ["Cường", "Thái", "Thịnh", "Linh", "Trang", "Vượng", "Hoàng anh", "Huy"],
  NguyenXien: ["An", "Kiên", "Trang anh", "Phú", "Trung", "Nam", "Hiệp", "Dương", "Đức anh", "Vinh"],
};
const ALL_STAFF = [...STAFF_BY_OFFICE.ThaiHa, ...STAFF_BY_OFFICE.NguyenXien];

const POS_DATA = {
  BV: {
    "THU TRANG 92A": ["1077", "8244", "1076", "8243"],
    "HONG QUAN": ["1732", "9318", "1731", "9317"],
    "XUAN HUNG": ["1864", "9426", "1865", "9427"],
  },
  AB: {
    "NGOC QUYNH JK M1": ["47"],
    "THIEN PHONG 83 M1": ["51"],
  },
  MB: {
    "MANH THANG - 1": ["T1"],
    "LƯƠNG TUYẾT LAN 1": ["L1"],
    "LƯƠNG TUYẾT LAN 2": ["L2"],
  },
  MBV: {
    "DUC MANH 1": ["DM1"],
    "LONG HA 1": ["LH1"],
  },
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
   HELPERS
========================================================= */
function parseCurrencyVND(str) {
  const digits = String(str ?? "").replace(/[^\d]/g, "");
  return digits ? Math.round(Number(digits)) : 0;
}
function formatVND(n) {
  return Number(n || 0).toLocaleString("vi-VN");
}

function formatCurrencyInput(el, allowEmpty = false) {
  const rawText = String(el.value ?? "");
  const onlyDigits = rawText.replace(/[^\d]/g, "");

  if (allowEmpty && onlyDigits.length === 0) {
    el.value = "";
    return;
  }

  const start = el.selectionStart ?? rawText.length;
  const before = rawText.slice(0, start);
  const digitsBefore = before.replace(/[^\d]/g, "").length;

  const raw = parseCurrencyVND(rawText);
  el.value = formatVND(raw);

  let pos = 0,
    seen = 0;
  while (pos < el.value.length && seen < digitsBefore) {
    if (/\d/.test(el.value[pos])) seen++;
    pos++;
  }
  el.setSelectionRange(pos, pos);
}

function setSelectOptions(selectEl, values, placeholder) {
  if (!selectEl) return;
  selectEl.innerHTML = [`<option value="">${placeholder}</option>`]
    .concat(values.map((v) => `<option value="${String(v)}">${String(v)}</option>`))
    .join("");
}

function getCardEls() {
  return Array.from(document.querySelectorAll(".card-item"));
}
function getAllCardIds() {
  return getCardEls().map((s) => s.dataset.cardId).filter(Boolean);
}
function getCardCount() {
  return getCardEls().length;
}
function getBillCount(cardId) {
  return document.querySelectorAll(`#billDetails_${cardId} .bill-row`).length;
}

function isTransferEl(el) {
  return String(el?.id || "").startsWith("transferAmount_");
}
function isBillAmountEl(el) {
  return !!el?.classList?.contains("bill-amount");
}
function isShipFeeEl(el) {
  return el?.id === "shipFee";
}
function isFeeFixedEl(el) {
  return el?.id === "feeFixedAll";
}
function currencyAllowEmpty(el) {
  return (
    isBillAmountEl(el) ||
    isTransferEl(el) ||
    isShipFeeEl(el) ||
    isFeeFixedEl(el) ||
    el?.id === "actualFeeReceived"
  );
}

function digitsOnly(str) {
  return String(str || "").replace(/[^\d]/g, "");
}

/* =========================================================
   UI LIMIT STATES
========================================================= */
function updateCardLimitUI() {
  const btn = document.getElementById("addCardBtn");
  if (!btn) return;
  btn.disabled = getCardCount() >= MAX_CARDS_PER_ORDER;
  btn.title = btn.disabled ? `Tối đa ${MAX_CARDS_PER_ORDER} thẻ/đơn` : "";
}
function updateBillLimitUI(cardId) {
  const card = document.querySelector(`.card-item[data-card-id="${cardId}"]`);
  if (!card) return;
  const btn = card.querySelector(".add-bill-btn");
  if (!btn) return;

  btn.disabled = getBillCount(cardId) >= MAX_BILLS_PER_CARD;
  btn.title = btn.disabled ? `Tối đa ${MAX_BILLS_PER_CARD} bill/thẻ` : "";
}

/* =========================================================
   (1) OFFICE -> STAFF
========================================================= */
function setupOfficeStaffLogic() {
  const form = document.getElementById("mainForm");

  const officeEl = document.getElementById("office");
  const staffEl = document.getElementById("staff");
  const shipEl = document.getElementById("staffShip");

  const contactField = document.getElementById("customerDetailField");
  const contactInput = document.getElementById("customerDetail");

  const contactLabel = contactField?.querySelector("label");
  if (contactLabel) contactLabel.textContent = "Liên hệ:";
  if (contactInput) contactInput.placeholder = "Nhập tên/SĐT liên hệ";

  // hidden field staffFinal
  let staffFinalEl = document.getElementById("staffFinal");
  if (!staffFinalEl) {
    staffFinalEl = document.createElement("input");
    staffFinalEl.type = "hidden";
    staffFinalEl.id = "staffFinal";
    staffFinalEl.name = "staffFinal";
    form?.appendChild(staffFinalEl);
  }

  if (shipEl) {
    shipEl.innerHTML = [`<o]()
