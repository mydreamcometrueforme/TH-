const MAX_CARDS_PER_ORDER = 10;
const MAX_BILLS_PER_CARD = 10;

/* =========================================================
   GOOGLE SHEET WEB APP (CẬP NHẬT 2 DÒNG NÀY)
========================================================= */
const GOOGLE_SHEET_WEBAPP_URL =
  "https://script.google.com/macros/s/AKfycby9GJ3Sk__YX5eE03f2oYd2DazE2ASrEgfrKvzCYbRnOcxFXh7o2Zbfpx8wo5YmXimA/exec";
const GOOGLE_SHEET_SECRET = "THỬ"; // trùng SECRET trong Apps Script

/* =========================================================
   SUBMISSION ID + CHẶN GỬI TRÙNG (CLIENT)
   (bổ sung – không ảnh hưởng phần cũ)
========================================================= */
let IS_SUBMITTING = false;
function makeSubmissionId() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}_${Math.random().toString(16).slice(2)}`;
}
const CURRENT_SUBMISSION_ID = makeSubmissionId();

/* =========================================================
   DATA
========================================================= */
// Nhân viên theo văn phòng
const STAFF_BY_OFFICE = {
  ThaiHa: ["Cường", "Thái", "Thịnh", "Linh", "Trang", "Vượng", "Hoàng anh", "Huy"],
  NguyenXien: ["An", "Kiên", "Trang anh", "Phú", "Trung", "Nam", "Hiệp", "Dương", "Đức anh", "Vinh"],
};
const ALL_STAFF = [...STAFF_BY_OFFICE.ThaiHa, ...STAFF_BY_OFFICE.NguyenXien];

// POS -> HKD -> Máy POS
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

// allowEmpty: true => user xóa hết thì giữ rỗng (không ép về 0)
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

/* =========================================================
   ✅ (6) digits-only helpers (bổ sung)
========================================================= */
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

  // Ship staff: all + Không
  if (shipEl) {
    shipEl.innerHTML = [
      `<option value="Không">Không</option>`,
      ...ALL_STAFF.map((n) => `<option value="${n}">${n}</option>`),
    ].join("");
    if (!shipEl.value) shipEl.value = "Không";
  }

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
    if (!staffEl) return;
    staffEl.disabled = true;
    staffEl.innerHTML = `<option value="">-- Vui lòng chọn Văn phòng trước --</option>`;
    staffEl.value = "";
  }
  function syncStaffFinal() {
    if (!staffFinalEl) return;
    const staffVal = staffEl?.value || "";
    staffFinalEl.value = staffVal === "Khách văn phòng" ? (contactInput?.value || "").trim() : staffVal;
  }

  officeEl?.addEventListener("change", () => {
    const office = officeEl.value;
    hideContact();
    resetStaff();

    if (!office || !STAFF_BY_OFFICE[office]) {
      syncStaffFinal();
      // ✅ (2) auto phí thu về
      updateReturnFeePercentAuto();
      recalcSummary();
      return;
    }

    const list = [...STAFF_BY_OFFICE[office], "Khách văn phòng"];
    staffEl.disabled = false;
    staffEl.innerHTML = [
      `<option value="">-- Chọn Nhân viên --</option>`,
      ...list.map((n) => `<option value="${n}">${n}</option>`),
    ].join("");

    syncStaffFinal();
    // ✅ (2) auto phí thu về
    updateReturnFeePercentAuto();
    recalcSummary();
  });

  staffEl?.addEventListener("change", () => {
    if (staffEl.value === "Khách văn phòng") showContact();
    else hideContact();
    syncStaffFinal();

    // ✅ (2) auto phí thu về
    updateReturnFeePercentAuto();
    recalcSummary();
  });

  contactInput?.addEventListener("input", () => {
    syncStaffFinal();
    // ✅ (2) nếu là khách VP thì phí thu về = phí thu %
    updateReturnFeePercentAuto();
  });

  hideContact();
  resetStaff();
  syncStaffFinal();
}

/* =========================================================
   (2) POS -> HKD -> MÁY (Bill row)
========================================================= */
function initBillRow(rowEl) {
  const posSel = rowEl.querySelector(".pos-select");
  const hkdSel = rowEl.querySelector(".hkd-select");
  const machineSel = rowEl.querySelector(".machine-select");
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
  const posSel = rowEl.querySelector(".pos-select");
  const hkdSel = rowEl.querySelector(".hkd-select");
  const machineSel = rowEl.querySelector(".machine-select");
  if (!posSel || !hkdSel || !machineSel) return;

  const pos = posSel.value;
  const hkds = pos && POS_DATA[pos] ? Object.keys(POS_DATA[pos]) : [];

  setSelectOptions(hkdSel, hkds, "-- Chọn HKD --");
  setSelectOptions(machineSel, [], "-- Chọn Máy POS --");
  machineSel.classList.remove("auto-locked");
}

function onHKDChange(rowEl) {
  const posSel = rowEl.querySelector(".pos-select");
  const hkdSel = rowEl.querySelector(".hkd-select");
  const machineSel = rowEl.querySelector(".machine-select");
  if (!posSel || !hkdSel || !machineSel) return;

  const pos = posSel.value;
  const hkd = hkdSel.value;
  const machines = pos && hkd && POS_DATA[pos] && POS_DATA[pos][hkd] ? POS_DATA[pos][hkd] : [];

  if (pos === "BV") {
    setSelectOptions(machineSel, machines, "-- Chọn Máy POS --");
    machineSel.classList.remove("auto-locked");
    machineSel.value = "";
    return;
  }

  if (machines.length > 0) {
    setSelectOptions(machineSel, machines, "-- Máy POS --");
    machineSel.value = String(machines[0]);
    machineSel.classList.add("auto-locked");
  } else {
    setSelectOptions(machineSel, [], "-- Chọn Máy POS --");
    machineSel.classList.remove("auto-locked");
  }
}

/* =========================================================
   ✅ (1) PHÍ % -> PHÍ CỨNG (bổ sung)
   - Nếu nhập % => tự nhảy phí cứng (format tiền)
   - Tính toán dùng phí cứng
========================================================= */
function getFeePercentInput() {
  const raw = String(document.getElementById("feePercentAll")?.value || "").trim();
  if (!raw) return 0;
  const n = Number(raw);
  return Number.isFinite(n) ? n : 0;
}
function getFeeFixedInput() {
  return parseCurrencyVND(document.getElementById("feeFixedAll")?.value || "");
}
function setFeeFixedInput(v) {
  const el = document.getElementById("feeFixedAll");
  if (!el) return;
  el.value = v > 0 ? formatVND(v) : "";
}
function showFeeFixedGroup(show) {
  const g = document.getElementById("feeFixedGroup");
  if (!g) return; // nếu HTML chưa thêm thì bỏ qua
  g.classList.toggle("hidden", !show);
}
function markFeeFixedManual() {
  const el = document.getElementById("feeFixedAll");
  if (!el) return;
  el.dataset.manual = "1";
  if (parseCurrencyVND(el.value || "") === 0) {
    // nếu user xóa sạch => cho phép auto lại
    delete el.dataset.manual;
  }
}
function syncFeeFixedFromPercent(totalBillAll) {
  const percent = getFeePercentInput();
  const fixedEl = document.getElementById("feeFixedAll");
  const fixedCur = getFeeFixedInput();

  // show/hide group
  showFeeFixedGroup(percent > 0 || fixedCur > 0);

  if (!fixedEl) return;

  // nếu user đã nhập tay phí cứng thì không auto đè
  const isManual = fixedEl.dataset.manual === "1";

  if (percent > 0 && !isManual) {
    const calc = Math.round((totalBillAll * percent) / 100);
    setFeeFixedInput(calc);
  }
}

/* =========================================================
   ✅ (2) PHÍ THU VỀ (%) AUTO (bổ sung)
   - NV => 1.45
   - Khách VP => = phí thu khách (%) (feePercentAll)
========================================================= */
function updateReturnFeePercentAuto() {
  const staffEl = document.getElementById("staff");
  const returnEl = document.getElementById("returnFeePercentAll");
  if (!returnEl) return;

  // auto
  returnEl.readOnly = true;
  returnEl.tabIndex = -1;

  const staffVal = String(staffEl?.value || "");
  if (!staffVal) {
    returnEl.value = "";
    return;
  }

  if (staffVal === "Khách văn phòng") {
    const p = String(document.getElementById("feePercentAll")?.value || "").trim();
    returnEl.value = p; // có thể rỗng
  } else {
    returnEl.value = "1.45";
  }
}

/* =========================================================
   ✅ chia tiền theo tỉ lệ tiền làm (bổ sung)
========================================================= */
function allocateByProportion(total, items, getWeight) {
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
   (3) TÍNH TOÁN THEO DỊCH VỤ
   ✅ sửa: dùng PHÍ CỨNG thay cho phí %
   ✅ sửa: RUT cộng phí thực thu: tiền làm - tiền phí + phí thực thu (chia theo thẻ)
========================================================= */
function normalizeServiceValue(raw) {
  const s = String(raw || "").trim().toUpperCase();
  if (!s) return "";
  if (s.includes("DAO") && s.includes("RUT")) return "DAO_RUT";
  if (s.includes("RUT")) return "RUT";
  if (s.includes("DAO")) return "DAO";
  return s;
}

function getShipFee() {
  return parseCurrencyVND(document.getElementById("shipFee")?.value || "");
}

function getCardBillTotal(cardId) {
  const billWrap = document.getElementById(`billDetails_${cardId}`);
  let totalBill = 0;

  if (billWrap) {
    billWrap.querySelectorAll(".bill-amount").forEach((inp) => {
      totalBill += parseCurrencyVND(inp.value);
    });
  }
  return totalBill;
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

// ✅ Rule (2)+(4): RÚT => auto Tiền chuyển = Tiền làm - Tiền phí + Phí thực thu (chia theo thẻ)
function syncTransferForWithdraw(cardId, totalBill, feeShare, actualShare) {
  const serviceRaw = document.getElementById(`serviceType_${cardId}`)?.value || "";
  const service = normalizeServiceValue(serviceRaw);

  const transferEl = document.getElementById(`transferAmount_${cardId}`);
  if (!transferEl) return;

  if (service === "RUT") {
    const autoTransfer = Math.max(0, totalBill - (feeShare || 0) + (actualShare || 0));
    lockTransferInput(transferEl, true);

    const cur = parseCurrencyVND(transferEl.value);
    if (cur !== autoTransfer) transferEl.value = autoTransfer > 0 ? formatVND(autoTransfer) : "";
  } else {
    lockTransferInput(transferEl, false);
  }
}

/* ✅ sửa getCardTotals: nhận ctx để lấy feeShare/actualShare */
function getCardTotals(cardId, ctx) {
  const totalBill = getCardBillTotal(cardId);

  const serviceRaw = document.getElementById(`serviceType_${cardId}`)?.value || "";
  const service = normalizeServiceValue(serviceRaw);

  // phí cứng chia theo thẻ
  const feeShare = ctx?.feeShareMap?.get(cardId) || 0;

  // phí thực thu chia theo thẻ RUT
  const actualShare = ctx?.actualShareMap?.get(cardId) || 0;

  // auto transfer nếu RUT
  syncTransferForWithdraw(cardId, totalBill, feeShare, actualShare);

  const transfer = parseCurrencyVND(document.getElementById(`transferAmount_${cardId}`)?.value);

  // phí thu về (chỉ dùng cho DAO_RUT)
  const returnFeePercent = Number(document.getElementById("returnFeePercentAll")?.value || 0);
  const returnFee = Math.round(totalBill * returnFeePercent / 100);

  const feeForCard = service === "DAO_RUT" ? feeShare + returnFee : feeShare;

  return { cardId, totalBill, transfer, service, feeShare, returnFee, feeForCard };
}

function updateCardMetricsUI(cardId, totalBill, transfer) {
  document
    .getElementById(`totalBillAmount_${cardId}`)
    ?.replaceChildren(document.createTextNode(formatVND(totalBill)));
  document
    .getElementById(`differenceAmount_${cardId}`)
    ?.replaceChildren(document.createTextNode(formatVND(totalBill - transfer)));
}

function calcDAO(card) {
  // (1) DAO: phí = tiền chuyển - tiền làm + phí_cứng_share
  const feeCalc = card.transfer - card.totalBill + card.feeShare;
  if (feeCalc >= 0) return { collect: feeCalc, pay: 0 };
  return { collect: 0, pay: Math.abs(feeCalc) };
}

function calcDAO_RUT(card) {
  // (3) DAO+RÚT: |diff| - feeForCard => tiền trả khách
  const remain = Math.abs(card.transfer - card.totalBill) - card.feeForCard;
  return { collect: card.feeForCard, pay: Math.max(0, remain) };
}

function recalcSummary() {
  const shipFee = getShipFee();

  // ===== 1) Tổng bill toàn bộ thẻ =====
  let totalBillAll = 0;
  const meta = getAllCardIds().map((cardId) => {
    const totalBill = getCardBillTotal(cardId);
    totalBillAll += totalBill;
    const serviceRaw = document.getElementById(`serviceType_${cardId}`)?.value || "";
    return { cardId, totalBill, service: normalizeServiceValue(serviceRaw) };
  });

  // ===== 2) sync phí cứng từ % (yêu cầu 1) =====
  syncFeeFixedFromPercent(totalBillAll);

  // ===== 3) phí cứng tổng (ưu tiên feeFixedAll, nếu chưa có thì lấy % tính ra) =====
  const feeFixedInput = getFeeFixedInput();
  const feePercent = getFeePercentInput();
  const feeFixedTotal =
    feeFixedInput > 0 ? feeFixedInput : feePercent > 0 ? Math.round(totalBillAll * feePercent / 100) : 0;

  // ===== 4) chia phí cứng theo tỉ lệ bill cho các thẻ có dịch vụ =====
  const cardsWithService = meta.filter((c) => !!c.service);
  const feeShareMap = allocateByProportion(feeFixedTotal, cardsWithService, (c) => c.totalBill);

  // ===== 5) chia phí thực thu cho các thẻ RÚT (yêu cầu 4) =====
  const actualFeeReceived = parseCurrencyVND(document.getElementById("actualFeeReceived")?.value || "");
  const withdrawCards = meta.filter((c) => c.service === "RUT");
  const actualShareMap = allocateByProportion(actualFeeReceived, withdrawCards, (c) => c.totalBill);

  const ctx = { feeShareMap, actualShareMap };

  let totalTransferAll = 0;

  let totalFeeAll = 0;     // tổng phí theo thẻ (DAO_RUT có returnFee)
  let totalFeeFixedAll = 0; // tổng phí cứng share

  const servicesSet = new Set();
  const cards = [];

  getAllCardIds().forEach((cardId) => {
    const c = getCardTotals(cardId, ctx);

    updateCardMetricsUI(cardId, c.totalBill, c.transfer);

    totalTransferAll += c.transfer;

    if (c.service) {
      totalFeeAll += c.feeForCard;
      totalFeeFixedAll += c.feeShare;
      servicesSet.add(c.service);
    }

    cards.push(c);
  });

  // Thẻ âm hiển thị theo (tổng bill - tổng chuyển) âm => abs
  const totalDiffAll = totalBillAll - totalTransferAll;
  const negativeCardValue = Math.max(0, -totalDiffAll);
  const negativeCardEl = document.getElementById("negativeCardFee");
  if (negativeCardEl) {
    negativeCardEl.readOnly = true;
    negativeCardEl.tabIndex = -1;
    negativeCardEl.value = formatVND(negativeCardValue);
  }

  let totalCollect = 0; // Tổng thu khách (chưa + ship + thẻ âm)
  let totalPay = 0;     // Tổng trả khách

  // (4) Mix nhiều dịch vụ:
  if (servicesSet.size > 1) {
    const net = totalTransferAll - totalBillAll + totalFeeAll;
    if (net >= 0) {
      totalCollect = net;
      totalPay = 0;
    } else {
      totalCollect = 0;
      totalPay = Math.abs(net);
    }
  } else {
    const onlyService = servicesSet.size === 1 ? Array.from(servicesSet)[0] : "";

    if (onlyService === "DAO") {
      cards.forEach((c) => {
        if (c.service !== "DAO") return;
        const r = calcDAO(c);
        totalCollect += r.collect;
        totalPay += r.pay;
      });
    } else if (onlyService === "RUT") {
      // (2) RUT: thu = tổng phí cứng; trả = tổng chuyển (auto)
      totalCollect = totalFeeFixedAll;
      totalPay = totalTransferAll;
    } else if (onlyService === "DAO_RUT") {
      totalCollect = totalFeeAll;
      cards.forEach((c) => {
        if (c.service !== "DAO_RUT") return;
        const r = calcDAO_RUT(c);
        totalPay += r.pay;
      });
    } else {
      totalCollect = 0;
      totalPay = 0;
    }
  }

  // ✅ tổng thu khách = phí thu + ship + thẻ âm
  const totalCollectFinal = totalCollect + shipFee + negativeCardValue;

  // UI tổng
  const elTotalBillAll = document.getElementById("totalBillAll");
  if (elTotalBillAll) elTotalBillAll.textContent = `${formatVND(totalBillAll)} VNĐ`;

  const elTotalFeeCollected = document.getElementById("totalFeeCollectedAll");
  if (elTotalFeeCollected) elTotalFeeCollected.textContent = `${formatVND(totalCollectFinal)} VNĐ`;

  const elTotalPay = document.getElementById("totalCustomerPayment");
  if (elTotalPay) elTotalPay.textContent = `${formatVND(totalPay)} VNĐ`;

  // auto chọn thanh toán phí theo phí thực thu
  const actualFeeInput = document.getElementById("actualFeeReceived");
  const payStatus = document.getElementById("feePaymentStatus");
  if (payStatus) {
    const actual = parseCurrencyVND(actualFeeInput?.value || "");
    payStatus.value = actual > 0 ? "da_thu" : "chua_thu";
  }
}

function recalcCard(cardId) {
  // giữ nguyên: giờ recalcSummary sẽ update luôn
  recalcSummary();
}

function toggleServiceDetails(cardId) {
  const service = document.getElementById(`serviceType_${cardId}`)?.value || "";
  const container = document.getElementById(`serviceDetails_${cardId}`);
  if (container) {
    if (service) container.classList.remove("hidden");
    else container.classList.add("hidden");
  }
  recalcCard(cardId);
}

/* =========================================================
   BILL: ADD / REMOVE / RENUMBER
========================================================= */
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
  const wrapper = document.getElementById(`billDetails_${cardId}`);
  if (!wrapper) return;

  const currentCount = wrapper.querySelectorAll(".bill-row").length;
  if (currentCount >= MAX_BILLS_PER_CARD) {
    alert(`Mỗi thẻ tối đa ${MAX_BILLS_PER_CARD} bill.`);
    updateBillLimitUI(cardId);
    return;
  }

  const nextIndex = currentCount + 1;
  wrapper.insertAdjacentHTML("beforeend", billRowMarkup(cardId, nextIndex));

  const row = wrapper.querySelector(`.bill-row[data-bill-index="${nextIndex}"]`);
  if (row) {
    initBillRow(row);
    row.querySelectorAll(".currency-input").forEach((el) => {
      // ✅ bill-amount cho phép rỗng nếu user xoá (nhưng giữ logic cũ)
      formatCurrencyInput(el, el.classList.contains("bill-amount"));
    });
  }

  updateBillLimitUI(cardId);
  recalcCard(cardId);
}

function renumberBills(cardId) {
  const wrapper = document.getElementById(`billDetails_${cardId}`);
  if (!wrapper) return;

  const rows = wrapper.querySelectorAll(".bill-row");
  rows.forEach((row, i) => {
    const idx = i + 1;
    row.dataset.billIndex = String(idx);

    const label = row.querySelector(".bill-label");
    const amount = row.querySelector(".bill-amount");
    const pos = row.querySelector(".pos-select");
    const hkd = row.querySelector(".hkd-select");
    const machine = row.querySelector(".machine-select");
    const batch = row.querySelector(".bill-batch");
    const invoice = row.querySelector(".bill-invoice");

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

  updateBillLimitUI(cardId);
}

/* =========================================================
   ✅ REINDEX CARD IDs (sau khi xóa thẻ)
========================================================= */
function replaceCardToken(str, oldId, newId) {
  if (!str) return str;
  str = str.replaceAll(`_${oldId}_`, `_${newId}_`);
  str = str.replace(new RegExp(`_${oldId}$`), `_${newId}`);
  return str;
}

function updateCardIdentifiers(cardEl, oldId, newId) {
  if (!cardEl) return;

  cardEl.dataset.cardId = String(newId);

  const h2 = cardEl.querySelector("h2");
  if (h2) h2.textContent = `2. Thông tin Thẻ #${newId}`;

  cardEl.querySelectorAll("[data-card-id]").forEach((el) => {
    el.dataset.cardId = String(newId);
  });

  cardEl.querySelectorAll("[id]").forEach((el) => {
    el.id = replaceCardToken(el.id, oldId, newId);
  });

  cardEl.querySelectorAll("[name]").forEach((el) => {
    el.name = replaceCardToken(el.name, oldId, newId);
  });

  cardEl.querySelectorAll("label[for]").forEach((lb) => {
    lb.setAttribute("for", replaceCardToken(lb.getAttribute("for"), oldId, newId));
  });
}

function reindexCards() {
  const cards = getCardEls();

  // pass 1: đổi qua TMP để tránh trùng id
  cards.forEach((cardEl, i) => {
    const oldId = cardEl.dataset.cardId;
    updateCardIdentifiers(cardEl, oldId, `TMP${i + 1}`);
  });

  // pass 2: TMP -> 1..n
  const cards2 = getCardEls();
  cards2.forEach((cardEl, i) => {
    const tmpId = cardEl.dataset.cardId;
    updateCardIdentifiers(cardEl, tmpId, String(i + 1));
  });

  // sau reindex: renumber bills + init rows + recalc
  getAllCardIds().forEach((cardId) => {
    const wrapper = document.getElementById(`billDetails_${cardId}`);
    wrapper?.querySelectorAll(".bill-row").forEach(initBillRow);

    renumberBills(cardId);
    toggleServiceDetails(cardId);
    recalcCard(cardId);
  });

  updateCardLimitUI();
  recalcSummary();
}

/* =========================================================
   ADD CARD
========================================================= */
function resetCardValues(cardEl) {
  const cardId = cardEl.dataset.cardId;

  cardEl.querySelectorAll("input").forEach((inp) => {
    if (inp.classList.contains("bill-label")) return;
    if (inp.classList.contains("currency-input")) inp.value = "0";
    else inp.value = "";
  });

  const serviceSel = cardEl.querySelector(".service-selector");
  if (serviceSel) serviceSel.value = "";

  cardEl.querySelector(".service-details-container")?.classList.add("hidden");

  const wrapper = cardEl.querySelector(`#billDetails_${cardId}`);
  if (wrapper) {
    wrapper.innerHTML = billRowMarkup(cardId, 1);
    wrapper.querySelectorAll(".bill-row").forEach(initBillRow);
  }

  cardEl.querySelectorAll(".currency-input").forEach((el) => {
    const allowEmpty =
      el.id === "actualFeeReceived" ||
      el.id === "shipFee" ||
      el.id === "feeFixedAll" ||
      String(el.id || "").startsWith("transferAmount_") ||
      el.classList.contains("bill-amount");
    formatCurrencyInput(el, allowEmpty);
  });

  const t = cardEl.querySelector(`#totalBillAmount_${cardId}`);
  if (t) t.textContent = "0";
  const d = cardEl.querySelector(`#differenceAmount_${cardId}`);
  if (d) d.textContent = "0";

  updateBillLimitUI(cardId);
}

function addNewCard() {
  const container = document.getElementById("cardContainer");
  const template = document.querySelector(`.card-item[data-card-id="1"]`);
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
  updateBillLimitUI(newId);
  recalcSummary();
}

/* =========================================================
   GOOGLE SHEET PAYLOAD (THÊM POS/HKD/MÁY/LÔ/HĐ)
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
  const office = document.getElementById("office")?.value || "";

  const order = {
    office,
    date: document.getElementById("date")?.value || "",
    staffFinal: document.getElementById("staffFinal")?.value || "",
    staffShip: document.getElementById("staffShip")?.value || "",
  };

  const feePaymentStatus = document.getElementById("feePaymentStatus")?.value || "chua_thu";

  const summary = {
    feePercentAll: Number(document.getElementById("feePercentAll")?.value || 0), // có thể 0 nếu bỏ trống
    // ✅ bổ sung gửi phí cứng (nếu Apps Script cần)
    feeFixedAll: getFeeFixedInput(),

    returnFeePercentAll: Number(document.getElementById("returnFeePercentAll")?.value || 0),
    shipFee: parseCurrencyVND(document.getElementById("shipFee")?.value || 0),
    negativeCardFee: parseCurrencyVND(document.getElementById("negativeCardFee")?.value || 0),

    totalFeeCollectedAll: parseCurrencyVND(document.getElementById("totalFeeCollectedAll")?.textContent || 0),
    actualFeeReceived: parseCurrencyVND(document.getElementById("actualFeeReceived")?.value || ""),

    feePaymentStatus,
    feePaymentStatusText: feeStatusToText(feePaymentStatus),
  };

  const cards = getCardEls().map((cardEl) => {
    const cardId = cardEl.dataset.cardId;

    const billRows = Array.from(cardEl.querySelectorAll(".bill-row"));
    const bills = billRows.map((row, idx) => ({
      // ✅ (5) cờ để Apps Script “chỉ hiện 1 lần” (nếu cậu dùng)
      isFirstOfCard: idx === 0,

      amount: parseCurrencyVND(row.querySelector(".bill-amount")?.value || 0),
      pos: row.querySelector(".pos-select")?.value || "",
      hkd: row.querySelector(".hkd-select")?.value || "",
      machine: row.querySelector(".machine-select")?.value || "",
      batch: row.querySelector(".bill-batch")?.value || "",
      invoice: row.querySelector(".bill-invoice")?.value || "",
    }));

    const cardType = document.getElementById(`cardType_${cardId}`)?.value || "";
    const serviceType = document.getElementById(`serviceType_${cardId}`)?.value || "";

    return {
      cardName: document.getElementById(`cardName_${cardId}`)?.value || "",
      cardNumber: document.getElementById(`cardNumber_${cardId}`)?.value || "",
      cardType: cardTypeToText(cardType),
      cardBank: document.getElementById(`cardBank_${cardId}`)?.value || "",
      serviceType: serviceToText(serviceType),

      transferAmount: parseCurrencyVND(document.getElementById(`transferAmount_${cardId}`)?.value || 0),
      totalBill: parseCurrencyVND(document.getElementById(`totalBillAmount_${cardId}`)?.textContent || 0),

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
   INIT + EVENTS
========================================================= */
document.addEventListener("DOMContentLoaded", () => {
  // set date today
  const dateInput = document.getElementById("date");
  if (dateInput && !dateInput.value) {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    dateInput.value = `${yyyy}-${mm}-${dd}`;
  }

  setupOfficeStaffLogic();

  // ✅ (2) auto phí thu về (set readonly)
  updateReturnFeePercentAuto();

  // init existing bill rows
  document.querySelectorAll(".bill-row").forEach(initBillRow);

  // format currency inputs
  document.querySelectorAll(".currency-input").forEach((el) => {
    const allowEmpty =
      el.id === "actualFeeReceived" ||
      el.id === "shipFee" ||
      el.id === "feeFixedAll" ||
      String(el.id || "").startsWith("transferAmount_") ||
      el.classList.contains("bill-amount");
    formatCurrencyInput(el, allowEmpty);
  });

  // thẻ âm readonly
  const neg = document.getElementById("negativeCardFee");
  if (neg) {
    neg.readOnly = true;
    neg.tabIndex = -1;
  }

  // init calculations
  toggleServiceDetails("1");
  recalcSummary();

  updateCardLimitUI();
  getAllCardIds().forEach(updateBillLimitUI);
});

document.addEventListener("click", (e) => {
  // add card
  if (e.target.id === "addCardBtn") {
    addNewCard();
    return;
  }

  // remove card (chặn khi còn 1 thẻ)
  if (e.target.classList.contains("remove-card-btn")) {
    if (getCardCount() <= 1) {
      alert("Mỗi đơn phải có ít nhất 1 thẻ. Không thể xóa thêm.");
      return;
    }

    const cardId = e.target.dataset.cardId;
    document.querySelector(`.card-item[data-card-id="${cardId}"]`)?.remove();
    reindexCards();
    return;
  }

  // add bill
  if (e.target.classList.contains("add-bill-btn")) {
    addBillRow(e.target.dataset.cardId);
    return;
  }

  // remove bill (giữ tối thiểu 1 bill)
  if (e.target.classList.contains("remove-bill-btn")) {
    const row = e.target.closest(".bill-row");
    const card = e.target.closest(".card-item");
    const cardId = card?.dataset.cardId;
    if (!row || !cardId) return;

    if (getBillCount(cardId) <= 1) {
      alert("Mỗi thẻ cần tối thiểu 1 bill.");
      return;
    }

    row.remove();
    renumberBills(cardId);
    recalcCard(cardId);
  }
});

document.addEventListener("change", (e) => {
  // service selector
  if (e.target.classList.contains("service-selector")) {
    const cardId = e.target.closest(".card-item")?.dataset.cardId;
    if (cardId) toggleServiceDetails(cardId);
  }

  // POS/HKD select in bill row
  const row = e.target.closest(".bill-row");
  if (row) {
    if (e.target.classList.contains("pos-select")) onPosChange(row);
    if (e.target.classList.contains("hkd-select")) onHKDChange(row);
  }
});

document.addEventListener("input", (e) => {
  // ✅ (6) digits-only: 4 số thẻ / số lô / số hóa đơn
  if (String(e.target?.id || "").startsWith("cardNumber_")) {
    e.target.value = digitsOnly(e.target.value).slice(0, 4);
  }
  if (e.target.classList?.contains("integer-only")) {
    e.target.value = digitsOnly(e.target.value);
  }

  // currency format
  if (e.target.classList.contains("currency-input")) {
    const allowEmpty =
      e.target.id === "actualFeeReceived" ||
      e.target.id === "shipFee" ||
      e.target.id === "feeFixedAll" ||
      String(e.target.id || "").startsWith("transferAmount_") ||
      e.target.classList.contains("bill-amount");
    formatCurrencyInput(e.target, allowEmpty);
  }

  // ✅ (1) nhập % => nhảy phí cứng + recalc
  if (e.target.id === "feePercentAll") {
    updateReturnFeePercentAuto(); // (2) nếu khách VP thì phí thu về = phí %
    recalcSummary();
    return;
  }

  // ✅ (1) user nhập tay phí cứng => mark manual
  if (e.target.id === "feeFixedAll") {
    markFeeFixedManual();
    recalcSummary();
    return;
  }

  const cardId = e.target.closest(".card-item")?.dataset.cardId;

  // per-card recalc when transfer/bill amount changes
  if (cardId && (e.target.id === `transferAmount_${cardId}` || e.target.classList.contains("bill-amount"))) {
    recalcCard(cardId);
  }

  // section 3 recalcs
  if (
    e.target.id === "returnFeePercentAll" || // thực tế readonly, nhưng cứ để
    e.target.id === "shipFee" ||
    e.target.id === "actualFeeReceived"
  ) {
    recalcSummary();
  }
});

// Submit: sync staffFinal + gửi Google Sheet
document.getElementById("mainForm")?.addEventListener("submit", async (e) => {
  e.preventDefault();

  if (IS_SUBMITTING) return;
  IS_SUBMITTING = true;

  // sync staffFinal lần cuối
  const staffFinal = document.getElementById("staffFinal");
  const staffEl = document.getElementById("staff");
  const contactEl = document.getElementById("customerDetail");

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

  // ✅ (6) validate 4 số đuôi thẻ, số lô, số hoá đơn là số nguyên
  for (const cardId of getAllCardIds()) {
    const v = String(document.getElementById(`cardNumber_${cardId}`)?.value || "").trim();
    if (!/^\d{4}$/.test(v)) {
      alert("4 số đuôi thẻ phải là số nguyên đúng 4 chữ số.");
      document.getElementById(`cardNumber_${cardId}`)?.focus();
      IS_SUBMITTING = false;
      return;
    }

    const rows = document.querySelectorAll(`#billDetails_${cardId} .bill-row`);
    for (const row of rows) {
      const batch = String(row.querySelector(".bill-batch")?.value || "").trim();
      const inv = String(row.querySelector(".bill-invoice")?.value || "").trim();

      if (batch && !/^\d+$/.test(batch)) {
        alert("Số lô phải là số nguyên.");
        row.querySelector(".bill-batch")?.focus();
        IS_SUBMITTING = false;
        return;
      }
      if (inv && !/^\d+$/.test(inv)) {
        alert("Số hóa đơn phải là số nguyên.");
        row.querySelector(".bill-invoice")?.focus();
        IS_SUBMITTING = false;
        return;
      }
    }
  }

  // đảm bảo tổng mới nhất
  updateReturnFeePercentAuto();
  recalcSummary();

  const payload = collectPayloadForSheet();

  try {
    await sendToGoogleSheet(payload);
    alert("Đã gửi đơn thành công. Em Hằng xin cảm ơn anh chị ạ!");
  } catch (err) {
    console.error(err);
    alert("Gửi đơn thất bại.Liên hệ em Hằng báo lỗi.");
    IS_SUBMITTING = false;
  }
});

