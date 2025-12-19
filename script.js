/* =========================================================
   LIMITS
========================================================= */
const MAX_CARDS_PER_ORDER = 10;
const MAX_BILLS_PER_CARD = 10;

/* =========================================================
   GOOGLE SHEET WEB APP (CẬP NHẬT 2 DÒNG NÀY)
========================================================= */
const GOOGLE_SHEET_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbyeJqGboIntGGg4l28EyzRl4zHdQkftp6lbns14czS83Z24Ym5uC8iUztaGnE2LfOtS/exec";
const GOOGLE_SHEET_SECRET = "1"; // trùng SECRET trong Apps Script

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

  let pos = 0, seen = 0;
  while (pos < el.value.length && seen < digitsBefore) {
    if (/\d/.test(el.value[pos])) seen++;
    pos++;
  }
  el.setSelectionRange(pos, pos);
}

function setSelectOptions(selectEl, values, placeholder) {
  if (!selectEl) return;
  selectEl.innerHTML = [`<option value="">${placeholder}</option>`]
    .concat(values.map(v => `<option value="${String(v)}">${String(v)}</option>`))
    .join("");
}

function getCardEls() {
  return Array.from(document.querySelectorAll(".card-item"));
}
function getAllCardIds() {
  return getCardEls().map(s => s.dataset.cardId).filter(Boolean);
}
function getCardCount() {
  return getCardEls().length;
}
function getBillCount(cardId) {
  return document.querySelectorAll(`#billDetails_${cardId} .bill-row`).length;
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
      ...ALL_STAFF.map(n => `<option value="${n}">${n}</option>`),
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
      return;
    }

    const list = [...STAFF_BY_OFFICE[office], "Khách văn phòng"];
    staffEl.disabled = false;
    staffEl.innerHTML = [
      `<option value="">-- Chọn Nhân viên --</option>`,
      ...list.map(n => `<option value="${n}">${n}</option>`),
    ].join("");

    syncStaffFinal();
  });

  staffEl?.addEventListener("change", () => {
    if (staffEl.value === "Khách văn phòng") showContact();
    else hideContact();
    syncStaffFinal();
  });

  contactInput?.addEventListener("input", syncStaffFinal);

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
    ...Object.keys(POS_DATA).map(k => `<option value="${k}">${k}</option>`),
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
  const machines = (pos && hkd && POS_DATA[pos] && POS_DATA[pos][hkd]) ? POS_DATA[pos][hkd] : [];

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
   (3) TÍNH TOÁN THEO DỊCH VỤ
========================================================= */
function normalizeServiceValue(raw) {
  const s = String(raw || "").trim().toUpperCase();
  if (!s) return "";
  if (s.includes("DAO") && s.includes("RUT")) return "DAO_RUT";
  if (s.includes("RUT")) return "RUT";
  if (s.includes("DAO")) return "DAO";
  return s;
}

function getFeePercents() {
  const feePercent = Number(document.getElementById("feePercentAll")?.value || 0);
  const returnFeePercent = Number(document.getElementById("returnFeePercentAll")?.value || 0);
  return { feePercent, returnFeePercent };
}

function getShipFee() {
  return parseCurrencyVND(document.getElementById("shipFee")?.value || 0);
}

function getCardBillTotal(cardId) {
  const billWrap = document.getElementById(`billDetails_${cardId}`);
  let totalBill = 0;

  if (billWrap) {
    billWrap.querySelectorAll(".bill-amount").forEach(inp => {
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

// Rule (2): RÚT => auto Tiền chuyển = Tiền làm - Tiền phí
function syncTransferForWithdraw(cardId, totalBill, feeBase) {
  const serviceRaw = document.getElementById(`serviceType_${cardId}`)?.value || "";
  const service = normalizeServiceValue(serviceRaw);

  const transferEl = document.getElementById(`transferAmount_${cardId}`);
  if (!transferEl) return;

  if (service === "RUT") {
    const autoTransfer = Math.max(0, totalBill - feeBase);
    lockTransferInput(transferEl, true);

    const cur = parseCurrencyVND(transferEl.value);
    if (cur !== autoTransfer) transferEl.value = formatVND(autoTransfer);
  } else {
    lockTransferInput(transferEl, false);
  }
}

function getCardTotals(cardId) {
  const totalBill = getCardBillTotal(cardId);

  const { feePercent, returnFeePercent } = getFeePercents();
  const feeBase = Math.round(totalBill * feePercent / 100);
  const returnFee = Math.round(totalBill * returnFeePercent / 100);

  // sync transfer nếu là RÚT
  syncTransferForWithdraw(cardId, totalBill, feeBase);

  const transfer = parseCurrencyVND(document.getElementById(`transferAmount_${cardId}`)?.value);
  const serviceRaw = document.getElementById(`serviceType_${cardId}`)?.value || "";
  const service = normalizeServiceValue(serviceRaw);

  // phí theo thẻ:
  // - DAO: feeBase
  // - RUT: feeBase
  // - DAO_RUT: feeBase + returnFee
  const feeForCard = service === "DAO_RUT" ? (feeBase + returnFee) : feeBase;

  return { cardId, totalBill, transfer, service, feeBase, returnFee, feeForCard };
}

function updateCardMetricsUI(cardId, totalBill, transfer) {
  document.getElementById(`totalBillAmount_${cardId}`)?.replaceChildren(document.createTextNode(formatVND(totalBill)));
  document.getElementById(`differenceAmount_${cardId}`)?.replaceChildren(document.createTextNode(formatVND(totalBill - transfer)));
}

function calcDAO(card) {
  // (1) DAO: phí = tiền chuyển - tiền làm + feeBase
  const feeCalc = card.transfer - card.totalBill + card.feeBase;
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

  let totalBillAll = 0;
  let totalTransferAll = 0;

  let totalFeeAll = 0;     // tổng phí theo từng thẻ (DAO_RUT có returnFee)
  let totalFeeBaseAll = 0; // tổng feeBase

  const servicesSet = new Set();
  const cards = [];

  getAllCardIds().forEach(cardId => {
    const c = getCardTotals(cardId);

    updateCardMetricsUI(cardId, c.totalBill, c.transfer);

    totalBillAll += c.totalBill;
    totalTransferAll += c.transfer;

    totalFeeAll += (c.service ? c.feeForCard : 0);
    totalFeeBaseAll += (c.service ? c.feeBase : 0);

    if (c.service) servicesSet.add(c.service);
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

  let totalCollect = 0; // Tổng thu khách
  let totalPay = 0;     // Tổng trả khách

  // (4) Mix nhiều dịch vụ:
  // tổng chuyển - tổng làm + phí + ship
  if (servicesSet.size > 1) {
    const net = totalTransferAll - totalBillAll + totalFeeAll + shipFee;
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
      cards.forEach(c => {
        if (c.service !== "DAO") return;
        const r = calcDAO(c);
        totalCollect += r.collect;
        totalPay += r.pay;
      });
      totalCollect += shipFee;
    }
    else if (onlyService === "RUT") {
      // (2) RUT: thu = tổng feeBase + ship ; trả = tổng chuyển (auto)
      totalCollect = totalFeeBaseAll + shipFee;
      totalPay = totalTransferAll;
    }
    else if (onlyService === "DAO_RUT") {
      totalCollect = totalFeeAll + shipFee;
      cards.forEach(c => {
        if (c.service !== "DAO_RUT") return;
        const r = calcDAO_RUT(c);
        totalPay += r.pay;
      });
    }
    else {
      totalCollect = shipFee;
      totalPay = 0;
    }
  }

  // UI tổng
  const elTotalBillAll = document.getElementById("totalBillAll");
  if (elTotalBillAll) elTotalBillAll.textContent = `${formatVND(totalBillAll)} VNĐ`;

  const elTotalFeeCollected = document.getElementById("totalFeeCollectedAll");
  if (elTotalFeeCollected) elTotalFeeCollected.textContent = `${formatVND(totalCollect)} VNĐ`;

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
  const c = getCardTotals(cardId);
  updateCardMetricsUI(cardId, c.totalBill, c.transfer);
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

      <input type="text" class="bill-batch card-input" name="billBatch_${cardId}_${billIndex}" placeholder="Số lô" />
      <input type="text" class="bill-invoice card-input" name="billInvoice_${cardId}_${billIndex}" placeholder="Số hóa đơn" />

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
    row.querySelectorAll(".currency-input").forEach(el => (el.value = formatVND(parseCurrencyVND(el.value))));
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

    if (label) { label.value = `Bill ${idx}`; label.name = `billLabel_${cardId}_${idx}`; }
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

  cardEl.querySelectorAll("[data-card-id]").forEach(el => {
    el.dataset.cardId = String(newId);
  });

  cardEl.querySelectorAll("[id]").forEach(el => {
    el.id = replaceCardToken(el.id, oldId, newId);
  });

  cardEl.querySelectorAll("[name]").forEach(el => {
    el.name = replaceCardToken(el.name, oldId, newId);
  });

  cardEl.querySelectorAll("label[for]").forEach(lb => {
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
  getAllCardIds().forEach(cardId => {
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

  cardEl.querySelectorAll("input").forEach(inp => {
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

  cardEl.querySelectorAll(".currency-input").forEach(el => {
    el.value = formatVND(parseCurrencyVND(el.value));
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
    case "V": return "Visa";
    case "M": return "MasterCard";
    case "J": return "JCB";
    case "n": return "Napas";
    default: return String(v || "");
  }
}
function serviceToText(v) {
  switch (String(v || "")) {
    case "DAO": return "ĐÁO";
    case "RUT": return "RÚT";
    case "DAO_RUT": return "ĐÁO+RÚT";
    default: return String(v || "");
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
    feePercentAll: Number(document.getElementById("feePercentAll")?.value || 0),
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

    const bills = Array.from(cardEl.querySelectorAll(".bill-row")).map((row) => ({
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

  // init existing bill rows
  document.querySelectorAll(".bill-row").forEach(initBillRow);

  // format currency inputs
  document.querySelectorAll(".currency-input").forEach(el => {
    const allowEmpty = el.id === "actualFeeReceived";
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
  recalcCard("1");
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
  // currency format
  if (e.target.classList.contains("currency-input")) {
    const allowEmpty = e.target.id === "actualFeeReceived";
    formatCurrencyInput(e.target, allowEmpty);
  }

  const cardId = e.target.closest(".card-item")?.dataset.cardId;

  // per-card recalc when transfer/bill amount changes
  if (
    cardId &&
    (e.target.id === `transferAmount_${cardId}` || e.target.classList.contains("bill-amount"))
  ) {
    recalcCard(cardId);
  }

  // section 3 recalcs
  if (
    e.target.id === "feePercentAll" ||
    e.target.id === "returnFeePercentAll" ||
    e.target.id === "shipFee" ||
    e.target.id === "actualFeeReceived"
  ) {
    recalcSummary();
  }
});

// Submit: sync staffFinal + gửi Google Sheet
document.getElementById("mainForm")?.addEventListener("submit", async (e) => {
  e.preventDefault();

  // sync staffFinal lần cuối
  const staffFinal = document.getElementById("staffFinal");
  const staffEl = document.getElementById("staff");
  const contactEl = document.getElementById("customerDetail");

  if (staffFinal) {
    staffFinal.value =
      staffEl?.value === "Khách văn phòng"
        ? (contactEl?.value || "").trim()
        : (staffEl?.value || "");
  }

  if (staffEl?.value === "Khách văn phòng" && !(contactEl?.value || "").trim()) {
    alert("Vui lòng nhập Liên hệ cho Khách văn phòng.");
    contactEl?.focus();
    return;
  }

  // đảm bảo tổng mới nhất
  recalcSummary();

  const payload = collectPayloadForSheet();

  try {
    await sendToGoogleSheet(payload);
    alert("Đã gửi đơn thành công. Em Hằng xin cảm ơn anh chị ạ!");
  } catch (err) {
    console.error(err);
    alert("Gửi đơn thất bại.Liên hệ em Hằng báo lỗi.");
  }
});
