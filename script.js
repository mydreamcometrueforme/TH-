/*    LIMITS */
const MAX_CARDS_PER_ORDER = 10;
const MAX_BILLS_PER_CARD = 10;

/*   GOOGLE SHEET WEB APP */
const GOOGLE_SHEET_WEBAPP_URL =
  "https://script.google.com/macros/s/AKfycbyeJqGboIntGGg4l28EyzRl4zHdQkftp6lbns14czS83Z24Ym5uC8iUztaGnE2LfOtS/exec";
const GOOGLE_SHEET_SECRET = "THỬ"; // phải trùng SECRET trong Apps Script

/*    SUBMIT GUARD (CHẶN GỬI NHIỀU LẦN) */
let IS_SUBMITTING = false;

function makeSubmissionId() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}_${Math.random().toString(16).slice(2)}`;
}
// ✅ 1 đơn / 1 ID cố định (reload trang sẽ có ID mới)
const CURRENT_SUBMISSION_ID = makeSubmissionId();

/*    DATA */
const STAFF_BY_OFFICE = {
  ThaiHa: ["Cường", "Thái", "Thịnh", "Linh", "Trang", "Vượng", "Hoàng Anh", "Huy"],
  NguyenXien: ["An", "Kiên", "Trang Anh", "Phú", "Trung", "Nam", "Hiệp", "Dương", "Đức Anh", "Vinh"],
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

/*  HELPERS */
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

/*    UI LIMIT STATES */
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

/*    (1) OFFICE -> STAFF */
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
    shipEl.innerHTML = [`<option value="Không">Không</option>`, ...ALL_STAFF.map((n) => `<option value="${n}">${n}</option>`)].join(
      ""
    );
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
      updateReturnFeePercentAuto();
      return;
    }

    const list = [...STAFF_BY_OFFICE[office], "Khách văn phòng"];
    staffEl.disabled = false;
    staffEl.innerHTML = [`<option value="">-- Chọn Nhân viên --</option>`, ...list.map((n) => `<option value="${n}">${n}</option>`)].join("");

    syncStaffFinal();
    updateReturnFeePercentAuto();
  });

  staffEl?.addEventListener("change", () => {
    if (staffEl.value === "Khách văn phòng") showContact();
    else hideContact();
    syncStaffFinal();
    updateReturnFeePercentAuto();
    recalcSummary();
  });

  contactInput?.addEventListener("input", () => {
    syncStaffFinal();
    updateReturnFeePercentAuto();
  });

  hideContact();
  resetStaff();
  syncStaffFinal();
}

/*   (2) POS -> HKD -> MÁY (Bill row) */
function initBillRow(rowEl) {
  const posSel = rowEl.querySelector(".pos-select");
  const hkdSel = rowEl.querySelector(".hkd-select");
  const machineSel = rowEl.querySelector(".machine-select");
  if (!posSel || !hkdSel || !machineSel) return;

  posSel.innerHTML = [`<option value="">-- Chọn POS --</option>`, ...Object.keys(POS_DATA).map((k) => `<option value="${k}">${k}</option>`)].join("");

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
   (3) PHÍ: % -> PHÍ CỨNG + PHÍ THU VỀ AUTO
========================================================= */
function ensureFeeFixedUI() {
  const feePercentEl = document.getElementById("feePercentAll");
  if (!feePercentEl) return;

  // Nếu đã tồn tại thì thôi
  if (document.getElementById("feeFixedAll")) return;

  const percentGroup = feePercentEl.closest(".form-group");
  if (!percentGroup) return;

  const wrap = document.createElement("div");
  wrap.className = "form-group hidden";
  wrap.id = "feeFixedGroup";

  wrap.innerHTML = `
    <label for="feeFixedAll">Phí cứng thu khách:</label>
    <input
      type="text"
      id="feeFixedAll"
      name="feeFixedAll"
      value=""
      class="card-input currency-input"
      inputmode="numeric"
      placeholder="Nhập phí cứng (tự động từ %)"
    />
    <small style="display:block;color:#666;margin-top:6px;">
      Nhập % sẽ tự ra phí cứng. Có thể sửa tay phí cứng, mọi tính toán dùng phí cứng.
    </small>
  `;

  percentGroup.insertAdjacentElement("afterend", wrap);
}

function showFeeFixedIfNeeded() {
  const percentEl = document.getElementById("feePercentAll");
  const fixedGroup = document.getElementById("feeFixedGroup");
  const fixedEl = document.getElementById("feeFixedAll");
  if (!fixedGroup || !percentEl || !fixedEl) return;

  const hasPercent = String(percentEl.value || "").trim() !== "";
  const hasFixed = parseCurrencyVND(fixedEl.value) > 0;

  if (hasPercent || hasFixed) fixedGroup.classList.remove("hidden");
  else fixedGroup.classList.add("hidden");
}

function getFeePercentInput() {
  const v = String(document.getElementById("feePercentAll")?.value || "").trim();
  if (v === "") return 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function getFeeFixedInput() {
  return parseCurrencyVND(document.getElementById("feeFixedAll")?.value || "");
}

// ✅ nếu có % thì tự nhảy phí cứng = tổng tiền làm * %
function syncFeeFixedFromPercent(totalBillAll) {
  const percent = getFeePercentInput();
  const fixedEl = document.getElementById("feeFixedAll");
  if (!fixedEl) return;

  if (percent > 0) {
    const calc = Math.round((totalBillAll * percent) / 100);
    // chỉ auto set khi user chưa gõ tay hoặc đang trống/0
    const cur = parseCurrencyVND(fixedEl.value || "");
    if (cur === 0) fixedEl.value = formatVND(calc);
    showFeeFixedIfNeeded();
  } else {
    showFeeFixedIfNeeded();
  }
}

function updateReturnFeePercentAuto() {
  const staffEl = document.getElementById("staff");
  const feePercentEl = document.getElementById("feePercentAll");
  const returnEl = document.getElementById("returnFeePercentAll");
  if (!returnEl) return;

  const staffVal = String(staffEl?.value || "");
  const feePercentVal = String(feePercentEl?.value || "").trim();

  // luôn auto
  returnEl.readOnly = true;
  returnEl.tabIndex = -1;

  if (!staffVal) {
    returnEl.value = "";
    return;
  }

  if (staffVal === "Khách văn phòng") {
    // khách văn phòng => phí thu về (%) = phí thu khách (%)
    returnEl.value = feePercentVal; // có thể rỗng
  } else {
    // nhân viên => 1.45
    returnEl.value = "1.45";
  }
}

/* =========================================================
   (4) TÍNH TOÁN THEO DỊCH VỤ (DÙNG PHÍ CỨNG)
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

function getReturnFeePercent() {
  const v = String(document.getElementById("returnFeePercentAll")?.value || "").trim();
  if (v === "") return 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function getActualFeeReceived() {
  return parseCurrencyVND(document.getElementById("actualFeeReceived")?.value || "");
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

function updateCardMetricsUI(cardId, totalBill, transfer) {
  document.getElementById(`totalBillAmount_${cardId}`)?.replaceChildren(document.createTextNode(formatVND(totalBill)));
  document.getElementById(`differenceAmount_${cardId}`)?.replaceChildren(document.createTextNode(formatVND(totalBill - transfer)));
}

function calcDAO(card) {
  // DAO: phí = tiền chuyển - tiền làm + phí_cứng_share
  const feeCalc = card.transfer - card.totalBill + card.feeShare;
  if (feeCalc >= 0) return { collect: feeCalc, pay: 0 };
  return { collect: 0, pay: Math.abs(feeCalc) };
}

function calcDAO_RUT(card) {
  // DAO+RUT: |diff| - (phí_cứng_share + phí_thu_về) => trả khách
  const remain = Math.abs(card.transfer - card.totalBill) - (card.feeShare + card.returnFee);
  return { collect: card.feeShare + card.returnFee, pay: Math.max(0, remain) };
}

// ✅ chia phí cứng theo tỉ lệ tiền làm
function allocateByProportion(total, items, getWeight, setValueKey) {
  const list = items.filter((x) => getWeight(x) > 0);
  const sumW = list.reduce((s, x) => s + getWeight(x), 0);

  if (total <= 0 || sumW <= 0) {
    items.forEach((x) => (x[setValueKey] = 0));
    return;
  }

  let used = 0;
  for (let i = 0; i < list.length; i++) {
    const x = list[i];
    if (i === list.length - 1) {
      x[setValueKey] = Math.max(0, total - used);
    } else {
      const share = Math.round((total * getWeight(x)) / sumW);
      x[setValueKey] = Math.max(0, share);
      used += x[setValueKey];
    }
  }

  // items có weight=0
  items.forEach((x) => {
    if (!list.includes(x)) x[setValueKey] = 0;
  });
}

function recalcSummary() {
  // 1) gom basic card
  const cards = getAllCardIds().map((cardId) => {
    const totalBill = getCardBillTotal(cardId);
    const serviceRaw = document.getElementById(`serviceType_${cardId}`)?.value || "";
    const service = normalizeServiceValue(serviceRaw);

    const transferEl = document.getElementById(`transferAmount_${cardId}`);
    const transfer = parseCurrencyVND(transferEl?.value || "");

    return {
      cardId,
      totalBill,
      service,
      transferEl,
      transfer,
      feeShare: 0,
      returnFee: 0,
      actualShare: 0,
    };
  });

  const totalBillAll = cards.reduce((s, c) => s + c.totalBill, 0);

  // 2) % -> phí cứng (tự nhảy)
  ensureFeeFixedUI();
  syncFeeFixedFromPercent(totalBillAll);
  showFeeFixedIfNeeded();

  // 3) lấy phí cứng thực dùng (ưu tiên phí cứng)
  let feeFixedTotal = getFeeFixedInput();
  if (feeFixedTotal <= 0) {
    const percent = getFeePercentInput();
    feeFixedTotal = percent > 0 ? Math.round((totalBillAll * percent) / 100) : 0;
  }

  // 4) phí thu về (%)
  const returnFeePercent = getReturnFeePercent();

  // 5) chia phí cứng theo tỉ lệ tiền làm (chỉ các thẻ có chọn dịch vụ)
  const cardsWithService = cards.filter((c) => !!c.service);
  allocateByProportion(feeFixedTotal, cardsWithService, (c) => c.totalBill, "feeShare");

  // gán lại cho list đầy đủ
  const feeShareById = new Map(cardsWithService.map((c) => [c.cardId, c.feeShare]));
  cards.forEach((c) => (c.feeShare = feeShareById.get(c.cardId) || 0));

  // 6) returnFee theo % (chỉ áp dụng cho DAO_RUT)
  cards.forEach((c) => {
    if (c.service === "DAO_RUT") c.returnFee = Math.round((c.totalBill * returnFeePercent) / 100);
    else c.returnFee = 0;
  });

  // 7) actualFeeReceived chia cho các thẻ RUT theo tỉ lệ tiền làm
  const actualFeeReceived = getActualFeeReceived();
  const withdrawCards = cards.filter((c) => c.service === "RUT");
  allocateByProportion(actualFeeReceived, withdrawCards, (c) => c.totalBill, "actualShare");

  // 8) auto tiền chuyển cho RUT: tiền làm - tiền phí + phí thực thu
  cards.forEach((c) => {
    if (!c.transferEl) return;

    if (c.service === "RUT") {
      const autoTransfer = Math.max(0, c.totalBill - c.feeShare + c.actualShare);
      lockTransferInput(c.transferEl, true);

      const cur = parseCurrencyVND(c.transferEl.value || "");
      if (cur !== autoTransfer) c.transferEl.value = autoTransfer === 0 ? "" : formatVND(autoTransfer);
      c.transfer = autoTransfer;
    } else {
      lockTransferInput(c.transferEl, false);
      c.transfer = parseCurrencyVND(c.transferEl.value || "");
    }
  });

  // 9) cập nhật UI từng thẻ + total transfer
  let totalTransferAll = 0;
  cards.forEach((c) => {
    updateCardMetricsUI(c.cardId, c.totalBill, c.transfer);
    totalTransferAll += c.transfer;
  });

  // 10) thẻ âm auto theo tổng bill - tổng chuyển
  const totalDiffAll = totalBillAll - totalTransferAll;
  const negativeCardValue = Math.max(0, -totalDiffAll);
  const negativeCardEl = document.getElementById("negativeCardFee");
  if (negativeCardEl) {
    negativeCardEl.readOnly = true;
    negativeCardEl.tabIndex = -1;
    negativeCardEl.value = negativeCardValue === 0 ? "0" : formatVND(negativeCardValue);
  }

  // 11) tính tổng thu/ trả theo dịch vụ (dùng phí_cứng_share thay %)
  const shipFee = getShipFee();
  const servicesSet = new Set(cards.filter((c) => !!c.service).map((c) => c.service));

  let collectNet = 0;
  let payNet = 0;

  if (servicesSet.size > 1) {
    // Mixed: tổng chuyển - tổng làm + tổng phí (feeShare + returnFee)
    const totalFeeAll = cards.reduce((s, c) => s + (c.service ? (c.feeShare + c.returnFee) : 0), 0);
    const net = totalTransferAll - totalBillAll + totalFeeAll;
    if (net >= 0) collectNet = net;
    else payNet = Math.abs(net);
  } else {
    const onlyService = servicesSet.size === 1 ? Array.from(servicesSet)[0] : "";

    if (onlyService === "DAO") {
      cards.forEach((c) => {
        if (c.service !== "DAO") return;
        const r = calcDAO(c);
        collectNet += r.collect;
        payNet += r.pay;
      });
    } else if (onlyService === "RUT") {
      // thu = tổng phí cứng; trả = tổng tiền chuyển (auto)
      collectNet = feeFixedTotal;
      payNet = totalTransferAll;
    } else if (onlyService === "DAO_RUT") {
      cards.forEach((c) => {
        if (c.service !== "DAO_RUT") return;
        const r = calcDAO_RUT(c);
        collectNet += r.collect;
        payNet += r.pay;
      });
    } else {
      collectNet = 0;
      payNet = 0;
    }
  }

  // Tổng thu khách = net thu + ship + thẻ âm
  const totalCollect = collectNet + shipFee + negativeCardValue;
  const totalPay = payNet;

  // 12) UI tổng
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

function toggleServiceDetails(cardId) {
  const service = document.getElementById(`serviceType_${cardId}`)?.value || "";
  const container = document.getElementById(`serviceDetails_${cardId}`);
  if (container) {
    if (service) container.classList.remove("hidden");
    else container.classList.add("hidden");
  }
  recalcSummary();
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
             value="" inputmode="numeric" placeholder="Số tiền" required />

      <select class="bill-pos card-input pos-select" name="billPOS_${cardId}_${billIndex}" required>
        <option value="">-- Chọn POS --</option>
      </select>

      <select class="bill-hkd card-input hkd-select" name="billHKD_${cardId}_${billIndex}" required>
        <option value="">-- Chọn HKD --</option>
      </select>

      <select class="bill-machine card-input machine-select" name="billMachine_${cardId}_${billIndex}" required>
        <option value="">-- Chọn Máy POS --</option>
      </select>

     <input type="text" class="bill-batch card-input integer-only"
       name="billBatch_${cardId}_${billIndex}" placeholder="Số lô"
       inputmode="numeric" pattern="\\d*" />

<input type="text" class="bill-invoice card-input integer-only"
       name="billInvoice_${cardId}_${billIndex}" placeholder="Số hóa đơn"
       inputmode="numeric" pattern="\\d*" />

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
  if (row) initBillRow(row);

  updateBillLimitUI(cardId);
  recalcSummary();
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

  cards.forEach((cardEl, i) => {
    const oldId = cardEl.dataset.cardId;
    updateCardIdentifiers(cardEl, oldId, `TMP${i + 1}`);
  });

  const cards2 = getCardEls();
  cards2.forEach((cardEl, i) => {
    const tmpId = cardEl.dataset.cardId;
    updateCardIdentifiers(cardEl, tmpId, String(i + 1));
  });

  getAllCardIds().forEach((cardId) => {
    const wrapper = document.getElementById(`billDetails_${cardId}`);
    wrapper?.querySelectorAll(".bill-row").forEach(initBillRow);
    renumberBills(cardId);
    toggleServiceDetails(cardId);
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

    if (inp.classList.contains("bill-amount")) {
      inp.value = "";
      return;
    }
    if (inp.id && inp.id.startsWith("transferAmount_")) {
      inp.value = "";
      return;
    }
    if (inp.classList.contains("currency-input")) {
      inp.value = inp.id === "negativeCardFee" ? "0" : "";
      return;
    }
    inp.value = "";
  });

  const serviceSel = cardEl.querySelector(".service-selector");
  if (serviceSel) serviceSel.value = "";

  cardEl.querySelector(".service-details-container")?.classList.add("hidden");

  const wrapper = cardEl.querySelector(`#billDetails_${cardId}`);
  if (wrapper) {
    wrapper.innerHTML = billRowMarkup(cardId, 1);
    wrapper.querySelectorAll(".bill-row").forEach(initBillRow);
  }

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
   GOOGLE SHEET PAYLOAD
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
    feePercentAll: String(document.getElementById("feePercentAll")?.value || "").trim(), // có thể rỗng
    feeFixedAll: parseCurrencyVND(document.getElementById("feeFixedAll")?.value || ""), // thêm để lưu nếu cần
    returnFeePercentAll: String(document.getElementById("returnFeePercentAll")?.value || "").trim(),
    shipFee: parseCurrencyVND(document.getElementById("shipFee")?.value || ""),
    negativeCardFee: parseCurrencyVND(document.getElementById("negativeCardFee")?.value || "0"),

    totalFeeCollectedAll: parseCurrencyVND(document.getElementById("totalFeeCollectedAll")?.textContent || 0),
    actualFeeReceived: parseCurrencyVND(document.getElementById("actualFeeReceived")?.value || ""),

    feePaymentStatus,
    feePaymentStatusText: feeStatusToText(feePaymentStatus),
  };

  const cards = getCardEls().map((cardEl) => {
    const cardId = cardEl.dataset.cardId;

    const bills = Array.from(cardEl.querySelectorAll(".bill-row")).map((row, idx) => ({
      isFirstOfCard: idx === 0, // ✅ dùng cho Apps Script để chỉ hiện 1 lần (yêu cầu #5)
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

      transferAmount: parseCurrencyVND(document.getElementById(`transferAmount_${cardId}`)?.value || ""),
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
   VALIDATION: SỐ NGUYÊN (YÊU CẦU #6)
========================================================= */
function validateDigitsFieldsOrThrow() {
  // 4 số đuôi thẻ
  for (const cardId of getAllCardIds()) {
    const el = document.getElementById(`cardNumber_${cardId}`);
    const v = String(el?.value || "").trim();
    if (!/^\d{4}$/.test(v)) {
      el?.focus();
      throw new Error("4 số đuôi thẻ phải là số nguyên đúng 4 chữ số.");
    }

    // số lô / số hóa đơn: nếu có nhập thì phải digits
    const rows = document.querySelectorAll(`#billDetails_${cardId} .bill-row`);
    for (const row of rows) {
      const batchEl = row.querySelector(".bill-batch");
      const invEl = row.querySelector(".bill-invoice");

      const batch = String(batchEl?.value || "").trim();
      const inv = String(invEl?.value || "").trim();

      if (batch && !/^\d+$/.test(batch)) {
        batchEl?.focus();
        throw new Error("Số lô phải là số nguyên (chỉ gồm chữ số).");
      }
      if (inv && !/^\d+$/.test(inv)) {
        invEl?.focus();
        throw new Error("Số hóa đơn phải là số nguyên (chỉ gồm chữ số).");
      }
    }
  }
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

  // ✅ đảm bảo UI phí cứng tồn tại
  ensureFeeFixedUI();

  // ✅ fee% và ship để trống (không 0)
  const feePercentEl = document.getElementById("feePercentAll");
  if (feePercentEl && String(feePercentEl.value) === "0") feePercentEl.value = "";

  const shipEl = document.getElementById("shipFee");
  if (shipEl && parseCurrencyVND(shipEl.value) === 0) shipEl.value = "";

  // phí thực thu để trống ok
  const actualEl = document.getElementById("actualFeeReceived");
  if (actualEl && parseCurrencyVND(actualEl.value) === 0) actualEl.value = "";

  setupOfficeStaffLogic();
  updateReturnFeePercentAuto();

  // init existing bill rows
  document.querySelectorAll(".bill-row").forEach(initBillRow);

  // format currency inputs (allow empty)
  document.querySelectorAll(".currency-input").forEach((el) => {
    formatCurrencyInput(el, currencyAllowEmpty(el));
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

  showFeeFixedIfNeeded();
});

document.addEventListener("click", (e) => {
  if (e.target.id === "addCardBtn") {
    addNewCard();
    return;
  }

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

  if (e.target.classList.contains("add-bill-btn")) {
    addBillRow(e.target.dataset.cardId);
    return;
  }

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
    recalcSummary();
  }
});

document.addEventListener("change", (e) => {
  if (e.target.classList.contains("service-selector")) {
    const cardId = e.target.closest(".card-item")?.dataset.cardId;
    if (cardId) toggleServiceDetails(cardId);
  }

  const row = e.target.closest(".bill-row");
  if (row) {
    if (e.target.classList.contains("pos-select")) onPosChange(row);
    if (e.target.classList.contains("hkd-select")) onHKDChange(row);
  }

  if (e.target.id === "staff") {
    updateReturnFeePercentAuto();
    recalcSummary();
  }
});

document.addEventListener("input", (e) => {
  // ✅ chặn ký tự chữ cho 4 số đuôi / số lô / số hóa đơn
  if (String(e.target?.id || "").startsWith("cardNumber_")) {
    const cleaned = digitsOnly(e.target.value).slice(0, 4);
    if (e.target.value !== cleaned) e.target.value = cleaned;
  }
  if (e.target.classList?.contains("bill-invoice") || e.target.classList?.contains("bill-batch")) {
    const cleaned = digitsOnly(e.target.value);
    if (e.target.value !== cleaned) e.target.value = cleaned;
  }

  // currency format (allow empty)
  if (e.target.classList.contains("currency-input")) {
    formatCurrencyInput(e.target, currencyAllowEmpty(e.target));
  }

  // fee% -> show & sync phí cứng + auto phí thu về nếu khách văn phòng
  if (e.target.id === "feePercentAll") {
    showFeeFixedIfNeeded();
    updateReturnFeePercentAuto();
    recalcSummary();
    return;
  }

  if (e.target.id === "feeFixedAll") {
    showFeeFixedIfNeeded();
    recalcSummary();
    return;
  }

  // ship/actual/transfer/bill => recalc
  if (
    e.target.id === "shipFee" ||
    e.target.id === "actualFeeReceived" ||
    isTransferEl(e.target) ||
    e.target.classList.contains("bill-amount")
  ) {
    recalcSummary();
  }
});

/* =========================================================
   SUBMIT
========================================================= */
document.getElementById("mainForm")?.addEventListener("submit", async (e) => {
  e.preventDefault();

  if (IS_SUBMITTING) return;
  IS_SUBMITTING = true;

  const submitBtn = document.querySelector('button[type="submit"]');
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.textContent = "Đang gửi...";
  }

  try {
    // sync staffFinal
    const staffFinal = document.getElementById("staffFinal");
    const staffEl = document.getElementById("staff");
    const contactEl = document.getElementById("customerDetail");

    if (staffFinal) {
      staffFinal.value =
        staffEl?.value === "Khách văn phòng" ? (contactEl?.value || "").trim() : (staffEl?.value || "");
    }

    if (staffEl?.value === "Khách văn phòng" && !(contactEl?.value || "").trim()) {
      alert("Vui lòng nhập Liên hệ cho Khách văn phòng.");
      contactEl?.focus();
      throw new Error("missing contact");
    }

    // validate bill amount > 0
    for (const cardId of getAllCardIds()) {
      const rows = document.querySelectorAll(`#billDetails_${cardId} .bill-row`);
      for (const row of rows) {
        const amountEl = row.querySelector(".bill-amount");
        const v = parseCurrencyVND(amountEl?.value || "");
        if (!v || v <= 0) {
          alert("Vui lòng nhập Số tiền BILL (không để trống/0).");
          amountEl?.focus();
          throw new Error("missing bill amount");
        }
      }
    }

    // validate số nguyên: 4 số thẻ / số lô / số hóa đơn
    validateDigitsFieldsOrThrow();

    // đảm bảo tính toán mới nhất
    recalcSummary();

    const payload = collectPayloadForSheet();
    await sendToGoogleSheet(payload);

    alert("Đã gửi đơn thành công!");
    setTimeout(() => location.reload(), 300);
  } catch (err) {
    console.error(err);
    if (String(err?.message || "") !== "missing contact" && String(err?.message || "") !== "missing bill amount") {
      alert(err?.message || "Gửi đơn thất bại.");
    }

    IS_SUBMITTING = false;
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = "Gửi Đơn";
    }
  }
});
