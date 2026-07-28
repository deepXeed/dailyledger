/* ============================================================
   Daily Ledger — app.js
   Firebase-backed ledger (Realtime Database + Auth) with Excel export (SheetJS)
   ============================================================ */

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const rtdb = firebase.database();

let currentUser = null; // set by onAuthStateChanged
let pendingSignupUsername = null; // avoids a display-name race right after signup

/* ---------------- Auth helpers ---------------- */
// Firebase Auth wants a real email; we turn "Username" into a fake one so the
// UI can stay simple (just Username + Secret Code) while still getting real
// per-account security from Firebase.
function usernameToEmail(username) {
  const clean = String(username || "").trim().toLowerCase().replace(/[^a-z0-9._-]/g, "");
  return clean ? `${clean}@daily-ledger.local` : null;
}

function friendlyAuthError(err) {
  const code = err && err.code;
  switch (code) {
    case "auth/invalid-email":
    case "auth/missing-email":
      return "Please enter a username.";
    case "auth/user-not-found":
      return "No account with that username. Tap \"Create New Account\" if you're new.";
    case "auth/wrong-password":
    case "auth/invalid-credential":
    case "auth/invalid-login-credentials":
      return "Wrong secret code for that username.";
    case "auth/email-already-in-use":
      return "That username is already taken — try logging in instead.";
    case "auth/weak-password":
      return "Secret code must be at least 6 characters.";
    default:
      return (err && err.message) || "Something went wrong. Please try again.";
  }
}

async function logIn(username, secretCode) {
  const email = usernameToEmail(username);
  if (!email) throw { code: "auth/missing-email" };
  await auth.signInWithEmailAndPassword(email, secretCode);
}
async function signUp(username, secretCode) {
  const email = usernameToEmail(username);
  if (!email) throw { code: "auth/missing-email" };
  pendingSignupUsername = username.trim();
  await auth.createUserWithEmailAndPassword(email, secretCode);
  // Store the human-readable username alongside the account for display purposes.
  await rtdb.ref(`users/${auth.currentUser.uid}/profile`).set({ username: username.trim() });
}
async function logOut() {
  await auth.signOut();
}

function userRef(path) {
  if (!currentUser) throw new Error("Not logged in");
  return rtdb.ref(`users/${currentUser.uid}/${path}`);
}

/* ---------------- Settings ---------------- */
async function getSettings() {
  const snap = await userRef("settings").once("value");
  const val = snap.val() || {};
  let startDate = val.startDate || null;
  const openingBalance = val.openingBalance || 0;
  const currencyCode = val.currencyCode || "AED";
  if (!startDate) {
    startDate = isoDate(new Date());
    await saveSettings({ startDate, openingBalance, currencyCode });
  }
  return { startDate, openingBalance, currencyCode };
}
async function saveSettings({ startDate, openingBalance, currencyCode }) {
  await userRef("settings").set({
    startDate,
    openingBalance: Number(openingBalance) || 0,
    currencyCode: currencyCode || "AED",
  });
}

const SUPPORTED_CURRENCIES = ["AED", "USD", "EUR", "GBP", "INR", "PKR", "SAR", "QAR", "KWD", "OMR", "BHD", "EGP", "PHP", "NGN"];
let CURRENT_CURRENCY = "AED";

/* ---------------- Collections CRUD ---------------- */
async function addCollection(date, name, amount) {
  await userRef("collections").push({ date, name, amount: Number(amount) });
}
async function deleteCollection(id) {
  await userRef(`collections/${id}`).remove();
}
function snapToArray(snapVal) {
  if (!snapVal) return [];
  return Object.keys(snapVal).map((id) => ({ id, ...snapVal[id] }));
}
async function getCollectionsForDate(date) {
  const snap = await userRef("collections").orderByChild("date").equalTo(date).once("value");
  return snapToArray(snap.val());
}
async function getAllCollections() {
  const snap = await userRef("collections").once("value");
  return snapToArray(snap.val());
}

/* ---------------- Petty Fund expense log CRUD ---------------- */
async function addExpense(date, place, amount) {
  await userRef("expenses").push({ date, place, amount: Number(amount) });
}
async function deleteExpense(id) {
  await userRef(`expenses/${id}`).remove();
}
async function getExpensesForDate(date) {
  const snap = await userRef("expenses").orderByChild("date").equalTo(date).once("value");
  return snapToArray(snap.val());
}
async function getAllExpenses() {
  const snap = await userRef("expenses").once("value");
  return snapToArray(snap.val());
}

/* ---------------- Top-ups (one value per date) ---------------- */
async function setTopup(date, amount) {
  const amt = Number(amount) || 0;
  if (amt === 0) {
    await userRef(`topups/${date}`).remove();
  } else {
    await userRef(`topups/${date}`).set(amt);
  }
}
async function getTopup(date) {
  const snap = await userRef(`topups/${date}`).once("value");
  return snap.val() || 0;
}
async function getAllTopups() {
  const snap = await userRef("topups").once("value");
  const val = snap.val() || {};
  return Object.keys(val).map((date) => ({ date, amount: val[date] }));
}

/* ---------------- Day options: sweep-to-net + deposit tracking ---------------- */
const DEFAULT_DAY_OPTIONS = { addClosingToNet: false, deposited: false, depositFull: true, depositCustomAmount: 0 };

async function getDayOptions(date) {
  const snap = await userRef(`dayOptions/${date}`).once("value");
  const rec = snap.val();
  return rec ? { ...DEFAULT_DAY_OPTIONS, ...rec } : { ...DEFAULT_DAY_OPTIONS, date };
}
async function setDayOptions(date, patch) {
  await userRef(`dayOptions/${date}`).update({ ...patch });
}
async function getAllDayOptions() {
  const snap = await userRef("dayOptions").once("value");
  const val = snap.val() || {};
  return Object.keys(val).map((date) => ({ date, ...val[date] }));
}

/* ---------------- Date helpers ---------------- */
function pad2(n) { return String(n).padStart(2, "0"); }
function isoDate(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function parseISO(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}
function addDaysISO(s, n) {
  const d = parseISO(s);
  d.setDate(d.getDate() + n);
  return isoDate(d);
}
function daysInMonth(year, month1based) { return new Date(year, month1based, 0).getDate(); }
const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
function fmtMoney(n) {
  const v = Number(n) || 0;
  return CURRENT_CURRENCY + " " + v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtDisplayDate(s) {
  const d = parseISO(s);
  return d.toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short", year: "numeric" });
}

/* ---------------- Ledger chain computation ---------------- */
// Computes the running Petty Fund chain from settings.startDate through endDateStr (inclusive).
// Uses cached raw data when available so deposit toggles stay snappy.
function computeChainFromData(endDateStr, data) {
  const settings = data.settings;
  const map = new Map();
  if (!settings || endDateStr < settings.startDate) return map;

  const collByDate = {};
  for (const c of data.collections) collByDate[c.date] = (collByDate[c.date] || 0) + c.amount;
  const expByDate = {};
  for (const e of data.expenses) expByDate[e.date] = (expByDate[e.date] || 0) + e.amount;
  const topupByDate = {};
  for (const t of data.topups) topupByDate[t.date] = t.amount;
  const optsByDate = {};
  for (const o of data.dayOptions) optsByDate[o.date] = o;

  let balance = Number(settings.openingBalance) || 0;
  let d = settings.startDate;
  while (d <= endDateStr) {
    const carry = balance;
    const topup = topupByDate[d] || 0;
    const open = carry + topup;
    const exp = expByDate[d] || 0;
    const shortfall = Math.max(exp - open, 0);
    const pettyClosing = Math.max(open - exp, 0);
    const coll = collByDate[d] || 0;
    const netCollection = coll - shortfall;

    const opts = { ...DEFAULT_DAY_OPTIONS, ...(optsByDate[d] || {}) };
    const addClosingToNet = !!opts.addClosingToNet;
    const netAmount = netCollection + (addClosingToNet ? pettyClosing : 0);

    let depositedAmount, remaining;
    if (!opts.deposited) {
      depositedAmount = 0;
      remaining = netAmount;
    } else if (opts.depositFull) {
      depositedAmount = netAmount;
      remaining = 0;
    } else {
      const custom = Math.min(Math.max(Number(opts.depositCustomAmount) || 0, 0), netAmount);
      depositedAmount = custom;
      remaining = netAmount - custom;
    }

    const pettyRetained = addClosingToNet ? 0 : pettyClosing;
    const nextCarry = pettyRetained + remaining;

    map.set(d, {
      date: d, carry, topup, open, exp, shortfall, pettyClosing, coll, netCollection,
      addClosingToNet, netAmount,
      deposited: !!opts.deposited, depositFull: !!opts.depositFull,
      depositCustomAmount: Number(opts.depositCustomAmount) || 0,
      depositedAmount, remaining, pettyRetained, nextCarry,
    });
    balance = nextCarry;
    d = addDaysISO(d, 1);
  }
  return map;
}

async function computeChain(endDateStr) {
  const data = await loadLedgerData();
  return computeChainFromData(endDateStr, data);
}

/* ---------------- App state ---------------- */
const state = {
  currentDate: isoDate(new Date()),
  activeView: "today",
};

/* Raw-data cache so deposit toggles don't re-fetch the whole ledger */
const dataCache = {
  collections: null,
  expenses: null,
  topups: null,
  dayOptions: null,
  settings: null,
  dirty: true,
};

function invalidateCache() {
  dataCache.dirty = true;
  dataCache.collections = null;
  dataCache.expenses = null;
  dataCache.topups = null;
  dataCache.dayOptions = null;
  dataCache.settings = null;
}

async function loadLedgerData(force = false) {
  if (!force && !dataCache.dirty && dataCache.collections) {
    return dataCache;
  }
  const [collections, expenses, topups, dayOptions, settings] = await Promise.all([
    getAllCollections(),
    getAllExpenses(),
    getAllTopups(),
    getAllDayOptions(),
    getSettings(),
  ]);
  dataCache.collections = collections;
  dataCache.expenses = expenses;
  dataCache.topups = topups;
  dataCache.dayOptions = dayOptions;
  dataCache.settings = settings;
  dataCache.dirty = false;
  return dataCache;
}

/* ---------------- Toast ---------------- */
let toastTimer = null;
function showToast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 2200);
}

/* ---------------- Render: Today view ---------------- */
async function renderToday() {
  document.getElementById("todayDateInput").value = state.currentDate;

  const [colls, exps, topup, settings] = await Promise.all([
    getCollectionsForDate(state.currentDate),
    getExpensesForDate(state.currentDate),
    getTopup(state.currentDate),
    getSettings(),
  ]);

  document.getElementById("topupInput").value = topup ? topup : "";

  const collRows = document.getElementById("collectionRows");
  const collTotal = colls.reduce((s, c) => s + c.amount, 0);
  document.getElementById("collectionTotal").textContent = fmtMoney(collTotal);
  collRows.innerHTML = "";
  if (colls.length === 0) {
    collRows.innerHTML = `<div class="empty-hint">No collections logged for this day yet.</div>`;
  } else {
    colls.sort((a, b) => (a.id < b.id ? -1 : 1)).forEach((c) => {
      const row = document.createElement("div");
      row.className = "row-item";
      row.innerHTML = `<span class="name">${escapeHtml(c.name)}</span><span class="amt">${fmtMoney(c.amount)}</span><button class="del" data-store="collections" data-id="${c.id}" aria-label="Delete">✕</button>`;
      collRows.appendChild(row);
    });
  }

  const expRows = document.getElementById("expenseRows");
  const expTotal = exps.reduce((s, e) => s + e.amount, 0);
  document.getElementById("expenseTotal").textContent = fmtMoney(expTotal);
  expRows.innerHTML = "";
  if (exps.length === 0) {
    expRows.innerHTML = `<div class="empty-hint">No petty fund expenses logged for this day yet.</div>`;
  } else {
    exps.sort((a, b) => (a.id < b.id ? -1 : 1)).forEach((e) => {
      const row = document.createElement("div");
      row.className = "row-item";
      row.innerHTML = `<span class="name">${escapeHtml(e.place)}</span><span class="amt">${fmtMoney(e.amount)}</span><button class="del" data-store="expenses" data-id="${e.id}" aria-label="Delete">✕</button>`;
      expRows.appendChild(row);
    });
  }

  const tape = document.getElementById("daySummaryTape");
  const depositCard = document.getElementById("depositCard");
  if (state.currentDate < settings.startDate) {
    tape.innerHTML = `<div class="empty-hint">This date is before your Ledger Start Date (set in Settings), so it isn't included in the Petty Fund chain yet.</div>`;
    depositCard.style.display = "none";
    return;
  }
  depositCard.style.display = "";
  // Warm the full-ledger cache, then paint from it (avoids a second full fetch on toggle)
  await loadLedgerData();
  const chain = computeChainFromData(state.currentDate, dataCache);
  const row = chain.get(state.currentDate);
  if (!row) return;
  paintDaySummaryTape(row);
  renderDepositCard(row);
}

function paintDaySummaryTape(row) {
  const tape = document.getElementById("daySummaryTape");
  tape.innerHTML = `
    <div class="tape-row"><span class="label">Carried Forward from Previous Day</span><span class="val">${fmtMoney(row.carry)}</span></div>
    <div class="tape-row"><span class="label">Additional Funds Added</span><span class="val">${fmtMoney(row.topup)}</span></div>
    <div class="tape-row total"><span class="label">Opening Balance (Total)</span><span class="val">${fmtMoney(row.open)}</span></div>
    <div class="tape-row"><span class="label">Total Petty Fund Expense</span><span class="val">${fmtMoney(row.exp)}</span></div>
    <div class="tape-row"><span class="label">Shortfall Taken from Collection</span><span class="val ${row.shortfall > 0 ? "warn" : ""}">${fmtMoney(row.shortfall)}</span></div>
    <div class="tape-row total"><span class="label">Petty Fund Closing Balance</span><span class="val">${fmtMoney(row.pettyClosing)}</span></div>
    <div class="tape-row"><span class="label">Total Customer Collection</span><span class="val">${fmtMoney(row.coll)}</span></div>
    <div class="tape-row total"><span class="label">Net Collection After Petty Deduction</span><span class="val ${row.shortfall > 0 ? "warn" : "good"}">${fmtMoney(row.netCollection)}</span></div>
  `;
}

function renderDepositCard(row) {
  const body = document.getElementById("depositCardBody");
  // Amounts first, Yes/No questions at the end (deposited / not deposited sit just above them)
  let html = `
    <div class="deposit-amounts">
      <div class="tape-row total"><span class="label">Net Amount (available for deposit)</span><span class="val">${fmtMoney(row.netAmount)}</span></div>
      <div class="tape-row total highlight"><span class="label">Deposited Amount</span><span class="val good">${fmtMoney(row.depositedAmount)}</span></div>
      <div class="tape-row total"><span class="label">Not Deposited — Carried to Petty Fund</span><span class="val ${row.remaining > 0 ? "warn" : ""}">${fmtMoney(row.remaining)}</span></div>
    </div>
    <div class="deposit-questions">
      <div class="toggle-group">
        <label>Add Petty Fund Closing Balance (${fmtMoney(row.pettyClosing)}) to Net Collection?</label>
        <div class="toggle-btns">
          <button type="button" class="toggle-btn ${row.addClosingToNet ? "active" : ""}" data-toggle="addClosingToNet" data-value="true">Yes</button>
          <button type="button" class="toggle-btn ${!row.addClosingToNet ? "active" : ""}" data-toggle="addClosingToNet" data-value="false">No</button>
        </div>
      </div>
      <div class="toggle-group">
        <label>Is today's Net Amount deposited?</label>
        <div class="toggle-btns">
          <button type="button" class="toggle-btn ${row.deposited ? "active" : ""}" data-toggle="deposited" data-value="true">Yes</button>
          <button type="button" class="toggle-btn ${!row.deposited ? "active" : ""}" data-toggle="deposited" data-value="false">No</button>
        </div>
      </div>
  `;

  if (row.deposited) {
    html += `
      <div class="toggle-group">
        <label>Deposit the full amount?</label>
        <div class="toggle-btns">
          <button type="button" class="toggle-btn ${row.depositFull ? "active" : ""}" data-toggle="depositFull" data-value="true">Yes, full</button>
          <button type="button" class="toggle-btn ${!row.depositFull ? "active" : ""}" data-toggle="depositFull" data-value="false">No, custom</button>
        </div>
      </div>`;
    if (!row.depositFull) {
      html += `
      <div class="topup-input-row deposit-custom">
        <label for="customDepositInput">Custom deposit amount</label>
        <input type="number" id="customDepositInput" step="0.01" min="0" max="${row.netAmount}" value="${row.depositCustomAmount || row.depositedAmount || ""}" placeholder="0.00" />
      </div>`;
    }
  }

  html += `</div>`;
  body.innerHTML = html;
  wireDepositCardEvents();
}

function wireDepositCardEvents() {
  const body = document.getElementById("depositCardBody");
  body.querySelectorAll(".toggle-btn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      if (btn.disabled) return;
      const field = btn.dataset.toggle;
      const value = btn.dataset.value === "true";

      // Optimistic active state for instant feedback
      const group = btn.closest(".toggle-btns");
      if (group) {
        group.querySelectorAll(".toggle-btn").forEach((b) => b.classList.toggle("active", b === btn));
      }
      body.querySelectorAll(".toggle-btn").forEach((b) => { b.disabled = true; });

      try {
        await setDayOptions(state.currentDate, { [field]: value });
        // Patch local cache so we don't re-fetch everything
        patchDayOptionsCache(state.currentDate, { [field]: value });
        await refreshSummaryAndDeposit();
      } catch (err) {
        console.error(err);
        showToast("Could not save — check connection");
        await refreshSummaryAndDeposit();
      }
    });
  });

  const customInput = document.getElementById("customDepositInput");
  if (customInput) {
    let debounceTimer = null;
    const saveCustom = async () => {
      const amount = Number(customInput.value) || 0;
      try {
        await setDayOptions(state.currentDate, { depositCustomAmount: amount });
        patchDayOptionsCache(state.currentDate, { depositCustomAmount: amount });
        await refreshSummaryAndDeposit();
      } catch (err) {
        console.error(err);
        showToast("Could not save — check connection");
      }
    };
    customInput.addEventListener("change", saveCustom);
    customInput.addEventListener("input", () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(saveCustom, 350);
    });
  }
}

function patchDayOptionsCache(date, patch) {
  if (!dataCache.dayOptions) return;
  const idx = dataCache.dayOptions.findIndex((o) => o.date === date);
  if (idx >= 0) {
    dataCache.dayOptions[idx] = { ...dataCache.dayOptions[idx], ...patch };
  } else {
    dataCache.dayOptions.push({ date, ...DEFAULT_DAY_OPTIONS, ...patch });
  }
}

/** Fast path: recompute chain from cache and repaint summary + deposit only */
async function refreshSummaryAndDeposit() {
  const settings = dataCache.settings || (await getSettings());
  const depositCard = document.getElementById("depositCard");
  const tape = document.getElementById("daySummaryTape");

  if (state.currentDate < settings.startDate) {
    tape.innerHTML = `<div class="empty-hint">This date is before your Ledger Start Date (set in Settings), so it isn't included in the Petty Fund chain yet.</div>`;
    depositCard.style.display = "none";
    return;
  }
  depositCard.style.display = "";

  const data = await loadLedgerData();
  const chain = computeChainFromData(state.currentDate, data);
  const row = chain.get(state.currentDate);
  if (!row) return;
  paintDaySummaryTape(row);
  renderDepositCard(row);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/* ---------------- Render: Month view ---------------- */
function populateMonthYearSelectors() {
  const monthSelect = document.getElementById("monthSelect");
  const yearSelect = document.getElementById("monthYearSelect");
  if (monthSelect.options.length === 0) {
    MONTH_NAMES.forEach((m, i) => {
      const opt = document.createElement("option");
      opt.value = i + 1;
      opt.textContent = m;
      monthSelect.appendChild(opt);
    });
  }
  if (yearSelect.options.length === 0) {
    const thisYear = new Date().getFullYear();
    for (let y = thisYear - 3; y <= thisYear + 3; y++) {
      const opt = document.createElement("option");
      opt.value = y;
      opt.textContent = y;
      yearSelect.appendChild(opt);
    }
    yearSelect.value = thisYear;
  }
  const now = new Date();
  if (!monthSelect.dataset.init) {
    monthSelect.value = now.getMonth() + 1;
    monthSelect.dataset.init = "1";
  }
}

async function renderMonth() {
  populateMonthYearSelectors();
  const month = Number(document.getElementById("monthSelect").value);
  const year = Number(document.getElementById("monthYearSelect").value);
  const dim = daysInMonth(year, month);
  const lastDateStr = `${year}-${pad2(month)}-${pad2(dim)}`;
  const chain = await computeChain(lastDateStr);

  let html = `<thead><tr><th>Date</th><th>Collection</th><th>Petty Exp.</th><th>Shortfall</th><th>Net Coll.</th><th>Net Amount</th><th>Deposited</th><th>Carried Fwd</th></tr></thead><tbody>`;
  let totColl = 0, totExp = 0, totShort = 0, totNet = 0, totAmount = 0, totDeposited = 0;
  let lastCarry = 0;
  for (let day = 1; day <= dim; day++) {
    const dstr = `${year}-${pad2(month)}-${pad2(day)}`;
    const row = chain.get(dstr);
    if (!row) {
      html += `<tr class="na-row"><td>${pad2(day)} ${MONTH_NAMES[month - 1].slice(0,3)}</td><td colspan="7">Before ledger start date</td></tr>`;
      continue;
    }
    totColl += row.coll; totExp += row.exp; totShort += row.shortfall; totNet += row.netCollection;
    totAmount += row.netAmount; totDeposited += row.depositedAmount; lastCarry = row.nextCarry;
    html += `<tr><td>${pad2(day)} ${MONTH_NAMES[month - 1].slice(0,3)}</td><td>${row.coll.toFixed(2)}</td><td>${row.exp.toFixed(2)}</td><td>${row.shortfall.toFixed(2)}</td><td>${row.netCollection.toFixed(2)}</td><td>${row.netAmount.toFixed(2)}</td><td>${row.depositedAmount.toFixed(2)}</td><td>${row.nextCarry.toFixed(2)}</td></tr>`;
  }
  html += `<tr class="total-row"><td>TOTAL</td><td>${totColl.toFixed(2)}</td><td>${totExp.toFixed(2)}</td><td>${totShort.toFixed(2)}</td><td>${totNet.toFixed(2)}</td><td>${totAmount.toFixed(2)}</td><td>${totDeposited.toFixed(2)}</td><td>${lastCarry.toFixed(2)}</td></tr>`;
  html += `</tbody>`;
  document.getElementById("monthTable").innerHTML = html;
}

/* ---------------- Render: Year view ---------------- */
function populateYearSelector() {
  const yearSelect = document.getElementById("yearSelect");
  if (yearSelect.options.length === 0) {
    const thisYear = new Date().getFullYear();
    for (let y = thisYear - 3; y <= thisYear + 3; y++) {
      const opt = document.createElement("option");
      opt.value = y;
      opt.textContent = y;
      yearSelect.appendChild(opt);
    }
    yearSelect.value = thisYear;
  }
}

async function renderYear() {
  populateYearSelector();
  const year = Number(document.getElementById("yearSelect").value);
  const lastDateStr = `${year}-12-31`;
  const chain = await computeChain(lastDateStr);

  let html = `<thead><tr><th>Month</th><th>Collection</th><th>Petty Exp.</th><th>Shortfall</th><th>Net Coll.</th><th>Net Amount</th><th>Deposited</th><th>Carried Fwd</th></tr></thead><tbody>`;
  let totColl = 0, totExp = 0, totShort = 0, totNet = 0, totAmount = 0, totDeposited = 0;
  let yearCarry = 0;
  for (let month = 1; month <= 12; month++) {
    const dim = daysInMonth(year, month);
    let mColl = 0, mExp = 0, mShort = 0, mNet = 0, mAmount = 0, mDeposited = 0, mCarry = null, any = false;
    for (let day = 1; day <= dim; day++) {
      const dstr = `${year}-${pad2(month)}-${pad2(day)}`;
      const row = chain.get(dstr);
      if (!row) continue;
      any = true;
      mColl += row.coll; mExp += row.exp; mShort += row.shortfall; mNet += row.netCollection;
      mAmount += row.netAmount; mDeposited += row.depositedAmount; mCarry = row.nextCarry;
    }
    if (!any) {
      html += `<tr class="na-row"><td>${MONTH_NAMES[month - 1]}</td><td colspan="7">Before ledger start date</td></tr>`;
      continue;
    }
    totColl += mColl; totExp += mExp; totShort += mShort; totNet += mNet;
    totAmount += mAmount; totDeposited += mDeposited; yearCarry = mCarry;
    html += `<tr><td>${MONTH_NAMES[month - 1]}</td><td>${mColl.toFixed(2)}</td><td>${mExp.toFixed(2)}</td><td>${mShort.toFixed(2)}</td><td>${mNet.toFixed(2)}</td><td>${mAmount.toFixed(2)}</td><td>${mDeposited.toFixed(2)}</td><td>${mCarry.toFixed(2)}</td></tr>`;
  }
  html += `<tr class="total-row"><td>YEAR TOTAL</td><td>${totColl.toFixed(2)}</td><td>${totExp.toFixed(2)}</td><td>${totShort.toFixed(2)}</td><td>${totNet.toFixed(2)}</td><td>${totAmount.toFixed(2)}</td><td>${totDeposited.toFixed(2)}</td><td>${yearCarry.toFixed(2)}</td></tr>`;
  html += `</tbody>`;
  document.getElementById("yearTable").innerHTML = html;
}

/* ---------------- Settings view ---------------- */
async function renderSettings() {
  const settings = await getSettings();
  document.getElementById("startDateInput").value = settings.startDate;
  document.getElementById("openingBalanceInput").value = settings.openingBalance;

  const exportYearSelect = document.getElementById("exportYearSelect");
  if (exportYearSelect.options.length === 0) {
    const thisYear = new Date().getFullYear();
    for (let y = thisYear - 3; y <= thisYear + 3; y++) {
      const opt = document.createElement("option");
      opt.value = y;
      opt.textContent = y;
      exportYearSelect.appendChild(opt);
    }
    exportYearSelect.value = thisYear;
  }

  if (window.matchMedia && window.matchMedia("(display-mode: standalone)").matches) {
    document.getElementById("installHint").style.display = "none";
  }
}

function populateCurrencySelectors(selected) {
  [document.getElementById("headerCurrencySelect"), document.getElementById("settingsCurrencySelect")].forEach((sel) => {
    if (!sel) return;
    if (sel.options.length === 0) {
      SUPPORTED_CURRENCIES.forEach((code) => {
        const opt = document.createElement("option");
        opt.value = code;
        opt.textContent = code;
        sel.appendChild(opt);
      });
    }
    sel.value = selected;
  });
}

async function applyCurrencyChange(newCode) {
  CURRENT_CURRENCY = newCode;
  populateCurrencySelectors(newCode);
  const settings = await getSettings();
  await saveSettings({ ...settings, currencyCode: newCode });
  invalidateCache();
  await switchView(state.activeView);
  showToast("Currency set to " + newCode);
}

/* ---------------- View switching ---------------- */
async function switchView(view) {
  state.activeView = view;
  document.querySelectorAll(".view").forEach((v) => v.classList.remove("active"));
  document.getElementById(`view-${view}`).classList.add("active");
  document.querySelectorAll("nav.bottomnav button").forEach((b) => b.classList.toggle("active", b.dataset.view === view));
  document.getElementById("dateNavToday").style.visibility = view === "today" ? "visible" : "hidden";
  if (view === "today") await renderToday();
  else if (view === "month") await renderMonth();
  else if (view === "year") await renderYear();
  else if (view === "settings") await renderSettings();
}

/* ---------------- Excel export ---------------- */
async function exportYearToExcel(year) {
  const lastDateStr = `${year}-12-31`;
  const chain = await computeChain(lastDateStr);
  const [allColl, allExp, allTopups] = await Promise.all([getAllCollections(), getAllExpenses(), getAllTopups()]);
  const cur = CURRENT_CURRENCY;

  const wb = XLSX.utils.book_new();
  const CUR_FMT = `"${cur}" #,##0.00`;

  function applyCurrencyCols(ws, colLetters, rowCount, headerRow = 1) {
    for (let r = headerRow + 1; r <= headerRow + rowCount; r++) {
      colLetters.forEach((col) => {
        const ref = `${col}${r}`;
        if (ws[ref]) ws[ref].z = CUR_FMT;
      });
    }
  }

  // ---- Yearly Summary sheet ----
  const yearHeader = ["Month", "Total Collection", "Total Petty Fund Expense", "Shortfall from Collection", "Net Collection", "Net Amount (after sweep)", "Deposited Amount", "Carried Forward to Petty Fund"];
  const yearAOA = [yearHeader];
  let yTotColl = 0, yTotExp = 0, yTotShort = 0, yTotNet = 0, yTotAmount = 0, yTotDeposited = 0, yCarry = 0;
  for (let month = 1; month <= 12; month++) {
    const dim = daysInMonth(year, month);
    let mColl = 0, mExp = 0, mShort = 0, mNet = 0, mAmount = 0, mDeposited = 0, mCarry = null, any = false;
    for (let day = 1; day <= dim; day++) {
      const row = chain.get(`${year}-${pad2(month)}-${pad2(day)}`);
      if (!row) continue;
      any = true;
      mColl += row.coll; mExp += row.exp; mShort += row.shortfall; mNet += row.netCollection;
      mAmount += row.netAmount; mDeposited += row.depositedAmount; mCarry = row.nextCarry;
    }
    if (!any) { yearAOA.push([MONTH_NAMES[month - 1], "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]); continue; }
    yTotColl += mColl; yTotExp += mExp; yTotShort += mShort; yTotNet += mNet;
    yTotAmount += mAmount; yTotDeposited += mDeposited; yCarry = mCarry;
    yearAOA.push([MONTH_NAMES[month - 1], mColl, mExp, mShort, mNet, mAmount, mDeposited, mCarry]);
  }
  yearAOA.push(["YEAR TOTAL", yTotColl, yTotExp, yTotShort, yTotNet, yTotAmount, yTotDeposited, yCarry]);
  const yearWs = XLSX.utils.aoa_to_sheet(yearAOA);
  yearWs["!cols"] = [{ wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 20 }];
  applyCurrencyCols(yearWs, ["B", "C", "D", "E", "F", "G", "H"], yearAOA.length - 1);
  XLSX.utils.book_append_sheet(wb, yearWs, "Yearly Summary");

  // ---- One sheet per month ----
  const monthHeader = ["Date", "Carried Forward", "Additional Funds Added", "Opening Balance", "Customer Collection", "Petty Fund Expense", "Shortfall from Collection", "Net Collection", "Added Petty Closing to Net?", "Net Amount", "Deposited?", "Deposited Amount", "Not Deposited (Carried Fwd)"];
  for (let month = 1; month <= 12; month++) {
    const dim = daysInMonth(year, month);
    const aoa = [monthHeader];
    let tColl = 0, tExp = 0, tShort = 0, tNet = 0, tAmount = 0, tDeposited = 0, tCarry = 0;
    for (let day = 1; day <= dim; day++) {
      const dstr = `${year}-${pad2(month)}-${pad2(day)}`;
      const row = chain.get(dstr);
      if (!row) { aoa.push([dstr, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"]); continue; }
      aoa.push([
        dstr, row.carry, row.topup, row.open, row.coll, row.exp, row.shortfall, row.netCollection,
        row.addClosingToNet ? "Yes" : "No", row.netAmount,
        row.deposited ? (row.depositFull ? "Yes (Full)" : "Yes (Custom)") : "No",
        row.depositedAmount, row.nextCarry,
      ]);
      tColl += row.coll; tExp += row.exp; tShort += row.shortfall; tNet += row.netCollection;
      tAmount += row.netAmount; tDeposited += row.depositedAmount; tCarry = row.nextCarry;
    }
    aoa.push(["TOTAL", "", "", "", tColl, tExp, tShort, tNet, "", tAmount, "", tDeposited, tCarry]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
    applyCurrencyCols(ws, ["B", "C", "D", "E", "F", "G", "H", "J", "L", "M"], aoa.length - 1);
    XLSX.utils.book_append_sheet(wb, ws, MONTH_NAMES[month - 1].slice(0, 31));
  }

  // ---- Raw logs (full history, for backup/audit) ----
  const collAOA = [["Date", "Customer Name", `Amount (${cur})`]].concat(
    allColl.sort((a, b) => (a.date < b.date ? -1 : 1)).map((c) => [c.date, c.name, c.amount])
  );
  const collWs = XLSX.utils.aoa_to_sheet(collAOA);
  collWs["!cols"] = [{ wch: 12 }, { wch: 26 }, { wch: 14 }];
  applyCurrencyCols(collWs, ["C"], collAOA.length - 1);
  XLSX.utils.book_append_sheet(wb, collWs, "Collections Log");

  const expAOA = [["Date", "Where Used (Place / Purpose)", `Amount (${cur})`]].concat(
    allExp.sort((a, b) => (a.date < b.date ? -1 : 1)).map((e) => [e.date, e.place, e.amount])
  );
  const expWs = XLSX.utils.aoa_to_sheet(expAOA);
  expWs["!cols"] = [{ wch: 12 }, { wch: 30 }, { wch: 14 }];
  applyCurrencyCols(expWs, ["C"], expAOA.length - 1);
  XLSX.utils.book_append_sheet(wb, expWs, "Petty Fund Expense Log");

  const topupAOA = [["Date", `Additional Funds Added (${cur})`]].concat(
    allTopups.sort((a, b) => (a.date < b.date ? -1 : 1)).map((t) => [t.date, t.amount])
  );
  const topupWs = XLSX.utils.aoa_to_sheet(topupAOA);
  topupWs["!cols"] = [{ wch: 12 }, { wch: 22 }];
  applyCurrencyCols(topupWs, ["B"], topupAOA.length - 1);
  XLSX.utils.book_append_sheet(wb, topupWs, "Fund Top-ups Log");

  XLSX.writeFile(wb, `Daily_Ledger_${cur}_${year}.xlsx`);
}

/* ---------------- Event wiring ---------------- */
function wireEvents() {
  document.querySelectorAll("nav.bottomnav button").forEach((btn) => {
    btn.addEventListener("click", () => switchView(btn.dataset.view));
  });

  document.getElementById("todayDateInput").addEventListener("change", async (e) => {
    state.currentDate = e.target.value;
    await renderToday();
  });
  document.getElementById("prevDayBtn").addEventListener("click", async () => {
    state.currentDate = addDaysISO(state.currentDate, -1);
    await renderToday();
  });
  document.getElementById("nextDayBtn").addEventListener("click", async () => {
    state.currentDate = addDaysISO(state.currentDate, 1);
    await renderToday();
  });
  document.getElementById("jumpTodayBtn").addEventListener("click", async () => {
    state.currentDate = isoDate(new Date());
    await renderToday();
  });

  document.getElementById("collectionForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const name = document.getElementById("collectionName").value.trim();
    const amount = document.getElementById("collectionAmount").value;
    if (!name || !amount) return;
    await addCollection(state.currentDate, name, amount);
    invalidateCache();
    e.target.reset();
    await renderToday();
    showToast("Collection added");
  });

  document.getElementById("expenseForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const place = document.getElementById("expensePlace").value.trim();
    const amount = document.getElementById("expenseAmount").value;
    if (!place || !amount) return;
    await addExpense(state.currentDate, place, amount);
    invalidateCache();
    e.target.reset();
    await renderToday();
    showToast("Petty fund expense logged");
  });

  document.getElementById("topupInput").addEventListener("change", async (e) => {
    await setTopup(state.currentDate, e.target.value);
    invalidateCache();
    await renderToday();
    showToast("Top-up saved");
  });

  document.addEventListener("click", async (e) => {
    const btn = e.target.closest(".del");
    if (!btn) return;
    const id = btn.dataset.id;
    if (btn.dataset.store === "collections") await deleteCollection(id);
    else await deleteExpense(id);
    invalidateCache();
    await renderToday();
  });

  document.getElementById("monthSelect").addEventListener("change", renderMonth);
  document.getElementById("monthYearSelect").addEventListener("change", renderMonth);
  document.getElementById("yearSelect").addEventListener("change", renderYear);

  document.getElementById("saveSettingsBtn").addEventListener("click", async () => {
    const startDate = document.getElementById("startDateInput").value;
    const openingBalance = document.getElementById("openingBalanceInput").value;
    if (!startDate) { showToast("Please pick a start date"); return; }
    const settings = await getSettings();
    await saveSettings({ startDate, openingBalance, currencyCode: settings.currencyCode });
    invalidateCache();
    showToast("Settings saved");
  });

  document.getElementById("headerCurrencySelect").addEventListener("change", (e) => applyCurrencyChange(e.target.value));
  document.getElementById("settingsCurrencySelect").addEventListener("change", (e) => applyCurrencyChange(e.target.value));

  document.getElementById("exportBtn").addEventListener("click", async () => {
    const year = Number(document.getElementById("exportYearSelect").value);
    showToast("Building Excel file…");
    try {
      await exportYearToExcel(year);
      showToast("Exported " + year);
    } catch (err) {
      console.error(err);
      showToast("Export failed — check connection for first-time load");
    }
  });

  document.getElementById("clearDataBtn").addEventListener("click", async () => {
    if (!confirm("This will permanently delete ALL ledger data in your account (all devices). Export a backup first if needed. Continue?")) return;
    await Promise.all(
      ["collections", "expenses", "topups", "settings", "dayOptions"].map((name) => userRef(name).remove())
    );
    invalidateCache();
    showToast("All data cleared");
    await switchView("today");
  });
}

/* ---------------- Auth UI wiring ---------------- */
function setLoginError(msg) {
  document.getElementById("loginError").textContent = msg || "";
}

function wireAuthEvents() {
  document.getElementById("loginBtn").addEventListener("click", async () => {
    const username = document.getElementById("loginUsername").value;
    const secretCode = document.getElementById("loginSecretCode").value;
    setLoginError("");
    if (!username.trim() || !secretCode) { setLoginError("Enter a username and secret code."); return; }
    try {
      await logIn(username, secretCode);
    } catch (err) {
      setLoginError(friendlyAuthError(err));
    }
  });

  document.getElementById("signupBtn").addEventListener("click", async () => {
    const username = document.getElementById("loginUsername").value;
    const secretCode = document.getElementById("loginSecretCode").value;
    setLoginError("");
    if (!username.trim() || !secretCode) { setLoginError("Enter a username and secret code."); return; }
    if (secretCode.length < 6) { setLoginError("Secret code must be at least 6 characters."); return; }
    try {
      await signUp(username, secretCode);
    } catch (err) {
      setLoginError(friendlyAuthError(err));
    }
  });

  document.getElementById("logoutBtn").addEventListener("click", async () => {
    if (!confirm("Log out of this device?")) return;
    await logOut();
  });
}

async function onLoggedIn(user) {
  currentUser = user;
  invalidateCache();
  document.body.classList.remove("auth-checking");
  document.body.classList.add("logged-in");
  document.getElementById("loginError").textContent = "";
  document.getElementById("loginUsername").value = "";
  document.getElementById("loginSecretCode").value = "";

  // Show a friendly username: prefer a just-completed signup, then the saved
  // profile name, then fall back to the local part of the internal email.
  let displayName = (user.email || "").split("@")[0];
  if (pendingSignupUsername) {
    displayName = pendingSignupUsername;
    pendingSignupUsername = null;
  } else {
    try {
      const snap = await rtdb.ref(`users/${user.uid}/profile/username`).once("value");
      if (snap.val()) displayName = snap.val();
    } catch (e) { /* ignore — fallback name above is fine */ }
  }
  document.getElementById("loggedInUsername").textContent = displayName;

  const settings = await getSettings();
  CURRENT_CURRENCY = settings.currencyCode;
  populateCurrencySelectors(settings.currencyCode);
  await switchView("today");
}

function onLoggedOut() {
  currentUser = null;
  document.body.classList.remove("auth-checking");
  document.body.classList.remove("logged-in");
}

/* ---------------- Init ---------------- */
window.addEventListener("DOMContentLoaded", () => {
  wireEvents();
  wireAuthEvents();

  auth.onAuthStateChanged(async (user) => {
    try {
      if (user) await onLoggedIn(user);
      else onLoggedOut();
    } catch (err) {
      console.error(err);
      document.body.classList.remove("auth-checking");
    }
  });

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
});
