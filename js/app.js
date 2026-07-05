// Family Finance Tracker - Core Controller Logic (Apple Minimalist Edition)

// Apps Script Deployment snippet for user copy-pasting (Dynamic Profiles & Simple Sync)
const APPS_SCRIPT_CODE = `// Google Apps Script Web App for Family Finance Tracker Sync
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action; 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ACTION: Create a sheet tab for a new profile family member
    if (action === "createSheet") {
      var sheetName = payload.name;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(["Transaction ID", "Date", "Category", "Description", "Type", "Amount"]);
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Sheet created for member " + sheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    // ACTION: Delete a sheet tab when a family member is removed
    if (action === "deleteSheet") {
      var sheetName = payload.name;
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.deleteSheet(sheet);
        return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Sheet deleted for member " + sheetName}))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Sheet not found"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    // TRANSACTION ACTIONS (Require specific member sheet tab)
    var tx = payload.data;
    var memberSheetName = tx.memberName || tx.memberId;
    
    var sheet = ss.getSheetByName(memberSheetName);
    if (!sheet) {
      sheet = ss.insertSheet(memberSheetName);
      sheet.appendRow(["Transaction ID", "Date", "Category", "Description", "Type", "Amount"]);
    }
    
    if (action === "add") {
      sheet.appendRow([
        tx.id, 
        tx.date, 
        tx.categoryName || tx.categoryId, 
        tx.description, 
        tx.type, 
        tx.amount
      ]);
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Transaction logged in " + memberSheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "delete") {
      var data = sheet.getDataRange().getValues();
      var deleted = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === tx.id) {
          sheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: deleted ? "Row deleted" : "Row not found"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid action"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// Fetch sheet tab names as profiles, and merge all rows as transactions
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    
    var profiles = [];
    var allTransactions = [];
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var sheetName = sheet.getName();
      
      // Skip the default Sheet1 if empty
      if (sheetName === "Sheet1" && sheet.getLastRow() <= 1) continue;
      
      // Each sheet tab represents a profile!
      profiles.push({
        id: sheetName.toLowerCase().replace(/[^a-z0-9]/g, '-'),
        name: sheetName,
        role: "Member",
        avatar: getAvatarForProfile(s, sheetName),
        color: getProfileColor(s),
        glow: "rgba(0, 122, 255, 0.15)"
      });
      
      var data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        // Expected headers: Transaction ID (0), Date (1), Category (2), Description (3), Type (4), Amount (5)
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (!row[0]) continue; // Skip blank rows
          
          allTransactions.push({
            id: row[0],
            date: Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
            categoryName: row[2],
            description: row[3],
            type: row[4],
            amount: parseFloat(row[5]),
            memberName: sheetName // Tab name represents the profile member name
          });
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success", 
      profiles: profiles, 
      transactions: allTransactions
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function getProfileColor(index) {
  var colors = ["#007aff", "#5856d6", "#ff9500", "#34c759", "#ff3b30"];
  return colors[index % colors.length];
}

function getAvatarForProfile(index, name) {
  var lower = name.toLowerCase();
  if (lower.indexOf("dad") !== -1 || lower.indexOf("marcus") !== -1) return "👨";
  if (lower.indexOf("mom") !== -1 || lower.indexOf("elena") !== -1) return "👩";
  if (lower.indexOf("teen") !== -1 || lower.indexOf("alex") !== -1) return "👦";
  if (lower.indexOf("kid") !== -1 || lower.indexOf("zoe") !== -1) return "👧";
  var avatars = ["👤", "👨", "👩", "👦", "👧", "👶"];
  return avatars[index % avatars.length];
}`;

// --- STATE MANAGEMENT ---
let transactions = [];
let budgets = {};
let goals = [];
let members = [];
let categories = {};

// Filter & App States
let selectedMemberId = 'all';
let currentView = 'dashboard-view';
let transactionType = 'expense';

// Settings States
let activeTheme = 'cyber-glass';
let currencySymbol = '₹'; // Default to Indian Rupee!
let currencyCode = 'INR';
let googleSheetUrl = 'https://script.google.com/macros/s/AKfycbwBgR14Sfj5jvaMqCaJ3tVEdhDjYRn5LgGf8ihBkYeK6ePYUYn6ov_Ax0w6zl0mSPky1A/exec';
let googleSheetSyncEnabled = true;

// PIN Settings States
let pinLockEnabled = false;
let enteredPin = '';

// DOM Elements
const totalBalanceEl = document.getElementById('total-balance');
const totalIncomeEl = document.getElementById('total-income');
const totalExpenseEl = document.getElementById('total-expense');
const membersContainer = document.getElementById('members-container');
const recentTransactionsContainer = document.getElementById('recent-transactions-container');
const fullTransactionsContainer = document.getElementById('full-transactions-container');
const budgetsContainer = document.getElementById('budgets-container');
const goalsContainer = document.getElementById('goals-container');
const txCountEl = document.getElementById('tx-count');

// Budget Tab Total Summary Elements
const totalBudgetRatioEl = document.getElementById('total-budget-ratio');
const totalBudgetProgressEl = document.getElementById('total-budget-progress');
const totalBudgetSpentEl = document.getElementById('total-budget-spent');
const totalBudgetLimitEl = document.getElementById('total-budget-limit');

// Modal Elements
const addTxOverlay = document.getElementById('add-transaction-overlay');
const addTxForm = document.getElementById('add-transaction-form');
const btnOpenAddTx = document.getElementById('btn-open-add-tx');
const btnCloseAddTx = document.getElementById('btn-close-add-tx');

const goalOverlay = document.getElementById('goal-contribution-overlay');
const goalForm = document.getElementById('goal-contribution-form');
const btnCloseGoalModal = document.getElementById('btn-close-goal-modal');
const contributionGoalIdInput = document.getElementById('contribution-goal-id');
const contributionGoalLabel = document.getElementById('contribution-goal-label');
const contributionAmountInput = document.getElementById('contribution-amount-input');

// Toggle Type Buttons
const typeExpenseBtn = document.getElementById('type-expense-btn');
const typeIncomeBtn = document.getElementById('type-income-btn');
const formGroupCategory = document.getElementById('form-group-category');

// Reset Button & Inputs
const btnResetData = document.getElementById('btn-reset-data');
const searchTxInput = document.getElementById('search-tx');

// --- SETTINGS DRAWER DOM ELEMENTS ---
const settingsDrawer = document.getElementById('settings-drawer');
const btnOpenSettings = document.getElementById('btn-open-settings');
const btnCloseSettings = document.getElementById('btn-close-settings');
const themeSelect = document.getElementById('theme-select');
const currencySelect = document.getElementById('currency-select');
const profilesEditList = document.getElementById('settings-profiles-list');

const btnToggleAddProfile = document.getElementById('btn-toggle-add-profile-form');
const addProfileForm = document.getElementById('add-profile-form');
const btnToggleAddCat = document.getElementById('btn-toggle-add-cat-form');
const addCatForm = document.getElementById('add-category-form');

const sheetSyncToggle = document.getElementById('sheet-sync-toggle');
const sheetUrlInput = document.getElementById('sheet-url-input');
const btnTestSheet = document.getElementById('btn-test-sheet');
const btnShowCode = document.getElementById('btn-show-code');
const codeSnippetBox = document.getElementById('code-snippet-box');

// --- INIT APP ---
function init() {
  // Register Service Worker for PWA compliance
  if ('serviceWorker' in navigator) {
    const isGithub = window.location.hostname.includes('github.io');
    const swPath = isGithub ? '/Finance-Tracker/sw.js' : './sw.js';
    navigator.serviceWorker.register(swPath)
      .then(() => console.log('Service Worker Registered:', swPath))
      .catch(err => console.error('Service Worker Registry Failed:', err));
  }

  loadData();
  setupEventListeners();
  
  // Set default form values
  document.getElementById('tx-date-input').value = new Date().toISOString().split('T')[0];
  
  // Setup Apps Script Text Block
  codeSnippetBox.innerText = APPS_SCRIPT_CODE;

  // Apply visual configurations
  applyTheme(activeTheme);
  applyCurrency(currencyCode);
  
  // Sync UI inputs with states
  syncSettingsUI();

  // Populate Dropdowns
  populateFormDropdowns();
  
  // Initial Render
  renderAll();

  // Pull records from Google Sheets on boot if sync is enabled
  if (googleSheetSyncEnabled && googleSheetUrl) {
    syncFromGoogleSheets();
  }
}

// Load configurations from storage
function loadData() {
  const localTx = localStorage.getItem('orbit_tx');
  const localBudgets = localStorage.getItem('orbit_budgets');
  const localGoals = localStorage.getItem('orbit_goals');
  const localMembers = localStorage.getItem('orbit_members');
  const localCategories = localStorage.getItem('orbit_categories');
  
  // Settings keys
  const localTheme = localStorage.getItem('orbit_theme');
  const localCurrencyCode = localStorage.getItem('orbit_currency_code');
  const localSheetUrl = localStorage.getItem('orbit_sheet_url');
  const localSheetSync = localStorage.getItem('orbit_sheet_sync');

  // Load Transactions, Budgets, Goals
  if (localTx && localBudgets && localGoals) {
    transactions = JSON.parse(localTx);
    budgets = JSON.parse(localBudgets);
    goals = JSON.parse(localGoals);
  } else {
    transactions = [];
    budgets = { ...window.INITIAL_BUDGETS };
    goals = JSON.parse(JSON.stringify(window.INITIAL_GOALS));
  }

  // Load Profiles (Members)
  if (localMembers) {
    members = JSON.parse(localMembers);
  } else {
    members = JSON.parse(JSON.stringify(window.INITIAL_MEMBERS));
  }

  // Load Categories
  if (localCategories) {
    categories = JSON.parse(localCategories);
  } else {
    categories = { ...window.CATEGORIES };
  }
  // Merge loaded categories with default categories to get new ones
  categories = { ...window.CATEGORIES, ...categories };

  // Load Setting variables
  activeTheme = localTheme || 'modern-light';
  currencyCode = localCurrencyCode || 'INR'; // Indian Rupee is default
  googleSheetUrl = (localSheetUrl && localSheetUrl.trim() !== '') ? localSheetUrl : 'https://script.google.com/macros/s/AKfycbwBgR14Sfj5jvaMqCaJ3tVEdhDjYRn5LgGf8ihBkYeK6ePYUYn6ov_Ax0w6zl0mSPky1A/exec';
  googleSheetSyncEnabled = localSheetSync !== null ? localSheetSync === 'true' : true;
  
  // Load PIN settings
  const localPinLock = localStorage.getItem('orbit_pin_lock');
  pinLockEnabled = localPinLock === 'true';
  
  saveToStorage();
}

function saveToStorage() {
  localStorage.setItem('orbit_tx', JSON.stringify(transactions));
  localStorage.setItem('orbit_budgets', JSON.stringify(budgets));
  localStorage.setItem('orbit_goals', JSON.stringify(goals));
  localStorage.setItem('orbit_members', JSON.stringify(members));
  localStorage.setItem('orbit_categories', JSON.stringify(categories));
  
  localStorage.setItem('orbit_theme', activeTheme);
  localStorage.setItem('orbit_currency_code', currencyCode);
  localStorage.setItem('orbit_sheet_url', googleSheetUrl);
  localStorage.setItem('orbit_sheet_sync', googleSheetSyncEnabled.toString());
  localStorage.setItem('orbit_pin_lock', pinLockEnabled.toString());
}

// Re-sync input states in Settings Drawer
function syncSettingsUI() {
  themeSelect.value = activeTheme;
  currencySelect.value = currencyCode;
  sheetSyncToggle.checked = googleSheetSyncEnabled;
  sheetUrlInput.value = googleSheetUrl;
  
  const pinLockToggle = document.getElementById('pin-lock-toggle');
  if (pinLockToggle) {
    pinLockToggle.checked = pinLockEnabled;
  }
}

// Reset Database completely
function resetDatabase() {
  if (confirm("Are you sure you want to delete your local storage cache? Ledger transactions will be wiped.")) {
    localStorage.clear();
    selectedMemberId = 'all';
    document.body.className = '';
    
    loadData();
    applyTheme(activeTheme);
    applyCurrency(currencyCode);
    syncSettingsUI();
    populateFormDropdowns();
    renderAll();
    
    showToast("Local data reset successfully.");
  }
}

// --- CORE RENDER PANEL ---
function renderAll() {
  renderMembers();
  updateMetricsAndCharts();
  renderRecentTransactions();
  renderFullTransactions();
  renderBudgets();
  renderGoals();
  renderSettingsProfiles();
  renderAccountsSlider(); // Render accounts balances slider card
}

// Renders the member avatar slider at the top
function renderMembers() {
  let html = `
    <div class="member-chip ${selectedMemberId === 'all' ? 'active' : ''}" data-id="all" style="--member-color: var(--apple-blue); --member-glow: rgba(0, 122, 255, 0.15)">
      <div class="avatar-wrapper">👥</div>
      <span>All Members</span>
    </div>
  `;
  
  members.forEach(m => {
    html += `
      <div class="member-chip ${selectedMemberId === m.id ? 'active' : ''}" data-id="${m.id}" style="--member-color: ${m.color}; --member-glow: ${m.glow || 'rgba(0,0,0,0.05)'}">
        <div class="avatar-wrapper">${m.avatar}</div>
        <span>${m.name.split(' ')[0]}</span>
      </div>
    `;
  });
  
  membersContainer.innerHTML = html;
  
  // Attach Click Handlers to Member Chips
  document.querySelectorAll('.member-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      selectedMemberId = chip.getAttribute('data-id');
      document.querySelectorAll('.member-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      
      updateMetricsAndCharts();
      renderRecentTransactions();
      renderFullTransactions();
      renderBudgets();
    });
  });
}

// Helper to filter transactions based on selected member
function getFilteredTransactions() {
  if (selectedMemberId === 'all') {
    return transactions;
  }
  return transactions.filter(t => t.memberId === selectedMemberId);
}

// Recalculates metrics and draws SVG charts
function updateMetricsAndCharts() {
  const filtered = getFilteredTransactions();
  
  let incomeSum = 0;
  let expenseSum = 0;
  
  filtered.forEach(t => {
    const amt = parseFloat(t.amount);
    if (t.type === 'income') {
      incomeSum += amt;
    } else {
      expenseSum += amt;
    }
  });
  
  const netBalance = incomeSum - expenseSum;
  
  // Count animations
  animateValue(totalBalanceEl, netBalance, true);
  animateValue(totalIncomeEl, incomeSum, false, '+');
  animateValue(totalExpenseEl, expenseSum, false, '-');
  
  // 1. Category sum calculations for Donut
  const categorySums = {};
  Object.keys(categories).forEach(k => {
    if (k !== 'salary') {
      categorySums[k] = 0;
    }
  });
  
  filtered.forEach(t => {
    if (t.type === 'expense' && categorySums[t.categoryId] !== undefined) {
      categorySums[t.categoryId] += parseFloat(t.amount);
    }
  });

  const donutSegments = Object.keys(categorySums).map(k => {
    return {
      id: k,
      value: categorySums[k],
      color: categories[k].color,
      label: categories[k].name
    };
  });
  
  window.Charts.renderDonutChart('donut-chart-container', donutSegments, currencySymbol);
  renderDonutLegend(donutSegments);

  // 2. Trend Timeline calculations
  const expenseTxs = filtered.filter(t => t.type === 'expense');
  const dateMap = {};
  expenseTxs.forEach(t => {
    dateMap[t.date] = (dateMap[t.date] || 0) + parseFloat(t.amount);
  });
  
  const sortedDates = Object.keys(dateMap).sort();
  const lastDates = sortedDates.slice(-6);
  const chartPoints = [];
  
  if (lastDates.length < 4) {
    const today = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(today.getDate() - i);
      const dateStr = d.toISOString().split('T')[0];
      const val = dateMap[dateStr] || 0;
      
      const parts = d.toDateString().split(' ');
      const label = `${parts[1]} ${parts[2]}`;
      chartPoints.push({ label, value: val });
    }
  } else {
    lastDates.forEach(dateStr => {
      const d = new Date(dateStr + 'T00:00:00');
      const parts = d.toDateString().split(' ');
      const label = `${parts[1]} ${parts[2]}`;
      chartPoints.push({ label, value: dateMap[dateStr] });
    });
  }

  // Chart line color matches theme preferences or selected member color
  let chartLineColor = activeTheme === 'solar-light' ? '#b07d00' : '#007aff';
  if (selectedMemberId !== 'all') {
    const member = members.find(m => m.id === selectedMemberId);
    if (member) chartLineColor = member.color;
  }
  window.Charts.renderLineChart('line-chart-container', chartPoints, chartLineColor, currencySymbol);
}

// Renders the legends of the donut breakdown
function renderDonutLegend(segments) {
  const legendContainer = document.getElementById('donut-legend');
  if (!legendContainer) return;
  
  const total = segments.reduce((acc, s) => acc + s.value, 0);
  if (total === 0) {
    legendContainer.innerHTML = '';
    return;
  }
  
  let html = '';
  segments.filter(s => s.value > 0).sort((a,b) => b.value - a.value).forEach(s => {
    const pct = ((s.value / total) * 100).toFixed(0);
    html += `
      <div class="legend-item">
        <span class="legend-dot" style="background: ${s.color};"></span>
        <span>${s.label}: ${pct}%</span>
      </div>
    `;
  });
  legendContainer.innerHTML = html;
}

// Renders the top 4 transactions on dashboard
function renderRecentTransactions() {
  const filtered = getFilteredTransactions();
  const sorted = [...filtered].sort((a, b) => new Date(b.date) - new Date(a.date));
  const recent = sorted.slice(0, 4);
  
  if (recent.length === 0) {
    recentTransactionsContainer.innerHTML = `
      <div class="empty-state">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
        </svg>
        <p>No transactions logged</p>
      </div>
    `;
    return;
  }
  
  recentTransactionsContainer.innerHTML = recent.map(t => getTransactionRowHTML(t)).join('');
  attachDeleteHandlers();
}

// Renders the full searchable transaction database
function renderFullTransactions() {
  const filtered = getFilteredTransactions();
  const searchVal = searchTxInput.value.toLowerCase().trim();
  
  let result = [...filtered].sort((a, b) => new Date(b.date) - new Date(a.date));
  
  if (searchVal) {
    result = result.filter(t => {
      const descMatch = t.description.toLowerCase().includes(searchVal);
      const cat = categories[t.categoryId];
      const catMatch = cat ? cat.name.toLowerCase().includes(searchVal) : false;
      const amountMatch = t.amount.toString().includes(searchVal);
      return descMatch || catMatch || amountMatch;
    });
  }
  
  txCountEl.innerText = `${result.length} Transaction${result.length === 1 ? '' : 's'}`;
  
  if (result.length === 0) {
    fullTransactionsContainer.innerHTML = `
      <div class="empty-state">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="11" cy="11" r="8"></circle>
          <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          <line x1="8" y1="11" x2="14" y2="11"></line>
        </svg>
        <p>No transactions match search criteria</p>
      </div>
    `;
    return;
  }
  
  fullTransactionsContainer.innerHTML = result.map(t => getTransactionRowHTML(t)).join('');
  attachDeleteHandlers();
}

// Templates single transaction list item
function getTransactionRowHTML(t) {
  const cat = categories[t.categoryId] || { name: 'Expense', icon: 'fa-shopping-basket', color: '#007aff' };
  const member = members.find(m => m.id === t.memberId) || { name: 'Member', color: '#8e8e93' };
  const isIncome = t.type === 'income';
  
  const d = new Date(t.date + 'T00:00:00');
  const formattedDate = d.toLocaleDateString(undefined, {month: 'short', day: 'numeric', year: 'numeric'});
  
  return `
    <div class="tx-item">
      <div class="tx-item-left">
        <div class="tx-flow-indicator ${isIncome ? 'income' : 'expense'}">
          <i class="fas ${isIncome ? 'fa-arrow-up' : 'fa-arrow-down'}"></i>
        </div>
        <div class="tx-icon-box" style="color: ${cat.color};">
          <i class="fas ${cat.icon}"></i>
        </div>
        <div class="tx-details">
          <span class="tx-desc">${t.description}</span>
          <span class="tx-meta">
            <span class="tx-member-tag" style="--member-color: ${member.color}">${member.name.split(' ')[0]}</span>
            <span>${cat.name}</span>
          </span>
        </div>
      </div>
      <div class="tx-amount-box">
        <span class="tx-amount ${isIncome ? 'text-income' : 'text-expense'}">
          ${isIncome ? '+' : '-'}${currencySymbol}${parseFloat(t.amount).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}
        </span>
        <span class="tx-date">${formattedDate}</span>
      </div>
      <button class="btn-delete-tx" data-id="${t.id}" title="Delete transaction">
        <i class="fas fa-trash-alt"></i>
      </button>
    </div>
  `;
}

function attachDeleteHandlers() {
  document.querySelectorAll('.btn-delete-tx').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-id');
      deleteTransaction(id);
    };
  });
}

function deleteTransaction(id) {
  const idx = transactions.findIndex(t => t.id === id);
  if (idx !== -1) {
    const tx = transactions[idx];
    transactions.splice(idx, 1);
    saveToStorage();
    
    // Sync deletion to Google Sheet tab background
    syncTransactionToGoogleSheet('delete', tx);
    
    renderAll();
    
    // Refresh category modal details if currently open
    const catOverlay = document.getElementById('category-details-overlay');
    if (catOverlay && catOverlay.classList.contains('active') && tx.type === 'expense') {
      openCategoryDetails(tx.categoryId);
    }
    
    showToast(`Transaction "${tx.description}" deleted.`);
  }
}

// Renders category budget allocations
function renderBudgets() {
  const filtered = getFilteredTransactions();
  const categorySpent = {};
  
  Object.keys(categories).forEach(k => {
    if (k !== 'salary') categorySpent[k] = 0;
  });
  
  filtered.forEach(t => {
    if (t.type === 'expense' && categorySpent[t.categoryId] !== undefined) {
      categorySpent[t.categoryId] += parseFloat(t.amount);
    }
  });

  let totalSpent = 0;
  let totalLimit = 0;
  let budgetsHtml = '';
  
  Object.keys(budgets).forEach(catId => {
    const limit = budgets[catId];
    const spent = categorySpent[catId] || 0;
    const cat = categories[catId];
    
    if (!cat) return;
    
    totalSpent += spent;
    totalLimit += limit;
    
    const percent = limit > 0 ? Math.min((spent / limit) * 100, 100) : 0;
    
    let warningClass = '';
    if (percent >= 90) warningClass = 'danger';
    else if (percent >= 75) warningClass = 'warning';
    
    budgetsHtml += `
      <div class="glass-panel budget-item">
        <div class="budget-meta">
          <span class="budget-cat-name" style="--cat-color: ${cat.color}">
            <i class="fas ${cat.icon}"></i>
            ${cat.name}
          </span>
          <span class="budget-amounts">
            <strong>${currencySymbol}${spent.toLocaleString(undefined, {maximumFractionDigits:0})}</strong> of 
            ${currencySymbol}<input type="number" class="budget-limit-input" data-cat-id="${catId}" value="${limit}" style="width: 55px; background: transparent; border: none; border-bottom: 1.5px dashed var(--apple-blue); color: var(--apple-text); font-weight: 700; text-align: right; font-size: 0.72rem; padding: 0 2px;">
          </span>
        </div>
        <div class="progress-track">
          <div class="progress-bar ${warningClass}" style="width: ${percent}%; --cat-color: ${cat.color}; background: ${cat.color};"></div>
        </div>
      </div>
    `;
  });
  
  budgetsContainer.innerHTML = budgetsHtml;
  
  // Attach inline budget limit change listeners
  document.querySelectorAll('.budget-limit-input').forEach(input => {
    input.addEventListener('change', () => {
      const catId = input.getAttribute('data-cat-id');
      const val = parseFloat(input.value);
      if (!isNaN(val) && val >= 0) {
        budgets[catId] = val;
        saveToStorage();
        renderAll();
        showToast(`Budget for ${categories[catId].name} updated to ${currencySymbol}${val}.`);
      }
    });
    // Prevent event propagation
    input.addEventListener('click', (e) => e.stopPropagation());
  });
  
  const totalPercent = totalLimit > 0 ? Math.min((totalSpent / totalLimit) * 100, 100) : 0;
  totalBudgetRatioEl.innerText = `${totalPercent.toFixed(0)}% Limit Reached`;
  totalBudgetProgressEl.style.width = `${totalPercent}%`;
  
  totalBudgetProgressEl.className = 'progress-bar';
  if (totalPercent >= 90) totalBudgetProgressEl.classList.add('danger');
  else if (totalPercent >= 75) totalBudgetProgressEl.classList.add('warning');

  totalBudgetSpentEl.innerText = `${currencySymbol}${totalSpent.toLocaleString(undefined, {maximumFractionDigits:0})} Spent`;
  totalBudgetLimitEl.innerText = `of ${currencySymbol}${totalLimit.toLocaleString(undefined, {maximumFractionDigits:0})} limit`;
}

// Renders family savings goals
function renderGoals() {
  goalsContainer.innerHTML = goals.map(g => {
    const pct = g.target > 0 ? ((g.current / g.target) * 100).toFixed(0) : 0;
    const progressSVG = window.Charts.getRadialProgressSVG(g.current, g.target, g.color);
    
    return `
      <div class="glass-panel goal-card" style="border-left: 3px solid ${g.color}; position: relative;">
        <div class="goal-details" style="flex: 1; padding-right: 8px;">
          <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
            <span class="goal-title" style="font-weight: 700; color: var(--apple-text);">${g.name}</span>
            <button class="btn-link btn-edit-goal" data-id="${g.id}" style="font-size: 0.68rem; color: var(--apple-blue); border: none; background: none; cursor: pointer; padding: 0;" title="Edit Goal">
              <i class="fas fa-pencil-alt"></i>
            </button>
            <button class="btn-link btn-delete-goal" data-id="${g.id}" style="font-size: 0.68rem; color: var(--apple-red); border: none; background: none; cursor: pointer; padding: 0;" title="Delete Goal">
              <i class="fas fa-trash-alt"></i>
            </button>
          </div>
          <div class="goal-numbers">
            <span style="--goal-color: ${g.color}; font-weight: 700;">${currencySymbol}${g.current.toLocaleString()}</span> of ${currencySymbol}${g.target.toLocaleString()} (${pct}%)
          </div>
        </div>
        <div class="goal-progress-box" style="flex-shrink: 0; display: flex; align-items: center; gap: 10px;">
          ${progressSVG}
          <button class="btn-contribute" data-id="${g.id}" data-name="${g.name}" data-color="${g.color}" title="Add Contribution">
            <i class="fas fa-plus"></i>
          </button>
        </div>
      </div>
    `;
  }).join('');
  
  // Attach Goal click listeners
  document.querySelectorAll('.btn-contribute').forEach(btn => {
    btn.onclick = () => {
      const id = btn.getAttribute('data-id');
      const name = btn.getAttribute('data-name');
      const color = btn.getAttribute('data-color');
      
      contributionGoalIdInput.value = id;
      contributionGoalLabel.innerText = `Contribute to ${name}`;
      contributionAmountInput.value = '';
      
      const submitBtn = goalForm.querySelector('.glass-submit-btn');
      submitBtn.style.background = color;
      submitBtn.style.boxShadow = `0 4px 10px rgba(${hexToRgb(color)}, 0.15)`;

      openOverlay(goalOverlay);
    };
  });

  // Attach Goal Edit/Delete listeners
  document.querySelectorAll('.btn-edit-goal').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-id');
      const g = goals.find(x => x.id === id);
      if (g) {
        document.getElementById('edit-goal-id').value = g.id;
        document.getElementById('edit-goal-name-input').value = g.name;
        document.getElementById('edit-goal-target-input').value = g.target;
        document.getElementById('edit-goal-current-input').value = g.current;
        document.getElementById('edit-goal-color-select').value = g.color;
        
        openOverlay(document.getElementById('edit-goal-overlay'));
      }
    };
  });

  document.querySelectorAll('.btn-delete-goal').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.getAttribute('data-id');
      const g = goals.find(x => x.id === id);
      if (g && confirm(`Delete savings goal "${g.name}"?`)) {
        goals = goals.filter(x => x.id !== id);
        saveToStorage();
        renderAll();
        showToast(`Goal "${g.name}" deleted.`);
      }
    };
  });
}

// Renders the list of Profiles editable in Settings Drawer
function renderSettingsProfiles() {
  let html = '';
  members.forEach(m => {
    const showDelete = members.length > 1;
    html += `
      <div class="profile-edit-item" style="border-left: 2px solid ${m.color}; display: flex; align-items: center; justify-content: space-between; padding: 10px 12px;">
        <div class="profile-edit-info" style="display: flex; align-items: center; gap: 10px;">
          <span class="profile-edit-avatar" style="font-size: 1.25rem;">${m.avatar}</span>
          <div class="profile-edit-text" style="display: flex; flex-direction: column;">
            <span class="profile-edit-name" style="font-weight: 700; font-size: 0.78rem;">${m.name}</span>
            <span class="profile-edit-role" style="font-size: 0.65rem; color: var(--apple-gray);">${m.role}</span>
          </div>
        </div>
        <div style="display: flex; gap: 8px; align-items: center;">
          <button class="btn-link btn-edit-profile" data-id="${m.id}" style="color: var(--apple-blue); cursor: pointer; border: none; background: none; font-size: 0.75rem;" title="Edit Profile">
            <i class="fas fa-pencil-alt"></i>
          </button>
          ${showDelete ? `
            <button class="btn-delete-profile" data-id="${m.id}" style="color: var(--apple-red); cursor: pointer; border: none; background: none; font-size: 0.75rem;" title="Remove Member">
              <i class="fas fa-trash-alt"></i>
            </button>
          ` : ''}
        </div>
      </div>
    `;
  });
  
  profilesEditList.innerHTML = html;
  
  // Hook profiles edit
  document.querySelectorAll('.btn-edit-profile').forEach(btn => {
    btn.onclick = () => {
      const id = btn.getAttribute('data-id');
      const m = members.find(x => x.id === id);
      if (m) {
        document.getElementById('edit-profile-id').value = m.id;
        document.getElementById('edit-profile-name-input').value = m.name;
        document.getElementById('edit-profile-role-input').value = m.role;
        document.getElementById('edit-profile-avatar-input').value = m.avatar;
        document.getElementById('edit-profile-color-select').value = m.color;
        
        openOverlay(document.getElementById('edit-profile-overlay'));
      }
    };
  });
  
  // Hook profiles deletion
  document.querySelectorAll('.btn-delete-profile').forEach(btn => {
    btn.onclick = () => {
      const id = btn.getAttribute('data-id');
      const member = members.find(m => m.id === id);
      if (confirm(`Remove family member "${member.name}"? This deletes their Google Sheet tab.`)) {
        
        // Remove locally
        members = members.filter(m => m.id !== id);
        if (selectedMemberId === id) {
          selectedMemberId = 'all';
        }
        saveToStorage();
        
        // Sync profile sheet deletion to Google Sheets in background
        syncProfileToGoogleSheet('delete', member.name);
        
        populateFormDropdowns();
        renderAll();
        showToast(`Member "${member.name}" removed.`);
      }
    };
  });
}

// --- APPLICATORS: THEME & CURRENCY ---

function applyTheme(themeName) {
  document.body.classList.remove('theme-modern-light', 'theme-modern-dark', 'theme-futuristic');
  
  // Add new theme class (cyber-glass=Apple Light, neo-minimalist=Apple Dark, solar-light=Solar Amber)
  document.body.classList.add(`theme-${themeName}`);
  activeTheme = themeName;
  saveToStorage();
  
  // Re-draw charts to adapt text colors and styles
  if (document.getElementById('dashboard-view').classList.contains('active')) {
    updateMetricsAndCharts();
  }
}

// Applies Currency changes
function applyCurrency(code) {
  currencyCode = code;
  switch (code) {
    case 'INR': currencySymbol = '₹'; break;
    case 'EUR': currencySymbol = '€'; break;
    case 'GBP': currencySymbol = '£'; break;
    case 'USD': 
    default: 
      currencySymbol = '$'; 
      currencyCode = 'USD';
      break;
  }
  
  saveToStorage();
  
  // Update UI Forms label
  const amountLabel = document.getElementById('tx-amount-label');
  if (amountLabel) {
    amountLabel.innerText = `Amount (${currencySymbol})`;
  }
  
  // Recount layout values
  renderAll();
}

// Populates dropdown lists in modals
function populateFormDropdowns() {
  // 1. Categories Select (Filtered by type)
  const catSelect = document.getElementById('tx-category-select');
  if (catSelect) {
    let catHtml = '';
    Object.keys(categories).forEach(k => {
      const cat = categories[k];
      const isIncome = transactionType === 'income';
      const catType = cat.type || (k === 'salary' ? 'income' : 'expense');
      
      if (isIncome && catType === 'income') {
        catHtml += `<option value="${k}">${cat.name}</option>`;
      } else if (!isIncome && catType === 'expense') {
        catHtml += `<option value="${k}">${cat.name}</option>`;
      }
    });
    catSelect.innerHTML = catHtml;
  }

  // 2. Members Select
  const memberSelect = document.getElementById('tx-member-select');
  if (memberSelect) {
    let memHtml = '';
    members.forEach(m => {
      memHtml += `<option value="${m.id}">${m.name}</option>`;
    });
    memberSelect.innerHTML = memHtml;
  }
}

// --- GOOGLE SHEETS NETWORK SYNC ---

// Fetches all sheets and synchronizes state locally (Only sheets records reflect)
async function syncFromGoogleSheets() {
  if (!googleSheetSyncEnabled || !googleSheetUrl) return;
  
  const syncBtnIcon = document.querySelector('#btn-reset-data i');
  if (syncBtnIcon) syncBtnIcon.classList.add('syncing-spin');
  showToast("Downloading transaction data...");
  
  try {
    const res = await fetch(googleSheetUrl);
    const result = await res.json();
    
    if (result.status === 'success') {
      // 1. Parse Profiles (Members) dynamically from Sheet Tab Names!
      if (result.profiles && result.profiles.length > 0) {
        members = result.profiles;
      } else {
        // Fallback default if sheet is completely empty
        members = [{ id: 'operator', name: 'Primary Member', role: 'Head of Family', avatar: '👨', color: '#007aff', glow: 'rgba(0, 122, 255, 0.15)' }];
      }

      // 2. Parse Transactions
      const fetchedTxs = [];
      if (result.transactions) {
        result.transactions.forEach(t => {
          // Resolve Member ID matching sheet name
          let member = members.find(m => m.name === t.memberName);
          if (!member) {
            // Backup creation
            const newId = t.memberName.toLowerCase().replace(/[^a-z0-9]/g, '-') + '-' + Date.now().toString().slice(-4);
            member = {
              id: newId,
              name: t.memberName,
              role: 'Member',
              avatar: '👤',
              color: '#007aff',
              glow: 'rgba(0, 122, 255, 0.15)'
            };
            members.push(member);
          }

          // Resolve Category ID matching category name
          let categoryId = 'daily'; // default
          Object.keys(categories).forEach(k => {
            if (categories[k].name === t.categoryName) {
              categoryId = k;
            }
          });
          if (t.type === 'income') categoryId = 'salary';

          fetchedTxs.push({
            id: t.id,
            date: t.date,
            memberId: member.id,
            categoryId: categoryId,
            description: t.description,
            type: t.type,
            amount: t.amount
          });
        });
      }

      transactions = fetchedTxs;
      saveToStorage();
      renderAll();
      populateFormDropdowns();
      
      showToast("Sync completed successfully.");
    } else {
      showToast("Sync warning: Empty database response.");
    }
  } catch (err) {
    console.error("GET Fetch error: ", err);
    showToast("Fetch failed. Using local cache.");
  } finally {
    if (syncBtnIcon) syncBtnIcon.classList.remove('syncing-spin');
  }
}

// Syncs Profile sheet additions/deletions
async function syncProfileToGoogleSheet(action, profileName) {
  if (!googleSheetSyncEnabled || !googleSheetUrl) return;
  
  const payload = {
    action: action === 'add' ? 'createSheet' : 'deleteSheet',
    name: profileName
  };
  
  showToast("Updating sheet tabs...");
  
  try {
    await fetch(googleSheetUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    showToast("Sheet tabs updated.");
  } catch (err) {
    console.error("Profile sync error: ", err);
    showToast("Failed to sync sheets.");
  }
}

// Syncs transaction logs to specific profile sheet tabs
async function syncTransactionToGoogleSheet(action, transaction) {
  if (!googleSheetSyncEnabled || !googleSheetUrl) return;
  
  const memberObj = members.find(m => m.id === transaction.memberId);
  const categoryObj = categories[transaction.categoryId];
  
  const enrichedTx = {
    ...transaction,
    memberName: memberObj ? memberObj.name : transaction.memberId,
    categoryName: categoryObj ? categoryObj.name : transaction.categoryId
  };
  
  const payload = {
    action: action,
    data: enrichedTx
  };
  
  showToast("Syncing transaction...");
  
  try {
    await fetch(googleSheetUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    showToast("Sync completed.");
  } catch (err) {
    console.error("Transaction sync error: ", err);
    showToast("Connection Sync Pipeline Error.");
  }
}

// Pings Google Sheet Apps Script URL with a dummy request
async function testGoogleSheetConnection() {
  if (!googleSheetUrl) {
    showToast("Please provide Web App URL.");
    return;
  }
  
  showToast("Connecting to Google Sheets...");
  
  try {
    const res = await fetch(googleSheetUrl);
    const data = await res.json();
    if (data.status === 'success') {
      alert(`GOOGLE SHEET CONNECTION ACTIVE:\nEndpoint is responding successfully.`);
      showToast("Handshake active.");
    } else {
      alert(`SYNC RESPONSE:\n${JSON.stringify(data)}`);
    }
  } catch (err) {
    console.error(err);
    alert(`Sync Handshake Failed:\n1. Verify script is deployed as Web App.\n2. Execute as: "Me".\n3. Access: "Anyone".\n\nDetails: ${err.message}`);
    showToast("Handshake Failed.");
  }
}

// --- CORE SYSTEM LISTENERS ---
function setupEventListeners() {
  // Navigation triggering views
  document.querySelectorAll('.nav-trigger').forEach(trigger => {
    trigger.addEventListener('click', (e) => {
      e.preventDefault();
      const target = trigger.getAttribute('data-target');
      
      document.querySelectorAll('.bottom-nav .nav-item').forEach(item => {
        item.classList.remove('active');
        if (item.getAttribute('data-target') === target) {
          item.classList.add('active');
        }
      });
      
      document.querySelectorAll('.view-panel').forEach(panel => {
        panel.classList.remove('active');
      });
      const targetPanel = document.getElementById(target);
      if (targetPanel) {
        targetPanel.classList.add('active');
      }
      
      currentView = target;
      
      if (target === 'dashboard-view') {
        updateMetricsAndCharts();
      }
    });
  });

  // Modal controls: Add Transaction
  btnOpenAddTx.addEventListener('click', () => {
    openOverlay(addTxOverlay);
  });
  btnCloseAddTx.addEventListener('click', () => {
    closeOverlay(addTxOverlay);
  });
  addTxOverlay.addEventListener('click', (e) => {
    if (e.target === addTxOverlay) closeOverlay(addTxOverlay);
  });

  // Modal controls: Goal Contribution
  btnCloseGoalModal.addEventListener('click', () => {
    closeOverlay(goalOverlay);
  });
  goalOverlay.addEventListener('click', (e) => {
    if (e.target === goalOverlay) closeOverlay(goalOverlay);
  });

  // Transaction form type toggle
  typeExpenseBtn.addEventListener('click', () => {
    transactionType = 'expense';
    typeExpenseBtn.classList.add('active');
    typeIncomeBtn.classList.remove('active');
    formGroupCategory.style.display = 'flex';
    document.getElementById('tx-category-select').required = true;
  });
  
  typeIncomeBtn.addEventListener('click', () => {
    transactionType = 'income';
    typeIncomeBtn.classList.add('active');
    typeExpenseBtn.classList.remove('active');
    formGroupCategory.style.display = 'none';
    document.getElementById('tx-category-select').required = false;
  });

  // Reset/Sync database button in header
  btnResetData.addEventListener('click', syncFromGoogleSheets);

  // Search input change events
  searchTxInput.addEventListener('input', renderFullTransactions);

  // Submit Add Transaction Form
  addTxForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const amount = parseFloat(document.getElementById('tx-amount-input').value);
    const memberId = document.getElementById('tx-member-select').value;
    const date = document.getElementById('tx-date-input').value;
    const description = document.getElementById('tx-desc-input').value.trim();
    const categoryId = document.getElementById('tx-category-select').value;
    const account = document.getElementById('tx-account-select').value;

    if (isNaN(amount) || amount <= 0 || !date || !description) {
      showToast("Fill in all transaction fields.");
      return;
    }

    const newTx = {
      id: 'tx-' + Date.now(),
      type: transactionType,
      memberId,
      categoryId,
      amount,
      date,
      description,
      account
    };

    transactions.push(newTx);
    saveToStorage();
    
    // Sync addition to Google Sheet tab background
    syncTransactionToGoogleSheet('add', newTx);

    renderAll();
    
    closeOverlay(addTxOverlay);
    addTxForm.reset();
    document.getElementById('tx-date-input').value = new Date().toISOString().split('T')[0];
    
    showToast(`Added: ${currencySymbol}${amount} for ${description}`);
  });

  // Submit Goal Contribution Form
  goalForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const goalId = contributionGoalIdInput.value;
    const amount = parseInt(contributionAmountInput.value);
    
    if (isNaN(amount) || amount <= 0) {
      showToast("Enter a valid contribution amount.");
      return;
    }
    
    const goal = goals.find(g => g.id === goalId);
    if (goal) {
      goal.current += amount;
      
      const newTx = {
        id: 'tx-' + Date.now(),
        type: 'expense',
        memberId: selectedMemberId === 'all' ? members[0].id : selectedMemberId,
        categoryId: 'investments',
        amount,
        date: new Date().toISOString().split('T')[0],
        description: `Goal Deposit: ${goal.name}`
      };
      
      transactions.push(newTx);
      saveToStorage();
      
      // Sync transaction addition to Google Sheets
      syncTransactionToGoogleSheet('add', newTx);
      
      renderAll();
      
      goalOverlay.classList.remove('active');
      showToast(`Deposited ${currencySymbol}${amount} into ${goal.name}.`);
    }
  });

  // --- SETTINGS EVENT LISTENERS ---
  
  // Drawer open/close
  btnOpenSettings.addEventListener('click', () => {
    openOverlay(settingsDrawer);
  });
  btnCloseSettings.addEventListener('click', () => {
    closeOverlay(settingsDrawer);
  });

  // Change Theme Trigger Dropdown
  themeSelect.addEventListener('change', () => {
    applyTheme(themeSelect.value);
  });

  // Change Currency Trigger
  currencySelect.addEventListener('change', () => {
    applyCurrency(currencySelect.value);
  });

  // Check & Install Updates click listener
  const btnUpdateApp = document.getElementById('btn-update-app');
  if (btnUpdateApp) {
    btnUpdateApp.addEventListener('click', () => {
      showToast("Checking for updates on GitHub...");
      if ('serviceWorker' in navigator) {
        navigator.serviceWorker.getRegistrations().then(registrations => {
          if (registrations.length === 0) {
            showToast("Cache refreshed. Reloading...");
            setTimeout(() => {
              window.location.reload(true);
            }, 1000);
            return;
          }
          
          let checked = 0;
          registrations.forEach(reg => {
            reg.update().then(() => {
              checked++;
              if (checked === registrations.length) {
                showToast("Updates installed! Reloading app...");
                setTimeout(() => {
                  window.location.reload(true);
                }, 1200);
              }
            }).catch(err => {
              console.error("SW Update fail:", err);
              window.location.reload(true);
            });
          });
        });
      } else {
        setTimeout(() => {
          window.location.reload(true);
        }, 1000);
      }
    });
  }

  // Toggle Accordions inside settings
  btnToggleAddProfile.addEventListener('click', () => {
    const isHidden = addProfileForm.style.display === 'none' || !addProfileForm.style.display;
    addProfileForm.style.display = isHidden ? 'flex' : 'none';
    btnToggleAddProfile.innerText = isHidden ? '- Hide Form' : '+ Add Family Member';
  });

  btnToggleAddCat.addEventListener('click', () => {
    const isHidden = addCatForm.style.display === 'none' || !addCatForm.style.display;
    addCatForm.style.display = isHidden ? 'flex' : 'none';
    btnToggleAddCat.innerText = isHidden ? '- Hide Form' : '+ Create New Category';
  });

  // Submit Add Family Member Form
  addProfileForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const name = document.getElementById('profile-name-input').value.trim();
    const role = document.getElementById('profile-role-input').value.trim();
    const avatar = document.getElementById('profile-avatar-input').value.trim();
    const color = document.getElementById('profile-color-select').value;
    
    const id = name.toLowerCase().replace(/[^a-z0-9]/g, '-') + '-' + Date.now().toString().slice(-4);
    const glow = `rgba(${hexToRgb(color)}, 0.15)`;
    
    const newMember = { id, name, role, avatar, color, glow };
    members.push(newMember);
    saveToStorage();
    
    // Create sheet for this profile in background
    syncProfileToGoogleSheet('add', name);
    
    populateFormDropdowns();
    renderAll();
    
    addProfileForm.reset();
    addProfileForm.style.display = 'none';
    btnToggleAddProfile.innerText = '+ Add Family Member';
    
    showToast(`Member "${name}" added.`);
  });

  // Submit Add Category Form
  addCatForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const id = document.getElementById('cat-id-input').value.trim().toLowerCase();
    const name = document.getElementById('cat-name-input').value.trim();
    const icon = document.getElementById('cat-icon-input').value.trim();
    const color = document.getElementById('cat-color-input').value;
    const initialBudget = parseFloat(document.getElementById('cat-budget-input').value);
    
    if (categories[id]) {
      alert("Error: Category key already exists.");
      return;
    }
    
    categories[id] = { name, icon, color };
    budgets[id] = isNaN(initialBudget) ? 10000 : initialBudget;
    
    saveToStorage();
    populateFormDropdowns();
    renderAll();
    
    addCatForm.reset();
    addCatForm.style.display = 'none';
    btnToggleAddCat.innerText = '+ Create New Category';
    
    showToast(`Category "${name}" created.`);
  });

  // Toggle Sync Handler
  sheetSyncToggle.addEventListener('change', () => {
    googleSheetSyncEnabled = sheetSyncToggle.checked;
    saveToStorage();
    showToast(`Sheets Sync ${googleSheetSyncEnabled ? 'Enabled' : 'Disabled'}.`);
    
    if (googleSheetSyncEnabled && googleSheetUrl) {
      syncFromGoogleSheets();
    }
  });

  // URL input handler
  sheetUrlInput.addEventListener('input', () => {
    googleSheetUrl = sheetUrlInput.value.trim();
    saveToStorage();
  });

  // Test sync connection
  btnTestSheet.addEventListener('click', testGoogleSheetConnection);

  // Toggle showing deployment code
  btnShowCode.addEventListener('click', () => {
    const isHidden = codeSnippetBox.style.display === 'none';
    codeSnippetBox.style.display = isHidden ? 'block' : 'none';
    btnShowCode.innerText = isHidden ? 'Hide Deployment Code' : 'Show Deployment Code';
  });

  // Toggle showing Sheets Setup Guide
  const btnToggleGuide = document.getElementById('btn-toggle-guide');
  const sheetsGuideBox = document.getElementById('sheets-guide-box');
  if (btnToggleGuide && sheetsGuideBox) {
    btnToggleGuide.addEventListener('click', () => {
      const isHidden = sheetsGuideBox.style.display === 'none';
      sheetsGuideBox.style.display = isHidden ? 'block' : 'none';
      btnToggleGuide.innerText = isHidden ? 'Hide Setup Guide' : 'Show Google Sheets Setup Guide';
    });
  }

  // PIN Lock switch listener
  const pinLockToggle = document.getElementById('pin-lock-toggle');
  if (pinLockToggle) {
    pinLockToggle.addEventListener('change', () => {
      pinLockEnabled = pinLockToggle.checked;
      saveToStorage();
      showToast(`PIN Lock ${pinLockEnabled ? 'Enabled (1234)' : 'Disabled'}.`);
    });
  }

  // Backup CSV Trigger
  const btnExportCSV = document.getElementById('btn-export-csv');
  if (btnExportCSV) {
    btnExportCSV.addEventListener('click', exportCSV);
  }

  // Restore CSV Trigger
  const restoreFileInput = document.getElementById('restore-file-input');
  if (restoreFileInput) {
    restoreFileInput.addEventListener('change', restoreCSV);
  }

  // Modal controls: Category Details
  const catOverlay = document.getElementById('category-details-overlay');
  const btnCloseCatModal = document.getElementById('btn-close-cat-modal');
  
  if (btnCloseCatModal && catOverlay) {
    btnCloseCatModal.addEventListener('click', () => {
      closeOverlay(catOverlay);
    });
    catOverlay.addEventListener('click', (e) => {
      if (e.target === catOverlay) closeOverlay(catOverlay);
    });
  }

  // Modal controls: Edit Family Profile
  const editProfileOverlay = document.getElementById('edit-profile-overlay');
  const btnCloseEditProfileModal = document.getElementById('btn-close-edit-profile-modal');
  const editProfileForm = document.getElementById('edit-profile-form');
  
  if (btnCloseEditProfileModal && editProfileOverlay) {
    btnCloseEditProfileModal.addEventListener('click', () => {
      closeOverlay(editProfileOverlay);
    });
    editProfileOverlay.addEventListener('click', (e) => {
      if (e.target === editProfileOverlay) closeOverlay(editProfileOverlay);
    });
  }
  
  if (editProfileForm) {
    editProfileForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const id = document.getElementById('edit-profile-id').value;
      const m = members.find(x => x.id === id);
      if (m) {
        const oldName = m.name;
        const newName = document.getElementById('edit-profile-name-input').value.trim();
        m.name = newName;
        m.role = document.getElementById('edit-profile-role-input').value.trim();
        m.avatar = document.getElementById('edit-profile-avatar-input').value.trim();
        m.color = document.getElementById('edit-profile-color-select').value;
        m.glow = `rgba(${hexToRgb(m.color)}, 0.15)`;
        
        saveToStorage();
        
        if (oldName !== newName) {
          syncProfileToGoogleSheet('rename', newName, oldName);
        }
        
        populateFormDropdowns();
        renderAll();
        closeOverlay(editProfileOverlay);
        showToast(`Member "${m.name}" updated.`);
      }
    });
  }

  // Modal controls: Edit Savings Goal
  const editGoalOverlay = document.getElementById('edit-goal-overlay');
  const btnCloseEditGoalModal = document.getElementById('btn-close-edit-goal-modal');
  const editGoalForm = document.getElementById('edit-goal-form');
  
  if (btnCloseEditGoalModal && editGoalOverlay) {
    btnCloseEditGoalModal.addEventListener('click', () => {
      closeOverlay(editGoalOverlay);
    });
    editGoalOverlay.addEventListener('click', (e) => {
      if (e.target === editGoalOverlay) closeOverlay(editGoalOverlay);
    });
  }
  
  if (editGoalForm) {
    editGoalForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const id = document.getElementById('edit-goal-id').value;
      const g = goals.find(x => x.id === id);
      if (g) {
        g.name = document.getElementById('edit-goal-name-input').value.trim();
        g.target = parseInt(document.getElementById('edit-goal-target-input').value) || 1000;
        g.current = parseInt(document.getElementById('edit-goal-current-input').value) || 0;
        g.color = document.getElementById('edit-goal-color-select').value;
        
        saveToStorage();
        renderAll();
        closeOverlay(editGoalOverlay);
        showToast(`Goal "${g.name}" updated.`);
      }
    });
  }

  // Event delegation for chart/list clicks
  const chartContainer = document.getElementById('donut-chart-container');
  const legendContainer = document.getElementById('donut-legend');
  
  if (chartContainer) {
    chartContainer.addEventListener('click', handleChartClick);
  }
  if (legendContainer) {
    legendContainer.addEventListener('click', handleChartClick);
  }

  function handleChartClick(e) {
    const item = e.target.closest('.spectrum-segment') || e.target.closest('.spectrum-list-item');
    if (item) {
      const catId = item.getAttribute('data-cat-id');
      if (catId) {
        openCategoryDetails(catId);
      }
    }
  }
}

// --- UTILS ---

// Animated counting number calculations
function animateValue(element, targetVal, isCurrency = false, prefix = '') {
  let startVal = parseFloat(element.innerText.replace(/[₹$,\+\-]/g, '')) || 0;
  if (isNaN(startVal)) startVal = 0;
  
  const duration = 750;
  const startTime = performance.now();
  
  function update(currentTime) {
    const elapsed = currentTime - startTime;
    const progress = Math.min(elapsed / duration, 1);
    
    const easeProgress = progress * (2 - progress);
    const currentVal = startVal + (targetVal - startVal) * easeProgress;
    
    if (isCurrency) {
      const sign = currentVal < 0 ? '-' : '';
      const absVal = Math.abs(currentVal);
      element.innerText = `${sign}${currencySymbol}${absVal.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
    } else {
      element.innerText = `${prefix}${currencySymbol}${Math.round(currentVal).toLocaleString()}`;
    }
    
    if (progress < 1) {
      requestAnimationFrame(update);
    }
  }
  
  requestAnimationFrame(update);
}

function showToast(message) {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerText = message;
  container.appendChild(toast);
  
  setTimeout(() => toast.classList.add('show'), 50);
  
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

function hexToRgb(hex) {
  var shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
  hex = hex.replace(shorthandRegex, function(m, r, g, b) {
    return r + r + g + g + b + b;
  });

  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? 
    `${parseInt(result[1], 16)}, ${parseInt(result[2], 16)}, ${parseInt(result[3], 16)}` 
    : '255, 255, 255';
}

function openCategoryDetails(catId) {
  const cat = categories[catId];
  if (!cat) return;
  
  const overlay = document.getElementById('category-details-overlay');
  const titleName = document.getElementById('cat-modal-name');
  const iconBadge = document.getElementById('cat-modal-icon-badge');
  const summaryBox = document.getElementById('cat-modal-summary-box');
  const txList = document.getElementById('cat-modal-tx-list');
  
  if (!overlay || !titleName || !iconBadge || !summaryBox || !txList) return;
  
  // Setup header
  titleName.innerText = cat.name;
  iconBadge.style.background = cat.color;
  iconBadge.innerHTML = `<i class="fas ${cat.icon}"></i>`;
  
  // Filter category transactions
  const filtered = getFilteredTransactions().filter(t => t.categoryId === catId && t.type === 'expense');
  const spent = filtered.reduce((sum, t) => sum + parseFloat(t.amount), 0);
  const limit = budgets[catId] || 0;
  const remaining = Math.max(limit - spent, 0);
  const percent = limit > 0 ? ((spent / limit) * 100).toFixed(0) : 0;
  
  // Build summary box
  summaryBox.innerHTML = `
    <div style="display: flex; justify-content: space-between;">
      <span>Total Spent:</span>
      <strong>${currencySymbol}${spent.toLocaleString(undefined, {minimumFractionDigits: 2})}</strong>
    </div>
    <div style="display: flex; justify-content: space-between; border-top: 1px dashed var(--apple-border); padding-top: 6px;">
      <span>Monthly Budget:</span>
      <span>${currencySymbol}${limit.toLocaleString()} (${percent}% Used)</span>
    </div>
    <div style="display: flex; justify-content: space-between; border-top: 1px dashed var(--apple-border); padding-top: 6px; font-weight: 600; color: ${remaining === 0 && limit > 0 ? 'var(--apple-red)' : 'var(--apple-text)'};">
      <span>Remaining Budget:</span>
      <span>${currencySymbol}${remaining.toLocaleString()}</span>
    </div>
  `;
  
  // Build transaction items
  if (filtered.length === 0) {
    txList.innerHTML = `
      <div class="empty-state" style="padding: 20px 0;">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
        </svg>
        <p style="font-size: 0.65rem;">No transactions in this category</p>
      </div>
    `;
  } else {
    txList.innerHTML = filtered.map(t => getTransactionRowHTML(t)).join('');
    attachDeleteHandlers();
  }
  
  overlay.classList.add('active');
}

// --- MOBILE OVERLAY HELPERS (History popstate triggers) ---
function openOverlay(el) {
  if (!el) return;
  el.classList.add('active');
  history.pushState({ modalOpen: true }, '');
}

function closeOverlay(el) {
  if (!el) return;
  el.classList.remove('active');
  if (history.state && history.state.modalOpen) {
    history.back();
  }
}

// Global browser popstate listener for back button interception
window.addEventListener('popstate', (e) => {
  const activeOverlay = document.querySelector('.modal-overlay.active, .settings-drawer.active');
  if (activeOverlay) {
    activeOverlay.classList.remove('active');
  }
});

// --- ACCOUNTS BALANCES SLIDER CARDS RENDERING ---
function renderAccountsSlider() {
  const carousel = document.getElementById('accounts-carousel-list');
  if (!carousel) return;
  
  // Base starting balances for each account type
  const accountsBalances = {
    cash: 10000,
    bank: 50000,
    card: -5000,
    upi: 5000,
    savings: 120000
  };
  
  transactions.forEach(t => {
    const acc = t.account || 'cash';
    if (accountsBalances[acc] !== undefined) {
      if (t.type === 'income') {
        accountsBalances[acc] += parseFloat(t.amount);
      } else {
        accountsBalances[acc] -= parseFloat(t.amount);
      }
    }
  });
  
  const accountsMetadata = {
    cash: { name: 'Cash', icon: '💵', color: '#34c759' },
    bank: { name: 'Bank Account', icon: '🏦', color: '#007aff' },
    card: { name: 'Credit Card', icon: '💳', color: '#ff3b30' },
    upi: { name: 'UPI Wallet', icon: '📱', color: '#af52de' },
    savings: { name: 'Savings Account', icon: '💰', color: '#ff9500' }
  };
  
  carousel.innerHTML = Object.keys(accountsMetadata).map(key => {
    const bal = accountsBalances[key];
    const meta = accountsMetadata[key];
    const sign = bal < 0 ? '-' : '';
    const absBal = Math.abs(bal);
    return `
      <div class="account-slide-card" style="border-left: 3.5px solid ${meta.color};">
        <span class="account-slide-icon">${meta.icon}</span>
        <span class="account-slide-name">${meta.name}</span>
        <span class="account-slide-balance">${sign}${currencySymbol}${absBal.toLocaleString(undefined, {maximumFractionDigits: 0})}</span>
      </div>
    `;
  }).join('');
}

// --- CSV BACKUP DATA EXPORTER & PARSER ---
function exportCSV() {
  let csvContent = "data:text/csv;charset=utf-8,";
  csvContent += "Transaction ID,Date,Type,Member ID,Category ID,Amount,Description,Account\n";
  
  transactions.forEach(t => {
    const row = [
      t.id,
      t.date,
      t.type,
      t.memberId,
      t.categoryId,
      t.amount,
      `"${t.description.replace(/"/g, '""')}"`,
      t.account || 'cash'
    ].join(",");
    csvContent += row + "\n";
  });
  
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", `family_ledger_backup_${new Date().toISOString().split('T')[0]}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast("CSV Exported successfully!");
}

function restoreCSV(e) {
  const file = e.target.files[0];
  if (!file) return;
  
  const reader = new FileReader();
  reader.onload = function(evt) {
    const text = evt.target.result;
    const lines = text.split('\n');
    const newTxs = [];
    
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      
      const cols = [];
      let insideQuote = false;
      let current = '';
      for (let c = 0; c < line.length; c++) {
        const char = line[c];
        if (char === '"') {
          insideQuote = !insideQuote;
        } else if (char === ',' && !insideQuote) {
          cols.push(current);
          current = '';
        } else {
          current += char;
        }
      }
      cols.push(current);
      
      if (cols.length >= 7) {
        newTxs.push({
          id: cols[0],
          date: cols[1],
          type: cols[2],
          memberId: cols[3],
          categoryId: cols[4],
          amount: parseFloat(cols[5]) || 0,
          description: cols[6],
          account: cols[7] || 'cash'
        });
      }
    }
    
    if (newTxs.length > 0) {
      transactions = newTxs;
      saveToStorage();
      renderAll();
      showToast(`Restored ${newTxs.length} transactions from CSV!`);
    } else {
      showToast("Invalid CSV Backup File.");
    }
  };
  reader.readAsText(file);
}

// --- PIN LOCK keypads setup ---
function setupPinLockpad() {
  const overlay = document.getElementById('pin-lock-overlay');
  const keys = overlay.querySelectorAll('.pin-key[data-val]');
  const btnClear = document.getElementById('btn-pin-clear');
  const btnDelete = document.getElementById('btn-pin-delete');
  
  enteredPin = '';
  updatePinDots();
  
  keys.forEach(k => {
    k.onclick = () => {
      if (enteredPin.length < 4) {
        enteredPin += k.getAttribute('data-val');
        updatePinDots();
        
        if (enteredPin.length === 4) {
          if (enteredPin === '1234') {
            overlay.style.display = 'none';
            showToast("App unlocked!");
          } else {
            enteredPin = '';
            updatePinDots();
            showToast("Incorrect PIN code! Try again.");
          }
        }
      }
    };
  });
  
  btnClear.onclick = () => {
    enteredPin = '';
    updatePinDots();
  };
  
  btnDelete.onclick = () => {
    enteredPin = enteredPin.slice(0, -1);
    updatePinDots();
  };
}

function updatePinDots() {
  for (let i = 0; i < 4; i++) {
    const dot = document.getElementById(`dot-${i}`);
    if (dot) {
      if (i < enteredPin.length) {
        dot.classList.add('filled');
      } else {
        dot.classList.remove('filled');
      }
    }
  }
}

// Intercept boot check
window.addEventListener('DOMContentLoaded', () => {
  init();
  
  // Trigger lock screen intercept if PIN lock setting is active
  const pinOverlay = document.getElementById('pin-lock-overlay');
  if (pinLockEnabled && pinOverlay) {
    pinOverlay.style.display = 'flex';
    setupPinLockpad();
  }
});
