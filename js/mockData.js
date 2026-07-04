// Family Finance Tracker - Mock Data (Apple minimalist version)
window.INITIAL_MEMBERS = [
  { id: 'marcus', name: 'Marcus (Dad)', role: 'Father', avatar: '👨', color: '#007aff', glow: 'rgba(0, 122, 255, 0.15)' },
  { id: 'elena', name: 'Elena (Mom)', role: 'Mother', avatar: '👩', color: '#5856d6', glow: 'rgba(88, 86, 214, 0.15)' },
  { id: 'alex', name: 'Alex (Teen)', role: 'Son', avatar: '👦', color: '#ff9500', glow: 'rgba(255, 149, 0, 0.15)' },
  { id: 'zoe', name: 'Zoe (Kid)', role: 'Daughter', avatar: '👧', color: '#34c759', glow: 'rgba(52, 199, 89, 0.15)' }
];

window.CATEGORIES = {
  daily: { name: 'Daily Expenses', icon: 'fa-shopping-basket', color: '#007aff' }, // Apple Blue
  utilities: { name: 'Utilities', icon: 'fa-bolt', color: '#ff9500' }, // Apple Orange
  health: { name: 'Health & Medical', icon: 'fa-heartbeat', color: '#ff3b30' }, // Apple Red
  travel: { name: 'Travel & Transit', icon: 'fa-car', color: '#5856d6' }, // Apple Purple
  investments: { name: 'Investments', icon: 'fa-chart-line', color: '#34c759' }, // Apple Green
  salary: { name: 'Salary', icon: 'fa-wallet', color: '#34c759' } // Apple Green
};

window.INITIAL_TRANSACTIONS = [
  { id: 'tx-1', type: 'income', memberId: 'marcus', categoryId: 'salary', amount: 85000, date: '2026-07-01', description: 'Monthly Salary Payment' },
  { id: 'tx-2', type: 'income', memberId: 'elena', categoryId: 'salary', amount: 92000, date: '2026-07-02', description: 'Business Revenue' },
  { id: 'tx-3', type: 'expense', memberId: 'marcus', categoryId: 'utilities', amount: 3500, date: '2026-07-02', description: 'Electricity & Water Bill' },
  { id: 'tx-4', type: 'expense', memberId: 'elena', categoryId: 'daily', amount: 4800, date: '2026-07-02', description: 'Supermarket Groceries' },
  { id: 'tx-5', type: 'expense', memberId: 'alex', categoryId: 'daily', amount: 1200, date: '2026-07-03', description: 'Mobile Recharges' },
  { id: 'tx-6', type: 'expense', memberId: 'zoe', categoryId: 'health', amount: 2500, date: '2026-07-03', description: 'Dental Clinic Visit' },
  { id: 'tx-7', type: 'expense', memberId: 'marcus', categoryId: 'investments', amount: 15000, date: '2026-07-03', description: 'Mutual Fund Investment' },
  { id: 'tx-8', type: 'expense', memberId: 'elena', categoryId: 'travel', amount: 8000, date: '2026-07-04', description: 'Weekly Fuel Refill' },
  { id: 'tx-9', type: 'expense', memberId: 'alex', categoryId: 'daily', amount: 850, date: '2026-07-04', description: 'Evening Snacks' },
  { id: 'tx-10', type: 'expense', memberId: 'zoe', categoryId: 'daily', amount: 450, date: '2026-07-04', description: 'Books and Stationery' },
  { id: 'tx-11', type: 'expense', memberId: 'marcus', categoryId: 'daily', amount: 3200, date: '2026-07-04', description: 'Family Dinner Outing' },
  { id: 'tx-12', type: 'expense', memberId: 'elena', categoryId: 'utilities', amount: 1500, date: '2026-07-04', description: 'Broadband Net Bill' }
];

window.INITIAL_BUDGETS = {
  daily: 15000,
  utilities: 10000,
  health: 12000,
  travel: 20000,
  investments: 50000
};

window.INITIAL_GOALS = [
  { id: 'goal-1', name: 'Family Vacation Fund', target: 250000, current: 185000, color: '#007aff' },
  { id: 'goal-2', name: 'New Living Room TV', target: 80000, current: 42000, color: '#ff9500' },
  { id: 'goal-3', name: 'Emergency Savings', target: 150000, current: 30000, color: '#34c759' }
];
