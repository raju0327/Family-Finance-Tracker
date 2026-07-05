// Family Finance Tracker - Mock Data (Apple minimalist version)
window.INITIAL_MEMBERS = [
  { id: 'marcus', name: 'Marcus (Dad)', role: 'Father', avatar: '👨', color: '#007aff', glow: 'rgba(0, 122, 255, 0.15)' },
  { id: 'elena', name: 'Elena (Mom)', role: 'Mother', avatar: '👩', color: '#5856d6', glow: 'rgba(88, 86, 214, 0.15)' },
  { id: 'alex', name: 'Alex (Teen)', role: 'Son', avatar: '👦', color: '#ff9500', glow: 'rgba(255, 149, 0, 0.15)' },
  { id: 'zoe', name: 'Zoe (Kid)', role: 'Daughter', avatar: '👧', color: '#34c759', glow: 'rgba(52, 199, 89, 0.15)' }
];

window.CATEGORIES = {
  // Expense Categories
  food: { name: 'Food & Groceries', icon: 'fa-shopping-basket', color: '#007aff', type: 'expense' },
  transport: { name: 'Transport & Fuel', icon: 'fa-car', color: '#5856d6', type: 'expense' },
  rent: { name: 'Rent & Living', icon: 'fa-home', color: '#ff9500', type: 'expense' },
  shopping: { name: 'Shopping', icon: 'fa-shopping-bag', color: '#ff2d55', type: 'expense' },
  bills: { name: 'Bills & Utilities', icon: 'fa-bolt', color: '#ffcc00', type: 'expense' },
  emi: { name: 'EMI & Loans', icon: 'fa-credit-card', color: '#ff3b30', type: 'expense' },
  entertainment: { name: 'Entertainment', icon: 'fa-play-circle', color: '#af52de', type: 'expense' },
  health: { name: 'Health & Medical', icon: 'fa-heartbeat', color: '#34c759', type: 'expense' },
  education: { name: 'Education', icon: 'fa-graduation-cap', color: '#00c7e2', type: 'expense' },
  investments: { name: 'Investments', icon: 'fa-chart-line', color: '#a2a2a2', type: 'expense' },
  
  // Income Categories
  salary: { name: 'Salary', icon: 'fa-wallet', color: '#34c759', type: 'income' },
  business: { name: 'Business Income', icon: 'fa-briefcase', color: '#34c759', type: 'income' },
  freelance: { name: 'Freelance Work', icon: 'fa-laptop-code', color: '#34c759', type: 'income' },
  'investments-income': { name: 'Investments Return', icon: 'fa-chart-pie', color: '#34c759', type: 'income' },
  'gifts-income': { name: 'Gifts & Awards', icon: 'fa-gift', color: '#34c759', type: 'income' }
};

window.INITIAL_TRANSACTIONS = [
  { id: 'tx-1', type: 'income', memberId: 'marcus', categoryId: 'salary', amount: 85000, date: '2026-07-01', description: 'Monthly Corporate Salary', account: 'bank' },
  { id: 'tx-2', type: 'income', memberId: 'elena', categoryId: 'business', amount: 92000, date: '2026-07-02', description: 'E-commerce Shop Sales', account: 'bank' },
  { id: 'tx-3', type: 'expense', memberId: 'marcus', categoryId: 'bills', amount: 3500, date: '2026-07-02', description: 'Electricity & Water Bill', account: 'upi' },
  { id: 'tx-4', type: 'expense', memberId: 'elena', categoryId: 'food', amount: 4800, date: '2026-07-02', description: 'Weekly Groceries Store', account: 'cash' },
  { id: 'tx-5', type: 'expense', memberId: 'alex', categoryId: 'bills', amount: 1200, date: '2026-07-03', description: 'Mobile Recharge Pack', account: 'upi' },
  { id: 'tx-6', type: 'expense', memberId: 'zoe', categoryId: 'health', amount: 2500, date: '2026-07-03', description: 'Pediatrician Checkup', account: 'card' },
  { id: 'tx-7', type: 'expense', memberId: 'marcus', categoryId: 'investments', amount: 15000, date: '2026-07-03', description: 'Index SIP Investment', account: 'savings' },
  { id: 'tx-8', type: 'expense', memberId: 'elena', categoryId: 'transport', amount: 8000, date: '2026-07-04', description: 'SUV Fuel Top-up', account: 'card' },
  { id: 'tx-9', type: 'expense', memberId: 'alex', categoryId: 'food', amount: 850, date: '2026-07-04', description: 'Burgers & Fries Delivery', account: 'upi' },
  { id: 'tx-10', type: 'expense', memberId: 'zoe', categoryId: 'education', amount: 450, date: '2026-07-04', description: 'Drawing Books & Crayons', account: 'cash' }
];

window.INITIAL_BUDGETS = {
  food: 15000,
  transport: 12000,
  rent: 30000,
  shopping: 10000,
  bills: 15000,
  emi: 20000,
  entertainment: 8000,
  health: 10000,
  education: 15000,
  investments: 25000
};

window.INITIAL_GOALS = [
  { id: 'goal-1', name: 'Family Vacation Fund', target: 250000, current: 185000, color: '#007aff' },
  { id: 'goal-2', name: 'New Living Room TV', target: 80000, current: 42000, color: '#ff9500' },
  { id: 'goal-3', name: 'Emergency Savings', target: 150000, current: 30000, color: '#34c759' }
];
