// Node.js script to seed the Google Sheets spreadsheet for Family Finance Tracker
// Usage: node seed.js

const url = "https://script.google.com/macros/s/AKfycbwBgR14Sfj5jvaMqCaJ3tVEdhDjYRn5LgGf8ihBkYeK6ePYUYn6ov_Ax0w6zl0mSPky1A/exec";

console.log("Starting database seeding on Google Sheet...");

const profiles = ["Marcus (Dad)", "Elena (Mom)", "Alex (Teen)"];

async function runSeeding() {
  // 1. Create Sheets for Profiles
  for (const profile of profiles) {
    console.log(`Creating sheet tab for operator: ${profile}...`);
    const payload = {
      action: "createSheet",
      name: profile
    };
    
    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      console.log("Result:", data);
    } catch (err) {
      console.error(`Failed to create sheet for ${profile}. Details:`, err.message);
    }
  }

  // 2. Add Transactions
  const txs = [
    {
      action: "add",
      data: {
        id: "tx-1",
        date: "2026-07-01",
        memberName: "Marcus (Dad)",
        categoryName: "Salary (In-flow)",
        description: "Quantum Corp Settlement",
        type: "income",
        amount: 85000
      }
    },
    {
      action: "add",
      data: {
        id: "tx-3",
        date: "2026-07-02",
        memberName: "Marcus (Dad)",
        categoryName: "Grid Utilities",
        description: "Fusion Grid Charging",
        type: "expense",
        amount: 3500
      }
    },
    {
      action: "add",
      data: {
        id: "tx-7",
        date: "2026-07-03",
        memberName: "Marcus (Dad)",
        categoryName: "Investments & Stocks",
        description: "Quantum Nodes purchase",
        type: "expense",
        amount: 15000
      }
    },
    {
      action: "add",
      data: {
        id: "tx-2",
        date: "2026-07-02",
        memberName: "Elena (Mom)",
        categoryName: "Salary (In-flow)",
        description: "BioTech Lab Grant",
        type: "income",
        amount: 92000
      }
    },
    {
      action: "add",
      data: {
        id: "tx-4",
        date: "2026-07-02",
        memberName: "Elena (Mom)",
        categoryName: "Daily Essentials",
        description: "Daily Hydroponics provisions",
        type: "expense",
        amount: 4800
      }
    },
    {
      action: "add",
      data: {
        id: "tx-8",
        date: "2026-07-04",
        memberName: "Elena (Mom)",
        categoryName: "Transit & Hyperloop",
        description: "Hyperloop Transit Pass",
        type: "expense",
        amount: 8000
      }
    },
    {
      action: "add",
      data: {
        id: "tx-5",
        date: "2026-07-03",
        memberName: "Alex (Teen)",
        categoryName: "Daily Essentials",
        description: "Holonet Gaming Log",
        type: "expense",
        amount: 1200
      }
    },
    {
      action: "add",
      data: {
        id: "tx-9",
        date: "2026-07-04",
        memberName: "Alex (Teen)",
        categoryName: "Daily Essentials",
        description: "Energy Nutrient Drinks",
        type: "expense",
        amount: 850
      }
    }
  ];

  for (const tx of txs) {
    console.log(`Appending transaction record: '${tx.data.description}'...`);
    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(tx)
      });
      const data = await res.json();
      console.log("Result:", data);
    } catch (err) {
      console.error(`Failed to append '${tx.data.description}'. Details:`, err.message);
    }
  }

  console.log("Seeding complete.");
}

runSeeding();
