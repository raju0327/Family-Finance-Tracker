# PowerShell script to seed the Google Sheets spreadsheet for Orbit Finance

$url = "https://script.google.com/macros/s/AKfycbwBgR14Sfj5jvaMqCaJ3tVEdhDjYRn5LgGf8ihBkYeK6ePYUYn6ov_Ax0w6zl0mSPky1A/exec"

Write-Host "Starting database seeding on Google Sheet..."

# 1. Create Sheets for Profiles
$profiles = @("Marcus (Dad)", "Elena (Mom)", "Alex (Teen)")

foreach ($profile in $profiles) {
    Write-Host "Creating sheet tab for operator: $profile..."
    $payload = @{
        action = "createSheet"
        name = $profile
    } | ConvertTo-Json
    
    try {
        $res = Invoke-RestMethod -Uri $url -Method Post -Body $payload -ContentType "application/json"
        Write-Host "Result: $($res | ConvertTo-Json -Compress)"
    } catch {
        Write-Host "Failed to create sheet for $profile. Details: $_"
    }
}

# 2. Add Transactions
$txs = @(
  # Marcus
  @{
    action = "add"
    data = @{
      id = "tx-1"
      date = "2026-07-01"
      memberName = "Marcus (Dad)"
      categoryName = "Salary (In-flow)"
      description = "Quantum Corp Settlement"
      type = "income"
      amount = 85000
    }
  },
  @{
    action = "add"
    data = @{
      id = "tx-3"
      date = "2026-07-02"
      memberName = "Marcus (Dad)"
      categoryName = "Grid Utilities"
      description = "Fusion Grid Charging"
      type = "expense"
      amount = 3500
    }
  },
  @{
    action = "add"
    data = @{
      id = "tx-7"
      date = "2026-07-03"
      memberName = "Marcus (Dad)"
      categoryName = "Investments & Stocks"
      description = "Quantum Nodes purchase"
      type = "expense"
      amount = 15000
    }
  },
  
  # Elena
  @{
    action = "add"
    data = @{
      id = "tx-2"
      date = "2026-07-02"
      memberName = "Elena (Mom)"
      categoryName = "Salary (In-flow)"
      description = "BioTech Lab Grant"
      type = "income"
      amount = 92000
    }
  },
  @{
    action = "add"
    data = @{
      id = "tx-4"
      date = "2026-07-02"
      memberName = "Elena (Mom)"
      categoryName = "Daily Essentials"
      description = "Daily Hydroponics provisions"
      type = "expense"
      amount = 4800
    }
  },
  @{
    action = "add"
    data = @{
      id = "tx-8"
      date = "2026-07-04"
      memberName = "Elena (Mom)"
      categoryName = "Transit & Hyperloop"
      description = "Hyperloop Transit Pass"
      type = "expense"
      amount = 8000
    }
  },
  
  # Alex
  @{
    action = "add"
    data = @{
      id = "tx-5"
      date = "2026-07-03"
      memberName = "Alex (Teen)"
      categoryName = "Daily Essentials"
      description = "Holonet Gaming Log"
      type = "expense"
      amount = 1200
    }
  },
  @{
    action = "add"
    data = @{
      id = "tx-9"
      date = "2026-07-04"
      memberName = "Alex (Teen)"
      categoryName = "Daily Essentials"
      description = "Energy Nutrient Drinks"
      type = "expense"
      amount = 850
    }
  }
)

foreach ($tx in $txs) {
    $desc = $tx.data.description
    Write-Host "Appending transaction record: '$desc'..."
    $payload = $tx | ConvertTo-Json
    
    try {
        $res = Invoke-RestMethod -Uri $url -Method Post -Body $payload -ContentType "application/json"
        Write-Host "Result: $($res | ConvertTo-Json -Compress)"
    } catch {
        Write-Host "Failed to append '$desc'. Details: $_"
    }
}

Write-Host "Seeding complete."
