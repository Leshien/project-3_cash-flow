
# Cash Flow Optimization Simulator

A dynamic, macro-enabled Excel tool that models how changes in working capital affect Free Cash Flow (FCF) — designed for FP&A analysts, controllers, and financial business partners.



##  Key Features

###  Realistic Financial Simulation
- Simulates how **Receivables Days**, **Inventory Days**, and **Payables Days** impact Net Working Capital (NWC) and FCF
- Uses a full 12-month synthetic dataset of sales, COGS, EBITDA, CapEx, and cash cycle metrics

###  Built-In Scenario Testing
- **Apply Random Shock**: Adds ±5-day volatility to simulate market impact
- **Reset to Baseline**: Returns assumptions to original dataset values
- **Run 100 Simulations**: Runs randomized working capital variations and logs total FCF across all runs

###  Analytical Summary
- Auto-updating summary dashboard shows:
  - Average FCF
  - Max & Min FCF across runs
  - Standard Deviation
- Ready for recruiter, hiring manager, or executive demo

---

##  File Contents

| `Cash_Flow_Optimization_Simulator.xlsm` - Main macro-enabled workbook
| `CashFlowSimulator.bas` - VBA macro module (import into Excel VBA Editor)

---

##  Why It Matters

This tool showcases:
- Strategic thinking in **cash cycle optimization**
- Technical capability with **Excel, Power Query, and VBA**
- Ability to model and communicate **business impact under uncertainty**
- The mindset of a **financial partner, not just a reporter**

---

## How to Use

1. Open the `.xlsm` file in Excel
2. Press `ALT + F11` → `File → Import` → load the `.bas` file
3. Use the buttons in the `Simulator` sheet
4. View results in the `Summary` sheet or explore hidden logs in `Simulation Runs`

