Expense Dashboard Automation (Excel VBA)

Excel VBA project that turns a raw expenses data sheet into a fully built dashboard — pivot tables, KPI cards, charts, and slicers — with a single macro call.

<img width="1760" height="679" alt="image" src="https://github.com/user-attachments/assets/bc5ea665-a569-44f1-8dab-aa39233e70da" />


What it does

Running `BuildExpensesDashboard()` does three things in sequence:

1. PrepareDataSheet – cleans and formats the raw `Data` sheet into a proper Excel Table (`tblExpenses`)
2. BuildPivotSheet – builds 7 pivot tables (expense by department, by category, monthly trend, approval status, KPI summary, payment method, budget vs actual)
3. BuildDashboardSheet – assembles the final dashboard: header, KPI cards, charts, and interactive slicers (Month, Department)

A status-bar progress indicator runs through each step so it's clear what stage the build is at.

 Why I built it this way

Manually rebuilding a pivot-based expense dashboard every time new data comes in is repetitive and error-prone. This automates the whole thing into one button click, and keeps the layout consistent every time it's rebuilt.

 Structure

Module1.bas – all macro logic (helpers, pivot builders, dashboard builder, chart formatter, utilities)
Sheet1.cls ThisWorkbook.cls` – sheet/workbook event hooks (currently unused, kept for future extension)

How to run

1. Open the workbook, enable macros
2. Alt+F8 → run `BuildExpensesDashboard`
3. Dashboard sheet gets rebuilt from the current `Data` sheet contents

