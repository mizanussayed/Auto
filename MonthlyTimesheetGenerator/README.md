# Monthly Timesheet Generator

This .NET console app copies the supplied Excel template into a new workbook for each month, updates the duration and date row, fills weekdays with `8`, weekends with `SA`/`SU`, marks public holidays as `PH`, colors only weekend and public-holiday date columns gray, and calculates working days. Active day cells use a white fill, and unused trailing-day cells have no border.

Run from the workspace folder:

```powershell
dotnet run --project .\MonthlyTimesheetGenerator
```

By default, it uses the payroll and project templates in the current folder, generates both file types for the current month and next month, and writes them to `.\\generated`. Unused trailing-day cells have no border.

Options:

```powershell
dotnet run --project .\MonthlyTimesheetGenerator -- `
  --template ".\Payroll Timesheet_Mizanur Rahman_ 1 Sep to 30.xlsx" `
  --output ".\generated" `
  --year 2027 `
  --month 2 `
  --public-holidays "16,25" `
  --medical-leave "10" `
  --annual-leave "20,21"
```

`--public-holidays`, `--medical-leave`, and `--annual-leave` accept comma-separated day numbers for the selected month. They write `PH`, `MC`, and `AL` respectively, use gray cells, and are excluded from working-day totals. Public holidays take priority over medical and annual leave; medical leave takes priority over annual leave. Omit `--month` to generate the current and next month; provide `--month` to generate only that month. Generated names use the format `Payroll Timesheet_Mizanur Rahman_ 1 Oct to 31.xlsx`.

By default, one run generates both payroll and project files. Use `--type payroll` or `--type project` to generate only one type. The selected month is the starting month for project timesheets: `--month 10` generates the range 3 October to 2 November and a file named `Project Timesheet_Mizanur Rahman_ 3 Oct to 2 Nov.xlsx`.

```powershell
dotnet run --project .\MonthlyTimesheetGenerator -- `
  --type project `
  --year 2026 `
  --month 10 `
  --public-holidays "16" `
  --medical-leave "10" `
  --annual-leave "20"
```
