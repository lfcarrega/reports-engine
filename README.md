# Reports Engine

> ⚠️ **Work in Progress** — functional but still under active development.

A lightweight desktop tool for generating web-based reports from SQL Server databases. Write a `.sql` file, drop it in the reports folder, and the tool executes it, renders the results in an interactive table, and lets you export to multiple formats — no web server setup required.

---

## How it works

1. The app starts a local [Caddy](https://caddyserver.com/) file server on `http://127.0.0.1:8080`
2. You select a report from the dropdown (any `.sql` or `.txt` file found in the `reports/` folder)
3. Optionally set a start/end date range
4. The query is executed against SQL Server via ADO, results are serialized to `data.json`
5. The browser opens with a [Tabulator](https://tabulator.info/) table loaded from that JSON

---

## Requirements

- Windows (uses AutoHotkey v2 + ActiveX/ADO)
- [AutoHotkey v2](https://www.autohotkey.com/)
- [SQL Server LocalDB 2022](https://download.microsoft.com/download/3/8/d/38de7036-2433-4207-8eae-06e247e17b25/SqlLocalDB.msi) — or any SQL Server instance
- [MSOLEDBSQL driver](https://learn.microsoft.com/en-us/sql/connect/oledb/download-oledb-driver-for-sql-server)

---

## Getting started

### 1. Set up the database

Edit the connection string near the top of `main.ahk` to point to your instance:

```ahk
ConnectionString := "Provider=MSOLEDBSQL;Data Source=(localdb)\MSSQLLocalDB;AttachDbFileName=" A_WorkingDir "\MyDatabase.mdf;Initial Catalog=MyDatabase;Integrated Security=SSPI;"
```

If you want a demo database to test with, run `setup_reports_db.sql` against your LocalDB instance:

```bat
sqlcmd -S "(localdb)\MSSQLLocalDB" -i setup_reports_db.sql
```

### 2. Create a report

Create a `.sql` file inside the `reports/` folder (you can also use `relatorios/` or `relatorios/`):

```sql
-- reports/sales_by_date.sql
SELECT
    v.id        AS Order_ID,
    c.nome      AS Customer,
    v.data_venda AS Date,
    v.total     AS Total,
    v.status    AS Status
FROM venda v
JOIN cliente c ON c.id = v.id_cliente
WHERE v.data_venda BETWEEN '{{DATAINICIAL}}' AND '{{DATAFINAL}}'
ORDER BY v.data_venda DESC
```

### 3. (Optional) Add metadata

Create a `.meta` file with the same name as your `.sql` to show friendly info in the UI:

```
title: Sales by Date
description: Lists all orders filtered by date range.
author: Your Name
```

### 4. Run

```bat
main.ahk
```

Select your report, set the date range if needed, and hit `→`.

---

## Supported placeholders

| Placeholder | Replaced with |
|---|---|
| `{{DATAINICIAL}}` | Start date (`yyyy-MM-dd 00:00:00`) or `00:00:00` if unchecked |
| `{{DATAFINAL}}` | End date (`yyyy-MM-dd 00:00:00`) or `00:00:00` if unchecked |
| `{{CODLOJA}}` | Store/branch code from the logged-in user |
| `{{PRIV}}` | User privilege level |
| `{{LOGIN}}` | Logged-in username |

---

## Customizing the output

By default the report renders a standard Tabulator table with download buttons. You can override the HTML, JS, and CSS per-report by placing files with the same name as your `.sql` alongside it:

| File | Overrides |
|---|---|
| `report_name.html` | The body HTML |
| `report_name.js` | The Tabulator initialization script |
| `report_name.css` | Additional styles |

Global defaults can be set with `__default.html`, `__default.js`, and `__default.css` in the root folder.

---

## Export formats

The default interface includes buttons to download the current table as:

- CSV
- JSON
- XLSX
- PDF (landscape)
- HTML (with inline styles)

---

## Theming

The app ships with all [Tabulator themes](https://tabulator.info/docs/6.3/theme). Pick one from the dropdown and star it (★) to save it as your default.

---

## Project structure

```
├── main.ahk                  # Entry point
├── adosql.ahk                # ADO query helper
├── reports/                  # Your .sql report files go here
│   ├── my_report.sql
│   ├── my_report.meta        # Optional metadata
│   ├── my_report.js          # Optional custom JS
│   └── my_report.css         # Optional custom CSS
├── srv/
│   ├── js/                   # Tabulator, jsPDF, SheetJS
│   ├── css/                  # Tabulator themes
│   ├── data.json             # Generated at runtime
│   └── index.html            # Generated at runtime
├── bin/win/amd64/caddy.exe   # Embedded file server
└── config.ini                # Saved theme preference
```

---

## Roadmap

- [ ] Custom parameter detection from query (auto-generate input fields based on `{{PARAM}}` tokens found in the `.sql`)
- [ ] Login / user management (optional, for multi-user deployments)
- [ ] Configurable connection string via UI
- [ ] Support for multiple database connections

---

## Dependencies & credits

- [Tabulator](https://tabulator.info/) — interactive tables
- [jsPDF](https://github.com/parallax/jsPDF) + [AutoTable](https://github.com/simonbengtsson/jsPDF-AutoTable) — PDF export
- [SheetJS](https://sheetjs.com/) — XLSX export
- [Caddy Server](https://caddyserver.com/) — embedded local file server
