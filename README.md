# Flash Test Report Manager (LONGi sample format)

## Supported sample
`LS-SIS-20260227_FR.xlsx`: 5 worksheets, header rows 1–6, serial in column A, electrical values in B:J starting at row 7. The workbook is saved **unaltered** to SQLite; only S/N row indexes are extracted. Output modifies the original XLSX (OOXML), retaining its shared styles, headers, merged cells, widths, and sheet names. The original summaries in rows 3–5 are recalculated for the selected serials.

## Install and run (Windows PowerShell)

```powershell
cd flash_report_app
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Open the local URL printed by Streamlit (usually http://localhost:8501).

## Workflow
1. Upload one or more original source XLSX files and click **บันทึกเข้าฐานข้อมูล**.
2. Select source files. Paste S/N values or upload `.xlsx` (column A), `.csv`, `.txt`.
3. Search. Inspect matches/unmatched S/N. Download one formatted Excel per original source file.
4. The database is stored under `data/flash_reports.sqlite3` by default. To relocate, set `FLASH_DB_PATH` in your environment.

## Notes / limits
- Currently tailored to the supplied vendor file layout, not any arbitrary Flash Test workbook. For a different supplier, add a parser/adapter.
- Sources and serial indexes persist between sessions on a persistent disk; ephemeral Streamlit hosting can reset local storage. Use persistent volume or managed DB for production.
- A repeated S/N in different sheets is included on all matching sheets, and a repeated S/N across files is exported separately per file.
- Exact original **cell styling/structure** is retained. New row counts alter pagination by design; printer drivers may change actual PDF output.
- The source workbook is never overwritten; output is newly generated.
- Excel reporting is implemented; PDF is intentionally not offered until a reliable Excel-to-PDF renderer is available (e.g. server-side Microsoft Excel automation) since PDF conversion may alter layout.
- For multi-user/public deployment, add login, upload limits, malware scanning, audit records, and backups; current build is for trusted internal users.
- Excel formulas or defined names targeting original row numbers are not supported in this template-specific implementation. The provided sample does not use defined print areas.

## การนำเข้า Flash Report หลายไฟล์พร้อมกัน

1. เปิดแท็บ **1. นำเข้าฐานข้อมูล**
2. กด Browse files แล้วเลือกไฟล์ `.xlsx` หลายไฟล์พร้อมกัน (กด Ctrl หรือ Shift เพื่อเลือกหลายรายการ)
3. กด **บันทึกไฟล์ทั้งหมดเข้าฐานข้อมูล**
4. ระบบแสดงสถานะของแต่ละไฟล์, Source ID, จำนวน S/N และจำนวนชีต

การนำเข้าแต่ละไฟล์แยกกัน หากไฟล์หนึ่งไม่ผ่านการตรวจสอบ ไฟล์อื่นจะยังถูกนำเข้าต่อไปตามปกติ
**หมายเหตุ:** หากกดนำเข้าไฟล์เดิมซ้ำ ระบบจะสร้าง Source ID ใหม่และอาจเกิด S/N ซ้ำระหว่างแหล่งข้อมูลได้ ควรตรวจสอบที่แท็บจัดการฐานข้อมูล
