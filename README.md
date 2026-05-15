# MCA Clients Import

This project contains the SQL and Python utility for importing shareholder /
debenture-holder data from company folders into Oracle tables.

## Files

- `sql/create_mca_client_tables.sql` adds the `COMPANIES.SEQ_NO` primary key and creates `FILES` and `CLIENTS`.
- `scripts/import_clients.py` scans company folders, extracts zip files, validates Excel workbooks, imports client rows, and writes a JSON skipped-file report.
- `requirements.txt` lists the Python dependencies.

## Install Dependencies

```powershell
pip install -r requirements.txt
```

## Run SQL

Run `sql/create_mca_client_tables.sql` against the Oracle schema that already has the `COMPANIES` table.

If `COMPANIES.SEQ_NO` is already a primary key, skip the first `ALTER TABLE` statement.

## Run Importer

```powershell
python scripts\import_clients.py "D:\path\to\root-folder" `
  --db-user "username" `
  --db-password "password" `
  --db-dsn "localhost:1521/FREEPDB1" `
  --company-name-column "COMPANY_NAME"
```

The company folder name is matched against `COMPANIES.<company-name-column>`.
If your column is named differently, pass it with `--company-name-column`.

The script creates:

- `import_clients.log`
- `skipped_files_report.json`
