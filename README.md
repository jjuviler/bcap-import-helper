# BCAP Import Helper

A browser-based tool that converts a Blog Content Action Plan (BCAP) XLSX file into an Asana-importable CSV file.

**Live tool:** https://jjuviler.github.io/bcap-import-helper

---

## What it does

Takes a BCAP workbook with **Marketing, Sales, Service, and Website** tabs and produces:

- An **Import-Ready CSV** formatted for Asana import
- A modified **XLSX** with an Import-Ready tab inserted as the first sheet

### Processing steps

1. Validates that all four required tabs are present
2. Validates that column headings are identical across all tabs
3. Consolidates all data rows into a single Import-Ready tab (Marketing → Sales → Service → Website)
4. Moves **Focus Keyword**, **Writer**, and **Blog Team** to the first three columns
5. Removes columns with no heading
6. Strips whitespace-only cell values
7. Fixes number formatting (removes comma formatting from numeric columns)
8. Standardizes date columns to MM/DD/YYYY text
9. Converts the Assignment Sprint column to plain text

---

## Errors

| Error | Cause |
|---|---|
| File must be an .xlsx file | Non-XLSX file selected |
| None of the required tabs were found | File has no Marketing, Sales, Service, or Website tabs |
| Missing tabs | One or more required tabs are absent — can proceed anyway |
| Heading mismatch | Column headings differ across tabs — fix the file and re-upload |
| Required column(s) not found | One or more required columns are missing from the file entirely |

## Warnings

After a successful run, a warning count is shown. Click to expand the full list. Warnings do not block downloads.

| Warning | Cause |
|---|---|
| Import-Ready tab already existed | Source file had an existing Import-Ready tab — it was replaced |
| Tab has no data rows | A required tab exists but contains no data below the heading row |
| Duplicate Focus Keyword | The same Focus Keyword appears in more than one row |
| Exact duplicate row | A row is identical to a previous row |
| Missing required fields | An active row (one with a Focus Keyword value) is missing one or more required fields |

### Required fields checked per active row

Blog Team · Action · Topic Cluster · Primary Keyword · URL · Content Brief · Timing · Assignee · Content Source · Property · Assignment Sprint · Due Date

---

## Stack

- Vanilla HTML, CSS, and JavaScript — no build step
- [SheetJS](https://sheetjs.com/) for XLSX reading and writing
- Hosted on GitHub Pages
