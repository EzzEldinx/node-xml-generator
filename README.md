# Numismatic XML Generator

## Overview

This Node.js project automates the transformation of numismatic data from an Excel spreadsheet (`Sample_data.xlsx`) into structured XML files, one per row, using a template (`script.xml`). The script is tailored for the Centre d'Études Alexandrines (CEAlex) and robustly fetches, processes, and outputs data for coin records, including external lookups, special XML formatting, and comprehensive error handling.

---

## Features

- **Reads Excel Data:** Extracts data from each row of the provided Excel file, skipping header rows.
- **Template-Based XML Generation:** Uses `script.xml` as a base, updating fields with row-specific data.
- **External Data Fetching:** Fetches additional labels and titles from URLs (e.g., nomisma.org) with retry logic and timeouts.
- **Special XML Handling:**
  - Escapes "corne d'abondance" as `corne d&apos;abondance` in all XML output.
  - Ensures copyright holder uses a literal apostrophe.
  - Formats `<date>` tags with both ISO 8601 and human-readable text, and inserts a comment.
  - Handles missing or empty fields by commenting out the corresponding XML tags.
  - Supports new tags: `<symbol>`, `<hoard>`, `<fallsWithin>`, and improved handling for `<department>`, `<repository>`, `<fileLocation>`, and more.
- **Comprehensive Debugging:** Logs detailed information for each processed row and field.
- **Output:** Writes a separate XML file for each valid row in the `output/` directory, named by the row's index.
- **Reporting:** At the end, prints a report showing which fields were filled or left empty for each row.
- **Post-processing:** Includes a script (`update-denominations.js`) to update denomination names in output XMLs by fetching the latest labels from the web.

---

## File Structure

- `Sample_data.xlsx`: Source Excel file with numismatic data.
- `script.xml`: XML template file.
- `output/`: Directory where generated XML files are saved.
- `process_excel.js`: Main script for processing and generating XML files.
- `update-denominations.js`: Script to update denomination names in output XMLs.
- `check_excel.js`: Utility to inspect and debug Excel columns.
- `process-xml.js`, `script.js`: Alternative or legacy scripts for XML processing.

---

## How It Works

### 1. Initialization

- Loads required modules: `xlsx`, `fs`, `path`, `axios`, `cheerio`, `fast-xml-parser`, `fs-extra`, `xml2js`.
- Reads the Excel file and the XML template.

### 2. Row Processing

For each data row (skipping headers):

- **Extracts values** from specific columns (see mapping below).
- **Fetches external data** (labels/titles) from URLs using robust retry logic.
- **Updates the XML template** with row-specific data:
  - Fills in or comments out tags as needed.
  - Handles special cases (e.g., escaping, copyright).
  - Formats and inserts date tags with comments.
  - Handles new tags: `<symbol>`, `<hoard>`, `<fallsWithin>`, and more.
- **Writes the resulting XML** to `output/{index}.xml`.

### 3. Reporting

- Collects a report for each row, indicating which fields were filled or empty.
- Prints a summary report to the console after processing all rows.

---

## Key Functions

- `fetchWithRetry(url, options, retries, delayMs)`: Fetches a URL with retries and timeout.
- `getMintLabel(url)`, `getReferenceTitle(url)`: Extracts the main heading from a mint/material/department/reference URL.
- `processExcelFile()`: Main function orchestrating the reading, processing, and writing logic.
- `updateDenominations()`: Updates denomination names in output XMLs by fetching the latest labels from the web.

---

## Column Mapping (Excel → XML)

| Excel Column | XML Tag/Field                | Notes                                      |
|--------------|------------------------------|--------------------------------------------|
| A            | `<recordId>`                 | Used for file naming and recordId          |
| B            | `<title>`                    | Coin title                                 |
| C, D         | `<fromDate>`, `<toDate>`     | Date range                                 |
| F, AJ        | `<denomination>`             | URL and name                               |
| G, AK        | `<material>`                 | URL and name                               |
| H, AL        | `<authority>`                | URL and name                               |
| J            | `<typeSeries>`               | Optional                                   |
| K, AM        | `<geogname role="mint">`    | Mint URL and name                          |
| L, M         | `<obverse>`                  | Description and legend                     |
| N, O         | `<reverse>`                  | Description and legend                     |
| P, AN        | `<symbol>`                   | Symbol URL and name                        |
| Q            | `<axis>`                     | Physical axis                              |
| R            | `<weight>`                   | Physical weight                            |
| S            | `<diameter>`                 | Physical diameter                          |
| T            | `<countermark>`              | Optional                                   |
| U, V, W      | `<reference>`, `<reference xlink:href>`, `<tei:idno>` | Reference info, URL, and CEAlex idno |
| X            | `<geogname role="stratigraphicUnit">` | Stratigraphic unit                  |
| Y, AH        | `<fallsWithin>`              | Findspot URL and name                      |
| Z, AO        | `<hoard>`                    | Hoard URL and name                         |
| AA           | `<date>`                     | ISO and readable, with comment             |
| AC           | `<identifier localType="Coin Number">` | Coin number                    |
| AD           | `<identifier localType="Inventory Number">` | Inventory number                |
| AE, AP       | `<department>`               | Department URL and name                    |
| AF           | `<repository>`               | Repository name                            |
| AG           | `<mets:FLocat>`              | File location URL                          |

---

## Special Logic

- **Escaping:** All instances of "corne d'abondance" are replaced with `corne d&apos;abondance` in XML.
- **Copyright:** The copyright holder is always set to `Centre d'Études Alexandrines`.
- **Date Tag:** If a date is present in column AA, it is formatted as both ISO 8601 and readable text, with a comment above the tag.
- **External Fetching:** If a label/title cannot be fetched after retries, the tag is left empty but the URL is still included.
- **Commenting:** If a field is missing or empty, the corresponding XML tag is commented out in the output.
- **New Tags:** Supports `<symbol>`, `<hoard>`, `<fallsWithin>`, and improved handling for all admin and findspot tags.

---

## Running the Script

1. Ensure all dependencies are installed (`npm install`).
2. Place `Sample_data.xlsx` and `script.xml` in the project root.
3. Run the main script:
   ```bash
   node process_excel.js
   ```
4. (Optional) Update denomination names in output XMLs:
   ```bash
   npm run update-denominations
   ```
5. Check the `output/` directory for generated XML files.
6. Review the console for the fetch report and any debug information.

---

## Troubleshooting

- **No Output Files:** Ensure the Excel file and template exist and are readable. Check for errors in the console.
- **Network Issues:** The script retries failed HTTP requests, but persistent network issues may leave some fields empty.
- **Template Changes:** If the XML structure changes, update `script.xml` and adjust the script's field mappings as needed.
- **Column Changes:** If the Excel file changes, update the column mapping in `process_excel.js` accordingly.

---

## Customization

- To process only a specific row, modify the row loop in `processExcelFile()`.
- To add or change field mappings, update the column extraction and XML update logic accordingly.
- To add new post-processing steps, create or modify scripts as needed (see `update-denominations.js`). 