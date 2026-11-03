<div align="center">
  <h1>Variable Data Studio</h1>
  <p><strong>Automate. Personalize. Generate.</strong></p>
  <p>A variable-data automation script for Adobe Illustrator.</p>
</div>
<img width="6967" height="3921" alt="01-variable-data-studio-cover" src="https://github.com/user-attachments/assets/8374e8ac-ce40-443f-b305-ad88a474251b" />
---

## Overview

**Variable Data Studio** is an Adobe Illustrator JSX script that helps designers create multiple personalized versions of one artwork from a structured data list.

Instead of manually editing a name, number, date, or other text field and duplicating the artwork repeatedly, you map live text frames in your template to custom variable names, enter or import records, choose an output layout, and let the script generate the variations.

It can be used for certificates, badges, ID cards, invitations, labels, tickets, product tags, numbered designs, and other layouts that contain changing text.

## Features

- **Manual text mapping:** Choose the exact live text frame for each variable. Give variables meaningful names such as `Full Name`, `ID`, `Date`, or `Code`.
- **Editable data grid:** Add and edit records directly in the second-stage table.
- **Spreadsheet input:** Paste cells from Excel or import supported workbook/data files.
- **Excel workbook import:** Import `.xlsx`, `.xls`, and `.xlsm` workbooks through Microsoft Excel automation on Windows with Microsoft Excel installed.
- **Text-file import:** Read CSV, TSV, and TXT files, with column-mapping support when source headers do not match variable names.
- **Large data lists:** Work with paginated records (up to 50 records per page), vertical scrolling, and horizontal scrolling for additional variables.
- **Flexible deletion:** Delete selected records or variables using individual indices and ranges, such as `2,4` or `2-15`.
- **Five output layouts:** Separate Artboards, Grid, Horizontal, Vertical, and Same Position.
- **Artboard arrangement:** Arrange Separate Artboards in a grid, one horizontal row, or one vertical column, with configurable spacing.
- **Template handling:** Keep the original template, delete it after generation, or replace it with the first generated output.
- **Blank-cell behavior:** Clear the mapped text when a cell is empty, or keep the template text.
- **Optional artwork flattening:** Rasterize static artwork while keeping mapped variable text editable. Available rasterization settings include 150, 300, and 600 ppi.
- **Preview:** Preview the first record before generating the full set.
- **PDF output:** Export all records into one multi-page PDF from the data stage, or export generated outputs as separately numbered PDF files after generation.
- **Reusable presets:** Save and load `.vds` presets containing mappings, configuration, and records. Legacy `.cbg` presets are supported for import.
- **Back to Setup:** Return from the data stage to edit the variable mapping and layout settings while keeping the entered records.
- **Temporary generation workflow:** Generated artwork is added to the open Illustrator document. The document is not saved automatically.

## Workflow

<img width="1672" height="941" alt="02-workflow-overview" src="https://github.com/user-attachments/assets/f02e7023-f13f-4ba8-9316-0bc6b77439ea" />

1. Open an Illustrator document and select the complete grouped artwork.
2. Add variables and map each one to a live text frame.
3. Enter records in the data grid, paste spreadsheet cells, or import a supported file.
4. Choose how the generated artwork should be arranged.
5. Preview the first record if needed, export a PDF, and/or finish generation.

## Output Layouts

<img width="1672" height="941" alt="03-layout-options" src="https://github.com/user-attachments/assets/d9d23cd9-275c-435b-b94f-edcb2e699295" />

| Layout | Result |
|---|---|
| **Separate Artboards** | Creates one artboard per output. Choose Grid, Horizontal only, or Vertical only arrangement and set the spacing. |
| **Grid** | Places all generated items in a grid on the existing document canvas. |
| **Horizontal** | Places generated items in one horizontal row. |
| **Vertical** | Places generated items in one vertical column. |
| **Same Position** | Places all generated items at the same position, stacked on top of one another. |

For Grid, Horizontal, Vertical, and Same Position, the existing artboard dimensions are kept unchanged by the generation workflow. Separate Artboards use the dimensions of the source artboard.

## Installation


1. Download `Variable_Data_Studio_v1.jsx` file
2. Cut or Copy The File & Paste it in Adobe illustrator Scripts file `example : C:\Program Files\Adobe\Adobe Illustrator\Presets\en\Scripts`
3. Restart Adobe Illustrator
4. Done!

## How to Use

### Stage 1 — Setup

1. Select the complete source artwork Group before launching the script. If you are restoring a saved setup, use **Load Preset...** and make sure its template backup is available when required.
2. In **Variable Mapping**, add a variable for every text field that should change.
3. Rename variables to descriptive labels, for example `Full Name`, `Reference`, `Issue Date`, or `Serial Number`.
4. For each variable, select its corresponding live text frame from the dropdown list. A text frame should not be assigned to more than one variable.
5. Select an output layout and configure the layout-specific settings. For Separate Artboards, choose Grid, Horizontal only, or Vertical only and set the spacing in centimeters.
6. Choose how the original template should be handled: **Keep Original**, **Delete Original**, or **Replace Original**.
7. Optionally enable **Flatten static artwork; keep mapped text editable** if the template contains complex static artwork and you want to reduce the amount of editable artwork in each generated copy.
8. Continue to the data stage.

### Stage 2 — Record Data

1. Add records with **Add Row +** and type directly into the grid.
2. To paste spreadsheet data, copy a rectangular range in Excel and use **Paste from Excel...**.
3. To import a file, choose **Import File...**. Select a supported Excel workbook or a CSV/TSV/TXT file. For Excel workbooks, select the worksheet containing your data.
4. Review the column-mapping dialog if the source headings do not match your variable names. Confirm whether the first row contains headers and verify the source-column assignments before importing.
5. Use **Delete Rows...** to remove records. Examples: `2`, `2-15`, `2,4`, or `2-5,8,11-13`.
6. Use **Back to Setup** to change variable mapping or layout settings. Your records are kept while you return to Stage 1.
7. Choose the blank-cell behavior:
   - **Clear mapped text:** an empty cell produces an empty mapped text frame.
   - **Keep template text:** an empty cell leaves the original template text unchanged.
8. Use **Preview #1...** to inspect the first record before generating the full batch.
9. Use **Export PDF...** to create one multi-page PDF while keeping the data window open, or choose **Finish** to generate the artwork in the active Illustrator document.

After generation, you can save the Illustrator document normally. The script does not automatically save the `.ai` document. Close the document and choose **Don't Save** if you want to discard that generation session.

## Importing Data

The data grid uses **one record per row** and **one variable per column**.

Example:

```csv
Full Name,Reference,Issue Date
Alex Morgan,VD-001,2026-11-03
Sam Taylor,VD-002,2026-11-04
Jamie Lee,VD-003,2026-11-05
```

- Header names that match your variable names can be mapped automatically.
- If your spreadsheet uses different headings, use the column-mapping dialog to assign the correct source column to each variable.
- Imported records are appended to the existing table; they do not replace earlier records.
- Empty cells are allowed and follow the selected blank-cell behavior.
- For direct `.xlsx`, `.xls`, and `.xlsm` import, use Windows with Microsoft Excel installed. The script uses local Excel automation to export the selected worksheet as Unicode tab-delimited data.
- On macOS, or when Microsoft Excel automation is unavailable, export the worksheet as **CSV UTF-8** or **Tab Delimited Text (TXT/TSV)**, or copy and paste the cells from Excel.

## Tips and Best Practices

- Always test the workflow on a **duplicate of your Illustrator document** first.
- Start with two or three records and verify the text, layout, artboard size, and PDF output before processing a large list.
- Use a unique live text frame for each mapped variable.
- Keep mapped text frames consistently formatted where possible. The script attempts to preserve character and paragraph styling, but complex mixed-format text may behave differently across Illustrator versions.
- The script changes text content; it does not automatically redesign the text frame or guarantee that every long value will fit. Make the text area wide enough for the longest expected value and check the preview.
- Use clear variable names that match your spreadsheet headings to make column mapping easier.
- Save a `.vds` preset when you want to reuse the same mapping and data structure in another session.
- Use **Same Position** only when intentionally stacking outputs.
- Use **Separate Artboards** when each output needs its own page/artboard.
- If you enable artwork flattening, keep an editable backup of the original design. Static artwork is rasterized at the selected resolution; mapped variable text remains editable.

## Common Mistakes and Troubleshooting

| Problem | Likely cause | What to do |
|---|---|---|
| The script says no source Group is selected | No complete artwork Group was selected | Group the complete template and select that Group before running the script, or load a preset with an available template backup. |
| No text frames appear in the mapping list | The selected Group contains no live text, or the text was outlined | Keep variable text as editable point or area text. Outlined text cannot be mapped as a live text frame. |
| A variable is not assigned | A variable row still shows `< Select a text frame >` | Assign a live text frame to every variable before continuing. |
| A generated field contains the wrong value | The wrong source frame or spreadsheet column was mapped | Recheck Stage 1 mappings and the import column assignments. Test with a small dataset. |
| The first spreadsheet row is missing or imported as data | Header-row selection does not match the file | Review the header-row setting and source-column mapping. Confirm whether the first row contains headings or actual data. |
| Direct Excel import does not start | Microsoft Excel or Windows Script Host is unavailable, or an Excel dialog is blocking automation | Make sure Excel is installed, close modal Excel dialogs, and retry. Otherwise import CSV UTF-8 or TSV. |
| Arabic or special characters appear incorrectly | The file encoding or export format is unsuitable | Use Excel's Unicode text export, CSV UTF-8, or direct Excel import on supported Windows systems. |
| The grid does not have enough room for a Grid layout | The selected number of columns and rows is smaller than the record count | Increase the grid capacity or choose another layout. |
| The same-position outputs overlap | This layout intentionally stacks every generated item | Choose Grid, Horizontal, Vertical, or Separate Artboards if you need to see every output at once. |
| Text is clipped or overflows | The replacement value is longer than the space designed in the template | Widen or redesign the text frame in the source template, then preview with the longest values. |
| PDF export fails | The destination cannot be written, the file is locked, or Illustrator rejects an export option | Choose a writable folder, close an open PDF with the same name, confirm overwrite when prompted, and test with a small batch. |
| Generation stops with an Illustrator error | A document object or text reference may be unavailable, or the document structure is complex | Test on a fresh document copy, reduce the batch to one or two records, and note the record number and operation stage shown in the error. |

## Compatibility and Limitations

- This is a **desktop Adobe Illustrator JSX / ExtendScript** script. It is not an Illustrator web extension.
- The interface uses ScriptUI and ES3-compatible ExtendScript syntax. Behavior can vary across Illustrator releases; test with your installed version.
- Direct `.xlsx`, `.xls`, and `.xlsm` workbook import requires **Windows and Microsoft Excel installed**. CSV, TSV, TXT, and spreadsheet paste are alternatives.
- The Separate Artboards workflow checks Illustrator's 100-artboard limit. The combined multi-page PDF workflow in this script is also limited to 100 output pages per export. For larger jobs, split the batch into smaller groups or use the numbered PDF export after generation.
- Rasterizing static artwork reduces editability and depends on the chosen resolution. Keep the original editable document.
- The script has been source-checked, but it has not been validated across every Illustrator version, operating system, font setup, or document structure. Always test before production use.

## Developer

**Created by Ibraheem Sabbah**  
Graphic Designer | Automation and Creative Tools  
LinkedIn: [ibraheem-sabbah](https://www.linkedin.com/in/ibraheem-sabbah)

The script includes the author credit:

> This Script Made By Ibraheem Sabbah - www.linkedin.com/in/ibraheem-sabbah

## Copyright

Copyright © 2026 Ibraheem Sabbah. All rights reserved.

No open-source license is currently included in this repository. The project is not affiliated with, endorsed by, or sponsored by Adobe. Adobe Illustrator is a trademark of Adobe Inc.

---

<div align="center">
  <strong>Variable Data Studio</strong><br />
  <sub>Automate repetitive data-driven artwork production in Adobe Illustrator.</sub>
</div>
