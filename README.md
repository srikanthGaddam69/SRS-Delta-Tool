# SRS Delta Tool

Compare two versions of an SRS Excel export (for example `SRS_11.00.xlsx` and `SRS_11.01.xlsx`) and generate a delta workbook with three tabs:

| Tab | Contents |
|---|---|
| **Delta** | One row per requirement ID: previous value, new value and a Yes/No change flag for each compared column, plus a Status column (New / Deleted / Modified / Unchanged). |
| **Summary** | Counts of new, deleted, modified and unchanged requirements; changes per factor; Safety Related breakdown; old → new value changes per column. Counts are live Excel formulas. |
| **Revision History** | Revision number, date, description of change and source files. Carried forward from the previous delta workbook. |

**Open the tool:** https://srikanthGaddam69.github.io/SRS-Delta-Tool/

Your files are processed entirely in your browser. Nothing is uploaded to any server. The tool also works offline: download `index.html` and open it in Edge or Chrome.

## How to use

1. **Load files:** the previous SRS and the new SRS. Optionally, also load the last delta workbook so that its Revision History is carried forward and the revision number goes up by one.
2. **Check the set-up:** the sheet, version labels, Requirement ID column and column pairing are detected from the headers automatically. Untick, rename or re-pair columns if needed.
3. **Revision:** check the revision number and enter a description of the change.
4. **Run comparison:** review the counts and the filterable list of changed requirements (click a row to see the old and new values), then download the workbook.

## Comparison rules

- Requirements are matched by **Requirement ID**. An ID found only in the new file is **New**; an ID found only in the previous file is **Deleted**.
- A requirement is **Modified** if any selected column differs after ignoring:
  - spaces, line breaks and tabs;
  - punctuation: `. , ; : ! ?`, quotes, dashes, bullets (`•`), brackets and underscore;
  - letter case (optional, on by default).
- Symbols such as `< > = / %` are **kept**, because they can change the meaning of a requirement.
- **List** columns (Geographical Areas, Migration Stages, Stage Introduced, Subsystem Allocation) are compared as sets of items. Changing the separator (line break vs comma) or the order of items is therefore not counted as a change.
- Formula cells are compared on their saved values.

## Repository layout

```
index.html        The tool: one self-contained file, ExcelJS included (works offline)
src/core.js       Comparison and workbook-building logic
src/page.html     User interface (template)
src/build.js      Builds index.html: node src/build.js (needs `npm i exceljs@4.4.0`)
```

## Version history

| Version | Date | Change |
|---|---|---|
| 1.1 | 2026-11-09 | Run comparison button always visible, with a status hint |
| 1.0 | 2026-11-09 | Initial release |
