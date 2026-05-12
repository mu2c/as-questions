# ConsolidateTables — PowerPoint Q&A tables → one Excel workbook

A one-button macro that scans a folder of PowerPoint files, finds every table
whose first row is exactly **`#` | `Question` | `Answer`**, and copies all
those rows into a single Excel workbook (`consolidated.xlsx`) — one sheet per
PowerPoint file, plus a `_log` sheet that explains anything that was skipped.

You do not need to know how to code to use this. Follow the steps below.

---

## What you get

- `ConsolidateTables.bas` — the macro module (you import this once).
- `README.md` — this file.

You will create one extra file yourself during setup:
`ConsolidateTables.xlsm` — a macro-enabled Excel workbook that holds the
macro. Keep it on your Desktop or in Documents.

---

## One-time setup (about 5 minutes)

You only do this once.

1. Save `ConsolidateTables.bas` somewhere you can find it (e.g. Desktop).
2. Open Excel and create a **blank workbook**.
3. Press **Alt + F11**. A second window opens called the *Visual Basic
   Editor*. It looks technical — that's fine, you only need one menu.
4. In that window's menu: **File → Import File…**, choose
   `ConsolidateTables.bas`, click **Open**.
5. Close the Visual Basic Editor (just click its X).
6. Back in Excel: **File → Save As**.
   - Change *Save as type* to **Excel Macro-Enabled Workbook (*.xlsm)**.
   - Name the file **`ConsolidateTables.xlsm`**.
   - Save it to your Desktop or Documents.
7. **First time you reopen the file**, Excel will show a yellow bar:
   *SECURITY WARNING — Macros have been disabled*. Click **Enable Content**.
   You will not be asked again on this PC.

That's it. From now on, you only ever open `ConsolidateTables.xlsm`.

> **Tip — pin it to the taskbar.** Right-click the file → *Pin to taskbar*
> so it's always one click away.

---

## How to run it (about 30 seconds of clicks)

1. Open **`ConsolidateTables.xlsm`** (double-click it).
2. If a yellow security bar appears, click **Enable Content**.
3. Press **Alt + F8**. A small dialog appears listing macros.
4. Select **`ConsolidateTables`** and click **Run**.
5. A folder picker opens. Browse to the folder that contains your PowerPoint
   files and click **OK**.
6. If `consolidated.xlsx` already exists in that folder, you are asked
   *Overwrite it?* — click **Yes** to replace it, **No** to cancel.
7. Wait. The bottom-left of Excel shows progress like
   *Processing 7 of 18: ProjectAlpha.pptx*.
   - 5 files: ~30 seconds.
   - 20 files: 2–3 minutes.
   - **Do not click anywhere in Excel while it runs.** PowerPoint is being
     driven invisibly in the background and clicks can interrupt it.
8. When done, a summary box appears:

   > Done.
   >
   > Files processed:  18
   > Sheets created:   16
   > Rows extracted:   423
   > Warnings:         3
   > Errors:           0
   >
   > Output: C:\Users\you\Decks\consolidated.xlsx
   >
   > See the _log sheet for per-slide details.

9. Click **OK**, then open **`consolidated.xlsx`** from the folder you
   picked. That is your result.

---

## Reading the output

`consolidated.xlsx` has:

- **One tab per PowerPoint file**, named after the file (truncated to fit
  Excel's 31-character limit; `_2`, `_3` suffix is added if two files would
  collide).
- Each tab has three columns: `#`, `Question`, `Answer`.
- One extra tab at the end called **`_log`** with these columns:
  `Timestamp | Level | File | Slide | TableIdx | RowsExtracted | Message`.
  - Level **INFO** = a table was extracted.
  - Level **WARN** = something was skipped (header didn't match, multiple
    tables on a slide, file produced no data).
  - Level **ERROR** = a file could not be opened or processed.

If anything was skipped, the macro automatically opens the `_log` tab so you
notice it.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Macro hangs and Excel becomes unresponsive | One of the .pptx files is password-protected or in PowerPoint's *Protected View*. PowerPoint is silently waiting on a hidden dialog. | Open Task Manager (Ctrl+Shift+Esc) → end **POWERPNT.EXE**. The file will be logged as an error. Remove or unlock that file and rerun. |
| "Cannot save consolidated.xlsx" message | The output file is open in another Excel window. | Close every open copy of `consolidated.xlsx` and rerun. |
| Empty result, or fewer sheets than files | Header text in the slide tables does not match `#` / `Question` / `Answer` **exactly** (case-sensitive). | Open the `_log` tab — every skip is listed with file name and slide number. Fix the header text in the source deck (or ask for the case-sensitivity rule to be relaxed). |
| `_log` shows "Slide has 2 tables; skipped per multi-table rule" | The slide has more than one table on it; the macro skips the entire slide on purpose. | Either consolidate the tables in the source deck, or accept the omission. |
| File appears as **ERROR** in the log | The file is on OneDrive and marked *online-only*; or it is a `.ppt` (legacy format), `.pptm`, or `.ppsx` — only `.pptx` is supported. | In File Explorer, right-click the file → *Always keep on this device* and rerun. For non-`.pptx` formats, save them as `.pptx` first. |
| `MsgBox` says "Could not start PowerPoint" | PowerPoint is not installed on this PC, or is broken. | Reinstall / repair Office. |
| Macro doesn't appear in the Alt+F8 list | The file you opened is the wrong workbook (the `.bas` was imported into a different file). | Reopen `ConsolidateTables.xlsm` specifically. |

---

## What the macro does NOT do

- Does **not** scan subfolders.
- Does **not** read `.ppt`, `.pptm`, or `.ppsx` — only `.pptx`.
- Does **not** open password-protected files (will hang if it tries).
- Does **not** modify your source decks in any way; it opens them read-only.
- Does **not** quit PowerPoint if you already had it open with another file.

---

## Optional: build a fixture deck to test it

To make sure the macro behaves as expected on your machine before running it
against real data, build a small test deck (one PowerPoint file, six slides):

| Slide | Build this on the slide                                                                          | Expected result in `consolidated.xlsx`                            |
| ----- | ------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------- |
| A     | A 3-column table, header `# / Question / Answer`, three data rows                                | The three rows appear in the output sheet                         |
| B     | The same table as A, but **grouped** with another shape (Insert → Shape, then group both)        | Rows still appear (the macro looks inside groups)                 |
| C     | A 3-column table with header `Num / Q / A`                                                       | Slide is skipped; `_log` shows a WARN                             |
| D     | Two separate tables (one matching, one not)                                                      | Slide is skipped entirely; `_log` shows a WARN                    |
| E     | A matching table with one empty row, one cell containing two paragraphs, and one cell starting with `=` | Empty row dropped, line break preserved, `=`-cell shown as text, no error |
| F     | No table at all                                                                                  | Silently ignored                                                  |

Save the deck as `test.pptx` in any folder, run the macro against that folder,
and confirm the output matches the table above.

---

## Editable settings (advanced — only if you want to)

Open the macro in the VBA editor (Alt+F11, double-click `ConsolidateTables`
in the left pane). Near the top you'll see:

```vba
Private Const DRY_RUN  As Boolean = False     ' True = log only, no data sheets
Private Const VERBOSE  As Boolean = False     ' True = log every slide examined
Private Const HEADER_COL1 As String = "#"
Private Const HEADER_COL2 As String = "Question"
Private Const HEADER_COL3 As String = "Answer"
Private Const OUTPUT_FILENAME As String = "consolidated.xlsx"
```

- Set `DRY_RUN = True` to do a "what would happen" run that produces only the
  `_log` sheet — no data sheets, no overwrite of your existing output. Good
  for diagnosing why slides are being skipped.
- Set `VERBOSE = True` to log every slide examined, not just the ones with
  problems.
- Change `HEADER_COL1/2/3` if your decks use different header text. The
  comparison is case-sensitive.

Save the file (Ctrl+S) after editing.
