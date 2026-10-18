# WorkerLeadDocMaker

Matches the unassigned workers (`פרטי מידע לא משוייכים`) to the businesses that asked for workers (`לידים עובדים`),
and writes a Hebrew Word report of which workers can be offered to which business, plus a PDF copy of it.
A sibling of WorkerDocMaker and ExpandedWorkerDocMaker, built the same way.

```
workers.csv ──► WorkerParser ──► WorkerSheet ──┐
                                               ├──► ReportFactory ──► שיבוץ עובדים לעסקים <date>.docx ──► PdfConverter ──► .pdf
leads.csv ────► LeadParser ────► Business… ────┘
```

## What the report contains

The report is right-to-left Hebrew on landscape A4 pages, with a 9pt table font. The businesses follow each other,
in the order they first appear in the leads sheet:

| Section | Content |
|---|---|
| Title + date line | `שיבוץ עובדים לעסקים`, today's date, the number of businesses, the workers they asked for, and how many of those we can supply |
| Business name (Heading 1) | One section per business |
| Details line | The business's filled-in details: ח.פ., קבלן, תחום, איש קשר, טלפון (empty ones are left out) |
| One line per profession | How many it wants, the profession, and how many of those we can supply out of how many it wants, e.g. `3 צבע - 2/3` |
| עובדים מתאימים - סך הכל (Heading 2) | A table of every unassigned worker with one of the wanted professions, by profession, in the workers sheet's order. Same columns as ExpandedWorkerDocMaker's report. The header row repeats on every page |

How the matching works:

- A worker matches a business when the worker's `מקצוע` is the same text as the lead's `מקצוע` (extra spaces are
  ignored). So the professions in both sheets must be written the same way, e.g. from the leads sheet's
  `אינדקס מקצועות` list.
- Every matching worker is offered to **every** business that wants the profession, so the same worker can appear
  under several businesses.
- The table lists **all** matching workers, even when there are more than the business asked for, so there is a
  choice. The line above it shows how many it asked for: with 5 painters for a request of 3, it reads `3 צבע - 3/3`.
- A business with no matching worker gets the line `אין כרגע עובדים מתאימים` instead of a table.
- The worker table fills the page width: columns with values are only as wide as their contents, and the empty
  columns share the rest, so there is room to fill them in on the printout.

## Requirements

- JDK 21 or newer (the project is developed with JDK 26)
- Gradle, through the included wrapper (`gradlew`); no external libraries are needed
- Excel, to export the sheets to CSV
- Microsoft Word (desktop), to convert the report to PDF. `PdfConverter` drives it through PowerShell, so this
  needs Windows. Without Word the run fails after the .docx is written.

## Running it

Step-by-step instructions in Hebrew for non-technical users: [`הוראות הפעלה.txt`](הוראות%20הפעלה.txt).

1. **Export both sheets.** Open each workbook (the unassigned-workers sheet, and the worker-leads sheet) and use
   File → Save As → **CSV UTF-8 (Comma delimited)**. Save both in `C:\workerleaddocmaker`, under any names ending
   in `.csv`. It must be the UTF-8 variant, or the Hebrew text will be garbled.
2. **Build** (once, and again after code changes):
   ```
   powershell -ExecutionPolicy Bypass -File scripts\build.ps1
   ```
   This produces `build\libs\WorkerLeadDocMaker.jar`. The script uses `JAVA_HOME` if it is set, otherwise the
   newest JDK in `~\.jdks` (where IntelliJ downloads JDKs).
3. **Run:**
   ```
   powershell -ExecutionPolicy Bypass -File scripts\run.ps1
   ```
   Or just **double-click `scripts\run.cmd`**, which does the same and keeps the window open to show the result.
   You can also run `Main` from IntelliJ. Options:
   - `-InputFolder D:\some\folder`: use a different folder
   - `-Debug`: also print every worker read, every business with its `x/y` lines, and the full stack trace of an
     error
4. The program tells the two CSVs apart **by their header row**, so file names don't matter. If there are
   several of a kind, it uses the **newest** one. It writes the report to the same folder as
   `שיבוץ עובדים לעסקים <yyyy-mm-dd>.docx`, plus a PDF copy (`.pdf`, same name) made with Word.

The default input folder is `Main.DEFAULT_INPUT_FOLDER`. It is a separate folder from the other two programs'
(`C:\workerdocmaker`, `C:\expandedworkerdocmaker`), so the workers CSV is saved here too. The jar takes the same
options directly: `java -jar WorkerLeadDocMaker.jar [folder] [--debug]`. Excel lock files (`~$…`) are ignored.
A CSV that is neither sheet is ignored with a warning. `-ExecutionPolicy Bypass` is only needed if Windows blocks
running scripts.

### When something goes wrong

The run stops with exit code 1 and a line starting with `ERROR:` that says what is wrong and how to fix it. For
example, it reports a missing folder, a missing workers or leads CSV, a CSV saved as plain CSV instead of CSV
UTF-8, a lead without a profession or with a quantity that isn't a whole number, a sheet without workers or
leads, a report left open in Word, and a missing Word install. `run.cmd` then shows "Something went wrong". An
`ERROR: Unexpected error` comes with a stack trace; that one is a bug or something unusual on the computer.

## Input formats

In both sheets the first row is the header. **Columns are found by their header**, not their position, so they can
be reordered and new ones added. Cells are trimmed.

### Unassigned workers (`פרטי מידע לא משוייכים`)

Every row after the header is one worker. The same sheet ExpandedWorkerDocMaker reads:

| Header | Meaning |
|---|---|
| מסד | Running number (**shown**) |
| מספר עובד | Worker ID (not shown) |
| שם מלא | Full name (**required**; not shown, but tells workers from empty rows) |
| מספר דרכון | Passport number (**shown**) |
| ארץ מוצא | Country of origin (**required, shown**) |
| ותק בארץ | Time in Israel (**shown**) |
| מקצוע | Profession (**required, shown**; matched against the leads) |
| שפות, רשיון נהיגה, חווד על עובד, היסטורית לקוחות, האם פוטר מאיפה ולמה | Details (**shown**, usually empty) |

- Only the columns marked "shown" are in the tables, in the sheet's order. The list is `SHOWN_HEADERS` in
  `WorkerParser` (quotes are ignored when comparing, so `חוו"ד` matches `חווד`).
- Rows without a name are left out; a nameless row with other content gets a warning.

### Worker leads (`לידים עובדים`)

Every row after the header says how many workers of one profession a business wants. A business that wants
several professions has one row for each.

| Header | Meaning |
|---|---|
| שם עסק | Business name (**required**). Rows with the same name are put together |
| ח.פ. | Company number (details line) |
| קבלן (כן/רגי) | Contractor (details line, shown as `קבלן`) |
| תחום | Field (details line) |
| איש קשר | Contact person (details line) |
| טלפון | Phone (details line) |
| מקצוע | Wanted profession (**required**) |
| כמות | How many workers of it (**required**, a whole number, at least 1) |
| אינדקס מקצועות | A separate list of profession names; ignored |

- The three required columns are constants in `LeadParser`; the details columns are `DETAILS` there. A missing
  details column is left out with a warning.
- Rows without a business name are skipped. That covers the rows that only hold the profession index. A nameless
  row with a profession, quantity or detail gets a warning.
- If a business lists the same profession twice, the quantities are added up. When its rows have different
  details, the first filled-in value is used.

## Project structure

```
scripts/
├── build.ps1                 builds build\libs\WorkerLeadDocMaker.jar
├── run.ps1                   runs the jar on the newest CSVs in C:\workerleaddocmaker
└── run.cmd                   double-click launcher for run.ps1
src/main/java/
├── Main.java                 entry point: find both CSVs → parse → write report → PDF
├── models/
│   ├── WorkerSheet.java      the workers sheet: header + one row per worker, and profession matching
│   └── Business.java         a business from the leads: details and wanted professions with quantities
└── utils/
    ├── CsvReader.java        reads a CSV UTF-8 export into records (shared by both parsers)
    ├── WorkerParser.java     reads the workers CSV into a WorkerSheet
    ├── LeadParser.java       reads the leads CSV into businesses
    ├── ReportFactory.java    lays out the report (sections, x/y lines, worker tables)
    ├── DocxWriter.java       minimal .docx writer (titles, headings, RTL, landscape, colored and fitted tables) without libraries
    ├── PdfConverter.java     .docx → PDF through Word (PowerShell + COM); always closes the Word it starts
    ├── UserFacingException.java  a problem the user can fix; Main prints its message without a stack trace
    └── FileIOHandler.java    small UTF-8 file helpers
```

`DocxWriter` writes the Office Open XML parts by hand and zips them, so the project has no dependencies.
Headings use Word's built-in Heading styles, so the businesses appear in Word's navigation pane.

## License

Copyright (C) 2026 Binenfeld. All rights reserved.

Proprietary and confidential: no one may access, use, copy, modify or distribute this software without written
permission from Binenfeld. See [LICENSE](LICENSE).
