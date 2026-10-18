package utils;

import models.WorkerSheet;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * Reads the unassigned-workers sheet ("פרטי מידע לא משוייכים"), exported from Excel as "CSV UTF-8".
 *
 * <p>The first row is the header; every other row is one worker. The sheet's columns, as of writing:
 * <pre>
 *   מסד | מספר עובד | שם מלא | מספר דרכון | ארץ מוצא | ותק בארץ | מקצוע | שפות | רשיון נהיגה |
 *   חווד על עובד | היסטורית לקוחות | האם פוטר מאיפה ולמה
 * </pre>
 * Columns are found by their header, not their position, so they can be reordered and others can be added.
 * Only {@value WorkerSheet#NAME_HEADER}, {@value WorkerSheet#COUNTRY_HEADER} and
 * {@value WorkerSheet#PROFESSION_HEADER} must exist. Only the columns in {@link #SHOWN_HEADERS} are kept for the
 * report, in the sheet's order; the name is used to tell workers from empty rows, but is not shown.
 */
public final class WorkerParser {
    private static final List<String> REQUIRED_HEADERS =
            List.of(WorkerSheet.NAME_HEADER, WorkerSheet.COUNTRY_HEADER, WorkerSheet.PROFESSION_HEADER);

    // The columns shown in the report; any other column is left out, even when filled in.
    // Compared without quotes, so "חוו""ד על עובד" also matches. The profession is matched against the leads,
    // so it must stay in this list.
    private static final List<String> SHOWN_HEADERS = List.of("מסד", "מספר דרכון", WorkerSheet.COUNTRY_HEADER,
            "ותק בארץ", WorkerSheet.PROFESSION_HEADER, "שפות", "רשיון נהיגה", "חווד על עובד", "היסטורית לקוחות",
            "האם פוטר מאיפה ולמה");

    private WorkerParser() {
    }

    /** True if this header row is the unassigned-workers sheet's: it has all the required columns. */
    public static boolean looksLikeWorkers(List<String> header) {
        return header.containsAll(REQUIRED_HEADERS);
    }

    /** The required column headers, for messages. */
    public static List<String> requiredHeaders() {
        return REQUIRED_HEADERS;
    }

    /**
     * Returns the sheet's header and worker rows. Rows without a name are skipped: empty rows, and rows Excel
     * exported because a cell further down was once formatted. Only the shown columns are kept.
     */
    public static WorkerSheet parseWorkerCsv(Path csvFile) throws IOException, UserFacingException {
        List<List<String>> records = CsvReader.readRecords(csvFile);
        if (records.isEmpty() || records.getFirst().stream().allMatch(String::isEmpty)) {
            throw new UserFacingException(csvFile + " is empty, or its first row is. Export the sheet again,"
                    + " with the header in the first row.");
        }
        List<String> header = records.getFirst();
        checkHeader(csvFile, header);

        int nameColumn = header.indexOf(WorkerSheet.NAME_HEADER);
        List<List<String>> rows = new ArrayList<>();
        for (int i = 1; i < records.size(); i++) {
            List<String> row = CsvReader.padded(records.get(i), header.size());
            if (row.get(nameColumn).isEmpty()) {
                long filled = row.stream().filter(cell -> !cell.isEmpty()).count();
                // A lone number is just the running number (מסד) of an unused row.
                if (filled > 1) {
                    System.out.println("Warning: row " + (i + 1) + " of the workers sheet has no name ("
                            + WorkerSheet.NAME_HEADER + ") and is left out: " + String.join(" | ", row));
                }
                continue;
            }
            rows.add(row);
        }
        if (rows.isEmpty()) {
            throw new UserFacingException(csvFile + " has the workers sheet's header but no workers (no row has a"
                    + " name in the \"" + WorkerSheet.NAME_HEADER + "\" column). Check that the right sheet was exported.");
        }
        return keepShownColumns(header, rows);
    }

    // Makes sure this is the unassigned-workers sheet: its main columns must be there.
    private static void checkHeader(Path csvFile, List<String> header) throws UserFacingException {
        List<String> missing = REQUIRED_HEADERS.stream().filter(h -> !header.contains(h)).toList();
        if (!missing.isEmpty()) {
            throw new UserFacingException(csvFile + " does not look like the unassigned-workers sheet: the first row"
                    + " has no " + String.join(", ", missing.stream().map(h -> "\"" + h + "\"").toList())
                    + " column (found: " + String.join(", ", header) + "). Check that the right sheet was exported,"
                    + " with the header in the first row. If a column was renamed, update the header constants in"
                    + " WorkerSheet.");
        }
    }

    // Keeps only the shown columns, in the sheet's order. A shown column missing from the sheet gets a warning.
    private static WorkerSheet keepShownColumns(List<String> header, List<List<String>> rows) {
        List<String> shown = SHOWN_HEADERS.stream().map(CsvReader::withoutQuotes).toList();
        List<String> found = header.stream().map(CsvReader::withoutQuotes).toList();
        for (int i = 0; i < shown.size(); i++) {
            if (!found.contains(shown.get(i))) {
                System.out.println("Warning: the workers sheet has no \"" + SHOWN_HEADERS.get(i) + "\" column, so"
                        + " the report leaves it out. If it was renamed, update SHOWN_HEADERS in WorkerParser.");
            }
        }
        List<Integer> kept = new ArrayList<>();
        for (int c = 0; c < header.size(); c++) {
            if (shown.contains(found.get(c))) {
                kept.add(c);
            }
        }
        List<String> keptHeader = kept.stream().map(header::get).toList();
        List<List<String>> keptRows = rows.stream().map(row -> kept.stream().map(row::get).toList()).toList();
        return new WorkerSheet(keptHeader, keptRows);
    }
}
