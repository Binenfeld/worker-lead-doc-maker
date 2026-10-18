package models;

import java.util.List;

/**
 * The unassigned-workers sheet as read from the CSV: its header row and one row per worker, in the sheet's order.
 * Every row has exactly one cell per header; cells are trimmed and never null.
 */
public record WorkerSheet(List<String> headers, List<List<String>> rows) {
    // Header texts of the columns the sheet must have.
    public static final String NAME_HEADER = "שם מלא";
    public static final String COUNTRY_HEADER = "ארץ מוצא";
    public static final String PROFESSION_HEADER = "מקצוע";

    public WorkerSheet {
        headers = List.copyOf(headers);
        rows = rows.stream().map(List::copyOf).toList();
    }

    /** The 0-based index of the column with this header, or -1 if the sheet has no such column. */
    public int column(String header) {
        return headers.indexOf(header);
    }

    /** The workers of the given profession, in the sheet's order (see {@link #sameProfession}). */
    public List<List<String>> withProfession(String profession) {
        int column = column(PROFESSION_HEADER);
        return rows.stream().filter(row -> sameProfession(row.get(column), profession)).toList();
    }

    /** True if the two professions are the same text, ignoring extra spaces ("שיפוצים  כללי" = "שיפוצים כללי"). */
    public static boolean sameProfession(String a, String b) {
        return normalize(a).equals(normalize(b));
    }

    private static String normalize(String profession) {
        return profession.strip().replaceAll("\\s+", " ");
    }

    @Override
    public String toString() {
        return "WorkerSheet[" + headers.size() + " columns, " + rows.size() + " workers]";
    }
}
