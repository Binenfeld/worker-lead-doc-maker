package utils;

import models.Business;
import models.WorkerSheet;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Reads the worker-leads sheet ("לידים עובדים"), exported from Excel as "CSV UTF-8".
 *
 * <p>The first row is the header; every other row says how many workers of one profession a business wants.
 * The sheet's columns, as of writing:
 * <pre>
 *   שם עסק | ח.פ. | קבלן (כן/רגי) | תחום | איש קשר | טלפון | מקצוע | כמות | … | אינדקס מקצועות
 * </pre>
 * Columns are found by their header, not their position. Only {@value #BUSINESS_HEADER},
 * {@value #PROFESSION_HEADER} and {@value #QUANTITY_HEADER} must exist; the details in {@link #DETAILS} are shown
 * under the business name when filled in. Other columns, such as the profession index, are ignored.
 */
public final class LeadParser {
    // Header texts of the columns the sheet must have.
    public static final String BUSINESS_HEADER = "שם עסק";
    public static final String PROFESSION_HEADER = "מקצוע";
    public static final String QUANTITY_HEADER = "כמות";

    private static final List<String> REQUIRED_HEADERS = List.of(BUSINESS_HEADER, PROFESSION_HEADER, QUANTITY_HEADER);

    // The detail columns shown under the business name: header in the sheet to label in the report.
    private static final Map<String, String> DETAILS = new LinkedHashMap<>();

    static {
        DETAILS.put("ח.פ.", "ח.פ.");
        DETAILS.put("קבלן (כן/רגי)", "קבלן");
        DETAILS.put("תחום", "תחום");
        DETAILS.put("איש קשר", "איש קשר");
        DETAILS.put("טלפון", "טלפון");
    }

    private LeadParser() {
    }

    /** True if this header row is the leads sheet's: it has all the required columns. */
    public static boolean looksLikeLeads(List<String> header) {
        return header.containsAll(REQUIRED_HEADERS);
    }

    /** The required column headers, for messages. */
    public static List<String> requiredHeaders() {
        return REQUIRED_HEADERS;
    }

    /**
     * Returns the businesses in the order they first appear. Rows of the same business are put together, and the
     * quantities of a profession it lists twice are added up. Rows without a business name are skipped.
     */
    public static List<Business> parseLeadCsv(Path csvFile) throws IOException, UserFacingException {
        List<List<String>> records = CsvReader.readRecords(csvFile);
        if (records.isEmpty() || records.getFirst().stream().allMatch(String::isEmpty)) {
            throw new UserFacingException(csvFile + " is empty, or its first row is. Export the leads sheet again,"
                    + " with the header in the first row.");
        }
        List<String> header = records.getFirst();
        checkHeader(csvFile, header);

        int businessColumn = header.indexOf(BUSINESS_HEADER);
        int professionColumn = header.indexOf(PROFESSION_HEADER);
        int quantityColumn = header.indexOf(QUANTITY_HEADER);
        Map<String, Integer> detailColumns = new LinkedHashMap<>();
        DETAILS.forEach((detailHeader, label) -> {
            int column = header.indexOf(detailHeader);
            if (column < 0) {
                System.out.println("Warning: the leads sheet has no \"" + detailHeader + "\" column, so the report"
                        + " leaves it out. If it was renamed, update DETAILS in LeadParser.");
            } else {
                detailColumns.put(label, column);
            }
        });
        List<Integer> leadColumns = new ArrayList<>(List.of(businessColumn, professionColumn, quantityColumn));
        leadColumns.addAll(detailColumns.values());

        // Business name to its details and wanted professions, in first-appearance order.
        Map<String, Map<String, String>> details = new LinkedHashMap<>();
        Map<String, Map<String, Integer>> wanted = new LinkedHashMap<>();
        for (int i = 1; i < records.size(); i++) {
            List<String> row = CsvReader.padded(records.get(i), header.size());
            int rowNumber = i + 1;
            String business = row.get(businessColumn);
            if (business.isEmpty()) {
                // Rows with only the profession index (a separate list on the side of the sheet) are expected.
                if (leadColumns.stream().anyMatch(c -> !row.get(c).isEmpty())) {
                    System.out.println("Warning: row " + rowNumber + " of the leads sheet has no business name ("
                            + BUSINESS_HEADER + ") and is left out: " + String.join(" | ", row));
                }
                continue;
            }
            String profession = row.get(professionColumn);
            if (profession.isEmpty()) {
                throw new UserFacingException("Row " + rowNumber + " of the leads sheet (" + business + ") has no"
                        + " profession in the \"" + PROFESSION_HEADER + "\" column. Fill it in, save the sheet as"
                        + " CSV UTF-8 again, and run again.");
            }
            int quantity = parseQuantity(row.get(quantityColumn), rowNumber, business);

            Map<String, String> businessDetails = details.computeIfAbsent(business, b -> new LinkedHashMap<>());
            detailColumns.forEach((label, column) -> {
                // The first filled-in value wins when the business's rows differ.
                if (!row.get(column).isEmpty()) {
                    businessDetails.putIfAbsent(label, row.get(column));
                }
            });
            Map<String, Integer> businessWanted = wanted.computeIfAbsent(business, b -> new LinkedHashMap<>());
            String key = businessWanted.keySet().stream()
                    .filter(p -> WorkerSheet.sameProfession(p, profession))
                    .findFirst()
                    .orElse(profession);
            businessWanted.merge(key, quantity, Integer::sum);
        }
        if (wanted.isEmpty()) {
            throw new UserFacingException(csvFile + " has the leads sheet's header but no leads (no row has a"
                    + " business name in the \"" + BUSINESS_HEADER + "\" column). Check that the right sheet was exported.");
        }

        // Details are put in the sheet's column order, whatever row they came from.
        List<Business> businesses = new ArrayList<>();
        wanted.forEach((name, professions) -> {
            Map<String, String> ordered = new LinkedHashMap<>();
            detailColumns.keySet().forEach(label -> {
                String value = details.get(name).get(label);
                if (value != null) {
                    ordered.put(label, value);
                }
            });
            businesses.add(new Business(name, ordered, professions));
        });
        return businesses;
    }

    // A whole number of workers, at least 1.
    private static int parseQuantity(String text, int rowNumber, String business) throws UserFacingException {
        String fix = " Fix it in the leads sheet, save it as CSV UTF-8 again, and run again.";
        if (text.isEmpty()) {
            throw new UserFacingException("Row " + rowNumber + " of the leads sheet (" + business + ") has no"
                    + " number of workers in the \"" + QUANTITY_HEADER + "\" column." + fix);
        }
        int quantity;
        try {
            quantity = Integer.parseInt(text);
        } catch (NumberFormatException e) {
            throw new UserFacingException("Row " + rowNumber + " of the leads sheet (" + business + ") has \"" + text
                    + "\" in the \"" + QUANTITY_HEADER + "\" column, which is not a whole number." + fix, e);
        }
        if (quantity < 1) {
            throw new UserFacingException("Row " + rowNumber + " of the leads sheet (" + business + ") asks for "
                    + quantity + " workers in the \"" + QUANTITY_HEADER + "\" column; it must be at least 1." + fix);
        }
        return quantity;
    }

    // Makes sure this is the leads sheet: its main columns must be there.
    private static void checkHeader(Path csvFile, List<String> header) throws UserFacingException {
        List<String> missing = REQUIRED_HEADERS.stream().filter(h -> !header.contains(h)).toList();
        if (!missing.isEmpty()) {
            throw new UserFacingException(csvFile + " does not look like the leads sheet: the first row has no "
                    + String.join(", ", missing.stream().map(h -> "\"" + h + "\"").toList())
                    + " column (found: " + String.join(", ", header) + "). Check that the right sheet was exported,"
                    + " with the header in the first row. If a column was renamed, update the header constants in"
                    + " LeadParser.");
        }
    }
}
