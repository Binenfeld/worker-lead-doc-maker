package utils;

import models.Business;
import models.WorkerSheet;
import utils.DocxWriter.TableColors;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Builds the Hebrew report (.docx) of which unassigned workers can be offered to which business, on landscape pages.
 *
 * <p>Document layout:
 * <pre>
 *   Title:      שיבוץ עובדים לעסקים
 *   Paragraph:  date, number of businesses, workers asked for, workers we can supply
 *   For each business, in the leads sheet's order:
 *     Heading 1:  business name
 *     Paragraph:  its filled-in details (ח.פ., תחום, איש קשר, טלפון…)
 *     Paragraph:  per profession: "wanted profession - canSupply/wanted", e.g. "3 צבע - 2/3"
 *     Heading 2:  עובדים מתאימים - סך הכל: amount
 *       Table:      every worker of the wanted professions, by profession, in the sheet's order  [blue]
 * </pre>
 * A worker is offered to every business that wants their profession, so the same worker can appear under several
 * businesses. The table lists all of them, even when there are more than the business asked for. Each table fills
 * the page width: columns with values are only as wide as their contents, and the empty columns share the rest.
 */
public final class ReportFactory {
    private static final String TITLE = "שיבוץ עובדים לעסקים";
    private static final String TOTAL = "סך הכל";

    private ReportFactory() {
    }

    /** Writes the report for the given workers and businesses to {@code path}, overwriting any existing file. */
    public static void writeReport(WorkerSheet workers, List<Business> businesses, Path path) throws IOException {
        int wanted = businesses.stream().mapToInt(Business::totalWanted).sum();
        int canSupply = businesses.stream().mapToInt(b -> canSupply(workers, b)).sum();
        String today = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        DocxWriter doc = new DocxWriter(true, true)
                .addTitle(TITLE)
                .addParagraph("נכון לתאריך " + today + "  |  עסקים: " + businesses.size()
                        + "  |  עובדים מבוקשים: " + wanted + "  |  ניתן לספק: " + canSupply);

        for (Business business : businesses) {
            addBusinessSection(doc, workers, business);
        }

        doc.save(path);
    }

    /** "3 צבע - 2/3": how many workers of the profession are wanted, and how many of them we have. */
    public static String supplyLine(WorkerSheet workers, String profession, int wanted) {
        int available = Math.min(workers.withProfession(profession).size(), wanted);
        return wanted + " " + profession + " - " + available + "/" + wanted;
    }

    private static void addBusinessSection(DocxWriter doc, WorkerSheet workers, Business business) {
        doc.addHeading(1, business.name());
        if (!business.details().isEmpty()) {
            doc.addParagraph(business.details().entrySet().stream()
                    .map(detail -> detail.getKey() + ": " + detail.getValue())
                    .collect(Collectors.joining("  |  ")));
        }
        List<List<String>> matching = new ArrayList<>();
        for (Map.Entry<String, Integer> profession : business.wanted().entrySet()) {
            doc.addParagraph(supplyLine(workers, profession.getKey(), profession.getValue()));
            matching.addAll(workers.withProfession(profession.getKey()));
        }

        doc.addHeading(2, withTotal("עובדים מתאימים", matching.size()));
        if (matching.isEmpty()) {
            doc.addParagraph("אין כרגע עובדים מתאימים.");
        } else {
            doc.addFittedTable(null, TableColors.BLUE, workers.headers(), matching, null, columnsToStretch(matching));
        }
    }

    // How many of the workers the business asked for we can supply, over all its professions.
    private static int canSupply(WorkerSheet workers, Business business) {
        return business.wanted().entrySet().stream()
                .mapToInt(p -> Math.min(workers.withProfession(p.getKey()).size(), p.getValue()))
                .sum();
    }

    // The columns with no value in any of the rows, or every column if none is empty.
    private static List<Integer> columnsToStretch(List<List<String>> rows) {
        List<Integer> empty = new ArrayList<>();
        List<Integer> all = new ArrayList<>();
        for (int c = 0; c < rows.getFirst().size(); c++) {
            int column = c;
            all.add(c);
            if (rows.stream().allMatch(row -> row.get(column).isEmpty())) {
                empty.add(c);
            }
        }
        return empty.isEmpty() ? all : empty;
    }

    // "title - סך הכל: amount", used for every heading that shows a number of workers.
    private static String withTotal(String title, int amount) {
        return title + " - " + TOTAL + ": " + amount;
    }
}
