package utils;

import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Builds a simple Word (.docx) document with a title, headings (levels 1-3), paragraphs and tables, on A4 pages
 * in portrait or landscape.
 * Headings use Word's built-in "Heading 1/2/3" styles, so they appear in the navigation pane
 * and can be used for a table of contents.
 *
 * Usage:
 *   new DocxWriter(true)
 *       .addTitle("Report")
 *       .addHeading(1, "Section")
 *       .addParagraph("Some text")
 *       .save(Path.of("report.docx"));
 */
public final class DocxWriter {
    private static final String FONT = "Arial";

    // A4 page in twips (1 cm = 567), and the margin on each side (1.27 cm, so as much as possible fits).
    private static final int PAGE_SHORT_SIDE = 11906;
    private static final int PAGE_LONG_SIDE = 16838;
    private static final int MARGIN = 720;

    // Table text size in half-points (18 = 9pt), smaller than body text so as much as possible fits.
    private static final int TABLE_FONT_SIZE = 18;

    // For measuring text in addFittedTable: the table font at its size in points.
    private static final Font TABLE_FONT_PLAIN = new Font(FONT, Font.PLAIN, TABLE_FONT_SIZE / 2);
    private static final Font TABLE_FONT_BOLD = TABLE_FONT_PLAIN.deriveFont(Font.BOLD);
    private static final FontRenderContext FONT_RENDER_CONTEXT = new FontRenderContext(null, true, true);

    // Added to a fitted column's text width (twips): the cell's left and right margins (80 each), plus a little
    // room so Word's own measuring never wraps a value that was measured to fit.
    private static final int FITTED_CELL_PADDING = 2 * 80 + 120;

    // Space between the two tables of addSideBySide (twips; 1 cm = 567).
    private static final int SIDE_BY_SIDE_GAP = 400;

    private static final String TABLE_BORDER = "A6A6A6";

    /**
     * A table with an optional heading above it, for {@link #addSideBySide(Table, Table)}.
     *
     * @param heading       level-2 heading shown above the table, or null for none
     * @param colors        color theme of the table
     * @param headers       header row (shaded, bold); Word does not repeat it on the next page for side-by-side
     *                      tables, because they are nested in a layout table
     * @param rows          body rows; each row must have one cell per header
     * @param footer        optional bold total row at the bottom, or null for none
     * @param columnWeights relative column widths, e.g. {3, 1} makes the first column three times wider;
     *                      pass nothing for equal widths
     */
    public record Table(String heading, TableColors colors, List<String> headers, List<List<String>> rows,
                        List<String> footer, int... columnWeights) {
        public Table {
            columnWeights = columnWeights.clone(); // so a caller changing its array afterwards doesn't affect the table
        }
    }

    /**
     * Colors for a table (hex RGB, no '#').
     *
     * @param header     header row background (text on it is white, so keep it dark)
     * @param band       background of every second body row
     * @param footer     background of the total row
     */
    public record TableColors(String header, String band, String footer) {
        public static final TableColors BLUE = new TableColors("2E74B5", "EAF1FB", "D9E2F3");
        public static final TableColors GREEN = new TableColors("548235", "EEF5E9", "D5E8CC");
        public static final TableColors ORANGE = new TableColors("C55A11", "FCEFE6", "F8CBAD");
        public static final TableColors PURPLE = new TableColors("7030A0", "F3ECF8", "E1D1EE");
        public static final TableColors TEAL = new TableColors("1F7A7A", "E8F4F4", "C7E4E4");
        public static final TableColors BURGUNDY = new TableColors("9E2A3F", "FAEBEE", "F0CDD4");
        public static final TableColors GOLD = new TableColors("8C6D00", "FFF8E1", "FFE699");
    }

    private final boolean rightToLeft;
    private final boolean landscape;
    private final int contentWidth; // page width minus the margins
    private final StringBuilder body = new StringBuilder();

    /** @param rightToLeft true for Hebrew documents (paragraphs read right to left) */
    public DocxWriter(boolean rightToLeft) {
        this(rightToLeft, false);
    }

    /**
     * @param rightToLeft true for Hebrew documents (paragraphs read right to left)
     * @param landscape   true for landscape pages, which fit tables with many columns
     */
    public DocxWriter(boolean rightToLeft, boolean landscape) {
        this.rightToLeft = rightToLeft;
        this.landscape = landscape;
        this.contentWidth = (landscape ? PAGE_LONG_SIDE : PAGE_SHORT_SIDE) - 2 * MARGIN;
    }

    public DocxWriter addTitle(String text) {
        return addStyledParagraph("Title", text, false);
    }

    /** Adds a heading; level must be 1, 2 or 3. */
    public DocxWriter addHeading(int level, String text) {
        return addStyledParagraph(headingStyle(level), text, false);
    }

    /** Same as {@link #addHeading(int, String)}, but the heading starts a new page. */
    public DocxWriter addHeadingOnNewPage(int level, String text) {
        return addStyledParagraph(headingStyle(level), text, true);
    }

    private static String headingStyle(int level) {
        if (level < 1 || level > 3) {
            throw new IllegalArgumentException("Heading level must be 1, 2 or 3, got " + level);
        }
        return "Heading" + level;
    }

    /** Adds a normal paragraph. Line breaks (\n) inside the text are kept. */
    public DocxWriter addParagraph(String text) {
        return addStyledParagraph(null, text, false);
    }

    public DocxWriter addPageBreak() {
        body.append("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>");
        return this;
    }

    /**
     * Adds a table spanning the page width.
     *
     * @param headers       header row (shaded, bold, repeated at the top of each page)
     * @param rows          body rows; each row must have one cell per header
     * @param footer        optional bold total row at the bottom, or null for none
     * @param columnWeights relative column widths, e.g. {3, 1} makes the first column three times wider;
     *                      pass nothing for equal widths
     */
    public DocxWriter addTable(List<String> headers, List<List<String>> rows, List<String> footer, int... columnWeights) {
        return addTable(TableColors.BLUE, headers, rows, footer, columnWeights);
    }

    /** Same as {@link #addTable(List, List, List, int...)} with a chosen color theme. */
    public DocxWriter addTable(TableColors colors, List<String> headers, List<List<String>> rows, List<String> footer,
                               int... columnWeights) {
        appendTable(new Table(null, colors, headers, rows, footer, columnWeights), contentWidth);
        // Word needs a paragraph between a table and whatever follows it; keep it small.
        appendSpacerParagraph();
        return this;
    }

    /**
     * Adds a table only as wide as its contents: each column is as wide as its longest value, or the longest word
     * of its header (headers wrap between words), so short columns stay narrow. If that is wider than the page,
     * only the widest columns are narrowed (their long values wrap), so short columns such as numbers stay whole.
     *
     * @param heading optional level-2 heading above the table, or null for none
     * @param footer  optional bold total row at the bottom, or null for none
     */
    public DocxWriter addFittedTable(String heading, TableColors colors, List<String> headers,
                                     List<List<String>> rows, List<String> footer) {
        return addFittedTable(heading, colors, headers, rows, footer, List.of());
    }

    /**
     * Same as {@link #addFittedTable(String, TableColors, List, List, List)}, but the table then fills the page
     * width: the space left over is shared equally by the given columns (0-based), the rest keep their fitted width.
     */
    public DocxWriter addFittedTable(String heading, TableColors colors, List<String> headers,
                                     List<List<String>> rows, List<String> footer, List<Integer> stretchColumns) {
        int[] widths = new int[headers.size()];
        for (int c = 0; c < widths.length; c++) {
            for (String word : headers.get(c).split("\\s+")) {
                widths[c] = Math.max(widths[c], textWidth(word, true));
            }
            for (List<String> row : rows) {
                for (String line : row.get(c).split("\n")) {
                    widths[c] = Math.max(widths[c], textWidth(line, false));
                }
            }
            if (footer != null) {
                widths[c] = Math.max(widths[c], textWidth(footer.get(c), true));
            }
            widths[c] += FITTED_CELL_PADDING;
        }
        capToWidth(widths, contentWidth);
        int leftOver = contentWidth - Arrays.stream(widths).sum();
        for (int i = 0; i < stretchColumns.size() && leftOver > 0; i++) {
            // The first columns get any remainder of the division, so the total is exactly the page width.
            int share = leftOver / stretchColumns.size() + (i < leftOver % stretchColumns.size() ? 1 : 0);
            widths[stretchColumns.get(i)] += share;
        }
        if (heading != null) {
            addStyledParagraph("Heading2", heading, false);
        }
        appendTable(new Table(null, colors, headers, rows, footer, widths), Arrays.stream(widths).sum());
        appendSpacerParagraph();
        return this;
    }

    // If the columns are wider than maxTotal together, lowers the widest ones to a common cap, the largest cap
    // that makes them fit. Columns narrower than the cap keep their width.
    private static void capToWidth(int[] widths, int maxTotal) {
        if (Arrays.stream(widths).sum() <= maxTotal) {
            return;
        }
        int low = 0;
        int high = Arrays.stream(widths).max().orElse(0);
        while (low < high) {
            int cap = (low + high + 1) / 2;
            int total = Arrays.stream(widths).map(w -> Math.min(w, cap)).sum();
            if (total <= maxTotal) {
                low = cap;
            } else {
                high = cap - 1;
            }
        }
        for (int c = 0; c < widths.length; c++) {
            widths[c] = Math.min(widths[c], low);
        }
    }

    // Width of one line of table text in twips, measured with the table font.
    private static int textWidth(String text, boolean bold) {
        Font font = bold ? TABLE_FONT_BOLD : TABLE_FONT_PLAIN;
        double points = font.getStringBounds(text, FONT_RENDER_CONTEXT).getWidth();
        return (int) Math.ceil(points * 20); // 20 twips per point
    }

    /**
     * Puts two tables next to each other, each half the page wide. The first is on the reading-direction start
     * side (the right in a right-to-left document), the second on the other side.
     *
     * @param second the other half, or null to leave it empty
     */
    public DocxWriter addSideBySide(Table first, Table second) {
        int cellWidth = contentWidth / 2;
        String cellMargin = String.valueOf(SIDE_BY_SIDE_GAP / 2);

        // An invisible two-column table, with one table nested in each cell.
        body.append("<w:tbl><w:tblPr>");
        if (rightToLeft) {
            body.append("<w:bidiVisual/>");
        }
        body.append("<w:tblW w:w=\"").append(contentWidth).append("\" w:type=\"dxa\"/>")
                .append("<w:tblBorders>")
                .append(noBorder("top")).append(noBorder("left")).append(noBorder("bottom")).append(noBorder("right"))
                .append(noBorder("insideH")).append(noBorder("insideV"))
                .append("</w:tblBorders>")
                .append("<w:tblLayout w:type=\"fixed\"/>")
                .append("<w:tblCellMar><w:left w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"0\" w:type=\"dxa\"/></w:tblCellMar>")
                .append("</w:tblPr><w:tblGrid>")
                .append("<w:gridCol w:w=\"").append(cellWidth).append("\"/>")
                .append("<w:gridCol w:w=\"").append(cellWidth).append("\"/>")
                .append("</w:tblGrid><w:tr>");
        for (Table table : Arrays.asList(first, second)) {
            // Half the gap on each side of each cell keeps the two halves equal; the outer edges move in a little.
            body.append("<w:tc><w:tcPr><w:tcW w:w=\"").append(cellWidth).append("\" w:type=\"dxa\"/>")
                    .append("<w:tcMar><w:left w:w=\"").append(cellMargin).append("\" w:type=\"dxa\"/>")
                    .append("<w:right w:w=\"").append(cellMargin).append("\" w:type=\"dxa\"/></w:tcMar>")
                    .append("</w:tcPr>");
            if (table != null) {
                if (table.heading() != null) {
                    addStyledParagraph("Heading2", table.heading(), false);
                }
                appendTable(table, cellWidth - SIDE_BY_SIDE_GAP);
            }
            // A table cell must end with a paragraph.
            appendSpacerParagraph();
            body.append("</w:tc>");
        }
        body.append("</w:tr></w:tbl>");
        appendSpacerParagraph();
        return this;
    }

    private void appendTable(Table table, int tableWidth) {
        List<String> headers = table.headers();
        List<List<String>> rows = table.rows();
        TableColors colors = table.colors();
        int columns = headers.size();
        int[] widths = columnWidths(tableWidth, columns, table.columnWeights());

        body.append("<w:tbl><w:tblPr>");
        if (rightToLeft) {
            body.append("<w:bidiVisual/>"); // first column on the right
        }
        body.append("<w:tblW w:w=\"").append(tableWidth).append("\" w:type=\"dxa\"/>")
                .append("<w:tblBorders>")
                .append(border("top")).append(border("left")).append(border("bottom")).append(border("right"))
                .append(border("insideH")).append(border("insideV"))
                .append("</w:tblBorders>")
                .append("<w:tblLayout w:type=\"fixed\"/>")
                .append("<w:tblCellMar><w:left w:w=\"80\" w:type=\"dxa\"/><w:right w:w=\"80\" w:type=\"dxa\"/></w:tblCellMar>")
                .append("</w:tblPr><w:tblGrid>");
        for (int width : widths) {
            body.append("<w:gridCol w:w=\"").append(width).append("\"/>");
        }
        body.append("</w:tblGrid>");

        addTableRow(headers, widths, colors.header(), "FFFFFF", true, true);
        for (int i = 0; i < rows.size(); i++) {
            List<String> row = rows.get(i);
            if (row.size() != columns) {
                throw new IllegalArgumentException("Row " + i + " has " + row.size() + " cells, expected " + columns);
            }
            addTableRow(row, widths, i % 2 == 1 ? colors.band() : null, null, false, false);
        }
        if (table.footer() != null) {
            addTableRow(table.footer(), widths, colors.footer(), null, true, false);
        }
        body.append("</w:tbl>");
    }

    private void appendSpacerParagraph() {
        body.append("<w:p><w:pPr><w:spacing w:after=\"0\"/></w:pPr></w:p>");
    }

    /** Writes the document to the given path, creating parent folders if needed. */
    public void save(Path path) throws IOException {
        new FileIOHandler().ensureParentDirectories(path);
        try (OutputStream out = Files.newOutputStream(path);
             ZipOutputStream zip = new ZipOutputStream(out)) {
            putEntry(zip, "[Content_Types].xml", CONTENT_TYPES);
            putEntry(zip, "_rels/.rels", ROOT_RELS);
            putEntry(zip, "word/_rels/document.xml.rels", DOCUMENT_RELS);
            putEntry(zip, "word/styles.xml", stylesXml());
            putEntry(zip, "word/document.xml", documentXml());
        }
    }

    // ---------- XML building ----------

    private DocxWriter addStyledParagraph(String styleId, String text, boolean newPage) {
        body.append("<w:p><w:pPr>");
        if (styleId != null) {
            body.append("<w:pStyle w:val=\"").append(styleId).append("\"/>");
        }
        if (newPage) {
            body.append("<w:pageBreakBefore/>");
        }
        if (rightToLeft) {
            body.append("<w:bidi/>");
        }
        body.append("</w:pPr>");

        String[] lines = (text == null ? "" : text).split("\n", -1);
        for (int i = 0; i < lines.length; i++) {
            body.append("<w:r>");
            if (rightToLeft) {
                body.append("<w:rPr><w:rtl/></w:rPr>");
            }
            if (i > 0) {
                body.append("<w:br/>");
            }
            body.append("<w:t xml:space=\"preserve\">").append(escapeXml(lines[i])).append("</w:t></w:r>");
        }
        body.append("</w:p>");
        return this;
    }

    private void addTableRow(List<String> cells, int[] widths, String fill, String textColor,
                             boolean bold, boolean isHeader) {
        body.append("<w:tr><w:trPr><w:cantSplit/>");
        if (isHeader) {
            body.append("<w:tblHeader/>");
        }
        body.append("</w:trPr>");

        for (int c = 0; c < cells.size(); c++) {
            body.append("<w:tc><w:tcPr><w:tcW w:w=\"").append(widths[c]).append("\" w:type=\"dxa\"/>");
            if (fill != null) {
                body.append("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"").append(fill).append("\"/>");
            }
            body.append("<w:vAlign w:val=\"center\"/></w:tcPr>");

            // First column follows the reading direction; the other columns (numbers) are centered.
            body.append("<w:p><w:pPr>");
            if (rightToLeft) {
                body.append("<w:bidi/>");
            }
            body.append("<w:spacing w:before=\"30\" w:after=\"30\" w:line=\"240\" w:lineRule=\"auto\"/>");
            if (c > 0) {
                body.append("<w:jc w:val=\"center\"/>");
            }
            body.append("</w:pPr>");

            // Line breaks (\n) inside a cell are kept.
            String[] lines = (cells.get(c) == null ? "" : cells.get(c)).split("\n", -1);
            for (int i = 0; i < lines.length; i++) {
                body.append("<w:r><w:rPr>");
                if (bold) {
                    body.append("<w:b/><w:bCs/>");
                }
                if (textColor != null) {
                    body.append("<w:color w:val=\"").append(textColor).append("\"/>");
                }
                body.append("<w:sz w:val=\"").append(TABLE_FONT_SIZE).append("\"/>")
                        .append("<w:szCs w:val=\"").append(TABLE_FONT_SIZE).append("\"/>");
                if (rightToLeft) {
                    body.append("<w:rtl/>");
                }
                body.append("</w:rPr>");
                if (i > 0) {
                    body.append("<w:br/>");
                }
                body.append("<w:t xml:space=\"preserve\">").append(escapeXml(lines[i])).append("</w:t></w:r>");
            }
            body.append("</w:p></w:tc>");
        }
        body.append("</w:tr>");
    }

    private static int[] columnWidths(int tableWidth, int columns, int... weights) {
        if (weights.length != 0 && weights.length != columns) {
            throw new IllegalArgumentException("Expected " + columns + " column weights, got " + weights.length);
        }
        int totalWeight = 0;
        for (int c = 0; c < columns; c++) {
            totalWeight += weights.length == 0 ? 1 : weights[c];
        }
        int[] widths = new int[columns];
        for (int c = 0; c < columns; c++) {
            widths[c] = tableWidth * (weights.length == 0 ? 1 : weights[c]) / totalWeight;
        }
        return widths;
    }

    private static String border(String side) {
        return "<w:" + side + " w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"" + TABLE_BORDER + "\"/>";
    }

    private static String noBorder(String side) {
        return "<w:" + side + " w:val=\"nil\"/>";
    }

    private String documentXml() {
        String pageSize = landscape
                ? "<w:pgSz w:w=\"" + PAGE_LONG_SIDE + "\" w:h=\"" + PAGE_SHORT_SIDE + "\" w:orient=\"landscape\"/>"
                : "<w:pgSz w:w=\"" + PAGE_SHORT_SIDE + "\" w:h=\"" + PAGE_LONG_SIDE + "\"/>";
        String margin = String.valueOf(MARGIN);
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body>" + body
                + "<w:sectPr>" + pageSize
                + "<w:pgMar w:top=\"" + margin + "\" w:right=\"" + margin + "\" w:bottom=\"" + margin + "\" w:left=\""
                + margin + "\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>"
                + (rightToLeft ? "<w:bidi/>" : "")
                + "</w:sectPr></w:body></w:document>";
    }

    private String stylesXml() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:docDefaults><w:rPrDefault><w:rPr>"
                + "<w:rFonts w:ascii=\"" + FONT + "\" w:hAnsi=\"" + FONT + "\" w:cs=\"" + FONT + "\" w:eastAsia=\"" + FONT + "\"/>"
                + "<w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/>"
                + (rightToLeft ? "<w:lang w:val=\"en-US\" w:bidi=\"he-IL\"/>" : "<w:lang w:val=\"en-US\"/>")
                + "</w:rPr></w:rPrDefault>"
                + "<w:pPrDefault><w:pPr><w:spacing w:after=\"120\" w:line=\"276\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault>"
                + "</w:docDefaults>"
                + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>"
                + headingStyle("Title", "Title", -1, 56, "17365D", 0, 300)
                + headingStyle("Heading1", "heading 1", 0, 36, "2E74B5", 360, 120)
                + headingStyle("Heading2", "heading 2", 1, 30, "2E74B5", 240, 80)
                + headingStyle("Heading3", "heading 3", 2, 26, "1F4D78", 200, 60)
                + "</w:styles>";
    }

    /**
     * @param outlineLevel 0-based outline level (-1 for none); this is what makes Word treat it as a heading
     * @param halfPoints   font size in half-points (36 = 18pt)
     */
    private static String headingStyle(String id, String name, int outlineLevel, int halfPoints,
                                       String color, int spaceBefore, int spaceAfter) {
        return "<w:style w:type=\"paragraph\" w:styleId=\"" + id + "\">"
                + "<w:name w:val=\"" + name + "\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/>"
                + "<w:uiPriority w:val=\"9\"/><w:qFormat/>"
                + "<w:pPr><w:keepNext/><w:keepLines/>"
                + "<w:spacing w:before=\"" + spaceBefore + "\" w:after=\"" + spaceAfter + "\"/>"
                + (outlineLevel >= 0 ? "<w:outlineLvl w:val=\"" + outlineLevel + "\"/>" : "")
                + "</w:pPr>"
                + "<w:rPr><w:b/><w:bCs/><w:color w:val=\"" + color + "\"/>"
                + "<w:sz w:val=\"" + halfPoints + "\"/><w:szCs w:val=\"" + halfPoints + "\"/></w:rPr>"
                + "</w:style>";
    }

    /** Escapes XML special characters and drops characters that are not allowed in XML. */
    private static String escapeXml(String text) {
        StringBuilder sb = new StringBuilder(text.length());
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            switch (c) {
                case '&' -> sb.append("&amp;");
                case '<' -> sb.append("&lt;");
                case '>' -> sb.append("&gt;");
                case '"' -> sb.append("&quot;");
                default -> {
                    if (c >= 0x20 || c == '\t') {
                        sb.append(c);
                    }
                }
            }
        }
        return sb.toString();
    }

    private static void putEntry(ZipOutputStream zip, String name, String content) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(content.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }

    private static final String CONTENT_TYPES =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                    + "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
                    + "</Types>";

    private static final String ROOT_RELS =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                    + "</Relationships>";

    private static final String DOCUMENT_RELS =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
                    + "</Relationships>";
}
