package utils;

import java.io.IOException;
import java.nio.charset.MalformedInputException;
import java.nio.file.FileSystemException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Reads a sheet exported from Excel as "CSV UTF-8" into records of trimmed fields. Shared by the workers and the
 * leads parsers.
 */
public final class CsvReader {
    // Excel starts a UTF-8 CSV with this invisible character.
    private static final String BYTE_ORDER_MARK = String.valueOf((char) 0xFEFF);

    private CsvReader() {
    }

    /** Returns every record of the file, the header row first. A file saved as plain CSV (not UTF-8) is rejected. */
    public static List<List<String>> readRecords(Path csvFile) throws IOException, UserFacingException {
        return parseCsv(readText(csvFile));
    }

    /** The file's first row, or an empty list if the file is empty. */
    public static List<String> readHeader(Path csvFile) throws IOException, UserFacingException {
        List<List<String>> records = readRecords(csvFile);
        return records.isEmpty() ? List.of() : records.getFirst();
    }

    /** The row with exactly the given number of cells: shorter rows are filled with "", longer ones cut. */
    public static List<String> padded(List<String> row, int size) {
        List<String> result = new ArrayList<>(row.subList(0, Math.min(row.size(), size)));
        while (result.size() < size) {
            result.add("");
        }
        return result;
    }

    /** A header without double or single quotes, so חוו"ד and חווד compare equal. */
    public static String withoutQuotes(String header) {
        return header.replace("\"", "").replace("'", "");
    }

    private static String readText(Path csvFile) throws IOException, UserFacingException {
        try {
            return new FileIOHandler().readString(csvFile).replace(BYTE_ORDER_MARK, "");
        } catch (MalformedInputException e) {
            throw new UserFacingException(csvFile + " is not saved as UTF-8, so its Hebrew text cannot be read."
                    + " In Excel, save the sheet again with File > Save As > \"CSV UTF-8 (Comma delimited)\""
                    + " (not the plain \"CSV (Comma delimited)\").", e);
        } catch (FileSystemException e) {
            throw new UserFacingException("Cannot read " + csvFile + ": " + Objects.requireNonNullElse(e.getReason(), e.toString())
                    + ". If it is open in another program, close it and run again.", e);
        }
    }

    /**
     * Splits CSV text into records of trimmed fields. Commas and line breaks inside double quotes are part of
     * the field (a free-text cell can hold several lines), and a doubled quote inside quotes ("חוו""ד") is read
     * as a single quote character.
     */
    private static List<List<String>> parseCsv(String text) {
        List<List<String>> records = new ArrayList<>();
        List<String> fields = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (inQuotes) {
                if (c == '"' && i + 1 < text.length() && text.charAt(i + 1) == '"') {
                    current.append('"');
                    i++;
                } else if (c == '"') {
                    inQuotes = false;
                } else if (c != '\r') {
                    current.append(c);
                }
            } else if (c == '"') {
                inQuotes = true;
            } else if (c == ',') {
                fields.add(current.toString().strip());
                current.setLength(0);
            } else if (c == '\n') {
                fields.add(current.toString().strip());
                current.setLength(0);
                records.add(fields);
                fields = new ArrayList<>();
            } else if (c != '\r') {
                current.append(c);
            }
        }
        // The last record, unless the file ends with a line break.
        if (!current.isEmpty() || !fields.isEmpty()) {
            fields.add(current.toString().strip());
            records.add(fields);
        }
        return records;
    }
}
