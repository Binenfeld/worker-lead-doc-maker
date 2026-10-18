import models.Business;
import models.WorkerSheet;
import utils.CsvReader;
import utils.FileIOHandler;
import utils.LeadParser;
import utils.PdfConverter;
import utils.ReportFactory;
import utils.UserFacingException;
import utils.WorkerParser;

import java.io.IOException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/**
 * Reads the newest unassigned-workers CSV and the newest worker-leads CSV in the input folder, writes the Hebrew
 * report of which workers can be offered to which business (.docx) next to them, and converts it to PDF with Word.
 *
 * <p>Usage: {@code java -jar WorkerLeadDocMaker.jar [inputFolder] [--debug]} (default folder
 * {@value #DEFAULT_INPUT_FOLDER}). {@code --debug} also prints every row read, and the stack trace of any error.
 *
 * <p>Any problem ends the program with exit code 1 and a message saying what went wrong and how to fix it.
 */
public class Main {
    // Where both sheets, exported from Excel as "CSV UTF-8", are dropped. Each is recognized by its header row.
    private static final String DEFAULT_INPUT_FOLDER = "C:\\workerleaddocmaker";

    private static final String DEBUG_FLAG = "--debug";

    private static final String SAVE_AS = " Open the workbook, select the sheet, and use File > Save As > \"CSV UTF-8"
            + " (Comma delimited)\" to save it in this folder.";

    // The two input CSVs.
    private record Inputs(Path workers, Path leads) {
    }

    public static void main(String[] args) {
        boolean debug = List.of(args).contains(DEBUG_FLAG);
        try {
            run(args, debug);
        } catch (UserFacingException e) {
            printError(e.getMessage(), e, debug);
            System.exit(1);
        } catch (Throwable e) {
            // Not one of the problems the program knows how to explain: a bug, or something unusual on this computer.
            printError("Unexpected error: " + describe(e), e, true);
            System.exit(1);
        }
    }

    private static void run(String[] args, boolean debug) throws Exception {
        List<String> folders = new ArrayList<>();
        for (String arg : args) {
            if (arg.equals(DEBUG_FLAG)) continue;
            if (arg.startsWith("-")) {
                throw new UserFacingException("Unknown option \"" + arg + "\". Options: " + DEBUG_FLAG + ".");
            }
            folders.add(arg);
        }
        if (folders.size() > 1) {
            throw new UserFacingException("Expected at most one input folder, got " + folders.size() + ": " + folders
                    + ". Put a folder name that contains spaces in quotes.");
        }
        Path inputFolder = Path.of(folders.isEmpty() ? DEFAULT_INPUT_FOLDER : folders.getFirst()).toAbsolutePath();
        if (!Files.isDirectory(inputFolder)) {
            throw new UserFacingException("The input folder " + inputFolder + " does not exist. Create it and save"
                    + " the unassigned-workers sheet and the worker-leads sheet there (Excel: File > Save As >"
                    + " CSV UTF-8), or pass a different folder.");
        }

        Inputs inputs = findInputs(inputFolder);
        System.out.println("Reading workers from " + inputs.workers());
        WorkerSheet workers = WorkerParser.parseWorkerCsv(inputs.workers());
        System.out.println("Read " + workers.rows().size() + " workers, " + workers.headers().size() + " columns");
        if (debug) {
            System.out.println(String.join(" | ", workers.headers()));
            workers.rows().forEach(row -> System.out.println(String.join(" | ", row)));
        }

        System.out.println("Reading leads from " + inputs.leads());
        List<Business> businesses = LeadParser.parseLeadCsv(inputs.leads());
        System.out.println("Read " + businesses.size() + " businesses, "
                + businesses.stream().mapToInt(Business::totalWanted).sum() + " workers asked for");
        if (debug) {
            for (Business business : businesses) {
                System.out.println(business.name() + " " + business.details());
                business.wanted().forEach((profession, wanted) ->
                        System.out.println("  " + ReportFactory.supplyLine(workers, profession, wanted)));
            }
        }

        String reportName = "שיבוץ עובדים לעסקים " + LocalDate.now();
        Path report = inputFolder.resolve(reportName + ".docx");
        try {
            ReportFactory.writeReport(workers, businesses, report);
        } catch (FileSystemException e) {
            throw new UserFacingException("Cannot write " + report + " - it is probably open in Word."
                    + " Close it and run again. (" + Objects.requireNonNullElse(e.getReason(), e.toString()) + ")", e);
        } catch (IOException e) {
            throw new UserFacingException("Cannot write " + report + ": " + describe(e)
                    + ". Check that the folder is writable and the disk is not full.", e);
        }
        System.out.println("Report written to " + report);

        Path pdf = inputFolder.resolve(reportName + ".pdf");
        try {
            PdfConverter.convert(report, pdf);
        } catch (IOException e) {
            throw new UserFacingException(e.getMessage(), e);
        }
        System.out.println("PDF written to " + pdf);
    }

    // The newest CSV of each kind, told apart by their header rows. Only files newer than the ones chosen are read,
    // so an old unreadable CSV in the folder does no harm.
    private static Inputs findInputs(Path inputFolder) throws UserFacingException, IOException {
        List<Path> csvs = new FileIOHandler().listFiles(inputFolder).stream()
                .filter(p -> {
                    String name = p.getFileName().toString();
                    return name.toLowerCase(Locale.ROOT).endsWith(".csv") && !name.startsWith("~$");
                })
                .sorted(Comparator.comparingLong((Path p) -> p.toFile().lastModified()).reversed())
                .toList();

        Path workers = null;
        Path leads = null;
        for (Path csv : csvs) {
            if (workers != null && leads != null) break;
            List<String> header = CsvReader.readHeader(csv);
            if (WorkerParser.looksLikeWorkers(header)) {
                if (workers == null) workers = csv;
            } else if (LeadParser.looksLikeLeads(header)) {
                if (leads == null) leads = csv;
            } else {
                System.out.println("Warning: " + csv.getFileName() + " is neither the unassigned-workers sheet nor"
                        + " the worker-leads sheet, so it is ignored (first row: " + String.join(", ", header) + ").");
            }
        }

        if (workers == null || leads == null) {
            boolean hasWorkbook = new FileIOHandler().newestFile(inputFolder, ".xlsb").isPresent()
                    || new FileIOHandler().newestFile(inputFolder, ".xlsx").isPresent();
            String workbookNote = hasWorkbook
                    ? " There is an Excel workbook in the folder, but the program reads CSV only." : "";
            if (workers == null && leads == null) {
                throw new UserFacingException("No unassigned-workers CSV and no worker-leads CSV in " + inputFolder + "."
                        + workbookNote + " Both are needed: save the \"פרטי מידע לא משוייכים\" sheet and the"
                        + " \"לידים עובדים\" sheet as CSV UTF-8 in this folder. For each one:" + SAVE_AS);
            }
            String missing = workers == null
                    ? "unassigned-workers CSV (\"פרטי מידע לא משוייכים\", with the columns "
                    + quoted(WorkerParser.requiredHeaders()) + ")"
                    : "worker-leads CSV (\"לידים עובדים\", with the columns " + quoted(LeadParser.requiredHeaders()) + ")";
            String found = workers != null
                    ? "the unassigned-workers CSV " + workers.getFileName()
                    : "the worker-leads CSV " + leads.getFileName();
            throw new UserFacingException("No " + missing + " in " + inputFolder + " (found only " + found + ")."
                    + workbookNote + SAVE_AS
                    + " If a CSV is there but was ignored, see the warning above: its first row must be the header.");
        }
        return new Inputs(workers, leads);
    }

    private static String quoted(List<String> headers) {
        return String.join(", ", headers.stream().map(h -> "\"" + h + "\"").toList());
    }

    private static void printError(String message, Throwable e, boolean withStackTrace) {
        System.out.flush();
        System.err.println();
        System.err.println("ERROR: " + message);
        if (withStackTrace) {
            System.err.println();
            e.printStackTrace();
        }
        System.err.flush();
    }

    // "message (ExceptionType)", or just the type when there is no message (e.g. a NullPointerException).
    private static String describe(Throwable e) {
        String type = e.getClass().getSimpleName();
        return e.getMessage() == null || e.getMessage().isBlank() ? type : e.getMessage() + " (" + type + ")";
    }
}
