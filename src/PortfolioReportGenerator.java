import csv.CsvLoader;
import csv.TransactionStore;
import report.ReportWriter;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.net.URI;

public class PortfolioReportGenerator {

    private static final String INPUT_DIRECTORY = "transaction_files";
    private static final String OUTPUT_FILE = "portfolio-report.html";

    public static void main(String[] args) throws IOException {
        ensureInputDirectoryExists();
        System.out.println("=== Portfolio Report Generator ===");

        TransactionStore store = new TransactionStore();
        int filesProcessed = CsvLoader.readAllTransactionFiles(INPUT_DIRECTORY, store);

        if (filesProcessed == 0) {
            System.out.println("No CSV files found in '" + INPUT_DIRECTORY + "'.");
            System.out.println("Place one or more transaction exports in that folder and run again.");
        } else {
            System.out.println("Loaded " + filesProcessed + " CSV file(s) from '" + INPUT_DIRECTORY + "'.");
        }

        System.out.println("Generating report...");

        ReportWriter.writeHtmlReport(store, OUTPUT_FILE);

        System.out.println("Done! Report generated: " + OUTPUT_FILE);
        openGeneratedReport();
    }

    private static void ensureInputDirectoryExists() {
        File directory = new File(INPUT_DIRECTORY);
        if (!directory.exists() && directory.mkdirs()) {
            System.out.println("Created input folder: " + INPUT_DIRECTORY);
        }
    }

    private static void openGeneratedReport() {
        File reportFile = new File(OUTPUT_FILE).getAbsoluteFile();
        if (!reportFile.exists()) {
            return;
        }

        try {
            if (Desktop.isDesktopSupported()) {
                Desktop desktop = Desktop.getDesktop();
                URI reportUri = reportFile.toURI();
                if (desktop.isSupported(Desktop.Action.BROWSE)) {
                    desktop.browse(reportUri);
                    return;
                }
                if (desktop.isSupported(Desktop.Action.OPEN)) {
                    desktop.open(reportFile);
                    return;
                }
            }
            System.out.println("Open this file manually in your browser: " + reportFile.getAbsolutePath());
        } catch (Exception e) {
            System.out.println("Could not open browser automatically: " + e.getMessage());
            System.out.println("Open this file manually: " + reportFile.getAbsolutePath());
        }
    }
}