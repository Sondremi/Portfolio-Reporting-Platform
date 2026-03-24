import csv.CsvLoader;
import csv.TransactionStore;
import report.ReportWriter;

import java.io.File;
import java.io.IOException;

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
    }

    private static void ensureInputDirectoryExists() {
        File directory = new File(INPUT_DIRECTORY);
        if (!directory.exists() && directory.mkdirs()) {
            System.out.println("Created input folder: " + INPUT_DIRECTORY);
        }
    }
}