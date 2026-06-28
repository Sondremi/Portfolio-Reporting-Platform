package tests;

import csv.CsvLoader;
import csv.TransactionStore;
import report.AnnualPerformanceSummary;
import report.PortfolioCalculator;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Golden-master test for the transaction-derived fields of the annual report:
 * realized gain, dividends and their total for a given year. Rates are pinned to
 * 1.0 so the NOK figures equal the sum of native per-security sale gains and
 * dividends for that year, making the check deterministic and offline.
 *
 * Portfolio value, benchmark, analytics and Monte Carlo fields depend on market
 * prices / network history and are intentionally NOT asserted here.
 *
 * Run: java -cp out tests.AnnualGoldenMasterTest
 */
public class AnnualGoldenMasterTest {

    private static final double EPS = 1e-4;
    private static final String DEFAULT_FIXTURE = "transaction_files/transactions_example.csv";
    private static final int YEAR = 2025;

    // Golden values captured from the example portfolio (all FX rates = 1.0).
    private static final double EXPECTED_REALIZED = 69793.0;
    private static final double EXPECTED_DIVIDENDS = 1095.75;

    private static int failures = 0;

    public static void main(String[] args) throws IOException {
        String fixture = args.length > 0 ? args[0] : DEFAULT_FIXTURE;

        Path tempDir = Files.createTempDirectory("annual-golden-test");
        Path target = tempDir.resolve("golden.csv");
        Files.copy(Path.of(fixture), target);

        TransactionStore store = new TransactionStore();
        CsvLoader.readAllTransactionFiles(tempDir.toString(), store);

        Map<String, Double> rates = new LinkedHashMap<>();
        rates.put("NOK", 1.0);
        rates.put("USD", 1.0);
        rates.put("EUR", 1.0);

        AnnualPerformanceSummary summary =
                PortfolioCalculator.buildAnnualPerformanceSummary(store, rates, YEAR, "^OSEAX");

        check("year echoed", summary.year == YEAR, String.valueOf(summary.year));
        assertClose("realized gain (NOK)", EXPECTED_REALIZED, summary.realizedGainNok);
        assertClose("dividends (NOK)", EXPECTED_DIVIDENDS, summary.dividendsNok);
        assertClose("realized total (NOK)", EXPECTED_REALIZED + EXPECTED_DIVIDENDS, summary.realizedTotalNok);

        Files.deleteIfExists(target);
        Files.deleteIfExists(tempDir);

        if (failures == 0) {
            System.out.println("\nANNUAL GOLDEN MASTER: PASS");
        } else {
            System.out.println("\nANNUAL GOLDEN MASTER: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void assertClose(String label, double expected, double actual) {
        check(label, Math.abs(expected - actual) <= EPS,
                String.format(java.util.Locale.US, "%.4f (expected %.4f)", actual, expected));
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
