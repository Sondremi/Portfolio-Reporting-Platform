package tests;

import csv.CsvLoader;
import csv.TransactionStore;
import model.Security;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

/**
 * Golden-master (characterization) test for the transaction-derived calculation
 * logic in the local Java report engine. Loads the tracked example transactions
 * and asserts the per-security units, average cost, realized gain, dividends and
 * realized cost basis against known-good values.
 *
 * These values depend only on the CSV transactions (FIFO cost basis, realized
 * P&L, dividends) — NOT on any network calls — so the test is fully
 * deterministic and runs offline.
 *
 * This guards the Java side of the intentional Java/JS split. The JS engine
 * lives under public/ (not version-controlled here), so it is verified
 * separately; keep both in sync per CLAUDE.md.
 *
 * Run:
 *   javac -d out src/*.java src/*&#47;*.java
 *   java -cp out tests.CalculationGoldenMasterTest
 */
public class CalculationGoldenMasterTest {

    private static final double EPS = 1e-4;
    private static final String DEFAULT_FIXTURE = "transaction_files/transactions_example.csv";

    private static int failures = 0;

    private static final class Expected {
        final double units, avgCost, realizedGain, dividends, realizedCostBasis;
        Expected(double units, double avgCost, double realizedGain, double dividends, double realizedCostBasis) {
            this.units = units;
            this.avgCost = avgCost;
            this.realizedGain = realizedGain;
            this.dividends = dividends;
            this.realizedCostBasis = realizedCostBasis;
        }
    }

    // Golden values captured from the example portfolio. A change here means a
    // change in calculation behavior — verify it is intentional before updating.
    private static Map<String, Expected> expectedByName() {
        Map<String, Expected> m = new HashMap<>();
        m.put("Apple Inc.",            new Expected(660, 220.82424242,   78.0,  17.25,    1098.0));
        m.put("Equinor ASA",           new Expected(680, 299.33764706, 5631.0, 1074.0,   59771.0));
        m.put("KLP AksjeNorden",       new Expected(320, 751.74416667, -1156.0,   0.0,  281200.0));
        m.put("Microsoft Corp.",       new Expected(396, 343.69696970, -176.0,   4.5,    2888.0));
        m.put("Skagen Global",         new Expected(51, 3335.54823529, 64764.0,   0.0, 1564000.0));
        m.put("Tesla Inc.",            new Expected(303, 293.82343234,    8.0,   0.0,    1794.0));
        m.put("Vanguard S&P 500 ETF",  new Expected(155, 508.05591398,  644.0,   0.0,   11200.0));
        return m;
    }

    private static final double EXPECTED_CASH = 352931.75;

    public static void main(String[] args) throws IOException {
        String fixture = args.length > 0 ? args[0] : DEFAULT_FIXTURE;

        // CsvLoader skips files whose name contains "example", so copy the fixture
        // into a temp directory under a neutral name before loading.
        Path tempDir = Files.createTempDirectory("golden-master-test");
        Path target = tempDir.resolve("golden.csv");
        Files.copy(Path.of(fixture), target);

        TransactionStore store = new TransactionStore();
        int files = CsvLoader.readAllTransactionFiles(tempDir.toString(), store);
        check("loaded exactly one CSV file", files == 1, String.valueOf(files));

        Map<String, Security> byName = new HashMap<>();
        for (Security s : store.getSecurities()) {
            byName.put(s.getName(), s);
        }

        Map<String, Expected> expected = expectedByName();
        check("security count", store.getSecurities().size() == expected.size(),
                store.getSecurities().size() + " (expected " + expected.size() + ")");

        for (Map.Entry<String, Expected> entry : expected.entrySet()) {
            String name = entry.getKey();
            Expected e = entry.getValue();
            Security s = byName.get(name);
            if (s == null) {
                check(name + " present", false, "missing");
                continue;
            }
            assertClose(name + " units", e.units, s.getUnitsOwned());
            assertClose(name + " avgCost", e.avgCost, s.getAverageCost());
            assertClose(name + " realizedGain", e.realizedGain, s.getRealizedGain());
            assertClose(name + " dividends", e.dividends, s.getDividends());
            assertClose(name + " realizedCostBasis", e.realizedCostBasis, s.getRealizedCostBasis());
        }

        assertClose("current cash holdings", EXPECTED_CASH, store.getCurrentCashHoldings());

        // Cleanup
        Files.deleteIfExists(target);
        Files.deleteIfExists(tempDir);

        if (failures == 0) {
            System.out.println("\nGOLDEN MASTER: PASS");
        } else {
            System.out.println("\nGOLDEN MASTER: FAIL (" + failures + " check(s) failed)");
            System.exit(1);
        }
    }

    private static void assertClose(String label, double expected, double actual) {
        boolean ok = Math.abs(expected - actual) <= EPS;
        check(label, ok, String.format(java.util.Locale.US, "%.8f (expected %.8f)", actual, expected));
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) {
            System.out.println("  ok   " + label);
        } else {
            failures++;
            System.out.println("  FAIL " + label + " -> " + detail);
        }
    }
}
