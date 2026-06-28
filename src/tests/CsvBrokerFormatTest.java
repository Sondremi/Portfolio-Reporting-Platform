package tests;

import csv.CsvLoader;
import csv.TransactionStore;
import model.Security;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Broker-format tolerance test. The same logical portfolio (two buys, one sell,
 * one dividend) is expressed in two different export layouts:
 *   - semicolon-delimited with common Nordnet header spellings,
 *   - tab-delimited with alternate column names (Handelsdato, Kurs per andel,
 *     Handelsbeløp, Anskaffelsesverdi).
 * Both must parse to identical units, realized gain and dividends. A header
 * missing the required columns must be rejected outright.
 *
 * Offline + deterministic.
 *
 * Run: java -cp out tests.CsvBrokerFormatTest
 */
public class CsvBrokerFormatTest {

    private static final double EPS = 1e-4;
    private static int failures = 0;

    public static void main(String[] args) throws IOException {
        // Variant A: semicolon delimiter, Nordnet-style headers.
        String semicolon = String.join("\n",
            "Bokføringsdag;Transaksjonstype;Verdipapir;ISIN;Antall;Kurs;Beløp;Valuta;Kjøpsverdi;Resultat",
            "2025-01-02;KJØPT;TestCo;NO0010000001;100;10;-1000;NOK;1000;0",
            "2025-02-02;KJØPT;TestCo;NO0010000001;50;12;-600;NOK;600;0",
            "2025-03-02;SALG;TestCo;NO0010000001;30;15;450;NOK;300;150",
            "2025-04-02;UTBYTTE;TestCo;NO0010000001;0;0;25;NOK;0;0");

        // Variant B: tab delimiter, alternate header spellings + a Konto column.
        String tab = String.join("\n",
            "Handelsdato\tKonto\tTransaksjonstype\tVerdipapir\tISIN\tAntall\tKurs per andel\tHandelsbeløp\tValuta\tAnskaffelsesverdi\tResultat",
            "2025-01-02\t1\tKJØPT\tTestCo\tNO0010000001\t100\t10\t-1000\tNOK\t1000\t0",
            "2025-02-02\t1\tKJØPT\tTestCo\tNO0010000001\t50\t12\t-600\tNOK\t600\t0",
            "2025-03-02\t1\tSALG\tTestCo\tNO0010000001\t30\t15\t450\tNOK\t300\t150",
            "2025-04-02\t1\tUTBYTTE\tTestCo\tNO0010000001\t0\t0\t25\tNOK\t0\t0");

        Security a = loadSingle("semicolon.csv", semicolon);
        Security b = loadSingle("tab.csv", tab);

        check("semicolon variant parsed", a != null, "null");
        check("tab variant parsed", b != null, "null");

        if (a != null) {
            assertClose("A units", 120.0, a.getUnitsOwned());
            assertClose("A realized gain", 150.0, a.getRealizedGain());
            assertClose("A dividends", 25.0, a.getDividends());
        }
        if (a != null && b != null) {
            assertClose("units match across formats", a.getUnitsOwned(), b.getUnitsOwned());
            assertClose("realized match across formats", a.getRealizedGain(), b.getRealizedGain());
            assertClose("dividends match across formats", a.getDividends(), b.getDividends());
        }

        // Negative: a header without Transaksjonstype lacks required columns.
        String badHeader = String.join("\n",
            "Dato;Verdipapir;Antall;Kurs",
            "2025-01-02;TestCo;100;10");
        TransactionStore store = new TransactionStore();
        Path dir = Files.createTempDirectory("broker-bad");
        Path csv = dir.resolve("bad.csv");
        Files.writeString(csv, badHeader);
        int processed = CsvLoader.readAllTransactionFiles(dir.toString(), store);
        check("header missing required columns is rejected",
                processed == 0 && store.getSecurities().isEmpty(),
                "processed=" + processed + " securities=" + store.getSecurities().size());
        Files.deleteIfExists(csv);
        Files.deleteIfExists(dir);

        if (failures == 0) {
            System.out.println("\nCSV BROKER FORMAT TEST: PASS");
        } else {
            System.out.println("\nCSV BROKER FORMAT TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static Security loadSingle(String fileName, String content) throws IOException {
        Path dir = Files.createTempDirectory("broker-fmt");
        Path csv = dir.resolve(fileName);
        Files.writeString(csv, content);
        TransactionStore store = new TransactionStore();
        CsvLoader.readAllTransactionFiles(dir.toString(), store);
        Security s = store.getSecurities().stream()
                .filter(x -> "TestCo".equals(x.getName())).findFirst().orElse(null);
        Files.deleteIfExists(csv);
        Files.deleteIfExists(dir);
        return s;
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
