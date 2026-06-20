package tests;

import csv.CsvLoader;
import csv.TransactionStore;
import model.Security;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

/**
 * Verifies cancelled-order bookkeeping (e.g. a corrected fund NAV re-booking):
 * a "Makulert kjøp" reversal and the cancelled original buy (both carrying a
 * Makuleringsdato) must be ignored, leaving only the valid/corrected buys.
 *
 * Offline + deterministic (no network — Security enrichment is deferred).
 *
 * Run: java -cp out tests.CancellationTransactionTest
 */
public class CancellationTransactionTest {

    private static final double EPS = 1e-4;
    private static int failures = 0;

    // Nordnet/Handelsbanken-style header; Makuleringsdato is column index 22.
    private static final String HEADER = String.join(";",
        "Bokføringsdag","Handelsdag","Oppgjørsdag","Portefølje","Transaksjonstype","Verdipapir","ISIN",
        "Antall","Kurs","Rente","Totale Avgifter","Valuta","Beløp","Valuta","Kjøpsverdi","Valuta",
        "Resultat","Valuta","Totalt antall","Saldo","Vekslingskurs","Transaksjonstekst","Makuleringsdato",
        "Sluttseddelnummer","Verifikationsnummer","Kurtasje","Valuta","Valutakurs","Innledende rente");

    private static String row(String date, String type, String antall, String kurs,
                              String belop, String kjopsverdi, String makdato, String id) {
        String[] f = new String[29];
        Arrays.fill(f, "");
        f[0] = date; f[1] = date; f[2] = date; f[3] = "111";
        f[4] = type; f[5] = "TestCo"; f[6] = "NO0000000001";
        f[7] = antall; f[8] = kurs; f[10] = "0"; f[11] = "NOK";
        f[12] = belop; f[13] = "NOK"; f[14] = kjopsverdi; f[15] = "NOK";
        f[16] = "0"; f[17] = "NOK"; f[19] = "0"; f[22] = makdato; f[24] = id; f[25] = "0"; f[26] = "NOK";
        return String.join(";", f);
    }

    public static void main(String[] args) throws IOException {
        Path dir = Files.createTempDirectory("cancellation-test");
        Path csv = dir.resolve("cancellation.csv");
        String content = String.join("\n",
            HEADER,
            // valid buy: 100 units @ 10
            row("2025-01-10", "KJØPT", "100", "10.00", "-1000.00", "1000.00", "", "100001"),
            // wrong-NAV buy that was cancelled (carries a Makuleringsdato): 50 @ 12
            row("2025-05-13", "KJØPT", "50", "12.00", "-600.00", "600.00", "2025-05-20", "100002"),
            // the cancellation marker (also carries the Makuleringsdato)
            row("2025-05-13", "Makulert kjøp", "50", "12.00", "600.00", "", "2025-05-20", "100003"),
            // corrected buy at the right NAV (no Makuleringsdato): 50 @ 11
            row("2025-05-27", "KJØPT", "50", "11.00", "-550.00", "550.00", "", "100004"));
        Files.writeString(csv, content);

        TransactionStore store = new TransactionStore();
        CsvLoader.readAllTransactionFiles(dir.toString(), store);

        Security s = store.getSecurities().stream()
            .filter(x -> "TestCo".equals(x.getName())).findFirst().orElse(null);
        check("security present", s != null, String.valueOf(s));
        if (s != null) {
            // Only the valid (100) + corrected (50) buys count; cancelled + makulert ignored.
            assertClose("units owned", 150.0, s.getUnitsOwned());
            // Cost basis 1000 + 550 = 1550 over 150 units.
            assertClose("average cost", 1550.0 / 150.0, s.getAverageCost());
            assertClose("realized gain", 0.0, s.getRealizedGain());
        }

        Files.deleteIfExists(csv);
        Files.deleteIfExists(dir);

        if (failures == 0) {
            System.out.println("\nCANCELLATION TEST: PASS");
        } else {
            System.out.println("\nCANCELLATION TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void assertClose(String label, double expected, double actual) {
        boolean ok = Math.abs(expected - actual) <= EPS;
        check(label, ok, String.format(java.util.Locale.US, "%.6f (expected %.6f)", actual, expected));
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
