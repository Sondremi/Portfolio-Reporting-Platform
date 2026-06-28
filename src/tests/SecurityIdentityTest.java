package tests;

import csv.TransactionStore;
import model.Security;

/**
 * Unit test for TransactionStore.getOrCreateSecurity identity rules — the
 * "security identity collision when ISIN is missing" risk. Verifies:
 *   - same ISIN (any case) dedupes to one instance, keeping the first name,
 *   - different ISIN never merges even with an identical name,
 *   - blank/null ISIN falls back to a normalized-name key,
 *   - blank name + blank ISIN collapses to a single UNKNOWN bucket.
 *
 * Offline + deterministic (Security market-data enrichment is deferred).
 *
 * Run: java -cp out tests.SecurityIdentityTest
 */
public class SecurityIdentityTest {

    private static int failures = 0;

    public static void main(String[] args) {
        TransactionStore store = new TransactionStore();

        // Same ISIN, different name/case -> one instance, first name wins.
        Security a1 = store.getOrCreateSecurity("Apple Inc.", "US0378331005");
        Security a2 = store.getOrCreateSecurity("APPLE INCORPORATED", "us0378331005");
        check("same ISIN dedupes to one instance", a1 == a2, "");
        check("first name retained", "Apple Inc.".equals(a1.getName()), a1.getName());

        // Different ISIN, identical name -> must NOT merge.
        Security b1 = store.getOrCreateSecurity("Shared Name", "NO0000000001");
        Security b2 = store.getOrCreateSecurity("Shared Name", "NO0000000002");
        check("distinct ISIN stays distinct", b1 != b2, "");

        // Blank/null ISIN -> keyed by normalized name (case + whitespace insensitive).
        Security c1 = store.getOrCreateSecurity("Cash Fund", "");
        Security c2 = store.getOrCreateSecurity("  cash   fund ", null);
        check("blank ISIN keys by normalized name", c1 == c2, "");
        Security c3 = store.getOrCreateSecurity("Other Fund", "");
        check("different blank-ISIN names stay distinct", c1 != c3, "");

        // Blank name + blank ISIN -> single UNKNOWN bucket.
        Security u1 = store.getOrCreateSecurity("", "");
        Security u2 = store.getOrCreateSecurity(null, "  ");
        check("unknown securities collapse to one bucket", u1 == u2, "");

        // Total distinct instances: {Apple, Shared#1, Shared#2, Cash, Other, Unknown} = 6.
        check("total distinct securities", store.getSecurities().size() == 6,
                String.valueOf(store.getSecurities().size()) + " (expected 6)");

        if (failures == 0) {
            System.out.println("\nSECURITY IDENTITY TEST: PASS");
        } else {
            System.out.println("\nSECURITY IDENTITY TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
