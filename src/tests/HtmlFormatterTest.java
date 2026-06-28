package tests;

import report.HtmlFormatter;

/**
 * Unit test for HtmlFormatter. Locks two hard report rules into code:
 *   - numeric values are never truncated/ellipsized,
 *   - escapeHtml neutralizes markup (XSS safety).
 * Also covers currency fallback, MIXED handling, percent and unit trimming.
 *
 * Offline + deterministic.
 *
 * Run: java -cp out tests.HtmlFormatterTest
 */
public class HtmlFormatterTest {

    private static int failures = 0;

    public static void main(String[] args) {
        // Money: US grouping with space separators, dot decimal, currency suffix.
        assertEquals("formatMoney basic", "1 234 567.89 NOK",
                HtmlFormatter.formatMoney(1234567.89, "NOK", 2));
        assertEquals("formatMoney negative", "-1 000.50 USD",
                HtmlFormatter.formatMoney(-1000.5, "USD", 2));
        assertEquals("formatMoney blank currency falls back to NOK", "100.00 NOK",
                HtmlFormatter.formatMoney(100.0, "", 2));
        assertEquals("formatMoney null currency falls back to NOK", "100.00 NOK",
                HtmlFormatter.formatMoney(100.0, null, 2));
        assertEquals("formatMoney MIXED", "100.00 mixed",
                HtmlFormatter.formatMoney(100.0, "MIXED", 2));

        // Never-truncate rule: a huge value must render every digit, no ellipsis.
        String big = HtmlFormatter.formatMoney(1234567890123.45, "NOK", 2);
        check("no unicode ellipsis", !big.contains("…"), big);
        check("no ascii ellipsis", !big.contains("..."), big);
        assertEquals("full big value", "1 234 567 890 123.45 NOK", big);

        // Percent.
        assertEquals("formatPercent 2dp", "12.50%", HtmlFormatter.formatPercent(12.5, 2));
        assertEquals("formatPercent default dp", "12.50%", HtmlFormatter.formatPercent(12.5));
        assertEquals("formatPercent negative", "-3.25%", HtmlFormatter.formatPercent(-3.25, 2));

        // Units: trailing zeros and trailing dot trimmed; integers stay clean.
        assertEquals("formatUnits integer", "100", HtmlFormatter.formatUnits(100.0));
        assertEquals("formatUnits one decimal", "1.5", HtmlFormatter.formatUnits(1.5));
        assertEquals("formatUnits four decimals kept", "0.0001", HtmlFormatter.formatUnits(0.0001));
        assertEquals("formatUnits grouped", "1 234", HtmlFormatter.formatUnits(1234.0));

        // escapeHtml: all five entities, in the right order (& not double-escaped).
        assertEquals("escapeHtml null -> empty", "", HtmlFormatter.escapeHtml(null));
        assertEquals("escapeHtml markup",
                "&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;",
                HtmlFormatter.escapeHtml("<script>alert(\"x\")</script>"));
        assertEquals("escapeHtml ampersand + apostrophe", "&amp;&#39;",
                HtmlFormatter.escapeHtml("&'"));

        // svgNumber.
        assertEquals("svgNumber", "3.14", HtmlFormatter.svgNumber(3.14159));

        if (failures == 0) {
            System.out.println("\nHTML FORMATTER TEST: PASS");
        } else {
            System.out.println("\nHTML FORMATTER TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void assertEquals(String label, String expected, String actual) {
        check(label, expected.equals(actual), "\"" + actual + "\" (expected \"" + expected + "\")");
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
