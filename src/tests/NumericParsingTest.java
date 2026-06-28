package tests;

import csv.CsvLoader;

/**
 * Unit test for CsvLoader's tolerant numeric parsing — the source of the
 * "silent 0.0 on a critical cash field" risk. Pins which inputs are valid
 * numbers, which legitimately mean null/0, and which European formats
 * (thousands separators, comma decimals, nbsp, unicode minus) are accepted.
 *
 * Offline + deterministic.
 *
 * Run: java -cp out tests.NumericParsingTest
 */
public class NumericParsingTest {

    private static final double EPS = 1e-9;
    private static int failures = 0;

    public static void main(String[] args) {
        // Plain numbers.
        assertValue("integer", "1000", 1000.0);
        assertValue("decimal point", "1234.56", 1234.56);
        assertValue("negative", "-42.5", -42.5);
        assertValue("zero", "0", 0.0);

        // European formats.
        assertValue("comma decimal", "1234,56", 1234.56);
        assertValue("dot thousands + comma decimal", "1.234,56", 1234.56);
        assertValue("dot thousands + comma decimal (millions)", "1.234.567,89", 1234567.89);
        assertValue("space thousands", "1 234,56", 1234.56);
        assertValue("nbsp thousands", "1 234,56", 1234.56);
        assertValue("unicode minus", "−500,25", -500.25);
        assertValue("leading/trailing whitespace", "  -1 000,00  ", -1000.0);

        // Inputs that legitimately map to null (and therefore 0.0 via OrZero).
        assertNull("null", null);
        assertNull("empty", "");
        assertNull("whitespace only", "   ");

        // Garbage must NOT become 0.0 silently through parseDoubleOrNull —
        // it returns null so callers can distinguish "missing" from "zero".
        assertNull("alphabetic garbage", "abc");
        assertNull("currency suffix not stripped", "100 NOK");
        assertNull("stray dash", "-");

        // parseDoubleOrZero collapses both null and garbage to 0.0 — this is the
        // documented silent-zero behavior; locking it makes the risk explicit.
        assertZero("OrZero null", null);
        assertZero("OrZero empty", "");
        assertZero("OrZero garbage", "abc");
        assertCloseZeroVariant("OrZero valid passes through", "1.234,56", 1234.56);

        if (failures == 0) {
            System.out.println("\nNUMERIC PARSING TEST: PASS");
        } else {
            System.out.println("\nNUMERIC PARSING TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void assertValue(String label, String input, double expected) {
        Double actual = CsvLoader.parseDoubleOrNull(input);
        boolean ok = actual != null && Math.abs(actual - expected) <= EPS;
        check(label, ok, actual == null ? "null (expected " + expected + ")"
                : actual + " (expected " + expected + ")");
    }

    private static void assertNull(String label, String input) {
        Double actual = CsvLoader.parseDoubleOrNull(input);
        check(label, actual == null, String.valueOf(actual) + " (expected null)");
    }

    private static void assertZero(String label, String input) {
        double actual = CsvLoader.parseDoubleOrZero(input);
        check(label, Math.abs(actual) <= EPS, actual + " (expected 0.0)");
    }

    private static void assertCloseZeroVariant(String label, String input, double expected) {
        double actual = CsvLoader.parseDoubleOrZero(input);
        check(label, Math.abs(actual - expected) <= EPS, actual + " (expected " + expected + ")");
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
