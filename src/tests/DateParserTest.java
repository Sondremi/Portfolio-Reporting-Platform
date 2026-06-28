package tests;

import util.DateParser;

import java.time.LocalDate;

/**
 * Unit test for DateParser.parseTradeDate: ISO and dd.MM.yyyy broker formats,
 * plus the LocalDate.MIN sentinel returned for blank/unparseable input.
 *
 * Offline + deterministic.
 *
 * Run: java -cp out tests.DateParserTest
 */
public class DateParserTest {

    private static int failures = 0;

    public static void main(String[] args) {
        assertDate("ISO date", "2025-05-13", LocalDate.of(2025, 5, 13));
        assertDate("dot date", "13.05.2025", LocalDate.of(2025, 5, 13));
        assertDate("ISO with surrounding whitespace", "  2024-01-02 ", LocalDate.of(2024, 1, 2));

        assertMin("null", null);
        assertMin("blank", "   ");
        assertMin("garbage", "not-a-date");
        assertMin("US slash format unsupported", "05/13/2025");

        if (failures == 0) {
            System.out.println("\nDATE PARSER TEST: PASS");
        } else {
            System.out.println("\nDATE PARSER TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void assertDate(String label, String input, LocalDate expected) {
        LocalDate actual = DateParser.parseTradeDate(input);
        check(label, expected.equals(actual), actual + " (expected " + expected + ")");
    }

    private static void assertMin(String label, String input) {
        LocalDate actual = DateParser.parseTradeDate(input);
        check(label, LocalDate.MIN.equals(actual), actual + " (expected MIN)");
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
