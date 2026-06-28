package tests;

import csv.TransactionStore;
import report.CurrencyConversionService;
import report.HeaderSummary;
import report.OverviewRow;
import report.PortfolioCalculator;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Tests mixed-currency header aggregation (risk #1) and FX JSON serialization.
 *
 * buildHeaderSummary converts every row's market value / return / cost basis to
 * NOK using an injected rate map (no network), so the totals and best/worst
 * picks are fully deterministic. Also pins the fallback behavior for a currency
 * missing from the rate map, and CurrencyConversionService.toJson formatting.
 *
 * Offline + deterministic.
 *
 * Run: java -cp out tests.CurrencyConversionTest
 */
public class CurrencyConversionTest {

    private static final double EPS = 1e-6;
    private static int failures = 0;

    public static void main(String[] args) {
        Map<String, Double> rates = new LinkedHashMap<>();
        rates.put("NOK", 1.0);
        rates.put("USD", 10.0);
        rates.put("EUR", 11.0);

        List<OverviewRow> rows = new ArrayList<>();
        rows.add(row("NOK Co", "NOK", 1000.0, 100.0, 900.0, 11.11));
        rows.add(row("USD Co", "USD", 200.0, 50.0, 150.0, 33.33));   // -> 2000 mv, 500 ret, 1500 cost
        rows.add(row("EUR Co", "EUR", 100.0, -20.0, 120.0, -16.67)); // -> 1100 mv, -220 ret, 1320 cost

        TransactionStore store = new TransactionStore(); // no events -> cash 0, no sparkline
        HeaderSummary h = PortfolioCalculator.buildHeaderSummary(store, rows, rates);

        assertClose("total market value (NOK)", 4100.0, h.totalMarketValue);
        assertClose("total return (NOK)", 380.0, h.totalReturn);
        assertClose("total return pct", 380.0 / 3720.0 * 100.0, h.totalReturnPct);
        assertClose("cash holdings (empty store)", 0.0, h.cashHoldings);
        check("holdings count", h.holdingsCount == 3, String.valueOf(h.holdingsCount));
        check("total currency code", "NOK".equals(h.totalCurrencyCode), h.totalCurrencyCode);

        // Best = USD row (500 NOK), worst = EUR row (-220 NOK), measured in NOK.
        assertClose("best return (NOK)", 500.0, h.bestReturn);
        assertClose("worst return (NOK)", -220.0, h.worstReturn);

        // Fallback: a currency absent from the rate map is treated as NOK (rate 1.0),
        // never dropped — so its raw value still contributes.
        List<OverviewRow> fallbackRows = new ArrayList<>();
        fallbackRows.add(row("GBP Co", "GBP", 50.0, 5.0, 45.0, 11.11));
        HeaderSummary hf = PortfolioCalculator.buildHeaderSummary(new TransactionStore(), fallbackRows, rates);
        assertClose("missing-currency falls back to NOK rate", 50.0, hf.totalMarketValue);

        // toJson: skips non-positive rates and non-3-letter codes; uppercases; 8 dp.
        Map<String, Double> fx = new LinkedHashMap<>();
        fx.put("NOK", 1.0);
        fx.put("usd", 10.5);
        fx.put("EUR", 0.0);    // dropped: non-positive
        fx.put("EURO", 5.0);   // dropped: not a 3-letter code
        String json = CurrencyConversionService.toJson(fx);
        check("toJson formatting",
                "{\"NOK\":1.00000000,\"USD\":10.50000000}".equals(json), json);

        if (failures == 0) {
            System.out.println("\nCURRENCY CONVERSION TEST: PASS");
        } else {
            System.out.println("\nCURRENCY CONVERSION TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static OverviewRow row(String name, String currency, double marketValue,
                                   double totalReturn, double historicalCostBasis, double totalReturnPct) {
        return new OverviewRow(
                name, "T", name, 0.0, 0.0, false,
                "STOCK", "Other", null, "Global", null,
                currency, "0.00",
                1.0, 0.0, 0.0, 0.0, 0.0, historicalCostBasis,
                marketValue, 0.0, 0.0, 0.0, 0.0,
                totalReturn, totalReturnPct, true);
    }

    private static void assertClose(String label, double expected, double actual) {
        check(label, Math.abs(expected - actual) <= EPS,
                String.format(java.util.Locale.US, "%.6f (expected %.6f)", actual, expected));
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
