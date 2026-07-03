package report;

import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

public class HtmlFormatter {

    // DecimalFormat is not thread-safe; keep one instance per decimals-count per thread.
    private static final ThreadLocal<Map<Integer, DecimalFormat>> NUMBER_FORMATS =
            ThreadLocal.withInitial(HashMap::new);

    public static String formatMoney(double value, String currency, int decimals) {
        if (currency == null || currency.isBlank()) {
            currency = "NOK";
        }

        if ("MIXED".equalsIgnoreCase(currency)) {
            return formatNumber(value, decimals) + " mixed";
        }

        return formatNumber(value, decimals) + " " + currency;
    }

    public static String formatPercent(double value, int decimals) {
        return formatNumber(value, decimals) + "%";
    }

    public static String formatPercent(double value) {
        return formatPercent(value, 2);
    }

    public static String formatUnits(double value) {
        String text = formatNumber(value, 4);
        if (text.contains(".")) {
            text = text.replaceAll("0+$", "").replaceAll("\\.$", "");
        }
        return text;
    }

    // Uses space thousands-grouping; ChartBuilder.formatNumber uses comma grouping.
    // Both are baked into current output — intentionally divergent, do not merge.
    private static String formatNumber(double value, int decimals) {
        DecimalFormat format = NUMBER_FORMATS.get().computeIfAbsent(decimals, HtmlFormatter::buildNumberFormat);
        return format.format(value);
    }

    private static DecimalFormat buildNumberFormat(int decimals) {
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(Locale.US);
        symbols.setGroupingSeparator(' ');
        symbols.setDecimalSeparator('.');

        DecimalFormat format = new DecimalFormat();
        format.setDecimalFormatSymbols(symbols);
        format.setGroupingUsed(true);
        format.setMinimumFractionDigits(decimals);
        format.setMaximumFractionDigits(decimals);
        return format;
    }

    public static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("\"", "&quot;")
                   .replace("'", "&#39;");
    }

    // ChartBuilder keeps a formatting-identical private copy; keep the two in sync.
    public static String svgNumber(double value) {
        return String.format(Locale.US, "%.2f", value);
    }
}