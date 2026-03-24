package report;

import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.Locale;

public class HtmlFormatter {

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

    private static String formatNumber(double value, int decimals) {
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(Locale.US);
        symbols.setGroupingSeparator(' ');
        symbols.setDecimalSeparator('.');

        DecimalFormat format = new DecimalFormat();
        format.setDecimalFormatSymbols(symbols);
        format.setGroupingUsed(true);
        format.setMinimumFractionDigits(decimals);
        format.setMaximumFractionDigits(decimals);
        return format.format(value);
    }

    public static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("\"", "&quot;")
                   .replace("'", "&#39;");
    }

    public static String svgNumber(double value) {
        return String.format(Locale.US, "%.2f", value);
    }
}