package report;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class CurrencyConversionService {

    private static final Pattern YAHOO_CLOSE_ARRAY = Pattern.compile("\\\"close\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Set<String> COMMON_TARGET_CURRENCIES = Set.of(
            "NOK", "USD", "EUR", "SEK", "DKK", "GBP", "CHF", "CAD", "AUD", "JPY"
    );

    private CurrencyConversionService() {
    }

    public static Map<String, Double> loadRatesToNok(Set<String> requestedCurrencies) {
        LinkedHashSet<String> currencies = new LinkedHashSet<>();
        currencies.add("NOK");
        currencies.addAll(COMMON_TARGET_CURRENCIES);

        if (requestedCurrencies != null) {
            for (String code : requestedCurrencies) {
                String normalized = normalizeCurrencyCode(code);
                if (normalized != null) {
                    currencies.add(normalized);
                }
            }
        }

        LinkedHashMap<String, Double> ratesToNok = new LinkedHashMap<>();
        ratesToNok.put("NOK", 1.0);

        for (String currency : currencies) {
            if ("NOK".equals(currency)) {
                continue;
            }

            Double rate = fetchYahooCloseRate(currency + "NOK=X");
            if (rate != null && rate > 0.0) {
                ratesToNok.put(currency, rate);
            }
        }

        return ratesToNok;
    }

    private static String normalizeCurrencyCode(String code) {
        if (code == null || code.isBlank()) {
            return null;
        }

        String normalized = code.trim().toUpperCase(Locale.ROOT);
        if (!normalized.matches("[A-Z]{3}")) {
            return null;
        }
        return normalized;
    }

    private static Double fetchYahooCloseRate(String symbol) {
        HttpURLConnection connection = null;
        try {
                URL url = URI.create("https://query1.finance.yahoo.com/v8/finance/chart/"
                    + symbol
                    + "?range=10d&interval=1d&events=history").toURL();

            connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);
            connection.setRequestProperty("User-Agent", "Mozilla/5.0");

            int status = connection.getResponseCode();
            if (status != 200) {
                return null;
            }

            StringBuilder body = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    body.append(line);
                }
            }

            Matcher closeMatcher = YAHOO_CLOSE_ARRAY.matcher(body);
            if (!closeMatcher.find()) {
                return null;
            }

            String[] parts = closeMatcher.group(1).split(",");
            ArrayList<Double> values = new ArrayList<>();
            for (String part : parts) {
                String token = part.trim();
                if (token.isEmpty() || "null".equalsIgnoreCase(token)) {
                    continue;
                }

                try {
                    double value = Double.parseDouble(token);
                    if (value > 0.0) {
                        values.add(value);
                    }
                } catch (NumberFormatException ignored) {
                    // Ignore malformed entries and continue.
                }
            }

            if (values.isEmpty()) {
                return null;
            }

            return values.get(values.size() - 1);
        } catch (IOException ex) {
            return null;
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
    }

    public static String toJson(Map<String, Double> ratesToNok) {
        StringBuilder json = new StringBuilder();
        json.append("{");

        boolean first = true;
        for (Map.Entry<String, Double> entry : ratesToNok.entrySet()) {
            String code = normalizeCurrencyCode(entry.getKey());
            Double value = entry.getValue();
            if (code == null || value == null || value <= 0.0) {
                continue;
            }

            if (!first) {
                json.append(",");
            }
            first = false;
            json.append("\"").append(code).append("\":")
                    .append(String.format(Locale.US, "%.8f", value));
        }

        json.append("}");
        return json.toString();
    }
}
