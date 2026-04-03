package report;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class CurrencyConversionService {

    private static final Pattern YAHOO_CLOSE_ARRAY = Pattern.compile("\\\"close\\\"\\s*:\\s*\\[(.*?)\\]", Pattern.DOTALL);
    private static final Path FX_CACHE_FILE = Path.of(".cache", "fx-rates-to-nok.csv");
    private static final Duration FX_CACHE_TTL = Duration.ofHours(12);
    private static final Map<String, CacheEntry> MEMORY_CACHE = Collections.synchronizedMap(new LinkedHashMap<>());
    private static final Set<String> COMMON_TARGET_CURRENCIES = Set.of(
            "NOK", "USD", "EUR", "SEK", "DKK", "GBP", "CHF", "CAD", "AUD", "JPY"
    );

    private static final class CacheEntry {
        private final double rate;
        private final Instant fetchedAt;

        private CacheEntry(double rate, Instant fetchedAt) {
            this.rate = rate;
            this.fetchedAt = fetchedAt;
        }

        private boolean isFresh(Instant now) {
            if (fetchedAt == null) {
                return false;
            }
            return fetchedAt.plus(FX_CACHE_TTL).isAfter(now);
        }
    }

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

        Instant now = Instant.now();
        Map<String, CacheEntry> diskCache = readDiskCache();
        LinkedHashMap<String, Double> ratesToNok = new LinkedHashMap<>();
        ratesToNok.put("NOK", 1.0);
        LinkedHashSet<String> needsFetch = new LinkedHashSet<>();

        for (String currency : currencies) {
            if ("NOK".equals(currency)) {
                continue;
            }

            CacheEntry memoryEntry = MEMORY_CACHE.get(currency);
            if (memoryEntry != null && memoryEntry.rate > 0.0 && memoryEntry.isFresh(now)) {
                ratesToNok.put(currency, memoryEntry.rate);
                continue;
            }

            CacheEntry diskEntry = diskCache.get(currency);
            if (diskEntry != null && diskEntry.rate > 0.0 && diskEntry.isFresh(now)) {
                ratesToNok.put(currency, diskEntry.rate);
                MEMORY_CACHE.put(currency, diskEntry);
                continue;
            }

            needsFetch.add(currency);
        }

        if (!needsFetch.isEmpty()) {
            Map<String, Double> fetchedRates = fetchRatesInParallel(needsFetch);
            for (String currency : needsFetch) {
                Double fetchedRate = fetchedRates.get(currency);
                if (fetchedRate != null && fetchedRate > 0.0) {
                    ratesToNok.put(currency, fetchedRate);
                    CacheEntry newEntry = new CacheEntry(fetchedRate, now);
                    MEMORY_CACHE.put(currency, newEntry);
                    diskCache.put(currency, newEntry);
                    continue;
                }

                CacheEntry fallback = diskCache.get(currency);
                if (fallback != null && fallback.rate > 0.0) {
                    ratesToNok.put(currency, fallback.rate);
                    MEMORY_CACHE.put(currency, fallback);
                }
            }
        }

        writeDiskCache(diskCache);
        return ratesToNok;
    }

    private static Map<String, Double> fetchRatesInParallel(Set<String> currencies) {
        LinkedHashMap<String, Double> results = new LinkedHashMap<>();
        if (currencies == null || currencies.isEmpty()) {
            return results;
        }

        int workerCount = Math.max(1, Math.min(6, currencies.size()));
        ExecutorService executor = Executors.newFixedThreadPool(workerCount);
        try {
            LinkedHashMap<String, Future<Double>> futures = new LinkedHashMap<>();
            for (String currency : currencies) {
                Callable<Double> task = () -> fetchYahooCloseRate(currency + "NOK=X");
                futures.put(currency, executor.submit(task));
            }

            for (Map.Entry<String, Future<Double>> entry : futures.entrySet()) {
                try {
                    Double value = entry.getValue().get();
                    if (value != null && value > 0.0) {
                        results.put(entry.getKey(), value);
                    }
                } catch (InterruptedException ex) {
                    Thread.currentThread().interrupt();
                    break;
                } catch (ExecutionException ignored) {
                    // Keep missing value and fall back to cache if available.
                }
            }

            return results;
        } finally {
            executor.shutdownNow();
        }
    }

    private static Map<String, CacheEntry> readDiskCache() {
        LinkedHashMap<String, CacheEntry> cache = new LinkedHashMap<>();
        if (!Files.exists(FX_CACHE_FILE)) {
            return cache;
        }

        try (BufferedReader reader = Files.newBufferedReader(FX_CACHE_FILE, StandardCharsets.UTF_8)) {
            String line;
            while ((line = reader.readLine()) != null) {
                String trimmed = line.trim();
                if (trimmed.isEmpty() || trimmed.startsWith("#")) {
                    continue;
                }

                String[] parts = trimmed.split(",");
                if (parts.length != 3) {
                    continue;
                }

                String code = normalizeCurrencyCode(parts[0]);
                if (code == null || "NOK".equals(code)) {
                    continue;
                }

                try {
                    double rate = Double.parseDouble(parts[1]);
                    long epochMillis = Long.parseLong(parts[2]);
                    if (rate <= 0.0) {
                        continue;
                    }
                    cache.put(code, new CacheEntry(rate, Instant.ofEpochMilli(epochMillis)));
                } catch (NumberFormatException ignored) {
                    // Ignore malformed cache rows.
                }
            }
        } catch (IOException ignored) {
            return cache;
        }

        return cache;
    }

    private static void writeDiskCache(Map<String, CacheEntry> cache) {
        if (cache == null || cache.isEmpty()) {
            return;
        }

        try {
            Path parent = FX_CACHE_FILE.getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }

            try (BufferedWriter writer = Files.newBufferedWriter(
                    FX_CACHE_FILE,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.TRUNCATE_EXISTING,
                    StandardOpenOption.WRITE)) {
                writer.write("# currency,rateToNok,fetchedAtEpochMillis\n");
                for (Map.Entry<String, CacheEntry> entry : cache.entrySet()) {
                    String code = normalizeCurrencyCode(entry.getKey());
                    CacheEntry value = entry.getValue();
                    if (code == null || value == null || value.rate <= 0.0 || value.fetchedAt == null) {
                        continue;
                    }
                    writer.write(code + "," + String.format(Locale.US, "%.8f", value.rate) + "," + value.fetchedAt.toEpochMilli());
                    writer.newLine();
                }
            }
        } catch (IOException ignored) {
            // Cache write should never fail report generation.
        }
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
