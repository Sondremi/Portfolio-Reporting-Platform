package util;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class DateParser {

    // Immutable/thread-safe; hoisted so repeated parses reuse one instance.
    private static final DateTimeFormatter[] FORMATS = {
        DateTimeFormatter.ISO_LOCAL_DATE,
        DateTimeFormatter.ofPattern("dd.MM.yyyy")
    };

    // The same date strings recur across every row and comparison; memoize the parse.
    private static final Map<String, LocalDate> CACHE = new ConcurrentHashMap<>();

    public static LocalDate parseTradeDate(String text) {
        if (text == null || text.isBlank()) return LocalDate.MIN;

        String value = text.trim();
        return CACHE.computeIfAbsent(value, DateParser::parseUncached);
    }

    private static LocalDate parseUncached(String value) {
        for (DateTimeFormatter fmt : FORMATS) {
            try {
                return LocalDate.parse(value, fmt);
            } catch (DateTimeParseException ignored) {}
        }
        return LocalDate.MIN;
    }
}
