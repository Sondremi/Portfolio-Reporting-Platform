package tests;

import util.SimpleJson;

import java.util.List;
import java.util.Map;

/**
 * Unit test for SimpleJson.parse: objects, arrays, strings (with escapes),
 * booleans, null, numbers (always Double), nesting, and lenient fallbacks.
 *
 * Offline + deterministic.
 *
 * Run: java -cp out tests.SimpleJsonTest
 */
public class SimpleJsonTest {

    private static int failures = 0;

    @SuppressWarnings("unchecked")
    public static void main(String[] args) {
        check("null input -> null", SimpleJson.parse(null) == null, "");

        // Scalars.
        check("number is Double", SimpleJson.parse("42.5").equals(42.5), String.valueOf(SimpleJson.parse("42.5")));
        check("negative exponent number", SimpleJson.parse("-1.5e2").equals(-150.0), String.valueOf(SimpleJson.parse("-1.5e2")));
        check("true", Boolean.TRUE.equals(SimpleJson.parse("true")), "");
        check("false", Boolean.FALSE.equals(SimpleJson.parse("false")), "");
        check("null literal", SimpleJson.parse("null") == null, "");
        check("string", "hello".equals(SimpleJson.parse("\"hello\"")), "");
        check("string with escapes", "a\nb\t\"c\"".equals(SimpleJson.parse("\"a\\nb\\t\\\"c\\\"\"")), "");

        // Object.
        Object obj = SimpleJson.parse("{\"a\":1,\"b\":\"x\",\"c\":true}");
        check("object is Map", obj instanceof Map, obj == null ? "null" : obj.getClass().getSimpleName());
        Map<String, Object> map = (Map<String, Object>) obj;
        check("object number value", map.get("a").equals(1.0), String.valueOf(map.get("a")));
        check("object string value", "x".equals(map.get("b")), String.valueOf(map.get("b")));
        check("object boolean value", Boolean.TRUE.equals(map.get("c")), String.valueOf(map.get("c")));

        // Array.
        Object arr = SimpleJson.parse("[1,2,3]");
        check("array is List", arr instanceof List, arr == null ? "null" : arr.getClass().getSimpleName());
        List<Object> list = (List<Object>) arr;
        check("array size", list.size() == 3, String.valueOf(list.size()));
        check("array first element", list.get(0).equals(1.0), String.valueOf(list.get(0)));

        // Nesting + whitespace tolerance.
        Object nested = SimpleJson.parse("{ \"rates\": { \"USD\": 10.5 }, \"flags\": [true, null] }");
        Map<String, Object> nestedMap = (Map<String, Object>) nested;
        Map<String, Object> rates = (Map<String, Object>) nestedMap.get("rates");
        check("nested object value", rates.get("USD").equals(10.5), String.valueOf(rates.get("USD")));
        List<Object> flags = (List<Object>) nestedMap.get("flags");
        check("nested array has null element", flags.size() == 2 && flags.get(1) == null, String.valueOf(flags));

        // Lenient fallback: malformed number token yields 0.0 rather than throwing.
        check("malformed number -> 0.0", SimpleJson.parse("--").equals(0.0), String.valueOf(SimpleJson.parse("--")));

        if (failures == 0) {
            System.out.println("\nSIMPLE JSON TEST: PASS");
        } else {
            System.out.println("\nSIMPLE JSON TEST: FAIL (" + failures + ")");
            System.exit(1);
        }
    }

    private static void check(String label, boolean ok, String detail) {
        if (ok) System.out.println("  ok   " + label);
        else { failures++; System.out.println("  FAIL " + label + " -> " + detail); }
    }
}
