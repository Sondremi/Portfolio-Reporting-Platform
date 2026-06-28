package tests;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;

/**
 * Runs every test in this package and aggregates the results. Each test is a
 * standalone main that exits non-zero on failure, so they are launched as
 * separate JVM processes and their exit codes are collected.
 *
 * Run:
 *   javac -d out src/*.java src/*&#47;*.java
 *   java -cp out tests.RunAll
 */
public class RunAll {

    private static final String[] TESTS = {
        "tests.NumericParsingTest",
        "tests.DateParserTest",
        "tests.HtmlFormatterTest",
        "tests.SimpleJsonTest",
        "tests.SecurityIdentityTest",
        "tests.CurrencyConversionTest",
        "tests.CsvBrokerFormatTest",
        "tests.CalculationGoldenMasterTest",
        "tests.CancellationTransactionTest",
        "tests.AnnualGoldenMasterTest",
    };

    public static void main(String[] args) throws Exception {
        String javaBin = System.getProperty("java.home") + "/bin/java";
        String classpath = System.getProperty("java.class.path");

        int failed = 0;
        for (String test : TESTS) {
            boolean ok = run(javaBin, classpath, test);
            System.out.printf("%-6s %s%n", ok ? "PASS" : "FAIL", test);
            if (!ok) failed++;
        }

        System.out.println();
        System.out.printf("%d/%d test class(es) passed.%n", TESTS.length - failed, TESTS.length);
        if (failed > 0) {
            System.exit(1);
        }
    }

    private static boolean run(String javaBin, String classpath, String test) throws Exception {
        ProcessBuilder builder = new ProcessBuilder(javaBin, "-cp", classpath, test);
        builder.redirectErrorStream(true);
        Process process = builder.start();

        // Drain output so the child never blocks on a full pipe; surface only on failure.
        StringBuilder output = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                output.append(line).append(System.lineSeparator());
            }
        }

        int exit = process.waitFor();
        if (exit != 0) {
            System.out.println(output);
        }
        return exit == 0;
    }
}
