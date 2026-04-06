package report;

public class AnnualPerformanceSummary {
    public final int year;
    public final String benchmarkTicker;
    public final boolean hasPortfolioData;
    public final double startValueNok;
    public final double endValueNok;
    public final double portfolioReturnNok;
    public final double portfolioReturnPct;
    public final double realizedGainNok;
    public final double dividendsNok;
    public final double realizedTotalNok;
    public final boolean hasBenchmarkData;
    public final double benchmarkReturnPct;

    public AnnualPerformanceSummary(
            int year,
            String benchmarkTicker,
            boolean hasPortfolioData,
            double startValueNok,
            double endValueNok,
            double portfolioReturnNok,
            double portfolioReturnPct,
            double realizedGainNok,
            double dividendsNok,
            double realizedTotalNok,
            boolean hasBenchmarkData,
            double benchmarkReturnPct) {
        this.year = year;
        this.benchmarkTicker = benchmarkTicker;
        this.hasPortfolioData = hasPortfolioData;
        this.startValueNok = startValueNok;
        this.endValueNok = endValueNok;
        this.portfolioReturnNok = portfolioReturnNok;
        this.portfolioReturnPct = portfolioReturnPct;
        this.realizedGainNok = realizedGainNok;
        this.dividendsNok = dividendsNok;
        this.realizedTotalNok = realizedTotalNok;
        this.hasBenchmarkData = hasBenchmarkData;
        this.benchmarkReturnPct = benchmarkReturnPct;
    }
}
