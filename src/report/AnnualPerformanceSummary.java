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
    public final boolean hasAnalytics;
    public final double annualizedVolatilityPct;
    public final double sharpeRatio;
    public final boolean hasBeta;
    public final double beta;
    public final boolean hasMonteCarlo;
    public final int monteCarloHorizonMonths;
    public final int monteCarloIterations;
    public final double monteCarloStartValueNok;
    public final double monteCarloMedianEndValueNok;
    public final double monteCarloP10EndValueNok;
    public final double monteCarloP90EndValueNok;
    public final double monteCarloExpectedEndValueNok;

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
            double benchmarkReturnPct,
            boolean hasAnalytics,
            double annualizedVolatilityPct,
            double sharpeRatio,
            boolean hasBeta,
            double beta,
            boolean hasMonteCarlo,
            int monteCarloHorizonMonths,
            int monteCarloIterations,
            double monteCarloStartValueNok,
            double monteCarloMedianEndValueNok,
            double monteCarloP10EndValueNok,
            double monteCarloP90EndValueNok,
            double monteCarloExpectedEndValueNok) {
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
        this.hasAnalytics = hasAnalytics;
        this.annualizedVolatilityPct = annualizedVolatilityPct;
        this.sharpeRatio = sharpeRatio;
        this.hasBeta = hasBeta;
        this.beta = beta;
        this.hasMonteCarlo = hasMonteCarlo;
        this.monteCarloHorizonMonths = monteCarloHorizonMonths;
        this.monteCarloIterations = monteCarloIterations;
        this.monteCarloStartValueNok = monteCarloStartValueNok;
        this.monteCarloMedianEndValueNok = monteCarloMedianEndValueNok;
        this.monteCarloP10EndValueNok = monteCarloP10EndValueNok;
        this.monteCarloP90EndValueNok = monteCarloP90EndValueNok;
        this.monteCarloExpectedEndValueNok = monteCarloExpectedEndValueNok;
    }
}
