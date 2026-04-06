package model;

import java.time.LocalDate;

public final class Events {

    public record UnitEvent(LocalDate tradeDate, String securityKey, double unitsDelta) {}
    public record CashEvent(LocalDate tradeDate, double cashDelta, boolean externalFlow) {
        public CashEvent(LocalDate tradeDate, double cashDelta) {
            this(tradeDate, cashDelta, false);
        }
    }
    public record PortfolioCashSnapshot(LocalDate tradeDate, long sortId, String portfolioId, double balance) {}

    private Events() {}
}