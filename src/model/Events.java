package model;

import java.time.LocalDate;

public final class Events {

    public record UnitEvent(LocalDate tradeDate, String securityKey, double unitsDelta) {}
    public record CashEvent(LocalDate tradeDate, double cashDelta) {}
    public record PortfolioCashSnapshot(LocalDate tradeDate, long sortId, String portfolioId, double balance) {}

    private Events() {}
}