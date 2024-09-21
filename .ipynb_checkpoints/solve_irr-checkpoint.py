import pandas as pd
import xlwings as xw
from scipy.optimize import brentq
from numpy import exp, log
from typing import Tuple, List

data_sheet_name, market_sheet_name = 'Data', 'MarketData'
data_table_range, market_table_range = 'A1:ZZ1', 'D1:ZZ1'
investment_funds = ['Colchón', 'Fondo A', 'APV', 'Acciones', 'Crypto']
benchmarks = ['TPM', 'CLFCLP', 'USDCLP OBS', 'S&P 500']

def npv(days: pd.Series, cashflows: pd.Series, rate: float) -> float:
    discount_factors = exp(-rate * days / 365)
    return (cashflows * discount_factors).sum()

def get_yield(rate, days):
    return exp(rate * days / 365) - 1

def simple_irr_solve(days: int, start_flow, end_flow) -> Tuple[float, float]:
    irr_solution = 365 / days * log(-end_flow / start_flow) # end_flow < 0
    effective_yield = -end_flow / start_flow - 1
    return irr_solution, effective_yield


def solve_irr(days: pd.Series, cashflows: pd.Series) -> Tuple[float, float]:
    if (cashflows != 0).sum() == 2:
        return simple_irr_solve(days.iloc[-1], cashflows.iloc[0], cashflows.iloc[-1])
    try:
        irr_solution = brentq(lambda r: npv(days, cashflows, r), -1, 1)
    except ValueError:
        irr_solution = brentq(lambda r: npv(days, cashflows, r), -1000, 1000)

    effective_yield = exp(irr_solution * days.iloc[-1] / 365) - 1.
    if abs(npv(days, cashflows, irr_solution)) > 1000: 
        raise ValueError('IRR calculation failed')
    return irr_solution, effective_yield


def get_effective_cashflow(cashflow: pd.Series, pnl: pd.Series) -> Tuple[pd.Series, pd.Series]:
    eff_cashflow = cashflow * -1.0
    eff_cashflow.iloc[[0, -1]] = -pnl.iloc[0], eff_cashflow.iloc[-1] + pnl.iloc[-1]
    days = eff_cashflow.index.to_series().diff().dt.days.cumsum().fillna(0)

    return days, eff_cashflow


def get_fund_performance(data: pd.DataFrame, funds: List[str]) -> pd.DataFrame:
    dates = data.index.unique()
    table = data.pivot_table(index='Fecha', columns='Linea', values=['Total', 'Transferencias'], aggfunc='sum')
    table_irr = pd.DataFrame(index=dates, columns=funds).fillna(0)
    table_yield = table_irr.copy()

    for t in dates[1:]:
        for fund in funds:
            cashflow = table.loc[:t, ('Transferencias', fund)]
            pnl = table.loc[:t, ('Total', fund)]
            days, cashflow = get_effective_cashflow(cashflow, pnl)
            irr, y = solve_irr(days, cashflow)
            table_irr.loc[t, fund] = irr
            table_yield.loc[t, fund] = y

    table_irr.columns = [column + ' IRR' for column in table_irr.columns]
    table_yield.columns = [column + ' Yield' for column in table_yield.columns]

    return pd.concat([table_irr, table_yield], axis=1)

def get_benchmark_performance(market_data: pd.DataFrame, benchmarks: List[str]):
    dates = market_data.index.unique()
    table_irr = pd.DataFrame(index=market_data.index.unique(), columns=benchmarks).fillna(0)
    table_yield = table_irr.copy()

    for t in dates[1:]:
        for benchmark in benchmarks:
            pnl = market_data.loc[:t, benchmark]
            cashflow = pd.Series(0, index=pnl.index)
            days, cashflow = get_effective_cashflow(cashflow, pnl)
            irr, y = solve_irr(days, cashflow)
            table_irr.loc[t, benchmark] = irr
            table_yield.loc[t, benchmark] = y

    table_irr.columns = [column + ' IRR' for column in table_irr.columns]
    table_yield.columns = [column + ' Yield' for column in table_yield.columns]

    return pd.concat([table_irr, table_yield], axis=1)

def main():
    xw.Book("Ledger.xlsm").set_mock_caller()
    wb = xw.Book.caller()
    sheet = wb.sheets['IRR']
    start_date, end_date = pd.to_datetime(sheet.range("B1").value), pd.to_datetime(sheet.range("B2").value)

    data = (
    wb.sheets[data_sheet_name]
    .range(data_table_range)
    .options(pd.DataFrame, expand='table')
    .value
    .loc[start_date:end_date]
    .query("`Linea` in @investment_funds")
    .fillna(0)
    )
    fund_performance = get_fund_performance(data, investment_funds)

    market_data = (
        wb.sheets[market_sheet_name]
        .range(market_table_range)
        .options(pd.DataFrame, expand='table')
        .value
        .loc[start_date:end_date, benchmarks]
        .fillna(method='ffill')
        .fillna(method='bfill')
    )

    market_data['S&P 500'] *= market_data['USDCLP OBS']

    days = market_data.index.to_series().diff().dt.days.cumsum().fillna(0)
    market_data['TPM'] = (1 + market_data['TPM'] / 360 / 100) ** days.diff()
    market_data['TPM'] = market_data['TPM'].cumprod().fillna(1)

    market_data = market_data.loc[fund_performance.index.unique()]
    benchmark_performance = get_benchmark_performance(market_data, benchmarks)

    sheet.range('D:ZZ').clear_contents()
    first_range = sheet.range('D1')
    first_range.value = fund_performance

    second_range = first_range.offset(0, fund_performance.shape[1] + 9)
    second_range.options(index=False).value = benchmark_performance


