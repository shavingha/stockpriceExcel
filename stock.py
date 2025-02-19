import xlwings as xw
import akshare as ak
from datetime import datetime


def get_hk_index_daily(symbol, query_date):
    try:
        query_date = query_date.strftime("%Y-%m-%d")
        index_daily = ak.stock_hk_index_daily_em(symbol=symbol)
        if not query_date in index_daily["date"].values:
            return None
        return index_daily[index_daily["date"] == query_date].iloc[0]["latest"]
    except Exception as e:
        return "error: " + str(e)


def get_cn_index_daily(symbol, query_date):
    try:
        start_date = query_date.strftime("%Y-%m-%d")
        end_date = start_date
        index_daily = ak.index_zh_a_hist(
            symbol=symbol, period="daily", start_date=start_date, end_date=end_date
        )
        if len(index_daily) == 0:
            return None
        return index_daily.iloc[0]["收盘"]
    except Exception as e:
        return "error: " + str(e)


@xw.func
def get_index_daily(symbol, query_date):
    symbol, region = symbol.split(".")
    if region == "HK":
        return get_hk_index_daily(symbol, query_date)
    elif region in ["SH", "SZ", "CSI"]:
        return get_cn_index_daily(symbol, query_date)
    else:
        return "ERROR SYMBOL"


# 获得单位净值
@xw.func
def get_fund_nav_daily(symbol, query_date):
    start_date = query_date.strftime("%Y%m%d")
    end_date = start_date
    symbol, _ = symbol.split(".")
    print(start_date, end_date)
    try:
        fund_info = ak.fund_etf_fund_info_em(
            fund=symbol, start_date=start_date, end_date=end_date
        )
        print(fund_info)
        if len(fund_info) == 0:
            return None
        return fund_info.iloc[0]["单位净值"]
    except Exception as e:
        return "error: " + str(e)


# adjust: qfq, hfq
@xw.func
def get_stock_daily(symbol, query_date, sdjust=""):
    start_date = query_date.strftime("%Y%m%d")
    end_date = start_date
    symbol, _ = symbol.split(".")
    df = ak.stock_zh_a_hist(
        symbol=symbol,
        period="daily",
        start_date=start_date,
        end_date=end_date,
        adjust=sdjust,
    )
    if len(df) == 0:
        return None
    return df.iloc[0]["收盘"]


def main():
    symbol = "000001.SZ"
    query_date = datetime.strptime("2025-02-14", "%Y-%m-%d")
    print(get_stock_daily(symbol, query_date))


if __name__ == "__main__":
    print("start main")
    xw.Book("stock.xlsm").set_mock_caller()
