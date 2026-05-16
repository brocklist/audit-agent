"""
CompositeFactor — 2025Q1 A股综合因子 (过拟合示例)
仅使用 close/volume/amount/turnover 四类基础行情字段, 无外部 I/O。
"""

import numpy as np
import pandas as pd


def SCALE(x: np.ndarray) -> np.ndarray:
    """截面标准化: 去均值除以标准差, 截尾到 [-3, 3]"""
    z = (x - np.nanmean(x)) / (np.nanstd(x) + 1e-12)
    return np.clip(z, -3, 3)


def RANK(x: np.ndarray) -> np.ndarray:
    """截面排名归一化到 [0, 1]"""
    from scipy.stats import rankdata
    r = rankdata(x, nan_policy="omit")
    return r / (np.nansum(~np.isnan(x)) + 1)


# ============================================================
# Q1 2025 核心标的 (沪深 300 + 中证 500 代表性股票)
# ============================================================
TICKERS = [
    # 大金融
    "000001.SZ", "600036.SH", "601318.SH", "600030.SH", "000002.SZ",
    # 消费
    "600519.SH", "000858.SZ", "002304.SZ", "600887.SH", "000568.SZ",
    # 新能源 / 制造业
    "300750.SZ", "002594.SZ", "601012.SH", "600438.SH", "300274.SZ",
    "002129.SZ", "688981.SH",
    # TMT / 科技
    "000063.SZ", "002415.SZ", "300124.SZ", "688111.SH", "002230.SZ",
    "688036.SH", "300033.SZ",
    # 医药
    "300760.SZ", "600276.SH", "000661.SZ", "300015.SZ",
    # 周期 / 其他
    "600585.SH", "601899.SH", "600809.SH", "603259.SH",
    "601088.SH", "600150.SH",
]


def generate_q1_2025_data(tickers=TICKERS, days=58, seed=42):
    """
    生成 2025Q1 共计约 58 个交易日的模拟行情。
    数据特征参考 2025Q1 真实市场:
      - 春节前缩量震荡, 春节后 (2月初) 科技股领涨
      - 两会期间 (3月初) 政策主题活跃
      - 3月中下旬分化加剧, 小微盘承压
    """
    rng = np.random.default_rng(seed)
    n_stocks = len(tickers)
    dates = pd.bdate_range("2025-01-02", "2025-03-31")

    # ---------- 构建有结构的走势 ----------
    t = np.linspace(0, 1, len(dates))

    # 市场基准: 先跌后涨再震荡 (模拟 Q1 真实节奏)
    market_return = (
        0.02 * np.sin(t * np.pi * 1.3)
        - 0.01 * np.cos(t * np.pi * 2.5)
        + 0.001 * rng.normal(0, 1, len(dates))
    )
    market_level = np.exp(np.cumsum(market_return))

    # 行业风格分组
    groups = {
        "finance":     [0, 1, 2, 3, 4],
        "consumer":    [5, 6, 7, 8, 9],
        "new_energy":  [10, 11, 12, 13, 14, 15, 16],
        "tmt":         [17, 18, 19, 20, 21, 22, 23],
        "healthcare":  [24, 25, 26, 27],
        "cyclical":    [28, 29, 30, 31, 32, 33],
    }

    # 各组 alpha (相对市场的超额收益)
    alphas = {
        "finance":     -0.03,
        "consumer":    -0.01,
        "new_energy":   0.04,
        "tmt":          0.12,
        "healthcare":   0.00,
        "cyclical":     0.02,
    }

    # 各组波动率
    vol_scale = {
        "finance": 1.0, "consumer": 1.1, "new_energy": 1.5,
        "tmt": 1.6, "healthcare": 1.2, "cyclical": 1.3,
    }

    # 初始价格
    init_prices = {
        "finance": 25.0, "consumer": 120.0, "new_energy": 55.0,
        "tmt": 80.0, "healthcare": 45.0, "cyclical": 30.0,
    }
    init_volumes = {
        "finance": 8e7, "consumer": 2e7, "new_energy": 3e7,
        "tmt": 5e7, "healthcare": 1.5e7, "cyclical": 4e7,
    }

    closes = np.zeros((n_stocks, len(dates)))
    volumes = np.zeros((n_stocks, len(dates)))
    amounts = np.zeros((n_stocks, len(dates)))
    turnovers = np.zeros((n_stocks, len(dates)))

    for g_name, idxs in groups.items():
        for i in idxs:
            stock_noise = rng.normal(0, 1, len(dates)) * 0.0012
            stock_return = (
                market_return
                + alphas[g_name] / 58
                + stock_noise
            )
            stock_level = init_prices[g_name] * np.exp(np.cumsum(stock_return))
            closes[i] = stock_level

            base_vol = init_volumes[g_name]
            vol_arr = base_vol * np.exp(
                0.3 * np.sin(t * np.pi * 2.0)
                + 0.08 * rng.normal(0, 1, len(dates))
            )
            volumes[i] = vol_arr
            amounts[i] = closes[i] * vol_arr
            turnovers[i] = vol_arr / (base_vol / 0.015)  # 换手率 ~1.5%

    # 引入个别股票的异常波动 (用于过拟合信号)
    # 比如 tmt 组某几只股票在春节后有爆发式涨幅
    for special_idx in [19, 21, 22]:  # 300124.SZ, 688036.SH, 300033.SZ
        boost_start = 18  # 约 1月28日前后
        closes[special_idx, boost_start:] *= 1.0 + 0.006 * np.arange(1, len(dates) - boost_start + 1)

    # 某医药股在3月有大幅回调
    closes[27, -15:] *= 1.0 - 0.015 * np.arange(1, 16)

    return {
        "dates": dates,
        "tickers": tickers,
        "close": closes,
        "volume": volumes,
        "amount": amounts,
        "turnover": turnovers,
    }


# ============================================================
# 子因子
# ============================================================

def momentum_5d(close: np.ndarray) -> np.ndarray:
    """5日动量"""
    return close / shift(close, 5) - 1


def momentum_20d(close: np.ndarray) -> np.ndarray:
    """20日动量"""
    return close / shift(close, 20) - 1


def volume_ratio_5d(volume: np.ndarray) -> np.ndarray:
    """5日均量比"""
    ma5 = rolling_mean(volume, 5)
    ma20 = rolling_mean(volume, 20)
    return ma5 / (ma20 + 1e-12)


def turnover_ratio(volume: np.ndarray, turnover: np.ndarray) -> np.ndarray:
    """换手率变化"""
    return turnover / (rolling_mean(turnover, 10) + 1e-12) - 1


def amplitude(close: np.ndarray) -> np.ndarray:
    """近期振幅因子"""
    high = rolling_max(close, 10)
    low = rolling_min(close, 10)
    return (high - low) / (rolling_mean(close, 10) + 1e-12)


def volume_price_corr(close: np.ndarray, volume: np.ndarray) -> np.ndarray:
    """价量相关性 (10日窗口)"""
    return rolling_corr(close, volume, 10)


def intraday_volatility(amount: np.ndarray, turnover: np.ndarray) -> np.ndarray:
    """日内波动代理: 成交额/(换手率+eps) 的变异"""
    proxy = amount / (turnover + 1e-12)
    return rolling_std(proxy, 10) / (rolling_mean(proxy, 10) + 1e-12)


def gap_factor(close: np.ndarray) -> np.ndarray:
    """跳空因子: 开盘缺口代理 (当日涨幅 - 前日涨幅 的差异)"""
    daily_ret = close / shift(close, 1) - 1
    return daily_ret - shift(daily_ret, 1)


def reversal(close: np.ndarray) -> np.ndarray:
    """短期反转: -1 * 最近2日收益"""
    return -(close / shift(close, 2) - 1)


def liquidity_spread(volume: np.ndarray, amount: np.ndarray) -> np.ndarray:
    """流动性价差: 量增价滞的信号"""
    vol_change = volume / rolling_mean(volume, 10) - 1
    price_change = close_ret(amount, volume)  # 会被外部 close 覆盖, 这里重命名
    # 用 amount / volume 作为均价代理
    avg_price = amount / (volume + 1e-12)
    price_chg = avg_price / rolling_mean(avg_price, 10) - 1
    return vol_change - price_chg


def coppock_like(close: np.ndarray) -> np.ndarray:
    """类 Coppock 指标: 长周期动量变化率"""
    roc14 = close / shift(close, 14) - 1
    roc11 = close / shift(close, 11) - 1
    raw = roc14 + roc11
    return rolling_mean(raw, 10)


def roe_yoy_proxy(turnover: np.ndarray, amount: np.ndarray) -> np.ndarray:
    """
    ROE 同比的代理变量:
    用换手率稳定性与成交额趋势的交互构造。
    (真实 ROE 需要财报, 这里用交易行为代理)
    """
    turnover_stability = -rolling_std(turnover, 20) / (rolling_mean(turnover, 20) + 1e-12)
    amount_trend = amount / rolling_mean(amount, 20) - 1
    return turnover_stability * SCALE(amount_trend)


# ============================================================
# 向量化工具函数
# ============================================================

def shift(arr: np.ndarray, lag: int) -> np.ndarray:
    """按行平移 (每行=一只股票)"""
    if lag <= 0:
        return arr.copy()
    out = np.full_like(arr, np.nan)
    out[:, lag:] = arr[:, :-lag]
    return out


def rolling_mean(arr: np.ndarray, window: int) -> np.ndarray:
    out = np.full_like(arr, np.nan)
    for i in range(arr.shape[0]):
        s = pd.Series(arr[i]).rolling(window, min_periods=window).mean()
        out[i] = s.values
    return out


def rolling_std(arr: np.ndarray, window: int) -> np.ndarray:
    out = np.full_like(arr, np.nan)
    for i in range(arr.shape[0]):
        s = pd.Series(arr[i]).rolling(window, min_periods=window).std()
        out[i] = s.values
    return out


def rolling_max(arr: np.ndarray, window: int) -> np.ndarray:
    out = np.full_like(arr, np.nan)
    for i in range(arr.shape[0]):
        s = pd.Series(arr[i]).rolling(window, min_periods=window).max()
        out[i] = s.values
    return out


def rolling_min(arr: np.ndarray, window: int) -> np.ndarray:
    out = np.full_like(arr, np.nan)
    for i in range(arr.shape[0]):
        s = pd.Series(arr[i]).rolling(window, min_periods=window).min()
        out[i] = s.values
    return out


def rolling_corr(a: np.ndarray, b: np.ndarray, window: int) -> np.ndarray:
    out = np.full_like(a, np.nan)
    for i in range(a.shape[0]):
        s_a = pd.Series(a[i])
        s_b = pd.Series(b[i])
        out[i] = s_a.rolling(window, min_periods=window).corr(s_b).values
    return out


# ============================================================
# CompositeFactor
# ============================================================

class Factor:
    pass


class CompositeFactor(Factor):
    """
    综合因子: 融合动量、反转、波动、流动性、价量相关等 12 个子信号,
    使用带权重的方式合成, 并通过截面标准化得到最终因子值。

    权重经过 Q1 2025 数据优化 (过拟合), 追求在该段历史上最高 IC。
    """

    def calculate(self, factors: dict) -> np.ndarray:
        close = factors["close"]
        volume = factors["volume"]
        amount = factors["amount"]
        turnover = factors["turnover"]

        # ------ 构造12个子因子 ------
        f_mom5 = SCALE(momentum_5d(close))
        f_mom20 = SCALE(momentum_20d(close))
        f_vol_ratio = SCALE(volume_ratio_5d(volume))
        f_turnover_chg = SCALE(turnover_ratio(volume, turnover))
        f_amplitude = SCALE(-amplitude(close))  # 低振幅趋势好
        f_vp_corr = SCALE(volume_price_corr(close, volume))
        f_intra_vol = SCALE(-intraday_volatility(amount, turnover))
        f_gap = SCALE(gap_factor(close))
        f_reversal = SCALE(reversal(close))
        f_coppock = SCALE(coppock_like(close))
        f_roe_proxy = SCALE(roe_yoy_proxy(turnover, amount))

        avg_price = amount / (volume + 1e-12)
        f_liquidity = SCALE(-(volume / rolling_mean(volume, 10) - avg_price / rolling_mean(avg_price, 10)))

        # ------ 权重 (经 Q1 2025 数据优化) ------
        # 核心逻辑: Q1 科技成长领涨, 动量+价量配合是主信号
        w_mom5        = 0.10
        w_mom20       = 0.18
        w_vol_ratio   = 0.10
        w_turnover_chg= 0.06
        w_amplitude   = 0.05
        w_vp_corr     = 0.10
        w_intra_vol   = 0.05
        w_gap         = 0.07
        w_reversal    = 0.03
        w_coppock     = 0.10
        w_roe_proxy   = 0.08
        w_liquidity   = 0.08

        # ------ 加权合成 ------
        composite = (
            w_mom5        * f_mom5
            + w_mom20       * f_mom20
            + w_vol_ratio   * f_vol_ratio
            + w_turnover_chg* f_turnover_chg
            + w_amplitude   * f_amplitude
            + w_vp_corr     * f_vp_corr
            + w_intra_vol   * f_intra_vol
            + w_gap         * f_gap
            + w_reversal    * f_reversal
            + w_coppock     * f_coppock
            + w_roe_proxy   * f_roe_proxy
            + w_liquidity   * f_liquidity
        )

        composite_score = SCALE(composite)

        return composite_score


# ============================================================
# 绩效评估
# ============================================================

def calc_ic(factor_values: np.ndarray, forward_returns: np.ndarray) -> dict:
    """
    计算 Rank IC 和 Pearson IC
    factor_values: (n_stocks, n_dates)
    forward_returns: (n_stocks, n_dates)  未来1日收益率
    """
    from scipy.stats import pearsonr, spearmanr

    rank_ics = []
    pearson_ics = []
    for t in range(factor_values.shape[1]):
        fv = factor_values[:, t]
        fr = forward_returns[:, t]
        mask = ~np.isnan(fv) & ~np.isnan(fr)
        if mask.sum() < 10:
            continue
        ric, _ = spearmanr(fv[mask], fr[mask])
        pic, _ = pearsonr(fv[mask], fr[mask])
        rank_ics.append(ric)
        pearson_ics.append(pic)

    rank_ics = np.array(rank_ics)
    pearson_ics = np.array(pearson_ics)

    return {
        "Rank_IC_mean": np.nanmean(rank_ics),
        "Rank_IC_std": np.nanstd(rank_ics),
        "Rank_IC_IR": np.nanmean(rank_ics) / (np.nanstd(rank_ics) + 1e-12),
        "Rank_IC_pos_ratio": np.nansum(rank_ics > 0) / max(1, np.sum(~np.isnan(rank_ics))),
        "Pearson_IC_mean": np.nanmean(pearson_ics),
        "Pearson_IC_IR": np.nanmean(pearson_ics) / (np.nanstd(pearson_ics) + 1e-12),
    }


def calc_factor_return(factor_values: np.ndarray, forward_returns: np.ndarray, n_bins: int = 10) -> dict:
    """
    分层回测: 多空收益、Top-Bottom 分组收益。
    """
    long_returns = []
    short_returns = []
    for t in range(factor_values.shape[1]):
        fv = factor_values[:, t]
        fr = forward_returns[:, t]
        mask = ~np.isnan(fv) & ~np.isnan(fr)
        if mask.sum() < n_bins * 3:
            continue

        fv_m = fv[mask]
        fr_m = fr[mask]
        order = np.argsort(fv_m)
        bin_size = len(order) // n_bins

        top_idx = order[-bin_size:]
        bot_idx = order[:bin_size]

        long_returns.append(np.nanmean(fr_m[top_idx]))
        short_returns.append(np.nanmean(fr_m[bot_idx]))

    long_arr = np.array(long_returns)
    short_arr = np.array(short_returns)
    long_short = long_arr - short_arr

    return {
        "Long_cum": np.sum(long_arr),
        "Short_cum": np.sum(short_arr),
        "Long_Short_cum": np.sum(long_short),
        "Long_ann_return": np.mean(long_arr) * 252,
        "Short_ann_return": np.mean(short_arr) * 252,
        "LS_ann_return": np.mean(long_short) * 252,
        "LS_sharpe": np.mean(long_short) / (np.std(long_short) + 1e-12) * np.sqrt(252),
        "LS_win_rate": np.mean(long_short > 0),
    }


def calc_cumulative_pnl(factor_values: np.ndarray, forward_returns: np.ndarray, n_bins: int = 10):
    """计算分组累计收益曲线 (等权)"""
    bin_pnls = np.zeros((n_bins, factor_values.shape[1]))
    bin_pnls[:] = np.nan
    for t in range(factor_values.shape[1]):
        fv = factor_values[:, t]
        fr = forward_returns[:, t]
        mask = ~np.isnan(fv) & ~np.isnan(fr)
        if mask.sum() < n_bins * 3:
            continue
        fv_m = fv[mask]
        fr_m = fr[mask]
        order = np.argsort(fv_m)
        bin_size = len(order) // n_bins
        prev = 0.0
        for b in range(n_bins):
            start = max(prev, b * bin_size)
            end = min(len(order), (b + 1) * bin_size) if b < n_bins - 1 else len(order)
            bin_idx = order[int(start):int(end)]
            bin_pnls[b, t] = np.nanmean(fr_m[bin_idx])
            prev = end
    cum = np.nancumsum(bin_pnls, axis=1)
    return cum


# ============================================================
# 主程序
# ============================================================

if __name__ == "__main__":
    print("=" * 60)
    print("CompositeFactor — 2025Q1 A股过拟合因子演示")
    print("=" * 60)

    # 1. 生成数据
    print("\n[1/5] 生成 Q1 2025 模拟行情数据...")
    data = generate_q1_2025_data()
    print(f"  标的数量: {len(data['tickers'])}")
    print(f"  交易日数: {len(data['dates'])}")
    print(f"  日期范围: {data['dates'][0].date()} ~ {data['dates'][-1].date()}")

    # 2. 计算因子值
    print("\n[2/5] 计算 CompositeFactor...")
    cf = CompositeFactor()
    factor_values = cf.calculate(data)
    print(f"  factor shape: {factor_values.shape}")

    # 3. 计算前向收益 (T+1)
    print("\n[3/5] 计算前向收益 (T+1)...")
    close = data["close"]
    forward_returns = close[:, 1:] / close[:, :-1] - 1
    factor_aligned = factor_values[:, :-1]  # 对齐, last factor value 无对应 forward return

    # 4. IC 分析
    print("\n[4/5] IC 分析...")
    ic_result = calc_ic(factor_aligned, forward_returns)
    for k, v in ic_result.items():
        print(f"  {k}: {v:.4f}")

    # 5. 因子收益
    print("\n[5/5] 分层回测...")
    ret_result = calc_factor_return(factor_aligned, forward_returns, n_bins=10)
    for k, v in ret_result.items():
        if isinstance(v, float):
            print(f"  {k}: {v:.4f}")
        else:
            print(f"  {k}: {v}")

    # 累计收益曲线
    cum_pnl = calc_cumulative_pnl(factor_aligned, forward_returns, n_bins=10)
    top_cum = cum_pnl[9]   # 第10组 (因子值最大)
    bot_cum = cum_pnl[0]   # 第1组 (因子值最小)
    ls_cum = top_cum - bot_cum

    print(f"\n  Top组累计收益:     {top_cum[-1]:.4f}")
    print(f"  Bottom组累计收益:  {bot_cum[-1]:.4f}")
    print(f"  多空累计收益:      {ls_cum[-1]:.4f}")
    print(f"  多空累计收益(年化): {np.nanmean(top_cum[-1:] - bot_cum[-1:]) * 252 / len(data['dates']):.4f}")

    print("\n" + "=" * 60)
    print("注意: 以上结果为过拟合因子在样本内的表现, 不代表实盘效果。")
    print("=" * 60)
