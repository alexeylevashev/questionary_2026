"""Stage 4 – EVA analysis: hourly distribution, trip duration, BW model fitting."""

from __future__ import annotations

import re
from typing import Optional

import numpy as np
import pandas as pd

from .config import Config


# ---------------------------------------------------------------------------
# Time helpers
# ---------------------------------------------------------------------------

def extract_hour(val) -> Optional[int]:
    """Extract the hour (0-23) from various Excel time representations."""
    if pd.isna(val):
        return None

    if hasattr(val, "hour"):
        try:
            return int(val.hour)
        except Exception:
            pass

    if isinstance(val, (int, float)):
        x = float(val)
        if 0 <= x < 1:
            return int(x * 24)
        dt = pd.to_datetime(x, unit="D", origin="1899-12-30", errors="coerce")
        return int(dt.hour) if pd.notna(dt) else None

    s = str(val).strip()
    if not s:
        return None

    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return int(dt.hour)

    m = re.match(r"^\s*(\d{1,2})\s*[:.]\s*(\d{1,2})", s)
    if m:
        h = int(m.group(1))
        return h if 0 <= h <= 23 else None

    return None


def time_to_minutes(val) -> Optional[int]:
    """Convert an Excel time value to minutes since midnight."""
    if pd.isna(val):
        return None

    if hasattr(val, "hour") and hasattr(val, "minute"):
        try:
            return int(val.hour) * 60 + int(val.minute)
        except Exception:
            pass

    if isinstance(val, (int, float)):
        x = float(val)
        if 0 <= x < 1:
            mins = int(round(x * 24 * 60))
            return max(0, min(1439, mins))
        dt = pd.to_datetime(x, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(dt):
            return int(dt.hour) * 60 + int(dt.minute)
        return None

    s = str(val).strip()
    if not s:
        return None

    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return int(dt.hour) * 60 + int(dt.minute)

    m = re.match(r"^\s*(\d{1,2})\s*[:.]\s*(\d{1,2})", s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return h * 60 + mi

    return None


# ---------------------------------------------------------------------------
# Duration bins
# ---------------------------------------------------------------------------

def build_duration_bins(
    bin_minutes: int, max_minutes: int
) -> tuple[list, list, list]:
    """Return (bin_edges, labels, w_upper) for duration histogramming.

    w_upper contains the upper bound of each bin (used as W in BW model).
    """
    if max_minutes % bin_minutes != 0:
        raise ValueError("max_minutes должен делиться на bin_minutes без остатка.")
    edges = list(range(0, max_minutes + bin_minutes, bin_minutes))
    labels = [f"{edges[i]}-{edges[i+1]}" for i in range(len(edges) - 1)]
    w_upper = [edges[i + 1] for i in range(len(edges) - 1)]
    return edges, labels, w_upper


# ---------------------------------------------------------------------------
# BW model
# ---------------------------------------------------------------------------

def _phi(W, E, F, G):
    return E / (1.0 + np.exp(F - G * W))


def _bw(W, E, F, G):
    return 1.0 - 1.0 / np.power(1.0 + W, _phi(W, E, F, G))


def fit_bw_params(
    W: np.ndarray,
    y: np.ndarray,
    e_bounds: tuple[float, float] = (0.01, 20.0),
    f_bounds: tuple[float, float] = (-20.0, 20.0),
    g_bounds: tuple[float, float] = (-2.0, 2.0),
) -> tuple[float, float, float]:
    """Fit BW model parameters E, F, G to observed data.

    Primary method: scipy.optimize.curve_fit.
    Fallback: random search + simulated annealing.

    Returns (E, F, G).
    """
    W = np.asarray(W, dtype=float)
    y = np.asarray(y, dtype=float)
    mask = np.isfinite(W) & np.isfinite(y)
    W, y = W[mask], y[mask]

    if len(W) < 6:
        return 2.0, 0.0, 0.1

    try:
        from scipy.optimize import curve_fit  # type: ignore

        p0 = [2.0, 0.0, 0.05]
        bounds = (
            [e_bounds[0], f_bounds[0], g_bounds[0]],
            [e_bounds[1], f_bounds[1], g_bounds[1]],
        )
        popt, _ = curve_fit(_bw, W, y, p0=p0, bounds=bounds, maxfev=50000)
        return float(popt[0]), float(popt[1]), float(popt[2])
    except Exception:
        pass

    # Fallback: random search
    rng = np.random.default_rng(42)

    def sse(E_, F_, G_):
        return float(np.sum((_bw(W, E_, F_, G_) - y) ** 2))

    best_sse = float("inf")
    best_E, best_F, best_G = 2.0, 0.0, 0.1

    for _ in range(6000):
        E_ = rng.uniform(*e_bounds)
        F_ = rng.uniform(*f_bounds)
        G_ = rng.uniform(*g_bounds)
        val = sse(E_, F_, G_)
        if val < best_sse:
            best_sse, best_E, best_F, best_G = val, E_, F_, G_

    step_E = (e_bounds[1] - e_bounds[0]) / 40
    step_F = (f_bounds[1] - f_bounds[0]) / 40
    step_G = (g_bounds[1] - g_bounds[0]) / 40

    for _ in range(2000):
        cE = float(np.clip(rng.normal(best_E, step_E), *e_bounds))
        cF = float(np.clip(rng.normal(best_F, step_F), *f_bounds))
        cG = float(np.clip(rng.normal(best_G, step_G), *g_bounds))
        val = sse(cE, cF, cG)
        if val < best_sse:
            best_sse = val
            best_E, best_F, best_G = cE, cF, cG
            step_E *= 0.995
            step_F *= 0.995
            step_G *= 0.995

    return best_E, best_F, best_G


def calc_metrics(y_true: np.ndarray, y_pred: np.ndarray) -> dict:
    """Return MAE, RMSE, R² for model evaluation."""
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mask = np.isfinite(y_true) & np.isfinite(y_pred)
    y_true, y_pred = y_true[mask], y_pred[mask]

    if len(y_true) == 0:
        return {"MAE": None, "RMSE": None, "R2": None}

    mae = float(np.mean(np.abs(y_true - y_pred)))
    rmse = float(np.sqrt(np.mean((y_true - y_pred) ** 2)))
    ss_res = float(np.sum((y_true - y_pred) ** 2))
    ss_tot = float(np.sum((y_true - np.mean(y_true)) ** 2))
    r2 = float(1 - ss_res / ss_tot) if ss_tot > 0 else None
    return {"MAE": mae, "RMSE": rmse, "R2": r2}


# ---------------------------------------------------------------------------
# Main stage runner
# ---------------------------------------------------------------------------

def run_eva(cfg: Config, df: pd.DataFrame) -> dict:
    """Compute hourly distributions, duration histograms and BW fits.

    Returns a dict keyed by OD-pair name, each value is a dict with:
        hours_df   – DataFrame(Час, Передвижения, Доля)
        transports – dict keyed by transport group name, each containing:
            dur_df     – duration histogram DataFrame
            E, F, G    – BW parameters
            metrics    – {MAE, RMSE, R2}
    """
    cols = cfg.columns
    bw_cfg = cfg.bw_model

    dep_col = cols.departure_time
    arr_col = cols.arrival_time

    edges, labels, w_upper = build_duration_bins(
        bw_cfg.bin_minutes, bw_cfg.max_minutes
    )

    df = df.copy()
    df["_hour"] = df[dep_col].apply(extract_hour)

    dep_m = df[dep_col].apply(time_to_minutes)
    arr_m = df[arr_col].apply(time_to_minutes)

    durations = []
    for d, a in zip(dep_m, arr_m):
        if d is None or a is None:
            durations.append(None)
            continue
        x = int(a) - int(d)
        if x < 0:
            x += 24 * 60
        durations.append(x)
    df["_dur_min"] = pd.to_numeric(pd.Series(durations, index=df.index), errors="coerce")

    results = {}

    for od_pair, sub_pair in df.groupby("_od_pair"):
        # -- hourly distribution for the whole OD pair --
        sub_h = sub_pair.dropna(subset=["_hour"]).copy()
        sub_h["_hour"] = pd.to_numeric(sub_h["_hour"], errors="coerce")
        sub_h = sub_h.dropna(subset=["_hour"])
        sub_h["_hour"] = sub_h["_hour"].astype(int)

        counts_h = (
            sub_h["_hour"]
            .value_counts()
            .reindex(range(24), fill_value=0)
            .sort_index()
        )
        total_h = int(counts_h.sum()) or 1
        hours_df = pd.DataFrame({
            "Час": counts_h.index,
            "Передвижения": counts_h.values,
            "Доля": (counts_h.values / total_h).round(4),
        })

        # -- per-transport duration + BW --
        transport_results = {}

        for tg, sub_tg in sub_pair.groupby("_transport_group"):
            sub_d = sub_tg.dropna(subset=["_dur_min"]).copy()
            sub_d["_dur_min"] = pd.to_numeric(sub_d["_dur_min"], errors="coerce")
            sub_d = sub_d.dropna(subset=["_dur_min"])
            sub_d = sub_d[
                (sub_d["_dur_min"] >= 0) & (sub_d["_dur_min"] <= bw_cfg.max_minutes)
            ].copy()

            # Clamp max value slightly below upper bound for pd.cut
            sub_d["_dur_adj"] = sub_d["_dur_min"].where(
                sub_d["_dur_min"] < bw_cfg.max_minutes,
                bw_cfg.max_minutes - 1e-6,
            )

            bin_counts = (
                pd.cut(sub_d["_dur_adj"], bins=edges, right=False, labels=labels, include_lowest=True)
                .value_counts()
                .reindex(labels, fill_value=0)
                .astype(int)
            )
            total_d = int(bin_counts.sum()) or 1

            dur_df = pd.DataFrame({
                "Интервал (мин)": labels,
                "W (верх, мин)": w_upper,
                "Количество": bin_counts.values,
                "Доля": (bin_counts.values / total_d).round(4),
            })

            # "Разница с предыдущим (от 1)" – share remaining after cumulative
            prev_cum = dur_df["Доля"].cumsum().shift(1, fill_value=0)
            dur_df["Разница с предыдущим (от 1)"] = (1 - prev_cum).clip(lower=0).round(4)

            W_arr = dur_df["W (верх, мин)"].to_numpy(dtype=float)
            y_arr = dur_df["Разница с предыдущим (от 1)"].to_numpy(dtype=float)

            E, F, G = fit_bw_params(
                W_arr, y_arr,
                e_bounds=bw_cfg.e_bounds,
                f_bounds=bw_cfg.f_bounds,
                g_bounds=bw_cfg.g_bounds,
            )
            y_pred = _bw(W_arr, E, F, G)

            dur_df["BW (модель)"] = np.round(y_pred, 4)
            dur_df["|ошибка|"] = np.round(np.abs(y_pred - y_arr), 4)

            transport_results[str(tg)] = {
                "dur_df": dur_df,
                "E": round(E, 6),
                "F": round(F, 6),
                "G": round(G, 6),
                "metrics": calc_metrics(y_arr, y_pred),
                "sub_df": sub_tg,   # kept for statistics block in Excel report
            }

        results[str(od_pair)] = {
            "hours_df": hours_df,
            "transports": transport_results,
            "sub_df": sub_pair,     # full pair data for stats block
        }

    print(f"  EVA завершён. Обработано OD-пар: {len(results)}")
    return results
