import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import io
import json
import os
import random
import re
import string
import subprocess
import sys
import time
import urllib.parse
import warnings
import numpy as np
import pandas as pd
import requests
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

# Tự động kiểm tra và cài đặt các thư viện đồ họa, Excel & WebSocket nếu thiếu
for pkg_name, import_name in [
    ("matplotlib", "matplotlib"),
    ("Pillow", "PIL"),
    ("openpyxl", "openpyxl"),
    ("websocket-client", "websocket"),
]:
    try:
        __import__(import_name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from websocket import create_connection

warnings.filterwarnings("ignore")

# =========================================================================
# CẤU HÌNH KẾT NỐI TỰ ĐỘNG TỪ GITHUB SECRETS (HOẶC DÁN TRỰC TIẾP KHI CHẠY PYCHARM)
# =========================================================================
DISCORD_WEBHOOK_URL = (
    os.environ.get("DISCORD_WEBHOOK_URL", "https://discord.com/api/webhooks/1560172581124833294/tFvPQ51-MsxOMGz5JQL5bbGV18Vp8vx71ctLFIPff8kD-xtE8_JK22iBDLj7bNxBvW4M").strip().strip('"').strip("'")
)
GOOGLE_SHEETS_WEBHOOK_URL = (
    os.environ.get("GOOGLE_SHEETS_WEBHOOK_URL", "https://script.google.com/macros/s/AKfycbwQ3orxyV8yVRvChxuGsWIrSpeePae6AVUrJud8m1srbcuK6zZHZ5OOXi4vlLg-lFmL/exec").strip().strip('"').strip("'")
)

MIN_VALUE_TRADED_BILLION_VND = 5.0
TOP_N_TO_TRACK_PER_ALGO = 5
CSV_EXPORT_FILENAME = "toan_bo_co_phieu_vn_30_algo.csv"
EXCEL_TRACKER_FILENAME = "thong_ke_xac_suat_thang_30_algo.xlsx"
CHARTS_DIR = "charts_t0"
EMBED_THUMBNAIL_IN_EXCEL = True

LEN = 20
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Content-Type": "application/json",
}


def get_vn_now():
    return datetime.datetime.now(VN_TZ)


def get_vn_today():
    return get_vn_now().date()


def clean_for_json(obj):
    if isinstance(obj, dict):
        return {k: clean_for_json(v) for k, v in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [clean_for_json(x) for x in obj]
    elif isinstance(obj, (np.bool_, bool)):
        return bool(obj)
    elif isinstance(obj, (np.integer, int)):
        return int(obj)
    elif isinstance(obj, (np.floating, float)):
        if np.isnan(obj) or np.isinf(obj):
            return ""
        return float(obj)
    elif obj is None or (isinstance(obj, float) and pd.isna(obj)):
        return ""
    return str(obj) if not isinstance(obj, str) else obj


# =========================================================================
# 1. BẢNG PHÂN LOẠI NGÀNH & HÀM CHUẨN HÓA KHỐI LƯỢNG THEO GIỜ GIAO DỊCH (UTC+7)
# =========================================================================
CORE_TICKER_SECTOR = {
    **{
        tk: "Ngân hàng"
        for tk in [
            "VCB", "BID", "CTG", "TCB", "MBB", "VPB", "ACB", "STB", "HDB",
            "LPB", "VIB", "SHB", "TPB", "SSB", "EIB", "OCB", "MSB", "NAB",
            "BAB", "BVB", "ABB", "NVB", "KLB", "PGB", "SGB", "VAB", "VBB",
        ]
    },
    **{
        tk: "Chứng khoán"
        for tk in [
            "SSI", "VND", "VCI", "HCM", "SHS", "MBS", "FTS", "BSI", "CTS",
            "VIX", "AGR", "ORS", "VDS", "BVS", "TVS", "DSC", "TCI", "AAS",
            "SBS", "EVS", "APG", "WSS", "VFS", "BMS", "APS", "PSI",
        ]
    },
    **{
        tk: "Thép & VLXD"
        for tk in [
            "HPG", "HSG", "NKG", "VGS", "TLH", "SMC", "TVN", "TIS", "POM",
            "GDA", "HT1", "BCC", "KSB", "DHA", "VCS", "PTB", "BMP", "NTP",
        ]
    },
    **{
        tk: "BĐS Khu Công Nghiệp"
        for tk in [
            "KBC", "IDC", "SZC", "VGC", "GVR", "BCM", "PHR", "DPR", "NTC",
            "SIP", "LHG", "TIP", "ITA", "DTD", "D2D", "VRG", "MH3",
        ]
    },
    **{
        tk: "Đầu tư công & Xây dựng"
        for tk in [
            "VCG", "HHV", "LCG", "FCN", "CII", "CTD", "HBC", "PC1", "TV2",
            "DPG", "C4G", "G36", "HUT", "SC5", "PHC", "CSC",
        ]
    },
    **{
        tk: "Hóa chất & Phân bón"
        for tk in [
            "DGC", "DCM", "DPM", "CSV", "LAS", "BFC", "DDV", "SFG", "PMB",
            "DRC", "CSM", "SRC", "AAA", "APF",
        ]
    },
    **{
        tk: "Cảng biển & Vận tải"
        for tk in [
            "GMD", "HAH", "PVT", "VSC", "VOS", "VIP", "TCO", "SGP", "PHP",
            "DSP", "TCL", "SWC", "VTO", "GSP", "VTP", "SCS", "ACV", "HVN",
            "VJC", "AST",
        ]
    },
    **{
        tk: "Thủy sản & Xuất khẩu"
        for tk in [
            "VHC", "ANV", "IDI", "CMX", "FMC", "MPC", "ACL", "ASM", "TNG",
            "GIL", "MSH", "TCM", "VGT", "STK", "ADS", "HTG", "PPH",
        ]
    },
}

TV_SECTOR_MAP = {
    "Finance": "Bất động sản & Tài chính",
    "Commercial Services": "Dịch vụ & Thương mại",
    "Distribution Services": "Bán lẻ & Phân phối",
    "Retail Trade": "Bán lẻ & Tiêu dùng",
    "Consumer Non-Durables": "Thực phẩm & Tiêu dùng",
    "Consumer Durables": "Hàng tiêu dùng",
    "Consumer Services": "Dịch vụ Tiêu dùng",
    "Electronic Technology": "Công nghệ & Viễn thông",
    "Technology Services": "Công nghệ & Phần mềm",
    "Communications": "Viễn thông",
    "Energy Minerals": "Dầu khí & Năng lượng",
    "Utilities": "Điện & Năng lượng",
    "Process Industries": "Hóa chất & Chế biến",
    "Non-Energy Minerals": "Khai khoáng & VLXD",
    "Producer Manufacturing": "Sản xuất Công nghiệp",
    "Industrial Services": "Đầu tư công & Xây dựng",
    "Transportation": "Cảng biển & Vận tải",
    "Health Technology": "Y tế & Dược phẩm",
    "Health Services": "Y tế & Dược phẩm",
}


def classify_vn_sector(ticker, tv_sector, description):
    if ticker in CORE_TICKER_SECTOR:
        return CORE_TICKER_SECTOR[ticker]
    desc_up = str(description).upper()
    if "NGÂN HÀNG" in desc_up or "BANK" in desc_up:
        return "Ngân hàng"
    if "CHỨNG KHOÁN" in desc_up or "SECURITIES" in desc_up:
        return "Chứng khoán"
    if (
        "BẤT ĐỘNG SẢN" in desc_up
        or "ĐỊA ỐC" in desc_up
        or "NHÀ " in desc_up
        or "LAND" in desc_up
    ):
        return "Bất động sản Dân dụng"
    if "KHU CÔNG NGHIỆP" in desc_up or "CAO SU" in desc_up:
        return "BĐS Khu Công Nghiệp"
    if "THÉP" in desc_up or "XI MĂNG" in desc_up or "ĐÁ " in desc_up:
        return "Thép & VLXD"
    if (
        "DẦU KHÍ" in desc_up
        or "PETRO" in desc_up
        or "KHÍ " in desc_up
        or "ĐIỆN" in desc_up
    ):
        return "Dầu khí & Năng lượng"
    if (
        "CẢNG" in desc_up
        or "VẬN TẢI" in desc_up
        or "HÀNG HẢI" in desc_up
        or "LOGISTICS" in desc_up
    ):
        return "Cảng biển & Vận tải"
    if (
        "THỦY SẢN" in desc_up
        or "DỆT MAY" in desc_up
        or "MAY " in desc_up
        or "GỖ " in desc_up
    ):
        return "Thủy sản & Xuất khẩu"
    if (
        "XÂY DỰNG" in desc_up
        or "CẦU ĐƯỜNG" in desc_up
        or "HẠ TẦNG" in desc_up
        or "LẮP MÁY" in desc_up
    ):
        return "Đầu tư công & Xây dựng"
    if "PHÂN BÓN" in desc_up or "HÓA CHẤT" in desc_up or "NHỰA" in desc_up:
        return "Hóa chất & Phân bón"
    if tv_sector == "Finance":
        return "Bất động sản Dân dụng"
    return TV_SECTOR_MAP.get(tv_sector, "Sản xuất & Khác")


def get_intraday_rvol_multiplier():
    now = get_vn_now()
    if now.weekday() >= 5:
        return 1.0, "Đã đóng cửa (Cuối tuần)"

    cur_mins = now.hour * 60 + now.minute
    m_start, m_end = 9 * 60 + 15, 11 * 60 + 30
    a_start, a_end = 13 * 60, 14 * 60 + 45

    if cur_mins < m_start or cur_mins >= a_end:
        return 1.0, "Ngoài giờ / Đóng cửa (Chuẩn EOD)"
    elif m_start <= cur_mins <= m_end:
        elapsed = max(cur_mins - m_start, 15)
        mult = min(240.0 / elapsed, 4.0)
        return mult, f"Phiên Sáng ({elapsed}/240p - Hệ số RVOL x{mult:.2f})"
    elif m_end < cur_mins < a_start:
        mult = 240.0 / 135.0
        return mult, f"Nghỉ Trưa (135/240p - Hệ số RVOL x{mult:.2f})"
    else:
        elapsed = 135 + (cur_mins - a_start)
        mult = min(240.0 / max(elapsed, 135), 1.8)
        return mult, f"Phiên Chiều ({elapsed}/240p - Hệ số RVOL x{mult:.2f})"


def fetch_foreign_net_flow_map():
    f_map = {}
    try:
        url = "https://finfo-api.vndirect.com.vn/v4/stock_prices?q=date:gte:2025-01-01&size=1800&sort=date:desc"
        r = requests.get(url, headers=HEADERS, timeout=4).json()
        for item in r.get("data", []):
            tk = str(item.get("code", "")).strip()
            if tk and tk not in f_map:
                f_val = (
                    float(item.get("foreignBuyVal", 0) or 0)
                    - float(item.get("foreignSellVal", 0) or 0)
                ) / 1e9
                f_map[tk] = f_val
    except Exception:
        pass
    return f_map


# =========================================================================
# 2. MODULE QUÉT TOÀN BỘ ~1.600+ CỔ PHIẾU & 4 TẦNG LẤY DỮ LIỆU NẾN OHLCV
# =========================================================================


def fetch_entire_vietnam_stock_market():
    url = "https://scanner.tradingview.com/vietnam/scan"
    columns_40 = [
        "name", "description", "exchange", "sector", "close", "open", "high",
        "low", "volume", "Value.Traded", "market_cap_basic", "change",
        "Perf.W", "Perf.1M", "Perf.3M", "relative_volume_10d_calc",
        "average_volume_30d_calc", "RSI", "RSI7", "SMA10", "SMA20", "SMA50",
        "SMA200", "EMA10", "EMA20", "EMA50", "MACD.macd", "MACD.signal",
        "BB.upper", "BB.lower", "ADX", "Stoch.K", "Stoch.D", "CCI20", "VWMA",
        "ATR", "price_52_week_high", "price_52_week_low",
        "price_earnings_ttm", "return_on_equity",
    ]

    payload = {
        "filter": [
            {"left": "type", "operation": "equal", "right": "stock"},
            {
                "left": "exchange",
                "operation": "in_range",
                "right": ["HOSE", "HNX", "UPCOM"],
            },
        ],
        "options": {"lang": "vi"},
        "markets": ["vietnam"],
        "symbols": {"query": {"types": []}, "tickers": []},
        "columns": columns_40,
        "sort": {"sortBy": "Value.Traded", "sortOrder": "desc"},
        "range": [0, 2500],
    }

    rvol_mult, session_label = get_intraday_rvol_multiplier()
    foreign_map = fetch_foreign_net_flow_map()
    all_rows = []

    try:
        r = requests.post(url, json=payload, headers=HEADERS, timeout=15).json()
        for item in r.get("data", []):
            d = item.get("d", [])
            if len(d) < 40:
                continue
            ticker = str(d[0]).strip()
            if len(ticker) > 3:
                continue

            desc = str(d[1] or ticker)
            exch = str(d[2] or "HOSE")
            tv_sec = str(d[3] or "")
            close_raw = float(d[4]) if d[4] is not None else 0.0
            if close_raw <= 0:
                continue

            open_raw = float(d[5]) if d[5] is not None else close_raw
            high_raw = float(d[6]) if d[6] is not None else close_raw
            low_raw = float(d[7]) if d[7] is not None else close_raw
            vol = float(d[8]) if d[8] is not None else 0.0

            val_traded_raw = (
                float(d[9]) if d[9] is not None else (close_raw * vol)
            )
            val_traded_billion = (val_traded_raw / 1e9) * (
                rvol_mult if rvol_mult > 1.0 else 1.0
            )
            val_traded_actual = val_traded_raw / 1e9
            mcap_billion = (float(d[10]) / 1e9) if d[10] is not None else 0.0

            r1d = float(d[11]) if d[11] is not None else 0.0
            r5d = float(d[12]) if d[12] is not None else r1d
            r20d = float(d[13]) if d[13] is not None else r5d
            r60d = float(d[14]) if d[14] is not None else r20d

            avg_vol_30d = float(d[16]) if d[16] is not None else vol
            if d[15] is not None and float(d[15]) > 0:
                rvol_raw = float(d[15])
            elif avg_vol_30d > 0:
                rvol_raw = vol / avg_vol_30d
            else:
                rvol_raw = 0.0

            rvol = (
                max(rvol_raw, rvol_raw * rvol_mult)
                if rvol_mult > 1.0
                else rvol_raw
            )

            rsi = float(d[17]) if d[17] is not None else 50.0
            rsi7 = float(d[18]) if d[18] is not None else rsi
            sma10 = float(d[19]) if d[19] is not None else close_raw
            sma20 = float(d[20]) if d[20] is not None else close_raw
            sma50 = float(d[21]) if d[21] is not None else sma20
            sma200 = float(d[22]) if d[22] is not None else sma50
            ema10 = float(d[23]) if d[23] is not None else sma10
            ema20 = float(d[24]) if d[24] is not None else sma20
            ema50 = float(d[25]) if d[25] is not None else sma50

            macd_val = float(d[26]) if d[26] is not None else 0.0
            macd_sig = float(d[27]) if d[27] is not None else 0.0
            macd_hist = macd_val - macd_sig

            bb_upper = float(d[28]) if d[28] is not None else (close_raw * 1.05)
            bb_lower = float(d[29]) if d[29] is not None else (close_raw * 0.95)
            bb_width_pct = (
                ((bb_upper - bb_lower) / sma20) * 100 if sma20 > 0 else 12.0
            )

            adx = float(d[30]) if d[30] is not None else 20.0
            stoch_k = float(d[31]) if d[31] is not None else 50.0
            stoch_d = float(d[32]) if d[32] is not None else 50.0
            cci20 = float(d[33]) if d[33] is not None else 0.0
            vwma = float(d[34]) if d[34] is not None else sma20
            atr = float(d[35]) if d[35] is not None else (close_raw * 0.025)

            high_52w = float(d[36]) if d[36] is not None else close_raw
            low_52w = float(d[37]) if d[37] is not None else (close_raw * 0.75)
            pe_ttm = float(d[38]) if d[38] is not None else 0.0
            roe = float(d[39]) if d[39] is not None else 0.0

            prev_close = (
                close_raw / (1.0 + r1d / 100.0)
                if (1.0 + r1d / 100.0) != 0
                else close_raw
            )
            gap_up_pct = (
                ((open_raw - prev_close) / prev_close) * 100
                if prev_close > 0
                else 0.0
            )
            intraday_range_pct = (
                ((high_raw - low_raw) / low_raw) * 100 if low_raw > 0 else 0.0
            )

            candle_range = high_raw - low_raw
            close_pos = (
                (close_raw - low_raw) / candle_range
                if candle_range > 0
                else 0.5
            )
            body_ratio = (
                (close_raw - open_raw) / candle_range
                if candle_range > 0
                else 0.0
            )
            mf_multiplier = (
                ((close_raw - low_raw) - (high_raw - close_raw)) / candle_range
                if candle_range > 0
                else 0.0
            )

            above_ma20 = close_raw > sma20 if sma20 > 0 else False
            above_ma50 = close_raw > sma50 if sma50 > 0 else False
            above_ma200 = close_raw > sma200 if sma200 > 0 else False

            dist_ma20_pct = (
                ((close_raw - sma20) / sma20) * 100 if sma20 > 0 else 0.0
            )
            dist_ma50_pct = (
                ((close_raw - sma50) / sma50) * 100 if sma50 > 0 else 0.0
            )
            dist_ma200_pct = (
                ((close_raw - sma200) / sma200) * 100 if sma200 > 0 else 0.0
            )
            dist_vwma_pct = (
                ((close_raw - vwma) / vwma) * 100 if vwma > 0 else 0.0
            )
            sma20_50_diff_pct = (
                ((sma20 - sma50) / sma50) * 100 if sma50 > 0 else 0.0
            )

            pct_from_52w_high = (
                (close_raw / high_52w) * 100 if high_52w > 0 else 80.0
            )
            gain_from_52w_low = (
                ((close_raw - low_52w) / low_52w) * 100 if low_52w > 0 else 20.0
            )

            atr_pct = float(
                np.clip(
                    (atr / close_raw) * 100 if close_raw > 0 else 2.5, 1.8, 4.5
                )
            )
            stop_loss_pct = float(np.clip(1.8 * atr_pct, 3.2, 6.5))
            target_pct = float(np.clip(2.2 * stop_loss_pct, 6.0, 15.0))
            atr_break_ratio = r1d / atr_pct if atr_pct > 0 else 0.0

            price_k = close_raw / 1000.0 if close_raw > 500 else close_raw
            target_price_k = round(price_k * (1.0 + target_pct / 100.0), 2)
            stop_price_k = round(price_k * (1.0 - stop_loss_pct / 100.0), 2)

            is_bo_cung = (intraday_range_pct > 5.5 and abs(r1d) < 0.35) or (
                bb_width_pct < 2.2 and rvol_raw > 0.85 and rvol_raw < 1.15
            )

            ceil_thresh = (
                6.5 if exch == "HOSE" else (8.8 if exch == "HNX" else 12.5)
            )
            if r1d >= ceil_thresh:
                buy_zone = "🟣 ĐÓNG TRẦN (Không đua)"
                chase_risk = True
            elif r1d >= 5.2:
                buy_zone = "⚠️ KÉO CAO (>5%)"
                chase_risk = True
            else:
                buy_zone = "🟢 VÙNG MUA CHUẨN"
                chase_risk = False

            foreign_net_b = foreign_map.get(ticker, 0.0)
            vn_sector = classify_vn_sector(ticker, tv_sec, desc)

            all_rows.append({
                "ticker": ticker,
                "exchange": exch,
                "company": desc,
                "sector": vn_sector,
                "price": price_k,
                "close_raw": close_raw,
                "open_raw": open_raw,
                "high_raw": high_raw,
                "low_raw": low_raw,
                "volume": vol,
                "val_billion": val_traded_actual,
                "val_proj_billion": val_traded_billion,
                "mcap_billion": mcap_billion,
                "r1d": r1d,
                "r5d": r5d,
                "r20d": r20d,
                "r60d": r60d,
                "rvol": rvol,
                "rvol_raw": rvol_raw,
                "rsi": rsi,
                "rsi7": rsi7,
                "sma10": sma10,
                "sma20": sma20,
                "sma50": sma50,
                "sma200": sma200,
                "ema10": ema10,
                "ema20": ema20,
                "ema50": ema50,
                "above_ma20": above_ma20,
                "above_ma50": above_ma50,
                "above_ma200": above_ma200,
                "dist_ma20_pct": dist_ma20_pct,
                "dist_ma50_pct": dist_ma50_pct,
                "dist_ma200_pct": dist_ma200_pct,
                "dist_vwma_pct": dist_vwma_pct,
                "sma20_50_diff_pct": sma20_50_diff_pct,
                "macd_val": macd_val,
                "macd_hist": macd_hist,
                "bb_upper": bb_upper,
                "bb_width_pct": bb_width_pct,
                "adx": adx,
                "stoch_k": stoch_k,
                "stoch_d": stoch_d,
                "cci20": cci20,
                "vwma": vwma,
                "atr_pct": atr_pct,
                "target_pct": target_pct,
                "stop_loss_pct": stop_loss_pct,
                "target_price": target_price_k,
                "stop_price": stop_price_k,
                "atr_break_ratio": atr_break_ratio,
                "pct_52w_high": pct_from_52w_high,
                "gain_52w_low": gain_from_52w_low,
                "pe": pe_ttm,
                "roe": roe,
                "gap_up_pct": gap_up_pct,
                "intraday_range_pct": intraday_range_pct,
                "close_pos": close_pos,
                "body_ratio": body_ratio,
                "mf_mult": mf_multiplier,
                "is_bo_cung": is_bo_cung,
                "buy_zone": buy_zone,
                "chase_risk": chase_risk,
                "foreign_net_b": foreign_net_b,
                "session_label": session_label,
            })
    except Exception as e:
        print(f"Lỗi quét Scanner: {e}")

    return pd.DataFrame(all_rows)


def fetch_vn_api_ohlcv(symbol, n_bars=65):
    """Lấy dữ liệu Nến Nhật (Open, High, Low, Close, Volume) đa tầng: DNSE -> VNDirect -> Yahoo Finance (.VN)"""
    now_ts = int(time.time())
    from_ts = now_ts - (n_bars * 2 * 86400)
    # Tầng 1 & 2: Cổng API Việt Nam
    for url in [
        f"https://services.entrade.com.vn/chart-api/v2/ohlcs/stock?from={from_ts}&to={now_ts}&symbol={symbol}&resolution=1D",
        f"https://dchart-api.vndirect.com.vn/dchart/history?resolution=D&symbol={symbol}&from={from_ts}&to={now_ts}",
    ]:
        try:
            r = requests.get(url, headers=HEADERS, timeout=4).json()
            if "c" in r and len(r["c"]) >= 10:
                c_list = [float(x) for x in r["c"]]
                o_list = [float(x) for x in r.get("o", c_list)]
                h_list = [float(x) for x in r.get("h", c_list)]
                l_list = [float(x) for x in r.get("l", c_list)]
                v_list = [float(x) for x in r.get("v", [0] * len(c_list))]
                t_list = r.get("t", [0] * len(c_list))
                return (
                    pd.DataFrame({
                        "time": t_list,
                        "open": o_list,
                        "high": h_list,
                        "low": l_list,
                        "close": c_list,
                        "volume": v_list,
                    })
                    .iloc[-n_bars:]
                    .reset_index(drop=True)
                )
        except Exception:
            pass

    # Tầng 3: Yahoo Finance (.VN) - Cực nhạy khi chạy trên máy chủ GitHub Actions tại Mỹ
    if len(symbol) == 3:
        df_yf = fetch_yahoo_ohlcv(f"{symbol}.VN", n_bars)
        if df_yf is not None and len(df_yf) >= 10:
            return df_yf

    return None


def fetch_yahoo_ohlcv(ticker, n_bars=65):
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=3mo"
    try:
        r = requests.get(url, headers=HEADERS, timeout=5).json()
        res = r["chart"]["result"][0]
        timestamps = res.get("timestamp", [])
        q = res["indicators"]["quote"][0]
        closes = q.get("close", [])
        opens = q.get("open", closes)
        highs = q.get("high", closes)
        lows = q.get("low", closes)
        vols = q.get("volume", [0] * len(closes))
        rows = []
        for idx, c_v in enumerate(closes):
            if c_v is not None:
                c_f = float(c_v)
                o_f = float(opens[idx]) if idx < len(opens) and opens[idx] is not None else c_f
                h_f = float(highs[idx]) if idx < len(highs) and highs[idx] is not None else max(o_f, c_f)
                l_f = float(lows[idx]) if idx < len(lows) and lows[idx] is not None else min(o_f, c_f)
                v_f = float(vols[idx]) if idx < len(vols) and vols[idx] is not None else 0.0
                t_f = int(timestamps[idx]) if idx < len(timestamps) and timestamps[idx] is not None else 0
                rows.append((t_f, o_f, h_f, l_f, c_f, v_f))
        if len(rows) >= 5:
            return (
                pd.DataFrame(
                    rows,
                    columns=["time", "open", "high", "low", "close", "volume"],
                )
                .iloc[-n_bars:]
                .reset_index(drop=True)
            )
    except Exception:
        pass
    return None


def fetch_tv_websocket_ohlcv(requests_dict, n_bars=65):
    """Tầng 4: Lấy dữ liệu Nến Nhật trực tiếp qua TradingView WebSocket cho các mã UPCOM/HNX"""
    results = {}
    ws = None
    try:
        ws = create_connection(
            "wss://data.tradingview.com/socket.io/websocket",
            headers={
                "Origin": "https://data.tradingview.com",
                "User-Agent": HEADERS["User-Agent"],
            },
            timeout=10,
        )
        auth_msg = json.dumps(
            {"m": "set_auth_token", "p": ["unauthorized_user_token"]}
        )
        ws.send(f"~m~{len(auth_msg)}~m~{auth_msg}")

        idx = 0
        for key, candidates in requests_dict.items():
            for sym, exch in candidates:
                idx += 1
                session = "cs_" + "".join(
                    random.choices(string.ascii_lowercase + string.digits, k=12)
                )
                sym_id, ser_id = f"sym_{idx}", f"s_{idx}"
                for func, args in [
                    ("chart_create_session", [session, ""]),
                    (
                        "resolve_symbol",
                        [
                            session,
                            sym_id,
                            f'={{"symbol":"{exch}:{sym}","adjustment":"splits"}}',
                        ],
                    ),
                    (
                        "create_series",
                        [session, ser_id, ser_id, sym_id, "1D", n_bars, ""],
                    ),
                ]:
                    m_str = json.dumps({"m": func, "p": args})
                    ws.send(f"~m~{len(m_str)}~m~{m_str}")

                rows = []
                while True:
                    try:
                        res = ws.recv()
                        if res.startswith("~m~") and "~h~" in res:
                            ws.send(res)
                            continue
                        for pkt in re.split(r"~m~\d+~m~", res):
                            if not pkt:
                                continue
                            try:
                                obj = json.loads(pkt)
                                if isinstance(obj, dict) and obj.get("m") in [
                                    "timescale_update",
                                    "du",
                                ]:
                                    p_data = obj.get("p", [])
                                    if len(p_data) > 1 and isinstance(
                                        p_data[1], dict
                                    ):
                                        for bar in (
                                            p_data[1]
                                            .get(ser_id, {})
                                            .get("s", [])
                                        ):
                                            v = bar.get("v", [])
                                            if len(v) >= 5:
                                                rows.append((
                                                    int(float(v[0])),
                                                    float(v[1]),
                                                    float(v[2]),
                                                    float(v[3]),
                                                    float(v[4]),
                                                    float(v[5])
                                                    if len(v) >= 6
                                                    else 0.0,
                                                ))
                            except Exception:
                                pass
                        if (
                            rows
                            or "series_completed" in res
                            or "symbol_error" in res
                        ):
                            break
                    except Exception:
                        break
                if len(rows) >= 10 and rows[-1][4] > 0:
                    results[key] = (
                        pd.DataFrame(
                            rows,
                            columns=[
                                "time", "open", "high", "low", "close", "volume"
                            ],
                        )
                        .iloc[-n_bars:]
                        .reset_index(drop=True)
                    )
                    break
    except Exception:
        pass
    finally:
        if ws:
            ws.close()
    return results


def build_synthetic_ohlcv_from_scanner(st_row, n_bars=30):
    """Tầng 5 (Đảm bảo 100% không mã nào bị thiếu biểu đồ nến): Tái tạo chuỗi nến từ các mốc giá thực của Scanner"""
    c0 = float(st_row["close_raw"])
    o0 = float(st_row["open_raw"])
    h0 = float(st_row["high_raw"])
    l0 = float(st_row["low_raw"])
    v0 = float(st_row["volume"])
    rvol = max(float(st_row.get("rvol", 1.0)), 0.3)
    avg_v = v0 / rvol if rvol > 0 else v0

    r1d = float(st_row.get("r1d", 0.0))
    r5d = float(st_row.get("r5d", 0.0))
    r20d = float(st_row.get("r20d", 0.0))
    atr_p = float(st_row.get("atr_pct", 2.5)) / 100.0

    c_1 = c0 / (1.0 + r1d / 100.0) if r1d != -100 else c0
    c_5 = c0 / (1.0 + r5d / 100.0) if r5d != -100 else c_1
    c_20 = c0 / (1.0 + r20d / 100.0) if r20d != -100 else c_5

    closes = np.concatenate([
        np.linspace(c_20 * 0.99, c_20, 10),
        np.linspace(c_20, c_5, 14),
        np.linspace(c_5, c_1, 5),
        [c0],
    ])[-n_bars:]

    now_ts = int(time.time())
    rows = []
    for i in range(len(closes)):
        t_val = now_ts - (len(closes) - 1 - i) * 86400
        if i == len(closes) - 1:
            rows.append((t_val, o0, h0, l0, c0, v0))
        else:
            ci = float(closes[i])
            oi = float(closes[i - 1]) if i > 0 else ci * 0.995
            hi = max(oi, ci) * (1.0 + atr_p * 0.45)
            li = min(oi, ci) * (1.0 - atr_p * 0.45)
            vi = avg_v * (0.85 + 0.3 * ((i % 3) / 2.0))
            rows.append((t_val, oi, hi, li, ci, vi))

    return pd.DataFrame(
        rows, columns=["time", "open", "high", "low", "close", "volume"]
    )


def fetch_tv_rest_macro_snapshot():
    url = "https://scanner.tradingview.com/global/scan"
    payload = {
        "symbols": {
            "tickers": [
                "HOSE:VNINDEX",
                "HOSE:VN30",
                "HOSE:VN301!",
                "FX_IDC:USDVND",
                "TVC:DXY",
                "TVC:US10Y",
            ],
            "query": {"types": []},
        },
        "columns": ["close", "change", "Perf.W", "Perf.1M", "SMA20", "volume"],
    }
    snap = {}
    try:
        r = requests.post(url, json=payload, headers=HEADERS, timeout=6).json()
        for item in r.get("data", []):
            sym = item.get("s", "").split(":")[-1]
            d = item.get("d", [])
            if len(d) >= 6 and d[0] is not None:
                snap[sym] = {
                    "close": float(d[0]),
                    "r1d": float(d[1] or 0.0),
                    "r5d": float(d[2] or 0.0),
                    "r20d": float(d[3] or 0.0),
                    "sma20": float(d[4] or d[0]),
                    "volume": float(d[5] or 0.0),
                }
    except Exception:
        pass
    return snap


def load_macro_and_full_market():
    console = Console()
    macro_data = {}

    with console.status(
        "[bold cyan]🌐 [HYBRID CLOUD ENGINE] Đang quét ~1.600+ cổ phiếu & Vĩ mô qua HTTP API...",
        spinner="dots",
    ):
        df_all_stocks = fetch_entire_vietnam_stock_market()

        for k in ["VNINDEX", "VN30", "VN30F1M"]:
            df_idx = fetch_vn_api_ohlcv(k, 60)
            if df_idx is not None:
                macro_data[k] = df_idx

        for k, yf_sym in [
            ("DXY", "DX-Y.NYB"),
            ("US10Y", "^TNX"),
            ("USDVND", "VND=X"),
        ]:
            df_yf = fetch_yahoo_ohlcv(yf_sym, 60)
            if df_yf is not None:
                macro_data[k] = df_yf

        tv_snap = fetch_tv_rest_macro_snapshot()
        map_snap = {
            "VNINDEX": "VNINDEX",
            "VN30": "VN30",
            "VN30F1M": "VN301!",
            "USDVND": "USDVND",
            "DXY": "DXY",
            "US10Y": "US10Y",
        }
        for k, s_key in map_snap.items():
            if k not in macro_data or macro_data[k]["close"].iloc[-1] == 0:
                if s_key in tv_snap:
                    st = tv_snap[s_key]
                    c_now = st["close"]
                    c_1d = (
                        c_now / (1.0 + st["r1d"] / 100.0)
                        if st["r1d"] != -100
                        else c_now
                    )
                    c_5d = (
                        c_now / (1.0 + st["r5d"] / 100.0)
                        if st["r5d"] != -100
                        else c_now
                    )
                    c_20d = (
                        c_now / (1.0 + st["r20d"] / 100.0)
                        if st["r20d"] != -100
                        else c_now
                    )
                    arr = np.linspace(c_20d, c_now, 25)
                    arr[-2] = c_1d
                    arr[-6] = c_5d
                    macro_data[k] = pd.DataFrame(
                        {"close": arr, "volume": [st["volume"]] * 25}
                    )
                else:
                    macro_data[k] = pd.DataFrame(
                        {"close": [0.0] * 25, "volume": [0.0] * 25}
                    )

    return macro_data, df_all_stocks


# =========================================================================
# 3. MODULE VẼ BIỂU ĐỒ NẾN NHẬT (MATPLOTLIB + QUICKCHART CANDLESTICK URL)
# =========================================================================


def build_quickchart_candlestick_url(ticker, exch, df_ohlcv, st_row):
    """
    Tạo URL Biểu đồ Nến Nhật thực thụ (type: 'candlestick' có thân nến Xanh/Đỏ, râu nến High/Low,
    đường MA20, Target & Cắt lỗ) qua QuickChart Short API để hàm =IMAGE() trên Google Sheets hiển thị chuẩn Nến!
    """
    if df_ohlcv is None or len(df_ohlcv) < 5:
        df_ohlcv = build_synthetic_ohlcv_from_scanner(st_row, 25)

    df_c = df_ohlcv.tail(25).copy().reset_index(drop=True)
    scale = 1000.0 if df_c["close"].iloc[-1] > 500 else 1.0

    closes_s = pd.Series(df_c["close"].values / scale)
    ma20_s = closes_s.rolling(15, min_periods=1).mean().values

    now_ms = int(time.time()) * 1000
    candles_data = []
    ma20_data = []
    tp_data = []
    sl_data = []

    tp_val = round(float(st_row.get("target_price", closes_s.iloc[-1] * 1.07)), 2)
    sl_val = round(float(st_row.get("stop_price", closes_s.iloc[-1] * 0.96)), 2)

    for i in range(len(df_c)):
        t_raw = float(df_c["time"].iloc[i]) if "time" in df_c else 0
        x_ms = int(t_raw * 1000) if t_raw > 1e8 else (now_ms - (len(df_c) - 1 - i) * 86400000)
        o_v = round(float(df_c["open"].iloc[i]) / scale, 2)
        h_v = round(float(df_c["high"].iloc[i]) / scale, 2)
        l_v = round(float(df_c["low"].iloc[i]) / scale, 2)
        c_v = round(float(df_c["close"].iloc[i]) / scale, 2)

        candles_data.append({"x": x_ms, "o": o_v, "h": h_v, "l": l_v, "c": c_v})
        ma20_data.append({"x": x_ms, "y": round(float(ma20_s[i]), 2)})
        tp_data.append({"x": x_ms, "y": tp_val})
        sl_data.append({"x": x_ms, "y": sl_val})

    r1d = float(st_row.get("r1d", 0.0))
    rvol = float(st_row.get("rvol", 1.0))

    chart_config = {
        "type": "candlestick",
        "data": {
            "datasets": [
                {
                    "label": f"{ticker} ({closes_s.iloc[-1]:.2f})",
                    "data": candles_data,
                    "color": {
                        "up": "#00e676",
                        "down": "#ff5252",
                        "unchanged": "#999999",
                    },
                    "borderColor": {
                        "up": "#00e676",
                        "down": "#ff5252",
                        "unchanged": "#999999",
                    },
                },
                {
                    "type": "line",
                    "label": "MA20",
                    "data": ma20_data,
                    "borderColor": "#ffd600",
                    "borderWidth": 1.5,
                    "pointRadius": 0,
                    "fill": False,
                },
                {
                    "type": "line",
                    "label": f"TP ({tp_val})",
                    "data": tp_data,
                    "borderColor": "#00e5ff",
                    "borderDash": [4, 4],
                    "borderWidth": 1.2,
                    "pointRadius": 0,
                    "fill": False,
                },
                {
                    "type": "line",
                    "label": f"SL ({sl_val})",
                    "data": sl_data,
                    "borderColor": "#ff5252",
                    "borderDash": [4, 4],
                    "borderWidth": 1.2,
                    "pointRadius": 0,
                    "fill": False,
                },
            ]
        },
        "options": {
            "plugins": {
                "title": {
                    "display": True,
                    "text": f"{ticker} ({exch}) | T+0: {closes_s.iloc[-1]:.2f} ({r1d:+.2f}%) | RVOL: {rvol:.2f}x",
                    "color": "#ffffff",
                    "font": {"size": 11},
                },
                "legend": {
                    "labels": {"color": "#ffffff", "font": {"size": 9}}
                },
            },
            "scales": {
                "x": {"ticks": {"color": "#aaaaaa", "maxTicksLimit": 6}},
                "y": {"ticks": {"color": "#ffffff"}},
            },
        },
    }

    # Gọi cổng tạo Short URL của QuickChart (Version 3 hỗ trợ Candlestick Chart)
    try:
        r_qc = requests.post(
            "https://quickchart.io/chart/create",
            json={
                "version": "3",
                "width": 480,
                "height": 260,
                "backgroundColor": "#0d1117",
                "chart": chart_config,
            },
            timeout=4,
        ).json()
        if r_qc.get("success") and r_qc.get("url"):
            return r_qc["url"]
    except Exception:
        pass

    # Nếu cổng Short URL bận, dùng URL GET trực tiếp với 14 cây nến gần nhất
    chart_config["data"]["datasets"][0]["data"] = candles_data[-14:]
    chart_config["data"]["datasets"][1]["data"] = ma20_data[-14:]
    chart_config["data"]["datasets"][2]["data"] = tp_data[-14:]
    chart_config["data"]["datasets"][3]["data"] = sl_data[-14:]
    encoded = urllib.parse.quote(json.dumps(chart_config, separators=(",", ":")))
    return f"https://quickchart.io/chart?v=3&w=480&h=260&bkg=%230d1117&c={encoded}"


def render_single_t0_candlestick_bytes(
    ticker, exch, sector, df_ohlcv, st_row, algo_tag
):
    """Vẽ Đồ thị Nến Nhật 60 phiên chuẩn TradingView Dark Mode bằng Matplotlib"""
    if df_ohlcv is None or len(df_ohlcv) < 5:
        df_ohlcv = build_synthetic_ohlcv_from_scanner(st_row, 30)

    df_c = df_ohlcv.tail(60).copy().reset_index(drop=True)
    scale = 1000.0 if df_c["close"].iloc[-1] > 500 else 1.0
    o = df_c["open"].values / scale
    h = df_c["high"].values / scale
    l = df_c["low"].values / scale
    c = df_c["close"].values / scale
    v = df_c["volume"].values / 1e6

    s_close = pd.Series(c)
    ma20 = s_close.rolling(20, min_periods=1).mean().values
    ma50 = s_close.rolling(50, min_periods=1).mean().values
    std20 = s_close.rolling(20, min_periods=1).std().fillna(0).values
    bb_up = ma20 + 2.0 * std20
    bb_dn = ma20 - 2.0 * std20
    v_ma20 = pd.Series(v).rolling(20, min_periods=1).mean().values

    fig = plt.figure(figsize=(10, 5.8), dpi=110, facecolor="#0d1117")
    gs = fig.add_gridspec(2, 1, height_ratios=[3.2, 1.0], hspace=0.06)
    ax1 = fig.add_subplot(gs[0], facecolor="#0d1117")
    ax2 = fig.add_subplot(gs[1], facecolor="#0d1117", sharex=ax1)

    n = len(c)
    x = np.arange(n)

    ax1.fill_between(x, bb_dn, bb_up, color="#1f6feb", alpha=0.10)
    ax1.plot(x, ma20, color="#ffd600", linewidth=1.4, label="MA20")
    ax1.plot(x, ma50, color="#00e5ff", linewidth=1.4, label="MA50")

    for i in range(n):
        is_bull = c[i] >= o[i]
        col = "#00e676" if is_bull else "#ff5252"
        ax1.plot([i, i], [l[i], h[i]], color=col, linewidth=1.1, zorder=3)
        body_bottom = min(o[i], c[i])
        body_h = max(abs(c[i] - o[i]), (h.max() - l.min()) * 0.004)
        rect = Rectangle(
            (i - 0.32, body_bottom),
            0.64,
            body_h,
            facecolor=col,
            edgecolor=col,
            zorder=4,
        )
        ax1.add_patch(rect)
        ax2.bar(i, v[i], width=0.64, color=col, alpha=0.80)

    ax2.plot(x, v_ma20, color="#ffd600", linewidth=1.1, label="Vol MA20")

    entry_p = float(c[-1])
    tp_pct = float(st_row.get("target_pct", 7.0))
    sl_pct = float(st_row.get("stop_loss_pct", 4.0))
    target_p = float(st_row.get("target_price", entry_p * (1 + tp_pct / 100.0)))
    stop_p = float(st_row.get("stop_price", entry_p * (1 - sl_pct / 100.0)))

    ax1.axhline(target_p, color="#00e676", linestyle="--", linewidth=0.9, alpha=0.85)
    ax1.axhline(entry_p, color="#ffffff", linestyle=":", linewidth=1.0, alpha=0.90)
    ax1.axhline(stop_p, color="#ff5252", linestyle="--", linewidth=0.9, alpha=0.85)

    ax1.text(
        n + 0.5, target_p, f" Target (+{tp_pct:.1f}%): {target_p:.2f}",
        color="#00e676", fontsize=8.5, fontweight="bold", va="center",
    )
    ax1.text(
        n + 0.5, entry_p, f" Mua T+0: {entry_p:.2f}",
        color="#ffffff", fontsize=8.5, fontweight="bold", va="center",
    )
    ax1.text(
        n + 0.5, stop_p, f" Cắt lỗ (-{sl_pct:.1f}%): {stop_p:.2f}",
        color="#ff5252", fontsize=8.5, fontweight="bold", va="center",
    )

    ax1.annotate(
        "MUA T+0",
        xy=(n - 1, l[-1]),
        xytext=(max(n - 7, 0), l[-1] - (h.max() - l.min()) * 0.14),
        arrowprops=dict(
            facecolor="#ffd600", edgecolor="#ffd600", shrink=0.05, width=1.5, headwidth=6
        ),
        color="#ffd600",
        fontsize=8.5,
        fontweight="bold",
    )

    r1d = float(st_row.get("r1d", 0.0))
    rvol = float(st_row.get("rvol", 1.0))
    rsi = float(st_row.get("rsi", 50.0))
    val_b = float(st_row.get("val_billion", 0.0))
    bzone = str(st_row.get("buy_zone", ""))
    today_s = get_vn_today().strftime("%Y-%m-%d")

    title_main = (
        f"[{today_s}] {ticker} ({exch} - {sector}) | Giá: {entry_p:.2f}"
        f" ({r1d:+.2f}%) | GTGD: {val_b:.1f} Tỷ | RVOL: {rvol:.2f}x | RSI: {rsi:.0f}"
    )
    ax1.set_title(
        f"{title_main}\n{algo_tag} | {bzone}",
        color="white",
        fontsize=10.0,
        fontweight="bold",
        pad=10,
    )

    y_min = min(l.min(), stop_p) * 0.98
    y_max = max(h.max(), target_p) * 1.02
    ax1.set_ylim(y_min, y_max)
    ax1.set_xlim(-1, n + 13)

    for ax in [ax1, ax2]:
        ax.grid(color="#30363d", linestyle="--", linewidth=0.5, alpha=0.5)
        ax.tick_params(colors="#c9d1d9", labelsize=8)
        for spine in ax.spines.values():
            spine.set_color("#30363d")

    ax1.legend(
        loc="upper left", facecolor="#161b22", edgecolor="#30363d",
        labelcolor="white", fontsize=8,
    )
    ax2.set_ylabel("Vol (Tr)", color="#c9d1d9", fontsize=8)
    plt.setp(ax1.get_xticklabels(), visible=False)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", facecolor="#0d1117")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def compress_png_to_base64_uri(png_bytes):
    """Nén ảnh Đồ thị Nến Matplotlib sang Base64 gọn nhẹ (~18KB) để nhúng thẳng vào ô Google Sheets"""
    try:
        im = PILImage.open(io.BytesIO(png_bytes)).convert("RGB")
        im.thumbnail((460, 260))
        out = io.BytesIO()
        im.save(out, format="JPEG", quality=82, optimize=True)
        b64_str = base64.b64encode(out.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{b64_str}"
    except Exception:
        return ""


def generate_t0_memory_charts(algo_display_results, algo_strict_results):
    console = Console()
    os.makedirs(CHARTS_DIR, exist_ok=True)
    today_str = get_vn_today().strftime("%Y-%m-%d")

    tracked_tickers = {}
    for idx_algo, title_str, _, _, _ in ALGO_METADATA:
        for df_pool in [
            algo_strict_results[idx_algo].head(TOP_N_TO_TRACK_PER_ALGO),
            algo_display_results[idx_algo].head(2),
        ]:
            for _, st in df_pool.iterrows():
                tk = str(st["ticker"]).strip()
                if tk not in tracked_tickers:
                    tracked_tickers[tk] = (
                        st,
                        f"ALGO_{idx_algo:02d} - {title_str[:40]}",
                    )

    chart_bytes_map = {}
    chart_url_map = {}
    chart_b64_map = {}

    with console.status(
        f"[bold magenta]📸 Đang dựng {len(tracked_tickers)} Biểu đồ Nến Nhật T+0 (4 tầng dữ liệu)...",
        spinner="dots",
    ):
        ohlcv_map = {}
        # Tầng 1, 2, 3: DNSE -> VNDirect -> Yahoo Finance (.VN)
        with ThreadPoolExecutor(max_workers=12) as executor:
            future_to_tk = {
                executor.submit(fetch_vn_api_ohlcv, tk, 65): tk
                for tk in tracked_tickers.keys()
            }
            for future in as_completed(future_to_tk):
                tk = future_to_tk[future]
                df_res = future.result()
                if df_res is not None:
                    ohlcv_map[tk] = df_res

        # Tầng 4: TradingView WebSocket cho các mã còn thiếu
        missing_tv = {
            tk: [(tk, tracked_tickers[tk][0]["exchange"])]
            for tk in tracked_tickers.keys()
            if tk not in ohlcv_map
        }
        if missing_tv:
            ohlcv_map.update(fetch_tv_websocket_ohlcv(missing_tv, 65))

        # Vẽ đồ thị nến cho 100% các mã (Nếu vẫn thiếu thì dùng Tầng 5: Tái tạo từ Scanner)
        for tk, (st_row, algo_tag) in tracked_tickers.items():
            df_ohlcv = ohlcv_map.get(tk)
            if df_ohlcv is None or len(df_ohlcv) < 5:
                df_ohlcv = build_synthetic_ohlcv_from_scanner(st_row, 30)

            png_bytes = render_single_t0_candlestick_bytes(
                tk,
                st_row["exchange"],
                st_row["sector"],
                df_ohlcv,
                st_row,
                algo_tag,
            )
            if png_bytes:
                chart_bytes_map[tk] = png_bytes
                chart_b64_map[f"{today_str}_{tk}"] = compress_png_to_base64_uri(png_bytes)
                img_file_path = os.path.join(CHARTS_DIR, f"{today_str}_{tk}.png")
                if not os.path.exists(img_file_path):
                    try:
                        with open(img_file_path, "wb") as f_out:
                            f_out.write(png_bytes)
                    except Exception:
                        pass

            chart_url_map[tk] = build_quickchart_candlestick_url(
                tk, st_row["exchange"], df_ohlcv, st_row
            )

    return chart_bytes_map, chart_url_map, chart_b64_map


def create_discord_grid_bytes(
    start_a, end_a, algo_display_results, chart_bytes_map
):
    pil_imgs = []
    for idx_algo in range(start_a, end_a + 1):
        df_sub = algo_display_results[idx_algo]
        if not df_sub.empty:
            tk = str(df_sub.iloc[0]["ticker"]).strip()
            if tk in chart_bytes_map:
                try:
                    pil_imgs.append(
                        PILImage.open(io.BytesIO(chart_bytes_map[tk]))
                    )
                except Exception:
                    pass

    if not pil_imgs:
        return None

    try:
        w, h = pil_imgs[0].size
        cols = 2 if len(pil_imgs) > 1 else 1
        rows = (len(pil_imgs) + 1) // 2
        grid_img = PILImage.new("RGB", (cols * w, rows * h), (13, 17, 23))
        for i, im in enumerate(pil_imgs[:6]):
            im_resized = im.resize((w, h))
            grid_img.paste(im_resized, ((i % cols) * w, (i // cols) * h))

        out_buf = io.BytesIO()
        grid_img.save(out_buf, format="PNG")
        out_buf.seek(0)
        return out_buf.getvalue()
    except Exception:
        return None


# =========================================================================
# 4. HỆ THỐNG 30 BỘ THUẬT TOÁN ĐỊNH LƯỢNG
# =========================================================================


def calc_roc(series, length):
    if (
        series is None
        or len(series) <= length
        or series.iloc[-1 - length] == 0
    ):
        return 0.0
    return float(
        ((series.iloc[-1] - series.iloc[-1 - length]) / series.iloc[-1 - length])
        * 100
    )


def calc_zscore(series, length=LEN):
    if series is None or len(series) < length or series.iloc[-1] == 0:
        return 0.0
    w = series.iloc[-length:]
    std = w.std()
    return float((series.iloc[-1] - w.mean()) / std) if std != 0 else 0.0


ALGO_METADATA = [
    (1, "🔥 BỘ 01: BÙNG NỔ DÒNG TIỀN & VƯỢT NỀN (CANSLIM BREAKOUT)", "🔥 Nổ Vol Vượt Nền (T+0)", "bold black on green", "🔥"),
    (2, "🛡️ BỘ 02: GOM NỀN TÍCH LŨY SIẾT BIÊN ĐỘ VCP (SEPA ACCUMULATION)", "🛡️ Siết Nền Cạn Cung", "bold white on blue", "🛡️"),
    (3, "⚡ BỘ 03: BẮT ĐÁY RÚT CHÂN & ĐẢO CHIỀU QUÁ BÁN (WYCKOFF SPRING)", "⚡ Rút Chân Bắt Đáy", "bold white on magenta", "⚡"),
    (4, "👑 BỘ 04: SIÊU CỔ PHIẾU KHÁNG BÃO & ĐỈNH 52W (RS ALPHA LEADER)", "👑 Leader Vượt Đỉnh 52W", "bold black on yellow", "👑"),
    (5, "🎯 BỘ 05: ĐIỂM MUA KÉO NGƯỢC HỖ TRỢ NGẮN HẠN (MA20 PULLBACK)", "🎯 Test MA20 Bật Tăng", "bold black on cyan", "🎯"),
    (6, "📈 BỘ 06: ĐỒNG THUẬN ĐỘNG LƯỢNG MACD & XU HƯỚNG ADX (TREND)", "📈 MACD + ADX Đồng Thuận", "bold white on green", "📈"),
    (7, "🍾 BỘ 07: NÉN CHẶT CỔ CHAI & BUNG NÚT BOLLINGER (BB SQUEEZE)", "🍾 Bung Nút Bollinger", "bold white on red", "🍾"),
    (8, "🐳 BỘ 08: DÒNG TIỀN CÁ MẬP GOM NGẦM THANH KHOẢN LỚN (STEALTH)", "🐳 Cá Mập Gom Ngầm", "bold white on blue", "🐳"),
    (9, "🔄 BỘ 09: CHUYỂN PHA DÀI HẠN & VƯỢT MA200 (MA200 TURNAROUND)", "🔄 Vượt MA200 Dài Hạn", "bold black on yellow", "🔄"),
    (10, "🚀 BỘ 10: SÓNG ĐẦU CƠ BETA CAO & DÒNG TIỀN NÓNG (HIGH-BETA)", "🚀 Đầu Cơ Nhạy Sóng", "bold white on magenta", "🚀"),
    (11, "💎 BỘ 11: MUA SỚM TRONG NỀN TÍCH LŨY (GIL MORALES POCKET PIVOT)", "💎 Pocket Pivot Trong Nền", "bold black on cyan", "💎"),
    (12, "🏆 BỘ 12: MẪU HÌNH SIÊU CỔ PHIẾU SEPA (MINERVINI TREND TEMPLATE)", "🏆 Chuẩn 8 Tiêu Chí Minervini", "bold black on green", "🏆"),
    (13, "☄️ BỘ 13: BỨT PHÁ KHOẢNG TRỐNG GIÁ KÈM CẦU LỚN (POWER GAP-UP)", "☄️ Mở Gap Bứt Phá Mạnh", "bold white on red", "☄️"),
    (14, "🟩 BỘ 14: NẾN CƯỜNG LỰC MARUBOZU ÁP ĐẢO (BULLISH MARUBOZU)", "🟩 Nến Marubozu Đóng Trần/Đỉnh", "bold black on green", "🟩"),
    (15, "🌊 BỘ 15: ÁP LỰC MUA CHỦ ĐỘNG TRONG PHIÊN (MONEY FLOW SURGE)", "🌊 Dòng Tiền Chủ Động Áp Đảo", "bold white on blue", "🌊"),
    (16, "✨ BỘ 16: GIAO CẮT VÀNG TRUNG HẠN (GOLDEN CROSS SMA20 x SMA50)", "✨ Giao Cắt Vàng MA20xMA50", "bold black on yellow", "✨"),
    (17, "🌈 BỘ 17: DẢI RUY-BĂNG EMA ĐỒNG THUẬN TĂNG (EMA RIBBON TREND)", "🌈 EMA10 > EMA20 > EMA50", "bold white on green", "🌈"),
    (18, "🔁 BỘ 18: ĐẢO CHIỀU ĐỘNG LƯỢNG STOCHASTIC (STOCHASTIC CROSS)", "🔁 Stoch %K Cắt Lên %D", "bold white on magenta", "🔁"),
    (19, "⚡ BỘ 19: BỨT PHÁ XUNG LỰC ĐỘNG LƯỢNG CCI > 100 (CCI THRUST)", "⚡ Xung Lực CCI Vượt 100", "bold black on cyan", "⚡"),
    (20, "⚓ BỘ 20: KÉO NGƯỢC VỀ HỖ TRỢ TỔ CHỨC MA50 (MA50 PULLBACK)", "⚓ Chạm Hỗ Trợ MA50 Nảy Lên", "bold white on blue", "⚓"),
    (21, "🚩 BỘ 21: HẤP THỤ CUNG GIÁ CAO - NỀN CHỒNG NỀN (HIGH-TIGHT FLAG)", "🚩 Giữ Giá Cao Siết Cờ", "bold black on yellow", "🚩"),
    (22, "🛡️ BỘ 22: LÁ CHẮN PHÂN KỲ NGƯỢC BÃO VN-INDEX (ANTI-INDEX SHIELD)", "🛡️ Xanh Ngược Bão Thị Trường", "bold black on green", "🛡️"),
    (23, "🏦 BỘ 23: BỨT PHÁ GIÁ VỐN TỔ CHỨC VWMA (VWMA BREAKOUT)", "🏦 Vượt Giá Vốn VWMA20", "bold white on blue", "🏦"),
    (24, "💥 BỘ 24: BÙNG NỔ BIÊN ĐỘ BIẾN ĐỘNG ATR (ATR EXPANSION)", "💥 Biên Độ Vượt Chuẩn ATR", "bold white on red", "💥"),
    (25, "🏹 BỘ 25: ĐẢO CHIỀU NHANH T+2.5 QUÁ BÁN (CONNORS RSI-7 DIP)", "🏹 RSI(7) Quá Bán Bật Xanh", "bold white on magenta", "🏹"),
    (26, "🔇 BỘ 26: CẠN KIỆT NGUỒN CUNG TRƯỚC GIỜ NỔ (VSA NO-SUPPLY TEST)", "🔇 Cạn Cung VSA Chờ Nổ", "bold black on cyan", "🔇"),
    (27, "💰 BỘ 27: ĐỊNH GIÁ HỢP LÝ P/E & ROE CAO HÚT TIỀN (GARP VALUE)", "💰 Cơ Bản Tốt (ROE Cao) + Tiền Vào", "bold black on green", "💰"),
    (28, "🏛️ BỘ 28: TRỤ CỘT VỐN HÓA LỚN DẪN DẮT CHỈ SỐ (BLUECHIP PILLAR)", "🏛️ Dòng Tiền Kéo Trụ Lớn", "bold black on yellow", "🏛️"),
    (29, "🧭 BỘ 29: ĐỒNG THUẬN ĐA KHUNG 1D-1W-1M-3M (MTF SYNERGY)", "🧭 Đồng Thuận Tăng 4 Khung", "bold white on green", "🧭"),
    (30, "🌟 BỘ 30: ĐIỂM HỘI TỤ SIÊU ĐẲNG ĐA THUẬT TOÁN (MASTER CONFLUENCE)", "🌟 Hội Tụ Nhiều Thuật Toán Nhất", "bold black on yellow", "🌟"),
]


def run_thirty_independent_algorithms(df):
    clean_stock = ~df["is_bo_cung"]
    val_ok = (df["val_proj_billion"] >= MIN_VALUE_TRADED_BILLION_VND) & clean_stock
    green_bar = df["close_raw"] >= df["open_raw"]
    rvol_cap = df["rvol"].clip(upper=6.0)
    f_bonus = (df["foreign_net_b"].clip(lower=-10.0, upper=20.0)) * 0.25

    df["algo1_pass"] = val_ok & df["above_ma20"] & df["above_ma50"] & green_bar & (df["r1d"] >= 2.2) & (df["rvol"] >= 1.35) & (df["close_pos"] >= 0.70) & df["rsi"].between(54, 78)
    df["algo1_score"] = (rvol_cap * 12.0) + (df["r1d"] * 3.2) + (df["rs_5d"] * 1.5) + (df["close_pos"] * 6.0) + f_bonus

    df["algo2_pass"] = val_ok & df["above_ma20"] & df["above_ma50"] & df["dist_ma20_pct"].between(0.0, 4.5) & (df["rs_5d"] > 0.3) & (df["rs_20d"] > 1.8) & df["r1d"].between(-1.2, 1.8) & df["rvol"].between(0.40, 1.15) & (df["intraday_range_pct"] <= 3.2) & df["rsi"].between(48, 65)
    df["algo2_score"] = (df["rs_20d"] * 1.6) + (df["rs_5d"] * 2.5) - (df["intraday_range_pct"] * 2.5) - ((df["rvol"] - 0.80).abs() * 8.0) + f_bonus

    is_dip = (df["r5d"] <= -3.0) | (df["r20d"] <= -5.5) | (df["rsi"] <= 43.0)
    df["algo3_pass"] = val_ok & is_dip & green_bar & (df["close_pos"] >= 0.65) & (df["r1d"] >= 0.5) & (df["rvol"] >= 0.95)
    df["algo3_score"] = (df["r20d"].clip(upper=0).abs() * 1.3) + ((50.0 - df["rsi"].clip(upper=50.0)) * 1.3) + (df["r1d"] * 3.5) + (rvol_cap * 6.0) + (df["close_pos"] * 8.0) + f_bonus

    df["algo4_pass"] = val_ok & df["above_ma20"] & df["above_ma50"] & df["above_ma200"] & (df["rs_1d"] > 0.0) & (df["rs_5d"] > 2.0) & (df["rs_20d"] > 5.0) & (df["pct_52w_high"] >= 86.0) & df["rsi"].between(55, 78)
    df["algo4_score"] = (df["rs_20d"] * 2.0) + (df["rs_5d"] * 3.0) + (df["pct_52w_high"] * 0.45) + (df["rs_1d"] * 2.0) + f_bonus

    df["algo5_pass"] = val_ok & (df["sma20"] > df["sma50"]) & (df["r20d"] > 3.0) & df["dist_ma20_pct"].between(-1.8, 2.5) & (df["r5d"] <= 1.5) & (df["r1d"] >= 0.3) & (df["close_pos"] >= 0.55) & (df["rvol"] <= 1.20)
    df["algo5_score"] = (df["r20d"] * 1.8) - (df["dist_ma20_pct"].abs() * 5.5) + (df["close_pos"] * 10.0) + (df["r1d"] * 2.5) + f_bonus

    df["algo6_pass"] = val_ok & df["above_ma20"] & (df["macd_hist"] > 0) & (df["macd_val"] > 0) & (df["adx"] >= 22.0) & (df["r1d"] >= 0.8) & (df["rvol"] >= 1.05)
    df["algo6_score"] = (df["adx"] * 1.3) + (df["rs_5d"] * 2.0) + (rvol_cap * 6.0) + (df["r1d"] * 2.2) + f_bonus

    df["algo7_pass"] = val_ok & df["above_ma20"] & (df["close_raw"] >= df["bb_upper"] * 0.98) & (df["bb_width_pct"] <= 19.0) & (df["r1d"] >= 1.8) & (df["rvol"] >= 1.20)
    df["algo7_score"] = ((24.0 - df["bb_width_pct"].clip(upper=24.0)) * 2.0) + (rvol_cap * 8.0) + (df["r1d"] * 3.0) + (df["close_pos"] * 6.0) + f_bonus

    df["algo8_pass"] = clean_stock & (df["val_proj_billion"] >= 18.0) & df["above_ma20"] & (df["rvol"] >= 1.30) & df["r1d"].between(0.2, 3.2) & (df["close_pos"] >= 0.62) & df["rsi"].between(48, 69)
    df["algo8_score"] = (np.log1p(df["val_proj_billion"]) * 6.5) + (rvol_cap * 10.0) + (df["close_pos"] * 8.5) + (df["rs_5d"] * 1.5) + f_bonus

    df["algo9_pass"] = val_ok & df["above_ma200"] & df["above_ma20"] & df["dist_ma200_pct"].between(0.0, 7.5) & (df["r5d"] >= 1.2) & (df["rvol"] >= 1.05)
    df["algo9_score"] = ((9.0 - (df["dist_ma200_pct"] - 2.5).abs().clip(upper=9.0)) * 2.8) + (df["r5d"] * 2.5) + (rvol_cap * 7.0) + (df["r1d"] * 2.0) + f_bonus

    df["algo10_pass"] = clean_stock & df["val_proj_billion"].between(2.0, 80.0) & df["above_ma20"] & (df["r1d"] >= 2.5) & (df["r5d"] >= 3.5) & (df["rvol"] >= 1.35) & (df["close_pos"] >= 0.65)
    df["algo10_score"] = (df["r1d"] * 3.5) + (df["r5d"] * 2.5) + (rvol_cap * 8.0) + (df["rs_20d"] * 1.0)

    df["algo11_pass"] = val_ok & df["above_ma20"] & df["above_ma50"] & df["dist_ma20_pct"].between(-0.5, 4.2) & (df["r20d"] <= 8.5) & df["r1d"].between(1.0, 4.0) & (df["rvol"] >= 1.35) & (df["close_pos"] >= 0.65)
    df["algo11_score"] = (rvol_cap * 12.0) + (df["close_pos"] * 8.0) + (df["r1d"] * 3.0) - (df["dist_ma20_pct"].abs() * 2.0) + f_bonus

    df["algo12_pass"] = val_ok & (df["close_raw"] > df["ema10"]) & (df["ema10"] > df["sma20"]) & (df["sma20"] > df["sma50"]) & (df["sma50"] > df["sma200"]) & (df["pct_52w_high"] >= 85.0) & (df["gain_52w_low"] >= 25.0) & (df["rs_20d"] > 3.5)
    df["algo12_score"] = (df["pct_52w_high"] * 0.5) + (df["gain_52w_low"] * 0.25) + (df["rs_20d"] * 2.2) + (df["rs_5d"] * 2.0) + f_bonus

    df["algo13_pass"] = val_ok & (df["gap_up_pct"] >= 0.5) & green_bar & (df["r1d"] >= 2.0) & (df["rvol"] >= 1.30) & (df["close_pos"] >= 0.68)
    df["algo13_score"] = (df["gap_up_pct"] * 6.0) + (rvol_cap * 10.0) + (df["r1d"] * 3.0) + (df["close_pos"] * 6.0) + f_bonus

    df["algo14_pass"] = val_ok & (df["body_ratio"] >= 0.75) & (df["close_pos"] >= 0.85) & (df["r1d"] >= 2.5) & (df["rvol"] >= 1.15)
    df["algo14_score"] = (df["body_ratio"] * 25.0) + (df["r1d"] * 4.0) + (rvol_cap * 7.0) + f_bonus

    df["algo15_pass"] = val_ok & (df["mf_mult"] >= 0.55) & (df["rvol"] >= 1.25) & (df["close_raw"] > df["vwma"]) & (df["r1d"] > 0.5)
    df["algo15_score"] = (df["mf_mult"] * 20.0) + (rvol_cap * 9.0) + (np.log1p(df["val_proj_billion"]) * 4.0) + (df["rs_5d"] * 1.5) + f_bonus

    df["algo16_pass"] = val_ok & df["above_ma20"] & df["sma20_50_diff_pct"].between(0.0, 3.5) & (df["r5d"] > 0.5) & (df["r1d"] > 0.0) & (df["rvol"] >= 1.0)
    df["algo16_score"] = ((4.0 - df["sma20_50_diff_pct"].clip(lower=0, upper=4.0)) * 5.0) + (df["r5d"] * 3.0) + (rvol_cap * 6.0) + (df["r1d"] * 2.0) + f_bonus

    df["algo17_pass"] = val_ok & (df["close_raw"] > df["ema10"]) & (df["ema10"] > df["ema20"]) & (df["ema20"] > df["ema50"]) & (df["ema50"] > df["sma200"]) & (df["adx"] >= 20.0) & (df["r1d"] > 0.3)
    df["algo17_score"] = (df["adx"] * 1.2) + (df["r20d"] * 1.8) + (df["r5d"] * 2.2) + (df["r1d"] * 2.0) + f_bonus

    df["algo18_pass"] = val_ok & (df["stoch_k"] > df["stoch_d"]) & ((df["stoch_k"] - df["stoch_d"]) <= 15.0) & df["stoch_k"].between(18.0, 58.0) & (df["r1d"] >= 0.5) & df["above_ma50"]
    df["algo18_score"] = ((60.0 - df["stoch_k"].clip(upper=60.0)) * 0.6) + ((df["stoch_k"] - df["stoch_d"]) * 2.5) + (df["r1d"] * 3.5) + (rvol_cap * 5.0) + f_bonus

    df["algo19_pass"] = val_ok & df["cci20"].between(100.0, 230.0) & df["above_ma20"] & (df["r1d"] >= 1.2) & (df["rvol"] >= 1.15)
    df["algo19_score"] = (df["cci20"] * 0.18) + (rvol_cap * 8.0) + (df["r1d"] * 3.0) + (df["rs_5d"] * 1.5) + f_bonus

    df["algo20_pass"] = val_ok & (df["sma50"] > df["sma200"]) & (df["r60d"] > 4.0) & df["dist_ma50_pct"].between(-2.0, 3.0) & (df["r1d"] >= 0.5) & (df["close_pos"] >= 0.55)
    df["algo20_score"] = (df["r60d"] * 1.2) - (df["dist_ma50_pct"].abs() * 5.0) + (df["close_pos"] * 10.0) + (df["r1d"] * 3.0) + f_bonus

    df["algo21_pass"] = val_ok & ((df["r20d"] >= 8.0) | (df["r60d"] >= 14.0)) & df["r5d"].between(-2.2, 3.2) & (df["pct_52w_high"] >= 85.0) & df["above_ma20"] & (df["rvol"] <= 1.15)
    df["algo21_score"] = (df["r20d"] * 1.8) + (df["pct_52w_high"] * 0.4) - (df["r5d"].abs() * 2.0) - (df["rvol"] * 5.0) + f_bonus

    df["algo22_pass"] = clean_stock & (df["val_proj_billion"] >= 8.0) & (df["rs_1d"] >= 1.2) & (df["rs_5d"] >= 2.5) & (df["r1d"] > 0.3) & (df["close_pos"] >= 0.60) & df["above_ma20"]
    df["algo22_score"] = (df["rs_1d"] * 5.0) + (df["rs_5d"] * 3.5) + (df["rs_20d"] * 1.5) + (df["close_pos"] * 6.0) + f_bonus

    df["algo23_pass"] = val_ok & df["dist_vwma_pct"].between(0.1, 4.5) & df["above_ma20"] & (df["rvol"] >= 1.20) & (df["r1d"] >= 1.0) & (df["close_pos"] >= 0.60)
    df["algo23_score"] = (rvol_cap * 10.0) + (df["r1d"] * 3.0) + ((5.0 - df["dist_vwma_pct"].clip(lower=0, upper=5.0)) * 3.0) + (df["rs_5d"] * 1.5) + f_bonus

    df["algo24_pass"] = val_ok & (df["atr_break_ratio"] >= 1.15) & df["above_ma20"] & (df["close_pos"] >= 0.70) & (df["rvol"] >= 1.20)
    df["algo24_score"] = (df["atr_break_ratio"] * 12.0) + (rvol_cap * 8.0) + (df["close_pos"] * 8.0) + f_bonus

    df["algo25_pass"] = val_ok & df["above_ma200"] & df["rsi7"].between(18.0, 46.0) & (df["r1d"] >= 0.5) & green_bar & (df["close_pos"] >= 0.58)
    df["algo25_score"] = ((50.0 - df["rsi7"].clip(upper=50.0)) * 1.5) + (df["r1d"] * 4.0) + (df["close_pos"] * 8.0) + (rvol_cap * 4.0) + f_bonus

    df["algo26_pass"] = val_ok & df["above_ma20"] & df["above_ma50"] & (df["rvol"] <= 0.68) & df["r1d"].between(-0.8, 1.5) & (df["close_pos"] >= 0.50) & (df["intraday_range_pct"] <= 2.8)
    df["algo26_score"] = ((1.0 - df["rvol"].clip(upper=1.0)) * 25.0) + (df["rs_20d"] * 1.5) + (df["close_pos"] * 8.0) - (df["intraday_range_pct"] * 2.0) + f_bonus

    df["algo27_pass"] = val_ok & df["pe"].between(3.0, 18.5) & (df["roe"] >= 12.0) & df["above_ma20"] & (df["r5d"] > 0.0) & (df["rvol"] >= 1.0)
    df["algo27_score"] = (df["roe"] * 1.2) + ((20.0 - df["pe"].clip(lower=3.0, upper=20.0)) * 1.2) + (df["rs_20d"] * 1.5) + (rvol_cap * 6.0) + f_bonus

    is_pillar = clean_stock & ((df["mcap_billion"] >= 12000.0) | (df["val_proj_billion"] >= 50.0))
    df["algo28_pass"] = is_pillar & df["above_ma20"] & (df["r1d"] > 0.2) & (df["rs_5d"] > -0.5) & (df["close_pos"] >= 0.55)
    df["algo28_score"] = (np.log1p(df["mcap_billion"]) * 3.0) + (df["rs_5d"] * 3.0) + (df["r1d"] * 3.5) + (rvol_cap * 8.0) + f_bonus

    df["algo29_pass"] = val_ok & (df["r1d"] > 0.5) & (df["r5d"] > 1.5) & (df["r20d"] > 3.5) & (df["r60d"] > 6.0) & df["above_ma20"] & df["rsi"].between(54, 76)
    df["algo29_score"] = (df["r1d"] * 2.5) + (df["r5d"] * 2.0) + (df["r20d"] * 1.2) + (df["r60d"] * 0.6) + (rvol_cap * 5.0) + f_bonus

    pass_cols_29 = [f"algo{i}_pass" for i in range(1, 30)]
    df["confluence_count"] = df[pass_cols_29].sum(axis=1)
    df["algo30_pass"] = val_ok & (df["confluence_count"] >= 4) & (df["r1d"] > 0.0)
    df["algo30_score"] = (df["confluence_count"] * 10.0) + (df["rs_5d"] * 2.0) + (df["r1d"] * 2.5) + (rvol_cap * 5.0) + f_bonus

    algo_strict_results = {}
    algo_display_results = {}

    for idx_algo in range(1, 31):
        pass_col = f"algo{idx_algo}_pass"
        score_col = f"algo{idx_algo}_score"

        df_strict = (
            df[df[pass_col] & (~df["chase_risk"])]
            .sort_values(by=score_col, ascending=False)
            .copy()
        )
        algo_strict_results[idx_algo] = df_strict

        df_disp = df[df[pass_col]].sort_values(by=score_col, ascending=False).copy()
        if len(df_disp) < 5:
            min_v = 2.0 if idx_algo == 10 else (12.0 if idx_algo == 28 else MIN_VALUE_TRADED_BILLION_VND)
            pool = df[(df["val_proj_billion"] >= min_v) & clean_stock]
            if pool.empty:
                pool = df
            df_disp = pool.sort_values(by=score_col, ascending=False).head(6).copy()
        algo_display_results[idx_algo] = df_disp

    return algo_strict_results, algo_display_results


# =========================================================================
# 5. ĐỒNG BỘ KÉP: GOOGLE SHEETS (NHÚNG ẢNH BIỂU ĐỒ NẾN) + FILE EXCEL (.XLSX)
# =========================================================================


def sync_google_sheets_and_local_excel(
    df_all,
    algo_strict_results,
    df_export_full,
    vni_price_now=0.0,
    chart_url_map=None,
    chart_b64_map=None,
):
    console = Console()
    if chart_url_map is None:
        chart_url_map = {}
    if chart_b64_map is None:
        chart_b64_map = {}

    today_date = get_vn_today()
    today_str = today_date.strftime("%Y-%m-%d")
    price_map = dict(zip(df_all["ticker"], df_all["price"]))

    log_cols = [
        "Ngày Báo",              # A (1)
        "Mã Bộ",                 # B (2)
        "Tên Bộ Thuật Toán",     # C (3)
        "Hạng",                  # D (4)
        "Mã CP",                 # E (5)
        "Sàn",                   # F (6)
        "Nhóm Ngành",            # G (7)
        "Giá Mua (T+0)",         # H (8)
        "VNINDEX (T+0)",         # I (9)
        "Điểm Algo",             # J (10)
        "RVOL (T+0)",            # K (11)
        "Mục Tiêu Động (ATR)",   # L (12)
        "Cắt Lỗ Động (ATR)",     # M (13)
        "Giá Hiện Tại",          # N (14)
        "VNINDEX Hiện Tại",      # O (15)
        "Số Phiên GD (T+)",      # P (16)
        "Đỉnh Cao Nhất (MFE)",   # Q (17)
        "Đáy Thấp Nhất (MAE)",   # R (18)
        "Giá T+3 (Hàng Về)",     # S (19)
        "Giá T+5 (1 Tuần)",      # T (20)
        "Giá T+10 (2 Tuần)",     # U (21)
        "Giá T+20 (1 Tháng)",    # V (22)
        "Giá Chốt Lệnh (Khóa)",  # W (23)
        "Trạng Thái Lệnh",       # X (24)
    ]

    existing_records = []
    if GOOGLE_SHEETS_WEBHOOK_URL:
        try:
            r_get = requests.get(
                GOOGLE_SHEETS_WEBHOOK_URL,
                params={"action": "read_log"},
                timeout=12,
            )
            if r_get.status_code == 200:
                j_get = r_get.json()
                for row in j_get.get("records", []):
                    if not row.get("Mã CP") or not row.get("Mã Bộ"):
                        continue
                    rec_item = {col: row.get(col, "") for col in log_cols}
                    rec_item["chart_url"] = row.get("chart_url", "")
                    existing_records.append(rec_item)
        except Exception as e:
            console.print(
                f"[bold yellow]⚠️ Đọc lịch sử Google Sheets: {e}[/bold yellow]"
            )

    if not existing_records and os.path.exists(EXCEL_TRACKER_FILENAME):
        try:
            df_old = pd.read_excel(
                EXCEL_TRACKER_FILENAME,
                sheet_name="2_Nhat_Ky_Theo_Doi_Tuong_Lai",
            )
            for _, row in df_old.iterrows():
                if pd.isna(row.get("Mã CP")) or pd.isna(row.get("Mã Bộ")):
                    continue
                rec_item = {
                    col: ("" if pd.isna(row.get(col)) else row.get(col))
                    for col in log_cols
                }
                rec_item["chart_url"] = (
                    ""
                    if pd.isna(row.get("Link Ảnh Nến Cloud (T+0)"))
                    else str(row.get("Link Ảnh Nến Cloud (T+0)"))
                )
                existing_records.append(rec_item)
        except Exception:
            pass

    def to_num(val):
        try:
            if val is None or val == "" or pd.isna(val):
                return 0.0
            return float(val)
        except Exception:
            return 0.0

    open_tickers_to_check = set()
    for rec in existing_records:
        sig_d = str(rec["Ngày Báo"])[:10]
        tk = str(rec["Mã CP"]).strip()
        is_locked = to_num(rec["Giá Chốt Lệnh (Khóa)"]) > 0
        if sig_d != today_str and not is_locked and tk:
            open_tickers_to_check.add(tk)

    hist_ohlcv_cache = {}
    if open_tickers_to_check:
        with ThreadPoolExecutor(max_workers=10) as executor:
            f_map = {
                executor.submit(fetch_vn_api_ohlcv, tk, 35): tk
                for tk in open_tickers_to_check
            }
            for fut in as_completed(f_map):
                tk = f_map[fut]
                res_df = fut.result()
                if res_df is not None and not res_df.empty:
                    hist_ohlcv_cache[tk] = res_df

    logged_keys = set()
    for rec in existing_records:
        sig_date_str = str(rec["Ngày Báo"])[:10]
        rec["Ngày Báo"] = sig_date_str
        algo_code = str(rec["Mã Bộ"]).strip()
        tk = str(rec["Mã CP"]).strip()
        logged_keys.add((sig_date_str, algo_code, tk))

        try:
            d0 = datetime.datetime.strptime(sig_date_str, "%Y-%m-%d").date()
            trading_days = (
                int(np.busday_count(d0, today_date)) if today_date > d0 else 0
            )
        except Exception:
            trading_days = 0
        rec["Số Phiên GD (T+)"] = trading_days

        if vni_price_now > 0:
            rec["VNINDEX Hiện Tại"] = round(float(vni_price_now), 2)
            if to_num(rec["VNINDEX (T+0)"]) <= 0:
                rec["VNINDEX (T+0)"] = round(float(vni_price_now), 2)

        if tk in price_map and price_map[tk] > 0:
            curr_p = round(float(price_map[tk]), 2)
            entry_p = to_num(rec["Giá Mua (T+0)"]) or curr_p
            target_p = to_num(rec["Mục Tiêu Động (ATR)"]) or round(entry_p * 1.07, 2)
            stop_p = to_num(rec["Cắt Lỗ Động (ATR)"]) or round(entry_p * 0.96, 2)

            rec["Giá Hiện Tại"] = curr_p
            old_mfe = to_num(rec["Đỉnh Cao Nhất (MFE)"]) or entry_p
            old_mae = to_num(rec["Đáy Thấp Nhất (MAE)"]) or entry_p

            df_h = hist_ohlcv_cache.get(tk)
            hit_tp_in_hist = False
            hit_sl_in_hist = False

            if df_h is not None and trading_days >= 1 and len(df_h) > trading_days:
                scale = 1000.0 if df_h["close"].iloc[-1] > 500 else 1.0
                bars_since_entry = df_h.iloc[-trading_days:]
                max_h = float(bars_since_entry["high"].max()) / scale
                min_l = float(bars_since_entry["low"].min()) / scale
                old_mfe = max(old_mfe, max_h)
                old_mae = min(old_mae, min_l)

                for t_step, col_t in [
                    (3, "Giá T+3 (Hàng Về)"),
                    (5, "Giá T+5 (1 Tuần)"),
                    (10, "Giá T+10 (2 Tuần)"),
                    (20, "Giá T+20 (1 Tháng)"),
                ]:
                    if trading_days >= t_step and to_num(rec[col_t]) <= 0:
                        idx_bar = -trading_days + (t_step - 1)
                        if abs(idx_bar) <= len(df_h):
                            rec[col_t] = round(
                                float(df_h["close"].iloc[idx_bar]) / scale, 2
                            )

                if trading_days >= 3:
                    bars_after_t3 = df_h.iloc[-(trading_days - 2) :]
                    for _, b_row in bars_after_t3.iterrows():
                        bh = float(b_row["high"]) / scale
                        bl = float(b_row["low"]) / scale
                        if bh >= target_p:
                            hit_tp_in_hist = True
                            break
                        if bl <= stop_p:
                            hit_sl_in_hist = True
                            break

            rec["Đỉnh Cao Nhất (MFE)"] = round(max(old_mfe, curr_p), 2)
            rec["Đáy Thấp Nhất (MAE)"] = round(min(old_mae, curr_p), 2)

            if trading_days >= 3 and to_num(rec["Giá T+3 (Hàng Về)"]) <= 0:
                rec["Giá T+3 (Hàng Về)"] = curr_p
            if trading_days >= 5 and to_num(rec["Giá T+5 (1 Tuần)"]) <= 0:
                rec["Giá T+5 (1 Tuần)"] = curr_p
            if trading_days >= 10 and to_num(rec["Giá T+10 (2 Tuần)"]) <= 0:
                rec["Giá T+10 (2 Tuần)"] = curr_p
            if trading_days >= 20 and to_num(rec["Giá T+20 (1 Tháng)"]) <= 0:
                rec["Giá T+20 (1 Tháng)"] = curr_p

            if to_num(rec["Giá Chốt Lệnh (Khóa)"]) <= 0:
                if trading_days < 3:
                    rec["Trạng Thái Lệnh"] = f"⏳ Chờ hàng về (T+{trading_days})"
                elif hit_tp_in_hist or curr_p >= target_p:
                    rec["Giá Chốt Lệnh (Khóa)"] = (
                        target_p if hit_tp_in_hist else curr_p
                    )
                    rec["Trạng Thái Lệnh"] = "🎯 ĐÃ CHỐT LỜI (ATR)"
                elif hit_sl_in_hist or curr_p <= stop_p:
                    rec["Giá Chốt Lệnh (Khóa)"] = (
                        stop_p if hit_sl_in_hist else curr_p
                    )
                    rec["Trạng Thái Lệnh"] = "🛑 ĐÃ CẮT LỖ (ATR)"
                elif trading_days >= 20:
                    rec["Giá Chốt Lệnh (Khóa)"] = curr_p
                    rec["Trạng Thái Lệnh"] = "🏁 ĐÓNG LỆNH T+20"
                else:
                    rec["Trạng Thái Lệnh"] = f"📈 Đang nắm giữ (T+{trading_days})"

    new_signals_added = 0
    for idx_algo, title_str, _, _, _ in ALGO_METADATA:
        algo_code = f"ALGO_{idx_algo:02d}"
        df_sub_strict = algo_strict_results[idx_algo].head(TOP_N_TO_TRACK_PER_ALGO)
        score_col = f"algo{idx_algo}_score"

        for rank_i, (_, st) in enumerate(df_sub_strict.iterrows(), 1):
            tk = str(st["ticker"]).strip()
            key_tuple = (today_str, algo_code, tk)
            if key_tuple in logged_keys:
                continue

            entry_p = round(float(st["price"]), 2)
            vni_val = (
                round(float(vni_price_now), 2) if vni_price_now > 0 else 1250.0
            )
            existing_records.append({
                "Ngày Báo": today_str,
                "Mã Bộ": algo_code,
                "Tên Bộ Thuật Toán": title_str,
                "Hạng": f"#{rank_i}",
                "Mã CP": tk,
                "Sàn": st["exchange"],
                "Nhóm Ngành": st["sector"],
                "Giá Mua (T+0)": entry_p,
                "VNINDEX (T+0)": vni_val,
                "Điểm Algo": round(float(st[score_col]), 2),
                "RVOL (T+0)": round(float(st["rvol"]), 2),
                "Mục Tiêu Động (ATR)": float(st["target_price"]),
                "Cắt Lỗ Động (ATR)": float(st["stop_price"]),
                "Giá Hiện Tại": entry_p,
                "VNINDEX Hiện Tại": vni_val,
                "Số Phiên GD (T+)": 0,
                "Đỉnh Cao Nhất (MFE)": entry_p,
                "Đáy Thấp Nhất (MAE)": entry_p,
                "Giá T+3 (Hàng Về)": "",
                "Giá T+5 (1 Tuần)": "",
                "Giá T+10 (2 Tuần)": "",
                "Giá T+20 (1 Tháng)": "",
                "Giá Chốt Lệnh (Khóa)": "",
                "Trạng Thái Lệnh": "⏳ Mới mua (T+0)",
                "chart_url": chart_url_map.get(tk, ""),
            })
            logged_keys.add(key_tuple)
            new_signals_added += 1

    sheet2_headers = log_cols + [
        "% Lãi/Lỗ T+3 (Ròng -0.4% phí)",   # Y (25)
        "% Lãi/Lỗ T+5 (Ròng -0.4% phí)",   # Z (26)
        "% Lãi/Lỗ Thực Tế (Sau Phí)",      # AA (27)
        "% Alpha vs VNINDEX",              # AB (28)
        "% Đỉnh MFE (Lãi Tối Đa)",         # AC (29)
        "% Đáy MAE (Sụt Giảm Tối Đa)",     # AD (30)
        "Kết Quả T+3 (Hàng Về)",           # AE (31)
        "Kết Quả Cuối Cùng",               # AF (32)
        "Link Ảnh Nến Cloud (T+0)",        # AG (33)
        "Ảnh Đồ Thị Nến Lúc Lọc (T+0)",   # AH (34)
    ]
    sheet2_rows = [sheet2_headers]

    for row_idx, rec in enumerate(existing_records, start=2):
        row_vals = []
        for col_name in log_cols:
            v = rec.get(col_name, "")
            if v is None or (isinstance(v, float) and np.isnan(v)):
                v = ""
            row_vals.append(v)

        tk_code = str(rec.get("Mã CP", "")).strip()
        sig_d = str(rec.get("Ngày Báo", ""))[:10]
        c_url = str(rec.get("chart_url", "") or "")
        if (not c_url or "type" in c_url) and tk_code in chart_url_map:
            c_url = chart_url_map[tk_code]

        f_y = f'=IF(ISNUMBER(S{row_idx}), ((S{row_idx}-H{row_idx})/H{row_idx})-0.004, "")'
        f_z = f'=IF(ISNUMBER(T{row_idx}), ((T{row_idx}-H{row_idx})/H{row_idx})-0.004, "")'
        f_aa = f"=IF(ISNUMBER(W{row_idx}), ((W{row_idx}-H{row_idx})/H{row_idx})-0.004, IF(P{row_idx}>0, ((N{row_idx}-H{row_idx})/H{row_idx})-0.004, 0))"
        f_ab = f"=IF(AND(ISNUMBER(I{row_idx}), I{row_idx}>0, ISNUMBER(O{row_idx}), O{row_idx}>0), AA{row_idx}-((O{row_idx}-I{row_idx})/I{row_idx}), AA{row_idx})"
        f_ac = f"=IF(ISNUMBER(Q{row_idx}), (Q{row_idx}-H{row_idx})/H{row_idx}, 0)"
        f_ad = f"=IF(ISNUMBER(R{row_idx}), (R{row_idx}-H{row_idx})/H{row_idx}, 0)"
        f_ae = f'=IF(ISNUMBER(Y{row_idx}), IF(Y{row_idx}>0, "THẮNG", "THUA"), "CHỜ T+3")'
        f_af = f'=IF(AND(P{row_idx}=0, N{row_idx}=H{row_idx}, NOT(ISNUMBER(W{row_idx}))), "ĐANG CHỜ", IF(AA{row_idx}>0, "THẮNG", "THUA"))'
        f_ah = f'=IF(AG{row_idx}<>"", IMAGE(AG{row_idx}, 1), "Chưa có ảnh")'

        row_vals.extend([f_y, f_z, f_aa, f_ab, f_ac, f_ad, f_ae, f_af, c_url, f_ah])
        sheet2_rows.append(row_vals)

    max_log_row = max(len(existing_records) + 50, 1000)
    s2_ref = "'2_Nhat_Ky_Theo_Doi_Tuong_Lai'"

    sheet1_headers = [
        "Mã Bộ", "Tên Bộ Thuật Toán", "Tổng Tín Hiệu Chuẩn", "Đã Phân Định",
        "Lệnh THẮNG", "Lệnh THUA", "Win Rate T+3 (Hàng Về)",
        "Win Rate Tổng (%)", "Lãi TB Lệnh Thắng (Avg Win)",
        "Lỗ TB Lệnh Thua (Avg Loss)", "Tỷ Lệ R:R (Payoff Ratio)",
        "Profit Factor (Tổng Lãi/Lỗ)", "Kỳ Vọng Toán Học (Expectancy)",
        "Alpha TB vs VNINDEX", "Đánh Giá Chuẩn Quỹ Định Lượng",
    ]
    sheet1_rows = [sheet1_headers]

    for idx_algo, title_str, _, _, _ in ALGO_METADATA:
        r_i = idx_algo + 1
        algo_code = f"ALGO_{idx_algo:02d}"
        sheet1_rows.append([
            algo_code,
            title_str,
            f"=COUNTIF({s2_ref}!$B$2:$B${max_log_row}, A{r_i})",
            f"=E{r_i}+F{r_i}",
            f'=COUNTIFS({s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THẮNG")',
            f'=COUNTIFS({s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THUA")',
            f'=IFERROR(COUNTIFS({s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AE$2:$AE${max_log_row}, "THẮNG")/COUNTIFS({s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AE$2:$AE${max_log_row}, "<>CHỜ T+3"), 0)',
            f"=IF(D{r_i}>0, E{r_i}/D{r_i}, 0)",
            f'=IFERROR(AVERAGEIFS({s2_ref}!$AA$2:$AA${max_log_row}, {s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THẮNG"), 0)',
            f'=IFERROR(AVERAGEIFS({s2_ref}!$AA$2:$AA${max_log_row}, {s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THUA"), 0)',
            f"=IF(AND(F{r_i}>0, ABS(J{r_i})>0), I{r_i}/ABS(J{r_i}), IF(E{r_i}>0, 5.0, 0))",
            f'=IF(ABS(SUMIFS({s2_ref}!$AA$2:$AA${max_log_row}, {s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THUA"))>0, SUMIFS({s2_ref}!$AA$2:$AA${max_log_row}, {s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THẮNG")/ABS(SUMIFS({s2_ref}!$AA$2:$AA${max_log_row}, {s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "THUA")), IF(E{r_i}>0, 5.0, 0))',
            f"=IF(D{r_i}>0, (H{r_i}*I{r_i})+((1-H{r_i})*J{r_i}), 0)",
            f'=IFERROR(AVERAGEIFS({s2_ref}!$AB$2:$AB${max_log_row}, {s2_ref}!$B$2:$B${max_log_row}, A{r_i}, {s2_ref}!$AF$2:$AF${max_log_row}, "<>ĐANG CHỜ"), 0)',
            f'=IF(D{r_i}=0, "⏳ Đang chờ T+1 -> T+3", IF(AND(M{r_i}>=0.015, L{r_i}>=1.5), "🏆 CHÉN THÁNH (Giải ngân Lớn)", IF(M{r_i}>0, "🟢 KỲ VỌNG DƯƠNG (Có Lợi Thế)", "🔴 KỲ VỌNG ÂM (Tạm Ngưng)")))',
        ])

    sheet3_headers = [
        "STT", "Tên Nhóm Ngành Chứng Khoán Việt Nam", "Tổng Tín Hiệu Ngành",
        "Số Lệnh Đã Phân Định", "Số Lệnh THẮNG", "Số Lệnh THUA",
        "Xác Suất Thắng Ngành (Win Rate %)", "Lãi/Lỗ Ròng Trung Bình (%)",
        "Alpha TB vs VNINDEX (%)", "Trạng Thái Dòng Tiền Ngành",
    ]
    sheet3_rows = [sheet3_headers]
    unique_sectors = sorted(list(df_all["sector"].unique()))
    for s_idx, sec_name in enumerate(unique_sectors, start=1):
        sr = s_idx + 1
        sheet3_rows.append([
            s_idx,
            sec_name,
            f"=COUNTIF({s2_ref}!$G$2:$G${max_log_row}, B{sr})",
            f"=E{sr}+F{sr}",
            f'=COUNTIFS({s2_ref}!$G$2:$G${max_log_row}, B{sr}, {s2_ref}!$AF$2:$AF${max_log_row}, "THẮNG")',
            f'=COUNTIFS({s2_ref}!$G$2:$G${max_log_row}, B{sr}, {s2_ref}!$AF$2:$AF${max_log_row}, "THUA")',
            f"=IF(D{sr}>0, E{sr}/D{sr}, 0)",
            f'=IFERROR(AVERAGEIFS({s2_ref}!$AA$2:$AA${max_log_row}, {s2_ref}!$G$2:$G${max_log_row}, B{sr}, {s2_ref}!$AF$2:$AF${max_log_row}, "<>ĐANG CHỜ"), 0)',
            f'=IFERROR(AVERAGEIFS({s2_ref}!$AB$2:$AB${max_log_row}, {s2_ref}!$G$2:$G${max_log_row}, B{sr}, {s2_ref}!$AF$2:$AF${max_log_row}, "<>ĐANG CHỜ"), 0)',
            f'=IF(D{sr}=0, "⏳ Đang theo dõi", IF(AND(G{sr}>=0.60, H{sr}>0.01), "🔥 NGÀNH DẪN SÓNG (Tỷ lệ ăn cao)", IF(H{sr}>0, "🟢 Có sinh lời", "🩸 Dòng tiền yếu / Hay xịt")))',
        ])

    df_s4 = df_export_full.head(300).fillna("")
    sheet4_rows = [list(df_s4.columns)] + df_s4.values.tolist()

    # Lưu bản sao File Excel (.xlsx) vào kho GitHub
    try:
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "1_Bang_Xac_Suat_Thang"
        ws2 = wb.create_sheet(title="2_Nhat_Ky_Theo_Doi_Tuong_Lai")
        ws3 = wb.create_sheet(title="3_WinRate_Theo_Nhom_Nganh")
        ws4 = wb.create_sheet(title="4_Du_Lieu_Toan_San_Hom_Nay")

        for r_data in clean_for_json(sheet1_rows):
            ws1.append(r_data)
        for r_idx, r_data in enumerate(clean_for_json(sheet2_rows), start=1):
            ws2.append(r_data)
            if r_idx >= 2 and EMBED_THUMBNAIL_IN_EXCEL:
                sig_d = str(r_data[0])[:10]
                tk_c = str(r_data[4]).strip()
                local_img = os.path.join(CHARTS_DIR, f"{sig_d}_{tk_c}.png")
                if os.path.exists(local_img):
                    try:
                        ws2.row_dimensions[r_idx].height = 95
                        xl_img = XLImage(local_img)
                        xl_img.width = 210
                        xl_img.height = 115
                        ws2.add_image(xl_img, f"AH{r_idx}")
                    except Exception:
                        pass
        for r_data in clean_for_json(sheet3_rows):
            ws3.append(r_data)
        for r_data in clean_for_json(sheet4_rows):
            ws4.append(r_data)
        wb.save(EXCEL_TRACKER_FILENAME)
    except Exception:
        pass

    # Đẩy toàn bộ 4 Sheet + Bản đồ Ảnh Biểu đồ Nến lên Google Sheets
    gsheet_status = "Chưa cấu hình GOOGLE_SHEETS_WEBHOOK_URL"
    if GOOGLE_SHEETS_WEBHOOK_URL:
        try:
            payload_gs = clean_for_json({
                "sheet1_rows": sheet1_rows,
                "sheet2_rows": sheet2_rows,
                "sheet3_rows": sheet3_rows,
                "sheet4_rows": sheet4_rows,
                "chart_b64_map": chart_b64_map,
            })
            r_post = requests.post(
                GOOGLE_SHEETS_WEBHOOK_URL,
                data=json.dumps(payload_gs),
                headers={"Content-Type": "application/json"},
                allow_redirects=True,
                timeout=40,
            )
            try:
                res_json = r_post.json()
                if res_json.get("status") == "ok":
                    gsheet_status = f"✅ Đã đồng bộ Google Sheets + Ảnh Biểu đồ Nến ({res_json.get('updated_rows', 0)} lệnh)!"
                    console.print(f"[bold green]{gsheet_status}[/bold green]")
                else:
                    gsheet_status = f"❌ Lỗi từ Apps Script: {res_json.get('message')}"
                    console.print(f"[bold red]{gsheet_status}[/bold red]")
            except Exception:
                gsheet_status = f"⚠️ Link Google Sheets bị chặn quyền (Hãy chọn 'Anyone/Bất kỳ ai' khi Deploy): {r_post.text[:120]}"
                console.print(f"[bold yellow]{gsheet_status}[/bold yellow]")
        except Exception as e:
            gsheet_status = f"❌ Lỗi kết nối Google Sheets: {e}"
            console.print(f"[bold red]{gsheet_status}[/bold red]")

    return gsheet_status, len(existing_records), new_signals_added


# =========================================================================
# 6. BỘ CHUYỂN MẠCH CHẾ ĐỘ THỊ TRƯỜNG & PHÂN TÍCH TỔNG HỢP
# =========================================================================


def detect_market_regime(
    vni_1d, vni_5d, pct_above_ma20, advancers, decliners, z_usdvnd
):
    if (
        z_usdvnd > 1.3
        or pct_above_ma20 < 28.0
        or (vni_1d < -1.5 and decliners > advancers * 2.0)
    ):
        return (
            "🔴 CHẾ ĐỘ 4: RỦI RO CAO / SẬP GÃY NỀN",
            "Khóa các bộ mua Breakout vượt đỉnh! Chỉ canh Bắt đáy Quá bán (Bộ 03, 25) hoặc Phòng thủ (Bộ 22, 27).",
            0xFF0000,
        )
    elif vni_1d > 0.1 and decliners > advancers * 1.25:
        return (
            "🟠 CHẾ ĐỘ 3: XANH VỎ ĐỎ LÒNG (KÉO TRỤ XẢ MIDCAP)",
            "Cẩn trọng bẫy Bull-trap ở Midcap! Ưu tiên Trụ (Bộ 28), Kháng bão (Bộ 04, 22) & Kéo ngược MA20/50 (Bộ 05, 20).",
            0xFF9900,
        )
    elif pct_above_ma20 >= 50.0 and vni_5d >= 0.0:
        return (
            "🟢 CHẾ ĐỘ 1: UPTREND ĐỒNG THUẬN (BÙNG NỔ)",
            "Mở khóa toàn bộ 30 Bộ Thuật Toán! Ưu tiên đánh Breakout & Dòng tiền mạnh (Bộ 01, 04, 11, 12, 13, 14, 30).",
            0x00FF66,
        )
    else:
        return (
            "🟡 CHẾ ĐỘ 2: SIDEWAY PHÂN HÓA / TÍCH LŨY",
            "Mua tại nền, bán khi nổ! Ưu tiên Gom nền VCP, Cạn cung VSA & Cá mập gom ngầm (Bộ 02, 05, 08, 15, 21, 26).",
            0x00BFFF,
        )


def analyze_full_vn_market(m_data, df_all):
    console = Console()
    if df_all.empty:
        console.print(
            "[bold red]❌ Không tải được danh sách cổ phiếu. Vui lòng kiểm tra kết nối mạng![/bold red]"
        )
        return

    total_tickers = len(df_all)
    df_active = df_all[df_all["volume"] > 0].copy()
    active_count = len(df_active) if len(df_active) > 0 else 1

    advancers = len(df_active[df_active["r1d"] > 0.1])
    decliners = len(df_active[df_active["r1d"] < -0.1])
    bo_cung_count = int(df_all["is_bo_cung"].sum())
    pct_above_ma20 = (df_active["above_ma20"].sum() / active_count) * 100
    total_market_val_billion = df_active["val_billion"].sum()
    session_status_str = str(df_all["session_label"].iloc[0])

    z_dxy = calc_zscore(m_data["DXY"]["close"])
    z_usdvnd = calc_zscore(m_data["USDVND"]["close"])
    usdvnd_price = m_data["USDVND"]["close"].iloc[-1]

    vni_c = m_data["VNINDEX"]["close"]
    vni_price = vni_c.iloc[-1]
    vni_1d = calc_roc(vni_c, 1)
    vni_5d = calc_roc(vni_c, 5)
    vni_20d = calc_roc(vni_c, 20)

    vn30_val = m_data["VN30"]["close"].iloc[-1]
    f1m_val = m_data["VN30F1M"]["close"].iloc[-1]
    basis_spread = (
        (f1m_val - vn30_val) if (f1m_val > 500 and vn30_val > 500) else 0.0
    )

    df_all["rs_1d"] = df_all["r1d"] - vni_1d
    df_all["rs_5d"] = df_all["r5d"] - vni_5d
    df_all["rs_20d"] = df_all["r20d"] - vni_20d

    def get_cap_tier(row):
        if row["mcap_billion"] >= 10000 or row["val_proj_billion"] >= 80:
            return "Large-Cap (Bluechip)"
        elif row["mcap_billion"] >= 1500 or row["val_proj_billion"] >= 10:
            return "Mid-Cap (Dòng tiền)"
        elif row["val_proj_billion"] >= 1.0:
            return "Small-Cap (Penny)"
        return "Micro / Kém TK"

    df_all["cap_tier"] = df_all.apply(get_cap_tier, axis=1)

    # 1. CHẠY 30 BỘ THUẬT TOÁN ĐỘC LẬP
    algo_strict_results, algo_display_results = (
        run_thirty_independent_algorithms(df_all)
    )

    # 2. DỰNG BIỂU ĐỒ NẾN T+0 (MATPLOTLIB + QUICKCHART CANDLESTICK SHORT URL)
    chart_bytes_map, chart_url_map, chart_b64_map = generate_t0_memory_charts(
        algo_display_results, algo_strict_results
    )

    export_cols = [
        "ticker", "exchange", "sector", "cap_tier", "price", "val_billion",
        "r1d", "r5d", "r20d", "r60d", "rs_5d", "rvol", "rsi", "pe", "roe",
        "target_price", "stop_price", "buy_zone", "is_bo_cung",
        "foreign_net_b", "confluence_count", "company",
    ]
    df_export = (
        df_all[export_cols]
        .sort_values(by="confluence_count", ascending=False)
        .copy()
    )
    try:
        df_export.to_csv(
            CSV_EXPORT_FILENAME, index=False, encoding="utf-8-sig"
        )
    except Exception:
        pass

    df_large = df_all[df_all["cap_tier"] == "Large-Cap (Bluechip)"]
    df_mid = df_all[df_all["cap_tier"] == "Mid-Cap (Dòng tiền)"]
    df_small = df_all[df_all["cap_tier"] == "Small-Cap (Penny)"]

    large_5d = df_large["r5d"].mean() if not df_large.empty else 0.0
    mid_5d = df_mid["r5d"].mean() if not df_mid.empty else 0.0
    small_5d = df_small["r5d"].mean() if not df_small.empty else 0.0
    large_val, mid_val, small_val = (
        df_large["val_billion"].sum(),
        df_mid["val_billion"].sum(),
        df_small["val_billion"].sum(),
    )

    df_liquid_sec = df_all[
        (df_all["val_proj_billion"] >= 1.0) & (~df_all["is_bo_cung"])
    ].copy()
    sector_ranks = []
    for sec_name, grp in df_liquid_sec.groupby("sector"):
        if len(grp) < 2:
            continue
        r1d_avg, r5d_avg, r20d_avg = (
            float(grp["r1d"].mean()),
            float(grp["r5d"].mean()),
            float(grp["r20d"].mean()),
        )
        rvol_avg = float(grp["rvol"].clip(upper=4.0).mean())
        sec_val = float(grp["val_billion"].sum())
        sec_score = (
            (0.35 * r1d_avg)
            + (0.40 * r5d_avg)
            + (0.25 * r20d_avg)
            + ((rvol_avg - 1.0) * 3.0 if r1d_avg > 0 else 0.0)
        )

        top_grp = grp[
            grp["val_proj_billion"] >= MIN_VALUE_TRADED_BILLION_VND
        ].sort_values(by="algo30_score", ascending=False)
        if top_grp.empty:
            top_grp = grp.sort_values(by="algo30_score", ascending=False)
        leaders_str = ", ".join(
            [
                f"{r['ticker']} ({r['r1d']:+.1f}%)"
                for _, r in top_grp.head(3).iterrows()
            ]
        )

        sector_ranks.append({
            "name": sec_name,
            "count": len(grp),
            "val_billion": sec_val,
            "r1d": r1d_avg,
            "r5d": r5d_avg,
            "r20d": r20d_avg,
            "rvol": rvol_avg,
            "score": sec_score,
            "leaders": leaders_str,
        })
    sector_ranks.sort(key=lambda x: x["score"], reverse=True)

    eval_fx = (
        "🔴 Tỷ giá căng thẳng"
        if z_usdvnd > 1.0
        else "🟢 Tỷ giá USD/VND Ổn định"
    )
    top1_sec = (
        sector_ranks[0]
        if sector_ranks
        else {"name": "N/A", "leaders": "N/A", "val_billion": 0}
    )
    top2_sec = sector_ranks[1] if len(sector_ranks) > 1 else top1_sec
    worst_sec = sector_ranks[-1] if sector_ranks else top1_sec

    regime_title, regime_advice, hex_col = detect_market_regime(
        vni_1d, vni_5d, pct_above_ma20, advancers, decliners, z_usdvnd
    )
    from_w = f"Nhóm suy yếu ({worst_sec['name']}) & Cổ phiếu gãy nền"
    to_w = f"Tập trung vào {top1_sec['name']} ({top1_sec['val_billion']:,.0f} Tỷ) & {top2_sec['name']}"
    ben = f"{regime_title} ➔ {regime_advice}"

    f_pct = lambda v: f"{'+' if v > 0 else ''}{v:.2f}%"

    # 3. ĐỒNG BỘ LÊN GOOGLE SHEETS (4 SHEET KÈM HÌNH ẢNH BIỂU ĐỒ NẾN)
    gs_status, total_tracked, new_sigs = sync_google_sheets_and_local_excel(
        df_all,
        algo_strict_results,
        df_export,
        vni_price,
        chart_url_map,
        chart_b64_map,
    )

    # 4. GỬI 5 GÓI TIN BÁO CÁO KÈM LƯỚI ẢNH ĐỒ THỊ NẾN T+0 SANG DISCORD
    def format_algo_discord_lines(idx_algo, df_sub, score_col, badge):
        pass_col = f"algo{idx_algo}_pass"
        lines = []
        for idx_r, (_, st) in enumerate(df_sub.head(5).iterrows(), 1):
            is_p = bool(st[pass_col])
            p_tag = "✅" if is_p else "⏳"
            conf_tag = (
                f" | `{int(st['confluence_count'])} Bộ`"
                if idx_algo == 30
                else ""
            )
            lines.append(
                f"{badge}{p_tag} **#{idx_r} {st['ticker']}**"
                f" (`{st['price']:.2f}` | `{st['buy_zone'][:10]}`) | GTGD:"
                f" `{st['val_billion']:,.1f}Tỷ` | 1D: `{f_pct(st['r1d'])}` |"
                f" RS: `{f_pct(st['rs_5d'])}` | Vol: `{st['rvol']:.2f}x` |"
                f" TP/SL: `{st['target_price']:.1f}/{st['stop_price']:.1f}` |"
                f" Điểm: **`{st[score_col]:+.1f}`**{conf_tag}"
            )
        return (
            "\n".join(lines)[:1020]
            if lines
            else "Không có mã nào đạt tiêu chuẩn trong phiên hôm nay."
        )

    sec_lines = []
    for i, s in enumerate(sector_ranks[:5]):
        icon = "🥇" if i == 0 else ("🥈" if i == 1 else ("🥉" if i == 2 else "🔹"))
        sec_lines.append(
            f"{icon} **{s['name']}** (`{s['val_billion']:,.0f} Tỷ`): 1D"
            f" `{f_pct(s['r1d'])}` | 5D `{f_pct(s['r5d'])}` ➔"
            f" **{s['leaders']}**"
        )

    if DISCORD_WEBHOOK_URL:
        batches = [
            (1, 6, "PHẦN 1/5: VĨ MÔ, CHẾ ĐỘ THỊ TRƯỜNG & BỘ 01 ➔ 06"),
            (7, 12, "PHẦN 2/5: BỘ THUẬT TOÁN 07 ➔ 12"),
            (13, 18, "PHẦN 3/5: BỘ THUẬT TOÁN 13 ➔ 18"),
            (19, 24, "PHẦN 4/5: BỘ THUẬT TOÁN 19 ➔ 24"),
            (25, 30, "PHẦN 5/5: BỘ THUẬT TOÁN 25 ➔ 30 & ĐỒNG BỘ GOOGLE SHEETS"),
        ]

        for b_idx, (start_a, end_a, batch_title) in enumerate(batches, 1):
            fields_list = []

            if b_idx == 1:
                fields_list.append({
                    "name": f"🧭 {regime_title}",
                    "value": (
                        f"• **Chiến lược:** **{regime_advice}**\n• **Phiên:**"
                        f" `{session_status_str}` | Đã lọc bỏ"
                        f" `{bo_cung_count}` mã Bo cung\n• **Tỷ giá USD/VND:**"
                        f" `{usdvnd_price:,.0f} đ` (`{z_usdvnd:+.2f}σ`) ➔"
                        f" **{eval_fx}** | **VN30F1M Basis:**"
                        f" `{basis_spread:+.1f} đ`\n" + "\n".join(sec_lines)
                    )[:1020],
                    "inline": False,
                })

            for idx_algo in range(start_a, end_a + 1):
                _, title_str, _, _, badge_ico = ALGO_METADATA[idx_algo - 1]
                strict_c = len(algo_strict_results[idx_algo])
                fields_list.append({
                    "name": f"{title_str} ({strict_c} mã Đạt chuẩn ✅)",
                    "value": format_algo_discord_lines(
                        idx_algo,
                        algo_display_results[idx_algo],
                        f"algo{idx_algo}_score",
                        badge_ico,
                    ),
                    "inline": False,
                })

            if b_idx == 5:
                fields_list.append({
                    "name": "🔄 VECTOR DỊCH CHUYỂN & SỔ CÁI GOOGLE SHEETS (KÈM ẢNH NẾN)",
                    "value": (
                        f"• **Dịch chuyển:** Từ **[{from_w}]** ➔ Sang"
                        f" **[{to_w}]**\n• **Trạng thái Cloud:** `{gs_status}`"
                        f" (Tổng theo dõi: `{total_tracked}` lệnh | Mới hôm"
                        f" nay: `+{new_sigs}` lệnh đạt chuẩn)"
                    )[:1020],
                    "inline": False,
                })

            grid_bytes = create_discord_grid_bytes(
                start_a, end_a, algo_display_results, chart_bytes_map
            )

            embed_obj = {
                "title": (
                    f"🇻🇳 BÁO CÁO QUÉT {total_tickers} CỔ PHIẾU - {batch_title}"
                ),
                "description": (
                    f"**VN-INDEX:** `{vni_price:,.1f}` (`{f_pct(vni_1d)}`) |"
                    f" **Tổng GTGD:** `{total_market_val_billion:,.0f} Tỷ VNĐ`"
                    f" | **Độ rộng:** 🟢`{advancers}` Tăng / 🔴`{decliners}`"
                    f" Giảm (`{pct_above_ma20:.1f}%` > MA20)"
                    if b_idx == 1
                    else "Chú thích: `✅` = Đạt chuẩn 100% (Lưu vào Google Sheets kèm Ảnh Nến T+0) | `⏳` = Theo dõi | `TP/SL` = Động theo ATR(14)"
                ),
                "color": hex_col,
                "fields": fields_list,
            }

            if grid_bytes:
                embed_obj["image"] = {
                    "url": f"attachment://grid_part_{b_idx}.png"
                }

            payload_discord = {
                "username": "VN Hybrid Cloud 30-Algo Quant Bot",
                "embeds": [embed_obj],
            }

            try:
                if grid_bytes:
                    requests.post(
                        DISCORD_WEBHOOK_URL,
                        data={"payload_json": json.dumps(payload_discord)},
                        files={
                            "file": (
                                f"grid_part_{b_idx}.png",
                                grid_bytes,
                                "image/png",
                            )
                        },
                        timeout=20,
                    )
                else:
                    requests.post(
                        DISCORD_WEBHOOK_URL, json=payload_discord, timeout=15
                    )
                time.sleep(0.8)
            except Exception as e:
                console.print(
                    f"[bold red]❌ Lỗi kết nối Discord gói {b_idx}: {e}[/bold red]"
                )

    summary_text = (
        f"⏱️ [bold white]TRẠNG THÁI PHIÊN & RVOL:[/bold white] [bold cyan]{session_status_str}[/bold cyan] | Đã lọc bỏ [bold red]{bo_cung_count}[/bold red] cổ phiếu Bo cung\n"
        f"☁️ [bold green]ĐỒNG BỘ GOOGLE SHEETS & ẢNH NẾN:[/bold green] [bold yellow]{gs_status}[/bold yellow] (Mới hôm nay: [bold green]+{new_sigs}[/bold green] lệnh đạt chuẩn)\n"
        f"🧭 [bold white]CHẾ ĐỘ THỊ TRƯỜNG:[/bold white] [bold yellow]{ben}[/bold yellow]"
    )
    console.print(Panel(summary_text, box=box.HEAVY, border_style="bright_blue"))


def main_handler():
    macro_df, full_stocks_df = load_macro_and_full_market()
    analyze_full_vn_market(macro_df, full_stocks_df)
    return {"status": "success", "timestamp": int(time.time())}


if __name__ == "__main__":
    main_handler()
