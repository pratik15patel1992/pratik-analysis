import os
import json
import math
import time
import threading
import hmac
from pathlib import Path
from datetime import datetime, timedelta
from io import BytesIO

from flask import (
    Flask,
    render_template,
    redirect,
    request,
    jsonify,
    send_file,
    session,
    url_for,
)

from kiteconnect import KiteConnect, KiteTicker

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment

import psycopg
from psycopg.types.json import Jsonb


# ============================================================
# CONFIGURATION
# ============================================================

APP_SECRET = os.environ.get("APP_SECRET", "change-me")
KITE_API_KEY = os.environ.get("KITE_API_KEY", "")
KITE_API_SECRET = os.environ.get("KITE_API_SECRET", "")
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")
DATABASE_URL = os.environ.get("DATABASE_URL", "")
PORTAL_USERNAME = os.environ.get("PORTAL_USERNAME", "")
PORTAL_PASSWORD = os.environ.get("PORTAL_PASSWORD", "")

app = Flask(__name__)
app.secret_key = APP_SECRET
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=True,
    PERMANENT_SESSION_LIFETIME=timedelta(days=7),
)

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

ACCESS_TOKEN_FILE = DATA / "access_token.json"
BASELINE_FILE = DATA / "oi_baseline.json"

HISTORY_DIR = DATA / "history"
HISTORY_DIR.mkdir(exist_ok=True)

lock = threading.RLock()

IST_OFFSET = timedelta(hours=5, minutes=30)

MARKET_START_HOUR = 9
MARKET_START_MINUTE = 15
MARKET_END_HOUR = 15
MARKET_END_MINUTE = 30


# ============================================================
# STATE
# ============================================================

state = {
    "configured": bool(KITE_API_KEY and KITE_API_SECRET),
    "connected": False,
    "message": "Waiting for Zerodha login",
    "last_update": None,

    "date": None,
    "expiry": None,

    "nifty": {
        "price": None,
        "previous_close": None,
        "open": None,
        "high": None,
        "low": None,
        "change": None,
        "change_pct": None,
    },

    "vix": {
        "price": None,
        "range": None,
        "interpretation": None,
    },

    "zone": {
        "quadrant": None,
        "zone": None,
        "opening_pct": None,
        "levels": {},
    },

    "opening_atm": None,

    "oic": {
        "atm": None,
        "minus100": None,
        "plus100": None,
    },

    # Locked six option contracts used for premium analytics.
    # Populated once opening ATM is known.
    "premium": {
        "atm": None,
        "minus100": None,
        "plus100": None,
    },

    "series": {
        "nifty": [],
        "atm": [],
        "minus100": [],
        "plus100": [],
        "cio": [],
        # True 1-minute option premium OHLC, built from live ticks.
        # Each strike bucket contains CE and PE candles.
        "premium": {
            "minus100": {"CE": [], "PE": []},
            "atm": {"CE": [], "PE": []},
            "plus100": {"CE": [], "PE": []},
        },
    },

    "history_dates": [],

    # Strategy trade ledger. The live S1/S2/PNA engine will populate
    # these lists once the exact signal rules are activated.
    "strategies": {
        "S1": {"trades": [], "active": None, "signals": []},
        "S2": {"trades": [], "active": None, "signals": []},
        "PNA": {"trades": [], "active": None, "signals": []},
    },
    "strategy_engine": {
        "enabled": True,
        "version": "LIVE-OIC-CIO-1",
        "last_eval_minute": None,
        "metrics": {},
    },
}


# ============================================================
# GLOBAL LIVE OBJECTS
# ============================================================

kite = None
ticker = None

nifty_token = None
vix_token = None

option_instruments = []
token_meta = {}
oic_tokens = {}
premium_token_map = {}

latest_oi = {}
prev_oi = {}

baseline_ready = False
baseline_thread_started = False

# Stage 7C: server-side Kite watchdog.  Render free instances can restart or
# temporarily lose the WebSocket even though today's access token is still
# valid in Neon.  The watchdog reconnects without requiring the browser.
reconnect_watchdog_started = False
reconnect_in_progress = False
last_reconnect_attempt = 0.0
live_start_mutex = threading.Lock()
snapshot_thread_started = False

latest_nifty = {}
latest_vix = {}

# Live 1-minute NIFTY OHLC aggregator. Kite's tick ``ohlc`` field is
# session/day OHLC, not a 1-minute candle, so we build true minute candles
# ourselves from the live NIFTY last_price ticks.
nifty_minute_bucket = None

# Live minute buckets for the six locked option contracts.
# Keyed by instrument token so CE/PE candles are independently aggregated.
premium_minute_buckets = {}

tick_counter = 0
oi_tick_counter = 0


# ============================================================
# TIME
# ============================================================

def now_ist():
    return datetime.utcnow() + IST_OFFSET


def today_key():
    return now_ist().strftime("%Y-%m-%d")


def minute_label():
    return now_ist().strftime("%H:%M")


def is_market_session():
    n = now_ist()

    current = n.hour * 60 + n.minute
    start = MARKET_START_HOUR * 60 + MARKET_START_MINUTE
    end = MARKET_END_HOUR * 60 + MARKET_END_MINUTE

    return start <= current <= end


# ============================================================
# JSON
# ============================================================

def safe_json_load(path, default=None):
    if default is None:
        default = {}

    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)

    except Exception as e:
        print(f"[DIAG] JSON load error: {e}", flush=True)

    return default


def safe_json_save(path, data):
    try:
        temp = path.with_suffix(path.suffix + ".tmp")

        with open(temp, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2,
            )

        temp.replace(path)

    except Exception as e:
        print(f"[DIAG] JSON save error: {e}", flush=True)


# ============================================================
# PERMANENT DATABASE (NEON POSTGRESQL)
# ============================================================

def db_enabled():
    return bool(DATABASE_URL)


def db_connect():
    if not DATABASE_URL:
        return None

    return psycopg.connect(
        DATABASE_URL,
        connect_timeout=10,
    )


def init_db():
    if not db_enabled():
        print(
            "[DB] DATABASE_URL not configured; using local history fallback.",
            flush=True,
        )
        return False

    try:
        with db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS pratik_daily_history (
                        day DATE PRIMARY KEY,
                        payload JSONB NOT NULL,
                        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                    )
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS pratik_kite_session (
                        session_id SMALLINT PRIMARY KEY CHECK (session_id = 1),
                        token_day DATE NOT NULL,
                        access_token TEXT NOT NULL,
                        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                    )
                    """
                )
            conn.commit()

        print(
            "[DB] Neon history table ready.",
            flush=True,
        )
        return True

    except Exception as e:
        print(
            f"[DB] Database initialization failed: {e}",
            flush=True,
        )
        return False


def save_history_to_db(day, payload):
    if not db_enabled():
        return False

    try:
        with db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO pratik_daily_history (day, payload, updated_at)
                    VALUES (%s, %s, NOW())
                    ON CONFLICT (day)
                    DO UPDATE SET
                        payload = EXCLUDED.payload,
                        updated_at = NOW()
                    """,
                    (day, Jsonb(payload)),
                )
            conn.commit()

        return True

    except Exception as e:
        print(
            f"[DB] History save failed for {day}: {e}",
            flush=True,
        )
        return False


def load_history_from_db(day):
    if not db_enabled():
        return None

    try:
        with db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT payload
                    FROM pratik_daily_history
                    WHERE day = %s
                    """,
                    (day,),
                )
                row = cur.fetchone()

        if not row:
            return None

        payload = row[0]

        if isinstance(payload, str):
            payload = json.loads(payload)

        return payload

    except Exception as e:
        print(
            f"[DB] History load failed for {day}: {e}",
            flush=True,
        )
        return None


def db_history_dates():
    if not db_enabled():
        return []

    try:
        with db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT day
                    FROM pratik_daily_history
                    ORDER BY day DESC
                    """
                )
                rows = cur.fetchall()

        return [row[0].isoformat() for row in rows]

    except Exception as e:
        print(
            f"[DB] History date load failed: {e}",
            flush=True,
        )
        return []


def history_exists(day):
    if load_history_from_db(day) is not None:
        return True

    return history_file(day).exists()


# ============================================================
# HISTORY
# ============================================================

def history_file(day):
    return HISTORY_DIR / f"{day}.json"


def load_day_history(day):
    db_data = load_history_from_db(day)

    if db_data is not None:
        return db_data

    return safe_json_load(
        history_file(day),
        {
            "date": day,
            "expiry": None,
            "opening_atm": None,
            "nifty": {},
            "vix": {},
            "zone": {},
            "oic": {},
            "series": {
                "nifty": [],
                "atm": [],
                "minus100": [],
                "plus100": [],
                "cio": [],
                "premium": {
                    "minus100": {"CE": [], "PE": []},
                    "atm": {"CE": [], "PE": []},
                    "plus100": {"CE": [], "PE": []},
                },
            },
            "premium": {
                "atm": None,
                "minus100": None,
                "plus100": None,
            },
            "strategies": {
                "S1": {"trades": [], "active": None, "signals": []},
                "S2": {"trades": [], "active": None, "signals": []},
                "PNA": {"trades": [], "active": None, "signals": []},
            },
            "strategy_engine": {"enabled": True, "version": "LIVE-OIC-CIO-1", "last_eval_minute": None, "metrics": {}},
        },
    )


def refresh_history_dates():
    dates = set(db_history_dates())

    for f in HISTORY_DIR.glob("*.json"):
        dates.add(f.stem)

    dates = sorted(
        dates,
        reverse=True,
    )

    with lock:
        state["history_dates"] = dates


def save_current_history():
    with lock:
        day = state.get("date") or today_key()

        payload = {
            "date": day,
            "expiry": state.get("expiry"),
            "opening_atm": state.get("opening_atm"),
            "nifty": state.get("nifty", {}),
            "vix": state.get("vix", {}),
            "zone": state.get("zone", {}),
            "oic": state.get("oic", {}),
            "premium": state.get("premium", {}),
            "series": state.get("series", {}),
            "strategies": state.get("strategies", {}),
            "strategy_engine": state.get("strategy_engine", {}),
            "saved_at": now_ist().isoformat(),
        }

    # Local file remains as a temporary fallback.
    safe_json_save(
        history_file(day),
        payload,
    )

    # Neon PostgreSQL is the permanent source of truth.
    save_history_to_db(
        day,
        payload,
    )

    refresh_history_dates()


# ============================================================
# ACCESS TOKEN
# ============================================================
# The daily Kite access token is persisted in Neon so a Render restart
# does not depend on Render's temporary local filesystem.  We keep the
# local JSON file only as a fallback when the database is unavailable.

def save_access_token(token):
    token_day = now_ist().date()
    saved_to_db = False

    if db_enabled():
        try:
            with db_connect() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        INSERT INTO pratik_kite_session
                            (session_id, token_day, access_token, updated_at)
                        VALUES (1, %s, %s, NOW())
                        ON CONFLICT (session_id)
                        DO UPDATE SET
                            token_day = EXCLUDED.token_day,
                            access_token = EXCLUDED.access_token,
                            updated_at = NOW()
                        """,
                        (token_day, token),
                    )
                conn.commit()
            saved_to_db = True
            print("[KITE] Today's access token saved to Neon.", flush=True)
        except Exception as e:
            print(f"[KITE] Neon token save failed: {e}", flush=True)

    # Fallback copy. It is not relied on across Render restarts.
    safe_json_save(
        ACCESS_TOKEN_FILE,
        {
            "access_token": token,
            "token_day": token_day.isoformat(),
            "saved_at": now_ist().isoformat(),
            "database_saved": saved_to_db,
        },
    )


def load_access_token():
    today = now_ist().date()

    if db_enabled():
        try:
            with db_connect() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        SELECT access_token
                        FROM pratik_kite_session
                        WHERE session_id = 1 AND token_day = %s
                        """,
                        (today,),
                    )
                    row = cur.fetchone()
            if row and row[0]:
                print("[KITE] Restored today's access token from Neon.", flush=True)
                return row[0]
        except Exception as e:
            print(f"[KITE] Neon token load failed: {e}", flush=True)

    # Local fallback is accepted only if it belongs to today.
    data = safe_json_load(ACCESS_TOKEN_FILE, {})
    token = data.get("access_token")
    token_day = data.get("token_day")
    if token and token_day == today.isoformat():
        print("[KITE] Restored today's access token from local fallback.", flush=True)
        return token

    return None


def clear_access_token():
    if db_enabled():
        try:
            with db_connect() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "DELETE FROM pratik_kite_session WHERE session_id = 1"
                    )
                conn.commit()
        except Exception as e:
            print(f"[KITE] Neon token clear failed: {e}", flush=True)

    try:
        ACCESS_TOKEN_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def restore_kite_session():
    """Reconnect server-side after a Render restart, without a browser."""
    if not (KITE_API_KEY and KITE_API_SECRET):
        return False

    token = load_access_token()
    if not token:
        return False

    try:
        start_live(token)
        print("[KITE] Server-side session restored successfully.", flush=True)
        return True
    except Exception as e:
        print(f"[KITE] Saved session restore failed: {e}", flush=True)
        clear_access_token()
        with lock:
            state["connected"] = False
            state["message"] = "Login required"
        return False


# ============================================================
# ZONE ENGINE
# ============================================================

def round_to_100(value):
    return int(
        round(float(value) / 100.0) * 100
    )


def calculate_zone(open_price, previous_close):
    if not open_price or not previous_close:
        return {
            "quadrant": None,
            "zone": None,
            "opening_pct": None,
            "levels": {},
        }

    opening_pct = (
        (open_price - previous_close)
        / previous_close
        * 100
    )

    quadrant = "Q1" if opening_pct >= 0 else "Q2"

    abs_pct = abs(opening_pct)

    if abs_pct <= 0.25:
        zone = "Z0"
    elif abs_pct <= 0.50:
        zone = "Z1"
    elif abs_pct <= 0.75:
        zone = "Z2"
    elif abs_pct <= 1.00:
        zone = "Z3"
    else:
        zone = "Outside Z3"

    percentages = [
        -1.00,
        -0.75,
        -0.50,
        -0.25,
        0.00,
        0.25,
        0.50,
        0.75,
        1.00,
    ]

    levels = {}

    for pct in percentages:
        value = previous_close * (
            1 + pct / 100
        )

        if pct > 0:
            key = f"+{pct:.2f}%"
        else:
            key = f"{pct:.2f}%"

        levels[key] = round(value, 2)

    return {
        "quadrant": quadrant,
        "zone": zone,
        "opening_pct": round(
            opening_pct,
            3,
        ),
        "levels": levels,
    }


# ============================================================
# INDIA VIX
# ============================================================

def classify_vix(value):
    if value is None:
        return None, None

    value = float(value)

    if value < 12:
        return (
            "<12 LOW",
            "Low volatility. Market may be relatively calm; option premiums can be lower.",
        )

    if value < 15:
        return (
            "12–15 NORMAL",
            "Normal volatility zone.",
        )

    if value < 20:
        return (
            "15–20 ELEVATED",
            "Elevated volatility. Expect larger intraday movement.",
        )

    if value < 25:
        return (
            "20–25 HIGH",
            "High volatility. Use additional caution.",
        )

    return (
        "≥25 VERY HIGH",
        "Very high volatility. Large and rapid market movement is possible.",
    )


# ============================================================
# INSTRUMENT DISCOVERY
# ============================================================

def discover_instruments(k):
    global nifty_token
    global vix_token
    global option_instruments
    global token_meta

    print(
        "[DIAG] Loading instruments...",
        flush=True,
    )

    nse = k.instruments("NSE")
    nfo = k.instruments("NFO")

    print(
        f"[DIAG] NSE={len(nse)} NFO={len(nfo)}",
        flush=True,
    )

    for row in nse:
        symbol = str(
            row.get("tradingsymbol", "")
        ).upper()

        name = str(
            row.get("name", "")
        ).upper()

        if symbol == "NIFTY 50" or name == "NIFTY 50":
            nifty_token = int(
                row["instrument_token"]
            )

        if symbol == "INDIA VIX" or name == "INDIA VIX":
            vix_token = int(
                row["instrument_token"]
            )

    today = now_ist().date()

    candidates = []

    for row in nfo:
        name = str(
            row.get("name", "")
        ).upper()

        inst_type = str(
            row.get("instrument_type", "")
        ).upper()

        if name != "NIFTY":
            continue

        if inst_type not in ("CE", "PE"):
            continue

        expiry = row.get("expiry")

        if not expiry:
            continue

        if isinstance(expiry, str):
            try:
                expiry = datetime.strptime(
                    expiry,
                    "%Y-%m-%d",
                ).date()
            except Exception:
                continue

        if expiry < today:
            continue

        candidates.append(row)

    if not candidates:
        raise RuntimeError(
            "No active NIFTY option contracts found."
        )

    nearest_expiry = min(
        row["expiry"]
        for row in candidates
    )

    option_instruments = [
        row
        for row in candidates
        if row["expiry"] == nearest_expiry
    ]

    token_meta = {}

    for row in option_instruments:
        token = int(
            row["instrument_token"]
        )

        token_meta[token] = {
            "strike": int(
                float(row["strike"])
            ),
            "type": row["instrument_type"],
            "symbol": row["tradingsymbol"],
            "expiry": str(row["expiry"]),
        }

    with lock:
        state["expiry"] = str(
            nearest_expiry
        )

    print(
        f"[DIAG] NIFTY token={nifty_token}",
        flush=True,
    )

    print(
        f"[DIAG] VIX token={vix_token}",
        flush=True,
    )

    print(
        f"[DIAG] Nearest expiry={nearest_expiry}",
        flush=True,
    )

    print(
        f"[DIAG] Option contracts loaded={len(option_instruments)}",
        flush=True,
    )


# ============================================================
# OIC STRIKE LOCK
# ============================================================

def lock_oic_strikes():
    global oic_tokens
    global premium_token_map

    with lock:
        atm = state.get(
            "opening_atm"
        )

    if not atm:
        print(
            "[DIAG] Opening ATM missing",
            flush=True,
        )
        return False

    wanted = {
        "atm": atm,
        "minus100": atm - 100,
        "plus100": atm + 100,
    }

    mapping = {}

    for key, strike in wanted.items():
        ce_token = None
        pe_token = None

        for token, meta in token_meta.items():
            if meta["strike"] != strike:
                continue

            if meta["type"] == "CE":
                ce_token = token

            elif meta["type"] == "PE":
                pe_token = token

        if ce_token and pe_token:
            mapping[key] = {
                "strike": strike,
                "CE": ce_token,
                "PE": pe_token,
            }

    oic_tokens = mapping

    # Reverse token lookup used by the premium OHLC aggregator.
    premium_token_map = {}
    for key, legs in mapping.items():
        for option_type in ("CE", "PE"):
            token = legs.get(option_type)
            if token:
                premium_token_map[int(token)] = {
                    "key": key,
                    "type": option_type,
                    "strike": int(legs["strike"]),
                }

    with lock:
        state["oic"]["atm"] = atm
        state["oic"]["minus100"] = atm - 100
        state["oic"]["plus100"] = atm + 100
        state["premium"]["atm"] = atm
        state["premium"]["minus100"] = atm - 100
        state["premium"]["plus100"] = atm + 100

    print(
        f"[DIAG] OIC strike mapping={mapping}",
        flush=True,
    )

    return len(mapping) == 3


def ensure_locked_strike_mappings(k=None):
    """
    Fail-safe for the three locked OIC/premium strikes.

    Stage-7 could remain with empty oic_tokens/premium_token_map when the
    server restored mid-session before state['opening_atm'] had been rebuilt.
    This routine reconstructs the opening ATM from the best available source
    and then rebuilds both token maps. It is safe to call repeatedly.
    """
    global oic_tokens
    global premium_token_map

    # Already healthy.
    if len(oic_tokens) == 3 and len(premium_token_map) == 6:
        return True

    with lock:
        atm = state.get("opening_atm")
        state_nifty = dict(state.get("nifty") or {})

    # First prefer an already-known session open.
    open_price = state_nifty.get("open") or latest_nifty.get("open")
    previous_close = (
        state_nifty.get("previous_close")
        or latest_nifty.get("previous_close")
    )

    # If a Render restart happened mid-session and no NIFTY FULL tick has yet
    # rebuilt the open, fetch one current quote from Kite REST.
    if not atm and open_price is None and k is not None:
        try:
            quote = k.quote(["NSE:NIFTY 50"]) or {}
            row = quote.get("NSE:NIFTY 50") or {}
            q_ohlc = row.get("ohlc") or {}
            open_price = q_ohlc.get("open")
            previous_close = previous_close or q_ohlc.get("close")

            if row.get("last_price") is not None:
                latest_nifty["price"] = float(row["last_price"])
            if open_price is not None:
                latest_nifty["open"] = float(open_price)
            if q_ohlc.get("high") is not None:
                latest_nifty["high"] = float(q_ohlc["high"])
            if q_ohlc.get("low") is not None:
                latest_nifty["low"] = float(q_ohlc["low"])
            if previous_close is not None:
                latest_nifty["previous_close"] = float(previous_close)

            with lock:
                state["nifty"] = dict(latest_nifty)

            print(
                f"[DIAG] Mapping repair quote open={open_price} pc={previous_close}",
                flush=True,
            )
        except Exception as e:
            print(f"[DIAG] Mapping repair quote failed: {e}", flush=True)

    if not atm and open_price is not None:
        atm = round_to_100(open_price)
        with lock:
            state["opening_atm"] = atm
            if previous_close:
                state["zone"] = calculate_zone(open_price, previous_close)
        print(f"[DIAG] Reconstructed opening ATM={atm}", flush=True)

    if not atm:
        print("[DIAG] Mapping repair waiting for NIFTY session open", flush=True)
        return False

    ok = lock_oic_strikes()

    print(
        f"[DIAG] Mapping repair result ok={ok} OIC={len(oic_tokens)}/3 PREMIUM={len(premium_token_map)}/6",
        flush=True,
    )
    return ok and len(premium_token_map) == 6


# ============================================================
# PREVIOUS DAY OI BASELINE
# ============================================================

def previous_oi_for_token(k, token):
    end = (
        now_ist().date()
        - timedelta(days=1)
    )

    start = (
        end
        - timedelta(days=10)
    )

    try:
        candles = k.historical_data(
            token,
            start,
            end,
            "day",
            oi=True,
        )

        if not candles:
            return None

        for candle in reversed(candles):
            if candle.get("oi") is not None:
                return int(candle["oi"])

    except Exception as e:
        print(
            f"[DIAG] Historical OI failed token={token}: {e}",
            flush=True,
        )

    return None


def build_oi_baseline():
    global baseline_ready
    global baseline_thread_started
    global prev_oi

    baseline_thread_started = True

    try:
        cached = safe_json_load(
            BASELINE_FILE,
            {},
        )

        cache_date = cached.get("date")
        cache_expiry = cached.get("expiry")

        with lock:
            current_expiry = state.get(
                "expiry"
            )

        if (
            cache_date == today_key()
            and cache_expiry == current_expiry
            and cached.get("oi")
        ):
            prev_oi = {
                int(k): int(v)
                for k, v in cached[
                    "oi"
                ].items()
            }

            baseline_ready = True

            print(
                f"[DIAG] CIO baseline loaded. Contracts={len(prev_oi)}",
                flush=True,
            )

            return

        result = {}

        for index, row in enumerate(
            option_instruments,
            start=1,
        ):
            token = int(
                row["instrument_token"]
            )

            value = previous_oi_for_token(
                kite,
                token,
            )

            if value is not None:
                result[token] = value

            if index % 20 == 0:
                print(
                    f"[DIAG] Baseline {index}/{len(option_instruments)}",
                    flush=True,
                )

            time.sleep(0.35)

        prev_oi = result

        safe_json_save(
            BASELINE_FILE,
            {
                "date": today_key(),
                "expiry": state.get(
                    "expiry"
                ),
                "oi": {
                    str(k): v
                    for k, v in result.items()
                },
            },
        )

        baseline_ready = True

        print(
            f"[DIAG] CIO baseline ready. Contracts={len(prev_oi)}",
            flush=True,
        )

    except Exception as e:
        baseline_ready = False

        print(
            f"[DIAG] CIO baseline ERROR: {e}",
            flush=True,
        )

    finally:
        baseline_thread_started = False


# ============================================================
# CIO
# ============================================================

def cio_totals():
    ce = 0
    pe = 0

    for token, current in latest_oi.items():
        baseline = prev_oi.get(token)
        meta = token_meta.get(token)

        if baseline is None or not meta:
            continue

        delta = int(current) - int(
            baseline
        )

        if delta >= 0:
            continue

        if meta["type"] == "CE":
            ce += delta

        elif meta["type"] == "PE":
            pe += delta

    return ce, pe


# ============================================================
# SERIES
# ============================================================

def append_or_replace_minute(
    series,
    point,
):
    if not series:
        series.append(point)
        return

    if (
        series[-1].get("time")
        == point.get("time")
    ):
        series[-1] = point

    else:
        series.append(point)


def oic_point(key):
    legs = oic_tokens.get(key)

    if not legs:
        return None

    ce = latest_oi.get(
        legs["CE"]
    )

    pe = latest_oi.get(
        legs["PE"]
    )

    if ce is None or pe is None:
        return None

    return {
        "time": minute_label(),
        "timestamp": now_ist().isoformat(),
        "ce": int(ce),
        "pe": int(pe),
    }


def update_nifty_minute_ohlc(price):
    """Aggregate live NIFTY ticks into a true 1-minute OHLC candle."""
    global nifty_minute_bucket

    if price is None:
        return

    p = float(price)
    n = now_ist()
    minute = n.strftime("%H:%M")
    timestamp = n.isoformat()

    with lock:
        # Minute rollover: finalise the previous candle in memory first.
        if (
            nifty_minute_bucket
            and nifty_minute_bucket.get("time") != minute
        ):
            finalised = dict(nifty_minute_bucket)
            finalised["price"] = finalised.get("close")
            append_or_replace_minute(
                state["series"]["nifty"],
                finalised,
            )
            nifty_minute_bucket = None

        if nifty_minute_bucket is None:
            nifty_minute_bucket = {
                "time": minute,
                "timestamp": timestamp,
                "open": p,
                "high": p,
                "low": p,
                "close": p,
                # Backward-compatible alias used by the existing chart/API.
                "price": p,
            }
        else:
            nifty_minute_bucket["high"] = max(
                float(nifty_minute_bucket.get("high", p)),
                p,
            )
            nifty_minute_bucket["low"] = min(
                float(nifty_minute_bucket.get("low", p)),
                p,
            )
            nifty_minute_bucket["close"] = p
            nifty_minute_bucket["price"] = p
            nifty_minute_bucket["timestamp"] = timestamp


def snapshot_current_nifty_candle():
    """Copy the current in-progress minute candle into state history."""
    with lock:
        if not nifty_minute_bucket:
            return

        point = dict(nifty_minute_bucket)
        point["price"] = point.get("close")
        append_or_replace_minute(
            state["series"]["nifty"],
            point,
        )


def update_premium_minute_ohlc(token, price, oi=None):
    """Aggregate one of the six locked option contracts into true 1-minute OHLC."""
    global premium_minute_buckets

    meta = premium_token_map.get(int(token))
    if not meta or price is None:
        return

    p = float(price)
    n = now_ist()
    minute = n.strftime("%H:%M")
    timestamp = n.isoformat()
    token = int(token)

    with lock:
        bucket = premium_minute_buckets.get(token)

        if bucket and bucket.get("time") != minute:
            finalised = dict(bucket)
            append_or_replace_minute(
                state["series"]["premium"][meta["key"]][meta["type"]],
                finalised,
            )
            bucket = None

        if bucket is None:
            bucket = {
                "time": minute,
                "timestamp": timestamp,
                "strike": meta["strike"],
                "type": meta["type"],
                "open": p,
                "high": p,
                "low": p,
                "close": p,
                "ltp": p,
                "oi": int(oi) if oi is not None else latest_oi.get(token),
            }
        else:
            bucket["high"] = max(float(bucket.get("high", p)), p)
            bucket["low"] = min(float(bucket.get("low", p)), p)
            bucket["close"] = p
            bucket["ltp"] = p
            bucket["timestamp"] = timestamp
            if oi is not None:
                bucket["oi"] = int(oi)
            elif latest_oi.get(token) is not None:
                bucket["oi"] = int(latest_oi[token])

        premium_minute_buckets[token] = bucket


def snapshot_current_premium_candles():
    """Copy all in-progress premium candles into state before persistence."""
    with lock:
        for token, bucket in list(premium_minute_buckets.items()):
            if not bucket:
                continue
            meta = premium_token_map.get(int(token))
            if not meta:
                continue
            append_or_replace_minute(
                state["series"]["premium"][meta["key"]][meta["type"]],
                dict(bucket),
            )



# ============================================================
# LIVE S1 / S2 / PNA SIGNAL ENGINE
# ============================================================
# The engine uses the locked three opening strikes and the OIC/CIO rules:
# - no entries during 09:15-09:30 observation period
# - direction is measured from rolling 5-minute OI change at ATM-100/ATM/ATM+100
# - bullish OI flow = PE OI rising while CE OI falls -> SELL PE
# - bearish OI flow = CE OI rising while PE OI falls -> SELL CE
# - CIO confirms bullish when CE negative-change is stronger than PE; bearish is the mirror
# - S1 requires >=2/3 agreement + CIO for 3 consecutive one-minute readings
# - S2 reacts to a visible momentum shift on >=2/3 strikes + CIO (earlier / no 3-read wait)
# - PNA requires >=2/3 agreement, weighted MasterD clear (|MasterD| >= 35),
#   CIO confirmation and 3 consecutive readings
# - hard SL = 30 NIFTY points; all open trades time-exit at/after 14:45
# Strategies are evaluated independently so S1/S2/PNA can coexist for research.

STRATEGY_MASTER_CLEAR = 35.0
STRATEGY_HARD_SL_POINTS = 30.0
STRATEGY_MAX_TRADES_PER_STRATEGY = 3


def _row_minutes(time_text):
    try:
        hh, mm = str(time_text).split(":")[:2]
        return int(hh) * 60 + int(mm)
    except Exception:
        return None


def _point_at_or_before(series, target_minutes):
    best = None
    for row in series or []:
        m = _row_minutes(row.get("time"))
        if m is None or m > target_minutes:
            continue
        if best is None or m > best[0]:
            best = (m, row)
    return best[1] if best else None


def _five_minute_delta(series):
    if not series:
        return None
    cur = series[-1]
    cur_m = _row_minutes(cur.get("time"))
    if cur_m is None:
        return None
    old = _point_at_or_before(series, cur_m - 5)
    if old is None:
        return None
    return {
        "ce": float(cur.get("ce", 0)) - float(old.get("ce", 0)),
        "pe": float(cur.get("pe", 0)) - float(old.get("pe", 0)),
    }


def _dominance(delta):
    if not delta:
        return 0.0
    ce, pe = float(delta["ce"]), float(delta["pe"])
    den = abs(ce) + abs(pe)
    return ((ce - pe) / den * 100.0) if den else 0.0


def _flow_direction(delta):
    if not delta:
        return None
    ce, pe = float(delta["ce"]), float(delta["pe"])
    if pe > 0 and ce < 0:
        return "PE"  # bullish -> sell PE
    if ce > 0 and pe < 0:
        return "CE"  # bearish -> sell CE
    return None


def _strategy_metrics():
    deltas = {}
    dirs = {}
    dom = {}
    for key in ("minus100", "atm", "plus100"):
        d = _five_minute_delta(state.get("series", {}).get(key, []))
        deltas[key] = d
        dirs[key] = _flow_direction(d)
        dom[key] = _dominance(d)

    pe_votes = sum(1 for x in dirs.values() if x == "PE")
    ce_votes = sum(1 for x in dirs.values() if x == "CE")
    master = 0.25 * dom["minus100"] + 0.50 * dom["atm"] + 0.25 * dom["plus100"]

    cio_rows = state.get("series", {}).get("cio", [])
    cio = cio_rows[-1] if cio_rows else {}
    cio_ce = float(cio.get("ce", 0) or 0)
    cio_pe = float(cio.get("pe", 0) or 0)
    cio_dir = "PE" if cio_ce < cio_pe else ("CE" if cio_pe < cio_ce else None)

    return {
        "time": minute_label(), "deltas": deltas, "directions": dirs,
        "pe_votes": pe_votes, "ce_votes": ce_votes,
        "master_d": round(master, 2), "cio_ce": int(cio_ce), "cio_pe": int(cio_pe),
        "cio_direction": cio_dir,
    }


def _recent_direction_count(strategy_name, direction):
    runtime = state.setdefault("strategy_engine", {}).setdefault("recent", {})
    key = f"{strategy_name}:{direction}"
    return int(runtime.get(key, 0))


def _set_direction_count(strategy_name, direction, value):
    runtime = state.setdefault("strategy_engine", {}).setdefault("recent", {})
    runtime[f"{strategy_name}:{direction}"] = int(value)


def _update_persistence(strategy_name, direction, condition):
    other = "CE" if direction == "PE" else "PE"
    if condition:
        _set_direction_count(strategy_name, direction, _recent_direction_count(strategy_name, direction) + 1)
    else:
        _set_direction_count(strategy_name, direction, 0)
    if condition:
        _set_direction_count(strategy_name, other, 0)
    return _recent_direction_count(strategy_name, direction)


def _signal_record(strategy_name, action, option_type, price, reason, metrics):
    return {
        "time": minute_label(), "timestamp": now_ist().isoformat(), "strategy": strategy_name,
        "action": action, "type": option_type, "nifty_level": round(float(price), 2),
        "reason": reason, "master_d": metrics.get("master_d"),
        "pe_votes": metrics.get("pe_votes"), "ce_votes": metrics.get("ce_votes"),
        "cio_direction": metrics.get("cio_direction"),
    }


def _enter_strategy(strategy_name, option_type, price, reason, metrics):
    s = state["strategies"][strategy_name]
    if s.get("active") or len(s.get("trades", [])) >= STRATEGY_MAX_TRADES_PER_STRATEGY:
        return
    p = float(price)
    sl = p - STRATEGY_HARD_SL_POINTS if option_type == "PE" else p + STRATEGY_HARD_SL_POINTS
    trade = {
        "date": today_key(), "type": option_type, "side": f"SELL {option_type}",
        "entry_time": minute_label(), "entry_timestamp": now_ist().isoformat(),
        "entry_level": round(p, 2), "sl_level": round(sl, 2), "exit_time": None,
        "exit_level": None, "points": None, "mae": 0.0, "mfe": 0.0,
        "sl_hit": False, "result": "OPEN", "exit_reason": None,
        "entry_reason": reason, "master_d_entry": metrics.get("master_d"),
    }
    s["active"] = trade
    s.setdefault("signals", []).append(_signal_record(strategy_name, "ENTRY", option_type, p, reason, metrics))


def _update_excursions(trade, price):
    p, e = float(price), float(trade["entry_level"])
    if trade.get("type") == "PE":
        favorable, adverse = max(0.0, p - e), max(0.0, e - p)
    else:
        favorable, adverse = max(0.0, e - p), max(0.0, p - e)
    trade["mfe"] = round(max(float(trade.get("mfe") or 0), favorable), 2)
    trade["mae"] = round(max(float(trade.get("mae") or 0), adverse), 2)


def _exit_strategy(strategy_name, price, reason, metrics, sl_hit=False):
    s = state["strategies"][strategy_name]
    trade = s.get("active")
    if not trade:
        return
    p, e = float(price), float(trade["entry_level"])
    points = p - e if trade.get("type") == "PE" else e - p
    trade.update({
        "exit_time": minute_label(), "exit_timestamp": now_ist().isoformat(),
        "exit_level": round(p, 2), "points": round(points, 2), "sl_hit": bool(sl_hit),
        "result": "WIN" if points > 0 else ("LOSS" if points < 0 else "FLAT"),
        "exit_reason": reason,
    })
    s.setdefault("trades", []).append(dict(trade))
    s.setdefault("signals", []).append(_signal_record(strategy_name, "EXIT", trade.get("type"), p, reason, metrics))
    s["active"] = None
    _set_direction_count(strategy_name, "PE", 0)
    _set_direction_count(strategy_name, "CE", 0)


def evaluate_live_strategies():
    """Evaluate once per completed/snapshotted minute from the recorded OIC+CIO series."""
    price = state.get("nifty", {}).get("price")
    if price is None or not baseline_ready:
        return
    now_m = _row_minutes(minute_label())
    if now_m is None:
        return
    engine = state.setdefault("strategy_engine", {})
    if engine.get("last_eval_minute") == minute_label():
        return
    engine["last_eval_minute"] = minute_label()

    metrics = _strategy_metrics()
    engine["metrics"] = metrics
    p = float(price)

    # First manage any active positions. Hard SL has priority.
    for name in ("S1", "S2", "PNA"):
        trade = state["strategies"][name].get("active")
        if not trade:
            continue
        _update_excursions(trade, p)
        if (trade["type"] == "PE" and p <= float(trade["sl_level"])) or (trade["type"] == "CE" and p >= float(trade["sl_level"])):
            _exit_strategy(name, p, "30-POINT HARD SL", metrics, True)
            continue
        if now_m >= 14 * 60 + 45:
            _exit_strategy(name, p, "14:45 TIME EXIT", metrics, False)
            continue

        opposite = "CE" if trade["type"] == "PE" else "PE"
        votes = metrics["ce_votes"] if opposite == "CE" else metrics["pe_votes"]
        cio_ok = metrics["cio_direction"] == opposite
        if name == "S1":
            rev_count = _update_persistence("S1_EXIT", opposite, votes >= 2)
            # Locked S1 exit: 3-reading OIC reversal OR CIO no longer supports current trade.
            if rev_count >= 3 or metrics["cio_direction"] not in (trade["type"], None):
                _exit_strategy(name, p, "OIC/CIO REVERSAL", metrics, False)
        elif name == "S2":
            if votes >= 2 and cio_ok:
                _exit_strategy(name, p, "MOMENTUM REVERSAL", metrics, False)
        else:  # PNA
            master_opposite = metrics["master_d"] >= STRATEGY_MASTER_CLEAR if opposite == "CE" else metrics["master_d"] <= -STRATEGY_MASTER_CLEAR
            if votes >= 2 and cio_ok and master_opposite:
                _exit_strategy(name, p, "MASTER OIC/CIO REVERSAL", metrics, False)

    # Observation window: collect data but no new entries until after 09:30.
    if now_m < 9 * 60 + 30 or now_m >= 14 * 60 + 45:
        return

    for direction in ("PE", "CE"):
        votes = metrics["pe_votes"] if direction == "PE" else metrics["ce_votes"]
        cio_ok = metrics["cio_direction"] == direction
        common = votes >= 2 and cio_ok

        # S2: earliest visible >=2/3 momentum shift with CIO support.
        if common and not state["strategies"]["S2"].get("active"):
            _enter_strategy("S2", direction, p, f"Momentum shift {votes}/3 + CIO", metrics)

        # S1: same clear dominance must persist for 3 consecutive 1-minute readings.
        s1_count = _update_persistence("S1", direction, common)
        if s1_count >= 3 and not state["strategies"]["S1"].get("active"):
            _enter_strategy("S1", direction, p, f"OIC dominance {votes}/3 + CIO, 3 readings", metrics)

        # PNA: multi-layer confirmation with weighted MasterD clear + persistence.
        master_ok = metrics["master_d"] <= -STRATEGY_MASTER_CLEAR if direction == "PE" else metrics["master_d"] >= STRATEGY_MASTER_CLEAR
        pna_cond = common and master_ok
        pna_count = _update_persistence("PNA", direction, pna_cond)
        if pna_count >= 3 and not state["strategies"]["PNA"].get("active"):
            _enter_strategy("PNA", direction, p, f"{votes}/3 + MasterD {metrics['master_d']:+.1f} + CIO, 3 readings", metrics)

def make_snapshot():
    if not state.get("connected"):
        return

    with lock:
        # Persist the true 1-minute NIFTY OHLC candle assembled from ticks.
        # ``price`` remains an alias of ``close`` so the current frontend
        # continues to work without any HTML/JavaScript change.
        snapshot_current_nifty_candle()
        snapshot_current_premium_candles()

        for key in (
            "atm",
            "minus100",
            "plus100",
        ):
            point = oic_point(key)

            if point:
                append_or_replace_minute(
                    state["series"][key],
                    point,
                )

        if baseline_ready:
            ce, pe = cio_totals()

            point = {
                "time": minute_label(),
                "timestamp": now_ist().isoformat(),
                "ce": int(ce),
                "pe": int(pe),
            }

            append_or_replace_minute(
                state["series"]["cio"],
                point,
            )

        # Generate live S1/S2/PNA entries/exits from the same minute snapshot.
        evaluate_live_strategies()

        state["last_update"] = (
            now_ist().isoformat()
        )

    save_current_history()


def snapshot_worker():
    print(
        "[DIAG] Snapshot worker started",
        flush=True,
    )

    while True:
        try:
            if (
                state.get("connected")
                and is_market_session()
            ):
                make_snapshot()

        except Exception as e:
            print(
                f"[DIAG] Snapshot ERROR: {e}",
                flush=True,
            )

        n = now_ist()

        seconds_to_next = (
            60 - n.second
        )

        if seconds_to_next < 2:
            seconds_to_next = 2

        time.sleep(seconds_to_next)


def ensure_snapshot_worker():
    global snapshot_thread_started

    if snapshot_thread_started:
        return

    snapshot_thread_started = True

    threading.Thread(
        target=snapshot_worker,
        daemon=True,
    ).start()


# ============================================================
# WEBSOCKET
# ============================================================

def on_ticks(ws, ticks):
    global latest_nifty
    global nifty_minute_bucket
    global premium_minute_buckets
    global latest_vix
    global tick_counter
    global oi_tick_counter

    tick_counter += len(ticks)

    option_oi_in_batch = 0

    for q in ticks:
        token = int(
            q.get(
                "instrument_token",
                0,
            )
        )

        # NIFTY
        if token == nifty_token:
            price = q.get(
                "last_price"
            )

            ohlc = q.get(
                "ohlc"
            ) or {}

            open_price = ohlc.get(
                "open"
            )

            high = ohlc.get(
                "high"
            )

            low = ohlc.get(
                "low"
            )

            previous_close = ohlc.get(
                "close"
            )

            if price is not None:
                latest_nifty[
                    "price"
                ] = float(price)

                # Build true minute OHLC from each live NIFTY tick.
                update_nifty_minute_ohlc(price)

            if open_price is not None:
                latest_nifty[
                    "open"
                ] = float(open_price)

            if high is not None:
                latest_nifty[
                    "high"
                ] = float(high)

            if low is not None:
                latest_nifty[
                    "low"
                ] = float(low)

            if previous_close is not None:
                latest_nifty[
                    "previous_close"
                ] = float(
                    previous_close
                )

            p = latest_nifty.get(
                "price"
            )

            pc = latest_nifty.get(
                "previous_close"
            )

            op = latest_nifty.get(
                "open"
            )

            if p is not None and pc:
                latest_nifty[
                    "change"
                ] = round(
                    p - pc,
                    2,
                )

                latest_nifty[
                    "change_pct"
                ] = round(
                    ((p - pc) / pc) * 100,
                    3,
                )

            with lock:
                state["nifty"] = dict(
                    latest_nifty
                )

            # Self-heal locked OIC/premium mappings if the process restarted
            # or mapping initialization was missed.
            if len(oic_tokens) != 3 or len(premium_token_map) != 6:
                ensure_locked_strike_mappings()

            if (
                state.get("opening_atm")
                is None
                and op is not None
            ):
                atm = round_to_100(op)

                with lock:
                    state["opening_atm"] = atm

                    state[
                        "zone"
                    ] = calculate_zone(
                        op,
                        pc,
                    )

                lock_oic_strikes()

            elif op is not None and pc:
                with lock:
                    state[
                        "zone"
                    ] = calculate_zone(
                        op,
                        pc,
                    )

        # VIX
        elif token == vix_token:
            price = q.get(
                "last_price"
            )

            if price is not None:
                value = float(price)

                band, interpretation = (
                    classify_vix(value)
                )

                latest_vix = {
                    "price": value,
                    "range": band,
                    "interpretation": interpretation,
                }

                with lock:
                    state["vix"] = dict(
                        latest_vix
                    )

        # OPTION PREMIUM (six locked strikes only)
        if token in premium_token_map:
            premium_price = q.get("last_price")
            premium_oi = q.get("oi")
            if premium_price is not None:
                update_premium_minute_ohlc(token, premium_price, premium_oi)

        # OPTION OI
        if token in token_meta:
            oi = q.get("oi")

            if oi is not None:
                latest_oi[
                    token
                ] = int(oi)

                oi_tick_counter += 1
                option_oi_in_batch += 1

    if option_oi_in_batch:
        print(
            f"[DIAG] OI batch={option_oi_in_batch}, unique={len(latest_oi)}",
            flush=True,
        )

    with lock:
        state["last_update"] = (
            now_ist().isoformat()
        )


def on_connect(ws, response):
    tokens = []

    if nifty_token:
        tokens.append(nifty_token)

    if vix_token:
        tokens.append(vix_token)

    tokens.extend(
        token_meta.keys()
    )

    tokens = list(set(tokens))

    print(
        f"[DIAG] WebSocket connected. Subscribing {len(tokens)} tokens.",
        flush=True,
    )

    if tokens:
        ws.subscribe(tokens)

        ws.set_mode(
            ws.MODE_FULL,
            tokens,
        )

    with lock:
        state["connected"] = True
        state["message"] = (
            "LIVE — Zerodha connected"
        )


def on_close(ws, code, reason):
    print(
        f"[DIAG] WebSocket closed {code} {reason}",
        flush=True,
    )

    with lock:
        state["connected"] = False
        state["message"] = (
            f"Disconnected — {reason or code}"
        )


def on_error(ws, code, reason):
    print(
        f"[DIAG] WebSocket ERROR {code} {reason}",
        flush=True,
    )

    with lock:
        state["message"] = (
            f"WebSocket error ({code}) — {reason}"
        )


# ============================================================
# START LIVE
# ============================================================

def start_live(access_token):
    global kite
    global ticker
    global baseline_ready

    # Prevent the index route, startup restore and watchdog from creating
    # competing KiteTicker instances at the same time.
    with live_start_mutex:
        print(
            "[DIAG] start_live() called",
            flush=True,
        )

        with lock:
            state["date"] = today_key()
            state["message"] = (
                "Starting Zerodha live feed..."
            )

        # Dispose of a stale ticker object before creating a fresh one.
        old_ticker = ticker
        ticker = None
        if old_ticker is not None:
            try:
                old_ticker.close()
            except Exception:
                pass

        kite = KiteConnect(
            api_key=KITE_API_KEY
        )

        kite.set_access_token(
            access_token
        )

        # Synchronous API validation.  If the saved daily token is invalid,
        # this raises immediately; transient WebSocket failures are handled
        # separately by the watchdog and do not erase the token.
        kite.profile()

        discover_instruments(kite)

    # Rebuild the three locked OIC + six premium mappings immediately.
    # The fail-safe also reconstructs opening ATM from the live NIFTY quote
    # after a Render restart when today's history did not contain it yet.
    ensure_locked_strike_mappings(kite)

    baseline_ready = False

    if not baseline_thread_started:
        threading.Thread(
            target=build_oi_baseline,
            daemon=True,
        ).start()

    ticker = KiteTicker(
        KITE_API_KEY,
        access_token,
    )

    ticker.on_ticks = on_ticks
    ticker.on_connect = on_connect
    ticker.on_close = on_close
    ticker.on_error = on_error

    ticker.connect(
        threaded=True
    )

    ensure_snapshot_worker()
    ensure_reconnect_watchdog()


def reconnect_watchdog():
    """Keep the server-side Kite stream alive during the market session.

    This is intentionally independent of the browser.  It uses the access
    token already persisted for today in Neon and retries a disconnected
    stream at a conservative interval.
    """
    global reconnect_in_progress
    global last_reconnect_attempt

    print("[KITE] Reconnect watchdog started", flush=True)

    while True:
        try:
            if is_market_session() and not state.get("connected"):
                now_ts = time.time()
                if (
                    not reconnect_in_progress
                    and now_ts - last_reconnect_attempt >= 15
                ):
                    token = load_access_token()
                    if token:
                        reconnect_in_progress = True
                        last_reconnect_attempt = now_ts
                        try:
                            print(
                                "[KITE] Feed disconnected; attempting automatic reconnect...",
                                flush=True,
                            )
                            start_live(token)
                        except Exception as e:
                            # Do not delete today's token for a transient
                            # WebSocket/network failure.  A genuinely invalid
                            # token will keep failing profile() and can be
                            # replaced by the normal Zerodha login flow.
                            print(
                                f"[KITE] Automatic reconnect failed: {e}",
                                flush=True,
                            )
                            with lock:
                                state["connected"] = False
                                state["message"] = "Reconnecting Zerodha live feed..."
                        finally:
                            reconnect_in_progress = False
        except Exception as e:
            print(f"[KITE] Watchdog error: {e}", flush=True)

        time.sleep(5)


def ensure_reconnect_watchdog():
    global reconnect_watchdog_started
    if reconnect_watchdog_started:
        return
    reconnect_watchdog_started = True
    threading.Thread(
        target=reconnect_watchdog,
        daemon=True,
        name="kite-reconnect-watchdog",
    ).start()


# ============================================================
# RESTORE TODAY
# ============================================================

def restore_today_history():
    day = today_key()

    saved = load_day_history(day)

    series = saved.get(
        "series"
    ) or {}

    with lock:
        state["date"] = day

        for key in (
            "nifty",
            "atm",
            "minus100",
            "plus100",
            "cio",
        ):
            if isinstance(
                series.get(key),
                list,
            ):
                state[
                    "series"
                ][key] = series[key]

        premium_series = series.get("premium") or {}
        for premium_key in ("minus100", "atm", "plus100"):
            legs = premium_series.get(premium_key) or {}
            for option_type in ("CE", "PE"):
                rows = legs.get(option_type) or []
                if isinstance(rows, list):
                    state["series"]["premium"][premium_key][option_type] = rows

        saved_premium = saved.get("premium") or {}
        for premium_key in ("minus100", "atm", "plus100"):
            if saved_premium.get(premium_key) is not None:
                state["premium"][premium_key] = saved_premium.get(premium_key)

        if saved.get(
            "opening_atm"
        ):
            state[
                "opening_atm"
            ] = saved[
                "opening_atm"
            ]

        if saved.get("expiry"):
            state["expiry"] = saved[
                "expiry"
            ]

        if saved.get("zone"):
            state["zone"] = saved[
                "zone"
            ]

        if saved.get("nifty"):
            state["nifty"] = saved[
                "nifty"
            ]

        if saved.get("vix"):
            state["vix"] = saved[
                "vix"
            ]

        saved_strategies = saved.get("strategies") or {}
        for strategy_name in ("S1", "S2", "PNA"):
            strategy_data = saved_strategies.get(strategy_name) or {}
            trades = strategy_data.get("trades") or []
            if isinstance(trades, list):
                state["strategies"][strategy_name]["trades"] = trades
            state["strategies"][strategy_name]["active"] = strategy_data.get("active")
            signals = strategy_data.get("signals") or []
            if isinstance(signals, list):
                state["strategies"][strategy_name]["signals"] = signals
        saved_engine = saved.get("strategy_engine") or {}
        if isinstance(saved_engine, dict):
            state["strategy_engine"].update(saved_engine)


init_db()
restore_today_history()
refresh_history_dates()
# Important for free Render: when GitHub Actions wakes/restarts the service,
# reconnect to Kite from today's Neon-persisted token without requiring the
# dashboard to be opened in a browser.
restore_kite_session()
ensure_snapshot_worker()
ensure_reconnect_watchdog()


# ============================================================
# ROUTES
# ============================================================

# ============================================================
# PORTAL AUTHENTICATION
# ============================================================

def portal_login_configured():
    return bool(PORTAL_USERNAME and PORTAL_PASSWORD)


def portal_authenticated():
    return bool(session.get("portal_authenticated"))


@app.before_request
def require_portal_login():
    # Login page, static assets and health check stay public.
    if request.endpoint in {"portal_login", "static", "health"}:
        return None

    if not portal_authenticated():
        next_url = request.full_path if request.query_string else request.path
        return redirect(url_for("portal_login", next=next_url))

    return None


@app.route("/login", methods=["GET", "POST"])
def portal_login():
    if portal_authenticated():
        return redirect(url_for("index"))

    error = None
    configured = portal_login_configured()

    if request.method == "POST":
        if not configured:
            error = "Portal login is not configured on the server."
        else:
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")

            user_ok = hmac.compare_digest(username, PORTAL_USERNAME)
            pass_ok = hmac.compare_digest(password, PORTAL_PASSWORD)

            if user_ok and pass_ok:
                session.clear()
                session["portal_authenticated"] = True
                session["portal_username"] = PORTAL_USERNAME
                session.permanent = request.form.get("remember") == "on"

                next_url = request.args.get("next", "")
                if not next_url.startswith("/") or next_url.startswith("//"):
                    next_url = url_for("index")
                return redirect(next_url)

            error = "Invalid username or password."

    return render_template(
        "login.html",
        error=error,
        configured=configured,
    )


@app.route("/logout")
def portal_logout():
    session.clear()
    return redirect(url_for("portal_login"))


@app.route("/")
def index():
    if (
        not KITE_API_KEY
        or not KITE_API_SECRET
    ):
        return render_template(
            "index.html",
            configured=False,
            base_url=PUBLIC_BASE_URL,
        )

    token = load_access_token()

    if token and not state["connected"]:
        try:
            start_live(token)

        except Exception as e:
            print(
                f"[DIAG] Existing token failed: {e}",
                flush=True,
            )

            clear_access_token()

            with lock:
                state["connected"] = False
                state["message"] = (
                    "Login required"
                )

    return render_template(
        "index.html",
        configured=True,
        base_url=PUBLIC_BASE_URL,
    )


@app.route("/kite/login")
def kite_login():
    if not KITE_API_KEY:
        return (
            "KITE_API_KEY is not configured",
            400,
        )

    k = KiteConnect(
        api_key=KITE_API_KEY
    )

    return redirect(
        k.login_url()
    )


@app.route("/kite/callback")
def kite_callback():
    request_token = request.args.get(
        "request_token"
    )

    if not request_token:
        return (
            "Zerodha did not return a request_token",
            400,
        )

    k = KiteConnect(
        api_key=KITE_API_KEY
    )

    session = k.generate_session(
        request_token,
        api_secret=KITE_API_SECRET,
    )

    access_token = session[
        "access_token"
    ]

    save_access_token(
        access_token
    )

    start_live(
        access_token
    )

    return redirect("/")


@app.route("/kite/logout")
def kite_logout():
    global ticker

    try:
        if ticker:
            ticker.close()

    except Exception as e:
        print(
            f"[DIAG] Ticker close error: {e}",
            flush=True,
        )

    clear_access_token()

    with lock:
        state["connected"] = False
        state["message"] = (
            "Logged out — Zerodha login required"
        )

    return redirect("/")


@app.route("/api/state")
def api_state():
    with lock:
        return jsonify(state)


@app.route("/api/history/dates")
def api_history_dates():
    refresh_history_dates()

    with lock:
        return jsonify(
            {
                "dates": state[
                    "history_dates"
                ]
            }
        )


@app.route("/api/history/<day>")
def api_history_day(day):
    if not history_exists(day):
        return jsonify(
            {
                "error": (
                    "No stored data for selected date."
                )
            }
        ), 404

    return jsonify(
        load_day_history(day)
    )


@app.route("/api/history/<day>/cio")
def api_history_cio(day):
    # For today use current live memory.
    if day == today_key():
        with lock:
            return jsonify(
                {
                    "date": day,
                    "expiry": state.get(
                        "expiry"
                    ),
                    "opening_atm": state.get(
                        "opening_atm"
                    ),
                    "cio": list(
                        state.get(
                            "series",
                            {},
                        ).get(
                            "cio",
                            [],
                        )
                    ),
                }
            )

    if not history_exists(day):
        return jsonify(
            {
                "error": (
                    "No stored data for selected date."
                )
            }
        ), 404

    data = load_day_history(day)

    return jsonify(
        {
            "date": day,
            "expiry": data.get(
                "expiry"
            ),
            "opening_atm": data.get(
                "opening_atm"
            ),
            "cio": data.get(
                "series",
                {},
            ).get(
                "cio",
                [],
            ),
        }
    )


@app.route("/api/premium/<day>")
def api_premium_day(day):
    """Return the six recorded option-premium OHLC series for a session."""
    if day == today_key():
        with lock:
            snapshot_current_premium_candles()
            return jsonify({
                "date": day,
                "expiry": state.get("expiry"),
                "opening_atm": state.get("opening_atm"),
                "premium": state.get("premium", {}),
                "series": state.get("series", {}).get("premium", {}),
            })

    if not history_exists(day):
        return jsonify({"error": "No stored data for selected date."}), 404

    data = load_day_history(day)
    return jsonify({
        "date": day,
        "expiry": data.get("expiry"),
        "opening_atm": data.get("opening_atm"),
        "premium": data.get("premium", {}),
        "series": (data.get("series", {}) or {}).get("premium", {}),
    })


# ============================================================
# OIC + CIO + NIFTY EXCEL DOWNLOAD
# Existing route is preserved for frontend compatibility.
# ============================================================

@app.route("/api/download/cio/<day>")
def download_cio_excel(day):
    if day == today_key():
        with lock:
            data = {
                "date": day,
                "expiry": state.get("expiry"),
                "opening_atm": state.get("opening_atm"),
                "series": {
                    key: list(state.get("series", {}).get(key, []))
                    for key in ("nifty", "minus100", "atm", "plus100", "cio")
                },
            }
    else:
        if not history_exists(day):
            return jsonify({"error": "No stored data for selected date."}), 404
        data = load_day_history(day)

    series = data.get("series", {}) or {}
    nifty_data = series.get("nifty", []) or []
    minus100 = series.get("minus100", []) or []
    atm = series.get("atm", []) or []
    plus100 = series.get("plus100", []) or []
    cio = series.get("cio", []) or []

    if not any((nifty_data, minus100, atm, plus100, cio)):
        return jsonify({"error": "No OIC/CIO/NIFTY data available for selected date."}), 404

    def by_time(rows):
        return {str(p.get("time")): p for p in rows if p.get("time")}

    maps = {
        "nifty": by_time(nifty_data),
        "minus100": by_time(minus100),
        "atm": by_time(atm),
        "plus100": by_time(plus100),
        "cio": by_time(cio),
    }
    all_times = sorted(set().union(*(m.keys() for m in maps.values())))

    wb = Workbook()
    ws = wb.active
    ws.title = "OIC + CIO + NIFTY"

    ws["A1"] = "Pratik Analysis"
    ws["A2"] = "NIFTY 1-Min OHLC + OIC + CIO — Minute-wise Data"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"].font = Font(bold=True, size=13)

    ws["A4"] = "Date"
    ws["B4"] = day
    ws["A5"] = "NIFTY Opening ATM"
    ws["B5"] = data.get("opening_atm")
    ws["A6"] = "Nearest Expiry"
    ws["B6"] = data.get("expiry")

    headers = [
        "Time",
        "NIFTY Open",
        "NIFTY High",
        "NIFTY Low",
        "NIFTY Close",
        "ATM -100 CE OI",
        "ATM -100 PE OI",
        "ATM CE OI",
        "ATM PE OI",
        "ATM +100 CE OI",
        "ATM +100 PE OI",
        "CIO CE Negative Change in OI",
        "CIO PE Negative Change in OI",
    ]
    header_row = 8
    for col, value in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col, value=value)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for row_no, t in enumerate(all_times, start=header_row + 1):
        n = maps["nifty"].get(t, {})
        m = maps["minus100"].get(t, {})
        a = maps["atm"].get(t, {})
        p = maps["plus100"].get(t, {})
        c = maps["cio"].get(t, {})
        # Older stored days may contain only ``price``.  For those rows,
        # fall back to that value so historical downloads remain readable.
        fallback_price = n.get("price")
        values = [
            t,
            n.get("open", fallback_price),
            n.get("high", fallback_price),
            n.get("low", fallback_price),
            n.get("close", fallback_price),
            m.get("ce"), m.get("pe"),
            a.get("ce"), a.get("pe"),
            p.get("ce"), p.get("pe"),
            c.get("ce"), c.get("pe"),
        ]
        for col, value in enumerate(values, 1):
            ws.cell(row=row_no, column=col, value=value)

    ws.freeze_panes = "A9"
    widths = [14, 16, 16, 16, 16, 20, 20, 20, 20, 20, 20, 32, 32]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64+i)].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Pratik_Analysis_NIFTY_OIC_CIO_{day}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# STRATEGY REPORTS + EXCEL EXPORT
# ============================================================

def _parse_report_day(value):
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except Exception:
        return None


def _strategy_trades_from_day(day, data):
    strategies = (data or {}).get("strategies") or {}
    rows = []

    for strategy_name in ("S1", "S2", "PNA"):
        strategy_data = strategies.get(strategy_name) or {}
        trades = strategy_data.get("trades") or []

        if not isinstance(trades, list):
            continue

        for i, trade in enumerate(trades, start=1):
            if not isinstance(trade, dict):
                continue

            row = dict(trade)
            row.setdefault("date", day)
            row.setdefault("strategy", strategy_name)
            row.setdefault("trade_no", i)
            rows.append(row)

    return rows


def _report_days(start_day, end_day):
    refresh_history_dates()
    with lock:
        available = list(state.get("history_dates", []))
        today_has_data = bool(state.get("date") == today_key())

    if today_has_data and today_key() not in available:
        available.append(today_key())

    selected = []
    for day in available:
        d = _parse_report_day(day)
        if d and start_day <= d <= end_day:
            selected.append(day)

    return sorted(set(selected))


def _collect_report(start_text, end_text):
    start_day = _parse_report_day(start_text)
    end_day = _parse_report_day(end_text)

    if not start_day or not end_day:
        return None, "Please select a valid From and To date."

    if start_day > end_day:
        return None, "From date cannot be after To date."

    days = _report_days(start_day, end_day)
    all_trades = []

    for day in days:
        if day == today_key():
            with lock:
                data = {
                    "date": day,
                    "strategies": json.loads(json.dumps(state.get("strategies", {}))),
                }
        else:
            data = load_day_history(day)

        all_trades.extend(_strategy_trades_from_day(day, data))

    def num(value):
        try:
            return float(value)
        except Exception:
            return 0.0

    summary = {}
    for name in ("S1", "S2", "PNA"):
        trades = [t for t in all_trades if str(t.get("strategy", "")).upper() == name]
        completed = [t for t in trades if t.get("exit_time") or t.get("exit_level") is not None]
        points = [num(t.get("points")) for t in completed]
        wins = sum(1 for p in points if p > 0)
        losses = sum(1 for p in points if p < 0)
        flat = sum(1 for p in points if p == 0)
        sl_hits = sum(1 for t in completed if bool(t.get("sl_hit")))
        maes = [num(t.get("mae")) for t in completed if t.get("mae") is not None]
        mfes = [num(t.get("mfe")) for t in completed if t.get("mfe") is not None]

        summary[name] = {
            "trades": len(completed),
            "wins": wins,
            "losses": losses,
            "flat": flat,
            "win_rate": round((wins / len(completed) * 100), 2) if completed else 0.0,
            "total_points": round(sum(points), 2),
            "avg_points": round((sum(points) / len(completed)), 2) if completed else 0.0,
            "sl_hits": sl_hits,
            "avg_mae": round((sum(maes) / len(maes)), 2) if maes else 0.0,
            "avg_mfe": round((sum(mfes) / len(mfes)), 2) if mfes else 0.0,
            "best_trade": round(max(points), 2) if points else 0.0,
            "worst_trade": round(min(points), 2) if points else 0.0,
        }

    daywise = []
    for day in days:
        item = {"date": day}
        for name in ("S1", "S2", "PNA"):
            day_trades = [
                t for t in all_trades
                if t.get("date") == day
                and str(t.get("strategy", "")).upper() == name
                and (t.get("exit_time") or t.get("exit_level") is not None)
            ]
            item[f"{name}_trades"] = len(day_trades)
            item[f"{name}_points"] = round(sum(num(t.get("points")) for t in day_trades), 2)
        daywise.append(item)

    return {
        "from": start_text,
        "to": end_text,
        "days": days,
        "summary": summary,
        "daywise": daywise,
        "trades": all_trades,
    }, None


@app.route("/api/reports")
def api_strategy_reports():
    start_text = request.args.get("from", "")
    end_text = request.args.get("to", "")
    report, error = _collect_report(start_text, end_text)
    if error:
        return jsonify({"error": error}), 400
    return jsonify(report)


@app.route("/api/download/reports")
def download_strategy_reports_excel():
    start_text = request.args.get("from", "")
    end_text = request.args.get("to", "")
    report, error = _collect_report(start_text, end_text)
    if error:
        return jsonify({"error": error}), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "Overall Summary"

    title_font = Font(bold=True, size=16)
    header_font = Font(bold=True)

    ws["A1"] = "Pratik Analysis"
    ws["A2"] = "S1 / S2 / PNA Strategy Performance Report"
    ws["A1"].font = title_font
    ws["A2"].font = Font(bold=True, size=13)
    ws["A4"] = "From"
    ws["B4"] = start_text
    ws["C4"] = "To"
    ws["D4"] = end_text

    summary_headers = [
        "Strategy", "Trades", "Wins", "Losses", "Flat", "Win %",
        "Total Points", "Avg Points/Trade", "30-Point SL Hits",
        "Avg MAE", "Avg MFE", "Best Trade", "Worst Trade",
    ]
    for col, value in enumerate(summary_headers, 1):
        cell = ws.cell(row=6, column=col, value=value)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_no, name in enumerate(("S1", "S2", "PNA"), start=7):
        m = report["summary"][name]
        values = [
            name, m["trades"], m["wins"], m["losses"], m["flat"], m["win_rate"],
            m["total_points"], m["avg_points"], m["sl_hits"], m["avg_mae"],
            m["avg_mfe"], m["best_trade"], m["worst_trade"],
        ]
        for col, value in enumerate(values, 1):
            ws.cell(row=row_no, column=col, value=value)

    ws_day = wb.create_sheet("Day-wise Summary")
    day_headers = ["Date", "S1 Trades", "S1 Points", "S2 Trades", "S2 Points", "PNA Trades", "PNA Points"]
    for col, value in enumerate(day_headers, 1):
        c = ws_day.cell(row=1, column=col, value=value)
        c.font = header_font
    for row_no, d in enumerate(report["daywise"], start=2):
        vals = [
            d["date"], d["S1_trades"], d["S1_points"], d["S2_trades"],
            d["S2_points"], d["PNA_trades"], d["PNA_points"],
        ]
        for col, value in enumerate(vals, 1):
            ws_day.cell(row=row_no, column=col, value=value)

    trade_headers = [
        "Date", "Trade #", "Type", "Entry Time", "Entry Level", "30-Point SL",
        "Exit Time", "Exit Level", "Points", "MAE", "MFE", "SL Hit?",
        "Result", "Exit Reason", "Comments",
    ]

    for strategy_name in ("S1", "S2", "PNA"):
        sheet = wb.create_sheet(f"{strategy_name} Trades")
        for col, value in enumerate(trade_headers, 1):
            c = sheet.cell(row=1, column=col, value=value)
            c.font = header_font

        rows = [t for t in report["trades"] if str(t.get("strategy", "")).upper() == strategy_name]
        for row_no, t in enumerate(rows, start=2):
            points = t.get("points")
            result = t.get("result")
            if not result and points is not None:
                try:
                    p = float(points)
                    result = "WIN" if p > 0 else "LOSS" if p < 0 else "FLAT"
                except Exception:
                    result = ""

            vals = [
                t.get("date"), t.get("trade_no"), t.get("type") or t.get("side"),
                t.get("entry_time"), t.get("entry_level"), t.get("sl_level"),
                t.get("exit_time"), t.get("exit_level"), points, t.get("mae"),
                t.get("mfe"), "YES" if t.get("sl_hit") else "NO", result,
                t.get("exit_reason"), t.get("comments"),
            ]
            for col, value in enumerate(vals, 1):
                sheet.cell(row=row_no, column=col, value=value)

    risk = wb.create_sheet("Risk & SL")
    risk_headers = ["Strategy", "Trades", "30-Point SL Hits", "Avg MAE", "Avg MFE", "Best Trade", "Worst Trade"]
    for col, value in enumerate(risk_headers, 1):
        c = risk.cell(row=1, column=col, value=value)
        c.font = header_font
    for row_no, name in enumerate(("S1", "S2", "PNA"), start=2):
        m = report["summary"][name]
        vals = [name, m["trades"], m["sl_hits"], m["avg_mae"], m["avg_mfe"], m["best_trade"], m["worst_trade"]]
        for col, value in enumerate(vals, 1):
            risk.cell(row=row_no, column=col, value=value)

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2" if sheet.title != "Overall Summary" else "A6"
        for column_cells in sheet.columns:
            max_len = 0
            letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            sheet.column_dimensions[letter].width = min(max(max_len + 2, 11), 30)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"Pratik_Analysis_Strategy_Report_{start_text}_to_{end_text}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )



# ============================================================
# OPTION PREMIUM PERFORMANCE + EXCEL EXPORT
# ============================================================

PREMIUM_KEYS = ("minus100", "atm", "plus100")
PREMIUM_LABELS = {
    "minus100": "ATM -100",
    "atm": "ATM",
    "plus100": "ATM +100",
}


def _normalise_hhmm(value):
    """Return HH:MM from common stored time formats."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    # ISO datetime
    if "T" in text:
        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00")).strftime("%H:%M")
        except Exception:
            pass
    # HH:MM[:SS] or labels containing a time
    import re
    m = re.search(r"(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?", text)
    if m:
        try:
            hh = int(m.group(1)); mm = int(m.group(2))
            if 0 <= hh <= 23 and 0 <= mm <= 59:
                return f"{hh:02d}:{mm:02d}"
        except Exception:
            pass
    return None


def _trade_option_type(trade):
    """Map strategy trade direction to the sold option leg (CE or PE)."""
    values = [
        trade.get("type"), trade.get("side"), trade.get("signal"),
        trade.get("option_type"), trade.get("direction"),
    ]
    text = " ".join(str(v or "") for v in values).upper()
    if "PE" in text:
        return "PE"
    if "CE" in text:
        return "CE"
    return None


def _series_for_premium(data, key, option_type):
    return ((((data or {}).get("series") or {}).get("premium") or {}).get(key) or {}).get(option_type) or []


def _candle_by_minute(series):
    result = {}
    for p in series or []:
        if not isinstance(p, dict):
            continue
        hhmm = _normalise_hhmm(p.get("time") or p.get("timestamp"))
        if hhmm:
            result[hhmm] = p
    return result


def _premium_trade_metrics(trade, day_data, key, option_type):
    """Calculate sold-option performance from stored 1-minute OHLC.

    Entry/exit use the candle close at the strategy entry/exit minute. This is the
    most precise reproducible value available from the stored 1-minute premium
    history. MAE/MFE use minute highs/lows while the strategy trade is open.
    """
    series = _series_for_premium(day_data, key, option_type)
    if not series:
        return None

    entry_time = _normalise_hhmm(trade.get("entry_time") or trade.get("entry_timestamp"))
    exit_time = _normalise_hhmm(trade.get("exit_time") or trade.get("exit_timestamp"))
    if not entry_time or not exit_time:
        return None

    by_minute = _candle_by_minute(series)
    entry_candle = by_minute.get(entry_time)
    exit_candle = by_minute.get(exit_time)
    if not entry_candle or not exit_candle:
        return None

    def f(value):
        try:
            n = float(value)
            return n if math.isfinite(n) else None
        except Exception:
            return None

    entry = f(entry_candle.get("close", entry_candle.get("ltp")))
    exit_ = f(exit_candle.get("close", exit_candle.get("ltp")))
    if entry is None or exit_ is None:
        return None

    window = []
    for p in series:
        t = _normalise_hhmm(p.get("time") or p.get("timestamp"))
        if t and entry_time <= t <= exit_time:
            window.append(p)

    highs = [f(p.get("high")) for p in window]
    lows = [f(p.get("low")) for p in window]
    highs = [v for v in highs if v is not None]
    lows = [v for v in lows if v is not None]

    highest = max(highs) if highs else max(entry, exit_)
    lowest = min(lows) if lows else min(entry, exit_)

    # All tracked strategy legs are option SELLs.
    premium_points = entry - exit_
    premium_pct = (premium_points / entry * 100.0) if entry else 0.0
    mae = max(0.0, highest - entry)   # premium rise is adverse for a seller
    mfe = max(0.0, entry - lowest)    # premium fall is favourable for a seller

    strike = None
    for c in (entry_candle, exit_candle):
        try:
            if c.get("strike") is not None:
                strike = int(float(c.get("strike")))
                break
        except Exception:
            pass
    if strike is None:
        premium_map = (day_data or {}).get("premium") or {}
        try:
            strike = int(float(premium_map.get(key)))
        except Exception:
            strike = None

    entry_oi = entry_candle.get("oi")
    exit_oi = exit_candle.get("oi")
    try:
        oi_change = int(exit_oi) - int(entry_oi) if entry_oi is not None and exit_oi is not None else None
    except Exception:
        oi_change = None

    return {
        "strike_key": key,
        "strike_label": PREMIUM_LABELS[key],
        "strike": strike,
        "option_type": option_type,
        "entry_time": entry_time,
        "exit_time": exit_time,
        "entry_premium": round(entry, 2),
        "exit_premium": round(exit_, 2),
        "premium_points": round(premium_points, 2),
        "premium_pct": round(premium_pct, 2),
        "highest_premium": round(highest, 2),
        "lowest_premium": round(lowest, 2),
        "premium_mae": round(mae, 2),
        "premium_mfe": round(mfe, 2),
        "entry_oi": entry_oi,
        "exit_oi": exit_oi,
        "oi_change": oi_change,
        "result": "WIN" if premium_points > 0 else "LOSS" if premium_points < 0 else "FLAT",
    }


def _collect_premium_report(start_text, end_text):
    start_day = _parse_report_day(start_text)
    end_day = _parse_report_day(end_text)
    if not start_day or not end_day:
        return None, "Please select a valid From and To date."
    if start_day > end_day:
        return None, "From date cannot be after To date."

    days = _report_days(start_day, end_day)
    rows = []

    for day in days:
        if day == today_key():
            with lock:
                snapshot_current_premium_candles()
                day_data = json.loads(json.dumps({
                    "date": day,
                    "opening_atm": state.get("opening_atm"),
                    "premium": state.get("premium", {}),
                    "series": state.get("series", {}),
                    "strategies": state.get("strategies", {}),
                }))
        else:
            day_data = load_day_history(day)

        trades = _strategy_trades_from_day(day, day_data)
        for trade in trades:
            # Only completed strategy trades have a reproducible entry -> exit interval.
            if not (trade.get("exit_time") or trade.get("exit_timestamp")):
                continue
            option_type = _trade_option_type(trade)
            if option_type not in ("CE", "PE"):
                continue

            for key in PREMIUM_KEYS:
                metrics = _premium_trade_metrics(trade, day_data, key, option_type)
                if not metrics:
                    continue
                row = {
                    "date": day,
                    "strategy": str(trade.get("strategy") or "").upper(),
                    "trade_no": trade.get("trade_no"),
                    "signal": trade.get("type") or trade.get("side") or f"SELL {option_type}",
                    "nifty_entry": trade.get("entry_level"),
                    "nifty_exit": trade.get("exit_level"),
                    "nifty_points": trade.get("points"),
                    "exit_reason": trade.get("exit_reason"),
                    "sl_hit": bool(trade.get("sl_hit")),
                }
                row.update(metrics)
                rows.append(row)

    def _summary(group_rows):
        pts = [float(r["premium_points"]) for r in group_rows]
        wins = sum(1 for p in pts if p > 0)
        losses = sum(1 for p in pts if p < 0)
        flat = sum(1 for p in pts if p == 0)
        return {
            "trades": len(group_rows),
            "wins": wins,
            "losses": losses,
            "flat": flat,
            "win_rate": round(wins / len(group_rows) * 100, 2) if group_rows else 0.0,
            "total_premium_points": round(sum(pts), 2),
            "avg_premium_points": round(sum(pts) / len(pts), 2) if pts else 0.0,
            "avg_mae": round(sum(float(r["premium_mae"]) for r in group_rows) / len(group_rows), 2) if group_rows else 0.0,
            "avg_mfe": round(sum(float(r["premium_mfe"]) for r in group_rows) / len(group_rows), 2) if group_rows else 0.0,
            "best_trade": round(max(pts), 2) if pts else 0.0,
            "worst_trade": round(min(pts), 2) if pts else 0.0,
        }

    strategy_summary = {}
    for name in ("S1", "S2", "PNA"):
        strategy_summary[name] = _summary([r for r in rows if r["strategy"] == name])

    strike_summary = {}
    for key in PREMIUM_KEYS:
        strike_summary[key] = _summary([r for r in rows if r["strike_key"] == key])

    strategy_strike_summary = []
    for name in ("S1", "S2", "PNA"):
        for key in PREMIUM_KEYS:
            group = [r for r in rows if r["strategy"] == name and r["strike_key"] == key]
            s = _summary(group)
            strategy_strike_summary.append({"strategy": name, "strike_key": key, "strike_label": PREMIUM_LABELS[key], **s})

    daily = []
    for day in days:
        item = {"date": day}
        day_rows = [r for r in rows if r["date"] == day]
        item.update(_summary(day_rows))
        daily.append(item)

    overall = _summary(rows)
    return {
        "from": start_text,
        "to": end_text,
        "days": days,
        "overall": overall,
        "strategy_summary": strategy_summary,
        "strike_summary": strike_summary,
        "strategy_strike_summary": strategy_strike_summary,
        "daily": daily,
        "trades": rows,
        "note": "Premium entry/exit uses stored 1-minute option candle close at each strategy signal minute; MAE/MFE uses minute high/low while the trade is open.",
    }, None


@app.route("/api/premium/reports")
def api_premium_reports():
    report, error = _collect_premium_report(request.args.get("from", ""), request.args.get("to", ""))
    if error:
        return jsonify({"error": error}), 400
    return jsonify(report)


def _style_workbook(wb):
    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        for row in sheet.iter_rows():
            for cell in row:
                if cell.row == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center")
        for column_cells in sheet.columns:
            letter = column_cells[0].column_letter
            max_len = max((len(str(c.value or "")) for c in column_cells), default=8)
            sheet.column_dimensions[letter].width = min(max(max_len + 2, 11), 34)


@app.route("/api/download/premium/reports")
def download_premium_reports_excel():
    start_text = request.args.get("from", "")
    end_text = request.args.get("to", "")
    report, error = _collect_premium_report(start_text, end_text)
    if error:
        return jsonify({"error": error}), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "Premium Overall Summary"
    ws.append(["PRATIK ANALYSIS — OPTION PREMIUM PERFORMANCE"])
    ws.append(["From", start_text, "To", end_text])
    ws.append(["Method", report["note"]])
    ws.append([])
    ws.append(["Metric", "Value"])
    for key, label in [
        ("trades", "Premium observations (3 strikes per completed trade)"),
        ("wins", "Profitable premium observations"),
        ("losses", "Losing premium observations"),
        ("flat", "Flat premium observations"),
        ("win_rate", "Win %"),
        ("total_premium_points", "Total premium points"),
        ("avg_premium_points", "Avg premium points / observation"),
        ("avg_mae", "Avg premium MAE"),
        ("avg_mfe", "Avg premium MFE"),
        ("best_trade", "Best premium observation"),
        ("worst_trade", "Worst premium observation"),
    ]:
        ws.append([label, report["overall"][key]])

    ss = wb.create_sheet("Strategy-wise Summary")
    headers = ["Strategy", "Observations", "Wins", "Losses", "Flat", "Win %", "Total Premium Pts", "Avg Premium Pts", "Avg MAE", "Avg MFE", "Best", "Worst"]
    ss.append(headers)
    for name in ("S1", "S2", "PNA"):
        m = report["strategy_summary"][name]
        ss.append([name, m["trades"], m["wins"], m["losses"], m["flat"], m["win_rate"], m["total_premium_points"], m["avg_premium_points"], m["avg_mae"], m["avg_mfe"], m["best_trade"], m["worst_trade"]])

    st = wb.create_sheet("Strike-wise Summary")
    st.append(["Strike Group", "Observations", "Wins", "Losses", "Flat", "Win %", "Total Premium Pts", "Avg Premium Pts", "Avg MAE", "Avg MFE", "Best", "Worst"])
    for key in PREMIUM_KEYS:
        m = report["strike_summary"][key]
        st.append([PREMIUM_LABELS[key], m["trades"], m["wins"], m["losses"], m["flat"], m["win_rate"], m["total_premium_points"], m["avg_premium_points"], m["avg_mae"], m["avg_mfe"], m["best_trade"], m["worst_trade"]])

    combo = wb.create_sheet("Strategy x Strike")
    combo.append(["Strategy", "Strike Group", "Observations", "Wins", "Losses", "Flat", "Win %", "Total Premium Pts", "Avg Premium Pts", "Avg MAE", "Avg MFE", "Best", "Worst"])
    for x in report["strategy_strike_summary"]:
        combo.append([x["strategy"], x["strike_label"], x["trades"], x["wins"], x["losses"], x["flat"], x["win_rate"], x["total_premium_points"], x["avg_premium_points"], x["avg_mae"], x["avg_mfe"], x["best_trade"], x["worst_trade"]])

    trades = wb.create_sheet("Premium Entry Exit")
    trade_headers = [
        "Date", "Strategy", "Trade #", "Signal", "Strike Group", "Strike", "Option",
        "Entry Time", "Entry Premium", "Exit Time", "Exit Premium", "Premium Points", "Premium %",
        "Highest Premium", "Lowest Premium", "Premium MAE", "Premium MFE", "Entry OI", "Exit OI", "OI Change",
        "NIFTY Entry", "NIFTY Exit", "NIFTY Points", "SL Hit?", "Premium Result", "Exit Reason",
    ]
    trades.append(trade_headers)
    for r in report["trades"]:
        trades.append([
            r["date"], r["strategy"], r["trade_no"], r["signal"], r["strike_label"], r["strike"], r["option_type"],
            r["entry_time"], r["entry_premium"], r["exit_time"], r["exit_premium"], r["premium_points"], r["premium_pct"],
            r["highest_premium"], r["lowest_premium"], r["premium_mae"], r["premium_mfe"], r["entry_oi"], r["exit_oi"], r["oi_change"],
            r["nifty_entry"], r["nifty_exit"], r["nifty_points"], "YES" if r["sl_hit"] else "NO", r["result"], r["exit_reason"],
        ])

    mm = wb.create_sheet("Premium MAE MFE")
    mm.append(["Date", "Strategy", "Trade #", "Strike Group", "Strike", "Option", "Entry Premium", "Highest", "Lowest", "MAE", "MFE", "Premium Points"])
    for r in report["trades"]:
        mm.append([r["date"], r["strategy"], r["trade_no"], r["strike_label"], r["strike"], r["option_type"], r["entry_premium"], r["highest_premium"], r["lowest_premium"], r["premium_mae"], r["premium_mfe"], r["premium_points"]])

    daily = wb.create_sheet("Daily Premium Summary")
    daily.append(["Date", "Observations", "Wins", "Losses", "Flat", "Win %", "Total Premium Pts", "Avg Premium Pts", "Avg MAE", "Avg MFE", "Best", "Worst"])
    for d in report["daily"]:
        daily.append([d["date"], d["trades"], d["wins"], d["losses"], d["flat"], d["win_rate"], d["total_premium_points"], d["avg_premium_points"], d["avg_mae"], d["avg_mfe"], d["best_trade"], d["worst_trade"]])

    compare = wb.create_sheet("NIFTY vs Premium")
    compare.append(["Date", "Strategy", "Trade #", "Signal", "Strike Group", "Strike", "Option", "NIFTY Points", "Premium Points", "Premium %", "Premium Result"])
    for r in report["trades"]:
        compare.append([r["date"], r["strategy"], r["trade_no"], r["signal"], r["strike_label"], r["strike"], r["option_type"], r["nifty_points"], r["premium_points"], r["premium_pct"], r["result"]])

    _style_workbook(wb)
    # Restore the intended premium-overall title presentation after generic styling.
    ws["A1"].font = Font(bold=True, size=16)
    ws["A5"].font = Font(bold=True)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    filename = f"Pratik_Analysis_Premium_Report_{start_text}_to_{end_text}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/download/premium/raw/<day>")
def download_raw_premium_excel(day):
    if day == today_key():
        with lock:
            snapshot_current_premium_candles()
            data = json.loads(json.dumps({
                "date": day,
                "expiry": state.get("expiry"),
                "opening_atm": state.get("opening_atm"),
                "premium": state.get("premium", {}),
                "series": state.get("series", {}),
            }))
    else:
        if not history_exists(day):
            return jsonify({"error": "No stored data for selected date."}), 404
        data = load_day_history(day)

    wb = Workbook()
    ws = wb.active
    ws.title = "Raw Premium 1-Min"
    ws.append(["Pratik Analysis — 1-Min Option Premium OHLC + OI"])
    ws.append(["Date", day, "Expiry", data.get("expiry"), "Opening ATM", data.get("opening_atm")])
    ws.append([])
    headers = [
        "Time",
        "ATM -100 CE O", "H", "L", "C", "OI",
        "ATM -100 PE O", "H", "L", "C", "OI",
        "ATM CE O", "H", "L", "C", "OI",
        "ATM PE O", "H", "L", "C", "OI",
        "ATM +100 CE O", "H", "L", "C", "OI",
        "ATM +100 PE O", "H", "L", "C", "OI",
    ]
    ws.append(headers)

    maps = {}
    all_times = set()
    for key in PREMIUM_KEYS:
        for option_type in ("CE", "PE"):
            m = _candle_by_minute(_series_for_premium(data, key, option_type))
            maps[(key, option_type)] = m
            all_times.update(m.keys())

    for time_key in sorted(all_times):
        row = [time_key]
        for key in PREMIUM_KEYS:
            for option_type in ("CE", "PE"):
                p = maps[(key, option_type)].get(time_key) or {}
                row.extend([p.get("open"), p.get("high"), p.get("low"), p.get("close", p.get("ltp")), p.get("oi")])
        ws.append(row)

    _style_workbook(wb)
    ws["A1"].font = Font(bold=True, size=16)
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Pratik_Analysis_Raw_Premium_{day}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# HEALTH
# ============================================================

@app.route("/health")
def health():
    return jsonify(
        {
            "ok": True,
            "connected": bool(
                state.get(
                    "connected"
                )
            ),
            "feed_message": state.get("message"),
            "reconnect_watchdog": reconnect_watchdog_started,
            "reconnect_in_progress": reconnect_in_progress,
            "date": state.get(
                "date"
            ),
            "option_tokens": len(
                token_meta
            ),
            "latest_oi_tokens": len(
                latest_oi
            ),
            "baseline_ready": baseline_ready,
            "opening_atm": state.get("opening_atm"),
            "oic_tokens": oic_tokens,
            "oic_mapping_count": len(oic_tokens),
            "premium_tokens": premium_token_map,
            "premium_mapping_count": len(premium_token_map),
            "premium_points": {
                key: {
                    option_type: len(
                        state.get("series", {}).get("premium", {}).get(key, {}).get(option_type, [])
                    )
                    for option_type in ("CE", "PE")
                }
                for key in ("minus100", "atm", "plus100")
            },
            "database_configured": db_enabled(),
            "database_history_days": len(db_history_dates()),
            "cio_points": len(
                state.get(
                    "series",
                    {},
                ).get(
                    "cio",
                    [],
                )
            ),
        }
    )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    port = int(
        os.environ.get(
            "PORT",
            "8000",
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
    )
