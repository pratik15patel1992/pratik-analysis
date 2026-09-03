import os
import json
import time
import threading
from pathlib import Path
from datetime import datetime, timedelta, date

from flask import (
    Flask,
    render_template,
    redirect,
    request,
    jsonify,
    send_file,
)

from kiteconnect import KiteConnect, KiteTicker


# ============================================================
# CONFIGURATION
# ============================================================

APP_SECRET = os.environ.get("APP_SECRET", "change-me")
KITE_API_KEY = os.environ.get("KITE_API_KEY", "")
KITE_API_SECRET = os.environ.get("KITE_API_SECRET", "")
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")

app = Flask(__name__)
app.secret_key = APP_SECRET

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

SNAPSHOT_SECONDS = 60


# ============================================================
# RUNTIME STATE
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

    "series": {
        "atm": [],
        "minus100": [],
        "plus100": [],
        "cio": [],
    },

    "history_dates": [],
}

kite = None
ticker = None

nifty_token = None
vix_token = None

option_instruments = []
token_meta = {}

oic_tokens = {}

latest_oi = {}
prev_oi = {}
baseline_ready = False

latest_nifty = {}
latest_vix = {}

snapshot_thread_started = False
baseline_thread_started = False


# ============================================================
# TIME HELPERS
# ============================================================

def now_ist():
    return datetime.utcnow() + IST_OFFSET


def today_key():
    return now_ist().strftime("%Y-%m-%d")


def timestamp_label():
    return now_ist().strftime("%H:%M:%S")


def minute_label():
    return now_ist().strftime("%H:%M")


def is_market_session():
    n = now_ist()
    current = n.hour * 60 + n.minute
    start = MARKET_START_HOUR * 60 + MARKET_START_MINUTE
    end = MARKET_END_HOUR * 60 + MARKET_END_MINUTE
    return start <= current <= end


# ============================================================
# FILE HELPERS
# ============================================================

def safe_json_load(path, default=None):
    if default is None:
        default = {}

    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass

    return default


def safe_json_save(path, data):
    temp = path.with_suffix(path.suffix + ".tmp")

    with open(temp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    temp.replace(path)


def history_file(day):
    return HISTORY_DIR / f"{day}.json"


def load_day_history(day):
    return safe_json_load(
        history_file(day),
        {
            "date": day,
            "expiry": None,
            "opening_atm": None,
            "nifty": {},
            "vix": {},
            "zone": {},
            "series": {
                "atm": [],
                "minus100": [],
                "plus100": [],
                "cio": [],
            },
        },
    )


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
            "series": state.get("series", {}),
            "saved_at": now_ist().isoformat(),
        }

    safe_json_save(history_file(day), payload)
    refresh_history_dates()


def refresh_history_dates():
    dates = []

    for f in HISTORY_DIR.glob("*.json"):
        try:
            dates.append(f.stem)
        except Exception:
            pass

    dates.sort(reverse=True)

    with lock:
        state["history_dates"] = dates


# ============================================================
# ACCESS TOKEN
# ============================================================

def save_access_token(token):
    safe_json_save(
        ACCESS_TOKEN_FILE,
        {
            "access_token": token,
            "saved_at": now_ist().isoformat(),
        },
    )


def load_access_token():
    data = safe_json_load(ACCESS_TOKEN_FILE, {})

    token = data.get("access_token")

    if not token:
        return None

    return token


def clear_access_token():
    try:
        ACCESS_TOKEN_FILE.unlink(missing_ok=True)
    except Exception:
        pass


# ============================================================
# MARKET / ZONE HELPERS
# ============================================================

def round_to_100(value):
    return int(round(float(value) / 100.0) * 100)


def calculate_zone(open_price, previous_close):
    if not open_price or not previous_close:
        return {
            "quadrant": None,
            "zone": None,
            "opening_pct": None,
            "levels": {},
        }

    opening_pct = ((open_price - previous_close) / previous_close) * 100

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

    multipliers = [
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

    for pct in multipliers:
        value = previous_close * (1 + pct / 100.0)

        if pct > 0:
            key = f"+{pct:.2f}%"
        else:
            key = f"{pct:.2f}%"

        levels[key] = round(value, 2)

    return {
        "quadrant": quadrant,
        "zone": zone,
        "opening_pct": round(opening_pct, 3),
        "levels": levels,
    }


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
            "Elevated volatility. Expect larger intraday movement and higher option premiums.",
        )

    if value < 25:
        return (
            "20–25 HIGH",
            "High volatility. Use additional caution and tighter risk control.",
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

    with lock:
        state["message"] = "Loading Zerodha instruments..."

    nse = k.instruments("NSE")
    nfo = k.instruments("NFO")

    for row in nse:
        symbol = str(row.get("tradingsymbol", "")).upper()
        name = str(row.get("name", "")).upper()

        if symbol == "NIFTY 50" or name == "NIFTY 50":
            nifty_token = int(row["instrument_token"])

        if symbol == "INDIA VIX" or name == "INDIA VIX":
            vix_token = int(row["instrument_token"])

    today = now_ist().date()

    candidates = []

    for row in nfo:
        name = str(row.get("name", "")).upper()
        inst_type = str(row.get("instrument_type", "")).upper()

        if name != "NIFTY":
            continue

        if inst_type not in ("CE", "PE"):
            continue

        expiry = row.get("expiry")

        if not expiry:
            continue

        if isinstance(expiry, str):
            try:
                expiry = datetime.strptime(expiry, "%Y-%m-%d").date()
            except Exception:
                continue

        if expiry < today:
            continue

        candidates.append(row)

    if not candidates:
        raise RuntimeError("No active NIFTY option contracts found.")

    nearest_expiry = min(x["expiry"] for x in candidates)

    option_instruments = [
        x for x in candidates if x["expiry"] == nearest_expiry
    ]

    token_meta = {}

    for row in option_instruments:
        token = int(row["instrument_token"])

        token_meta[token] = {
            "strike": int(float(row["strike"])),
            "type": row["instrument_type"],
            "symbol": row["tradingsymbol"],
            "expiry": str(row["expiry"]),
        }

    with lock:
        state["expiry"] = str(nearest_expiry)


# ============================================================
# OIC TOKEN LOCK
# ============================================================

def lock_oic_strikes():
    global oic_tokens

    with lock:
        atm = state.get("opening_atm")

    if not atm:
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

    if len(mapping) != 3:
        return False

    oic_tokens = mapping

    with lock:
        state["oic"]["atm"] = atm
        state["oic"]["minus100"] = atm - 100
        state["oic"]["plus100"] = atm + 100

    return True


# ============================================================
# PREVIOUS-DAY OI BASELINE FOR CIO
# ============================================================

def previous_oi_for_token(k, token):
    end = now_ist().date() - timedelta(days=1)
    start = end - timedelta(days=10)

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

    except Exception:
        return None

    return None


def build_oi_baseline():
    global baseline_ready
    global baseline_thread_started
    global prev_oi

    baseline_thread_started = True

    try:
        cached = safe_json_load(BASELINE_FILE, {})

        cache_date = cached.get("date")
        cache_expiry = cached.get("expiry")

        with lock:
            current_expiry = state.get("expiry")

        if (
            cache_date == today_key()
            and cache_expiry == current_expiry
            and cached.get("oi")
        ):
            prev_oi = {
                int(k): int(v)
                for k, v in cached["oi"].items()
            }

            baseline_ready = True

            with lock:
                state["message"] = "LIVE — Zerodha connected"

            return

        result = {}

        with lock:
            state["message"] = (
                "Preparing previous-day OI baseline for CIO..."
            )

        for row in option_instruments:
            token = int(row["instrument_token"])

            value = previous_oi_for_token(kite, token)

            if value is not None:
                result[token] = value

            time.sleep(0.35)

        prev_oi = result

        safe_json_save(
            BASELINE_FILE,
            {
                "date": today_key(),
                "expiry": state.get("expiry"),
                "oi": {
                    str(k): v
                    for k, v in result.items()
                },
            },
        )

        baseline_ready = True

        with lock:
            state["message"] = "LIVE — Zerodha connected"

    except Exception as e:
        baseline_ready = False

        with lock:
            state["message"] = f"CIO baseline error — {e}"

    finally:
        baseline_thread_started = False


# ============================================================
# CIO CALCULATION
# ============================================================

def cio_totals():
    ce = 0
    pe = 0

    for token, current in latest_oi.items():
        baseline = prev_oi.get(token)
        meta = token_meta.get(token)

        if baseline is None or not meta:
            continue

        delta = int(current) - int(baseline)

        if delta >= 0:
            continue

        if meta["type"] == "CE":
            ce += delta

        elif meta["type"] == "PE":
            pe += delta

    return ce, pe


# ============================================================
# DATA SERIES HELPERS
# ============================================================

def append_or_replace_minute(series, point):
    if not series:
        series.append(point)
        return

    if series[-1].get("time") == point.get("time"):
        series[-1] = point
    else:
        series.append(point)


def oic_point(key):
    legs = oic_tokens.get(key)

    if not legs:
        return None

    ce = latest_oi.get(legs["CE"])
    pe = latest_oi.get(legs["PE"])

    if ce is None or pe is None:
        return None

    return {
        "time": minute_label(),
        "timestamp": now_ist().isoformat(),
        "ce": int(ce),
        "pe": int(pe),
    }


def make_snapshot():
    global state

    if not state.get("connected"):
        return

    with lock:
        for key in ("atm", "minus100", "plus100"):
            point = oic_point(key)

            if point:
                append_or_replace_minute(
                    state["series"][key],
                    point,
                )

        if baseline_ready:
            ce, pe = cio_totals()

            cio_point = {
                "time": minute_label(),
                "timestamp": now_ist().isoformat(),
                "ce": int(ce),
                "pe": int(pe),
            }

            append_or_replace_minute(
                state["series"]["cio"],
                cio_point,
            )

        state["last_update"] = now_ist().isoformat()

    save_current_history()


def snapshot_worker():
    while True:
        try:
            if state.get("connected") and is_market_session():
                make_snapshot()

        except Exception as e:
            with lock:
                state["message"] = f"Snapshot error — {e}"

        # Align close to minute boundary.
        n = now_ist()
        seconds_to_next = 60 - n.second

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
# WEBSOCKET CALLBACKS
# ============================================================

def on_ticks(ws, ticks):
    global latest_nifty
    global latest_vix

    for q in ticks:
        token = int(q.get("instrument_token", 0))

        # ------------------------------------
        # NIFTY SPOT
        # ------------------------------------
        if token == nifty_token:
            price = q.get("last_price")

            ohlc = q.get("ohlc") or {}

            open_price = ohlc.get("open")
            high = ohlc.get("high")
            low = ohlc.get("low")
            previous_close = ohlc.get("close")

            if price is not None:
                latest_nifty["price"] = float(price)

            if open_price:
                latest_nifty["open"] = float(open_price)

            if high:
                latest_nifty["high"] = float(high)

            if low:
                latest_nifty["low"] = float(low)

            if previous_close:
                latest_nifty["previous_close"] = float(previous_close)

            p = latest_nifty.get("price")
            pc = latest_nifty.get("previous_close")
            op = latest_nifty.get("open")

            if p is not None and pc:
                latest_nifty["change"] = round(p - pc, 2)
                latest_nifty["change_pct"] = round(
                    ((p - pc) / pc) * 100,
                    3,
                )

            with lock:
                state["nifty"] = dict(latest_nifty)

            if (
                state.get("opening_atm") is None
                and op is not None
            ):
                atm = round_to_100(op)

                with lock:
                    state["opening_atm"] = atm
                    state["zone"] = calculate_zone(op, pc)

                lock_oic_strikes()

            elif op is not None and pc:
                with lock:
                    state["zone"] = calculate_zone(op, pc)

        # ------------------------------------
        # INDIA VIX
        # ------------------------------------
        elif token == vix_token:
            price = q.get("last_price")

            if price is not None:
                value = float(price)
                band, interpretation = classify_vix(value)

                latest_vix = {
                    "price": value,
                    "range": band,
                    "interpretation": interpretation,
                }

                with lock:
                    state["vix"] = dict(latest_vix)

        # ------------------------------------
        # OPTIONS OI
        # ------------------------------------
        if token in token_meta:
            oi = q.get("oi")

            if oi is not None:
                latest_oi[token] = int(oi)

    with lock:
        state["last_update"] = now_ist().isoformat()


def on_connect(ws, response):
    tokens = []

    if nifty_token:
        tokens.append(nifty_token)

    if vix_token:
        tokens.append(vix_token)

    tokens.extend(token_meta.keys())

    tokens = list(set(tokens))

    if tokens:
        ws.subscribe(tokens)
        ws.set_mode(ws.MODE_FULL, tokens)

    with lock:
        state["connected"] = True
        state["message"] = "LIVE — Zerodha connected"


def on_close(ws, code, reason):
    with lock:
        state["connected"] = False
        state["message"] = (
            f"Disconnected — {reason or code}"
        )


def on_error(ws, code, reason):
    with lock:
        state["message"] = (
            f"WebSocket error ({code}) — {reason}"
        )


# ============================================================
# START LIVE CONNECTION
# ============================================================

def start_live(access_token):
    global kite
    global ticker
    global baseline_ready

    with lock:
        state["date"] = today_key()
        state["message"] = "Starting Zerodha live feed..."

    kite = KiteConnect(api_key=KITE_API_KEY)
    kite.set_access_token(access_token)

    # Verify token.
    kite.profile()

    discover_instruments(kite)

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

    ticker.connect(threaded=True)

    ensure_snapshot_worker()


# ============================================================
# RESTORE TODAY HISTORY
# ============================================================

def restore_today_history():
    day = today_key()
    saved = load_day_history(day)

    series = saved.get("series") or {}

    with lock:
        state["date"] = day

        for key in (
            "atm",
            "minus100",
            "plus100",
            "cio",
        ):
            if isinstance(series.get(key), list):
                state["series"][key] = series[key]

        if saved.get("opening_atm"):
            state["opening_atm"] = saved["opening_atm"]

        if saved.get("expiry"):
            state["expiry"] = saved["expiry"]

        if saved.get("zone"):
            state["zone"] = saved["zone"]

        if saved.get("nifty"):
            state["nifty"] = saved["nifty"]

        if saved.get("vix"):
            state["vix"] = saved["vix"]


restore_today_history()
refresh_history_dates()


# ============================================================
# FLASK ROUTES
# ============================================================

@app.route("/")
def index():
    if not KITE_API_KEY or not KITE_API_SECRET:
        return render_template(
            "index.html",
            configured=False,
            base_url=PUBLIC_BASE_URL,
        )

    token = load_access_token()

    if token and not state["connected"]:
        try:
            start_live(token)

        except Exception:
            clear_access_token()

            with lock:
                state["connected"] = False
                state["message"] = "Login required"

    return render_template(
        "index.html",
        configured=True,
        base_url=PUBLIC_BASE_URL,
    )


@app.route("/kite/login")
def kite_login():
    if not KITE_API_KEY:
        return "KITE_API_KEY is not configured", 400

    k = KiteConnect(api_key=KITE_API_KEY)

    return redirect(k.login_url())


@app.route("/kite/callback")
def kite_callback():
    request_token = request.args.get("request_token")

    if not request_token:
        return "Zerodha did not return a request_token", 400

    k = KiteConnect(api_key=KITE_API_KEY)

    session = k.generate_session(
        request_token,
        api_secret=KITE_API_SECRET,
    )

    access_token = session["access_token"]

    save_access_token(access_token)

    start_live(access_token)

    return redirect("/")


@app.route("/kite/logout")
def kite_logout():
    global ticker

    try:
        if ticker:
            ticker.close()

    except Exception:
        pass

    clear_access_token()

    with lock:
        state["connected"] = False
        state["message"] = "Logged out — Zerodha login required"

    return redirect("/")


@app.route("/api/state")
def api_state():
    with lock:
        return jsonify(state)


@app.route("/api/history/dates")
def api_history_dates():
    refresh_history_dates()

    with lock:
        return jsonify({
            "dates": state["history_dates"]
        })


@app.route("/api/history/<day>")
def api_history_day(day):
    path = history_file(day)

    if not path.exists():
        return jsonify({
            "error": "No stored data for selected date."
        }), 404

    return jsonify(load_day_history(day))


@app.route("/api/history/<day>/cio")
def api_history_cio(day):
    path = history_file(day)

    if not path.exists():
        return jsonify({
            "error": "No stored data for selected date."
        }), 404

    data = load_day_history(day)

    return jsonify({
        "date": day,
        "expiry": data.get("expiry"),
        "opening_atm": data.get("opening_atm"),
        "cio": (
            data.get("series", {})
            .get("cio", [])
        ),
    })


@app.route("/health")
def health():
    return jsonify({
        "ok": True,
        "connected": bool(state.get("connected")),
        "date": state.get("date"),
    })


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
    )
