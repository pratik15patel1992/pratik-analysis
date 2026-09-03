
import os, json, time, threading
from datetime import datetime, timedelta, date
from pathlib import Path

from flask import Flask, render_template, redirect, request, jsonify
from kiteconnect import KiteConnect, KiteTicker

APP_SECRET = os.environ.get("APP_SECRET", "change-me")
KITE_API_KEY = os.environ.get("KITE_API_KEY", "")
KITE_API_SECRET = os.environ.get("KITE_API_SECRET", "")
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")

app = Flask(__name__)
app.secret_key = APP_SECRET

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

ACCESS_FILE = DATA / "access_token.json"
BASELINE_FILE = DATA / "oi_baseline.json"

lock = threading.Lock()
kite = None
ticker = None
token_meta = {}
latest_oi = {}
prev_oi = {}
oic_tokens = {}

state = {
    "connected": False,
    "message": "Waiting for Zerodha login",
    "session_date": str(date.today()),
    "oic": {
        "atm": None,
        "minus100": None,
        "plus100": None,
        "series": {"atm": [], "minus100": [], "plus100": []}
    },
    "cio": {"series": [], "baseline_ready": False},
    "last_update": None
}

def ist_now():
    # Render usually runs UTC; IST = UTC + 5:30.
    return datetime.utcnow() + timedelta(hours=5, minutes=30)

def tstamp():
    return ist_now().strftime("%H:%M:%S")

def trim(arr, n=800):
    if len(arr) > n:
        del arr[:-n]

def save_access_token(tok):
    ACCESS_FILE.write_text(json.dumps({"date": str(date.today()), "access_token": tok}))

def load_access_token():
    try:
        d = json.loads(ACCESS_FILE.read_text())
        if d.get("date") == str(date.today()):
            return d.get("access_token")
    except Exception:
        pass
    return None

def nifty_nearest_expiry_options(instruments):
    rows = [
        x for x in instruments
        if x.get("exchange") == "NFO"
        and x.get("name") == "NIFTY"
        and x.get("instrument_type") in ("CE", "PE")
        and x.get("expiry")
        and x.get("expiry") >= date.today()
    ]
    if not rows:
        return []
    expiry = min(x["expiry"] for x in rows)
    return [x for x in rows if x["expiry"] == expiry]

def build_oic_map(options, opening_price):
    if not opening_price:
        return
    # Locked convention: ATM rounded to nearest 100; then +/-100.
    atm = int(round(float(opening_price) / 100.0) * 100)
    strikes = {"atm": atm, "minus100": atm - 100, "plus100": atm + 100}
    temp = {}
    for key, strike in strikes.items():
        legs = {}
        for x in options:
            if int(float(x["strike"])) == strike and x["instrument_type"] in ("CE","PE"):
                legs[x["instrument_type"]] = int(x["instrument_token"])
        if len(legs) == 2:
            temp[key] = legs
    oic_tokens.clear()
    oic_tokens.update(temp)
    with lock:
        state["oic"]["atm"] = strikes["atm"]
        state["oic"]["minus100"] = strikes["minus100"]
        state["oic"]["plus100"] = strikes["plus100"]

def load_prev_oi(options, access_token):
    global prev_oi
    try:
        d = json.loads(BASELINE_FILE.read_text())
        if d.get("for_date") == str(date.today()):
            prev_oi = {int(k): int(v) for k, v in d.get("oi", {}).items()}
            with lock:
                state["cio"]["baseline_ready"] = bool(prev_oi)
            return
    except Exception:
        pass

    k = KiteConnect(api_key=KITE_API_KEY)
    k.set_access_token(access_token)
    end = ist_now().replace(hour=9, minute=0, second=0, microsecond=0)
    start = end - timedelta(days=10)
    baseline = {}

    # Historical OI baseline = last available daily OI before today.
    # A small sleep keeps requests conservative.
    for x in options:
        tok = int(x["instrument_token"])
        try:
            rows = k.historical_data(tok, start, end, "day", oi=True)
            valid = []
            for r in rows:
                dt = r.get("date")
                dte = dt.date() if hasattr(dt, "date") else None
                if dte and dte < date.today() and r.get("oi") is not None:
                    valid.append(r)
            if valid:
                baseline[tok] = int(valid[-1]["oi"])
        except Exception:
            pass
        time.sleep(0.38)

    prev_oi = baseline
    try:
        BASELINE_FILE.write_text(json.dumps({"for_date": str(date.today()), "oi": baseline}))
    except Exception:
        pass
    with lock:
        state["cio"]["baseline_ready"] = bool(prev_oi)

def cio_totals():
    ce = 0
    pe = 0
    for tok, cur in latest_oi.items():
        base = prev_oi.get(tok)
        meta = token_meta.get(tok)
        if base is None or not meta:
            continue
        d = int(cur) - int(base)
        if d < 0:
            if meta["type"] == "CE":
                ce += d
            elif meta["type"] == "PE":
                pe += d
    return ce, pe

def on_ticks(ws, ticks):
    ts = tstamp()
    for q in ticks:
        tok = int(q.get("instrument_token", 0))
        if tok in token_meta and q.get("oi") is not None:
            latest_oi[tok] = int(q["oi"])

    with lock:
        for key, legs in oic_tokens.items():
            ce = latest_oi.get(legs.get("CE"))
            pe = latest_oi.get(legs.get("PE"))
            if ce is not None and pe is not None:
                arr = state["oic"]["series"][key]
                if not arr or arr[-1]["t"] != ts:
                    arr.append({"t": ts, "ce": ce, "pe": pe})
                    trim(arr)

        ce_neg, pe_neg = cio_totals()
        carr = state["cio"]["series"]
        if state["cio"]["baseline_ready"] and (not carr or carr[-1]["t"] != ts):
            carr.append({"t": ts, "ce": ce_neg, "pe": pe_neg})
            trim(carr)

        state["last_update"] = ts

def on_connect(ws, response):
    tokens = list(token_meta.keys())
    if tokens:
        ws.subscribe(tokens)
        ws.set_mode(ws.MODE_FULL, tokens)
    with lock:
        state["connected"] = True
        state["message"] = "LIVE — Zerodha connected"

def on_close(ws, code, reason):
    with lock:
        state["connected"] = False
        state["message"] = f"Disconnected — {reason or code}"

def start_live(access_token):
    global kite, ticker, token_meta
    kite = KiteConnect(api_key=KITE_API_KEY)
    kite.set_access_token(access_token)

    instruments = kite.instruments()
    options = nifty_nearest_expiry_options(instruments)
    token_meta = {
        int(x["instrument_token"]): {
            "strike": float(x["strike"]),
            "type": x["instrument_type"],
            "symbol": x["tradingsymbol"]
        }
        for x in options
    }

    # NIFTY is used only internally to lock opening ATM; it is not shown.
    try:
        q = kite.quote(["NSE:NIFTY 50"]).get("NSE:NIFTY 50", {})
        opening = (q.get("ohlc") or {}).get("open") or q.get("last_price")
        build_oic_map(options, opening)
    except Exception as e:
        with lock:
            state["message"] = f"Connected, but ATM setup failed: {e}"

    threading.Thread(target=load_prev_oi, args=(options, access_token), daemon=True).start()

    ticker = KiteTicker(KITE_API_KEY, access_token)
    ticker.on_ticks = on_ticks
    ticker.on_connect = on_connect
    ticker.on_close = on_close
    ticker.connect(threaded=True)

@app.route("/")
def index():
    if not KITE_API_KEY or not KITE_API_SECRET:
        return render_template("index.html", configured=False, base_url=PUBLIC_BASE_URL)

    tok = load_access_token()
    if tok and not state["connected"]:
        try:
            start_live(tok)
        except Exception as e:
            with lock:
                state["message"] = f"Login required — {e}"
    return render_template("index.html", configured=True, base_url=PUBLIC_BASE_URL)

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
    sess = k.generate_session(request_token, api_secret=KITE_API_SECRET)
    access_token = sess["access_token"]
    save_access_token(access_token)
    start_live(access_token)
    return redirect("/")

@app.route("/api/state")
def api_state():
    with lock:
        return jsonify(state)

@app.route("/health")
def health():
    return jsonify({"ok": True})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
