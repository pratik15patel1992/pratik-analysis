
import os
import math
import time
import threading
from datetime import datetime, timedelta
from io import BytesIO

from flask import Flask, render_template, redirect, request, jsonify, send_file
from kiteconnect import KiteConnect
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

APP_SECRET = os.environ.get("APP_SECRET", "change-me")
KITE_API_KEY = os.environ.get("KITE_API_KEY", "")
KITE_API_SECRET = os.environ.get("KITE_API_SECRET", "")
PORT = int(os.environ.get("PORT", "8000"))

IST_OFFSET = timedelta(hours=5, minutes=30)
MARKET_START = 9 * 60 + 15
MARKET_END = 15 * 60 + 30
STRATEGY_START = 9 * 60 + 30
STRATEGY_EXIT = 14 * 60 + 45
HARD_SL = 30.0
MAX_TRADES = 3
MASTER_CLEAR = 35.0

app = Flask(__name__)
app.secret_key = APP_SECRET

lock = threading.RLock()
rest_lock = threading.Lock()

kite = None
access_token = None
nifty_token = None
option_rows = []
token_meta = {}
locked = {}
latest_oi = {}
baseline_oi = {}
baseline_ready = False
last_poll = 0.0
worker_started = False

state = {
    "configured": bool(KITE_API_KEY and KITE_API_SECRET),
    "connected": False,
    "message": "Waiting for Zerodha login",
    "last_update": None,
    "date": None,
    "expiry": None,
    "opening_atm": None,
    "nifty": {"price": None, "open": None, "high": None, "low": None, "previous_close": None},
    "series": {"nifty": [], "atm": [], "minus100": [], "plus100": [], "cio": []},
    "strategies": {
        "S1": {"trades": [], "active": None},
        "S2": {"trades": [], "active": None},
        "PNA": {"trades": [], "active": None},
    },
    "strategy_engine": {"last_eval_minute": None, "recent": {}},
}

def now_ist():
    return datetime.utcnow() + IST_OFFSET

def today_key():
    return now_ist().strftime("%Y-%m-%d")

def market_minute():
    n = now_ist()
    return n.hour * 60 + n.minute

def is_market():
    m = market_minute()
    return MARKET_START <= m <= MARKET_END

def minute_label():
    return now_ist().strftime("%H:%M")

def round_to_100(v):
    return int(round(float(v) / 100.0) * 100)

def append_or_replace(series, point):
    if series and series[-1].get("time") == point.get("time"):
        series[-1] = point
    else:
        series.append(point)

def discover(k):
    global nifty_token, option_rows, token_meta
    nse = k.instruments("NSE")
    nfo = k.instruments("NFO")

    nifty_token = None
    for r in nse:
        sym = str(r.get("tradingsymbol", "")).upper()
        name = str(r.get("name", "")).upper()
        if sym == "NIFTY 50" or name == "NIFTY 50":
            nifty_token = int(r["instrument_token"])
            break
    if not nifty_token:
        raise RuntimeError("NIFTY 50 token not found")

    today = now_ist().date()
    candidates = []
    for r in nfo:
        if str(r.get("name", "")).upper() != "NIFTY":
            continue
        typ = str(r.get("instrument_type", "")).upper()
        if typ not in ("CE", "PE"):
            continue
        exp = r.get("expiry")
        if isinstance(exp, str):
            try:
                exp = datetime.strptime(exp, "%Y-%m-%d").date()
            except Exception:
                continue
        if not exp or exp < today:
            continue
        candidates.append(r)

    if not candidates:
        raise RuntimeError("No active NIFTY options found")

    expiry = min(r["expiry"] for r in candidates)
    option_rows = [r for r in candidates if r["expiry"] == expiry]
    token_meta = {}
    for r in option_rows:
        token_meta[int(r["instrument_token"])] = {
            "strike": int(float(r["strike"])),
            "type": str(r["instrument_type"]).upper(),
            "symbol": r["tradingsymbol"],
        }

    with lock:
        state["expiry"] = str(expiry)

def quote_nifty(k):
    q = k.quote(["NSE:NIFTY 50"]).get("NSE:NIFTY 50") or {}
    ohlc = q.get("ohlc") or {}
    return {
        "price": float(q["last_price"]) if q.get("last_price") is not None else None,
        "open": float(ohlc["open"]) if ohlc.get("open") is not None else None,
        "high": float(ohlc["high"]) if ohlc.get("high") is not None else None,
        "low": float(ohlc["low"]) if ohlc.get("low") is not None else None,
        "previous_close": float(ohlc["close"]) if ohlc.get("close") is not None else None,
    }

def lock_strikes(k):
    global locked
    with lock:
        atm = state.get("opening_atm")
    if not atm:
        n = quote_nifty(k)
        if n.get("open") is None:
            return False
        atm = round_to_100(n["open"])
        with lock:
            state["opening_atm"] = atm
            state["nifty"].update(n)

    wanted = {"minus100": atm - 100, "atm": atm, "plus100": atm + 100}
    m = {}
    for key, strike in wanted.items():
        ce = pe = None
        for token, meta in token_meta.items():
            if meta["strike"] != strike:
                continue
            if meta["type"] == "CE":
                ce = token
            elif meta["type"] == "PE":
                pe = token
        if ce and pe:
            m[key] = {"strike": strike, "CE": ce, "PE": pe}
    locked = m
    return len(locked) == 3

def previous_oi(k, token):
    end = now_ist().date() - timedelta(days=1)
    start = end - timedelta(days=10)
    try:
        candles = k.historical_data(token, start, end, "day", oi=True)
        for c in reversed(candles or []):
            if c.get("oi") is not None:
                return int(c["oi"])
    except Exception:
        pass
    return None

def build_baseline():
    global baseline_ready, baseline_oi
    result = {}
    try:
        for i, token in enumerate(token_meta):
            v = previous_oi(kite, token)
            if v is not None:
                result[token] = v
            time.sleep(0.35)
        baseline_oi = result
        baseline_ready = True
        print(f"[BASELINE] ready {len(result)} contracts", flush=True)
    except Exception as e:
        baseline_ready = False
        print(f"[BASELINE] error: {e}", flush=True)

def cio_totals():
    ce = pe = 0
    for token, cur in latest_oi.items():
        base = baseline_oi.get(token)
        meta = token_meta.get(token)
        if base is None or not meta:
            continue
        delta = int(cur) - int(base)
        if delta >= 0:
            continue
        if meta["type"] == "CE":
            ce += delta
        elif meta["type"] == "PE":
            pe += delta
    return ce, pe

def _row_minutes(t):
    try:
        h, m = str(t).split(":")[:2]
        return int(h) * 60 + int(m)
    except Exception:
        return None

def _point_at_or_before(rows, target):
    best = None
    for r in rows or []:
        m = _row_minutes(r.get("time"))
        if m is None or m > target:
            continue
        if best is None or m > best[0]:
            best = (m, r)
    return best[1] if best else None

def _five_delta(rows):
    if not rows:
        return None
    cur = rows[-1]
    cm = _row_minutes(cur.get("time"))
    old = _point_at_or_before(rows, cm - 5) if cm is not None else None
    if not old:
        return None
    return {
        "ce": float(cur.get("ce", 0)) - float(old.get("ce", 0)),
        "pe": float(cur.get("pe", 0)) - float(old.get("pe", 0)),
    }

def _flow(d):
    if not d:
        return None
    ce, pe = d["ce"], d["pe"]
    if pe > 0 and ce < 0:
        return "PE"
    if ce > 0 and pe < 0:
        return "CE"
    return None

def _dom(d):
    if not d:
        return 0.0
    ce, pe = float(d["ce"]), float(d["pe"])
    den = abs(ce) + abs(pe)
    return ((ce - pe) / den * 100) if den else 0.0

def metrics():
    ds, dirs, dom = {}, {}, {}
    for key in ("minus100", "atm", "plus100"):
        d = _five_delta(state["series"][key])
        ds[key] = d
        dirs[key] = _flow(d)
        dom[key] = _dom(d)
    pe_votes = sum(1 for x in dirs.values() if x == "PE")
    ce_votes = sum(1 for x in dirs.values() if x == "CE")
    master = 0.25 * dom["minus100"] + 0.50 * dom["atm"] + 0.25 * dom["plus100"]
    cio = state["series"]["cio"][-1] if state["series"]["cio"] else {}
    cce = float(cio.get("ce", 0) or 0)
    cpe = float(cio.get("pe", 0) or 0)
    cdir = "PE" if cce < cpe else ("CE" if cpe < cce else None)
    return {"pe_votes": pe_votes, "ce_votes": ce_votes, "master_d": master, "cio_direction": cdir}

def recent(name, direction):
    return int(state["strategy_engine"]["recent"].get(f"{name}:{direction}", 0))

def set_recent(name, direction, v):
    state["strategy_engine"]["recent"][f"{name}:{direction}"] = int(v)

def persist(name, direction, cond):
    other = "CE" if direction == "PE" else "PE"
    set_recent(name, direction, recent(name, direction) + 1 if cond else 0)
    if cond:
        set_recent(name, other, 0)
    return recent(name, direction)

def enter(name, typ, price):
    s = state["strategies"][name]
    if s["active"] or len(s["trades"]) >= MAX_TRADES:
        return
    p = float(price)
    sl = p - HARD_SL if typ == "PE" else p + HARD_SL
    s["active"] = {
        "type": typ,
        "entry_time": minute_label(),
        "entry_level": round(p, 2),
        "sl_level": round(sl, 2),
        "exit_time": None,
        "exit_level": None,
        "points": None,
    }

def exit_trade(name, price, reason):
    s = state["strategies"][name]
    t = s.get("active")
    if not t:
        return
    p, e = float(price), float(t["entry_level"])
    pts = p - e if t["type"] == "PE" else e - p
    t.update({
        "exit_time": minute_label(),
        "exit_level": round(p, 2),
        "points": round(pts, 2),
        "exit_reason": reason,
    })
    s["trades"].append(dict(t))
    s["active"] = None
    set_recent(name, "PE", 0)
    set_recent(name, "CE", 0)

def evaluate_strategies():
    price = state["nifty"].get("price")
    if price is None or not baseline_ready:
        return
    m = market_minute()
    label = minute_label()
    if state["strategy_engine"].get("last_eval_minute") == label:
        return
    state["strategy_engine"]["last_eval_minute"] = label
    met = metrics()

    for name in ("S1", "S2", "PNA"):
        t = state["strategies"][name].get("active")
        if not t:
            continue
        if (t["type"] == "PE" and price <= t["sl_level"]) or (t["type"] == "CE" and price >= t["sl_level"]):
            exit_trade(name, price, "30-POINT HARD SL")
            continue
        if m >= STRATEGY_EXIT:
            exit_trade(name, price, "14:45 TIME EXIT")
            continue
        opposite = "CE" if t["type"] == "PE" else "PE"
        votes = met["ce_votes"] if opposite == "CE" else met["pe_votes"]
        cio_ok = met["cio_direction"] == opposite
        if name == "S1":
            if persist("S1_EXIT", opposite, votes >= 2) >= 3 or met["cio_direction"] not in (t["type"], None):
                exit_trade(name, price, "OIC/CIO REVERSAL")
        elif name == "S2":
            if votes >= 2 and cio_ok:
                exit_trade(name, price, "MOMENTUM REVERSAL")
        else:
            master_ok = met["master_d"] >= MASTER_CLEAR if opposite == "CE" else met["master_d"] <= -MASTER_CLEAR
            if votes >= 2 and cio_ok and master_ok:
                exit_trade(name, price, "MASTER OIC/CIO REVERSAL")

    if m < STRATEGY_START or m >= STRATEGY_EXIT:
        return

    for direction in ("PE", "CE"):
        votes = met["pe_votes"] if direction == "PE" else met["ce_votes"]
        common = votes >= 2 and met["cio_direction"] == direction

        if common and not state["strategies"]["S2"]["active"]:
            enter("S2", direction, price)

        if persist("S1", direction, common) >= 3 and not state["strategies"]["S1"]["active"]:
            enter("S1", direction, price)

        master_ok = met["master_d"] <= -MASTER_CLEAR if direction == "PE" else met["master_d"] >= MASTER_CLEAR
        if persist("PNA", direction, common and master_ok) >= 3 and not state["strategies"]["PNA"]["active"]:
            enter("PNA", direction, price)

def poll_once():
    global last_poll
    if kite is None or not token_meta or not locked or not is_market():
        return

    if time.time() - last_poll < 1.8:
        return
    if not rest_lock.acquire(timeout=0.1):
        return
    try:
        last_poll = time.time()
        symbols = ["NSE:NIFTY 50"] + [f"NFO:{m['symbol']}" for m in token_meta.values()]
        quotes = kite.quote(symbols)

        nq = quotes.get("NSE:NIFTY 50") or {}
        ohlc = nq.get("ohlc") or {}
        price = nq.get("last_price")
        with lock:
            if price is not None:
                state["nifty"]["price"] = float(price)
            for a, b in (("open","open"),("high","high"),("low","low"),("close","previous_close")):
                if ohlc.get(a) is not None:
                    state["nifty"][b] = float(ohlc[a])

        for q in quotes.values():
            token = int(q.get("instrument_token") or 0)
            if token in token_meta and q.get("oi") is not None:
                latest_oi[token] = int(q["oi"])

        n = state["nifty"]
        if state.get("opening_atm") is None and n.get("open") is not None:
            state["opening_atm"] = round_to_100(n["open"])
            lock_strikes(kite)

        ts = now_ist().isoformat()
        with lock:
            # NIFTY minute close for Excel + strategy cutoff reference
            if n.get("price") is not None:
                p = float(n["price"])
                append_or_replace(state["series"]["nifty"], {
                    "time": minute_label(),
                    "timestamp": ts,
                    "open": p, "high": p, "low": p, "close": p, "price": p
                })

            fresh = True
            for key, legs in locked.items():
                ce = latest_oi.get(legs["CE"])
                pe = latest_oi.get(legs["PE"])
                if ce is None or pe is None:
                    fresh = False
                    continue
                append_or_replace(state["series"][key], {
                    "time": minute_label(), "timestamp": ts,
                    "ce": int(ce), "pe": int(pe), "source": "LIVE"
                })

            if baseline_ready:
                ce, pe = cio_totals()
                append_or_replace(state["series"]["cio"], {
                    "time": minute_label(), "timestamp": ts,
                    "ce": int(ce), "pe": int(pe), "source": "LIVE"
                })

            if fresh and baseline_ready:
                evaluate_strategies()

            state["date"] = today_key()
            state["connected"] = True
            state["message"] = "LIVE — Zerodha REST"
            state["last_update"] = ts

    except Exception as e:
        with lock:
            state["connected"] = False
            state["message"] = f"Feed warning: {type(e).__name__}"
        print(f"[POLL] {type(e).__name__}: {e}", flush=True)
    finally:
        rest_lock.release()

def worker():
    while True:
        try:
            poll_once()
        except Exception as e:
            print(f"[WORKER] {e}", flush=True)
        time.sleep(0.5)

def start_worker():
    global worker_started
    if worker_started:
        return
    worker_started = True
    threading.Thread(target=worker, daemon=True).start()

def start_live(token):
    global kite, access_token, baseline_ready
    k = KiteConnect(api_key=KITE_API_KEY)
    k.set_access_token(token)
    k.profile()
    discover(k)
    kite = k
    access_token = token

    with lock:
        state["date"] = today_key()
        state["connected"] = True
        state["message"] = "Zerodha login accepted — starting feed"

    lock_strikes(k)
    baseline_ready = False
    threading.Thread(target=build_baseline, daemon=True).start()
    start_worker()

@app.after_request
def no_cache(resp):
    if request.path.startswith("/api/") or request.path == "/health":
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return resp

@app.route("/")
def index():
    return render_template("index.html", configured=bool(KITE_API_KEY and KITE_API_SECRET))

@app.route("/kite/login")
def kite_login():
    if not KITE_API_KEY:
        return "KITE_API_KEY not configured", 400
    return redirect(KiteConnect(api_key=KITE_API_KEY).login_url())

@app.route("/kite/callback")
def kite_callback():
    token = request.args.get("request_token")
    if not token:
        return "Missing request_token", 400
    k = KiteConnect(api_key=KITE_API_KEY)
    s = k.generate_session(token, api_secret=KITE_API_SECRET)
    start_live(s["access_token"])
    return redirect("/")

@app.route("/api/state")
def api_state():
    with lock:
        return jsonify(state)

@app.route("/api/download/cio")
def download_excel():
    with lock:
        s = {
            "date": state.get("date") or today_key(),
            "expiry": state.get("expiry"),
            "opening_atm": state.get("opening_atm"),
            "series": {k: list(state["series"].get(k, [])) for k in ("nifty","minus100","atm","plus100","cio")},
        }

    def by_time(rows):
        return {str(r.get("time")): r for r in rows if r.get("time")}
    maps = {k: by_time(v) for k,v in s["series"].items()}
    all_times = sorted(set().union(*(m.keys() for m in maps.values())))

    wb = Workbook()
    ws = wb.active
    ws.title = "NIFTY + OIC + CIO"
    ws["A1"] = "Pratik Analysis"
    ws["A2"] = "NIFTY + OIC + CIO — Minute-wise Data"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"].font = Font(bold=True, size=13)
    ws["A4"] = "Date"; ws["B4"] = s["date"]
    ws["A5"] = "Opening ATM"; ws["B5"] = s["opening_atm"]
    ws["A6"] = "Expiry"; ws["B6"] = s["expiry"]

    headers = [
        "Time","NIFTY Close",
        "ATM -100 CE OI","ATM -100 PE OI",
        "ATM CE OI","ATM PE OI",
        "ATM +100 CE OI","ATM +100 PE OI",
        "CIO CE Negative Change in OI","CIO PE Negative Change in OI",
    ]
    for c,h in enumerate(headers,1):
        ws.cell(8,c,h).font = Font(bold=True)

    for rno,t in enumerate(all_times,9):
        n=maps["nifty"].get(t,{})
        m=maps["minus100"].get(t,{})
        a=maps["atm"].get(t,{})
        p=maps["plus100"].get(t,{})
        c=maps["cio"].get(t,{})
        vals=[t,n.get("close",n.get("price")),m.get("ce"),m.get("pe"),a.get("ce"),a.get("pe"),p.get("ce"),p.get("pe"),c.get("ce"),c.get("pe")]
        for col,val in enumerate(vals,1):
            ws.cell(rno,col,val)

    ws.freeze_panes = "A9"
    for col in range(1, 11):
        ws.column_dimensions[chr(64+col)].width = 24 if col > 2 else 16

    buf = BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name=f"Pratik_NIFTY_OIC_CIO_{s['date']}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/health")
def health():
    return jsonify({
        "ok": True,
        "connected": state.get("connected"),
        "message": state.get("message"),
        "date": state.get("date"),
        "expiry": state.get("expiry"),
        "opening_atm": state.get("opening_atm"),
        "locked_strikes": locked,
        "latest_oi_tokens": len(latest_oi),
        "baseline_ready": baseline_ready,
        "points": {k: len(state["series"][k]) for k in ("atm","minus100","plus100","cio")},
    })

if __name__ == "__main__":
    start_worker()
    app.run(host="0.0.0.0", port=PORT, debug=False)
