import os
import math
import time
import json
import threading
from datetime import datetime, timedelta, date, time as dt_time
from io import BytesIO

from flask import Flask, render_template, redirect, request, jsonify, send_file
from kiteconnect import KiteConnect
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

try:
    import psycopg
    from psycopg.types.json import Jsonb
except Exception:
    psycopg = None
    Jsonb = None

APP_SECRET = os.environ.get("APP_SECRET", "change-me")
KITE_API_KEY = os.environ.get("KITE_API_KEY", "")
KITE_API_SECRET = os.environ.get("KITE_API_SECRET", "")
DATABASE_URL = os.environ.get("DATABASE_URL", "")
PORT = int(os.environ.get("PORT", "8000"))

IST_OFFSET = timedelta(hours=5, minutes=30)
MARKET_START = 9 * 60 + 15
MARKET_END = 15 * 60 + 30
STRATEGY_START = 9 * 60 + 30
STRATEGY_EXIT = 14 * 60 + 45
HARD_SL = 30.0
MAX_TRADES = 3
MASTER_CLEAR = 35.0
SEED_FILE = os.path.join(os.path.dirname(__file__), "seed_2026-09-08.json")

app = Flask(__name__)
app.secret_key = APP_SECRET

lock = threading.RLock()
rest_lock = threading.Lock()
db_lock = threading.Lock()

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
last_persist_label = None
last_persist_epoch = 0.0

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


def update_nifty_minute(price, ts):
    p = float(price)
    label = minute_label()
    rows = state["series"]["nifty"]
    if rows and rows[-1].get("time") == label:
        r = rows[-1]
        op = r.get("open")
        hi = r.get("high")
        lo = r.get("low")
        r.update({
            "timestamp": ts,
            "open": p if op is None else op,
            "high": p if hi is None else max(float(hi), p),
            "low": p if lo is None else min(float(lo), p),
            "close": p,
            "price": p,
        })
    else:
        rows.append({
            "time": label,
            "timestamp": ts,
            "open": p,
            "high": p,
            "low": p,
            "close": p,
            "price": p,
        })


# -----------------------
# PostgreSQL persistence
# -----------------------

def db_enabled():
    return bool(DATABASE_URL and psycopg is not None)


def db_connect():
    return psycopg.connect(DATABASE_URL, autocommit=True)


def init_db():
    if not db_enabled():
        print("[DB] DATABASE_URL/psycopg unavailable; running memory-only", flush=True)
        return False
    try:
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS pna_session_meta (
                        trade_date DATE PRIMARY KEY,
                        expiry DATE,
                        opening_atm INTEGER,
                        last_update TIMESTAMPTZ,
                        strategies JSONB NOT NULL DEFAULT '{}'::jsonb,
                        strategy_engine JSONB NOT NULL DEFAULT '{}'::jsonb,
                        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS pna_minute_data (
                        trade_date DATE NOT NULL,
                        time_label TEXT NOT NULL,
                        ts TIMESTAMPTZ,
                        nifty_open DOUBLE PRECISION,
                        nifty_high DOUBLE PRECISION,
                        nifty_low DOUBLE PRECISION,
                        nifty_close DOUBLE PRECISION,
                        minus100_ce BIGINT,
                        minus100_pe BIGINT,
                        atm_ce BIGINT,
                        atm_pe BIGINT,
                        plus100_ce BIGINT,
                        plus100_pe BIGINT,
                        cio_ce BIGINT,
                        cio_pe BIGINT,
                        PRIMARY KEY (trade_date, time_label)
                    )
                """)
        print("[DB] ready", flush=True)
        return True
    except Exception as e:
        print(f"[DB] init error: {type(e).__name__}: {e}", flush=True)
        return False


def _latest_for_time(rows, label):
    if not rows:
        return {}
    if rows[-1].get("time") == label:
        return rows[-1]
    for r in reversed(rows):
        if r.get("time") == label:
            return r
    return {}


def persist_current_minute(force=False):
    global last_persist_label, last_persist_epoch
    if not db_enabled():
        return
    day = state.get("date") or today_key()
    label = minute_label()
    now_epoch = time.time()
    if not force and label == last_persist_label and now_epoch - last_persist_epoch < 10:
        return

    with lock:
        n = _latest_for_time(state["series"]["nifty"], label)
        m = _latest_for_time(state["series"]["minus100"], label)
        a = _latest_for_time(state["series"]["atm"], label)
        p = _latest_for_time(state["series"]["plus100"], label)
        c = _latest_for_time(state["series"]["cio"], label)
        if not any((n, m, a, p, c)):
            return
        expiry = state.get("expiry")
        opening_atm = state.get("opening_atm")
        last_update = state.get("last_update")
        strategies = json.loads(json.dumps(state["strategies"]))
        strategy_engine = json.loads(json.dumps(state["strategy_engine"]))
        ts = n.get("timestamp") or m.get("timestamp") or a.get("timestamp") or p.get("timestamp") or c.get("timestamp")
        vals = (
            day, label, ts,
            n.get("open"), n.get("high"), n.get("low"), n.get("close", n.get("price")),
            m.get("ce"), m.get("pe"), a.get("ce"), a.get("pe"), p.get("ce"), p.get("pe"),
            c.get("ce"), c.get("pe"),
        )

    try:
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO pna_minute_data (
                        trade_date,time_label,ts,nifty_open,nifty_high,nifty_low,nifty_close,
                        minus100_ce,minus100_pe,atm_ce,atm_pe,plus100_ce,plus100_pe,cio_ce,cio_pe
                    ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (trade_date,time_label) DO UPDATE SET
                        ts=EXCLUDED.ts,
                        nifty_open=COALESCE(EXCLUDED.nifty_open,pna_minute_data.nifty_open),
                        nifty_high=COALESCE(EXCLUDED.nifty_high,pna_minute_data.nifty_high),
                        nifty_low=COALESCE(EXCLUDED.nifty_low,pna_minute_data.nifty_low),
                        nifty_close=COALESCE(EXCLUDED.nifty_close,pna_minute_data.nifty_close),
                        minus100_ce=COALESCE(EXCLUDED.minus100_ce,pna_minute_data.minus100_ce),
                        minus100_pe=COALESCE(EXCLUDED.minus100_pe,pna_minute_data.minus100_pe),
                        atm_ce=COALESCE(EXCLUDED.atm_ce,pna_minute_data.atm_ce),
                        atm_pe=COALESCE(EXCLUDED.atm_pe,pna_minute_data.atm_pe),
                        plus100_ce=COALESCE(EXCLUDED.plus100_ce,pna_minute_data.plus100_ce),
                        plus100_pe=COALESCE(EXCLUDED.plus100_pe,pna_minute_data.plus100_pe),
                        cio_ce=COALESCE(EXCLUDED.cio_ce,pna_minute_data.cio_ce),
                        cio_pe=COALESCE(EXCLUDED.cio_pe,pna_minute_data.cio_pe)
                """, vals)
                cur.execute("""
                    INSERT INTO pna_session_meta
                        (trade_date,expiry,opening_atm,last_update,strategies,strategy_engine,updated_at)
                    VALUES (%s,%s,%s,%s,%s,%s,NOW())
                    ON CONFLICT (trade_date) DO UPDATE SET
                        expiry=COALESCE(EXCLUDED.expiry,pna_session_meta.expiry),
                        opening_atm=COALESCE(EXCLUDED.opening_atm,pna_session_meta.opening_atm),
                        last_update=EXCLUDED.last_update,
                        strategies=EXCLUDED.strategies,
                        strategy_engine=EXCLUDED.strategy_engine,
                        updated_at=NOW()
                """, (
                    day, expiry, opening_atm, last_update,
                    Jsonb(strategies), Jsonb(strategy_engine)
                ))
        last_persist_label = label
        last_persist_epoch = now_epoch
    except Exception as e:
        print(f"[DB] persist error: {type(e).__name__}: {e}", flush=True)


def load_history(day):
    if not db_enabled():
        return None
    try:
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT expiry, opening_atm, last_update, strategies, strategy_engine
                    FROM pna_session_meta WHERE trade_date=%s
                """, (day,))
                meta = cur.fetchone()
                if not meta:
                    return None
                cur.execute("""
                    SELECT time_label, ts, nifty_open, nifty_high, nifty_low, nifty_close,
                           minus100_ce, minus100_pe, atm_ce, atm_pe,
                           plus100_ce, plus100_pe, cio_ce, cio_pe
                    FROM pna_minute_data
                    WHERE trade_date=%s
                    ORDER BY time_label
                """, (day,))
                rows = cur.fetchall()

        expiry, opening_atm, last_update, strategies, strategy_engine = meta
        out = {
            "configured": bool(KITE_API_KEY and KITE_API_SECRET),
            "connected": False,
            "message": f"HISTORICAL — {day}",
            "last_update": last_update.isoformat() if hasattr(last_update, "isoformat") else last_update,
            "date": str(day),
            "expiry": str(expiry) if expiry else None,
            "opening_atm": opening_atm,
            "nifty": {"price": None, "open": None, "high": None, "low": None, "previous_close": None},
            "series": {"nifty": [], "atm": [], "minus100": [], "plus100": [], "cio": []},
            "strategies": strategies or {
                "S1": {"trades": [], "active": None},
                "S2": {"trades": [], "active": None},
                "PNA": {"trades": [], "active": None},
            },
            "strategy_engine": strategy_engine or {"last_eval_minute": None, "recent": {}},
            "historical": True,
        }
        for r in rows:
            t, ts, no, nh, nl, nc, mce, mpe, ace, ape, pce, ppe, cce, cpe = r
            ts_s = ts.isoformat() if hasattr(ts, "isoformat") else ts
            if any(v is not None for v in (no, nh, nl, nc)):
                out["series"]["nifty"].append({
                    "time": t, "timestamp": ts_s,
                    "open": no, "high": nh, "low": nl, "close": nc, "price": nc
                })
            if mce is not None or mpe is not None:
                out["series"]["minus100"].append({"time": t, "timestamp": ts_s, "ce": mce, "pe": mpe, "source": "STORED"})
            if ace is not None or ape is not None:
                out["series"]["atm"].append({"time": t, "timestamp": ts_s, "ce": ace, "pe": ape, "source": "STORED"})
            if pce is not None or ppe is not None:
                out["series"]["plus100"].append({"time": t, "timestamp": ts_s, "ce": pce, "pe": ppe, "source": "STORED"})
            if cce is not None or cpe is not None:
                out["series"]["cio"].append({"time": t, "timestamp": ts_s, "ce": cce, "pe": cpe, "source": "STORED"})
        if out["series"]["nifty"]:
            out["nifty"]["price"] = out["series"]["nifty"][-1].get("close")
        return out
    except Exception as e:
        print(f"[DB] load error: {type(e).__name__}: {e}", flush=True)
        return None


def restore_today_from_db():
    hist = load_history(today_key())
    if not hist:
        return
    with lock:
        state["date"] = hist["date"]
        state["expiry"] = hist["expiry"]
        state["opening_atm"] = hist["opening_atm"]
        state["last_update"] = hist["last_update"]
        state["series"] = hist["series"]
        state["strategies"] = hist["strategies"]
        state["strategy_engine"] = hist["strategy_engine"]
        if state["series"]["nifty"]:
            state["nifty"]["price"] = state["series"]["nifty"][-1].get("close")
        state["message"] = "Stored today loaded — Zerodha login required for live feed"
    print(f"[DB] restored today: {sum(len(v) for v in state['series'].values())} points", flush=True)


def import_seed_if_needed():
    if not db_enabled() or not os.path.exists(SEED_FILE):
        return
    try:
        with open(SEED_FILE, "r", encoding="utf-8") as f:
            seed = json.load(f)
        day = seed.get("date")
        if not day:
            return
        if load_history(day):
            print(f"[SEED] {day} already present", flush=True)
            return
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                for r in seed.get("rows", []):
                    cur.execute("""
                        INSERT INTO pna_minute_data (
                            trade_date,time_label,ts,nifty_open,nifty_high,nifty_low,nifty_close,
                            minus100_ce,minus100_pe,atm_ce,atm_pe,plus100_ce,plus100_pe,cio_ce,cio_pe
                        ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (trade_date,time_label) DO NOTHING
                    """, (
                        day, r.get("time"), r.get("timestamp"),
                        None, None, None, r.get("nifty_close"),
                        r.get("minus100_ce"), r.get("minus100_pe"),
                        r.get("atm_ce"), r.get("atm_pe"),
                        r.get("plus100_ce"), r.get("plus100_pe"),
                        r.get("cio_ce"), r.get("cio_pe"),
                    ))
                cur.execute("""
                    INSERT INTO pna_session_meta
                        (trade_date,expiry,opening_atm,last_update,strategies,strategy_engine,updated_at)
                    VALUES (%s,%s,%s,%s,%s,%s,NOW())
                    ON CONFLICT (trade_date) DO NOTHING
                """, (
                    day, seed.get("expiry"), seed.get("opening_atm"),
                    seed.get("rows", [{}])[-1].get("timestamp") if seed.get("rows") else None,
                    Jsonb(seed.get("strategies") or {}),
                    Jsonb({"last_eval_minute": None, "recent": {}}),
                ))
        print(f"[SEED] imported {day}: {len(seed.get('rows', []))} rows", flush=True)
    except Exception as e:
        print(f"[SEED] import error: {type(e).__name__}: {e}", flush=True)


def stored_dates():
    if not db_enabled():
        return []
    try:
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT trade_date FROM pna_session_meta ORDER BY trade_date DESC")
                return [str(r[0]) for r in cur.fetchall()]
    except Exception:
        return []


def backfill_nifty_ohlc(k, day):
    """Replace stored/sampled NIFTY OHLC with Zerodha 1-minute historical candles when available."""
    if not db_enabled() or not nifty_token:
        return False
    try:
        d = datetime.strptime(str(day), "%Y-%m-%d").date()
        start = datetime.combine(d, dt_time(9, 15))
        end = datetime.combine(d, dt_time(15, 30))
        candles = k.historical_data(nifty_token, start, end, "minute") or []
        if not candles:
            return False
        updated = 0
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                for c in candles:
                    cd = c.get("date")
                    label = cd.strftime("%H:%M") if hasattr(cd, "strftime") else str(cd)[11:16]
                    cur.execute("""
                        UPDATE pna_minute_data SET
                            nifty_open=%s,nifty_high=%s,nifty_low=%s,nifty_close=%s
                        WHERE trade_date=%s AND time_label=%s
                    """, (c.get("open"), c.get("high"), c.get("low"), c.get("close"), day, label))
                    updated += cur.rowcount or 0
        if updated:
            print(f"[OHLC] backfilled {day}: {updated} rows", flush=True)
        return updated > 0
    except Exception as e:
        print(f"[OHLC] backfill error {day}: {type(e).__name__}: {e}", flush=True)
        return False


def backfill_missing_ohlc(k):
    if not db_enabled():
        return
    try:
        with db_lock, db_connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT DISTINCT trade_date
                    FROM pna_minute_data
                    WHERE nifty_close IS NOT NULL
                      AND (nifty_open IS NULL OR nifty_high IS NULL OR nifty_low IS NULL)
                    ORDER BY trade_date DESC
                    LIMIT 20
                """)
                days = [str(r[0]) for r in cur.fetchall()]
        for d in days:
            backfill_nifty_ohlc(k, d)
            time.sleep(0.35)
    except Exception as e:
        print(f"[OHLC] scan error: {type(e).__name__}: {e}", flush=True)


# -----------------------
# Zerodha discovery/feed
# -----------------------

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
            for a, b in (("open", "open"), ("high", "high"), ("low", "low"), ("close", "previous_close")):
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
            if n.get("price") is not None:
                update_nifty_minute(float(n["price"]), ts)

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

        persist_current_minute()

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


def reset_for_new_day():
    global locked, latest_oi, baseline_oi, baseline_ready, last_persist_label, last_persist_epoch
    with lock:
        state["date"] = today_key()
        state["expiry"] = None
        state["opening_atm"] = None
        state["last_update"] = None
        state["nifty"] = {"price": None, "open": None, "high": None, "low": None, "previous_close": None}
        state["series"] = {"nifty": [], "atm": [], "minus100": [], "plus100": [], "cio": []}
        state["strategies"] = {
            "S1": {"trades": [], "active": None},
            "S2": {"trades": [], "active": None},
            "PNA": {"trades": [], "active": None},
        }
        state["strategy_engine"] = {"last_eval_minute": None, "recent": {}}
    locked = {}
    latest_oi = {}
    baseline_oi = {}
    baseline_ready = False
    last_persist_label = None
    last_persist_epoch = 0.0


def start_live(token):
    global kite, access_token, baseline_ready
    # Prevent yesterday's in-memory points from carrying into a new trading day.
    if state.get("date") != today_key():
        reset_for_new_day()

    k = KiteConnect(api_key=KITE_API_KEY)
    k.set_access_token(token)
    k.profile()
    discover(k)
    kite = k
    access_token = token

    # Restore current day first if server restarted, then continue the same session.
    restore_today_from_db()

    with lock:
        state["date"] = today_key()
        state["connected"] = True
        state["message"] = "Zerodha login accepted — starting feed"

    lock_strikes(k)
    baseline_ready = False
    threading.Thread(target=build_baseline, daemon=True).start()
    threading.Thread(target=backfill_missing_ohlc, args=(k,), daemon=True).start()
    start_worker()


# -----------------------
# Flask routes
# -----------------------

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
    requested = (request.args.get("date") or "").strip()
    if requested and requested != today_key():
        hist = load_history(requested)
        if hist:
            return jsonify(hist)
        return jsonify({
            "configured": bool(KITE_API_KEY and KITE_API_SECRET),
            "connected": False,
            "message": f"No stored data for {requested}",
            "last_update": None,
            "date": requested,
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
            "historical": True,
        })
    with lock:
        return jsonify(state)


@app.route("/api/dates")
def api_dates():
    return jsonify({"dates": stored_dates(), "today": today_key()})


def snapshot_for_date(day):
    if day == today_key():
        with lock:
            return {
                "date": state.get("date") or today_key(),
                "expiry": state.get("expiry"),
                "opening_atm": state.get("opening_atm"),
                "series": {k: list(state["series"].get(k, [])) for k in ("nifty", "minus100", "atm", "plus100", "cio")},
                "strategies": json.loads(json.dumps(state["strategies"])),
            }
    hist = load_history(day)
    if not hist:
        return None
    return {
        "date": hist.get("date"),
        "expiry": hist.get("expiry"),
        "opening_atm": hist.get("opening_atm"),
        "series": hist.get("series"),
        "strategies": hist.get("strategies"),
    }


@app.route("/api/download/cio")
def download_excel():
    requested = (request.args.get("date") or today_key()).strip()

    # If Zerodha is connected, reconcile OHLC with actual 1-minute candles before export.
    if kite is not None and db_enabled():
        backfill_nifty_ohlc(kite, requested)
        if requested == today_key():
            hist = load_history(requested)
            if hist and hist.get("series", {}).get("nifty"):
                with lock:
                    # Only refresh NIFTY OHLC rows; OI/CIO/strategy live state remains untouched.
                    state["series"]["nifty"] = hist["series"]["nifty"]

    s = snapshot_for_date(requested)
    if not s:
        return jsonify({"error": f"No stored data for {requested}"}), 404

    def by_time(rows):
        return {str(r.get("time")): r for r in rows if r.get("time")}

    maps = {k: by_time(v) for k, v in s["series"].items()}
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
        "Time", "NIFTY Open", "NIFTY High", "NIFTY Low", "NIFTY Close",
        "ATM -100 CE OI", "ATM -100 PE OI",
        "ATM CE OI", "ATM PE OI",
        "ATM +100 CE OI", "ATM +100 PE OI",
        "CIO CE Negative Change in OI", "CIO PE Negative Change in OI",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(8, c, h).font = Font(bold=True)
        ws.cell(8, c).alignment = Alignment(horizontal="center")

    for rno, t in enumerate(all_times, 9):
        n = maps["nifty"].get(t, {})
        m = maps["minus100"].get(t, {})
        a = maps["atm"].get(t, {})
        p = maps["plus100"].get(t, {})
        c = maps["cio"].get(t, {})
        vals = [
            t,
            n.get("open"), n.get("high"), n.get("low"), n.get("close", n.get("price")),
            m.get("ce"), m.get("pe"), a.get("ce"), a.get("pe"), p.get("ce"), p.get("pe"),
            c.get("ce"), c.get("pe")
        ]
        for col, val in enumerate(vals, 1):
            ws.cell(rno, col, val)

    ws.freeze_panes = "A9"
    widths = [12, 14, 14, 14, 14, 19, 19, 18, 18, 19, 19, 31, 31]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    # Strategy summary report sheet
    ss = wb.create_sheet("S1 S2 PNA Summary")
    ss["A1"] = "Pratik Analysis"
    ss["A2"] = "S1 / S2 / PNA Strategy Summary Report"
    ss["A1"].font = Font(bold=True, size=16)
    ss["A2"].font = Font(bold=True, size=13)
    ss["A4"] = "Date"; ss["B4"] = s["date"]

    summary_headers = ["Strategy", "Total Trades", "Profitable", "Losing", "Total Points"]
    for c, h in enumerate(summary_headers, 1):
        ss.cell(6, c, h).font = Font(bold=True)

    strategies = s.get("strategies") or {}
    summary_row = 7
    all_trade_rows = []
    for name in ("S1", "S2", "PNA"):
        st = strategies.get(name) or {"trades": [], "active": None}
        trades = list(st.get("trades") or [])
        wins = sum(1 for t in trades if float(t.get("points") or 0) >= 0)
        losses = sum(1 for t in trades if float(t.get("points") or 0) < 0)
        total_points = round(sum(float(t.get("points") or 0) for t in trades), 2)
        vals = [name, len(trades), wins, losses, total_points]
        for c, v in enumerate(vals, 1):
            ss.cell(summary_row, c, v)
        summary_row += 1

        for idx, t in enumerate(trades, 1):
            pts = t.get("points")
            result = "PROFIT" if pts is not None and float(pts) >= 0 else "LOSS"
            all_trade_rows.append([
                name, idx, t.get("type"), t.get("entry_time"), t.get("entry_level"), t.get("sl_level"),
                t.get("exit_time"), t.get("exit_level"), t.get("exit_reason"), pts, result
            ])
        active = st.get("active")
        if active:
            all_trade_rows.append([
                name, len(trades) + 1, active.get("type"), active.get("entry_time"), active.get("entry_level"), active.get("sl_level"),
                None, None, "OPEN", None, "OPEN"
            ])

    detail_start = 12
    detail_headers = [
        "Strategy", "Trade #", "Type", "Entry Time", "Entry NIFTY", "SL Level",
        "Exit Time", "Exit NIFTY", "Exit Reason", "Points", "Result"
    ]
    for c, h in enumerate(detail_headers, 1):
        ss.cell(detail_start, c, h).font = Font(bold=True)
    if all_trade_rows:
        for rno, row in enumerate(all_trade_rows, detail_start + 1):
            for c, v in enumerate(row, 1):
                ss.cell(rno, c, v)
    else:
        ss.cell(detail_start + 1, 1, "No strategy trades for this date")

    ss.freeze_panes = f"A{detail_start + 1}"
    ss_widths = [14, 10, 10, 14, 16, 16, 14, 16, 30, 12, 12]
    for i, w in enumerate(ss_widths, 1):
        ss.column_dimensions[chr(64 + i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
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
        "database": db_enabled(),
        "stored_dates": stored_dates(),
        "points": {k: len(state["series"][k]) for k in ("atm", "minus100", "plus100", "cio")},
    })


# Safe startup: initialize DB, seed 08-09-2026 if absent, restore today if present.
init_db()
import_seed_if_needed()
restore_today_from_db()
start_worker()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=False)
