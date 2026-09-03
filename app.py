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


# ============================================================
# RUNTIME STATE
# ============================================================

state = {

    "configured":
        bool(KITE_API_KEY and KITE_API_SECRET),

    "connected":
        False,

    "message":
        "Waiting for Zerodha login",

    "last_update":
        None,

    "date":
        None,

    "expiry":
        None,

    "nifty": {

        "price":
            None,

        "previous_close":
            None,

        "open":
            None,

        "high":
            None,

        "low":
            None,

        "change":
            None,

        "change_pct":
            None,
    },

    "vix": {

        "price":
            None,

        "range":
            None,

        "interpretation":
            None,
    },

    "zone": {

        "quadrant":
            None,

        "zone":
            None,

        "opening_pct":
            None,

        "levels":
            {},
    },

    "opening_atm":
        None,

    "oic": {

        "atm":
            None,

        "minus100":
            None,

        "plus100":
            None,
    },

    "series": {

        "atm":
            [],

        "minus100":
            [],

        "plus100":
            [],

        "cio":
            [],
    },

    "history_dates":
        [],
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

    current =
        n.hour * 60 + n.minute

    start =
        MARKET_START_HOUR * 60 + MARKET_START_MINUTE

    end =
        MARKET_END_HOUR * 60 + MARKET_END_MINUTE

    return start <= current <= end
