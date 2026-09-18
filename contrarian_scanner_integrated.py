import html
import json
import logging
import os
import smtplib
import sys
import time
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO

import pandas as pd
import pandas_market_calendars as mcal
import requests
import yfinance as yf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# Configuratie
# ============================================================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")
TOP_N = int(os.getenv("TOP_N", "5"))
PERIOD = os.getenv("PERIOD", "1y")
MIN_HISTORY = 200
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "90"))

SIGNAAL_DREMPEL = 2.0
ATR_STOP_BUFFER = 1.0

# Retry-instellingen voor de yfinance-download (aandelen)
DOWNLOAD_MAX_POGINGEN = 3
DOWNLOAD_BACKOFF_SECONDEN = 5  # verdubbelt elke poging

# Bitvavo-specifiek
BITVAVO_API_BASE = "https://api.bitvavo.com/v2"
CANDLE_LIMIT = int(os.getenv("CANDLE_LIMIT", "300"))
VERTRAGING_TUSSEN_REQUESTS = 0.3  # seconden, uit beleefdheid richting Bitvavo's rate limit

# ---- Vier universes ----
AEX_TICKERS = [
    "ASML.AS", "SHELL.AS", "UNA.AS", "PRX.AS", "INGA.AS", "REN.AS",
    "ASM.AS", "MT.AS", "HEIA.AS", "AD.AS", "ADYEN.AS", "ABN.AS",
    "PHIA.AS", "NN.AS", "BESI.AS", "KPN.AS", "WKL.AS", "ASRNL.AS",
    "AGN.AS", "DSFIR.AS", "UMG.AS", "EXO.AS", "SBMO.AS", "IMCD.AS",
    "AALB.AS",
]

DOW_TICKERS = [
    "AAPL", "AMGN", "AMZN", "AXP", "BA", "CAT", "CRM", "CSCO", "CVX", "DIS",
    "GS", "HD", "HON", "IBM", "JNJ", "JPM", "KO", "MCD", "MMM", "MRK",
    "MSFT", "NKE", "NVDA", "PG", "SHW", "TRV", "UNH", "V", "VZ", "WMT"
]

NASDAQ_TICKERS = [
    "AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "GOOG", "META", "TSLA", "AVGO", "AMD",
    "INTC", "QCOM", "TXN", "AMAT", "LRCX", "KLAC", "ADI", "MU", "ADBE", "ORCL",
    "NOW", "PANW", "NFLX", "COST", "PEP", "SBUX", "BKNG", "ABNB", "PYPL", "COIN",
    "SHOP", "UBER", "CRM", "CSCO"
]

# Bitvavo: alleen gevestigde, liquide assets op de REGULIERE beurs
# (klassiek orderboek via de publieke /candles-API). GEEN memecoins (DOGE
# e.d. bewust weggelaten) en GEEN Web3-wallet-only tokens — die laatste
# lopen sowieso niet via deze API: de Web3-wallet is een apart, non-custodiaal
# DeFi/swap-onderdeel (15.000+ tokens, niet door Bitvavo geverifieerd) dat
# volledig losstaat van de gereguleerde beurs-orderboeken die /candles
# ontsluit. Door uitsluitend deze API te gebruiken, is die scheiding al
# geborgd; de onderstaande lijst bevat daarnaast bewust alleen bekende,
# lang gevestigde munten — geen nieuwe/speculatieve kleine caps.
BITVAVO_MARKETS = [
    "BTC-EUR", "ETH-EUR", "XRP-EUR", "ADA-EUR", "SOL-EUR",
    "DOT-EUR", "LTC-EUR", "LINK-EUR", "XLM-EUR", "ATOM-EUR", "ALGO-EUR",
    "AVAX-EUR", "UNI-EUR", "AAVE-EUR", "FIL-EUR", "NEAR-EUR", "ICP-EUR",
    "ETC-EUR", "BCH-EUR", "TRX-EUR", "VET-EUR", "XTZ-EUR", "EGLD-EUR",
    "EOS-EUR", "MATIC-EUR",
]

# (naam, tickerlijst, valutasymbool)
STOCK_UNIVERSES = [
    ("AEX", AEX_TICKERS, "€"),
    ("Dow", DOW_TICKERS, "$"),
    ("Nasdaq", NASDAQ_TICKERS, "$"),
]
ALLE_UNIVERSE_NAMEN = ["AEX", "Dow", "Nasdaq", "Bitvavo"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


class DataOphalenMislukt(Exception):
    """Wordt gegooid als de koersdata na alle retries niet opgehaald kon worden."""
    pass


def is_trading_day(calendar_code, date=None):
    date = date or datetime.now(timezone.utc).date()
    return not mcal.get_calendar(calendar_code).schedule(start_date=date, end_date=date).empty


def get_ticker_frame(data, ticker, only_one_ticker):
    try:
        if only_one_ticker or not isinstance(data.columns, pd.MultiIndex):
            df = data.copy()
        else:
            if ticker not in data.columns.get_level_values(0):
                return None
            df = data[ticker].copy()
        df = df.dropna(how="all")
        if df.empty:
            return None
        df.columns = [str(c).lower() for c in df.columns]
        if "adj close" in df.columns and "close" not in df.columns:
            df["close"] = df["adj close"]
        return df
    except Exception:
        return None


def add_indicators(df, jaar_venster=252, jaar_min_periods=100):
    """
    Indicatoren gericht op CONTRAIR beleggen — werkt voor zowel aandelen als
    crypto. `jaar_venster` is 252 (handelsdagen) voor aandelen en 365
    (kalenderdagen) voor crypto, dat 24/7 verhandeld wordt.
    """
    df = df.copy()

    delta = df["close"].diff()
    gain = delta.where(delta > 0, 0.0)
    loss = -delta.where(delta < 0, 0.0)
    avg_gain = gain.ewm(alpha=1 / 14, min_periods=14, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / 14, min_periods=14, adjust=False).mean()
    rs = avg_gain / avg_loss
    df["rsi_14"] = 100 - (100 / (1 + rs))

    df["sma_20"] = df["close"].rolling(20).mean()
    df["std_20"] = df["close"].rolling(20).std()
    df["bb_upper"] = df["sma_20"] + 2 * df["std_20"]
    df["bb_lower"] = df["sma_20"] - 2 * df["std_20"]
    df["percent_b"] = (df["close"] - df["bb_lower"]) / (df["bb_upper"] - df["bb_lower"])

    df["sma_50"] = df["close"].rolling(50).mean()
    df["std_50"] = df["close"].rolling(50).std()
    df["z_score_50"] = (df["close"] - df["sma_50"]) / df["std_50"]

    df["hoog_jaar"] = df["high"].rolling(jaar_venster, min_periods=jaar_min_periods).max()
    df["laag_jaar"] = df["low"].rolling(jaar_venster, min_periods=jaar_min_periods).min()
    df["afstand_tot_laag_jaar_pct"] = (df["close"] - df["laag_jaar"]) / df["laag_jaar"] * 100
    df["afstand_tot_hoog_jaar_pct"] = (df["hoog_jaar"] - df["close"]) / df["hoog_jaar"] * 100

    df["vol_sma"] = df["volume"].rolling(20).mean()
    df["rel_volume"] = df["volume"] / df["vol_sma"]
    df["dagrendement_pct"] = df["close"].pct_change() * 100

    prev_close = df["close"].shift(1)
    true_range = pd.concat([
        df["high"] - df["low"],
        (df["high"] - prev_close).abs(),
        (df["low"] - prev_close).abs(),
    ], axis=1).max(axis=1)
    df["atr_14"] = true_range.ewm(alpha=1 / 14, min_periods=14, adjust=False).mean()

    df["laag_10d"] = df["low"].rolling(10).min()
    df["hoog_10d"] = df["high"].rolling(10).max()

    return df


def _bereken_target_en_stop(latest, direction):
    entry = latest.get("close")
    sma_20 = latest.get("sma_20")
    atr = latest.get("atr_14")

    if direction == "Koop":
        laag_10d = latest.get("laag_10d")
        if pd.isna(entry) or pd.isna(sma_20) or pd.isna(atr) or pd.isna(laag_10d):
            return "N/A", "N/A", "N/A"
        target = sma_20
        stop = min(laag_10d, entry) - ATR_STOP_BUFFER * atr
        if stop >= entry or target <= entry:
            return "N/A", "N/A", "N/A"
        risico = entry - stop
        beloning = target - entry
        rr = round(beloning / risico, 2) if risico > 0 else "N/A"
        return round(float(target), 4), round(float(stop), 4), rr

    elif direction == "Verkoop":
        hoog_10d = latest.get("hoog_10d")
        if pd.isna(entry) or pd.isna(sma_20) or pd.isna(atr) or pd.isna(hoog_10d):
            return "N/A", "N/A", "N/A"
        target = sma_20
        stop = max(hoog_10d, entry) + ATR_STOP_BUFFER * atr
        if stop <= entry or target >= entry:
            return "N/A", "N/A", "N/A"
        risico = stop - entry
        beloning = entry - target
        rr = round(beloning / risico, 2) if risico > 0 else "N/A"
        return round(float(target), 4), round(float(stop), 4), rr

    return "N/A", "N/A", "N/A"


def score_signal(df, jaar_label="jaar", jaar_nabijheid_pct=3, dagrendement_drempel=3):
    """
    Contraire scoring — gedeeld tussen aandelen en crypto. `jaar_label` is
    alleen voor de leesbare reden-tekst ("52-weken" vs "365-dagen");
    `jaar_nabijheid_pct` en `dagrendement_drempel` liggen iets ruimer voor
    crypto (van nature volatieler) dan voor aandelen.
    """
    if len(df) < 60:
        return None

    latest = df.iloc[-1]
    score = 0.0
    reasons = []

    rsi = latest.get("rsi_14")
    if pd.isna(rsi):
        rsi = 50

    if rsi < 25:
        score += 2.2
        reasons.append(f"RSI extreem oversold ({rsi:.1f}) — mogelijke capitulatie")
    elif rsi < 32:
        score += 1.0
        reasons.append(f"RSI oversold ({rsi:.1f})")
    elif rsi > 75:
        score -= 2.2
        reasons.append(f"RSI extreem overbought ({rsi:.1f}) — mogelijke euforie-top")
    elif rsi > 68:
        score -= 1.0
        reasons.append(f"RSI overbought ({rsi:.1f})")

    percent_b = latest.get("percent_b")
    if pd.notna(percent_b):
        if percent_b < 0:
            score += 1.5
            reasons.append("Koers doorbreekt de onderkant van de Bollinger Band")
        elif percent_b > 1:
            score -= 1.5
            reasons.append("Koers doorbreekt de bovenkant van de Bollinger Band")

    z_score = latest.get("z_score_50")
    if pd.notna(z_score):
        if z_score <= -2.0:
            score += 1.8
            reasons.append(f"Koers {abs(z_score):.1f}σ onder 50-daags gemiddelde (statistische uitschieter)")
        elif z_score >= 2.0:
            score -= 1.8
            reasons.append(f"Koers {z_score:.1f}σ boven 50-daags gemiddelde (statistische uitschieter)")

    afstand_laag = latest.get("afstand_tot_laag_jaar_pct")
    afstand_hoog = latest.get("afstand_tot_hoog_jaar_pct")
    if pd.notna(afstand_laag) and afstand_laag <= jaar_nabijheid_pct:
        score += 1.0
        reasons.append(f"Koers nabij {jaar_label}-dieptepunt")
    if pd.notna(afstand_hoog) and afstand_hoog <= jaar_nabijheid_pct:
        score -= 1.0
        reasons.append(f"Koers nabij {jaar_label}-hoogtepunt")

    rel_vol = latest.get("rel_volume")
    dagrendement = latest.get("dagrendement_pct")
    if pd.notna(rel_vol) and pd.notna(dagrendement) and rel_vol > 1.8:
        if dagrendement < -dagrendement_drempel:
            score += 0.8
            reasons.append("Paniekvolume op een sterk dalende dag (mogelijke uitverkoop-bodem)")
        elif dagrendement > dagrendement_drempel:
            score -= 0.8
            reasons.append("Uitzonderlijk volume op een sterk stijgende dag (mogelijke euforie-top)")

    if score >= SIGNAAL_DREMPEL:
        direction = "Koop"
    elif score <= -SIGNAAL_DREMPEL:
        direction = "Verkoop"
    else:
        direction = "Neutraal"

    target, stop, rr = _bereken_target_en_stop(latest, direction)

    return {
        "Contrarian Score": round(score, 2),
        "Direction": direction,
        "RSI": round(float(rsi), 1),
        "%B (Bollinger)": round(float(percent_b), 2) if pd.notna(percent_b) else "N/A",
        "Z-score(50)": round(float(z_score), 2) if pd.notna(z_score) else "N/A",
        "Rel. Volume": round(float(rel_vol), 2) if pd.notna(rel_vol) else "N/A",
        "Koersdoel": target,
        "Stop-loss": stop,
        "Risk/Reward": rr,
        "Reasons": " | ".join(reasons) if reasons else "Geen extreme afwijking",
    }


def safe_round(val, digits=2):
    try:
        return round(float(val), digits)
    except Exception:
        return "N/A"


# ---------------------------------------------------------------
# Aandelen: bulk-download via yfinance
# ---------------------------------------------------------------
def download_met_retry(tickers):
    laatste_fout = None
    wachttijd = DOWNLOAD_BACKOFF_SECONDEN

    for poging in range(1, DOWNLOAD_MAX_POGINGEN + 1):
        try:
            data = yf.download(
                tickers, period=PERIOD, interval="1d", group_by="ticker",
                auto_adjust=False, threads=True, progress=False,
            )
            if data is None or data.empty:
                raise ValueError("yfinance gaf een lege dataset terug")
            return data
        except Exception as e:
            laatste_fout = e
            logger.warning(f"Download poging {poging}/{DOWNLOAD_MAX_POGINGEN} mislukt: {e}")
            if poging < DOWNLOAD_MAX_POGINGEN:
                time.sleep(wachttijd)
                wachttijd *= 2

    raise DataOphalenMislukt(f"Download definitief mislukt na {DOWNLOAD_MAX_POGINGEN} pogingen: {laatste_fout}")


def process_stock_universe(tickers, data, universe_name, currency_symbol, only_one_ticker):
    results = []
    for ticker in tickers:
        try:
            df = get_ticker_frame(data, ticker, only_one_ticker)
            if df is None or len(df) < MIN_HISTORY:
                continue
            df = add_indicators(df.reset_index(drop=True), jaar_venster=252, jaar_min_periods=100)
            signal = score_signal(df, jaar_label="52-weken", jaar_nabijheid_pct=3, dagrendement_drempel=3)
            if not signal:
                continue
            latest = df.iloc[-1]
            results.append({
                "Ticker": ticker,
                "Universe": universe_name,
                "Valuta": currency_symbol,
                "Koers": safe_round(latest.get("close")),
                **signal
            })
        except Exception as e:
            logger.warning(f"Fout bij {ticker}: {e}")

    buys = [r for r in results if r["Direction"] == "Koop"]
    sells = [r for r in results if r["Direction"] == "Verkoop"]
    buys.sort(key=lambda x: x["Contrarian Score"], reverse=True)
    sells.sort(key=lambda x: x["Contrarian Score"])
    return buys[:TOP_N], sells[:TOP_N]


# ---------------------------------------------------------------
# Bitvavo: per-market REST-aanroepen (geen bulk-endpoint beschikbaar)
# ---------------------------------------------------------------
def haal_candles_op(market, interval="1d", limit=CANDLE_LIMIT):
    url = f"{BITVAVO_API_BASE}/{market}/candles"
    params = {"interval": interval, "limit": limit}

    laatste_fout = None
    wachttijd = DOWNLOAD_BACKOFF_SECONDEN

    for poging in range(1, DOWNLOAD_MAX_POGINGEN + 1):
        try:
            resp = requests.get(url, params=params, timeout=20)
            resp.raise_for_status()
            data = resp.json()
            if not isinstance(data, list) or not data:
                raise ValueError("Bitvavo gaf een lege of onverwachte candle-response terug")

            df = pd.DataFrame(data, columns=["timestamp", "open", "high", "low", "close", "volume"])
            df = df.astype({
                "open": "float64", "high": "float64", "low": "float64",
                "close": "float64", "volume": "float64",
            })
            # Bitvavo levert candles nieuwste-eerst; wij hebben oudste-eerst nodig.
            df = df.sort_values("timestamp").reset_index(drop=True)
            return df

        except Exception as e:
            laatste_fout = e
            logger.warning(f"[{market}] Poging {poging}/{DOWNLOAD_MAX_POGINGEN} mislukt: {e}")
            if poging < DOWNLOAD_MAX_POGINGEN:
                time.sleep(wachttijd)
                wachttijd *= 2

    raise DataOphalenMislukt(f"[{market}] Definitief mislukt na {DOWNLOAD_MAX_POGINGEN} pogingen: {laatste_fout}")


def process_bitvavo_universe(markets, currency_symbol):
    results = []
    aantal_geslaagd = 0

    for market in markets:
        try:
            df = haal_candles_op(market)
            if df is None or len(df) < MIN_HISTORY:
                logger.info(f"[{market}] Onvoldoende historie, overgeslagen")
                continue
            aantal_geslaagd += 1

            df = add_indicators(df, jaar_venster=365, jaar_min_periods=150)
            signal = score_signal(df, jaar_label="365-dagen", jaar_nabijheid_pct=5, dagrendement_drempel=5)
            if not signal:
                continue
            latest = df.iloc[-1]
            results.append({
                "Ticker": market,
                "Universe": "Bitvavo",
                "Valuta": currency_symbol,
                "Koers": safe_round(latest.get("close"), 4),
                **signal
            })
        except DataOphalenMislukt as e:
            logger.warning(f"Market overgeslagen: {e}")
        except Exception as e:
            logger.warning(f"Fout bij {market}: {e}")
        finally:
            time.sleep(VERTRAGING_TUSSEN_REQUESTS)

    if aantal_geslaagd == 0:
        raise DataOphalenMislukt("Voor geen enkele Bitvavo-market kon data opgehaald worden")

    buys = [r for r in results if r["Direction"] == "Koop"]
    sells = [r for r in results if r["Direction"] == "Verkoop"]
    buys.sort(key=lambda x: x["Contrarian Score"], reverse=True)
    sells.sort(key=lambda x: x["Contrarian Score"])
    logger.info(f"Bitvavo → {len(buys)} contraire koop / {len(sells)} contraire verkoop "
                f"(op basis van {aantal_geslaagd}/{len(markets)} markets)")
    return buys[:TOP_N], sells[:TOP_N]


def fetch_all(aex_open, us_open):
    """
    Verwerkt alle vier universes. AEX wordt overgeslagen als Euronext dicht
    is; Dow/Nasdaq worden overgeslagen als de NYSE dicht is. Bitvavo wordt
    ALTIJD verwerkt (crypto handelt 24/7). Overgeslagen universes krijgen
    None i.p.v. lege lijsten, zodat de rapportage "markt gesloten" kan tonen
    i.p.v. "geen signalen gevonden".
    """
    resultaten = {}

    actieve_stock_universes = []
    if aex_open:
        actieve_stock_universes.append(("AEX", AEX_TICKERS, "€"))
    else:
        logger.info("Euronext gesloten — AEX wordt overgeslagen")
        resultaten["aex_buys"], resultaten["aex_sells"] = None, None

    if us_open:
        actieve_stock_universes.append(("Dow", DOW_TICKERS, "$"))
        actieve_stock_universes.append(("Nasdaq", NASDAQ_TICKERS, "$"))
    else:
        logger.info("NYSE gesloten — Dow en Nasdaq worden overgeslagen")
        resultaten["dow_buys"], resultaten["dow_sells"] = None, None
        resultaten["nasdaq_buys"], resultaten["nasdaq_sells"] = None, None

    if actieve_stock_universes:
        alle_stock_tickers = list(set(t for _, tickers, _ in actieve_stock_universes for t in tickers))
        logger.info(f"Data ophalen voor {len(alle_stock_tickers)} unieke aandelen-tickers...")
        data = download_met_retry(alle_stock_tickers)  # gooit DataOphalenMislukt bij falen
        only_one_ticker = len(alle_stock_tickers) == 1

        for naam, tickers, valuta in actieve_stock_universes:
            buys, sells = process_stock_universe(tickers, data, naam, valuta, only_one_ticker)
            resultaten[f"{naam.lower()}_buys"] = buys
            resultaten[f"{naam.lower()}_sells"] = sells
            logger.info(f"{naam:8s} → {len(buys)} contraire koop / {len(sells)} contraire verkoop")

    # Bitvavo altijd verwerken
    bv_buys, bv_sells = process_bitvavo_universe(BITVAVO_MARKETS, "€")  # gooit DataOphalenMislukt bij totale falen
    resultaten["bitvavo_buys"] = bv_buys
    resultaten["bitvavo_sells"] = bv_sells

    return resultaten


def analyze_with_ai(resultaten):
    if not DEEPSEEK_API_KEY:
        return "Geen AI-analyse (API key ontbreekt)."

    system_prompt = (
        "Je bent een ervaren contraire belegger ('contrarian investor') actief op de "
        "AEX, Dow 30, Nasdaq en Bitvavo (crypto). Je zoekt bewust extremen tegen de "
        "heersende marktstemming in: paniek en capitulatie als koopkans, euforie en "
        "overdreven optimisme als verkoopkans. Elk signaal bevat al een berekend "
        "koersdoel, stop-loss en risk/reward — reken deze niet zelf opnieuw uit, maar "
        "interpreteer en beoordeel ze. Je antwoordt uitsluitend in het Nederlands, "
        "volgt exact de gevraagde structuur, en waarschuwt expliciet dat een contraire "
        "positie tegen de trend in kan lopen ('probeer geen vallend mes te vangen') — "
        "en dat dit risico bij crypto nog groter is dan bij aandelen."
    )

    def sectie(key):
        val = resultaten.get(key)
        return json.dumps(val, indent=2, ensure_ascii=False) if val is not None else '"Markt was vandaag gesloten"'

    prompt = f"""
Hier zijn de contraire signalen voor vier universes. Elk signaal bevat al een
berekend koersdoel (reversie naar 20-daags gemiddelde), stop-loss
(ATR-gecorrigeerd t.o.v. het recente 10-daagse swing-niveau) en Risk/Reward-ratio.

=== AEX ===
Koopkansen: {sectie('aex_buys')}
Verkoopkansen: {sectie('aex_sells')}

=== DOW 30 ===
Koopkansen: {sectie('dow_buys')}
Verkoopkansen: {sectie('dow_sells')}

=== NASDAQ ===
Koopkansen: {sectie('nasdaq_buys')}
Verkoopkansen: {sectie('nasdaq_sells')}

=== BITVAVO (crypto) ===
Koopkansen: {sectie('bitvavo_buys')}
Verkoopkansen: {sectie('bitvavo_sells')}

Schrijf een gestructureerd rapport in het Nederlands met EXACT deze opbouw:

## 1. AEX – Beste Contraire Kansen
## 2. Dow 30 – Beste Contraire Kansen
## 3. Nasdaq – Beste Contraire Kansen
## 4. Bitvavo – Beste Contraire Kansen

Voor elke sectie (indien de markt open was): de 1-2 sterkste koop- én
verkoopsignalen — ticker, koers, status (KOPEN/VERKOPEN, contrair), waarom dit
overdreven pessimisme/optimisme lijkt, en een beoordeling van het gegeven
koersdoel/stop-loss/risk-reward. Was de markt gesloten, meld dat kort.

## 5. Marktobservatie
Eén korte observatie: waar zitten de meeste extremen vandaag, en zegt het
verschil tussen aandelen en crypto iets over de bredere risicobereidheid?

## 6. Waarschuwing
Een korte, nuchtere waarschuwing over de risico's van contrair beleggen —
en specifiek de extra risico's bij crypto (hogere volatiliteit, dunnere
liquiditeit, mogelijke structurele project-risico's i.p.v. tijdelijk sentiment).

BELANGRIJK:
- Geef NOOIT tegenstrijdige adviezen.
- Herbereken koersdoel/stop-loss niet zelf — beoordeel de gegeven waarden.
- Wees nuchter en realistisch. Geen garanties.
"""

    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            json={
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.15,
            },
            headers={
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json",
            },
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        logger.error(f"AI mislukt: {e}")
        return f"AI-analyse mislukt: {e}"


def _stijl_tabel_header(table, achtergrondkleur):
    for cell in table.rows[0].cells:
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), achtergrondkleur)
        cell._tc.get_or_add_tcPr().append(shading)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(9)


def _voeg_signalen_tabel_toe(doc, titel, resultaten, markt_gesloten_tekst, geen_data_tekst, header_kleur):
    doc.add_heading(titel, level=2)
    if resultaten is None:
        p = doc.add_paragraph(markt_gesloten_tekst)
        p.italic = True
        return
    if not resultaten:
        p = doc.add_paragraph(geen_data_tekst)
        p.italic = True
        return

    headers = ["Ticker", "Koers", "RSI", "%B", "Z-score", "Koersdoel", "Stop-loss", "R/R", "Score"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
    _stijl_tabel_header(table, header_kleur)

    for r in resultaten:
        row = table.add_row().cells
        row[0].text = r["Ticker"]
        row[1].text = f"{r['Valuta']}{r['Koers']}"
        row[2].text = str(r["RSI"])
        row[3].text = str(r["%B (Bollinger)"])
        row[4].text = str(r["Z-score(50)"])
        row[5].text = f"{r['Valuta']}{r['Koersdoel']}" if r["Koersdoel"] != "N/A" else "N/A"
        row[6].text = f"{r['Valuta']}{r['Stop-loss']}" if r["Stop-loss"] != "N/A" else "N/A"
        row[7].text = str(r["Risk/Reward"])
        row[8].text = str(r["Contrarian Score"])


def create_docx_report(resultaten, ai_report):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    title = doc.add_heading(f"Contrarian Scanner (AEX / Dow / Nasdaq / Bitvavo) – {datetime.now():%d-%m-%Y}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Gegenereerd: {datetime.now():%d-%m-%Y %H:%M UTC}")
    doc.add_paragraph(
        "Focus: contraire signalen met berekend koersdoel (reversie naar 20-daags "
        "gemiddelde) en ATR-gecorrigeerde stop-loss — aandelen + crypto (Bitvavo, "
        "hoofdbeurs, geen memecoins)"
    )

    for naam in ALLE_UNIVERSE_NAMEN:
        key = naam.lower()
        _voeg_signalen_tabel_toe(
            doc, f"{naam} – Contraire Koopkansen (overdreven pessimisme)",
            resultaten[f"{key}_buys"],
            f"{naam} was vandaag gesloten.",
            f"Geen extreme oversold-signalen gevonden voor {naam} vandaag.", "2E7D32")
        _voeg_signalen_tabel_toe(
            doc, f"{naam} – Contraire Verkoopkansen (overdreven optimisme)",
            resultaten[f"{key}_sells"],
            f"{naam} was vandaag gesloten.",
            f"Geen extreme overbought-signalen gevonden voor {naam} vandaag.", "C62828")

    doc.add_heading("DeepSeek Analyse – Contraire Kansen + Risicowaarschuwing", level=2)
    doc.add_paragraph(ai_report)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def send_email(resultaten, ai_report):
    if not all([SENDER_EMAIL, SENDER_PASSWORD, RECEIVER_EMAIL]):
        logger.error("E-mailgegevens ontbreken (SENDER_EMAIL / SENDER_PASSWORD / RECEIVER_EMAIL)")
        return False

    receivers = [e.strip() for e in RECEIVER_EMAIL.split(",") if e.strip()]
    if not receivers:
        logger.error("RECEIVER_EMAIL bevat geen geldige adressen")
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"Contrarian Scanner (AEX/Dow/Nasdaq/Bitvavo) – {datetime.now():%d-%m-%Y}"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(receivers)

    def maak_tabel(data, universe_naam):
        if data is None:
            return f"<p><em>{universe_naam} was vandaag gesloten.</em></p>"
        if not data:
            return "<p>Geen extreme signalen</p>"
        df = pd.DataFrame(data)
        df["Koers"] = df.apply(lambda r: f"{r['Valuta']}{r['Koers']}", axis=1)
        df["Koersdoel"] = df.apply(lambda r: f"{r['Valuta']}{r['Koersdoel']}" if r["Koersdoel"] != "N/A" else "N/A", axis=1)
        df["Stop-loss"] = df.apply(lambda r: f"{r['Valuta']}{r['Stop-loss']}" if r["Stop-loss"] != "N/A" else "N/A", axis=1)
        kolommen = ["Ticker", "Koers", "RSI", "Koersdoel", "Stop-loss", "Risk/Reward", "Contrarian Score", "Reasons"]
        return df[kolommen].to_html(index=False, escape=True)

    secties = ""
    for naam in ALLE_UNIVERSE_NAMEN:
        key = naam.lower()
        secties += f"""
        <h3 style="color:green">{naam} – Contraire Koopkansen</h3>
        {maak_tabel(resultaten[f'{key}_buys'], naam)}
        <h3 style="color:red">{naam} – Contraire Verkoopkansen</h3>
        {maak_tabel(resultaten[f'{key}_sells'], naam)}
        """

    body = f"""
    <html><body style="font-family:Arial,sans-serif;line-height:1.5">
    <h2>Contrarian Scanner (AEX / Dow / Nasdaq / Bitvavo) – {datetime.now():%d-%m-%Y}</h2>
    {secties}
    <h3>AI Analyse + Risicowaarschuwing</h3>
    <div>{html.escape(ai_report).replace(chr(10), "<br>")}</div>

    <p style="color:#666;font-size:12px;margin-top:25px">
    Dit is geen beleggingsadvies. Koersdoel en stop-loss zijn statistisch afgeleid
    en geen garantie. Contrair beleggen kent verhoogd risico, bij crypto nog meer
    dan bij aandelen.
    </p>
    </body></html>
    """
    msg.attach(MIMEText(body, "html"))

    try:
        attachment = MIMEBase("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
        attachment.set_payload(create_docx_report(resultaten, ai_report).read())
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", "attachment",
                              filename=f"Contrarian_Scanner_{datetime.now():%Y%m%d}.docx")
        msg.attach(attachment)
    except Exception as e:
        logger.warning(f"DOCX mislukt: {e}")

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, receivers, msg.as_string())
        logger.info(f"E-mail verzonden naar: {', '.join(receivers)}")
        return True
    except Exception as e:
        logger.error(f"E-mail mislukt: {e}")
        return False


def main():
    logger.info("=" * 60)
    logger.info("START Contrarian Scanner (AEX / Dow / Nasdaq / Bitvavo)")
    logger.info("=" * 60)

    is_manual = os.getenv("GITHUB_EVENT_NAME") == "workflow_dispatch"
    aex_open = True if is_manual else is_trading_day("XAMS")
    us_open = True if is_manual else is_trading_day("NYSE")

    try:
        resultaten = fetch_all(aex_open, us_open)
    except DataOphalenMislukt as e:
        logger.critical(f"Kon geen koersdata ophalen: {e}")
        sys.exit(1)

    verwerkte_resultaten = [v for v in resultaten.values() if v is not None]
    if not any(verwerkte_resultaten):
        # Bij contrair beleggen is "geen extremen" een normale, verwachte
        # uitkomst op een rustige dag over meerdere markten heen — geen fout.
        logger.info("Geen extreme afwijkingen gevonden in de verwerkte markten — rustige dag")
        sys.exit(0)

    ai_report = analyze_with_ai(resultaten)

    output = os.getenv("OUTPUT_DIR", "output")
    os.makedirs(output, exist_ok=True)
    with open(f"{output}/contrarian_results.json", "w", encoding="utf-8") as f:
        json.dump(resultaten, f, indent=2, ensure_ascii=False)
    with open(f"{output}/contrarian_ai_report.txt", "w", encoding="utf-8") as f:
        f.write(ai_report)

    email_verzonden = send_email(resultaten, ai_report)
    if not email_verzonden:
        logger.critical("Analyse geslaagd, maar e-mail kon niet worden verzonden")
        sys.exit(1)

    logger.info("SCRIPT VOLLEDIG AFGEROND")


if __name__ == "__main__":
    main()
