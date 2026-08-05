#!/usr/bin/env python3
"""
cboe_option_chain.py

Telecharge la chaine d'options COMPLETE (toutes echeances : hebdo, mensuelles,
trimestrielles, LEAPS) pour UNE OU PLUSIEURS valeurs sous-jacentes (UL) depuis
le flux "delayed quotes" public du CBOE, et exporte UN CSV PAR UL (toutes les
options de cet UL dans un seul fichier).

API JSON CBOE :
    - Actions/ETF   : https://cdn.cboe.com/api/global/delayed_quotes/options/<TICKER>.json
    - Indices       : https://cdn.cboe.com/api/global/delayed_quotes/options/_<TICKER>.json
      (ex: _SPX, _VIX, _RUT, _NDX, _DJX ...)
Le script essaie automatiquement les deux formes, pas besoin de connaitre le prefixe.

Donnees fournies par contrat : bid/ask (+ tailles), dernier prix traite,
volume, open interest, IV, delta/gamma/vega/theta/rho, theorique, etc.
Delai CBOE standard : 15 minutes (flux "delayed quotes").

Installation :
    pip install requests pandas xlsxwriter

Usage :
    python cboe_option_chain.py SPX                     # un seul UL
    python cboe_option_chain.py SPX AAPL VIX TSLA        # plusieurs UL -> 1 CSV chacun
    python cboe_option_chain.py --symbols-file tickers.txt   # liste dans un fichier (1 par ligne)
    python cboe_option_chain.py SPX AAPL --xlsx          # + fichier Excel (1 onglet/echeance) par UL
    python cboe_option_chain.py SPX --raw-json           # + JSON brut telecharge
    python cboe_option_chain.py SPX -o ./data            # dossier de sortie
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import requests
    import pandas as pd
except ImportError as e:
    sys.exit(
        f"Module manquant ({e}). Installez les dependances avec :\n"
        f"    pip install requests pandas xlsxwriter"
    )

BASE_URL = "https://cdn.cboe.com/api/global/delayed_quotes/options/{sym}.json"

# Format d'un symbole d'option CBOE : <ROOT><YYMMDD><C|P><STRIKE x1000, 8 chiffres>
# ex: SPX260821C00200000  -> root=SPX,  echeance=2026-08-21, Call, strike=200.00
# ex: AAPL260805C00205000 -> root=AAPL, echeance=2026-08-05, Call, strike=205.00
OPT_RE = re.compile(r"^(?P<root>[A-Z0-9]+?)(?P<exp>\d{6})(?P<cp>[CP])(?P<strike>\d{8})$")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json",
}


def fetch_raw_for_symbol(symbol: str, timeout: int = 20) -> dict:
    """Essaie d'abord <TICKER>.json (actions/ETF), puis _<TICKER>.json (indices)."""
    symbol = symbol.strip().upper().lstrip("_")
    candidates = [symbol, f"_{symbol}"]

    last_err = None
    for cand in candidates:
        url = BASE_URL.format(sym=cand)
        try:
            resp = requests.get(url, headers=HEADERS, timeout=timeout)
            if resp.status_code == 404:
                last_err = f"404 sur {url}"
                continue
            resp.raise_for_status()
            data = resp.json()
            if data.get("data", {}).get("options"):
                return data
            last_err = f"reponse vide sur {url}"
        except requests.RequestException as e:
            last_err = str(e)

    raise RuntimeError(f"Impossible de recuperer la chaine pour '{symbol}' ({last_err})")


def parse_symbol(sym: str):
    m = OPT_RE.match(sym)
    if not m:
        return None
    root = m.group("root")
    expiry = datetime.strptime(m.group("exp"), "%y%m%d").date()
    cp = "Call" if m.group("cp") == "C" else "Put"
    strike = int(m.group("strike")) / 1000.0
    return root, expiry, cp, strike


def build_dataframe(raw: dict) -> "pd.DataFrame":
    options = raw.get("data", {}).get("options", [])
    if not options:
        raise ValueError("Aucune option trouvee dans la reponse CBOE (format inattendu ?).")

    rows = []
    for o in options:
        parsed = parse_symbol(o.get("option", ""))
        if parsed is None:
            continue
        root, expiry, cp, strike = parsed
        rows.append({"root": root, "expiration": expiry, "type": cp, "strike": strike, **o})

    df = pd.DataFrame(rows)
    if df.empty:
        raise ValueError("Impossible de parser les symboles d'options recus.")

    df["dte"] = (pd.to_datetime(df["expiration"]) - pd.Timestamp.today().normalize()).dt.days
    df["mid"] = (df["bid"] + df["ask"]) / 2

    preferred = [
        "root", "expiration", "dte", "type", "strike", "option",
        "bid", "ask", "mid", "last_trade_price", "change", "percent_change",
        "volume", "open_interest", "iv", "delta", "gamma", "vega", "theta", "rho",
        "theo", "open", "high", "low", "prev_day_close", "tick", "last_trade_time",
    ]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    df = df[cols].sort_values(["expiration", "strike", "type"]).reset_index(drop=True)
    return df


def load_symbols(args) -> list[str]:
    symbols = list(args.symbols)
    if args.symbols_file:
        path = Path(args.symbols_file)
        if not path.exists():
            sys.exit(f"Fichier introuvable : {path}")
        for line in path.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                symbols.append(line)

    # dedoublonnage en conservant l'ordre
    seen = set()
    ordered = []
    for s in symbols:
        s_up = s.strip().upper()
        if s_up and s_up not in seen:
            seen.add(s_up)
            ordered.append(s_up)

    if not ordered:
        sys.exit("Aucun UL fourni. Exemple : python cboe_option_chain.py SPX AAPL VIX")
    return ordered


def process_symbol(symbol: str, outdir: Path, stamp: str, want_xlsx: bool, want_raw: bool) -> dict:
    print(f"\n--- {symbol} ---")
    print(f"Telechargement de la chaine d'options...")
    raw = fetch_raw_for_symbol(symbol)

    if want_raw:
        raw_path = outdir / f"{symbol}_raw_{stamp}.json"
        raw_path.write_text(json.dumps(raw, indent=2))
        print(f"JSON brut sauvegarde -> {raw_path}")

    df = build_dataframe(raw)

    n_exp = df["expiration"].nunique()
    n_contracts = len(df)
    print(f"{n_contracts} contrats recuperes sur {n_exp} echeances.")

    csv_path = outdir / f"{symbol}_chain_{stamp}.csv"
    df.to_csv(csv_path, index=False)
    print(f"CSV sauvegarde -> {csv_path}")

    if want_xlsx:
        xlsx_path = outdir / f"{symbol}_chain_{stamp}.xlsx"
        try:
            with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="ALL", index=False)
                for exp, sub in df.groupby("expiration"):
                    sub.to_excel(writer, sheet_name=exp.strftime("%Y-%m-%d"), index=False)
            print(f"Excel sauvegarde -> {xlsx_path}")
        except ImportError:
            print("xlsxwriter non installe (pip install xlsxwriter) -> export Excel ignore.")

    return {"symbol": symbol, "contracts": n_contracts, "expirations": n_exp, "csv": str(csv_path)}


def main():
    ap = argparse.ArgumentParser(description="Telecharge la chaine d'options complete depuis le CBOE pour un ou plusieurs UL.")
    ap.add_argument("symbols", nargs="*", help="Tickers des sous-jacents (ex: SPX AAPL VIX TSLA)")
    ap.add_argument("--symbols-file", help="Fichier texte avec un ticker par ligne")
    ap.add_argument("-o", "--outdir", default=".", help="Dossier de sortie (defaut: repertoire courant)")
    ap.add_argument("--xlsx", action="store_true", help="Exporter aussi en Excel (un onglet par echeance) pour chaque UL")
    ap.add_argument("--raw-json", action="store_true", help="Sauvegarder aussi le JSON brut telecharge pour chaque UL")
    # parse_known_args : sous Jupyter/IPython, l'interpreteur ajoute son propre
    # argument (ex. -f kernel-....json) qu'on doit ignorer.
    args, _unknown = ap.parse_known_args()

    symbols = load_symbols(args)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    print(f"UL a traiter ({len(symbols)}) : {', '.join(symbols)}")

    results, errors = [], []
    for symbol in symbols:
        try:
            results.append(process_symbol(symbol, outdir, stamp, args.xlsx, args.raw_json))
        except Exception as e:
            print(f"ECHEC pour {symbol} : {e}")
            errors.append((symbol, str(e)))

    print("\n=== Resume ===")
    for r in results:
        print(f"  {r['symbol']:<8} {r['contracts']:>6} contrats / {r['expirations']:>3} echeances -> {r['csv']}")
    if errors:
        print("  Echecs :")
        for sym, err in errors:
            print(f"    {sym}: {err}")


if __name__ == "__main__":
    main()
