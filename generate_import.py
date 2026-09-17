#!/usr/bin/env python3
"""
V4.4 Portfolio Import Generator
Erstellt portfolio_import.json aus dem VEM AW.pdf (Kurswerte Stand 15.04.2026)
Alle Werte wurden aus dem Scan-PDF manuell/visuell ausgelesen.
"""

import json

# ── Alle Fonds aus dem PDF mit WKN → Kurswert (Einmalbetrag) ─────────────────
# Seite 2: Kasse / Defensiv
# Seite 3: Märkte
# Seite 4: Ausgewogen / Dynamisch
# Seite 5: Spezialitäten
pdf_funds = [
    # Seite 2 ─ Kasse / Defensiv / Märkte-Beginn
    {"wkn": "A0MUWS",  "name": "ZinsPlus",                          "kurswert": 25037.42},
    {"wkn": "847809",  "name": "Basis-Fonds I Nachhaltig",           "kurswert": 50050.50},
    {"wkn": "979952",  "name": "Renten Strategie K",                 "kurswert": 25055.31},
    {"wkn": "A3D8E7",  "name": "iShares iBonds Dec 2028 Term E.",    "kurswert": 24805.24},
    {"wkn": "A0X758",  "name": "ACATIS IFK Value Renten",            "kurswert": 15063.73},
    {"wkn": "A143JK",  "name": "Vanguard EUR Corporate Bond",        "kurswert": 9973.54},
    {"wkn": "A3DK4P",  "name": "Carmignac Credit 2027",              "kurswert": 10023.15},
    {"wkn": "A3EXGB",  "name": "Carmignac Credit 2029",              "kurswert": 10050.68},
    {"wkn": "A417MG",  "name": "Carmignac Portfolio Flexible Bond",  "kurswert": 21841.10},
    {"wkn": "A0RPWH",  "name": "iShares Core MSCI World UCITS ETF", "kurswert": 10258.88},
    {"wkn": "A2PSGE",  "name": "smarTrack dynamic B",                "kurswert": 15449.00},
    {"wkn": "A0M5RD",  "name": "Aktienstrategie MultiManager",       "kurswert": 15398.61},

    # Seite 3 ─ Märkte (Fortsetzung)
    {"wkn": "A2PKNU",  "name": "Eleva European Selection Fund",      "kurswert": 10441.75},
    {"wkn": "986838",  "name": "AB SICAV - American Growth Port.",   "kurswert": 10465.18},
    {"wkn": "A0YEDG",  "name": "iShares Core S&P 500 UCITS ETF",    "kurswert": 10236.53},
    {"wkn": "A3DXEB",  "name": "AXA IM Nasdaq 100 UCITS ETF",        "kurswert": 5160.87},
    {"wkn": "DBX0NJ",  "name": "Xtrackers Nikkei 225 UCITS ETF",    "kurswert": 15623.99},
    {"wkn": "A1C1H5",  "name": "iShares MSCI EM Asia UCITS ETF",    "kurswert": 15477.86},
    {"wkn": "A2PGYQ",  "name": "Wellington Enduring Infrastructure Assets", "kurswert": 9870.40},
    {"wkn": "986932",  "name": "BGF World Mining Fund",              "kurswert": 5374.72},
    {"wkn": "974119",  "name": "BGF World Gold Fund",                "kurswert": 10749.75},
    {"wkn": "A2NGLC",  "name": "Xtrackers Art. Intel. & Big Data ETF", "kurswert": 5180.62},
    {"wkn": "LB6B0M",  "name": "LBBW Sicher Leben",                  "kurswert": 15938.94},
    {"wkn": "214466",  "name": "Sauren Global Defensiv",             "kurswert": 5040.11},
    {"wkn": "A0YFQ9",  "name": "BKC Treuhand Portfolio",             "kurswert": 15107.20},

    # Seite 4 ─ Defensiv / Ausgewogen / Dynamisch
    {"wkn": "A1JUU9",  "name": "EB - Multi Asset Conservative",      "kurswert": 15170.36},
    {"wkn": "930920",  "name": "Sauren Global Balanced",             "kurswert": 15243.26},
    {"wkn": "975745",  "name": "MEAG EuroBalance",                   "kurswert": 15404.06},
    {"wkn": "A2DR2M",  "name": "ACATIS Value Event Fonds",           "kurswert": 10222.58},
    {"wkn": "A407LK",  "name": "GANÉ Value Event Fund M",            "kurswert": 10018.52},
    {"wkn": "A41MQN",  "name": "Swisscanto Portfolio Fund Sust. Balanced", "kurswert": 10273.49},
    {"wkn": "847811",  "name": "FMM-Fonds",                          "kurswert": 14944.88},
    {"wkn": "930921",  "name": "Sauren Global Opportunities",        "kurswert": 25749.53},
    {"wkn": "986855",  "name": "BL Global 75",                       "kurswert": 25352.16},
    {"wkn": "A0M430",  "name": "FvS - Multiple Opportunities R",     "kurswert": 20145.75},
    {"wkn": "A12GMG",  "name": "FERI Core Strategy Dynamic F",       "kurswert": 17770.63},
    {"wkn": "A2PT6U",  "name": "GlobalPortfolioOne",                 "kurswert": 25168.82},

    # Seite 5 ─ Spezialitäten
    {"wkn": "A12E0R",  "name": "US EquityFlex I",                    "kurswert": 20897.22},
    {"wkn": "A1H72F",  "name": "ACATIS Datini Valueflex Fonds A",    "kurswert": 15467.71},
]

# ── WKN → Block-Mapping aus fund_data.js ─────────────────────────────────────
# Manuelle Zuordnung basierend auf fund_data.js
WKN_TO_BLOCK = {
    # block-kasse (Kapitalreservefonds)
    "A0MUWS": "block-kasse",
    "847809": "block-kasse",   # auch in block-tagesgeld – kasse hat Vorrang
    "979952": "block-kasse",
    "A3D8E7": "block-kasse",
    "A3DK4P": "block-kasse",
    "A3EXGB": "block-kasse",
    "A417MG": "block-kasse",

    # block-defensiv (Defensive Vermögensverwalter)
    "A0X758": "block-defensiv",
    "A143JK": "block-defensiv",
    "214466": "block-defensiv",
    "A0YFQ9": "block-defensiv",
    "A1JUU9": "block-defensiv",

    # block-ausgewogen (Ausgewogene Vermögensverwalter)
    "930920": "block-ausgewogen",
    "975745": "block-ausgewogen",
    "A2DR2M": "block-dynamisch",   # ACATIS Value Event Fonds → Dynamisch!
    "A407LK": "block-dynamisch",   # GANÉ Value Event Fund M → Dynamisch!
    "A41MQN": "block-ausgewogen",

    # block-dynamisch (Dynamische Vermögensverwalter)
    "847811": "block-dynamisch",
    "930921": "block-maerkte",     # Sauren Global Opportunities → Märkte
    "986855": "block-dynamisch",
    "A0M430": "block-dynamisch",
    "A12GMG": "block-dynamisch",
    "A2PT6U": "block-dynamisch",

    # block-maerkte (Märkte)
    "A0RPWH": "block-maerkte",
    "A2PSGE": "block-dynamisch",   # smarTrack dynamic B → Dynamisch
    "A0M5RD": "block-maerkte",
    "A2PKNU": "block-maerkte",
    "986838": "block-maerkte",
    "A0YEDG": "block-maerkte",
    "A3DXEB": "block-maerkte",
    "DBX0NJ": "block-maerkte",
    "A1C1H5": "block-maerkte",
    "986932": "block-maerkte",
    "974119": "block-maerkte",
    "A2NGLC": "block-maerkte",
    "LB6B0M": "block-maerkte",

    # block-spezial (Spezialitäten / Themen)
    "A2PGYQ": "block-spezial",
    "A12E0R": "block-spezial",
    "A1H72F": "block-dynamisch",   # ACATIS Datini Valueflex → Dynamisch
}

# Fund name mapping for exact match against fund_data.js
WKN_TO_FUNDNAME = {
    "A0MUWS": "ZinsPlus",
    "847809": "Basis-Fonds I Nachhaltig",
    "979952": "Renten Strategie K",
    "A3D8E7": "iShares iBonds Dec 2028 Term E.",
    "A0X758": "ACATIS IFK Value Renten",
    "A143JK": "Vanguard EUR Corporate Bond",
    "A3DK4P": "Carmignac Credit 2027",
    "A3EXGB": "Carmignac Credit 2029",
    "A417MG": "Carmignac Portfolio Flexible Bond",
    "A0RPWH": "iShares Core MSCI World UCITS ETF",
    "A2PSGE": "smarTrack dynamic B",
    "A0M5RD": "Aktienstrategie MultiManager",
    "A2PKNU": "Eleva European Selection Fund",
    "986838": "AB SICAV - American Growth Port.",
    "A0YEDG": "iShares Core S&P 500 UCITS ETF",
    "A3DXEB": "AXA IM Nasdaq 100 UCITS ETF",
    "DBX0NJ": "Xtrackers Nikkei 225 UCITS ETF",
    "A1C1H5": "iShares MSCI EM Asia UCITS ETF",
    "A2PGYQ": "Wellington Enduring Infrastructure Assets",
    "986932": "BGF World Mining Fund",
    "974119": "BGF World Gold Fund",
    "A2NGLC": "Xtrackers Art. Intel. & Big Data ETF",
    "LB6B0M": "LBBW Sicher Leben",
    "214466": "Sauren Global Defensiv",
    "A0YFQ9": "BKC Treuhand Portfolio",
    "A1JUU9": "EB - Multi Asset Conservative",
    "930920": "Sauren Global Balanced",
    "975745": "MEAG EuroBalance",
    "A2DR2M": "ACATIS Value Event Fonds",
    "A407LK": "GANÉ Value Event Fund M",
    "A41MQN": "Swisscanto Portfolio Fund Sust. Balanced",
    "847811": "FMM-Fonds",
    "930921": "Sauren Global Opportunities",
    "986855": "BL Global 75",
    "A0M430": "FvS - Multiple Opportunities R",
    "A12GMG": "FERI Core Strategy Dynamic F",
    "A2PT6U": "GlobalPortfolioOne",
    "A12E0R": "US EquityFlex I",
    "A1H72F": "ACATIS Datini Valueflex Fonds A",
}

# Block info mapping
BLOCK_INFO = {
    "block-kasse":     {"title": "Kapitalreservefonds"},
    "block-defensiv":  {"title": "Defensive Vermögensverwalter"},
    "block-ausgewogen":{"title": "Ausgewogene Vermögensverwalter"},
    "block-dynamisch": {"title": "Dynamische Vermögensverwalter"},
    "block-maerkte":   {"title": "Märkte"},
    "block-spezial":   {"title": "Spezialitäten / Themen"},
}

# ── Import-JSON aufbauen ──────────────────────────────────────────────────────
fund_investments = {}
fund_sparrates   = {}
portfolio_layers = {}
unmatched        = []

total_einmal = 0.0

for fund in pdf_funds:
    wkn = fund["wkn"]
    block_id = WKN_TO_BLOCK.get(wkn)
    fund_name = WKN_TO_FUNDNAME.get(wkn)
    
    if block_id and fund_name:
        key = f"{block_id}::{fund_name}"
        fund_investments[key] = round(fund["kurswert"], 2)
        total_einmal += fund["kurswert"]
        
        # Layer aufbauen
        if block_id not in portfolio_layers:
            portfolio_layers[block_id] = {
                "blockId": block_id,
                "blockTitle": BLOCK_INFO[block_id]["title"],
                "allocation": 0,
                "funds": []
            }
        
        # Fund zur Layer hinzufügen (falls nicht schon drin)
        existing = [f for f in portfolio_layers[block_id]["funds"] if f["name"] == fund_name]
        if not existing:
            portfolio_layers[block_id]["funds"].append({
                "name": fund_name,
                "info": f"WKN: {wkn}",
                "type": "",
                "ertrag": ""
            })
        
        print(f"✅ {wkn:10s} → {block_id:20s} :: {fund_name} = {fund['kurswert']:,.2f} €")
    else:
        unmatched.append(fund)
        print(f"❌ {wkn:10s} → NICHT GEFUNDEN: {fund['name']} ({fund['kurswert']:,.2f} €)")

# Reihenfolge der Layer wie im Turm
LAYER_ORDER = ["block-kasse", "block-defensiv", "block-ausgewogen", "block-dynamisch", "block-maerkte", "block-spezial"]
portfolio_list = [portfolio_layers[lid] for lid in LAYER_ORDER if lid in portfolio_layers]

# Investment-Objekt aufbauen
portfolio_globals = {
    "totalInvestment": round(total_einmal, 2),
    "totalSparrate": 0,
    "fundInvestments": fund_investments,
    "fundSparrates": fund_sparrates
}

import_data = {
    "portfolio": portfolio_list,
    "globals": portfolio_globals
}

# Ausgabe als JSON
out_path = "portfolio_import.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(import_data, f, ensure_ascii=False, indent=2)

print(f"\n{'='*60}")
print(f"✅ Import-JSON erstellt: {out_path}")
print(f"   Gesamtinvestment: {total_einmal:,.2f} €")
print(f"   Zugeordnete Fonds: {len(fund_investments)}")
print(f"   Nicht gefunden: {len(unmatched)}")
if unmatched:
    print(f"\n⚠️  Nicht zugeordnet:")
    for u in unmatched:
        print(f"   {u['wkn']:10s} {u['name']} ({u['kurswert']:,.2f} €)")
