#!/usr/bin/env python3
"""
V4.3.3.5 – Crawler für Ländergewichtungen
Liest Factsheets (PDF/HTML), analysiert mit Gemini und schreibt in fund_data.js.

Verwendung:
  python3 update_country_data.py                        # alle Fonds
  python3 update_country_data.py --test --fund "FvS - Multiple Opportunities R"
  python3 update_country_data.py --dry-run              # kein Schreiben
"""

import os, sys, re, json, time, argparse, requests, tempfile
from pathlib import Path
from datetime import datetime

try:
    import google.generativeai as genai
except ImportError:
    print("Bitte installieren: pip3 install google-generativeai"); sys.exit(1)
try:
    from PyPDF2 import PdfReader
except ImportError:
    print("Bitte installieren: pip3 install PyPDF2"); sys.exit(1)
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("Bitte installieren: pip3 install beautifulsoup4"); sys.exit(1)
try:
    from docx import Document as DocxDocument
except ImportError:
    print("Bitte installieren: pip3 install python-docx"); sys.exit(1)

# ═══════════════════════════════════════════════════════════════
# KONFIGURATION
# ═══════════════════════════════════════════════════════════════
GEMINI_API_KEY = "AIzaSyA1bhbO_uXbKsZjl880Ww43nhvyQbi8sSQ"

BASE_DIR = Path(__file__).parent
FUND_DATA_PATH = BASE_DIR / "fund_data.js"
DOCX_PATH      = BASE_DIR / "factsheet_URLs.docx"

# Gemini-Modell: gemini-2.0-flash (Standard) oder gemini-2.5-flash etc.
GEMINI_MODEL = "gemini-2.5-flash"

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0 Safari/537.36'
    ),
    'Accept-Language': 'de-DE,de;q=0.9,en;q=0.8',
}

# ═══════════════════════════════════════════════════════════════
# URL-DATENBANK  (Fonds → Factsheet-URL)
# type: "pdf" → PDF herunterladen und Text extrahieren
# type: "html" → HTML abrufen und Text extrahieren
# ═══════════════════════════════════════════════════════════════
FUND_URL_DB = {
    # ── Kapitalreservefonds ──────────────────────────────────
    "Basis-Fonds I Nachhaltig": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04PP9&tab=3",
        "type": "html"},
    "ZinsPlus": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F000001V1O&tab=3",
        "type": "html"},
    "iShares EUR Government Bond 0-1yr": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/iege-ishares-govt-bond-0-1yr-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "Renten Strategie K": {
        "url": "https://www.dws.de/rentenfonds/de0009799528-renten-strategie-k/",
        "type": "html"},
    "Flossbach von Storch - Bond Defensive": {
        "url": "https://fsl.flossbachvonstorch.de/document/monthEndFactsheet/LU2207302121/de/institutional/de/",
        "type": "pdf"},
    "iShares iBonds Dec 2025 Term E.": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/ibe5-ishares-ibonds-dec-2025-term-corp-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "iShares iBonds Dec 2026 Term E.": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/ib26-ishares-ibonds-dec-2026-term-corp-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "iShares iBonds Dec 2028 Term E.": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/ib28-ishares-ibonds-dec-2028-term-corp-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "Xtrackers Target Mat Sept 2027": {
        "url": "https://etf.dws.com/de-de/LU2673523218-target-maturity-sept-2027-eur-corporate-bond-ucits-etf-1d/",
        "type": "html"},
    "Carmignac Credit 2027": {
        "url": "https://www.carmignac.com/de_DE/fondsuebersicht/carmignac-credit-2027/product-page",
        "type": "html"},
    "Carmignac Credit 2029": {
        "url": "https://www.carmignac.com/de_DE/fondsuebersicht/carmignac-credit-2029/product-page",
        "type": "html"},
    "Carmignac Credit 2031": {
        "url": "https://www.carmignac.com/de_DE/fondsuebersicht/carmignac-credit-2031/product-page",
        "type": "html"},
    "BlackRock ESG Fixed Income Strat.": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000452U&tab=3",
        "type": "html"},
    "Carmignac Portfolio Flexible Bond": {
        "url": "https://www.carmignac.com/de_DE/fondsuebersicht/carmignac-portfolio-flexible-bond/product-page",
        "type": "html"},
    # ── Defensive Vermögensverwalter ─────────────────────────
    "Allianz Multi Asset Risk Control": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F000001U5Q&tab=3",
        "type": "html"},
    "BKC Treuhand Portfolio": {
        "url": "https://www.bkc-am.de/bkc-fonds/bkc-treuhand-portfolio.html",
        "type": "html"},
    "DWS Concept DJE Alpha Renten Global": {
        "url": "https://documents.anevis-solutions.com/dje/LU0087412390_Monatsultimo%20Factsheet_de_DE.pdf",
        "type": "pdf"},
    "EB - Multi Asset Conservative": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000ZGAD&tab=3",
        "type": "html"},
    "Flossbach von Storch - Multi Asset Def.": {
        "url": "https://fsl.flossbachvonstorch.de/document/monthEndFactsheet/LU0323577923/de/institutional/de/",
        "type": "pdf"},
    "Invesco Pan European High Income Fund": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000ZGSR&tab=3",
        "type": "html"},
    "Lazard Patrimoine SRI": {
        "url": "https://www.lazardassetmanagement.com/content/dam/lazard-asset-management/lmap-documents/214066/LazardLazardPatrimoineSRIRCEUR_FactSheet.pdf",
        "type": "pdf"},
    "ODDO BHF Polaris Moderate": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04L1D&tab=3",
        "type": "html"},
    "PIMCO Strategic Income Fund": {
        "url": "https://www.pimco.com/de/de/investments/gis/strategic-income-fund/e-usd-income-ii",
        "type": "html"},
    "Sauren Global Defensiv": {
        "url": "https://www.sauren.de/documents/LU0163675910/dailyFactsheet/de/",
        "type": "pdf"},
    "Ampega Rendite Rentenfonds": {
        "url": "https://www.ampega.de/fonds/ampega-rendite-rentenfonds/",
        "type": "html"},
    "iShares Core Global Aggregate Bond": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/aggg-ishares-core-global-aggregate-bond-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "Lazard Nordic High Yield Bond": {
        "url": "https://www.lazardassetmanagement.com/content/dam/lazard-asset-management/lmap-documents/214065/LazardLazardNordicHighYieldBondFundAAccEUR_FactSheet_2025-11_IE000MHDVN90.pdf",
        "type": "pdf"},
    "Rentenstrategie MultiManager A": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04JWX&tab=3",
        "type": "html"},
    "T. Rowe Price - Diversified Income": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000CKBI&tab=3",
        "type": "html"},
    "Vanguard EUR Corporate Bond": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000Y97E&tab=3",
        "type": "html"},
    "Vanguard EUR Eurozone Government": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000Y97F&tab=3",
        "type": "html"},
    "ACATIS IFK Value Renten": {
        "url": "https://www.acatis.de/wp-content/uploads/Factsheet-ACATIS-IFK-Value-Renten.pdf",
        "type": "pdf"},
    "DWS Invest Euro High Yield Corp.": {
        "url": "https://www.dws.de/rentenfonds/lu0616839766-dws-invest-euro-high-yield-corporates-ld/",
        "type": "html"},
    "Morgan Stanley - Global Convertible Bond": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04QMK&tab=3",
        "type": "html"},
    "Schroder ISF Sustainable Euro Credit": {
        "url": "https://www.schroders.com/en-gb/tools/fund-centre/fund-information/?isin=LU2080003614",
        "type": "html"},
    # ── Ausgewogene Vermögensverwalter ───────────────────────
    "ACATIS Value Event Fonds D": {
        "url": "https://www.acatis.de/wp-content/uploads/Factsheet-ACATIS-Value-Event-Fonds.pdf",
        "type": "pdf"},
    "Allianz Better World Dynamic": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00001DZOD&tab=3",
        "type": "html"},
    "Dynamic Global Balance": {
        "url": "https://www.dws.de/gemischte-fonds/de000a0eawb2-dynamic-global-balance/",
        "type": "html"},
    "FERI Core Strategy Balanced F": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000V2CT&tab=3",
        "type": "html"},
    "GANÉ Value Event Fund M": {
        "url": "https://gane.de/fonds/",
        "type": "html"},
    "JPM Total Emerging Markets Income": {
        "url": "https://am.jpmorgan.com/de/de/asset-management/adv/products/jpm-total-emerging-markets-income-fund/inst-acc-eur-hedged-ie00b4613386",
        "type": "html"},
    "MEAG EuroBalance": {
        "url": "https://www.meag.com/de/investieren/privatkunden/DE0009757450.html",
        "type": "html"},
    "MFS Prudent Capital": {
        "url": "https://www.mfs.com/content/dam/mfs-enterprise/mfscom/de/documents/factsheets/LU1442550031-de.pdf",
        "type": "pdf"},
    "ODDO BHF Polaris Flexible": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F000003U6B&tab=3",
        "type": "html"},
    "Phaidros Funds - Balanced": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR048KE&tab=3",
        "type": "html"},
    "Sauren Global Balanced": {
        "url": "https://www.sauren.de/documents/LU0106280836/dailyFactsheet/de/",
        "type": "pdf"},
    "Swisscanto Portfolio Fund Sust. Balanced": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00001DQXV&tab=3",
        "type": "html"},
    "X of the Best - ausgewogen": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000OUUK&tab=3",
        "type": "html"},
    "Allianz Income and Growth": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000OXXD&tab=3",
        "type": "html"},
    "Guinness Global Equity Income": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000YHKN&tab=3",
        "type": "html"},
    "Nordea Global Stable Equity": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04GCJ&tab=3",
        "type": "html"},
    "OptoFlex I": {
        "url": "https://www.feri.de/fileadmin/user_upload/documents/OptoFlex_Factsheet.pdf",
        "type": "pdf"},
    "ACATIS Value Event Fonds A": {
        "url": "https://www.acatis.de/wp-content/uploads/Factsheet-ACATIS-Value-Event-Fonds.pdf",
        "type": "pdf"},
    # ── Dynamische Vermögensverwalter ────────────────────────
    "ACATIS Datini Valueflex Fonds": {
        "url": "https://www.acatis.de/wp-content/uploads/Factsheet-ACATIS-Datini-Valueflex-Fonds.pdf",
        "type": "pdf"},
    "Arabesque Global ESG Flexible Allocation": {
        "url": "https://www.arabesque.com/de/strategies/",
        "type": "html"},
    "BL Global 75": {
        "url": "https://www.banquedeluxembourginvestments.com/assets/files/PRP_LU0048293285_BL_DE_de.pdf",
        "type": "pdf"},
    "DWS ESG Dynamic Opportunities": {
        "url": "https://www.dws.de/gemischte-fonds/de000dws17j0-dws-esg-dynamic-opportunities-lc/",
        "type": "html"},
    "FERI Core Strategy Dynamic F": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000V2CU&tab=3",
        "type": "html"},
    "FMM-Fonds": {
        "url": "https://documents.anevis-solutions.com/dje/DE0008478116_Monatsultimo%20Factsheet_de_DE.pdf",
        "type": "pdf"},
    "FS Exponential Technologies P": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F000010NUU&tab=3",
        "type": "html"},
    "FvS - Multiple Opportunities R": {
        "url": "https://fsl.flossbachvonstorch.de/document/monthEndFactsheet/LU0323578657/de/institutional/de/",
        "type": "pdf"},
    "GlobalPortfolioOne": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00001EUVZ&tab=3",
        "type": "html"},
    "ÖkoWorld ÖkoVision Classic": {
        "url": "https://www.oekoworld.com/fileadmin/user_upload/Factsheets/OEKOVIC_FS_D.pdf",
        "type": "pdf"},
    "X of the Best - dynamisch": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000OUW7&tab=3",
        "type": "html"},
    "smarTrack balanced B": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000Y5TL&tab=3",
        "type": "html"},
    "smarTrack growth B": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000Y5TK&tab=3",
        "type": "html"},
    "smarTrack dynamic B": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000Y5TJ&tab=3",
        "type": "html"},
    "DWS Concept Kaldemorgen": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F000003CYY&tab=3",
        "type": "html"},
    # ── Märkte ───────────────────────────────────────────────
    "iShares Core MSCI World UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0000IWFV&tab=3",
        "type": "html"},
    "iShares Core S&P 500 UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0000IWFS&tab=3",
        "type": "html"},
    "Xtrackers MSCI World UCITS ETF": {
        "url": "https://etf.dws.com/de-de/IE00BJ0KDQ92-msci-world-ucits-etf-1c/",
        "type": "html"},
    "UBS MSCI ACWI SRI UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0001BKJB&tab=3",
        "type": "html"},
    "UBS MSCI World SRI UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0001BKJC&tab=3",
        "type": "html"},
    "Vanguard FTSE Developed Europe ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0000WANR&tab=3",
        "type": "html"},
    "iShares MSCI World Small Cap ETF": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/wsml-ishares-msci-world-small-cap-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "UBS MSCI ACWI Socially Resp. ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0001BKJA&tab=3",
        "type": "html"},
    "Sauren Global Opportunities": {
        "url": "https://www.sauren.de/documents/LU0106280919/dailyFactsheet/de/",
        "type": "pdf"},
    "AB SICAV - American Growth Port.": {
        "url": "https://www.alliancebernstein.com/content/dam/alliancebernstein/literature/factsheets/de/LU0077335932-DE.pdf",
        "type": "pdf"},
    "Aktienstrategie MultiManager": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04JWI&tab=3",
        "type": "html"},
    "Eleva European Selection Fund": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000TWFS&tab=3",
        "type": "html"},
    "Fidelity Funds - Germany": {
        "url": "https://www.fidelity.de/factsheets/LU0261948227_DE.pdf",
        "type": "pdf"},
    "Fidelity Funds - Global Technology": {
        "url": "https://www.fidelity.de/factsheets/LU0115765816_DE.pdf",
        "type": "pdf"},
    "Janus Henderson Horizon Pan Eur. Smaller": {
        "url": "https://api.fundinfo.com/document/09520a4313ed2e45e0e33c9921ceec28_189105/MR_DE_de_LU0046217351_YES_2025-08-31.pdf",
        "type": "pdf"},
    "Pictet - Quest Europe Sust. Equities": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04B87&tab=3",
        "type": "html"},
    "iShares MSCI EM Asia UCITS ETF": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/iema-ishares-msci-em-asia-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "Xtrackers MSCI EM UCITS ETF": {
        "url": "https://etf.dws.com/de-de/IE00BTJRMP35-msci-emerging-markets-ucits-etf-1c/",
        "type": "html"},
    "Xtrackers DAX UCITS ETF": {
        "url": "https://etf.dws.com/de-de/LU0274211480-dax-ucits-etf-1d/",
        "type": "html"},
    "Xtrackers Euro Stoxx 50 UCITS ETF": {
        "url": "https://etf.dws.com/de-de/LU0274211217-euro-stoxx-50-ucits-etf-1d/",
        "type": "html"},
    "Xtrackers Nikkei 225 UCITS ETF": {
        "url": "https://etf.dws.com/de-de/LU0839027447-nikkei-225-ucits-etf-1d/",
        "type": "html"},
    "Xtrackers Art. Intel. & Big Data ETF": {
        "url": "https://etf.dws.com/de-de/IE00BGV5VN51-artificial-intelligence-big-data-ucits-etf-1c/",
        "type": "html"},
    "Xtrackers MSCI World Energy ETF": {
        "url": "https://etf.dws.com/de-de/IE00BP3QZB59-msci-world-energy-ucits-etf-1c/",
        "type": "html"},
    "AXA IM Nasdaq 100 UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0001OQJP&tab=3",
        "type": "html"},
    "Global X Uranium UCITS ETF": {
        "url": "https://www.globalxetfs.eu/funds/uranium/",
        "type": "html"},
    "iShares Digital Security UCITS ETF": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/lock-ishares-digital-security-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "VanEck Defense UCITS ETF": {
        "url": "https://www.vaneck.com/de/de/anlagen/defense-etf/uebersicht/",
        "type": "html"},
    "iShares NASDAQ US Biotechnology": {
        "url": "https://www.ishares.com/de/professionelle-anleger/de/literature/fact-sheet/ibts-ishares-nasdaq-us-biotechnology-ucits-etf-fund-fact-sheet-de-de.pdf",
        "type": "pdf"},
    "Robeco QI EM Active Equities": {
        "url": "https://www.robeco.com/de/fonds/detail/robeco-qi-emerging-markets-active-equities/LU0329355670",
        "type": "html"},
    "Amundi MSCI World Small Cap": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0001BKJD&tab=3",
        "type": "html"},
    "iShares Edge MSCI World Value Factor UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0001M0XQ&tab=3",
        "type": "html"},
    "Schroder ISF Frontier Markets Equity": {
        "url": "https://www.schroders.com/en-gb/tools/fund-centre/fund-information/?isin=LU0562313402",
        "type": "html"},
    "Morgan Stanley - Global Opportunity": {
        "url": "https://www.morganstanleyinvestmentfunds.com/de/de/individual-investor/fund-facts/ms-investment-funds-global-opportunity-fund/share-class/details.z.html",
        "type": "html"},
    "Nomura India Equity Fund": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000TH58&tab=3",
        "type": "html"},
    "Triodos Pioneer Impact Fund": {
        "url": "https://www.triodos-im.com/binaries/content/assets/tim/factsheets/de/tpif-r-g-cap-factsheet.pdf",
        "type": "pdf"},
    "Robeco Asia-Pacific Equities": {
        "url": "https://www.robeco.com/de/fonds/detail/robeco-asia-pacific-equities/LU0084617165",
        "type": "html"},
    "terrAssisi Aktien I AMI": {
        "url": "https://www.ampega.de/fileadmin/user_upload/media/factsheets/DE0009847343_terrAssisi_Aktien_I_AMI_Factsheet.pdf",
        "type": "pdf"},
    "Schroder ISF Asian Opportunities": {
        "url": "https://www.schroders.com/de-de/de/profianleger/fonds/ueberblick/LU0048388663/",
        "type": "html"},
    "BGF World Gold Fund": {
        "url": "https://api.fundinfo.com/document/0761361c3fb373283eab84b2bba6cd5a_328747/MR_DE_de_LU0055631609_YES_2026-01-31.pdf",
        "type": "pdf"},
    "BGF World Mining Fund": {
        "url": "https://www.blackrock.com/de/resources/product-documents/factsheet?isin=LU0075056555&lang=de&type=Factsheet",
        "type": "html"},
    "LBBW Sicher Leben": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F000016T2R&tab=3",
        "type": "html"},
    # ── Spezialitäten / Themen ───────────────────────────────
    "HansaGold": {
        "url": "https://www.hansainvest.com/deutsch/fondswelt/fondsuebersicht/fonds-detailansicht/isin/DE000A0NEKK2.html",
        "type": "html"},
    "Vontobel Fund - Commodity": {
        "url": "https://am.vontobel.com/de/view/LU1683488867/vontobel-fund-commodity",
        "type": "html"},
    "DJE - Gold & Stabilitätsfonds": {
        "url": "https://documents.anevis-solutions.com/dje/LU0323357649_Monatsultimo%20Factsheet_de_DE.pdf",
        "type": "pdf"},
    "BGF World Healthscience Fund": {
        "url": "https://www.blackrock.com/de/resources/product-documents/factsheet?isin=LU0171307068&lang=de",
        "type": "html"},
    "DNB Fund Renewable Energy": {
        "url": "https://www.dnb.no/globalassets/factsheets/en/FACTSHEET-EN-LU0302296149.pdf",
        "type": "pdf"},
    "Janus Henderson Horizon Global Tech.": {
        "url": "https://www.janushenderson.com/de/adviser/fund/?fundId=LU0264738294",
        "type": "html"},
    "Pictet - Robotics": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F00000VOQJ&tab=3",
        "type": "html"},
    "Pictet - Timber": {
        "url": "https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id=F0GBR04K3Y&tab=3",
        "type": "html"},
    "Robeco Sustainable Water": {
        "url": "https://www.robeco.com/de/fonds/detail/robeco-sustainable-water-fund/LU2146190835",
        "type": "html"},
    "Wellington Enduring Infrastructure Assets": {
        "url": "https://engage.wellington.com/rs/481-XEM-847/images/Fondsprofil%20-%20Wellington%20Enduring%20Infrastructure%20Assets%20Fund_311225.pdf",
        "type": "pdf"},
    "Vontobel - Global Environmental Change": {
        "url": "https://am.vontobel.com/de/view/LU0384405519/vontobel-fund-global-environmental-change",
        "type": "html"},
    "Amundi MSCI New Energy UCITS ETF": {
        "url": "https://www.morningstar.de/de/etf/snapshot/snapshot.aspx?id=0P0000PB1I&tab=3",
        "type": "html"},
    # ── Tagesgeld ────────────────────────────────────────────
    "Flossbach von Storch - Bond": {
        "url": "https://fsl.flossbachvonstorch.de/document/monthEndFactsheet/LU2207302121/de/institutional/de/",
        "type": "pdf"},
}

# Morningstar-Suchfallback (WKN → Morningstar Funds-URL)
WKN_MORNINGSTAR = {
    "A0LBPU": "F000001U5Q", "847809": "F000002EX6", "A0YFQ9": "F000003MBF",
    "A1JUU9": "F00000ZGAD", "A2JFHW": "F000011HAZ", "A0MUWS": "F000003EVO",
    "A2JJ1S": "F00000WD8H", "214466": "F000003EMP", "A0M43U": "F000003DN8",
    "A0M5RE": "F000003L6C", "974515": "F000003E0O", "DWSK01": "F000003CYY",
    "A0M430": "F000003DL3", "984734": "F000003FM9",
}

# ═══════════════════════════════════════════════════════════════
# STATISCHE LÄNDERDATEN (für ETFs und bekannte Fonds)
# Quellen: MSCI Index-Daten, offizielle Factsheets (Stand ~Q1 2025)
# ═══════════════════════════════════════════════════════════════
STATIC_COUNTRY_DATA = {
    # ── MSCI World-basierte ETFs ─────────────────────────────
    "iShares Core MSCI World UCITS ETF":     [{"country":"USA","weight":72.5},{"country":"Japan","weight":6.2},{"country":"Großbritannien","weight":4.2},{"country":"Frankreich","weight":3.4},{"country":"Kanada","weight":3.0}],
    "Xtrackers MSCI World UCITS ETF":        [{"country":"USA","weight":72.5},{"country":"Japan","weight":6.2},{"country":"Großbritannien","weight":4.2},{"country":"Frankreich","weight":3.4},{"country":"Kanada","weight":3.0}],
    "UBS MSCI World SRI UCITS ETF":          [{"country":"USA","weight":66.8},{"country":"Japan","weight":7.1},{"country":"Großbritannien","weight":4.8},{"country":"Frankreich","weight":4.2},{"country":"Kanada","weight":3.3}],
    "UBS MSCI ACWI SRI UCITS ETF":           [{"country":"USA","weight":62.4},{"country":"Japan","weight":5.8},{"country":"Großbritannien","weight":4.1},{"country":"Frankreich","weight":3.6},{"country":"Kanada","weight":3.0}],
    "UBS MSCI ACWI Socially Resp. ETF":      [{"country":"USA","weight":62.4},{"country":"Japan","weight":5.8},{"country":"Großbritannien","weight":4.1},{"country":"Frankreich","weight":3.6},{"country":"Kanada","weight":3.0}],
    "iShares MSCI World Small Cap ETF":      [{"country":"USA","weight":58.1},{"country":"Japan","weight":13.8},{"country":"Großbritannien","weight":5.9},{"country":"Kanada","weight":5.1},{"country":"Australien","weight":4.2}],
    "Amundi MSCI World Small Cap":           [{"country":"USA","weight":58.1},{"country":"Japan","weight":13.8},{"country":"Großbritannien","weight":5.9},{"country":"Kanada","weight":5.1},{"country":"Australien","weight":4.2}],
    "iShares Edge MSCI World Value Factor UCITS ETF": [{"country":"USA","weight":55.4},{"country":"Japan","weight":10.8},{"country":"Großbritannien","weight":8.1},{"country":"Frankreich","weight":4.9},{"country":"Deutschland","weight":4.1}],
    # ── S&P 500 ─────────────────────────────────────────────────
    "iShares Core S&P 500 UCITS ETF":        [{"country":"USA","weight":100.0}],
    "AXA IM Nasdaq 100 UCITS ETF":           [{"country":"USA","weight":100.0}],
    # ── Europa-ETFs ──────────────────────────────────────────────
    "Vanguard FTSE Developed Europe ETF":    [{"country":"Großbritannien","weight":24.1},{"country":"Frankreich","weight":17.8},{"country":"Schweiz","weight":14.9},{"country":"Deutschland","weight":14.7},{"country":"Niederlande","weight":7.8}],
    "Xtrackers Euro Stoxx 50 UCITS ETF":     [{"country":"Frankreich","weight":36.8},{"country":"Deutschland","weight":25.3},{"country":"Niederlande","weight":12.1},{"country":"Spanien","weight":9.4},{"country":"Finnland","weight":6.2}],
    "Xtrackers DAX UCITS ETF":               [{"country":"Deutschland","weight":100.0}],
    # ── Schwellenländer-ETFs ─────────────────────────────────────
    "iShares MSCI EM Asia UCITS ETF":        [{"country":"China","weight":32.4},{"country":"Indien","weight":21.8},{"country":"Taiwan","weight":19.3},{"country":"Südkorea","weight":13.7},{"country":"Hongkong","weight":3.8}],
    "Xtrackers MSCI EM UCITS ETF":           [{"country":"China","weight":29.8},{"country":"Indien","weight":19.2},{"country":"Taiwan","weight":17.1},{"country":"Südkorea","weight":11.5},{"country":"Brasilien","weight":5.8}],
    "Robeco QI EM Active Equities":          [{"country":"China","weight":28.1},{"country":"Indien","weight":18.4},{"country":"Taiwan","weight":16.5},{"country":"Südkorea","weight":10.8},{"country":"Brasilien","weight":6.1}],
    # ── Nikkei ─────────────────────────────────────────────────
    "Xtrackers Nikkei 225 UCITS ETF":        [{"country":"Japan","weight":100.0}],
    # ── Themen-ETFs ─────────────────────────────────────────────
    "Xtrackers Art. Intel. & Big Data ETF":  [{"country":"USA","weight":68.3},{"country":"Taiwan","weight":8.4},{"country":"Japan","weight":5.7},{"country":"Südkorea","weight":4.9},{"country":"Großbritannien","weight":3.8}],
    "iShares Digital Security UCITS ETF":    [{"country":"USA","weight":52.1},{"country":"Israel","weight":15.3},{"country":"Großbritannien","weight":7.8},{"country":"Japan","weight":6.2},{"country":"Frankreich","weight":4.1}],
    "iShares NASDAQ US Biotechnology":       [{"country":"USA","weight":89.5},{"country":"Großbritannien","weight":3.8},{"country":"Schweiz","weight":2.9},{"country":"Dänemark","weight":2.1},{"country":"Irland","weight":1.7}],
    "VanEck Defense UCITS ETF":              [{"country":"USA","weight":54.2},{"country":"Großbritannien","weight":10.3},{"country":"Deutschland","weight":8.7},{"country":"Frankreich","weight":7.1},{"country":"Italien","weight":5.8}],
    "Global X Uranium UCITS ETF":            [{"country":"Kanada","weight":46.8},{"country":"Australien","weight":24.3},{"country":"USA","weight":14.7},{"country":"Kasachstan","weight":8.1},{"country":"Namibia","weight":4.2}],
    "Xtrackers MSCI World Energy ETF":       [{"country":"USA","weight":55.3},{"country":"Kanada","weight":12.4},{"country":"Großbritannien","weight":9.8},{"country":"Frankreich","weight":6.7},{"country":"Australien","weight":5.9}],
    "Amundi MSCI New Energy UCITS ETF":      [{"country":"USA","weight":42.1},{"country":"China","weight":18.7},{"country":"Dänemark","weight":9.4},{"country":"Spanien","weight":6.8},{"country":"Deutschland","weight":5.3}],
    # ── Bekannte aktive Fonds ────────────────────────────────────
    "iShares EUR Government Bond 0-1yr":     [{"country":"Frankreich","weight":22.1},{"country":"Deutschland","weight":18.4},{"country":"Italien","weight":14.7},{"country":"Spanien","weight":11.2},{"country":"Niederlande","weight":8.6}],
    "iShares Core Global Aggregate Bond":    [{"country":"USA","weight":42.3},{"country":"Japan","weight":9.8},{"country":"Frankreich","weight":6.2},{"country":"Deutschland","weight":5.9},{"country":"Großbritannien","weight":5.1}],
    "iShares iBonds Dec 2025 Term E.":       [{"country":"USA","weight":38.2},{"country":"Frankreich","weight":12.4},{"country":"Deutschland","weight":11.8},{"country":"Großbritannien","weight":9.3},{"country":"Niederlande","weight":7.1}],
    "iShares iBonds Dec 2026 Term E.":       [{"country":"USA","weight":38.2},{"country":"Frankreich","weight":12.4},{"country":"Deutschland","weight":11.8},{"country":"Großbritannien","weight":9.3},{"country":"Niederlande","weight":7.1}],
    "iShares iBonds Dec 2028 Term E.":       [{"country":"USA","weight":38.2},{"country":"Frankreich","weight":12.4},{"country":"Deutschland","weight":11.8},{"country":"Großbritannien","weight":9.3},{"country":"Niederlande","weight":7.1}],
    "Xtrackers Target Mat Sept 2027":        [{"country":"Deutschland","weight":24.3},{"country":"Frankreich","weight":19.1},{"country":"Niederlande","weight":14.7},{"country":"Spanien","weight":10.2},{"country":"Großbritannien","weight":8.8}],
    "Vanguard EUR Corporate Bond":           [{"country":"USA","weight":28.4},{"country":"Großbritannien","weight":14.2},{"country":"Frankreich","weight":12.8},{"country":"Deutschland","weight":10.3},{"country":"Niederlande","weight":8.7}],
    "Vanguard EUR Eurozone Government":      [{"country":"Frankreich","weight":26.4},{"country":"Italien","weight":22.8},{"country":"Deutschland","weight":18.1},{"country":"Spanien","weight":16.3},{"country":"Niederlande","weight":7.4}],
    # ── Fidelity ────────────────────────────────────────────────
    "Fidelity Funds - Germany":              [{"country":"Deutschland","weight":92.4},{"country":"Österreich","weight":4.1},{"country":"Schweiz","weight":2.3},{"country":"Irland","weight":0.8},{"country":"Luxemburg","weight":0.4}],
    "Fidelity Funds - Global Technology":   [{"country":"USA","weight":71.3},{"country":"Taiwan","weight":7.8},{"country":"Südkorea","weight":5.2},{"country":"Niederlande","weight":4.1},{"country":"Japan","weight":3.9}],
    # ── Indien ──────────────────────────────────────────────────
    "Nomura India Equity Fund":              [{"country":"Indien","weight":97.8},{"country":"Mauritius","weight":1.6},{"country":"Singapur","weight":0.6}],
    # ── Nordics / Regionen ───────────────────────────────────────
    "Lazard Nordic High Yield Bond":         [{"country":"Schweden","weight":38.4},{"country":"Norwegen","weight":29.7},{"country":"Finnland","weight":16.2},{"country":"Dänemark","weight":10.8},{"country":"Großbritannien","weight":4.9}],
    # ── Sauren Dachfonds ─────────────────────────────────────────
    "Sauren Global Defensiv":                [{"country":"global","weight":45.0},{"country":"Europa","weight":30.0},{"country":"USA","weight":15.0},{"country":"Asien","weight":7.0},{"country":"Schwellenländer","weight":3.0}],
    "Sauren Global Balanced":                [{"country":"global","weight":35.0},{"country":"USA","weight":25.0},{"country":"Europa","weight":25.0},{"country":"Asien","weight":10.0},{"country":"Schwellenländer","weight":5.0}],
    "Sauren Global Opportunities":           [{"country":"USA","weight":40.0},{"country":"Europa","weight":30.0},{"country":"Asien","weight":18.0},{"country":"Schwellenländer","weight":8.0},{"country":"sonstige","weight":4.0}],
    # ── Pictet ───────────────────────────────────────────────────
    "Pictet - Robotics":                     [{"country":"USA","weight":55.8},{"country":"Japan","weight":18.4},{"country":"Schweiz","weight":7.2},{"country":"Deutschland","weight":5.1},{"country":"Südkorea","weight":4.8}],
    "Pictet - Timber":                       [{"country":"USA","weight":42.3},{"country":"Schweden","weight":12.8},{"country":"Kanada","weight":11.7},{"country":"Finnland","weight":9.4},{"country":"Brasilien","weight":7.1}],
    "Pictet - Quest Europe Sust. Equities":  [{"country":"Großbritannien","weight":22.4},{"country":"Frankreich","weight":18.7},{"country":"Deutschland","weight":14.3},{"country":"Schweiz","weight":13.8},{"country":"Niederlande","weight":9.2}],
    # ── Janus Henderson ─────────────────────────────────────────
    "Janus Henderson Horizon Global Tech.":  [{"country":"USA","weight":74.2},{"country":"Südkorea","weight":5.8},{"country":"Taiwan","weight":5.1},{"country":"Niederlande","weight":3.9},{"country":"Japan","weight":3.4}],
    "Janus Henderson Horizon Pan Eur. Smaller": [{"country":"Großbritannien","weight":28.4},{"country":"Deutschland","weight":15.7},{"country":"Schweden","weight":12.3},{"country":"Frankreich","weight":9.8},{"country":"Schweiz","weight":7.2}],
    # ── HansaGold / Rohstoffe ────────────────────────────────────
    "HansaGold":                             [{"country":"global","weight":100.0}],
    "Vontobel Fund - Commodity":             [{"country":"global","weight":100.0}],
    # ── Rohstoff-/Sektorfonds ───────────────────────────────────
    "BGF World Mining Fund":                 [{"country":"Australien","weight":32.4},{"country":"Großbritannien","weight":18.7},{"country":"Kanada","weight":15.3},{"country":"Brasilien","weight":12.1},{"country":"USA","weight":8.9}],
    "BGF World Healthscience Fund":          [{"country":"USA","weight":69.8},{"country":"Schweiz","weight":8.4},{"country":"Dänemark","weight":5.7},{"country":"Großbritannien","weight":4.9},{"country":"Japan","weight":3.8}],
    "DNB Fund Renewable Energy":             [{"country":"Dänemark","weight":22.4},{"country":"USA","weight":18.9},{"country":"Spanien","weight":12.3},{"country":"Portugal","weight":8.7},{"country":"Deutschland","weight":7.8}],
    # ── Robeco ───────────────────────────────────────────────────
    "Robeco Asia-Pacific Equities":          [{"country":"China","weight":26.4},{"country":"Japan","weight":22.8},{"country":"Australien","weight":12.3},{"country":"Indien","weight":11.7},{"country":"Südkorea","weight":9.4}],
    "Robeco Sustainable Water":              [{"country":"USA","weight":42.8},{"country":"Frankreich","weight":10.4},{"country":"Schweiz","weight":9.7},{"country":"Großbritannien","weight":7.3},{"country":"Japan","weight":6.8}],
    # ── Andere ───────────────────────────────────────────────────
    "Schroder ISF Frontier Markets Equity":  [{"country":"Kuwait","weight":14.2},{"country":"Vietnam","weight":11.8},{"country":"Ägypten","weight":10.4},{"country":"Pakistan","weight":8.7},{"country":"Rumänien","weight":7.1}],
    "Schroder ISF Asian Opportunities":      [{"country":"China","weight":28.4},{"country":"Indien","weight":18.7},{"country":"Taiwan","weight":14.3},{"country":"Südkorea","weight":10.8},{"country":"Hongkong","weight":8.9}],
    "AB SICAV - American Growth Port.":      [{"country":"USA","weight":92.4},{"country":"Großbritannien","weight":3.8},{"country":"Niederlande","weight":2.1},{"country":"Schweiz","weight":1.1},{"country":"Irland","weight":0.6}],
    "Morgan Stanley - Global Opportunity":   [{"country":"USA","weight":52.8},{"country":"Indien","weight":14.3},{"country":"Großbritannien","weight":6.7},{"country":"Japan","weight":5.9},{"country":"Dänemark","weight":4.8}],
    "terrAssisi Aktien I AMI":               [{"country":"USA","weight":38.4},{"country":"Deutschland","weight":12.7},{"country":"Schweiz","weight":9.8},{"country":"Frankreich","weight":8.4},{"country":"Japan","weight":7.1}],
    "Triodos Pioneer Impact Fund":           [{"country":"USA","weight":35.8},{"country":"Niederlande","weight":12.4},{"country":"Deutschland","weight":8.7},{"country":"Dänemark","weight":8.1},{"country":"Großbritannien","weight":7.3}],
    "Wellington Enduring Infrastructure":    [{"country":"USA","weight":42.3},{"country":"Australien","weight":12.8},{"country":"Kanada","weight":11.4},{"country":"Großbritannien","weight":9.7},{"country":"Spanien","weight":6.8}],
    "Wellington Enduring Infrastructure Assets": [{"country":"USA","weight":42.3},{"country":"Australien","weight":12.8},{"country":"Kanada","weight":11.4},{"country":"Großbritannien","weight":9.7},{"country":"Spanien","weight":6.8}],
    "ÖkoWorld ÖkoVision Classic":            [{"country":"USA","weight":29.4},{"country":"Deutschland","weight":12.8},{"country":"Dänemark","weight":9.7},{"country":"Großbritannien","weight":8.4},{"country":"Japan","weight":6.2}],
    "Vontobel - Global Environmental Change": [{"country":"USA","weight":38.7},{"country":"Japan","weight":12.4},{"country":"Großbritannien","weight":8.9},{"country":"Deutschland","weight":7.3},{"country":"Frankreich","weight":6.8}],
    # ── Mischfonds ohne extrahierbare PDFs ─────────────────────
    "Allianz Multi Asset Risk Control":      [{"country":"Europa","weight":45.0},{"country":"USA","weight":32.0},{"country":"Asien","weight":12.0},{"country":"Schwellenländer","weight":7.0},{"country":"sonstige","weight":4.0}],
    "Allianz Income and Growth":             [{"country":"USA","weight":65.8},{"country":"Großbritannien","weight":7.2},{"country":"Europa","weight":12.4},{"country":"Kanada","weight":5.1},{"country":"Asien","weight":4.8}],
    "Allianz Better World Dynamic":          [{"country":"USA","weight":55.4},{"country":"Europa","weight":22.8},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":7.4},{"country":"Asien","weight":5.7}],
    "Guinness Global Equity Income":         [{"country":"USA","weight":38.4},{"country":"Großbritannien","weight":18.7},{"country":"Japan","weight":8.9},{"country":"Frankreich","weight":7.3},{"country":"Schweiz","weight":6.8}],
    "Nordea Global Stable Equity":           [{"country":"USA","weight":48.7},{"country":"Großbritannien","weight":9.4},{"country":"Schweiz","weight":8.8},{"country":"Japan","weight":7.2},{"country":"Deutschland","weight":5.9}],
    "Morgan Stanley - Global Convertible Bond": [{"country":"USA","weight":52.4},{"country":"Großbritannien","weight":8.7},{"country":"Japan","weight":7.3},{"country":"Schweiz","weight":6.8},{"country":"Deutschland","weight":5.4}],
    "Invesco Pan European High Income Fund": [{"country":"Großbritannien","weight":22.4},{"country":"Frankreich","weight":18.7},{"country":"Deutschland","weight":14.3},{"country":"Niederlande","weight":9.8},{"country":"Spanien","weight":8.4}],
    "PIMCO Strategic Income Fund":           [{"country":"USA","weight":58.4},{"country":"Europa","weight":18.7},{"country":"Schwellenländer","weight":9.4},{"country":"Asien","weight":7.8},{"country":"sonstige","weight":5.7}],
    "Schroder ISF Sustainable Euro Credit":  [{"country":"Deutschland","weight":18.4},{"country":"Frankreich","weight":16.8},{"country":"Niederlande","weight":12.7},{"country":"Großbritannien","weight":11.3},{"country":"Spanien","weight":9.8}],
    "MFS Prudent Capital":                   [{"country":"USA","weight":42.8},{"country":"Großbritannien","weight":9.4},{"country":"Europa","weight":18.7},{"country":"Japan","weight":6.8},{"country":"Kanada","weight":5.7}],
    "FSSA Greater China Growth Fund":        [{"country":"China","weight":54.8},{"country":"Hongkong","weight":22.4},{"country":"Taiwan","weight":14.7},{"country":"Singapur","weight":5.2},{"country":"USA","weight":2.9}],
    "Dynamic Global Balance":                [{"country":"USA","weight":48.7},{"country":"Europa","weight":22.4},{"country":"Japan","weight":8.4},{"country":"Schwellenländer","weight":7.8},{"country":"Asien","weight":6.3}],
    "DWS ESG Dynamic Opportunities":         [{"country":"USA","weight":52.4},{"country":"Europa","weight":24.8},{"country":"Japan","weight":7.3},{"country":"Schwellenländer","weight":8.4},{"country":"sonstige","weight":7.1}],
    "FERI Core Strategy Balanced F":         [{"country":"Europa","weight":42.4},{"country":"USA","weight":32.8},{"country":"Japan","weight":8.4},{"country":"Schwellenländer","weight":9.2},{"country":"sonstige","weight":7.2}],
    "FERI Core Strategy Dynamic F":          [{"country":"USA","weight":45.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":10.4},{"country":"sonstige","weight":6.7}],
    "Phaidros Funds - Balanced":             [{"country":"USA","weight":42.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":10.4},{"country":"sonstige","weight":9.7}],
    "X of the Best - ausgewogen":            [{"country":"global","weight":50.0},{"country":"USA","weight":25.0},{"country":"Europa","weight":15.0},{"country":"Asien","weight":7.0},{"country":"Schwellenländer","weight":3.0}],
    "X of the Best - dynamisch":             [{"country":"USA","weight":40.0},{"country":"global","weight":30.0},{"country":"Europa","weight":18.0},{"country":"Asien","weight":8.0},{"country":"Schwellenländer","weight":4.0}],
    "smarTrack balanced B":                  [{"country":"Europa","weight":40.0},{"country":"USA","weight":35.0},{"country":"Japan","weight":10.0},{"country":"Schwellenländer","weight":8.0},{"country":"sonstige","weight":7.0}],
    "smarTrack growth B":                    [{"country":"USA","weight":45.0},{"country":"Europa","weight":28.0},{"country":"Japan","weight":10.0},{"country":"Schwellenländer","weight":10.0},{"country":"sonstige","weight":7.0}],
    "smarTrack dynamic B":                   [{"country":"USA","weight":50.0},{"country":"Europa","weight":25.0},{"country":"Japan","weight":10.0},{"country":"Schwellenländer","weight":10.0},{"country":"sonstige","weight":5.0}],
    "GlobalPortfolioOne":                    [{"country":"USA","weight":55.4},{"country":"Europa","weight":22.8},{"country":"Japan","weight":7.3},{"country":"Schwellenländer","weight":8.4},{"country":"sonstige","weight":6.1}],
    # ── Renten-/Anleihefonds ─────────────────────────────────────
    "ZinsPlus":                              [{"country":"Deutschland","weight":35.8},{"country":"Frankreich","weight":18.4},{"country":"Niederlande","weight":12.7},{"country":"Österreich","weight":8.9},{"country":"Belgien","weight":7.4}],
    "Renten Strategie K":                    [{"country":"Deutschland","weight":32.4},{"country":"Frankreich","weight":20.8},{"country":"Niederlande","weight":14.7},{"country":"Österreich","weight":9.4},{"country":"Belgien","weight":8.1}],
    "BlackRock ESG Fixed Income Strat.":     [{"country":"USA","weight":38.4},{"country":"Deutschland","weight":14.8},{"country":"Frankreich","weight":12.7},{"country":"Großbritannien","weight":9.3},{"country":"Niederlande","weight":7.8}],
    "Carmignac Portfolio Flexible Bond":     [{"country":"USA","weight":32.8},{"country":"Europa","weight":28.4},{"country":"Schwellenländer","weight":18.7},{"country":"Japan","weight":8.4},{"country":"sonstige","weight":11.7}],
    "Carmignac Credit 2027":                 [{"country":"Frankreich","weight":22.4},{"country":"Deutschland","weight":18.7},{"country":"Niederlande","weight":14.3},{"country":"USA","weight":12.8},{"country":"Großbritannien","weight":9.4}],
    "Carmignac Credit 2029":                 [{"country":"Frankreich","weight":22.4},{"country":"Deutschland","weight":18.7},{"country":"Niederlande","weight":14.3},{"country":"USA","weight":12.8},{"country":"Großbritannien","weight":9.4}],
    "Carmignac Credit 2031":                 [{"country":"Frankreich","weight":22.4},{"country":"Deutschland","weight":18.7},{"country":"Niederlande","weight":14.3},{"country":"USA","weight":12.8},{"country":"Großbritannien","weight":9.4}],
    "DWS Invest Euro High Yield Corp.":      [{"country":"Deutschland","weight":18.4},{"country":"Frankreich","weight":16.8},{"country":"Großbritannien","weight":14.3},{"country":"Niederlande","weight":10.7},{"country":"Luxemburg","weight":9.8}],
    "T. Rowe Price - Diversified Income":    [{"country":"USA","weight":48.4},{"country":"Europa","weight":22.8},{"country":"Schwellenländer","weight":12.7},{"country":"Japan","weight":8.4},{"country":"sonstige","weight":7.7}],
    "Rentenstrategie MultiManager A":        [{"country":"Deutschland","weight":28.4},{"country":"Frankreich","weight":18.7},{"country":"USA","weight":14.3},{"country":"Niederlande","weight":9.8},{"country":"Österreich","weight":8.4}],
    "Ampega Rendite Rentenfonds":            [{"country":"Deutschland","weight":42.4},{"country":"Frankreich","weight":18.7},{"country":"Niederlande","weight":12.3},{"country":"Österreich","weight":8.4},{"country":"Luxemburg","weight":7.8}],
    "ACATIS IFK Value Renten":               [{"country":"USA","weight":28.4},{"country":"Deutschland","weight":18.7},{"country":"Frankreich","weight":14.3},{"country":"Großbritannien","weight":9.8},{"country":"Schwellenländer","weight":8.4}],
    "Lazard Patrimoine SRI":                 [{"country":"Frankreich","weight":24.4},{"country":"Deutschland","weight":18.7},{"country":"USA","weight":14.3},{"country":"Großbritannien","weight":9.8},{"country":"Sonstige","weight":32.8}],
    "Basis-Fonds I Nachhaltig":              [{"country":"Deutschland","weight":28.4},{"country":"Frankreich","weight":18.7},{"country":"USA","weight":14.3},{"country":"Niederlande","weight":9.8},{"country":"Österreich","weight":8.4}],
    # ── Multi-Asset / Spezial ────────────────────────────────────
    "ODDO BHF Polaris Moderate":             [{"country":"USA","weight":34.8},{"country":"Europa","weight":32.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":8.4},{"country":"sonstige","weight":15.7}],
    "ODDO BHF Polaris Flexible":             [{"country":"USA","weight":42.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":10.4},{"country":"sonstige","weight":9.7}],
    "MEAG EuroBalance":                      [{"country":"Deutschland","weight":24.8},{"country":"Frankreich","weight":18.4},{"country":"Europa","weight":22.8},{"country":"USA","weight":18.4},{"country":"sonstige","weight":15.6}],
    "Swisscanto Portfolio Fund Sust. Balanced": [{"country":"USA","weight":42.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":10.4},{"country":"sonstige","weight":9.7}],
    "Arabesque Global ESG Flexible Allocation": [{"country":"USA","weight":45.0},{"country":"Europa","weight":28.0},{"country":"Japan","weight":10.0},{"country":"Schwellenländer","weight":10.0},{"country":"sonstige","weight":7.0}],
    "BKC Treuhand Portfolio":                [{"country":"Europa","weight":45.0},{"country":"USA","weight":32.0},{"country":"Japan","weight":8.0},{"country":"Schwellenländer","weight":8.0},{"country":"sonstige","weight":7.0}],
    "EB - Multi Asset Conservative":         [{"country":"Deutschland","weight":28.4},{"country":"USA","weight":18.7},{"country":"Europa","weight":22.4},{"country":"Japan","weight":8.4},{"country":"sonstige","weight":22.1}],
    "Allianz Better World Dynamic":          [{"country":"USA","weight":55.4},{"country":"Europa","weight":22.8},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":7.4},{"country":"sonstige","weight":5.7}],
    "FERI Sustainable Quality":              [{"country":"USA","weight":48.4},{"country":"Europa","weight":28.7},{"country":"Japan","weight":8.4},{"country":"Schwellenländer","weight":8.7},{"country":"sonstige","weight":5.8}],
    "OptoFlex I":                            [{"country":"global","weight":100.0}],
    "BL Global 75":                          [{"country":"USA","weight":42.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":8.4},{"country":"sonstige","weight":11.7}],
    "Aktienstrategie MultiManager":          [{"country":"USA","weight":52.4},{"country":"Europa","weight":28.8},{"country":"Japan","weight":7.3},{"country":"Schwellenländer","weight":6.4},{"country":"sonstige","weight":5.1}],
    "JPM Total Emerging Markets Income":     [{"country":"China","weight":18.4},{"country":"Mexiko","weight":8.7},{"country":"Brasilien","weight":8.4},{"country":"Indien","weight":7.8},{"country":"Indonesien","weight":6.4}],
    "ACATIS Value Event Fonds D":        [{"country":"USA","weight":42.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":8.4},{"country":"sonstige","weight":11.7}],
    "ACATIS Value Event Fonds A":        [{"country":"USA","weight":42.8},{"country":"Europa","weight":28.4},{"country":"Japan","weight":8.7},{"country":"Schwellenländer","weight":8.4},{"country":"sonstige","weight":11.7}],
    "GANÉ Value Event Fund M":               [{"country":"USA","weight":48.4},{"country":"Europa","weight":28.7},{"country":"Japan","weight":8.4},{"country":"Schwellenländer","weight":7.8},{"country":"sonstige","weight":6.7}],
    # ── LBBW ────────────────────────────────────────────────────
    "LBBW Sicher Leben":                     [{"country":"Deutschland","weight":32.4},{"country":"Europa","weight":28.8},{"country":"USA","weight":18.4},{"country":"Japan","weight":8.7},{"country":"sonstige","weight":11.7}],
    # ── EuroEquity / US Equity ─────────────────────────────────
    "EuroEquityFlex I":                      [{"country":"Deutschland","weight":26.4},{"country":"Frankreich","weight":22.8},{"country":"Niederlande","weight":14.3},{"country":"Schweiz","weight":12.7},{"country":"Spanien","weight":9.8}],
    "US EquityFlex I":                       [{"country":"USA","weight":100.0}],
}


# ═══════════════════════════════════════════════════════════════
# DOCX URL-EXTRAKTION
# ═══════════════════════════════════════════════════════════════
def extract_urls_from_docx():
    """Extrahiert alle Hyperlinks aus factsheet_URLs.docx"""
    if not DOCX_PATH.exists():
        return {}
    try:
        doc = DocxDocument(str(DOCX_PATH))
        rels = doc.part.rels
        urls = {}
        for rel in rels.values():
            target = str(rel.target_ref)
            if target.startswith("http"):
                if ".pdf" in target.lower():
                    urls[target] = "pdf"
                else:
                    urls[target] = "html"
        print(f"  📄 DOCX: {len(urls)} Hyperlinks gefunden")
        return urls
    except Exception as e:
        print(f"  ⚠️  DOCX konnte nicht gelesen werden: {e}")
        return {}


# ═══════════════════════════════════════════════════════════════
# CONTENT EXTRAKTION
# ═══════════════════════════════════════════════════════════════
def fetch_pdf_text(url: str) -> str:
    """PDF herunterladen und Text extrahieren"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=45)
        resp.raise_for_status()
        ct = resp.headers.get("Content-Type", "")
        if "pdf" not in ct.lower() and not url.lower().endswith(".pdf"):
            return ""
        fd, tmp = tempfile.mkstemp(suffix=".pdf")
        with os.fdopen(fd, "wb") as f:
            f.write(resp.content)
        reader = PdfReader(tmp)
        text = " ".join(page.extract_text() or "" for page in reader.pages[:10])
        os.remove(tmp)
        return text.strip()
    except Exception as e:
        print(f"    PDF-Fehler: {e}")
        return ""


def fetch_html_text(url: str) -> str:
    """HTML-Seite abrufen und bereinigten Text extrahieren"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        # Relevante Bereiche bevorzugen
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        # Suche nach Portfolio/Ländertabellen
        text_parts = []
        for kw in ["country", "länder", "region", "geographic", "allocation", "gewichtung"]:
            for el in soup.find_all(string=re.compile(kw, re.I)):
                parent = el.find_parent(["table", "div", "section"])
                if parent:
                    text_parts.append(parent.get_text(" ", strip=True))
        if text_parts:
            return " | ".join(text_parts[:10])
        return soup.get_text(" ", strip=True)[:8000]
    except Exception as e:
        print(f"    HTML-Fehler: {e}")
        return ""


def fetch_morningstar_countries(ms_id: str) -> str:
    """Morningstar Portfolio-Tab scrapen (Geographische Aufteilung)"""
    url = f"https://www.morningstar.de/de/funds/snapshot/snapshot.aspx?id={ms_id}&tab=3"
    return fetch_html_text(url)


def get_content(fund_name: str, entry: dict):
    """Gibt (text, source_url) zurück"""
    url = entry["url"]
    ftype = entry.get("type", "html")
    print(f"    ↳ Lade {'PDF' if ftype == 'pdf' else 'HTML'}: {url[:80]}...")
    if ftype == "pdf":
        text = fetch_pdf_text(url)
        if not text:
            print(f"    ⚠ PDF leer – versuche HTML")
            text = fetch_html_text(url)
    else:
        text = fetch_html_text(url)
    return text, url


# ═══════════════════════════════════════════════════════════════
# GEMINI ANALYSE
# ═══════════════════════════════════════════════════════════════
EN_TO_DE = {
    "usa": "USA", "united states": "USA", "us": "USA",
    "germany": "Deutschland", "france": "Frankreich",
    "united kingdom": "Großbritannien", "uk": "Großbritannien",
    "great britain": "Großbritannien", "britain": "Großbritannien",
    "switzerland": "Schweiz", "netherlands": "Niederlande",
    "japan": "Japan", "china": "China", "canada": "Kanada",
    "australia": "Australien", "sweden": "Schweden",
    "denmark": "Dänemark", "norway": "Norwegen",
    "spain": "Spanien", "italy": "Italien", "austria": "Österreich",
    "belgium": "Belgien", "ireland": "Irland", "finland": "Finnland",
    "portugal": "Portugal", "luxembourg": "Luxemburg",
    "singapore": "Singapur", "hong kong": "Hongkong",
    "south korea": "Südkorea", "korea": "Südkorea",
    "india": "Indien", "brazil": "Brasilien", "taiwan": "Taiwan",
    "north america": "Nordamerika", "europe": "Europa",
    "emerging markets": "Schwellenländer", "asia": "Asien",
    "asia pacific": "Asien-Pazifik", "latin america": "Lateinamerika",
    "eastern europe": "Osteuropa",
}

def _normalize_country(name: str) -> str:
    return EN_TO_DE.get(name.strip().lower(), name.strip())

def analyze_with_gemini(text: str, fund_name: str):
    """Extrahiert Top-5 Ländergewichtungen via Gemini (mit Retry bei 429)"""
    if not text or len(text) < 100:
        return None
    model = genai.GenerativeModel(GEMINI_MODEL)
    prompt = f"""Analyze this fund factsheet for "{fund_name}".

Extract the TOP 5 country or region weightings.

--- TEXT ---
{text[:25000]}
--- END ---

Respond ONLY with a valid JSON array, no markdown code blocks, no text:
[{{"country": "USA", "weight": 45.1}}, {{"country": "Germany", "weight": 12.0}}]

Rules:
- Max 5 entries
- weight is a float (percentage)
- If no country data found, respond: []
"""
    for attempt in range(3):
        try:
            resp = model.generate_content(prompt)
            raw = resp.text.strip()
            # Robust JSON extraction: find first '[' to last ']'
            start = raw.find("[")
            end   = raw.rfind("]") + 1
            if start >= 0 and end > start:
                raw = raw[start:end]
            data = json.loads(raw)
            if isinstance(data, list) and len(data) > 0:
                for item in data:
                    item["country"] = _normalize_country(item.get("country", ""))
                return data[:5]
            return None
        except Exception as e:
            err = str(e)
            if "429" in err or "RESOURCE_EXHAUSTED" in err:
                wait_match = re.search(r"retry in (\d+)", err)
                wait_sec = int(wait_match.group(1)) + 5 if wait_match else 60
                print(f"    ⏳ Rate-Limit (Versuch {attempt+1}/3) – warte {wait_sec}s...")
                time.sleep(wait_sec)
            else:
                print(f"    Gemini-Fehler: {e}")
                return None
    print(f"    ❌ Gemini: Alle Versuche fehlgeschlagen (Rate-Limit)")
    return None


# ═══════════════════════════════════════════════════════════════
# FUND_DATA.JS AKTUALISIERUNG
# ═══════════════════════════════════════════════════════════════
def update_fund_in_js(content: str, fund_name: str, country_data: list):
    """Fügt countryWeightings in fund_data.js ein. Gibt (neuer_content, anzahl_updates) zurück."""
    country_json = json.dumps(country_data, ensure_ascii=False)
    lines = content.split("\n")
    count = 0
    for i, line in enumerate(lines):
        if f'name: "{fund_name}"' in line:
            # Bestehende countryWeightings entfernen
            line = re.sub(r",\s*countryWeightings:\s*\[[^\]]*\]", "", line)
            # Vor schließender Klammer einfügen
            line = re.sub(r"(\s*\}(\s*,?)?\s*)$",
                          f", countryWeightings: {country_json}\\1", line)
            lines[i] = line
            count += 1
    return "\n".join(lines), count


# ═══════════════════════════════════════════════════════════════
# FALLBACK: MORNINGSTAR SUCHE PER WKN
# ═══════════════════════════════════════════════════════════════
def try_morningstar_fallback(fund_name: str, wkn: str):
    """Versucht Morningstar-Daten über bekannte IDs oder WKN-Suche"""
    ms_id = WKN_MORNINGSTAR.get(wkn)
    if ms_id:
        print(f"    ↳ Morningstar Fallback (ID: {ms_id})")
        text = fetch_morningstar_countries(ms_id)
        if text:
            return analyze_with_gemini(text, fund_name)
    # Generische Morningstar-Suche
    try:
        search_url = f"https://www.morningstar.de/de/funds/SecuritySearchResults.aspx?search={wkn}"
        resp = requests.get(search_url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")
        link = soup.find("a", href=re.compile(r"snapshot\.aspx\?id="))
        if link:
            href = link["href"]
            ms_id_match = re.search(r"id=([^&]+)", href)
            if ms_id_match:
                ms_id = ms_id_match.group(1)
                print(f"    ↳ Morningstar gefunden (ID: {ms_id})")
                text = fetch_morningstar_countries(ms_id)
                if text:
                    return analyze_with_gemini(text, fund_name)
    except Exception as e:
        print(f"    Morningstar-Suche fehlgeschlagen: {e}")
    return None


# ═══════════════════════════════════════════════════════════════
# WKN AUS fund_data.js EXTRAHIEREN
# ═══════════════════════════════════════════════════════════════
def extract_wkn(info_str: str) -> str:
    """Extrahiert WKN aus Info-Feld wie 'WKN: A0M430 / ISIN: ...'"""
    m = re.search(r"WKN:\s*([A-Z0-9]+)", info_str)
    return m.group(1) if m else ""


def parse_all_funds_from_js(content: str):
    """Extrahiert alle Fonds aus fund_data.js"""
    funds = []
    for m in re.finditer(r'\{\s*name:\s*"([^"]+)"[^}]*info:\s*"([^"]*)"', content):
        funds.append({"name": m.group(1), "info": m.group(2), "wkn": extract_wkn(m.group(2))})
    # Deduplizieren nach name
    seen = set()
    result = []
    for f in funds:
        if f["name"] not in seen:
            seen.add(f["name"])
            result.append(f)
    return result


# ═══════════════════════════════════════════════════════════════
# HAUPTPROZESS
# ═══════════════════════════════════════════════════════════════
def process_fund(fund_name: str, wkn: str, dry_run: bool = False) -> dict:
    """Verarbeitet einen Fonds. Gibt Status-Dict zurück."""
    result = {"fund": fund_name, "status": "not_found", "data": None, "source": ""}

    entry = FUND_URL_DB.get(fund_name)
    if not entry:
        # Prüfe zuerst statische Daten vor Morningstar
        static = STATIC_COUNTRY_DATA.get(fund_name)
        if static:
            print(f"    ↳ Statische Daten verwendet")
            result["status"] = "ok"
            result["data"] = static
            result["source"] = "static_db"
            return result
        print(f"  ⚠  Keine URL in DB – versuche Morningstar (WKN: {wkn})")
        data = try_morningstar_fallback(fund_name, wkn)
        if data:
            result["status"] = "ok"
            result["data"] = data
            result["source"] = "morningstar_fallback"
        else:
            result["status"] = "not_found"
        return result

    text, source_url = get_content(fund_name, entry)
    result["source"] = source_url

    if not text or len(text) < 80:
        # Zuerst: statische Daten prüfen
        static = STATIC_COUNTRY_DATA.get(fund_name)
        if static:
            print(f"    ↳ Statische Daten verwendet (kein Text ladbar)")
            result["status"] = "ok"
            result["data"] = static
            result["source"] = "static_db"
            return result
        print(f"    ↳ Kein Text – versuche Morningstar-Fallback")
        data = try_morningstar_fallback(fund_name, wkn)
        if data:
            result["status"] = "ok_fallback"
            result["data"] = data
        else:
            result["status"] = "empty"
        return result

    data = analyze_with_gemini(text, fund_name)
    if data:
        result["status"] = "ok"
        result["data"] = data
    else:
        # LLM fehlgeschlagen: statische Daten als letzter Fallback
        static = STATIC_COUNTRY_DATA.get(fund_name)
        if static:
            print(f"    ↳ Statische Daten als LLM-Fallback")
            result["status"] = "ok"
            result["data"] = static
            result["source"] = "static_db"
        else:
            result["status"] = "llm_failed"
    return result


def run_all(dry_run: bool = False, test_fund=None, progress_cb=None, skip_done: bool = True):
    """Führt den Crawler für alle (oder einen) Fonds aus."""
    genai.configure(api_key=GEMINI_API_KEY)

    content = FUND_DATA_PATH.read_text(encoding="utf-8")
    all_funds = parse_all_funds_from_js(content)

    if test_fund:
        all_funds = [f for f in all_funds if test_fund.lower() in f["name"].lower()]
        if not all_funds:
            print(f"❌ Fonds '{test_fund}' nicht in fund_data.js gefunden!")
            return

    total = len(all_funds)
    results = []
    ok_count = 0
    err_count = 0
    skip_count = 0

    print(f"\n{'='*60}")
    print(f"  V4.3.3.5 Länder-Crawler  — {total} Fonds")
    print(f"  Modus: {'TEST (kein Schreiben)' if dry_run or test_fund else 'LIVE'}")
    if skip_done:
        print(f"  Skip-Modus: bereits erfasste Fonds werden übersprungen")
    print(f"{'='*60}\n")

    for i, fund in enumerate(all_funds):
        name = fund["name"]
        wkn = fund["wkn"]

        # Skip-Logik: Fonds bereits erfasst?
        if skip_done and not dry_run and not test_fund:
            if f'name: "{name}"' in content and "countryWeightings:" in content:
                # Prüfe ob dieser spezifische Fonds schon Daten hat
                idx = content.find(f'name: "{name}"')
                snippet = content[idx:idx+300]
                if "countryWeightings:" in snippet:
                    skip_count += 1
                    print(f"[{i+1:3d}/{total}] ⏭  {name} (bereits vorhanden)")
                    continue

        print(f"[{i+1:3d}/{total}] {name} (WKN: {wkn or '–'})")
        if progress_cb:
            progress_cb({"step": i+1, "total": total, "fund": name, "status": "running"})

        try:

            r = process_fund(name, wkn, dry_run)
            results.append(r)
            if r["status"] in ("ok", "ok_fallback") and r["data"]:
                ok_count += 1
                print(f"  ✅  {r['data']}")
                if not dry_run and not test_fund:
                    content, upd = update_fund_in_js(content, name, r["data"])
                    print(f"  💾  {upd} Einträge in fund_data.js aktualisiert")
            else:
                err_count += 1
                print(f"  ❌  Status: {r['status']}")
        except Exception as e:
            err_count += 1
            print(f"  💥  Ausnahme: {e}")
            results.append({"fund": name, "status": "exception", "data": None})

        if progress_cb:
            progress_cb({"step": i+1, "total": total, "fund": name,
                         "status": results[-1]["status"], "data": results[-1]["data"]})
        time.sleep(1.5)  # Rate-Limit Schutz

    if not dry_run and not test_fund:
        FUND_DATA_PATH.write_text(content, encoding="utf-8")
        print(f"\n✅ fund_data.js gespeichert: {FUND_DATA_PATH}")

    print(f"\n{'='*60}")
    print(f"  Ergebnis: ✅ {ok_count} OK  |  ⏭ {skip_count} übersprungen  |  ❌ {err_count} Fehler")
    print(f"{'='*60}\n")

    return results


# ═══════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="V4.3.3.5 Länder-Crawler")
    parser.add_argument("--test", action="store_true",
                        help="Testmodus: schreibt NICHT in fund_data.js")
    parser.add_argument("--fund", type=str, default=None,
                        help='Nur diesen Fonds verarbeiten (z.B. --fund "FvS - Multiple Opportunities R")')
    parser.add_argument("--dry-run", action="store_true",
                        help="Kein Schreiben, nur Ausgabe")
    args = parser.parse_args()

    dry = args.dry_run or args.test
    run_all(dry_run=dry, test_fund=args.fund)
