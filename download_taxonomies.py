#!/usr/bin/env python3
"""
US 311613 - B4: Survey industry taxonomies beyond ICAO
Evidence 1 helper: download public copies of every openly available taxonomy
referenced in "US311613_B4_Industry_Taxonomies_Beyond_ICAO.docx".

Usage
    python download_taxonomies.py                  # download everything into ./B4_Taxonomy_Library
    python download_taxonomies.py --out D:\\Share\\B4
    python download_taxonomies.py --only 03 04     # only some folders (by number prefix)
    python download_taxonomies.py --skip-icao      # leave out the ICAO baseline folder
    python download_taxonomies.py --dry-run        # list what would be downloaded
    python download_taxonomies.py --force          # re-download files that already exist

Standard library only (Python 3.8+). No pip install needed.

Output
    <out>/<folder>/<NN>_<name>.pdf (or .html for web pages / portals)
    <out>/00_Index/download_manifest.csv   - one row per item: status, size, SHA-256, URL, date
    <out>/00_Index/paid_and_members_only.md - citation sheet for sources that cannot be downloaded

Notes
- "page" items are publisher pages or portals (no direct file). They are saved as .html
  snapshots so the file share still has a dated copy and the link; open them to get the file.
- Some publishers block scripted downloads or require a click-through (IOGP asks for a
  name/email). Those show as FAILED or PAGE in the manifest - download them by hand.
- Do not put paid standards (ATA, SAE, ISO, IEC, API) on the share unless covered by a license.
"""

import argparse
import csv
import datetime as dt
import hashlib
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

# ---------------------------------------------------------------------------
# Catalog: (id, folder, short_name, title, url, kind)   kind = "file" | "page"
# ---------------------------------------------------------------------------
F_ICAO = "01_ICAO_Baseline"
F_RISK = "02_Occurrence_Risk"
F_SYS = "03_System_Design"
F_HF = "04_Causal_HF"
F_INV = "05_Investigation_Reporting"
F_ADJ = "06_Adjacent_Industries"
F_RES = "08_Research"

CATALOG = [
    # --- ICAO baseline (reference copies) ---------------------------------
    ("B01", F_ICAO, "CICTT_Occurrence_Categories_v4.8", "CICTT Aviation Occurrence Categories v4.8 (May 2021)",
     "https://www.intlaviationstandards.org/Documents/OccurrenceCategoryDefinitions.pdf", "file"),
    ("B02", F_ICAO, "CICTT_Occurrence_Categories_v4.7", "CICTT Aviation Occurrence Categories v4.7 (Dec 2017)",
     "https://www.enac.gov.it/sites/default/files/allegati/2018-Lug/CICTT%20Occ.%20Cat.%20(v%204.7%20-%20Dec%202017).pdf", "file"),
    ("B03", F_ICAO, "CICTT_Occurrence_Categories_NTSB_copy", "CICTT Occurrence Categories (NTSB-hosted)",
     "https://www.ntsb.gov/safety/data/Documents/datafiles/OccurrenceCategoryDefinitions.pdf", "file"),
    ("B04", F_ICAO, "CICTT_Hazard_Taxonomy_v1.0", "CICTT Hazard Taxonomy v1.0",
     "https://sassofia.com/wp-content/uploads/2022/01/HazardDefinitions.pdf", "file"),
    ("B05", F_ICAO, "CICTT_Positive_Taxonomy_v1.2", "CICTT Positive Taxonomy v1.2",
     "https://omnisms.aero/wp-content/uploads/2016/12/CICTT-Positive-Taxonomy.pdf", "file"),
    ("B06", F_ICAO, "ECCAIRS_Coding_Guidance_Ch2", "ECCAIRS coding guidance, chapter 2",
     "https://aviationreporting.eu/sites/default/files/2022-07/ECG%20Chapter%202_v1.0.pdf", "file"),
    ("B07", F_ICAO, "ICAO_GASP_Doc10004_2026-2028", "ICAO Global Aviation Safety Plan 2026-2028 (Doc 10004)",
     "https://www.icao.int/sites/default/files/safety/GASP/Documents/Doc.10004.2026-2028.GASP_.EN_.pdf", "file"),

    # --- Occurrence risk / severity ---------------------------------------
    ("A02", F_RISK, "EU_2020-2034_ERCS", "Commission Delegated Regulation (EU) 2020/2034 - ERCS",
     "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX%3A32020R2034&from=EN", "file"),
    ("A03", F_RISK, "ARMS_Methodology_v4.1_2010", "ARMS Methodology for Operational Risk Assessment v4.1",
     "https://www.helicopter.su/assets/media_sources/ehest-ihts/2016/ARMS/2- SMS - ARMS_original EN.pdf", "file"),
    ("A03b", F_RISK, "ARMS_EASA_page", "EASA page for the ARMS methodology (links to SKYbrary copy)",
     "https://www.easa.europa.eu/en/node/24193", "page"),
    ("A04", F_RISK, "EASA_EPAS_Vol3_2025", "EASA EPAS Volume III - Safety Risk Portfolios, 2025 edition",
     "https://www.easa.europa.eu/sites/default/files/dfu/easa_mb_2024-02_-_wp16b_-_epas_vol._iii_-_2025_edition.pdf", "file"),
    ("A08", F_RISK, "FAA_Order_8040.4C", "FAA Order 8040.4C - Safety Risk Management Policy",
     "https://www.faa.gov/documentLibrary/media/Order/FAA_Order_8040.4C.pdf", "file"),
    ("A11", F_RISK, "MIL-STD-882E", "MIL-STD-882E - DoD Standard Practice: System Safety",
     "https://mail.system-safety.org/Documents/MIL-STD-882E.pdf", "file"),

    # --- System / design --------------------------------------------------
    ("A01", F_SYS, "FAA_JASC_Code_Table", "FAA JASC Code Table (ATA-based system codes)",
     "https://sdrs.faa.gov/documents/JASC_Code.pdf", "file"),
    ("A01b", F_SYS, "FAA_JASC_Code_Table_mirror", "FAA JASC Code Table (faa.gov mirror)",
     "https://www.faa.gov/sites/faa.gov/files/JASC_Code.pdf", "file"),
    ("A09", F_SYS, "FAA_AC_25.1309-1B", "FAA AC 25.1309-1B - System Design and Analysis (2024)",
     "https://www.faa.gov/documentLibrary/media/Advisory_Circular/AC_25.1309-1B.pdf", "file"),
    ("A10", F_SYS, "EASA_CS-25_Easy_Access_Rules_page", "EASA CS-25 Easy Access Rules (incl. AMC 25.1309)",
     "https://www.easa.europa.eu/en/document-library/easy-access-rules/easy-access-rules-large-aeroplanes-cs-25", "page"),

    # --- Causal / human factors -------------------------------------------
    ("A05", F_HF, "DoD_HFACS_7.0", "DoD HFACS v7.0",
     "https://www.dcms.uscg.mil/Portals/10/CG-1/cg113/docs/pdf/DoD_HFACS7.0.pdf?ver=2017-02-23-152408-007", "file"),
    ("A06", F_HF, "Boeing_MEDA_Results_Form_RevN", "MEDA Results Form Rev N (Feb 2022)",
     "https://www.faa.gov/about/initiatives/maintenance_hf/procedural_non-compliance/MEDA_Results_Form_rev_N_February_2022.pdf", "file"),
    ("A07", F_HF, "Boeing_MEDA_Users_Guide_2022", "MEDA User's Guide (Feb 2022)",
     "https://www.faa.gov/about/initiatives/maintenance_hf/procedural_non-compliance/MEDA_Users_Guide_rev_February_2022.pdf", "file"),
    ("A17", F_HF, "IATA_CBTA_WhitePaper_ACTF_Annex", "IATA CBTA White Paper (2024) - Annex 3 ACTF threats/errors/UAS",
     "https://www.iata.org/contentassets/c0f61fc821dc4f62bb6441d7abedb076/cbta-expansion-within-the-aviation-system.pdf", "file"),
    ("A18", F_HF, "IATA_Safety_Report_2024_Summary", "IATA Safety Report 2024 - executive summary",
     "https://www.iata.org/contentassets/a8e49941e8824a058fee3f5ae0c005d9/safety-report-executive-summary-and-safety-overview-2024_final.pdf", "file"),
    ("A19", F_HF, "IATA_LOC-I_Environmental_Factors", "IATA guidance: environmental factors affecting LOC-I",
     "https://flightsafety.org/wp-content/uploads/2017/07/IATA-guidance-loci-environmental-factors-affecting-loci-1st-edition.pdf", "file"),
    ("A20", F_HF, "CAST_SE219_Runway_Excursion", "CAST runway excursion JSAIT material (SKYbrary 2810)",
     "https://skybrary.aero/sites/default/files/bookshelf/2810.pdf", "file"),

    # --- Investigation / reporting ----------------------------------------
    ("A12", F_INV, "NTSB_Coding_Manual_legacy", "NTSB Aviation Coding Manual (pre-2008)",
     "https://www.ntsb.gov/GILS/Documents/codman.pdf", "file"),
    ("A13", F_INV, "NTSB_CAROL_page", "NTSB CAROL query and data download portal",
     "https://data.ntsb.gov/carol-main-public/basic-search", "page"),
    ("A14", F_INV, "NASA_TP-2006-213466_ASRS", "NASA/TP-2006-213466 (ASRS anomaly taxonomy reference)",
     "https://ntrs.nasa.gov/api/citations/20060023334/downloads/20060023334.pdf?attachment=true", "file"),
    ("A15", F_INV, "ASRS_Review_and_Analysis_Process", "ASRS review and analysis process (NTSB docket exhibit 17-B)",
     "https://data.ntsb.gov/Docket/Document/docBLOB?ID=19038204&FileExtension=pdf&FileName=17-B+Aviation+Safety+Reporting+System+Review+and+Analysis+Process-Rel.pdf", "file"),
    ("A16", F_INV, "NASA_ASRS_Database_page", "NASA ASRS online database",
     "https://asrs.arc.nasa.gov/search/database.html", "page"),
    ("A21", F_INV, "FAA_ASIAS_Turbulence_Study", "FAA ASIAS turbulence study (uses NTSB coding)",
     "https://www.asias.faa.gov/i/studies/turbulence_study_new.pdf", "file"),
    ("A22", F_INV, "Aviation_Safety_Network_page", "Aviation Safety Network database",
     "https://aviation-safety.net/database/", "page"),

    # --- Adjacent industries ----------------------------------------------
    ("A23", F_ADJ, "IAEA_INES_Users_Manual_2008", "IAEA INES User's Manual, 2008 Edition (English)",
     "https://www.bnra.bg/media/2021/05/ines-2009_web.pdf", "file"),
    ("A23b", F_ADJ, "IAEA_INES_page", "IAEA INES page",
     "https://www.iaea.org/resources/databases/international-nuclear-and-radiological-event-scale", "page"),
    ("A24", F_ADJ, "IAEA_IRS_Coding_Manual_2022", "Manual for IRS Coding, 2022 Edition (IAEA-SVS-20 Rev. 1)",
     "https://www-pub.iaea.org/MTCD/Publications/PDF/SVS-20(Rev1)web.pdf", "file"),
    ("A25", F_ADJ, "IMO_MSC.255(84)_Casualty_Code", "IMO Casualty Investigation Code - Res. MSC.255(84)",
     "https://wwwcdn.imo.org/localresources/en/OurWork/MSAS/Documents/Res.MSC.255(84)CasualtyIinvestigationCode.pdf", "file"),
    ("A26", F_ADJ, "EMSA_EMCIP_taxonomy", "EMSA EMCIP reporting attributes / taxonomy",
     "https://emsa.europa.eu/emcip/download/4719/3024/23.html", "file"),
    ("A26b", F_ADJ, "EMSA_EMCIP_related_publication", "EMSA related EMCIP publication",
     "https://www.emsa.europa.eu/publications/reports/download/8745/5844/23.html", "file"),
    ("A27", F_ADJ, "EU_Directive_2016-798_Rail_CSIs", "Directive (EU) 2016/798 - Annex I Common Safety Indicators",
     "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX:32016L0798", "file"),
    ("A28", F_ADJ, "ERA_Rail_Safety_Report_2022", "ERA Report on Railway Safety and Interoperability 2022",
     "https://www.era.europa.eu/system/files/2022-10/Report%20on%20Railway%20Safety%20and%20Interoperability%20in%20the%20EU%202022.pdf", "file"),
    ("A29", F_ADJ, "WHO_ICPS_page", "WHO Conceptual Framework for the ICPS v1.1 (publication page)",
     "https://www.who.int/publications/i/item/the-conceptual-framework-for-the-international-classification-for-patient-safety-(icps)", "page"),
    ("A30", F_ADJ, "AHRQ_Common_Formats_page", "AHRQ Common Formats (PSO PPC)",
     "https://www.psoppc.org/psoppc_web/publicpages/commonFormatsOverview", "page"),
    ("A31", F_ADJ, "IOGP_456_page", "IOGP Report 456 - Process safety KPIs (free download after registration)",
     "https://www.iogp.org/bookstore/product/process-safety-recommended-practice-on-key-performance-indicators/", "page"),
    ("A32", F_ADJ, "NASA_NPR_8621.1D_page", "NASA NPR 8621.1D - Mishap and Close Call Reporting",
     "https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPR&c=8621&s=1D", "page"),

    # --- Research papers --------------------------------------------------
    ("C01", F_RES, "Herr_ERAU_ASRS_to_HFACS_TEM", "Herr - ASRS to HFACS/TEM/ACCERS mapping (ERAU dissertation)",
     "https://commons.erau.edu/cgi/viewcontent.cgi?httpsredir=1&article=1025&context=edt", "file"),
    ("C02", F_RES, "WrightState_HFACS7_Reliability_2015", "Inter-rater reliability of DoD HFACS 7.0 (ISAP 2015)",
     "https://corescholar.libraries.wright.edu/cgi/viewcontent.cgi?article=1006&context=isap_2015", "file"),
    ("C03", F_RES, "USAARL_TR_2024-25", "USAARL Technical Report 2024-25",
     "https://usaarl.health.mil/assets/docs/techReports/2024-25.pdf", "file"),
    ("C04", F_RES, "Zhao_arXiv_2403.17914", "Hierarchy-aware classification of NTSB events",
     "https://arxiv.org/pdf/2403.17914", "file"),
    ("C05", F_RES, "NLR_arXiv_2301.05663", "BERT labeling of ASRS and CADORS (NLR)",
     "https://arxiv.org/pdf/2301.05663", "file"),
    ("C06", F_RES, "AviationKG_arXiv_2205.15952", "Aviation-KG knowledge graph",
     "https://arxiv.org/pdf/2205.15952", "file"),
    ("C07", F_RES, "FAA_DOT_ROSAP_93350", "FAA/DOT Aviation-BERT and knowledge graph report",
     "https://rosap.ntl.bts.gov/view/dot/93350/dot_93350_DS1.pdf", "file"),
    ("C08", F_RES, "GaTech_Aviation-BERT", "Aviation-BERT (Georgia Tech)",
     "https://repository.gatech.edu/server/api/core/bitstreams/94b3d7cf-a068-495a-96a2-d4e114bbd637/content", "file"),
    ("C09", F_RES, "WHO_ICPS_Framework_Paper_2009", "Towards an ICPS: the conceptual framework (PMC)",
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC2638753", "page"),
]

PAID = [
    ("ATA iSpec 2200", "Information Standards for Aviation Maintenance", "https://publications.airlines.org"),
    ("SAE ARP4761A", "Guidelines for Conducting the Safety Assessment Process", "https://www.sae.org/standards/content/arp4761a/"),
    ("SAE ARP4754B", "Guidelines for Development of Civil Aircraft and Systems", "https://www.sae.org/standards/content/arp4754b/"),
    ("ISO 26262-3:2018", "Road vehicles - Functional safety - Part 3", "https://www.iso.org/standard/68385.html"),
    ("ISO 21448:2022", "Road vehicles - SOTIF", "https://www.iso.org/standard/77490.html"),
    ("ISO 31000:2018", "Risk management - Guidelines", "https://www.iso.org/standard/65694.html"),
    ("IEC 31010:2019", "Risk assessment techniques", "https://webstore.iec.ch/publication/59809"),
    ("API RP 754 (3rd ed.)", "Process Safety Performance Indicators", "https://www.api.org"),
    ("IATA GADM / STEADES", "Members only", "https://www.iata.org/en/services/statistics/gadm/"),
    ("FAA ASIAS protected data", "Participants only", "https://www.asias.faa.gov"),
    ("RSSB Safe Insights", "Members only (replaced SMIS, Mar 2025)",
     "https://www.rssb.co.uk/about-rssb/groups-and-committees/safety/system-safety-risk-group/safe-insights-advisory-group"),
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

EXT_BY_TYPE = {
    "application/pdf": ".pdf",
    "text/html": ".html",
    "application/xhtml+xml": ".html",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/zip": ".zip",
}


def safe_url(url: str) -> str:
    """Percent-encode spaces and other unsafe characters, keep existing escapes."""
    parts = urllib.parse.urlsplit(url)
    path = urllib.parse.quote(parts.path, safe="/%()~-._!*'+,;=:@")
    query = urllib.parse.quote(parts.query, safe="=&%+:/?,;()")
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))


def fetch(url: str, timeout: int, retries: int):
    last = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(safe_url(url), headers={
                "User-Agent": UA,
                "Accept": "application/pdf,text/html;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=timeout) as r:
                ctype = (r.headers.get("Content-Type") or "").split(";")[0].strip().lower()
                return r.read(), ctype, r.geturl()
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
            if e.code in (401, 403, 404, 410):
                break  # not worth retrying
        except Exception as e:  # noqa: BLE001
            last = f"{type(e).__name__}: {e}"
        time.sleep(2 * attempt)
    raise RuntimeError(last or "unknown error")


def pick_ext(data: bytes, ctype: str, kind: str) -> str:
    if data[:5] == b"%PDF-":
        return ".pdf"
    if data[:4] == b"PK\x03\x04":
        return EXT_BY_TYPE.get(ctype, ".zip")
    if ctype in EXT_BY_TYPE:
        return EXT_BY_TYPE[ctype]
    return ".html" if kind == "page" else ".bin"


def existing(folder: str, stem: str):
    if not os.path.isdir(folder):
        return None
    for f in os.listdir(folder):
        if os.path.splitext(f)[0] == stem:
            return os.path.join(folder, f)
    return None


def write_paid_sheet(index_dir: str):
    path = os.path.join(index_dir, "paid_and_members_only.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("# Paid standards and members-only sources (cite only; do not redistribute)\n\n")
        f.write("| Taxonomy | Document | Where to obtain |\n|---|---|---|\n")
        for name, doc, url in PAID:
            f.write(f"| {name} | {doc} | {url} |\n")
    return path


def main():
    ap = argparse.ArgumentParser(description="Download public safety taxonomies for US 311613 (B4).")
    ap.add_argument("--out", default="B4_Taxonomy_Library", help="output folder (default: ./B4_Taxonomy_Library)")
    ap.add_argument("--only", nargs="*", help="folder number prefixes to include, e.g. 02 04")
    ap.add_argument("--skip-icao", action="store_true", help="skip the 01_ICAO_Baseline folder")
    ap.add_argument("--force", action="store_true", help="re-download files that already exist")
    ap.add_argument("--dry-run", action="store_true", help="list items without downloading")
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument("--retries", type=int, default=3)
    args = ap.parse_args()

    items = [c for c in CATALOG
             if not (args.skip_icao and c[1] == F_ICAO)
             and (not args.only or any(c[1].startswith(p) for p in args.only))]

    if args.dry_run:
        for iid, folder, name, title, url, kind in items:
            print(f"[{iid:>4}] {folder:<28} {kind:<4} {title}\n        {url}")
        print(f"\n{len(items)} items. Run without --dry-run to download.")
        return 0

    out = os.path.abspath(args.out)
    index_dir = os.path.join(out, "00_Index")
    os.makedirs(index_dir, exist_ok=True)
    os.makedirs(os.path.join(out, "07_Crosswalks"), exist_ok=True)  # team-owned, left empty

    rows, ok, pages, failed, skipped = [], 0, 0, 0, 0
    today = dt.date.today().isoformat()
    for n, (iid, folder, name, title, url, kind) in enumerate(items, 1):
        fdir = os.path.join(out, folder)
        os.makedirs(fdir, exist_ok=True)
        stem = f"{iid}_{name}"
        prefix = f"[{n:>2}/{len(items)}] {iid:<4}"

        prev = existing(fdir, stem)
        if prev and not args.force:
            print(f"{prefix} exists   {os.path.relpath(prev, out)}")
            skipped += 1
            data = open(prev, "rb").read()
            rows.append([iid, folder, title, "EXISTS", os.path.basename(prev), len(data),
                         hashlib.sha256(data).hexdigest(), url, "", today])
            continue

        try:
            data, ctype, final_url = fetch(url, args.timeout, args.retries)
            ext = pick_ext(data, ctype, kind)
            path = os.path.join(fdir, stem + ext)
            with open(path, "wb") as f:
                f.write(data)
            # A "file" that came back as HTML is usually a landing/consent page - flag it.
            if kind == "file" and ext == ".html":
                status, note = "CHECK", "expected a file but got a web page (consent/landing page?)"
                pages += 1
            elif kind == "page":
                status, note = "PAGE", "web page snapshot - open it to download the document"
                pages += 1
            else:
                status, note = "OK", ""
                ok += 1
            print(f"{prefix} {status:<8} {os.path.relpath(path, out)}  ({len(data)/1024:,.0f} KB)")
            rows.append([iid, folder, title, status, os.path.basename(path), len(data),
                         hashlib.sha256(data).hexdigest(), url,
                         note if final_url == safe_url(url) else f"{note} redirected to {final_url}".strip(), today])
        except Exception as e:  # noqa: BLE001
            failed += 1
            print(f"{prefix} FAILED   {title}  -> {e}")
            rows.append([iid, folder, title, "FAILED", "", 0, "", url, str(e), today])
        time.sleep(1)  # be polite to publishers

    manifest = os.path.join(index_dir, "download_manifest.csv")
    with open(manifest, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["id", "folder", "title", "status", "file", "bytes", "sha256", "source_url", "note", "date"])
        w.writerows(rows)
    paid = write_paid_sheet(index_dir)

    print("\n" + "-" * 70)
    print(f"Downloaded: {ok}   Pages/check: {pages}   Already present: {skipped}   Failed: {failed}")
    print(f"Library:    {out}")
    print(f"Manifest:   {manifest}")
    print(f"Paid list:  {paid}")
    if failed or pages:
        print("Open the manifest and download FAILED / CHECK / PAGE items by hand from source_url.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
