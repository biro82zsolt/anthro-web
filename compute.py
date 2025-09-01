import os, re
import io
import zipfile
from datetime import datetime

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from models import Result
from anthro_calc import compute_all_metrics  # a Te kalkulátorod (ref_path-ot lásd lent)

# --- i18n: egyszerű kulcs-érték fordítások ---
TRANSLATIONS = {
    "hu": {
        "title": "Egyéni visszajelzés",
        "subtitle": "Sporttudományi / antropometriai összefoglaló",
        "name": "Név",
        "birth_date": "Születési dátum",
        "meas_date": "Mérés dátuma",
        "height": "Testmagasság",
        "weight": "Testsúly",
        "bmi": "BMI",
        "body_fat": "Testzsír",
        "final_height": "Várható végső magasság",
        "table_hdr_metric": "Mutató",
        "table_hdr_value": "Érték",
        "table_hdr_note": "Megjegyzés",
        "endomorphy": "Endomorfia",
        "mesomorphy": "Mezomorfia",
        "ectomorphy": "Ektomorfia",
        "phv": "PHV",
        "mk_corr": "MK (korrekció)",
        "multiplier": "szorzó",
        "plx": "PLX",
        "sum6": "Sum of 6 skinfolds",
        "section_notes": "Rövid értelmezés",
        "note_bmi": "BMI kategória",
        "note_phv": "Növekedési csúcs (PHV) státusz",
        "note_delta_height": "A várható végső magasság és az aktuális magasság különbsége",
        "disclaimer": "Ez a visszajelzés tájékoztató jellegű. A méréseket standardizált protokoll szerint érdemes ismételni.",
        # egységek
        "cm": "cm",
        "kg": "kg",
        "pct": "%"
    },
    "en": {
        "title": "Individual Feedback",
        "subtitle": "Sports science / anthropometry summary",
        "name": "Name",
        "birth_date": "Birth date",
        "meas_date": "Measurement date",
        "height": "Height",
        "weight": "Weight",
        "bmi": "BMI",
        "body_fat": "Body fat",
        "final_height": "Predicted adult height",
        "table_hdr_metric": "Metric",
        "table_hdr_value": "Value",
        "table_hdr_note": "Note",
        "endomorphy": "Endomorphy",
        "mesomorphy": "Mesomorphy",
        "ectomorphy": "Ectomorphy",
        "phv": "PHV",
        "mk_corr": "MK (correction)",
        "multiplier": "multiplier",
        "plx": "PLX",
        "sum6": "Sum of 6 skinfolds",
        "section_notes": "Short interpretation",
        "note_bmi": "BMI category",
        "note_phv": "Peak height velocity status",
        "note_delta_height": "Difference between predicted adult height and current height",
        "disclaimer": "This feedback is informational. Measurements should be repeated using a standardized protocol.",
        # units
        "cm": "cm",
        "kg": "kg",
        "pct": "%"
    },
}

def _localize_bmi_cat(cat: str, lang: str) -> str:
    if not cat:
        return ""
    cat = cat.strip().lower()
    if lang == "en":
        mapping = {
            "túlsúlyos": "overweight",
            "normális testsúly": "normal weight",
            "enyhe soványság": "mild thinness",
            "mérsékelt soványság": "moderate thinness",
            "súlyos soványság": "severe thinness",
        }
        return mapping.get(cat, cat)
    return cat

def _localize_phv_cat(cat: str, lang: str) -> str:
    if not cat:
        return ""
    cat = cat.strip().lower()
    if lang == "en":
        mapping = {
            "magas": "high",
            "alacsony": "low",
            "normál": "normal",
            "ismeretlen": "unknown",
        }
        return mapping.get(cat, cat)
    return cat

def _localize_somatotype(cat: str, kind: str, lang: str) -> str:
    """
    kind ∈ {"endo","mezo","ekto"} – HU → EN fordítás csak akkor, ha lang=="en".
    HU esetben az eredeti magyar szöveget adja vissza.
    """
    if not cat:
        return ""
    if lang != "en":
        return cat

    c = cat.strip().lower()

    if kind == "endo":
        mapping = {
            "hízásra hajlamos testalkat": "high propensity to gain fat",
            "hízásra közepes mértékben hajlamos testalkat": "moderate propensity to gain fat",
            "hízásra nem hajlamos testalkat": "low propensity to gain fat",
        }
    elif kind == "mezo":
        mapping = {
            "nagy mértékben fejleszthető izomzat": "high muscular development potential",
            "közepes mértékben fejleszthető izomzat": "moderate muscular development potential",
            "kis mértékben fejleszthető izomzat": "low muscular development potential",
        }
    elif kind == "ekto":
        mapping = {
            "kifejezetten nyúlánk alkat": "high linearity",
            "közepesen nyúlánk alkat": "moderate linearity",
            "alacsony fokú relatív nyúlánkság": "low linearity",
        }
    else:
        mapping = {}

    return mapping.get(c, cat)


def _t(key: str, lang: str = "hu") -> str:
    """Egyszerű i18n: kulcs alapján visszaadja a fordítást a megadott nyelven."""
    lang = (lang or "hu").lower()
    if lang not in ("hu", "en"):
        lang = "hu"
    return TRANSLATIONS.get(lang, {}).get(key, key)

def _get_first(obj, *names, default=None):
    """Több lehetséges attribútumnév közül az első létező érték."""
    for n in names:
        v = getattr(obj, n, None)
        if v is not None:
            return v
    return default

def _safe_date_str(d):
    try:
        if d is None:
            return "-"
        if hasattr(d, "strftime"):
            return d.strftime("%Y-%m-%d")
        s = str(d)
        return s[:10]
    except Exception:
        return str(d) if d is not None else "-"

# ---------- Excel feldolgozás ----------

def _to_date_like(x):
    """pandas Timestamp / datetime / date -> date vagy None"""
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    if hasattr(x, "date"):
        return x.date()
    return x

def _normalize_record_keys(rec: dict) -> dict:
    out = dict(rec)

    # 0) Zárójeles kódok felvétele külön kulcsként (pl. "Height (TTM)" -> out["TTM"] = érték)
    for k, v in list(out.items()):
        if k is None:
            continue
        m = re.search(r"\(([^)]+)\)\s*$", str(k))
        if m:
            code = m.group(1).strip()
            if code and code not in out and v not in (None, ""):
                out[code] = v

    # 1) Aliases – HU/EN oszlopnevek egységesítése a számoló függvény magyar kulcsaira
    aliases = {
        # név, nem
        "Name": "Név",
        "Sex": "Nem",

        # dátumok  ⟵ EZ HIÁNYZOTT
        "Birth date": "Születési dátum",
        "Measurement date": "Mérés dátuma",

        # TTM/TTS/ÜLŐ
        "Height": "TTM",
        "Height (TTM)": "TTM",
        "Stature": "TTM",               # gyakori angol szinonima
        "Weight": "TTS",
        "Body mass": "TTS",             # gyakori angol szinonima
        "Weight (TTS)": "TTS",
        "Sitting height": "ÜLŐ",
        "Sitting Height": "ÜLŐ",
        "Sitting height (cm)": "ÜLŐ",
    }
    for src, dst in aliases.items():
        if dst not in out and src in out and out[src] not in (None, ""):
            out[dst] = out[src]

    return out

def process_excel_to_results(xlsx_path: str, user_id: int):
    df = pd.read_excel(xlsx_path)
    rows = []
    for _, row in df.iterrows():
        rec = _normalize_record_keys(row.to_dict())

        name = rec.get("Név") or rec.get("Name") or "N/A"
        sex  = rec.get("Nem") or rec.get("Sex")
        birth = rec.get("Születési dátum") or rec.get("Birth date")
        meas  = rec.get("Mérés dátuma") or rec.get("Measurement date")

        # számítás a normalizált rekorddal
        res = compute_all_metrics(rec, sex=sex, ref_path="mk_components.xlsx")

        r = Result(
            user_id=user_id,
            name=name,  # <- biztos, hogy mindig van érték
            birth_date=_to_date_like(rec.get("Születési dátum")),
            meas_date=_to_date_like(rec.get("Mérés dátuma")),
            ttm=rec.get("TTM"),
            tts=rec.get("TTS"),
            ca_years=res.age_years,
            plx=res.plx,
            mk_raw=res.mk_raw,
            mk=res.mk,
            mk_corr_factor=res.mk_corr_factor,
            vttm=res.vttm,
            sum6=res.sum6,
            bodyfat_percent=res.bodyfat_percent,
            bmi=res.bmi,
            bmi_cat=res.bmi_cat,
            endo=res.endomorphy,
            endo_cat=res.endomorphy_cat,
            mezo=res.mesomorphy,
            mezo_cat=res.mesomorphy_cat,
            ekto=res.ectomorphy,
            ekto_cat=res.ectomorphy_cat,
            phv=res.phv,
            phv_cat=res.phv_cat,
        )
        rows.append(r)
    return rows

def export_results_excel(results: list[Result]):
    data = []
    for r in results:
        data.append({
            "Név": r.name,
            "Születési dátum": r.birth_date,
            "Mérés dátuma": r.meas_date,
            "Testmagasság (TTM)": r.ttm,
            "Testsúly (TTS)": r.tts,
            "Életkor (CA)": round(r.ca_years or 0, 2) if r.ca_years is not None else None,
            "PLX": r.plx,
            "MK (nyers)": r.mk_raw,
            "MK": r.mk,
            "MK szorzó": r.mk_corr_factor,
            "VTTM": r.vttm,
            "Sum of 6 skinfolds": r.sum6,
            "Testzsír %": r.bodyfat_percent,
            "BMI": r.bmi,
            "BMI kategória": r.bmi_cat,
            "Endomorfia": r.endo,
            "Endo kategória": r.endo_cat,
            "Mezomorfia": r.mezo,
            "Mezo kategória": r.mezo_cat,
            "Ektomorfia": r.ekto,
            "Ekto kategória": r.ekto_cat,
            "PHV": r.phv,
            "PHV kategória": r.phv_cat,
        })
    df = pd.DataFrame(data)

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    fname = f"results_{ts}.xlsx"
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue(), fname

# ---------- PDF export (szép egyoldalas lapok) ----------

_FONT_REGISTERED = False
def _ensure_font():
    """DejaVuSans regisztráció, ha elérhető (static/fonts/DejaVuSans.ttf)."""
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return
    here = os.path.dirname(os.path.abspath(__file__))  # <-- helyesen __file__
    font_path = os.path.join(here, "static", "fonts", "DejaVuSans.ttf")
    if os.path.isfile(font_path):
        try:
            pdfmetrics.registerFont(TTFont("DejaVuSans", font_path))
        except Exception:
            pass
    _FONT_REGISTERED = True

def _styles():
    ss = getSampleStyleSheet()
    has_djv = any(getattr(f, "faceName", "") == "DejaVuSans" for f in pdfmetrics._fonts.values())
    base_font = "DejaVuSans" if has_djv else "Helvetica"

    ss["Normal"].fontName = base_font
    ss["Normal"].fontSize = 10
    ss["Normal"].leading = 14

    ss["Title"].fontName = base_font
    ss["Title"].fontSize = 18
    ss["Title"].leading = 22

    ss["Heading2"].fontName = base_font
    ss["Heading2"].fontSize = 14
    ss["Heading2"].leading = 18

    badge = ParagraphStyle(
        "Badge",
        parent=ss["Normal"],
        backColor=colors.whitesmoke,
        textColor=colors.HexColor("#111827"),
        fontName=base_font,
        fontSize=11,
        leading=14,
        leftIndent=6, rightIndent=6, spaceBefore=4, spaceAfter=4
    )

    kpi = ParagraphStyle(
        "KPI",
        parent=ss["Normal"],
        fontName=base_font,
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#111827"),
    )

    smallmuted = ParagraphStyle(
        "SmallMuted",
        parent=ss["Normal"],
        fontSize=8,
        textColor=colors.HexColor("#6b7280"),
    )

    return ss, badge, kpi, smallmuted

def _fmt(value, decimals=2, empty="-"):
    try:
        if value is None:
            return empty
        if isinstance(value, int):
            return str(value)
        return f"{float(value):.{decimals}f}"
    except Exception:
        return str(value) if value is not None else empty

def _kpi_card(title, value, unit="", hint=None):
    label = f"<b>{title}</b><br/>{value}{(' ' + unit) if unit else ''}"
    if hint:
        label += f"<br/><font size='8' color='#6b7280'>{hint}</font>"
    return Paragraph(label, getSampleStyleSheet()["Normal"])

def _build_single_pdf(res, logo_path=None, lang="hu"):
    """Egy egyéni visszajelző PDF BytesIO-ként (HU/EN)."""
    _ensure_font()
    ss, badge, kpi, smallmuted = _styles()

    buff = io.BytesIO()
    doc = SimpleDocTemplate(
        buff, pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm, topMargin=16*mm, bottomMargin=16*mm
    )
    story = []

    # --- fejléc: meta balra, logó jobbra
    header_cells = []

    # bal: meta (név + dátumok) – BALRA igazítva
    name = _get_first(res, "name", "full_name", "student_name", default="-")
    birth = _get_first(res, "birth_date", "birth")
    meas  = _get_first(res, "measure_date", "meas_date")

    meta_html = (
        f"<para align='left'><font size='16'><b>{name}</b></font><br/>"
        f"{_t('birth_date', lang)}: <b>{_safe_date_str(birth)}</b><br/>"
        f"{_t('meas_date', lang)}: <b>{_safe_date_str(meas)}</b></para>"
    )
    meta_para = Paragraph(meta_html, ss["Normal"])

    # jobb: logó (ha van)
    logo_cell = [""]
    if logo_path and os.path.isfile(logo_path):
        try:
            img = Image(logo_path, width=28*mm, height=28*mm, kind='proportional')
            logo_cell = [img]
        except Exception:
            pass

    # táblázat: bal meta, jobb logó (jobb oldali cella jobbra igazítva)
    header = Table([[meta_para, logo_cell[0]]], colWidths=[None, 30*mm])
    header.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (1,0), (1,0), "RIGHT"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story += [header, Spacer(1, 8)]

    # címek
    story += [Paragraph(_t("title", lang), ss["Title"]), Spacer(1, 6)]
    story += [Paragraph(_t("subtitle", lang), smallmuted), Spacer(1, 10)]

    # --- KPI kártyák
    height_cm = _get_first(res, "height_cm", "ttm")
    weight_kg = _get_first(res, "weight_kg", "tts")
    bmi       = _get_first(res, "bmi")
    bf        = _get_first(res, "bodyfat_percent", "testzsir_percent")
    vtm       = _get_first(res, "vttm", "vtm")

    kpis = [
        _kpi_card(_t("height", lang), _fmt(height_cm, 1), _t("cm", lang)),
        _kpi_card(_t("weight", lang), _fmt(weight_kg, 1), _t("kg", lang)),
        _kpi_card(_t("bmi", lang),    _fmt(bmi, 1), hint=_get_first(res, "bmi_cat")),
        _kpi_card(_t("body_fat", lang), _fmt(bf, 1), _t("pct", lang)),
    ]
    if vtm is not None:
        kpis.append(_kpi_card(_t("final_height", lang), _fmt(vtm, 1), _t("cm", lang)))

    kpi_tbl = Table([kpis], colWidths=[None]*len(kpis))
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f9fafb")),
        ("BOX", (0,0), (-1,0), 0.5, colors.HexColor("#e5e7eb")),
        ("INNERGRID", (0,0), (-1,0), 0.5, colors.HexColor("#e5e7eb")),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story += [kpi_tbl, Spacer(1, 12)]

    # --- részletes táblázat
    table_data = [[
        _t("table_hdr_metric", lang),
        _t("table_hdr_value", lang),
        _t("table_hdr_note", lang),
    ]]

    # lokalizált kategória-szövegek az EN riporthoz
    endo_note = _localize_somatotype(_get_first(res, "endomorphy_cat", "endo_cat", default=""), "endo", lang)
    mezo_note = _localize_somatotype(_get_first(res, "mesomorphy_cat", "mezo_cat", default=""), "mezo", lang)
    ekto_note = _localize_somatotype(_get_first(res, "ectomorphy_cat", "ekto_cat", default=""), "ekto", lang)

    data_rows = []
    data_rows.append([
        _t("endomorphy", lang),
        _fmt(_get_first(res, "endomorphy", "endo"), 2),
        endo_note
    ])
    data_rows.append([
        _t("mesomorphy", lang),
        _fmt(_get_first(res, "mesomorphy", "mezo"), 2),
        mezo_note
    ])
    data_rows.append([
        _t("ectomorphy", lang),
        _fmt(_get_first(res, "ectomorphy", "ekto"), 2),
        ekto_note
    ])
    data_rows.append([
        _t("phv", lang),
        _fmt(_get_first(res, "phv"), 2),
        _localize_phv_cat(_get_first(res, "phv_cat", default=""), lang)
    ])
    data_rows.append([
        _t("mk_corr", lang),
        _fmt(_get_first(res, "mk"), 2),
        f"{_t('multiplier', lang)}: {_fmt(_get_first(res, 'mk_corr_factor'), 2)}"
    ])
    data_rows.append([_t("plx", lang), _fmt(_get_first(res, "plx"), 2), ""])
    data_rows.append([_t("sum6", lang), _fmt(_get_first(res, "sum6"), 2), ""])

    table_data += data_rows

    tbl = Table(table_data, colWidths=[60*mm, 35*mm, None])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2563eb")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "DejaVuSans" if any(getattr(f, "faceName", "") == "DejaVuSans" for f in pdfmetrics._fonts.values()) else "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 11),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#e5e7eb")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f9fafb")]),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story += [tbl, Spacer(1, 12)]

    # --- rövid értelmezés
    tips = []
    bmi_cat_raw = _get_first(res, "bmi_cat")
    bmi_cat_loc = _localize_bmi_cat(bmi_cat_raw, lang)
    if bmi_cat_loc:
        tips.append(f"{_t('note_bmi', lang)}: <b>{bmi_cat_loc}</b>.")

    phv_cat_raw = _get_first(res, "phv_cat")
    phv_cat_loc = _localize_phv_cat(phv_cat_raw, lang)
    if phv_cat_loc:
        tips.append(f"{_t('note_phv', lang)}: <b>{phv_cat_loc}</b>.")

    if vtm is not None and height_cm is not None:
        try:
            diff = float(vtm) - float(height_cm)
            tips.append(f"{_t('note_delta_height', lang)}: <b>{_fmt(diff, 1)} cm</b>.")
        except Exception:
            pass

    if tips:
        story += [Paragraph(_t("section_notes", lang), ss["Heading2"]), Spacer(1, 4)]
        story += [Paragraph(" ".join(tips), ss["Normal"]), Spacer(1, 6)]

    story += [Spacer(1, 6), Paragraph("—", smallmuted)]
    story += [Paragraph(_t("disclaimer", lang), smallmuted)]

    doc.build(story)
    buff.seek(0)
    return buff

def export_results_pdfs(results, lang="hu"):
    """
    Egyéni PDF-ek készítése és ZIP-be csomagolása a megadott nyelven (hu/en).
    """
    here = os.path.dirname(os.path.abspath(__file__))
    logo_path = os.path.join(here, "static", "logo.png")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            raw_name = _get_first(r, "name", "full_name", "student_name", default="unknown")
            safe = "".join(ch for ch in str(raw_name) if ch.isalnum() or ch in (" ", "-", "_")).strip() or "unknown"

            pdf_io = _build_single_pdf(
                r,
                logo_path=logo_path if os.path.isfile(logo_path) else None,
                lang=lang
            )
            zf.writestr(f"{safe}.pdf", pdf_io.read())

    zip_buf.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return zip_buf.read(), f"feedback_{lang}_{timestamp}.zip"
