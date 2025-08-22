import os
import io
import zipfile
from datetime import datetime
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from models import Result
from anthro_calc import compute_all_metrics  # <- a Te kalkulátorod

# A referenciafájl az anthro_calc-on belül van hivatkozva ref_path-ként.
# Ha máshol van, add át a compute_all_metrics(..., ref_path="mk_components.xlsx") paramétert.

def _safe_date(x):
    # pandas Timestamp / datetime / date -> date
    if pd.isna(x) or x is None:
        return None
    if hasattr(x, "date"):
        return x.date()
    return x

def process_excel_to_results(xlsx_path: str, user_id: int):
    df = pd.read_excel(xlsx_path)
    rows = []
    for _, row in df.iterrows():
        rec = row.to_dict()

        # kötelező mezők (engedékenyen kezelheted, ha szeretnéd)
        name = rec.get("Név") or rec.get("Name") or "N/A"

        res = compute_all_metrics(rec, sex=rec.get("Nem"), ref_path="mk_components.xlsx")

        r = Result(
            user_id=user_id,
            name=name,
            birth_date=_safe_date(rec.get("Születési dátum")),
            meas_date=_safe_date(rec.get("Mérés dátuma")),
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

    # dátumos fájlnév
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    fname = f"results_{ts}.xlsx"
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue(), fname

def _draw_line(c: canvas.Canvas, x, y, label, value):
    c.drawString(x, y, f"{label}: {value}")

def _result_pdf_bytes(r: Result):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    x = 40
    y = h - 60
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x, y, f"Anthropometry Feedback / Visszajelzés")
    y -= 24
    c.setFont("Helvetica", 11)
    _draw_line(c, x, y, "Name / Név", r.name); y -= 16
    _draw_line(c, x, y, "Birth date", r.birth_date); y -= 16
    _draw_line(c, x, y, "Measurement date", r.meas_date); y -= 16
    _draw_line(c, x, y, "Height (cm)", r.ttm); y -= 16
    _draw_line(c, x, y, "Weight (kg)", r.tts); y -= 24

    cols = [
        ("CA (years)", r.ca_years),
        ("PLX", r.plx),
        ("MK (raw)", r.mk_raw),
        ("MK", r.mk),
        ("VTTM", r.vttm),
        ("Sum of 6", r.sum6),
        ("Body fat %", r.bodyfat_percent),
        ("BMI", f"{(r.bmi or 0):.2f} ({r.bmi_cat})" if r.bmi is not None else None),
        ("Endomorphy", f"{(r.endo or 0):.2f} ({r.endo_cat})" if r.endo is not None else None),
        ("Mesomorphy", f"{(r.mezo or 0):.2f} ({r.mezo_cat})" if r.mezo is not None else None),
        ("Ectomorphy", f"{(r.ekto or 0):.2f} ({r.ekto_cat})" if r.ekto is not None else None),
        ("PHV", f"{(r.phv or 0):.2f} ({r.phv_cat})" if r.phv is not None else None),
    ]
    for label, val in cols:
        _draw_line(c, x, y, label, val)
        y -= 16
        if y < 60:
            c.showPage()
            y = h - 60

    c.showPage()
    c.save()
    return buf.getvalue()

def export_results_pdfs(results: list[Result]):
    # ZIP-be csomagolt, soronként generált PDF-ek (Név.pdf)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    zip_name = f"individual_feedback_{ts}.zip"
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            pdf_bytes = _result_pdf_bytes(r)
            # fájlnév: Nev.pdf
            base = (r.name or "result").replace("/", "_").replace("\\", "_")
            zf.writestr(f"{base}.pdf", pdf_bytes)
    return zbuf.getvalue(), zip_name
