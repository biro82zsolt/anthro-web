from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, Any, Optional, Tuple
from datetime import datetime, date
import math
import pandas as pd
import numpy as np


# ====== BEÁLLÍTÁS: referencia CSV oszlopnevek ======
# A CSV-nek monoton (életkor szerinti) referencia-oszlopai legyenek, fiúk/lányok bontásban.
# A közelítő keresés (VLOOKUP approx) a "mért érték" szerint a legnagyobb nem nagyobb referenciaértékhez rendelt
# életkort adja vissza. Ha a te táblád másképp épül fel, csak ezt a mappinget igazítsd.
REF_COLUMNS = {
    # PLX -> „rendelt kor” (azaz age)
    "plx": {"value_col": "plx_ref", "age_col": "age"},

    # TTS/TTM -> „rendelt kor” (mindkettő age)
    "tts_boy":  {"value_col": "weight_ref", "age_col": "age"},
    "tts_girl": {"value_col": "weight_ref", "age_col": "age"},
    "ttm_boy":  {"value_col": "height_ref", "age_col": "age"},
    "ttm_girl": {"value_col": "height_ref", "age_col": "age"},

    # MK%: életkor (age) -> százalék (%)  — figyelj: az oszlop neve tényleg "%%"
    "mkpct": {"value_col": "age", "age_col": "%"},
}

# ====== Hasznos típus ======
@dataclass
class CalcResult:
    plx: float
    age_years: float
    mk_raw: float
    mk_corr_factor: float   # MK_kor% (szorzó)
    mk: float
    vttm: Optional[float]
    endomorphy: float
    endomorphy_cat: str
    mesomorphy: float
    mesomorphy_cat: str
    ectomorphy: float
    ectomorphy_cat: str
    sum6: float
    bodyfat_percent: Optional[float]
    bmi: float
    bmi_cat: str
    phv: float
    phv_cat: str


# ====== Segédfüggvények ======
def _to_float(x: Any) -> float:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return float("nan")
    if isinstance(x, str):
        x = x.replace(",", ".")
    return float(x)


def parse_date(s):
    """Fogad: str, pandas.Timestamp, numpy.datetime64, datetime, vagy None."""
    if s is None:
        raise ValueError("Hiányzó dátum.")
    # Pandas Timestamp
    if isinstance(s, pd.Timestamp):
        return s.to_pydatetime()
    # numpy datetime64
    if isinstance(s, np.datetime64):
        return pd.to_datetime(s).to_pydatetime()
    # már datetime
    if isinstance(s, (datetime, date)):
        # date -> datetime
        return datetime(s.year, s.month, s.day) if isinstance(s, date) and not isinstance(s, datetime) else s
    # string -> próbáljuk a formátumokat
    if isinstance(s, str):
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        # utolsó esély: pandas parse
        return pd.to_datetime(s).to_pydatetime()
    # bármi más típus
    return pd.to_datetime(s).to_pydatetime()


def years_between(d1: datetime, d2: datetime) -> float:
    return abs((d2 - d1).days) / 365.25


def approx_lookup(value: float, ref_df: pd.DataFrame, value_col: str, target_col: str) -> Optional[float]:
    """
    Excel VLOOKUP approximate (range_lookup=True) viselkedés:
    - rendez value_col szerint növekvőbe
    - visszaadja a legnagyobb olyan target_col értéket, amelynek value_col <= value
    - ha value a legkisebbnél is kisebb, akkor az első targetet adja
    """
    if value is None or pd.isna(value):
        return None
    df = ref_df[[value_col, target_col]].dropna().copy()
    if df.empty:
        return None
    df = df.sort_values(value_col)
    # merge_asof "backward" = legnagyobb <= value
    matched = pd.merge_asof(
        pd.DataFrame({value_col: [value]}),
        df,
        on=value_col,
        direction="backward",
        allow_exact_matches=True,
    )
    out = matched.iloc[0][target_col]
    if pd.isna(out):
        # ha nincs backward match (érték kisebb mint a legkisebb), vegyük az első targetet
        out = df.iloc[0][target_col]
    return float(out)


def load_reference_table(path: str, is_boy: bool) -> pd.DataFrame:
    """
    Beolvassa az Excel táblát a megfelelő munkalapról.
    - boys lap a fiúkhoz
    - girls lap a lányokhoz
    """
    sheet = "boys" if is_boy else "girls"
    return pd.read_excel(path, sheet_name=sheet)


def mk_correction_factor(mk_minus_ca: float) -> float:
    """
    MK-kor% szabályod szerint:
    MK - CA (év) különbség alapján:
      < -1.99  -> 1.08
      -1.99 .. -0.99 -> 1.05
      -0.99 .. 0.99 -> 1.00
      0.99 .. 1.99 -> 0.95
      > 1.99 -> 0.92
    """
    if mk_minus_ca < -1.99:
        return 1.08
    if -1.99 <= mk_minus_ca < -0.99:
        return 1.05
    if -0.99 <= mk_minus_ca <= 0.99:
        return 1.00
    if 0.99 < mk_minus_ca <= 1.99:
        return 0.95
    return 0.92


def bmi_category(bmi: float) -> str:
    # A megadott küszöbök alapján (kiegészítve az "..."-ot)
    if bmi >= 25:
        return "túlsúlyos"
    if bmi >= 18:
        return "normális testsúly"
    if bmi > 17:
        return "enyhe soványság"
    if bmi > 16:
        return "mérsékelt soványság"
    return "súlyos soványság"


def phv_category(phv: float) -> str:
    if phv is None or (isinstance(phv, float) and math.isnan(phv)):
        return "ismeretlen"
    if phv > 1:
        return "magas"
    if phv < -1.5:
        return "alacsony"
    return "normál"


def bucket(value: float, high=5.5, mid=2.6, high_txt="", mid_txt="", low_txt="") -> str:
    if value >= high:
        return high_txt
    if value >= mid:
        return mid_txt
    return low_txt


# ====== Fő kalkulátor ======
def compute_all_metrics(
    record: Dict[str, Any],
    sex: str,
    ref_path: str = "mk_components.xlsx",
    mkkorrekcio: float = 0.125
) -> CalcResult:
    """
    record kulcsok (a felsorolásod alapján, mind cm-ben / kg-ban, dátumok stringben):
    Név, Nem, Korosztály, Születési dátum, Mérés dátuma,
    TTS, TTM, ÜLŐ, TR, LPR, BR, eCSR, HR, COR, MSR,
    FK, FFK, AKK, CUK, KZK, MKK, COK, ASK, BOK, VAS,
    CRS, MKS, MMG, HUS, TDS, FKF, DKF, BoSZ
    """
    # --- alap mérések ---
    TTS = _to_float(record.get("TTS"))             # testsúly kg
    TTM = _to_float(record.get("TTM"))             # testmagasság cm
    ULO = _to_float(record.get("ÜLŐ"))             # ülőmagasság cm

    TR  = _to_float(record.get("TR"))
    LPR = _to_float(record.get("LPR"))
    BR  = _to_float(record.get("BR"))
    eCSR= _to_float(record.get("eCSR"))
    HR  = _to_float(record.get("HR"))
    COR = _to_float(record.get("COR"))
    MSR = _to_float(record.get("MSR"))

    FFK = _to_float(record.get("FFK"))
    AKK = _to_float(record.get("AKK"))
    KZK = _to_float(record.get("KZK"))
    VAS = _to_float(record.get("VAS"))
    ASK = _to_float(record.get("ASK"))

    HUS = _to_float(record.get("HUS"))
    TDS = _to_float(record.get("TDS"))

    birth = parse_date(record.get("Születési dátum") or record.get("Birth date"))
    meas = parse_date(record.get("Mérés dátuma") or record.get("Measurement date"))
    CA = years_between(birth, meas)  # kronológiai életkor (év)

    raw_sex = sex if sex is not None else record.get("Nem")
    s = str(raw_sex).strip().lower()

    BOY_TOKENS = {"fiú", "fiu", "f", "boy", "b", "m", "male", "1", "ferfi", "férfi"}
    GIRL_TOKENS = {"lány", "lany", "l", "girl", "g", "fem", "female", "0", "2", "no", "nő"}

    if s in BOY_TOKENS:
        is_boy = True
    elif s in GIRL_TOKENS:
        is_boy = False
    else:
        raise ValueError(f"Ismeretlen 'Nem' érték: {raw_sex!r}. Elfogadottak: {BOY_TOKENS | GIRL_TOKENS}")

    # --- referencia tábla ---
    ref = load_reference_table(ref_path, is_boy)

    # --- PLX ---
    plx = AKK + KZK + VAS

    # --- VLOOKUP-szerű „rendelt korok” ---
    # PLX_kor
    plx_age = approx_lookup(plx, ref, REF_COLUMNS["plx"]["value_col"], REF_COLUMNS["plx"]["age_col"])

    # TTS_kor (nemi bontás)
    if is_boy:
        tts_age = approx_lookup(TTS, ref, REF_COLUMNS["tts_boy"]["value_col"], REF_COLUMNS["tts_boy"]["age_col"])
        ttm_age = approx_lookup(TTM, ref, REF_COLUMNS["ttm_boy"]["value_col"], REF_COLUMNS["ttm_boy"]["age_col"])
    else:
        tts_age = approx_lookup(TTS, ref, REF_COLUMNS["tts_girl"]["value_col"], REF_COLUMNS["tts_girl"]["age_col"])
        ttm_age = approx_lookup(TTM, ref, REF_COLUMNS["ttm_girl"]["value_col"], REF_COLUMNS["ttm_girl"]["age_col"])

    # --- MK_nyers ---
    # MK_nyers = (PLX-kor + Testsúly-kor + Testmagasság-kor + Életkor + 3×Mkkorrekció) / 4
    mk_raw = (float(plx_age) + float(tts_age) + float(ttm_age) + CA + 3.0 * mkkorrekcio) / 4.0

    # --- MK-kor% faktor a szabályrendszered szerint ---
    mk_minus_ca = mk_raw - CA
    mk_corr = mk_correction_factor(mk_minus_ca)

    # --- MK végleges ---
    mk = mk_raw * mk_corr

    # --- VTTM ---
    # A megadott képleted értelmezése:
    #   VTTM = TTM * 100 / ((MK% (az életkorhoz) + MK%(+0.25 év)) / 2)
    # Ehhez kell egy MK%-t adó referencia (életkor -> százalék). Ha nincs, a VTTM-et None-ra állítjuk.
    vttm = None
    try:
        # 1) MK%-lookup az MK életkorára
        mkpct_now = approx_lookup(mk, ref, REF_COLUMNS["mkpct"]["value_col"], REF_COLUMNS["mkpct"]["age_col"])
        mkpct_plus = approx_lookup(mk + 0.25, ref, REF_COLUMNS["mkpct"]["value_col"], REF_COLUMNS["mkpct"]["age_col"])
        if mkpct_now is not None and mkpct_plus is not None and mkpct_now > 0 and mkpct_plus > 0:
            denom = (mkpct_now + mkpct_plus) / 2.0
            vttm = TTM * 100.0 / denom
    except KeyError:
        # Ha nincs MK%-oszlop a táblában, VTTM nem számítható ezzel a módszerrel
        vttm = None

    # --- 6 bőrredő összege ---
    sum6 = TR + LPR + BR + eCSR + HR + COR + MSR  # megjegyzésedben "LPRB" is szerepelt, itt LPR-rel számolok

    # --- Testzsír% (Deurenberg-szerű lineáris becslés a megadott együtthatókkal) ---
    # fiú: zsir1 * sum6 + zsir2 ; lány: zsir3 * sum6 + zsir4
    zsir1, zsir2 = 0.1051, 2.585
    zsir3, zsir4 = 0.1548, 3.58
    bodyfat = zsir1 * sum6 + zsir2 if is_boy else zsir3 * sum6 + zsir4

    # --- Endomorfia ---
    # S = (TR+LPR+eCSR) * (170.18 / TTM)
    S = (TR + LPR + eCSR) * (170.18 / TTM)
    endo = -0.7182 + 0.1451 * S - 0.00068 * (S ** 2) + 0.0000014 * (S ** 3)
    endo_cat = bucket(
        endo,
        high=5.5, mid=2.6,
        high_txt="hízásra hajlamos testalkat",
        mid_txt="hízásra közepes mértékben hajlamos testalkat",
        low_txt="hízásra nem hajlamos testalkat",
    )

    # --- Mezomorfia ---
    mezo = (0.858 * HUS + 0.601 * TDS + (0.188 * (FFK - (TR / 10.0))) + (0.161 * (ASK - (MSR / 10.0)))) - (0.131 * TTM) + 4.5
    mezo_cat = bucket(
        mezo,
        high=5.5, mid=2.6,
        high_txt="nagy mértékben fejleszthető izomzat",
        mid_txt="közepes mértékben fejleszthető izomzat",
        low_txt="kis mértékben fejleszthető izomzat",
    )

    # --- Ektomorfia ---
    hwr = TTM / (TTS ** (1.0 / 3.0))
    if hwr > 40.75:
        ekto = 0.732 * hwr - 28.58
    elif hwr > 38.28:
        ekto = 0.463 * hwr - 17.63
    else:
        ekto = 0.1
    ekto_cat = bucket(
        ekto,
        high=5.5, mid=2.6,
        high_txt="kifejezetten nyúlánk alkat",
        mid_txt="közepesen nyúlánk alkat",
        low_txt="alacsony fokú relatív nyúlánkság",
    )

    # --- BMI és kategória ---
    bmi = TTS / ((TTM / 100.0) ** 2)  # kg / m^2
    bmi_cat = bmi_category(bmi)

    # --- PHV (Mirwald-típusú képletek mintájára – a megadott együtthatókat használva) ---
    if is_boy:
        phv = -9.236 + (0.0002708 * ((TTM - ULO) * ULO)) - (0.001663 * CA * (TTM - ULO)) + (0.007216 * CA * ULO) + 0.02292 * (ULO / TTM)
    else:
        phv = -9.376 + (0.0001882 * ((TTM - ULO) * ULO)) + (0.0022 * CA * (TTM - ULO)) + (0.005841 * CA * ULO) - 0.002658 * (CA * ULO) + 0.07693 * (ULO / TTM)
    phv_cat = phv_category(phv)

    return CalcResult(
        plx=plx,
        age_years=CA,
        mk_raw=mk_raw,
        mk_corr_factor=mk_corr,
        mk=mk,
        vttm=vttm,
        endomorphy=endo,
        endomorphy_cat=endo_cat,
        mesomorphy=mezo,
        mesomorphy_cat=mezo_cat,
        ectomorphy=ekto,
        ectomorphy_cat=ekto_cat,
        sum6=sum6,
        bodyfat_percent=bodyfat,
        bmi=bmi,
        bmi_cat=bmi_cat,
        phv=phv,
        phv_cat=phv_cat,
    )


# ====== Példa futtatás ======
if __name__ == "__main__":
    sample = {
        "Név": "Minta János",
        "Nem": "fiú",
        "Születési dátum": "2010-05-12",
        "Mérés dátuma": "2025-05-20",
        # kg / cm
        "TTS": 55.0, "TTM": 170.0, "ÜLŐ": 88.0,
        # redők (mm) és körméretek / szélességek (cm) – a saját mértékegységed szerint add meg
        "TR": 8.0, "LPR": 9.0, "BR": 10.0, "eCSR": 7.0, "HR": 8.0, "COR": 6.0, "MSR": 7.0,
        "FFK": 32.0, "AKK": 26.0, "KZK": 27.0, "VAS": 38.0, "ASK": 40.0,
        "HUS": 14.0, "TDS": 10.0,
    }
    # FIGYELEM: a mk_components.csv-nek egyeznie kell a REF_COLUMNS fejléceivel!
    res = compute_all_metrics(sample, sex=sample["Nem"], ref_path="mk_components.xlsx")
    print(res)
