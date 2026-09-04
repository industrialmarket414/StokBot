"""Bayi eslesmeyenler raporlarindan 3 sade urun yonetim Excel'i olusturur."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import re
import unicodedata

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
VARSAYILAN_CIKTI_DIR = BASE_DIR / "Cikti"
FIRMALAR = [
    "KIRAC", "KONYA", "SANTRAL", "MET", "KLTSHOP", "MEANWELL", "YILDIZ",
    "VEKMAR",
]
MANUEL_KOLONLAR = [
    "Nihai Karar",
    "Ana Tedarikci",
    "Kontrol Notu",
    "Kontrol Tarihi",
    "T-Soft'a Eklendi",
    "T-Soft Web Servis Kodu",
]
AYLIK_KIMLIK_KOLONLARI = ["Marka", "Web Servis Kodu"]
AYLIK_TOPLAM_KOLONU = "Tum Cekimlerin Toplami"
TURKCE_AYLAR = {
    1: "Ocak", 2: "Subat", 3: "Mart", 4: "Nisan", 5: "Mayis", 6: "Haziran",
    7: "Temmuz", 8: "Agustos", 9: "Eylul", 10: "Ekim", 11: "Kasim", 12: "Aralik",
}
MARKA_GORUNUM_ESLESMELERI = {
    "ORTAC": "ORTAÇ",
    "OZNUR": "ÖZNUR",
    "OZNUR KABLO": "ÖZNUR KABLO",
    "WEIDMULLER": "WEIDMÜLLER",
    "SAFAK": "ŞAFAK ELEKTRİK",
    "SAFAK ELEKTRIK": "ŞAFAK ELEKTRİK",
    "KARDES": "KARDEŞ ELEKTRİK",
    "KARDES ELEKTRIK": "KARDEŞ ELEKTRİK",
    "KRAUS NAIMER": "KRAUS & NAIMER",
    "KRAUS & NAIMER": "KRAUS & NAIMER",
    "SCHNEIDER": "SCHNEIDER ELECTRIC",
    "VIKO-PANASONIC": "VIKO",
    "VIKO PANASONIC": "VIKO",
}


def metin_norm(deger) -> str:
    if pd.isna(deger):
        return ""
    metin = str(deger).replace("\xa0", " ").strip().upper()
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(ch for ch in metin if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", metin)


def marka_gorunum_standart(deger) -> str:
    norm = metin_norm(deger)
    return MARKA_GORUNUM_ESLESMELERI.get(norm, norm)


def baslik_norm(deger) -> str:
    return re.sub(r"[^A-Z0-9]", "", metin_norm(deger))


def kolon_bul(df: pd.DataFrame, adaylar: list[str]) -> str | None:
    mevcut = {baslik_norm(kolon): kolon for kolon in df.columns}
    for aday in adaylar:
        if baslik_norm(aday) in mevcut:
            return mevcut[baslik_norm(aday)]
    return None


def stok_sayisina_cevir(deger) -> int | None:
    if pd.isna(deger):
        return None
    metin = metin_norm(deger)
    if not metin:
        return None
    if "YOK" in metin:
        return 0
    metin = re.sub(r"\s*(ADET|ADT|AD|TANE|SET|PCS|PC|METRE|MTR|MT|M)\.?\s*$", "", metin)
    metin = re.sub(r"\s+", "", metin)
    if "." in metin and "," in metin:
        metin = metin.replace(".", "").replace(",", ".") if metin.rfind(",") > metin.rfind(".") else metin.replace(",", "")
    elif re.fullmatch(r"-?\d{1,3}(?:[.,]\d{3})+", metin):
        metin = metin.replace(".", "").replace(",", "")
    else:
        metin = metin.replace(",", ".")
    try:
        return max(0, int(float(metin)))
    except ValueError:
        eslesme = re.search(r"-?\d+(?:[.,]\d+)*", metin)
        return stok_sayisina_cevir(eslesme.group(0)) if eslesme else None


def kaynak_dosyalari_bul(cikti_dir: Path) -> list[tuple[str, Path]]:
    """8 bayi icin eslesmeyen, stoklu aday ve stoksuz izleme dosyalarini bulur."""
    bulunan: list[tuple[str, Path]] = []
    for firma in FIRMALAR:
        firma_dir = cikti_dir / firma
        firma_kucuk = firma.lower()
        beklenenler = [
            firma_dir / f"{firma_kucuk}_eslesmeyenler.xlsx",
            firma_dir / f"{firma_kucuk}_tsoftta_yok_bayide_var_siteye_yuklenecekler.xlsx",
            firma_dir / f"{firma_kucuk}_tsoftta_yok_bayide_var_stok_yok.xlsx",
        ]
        for dosya in beklenenler:
            if dosya.exists() and not dosya.name.startswith("~$"):
                bulunan.append((firma, dosya))
    return bulunan


def eslesen_kaynak_dosyalari_bul(cikti_dir: Path) -> list[tuple[str, Path]]:
    """B2B ve T-Soft'ta eslesen firma bazli stok guncelleme dosyalarini bulur."""
    bulunan: list[tuple[str, Path]] = []
    for firma in FIRMALAR:
        dosya = cikti_dir / firma / f"{firma.lower()}_stok_guncellenecekler.xlsx"
        if dosya.exists() and not dosya.name.startswith("~$"):
            bulunan.append((firma, dosya))
    return bulunan


def tsoftta_yok_satiri_mi(neden) -> bool:
    metin = metin_norm(neden)
    if not metin:
        return True
    return (
        "B2B'DE VAR, SITEDE YOK" in metin
        or "T-SOFT URUN LISTESINDE ESLESMEDI" in metin
        or "T-SOFT'TA YOK" in metin
    )


def kaynaklari_oku(dosyalar: list[tuple[str, Path]]) -> tuple[pd.DataFrame, list[str]]:
    kayitlar: list[dict] = []
    sorunlar: list[str] = []

    for firma, dosya in dosyalar:
        try:
            df = pd.read_excel(dosya, dtype=str).fillna("")
        except Exception as exc:
            sorunlar.append(f"{firma}: dosya okunamadi ({exc})")
            continue

        marka_kolonu = kolon_bul(df, ["B2B Marka", "Marka"])
        kod_kolonu = kolon_bul(df, [
            "B2B Web Servis Kodu", "Web Servis Kodu", "Urun Kodu",
            "Stok Kodu", "Model No", "Model",
        ])
        stok_kolonu = kolon_bul(df, ["Yeni Stok", "B2B Stok Metni", "Stok", "Kalan Stok", "Stok Adeti"])
        neden_kolonu = kolon_bul(df, ["Neden", "Eslesme Durumu"])

        if kod_kolonu is None or stok_kolonu is None:
            sorunlar.append(f"{firma}: urun kodu veya stok kolonu bulunamadi")
            continue

        for sira, row in df.iterrows():
            neden = row.get(neden_kolonu, "") if neden_kolonu else ""
            if neden_kolonu and not tsoftta_yok_satiri_mi(neden):
                continue

            kod = str(row.get(kod_kolonu, "")).strip()
            if not kod:
                continue
            stok = stok_sayisina_cevir(row.get(stok_kolonu, ""))
            if stok is None:
                sorunlar.append(f"{firma}: {sira + 2}. satirin stoku okunamadi")
                continue
            marka = str(row.get(marka_kolonu, "BILINMIYOR")).strip() if marka_kolonu else "BILINMIYOR"
            kayitlar.append({
                "Marka": marka or "BILINMIYOR",
                "Urun / Model Kodu": kod,
                "Firma": firma,
                "Stok": stok,
            })

    return pd.DataFrame(kayitlar), sorunlar


def birlesik_tsoftta_yok_urunleri_oku(dosya: Path) -> pd.DataFrame:
    """stok_eslestir.py tarafindan uretilen ortak TSoft'ta yok listesini okur."""
    if not dosya.exists():
        return pd.DataFrame()

    gerekli = {
        "Marka",
        "Urun / Model Kodu",
        *FIRMALAR,
        "Toplam Stok",
        "Firma Sayisi",
    }
    try:
        df = pd.read_excel(dosya).fillna("")
    except Exception:
        return pd.DataFrame()
    if not gerekli.issubset(df.columns):
        return pd.DataFrame()

    sonuc = df.copy()
    sonuc["Marka"] = sonuc["Marka"].apply(marka_gorunum_standart)
    sonuc["Urun / Model Kodu"] = sonuc["Urun / Model Kodu"].astype(str).str.strip()
    sonuc = sonuc[sonuc["Urun / Model Kodu"].ne("")].copy()
    for kolon in [*FIRMALAR, "Toplam Stok", "Firma Sayisi"]:
        sonuc[kolon] = pd.to_numeric(sonuc[kolon], errors="coerce").fillna(0).astype(int)
    sonuc["Karar"] = "INCELENECEK"
    sonuc.loc[sonuc["Toplam Stok"] <= 0, "Karar"] = "STOK BEKLENECEK"
    sonuc.loc[
        (sonuc["Toplam Stok"] >= 10) & (sonuc["Firma Sayisi"] >= 2),
        "Karar",
    ] = "EKLENECEK"
    sonuc["Marka Anahtar"] = sonuc["Marka"].apply(metin_norm)
    sonuc["Kod Anahtar"] = sonuc["Urun / Model Kodu"].apply(metin_norm)
    return sonuc


def eslesen_kaynaklari_oku(dosyalar: list[tuple[str, Path]]) -> tuple[pd.DataFrame, list[str]]:
    kayitlar: list[dict] = []
    sorunlar: list[str] = []
    for firma, dosya in dosyalar:
        try:
            df = pd.read_excel(dosya, dtype=str).fillna("")
        except Exception as exc:
            sorunlar.append(f"{firma}: eslesen stok dosyasi okunamadi ({exc})")
            continue

        gerekenler = {"B2B Marka", "Web Servis Kodu", "Yeni Stok"}
        if not gerekenler.issubset(df.columns):
            eksik = ", ".join(sorted(gerekenler - set(df.columns)))
            sorunlar.append(f"{firma}: eslesen stok dosyasinda eksik kolon ({eksik})")
            continue

        for _, row in df.iterrows():
            kod = str(row.get("Web Servis Kodu", "")).strip()
            if not kod:
                continue
            stok = stok_sayisina_cevir(row.get("Yeni Stok", ""))
            if stok is None:
                stok = stok_sayisina_cevir(row.get("B2B Stok Metni", ""))
            if stok is None:
                stok = 0
            marka = marka_gorunum_standart(row.get("B2B Marka", "BILINMIYOR"))
            kayitlar.append({
                "Marka": marka or "BILINMIYOR",
                "Web Servis Kodu": kod,
                "Firma": firma,
                "Stok": stok,
            })
    return pd.DataFrame(kayitlar), sorunlar


def eslesen_marka_ozeti_hazirla(kayitlar: pd.DataFrame) -> pd.DataFrame:
    kolonlar = [
        "Marka", *FIRMALAR, "Toplam Stok", "Toplam Urun",
        "Stoklu Urun", "Stoksuz Urun", "Firma Sayisi",
    ]
    if kayitlar.empty:
        return pd.DataFrame(columns=kolonlar)

    firma_stok = (
        kayitlar.groupby(["Marka", "Web Servis Kodu", "Firma"], as_index=False)["Stok"]
        .max()
    )
    urunler = firma_stok.pivot_table(
        index=["Marka", "Web Servis Kodu"],
        columns="Firma",
        values="Stok",
        aggfunc="max",
        fill_value=0,
    ).reset_index()
    for firma in FIRMALAR:
        if firma not in urunler.columns:
            urunler[firma] = 0
        urunler[firma] = pd.to_numeric(urunler[firma], errors="coerce").fillna(0).astype(int)
    urunler["Toplam Stok"] = urunler[FIRMALAR].sum(axis=1).astype(int)

    ozet = urunler.groupby("Marka", as_index=False).agg({
        **{firma: "sum" for firma in FIRMALAR},
        "Toplam Stok": "sum",
        "Web Servis Kodu": "count",
    }).rename(columns={"Web Servis Kodu": "Toplam Urun"})
    stoklu = urunler[urunler["Toplam Stok"] > 0].groupby("Marka").size()
    stoksuz = urunler[urunler["Toplam Stok"] == 0].groupby("Marka").size()
    ozet["Stoklu Urun"] = ozet["Marka"].map(stoklu).fillna(0).astype(int)
    ozet["Stoksuz Urun"] = ozet["Marka"].map(stoksuz).fillna(0).astype(int)
    ozet["Firma Sayisi"] = (ozet[FIRMALAR] > 0).sum(axis=1).astype(int)
    return ozet[kolonlar].sort_values("Marka").reset_index(drop=True)


def urunleri_birlestir(kayitlar: pd.DataFrame) -> pd.DataFrame:
    if kayitlar.empty:
        return pd.DataFrame()

    df = kayitlar.copy()
    df["Marka Anahtar"] = df["Marka"].apply(metin_norm)
    df["Kod Anahtar"] = df["Urun / Model Kodu"].apply(metin_norm)

    # Ayni bayi + ayni urun tekrarlanirsa toplama yapma; en yuksek stogu kullan.
    firma_stok = (
        df.groupby(["Marka Anahtar", "Kod Anahtar", "Firma"], as_index=False)["Stok"]
        .max()
    )
    gorunum = (
        df.groupby(["Marka Anahtar", "Kod Anahtar"], as_index=False)
        .agg({"Marka": "first", "Urun / Model Kodu": "first"})
    )
    pivot = firma_stok.pivot_table(
        index=["Marka Anahtar", "Kod Anahtar"],
        columns="Firma",
        values="Stok",
        aggfunc="max",
        fill_value=0,
    ).reset_index()
    sonuc = gorunum.merge(pivot, on=["Marka Anahtar", "Kod Anahtar"], how="left")

    for firma in FIRMALAR:
        if firma not in sonuc.columns:
            sonuc[firma] = 0
        sonuc[firma] = pd.to_numeric(sonuc[firma], errors="coerce").fillna(0).astype(int)

    sonuc["Toplam Stok"] = sonuc[FIRMALAR].sum(axis=1).astype(int)
    sonuc["Firma Sayisi"] = (sonuc[FIRMALAR] > 0).sum(axis=1).astype(int)
    sonuc["Karar"] = sonuc.apply(
        lambda row: "STOK BEKLENECEK" if row["Toplam Stok"] == 0 else (
            "EKLENECEK" if row["Toplam Stok"] >= 10 and row["Firma Sayisi"] >= 2 else "INCELENECEK"
        ),
        axis=1,
    )
    return sonuc


def eski_manuel_bilgileri_oku(rapor: Path) -> dict[tuple[str, str], dict]:
    eski: dict[tuple[str, str], dict] = {}
    if not rapor.exists():
        return eski
    for sayfa in [
        "TSoft Yok - Stoklu Ürünler",
        "TSoft Yok - Stoksuz Ürünler",
        "Urun Adaylari",
        "Stoksuz Izleme",
    ]:
        try:
            df = pd.read_excel(rapor, sheet_name=sayfa, header=2, dtype=str).fillna("")
        except Exception:
            continue
        if "Marka" not in df.columns or "Urun / Model Kodu" not in df.columns:
            continue
        for _, row in df.iterrows():
            anahtar = (metin_norm(row["Marka"]), metin_norm(row["Urun / Model Kodu"]))
            eski[anahtar] = {kolon: row.get(kolon, "") for kolon in MANUEL_KOLONLAR}
    return eski


def gun_basligi(gun: date) -> str:
    return f"{gun.day:02d} {TURKCE_AYLAR[gun.month]}"


def eski_aylik_stok_takibini_oku(rapor: Path) -> pd.DataFrame:
    temel = [*AYLIK_KIMLIK_KOLONLARI, AYLIK_TOPLAM_KOLONU]
    if not rapor.exists():
        return pd.DataFrame(columns=temel)
    eski = None
    for sayfa in ["31 Günlük Stok Takibi", "Aylik Stok Takibi"]:
        try:
            eski = pd.read_excel(rapor, sheet_name=sayfa, header=2, dtype=str)
            break
        except Exception:
            continue
    if eski is None:
        return pd.DataFrame(columns=temel)

    eski.columns = [str(kolon).strip() for kolon in eski.columns]
    if not set(AYLIK_KIMLIK_KOLONLARI).issubset(eski.columns):
        return pd.DataFrame(columns=temel)
    if AYLIK_TOPLAM_KOLONU not in eski.columns:
        eski[AYLIK_TOPLAM_KOLONU] = 0
    eski["Marka"] = eski["Marka"].fillna("").astype(str).str.strip()
    eski["Web Servis Kodu"] = eski["Web Servis Kodu"].fillna("").astype(str).str.strip()
    return eski[eski["Web Servis Kodu"] != ""].copy()


def aylik_stok_takibini_hazirla(
    rapor: Path,
    gunluk_stok_dosyasi: Path,
    bugun: date | None = None,
) -> pd.DataFrame:
    """Tum bayiler guncelleme listesinden son 31 gunluk urun bazli stok gecmisi olusturur."""
    bugun = bugun or date.today()
    bugun_kolonu = gun_basligi(bugun)
    eski = eski_aylik_stok_takibini_oku(rapor)
    sabit_kolonlar = [*AYLIK_KIMLIK_KOLONLARI, AYLIK_TOPLAM_KOLONU]
    tarih_kolonlari = [kolon for kolon in eski.columns if kolon not in sabit_kolonlar]

    if not gunluk_stok_dosyasi.exists():
        print(f"Aylik stok takibi kaynagi bulunamadi: {gunluk_stok_dosyasi}")
        return eski

    gunluk = pd.read_excel(gunluk_stok_dosyasi, dtype=str).fillna("")
    gerekenler = {"B2B Marka", "Web Servis Kodu", "Guncellenecek Stok"}
    if not gerekenler.issubset(gunluk.columns):
        eksik = ", ".join(sorted(gerekenler - set(gunluk.columns)))
        print(f"Aylik stok takibi kaynaginda eksik kolon: {eksik}")
        return eski

    if "Guncelleme Durumu" in gunluk.columns:
        gunluk = gunluk[
            gunluk["Guncelleme Durumu"].apply(metin_norm) == "GONDERILECEK"
        ].copy()

    gunluk["Web Servis Kodu"] = gunluk["Web Servis Kodu"].astype(str).str.strip()
    gunluk["Marka"] = gunluk["B2B Marka"].astype(str).str.strip()
    gunluk[bugun_kolonu] = pd.to_numeric(
        gunluk["Guncellenecek Stok"], errors="coerce"
    ).fillna(0).astype(int)
    gunluk = gunluk[gunluk["Web Servis Kodu"] != ""].copy()
    gunluk = (
        gunluk.groupby("Web Servis Kodu", as_index=False)
        .agg({"Marka": "first", bugun_kolonu: "max"})
    )

    if bugun_kolonu not in tarih_kolonlari:
        tarih_kolonlari.append(bugun_kolonu)
    tarih_kolonlari = tarih_kolonlari[-31:]

    if eski.empty:
        takip = gunluk.set_index("Web Servis Kodu")
    else:
        takip = eski.drop(columns=[AYLIK_TOPLAM_KOLONU], errors="ignore")
        takip = takip.drop_duplicates("Web Servis Kodu", keep="last").set_index("Web Servis Kodu")
        if bugun_kolonu not in takip.columns:
            takip[bugun_kolonu] = pd.NA
        else:
            takip[bugun_kolonu] = pd.NA
        guncel = gunluk.set_index("Web Servis Kodu")
        yeni_kodlar = guncel.index.difference(takip.index)
        if len(yeni_kodlar):
            takip = pd.concat([takip, guncel.loc[yeni_kodlar]], axis=0)
        takip.loc[guncel.index, "Marka"] = guncel["Marka"]
        takip.loc[guncel.index, bugun_kolonu] = guncel[bugun_kolonu]

    for kolon in tarih_kolonlari:
        if kolon not in takip.columns:
            takip[kolon] = pd.NA
        takip[kolon] = pd.to_numeric(takip[kolon], errors="coerce")

    takip = takip[["Marka", *tarih_kolonlari]]
    takip = takip.dropna(subset=tarih_kolonlari, how="all")
    takip[AYLIK_TOPLAM_KOLONU] = (
        takip[tarih_kolonlari].sum(axis=1, min_count=1).fillna(0).astype(int)
    )
    takip = takip.reset_index()
    takip = takip[["Marka", "Web Servis Kodu", AYLIK_TOPLAM_KOLONU, *tarih_kolonlari]]
    return takip.sort_values(["Marka", "Web Servis Kodu"]).reset_index(drop=True)


def stok_guncelleme_sayilari_oku(gunluk_stok_dosyasi: Path) -> dict[str, int]:
    sayilar = {
        "toplam": 0,
        "gonderilecek": 0,
        "gonderilmeyecek": 0,
        "stoklu": 0,
        "stoksuz": 0,
        "toplam_stok": 0,
    }
    if not gunluk_stok_dosyasi.exists():
        return sayilar
    try:
        df = pd.read_excel(
            gunluk_stok_dosyasi,
            usecols=["Web Servis Kodu", "Guncelleme Durumu", "Bayi Toplam Stok"],
        ).fillna("")
    except Exception:
        return sayilar
    df["Web Servis Kodu"] = df["Web Servis Kodu"].astype(str).str.strip()
    df = df[df["Web Servis Kodu"] != ""].drop_duplicates("Web Servis Kodu", keep="last")
    durum = df["Guncelleme Durumu"].apply(metin_norm)
    sayilar["toplam"] = len(df)
    sayilar["gonderilecek"] = int((durum == "GONDERILECEK").sum())
    sayilar["gonderilmeyecek"] = sayilar["toplam"] - sayilar["gonderilecek"]
    stok = pd.to_numeric(df["Bayi Toplam Stok"], errors="coerce").fillna(0)
    sayilar["stoklu"] = int((stok > 0).sum())
    sayilar["stoksuz"] = int((stok <= 0).sum())
    sayilar["toplam_stok"] = int(stok.sum())
    return sayilar


def uc_tabloyu_hazirla(urunler: pd.DataFrame, eski: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    temel_kolonlar = ["Marka", "Urun / Model Kodu", *FIRMALAR, "Toplam Stok", "Firma Sayisi", "Karar"]
    if urunler.empty:
        bos = pd.DataFrame(columns=[*temel_kolonlar, *MANUEL_KOLONLAR])
        marka_bos = pd.DataFrame(columns=["Marka", *FIRMALAR, "Toplam Stok", "Urun Adayi", "Stoksuz Urun", "Firma Sayisi"])
        return bos.copy(), bos.copy(), marka_bos

    tablo = urunler.copy()
    for kolon in MANUEL_KOLONLAR:
        tablo[kolon] = ""
    for idx, row in tablo.iterrows():
        onceki = eski.get((row["Marka Anahtar"], row["Kod Anahtar"]), {})
        for kolon in MANUEL_KOLONLAR:
            tablo.at[idx, kolon] = onceki.get(kolon, "")

    tablo = tablo[[*temel_kolonlar, *MANUEL_KOLONLAR]]
    adaylar = tablo[tablo["Toplam Stok"] > 0].sort_values(
        ["Marka", "Toplam Stok"], ascending=[True, False]
    ).reset_index(drop=True)
    stoksuz = tablo[tablo["Toplam Stok"] == 0].sort_values(
        ["Marka", "Urun / Model Kodu"]
    ).reset_index(drop=True)

    marka = tablo.groupby("Marka", as_index=False).agg({
        **{firma: "sum" for firma in FIRMALAR},
        "Toplam Stok": "sum",
        "Urun / Model Kodu": "count",
    }).rename(columns={"Urun / Model Kodu": "Toplam Urun"})
    marka["Urun Adayi"] = marka["Marka"].map(adaylar.groupby("Marka").size()).fillna(0).astype(int)
    marka["Stoksuz Urun"] = marka["Marka"].map(stoksuz.groupby("Marka").size()).fillna(0).astype(int)
    marka["Firma Sayisi"] = (marka[FIRMALAR] > 0).sum(axis=1).astype(int)
    marka = marka[["Marka", *FIRMALAR, "Toplam Stok", "Toplam Urun", "Urun Adayi", "Stoksuz Urun", "Firma Sayisi"]]
    return adaylar, stoksuz, marka.sort_values("Marka").reset_index(drop=True)


def sade_excel_yaz(df: pd.DataFrame, dosya: Path, sayfa_adi: str):
    dosya.parent.mkdir(parents=True, exist_ok=True)
    gecici = dosya.with_name(f"~{dosya.stem}_gecici.xlsx")
    df.to_excel(gecici, index=False, sheet_name=sayfa_adi)

    # Renk yok; sadece okunabilirlik icin filtre, sabit baslik ve kolon genisligi.
    wb = load_workbook(gecici)
    ws = wb[sayfa_adi]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for kolon_hucreleri in ws.columns:
        en_uzun = max(len(str(hucre.value or "")) for hucre in kolon_hucreleri[:300])
        ws.column_dimensions[kolon_hucreleri[0].column_letter].width = min(max(en_uzun + 2, 10), 35)
    wb.save(gecici)
    gecici.replace(dosya)


def tek_yonetim_raporu_yaz(
    adaylar: pd.DataFrame,
    stoksuz: pd.DataFrame,
    marka: pd.DataFrame,
    eslesen_marka: pd.DataFrame,
    aylik_stok: pd.DataFrame,
    guncelleme_sayilari: dict[str, int],
    rapor: Path,
):
    """Eski duzendeki tek, renkli ve cok sayfali yonetim raporunu yazar."""
    rapor.parent.mkdir(parents=True, exist_ok=True)
    gecici = rapor.with_name(f"~{rapor.stem}_gecici.xlsx")
    eklenecek_sayisi = (
        int((adaylar["Karar"] == "EKLENECEK").sum())
        if not adaylar.empty and "Karar" in adaylar.columns else 0
    )
    incelenecek_sayisi = (
        int((adaylar["Karar"] == "INCELENECEK").sum())
        if not adaylar.empty and "Karar" in adaylar.columns else 0
    )
    eslesen_toplam_urun = int(guncelleme_sayilari.get("toplam", 0))
    eslesen_stoklu_urun = int(guncelleme_sayilari.get("stoklu", 0))
    eslesen_stoksuz_urun = int(guncelleme_sayilari.get("stoksuz", 0))
    ozet = pd.DataFrame([
        ["TSoft'ta Yok", "B2B'de var toplam ürün", len(adaylar) + len(stoksuz),
         "Yedi bayide tekilleştirilmiş toplam ürün sayısı"],
        ["TSoft'ta Yok", "Stoklu toplam ürün", len(adaylar),
         "Eklenecek ve incelenecek stoklu ürünlerin toplamı"],
        ["TSoft'ta Yok", "Eklenecek ürün", eklenecek_sayisi,
         "Toplam stok >= 10 ve firma sayısı >= 2"],
        ["TSoft'ta Yok", "İncelenecek ürün", incelenecek_sayisi,
         "Stoklu; ancak ekleme eşiklerinden en az biri yetersiz"],
        ["TSoft'ta Yok", "Stoksuz izlenen ürün", len(stoksuz),
         "Bütün bayilerde toplam stok = 0"],
        ["TSoft'ta Yok", "Marka sayısı", len(marka),
         "TSoft'ta eşleşmeyen ürünlerin standart marka sayısı"],
        ["TSoft'ta Var", "Eşleşen toplam ürün", eslesen_toplam_urun,
         "B2B ve TSoft'ta Web Servis Kodu ile eşleşen ürünler"],
        ["TSoft'ta Var", "Stok güncellemesine gönderilecek ürün",
         int(guncelleme_sayilari.get("gonderilecek", 0)),
         "Güncelleme Durumu = GONDERILECEK"],
        ["TSoft'ta Var", "Güncellemeye alınmayacak ürün",
         int(guncelleme_sayilari.get("gonderilmeyecek", 0)),
         "Güvenlik veya marka kuralı nedeniyle gönderilmeyecek"],
        ["TSoft'ta Var", "Stoklu ürün", eslesen_stoklu_urun,
         "Eşleşen ve toplam bayi stoku sıfırdan büyük ürünler"],
        ["TSoft'ta Var", "Stoksuz ürün", eslesen_stoksuz_urun,
         "Eşleşen fakat bütün bayilerde toplam stoku 0 olan ürünler"],
        ["TSoft'ta Var", "Marka sayısı", len(eslesen_marka),
         "B2B ve TSoft'ta eşleşen standart marka sayısı"],
    ], columns=["Kapsam", "Gösterge", "Değer", "Açıklama / Karar Kuralı"])

    with pd.ExcelWriter(gecici, engine="openpyxl") as writer:
        ozet.to_excel(writer, sheet_name="Yönetim Özeti", index=False, startrow=2)
        adaylar.to_excel(writer, sheet_name="TSoft Yok - Stoklu Ürünler", index=False, startrow=2)
        stoksuz.to_excel(writer, sheet_name="TSoft Yok - Stoksuz Ürünler", index=False, startrow=2)
        marka.to_excel(writer, sheet_name="TSoft Yok - Marka Özeti", index=False, startrow=2)
        eslesen_marka.to_excel(writer, sheet_name="TSoft Var - Marka Özeti", index=False, startrow=2)
        aylik_stok.to_excel(writer, sheet_name="31 Günlük Stok Takibi", index=False, startrow=2)

    wb = load_workbook(gecici)
    basliklar = {
        "Yönetim Özeti": "B2B -> T-Soft Ürün Yönetim Raporu",
        "TSoft Yok - Stoklu Ürünler": "B2B'de Var, TSoft'ta Yok - Stoklu Ürünler",
        "TSoft Yok - Stoksuz Ürünler": "B2B'de Var, TSoft'ta Yok - Stoksuz Ürünler",
        "TSoft Yok - Marka Özeti": "TSoft'ta Olmayan Ürünler - Marka ve Firma Bazlı Stok Özeti",
        "TSoft Var - Marka Özeti": "B2B ve TSoft'ta Eşleşen Ürünler - Marka ve Firma Bazlı Stok Özeti",
        "31 Günlük Stok Takibi": "Tüm Bayiler - Son 31 Günlük Ürün Bazlı Stok Takibi",
    }
    lacivert = PatternFill("solid", fgColor="16324F")
    yesil = PatternFill("solid", fgColor="0F766E")
    durum_renkleri = {
        "EKLENECEK": PatternFill("solid", fgColor="DCFCE7"),
        "INCELENECEK": PatternFill("solid", fgColor="FEF3C7"),
        "STOK BEKLENECEK": PatternFill("solid", fgColor="E5E7EB"),
    }

    for sayfa_adi, ws in [(ad, wb[ad]) for ad in basliklar]:
        ws.sheet_view.showGridLines = False
        son_kolon = max(1, ws.max_column)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=son_kolon)
        ws.cell(1, 1, basliklar[sayfa_adi])
        ws.cell(1, 1).fill = lacivert
        ws.cell(1, 1).font = Font(color="FFFFFF", bold=True, size=15)
        ws.cell(1, 1).alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 26
        for cell in ws[3]:
            if cell.value is not None:
                cell.fill = yesil
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if sayfa_adi == "Yönetim Özeti":
            ws.freeze_panes = "A4"
        elif sayfa_adi == "31 Günlük Stok Takibi":
            ws.freeze_panes = "D4"
        else:
            ws.freeze_panes = "C4"
        ws.auto_filter.ref = f"A3:{ws.cell(ws.max_row, ws.max_column).coordinate}"
        for kolon_no in range(1, ws.max_column + 1):
            hucreler = [ws.cell(satir, kolon_no) for satir in range(1, min(ws.max_row, 300) + 1)]
            en_uzun = max(len(str(hucre.value or "")) for hucre in hucreler[:300])
            ws.column_dimensions[get_column_letter(kolon_no)].width = min(max(en_uzun + 2, 11), 35)

        if sayfa_adi in ["TSoft Yok - Stoklu Ürünler", "TSoft Yok - Stoksuz Ürünler"]:
            kolonlar = {ws.cell(3, col).value: col for col in range(1, ws.max_column + 1)}
            karar_col = kolonlar.get("Karar")
            if karar_col:
                for satir in range(4, ws.max_row + 1):
                    deger = str(ws.cell(satir, karar_col).value or "")
                    if deger in durum_renkleri:
                        ws.cell(satir, karar_col).fill = durum_renkleri[deger]

        if sayfa_adi == "31 Günlük Stok Takibi":
            ws.column_dimensions["A"].width = 24
            ws.column_dimensions["B"].width = 22
            ws.column_dimensions["C"].width = 24
            for kolon_no in range(3, ws.max_column + 1):
                for satir in range(4, ws.max_row + 1):
                    ws.cell(satir, kolon_no).number_format = "#,##0"
            for kolon_no in range(4, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(kolon_no)].width = 14

    # Yonetim ozeti iki kapsami renklerle ayirarak okunabilir tutar.
    ws = wb["Yönetim Özeti"]
    tsoft_yok_dolgu = PatternFill("solid", fgColor="FFF2CC")
    tsoft_var_dolgu = PatternFill("solid", fgColor="E2F0D9")
    for satir in range(4, ws.max_row + 1):
        kapsam = str(ws.cell(satir, 1).value or "")
        dolgu = tsoft_yok_dolgu if kapsam == "TSoft'ta Yok" else tsoft_var_dolgu
        for kolon in range(1, 5):
            ws.cell(satir, kolon).fill = dolgu
        ws.cell(satir, 1).font = Font(bold=True)
        ws.cell(satir, 2).font = Font(bold=True)
        ws.cell(satir, 3).number_format = "#,##0"
        ws.cell(satir, 4).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 68

    wb.save(gecici)
    gecici.replace(rapor)


def argumanlar():
    parser = argparse.ArgumentParser(description="21 bayi kaynagindan tek urun yonetim Excel'i olusturur.")
    parser.add_argument("--cikti-dir", type=Path, default=VARSAYILAN_CIKTI_DIR)
    return parser.parse_args()


def main():
    args = argumanlar()
    cikti_dir = args.cikti_dir.resolve()
    birlesik_tsoftta_yok_dosyasi = (
        cikti_dir / "tum_bayiler_tsoftta_yok_b2bde_var.xlsx"
    )
    birlesik_tsoftta_yok = birlesik_tsoftta_yok_urunleri_oku(
        birlesik_tsoftta_yok_dosyasi
    )
    kaynaklar = kaynak_dosyalari_bul(cikti_dir)

    print(f"Okunacak kaynak: {len(kaynaklar)}")
    for firma, dosya in kaynaklar:
        print(f"  {firma}: {dosya.name}")

    if not kaynaklar and birlesik_tsoftta_yok.empty:
        print(f"Kaynak bayi raporu bulunamadi: {cikti_dir}")
        raise SystemExit(1)

    rapor = cikti_dir / "B2B_TSoft_Urun_Yonetim_Raporu.xlsx"

    sorunlar: list[str] = []
    if not birlesik_tsoftta_yok.empty:
        urunler = birlesik_tsoftta_yok
        print(f"Ortak TSoft'ta yok listesi kullanildi: {birlesik_tsoftta_yok_dosyasi.name}")
    else:
        kayitlar, kaynak_sorunlari = kaynaklari_oku(kaynaklar)
        sorunlar.extend(kaynak_sorunlari)
        urunler = urunleri_birlestir(kayitlar)
    eski = eski_manuel_bilgileri_oku(rapor)
    adaylar, stoksuz, marka = uc_tabloyu_hazirla(urunler, eski)
    eslesen_kaynaklar = eslesen_kaynak_dosyalari_bul(cikti_dir)
    eslesen_kayitlar, eslesen_sorunlar = eslesen_kaynaklari_oku(eslesen_kaynaklar)
    sorunlar.extend(eslesen_sorunlar)
    eslesen_marka = eslesen_marka_ozeti_hazirla(eslesen_kayitlar)
    tum_stok_dosyasi = cikti_dir / "tum_bayiler_stok_guncellenecekler.xlsx"
    aylik_stok = aylik_stok_takibini_hazirla(
        rapor,
        tum_stok_dosyasi,
    )
    guncelleme_sayilari = stok_guncelleme_sayilari_oku(tum_stok_dosyasi)

    tek_yonetim_raporu_yaz(
        adaylar,
        stoksuz,
        marka,
        eslesen_marka,
        aylik_stok,
        guncelleme_sayilari,
        rapor,
    )

    print(f"Urun adayi: {len(adaylar)}")
    print(f"Stoksuz urun: {len(stoksuz)}")
    print(f"Marka: {len(marka)}")
    print(f"B2B ve T-Soft'ta eslesen marka: {len(eslesen_marka)}")
    print(f"Stok guncellemesine gonderilecek: {guncelleme_sayilari['gonderilecek']}")
    print(f"Stok guncellemesinden haric: {guncelleme_sayilari['gonderilmeyecek']}")
    print(f"Aylik stok takip urunu: {len(aylik_stok)}")
    print(f"Rapor: {rapor}")
    if sorunlar:
        print("Uyarilar:")
        for sorun in sorunlar:
            print(f"  - {sorun}")


if __name__ == "__main__":
    main()
