"""B2B alışları, iki kâr senaryosu ve rakip fiyatlarını tek Excel'de birleştirir.

Bu betik T-Soft'tan fiyat çekmez ve T-Soft'a veri göndermez.
T-Soft yalnızca ürün kimliği, ürün bilgisi ve fiyat birimini sağlar.
"""

from __future__ import annotations

import argparse
import re
import unicodedata
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


BASE_DIR = Path(__file__).resolve().parent
TSOFT_DOSYASI = BASE_DIR / "Tsoft" / "tsoft_secili_alanlar.xlsx"
ESLESME_DOSYASI = BASE_DIR / "Cikti" / "fiyatlandirma_eslesme_listesi.xlsx"
CIKTI_DOSYASI = BASE_DIR / "Cikti" / "fiyat_onay_listesi.xlsx"
TCMB_KUR_URL = "https://www.tcmb.gov.tr/kurlar/today.xml"

KOLON_ADAYLARI = {
    "liste": [
        "Liste Fiyatı",
        "Liste Fiyati",
        "ListeFiyati",
        "Liste Fiyat",
        "Eski Fiyat",
        "Satış Fiyatı",
        "Fiyat",
        "List Price",
    ],
    "iskonto": [
        "İskonto",
        "İskonto %",
        "İskonto oranları",
        "İndirim %",
        "Iskonto",
        "Iskonto %",
        "Indirim %",
        "İskonto Oranı",
        "Iskonto Orani",
        "Discount",
    ],
    "min_paket": [
        "Min. Paket Adedi",
        "Min Paket Adedi",
        "Minimum Paket Adedi",
        "Paket Adedi",
        "Kutu Adedi",
    ],
    "para_birimi": [
        "Para Birimi",
        "Döviz",
        "Doviz",
        "Currency",
        "Fiyat Birimi",
        "Fiyat Birim",
        "Cinsi",
    ],
}

RAKIP_AYAR_KOLONLARI = [
    "Firma",
    "Aktif",
    "Dosya",
    "KodKolonu",
    "StokKolonu",
    "ListeFiyatiKolonu",
    "SatisFiyatiKolonu",
    "IskontoKolonu",
    "NetSatisKolonu",
    "ParaBirimiKolonu",
    "VarsayilanParaBirimi",
    "FiyatKDVTipi",
]
VARSAYILAN_RAKIP_FIRMALAR = ["STAR AKIM"]


def metin_norm(deger) -> str:
    metin = "" if pd.isna(deger) else str(deger).strip().upper()
    metin = metin.translate(
        str.maketrans(
            {"Ç": "C", "Ğ": "G", "İ": "I", "Ö": "O", "Ş": "S", "Ü": "U", "ı": "I"}
        )
    )
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(c for c in metin if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", " ", metin).strip()


def ondalik(deger, varsayilan=None):
    if deger is None or pd.isna(deger):
        return varsayilan
    metin = str(deger).strip().replace("\xa0", "").replace(" ", "")
    if not metin or metin.lower() == "nan":
        return varsayilan
    metin = re.sub(r"[^0-9,.\-]", "", metin)
    if not metin:
        return varsayilan
    if "," in metin and "." in metin:
        if metin.rfind(",") > metin.rfind("."):
            metin = metin.replace(".", "").replace(",", ".")
        else:
            metin = metin.replace(",", "")
    elif "," in metin:
        metin = metin.replace(",", ".")
    try:
        return Decimal(metin)
    except InvalidOperation:
        return varsayilan


def oran(deger) -> Decimal:
    sayi = ondalik(deger, Decimal("0"))
    return sayi / Decimal("100") if abs(sayi) > 1 else sayi


def iskonto_carpani(deger) -> Decimal:
    """55 veya 20+10 gibi iskonto ifadelerini çarpana dönüştürür."""
    if deger is None or pd.isna(deger) or not str(deger).strip():
        raise ValueError("İskonto boş")
    sonuc = Decimal("1")
    parca_bulundu = False
    for parca in re.split(r"[+;/]", str(deger)):
        if not parca.strip():
            continue
        parca_bulundu = True
        indirim = oran(parca)
        if indirim < 0 or indirim >= 1:
            raise ValueError(f"Geçersiz iskonto: {deger}")
        sonuc *= Decimal("1") - indirim
    if not parca_bulundu:
        raise ValueError("İskonto boş")
    return sonuc


def etkin_iskonto_yuzde(deger) -> Decimal:
    return ((Decimal("1") - iskonto_carpani(deger)) * Decimal("100")).quantize(
        Decimal("0.0001"), rounding=ROUND_HALF_UP
    )


def kolon_bul(kolonlar, adaylar):
    lookup = {metin_norm(k): k for k in kolonlar}
    for aday in adaylar:
        bulunan = lookup.get(metin_norm(aday))
        if bulunan:
            return bulunan
    return None


def para_birimi_bul(liste_degeri="", ayri_para_birimi="") -> str:
    metin = (
        str(liste_degeri)
        .upper()
        .replace("₺", " TL ")
        .replace("€", " EUR ")
        .replace("$", " USD ")
    )
    for kod in ("EUR", "USD", "GBP", "TRL", "TRY", "TL"):
        if re.search(rf"(^|[^A-Z]){kod}([^A-Z]|$)", metin):
            return "TRY" if kod in {"TRL", "TRY", "TL"} else kod
    kod = metin_norm(ayri_para_birimi)
    return "TRY" if kod in {"", "TRL", "TRY", "TL", "TURK LIRASI"} else kod


def dosya_oku(dosya: Path) -> pd.DataFrame:
    if dosya.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(dosya, dtype=str).fillna("")
    son_hata = None
    for encoding in ("utf-8-sig", "cp1254", "latin1"):
        for sep in (";", ",", "\t", "|"):
            try:
                df = pd.read_csv(dosya, dtype=str, encoding=encoding, sep=sep).fillna("")
                if len(df.columns) > 1:
                    return df
            except Exception as exc:
                son_hata = exc
    raise ValueError(f"Dosya okunamadı: {dosya}: {son_hata}")


def mutlak_dosya_yolu(deger) -> Path:
    dosya = Path(str(deger).strip())
    return dosya if dosya.is_absolute() else BASE_DIR / dosya


def tcmb_xml_kurlarini_ayikla(xml_verisi: bytes) -> dict:
    try:
        kok = ET.fromstring(xml_verisi)
    except ET.ParseError as exc:
        raise ValueError(f"TCMB kur XML'i okunamadı: {exc}") from exc
    kurlar = {"TRY": Decimal("1"), "TL": Decimal("1")}
    for para_birimi in kok.findall("Currency"):
        kod = str(para_birimi.attrib.get("CurrencyCode", "")).strip().upper()
        if kod not in {"USD", "EUR"}:
            continue
        satis = ondalik(para_birimi.findtext("ForexSelling", ""))
        birim = ondalik(para_birimi.findtext("Unit", "1"), Decimal("1"))
        if satis is not None and satis > 0 and birim is not None and birim > 0:
            kurlar[kod] = satis / birim
    eksik = [kod for kod in ("USD", "EUR") if kod not in kurlar]
    if eksik:
        raise ValueError(f"TCMB cevabında kur bulunamadı: {', '.join(eksik)}")
    return {
        "kdv": Decimal("0.20"),
        "kurlar": kurlar,
        "rakipler": [],
        "kur_kaynagi": TCMB_KUR_URL,
        "kur_tarihi": str(kok.attrib.get("Tarih", "")).strip(),
        "kur_cekim_zamani": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def tcmb_kurlarini_cek(url: str = TCMB_KUR_URL) -> dict:
    istek = urllib.request.Request(
        url,
        headers={"User-Agent": "EntesStokBot/1.0", "Accept": "application/xml,text/xml"},
    )
    try:
        with urllib.request.urlopen(istek, timeout=30) as cevap:
            xml_verisi = cevap.read()
    except Exception as exc:
        raise RuntimeError(
            "TCMB kurları alınamadı. İnternet bağlantısını kontrol edip tekrar deneyin: "
            f"{exc}"
        ) from exc
    return tcmb_xml_kurlarini_ayikla(xml_verisi)


def kaynak_fiyatini_oku(
    kaynak_dosya,
    satir_no,
    cache: dict,
    ayarlar: dict,
) -> dict:
    dosya = mutlak_dosya_yolu(kaynak_dosya)
    if not dosya.exists():
        raise FileNotFoundError(f"B2B kaynak dosyası yok: {dosya}")
    anahtar = str(dosya.resolve())
    if anahtar not in cache:
        cache[anahtar] = dosya_oku(dosya)
    df = cache[anahtar]
    indeks = int(float(str(satir_no))) - 2
    if indeks < 0 or indeks >= len(df):
        raise IndexError(f"B2B satırı bulunamadı: {dosya}, satır {satir_no}")
    row = df.iloc[indeks]
    kolonlar = {alan: kolon_bul(df.columns, adaylar) for alan, adaylar in KOLON_ADAYLARI.items()}
    liste_ham = row.get(kolonlar["liste"], "") if kolonlar["liste"] else ""
    iskonto_ham = row.get(kolonlar["iskonto"], "") if kolonlar["iskonto"] else ""
    min_paket_ham = row.get(kolonlar["min_paket"], "") if kolonlar["min_paket"] else ""
    para_ham = row.get(kolonlar["para_birimi"], "") if kolonlar["para_birimi"] else ""
    liste = ondalik(liste_ham)
    if liste is None or liste <= 0:
        raise ValueError("Liste fiyatı boş veya sıfır")
    doviz = para_birimi_bul(liste_ham, para_ham)
    kur = ayarlar["kurlar"].get(doviz)
    if kur is None:
        raise ValueError(f"{doviz} kuru ayarlanmamış")
    liste_tl = (liste * kur).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    iskonto_yuzde = etkin_iskonto_yuzde(iskonto_ham)
    net_tl_kdv_haric = (liste * iskonto_carpani(iskonto_ham) * kur).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    min_paket = ondalik(min_paket_ham, Decimal("1")) or Decimal("1")
    return {
        "liste_tl": liste_tl,
        "iskonto_yuzde": iskonto_yuzde,
        "net_tl_kdv_haric": net_tl_kdv_haric,
        "min_paket": min_paket,
        "doviz": doviz,
    }


def rakip_verilerini_oku(rakip_ayarlari: list[dict], ayarlar: dict):
    firmalar, harita, hatalar = [], {}, []
    for ayar in rakip_ayarlari:
        firma = str(ayar["Firma"]).strip()
        firma_norm = metin_norm(firma)
        firmalar.append(firma)
        try:
            dosya = mutlak_dosya_yolu(ayar["Dosya"])
            if not dosya.exists():
                raise FileNotFoundError(f"Rakip dosyası yok: {dosya}")
            df = dosya_oku(dosya)
            kod_col = kolon_bul(df.columns, [ayar["KodKolonu"], "Web Servis Kodu", "Ürün Kodu", "Kod"])
            stok_col = kolon_bul(df.columns, [ayar["StokKolonu"], "Stok"])
            liste_col = kolon_bul(
                df.columns,
                [ayar.get("ListeFiyatiKolonu", ""), "Liste Fiyatı", "Liste"],
            )
            satis_col = kolon_bul(
                df.columns,
                [
                    ayar.get("SatisFiyatiKolonu", ""),
                    "Satış Fiyatı",
                    "Satis Fiyati",
                    "Fiyat",
                ],
            )
            isk_col = kolon_bul(
                df.columns,
                [ayar.get("IskontoKolonu", ""), "İskonto", "Iskonto"],
            )
            net_col = kolon_bul(
                df.columns,
                [
                    ayar.get("NetSatisKolonu", ""),
                    "Net Satış",
                    "Net Satış (KDV Hariç)",
                    "Net Satış KDV Hariç",
                ],
            )
            para_col = kolon_bul(
                df.columns,
                [ayar.get("ParaBirimiKolonu", ""), "Para Birimi", "Fiyat Birimi"],
            )
            if not kod_col:
                raise ValueError("Rakip ürün kodu kolonu bulunamadı")
            if not satis_col and not net_col and not liste_col:
                raise ValueError("Rakip satış fiyatı kolonu bulunamadı")
            for i, row in df.iterrows():
                kod = metin_norm(row.get(kod_col, ""))
                if not kod:
                    continue
                para_ham = (
                    row.get(para_col, "")
                    if para_col
                    else ayar.get("VarsayilanParaBirimi", "TRY")
                )
                fiyat_col = satis_col or net_col or liste_col
                fiyat_ham = row.get(fiyat_col, "")
                doviz = para_birimi_bul(fiyat_ham, para_ham)
                kur = ayarlar["kurlar"].get(doviz)
                if kur is None:
                    raise ValueError(f"{doviz} kuru ayarlanmamış")
                iskonto_ham = row.get(isk_col, "") if isk_col else ""
                fiyat = ondalik(fiyat_ham)
                if fiyat is None or fiyat <= 0:
                    continue
                satis_tl = fiyat * kur
                if satis_col or liste_col:
                    if iskonto_ham is None or not str(iskonto_ham).strip():
                        net_tl = satis_tl
                    else:
                        net_tl = satis_tl * iskonto_carpani(iskonto_ham)
                else:
                    net_tl = satis_tl
                if metin_norm(ayar.get("FiyatKDVTipi", "KDV DAHIL")) in {
                    "KDV DAHIL",
                    "DAHIL",
                }:
                    net_tl /= Decimal("1") + ayarlar["kdv"]
                if net_tl <= 0:
                        continue
                kayit = {
                    "stok": str(row.get(stok_col, "")).strip() if stok_col else "",
                    "iskonto": (
                        float(etkin_iskonto_yuzde(iskonto_ham))
                        if iskonto_ham
                        else None
                    ),
                    "satis_fiyati": satis_tl.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "net_satis_kdv_haric": net_tl.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                }
                anahtar = (firma_norm, kod)
                onceki = harita.get(anahtar)
                if (
                    onceki is None
                    or kayit["net_satis_kdv_haric"] < onceki["net_satis_kdv_haric"]
                ):
                    harita[anahtar] = kayit
        except Exception as exc:
            hatalar.append(
                {
                    "ProductId": "",
                    "Web Servis Kodu": "",
                    "Firma": firma,
                    "Kaynak Dosya": ayar.get("Dosya", ""),
                    "Kaynak Satır": "",
                    "Hata": str(exc),
                }
            )
    return firmalar, harita, hatalar


def tsoft_haritasi_olustur(tsoft: pd.DataFrame) -> tuple[dict, dict]:
    product_map, kod_map = {}, {}
    for _, row in tsoft.iterrows():
        kayit = row.to_dict()
        pid = str(row.get("ProductId", "")).strip()
        kod = metin_norm(row.get("Web Servis Kodu", ""))
        if pid:
            product_map[pid] = kayit
        if kod:
            kod_map[kod] = kayit
    return product_map, kod_map


def fiyat_verisini_hazirla(eslesme: pd.DataFrame, tsoft: pd.DataFrame, ayarlar: dict):
    eslesme = eslesme.copy()
    eslesme_adaylari = {
        "ProductId": ["ProductId"],
        "Web Servis Kodu": ["Web Servis Kodu"],
        "Tedarikci": ["Tedarikci", "Tedarikçi"],
        "B2B Kaynak Dosya": ["B2B Kaynak Dosya"],
        "B2B Satır No": ["B2B Satır No", "B2B Satir No", "B2B SatÄ±r No"],
    }
    yeniden_adlandir = {}
    for standart, adaylar in eslesme_adaylari.items():
        bulunan = kolon_bul(eslesme.columns, adaylar)
        if bulunan and bulunan != standart:
            yeniden_adlandir[bulunan] = standart
    if yeniden_adlandir:
        eslesme = eslesme.rename(columns=yeniden_adlandir)
    gerekli = list(eslesme_adaylari)
    eksik = [k for k in gerekli if k not in eslesme.columns]
    if eksik:
        raise ValueError(f"Eşleşme dosyasında eksik kolonlar: {', '.join(eksik)}")
    product_map, kod_map = tsoft_haritasi_olustur(tsoft)
    cache, fiyatlar, hatalar, tedarikciler = {}, {}, [], set()

    for _, row in eslesme.iterrows():
        pid = str(row.get("ProductId", "")).strip()
        kod = str(row.get("Web Servis Kodu", "")).strip()
        tedarikci = metin_norm(row.get("Tedarikci", ""))
        if not pid or not kod or not tedarikci:
            continue
        tedarikciler.add(tedarikci)
        try:
            fiyat = kaynak_fiyatini_oku(
                row.get("B2B Kaynak Dosya", ""),
                row.get("B2B Satır No", ""),
                cache,
                ayarlar,
            )
            anahtar = (pid, tedarikci)
            if (
                anahtar not in fiyatlar
                or fiyat["net_tl_kdv_haric"] < fiyatlar[anahtar]["net_tl_kdv_haric"]
            ):
                fiyatlar[anahtar] = fiyat
        except Exception as exc:
            hatalar.append(
                {
                    "ProductId": pid,
                    "Web Servis Kodu": kod,
                    "Firma": tedarikci,
                    "Kaynak Dosya": row.get("B2B Kaynak Dosya", ""),
                    "Kaynak Satır": row.get("B2B Satır No", ""),
                    "Hata": str(exc),
                }
            )

    urunler = []
    urun_anahtarlari = eslesme[["ProductId", "Web Servis Kodu"]].drop_duplicates()
    for _, kimlik in urun_anahtarlari.iterrows():
        pid = str(kimlik["ProductId"]).strip()
        kod = str(kimlik["Web Servis Kodu"]).strip()
        ts = product_map.get(pid) or kod_map.get(metin_norm(kod), {})
        eslesen_satirlar = eslesme[
            (eslesme["ProductId"].astype(str).str.strip() == pid)
            & (eslesme["Web Servis Kodu"].astype(str).str.strip() == kod)
        ]
        marka = ""
        for kolon in ("B2B Marka", "Marka"):
            if kolon in eslesen_satirlar.columns:
                degerler = eslesen_satirlar[kolon].astype(str).str.strip()
                degerler = degerler[(degerler != "") & (degerler.str.lower() != "nan")]
                if not degerler.empty:
                    marka = degerler.iloc[0]
                    break
        urun = {
            "ProductId": pid,
            "Marka": marka,
            "Kategori": ts.get("Ana Kategori Adı", ""),
            "Web Servis Kodu": kod,
            "Ürün Adı": ts.get("Ürün Adı", "") or ts.get("Model", ""),
            "Fiyat Birimi": ts.get("Fiyat Birimi", "TRY") or "TRY",
        }
        for tedarikci in sorted(tedarikciler):
            urun[tedarikci] = fiyatlar.get((pid, tedarikci))
        urunler.append(urun)
    return urunler, sorted(tedarikciler), hatalar


def _formul_ortalama(kolonlar, satir_no):
    hucreler = [f"{get_column_letter(c)}{satir_no}" for c in kolonlar]
    return f'=IFERROR(AVERAGE({",".join(hucreler)}),"")' if hucreler else '=""'


def rapor_olustur(
    urunler,
    tedarikciler,
    rakip_firmalar,
    rakip_haritasi,
    hatalar,
    ayarlar,
    cikti: Path,
) -> Path:
    rakip_firmalar = list(
        dict.fromkeys([*VARSAYILAN_RAKIP_FIRMALAR, *rakip_firmalar])
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "Fiyatlandırma"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "F3"

    sabit = ["ProductId", "Marka", "Kategori", "Ürün Web Servis Kodu", "Ürün Adı", "Fiyat Birimi"]
    b2b = []
    for t in tedarikciler:
        b2b.extend(
            [
                f"{t} LİSTE",
                f"{t} İSKONTO",
                f"{t} NET ALIŞ KDV HARİÇ",
                f"{t} MİN. PAKET",
            ]
        )
    min_kar = [
        "MİN. ORT. İSKONTO",
        "MİN. ORT. NET ALIŞ",
        "MİN. KÂR ORANI (MANUEL)",
        "MİN. SATIŞ KDV HARİÇ",
        "MİN. HESAPLANAN İSKONTO",
    ]
    hedef_kar = [
        "HEDEF ORT. İSKONTO",
        "HEDEF ORT. NET ALIŞ",
        "HEDEF KÂR ORANI (MANUEL)",
        "HEDEF SATIŞ KDV HARİÇ",
        "HEDEF HESAPLANAN İSKONTO",
    ]
    rakip = []
    for firma in rakip_firmalar:
        rakip.extend(
            [
                f"{firma} STOK",
                f"{firma} İSKONTO",
                f"{firma} SATIŞ FİYATI",
                f"{firma} NET SATIŞ KDV HARİÇ",
            ]
        )
    karar = [
        "SEÇİLEN İSKONTO",
        "ONAY",
        "AKTARIM DURUMU",
    ]
    basliklar = sabit + b2b + min_kar + hedef_kar + rakip + karar
    index = {ad: i + 1 for i, ad in enumerate(basliklar)}

    gruplar = [
        ("ÜRÜN BİLGİLERİ", len(sabit), "1F4E78"),
        ("B2B TEDARİKÇİLER", len(b2b), "548235"),
        ("MİN. KÂR ORANI", len(min_kar), "BF9000"),
        ("HEDEFLENEN KÂR ORANI", len(hedef_kar), "7030A0"),
        ("RAKİP FİRMALAR", len(rakip), "C00000"),
        ("FİYAT KARARI", len(karar), "305496"),
    ]
    baslangic = 1
    for grup_adi, uzunluk, renk in gruplar:
        if uzunluk <= 0:
            continue
        bitis = baslangic + uzunluk - 1
        ws.merge_cells(start_row=1, start_column=baslangic, end_row=1, end_column=bitis)
        cell = ws.cell(1, baslangic, grup_adi)
        cell.fill = PatternFill("solid", fgColor=renk)
        cell.font = Font(color="FFFFFF", bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        baslangic = bitis + 1

    for c, ad in enumerate(basliklar, 1):
        cell = ws.cell(2, c, ad)
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    liste_kolonlari = [index[f"{t} LİSTE"] for t in tedarikciler]
    iskonto_kolonlari = [index[f"{t} İSKONTO"] for t in tedarikciler]
    net_kolonlari = [index[f"{t} NET ALIŞ KDV HARİÇ"] for t in tedarikciler]
    for satir_no, urun in enumerate(urunler, 3):
        ws.cell(satir_no, index["ProductId"], urun["ProductId"])
        ws.cell(satir_no, index["Marka"], urun["Marka"])
        ws.cell(satir_no, index["Kategori"], urun["Kategori"])
        ws.cell(satir_no, index["Ürün Web Servis Kodu"], urun["Web Servis Kodu"])
        ws.cell(satir_no, index["Ürün Adı"], urun["Ürün Adı"])
        ws.cell(satir_no, index["Fiyat Birimi"], urun["Fiyat Birimi"])
        for t in tedarikciler:
            f = urun.get(t)
            if not f:
                continue
            liste_col = index[f"{t} LİSTE"]
            isk_col = index[f"{t} İSKONTO"]
            net_haric_col = index[f"{t} NET ALIŞ KDV HARİÇ"]
            ws.cell(satir_no, liste_col, float(f["liste_tl"]))
            ws.cell(satir_no, isk_col, float(f["iskonto_yuzde"]))
            liste_ref = f"{get_column_letter(liste_col)}{satir_no}"
            isk_ref = f"{get_column_letter(isk_col)}{satir_no}"
            ws.cell(satir_no, net_haric_col, f'=IFERROR({liste_ref}*(1-{isk_ref}/100),"")')
            ws.cell(satir_no, index[f"{t} MİN. PAKET"], float(f["min_paket"]))

        ort_isk_degerleri = []
        ort_net_degerleri = []
        ort_liste_degerleri = []
        ort_isk_sayimlari = []
        ort_net_sayimlari = []
        ort_liste_sayimlari = []
        for t in tedarikciler:
            liste_ref = f"{get_column_letter(index[f'{t} LİSTE'])}{satir_no}"
            isk_ref = f"{get_column_letter(index[f'{t} İSKONTO'])}{satir_no}"
            net_ref = (
                f"{get_column_letter(index[f'{t} NET ALIŞ KDV HARİÇ'])}{satir_no}"
            )
            if t == "KONYA":
                ort_isk_degerleri.append(f"IF({isk_ref}<>0,{isk_ref},0)")
                ort_net_degerleri.append(f"IF({isk_ref}<>0,{net_ref},0)")
                ort_liste_degerleri.append(f"IF({isk_ref}<>0,{liste_ref},0)")
                ort_isk_sayimlari.append(
                    f"IF(AND({isk_ref}<>0,ISNUMBER({isk_ref})),1,0)"
                )
                ort_net_sayimlari.append(
                    f"IF(AND({isk_ref}<>0,ISNUMBER({net_ref})),1,0)"
                )
                ort_liste_sayimlari.append(
                    f"IF(AND({isk_ref}<>0,ISNUMBER({liste_ref})),1,0)"
                )
            else:
                ort_isk_degerleri.append(isk_ref)
                ort_net_degerleri.append(net_ref)
                ort_liste_degerleri.append(liste_ref)
                ort_isk_sayimlari.append(f"IF(ISNUMBER({isk_ref}),1,0)")
                ort_net_sayimlari.append(f"IF(ISNUMBER({net_ref}),1,0)")
                ort_liste_sayimlari.append(f"IF(ISNUMBER({liste_ref}),1,0)")
        ort_isk_ifade = (
            f'IFERROR(SUM({",".join(ort_isk_degerleri)})/'
            f'SUM({",".join(ort_isk_sayimlari)}),"")'
        )
        ort_net_ifade = (
            f'IFERROR(SUM({",".join(ort_net_degerleri)})/'
            f'SUM({",".join(ort_net_sayimlari)}),"")'
        )
        ort_liste_ifade = (
            f'IFERROR(SUM({",".join(ort_liste_degerleri)})/'
            f'SUM({",".join(ort_liste_sayimlari)}),"")'
        )
        ort_isk_formul = (
            f"={ort_isk_ifade}"
            if ort_isk_degerleri
            else '=""'
        )
        ort_net_formul = (
            f"={ort_net_ifade}"
            if ort_net_degerleri
            else '=""'
        )
        ort_liste_formul = (
            ort_liste_ifade
            if ort_liste_degerleri
            else '""'
        )

        ws.cell(satir_no, index["MİN. ORT. İSKONTO"], ort_isk_formul)
        ws.cell(satir_no, index["MİN. ORT. NET ALIŞ"], ort_net_formul)
        min_ort = f"{get_column_letter(index['MİN. ORT. NET ALIŞ'])}{satir_no}"
        min_kar_ref = f"{get_column_letter(index['MİN. KÂR ORANI (MANUEL)'])}{satir_no}"
        min_satis = f"{get_column_letter(index['MİN. SATIŞ KDV HARİÇ'])}{satir_no}"
        ws.cell(
            satir_no,
            index["MİN. SATIŞ KDV HARİÇ"],
            f'=IF(AND({min_ort}>0,{min_kar_ref}<>""),{min_ort}*(1+{min_kar_ref}/100),"")',
        )
        ws.cell(
            satir_no,
            index["MİN. HESAPLANAN İSKONTO"],
            f'=IFERROR((1-{min_satis}/{ort_liste_formul})*100,"")',
        )

        ws.cell(satir_no, index["HEDEF ORT. İSKONTO"], ort_isk_formul)
        ws.cell(satir_no, index["HEDEF ORT. NET ALIŞ"], ort_net_formul)
        hedef_ort = f"{get_column_letter(index['HEDEF ORT. NET ALIŞ'])}{satir_no}"
        hedef_kar_ref = (
            f"{get_column_letter(index['HEDEF KÂR ORANI (MANUEL)'])}{satir_no}"
        )
        hedef_satis = f"{get_column_letter(index['HEDEF SATIŞ KDV HARİÇ'])}{satir_no}"
        ws.cell(
            satir_no,
            index["HEDEF SATIŞ KDV HARİÇ"],
            f'=IF(AND({hedef_ort}>0,{hedef_kar_ref}<>""),{hedef_ort}*(1+{hedef_kar_ref}/100),"")',
        )
        ws.cell(
            satir_no,
            index["HEDEF HESAPLANAN İSKONTO"],
            f'=IFERROR((1-{hedef_satis}/{ort_liste_formul})*100,"")',
        )

        kod_norm = metin_norm(urun["Web Servis Kodu"])
        for firma in rakip_firmalar:
            kayit = rakip_haritasi.get((metin_norm(firma), kod_norm))
            if not kayit:
                continue
            ws.cell(satir_no, index[f"{firma} STOK"], kayit["stok"])
            ws.cell(satir_no, index[f"{firma} İSKONTO"], kayit["iskonto"])
            ws.cell(
                satir_no,
                index[f"{firma} SATIŞ FİYATI"],
                float(kayit["satis_fiyati"]),
            )
            ws.cell(
                satir_no,
                index[f"{firma} NET SATIŞ KDV HARİÇ"],
                float(kayit["net_satis_kdv_haric"]),
            )

        secilen_iskonto = f"{get_column_letter(index['SEÇİLEN İSKONTO'])}{satir_no}"
        secilen = f"({ort_liste_formul}*(1-{secilen_iskonto}/100))"
        onay = f"{get_column_letter(index['ONAY'])}{satir_no}"
        ws.cell(
            satir_no,
            index["AKTARIM DURUMU"],
            (
                f'=IF({min_ort}="","ALIŞ FİYATI YOK",'
                f'IF(OR({min_kar_ref}="",{hedef_kar_ref}=""),"KÂR ORANI BEKLİYOR",'
                f'IF({min_satis}>{hedef_satis},"KÂR ARALIĞI HATALI",'
                f'IF({secilen_iskonto}="","İSKONTO SEÇİMİ BEKLİYOR",'
                f'IF(OR({secilen}<{min_satis},{secilen}>{hedef_satis}),"FİYAT ARALIK DIŞI",'
                f'IF({onay}="EVET","AKTARIMA HAZIR",'
                f'IF({onay}="HAYIR","REDDEDİLDİ","ONAY BEKLİYOR")))))))'
            ),
        )

    son_satir = max(len(urunler) + 2, 3)
    thin = Side(style="thin", color="D9E2F3")
    kaynak_fill = PatternFill("solid", fgColor="D9EAF7")
    b2b_fill = PatternFill("solid", fgColor="E2F0D9")
    min_fill = PatternFill("solid", fgColor="FFF2CC")
    hedef_fill = PatternFill("solid", fgColor="E4DFEC")
    rakip_fill = PatternFill("solid", fgColor="FCE4D6")
    karar_fill = PatternFill("solid", fgColor="DDEBF7")
    input_fill = PatternFill("solid", fgColor="FFD966")
    bolum_sinirlari = [
        (1, len(sabit), kaynak_fill),
        (len(sabit) + 1, len(sabit) + len(b2b), b2b_fill),
        (
            len(sabit) + len(b2b) + 1,
            len(sabit) + len(b2b) + len(min_kar),
            min_fill,
        ),
        (
            len(sabit) + len(b2b) + len(min_kar) + 1,
            len(sabit) + len(b2b) + len(min_kar) + len(hedef_kar),
            hedef_fill,
        ),
        (
            len(sabit) + len(b2b) + len(min_kar) + len(hedef_kar) + 1,
            len(sabit) + len(b2b) + len(min_kar) + len(hedef_kar) + len(rakip),
            rakip_fill,
        ),
        (
            len(basliklar) - len(karar) + 1,
            len(basliklar),
            karar_fill,
        ),
    ]
    for satir in ws.iter_rows(min_row=3, max_row=son_satir, min_col=1, max_col=len(basliklar)):
        for cell in satir:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="center")
            for ilk, son, fill in bolum_sinirlari:
                if ilk <= cell.column <= son:
                    cell.fill = fill
                    break
    for ad in [
        "MİN. KÂR ORANI (MANUEL)",
        "HEDEF KÂR ORANI (MANUEL)",
        "SEÇİLEN İSKONTO",
        "ONAY",
    ]:
        for r in range(3, son_satir + 1):
            ws.cell(r, index[ad]).fill = input_fill

    kar_dv = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="100",
        allow_blank=True,
    )
    ws.add_data_validation(kar_dv)
    for ad in ["MİN. KÂR ORANI (MANUEL)", "HEDEF KÂR ORANI (MANUEL)"]:
        col = get_column_letter(index[ad])
        kar_dv.add(f"{col}3:{col}{son_satir}")
    iskonto_dv = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="99.99",
        allow_blank=True,
    )
    ws.add_data_validation(iskonto_dv)
    iskonto_col = get_column_letter(index["SEÇİLEN İSKONTO"])
    iskonto_dv.add(f"{iskonto_col}3:{iskonto_col}{son_satir}")
    onay_dv = DataValidation(type="list", formula1='"EVET,HAYIR"', allow_blank=True)
    ws.add_data_validation(onay_dv)
    onay_col = get_column_letter(index["ONAY"])
    onay_dv.add(f"{onay_col}3:{onay_col}{son_satir}")

    durum_col = get_column_letter(index["AKTARIM DURUMU"])
    ws.conditional_formatting.add(
        f"{durum_col}3:{durum_col}{son_satir}",
        FormulaRule(
            formula=[f'${durum_col}3="AKTARIMA HAZIR"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
        ),
    )
    ws.conditional_formatting.add(
        f"{durum_col}3:{durum_col}{son_satir}",
        FormulaRule(
            formula=[f'OR(${durum_col}3="FİYAT ARALIK DIŞI",${durum_col}3="KÂR ARALIĞI HATALI")'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
        ),
    )

    for ad in basliklar:
        if any(k in ad for k in ["LİSTE", "NET ALIŞ", "SATIŞ"]):
            for r in range(3, son_satir + 1):
                ws.cell(r, index[ad]).number_format = '#,##0.00'
        if any(k in ad for k in ["İSKONTO", "KÂR ORANI"]):
            for r in range(3, son_satir + 1):
                ws.cell(r, index[ad]).number_format = '0.00'

    genislik = {
        "ProductId": 12,
        "Marka": 18,
        "Kategori": 25,
        "Ürün Web Servis Kodu": 25,
        "Ürün Adı": 52,
        "Fiyat Birimi": 13,
        "ONAY": 12,
        "AKTARIM DURUMU": 24,
    }
    for c, ad in enumerate(basliklar, 1):
        ws.column_dimensions[get_column_letter(c)].width = genislik.get(ad, 18)
    ws.auto_filter.ref = f"A2:{get_column_letter(len(basliklar))}{son_satir}"
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 55

    kur_ws = wb.create_sheet("Kurlar")
    kur_ws.append(["Para Birimi", "TL Kuru", "Kaynak", "Kur Tarihi", "Çekim Zamanı"])
    for kod in ("TRY", "USD", "EUR"):
        kur_ws.append(
            [
                kod,
                float(ayarlar["kurlar"][kod]),
                "TCMB Döviz Satış" if kod != "TRY" else "Sabit",
                ayarlar.get("kur_tarihi", ""),
                ayarlar.get("kur_cekim_zamani", ""),
            ]
        )
    kur_ws.append([])
    kur_ws.append(
        [
            "KDV Oranı %",
            float(ayarlar["kdv"] * 100),
            "KDV hariç fiyatlardan KDV dahil fiyat üretir",
            "",
            "",
        ]
    )
    kur_ws.append(
        [
            "Kur Kaynağı",
            ayarlar.get("kur_kaynagi", TCMB_KUR_URL),
            "Fiyat raporu oluşturulurken otomatik alınır",
            "",
            "",
        ]
    )
    for cell in kur_ws[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
    kur_ws.column_dimensions["A"].width = 20
    kur_ws.column_dimensions["B"].width = 22
    kur_ws.column_dimensions["C"].width = 36
    kur_ws.column_dimensions["D"].width = 16
    kur_ws.column_dimensions["E"].width = 22

    hata_ws = wb.create_sheet("Hatalar")
    hata_baslik = ["ProductId", "Web Servis Kodu", "Firma", "Kaynak Dosya", "Kaynak Satır", "Hata"]
    hata_ws.append(hata_baslik)
    for hata in hatalar:
        hata_ws.append([hata.get(k, "") for k in hata_baslik])
    for cell in hata_ws[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
    hata_ws.freeze_panes = "A2"
    for c, width in enumerate([12, 24, 18, 55, 14, 55], 1):
        hata_ws.column_dimensions[get_column_letter(c)].width = width

    cikti.parent.mkdir(exist_ok=True)
    wb.save(cikti)
    return cikti


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="B2B alış + minimum/hedef kâr + rakip fiyat Excel'i oluşturur"
    )
    parser.add_argument("--tsoft", type=Path, default=TSOFT_DOSYASI)
    parser.add_argument("--eslesme", type=Path, default=ESLESME_DOSYASI)
    parser.add_argument("--cikti", type=Path, default=CIKTI_DOSYASI)
    args = parser.parse_args(argv)
    try:
        if not args.tsoft.exists():
            raise FileNotFoundError(
                f"T-Soft ürün dosyası yok; önce tsoft_products_export.py çalıştırın: {args.tsoft}"
            )
        if not args.eslesme.exists():
            raise FileNotFoundError(
                "Fiyatlandırma eşleşme dosyası yok; önce güncel "
                f"stok_eslestir.py dosyasını çalıştırın: {args.eslesme}"
            )
        print(f"TCMB döviz satış kurları alınıyor: {TCMB_KUR_URL}")
        ayarlar = tcmb_kurlarini_cek()
        print(
            "Kurlar: "
            f"USD={ayarlar['kurlar']['USD']} TL, "
            f"EUR={ayarlar['kurlar']['EUR']} TL "
            f"(TCMB tarihi: {ayarlar.get('kur_tarihi', '')})"
        )
        tsoft = pd.read_excel(args.tsoft, dtype=str).fillna("")
        eslesme = pd.read_excel(args.eslesme, dtype=str).fillna("")
        urunler, tedarikciler, hatalar = fiyat_verisini_hazirla(eslesme, tsoft, ayarlar)
        rakip_firmalar, rakip_haritasi, rakip_hatalari = rakip_verilerini_oku(
            ayarlar["rakipler"], ayarlar
        )
        rapor = rapor_olustur(
            urunler,
            tedarikciler,
            rakip_firmalar,
            rakip_haritasi,
            hatalar + rakip_hatalari,
            ayarlar,
            args.cikti,
        )
        print(f"Fiyatlandırılan ürün: {len(urunler)}")
        print(f"B2B tedarikçi: {len(tedarikciler)}")
        print(f"Rakip firma: {len(rakip_firmalar)}")
        print(f"Veri hatası: {len(hatalar) + len(rakip_hatalari)}")
        print(f"Fiyatlandırma dosyası: {rapor}")
        print("Bu işlem T-Soft'a fiyat göndermez.")
        return 0
    except Exception as exc:
        print(f"HATA: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
