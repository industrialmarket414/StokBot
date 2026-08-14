"""B2B alışları, iki kâr senaryosu ve rakip fiyatlarını tek Excel'de birleştirir.

Bu betik T-Soft'a veri göndermez. T-Soft ürün aktarımının oluşturduğu Excel'den
ürün bilgileri, fiyat birimi ve liste fiyatını okur.
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
RAKIP_DIZINI = BASE_DIR / "Rakipler"
KATEGORI_KAR_ORANI_DOSYASI = BASE_DIR / "Ayarlar" / "kategori_kar_oranlari.xlsx"
PAKET_BIRIM_AYAR_DOSYASI = BASE_DIR / "Ayarlar" / "paket_birim_ayarlari.xlsx"
KATEGORI_KAR_YEDEK_DOSYASI = Path.home() / "OneDrive" / "Desktop" / "karyedek.xlsx"
ENTES_FIYAT_LISTESI_DOSYASI = (
    BASE_DIR / "FiyatListesi" / "ENTES_FIYAT_LISTESI.xlsx"
)
TCMB_KUR_URL = "https://www.tcmb.gov.tr/kurlar/today.xml"
GENEL_MIN_KAR_ORANI = Decimal("7")

# T-Soft'ta bu kisa kodlar GWEST urunleridir. Ayni kodlari kullanan NEAD
# urunleri hem B2B alisinda hem rakip eslestirmesinde bu urunlere baglanamaz.
FIYATLANDIRMA_KOD_MARKA_ONCELIGI = {
    # KIRAC/Elektrix'teki SAFAK KM-240 kablo makasi, T-Soft'taki
    # KARDES ELEKTRIK KM 240 ek muf ile ayni urun degildir.
    "KM 240": "KARDES ELEKTRIK",
    # Uc Faz'daki ayni kodlu VIKO telefon prizi, T-Soft'taki ARTLINE VGA
    # kapagiyla ayni urun degildir. Bu kodda yalniz ARTLINE rakibi kabul edilir.
    "92605513": "ARTLINE",
    "1010": "GWEST",
    "2007": "GWEST",
    "2008": "GWEST",
    "4005": "GWEST",
    "4007": "GWEST",
    "4008": "GWEST",
    "4009": "GWEST",
    "4010": "GWEST",
    "4011": "GWEST",
}

# Dogrulanmis yanlis paket kurallari. Bu satirlar kaynak ayar dosyasinda kalsa
# bile fiyat hesabina ve yeni raporun Paket Birim Ayarlari sayfasina alinmaz.
GECERSIZ_PAKET_BIRIM_KURALLARI = {
    # T-Soft KARDES KM 240 ek muf; KIRAC/ELEKTRIX KM-240 ise SAFAK makastir.
    ("KM 240", "KIRAC", "ALIS"),
    ("KM 240", "ELEKTRIX", "RAKIP"),
    # KONYA 570040 fiyati zaten tek adet icindir; yalniz rakipteki 500'lu
    # paket fiyati bolunmelidir.
    ("570040", "KONYA", "ALIS"),
    # Bu iki KONYA fiyati da tek adet icindir; 100'lu paket yalniz rakip
    # Elektrik Malzemesi fiyatinda uygulanir.
    ("590024", "KONYA", "ALIS"),
    ("591012", "KONYA", "ALIS"),
    # Eski kurallar ayni kodlu NEAD fiyatlarini GWEST paket adedine boluyordu.
    # GWEST B2B liste fiyatlari zaten birimdir; yalniz paketli rakipler bolunur.
    ("1010", "KONYA", "ALIS"),
    ("2007", "YILDIZ", "ALIS"),
    ("2008", "YILDIZ", "ALIS"),
    ("4005", "YILDIZ", "ALIS"),
    ("4005", "STAR AKIM", "RAKIP"),
    ("4007", "KONYA", "ALIS"),
    ("4008", "YILDIZ", "ALIS"),
}

# Kaynak urun adindan ve rakip sayfasindan elle dogrulanmis paket fiyatlari.
# Bu kurallar ayar Excel'i yenilense bile kaybolmaz.
DOGRULANMIS_RAKIP_PAKET_KURALLARI = {
    # Buyuk Kalpler satis fiyati 5 metrelik paketin tamamidir.
    ("YL42-1219", "BUYUK KALPLER"): Decimal("5"),
}


def paket_birim_ayarlarini_oku(dosya: Path = PAKET_BIRIM_AYAR_DOSYASI) -> dict:
    """Onaylanmış ürün+firma paket fiyatlarını birim fiyata dönüştürme haritasını okur."""
    dosya = Path(dosya)
    sonuc = {"alis": {}, "alis_marka": {}, "rakip": {}, "satirlar": []}
    if not dosya.exists():
        return sonuc
    df = pd.read_excel(
        dosya,
        sheet_name="Paket Birim Ayarları",
        header=2,
        dtype=str,
    ).fillna("")
    gerekli = {
        "Aktif", "ProductId", "Marka", "Kategori", "Ürün Web Servis Kodu",
        "Firma", "Kaynak Türü", "Paket Adedi", "Fiyatın Kapsadığı Adet",
        "Kontrol Durumu",
    }
    eksik = gerekli - set(df.columns)
    if eksik:
        raise ValueError(
            "paket_birim_ayarlari.xlsx eksik kolonlar: " + ", ".join(sorted(eksik))
        )
    for _, row in df.iterrows():
        kayit = {kolon: str(row.get(kolon, "")).strip() for kolon in gerekli}
        if metin_norm(kayit["Aktif"]) != "EVET":
            continue
        if metin_norm(kayit["Kontrol Durumu"]) != "ONAYLANDI":
            continue
        kod = metin_norm(kayit["Ürün Web Servis Kodu"])
        firma = metin_norm(kayit["Firma"])
        tur = metin_norm(kayit["Kaynak Türü"])
        if (kod, firma, tur) in GECERSIZ_PAKET_BIRIM_KURALLARI:
            continue
        kapsam = ondalik(kayit["Fiyatın Kapsadığı Adet"])
        paket = ondalik(kayit["Paket Adedi"], Decimal("1")) or Decimal("1")
        if not kod or not firma or kapsam is None or kapsam <= 0:
            raise ValueError(
                "Paket birim ayarında kod, firma veya kapsadığı adet hatalı: "
                f"{kayit}"
            )
        if tur == "ALIS":
            hedef = sonuc["alis"]
        elif tur == "RAKIP":
            hedef = sonuc["rakip"]
        else:
            raise ValueError(f"Paket birim ayarında geçersiz Kaynak Türü: {tur}")
        hedef[(kod, firma)] = kapsam
        if tur == "ALIS":
            sonuc["alis_marka"][(kod, firma)] = metin_norm(kayit["Marka"])
        sonuc["satirlar"].append({
            **kayit,
            "Paket Adedi": paket,
            "Fiyatın Kapsadığı Adet": kapsam,
        })
    for anahtar, kapsam in DOGRULANMIS_RAKIP_PAKET_KURALLARI.items():
        norm_anahtar = (metin_norm(anahtar[0]), metin_norm(anahtar[1]))
        sonuc["rakip"][norm_anahtar] = kapsam
        if not any(
            metin_norm(satir.get("Ürün Web Servis Kodu", "")) == norm_anahtar[0]
            and metin_norm(satir.get("Firma", "")) == norm_anahtar[1]
            and metin_norm(satir.get("Kaynak Türü", "")) == "RAKIP"
            for satir in sonuc["satirlar"]
        ):
            sonuc["satirlar"].append({
                "Aktif": "EVET",
                "ProductId": "",
                "Marka": "NOAS",
                "Kategori": "Neon Ledler",
                "Ürün Web Servis Kodu": anahtar[0],
                "Firma": "BÜYÜK KALPLER",
                "Kaynak Türü": "RAKİP",
                "Paket Adedi": kapsam,
                "Fiyatın Kapsadığı Adet": kapsam,
                "Kontrol Durumu": "ONAYLANDI",
            })
    return sonuc


def paket_fiyatini_birime_cevir(
    kayit: dict,
    ayarlar: dict,
    kod,
    firma,
    kaynak_turu: str,
    marka="",
) -> dict:
    """Fiyat alanlarını, ayarda belirtilen kapsadığı adede bölerek birim baza getirir."""
    harita = ayarlar.get("paket_birim", {}).get(kaynak_turu, {})
    anahtar = (metin_norm(kod), metin_norm(firma))
    kapsam = harita.get(anahtar)
    if kaynak_turu == "alis":
        beklenen_marka = ayarlar.get("paket_birim", {}).get(
            "alis_marka", {}
        ).get(anahtar, "")
        gelen_marka = metin_norm(marka)
        if beklenen_marka and gelen_marka and beklenen_marka not in gelen_marka:
            return kayit
    if kapsam is None or kapsam <= 1:
        return kayit
    sonuc = dict(kayit)
    fiyat_alanlari = (
        ("liste_tl", "net_tl_kdv_haric")
        if kaynak_turu == "alis"
        else ("satis_fiyati", "net_satis_kdv_haric")
    )
    for alan in fiyat_alanlari:
        deger = ondalik(sonuc.get(alan))
        if deger is not None:
            sonuc[alan] = (deger / kapsam).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
    sonuc["paket_birim_kapsam"] = kapsam
    return sonuc

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
        "Alış İskontosu",
        "Alış İskonto",
        "Alış İskonto %",
        "Alış İskonto Oranı",
        "Alis Iskontosu",
        "Alis Iskonto",
        "Alis Iskonto %",
        "Alis Iskonto Orani",
        "Alış İndirim",
        "Alis Indirim",
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
    "aksiyon_fiyat": [
        "Aksiyon Fiyat",
        "Aksiyon Fiyatı",
        "Aksiyon Fiyati",
        "AksiyonFiyati",
        "Kampanya Fiyat",
        "Kampanya Fiyatı",
        "Kampanya Fiyati",
    ],
    "min_paket": [
        "Min. Paket Adedi",
        "Min Paket Adedi",
        "Minimum Paket Adedi",
        "Paket Adedi",
        "Kutu Adedi",
    ],
    "stok": [
        "Stok",
        "Stok Adedi",
        "Stok Adeti",
        "Stok Miktarı",
        "Stok Miktari",
        "Miktar",
        "Mevcut Stok",
        "Kalan Stok",
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
VARSAYILAN_RAKIP_FIRMALAR = [
    "STAR AKIM",
    "ELEKTROMARKETİM",
    "ELEKTRİK MARKET",
    "AYDINLATMA DÜNYAM",
    "BÜYÜK KALPLER",
    "ÇİMEN ELEKTRİK",
    "ELEKTRİKÇİDEN",
    "ELEKTRİK MALZEMESİ",
    "ELEKTRIX",
    "KOZ AYDINLATMA",
    "KRC ELEKTRİK",
    "TÜMPA ELEKTRİK",
    "ÜÇ FAZ",
    "ZEYBEK MARKET",
]

RAKIP_DOSYA_ESLESMELERI = [
    ("STAR AKIM", ("STAR", "AKIM")),
    ("ELEKTROMARKETİM", ("ELEKTROMARKETIM",)),
    ("ELEKTRİK MARKET", ("ELEKTRIK", "MARKET")),
    ("AYDINLATMA DÜNYAM", ("AYDINLATMADUNYAM",)),
    ("BÜYÜK KALPLER", ("BUYUKKALPLER",)),
    ("ÇİMEN ELEKTRİK", ("CIMENELEKTRIK",)),
    ("ELEKTRİKÇİDEN", ("ELEKTRIKCIDEN",)),
    ("ELEKTRİK MALZEMESİ", ("ELEKTRIKMALZEMESI",)),
    ("ELEKTRIX", ("ELEKTRIX",)),
    ("KOZ AYDINLATMA", ("KOZAYDINLATMA",)),
    ("KRC ELEKTRİK", ("KRCELEKTRIK",)),
    ("TÜMPA ELEKTRİK", ("TUMPAELEKTRIK",)),
    ("ÜÇ FAZ", ("UCFAZ",)),
    ("ZEYBEK MARKET", ("ZEYBEKMARKET",)),
]

ELEKTROMARKETIM_ONAYLI_KOD_ESLESMELERI = {
    "S100NK5SP0": "S100NK",
    "S140NY1": "S140NY",
}

ELEKTROMARKETIM_KOD_MARKA_ONCELIGI = {
    "1002": "NADE",
    "1012": "NEAD",
    "2001": "NEAD",
    "2002": "NEAD",
    "3501": "NEAD",
    "3516": "NEAD",
    "4003": "NEAD",
    "4012": "NEAD",
    "4102": "NEAD",
    "4103": "NEAD",
    "4104": "NEAD",
    "4280": "LEGRAND",
    "4285": "LEGRAND",
    "5021": "NEAD",
    "5022": "NEAD",
    "7001": "NADE",
}

ELEKTRIK_MARKET_CATA_ONAYLI_KOD_ESLESMELERI = {
    "CT-4083": "CT-4083B",
    "CT-4177": "CT-4177B",
    "CT-4178": "CT-4178B",
    "CT-4178 G": "CT-4178G",
    "CT-4540": "CT-4540B",
    "CT-4795-G": "CT-4795G",
    "CT-5170B B": "CT-5170B",
    "CT-5170BG": "CT-5170B",
    "CT-5179S G": "CT-5179SG",
    "CT-5179S B": "CT-5179SB",
    "CT-5270BG": "CT-5270B",
    "CT-5274": "CT-5274B",
    "CT-9080": "CT-9080B",
}

ELEKTRIK_MARKET_JUPITER_ONAYLI_KOD_ESLESMELERI = {
    "LE496 S": "LE496",
    "LE210 N": "LE210",
    "LE208 S": "LE208",
    "LE208 N": "LE208",
    "LE208 B": "LE208",
    "JK120 N": "JK120",
}

ELEKTRIK_MARKET_VIKO_ONAYLI_KOD_ESLESMELERI = {
    "90967008 1": "90967008",
    "90967008MR 1": "90967008",
    "90967056C": "90967056",
    "90979005MR": "90979005",
    "90979004MRD": "90979004",
    "90967013MR": "90967013",
    "90979001MR": "90979001",
    "90967008MR": "90967008",
    "90967009MR": "90967009",
    "90979003MR": "90979003",
}

ELEXTRIX_MARKA_KOD_ONEKLERI = {
    "SCHNEIDER ELECTRIC": {"SCHMT", "SCHMG", "SCH TM", "SCH MG"},
    "SIEMENS": {"SIE"},
    "LEGRAND": {"LEG"},
    "EMAS": {"EMAS"},
    "BEMIS": {"BEM"},
    "WEIDMULLER": {"WM"},
    "TP ELECTRIC": {"TPLS"},
    "BTICINO": {"LEG"},
    "PHOENIX CONTACT": {"PHONX"},
    "SAFAK ELEKTRIK": {"SAFAK"},
    "KRAUS NAIMER": {"KRAUS"},
    "KARDES ELEKTRIK": {"KARDES"},
    "MAVIGARD": {"MAVI"},
    "VIKO": {"VIKO"},
    "BYES": {"BYES"},
    "BOSCH": {"BSH"},
    "ENTES": {"ENT"},
    "ISISO": {"IMR"},
    "TELEMECANIQUE SENSORS": {"SCH TM", "SCH MG"},
}

ELEXTRIX_SIEMENS_ONAYLI_KOD_ESLESMELERI = {
    "SIE 5SL6106 6": "5SL6106 6YA",
    "SIE 5SL6306 7": "5SL6306 7YA",
    "SIE 5SL6310 7": "5SL6310 7YA",
}

TUMPA_ONAYLI_KOD_ESLESMELERI = {
    ("SIEMENS", "5SL6306 7"): "5SL6306 7YA",
    ("SIEMENS", "5SL6310 7"): "5SL6310 7YA",
    ("SIEMENS", "LZS PT78740"): "PT78740",
    ("SIEMENS", "LZX PT370024"): "PT370024",
    ("SIEMENS", "LZX PT370730"): "PT370730",
}

ELEKTRIK_MALZEMESI_SCHNEIDER_ONAYLI_KOD_ESLESMELERI = {
    "LC1D258BD": "LC1D25BD",
    "XALK178E": "XALK178",
    "ABLS1A24050E": "ABLS1A24050",
    "ABLS1A24100E": "ABLS1A24100",
    "ATV340U22N4E": "ATV340U22N4",
    "ATV340U75N4E": "ATV340U75N4",
    "LC1D256M7": "LC1D25M7",
    "A9A26904": "A9N26904",
    "ZBY2101": "ZBYF2101",
}

ELEKTRIK_MALZEMESI_SIEMENS_ONAYLI_KOD_ESLESMELERI = {
    "5SL6306 7": "5SL6306 7YA",
    "5SL6310 7": "5SL6310 7YA",
    "LZX PT370024": "PT370024",
}

ZEYBEK_MARKET_ONAYLI_KOD_ESLESMELERI = {
    ("NEAD", "NEAD-4003"): "4003",
    ("NEAD", "NEAD-4001"): "4001",
    ("NEAD", "NEAD-4103"): "4103",
    ("NEAD", "NEAD-2001-MP"): "2001-MP",
    ("NEAD", "NEAD-2001"): "2001",
    ("NEAD", "NEAD-7450-B"): "7450-B",
    ("NEAD", "NEAD-4012"): "4012",
    ("NEAD", "NEAD-1005-MP"): "1005-MP",
    ("NEAD", "NEAD-7444-B"): "7444-B",
    ("NEAD", "NEAD-7446-B"): "7446-B",
    ("NEAD", "NEAD-7448-B"): "7448-B",
    ("NEAD", "NEAD-1012"): "1012",
    ("ORTAÇ", "ORTAÇ-001"): "ORG01",
    ("ORTAÇ", "ORTAÇ-002"): "ORG02",
    ("ORTAÇ", "ORTAÇ-003"): "ORG03",
    ("ORTAÇ", "ORTAÇ-004"): "ORG04",
    ("ORTAÇ", "ORTAÇ-005"): "ORG05",
    ("ORTAÇ", "ORTAÇ-006"): "ORG06",
    ("ORTAÇ", "ORTAÇ-008"): "ORG08",
    ("ORTAÇ", "ORTAÇ-009"): "ORG09",
    ("WAGO", "WAGO 2273-202"): "2273-202",
    ("WAGO", "WAGO 221-412"): "221-412",
    ("WAGO", "WAGO 2273-203"): "2273-203",
    ("WAGO", "WAGO 221-413"): "221-413",
    ("WAGO", "WAGO 2273-204"): "2273-204",
    ("WAGO", "WAGO 2273-205"): "2273-205",
    ("WAGO", "WAGO 221-415"): "221-415",
    ("ENTES", "ENTES FG-4"): "M1454",
    ("ENTES", "ENTES SER"): "M4091",
}

ZEYBEK_TSOFT_MARKA_ADLARI = {
    "CATA": ("CATA",),
    "EMAS": ("EMAS",),
    "SIEMENS": ("SIEMENS",),
    "BEMIS": ("BEMIS", "BYES"),
    "VIKO": ("VIKO",),
    "ENTES": ("ENTES",),
    "OPAS": ("OPAS",),
}

UCFAZ_TSOFT_MARKA_ADLARI = {
    "ACK": ("ACK",),
    "AUDIO": ("AUDIO",),
    "CATA": ("CATA",),
    "ÇETİNKAYA": ("ÇETİNKAYA",),
    "ENTES": ("ENTES",),
    "ERSE": ("ERSE",),
    "JUPITER": ("JUPITER",),
    "KLEMSAN": ("KLEMSAN",),
    "MEAN WELL": ("MEAN WELL", "MEANWELL"),
    "METE ENERJİ": ("METE ENERJİ", "METE"),
    "MUTLUSAN": ("MUTLUSAN",),
    "NADE": ("NADE", "NEAD"),
    "SCHNEIDER ELECTRIC": ("SCHNEIDER",),
    "SIEMENS": ("SIEMENS",),
    "VIKO": ("VIKO", "ARTLINE"),
    "WAGO": ("WAGO",),
}

KRC_TSOFT_MARKA_ADLARI = {
    "3M": ("3M",),
    "ABB": ("ABB",),
    "ALCU": ("ALCU",),
    "AUDIO": ("AUDIO",),
    "AUTONICS": ("AUTONICS",),
    "BEMİS": ("BEMIS", "BYES"),
    "BEMİS E-V CHARGE": ("BEMIS", "BYES"),
    "BYES": ("BYES", "BEMIS"),
    "CATA": ("CATA",),
    "CETA FORM": ("CETA FORM",),
    "CHINT": ("CHINT",),
    "CKS": ("CKS",),
    "ÇETİNKAYA PANO": ("ÇETINKAYA",),
    "DÜNDAR": ("DÜNDAR",),
    "EATON": ("EATON",),
    "EFEN": ("EFEN",),
    "EMAS": ("EMAS",),
    "ENDA": ("ENDA",),
    "ENTES": ("ENTES",),
    "ERSE": ("ERSE",),
    "FINDER": ("FINDER",),
    "FLENDER": ("FLENDER",),
    "GECEM": ("GECEM",),
    "GOYA": ("GOYA",),
    "GRUP ARGE": ("GRUP ARGE",),
    "GWEST": ("GWEST",),
    "HELLERMANN TYTON": ("HELLERMANN", "TYTON"),
    "HENSEL": ("HENSEL",),
    "HYUNDAI ELEKTRİK": ("HYUNDAI",),
    "IBOCO": ("IBOCO",),
    "INNOMOTICS": ("INNOMOTICS",),
    "İNOLED": ("İNOLED",),
    "KARDEŞ ELEKTRİK": ("KARDEŞ",),
    "KLEMSAN": ("KLEMSAN",),
    "KLAUKE": ("KLAUKE",),
    "KRAUS & NAIMER": ("KRAUS", "NAIMER"),
    "LEGRAND": ("LEGRAND",),
    "LOOBAR": ("LOOBAR",),
    "LOSPAR": ("LOSPAR",),
    "LS ELEKTRİK": ("LS ELEKTRİK", "LS ELECTRIC"),
    "MAVİLİ": ("MAVİLİ",),
    "MEANWELL": ("MEAN WELL", "MEANWELL"),
    "METE ENERJİ": ("METE ENERJİ", "METE"),
    "METOP": ("METOP",),
    "MUCCO": ("MUCCO",),
    "MUTLUSAN": ("MUTLUSAN",),
    "NETELSAN": ("NETELSAN",),
    "OBO BETTERMANN": ("OBO", "BETTERMANN"),
    "OPAS": ("OPAS",),
    "ORTAÇ": ("ORTAÇ",),
    "ÖZNUR KABLO": ("ÖZNUR",),
    "PANASONİC": ("PANASONIC",),
    "PELSAN": ("PELSAN",),
    "PHILIPS": ("PHILIPS",),
    "PHOENIX CONTACT": ("PHOENIX CONTACT",),
    "RADE KONCAR": ("RADE", "KONCAR"),
    "REÇBER KABLO": ("REÇBER",),
    "RELPOL": ("RELPOL",),
    "RITTAL": ("RITTAL",),
    "SCHNEIDER": ("SCHNEIDER",),
    "SCHRACK": ("SCHRACK",),
    "SHREEM": ("SHREEM",),
    "SIEMENS": ("SIEMENS",),
    "SIEMENS HIZ KONTROL": ("SIEMENS",),
    "SIEMENS OTOMASYON": ("SIEMENS",),
    "SİGMA": ("SİGMA",),
    "TAM GRUP": ("TAM GRUP",),
    "TEKPAN": ("TEKPAN",),
    "TELEMECANIQUE SENSORS": ("TELEMECANIQUE",),
    "TORK": ("TORK",),
    "TP ELEKTRİK": ("TP ELECTRIC", "TP ELEKTRİK"),
    "TUNÇMATİK": ("TUNÇMATİK",),
    "TURKUAZ": ("TURKUAZ",),
    "TYITS": ("TYITS",),
    "UNI-T": ("UNI-T", "UNI T"),
    "VARTA": ("VARTA",),
    "VİKO": ("VİKO", "VIKO"),
    "WEIDMULLER": ("WEIDMULLER", "WEIDMÜLLER"),
    "ZAZU": ("ZAZU",),
    "ZİGAR": ("ZİGAR",),
}

BUYUK_KALPLER_TSOFT_MARKA_ADLARI = {
    **KRC_TSOFT_MARKA_ADLARI,
    "ACK": ("ACK",),
    "AUDIO": ("AUDIO",),
    "BVN BAHCIVAN": ("BVN", "BAHCIVAN", "BAHÇIVAN"),
    "HOROZ ELECTRIC": ("HOROZ",),
    "INOLED": ("INOLED", "İNOLED"),
    "LEDVANCE": ("LEDVANCE", "OSRAM"),
    "MITSUBISHI ELECTRIC": ("MITSUBISHI", "MITSUBISHI ELECTRIC"),
    "NOAS": ("NOAS",),
    "OSRAM": ("OSRAM", "LEDVANCE"),
    "OZNUR KABLO": ("OZNUR", "ÖZNUR"),
    "SUNLIGHT": ("SUNLIGHT",),
    "TUNCMATIK": ("TUNCMATIK", "TUNÇMATİK"),
    "VOLTEK": ("VOLTEK",),
}

UCFAZ_ERSE_ONAYLI_KOD_ESLESMELERI = {
    "2X0 75 ERVITAL LIHCH FE180 PH120": "MDER009800200075O",
    "2X0 75 LIYCY": "MDER005700200075G",
    "2X1 5 ERVITAL LIHCH FE180 PH120": "MDER009800200150O",
    "2X1 5 LIYCY": "MDER005700200150G",
    "2X1 ERVITAL LIHCH FE180 PH120": "MDER009800200100O",
    "3X0 75 ERVITAL LIHCH FE180 PH120": "MDER009800300075O",
    "3X1 5 ERVITAL LIHCH FE180 PH120": "MDER009800300150O",
    "3X1 5 LIYCY": "MDER005700300150G",
    "3X1 ERVITAL LIHCH FE180 PH120": "MDER009800300100O",
    "4X0 75 ERVITAL LIHCH FE180 PH120": "MDER009800400075O",
    "4X1 5 ERVITAL LIHCH FE180 PH120": "MDER009800400150O",
    "4X1 ERVITAL LIHCH FE180 PH120": "MDER009800400100O",
    "6X1 5 LIYCY": "MDER005700600150G",
    "6X1 LIYCY": "MDER005700600100G",
    "1X2X0 80 0 80 JE H ST H BD": "MDER009200100080O",
    "1X2X1 0 80 JE H ST H BD": "MDER009200100102O",
    "2X2X0 80 0 80 JE H ST H BD": "MDER009200200080O",
    "2X2X1 0 80 JE H ST H BD": "MDER009200200102O",
}

UCFAZ_VIKO_ONAYLI_KOD_ESLESMELERI = {
    "90960013": "90967013",
    "92192551": "92192251",
    "92192552": "92192252",
}

UCFAZ_WAGO_ONAYLI_KOD_ESLESMELERI = {
    "24VDC2X8A": "5119 9793",
}

UCFAZ_MEANWELL_ONAYLI_KOD_ESLESMELERI = {
    "LRS 200 12": "LRS 200 12KO",
    "LRS 200 24": "LRS 200 24KO",
}

UCFAZ_CATA_ONAYLI_KOD_ESLESMELERI = {
    "CT 5252": "CT 5252B",
    "CT 4217G": "CT 4217",
    "CT 4329 64": "CT 4329",
}

ELEKTRIKCIDEN_ENTES_MODEL_ESLESMELERI = {
    "DTR-10T": "M5307",
    "FG-4": "M1454",
    "MKS-01": "M1131",
    "RG3-12E": "M3682",
    "SSRC-04": "M1656",
    "MCB-30": "M4114",
    "MCB-60": "M4115",
}

ELEKTRIKCIDEN_VIKO_ONAYLI_AD_ESLESMELERI = {
    "VIKO KARRE BESLI YATAY CERCEVE BEYAZ": "90960204",
    "VIKO KARRE IKILI YATAY CERCEVE BEYAZ": "90960201",
    "VIKO KARRE ALTILI YATAY CERCEVE BEYAZ": "90960205",
    "VIKO KARRE DORTLU YATAY CERCEVE BEYAZ": "90960203",
    "VIKO KARRE SIVAALTI 2 LI PRIZ CERCEVESI BEYAZ": "90960206",
    "VIKO KARRE TEKLI CERCEVE BEYAZ": "90960200",
    "VIKO KARRE UCLU YATAY CERCEVE BEYAZ": "90960202",
    "VIKO KARRE MERIDIAN 2 LI JALUZI KUMANDA MEKANIZMA BEYAZ CERCEVE HARIC": "90967016",
    "VIKO KARRE MERIDIAN 3 LU ANAHTAR MEKANIZMA BEYAZ CERCEVE HARIC": "90967068",
    "VIKO KARRE MERIDIAN ANAHTAR MEKANIZMA BEYAZ CERCEVE HARIC": "90967001",
    "VIKO KARRE MERIDIAN ARA VAVIEN PERMUTATOR MEKANIZMA BEYAZ CERCEVE HARIC": "90967031",
    "VIKO KARRE MERIDIAN DATA PRIZI JAKSIZ MEKANIZMA BEYAZ CERCEVE HARIC": "90967062",
    "VIKO KARRE MERIDIAN DIMMER 600W MEKANIZMA BEYAZ CERCEVE HARIC": "90967047",
    "VIKO KARRE MERIDIAN KOMITATOR MEKANIZMA BEYAZ CERCEVE HARIC": "90967002",
    "VIKO KARRE MERIDIAN KOMITATOR VAVIEN MEKANIZMA BEYAZ CERCEVE HARIC": "90967017",
    "VIKO KARRE MERIDIAN LIGHT MEKANIZMA BEYAZ CERCEVE HARIC": "90967003",
    "VIKO KARRE MERIDIAN TOPRAKLI PRIZ MEKANIZMA BEYAZ CERCEVE HARIC": "90967008",
    "VIKO KARRE MERIDIAN UYDU PRIZI F KON SONLU MEKANIZMA BEYAZ CERCEVE HARIC": "90967009",
    "VIKO KARRE MERIDIAN VAVIEN MEKANIZMA BEYAZ CERCEVE HARIC": "90967004",
    "VIKO KARRE MERIDIAN COCUK KORUMALI TOP PRIZ MEKANIZMA BEYAZ CERCEVE HARIC": "90967042",
    "VIKO MERIDIAN ALTILI YATAY CERCEVE BEYAZ": "90979006",
    "VIKO MERIDIAN BESLI YATAY CERCEVE BEYAZ": "90979005",
    "VIKO MERIDIAN DORTLU YATAY CERCEVE BEYAZ": "90979004",
    "VIKO MERIDIAN IKILI YATAY CERCEVE BEYAZ": "90979002",
    "VIKO MERIDIAN TEKLI CERCEVE BEYAZ": "90979001",
    "VIKO MERIDIAN UCLU YATAY CERCEVE BEYAZ": "90979003",
    "VIKO MIMOZA TOPRAKSIZ ERKEK FIS BEYAZ": "90305400",
    "VIKO TOPRAKLI DISI FIS": "90301800",
    "VIKO TOPRAKLI ERKEK FIS BEYAZ": "90304100",
}

ELEKTRIKCIDEN_CATA_ONAYLI_AD_ESLESMELERI = {
    "CATA 1 5W SATIN KASA YILDIZ LED SPOT 3000K GUNISIGI CT 5270": "CT 5270SG",
    "CATA 1 5W SATIN KASA YILDIZ LED SPOT 6500K BEYAZ CT 5270": "CT 5270SB",
    "CATA 1 5W SIYAH KASA YILDIZ LED SPOT 3200K GUNISIGI CT 5270": "CT 5270SYHG",
    "CATA 1 5W SIYAH KASA YILDIZ LED SPOT 6500K BEYAZ CT 5270": "CT 5270SYHB",
    "CATA 1 5W YUVARLAK SENSORLU KORIDOR ARMATURU GUNISIGI 3200K CT 5174": "CT 5174G",
    "CATA 10W T5 LED BANT ARMATURU 60CM CT 2467 BEYAZ": "CT 2467B",
    "CATA 12V DIS MEKAN NEON LED 50M YESIL CT 4455BY": "CT 4455Y",
    "CATA 13W TUVALU YATAK BASI APLIK SIYAH KASA 3200K CT 5253S": "CT 5253",
    "CATA 15W T5 LED BANT ARMATURU 90CM CT 2468 BEYAZ": "CT 2468B",
    "CATA 1W SENSORLU FUME KORIDOR ARMATURU BEYAZ 6500K CT 5170F": "CT 5170FB",
    "CATA 1W SENSORLU FUME KORIDOR ARMATURU GUNISIGI 3000K CT 5170F": "CT 5170FG",
    "CATA 1W SENSORLU SIYAH KORIDOR ARMATURU BEYAZ 6500K CT 5170S": "CT 5170SB",
    "CATA 1W SENSORLU SIYAH KORIDOR ARMATURU GUNISIGI 3000K CT 5170S": "CT 5170SG",
    "CATA 20W PLATINIUM LED PROJEKTOR 6500K BEYAZ CT 4656B": "CT 4656",
    "CATA 20W T5 LED BANT ARMATURU 120CM CT 2469 BEYAZ": "CT 2469B",
    "CATA 25W TORCH LED AMPUL E27 DUY 6400K BEYAZ CT 4330B": "CT 4330",
    "CATA 30W BABIL RAY SPOT SIYAH KASA 4000K CT 5334": "CT 5334GG",
    "CATA 30W LED PERVANE AMPUL E27 DUY 6500K BEYAZ CT 1151": "CT 1151B",
    "CATA 3W BAYKUS YATAK BASI APLIK SIYAH KASA 3200K CT 5232S": "CT 5232",
    "CATA 3W SIVAALTI PANEL LED 6500K BEYAZ CT 5144": "CT 5144B",
    "CATA 40W ETANJ ACIL AYD LED ARMATUR 120CM 6500K BEYAZ CT 2480B": "CT 2480",
    "CATA 40W ETANJ DIS MEKAN LED ARMATUR120CM 6500K BEYAZ CT 2477B": "CT 2477",
    "CATA 45W TORCH LED AMPUL E27 DUY 6400K BEYAZ CT 4242B": "CT 4242",
    "CATA 4W EDISON LED AMPUL GUNISIGI CT 4288G": "CT 4288",
    "CATA 4W BUJI RUSTIK KIVRIK LED AMPUL E 14 DUY 3200K CT 4062G": "CT 4062",
    "CATA 4W BUJI RUSTIK LED AMPUL E 14 DUY 3200K CT 4066G": "CT 4066",
    "CATA 55W TORCH LED AMPUL E27 DUY 6400K BEYAZ CT 4262B": "CT 4262",
    "CATA 5W T5 LED BANT ARMATURU 30CM CT 2466 BEYAZ": "CT 2466B",
    "CATA 60X60CM 60W SIVAUSTU PANEL LED BEYAZ CT 5281B": "CT 5281",
    "CATA 65W TORCH LED AMPUL E27 DUY 6400K BEYAZ CT 4328B": "CT 4328",
    "CATA 6W DEKORATIF EDISON RUSTIK LED AMPUL GUNISIGI CT 4231G": "CT 4231",
    "CATA 7W AKIK COB LED ARMATUR 6500K BEYAZ CT 5252": "CT 5252B",
    "CATA 80W KRISTAL LED BANT ARMATURU 120CM BEYAZ CT 2482B": "CT 2482",
    "CATA 8W DEKORATIF EDISON RUSTIK LED AMPUL BEYAZ CT 4353B": "CT 4353",
    "CATA 8W LED CANAK AMPUL GU10 DUY 4000K CT 4215": "CT 4215GG",
    "CATA 9W 60CM LED FLORASAN AMPUL 6500K BEYAZ CT 4223B": "CT 4223",
    "CATA 9W SARJLI LED AMPUL BEYAZ CT 4229B": "CT 4229",
    "CATA IC DIS MEKAN YAPISKANLI 220V SERIT LED 1M BEYAZ CT 4595B": "CT 4595",
    "CATA KABLOSUZ SARJLI SENSORLU MIKNATISLI KABIN AYDINLATMA CT 5180G": "CT 5180",
    "CATA OKYANUS 6W CIFT RENKLI SIVAALTI PANEL LED BEYAZ MAVI CT 5150B": "CT 5150",
    "CATA 1 5W SIYAH KASA UFO YILDIZ LED SPOT 6500K BEYAZ CT 5268SB": "CT 5268SB",
}

STAR_AKIM_ONAYLI_KOD_ESLESMELERI = {
    ("ACK", "ACK-AS90-00800F"): "AS90-00800",
    ("ACK", "ACK-AS90-00803F"): "AS90-00803",
    ("ACK", "ACK-AY02-00600-PRO"): "AY02-00600",
    ("AUDIO", "AUDIO 001180C"): "AUDIO 001180",
    ("BTICINO", "RG4001P"): "RG4001",
    ("BTICINO", "RG4003P"): "RG4003",
    ("BTICINO", "RG4004P"): "RG4004",
    ("BTICINO", "RG4005P"): "RG4005",
    ("BTICINO", "RG4141P"): "RG4141",
    ("BTICINO", "RW4005P"): "RW4005",
    ("CATA", "CT-2477B"): "CT-2477",
    ("CATA", "CT-2479B"): "CT-2479",
    ("CATA", "CT-4058B"): "CT-4058",
    ("CATA", "CT-4062G"): "CT-4062",
    ("CATA", "CT-4066G"): "CT-4066",
    ("CATA", "CT-4223B"): "CT-4223",
    ("CATA", "CT-4224B"): "CT-4224",
    ("CATA", "CT-4229B"): "CT-4229",
    ("CATA", "CT-4242B"): "CT-4242",
    ("CATA", "CT-4263B"): "CT-4263",
    ("CATA", "CT-4278B"): "CT-4278",
    ("CATA", "CT-4328B"): "CT-4328",
    ("CATA", "CT-4329B"): "CT-4329",
    ("CATA", "CT-4330B"): "CT-4330",
    ("CATA", "CT-4648-100W"): "CT-4648",
    ("CATA", "CT-4664B"): "CT-4664",
    ("CATA", "CT-5110B"): "CT-5110",
    ("CATA", "CT-5170 (SYH) GÜNIŞIĞI"): "CT-5170SG",
    ("CATA", "CT-5170FM"): "CT-5170FG",
    ("CATA", "CT-5179 B"): "CT-5179SB",
    ("CATA", "CT-5179G SYH"): "CT-5179SG",
    ("CATA", "CT-5213S"): "CT-5213",
    ("CATA", "CT-5220B"): "CT-5220",
    ("CATA", "CT-5252 B"): "CT-5252B",
    ("CATA", "CT-5252 G"): "CT-5252G",
    ("CATA", "CT-5253S"): "CT-5253",
    ("CATA", "CT-5268G SYH"): "CT-5268G",
    ("CATA", "CT-5270B SYH"): "CT-5270SYHB",
    ("CATA", "CT-5270G SYH"): "CT-5270SYHG",
    ("CATA", "CT-5271G-ALÜMİNYUM-"): "CT-5271G",
    ("CATA", "CT-5272B-ALÜMİNYUM-"): "CT-5272B",
    ("CATA", "CT-5273B-ALÜMİNYUM-"): "CT-5273B",
    ("CATA", "CT-5273G-ALÜMİNYUM"): "CT-5273G",
    ("CATA", "CT-5274G-ALÜMİNYUM"): "CT-5274G",
    ("CATA", "CT-5280B"): "CT-5280",
    ("CATA", "CT-5281B"): "CT-5281",
    ("CATA", "CT-5420"): "CT-5420SB",
    ("CATA", "CT-5421 SB"): "CT-5421SB",
    ("CATA", "CT-5421 SP"): "CT-5421SP",
    ("CATA", "CT-5820G"): "CT-5820",
    ("CATA", "CT-5825G"): "CT-5825",
    ("CATA", "CT-5841G"): "CT-5841",
    ("CATA", "CT-9081 4000K"): "CT-9081K",
    ("JUPITER", "JH639 SYH"): "JH639 S",
    ("JUPITER", "JH640SN"): "JH640 SN",
    ("JUPITER", "LD511 N"): "LD511",
    ("JUPITER", "LD469S SİYAH"): "LD469 S",
    ("JUPITER", "LD471 S SYH"): "LD471 S",
    ("JUPITER", "LD472 S SYH"): "LD472 S",
    ("JUPITER", "LD474 S 830"): "LD474 S",
    ("JUPITER", "LE206 S"): "LE206",
    ("JUPITER", "LE494 S"): "LE494",
    ("JUPITER", "LG984 S"): "LG984",
    ("JUPITER", "LW427 S"): "LW427",
    ("JUPITER", "LD513 N"): "LD513",
    ("LEGRAND", "433708"): "403708",
    ("NADE", "NADE FD 3050"): "FD3050",
    ("NADE", "NADE FD 7130"): "FD7130",
    ("NOAS", "YL10-0901-B"): "YL10-0901",
    ("SIEMENS", "SIE.001.030"): "5SL6110-6",
    ("SIEMENS", "SIE.001.036"): "5SL6102-7",
    ("SIEMENS", "SIE.001.037"): "5SL6103-7",
    ("SIEMENS", "SIE.001.038"): "5SL6104-7",
    ("SIEMENS", "SIE.001.045"): "5SL6140-7",
    ("SIEMENS", "SIE.001.047"): "5SL6163-7",
    ("SIEMENS", "SIE.001.050"): "5SL6316-7",
    ("SIEMENS", "SIE.001.057"): "5SP4380-7",
    ("SIEMENS", "SIE.001.059"): "5SP4392-7",
    ("SIEMENS", "SIE.001.078"): "5SL6206-7",
    ("SIEMENS", "SIE.001.099"): "5SL6506-7",
    ("SIEMENS", "SIE.001.123"): "5SL6650-7",
    ("SIEMENS", "SIE.001.136"): "5SL4350-7",
    ("SIEMENS", "SIE.001.158"): "5SL6310-7YA",
    ("SIEMENS", "SIE.001.170"): "5SL6516-6YA",
    ("SIEMENS", "SIE.002.050"): "5SV5614-6",
    ("SIEMENS", "SIE.003.034"): "3RT2015-1AP01",
    ("SIEMENS", "SIE.003.044"): "3RT2025-1AP00",
    ("SIEMENS", "SIE.003.046"): "3RT2027-1AP00",
    ("SIEMENS", "SIE.003.071"): "3RH2911-1HA11",
    ("SIEMENS", "SIE.003.083"): "3RH2911-1LA11",
    ("SIEMENS", "SIE.004.055"): "3RU2116-1HB0",
    ("SIEMENS", "SIE.004.057"): "3RU2116-1KB0",
    ("SIEMENS", "SIE.004.058"): "3RU2116-4AB0",
    ("SIEMENS", "SIE.004.079"): "3RU2116-1CB1",
    ("SIEMENS", "SIE.004.081"): "3RU2116-0JB1",
    ("SIEMENS", "SIE.004.085"): "3RU2126-4BB0",
    ("SIEMENS", "SIE.005.055"): "3RV2011-1HA10",
    ("SIEMENS", "SIE.005.058"): "3RV2011-1AA10",
    ("SIEMENS", "SIE.005.060"): "3RV2011-1DA10",
    ("SIEMENS", "SIE.006.055"): "3VM1110-3EE32-0AA0",
    ("SIEMENS", "SIE.012.001"): "3TY7561-1AA00",
    ("VIKO", "10VTB-4C50N"): "10VTB-4C50",
    ("VIKO", "10VTB-4C63N"): "10VTB-4C63",
    ("VIKO", "10VTR4-10030/A"): "VTR4-10030",
    ("VIKO", "4VTB-1C16N"): "4VTB-1C16",
    ("VIKO", "4VTB-1C40N"): "4VTB-1C40",
    ("VIKO", "VTC-18/11/S-SN2"): "VTC-18/11/S",
    ("VIKO ARTLINE", "92615158-C1"): "92615158",
}


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


def tsoft_aktiflik_durumu(deger) -> str:
    norm = metin_norm(deger)
    if norm in {"1", "TRUE", "ACTIVE", "AKTIF", "EVET"}:
        return "AKTİF"
    if norm in {"0", "FALSE", "INACTIVE", "PASIF", "HAYIR"}:
        return "PASİF"
    return "BİLİNMİYOR"


def zeybek_kod_anlamli_mi(kod: str) -> bool:
    kompakt = re.sub(r"[^A-Z0-9]", "", kod)
    if len(kompakt) < 4:
        return False
    return not (kompakt.isdigit() and len(kompakt) < 6)


def zeybek_ifade_iceriyor(metin: str, ifade: str) -> bool:
    return bool(ifade and f" {ifade} " in f" {metin} ")


def zeybek_market_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_marka_urunleri: dict[str, list[tuple[str, str]]],
) -> str:
    """Zeybek iç kodunu, marka içinde tekil doğrulanan T-Soft koduna çevirir."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    kaynak_urun_adi = metin_norm(urun_adi)
    hedef_urunler = tsoft_marka_urunleri.get(marka_norm, [])
    if not kaynak_kod or not hedef_urunler:
        return ""

    kaynak_model = kaynak_kod
    if kaynak_model.startswith(f"{marka_norm} "):
        kaynak_model = kaynak_model[len(marka_norm) + 1 :].strip()
    ek_onek = {
        metin_norm("SIEMENS"): "S ",
        metin_norm("VIKO"): "VK ",
        metin_norm("BEMIS"): "BEM ",
    }.get(marka_norm, "")
    if ek_onek and kaynak_model.startswith(ek_onek):
        kaynak_model = kaynak_model[len(ek_onek) :].strip()
    if marka_norm == metin_norm("OPAS"):
        kaynak_model = re.sub(
            r"^(?:YPT|YKL|LB2S)\s*",
            "",
            kaynak_model,
        ).strip()

    hedef_kodlari = {hedef_kod for hedef_kod, _ in hedef_urunler}
    if kaynak_model in hedef_kodlari:
        return kaynak_model

    adaylar = set()
    for hedef_kod, hedef_urun_adi in hedef_urunler:
        if (
            zeybek_kod_anlamli_mi(hedef_kod)
            and zeybek_ifade_iceriyor(kaynak_urun_adi, hedef_kod)
        ):
            adaylar.add(hedef_kod)
        if (
            zeybek_kod_anlamli_mi(kaynak_model)
            and zeybek_ifade_iceriyor(hedef_urun_adi, kaynak_model)
        ):
            adaylar.add(hedef_kod)
    return next(iter(adaylar)) if len(adaylar) == 1 else ""


def normlanmis_kod_ifadesi_iceriyor(
    metin_normlanmis: str,
    kod_normlanmis: str,
) -> bool:
    if not metin_normlanmis or not kod_normlanmis:
        return False
    son_parca = kod_normlanmis.split()[-1]
    bitis = r"(?=$| [A-Z])" if son_parca.isdigit() else r"(?=$| )"
    return bool(
        re.search(
            rf"(?:^| ){re.escape(kod_normlanmis)}{bitis}",
            metin_normlanmis,
        )
    )


def kod_ifadesi_iceriyor(metin: str, kod: str) -> bool:
    """Normalize kodu tam ifade olarak bulur; sayısal devamı kısmi eşleşme saymaz."""
    return normlanmis_kod_ifadesi_iceriyor(metin_norm(metin), metin_norm(kod))


def ham_metin_model_norm(deger) -> str:
    """Model kontrolü için ayraçları koruyarak metni büyük harfe dönüştürür."""
    metin = "" if pd.isna(deger) else str(deger).strip().upper()
    metin = metin.translate(
        str.maketrans(
            {
                "Ç": "C",
                "Ğ": "G",
                "İ": "I",
                "Ö": "O",
                "Ş": "S",
                "Ü": "U",
                "ı": "I",
            }
        )
    )
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(c for c in metin if not unicodedata.combining(c))


def ham_metin_model_kodu_iceriyor(metin: str, kod: str) -> bool:
    """Ayraç biçimi değişse de tam modeli bulur; ondalık/varyant devamını reddeder."""
    metin_ham = ham_metin_model_norm(metin)
    kod_parcalari = re.findall(r"[A-Z0-9]+", ham_metin_model_norm(kod))
    if not metin_ham or not kod_parcalari:
        return False
    desen = r"(?<![A-Z0-9])" + r"[^A-Z0-9]+".join(
        re.escape(parca) for parca in kod_parcalari
    ) + r"(?![A-Z0-9])"
    for eslesme in re.finditer(desen, metin_ham):
        devam = metin_ham[eslesme.end() :]
        if re.match(r"^\s*[.,/-]\s*\d", devam):
            continue
        return True
    return False


AYDINLATMA_DUNYAM_KOD_ONEKLERI = {
    "AUDIO": ("AUDIO ",),
    "GWEST": ("GWEST ",),
    "LEGRAND": ("LEG ",),
    "MUTLUSAN": ("MTS ",),
    "SCHNEIDER": ("SCH ",),
    "VIKO": ("VIKO ",),
}

AYDINLATMA_DUNYAM_ONAYLI_KOD_ESLESMELERI = {
    ("AUDIO", "AUDIO 001230"): "1230",
    ("AUDIO", "AUDIO 001328"): "1328",
    ("AUDIO", "AUDIO 001361"): "1361",
    ("AUDIO", "AUDIO 002268"): "2268",
    ("AUDIO", "AUDIO 004849 E18"): "AUDIO 004849 E08",
    ("CATA", "CT 105B"): "CT 105",
    ("CATA", "CT 2477B"): "CT 2477",
    ("CATA", "CT 2482B"): "CT 2482",
    ("CATA", "CT 4058B"): "CT 4058",
    ("CATA", "CT 4064G"): "CT 4064",
    ("CATA", "CT 4229B"): "CT 4229",
    ("CATA", "CT 4242B"): "CT 4242",
    ("CATA", "CT 4328B"): "CT 4328",
    ("CATA", "CT 4329B"): "CT 4329",
    ("CATA", "CT 4656B"): "CT 4656",
    ("CATA", "CT 5150B"): "CT 5150",
    ("CATA", "CT 5212KG"): "CT 5212G",
    ("CATA", "CT 5253S"): "CT 5253",
    ("CATA", "CT 5421"): "CT 5421SB",
    ("CATA", "CT 5816G"): "CT 5816",
    ("CATA", "CT 5821G"): "CT 5821",
    ("CATA", "CT 8010G"): "CT 8010",
    ("CATA", "CT 9500A"): "CT 9500",
    ("CATA", "CT 9500B"): "CT 9500",
    ("CATA", "CT 9500M"): "CT 9500",
    ("CATA", "CT 9500P"): "CT 9500",
    ("JUPITER", "LD472B"): "LD472 B",
    ("JUPITER", "LD474B"): "LD474 B",
}

CIMEN_ELEKTRIK_CATA_ONAYLI_KOD_ESLESMELERI = {
    "CT 1151": "CT 1151B",
    "CT 2582 24": "CT 2582",
    "CT 4050 KAMERA": "CT 4050",
    "CT 4062G": "CT 4062",
    "CT 4066G": "CT 4066",
    "CT 4215 ILIK BEYAZ": "CT 4215GG",
    "CT 4229B": "CT 4229",
    "CT 4242B": "CT 4242",
    "CT 4329B": "CT 4329",
    "CT 4656B": "CT 4656",
    "CT 4657B": "CT 4657",
    "CT 5150B": "CT 5150",
    "CT 5170G SK": "CT 5170SG",
    "CT 5170B SK": "CT 5170SB",
    "CT 5180G": "CT 5180",
    "CT 5212 EB Y": "CT 5212EB",
    "CT 5212 KG Y": "CT 5212G",
    "CT 5212 KB Y": "CT 5212B",
    "CT 5212 SB Y": "CT 5212SB",
    "CT 5232G SK": "CT 5232",
    "CT 5233G BK": "CT 5233G",
    "CT 5234B BK": "CT 5234B",
    "CT 5250G B": "CT 5250",
    "CT 5253S": "CT 5253",
    "CT 5258R": "CT 5258",
    "CT 5270SK": "CT 5270SB",
    "CT 5270SK B": "CT 5270SYHB",
    "CT 5273B BK": "CT 5273B",
    "CT 5274B BK": "CT 5274B",
    "CT 5810G": "CT 5810",
    "CT 5811G": "CT 5811",
    "CT 5812G": "CT 5812",
    "CT 5815G": "CT 5815",
    "CT 5816G": "CT 5816",
    "CT 5820G": "CT 5820",
    "CT 5822G": "CT 5822",
    "CT 5825G": "CT 5825",
    "CT 7321RGB": "CT 7321",
    "CT 8010G": "CT 8010",
    "CT 8438G": "CT 8438",
    "CT 8439 K": "CT 8439",
    "CT 8440 BK Y": "CT 8440",
    "CT 9082 G": "CT 9082G",
    "CT 9152 10": "CT 9152",
    "CT 9500A": "CT 9500",
    "CT 9500B": "CT 9500",
    "CT 9500M": "CT 9500",
    "CT 9500P": "CT 9500",
}

CIMEN_ELEKTRIK_NOAS_ONAYLI_KOD_ESLESMELERI = {
    "YL29 3002 BG": "YL29 3002",
    "YL29 4001B": "YL29 4001",
}

CIMEN_ELEKTRIK_ISILDAR_ONAYLI_KOD_ESLESMELERI = {
    "140000006102013": "14 000 000 610 2013",
}

KOZ_AYDINLATMA_CATA_ONAYLI_KOD_ESLESMELERI = {
    "CT 104GO": "CT 104",
    "CT 104G": "CT 104",
    "CT 104S": "CT 104",
    "CT 104GS": "CT 104",
    "CT 105 BEYAZ": "CT 105",
    "CT 105 GUMUS": "CT 105",
    "CT 105 SIYAH": "CT 105",
    "CT 4215R": "CT 4215K",
    "CT 5421": "CT 5421SB",
    "CT 9115S": "CT 9115",
    "CT 9500A": "CT 9500",
    "CT 9500B": "CT 9500",
    "CT 9500M": "CT 9500",
    "CT 9500P": "CT 9500",
}


def aydinlatma_dunyam_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_kodlari: set[str],
    tsoft_oznur_teknik_indeksi: dict[tuple[str, str, str, str], set[str]],
) -> str:
    """Rakibin eklediği marka önekini yalnız birebir T-Soft kodu varsa kaldırır."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    onayli_hedef = AYDINLATMA_DUNYAM_ONAYLI_KOD_ESLESMELERI.get(
        (marka_norm, kaynak_kod),
        "",
    )
    if onayli_hedef in tsoft_kodlari:
        return onayli_hedef
    if marka_norm == metin_norm("ÖZNUR KABLO"):
        teknik_anahtar = elektrikciden_oznur_teknik_anahtar(urun_adi)
        teknik_adaylar = (
            tsoft_oznur_teknik_indeksi.get(teknik_anahtar, set())
            if teknik_anahtar
            else set()
        )
        if len(teknik_adaylar) == 1:
            return next(iter(teknik_adaylar))
    for onek in AYDINLATMA_DUNYAM_KOD_ONEKLERI.get(marka_norm, ()):
        if not kaynak_kod.startswith(onek):
            continue
        hedef_kod = kaynak_kod[len(onek) :].strip()
        if hedef_kod in tsoft_kodlari:
            return hedef_kod
    return ""


def cimen_elektrik_tsoft_kodu(
    marka: str,
    kod: str,
    tsoft_kodlari: set[str],
) -> str:
    """Çimen'in ürün kodu sonundaki site kayıt numarasını güvenle kaldırır."""
    kaynak_kod = metin_norm(kod)
    eslesme = re.fullmatch(r"(.+?)\s+\d{4,6}", kaynak_kod)
    if not eslesme:
        return ""
    hedef_kod = eslesme.group(1).strip()
    if hedef_kod in tsoft_kodlari:
        return hedef_kod
    marka_norm = metin_norm(marka)
    if marka_norm == metin_norm("AUDIO"):
        audio_hedefi = f"AUDIO {hedef_kod}"
        if audio_hedefi in tsoft_kodlari:
            return audio_hedefi
    if (
        marka_norm == metin_norm("SCHNEIDER")
        and hedef_kod.startswith("SCH ")
    ):
        schneider_hedefi = hedef_kod[len("SCH ") :].strip()
        if schneider_hedefi in tsoft_kodlari:
            return schneider_hedefi
    if marka_norm == metin_norm("CATA"):
        onayli_eslesmeler = CIMEN_ELEKTRIK_CATA_ONAYLI_KOD_ESLESMELERI
    elif marka_norm == metin_norm("NOAS"):
        onayli_eslesmeler = CIMEN_ELEKTRIK_NOAS_ONAYLI_KOD_ESLESMELERI
    elif marka_norm == metin_norm("ISILDAR"):
        onayli_eslesmeler = CIMEN_ELEKTRIK_ISILDAR_ONAYLI_KOD_ESLESMELERI
    else:
        onayli_eslesmeler = {}
    onayli_hedef = onayli_eslesmeler.get(hedef_kod, "")
    return onayli_hedef if onayli_hedef in tsoft_kodlari else ""


def koz_aydinlatma_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_cata_kodlari: set[str],
    tsoft_noas_kodlari: set[str],
) -> str:
    """Koz kodunu yalnız tekil veya açıkça belirtilmiş T-Soft varyantına bağlar."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    kaynak_urun_adi = metin_norm(urun_adi)

    if marka_norm == metin_norm("NOAS"):
        # Koz koduna 6500K/S1/ST gibi site ekleri koyabiliyor. Ürün adındaki
        # gerçek NOAS modeli T-Soft'ta birebir varsa onu kullanır; T10 paketini
        # tekli ürüne çevirmek gibi tahminler yapmaz.
        ad_kodlari = re.findall(
            r"\bYL\d+\s+\d+(?:\s+[A-Z]\d*)?\b",
            kaynak_urun_adi,
        )
        for ad_kodu in reversed(ad_kodlari):
            if ad_kodu in tsoft_noas_kodlari:
                return ad_kodu
        return ""

    if marka_norm != metin_norm("CATA"):
        return ""

    kaynak_kompakt = re.sub(r"[^A-Z0-9]", "", kaynak_kod)
    kompakt_adaylar = {
        hedef
        for hedef in tsoft_cata_kodlari
        if re.sub(r"[^A-Z0-9]", "", hedef) == kaynak_kompakt
    }
    if len(kompakt_adaylar) == 1:
        return next(iter(kompakt_adaylar))

    onayli_hedef = KOZ_AYDINLATMA_CATA_ONAYLI_KOD_ESLESMELERI.get(
        kaynak_kod,
        "",
    )
    if onayli_hedef in tsoft_cata_kodlari:
        return onayli_hedef

    model_eslesmesi = re.search(r"\bCT\s*(\d{3,4})\b", kaynak_kod)
    if not model_eslesmesi:
        model_eslesmesi = re.search(r"\bCT\s*(\d{3,4})\b", kaynak_urun_adi)
    if not model_eslesmesi:
        return ""
    model = model_eslesmesi.group(1)
    model_kompakt = f"CT{model}"
    aile_adaylari = {
        hedef
        for hedef in tsoft_cata_kodlari
        if re.sub(r"[^A-Z0-9]", "", hedef).startswith(model_kompakt)
    }
    if kaynak_kompakt == model_kompakt and len(aile_adaylari) == 1:
        return next(iter(aile_adaylari))

    renk_ekleri = (
        ("NATURAL ISIK", "K"),
        ("BEYAZ ISIK", "B"),
        ("GUN ISIGI", "G"),
        ("GUNISIGI", "G"),
        ("AMBER", "A"),
        ("YESIL", "Y"),
    )
    for renk_metni, renk_eki in renk_ekleri:
        if renk_metni not in kaynak_kod:
            continue
        renk_hedefi = f"CT {model}{renk_eki}"
        if renk_hedefi in tsoft_cata_kodlari:
            return renk_hedefi
    return ""


def ucfaz_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_marka_urunleri: dict[str, dict],
    tsoft_kodlari: set[str],
    tsoft_meanwell_indeksi: dict[str, set[str]],
) -> str:
    """3Faz kodunu yalnız marka içinde tekil ve doğrulanabilir modele dönüştürür."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    kaynak_urun_adi = metin_norm(urun_adi)
    hedef_bilgisi = tsoft_marka_urunleri.get(marka_norm, {})
    hedef_urunler = hedef_bilgisi.get("urunler", [])
    hedef_kodlari = hedef_bilgisi.get("kodlar", set())
    if not kaynak_kod:
        return ""

    onayli_hedef = STAR_AKIM_ONAYLI_KOD_ESLESMELERI.get(
        (marka_norm, kaynak_kod),
        "",
    )
    if (
        marka_norm == metin_norm("CATA")
        and not onayli_hedef
    ):
        onayli_hedef = UCFAZ_CATA_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("CATA")
        and not onayli_hedef
    ):
        onayli_hedef = ELEKTRIK_MARKET_CATA_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("JUPITER")
        and not onayli_hedef
    ):
        onayli_hedef = ELEKTRIK_MARKET_JUPITER_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("VIKO")
        and not onayli_hedef
    ):
        onayli_hedef = ELEKTRIK_MARKET_VIKO_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("VIKO")
        and not onayli_hedef
    ):
        onayli_hedef = UCFAZ_VIKO_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("ERSE")
        and not onayli_hedef
    ):
        onayli_hedef = UCFAZ_ERSE_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("WAGO")
        and not onayli_hedef
    ):
        onayli_hedef = UCFAZ_WAGO_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if (
        marka_norm == metin_norm("MEAN WELL")
        and not onayli_hedef
    ):
        onayli_hedef = UCFAZ_MEANWELL_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod,
            "",
        )
    if onayli_hedef in tsoft_kodlari:
        return onayli_hedef

    # Bazı Cata kartlarında ürün adında marka bulunmuyor. Kaynak kodu CT
    # ile başlıyor ve T-Soft'ta birebir aynı kod varsa marka adına bakmadan
    # güvenli biçimde aynı karta bağla.
    if (
        marka_norm == metin_norm("CATA")
        and kaynak_kod.startswith("CT ")
        and kaynak_kod in tsoft_kodlari
    ):
        return kaynak_kod

    if kaynak_kod in hedef_kodlari:
        return kaynak_kod

    kompakt_kod = re.sub(r"[^A-Z0-9]", "", kaynak_kod)
    if len(kompakt_kod) >= 4:
        kompakt_adaylar = hedef_bilgisi.get("kompakt", {}).get(
            kompakt_kod,
            set(),
        )
        if len(kompakt_adaylar) == 1:
            return next(iter(kompakt_adaylar))

    if marka_norm == metin_norm("MEAN WELL"):
        meanwell_eslesmesi = elektrik_market_meanwell_tsoft_kodu(
            kaynak_kod,
            tsoft_meanwell_indeksi,
        )
        if meanwell_eslesmesi:
            return meanwell_eslesmesi

    adaylar = set()
    ad_parcalari = kaynak_urun_adi.split()
    maks_parca = min(
        hedef_bilgisi.get("maks_parca", 1),
        len(ad_parcalari),
    )
    for parca_sayisi in range(1, maks_parca + 1):
        for baslangic in range(0, len(ad_parcalari) - parca_sayisi + 1):
            aday_kod = " ".join(
                ad_parcalari[baslangic : baslangic + parca_sayisi]
            )
            if (
                aday_kod in hedef_kodlari
                and zeybek_kod_anlamli_mi(aday_kod)
                and normlanmis_kod_ifadesi_iceriyor(
                    kaynak_urun_adi,
                    aday_kod,
                )
            ):
                adaylar.add(aday_kod)
    kaynak_kompakt = re.sub(r"[^A-Z0-9]", "", kaynak_kod)
    entes_modeli_anlamli = (
        len(kaynak_kompakt) >= 3
        and bool(re.search(r"[A-Z]", kaynak_kompakt))
        and bool(re.search(r"\d", kaynak_kompakt))
    )
    if marka_norm == metin_norm("ENTES") and entes_modeli_anlamli:
        adaylar.update(
            hedef_kod
            for hedef_kod, _, hedef_urun_adi_ham in hedef_urunler
            if ham_metin_model_kodu_iceriyor(
                hedef_urun_adi_ham,
                kod,
            )
        )
    return next(iter(adaylar)) if len(adaylar) == 1 else ""


def erse_teknik_anahtari(deger, tsoft_kaynagi: bool = False):
    """Erse kablosunu aile + damar/kesit bilgisiyle karşılaştırılabilir yapar."""
    raw = "" if pd.isna(deger) else str(deger).upper()
    raw = unicodedata.normalize("NFKD", raw)
    raw = "".join(c for c in raw if not unicodedata.combining(c))
    raw = raw.replace("İ", "I")

    if "LIYCY" in raw:
        aile = "LIYCY"
    elif "LIHCH" in raw and ("FE180" in raw or "PH120" in raw):
        aile = "LIHCH-FE"
    elif (
        ("ERVITAL" in raw or "JE-H" in raw or "JEHSTH" in raw)
        and "X2X" in raw
    ):
        aile = "JEHSTH"
    elif (
        "YSLYJZ" in raw
        or (
            tsoft_kaynagi
            and (
                "ERFLEX 130-Y" in raw
                or re.search(r"\bYSLY\b", raw)
            )
        )
    ):
        aile = "YSLYJZ"
    else:
        return None

    def olcu(deger):
        try:
            sayi = Decimal(str(deger).replace(",", "."))
        except InvalidOperation:
            return ""
        return format(sayi.normalize(), "f")

    if aile == "JEHSTH":
        eslesme = re.search(
            r"(?<!\d)(\d+)X2X(\d+(?:[.,]\d+)?)\s*\+\s*(\d+(?:[.,]\d+)?)",
            raw,
        )
        if not eslesme:
            return None
        return (
            aile,
            int(eslesme.group(1)),
            olcu(eslesme.group(2)),
            olcu(eslesme.group(3)),
        )

    eslesme = re.search(
        r"(?<!\d)(\d+)X(\d+(?:[.,]\d+)?)",
        raw,
    )
    if not eslesme:
        return None
    return (
        aile,
        int(eslesme.group(1)),
        olcu(eslesme.group(2)),
    )


def krc_erse_tsoft_kodu(
    urun_adi: str,
    tsoft_urunler: list[tuple[str, str]],
) -> str:
    """KRC Erse kablosunu aynı aile ve aynı teknik ölçüdeki tek T-Soft kartına bağlar."""
    kaynak_anahtari = erse_teknik_anahtari(urun_adi)
    if not kaynak_anahtari:
        return ""
    adaylar = {
        kod
        for kod, tsoft_urun_adi in tsoft_urunler
        if erse_teknik_anahtari(
            tsoft_urun_adi,
            tsoft_kaynagi=True,
        )
        == kaynak_anahtari
    }
    return next(iter(adaylar)) if len(adaylar) == 1 else ""


def entes_model_anahtari(deger: object) -> str:
    """ENTES modelini ENT/ENTES öneki ve ayraç farklarından arındırır."""
    anahtar = metin_norm(deger)
    anahtar = re.sub(r"^(?:ENTES|ENT)(?=[^A-Z0-9]|[A-Z])", "", anahtar)
    return re.sub(r"[^A-Z0-9]", "", anahtar)


def entes_fiyat_model_indeksi_olustur(
    fiyat_listesi_dosyasi: Path,
    tsoft_entes_kodlari: set[str],
) -> dict[str, set[str]]:
    """Resmî ENTES fiyat listesinden tekil Model -> T-Soft ürün kodu indeksi kurar."""
    if not fiyat_listesi_dosyasi.exists():
        return {}

    ham = pd.read_excel(
        fiyat_listesi_dosyasi,
        dtype=str,
        header=None,
    ).fillna("")
    baslik_satiri = None
    for satir_no, row in ham.iterrows():
        basliklar = {metin_norm(deger) for deger in row.tolist()}
        if "URUN KODU" in basliklar and "URUN MODELI" in basliklar:
            baslik_satiri = satir_no
            break
    if baslik_satiri is None:
        return {}

    fiyat_listesi = pd.read_excel(
        fiyat_listesi_dosyasi,
        dtype=str,
        header=baslik_satiri,
    ).fillna("")
    kolonlar = {metin_norm(kolon): kolon for kolon in fiyat_listesi.columns}
    kod_kolonu = kolonlar.get("URUN KODU")
    model_kolonu = kolonlar.get("URUN MODELI")
    if not kod_kolonu or not model_kolonu:
        return {}

    indeks: dict[str, set[str]] = {}
    for _, row in fiyat_listesi.iterrows():
        urun_kodu = metin_norm(row.get(kod_kolonu, ""))
        model_anahtari = entes_model_anahtari(row.get(model_kolonu, ""))
        if (
            not model_anahtari
            or urun_kodu not in tsoft_entes_kodlari
        ):
            continue
        indeks.setdefault(model_anahtari, set()).add(urun_kodu)
    return indeks


def entes_akim_trafosu_anahtari(
    kod: object,
    urun_adi: object,
) -> tuple[str, int, Decimal, bool] | None:
    """ENTES akım trafosunu seri, akım oranı, VA ve mühür bilgisiyle tanımlar."""
    def teknik_norm(deger: object) -> str:
        metin = unicodedata.normalize("NFKD", str(deger).upper())
        metin = "".join(
            karakter
            for karakter in metin
            if not unicodedata.combining(karakter)
        )
        return metin.replace("İ", "I").replace("ı", "I").replace(",", ".")

    kod_norm = teknik_norm(kod)
    ad_norm = teknik_norm(urun_adi)
    seri_eslesmesi = re.search(
        r"(?<![A-Z0-9])(A100T|A100|A80T|A80|A60T|A60|A40|A30|AB)\b",
        kod_norm,
    ) or re.search(
        r"ENT[. ]*(A100T|A100|A80T|A80|A60T|A60|A40|A30|AB)\b",
        ad_norm,
    )
    oran_eslesmesi = re.search(
        r"(?<!\d)(\d+)\s*/\s*5\s*A?\b",
        f"{kod_norm} {ad_norm}",
    )
    va_eslesmesi = re.search(r"(\d+(?:\.\d+)?)\s*VA\b", ad_norm)
    if not seri_eslesmesi or not oran_eslesmesi or not va_eslesmesi:
        return None
    try:
        va = Decimal(va_eslesmesi.group(1))
    except InvalidOperation:
        return None
    muhurlu = "MUH" in kod_norm or "MUHURLU" in ad_norm
    return (
        seri_eslesmesi.group(1),
        int(oran_eslesmesi.group(1)),
        va,
        muhurlu,
    )


def entes_tsoft_akim_trafosu_anahtari(
    urun_adi: object,
) -> tuple[str, int, Decimal, bool] | None:
    """T-Soft ENTES akım trafosu adından aynı teknik anahtarı çıkarır."""
    ad_norm = unicodedata.normalize("NFKD", str(urun_adi).upper())
    ad_norm = "".join(
        karakter
        for karakter in ad_norm
        if not unicodedata.combining(karakter)
    )
    ad_norm = (
        ad_norm.replace("İ", "I")
        .replace("ı", "I")
        .replace(",", ".")
    )
    seri_eslesmesi = re.search(
        r"ENT[. ]*(A100T|A100|A80T|A80|A60T|A60|A40|A30|AB)\b",
        ad_norm,
    )
    if not seri_eslesmesi:
        return None
    oran_eslesmesi = re.search(
        r"(?<!\d)(\d+)\s*/\s*5\s*A?\b",
        ad_norm,
    )
    if not oran_eslesmesi:
        oran_eslesmesi = re.search(
            re.escape(seri_eslesmesi.group(0)) + r"\s*(\d+)",
            ad_norm,
        )
    va_eslesmesi = re.search(r"(\d+(?:\.\d+)?)\s*VA\b", ad_norm)
    if not oran_eslesmesi or not va_eslesmesi:
        return None
    try:
        va = Decimal(va_eslesmesi.group(1))
    except InvalidOperation:
        return None
    return (
        seri_eslesmesi.group(1),
        int(oran_eslesmesi.group(1)),
        va,
        "MUHURLU" in ad_norm,
    )


def krc_entes_tsoft_kodu(
    kod: str,
    urun_adi: str,
    hedef_bilgisi: dict,
) -> str:
    """KRC ENTES modelini fiyat listesi ve teknik akım trafosu anahtarıyla eşleştirir."""
    kaynak_anahtari = entes_model_anahtari(kod)
    model_adaylari = set(
        hedef_bilgisi.get("entes_model", {}).get(kaynak_anahtari, set())
    )

    # KRC MPR-45S/MPR-47S yazar; fiyat listesinde aynı ürün IP54 ekiyle bulunur.
    if re.fullmatch(r"MPR\d+[A-Z]+", kaynak_anahtari):
        model_adaylari.update(
            hedef_bilgisi.get("entes_model", {}).get(
                f"{kaynak_anahtari}IP54",
                set(),
            )
        )
    if len(model_adaylari) == 1:
        return next(iter(model_adaylari))
    if len(model_adaylari) > 1:
        return ""

    trafo_anahtari = entes_akim_trafosu_anahtari(kod, urun_adi)
    trafo_adaylari = (
        hedef_bilgisi.get("entes_akim_trafosu", {}).get(
            trafo_anahtari,
            set(),
        )
        if trafo_anahtari
        else set()
    )
    if len(trafo_adaylari) == 1:
        return next(iter(trafo_adaylari))

    # Fiyat listesinde model adı farklı yazılan ama T-Soft kartı tekil doğrulananlar.
    onayli_ozel_eslesmeler = {
        "ENSCCT24100A": "M5129",
        "3X150330MV3PMD25": "M3649",
        "ERTC01": "M0219",
    }
    ozel_hedef = onayli_ozel_eslesmeler.get(kaynak_anahtari, "")
    if ozel_hedef in hedef_bilgisi.get("kodlar", set()):
        return ozel_hedef
    return ""


def krc_elektrik_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_marka_urunleri: dict[str, dict],
) -> str:
    """KRC'nin firma önekli kodunu aynı marka içindeki tekil T-Soft modeline bağlar."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    kaynak_urun_adi = metin_norm(urun_adi)
    hedef_bilgisi = tsoft_marka_urunleri.get(marka_norm, {})
    hedef_kodlari = hedef_bilgisi.get("kodlar", set())
    kompakt_indeks = hedef_bilgisi.get("kompakt", {})
    if not kaynak_kod or not hedef_kodlari:
        return ""

    if marka_norm == metin_norm("ENTES"):
        entes_hedefi = krc_entes_tsoft_kodu(
            kod,
            urun_adi,
            hedef_bilgisi,
        )
        if entes_hedefi:
            return entes_hedefi

    if marka_norm == metin_norm("ERSE"):
        erse_hedefi = krc_erse_tsoft_kodu(
            urun_adi,
            hedef_bilgisi.get("ham_urunler", []),
        )
        if erse_hedefi:
            return erse_hedefi

    kod_adaylari = set()
    if kaynak_kod in hedef_kodlari:
        kod_adaylari.add(kaynak_kod)

    # KRC kodlarının ilk parçası çoğunlukla firma kısaltmasıdır:
    # "S 3KD34300NE100", "SN A9F84206", "LR 004846" gibi.
    kod_parcalari = kaynak_kod.split()
    kaynak_bicimleri = [kaynak_kod]
    sifirsiz_kod = ""
    if len(kod_parcalari) > 1:
        onejsiz_kod = " ".join(kod_parcalari[1:]).strip()
        kaynak_bicimleri.append(onejsiz_kod)
        if (
            marka_norm == metin_norm("EMAS")
            and onejsiz_kod in hedef_kodlari
        ):
            kod_adaylari.add(onejsiz_kod)
        if marka_norm == metin_norm("VIKO"):
            viko_parcalari = kod_parcalari[1:]
            if (
                len(viko_parcalari) >= 2
                and re.fullmatch(r"\d{6,10}", viko_parcalari[0])
                and re.fullmatch(r"[A-Z]", viko_parcalari[1])
                and viko_parcalari[0] in hedef_kodlari
            ):
                kod_adaylari.add(viko_parcalari[0])
        if marka_norm in {
            metin_norm("LEGRAND"),
            metin_norm("PHOENIX CONTACT"),
        }:
            # KRC'de bazı Legrand/Phoenix katalog kodları başında sıfırla,
            # T-Soft'ta sıfırsız tutuluyor. Yalnız sayısal kod ve aynı marka
            # içinde tek karşılık olduğunda kullanılır.
            sayisal_kod = "".join(kod_parcalari[1:])
            if sayisal_kod.isdigit():
                sifirsiz_kod = sayisal_kod.lstrip("0") or "0"
    for kaynak_bicimi in kaynak_bicimleri:
        kompakt = re.sub(r"[^A-Z0-9]", "", kaynak_bicimi)
        if len(kompakt) < 4:
            continue
        kompakt_adaylar = kompakt_indeks.get(kompakt, set())
        if len(kompakt_adaylar) == 1:
            kod_adaylari.update(kompakt_adaylar)

    if len(kod_adaylari) == 1:
        return next(iter(kod_adaylari))
    if len(kod_adaylari) > 1:
        return ""

    if sifirsiz_kod:
        sifirsiz_adaylar = kompakt_indeks.get(sifirsiz_kod, set())
        if len(sifirsiz_adaylar) == 1:
            return next(iter(sifirsiz_adaylar))

    if marka_norm == metin_norm("ORTAÇ"):
        return ""

    # Ürün adında gerçek üretici modeli bulunuyorsa yalnız aynı marka içindeki
    # tek bir T-Soft koduyla eşleşmesine izin verilir.
    adaylar = set()
    ad_parcalari = kaynak_urun_adi.split()
    maks_parca = min(
        hedef_bilgisi.get("maks_parca", 1),
        len(ad_parcalari),
    )
    for parca_sayisi in range(1, maks_parca + 1):
        for baslangic in range(0, len(ad_parcalari) - parca_sayisi + 1):
            aday_kod = " ".join(
                ad_parcalari[baslangic : baslangic + parca_sayisi]
            )
            if (
                aday_kod in hedef_kodlari
                and zeybek_kod_anlamli_mi(aday_kod)
                and normlanmis_kod_ifadesi_iceriyor(
                    kaynak_urun_adi,
                    aday_kod,
                )
            ):
                adaylar.add(aday_kod)

    return next(iter(adaylar)) if len(adaylar) == 1 else ""


def buyuk_kalpler_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_marka_urunleri: dict[str, dict],
    tsoft_kodlari: set[str],
) -> str:
    """Büyük Kalpler kodlarını yalnız doğrulanan T-Soft model kodlarına bağlar."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    kod_adaylari = set()

    if marka_norm == metin_norm("CATA"):
        kompakt_kod = re.sub(r"[^A-Z0-9]", "", kaynak_kod)
        if kompakt_kod.startswith("CT") and len(kompakt_kod) > 2:
            govde = kompakt_kod[2:]
            cata_adaylari = {
                f"CT {govde}",
                f"CT-{govde}",
            }
            if govde[-1].isalpha():
                cata_adaylari.update(
                    {
                        f"CT {govde[:-1]}",
                        f"CT-{govde[:-1]}",
                    }
                )
            if len(govde) > 1 and govde[-1].isalpha() and govde[-2].isalpha():
                cata_adaylari.update(
                    {
                        f"CT {govde[:-2]}",
                        f"CT-{govde[:-2]}",
                    }
                )
            kod_adaylari.update(
                metin_norm(aday)
                for aday in cata_adaylari
                if metin_norm(aday) in tsoft_kodlari
            )
            kaynak_urun_adi = metin_norm(urun_adi)
            for ad_kodu in re.findall(r"\bCT[\s\-_]*\d+[A-Z]*\b", kaynak_urun_adi):
                ad_kompakt = re.sub(r"[^A-Z0-9]", "", ad_kodu)
                if not ad_kompakt.startswith("CT"):
                    continue
                ad_govde = ad_kompakt[2:]
                ad_adaylari = {
                    f"CT {ad_govde}",
                    f"CT-{ad_govde}",
                }
                kod_adaylari.update(
                    metin_norm(aday)
                    for aday in ad_adaylari
                    if metin_norm(aday) in tsoft_kodlari
                )

    if marka_norm == metin_norm("ACK"):
        ack_eslesmesi = re.fullmatch(r"(.+\d)[A-Z]{1,2}", kaynak_kod)
        if ack_eslesmesi and ack_eslesmesi.group(1) in tsoft_kodlari:
            kod_adaylari.add(ack_eslesmesi.group(1))
        if kaynak_kod.endswith(" PRO"):
            pro_siz_kod = kaynak_kod[: -len(" PRO")].strip()
            if pro_siz_kod in tsoft_kodlari:
                kod_adaylari.add(pro_siz_kod)
        if kaynak_kod.startswith("ACK"):
            ack_siz_kod = kaynak_kod[len("ACK") :].strip()
            if ack_siz_kod in tsoft_kodlari:
                kod_adaylari.add(ack_siz_kod)

    if marka_norm == metin_norm("OZNUR KABLO"):
        oznur_eslesmesi = elektromarketim_oznur_tsoft_kodu(kaynak_kod)
        if oznur_eslesmesi in tsoft_kodlari:
            kod_adaylari.add(oznur_eslesmesi)

    if marka_norm == metin_norm("AUDIO"):
        audio_ozel_eslesmeler = {
            metin_norm("001284"): metin_norm("1284"),
            metin_norm("001460"): metin_norm("AUDIO 001460"),
            metin_norm("001463"): metin_norm("AUDIO 001463"),
        }
        audio_hedef = audio_ozel_eslesmeler.get(kaynak_kod)
        if audio_hedef in tsoft_kodlari:
            kod_adaylari.add(audio_hedef)

    genel_eslesme = krc_elektrik_tsoft_kodu(
        marka,
        kod,
        urun_adi,
        tsoft_marka_urunleri,
    )
    if genel_eslesme in tsoft_kodlari:
        kod_adaylari.add(genel_eslesme)

    return next(iter(kod_adaylari)) if len(kod_adaylari) == 1 else ""


def elektrik_malzemesi_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_klemsan_kodlari: set[str],
    tsoft_gwest_kodlari: set[str],
    tsoft_kraus_naimer_kodlari: set[str],
    tsoft_ortac_kompakt_indeksi: dict[str, set[str]],
    tsoft_schneider_kodlari: set[str],
    tsoft_kodlari: set[str],
) -> str:
    """Elektrik Malzemesi iç kodunu yalnız doğrulanan marka kuralıyla dönüştürür."""
    marka_norm = metin_norm(marka)
    kaynak_kod = metin_norm(kod)
    if marka_norm == metin_norm("KLEMSAN"):
        hedef_kod = re.sub(r"[^A-Z0-9]", "", kaynak_kod).lstrip("0")
        return hedef_kod if hedef_kod in tsoft_klemsan_kodlari else ""
    if marka_norm == metin_norm("GWEST"):
        hedef_kod = re.sub(r"^GVN\s+", "", kaynak_kod).strip()
        return hedef_kod if hedef_kod in tsoft_gwest_kodlari else ""
    if marka_norm == metin_norm("K. NAIMER"):
        kaynak_urun_adi = metin_norm(urun_adi)
        kompakt_urun_adi = re.sub(r"[^A-Z0-9]", "", kaynak_urun_adi)
        kaynak_kutulu_varyantlari = {
            re.sub(r"[^A-Z0-9]", "", varyant)
            for varyant in re.findall(r"\bTRA\s*\d+\b", kaynak_urun_adi)
        }
        adaylar = set()
        for tsoft_kodu in tsoft_kraus_naimer_kodlari:
            kompakt_tsoft_kodu = re.sub(r"[^A-Z0-9]", "", tsoft_kodu)
            if len(kompakt_tsoft_kodu) < 4:
                continue
            if kompakt_tsoft_kodu not in kompakt_urun_adi:
                continue
            if any(
                varyant not in kompakt_tsoft_kodu
                for varyant in kaynak_kutulu_varyantlari
            ):
                continue
            adaylar.add(tsoft_kodu)
        return next(iter(adaylar)) if len(adaylar) == 1 else ""
    if marka_norm == metin_norm("ORTAÇ"):
        kompakt_kod = re.sub(r"[^A-Z0-9]", "", kaynak_kod)
        adaylar = tsoft_ortac_kompakt_indeksi.get(kompakt_kod, set())
        return next(iter(adaylar)) if len(adaylar) == 1 else ""
    if marka_norm == metin_norm("SCHNEIDER"):
        hedef_kod = ELEKTRIK_MALZEMESI_SCHNEIDER_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod, ""
        )
        return hedef_kod if hedef_kod in tsoft_schneider_kodlari else ""
    if marka_norm == metin_norm("SIEMENS"):
        hedef_kod = ELEKTRIK_MALZEMESI_SIEMENS_ONAYLI_KOD_ESLESMELERI.get(
            kaynak_kod, ""
        )
        return hedef_kod if hedef_kod in tsoft_kodlari else ""
    return ""


def elektrikciden_tsoft_kodu(
    marka: str,
    kod: str,
    urun_adi: str,
    tsoft_cata_kodlari: set[str],
    belirsiz_cata_kodlari: set[str],
    tsoft_oznur_teknik_indeksi: dict[tuple[str, str, str, str], set[str]],
    tsoft_viko_kodlari: set[str],
    tsoft_mutlusan_kodlari: set[str],
    tsoft_siemens_kodlari: set[str],
    tsoft_schneider_kodlari: set[str],
    tsoft_noas_kodlari: set[str],
    tsoft_audio_sayisal_indeksi: dict[str, set[str]],
    belirsiz_audio_kodlari: set[str],
    tsoft_entes_kodlari: set[str],
) -> str:
    """Elektrikçiden ürün adındaki doğrulanmış marka modelini T-Soft koduna çevirir."""
    marka_norm = metin_norm(marka)
    if marka_norm == metin_norm("ÖZNUR"):
        kaynak_urun_adi = str(urun_adi)
        if re.search(r"-\s*100\s*M\b", kaynak_urun_adi.upper()):
            return ""
        if not re.search(r"-\s*1\s*M\b", kaynak_urun_adi.upper()):
            return ""
        anahtar = elektrikciden_oznur_teknik_anahtar(
            kaynak_urun_adi,
            kod,
        )
        adaylar = tsoft_oznur_teknik_indeksi.get(anahtar, set())
        return next(iter(adaylar)) if len(adaylar) == 1 else ""

    if marka_norm == metin_norm("AUDIO"):
        sayisal_kodlar = {
            (eslesme.lstrip("0") or "0")
            for eslesme in re.findall(
                r"(?<!\d)(\d{4,6})(?![A-Z0-9])",
                str(urun_adi).upper(),
            )
        } - belirsiz_audio_kodlari
        adaylar = set()
        for sayisal_kod in sayisal_kodlar:
            adaylar.update(tsoft_audio_sayisal_indeksi.get(sayisal_kod, set()))
        return next(iter(adaylar)) if len(adaylar) == 1 else ""

    kodlu_marka_ayarlari = {
        metin_norm("VIKO"): (tsoft_viko_kodlari, 6),
        metin_norm("MUTLUSAN"): (tsoft_mutlusan_kodlari, 5),
        metin_norm("SIEMENS"): (tsoft_siemens_kodlari, 5),
        metin_norm("SCHNEIDER"): (tsoft_schneider_kodlari, 5),
        metin_norm("NOAS"): (tsoft_noas_kodlari, 5),
    }
    kodlu_marka_ayari = kodlu_marka_ayarlari.get(marka_norm)
    if kodlu_marka_ayari:
        marka_kodlari, minimum_kod_uzunlugu = kodlu_marka_ayari
        if marka_norm == metin_norm("VIKO"):
            onayli_ad_eslesmesi = (
                ELEKTRIKCIDEN_VIKO_ONAYLI_AD_ESLESMELERI.get(
                    metin_norm(urun_adi), ""
                )
            )
            if (
                onayli_ad_eslesmesi
                and metin_norm(onayli_ad_eslesmesi) in marka_kodlari
            ):
                return metin_norm(onayli_ad_eslesmesi)
        urun_adi_kompakt = re.sub(r"[^A-Z0-9]", "", metin_norm(urun_adi))
        adaylar = {
            tsoft_kodu
            for tsoft_kodu in marka_kodlari
            if (
                len(re.sub(r"[^A-Z0-9]", "", tsoft_kodu))
                >= minimum_kod_uzunlugu
                and any(ch.isdigit() for ch in tsoft_kodu)
                and re.sub(r"[^A-Z0-9]", "", tsoft_kodu) in urun_adi_kompakt
            )
        }
        if not adaylar:
            return ""
        en_uzun = max(
            len(re.sub(r"[^A-Z0-9]", "", aday))
            for aday in adaylar
        )
        en_uzun_adaylar = {
            aday
            for aday in adaylar
            if len(re.sub(r"[^A-Z0-9]", "", aday)) == en_uzun
        }
        return next(iter(en_uzun_adaylar)) if len(en_uzun_adaylar) == 1 else ""

    if marka_norm == metin_norm("ENTES"):
        kaynak_urun_adi = metin_norm(urun_adi)
        for model, hedef in ELEKTRIKCIDEN_ENTES_MODEL_ESLESMELERI.items():
            model_norm = metin_norm(model)
            if (
                re.search(
                    rf"(?:^| ){re.escape(model_norm)}(?: |$)",
                    kaynak_urun_adi,
                )
                and metin_norm(hedef) in tsoft_entes_kodlari
            ):
                return metin_norm(hedef)
        if (
            re.search(r"(?:^| )125 5A(?: |$)", kaynak_urun_adi)
            and "10VA" in kaynak_urun_adi.replace(" ", "")
            and metin_norm("M4225") in tsoft_entes_kodlari
        ):
            return metin_norm("M4225")
        return ""

    if marka_norm == metin_norm("CATA"):
        onayli_ad_eslesmesi = ELEKTRIKCIDEN_CATA_ONAYLI_AD_ESLESMELERI.get(
            metin_norm(urun_adi), ""
        )
        if (
            onayli_ad_eslesmesi
            and metin_norm(onayli_ad_eslesmesi) in tsoft_cata_kodlari
        ):
            return metin_norm(onayli_ad_eslesmesi)

    if marka_norm != metin_norm("CATA"):
        return ""

    kaynak_urun_adi = metin_norm(urun_adi)
    adaylar = {
        metin_norm(eslesme)
        for eslesme in re.findall(r"\bCT\s*\d+[A-Z]*\b", kaynak_urun_adi)
    }
    mevcut_adaylar = (adaylar & tsoft_cata_kodlari) - belirsiz_cata_kodlari
    return next(iter(mevcut_adaylar)) if len(mevcut_adaylar) == 1 else ""


def elektrikciden_oznur_teknik_anahtar(
    urun_adi: str,
    kaynak_kod: str = "",
) -> tuple[str, str, str, str] | None:
    """Öznur kablo ailesi, ölçüsü, rengi ve HFFR bilgisinden güvenli anahtar üretir."""
    raw = (
        str(urun_adi)
        .upper()
        .replace("MM²", "MM ")
        .replace("MM2", "MM ")
    )
    raw_norm = metin_norm(raw)
    aile = next(
        (
            aday
            for aday in ("NHXMH", "N2XH", "NAYY", "NYAF", "NYM", "NYA", "TTR")
            if re.search(rf"\b{aday}\b", raw_norm)
        ),
        "",
    )
    if not aile:
        return None

    eslesme = re.search(
        r"\b(\d+)\s*[Xx]\s*(\d+(?:[,.]\d+)?)"
        r"(?:\s*\+\s*(\d+(?:[,.]\d+)?))?\s*MM\b",
        raw,
    )
    if eslesme:
        damar, kesit, ek_kesit = eslesme.groups()
        olcu = f"{damar}X{kesit.replace('.', ',')}"
        if ek_kesit:
            olcu += f"+{ek_kesit.replace('.', ',')}"
    elif aile in {"NYA", "NYAF"}:
        eslesme = re.search(r"\b(\d+(?:[,.]\d+)?)\s*MM\b", raw)
        if not eslesme:
            return None
        olcu = f"1X{eslesme.group(1).replace('.', ',')}"
    else:
        return None

    renk = ""
    kaynak_kod_norm = metin_norm(kaynak_kod)
    if kaynak_kod_norm:
        kompakt_kaynak_kod = re.sub(r"[^A-Z0-9]", "", kaynak_kod_norm)
        for kod_eki, renk_adi in (
            ("SY", "SARI YESIL"),
            ("KV", "KAHVE"),
            ("K", "KIRMIZI"),
            ("M", "MAVI"),
            ("S", "SIYAH"),
            ("B", "BEYAZ"),
        ):
            if kompakt_kaynak_kod.endswith(kod_eki):
                renk = renk_adi
                break
    if not renk:
        for renk_ifadesi, renk_adi in (
            ("SARI YESIL", "SARI YESIL"),
            ("KIRMIZI", "KIRMIZI"),
            ("KAHVE", "KAHVE"),
            ("MAVI", "MAVI"),
            ("SIYAH", "SIYAH"),
            ("BEYAZ", "BEYAZ"),
        ):
            if renk_ifadesi in raw_norm:
                renk = renk_adi
                break

    if aile in {"NYA", "NYAF", "TTR"} and not renk:
        return None
    hffr = "HFFR" if ("HALOGEN FREE" in raw_norm or "HFFR" in raw_norm) else "NORMAL"
    return aile, olcu, renk, hffr


def elektromarketim_oznur_tsoft_kodu(kod: str) -> str:
    """Elektromarketim Öznur teknik kodunu T-Soft iç kod biçimine dönüştürür."""
    eslesme = re.fullmatch(
        r"OZNUR (H07V K|H07V R|H07V U|H07Z1 K) (.+?)(KH|M|S|K)",
        kod,
    )
    if eslesme:
        kablo_tipi, kesit, renk_kodu = eslesme.groups()
        renk = {
            "KH": "KAHVE",
            "M": "MAVI",
            "S": "SIYAH",
            "K": "KIRMIZI",
        }[renk_kodu]
        urun_tipi = "NYAF" if kablo_tipi in {"H07V K", "H07Z1 K"} else "NYA"
        hffr = " HFFR" if kablo_tipi == "H07Z1 K" else ""
        return f"O 1X{kesit} {renk} {urun_tipi}{hffr}"

    eslesme = re.fullmatch(r"OZNUR H05VV F (.+?)S", kod)
    if eslesme:
        return f"O {eslesme.group(1)} SIYAH TTR"

    for kablo_tipi, ek in (
        ("N2XH", "HFFR"),
        ("NHXMH", "ANTIGRON"),
        ("NYM", "ANTIGRON"),
    ):
        onek = f"OZNUR {kablo_tipi} "
        if kod.startswith(onek):
            return f"O {kablo_tipi} {kod[len(onek):]} {ek}"
    return ""


def elektrik_market_jupiter_tsoft_kodu(
    kod: str,
    urun_adi: str,
    tsoft_kod_indeksi: dict[str, set[str]],
) -> str:
    """Jupiter varyant kodunu yalnızca açık ve tekil olduğunda T-Soft koduna bağlar."""
    kaynak = metin_norm(kod)
    ad = metin_norm(urun_adi)
    ilk_parca = kaynak.split()[0] if kaynak else ""
    if not re.fullmatch(r"[A-Z]{1,3}\d+", ilk_parca):
        return ""

    adaylar = tsoft_kod_indeksi.get(ilk_parca, set())
    if not adaylar:
        return ""

    donusen_kodlar = {kaynak}
    donusen_kodlar.add(
        re.sub(r" ([A-Z]{1,2})(?:827|830|840|865)$", r" \1", kaynak)
    )
    donusen_kodlar.add(re.sub(r" ([A-Z]) \1$", r" \1", kaynak))

    guvenli_adaylar = set()
    for hedef in adaylar:
        if hedef in donusen_kodlar:
            guvenli_adaylar.add(hedef)
            continue
        if hedef == ilk_parca:
            continue
        if re.search(rf"(?:^| ){re.escape(hedef)}(?: (?:827|830|840|865))?$", ad):
            guvenli_adaylar.add(hedef)
            continue
        if hedef.endswith(" B") and re.search(
            rf"(?:^| ){re.escape(ilk_parca)} B S$", ad
        ):
            guvenli_adaylar.add(hedef)
            continue
        if hedef.endswith(" S") and re.search(
            rf"(?:^| ){re.escape(ilk_parca)} S S$", ad
        ):
            guvenli_adaylar.add(hedef)

    return next(iter(guvenli_adaylar)) if len(guvenli_adaylar) == 1 else ""


def meanwell_bicim_anahtari(kod: str) -> str:
    """Meanwell kodundaki ayraçları ve sayısal ön sıfırları eşitleyen anahtar."""
    parcalar = re.findall(r"[A-Z]+|\d+(?:\.\d+)?", metin_norm(kod))
    sonuc = []
    for parca in parcalar:
        if re.fullmatch(r"\d+", parca):
            sonuc.append(str(int(parca)))
        else:
            sonuc.append(parca)
    return " ".join(sonuc)


def elektrik_market_meanwell_tsoft_kodu(
    kod: str,
    tsoft_meanwell_indeksi: dict[str, set[str]],
) -> str:
    """Meanwell kodunu yalnızca güvenli biçim farkı veya ELG 3Y ekiyle eşleştirir."""
    kaynak = metin_norm(kod)
    if not kaynak:
        return ""

    anahtarlar = [meanwell_bicim_anahtari(kaynak)]
    if kaynak.startswith("ELG "):
        anahtarlar.append(meanwell_bicim_anahtari(f"{kaynak} 3Y"))

    for anahtar in anahtarlar:
        adaylar = tsoft_meanwell_indeksi.get(anahtar, set())
        if len(adaylar) == 1:
            hedef = next(iter(adaylar))
            if hedef != kaynak:
                return hedef
    return ""


def elektrik_market_mutlusan_tsoft_kodu(
    urun_adi: str,
    tsoft_mutlusan_kodlari: set[str],
) -> str:
    """Ürün adında açıkça bulunan tekil Mutlusan T-Soft kodunu döndürür."""
    ad = metin_norm(urun_adi)
    if not ad:
        return ""

    adaylar = []
    for kod in tsoft_mutlusan_kodlari:
        yalniz_harf_sayi = re.sub(r"[^A-Z0-9]", "", kod)
        if len(yalniz_harf_sayi) < 6:
            continue
        if re.search(rf"(?:^| ){re.escape(kod)}(?: |$)", ad):
            adaylar.append(kod)

    en_uzunlar = [
        kod
        for kod in adaylar
        if not any(kod != diger and kod in diger for diger in adaylar)
    ]
    en_uzunlar = sorted(set(en_uzunlar))
    return en_uzunlar[0] if len(en_uzunlar) == 1 else ""


def elextrix_tsoft_kodu(
    marka: str,
    kod: str,
    tsoft_kodlari: set[str],
    tsoft_kompakt_kod_indeksi: dict[str, set[str]],
) -> str:
    """Elextrix marka önekini yalnızca tekil bir T-Soft koduna ulaşıyorsa kaldırır."""
    kaynak_marka = metin_norm(marka)
    kaynak_kod = metin_norm(kod)

    # SAFAK-KM240 kablo kesme makasidir. T-Soft'taki KM 240 kodu KARDES
    # ELEKTRIK ek mufa aittir; kompakt kod benzerligiyle eslestirilmemeli.
    if (
        kaynak_marka == metin_norm("SAFAK ELEKTRIK")
        and re.sub(r"[^A-Z0-9]", "", kaynak_kod) == "SAFAKKM240"
    ):
        return ""
    if kaynak_marka == metin_norm("SIEMENS"):
        onayli_eslesmeler = {
            metin_norm(kaynak): metin_norm(hedef)
            for kaynak, hedef in ELEXTRIX_SIEMENS_ONAYLI_KOD_ESLESMELERI.items()
        }
        onayli_hedef = onayli_eslesmeler.get(kaynak_kod, "")
        if onayli_hedef in tsoft_kodlari:
            return onayli_hedef

    izinli_onekler = {
        metin_norm(onek)
        for onek in ELEXTRIX_MARKA_KOD_ONEKLERI.get(kaynak_marka, set())
    }
    bulunan_onek = next(
        (
            onek
            for onek in sorted(izinli_onekler, key=len, reverse=True)
            if kaynak_kod.startswith(f"{onek} ")
        ),
        "",
    )
    if not bulunan_onek:
        return ""

    oneiksiz_kod = kaynak_kod[len(bulunan_onek) :].strip()
    if oneiksiz_kod in tsoft_kodlari:
        return oneiksiz_kod

    kompakt_kod = re.sub(r"[^A-Z0-9]", "", oneiksiz_kod)
    adaylar = tsoft_kompakt_kod_indeksi.get(kompakt_kod, set())
    if len(adaylar) == 1:
        return next(iter(adaylar))

    if kaynak_marka == metin_norm("LEGRAND") and kompakt_kod.isdigit():
        sifirsiz_kod = kompakt_kod.lstrip("0") or "0"
        if sifirsiz_kod in tsoft_kodlari:
            return sifirsiz_kod
    return ""


def para_birimi_kodu(deger) -> str:
    kod = metin_norm(deger)
    return {
        "TL": "TRY",
        "TRL": "TRY",
        "TURK LIRASI": "TRY",
        "EURO": "EUR",
        "DOLAR": "USD",
        "US DOLLAR": "USD",
    }.get(kod, kod or "TRY")


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


def stok_degerini_oku(deger, varsayilan=None):
    """B2B stok metnini birim ve Turkce binlik ayiraciyla birlikte okur.

    Ornekler: ``28.000 AD.`` -> 28000, ``1.018 ADET`` -> 1018,
    ``2,5`` -> 2.5. Fiyatlar icin kullanilan genel ``ondalik`` fonksiyonundan
    ayridir; stok alanindaki tek nokta + uc hane binlik ayiraci kabul edilir.
    """
    if deger is None or pd.isna(deger):
        return varsayilan
    metin = str(deger).strip().replace("\xa0", " ")
    if not metin or metin.lower() == "nan":
        return varsayilan
    if any(ifade in metin_norm(metin) for ifade in ("YOK", "TUKENDI")):
        return Decimal("0")
    eslesme = re.search(r"-?\d+(?:[.,]\d+)*", metin)
    if not eslesme:
        return varsayilan
    sayi = eslesme.group(0)
    if "," in sayi and "." in sayi:
        if sayi.rfind(",") > sayi.rfind("."):
            sayi = sayi.replace(".", "").replace(",", ".")
        else:
            sayi = sayi.replace(",", "")
    elif "." in sayi:
        parcalar = sayi.split(".")
        if len(parcalar) > 1 and all(len(parca) == 3 for parca in parcalar[1:]):
            sayi = "".join(parcalar)
    elif "," in sayi:
        parcalar = sayi.split(",")
        if len(parcalar) > 1 and all(len(parca) == 3 for parca in parcalar[1:]):
            sayi = "".join(parcalar)
        else:
            sayi = sayi.replace(",", ".")
    try:
        return Decimal(sayi)
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


def varsayilan_rakip_ayarlari() -> list[dict]:
    """Rakipler klasöründeki üç rakip dosyasını dosya adından otomatik bulur."""
    desteklenenler = {".xlsx", ".xls", ".csv"}
    dosyalar = (
        [p for p in RAKIP_DIZINI.iterdir() if p.is_file() and p.suffix.lower() in desteklenenler]
        if RAKIP_DIZINI.exists()
        else []
    )
    sonuc = []
    kullanilanlar = set()
    for firma, anahtarlar in RAKIP_DOSYA_ESLESMELERI:
        bulunan = None
        for dosya in dosyalar:
            if dosya in kullanilanlar:
                continue
            dosya_adi = metin_norm(dosya.stem).replace(" ", "")
            if all(metin_norm(a).replace(" ", "") in dosya_adi for a in anahtarlar):
                bulunan = dosya
                break
        if bulunan:
            kullanilanlar.add(bulunan)
        sonuc.append(
            {
                "Firma": firma,
                "Dosya": bulunan or (RAKIP_DIZINI / f"{metin_norm(firma).replace(' ', '_')}.xlsx"),
                "KodKolonu": "",
                "StokKolonu": "",
                "ListeFiyatiKolonu": "",
                "SatisFiyatiKolonu": "",
                "IskontoKolonu": "",
                "NetSatisKolonu": "",
                "ParaBirimiKolonu": "",
                "VarsayilanParaBirimi": "TRY",
                "FiyatKDVTipi": "KDV DAHIL",
            }
        )
    return sonuc


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
        "rakipler": varsayilan_rakip_ayarlari(),
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
    aksiyon_ham = (
        row.get(kolonlar["aksiyon_fiyat"], "")
        if kolonlar["aksiyon_fiyat"]
        else ""
    )
    min_paket_ham = row.get(kolonlar["min_paket"], "") if kolonlar["min_paket"] else ""
    stok_ham = row.get(kolonlar["stok"], "") if kolonlar["stok"] else ""
    para_ham = row.get(kolonlar["para_birimi"], "") if kolonlar["para_birimi"] else ""
    aksiyon_fiyat = ondalik(aksiyon_ham)
    liste = ondalik(liste_ham)
    if (liste is None or liste <= 0) and (
        aksiyon_fiyat is None or aksiyon_fiyat <= 0
    ):
        raise ValueError("Liste fiyatı ve aksiyon fiyatı boş veya sıfır")

    if aksiyon_fiyat is not None and aksiyon_fiyat > 0:
        aksiyon_doviz = para_birimi_bul(aksiyon_ham, para_ham)
        aksiyon_kur = ayarlar["kurlar"].get(aksiyon_doviz)
        if aksiyon_kur is None:
            raise ValueError(f"{aksiyon_doviz} kuru ayarlanmamış")
        # Aksiyon fiyatı net alış fiyatıdır; alış iskontosu ikinci kez uygulanmaz.
        iskonto_yuzde = None
        net_tl_kdv_haric = (aksiyon_fiyat * aksiyon_kur).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        if liste is not None and liste > 0:
            doviz = para_birimi_bul(liste_ham, para_ham)
            kur = ayarlar["kurlar"].get(doviz)
            if kur is None:
                raise ValueError(f"{doviz} kuru ayarlanmamış")
            liste_tl = (liste * kur).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
        else:
            doviz = aksiyon_doviz
            liste_tl = net_tl_kdv_haric
    else:
        doviz = para_birimi_bul(liste_ham, para_ham)
        kur = ayarlar["kurlar"].get(doviz)
        if kur is None:
            raise ValueError(f"{doviz} kuru ayarlanmamış")
        liste_tl = (liste * kur).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        iskonto_yuzde = etkin_iskonto_yuzde(iskonto_ham)
        net_tl_kdv_haric = (liste * iskonto_carpani(iskonto_ham) * kur).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    min_paket = ondalik(min_paket_ham, Decimal("1")) or Decimal("1")
    stok = stok_degerini_oku(stok_ham)
    return {
        "liste_tl": liste_tl,
        "iskonto_yuzde": iskonto_yuzde,
        "net_tl_kdv_haric": net_tl_kdv_haric,
        "min_paket": min_paket,
        "stok": stok,
        "doviz": doviz,
        "aksiyon_fiyati_kullanildi": aksiyon_fiyat is not None and aksiyon_fiyat > 0,
    }


def wago_b2b_fiyat_satirlarini_cikar(eslesme: pd.DataFrame) -> pd.DataFrame:
    """WAGO SABİT dışındaki WAGO B2B fiyat kaynaklarını kaldırır."""
    sonuc = eslesme.copy()
    marka_kolonu = kolon_bul(sonuc.columns, ["B2B Marka", "Marka"])
    tedarikci_kolonu = kolon_bul(sonuc.columns, ["Tedarikci", "Tedarikçi"])
    if not marka_kolonu or not tedarikci_kolonu:
        return sonuc
    wago_b2b = sonuc[marka_kolonu].apply(metin_norm).eq("WAGO") & ~sonuc[
        tedarikci_kolonu
    ].apply(metin_norm).eq("WAGO SABIT")
    return sonuc[~wago_b2b].copy()


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
            stok_col = kolon_bul(
                df.columns,
                [
                    ayar["StokKolonu"],
                    "Stok Adedi",
                    "Stok Miktarı",
                    "Stok Miktari",
                    "Stok",
                    "Stok Durumu",
                ],
            )
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
                    "İndirimli Fiyat",
                    "Indirimli Fiyat",
                    "KDV Dahil Fiyat",
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
            marka_col = kolon_bul(df.columns, ["Marka", "Brand"])
            urun_adi_col = kolon_bul(
                df.columns,
                ["Ürün Adı", "Urun Adi", "Ürün İsmi", "Urun Ismi", "Product Name"],
            )
            onayli_eslesmeler = (
                {
                    metin_norm(kaynak): metin_norm(hedef)
                    for kaynak, hedef in ELEKTROMARKETIM_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("ELEKTROMARKETİM")
                else {}
            )
            kod_marka_onceligi = (
                {
                    metin_norm(kod): metin_norm(marka)
                    for kod, marka in ELEKTROMARKETIM_KOD_MARKA_ONCELIGI.items()
                }
                if firma_norm == metin_norm("ELEKTROMARKETİM")
                else {}
            )
            cata_onayli_eslesmeler = (
                {
                    metin_norm(kaynak): metin_norm(hedef)
                    for kaynak, hedef in ELEKTRIK_MARKET_CATA_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("ELEKTRİK MARKET")
                else {}
            )
            jupiter_onayli_eslesmeler = (
                {
                    metin_norm(kaynak): metin_norm(hedef)
                    for kaynak, hedef in ELEKTRIK_MARKET_JUPITER_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("ELEKTRİK MARKET")
                else {}
            )
            viko_onayli_eslesmeler = (
                {
                    metin_norm(kaynak): metin_norm(hedef)
                    for kaynak, hedef in ELEKTRIK_MARKET_VIKO_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("ELEKTRİK MARKET")
                else {}
            )
            star_onayli_eslesmeler = (
                {
                    (metin_norm(marka), metin_norm(kaynak)): metin_norm(hedef)
                    for (marka, kaynak), hedef in STAR_AKIM_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("STAR AKIM")
                else {}
            )
            tumpa_onayli_eslesmeler = (
                {
                    (metin_norm(marka), metin_norm(kaynak)): metin_norm(hedef)
                    for (marka, kaynak), hedef in TUMPA_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("TÜMPA ELEKTRİK")
                else {}
            )
            zeybek_onayli_eslesmeler = (
                {
                    (metin_norm(marka), metin_norm(kaynak)): metin_norm(hedef)
                    for (marka, kaynak), hedef in ZEYBEK_MARKET_ONAYLI_KOD_ESLESMELERI.items()
                }
                if firma_norm == metin_norm("ZEYBEK MARKET")
                else {}
            )
            tsoft_kodlari = ayarlar.get("tsoft_kodlari", set())
            tsoft_cata_kodlari = ayarlar.get("tsoft_cata_kodlari", set())
            tsoft_oznur_teknik_indeksi = ayarlar.get(
                "tsoft_oznur_teknik_indeksi", {}
            )
            tsoft_viko_kodlari = ayarlar.get("tsoft_viko_kodlari", set())
            tsoft_noas_kodlari = ayarlar.get("tsoft_noas_kodlari", set())
            tsoft_audio_sayisal_indeksi = ayarlar.get(
                "tsoft_audio_sayisal_indeksi", {}
            )
            tsoft_gwest_kodlari = ayarlar.get("tsoft_gwest_kodlari", set())
            tsoft_mutlusan_kodlari = ayarlar.get("tsoft_mutlusan_kodlari", set())
            tsoft_klemsan_kodlari = ayarlar.get("tsoft_klemsan_kodlari", set())
            tsoft_kraus_naimer_kodlari = ayarlar.get(
                "tsoft_kraus_naimer_kodlari", set()
            )
            tsoft_molwex_kodlari = ayarlar.get("tsoft_molwex_kodlari", set())
            tsoft_ortac_kodlari = ayarlar.get("tsoft_ortac_kodlari", set())
            tsoft_schneider_kodlari = ayarlar.get(
                "tsoft_schneider_kodlari", set()
            )
            tsoft_siemens_kodlari = ayarlar.get("tsoft_siemens_kodlari", set())
            tsoft_telemecanique_kodlari = ayarlar.get(
                "tsoft_telemecanique_kodlari", set()
            )
            tsoft_omron_kodlari = ayarlar.get("tsoft_omron_kodlari", set())
            tsoft_emas_kodlari = ayarlar.get("tsoft_emas_kodlari", set())
            tsoft_entes_kodlari = ayarlar.get("tsoft_entes_kodlari", set())
            tsoft_safak_kodlari = ayarlar.get("tsoft_safak_kodlari", set())
            tsoft_zeybek_marka_urunleri = ayarlar.get(
                "tsoft_zeybek_marka_urunleri", {}
            )
            tsoft_ucfaz_marka_urunleri = ayarlar.get(
                "tsoft_ucfaz_marka_urunleri", {}
            )
            tsoft_krc_marka_urunleri = ayarlar.get(
                "tsoft_krc_marka_urunleri", {}
            )
            tsoft_buyuk_kalpler_marka_urunleri = ayarlar.get(
                "tsoft_buyuk_kalpler_marka_urunleri", {}
            )
            tsoft_ortac_kompakt_indeksi = {}
            for tsoft_ortac_kodu in tsoft_ortac_kodlari:
                kompakt_ortac_kodu = re.sub(r"[^A-Z0-9]", "", tsoft_ortac_kodu)
                if kompakt_ortac_kodu:
                    tsoft_ortac_kompakt_indeksi.setdefault(
                        kompakt_ortac_kodu, set()
                    ).add(tsoft_ortac_kodu)
            elektrikciden_belirsiz_cata_kodlari = set()
            elektrikciden_belirsiz_audio_kodlari = set()
            if firma_norm == metin_norm("ELEKTRİKÇİDEN"):
                cata_kod_sayilari = {}
                audio_kod_sayilari = {}
                for _, cata_row in df.iterrows():
                    if (
                        marka_col
                        and metin_norm(cata_row.get(marka_col, ""))
                        == metin_norm("AUDIO")
                    ):
                        audio_urun_adi = str(
                            cata_row.get(urun_adi_col, "")
                            if urun_adi_col
                            else ""
                        ).upper()
                        audio_kodlari = {
                            (eslesme.lstrip("0") or "0")
                            for eslesme in re.findall(
                                r"(?<!\d)(\d{4,6})(?![A-Z0-9])",
                                audio_urun_adi,
                            )
                        }
                        for audio_kodu in audio_kodlari:
                            audio_kod_sayilari[audio_kodu] = (
                                audio_kod_sayilari.get(audio_kodu, 0) + 1
                            )
                    if (
                        marka_col
                        and metin_norm(cata_row.get(marka_col, ""))
                        != metin_norm("CATA")
                    ):
                        continue
                    cata_urun_adi = metin_norm(
                        cata_row.get(urun_adi_col, "") if urun_adi_col else ""
                    )
                    cata_adaylari = {
                        metin_norm(eslesme)
                        for eslesme in re.findall(
                            r"\bCT\s*\d+[A-Z]*\b", cata_urun_adi
                        )
                    } & tsoft_cata_kodlari
                    if len(cata_adaylari) == 1:
                        cata_kodu = next(iter(cata_adaylari))
                        cata_kod_sayilari[cata_kodu] = (
                            cata_kod_sayilari.get(cata_kodu, 0) + 1
                        )
                elektrikciden_belirsiz_cata_kodlari = {
                    cata_kodu
                    for cata_kodu, adet in cata_kod_sayilari.items()
                    if adet > 1
                }
                elektrikciden_belirsiz_audio_kodlari = {
                    audio_kodu
                    for audio_kodu, adet in audio_kod_sayilari.items()
                    if adet > 1
                }
            tsoft_kod_indeksi = {}
            tsoft_meanwell_indeksi = {}
            tsoft_kompakt_kod_indeksi = {}
            for tsoft_kodu in tsoft_kodlari:
                ilk_parca = tsoft_kodu.split()[0] if tsoft_kodu else ""
                if ilk_parca:
                    tsoft_kod_indeksi.setdefault(ilk_parca, set()).add(tsoft_kodu)
                meanwell_anahtari = meanwell_bicim_anahtari(tsoft_kodu)
                if meanwell_anahtari:
                    tsoft_meanwell_indeksi.setdefault(meanwell_anahtari, set()).add(
                        tsoft_kodu
                    )
                kompakt_kod = re.sub(r"[^A-Z0-9]", "", tsoft_kodu)
                if kompakt_kod:
                    tsoft_kompakt_kod_indeksi.setdefault(kompakt_kod, set()).add(
                        tsoft_kodu
                    )
            for i, row in df.iterrows():
                kod = metin_norm(row.get(kod_col, ""))
                if not kod:
                    continue
                kod_anahtarlari = {kod}
                rakip_marka = metin_norm(row.get(marka_col, "")) if marka_col else ""
                if (
                    firma_norm == metin_norm("ÜÇ FAZ")
                    and rakip_marka in tsoft_ucfaz_marka_urunleri
                ):
                    ucfaz_hedef_kodlari = tsoft_ucfaz_marka_urunleri[
                        rakip_marka
                    ].get("kodlar", set())
                    if kod not in ucfaz_hedef_kodlari:
                        # Uc Faz'da ayni kod farkli markalarda farkli urunleri
                        # gosterebilir. Dogrudan kodu yalniz rakip markasinin
                        # T-Soft urunleri arasinda varsa kabul et. Ornek:
                        # VIKO 92605513 telefon prizi, ARTLINE 92605513 VGA
                        # kapagiyla ayni urun degildir.
                        kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("KOZ AYDINLATMA")
                    and rakip_marka == metin_norm("ÖZEL ÜRETİM")
                ):
                    # Koz'un P-/PNYA/NYMK/TTRK kodları siteye özel ürün
                    # kimlikleridir; farklı markalı T-Soft ürünlerine bağlanmaz.
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("K. NAIMER")
                    and kod not in tsoft_kraus_naimer_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("MOLWEX")
                    and kod not in tsoft_molwex_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("ORTAÇ")
                    and kod not in tsoft_ortac_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("SCHNEIDER")
                    and kod not in tsoft_schneider_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("SIEMENS")
                    and kod not in tsoft_siemens_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("TELEMECANIQUE")
                    and kod not in tsoft_telemecanique_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("OMRON")
                    and kod not in tsoft_omron_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("EMAS")
                    and kod not in tsoft_emas_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("ENTES")
                    and kod not in tsoft_entes_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if (
                    firma_norm == metin_norm("ELEKTRİK MALZEMESİ")
                    and rakip_marka == metin_norm("ŞAFAK")
                    and kod not in tsoft_safak_kodlari
                ):
                    kod_anahtarlari.discard(kod)
                if firma_norm == metin_norm("ELEKTRİKÇİDEN"):
                    # Bu kaynaktaki "Ürün Kodu" rakip sitenin kendi iç kimliğidir.
                    # Yalnızca aşağıdaki doğrulanmış marka/model kuralları kullanılmalıdır.
                    kod_anahtarlari.discard(kod)
                if firma_norm == metin_norm("STAR AKIM") and rakip_marka:
                    marka_oneki = f"{rakip_marka} "
                    if kod.startswith(marka_oneki):
                        kod_anahtarlari.add(kod[len(marka_oneki) :].strip())
                    star_eslesmesi = star_onayli_eslesmeler.get((rakip_marka, kod))
                    if star_eslesmesi:
                        kod_anahtarlari.add(star_eslesmesi)
                    if rakip_marka == metin_norm("LEGRAND") and kod.isdigit():
                        sifirsiz_kod = kod.lstrip("0") or "0"
                        if sifirsiz_kod in tsoft_kodlari:
                            kod_anahtarlari.add(sifirsiz_kod)
                    if rakip_marka in {
                        metin_norm("HES"),
                        metin_norm("HES KABLO"),
                    }:
                        rakip_urun_adi = metin_norm(
                            row.get(urun_adi_col, "") if urun_adi_col else ""
                        )
                        if (
                            kod == metin_norm("H.002.004")
                            and "MAVI" in rakip_urun_adi
                            and re.search(r"(?:^| )0 50(?: |$)", rakip_urun_adi)
                        ):
                            kod_anahtarlari = {metin_norm("H.002.002")}
                        elif (
                            kod == metin_norm("H.104.007")
                            and "NHXMH" in rakip_urun_adi
                            and re.search(r"(?:^| )5X10(?: |$)", rakip_urun_adi)
                        ):
                            kod_anahtarlari = {metin_norm("H.104.022")}
                if (
                    firma_norm == metin_norm("ELEKTROMARKETİM")
                    and rakip_marka == metin_norm("MEAN WELL")
                    and kod.startswith("ELG ")
                    and not kod.endswith(" 3Y")
                ):
                    kod_anahtarlari.add(f"{kod} 3Y")
                if kod in onayli_eslesmeler:
                    kod_anahtarlari.add(onayli_eslesmeler[kod])
                if firma_norm == metin_norm("ELEKTROMARKETİM"):
                    if (
                        rakip_marka == metin_norm("GWEST")
                        and kod.startswith("GWEST")
                        and len(kod) > len("GWEST")
                    ):
                        gwest_web_servis_kodu = kod[len("GWEST") :].strip()
                        beklenen_marka = kod_marka_onceligi.get(
                            gwest_web_servis_kodu
                        )
                        if (
                            not beklenen_marka
                            and gwest_web_servis_kodu in tsoft_gwest_kodlari
                        ):
                            kod_anahtarlari.add(gwest_web_servis_kodu)
                    if rakip_marka == metin_norm("ÖZNUR KABLO"):
                        oznur_kodu = elektromarketim_oznur_tsoft_kodu(kod)
                        if oznur_kodu:
                            kod_anahtarlari.add(oznur_kodu)
                    if (
                        rakip_marka == metin_norm("SCHNEIDER ELECTRIC")
                        and re.fullmatch(r"ZB4B(?:V0|W3)\d+3", kod)
                    ):
                        kod_anahtarlari.add(kod[:-1])
                    if (
                        rakip_marka == metin_norm("BEMIS")
                        and kod.endswith("KZ")
                    ):
                        kod_anahtarlari.add(kod[:-2].strip())
                    if (
                        rakip_marka == metin_norm("ŞAFAK")
                        and re.fullmatch(r"UY \d+[ES]", kod)
                    ):
                        kod_anahtarlari.add(kod[:-1])
                if (
                    firma_norm == metin_norm("ELEKTRİK MARKET")
                    and rakip_marka == metin_norm("ACK")
                ):
                    ack_eslesmesi = re.fullmatch(r"(.+\d)[A-Z]{1,2}", kod)
                    if ack_eslesmesi:
                        kod_anahtarlari.add(ack_eslesmesi.group(1))
                if (
                    firma_norm == metin_norm("ELEKTRİK MARKET")
                    and rakip_marka == metin_norm("CATA")
                ):
                    if kod in cata_onayli_eslesmeler:
                        kod_anahtarlari.add(cata_onayli_eslesmeler[kod])
                    kompakt_kod = kod.replace(" ", "")
                    if kompakt_kod.startswith("CT") and len(kompakt_kod) > 2:
                        govde = kompakt_kod[2:]
                        cata_adaylari = {f"CT {govde}"}
                        if govde[-1].isalpha():
                            cata_adaylari.add(f"CT {govde[:-1]}")
                        if len(govde) > 1 and govde[-1].isalpha() and govde[-2].isalpha():
                            cata_adaylari.add(f"CT {govde[:-2]}")
                        mevcut_adaylar = cata_adaylari & tsoft_kodlari
                        if len(mevcut_adaylar) == 1:
                            kod_anahtarlari.update(mevcut_adaylar)
                if (
                    firma_norm == metin_norm("ELEKTRİK MARKET")
                    and rakip_marka == metin_norm("JUPITER")
                ):
                    jupiter_eslesmesi = (
                        jupiter_onayli_eslesmeler.get(kod)
                        or elektrik_market_jupiter_tsoft_kodu(
                            kod,
                            row.get(urun_adi_col, "") if urun_adi_col else "",
                            tsoft_kod_indeksi,
                        )
                    )
                    if jupiter_eslesmesi:
                        kod_anahtarlari.add(jupiter_eslesmesi)
                if (
                    firma_norm == metin_norm("ELEKTRİK MARKET")
                    and rakip_marka == metin_norm("MEANWELL")
                ):
                    meanwell_eslesmesi = elektrik_market_meanwell_tsoft_kodu(
                        kod,
                        tsoft_meanwell_indeksi,
                    )
                    if meanwell_eslesmesi:
                        kod_anahtarlari.add(meanwell_eslesmesi)
                if (
                    firma_norm == metin_norm("ELEKTRİK MARKET")
                    and rakip_marka == metin_norm("MUTLUSAN")
                ):
                    mutlusan_eslesmesi = elektrik_market_mutlusan_tsoft_kodu(
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_mutlusan_kodlari,
                    )
                    if mutlusan_eslesmesi:
                        kod_anahtarlari.add(mutlusan_eslesmesi)
                if (
                    firma_norm == metin_norm("ELEKTRİK MARKET")
                    and rakip_marka == metin_norm("VIKO BY PANASONIC")
                    and kod in viko_onayli_eslesmeler
                ):
                    kod_anahtarlari.add(viko_onayli_eslesmeler[kod])
                if firma_norm == metin_norm("ELEKTRIX"):
                    elextrix_eslesmesi = elextrix_tsoft_kodu(
                        rakip_marka,
                        kod,
                        tsoft_kodlari,
                        tsoft_kompakt_kod_indeksi,
                    )
                    if elextrix_eslesmesi:
                        kod_anahtarlari.add(elextrix_eslesmesi)
                if firma_norm == metin_norm("TÜMPA ELEKTRİK"):
                    tumpa_eslesmesi = tumpa_onayli_eslesmeler.get(
                        (rakip_marka, kod)
                    )
                    if tumpa_eslesmesi in tsoft_kodlari:
                        kod_anahtarlari.add(tumpa_eslesmesi)
                if firma_norm == metin_norm("ZEYBEK MARKET"):
                    zeybek_eslesmesi = (
                        zeybek_onayli_eslesmeler.get((rakip_marka, kod))
                        or zeybek_market_tsoft_kodu(
                            rakip_marka,
                            kod,
                            row.get(urun_adi_col, "") if urun_adi_col else "",
                            tsoft_zeybek_marka_urunleri,
                        )
                    )
                    if zeybek_eslesmesi in tsoft_kodlari:
                        kod_anahtarlari.add(zeybek_eslesmesi)
                if firma_norm == metin_norm("ÜÇ FAZ"):
                    ucfaz_eslesmesi = ucfaz_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_ucfaz_marka_urunleri,
                        tsoft_kodlari,
                        tsoft_meanwell_indeksi,
                    )
                    if ucfaz_eslesmesi:
                        kod_anahtarlari.add(ucfaz_eslesmesi)
                if firma_norm == metin_norm("AYDINLATMA DÜNYAM"):
                    aydinlatma_dunyam_eslesmesi = aydinlatma_dunyam_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_kodlari,
                        tsoft_oznur_teknik_indeksi,
                    )
                    if aydinlatma_dunyam_eslesmesi:
                        kod_anahtarlari.add(aydinlatma_dunyam_eslesmesi)
                if firma_norm == metin_norm("ÇİMEN ELEKTRİK"):
                    cimen_elektrik_eslesmesi = cimen_elektrik_tsoft_kodu(
                        rakip_marka,
                        kod,
                        tsoft_kodlari,
                    )
                    if cimen_elektrik_eslesmesi:
                        kod_anahtarlari.add(cimen_elektrik_eslesmesi)
                if firma_norm == metin_norm("KOZ AYDINLATMA"):
                    koz_aydinlatma_eslesmesi = koz_aydinlatma_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_cata_kodlari,
                        tsoft_noas_kodlari,
                    )
                    if koz_aydinlatma_eslesmesi:
                        kod_anahtarlari.add(koz_aydinlatma_eslesmesi)
                if firma_norm == metin_norm("KRC ELEKTRİK"):
                    krc_elektrik_eslesmesi = krc_elektrik_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_krc_marka_urunleri,
                    )
                    if krc_elektrik_eslesmesi:
                        kod_anahtarlari.add(krc_elektrik_eslesmesi)
                if firma_norm == metin_norm("BÜYÜK KALPLER"):
                    buyuk_kalpler_eslesmesi = buyuk_kalpler_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_buyuk_kalpler_marka_urunleri,
                        tsoft_kodlari,
                    )
                    if buyuk_kalpler_eslesmesi:
                        kod_anahtarlari.add(buyuk_kalpler_eslesmesi)
                if firma_norm == metin_norm("ELEKTRİK MALZEMESİ"):
                    elektrik_malzemesi_eslesmesi = elektrik_malzemesi_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_klemsan_kodlari,
                        tsoft_gwest_kodlari,
                        tsoft_kraus_naimer_kodlari,
                        tsoft_ortac_kompakt_indeksi,
                        tsoft_schneider_kodlari,
                        tsoft_kodlari,
                    )
                    if elektrik_malzemesi_eslesmesi:
                        kod_anahtarlari.add(elektrik_malzemesi_eslesmesi)
                if firma_norm == metin_norm("ELEKTRİKÇİDEN"):
                    elektrikciden_eslesmesi = elektrikciden_tsoft_kodu(
                        rakip_marka,
                        kod,
                        row.get(urun_adi_col, "") if urun_adi_col else "",
                        tsoft_cata_kodlari,
                        elektrikciden_belirsiz_cata_kodlari,
                        tsoft_oznur_teknik_indeksi,
                        tsoft_viko_kodlari,
                        tsoft_mutlusan_kodlari,
                        tsoft_siemens_kodlari,
                        tsoft_schneider_kodlari,
                        tsoft_noas_kodlari,
                        tsoft_audio_sayisal_indeksi,
                        elektrikciden_belirsiz_audio_kodlari,
                        tsoft_entes_kodlari,
                    )
                    if elektrikciden_eslesmesi:
                        kod_anahtarlari.add(elektrikciden_eslesmesi)
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
                liste = ondalik(row.get(liste_col, "")) if liste_col else None
                satis_tl = fiyat * kur
                if satis_col:
                    net_tl = satis_tl
                    if (
                        (iskonto_ham is None or not str(iskonto_ham).strip())
                        and liste is not None
                        and liste > 0
                    ):
                        iskonto_ham = (Decimal("1") - fiyat / liste) * Decimal("100")
                elif liste_col:
                    net_tl = satis_tl * iskonto_carpani(iskonto_ham)
                    satis_tl = net_tl
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
                for kod_anahtari in kod_anahtarlari:
                    beklenen_marka = FIYATLANDIRMA_KOD_MARKA_ONCELIGI.get(
                        metin_norm(kod_anahtari)
                    )
                    if (
                        beklenen_marka
                        and metin_norm(rakip_marka) != metin_norm(beklenen_marka)
                    ):
                        continue
                    anahtar = (firma_norm, kod_anahtari)
                    ayarlanmis_kayit = paket_fiyatini_birime_cevir(
                        kayit, ayarlar, kod_anahtari, firma_norm, "rakip"
                    )
                    onceki = harita.get(anahtar)
                    if (
                        onceki is None
                        or ayarlanmis_kayit["net_satis_kdv_haric"]
                        < onceki["net_satis_kdv_haric"]
                    ):
                        harita[anahtar] = ayarlanmis_kayit
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
    if isinstance(eslesme, (str, Path)):
        eslesme = pd.read_excel(eslesme, dtype=str).fillna("")
    if isinstance(tsoft, (str, Path)):
        tsoft = pd.read_excel(tsoft, dtype=str).fillna("")
    eslesme = eslesme.copy()
    eslesme_adaylari = {
        "ProductId": ["ProductId"],
        "Web Servis Kodu": ["Web Servis Kodu"],
        "B2B Marka": ["B2B Marka"],
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
    eslesme = wago_b2b_fiyat_satirlarini_cikar(eslesme)
    gerekli = list(eslesme_adaylari)
    eksik = [k for k in gerekli if k not in eslesme.columns]
    if eksik:
        raise ValueError(f"Eşleşme dosyasında eksik kolonlar: {', '.join(eksik)}")

    # Marka cakismasi nedeniyle tum B2B satirlari elense bile T-Soft urunu
    # raporda kalir ve ALIS FIYATI YOK olarak gorunur.
    urun_anahtarlari = eslesme[["ProductId", "Web Servis Kodu"]].drop_duplicates().copy()

    # Paket ayarinda marka belirtilmisse ayni urun+bayi icindeki kod cakismasinda
    # yalniz o markaya ait kaynak satirini kullan. Ornek: YILDIZ'da hem GWEST
    # 2008 hem NEAD 2008 ayni T-Soft urunune eslesebiliyor; NEAD paket ayari
    # GWEST fiyatina uygulanmamali ve iki marka ayni bayi ortalamasina karismamali.
    paket_alis_markalari = ayarlar.get("paket_birim", {}).get("alis_marka", {})
    if paket_alis_markalari and "B2B Marka" in eslesme.columns:
        tutulacak = pd.Series(True, index=eslesme.index)
        for (kod_norm, firma_norm), beklenen_marka in paket_alis_markalari.items():
            aday = (
                eslesme["Web Servis Kodu"].apply(metin_norm).eq(kod_norm)
                & eslesme["Tedarikci"].apply(metin_norm).eq(firma_norm)
            )
            marka_uyumlu = aday & eslesme["B2B Marka"].apply(metin_norm).apply(
                lambda gelen: bool(gelen) and beklenen_marka in gelen
            )
            if marka_uyumlu.any():
                tutulacak.loc[aday & ~marka_uyumlu] = False
        eslesme = eslesme[tutulacak].copy()

    # Eslesme dosyasi daha once uretilmis olsa bile dogrulanmis kisa kod/marka
    # cakismalarini rapor olusturulurken yeniden engelle.
    kod_norm = eslesme["Web Servis Kodu"].apply(metin_norm)
    marka_norm = eslesme["B2B Marka"].apply(metin_norm)
    beklenen_marka = kod_norm.map(FIYATLANDIRMA_KOD_MARKA_ONCELIGI)
    eslesme = eslesme[
        beklenen_marka.isna()
        | marka_norm.eq(beklenen_marka.fillna("").apply(metin_norm))
    ].copy()
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
            fiyat = paket_fiyatini_birime_cevir(
                fiyat,
                ayarlar,
                kod,
                tedarikci,
                "alis",
                row.get("B2B Marka", ""),
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
    for _, kimlik in urun_anahtarlari.iterrows():
        pid = str(kimlik["ProductId"]).strip()
        kod = str(kimlik["Web Servis Kodu"]).strip()
        ts = product_map.get(pid) or kod_map.get(metin_norm(kod), {})
        eslesen_satirlar = eslesme[
            (eslesme["ProductId"].astype(str).str.strip() == pid)
            & (eslesme["Web Servis Kodu"].astype(str).str.strip() == kod)
        ]
        marka = FIYATLANDIRMA_KOD_MARKA_ONCELIGI.get(metin_norm(kod), "")
        for kolon in ("B2B Marka", "Marka"):
            if marka:
                break
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
            "Stok": ts.get("Stok", ""),
            "T-SOFT AKTİFLİK": tsoft_aktiflik_durumu(
                ts.get("T-SOFT AKTİFLİK", ts.get("IsActive", ""))
            ),
            "Fiyat Birimi": ts.get("Fiyat Birimi", "TRY") or "TRY",
            "T-SOFT LİSTE FİYATI": ondalik(ts.get("T-SOFT LİSTE FİYATI", "")),
            "T-SOFT İNDİRİMLİ FİYATI": ondalik(
                ts.get("T-SOFT İNDİRİMLİ FİYATI", "")
            ),
        }
        for tedarikci in sorted(tedarikciler):
            urun[tedarikci] = fiyatlar.get((pid, tedarikci))
        urunler.append(urun)
    return urunler, sorted(tedarikciler), hatalar


def _formul_ortalama(kolonlar, satir_no):
    hucreler = [f"{get_column_letter(c)}{satir_no}" for c in kolonlar]
    return f'=IFERROR(AVERAGE({",".join(hucreler)}),"")' if hucreler else '=""'


def rakip_stokta_var_mi(stok) -> bool:
    sayi = ondalik(stok)
    if sayi is not None:
        return sayi > 0
    return metin_norm(stok) in {
        "VAR",
        "STOK VAR",
        "STOKTA",
        "STOKTA VAR",
        "MEVCUT",
        "IN STOCK",
    }


def alis_kaynagi_kullanilabilir_mi(
    fiyat: dict | None,
    *,
    tek_kaynak: bool = False,
) -> bool:
    """Ürünün kaynak sayısına göre kullanılabilir alış fiyatını belirler.

    Birden fazla kaynak varsa yalnız stoklu ve pozitif alış iskontolu kaynak
    hesaba girer. Ürün yalnız tek kaynakta bulunuyorsa stok olmasa da pozitif
    liste + pozitif iskonto ile hesaplanan net alış kullanılabilir. İskontosu
    yüzde 0/boş olan normal kayıt hiçbir durumda kullanılmaz. Açık aksiyon
    fiyatı ise mevcut davranışı korur ve yalnız stok varsa doğrudan net alış
    istisnası olarak kullanılır.
    """
    if not fiyat:
        return False
    net = ondalik(fiyat.get("net_tl_kdv_haric"))
    if net is None or net <= 0:
        return False
    iskonto = ondalik(fiyat.get("iskonto_yuzde"))
    stok = ondalik(fiyat.get("stok"))
    stoklu = stok is not None and stok > 0
    if fiyat.get("aksiyon_fiyati_kullanildi") is True:
        return stoklu
    liste = ondalik(fiyat.get("liste_tl"))
    pozitif_liste_iskonto = bool(
        liste is not None
        and liste > 0
        and iskonto is not None
        and iskonto > 0
    )
    return pozitif_liste_iskonto and (tek_kaynak or stoklu)


def alis_kaynagi_stokta_var_mi(fiyat: dict | None) -> bool:
    """Geriye uyumlu ad; kullanılabilir alış kaynağı kuralını uygular."""
    return alis_kaynagi_kullanilabilir_mi(fiyat)


def kategori_kar_oranlarini_oku(kategoriler: list[str]) -> dict:
    """Kategori bazlı karar kâr oranlarını ayar Excel'inden okur.

    Dosya yoksa kullanıcının doldurması için örnek şablon oluşturulur.
    Beklenen kolonlar: Kategori, Kar Orani
    """
    KATEGORI_KAR_ORANI_DOSYASI.parent.mkdir(exist_ok=True)
    benzersiz_kategoriler = sorted(
        {
            str(kategori).strip()
            for kategori in kategoriler
            if str(kategori).strip()
        }
    )
    if not KATEGORI_KAR_ORANI_DOSYASI.exists():
        pd.DataFrame(
            {
                "Kategori": benzersiz_kategoriler,
                "Kar Orani": ["" for _ in benzersiz_kategoriler],
            }
        ).to_excel(KATEGORI_KAR_ORANI_DOSYASI, index=False)
        print(
            "Kategori kâr oranı şablonu oluşturuldu; oranları doldurup raporu tekrar çalıştırın: "
            f"{KATEGORI_KAR_ORANI_DOSYASI}"
        )
        return {}

    kaynak_dosya = (
        KATEGORI_KAR_YEDEK_DOSYASI
        if KATEGORI_KAR_YEDEK_DOSYASI.exists()
        else KATEGORI_KAR_ORANI_DOSYASI
    )
    df = pd.read_excel(kaynak_dosya, dtype=str).fillna("")
    kategori_col = kolon_bul(df.columns, ["Kategori", "Ana Kategori Adı", "Ana Kategori Adi"])
    kar_col = kolon_bul(df.columns, ["Kar Orani", "Kâr Oranı", "Kar Oranı", "Kar %", "Kâr %"])
    if not kategori_col or not kar_col:
        raise ValueError(
            "kategori_kar_oranlari.xlsx içinde 'Kategori' ve 'Kar Orani' kolonları olmalı"
        )

    oranlar = {}
    for _, row in df.iterrows():
        kategori = str(row.get(kategori_col, "")).strip()
        oran = ondalik(row.get(kar_col, ""))
        if kategori and oran is not None:
            oranlar[metin_norm(kategori)] = float(oran)
    return oranlar


def uygulanan_kar_oranini_hesapla(kategori_hedefi, en_uygun_fiyat):
    """Fiyat Onay sayfasına yazılacak gerçek kâr tabanını hesaplar."""
    hedef = ondalik(kategori_hedefi)
    if hedef is None:
        return ""
    rakip = ondalik(en_uygun_fiyat)
    if rakip is not None and rakip > 0:
        if hedef in {Decimal("30"), Decimal("50")}:
            return float(hedef)
        return float(GENEL_MIN_KAR_ORANI)
    return float(hedef)


def paket_birim_ayarlari_sayfasi_ekle(wb, ayarlar: dict) -> None:
    """Kullanılan paket/birim kurallarını fiyat onay dosyasına görünür ayar sayfası olarak ekler."""
    ws = wb.create_sheet("Paket Birim Ayarları")
    ws.sheet_view.showGridLines = False
    basliklar = [
        "Aktif", "ProductId", "Marka", "Kategori", "Ürün Web Servis Kodu",
        "Firma", "Kaynak Türü", "Paket Adedi", "Fiyatın Kapsadığı Adet",
        "Kontrol Durumu",
    ]
    ws.merge_cells("A1:J1")
    ws["A1"] = "PAKET / BİRİM FİYAT AYARLARI"
    ws["A1"].fill = PatternFill("solid", fgColor="173F68")
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=14)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:J2")
    ws["A2"] = (
        "Yalnız fiyatın kapsadığı adet dönüşümüdür; zam, satış fiyatı veya "
        "iskonto sonucu içermez."
    )
    ws["A2"].fill = PatternFill("solid", fgColor="EAF2F8")
    ws["A2"].font = Font(color="244A6D", italic=True)
    ws.append(basliklar)
    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor="245986")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for kayit in ayarlar.get("paket_birim", {}).get("satirlar", []):
        ws.append([
            kayit.get("Aktif", ""), kayit.get("ProductId", ""),
            kayit.get("Marka", ""), kayit.get("Kategori", ""),
            kayit.get("Ürün Web Servis Kodu", ""), kayit.get("Firma", ""),
            kayit.get("Kaynak Türü", ""), float(kayit.get("Paket Adedi", 1)),
            float(kayit.get("Fiyatın Kapsadığı Adet", 1)),
            kayit.get("Kontrol Durumu", ""),
        ])
    genislikler = [9, 11, 18, 28, 22, 23, 13, 12, 20, 16]
    for index, genislik in enumerate(genislikler, 1):
        ws.column_dimensions[get_column_letter(index)].width = genislik
    for satir in range(4, ws.max_row + 1):
        for kolon in (1, 8, 9, 10):
            ws.cell(satir, kolon).fill = PatternFill("solid", fgColor="FFF2CC")
        ws.cell(satir, 8).number_format = "0"
        ws.cell(satir, 9).number_format = "0"
    ws.freeze_panes = "F4"
    ws.auto_filter.ref = f"A3:J{max(ws.max_row, 3)}"


def fiyat_onay_sayfasi_ekle(
    wb,
    index: dict,
    urunler: list[dict],
    tedarikciler: list[str],
    rakip_firmalar: list[str],
    rakip_haritasi: dict,
    son_satir: int,
) -> None:
    kategori_oranlari = kategori_kar_oranlarini_oku(
        [urun.get("Kategori", "") for urun in urunler]
    )

    kategori_ws = wb.create_sheet("Kategori Kar Oranları")
    kategori_ws.append(["Kategori", "Kar Orani"])
    kategoriler = sorted(
        {
            str(urun.get("Kategori", "")).strip()
            for urun in urunler
            if str(urun.get("Kategori", "")).strip()
        }
    )
    for kategori in kategoriler:
        kategori_ws.append(
            [kategori, kategori_oranlari.get(metin_norm(kategori), "")]
        )
    for cell in kategori_ws[1]:
        cell.fill = PatternFill("solid", fgColor="305496")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    kategori_ws.freeze_panes = "A2"
    kategori_ws.column_dimensions["A"].width = 38
    kategori_ws.column_dimensions["B"].width = 14
    for row in range(2, len(kategoriler) + 2):
        kategori_ws.cell(row, 2).number_format = "0.00"

    onay_ws = wb.create_sheet("Fiyat Onay", 1)
    onay_ws.sheet_view.showGridLines = False
    urun_bilgileri = [
        "ProductId",
        "Marka",
        "Kategori",
        "Ürün Web Servis Kodu",
        "Stok",
        "T-SOFT AKTİFLİK",
    ]
    hesap_bilgileri = [
        "EN İYİ ALIŞ B2B",
        "ORTALAMA ALIŞ",
        "ORTALAMA STOK",
        "ORTALAMA İSKONTO",
        "EN UYGUN RAKİP FİRMA",
        "RAKİP STOK DURUMU",
        "EN UYGUN SATIŞ FİYATI",
        "EN UYGUN İSKONTO",
    ]
    karar = [
        "KATEGORİ KÂR ORANI",
        "SATIŞ FİYATI",
        "REEL KÂR ORANI",
        "HESAPLANAN İSKONTO",
        "SEÇİLEN İSKONTO",
        "ONAY",
        "AKTARIM DURUMU",
    ]
    basliklar = urun_bilgileri + hesap_bilgileri + karar
    onay_index = {ad: i + 1 for i, ad in enumerate(basliklar)}

    gruplar = [
        ("ÜRÜN BİLGİLERİ", len(urun_bilgileri), "1F4E78"),
        ("ALIŞ / PİYASA ÖZETİ", len(hesap_bilgileri), "548235"),
        ("FİYATLANDIRMA ONAYI", len(karar), "305496"),
    ]
    baslangic = 1
    for grup_adi, uzunluk, renk in gruplar:
        bitis = baslangic + uzunluk - 1
        onay_ws.merge_cells(
            start_row=1,
            start_column=baslangic,
            end_row=1,
            end_column=bitis,
        )
        cell = onay_ws.cell(1, baslangic, grup_adi)
        cell.fill = PatternFill("solid", fgColor=renk)
        cell.font = Font(color="FFFFFF", bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        baslangic = bitis + 1

    for col, baslik in enumerate(basliklar, 1):
        cell = onay_ws.cell(2, col, baslik)
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ana_sayfa = "'Fiyatlandırma'"

    for kaynak_satir in range(3, son_satir + 1):
        hedef_satir = kaynak_satir
        urun = urunler[kaynak_satir - 3]
        for baslik in [
            "ProductId",
            "Marka",
            "Kategori",
            "Ürün Web Servis Kodu",
            "T-SOFT AKTİFLİK",
        ]:
            kaynak = f"{get_column_letter(index[baslik])}{kaynak_satir}"
            onay_ws.cell(
                hedef_satir,
                onay_index[baslik],
                f"={ana_sayfa}!{kaynak}",
            )
        onay_ws.cell(hedef_satir, onay_index["Stok"], urun.get("Stok", ""))

        netler = []
        stoklar = []
        iskonto_degerleri = []
        en_iyi_b2b = ""
        en_iyi_net = None
        urun_kaynak_sayisi = sum(
            1 for tedarikci in tedarikciler if urun.get(tedarikci)
        )
        tek_alis_kaynagi = urun_kaynak_sayisi == 1
        for tedarikci in tedarikciler:
            fiyat = urun.get(tedarikci)
            if not fiyat:
                continue
            if not alis_kaynagi_kullanilabilir_mi(
                fiyat,
                tek_kaynak=tek_alis_kaynagi,
            ):
                continue
            stok = fiyat.get("stok")
            if stok is not None:
                stoklar.append(float(stok))
            net = fiyat.get("net_tl_kdv_haric")
            if net is not None:
                net_float = float(net)
                netler.append(net_float)
                if en_iyi_net is None or net_float < en_iyi_net:
                    en_iyi_net = net_float
                    en_iyi_b2b = tedarikci
            iskonto = fiyat.get("iskonto_yuzde")
            if iskonto is not None:
                iskonto_degerleri.append(float(iskonto))

        onay_ws.cell(hedef_satir, onay_index["EN İYİ ALIŞ B2B"], en_iyi_b2b)
        ort_alis_ref = (
            f"{get_column_letter(index['MİN. ORT. NET ALIŞ'])}{kaynak_satir}"
        )
        ort_isk_ref = (
            f"{get_column_letter(index['MİN. ORT. İSKONTO'])}{kaynak_satir}"
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["ORTALAMA ALIŞ"],
            f"={ana_sayfa}!{ort_alis_ref}",
        )
        if stoklar:
            onay_ws.cell(
                hedef_satir,
                onay_index["ORTALAMA STOK"],
                sum(stoklar) / len(stoklar),
            )
        onay_ws.cell(
            hedef_satir,
            onay_index["ORTALAMA İSKONTO"],
            f"={ana_sayfa}!{ort_isk_ref}",
        )

        kod_norm = metin_norm(urun.get("Web Servis Kodu", ""))
        en_uygun_rakip_firma = ""
        en_uygun_rakip_stok = ""
        en_uygun_iskonto = ""
        en_uygun_fiyat = None
        herhangi_rakip_var = False
        for firma in rakip_firmalar:
            kayit = rakip_haritasi.get((metin_norm(firma), kod_norm))
            if kayit:
                herhangi_rakip_var = True
                stok_durum = "VAR" if rakip_stokta_var_mi(kayit.get("stok")) else "YOK"
                if (
                    stok_durum == "VAR"
                    and kayit.get("net_satis_kdv_haric") is not None
                    and kayit["net_satis_kdv_haric"] > 0
                    and (
                        en_uygun_fiyat is None
                        or float(kayit["net_satis_kdv_haric"]) < en_uygun_fiyat
                    )
                ):
                    en_uygun_fiyat = float(kayit["net_satis_kdv_haric"])
                    en_uygun_rakip_firma = firma
                    en_uygun_rakip_stok = "VAR"
                    # Yuzde 0 gercek bir degerdir; bos metne cevrilirse rakip
                    # iskontosu yok kuralinin denetlenmesi mumkun olmaz.
                    rakip_iskonto = kayit.get("iskonto")
                    en_uygun_iskonto = (
                        rakip_iskonto if rakip_iskonto is not None else ""
                    )
        if herhangi_rakip_var and not en_uygun_rakip_stok:
            en_uygun_rakip_stok = "YOK"

        en_uygun_ref = (
            f"{get_column_letter(index['EN DÜŞÜK RAKİP FİYATI'])}{kaynak_satir}"
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["EN UYGUN RAKİP FİRMA"],
            en_uygun_rakip_firma,
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["RAKİP STOK DURUMU"],
            en_uygun_rakip_stok,
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["EN UYGUN SATIŞ FİYATI"],
            f"={ana_sayfa}!{en_uygun_ref}",
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["EN UYGUN İSKONTO"],
            en_uygun_iskonto,
        )

        ort_alis = f"{get_column_letter(onay_index['ORTALAMA ALIŞ'])}{hedef_satir}"
        en_uygun = f"{get_column_letter(onay_index['EN UYGUN SATIŞ FİYATI'])}{hedef_satir}"
        en_uygun_isk = f"{get_column_letter(onay_index['EN UYGUN İSKONTO'])}{hedef_satir}"
        kategori_kar = f"{get_column_letter(onay_index['KATEGORİ KÂR ORANI'])}{hedef_satir}"
        satis = f"{get_column_letter(onay_index['SATIŞ FİYATI'])}{hedef_satir}"
        reel_kar = f"{get_column_letter(onay_index['REEL KÂR ORANI'])}{hedef_satir}"
        hesaplanan_isk = f"{get_column_letter(onay_index['HESAPLANAN İSKONTO'])}{hedef_satir}"
        secilen_isk = f"{get_column_letter(onay_index['SEÇİLEN İSKONTO'])}{hedef_satir}"
        onay = f"{get_column_letter(onay_index['ONAY'])}{hedef_satir}"
        aktiflik = f"{get_column_letter(onay_index['T-SOFT AKTİFLİK'])}{hedef_satir}"
        tsoft_liste_ref = (
            f"{get_column_letter(index['T-SOFT LİSTE FİYATI TL'])}{kaynak_satir}"
        )
        kategori_hedefi = kategori_oranlari.get(
            metin_norm(urun.get("Kategori", "")),
            "",
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["KATEGORİ KÂR ORANI"],
            uygulanan_kar_oranini_hesapla(kategori_hedefi, en_uygun_fiyat),
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["SATIŞ FİYATI"],
            (
                f'=IF(OR({ort_alis}="",{kategori_kar}=""),"",'
                f'IF(AND({en_uygun}<>"",{en_uygun}>={ort_alis}*(1+{kategori_kar}/100)),'
                f'{en_uygun},{ort_alis}*(1+{kategori_kar}/100)))'
            ),
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["REEL KÂR ORANI"],
            f'=IFERROR(({satis}/{ort_alis}-1)*100,"")',
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["HESAPLANAN İSKONTO"],
            f'=IFERROR((1-{satis}/{ana_sayfa}!{tsoft_liste_ref})*100,"")',
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["SEÇİLEN İSKONTO"],
            f'=IF({hesaplanan_isk}="","",{hesaplanan_isk})',
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["ONAY"],
            (
                f'=IF({aktiflik}<>"AKTİF","HAYIR",'
                f'IF(OR({ort_alis}="",{kategori_kar}="",{satis}="",'
                f'{reel_kar}="",{secilen_isk}="",'
                f'AND({en_uygun}<>"",OR({en_uygun_isk}="",{en_uygun_isk}<=0)),'
                f'{secilen_isk}<0,{secilen_isk}>=100,'
                f'{reel_kar}<{kategori_kar}),"HAYIR","EVET"))'
            ),
        )
        onay_ws.cell(
            hedef_satir,
            onay_index["AKTARIM DURUMU"],
            (
                f'=IF({aktiflik}="PASİF","PASİF ÜRÜN",'
                f'IF({aktiflik}<>"AKTİF","AKTİFLİK BİLGİSİ YOK",'
                f'IF({ort_alis}="","ALIŞ FİYATI YOK",'
                f'IF({kategori_kar}="","KATEGORİ KÂR ORANI YOK",'
                f'IF(AND({en_uygun}<>"",OR({en_uygun_isk}="",{en_uygun_isk}<=0)),"RAKİP İSKONTOSU YOK",'
                f'IF({secilen_isk}="","İSKONTO BEKLİYOR",'
                f'IF(OR({secilen_isk}<0,{secilen_isk}>=100),"İSKONTO HATALI",'
                f'IF({reel_kar}<{kategori_kar},"KÂR ORANI DÜŞÜK",'
                f'IF({onay}="EVET","AKTARIMA HAZIR","AKTARIMA UYGUN DEĞİL")))))))))'
            ),
        )

    hedef_son_satir = max(son_satir, 3)
    thin = Side(style="thin", color="D9E2F3")
    input_fill = PatternFill("solid", fgColor="FFD966")
    karar_fill = PatternFill("solid", fgColor="DDEBF7")
    for row in onay_ws.iter_rows(
        min_row=3,
        max_row=hedef_son_satir,
        min_col=1,
        max_col=len(basliklar),
    ):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="center")
            if cell.column >= onay_index["KATEGORİ KÂR ORANI"]:
                cell.fill = karar_fill
    for ad in ["SEÇİLEN İSKONTO"]:
        for row in range(3, hedef_son_satir + 1):
            onay_ws.cell(row, onay_index[ad]).fill = input_fill

    iskonto_dv = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="99.99",
        allow_blank=True,
    )
    onay_ws.add_data_validation(iskonto_dv)
    iskonto_col = get_column_letter(onay_index["SEÇİLEN İSKONTO"])
    iskonto_dv.add(f"{iskonto_col}3:{iskonto_col}{hedef_son_satir}")
    durum_col = get_column_letter(onay_index["AKTARIM DURUMU"])
    onay_ws.conditional_formatting.add(
        f"{durum_col}3:{durum_col}{hedef_son_satir}",
        FormulaRule(
            formula=[f'${durum_col}3="AKTARIMA HAZIR"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
        ),
    )
    onay_ws.conditional_formatting.add(
        f"{durum_col}3:{durum_col}{hedef_son_satir}",
        FormulaRule(
            formula=[
                f'OR(${durum_col}3="KÂR ORANI DÜŞÜK",'
                f'${durum_col}3="ALIŞ FİYATI YOK",'
                f'${durum_col}3="KATEGORİ KÂR ORANI YOK")'
            ],
            fill=PatternFill("solid", fgColor="FFC7CE"),
        ),
    )

    sayisal_basliklar = {
        "Stok",
        "ORTALAMA ALIŞ",
        "ORTALAMA STOK",
        "EN UYGUN SATIŞ FİYATI",
        "SATIŞ FİYATI",
    }
    yuzde_basliklar = {
        "ORTALAMA İSKONTO",
        "EN UYGUN İSKONTO",
        "KATEGORİ KÂR ORANI",
        "REEL KÂR ORANI",
        "HESAPLANAN İSKONTO",
        "SEÇİLEN İSKONTO",
    }
    for row in range(3, hedef_son_satir + 1):
        for baslik in sayisal_basliklar:
            onay_ws.cell(row, onay_index[baslik]).number_format = "#,##0.00"
        for baslik in yuzde_basliklar:
            onay_ws.cell(row, onay_index[baslik]).number_format = "0.00"

    genislik = {
        "ProductId": 12,
        "Marka": 18,
        "Kategori": 25,
        "Ürün Web Servis Kodu": 25,
        "Stok": 12,
        "T-SOFT AKTİFLİK": 18,
        "EN İYİ ALIŞ B2B": 18,
        "ORTALAMA ALIŞ": 16,
        "ORTALAMA STOK": 16,
        "ORTALAMA İSKONTO": 18,
        "EN UYGUN RAKİP FİRMA": 24,
        "RAKİP STOK DURUMU": 18,
        "EN UYGUN SATIŞ FİYATI": 20,
        "EN UYGUN İSKONTO": 18,
        "KATEGORİ KÂR ORANI": 20,
        "SATIŞ FİYATI": 18,
        "REEL KÂR ORANI": 18,
        "HESAPLANAN İSKONTO": 20,
        "SEÇİLEN İSKONTO": 18,
        "ONAY": 12,
        "AKTARIM DURUMU": 24,
    }
    for col, baslik in enumerate(basliklar, 1):
        onay_ws.column_dimensions[get_column_letter(col)].width = genislik.get(
            baslik, 14
        )
    onay_ws.freeze_panes = "F3"
    onay_ws.auto_filter.ref = (
        f"A2:{get_column_letter(len(basliklar))}{hedef_son_satir}"
    )
    onay_ws.row_dimensions[1].height = 26
    onay_ws.row_dimensions[2].height = 55


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
    paket_birim_ayarlari_sayfasi_ekle(wb, ayarlar)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "F3"
    rakip_analiz_ws = wb.create_sheet("_RakipAnaliz")
    rakip_analiz_ws.sheet_state = "veryHidden"
    for sutun, firma in enumerate(rakip_firmalar, 1):
        rakip_analiz_ws.cell(2, sutun, firma)

    sabit = [
        "ProductId",
        "Marka",
        "Kategori",
        "Ürün Web Servis Kodu",
        "Ürün Adı",
        "T-SOFT AKTİFLİK",
        "Fiyat Birimi",
    ]
    b2b = []
    for t in tedarikciler:
        b2b.extend(
            [
                f"{t} LİSTE",
                f"{t} İSKONTO",
                f"{t} NET ALIŞ KDV HARİÇ",
                f"{t} STOK",
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
    tsoft_fiyat = [
        "T-SOFT LİSTE FİYATI TL",
        "T-SOFT İNDİRİMLİ FİYATI TL",
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
    rekabet = [
        "BİZİM FİYATIMIZ KDV HARİÇ",
        "EN DÜŞÜK RAKİP FİYATI",
        "ORTALAMA RAKİP FİYATI",
        "EN YÜKSEK RAKİP FİYATI",
        "FİYAT KONUMUMUZ",
        "RAKİP FARK %",
    ]
    karar = [
        "SEÇİLEN İSKONTO",
        "KÂR KONTROLÜ",
        "ONAY",
        "AKTARIM DURUMU",
    ]
    basliklar = sabit + b2b + min_kar + tsoft_fiyat + hedef_kar + rakip + rekabet + karar
    index = {ad: i + 1 for i, ad in enumerate(basliklar)}

    gruplar = [
        ("ÜRÜN BİLGİLERİ", len(sabit), "1F4E78"),
        ("B2B TEDARİKÇİLER", len(b2b), "548235"),
        ("MİN. KÂR ORANI", len(min_kar), "BF9000"),
        ("T-SOFT LİSTE FİYATI", len(tsoft_fiyat), "4472C4"),
        ("HEDEFLENEN KÂR ORANI", len(hedef_kar), "7030A0"),
        ("RAKİP FİRMALAR", len(rakip), "C00000"),
        ("REKABET ANALİZİ", len(rekabet), "9C6500"),
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
        ws.cell(satir_no, index["T-SOFT AKTİFLİK"], urun["T-SOFT AKTİFLİK"])
        ws.cell(satir_no, index["Fiyat Birimi"], urun["Fiyat Birimi"])
        tsoft_liste = urun.get("T-SOFT LİSTE FİYATI")
        para_birimi = para_birimi_kodu(urun["Fiyat Birimi"])
        tsoft_kuru = ayarlar["kurlar"].get(para_birimi)
        if (
            tsoft_liste is not None
            and tsoft_liste > 0
            and tsoft_kuru is not None
            and tsoft_kuru > 0
        ):
            ws.cell(
                satir_no,
                index["T-SOFT LİSTE FİYATI TL"],
                float(tsoft_liste * tsoft_kuru),
            )
        tsoft_indirimli = urun.get("T-SOFT İNDİRİMLİ FİYATI")
        if (
            tsoft_indirimli is not None
            and tsoft_indirimli > 0
            and tsoft_kuru is not None
            and tsoft_kuru > 0
        ):
            ws.cell(
                satir_no,
                index["T-SOFT İNDİRİMLİ FİYATI TL"],
                float(tsoft_indirimli * tsoft_kuru),
            )
        for t in tedarikciler:
            f = urun.get(t)
            if not f:
                continue
            liste_col = index[f"{t} LİSTE"]
            isk_col = index[f"{t} İSKONTO"]
            net_haric_col = index[f"{t} NET ALIŞ KDV HARİÇ"]
            ws.cell(satir_no, liste_col, float(f["liste_tl"]))
            liste_ref = f"{get_column_letter(liste_col)}{satir_no}"
            isk_ref = f"{get_column_letter(isk_col)}{satir_no}"
            if f.get("aksiyon_fiyati_kullanildi"):
                # Aksiyon fiyatı doğrudan net alıştır; iskonto hücresi boş kalır
                # ve liste fiyatı üzerinden ikinci kez hesap yapılmaz.
                ws.cell(satir_no, isk_col, "")
                ws.cell(
                    satir_no,
                    net_haric_col,
                    float(f["net_tl_kdv_haric"]),
                )
            else:
                ws.cell(satir_no, isk_col, float(f["iskonto_yuzde"]))
                ws.cell(
                    satir_no,
                    net_haric_col,
                    f'=IFERROR({liste_ref}*(1-{isk_ref}/100),"")',
                )
            stok = f.get("stok")
            if stok is not None:
                ws.cell(satir_no, index[f"{t} STOK"], float(stok))
            ws.cell(satir_no, index[f"{t} MİN. PAKET"], float(f["min_paket"]))

        ort_isk_degerleri = []
        ort_net_degerleri = []
        ort_liste_degerleri = []
        ort_isk_sayimlari = []
        ort_net_sayimlari = []
        ort_liste_sayimlari = []
        urun_kaynak_sayisi = sum(1 for t in tedarikciler if urun.get(t))
        tek_alis_kaynagi = urun_kaynak_sayisi == 1
        for t in tedarikciler:
            liste_ref = f"{get_column_letter(index[f'{t} LİSTE'])}{satir_no}"
            isk_ref = f"{get_column_letter(index[f'{t} İSKONTO'])}{satir_no}"
            net_ref = (
                f"{get_column_letter(index[f'{t} NET ALIŞ KDV HARİÇ'])}{satir_no}"
            )
            stok_ref = f"{get_column_letter(index[f'{t} STOK'])}{satir_no}"
            # Tek kaynaklı üründe pozitif liste + pozitif iskonto stok olmasa da
            # kullanılabilir. Çok kaynaklı üründe stok şartı devam eder. İskontosu
            # boş ve net hücresi sabit sayı olan açık aksiyon fiyatı yalnız stok
            # varsa doğrudan net alış istisnasıdır.
            normal_kaynak_gecerli = (
                f"AND({liste_ref}>0,{isk_ref}>0)"
                if tek_alis_kaynagi
                else f"AND({stok_ref}>0,{liste_ref}>0,{isk_ref}>0)"
            )
            kaynak_gecerli = (
                f"AND(ISNUMBER({net_ref}),{net_ref}>0,OR("
                f"{normal_kaynak_gecerli},"
                f"AND({stok_ref}>0,{isk_ref}=\"\",NOT(ISFORMULA({net_ref})))))"
            )
            ort_isk_degerleri.append(f"IF({kaynak_gecerli},{isk_ref},0)")
            ort_net_degerleri.append(f"IF({kaynak_gecerli},{net_ref},0)")
            ort_liste_degerleri.append(f"IF({kaynak_gecerli},{liste_ref},0)")
            ort_isk_sayimlari.append(
                f"IF(AND({kaynak_gecerli},ISNUMBER({isk_ref})),1,0)"
            )
            ort_net_sayimlari.append(
                f"IF(AND({kaynak_gecerli},ISNUMBER({net_ref})),1,0)"
            )
            ort_liste_sayimlari.append(
                f"IF(AND({kaynak_gecerli},ISNUMBER({liste_ref})),1,0)"
            )
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
        for firma_sutun, firma in enumerate(rakip_firmalar, 1):
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
            if (
                rakip_stokta_var_mi(kayit.get("stok"))
                and kayit["net_satis_kdv_haric"] > 0
            ):
                rakip_analiz_ws.cell(
                    satir_no,
                    firma_sutun,
                    float(kayit["net_satis_kdv_haric"]),
                )

        secilen_iskonto = f"{get_column_letter(index['SEÇİLEN İSKONTO'])}{satir_no}"
        secilen = f"({ort_liste_formul}*(1-{secilen_iskonto}/100))"
        tsoft_liste_tl_ref = (
            f"{get_column_letter(index['T-SOFT LİSTE FİYATI TL'])}{satir_no}"
        )
        kar_kontrol_ref = (
            f"{get_column_letter(index['KÂR KONTROLÜ'])}{satir_no}"
        )
        secilen_tsoft_satis = (
            f"({tsoft_liste_tl_ref}*(1-{secilen_iskonto}/100))"
        )
        tsoft_indirimli_ref = (
            f"{get_column_letter(index['T-SOFT İNDİRİMLİ FİYATI TL'])}{satir_no}"
        )
        bizim_fiyat_ref = (
            f"{get_column_letter(index['BİZİM FİYATIMIZ KDV HARİÇ'])}{satir_no}"
        )
        son_rakip_analiz_kolonu = get_column_letter(len(rakip_firmalar))
        stoklu_rakip_aralik = (
            f"'_RakipAnaliz'!A{satir_no}:"
            f"{son_rakip_analiz_kolonu}{satir_no}"
        )
        min_rakip_ref = f"{get_column_letter(index['EN DÜŞÜK RAKİP FİYATI'])}{satir_no}"
        ort_rakip_ref = f"{get_column_letter(index['ORTALAMA RAKİP FİYATI'])}{satir_no}"
        maks_rakip_ref = f"{get_column_letter(index['EN YÜKSEK RAKİP FİYATI'])}{satir_no}"
        ws.cell(
            satir_no,
            index["BİZİM FİYATIMIZ KDV HARİÇ"],
            (
                f'=IF(AND({secilen_iskonto}<>"",{tsoft_liste_tl_ref}>0),'
                f'{secilen_tsoft_satis},IF({tsoft_indirimli_ref}>0,{tsoft_indirimli_ref},""))'
            ),
        )
        ws.cell(
            satir_no,
            index["EN DÜŞÜK RAKİP FİYATI"],
            (
                f'=IF(COUNT({stoklu_rakip_aralik})=0,"",'
                f'MIN({stoklu_rakip_aralik}))'
            ),
        )
        ws.cell(
            satir_no,
            index["ORTALAMA RAKİP FİYATI"],
            (
                f'=IF(COUNT({stoklu_rakip_aralik})=0,"",'
                f'AVERAGE({stoklu_rakip_aralik}))'
            ),
        )
        ws.cell(
            satir_no,
            index["EN YÜKSEK RAKİP FİYATI"],
            (
                f'=IF(COUNT({stoklu_rakip_aralik})=0,"",'
                f'MAX({stoklu_rakip_aralik}))'
            ),
        )
        ws.cell(
            satir_no,
            index["FİYAT KONUMUMUZ"],
            (
                f'=IF({bizim_fiyat_ref}="","BİZİM FİYAT YOK",'
                f'IF({ort_rakip_ref}="","STOKLU RAKİP FİYATI YOK",'
                f'IF({bizim_fiyat_ref}<={min_rakip_ref},"EN UCUZ",'
                f'IF({bizim_fiyat_ref}>={maks_rakip_ref},"EN PAHALI",'
                f'IF({bizim_fiyat_ref}<{ort_rakip_ref},"ORTALAMANIN ALTINDA",'
                f'"ORTALAMANIN ÜSTÜNDE")))))'
            ),
        )
        ws.cell(
            satir_no,
            index["RAKİP FARK %"],
            f'=IFERROR(({bizim_fiyat_ref}/{ort_rakip_ref}-1)*100,"")',
        )
        gercek_kar = f"(({secilen_tsoft_satis}/{min_ort}-1)*100)"
        ws.cell(
            satir_no,
            index["KÂR KONTROLÜ"],
            (
                f'=IF({secilen_iskonto}="","İSKONTO BEKLİYOR",'
                f'IF(OR({tsoft_liste_tl_ref}="",{tsoft_liste_tl_ref}<=0,'
                f'{min_ort}="",{min_ort}<=0),"KONTROL EDİLEMEDİ",'
                f'IF({gercek_kar}<{GENEL_MIN_KAR_ORANI},"KÂR %7 ALTI: %"&TEXT({gercek_kar},"0.00"),'
                f'"KÂR UYGUN: %"&TEXT({gercek_kar},"0.00"))))'
            ),
        )
        onay = f"{get_column_letter(index['ONAY'])}{satir_no}"
        aktiflik = f"{get_column_letter(index['T-SOFT AKTİFLİK'])}{satir_no}"
        ws.cell(
            satir_no,
            index["AKTARIM DURUMU"],
            (
                f'=IF({aktiflik}="PASİF","PASİF ÜRÜN",'
                f'IF({aktiflik}<>"AKTİF","AKTİFLİK BİLGİSİ YOK",'
                f'IF({min_ort}="","ALIŞ FİYATI YOK",'
                f'IF(OR({min_kar_ref}="",{hedef_kar_ref}=""),"KÂR ORANI BEKLİYOR",'
                f'IF({min_satis}>{hedef_satis},"KÂR ARALIĞI HATALI",'
                f'IF({secilen_iskonto}="","İSKONTO SEÇİMİ BEKLİYOR",'
                f'IF({kar_kontrol_ref}="KONTROL EDİLEMEDİ","KÂR KONTROL EDİLEMEDİ",'
                f'IF(LEFT({kar_kontrol_ref},11)="KÂR %7 ALTI","KÂR %7 ALTI",'
                f'IF(OR({secilen}<{min_satis},{secilen}>{hedef_satis}),"FİYAT ARALIK DIŞI",'
                f'IF({onay}="EVET","AKTARIMA HAZIR",'
                f'IF({onay}="HAYIR","REDDEDİLDİ","ONAY BEKLİYOR")))))))))))'
            ),
        )

    son_satir = max(len(urunler) + 2, 3)
    thin = Side(style="thin", color="D9E2F3")
    kaynak_fill = PatternFill("solid", fgColor="D9EAF7")
    b2b_fill = PatternFill("solid", fgColor="E2F0D9")
    min_fill = PatternFill("solid", fgColor="FFF2CC")
    hedef_fill = PatternFill("solid", fgColor="E4DFEC")
    rakip_fill = PatternFill("solid", fgColor="FCE4D6")
    rekabet_fill = PatternFill("solid", fgColor="FFF2CC")
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
            len(sabit) + len(b2b) + len(min_kar) + len(tsoft_fiyat),
            kaynak_fill,
        ),
        (
            len(sabit) + len(b2b) + len(min_kar) + len(tsoft_fiyat) + 1,
            len(sabit) + len(b2b) + len(min_kar) + len(tsoft_fiyat) + len(hedef_kar),
            hedef_fill,
        ),
        (
            len(sabit) + len(b2b) + len(min_kar) + len(tsoft_fiyat) + len(hedef_kar) + 1,
            len(sabit)
            + len(b2b)
            + len(min_kar)
            + len(tsoft_fiyat)
            + len(hedef_kar)
            + len(rakip),
            rakip_fill,
        ),
        (
            len(sabit)
            + len(b2b)
            + len(min_kar)
            + len(tsoft_fiyat)
            + len(hedef_kar)
            + len(rakip)
            + 1,
            len(sabit)
            + len(b2b)
            + len(min_kar)
            + len(tsoft_fiyat)
            + len(hedef_kar)
            + len(rakip)
            + len(rekabet),
            rekabet_fill,
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
            formula=[
                f'OR(${durum_col}3="FİYAT ARALIK DIŞI",'
                f'${durum_col}3="KÂR ARALIĞI HATALI",'
                f'${durum_col}3="KÂR %7 ALTI",'
                f'${durum_col}3="KÂR KONTROL EDİLEMEDİ")'
            ],
            fill=PatternFill("solid", fgColor="FFC7CE"),
        ),
    )
    kar_kontrol_col = get_column_letter(index["KÂR KONTROLÜ"])
    ws.conditional_formatting.add(
        f"{kar_kontrol_col}3:{kar_kontrol_col}{son_satir}",
        FormulaRule(
            formula=[f'LEFT(${kar_kontrol_col}3,9)="KÂR UYGUN"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
        ),
    )
    ws.conditional_formatting.add(
        f"{kar_kontrol_col}3:{kar_kontrol_col}{son_satir}",
        FormulaRule(
            formula=[f'LEFT(${kar_kontrol_col}3,11)="KÂR %7 ALTI"'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
        ),
    )

    for ad in basliklar:
        if any(k in ad for k in ["LİSTE", "NET ALIŞ", "SATIŞ"]):
            for r in range(3, son_satir + 1):
                ws.cell(r, index[ad]).number_format = '#,##0.00'
        if any(k in ad for k in ["İSKONTO", "KÂR ORANI", "FARK %"]):
            for r in range(3, son_satir + 1):
                ws.cell(r, index[ad]).number_format = '0.00'

    genislik = {
        "ProductId": 12,
        "Marka": 18,
        "Kategori": 25,
        "Ürün Web Servis Kodu": 25,
        "Ürün Adı": 52,
        "T-SOFT AKTİFLİK": 18,
        "Fiyat Birimi": 13,
        "KÂR KONTROLÜ": 24,
        "FİYAT KONUMUMUZ": 24,
        "RAKİP FARK %": 16,
        "ONAY": 12,
        "AKTARIM DURUMU": 24,
    }
    for c, ad in enumerate(basliklar, 1):
        ws.column_dimensions[get_column_letter(c)].width = genislik.get(ad, 18)
    ws.auto_filter.ref = f"A2:{get_column_letter(len(basliklar))}{son_satir}"
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 55

    fiyat_onay_sayfasi_ekle(
        wb=wb,
        index=index,
        urunler=urunler,
        tedarikciler=tedarikciler,
        rakip_firmalar=rakip_firmalar,
        rakip_haritasi=rakip_haritasi,
        son_satir=son_satir,
    )

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
        paket_ayar_kaynagi = PAKET_BIRIM_AYAR_DOSYASI
        if args.cikti.exists():
            try:
                with pd.ExcelFile(args.cikti) as onceki_rapor:
                    if "Paket Birim Ayarları" in onceki_rapor.sheet_names:
                        paket_ayar_kaynagi = args.cikti
            except Exception:
                paket_ayar_kaynagi = PAKET_BIRIM_AYAR_DOSYASI
        ayarlar["paket_birim"] = paket_birim_ayarlarini_oku(paket_ayar_kaynagi)
        print(
            "Paket/birim fiyat kuralı: "
            f"{len(ayarlar['paket_birim']['satirlar'])} firma kuralı"
        )
        print(
            "Kurlar: "
            f"USD={ayarlar['kurlar']['USD']} TL, "
            f"EUR={ayarlar['kurlar']['EUR']} TL "
            f"(TCMB tarihi: {ayarlar.get('kur_tarihi', '')})"
        )
        tsoft = pd.read_excel(args.tsoft, dtype=str).fillna("")
        eslesme = pd.read_excel(args.eslesme, dtype=str).fillna("")
        ayarlar["tsoft_kodlari"] = {
            metin_norm(kod) for kod in tsoft["Web Servis Kodu"] if metin_norm(kod)
        }
        ayarlar["tsoft_cata_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if metin_norm(row.get("Web Servis Kodu", ""))
            and (
                metin_norm(row.iloc[2]).startswith("CATA ")
                or re.fullmatch(
                    r"CT\s*\d+[A-Z0-9 ]*",
                    metin_norm(row.get("Web Servis Kodu", "")),
                )
            )
        }
        tsoft_oznur_teknik_indeksi = {}
        for _, row in tsoft.iterrows():
            if not metin_norm(row.iloc[2]).startswith("OZNUR "):
                continue
            tsoft_kodu = metin_norm(row.get("Web Servis Kodu", ""))
            tsoft_anahtari = elektrikciden_oznur_teknik_anahtar(row.iloc[2])
            if tsoft_kodu and tsoft_anahtari:
                tsoft_oznur_teknik_indeksi.setdefault(
                    tsoft_anahtari, set()
                ).add(tsoft_kodu)
        ayarlar["tsoft_oznur_teknik_indeksi"] = tsoft_oznur_teknik_indeksi
        ayarlar["tsoft_viko_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if metin_norm(row.iloc[2]).startswith("VIKO ")
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_noas_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if metin_norm(row.iloc[2]).startswith("NOAS ")
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        tsoft_audio_sayisal_indeksi = {}
        for _, row in tsoft.iterrows():
            if not metin_norm(row.iloc[2]).startswith("AUDIO "):
                continue
            tsoft_kodu = metin_norm(row.get("Web Servis Kodu", ""))
            if not tsoft_kodu:
                continue
            for audio_kod_eslesmesi in re.findall(r"\d{4,6}", tsoft_kodu):
                sayisal_kod = audio_kod_eslesmesi.lstrip("0") or "0"
                tsoft_audio_sayisal_indeksi.setdefault(
                    sayisal_kod, set()
                ).add(tsoft_kodu)
        ayarlar["tsoft_audio_sayisal_indeksi"] = (
            tsoft_audio_sayisal_indeksi
        )
        ayarlar["tsoft_gwest_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if metin_norm(row.get("Ürün Adı", "")).startswith("GWEST ")
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_mutlusan_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "MUTLUSAN" in metin_norm(row.get("Ürün Adı", ""))
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_klemsan_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "KLEMSAN" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_kraus_naimer_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if (
                "KRAUS" in metin_norm(row.iloc[2])
                or "NAIMER" in metin_norm(row.iloc[2])
            )
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_molwex_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "MOLWEX" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_ortac_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "ORTAC" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_schneider_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "SCHNEIDER" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_siemens_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "SIEMENS" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_telemecanique_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "TELEMECANIQUE" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_omron_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "OMRON" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_emas_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "EMAS" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_entes_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "ENTES" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        ayarlar["tsoft_safak_kodlari"] = {
            metin_norm(row["Web Servis Kodu"])
            for _, row in tsoft.iterrows()
            if "SAFAK" in metin_norm(row.iloc[2])
            and metin_norm(row.get("Web Servis Kodu", ""))
        }
        zeybek_marka_urunleri = {
            metin_norm(marka): [] for marka in ZEYBEK_TSOFT_MARKA_ADLARI
        }
        for _, row in tsoft.iterrows():
            tsoft_kodu = metin_norm(row.get("Web Servis Kodu", ""))
            tsoft_urun_adi_ham = str(row.get("Ürün Adı", ""))
            tsoft_urun_adi = metin_norm(tsoft_urun_adi_ham)
            if not tsoft_kodu or not tsoft_urun_adi:
                continue
            for marka, marka_adlari in ZEYBEK_TSOFT_MARKA_ADLARI.items():
                if any(
                    zeybek_ifade_iceriyor(tsoft_urun_adi, metin_norm(marka_adi))
                    for marka_adi in marka_adlari
                ):
                    zeybek_marka_urunleri[metin_norm(marka)].append(
                        (tsoft_kodu, tsoft_urun_adi)
                    )
        ayarlar["tsoft_zeybek_marka_urunleri"] = zeybek_marka_urunleri
        ucfaz_marka_urunleri = {
            metin_norm(marka): {
                "urunler": [],
                "ham_urunler": [],
                "kodlar": set(),
                "kompakt": {},
                "maks_parca": 1,
            }
            for marka in UCFAZ_TSOFT_MARKA_ADLARI
        }
        for _, row in tsoft.iterrows():
            tsoft_kodu = metin_norm(row.get("Web Servis Kodu", ""))
            tsoft_urun_adi = metin_norm(row.get("Ürün Adı", ""))
            if not tsoft_kodu or not tsoft_urun_adi:
                continue
            for marka, marka_adlari in UCFAZ_TSOFT_MARKA_ADLARI.items():
                if any(
                    zeybek_ifade_iceriyor(tsoft_urun_adi, metin_norm(marka_adi))
                    for marka_adi in marka_adlari
                ):
                    ucfaz_marka_urunleri[metin_norm(marka)]["urunler"].append(
                        (tsoft_kodu, tsoft_urun_adi, tsoft_urun_adi_ham)
                    )
        for marka_bilgisi in ucfaz_marka_urunleri.values():
            marka_bilgisi["kodlar"] = {
                kod for kod, _, _ in marka_bilgisi["urunler"]
            }
            for kod in marka_bilgisi["kodlar"]:
                kompakt = re.sub(r"[^A-Z0-9]", "", kod)
                if kompakt:
                    marka_bilgisi["kompakt"].setdefault(
                        kompakt,
                        set(),
                    ).add(kod)
                marka_bilgisi["maks_parca"] = max(
                    marka_bilgisi["maks_parca"],
                    len(kod.split()),
                )
        ayarlar["tsoft_ucfaz_marka_urunleri"] = ucfaz_marka_urunleri
        krc_marka_urunleri = {
            metin_norm(marka): {
                "urunler": [],
                "ham_urunler": [],
                "kodlar": set(),
                "kompakt": {},
                "maks_parca": 1,
            }
            for marka in KRC_TSOFT_MARKA_ADLARI
        }
        for _, row in tsoft.iterrows():
            tsoft_kodu = metin_norm(row.get("Web Servis Kodu", ""))
            tsoft_urun_adi = metin_norm(row.get("Ürün Adı", ""))
            if not tsoft_kodu or not tsoft_urun_adi:
                continue
            for marka, marka_adlari in KRC_TSOFT_MARKA_ADLARI.items():
                if any(
                    zeybek_ifade_iceriyor(
                        tsoft_urun_adi,
                        metin_norm(marka_adi),
                    )
                    for marka_adi in marka_adlari
                ):
                    krc_marka_urunleri[metin_norm(marka)]["urunler"].append(
                        (tsoft_kodu, tsoft_urun_adi)
                    )
                    krc_marka_urunleri[metin_norm(marka)]["ham_urunler"].append(
                        (
                            tsoft_kodu,
                            str(row.get("Ürün Adı", "")),
                        )
                    )
        for marka_bilgisi in krc_marka_urunleri.values():
            marka_bilgisi["kodlar"] = {
                kod for kod, _ in marka_bilgisi["urunler"]
            }
            for kod in marka_bilgisi["kodlar"]:
                kompakt = re.sub(r"[^A-Z0-9]", "", kod)
                if kompakt:
                    marka_bilgisi["kompakt"].setdefault(
                        kompakt,
                        set(),
                    ).add(kod)
                marka_bilgisi["maks_parca"] = max(
                    marka_bilgisi["maks_parca"],
                    len(kod.split()),
                )
        krc_entes_bilgisi = krc_marka_urunleri.get(
            metin_norm("ENTES"),
            {},
        )
        krc_entes_bilgisi["entes_model"] = (
            entes_fiyat_model_indeksi_olustur(
                ENTES_FIYAT_LISTESI_DOSYASI,
                ayarlar["tsoft_entes_kodlari"],
            )
        )
        krc_entes_bilgisi["entes_akim_trafosu"] = {}
        for tsoft_kodu, tsoft_urun_adi in krc_entes_bilgisi.get(
            "ham_urunler",
            [],
        ):
            trafo_anahtari = entes_tsoft_akim_trafosu_anahtari(
                tsoft_urun_adi
            )
            if trafo_anahtari:
                krc_entes_bilgisi["entes_akim_trafosu"].setdefault(
                    trafo_anahtari,
                    set(),
                ).add(tsoft_kodu)
        ayarlar["tsoft_krc_marka_urunleri"] = krc_marka_urunleri
        buyuk_kalpler_marka_urunleri = {
            metin_norm(marka): {
                "urunler": [],
                "ham_urunler": [],
                "kodlar": set(),
                "kompakt": {},
                "maks_parca": 1,
            }
            for marka in BUYUK_KALPLER_TSOFT_MARKA_ADLARI
        }
        for _, row in tsoft.iterrows():
            tsoft_kodu = metin_norm(row.get("Web Servis Kodu", ""))
            tsoft_urun_adi = metin_norm(row.get("Ürün Adı", ""))
            if not tsoft_kodu or not tsoft_urun_adi:
                continue
            for marka, marka_adlari in BUYUK_KALPLER_TSOFT_MARKA_ADLARI.items():
                if any(
                    zeybek_ifade_iceriyor(
                        tsoft_urun_adi,
                        metin_norm(marka_adi),
                    )
                    for marka_adi in marka_adlari
                ):
                    buyuk_kalpler_marka_urunleri[metin_norm(marka)][
                        "urunler"
                    ].append((tsoft_kodu, tsoft_urun_adi))
                    buyuk_kalpler_marka_urunleri[metin_norm(marka)][
                        "ham_urunler"
                    ].append((tsoft_kodu, str(row.get("Ürün Adı", ""))))
        for marka_bilgisi in buyuk_kalpler_marka_urunleri.values():
            marka_bilgisi["kodlar"] = {
                kod for kod, _ in marka_bilgisi["urunler"]
            }
            for kod in marka_bilgisi["kodlar"]:
                kompakt = re.sub(r"[^A-Z0-9]", "", kod)
                if kompakt:
                    marka_bilgisi["kompakt"].setdefault(
                        kompakt,
                        set(),
                    ).add(kod)
                marka_bilgisi["maks_parca"] = max(
                    marka_bilgisi["maks_parca"],
                    len(kod.split()),
                )
        ayarlar["tsoft_buyuk_kalpler_marka_urunleri"] = (
            buyuk_kalpler_marka_urunleri
        )
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
