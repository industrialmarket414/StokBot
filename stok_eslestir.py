
from pathlib import Path
from datetime import datetime
import re
import unicodedata
import pandas as pd
import os
import re


BASE_DIR = Path(__file__).parent

AYAR_FILE = BASE_DIR / "Ayarlar" / "tedarikci_ayar.xlsx"
SANTRAL_ESLESME_FILE = BASE_DIR / "Ayarlar" / "santral_model_eslesme.xlsx"

B2B_DIR = BASE_DIR / "B2B"
TSOFT_FILE = BASE_DIR / "Tsoft" / "tsoft_secili_alanlar.xlsx"
DIA_STOK_FILE = BASE_DIR / "DIA" / "DIA_STOK.xlsx"
SABIT_STOK_DIR = BASE_DIR / "SabitStok"

CIKTI_DIR = BASE_DIR / "Cikti"
LOG_DIR = BASE_DIR / "Log"

CIKTI_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# KIRAC ürün kodu alanına yanlışlıkla ürün modeli yazılan satırlar için
# Bu sürümde KIRAC düzeltmesi stok eşleştirmeden ÖNCE yapılır; eşleştirme düzeltilmiş kodla devam eder.
# ENTES fiyat listesinden Model -> Ürün Kodu dönüşümü yapılır.
FIYAT_LISTESI_DIR = BASE_DIR / "FiyatListesi"
KIRAC_DUZELTILMIS_FILE = CIKTI_DIR / "kirac_duzeltilmis.xlsx"
KIRAC_DUZELTME_RAPOR_FILE = CIKTI_DIR / "kirac_model_kod_duzeltme_raporu.xlsx"
KIRAC_DUZELTME_SAYILARI_FILE = CIKTI_DIR / "kirac_duzeltme_sayilari.xlsx"
KIRAC_DUZELTME_SAYILARI_GUNCEL_FILE = CIKTI_DIR / "kirac_duzeltme_sayilari_GUNCEL.xlsx"
KIRAC_SON_CALISMA_FILE = CIKTI_DIR / "kirac_son_calisma.txt"

# SANTRAL ürün kodu alanında ENTES ürün modeli/model numarası geldiği için
# manuel eşleşme yerine ENTES fiyat listesinden Model -> Ürün Kodu dönüşümü yapılır.
SANTRAL_DUZELTILMIS_FILE = CIKTI_DIR / "santral_duzeltilmis.xlsx"
SANTRAL_DUZELTME_RAPOR_FILE = CIKTI_DIR / "santral_model_kod_duzeltme_raporu.xlsx"
SANTRAL_DUZELTME_SAYILARI_FILE = CIKTI_DIR / "santral_duzeltme_sayilari.xlsx"
SANTRAL_DUZELTME_SAYILARI_GUNCEL_FILE = CIKTI_DIR / "santral_duzeltme_sayilari_GUNCEL.xlsx"
SANTRAL_SON_CALISMA_FILE = CIKTI_DIR / "santral_son_calisma.txt"

KOD_SURUMU = "DIA_SABIT_KONYA_OPAS_PREFIX_V127_2026_07_21"

# Kendi B2B aktariminda yalnizca tek marka bulunan tedarikciler. Bu kaynaklarda
# marka kolonu olmasa veya ayarda farkli bir kolon yazsa bile tum satirlar ilgili
# marka ile T-Soft Web Servis Kodu uzerinden eslestirilir.
SABIT_TEDARIKCI_MARKALARI = {
    "MEANWELL": "MEANWELL",
}

MEANWELL_TEDARIKCI_AYARI = {
    "Tedarikci": "MEANWELL",
    "Aktif": "1",
    "MarkaFiltresi": "MEANWELL",
    "DosyaUzantisi": ".xlsx",
    "KodKolonu": "Model No",
    "StokKolonu": "Stok",
    "MarkaKolonu": "__MARKA_FILTRESI__",
    "EslesmeTipi": "WEBSERVIS",
}

# Bu markalar tum bayilerde eslestirilir ve raporlanir; canli stok guncellemesinde
# ise yalnizca YILDIZ B2B kaynagindan gelen satirlar kullanilir.
YALNIZCA_YILDIZDAN_GUNCELLENECEK_MARKALAR = {
    "VIKO",
    "ARTLINE",
    "THEA",
    "ACK",
}

# Kalici olarak canli stok guncellemesinden cikartilan markalar. OZNUR tum
# kaynaklarda eslestirilip raporlanir, ancak canli guncellemeye hicbir sekilde gitmez.
GUNCELLEMEYE_HIC_GONDERILMEYECEK_MARKALAR = {
    "OZNUR",
    "OZNUR KABLO",
}

DIA_TOPLAMA_MARKALARI = {
    "WAGO": "WAGO",
    "ORTAC": "ORTAÇ",
    "EMAS": "EMAS",
    "OPAS": "OPAS",
    "MUCCO": "MUCCO",
}

SABIT_STOK_DOSYALARI = {
    "WAGO": {
        "dosya": SABIT_STOK_DIR / "Wago.xlsx",
        "kod_kolonlari": ["Web Servis Kodu"],
        "stok_kolonlari": ["Stok"],
    },
    "ORTAC": {
        "dosya": SABIT_STOK_DIR / "Ortaç.xlsx",
        "kod_kolonlari": ["Web Servis Kodu", "KOD"],
        "stok_kolonlari": ["Stok"],
    },
    "EMAS": {
        "dosya": SABIT_STOK_DIR / "Emas.xlsx",
        "kod_kolonlari": ["Web Servis Kodu"],
        # Mevcut dosyada yeni stok kolonunun basligi bos (Unnamed: 4).
        # Dosyada baslik Stok yapildiginda da ayni kural calisir.
        "stok_kolonlari": ["Stok", "Unnamed: 4"],
    },
}

# Ayni T-Soft Web Servis Kodu birden fazla gercek B2B markasinda kullanildiginda
# canli stokta yalnizca burada secilen markanin satiri hesaba katilir.
KOD_MARKA_ONCELIGI = {
    "15363": "LEGRAND",
    "28920": "SCHNEIDER ELECTRIC",
    "422060": "LEGRAND",
    "422037": "LEGRAND",
    "MK05": "ORTAC",
    "10521": "NADE",
    "770101": "LEGRAND",
    "2006": "GWEST",
    "2007": "GWEST",
    "2008": "GWEST",
    "2009": "GWEST",
    "4006": "GWEST",
    "4007": "GWEST",
    "4008": "GWEST",
    "4009": "GWEST",
    "4010": "GWEST",
    "4011": "GWEST",
    "KM 240": "KARDES ELEKTRIK",
    "OGS-60": "SAFAK ELEKTRIK",
    "2010": "GWEST",
    "3608": "GWEST",
    "3701": "GWEST",
    "3702": "GWEST",
    "4004": "GWEST",
    "7050": "GWEST",
    "RT424024": "SCHRACK",
    "RT424730": "SCHRACK",
    "PT78740": "SCHRACK",
    "RT78726": "SCHRACK",
    "XZCP1141L5": "TELEMECANIQUE SENSORS",
    "XCKN2145G11": "TELEMECANIQUE SENSORS",
    "XCSZ05": "TELEMECANIQUE SENSORS",
    "KE01": "KRAUS & NAIMER",
    "KE03": "KRAUS & NAIMER",
    "RT314730": "SCHRACK",
    "RT16016": "SCHRACK",
    "PT370024": "SCHRACK",
    "PT370110": "SCHRACK",
    "PT370730": "SCHRACK",
    "PT570524": "SCHRACK",
    "PTML0730": "SCHRACK",
    "PT17016": "SCHRACK",
    "PTML0024": "SCHRACK",
    "MT221024": "SCHRACK",
    "PT570730": "SCHRACK",
    "PT570024": "SCHRACK",
}

# YILDIZ B2B'deki OZN.001.xxx kodlari T-Soft kodunun kisaltilmis hali degildir;
# YILDIZ'in urun karti kodudur. Urun kartlarindaki teknik ad/kesit/renk bilgisi
# dogrulanarak T-Soft'un OZNUR KABLO web servis koduna cevrilir.
YILDIZ_OZNUR_KOD_ESLESMELERI = {
    "OZN.001.021": "O.1X4.Kirmizi.NYA",
    "OZN.001.031": "O.1X10.Sari-Yesil.NYA",
    "OZN.001.034": "O.1X16.Kahve.NYA",
    "OZN.001.037": "O.1X16.Siyah.NYA",
    "OZN.001.038": "O.1X16.Kirmizi.NYA",
    "OZN.001.049": "O.1X50.Kahve.NYA",
    "OZN.001.052": "O.1X50.Siyah.NYA",
    "OZN.001.246": "O.4X6.Siyah.TTR",
    "OZN.001.268": "O.NYM.2X1,5.Antigron",
    "OZN.001.274": "O.NYM.3X1,5.Antigron",
    "OZN.001.275": "O.NYM.3X2,5.Antigron",
    "OZN.001.276": "O.NYM.3X4.Antigron",
    "OZN.001.281": "O.NYM.4X1,5.Antigron",
    "OZN.001.282": "O.NYM.4X2,5.Antigron",
    "OZN.001.287": "O.NYM.5X1,5.Antigron",
    "OZN.001.288": "O.NYM.5X2,5.Antigron",
    "OZN.001.308": "O.NYY.1X185",
    "OZN.001.311": "O.NYY.2X1,5",
    "OZN.001.323": "O.NYY.3X6",
    "OZN.001.382": "O.NYFGBY.4X10",
    "OZN.001.437": "O.YXC7V-R.1X240+25",
    "OZN.001.441": "O.YAXC7V-R.1X95+16",
    "OZN.001.460": "O.1X1,5.Kahve.NYA.HFFR",
    "OZN.001.461": "O.1X1,5.Mavi.NYA.HFFR",
    "OZN.001.697": "O.NHXMH.4X6.Antigron",
    "OZN.001.751": "O.N2XH.3X240+120.HFFR",
}

# Santral ayni urun kodunu birden fazla aciklamada yayinladiginda, kullanicinin
# T-Soft'tan dogruladigi ProductId dogrudan kullanilir.
SANTRAL_KOD_PRODUCT_ID_ESLESME = {
    "4RB8083-3FC10": "27787",
}

# Yalnizca KONYA + BYLION icin satis fiyat listesi yerine kullanilan
# dogrudan model -> urun kodu tablosu.
KONYA_BYLION_MODEL_KOD_ESLESME = {
    "MMP-10000-F": "124-300-0000",
    "PMK-10000-F": "100-300-0000",
    "PMP-10000-F": "102-300-0000",
    "PSE-10000-F": "128-340-0000",
}

def metin_norm(metin):
    if pd.isna(metin):
        return ""

    metin = str(metin)
    metin = metin.replace("\xa0", " ")
    metin = metin.strip()
    metin = re.sub(r"\s+", " ", metin)

    ceviri = str.maketrans({
        "İ": "I",
        "ı": "I",
        "Ş": "S",
        "ş": "S",
        "Ğ": "G",
        "ğ": "G",
        "Ü": "U",
        "ü": "U",
        "Ö": "O",
        "ö": "O",
        "Ç": "C",
        "ç": "C",
    })

    metin = metin.translate(ceviri)
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(ch for ch in metin if not unicodedata.combining(ch))
    metin = metin.upper().strip()

    return metin


def marka_adi_standart(deger):
    norm = metin_norm(deger)
    gorunum = {
        "SAFAK": "ŞAFAK ELEKTRİK",
        "SAFAK ELEKTRIK": "ŞAFAK ELEKTRİK",
        "KARDES": "KARDEŞ ELEKTRİK",
        "KARDES ELEKTRIK": "KARDEŞ ELEKTRİK",
        "KRAUS NAIMER": "KRAUS & NAIMER",
        "KRAUS & NAIMER": "KRAUS & NAIMER",
        "TELEMECANIQUE": "TELEMECANIQUE SENSORS",
        "TELEMECANIQUE SENSORS": "TELEMECANIQUE SENSORS",
        "WEIDMULLER": "WEIDMÜLLER",
        "SCHNEIDER": "SCHNEIDER ELECTRIC",
        "VIKO-PANASONIC": "VIKO",
        "VIKO PANASONIC": "VIKO",
        "ORTAC": "ORTAÇ",
        "OPAS": "OPAŞ",
        "OZNUR": "ÖZNUR KABLO",
        "OZNUR KABLO": "ÖZNUR KABLO",
    }
    return gorunum.get(norm, norm)


def guncelleme_kaynagi_kurallarini_uygula(df):
    """Eslesmeleri koruyup yalnizca canli stok kaynagi kurallarini uygular."""
    if df is None:
        return pd.DataFrame()

    if df.empty or "Tedarikci" not in df.columns or "B2B Marka" not in df.columns:
        return df.copy()

    sonuc = df.copy()
    tedarikci_norm = sonuc["Tedarikci"].apply(metin_norm)
    marka_norm = sonuc["B2B Marka"].apply(
        lambda deger: metin_norm(marka_adi_standart(deger))
    )

    konya_hes_maskesi = tedarikci_norm.eq("KONYA") & marka_norm.eq("HES")
    yildiz_disi_kaynak_maskesi = (
        marka_norm.isin(YALNIZCA_YILDIZDAN_GUNCELLENECEK_MARKALAR)
        & ~tedarikci_norm.eq("YILDIZ")
    )
    tamamen_haric_marka_maskesi = marka_norm.isin(
        GUNCELLEMEYE_HIC_GONDERILMEYECEK_MARKALAR
    )
    # OZNUR canli stokta yalnizca kurali tanimli kaynaklardan kullanilir.
    oznur_marka_maskesi = marka_norm.isin({"OZNUR", "OZNUR KABLO"})
    oznur_izinli_kaynak_maskesi = tedarikci_norm.isin(
        {"KIRAC", "SANTRAL", "YILDIZ"}
    )
    oznur_izin_disindaki_kaynak_maskesi = (
        oznur_marka_maskesi & ~oznur_izinli_kaynak_maskesi
    )

    if "Web Servis Kodu" in sonuc.columns:
        web_servis_kodu_norm = sonuc["Web Servis Kodu"].apply(metin_norm)
        beklenen_marka = web_servis_kodu_norm.map(KOD_MARKA_ONCELIGI)
        kod_marka_onceligi_maskesi = (
            beklenen_marka.notna()
            & marka_norm.ne(beklenen_marka.fillna(""))
        )
    else:
        kod_marka_onceligi_maskesi = pd.Series(False, index=sonuc.index)

    sonuc["Guncelleme Durumu"] = "GONDERILECEK"
    sonuc.loc[tamamen_haric_marka_maskesi, "Guncelleme Durumu"] = "GONDERILMEYECEK"

    haric_maskesi = (
        konya_hes_maskesi
        | yildiz_disi_kaynak_maskesi
        | oznur_izin_disindaki_kaynak_maskesi
        | kod_marka_onceligi_maskesi
    )
    return sonuc[~haric_maskesi].copy()


def kod_temizle(kod):
    return metin_norm(kod)


def marka_uygun_mu(marka, marka_filtresi):
    # Bazi B2B dosyalarinda marka adindaki bosluk URL biciminde %20 gelebilir.
    marka = metin_norm(marka).replace("%20", " ")
    marka_filtresi = metin_norm(marka_filtresi)

    if not marka_filtresi:
        return True

    if not marka:
        return False

    viko_adlari = {"VIKO", "VIKO-PANASONIC", "VIKO PANASONIC"}
    if marka in viko_adlari and marka_filtresi in viko_adlari:
        return True

    kardes_adlari = {"KARDES", "KARDES ELEKTRIK"}
    if marka in kardes_adlari and marka_filtresi in kardes_adlari:
        return True

    kraus_adlari = {"KRAUS NAIMER", "KRAUS & NAIMER"}
    if marka in kraus_adlari and marka_filtresi in kraus_adlari:
        return True

    telemecanique_adlari = {"TELEMECANIQUE", "TELEMECANIQUE SENSORS"}
    if marka in telemecanique_adlari and marka_filtresi in telemecanique_adlari:
        return True

    safak_adlari = {"SAFAK", "SAFAK ELEKTRIK"}
    if marka in safak_adlari and marka_filtresi in safak_adlari:
        return True

    return marka_filtresi in marka


def entes_prefix_sil(kod):
    kod = kod_temizle(kod)

    temiz = kod
    temiz = re.sub(r"^ENTES[\.\-\s_]*", "", temiz).strip()
    temiz = re.sub(r"^ENT[\.\-\s_]*", "", temiz).strip()

    return temiz


def ack_prefix_sil(kod):
    """Yalnizca ACK markasinda urun kodunun basindaki ACK on ekini siler."""
    kod = kod_temizle(kod)
    return re.sub(r"^ACK[\.\-\s_]+", "", kod).strip()


def bemis_prefix_sil(kod):
    """Yalnizca BEMIS markasinda urun kodunun basindaki BEM on ekini siler."""
    kod = kod_temizle(kod)
    return re.sub(r"^BEM[\.\-\s_]+", "", kod).strip()


def bylion_model_urun_koduna_cevir(kod):
    """Konya BYLION on ekini siler ve bilinen modeli dogrudan urun koduna cevirir."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^BYLION[\.\-\s_]+", "", temiz).strip()
    return KONYA_BYLION_MODEL_KOD_ESLESME.get(temiz, temiz)


def cata_urun_kodu_temizle(kod):
    """CATA marka on ekini ve sondaki harften onceki gereksiz tireyi siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^CATA[\.\-\s_]+", "", temiz).strip()
    temiz = re.sub(r"(?<=\d)-(?=[A-Z]+$)", "", temiz)
    return temiz


def cetinkaya_urun_kodu_temizle(kod):
    """CETINKAYA urun kodunun basindaki CETIN marka on ekini siler."""
    if pd.isna(kod):
        return ""

    temiz = str(kod).replace("\xa0", " ").strip()
    temiz = re.sub(
        r"^(?:ÇETİN|CETIN)[\.\-\s_]+",
        "",
        temiz,
        flags=re.IGNORECASE,
    )
    return temiz.strip()


def emas_urun_kodu_temizle(kod):
    """Yalnizca EMAS markasinda urun kodunun basindaki EMAS on ekini siler."""
    kod = kod_temizle(kod)
    return re.sub(r"^EMAS[\.\-\s_]+", "", kod).strip()


def kardes_urun_kodu_temizle(kod):
    """KARDES urun kodunun basindaki U.KRD on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^(?:U[\.\-\s_]*)?KRD[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def klemsan_urun_kodu_temizle(kod):
    """KLEMSAN urun kodunun basindaki KLM on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^KLM[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def kraus_naimer_urun_kodu_temizle(kod):
    """KRAUS NAIMER urun kodunun basindaki U.KRAU on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^(?:U[\.\-\s_]*)?KRAU[\.\-\s_]+", "", temiz).strip()


def konya_legrand_urun_koduna_cevir(kod):
    """Konya LEGRAND on ekini ve urun kodunun basindaki sifirlari siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^LEGR[\.\-\s_]+", "", temiz).strip()
    return bas_sifir_sil(temiz)


def mete_urun_kodu_temizle(kod):
    """METE urun kodunun basindaki METE on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^METE[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def mutlusan_urun_kodu_temizle(kod):
    """MUTLUSAN urun kodunun basindaki MUT on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^MUT[\.\-\s_]+", "", temiz).strip()


def nade_urun_kodu_temizle(kod):
    """NADE urun kodunun basindaki NADE marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^NADE[\.\-\s_]+", "", temiz).strip()


def nead_urun_kodu_temizle(kod):
    """NEAD urun kodunun basindaki NEAD marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^NEAD[\.\-\s_]+", "", temiz).strip()


def omron_urun_kodu_temizle(kod):
    """OMRON urun kodunun basindaki OMR marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^OMR[\.\-\s_]+", "", temiz).strip()


def opas_konya_urun_kodu_temizle(kod):
    """Konya OPAŞ kodundan OPAS marka on ekini ve para birimini siler.

    Ornek: ``OPAS. IPE3132`` -> ``IPE3132``. OPAŞ kataloglarindaki
    ``IPE 3132`` ile ``IPE 3 1A1K32`` farkli urunler oldugu icin urun kodunun
    kalan bolumu degistirilmez; yalnizca bayi on eki temizlenir.
    """
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^OPAS[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def ortac_konya_urun_kodu_temizle(kod):
    """Konya ORTAC urun kodunun basindaki ORT on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^ORT[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def osram_urun_kodu_temizle(kod):
    """OSRAM urun kodunun basindaki AMP marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^AMP[\.\-\s_]+", "", temiz).strip()


def philips_urun_kodu_temizle(kod):
    """PHILIPS urun kodunun basindaki AMP marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^AMP[\.\-\s_]+", "", temiz).strip()


def phoenix_contact_urun_kodu_temizle(kod):
    """PHOENIX CONTACT urun kodunun basindaki PHOE marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^PHOE[\.\-\s_]+", "", temiz).strip()


def schneider_konya_urun_kodu_temizle(kod):
    """Konya SCHNEIDER kodlarindaki S.TM/S.ETKS/S.HC/S.MG on eklerini siler."""
    temiz = kod_temizle(kod)
    eslesme = re.match(r"^S[\.\-\s_]*(TM|ETKS|HC|MG)[\.\-\s_]*", temiz)

    if eslesme:
        on_ek = eslesme.group(1)
        temiz = temiz[eslesme.end():].strip()

        # S.ETKS kaldirildiktan sonra kalan EPH de bu grupta marka on ekidir.
        if on_ek == "ETKS":
            temiz = re.sub(r"^EPH[\.\-\s_]*", "", temiz).strip()

    return para_birimi_temizle(temiz)


def siemens_konya_urun_kodu_temizle(kod):
    """Konya SIEMENS urun kodunun basindaki SIE on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^SIE[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def safak_konya_urun_kodu_temizle(kod):
    """Konya SAFAK urun kodunun basindaki SFK on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^SFK[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def telemecanique_konya_urun_kodu_temizle(kod):
    """Konya TELEMECANIQUE kodunun basindaki TM on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^TM[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def viko_konya_urun_kodu_temizle(kod):
    """Konya VIKO urun kodunun basindaki VIKO marka on ekini siler."""
    temiz = kod_temizle(kod)
    return re.sub(r"^VIKO[\.\-\s_]+", "", temiz).strip()


def wago_konya_urun_kodu_temizle(kod):
    """Konya WAGO kodunun basindaki WAGO on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^WAGO[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def weidmuller_konya_urun_kodu_temizle(kod):
    """Konya WEIDMULLER kodunun basindaki WDM on ekini ve para birimini siler."""
    temiz = kod_temizle(kod)
    temiz = re.sub(r"^WDM[\.\-\s_]+", "", temiz).strip()
    return para_birimi_temizle(temiz)


def woer_konya_urun_kodu_temizle(kod):
    """Konya WOER sayisal kodunu WOER-W-<sayi> T-Soft bicimine cevirir."""
    temiz = para_birimi_temizle(kod_temizle(kod))
    eslesme = re.fullmatch(r"WOER[\.\-\s_]+(\d+)", temiz)

    if not eslesme:
        return temiz

    sayi = eslesme.group(1).lstrip("0") or "0"
    return f"WOER-W-{sayi}"


def para_birimi_temizle(kod):
    kod = kod_temizle(kod)

    # Kod veya sayi sabit degildir; yalnizca sondaki para birimi temizlenir.
    # Ornek: ENT. M2084-€ -> M2084, ENT. X123-$ -> X123
    kod = re.sub(
        r"(?:[\-\._\s]*(?:€|\$|₺|£|EURO|EUR|USD|DOLAR|TL|GBP|STERLIN))+\s*$",
        "",
        kod,
    ).strip()

    kod = re.sub(r"[\-\._\s]+$", "", kod)
    kod = kod.strip()

    return kod

def bas_sifir_sil(kod):
    """
    Sadece KIRAC + LEGRAND için kullanılacak.
    Baştaki sıfırları siler.

    Örnek:
    004280  -> 4280
    04280   -> 4280
    0004280 -> 4280
    000     -> 0
    """
    kod = kod_temizle(kod)

    if not kod:
        return ""

    temiz = re.sub(r"^0+", "", kod).strip()

    if temiz == "":
        return "0"

    return temiz


def bosch_bosluklu_kod(kod):
    """
    KIRAC + BOSCH icin kullanilir.
    B2B'den bosluksuz gelen Bosch kodunu T-Soft web servis kodu formatina cevirir.

    Ornek:
    06019G8100 -> 0 601 9G8 100
    """
    temiz = sade_kod(kod)

    if len(temiz) == 10 and temiz.startswith("0"):
        return f"{temiz[0]} {temiz[1:4]} {temiz[4:7]} {temiz[7:10]}"

    return kod_temizle(kod)


def ortac_ospa_bosluklu_kod(kod):
    """
    KIRAC + ORTAC icin kullanilir.
    OSPA sonrasinda gelen sayidan once T-Soft web servis kodundaki boslugu ekler.

    Ornek:
    OSPA11 -> OSPA 11
    OSPA37 -> OSPA 37
    """
    temiz = sade_kod(kod)
    eslesme = re.fullmatch(r"(OSPA)(\d+)", temiz)

    if eslesme:
        return f"{eslesme.group(1)} {eslesme.group(2)}"

    return kod_temizle(kod)


def sade_kod(kod):
    kod = kod_temizle(kod)
    kod = re.sub(r"[^A-Z0-9]", "", kod)
    return kod


def safak_km_bosluklu_kod(kod):
    """
    Sadece KIRAC + SAFAK ELEKTRIK icin kullanilir.
    KIRAC B2B'de tireli gelen KM kodunu T-Soft'taki
    ``KM <sayı>`` web servis kodu biçimine çevirir. Sayı sabit değildir.

    Örnekler:
    KM-240  -> KM 240
    KM-250  -> KM 250
    KM300   -> KM 300
    KM 480  -> KM 480
    """
    raw = kod_temizle(kod)
    if not raw:
        return ""

    eslesme = re.fullmatch(r"KM[\s._-]*(\d+)", raw)
    if eslesme:
        return f"KM {eslesme.group(1)}"

    return raw


def kirac_oznurkablo_tsoft_koduna_cevir(kod):
    """KIRAC + OZNUR urun kodunu T-Soft Web Servis Kodu bicimine cevirir.

    Kural tek tek urun ezberlemez; kablo ailesi, damar/kesit ve renk bilgisini
    KIRAC kodundan ayirir. T-Soft'ta olmayan hedefler daha sonra normal eslesme
    kontrolunde eslesmez ve guncelleme listesine alinmaz.
    """
    raw = kod_temizle(kod).replace(" ", "")
    if not raw:
        return ""

    # KIRAC'taki paket/mukerrer satir ekleri urunun teknik koduna ait degildir.
    raw = re.sub(r"-MK\d+$", "", raw)
    raw = re.sub(r"PRC\d+$", "", raw)

    def olcu_duzelt(deger, arti_eki=False):
        sonuc = str(deger).replace(".", ",").upper()
        if arti_eki:
            sonuc = re.sub(r"-(?=\d+$)", "+", sonuc)
        return sonuc

    def renk_adi(deger):
        renk = sade_kod(deger)
        return {
            "SY": "Sari-Yesil",
            "SARIYESIL": "Sari-Yesil",
            "S": "Siyah",
            "SIYAH": "Siyah",
            "M": "Mavi",
            "MAVI": "Mavi",
            "K": "Kirmizi",
            "KIRMIZI": "Kirmizi",
            "KA": "Kahve",
            "KAHVE": "Kahve",
            "B": "Beyaz",
            "BEYAZ": "Beyaz",
            "G": "Gri",
            "GRI": "Gri",
            "SARI": "Sari",
            "TURUNCU": "Turuncu",
            "YESIL": "Yesil",
        }.get(renk, "")

    # H07Z1-R / H07Z1-U -> NYA.HFFR, H07Z1-K -> NYAF.HFFR.
    eslesme = re.fullmatch(
        r"H?07Z1-?([RKU])-?([0-9]+(?:[,.][0-9]+)?)-(.+)",
        raw,
    )
    if eslesme:
        iletken_tipi, kesit, renk_raw = eslesme.groups()
        renk = renk_adi(renk_raw)
        if not renk:
            return ""
        aile = "NYAF" if iletken_tipi == "K" else "NYA"
        return f"O.1X{olcu_duzelt(kesit)}.{renk}.{aile}.HFFR"

    # NYA / NYAF: NYAF1,5-KIRMIZI -> O.1X1,5.Kirmizi.NYAF
    eslesme = re.fullmatch(
        r"(NYAF|NYA)-?([0-9]+(?:[,.][0-9]+)?)-?(.+)?",
        raw,
    )
    if eslesme:
        aile, kesit, renk_raw = eslesme.groups()
        renk_raw = re.sub(r"-1$", "", renk_raw or "")
        renk = renk_adi(renk_raw)
        if not renk:
            return ""
        return f"O.1X{olcu_duzelt(kesit)}.{renk}.{aile}"

    # NYM: NYM2X1,5 -> O.NYM.2X1,5.Antigron
    eslesme = re.fullmatch(r"NYM-?([0-9]+X[0-9]+(?:[,.][0-9]+)?)", raw)
    if eslesme:
        return f"O.NYM.{olcu_duzelt(eslesme.group(1))}.Antigron"

    # N2XH: N2XH-3X25-16 -> O.N2XH.3X25+16.HFFR
    eslesme = re.fullmatch(
        r"N2XH-?([0-9]+X[0-9]+(?:[,.][0-9]+)?(?:[-+][0-9]+)?)",
        raw,
    )
    if eslesme:
        return f"O.N2XH.{olcu_duzelt(eslesme.group(1), arti_eki=True)}.HFFR"

    # TTR: renk eki yoksa Beyaz, -S/-SIYAH varsa Siyah.
    ttr_raw = re.sub(r"-1$", "", raw)
    eslesme = re.fullmatch(
        r"TTR([0-9]+X[0-9]+(?:[,.][0-9]+)?)(?:-(S|SIYAH))?",
        ttr_raw,
    )
    if eslesme:
        olcu, siyah_eki = eslesme.groups()
        renk = "Siyah" if siyah_eki else "Beyaz"
        return f"O.{olcu_duzelt(olcu)}.{renk}.TTR"

    # Alvinal/NAYY ailesi.
    eslesme = re.fullmatch(r"ALVINAL-?(.+)", raw)
    if eslesme:
        return f"O.NAYY.{olcu_duzelt(eslesme.group(1), arti_eki=True)}.Alvinal"

    eslesme = re.fullmatch(r"AL-NAYY-YAVV-(.+)", raw)
    if eslesme:
        return f"O.NAYY.{olcu_duzelt(eslesme.group(1), arti_eki=True)}.Alvinal"

    eslesme = re.fullmatch(r"AL-([0-9]+X[0-9]+-[0-9]+)", raw)
    if eslesme:
        return f"O.NAYY.{olcu_duzelt(eslesme.group(1), arti_eki=True)}.Alvinal"

    # T-Soft mevcut listesinde NYY karsiligi yoksa uretilen hedef eslesmez.
    eslesme = re.fullmatch(r"NYY-?(.+)", raw)
    if eslesme:
        return f"O.NYY.{olcu_duzelt(eslesme.group(1), arti_eki=True)}"

    # Olcu icermeyen NHXMH ve T-Soft'ta ailesi bulunmayan diger kodlar
    # tahminle eslestirilmez.
    return ""


def santral_oznurkablo_tsoft_koduna_cevir(kod):
    """SANTRAL + OZNUR urun adini T-Soft Web Servis Kodu bicimine cevirir.

    Donusum tek tek urun listesi degil, kablo ailesi + damar/kesit + renk
    kuralidir. T-Soft'ta bulunmayan olculer normal eslesme kontrolunde elenir.
    Teknik olarak guvenli karsiligi olmayan orta gerilim aileleri tahminle
    baska bir urune baglanmaz.
    """
    raw = kod_temizle(kod)
    if not raw:
        return ""

    def olcu_duzelt(deger):
        sonuc = re.sub(r"\s+", "", str(deger).upper())
        sonuc = sonuc.replace(".", ",")
        sonuc = sonuc.replace("/", "+")
        return sonuc

    def renk_adi(deger):
        renk = sade_kod(deger)
        return {
            "SIYAH": "Siyah",
            "MAVI": "Mavi",
            "KAHVE": "Kahve",
            "KIRMIZI": "Kirmizi",
            "SYESIL": "Sari-Yesil",
            "SARIYESIL": "Sari-Yesil",
            "BEYAZ": "Beyaz",
            "GRI": "Gri",
            "SARI": "Sari",
            "YESIL": "Yesil",
        }.get(renk, "")

    # PVC tek damar: H05V-U / H07V-U -> NYA.
    # Halojensiz tek damar: H05Z1-U / H07Z1-U -> NYA.HFFR.
    eslesme = re.fullmatch(
        r"(H05V-U|H07V-U|H05Z1-U|H07Z1-U)\s+"
        r"([0-9]+(?:[,.][0-9]+)?)\s+(.+)",
        raw,
    )
    if eslesme:
        aile, kesit, renk_raw = eslesme.groups()
        renk = renk_adi(renk_raw)
        if not renk:
            return ""
        hffr_eki = ".HFFR" if "Z1" in aile else ""
        return f"O.1X{olcu_duzelt(kesit)}.{renk}.NYA{hffr_eki}"

    # SANTRAL YXV adlandirmasi T-Soft'ta N2XH.HFFR ailesiyle tutuluyor.
    eslesme = re.fullmatch(r"YXV\s+(.+)", raw)
    if eslesme:
        return f"O.N2XH.{olcu_duzelt(eslesme.group(1))}.HFFR"

    # Aluminyum NAYY, T-Soft'ta Alvinal son ekiyle tutuluyor.
    eslesme = re.fullmatch(r"NAYY\s+(.+)", raw)
    if eslesme:
        return f"O.NAYY.{olcu_duzelt(eslesme.group(1))}.Alvinal"

    # YVV, NA2XH, YXC7V-R, YXC8VZ3V-R ve YAXC... ailelerinin mevcut
    # T-Soft listesinde guvenli birebir web servis kodu yoktur.
    return ""


def yildiz_oznurkablo_tsoft_koduna_cevir(kod):
    """YILDIZ OZNUR urun karti kodunu dogrulanmis teknik T-Soft koduna cevirir."""
    return YILDIZ_OZNUR_KOD_ESLESMELERI.get(kod_temizle(kod), "")


def b2b_temiz_kod_uret(kod):
    kod = kod_temizle(kod)
    kod = entes_prefix_sil(kod)
    kod = para_birimi_temizle(kod)
    return kod


def eslesme_anahtarlari_uret(kod):
    anahtarlar = []
    gorulen = set()

    def ekle(tip, deger):
        deger = kod_temizle(deger)

        if not deger:
            return

        key = f"{tip}|{deger}"

        if key not in gorulen:
            gorulen.add(key)
            anahtarlar.append((key, tip, deger))

    raw = kod_temizle(kod)
    para_temiz = para_birimi_temizle(raw)

    prefixsiz = entes_prefix_sil(raw)
    prefixsiz_para_temiz = para_birimi_temizle(prefixsiz)

    ekle("BIREBIR", raw)
    ekle("PARA_TEMIZ", para_temiz)
    ekle("PREFIX_TEMIZ", prefixsiz)
    ekle("PREFIX_PARA_TEMIZ", prefixsiz_para_temiz)

    for deger in [raw, para_temiz, prefixsiz, prefixsiz_para_temiz]:
        sade = sade_kod(deger)

        if sade and len(sade) >= 4:
            ekle("SADE", sade)

    return anahtarlar


def stok_birim_metnini_sil(stok_metni):
    """Tum tedarikci ve markalarda stok sayisinin sonundaki birimi siler."""
    metin = metin_norm(stok_metni)

    # Noktali/noktasiz yaygin adet ve metre kisaltmalari.
    # Ornek: 48 ADET, 48 ADT., 48 AD., 1.000 MT.
    return re.sub(
        r"\s*(ADET|ADT|AD|TANE|SET|PCS|PC|METRE|MTR|MT|M)\.?\s*$",
        "",
        metin,
    ).strip()


def stok_degerine_cevir(stok):
    """
    STOKTA YOK -> 0
    48 ADET    -> 48
    48 ADT     -> 48
    48 ADT.    -> 48
    48 AD.     -> 48
    48 AD      -> 48

    1.000 MT. -> 1000

    Birim temizleme kurali tum tedarikci ve markalar icin ortaktir.
    STOKTA VAR -> 20 yok.
    Sayı yoksa hatalı stok sayılır.
    """

    if pd.isna(stok):
        return None

    metin = metin_norm(stok)

    if metin == "":
        return None

    if "YOK" in metin:
        return 0

    temiz_stok = stok_birim_metnini_sil(metin)
    sayi_eslesme = re.fullmatch(r"-?\d+(?:[.,]\d+)*", temiz_stok)

    if sayi_eslesme:
        sayi = sayiya_cevir(temiz_stok, varsayilan=None)

        if sayi is None:
            return None

        return max(0, sayi)

    return None


def sayiya_cevir(deger, varsayilan=0):
    if pd.isna(deger):
        return varsayilan

    metin = re.sub(r"\s+", "", str(deger).strip())

    if metin == "" or metin.lower() == "nan":
        return varsayilan

    if "." in metin and "," in metin:
        if metin.rfind(",") > metin.rfind("."):
            metin = metin.replace(".", "").replace(",", ".")
        else:
            metin = metin.replace(",", "")
    elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", metin):
        metin = metin.replace(".", "")
    elif re.fullmatch(r"-?\d{1,3}(?:,\d{3})+", metin):
        metin = metin.replace(",", "")
    else:
        metin = metin.replace(",", ".")

    try:
        return int(float(metin))
    except Exception:
        return varsayilan


def tamamen_bos_satirlari_sil(df):
    ilk_sayi = len(df)

    def satir_tamamen_bos_mu(row):
        for value in row.values:
            metin = str(value).replace("\xa0", " ").strip()

            if metin and metin.lower() != "nan":
                return False

        return True

    df = df[~df.apply(satir_tamamen_bos_mu, axis=1)].copy()

    silinen = ilk_sayi - len(df)

    if silinen > 0:
        print(f"Tamamen boş satır silindi: {silinen}")

    return df


def ayarlari_oku():
    if not AYAR_FILE.exists():
        print("Tedarikçi ayar dosyası bulunamadı:")
        print(AYAR_FILE)
        print("Önce python kurulum.py veya python tedarikci_ayar_guncelle.py çalıştır.")
        raise SystemExit

    df = pd.read_excel(AYAR_FILE, dtype=str).fillna("")

    gerekli = [
        "Tedarikci",
        "Aktif",
        "MarkaFiltresi",
        "KodKolonu",
        "StokKolonu",
        "MarkaKolonu"
    ]

    for kolon in gerekli:
        if kolon not in df.columns:
            print(f"tedarikci_ayar.xlsx içinde eksik kolon var: {kolon}")
            raise SystemExit

    if "EslesmeTipi" not in df.columns:
        df["EslesmeTipi"] = "WEBSERVIS"

    # MEANWELL B2B dosyasinda marka kolonu yoktur. Model No dogrudan T-Soft
    # Web Servis Kodu, Stok ise bayi stogu olarak kullanilir. Ayar dosyasinda
    # satir olmasa da bu B2B kaynagi otomatik olarak calismaya dahil edilir.
    meanwell_mask = df["Tedarikci"].apply(metin_norm).eq("MEANWELL")
    if meanwell_mask.any():
        for kolon, deger in MEANWELL_TEDARIKCI_AYARI.items():
            if kolon not in df.columns:
                df[kolon] = ""
            df.loc[meanwell_mask, kolon] = deger
    else:
        meanwell_satiri = {
            kolon: MEANWELL_TEDARIKCI_AYARI.get(kolon, "")
            for kolon in df.columns
        }
        df = pd.concat(
            [df, pd.DataFrame([meanwell_satiri])],
            ignore_index=True,
        )

    df["Tedarikci"] = df["Tedarikci"].astype(str).str.strip()
    df["Aktif"] = df["Aktif"].astype(str).str.strip()
    df["MarkaFiltresi"] = df["MarkaFiltresi"].astype(str).str.strip()

    df["KodKolonu"] = df["KodKolonu"].replace("", "Ürün Kodu")
    df["StokKolonu"] = df["StokKolonu"].replace("", "Stok")
    df["MarkaKolonu"] = df["MarkaKolonu"].replace("", "Marka")
    df["EslesmeTipi"] = df["EslesmeTipi"].replace("", "WEBSERVIS")

    df["EslesmeTipi"] = df["EslesmeTipi"].astype(str).str.strip().str.upper()

    aktif_df = df[
        df["Aktif"].isin(["1", "True", "TRUE", "aktif", "AKTIF"])
    ].copy()

    return aktif_df


def santral_eslesme_dosyasi_olustur():
    SANTRAL_ESLESME_FILE.parent.mkdir(exist_ok=True)

    if SANTRAL_ESLESME_FILE.exists():
        return

    df = pd.DataFrame([
        {
            "SantralKodu": "MCB-7",
            "TsoftWebServisKodu": "M4116"
        }
    ])

    with pd.ExcelWriter(SANTRAL_ESLESME_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="SantralEslesme", index=False)

    print("Santral eşleşme dosyası oluşturuldu:")
    print(SANTRAL_ESLESME_FILE)


def santral_eslesme_oku():
    santral_eslesme_dosyasi_olustur()

    df = pd.read_excel(SANTRAL_ESLESME_FILE, dtype=str).fillna("")

    gerekli = [
        "SantralKodu",
        "TsoftWebServisKodu"
    ]

    for kolon in gerekli:
        if kolon not in df.columns:
            print(f"santral_model_eslesme.xlsx içinde eksik kolon var: {kolon}")
            print("Olması gereken kolonlar:")
            print("SantralKodu")
            print("TsoftWebServisKodu")
            raise SystemExit

    lookup = {}

    for _, row in df.iterrows():
        santral_kodu = str(row["SantralKodu"]).strip()
        tsoft_kodu = str(row["TsoftWebServisKodu"]).strip()

        if not santral_kodu or not tsoft_kodu:
            continue

        olasi_anahtarlar = [
            kod_temizle(santral_kodu),
            b2b_temiz_kod_uret(santral_kodu),
            sade_kod(santral_kodu)
        ]

        for anahtar in olasi_anahtarlar:
            if anahtar:
                lookup.setdefault(anahtar, tsoft_kodu)

    return lookup


def manuel_tsoft_kodu_bul(santral_kodu, lookup):
    olasi_anahtarlar = [
        kod_temizle(santral_kodu),
        b2b_temiz_kod_uret(santral_kodu),
        sade_kod(santral_kodu)
    ]

    for anahtar in olasi_anahtarlar:
        if anahtar in lookup:
            return lookup[anahtar]

    return ""


def tedarikci_dosyasi_bul(tedarikci):
    klasor = B2B_DIR / tedarikci
    dosyalar = []

    if klasor.exists():
        for uzanti in ["*.xlsx", "*.xls", "*.csv"]:
            for dosya in klasor.glob(uzanti):
                ad = dosya.name.strip()

                # Excel acikken olusan gecici kilit dosyalarini alma.
                if ad.startswith("~$") or ad.startswith("."):
                    continue

                dosyalar.append(dosya)

    # Kok klasore konan KONYA_latest.xlsx gibi aktarim dosyalarini da bul.
    # B2B/<TEDARIKCI> klasorundeki dosya varsa her zaman ona oncelik verilir.
    if not dosyalar:
        aranan_kok_ad = f"{metin_norm(tedarikci)}_LATEST"
        for dosya in BASE_DIR.iterdir():
            if not dosya.is_file() or dosya.suffix.lower() not in {".xlsx", ".xls", ".csv"}:
                continue
            if metin_norm(dosya.stem) == aranan_kok_ad:
                dosyalar.append(dosya)

    if not dosyalar:
        return None

    return max(dosyalar, key=lambda x: x.stat().st_mtime)


def csv_oku(dosya):
    ayiraclar = [";", ",", "\t", "|"]
    encodingler = ["utf-8-sig", "cp1254", "latin1"]

    for encoding in encodingler:
        for ayirac in ayiraclar:
            try:
                df = pd.read_csv(
                    dosya,
                    sep=ayirac,
                    dtype=str,
                    encoding=encoding
                )

                if len(df.columns) > 1:
                    return df.fillna("")
            except Exception:
                pass

    raise Exception("CSV dosyası okunamadı.")


def dosya_oku(dosya):
    uzanti = dosya.suffix.lower()

    if uzanti in [".xlsx", ".xls"]:
        return pd.read_excel(dosya, dtype=str).fillna("")

    if uzanti == ".csv":
        return csv_oku(dosya)

    raise Exception(f"Desteklenmeyen dosya türü: {uzanti}")



def entes_fiyat_listesi_oku(dosya):
    """
    ENTES fiyat listesi bazen logo ve başlık nedeniyle 1. satırdan başlamaz.
    Bu fonksiyon ÜRÜN KODU ve ÜRÜN MODELİ geçen satırı otomatik başlık olarak bulur.
    Senin örneğinde başlık 4. satırda olduğu için bu şart.
    """
    uzanti = dosya.suffix.lower()

    if uzanti == ".csv":
        return csv_oku(dosya)

    if uzanti not in [".xlsx", ".xls"]:
        raise Exception(f"Desteklenmeyen fiyat listesi türü: {uzanti}")

    ham_df = pd.read_excel(dosya, dtype=str, header=None).fillna("")

    baslik_satiri = None

    for i, row in ham_df.iterrows():
        satir_metinleri = [metin_norm(x) for x in row.values]

        urun_kodu_var = any("URUN KODU" in x for x in satir_metinleri)
        urun_modeli_var = any(
            "URUN MODELI" in x or x in ["MODEL", "URUN MODEL", "MODEL KODU", "TIP"]
            for x in satir_metinleri
        )

        if urun_kodu_var and urun_modeli_var:
            baslik_satiri = i
            break

    if baslik_satiri is None:
        print("ENTES fiyat listesinde başlık satırı bulunamadı.")
        print("Aranan başlıklar: ÜRÜN KODU ve ÜRÜN MODELİ")
        print("Bulunan ilk satırlar:")
        print(ham_df.head(8).to_string(index=False))
        raise SystemExit

    df = pd.read_excel(dosya, dtype=str, header=baslik_satiri).fillna("")
    df = tamamen_bos_satirlari_sil(df)

    return df


def kolon_bul_esnek(df, adaylar):
    """Kolon adlarını Türkçe karakter / boşluk farklarına takılmadan bulur."""
    kolon_lookup = {metin_norm(kolon): kolon for kolon in df.columns}

    for aday in adaylar:
        aday_norm = metin_norm(aday)
        if aday_norm in kolon_lookup:
            return kolon_lookup[aday_norm]

    return None


def entes_fiyat_listesi_bul():
    """
    ENTES fiyat listesini otomatik bulur.
    Öncelik: FiyatListesi klasörü, sonra Ayarlar, sonra ana klasör.
    Dosya adında ENTES ve FIYAT/LISTE geçmesi yeterlidir.
    """
    aday_klasorler = [
        FIYAT_LISTESI_DIR,
        BASE_DIR / "Fiyat Listesi",
        BASE_DIR / "Fiyat",
        BASE_DIR / "Ayarlar",
        BASE_DIR,
    ]

    aday_dosyalar = []

    for klasor in aday_klasorler:
        if not klasor.exists():
            continue

        for uzanti in ["*.xlsx", "*.xls"]:
            for dosya in klasor.glob(uzanti):
                ad_norm = metin_norm(dosya.name)

                if "ENTES" in ad_norm and ("FIYAT" in ad_norm or "LISTE" in ad_norm):
                    aday_dosyalar.append(dosya)

    if not aday_dosyalar:
        return None

    return max(aday_dosyalar, key=lambda x: x.stat().st_mtime)


def kirac_model_kod_duzelt(df, kod_kolonu, kaynak_dosya, kontrol_mask=None):
    """
    Sadece KIRAC için çalışır.
    KIRAC ürün kodu kolonunda model yazılmışsa ENTES fiyat listesinden gerçek ürün koduna çevirir.
    Normal ürün kodu olan satırlara dokunmaz.
    """
    fiyat_dosya = entes_fiyat_listesi_bul()

    if fiyat_dosya is None:
        print()
        print("KIRAC düzeltmesi için ENTES fiyat listesi bulunamadı.")
        print("Lütfen ENTES fiyat listesini şu klasöre koy:")
        print(FIYAT_LISTESI_DIR)
        print("Dosya adında ENTES ve FIYAT veya LISTE geçsin. Örnek:")
        print("ENTES_FIYAT_LISTESI.xlsx")
        raise SystemExit

    fiyat_df = entes_fiyat_listesi_oku(fiyat_dosya)

    model_kolonu = kolon_bul_esnek(
        fiyat_df,
        [
            "Ürün Modeli",
            "Urun Modeli",
            "Ürün Model",
            "Urun Model",
            "Model",
            "Model Kodu",
            "Tip",
        ]
    )

    urun_kod_kolonu = kolon_bul_esnek(
        fiyat_df,
        [
            "Ürün Kodu",
            "Urun Kodu",
            "Stok Kodu",
            "Kod",
            "Malzeme Kodu",
            "Web Servis Kodu",
            "Ürün No",
            "Urun No",
        ]
    )

    if not model_kolonu or not urun_kod_kolonu:
        print("ENTES fiyat listesinde ürün modeli veya ürün kodu kolonu bulunamadı.")
        print("Bulunan kolonlar:")
        for c in fiyat_df.columns:
            print("-", c)
        print()
        print("Beklenen örnek kolonlar:")
        print("- Ürün Modeli / Model")
        print("- Ürün Kodu / Stok Kodu / Web Servis Kodu")
        raise SystemExit

    kod_anahtar_set = set()
    model_to_kod_set = {}

    # KIRAC düzeltmesinde M kodu ezberlenmez / elle sabitlenmez.
    # Özel örnekler kural olarak hedef model yazımına çevrilir;
    # son ürün kodu her zaman ENTES fiyat listesindeki Ürün Kodu kolonundan alınır.

    def genel_anahtar_gruplari(deger):
        """
        Normal kod/model karşılaştırması.

        ÖNEMLİ:
        Önce birebir/güçlü anahtarlar ayrı grup olarak denenir.
        Panel ölçüsü silme veya IP54 ekini kaldırma gibi gevşek aliaslar
        aynı gruba konmaz. Böylece şu tip ürünler yanlışlıkla çokluya düşmez:

        EVM-3C-48  -> önce tam EVM-3C-48 aranır
        EPM-4A-72  -> önce tam EPM-4A-72 aranır

        Eğer tam model fiyat listesinde yoksa, ancak o zaman panelsiz/IP54'süz
        gevşek alternatifler denenir. Bu da EMM-04CS-96 -> EMM-04CS gibi
        gerekli aliasları bozmadan çalıştırır.
        """
        raw = kod_temizle(deger)
        if not raw:
            return []

        para_temiz = para_birimi_temizle(raw)
        prefixsiz = entes_prefix_sil(raw)
        prefixsiz_para_temiz = para_birimi_temizle(prefixsiz)

        def tekil_grup_olustur(degerler):
            grup = []
            gorulen = set()

            for deger2 in degerler:
                temiz2 = kod_temizle(deger2)
                if temiz2 and temiz2 not in gorulen:
                    gorulen.add(temiz2)
                    grup.append(temiz2)

                sade2 = sade_kod(temiz2)
                if sade2 and sade2 not in gorulen:
                    gorulen.add(sade2)
                    grup.append(sade2)

            return grup

        gruplar = []

        # 1) En güçlü/birebir grup.
        # ENT/ENTES prefix temizliği burada kalabilir; bu hâlâ aynı ürün yazımıdır.
        birebir_grup = tekil_grup_olustur([
            raw,
            para_temiz,
            prefixsiz,
            prefixsiz_para_temiz,
        ])

        if birebir_grup:
            gruplar.append(birebir_grup)

        # 2) Gevşek alias grubu.
        # Bu grup sadece birebir grup sonuç vermezse denenir.
        alias_kaynaklar = []

        for kaynak in [raw, para_temiz, prefixsiz, prefixsiz_para_temiz]:
            kaynak = kod_temizle(kaynak)
            if not kaynak:
                continue

            # IP54 eki fiyat listesinde olup KIRAC'ta olmayabilir veya tersi olabilir.
            ip54siz = re.sub(r"[\s\-\._/]*IP54$", "", kaynak, flags=re.IGNORECASE).strip()
            if ip54siz and ip54siz != kaynak:
                alias_kaynaklar.append(ip54siz)

            # Panel ölçüsü aliası.
            # Örnek: EMM-04CS-96 -> EMM-04CS
            # Fakat bu alias birebirden sonra denendiği için
            # EVM-3C-48 / EPM-4A-72 gibi doğrudan fiyat listesinde olanlar çokluya düşmez.
            panelsiz = re.sub(
                r"^((?:EMM|EMR|EPM|EVM)[A-Z0-9\s\-\._/]*?)[\s\-\._/]+(?:48|72|96)$",
                r"\1",
                kaynak,
                flags=re.IGNORECASE
            ).strip()
            if panelsiz and panelsiz != kaynak and len(sade_kod(panelsiz)) >= 5:
                alias_kaynaklar.append(panelsiz)

            sade_kaynak = sade_kod(kaynak)

            m = re.match(r"^((?:MPR|EMM|EMR|EPM|EVM).+?)IP54$", sade_kaynak, flags=re.IGNORECASE)
            if m:
                alias_kaynaklar.append(m.group(1))

            # KIRAC bazı MPR ürünlerini V2 yazar; ENTES fiyat listesinde aynı ürün IP54 olarak geçer.
            # Örnekler:
            # MPR45S-V2  -> MPR-45S-IP54
            # MPR47SE-V2 -> MPR-47SE-IP54
            m = re.match(r"^(MPR\d+[A-Z]+)V2$", sade_kaynak, flags=re.IGNORECASE)
            if m:
                mpr_base = m.group(1)
                alias_kaynaklar.append(mpr_base)
                alias_kaynaklar.append(mpr_base + "IP54")

            # MPR-47SE-0,5 IP54 gibi yazımları, fiyat listesindeki
            # MPR-47SE-IP54-0,5 formuna yaklaştır.
            m = re.match(r"^(MPR\d+[A-Z]+)(0D5|05)IP54$", sade_kaynak, flags=re.IGNORECASE)
            if m:
                mpr_base = m.group(1)
                class_part = m.group(2)
                alias_kaynaklar.append(mpr_base + "IP54" + class_part)
                alias_kaynaklar.append(mpr_base + class_part)

            m = re.match(r"^(MPR\d+[A-Z]+)IP54(0D5|05)$", sade_kaynak, flags=re.IGNORECASE)
            if m:
                mpr_base = m.group(1)
                class_part = m.group(2)
                alias_kaynaklar.append(mpr_base + class_part + "IP54")
                alias_kaynaklar.append(mpr_base + class_part)

            # KIRAC: CM-8,33-KAPAK; ENTES fiyat listesi: CM serisi 8,33kVAr
            m = re.match(r"^CM[\s\-\._/]*(\d+(?:[,.]\d+)?)(?:[\s\-\._/]*KAPAK)?$", kaynak, flags=re.IGNORECASE)
            if m:
                kvar = m.group(1).replace(".", ",")
                alias_kaynaklar.append(f"CM SERISI {kvar}KVAR")
                alias_kaynaklar.append(f"CM SERİSİ {kvar}KVAR")
                alias_kaynaklar.append(f"CM SERISI {kvar} KVAR")
                alias_kaynaklar.append(f"CM SERİSİ {kvar} KVAR")

            m = re.match(r"^((?:EMM|EMR|EPM|EVM).+?)(?:48|72|96)$", sade_kaynak, flags=re.IGNORECASE)
            if m and len(m.group(1)) >= 5:
                alias_kaynaklar.append(m.group(1))

        alias_grup = tekil_grup_olustur(alias_kaynaklar)
        if alias_grup:
            gruplar.append(alias_grup)

        return gruplar

    def sayi_anahtar(deger, class_degeri=False):
        """
        0,5 / 0.5 / 05 gibi değerleri aynı anahtara çevirir.
        Örnek: 0,5 -> 0D5, 05 -> 0D5, 30VA -> 30
        """
        if deger is None or pd.isna(deger):
            return ""

        s = str(deger).upper().strip()
        s = s.replace(",", ".")
        s = re.sub(r"[^0-9.]", "", s)

        if not s:
            return ""

        # KIRAC tarafında 05 çoğu zaman class 0,5 anlamına geliyor.
        if class_degeri and re.fullmatch(r"0\d", s):
            s = "0." + s[-1]

        try:
            sayi = float(s)
            if sayi.is_integer():
                out = str(int(sayi))
            else:
                out = (f"{sayi:.6f}").rstrip("0").rstrip(".")
        except Exception:
            out = s

        return out.replace(".", "D")

    def ct_aile_duzelt(aile):
        """
        Akım trafosu aile adlarını aynı formata çevirir.

        KIRAC bazen AMN40-400-5-0,5 gibi yazar.
        ENTES fiyat listesinde aynı ürün ENT.A40 MN 400- 5 VA CL:0,5
        şeklinde yazılır. Bu yüzden AMN40 ile A40MN aynı aile kabul edilir.
        """
        a = sade_kod(aile)
        if not a:
            return ""

        # AMN40 -> A40MN
        m = re.fullmatch(r"AMN(\d+)", a)
        if m:
            return f"A{m.group(1)}MN"

        # A40MN zaten doğru format.
        m = re.fullmatch(r"A(\d+)MN", a)
        if m:
            return f"A{m.group(1)}MN"

        return a

    def akim_trafo_bilesenleri(deger):
        """
        KIRAC'ın kısa yazdığı akım trafosu modellerini ve ENTES fiyat listesindeki
        uzun açıklamalı model yazımlarını aynı parçalarla karşılaştırır.

        Örnekler:
        KIRAC: A100-2500-5-0,5-30VA
        Fiyat listesi: ENT.A100 2500/5 30VA CL:0,5 A.G. Akım Trafosu

        KIRAC: CXDH-450-15
        Fiyat listesi: ENT.CXD.H-450-15

        KIRAC: AB-50-5
        Fiyat listesi: ENT.AB 50- 10 VA CL:0,5

        KIRAC: AMN40-400-5-0,5
        Fiyat listesi: ENT.A40 MN 400- 5 VA CL:0,5
        """
        raw = kod_temizle(deger)
        if not raw:
            return None

        txt = raw
        txt = txt.replace("₺", " ").replace("€", " ").replace("$", " ")
        txt = re.sub(r"^(ENTES|ENT)[\.\-\s_]*", "", txt).strip()
        txt = txt.replace("/", " / ")
        txt = txt.replace(".", " ")
        txt = txt.replace("-", " - ")
        txt = re.sub(r"\s+", " ", txt).strip()

        # Başta yer alan harfli model ailesini alıyoruz.
        # ENT.CXD.H-450-15 -> CXD + H -> CXDH
        # ENT.A100 4000/5 -> A100
        tokenler = re.findall(r"\d+(?:[,.]\d+)?|[A-Z]+[0-9]*[A-Z]*", txt)
        if not tokenler:
            return None

        aile_parcalari = []
        baslangic_sayi_index = 0

        for i, token in enumerate(tokenler):
            if re.search(r"[A-Z]", token) and token not in ["VA", "CL", "AG", "A", "G", "AKIM", "TRAFOSU", "TRAFO"]:
                aile_parcalari.append(token)
                baslangic_sayi_index = i + 1
                continue

            if re.fullmatch(r"\d+(?:[,.]\d+)?", token):
                baslangic_sayi_index = i
                break

        aile = ct_aile_duzelt("".join(aile_parcalari))
        kalan_tokenler = tokenler[baslangic_sayi_index:]
        sayilar_raw = [t for t in kalan_tokenler if re.fullmatch(r"\d+(?:[,.]\d+)?", t)]

        if not aile and len(sayilar_raw) < 2:
            return None

        va_eslesme = re.search(r"(\d+(?:[,.]\d+)?)\s*VA\b", raw)
        cl_eslesme = re.search(r"\bCL\s*:?\s*(\d+(?:[,.]\d+)?)", raw)
        oran_eslesme = re.search(r"(\d+(?:[,.]\d+)?)\s*/\s*(\d+(?:[,.]\d+)?)", raw)

        primary = ""
        secondary = ""
        cls = ""
        va = ""

        kullanilan_indexler = set()

        if va_eslesme:
            va = sayi_anahtar(va_eslesme.group(1))

        if cl_eslesme:
            cls = sayi_anahtar(cl_eslesme.group(1), class_degeri=True)

        if oran_eslesme:
            primary = sayi_anahtar(oran_eslesme.group(1))
            secondary = sayi_anahtar(oran_eslesme.group(2))
        else:
            if sayilar_raw:
                primary = sayi_anahtar(sayilar_raw[0])
                kullanilan_indexler.add(0)

            # KIRAC kısa yazımında ikinci sayı genelde sekonderdir: 2500-5, 125-5, 250-2,5 gibi.
            # Ancak fiyat listesinde "50 10VA" gibi ikinci sayı VA ise sekonder saymayız.
            if len(sayilar_raw) >= 2:
                ikinci = sayi_anahtar(sayilar_raw[1])
                ikinci_va_mi = bool(va and ikinci == va)
                if not ikinci_va_mi:
                    secondary = ikinci
                    kullanilan_indexler.add(1)

        if not cls:
            # Class genelde sekonderden sonra gelir: 2500-5-0,5 veya 2500-5-05
            for i, sayi_raw in enumerate(sayilar_raw):
                if i in kullanilan_indexler:
                    continue

                aday = sayi_anahtar(sayi_raw, class_degeri=True)

                if va and aday == va:
                    continue

                # Class değerleri genelde 0,2 / 0,5 / 1 / 2,5 gibi olur.
                # 30, 75 gibi büyük sayılar VA olabileceği için class yapmayalım.
                try:
                    aday_float = float(aday.replace("D", "."))
                except Exception:
                    aday_float = 999

                if aday_float <= 5:
                    cls = aday
                    kullanilan_indexler.add(i)
                    break

        if not va:
            # KIRAC kısa yazımında son sayı VA olabilir: A100-3200-5-0,5-30
            for i in range(len(sayilar_raw) - 1, -1, -1):
                if i in kullanilan_indexler:
                    continue

                aday = sayi_anahtar(sayilar_raw[i])

                if aday and aday not in [primary, secondary, cls]:
                    try:
                        aday_float = float(aday.replace("D", "."))
                    except Exception:
                        aday_float = 0

                    if aday_float >= 5:
                        va = aday
                        kullanilan_indexler.add(i)
                        break

        if not primary:
            return None

        return {
            "aile": aile,
            "primary": primary,
            "secondary": secondary,
            "class": cls,
            "va": va,
        }

    def akim_trafo_anahtar_gruplari(deger):
        b = akim_trafo_bilesenleri(deger)
        if not b:
            return []

        aile = b.get("aile", "")
        p = b.get("primary", "")
        s = b.get("secondary", "")
        c = b.get("class", "")
        va = b.get("va", "")

        gruplar = []

        def grup_ekle(*keys):
            temizler = []
            gorulen = set()
            for key in keys:
                if key and key not in gorulen:
                    gorulen.add(key)
                    temizler.append(key)
            if temizler:
                gruplar.append(temizler)

        def k(*parcalar):
            return "CT|" + "|".join([p for p in parcalar if p])

        # En güçlü eşleşmeler önce. Arama sırasında ilk bulunan grup kullanılacak.
        if aile and p and s and c and va:
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"S:{s}", f"C:{c}", f"VA:{va}"))

        if aile and p and c and va:
            # Fiyat listesinde sekonder bazen yazılmaz: ENT.A80T 2500 15VA CL:0.5
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"C:{c}", f"VA:{va}"))

        if aile and p and s and c:
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"S:{s}", f"C:{c}"))

        if aile and p and c:
            # KIRAC: A80T-2500-5-05 / Fiyat: ENT.A80T 2500 15VA CL:0.5
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"C:{c}"))

        if aile and p and s and va:
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"S:{s}", f"VA:{va}"))

        if aile and p and va:
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"VA:{va}"))

        if aile and p and s:
            # KIRAC: AB-50-5 / Fiyat: ENT.AB 50- 10 VA CL:0,5 varsa sonraki grup F+P yakalar.
            grup_ekle(k(f"F:{aile}", f"P:{p}", f"S:{s}"))

        if aile and p:
            grup_ekle(k(f"F:{aile}", f"P:{p}"))

        # GÜVENLİK NOTU:
        # Önceki sürümlerde ailesiz/numerik anahtarlar da üretiliyordu:
        #   P:400|S:1 veya P:400|C:1 gibi.
        # Bu, ENT.C10-400-1 ürününü ENT.CMD-400-1 veya ENT.SRS3 400/1 400V 1 kVar
        # gibi tamamen farklı ailedeki ürünlerle hatalı eşleştirebiliyordu.
        # Bu yüzden artık CT eşleşmesinde aile zorunludur.
        # Ailesiz özel ürünler gerekiyorsa genel numerik kuraldan değil,
        # açık hedef-model kuralından çözülmelidir.

        return gruplar

    def anahtar_gruplari(deger):
        gruplar = []
        gruplar.extend(genel_anahtar_gruplari(deger))
        gruplar.extend(akim_trafo_anahtar_gruplari(deger))
        return gruplar

    def anahtar_listesi(deger):
        anahtarlar = []
        gorulen = set()
        for grup in anahtar_gruplari(deger):
            for anahtar in grup:
                if anahtar and anahtar not in gorulen:
                    gorulen.add(anahtar)
                    anahtarlar.append(anahtar)
        return anahtarlar

    def ab_akim_trafo_ozel_kod_bul(deger):
        """
        KIRAC tarafında AB akım trafoları kısa yazılır:
        AB-25-5-0,5 / AB-50-5 / AB-75-5-0,5 / AB-125-5-0,5 gibi.

        ENTES fiyat listesinde ise genelde açıklamalı görünür:
        ENT.AB 25- 10 VA CL:0,5
        ENT.AB 125- 10 VA CL:0,5 (Mühürlü)

        Bu fonksiyon sadece KIRAC düzeltme adımında çalışır.
        Tek tek manuel listeye yazılmamış AB ürünlerini de kuraldan bulur:
        - Aile AB olacak
        - Primer aynı olacak
        - Class KIRAC'ta varsa aynı olacak; yoksa AB için varsayılan 0,5 kabul edilir
        - VA KIRAC'ta yoksa AB için 10VA tercih edilir
        - Sekonder fiyat listesinde yazılmışsa KIRAC ile çelişmeyecek
        """
        raw = kod_temizle(deger)
        if not raw:
            return ""

        raw_sade = sade_kod(raw)

        # Önce genel CT ayrıştırıcıyı kullan.
        b = akim_trafo_bilesenleri(raw)

        # KIRAC kısa AB yazımları için daha net regex.
        # Örnekler:
        # AB-25-5-0,5  -> P=25, S=5, C=0,5
        # AB-50-5      -> P=50, S=5, C=0,5 varsayılan
        # AB-80-5-05   -> P=80, S=5, C=0,5
        regex_b = None
        m = re.match(
            r"^AB[\s\-\._/]*(\d+(?:[,.]\d+)?)(?:[\s\-\._/]+(\d+(?:[,.]\d+)?))?(?:[\s\-\._/]+(\d+(?:[,.]\d+)?))?",
            raw,
            flags=re.IGNORECASE
        )
        if m:
            regex_b = {
                "aile": "AB",
                "primary": sayi_anahtar(m.group(1)),
                "secondary": sayi_anahtar(m.group(2)) if m.group(2) else "",
                "class": sayi_anahtar(m.group(3), class_degeri=True) if m.group(3) else "",
                "va": "",
            }

        # AB ile başlamayan satırlarda bu özel kural çalışmasın.
        if not (raw_sade.startswith("AB") or (b and b.get("aile") == "AB") or regex_b):
            return ""

        # Regex sonucu daha güvenilir; yoksa genel ayrıştırıcıdan geleni kullan.
        if regex_b and regex_b.get("primary"):
            b = regex_b

        if not b or not b.get("primary"):
            return ""

        p = b.get("primary", "")
        s = b.get("secondary", "")
        c = b.get("class", "")
        va = b.get("va", "")

        # KIRAC AB kısa yazımlarında class yazmıyorsa, ürün açıklamalarında
        # karşılığı genelde CL:0,5 olduğu için varsayılan 0,5 kabul ediyoruz.
        if not c:
            c = "0D5"

        adaylar = []

        def ab_modelinden_bilesen_al(model_degeri):
            """
            ENTES fiyat listesinde AB ürünleri çoğunlukla açıklamalı yazılır:
            ENT.AB 80- 10 VA CL:0,5
            ENT.AB 125- 10 VA CL:0,5 (Mühürlü)

            Genel akım trafosu ayrıştırıcısı bazen bu satırlarda 10VA'yı sekonder gibi
            yorumlayabildiği için AB ailesinde fiyat listesi metninden doğrudan okuma yapıyoruz.
            """
            norm = metin_norm(model_degeri)
            if not norm:
                return None

            # ENT/ENTES prefixi opsiyonel, ama aile mutlaka AB olmalı. A80/A100 gibi ailelerle karışmasın.
            m_ab = re.search(
                r"(?:^|[^A-Z0-9])(?:ENTES|ENT)?[\.\-\s_]*AB(?:[^A-Z0-9]|$)",
                norm
            )
            if not m_ab:
                return None

            # AB'den hemen sonra gelen ilk sayı primerdir: ENT.AB 80- 10 VA ... -> 80
            m_p = re.search(
                r"(?:ENTES|ENT)?[\.\-\s_]*AB[\s\.\-_]*(\d+(?:[,.]\d+)?)",
                norm
            )
            if not m_p:
                return None

            m_va = re.search(r"(\d+(?:[,.]\d+)?)\s*VA\b", norm)
            m_cl = re.search(r"\bCL\s*:?\s*(\d+(?:[,.]\d+)?)", norm)

            return {
                "aile": "AB",
                "primary": sayi_anahtar(m_p.group(1)),
                "secondary": "",
                "class": sayi_anahtar(m_cl.group(1), class_degeri=True) if m_cl else "",
                "va": sayi_anahtar(m_va.group(1)) if m_va else "",
            }

        for _, fiyat_row in fiyat_df.iterrows():
            model_raw2 = fiyat_row.get(model_kolonu, "")
            kod_raw2 = fiyat_row.get(urun_kod_kolonu, "")
            kod2 = str(kod_raw2).strip() if not pd.isna(kod_raw2) else ""
            if not kod2:
                continue

            # AB ailesinde önce fiyat listesi metninden doğrudan oku.
            # Böylece AB-80-5-0,5 -> ENT.AB 80- 10 VA CL:0,5 kuraldan bulunur.
            mb = ab_modelinden_bilesen_al(model_raw2)

            # Doğrudan AB okuma olmazsa eski genel ayrıştırıcıya düş.
            if not mb:
                mb = akim_trafo_bilesenleri(model_raw2)

            if not mb:
                continue

            if mb.get("aile") != "AB":
                continue

            if mb.get("primary") != p:
                continue

            mb_c = mb.get("class", "")
            mb_va = mb.get("va", "")
            mb_s = mb.get("secondary", "")

            # Class çelişirse alma. Fiyat listesinde class varsa KIRAC class'ıyla aynı olmalı.
            if c and mb_c and mb_c != c:
                continue

            # KIRAC'ta sekonder varsa, fiyat listesinde açıkça farklı sekonder yazılmışsa alma.
            # AB açıklamalarında çoğunlukla sekonder yazmaz, bu durumda çelişki sayılmaz.
            if s and mb_s and mb_s != s:
                continue

            # KIRAC'ta VA açıkça varsa aynı olmalı. Yoksa AB için 10VA tercih edeceğiz.
            if va and mb_va and mb_va != va:
                continue

            model_norm = metin_norm(model_raw2)
            adaylar.append({
                "kod": kod2,
                "model": str(model_raw2),
                "va": mb_va,
                "class": mb_c,
                "secondary": mb_s,
                "muhurlu": "MUHURLU" in model_norm or "MÜHÜRLÜ" in str(model_raw2).upper(),
            })

        # Aynı ürün kodu birden fazla satırda görünürse tekilleştir.
        tekil = {}
        for aday in adaylar:
            tekil.setdefault(aday["kod"], aday)
        adaylar = list(tekil.values())

        if not adaylar:
            return ""

        # En doğru adayları puanla. AB kısa kodlarında VA genelde yazmaz;
        # kullanıcı örneklerine göre doğru karşılık 10VA olan modeldir.
        puanli = []
        for aday in adaylar:
            puan = 0

            if aday.get("class") == c:
                puan += 50

            if s and (not aday.get("secondary") or aday.get("secondary") == s):
                puan += 20

            if va:
                if aday.get("va") == va:
                    puan += 40
            else:
                if aday.get("va") == "10":
                    puan += 100

            # Mühürlü ibaresi tek başına seçtirmesin; sadece eşitlik bozucu küçük tercih olsun.
            if aday.get("muhurlu"):
                puan += 1

            puanli.append((puan, aday))

        puanli.sort(key=lambda x: x[0], reverse=True)
        en_yuksek_puan = puanli[0][0]
        en_iyi = [aday for puan, aday in puanli if puan == en_yuksek_puan]

        # En yüksek puanlı tek ürün kodu varsa onu seç.
        en_iyi_kodlar = sorted(set(a["kod"] for a in en_iyi))
        if len(en_iyi_kodlar) == 1:
            return en_iyi_kodlar[0]

        # Puan eşitliği varsa ama 10VA + class eşleşen tek aday varsa onu seç.
        va10_class_adaylar = [
            a for a in adaylar
            if a.get("va") == "10" and (not c or not a.get("class") or a.get("class") == c)
        ]
        va10_class_kodlar = sorted(set(a["kod"] for a in va10_class_adaylar))
        if len(va10_class_kodlar) == 1:
            return va10_class_kodlar[0]

        return ""


    def cm_serisi_ozel_kod_bul(deger):
        """
        KIRAC: CM-8,33-KAPAK
        ENTES fiyat listesi: CM serisi 8,33kVAr

        M kodu elle yazılmaz; fiyat listesinde CM serisi + aynı kVAr değerini bulup
        karşısındaki ürün kodunu döndürür. Sadece KIRAC düzeltme adımında çalışır.
        """
        raw = kod_temizle(deger)
        if not raw:
            return ""

        m = re.match(r"^CM[\s\-\._/]*(\d+(?:[,.]\d+)?)(?:[\s\-\._/]*KAPAK)?$", raw, flags=re.IGNORECASE)
        if not m:
            return ""

        hedef_kvar = sayi_anahtar(m.group(1))
        if not hedef_kvar:
            return ""

        adaylar = []

        for _, fiyat_row in fiyat_df.iterrows():
            model_raw2 = fiyat_row.get(model_kolonu, "")
            kod_raw2 = fiyat_row.get(urun_kod_kolonu, "")
            kod2 = str(kod_raw2).strip() if not pd.isna(kod_raw2) else ""
            if not kod2:
                continue

            norm = metin_norm(model_raw2)
            sade = sade_kod(norm)

            # CM serisi dışındaki ürünlerle karışmasın.
            if not ("CMSERISI" in sade or "CMSERİSİ" in sade or re.search(r"\bCM\b", norm)):
                continue

            # kVAr değeri model içinde geçmeli: 8,33kVAr / 8.33 kVAr gibi.
            kvar_eslesmeleri = re.findall(r"(\d+(?:[,.]\d+)?)\s*K\s*V\s*A\s*R", norm)
            if not kvar_eslesmeleri:
                # Bazı listelerde sadece 8,33 yazabilir; CM serisi şartı sağlandığı için bunu da deneriz.
                kvar_eslesmeleri = re.findall(r"\d+(?:[,.]\d+)?", norm)

            for kvar_raw in kvar_eslesmeleri:
                if sayi_anahtar(kvar_raw) == hedef_kvar:
                    adaylar.append(kod2)
                    break

        aday_kodlar = sorted(set(adaylar))
        if len(aday_kodlar) == 1:
            return aday_kodlar[0]

        return ""

    def mpr_v2_ip54_ozel_kod_bul(deger):
        """
        KIRAC: MPR45S-V2 / MPR47SE-V2
        ENTES fiyat listesi: MPR-45S-IP54 / MPR-47SE-IP54

        M kodu elle bilinmeden fiyat listesindeki IP54 model satırından bulunur.
        """
        raw = kod_temizle(deger)
        if not raw:
            return ""

        raw_sade = sade_kod(raw)
        m = re.match(r"^(MPR)(\d+)([A-Z]+)V2$", raw_sade, flags=re.IGNORECASE)
        if not m:
            return ""

        hedef_sade = f"MPR{m.group(2)}{m.group(3)}IP54"
        adaylar = []

        for _, fiyat_row in fiyat_df.iterrows():
            model_raw2 = fiyat_row.get(model_kolonu, "")
            kod_raw2 = fiyat_row.get(urun_kod_kolonu, "")
            kod2 = str(kod_raw2).strip() if not pd.isna(kod_raw2) else ""
            if not kod2:
                continue

            model_sade = sade_kod(model_raw2)
            if model_sade == hedef_sade:
                adaylar.append(kod2)

        aday_kodlar = sorted(set(adaylar))
        if len(aday_kodlar) == 1:
            return aday_kodlar[0]

        return ""


    def ens_ozel_kod_bul(deger):
        """
        KIRAC ENS kısa kodlarını M kodu elle yazmadan ENTES fiyat listesinden çözer.

        Örnekler:
        ENS-CCT24-300-5    -> fiyat listesinde ENS.CCT 24 300/5A satırı
        ENS-3PMD-50-3X1000 -> fiyat listesinde ENS.3PMD 50 3x1000A/0,330V satırı
        ENS-CYS816-1200    -> fiyat listesinde ENS.CYS 816 1200/5 satırı
        ENS-CYS816-5000    -> fiyat listesinde ENS.CYS 816 5000/5 satırı

        Bu fonksiyon ürün kodunu manuel tablodan değil, fiyat listesindeki eşleşen
        model satırının ÜRÜN KODU kolonundan alır.
        """
        raw = kod_temizle(deger)
        raw_sade = sade_kod(raw)
        if not raw_sade.startswith("ENS"):
            return ""

        hedef = None

        # ENS kısa kodlarında rakamları greedy regex ile ayırmak hatalı sonuç verebilir.
        # Bu yüzden önce tire/nokta/boşluklu gerçek KIRAC yazımından parse ediyoruz.
        # ENS-CCT24-300-5 => seri=24, primer=300, sekonder=5
        m = re.match(
            r"^ENS[\s\-\._/]*CCT[\s\-\._/]*(\d+)[\s\-\._/]+(\d+)[\s\-\._/]+(\d+)A?$",
            raw,
            flags=re.IGNORECASE
        )
        if m:
            hedef = {
                "tip": "CCT",
                "seri": m.group(1),
                "primer": m.group(2),
                "sekonder": m.group(3),
            }

        # Fallback: ENSCCT243005 gibi sade anahtar gelirse CCT seri genelde 2 hanelidir,
        # son hane sekonderdir, aradaki bölüm primerdir.
        if hedef is None:
            m = re.match(r"^ENSCCT(\d{2})(\d+)(\d)$", raw_sade, flags=re.IGNORECASE)
            if m:
                hedef = {
                    "tip": "CCT",
                    "seri": m.group(1),
                    "primer": m.group(2),
                    "sekonder": m.group(3),
                }

        # ENS-3PMD-50-3X1000 => ENS.3PMD 50 3x1000A/0,330V
        if hedef is None:
            m = re.match(r"^ENS[\s\-\._/]*3PMD[\s\-\._/]*(\d+)[\s\-\._/]*3\s*[Xx]\s*(\d+)$", raw, flags=re.IGNORECASE)
            if not m:
                m = re.match(r"^ENS3PMD(\d+)3X(\d+)$", raw_sade, flags=re.IGNORECASE)
            if m:
                hedef = {
                    "tip": "3PMD",
                    "seri": m.group(1),
                    "faz": "3",
                    "primer": m.group(2),
                }

        # ENS-CYS816-1200 => seri=816, primer=1200, sekonder=5
        if hedef is None:
            m = re.match(
                r"^ENS[\s\-\._/]*CYS[\s\-\._/]*(\d+)[\s\-\._/]+(\d+)(?:[\s\-\._/]+(\d+))?$",
                raw,
                flags=re.IGNORECASE
            )
            if m:
                hedef = {
                    "tip": "CYS",
                    "seri": m.group(1),
                    "primer": m.group(2),
                    "sekonder": m.group(3) or "5",
                }

        # Fallback: ENSCYS8161200 gibi sade anahtar gelirse CYS seri genelde 3 hanelidir,
        # kalan bölüm primerdir; sekonder varsayılan 5'tir.
        if hedef is None:
            m = re.match(r"^ENSCYS(\d{3})(\d+)$", raw_sade, flags=re.IGNORECASE)
            if m:
                hedef = {
                    "tip": "CYS",
                    "seri": m.group(1),
                    "primer": m.group(2),
                    "sekonder": "5",
                }

        if hedef is None:
            return ""

        adaylar = []

        for _, fiyat_row in fiyat_df.iterrows():
            model_raw2 = fiyat_row.get(model_kolonu, "")
            kod_raw2 = fiyat_row.get(urun_kod_kolonu, "")
            kod2 = str(kod_raw2).strip() if not pd.isna(kod_raw2) else ""
            if not kod2:
                continue

            norm = metin_norm(model_raw2)
            sade = sade_kod(norm)

            if hedef["tip"] == "CCT":
                # ENS.CCT 24 300/5A, ENS CCT 24 300 5A vb. yazımları yakalar.
                if not sade.startswith("ENSCCT"):
                    continue

                sayilar = re.findall(r"\d+", norm)
                if len(sayilar) < 3:
                    continue

                # CCT sonrası: seri, primer, sekonder
                if (
                    sayilar[0] == hedef["seri"]
                    and sayilar[1] == hedef["primer"]
                    and sayilar[2] == hedef["sekonder"]
                ):
                    adaylar.append(kod2)

            elif hedef["tip"] == "3PMD":
                # ENS.3PMD 50 3x1000A/0,330V
                if "ENS3PMD" not in sade:
                    continue

                # Seri değeri ve 3x primer değeri birlikte aranır.
                seri_ok = re.search(r"3\s*PMD\s*" + re.escape(hedef["seri"]), norm)
                primer_ok = re.search(r"3\s*[Xx]\s*" + re.escape(hedef["primer"]) + r"\s*A?", norm)

                if not seri_ok:
                    # sade anahtarda ENS3PMD50 gibi geçebilir.
                    seri_ok = (f"ENS3PMD{hedef['seri']}" in sade)

                if not primer_ok:
                    primer_ok = (f"3X{hedef['primer']}" in sade)

                if seri_ok and primer_ok:
                    adaylar.append(kod2)

            elif hedef["tip"] == "CYS":
                # ENS.CYS 816 1200/5
                if not sade.startswith("ENSCYS"):
                    continue

                sayilar = re.findall(r"\d+", norm)
                if len(sayilar) < 3:
                    continue

                if (
                    sayilar[0] == hedef["seri"]
                    and sayilar[1] == hedef["primer"]
                    and sayilar[2] == hedef["sekonder"]
                ):
                    adaylar.append(kod2)

        aday_kodlar = sorted(set(adaylar))
        if len(aday_kodlar) == 1:
            return aday_kodlar[0]

        return ""

    def fiyat_model_hedeflerinden_kod_bul(hedefler, tercih_muhurlu=False):
        """
        Verilen hedef model yazımlarını ENTES fiyat listesindeki Ürün Modeli
        kolonunda arar ve eşleşen satırın Ürün Kodu değerini döndürür.

        Burada M kodu elle yazılmaz. Örnek:
        KIRAC: ENS-CCT24-300-5
        Hedef model: ENS.CCT 24 300/5A
        Sonuç: fiyat listesindeki o satırın Ürün Kodu.
        """
        if not hedefler:
            return ""

        hedef_sadeler = []
        for hedef in hedefler:
            h_sade = sade_kod(hedef)
            if h_sade:
                hedef_sadeler.append((str(hedef), h_sade))

        if not hedef_sadeler:
            return ""

        adaylar = []

        for _, fiyat_row in fiyat_df.iterrows():
            model_raw2 = fiyat_row.get(model_kolonu, "")
            kod_raw2 = fiyat_row.get(urun_kod_kolonu, "")
            kod2 = str(kod_raw2).strip() if not pd.isna(kod_raw2) else ""
            if not kod2:
                continue

            model_norm = metin_norm(model_raw2)
            model_sade = sade_kod(model_norm)
            if not model_sade:
                continue

            en_iyi_puan = 0
            for _, hedef_sade in hedef_sadeler:
                if model_sade == hedef_sade:
                    en_iyi_puan = max(en_iyi_puan, 1000)
                elif hedef_sade in model_sade:
                    en_iyi_puan = max(en_iyi_puan, 800)
                elif model_sade in hedef_sade and len(model_sade) >= 6:
                    en_iyi_puan = max(en_iyi_puan, 600)

            if en_iyi_puan <= 0:
                continue

            if tercih_muhurlu and ("MUHURLU" in model_sade or "MÜHÜRLÜ" in str(model_raw2).upper()):
                en_iyi_puan += 50

            adaylar.append({
                "kod": kod2,
                "model": str(model_raw2),
                "puan": en_iyi_puan,
            })

        if not adaylar:
            return ""

        adaylar.sort(key=lambda x: x["puan"], reverse=True)
        en_yuksek = adaylar[0]["puan"]
        en_iyi = [a for a in adaylar if a["puan"] == en_yuksek]
        en_iyi_kodlar = sorted(set(a["kod"] for a in en_iyi))

        if len(en_iyi_kodlar) == 1:
            return en_iyi_kodlar[0]

        return ""

    def kirac_kural_modelden_kod_bul(deger):
        """
        KIRAC'ta kısa/eski yazılan modelleri M kodu elle yazmadan,
        ENTES fiyat listesindeki model yazımına dönüştürerek arar.
        Bulduğu satırın Ürün Kodu değerini döndürür.
        """
        raw = kod_temizle(deger)
        raw_sade = sade_kod(raw)
        if not raw_sade:
            return ""

        hedefler = []
        tercih_muhurlu = False

        # ENT.SRS1 -> ENT.SRS1 230/1,66
        if raw_sade in {"ENTSRS1", "SRS1", "ENTESSRS1"}:
            hedefler.extend([
                "ENT.SRS1 230/1,66",
                "SRS1 230/1,66",
            ])

        # 150-5-0,5M -> ENT.AB 150- 10 VA CL:0,5(Mühürlü)
        if raw_sade in {"150505M", "150505"}:
            hedefler.extend([
                "ENT.AB 150- 10 VA CL:0,5",
                "ENT.AB 150- 10 VA CL:0.5",
            ])
            tercih_muhurlu = True

        # EMM panel/ölçü farkları fiyat listesindeki model satırından çözülür.
        # M kodu elle yazılmaz; hedef model fiyat listesinde aranır.
        # Örnek: EMM-04CS-96 -> EMM-04CS
        if raw_sade in {"EMM04CS96", "EMM04CS"}:
            hedefler.append("EMM-04CS")

        # Örnek: EMM-04S-96 -> EMM-04
        if raw_sade in {"EMM04S96", "EMM04S", "EMM04"}:
            hedefler.append("EMM-04")

        # MPR IP54 / V2 / class farkları fiyat listesindeki model satırından çözülür.
        # M kodu elle yazılmaz; hedef model fiyat listesinde aranır.
        # Örnek: MPR-47SE -> MPR-47SE-IP54
        if raw_sade in {"MPR47SE", "MPR47SEV2", "MPR47SEIP54"}:
            hedefler.append("MPR-47SE-IP54")

        # Örnekler:
        # MPR-47SE-0,5      -> MPR-47SE-IP54-0,5
        # MPR-47SE-0,5 IP54 -> MPR-47SE-IP54-0,5
        # MPR-47SE IP54 0,5 -> MPR-47SE-IP54-0,5
        # M kodu elle yazılmaz; hedef model fiyat listesinde aranır.
        if raw_sade in {
            "MPR47SE05",
            "MPR47SE0D5",
            "MPR47SE05IP54",
            "MPR47SE0D5IP54",
            "MPR47SEIP5405",
            "MPR47SEIP540D5",
        }:
            hedefler.extend([
                "MPR-47SE-IP54-0,5",
                "MPR-47SE-IP54 0,5",
                "MPR-47SE-IP54-0.5",
                "MPR-47SE 0,5 IP54",
                "MPR-47SE-0,5 IP54",
            ])

        # MPR45S-V2 -> MPR-45S-IP54
        if raw_sade in {"MPR45SV2", "MPR45SIP54"}:
            hedefler.append("MPR-45S-IP54")

        # MPR-47SE-0,5 IP54 -> MPR-47SE-IP54-0,5
        # Bu da genel class/IP54 yazım kuralıdır; M kodu elle yazılmaz.
        if raw_sade in {
            "MPR47SE05",
            "MPR47SE0D5",
            "MPR47SE05IP54",
            "MPR47SE0D5IP54",
            "MPR47SEIP5405",
            "MPR47SEIP540D5",
        }:
            hedefler.extend([
                "MPR-47SE-IP54-0,5",
                "MPR-47SE-IP54-0.5",
                "MPR-47SE-IP54 0,5",
            ])

        # MPR-47S OG-D 0,5 IP51 -> MPR-47S-OG-D-0,5
        # Bu yazımda 0,5 class bilgisi varsa hedef model IP54 değil, classlı modeldir.
        if raw_sade in {
            "MPR47SOGD05IP51",
            "MPR47SOGD0D5IP51",
            "MPR47SOGD05IP54",
            "MPR47SOGD0D5IP54",
            "MPR47SOGDIP5105",
            "MPR47SOGDIP510D5",
            "MPR47SOGD05",
            "MPR47SOGD0D5",
        }:
            hedefler.extend([
                "MPR-47S-OG-D-0,5",
                "MPR-47S-OG-D-0.5",
                "MPR-47S-OG-D 0,5",
                "MPR-47S OG-D 0,5",
            ])

        # MPR-47S OG-D IP51 -> MPR-47S-OG-D-IP54
        # IP51 gelen SANTRAL yazımı fiyat listesindeki IP54 modeline gider.
        if raw_sade in {"MPR47SOGDIP51", "MPR47SOGDIP54", "MPR47SOGD"}:
            hedefler.append("MPR-47S-OG-D-IP54")

        # MPR-47D-0,5 IP54 -> MPR-47S-D-0,5-IP54
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {
            "MPR47D05IP54",
            "MPR47D0D5IP54",
            "MPR47DIP5405",
            "MPR47DIP540D5",
            "MPR47D05",
            "MPR47D0D5",
        }:
            hedefler.extend([
                "MPR-47S-D-0,5-IP54",
                "MPR-47S-D-0.5-IP54",
                "MPR-47S-D-IP54-0,5",
                "MPR-47S-D-IP54-0.5",
            ])

        # MPR-47D IP54 -> MPR-47S-D-IP54
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {"MPR47DIP54", "MPR47D"}:
            hedefler.append("MPR-47S-D-IP54")

        # MPR-47S OG-D IP51 -> MPR-47S-OG-D-IP54
        # Sonda IP51 yazsa bile fiyat listesindeki ürün IP54 yazımıyla bulunur.
        if raw_sade in {"MPR47SOGDIP51", "MPR47SOGDIP54", "MPR47SOGD"}:
            hedefler.append("MPR-47S-OG-D-IP54")

        # RGSR SVC yazım farkları fiyat listesindeki gerçek model satırından çözülür.
        # M kodu elle yazılmaz; hedef model fiyat listesinde aranır.
        # Örnek: RGSR-20S (SVC) -> RGSR-20S-OG
        # Örnek: RGSR-15S(SVC)  -> RGSR-15S
        if raw_sade in {"RGSR20SSVC", "RGSR20SVC"}:
            hedefler.append("RGSR-20S-OG")

        if raw_sade in {"RGSR15SSVC", "RGSR15SVC"}:
            hedefler.append("RGSR-15S")

        # NH bıçaklı sigorta yazım farkı fiyat listesindeki model satırından çözülür.
        # M kodu elle yazılmaz; ENTES fiyat listesinde bu model aranır.
        # GENEL VE KATI KURAL:
        # ENT.NH00 25A -> ENT.NH00 25A Bıçaklı Sigorta
        # ENT.NH1 160A -> ENT.NH1 160A Bıçaklı Sigorta
        # NH00 / NH1 ailesi ve amper değeri birebir aynı olmak zorundadır.
        # Sadece sayı benziyor diye farklı NH ailesi veya farklı amperle eşleşme yapılmaz.
        nh_eslesme = re.fullmatch(r"(?:ENT|ENTES)?(NH00|NH1)(\d+)A", raw_sade)
        if nh_eslesme:
            nh_aile = nh_eslesme.group(1)
            amper = nh_eslesme.group(2)
            hedefler.extend([
                f"ENT.{nh_aile} {amper}A Bıçaklı Sigorta",
                f"{nh_aile} {amper}A Bıçaklı Sigorta",
                f"ENT.{nh_aile} {amper} A Bıçaklı Sigorta",
                f"{nh_aile} {amper} A Bıçaklı Sigorta",
            ])

        # Kullanıcının verdiği CT kısa kodları: hedef model yazımı fiyat listesinden aranır.
        # M kodu elle yazılmaz; hedef model bulununca o satırın Ürün Kodu alınır.
        b = akim_trafo_bilesenleri(raw)
        if b:
            aile = b.get("aile", "")
            p = b.get("primary", "")
            c = b.get("class", "")
            va = b.get("va", "")

            # KIRAC tarafında class yoksa bu özel kurallarda çoğunlukla 0,5 kabul edilir.
            if not c:
                c = "0D5"

            # Örneklerden gelen VA tercihleri. Bu map M kodu değildir;
            # sadece fiyat listesindeki doğru model satırını bulmak için VA seçer.
            tercih_va = {
                ("A100", "2000", "0D5"): "15",
                ("A100", "2500", "0D5"): "30",
                ("A100", "3200", "0D5"): "15",
                ("A40", "600", "0D5"): "10",
                ("A60", "1200", "0D5"): "10",
                ("A60", "1250", "0D5"): "10",
                ("A80", "1500", "0D5"): "15",
                ("A80", "1600", "0D5"): "30",
            }

            hedef_va = va or tercih_va.get((aile, p, c), "")

            if aile and p and c and hedef_va:
                class_text = c.replace("D", ",")
                # ENT.A100 2500 30 VA Cl:0,5 gibi açıklamalı fiyat listesi modelleri
                hedefler.extend([
                    f"ENT.{aile} {p} {hedef_va} VA CL:{class_text}",
                    f"ENT.{aile} {p} {hedef_va}VA CL:{class_text}",
                    f"ENT.{aile} {p}- {hedef_va} VA CL:{class_text}",
                    f"{aile} {p} {hedef_va} VA CL:{class_text}",
                ])

        if not hedefler:
            return ""

        return fiyat_model_hedeflerinden_kod_bul(hedefler, tercih_muhurlu=tercih_muhurlu)

    for _, row in fiyat_df.iterrows():
        model_raw = row.get(model_kolonu, "")
        kod_raw = row.get(urun_kod_kolonu, "")

        kod = str(kod_raw).strip() if not pd.isna(kod_raw) else ""
        if not kod:
            continue

        for kod_key in anahtar_listesi(kod):
            kod_anahtar_set.add(kod_key)

        for model_key in anahtar_listesi(model_raw):
            if not model_key:
                continue

            model_to_kod_set.setdefault(model_key, set()).add(kod)

    df = df.copy()

    # Sayım ve düzeltme sadece marka filtresine giren KIRAC satırlarında yapılsın.
    # Böylece KIRAC dosyasındaki diğer markalar sayıları şişirmez.
    if kontrol_mask is None:
        kontrol_mask = pd.Series([True] * len(df), index=df.index)
    else:
        kontrol_mask = pd.Series(kontrol_mask, index=df.index).fillna(False).astype(bool)

    orijinal_kolon = "KIRAC Orijinal Ürün Kodu"
    durum_kolon = "KIRAC Kod Düzeltme Durumu"
    not_kolon = "KIRAC Kod Düzeltme Notu"

    if orijinal_kolon not in df.columns:
        df[orijinal_kolon] = df[kod_kolonu]

    df[durum_kolon] = "Dokunulmadı"
    df[not_kolon] = ""

    rapor_satirlari = []
    modelden_bulunan_satirlar = []
    modeli_bulunmayan_satirlar = []
    urun_kodu_olan_satirlar = []
    coklu_model_satirlari = []
    bos_kod_satirlari = []
    geriye_kalan_satirlar = []

    cevrilen = 0
    zaten_kod = 0
    bulunamayan = 0
    coklu = 0
    bos_kod = 0

    # KIRAC tarafında karşılığı olmadığı belirtilen modeller.
    # Bunlar kural/alias sistemine yakalansa bile ürün koduna çevrilmez;
    # raporda doğrudan "Model karşılığı bulunamadı" olarak kalır.
    kirac_model_karsiligi_bulunamayan_sade_kodlar = {
        sade_kod("CF-230-2,5"),
        # ENS-CCT24-300-5 karşılıksız değildir.
        # Kuralı: ENS-CCT24-300-5 -> ENS.CCT 24 300/5A
        # Bu yüzden bulunamayan listesine alınmaz; ens_ozel_kod_bul ile fiyat listesinden çözülür.
        sade_kod("EPM-07-S"),
        sade_kod("ENT.C10-400-1"),
        sade_kod("C10-400-1"),
    }

    for idx, row in df.iterrows():
        if not bool(kontrol_mask.loc[idx]):
            df.at[idx, durum_kolon] = "Marka filtresi dışında"
            df.at[idx, not_kolon] = "Bu satır KIRAC marka filtresine girmediği için ENTES düzeltme sayımına dahil edilmedi."
            continue

        eski_deger = row.get(kod_kolonu, "")
        eski_deger_str = str(eski_deger).strip() if not pd.isna(eski_deger) else ""

        if not eski_deger_str:
            df.at[idx, durum_kolon] = "Boş"
            df.at[idx, not_kolon] = "Ürün kodu boş"
            bos_kod += 1

            bos_satir = {
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonundaki Değer": eski_deger_str,
                "Kalan Grup": "Boş ürün kodu",
                "Durum": "Ürün kodu boş olduğu için çevrilemedi",
            }
            bos_kod_satirlari.append(bos_satir)
            geriye_kalan_satirlar.append(bos_satir)
            continue

        if sade_kod(eski_deger_str) in kirac_model_karsiligi_bulunamayan_sade_kodlar:
            df.at[idx, durum_kolon] = "Model karşılığı bulunamadı"
            df.at[idx, not_kolon] = "Bu KIRAC kodu için model karşılığı bulunamadığı belirtildi; ürün koduna çevrilmedi."
            bulunamayan += 1

            bulunamayan_satir = {
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonundaki Değer": eski_deger_str,
                "Kalan Grup": "Modeli bulunamayan",
                "Durum": "Model karşılığı bulunamadı; ürün koduna çevrilmedi",
            }
            modeli_bulunmayan_satirlar.append(bulunamayan_satir)
            geriye_kalan_satirlar.append(bulunamayan_satir)
            continue

        satir_anahtarlari = anahtar_listesi(eski_deger_str)

        # Önce kural bazlı model eşleşmeleri denenir.
        # Bu adım M kodu elle yazmaz; ENTES fiyat listesinde hedef model satırını bulur
        # ve o satırın Ürün Kodu değerini kullanır.
        kural_kod = kirac_kural_modelden_kod_bul(eski_deger_str)
        if kural_kod:
            df.at[idx, kod_kolonu] = kural_kod
            df.at[idx, durum_kolon] = "Kuraldan modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger_str} -> {kural_kod}"
            cevrilen += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": kural_kod,
                "Durum": "Kuraldan fiyat listesindeki model satırından çevrildi",
            })

            modelden_bulunan_satirlar.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonunda Yazılan Model": eski_deger_str,
                "Bulunan / Yazılan Ürün Kodu": kural_kod,
                "Durum": "Kuraldan fiyat listesindeki model satırından ürün koduna çevrildi",
            })

            continue

        # Değer zaten ENTES fiyat listesindeki gerçek ürün kodlarından biriyse dokunma.
        if any(anahtar in kod_anahtar_set for anahtar in satir_anahtarlari):
            df.at[idx, durum_kolon] = "Zaten ürün kodu"
            zaten_kod += 1

            urun_kodu_satir = {
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Verilen Ürün Kodu": eski_deger_str,
                "Kalan Grup": "Zaten ürün kodu",
                "Durum": "Zaten ürün kodu olarak verilmiş, değişiklik yapılmadı",
            }
            urun_kodu_olan_satirlar.append(urun_kodu_satir)

            continue

        # AB ailesi için özel kısa kod çözümü. Bu, genel CT anahtarlarının çokluya düşmesini engeller.
        ab_ozel_kod = ab_akim_trafo_ozel_kod_bul(eski_deger_str)
        if ab_ozel_kod:
            df.at[idx, kod_kolonu] = ab_ozel_kod
            df.at[idx, durum_kolon] = "AB akım trafosu modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger_str} -> {ab_ozel_kod}"
            cevrilen += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": ab_ozel_kod,
                "Durum": "AB özel kuralıyla çevrildi",
            })

            modelden_bulunan_satirlar.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonunda Yazılan Model": eski_deger_str,
                "Bulunan / Yazılan Ürün Kodu": ab_ozel_kod,
                "Durum": "AB akım trafosu özel kuralıyla ürün koduna çevrildi",
            })

            continue

        # CM kapak ürünleri için özel model çözümü.
        # Örnek: CM-8,33-KAPAK -> fiyat listesindeki "CM serisi 8,33kVAr" satırı.
        cm_ozel_kod = cm_serisi_ozel_kod_bul(eski_deger_str)
        if cm_ozel_kod:
            df.at[idx, kod_kolonu] = cm_ozel_kod
            df.at[idx, durum_kolon] = "CM serisi modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger_str} -> {cm_ozel_kod}"
            cevrilen += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": cm_ozel_kod,
                "Durum": "CM serisi özel kuralıyla çevrildi",
            })

            modelden_bulunan_satirlar.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonunda Yazılan Model": eski_deger_str,
                "Bulunan / Yazılan Ürün Kodu": cm_ozel_kod,
                "Durum": "CM serisi özel kuralıyla ürün koduna çevrildi",
            })

            continue

        # MPR V2 ürünleri fiyat listesinde IP54 yazımıyla geçebilir.
        # Örnek: MPR45S-V2 -> MPR-45S-IP54, MPR47SE-V2 -> MPR-47SE-IP54.
        mpr_ozel_kod = mpr_v2_ip54_ozel_kod_bul(eski_deger_str)
        if mpr_ozel_kod:
            df.at[idx, kod_kolonu] = mpr_ozel_kod
            df.at[idx, durum_kolon] = "MPR V2/IP54 modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger_str} -> {mpr_ozel_kod}"
            cevrilen += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": mpr_ozel_kod,
                "Durum": "MPR V2/IP54 özel kuralıyla çevrildi",
            })

            modelden_bulunan_satirlar.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonunda Yazılan Model": eski_deger_str,
                "Bulunan / Yazılan Ürün Kodu": mpr_ozel_kod,
                "Durum": "MPR V2/IP54 özel kuralıyla ürün koduna çevrildi",
            })

            continue

        # ENS özel kısa kodları M kodu manuel yazılmadan fiyat listesinden çözülür.
        # Örnek: ENS-CCT24-300-5 -> fiyat listesindeki ENS.CCT 24 300/5A satırının ürün kodu.
        ens_ozel_kod = ens_ozel_kod_bul(eski_deger_str)
        if ens_ozel_kod:
            df.at[idx, kod_kolonu] = ens_ozel_kod
            df.at[idx, durum_kolon] = "ENS modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger_str} -> {ens_ozel_kod}"
            cevrilen += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": ens_ozel_kod,
                "Durum": "ENS özel kuralıyla fiyat listesinden çevrildi",
            })

            modelden_bulunan_satirlar.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonunda Yazılan Model": eski_deger_str,
                "Bulunan / Yazılan Ürün Kodu": ens_ozel_kod,
                "Durum": "ENS özel kuralıyla fiyat listesinden ürün koduna çevrildi",
            })

            continue

        bulunan_kodlar = set()

        # Model aramasını öncelikli gruplarla yapıyoruz.
        # Önce birebir/güçlü anahtarlar denenir, olmazsa akım trafosu için daha esnek anahtarlar denenir.
        # Böylece AB-50-5 gibi kısa KIRAC yazımları, ENT.AB 50- 10 VA CL:0,5 gibi fiyat listesi yazımlarını bulabilir.
        for anahtar_grubu in anahtar_gruplari(eski_deger_str):
            grup_bulunan_kodlar = set()

            for anahtar in anahtar_grubu:
                grup_bulunan_kodlar.update(model_to_kod_set.get(anahtar, set()))

            if grup_bulunan_kodlar:
                bulunan_kodlar = grup_bulunan_kodlar
                break

        if len(bulunan_kodlar) == 1:
            yeni_kod = list(bulunan_kodlar)[0]
            df.at[idx, kod_kolonu] = yeni_kod
            df.at[idx, durum_kolon] = "Modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger_str} -> {yeni_kod}"
            cevrilen += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": yeni_kod,
                "Durum": "Çevrildi",
            })

            modelden_bulunan_satirlar.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonunda Yazılan Model": eski_deger_str,
                "Bulunan / Yazılan Ürün Kodu": yeni_kod,
                "Durum": "Ürün kodu kolonunda model bulundu ve ürün koduna çevrildi",
            })

        elif len(bulunan_kodlar) > 1:
            df.at[idx, durum_kolon] = "Çoklu model eşleşmesi"
            df.at[idx, not_kolon] = ", ".join(sorted(bulunan_kodlar))
            coklu += 1

            rapor_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger_str,
                "Yeni Ürün Kodu": "",
                "Durum": "Çoklu eşleşme",
                "Bulunan Kodlar": ", ".join(sorted(bulunan_kodlar)),
            })

            coklu_satir = {
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonundaki Değer": eski_deger_str,
                "Bulunan Kodlar": ", ".join(sorted(bulunan_kodlar)),
                "Kalan Grup": "Çoklu model eşleşmesi",
                "Durum": "Çoklu model eşleşmesi - manuel kontrol gerekli",
            }
            coklu_model_satirlari.append(coklu_satir)
            geriye_kalan_satirlar.append(coklu_satir)

        else:
            df.at[idx, durum_kolon] = "Model karşılığı bulunamadı"
            bulunamayan += 1

            bulunamayan_satir = {
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Ürün Kodu Kolonundaki Değer": eski_deger_str,
                "Kalan Grup": "Modeli bulunamayan",
                "Durum": "Ne ürün kodu olarak bulundu ne de model karşılığı bulundu",
            }
            modeli_bulunmayan_satirlar.append(bulunamayan_satir)
            geriye_kalan_satirlar.append(bulunamayan_satir)

    rapor_df = pd.DataFrame(rapor_satirlari)

    kontrol_edilen_sayi = int(kontrol_mask.sum())
    bulunan_toplam = zaten_kod + cevrilen
    geriye_kalan_toplam = bulunamayan + coklu + bos_kod
    kontrol_toplam = bulunan_toplam + geriye_kalan_toplam

    rapor_zamani = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    sayi_df = pd.DataFrame([
        {
            "Kalem": "Rapor oluşturma zamanı",
            "Adet": rapor_zamani,
            "Açıklama": f"Bu dosya bu çalıştırmada oluşturuldu. Kod sürümü: {KOD_SURUMU}",
        },
        {
            "Kalem": "Toplam kontrol edilen KIRAC ENTES satırı",
            "Adet": kontrol_edilen_sayi,
            "Açıklama": "Marka filtresine giren ve KIRAC düzeltme kontrolüne dahil edilen satır sayısı. Diğer markalar bu sayıya dahil edilmez.",
        },
        {
            "Kalem": "Bulunan toplam",
            "Adet": bulunan_toplam,
            "Açıklama": "Ürün kodu olarak bulunanlar + modelden ürün koduna çevrilenler. Yani sorun çıkarmayan toplam satır.",
        },
        {
            "Kalem": "Ürün kodu olarak verilenlerin sayısı",
            "Adet": zaten_kod,
            "Açıklama": "KIRAC ürün kodu kolonundaki değer ENTES fiyat listesinde zaten ürün kodu olarak bulundu, değiştirilmedi.",
        },
        {
            "Kalem": "Ürün kodu kolonunda model bulunup düzeltilenlerin sayısı",
            "Adet": cevrilen,
            "Açıklama": "KIRAC ürün kodu kolonundaki değer ENTES fiyat listesinde model olarak bulundu ve ürün koduna çevrildi.",
        },
        {
            "Kalem": "Geriye kalan toplam",
            "Adet": geriye_kalan_toplam,
            "Açıklama": "Bulunan toplam dışında kalan satırlar. Dağılımı aşağıdaki satırlarda ve GeriyeKalanlar sayfasında görülür.",
        },
        {
            "Kalem": "Geriye kalan - modeli bulunamayanların sayısı",
            "Adet": bulunamayan,
            "Açıklama": "KIRAC ürün kodu kolonundaki değer ENTES fiyat listesinde ne ürün kodu ne de model olarak bulunamadı.",
        },
        {
            "Kalem": "Geriye kalan - çoklu model eşleşmesi sayısı",
            "Adet": coklu,
            "Açıklama": "Aynı model ENTES fiyat listesinde birden fazla ürün koduna denk geldi, manuel kontrol gerekir.",
        },
        {
            "Kalem": "Geriye kalan - boş ürün kodu sayısı",
            "Adet": bos_kod,
            "Açıklama": "KIRAC ürün kodu kolonu boş olan satırlar.",
        },
        {
            "Kalem": "Kontrol toplamı",
            "Adet": kontrol_toplam,
            "Açıklama": "Bulunan toplam + geriye kalan toplam. Bu sayı Toplam kontrol edilen KIRAC ENTES satırı ile aynı olmalıdır.",
        },
    ])

    modelden_bulunan_df = pd.DataFrame(modelden_bulunan_satirlar)
    modeli_bulunmayan_df = pd.DataFrame(modeli_bulunmayan_satirlar)
    urun_kodu_olan_df = pd.DataFrame(urun_kodu_olan_satirlar)
    coklu_model_df = pd.DataFrame(coklu_model_satirlari)
    bos_kod_df = pd.DataFrame(bos_kod_satirlari)
    geriye_kalan_df = pd.DataFrame(geriye_kalan_satirlar)

    return df



def santral_model_kod_duzelt(df, kod_kolonu, kaynak_dosya, kontrol_mask=None):
    """
    SANTRAL için çalışır.
    SANTRAL B2B'de ürün kodu alanında ENTES ürün modeli/model numarası geldiği için,
    ENTES fiyat listesinden Model -> Ürün Kodu dönüşümü yapar.

    Bu fonksiyon M kodu ezberlemez; fiyat listesindeki ÜRÜN MODELİ satırını bulur,
    karşısındaki ÜRÜN KODU neyse onu kullanır.
    """
    fiyat_dosyasi = entes_fiyat_listesi_bul()

    orijinal_kolon = "SANTRAL Orijinal Ürün Kodu"
    durum_kolon = "SANTRAL Kod Düzeltme Durumu"
    not_kolon = "SANTRAL Kod Düzeltme Notu"

    df = df.copy()

    if orijinal_kolon not in df.columns:
        df[orijinal_kolon] = df[kod_kolonu]
    if durum_kolon not in df.columns:
        df[durum_kolon] = ""
    if not_kolon not in df.columns:
        df[not_kolon] = ""

    if kontrol_mask is None:
        kontrol_mask = pd.Series([True] * len(df), index=df.index)
    else:
        kontrol_mask = pd.Series(kontrol_mask, index=df.index).fillna(False).astype(bool)

    if fiyat_dosyasi is None:
        print("SANTRAL düzeltmesi için ENTES fiyat listesi bulunamadı.")
        print("Fiyat listesini şu klasöre koy:")
        print(FIYAT_LISTESI_DIR)
        print("Örnek dosya adı: ENTES_FIYAT_LISTESI.xlsx")
        raise SystemExit

    entes_df = entes_fiyat_listesi_oku(fiyat_dosyasi)

    entes_model_col = kolon_bul_esnek(
        entes_df,
        [
            "Ürün Modeli",
            "Urun Modeli",
            "Model",
            "Ürün Model",
            "Urun Model",
            "Model Kodu",
            "Tip",
        ]
    )

    entes_kod_col = kolon_bul_esnek(
        entes_df,
        [
            "Ürün Kodu",
            "Urun Kodu",
            "Stok Kodu",
            "Kod",
            "Malzeme Kodu",
            "Web Servis Kodu",
            "Ürün No",
            "Urun No",
        ]
    )

    if not entes_model_col or not entes_kod_col:
        print("ENTES fiyat listesinde Ürün Kodu / Ürün Modeli kolonları bulunamadı.")
        print("Gelen kolonlar:")
        for c in entes_df.columns:
            print("-", c)
        raise SystemExit

    def santral_anahtarlar_uret(deger):
        """
        SANTRAL için kontrollü anahtar üretir.
        Sadece tam model/kod yazım farklarını temizler; tek başına sayı benzerliğiyle eşleştirme yapmaz.
        """
        anahtarlar = []
        gorulen = set()

        def ekle(deger2):
            deger2 = kod_temizle(deger2)
            if not deger2:
                return
            if deger2 not in gorulen:
                gorulen.add(deger2)
                anahtarlar.append(deger2)

        raw = kod_temizle(deger)
        prefixsiz = entes_prefix_sil(raw)
        temiz = b2b_temiz_kod_uret(raw)

        ekle(raw)
        ekle(prefixsiz)
        ekle(temiz)

        # Nokta / tire / boşluk farklarını kaldırılmış tam model anahtarı.
        for item in [raw, prefixsiz, temiz]:
            sade = sade_kod(item)
            # Sadece çok kısa sayısal anahtarları kullanma; yanlış eşleşme riski yüksek.
            if sade and len(sade) >= 4 and not sade.isdigit():
                ekle(sade)
            elif sade and len(sade) >= 5 and sade.isdigit():
                ekle(sade)

        return anahtarlar

    urun_kodu_lookup = {}
    model_lookup = {}

    for _, row in entes_df.iterrows():
        kod_raw = str(row.get(entes_kod_col, "")).strip()
        model_raw = str(row.get(entes_model_col, "")).strip()

        if not kod_raw:
            continue

        for anahtar in santral_anahtarlar_uret(kod_raw):
            urun_kodu_lookup.setdefault(anahtar, set()).add(kod_raw)

        if model_raw:
            for anahtar in santral_anahtarlar_uret(model_raw):
                model_lookup.setdefault(anahtar, set()).add((kod_raw, model_raw))

    def santral_kablo_bagi_parcala(deger):
        """
        SANTRAL kablo bağı yazımını parçalar.
        Örnekler:
        - 2,5x80 KABLO BAĞI B  -> genişlik 2,5, boy 80, renk Beyaz
        - 2,5x80 KABLO BAĞI S  -> genişlik 2,5, boy 80, renk Siyah
        - 3,6x280 KABLO BAĞI SARI -> genişlik 3,6, boy 280, renk Sarı
        - 2,5x100 UV K.BAĞI S -> genişlik 2,5, boy 100, UV, renk Siyah

        Burada M kodu kullanılmaz. Eşleşme fiyat listesindeki model metninden yapılır.
        """

        raw_norm = metin_norm(deger)

        eslesme = re.search(
            r"(\d+(?:[,.]\d+)?)\s*[Xx×]\s*(\d+)\s*(?:MM)?\s*(UV\s*)?(?:KABLO|K)\.?\s*BAGI\s*([A-Z0-9]*)",
            raw_norm
        )

        if not eslesme:
            return None

        genislik = eslesme.group(1).replace(".", ",")
        boy = str(int(eslesme.group(2)))
        uv_var = bool(eslesme.group(3)) or "UV" in raw_norm
        renk_token = metin_norm(eslesme.group(4) or "")

        renk_map = {
            "B": "BEYAZ",
            "BEYAZ": "BEYAZ",
            "N": "BEYAZ",
            "W": "BEYAZ",
            "S": "SIYAH",
            "SIYAH": "SIYAH",
            "SYH": "SIYAH",
            "BLACK": "SIYAH",
            "Y": "SARI",
            "SARI": "SARI",
            "R": "KIRMIZI",
            "RED": "KIRMIZI",
            "K": "KIRMIZI",
            "KIRMIZI": "KIRMIZI",
            "KIRMZI": "KIRMIZI",
            "KRMIZI": "KIRMIZI",
            "KRMZI": "KIRMIZI",
            "KRMZ": "KIRMIZI",
            "KIR": "KIRMIZI",
            "BL": "MAVI",
            "MAVI": "MAVI",
            "MAVİ": "MAVI",
            "P": "MOR",
            "MOR": "MOR",
            "G": "YESIL",
            "YESIL": "YESIL",
            "YEŞIL": "YESIL",
            "YESİL": "YESIL",
        }

        renk = renk_map.get(renk_token, "")

        return {
            "genislik": genislik,
            "genislik_nokta": genislik.replace(",", "."),
            "boy": boy,
            "renk": renk,
            "renk_token": renk_token,
            "uv": uv_var,
        }

    def santral_kablo_bagi_genislik_adaylari(parca):
        """
        SANTRAL kablo bağı ölçülerinde bazı genişlikler yuvarlanmış veya eksik yazılabiliyor.
        Burada M kodu ezberlenmez; sadece fiyat listesindeki ölçü metnine ulaşmak için
        güvenli genişlik adayları üretilir.

        Örnekler:
        - 9,0x920 / 9x920  -> 8,8x920
        - 9x1120          -> 8,8x1120
        - 2,6x1030        -> 12,6x1030  (uzun kablo bağında baştaki 1 eksik gelebiliyor)
        """

        adaylar = []
        gorulen = set()

        def ekle(x):
            x = str(x).strip().replace(".", ",")
            if not x:
                return
            if x not in gorulen:
                gorulen.add(x)
                adaylar.append(x)

        genislik = str(parca.get("genislik", "")).strip().replace(".", ",")
        boy = sayiya_cevir(parca.get("boy", "0"), varsayilan=0)

        ekle(genislik)

        try:
            genislik_float = float(genislik.replace(",", "."))
        except Exception:
            genislik_float = None

        # SANTRAL bazen 8,8 mm kablo bağını 9 / 9,0 olarak yazıyor.
        # Fiyat listesinde ürün 8,8x... olarak geçiyor.
        if genislik_float is not None and abs(genislik_float - 9.0) < 0.11:
            ekle("8,8")

        # Uzun 12,6 mm kablo bağında SANTRAL tarafında 2,6x1030 gibi baştaki 1 düşebiliyor.
        # Bunu sadece uzun boylarda uygula ki 2,5 / 3,6 aileleriyle karışmasın.
        if genislik_float is not None and abs(genislik_float - 2.6) < 0.11 and boy >= 650:
            ekle("12,6")

        return adaylar

    def santral_cirt_cirtli_kablo_bagi_fiyat_listesinden_bul_parca(parca):
        """
        SANTRAL cırt cırtlı kablo bağlarını fiyat listesindeki model metninden bulur.

        Örnekler:
        - 12,5x100 K.BAĞI BEYAZ  -> 12,5X130mm Cırt Cırtlı Kablo Bağları, Beyaz
        - 12,5x200 K.BAĞI SİYAH  -> 12,5X200mm Cırt Cırtlı Kablo Bağları, Siyah
        - 17,0x300 K.BAĞI KIRMIZI -> 17.0X300mm Cırt Cırtlı Kablo Bağları, Kırmızı

        M kodu veya TS1-2 kodu ezberlenmez. Ölçü + renk + "Cırt Cırtlı" model metni
        fiyat listesinde bulunur ve o satırdaki Ürün Kodu alınır.
        """

        if not parca or parca.get("uv"):
            return "", ""

        renk = parca.get("renk", "")

        # Cırt cırtlı kablo bağı tarafında şu an güvenli renkler bunlar.
        if renk not in {"BEYAZ", "SIYAH", "KIRMIZI"}:
            return "", ""

        genislik = str(parca.get("genislik", "")).strip().replace(".", ",")
        boy = sayiya_cevir(parca.get("boy", "0"), varsayilan=0)

        try:
            genislik_float = float(genislik.replace(",", "."))
        except Exception:
            genislik_float = None

        genislik_adaylari = []
        boy_adaylari = []

        if genislik_float is not None and abs(genislik_float - 12.5) < 0.11:
            genislik_adaylari = ["12,5", "12.5"]

            # SANTRAL tarafında 12,5x100 diye gelen ürün fiyat listesinde
            # 12,5X130mm cırt cırtlı kablo bağı olarak geçiyor.
            if boy == 100:
                boy_adaylari = ["130"]
            elif boy in {130, 200}:
                boy_adaylari = [str(boy)]
            else:
                return "", ""

        elif genislik_float is not None and abs(genislik_float - 17.0) < 0.11:
            genislik_adaylari = ["17,0", "17.0", "17"]

            if boy == 300:
                boy_adaylari = ["300"]
            else:
                return "", ""
        else:
            return "", ""

        olcu_anahtarlari = set()

        for g in genislik_adaylari:
            for b in boy_adaylari:
                olcu_anahtarlari.add(metin_norm(f"{g}x{b}mm").replace(" ", ""))
                olcu_anahtarlari.add(metin_norm(f"{g} x {b} mm").replace(" ", ""))

        adaylar = {}

        # Fiyat listesinde cırt cırtlı grup A kolonunda / ÜRÜN GRUBU kolonunda duruyor.
        # Model satırında da TS1-2 ve Cırt Cırtlı metni var. İkisini beraber kullanıyoruz.
        entes_grup_col_local = None
        for col in entes_df.columns:
            if metin_norm(col) == "URUN GRUBU":
                entes_grup_col_local = col
                break

        for _, entes_row in entes_df.iterrows():
            kod_raw = str(entes_row.get(entes_kod_col, "")).strip()
            model_raw = str(entes_row.get(entes_model_col, "")).strip()
            grup_raw = str(entes_row.get(entes_grup_col_local, "")).strip() if entes_grup_col_local else ""

            if not kod_raw or not model_raw:
                continue

            model_norm_bosluklu = metin_norm(model_raw)
            model_norm_bosuksuz = model_norm_bosluklu.replace(" ", "")
            model_sade = sade_kod(model_raw)
            grup_norm_bosluklu = metin_norm(grup_raw)

            cirt_grubu_mu = (
                "CIRT CIRTLI" in grup_norm_bosluklu
                or "CIRT CIRTLI" in model_norm_bosluklu
                or model_sade.startswith("TS12")
            )

            if not cirt_grubu_mu:
                continue

            if "KABLO BAG" not in model_norm_bosluklu and "KABLO BAG" not in grup_norm_bosluklu:
                continue

            # Cırt cırtlı kablo bağları fiyat listesinde TS1-2 ile başlar.
            # Bu kontrol normal TS1-1 kablo bağlarıyla karışmayı engeller.
            if not model_sade.startswith("TS12"):
                continue

            if not any(olcu in model_norm_bosuksuz for olcu in olcu_anahtarlari):
                continue

            if renk not in model_norm_bosluklu:
                continue

            adaylar.setdefault(kod_raw, model_raw)

        if len(adaylar) == 1:
            yeni_kod = list(adaylar.keys())[0]
            bulunan_model = adaylar[yeni_kod]
            return yeni_kod, bulunan_model

        return "", ""

    def santral_ozel_kablo_bagi_fiyat_listesinden_bul(deger):
        """
        SANTRAL özel kablo bağı ailelerini ölçü + tip + renk üzerinden fiyat listesinden bulur.

        M kodu veya TS kodu ezberlenmez. SANTRAL'deki yazımdan ölçü ve ürün tipi okunur;
        ENTES fiyat listesinde aynı ölçü, aynı tip ve renk aranır.

        Örnekler:
        - 7,6x200 AÇILAN K.BAĞ   -> 7,6x200mm Kablo Bağı, Açılabilir, Beyaz
        - 2,5x100 ETİKET K.BAĞ   -> 2,5x100mm Kablo Bağı, Etiketli, Beyaz
        - 4,8x190 ETİKET K.BAĞ   -> 4,8x190mm Kablo Bağı, Etiketli, Beyaz
        - 2,5x100 BONCUK K.BAĞ   -> 2,5x100mm Kablo Bağı, Boncuk Tip, Beyaz
        """

        raw_norm = metin_norm(deger)

        if "K.BAG" not in raw_norm and "KABLO BAG" not in raw_norm:
            return "", ""

        tip = ""
        if "ACILAN" in raw_norm or "ACILABILIR" in raw_norm:
            tip = "ACILABILIR"
        elif "ETIKET" in raw_norm:
            tip = "ETIKETLI"
        elif "BONCUK" in raw_norm:
            tip = "BONCUK TIP"

        if not tip:
            return "", ""

        eslesme = re.search(
            r"(\d+(?:[,.]\d+)?)\s*[Xx×]\s*(\d+)",
            raw_norm
        )

        if not eslesme:
            return "", ""

        genislik = eslesme.group(1).replace(".", ",")
        boy = str(int(eslesme.group(2)))

        renk = "BEYAZ"
        if "SIYAH" in raw_norm or re.search(r"\bS\b", raw_norm):
            renk = "SIYAH"
        elif (
            "KIRMIZI" in raw_norm
            or "KIRMZI" in raw_norm
            or "KRMIZI" in raw_norm
            or "KRMZI" in raw_norm
            or "KRMZ" in raw_norm
            or re.search(r"\bK\b", raw_norm)
        ):
            renk = "KIRMIZI"
        elif "SARI" in raw_norm:
            renk = "SARI"
        elif "MAVI" in raw_norm:
            renk = "MAVI"
        elif "MOR" in raw_norm:
            renk = "MOR"
        elif "YESIL" in raw_norm:
            renk = "YESIL"

        parca = {
            "genislik": genislik,
            "boy": boy,
            "renk": renk,
            "uv": False,
        }

        olcu_anahtarlari = set()

        for genislik_adayi in santral_kablo_bagi_genislik_adaylari(parca):
            genislik_nokta = str(genislik_adayi).replace(",", ".")

            for g in [genislik_adayi, genislik_nokta]:
                olcu_anahtarlari.add(metin_norm(f"{g}x{boy}mm").replace(" ", ""))
                olcu_anahtarlari.add(metin_norm(f"{g} x {boy} mm").replace(" ", ""))

        if tip == "ACILABILIR":
            grup_ipucu = "ACILABILIR"
            ts_prefix = "TS18"
            model_ipucu = "ACILABILIR"
        elif tip == "ETIKETLI":
            grup_ipucu = "ETIKETLI"
            ts_prefix = "TS19"
            model_ipucu = "ETIKETLI"
        else:
            grup_ipucu = "BONCUK"
            ts_prefix = "TS113"
            model_ipucu = "BONCUK"

        entes_grup_col_local = None
        for col in entes_df.columns:
            if metin_norm(col) == "URUN GRUBU":
                entes_grup_col_local = col
                break

        adaylar = {}

        for _, entes_row in entes_df.iterrows():
            kod_raw = str(entes_row.get(entes_kod_col, "")).strip()
            model_raw = str(entes_row.get(entes_model_col, "")).strip()
            grup_raw = str(entes_row.get(entes_grup_col_local, "")).strip() if entes_grup_col_local else ""

            if not kod_raw or not model_raw:
                continue

            model_norm_bosluklu = metin_norm(model_raw)
            model_norm_bosluksuz = model_norm_bosluklu.replace(" ", "")
            model_sade = sade_kod(model_raw)
            grup_norm_bosluklu = metin_norm(grup_raw)

            if "KABLO BAG" not in model_norm_bosluklu and "KABLO BAG" not in grup_norm_bosluklu:
                continue

            # Aileyi mutlaka kontrol et. Sadece ölçü benziyor diye normal / UV / cırt cırtlı
            # kablo bağlarıyla karışmasın.
            aile_uygun = (
                grup_ipucu in grup_norm_bosluklu
                or model_ipucu in model_norm_bosluklu
                or model_sade.startswith(ts_prefix)
            )

            if not aile_uygun:
                continue

            if not model_sade.startswith(ts_prefix):
                continue

            if not any(olcu in model_norm_bosluksuz for olcu in olcu_anahtarlari):
                continue

            if renk not in model_norm_bosluklu:
                continue

            adaylar.setdefault(kod_raw, model_raw)

        if len(adaylar) == 1:
            yeni_kod = list(adaylar.keys())[0]
            bulunan_model = adaylar[yeni_kod]
            return yeni_kod, bulunan_model

        return "", ""

    def santral_kablo_bagi_fiyat_listesinden_bul(deger):
        """
        Kablo bağlarında M kodlarını dikkate almadan, ölçü + renk üzerinden fiyat listesindeki
        Ürün Modeli satırını bulur.

        Fiyat listesi örnekleri:
        TS1-1-25080N 2,5x80mm Kablo Bağı, Beyaz
        TS1-1-25080B 2,5x80mm Kablo Bağı, Siyah
        TS1-1-36280Y 3,6x280mm Kablo Bağı, Sarı
        TS1-1-90920N 8,8x920mm Kablo Bağı, Beyaz
        TS1-1-12610N 12,6x1030mm Kablo Bağı, Beyaz
        TS1-1-25100UVB 2,5x100mm Kablo Bağı, UV'ye Dayanıklı, Siyah

        Kod burada TS1-1 kodunu veya M kodunu ezberlemez; asıl güvenli kontrol
        model metnindeki ölçü ve renktir.
        """

        ozel_kod, ozel_model = santral_ozel_kablo_bagi_fiyat_listesinden_bul(deger)
        if ozel_kod:
            return ozel_kod, ozel_model

        parca = santral_kablo_bagi_parcala(deger)

        if not parca:
            return "", ""

        # Cırt cırtlı kablo bağı özel ailesini normal/UV kablo bağından önce dene.
        # Böylece 12,5 / 17,0 ailesi TS1-1 normal kablo bağıyla karışmaz.
        cirt_kod, cirt_model = santral_cirt_cirtli_kablo_bagi_fiyat_listesinden_bul_parca(parca)
        if cirt_kod:
            return cirt_kod, cirt_model

        olcu_anahtarlari = set()

        for genislik in santral_kablo_bagi_genislik_adaylari(parca):
            genislik_nokta = str(genislik).replace(",", ".")

            for g in [genislik, genislik_nokta]:
                olcu_anahtarlari.add(metin_norm(f"{g}x{parca['boy']}mm").replace(" ", ""))
                olcu_anahtarlari.add(metin_norm(f"{g} x {parca['boy']} mm").replace(" ", ""))

        renk = parca["renk"]
        uv_isteniyor = bool(parca.get("uv"))

        adaylar = {}

        for _, entes_row in entes_df.iterrows():
            kod_raw = str(entes_row.get(entes_kod_col, "")).strip()
            model_raw = str(entes_row.get(entes_model_col, "")).strip()

            if not kod_raw or not model_raw:
                continue

            model_norm_bosluklu = metin_norm(model_raw)
            model_norm_bosuksuz = model_norm_bosluklu.replace(" ", "")
            model_sade = sade_kod(model_raw)

            if "KABLO BAGI" not in model_norm_bosluklu:
                continue

            # Kablo bağı ürünleri fiyat listesinde TS1-1 ile başlıyor.
            # Bunu kontrol ederek aynı ölçüde başka ürünlerin karışmasını engelliyoruz.
            if not model_sade.startswith("TS11"):
                continue

            if not any(olcu in model_norm_bosuksuz for olcu in olcu_anahtarlari):
                continue

            if renk and renk not in model_norm_bosluklu:
                continue

            model_uv_mu = ("UV" in model_norm_bosluklu) or ("DAYANIKLI" in model_norm_bosluklu)

            # SANTRAL'de "UV K.BAĞI" yazıyorsa yalnızca UV'ye dayanıklı kablo bağıyla eşleştir.
            # UV yazmıyorsa aynı ölçü/renkteki normal kablo bağını UV seriyle karıştırma.
            if uv_isteniyor and not model_uv_mu:
                continue

            if not uv_isteniyor and model_uv_mu:
                continue

            adaylar.setdefault(kod_raw, model_raw)

        if len(adaylar) == 1:
            yeni_kod = list(adaylar.keys())[0]
            bulunan_model = adaylar[yeni_kod]
            return yeni_kod, bulunan_model

        return "", ""

    def santral_kural_model_hedefleri(deger):
        """
        SANTRAL tarafındaki bazı tedarikçi yazımlarını ENTES fiyat listesindeki
        model yazımına çevirir. M kodu ezberlemez; hedef model fiyat listesinden
        bulunur ve o satırın Ürün Kodu kullanılır.
        """
        raw_sade = sade_kod(deger)
        hedefler = []

        # SRS3 reaktif güç kontrol modülü yazım farkı:
        # SANTRAL: ENT.SRS3 400/1   -> ENTES model: ENT.SRS3 400/1 400V 1 kVar
        # SANTRAL: ENT.SRS3 400/1,5 -> ENTES model: ENT.SRS3 400/1,5 400V 1,5 kVar
        # Burada 400 ve 1 / 1,5 gibi değerler sabit değildir; parçalanıp hedef model üretilir.
        # M kodu elle yazılmaz; fiyat listesindeki model satırından Ürün Kodu alınır.
        srs3_eslesme = re.search(
            r"(?:ENT|ENTES)?\s*\.?\s*SRS\s*3[\s\-]*(\d+)\s*/\s*([0-9]+(?:[,.][0-9]+)?)",
            kod_temizle(deger)
        )
        if srs3_eslesme:
            srs3_gerilim = srs3_eslesme.group(1)
            srs3_kvar = srs3_eslesme.group(2).replace(".", ",")
            hedefler.extend([
                f"ENT.SRS3 {srs3_gerilim}/{srs3_kvar} {srs3_gerilim}V {srs3_kvar} kVar",
                f"ENT.SRS3 {srs3_gerilim}/{srs3_kvar} {srs3_gerilim} V {srs3_kvar} kVar",
                f"SRS3 {srs3_gerilim}/{srs3_kvar} {srs3_gerilim}V {srs3_kvar} kVar",
                f"ENT.SRS3 {srs3_gerilim}/{srs3_kvar}",
            ])

        # SRS1 shunt/şönt reaktör yazım farkı:
        # SANTRAL: ENT.SRS1.230/2   -> ENTES model: ENT.SRS1 230/2 230V 2 kVAr Shunt Reaktor
        # SANTRAL: ENT.SRS1 230/7,5 -> ENTES model: ENT.SRS1 230/7,5 230V 7,5 kVAr Şönt Reaktör
        # Türkçe / İngilizce açıklama farkını destekler. M kodu elle yazılmaz.
        hedefler.extend(santral_srs1_hedefleri(deger))

        # ERH-E decimal yazımlar birebir modeldir:
        # SANTRAL: ENT.ERH-E50-7-400-1,25 -> ENTES model: ENT.ERH-E50-7-400-1,25
        # SANTRAL: ENT.ERH-E50-7-400-1,5  -> ENTES model: ENT.ERH-E50-7-400-1,5
        # Bu iki tipte sade anahtar virgülü kaldırınca 12,5 gibi başka değerlerle karışabilir.
        # Bu yüzden önce birebir hedef model verilir, çoklu eşleşmeye düşürülmez.
        erh_e_decimal_eslesme = re.search(
            r"(?:ENT|ENTES)?\s*[.\-]?\s*ERH\s*[- ]\s*E\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*([0-9]+(?:[,.][0-9]+)?)",
            kod_temizle(deger),
            flags=re.IGNORECASE,
        )
        if erh_e_decimal_eslesme:
            e_govde, e_sira, e3, e4 = erh_e_decimal_eslesme.groups()
            e4 = e4.replace(".", ",")
            hedefler.extend([
                f"ENT.ERH-E{e_govde}-{e_sira}-{e3}-{e4}",
                f"ENT ERH E{e_govde} {e_sira} {e3} {e4}",
            ])

        # ERH / SC özel yazım kuralları:
        # SANTRAL: ENT.ERH-50-7-400-4 -> ENTES model: ENT.ERH-E50-7-400-4
        # Kural: ENT.ERH-X-... -> ENT.ERH-EX-...
        # SANTRAL: ENT-SC-225-25  -> ENTES model: ENT-SC-225
        # SANTRAL: ENT-SC-225S-25 -> ENTES model: ENT-SC-225S
        # Kural: ENT-SC-<gövde>-<son ek> -> ENT-SC-<gövde>
        # M kodu elle yazılmaz; hedef model fiyat listesinden bulunur ve o satırdaki Ürün Kodu alınır.
        erh_eslesme = re.fullmatch(r"(?:ENT|ENTES)?ERH(\d+)(\d+)(\d+)(\d+)", raw_sade)
        if erh_eslesme:
            erh_govde = erh_eslesme.group(1)
            erh_parca2 = erh_eslesme.group(2)
            erh_parca3 = erh_eslesme.group(3)
            erh_parca4 = erh_eslesme.group(4)
            # Raw sade anahtarda parça sınırları kaybolduğu için, normalleştirilmiş metinden
            # tireli yazımı da yakala; başarırsa daha güvenli hedef üret.
            erh_text_eslesme = re.search(
                r"(?:ENT|ENTES)?\s*[.\-]?\s*ERH\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)",
                kod_temizle(deger)
            )
            if erh_text_eslesme:
                e1, e2, e3, e4 = erh_text_eslesme.groups()
                hedefler.extend([
                    f"ENT.ERH-E{e1}-{e2}-{e3}-{e4}",
                    f"ENT ERH E{e1} {e2} {e3} {e4}",
                ])

        # ENT.ERH-50-7-400-4 gibi tireli yazımlar için öncelikli güvenli regex.
        erh_text_eslesme = re.search(
            r"(?:ENT|ENTES)?\s*[.\-]?\s*ERH\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)",
            kod_temizle(deger)
        )
        if erh_text_eslesme:
            e1, e2, e3, e4 = erh_text_eslesme.groups()
            hedefler.extend([
                f"ENT.ERH-E{e1}-{e2}-{e3}-{e4}",
                f"ENT ERH E{e1} {e2} {e3} {e4}",
            ])

        # ENT.ERH-7-E50-400-15 -> ENT.ERH-E50-7-400-15
        # Bu yazımda E50 ikinci parçada geldiği için ilk iki parça yer değiştirir.
        erh_e_text_eslesme = re.search(
            r"(?:ENT|ENTES)?\s*[.\-]?\s*ERH\s*[- ]\s*(\d+)\s*[- ]\s*E\s*(\d+)\s*[- ]\s*(\d+)\s*[- ]\s*(\d+)",
            kod_temizle(deger),
            flags=re.IGNORECASE,
        )
        if erh_e_text_eslesme:
            e_sira, e_govde, e3, e4 = erh_e_text_eslesme.groups()
            hedefler.extend([
                f"ENT.ERH-E{e_govde}-{e_sira}-{e3}-{e4}",
                f"ENT ERH E{e_govde} {e_sira} {e3} {e4}",
            ])

        # CM / CF kapak yazımları:
        # SANTRAL: CM 8,33 KVAR KAPAK -> ENTES model: CM serisi 8,33kVAr
        # SANTRAL: CF 1,67 KVAR KAPAK -> ENTES model: CF SERİSİ 1,67kVAr
        # CF 2,5 için fiyat listesi karşılığı kullanıcı doğrulamasına göre 2,57kVAr'dır.
        cmcf_eslesme = re.search(
            r"\b(CM|CF)\s*[- ]?\s*([0-9]+(?:[,.][0-9]+)?)\s*(?:KVAR|KVA[R]?|KVAr|kVAr)\s*KAPAK\b",
            kod_temizle(deger),
            flags=re.IGNORECASE,
        )
        if cmcf_eslesme:
            seri = cmcf_eslesme.group(1).upper()
            kvar_txt = cmcf_eslesme.group(2).replace(".", ",")
            if seri == "CF" and santral_sayi_key(kvar_txt) == "2.5":
                kvar_txt = "2,57"
            hedefler.extend([
                f"{seri} serisi {kvar_txt}kVAr",
                f"{seri} SERİSİ {kvar_txt}kVAr",
                f"{seri} Serisi {kvar_txt} kVAr",
                f"{seri} SERISI {kvar_txt} kVAr",
                f"{seri} serisi {kvar_txt} KVAR",
            ])

        sc_text_eslesme = re.search(
            r"(?:ENT|ENTES)?\s*[.\-]?\s*SC\s*[- ]\s*(\d+S?)\s*[- ]\s*(\d+)$",
            kod_temizle(deger)
        )
        if sc_text_eslesme:
            sc_govde = sc_text_eslesme.group(1)
            hedefler.extend([
                f"ENT-SC-{sc_govde}",
                f"ENT.SC-{sc_govde}",
                f"ENT SC {sc_govde}",
            ])

        sc_sade_eslesme = re.fullmatch(r"(?:ENT|ENTES)?SC(\d+S?)(?:25|50)", raw_sade)
        if sc_sade_eslesme:
            sc_govde = sc_sade_eslesme.group(1)
            hedefler.extend([
                f"ENT-SC-{sc_govde}",
                f"ENT.SC-{sc_govde}",
                f"ENT SC {sc_govde}",
            ])

        # RGSR SVC yazım farkları:
        # SANTRAL: RGSR-20S (SVC)      -> ENTES model: RGSR-20S-OG
        # SANTRAL: RGSR-24S (SVC)      -> ENTES model: RGSR-24S-OG
        # SANTRAL: RGSR-20S-OG(SVC)    -> ENTES model: RGSR-20S
        # Kural:
        # - SVC var, OG yoksa hedefe -OG eklenir.
        # - SVC var, OG zaten yazıyorsa hedefte OG kaldırılır.
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        rgsr_text_eslesme = re.search(
            r"RGSR\s*[- ]?\s*(\d+)\s*S(?:\s*[- ]?\s*OG)?\s*(?:\(|\s)*SVC",
            kod_temizle(deger),
            flags=re.IGNORECASE,
        )
        if rgsr_text_eslesme:
            rgsr_no = rgsr_text_eslesme.group(1)
            rgsr_metni = kod_temizle(deger).upper()
            if "OG" in rgsr_metni:
                hedefler.append(f"RGSR-{rgsr_no}S")
            else:
                hedefler.append(f"RGSR-{rgsr_no}S-OG")

        # Eski özel yazımlar da güvenlik için korunur.
        if raw_sade in {"RGSR20SSVC", "RGSR20SVC"}:
            hedefler.append("RGSR-20S-OG")

        if raw_sade in {"RGSR15SSVC", "RGSR15SVC"}:
            hedefler.append("RGSR-15S")

        if raw_sade in {"RGSR24SSVC", "RGSR24SVC"}:
            hedefler.append("RGSR-24S-OG")

        if raw_sade in {"RGSR20SOGSVC", "RGSR20SOGS"}:
            hedefler.append("RGSR-20S")

        # NH bıçaklı sigorta yazım farkı:
        # SANTRAL: ENT.NH00 25A -> ENTES model: ENT.NH00 25A Bıçaklı Sigorta
        # SANTRAL: ENT.NH1 160A -> ENTES model: ENT.NH1 160A Bıçaklı Sigorta
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        # NH00 / NH1 ailesi ve amper değeri birebir aynı olmak zorundadır.
        nh_eslesme = re.fullmatch(r"(?:ENT|ENTES)?(NH00|NH1)(\d+)A", raw_sade)
        if nh_eslesme:
            nh_aile = nh_eslesme.group(1)
            amper = nh_eslesme.group(2)
            hedefler.extend([
                f"ENT.{nh_aile} {amper}A Bıçaklı Sigorta",
                f"{nh_aile} {amper}A Bıçaklı Sigorta",
                f"ENT.{nh_aile} {amper} A Bıçaklı Sigorta",
                f"{nh_aile} {amper} A Bıçaklı Sigorta",
            ])

        # MPR-47SE-0,5 IP54 -> MPR-47SE-IP54-0,5
        # Bu da genel class/IP54 yazım kuralıdır; M kodu elle yazılmaz.
        if raw_sade in {
            "MPR47SE05",
            "MPR47SE0D5",
            "MPR47SE05IP54",
            "MPR47SE0D5IP54",
            "MPR47SEIP5405",
            "MPR47SEIP540D5",
        }:
            hedefler.extend([
                "MPR-47SE-IP54-0,5",
                "MPR-47SE-IP54-0.5",
                "MPR-47SE-IP54 0,5",
            ])

        # MPR-47S OG-D 0,5 IP51 -> MPR-47S-OG-D-0,5
        # Bu yazımda 0,5 class bilgisi varsa hedef model IP54 değil, classlı modeldir.
        if raw_sade in {
            "MPR47SOGD05IP51",
            "MPR47SOGD0D5IP51",
            "MPR47SOGD05IP54",
            "MPR47SOGD0D5IP54",
            "MPR47SOGDIP5105",
            "MPR47SOGDIP510D5",
            "MPR47SOGD05",
            "MPR47SOGD0D5",
        }:
            hedefler.extend([
                "MPR-47S-OG-D-0,5",
                "MPR-47S-OG-D-0.5",
                "MPR-47S-OG-D 0,5",
                "MPR-47S OG-D 0,5",
            ])

        # MPR-47S OG-D IP51 -> MPR-47S-OG-D-IP54
        # IP51 gelen SANTRAL yazımı fiyat listesindeki IP54 modeline gider.
        if raw_sade in {"MPR47SOGDIP51", "MPR47SOGDIP54", "MPR47SOGD"}:
            hedefler.append("MPR-47S-OG-D-IP54")

        # MPR-47D-0,5 IP54 -> MPR-47S-D-0,5-IP54
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {
            "MPR47D05IP54",
            "MPR47D0D5IP54",
            "MPR47DIP5405",
            "MPR47DIP540D5",
            "MPR47D05",
            "MPR47D0D5",
        }:
            hedefler.extend([
                "MPR-47S-D-0,5-IP54",
                "MPR-47S-D-0.5-IP54",
                "MPR-47S-D-IP54-0,5",
                "MPR-47S-D-IP54-0.5",
            ])

        # MPR-47D IP54 -> MPR-47S-D-IP54
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {"MPR47DIP54", "MPR47D"}:
            hedefler.append("MPR-47S-D-IP54")

        # MM modül tipi kuralları:
        # SANTRAL kısa modelini ENTES fiyat listesindeki yıldızlı model adına çevirir.
        # M kodu elle yazılmaz; hedef model fiyat listesinden bulunur ve o satırdaki Ürün Kodu alınır.
        mm_ozel_hedefler = {
            "MM102": "* MM-102 Dijital Çıkış Modülü",
            "MM002": "* MM-002 Röle Çıkış Modülü",
            "MM122": "* MM-122 Dijital Giriş-Çıkış Modülü",
            "MM202": "* MM-202 Analog Çıkış Modülü",
            "MM144": "* MM-144 Dijital Giriş / Çıkış",
        }

        if raw_sade in mm_ozel_hedefler:
            hedefler.append(mm_ozel_hedefler[raw_sade])

        # Diğer MM modellerinde eski genel kural korunur:
        # Örnek: MM-120 -> * MM-120 Dijital Giriş Modülü
        # Özel tabloda olanlar tekrar genel giriş modülüne düşürülmez.
        mm_eslesme = re.fullmatch(r"MM(\d+)", raw_sade)
        if mm_eslesme and raw_sade not in mm_ozel_hedefler:
            mm_no = mm_eslesme.group(1)
            hedefler.append(f"* MM-{mm_no} Dijital Giriş Modülü")


        # PGS soket yazım farkı:
        # SANTRAL: PGS-108 -> ENTES model: PGS-108(8 PIN SOKET)
        # Buradaki PIN sabit 8 değildir; model numarasının son kısmından gelir.
        # Örnekler:
        # PGS-108 -> 8 PIN
        # PGS-111 -> 11 PIN
        # PGS-114 -> 14 PIN
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        pgs_eslesme = re.fullmatch(r"PGS(\d+)", raw_sade)
        if pgs_eslesme:
            pgs_no = pgs_eslesme.group(1)

            pin_adaylari = []

            # PGS-108 gibi modellerde 108 içindeki 08, pin sayısıdır.
            # PGS-111 -> 11, PGS-114 -> 14 şeklinde çalışır.
            if len(pgs_no) >= 3 and pgs_no.startswith("10"):
                pin_raw = pgs_no[2:]
                if pin_raw:
                    try:
                        pin_adaylari.append(str(int(pin_raw)))
                    except Exception:
                        pin_adaylari.append(pin_raw)

            # Bazı yazımlarda direkt son iki rakam pin olarak kullanılabilir.
            if len(pgs_no) >= 2:
                pin_raw = pgs_no[-2:]
                try:
                    pin_adaylari.append(str(int(pin_raw)))
                except Exception:
                    pin_adaylari.append(pin_raw)

            # Tam numarayı da düşük öncelikli aday olarak ekle; bazı fiyat listelerinde farklı yazım olabilir.
            pin_adaylari.append(pgs_no)

            gorulen_pin = set()
            for pin_sayisi in pin_adaylari:
                pin_sayisi = str(pin_sayisi).strip()
                if not pin_sayisi or pin_sayisi in gorulen_pin:
                    continue

                gorulen_pin.add(pin_sayisi)

                hedefler.extend([
                    f"PGS-{pgs_no}({pin_sayisi} PIN SOKET)",
                    f"PGS-{pgs_no} ({pin_sayisi} PIN SOKET)",
                    f"PGS {pgs_no} {pin_sayisi} PIN SOKET",
                ])

        # EMK gerilim besleme yazım farkı:
        # SANTRAL: EMK-01 -> ENTES model: EMK-01 (85-265 VAC/DC)
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {"EMK01", "ENTEMK01", "ENTESEMK01"}:
            hedefler.extend([
                "EMK-01 (85-265 VAC/DC)",
                "EMK-01 85-265 VAC/DC",
                "EMK-01",
            ])

        # MPR-53-96 ölçü / CT25 yazım farkı:
        # SANTRAL: MPR-53-96 -> ENTES model: MPR-53-96-CT25
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {"MPR5396", "ENTMPR5396", "ENTESMPR5396"}:
            hedefler.extend([
                "MPR-53-96-CT25",
                "MPR53-96CT25",
                "MPR-53 96 CT25",
            ])

        # ETMO ethernet gateway yazım farkı:
        # SANTRAL: ETMO-02 -> ENTES model: ETMO-02 Ethernet Gateway
        # SANTRAL: ETMO-10 -> ENTES model: ETMO-10 Ethernet Gateway
        # Kural: ETMO-NO -> ETMO-NO Ethernet Gateway
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        etmo_eslesme = re.fullmatch(r"(?:ENT|ENTES)?ETMO(\d+)", raw_sade)
        if etmo_eslesme:
            etmo_no = etmo_eslesme.group(1)
            hedefler.extend([
                f"ETMO-{etmo_no} Ethernet Gateway",
                f"ETMO {etmo_no} Ethernet Gateway",
                f"ETMO-{etmo_no}",
            ])

        # GEMO 4G gateway yazım farkı:
        # SANTRAL: GEMO-10-DR -> ENTES model: GEMO-10-DR-4G
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        if raw_sade in {"GEMO10DR", "ENTGEMO10DR", "ENTESGEMO10DR"}:
            hedefler.extend([
                "GEMO-10-DR-4G",
                "GEMO 10 DR 4G",
                "GEMO-10-DR",
            ])

        # SA60 class yazım farkı:
        # SANTRAL: SA60-30 -> ENTES model: SA60-30-0,5
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        sa60_eslesme = re.fullmatch(r"SA60(\d+)", raw_sade)
        if sa60_eslesme:
            sa60_no = sa60_eslesme.group(1)
            hedefler.extend([
                f"SA60-{sa60_no}-0,5",
                f"SA60-{sa60_no}-0.5",
            ])

        # CKR besleme gerilimi yazım farkı:
        # SANTRAL: CKR-93T -> ENTES model: CKR-93T-85-265- VAC DC
        # M kodu elle yazılmaz; fiyat listesindeki model satırından ürün kodu alınır.
        ckr_eslesme = re.fullmatch(r"CKR(\d+T)", raw_sade)
        if ckr_eslesme:
            ckr_no = ckr_eslesme.group(1)
            hedefler.extend([
                f"CKR-{ckr_no}-85-265- VAC DC",
                f"CKR-{ckr_no}-85-265 VAC DC",
                f"CKR-{ckr_no} 85-265 VAC/DC",
                f"CKR-{ckr_no}-85-265-VAC-DC",
            ])

        # MKC bobin gerilimi yazım farkı:
        # SANTRAL: MKC-04 -> ENTES model: MKC-04-220/230VAC
        # Benzer MKC modellerinde numara korunur, 220/230VAC hedeflenir.
        mkc_eslesme = re.fullmatch(r"MKC(\d+)", raw_sade)
        if mkc_eslesme:
            mkc_no = mkc_eslesme.group(1)
            hedefler.extend([
                f"MKC-{mkc_no}-220/230VAC",
                f"MKC-{mkc_no} 220/230VAC",
                f"MKC-{mkc_no}-220-230VAC",
                f"MKC-{mkc_no}-220/230 VAC",
            ])

        # EF besleme gerilimi yazım farkı:
        # SANTRAL: EF-10 -> ENTES model: EF-10-220V AC
        # Benzer EF modellerinde numara korunur, 220V AC hedeflenir.
        ef_eslesme = re.fullmatch(r"EF(\d+)", raw_sade)
        if ef_eslesme:
            ef_no = ef_eslesme.group(1)
            hedefler.extend([
                f"EF-{ef_no}-220V AC",
                f"EF-{ef_no} 220V AC",
                f"EF-{ef_no}-220VAC",
            ])

        # ERB buzzer yazım farkı:
        # SANTRAL: ERB-50-BUZZER -> ENTES model: ERB-50 (BUZZER)Buzerli
        # Benzer ERB buzzer modellerinde numara korunur.
        erb_buzzer_eslesme = re.fullmatch(r"ERB(\d+)BUZZER", raw_sade)
        if erb_buzzer_eslesme:
            erb_no = erb_buzzer_eslesme.group(1)
            hedefler.extend([
                f"ERB-{erb_no} (BUZZER)Buzerli",
                f"ERB-{erb_no}(BUZZER)Buzerli",
                f"ERB-{erb_no} BUZZER Buzerli",
                f"ERB-{erb_no} Buzerli",
            ])

        # Özel kablo bağı kesin hedef kuralları:
        # Bu satırlarda M kodu ezberlenmez. SANTRAL yazımı doğrudan ENTES fiyat listesindeki
        # ÜRÜN MODELİ metnine çevrilir; sonra o model satırının ÜRÜN KODU kullanılır.
        # Bu grup ölçü mantığıyla güvenli çözülemediği için kesin hedef model olarak tanımlandı.
        ozel_kablo_kesin_hedefler = {
            "76X200ACILANKBAG": [
                "TS1-8-75200N 7,6x200mm Kablo Bağı, Açılabilir, Beyaz",
                "TS1-8-75200N 7.6x200mm Kablo Bagi, Acilabilir, Beyaz",
            ],
            "76X300ACILANKBAG": [
                "TS1-8-75300N 7,6x300mm Kablo Bağı, Açılabilir, Beyaz",
                "TS1-8-75300N 7.6x300mm Kablo Bagi, Acilabilir, Beyaz",
            ],
            "25X100ETIKETKBAG": [
                "TS1-9-25100N 2,5x100mm Kablo Bağı, Etiketli, Beyaz",
                "TS1-9-25100N 2.5x100mm Kablo Bagi, Etiketli, Beyaz",
            ],
            "48X190ETIKETKBAG": [
                "TS1-9-46190N 4,8x190mm Kablo Bağı, Etiketli, Beyaz",
                "TS1-9-46190N 4.8x190mm Kablo Bagi, Etiketli, Beyaz",
            ],
            "25X100BONCUKKBAG": [
                "TS1-13-25100N 2,5x100mm Kablo Bağı, Boncuk Tip, Beyaz",
                "TS1-13-25100N 2.5x100mm Kablo Bagi, Boncuk Tip, Beyaz",
            ],
            "25X150BONCUKKBAG": [
                "TS1-13-25150N 2,5x150mm Kablo Bağı, Boncuk Tip, Beyaz",
                "TS1-13-25150N 2.5x150mm Kablo Bagi, Boncuk Tip, Beyaz",
            ],
            "25X100DUBELLIKBAG": [
                "TS2-5-02N 22X15mm Kablo Bağı Dübelli,Beyaz",
                "TS2-5-02N 22X15mm Kablo Bagi Dubelli, Beyaz",
            ],
            "48X190DUBELLIKBAG": [
                "TS2-5-02B 22X15mm Kablo Bağı Dübelli,Siyah",
                "TS2-5-02B 22X15mm Kablo Bagi Dubelli, Siyah",
            ],
        }

        if raw_sade in ozel_kablo_kesin_hedefler:
            hedefler.extend(ozel_kablo_kesin_hedefler[raw_sade])


        # Yapışkan kroşe kuralları:
        # SANTRAL: 19x19 YAPIŞKAN KROŞE
        # ENTES fiyat listesi adayları:
        # TS2-1-2020N 19x19mm Kroşe, Yapışkanlı, Beyaz
        # TS2-1-2020B 19x19mm Kroşe, Yapışkanlı, Siyah
        # Not: 19x19 SANTRAL yazımında renk yoksa iki aday da üretilir;
        # fiyat listesinde ikisi de varsa güvenlik için çoklu eşleşme olarak kalır.
        # Renk yazıyorsa BEYAZ/SİYAH ayrımı yapılır.
        krose_eslesme = re.search(r"(\d+)\s*[xX]\s*(\d+).*YAPISKAN.*KROSE", raw_sade)
        if krose_eslesme:
            en = krose_eslesme.group(1)
            boy = krose_eslesme.group(2)

            renk_adaylari = []
            if "SIYAH" in raw_sade or raw_sade.endswith("S"):
                renk_adaylari.append(("B", "Siyah"))
            elif "BEYAZ" in raw_sade or raw_sade.endswith("B"):
                renk_adaylari.append(("N", "Beyaz"))
            else:
                renk_adaylari.extend([("N", "Beyaz"), ("B", "Siyah")])

            # Fiyat listesinde 19x19 kodu TS2-1-2020, 28x28 kodu TS2-1-3030 şeklinde geçiyor.
            # Önce ölçü metniyle güvenli arama yapılır; TS kodu da hedefe eklenir.
            ts_olcu_kod = {
                ("19", "19"): "2020",
                ("28", "28"): "3030",
            }.get((en, boy), f"{int(en):02d}{int(boy):02d}")

            for suffix, renk_adi in renk_adaylari:
                hedefler.extend([
                    f"TS2-1-{ts_olcu_kod}{suffix} {en}x{boy}mm Kroşe, Yapışkanlı, {renk_adi}",
                    f"TS2-1-{ts_olcu_kod}{suffix} {en}X{boy}mm Kroşe, Yapışkanlı, {renk_adi}",
                    f"{en}x{boy}mm Kroşe, Yapışkanlı, {renk_adi}",
                    f"{en}X{boy}mm Kroşe, Yapışkanlı, {renk_adi}",
                ])

        # Kablo bağı ölçü/renk yazım farkı:
        # SANTRAL: 2,5x80 KABLO BAĞI B
        # Fiyat listesi: TS1-1-25080N 2,5x80mm Kablo Bağı, Beyaz
        # Bu blok yalnızca hedef model adayları üretir. Asıl güvenli eşleşme
        # santral_kablo_bagi_fiyat_listesinden_bul() fonksiyonunda ölçü + renk ile yapılır.
        parca_kablo = santral_kablo_bagi_parcala(deger)
        if parca_kablo:
            renk_to_suffix_ad = {
                "BEYAZ": ("N", "Beyaz"),
                "SIYAH": ("B", "Siyah"),
                "SARI": ("Y", "Sarı"),
                "KIRMIZI": ("R", "Kırmızı"),
                "MAVI": ("BL", "Mavi"),
                "MOR": ("P", "Mor"),
                "YESIL": ("G", "Yeşil"),
            }

            renk_bilgisi = renk_to_suffix_ad.get(parca_kablo["renk"])

            if renk_bilgisi:
                suffix, renk_ad = renk_bilgisi
                genislik_raw = parca_kablo["genislik"]
                boy_raw = parca_kablo["boy"]
                genislik_kod = re.sub(r"[^0-9]", "", genislik_raw)
                boy_kod = str(int(boy_raw)).zfill(3)
                uv_ek = "UV" if parca_kablo.get("uv") else ""
                ts_kod = f"TS1-1-{genislik_kod}{boy_kod}{uv_ek}{suffix}"

                for genislik_yazim in [parca_kablo["genislik"], parca_kablo["genislik_nokta"]]:
                    if parca_kablo.get("uv"):
                        hedefler.extend([
                            f"{ts_kod} {genislik_yazim}x{int(boy_raw)}mm Kablo Bağı, UV'ye Dayanıklı, {renk_ad}",
                            f"{ts_kod} {genislik_yazim} x {int(boy_raw)} mm Kablo Bağı, UV'ye Dayanıklı, {renk_ad}",
                            f"{ts_kod} {genislik_yazim}x{int(boy_raw)}mm Kablo Bagi, UV'ye Dayanikli, {renk_ad}",
                        ])
                    else:
                        hedefler.extend([
                            f"{ts_kod} {genislik_yazim}x{int(boy_raw)}mm Kablo Bağı, {renk_ad}",
                            f"{ts_kod} {genislik_yazim} x {int(boy_raw)} mm Kablo Bağı, {renk_ad}",
                            f"{ts_kod} {genislik_yazim}x{int(boy_raw)}mm Kablo Bagi, {renk_ad}",
                        ])

                hedefler.append(ts_kod)

        return hedefler

    def santral_hedef_modelden_kod_bul(hedefler):
        """
        Verilen hedef ENTES model yazımlarını fiyat listesindeki Ürün Modeli kolonunda arar.
        Bu fonksiyon M kodu ezberlemez; hedef model bulunursa o satırın Ürün Kodu alınır.

        Önemli fark: artık model metnini parçalamaya güvenmiyoruz.
        Önce SANTRAL kısa yazımından hedef ENTES model yazımı üretiliyor,
        sonra fiyat listesinde sadeleştirilmiş model metniyle aranıyor.
        """
        if not hedefler:
            return "", ""

        hedef_sadeler = []
        for hedef in hedefler:
            hedef_sade = sade_kod(hedef)
            if hedef_sade and hedef_sade not in hedef_sadeler:
                hedef_sadeler.append(hedef_sade)

        if not hedef_sadeler:
            return "", ""

        def tekille(bulunanlar):
            koddan_model = {}
            for kod, model in bulunanlar:
                if kod:
                    koddan_model.setdefault(kod, model)

            if len(koddan_model) == 1:
                yeni_kod = list(koddan_model.keys())[0]
                return yeni_kod, koddan_model[yeni_kod]

            return "", ""

        # 1) Tam sade eşleşme.
        bulunanlar = []
        for _, fiyat_row in entes_df.iterrows():
            model_raw = str(fiyat_row.get(entes_model_col, "")).strip()
            kod_raw = str(fiyat_row.get(entes_kod_col, "")).strip()
            if not model_raw or not kod_raw:
                continue

            model_sade = sade_kod(model_raw)
            if model_sade in hedef_sadeler:
                bulunanlar.append((kod_raw, model_raw))

        yeni_kod, bulunan_model = tekille(bulunanlar)
        if yeni_kod:
            return yeni_kod, bulunan_model

        # 2) Fiyat listesinde model sonunda ek açıklama varsa hedef model modelin içinde geçebilir.
        # Örn: hedef "ENT.A20 MN 400-5VA CL:0,5", fiyat listesi "ENT.A20 MN 400-5VA CL:0,5 A.G. Akım Trafosu"
        bulunanlar = []
        for _, fiyat_row in entes_df.iterrows():
            model_raw = str(fiyat_row.get(entes_model_col, "")).strip()
            kod_raw = str(fiyat_row.get(entes_kod_col, "")).strip()
            if not model_raw or not kod_raw:
                continue

            model_sade = sade_kod(model_raw)
            if any(hedef_sade and hedef_sade in model_sade for hedef_sade in hedef_sadeler):
                bulunanlar.append((kod_raw, model_raw))

        return tekille(bulunanlar)

    def santral_sayi_key(val):
        txt = metin_norm(val)
        txt = txt.replace("VA", "").replace("V", "")
        txt = txt.replace("CL", "").replace("KL", "")
        txt = txt.replace(":", "").replace(";", "")
        txt = txt.replace(" ", "").replace(",", ".")
        s_eki = txt.endswith("S")
        txt = txt.rstrip("S")

        try:
            f = float(txt)
        except Exception:
            return ""

        ana = str(int(f)) if abs(f - int(f)) < 0.0001 else f"{f:.6f}".rstrip("0").rstrip(".")
        return ana + ("S" if s_eki else "")

    def santral_sayi_tr(val):
        key = santral_sayi_key(val)
        if not key:
            return ""

        s_eki = key.endswith("S")
        ana = key[:-1] if s_eki else key
        ana = ana.replace(".", ",")
        return ana + ("S" if s_eki else "")

    def santral_ag_hedef_yazimlari(aile, primer, va, cl, slash=False):
        aile = int(aile)
        primer = int(primer)
        va_txt = santral_sayi_tr(va)
        cl_txt = santral_sayi_tr(cl)

        if not va_txt or not cl_txt:
            return []

        hedefler = []

        if slash:
            hedefler.extend([
                f"ENT.A{aile}MN {primer}/5 {va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile} MN {primer}/5 {va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile}MN {primer} / 5 {va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile} MN {primer} / 5 {va_txt}VA CL:{cl_txt}",
            ])
        else:
            hedefler.extend([
                f"ENT.A{aile} MN {primer}-{va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile} MN {primer}- {va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile} MN {primer} {va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile}MN {primer}-{va_txt}VA CL:{cl_txt}",
                f"ENT.A{aile}MN {primer} {va_txt}VA CL:{cl_txt}",
            ])

        return hedefler

    def santral_ag_akim_trafo_hedefleri(deger):
        """
        SANTRAL akım trafosu kısa yazımından hedef ENTES model yazımlarını üretir.

        Örnek:
        400/5 5VA KL-0,5 Mini -> ENT.A20 MN 400-5VA CL:0,5
        500/5 5VA KL-0,25 Mini -> ENT.A30MN 500/5 5VA CL:0,2S
        """
        raw = metin_norm(deger)
        raw = raw.replace("CLASS", "KL")
        raw = raw.replace("KLAS", "KL")

        es = re.search(
            r"\b(\d+)\s*/\s*5\s*(\d+(?:[,.]\d+)?)\s*VA\s*(?:KL|CL)\s*[-:]?\s*([0-9]+(?:[,.][0-9]+)?)",
            raw
        )
        if not es:
            return []

        primer = int(es.group(1))
        va_key = santral_sayi_key(es.group(2))
        kl_key = santral_sayi_key(es.group(3))
        mini_var = "MINI" in raw
        ab_b_serisi = bool(re.search(r"\bB\b", raw))

        hedefler = []

        # AB B tipi A.G. akım trafosu genel kuralı:
        # SANTRAL: 15/5 2,5VA Klas-0,2 B
        # ENTES:   ENT.AB 15/5 2,5VA Cl:0,2 A.G. Akım Trafosu
        # Burada primer / VA / class değişken; M kodu ezberlenmez.
        if ab_b_serisi and not mini_var:
            va_txt_ab = santral_sayi_tr(va_key)
            cl_txt_ab = santral_sayi_tr(kl_key)
            if va_txt_ab and cl_txt_ab:
                hedefler.extend([
                    # Fiyat listesinde AB ailesi iki farklı yazımla gelebiliyor:
                    # 1) ENT.AB 15/5 2,5VA Cl:0,2 A.G. Akım Trafosu
                    # 2) ENT.AB 15- 10 VA CL:0,5
                    # İkisini de hedefe koyuyoruz; fiyat listesinde hangisi varsa o satırın Ürün Kodu alınır.
                    f"ENT.AB {primer}/5 {va_txt_ab}VA Cl:{cl_txt_ab} A.G. Akım Trafosu",
                    f"ENT.AB {primer}/5 {va_txt_ab}VA CL:{cl_txt_ab} A.G. Akım Trafosu",
                    f"ENT.AB {primer}/5 {va_txt_ab}VA Cl:{cl_txt_ab} AG Akım Trafosu",
                    f"ENT.AB {primer}/5 {va_txt_ab}VA CL:{cl_txt_ab} AG Akim Trafosu",
                    f"ENT.AB {primer}/5 {va_txt_ab}VA Cl:{cl_txt_ab}",
                    f"ENT.AB {primer}/5 {va_txt_ab}VA CL:{cl_txt_ab}",
                    f"ENT.AB {primer}- {va_txt_ab} VA CL:{cl_txt_ab}",
                    f"ENT.AB {primer}-{va_txt_ab} VA CL:{cl_txt_ab}",
                    f"ENT.AB {primer}- {va_txt_ab}VA CL:{cl_txt_ab}",
                    f"ENT.AB {primer}-{va_txt_ab}VA CL:{cl_txt_ab}",
                ])

        # Kullanıcının elle doğruladığı SANTRAL -> ENTES model karşılıkları.
        # Bunlar kesin hedef model kuralıdır; M kodu ezberlenmez.
        # Kod, aşağıdaki hedef modeli ENTES fiyat listesinde bulur ve o satırdaki Ürün Kodu'nu alır.
        kesin_ag_hedefleri = {
            # Mini olmayanlar
            (40, "1.5", "3", False): ["ENT.A20 MN 40-1VA CL:3"],
            (50, "1.5", "3", False): ["ENT.A20 MN 50-1,25VA CL:1"],

            # Mini olan özel A20/A30/A40 MN dönüşümleri
            (40, "1.5", "3", True): ["ENT.A30 MN 40- 1VA CL:3"],
            (50, "1.5", "1", True): ["ENT.A30 MN 50- 1VA CL:1"],
            (80, "1.5", "1", True): ["ENT.A30 MN 80- 2,5VA CL:1"],
            (100, "1.5", "1", True): ["ENT.A30MN 100/5 1VA CL:0,2S"],
            (125, "1.5", "1", True): ["ENT.A30 MN 125- 1,5VA CL:0,5"],
            (250, "2.5", "0.5", True): ["ENT.A30 MN 250- 2,5VA CL:0,5"],

            (300, "5", "0.25", True): ["ENT.A30MN 300/5 2,5VA CL:0,2S"],
            (400, "5", "0.25", True): ["ENT.A30MN 400/5 2,5VA CL:0,2S"],
            (400, "5", "0.5", True): ["ENT.A30 MN 400- 5VA CL:0,5"],
            (500, "5", "0.25", True): ["ENT.A40MN 500/5 2,5VA CL:0,2S"],
            (500, "5", "0.5", True): ["ENT.A30 MN 500- 5VA CL:0,5"],
            (600, "5", "0.25", True): ["ENT.A40MN 600/5 2,5VA CL:0,2S"],
            (600, "5", "0.5", True): ["ENT.A30 MN 600- 5VA CL:0,5"],
            (700, "5", "0.25", True): ["ENT.A40MN 750/5 2,5VA CL:0,2S"],
            (800, "5", "0.25", True): ["ENT.A40MN 800/5 2,5VA CL:0,2S"],
        }

        kesin_key = (primer, va_key, kl_key, mini_var)
        if kesin_key in kesin_ag_hedefleri:
            hedefler.extend(kesin_ag_hedefleri[kesin_key])

        def ekle(aile, hedef_va, hedef_cl, slash=False):
            hedefler.extend(santral_ag_hedef_yazimlari(aile, primer, hedef_va, hedef_cl, slash=slash))

        if hedefler:
            pass
        elif mini_var:
            # Kullanıcının verdiği örneklerden çıkarılan genel kurallar.
            # Burada primer değişken; M kodu yok.
            if va_key == "1.5" and kl_key == "1" and primer == 60:
                ekle(20, 1.25, 1)
            elif va_key == "1.5" and kl_key == "1" and primer == 75:
                ekle(20, 1.5, 1)
            elif va_key == "1.5" and kl_key == "0.5" and primer == 100:
                ekle(30, 1.5, "0,5")
            elif va_key == "1.5" and kl_key == "0.5" and primer >= 150:
                ekle(30, 1.5, "0,2S", slash=True)
            elif va_key == "1.5" and kl_key in {"0.25", "0.2", "0.2S"}:
                ekle(30, 1.5, "0,2S", slash=True)
            elif va_key == "2.5" and kl_key == "1":
                ekle(30, 2.5, 1)
            elif va_key == "2.5" and kl_key == "0.5":
                ekle(20, 2.5, "0,5")
            elif va_key == "5" and kl_key == "1":
                ekle(20, 5, 1)
            elif va_key == "5" and kl_key == "0.5":
                ekle(20, 5, "0,5")
            elif va_key == "5" and kl_key in {"0.25", "0.2", "0.2S"}:
                ekle(30, 5, "0,2S", slash=True)
            elif va_key == "7.5" and kl_key == "1":
                ekle(20, 7.5, 1)
        else:
            # Mini olmayan özel kurallar.
            if primer == 100 and va_key == "2.5" and kl_key == "0.5":
                ekle(20, 1.5, "0,5")
            elif primer == 100 and va_key == "2.5" and kl_key == "1":
                ekle(20, 2.5, 1)
            elif primer == 125 and va_key == "2" and kl_key == "0.5":
                ekle(20, 2, "0,5")
            elif primer == 125 and va_key == "4" and kl_key == "1":
                ekle(20, 4, 1)

        # Tekilleştir.
        sonuc = []
        gorulen = set()
        for h in hedefler:
            k = sade_kod(h)
            if k and k not in gorulen:
                gorulen.add(k)
                sonuc.append(h)
        return sonuc

    def santral_ag_akim_trafo_fiyat_listesinden_bul(deger):
        hedefler = santral_ag_akim_trafo_hedefleri(deger)
        return santral_hedef_modelden_kod_bul(hedefler)

    def santral_srs3_parcala(deger):
        """
        ENT.SRS3 400/1, ENT.SRS3 400/1,5, ENT.SRS3 400/25 gibi
        SANTRAL/fiyat listesi SRS3 yazımlarından gerilim ve kVar değerini ayırır.
        """
        raw = metin_norm(deger)
        if not raw:
            return None

        es = re.search(
            r"(?:ENT|ENTES)?\s*\.?\s*SRS\s*3\s*(\d+)\s*/\s*([0-9]+(?:[,.][0-9]+)?)",
            raw
        )
        if not es:
            return None

        gerilim = int(es.group(1))
        kvar_raw = es.group(2).replace(",", ".")
        try:
            kvar_float = float(kvar_raw)
        except Exception:
            return None

        # 1.0 -> 1, 2.5 -> 2,5 olacak şekilde hem sayı anahtarı hem TR yazım üret.
        if kvar_float.is_integer():
            kvar_key = str(int(kvar_float))
            kvar_txt = str(int(kvar_float))
        else:
            kvar_key = str(kvar_float).rstrip("0").rstrip(".")
            kvar_txt = kvar_key.replace(".", ",")

        return {
            "gerilim": gerilim,
            "kvar_key": kvar_key,
            "kvar_txt": kvar_txt,
        }

    def santral_srs3_hedefleri(deger):
        """
        SANTRAL SRS3 kısa yazımından genel hedef ENTES model yazımlarını üretir.
        Kural:
            ENT.SRS3 400/X -> ENT.SRS3 400/X 400V X kVar

        X değeri değişebilir: 1 / 1,5 / 2 / 2,5 / 3 / 5 / 7,5 / 10 / 12,5 / 15 / 20 / 25 / 30 / 40 / 50...
        """
        parca = santral_srs3_parcala(deger)
        if not parca:
            return []

        gerilim = parca["gerilim"]
        kvar_txt = parca["kvar_txt"]

        # Fiyat listesinde kVar bazen bitişik, bazen ayrı yazılabiliyor:
        # 2,5kVar / 2,5 kVar gibi. İkisini de hedefe koyuyoruz.
        return [
            f"ENT.SRS3 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt}kVar",
            f"ENT.SRS3 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt} kVar",
            f"ENT.SRS3 {gerilim}/{kvar_txt} {gerilim} V {kvar_txt}kVar",
            f"ENT.SRS3 {gerilim}/{kvar_txt} {gerilim} V {kvar_txt} kVar",
            f"SRS3 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt}kVar",
            f"SRS3 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt} kVar",
            f"ENT.SRS3 {gerilim}/{kvar_txt}",
        ]

    def santral_srs3_fiyat_listesinden_bul(deger):
        """
        SRS3 için önce gerçek parçalı kontrol yapılır:
        SANTRAL'de ve fiyat listesinde aynı 400/X değeri varsa eşleşir.
        Böylece 2,5 kVar / 2,5kVar / çift boşluk gibi yazım farkları sorun olmaz.
        """
        parca = santral_srs3_parcala(deger)
        if not parca:
            return "", ""

        hedef_gerilim = parca["gerilim"]
        hedef_kvar = parca["kvar_key"]

        bulunanlar = []
        for _, fiyat_row in entes_df.iterrows():
            model_raw = str(fiyat_row.get(entes_model_col, "")).strip()
            kod_raw = str(fiyat_row.get(entes_kod_col, "")).strip()
            if not model_raw or not kod_raw:
                continue

            model_norm = metin_norm(model_raw)
            if "SRS" not in model_norm or "3" not in model_norm:
                continue

            model_parca = santral_srs3_parcala(model_raw)
            if not model_parca:
                continue

            if model_parca["gerilim"] == hedef_gerilim and model_parca["kvar_key"] == hedef_kvar:
                bulunanlar.append((kod_raw, model_raw))

        koddan_model = {}
        for kod, model in bulunanlar:
            koddan_model.setdefault(kod, model)

        if len(koddan_model) == 1:
            yeni_kod = list(koddan_model.keys())[0]
            return yeni_kod, koddan_model[yeni_kod]

        # Fallback: hedef yazımlarıyla sade metin araması.
        hedefler = santral_srs3_hedefleri(deger)
        return santral_hedef_modelden_kod_bul(hedefler)

    def santral_srs1_parcala(deger):
        """
        ENT.SRS1.230/2, ENT.SRS1 230/7,5 gibi SANTRAL/fiyat listesi
        SRS1 yazımlarından gerilim ve kVAr değerini ayırır.
        Türkçe/İngilizce açıklama farkını önemsemez; anahtar 230/X değeridir.
        """
        raw = metin_norm(deger)
        if not raw:
            return None

        es = re.search(
            r"(?:ENT|ENTES)?\s*\.?\s*SRS\s*1[\s.\-]*(\d+)\s*/\s*([0-9]+(?:[,.][0-9]+)?)",
            raw
        )
        if not es:
            return None

        gerilim = int(es.group(1))
        kvar_raw = es.group(2).replace(",", ".")
        try:
            kvar_float = float(kvar_raw)
        except Exception:
            return None

        if kvar_float.is_integer():
            kvar_key = str(int(kvar_float))
            kvar_txt = str(int(kvar_float))
        else:
            kvar_key = str(kvar_float).rstrip("0").rstrip(".")
            kvar_txt = kvar_key.replace(".", ",")

        return {
            "gerilim": gerilim,
            "kvar_key": kvar_key,
            "kvar_txt": kvar_txt,
        }

    def santral_srs1_hedefleri(deger):
        """
        SANTRAL SRS1 kısa yazımından genel hedef ENTES model yazımlarını üretir.
        Kural:
            ENT.SRS1 230/X -> ENT.SRS1 230/X 230V X kVAr Shunt Reaktor / Şönt Reaktör

        X değeri değişebilir: 2 / 3 / 5 / 7,5 vb.
        """
        parca = santral_srs1_parcala(deger)
        if not parca:
            return []

        gerilim = parca["gerilim"]
        kvar_txt = parca["kvar_txt"]

        hedefler = []
        aciklama_adaylari = [
            "Shunt Reaktor",
            "Shunt Reactor",
            "Şönt Reaktör",
            "Sönt Reaktör",
            "Sont Reaktor",
            "Reaktör",
            "Reaktor",
        ]

        for aciklama in aciklama_adaylari:
            hedefler.extend([
                f"ENT.SRS1 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt}kVAr {aciklama}",
                f"ENT.SRS1 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt} kVAr {aciklama}",
                f"ENT.SRS1 {gerilim}/{kvar_txt} {gerilim} V {kvar_txt}kVAr {aciklama}",
                f"ENT.SRS1 {gerilim}/{kvar_txt} {gerilim} V {kvar_txt} kVAr {aciklama}",
            ])

        hedefler.extend([
            f"ENT.SRS1 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt}kVAr",
            f"ENT.SRS1 {gerilim}/{kvar_txt} {gerilim}V {kvar_txt} kVAr",
            f"ENT.SRS1 {gerilim}/{kvar_txt}",
        ])

        return hedefler

    def santral_srs1_fiyat_listesinden_bul(deger):
        """
        SRS1 için önce gerçek parçalı kontrol yapılır.
        Böylece fiyat listesindeki açıklama İngilizce olsa da Türkçe olsa da
        230/X değeri aynıysa doğru model satırı bulunur.
        """
        parca = santral_srs1_parcala(deger)
        if not parca:
            return "", ""

        hedef_gerilim = parca["gerilim"]
        hedef_kvar = parca["kvar_key"]

        bulunanlar = []
        for _, fiyat_row in entes_df.iterrows():
            model_raw = str(fiyat_row.get(entes_model_col, "")).strip()
            kod_raw = str(fiyat_row.get(entes_kod_col, "")).strip()
            if not model_raw or not kod_raw:
                continue

            model_norm = metin_norm(model_raw)
            if "SRS" not in model_norm or "1" not in model_norm:
                continue

            model_parca = santral_srs1_parcala(model_raw)
            if not model_parca:
                continue

            if model_parca["gerilim"] == hedef_gerilim and model_parca["kvar_key"] == hedef_kvar:
                bulunanlar.append((kod_raw, model_raw))

        koddan_model = {}
        for kod, model in bulunanlar:
            koddan_model.setdefault(kod, model)

        if len(koddan_model) == 1:
            yeni_kod = list(koddan_model.keys())[0]
            return yeni_kod, koddan_model[yeni_kod]

        # Fallback: Türkçe/İngilizce hedef yazımlarıyla sade metin araması.
        hedefler = santral_srs1_hedefleri(deger)
        return santral_hedef_modelden_kod_bul(hedefler)

    def santral_kural_modelden_kod_bul(deger):
        # SRS1 kısa yazımlarını önce parçalı kontrolle fiyat listesindeki ENT.SRS1 modelinden çöz.
        # Örnek: ENT.SRS1.230/7,5 -> ENT.SRS1 230/7,5 230V 7,5 kVAr Shunt Reaktor / Şönt Reaktör
        srs1_kod, srs1_model = santral_srs1_fiyat_listesinden_bul(deger)
        if srs1_kod:
            return srs1_kod, srs1_model

        # SRS3 kısa yazımlarını önce parçalı kontrolle fiyat listesindeki ENT.SRS3 modelinden çöz.
        # Örnek: ENT.SRS3 400/1,5 -> ENT.SRS3 400/1,5 400V 1,5 kVar
        srs3_kod, srs3_model = santral_srs3_fiyat_listesinden_bul(deger)
        if srs3_kod:
            return srs3_kod, srs3_model

        # Alçak gerilim akım trafosu kısa yazımlarını önce fiyat listesindeki ENT.Axx MN modelinden çöz.
        # Örnek: 50/5 1,5VA KL-3 -> ENT.A20 MN 50-1,25VA CL:1
        ag_akim_kod, ag_akim_model = santral_ag_akim_trafo_fiyat_listesinden_bul(deger)
        if ag_akim_kod:
            return ag_akim_kod, ag_akim_model

        # Kablo bağlarında önce ölçü + renk üzerinden fiyat listesindeki model satırını ara.
        # M kodları burada dikkate alınmaz; fiyat listesindeki Ürün Modeli belirleyicidir.
        kablo_bagi_kod, kablo_bagi_model = santral_kablo_bagi_fiyat_listesinden_bul(deger)
        if kablo_bagi_kod:
            return kablo_bagi_kod, kablo_bagi_model

        hedefler = santral_kural_model_hedefleri(deger)

        if not hedefler:
            return "", ""

        # Önce hedef modelin birebir temizlenmiş halini dene.
        # Özellikle "* MM-120 Dijital Giriş Modülü" gibi yıldızın ayırt edici olduğu
        # modellerde gevşek/sade anahtarlara geçmeden doğru satırı yakalamak için.
        for hedef in hedefler:
            birebir_anahtar = kod_temizle(hedef)
            birebir_adaylar = model_lookup.get(birebir_anahtar, set())
            birebir_koddan_model = {}
            for kod, model in birebir_adaylar:
                birebir_koddan_model.setdefault(kod, model)

            if len(birebir_koddan_model) == 1:
                yeni_kod = list(birebir_koddan_model.keys())[0]
                bulunan_model = birebir_koddan_model[yeni_kod]
                return yeni_kod, bulunan_model

        adaylar = set()

        for hedef in hedefler:
            for anahtar in santral_anahtarlar_uret(hedef):
                adaylar.update(model_lookup.get(anahtar, set()))

        koddan_model = {}
        for kod, model in adaylar:
            koddan_model.setdefault(kod, model)

        if len(koddan_model) == 1:
            yeni_kod = list(koddan_model.keys())[0]
            bulunan_model = koddan_model[yeni_kod]
            return yeni_kod, bulunan_model

        return "", ""

    rapor_satirlari = []

    say_zaten_urun_kodu = 0
    say_modelden_cevrilen = 0
    say_model_bulunamayan = 0
    say_coklu = 0
    say_bos = 0
    kontrol_edilen_sayi = int(kontrol_mask.sum())

    for idx, row in df.iterrows():
        if not bool(kontrol_mask.loc[idx]):
            df.at[idx, durum_kolon] = "Kontrol dışı"
            df.at[idx, not_kolon] = "Bu satır SANTRAL marka filtresine girmediği için ENTES düzeltme sayımına dahil edilmedi."
            continue

        eski_deger = str(row.get(kod_kolonu, "")).strip()
        df.at[idx, orijinal_kolon] = eski_deger

        if not eski_deger or eski_deger.lower() == "nan":
            df.at[idx, durum_kolon] = "Boş ürün kodu"
            df.at[idx, not_kolon] = "SANTRAL ürün kodu/model alanı boş."
            say_bos += 1
            rapor_satirlari.append({
                "B2B Satır No": idx + 2,
                "Eski Değer": eski_deger,
                "Yeni Ürün Kodu": "",
                "Durum": "Boş ürün kodu",
                "Not": "SANTRAL ürün kodu/model alanı boş.",
                "Kaynak Dosya": str(kaynak_dosya),
            })
            continue

        anahtarlar = santral_anahtarlar_uret(eski_deger)

        # 1) Değer ENTES fiyat listesinde ürün kodu olarak geçiyorsa bunu kullan.
        # SANTRAL tarafında bazen ürün kodu doğrudan gelir, bazen ENT. / ENTES prefixiyle gelir.
        # Örnek: M0005, ENT.M0005, ENTES-M0005 gibi değerler fiyat listesindeki gerçek Ürün Kodu ile normalize edilir.
        urun_kodu_adaylari = set()
        for anahtar in anahtarlar:
            urun_kodu_adaylari.update(urun_kodu_lookup.get(anahtar, set()))

        if urun_kodu_adaylari:
            tercih_anahtarlari = set()
            for deger in [
                eski_deger,
                entes_prefix_sil(eski_deger),
                b2b_temiz_kod_uret(eski_deger),
                sade_kod(eski_deger),
                sade_kod(entes_prefix_sil(eski_deger)),
                sade_kod(b2b_temiz_kod_uret(eski_deger)),
            ]:
                temiz_deger = kod_temizle(deger)
                sade_deger = sade_kod(deger)
                if temiz_deger:
                    tercih_anahtarlari.add(temiz_deger)
                if sade_deger:
                    tercih_anahtarlari.add(sade_deger)

            tam_uyan_kodlar = []
            for kod in sorted(urun_kodu_adaylari):
                kod_anahtarlari = set()
                for deger in [kod, entes_prefix_sil(kod), b2b_temiz_kod_uret(kod), sade_kod(kod)]:
                    temiz_deger = kod_temizle(deger)
                    sade_deger = sade_kod(deger)
                    if temiz_deger:
                        kod_anahtarlari.add(temiz_deger)
                    if sade_deger:
                        kod_anahtarlari.add(sade_deger)

                if tercih_anahtarlari.intersection(kod_anahtarlari):
                    tam_uyan_kodlar.append(kod)

            if len(tam_uyan_kodlar) == 1:
                yeni_kod = tam_uyan_kodlar[0]
            elif len(urun_kodu_adaylari) == 1:
                yeni_kod = list(urun_kodu_adaylari)[0]
            else:
                yeni_kod = ""

            if yeni_kod:
                df.at[idx, kod_kolonu] = yeni_kod
                df.at[idx, durum_kolon] = "Ürün kodu olarak bulundu"

                if kod_temizle(eski_deger) == kod_temizle(yeni_kod):
                    not_metni = "SANTRAL değeri ENTES fiyat listesinde ürün kodu olarak bulundu, değiştirilmedi."
                else:
                    not_metni = f"SANTRAL değeri ENTES fiyat listesinde ürün kodu olarak bulundu ve gerçek ürün koduna normalize edildi: {eski_deger} -> {yeni_kod}"

                df.at[idx, not_kolon] = not_metni
                say_zaten_urun_kodu += 1
                rapor_satirlari.append({
                    "B2B Satır No": idx + 2,
                    "Eski Değer": eski_deger,
                    "Yeni Ürün Kodu": yeni_kod,
                    "Durum": "Ürün kodu olarak bulundu",
                    "Not": not_metni,
                    "Kaynak Dosya": str(kaynak_dosya),
                })
                continue

            bulunan_kodlar = ", ".join(sorted(urun_kodu_adaylari))
            df.at[idx, durum_kolon] = "Çoklu ürün kodu eşleşmesi"
            df.at[idx, not_kolon] = bulunan_kodlar
            say_coklu += 1
            rapor_satirlari.append({
                "B2B Satır No": idx + 2,
                "Eski Değer": eski_deger,
                "Yeni Ürün Kodu": "",
                "Bulunan Kodlar": bulunan_kodlar,
                "Durum": "Çoklu ürün kodu eşleşmesi",
                "Not": "SANTRAL değeri fiyat listesinde birden fazla ürün koduyla eşleşti; manuel kontrol gerekir.",
                "Kaynak Dosya": str(kaynak_dosya),
            })
            continue

        # 2) Önce SANTRAL özel model yazım kuralları denenir.
        # Örnek: RGSR-20S (SVC) -> RGSR-20S-OG, RGSR-15S(SVC) -> RGSR-15S.
        # M kodu elle yazılmaz; fiyat listesindeki hedef model bulunup o satırın Ürün Kodu alınır.
        kural_kod, kural_model = santral_kural_modelden_kod_bul(eski_deger)

        if kural_kod:
            df.at[idx, kod_kolonu] = kural_kod
            df.at[idx, durum_kolon] = "Modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"SANTRAL özel kuralıyla fiyat listesi modeli bulundu: {kural_model} -> {kural_kod}"
            say_modelden_cevrilen += 1

            rapor_satirlari.append({
                "B2B Satır No": idx + 2,
                "Eski Değer": eski_deger,
                "Yeni Ürün Kodu": kural_kod,
                "Bulunan ENTES Modeli": kural_model,
                "Durum": "Modelden ürün koduna çevrildi",
                "Not": "SANTRAL özel kuralıyla model fiyat listesinden ürün koduna çevrildi.",
                "Kaynak Dosya": str(kaynak_dosya),
            })
            continue

        # 3) Değer fiyat listesinde ürün modeli/model numarası olarak aranır.
        adaylar = set()
        for anahtar in anahtarlar:
            adaylar.update(model_lookup.get(anahtar, set()))

        # Aynı ürün kodu aynı modelin farklı temiz anahtarlarından gelmiş olabilir.
        koddan_model = {}
        for kod, model in adaylar:
            koddan_model.setdefault(kod, model)

        if len(koddan_model) == 1:
            yeni_kod = list(koddan_model.keys())[0]
            bulunan_model = koddan_model[yeni_kod]

            df.at[idx, kod_kolonu] = yeni_kod
            df.at[idx, durum_kolon] = "Modelden ürün koduna çevrildi"
            df.at[idx, not_kolon] = f"SANTRAL modeli ENTES fiyat listesinde bulundu: {bulunan_model} -> {yeni_kod}"
            say_modelden_cevrilen += 1

            rapor_satirlari.append({
                "B2B Satır No": idx + 2,
                "Eski Değer": eski_deger,
                "Yeni Ürün Kodu": yeni_kod,
                "Bulunan ENTES Modeli": bulunan_model,
                "Durum": "Modelden ürün koduna çevrildi",
                "Not": "SANTRAL modeli fiyat listesinden ürün koduna çevrildi.",
                "Kaynak Dosya": str(kaynak_dosya),
            })
            continue

        if len(koddan_model) > 1:
            bulunan_kodlar = ", ".join(sorted(koddan_model.keys()))
            bulunan_modeller = " | ".join([f"{k}: {m}" for k, m in sorted(koddan_model.items())])

            df.at[idx, durum_kolon] = "Çoklu model eşleşmesi"
            df.at[idx, not_kolon] = bulunan_kodlar
            say_coklu += 1

            rapor_satirlari.append({
                "B2B Satır No": idx + 2,
                "Eski Değer": eski_deger,
                "Yeni Ürün Kodu": "",
                "Bulunan Kodlar": bulunan_kodlar,
                "Bulunan Modeller": bulunan_modeller,
                "Durum": "Çoklu model eşleşmesi",
                "Not": "Aynı SANTRAL modeli birden fazla ENTES ürün koduyla eşleşti; manuel kontrol gerekir.",
                "Kaynak Dosya": str(kaynak_dosya),
            })
            continue

        df.at[idx, durum_kolon] = "Model karşılığı bulunamadı"
        df.at[idx, not_kolon] = "SANTRAL değeri ENTES fiyat listesinde ne ürün kodu ne de ürün modeli olarak bulunamadı."
        say_model_bulunamayan += 1

        rapor_satirlari.append({
            "B2B Satır No": idx + 2,
            "Eski Değer": eski_deger,
            "Yeni Ürün Kodu": "",
            "Durum": "Model karşılığı bulunamadı",
            "Not": "Fiyat listesinde ürün kodu/model karşılığı bulunamadı.",
            "Kaynak Dosya": str(kaynak_dosya),
        })

    rapor_df = pd.DataFrame(rapor_satirlari)

    bulunan_toplam = say_zaten_urun_kodu + say_modelden_cevrilen
    geriye_kalan_toplam = say_model_bulunamayan + say_coklu + say_bos

    ozet_df = pd.DataFrame([
        {
            "Kalem": "Rapor oluşturma zamanı",
            "Adet": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Açıklama": f"Bu dosya bu çalıştırmada oluşturuldu. Kod sürümü: {KOD_SURUMU}",
        },
        {
            "Kalem": "Toplam kontrol edilen SANTRAL ENTES satırı",
            "Adet": kontrol_edilen_sayi,
            "Açıklama": "Marka filtresine giren ve SANTRAL düzeltme kontrolüne dahil edilen satır sayısı.",
        },
        {
            "Kalem": "Bulunan toplam",
            "Adet": bulunan_toplam,
            "Açıklama": "Ürün kodu olarak bulunanlar + modelden ürün koduna çevrilenler.",
        },
        {
            "Kalem": "Ürün kodu olarak verilenlerin sayısı",
            "Adet": say_zaten_urun_kodu,
            "Açıklama": "SANTRAL ürün kodu/model alanındaki değer ENTES fiyat listesinde zaten ürün kodu olarak bulundu.",
        },
        {
            "Kalem": "Ürün kodu kolonunda model bulunup düzeltilenlerin sayısı",
            "Adet": say_modelden_cevrilen,
            "Açıklama": "SANTRAL ürün kodu/model alanındaki değer ENTES fiyat listesinde model olarak bulundu ve ürün koduna çevrildi.",
        },
        {
            "Kalem": "Geriye kalan toplam",
            "Adet": geriye_kalan_toplam,
            "Açıklama": "Bulunan toplam dışında kalan satırlar.",
        },
        {
            "Kalem": "Geriye kalan - modeli bulunamayanların sayısı",
            "Adet": say_model_bulunamayan,
            "Açıklama": "SANTRAL değeri ENTES fiyat listesinde ne ürün kodu ne de model olarak bulunamadı.",
        },
        {
            "Kalem": "Geriye kalan - çoklu model eşleşmesi sayısı",
            "Adet": say_coklu,
            "Açıklama": "Aynı SANTRAL modeli fiyat listesinde birden fazla ürün koduna denk geldi.",
        },
        {
            "Kalem": "Geriye kalan - boş ürün kodu sayısı",
            "Adet": say_bos,
            "Açıklama": "SANTRAL ürün kodu/model alanı boş olan satırlar.",
        },
        {
            "Kalem": "Kontrol toplamı",
            "Adet": bulunan_toplam + geriye_kalan_toplam,
            "Açıklama": "Bulunan toplam + geriye kalan toplam. Toplam kontrol edilen satırla aynı olmalıdır.",
        },
    ])

    def filtre_durum(durum):
        if rapor_df.empty or "Durum" not in rapor_df.columns:
            return pd.DataFrame()
        return rapor_df[rapor_df["Durum"] == durum].copy()

    modelden_bulunanlar = filtre_durum("Modelden ürün koduna çevrildi")
    modeli_bulunmayanlar = filtre_durum("Model karşılığı bulunamadı")
    urun_kodu_olanlar = pd.concat([d for d in [filtre_durum("Ürün kodu olarak bulundu"), filtre_durum("Zaten ürün kodu")] if d is not None and not d.empty], ignore_index=True, sort=False) if any(d is not None and not d.empty for d in [filtre_durum("Ürün kodu olarak bulundu"), filtre_durum("Zaten ürün kodu")]) else pd.DataFrame()
    coklu_eslesme = filtre_durum("Çoklu model eşleşmesi")
    bos_urun_kodu = filtre_durum("Boş ürün kodu")
    geriye_kalanlar = pd.concat(
        [d for d in [modeli_bulunmayanlar, coklu_eslesme, bos_urun_kodu] if d is not None and not d.empty],
        ignore_index=True,
        sort=False
    ) if any(d is not None and not d.empty for d in [modeli_bulunmayanlar, coklu_eslesme, bos_urun_kodu]) else pd.DataFrame()

    return df

def klemsan_fiyat_listesi_bul():
    """
    KLEMSAN fiyat listesini otomatik bulur.
    Dosya adında KLEMSAN geçmesi yeterlidir; varsa FIYAT/LISTE geçen dosya önceliklidir.
    """
    aday_klasorler = [
        FIYAT_LISTESI_DIR,
        BASE_DIR / "Fiyat Listesi",
        BASE_DIR / "Fiyat",
        BASE_DIR / "Ayarlar",
        BASE_DIR,
    ]

    aday_dosyalar = []

    for klasor in aday_klasorler:
        if not klasor.exists():
            continue

        for uzanti in ["*.xlsx", "*.xls"]:
            for dosya in klasor.glob(uzanti):
                ad_norm = metin_norm(dosya.name)

                if "KLEMSAN" not in ad_norm:
                    continue

                oncelik = 1 if ("FIYAT" in ad_norm or "LISTE" in ad_norm) else 0
                aday_dosyalar.append((oncelik, dosya.stat().st_mtime, dosya))

    if not aday_dosyalar:
        return None

    return max(aday_dosyalar, key=lambda x: (x[0], x[1]))[2]


def klemsan_fiyat_listesi_oku(dosya):
    """
    KLEMSAN fiyat listesinde başlık satırı ilk satırda olmayabilir.
    Model adı ve Noktasız Stok Kodu kolonlarını içeren satırı başlık olarak bulur.
    """
    uzanti = dosya.suffix.lower()

    if uzanti == ".csv":
        return csv_oku(dosya)

    if uzanti not in [".xlsx", ".xls"]:
        raise Exception(f"Desteklenmeyen KLEMSAN fiyat listesi türü: {uzanti}")

    ham_df = pd.read_excel(dosya, dtype=str, header=None).fillna("")

    baslik_satiri = None

    for i, row in ham_df.iterrows():
        satir_metinleri = [metin_norm(x) for x in row.values]

        model_var = any(
            x in [
                "MODEL",
                "MODEL ADI",
                "URUN MODELI",
                "URUN MODEL",
                "MALZEME ADI",
                "URUN ADI",
                "STOK ADI",
            ]
            or "MODEL" in x
            for x in satir_metinleri
        )

        noktasiz_kod_var = any(
            ("NOKTASIZ" in x and ("STOK" in x or "KOD" in x))
            or x in ["NOKTASIZ STOK KODU", "NOKTASIZ KOD", "WEBSERVIS KODU", "WEB SERVIS KODU"]
            for x in satir_metinleri
        )

        if model_var and noktasiz_kod_var:
            baslik_satiri = i
            break

    if baslik_satiri is None:
        print("KLEMSAN fiyat listesinde başlık satırı bulunamadı.")
        print("Aranan başlıklar: Model adı ve Noktasız Stok Kodu")
        print("Bulunan ilk satırlar:")
        print(ham_df.head(8).to_string(index=False))
        raise SystemExit

    df = pd.read_excel(dosya, dtype=str, header=baslik_satiri).fillna("")
    df = tamamen_bos_satirlari_sil(df)

    return df


def klemsan_model_norm(deger):
    deger = metin_norm(deger)
    deger = re.sub(r"\s+", " ", deger).strip()
    return deger


def klemsan_model_anahtarlari_uret(deger):
    """
    KLEMSAN Stok Adı eşleşmesini yazım farklarına dayanıklı yapar.
    Örnek farklar: 2,5 / 2.5, parantez, tire, slash ve boşluk.
    """
    norm = klemsan_model_norm(deger)

    if not norm:
        return []

    adaylar = []
    gorulen = set()

    def ekle(x):
        x = klemsan_model_norm(x)
        if x and x not in gorulen:
            gorulen.add(x)
            adaylar.append(x)

    ekle(norm)
    ekle(norm.replace(".", ","))
    ekle(norm.replace(",", "."))

    sade = sade_kod(norm)
    if sade and len(sade) >= 5:
        ekle(sade)

    parantez_oncesi = norm.split("(", 1)[0].strip()
    if parantez_oncesi and parantez_oncesi != norm:
        ekle(parantez_oncesi)
        parantez_oncesi_sade = sade_kod(parantez_oncesi)
        if parantez_oncesi_sade and len(parantez_oncesi_sade) >= 5:
            ekle(parantez_oncesi_sade)

    # Renk parantezleri fiyat listesinde bazen farklı yazılabiliyor.
    parantezsiz = norm.replace("(", " ").replace(")", " ")
    parantezsiz = re.sub(r"\s+", " ", parantezsiz).strip()
    ekle(parantezsiz)

    parantezsiz_sade = sade_kod(parantezsiz)
    if parantezsiz_sade and len(parantezsiz_sade) >= 5:
        ekle(parantezsiz_sade)

    return adaylar


def santral_klemsan_model_adaylari_uret(deger):
    """
    SANTRAL KLEMSAN ürün adını KLEMSAN fiyat listesindeki model yazımına çevirir.
    Soldaki değer SANTRAL'den gelir; üretilen adaylar fiyat listesindeki model adıyla aranır.
    """
    raw = klemsan_model_norm(deger)
    if not raw:
        return []

    birebir_eslesmeler = {
        "AVK-2,5/4 T": "AVK 2,5/4T",
        "AVK-6/10 T": "AVK 6/10T",
        "AVK-50 T": "AVK 50T",
        "PYK-1,5 MT": "PYK 1,5MT",
        "PYK-4 T": "PYK 4T",
        "PYK-4 2F GRI": "PYK 4-2FN (GRI)",
        "PIK-2,5 N GRI": "PIK 2,5N GRI",
        "PIK-2,5 N BEJ": "PIK 2,5N BEJ",
        "PUK-3 GRI": "PUK 3 GRI",
        "PUK-2 T GRI": "PUK 2T GRI",
        "AVK-4 FS-GRI": "AVK 4FS GRI",
        "ASK-5 LD 24V GRI": "ASK 5LD 24V (NONPOLARITY) (GRI)",
        "ASK-5 LD 220V GRI": "ASK 5LD 220V (NONPOLARITY) (GRI)",
        "SK-1 2,5 MM SIYAH": "SK 1 SIYAH",
        "SK-3 6 MM SIYAH": "SK 3 SIYAH",
        "SK-1 2,5 MM BEYAZ": "SK 1 BEYAZ",
        "TRV-4 S": "TRV 4S (BEJ)",
        "TRV-10": "TRV 10 (TURUNCU)",
        "TRV-4 M": "TRV 4M (TURUNCU)",
        "TRV-4 B": "TRV 4B (TURUNCU)",
        "DARALAN MAKARON APARATI": "TTP345 HSL Makaron Aparatı",
    }

    adaylar = []
    gorulen = set()

    def ekle(x):
        x = klemsan_model_norm(x)
        if x and x not in gorulen:
            gorulen.add(x)
            adaylar.append(x)

    ekle(raw)

    if raw in birebir_eslesmeler:
        ekle(birebir_eslesmeler[raw])

    genel = raw.replace("-", " ")
    genel = re.sub(r"\s+", " ", genel).strip()
    genel = re.sub(r"\b(AVK|PYK|PUK)\s+([0-9,\/]+)\s+(MT|T)\b", r"\1 \2\3", genel)
    genel = re.sub(r"\b(AVK)\s+([0-9,\/]+)\s+(T)\s+(RD|RDS|IRDS)\b", r"\1 \2\3 \4", genel)
    genel = re.sub(r"\b(PIK)\s+([0-9,]+)\s+N\b", r"\1 \2N", genel)
    genel = re.sub(r"\b(PUK)\s+([0-9,]+)\s+T\b", r"\1 \2T", genel)
    genel = re.sub(r"\b(AVK)\s+([0-9,]+)\s+FS\b", r"\1 \2FS", genel)
    genel = re.sub(r"\b(ASK)\s+([0-9]+)\s+LD\b", r"\1 \2LD", genel)

    ekle(genel)

    renk_parantez = {
        "GRI": "GRI",
        "SARI YESIL": "SARI-YESIL",
        "SARI-YESIL": "SARI-YESIL",
        "BEJ": "BEJ",
        "SIYAH": "SIYAH",
        "BEYAZ": "BEYAZ",
    }

    # CTP-2,5 Gri / CTP-2,5E Gri -> CTP 2,5(E) (GRI)
    eslesme = re.match(r"^CTP-([0-9,]+)([A-Z]*)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"CTP {olcu}{ek} ({renk})")

    # SC-PT-2,5 Gri / SC-PTR-2,5 Gri / SC-PTS-2,5 Gri
    # -> SC-PT 2,5 (GRI) / SC-PTR 2,5 (GRI) / SC-PTs 2,5 (GRI)
    eslesme = re.match(r"^(SC-PT(?:R|S)?)-([0-9,]+)([A-Z]*)\s+(.+)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        olcu = eslesme.group(2)
        ek = eslesme.group(3)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(4)))
        if renk:
            if seri == "SC-PTS":
                seri = "SC-PTs"
            ekle(f"{seri} {olcu}{ek} ({renk})")

    # SG-PT ailesi:
    # SG-PT-2,5-1 Gri             -> SG-PT 2,5/1 (GRI)
    # SG-PT-2,5-1L Sarı Yeşil     -> SG-PT 2,5/1L (SARI-YESIL)
    # SG-PT-2,5G-1M Gri           -> SG-PT 2,5G/1M (GRI)
    # SG-PT-2,5CS-4 Gri           -> SG-PT 2,5CS/4 (GRI)
    # SG-PT-2,5GCS-1R Gri         -> SG-PT 2,5GCS/1R (GRI)
    # SG-PT-2,5AN-10 Gri          -> SG-PT 2,5AN/10 (GRI)
    # SG-PT-2,5DB-1 Sarı Yeşil    -> SG-PT 2,5DB/1 (SARI-YESIL)
    eslesme = re.match(r"^(SG-PT)-([0-9,]+[A-Z]*?)-([0-9]+[A-Z]?)\s+(.+)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        olcu = eslesme.group(2)
        kat = eslesme.group(3)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(4)))
        if renk:
            ekle(f"{seri} {olcu}/{kat} ({renk})")

    # PB/NM 2,5 Gri          -> PB 2,5/NM (GRI)
    # PB/NM 2,5K Gri         -> PB 2,5K/NM (GRI)
    # PB/NM 2,5T Sarı Yeşil  -> PB 2,5T/NM
    eslesme = re.match(r"^PB/NM\s+([0-9,]+[A-Z]*)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        renk_raw = klemsan_model_norm(eslesme.group(2))

        if renk_raw in ["SARI YESIL", "SARI-YESIL"]:
            ekle(f"PB {olcu}/NM")
        else:
            renk = renk_parantez.get(renk_raw)
            if renk:
                ekle(f"PB {olcu}/NM ({renk})")

    # SRD-2,5K Siyah / Yeşil -> SRD_K2,5 (SIYAH/YESIL)
    eslesme = re.match(r"^SRD-([0-9,]+)([A-Z]*)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk_raw = klemsan_model_norm(eslesme.group(3))
        renk = {
            "SIYAH": "SIYAH",
            "YESIL": "YESIL",
            "YESİL": "YESIL",
        }.get(renk_raw)
        if renk:
            ekle(f"SRD_{ek}{olcu} ({renk})")

    # CTP-2,5-FT Sarı Yeşil -> CTP 2,5FT (SARI-YESIL) / CTP 2,5-FT (SARI-YESIL)
    eslesme = re.match(r"^CTP-([0-9,]+)-([A-Z]+)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"CTP {olcu}{ek} ({renk})")
            ekle(f"CTP {olcu}-{ek} ({renk})")

    # SRD-2,5T / SRD-2,5CT / SRD-2,5ET -> SRD 2,5T / SRD_T2,5 adayları
    eslesme = re.match(r"^SRD-([0-9,]+)(T|CT|ET)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        ekle(f"SRD {olcu}{ek}")
        ekle(f"SRD_{ek}{olcu}")

    # SRD-2,5-2T / SRD-4-2T -> SRD 2,5-2T / SRD 2,5/2T adayları
    eslesme = re.match(r"^SRD-([0-9,]+)-([0-9]+T)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        kat = eslesme.group(2)
        ekle(f"SRD {olcu}-{kat}")
        ekle(f"SRD {olcu}/{kat}")

    # SRD-2,5-2T S/YEŞİL -> SRD 2,5-2T (SARI-YESIL) / SRD 2,5/2T (SARI-YESIL)
    eslesme = re.match(r"^SRD-([0-9,]+)-([0-9]+T)\s+S/?YESIL$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        kat = eslesme.group(2)
        ekle(f"SRD {olcu}-{kat} (SARI-YESIL)")
        ekle(f"SRD {olcu}/{kat} (SARI-YESIL)")

    # SRD-4-FK GRİ -> SRD 4FK GRI / SRD 4-FK GRI
    eslesme = re.match(r"^SRD-([0-9,]+)-([A-Z]+)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"SRD {olcu}{ek} {renk}")
            ekle(f"SRD {olcu}-{ek} {renk}")
            if ek == "FK":
                ekle(f"SRD {olcu}-2FK ({renk})")

    # AVK-4 FSLD 24V Gri -> AVK 4FSLD 24V GRI / AVK 4FSLD 24V DC GRI adayları
    eslesme = re.match(r"^AVK-([0-9,]+)\s+FSLD\s+([0-9]+)V\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        volt = eslesme.group(2)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"AVK {olcu}FSLD {volt}V {renk}")
            ekle(f"AVK {olcu}FSLD {volt}V DC {renk}")
            ekle(f"AVK {olcu}FSLD {volt}VDC {renk}")
            ekle(f"AVK {olcu}FSLD {volt}V (NONPOLARITY) ({renk})")

    # ASK-5 Gri -> ASK 5S (GRI), ASK-5 LD 110V Gri -> ASK 5LD 110V (NONPOLARITY) (GRI)
    eslesme = re.match(r"^ASK-([0-9]+)\s+(.+)$", raw)
    if eslesme:
        no = eslesme.group(1)
        kalan = klemsan_model_norm(eslesme.group(2))
        renk = renk_parantez.get(kalan)
        if renk:
            ekle(f"ASK {no}S ({renk})")

    eslesme = re.match(r"^ASK-([0-9]+)\s+LD\s+([0-9]+)V\s+(.+)$", raw)
    if eslesme:
        no = eslesme.group(1)
        volt = eslesme.group(2)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"ASK {no}LD {volt}V (NONPOLARITY) ({renk})")

    # WGO-2N Kovansız/Bej     -> WGO 2N BEJ KOVANSIZ
    # WGO-2N Tek Kovan/Gri    -> WGO 2N GRI TEK KOVANLI
    # WGO-2N Çift Test/Bej    -> WGO 2N BEJ CIFT TEST KOVANLI
    # WGO-2N Tek Test/Gri     -> WGO 2N GRI TEK TEST KOVANLI
    # WGO-2N Çift Kovan/B     -> WGO 2N BEJ CIFT KOVANLI
    eslesme = re.match(r"^WGO-([0-9]+N?)\s+(.+?)/(.+)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        tip_raw = klemsan_model_norm(eslesme.group(2))
        renk_raw = klemsan_model_norm(eslesme.group(3))

        renk = {
            "B": "BEJ",
            "BEJ": "BEJ",
            "G": "GRI",
            "GRI": "GRI",
        }.get(renk_raw)

        tip = {
            "KOVANSIZ": "KOVANSIZ",
            "TEK KOVAN": "TEK KOVANLI",
            "TEK KOV": "TEK KOVANLI",
            "TEK TEST": "TEK TEST KOVANLI",
            "CIFT TEST": "CIFT TEST KOVANLI",
            "C TEST": "CIFT TEST KOVANLI",
            "C.TEST": "CIFT TEST KOVANLI",
            "CIFT KOVAN": "CIFT KOVANLI",
            "CIFT KOV": "CIFT KOVANLI",
        }.get(tip_raw)

        if renk and tip:
            ekle(f"WGO {seri} {renk} {tip}")

            # WGO 3 fiyat listesinde çift kovan ifadesi çift test kovanlı olarak geçebiliyor.
            if seri == "3" and tip == "CIFT KOVANLI":
                ekle(f"WGO {seri} {renk} CIFT TEST KOVANLI")

    # WGO-4 Ç.Test Kov/Gri -> WGO 4 GRI CIFT TEST KOVANLI + WGO 4 GRI yedek aday
    eslesme = re.match(r"^WGO-([0-9]+)\s+(.+?)/(.+)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        tip_raw = klemsan_model_norm(eslesme.group(2))
        renk_raw = klemsan_model_norm(eslesme.group(3))

        renk = {
            "B": "BEJ",
            "BEJ": "BEJ",
            "G": "GRI",
            "GRI": "GRI",
        }.get(renk_raw)

        if renk and (
            "CIFT" in tip_raw
            or "C TEST" in tip_raw
            or "C.TEST" in tip_raw
            or "TEST" in tip_raw
        ):
            ekle(f"WGO {seri} {renk} CIFT TEST KOVANLI")
            ekle(f"WGO {seri} {renk}")

    # WGO-5 D5 Gri -> WGO D5 GRI, normal WGO-5 Gri gelirse WGO 5 GRI
    eslesme = re.match(r"^WGO-([0-9]+)\s+(D[0-9]+)\s+(.+)$", raw)
    if eslesme:
        d_kodu = eslesme.group(2)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"WGO {d_kodu} {renk}")

    eslesme = re.match(r"^WGO-([0-9]+)\s+(.+)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(2)))
        if renk:
            ekle(f"WGO {seri} {renk}")

    # FPAL-10 SİYAH -> FPAL10 SPIRAL 10mm
    eslesme = re.match(r"^FPAL-([0-9]+)\s+.*$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ekle(f"FPAL{olcu} SPIRAL {olcu}mm")
        ekle(f"FPAL{olcu} SPIRAL {olcu}MM")

    # FPA10-M12 SİYAH -> FPA10-M12 RAKOR 10mm
    # FPA13-M20 SİYAH -> FPA13-M20B / FPA13-M20 RAKOR 13mm adayları
    eslesme = re.match(r"^FPA([0-9]+)-M([0-9]+)\s+.*$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        metrik = eslesme.group(2)
        ekle(f"FPA{olcu}-M{metrik} RAKOR {olcu}mm")
        ekle(f"FPA{olcu}-M{metrik} RAKOR {olcu}MM")
        ekle(f"FPA{olcu}-M{metrik}B")

    # FPA13-PG9 SİYAH -> FPA13-PG9 RAKOR 13mm
    # FPA42-PG36 SİYAH -> FPA42-PG36B RAKOR 42mm
    eslesme = re.match(r"^FPA([0-9]+)-PG([0-9]+)\s+.*$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        pg = eslesme.group(2)
        ekle(f"FPA{olcu}-PG{pg} RAKOR {olcu}mm")
        ekle(f"FPA{olcu}-PG{pg} RAKOR {olcu}MM")
        ekle(f"FPA{olcu}-PG{pg}B RAKOR {olcu}mm")
        ekle(f"FPA{olcu}-PG{pg}B RAKOR {olcu}MM")

    # FPA21D-PG16 SİYAH -> FPA21D-PG16
    eslesme = re.match(r"^FPA([0-9]+D)-PG([0-9]+)\s+.*$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        pg = eslesme.group(2)
        ekle(f"FPA{olcu}-PG{pg}")

    # FPA13-T3 SİYAH -> FPA 13 T3
    eslesme = re.match(r"^FPA([0-9]+)-T3\s+.*$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ekle(f"FPA {olcu} T3")

    def kablo_kanali_olcu(kod, seri=""):
        kod = str(kod or "").strip()
        if len(kod) == 4:
            if seri == "KKC":
                ilk = int(kod[:2])
            elif kod[:2] in ["10", "12", "15"]:
                ilk = int(kod[:2]) * 10
            else:
                ilk = int(kod[:2])

            son = int(kod[2:])
            if son == 1:
                son = 100
            elif son in [6, 8, 10]:
                son = son * 10
            return f"{ilk}x{son}"
        return ""

    # KKC-1518 PVC KANAL -> KKC 1518 15x18 KABLO KANALI (GRI)
    eslesme = re.match(r"^(KKC)-([0-9]{4})\s+PVC\s+KANAL$", raw)
    if eslesme:
        seri = eslesme.group(1)
        kod = eslesme.group(2)
        olcu = kablo_kanali_olcu(kod, seri)
        if olcu:
            ekle(f"{seri} {kod} {olcu} KABLO KANALI (GRI)")

    # KKC-1518 YAPISKANLI KANAL -> KKC 1518 15x18 YAPISKANLI KABLO KANALI (GRI)
    eslesme = re.match(r"^(KKC)-([0-9]{4})\s+YAPISKANLI\s+KANAL$", raw)
    if eslesme:
        seri = eslesme.group(1)
        kod = eslesme.group(2)
        olcu = kablo_kanali_olcu(kod, seri)
        if olcu:
            ekle(f"{seri} {kod} {olcu} YAPISKANLI KABLO KANALI (GRI)")

    # KKD-2530 PVC KANAL -> KKD 2530 25x30 KABLO KANALI (GRI)
    # KKD-4060 GENIS OLUK -> KKD 4060 40x60 KABLO KANALI (GRI) (KALIN SLOT)
    eslesme = re.match(r"^(KKD)-([0-9]{4})\s+(PVC\s+KANAL|GENIS\s+OLUK)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        kod = eslesme.group(2)
        tip = eslesme.group(3)
        olcu = kablo_kanali_olcu(kod, seri)
        if olcu:
            ekle(f"{seri} {kod} {olcu} KABLO KANALI (GRI)")
            if "GENIS" in tip:
                ekle(f"{seri} {kod} {olcu} KABLO KANALI (GRI) (KALIN SLOT)")

    # KKT-4040 -> KKT 4040 40x40 KABLO TUTUCU (SIYAH)
    eslesme = re.match(r"^(KKT)-([0-9]{4})$", raw)
    if eslesme:
        seri = eslesme.group(1)
        kod = eslesme.group(2)
        olcu = kablo_kanali_olcu(kod, seri)
        if olcu:
            ekle(f"{seri} {kod} {olcu} KABLO TUTUCU (SIYAH)")

    # KKK-15 PVC KANAL KAPAGI -> KKK 15mm KANAL KAPAGI (GRI)
    eslesme = re.match(r"^(KKK)-([0-9]+)\s+PVC\s+KANAL\s+KAPAGI$", raw)
    if eslesme:
        seri = eslesme.group(1)
        olcu = eslesme.group(2)
        ekle(f"{seri} {olcu}mm KANAL KAPAGI (GRI)")
        ekle(f"{seri} {olcu}MM KANAL KAPAGI (GRI)")

    # KDK-25 PVC KANAL KAPAGI -> KDK 25 25mm KANAL KAPAGI (GRI)
    eslesme = re.match(r"^(KDK)-([0-9]+)\s+PVC\s+KANAL\s+KAPAGI$", raw)
    if eslesme:
        seri = eslesme.group(1)
        olcu = eslesme.group(2)
        ekle(f"{seri} {olcu} {olcu}mm KANAL KAPAGI (GRI)")
        ekle(f"{seri} {olcu} {olcu}MM KANAL KAPAGI (GRI)")

    # KHF-2530 H.F. KANAL -> KHF 2530 25x30 H.F./HALOGEN FREE KABLO KANALI (GRI)
    eslesme = re.match(r"^(KHF)-([0-9]{4})\s+H\.F\.\s+KANAL$", raw)
    if eslesme:
        seri = eslesme.group(1)
        kod = eslesme.group(2)
        olcu = kablo_kanali_olcu(kod, seri)
        if olcu:
            ekle(f"{seri} {kod} {olcu} H.F. KABLO KANALI (GRI)")
            ekle(f"{seri} {kod} {olcu} HALOGEN FREE KABLO KANALI (GRI)")
            ekle(f"{seri} {kod} {olcu} KABLO KANALI (GRI)")

    # HFK-25 H.F. KANAL KAPAGI -> HFK 25 25mm H.F./HALOGEN FREE KANAL KAPAGI (GRI)
    eslesme = re.match(r"^(HFK)-([0-9]+)\s+H\.F\.\s+KANAL\s+KAPAGI$", raw)
    if eslesme:
        seri = eslesme.group(1)
        olcu = eslesme.group(2)
        ekle(f"{seri} {olcu} {olcu}mm H.F. KANAL KAPAGI (GRI)")
        ekle(f"{seri} {olcu} {olcu}MM H.F. KANAL KAPAGI (GRI)")
        ekle(f"{seri} {olcu} {olcu}mm HALOGEN FREE KANAL KAPAGI (GRI)")
        ekle(f"{seri} {olcu} {olcu}MM HALOGEN FREE KANAL KAPAGI (GRI)")
        ekle(f"{seri} {olcu}mm HALOGEN FREE KANAL KAPAK (GRI)")
        ekle(f"{seri} {olcu}MM HALOGEN FREE KANAL KAPAK (GRI)")

    # KKB / KKM kısa adları fiyat listesinde açıklamalı geçiyor.
    if raw == "KKB":
        ekle("KKB KABLO BAGI APARAT")
    elif raw == "KKM":
        ekle("KKM KABLO AYIRICI MANDAL")

    # SK-3 6 mm Beyaz -> SK 3 BEYAZ
    eslesme = re.match(r"^SK-([0-9]+)\s+[0-9]+\s+MM\s+(.+)$", raw)
    if eslesme:
        no = eslesme.group(1)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(2)))
        if renk:
            ekle(f"SK {no} {renk}")

    # KIK-6 -> KIK 1/6
    eslesme = re.match(r"^KIK-([0-9,]+)$", raw)
    if eslesme:
        ekle(f"KIK 1/{eslesme.group(1)}")

    # PPE 6,5 -> PLASTIK PERCIN 6,5 MM - PPE 6,5
    eslesme = re.match(r"^PPE\s+([0-9,]+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ekle(f"PLASTIK PERCIN {olcu} MM - PPE {olcu}")

    # SETR -> PERCIN CAKMA ALETI - SETR
    if raw == "SETR":
        ekle("PERCIN CAKMA ALETI - SETR")

    if raw == "BICAK SETI":
        ekle("STRIPPPER 16 ICIN YEDEK BICAK SETI")

    # KBE-3 -> KBE-3 - SI 1001
    if raw == "KBE-3":
        ekle("KBE-3 - SI 1001")

    # KBEP-209W Beyaz -> KBEP 209 W - SI HF 44W
    eslesme = re.match(r"^KBEP-([0-9]+)W\s+BEYAZ$", raw)
    if eslesme:
        ekle(f"KBEP {eslesme.group(1)} W - SI HF 44W")

    # KIP-10W Beyaz -> KIP 10 W - SI HF 2W/10
    # KIP-10Y Sari  -> KIP 10 Y - SI HF 2Y/10
    eslesme = re.match(r"^KIP-([0-9]+)([WY])\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        tip = eslesme.group(2)
        renk_raw = klemsan_model_norm(eslesme.group(3))
        if (tip == "W" and renk_raw == "BEYAZ") or (tip == "Y" and renk_raw == "SARI"):
            ekle(f"KIP {olcu} {tip} - SI HF 2{tip}/{olcu}")

    # MSB-2,5 -> MSB 5x20 2,5 A.
    eslesme = re.match(r"^MSB-([0-9]+(?:,[0-9]+)?)$", raw)
    if eslesme:
        amper = eslesme.group(1)
        ekle(f"MSB 5x20 {amper} A.")

    # SBK-16/5 -> SBK 16-5 YAY BASKILI (AKB VIDALI)
    eslesme = re.match(r"^SBK-([0-9]+)/([0-9]+)$", raw)
    if eslesme:
        govde = eslesme.group(1)
        kademe = eslesme.group(2)
        ekle(f"SBK {govde}/{kademe} YAY BASKILI (AKB VIDALI)")
        ekle(f"SBK {govde}-{kademe} YAY BASKILI (AKB VIDALI)")

    # HSL-02-Y -> HSL 02-Y 400m D:0.25-1
    eslesme = re.match(r"^HSL-([0-9]+(?:,[0-9]+)?)-Y$", raw)
    if eslesme:
        olcu = eslesme.group(1).replace(",", ".")
        hsl_kod = {
            "02": "02-Y 400m D:0.25-1",
            "03": "03-Y 400m D:1-1.5",
            "4.5": "4.5-Y 100m D:2.5",
            "06": "06-Y 200m D:4-6",
            "10": "10-Y 100m D:6-10",
        }.get(olcu)
        if hsl_kod:
            ekle(f"HSL {hsl_kod}")

    # OYE-07Y -> OYE-07 Y 44x7 YAPISKAN ETIKET
    eslesme = re.match(r"^OYE-([0-9]+)Y$", raw)
    if eslesme:
        no = eslesme.group(1)
        oye_olcu = {
            "07": "44x7",
            "06": "38x10",
            "05": "20x8",
            "03": "12x9",
            "01": "17x10",
            "04": "15x9",
            "02": "15x6",
        }.get(no)
        if oye_olcu:
            ekle(f"OYE-{no} Y {oye_olcu} YAPISKAN ETIKET")

    # TTSL-40-WH 12,7x19,5 -> TTSL 40-W 12,7 x 19,5mm D:3
    eslesme = re.match(r"^TTSL-40-WH\s+([0-9,]+)X([0-9,]+)$", raw)
    if eslesme:
        en = eslesme.group(1)
        boy = eslesme.group(2)
        d_degeri = {
            ("12,7", "19,5"): "3",
            ("19", "23,8"): "4",
            ("12,7", "36,5"): "7",
            ("25,4", "36,5"): "7",
            ("19", "44,5"): "10",
        }.get((en, boy))
        if d_degeri:
            ekle(f"TTSL 40-W {en} x {boy}mm D:{d_degeri}")

    # LSL-03SV GUMUS -> LSL 03-SV 3cm x 20m
    eslesme = re.match(r"^LSL-([0-9]+)(W|SV|YL)\s+(.+)$", raw)
    if eslesme:
        no = eslesme.group(1)
        kod = eslesme.group(2)
        renk_raw = klemsan_model_norm(eslesme.group(3))
        renk_ok = (
            (kod == "W" and renk_raw == "BEYAZ")
            or (kod == "SV" and renk_raw == "GUMUS")
            or (kod == "YL" and renk_raw == "SARI")
            or (kod == "SV" and renk_raw == "SARI")
        )
        if renk_ok:
            if renk_raw == "SARI":
                kod = "YL"
            cm = str(int(no))
            ekle(f"LSL {no}-{kod} {cm}cm x 20m")

    # YDR-21 SIYAH -> YDR 21- 2 X 16
    eslesme = re.match(r"^YDR-([0-9]+)\s+SIYAH$", raw)
    if eslesme:
        no = eslesme.group(1)
        ydr_karsi = {
            "21": "YDR 21- 2 X 16",
            "28": "YDR 28- 2 X 21",
            "34": "YDR 34- 2 X 28",
            "42": "YDR 42-2 X 34 SIYAH",
            "54": "YDR 54-2 X 42 SIYAH",
        }.get(no)
        if ydr_karsi:
            ekle(ydr_karsi)

    # KLEA 320D -> KLEA 320P-D, ECRAS 100 VCF -> ECRAS VCF
    eslesme = re.match(r"^KLEA\s+([0-9]+)D$", raw)
    if eslesme:
        ekle(f"KLEA {eslesme.group(1)}P-D")

    if raw == "ECRAS 100 VCF":
        ekle("ECRAS VCF")

    if raw == "KOIP-TT W":
        ekle("KOIP TT W - SI HF 32 MWK")

    # UTOR4i RS-485 -> UTOR-4I
    eslesme = re.match(r"^UTOR([0-9]+)I\s+RS-[0-9]+$", raw)
    if eslesme:
        ekle(f"UTOR-{eslesme.group(1)}I")

    eslesme = re.match(r"^UTORT([0-9]+)I\s+TTL\s+\([0-9,]+V\)$", raw)
    if eslesme:
        ekle(f"UTOR-T{eslesme.group(1)}I")

    # PAP-T Pano Prizi -> PAP T
    eslesme = re.match(r"^PAP-([A-Z]+)\s+PANO\s+PRIZI$", raw)
    if eslesme:
        ek = eslesme.group(1)
        ekle(f"PAP {ek}")

    # PG9-N SIYAH -> PG9-N SOMUN (PG9)
    eslesme = re.match(r"^(PG[0-9]+)-N\s+SIYAH$", raw)
    if eslesme:
        pg = eslesme.group(1)
        ekle(f"{pg}-N SOMUN ({pg})")
        ekle(f"{pg} -N SOMUN ({pg})")
        ekle(f"{pg} _ N SOMUN")

    # CC-13 SIYAH -> CC13 SPRIAL MESNEDI 13mm
    eslesme = re.match(r"^CC-([0-9]+)\s+SIYAH$", raw)
    if eslesme:
        no = eslesme.group(1)
        if no == "10":
            ekle("CC10B SPIRAL MESNEDI 10mm")
        else:
            ekle(f"CC{no} SPRIAL MESNEDI {no}mm")
            ekle(f"CC{no} SPIRAL MESNEDI {no}mm")

    # YK-1 B -> YK 1 BEYAZ, YK-2 S -> YK 2 SIYAH
    eslesme = re.match(r"^YK-([0-9]+)\s+([BS])$", raw)
    if eslesme:
        no = eslesme.group(1)
        renk = "BEYAZ" if eslesme.group(2) == "B" else "SIYAH"
        ekle(f"YK {no} {renk}")

    # DM-6 S -> DM 6-S
    eslesme = re.match(r"^DM-([0-9]+)\s+S$", raw)
    if eslesme:
        ekle(f"DM {eslesme.group(1)}-S")

    # KS-2K / KS-4K / KS-C4K / CLP/KRB-C-15L
    if raw == "KS-2K":
        ekle("KS 2CO SOCKET (12A)")
    elif raw == "KS-4K":
        ekle("KS 4CO SOCKET (10A)")
    elif raw == "KS-C4K":
        ekle("KS-C 4CO SOCKET (10A)")
    elif raw == "KS-C2K":
        ekle("KS-C 2CO SOCKET (12A)")
    elif raw == "LB-KS":
        ekle("LB/KS-ECO ID TAG")
    elif raw == "CLP/KRB-C-15L":
        ekle("CLP/KRB-C-15L PLASTIC CLIPS")
    elif raw == "KS-R 3K":
        ekle("KS-R 3CO SOCKET (12A)")
    elif raw == "CLM/KS-R":
        ekle("CLM/KS-R METAL CLIPS")

    # WG-EKI-N 110010N -> WG-EKI-C satırı, kodu fiyat listesinden bulunur.
    eslesme = re.match(r"^WG-EKI-N\s+([0-9]+N)$", raw)
    if eslesme:
        ekle("WG-EKI-C")

    # LSK-2 4lü 24VAC/DC -> LSK 4 LU 24V
    # LSK-2 6lı 220VAC/DC -> LSK 6 LI 220V
    eslesme = re.match(r"^LSK-2\s+([0-9]+)L[UI]\s+([0-9]+)VAC/DC$", raw)
    if eslesme:
        adet = eslesme.group(1)
        volt = eslesme.group(2)
        li_lu = "LI" if adet == "6" else "LU"
        ekle(f"LSK {adet} {li_lu} {volt}V")

    # KMY200024 -> KMY200024L
    # KMY200048 -> KMY200048LT
    # KMY200024LTD zaten doğrudan model olarak aranır.
    eslesme = re.match(r"^(KMY(?:200|400)(?:012|024|048|110|220|524|548|615|730))(LTD|LT|L)?$", raw)
    if eslesme:
        kmy_kod = eslesme.group(1)
        ek = eslesme.group(2)

        if ek:
            ekle(f"{kmy_kod}{ek}")
        elif kmy_kod.endswith("024"):
            ekle(f"{kmy_kod}L")
        elif kmy_kod.endswith("220"):
            ekle(kmy_kod)
            ekle(f"{kmy_kod}LT")
        else:
            ekle(f"{kmy_kod}LT")

    # KRG100012 -> KRG100012 (fiyat listesinde açıklama parantezli gelebilir)
    # KRG200024 -> KRG200024(24VDC,2CO,8A,RELAY)
    eslesme = re.match(r"^(KRG(?:100|200|201)(?:005|006|009|012|024|048|060|110|524|615|730))$", raw)
    if eslesme:
        ekle(eslesme.group(1))

    # KRB-1CO / KRC-2CO -> KRB1CO SOCKET / KRC2CO SOCKET
    eslesme = re.match(r"^(KR[BC])-([12]CO)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        kontak = eslesme.group(2)
        ekle(f"{seri}{kontak} SOCKET")

    # KRB300012LT -> KRB300012LT (...), KRB300024LT -> 24KRB300024LT (...)
    eslesme = re.match(r"^(KRB300(?:012|024|048|110|524|615|730)LT)$", raw)
    if eslesme:
        krb_kod = eslesme.group(1)
        ekle(krb_kod)
        if krb_kod == "KRB300024LT":
            ekle(f"24{krb_kod}")

    # YKB-100 B 2,5 X 100 -> YKB 100 (2,5X100) BEYAZ
    # YKB-450 S 8 X 450   -> YKB 450 (8,0X450) SIYAH
    eslesme = re.match(r"^YKB-([0-9]+)\s+([BS])\s+([0-9]+(?:,[0-9]+)?)\s*X\s*([0-9]+)$", raw)
    if eslesme:
        boy = eslesme.group(1)
        renk_harfi = eslesme.group(2)
        en = eslesme.group(3)
        uzunluk = eslesme.group(4)

        renk = {
            "B": "BEYAZ",
            "S": "SIYAH",
        }.get(renk_harfi)

        en_adaylari = [en]
        if "," not in en:
            en_adaylari.append(f"{en},0")

        if renk:
            for en_adayi in en_adaylari:
                ekle(f"YKB {boy} ({en_adayi}X{uzunluk}) {renk}")
                ekle(f"YKB {boy} ({en_adayi}x{uzunluk}) {renk}")
                ekle(f"YKB{boy} ({en_adayi}X{uzunluk}) {renk}")
                ekle(f"YKB{boy} ({en_adayi}x{uzunluk}) {renk}")

    # İKY-1/8 -> IKY 1,0/8
    # İKY-0,50/8 EKO -> IKY 0,50/8 ECO
    eslesme = re.match(r"^IKY-([0-9]+(?:,[0-9]+)?)/([0-9]+)(?:\s+(EKO|ECO))?$", raw)
    if eslesme:
        kesit = eslesme.group(1)
        boy = eslesme.group(2)
        ek = eslesme.group(3)

        ek_text = " ECO" if ek else ""

        kesit_adaylari = [kesit]
        if "," not in kesit:
            kesit_adaylari.append(f"{kesit},0")

        for kesit_adayi in kesit_adaylari:
            ekle(f"IKY {kesit_adayi}/{boy}{ek_text}")

    # KY-1/10 -> KY 1,0/10
    eslesme = re.match(r"^KY-([0-9]+(?:,[0-9]+)?)/([0-9]+)(?:\s+(EKO|ECO))?$", raw)
    if eslesme:
        kesit = eslesme.group(1)
        boy = eslesme.group(2)
        ek = eslesme.group(3)
        ek_text = " ECO" if ek else ""

        kesit_adaylari = [kesit]
        if "," not in kesit:
            kesit_adaylari.append(f"{kesit},0")

        for kesit_adayi in kesit_adaylari:
            ekle(f"KY {kesit_adayi}/{boy}{ek_text}")

    npp_renk = {
        "GRI": "GRI",
        "BEJ": "BEJ",
        "KRMZ": "KIRMIZI",
        "KIRMIZI": "KIRMIZI",
        "TRNCU": "TURUNCU",
        "TURUNCU": "TURUNCU",
        "MAVI": "MAVI",
        "YESIL": "YESIL",
        "SARI": "SARI",
        "SIYAH": "SIYAH",
        "BEYAZ": "BEYAZ",
    }

    # NPP-AVK 2,5-10 Krmz -> NPP / AVK 2,5-10 KIRMIZI
    eslesme = re.match(r"^NPP-AVK\s+([0-9,]+-[0-9]+)\s+([A-Z]+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(2)))
        if renk:
            ekle(f"NPP / AVK {olcu} {renk}")
            ekle(f"NPP/AVK {olcu} {renk}")

    # NPP-AVK Y4 AVK Y6Gri -> NPP / AVK-Y4/Y6 (GRI)
    eslesme = re.match(r"^NPP-AVK\s+Y([0-9]+)\s+AVK\s+Y([0-9]+)\s*([A-Z]+)$", raw)
    if eslesme:
        ilk = eslesme.group(1)
        ikinci = eslesme.group(2)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"NPP / AVK-Y{ilk}/Y{ikinci} ({renk})")
            ekle(f"NPP / AVK Y{ilk}/Y{ikinci} ({renk})")

    # NPP-AVK 4A AVK 4RGri -> NPP / AVK 4R-A (GRI)
    eslesme = re.match(r"^NPP-AVK\s+([0-9,]+)A\s+AVK\s+([0-9,]+)R\s*([A-Z]+)$", raw)
    if eslesme:
        olcu = eslesme.group(2)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"NPP / AVK {olcu}R-A ({renk})")

    # NPP-AVK 4CC Gri -> NPP / AVK 4CC-CCA (GRI)
    eslesme = re.match(r"^NPP-AVK\s+([0-9,]+)CC\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(2)))
        if renk:
            ekle(f"NPP / AVK {olcu}CC-CCA ({renk})")
            ekle(f"NPP / AVK {olcu}CC - CCA ({renk})")

    # NPP-PYK 2,5-2FT Gri -> NPP/PYK 2,5-2FT/3F GRI
    eslesme = re.match(r"^NPP-PYK\s+([0-9,]+)-([0-9]+FT)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"NPP/PYK {olcu}-{ek}/3F {renk}")
            ekle(f"NPP / PYK {olcu}-{ek}/3F {renk}")

    # NPP-PYK 4-2F Gri -> NPP / PYK 4-2FN GRI
    eslesme = re.match(r"^NPP-PYK\s+([0-9,]+)-([0-9]+F)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"NPP / PYK {olcu}-{ek}N {renk}")
            ekle(f"NPP/PYK {olcu}-{ek}N {renk}")
            ekle(f"NPP/PYK {olcu}-{ek} {renk}")

    # NPP-PYK 1,5C Gri / NPP-PYK 1,5E Gri -> NPP/PYK 1,5C/E GRI adayları
    eslesme = re.match(r"^NPP-PYK\s+([0-9,]+)([A-Z]+)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        renk = npp_renk.get(klemsan_model_norm(eslesme.group(3)))
        if renk:
            ekle(f"NPP/PYK {olcu}{ek} {renk}")
            ekle(f"NPP / PYK {olcu}{ek} {renk}")
            if ek in ["C", "E"]:
                ekle(f"NPP / PYK {olcu}M{ek} {renk}")
                ekle(f"NPP/PYK {olcu}M{ek} {renk}")

    # NPP-PYK 4T / 6T / 10T / 16T -> NPP/PYK 4 (SARI-YESIL)
    eslesme = re.match(r"^NPP-PYK\s+([0-9,]+)T$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ekle(f"NPP/PYK {olcu} (SARI-YESIL)")
        ekle(f"NPP / PYK {olcu} (SARI-YESIL)")

    # NPP-PYK 1,5T/2,5T -> NPP/PYK 1,5/2,5 (SARI-YESIL)
    eslesme = re.match(r"^NPP-PYK\s+([0-9,]+)T/([0-9,]+)T$", raw)
    if eslesme:
        ilk = eslesme.group(1)
        ikinci = eslesme.group(2)
        ekle(f"NPP/PYK {ilk}/{ikinci} (SARI-YESIL)")
        ekle(f"NPP / PYK {ilk}/{ikinci} (SARI-YESIL)")

    # MYPK-2,5 / MYSK-2,5 / PYKMR-2,5 / PYKM-2,5 -> seri ölçü GRI
    eslesme = re.match(r"^(MYPK|MYSK|PYKMR|PYKM)-([0-9,]+)$", raw)
    if eslesme:
        seri = eslesme.group(1)
        olcu = eslesme.group(2)
        ekle(f"{seri} {olcu} GRI")

    # PYK-6 ET / PYK-10 ET -> PYK 6ET (SARI-YESIL)
    # PYK-6 CT             -> PYK 6CT (SARI-YESIL)
    eslesme = re.match(r"^PYK-([0-9,]+)\s+(ET|CT)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        ek = eslesme.group(2)
        ekle(f"PYK {olcu}{ek} (SARI-YESIL)")
        # Bazı fiyat listelerinde ET renginde parantez olmayabilir.
        ekle(f"PYK {olcu}{ek} SARI-YESIL")

    # PYK-2,5 2FDA DİYOTLU -> PYK 2,5-2FD-A (GRI)
    # PYK-2,5 2FDB DİYOTLU -> PYK 2,5-2FD-B (GRI)
    eslesme = re.match(r"^PYK-([0-9,]+)\s+2FD([AB])\s+DIYOTLU$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        diyot_tipi = eslesme.group(2)
        ekle(f"PYK {olcu}-2FD-{diyot_tipi} (GRI)")

    # PUK-3 SLD/NPN Gri -> PUK 3SLD 24 VDC NPN GRI
    eslesme = re.match(r"^PUK-([0-9,]+)\s+SLD/(NPN|PNP)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        tip = eslesme.group(2)
        renk_raw = klemsan_model_norm(eslesme.group(3))
        renk = {
            "GRI": "GRI",
            "BEJ": "BEJ",
        }.get(renk_raw)
        if renk:
            ekle(f"PUK {olcu}SLD 24 VDC {tip} {renk}")

    # PYK-3SLD 24V NPN Gri -> PYK 3S NPN 24VDC GRI
    eslesme = re.match(r"^PYK-([0-9,]+)SLD\s+([0-9]+)V\s+(NPN|PNP)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        volt = eslesme.group(2)
        tip = eslesme.group(3)
        renk = renk_parantez.get(klemsan_model_norm(eslesme.group(4)))
        if renk:
            ekle(f"PYK {olcu}S {tip} {volt}VDC {renk}")

    # PYK-4S LD 24V DC Gri  -> PYK 4SLD 24V DC (5x20) GRI
    # PYK-4S LD 220V AC G   -> PYK 4SLD 220V AC (5x20) GRI
    eslesme = re.match(r"^PYK-([0-9,]+)S\s+LD\s+([0-9]+)V\s+(DC|AC)\s+(.+)$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        volt = eslesme.group(2)
        akim_tipi = eslesme.group(3)
        renk_raw = klemsan_model_norm(eslesme.group(4))
        renk = "GRI" if renk_raw in ["G", "GRI"] else renk_parantez.get(renk_raw)
        if renk:
            ekle(f"PYK {olcu}SLD {volt}V {akim_tipi} (5x20) {renk}")

    # PYK-4I 6,35x31,75 G -> PYK 4I (6,35x31,75) GRI
    # PYK-4I LD 24V DC G  -> PYK 4ILD 24V DC (6,35x31,75) (GRI)
    eslesme = re.match(r"^PYK-([0-9,]+)I\s+([0-9,]+X[0-9,]+)\s+G$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        sigorta = eslesme.group(2).replace("X", "x")
        ekle(f"PYK {olcu}I ({sigorta}) GRI")

    eslesme = re.match(r"^PYK-([0-9,]+)I\s+LD\s+([0-9]+)V\s+(DC|AC)\s+G$", raw)
    if eslesme:
        olcu = eslesme.group(1)
        volt = eslesme.group(2)
        akim = eslesme.group(3)
        ekle(f"PYK {olcu}ILD {volt}V {akim} (6,35x31,75) (GRI)")

    return adaylar


KLEMSAN_MANUEL_VARYANT_MODELLERI = {
    "PISO-DC-1": "PISO-DC-1 için 0-20mA/0-20mA, 4-20mA/4-20mA, 0-20mA/0-10V, 0-20mA/0-5V varyantları var.",
    "PISO-DC-2": "PISO-DC-2 için 0-20mA/0-20mA, 4-20mA/4-20mA, 0-20mA/0-10V, 0-20mA/0-5V varyantları var.",
    "PISO-DC-DUO": "PISO-DC-DUO için 0-20mA/0-20mA, 4-20mA/4-20mA, 0-20mA/0-10V, 0-20mA/0-5V varyantları var.",
    "TT-RTD-LP": "TT-RTD-LP için farklı sıcaklık aralığı varyantları var.",
    "TERMAL YAZICI RIBON": "KLEMARK ribon için 30/40/50/110mm gibi birden fazla ölçü var; ölçü detayı eksik.",
    "TERMAL YAZICI SERIT": "Termal yazıcı şerit için net tek karşılık bulunamadı; ölçü/model detayı eksik.",
    "YAPISKAN ETIKET GUMUS": "Yapışkan etiket için ölçü/model detayı eksik.",
    "YAPISKAN ETIKET SARI": "Yapışkan etiket için ölçü/model detayı eksik.",
    "BUTON ETIKETI GUMUS": "Buton etiketi için model/ölçü detayı eksik.",
    "PANO YUZEY ETIKETI": "Pano yüzey etiketi için model/ölçü detayı eksik.",
}


KLEMSAN_KARSILIGI_BULUNMAYAN_MODELLERI = {
    "AVK-2,5 A BEJ": "KLEMSAN fiyat listesinde AVK 2,5A için BEJ varyantı net bulunamadı.",
    "P1-SU": "KLEMSAN fiyat listesinde P1-SU için net tek karşılık bulunamadı.",
    "CT3N-AC": "KLEMSAN fiyat listesinde CT3N-AC için net tek karşılık bulunamadı.",
    "CT3N-AC-24": "KLEMSAN fiyat listesinde CT3N-AC-24 için net tek karşılık bulunamadı.",
    "QS40.361": "KLEMSAN fiyat listesinde QS40.361 için net tek karşılık bulunamadı.",
}


SANTRAL_KLEMSAN_DOGRUDAN_KOD_ESLESME = {
    "TRV-4 S": "422027",
    "TRV-10": "422037",
    "TRV-4 M": "422047",
    "TRV-4 B": "422057",
    "DARALAN MAKARON APARATI": "1020050",
    "RA-1": "250107",
    "PP-1": "250117",
    "PK-1": "250207",
    "YKB-450 B 8 X 450": "538450B",
    "YKB-762 B 9 X 762": "539762B",
    "YKB-450 S 8 X 450": "538450S",
    "YKB-762 S 9 X 762": "539762S",
    "KKD-4060 PVC KANAL": "5554060",
    "KKD-4060 GENIS OLUK": "5554060WS",
    "KKD-6080 PVC KANAL": "5556080",
    "KKD-6080 GENIS OLUK": "5556080WS",
}


def santral_klemsan_dogrudan_kod_bul(model):
    return SANTRAL_KLEMSAN_DOGRUDAN_KOD_ESLESME.get(
        klemsan_model_norm(model),
        "",
    )


def santral_klemsan_kod_duzelt(df, kod_kolonu, kaynak_dosya, kontrol_mask=None):
    """
    SANTRAL + KLEMSAN için çalışır.
    SANTRAL'den gelen ürün adını KLEMSAN fiyat listesindeki model adına çevirir,
    fiyat listesindeki Noktasız Stok Kodu değerini alıp eşleşmede kullanılacak koda yazar.
    """
    fiyat_dosyasi = klemsan_fiyat_listesi_bul()

    orijinal_kolon = "SANTRAL Orijinal Ürün Kodu"
    durum_kolon = "SANTRAL Kod Düzeltme Durumu"
    not_kolon = "SANTRAL Kod Düzeltme Notu"

    df = df.copy()

    if orijinal_kolon not in df.columns:
        df[orijinal_kolon] = df[kod_kolonu]
    if durum_kolon not in df.columns:
        df[durum_kolon] = ""
    if not_kolon not in df.columns:
        df[not_kolon] = ""

    if kontrol_mask is None:
        kontrol_mask = pd.Series([True] * len(df), index=df.index)
    else:
        kontrol_mask = pd.Series(kontrol_mask, index=df.index).fillna(False).astype(bool)

    if fiyat_dosyasi is None:
        print("SANTRAL KLEMSAN düzeltmesi için KLEMSAN fiyat listesi bulunamadı.")
        print("Fiyat listesini şu klasöre koy:")
        print(FIYAT_LISTESI_DIR)
        print("Dosya adında KLEMSAN geçsin.")
        raise SystemExit

    print()
    print("SANTRAL KLEMSAN model -> noktasız stok kodu düzeltmesi kontrol ediliyor.")
    print("KLEMSAN fiyat listesi:")
    print(fiyat_dosyasi)

    klemsan_df = klemsan_fiyat_listesi_oku(fiyat_dosyasi)

    model_kolonu = kolon_bul_esnek(
        klemsan_df,
        [
            "Model",
            "Model Adı",
            "Model Adi",
            "Ürün Modeli",
            "Urun Modeli",
            "Ürün Adı",
            "Urun Adi",
            "Malzeme Adı",
            "Malzeme Adi",
            "Stok Adı",
            "Stok Adi",
        ]
    )

    noktasiz_kod_kolonu = kolon_bul_esnek(
        klemsan_df,
        [
            "Noktasız Stok Kodu",
            "Noktasiz Stok Kodu",
            "Noktasız Kod",
            "Noktasiz Kod",
            "Noktasız Ürün Kodu",
            "Noktasiz Urun Kodu",
            "Web Servis Kodu",
            "WebServis Kodu",
        ]
    )

    if not model_kolonu or not noktasiz_kod_kolonu:
        print("KLEMSAN fiyat listesinde Model veya Noktasız Stok Kodu kolonu bulunamadı.")
        print("Gelen kolonlar:")
        for c in klemsan_df.columns:
            print("-", c)
        raise SystemExit

    model_lookup = {}

    for _, row in klemsan_df.iterrows():
        model_raw = str(row.get(model_kolonu, "")).strip()
        noktasiz_kod = str(row.get(noktasiz_kod_kolonu, "")).strip()

        if not model_raw or not noktasiz_kod:
            continue

        for anahtar in klemsan_model_anahtarlari_uret(model_raw):
            model_lookup.setdefault(anahtar, set()).add((noktasiz_kod, model_raw))

    duzeltme_satirlari = []

    for idx, row in df[kontrol_mask].iterrows():
        eski_deger = str(row.get(kod_kolonu, "")).strip()
        eski_deger_norm = klemsan_model_norm(eski_deger)

        dogrudan_kod = santral_klemsan_dogrudan_kod_bul(eski_deger)
        if dogrudan_kod:
            df.at[idx, kod_kolonu] = dogrudan_kod
            df.at[idx, durum_kolon] = "KLEMSAN doğrudan noktasız stok koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger} -> {dogrudan_kod}"
            duzeltme_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger,
                "KLEMSAN Model Adayı": eski_deger,
                "Yeni Noktasız Stok Kodu": dogrudan_kod,
                "Durum": "Doğrudan çevrildi",
            })
            continue

        if eski_deger_norm in KLEMSAN_KARSILIGI_BULUNMAYAN_MODELLERI:
            df.at[idx, durum_kolon] = "KLEMSAN model karşılığı bulunamadı"
            df.at[idx, not_kolon] = KLEMSAN_KARSILIGI_BULUNMAYAN_MODELLERI[eski_deger_norm]
            continue

        if eski_deger_norm in KLEMSAN_MANUEL_VARYANT_MODELLERI:
            df.at[idx, durum_kolon] = "KLEMSAN çoklu stok kodu"
            df.at[idx, not_kolon] = (
                f"{eski_deger} fiyat listesinde birden fazla varyanta karşılık geliyor; "
                f"{KLEMSAN_MANUEL_VARYANT_MODELLERI[eski_deger_norm]} Stok kodu otomatik değiştirilmedi."
            )
            continue

        wg_eki = re.match(r"^WG-EKI-N\s+([0-9]+N)$", eski_deger_norm)
        if wg_eki:
            yeni_kod = wg_eki.group(1)
            df.at[idx, kod_kolonu] = yeni_kod
            df.at[idx, durum_kolon] = "KLEMSAN modelden noktasız stok koduna çevrildi"
            df.at[idx, not_kolon] = f"{eski_deger} -> WG-EKI-C -> {yeni_kod}"
            duzeltme_satirlari.append({
                "Kaynak Dosya": str(kaynak_dosya),
                "Excel Satır No": int(idx) + 2,
                "Eski Değer": eski_deger,
                "KLEMSAN Model Adayı": "WG-EKI-C",
                "Yeni Noktasız Stok Kodu": yeni_kod,
                "Durum": "Çevrildi",
            })
            continue

        adaylar = santral_klemsan_model_adaylari_uret(eski_deger)

        bulunanlar = []
        kullanilan_model = ""

        for aday in adaylar:
            birebir_anahtar = klemsan_model_norm(aday)
            birebir_eslesenler = model_lookup.get(birebir_anahtar, set())
            birebir_kodlar = sorted({kod for kod, _ in birebir_eslesenler if kod})
            if len(birebir_kodlar) == 1:
                bulunanlar = sorted(birebir_eslesenler)
                kullanilan_model = aday
                break

            for anahtar in klemsan_model_anahtarlari_uret(aday):
                eslesenler = model_lookup.get(anahtar, set())
                if eslesenler:
                    bulunanlar = sorted(eslesenler)
                    kullanilan_model = aday
                    break

            if bulunanlar:
                break

        if not bulunanlar:
            df.at[idx, durum_kolon] = "KLEMSAN model karşılığı bulunamadı"
            df.at[idx, not_kolon] = "KLEMSAN fiyat listesinde model adı bulunamadı; stok koduna çevrilmedi."
            continue

        noktasiz_kodlar = sorted({kod for kod, _ in bulunanlar if kod})

        # SANTRAL IKY satırlarında ECO/F/Z gibi ek varyant yazmıyorsa normal ürün seçilir.
        # KLEMSAN fiyat listesinde aynı model IKY 0,50/8, IKY 0,50/8 F, IKY 0,50/8 Z
        # gibi aynı adla dönebildiği için harfsiz noktasız stok kodunu güvenli önceliklendiririz.
        if len(noktasiz_kodlar) != 1 and re.match(r"^İ?IKY-|^IKY-", eski_deger_norm):
            harfsiz_kodlar = [kod for kod in noktasiz_kodlar if re.fullmatch(r"[0-9]+", str(kod))]
            if len(harfsiz_kodlar) == 1:
                noktasiz_kodlar = harfsiz_kodlar

        if len(noktasiz_kodlar) != 1:
            df.at[idx, durum_kolon] = "KLEMSAN çoklu stok kodu"
            df.at[idx, not_kolon] = f"{kullanilan_model} için birden fazla noktasız stok kodu bulundu: {', '.join(noktasiz_kodlar)}"
            continue

        yeni_kod = noktasiz_kodlar[0]

        df.at[idx, kod_kolonu] = yeni_kod
        df.at[idx, durum_kolon] = "KLEMSAN modelden noktasız stok koduna çevrildi"
        df.at[idx, not_kolon] = f"{eski_deger} -> {kullanilan_model} -> {yeni_kod}"

        duzeltme_satirlari.append({
            "Kaynak Dosya": str(kaynak_dosya),
            "Excel Satır No": int(idx) + 2,
            "Eski Değer": eski_deger,
            "KLEMSAN Model Adayı": kullanilan_model,
            "Yeni Noktasız Stok Kodu": yeni_kod,
            "Durum": "Çevrildi",
        })

    rapor_df = pd.DataFrame(duzeltme_satirlari)
    if rapor_df.empty:
        rapor_df = pd.DataFrame(columns=[
            "Kaynak Dosya",
            "Excel Satır No",
            "Eski Değer",
            "KLEMSAN Model Adayı",
            "Yeni Noktasız Stok Kodu",
            "Durum",
        ])

    return df


def b2b_tedarikci_oku(ayar, santral_lookup):
    tedarikci = str(ayar["Tedarikci"]).strip()
    marka_filtresi = str(ayar["MarkaFiltresi"]).strip()
    kod_kolonu = str(ayar["KodKolonu"]).strip()
    stok_kolonu = str(ayar["StokKolonu"]).strip()
    marka_kolonu = str(ayar["MarkaKolonu"]).strip()
    eslesme_tipi = str(ayar.get("EslesmeTipi", "WEBSERVIS")).strip().upper()

    tedarikci_norm = metin_norm(tedarikci)
    marka_norm = metin_norm(marka_filtresi)

    sabit_tedarikci_markasi = SABIT_TEDARIKCI_MARKALARI.get(tedarikci_norm, "")
    if sabit_tedarikci_markasi:
        marka_filtresi = sabit_tedarikci_markasi
        marka_norm = metin_norm(sabit_tedarikci_markasi)
        marka_kolonu = "__Sabit Marka__"
        eslesme_tipi = "WEBSERVIS"

    dosya = tedarikci_dosyasi_bul(tedarikci)

    if dosya is None:
        print(f"{tedarikci} için dosya bulunamadı.")
        return pd.DataFrame()

    print()
    print(f"{tedarikci} dosyası okunuyor:")
    print(dosya)
    print(f"Marka filtresi: {marka_filtresi}")

    df = dosya_oku(dosya)
    df = tamamen_bos_satirlari_sil(df)

    sabit_marka_kolonu = (
        bool(sabit_tedarikci_markasi)
        or metin_norm(marka_kolonu) in [
            "SABIT",
            "SABIT MARKA",
            "__SABIT_MARKA__",
            "__MARKA_FILTRESI__",
        ]
    )

    if sabit_marka_kolonu:
        marka_kolonu = "__Sabit Marka__"
        df[marka_kolonu] = marka_filtresi

    gerekli = [kod_kolonu, stok_kolonu, marka_kolonu]

    eksikler = [kolon for kolon in gerekli if kolon not in df.columns]

    if eksikler:
        print(f"{tedarikci} dosyasında eksik kolon var:")
        for kolon in eksikler:
            print("-", kolon)

        print("Gelen kolonlar:")
        for c in df.columns:
            print("-", c)

        return pd.DataFrame()

    marka_kontrol_mask = df[marka_kolonu].apply(
        lambda x: marka_uygun_mu(x, marka_filtresi)
    )

    kirac_entes_dogrudan_sade_kodlar = set()
    if "KIRAC" in tedarikci_norm and marka_norm == "ENTES":
        kirac_entes_dogrudan_sade_kodlar = set(
            df.loc[marka_kontrol_mask, kod_kolonu]
            .astype(str)
            .apply(sade_kod)
            .dropna()
        )

    # KIRAC özel ENTES model -> ürün kodu düzeltmesi
    # SADECE KIRAC + ENTES için çalışır.
    # KIRAC'ın Legrand dahil diğer markalarında model düzeltme / ENTES fiyat listesi kontrolü yapılmaz.
    if "KIRAC" in tedarikci_norm and marka_norm == "ENTES":
        df = kirac_model_kod_duzelt(
            df=df,
            kod_kolonu=kod_kolonu,
            kaynak_dosya=dosya,
            kontrol_mask=marka_kontrol_mask
        )

    # SANTRAL özel ENTES model -> ürün kodu düzeltmesi
    # SADECE SANTRAL + ENTES için çalışır.
    # SANTRAL'in diğer markalarında direkt WEBSERVIS eşleşme yapılır.
    if "SANTRAL" in tedarikci_norm and marka_norm == "ENTES":
        df = santral_model_kod_duzelt(
            df=df,
            kod_kolonu=kod_kolonu,
            kaynak_dosya=dosya,
            kontrol_mask=marka_kontrol_mask
        )

        # SANTRAL ENTES düzeltme sonrası ürün kodu doğrudan
        # T-Soft Web Servis Kodu gibi eşleştirilir.
        eslesme_tipi = "WEBSERVIS"

    # SANTRAL özel KLEMSAN model -> noktasız stok kodu düzeltmesi
    # SANTRAL'den gelen KLEMSAN ürün adı önce KLEMSAN fiyat listesindeki model adına çevrilir.
    # Fiyat listesindeki Noktasız Stok Kodu, T-Soft Web Servis Kodu ile eşleşmede kullanılır.
    if "SANTRAL" in tedarikci_norm and marka_norm == "KLEMSAN":
        df = santral_klemsan_kod_duzelt(
            df=df,
            kod_kolonu=kod_kolonu,
            kaynak_dosya=dosya,
            kontrol_mask=marka_kontrol_mask
        )

        eslesme_tipi = "WEBSERVIS"

    filtre = df[marka_kolonu].apply(
        lambda x: marka_uygun_mu(x, marka_filtresi)
    )

    marka_df = df[filtre].copy()

    sonuc = pd.DataFrame()
    sonuc["Tedarikci"] = [tedarikci] * len(marka_df)
    sonuc["EslesmeTipi"] = [eslesme_tipi] * len(marka_df)
    sonuc["B2B Satır No"] = (marka_df.index + 2).astype(int).values
    sonuc["B2B Kaynak Dosya"] = str(dosya)
    sonuc["B2B Web Servis Kodu"] = marka_df[kod_kolonu].astype(str).str.strip().values
    sonuc["B2B Stok Metni"] = marka_df[stok_kolonu].astype(str).str.strip().values
    sonuc["B2B Marka"] = marka_df[marka_kolonu].apply(marka_adi_standart).values

    if "EM Kodu" in marka_df.columns:
        sonuc["EM Kodu"] = marka_df["EM Kodu"].astype(str).str.strip().values
    else:
        sonuc["EM Kodu"] = ""

    for kirac_kolon in [
        "KIRAC Orijinal Ürün Kodu",
        "KIRAC Kod Düzeltme Durumu",
        "KIRAC Kod Düzeltme Notu",
    ]:
        if kirac_kolon in marka_df.columns:
            sonuc[kirac_kolon] = marka_df[kirac_kolon].astype(str).str.strip().values
        else:
            sonuc[kirac_kolon] = ""

    for santral_kolon in [
        "SANTRAL Orijinal Ürün Kodu",
        "SANTRAL Kod Düzeltme Durumu",
        "SANTRAL Kod Düzeltme Notu",
    ]:
        if santral_kolon in marka_df.columns:
            sonuc[santral_kolon] = marka_df[santral_kolon].astype(str).str.strip().values
        else:
            sonuc[santral_kolon] = ""

    sonuc["B2B Temiz Kod"] = sonuc["B2B Web Servis Kodu"].apply(b2b_temiz_kod_uret)
    sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)
    sonuc["Yeni Stok"] = sonuc["B2B Stok Metni"].apply(stok_degerine_cevir)

    # Excel raporlarinda ham "STOKTA YOK" / "48 ADET" gibi metinler kalmasin.
    # B2B Stok Metni kolonu da temizlenmis sayisal stok degerini gostersin.
    sonuc["B2B Stok Metni"] = sonuc["Yeni Stok"]

    if "KIRAC" in tedarikci_norm and marka_norm == "ENTES":
        duzeltme_durumu = sonuc["KIRAC Kod Düzeltme Durumu"].astype(str).apply(metin_norm)
        modelden_cevrilen = duzeltme_durumu.str.contains("CEVRILDI", na=False)
        hedef_m_kodu = sonuc["B2B Sade Kod"].astype(str).str.startswith("M", na=False)
        dogrudan_kodu_da_var = sonuc["B2B Sade Kod"].isin(
            kirac_entes_dogrudan_sade_kodlar
        )
        elenecek_mask = modelden_cevrilen & hedef_m_kodu & dogrudan_kodu_da_var
        elenecekler = sonuc[elenecek_mask].copy()

        if not elenecekler.empty:
            sonuc = sonuc[~elenecek_mask].copy()

            print()
            print("KIRAC + ENTES ayni kod dogrudan da geldigi icin mukerrer satirlar cikarildi.")
            print(f"Mukerrer cikarilan satir: {len(elenecekler)}")

    sonuc["Manuel Tsoft Web Servis Kodu"] = ""
    sonuc["Eslesmede Kullanilacak Kod"] = sonuc["B2B Web Servis Kodu"]

    if eslesme_tipi == "MANUEL_ESLESME":
        sonuc["Manuel Tsoft Web Servis Kodu"] = sonuc["B2B Web Servis Kodu"].apply(
            lambda x: manuel_tsoft_kodu_bul(x, santral_lookup)
        )

        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Manuel Tsoft Web Servis Kodu"]

    # ACK ozel kurali: B2B'de "ACK AY02-01500" gelen kodu
    # T-Soft eslesmesi icin "AY02-01500" bicimine cevirir.
    if tedarikci_norm == "KONYA" and marka_norm == "ACK":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            ack_prefix_sil
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # BEMIS ozel kurali: B2B'de "BEM BAK-1000-0010" gelen kodu
    # T-Soft eslesmesi icin "BAK-1000-0010" bicimine cevirir.
    if tedarikci_norm == "KONYA" and marka_norm == "BEMIS":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            bemis_prefix_sil
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # BYLION ozel kurali: Once basindaki BYLION marka adini siler,
    # sonra bilinen model adini dogrudan T-Soft urun koduna cevirir.
    if tedarikci_norm == "KONYA" and marka_norm == "BYLION":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            bylion_model_urun_koduna_cevir
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # CATA ozel genel kural: "CATA CT-4221-G" kodunu "CT-4221G" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "CATA":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            cata_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # CETINKAYA ozel kurali: Yalnizca bastaki "CETIN." marka on ekini siler.
    if tedarikci_norm == "KONYA" and marka_norm == "CETINKAYA":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            cetinkaya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # EMAS ozel kurali: "EMAS B100HY" kodunu "B100HY" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "EMAS":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            emas_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KARDES ozel kurali: "U.KRD KP050-M08-EURO" kodunu "KP050-M08" yapar.
    if tedarikci_norm == "KONYA" and marka_norm in {"KARDES", "KARDES ELEKTRIK"}:
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            kardes_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KLEMSAN ozel kurali: "KLM 610208-EURO" kodunu "610208" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "KLEMSAN":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            klemsan_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KRAUS NAIMER ozel kurali: "U.KRAU C26 A200" kodunu "C26 A200" yapar.
    if tedarikci_norm == "KONYA" and marka_norm in {"KRAUS NAIMER", "KRAUS & NAIMER"}:
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            kraus_naimer_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + LEGRAND genel kural: "LEGR. 26092" -> "26092";
    # bastaki sifirlar da silinir: "LEGR. 003934" -> "3934".
    if tedarikci_norm == "KONYA" and marka_norm == "LEGRAND":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            konya_legrand_urun_koduna_cevir
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + METE ozel kurali: "METE 141044-EURO" kodunu "141044" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "METE":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            mete_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + MUTLUSAN ozel kurali: Bastaki "MUT." on ekini siler.
    if tedarikci_norm == "KONYA" and marka_norm == "MUTLUSAN":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            mutlusan_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + NADE ozel kurali: Bastaki "NADE" marka on ekini siler.
    if tedarikci_norm == "KONYA" and marka_norm == "NADE":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            nade_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + NEAD ozel kurali: Bastaki "NEAD" marka on ekini siler.
    if tedarikci_norm == "KONYA" and marka_norm == "NEAD":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            nead_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + OMRON ozel kurali: Bastaki "OMR." marka on ekini siler.
    if tedarikci_norm == "KONYA" and marka_norm == "OMRON":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            omron_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + OPAŞ: "OPAS. IPE3132" -> "IPE3132". T-Soft'taki
    # "IPE 3132" ile sade kod uzerinden eslesir. Yardimci kontakli
    # "IPE 3 1A1K32" ayri urundur ve bu kural ona donusum yapmaz.
    if tedarikci_norm == "KONYA" and marka_norm == "OPAS":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            opas_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + ORTAC ozel kurali: "ORT. ADGK13-EURO" kodunu "ADGK13" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "ORTAC":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            ortac_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + OSRAM ozel kurali: "AMP.OS 0086" kodunu "OS 0086" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "OSRAM":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            osram_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + PHILIPS ozel kurali: "AMP.PH 0025" kodunu "PH 0025" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "PHILIPS":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            philips_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + PHOENIX CONTACT ozel kurali: "PHOE 0719087" kodunu "0719087" yapar.
    if tedarikci_norm == "KONYA" and marka_norm == "PHOENIX CONTACT":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            phoenix_contact_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + SCHNEIDER: S.TM, S.ETKS(+EPH), S.HC ve S.MG on eklerini siler.
    if tedarikci_norm == "KONYA" and marka_norm == "SCHNEIDER":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            schneider_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + SIEMENS ozel kurali: "SIE. 3KA7114-4AA00-EURO" -> "3KA7114-4AA00".
    if tedarikci_norm == "KONYA" and marka_norm == "SIEMENS":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            siemens_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + SAFAK ozel kurali: "SFK. AL120-CU95-DOLAR" -> "AL120-CU95".
    if tedarikci_norm == "KONYA" and marka_norm in {"SAFAK", "SAFAK ELEKTRIK"}:
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            safak_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + TELEMECANIQUE: "TM ZCKJ4046H29-EURO" -> "ZCKJ4046H29".
    if tedarikci_norm == "KONYA" and marka_norm in {
        "TELEMECANIQUE",
        "TELEMECANIQUE SENSORS",
    }:
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            telemecanique_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + VIKO/PANASONIC: "VIKO 90110600" -> "90110600".
    if tedarikci_norm == "KONYA" and marka_norm in {
        "VIKO",
        "VIKO-PANASONIC",
        "VIKO PANASONIC",
    }:
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            viko_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + WAGO: "WAGO 2002-1611/1000-541EURO" -> "2002-1611/1000-541".
    if tedarikci_norm == "KONYA" and marka_norm == "WAGO":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            wago_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + WEIDMULLER: "WDM 0473460001-EURO" -> "0473460001".
    if tedarikci_norm == "KONYA" and marka_norm == "WEIDMULLER":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            weidmuller_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KONYA + WOER: "WOER 067-DOLAR" -> "WOER-W-67".
    if tedarikci_norm == "KONYA" and marka_norm == "WOER":
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            woer_konya_urun_kodu_temizle
        )
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

    # KIRAC + LEGRAND özel kuralı:
    # KIRAC'ta Legrand kodları 004280 gibi başında sıfırla gelebiliyor.
    # T-Soft'ta karşılığı 4280 olduğu için sadece KIRAC + LEGRAND satırlarında
    # baştaki sıfırları eşleşmede kullanılacak koddan siliyoruz.
    if "KIRAC" in tedarikci_norm and marka_norm == "LEGRAND":
        sonuc["KIRAC Legrand Orijinal Kod"] = sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()

        sonuc["KIRAC Legrand Sıfırsız Kod"] = sonuc["KIRAC Legrand Orijinal Kod"].apply(
            bas_sifir_sil
        )

        sonuc["KIRAC Legrand Sıfır Kuralı"] = sonuc.apply(
            lambda row: (
                f"{row['KIRAC Legrand Orijinal Kod']} -> {row['KIRAC Legrand Sıfırsız Kod']}"
                if str(row["KIRAC Legrand Orijinal Kod"]).strip() != str(row["KIRAC Legrand Sıfırsız Kod"]).strip()
                else ""
            ),
            axis=1
        )

        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["KIRAC Legrand Sıfırsız Kod"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)
    else:
        sonuc["KIRAC Legrand Orijinal Kod"] = ""
        sonuc["KIRAC Legrand Sıfırsız Kod"] = ""
        sonuc["KIRAC Legrand Sıfır Kuralı"] = ""

    # KIRAC + BOSCH özel kuralı:
    # KIRAC B2B'de Bosch kodu 06019G8100 gibi bosluksuz gelebiliyor.
    # T-Soft Web Servis Kodu 0 601 9G8 100 formatinda oldugu icin
    # sadece KIRAC + BOSCH satirlarinda 1-3-3-3 Bosch bosluk formatina ceviriyoruz.
    if "KIRAC" in tedarikci_norm and marka_norm == "BOSCH":
        sonuc["KIRAC Bosch Orijinal Kod"] = sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()

        sonuc["KIRAC Bosch Boşluklu Kod"] = sonuc["KIRAC Bosch Orijinal Kod"].apply(
            bosch_bosluklu_kod
        )

        sonuc["KIRAC Bosch Boşluk Kuralı"] = sonuc.apply(
            lambda row: (
                f"{row['KIRAC Bosch Orijinal Kod']} -> {row['KIRAC Bosch Boşluklu Kod']}"
                if str(row["KIRAC Bosch Orijinal Kod"]).strip() != str(row["KIRAC Bosch Boşluklu Kod"]).strip()
                else ""
            ),
            axis=1
        )

        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["KIRAC Bosch Boşluklu Kod"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)
    else:
        sonuc["KIRAC Bosch Orijinal Kod"] = ""
        sonuc["KIRAC Bosch Boşluklu Kod"] = ""
        sonuc["KIRAC Bosch Boşluk Kuralı"] = ""

    # KIRAC + ORTAC özel kuralı:
    # KIRAC B2B'de OSPA11 / OSPA37 gibi bosluksuz gelebiliyor.
    # T-Soft Web Servis Kodu OSPA 11 / OSPA 37 formatinda oldugu icin
    # sadece KIRAC + ORTAC satirlarinda OSPA ile sayi arasina bosluk ekliyoruz.
    if "KIRAC" in tedarikci_norm and marka_norm == "ORTAC":
        sonuc["KIRAC Ortac Orijinal Kod"] = sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()

        sonuc["KIRAC Ortac Boşluklu Kod"] = sonuc["KIRAC Ortac Orijinal Kod"].apply(
            ortac_ospa_bosluklu_kod
        )

        sonuc["KIRAC Ortac Boşluk Kuralı"] = sonuc.apply(
            lambda row: (
                f"{row['KIRAC Ortac Orijinal Kod']} -> {row['KIRAC Ortac Boşluklu Kod']}"
                if str(row["KIRAC Ortac Orijinal Kod"]).strip() != str(row["KIRAC Ortac Boşluklu Kod"]).strip()
                else ""
            ),
            axis=1
        )

        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["KIRAC Ortac Boşluklu Kod"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)
    else:
        sonuc["KIRAC Ortac Orijinal Kod"] = ""
        sonuc["KIRAC Ortac Boşluklu Kod"] = ""
        sonuc["KIRAC Ortac Boşluk Kuralı"] = ""

    # KIRAC + SAFAK ELEKTRIK özel kuralı:
    # KIRAC B2B'de KM-240 gibi tireli gelen kodlar T-Soft'ta KM 240
    # biçiminde tutulduğu için yalnızca bu marka ve tedarikçide boşluklu forma çevrilir.
    if "KIRAC" in tedarikci_norm and marka_norm == "SAFAK ELEKTRIK":
        sonuc["KIRAC Şafak Orijinal Kod"] = sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()

        sonuc["KIRAC Şafak Boşluklu Kod"] = sonuc["KIRAC Şafak Orijinal Kod"].apply(
            safak_km_bosluklu_kod
        )

        sonuc["KIRAC Şafak KM Kuralı"] = sonuc.apply(
            lambda row: (
                f"{row['KIRAC Şafak Orijinal Kod']} -> {row['KIRAC Şafak Boşluklu Kod']}"
                if str(row["KIRAC Şafak Orijinal Kod"]).strip() != str(row["KIRAC Şafak Boşluklu Kod"]).strip()
                else ""
            ),
            axis=1
        )

        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["KIRAC Şafak Boşluklu Kod"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(b2b_temiz_kod_uret)
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)
    else:
        sonuc["KIRAC Şafak Orijinal Kod"] = ""
        sonuc["KIRAC Şafak Boşluklu Kod"] = ""
        sonuc["KIRAC Şafak KM Kuralı"] = ""

    # KIRAC + SCHNEIDER ELECTRIC tekrar kuralı:
    # Aynı web servis kodu KIRAC dosyasında birden fazla kez gelirse stoklar toplanmaz.
    # Stok değeri en yüksek olan satır tutulur. Örnek: LADN22 0 ve 22 -> 22.
    # KIRAC + OZNUR KABLO marka bazli kural:
    # KIRAC'in kisa kablo kodu; aile, damar/kesit ve renk bilgisine gore
    # T-Soft Web Servis Kodu bicimine cevrilir. Bu blok KIRAC'in diger
    # markalarina ve diger tedarikcilerin OZNUR satirlarina dokunmaz.
    if "KIRAC" in tedarikci_norm and marka_norm in {"OZNUR", "OZNUR KABLO"}:
        sonuc["KIRAC OZNUR Orijinal Kod"] = (
            sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()
        )
        sonuc["KIRAC OZNUR Tsoft Kodu"] = sonuc[
            "KIRAC OZNUR Orijinal Kod"
        ].apply(kirac_oznurkablo_tsoft_koduna_cevir)
        sonuc["KIRAC OZNUR Kural"] = sonuc.apply(
            lambda row: (
                f"{row['KIRAC OZNUR Orijinal Kod']} -> {row['KIRAC OZNUR Tsoft Kodu']}"
                if str(row["KIRAC OZNUR Tsoft Kodu"]).strip()
                else f"{row['KIRAC OZNUR Orijinal Kod']} -> GUVENLI HEDEF BULUNAMADI"
            ),
            axis=1,
        )
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["KIRAC OZNUR Tsoft Kodu"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            b2b_temiz_kod_uret
        )
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

        # MK/PRC gibi farkli kaynak satirlari ayni teknik urune donusurse
        # stoklari iki kez toplama; ayni hedefte en yuksek stoku kullan.
        sonuc["_KIRAC_OZNUR_TEKRAR_ANAHTAR"] = sonuc[
            "Eslesmede Kullanilacak Kod"
        ].apply(kod_temizle)
        sonuc["_KIRAC_OZNUR_STOK_SECIM"] = pd.to_numeric(
            sonuc["Yeni Stok"], errors="coerce"
        ).fillna(-1)
        anahtar_var = sonuc["_KIRAC_OZNUR_TEKRAR_ANAHTAR"].astype(str).str.strip().ne("")
        kodlu = sonuc[anahtar_var].copy()
        kodsuz = sonuc[~anahtar_var].copy()
        kodlu = kodlu.sort_values(
            by=[
                "_KIRAC_OZNUR_TEKRAR_ANAHTAR",
                "_KIRAC_OZNUR_STOK_SECIM",
                "B2B Satır No",
            ],
            ascending=[True, False, True],
            kind="stable",
        ).drop_duplicates(
            subset=["_KIRAC_OZNUR_TEKRAR_ANAHTAR"],
            keep="first",
        )
        sonuc = pd.concat([kodlu, kodsuz], ignore_index=True, sort=False)
        sonuc = sonuc.sort_values("B2B Satır No", kind="stable").reset_index(drop=True)
        sonuc = sonuc.drop(
            columns=["_KIRAC_OZNUR_TEKRAR_ANAHTAR", "_KIRAC_OZNUR_STOK_SECIM"],
            errors="ignore",
        )
    else:
        sonuc["KIRAC OZNUR Orijinal Kod"] = ""
        sonuc["KIRAC OZNUR Tsoft Kodu"] = ""
        sonuc["KIRAC OZNUR Kural"] = ""

    # SANTRAL + OZNUR KABLO marka bazli kural:
    # Urun adindaki aile, damar/kesit ve renk bilgisiyle T-Soft kodu uretilir.
    # Kural SANTRAL'in diger markalarina ve diger tedarikcilere dokunmaz.
    if "SANTRAL" in tedarikci_norm and marka_norm in {"OZNUR", "OZNUR KABLO"}:
        sonuc["SANTRAL OZNUR Orijinal Kod"] = (
            sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()
        )
        sonuc["SANTRAL OZNUR Tsoft Kodu"] = sonuc[
            "SANTRAL OZNUR Orijinal Kod"
        ].apply(santral_oznurkablo_tsoft_koduna_cevir)
        sonuc["SANTRAL OZNUR Kural"] = sonuc.apply(
            lambda row: (
                f"{row['SANTRAL OZNUR Orijinal Kod']} -> {row['SANTRAL OZNUR Tsoft Kodu']}"
                if str(row["SANTRAL OZNUR Tsoft Kodu"]).strip()
                else f"{row['SANTRAL OZNUR Orijinal Kod']} -> GUVENLI HEDEF BULUNAMADI"
            ),
            axis=1,
        )
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["SANTRAL OZNUR Tsoft Kodu"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            b2b_temiz_kod_uret
        )
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)

        # Ayni teknik hedef birden fazla SANTRAL satirindan gelirse iki kez
        # toplama; en yuksek stoklu satiri kullan.
        sonuc["_SANTRAL_OZNUR_TEKRAR_ANAHTAR"] = sonuc[
            "Eslesmede Kullanilacak Kod"
        ].apply(kod_temizle)
        sonuc["_SANTRAL_OZNUR_STOK_SECIM"] = pd.to_numeric(
            sonuc["Yeni Stok"], errors="coerce"
        ).fillna(-1)
        anahtar_var = sonuc["_SANTRAL_OZNUR_TEKRAR_ANAHTAR"].astype(str).str.strip().ne("")
        kodlu = sonuc[anahtar_var].copy()
        kodsuz = sonuc[~anahtar_var].copy()
        kodlu = kodlu.sort_values(
            by=[
                "_SANTRAL_OZNUR_TEKRAR_ANAHTAR",
                "_SANTRAL_OZNUR_STOK_SECIM",
                "B2B Satır No",
            ],
            ascending=[True, False, True],
            kind="stable",
        ).drop_duplicates(
            subset=["_SANTRAL_OZNUR_TEKRAR_ANAHTAR"],
            keep="first",
        )
        sonuc = pd.concat([kodlu, kodsuz], ignore_index=True, sort=False)
        sonuc = sonuc.sort_values("B2B Satır No", kind="stable").reset_index(drop=True)
        sonuc = sonuc.drop(
            columns=["_SANTRAL_OZNUR_TEKRAR_ANAHTAR", "_SANTRAL_OZNUR_STOK_SECIM"],
            errors="ignore",
        )
    else:
        sonuc["SANTRAL OZNUR Orijinal Kod"] = ""
        sonuc["SANTRAL OZNUR Tsoft Kodu"] = ""
        sonuc["SANTRAL OZNUR Kural"] = ""

    # YILDIZ + OZNUR KABLO ozel kurali:
    # OZN.001.xxx bayi urun karti kodu, dogrulanmis teknik T-Soft koduna cevrilir.
    if "YILDIZ" in tedarikci_norm and marka_norm in {"OZNUR", "OZNUR KABLO"}:
        sonuc["YILDIZ OZNUR Orijinal Kod"] = (
            sonuc["Eslesmede Kullanilacak Kod"].astype(str).str.strip()
        )
        sonuc["YILDIZ OZNUR Tsoft Kodu"] = sonuc[
            "YILDIZ OZNUR Orijinal Kod"
        ].apply(yildiz_oznurkablo_tsoft_koduna_cevir)
        sonuc["YILDIZ OZNUR Kural"] = sonuc.apply(
            lambda row: (
                f"{row['YILDIZ OZNUR Orijinal Kod']} -> {row['YILDIZ OZNUR Tsoft Kodu']}"
                if str(row["YILDIZ OZNUR Tsoft Kodu"]).strip()
                else f"{row['YILDIZ OZNUR Orijinal Kod']} -> GUVENLI HEDEF BULUNAMADI"
            ),
            axis=1,
        )
        sonuc["Eslesmede Kullanilacak Kod"] = sonuc["YILDIZ OZNUR Tsoft Kodu"]
        sonuc["B2B Temiz Kod"] = sonuc["Eslesmede Kullanilacak Kod"].apply(
            b2b_temiz_kod_uret
        )
        sonuc["B2B Sade Kod"] = sonuc["B2B Temiz Kod"].apply(sade_kod)
    else:
        sonuc["YILDIZ OZNUR Orijinal Kod"] = ""
        sonuc["YILDIZ OZNUR Tsoft Kodu"] = ""
        sonuc["YILDIZ OZNUR Kural"] = ""

    if "KIRAC" in tedarikci_norm and "SCHNEIDER" in marka_norm and not sonuc.empty:
        sonuc = sonuc.copy()
        sonuc["_KIRAC_SCHNEIDER_TEKRAR_ANAHTAR"] = sonuc["Eslesmede Kullanilacak Kod"].apply(sade_kod)
        sonuc["_KIRAC_SCHNEIDER_STOK_SECIM"] = pd.to_numeric(
            sonuc["Yeni Stok"],
            errors="coerce"
        ).fillna(-1)

        anahtar_var = sonuc["_KIRAC_SCHNEIDER_TEKRAR_ANAHTAR"].astype(str).str.strip() != ""
        kodlu = sonuc[anahtar_var].copy()
        kodsuz = sonuc[~anahtar_var].copy()

        tekrar_mask = kodlu.duplicated(
            subset=["_KIRAC_SCHNEIDER_TEKRAR_ANAHTAR"],
            keep=False
        )
        tekrarlar = kodlu[tekrar_mask].copy()

        if not tekrarlar.empty:
            print()
            print("KIRAC + SCHNEIDER ELECTRIC tekrar eden kodlar tekillestiriliyor.")
            print("Stoklar toplanmayacak; ayni kodda en yuksek stok tutulacak.")
            print(f"Tekrar eden kaynak satir: {len(tekrarlar)}")

        # Stok azalan sırada dizildiği için aynı koddan ilk satır en yüksek stoktur.
        # Eşit stoklarda Excel'de üstte olan satır seçilir.
        kodlu = kodlu.sort_values(
            by=[
                "_KIRAC_SCHNEIDER_TEKRAR_ANAHTAR",
                "_KIRAC_SCHNEIDER_STOK_SECIM",
                "B2B Satır No",
            ],
            ascending=[True, False, True],
            kind="stable"
        )
        kodlu = kodlu.drop_duplicates(
            subset=["_KIRAC_SCHNEIDER_TEKRAR_ANAHTAR"],
            keep="first"
        )

        sonuc = pd.concat([kodlu, kodsuz], ignore_index=True, sort=False)
        sonuc = sonuc.sort_values("B2B Satır No", kind="stable").reset_index(drop=True)
        sonuc = sonuc.drop(
            columns=["_KIRAC_SCHNEIDER_TEKRAR_ANAHTAR", "_KIRAC_SCHNEIDER_STOK_SECIM"],
            errors="ignore"
        )

    # YILDIZ tekrar kuralı:
    # Aynı marka içinde T-Soft'ta aynı Web Servis Koduna gidecek birden fazla
    # satır varsa stoklar toplanmaz. En yüksek stoklu satır güncellemeye girer.
    if tedarikci_norm == "YILDIZ" and not sonuc.empty:
        sonuc = sonuc.copy()
        sonuc["_YILDIZ_TEKRAR_ANAHTAR"] = sonuc.apply(
            lambda row: yildiz_eslesme_kodu_uret(
                row["B2B Web Servis Kodu"],
                row["B2B Marka"],
            )[0],
            axis=1,
        )
        sonuc["_YILDIZ_STOK_SECIM"] = pd.to_numeric(
            sonuc["Yeni Stok"],
            errors="coerce",
        ).fillna(-1)

        anahtar_var = sonuc["_YILDIZ_TEKRAR_ANAHTAR"].astype(str).str.strip().ne("")
        kodlu = sonuc[anahtar_var].copy()
        kodsuz = sonuc[~anahtar_var].copy()
        tekrar_maskesi = kodlu.duplicated(
            subset=["_YILDIZ_TEKRAR_ANAHTAR"],
            keep=False,
        )
        tekrarlar = kodlu[tekrar_maskesi]

        if not tekrarlar.empty:
            print()
            print("YILDIZ tekrar eden kodlar tekillestiriliyor.")
            print("Stoklar toplanmayacak; ayni kodda en yuksek stok tutulacak.")
            print(f"Tekrar eden kaynak satir: {len(tekrarlar)}")

        kodlu = kodlu.sort_values(
            by=[
                "_YILDIZ_TEKRAR_ANAHTAR",
                "_YILDIZ_STOK_SECIM",
                "B2B Satır No",
            ],
            ascending=[True, False, True],
            kind="stable",
        )
        kodlu = kodlu.drop_duplicates(
            subset=["_YILDIZ_TEKRAR_ANAHTAR"],
            keep="first",
        )

        sonuc = pd.concat([kodlu, kodsuz], ignore_index=True, sort=False)
        sonuc = sonuc.sort_values("B2B Satır No", kind="stable").reset_index(drop=True)
        sonuc = sonuc.drop(
            columns=["_YILDIZ_TEKRAR_ANAHTAR", "_YILDIZ_STOK_SECIM"],
            errors="ignore",
        )

    return sonuc


def tsoft_oku():
    if not TSOFT_FILE.exists():
        print("T-Soft seçili alan dosyası bulunamadı:")
        print(TSOFT_FILE)
        print("Önce python tsoft_products_export.py çalıştır.")
        raise SystemExit

    df = pd.read_excel(TSOFT_FILE, dtype=str).fillna("")

    gerekli = [
        "ProductId",
        "Web Servis Kodu",
        "Stok",
        "Stok Birimi",
        "Fiyat Birimi"
    ]

    for kolon in gerekli:
        if kolon not in df.columns:
            print(f"T-Soft dosyasında eksik kolon var: {kolon}")
            print("Gelen kolonlar:")
            for c in df.columns:
                print("-", c)
            raise SystemExit

    if "Model" not in df.columns:
        df["Model"] = ""

    if "Ürün Adı" not in df.columns:
        df["Ürün Adı"] = ""

    if "Ana Kategori Adı" not in df.columns:
        df["Ana Kategori Adı"] = ""

    # Fiyatlandırma için T-Soft'tan mevcut fiyat çekilmez. Eski stok raporu
    # kolon düzeni bozulmasın diye fiyat kolonları boş olarak tutulur.
    for fiyat_kolonu in ["Satış Fiyatı", "KDV Dahil Fiyat"]:
        if fiyat_kolonu not in df.columns:
            df[fiyat_kolonu] = ""

    df = df[
        [
            "ProductId",
            "Web Servis Kodu",
            "Ürün Adı",
            "Model",
            "Ana Kategori Adı",
            "Stok",
            "Stok Birimi",
            "Fiyat Birimi",
            "Satış Fiyatı",
            "KDV Dahil Fiyat"
        ]
    ].copy()

    df["ProductId"] = df["ProductId"].astype(str).str.strip()
    df["Web Servis Kodu"] = df["Web Servis Kodu"].astype(str).str.strip()
    df["Ürün Adı"] = df["Ürün Adı"].astype(str).str.strip()
    df["Model"] = df["Model"].astype(str).str.strip()
    df["Ana Kategori Adı"] = df["Ana Kategori Adı"].astype(str).str.strip()
    df["Stok"] = df["Stok"].astype(str).str.strip()
    df["Stok Birimi"] = df["Stok Birimi"].astype(str).str.strip()
    df["Fiyat Birimi"] = df["Fiyat Birimi"].astype(str).str.strip()
    df["Satış Fiyatı"] = df["Satış Fiyatı"].astype(str).str.strip()
    df["KDV Dahil Fiyat"] = df["KDV Dahil Fiyat"].astype(str).str.strip()

    df["Tsoft Temiz Kod"] = df["Web Servis Kodu"].apply(kod_temizle)
    df["Tsoft Sade Kod"] = df["Web Servis Kodu"].apply(sade_kod)

    return df


def dia_kart_kodunu_tsoft_koduna_cevir(marka, kart_kodu):
    """DIA Kart Kodu teknik ad ise guvenli marka kuralıyla T-Soft koduna cevirir."""
    marka_norm = metin_norm(marka_adi_standart(marka))
    kart_norm = metin_norm(kart_kodu)
    if marka_norm != "OZNUR KABLO":
        return str(kart_kodu).strip()

    nyaf = re.fullmatch(
        r"H0[57]V-K\s*\(NYAF\)\s*1X([0-9]+(?:[,.][0-9]+)?)\s+(.+)",
        kart_norm,
    )
    nyaf_hffr = re.fullmatch(
        r"H0[57]Z1-K\s*\(NYAF\)\s*1X([0-9]+(?:[,.][0-9]+)?)\s+(.+)",
        kart_norm,
    )
    eslesme = nyaf or nyaf_hffr
    if not eslesme:
        return str(kart_kodu).strip()

    kesit = eslesme.group(1).replace(".", ",")
    renk_norm = metin_norm(eslesme.group(2)).replace("-", " ")
    renkler = {
        "KIRMIZI": "Kirmizi",
        "SIYAH": "Siyah",
        "MAVI": "Mavi",
        "KAHVE": "Kahve",
        "SARI YESIL": "Sari-Yesil",
    }
    renk = renkler.get(renk_norm)
    if not renk:
        return str(kart_kodu).strip()

    son_ek = ".NYAF.HFFR" if nyaf_hffr else ".NYAF"
    return f"O.1X{kesit}.{renk}{son_ek}"


def dia_stoklari_oku():
    """DIA'daki marka bazli Kart Kodu ve Gercek Stok degerlerini okur."""
    if not DIA_STOK_FILE.exists():
        raise FileNotFoundError(f"DIA stok dosyasi bulunamadi: {DIA_STOK_FILE}")

    ham_df = pd.read_excel(DIA_STOK_FILE, dtype=str, header=None).fillna("")
    baslik_satiri = None
    for index, row in ham_df.iterrows():
        alanlar = {metin_norm(deger) for deger in row.values if str(deger).strip()}
        if {"MARKA", "KART KODU", "GERCEK STOK"}.issubset(alanlar):
            baslik_satiri = index
            break

    if baslik_satiri is None:
        raise ValueError(
            "DIA dosyasinda Marka, Kart Kodu ve Gercek Stok basliklari bulunamadi."
        )

    df = pd.read_excel(DIA_STOK_FILE, dtype=str, header=baslik_satiri).fillna("")
    marka_kolonu = kolon_bul_esnek(df, ["Marka"])
    kod_kolonu = kolon_bul_esnek(df, ["Kart Kodu"])
    stok_kolonu = kolon_bul_esnek(df, ["Gercek Stok", "Gerçek Stok"])
    if not marka_kolonu or not kod_kolonu or not stok_kolonu:
        raise ValueError("DIA dosyasindaki zorunlu kolonlar okunamadi.")

    sonuc = df[[marka_kolonu, kod_kolonu, stok_kolonu]].copy()
    sonuc["Marka"] = sonuc[marka_kolonu].apply(
        lambda deger: DIA_TOPLAMA_MARKALARI.get(metin_norm(deger), "")
    )
    sonuc = sonuc[sonuc["Marka"].ne("")].copy()
    sonuc["Kaynak"] = "DIA"
    sonuc["Kaynak Kod"] = sonuc[kod_kolonu].astype(str).str.strip()
    sonuc["Eslesme Kodu"] = sonuc.apply(
        lambda row: dia_kart_kodunu_tsoft_koduna_cevir(
            row["Marka"], row["Kaynak Kod"]
        ),
        axis=1,
    )
    sonuc["Temiz Kod"] = sonuc["Eslesme Kodu"].apply(kod_temizle)
    sonuc["Kaynak Stok"] = sonuc[stok_kolonu].apply(stok_degerine_cevir)

    hatali = sonuc["Temiz Kod"].eq("") | sonuc["Kaynak Stok"].isna()
    if hatali.any():
        print("DIA hedef markalarinda bos Kart Kodu veya okunamayan Gercek Stok var:")
        print(sonuc.loc[hatali, ["Marka", "Kaynak Kod", "Kaynak Stok"]].to_string(index=False))
        raise SystemExit

    sonuc["Kaynak Stok"] = sonuc["Kaynak Stok"].astype(int)
    return (
        sonuc.groupby(["Kaynak", "Marka", "Temiz Kod"], as_index=False)
        .agg({"Kaynak Kod": "first", "Kaynak Stok": "sum"})
    )


def sabit_stoklari_oku():
    """WAGO, ORTAC ve EMAS sabit stok Excel'lerini ortak bicimde okur."""
    parcalar = []
    for marka_norm, ayar in SABIT_STOK_DOSYALARI.items():
        dosya = ayar["dosya"]
        if not dosya.exists():
            raise FileNotFoundError(f"Sabit stok dosyasi bulunamadi: {dosya}")

        df = pd.read_excel(dosya, dtype=str).fillna("")
        kod_kolonu = kolon_bul_esnek(df, ayar["kod_kolonlari"])
        stok_kolonu = kolon_bul_esnek(df, ayar["stok_kolonlari"])
        if not kod_kolonu or not stok_kolonu:
            raise ValueError(
                f"{dosya.name} icinde urun kodu veya Stok kolonu bulunamadi."
            )

        parca = df[[kod_kolonu, stok_kolonu]].copy()
        parca["Kaynak"] = "SABIT_STOK"
        parca["Marka"] = marka_adi_standart(marka_norm)
        parca["Kaynak Kod"] = parca[kod_kolonu].astype(str).str.strip()
        parca["Temiz Kod"] = parca["Kaynak Kod"].apply(kod_temizle)
        parca["Kaynak Stok"] = parca[stok_kolonu].apply(stok_degerine_cevir)

        hatali = parca["Temiz Kod"].eq("") | parca["Kaynak Stok"].isna()
        if hatali.any():
            print(f"{dosya.name} icinde bos kod veya okunamayan stok var:")
            print(parca.loc[hatali, ["Kaynak Kod", "Kaynak Stok"]].to_string(index=False))
            raise SystemExit

        parca["Kaynak Stok"] = parca["Kaynak Stok"].astype(int)
        parcalar.append(
            parca[["Kaynak", "Marka", "Kaynak Kod", "Temiz Kod", "Kaynak Stok"]]
        )

    birlesik = pd.concat(parcalar, ignore_index=True)
    return (
        birlesik.groupby(["Kaynak", "Marka", "Temiz Kod"], as_index=False)
        .agg({"Kaynak Kod": "first", "Kaynak Stok": "sum"})
    )


def _liste_degerine_ekle(mevcut, eklenecek, ayirici=", "):
    degerler = []
    for kaynak in [mevcut, eklenecek]:
        for parca in str(kaynak).split(ayirici):
            temiz = parca.strip()
            if temiz and temiz.lower() != "nan" and temiz not in degerler:
                degerler.append(temiz)
    return ayirici.join(degerler)


def marka_stok_kaynaklarini_uygula(hazir_df, tsoft_df):
    """Marka bazinda izinli stok kaynaklarini tek toplamda birlestirir.

    WAGO, ORTAC ve EMAS icin B2B + DIA + sabit stok; OPAS ve MUCCO icin
    yalnizca B2B + DIA kullanilir.
    """
    sonuc = hazir_df.copy()
    for kolon in ["DIA Stok", "Sabit Stok"]:
        if kolon not in sonuc.columns:
            sonuc[kolon] = 0
        sonuc[kolon] = pd.to_numeric(sonuc[kolon], errors="coerce").fillna(0).astype(int)

    hedef_marka_normlari = {
        metin_norm(marka_adi_standart(marka))
        for marka in DIA_TOPLAMA_MARKALARI.values()
    }
    sonuc_marka_norm = sonuc["B2B Marka"].apply(
        lambda deger: metin_norm(marka_adi_standart(deger))
    )
    hedef_maskesi = sonuc_marka_norm.isin(hedef_marka_normlari)
    sonuc.loc[hedef_maskesi, "Guncellenecek Stok"] = (
        pd.to_numeric(sonuc.loc[hedef_maskesi, "Bayi Toplam Stok"], errors="coerce")
        .fillna(0).astype(int)
    )
    sonuc.loc[hedef_maskesi, "Guncelleme Durumu"] = "GONDERILECEK"

    kaynak_df = pd.concat(
        [dia_stoklari_oku(), sabit_stoklari_oku()],
        ignore_index=True,
    )
    tsoft = tsoft_df.copy()
    tsoft["Kaynak Temiz Kod"] = tsoft["Web Servis Kodu"].apply(kod_temizle)
    tsoft_kod_sayisi = tsoft.groupby("Kaynak Temiz Kod").size()

    rapor_satirlari = []
    eslesen_katkilar = []
    for _, kaynak_row in kaynak_df.iterrows():
        temiz_kod = kaynak_row["Temiz Kod"]
        eslesen_sayi = int(tsoft_kod_sayisi.get(temiz_kod, 0))
        rapor = {
            "Kaynak": kaynak_row["Kaynak"],
            "Marka": kaynak_row["Marka"],
            "Kaynak Kod": kaynak_row["Kaynak Kod"],
            "Kaynak Stok": int(kaynak_row["Kaynak Stok"]),
            "T-Soft Web Servis Kodu": "",
            "Durum": "",
        }
        if eslesen_sayi == 0:
            rapor["Durum"] = "T-SOFTTA BULUNAMADI"
            rapor_satirlari.append(rapor)
            continue
        if eslesen_sayi > 1:
            rapor["Durum"] = "BIRDEN FAZLA T-SOFT ESLESMESI"
            rapor_satirlari.append(rapor)
            continue

        tsoft_row = tsoft[tsoft["Kaynak Temiz Kod"].eq(temiz_kod)].iloc[0]
        rapor["T-Soft Web Servis Kodu"] = tsoft_row["Web Servis Kodu"]
        rapor["Durum"] = "ESLESTI"
        rapor_satirlari.append(rapor)
        eslesen_katkilar.append({
            "ProductId": str(tsoft_row["ProductId"]).strip(),
            "Web Servis Kodu": str(tsoft_row["Web Servis Kodu"]).strip(),
            "Model": str(tsoft_row["Model"]).strip(),
            "Marka": kaynak_row["Marka"],
            "Marka Norm": metin_norm(marka_adi_standart(kaynak_row["Marka"])),
            "DIA Stok": int(kaynak_row["Kaynak Stok"]) if kaynak_row["Kaynak"] == "DIA" else 0,
            "Sabit Stok": int(kaynak_row["Kaynak Stok"]) if kaynak_row["Kaynak"] == "SABIT_STOK" else 0,
            "Kaynaklar": kaynak_row["Kaynak"],
            "Kaynak Kodlar": str(kaynak_row["Kaynak Kod"]).strip(),
            "Stok Birimi": str(tsoft_row["Stok Birimi"]).strip(),
            "Fiyat Birimi": str(tsoft_row["Fiyat Birimi"]).strip(),
            "Satış Fiyatı": str(tsoft_row["Satış Fiyatı"]).strip(),
        })

    if eslesen_katkilar:
        katki_df = pd.DataFrame(eslesen_katkilar)
        katki_df = (
            katki_df.groupby(["ProductId", "Web Servis Kodu", "Marka", "Marka Norm"], as_index=False)
            .agg({
                "Model": "first",
                "DIA Stok": "sum",
                "Sabit Stok": "sum",
                "Kaynaklar": lambda x: ", ".join(dict.fromkeys(map(str, x))),
                "Kaynak Kodlar": lambda x: " | ".join(dict.fromkeys(map(str, x))),
                "Stok Birimi": "first",
                "Fiyat Birimi": "first",
                "Satış Fiyatı": "first",
            })
        )
    else:
        katki_df = pd.DataFrame()

    yeni_satirlar = []
    sonuc_product_id = sonuc["ProductId"].astype(str).str.strip()
    mevcut_marka_norm = sonuc["B2B Marka"].apply(
        lambda deger: metin_norm(marka_adi_standart(deger))
    )
    yeni_product_markalari = {}
    for _, katki in katki_df.iterrows():
        product_id = str(katki["ProductId"]).strip()
        marka_norm = str(katki["Marka Norm"])
        mevcut_maskesi = (
            sonuc_product_id.eq(product_id)
            & mevcut_marka_norm.eq(marka_norm)
        )

        if mevcut_maskesi.any():
            indexler = sonuc.index[mevcut_maskesi]
            sonuc.loc[indexler, "DIA Stok"] = int(katki["DIA Stok"])
            sonuc.loc[indexler, "Sabit Stok"] = int(katki["Sabit Stok"])
            sonuc.loc[indexler, "Guncellenecek Stok"] = (
                pd.to_numeric(sonuc.loc[indexler, "Bayi Toplam Stok"], errors="coerce")
                .fillna(0).astype(int)
                + int(katki["DIA Stok"])
                + int(katki["Sabit Stok"])
            )
            sonuc.loc[indexler, "Guncelleme Durumu"] = "GONDERILECEK"
            sonuc.loc[indexler, "Tedarikciler"] = sonuc.loc[indexler, "Tedarikciler"].apply(
                lambda deger: _liste_degerine_ekle(deger, katki["Kaynaklar"])
            )
            sonuc.loc[indexler, "B2B Kodları"] = sonuc.loc[indexler, "B2B Kodları"].apply(
                lambda deger: _liste_degerine_ekle(deger, katki["Kaynak Kodlar"], " | ")
            )
            sonuc.loc[indexler, "Kaynak Satır Sayısı"] = (
                pd.to_numeric(sonuc.loc[indexler, "Kaynak Satır Sayısı"], errors="coerce")
                .fillna(0).astype(int)
                + len(str(katki["Kaynaklar"]).split(", "))
            )
            continue

        ayni_product = sonuc_product_id.eq(product_id)
        yeni_marka_norm = yeni_product_markalari.get(product_id)
        if ayni_product.any() or (yeni_marka_norm and yeni_marka_norm != marka_norm):
            rapor_satirlari.append({
                "Kaynak": katki["Kaynaklar"],
                "Marka": katki["Marka"],
                "Kaynak Kod": katki["Kaynak Kodlar"],
                "Kaynak Stok": int(katki["DIA Stok"]) + int(katki["Sabit Stok"]),
                "T-Soft Web Servis Kodu": katki["Web Servis Kodu"],
                "Durum": "MARKA CAKISMASI - GONDERILMEYECEK",
            })
            continue

        yeni_satir = {kolon: "" for kolon in sonuc.columns}
        yeni_satir.update({
            "ProductId": product_id,
            "Web Servis Kodu": katki["Web Servis Kodu"],
            "Model": katki["Model"],
            "B2B Marka": katki["Marka"],
            "Marka Uyuşmazlığı": "",
            "Tsoft Mevcut Stok": 0,
            "Bayi Toplam Stok": 0,
            "DIA Stok": int(katki["DIA Stok"]),
            "Sabit Stok": int(katki["Sabit Stok"]),
            "Guncellenecek Stok": int(katki["DIA Stok"]) + int(katki["Sabit Stok"]),
            "Guncelleme Durumu": "GONDERILECEK",
            "Stok Birimi": katki["Stok Birimi"],
            "Fiyat Birimi": katki["Fiyat Birimi"],
            "Satış Fiyatı": katki["Satış Fiyatı"],
            "Tedarikciler": katki["Kaynaklar"],
            "B2B Kodları": katki["Kaynak Kodlar"],
            "Kaynak Satır Sayısı": len(str(katki["Kaynaklar"]).split(", ")),
        })
        yeni_satirlar.append(yeni_satir)
        yeni_product_markalari[product_id] = marka_norm

    if yeni_satirlar:
        sonuc = pd.concat([sonuc, pd.DataFrame(yeni_satirlar)], ignore_index=True, sort=False)

    rapor_df = pd.DataFrame(rapor_satirlari).drop_duplicates()
    excel_kaydet_guvenli(rapor_df, CIKTI_DIR / "dia_sabit_stok_eslesme_raporu.xlsx")
    excel_kaydet_guvenli(
        rapor_df[rapor_df["Marka"].apply(metin_norm).eq("WAGO")].copy(),
        CIKTI_DIR / "wago_dia_eslesme_raporu.xlsx",
    )

    sonuc_marka_norm = sonuc["B2B Marka"].apply(
        lambda deger: metin_norm(marka_adi_standart(deger))
    )
    sabit_marka_normlari = {
        metin_norm(marka_adi_standart(marka))
        for marka in SABIT_STOK_DOSYALARI
    }
    for marka_norm in sorted(hedef_marka_normlari):
        marka_sonuc = sonuc[sonuc_marka_norm.eq(marka_norm)]
        kaynak_etiketi = (
            "DIA/Sabit/B2B"
            if marka_norm in sabit_marka_normlari
            else "DIA/B2B"
        )
        print(
            f"{marka_adi_standart(marka_norm)} {kaynak_etiketi} toplam urunu:",
            len(marka_sonuc),
        )
    return sonuc


def tsoft_lookup_olustur(tsoft_df):
    lookup = {}

    for idx, row in tsoft_df.iterrows():
        anahtarlar = eslesme_anahtarlari_uret(row["Web Servis Kodu"])

        for key, tip, deger in anahtarlar:
            lookup.setdefault(key, []).append(idx)

    return lookup


def b2b_eslesme_bul(eslesmede_kullanilacak_kod, tsoft_df, lookup):
    anahtarlar = eslesme_anahtarlari_uret(eslesmede_kullanilacak_kod)

    bulunanlar = []
    kullanilanlar = []

    for key, tip, deger in anahtarlar:
        idx_list = lookup.get(key, [])

        for idx in idx_list:
            if idx not in bulunanlar:
                bulunanlar.append(idx)
                kullanilanlar.append((tip, deger))

    if len(bulunanlar) == 0:
        return None, "", ""

    if len(bulunanlar) == 1:
        idx = bulunanlar[0]
        tip, deger = kullanilanlar[0]
        return tsoft_df.loc[idx], tip, deger

    return "COKLU", kullanilanlar[0][0], kullanilanlar[0][1]


def oznurkablo_direkt_eslesme_bul(b2b_row, tsoft_df):
    """KIRAC/SANTRAL/YILDIZ OZNUR hedefini birebir web servis koduyla bulur.

    Genel sade-kod anahtari virgulu attigi icin 2,5 ve 25 kesitleri ayni
    anahtara dusebilir. OZNUR donusumu zaten tam T-Soft kodu urettiginden bu
    iki kaynakta gevsek anahtar kullanmak yerine birebir kod tercih edilir.
    """
    tedarikci = metin_norm(b2b_row.get("Tedarikci", ""))
    marka = metin_norm(b2b_row.get("B2B Marka", ""))
    if tedarikci not in {"KIRAC", "SANTRAL", "YILDIZ"}:
        return None
    if marka not in {"OZNUR", "OZNUR KABLO"}:
        return None

    aranan_kod = kod_temizle(b2b_row.get("Eslesmede Kullanilacak Kod", ""))
    bulunanlar = tsoft_df.index[
        tsoft_df["Tsoft Temiz Kod"].eq(aranan_kod)
    ].tolist()

    if len(bulunanlar) == 0:
        return None, "OZNUR_BIREBIR", aranan_kod
    if len(bulunanlar) == 1:
        return tsoft_df.loc[bulunanlar[0]], "OZNUR_BIREBIR", aranan_kod
    return "COKLU", "OZNUR_BIREBIR", aranan_kod


def santral_product_id_eslesme_bul(b2b_row, tsoft_df):
    if metin_norm(b2b_row.get("Tedarikci", "")) != "SANTRAL":
        return None

    b2b_kod = metin_norm(b2b_row.get("B2B Web Servis Kodu", ""))
    product_id = SANTRAL_KOD_PRODUCT_ID_ESLESME.get(b2b_kod)
    if not product_id:
        return None

    eslesen = tsoft_df[
        tsoft_df["ProductId"].astype(str).str.strip().eq(product_id)
    ]
    if eslesen.empty:
        return None

    return eslesen.iloc[0]


YILDIZ_ACK_KOD_ESLESMELERI = {
    "ACK-AA09-00713": "AA13-01223",
}


def yildiz_eslesme_kodu_uret(web_servis_kodu, marka):
    """YILDIZ'a özel marka kurallarından sonra T-Soft eşleşme kodunu üretir."""
    aranan_kod = kod_temizle(web_servis_kodu)
    marka_anahtari = metin_norm(marka)

    if marka_anahtari.startswith("ACK"):
        # Bilinen tekil karşılıklar önce uygulanır. Diğer ACK kodlarında
        # B2B'deki ACK- öneki kaldırılarak T-Soft Web Servis Kodu aranır.
        duzeltilmis = YILDIZ_ACK_KOD_ESLESMELERI.get(aranan_kod)

        if duzeltilmis is None:
            duzeltilmis = re.sub(r"^ACK[\s\-_]*", "", aranan_kod).strip()

        eslesme_bicimi = (
            "YILDIZ_ACK_OZEL"
            if duzeltilmis != aranan_kod
            else "YILDIZ_DIREKT"
        )
        return duzeltilmis, eslesme_bicimi

    if marka_anahtari.startswith("NADE"):
        duzeltilmis = re.sub(r"^NADE[\s\-_]*", "", aranan_kod).strip()
        eslesme_bicimi = (
            "YILDIZ_NADE_PREFIX"
            if duzeltilmis != aranan_kod
            else "YILDIZ_DIREKT"
        )
        return duzeltilmis, eslesme_bicimi

    if marka_anahtari.startswith("NEAD"):
        duzeltilmis = re.sub(r"^NEAD[\s\-_]*", "", aranan_kod).strip()
        eslesme_bicimi = (
            "YILDIZ_NEAD_PREFIX"
            if duzeltilmis != aranan_kod
            else "YILDIZ_DIREKT"
        )
        return duzeltilmis, eslesme_bicimi

    if marka_anahtari.startswith("CETINKAYA"):
        # metin_norm Ç harfini C'ye çevirdiği için ÇP10954 burada CP10954 olur.
        # T-Soft'taki ÇP 10954 biçimiyle eşleşmesi için önekten sonra boşluk eklenir.
        duzeltilmis = re.sub(r"^CP[\s\-_]*(\d+)$", r"CP \1", aranan_kod).strip()
        eslesme_bicimi = (
            "YILDIZ_CETINKAYA_BOSLUK"
            if duzeltilmis != aranan_kod
            else "YILDIZ_DIREKT"
        )
        return duzeltilmis, eslesme_bicimi

    return aranan_kod, "YILDIZ_DIREKT"


def yildiz_direkt_eslesme_bul(web_servis_kodu, marka, tsoft_df):
    """YILDIZ kodlarını marka kurallarından sonra T-Soft koduyla eşleştirir."""
    aranan_kod, eslesme_bicimi = yildiz_eslesme_kodu_uret(
        web_servis_kodu,
        marka,
    )

    if not aranan_kod:
        return None, "", ""

    bulunanlar = tsoft_df.index[
        tsoft_df["Tsoft Temiz Kod"] == aranan_kod
    ].tolist()

    if len(bulunanlar) == 0:
        return None, "", ""

    if len(bulunanlar) == 1:
        return tsoft_df.loc[bulunanlar[0]], eslesme_bicimi, aranan_kod

    return "COKLU", eslesme_bicimi, aranan_kod


def bos_df(kolonlar):
    return pd.DataFrame(columns=kolonlar)


def hazir_stok_olustur(tedarikci_bazli_df):
    kolon_sirasi = [
        "ProductId",
        "Web Servis Kodu",
        "Model",
        "B2B Marka",
        "Marka Uyuşmazlığı",
        "Tsoft Mevcut Stok",
        "Bayi Toplam Stok",
        "Guncellenecek Stok",
        "Guncelleme Durumu",
        "Stok Birimi",
        "Fiyat Birimi",
        "Satış Fiyatı",
        "Tedarikciler",
        "B2B Kodları",
        "Kaynak Satır Sayısı",
    ]

    if tedarikci_bazli_df.empty or "Yeni Stok" not in tedarikci_bazli_df.columns:
        return bos_df(kolon_sirasi)

    df = tedarikci_bazli_df.copy()
    df["Yeni Stok"] = pd.to_numeric(df["Yeni Stok"], errors="coerce").fillna(0).astype(int)

    if "Model" not in df.columns:
        df["Model"] = ""

    if "B2B Marka" not in df.columns:
        df["B2B Marka"] = ""

    if "Guncelleme Durumu" not in df.columns:
        df["Guncelleme Durumu"] = "GONDERILECEK"

    df["ProductId"] = df["ProductId"].astype(str).str.strip()
    df["Web Servis Kodu"] = df["Web Servis Kodu"].astype(str).str.strip()

    grouped = df.groupby(
        ["ProductId", "Web Servis Kodu"],
        dropna=False
    ).agg({
        "Model": "first",
        "B2B Marka": lambda x: ", ".join(sorted(set(
            marka_adi_standart(v)
            for v in x
            if marka_adi_standart(v)
        ))),
        "Yeni Stok": "sum",
        "Guncelleme Durumu": lambda x: (
            "GONDERILMEYECEK"
            if x.astype(str).str.upper().eq("GONDERILMEYECEK").any()
            else "GONDERILECEK"
        ),
        "Tsoft Mevcut Stok": "first",
        "Stok Birimi": "first",
        "Fiyat Birimi": "first",
        "Satış Fiyatı": "first",
        "Tedarikci": lambda x: ", ".join(sorted(set(map(str, x)))),
        "B2B Web Servis Kodu": lambda x: " | ".join(list(dict.fromkeys(map(str, x)))[:30]),
    }).reset_index()

    grouped.rename(
        columns={
            "Yeni Stok": "Bayi Toplam Stok",
            "Tedarikci": "Tedarikciler",
            "B2B Web Servis Kodu": "B2B Kodları"
        },
        inplace=True
    )

    grouped["Marka Uyuşmazlığı"] = grouped["B2B Marka"].apply(
        lambda deger: "EVET" if "," in str(deger) else ""
    )

    kaynak_sayilari = df.groupby(
        ["ProductId", "Web Servis Kodu"],
        dropna=False
    ).size().reset_index(name="Kaynak Satır Sayısı")

    grouped = grouped.merge(
        kaynak_sayilari,
        on=["ProductId", "Web Servis Kodu"],
        how="left"
    )

    grouped["Guncellenecek Stok"] = grouped["Bayi Toplam Stok"].astype(int)

    return grouped[kolon_sirasi].copy()



def tedarikciye_gore_filtrele(df, tedarikci_adi):
    if df is None or df.empty:
        return pd.DataFrame(columns=[] if df is None else df.columns)

    if "Tedarikci" not in df.columns:
        return pd.DataFrame(columns=df.columns)

    aranan = metin_norm(tedarikci_adi)
    maske = df["Tedarikci"].apply(
        lambda x: aranan in metin_norm(x)
    )

    return df[maske].copy()


def santral_tekrarlarda_en_yuksek_stogu_sec(df, tedarikci="SANTRAL"):
    """Ayni T-Soft urununde toplama yapmadan en yuksek stoklu satiri tutar."""
    if df is None or df.empty or "Yeni Stok" not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()

    sonuc = df.copy()
    sonuc["_SANTRAL_KAYNAK_SIRA"] = range(len(sonuc))
    sonuc["_SANTRAL_STOK_SECIM"] = pd.to_numeric(
        sonuc["Yeni Stok"],
        errors="coerce",
    ).fillna(-1)

    def tekrar_anahtari(row):
        product_id = str(row.get("ProductId", "")).strip()
        if product_id:
            return f"PID:{product_id}"

        web_servis_kodu = metin_norm(row.get("Web Servis Kodu", ""))
        if web_servis_kodu:
            return f"KOD:{web_servis_kodu}"

        return ""

    sonuc["_SANTRAL_TEKRAR_ANAHTAR"] = sonuc.apply(tekrar_anahtari, axis=1)
    anahtar_var = sonuc["_SANTRAL_TEKRAR_ANAHTAR"].astype(str).str.strip() != ""
    anahtarli = sonuc[anahtar_var].copy()
    anahtarsiz = sonuc[~anahtar_var].copy()

    tekrar_maskesi = anahtarli.duplicated(
        subset=["_SANTRAL_TEKRAR_ANAHTAR"],
        keep=False,
    )
    tekrar_satir_sayisi = int(tekrar_maskesi.sum())

    # Stok azalan sirada oldugu icin her urunden en yuksek stok tutulur.
    # Stoklar esitse kaynak Excel'de ustte olan tek satir kalir.
    anahtarli = anahtarli.sort_values(
        by=[
            "_SANTRAL_TEKRAR_ANAHTAR",
            "_SANTRAL_STOK_SECIM",
            "_SANTRAL_KAYNAK_SIRA",
        ],
        ascending=[True, False, True],
        kind="stable",
    )
    anahtarli = anahtarli.drop_duplicates(
        subset=["_SANTRAL_TEKRAR_ANAHTAR"],
        keep="first",
    )

    sonuc = pd.concat([anahtarli, anahtarsiz], ignore_index=True, sort=False)
    sonuc = sonuc.sort_values("_SANTRAL_KAYNAK_SIRA", kind="stable")

    if tekrar_satir_sayisi:
        print()
        print(f"{tedarikci} stok guncelleneceklerde tekrar eden urunler tekillestirildi.")
        print("Stoklar toplanmadi; ayni urunde en yuksek stoklu tek satir secildi.")
        print(f"Tekrar eden kaynak satir: {tekrar_satir_sayisi}")

    return sonuc.drop(
        columns=[
            "_SANTRAL_KAYNAK_SIRA",
            "_SANTRAL_STOK_SECIM",
            "_SANTRAL_TEKRAR_ANAHTAR",
        ],
        errors="ignore",
    ).reset_index(drop=True)


def durum_tablosu_ekle(kaynak_df, durum, ekstra_kolonlar=None):
    if kaynak_df is None or kaynak_df.empty:
        return pd.DataFrame()

    df = kaynak_df.copy()
    df["Eşleşme Durumu"] = durum

    if "Neden" not in df.columns:
        df["Neden"] = ""

    kolonlar = [
        "B2B Satır No",
        "B2B Kaynak Dosya",
        "B2B Web Servis Kodu",
        "Eşleşme Durumu",
        "Neden",
    ]

    if ekstra_kolonlar:
        kolonlar.extend(ekstra_kolonlar)

    mevcut_kolonlar = [kolon for kolon in kolonlar if kolon in df.columns]

    return df[mevcut_kolonlar].copy()


def santral_eslesen_filtre_ekle(df):
    if df is None or df.empty:
        return df

    df = df.copy()

    if "Tedarikci" not in df.columns:
        return df

    santral_mask = df["Tedarikci"].apply(lambda x: "SANTRAL" in metin_norm(x))

    if not santral_mask.any():
        return df

    def filtre_uret(row):
        durum = str(row.get("SANTRAL Kod Düzeltme Durumu", "")).strip()
        durum_norm = metin_norm(durum)

        if "SCHRACK" in durum_norm:
            if "ESLESTI" in durum_norm:
                return "SCHRACK kod eşleşti"
            if "KARSILIGI BULUNAMADI" in durum_norm:
                return "SCHRACK karşılığı bulunamadı"
            return durum

        if "KLEMSAN" in durum_norm:
            return durum

        if "CEVRILDI" in durum_norm:
            return "Modelden koda çevrildi"

        return "Zaten ürün kodu"

    df.loc[santral_mask, "SANTRAL Dönüşüm Filtresi"] = df.loc[santral_mask].apply(
        filtre_uret,
        axis=1
    )

    return df


SANTRAL_OZEL_KOLONLAR = [
    "SANTRAL Orijinal Ürün Kodu",
    "SANTRAL Kod Düzeltme Durumu",
    "SANTRAL Kod Düzeltme Notu",
    "KIRAC Legrand Orijinal Kod",
    "KIRAC Legrand Sıfırsız Kod",
    "KIRAC Legrand Sıfır Kuralı",
    "KIRAC Bosch Orijinal Kod",
    "KIRAC Bosch Boşluklu Kod",
    "KIRAC Bosch Boşluk Kuralı",
    "KIRAC Ortac Orijinal Kod",
    "KIRAC Ortac Boşluklu Kod",
    "KIRAC Ortac Boşluk Kuralı",
    "KIRAC Şafak Orijinal Kod",
    "KIRAC Şafak Boşluklu Kod",
    "KIRAC Şafak KM Kuralı",
    "KIRAC OZNUR Orijinal Kod",
    "KIRAC OZNUR Tsoft Kodu",
    "KIRAC OZNUR Kural",
    "SANTRAL OZNUR Orijinal Kod",
    "SANTRAL OZNUR Tsoft Kodu",
    "SANTRAL OZNUR Kural",
    "YILDIZ OZNUR Orijinal Kod",
    "YILDIZ OZNUR Tsoft Kodu",
    "YILDIZ OZNUR Kural",
]


def santral_kolonlarini_cikar(df):
    if df is None or df.empty:
        return df

    return df.drop(
        columns=[kolon for kolon in SANTRAL_OZEL_KOLONLAR if kolon in df.columns],
        errors="ignore"
    )


def gercek_eslesmeyenleri_filtrele(df):
    if df is None or df.empty:
        return df

    tutulacak_ifadeler = [
        "B2B'DE VAR, SITEDE YOK",
        "MANUEL T-SOFT KODU BULUNDU AMA T-SOFT URUN LISTESINDE ESLESMEDI",
        "MODEL KARSILIGI BULUNAMADI",
        "KLEMSAN MODEL KARSILIGI BULUNAMADI",
        "KLEMSAN COKLU STOK KODU",
        "COKLU ESLESME",
        "MANUEL ESLESME TABLOSUNDA KARSILIK BULUNAMADI",
    ]

    kontrol_kolonlari = [
        "Neden",
        "SANTRAL Kod Düzeltme Durumu",
        "SANTRAL Kod Düzeltme Notu",
        "KIRAC Kod Düzeltme Durumu",
        "KIRAC Kod Düzeltme Notu",
    ]

    mevcut_kolonlar = [kolon for kolon in kontrol_kolonlari if kolon in df.columns]
    if not mevcut_kolonlar:
        return df.iloc[0:0].copy()

    def tutulacak_mi(row):
        metin = " ".join(str(row.get(kolon, "")) for kolon in mevcut_kolonlar)
        metin = metin_norm(metin)
        if not metin:
            return False
        return any(ifade in metin for ifade in tutulacak_ifadeler)

    return df[df.apply(tutulacak_mi, axis=1)].copy()


def konya_raporlari_olustur(
    b2b_all,
    tedarikci_bazli_df,
    b2bde_var_sitede_yok_df,
    bos_kod_df,
    hatali_stok_df,
    coklu_df
):
    """
    KONYA için özel kontrol dosyaları oluşturur:
    1- konya_eslesenler.xlsx
    2- konya_eslesmeyenler.xlsx
    3- konya_stok_guncellenecekler.xlsx
    4- konya_kod_kontrol.xlsx
    5- konya_tsoftta_yok_b2b_stok_var.xlsx
    6- konya_tsoftta_yok_b2b_stok_yok.xlsx
    7- konya_bos_urun_kodu.xlsx
    """

    konya_all = tedarikciye_gore_filtrele(b2b_all, "KONYA")
    konya_eslesenler = tedarikciye_gore_filtrele(tedarikci_bazli_df, "KONYA")

    konya_b2bde_var_sitede_yok = tedarikciye_gore_filtrele(
        b2bde_var_sitede_yok_df,
        "KONYA"
    )
    konya_bos_kod = tedarikciye_gore_filtrele(bos_kod_df, "KONYA")
    konya_hatali_stok = tedarikciye_gore_filtrele(hatali_stok_df, "KONYA")
    konya_coklu = tedarikciye_gore_filtrele(coklu_df, "KONYA")

    konya_eslesmeyenler_listesi = [
        df for df in [
            konya_b2bde_var_sitede_yok,
            konya_bos_kod,
            konya_hatali_stok,
            konya_coklu,
        ]
        if df is not None and not df.empty
    ]

    if konya_eslesmeyenler_listesi:
        konya_eslesmeyenler = pd.concat(
            konya_eslesmeyenler_listesi,
            ignore_index=True,
            sort=False
        )
        konya_eslesmeyenler = gercek_eslesmeyenleri_filtrele(konya_eslesmeyenler)
    else:
        konya_eslesmeyenler = pd.DataFrame()

    # KONYA eşleşmeyenlerini daha net ayırıyoruz:
    # 1- T-Soft'ta yok, B2B'de stok var
    # 2- T-Soft'ta yok, B2B'de stok yok
    # 3- B2B ürün kodu boş
    konya_bos_urun_kodu = konya_bos_kod.copy()

    if not konya_b2bde_var_sitede_yok.empty and "Yeni Stok" in konya_b2bde_var_sitede_yok.columns:
        konya_tsoftta_yok_df = konya_b2bde_var_sitede_yok.copy()
        konya_tsoftta_yok_df["Yeni Stok Sayı"] = pd.to_numeric(
            konya_tsoftta_yok_df["Yeni Stok"],
            errors="coerce"
        ).fillna(0).astype(int)

        konya_tsoftta_yok_b2b_stok_var = konya_tsoftta_yok_df[
            konya_tsoftta_yok_df["Yeni Stok Sayı"] > 0
        ].copy()

        konya_tsoftta_yok_b2b_stok_yok = konya_tsoftta_yok_df[
            konya_tsoftta_yok_df["Yeni Stok Sayı"] <= 0
        ].copy()

        for df in [
            konya_tsoftta_yok_b2b_stok_var,
            konya_tsoftta_yok_b2b_stok_yok,
        ]:
            df.drop(
                columns=["Yeni Stok Sayı"],
                inplace=True,
                errors="ignore"
            )
    else:
        konya_tsoftta_yok_b2b_stok_var = pd.DataFrame()
        konya_tsoftta_yok_b2b_stok_yok = pd.DataFrame()

    konya_stok_guncellenecekler = guncelleme_kaynagi_kurallarini_uygula(
        konya_eslesenler
    )
    konya_stok_guncellenecekler = santral_tekrarlarda_en_yuksek_stogu_sec(
        konya_stok_guncellenecekler,
        tedarikci="KONYA",
    )

    durum_df_listesi = []

    durum_df_listesi.append(
        durum_tablosu_ekle(
            konya_eslesenler,
            "EŞLEŞTİ",
            ekstra_kolonlar=[
                "ProductId",
                "Web Servis Kodu",
                "Model",
                "Yeni Stok",
                "Tsoft Mevcut Stok",
                "Eslesme Bicimi",
                "Kullanılan Kod",
            ]
        )
    )

    durum_df_listesi.append(
        durum_tablosu_ekle(konya_b2bde_var_sitede_yok, "EŞLEŞMEDİ")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(konya_bos_kod, "BOŞ KOD")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(konya_hatali_stok, "HATALI STOK")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(konya_coklu, "ÇOKLU EŞLEŞME")
    )

    durum_df_listesi = [df for df in durum_df_listesi if df is not None and not df.empty]

    if durum_df_listesi:
        durum_df = pd.concat(durum_df_listesi, ignore_index=True, sort=False)
    else:
        durum_df = pd.DataFrame()

    if not konya_all.empty:
        konya_kod_kontrol = konya_all.copy()

        if not durum_df.empty and "B2B Satır No" in konya_kod_kontrol.columns and "B2B Satır No" in durum_df.columns:
            durum_df = durum_df.drop_duplicates(subset=["B2B Satır No"], keep="first")

            eklenecek_kolonlar = [
                kolon for kolon in durum_df.columns
                if kolon not in [
                    "B2B Kaynak Dosya",
                    "B2B Web Servis Kodu",
                ]
            ]

            konya_kod_kontrol = konya_kod_kontrol.merge(
                durum_df[eklenecek_kolonlar],
                on="B2B Satır No",
                how="left"
            )
        else:
            konya_kod_kontrol["Eşleşme Durumu"] = "KONTROL EDİLEMEDİ"
            konya_kod_kontrol["Neden"] = "B2B Satır No bulunamadı"

        if "Eşleşme Durumu" not in konya_kod_kontrol.columns:
            konya_kod_kontrol["Eşleşme Durumu"] = "KONTROL EDİLEMEDİ"

        konya_kod_kontrol["Eşleşme Durumu"] = konya_kod_kontrol["Eşleşme Durumu"].fillna(
            "KONTROL EDİLEMEDİ"
        )

        if "Neden" not in konya_kod_kontrol.columns:
            konya_kod_kontrol["Neden"] = ""

        konya_kod_kontrol["Neden"] = konya_kod_kontrol["Neden"].fillna("")
    else:
        konya_kod_kontrol = pd.DataFrame()

    konya_eslesenler = santral_kolonlarini_cikar(konya_eslesenler)
    konya_eslesmeyenler = santral_kolonlarini_cikar(konya_eslesmeyenler)
    konya_stok_guncellenecekler = santral_kolonlarini_cikar(konya_stok_guncellenecekler)
    konya_kod_kontrol = santral_kolonlarini_cikar(konya_kod_kontrol)
    konya_tsoftta_yok_b2b_stok_var = santral_kolonlarini_cikar(konya_tsoftta_yok_b2b_stok_var)
    konya_tsoftta_yok_b2b_stok_yok = santral_kolonlarini_cikar(konya_tsoftta_yok_b2b_stok_yok)
    konya_bos_urun_kodu = santral_kolonlarini_cikar(konya_bos_urun_kodu)

    return {
        "konya_eslesenler": konya_eslesenler,
        "konya_eslesmeyenler": konya_eslesmeyenler,
        "konya_stok_guncellenecekler": konya_stok_guncellenecekler,
        "konya_kod_kontrol": konya_kod_kontrol,
        "konya_tsoftta_yok_b2b_stok_var": konya_tsoftta_yok_b2b_stok_var,
        "konya_tsoftta_yok_b2b_stok_yok": konya_tsoftta_yok_b2b_stok_yok,
        "konya_bos_urun_kodu": konya_bos_urun_kodu,
    }



def kirac_raporlari_olustur(
    b2b_all,
    tedarikci_bazli_df,
    b2bde_var_sitede_yok_df,
    bos_kod_df,
    hatali_stok_df,
    coklu_df
):
    """
    KIRAC için özel kontrol dosyaları oluşturur:
    1- kirac_eslesenler.xlsx
    2- kirac_eslesmeyenler.xlsx
    3- kirac_stok_guncellenecekler.xlsx
    4- kirac_kod_kontrol.xlsx
    5- kirac_tsoftta_yok_b2b_stok_var.xlsx
    6- kirac_tsoftta_yok_b2b_stok_yok.xlsx
    7- kirac_bos_urun_kodu.xlsx
    """

    kirac_all = tedarikciye_gore_filtrele(b2b_all, "KIRAC")
    kirac_eslesenler = tedarikciye_gore_filtrele(tedarikci_bazli_df, "KIRAC")

    kirac_b2bde_var_sitede_yok = tedarikciye_gore_filtrele(
        b2bde_var_sitede_yok_df,
        "KIRAC"
    )
    kirac_bos_kod = tedarikciye_gore_filtrele(bos_kod_df, "KIRAC")
    kirac_hatali_stok = tedarikciye_gore_filtrele(hatali_stok_df, "KIRAC")
    kirac_coklu = tedarikciye_gore_filtrele(coklu_df, "KIRAC")

    kirac_eslesmeyenler_listesi = [
        df for df in [
            kirac_b2bde_var_sitede_yok,
            kirac_bos_kod,
            kirac_hatali_stok,
            kirac_coklu,
        ]
        if df is not None and not df.empty
    ]

    if kirac_eslesmeyenler_listesi:
        kirac_eslesmeyenler = pd.concat(
            kirac_eslesmeyenler_listesi,
            ignore_index=True,
            sort=False
        )
        kirac_eslesmeyenler = gercek_eslesmeyenleri_filtrele(kirac_eslesmeyenler)
    else:
        kirac_eslesmeyenler = pd.DataFrame()

    # KIRAC eşleşmeyenlerini daha net ayırıyoruz:
    # 1- T-Soft'ta yok, B2B'de stok var
    # 2- T-Soft'ta yok, B2B'de stok yok
    # 3- B2B ürün kodu boş
    kirac_bos_urun_kodu = kirac_bos_kod.copy()

    if not kirac_b2bde_var_sitede_yok.empty and "Yeni Stok" in kirac_b2bde_var_sitede_yok.columns:
        kirac_tsoftta_yok_df = kirac_b2bde_var_sitede_yok.copy()
        kirac_tsoftta_yok_df["Yeni Stok Sayı"] = pd.to_numeric(
            kirac_tsoftta_yok_df["Yeni Stok"],
            errors="coerce"
        ).fillna(0).astype(int)

        kirac_tsoftta_yok_b2b_stok_var = kirac_tsoftta_yok_df[
            kirac_tsoftta_yok_df["Yeni Stok Sayı"] > 0
        ].copy()

        kirac_tsoftta_yok_b2b_stok_yok = kirac_tsoftta_yok_df[
            kirac_tsoftta_yok_df["Yeni Stok Sayı"] <= 0
        ].copy()

        for df in [
            kirac_tsoftta_yok_b2b_stok_var,
            kirac_tsoftta_yok_b2b_stok_yok,
        ]:
            df.drop(
                columns=["Yeni Stok Sayı"],
                inplace=True,
                errors="ignore"
            )
    else:
        kirac_tsoftta_yok_b2b_stok_var = pd.DataFrame()
        kirac_tsoftta_yok_b2b_stok_yok = pd.DataFrame()

    kirac_stok_guncellenecekler = guncelleme_kaynagi_kurallarini_uygula(
        kirac_eslesenler
    )

    if not kirac_stok_guncellenecekler.empty:
        tekil_anahtarlar = [
            kolon for kolon in ["ProductId", "Web Servis Kodu"]
            if kolon in kirac_stok_guncellenecekler.columns
        ]

        if tekil_anahtarlar:
            tekrarlar = kirac_stok_guncellenecekler[
                kirac_stok_guncellenecekler.duplicated(
                    subset=tekil_anahtarlar,
                    keep=False
                )
            ].copy()

            if not tekrarlar.empty:
                secilecek_satirlar = []

                def stok_sayisi_en_yuksek_satiri_sec(satirlar):
                    satirlar = satirlar.copy()
                    if "Yeni Stok" not in satirlar.columns:
                        return satirlar.tail(1)

                    satirlar["_secim_stok"] = pd.to_numeric(
                        satirlar["Yeni Stok"],
                        errors="coerce"
                    ).fillna(-1)

                    return satirlar.sort_values("_secim_stok").tail(1).drop(
                        columns=["_secim_stok"],
                        errors="ignore"
                    )

                for _, grup in kirac_stok_guncellenecekler.groupby(
                    tekil_anahtarlar,
                    dropna=False,
                    sort=False
                ):
                    web_kodu = ""
                    b2b_marka = ""

                    if "Web Servis Kodu" in grup.columns:
                        web_kodu = str(grup["Web Servis Kodu"].iloc[0]).strip()

                    if "B2B Marka" in grup.columns:
                        b2b_marka = str(grup["B2B Marka"].iloc[0]).strip()

                    kirac_entes_m_tekrari = (
                        web_kodu.upper().startswith("M")
                        and metin_norm(b2b_marka) == "ENTES"
                    )

                    if (
                        kirac_entes_m_tekrari
                        and "B2B Web Servis Kodu" in grup.columns
                    ):
                        dogrudan_m_satirlar = grup[
                            grup["B2B Web Servis Kodu"].apply(sade_kod)
                            == grup["Web Servis Kodu"].apply(sade_kod)
                        ].copy()

                        if not dogrudan_m_satirlar.empty:
                            secilecek_satirlar.append(
                                stok_sayisi_en_yuksek_satiri_sec(dogrudan_m_satirlar)
                            )
                            continue

                    secilecek_satirlar.append(
                        stok_sayisi_en_yuksek_satiri_sec(grup)
                    )

                if secilecek_satirlar:
                    kirac_stok_guncellenecekler = pd.concat(
                        secilecek_satirlar,
                        ignore_index=True,
                        sort=False
                    )
                else:
                    kirac_stok_guncellenecekler = kirac_stok_guncellenecekler.iloc[0:0].copy()

                print()
                print("KIRAC stok guncellenecekler icinde tekrar eden urunler tekillestirildi.")
                print("Toplama yapilmadi; T-Soft mevcut stok hesaba katilmadi.")
                print("Tekrar eden KIRAC + ENTES M kodlarinda sadece KIRAC'ta dogrudan M kodu olan satir kullanildi.")
                print("Dogrudan M kodu yoksa modelden cevrilen en yuksek stoklu satir kullanildi.")
                print("Tekrar eden satirlarda stok sayisi en yuksek olan satir secildi.")

    durum_df_listesi = []

    durum_df_listesi.append(
        durum_tablosu_ekle(
            kirac_eslesenler,
            "EŞLEŞTİ",
            ekstra_kolonlar=[
                "ProductId",
                "Web Servis Kodu",
                "Model",
                "Yeni Stok",
                "Tsoft Mevcut Stok",
                "Eslesme Bicimi",
                "Kullanılan Kod",
            ]
        )
    )

    durum_df_listesi.append(
        durum_tablosu_ekle(kirac_b2bde_var_sitede_yok, "EŞLEŞMEDİ")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(kirac_bos_kod, "BOŞ KOD")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(kirac_hatali_stok, "HATALI STOK")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(kirac_coklu, "ÇOKLU EŞLEŞME")
    )

    durum_df_listesi = [df for df in durum_df_listesi if df is not None and not df.empty]

    if durum_df_listesi:
        durum_df = pd.concat(durum_df_listesi, ignore_index=True, sort=False)
    else:
        durum_df = pd.DataFrame()

    if not kirac_all.empty:
        kirac_kod_kontrol = kirac_all.copy()

        if not durum_df.empty and "B2B Satır No" in kirac_kod_kontrol.columns and "B2B Satır No" in durum_df.columns:
            durum_df = durum_df.drop_duplicates(subset=["B2B Satır No"], keep="first")

            eklenecek_kolonlar = [
                kolon for kolon in durum_df.columns
                if kolon not in [
                    "B2B Kaynak Dosya",
                    "B2B Web Servis Kodu",
                ]
            ]

            kirac_kod_kontrol = kirac_kod_kontrol.merge(
                durum_df[eklenecek_kolonlar],
                on="B2B Satır No",
                how="left"
            )
        else:
            kirac_kod_kontrol["Eşleşme Durumu"] = "KONTROL EDİLEMEDİ"
            kirac_kod_kontrol["Neden"] = "B2B Satır No bulunamadı"

        if "Eşleşme Durumu" not in kirac_kod_kontrol.columns:
            kirac_kod_kontrol["Eşleşme Durumu"] = "KONTROL EDİLEMEDİ"

        kirac_kod_kontrol["Eşleşme Durumu"] = kirac_kod_kontrol["Eşleşme Durumu"].fillna(
            "KONTROL EDİLEMEDİ"
        )

        if "Neden" not in kirac_kod_kontrol.columns:
            kirac_kod_kontrol["Neden"] = ""

        kirac_kod_kontrol["Neden"] = kirac_kod_kontrol["Neden"].fillna("")
    else:
        kirac_kod_kontrol = pd.DataFrame()

    kirac_eslesenler = santral_kolonlarini_cikar(kirac_eslesenler)
    kirac_eslesmeyenler = santral_kolonlarini_cikar(kirac_eslesmeyenler)
    kirac_stok_guncellenecekler = santral_kolonlarini_cikar(kirac_stok_guncellenecekler)
    kirac_kod_kontrol = santral_kolonlarini_cikar(kirac_kod_kontrol)
    kirac_tsoftta_yok_b2b_stok_var = santral_kolonlarini_cikar(kirac_tsoftta_yok_b2b_stok_var)
    kirac_tsoftta_yok_b2b_stok_yok = santral_kolonlarini_cikar(kirac_tsoftta_yok_b2b_stok_yok)
    kirac_bos_urun_kodu = santral_kolonlarini_cikar(kirac_bos_urun_kodu)

    return {
        "kirac_eslesenler": kirac_eslesenler,
        "kirac_eslesmeyenler": kirac_eslesmeyenler,
        "kirac_stok_guncellenecekler": kirac_stok_guncellenecekler,
        "kirac_kod_kontrol": kirac_kod_kontrol,
        "kirac_tsoftta_yok_b2b_stok_var": kirac_tsoftta_yok_b2b_stok_var,
        "kirac_tsoftta_yok_b2b_stok_yok": kirac_tsoftta_yok_b2b_stok_yok,
        "kirac_bos_urun_kodu": kirac_bos_urun_kodu,
    }



def santral_raporlari_olustur(
    b2b_all,
    tedarikci_bazli_df,
    b2bde_var_sitede_yok_df,
    bos_kod_df,
    hatali_stok_df,
    coklu_df
):
    """
    SANTRAL için özel kontrol dosyaları oluşturur:
    1- santral_eslesenler.xlsx
    2- santral_eslesmeyenler.xlsx
    3- santral_stok_guncellenecekler.xlsx
    4- santral_kod_kontrol.xlsx
    5- santral_tsoftta_yok_b2b_stok_var.xlsx
    6- santral_tsoftta_yok_b2b_stok_yok.xlsx
    7- santral_bos_urun_kodu.xlsx
    """

    santral_all = tedarikciye_gore_filtrele(b2b_all, "SANTRAL")
    santral_eslesenler = tedarikciye_gore_filtrele(tedarikci_bazli_df, "SANTRAL")
    santral_eslesenler = santral_eslesen_filtre_ekle(santral_eslesenler)

    santral_b2bde_var_sitede_yok = tedarikciye_gore_filtrele(
        b2bde_var_sitede_yok_df,
        "SANTRAL"
    )
    santral_bos_kod = tedarikciye_gore_filtrele(bos_kod_df, "SANTRAL")
    santral_hatali_stok = tedarikciye_gore_filtrele(hatali_stok_df, "SANTRAL")
    santral_coklu = tedarikciye_gore_filtrele(coklu_df, "SANTRAL")

    santral_eslesmeyenler_listesi = [
        df for df in [
            santral_b2bde_var_sitede_yok,
            santral_bos_kod,
            santral_hatali_stok,
            santral_coklu,
        ]
        if df is not None and not df.empty
    ]

    if santral_eslesmeyenler_listesi:
        santral_eslesmeyenler = pd.concat(
            santral_eslesmeyenler_listesi,
            ignore_index=True,
            sort=False
        )
        santral_eslesmeyenler = gercek_eslesmeyenleri_filtrele(santral_eslesmeyenler)
    else:
        santral_eslesmeyenler = pd.DataFrame()

    # SANTRAL eşleşmeyenlerini daha net ayırıyoruz:
    # 1- T-Soft'ta yok, B2B'de stok var
    # 2- T-Soft'ta yok, B2B'de stok yok
    # 3- B2B ürün kodu boş
    santral_bos_urun_kodu = santral_bos_kod.copy()

    if not santral_b2bde_var_sitede_yok.empty and "Yeni Stok" in santral_b2bde_var_sitede_yok.columns:
        santral_tsoftta_yok_df = santral_b2bde_var_sitede_yok.copy()
        santral_tsoftta_yok_df["Yeni Stok Sayı"] = pd.to_numeric(
            santral_tsoftta_yok_df["Yeni Stok"],
            errors="coerce"
        ).fillna(0).astype(int)

        santral_tsoftta_yok_b2b_stok_var = santral_tsoftta_yok_df[
            santral_tsoftta_yok_df["Yeni Stok Sayı"] > 0
        ].copy()

        santral_tsoftta_yok_b2b_stok_yok = santral_tsoftta_yok_df[
            santral_tsoftta_yok_df["Yeni Stok Sayı"] <= 0
        ].copy()

        for df in [
            santral_tsoftta_yok_b2b_stok_var,
            santral_tsoftta_yok_b2b_stok_yok,
        ]:
            df.drop(
                columns=["Yeni Stok Sayı"],
                inplace=True,
                errors="ignore"
            )
    else:
        santral_tsoftta_yok_b2b_stok_var = pd.DataFrame()
        santral_tsoftta_yok_b2b_stok_yok = pd.DataFrame()

    santral_stok_guncellenecekler = guncelleme_kaynagi_kurallarini_uygula(
        santral_eslesenler
    )
    santral_stok_guncellenecekler = santral_tekrarlarda_en_yuksek_stogu_sec(
        santral_stok_guncellenecekler
    )

    durum_df_listesi = []

    durum_df_listesi.append(
        durum_tablosu_ekle(
            santral_eslesenler,
            "EŞLEŞTİ",
            ekstra_kolonlar=[
                "ProductId",
                "Web Servis Kodu",
                "Model",
                "Yeni Stok",
                "Tsoft Mevcut Stok",
                "Eslesme Bicimi",
                "Kullanılan Kod",
            ]
        )
    )

    durum_df_listesi.append(
        durum_tablosu_ekle(santral_b2bde_var_sitede_yok, "EŞLEŞMEDİ")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(santral_bos_kod, "BOŞ KOD")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(santral_hatali_stok, "HATALI STOK")
    )
    durum_df_listesi.append(
        durum_tablosu_ekle(santral_coklu, "ÇOKLU EŞLEŞME")
    )

    durum_df_listesi = [df for df in durum_df_listesi if df is not None and not df.empty]

    if durum_df_listesi:
        durum_df = pd.concat(durum_df_listesi, ignore_index=True, sort=False)
    else:
        durum_df = pd.DataFrame()

    if not santral_all.empty:
        santral_kod_kontrol = santral_all.copy()

        if not durum_df.empty and "B2B Satır No" in santral_kod_kontrol.columns and "B2B Satır No" in durum_df.columns:
            durum_df = durum_df.drop_duplicates(subset=["B2B Satır No"], keep="first")

            eklenecek_kolonlar = [
                kolon for kolon in durum_df.columns
                if kolon not in [
                    "B2B Kaynak Dosya",
                    "B2B Web Servis Kodu",
                ]
            ]

            santral_kod_kontrol = santral_kod_kontrol.merge(
                durum_df[eklenecek_kolonlar],
                on="B2B Satır No",
                how="left"
            )
        else:
            santral_kod_kontrol["Eşleşme Durumu"] = "KONTROL EDİLEMEDİ"
            santral_kod_kontrol["Neden"] = "B2B Satır No bulunamadı"

        if "Eşleşme Durumu" not in santral_kod_kontrol.columns:
            santral_kod_kontrol["Eşleşme Durumu"] = "KONTROL EDİLEMEDİ"

        santral_kod_kontrol["Eşleşme Durumu"] = santral_kod_kontrol["Eşleşme Durumu"].fillna(
            "KONTROL EDİLEMEDİ"
        )

        if "Neden" not in santral_kod_kontrol.columns:
            santral_kod_kontrol["Neden"] = ""

        santral_kod_kontrol["Neden"] = santral_kod_kontrol["Neden"].fillna("")
    else:
        santral_kod_kontrol = pd.DataFrame()

    return {
        "santral_eslesenler": santral_eslesenler,
        "santral_eslesmeyenler": santral_eslesmeyenler,
        "santral_stok_guncellenecekler": santral_stok_guncellenecekler,
        "santral_kod_kontrol": santral_kod_kontrol,
        "santral_tsoftta_yok_b2b_stok_var": santral_tsoftta_yok_b2b_stok_var,
        "santral_tsoftta_yok_b2b_stok_yok": santral_tsoftta_yok_b2b_stok_yok,
        "santral_bos_urun_kodu": santral_bos_urun_kodu,
    }


def excel_kaydet_guvenli(df, dosya, index=False):
    """
    Excel dosyası açık kaldığında üzerine yazılamaz.
    Böyle bir durumda aynı dosyayı tarih-saat ekleyerek kaydeder.
    """
    try:
        df.to_excel(dosya, index=index)
        return dosya
    except PermissionError:
        tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
        yeni_dosya = dosya.with_name(f"{dosya.stem}_{tarih}{dosya.suffix}")

        df.to_excel(yeni_dosya, index=index)

        print()
        print("UYARI: Dosya açık olduğu için üzerine yazılamadı:")
        print(dosya)
        print("Yeni dosya olarak kaydedildi:")
        print(yeni_dosya)

        return yeni_dosya


def excel_coklu_sayfa_kaydet_guvenli(dosya, sayfalar):
    """
    Çok sayfalı Excel dosyasını güvenli kaydeder.
    kontrol_raporu.xlsx açık kalırsa tarih-saatli yeni dosya oluşturur.
    """
    try:
        with pd.ExcelWriter(dosya, engine="openpyxl") as writer:
            for sayfa_adi, df in sayfalar:
                df.to_excel(writer, sheet_name=sayfa_adi, index=False)

        return dosya

    except PermissionError:
        tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
        yeni_dosya = dosya.with_name(f"{dosya.stem}_{tarih}{dosya.suffix}")

        with pd.ExcelWriter(yeni_dosya, engine="openpyxl") as writer:
            for sayfa_adi, df in sayfalar:
                df.to_excel(writer, sheet_name=sayfa_adi, index=False)

        print()
        print("UYARI: Çok sayfalı Excel dosyası açık olduğu için üzerine yazılamadı:")
        print(dosya)
        print("Yeni dosya olarak kaydedildi:")
        print(yeni_dosya)

        return yeni_dosya


def marka_firma_stok_ozeti_olustur(tedarikci_bazli_df):
    kolonlar = [
        "MARKA",
        "KIRAC STOK",
        "KONYA STOK",
        "SANTRAL STOK",
        "MET STOK",
        "KLTSHOP STOK",
        "TOPLAM BAYİ STOK",
    ]

    if tedarikci_bazli_df.empty:
        return pd.DataFrame(columns=kolonlar)

    gerekli = {"Tedarikci", "B2B Marka", "ProductId", "Tsoft Mevcut Stok", "Yeni Stok"}
    if not gerekli.issubset(tedarikci_bazli_df.columns):
        return pd.DataFrame(columns=kolonlar)

    kaynak = tedarikci_bazli_df.copy()
    kaynak["FİRMA"] = kaynak["Tedarikci"].astype(str).str.strip().str.upper()
    kaynak["MARKA"] = kaynak["B2B Marka"].apply(marka_adi_standart)
    kaynak["BAYİ STOK"] = kaynak["Yeni Stok"].apply(sayiya_cevir)

    # Aynı ürün bayi dosyasında birden fazla satırda bulunabilir. T-Soft stoku
    # ürün başına bir kez, bayi stoku ise ilgili satırların toplamı olarak alınır.
    urun_bazli = (
        kaynak
        .groupby(["MARKA", "FİRMA", "ProductId"], dropna=False, as_index=False)
        .agg({
            "BAYİ STOK": "sum",
        })
    )

    bayi_ozet = (
        urun_bazli
        .groupby(["MARKA", "FİRMA"], dropna=False, as_index=False)["BAYİ STOK"]
        .sum()
        .pivot(index="MARKA", columns="FİRMA", values="BAYİ STOK")
        .fillna(0)
        .reset_index()
    )
    bayi_ozet.columns.name = None

    for firma in ["KIRAC", "KONYA", "SANTRAL", "MET", "KLTSHOP"]:
        if firma not in bayi_ozet.columns:
            bayi_ozet[firma] = 0
        bayi_ozet = bayi_ozet.rename(columns={firma: f"{firma} STOK"})

    ozet = (
        bayi_ozet
        .fillna(0)
        .sort_values("MARKA", kind="stable")
        .reset_index(drop=True)
    )
    ozet["TOPLAM BAYİ STOK"] = (
        ozet["KIRAC STOK"]
        + ozet["KONYA STOK"]
        + ozet["SANTRAL STOK"]
        + ozet["MET STOK"]
        + ozet["KLTSHOP STOK"]
    )
    for kolon in kolonlar[1:]:
        ozet[kolon] = pd.to_numeric(ozet[kolon], errors="coerce").fillna(0).astype(int)

    return ozet[kolonlar]


def tum_bayiler_tsoftta_yok_urunleri_olustur(b2bde_var_sitede_yok_df):
    """TSoft'ta olmayan urunleri yedi bayi icin tek satirda birlestirir."""
    firmalar = [
        "KIRAC",
        "KONYA",
        "SANTRAL",
        "MET",
        "KLTSHOP",
        "MEANWELL",
        "YILDIZ",
    ]
    sonuc_kolonlari = [
        "Marka",
        "Urun / Model Kodu",
        *firmalar,
        "Toplam Stok",
        "Firma Sayisi",
    ]

    if b2bde_var_sitede_yok_df is None or b2bde_var_sitede_yok_df.empty:
        return pd.DataFrame(columns=sonuc_kolonlari)

    gerekli_kolonlar = {
        "Tedarikci",
        "B2B Web Servis Kodu",
        "B2B Marka",
        "Yeni Stok",
    }
    if not gerekli_kolonlar.issubset(b2bde_var_sitede_yok_df.columns):
        eksikler = sorted(gerekli_kolonlar - set(b2bde_var_sitede_yok_df.columns))
        print(
            "UYARI: TSoft'ta yok ortak urun raporu olusturulamadi; "
            f"eksik kolonlar: {', '.join(eksikler)}"
        )
        return pd.DataFrame(columns=sonuc_kolonlari)

    kaynak = b2bde_var_sitede_yok_df.copy()
    kaynak["Firma"] = kaynak["Tedarikci"].apply(metin_norm)
    kaynak = kaynak[kaynak["Firma"].isin(firmalar)].copy()
    kaynak["Marka"] = kaynak["B2B Marka"].apply(marka_adi_standart)
    kaynak["Marka"] = kaynak["Marka"].replace("", "BILINMIYOR")
    kaynak["Urun / Model Kodu"] = (
        kaynak["B2B Web Servis Kodu"].fillna("").astype(str).str.strip()
    )
    kaynak["Kod Anahtar"] = kaynak["Urun / Model Kodu"].apply(metin_norm)
    kaynak["Stok"] = kaynak["Yeni Stok"].apply(sayiya_cevir)
    kaynak = kaynak[kaynak["Kod Anahtar"].ne("")].copy()

    if kaynak.empty:
        return pd.DataFrame(columns=sonuc_kolonlari)

    # Ayni urun ayni bayi dosyasinda birden fazla kez gelirse stok toplanmaz.
    # Bayinin o urun icin bildirdigi en yuksek stok degeri kullanilir.
    firma_stok = (
        kaynak
        .groupby(["Marka", "Kod Anahtar", "Firma"], as_index=False)["Stok"]
        .max()
    )
    kod_gorunumu = (
        kaynak
        .groupby(["Marka", "Kod Anahtar"], as_index=False)["Urun / Model Kodu"]
        .first()
    )
    urunler = (
        firma_stok
        .pivot_table(
            index=["Marka", "Kod Anahtar"],
            columns="Firma",
            values="Stok",
            aggfunc="max",
            fill_value=0,
        )
        .reset_index()
    )
    urunler.columns.name = None
    urunler = urunler.merge(
        kod_gorunumu,
        on=["Marka", "Kod Anahtar"],
        how="left",
    )

    for firma in firmalar:
        if firma not in urunler.columns:
            urunler[firma] = 0
        urunler[firma] = (
            pd.to_numeric(urunler[firma], errors="coerce").fillna(0).astype(int)
        )

    urunler["Toplam Stok"] = urunler[firmalar].sum(axis=1).astype(int)
    urunler["Firma Sayisi"] = (urunler[firmalar] > 0).sum(axis=1).astype(int)

    urunler = (
        urunler[sonuc_kolonlari]
        .sort_values(["Marka", "Urun / Model Kodu"], kind="stable")
        .reset_index(drop=True)
    )
    return urunler


def duzenli_raporlari_kaydet(
    tsoft_df,
    tedarikci_bazli_df,
    b2bde_var_sitede_yok_df,
    bos_kod_df,
    hatali_stok_df,
    coklu_df,
    konya_raporlar,
    kirac_raporlar,
    santral_raporlar,
):
    def genel_eslesmeyenler_olustur(tedarikci):
        parcalar = [
            tedarikciye_gore_filtrele(b2bde_var_sitede_yok_df, tedarikci),
            tedarikciye_gore_filtrele(bos_kod_df, tedarikci),
            tedarikciye_gore_filtrele(hatali_stok_df, tedarikci),
            tedarikciye_gore_filtrele(coklu_df, tedarikci),
        ]
        parcalar = [df for df in parcalar if df is not None and not df.empty]

        if not parcalar:
            return pd.DataFrame()

        return pd.concat(parcalar, ignore_index=True, sort=False)

    def kod_eslesmeyenleri_olustur(tedarikci):
        """Yalnizca kodu T-Soft urun listesinde bulunmayan satirlari getirir."""
        return tedarikciye_gore_filtrele(
            b2bde_var_sitede_yok_df,
            tedarikci,
        )

    def tsoftta_yok_stoga_gore_ayir(tedarikci):
        """T-Soft'ta olmayan bayi satirlarini stok var/yok olarak ayirir."""
        tsoftta_yok = kod_eslesmeyenleri_olustur(tedarikci)

        if "Yeni Stok" not in tsoftta_yok.columns:
            return tsoftta_yok.copy(), tsoftta_yok.copy()

        stok_sayisi = pd.to_numeric(
            tsoftta_yok["Yeni Stok"],
            errors="coerce",
        ).fillna(0)

        stok_var = tsoftta_yok[stok_sayisi > 0].copy()
        stok_yok = tsoftta_yok[stok_sayisi <= 0].copy()
        return stok_var, stok_yok

    def rapor_kolonlarini_sadelestir(rapor_df, firma, dosya_adi):
        if rapor_df.empty:
            return rapor_df.copy()

        her_zaman_silinecekler = [
            "Model",
            "Tsoft Mevcut Stok",
            "EslesmeTipi",
            "B2B Kaynak Dosya",
            "B2B Temiz Kod",
            "B2B Sade Kod",
            "Eslesmede Kullanilacak Kod",
            "Eslesme Bicimi",
            "Kullanılan Kod",
        ]

        teknik_kolonlar = her_zaman_silinecekler + [
            kolon
            for kolon in rapor_df.columns
            if metin_norm(kolon).startswith("KIRAC ")
        ]

        santral_eslesmeyen_raporu = (
            firma == "SANTRAL"
            and dosya_adi == "santral_eslesmeyenler.xlsx"
        )
        if not santral_eslesmeyen_raporu:
            teknik_kolonlar.extend([
                "SANTRAL Orijinal Ürün Kodu",
                "SANTRAL Kod Düzeltme Durumu",
                "SANTRAL Kod Düzeltme Notu",
            ])

        for bos_ise_silinecek in ["EM Kodu", "Manuel Tsoft Web Servis Kodu"]:
            if bos_ise_silinecek not in rapor_df.columns:
                continue

            dolu_mu = (
                rapor_df[bos_ise_silinecek]
                .astype(str)
                .str.strip()
                .replace("nan", "")
                .ne("")
                .any()
            )
            if not dolu_mu:
                teknik_kolonlar.append(bos_ise_silinecek)

        sade_df = rapor_df.drop(columns=teknik_kolonlar, errors="ignore")

        if dosya_adi.endswith("_stok_guncellenecekler.xlsx"):
            stok_guncelleme_kolonlari = [
                "Tedarikci",
                "B2B Satır No",
                "ProductId",
                "Web Servis Kodu",
                "Yeni Stok",
                "Stok Birimi",
                "Fiyat Birimi",
                "Satış Fiyatı",
                "B2B Web Servis Kodu",
                "B2B Stok Metni",
                "B2B Marka",
                "Guncelleme Durumu",
            ]
            return sade_df[
                [kolon for kolon in stok_guncelleme_kolonlari if kolon in sade_df.columns]
            ].copy()

        if dosya_adi.endswith("_eslesmeyenler.xlsx"):
            sade_df = sade_df.drop(columns=["Yeni Stok"], errors="ignore")
            if "B2B Marka" in sade_df.columns:
                sade_df = sade_df.rename(columns={"B2B Marka": "Marka"})
            elif "Marka" not in sade_df.columns:
                sade_df["Marka"] = ""

        korunacak_kolonlar = {
            "Tedarikci",
            "B2B Satır No",
            "ProductId",
            "Web Servis Kodu",
            "Yeni Stok",
            "Tsoft Mevcut Stok",
            "B2B Stok Metni",
            "Marka",
            "B2B Marka",
            "Neden",
        }
        tamamen_bos_kolonlar = [
            kolon
            for kolon in sade_df.columns
            if kolon not in korunacak_kolonlar
            and sade_df[kolon].astype(str).str.strip().replace("nan", "").eq("").all()
        ]
        sade_df = sade_df.drop(columns=tamamen_bos_kolonlar, errors="ignore")

        return sade_df

    # T-Soft'ta olmayan bayi urunleri stoklu ve stoksuz olarak ayri raporlanir.
    met_tsoftta_yok_stok_var, met_tsoftta_yok_stok_yok = tsoftta_yok_stoga_gore_ayir("MET")
    kltshop_tsoftta_yok_stok_var, kltshop_tsoftta_yok_stok_yok = tsoftta_yok_stoga_gore_ayir("KLTSHOP")
    meanwell_tsoftta_yok_stok_var, meanwell_tsoftta_yok_stok_yok = tsoftta_yok_stoga_gore_ayir("MEANWELL")
    yildiz_tsoftta_yok_stok_var, yildiz_tsoftta_yok_stok_yok = (
        tsoftta_yok_stoga_gore_ayir("YILDIZ")
    )

    bayi_raporlari = {
        "KONYA": konya_raporlar or {},
        "KIRAC": kirac_raporlar or {},
        "SANTRAL": santral_raporlar or {},
        "MET": {
            "met_eslesmeyenler": kod_eslesmeyenleri_olustur("MET"),
            "met_stok_guncellenecekler": guncelleme_kaynagi_kurallarini_uygula(
                tedarikciye_gore_filtrele(tedarikci_bazli_df, "MET")
            ),
            "met_tsoftta_yok_b2b_stok_var": met_tsoftta_yok_stok_var,
            "met_tsoftta_yok_b2b_stok_yok": met_tsoftta_yok_stok_yok,
        },
        "KLTSHOP": {
            "kltshop_eslesmeyenler": kod_eslesmeyenleri_olustur("KLTSHOP"),
            "kltshop_stok_guncellenecekler": guncelleme_kaynagi_kurallarini_uygula(
                tedarikciye_gore_filtrele(tedarikci_bazli_df, "KLTSHOP")
            ),
            "kltshop_tsoftta_yok_b2b_stok_var": kltshop_tsoftta_yok_stok_var,
            "kltshop_tsoftta_yok_b2b_stok_yok": kltshop_tsoftta_yok_stok_yok,
        },
        "MEANWELL": {
            "meanwell_eslesmeyenler": kod_eslesmeyenleri_olustur("MEANWELL"),
            "meanwell_stok_guncellenecekler": guncelleme_kaynagi_kurallarini_uygula(
                tedarikciye_gore_filtrele(tedarikci_bazli_df, "MEANWELL")
            ),
            "meanwell_tsoftta_yok_b2b_stok_var": meanwell_tsoftta_yok_stok_var,
            "meanwell_tsoftta_yok_b2b_stok_yok": meanwell_tsoftta_yok_stok_yok,
        },
        "YILDIZ": {
            "yildiz_eslesmeyenler": genel_eslesmeyenler_olustur("YILDIZ"),
            "yildiz_stok_guncellenecekler": guncelleme_kaynagi_kurallarini_uygula(
                tedarikciye_gore_filtrele(tedarikci_bazli_df, "YILDIZ")
            ),
            "yildiz_tsoftta_yok_b2b_stok_var": yildiz_tsoftta_yok_stok_var,
            "yildiz_tsoftta_yok_b2b_stok_yok": yildiz_tsoftta_yok_stok_yok,
        },
    }

    for firma, raporlar in bayi_raporlari.items():
        firma_kucuk = firma.lower()
        firma_dir = CIKTI_DIR / firma
        firma_dir.mkdir(parents=True, exist_ok=True)

        kaydedilecekler = [
            (
                f"{firma_kucuk}_eslesmeyenler.xlsx",
                raporlar.get(f"{firma_kucuk}_eslesmeyenler", pd.DataFrame()),
            ),
            (
                f"{firma_kucuk}_stok_guncellenecekler.xlsx",
                raporlar.get(f"{firma_kucuk}_stok_guncellenecekler", pd.DataFrame()),
            ),
            (
                f"{firma_kucuk}_tsoftta_yok_bayide_var_siteye_yuklenecekler.xlsx",
                raporlar.get(
                    f"{firma_kucuk}_tsoftta_yok_b2b_stok_var",
                    pd.DataFrame(),
                ),
            ),
            (
                f"{firma_kucuk}_tsoftta_yok_bayide_var_stok_yok.xlsx",
                raporlar.get(
                    f"{firma_kucuk}_tsoftta_yok_b2b_stok_yok",
                    pd.DataFrame(),
                ),
            ),
        ]

        for dosya_adi, rapor_df in kaydedilecekler:
            sade_rapor_df = rapor_kolonlarini_sadelestir(
                rapor_df,
                firma,
                dosya_adi,
            )
            excel_kaydet_guvenli(sade_rapor_df, firma_dir / dosya_adi)

    tum_bayiler_tsoftta_yok = tum_bayiler_tsoftta_yok_urunleri_olustur(
        b2bde_var_sitede_yok_df
    )
    excel_kaydet_guvenli(
        tum_bayiler_tsoftta_yok,
        CIKTI_DIR / "tum_bayiler_tsoftta_yok_b2bde_var.xlsx",
    )

    ortak_ozet = marka_firma_stok_ozeti_olustur(tedarikci_bazli_df)
    excel_kaydet_guvenli(
        ortak_ozet,
        CIKTI_DIR / "tum_bayiler_marka_stok_ozeti.xlsx",
    )

    guncellemeye_uygun_df = guncelleme_kaynagi_kurallarini_uygula(
        tedarikci_bazli_df
    )
    tum_bayiler_guncelleme = hazir_stok_olustur(guncellemeye_uygun_df)
    tum_bayiler_guncelleme = marka_stok_kaynaklarini_uygula(
        tum_bayiler_guncelleme,
        tsoft_df,
    )
    marka_uyusmazliklari = tum_bayiler_guncelleme[
        tum_bayiler_guncelleme["Marka Uyuşmazlığı"] == "EVET"
    ].copy()
    tum_bayiler_guncelleme = tum_bayiler_guncelleme[
        tum_bayiler_guncelleme["Marka Uyuşmazlığı"] != "EVET"
    ].copy()

    excel_kaydet_guvenli(
        marka_uyusmazliklari,
        CIKTI_DIR / "tum_bayiler_marka_uyusmazliklari.xlsx",
    )

    tum_bayiler_guncelleme = tum_bayiler_guncelleme.drop(
        columns=["Tsoft Mevcut Stok", "Marka Uyuşmazlığı"],
        errors="ignore",
    )
    excel_kaydet_guvenli(
        tum_bayiler_guncelleme,
        CIKTI_DIR / "tum_bayiler_stok_guncellenecekler.xlsx",
    )

    excel_kaydet_guvenli(bos_kod_df, LOG_DIR / "bos_webservis_kodlari.xlsx")
    excel_kaydet_guvenli(hatali_stok_df, LOG_DIR / "hatali_stok_satirlari.xlsx")
    excel_kaydet_guvenli(coklu_df, LOG_DIR / "coklu_eslesmeler.xlsx")

    print()
    print("Çıktılar oluşturuldu.")
    print(f"Bayi klasörleri: {CIKTI_DIR}")
    print("TSoft'ta yok ortak ürün listesi: tum_bayiler_tsoftta_yok_b2bde_var.xlsx")
    print(
        "TSoft'ta yok ortak ürün sayısı: "
        f"{len(tum_bayiler_tsoftta_yok)}"
    )
    print("Ortak rapor: tum_bayiler_marka_stok_ozeti.xlsx")
    print("Canlı güncelleme listesi: tum_bayiler_stok_guncellenecekler.xlsx")
    print(f"Marka uyuşmazlığı: {len(marka_uyusmazliklari)}")


def raporlari_kaydet(
    tsoft_df,
    hazir_df,
    tedarikci_bazli_df,
    b2bde_var_sitede_yok_df,
    santral_eslesme_bekleyen_df,
    bos_kod_df,
    hatali_stok_df,
    coklu_df,
    ozet_df,
    konya_raporlar=None,
    kirac_raporlar=None,
    santral_raporlar=None
):
    duzenli_raporlari_kaydet(
        tsoft_df=tsoft_df,
        tedarikci_bazli_df=tedarikci_bazli_df,
        b2bde_var_sitede_yok_df=b2bde_var_sitede_yok_df,
        bos_kod_df=bos_kod_df,
        hatali_stok_df=hatali_stok_df,
        coklu_df=coklu_df,
        konya_raporlar=konya_raporlar,
        kirac_raporlar=kirac_raporlar,
        santral_raporlar=santral_raporlar,
    )

    # Fiyatlandırma sistemi stok özetini değil, her tedarikçinin eşleştiği
    # gerçek B2B kaynak dosyasını ve satırını kullanır. Bu yüzden güncel rapor
    # akışında fiyatlandırmaya özel, sade bir eşleşme dosyası ayrıca üretilir.
    fiyatlandirma_kolonlari = [
        "ProductId",
        "Web Servis Kodu",
        "B2B Marka",
        "Tedarikci",
        "B2B Kaynak Dosya",
        "B2B Satır No",
    ]
    fiyatlandirma_eslesme_df = tedarikci_bazli_df.reindex(
        columns=fiyatlandirma_kolonlari
    ).copy()

    # B2B fiyat eşleşmesi bulunmayan fakat DIA/SabitStok kuralıyla stok
    # güncellemesine giden EMAS ve ORTAÇ ürünlerini kendi fiyat dosyalarından
    # fiyatlandırma akışına ekle. Kaynak dosya ve satır bilgisi korunur; liste
    # veya iskonto hataları fiyatlandırma raporunun Hatalar sayfasında görünür.
    stok_guncelleme_dosya = CIKTI_DIR / "tum_bayiler_stok_guncellenecekler.xlsx"
    sabit_fiyat_satirlari = []
    if stok_guncelleme_dosya.exists():
        stok_guncelleme_df = pd.read_excel(
            stok_guncelleme_dosya, dtype=str
        ).fillna("")
        mevcut_fiyat_pidleri = set(
            fiyatlandirma_eslesme_df["ProductId"].astype(str).str.strip()
        )
        sabit_hedefler = stok_guncelleme_df[
            stok_guncelleme_df["Guncelleme Durumu"]
            .astype(str)
            .str.strip()
            .eq("GONDERILECEK")
        ].copy()
        sabit_hedefler["Marka Norm"] = sabit_hedefler["B2B Marka"].apply(metin_norm)
        sabit_hedefler["Temiz Kod"] = sabit_hedefler["Web Servis Kodu"].apply(
            kod_temizle
        )
        sabit_hedefler = sabit_hedefler[
            sabit_hedefler["Marka Norm"].isin({"WAGO", "EMAS", "ORTAC"})
            & ~sabit_hedefler["ProductId"].astype(str).str.strip().isin(
                mevcut_fiyat_pidleri
            )
        ].copy()
        hedef_haritasi = {
            (row["Marka Norm"], row["Temiz Kod"]): row
            for _, row in sabit_hedefler.iterrows()
            if row["Temiz Kod"]
        }

        for marka_norm, tedarikci_adi in [
            ("WAGO", "WAGO SABİT"),
            ("EMAS", "EMAS SABİT"),
            ("ORTAC", "ORTAÇ SABİT"),
        ]:
            ayar = SABIT_STOK_DOSYALARI[marka_norm]
            dosya = ayar["dosya"]
            kaynak_df = pd.read_excel(dosya, dtype=str).fillna("")
            kod_kolonu = kolon_bul_esnek(kaynak_df, ayar["kod_kolonlari"])
            if not kod_kolonu:
                raise ValueError(
                    f"{dosya.name} içinde fiyatlandırma için ürün kodu bulunamadı"
                )
            for kaynak_index, kaynak_row in kaynak_df.iterrows():
                temiz_kod = kod_temizle(kaynak_row.get(kod_kolonu, ""))
                hedef = hedef_haritasi.get((marka_norm, temiz_kod))
                if hedef is None:
                    continue
                sabit_fiyat_satirlari.append(
                    {
                        "ProductId": str(hedef["ProductId"]).strip(),
                        "Web Servis Kodu": str(hedef["Web Servis Kodu"]).strip(),
                        "B2B Marka": str(hedef["B2B Marka"]).strip(),
                        "Tedarikci": tedarikci_adi,
                        "B2B Kaynak Dosya": str(dosya),
                        "B2B Satır No": int(kaynak_index) + 2,
                    }
                )

    if sabit_fiyat_satirlari:
        fiyatlandirma_eslesme_df = pd.concat(
            [
                fiyatlandirma_eslesme_df,
                pd.DataFrame(sabit_fiyat_satirlari),
            ],
            ignore_index=True,
        )
        fiyatlandirma_eslesme_df = fiyatlandirma_eslesme_df.drop_duplicates(
            subset=["ProductId", "Tedarikci", "B2B Kaynak Dosya", "B2B Satır No"],
            keep="first",
        )
        print(
            "SabitStok fiyatlandırma eşleşmesi eklendi: "
            f"{len(sabit_fiyat_satirlari)}"
        )

    fiyatlandirma_eslesme_dosya = excel_kaydet_guvenli(
        fiyatlandirma_eslesme_df,
        CIKTI_DIR / "fiyatlandirma_eslesme_listesi.xlsx",
    )
    print(
        "Fiyatlandırma eşleşme listesi: "
        f"{fiyatlandirma_eslesme_dosya}"
    )
    return

    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")

    hazir_dosya = CIKTI_DIR / "hazir_stok_listesi.xlsx"
    tedarikci_bazli_dosya = CIKTI_DIR / "tedarikci_bazli_stok_listesi.xlsx"
    b2bde_var_sitede_yok_dosya = CIKTI_DIR / "b2bde_var_sitede_yok.xlsx"
    santral_eslesme_bekleyen_dosya = CIKTI_DIR / "santral_eslesme_bekleyenler.xlsx"

    konya_eslesenler_dosya = CIKTI_DIR / "konya_eslesenler.xlsx"
    konya_eslesmeyenler_dosya = CIKTI_DIR / "konya_eslesmeyenler.xlsx"
    konya_stok_guncellenecekler_dosya = CIKTI_DIR / "konya_stok_guncellenecekler.xlsx"
    konya_kod_kontrol_dosya = CIKTI_DIR / "konya_kod_kontrol.xlsx"
    konya_tsoftta_yok_b2b_stok_var_dosya = CIKTI_DIR / "konya_tsoftta_yok_b2b_stok_var.xlsx"
    konya_tsoftta_yok_b2b_stok_yok_dosya = CIKTI_DIR / "konya_tsoftta_yok_b2b_stok_yok.xlsx"
    konya_bos_urun_kodu_dosya = CIKTI_DIR / "konya_bos_urun_kodu.xlsx"

    kirac_eslesenler_dosya = CIKTI_DIR / "kirac_eslesenler.xlsx"
    kirac_eslesmeyenler_dosya = CIKTI_DIR / "kirac_eslesmeyenler.xlsx"
    kirac_stok_guncellenecekler_dosya = CIKTI_DIR / "kirac_stok_guncellenecekler.xlsx"
    kirac_kod_kontrol_dosya = CIKTI_DIR / "kirac_kod_kontrol.xlsx"
    kirac_tsoftta_yok_b2b_stok_var_dosya = CIKTI_DIR / "kirac_tsoftta_yok_b2b_stok_var.xlsx"
    kirac_tsoftta_yok_b2b_stok_yok_dosya = CIKTI_DIR / "kirac_tsoftta_yok_b2b_stok_yok.xlsx"
    kirac_bos_urun_kodu_dosya = CIKTI_DIR / "kirac_bos_urun_kodu.xlsx"

    santral_eslesenler_dosya = CIKTI_DIR / "santral_eslesenler.xlsx"
    santral_eslesmeyenler_dosya = CIKTI_DIR / "santral_eslesmeyenler.xlsx"
    santral_stok_guncellenecekler_dosya = CIKTI_DIR / "santral_stok_guncellenecekler.xlsx"
    santral_kod_kontrol_dosya = CIKTI_DIR / "santral_kod_kontrol.xlsx"
    santral_tsoftta_yok_b2b_stok_var_dosya = CIKTI_DIR / "santral_tsoftta_yok_b2b_stok_var.xlsx"
    santral_tsoftta_yok_b2b_stok_yok_dosya = CIKTI_DIR / "santral_tsoftta_yok_b2b_stok_yok.xlsx"
    santral_bos_urun_kodu_dosya = CIKTI_DIR / "santral_bos_urun_kodu.xlsx"

    kontrol_dosya = CIKTI_DIR / "kontrol_raporu.xlsx"
    tarihli_kontrol_dosya = CIKTI_DIR / f"kontrol_raporu_{tarih}.xlsx"

    bos_kod_dosya = LOG_DIR / "bos_webservis_kodlari.xlsx"
    hatali_stok_dosya = LOG_DIR / "hatali_stok_satirlari.xlsx"
    coklu_dosya = LOG_DIR / "coklu_eslesmeler.xlsx"

    hazir_dosya = excel_kaydet_guvenli(hazir_df, hazir_dosya)
    tedarikci_bazli_dosya = excel_kaydet_guvenli(tedarikci_bazli_df, tedarikci_bazli_dosya)
    b2bde_var_sitede_yok_dosya = excel_kaydet_guvenli(b2bde_var_sitede_yok_df, b2bde_var_sitede_yok_dosya)
    santral_eslesme_bekleyen_dosya = excel_kaydet_guvenli(santral_eslesme_bekleyen_df, santral_eslesme_bekleyen_dosya)

    if konya_raporlar is None:
        konya_raporlar = {}

    konya_eslesenler_df = konya_raporlar.get("konya_eslesenler", pd.DataFrame())
    konya_eslesmeyenler_df = konya_raporlar.get("konya_eslesmeyenler", pd.DataFrame())
    konya_stok_guncellenecekler_df = konya_raporlar.get("konya_stok_guncellenecekler", pd.DataFrame())
    konya_kod_kontrol_df = konya_raporlar.get("konya_kod_kontrol", pd.DataFrame())
    konya_tsoftta_yok_b2b_stok_var_df = konya_raporlar.get("konya_tsoftta_yok_b2b_stok_var", pd.DataFrame())
    konya_tsoftta_yok_b2b_stok_yok_df = konya_raporlar.get("konya_tsoftta_yok_b2b_stok_yok", pd.DataFrame())
    konya_bos_urun_kodu_df = konya_raporlar.get("konya_bos_urun_kodu", pd.DataFrame())

    if kirac_raporlar is None:
        kirac_raporlar = {}

    kirac_eslesenler_df = kirac_raporlar.get("kirac_eslesenler", pd.DataFrame())
    kirac_eslesmeyenler_df = kirac_raporlar.get("kirac_eslesmeyenler", pd.DataFrame())
    kirac_stok_guncellenecekler_df = kirac_raporlar.get("kirac_stok_guncellenecekler", pd.DataFrame())
    kirac_kod_kontrol_df = kirac_raporlar.get("kirac_kod_kontrol", pd.DataFrame())
    kirac_tsoftta_yok_b2b_stok_var_df = kirac_raporlar.get("kirac_tsoftta_yok_b2b_stok_var", pd.DataFrame())
    kirac_tsoftta_yok_b2b_stok_yok_df = kirac_raporlar.get("kirac_tsoftta_yok_b2b_stok_yok", pd.DataFrame())
    kirac_bos_urun_kodu_df = kirac_raporlar.get("kirac_bos_urun_kodu", pd.DataFrame())

    if santral_raporlar is None:
        santral_raporlar = {}

    santral_eslesenler_df = santral_raporlar.get("santral_eslesenler", pd.DataFrame())
    santral_eslesmeyenler_df = santral_raporlar.get("santral_eslesmeyenler", pd.DataFrame())
    santral_stok_guncellenecekler_df = santral_raporlar.get("santral_stok_guncellenecekler", pd.DataFrame())
    santral_kod_kontrol_df = santral_raporlar.get("santral_kod_kontrol", pd.DataFrame())
    santral_tsoftta_yok_b2b_stok_var_df = santral_raporlar.get("santral_tsoftta_yok_b2b_stok_var", pd.DataFrame())
    santral_tsoftta_yok_b2b_stok_yok_df = santral_raporlar.get("santral_tsoftta_yok_b2b_stok_yok", pd.DataFrame())
    santral_bos_urun_kodu_df = santral_raporlar.get("santral_bos_urun_kodu", pd.DataFrame())

    konya_eslesenler_dosya = excel_kaydet_guvenli(konya_eslesenler_df, konya_eslesenler_dosya)
    konya_eslesmeyenler_dosya = excel_kaydet_guvenli(konya_eslesmeyenler_df, konya_eslesmeyenler_dosya)
    konya_stok_guncellenecekler_dosya = excel_kaydet_guvenli(konya_stok_guncellenecekler_df, konya_stok_guncellenecekler_dosya)
    konya_kod_kontrol_dosya = excel_kaydet_guvenli(konya_kod_kontrol_df, konya_kod_kontrol_dosya)
    konya_tsoftta_yok_b2b_stok_var_dosya = excel_kaydet_guvenli(konya_tsoftta_yok_b2b_stok_var_df, konya_tsoftta_yok_b2b_stok_var_dosya)
    konya_tsoftta_yok_b2b_stok_yok_dosya = excel_kaydet_guvenli(konya_tsoftta_yok_b2b_stok_yok_df, konya_tsoftta_yok_b2b_stok_yok_dosya)
    konya_bos_urun_kodu_dosya = excel_kaydet_guvenli(konya_bos_urun_kodu_df, konya_bos_urun_kodu_dosya)

    kirac_eslesenler_dosya = excel_kaydet_guvenli(kirac_eslesenler_df, kirac_eslesenler_dosya)
    kirac_eslesmeyenler_dosya = excel_kaydet_guvenli(kirac_eslesmeyenler_df, kirac_eslesmeyenler_dosya)
    kirac_stok_guncellenecekler_dosya = excel_kaydet_guvenli(kirac_stok_guncellenecekler_df, kirac_stok_guncellenecekler_dosya)
    kirac_kod_kontrol_dosya = excel_kaydet_guvenli(kirac_kod_kontrol_df, kirac_kod_kontrol_dosya)
    kirac_tsoftta_yok_b2b_stok_var_dosya = excel_kaydet_guvenli(kirac_tsoftta_yok_b2b_stok_var_df, kirac_tsoftta_yok_b2b_stok_var_dosya)
    kirac_tsoftta_yok_b2b_stok_yok_dosya = excel_kaydet_guvenli(kirac_tsoftta_yok_b2b_stok_yok_df, kirac_tsoftta_yok_b2b_stok_yok_dosya)
    kirac_bos_urun_kodu_dosya = excel_kaydet_guvenli(kirac_bos_urun_kodu_df, kirac_bos_urun_kodu_dosya)

    santral_eslesenler_dosya = excel_kaydet_guvenli(santral_eslesenler_df, santral_eslesenler_dosya)
    santral_eslesmeyenler_dosya = excel_kaydet_guvenli(santral_eslesmeyenler_df, santral_eslesmeyenler_dosya)
    santral_stok_guncellenecekler_dosya = excel_kaydet_guvenli(santral_stok_guncellenecekler_df, santral_stok_guncellenecekler_dosya)
    santral_kod_kontrol_dosya = excel_kaydet_guvenli(santral_kod_kontrol_df, santral_kod_kontrol_dosya)
    santral_tsoftta_yok_b2b_stok_var_dosya = excel_kaydet_guvenli(santral_tsoftta_yok_b2b_stok_var_df, santral_tsoftta_yok_b2b_stok_var_dosya)
    santral_tsoftta_yok_b2b_stok_yok_dosya = excel_kaydet_guvenli(santral_tsoftta_yok_b2b_stok_yok_df, santral_tsoftta_yok_b2b_stok_yok_dosya)
    santral_bos_urun_kodu_dosya = excel_kaydet_guvenli(santral_bos_urun_kodu_df, santral_bos_urun_kodu_dosya)

    bos_kod_dosya = excel_kaydet_guvenli(bos_kod_df, bos_kod_dosya)
    hatali_stok_dosya = excel_kaydet_guvenli(hatali_stok_df, hatali_stok_dosya)
    coklu_dosya = excel_kaydet_guvenli(coklu_df, coklu_dosya)

    kontrol_sayfalari = [
        ("Ozet", ozet_df),
        ("HazirStok", hazir_df),
        ("TedarikciBazli", tedarikci_bazli_df),
        ("B2BdeVarSitedeYok", b2bde_var_sitede_yok_df),
        ("SantralBekleyen", santral_eslesme_bekleyen_df),
        ("KonyaEslesen", konya_eslesenler_df),
        ("KonyaEslesmeyen", konya_eslesmeyenler_df),
        ("KonyaStokGuncelle", konya_stok_guncellenecekler_df),
        ("KonyaKodKontrol", konya_kod_kontrol_df),
        ("KonyaYokB2BStokVar", konya_tsoftta_yok_b2b_stok_var_df),
        ("KonyaYokB2BStokYok", konya_tsoftta_yok_b2b_stok_yok_df),
        ("KonyaBosUrunKodu", konya_bos_urun_kodu_df),
        ("KiracEslesen", kirac_eslesenler_df),
        ("KiracEslesmeyen", kirac_eslesmeyenler_df),
        ("KiracStokGuncelle", kirac_stok_guncellenecekler_df),
        ("KiracKodKontrol", kirac_kod_kontrol_df),
        ("KiracYokB2BStokVar", kirac_tsoftta_yok_b2b_stok_var_df),
        ("KiracYokB2BStokYok", kirac_tsoftta_yok_b2b_stok_yok_df),
        ("KiracBosUrunKodu", kirac_bos_urun_kodu_df),
        ("SantralEslesen", santral_eslesenler_df),
        ("SantralEslesmeyen", santral_eslesmeyenler_df),
        ("SantralStokGuncelle", santral_stok_guncellenecekler_df),
        ("SantralKodKontrol", santral_kod_kontrol_df),
        ("SantralYokB2BStokVar", santral_tsoftta_yok_b2b_stok_var_df),
        ("SantralYokB2BStokYok", santral_tsoftta_yok_b2b_stok_yok_df),
        ("SantralBosUrunKodu", santral_bos_urun_kodu_df),
        ("BosKod", bos_kod_df),
        ("HataliStok", hatali_stok_df),
        ("CokluEslesme", coklu_df),
    ]

    kontrol_dosya = excel_coklu_sayfa_kaydet_guvenli(kontrol_dosya, kontrol_sayfalari)
    tarihli_kontrol_dosya = excel_coklu_sayfa_kaydet_guvenli(tarihli_kontrol_dosya, kontrol_sayfalari)

    print()
    print("Çıktılar oluşturuldu.")
    print(f"Çıktı klasörü: {CIKTI_DIR}")
    print(f"Log klasörü: {LOG_DIR}")


def main():
    print("=== B2B STOK EŞLEŞTİRME BAŞLADI ===")
    print(f"Kod sürümü: {KOD_SURUMU}")
    print("Mod: KONTROL MODU")
    print("Bu işlem T-Soft'a stok göndermez.")
    print()
    print("Stok mantığı:")
    print("Stok güncellenecekler = T-Soft ile eşleşen satırlar")
    print()

    ayar_df = ayarlari_oku()
    santral_lookup = santral_eslesme_oku()
    tsoft_df = tsoft_oku()
    lookup = tsoft_lookup_olustur(tsoft_df)

    tum_b2b_list = []

    for _, ayar in ayar_df.iterrows():
        b2b_df = b2b_tedarikci_oku(ayar, santral_lookup)

        if not b2b_df.empty:
            tum_b2b_list.append(b2b_df)

    b2b_kolonlar = [
        "Tedarikci",
        "EslesmeTipi",
        "B2B Satır No",
        "B2B Kaynak Dosya",
        "B2B Web Servis Kodu",
        "B2B Stok Metni",
        "B2B Marka",
        "EM Kodu",
        "KIRAC Orijinal Ürün Kodu",
        "KIRAC Kod Düzeltme Durumu",
        "KIRAC Kod Düzeltme Notu",
        "SANTRAL Orijinal Ürün Kodu",
        "SANTRAL Kod Düzeltme Durumu",
        "SANTRAL Kod Düzeltme Notu",
        "B2B Temiz Kod",
        "B2B Sade Kod",
        "Yeni Stok",
        "Manuel Tsoft Web Servis Kodu",
        "Eslesmede Kullanilacak Kod",
        "KIRAC Legrand Orijinal Kod",
        "KIRAC Legrand Sıfırsız Kod",
        "KIRAC Legrand Sıfır Kuralı",
        "KIRAC Bosch Orijinal Kod",
        "KIRAC Bosch Boşluklu Kod",
        "KIRAC Bosch Boşluk Kuralı",
        "KIRAC Ortac Orijinal Kod",
        "KIRAC Ortac Boşluklu Kod",
        "KIRAC Ortac Boşluk Kuralı",
        "KIRAC Şafak Orijinal Kod",
        "KIRAC Şafak Boşluklu Kod",
        "KIRAC Şafak KM Kuralı",
        "KIRAC OZNUR Orijinal Kod",
        "KIRAC OZNUR Tsoft Kodu",
        "KIRAC OZNUR Kural",
        "SANTRAL OZNUR Orijinal Kod",
        "SANTRAL OZNUR Tsoft Kodu",
        "SANTRAL OZNUR Kural",
        "YILDIZ OZNUR Orijinal Kod",
        "YILDIZ OZNUR Tsoft Kodu",
        "YILDIZ OZNUR Kural",
    ]

    if tum_b2b_list:
        b2b_all = pd.concat(tum_b2b_list, ignore_index=True)
    else:
        b2b_all = bos_df(b2b_kolonlar)

    print()
    print(f"Toplam B2B satırı: {len(b2b_all)}")
    print(f"T-Soft ürün satırı: {len(tsoft_df)}")

    tedarikci_bazli_satirlar = []
    b2bde_var_sitede_yok_satirlar = []
    santral_eslesme_bekleyen_satirlar = []
    bos_kod_satirlar = []
    hatali_stok_satirlar = []
    coklu_satirlar = []

    for _, b2b_row in b2b_all.iterrows():
        b2b_kod = str(b2b_row["B2B Web Servis Kodu"]).strip()
        eslesme_tipi = str(b2b_row["EslesmeTipi"]).strip().upper()
        eslesmede_kullanilacak_kod = str(b2b_row["Eslesmede Kullanilacak Kod"]).strip()
        yeni_stok = b2b_row["Yeni Stok"]
        santral_schrack_mi = (
            "SANTRAL" in metin_norm(b2b_row.get("Tedarikci", ""))
            and metin_norm(b2b_row.get("B2B Marka", "")) == "SCHRACK"
        )

        if not b2b_kod or b2b_kod.lower() == "nan":
            satir = b2b_row.to_dict()
            satir["Neden"] = "B2B ürün kodu boş"
            bos_kod_satirlar.append(satir)
            continue

        if pd.isna(yeni_stok):
            satir = b2b_row.to_dict()
            satir["Neden"] = "Stok değeri okunamadı"
            hatali_stok_satirlar.append(satir)
            continue

        if eslesme_tipi == "MANUEL_ESLESME" and not eslesmede_kullanilacak_kod:
            satir = b2b_row.to_dict()
            satir["Neden"] = "Santral manuel eşleşme tablosunda karşılık bulunamadı"
            satir["Eklenecek SantralKodu"] = b2b_row["B2B Web Servis Kodu"]
            satir["Eklenecek TsoftWebServisKodu"] = ""
            santral_eslesme_bekleyen_satirlar.append(satir)
            continue

        santral_manuel_tsoft_row = santral_product_id_eslesme_bul(
            b2b_row,
            tsoft_df,
        )
        oznur_direkt_sonuc = oznurkablo_direkt_eslesme_bul(b2b_row, tsoft_df)

        if santral_manuel_tsoft_row is not None:
            tsoft_row = santral_manuel_tsoft_row
            eslesme_bicimi = "SANTRAL_MANUEL_PRODUCT_ID"
            kullanilan_kod = str(tsoft_row["ProductId"]).strip()
        elif oznur_direkt_sonuc is not None:
            tsoft_row, eslesme_bicimi, kullanilan_kod = oznur_direkt_sonuc
        elif metin_norm(b2b_row.get("Tedarikci", "")) == "YILDIZ":
            tsoft_row, eslesme_bicimi, kullanilan_kod = yildiz_direkt_eslesme_bul(
                b2b_kod,
                b2b_row.get("B2B Marka", ""),
                tsoft_df,
            )
        else:
            tsoft_row, eslesme_bicimi, kullanilan_kod = b2b_eslesme_bul(
                eslesmede_kullanilacak_kod,
                tsoft_df,
                lookup
            )

        if tsoft_row is None:
            satir = b2b_row.to_dict()

            if eslesme_tipi == "MANUEL_ESLESME":
                satir["Neden"] = "Manuel T-Soft kodu bulundu ama T-Soft ürün listesinde eşleşmedi"
            else:
                satir["Neden"] = "B2B'de var, sitede yok"

            if santral_schrack_mi:
                satir["SANTRAL Kod Düzeltme Durumu"] = "SCHRACK karşılığı bulunamadı"
                satir["SANTRAL Kod Düzeltme Notu"] = "SCHRACK için kural uygulanmadı; kod T-Soft Web Servis Kodu ile direkt karşılaştırıldı ve eşleşmedi."

            b2bde_var_sitede_yok_satirlar.append(satir)
            continue

        if isinstance(tsoft_row, str) and tsoft_row == "COKLU":
            satir = b2b_row.to_dict()
            satir["Neden"] = "Birden fazla T-Soft ürünüyle eşleşti"
            satir["Eslesme Bicimi"] = eslesme_bicimi
            satir["Kullanılan Kod"] = kullanilan_kod
            coklu_satirlar.append(satir)
            continue

        guncel_bayi_stok = int(yeni_stok)

        santral_kod_duzeltme_durumu = b2b_row.get("SANTRAL Kod Düzeltme Durumu", "")
        santral_kod_duzeltme_notu = b2b_row.get("SANTRAL Kod Düzeltme Notu", "")

        if santral_schrack_mi:
            santral_kod_duzeltme_durumu = "SCHRACK kod eşleşti"
            santral_kod_duzeltme_notu = "SCHRACK için kural uygulanmadı; kod T-Soft Web Servis Kodu ile direkt karşılaştırıldı ve eşleşti."

        tedarikci_bazli_satirlar.append({
            "Tedarikci": b2b_row["Tedarikci"],
            "EslesmeTipi": eslesme_tipi,
            "B2B Satır No": b2b_row.get("B2B Satır No", ""),
            "B2B Kaynak Dosya": b2b_row.get("B2B Kaynak Dosya", ""),
            "ProductId": tsoft_row["ProductId"],
            "Web Servis Kodu": tsoft_row["Web Servis Kodu"],
            "Ürün Adı": tsoft_row.get("Ürün Adı", ""),
            "Model": tsoft_row["Model"],
            "Ana Kategori Adı": tsoft_row.get("Ana Kategori Adı", ""),
            "Yeni Stok": guncel_bayi_stok,
            "Tsoft Mevcut Stok": tsoft_row["Stok"],
            "Stok Birimi": tsoft_row["Stok Birimi"],
            "Fiyat Birimi": tsoft_row["Fiyat Birimi"],
            "Satış Fiyatı": tsoft_row["Satış Fiyatı"],
            "KDV Dahil Fiyat": tsoft_row.get("KDV Dahil Fiyat", ""),
            "B2B Web Servis Kodu": b2b_row["B2B Web Servis Kodu"],
            "B2B Temiz Kod": b2b_row["B2B Temiz Kod"],
            "B2B Sade Kod": b2b_row["B2B Sade Kod"],
            "B2B Stok Metni": b2b_row["B2B Stok Metni"],
            "B2B Marka": b2b_row["B2B Marka"],
            "EM Kodu": b2b_row["EM Kodu"],
            "KIRAC Orijinal Ürün Kodu": b2b_row.get("KIRAC Orijinal Ürün Kodu", ""),
            "KIRAC Kod Düzeltme Durumu": b2b_row.get("KIRAC Kod Düzeltme Durumu", ""),
            "KIRAC Kod Düzeltme Notu": b2b_row.get("KIRAC Kod Düzeltme Notu", ""),
            "SANTRAL Orijinal Ürün Kodu": b2b_row.get("SANTRAL Orijinal Ürün Kodu", ""),
            "SANTRAL Kod Düzeltme Durumu": santral_kod_duzeltme_durumu,
            "SANTRAL Kod Düzeltme Notu": santral_kod_duzeltme_notu,
            "Manuel Tsoft Web Servis Kodu": b2b_row["Manuel Tsoft Web Servis Kodu"],
            "Eslesmede Kullanilacak Kod": eslesmede_kullanilacak_kod,
            "KIRAC Legrand Orijinal Kod": b2b_row.get("KIRAC Legrand Orijinal Kod", ""),
            "KIRAC Legrand Sıfırsız Kod": b2b_row.get("KIRAC Legrand Sıfırsız Kod", ""),
            "KIRAC Legrand Sıfır Kuralı": b2b_row.get("KIRAC Legrand Sıfır Kuralı", ""),
            "KIRAC Bosch Orijinal Kod": b2b_row.get("KIRAC Bosch Orijinal Kod", ""),
            "KIRAC Bosch Boşluklu Kod": b2b_row.get("KIRAC Bosch Boşluklu Kod", ""),
            "KIRAC Bosch Boşluk Kuralı": b2b_row.get("KIRAC Bosch Boşluk Kuralı", ""),
            "KIRAC Ortac Orijinal Kod": b2b_row.get("KIRAC Ortac Orijinal Kod", ""),
            "KIRAC Ortac Boşluklu Kod": b2b_row.get("KIRAC Ortac Boşluklu Kod", ""),
            "KIRAC Ortac Boşluk Kuralı": b2b_row.get("KIRAC Ortac Boşluk Kuralı", ""),
            "KIRAC Şafak Orijinal Kod": b2b_row.get("KIRAC Şafak Orijinal Kod", ""),
            "KIRAC Şafak Boşluklu Kod": b2b_row.get("KIRAC Şafak Boşluklu Kod", ""),
            "KIRAC Şafak KM Kuralı": b2b_row.get("KIRAC Şafak KM Kuralı", ""),
            "KIRAC OZNUR Orijinal Kod": b2b_row.get("KIRAC OZNUR Orijinal Kod", ""),
            "KIRAC OZNUR Tsoft Kodu": b2b_row.get("KIRAC OZNUR Tsoft Kodu", ""),
            "KIRAC OZNUR Kural": b2b_row.get("KIRAC OZNUR Kural", ""),
            "SANTRAL OZNUR Orijinal Kod": b2b_row.get("SANTRAL OZNUR Orijinal Kod", ""),
            "SANTRAL OZNUR Tsoft Kodu": b2b_row.get("SANTRAL OZNUR Tsoft Kodu", ""),
            "SANTRAL OZNUR Kural": b2b_row.get("SANTRAL OZNUR Kural", ""),
            "YILDIZ OZNUR Orijinal Kod": b2b_row.get("YILDIZ OZNUR Orijinal Kod", ""),
            "YILDIZ OZNUR Tsoft Kodu": b2b_row.get("YILDIZ OZNUR Tsoft Kodu", ""),
            "YILDIZ OZNUR Kural": b2b_row.get("YILDIZ OZNUR Kural", ""),
            "Eslesme Bicimi": eslesme_bicimi,
            "Kullanılan Kod": kullanilan_kod
        })

    tedarikci_bazli_df = pd.DataFrame(tedarikci_bazli_satirlar)
    b2bde_var_sitede_yok_df = pd.DataFrame(b2bde_var_sitede_yok_satirlar)
    santral_eslesme_bekleyen_df = pd.DataFrame(santral_eslesme_bekleyen_satirlar)
    bos_kod_df = pd.DataFrame(bos_kod_satirlar)
    hatali_stok_df = pd.DataFrame(hatali_stok_satirlar)
    coklu_df = pd.DataFrame(coklu_satirlar)

    hazir_df = hazir_stok_olustur(tedarikci_bazli_df)

    toplam_entes_b2b_satiri = 0
    toplam_entes_eslesen_satir = 0

    if not b2b_all.empty and "B2B Marka" in b2b_all.columns:
        toplam_entes_b2b_satiri = int(
            b2b_all["B2B Marka"].apply(lambda x: metin_norm(x) == "ENTES").sum()
        )

    if not tedarikci_bazli_df.empty and "B2B Marka" in tedarikci_bazli_df.columns:
        toplam_entes_eslesen_satir = int(
            tedarikci_bazli_df["B2B Marka"].apply(lambda x: metin_norm(x) == "ENTES").sum()
        )

    ozet_satirlari = [
        ["Toplam B2B satırı", len(b2b_all)],
        ["Toplam ENTES B2B satırı", toplam_entes_b2b_satiri],
        ["T-Soft ürün satırı", len(tsoft_df)],
        ["Tedarikçi bazlı eşleşen satır", len(tedarikci_bazli_df)],
        ["Tedarikçi bazlı ENTES eşleşen satır", toplam_entes_eslesen_satir],
        ["Güncellenecek benzersiz ürün", len(hazir_df)],
        ["B2B'de var sitede yok", len(b2bde_var_sitede_yok_df)],
        ["Santral manuel eşleşme bekleyen", len(santral_eslesme_bekleyen_df)],
        ["Boş web servis kodu", len(bos_kod_df)],
        ["Hatalı stok satırı", len(hatali_stok_df)],
        ["Çoklu eşleşme", len(coklu_df)],
    ]

    if not b2b_all.empty and "Tedarikci" in b2b_all.columns:
        b2b_ozet_df = b2b_all.copy()
        b2b_ozet_df["Tedarikci"] = b2b_ozet_df["Tedarikci"].astype(str).str.strip()

        if "B2B Marka" in b2b_ozet_df.columns:
            b2b_ozet_df["B2B Marka Norm"] = b2b_ozet_df["B2B Marka"].apply(metin_norm)
        else:
            b2b_ozet_df["B2B Marka Norm"] = ""

        tedarikci_toplam_ozet = (
            b2b_ozet_df
            .groupby("Tedarikci", dropna=False)
            .size()
            .reset_index(name="Toplam B2B Satırı")
        )

        tedarikci_entes_ozet = (
            b2b_ozet_df[b2b_ozet_df["B2B Marka Norm"] == "ENTES"]
            .groupby("Tedarikci", dropna=False)
            .size()
            .reset_index(name="ENTES B2B Satırı")
        )

        tedarikci_ozet = tedarikci_toplam_ozet.merge(
            tedarikci_entes_ozet,
            on="Tedarikci",
            how="left"
        )
    else:
        tedarikci_ozet = pd.DataFrame(columns=[
            "Tedarikci",
            "Toplam B2B Satırı",
            "ENTES B2B Satırı",
        ])

    if not tedarikci_bazli_df.empty and "Tedarikci" in tedarikci_bazli_df.columns:
        eslesen_ozet_df = tedarikci_bazli_df.copy()
        eslesen_ozet_df["Tedarikci"] = eslesen_ozet_df["Tedarikci"].astype(str).str.strip()

        if "B2B Marka" in eslesen_ozet_df.columns:
            eslesen_ozet_df["B2B Marka Norm"] = eslesen_ozet_df["B2B Marka"].apply(metin_norm)
        else:
            eslesen_ozet_df["B2B Marka Norm"] = ""

        eslesen_ozet = (
            eslesen_ozet_df
            .groupby("Tedarikci", dropna=False)
            .size()
            .reset_index(name="Eşleşen Satır")
        )

        entes_eslesen_ozet = (
            eslesen_ozet_df[eslesen_ozet_df["B2B Marka Norm"] == "ENTES"]
            .groupby("Tedarikci", dropna=False)
            .size()
            .reset_index(name="ENTES Eşleşen Satır")
        )

        eslesen_ozet = eslesen_ozet.merge(
            entes_eslesen_ozet,
            on="Tedarikci",
            how="left"
        )
    else:
        eslesen_ozet = pd.DataFrame({
            "Tedarikci": pd.Series(dtype="str"),
            "Eşleşen Satır": pd.Series(dtype="int"),
            "ENTES Eşleşen Satır": pd.Series(dtype="int"),
        })

    tedarikci_ozet["Tedarikci"] = tedarikci_ozet["Tedarikci"].astype(str).str.strip()
    eslesen_ozet["Tedarikci"] = eslesen_ozet["Tedarikci"].astype(str).str.strip()

    tedarikci_ozet = tedarikci_ozet.merge(
        eslesen_ozet,
        on="Tedarikci",
        how="left"
    )

    for sayi_kolonu in [
        "Toplam B2B Satırı",
        "ENTES B2B Satırı",
        "Eşleşen Satır",
        "ENTES Eşleşen Satır",
    ]:
        if sayi_kolonu not in tedarikci_ozet.columns:
            tedarikci_ozet[sayi_kolonu] = 0

        tedarikci_ozet[sayi_kolonu] = pd.to_numeric(
            tedarikci_ozet[sayi_kolonu],
            errors="coerce"
        ).fillna(0).astype(int)

    for _, row in tedarikci_ozet.iterrows():
        ozet_satirlari.append([
            f"{row['Tedarikci']} toplam satır",
            int(row["Toplam B2B Satırı"])
        ])
        ozet_satirlari.append([
            f"{row['Tedarikci']} ENTES satırı",
            int(row["ENTES B2B Satırı"])
        ])
        ozet_satirlari.append([
            f"{row['Tedarikci']} eşleşen satır",
            int(row["Eşleşen Satır"])
        ])
        ozet_satirlari.append([
            f"{row['Tedarikci']} ENTES eşleşen satır",
            int(row["ENTES Eşleşen Satır"])
        ])

    ozet_df = pd.DataFrame(ozet_satirlari, columns=["Kalem", "Adet"])

    konya_raporlar = konya_raporlari_olustur(
        b2b_all=b2b_all,
        tedarikci_bazli_df=tedarikci_bazli_df,
        b2bde_var_sitede_yok_df=b2bde_var_sitede_yok_df,
        bos_kod_df=bos_kod_df,
        hatali_stok_df=hatali_stok_df,
        coklu_df=coklu_df
    )

    kirac_raporlar = kirac_raporlari_olustur(
        b2b_all=b2b_all,
        tedarikci_bazli_df=tedarikci_bazli_df,
        b2bde_var_sitede_yok_df=b2bde_var_sitede_yok_df,
        bos_kod_df=bos_kod_df,
        hatali_stok_df=hatali_stok_df,
        coklu_df=coklu_df
    )

    santral_raporlar = santral_raporlari_olustur(
        b2b_all=b2b_all,
        tedarikci_bazli_df=tedarikci_bazli_df,
        b2bde_var_sitede_yok_df=b2bde_var_sitede_yok_df,
        bos_kod_df=bos_kod_df,
        hatali_stok_df=hatali_stok_df,
        coklu_df=coklu_df
    )

    raporlari_kaydet(
        tsoft_df=tsoft_df,
        hazir_df=hazir_df,
        tedarikci_bazli_df=tedarikci_bazli_df,
        b2bde_var_sitede_yok_df=b2bde_var_sitede_yok_df,
        santral_eslesme_bekleyen_df=santral_eslesme_bekleyen_df,
        bos_kod_df=bos_kod_df,
        hatali_stok_df=hatali_stok_df,
        coklu_df=coklu_df,
        ozet_df=ozet_df,
        konya_raporlar=konya_raporlar,
        kirac_raporlar=kirac_raporlar,
        santral_raporlar=santral_raporlar
    )

    print()
    print("=== B2B STOK EŞLEŞTİRME TAMAMLANDI ===")


if __name__ == "__main__":
    main()

