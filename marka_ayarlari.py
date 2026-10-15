from pathlib import Path
from datetime import datetime
import shutil
import pandas as pd


BASE_DIR = Path(__file__).parent

AYARLAR_DIR = BASE_DIR / "Ayarlar"
AYAR_FILE = AYARLAR_DIR / "tedarikci_ayar.xlsx"

AYARLAR_DIR.mkdir(exist_ok=True)


KONYA_MARKALAR = [
    "ACK",
    "BEMİS",
    "BYLION",
    "CATA",
    "ÇETİNKAYA",
    "EMAS",
    "ENTES",
    "HES",
    "KARDEŞ",
    "KLEMSAN",
    "KRAUS NAIMER",
    "LEGRAND",
    "METE",
    "MUTLUSAN",
    "NADE",
    "NEAD",
    "OMRON",
    "OPAŞ",
    "ORTAÇ",
    "OSRAM",
    "PHILIPS",
    "PHOENIX CONTACT",
    "SCHNEIDER",
    "SIEMENS",
    "ŞAFAK",
    "TELEMECANIQUE",
    "VIKO-PANASONIC",
    "WAGO",
    "WEIDMULLER",
    "WOER",
]


KIRAC_MARKALAR = [
    "3M",
    "Bemis",
    "Birtaş",
    "Emas",
    "Bosch",
    "Bticino",
    "Byes",
    "Entes",
    "HellermannTyton",
    "Iboco",
    "Isiso",
    "Kardeş Elektrik",
    "Kraus & Naimer",
    "Legrand",
    "Mavigard",
    "Mesan",
    "Osram",
    "Ortaç",
    "Öznur",
    "Pelsan",
    "Philips",
    "Phoenix Contact",
    "Prysmian",
    "Schneider Electric",
    "Siemens",
    "SIEMENS",
    "Şafak Elektrik",
    "Telemecanique Sensors",
    "TP Electric",
    "Trimbox",
    "Weidmüller",
]



SANTRAL_MARKALAR = [
    "ENTES",
    "KLEMSAN",
    "ÖZNUR KABLO",
    "SCHRACK",
    "SIEMENS",
]


MET_MARKALAR = [
    "NOAS",
]


KLTSHOP_MARKALAR = [
    "Schneider",
]


MEANWELL_MARKALAR = [
    "MEANWELL",
]


YILDIZ_MARKALAR = [
    "ACK",
    "ARTLİNE",
    "ATILIM",
    "AUDIO",
    "BTİCİNO",
    "BYLION",
    "CATA",
    "ÇETİNKAYA",
    "ERAT",
    "ERSE",
    "GWEST",
    "HCS",
    "HES",
    "IŞILDAR",
    "JUPİTER",
    "LEGRAND",
    "MUTLUSAN",
    "NADE",
    "NEAD",
    "NEXANS",
    "ÖZNUR",
    "PRYSMİAN",
    "TAMER",
    "THEA",
    "ÜNAL",
    "VİKO",
    "WAGO",
]


ANKA_MARKALAR = [
    "ACK",
    "BEMİS",
    "CATA",
    "ÇETİNKAYA",
    "EMAS",
    "ENTES",
    "FINDER",
    "GWEST",
    "HİMEL",
    "IŞILDAR",
    "Kraus & Naimer",
    "LEGRAND",
    "MEANWELL",
    "METE ENERJİ",
    "MUCCO",
    "MUTLUSAN",
    "NADE",
    "NEAD",
    "OMRON",
    "OPAŞ",
    "ORTAÇ",
    "Phoenix Contact",
    "ŞAFAK",
    "SCHRACK",
    "SIEMENS",
    "TP ELEKTRİK",
    "ÜNAL",
    "VİKO",
    "WEIDMULLER",
]


VEKMAR_MARKALAR = [
    "SCHNEIDER",
    "TELEMECANIQUE SENSORS",
    "HİMEL",
    "KLEMSAN",
    "METE ENERJİ",
    "SIEMENS",
    "ŞAFAK",
]


def eski_ayari_yedekle():
    if not AYAR_FILE.exists():
        print("Eski tedarikci_ayar.xlsx bulunamadı, yeni dosya oluşturulacak.")
        return None

    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    yedek_dosya = AYARLAR_DIR / f"tedarikci_ayar_yedek_{tarih}.xlsx"

    shutil.copy2(AYAR_FILE, yedek_dosya)

    print("Eski tedarikci_ayar.xlsx yedeklendi:")
    print(yedek_dosya)

    return yedek_dosya


def konya_satirlari_olustur():
    satirlar = []

    for marka in KONYA_MARKALAR:
        satirlar.append({
            "Tedarikci": "KONYA",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def kirac_satirlari_olustur():
    satirlar = []

    for marka in KIRAC_MARKALAR:
        satirlar.append({
            "Tedarikci": "KIRAC",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def santral_satirlari_olustur():
    satirlar = []

    for marka in SANTRAL_MARKALAR:
        satirlar.append({
            "Tedarikci": "SANTRAL",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def met_satirlari_olustur():
    satirlar = []

    for marka in MET_MARKALAR:
        satirlar.append({
            "Tedarikci": "MET",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Stok Kodu",
            "StokKolonu": "Kalan Stok",
            "MarkaKolonu": "__MARKA_FILTRESI__",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def kltshop_satirlari_olustur():
    satirlar = []

    for marka in KLTSHOP_MARKALAR:
        satirlar.append({
            "Tedarikci": "KLTSHOP",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok Adeti",
            "MarkaKolonu": "__MARKA_FILTRESI__",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def meanwell_satirlari_olustur():
    satirlar = []

    for marka in MEANWELL_MARKALAR:
        satirlar.append({
            "Tedarikci": "MEANWELL",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Model No",
            "StokKolonu": "Stok",
            "MarkaKolonu": "__MARKA_FILTRESI__",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def yildiz_satirlari_olustur():
    satirlar = []

    for marka in YILDIZ_MARKALAR:
        satirlar.append({
            "Tedarikci": "YILDIZ",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def anka_satirlari_olustur():
    satirlar = []

    for marka in ANKA_MARKALAR:
        satirlar.append({
            "Tedarikci": "ANKA",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def vekmar_satirlari_olustur():
    satirlar = []

    for marka in VEKMAR_MARKALAR:
        satirlar.append({
            "Tedarikci": "VEKMAR",
            "Aktif": 1,
            "MarkaFiltresi": marka,
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka",
            "EslesmeTipi": "WEBSERVIS",
        })

    return satirlar


def ayar_dosyasi_yaz(df):
    try:
        with pd.ExcelWriter(AYAR_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="TedarikciAyar", index=False)

    except PermissionError:
        print()
        print("HATA: tedarikci_ayar.xlsx şu an açık görünüyor.")
        print("Excel dosyasını kapatıp tekrar çalıştır.")
        raise SystemExit


def main():
    print("=== TEDARIKCI AYAR MARKA GUNCELLEME ===")
    print()

    eski_ayari_yedekle()

    satirlar = []
    satirlar.extend(konya_satirlari_olustur())
    satirlar.extend(kirac_satirlari_olustur())
    satirlar.extend(santral_satirlari_olustur())
    satirlar.extend(met_satirlari_olustur())
    satirlar.extend(kltshop_satirlari_olustur())
    satirlar.extend(meanwell_satirlari_olustur())
    satirlar.extend(yildiz_satirlari_olustur())
    satirlar.extend(vekmar_satirlari_olustur())
    satirlar.extend(anka_satirlari_olustur())

    df = pd.DataFrame(satirlar)

    # Güvenlik: Boşlukları temizle.
    df["Tedarikci"] = df["Tedarikci"].astype(str).str.strip()
    df["MarkaFiltresi"] = df["MarkaFiltresi"].astype(str).str.strip()
    df["DosyaUzantisi"] = df["DosyaUzantisi"].astype(str).str.strip()
    df["KodKolonu"] = df["KodKolonu"].astype(str).str.strip()
    df["StokKolonu"] = df["StokKolonu"].astype(str).str.strip()
    df["MarkaKolonu"] = df["MarkaKolonu"].astype(str).str.strip()
    df["EslesmeTipi"] = df["EslesmeTipi"].astype(str).str.strip().str.upper()

    # Güvenlik: Aynı tedarikçi + marka tekrar yazıldıysa tekilleştir.
    df = df.drop_duplicates(
        subset=["Tedarikci", "MarkaFiltresi"],
        keep="first"
    )

    ayar_dosyasi_yaz(df)

    print()
    print("Yeni tedarikci_ayar.xlsx oluşturuldu/güncellendi:")
    print(AYAR_FILE)

    print()
    print("Özet:")
    print("KONYA marka sayısı:", len(KONYA_MARKALAR))
    print("KIRAC marka sayısı:", len(KIRAC_MARKALAR))
    print("SANTRAL marka sayısı:", len(SANTRAL_MARKALAR))
    print("MET marka sayısı:", len(MET_MARKALAR))
    print("KLTSHOP marka sayısı:", len(KLTSHOP_MARKALAR))
    print("MEANWELL marka sayısı:", len(MEANWELL_MARKALAR))
    print("YILDIZ marka sayısı:", len(YILDIZ_MARKALAR))
    print("VEKMAR marka sayısı:", len(VEKMAR_MARKALAR))
    print("ANKA marka sayısı:", len(ANKA_MARKALAR))
    print("Toplam ayar satırı:", len(df))

    print()
    print("Tedarikçi bazlı satır sayısı:")
    print(df["Tedarikci"].value_counts().to_string())

    print()
    print("İlk 10 satır:")
    print(df.head(10).to_string(index=False))

    print()
    print("Son 10 satır:")
    print(df.tail(10).to_string(index=False))

    print()
    print("=== ISLEM TAMAMLANDI ===")


if __name__ == "__main__":
    main()
