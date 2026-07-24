from pathlib import Path
import json
import pandas as pd

BASE_DIR = Path(__file__).parent

KLASORLER = [
    "Ayarlar",
    "B2B/KONYA",
    "B2B/KIRAC",
    "B2B/SANTRAL",
    "Tsoft",
    "Cikti",
    "Log",
    "Archive/B2B",
    "Archive/Tsoft",
    "Archive/Cikti",
    "Archive/Log",
]


def klasorleri_olustur():
    for klasor in KLASORLER:
        path = BASE_DIR / klasor
        path.mkdir(parents=True, exist_ok=True)
        print("Klasör hazır:", path)


def tsoft_api_olustur():
    dosya = BASE_DIR / "Ayarlar" / "tsoft_api.json"

    if dosya.exists():
        print("Zaten var, dokunulmadı:", dosya)
        return

    data = {
        "site_url": "https://www.endustriyelmarket.com",
        "username": "stokbot",
        "password": ""
    }

    with open(dosya, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print("Oluşturuldu:", dosya)
    print("Bu dosyada password alanına web servis şifresini yazmalısın.")


def tedarikci_ayar_olustur():
    dosya = BASE_DIR / "Ayarlar" / "tedarikci_ayar.xlsx"

    if dosya.exists():
        print("Zaten var, dokunulmadı:", dosya)
        return

    df = pd.DataFrame([
        {
            "Tedarikci": "KONYA",
            "Aktif": 1,
            "MarkaFiltresi": "ENTES",
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "Ürün Kodu",
            "StokKolonu": "Stok",
            "MarkaKolonu": "Marka"
        },
        {
            "Tedarikci": "KIRAC",
            "Aktif": 1,
            "MarkaFiltresi": "ENTES",
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "",
            "StokKolonu": "",
            "MarkaKolonu": ""
        },
        {
            "Tedarikci": "SANTRAL",
            "Aktif": 1,
            "MarkaFiltresi": "ENTES",
            "DosyaUzantisi": ".xlsx",
            "KodKolonu": "",
            "StokKolonu": "",
            "MarkaKolonu": ""
        },
    ])

    with pd.ExcelWriter(dosya, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="TedarikciAyar", index=False)

    print("Oluşturuldu:", dosya)


def kod_kurallari_olustur():
    dosya = BASE_DIR / "Ayarlar" / "kod_kurallari.xlsx"

    if dosya.exists():
        print("Zaten var, dokunulmadı:", dosya)
       