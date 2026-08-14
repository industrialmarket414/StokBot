from pathlib import Path
from datetime import datetime
import re

import pandas as pd

from tsoft_api import (
    api_login,
    config_oku,
    sayiya_cevir,
    stok_batch_gonder,
    urun_getir_product_code,
    urun_kodu_temizle,
    urun_stok_degeri,
)


BASE_DIR = Path(__file__).parent
CIKTI_DIR = BASE_DIR / "Cikti"
TEST_DIR = CIKTI_DIR / "TEST"

GUNCELLEME_DOSYA = CIKTI_DIR / "tum_bayiler_stok_guncellenecekler.xlsx"


def test_urununu_oku(product_code):
    if not GUNCELLEME_DOSYA.exists():
        print("Tum bayiler guncelleme dosyasi bulunamadi:")
        print(GUNCELLEME_DOSYA)
        print("Once stok_eslestir.py kodunu calistir.")
        raise SystemExit

    df = pd.read_excel(GUNCELLEME_DOSYA, dtype=str).fillna("")

    gerekli_kolonlar = [
        "ProductId",
        "Web Servis Kodu",
        "Bayi Toplam Stok",
        "Guncellenecek Stok",
    ]
    for kolon in gerekli_kolonlar:
        if kolon not in df.columns:
            print(f"Excel icinde eksik kolon var: {kolon}")
            raise SystemExit

    temiz_product_code = urun_kodu_temizle(product_code)
    df["_Temiz Kod"] = df["Web Servis Kodu"].apply(urun_kodu_temizle)
    bulunan = df[
        df["_Temiz Kod"].str.upper() == temiz_product_code.upper()
    ].copy()

    if bulunan.empty:
        print(f"{product_code} guncelleme listesinde bulunamadi.")
        raise SystemExit

    if len(bulunan) != 1:
        print(f"{product_code} icin birden fazla satir bulundu. Islem durduruldu.")
        print(bulunan.to_string(index=False))
        raise SystemExit

    row = bulunan.iloc[0]
    temiz_product_code = urun_kodu_temizle(row["Web Servis Kodu"])
    bayi_toplam_stok = sayiya_cevir(row["Bayi Toplam Stok"])
    guncellenecek_stok = sayiya_cevir(row["Guncellenecek Stok"])

    if str(row.get("Guncelleme Durumu", "")).strip().upper() == "GONDERILMEYECEK":
        print("Bu urun stok guncellemesine kapali. Islem durduruldu.")
        raise SystemExit

    if guncellenecek_stok < 0:
        print("Guncellenecek stok negatif olamaz.")
        raise SystemExit

    return {
        "ProductId": str(row["ProductId"]).strip(),
        "ProductCode": temiz_product_code,
        "Bayi Toplam Stok": bayi_toplam_stok,
        "Guncellenecek Stok": guncellenecek_stok,
        "Tedarikciler": str(row.get("Tedarikciler", "")).strip(),
        "B2B Marka": str(row.get("B2B Marka", "")).strip(),
        "DIA Stok": sayiya_cevir(row.get("DIA Stok", 0)),
        "Sabit Stok": sayiya_cevir(row.get("Sabit Stok", 0)),
    }


def api_stok_oku(site_url, token, product_code):
    urun = urun_getir_product_code(site_url, token, product_code)
    if not urun:
        print(f"T-Soft API'de urun bulunamadi: {product_code}")
        raise SystemExit

    stok = urun_stok_degeri(urun)
    if stok is None:
        print("API cevabinda stok alani bulunamadi.")
        raise SystemExit

    return sayiya_cevir(stok), urun


def tek_stok_gonder(site_url, token, product_id, product_code, stok):
    df = pd.DataFrame([{
        "ProductId": product_id,
        "ProductCode": product_code,
        "Guncellenecek Stok": int(stok),
    }])
    return stok_batch_gonder(site_url, token, 1, df)


def main():
    print("=== TEK URUN STOK GUNCELLEME ===")
    product_code = input("Guncellenecek urun kodunu yaz: ").strip()
    if not product_code:
        print("Urun kodu bos birakildi. Islem iptal edildi.")
        return

    print()

    TEST_DIR.mkdir(parents=True, exist_ok=True)
    test_urunu = test_urununu_oku(product_code)

    site_url, username, password = config_oku()
    token = api_login(site_url, username, password)
    eski_stok, _ = api_stok_oku(site_url, token, test_urunu["ProductCode"])

    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    dosya_kodu = re.sub(r"[^A-Za-z0-9._-]+", "_", test_urunu["ProductCode"])
    yedek_dosya = TEST_DIR / f"{dosya_kodu}_test_yedek_{tarih}.xlsx"
    pd.DataFrame([{
        "ProductId": test_urunu["ProductId"],
        "ProductCode": test_urunu["ProductCode"],
        "Eski Tsoft Stok": eski_stok,
        "Testte Gonderilecek Stok": test_urunu["Guncellenecek Stok"],
        "Bayi Toplam Stok": test_urunu["Bayi Toplam Stok"],
        "Tedarikciler": test_urunu["Tedarikciler"],
        "B2B Marka": test_urunu["B2B Marka"],
        "DIA Stok": test_urunu["DIA Stok"],
        "Sabit Stok": test_urunu["Sabit Stok"],
    }]).to_excel(yedek_dosya, index=False)

    print("ProductId:", test_urunu["ProductId"])
    print("ProductCode:", test_urunu["ProductCode"])
    print("Tedarikciler:", test_urunu["Tedarikciler"])
    print("B2B Marka:", test_urunu["B2B Marka"])
    print("T-Soft mevcut stok:", eski_stok)
    print("Tum bayiler toplam stogu:", test_urunu["Bayi Toplam Stok"])
    print("DIA stogu:", test_urunu["DIA Stok"])
    print("Sabit stok:", test_urunu["Sabit Stok"])
    print("API'ye gonderilecek stok:", test_urunu["Guncellenecek Stok"])
    print("Geri alma yedegi:", yedek_dosya)
    print()

    beklenen_onay = f"{test_urunu['ProductCode']} GUNCELLE".upper()
    onay = input(f"Devam etmek icin {beklenen_onay} yaz: ").strip().upper()
    if onay != beklenen_onay:
        print("Islem iptal edildi. API'ye stok gonderilmedi.")
        return

    sonuc = tek_stok_gonder(
        site_url,
        token,
        test_urunu["ProductId"],
        test_urunu["ProductCode"],
        test_urunu["Guncellenecek Stok"],
    )
    print(f"Gonderim sonucu: HTTP {sonuc['HTTP']} / {sonuc['API Basari']}")

    api_stok, _ = api_stok_oku(site_url, token, test_urunu["ProductCode"])
    beklenen = test_urunu["Guncellenecek Stok"]
    print("API'den okunan stok:", api_stok)
    print("Kontrol:", "TAMAM" if api_stok == beklenen else "KONTROL ET")

    if api_stok != beklenen:
        print("Beklenen stok API'de gorulmedi. Log dosyasini kontrol et.")
        return

    print()
    geri_onay = input(
        "Yeni stok kalici kalsin istiyorsan ENTER; eski stoga donmek icin GERI AL yaz: "
    ).strip().upper()
    if geri_onay != "GERI AL":
        print("Geri alma yapilmadi. Test stogu T-Soft'ta kaldi.")
        return

    geri_sonuc = tek_stok_gonder(
        site_url,
        token,
        test_urunu["ProductId"],
        test_urunu["ProductCode"],
        eski_stok,
    )
    print(f"Geri alma sonucu: HTTP {geri_sonuc['HTTP']} / {geri_sonuc['API Basari']}")

    geri_api_stok, _ = api_stok_oku(site_url, token, test_urunu["ProductCode"])
    print("Geri alma sonrasi API stogu:", geri_api_stok)
    print("Geri alma kontrol:", "TAMAM" if geri_api_stok == eski_stok else "KONTROL ET")


if __name__ == "__main__":
    main()
