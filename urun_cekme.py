from pathlib import Path
from datetime import datetime
import json

import pandas as pd
import requests


BASE_DIR = Path(__file__).parent

AYARLAR_DIR = BASE_DIR / "Ayarlar"
TSOFT_DIR = BASE_DIR / "Tsoft"
LOG_DIR = BASE_DIR / "Log"

AYARLAR_DIR.mkdir(exist_ok=True)
TSOFT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

CONFIG_FILE = AYARLAR_DIR / "tsoft_api.json"


def config_oku():
    if not CONFIG_FILE.exists():
        ornek = {
            "site_url": "https://www.endustriyelmarket.com",
            "username": "stokbot",
            "password": "",
        }

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(ornek, f, ensure_ascii=False, indent=4)

        print("Ayar dosyasi olusturuldu:")
        print(CONFIG_FILE)
        print("Bu dosyaya API sifresini yazip tekrar calistir.")
        raise SystemExit

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    site_url = str(data.get("site_url", "")).strip().rstrip("/")
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not site_url or not username or not password:
        print("tsoft_api.json icinde site_url, username veya password bos.")
        print("Dosya:")
        print(CONFIG_FILE)
        raise SystemExit

    return site_url, username, password


def token_cevaptan_bul(data):
    if not isinstance(data, dict):
        return None

    for key in ["token", "access_token", "bearer_token"]:
        if data.get(key):
            return data.get(key)

    if isinstance(data.get("data"), dict):
        for key in ["token", "access_token", "bearer_token"]:
            if data["data"].get(key):
                return data["data"].get(key)

    if isinstance(data.get("data"), list) and data["data"]:
        first = data["data"][0]

        if isinstance(first, dict):
            for key in ["token", "access_token", "bearer_token"]:
                if first.get(key):
                    return first.get(key)

    return None


def api_login(site_url, username, password):
    url = f"{site_url}/rest1/auth/login/{username}"

    print("Token aliniyor...")
    print(url)

    response = requests.post(
        url,
        data={"pass": password},
        headers={"Accept": "application/json"},
        timeout=30,
    )

    try:
        data = response.json()
    except Exception:
        print("Token JSON cevabi alinamadi.")
        print(response.text[:1000])
        raise SystemExit

    log_path = LOG_DIR / "tsoft_login_response.json"

    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    token = token_cevaptan_bul(data)

    if not token:
        print("Token alinamadi.")
        print("API cevabi:")
        print(json.dumps(data, ensure_ascii=False, indent=4)[:3000])
        raise SystemExit

    print("Token alindi.")
    return token


def urun_cevabindan_liste_al(data):
    if isinstance(data, list):
        return data

    if not isinstance(data, dict):
        return []

    olasi_anahtarlar = ["data", "products", "items", "list", "result"]

    for key in olasi_anahtarlar:
        value = data.get(key)

        if isinstance(value, list):
            return value

        if isinstance(value, dict):
            for inner_key in olasi_anahtarlar:
                inner_value = value.get(inner_key)

                if isinstance(inner_value, list):
                    return inner_value

    return []


def kayit_kimligi(kayit, index):
    if not isinstance(kayit, dict):
        return str(kayit)

    product_id = str(kayit.get("ProductId", "")).strip()
    if product_id:
        return f"ProductId:{product_id}"

    product_code = str(kayit.get("ProductCode", "")).strip()
    if product_code:
        return f"ProductCode:{product_code}"

    return f"Index:{index}:{json.dumps(kayit, ensure_ascii=False, sort_keys=True)}"


def urunleri_cek(site_url, token, limit=500):
    tum_urunler = []
    start = 0
    gorulen_sayfalar = set()

    kolonlar = ",".join([
        "ProductId",
        "ProductCode",
        "Marka",
        "Brand",
        "BrandId",
        "BrandName",
        "BrandCode",
        "BrandTitle",
        "ProductBrand",
        "ProductBrandId",
        "ProductBrandName",
        "Manufacturer",
        "ManufacturerName",
        "Model",
        "Stock",
        "StockUnit",
        "Currency",
        "CurrencyId",
        "SellingPrice",
    ])

    while True:
        print()
        print(f"Urunler cekiliyor... start={start}, limit={limit}")

        url = f"{site_url}/rest1/product/getProducts"

        response = requests.post(
            url,
            data={
                "token": token,
                "start": str(start),
                "limit": str(limit),
                "columns": kolonlar,
            },
            headers={"Accept": "application/json"},
            timeout=60,
        )

        try:
            data = response.json()
        except Exception:
            print("Urun JSON cevabi alinamadi.")
            print(response.text[:1000])
            raise SystemExit

        log_path = LOG_DIR / f"urun_cevap_start_{start}.json"

        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        urunler = urun_cevabindan_liste_al(data)

        if not urunler:
            print("Yeni urun gelmedi. Cekme islemi durduruldu.")
            break

        sayfa_imzasi = tuple(
            kayit_kimligi(urun, index)
            for index, urun in enumerate(urunler)
        )

        if sayfa_imzasi in gorulen_sayfalar:
            print("Ayni sayfa tekrar geldi. Sonsuz dongu olmamasi icin islem durduruldu.")
            break

        gorulen_sayfalar.add(sayfa_imzasi)
        tum_urunler.extend(urunler)

        print(f"Bu sayfada gelen urun: {len(urunler)}")
        print(f"Toplam urun: {len(tum_urunler)}")

        start += len(urunler)

    return tum_urunler


def duzlestir_kayit(kayit):
    yeni = {}

    if not isinstance(kayit, dict):
        return {"value": kayit}

    for key, value in kayit.items():
        if isinstance(value, (dict, list)):
            yeni[key] = json.dumps(value, ensure_ascii=False)
        else:
            yeni[key] = value

    return yeni


def kolon_al(df, kolon_adi):
    if kolon_adi in df.columns:
        return df[kolon_adi].astype(str).str.strip()

    return pd.Series([""] * len(df))


def ilk_dolu_kolon(df, kolon_adlari):
    sonuc = pd.Series([""] * len(df))

    for kolon in kolon_adlari:
        if kolon in df.columns:
            degerler = kolon_al(df, kolon)
            sonuc = sonuc.mask(sonuc == "", degerler)

    return sonuc


def marka_al(df):
    marka_kolonlari = [
        "Marka",
        "BrandName",
        "BrandTitle",
        "ProductBrandName",
        "ManufacturerName",
        "Brand",
        "ProductBrand",
        "Manufacturer",
        "BrandCode",
        "BrandId",
        "ProductBrandId",
    ]

    bulunanlar = [kolon for kolon in marka_kolonlari if kolon in df.columns]

    for kolon in df.columns:
        kolon_lower = str(kolon).lower()
        if (
            "marka" in kolon_lower
            or "brand" in kolon_lower
            or "manufacturer" in kolon_lower
        ) and kolon not in bulunanlar:
            bulunanlar.append(kolon)

    return ilk_dolu_kolon(df, bulunanlar)


def secili_alanlari_hazirla(df):
    secili_df = pd.DataFrame()

    secili_df["ProductId"] = kolon_al(df, "ProductId")
    secili_df["Web Servis Kodu"] = kolon_al(df, "ProductCode")
    secili_df["Marka"] = marka_al(df)
    secili_df["Model"] = kolon_al(df, "Model")
    secili_df["Stok"] = kolon_al(df, "Stock")
    secili_df["Stok Birimi"] = kolon_al(df, "StockUnit")

    if "Currency" in df.columns:
        secili_df["Fiyat Birimi"] = kolon_al(df, "Currency")
    else:
        secili_df["Fiyat Birimi"] = kolon_al(df, "CurrencyId")

    secili_df["Satis Fiyati"] = kolon_al(df, "SellingPrice")

    secili_df = secili_df[
        secili_df["Web Servis Kodu"].astype(str).str.strip() != ""
    ].copy()

    secili_df = secili_df[
        secili_df["Web Servis Kodu"].astype(str).str.lower() != "nan"
    ].copy()

    return secili_df


def urunleri_excele_kaydet(urunler):
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")

    ana_dosya = TSOFT_DIR / "urun_cekme_api_urun_listesi.xlsx"
    tarihli_dosya = TSOFT_DIR / f"urun_cekme_api_urun_listesi_{tarih}.xlsx"

    secili_dosya = TSOFT_DIR / "urun_cekme_secili_alanlar.xlsx"
    secili_csv = TSOFT_DIR / "urun_cekme_secili_alanlar.csv"

    if not urunler:
        print("Kaydedilecek urun bulunamadi.")
        return

    duz_urunler = [duzlestir_kayit(u) for u in urunler]
    df = pd.DataFrame(duz_urunler)

    if "Marka" not in df.columns:
        df["Marka"] = marka_al(df)

    with pd.ExcelWriter(ana_dosya, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Urunler", index=False)

    with pd.ExcelWriter(tarihli_dosya, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Urunler", index=False)

    secili_df = secili_alanlari_hazirla(df)

    with pd.ExcelWriter(secili_dosya, engine="openpyxl") as writer:
        secili_df.to_excel(writer, sheet_name="SeciliAlanlar", index=False)

    secili_df.to_csv(secili_csv, sep=";", index=False, encoding="utf-8-sig")

    print()
    print("Tum API urun listesi kaydedildi:")
    print(ana_dosya)
    print(tarihli_dosya)

    print()
    print("Secili alanlar kaydedildi:")
    print(secili_dosya)
    print(secili_csv)

    print()
    print(f"Toplam urun satiri: {len(df)}")
    print(f"Secili alan satiri: {len(secili_df)}")

    print()
    print("Secili kolonlar:")
    for col in secili_df.columns:
        print("-", col)


def main():
    print("=== T-SOFT API URUN LISTESI CEKME BASLADI ===")

    site_url, username, password = config_oku()

    print("Site URL:", site_url)
    print("Kullanici:", username)

    token = api_login(site_url, username, password)
    urunler = urunleri_cek(site_url, token)

    print()
    print(f"API'den cekilen toplam urun: {len(urunler)}")

    urunleri_excele_kaydet(urunler)

    print()
    print("=== ISLEM TAMAMLANDI ===")
    print("Tsoft klasorunu kontrol et:")
    print(TSOFT_DIR)


if __name__ == "__main__":
    main()
