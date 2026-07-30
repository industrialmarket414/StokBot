from pathlib import Path
import json
import requests
import pandas as pd

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
            "password": ""
        }

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(ornek, f, ensure_ascii=False, indent=4)

        print("Ayar dosyası oluşturuldu:")
        print(CONFIG_FILE)
        print("Bu dosyaya API şifresini yazıp tekrar çalıştır.")
        raise SystemExit

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    site_url = data.get("site_url", "").strip().rstrip("/")
    username = data.get("username", "").strip()
    password = data.get("password", "").strip()

    if not site_url or not username or not password:
        print("tsoft_api.json içinde site_url, username veya password boş.")
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

    print("Token alınıyor...")
    print(url)

    response = requests.post(
        url,
        data={"pass": password},
        headers={"Accept": "application/json"},
        timeout=30
    )

    try:
        data = response.json()
    except Exception:
        print("Token JSON cevabı alınamadı.")
        print(response.text[:1000])
        raise SystemExit

    log_path = LOG_DIR / "tsoft_login_response.json"

    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    token = token_cevaptan_bul(data)

    if not token:
        print("Token alınamadı.")
        print("API cevabı:")
        print(json.dumps(data, ensure_ascii=False, indent=4)[:3000])
        raise SystemExit

    print("Token alındı.")
    return token


def urun_cevabindan_liste_al(data):
    if isinstance(data, list):
        return data

    if not isinstance(data, dict):
        return []

    olasi_anahtarlar = [
        "data",
        "products",
        "items",
        "list",
        "result"
    ]

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
        "ProductName",
        "Model",
        "DefaultCategoryName",
        "Stock",
        "StockUnit",
        "Currency",
        "CurrencyId",
        "SellingPrice",
        "DiscountedSellingPrice"
    ])

    while True:
        print()
        print(f"Ürünler çekiliyor... start={start}, limit={limit}")

        url = f"{site_url}/rest1/product/getProducts"

        response = requests.post(
            url,
            data={
                "token": token,
                "start": str(start),
                "limit": str(limit),
                "columns": kolonlar,
                "CustomerGroupCode": "TZ0",
                "FetchDiscountedPriceDetail": "true",
            },
            headers={"Accept": "application/json"},
            timeout=60
        )

        try:
            data = response.json()
        except Exception:
            print("Ürün JSON cevabı alınamadı.")
            print(response.text[:1000])
            raise SystemExit

        log_path = LOG_DIR / f"urun_cevap_start_{start}.json"

        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        urunler = urun_cevabindan_liste_al(data)

        if not urunler:
            print("Yeni ürün gelmedi. Çekme işlemi durduruldu.")
            break

        sayfa_imzasi = tuple(
            kayit_kimligi(urun, index)
            for index, urun in enumerate(urunler)
        )

        if sayfa_imzasi in gorulen_sayfalar:
            print("Aynı sayfa tekrar geldi. Sonsuz döngü olmaması için işlem durduruldu.")
            break

        gorulen_sayfalar.add(sayfa_imzasi)

        tum_urunler.extend(urunler)

        print(f"Bu sayfada gelen ürün: {len(urunler)}")
        print(f"Toplam ürün: {len(tum_urunler)}")

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
        return (
            df[kolon_adi]
            .astype(str)
            .str.replace(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", regex=True)
            .str.strip()
        )

    return pd.Series([""] * len(df))


def secili_alanlari_hazirla(df):
    secili_df = pd.DataFrame()

    secili_df["ProductId"] = kolon_al(df, "ProductId")
    secili_df["Web Servis Kodu"] = kolon_al(df, "ProductCode")
    secili_df["Ürün Adı"] = kolon_al(df, "ProductName")
    secili_df["Model"] = kolon_al(df, "Model")
    secili_df["Ana Kategori Adı"] = kolon_al(df, "DefaultCategoryName")
    secili_df["Stok"] = kolon_al(df, "Stock")
    secili_df["Stok Birimi"] = kolon_al(df, "StockUnit")

    if "Currency" in df.columns:
        secili_df["Fiyat Birimi"] = kolon_al(df, "Currency")
    else:
        secili_df["Fiyat Birimi"] = kolon_al(df, "CurrencyId")

    secili_df["T-SOFT LİSTE FİYATI"] = kolon_al(df, "SellingPrice")
    secili_df["T-SOFT İNDİRİMLİ FİYATI"] = kolon_al(
        df, "DiscountedSellingPrice"
    )

    secili_df = secili_df[
        secili_df["Web Servis Kodu"].astype(str).str.strip() != ""
    ].copy()

    secili_df = secili_df[
        secili_df["Web Servis Kodu"].astype(str).str.lower() != "nan"
    ].copy()

    return secili_df


def urunleri_excele_kaydet(urunler):
    secili_dosya = TSOFT_DIR / "tsoft_secili_alanlar.xlsx"

    if not urunler:
        print("Kaydedilecek ürün bulunamadı.")
        return

    duz_urunler = [duzlestir_kayit(u) for u in urunler]
    df = pd.DataFrame(duz_urunler)

    secili_df = secili_alanlari_hazirla(df)

    with pd.ExcelWriter(secili_dosya, engine="openpyxl") as writer:
        secili_df.to_excel(writer, sheet_name="SeciliAlanlar", index=False)

    print()
    print("T-Soft ürün listesi kaydedildi:")
    print(secili_dosya)

    print()
    print(f"Toplam ürün satırı: {len(secili_df)}")

    print()
    print("Seçili kolonlar:")
    for col in secili_df.columns:
        print("-", col)


def main():
    print("=== T-SOFT API ÜRÜN LİSTESİ ÇEKME BAŞLADI ===")

    site_url, username, password = config_oku()

    print("Site URL:", site_url)
    print("Kullanıcı:", username)

    token = api_login(site_url, username, password)

    urunler = urunleri_cek(site_url, token)

    print()
    print(f"API'den çekilen toplam ürün: {len(urunler)}")

    urunleri_excele_kaydet(urunler)

    print()
    print("=== İŞLEM TAMAMLANDI ===")
    print("Tsoft klasörünü kontrol et:")
    print(TSOFT_DIR)


if __name__ == "__main__":
    main()
