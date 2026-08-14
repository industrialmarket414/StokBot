from pathlib import Path
from datetime import datetime
import json
import re

import pandas as pd
import requests


BASE_DIR = Path(__file__).parent
AYARLAR_DIR = BASE_DIR / "Ayarlar"
LOG_DIR = BASE_DIR / "Log"

CONFIG_FILE = AYARLAR_DIR / "tsoft_api.json"


def urun_kodu_temizle(deger):
    if pd.isna(deger):
        return ""

    return re.sub(r"\s+", "", str(deger))


def sayiya_cevir(deger, varsayilan=0):
    if pd.isna(deger):
        return varsayilan

    metin = str(deger).strip().replace(",", ".")

    if metin == "" or metin.lower() == "nan":
        return varsayilan

    try:
        return int(float(metin))
    except Exception:
        return varsayilan


def config_oku():
    if not CONFIG_FILE.exists():
        print("Ayar dosyasi bulunamadi:")
        print(CONFIG_FILE)
        raise SystemExit

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    site_url = str(data.get("site_url", "")).strip().rstrip("/")
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()

    if not site_url or not username or not password:
        print("tsoft_api.json icinde site_url, username veya password bos.")
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

    LOG_DIR.mkdir(exist_ok=True)
    log_path = LOG_DIR / "tsoft_api_login_response.json"

    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    token = token_cevaptan_bul(data)

    if not token:
        print("Token alinamadi. API cevabi loglandi:")
        print(log_path)
        raise SystemExit

    print("Token alindi.")
    return token


def cevap_liste_al(data):
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


def urun_getir_product_code(site_url, token, product_code):
    url = f"{site_url}/rest1/product/getProducts"

    response = requests.post(
        url,
        data={
            "token": token,
            "start": "0",
            "limit": "1",
            "columns": "ProductId,ProductCode,Stock",
            "f": f"ProductCode|{product_code}|equal|",
        },
        headers={"Accept": "application/json"},
        timeout=60,
    )

    try:
        data = response.json()
    except Exception:
        return None

    urunler = cevap_liste_al(data)

    if not urunler:
        return None

    return urunler[0]


def urun_getir_product_id(site_url, token, product_id):
    """T-Soft urununu ProductId ile getirir.

    ProductCode filtresi Turkce karakter veya uzun sayisal kodlarda sonuc
    dondurmezse guvenli ikinci arama olarak kullanilir.
    """
    url = f"{site_url}/rest1/product/getProducts"

    response = requests.post(
        url,
        data={
            "token": token,
            "start": "0",
            "limit": "1",
            "columns": "ProductId,ProductCode,Stock",
            "f": f"ProductId|{product_id}|equal|",
        },
        headers={"Accept": "application/json"},
        timeout=60,
    )

    try:
        data = response.json()
    except Exception:
        return None

    urunler = cevap_liste_al(data)
    if not urunler:
        return None

    return urunler[0]


def urun_stok_degeri(urun):
    if not isinstance(urun, dict):
        return None

    for alan in ["Stock", "stock", "Stok"]:
        if alan in urun:
            return urun.get(alan)

    return None


def api_cevabi_basarili_mi(data):
    if data is None:
        return False

    if isinstance(data, dict):
        success = data.get("success")
        mesajlar = data.get("message")

        if success is False:
            return False

        if success is True:
            if not isinstance(mesajlar, list) or not mesajlar:
                return True

            basarili_kodlar = {"PRI001", "GNI001"}
            return all(
                isinstance(mesaj, dict)
                and str(mesaj.get("code", "")).upper() in basarili_kodlar
                and not mesaj.get("errorField")
                for mesaj in mesajlar
            )

    metin = json.dumps(data, ensure_ascii=False).lower()

    hata_kelimeleri = [
        "hata",
        "basarisiz",
        "gecersiz",
        "yetkisiz",
        "unauthorized",
        "permission",
    ]
    basari_kelimeleri = ["basari", "success", "pri001", "gni001"]

    if any(kelime in metin for kelime in hata_kelimeleri):
        return False

    if any(kelime in metin for kelime in basari_kelimeleri):
        return True

    return False


def stok_batch_gonder(site_url, token, batch_no, batch_df):
    url = f"{site_url}/rest1/product/updateProducts"

    gonderilecekler = []
    for _, row in batch_df.iterrows():
        api_product_code = row.get("API ProductCode", "")
        if pd.isna(api_product_code) or not str(api_product_code).strip():
            product_code = urun_kodu_temizle(row["ProductCode"])
        else:
            # API on kontrolunde ProductId ile bulunan kaydin T-Soft'taki gercek
            # kodunu (bosluklar dahil) aynen kullan.
            product_code = str(api_product_code).strip()

        gonderilecekler.append(
            {
                "ProductCode": product_code,
                "Stock": int(row["Guncellenecek Stok"]),
            }
        )

    response = requests.post(
        url,
        data={
            "token": token,
            "data": json.dumps(gonderilecekler, ensure_ascii=False),
        },
        headers={"Accept": "application/json"},
        timeout=90,
    )

    try:
        cevap = response.json()
    except Exception:
        cevap = {"raw_response": response.text}

    LOG_DIR.mkdir(exist_ok=True)
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"tsoft_api_stok_batch_{batch_no:03d}_{tarih}.json"

    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(cevap, f, ensure_ascii=False, indent=4)

    http_basarili = response.status_code in [200, 201, 202]
    cevap_basarili = api_cevabi_basarili_mi(cevap)

    return {
        "Batch": batch_no,
        "Urun Sayisi": len(batch_df),
        "HTTP": response.status_code,
        "API Basari": "EVET" if http_basarili and cevap_basarili else "KONTROL",
        "Log": str(log_path),
    }
