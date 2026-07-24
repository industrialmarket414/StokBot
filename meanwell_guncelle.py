"""MEANWELL eslestirme ciktisindaki stoklari T-Soft'a gonderir."""

from pathlib import Path
from datetime import datetime
import json
import math
import re

import pandas as pd
import requests


BASE_DIR = Path(__file__).resolve().parent
AYARLAR_DIR = BASE_DIR / "Ayarlar"
CIKTI_DIR = BASE_DIR / "Cikti"
MEANWELL_CIKTI_DIR = CIKTI_DIR / "MEANWELL"
LOG_DIR = BASE_DIR / "Log"
CONFIG_FILE = AYARLAR_DIR / "tsoft_api.json"
STOK_DOSYA = MEANWELL_CIKTI_DIR / "meanwell_stok_guncellenecekler.xlsx"
BATCH_SIZE = 100
GUNCELLEME_HARIC_MARKALAR = {
    "WAGO", "ORTAÇ", "ORTAC", "OPAŞ", "OPAS", "ÖZNUR", "ÖZNUR KABLO",
    "OZNUR", "OZNUR KABLO", "MUCCO", "EMAS",
}


def guncelleme_haric_markalari_cikar(df):
    if df is None or df.empty or "B2B Marka" not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()
    maske = df["B2B Marka"].astype(str).str.strip().str.upper().isin(GUNCELLEME_HARIC_MARKALAR)
    if maske.any():
        print(f"Guvenlik filtresi: guncelleme disi marka satiri: {int(maske.sum())}")
    return df[~maske].copy()


def urun_kodu_temizle(deger):
    if pd.isna(deger):
        return ""
    return re.sub(r"\s+", "", str(deger))


def stok_sayisina_cevir(deger, varsayilan=None):
    if pd.isna(deger):
        return varsayilan

    metin = str(deger).strip().upper()
    if not metin or metin == "NAN":
        return varsayilan
    if "YOK" in metin:
        return 0

    metin = re.sub(r"\s*(ADET|ADT|TANE|PCS|PC)\.?\s*$", "", metin)
    metin = re.sub(r"\s+", "", metin)
    if "." in metin and "," in metin:
        if metin.rfind(",") > metin.rfind("."):
            metin = metin.replace(".", "").replace(",", ".")
        else:
            metin = metin.replace(",", "")
    elif re.fullmatch(r"-?\d{1,3}(?:[.,]\d{3})+", metin):
        metin = metin.replace(".", "").replace(",", "")
    else:
        metin = metin.replace(",", ".")

    try:
        sayi = float(metin)
    except ValueError:
        return varsayilan
    if not sayi.is_integer():
        return varsayilan
    return int(sayi)


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
    log_path = LOG_DIR / "tsoft_meanwell_tum_stok_login_response.json"
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
    response = requests.post(
        f"{site_url}/rest1/product/getProducts",
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
    return urunler[0] if urunler else None


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

    metin = json.dumps(data, ensure_ascii=False).lower()
    hata_kelimeleri = [
        "hata", "error", "basarisiz", "gecersiz", "yetkisiz",
        "unauthorized", "permission",
    ]
    basari_kelimeleri = ["basari", "success", "pri001", "gni001"]

    if any(kelime in metin for kelime in hata_kelimeleri):
        return False
    if any(kelime in metin for kelime in basari_kelimeleri):
        return True
    return False


def stok_listesini_oku():
    if not STOK_DOSYA.exists():
        print("MEANWELL stok dosyasi bulunamadi:")
        print(STOK_DOSYA)
        print("Once stok_eslestir.py calistirilmalidir.")
        raise SystemExit(1)

    df = pd.read_excel(STOK_DOSYA, dtype=str).fillna("")
    for kolon in ("ProductId", "Web Servis Kodu"):
        if kolon not in df.columns:
            print(f"Excel icinde eksik kolon var: {kolon}")
            raise SystemExit(1)

    if "B2B Stok Metni" not in df.columns:
        print("Excel icinde eksik kolon var: B2B Stok Metni")
        raise SystemExit(1)
    return guncelleme_haric_markalari_cikar(df)


def guncelleme_listesi_hazirla(df):
    satirlar = []
    hatalar = []

    for index, row in df.iterrows():
        product_id = str(row.get("ProductId", "")).strip()
        product_code = urun_kodu_temizle(row.get("Web Servis Kodu", ""))
        stok_ham = row.get("B2B Stok Metni", "")
        stok = stok_sayisina_cevir(stok_ham)

        if not product_id:
            hatalar.append(f"Excel satir {index + 2}: ProductId bos.")
        elif not product_code:
            hatalar.append(f"Excel satir {index + 2}: Web Servis Kodu bos.")
        elif stok is None:
            hatalar.append(f"Excel satir {index + 2}: Stok sayi degil ({stok_ham!r}).")
        elif stok < 0:
            hatalar.append(f"Excel satir {index + 2}: Stok negatif ({stok}).")
        else:
            satirlar.append({
                "ProductId": product_id,
                "ProductCode": product_code,
                "Web Servis Kodu": product_code,
                "Guncellenecek Stok": stok,
                "B2B Marka": str(row.get("B2B Marka", "MEANWELL")).strip() or "MEANWELL",
                "Model": str(row.get("Model", "")).strip(),
            })

    if hatalar:
        LOG_DIR.mkdir(exist_ok=True)
        hata_dosya = LOG_DIR / f"meanwell_hazirlik_hatalari_{datetime.now():%Y%m%d_%H%M%S}.txt"
        hata_dosya.write_text("\n".join(hatalar), encoding="utf-8")
        print("Hazirlik hatalari var; API'ye veri gonderilmedi.")
        print(hata_dosya)
        raise SystemExit(1)
    if not satirlar:
        print("Gonderilecek MEANWELL stok satiri yok.")
        raise SystemExit(1)

    return pd.DataFrame(satirlar)


def on_kontrol_yaz(df):
    CIKTI_DIR.mkdir(parents=True, exist_ok=True)
    dosya = CIKTI_DIR / f"meanwell_tum_stok_gonderim_on_kontrol_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    df.to_excel(dosya, index=False)
    return dosya


def batchlere_bol(df, batch_size):
    for baslangic in range(0, len(df), batch_size):
        yield (
            baslangic // batch_size + 1,
            df.iloc[baslangic:baslangic + batch_size].copy(),
        )


def stok_batch_gonder(site_url, token, batch_no, batch_df):
    url = f"{site_url}/rest1/product/updateProducts"
    gonderilecekler = [
        {
            "ProductCode": urun_kodu_temizle(row["ProductCode"]),
            "Stock": int(row["Guncellenecek Stok"]),
        }
        for _, row in batch_df.iterrows()
    ]

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
    log_path = LOG_DIR / f"tsoft_meanwell_tum_stok_batch_{batch_no:03d}_{tarih}.json"
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


def stoklari_gonder(site_url, token, df):
    sonuclar = []
    for batch_no, batch_df in batchlere_bol(df, BATCH_SIZE):
        print(f"Batch {batch_no} gonderiliyor ({len(batch_df)} urun)...")
        sonuc = stok_batch_gonder(site_url, token, batch_no, batch_df)
        sonuclar.append(sonuc)
        print(f"Batch {batch_no}: HTTP {sonuc['HTTP']} / {sonuc['API Basari']}")

    sonuc_df = pd.DataFrame(sonuclar)
    sonuc_dosya = CIKTI_DIR / f"meanwell_tum_stok_gonderim_sonuc_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    sonuc_df.to_excel(sonuc_dosya, index=False)
    return sonuc_df, sonuc_dosya


def ornek_kontrol(site_url, token, df):
    ornekler = pd.concat([df.head(5), df.tail(5)], ignore_index=True).drop_duplicates("ProductCode")
    satirlar = []
    for _, row in ornekler.iterrows():
        urun = urun_getir_product_code(site_url, token, row["ProductCode"])
        api_stok = urun_stok_degeri(urun)
        api_sayi = stok_sayisina_cevir(api_stok)
        beklenen = int(row["Guncellenecek Stok"])
        satirlar.append({
            "ProductCode": row["ProductCode"],
            "Beklenen Stok": beklenen,
            "API Okunan Stok": api_stok,
            "Kontrol": "TAMAM" if api_sayi == beklenen else "KONTROL ET",
        })
    kontrol = pd.DataFrame(satirlar)
    dosya = CIKTI_DIR / f"meanwell_tum_stok_ornek_kontrol_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    kontrol.to_excel(dosya, index=False)
    print(kontrol.to_string(index=False))
    print("Ornek kontrol dosyasi:", dosya)


def main():
    print("=== MEANWELL T-SOFT STOK GUNCELLEME ===")
    print("Bu islem meanwell_stok_guncellenecekler.xlsx icindeki tum MEANWELL urunlerini gonderir.")
    print()
    site_url, username, password = config_oku()
    guncelleme_df = guncelleme_listesi_hazirla(stok_listesini_oku())
    on_kontrol = on_kontrol_yaz(guncelleme_df)

    print("Site URL:", site_url)
    print("Gonderilecek urun:", len(guncelleme_df))
    print("Batch sayisi:", math.ceil(len(guncelleme_df) / BATCH_SIZE))
    print("On kontrol dosyasi:", on_kontrol)
    print(guncelleme_df[["ProductCode", "Guncellenecek Stok"]].head(10).to_string(index=False))
    print()
    print("DIKKAT: Bu islem T-Soft canli stoklarini gunceller.")
    if input("Gondermek icin MEANWELL yaz: ").strip().upper() != "MEANWELL":
        print("Islem iptal edildi; API'ye veri gonderilmedi.")
        return

    token = api_login(site_url, username, password)
    sonuc_df, sonuc_dosya = stoklari_gonder(site_url, token, guncelleme_df)
    print("Batch sonuc dosyasi:", sonuc_dosya)
    sorunlu = sonuc_df[sonuc_df["API Basari"].astype(str) != "EVET"]
    if not sorunlu.empty:
        print("Uyari: Kontrol edilmesi gereken batch var.")
        print(sorunlu.to_string(index=False))

    ornek_kontrol(site_url, token, guncelleme_df)
    print("MEANWELL stok guncelleme tamamlandi.")
    print("Son adim: T-Soft panelde ornek urunleri manuel kontrol et.")


if __name__ == "__main__":
    main()
