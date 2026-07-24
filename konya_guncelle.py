from pathlib import Path
from datetime import datetime
import json
import math
import re

import pandas as pd
import requests


BASE_DIR = Path(__file__).parent

AYARLAR_DIR = BASE_DIR / "Ayarlar"
CIKTI_DIR = BASE_DIR / "Cikti"
LOG_DIR = BASE_DIR / "Log"
KONYA_CIKTI_DIR = CIKTI_DIR / "KONYA"

CONFIG_FILE = AYARLAR_DIR / "tsoft_api.json"
KONYA_STOK_DOSYA = KONYA_CIKTI_DIR / "konya_stok_guncellenecekler.xlsx"

BATCH_SIZE = 100
KONYA_GUNCELLEME_HARIC_MARKALAR = {
    "HES",
    "VIKO",
    "VIKO-PANASONIC",
    "VIKO PANASONIC",
    "ARTLINE",
    "THEA",
    "ACK",
    "WAGO",
    "ORTAC",
    "OPAS",
    "OZNUR",
    "OZNUR KABLO",
    "MUCCO",
    "EMAS",
}


def urun_kodu_temizle(deger):
    if pd.isna(deger):
        return ""

    return re.sub(r"\s+", " ", str(deger)).strip()


def konya_guncelleme_hariclerini_cikar(df):
    if df is None or df.empty or "B2B Marka" not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()

    marka_norm = (
        df["B2B Marka"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.translate(str.maketrans("ÇĞİÖŞÜ", "CGIOSU"))
    )
    haric_maskesi = marka_norm.isin(KONYA_GUNCELLEME_HARIC_MARKALAR)
    haric_sayi = int(haric_maskesi.sum())

    if haric_sayi:
        print(
            "Guvenlik filtresi: KONYA guncelleme disi marka satiri: "
            f"{haric_sayi}"
        )

    return df[~haric_maskesi].copy()


def sayiya_cevir(deger, varsayilan=0):
    if pd.isna(deger):
        return varsayilan

    metin = re.sub(r"\s+", " ", str(deger).strip()).upper()

    if metin == "" or metin == "NAN":
        return varsayilan

    if "YOK" in metin:
        return 0

    metin = re.sub(
        r"\s*(ADET|ADT|AD|TANE|SET|PCS|PC|METRE|MTR|MT|M)\.?\s*$",
        "",
        metin,
    )
    metin = re.sub(r"\s+", "", metin)

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
        timeout=30
    )

    try:
        data = response.json()
    except Exception:
        print("Token JSON cevabi alinamadi.")
        print(response.text[:1000])
        raise SystemExit

    LOG_DIR.mkdir(exist_ok=True)
    log_path = LOG_DIR / "tsoft_konya_tum_stok_login_response.json"

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
            "f": f"ProductCode|{product_code}|equal|"
        },
        headers={"Accept": "application/json"},
        timeout=60
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

    metin = json.dumps(data, ensure_ascii=False).lower()

    hata_kelimeleri = [
        "hata",
        "error",
        "basarisiz",
        "gecersiz",
        "yetkisiz",
        "unauthorized",
        "permission"
    ]

    basari_kelimeleri = [
        "basari",
        "success",
        "pri001",
        "gni001"
    ]

    if any(kelime in metin for kelime in hata_kelimeleri):
        return False

    if any(kelime in metin for kelime in basari_kelimeleri):
        return True

    return False


def konya_stok_listesini_oku():
    if not KONYA_STOK_DOSYA.exists():
        print("KONYA stok dosyasi bulunamadi:")
        print(KONYA_STOK_DOSYA)
        raise SystemExit

    df = pd.read_excel(KONYA_STOK_DOSYA, dtype=str).fillna("")

    gerekli_kolonlar = [
        "ProductId",
        "Web Servis Kodu",
    ]

    for kolon in gerekli_kolonlar:
        if kolon not in df.columns:
            print(f"Excel icinde eksik kolon var: {kolon}")
            raise SystemExit

    if "Yeni Stok" not in df.columns and "B2B Stok Metni" not in df.columns:
        print("Excel icinde Yeni Stok veya B2B Stok Metni kolonu bulunamadi.")
        raise SystemExit

    return konya_guncelleme_hariclerini_cikar(df)


def guncelleme_listesi_hazirla(df):
    hazir_satirlar = []
    hatalar = []

    for index, row in df.iterrows():
        product_id = str(row["ProductId"]).strip()
        product_code = urun_kodu_temizle(row["Web Servis Kodu"])

        stok_degeri = row.get("Yeni Stok", "")
        if str(stok_degeri).strip() == "":
            stok_degeri = row.get("B2B Stok Metni", "")
        guncellenecek_stok = sayiya_cevir(stok_degeri)

        if not product_id:
            hatalar.append(f"Excel satir {index + 2}: ProductId bos.")
            continue

        if not product_code:
            hatalar.append(f"ProductId {product_id}: Web Servis Kodu bos.")
            continue

        if guncellenecek_stok < 0:
            hatalar.append(
                f"ProductId {product_id}: Guncellenecek Stok negatif ({guncellenecek_stok})."
            )
            continue

        hazir_satirlar.append({
            "ProductId": product_id,
            "ProductCode": product_code,
            "Web Servis Kodu": product_code,
            "Guncellenecek Stok": guncellenecek_stok,
            "B2B Marka": str(row.get("B2B Marka", "")).strip(),
            "Model": str(row.get("Model", "")).strip(),
        })

    if hatalar:
        tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
        hata_dosya = LOG_DIR / f"tsoft_konya_tum_stok_hazirlik_hatalari_{tarih}.txt"
        LOG_DIR.mkdir(exist_ok=True)

        with open(hata_dosya, "w", encoding="utf-8") as f:
            for hata in hatalar:
                f.write(hata + "\n")

        print("Hazirlik hatalari var. API'ye gonderim yapilmadi.")
        print("Hata dosyasi:")
        print(hata_dosya)
        raise SystemExit

    if not hazir_satirlar:
        print("Gonderilecek KONYA stok satiri yok.")
        raise SystemExit

    hazir_df = pd.DataFrame(hazir_satirlar)
    tekrarlar = hazir_df[hazir_df["ProductCode"].duplicated(keep=False)].copy()

    if not tekrarlar.empty:
        tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
        tekrar_dosya = KONYA_CIKTI_DIR / f"konya_tekrar_eden_urun_kodlari_son_satir_kullanildi_{tarih}.xlsx"
        tekrarlar.to_excel(tekrar_dosya, index=False)

        print()
        print("Ayni ProductCode birden fazla satirda bulundu.")
        print("Toplama yapilmadi; her ProductCode icin Excel'deki son satir gonderilecek.")
        print("Tekrar eden kod raporu:")
        print(tekrar_dosya)

        hazir_df = hazir_df.drop_duplicates(subset=["ProductCode"], keep="last").copy()

    return hazir_df


def on_kontrol_raporu_yaz(guncelleme_df):
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    rapor_dosya = KONYA_CIKTI_DIR / f"konya_tum_stok_gonderim_on_kontrol_{tarih}.xlsx"
    guncelleme_df.to_excel(rapor_dosya, index=False)
    return rapor_dosya


def ozet_yaz(guncelleme_df):
    print()
    print("KONYA tum marka stok guncelleme ozeti:")
    print(f"Gonderilecek urun sayisi: {len(guncelleme_df)}")
    print(f"Batch boyutu: {BATCH_SIZE}")
    print(f"Batch sayisi: {math.ceil(len(guncelleme_df) / BATCH_SIZE)}")

    if "B2B Marka" in guncelleme_df.columns:
        marka_ozet = (
            guncelleme_df
            .groupby("B2B Marka", dropna=False)
            .size()
            .reset_index(name="Urun Sayisi")
            .sort_values("Urun Sayisi", ascending=False)
        )

        print()
        print("Marka ozeti:")
        print(marka_ozet.to_string(index=False))

    print()
    print("Ilk 10 urun:")
    print(
        guncelleme_df[
            [
                "ProductId",
                "ProductCode",
                "B2B Marka",
                "Guncellenecek Stok",
            ]
        ].head(10).to_string(index=False)
    )


def batchlere_bol(df, batch_size):
    for baslangic in range(0, len(df), batch_size):
        yield baslangic // batch_size + 1, df.iloc[baslangic:baslangic + batch_size].copy()


def stok_batch_gonder(site_url, token, batch_no, batch_df):
    url = f"{site_url}/rest1/product/updateProducts"

    gonderilecekler = [
        {
            "ProductCode": urun_kodu_temizle(row["ProductCode"]),
            "Stock": int(row["Guncellenecek Stok"])
        }
        for _, row in batch_df.iterrows()
    ]

    response = requests.post(
        url,
        data={
            "token": token,
            "data": json.dumps(gonderilecekler, ensure_ascii=False)
        },
        headers={"Accept": "application/json"},
        timeout=90
    )

    try:
        cevap = response.json()
    except Exception:
        cevap = {"raw_response": response.text}

    LOG_DIR.mkdir(exist_ok=True)
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"tsoft_konya_tum_stok_batch_{batch_no:03d}_{tarih}.json"

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


def stoklari_gonder(site_url, token, guncelleme_df):
    sonuc_satirlari = []

    for batch_no, batch_df in batchlere_bol(guncelleme_df, BATCH_SIZE):
        print(f"Batch {batch_no} gonderiliyor... ({len(batch_df)} urun)")

        sonuc = stok_batch_gonder(
            site_url=site_url,
            token=token,
            batch_no=batch_no,
            batch_df=batch_df
        )

        sonuc_satirlari.append(sonuc)
        print(f"Batch {batch_no} sonuc: HTTP {sonuc['HTTP']} / {sonuc['API Basari']}")

    sonuc_df = pd.DataFrame(sonuc_satirlari)
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    sonuc_dosya = KONYA_CIKTI_DIR / f"konya_tum_stok_gonderim_sonuc_{tarih}.xlsx"
    sonuc_df.to_excel(sonuc_dosya, index=False)

    return sonuc_df, sonuc_dosya


def ornek_kontrol_yap(site_url, token, guncelleme_df):
    kontrol_df = pd.concat([
        guncelleme_df.head(5),
        guncelleme_df.tail(5),
    ], ignore_index=True).drop_duplicates(subset=["ProductCode"])

    kontrol_satirlari = []

    for _, row in kontrol_df.iterrows():
        product_code = urun_kodu_temizle(row["ProductCode"])
        beklenen_stok = int(row["Guncellenecek Stok"])
        urun = urun_getir_product_code(site_url, token, product_code)
        api_stok = urun_stok_degeri(urun)
        api_stok_sayi = sayiya_cevir(api_stok, varsayilan=None)

        kontrol_satirlari.append({
            "ProductId": row["ProductId"],
            "ProductCode": product_code,
            "Beklenen Stok": beklenen_stok,
            "API Okunan Stok": api_stok,
            "Kontrol": "TAMAM" if api_stok_sayi == beklenen_stok else "KONTROL ET",
        })

    kontrol_sonuc_df = pd.DataFrame(kontrol_satirlari)
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    kontrol_dosya = KONYA_CIKTI_DIR / f"konya_tum_stok_ornek_kontrol_{tarih}.xlsx"
    kontrol_sonuc_df.to_excel(kontrol_dosya, index=False)

    print()
    print("Ornek API kontrol sonucu:")
    print(kontrol_sonuc_df.to_string(index=False))
    print()
    print("Ornek kontrol dosyasi:")
    print(kontrol_dosya)


def main():
    print("=== KONYA TUM MARKA T-SOFT STOK GUNCELLEME ===")
    print("Bu islem konya_stok_guncellenecekler.xlsx icindeki KONYA markalarini gonderir.")
    print()

    KONYA_CIKTI_DIR.mkdir(parents=True, exist_ok=True)
    site_url, username, password = config_oku()
    df = konya_stok_listesini_oku()
    guncelleme_df = guncelleme_listesi_hazirla(df)
    on_kontrol_dosya = on_kontrol_raporu_yaz(guncelleme_df)

    print("Site URL:", site_url)
    print("Kullanici:", username)
    print("On kontrol dosyasi:")
    print(on_kontrol_dosya)

    ozet_yaz(guncelleme_df)

    print()
    print("DIKKAT: Bu islem T-Soft canli stoklarini gunceller.")
    onay = input("Tum KONYA stoklarini gondermek icin KONYA yaz: ").strip().upper()

    if onay != "KONYA":
        print("Islem iptal edildi. API'ye stok gonderilmedi.")
        return

    token = api_login(site_url, username, password)
    sonuc_df, sonuc_dosya = stoklari_gonder(site_url, token, guncelleme_df)

    print()
    print("Gonderim tamamlandi.")
    print("Batch sonuc dosyasi:")
    print(sonuc_dosya)

    sorunlu = sonuc_df[sonuc_df["API Basari"].astype(str) != "EVET"]

    if not sorunlu.empty:
        print()
        print("Uyari: Kontrol edilmesi gereken batch var.")
        print(sorunlu.to_string(index=False))

    ornek_kontrol_yap(site_url, token, guncelleme_df)

    print()
    print("KONYA tum marka stok guncelleme tamamlandi.")
    print("Son adim: T-Soft panelde ornek urunleri manuel kontrol et.")


if __name__ == "__main__":
    main()
