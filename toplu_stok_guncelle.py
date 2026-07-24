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

CONFIG_FILE = AYARLAR_DIR / "tsoft_api.json"
TOPLU_STOK_DOSYA = CIKTI_DIR / "tum_bayiler_stok_guncellenecekler.xlsx"

BATCH_SIZE = 100
YALNIZCA_YILDIZDAN_GUNCELLENECEK_MARKALAR = {
    "VIKO",
    "VIKO-PANASONIC",
    "VIKO PANASONIC",
    "ARTLINE",
    "THEA",
    "ACK",
}


def urun_kodu_temizle(deger):
    if pd.isna(deger):
        return ""

    # T-Soft kodlarinda anlamli bosluklar olabilir (ornegin Bosch kodlari).
    return re.sub(r"\s+", " ", str(deger)).strip()


def marka_norm(deger):
    return str(deger).strip().upper().translate(str.maketrans({
        "Ç": "C", "Ğ": "G", "İ": "I", "Ö": "O", "Ş": "S", "Ü": "U",
    }))


def marka_adi_standart(deger):
    norm = marka_norm(deger)
    return "SCHNEIDER ELECTRIC" if norm == "SCHNEIDER" else str(deger).strip().upper()


def sayiya_cevir(deger, varsayilan=0):
    if pd.isna(deger):
        return varsayilan

    metin = re.sub(r"\s+", "", str(deger).strip())

    if metin == "" or metin.lower() == "nan":
        return varsayilan

    if "YOK" in metin.upper():
        return 0

    # Stoklar tam sayidir. Turkce binlik ayiracini ondalik sanma:
    # 1.344 -> 1344, 12.500 -> 12500, ancak 5.0 -> 5.
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
        sayi_eslesmesi = re.search(r"-?\d+(?:[.,]\d+)*", metin)

        if not sayi_eslesmesi:
            return varsayilan

        try:
            return sayiya_cevir(sayi_eslesmesi.group(0), varsayilan)
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
    log_path = LOG_DIR / "tsoft_toplu_tum_stok_login_response.json"

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

    # T-Soft cevabinda acik bir basari degeri varsa once onu kullan.
    # Cevaplarda hata olmasa bile "errorField": [] alani bulundugu icin
    # tum JSON icinde "error" aramak basarili islemleri KONTROL gosteriyordu.
    if isinstance(data, dict):
        success = data.get("success")

        if success is True:
            return True

        if success is False:
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


def toplu_stok_listesini_oku():
    if not TOPLU_STOK_DOSYA.exists():
        print("TOPLU stok dosyasi bulunamadi:")
        print(TOPLU_STOK_DOSYA)
        raise SystemExit

    df = pd.read_excel(TOPLU_STOK_DOSYA, dtype=str).fillna("")

    gerekli_kolonlar = [
        "ProductId",
        "Web Servis Kodu",
        "Guncellenecek Stok",
    ]

    for kolon in gerekli_kolonlar:
        if kolon not in df.columns:
            print(f"Excel icinde eksik kolon var: {kolon}")
            raise SystemExit

    return df


def guncelleme_listesi_hazirla(df):
    hazir_satirlar = []
    hatalar = []
    konya_hes_haric_sayi = 0
    yildiz_tek_kaynak_haric_sayi = 0
    gonderilmeyecek_haric_sayi = 0
    oznur_guvenlik_haric_sayi = 0

    for index, row in df.iterrows():
        b2b_marka = marka_adi_standart(row.get("B2B Marka", ""))
        tedarikciler = str(row.get("Tedarikciler", "")).strip()
        product_code = urun_kodu_temizle(row.get("Web Servis Kodu", ""))

        if str(row.get("Guncelleme Durumu", "")).strip().upper() == "GONDERILMEYECEK":
            gonderilmeyecek_haric_sayi += 1
            continue

        # Ikinci guvenlik: eski veya elle degistirilmis bir kontrol Excel'i
        # kullanilsa bile OZNUR canli T-Soft guncellemesine gecemez.
        if marka_norm(b2b_marka) in {"OZNUR", "OZNUR KABLO"}:
            oznur_guvenlik_haric_sayi += 1
            continue

        if (
            marka_norm(b2b_marka) == "HES"
            and re.search(r"(^|[^A-Z0-9])KONYA([^A-Z0-9]|$)", marka_norm(tedarikciler))
        ):
            konya_hes_haric_sayi += 1
            continue

        if (
            marka_norm(b2b_marka) in YALNIZCA_YILDIZDAN_GUNCELLENECEK_MARKALAR
            and marka_norm(tedarikciler) != "YILDIZ"
        ):
            yildiz_tek_kaynak_haric_sayi += 1
            continue

        product_id = str(row["ProductId"]).strip()
        guncellenecek_stok = sayiya_cevir(row["Guncellenecek Stok"])

        if not product_id:
            hatalar.append(f"Excel satir {index + 2}: ProductId bos.")
            continue

        if not product_code:
            hatalar.append(f"ProductId {product_id}: Web Servis Kodu bos.")
            continue

        hazir_satirlar.append({
            "ProductId": product_id,
            "ProductCode": product_code,
            "Web Servis Kodu": product_code,
            "Guncellenecek Stok": guncellenecek_stok,
            "B2B Marka": b2b_marka,
            "Model": str(row.get("Model", "")).strip(),
            "Tedarikciler": tedarikciler,
        })

    if konya_hes_haric_sayi:
        print(
            "KONYA + HES guncelleme disi birakilan satir: "
            f"{konya_hes_haric_sayi}"
        )

    if yildiz_tek_kaynak_haric_sayi:
        print(
            "Yalnizca YILDIZ kaynagindan guncellenecek marka satiri atlandi: "
            f"{yildiz_tek_kaynak_haric_sayi}"
        )

    if gonderilmeyecek_haric_sayi:
        print(
            "GONDERILMEYECEK durumuyla guncelleme disi birakilan satir: "
            f"{gonderilmeyecek_haric_sayi}"
        )

    if oznur_guvenlik_haric_sayi:
        print(
            "OZNUR ikinci guvenlik kontroluyla guncelleme disi birakilan satir: "
            f"{oznur_guvenlik_haric_sayi}"
        )

    if hatalar:
        tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
        hata_dosya = LOG_DIR / f"tsoft_toplu_tum_stok_hazirlik_hatalari_{tarih}.txt"
        LOG_DIR.mkdir(exist_ok=True)

        with open(hata_dosya, "w", encoding="utf-8") as f:
            for hata in hatalar:
                f.write(hata + "\n")

        print("Hazirlik hatalari var. API'ye gonderim yapilmadi.")
        print("Hata dosyasi:")
        print(hata_dosya)
        raise SystemExit

    if not hazir_satirlar:
        print("Gonderilecek TOPLU stok satiri yok.")
        raise SystemExit

    return pd.DataFrame(hazir_satirlar)


def on_kontrol_raporu_yaz(guncelleme_df, dosya_etiketi="toplu_tum_stok"):
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    rapor_dosya = CIKTI_DIR / f"{dosya_etiketi}_gonderim_on_kontrol_{tarih}.xlsx"
    guncelleme_df.to_excel(rapor_dosya, index=False)
    return rapor_dosya


def ozet_yaz(guncelleme_df):
    print()
    print("Tum bayiler toplu stok guncelleme ozeti:")
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


def stok_batch_gonder(
    site_url,
    token,
    batch_no,
    batch_df,
    dosya_etiketi="toplu_tum_stok",
):
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
    log_path = LOG_DIR / f"tsoft_{dosya_etiketi}_batch_{batch_no:03d}_{tarih}.json"

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


def stoklari_gonder(
    site_url,
    token,
    guncelleme_df,
    dosya_etiketi="toplu_tum_stok",
):
    sonuc_satirlari = []

    for batch_no, batch_df in batchlere_bol(guncelleme_df, BATCH_SIZE):
        print(f"Batch {batch_no} gonderiliyor... ({len(batch_df)} urun)")

        sonuc = stok_batch_gonder(
            site_url=site_url,
            token=token,
            batch_no=batch_no,
            batch_df=batch_df,
            dosya_etiketi=dosya_etiketi,
        )

        sonuc_satirlari.append(sonuc)
        print(f"Batch {batch_no} sonuc: HTTP {sonuc['HTTP']} / {sonuc['API Basari']}")

    sonuc_df = pd.DataFrame(sonuc_satirlari)
    tarih = datetime.now().strftime("%Y%m%d_%H%M%S")
    sonuc_dosya = CIKTI_DIR / f"{dosya_etiketi}_gonderim_sonuc_{tarih}.xlsx"
    sonuc_df.to_excel(sonuc_dosya, index=False)

    return sonuc_df, sonuc_dosya


def ornek_kontrol_yap(
    site_url,
    token,
    guncelleme_df,
    dosya_etiketi="toplu_tum_stok",
    tumunu_kontrol_et=False,
):
    if tumunu_kontrol_et:
        kontrol_df = guncelleme_df.drop_duplicates(subset=["ProductCode"]).copy()
    else:
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
    kontrol_dosya = CIKTI_DIR / f"{dosya_etiketi}_ornek_kontrol_{tarih}.xlsx"
    kontrol_sonuc_df.to_excel(kontrol_dosya, index=False)

    print()
    print("Ornek API kontrol sonucu:")
    print(kontrol_sonuc_df.to_string(index=False))
    print()
    print("Ornek kontrol dosyasi:")
    print(kontrol_dosya)


def main():
    print("=== TOPLU TUM MARKA T-SOFT STOK GUNCELLEME ===")
    print("Bu islem tum_bayiler_stok_guncellenecekler.xlsx icindeki urunleri gonderir.")
    print()

    site_url, username, password = config_oku()
    df = toplu_stok_listesini_oku()
    guncelleme_df = guncelleme_listesi_hazirla(df)
    on_kontrol_dosya = on_kontrol_raporu_yaz(guncelleme_df)

    print("Site URL:", site_url)
    print("Kullanici:", username)
    print("On kontrol dosyasi:")
    print(on_kontrol_dosya)

    ozet_yaz(guncelleme_df)

    print()
    print("DIKKAT: Bu islem T-Soft canli stoklarini gunceller.")
    onay = input("Tum TOPLU stoklarini gondermek icin TOPLU yaz: ").strip().upper()

    if onay != "TOPLU":
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
    print("Tum bayiler toplu stok guncelleme tamamlandi.")
    print("Son adim: T-Soft panelde ornek urunleri manuel kontrol et.")


if __name__ == "__main__":
    main()
