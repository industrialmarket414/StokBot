"""Eski toplu gönderimde DIA katkısı eksik kalan ürünleri güvenli biçimde günceller.

Bu araç yalnızca:
  * güncel stok listesinde DIA Stok > 0 olan,
  * eski 2026-08-11 09:37 gönderiminde farklı stokla yer alan,
  * Guncelleme Durumu = GONDERILECEK olan
ürünleri aday kabul eder. Gönderimden önce T-Soft mevcut stokları okunur; zaten
doğru olanlar atlanır. Sonrasında gönderilen bütün ürünler API'den doğrulanır.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

from tsoft_api import (
    api_login,
    config_oku,
    sayiya_cevir,
    stok_batch_gonder,
    urun_getir_product_code,
    urun_getir_product_id,
    urun_kodu_temizle,
    urun_stok_degeri,
)


BASE_DIR = Path(__file__).resolve().parent
CIKTI_DIR = BASE_DIR / "Cikti"
TEST_DIR = CIKTI_DIR / "DIA_EKSIK_GONDERIM"

GUNCEL_STOK_DOSYASI = CIKTI_DIR / "tum_bayiler_stok_guncellenecekler.xlsx"
ESKI_GONDERIM_DOSYASI = (
    CIKTI_DIR / "toplu_tum_stok_gonderim_on_kontrol_20260811_093703.xlsx"
)

BATCH_SIZE = 50


def tsoft_urunu_guvenli_getir(site_url: str, token: str, row: pd.Series):
    """Urunu once kodla, gerekirse ProductId ile bulup kimligini dogrular."""
    beklenen_id = sayiya_cevir(row["ProductId"], varsayilan=None)
    product_code = row["ProductCode"]

    urun = urun_getir_product_code(site_url, token, product_code)
    gelen_id = sayiya_cevir(
        urun.get("ProductId") if isinstance(urun, dict) else None,
        varsayilan=None,
    )
    stok = urun_stok_degeri(urun)

    if gelen_id != beklenen_id or stok is None:
        urun = urun_getir_product_id(site_url, token, row["ProductId"])
        gelen_id = sayiya_cevir(
            urun.get("ProductId") if isinstance(urun, dict) else None,
            varsayilan=None,
        )
        if gelen_id != beklenen_id:
            return None, "BULUNAMADI"
        return urun, "PRODUCT_ID"

    return urun, "PRODUCT_CODE"


def gerekli_kolonlari_kontrol_et(df: pd.DataFrame, kolonlar: list[str], dosya: Path):
    eksik = [kolon for kolon in kolonlar if kolon not in df.columns]
    if eksik:
        raise ValueError(f"{dosya.name} eksik kolonlar: {', '.join(eksik)}")


def dia_adaylarini_hazirla() -> pd.DataFrame:
    if not GUNCEL_STOK_DOSYASI.exists():
        raise FileNotFoundError(
            "Güncel stok listesi bulunamadı. Önce stok_eslestir.py çalıştırın: "
            f"{GUNCEL_STOK_DOSYASI}"
        )
    if not ESKI_GONDERIM_DOSYASI.exists():
        raise FileNotFoundError(
            f"Karşılaştırılacak eski gönderim dosyası bulunamadı: {ESKI_GONDERIM_DOSYASI}"
        )

    guncel = pd.read_excel(GUNCEL_STOK_DOSYASI, dtype=str).fillna("")
    eski = pd.read_excel(ESKI_GONDERIM_DOSYASI, dtype=str).fillna("")

    gerekli_kolonlari_kontrol_et(
        guncel,
        [
            "ProductId",
            "Web Servis Kodu",
            "B2B Marka",
            "Guncellenecek Stok",
            "Guncelleme Durumu",
            "Tedarikciler",
            "DIA Stok",
        ],
        GUNCEL_STOK_DOSYASI,
    )
    gerekli_kolonlari_kontrol_et(
        eski,
        ["ProductId", "Guncellenecek Stok"],
        ESKI_GONDERIM_DOSYASI,
    )

    eski_stok = {
        str(row["ProductId"]).strip(): sayiya_cevir(row["Guncellenecek Stok"])
        for _, row in eski.iterrows()
        if str(row["ProductId"]).strip()
    }

    adaylar = []
    for _, row in guncel.iterrows():
        product_id = str(row["ProductId"]).strip()
        product_code = urun_kodu_temizle(row["Web Servis Kodu"])
        dia_stok = sayiya_cevir(row["DIA Stok"])
        yeni_stok = sayiya_cevir(row["Guncellenecek Stok"])
        eski_deger = eski_stok.get(product_id)
        durum = str(row["Guncelleme Durumu"]).strip().upper()

        if not product_id or not product_code:
            continue
        if durum != "GONDERILECEK":
            continue
        if dia_stok <= 0:
            continue
        if eski_deger is not None and eski_deger == yeni_stok:
            continue

        adaylar.append(
            {
                "ProductId": product_id,
                "ProductCode": product_code,
                "Web Servis Kodu": product_code,
                "B2B Marka": str(row["B2B Marka"]).strip(),
                "Eski Gönderilen Stok": eski_deger,
                "Guncellenecek Stok": yeni_stok,
                "DIA Stok": dia_stok,
                "Fark": yeni_stok - (eski_deger or 0),
                "Tedarikciler": str(row["Tedarikciler"]).strip(),
            }
        )

    sonuc = pd.DataFrame(adaylar)
    if sonuc.empty:
        return sonuc
    if sonuc["ProductId"].duplicated().any():
        tekrarlar = sonuc[sonuc["ProductId"].duplicated(False)]
        raise ValueError(
            "DIA aday listesinde tekrarlanan ProductId var:\n"
            + tekrarlar.to_string(index=False)
        )
    return sonuc.sort_values(
        ["B2B Marka", "Web Servis Kodu"], kind="stable"
    ).reset_index(drop=True)


def api_mevcut_stoklari_oku(
    site_url: str, token: str, adaylar: pd.DataFrame
) -> pd.DataFrame:
    satirlar = []
    toplam = len(adaylar)
    for sira, (_, row) in enumerate(adaylar.iterrows(), 1):
        product_code = row["ProductCode"]
        print(f"API ön kontrol {sira}/{toplam}: {product_code}")
        urun, arama_yontemi = tsoft_urunu_guvenli_getir(site_url, token, row)
        api_ham = urun_stok_degeri(urun)
        api_stok = sayiya_cevir(api_ham, varsayilan=None)
        kayit = row.to_dict()
        kayit["T-Soft Mevcut Stok"] = api_stok
        kayit["API Arama Yontemi"] = arama_yontemi
        kayit["API ProductCode"] = (
            str(urun.get("ProductCode", "")).strip()
            if isinstance(urun, dict)
            else ""
        )
        kayit["Ön Kontrol"] = (
            "OKUNAMADI"
            if api_stok is None
            else "ZATEN DOĞRU"
            if api_stok == int(row["Guncellenecek Stok"])
            else "GÖNDERİLECEK"
        )
        satirlar.append(kayit)
    return pd.DataFrame(satirlar)


def batchlere_bol(df: pd.DataFrame):
    for baslangic in range(0, len(df), BATCH_SIZE):
        yield baslangic // BATCH_SIZE + 1, df.iloc[
            baslangic : baslangic + BATCH_SIZE
        ].copy()


def on_kontrol_yaz(df: pd.DataFrame, etiket: str) -> Path:
    TEST_DIR.mkdir(parents=True, exist_ok=True)
    dosya = TEST_DIR / f"dia_eksik_stok_{etiket}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    df.to_excel(dosya, index=False)
    return dosya


def gonder(site_url: str, token: str, df: pd.DataFrame) -> pd.DataFrame:
    sonuclar = []
    for batch_no, batch_df in batchlere_bol(df):
        print(f"Batch {batch_no} gönderiliyor: {len(batch_df)} ürün")
        sonuc = stok_batch_gonder(site_url, token, batch_no, batch_df)
        sonuclar.append(sonuc)
        print(f"Batch {batch_no}: HTTP {sonuc['HTTP']} / {sonuc['API Basari']}")
        if sonuc["API Basari"] != "EVET":
            print("Batch kontrol gerektiriyor; sonraki batchler gönderilmeyecek.")
            break
    return pd.DataFrame(sonuclar)


def gonderimi_dogrula(
    site_url: str, token: str, gonderilen: pd.DataFrame
) -> pd.DataFrame:
    satirlar = []
    toplam = len(gonderilen)
    for sira, (_, row) in enumerate(gonderilen.iterrows(), 1):
        product_code = row["ProductCode"]
        print(f"API son kontrol {sira}/{toplam}: {product_code}")
        urun, arama_yontemi = tsoft_urunu_guvenli_getir(site_url, token, row)
        api_ham = urun_stok_degeri(urun)
        api_stok = sayiya_cevir(api_ham, varsayilan=None)
        beklenen = int(row["Guncellenecek Stok"])
        satirlar.append(
            {
                "ProductId": row["ProductId"],
                "ProductCode": product_code,
                "B2B Marka": row["B2B Marka"],
                "Beklenen Stok": beklenen,
                "API Okunan Stok": api_stok,
                "API Arama Yontemi": arama_yontemi,
                "Kontrol": "TAMAM" if api_stok == beklenen else "KONTROL ET",
            }
        )
    return pd.DataFrame(satirlar)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Eski gönderimde DIA stoku eksik kalan ürünleri günceller."
    )
    parser.add_argument(
        "--sadece-listele",
        action="store_true",
        help="API bağlantısı ve gönderim yapmadan aday listesini oluşturur.",
    )
    args = parser.parse_args(argv)

    print("=== DIA EKSİK STOK TOPLU GÜNCELLEME ===")
    adaylar = dia_adaylarini_hazirla()
    if adaylar.empty:
        print("Eski gönderime göre eksik DIA stoklu ürün bulunmadı.")
        return 0

    aday_dosya = on_kontrol_yaz(adaylar, "aday_listesi")
    print(f"DIA aday ürün: {len(adaylar)}")
    print(f"Aday kontrol dosyası: {aday_dosya}")
    print()
    print("Marka dağılımı:")
    print(adaylar.groupby("B2B Marka").size().sort_values(ascending=False).to_string())

    if args.sadece_listele:
        print("Sadece listeleme yapıldı. T-Soft API'ye bağlanılmadı ve stok gönderilmedi.")
        return 0

    site_url, username, password = config_oku()
    token = api_login(site_url, username, password)
    kontrol = api_mevcut_stoklari_oku(site_url, token, adaylar)

    okunamayan = kontrol[kontrol["Ön Kontrol"] == "OKUNAMADI"]
    if not okunamayan.empty:
        hata_dosya = on_kontrol_yaz(kontrol, "api_on_kontrol_hatasi")
        print("Bazı ürünlerin T-Soft mevcut stoku okunamadı; güvenlik için gönderim yapılmadı.")
        print(hata_dosya)
        return 1

    gonderilecek = kontrol[kontrol["Ön Kontrol"] == "GÖNDERİLECEK"].copy()
    zaten_dogru = len(kontrol) - len(gonderilecek)
    on_kontrol_dosya = on_kontrol_yaz(kontrol, "api_on_kontrol")
    print()
    print(f"Zaten doğru, atlanacak: {zaten_dogru}")
    print(f"Gerçekten gönderilecek: {len(gonderilecek)}")
    print(f"API ön kontrol dosyası: {on_kontrol_dosya}")

    if gonderilecek.empty:
        print("Tüm DIA stokları zaten doğru. Gönderim yapılmadı.")
        return 0

    beklenen_onay = f"DIA {len(gonderilecek)} GUNCELLE"
    print()
    print("DİKKAT: Yalnız yukarıdaki DIA ürünlerinin canlı T-Soft stokları güncellenecek.")
    onay = input(f"Göndermek için {beklenen_onay} yaz: ").strip().upper()
    if onay != beklenen_onay:
        print("İşlem iptal edildi. API'ye stok gönderilmedi.")
        return 0

    batch_sonuclari = gonder(site_url, token, gonderilecek)
    batch_dosya = on_kontrol_yaz(batch_sonuclari, "batch_sonuclari")
    if (batch_sonuclari["API Basari"] != "EVET").any():
        print("Gönderimde kontrol gerektiren batch var. Sonuç dosyası:")
        print(batch_dosya)
        return 1

    dogrulama = gonderimi_dogrula(site_url, token, gonderilecek)
    dogrulama_dosya = on_kontrol_yaz(dogrulama, "api_son_kontrol")
    sorunlu = dogrulama[dogrulama["Kontrol"] != "TAMAM"]

    print()
    print(f"Gönderilen ürün: {len(gonderilecek)}")
    print(f"Doğrulanan ürün: {len(dogrulama) - len(sorunlu)}")
    print(f"Kontrol gereken ürün: {len(sorunlu)}")
    print(f"Batch sonuçları: {batch_dosya}")
    print(f"API son kontrol: {dogrulama_dosya}")
    return 1 if not sorunlu.empty else 0


if __name__ == "__main__":
    raise SystemExit(main())
