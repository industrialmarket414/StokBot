"""Onaylı iskonto oranlarını T-Soft Ziyaretçiler özel grup fiyatına gönderir.

T-Soft'taki ana satış fiyatı değiştirilmez. Excel'de seçilen satış fiyatından
hesaplanan pozitif iskonto yüzdesi, Ziyaretçiler grubuna "Satış Üzerinden
Yüzde (Kalıcı)" türünde negatif Percentage değeri olarak yazılır.
"""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
from fiyat_kontrol_kurallari import aktarim_dogrula, fiyat_guncelleme_harici_mi


BASE_DIR = Path(__file__).resolve().parent
ONAY_DOSYASI = BASE_DIR / "Cikti" / "fiyat_onay_listesi.xlsx"
TSOFT_CONFIG = BASE_DIR / "Ayarlar" / "tsoft_api.json"
LOG_DIR = BASE_DIR / "Log"
BATCH_SIZE = 100
CUSTOMER_GROUP_CODE = "TZ0"  # [Ziyaretçiler]


def metin_norm(deger) -> str:
    metin = "" if pd.isna(deger) else str(deger).strip().upper()
    metin = metin.translate(
        str.maketrans(
            {"Ç": "C", "Ğ": "G", "İ": "I", "Ö": "O", "Ş": "S", "Ü": "U", "ı": "I"}
        )
    )
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(c for c in metin if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", " ", metin).strip()


def ondalik(deger):
    if deger is None or pd.isna(deger):
        return None
    metin = str(deger).strip().replace("\xa0", "").replace(" ", "")
    if not metin or metin.lower() == "nan" or metin.startswith("="):
        return None
    metin = re.sub(r"[^0-9,.\-]", "", metin)
    if "," in metin and "." in metin:
        if metin.rfind(",") > metin.rfind("."):
            metin = metin.replace(".", "").replace(",", ".")
        else:
            metin = metin.replace(",", "")
    elif "," in metin:
        metin = metin.replace(",", ".")
    try:
        return Decimal(metin)
    except InvalidOperation:
        return None


def kurlari_rapordan_oku(dosya: Path) -> dict:
    try:
        kur_df = pd.read_excel(dosya, sheet_name="Kurlar", dtype=str).fillna("")
    except ValueError as exc:
        raise ValueError("Fiyatlandırma dosyasında Kurlar sayfası bulunamadı") from exc
    sonuc = {"TRY": Decimal("1"), "TL": Decimal("1")}
    for _, row in kur_df.iterrows():
        kod = metin_norm(row.get("Para Birimi", ""))
        kod = "TRY" if kod in {"TL", "TRL", "TRY"} else kod
        if kod not in {"TRY", "USD", "EUR"}:
            continue
        kur = ondalik(row.get("TL Kuru", ""))
        if kod and kur is not None and kur > 0:
            sonuc[kod] = kur
    eksik = [kod for kod in ("USD", "EUR") if kod not in sonuc]
    if eksik:
        raise ValueError(f"Kurlar sayfasında kur eksik: {', '.join(eksik)}")
    return sonuc


def _tedarikci_kolonlari(df: pd.DataFrame):
    kolon_map = {metin_norm(c): c for c in df.columns}
    sonuc = []
    for norm, gercek in kolon_map.items():
        if not norm.endswith(" LISTE"):
            continue
        firma = norm[: -len(" LISTE")]
        iskonto = kolon_map.get(f"{firma} ISKONTO")
        if iskonto:
            sonuc.append((gercek, iskonto))
    return sonuc


def _ortalama_net_alis(row, kolon_ciftleri):
    netler = []
    for liste_col, iskonto_col in kolon_ciftleri:
        liste = ondalik(row.get(liste_col))
        iskonto = ondalik(row.get(iskonto_col))
        if metin_norm(liste_col) == "KONYA LISTE" and iskonto == 0:
            continue
        if (
            liste is None
            or liste <= 0
            or iskonto is None
            or iskonto < 0
            or iskonto >= 100
        ):
            continue
        netler.append(liste * (Decimal("1") - iskonto / Decimal("100")))
    if not netler:
        return None
    return sum(netler) / Decimal(len(netler))


def aktif_urun_dogrula(row) -> None:
    durum = metin_norm(row.get("T-SOFT AKTİFLİK", ""))
    if durum == "PASIF":
        raise ValueError("T-Soft ürünü pasif; fiyat aktarımı engellendi")
    if durum != "AKTIF":
        raise ValueError("T-Soft aktiflik bilgisi yok veya geçersiz")


def gonderim_listesi_hazirla(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    if "KATEGORİ KÂR ORANI" in df.columns and "AKTARIM DURUMU" in df.columns:
        return gonderim_listesi_hazirla_fiyat_onay(df)

    gerekli = [
        "ProductId",
        "Ürün Web Servis Kodu",
        "T-SOFT AKTİFLİK",
        "Fiyat Birimi",
        "MİN. KÂR ORANI (MANUEL)",
        "HEDEF KÂR ORANI (MANUEL)",
        "SEÇİLEN İSKONTO",
        "ONAY",
    ]
    eksik = [k for k in gerekli if k not in df.columns]
    if eksik:
        raise ValueError(f"Fiyatlandırma dosyasında eksik kolonlar: {', '.join(eksik)}")
    kolon_ciftleri = _tedarikci_kolonlari(df)
    if not kolon_ciftleri:
        raise ValueError("B2B LİSTE ve İSKONTO kolon çiftleri bulunamadı")

    hazir, hatalar = [], []
    for excel_satir, (_, row) in enumerate(df.iterrows(), 3):
        if metin_norm(row.get("ONAY", "")) != "EVET":
            continue
        kod = str(row.get("Ürün Web Servis Kodu", "")).strip()
        try:
            if not kod:
                raise ValueError("Web Servis Kodu boş")
            aktif_urun_dogrula(row)
            aktarim_dogrula(kod, ONAY_DOSYASI)
            ortalama = _ortalama_net_alis(row, kolon_ciftleri)
            if ortalama is None:
                raise ValueError("Geçerli B2B liste/iskonto çifti yok")
            min_kar = ondalik(row.get("MİN. KÂR ORANI (MANUEL)"))
            hedef_kar = ondalik(row.get("HEDEF KÂR ORANI (MANUEL)"))
            secilen_iskonto = ondalik(row.get("SEÇİLEN İSKONTO"))
            if min_kar is None or hedef_kar is None:
                raise ValueError("Minimum veya hedef kâr oranı boş")
            if min_kar < 0 or hedef_kar < 0 or min_kar > 100 or hedef_kar > 100:
                raise ValueError("Kâr oranları 0-100 aralığında olmalı")
            if min_kar > hedef_kar:
                raise ValueError("Minimum kâr oranı hedef kâr oranından büyük")
            if secilen_iskonto is None:
                raise ValueError("Seçilen iskonto boş")
            if secilen_iskonto >= 100:
                raise ValueError("Seçilen iskonto 100'den küçük olmalı")
            gecerli_listeler = []
            for liste_col, iskonto_col in kolon_ciftleri:
                liste = ondalik(row.get(liste_col))
                iskonto = ondalik(row.get(iskonto_col))
                if metin_norm(liste_col) == "KONYA LISTE" and iskonto == 0:
                    continue
                if (
                    liste is not None
                    and liste > 0
                    and iskonto is not None
                    and Decimal("0") <= iskonto < Decimal("100")
                ):
                    gecerli_listeler.append(liste)
            if not gecerli_listeler:
                raise ValueError("İskonto hesabı için geçerli B2B liste fiyatı yok")
            ortalama_liste = sum(gecerli_listeler) / Decimal(len(gecerli_listeler))
            secilen = ortalama_liste * (
                Decimal("1") - secilen_iskonto / Decimal("100")
            )
            min_fiyat = ortalama * (Decimal("1") + min_kar / Decimal("100"))
            hedef_fiyat = ortalama * (Decimal("1") + hedef_kar / Decimal("100"))
            tolerans = Decimal("0.011")
            if secilen < min_fiyat - tolerans or secilen > hedef_fiyat + tolerans:
                raise ValueError(
                    f"Seçilen fiyat {min_fiyat:.2f}-{hedef_fiyat:.2f} TL aralığında değil"
                )
            hazir.append(
                {
                    "ProductId": str(row.get("ProductId", "")).strip(),
                    "Web Servis Kodu": kod,
                    "Ortalama Net Alış KDV Hariç": ortalama.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "Ortalama B2B Liste": ortalama_liste.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "Minimum Satış TL KDV Hariç": min_fiyat.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "Hedef Satış TL KDV Hariç": hedef_fiyat.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "Seçilen İskonto": secilen_iskonto,
                    "Seçilen Satış TL KDV Hariç": secilen.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "TSoft Ziyaretçiler Yüzde Farkı": -secilen_iskonto,
                }
            )
        except Exception as exc:
            hatalar.append(
                {
                    "Excel Satır": excel_satir,
                    "Web Servis Kodu": kod,
                    "Hata": str(exc),
                }
            )
    return pd.DataFrame(hazir), hatalar


def gonderim_listesi_hazirla_fiyat_onay(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, list[dict]]:
    gerekli = [
        "ProductId",
        "Ürün Web Servis Kodu",
        "T-SOFT AKTİFLİK",
        "ORTALAMA ALIŞ",
        "KATEGORİ KÂR ORANI",
        "SATIŞ FİYATI",
        "REEL KÂR ORANI",
        "SEÇİLEN İSKONTO",
        "ONAY",
        "AKTARIM DURUMU",
    ]
    eksik = [k for k in gerekli if k not in df.columns]
    if eksik:
        raise ValueError(f"Fiyat Onay sayfasında eksik kolonlar: {', '.join(eksik)}")

    hazir, hatalar = [], []
    for excel_satir, (_, row) in enumerate(df.iterrows(), 3):
        if metin_norm(row.get("ONAY", "")) != "EVET":
            continue
        kod = str(row.get("Ürün Web Servis Kodu", "")).strip()
        try:
            if not kod:
                raise ValueError("Web Servis Kodu boş")
            aktif_urun_dogrula(row)
            aktarim_dogrula(kod, ONAY_DOSYASI)
            durum = str(row.get("AKTARIM DURUMU", "")).strip()
            if durum not in {"AKTARIMA HAZIR", "ZAM YAPILMALI"}:
                raise ValueError(f"Aktarım durumu uygun değil: {durum or 'boş'}")
            ortalama_alis = ondalik(row.get("ORTALAMA ALIŞ"))
            kategori_kar = ondalik(row.get("KATEGORİ KÂR ORANI"))
            satis_fiyati = ondalik(row.get("SATIŞ FİYATI"))
            reel_kar = ondalik(row.get("REEL KÂR ORANI"))
            secilen_iskonto = ondalik(row.get("SEÇİLEN İSKONTO"))
            if ortalama_alis is None or ortalama_alis <= 0:
                raise ValueError("Ortalama alış boş veya sıfır")
            if kategori_kar is None or kategori_kar < 0:
                raise ValueError("Kategori kâr oranı boş veya hatalı")
            if satis_fiyati is None or satis_fiyati <= 0:
                raise ValueError("Satış fiyatı boş veya sıfır")
            if reel_kar is None:
                raise ValueError("Reel kâr oranı hesaplanmamış")
            # Frontend hesaplari Python float olarak JSON'a yazildiginda
            # %7 degeri 6.999999999999984 gibi gelebilir. Gercek bir kar
            # acigini gizlemeyecek kadar kucuk ayni toleransi kullan.
            kar_toleransi = Decimal("0.0001")
            if reel_kar + kar_toleransi < kategori_kar:
                raise ValueError("Reel kâr oranı kategori kâr oranının altında")
            if secilen_iskonto is None:
                raise ValueError("Seçilen iskonto boş")
            if secilen_iskonto >= 100:
                raise ValueError("Seçilen iskonto 100'den küçük olmalı")
            secilen_iskonto = secilen_iskonto.quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            hazir.append(
                {
                    "ProductId": str(row.get("ProductId", "")).strip(),
                    "Web Servis Kodu": kod,
                    "Ortalama Net Alış KDV Hariç": ortalama_alis.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "Kategori Kâr Oranı": kategori_kar,
                    "Seçilen İskonto": secilen_iskonto,
                    "Seçilen Satış TL KDV Hariç": satis_fiyati.quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    ),
                    "Reel Kâr Oranı": reel_kar,
                    "TSoft Ziyaretçiler Yüzde Farkı": -secilen_iskonto,
                }
            )
        except Exception as exc:
            hatalar.append(
                {
                    "Excel Satır": excel_satir,
                    "Web Servis Kodu": kod,
                    "Hata": str(exc),
                }
            )
    return pd.DataFrame(hazir), hatalar


def requests_modulu():
    try:
        import requests
    except ImportError as exc:
        raise RuntimeError("Canlı aktarım için requests paketi gerekli") from exc
    return requests


def token_bul(data):
    if not isinstance(data, dict):
        return None
    for key in ("token", "access_token", "bearer_token"):
        if data.get(key):
            return data[key]
    if isinstance(data.get("data"), dict):
        return token_bul(data["data"])
    if isinstance(data.get("data"), list) and data["data"]:
        return token_bul(data["data"][0])
    return None


def login():
    requests = requests_modulu()
    with open(TSOFT_CONFIG, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    site = str(cfg.get("site_url", "")).strip().rstrip("/")
    user = str(cfg.get("username", "")).strip()
    password = str(cfg.get("password", "")).strip()
    if not site or not user or not password:
        raise ValueError("tsoft_api.json bilgileri eksik")
    response = requests.post(
        f"{site}/rest1/auth/login/{user}",
        data={"pass": password},
        timeout=30,
    )
    token = token_bul(response.json())
    if not token:
        raise RuntimeError("T-Soft token alınamadı")
    return site, token


def api_basarili_mi(data) -> bool:
    if isinstance(data, dict) and data.get("success") is True:
        mesajlar = data.get("message")
        if isinstance(mesajlar, list):
            hatali_mesaj = any(
                isinstance(mesaj, dict)
                and (
                    mesaj.get("type") not in (1, "1", None)
                    or bool(mesaj.get("errorField"))
                )
                for mesaj in mesajlar
            )
            if not hatali_mesaj:
                return True

    metin = json.dumps(data, ensure_ascii=False).lower()
    hata_kelimeleri = ("hata", "basarisiz", "başarısız", "yetkisiz", "unauthorized")
    if any(x in metin for x in hata_kelimeleri):
        return False
    return any(x in metin for x in ("success", "basari", "başarı", "pri001", "pri019", "gni001"))


def api_payload_hazirla(df: pd.DataFrame) -> list[dict]:
    return [
        {
            "ProductCode": str(r["Web Servis Kodu"]).strip(),
            "CustomerGroupCode": CUSTOMER_GROUP_CODE,
            "DiscountType": "1",
            "Percentage": float(r["TSoft Ziyaretçiler Yüzde Farkı"]),
        }
        for _, r in df.iterrows()
    ]


def frontend_json_oku(dosya: Path) -> pd.DataFrame:
    if not dosya.exists():
        raise FileNotFoundError(f"Frontend aktarım verisi yok: {dosya}")
    kayitlar = json.loads(dosya.read_text(encoding="utf-8"))
    if not isinstance(kayitlar, list) or not kayitlar:
        raise ValueError("Frontend aktarım verisi boş veya geçersiz")
    return pd.DataFrame(kayitlar).fillna("")


def tsofta_gonder(df: pd.DataFrame) -> list[dict]:
    # Eski frontend JSON'u / dogrudan cagri da dogrulamayi atlayamaz.
    for kod in df["Web Servis Kodu"]:
        aktarim_dogrula(kod, ONAY_DOSYASI)
    requests = requests_modulu()
    site, token = login()
    LOG_DIR.mkdir(exist_ok=True)
    sonuclar = []
    for baslangic in range(0, len(df), BATCH_SIZE):
        batch_no = baslangic // BATCH_SIZE + 1
        parca = df.iloc[baslangic : baslangic + BATCH_SIZE]
        payload = api_payload_hazirla(parca)
        zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
        istek_log = LOG_DIR / f"tsoft_iskonto_istek_{batch_no:03d}_{zaman}.json"
        istek_log.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        response = requests.post(
            f"{site}/rest1/product/updateCustomerGroupPrice",
            data={"token": token, "data": json.dumps(payload, ensure_ascii=False)},
            headers={"Accept": "application/json"},
            timeout=90,
        )
        try:
            cevap = response.json()
        except Exception:
            cevap = {"raw_response": response.text}
        cevap_log = LOG_DIR / f"tsoft_iskonto_cevap_{batch_no:03d}_{zaman}.json"
        cevap_log.write_text(
            json.dumps(cevap, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        sonuclar.append(
            {
                "Batch": batch_no,
                "Ürün Sayısı": len(parca),
                "HTTP": response.status_code,
                "Başarı": "EVET" if response.ok and api_basarili_mi(cevap) else "KONTROL",
                "İstek Log": str(istek_log),
                "Cevap Log": str(cevap_log),
            }
        )
    return sonuclar


def main(argv=None) -> int:
    global ONAY_DOSYASI
    parser = argparse.ArgumentParser(
        description="Onaylı iskontoları T-Soft Ziyaretçiler özel grup fiyatına gönderir"
    )
    parser.add_argument("--onay", type=Path, default=ONAY_DOSYASI)
    parser.add_argument("--frontend-json", type=Path)
    parser.add_argument("--gonder", action="store_true")
    args = parser.parse_args(argv)
    ONAY_DOSYASI = args.onay
    try:
        if args.frontend_json:
            df = frontend_json_oku(args.frontend_json)
            print(f"Okunan kaynak: Frontend ({len(df)} seçili ürün)")
        else:
            if not args.onay.exists():
                raise FileNotFoundError(f"Fiyatlandırma dosyası yok: {args.onay}")
            excel = pd.ExcelFile(args.onay)
            sayfa = "Fiyat Onay" if "Fiyat Onay" in excel.sheet_names else "Fiyatlandırma"
            df = pd.read_excel(
                args.onay,
                sheet_name=sayfa,
                header=1,
                dtype=str,
            ).fillna("")
            print(f"Okunan sayfa: {sayfa}")

        kod_kolonu = next(
            (kolon for kolon in ("Web Servis Kodu", "Ürün Web Servis Kodu") if kolon in df.columns),
            None,
        )
        if kod_kolonu:
            haric_maske = df[kod_kolonu].map(fiyat_guncelleme_harici_mi)
            haric_sayisi = int(haric_maske.sum())
            if haric_sayisi:
                haric_kodlar = sorted(
                    {
                        str(kod).strip()
                        for kod in df.loc[haric_maske, kod_kolonu]
                        if str(kod).strip()
                    }
                )
                print(
                    "Fiyat guncellemesi disinda birakilan: "
                    f"{haric_sayisi} ({', '.join(haric_kodlar)})"
                )
                df = df.loc[~haric_maske].copy()

        hazir, hatalar = gonderim_listesi_hazirla(df)
        print(f"Onaylı ve geçerli iskonto: {len(hazir)}")
        print(f"Onaylı fakat hatalı iskonto: {len(hatalar)}")
        for hata in hatalar[:20]:
            print(hata)
        if not args.gonder:
            print("Kontrol modu: T-Soft'a iskonto gönderilmedi.")
            return 0
        if hatalar:
            raise RuntimeError("Onaylı satırlarda hata var; hiçbir iskonto gönderilmedi")
        if hazir.empty:
            raise RuntimeError("Gönderilecek onaylı iskonto yok")
        sonuclar = tsofta_gonder(hazir)
        basarili = [s for s in sonuclar if s.get("Başarı") == "EVET"]
        kontrol = [s for s in sonuclar if s.get("Başarı") != "EVET"]
        basarili_urun = sum(int(s.get("Ürün Sayısı", 0)) for s in basarili)
        print("=" * 56)
        if not kontrol:
            print("T-SOFT GÜNCELLEMESİ BAŞARIYLA TAMAMLANDI")
            print(f"Başarıyla güncellenen ürün: {basarili_urun}")
            print(f"Başarılı batch: {len(basarili)}/{len(sonuclar)}")
            print("Hatalı batch: 0")
        else:
            print("T-SOFT GÜNCELLEMESİ KONTROL GEREKTİRİYOR")
            print(f"Başarıyla güncellenen ürün: {basarili_urun}")
            print(f"Başarılı batch: {len(basarili)}/{len(sonuclar)}")
            print(f"Kontrol gereken batch: {len(kontrol)}")
            for sonuc in kontrol:
                print(json.dumps(sonuc, ensure_ascii=False))
        print("=" * 56)
        if kontrol:
            return 1
        return 0
    except Exception as exc:
        print(f"HATA: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
