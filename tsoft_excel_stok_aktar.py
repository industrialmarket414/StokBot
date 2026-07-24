"""Secilen CSV dosyasindaki urun stoklarini T-Soft API'ye aktarir.

CSV'den yalnizca "Web Servisi" ve "Stok" kolonlari okunur. API'ye de
yalnizca ProductCode ve Stock alanlari gonderilir; diger urun bilgileri
degistirilmez.
"""

from __future__ import annotations

import argparse
import csv
from datetime import datetime
import json
from pathlib import Path
import re
import sys
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
AYARLAR_DIR = BASE_DIR / "Ayarlar"
LOG_DIR = BASE_DIR / "Log"
CONFIG_FILE = AYARLAR_DIR / "tsoft_api.json"
VARSAYILAN_BATCH = 100


def requests_modulu():
    try:
        import requests
    except ImportError as exc:
        raise RuntimeError(
            "Gercek API gonderimi icin 'requests' paketi gerekli. "
            "Bu projeyi calistirdiginiz Python'a requests paketini kurun."
        ) from exc
    return requests


def metin_norm(deger: object) -> str:
    metin = str(deger or "").strip().casefold()
    ceviri = str.maketrans({"ı": "i", "ş": "s", "ğ": "g", "ü": "u", "ö": "o", "ç": "c"})
    return re.sub(r"[^a-z0-9]", "", metin.translate(ceviri))


def urun_kodu_temizle(deger: object) -> str:
    # ProductCode T-Soft'taki degerle birebir ayni kalmali. Ornegin Bosch
    # kodlarindaki "0 601 9G9 100" ic bosluklari urun kodunun parcasidir.
    # Yalnizca hucre basindaki/sonundaki bosluklari temizle.
    return str(deger or "").replace("\u00a0", " ").strip()


def stok_cevir(deger: object, satir_no: int) -> int:
    metin = str(deger or "").strip().replace(" ", "")
    if not metin:
        raise ValueError(f"{satir_no}. satirda Stok bos.")

    # 1.000 gibi tam sayi binlik ayraclarini, 10,0 gibi ondaliklari destekle.
    if re.fullmatch(r"[-+]?\d{1,3}(?:\.\d{3})+", metin):
        metin = metin.replace(".", "")
    else:
        metin = metin.replace(",", ".")

    try:
        sayi = float(metin)
    except ValueError as exc:
        raise ValueError(f"{satir_no}. satirdaki Stok sayi degil: {deger!r}") from exc

    if not sayi.is_integer():
        raise ValueError(f"{satir_no}. satirdaki Stok tam sayi degil: {deger!r}")
    if sayi < 0:
        raise ValueError(f"{satir_no}. satirdaki Stok negatif olamaz: {deger!r}")

    return int(sayi)


def config_oku() -> tuple[str, str, str]:
    AYARLAR_DIR.mkdir(exist_ok=True)
    if not CONFIG_FILE.exists():
        ornek = {"site_url": "https://www.siteadresiniz.com", "username": "", "password": ""}
        CONFIG_FILE.write_text(json.dumps(ornek, ensure_ascii=False, indent=4), encoding="utf-8")
        raise RuntimeError(
            f"Ayar dosyasi olusturuldu: {CONFIG_FILE}\n"
            "site_url, username ve password alanlarini doldurup tekrar calistirin."
        )

    data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    site_url = str(data.get("site_url", "")).strip().rstrip("/")
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", "")).strip()
    if not site_url or not username or not password:
        raise RuntimeError(f"{CONFIG_FILE} icinde site_url, username veya password bos.")
    return site_url, username, password


def token_bul(data: Any) -> str | None:
    if not isinstance(data, dict):
        return None
    for kaynak in (data, data.get("data")):
        if isinstance(kaynak, dict):
            for anahtar in ("token", "access_token", "bearer_token"):
                if kaynak.get(anahtar):
                    return str(kaynak[anahtar])
        elif isinstance(kaynak, list) and kaynak and isinstance(kaynak[0], dict):
            for anahtar in ("token", "access_token", "bearer_token"):
                if kaynak[0].get(anahtar):
                    return str(kaynak[0][anahtar])
    return None


def api_login(site_url: str, username: str, password: str) -> str:
    requests = requests_modulu()
    try:
        response = requests.post(
            f"{site_url}/rest1/auth/login/{username}",
            data={"pass": password},
            headers={"Accept": "application/json"},
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(f"T-Soft oturumu acilamadi: {exc}") from exc
    try:
        data = response.json()
    except ValueError as exc:
        raise RuntimeError(f"T-Soft giris cevabi JSON degil: {response.text[:500]}") from exc

    token = token_bul(data)
    if not token:
        raise RuntimeError(f"T-Soft token alinamadi. API cevabi: {json.dumps(data, ensure_ascii=False)[:800]}")
    return token


def csv_metin_oku(dosya: Path) -> tuple[str, str]:
    ham = dosya.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
        try:
            return ham.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV karakter kodlamasi okunamadi.")


def csv_oku(dosya: Path) -> tuple[list[dict[str, object]], list[str], str]:
    metin, encoding = csv_metin_oku(dosya)
    ornek = metin[:8192]
    try:
        dialect = csv.Sniffer().sniff(ornek, delimiters=",;\t|")
        ayirac = dialect.delimiter
    except csv.Error:
        ayirac = ";" if ornek.count(";") > ornek.count(",") else ","

    okuyucu = csv.DictReader(metin.splitlines(), delimiter=ayirac)
    if not okuyucu.fieldnames:
        raise ValueError("CSV baslik satiri bulunamadi.")

    basliklar = {metin_norm(baslik): baslik for baslik in okuyucu.fieldnames if baslik}
    kod_basligi = basliklar.get("webservisi") or basliklar.get("webserviskodu")
    stok_basligi = basliklar.get("stok")
    if not kod_basligi or not stok_basligi:
        bulunan = ", ".join(str(x) for x in okuyucu.fieldnames)
        raise ValueError(
            "CSV'de 'Web Servisi' ve 'Stok' kolonlari zorunludur. "
            f"Bulunan kolonlar: {bulunan}"
        )

    son_kayitlar: dict[str, tuple[int, dict[str, object]]] = {}
    hatalar: list[str] = []
    tekrarlar: list[str] = []

    for satir_no, satir in enumerate(okuyucu, start=2):
        kod = urun_kodu_temizle(satir.get(kod_basligi))
        if not kod:
            hatalar.append(f"{satir_no}. satirda Web Servisi bos.")
            continue
        try:
            stok = stok_cevir(satir.get(stok_basligi), satir_no)
        except ValueError as exc:
            hatalar.append(str(exc))
            continue

        anahtar = kod.casefold()
        if anahtar in son_kayitlar:
            tekrarlar.append(kod)
        son_kayitlar[anahtar] = (satir_no, {"ProductCode": kod, "Stock": stok})

    kayitlar = [deger[1] for deger in sorted(son_kayitlar.values(), key=lambda x: x[0])]
    if hatalar:
        ozet = "\n".join(f"- {hata}" for hata in hatalar[:30])
        kalan = len(hatalar) - 30
        if kalan > 0:
            ozet += f"\n- ... ve {kalan} hata daha"
        raise ValueError(f"CSV gonderilmedi; hatali satirlar var:\n{ozet}")
    if not kayitlar:
        raise ValueError("CSV'de gonderilecek gecerli urun bulunamadi.")
    return kayitlar, sorted(set(tekrarlar)), f"{encoding}, ayirac={ayirac!r}"


def cevap_hata_mi(data: Any) -> bool:
    """Acik bir API hatasi varsa True dondurur; ``error: false`` hata sayilmaz."""
    if isinstance(data, dict):
        for anahtar, deger in data.items():
            norm_anahtar = metin_norm(anahtar)
            if norm_anahtar in {"error", "errors", "hata", "hatalar"}:
                if deger not in (None, False, 0, "", [], {}):
                    return True
            if norm_anahtar in {"success", "basari", "basarili"} and deger is False:
                return True
            if cevap_hata_mi(deger):
                return True
        return False
    if isinstance(data, list):
        return any(cevap_hata_mi(deger) for deger in data)
    if isinstance(data, str):
        metin = data.casefold()
        return any(kelime in metin for kelime in ("basarisiz", "başarısız", "unauthorized", "yetkisiz"))
    return False


def stoklari_gonder(
    site_url: str,
    token: str,
    kayitlar: list[dict[str, object]],
    batch_boyutu: int,
    baslangic_batch: int = 1,
) -> tuple[int, Path, Path | None, int]:
    LOG_DIR.mkdir(exist_ok=True)
    zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_dosyasi = LOG_DIR / f"tsoft_csv_stok_aktar_{zaman}.json"
    loglar: list[dict[str, object]] = []
    hatalilar: list[dict[str, object]] = []
    gonderilen = 0
    requests = requests_modulu()

    def istek_yap(veri: list[dict[str, object]], aciklama: str) -> tuple[Any, int, bool]:
        try:
            response = requests.post(
                f"{site_url}/rest1/product/updateProducts",
                data={"token": token, "data": json.dumps(veri, ensure_ascii=False)},
                headers={"Accept": "application/json"},
                timeout=90,
            )
        except requests.RequestException as exc:
            raise RuntimeError(f"{aciklama} T-Soft'a gonderilemedi: {exc}") from exc
        try:
            cevap: Any = response.json()
        except ValueError:
            cevap = {"raw_response": response.text[:3000]}
        basarili = response.ok and not cevap_hata_mi(cevap)
        return cevap, response.status_code, basarili

    try:
        for baslangic in range(0, len(kayitlar), batch_boyutu):
            batch = kayitlar[baslangic : baslangic + batch_boyutu]
            batch_no = baslangic // batch_boyutu + 1
            if batch_no < baslangic_batch:
                continue
            print(f"Batch {batch_no}: {len(batch)} urun gonderiliyor...")
            cevap, http, basarili = istek_yap(batch, f"Batch {batch_no}")
            loglar.append(
                {
                    "batch": batch_no,
                    "urun_sayisi": len(batch),
                    "http": http,
                    "basarili": basarili,
                    "istek": batch,
                    "cevap": cevap,
                }
            )
            if basarili:
                gonderilen += len(batch)
                continue

            if http == 429 or http >= 500:
                raise RuntimeError(
                    f"Batch {batch_no} gecici sunucu/limit hatasi verdi (HTTP {http}). "
                    "Sonraki batchler gonderilmedi; daha sonra ayni batch'ten devam edin."
                )

            print(
                f"Batch {batch_no} toplu cevapta hata verdi (HTTP {http}). "
                "Sorunlu urunu bulmak icin bu batch tek tek deneniyor..."
            )
            for sira, urun in enumerate(batch, start=1):
                tek_cevap, tek_http, tek_basarili = istek_yap(
                    [urun], f"Batch {batch_no}, urun {sira}"
                )
                loglar.append(
                    {
                        "batch": batch_no,
                        "tek_urun_sira": sira,
                        "http": tek_http,
                        "basarili": tek_basarili,
                        "istek": urun,
                        "cevap": tek_cevap,
                    }
                )
                if tek_basarili:
                    gonderilen += 1
                    continue
                if tek_http == 429 or tek_http >= 500:
                    raise RuntimeError(
                        f"Batch {batch_no}, urun {sira} gecici sunucu/limit hatasi verdi "
                        f"(HTTP {tek_http}). Islem durduruldu."
                    )
                cevap_ozeti = json.dumps(tek_cevap, ensure_ascii=False)
                if len(cevap_ozeti) > 600:
                    cevap_ozeti = cevap_ozeti[:600] + "..."
                raise RuntimeError(
                    "Urun atlanmadi; guvenlik icin islem durduruldu. "
                    f"Web Servis Kodu: {urun.get('ProductCode', '')} | "
                    f"Stok: {urun.get('Stock', '')} | HTTP: {tek_http} | "
                    f"T-Soft cevabi: {cevap_ozeti}"
                )
    finally:
        log_dosyasi.write_text(json.dumps(loglar, ensure_ascii=False, indent=2), encoding="utf-8")

    hatali_dosyasi: Path | None = None
    if hatalilar:
        hatali_dosyasi = LOG_DIR / f"tsoft_csv_stok_hatalilar_{zaman}.csv"
        with hatali_dosyasi.open("w", encoding="utf-8-sig", newline="") as dosya:
            yazici = csv.DictWriter(dosya, fieldnames=list(hatalilar[0].keys()), delimiter=";")
            yazici.writeheader()
            yazici.writerows(hatalilar)

    return gonderilen, log_dosyasi, hatali_dosyasi, len(hatalilar)


def dosya_sectir() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        secilen = filedialog.askopenfilename(
            title="T-Soft'a gonderilecek CSV dosyasini secin",
            filetypes=[("CSV dosyalari", "*.csv"), ("Tum dosyalar", "*.*")],
        )
        root.destroy()
        return Path(secilen) if secilen else None
    except Exception as exc:
        raise RuntimeError("Dosya secme penceresi acilamadi; CSV yolunu komut satirinda verin.") from exc


def argumanlar() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="CSV'deki Web Servisi ve Stok alanlarini T-Soft'a gonderir.")
    parser.add_argument("csv", nargs="?", type=Path, help="CSV dosya yolu; verilmezse dosya secme penceresi acilir")
    parser.add_argument("--batch", type=int, default=VARSAYILAN_BATCH, help="Bir istekteki urun sayisi")
    parser.add_argument(
        "--baslangic-batch",
        type=int,
        default=1,
        help="Onceki basarili batchleri atlayip bu batch numarasindan baslar",
    )
    parser.add_argument("--kontrol", action="store_true", help="CSV'yi kontrol eder, API'ye gondermez")
    parser.add_argument("--evet", action="store_true", help="Komut satiri onayini atlar")
    return parser.parse_args()


def main() -> int:
    args = argumanlar()
    try:
        if args.batch < 1 or args.batch > 500:
            raise ValueError("--batch 1 ile 500 arasinda olmalidir.")
        if args.baslangic_batch < 1:
            raise ValueError("--baslangic-batch en az 1 olmalidir.")
        csv_dosyasi = args.csv or dosya_sectir()
        if not csv_dosyasi:
            print("Dosya secilmedi; islem iptal edildi.")
            return 0
        csv_dosyasi = csv_dosyasi.resolve()
        if not csv_dosyasi.is_file():
            raise FileNotFoundError(f"CSV bulunamadi: {csv_dosyasi}")

        kayitlar, tekrarlar, csv_bilgisi = csv_oku(csv_dosyasi)
        print(f"CSV: {csv_dosyasi}")
        print(f"Okuma: {csv_bilgisi}")
        print(f"Gecerli ve benzersiz urun: {len(kayitlar)}")
        toplam_batch = (len(kayitlar) + args.batch - 1) // args.batch
        if args.baslangic_batch > toplam_batch:
            raise ValueError(
                f"--baslangic-batch {args.baslangic_batch} gecersiz; toplam batch: {toplam_batch}."
            )
        if args.baslangic_batch > 1:
            atlanan = min((args.baslangic_batch - 1) * args.batch, len(kayitlar))
            print(f"Ilk {atlanan} urun atlanacak; batch {args.baslangic_batch}'ten baslanacak.")
        if tekrarlar:
            print(f"Tekrar eden kod: {len(tekrarlar)} (her kodun son satiri kullanilacak)")

        if args.kontrol:
            print("Kontrol tamamlandi; --kontrol nedeniyle API'ye veri gonderilmedi.")
            return 0

        if not args.evet:
            cevap = input(
                "T-Soft'ta bu urunlerin SADECE stok degeri degisecek. "
                "Devam etmek icin EVET yazin: "
            ).strip()
            if cevap.casefold() != "evet":
                print("Onay verilmedi; API'ye veri gonderilmedi.")
                return 0

        site_url, username, password = config_oku()
        print(f"T-Soft oturumu aciliyor: {site_url}")
        token = api_login(site_url, username, password)
        gonderilen, log_dosyasi, hatali_dosyasi, hata_sayisi = stoklari_gonder(
            site_url, token, kayitlar, args.batch, args.baslangic_batch
        )
        print(f"Tamamlandi. Bu calistirmada gonderilen urun: {gonderilen}")
        print(f"API cevap logu: {log_dosyasi}")
        if hatali_dosyasi:
            print(f"Gonderilemeyen urun: {hata_sayisi}")
            print(f"Hatali urunler: {hatali_dosyasi}")
            return 2
        return 0
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError) as exc:
        print(f"HATA: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
