"""Stoklu T-Soft urun linkleri icin DAILY sitemap CSV'si hazirlar.

Mevcut YEARLY ureticinin stok ve link eslestirme kurallarini aynen kullanir;
yalnizca SiteMapChangeFreq alanini DAILY olarak yazar. T-Soft'a kendi basina
veri gondermez.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import sys

from tsoft_stoklu_urun_sitemap_yillik import (
    CIKTI_DIR,
    VARSAYILAN_LINK_DOSYASI,
    VARSAYILAN_STOK_DOSYASI,
    csv_yaz,
    eslesmeyenleri_yaz,
    link_csv_oku,
    stoklu_urunleri_oku,
    yillik_link_satirlari_hazirla,
)


def gunluk_link_satirlari_hazirla(
    stoklular: dict[str, dict[str, object]],
    linkler: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[str]]:
    """Yalniz urun linklerini eslestirir ve sitemap sikligini DAILY yapar."""
    satirlar, eslesmeyen = yillik_link_satirlari_hazirla(stoklular, linkler)
    for satir in satirlar:
        satir["SiteMapChangeFreq"] = "DAILY"
    return satirlar, eslesmeyen


def sitemap_csv_olustur(
    stok_dosyasi: Path = VARSAYILAN_STOK_DOSYASI,
    link_dosyasi: Path = VARSAYILAN_LINK_DOSYASI,
    cikti: Path | None = None,
) -> dict[str, object]:
    """DAILY ice aktarim CSV'sini olusturur ve islem ozetini dondurur."""
    stok_dosyasi = Path(stok_dosyasi).resolve()
    link_dosyasi = Path(link_dosyasi).resolve()
    if not stok_dosyasi.is_file():
        raise FileNotFoundError(f"Stok dosyasi bulunamadi: {stok_dosyasi}")
    if not link_dosyasi.is_file():
        raise FileNotFoundError(
            "T-Soft link disari aktarim dosyasi bulunamadi. CSV Urun Sihirbazi > "
            f"Ayarlar > Linkleri Disari Aktar dosyasini buraya koyun: {link_dosyasi}"
        )

    stoklular = stoklu_urunleri_oku(stok_dosyasi)
    linkler, encoding, ayirac = link_csv_oku(link_dosyasi)
    satirlar, eslesmeyen = gunluk_link_satirlari_hazirla(stoklular, linkler)
    if not satirlar:
        raise ValueError("Stoklu urunlerle eslesen hicbir T-Soft urun linki bulunamadi.")

    zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
    cikti = Path(cikti).resolve() if cikti else (
        CIKTI_DIR / f"tsoft_stoklu_urun_sitemap_gunluk_{zaman}.csv"
    )
    csv_yaz(satirlar, cikti)

    eslesmeyen_dosyasi = None
    if eslesmeyen:
        eslesmeyen_dosyasi = cikti.with_name(cikti.stem + "_eslesmeyen.csv")
        eslesmeyenleri_yaz(eslesmeyen, stoklular, eslesmeyen_dosyasi)

    return {
        "stoklu_sayisi": len(stoklular),
        "link_sayisi": len(linkler),
        "gunluk_sayisi": len(satirlar),
        "eslesmeyen_sayisi": len(eslesmeyen),
        "encoding": encoding,
        "ayirac": ayirac,
        "cikti": cikti,
        "eslesmeyen_dosyasi": eslesmeyen_dosyasi,
    }


def argumanlar() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Stoklu urun linklerini DAILY yapmak icin T-Soft ice aktarim CSV'si olusturur."
    )
    parser.add_argument("--stok", type=Path, default=VARSAYILAN_STOK_DOSYASI)
    parser.add_argument("--link-export", type=Path, default=VARSAYILAN_LINK_DOSYASI)
    parser.add_argument("--cikti", type=Path, help="Olusturulacak CSV yolu")
    return parser.parse_args()


def main() -> int:
    args = argumanlar()
    try:
        sonuc = sitemap_csv_olustur(args.stok, args.link_export, args.cikti)
        print(f"Stoklu benzersiz urun: {sonuc['stoklu_sayisi']}")
        print(f"DAILY yapilacak urun linki: {sonuc['gunluk_sayisi']}")
        print(f"Hazir T-Soft ice aktarim dosyasi: {sonuc['cikti']}")
        if sonuc["eslesmeyen_sayisi"]:
            print(f"Link ihracinda bulunamayan stoklu urun: {sonuc['eslesmeyen_sayisi']}")
            print(f"Eslesmeyen urunler: {sonuc['eslesmeyen_dosyasi']}")
        else:
            print("Tum stoklu urunler link ihraciyla eslesti.")
        print("Bu betik T-Soft'a veri gondermedi; yalnizca kontrol edilebilir CSV olusturdu.")
        return 0
    except (OSError, ValueError) as exc:
        print(f"HATA: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
