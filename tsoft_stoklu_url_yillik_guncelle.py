"""Stok aktarimini tekrarlamadan stoklu urun URL'lerini YEARLY yapar.

Son stok eslestirme dosyasindan T-Soft link CSV'sini hazirlar ve kullanicinin
YUKLE onayindan sonra panelde Linkleri Iceri Aktar islemini baslatir.
"""

from __future__ import annotations

import sys

from tsoft_sitemap_panel_yukle import panel_csv_yukle, panel_yukleme_aktif_mi
from tsoft_stoklu_urun_sitemap_yillik import (
    VARSAYILAN_LINK_DOSYASI,
    VARSAYILAN_STOK_DOSYASI,
    sitemap_csv_olustur,
)


def main() -> int:
    try:
        print("Stok guncellemesi calistirilmadi; mevcut son stok sonucu kullaniliyor.")
        print(f"Stok kaynagi: {VARSAYILAN_STOK_DOSYASI}")
        print(f"T-Soft link kaynagi: {VARSAYILAN_LINK_DOSYASI}")

        sonuc = sitemap_csv_olustur()
        print()
        print(f"Stoklu benzersiz urun: {sonuc['stoklu_sayisi']}")
        print(f"YEARLY yapilacak urun linki: {sonuc['yillik_sayisi']}")
        print(f"Hazirlanan CSV: {sonuc['cikti']}")
        if sonuc["eslesmeyen_sayisi"]:
            print(f"Eslesmeyen stoklu urun: {sonuc['eslesmeyen_sayisi']}")
            print(f"Eslesmeyenler: {sonuc['eslesmeyen_dosyasi']}")

        if not panel_yukleme_aktif_mi():
            raise RuntimeError("T-Soft panel otomatik yukleme ayarlarda pasif.")

        print("T-Soft panel yuklemesi otomatik baslatiliyor...")
        panel_sonuc = panel_csv_yukle(sonuc["cikti"], onay_iste=False)
        print()
        print(f"T-Soft panel YEARLY yukleme durumu: {panel_sonuc['durum']}")
        if panel_sonuc.get("log"):
            print(f"Panel yukleme logu: {panel_sonuc['log']}")
        return 0 if panel_sonuc["durum"] in {"TAMAM", "IPTAL"} else 2
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"HATA: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
