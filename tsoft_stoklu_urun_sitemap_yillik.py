"""Stoklu T-Soft urun linkleri icin YEARLY sitemap CSV'si hazirlar.

T-Soft'un link ice aktariminda Type ve TableId alanlari zorunludur. Bu nedenle
betik, T-Soft CSV Urun Sihirbazi'ndan disari aktarilan link dosyasini stok
guncelleme listesiyle TableId/ProductId uzerinden eslestirir. T-Soft'a kendi
basina veri gondermez; kontrol edilebilir bir ice aktarim CSV'si uretir.
"""

from __future__ import annotations

import argparse
import csv
from datetime import datetime
from pathlib import Path
import re
import sys

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
VARSAYILAN_STOK_DOSYASI = BASE_DIR / "Cikti" / "tum_bayiler_stok_guncellenecekler.xlsx"
VARSAYILAN_LINK_DOSYASI = BASE_DIR / "Tsoft" / "tsoft_linkler.csv"
CIKTI_DIR = BASE_DIR / "Cikti"


def metin_norm(deger: object) -> str:
    metin = str(deger or "").strip().casefold()
    ceviri = str.maketrans({"ı": "i", "ş": "s", "ğ": "g", "ü": "u", "ö": "o", "ç": "c"})
    return re.sub(r"[^a-z0-9]", "", metin.translate(ceviri))


def kimlik_norm(deger: object) -> str:
    """Excel'in 123.0 bicimine cevirdigi kimlikleri guvenle tekillestirir."""
    metin = str(deger or "").strip()
    if re.fullmatch(r"\d+\.0", metin):
        return metin[:-2]
    return metin


def sayiya_cevir(deger: object) -> float:
    metin = str(deger or "").strip().replace(" ", "")
    if not metin or metin.casefold() == "nan":
        return 0.0
    if re.fullmatch(r"[-+]?\d{1,3}(?:\.\d{3})+", metin):
        metin = metin.replace(".", "")
    else:
        metin = metin.replace(",", ".")
    try:
        return float(metin)
    except ValueError:
        return 0.0


def csv_metin_oku(dosya: Path) -> tuple[str, str]:
    ham = dosya.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
        try:
            return ham.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    raise ValueError("Link CSV dosyasinin karakter kodlamasi okunamadi.")


def link_csv_oku(dosya: Path) -> tuple[list[dict[str, str]], str, str]:
    metin, encoding = csv_metin_oku(dosya)
    ornek = metin[:8192]
    try:
        dialect = csv.Sniffer().sniff(ornek, delimiters=",;\t|")
        ayirac = dialect.delimiter
    except csv.Error:
        ayirac = ";" if ornek.count(";") > ornek.count(",") else ","

    okuyucu = csv.DictReader(metin.splitlines(), delimiter=ayirac)
    if not okuyucu.fieldnames:
        raise ValueError("Link CSV dosyasinda baslik satiri bulunamadi.")

    baslik_haritasi = {metin_norm(baslik): baslik for baslik in okuyucu.fieldnames if baslik}
    type_basligi = baslik_haritasi.get("type") or baslik_haritasi.get("tip")
    table_id_basligi = baslik_haritasi.get("tableid") or baslik_haritasi.get("tabloid")
    if not type_basligi or not table_id_basligi:
        bulunan = ", ".join(str(x) for x in okuyucu.fieldnames)
        raise ValueError(
            "T-Soft link disari aktariminda Tip/Type ve Tablo ID/TableId kolonlari zorunludur. "
            f"Bulunan kolonlar: {bulunan}"
        )

    satirlar = []
    for satir in okuyucu:
        satirlar.append(
            {
                "Type": str(satir.get(type_basligi, "")).strip(),
                "TableId": kimlik_norm(satir.get(table_id_basligi, "")),
            }
        )
    return satirlar, encoding, ayirac


def stoklu_urunleri_oku(dosya: Path) -> dict[str, dict[str, object]]:
    df = pd.read_excel(dosya, dtype=str).fillna("")
    gerekli = {"ProductId", "Guncellenecek Stok"}
    eksik = gerekli - set(df.columns)
    if eksik:
        raise ValueError("Stok dosyasinda eksik kolonlar: " + ", ".join(sorted(eksik)))

    sonuc: dict[str, dict[str, object]] = {}
    for _, satir in df.iterrows():
        product_id = kimlik_norm(satir.get("ProductId", ""))
        if not product_id:
            continue
        durum = str(satir.get("Guncelleme Durumu", "")).strip().upper()
        if durum == "GONDERILMEYECEK":
            continue
        stok = sayiya_cevir(satir.get("Guncellenecek Stok", 0))
        if stok <= 0:
            continue
        sonuc[product_id] = {
            "ProductId": product_id,
            "Web Servis Kodu": str(satir.get("Web Servis Kodu", "")).strip(),
            "Stok": stok,
        }
    return sonuc


def yillik_link_satirlari_hazirla(
    stoklular: dict[str, dict[str, object]],
    linkler: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[str]]:
    bulunan_kimlikler: set[str] = set()
    cikti: list[dict[str, str]] = []
    gorulen: set[tuple[str, str]] = set()

    for link in linkler:
        table_id = kimlik_norm(link.get("TableId", ""))
        link_type = str(link.get("Type", "")).strip()
        # TableId degeri kategori, etiket ve diger link turlerinde de ayni
        # sayisal degeri kullanabilir. Yalniz urun linklerini guncelle.
        if link_type.casefold() != "product":
            continue
        if table_id not in stoklular:
            continue
        anahtar = (link_type.casefold(), table_id)
        if anahtar in gorulen:
            continue
        gorulen.add(anahtar)
        bulunan_kimlikler.add(table_id)
        cikti.append(
            {
                "Type": link_type,
                "TableId": table_id,
                "SiteMapChangeFreq": "YEARLY",
            }
        )

    eslesmeyen = sorted(set(stoklular) - bulunan_kimlikler, key=lambda x: (len(x), x))
    return cikti, eslesmeyen


def csv_yaz(satirlar: list[dict[str, str]], dosya: Path) -> None:
    dosya.parent.mkdir(parents=True, exist_ok=True)
    with dosya.open("w", encoding="utf-8-sig", newline="") as handle:
        yazici = csv.DictWriter(
            handle,
            fieldnames=["Type", "TableId", "SiteMapChangeFreq"],
            delimiter=";",
        )
        yazici.writeheader()
        yazici.writerows(satirlar)


def eslesmeyenleri_yaz(
    eslesmeyen: list[str],
    stoklular: dict[str, dict[str, object]],
    dosya: Path,
) -> None:
    satirlar = [stoklular[product_id] for product_id in eslesmeyen]
    pd.DataFrame(satirlar, columns=["ProductId", "Web Servis Kodu", "Stok"]).to_csv(
        dosya,
        sep=";",
        index=False,
        encoding="utf-8-sig",
    )


def sitemap_csv_olustur(
    stok_dosyasi: Path = VARSAYILAN_STOK_DOSYASI,
    link_dosyasi: Path = VARSAYILAN_LINK_DOSYASI,
    cikti: Path | None = None,
) -> dict[str, object]:
    """YEARLY ice aktarim CSV'sini olusturur ve islem ozetini dondurur."""
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
    satirlar, eslesmeyen = yillik_link_satirlari_hazirla(stoklular, linkler)
    if not satirlar:
        raise ValueError("Stoklu urunlerle eslesen hicbir T-Soft urun linki bulunamadi.")

    zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
    cikti = Path(cikti).resolve() if cikti else (
        CIKTI_DIR / f"tsoft_stoklu_urun_sitemap_yillik_{zaman}.csv"
    )
    csv_yaz(satirlar, cikti)

    eslesmeyen_dosyasi = None
    if eslesmeyen:
        eslesmeyen_dosyasi = cikti.with_name(cikti.stem + "_eslesmeyen.csv")
        eslesmeyenleri_yaz(eslesmeyen, stoklular, eslesmeyen_dosyasi)

    return {
        "stoklu_sayisi": len(stoklular),
        "link_sayisi": len(linkler),
        "yillik_sayisi": len(satirlar),
        "eslesmeyen_sayisi": len(eslesmeyen),
        "encoding": encoding,
        "ayirac": ayirac,
        "cikti": cikti,
        "eslesmeyen_dosyasi": eslesmeyen_dosyasi,
    }


def argumanlar() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Stoklu urun linklerini YEARLY yapmak icin T-Soft ice aktarim CSV'si olusturur."
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
        print(
            f"Link disari aktarim satiri: {sonuc['link_sayisi']} "
            f"({sonuc['encoding']}, ayirac={sonuc['ayirac']!r})"
        )
        print(f"YEARLY yapilacak urun linki: {sonuc['yillik_sayisi']}")
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
