"""T-Soft link ice aktarim CSV'sini Chrome ile panele yukler.

Ilk kurulumda kullanici panelde oturum acar ve "Linkleri Iceri Aktar" sayfasina
bir kez gider. Sayfa adresi ve ayri bir Chrome profilindeki oturum sonraki
calistirmalar icin saklanir. Yukleme her zaman konsolda YUKLE onayi ister.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import json
from pathlib import Path
import re
import sys
import time
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
AYARLAR_DIR = BASE_DIR / "Ayarlar"
LOG_DIR = BASE_DIR / "Log"
API_CONFIG = AYARLAR_DIR / "tsoft_api.json"
PANEL_CONFIG = AYARLAR_DIR / "tsoft_panel_yukleme.json"
VARSAYILAN_CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
VARSAYILAN_PROFIL = AYARLAR_DIR / "tsoft_panel_chrome_profili"


def playwright_modulu():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Playwright kurulu degil. Proje sanal ortaminda 'python -m pip install playwright' calistirin."
        ) from exc
    return sync_playwright


def site_url_oku() -> str:
    if not API_CONFIG.is_file():
        return ""
    data = json.loads(API_CONFIG.read_text(encoding="utf-8"))
    return str(data.get("site_url", "")).strip().rstrip("/")


def varsayilan_ayarlar() -> dict[str, Any]:
    site_url = site_url_oku()
    return {
        "aktif": True,
        "panel_url": f"{site_url}/admin" if site_url else "",
        "link_ice_aktar_url": "",
        "chrome_executable": str(VARSAYILAN_CHROME),
        "profil_klasoru": str(VARSAYILAN_PROFIL),
        "yukleme_zaman_asimi_ms": 180000,
    }


def ayarlari_oku() -> dict[str, Any]:
    AYARLAR_DIR.mkdir(exist_ok=True)
    varsayilan = varsayilan_ayarlar()
    if not PANEL_CONFIG.is_file():
        PANEL_CONFIG.write_text(
            json.dumps(varsayilan, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return varsayilan
    mevcut = json.loads(PANEL_CONFIG.read_text(encoding="utf-8"))
    varsayilan.update(mevcut if isinstance(mevcut, dict) else {})
    return varsayilan


def ayarlari_yaz(ayarlar: dict[str, Any]) -> None:
    AYARLAR_DIR.mkdir(exist_ok=True)
    PANEL_CONFIG.write_text(
        json.dumps(ayarlar, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def panel_yukleme_aktif_mi() -> bool:
    return bool(ayarlari_oku().get("aktif", False))


def csv_satir_sayisi(csv_dosyasi: Path) -> int:
    with Path(csv_dosyasi).open("r", encoding="utf-8-sig") as handle:
        return max(sum(1 for _ in handle) - 1, 0)


def giris_ekrani_mi(page) -> bool:
    try:
        return page.locator("input[type='password']").count() > 0
    except Exception:
        return False


def ice_aktarim_yuzeyini_bul(context):
    """Dosya alanini yeni sekme veya iframe dahil tum tarayici yuzeylerinde bulur."""
    for aday_page in reversed(context.pages):
        for yuzey in reversed(aday_page.frames):
            try:
                if yuzey.locator("input[type='file']").count() < 1:
                    continue
                metin = yuzey.locator("body").inner_text(timeout=5000).casefold()
                adres = f"{aday_page.url} {yuzey.url}".casefold()
                if (
                    ("link" in metin and ("içe aktar" in metin or "ice aktar" in metin))
                    or "links-import" in adres
                ):
                    return aday_page, yuzey
            except Exception:
                continue
    return None, None


def ice_aktarim_ekrani_mi(context) -> bool:
    _, yuzey = ice_aktarim_yuzeyini_bul(context)
    return yuzey is not None


def link_ice_aktarim_sayfasina_git(context, ana_page):
    """CSV Sihirbazi iframe'inde Genel Islemler > Linkleri Iceri Aktar'i acar."""
    for _ in range(60):
        hedef_page, yuzey = ice_aktarim_yuzeyini_bul(context)
        if yuzey is not None:
            return hedef_page, yuzey

        for aday_page in reversed(context.pages):
            for frame in reversed(aday_page.frames):
                if "app.csvexport" not in frame.url.casefold():
                    continue
                try:
                    genel = frame.get_by_text(re.compile(r"^Genel\s+İşlemler$", re.I))
                    if genel.count() > 0:
                        genel.first.click(timeout=3000)
                        frame.wait_for_timeout(500)
                    ice_aktar = frame.get_by_text(
                        re.compile(r"^Linkleri\s+İçeri\s+Aktar$", re.I)
                    )
                    if ice_aktar.count() > 0:
                        ice_aktar.first.click(timeout=3000)
                        frame.wait_for_timeout(1000)
                except Exception:
                    continue
        ana_page.wait_for_timeout(1000)
    return None, None


def ice_aktarim_ekranini_bekle(context, ana_page, yenile: bool = True):
    """T-Soft'un gec yuklenen CSV ekranini bekler; gerekirse bir kez yeniler."""
    for deneme in range(2 if yenile else 1):
        for _ in range(30):
            hedef_page, yuzey = ice_aktarim_yuzeyini_bul(context)
            if yuzey is not None:
                return hedef_page, yuzey
            ana_page.wait_for_timeout(1000)
        if deneme == 0 and yenile:
            print("CSV ekrani gecikti; sayfa bir kez yenileniyor...")
            ana_page.reload(wait_until="domcontentloaded", timeout=120000)
    return None, None


def tanilama_kaydet(context, not_metni: str) -> Path:
    """Bos/bozuk panel ekraninda URL, iframe ve kisa govde metnini kaydeder."""
    LOG_DIR.mkdir(exist_ok=True)
    zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
    kayit = {"zaman": zaman, "not": not_metni, "sayfalar": []}
    for sira, aday_page in enumerate(context.pages, start=1):
        sayfa_bilgisi = {"url": aday_page.url, "baslik": "", "frameler": []}
        try:
            sayfa_bilgisi["baslik"] = aday_page.title(timeout=5000)
        except Exception:
            pass
        for frame in aday_page.frames:
            frame_bilgisi = {"url": frame.url, "metin": ""}
            try:
                frame_bilgisi["metin"] = frame.locator("body").inner_text(timeout=5000)[-4000:]
            except Exception as exc:
                frame_bilgisi["hata"] = str(exc)
            sayfa_bilgisi["frameler"].append(frame_bilgisi)
        kayit["sayfalar"].append(sayfa_bilgisi)
        try:
            aday_page.screenshot(
                path=str(LOG_DIR / f"tsoft_kurulum_hata_{zaman}_{sira}.png"),
                full_page=True,
            )
        except Exception:
            pass
    log_yolu = LOG_DIR / f"tsoft_kurulum_hata_{zaman}.json"
    log_yolu.write_text(json.dumps(kayit, ensure_ascii=False, indent=2), encoding="utf-8")
    return log_yolu


def ilk_kurulum(page, context, ayarlar: dict[str, Any]):
    panel_url = str(ayarlar.get("panel_url", "")).strip()
    if not panel_url:
        raise ValueError(f"{PANEL_CONFIG} icinde panel_url bos.")

    page.goto(panel_url, wait_until="domcontentloaded", timeout=120000)
    print()
    print("ILK T-SOFT PANEL KURULUMU")
    if giris_ekrani_mi(page):
        print("Acilan Chrome penceresinde T-Soft paneline giris yapin.")
        input("Giris tamamlaninca ENTER'a basin: ")
        page.goto(panel_url, wait_until="domcontentloaded", timeout=120000)
    print("Linkleri Iceri Aktar ekrani otomatik aciliyor...")
    hedef_page, yuzey = link_ice_aktarim_sayfasina_git(context, page)
    if yuzey is None:
        log_yolu = tanilama_kaydet(context, "Linkleri Iceri Aktar formu yuklenmedi")
        raise RuntimeError(
            "Linkleri Iceri Aktar formu yuklenmedi. "
            f"Tani kaydi: {log_yolu}"
        )

    ayarlar["link_ice_aktar_url"] = hedef_page.url
    ayarlari_yaz(ayarlar)
    print(f"Link ice aktarim sayfasi kaydedildi: {hedef_page.url}")
    return hedef_page, yuzey


def noktalivirgulu_sec(page) -> None:
    """Ayirac kontrolu varsa noktali virgulu secer; yoksa varsayilana dokunmaz."""
    secenekler = (
        page.get_by_text(re.compile(r"noktal[ıi]\s*virg[üu]l", re.I)),
        page.get_by_label(re.compile(r"ay[ıi]ra[cç]", re.I)),
    )
    for locator in secenekler:
        try:
            if locator.count() < 1:
                continue
            ilk = locator.first
            if ilk.evaluate("el => el.tagName === 'SELECT'"):
                try:
                    ilk.select_option(label=re.compile(r"noktal[ıi]\s*virg[üu]l", re.I))
                except Exception:
                    ilk.select_option(value=";")
            else:
                ilk.click(timeout=3000)
            return
        except Exception:
            continue


def yukle_butonunu_bul(page):
    adaylar = [
        page.get_by_role("button", name=re.compile(r"^y[üu]kle$", re.I)),
        page.locator("button:has-text('Yükle'), button:has-text('Yukle')"),
        page.locator("input[type='submit']"),
    ]
    for aday in adaylar:
        try:
            if aday.count() > 0 and aday.first.is_visible():
                return aday.first
        except Exception:
            continue
    raise RuntimeError("T-Soft Linkleri Iceri Aktar ekraninda Yukle butonu bulunamadi.")


def sonuc_metnini_oku(page) -> str:
    try:
        return page.locator("body").inner_text(timeout=10000)
    except Exception:
        return ""


def aktarimin_bitmesini_bekle(yuzey, page, timeout_ms: int) -> tuple[bool, int, str]:
    """T-Soft'un Satir Isleniyor penceresi bitene kadar ilerlemeyi izler."""
    # Buyuk dosyalarda panelin satir islemesi uc dakikadan uzun surebilir.
    son_zaman = time.monotonic() + max(timeout_ms, 1_800_000) / 1000
    islem_goruldu = False
    en_yuksek_yuzde = 0
    son_yazilan_yuzde = -1
    son_metin = ""

    while time.monotonic() < son_zaman:
        son_metin = sonuc_metnini_oku(yuzey)
        # Python casefold, Turkce buyuk I-noktali harfini "i + birlesen nokta"
        # yapar. Birlesen noktayi temizle ki "Satir Isleniyor" algilansin.
        norm = son_metin.casefold().replace("\u0307", "")
        isleniyor = bool(
            re.search(r"sat[ıi]r\s+i[sş]leniyor|i[sş]leniyor\.\.\.", norm)
        )
        yuzdeler = [
            int(eslesme.group(1) or eslesme.group(2))
            for eslesme in re.finditer(r"%\s*(\d{1,3})|(\d{1,3})\s*%", son_metin)
        ]
        if yuzdeler:
            en_yuksek_yuzde = max(en_yuksek_yuzde, max(yuzdeler))
            if en_yuksek_yuzde != son_yazilan_yuzde:
                print(f"T-Soft satir isleme ilerlemesi: %{en_yuksek_yuzde}")
                son_yazilan_yuzde = en_yuksek_yuzde

        if isleniyor:
            islem_goruldu = True
        elif islem_goruldu:
            page.wait_for_timeout(2000)
            son_metin = sonuc_metnini_oku(yuzey)
            return True, en_yuksek_yuzde, son_metin
        elif time.monotonic() > son_zaman - max(timeout_ms, 1_800_000) / 1000 + 15:
            return False, en_yuksek_yuzde, son_metin

        page.wait_for_timeout(2000)

    return False, en_yuksek_yuzde, son_metin


def panel_csv_yukle(csv_dosyasi: Path, onay_iste: bool = True) -> dict[str, Any]:
    csv_dosyasi = Path(csv_dosyasi).resolve()
    if not csv_dosyasi.is_file():
        raise FileNotFoundError(f"Yuklenecek CSV bulunamadi: {csv_dosyasi}")
    urun_sayisi = csv_satir_sayisi(csv_dosyasi)
    if urun_sayisi < 1:
        raise ValueError("Yuklenecek CSV'de urun satiri yok.")

    ayarlar = ayarlari_oku()
    if not ayarlar.get("aktif", False):
        raise RuntimeError(f"Panel otomatik yukleme pasif: {PANEL_CONFIG}")
    chrome = Path(str(ayarlar.get("chrome_executable", VARSAYILAN_CHROME)))
    if not chrome.is_file():
        raise FileNotFoundError(f"Chrome bulunamadi: {chrome}")
    profil = Path(str(ayarlar.get("profil_klasoru", VARSAYILAN_PROFIL)))
    profil.mkdir(parents=True, exist_ok=True)
    timeout_ms = int(ayarlar.get("yukleme_zaman_asimi_ms", 180000))

    sync_playwright = playwright_modulu()
    zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
    LOG_DIR.mkdir(exist_ok=True)
    ekran_dosyasi = LOG_DIR / f"tsoft_sitemap_yukleme_{zaman}.png"
    log_dosyasi = LOG_DIR / f"tsoft_sitemap_yukleme_{zaman}.json"

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profil),
            executable_path=str(chrome),
            ignore_default_args=["--no-sandbox"],
            headless=False,
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(30000)
        hedef_url = str(ayarlar.get("link_ice_aktar_url", "")).strip()
        if not hedef_url:
            page, yuzey = ilk_kurulum(page, context, ayarlar)
        else:
            page.goto(hedef_url, wait_until="domcontentloaded", timeout=120000)
            if giris_ekrani_mi(page):
                print("T-Soft panel oturumu sona ermis. Acilan Chrome'da tekrar giris yapin.")
                input("Giris yapip Linkleri Iceri Aktar sayfasina geldikten sonra ENTER'a basin: ")
            page, yuzey = link_ice_aktarim_sayfasina_git(context, page)
            if yuzey is None:
                context.close()
                raise RuntimeError("T-Soft Linkleri Iceri Aktar ekrani dogrulanamadi.")
            if page.url != hedef_url:
                ayarlar["link_ice_aktar_url"] = page.url
                ayarlari_yaz(ayarlar)

        noktalivirgulu_sec(yuzey)
        dosya_alani = yuzey.locator("input[type='file']").first
        dosya_alani.set_input_files(str(csv_dosyasi))

        if onay_iste:
            print()
            print(f"T-Soft'a yuklenecek YEARLY urun linki: {urun_sayisi}")
            print(f"Dosya: {csv_dosyasi}")
            onay = input("Panelde yuklemeyi baslatmak icin YUKLE yazin: ").strip().upper()
            if onay != "YUKLE":
                context.close()
                return {
                    "durum": "IPTAL",
                    "urun_sayisi": urun_sayisi,
                    "csv": str(csv_dosyasi),
                    "log": "",
                }

        yukle_butonunu_bul(yuzey).click()
        tamamlandi, son_yuzde, metin = aktarimin_bitmesini_bekle(
            yuzey, page, timeout_ms
        )
        page.screenshot(path=str(ekran_dosyasi), full_page=True)
        metin_norm = metin.casefold()
        hata_sayisi = re.search(
            r"hatal[ıi]\s+aktar[ıi]mlar?\s*[:(]?\s*(\d+)", metin_norm
        )
        hata_var = bool(hata_sayisi and int(hata_sayisi.group(1)) > 0)
        # T-Soft, tum satirlar bittigi halde son ilerleme degerini bazen %99'da
        # birakip islem penceresini kapatiyor. Pencerenin kapanmasi ve hata
        # sayisinin sifir olmasi ile birlikte %99 tamamlanmis kabul edilir.
        durum = "TAMAM" if tamamlandi and son_yuzde >= 99 and not hata_var else "KONTROL"

        log = {
            "zaman": datetime.now().isoformat(timespec="seconds"),
            "durum": durum,
            "urun_sayisi": urun_sayisi,
            "islem_tamamlandi": tamamlandi,
            "son_yuzde": son_yuzde,
            "csv": str(csv_dosyasi),
            "sayfa": page.url,
            "ekran_goruntusu": str(ekran_dosyasi),
            "sonuc_metni": metin[-12000:],
        }
        log_dosyasi.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
        context.close()
        log["log"] = str(log_dosyasi)
        return log


def argumanlar() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="T-Soft YEARLY link CSV'sini panele yukler.")
    parser.add_argument("csv", nargs="?", type=Path)
    parser.add_argument("--kurulum", action="store_true", help="Panel oturumunu ve link aktarim sayfasini kaydeder")
    parser.add_argument("--evet", action="store_true", help="YUKLE onay sorusunu atlar")
    return parser.parse_args()


def main() -> int:
    args = argumanlar()
    try:
        if args.kurulum:
            ayarlar = ayarlari_oku()
            sync_playwright = playwright_modulu()
            chrome = Path(str(ayarlar.get("chrome_executable", VARSAYILAN_CHROME)))
            profil = Path(str(ayarlar.get("profil_klasoru", VARSAYILAN_PROFIL)))
            with sync_playwright() as p:
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(profil),
                    executable_path=str(chrome),
                    ignore_default_args=["--no-sandbox"],
                    headless=False,
                    viewport={"width": 1440, "height": 900},
                )
                page = context.pages[0] if context.pages else context.new_page()
                ilk_kurulum(page, context, ayarlar)
                context.close()
            print("T-Soft panel yukleme kurulumu tamamlandi.")
            return 0
        if not args.csv:
            raise ValueError("Yuklenecek CSV yolu gerekli veya --kurulum kullanin.")
        sonuc = panel_csv_yukle(args.csv, onay_iste=not args.evet)
        print(f"Panel yukleme durumu: {sonuc['durum']}")
        if sonuc.get("log"):
            print(f"Panel yukleme logu: {sonuc['log']}")
        return 0 if sonuc["durum"] in {"TAMAM", "IPTAL"} else 2
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError) as exc:
        print(f"HATA: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
