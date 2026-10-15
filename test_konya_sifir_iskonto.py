from fiyatlandirma_raporu import alis_kaynagi_kullanilabilir_mi


def fiyat(*, iskonto, stok=1):
    return {
        "liste_tl": 100,
        "iskonto_yuzde": iskonto,
        "net_tl_kdv_haric": 100 if iskonto in (None, 0) else 75,
        "stok": stok,
    }


def test_konya_weidmuller_sayisal_sifir_iskonto_stoklu_kaynakta_gecerli():
    assert alis_kaynagi_kullanilabilir_mi(
        fiyat(iskonto=0),
        tedarikci="KONYA",
        marka="WEIDMÜLLER",
    )


def test_konya_weidmuller_sayisal_sifir_iskonto_tek_kaynakta_stoksuz_gecerli():
    assert alis_kaynagi_kullanilabilir_mi(
        fiyat(iskonto=0, stok=0),
        tedarikci="Konya",
        marka="WEIDMULLER",
        tek_kaynak=True,
    )


def test_konya_weidmuller_sayisal_sifir_iskonto_cok_kaynakta_stoksuz_gecersiz():
    assert not alis_kaynagi_kullanilabilir_mi(
        fiyat(iskonto=0, stok=0),
        tedarikci="KONYA",
        marka="WEIDMULLER",
    )


def test_konya_diger_markada_sifir_iskonto_gecersiz():
    assert not alis_kaynagi_kullanilabilir_mi(
        fiyat(iskonto=0),
        tedarikci="KONYA",
        marka="MUTLUSAN",
        tek_kaynak=True,
    )


def test_diger_tedarikcide_sifir_iskonto_gecersiz():
    assert not alis_kaynagi_kullanilabilir_mi(
        fiyat(iskonto=0),
        tedarikci="SANTRAL",
        marka="WEIDMULLER",
        tek_kaynak=True,
    )


def test_konya_bos_iskonto_gecersiz():
    assert not alis_kaynagi_kullanilabilir_mi(
        fiyat(iskonto=None),
        tedarikci="KONYA",
        marka="WEIDMULLER",
        tek_kaynak=True,
    )
