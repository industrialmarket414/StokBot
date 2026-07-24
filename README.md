# EntesStokBot

Tedarikçi B2B dosyalarını T-Soft ürünleriyle eşleştiren, stok ve fiyatlandırma
raporları oluşturan Python tabanlı otomasyon projesidir.

## Özellikler

- T-Soft ürünlerinin Excel'e aktarılması
- B2B ürünlerinin web servis kodu ve marka kurallarıyla eşleştirilmesi
- Tedarikçi liste fiyatı ve iskontosundan net alış fiyatı hesaplanması
- Minimum ve hedef kâr senaryolarının Excel raporunda oluşturulması
- Onaylanan iskontoların T-Soft ziyaretçi grubuna aktarılması
- Tedarikçi bazlı stok güncelleme dosyalarının hazırlanması ve aktarılması

## Kurulum

Python 3.11 veya daha yeni bir sürüm önerilir.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python kurulum.py
```

`Ayarlar/tsoft_api.example.json` dosyasını `Ayarlar/tsoft_api.json` adıyla
kopyalayın ve kendi T-Soft bilgilerinizle doldurun. Gerçek ayar dosyası Git
tarafından özellikle izlenmez.

Tedarikçi Excel'leri, T-Soft dışa aktarımları, raporlar ve API logları da ticari
veri içerebildiği için depoya dahil edilmez.

## Temel fiyatlandırma akışı

1. `python tsoft_products_export.py`
2. Güncel B2B Excel dosyalarını ilgili klasörlere yerleştirme
3. `python stok_eslestir.py`
4. `python fiyatlandirma_raporu.py`
5. `Cikti/fiyat_onay_listesi.xlsx` içinde kâr/iskonto seçimi ve `ONAY`
6. Kontrol: `python tsoft_fiyat_aktar.py`
7. Gerçek aktarım: `python tsoft_fiyat_aktar.py --gonder`

`--gonder` verilmediğinde fiyat aktarım aracı kontrol modunda çalışır ve
T-Soft'ta değişiklik yapmaz.

## Güvenlik

- `Ayarlar/tsoft_api.json` dosyasını commit etmeyin.
- `Log/` altındaki oturum ve API cevaplarını paylaşmayın.
- B2B fiyat listeleri ile `Cikti/` raporlarını herkese açık depoya yüklemeyin.

