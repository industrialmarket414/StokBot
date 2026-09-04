"""31.08.2026 kaynak denetimi: firma bazli paket kurallari ve inceleme bekleyenler.

Kaynaklari dogrulanmayan fiyatlar otomatik degistirilmez. Bu modul ag istegi
yapmaz; eski rapor/JSON ile hatali aktarimi da engeller.
"""
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

SURUM = "2026-09-02-paket-denetimi-v5"

def norm(value):
    text = str(value or "").strip().upper().translate(str.maketrans("ÇĞİÖŞÜ", "CGIOSU"))
    text = ''.join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))
    return re.sub(r'[^A-Z0-9]+', ' ', text).strip()

# code, supplier, source kind, covered units, brand, category
# Elektrik Malzemesi kaynak adlarinda acikca 100 ADET yazan SS urunleri.
SS_KODLARI = (
    'SS-100-2C', 'SS-100B-2C', 'SS-135-2C', 'SS-140-3C', 'SS-140-4C',
    'SS-160-2C', 'SS-1919B', 'SS-200-2C', 'SS-200-3C', 'SS-200-4C',
    'SS-200B-3C', 'SS-250-3C', 'SS-280-3C', 'SS-280-7C', 'SS-360-3C',
    'SS-360-4C', 'SS-360B-3C', 'SS-540B-7C', 'SS-HRCV-200',
    'SS-MV200-4C', 'SS-AM-1', 'SS-TM-2', 'SS-TM-2B',
)
PAKET_KURALLARI = [
    ('505000','SANTRAL','alis',500,'KLEMSAN','Etiketler'),
    ('505000','ELEKTRİK MALZEMESİ','rakip',500,'KLEMSAN','Etiketler'),
    ('517006','ELEKTRİK MALZEMESİ','rakip',500,'KLEMSAN','Kablo İşaretleme Etiketleri'),
    ('519005','ELEKTRİK MALZEMESİ','rakip',500,'KLEMSAN','Kablo İşaretleme Etiketleri'),
    ('001 054 160014 00 11','ELEKTRİK MALZEMESİ','rakip',50,'MUTLUSAN','Spiraller'),
    ('UY-23025','ELEKTRİK MALZEMESİ','rakip',250,'ŞAFAK ELEKTRİK','Kablo Yüksükleri'),
    # GGX20 rakip fiyatı 30 adetlik minimum sipariş toplamıdır.
    # T-Soft liste fiyatı adet bazlı olduğu için rakip toplamı 30'a bölünür.
    ('GGX20','ELEKTRİK MALZEMESİ','rakip',30,'ORTAÇ','Spiraller'),
    ('SS-135-2C','ELEKTROMARKETİM','rakip',100,'ORTAÇ','Kablo Bağları ve Sabitleme Klipsleri'),
    ('SS-200-4C','ELEKTRIX','rakip',100,'ŞAFAK ELEKTRİK','Kablo Bağları ve Sabitleme Klipsleri'),
    ('SS-TM-2','ELEKTRIX','rakip',100,'ORTAÇ','Kablo Bağları ve Sabitleme Klipsleri'),
    ('SS-TM-2B','ELEKTRIX','rakip',100,'ORTAÇ','Kablo Bağları ve Sabitleme Klipsleri'),
] + [(code,'ELEKTRİK MALZEMESİ','rakip',100,
       'ŞAFAK ELEKTRİK' if code == 'SS-200-4C' else 'ORTAÇ',
       'Kablo Bağları ve Sabitleme Klipsleri') for code in SS_KODLARI]

# KONYA 505000: kaynakta liste 0,05 EUR / adet; tekrar 500'e bolunemez.
IPTAL_KURALLAR = {(norm('505000'),norm('KONYA'),'ALIS')}
DUZELTILEN_KODLAR = {norm(rule[0]) for rule in PAKET_KURALLARI}

# MET dosyasında NOAS ürünlerinin Min. Paket Adedi 1 olarak doğrulandı.
# Bu ürünler normal fiyat hesabına girer; zam gerektirenler raporda ayrıca işaretlenir.
BEKLEYEN_KONTROLLER = {}

def kaynak_kontrol_nedeni(code):
    return BEKLEYEN_KONTROLLER.get(norm(code), '')

def dogrulanmis_paketleri_uygula(result):
    """Onayli kurallari haritaya ve rapor sayfasina birlikte yazar; mukerrer tutmaz."""
    rule_keys = {(norm(c),norm(f),norm(t)) for c,f,t,*_ in PAKET_KURALLARI}
    replaced = rule_keys | IPTAL_KURALLAR
    result['satirlar'] = [r for r in result['satirlar'] if
        (norm(r.get('Ürün Web Servis Kodu')),norm(r.get('Firma')),norm(r.get('Kaynak Türü'))) not in replaced]
    for code,firm,kind in IPTAL_KURALLAR:
        result['alis'].pop((code,firm),None)
        result['alis_marka'].pop((code,firm),None)
    for code,firm,kind,units,brand,category in PAKET_KURALLARI:
        key=(norm(code),norm(firm)); units=Decimal(units)
        result[kind][key]=units
        if kind=='alis': result['alis_marka'][key]=norm(brand)
        result['satirlar'].append({'Aktif':'EVET','ProductId':'','Marka':brand,'Kategori':category,
            'Ürün Web Servis Kodu':code,'Firma':firm,'Kaynak Türü':'ALIŞ' if kind=='alis' else 'RAKİP',
            'Paket Adedi':units,'Fiyatın Kapsadığı Adet':units,'Kontrol Durumu':'ONAYLANDI'})
    return result

def hazir_duzeltme_kodlari(records):
    actual={}
    for r in records:
        key=(norm(r.get('Ürün Web Servis Kodu')),norm(r.get('Firma')),norm(r.get('Kaynak Türü')))
        if norm(r.get('Aktif'))=='EVET' and norm(r.get('Kontrol Durumu'))=='ONAYLANDI':
            try: actual[key]=Decimal(str(r.get('Fiyatın Kapsadığı Adet')))
            except Exception: pass
    ready=set(DUZELTILEN_KODLAR)
    for code,firm,kind,units,*_ in PAKET_KURALLARI:
        if actual.get((norm(code),norm(firm),norm(kind))) != units: ready.discard(norm(code))
    for key in IPTAL_KURALLAR:
        if key in actual: ready.discard(key[0])
    return ready

@lru_cache(maxsize=8)
def _rapor_hazir_kodlar(path,mtime_ns,size):
    # Sadece kucuk paket sayfasi okunur; pickle ya da formuller calistirilmaz.
    ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    rel='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    with zipfile.ZipFile(path) as z:
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        sheet=next(s for s in wb.find(ns+'sheets') if s.attrib['name']=='Paket Birim Ayarları')
        links=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target=next(r.attrib['Target'] for r in links if r.attrib['Id']==sheet.attrib[rel]).lstrip('/')
        if not target.startswith('xl/'): target='xl/'+target
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=[''.join(x.itertext()) for x in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        xml=ET.fromstring(z.read(target)); headers={}; records=[]
        for row in xml.iter(ns+'row'):
            values={}
            for cell in row:
                col=re.sub(r'\d','',cell.attrib.get('r',''))
                if cell.attrib.get('t')=='inlineStr': val=''.join(t.text or '' for t in cell.iter(ns+'t'))
                else:
                    val=cell.findtext(ns+'v','')
                    if cell.attrib.get('t')=='s' and val: val=strings[int(val)]
                values[col]=val
            if 'Ürün Web Servis Kodu' in values.values(): headers=values
            elif headers: records.append({name:values.get(col,'') for col,name in headers.items()})
        return frozenset(hazir_duzeltme_kodlari(records))

def rapor_hazir_kodlar(path):
    try:
        path=Path(path); stat=path.stat()
        return _rapor_hazir_kodlar(str(path.resolve()),stat.st_mtime_ns,stat.st_size)
    except Exception:
        return frozenset()  # Eksik/bozuk rapor duzeltilmis varsayilmaz.

def aktarim_engeli(code,report_path):
    reason=kaynak_kontrol_nedeni(code)
    if reason: return reason
    if norm(code) in DUZELTILEN_KODLAR and norm(code) not in rapor_hazir_kodlar(report_path):
        return 'Paket/birim kuralları düzeltildi; Fiyat Raporunu Oluştur ile raporu yenileyin.'
    return ''

def aktarim_dogrula(code,report_path):
    reason=aktarim_engeli(code,report_path)
    if reason: raise ValueError('KAYNAK FİYAT KONTROLÜ: '+reason)
