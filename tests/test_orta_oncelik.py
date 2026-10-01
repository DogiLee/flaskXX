"""Denetimdeki orta önemli veri kaybı ve izlenebilirlik bulgularının regresyonları (27.09.2026).

B5: Excel'den DİZGİDE gelmiş kartta operatörün girdiği adet "uygulamada ilerleme" sayılmalı.
B6: Talep NO yazım hatası düzeltilince ayrılmış kart geçmişiyle geri bağlanmalı.
N2: Operatör "Üretilen adet" alanı dolu açılmamalı.
B4+N1: İşlemi yapan kişinin adı karta ve işlem loguna yazılmalı; hesap adı kişi adı sayılmamalı.
B8: Admin düzenlemesi değişen alanların eski/yeni değerini ve silinen not satırlarını loglamalı.
B7: Aynı Talep NO + Kart Stok No ile manuel kart ancak açık onayla eklenebilmeli.
"""
import json
import shutil
import unittest
from unittest.mock import patch

import depo
import excel_araclari as ex
from tests import test_excel_sync as fixtures, test_ui as ui
from tests.test_excel_sync import ROOT, row

HEADERS = {'X-CSRF-Token': 'test-token'}


class _Ortak(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    cards = fixtures.SyncTests.cards
    load_app = ui.UITests.load_app

    def upload(self, sheets, **kw):
        path = self.book(sheets)

        def snapshot(_):
            hedef = self.root / 'snapshot.xlsx'; shutil.copy2(path, hedef); return str(hedef)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            return ex.excelden_aktar(str(path), 'tester', **kw)

    def kararlar(self, sonuc):
        return {d['id']: d['gerileme'] for d in sonuc['degisiklikler'] if d['gerileme']}

    def loglar(self, islem):
        return [l for l in depo._loglar if l['islem'] == islem]


class KararOnerisiTests(_Ortak):
    def test_excelden_dizgide_karta_operatorun_girdigi_adet_korunmasi_onerilir(self):
        self.upload({'ELDE DİZGİ': [row(1, qty=6, status='DİZGİDE', end='2099-12-31')],
                     'MAKİNE': []})
        kid = depo._kartlar[0]['id']
        depo.kart_bitir(kid, 2, 'elle1', 'operator', 'elle_dizgi', isim='Ayşe Kaya')
        karar = self.kararlar(self.upload({'ELDE DİZGİ': [row(1, qty=6, status='DİZGİ İÇİN BEKLİYOR',
                                                             end='2099-12-31')], 'MAKİNE': []},
                                          onizleme=True))[kid]
        self.assertTrue(karar['uygulamada'])
        self.assertEqual(karar['onerilen'], 'DİZGİDE')
        # Öneri kabul edilince operatörün girdiği adet korunur.
        self.upload({'ELDE DİZGİ': [row(1, qty=6, status='DİZGİ İÇİN BEKLİYOR', end='2099-12-31')],
                     'MAKİNE': []}, gerileme_secimleri={kid: karar['onerilen']})
        kart = depo.kart_getir(kid)
        self.assertEqual((kart['durum'], kart['tamamlanan_adet']), ('DİZGİDE', 2))

    def test_eski_excel_teslim_kaydindan_kalan_adet_uygulama_ilerlemesi_sayilmaz(self):
        self.upload({'MAKİNE': [row(1, status='TESLİM EDİLDİ', actual='2026-09-18')]})
        kid = depo._kartlar[0]['id']
        self.upload({'MAKİNE': [row(1, status='DİZGİDE')]}, gerileme_secimleri={kid: 'DİZGİDE'})
        self.assertEqual((depo.kart_getir(kid)['tamamlanan_adet'], depo.kart_getir(kid)['operator']), (5, 'Excel'))
        karar = self.kararlar(self.upload({'MAKİNE': [row(1, status='PLANA ALINDI')]}, onizleme=True))[kid]
        self.assertFalse(karar['uygulamada'])
        self.assertEqual(karar['onerilen'], 'PLANA ALINDI')


class TalepGeriBaglamaTests(_Ortak):
    def satirlar(self, talep):
        return {'MAKİNE': [row(1, status='DİZGİDE', end='2099-12-31', request=talep)]}

    def test_talep_yazim_hatasi_duzeltilince_kart_gecmisiyle_geri_baglanir(self):
        self.upload({'MAKİNE': [row(1, end='2099-12-31', request='T1')]})
        asil = depo._kartlar[0]['id']
        depo.kart_baslat(asil, 5, 'makine1', 'operator', 'makine', isim='Ali Veli')
        depo.kart_bitir(asil, 2, 'makine1', 'operator', 'makine', aciklama='önemli not', isim='Ali Veli')

        self.assertEqual(self.upload(self.satirlar('T9'))['ayrilan'], 1)          # yazım hatası
        hatali = next(k['id'] for k in depo._kartlar if k['talep_no'] == 'T9')

        onizleme = self.upload(self.satirlar('T1'), onizleme=True)                  # düzeltildi
        self.assertEqual(onizleme['geri_baglanan'], 1)
        self.assertEqual(onizleme['yeni'], 0)
        geri = next(d for d in onizleme['degisiklikler'] if d['id'] == asil)
        self.assertEqual((geri['tur'], geri['geri_baglandi']), ('geri', True))

        sonuc = self.upload(self.satirlar('T1'))
        self.assertEqual((sonuc['yeni'], sonuc['geri_baglanan'], sonuc['ayrilan']), (0, 1, 1))
        aktif = depo.kartlari_getir()
        self.assertEqual([k['id'] for k in aktif], [asil])
        self.assertEqual((aktif[0]['talep_no'], aktif[0]['tamamlanan_adet'], aktif[0]['source_row_id']),
                         ('T1', 2, 'NO:1'))
        self.assertIn('önemli not', aktif[0]['aciklama'])
        self.assertEqual(depo.kart_getir(hatali)['source_active'], 0)
        self.assertEqual(len(depo._kartlar), 2)                                   # üçüncü kart açılmadı
        self.assertEqual(len(self.loglar('EXCEL NO ESKİ TALEBE DÖNDÜ')), 1)
        # Aynı dosya tekrar: kart yerinde kalır, yeni ayrılma olmaz.
        tekrar = self.upload(self.satirlar('T1'))
        self.assertEqual((tekrar['yeni'], tekrar['ayrilan'], tekrar['geri_baglanan'], tekrar['degismeyen']),
                         (0, 0, 0, 1))
        depo.kur()
        self.assertEqual([k['id'] for k in depo.kartlari_getir()], [asil])

    def test_birden_fazla_aday_varsa_tahmin_yapilmaz(self):
        self.upload({'MAKİNE': [row(1, request='T1')]})
        self.upload({'MAKİNE': [row(1, request='T9')]})
        ayrilmis = next(k for k in depo._kartlar if k['talep_no'] == 'T1')
        kopya = dict(ayrilmis, id=99, source_row_id='NO:1~99')
        kopya['source_key'] = kopya['anahtar'] = depo.kaynak_anahtari('MAKINE', 'NO:1~99')
        depo._kartlar.append(kopya)
        sonuc = self.upload({'MAKİNE': [row(1, request='T1')]})
        self.assertEqual((sonuc['geri_baglanan'], sonuc['yeni']), (0, 1))
        pasifler = {k['id'] for k in depo._kartlar if k['source_active'] == 0}
        self.assertTrue(pasifler.issuperset({ayrilmis['id'], 99}))


class IslemYapanTests(_Ortak):
    def post(self, client, yol, veri):
        return client.post(yol, json=veri, headers=HEADERS)

    def test_islem_yapan_karta_ve_her_log_satirina_yazilir(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1, end='2099-12-31')]})
        kid = depo._kartlar[0]['id']
        for yol, veri in (('/api/basla', {'adet': 5}), ('/api/bitir', {'adet': 5}), ('/api/teslim-et', {}),
                          ('/api/not', {'not': 'teslim notu'})):
            yanit = self.post(client, yol, {'kart_id': kid, 'isim': 'Ahmet Yılmaz', **veri})
            self.assertEqual(yanit.status_code, 200, (yol, yanit.get_data(as_text=True)))
        self.assertEqual(depo.kart_getir(kid)['operator'], 'Ahmet Yılmaz')
        for islem in ('DİZGİYE ALINDI', 'ÜRETİM ADEDİ TAMAMLANDI', 'TESLİM EDİLDİ', 'NOT EKLENDİ'):
            kayit = self.loglar(islem)[-1]
            self.assertEqual((kayit['kullanici'], kayit['islem_yapan']), ('tester', 'Ahmet Yılmaz'), islem)
        diskteki = depo._oku(depo.LOG_DOSYA, depo.LOG_ALANLARI)
        self.assertEqual(diskteki[-1]['islem_yapan'], 'Ahmet Yılmaz')
        self.assertIn('Ahmet Yılmaz (tester)', client.get('/yonetim').get_data(as_text=True))

    def test_isim_gonderilmezse_hesap_adi_uzun_isim_reddedilir(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1)]})
        kid = depo._kartlar[0]['id']
        self.assertEqual(self.post(client, '/api/basla', {'kart_id': kid, 'adet': 5, 'isim': 'x' * 81}).status_code,
                         400)
        self.assertEqual(depo.kart_getir(kid)['durum'], 'PLANA ALINDI')
        self.assertEqual(self.post(client, '/api/basla', {'kart_id': kid, 'adet': 5}).status_code, 200)
        self.assertEqual(depo.kart_getir(kid)['operator'], 'Tester')

    def test_eski_log_dosyasi_yeni_sutun_olmadan_okunur(self):
        eski_alanlar = [a for a in depo.LOG_ALANLARI if a[1] != 'islem_yapan']
        depo._yaz(depo.LOG_DOSYA, eski_alanlar, [depo._log_kaydi('u', 'admin', 'GİRİŞ YAPILDI')], 'İşlem Logu')
        depo.kur()
        self.assertEqual(depo._loglar[0]['islem'], 'GİRİŞ YAPILDI')
        self.assertIsNone(depo._loglar[0]['islem_yapan'])
        depo.log_ekle('u', 'admin', 'ÇIKIŞ YAPILDI')
        self.assertEqual(len(depo._oku(depo.LOG_DOSYA, depo.LOG_ALANLARI)), 2)

    def test_arayuz_hesap_adini_kisi_adi_saymaz_ve_adet_bos_acilir(self):
        js = (ROOT / 'static' / 'js' / 'operator.js').read_text(encoding='utf-8')
        self.assertNotIn('|| document.body.dataset.ad ||', js)
        self.assertIn('hesapAdiMi(isim)', js)
        self.assertIn('el("bitir-adet").value = "";', js)

    def test_cok_uzun_log_detayi_islemi_engellemez(self):
        kayit = depo._log_kaydi('u', 'admin', 'X', detay='a' * 40_000)
        self.assertLessEqual(len(kayit['detay']), depo.LOG_DETAY_SINIRI + 100)
        self.assertIn('kısaltıldı', kayit['detay'])


class AdminDuzenlemeIziTests(_Ortak):
    def test_degisen_alanlarin_eski_ve_yeni_degeri_ve_silinen_not_loglanir(self):
        self.upload({'MAKİNE': [row(1, end='2026-10-01')]})
        kid = depo._kartlar[0]['id']
        depo.kart_not_guncelle(kid, 'operatörün notu', 'makine1', 'operator', isim='Ali')
        depo.admin_kart_duzenle(kid, 'PLANA ALINDI', 0, 9, '', 'admin', plan_teslim='2026-12-01')
        duzenleme = self.loglar('ADMİN DÜZENLEDİ')[-1]['detay']
        degisen = json.loads(duzenleme.split('değişenler: ', 1)[1])
        self.assertEqual(degisen['toplam_adet'], [5, 9])
        self.assertEqual(degisen['plan_teslim'], ['2026-10-01', '2026-12-01'])
        not_kaydi = self.loglar('ADMİN NOT DÜZENLEDİ')[-1]['detay']
        self.assertIn('1 satır silindi', not_kaydi)
        self.assertIn('Ali: operatörün notu', not_kaydi)

    def test_not_degismezse_not_kaydi_yazilmaz(self):
        self.upload({'MAKİNE': [row(1)]})
        kid = depo._kartlar[0]['id']
        depo.admin_kart_duzenle(kid, 'DİZGİDE', 0, 5, None, 'admin')
        self.assertEqual(self.loglar('ADMİN NOT DÜZENLEDİ'), [])
        self.assertIn('"durum": ["PLANA ALINDI", "DİZGİDE"]', self.loglar('ADMİN DÜZENLEDİ')[-1]['detay'])


class TekrarKartTests(_Ortak):
    def test_ayni_talep_stok_icin_acik_onay_istenir(self):
        self.upload({'MAKİNE': [row(1, request='123', stock='ABC')]})
        with self.assertRaises(depo.TekrarKartOnayiGerekli) as hata:
            depo.admin_kart_ekle('123', 'abc', 3, 'admin')                        # büyük/küçük harf farkı
        self.assertEqual([k['kaynak'] for k in hata.exception.kartlar], ['EXCEL'])
        self.assertEqual(len(depo._kartlar), 1)
        birinci = depo.admin_kart_ekle('123', 'ABC', 3, 'admin', tekrar_onayi=True)
        ikinci = depo.admin_kart_ekle('123', 'ABC', 4, 'admin', tekrar_onayi=True)   # iki manuel de olabilir
        self.assertEqual({birinci['anahtar'], ikinci['anahtar']},
                         {f"MANUEL:{birinci['id']}", f"MANUEL:{ikinci['id']}"})
        self.assertIn('admin onayıyla eklendi', self.loglar('MANUEL KART EKLENDİ')[-1]['detay'])
        depo.kur()
        self.assertEqual(len(depo.kartlari_getir()), 3)

    def test_pasif_veya_gizli_kart_onay_istemez(self):
        self.upload({'MAKİNE': [row(1, request='123', stock='ABC'), row(2, request='456', stock='DEF')]})
        self.upload({'MAKİNE': [row(2, request='456', stock='DEF')]})                 # 123 pasifleşti
        depo.admin_kart_gizle(next(k['id'] for k in depo._kartlar if k['talep_no'] == '456'), 'admin')
        depo.admin_kart_ekle('123', 'ABC', 3, 'admin')
        depo.admin_kart_ekle('456', 'DEF', 3, 'admin')

    def test_api_onay_bayragi(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1, request='123', stock='ABC')]})
        veri = {'talep_no': '123', 'stok_no': 'ABC', 'toplam_adet': 3}
        yanit = client.post('/api/admin/kart-ekle', json=veri, headers=HEADERS)
        self.assertEqual(yanit.status_code, 409)
        self.assertTrue(yanit.get_json()['tekrar_onayi_gerekli'])
        self.assertEqual(yanit.get_json()['mevcut'][0]['id'], depo._kartlar[0]['id'])
        yanit = client.post('/api/admin/kart-ekle', json={**veri, 'tekrar_onayi': True}, headers=HEADERS)
        self.assertEqual(yanit.status_code, 201)
        js = (ROOT / 'static' / 'js' / 'yonetim.js').read_text(encoding='utf-8')
        self.assertIn('tekrar_onayi_gerekli', js)


if __name__ == '__main__':
    unittest.main()
