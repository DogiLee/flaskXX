"""Excel'de DURUM boş olan kartlar için durum kararı ve "Notları temizle" seçeneği.

1) Uygulamada bir durumu olan kartın DURUM hücresi Excel'de boşaltılırsa (veya
   MALZEME TEDARİK / PDGM ÖNERİ gibi iş akışı durumu olmayan bir metin yazılırsa)
   durum artık sessizce korunmaz: önizleme kartı "Karar gerekiyor" bölümüne koyar,
   admin durumsuz bırakmak ile uygulamadaki durumu korumak arasında seçer.
2) Durumu değişen kartın notları admin isterse temizlenir; silinen not işlem loguna yazılır.
Geçici klasörde çalışır; üretim verisine dokunmaz.
"""
import copy
import re
import shutil
import unittest
from pathlib import Path
from unittest.mock import patch

import depo
import excel_araclari as ex
from tests import test_excel_sync as fixtures, test_ui as ui
from tests.test_excel_sync import row

EXCEL = {'MAKİNE': [row(1, qty=10, status='DİZGİDE', end='2099-12-01'),
                    row(2, qty=4, status='PLANA ALINDI', end='2099-12-01', stock='B'),
                    row(3, qty=5, status=None, end='2099-12-01', stock='C'),
                    row(4, qty=6, status='PLANA ALINDI', end='2099-12-01', stock='D'),
                    row(5, qty=3, status='DİZGİDE', end='2099-12-01', stock='E')]}


class DurumsuzVeNotTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    load_app = ui.UITests.load_app

    def yukle(self, sayfalar, **kw):
        yol = self.book(sayfalar)
        def snapshot(_):
            hedef = self.root / 'snapshot.xlsx'; shutil.copy2(yol, hedef); return str(hedef)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            return ex.excelden_aktar(str(yol), 'tester', **kw)

    def kart(self, kart_id):
        return depo.kart_getir(kart_id)

    def hazirla(self):
        """1: DİZGİDE Excel'den geldi, Excel'de DURUM silinecek.
        2: operatör dizgiye aldı ve 2 adet girdi; Excel'de DURUM silinecek.
        3: Excel'de hep boştu, admin PLANA ALINDI atadı; Excel hâlâ boş.
        4: PLANA ALINDI -> Excel'de 'MALZEME TEDARİK'. 5: notlu, Excel DİZGİDE -> TESLİM."""
        self.yukle(EXCEL)
        k = {x['sira']: x['id'] for x in depo._kartlar}
        depo.kart_baslat(k[2], 4, 'op', 'operator', 'makine')
        depo.kart_bitir(k[2], 2, 'op', 'operator', 'makine')
        kart3 = self.kart(k[3])
        depo.admin_kart_duzenle(k[3], 'PLANA ALINDI', 0, kart3['toplam_adet'], None, 'admin')
        for no in (1, 5):
            depo.kart_not_guncelle(k[no], f'Kart {no} eski notu', 'op', 'operator')
        yeni = copy.deepcopy(EXCEL)
        yeni['MAKİNE'][0][7] = None
        yeni['MAKİNE'][1][7] = None
        yeni['MAKİNE'][3][7] = 'MALZEME TEDARİK'
        yeni['MAKİNE'][4][7], yeni['MAKİNE'][4][6] = 'TESLİM EDİLDİ', '2026-09-25'
        return k, yeni

    def test_onizleme_durumu_bos_kartlar_icin_karar_ister(self):
        k, yeni = self.hazirla()
        onizleme = self.yukle(yeni, onizleme=True)
        self.assertEqual((onizleme['durumsuz'], onizleme['gerileme']), (4, 0))
        g = {d['sira']: d['gerileme'] for d in onizleme['degisiklikler'] if d['tur'] == 'gerileme'}
        self.assertEqual(set(g), {1, 2, 3, 4})
        self.assertEqual({no: (x['karar_turu'], x['secenekler'], x['onerilen']) for no, x in g.items()},
                         {1: ('durumsuz', [None, 'DİZGİDE'], None),           # Excel'den gelmişti -> Excel'e uy
                          2: ('durumsuz', [None, 'DİZGİDE'], 'DİZGİDE'),      # operatör ilerletti -> koru
                          3: ('durumsuz', [None, 'PLANA ALINDI'], 'PLANA ALINDI'),  # admin atadı -> koru
                          4: ('durumsuz', [None, 'PLANA ALINDI'], None)})
        # Önizleme hiçbir şeyi değiştirmez.
        self.assertEqual(self.kart(k[1])['durum'], 'DİZGİDE')

    def test_durumsuz_birakma_ve_koruma_secimleri_uygulanir(self):
        k, yeni = self.hazirla()
        sonuc = self.yukle(yeni, gerileme_secimleri={k[1]: '', k[2]: 'DİZGİDE', k[3]: 'PLANA ALINDI', k[4]: ''})
        self.assertEqual((sonuc['durumsuz'], sonuc['durumsuz_korunan'], sonuc['gerileme']), (4, 2, 0))
        kart1 = self.kart(k[1])
        self.assertEqual((kart1['durum'], kart1['tamamlanan_adet'], kart1['baslama_zamani']), (None, 0, None))
        self.assertIn(k[1], {x['id'] for x in depo.durumu_eksik_kartlari_getir()})
        self.assertNotIn(k[1], {x['id'] for x in depo.kartlari_getir()})
        self.assertEqual((self.kart(k[2])['durum'], self.kart(k[2])['tamamlanan_adet']), ('DİZGİDE', 2))
        self.assertEqual(self.kart(k[3])['durum'], 'PLANA ALINDI')
        kart4 = self.kart(k[4])
        self.assertEqual((kart4['durum'], kart4['excel_durum'], kart4['malzeme_bekliyor']), (None, 'MALZEME TEDARİK', True))
        loglar = [l for l in depo._loglar if l['islem'] == 'EXCEL DURUM BOŞ: DURUM KARARI']
        self.assertEqual(len(loglar), 4)
        self.assertTrue(any('durumsuz bırakıldı' in l['detay'] for l in loglar))
        self.assertTrue(any("'MALZEME TEDARİK' (iş akışı durumu değil)" in l['detay'] for l in loglar))

    def test_teslim_edilen_kart_korunursa_teslim_tarihi_kalir(self):
        self.yukle({'MAKİNE': [row(1, qty=2, status='DİZGİDE', end='2099-12-01')]})
        kid = depo._kartlar[0]['id']
        depo.kart_bitir(kid, 2, 'op', 'operator', 'makine')
        depo.kart_teslim_et(kid, 'op', 'operator', 'makine')
        teslim = self.kart(kid)['gerceklesen_teslim']
        bos = {'MAKİNE': [row(1, qty=2, status=None, end='2099-12-01')]}
        onizleme = self.yukle(bos, onizleme=True)
        self.assertEqual(onizleme['degisiklikler'][0]['gerileme']['onerilen'], 'TESLİM EDİLDİ')
        self.yukle(bos, gerileme_secimleri={kid: 'TESLİM EDİLDİ'})
        self.assertEqual((self.kart(kid)['durum'], self.kart(kid)['gerceklesen_teslim']), ('TESLİM EDİLDİ', teslim))

    def test_eksik_veya_gecersiz_karar_hicbir_seyi_degistirmez(self):
        k, yeni = self.hazirla()
        once = copy.deepcopy(depo._kartlar)
        disk = Path(depo.KARTLAR_DOSYA).read_bytes()
        tam = {k[1]: '', k[2]: 'DİZGİDE', k[3]: 'PLANA ALINDI', k[4]: ''}
        for secim, mesaj in [
            ({a: b for a, b in tam.items() if a != k[1]}, "DURUM'u boş olan kart için geçerli bir durum seçilmedi"),
            ({**tam, k[2]: 'TESLİM EDİLDİ'}, "DURUM'u boş olan kart için geçerli bir durum seçilmedi"),
        ]:
            with self.subTest(mesaj=mesaj, secim=secim):
                with self.assertRaisesRegex(depo.IsKuralHatasi, mesaj):
                    self.yukle(yeni, gerileme_secimleri=secim)
                self.assertEqual(depo._kartlar, once)
                self.assertEqual(Path(depo.KARTLAR_DOSYA).read_bytes(), disk)

    def test_onizlemesiz_dogrudan_import_durumu_korur(self):
        k, yeni = self.hazirla()
        sonuc = self.yukle(yeni)
        self.assertEqual((self.kart(k[1])['durum'], self.kart(k[2])['tamamlanan_adet']), ('DİZGİDE', 2))
        self.assertEqual(sonuc['durumsuz'], 0)

    def test_notlar_yalniz_durumu_degisen_kartta_ve_secilirse_temizlenir(self):
        k, yeni = self.hazirla()
        yeni['MAKİNE'][2][5] = '2099-12-05'                  # 3: yalnız tarih değişir, durum aynı
        depo.kart_not_guncelle(k[3], 'Kart 3 notu', 'op', 'operator')
        onizleme = self.yukle(yeni, onizleme=True)
        # 5: durum DİZGİDE -> TESLİM; 1: karar kartı; 3: karar kartı (Excel boş, admin atadı).
        self.assertEqual(set(onizleme['not_temizlenebilir']), {k[1], k[3], k[5]})
        kart5 = next(d for d in onizleme['degisiklikler'] if d['id'] == k[5])
        self.assertTrue(kart5['not_temizlenebilir'])
        self.assertIn('Kart 5 eski notu', kart5['aciklama'])
        secim = {k[1]: '', k[2]: 'DİZGİDE', k[3]: 'PLANA ALINDI', k[4]: ''}
        sonuc = self.yukle(yeni, gerileme_secimleri=secim, notlari_temizle=[k[5]])
        self.assertEqual(sonuc['notu_temizlenen'], 1)
        self.assertIsNone(self.kart(k[5])['aciklama'])
        self.assertIn('Kart 1 eski notu', self.kart(k[1])['aciklama'])       # seçilmeyen korunur
        log = next(l for l in depo._loglar if l['islem'] == 'EXCEL: NOTLAR TEMİZLENDİ')
        self.assertIn('Kart 5 eski notu', log['detay'])
        self.assertIn('DİZGİDE -> TESLİM EDİLDİ', log['detay'])

    def test_durumu_degismeyen_kartin_notu_temizlenemez(self):
        self.yukle({'MAKİNE': [row(1, qty=4, status='PLANA ALINDI', end='2099-12-01')]})
        kid = depo._kartlar[0]['id']
        depo.kart_not_guncelle(kid, 'Kalsın', 'op', 'operator')
        yeni = {'MAKİNE': [row(1, qty=4, status='PLANA ALINDI', end='2099-12-09')]}
        self.assertEqual(self.yukle(yeni, onizleme=True)['not_temizlenebilir'], [])
        once = copy.deepcopy(depo._kartlar)
        with self.assertRaisesRegex(depo.IsKuralHatasi, 'Notları temizlenmek üzere seçilen kartın'):
            self.yukle(yeni, notlari_temizle=[kid])
        self.assertEqual(depo._kartlar, once)

    def test_arayuz_durumsuz_karar_ve_not_temizleme(self):
        _, client = self.load_app()
        k, yeni = self.hazirla()
        yol = self.book(yeni)

        def snapshot(kaynak):
            hedef = self.root / 'onizleme.xlsx'; shutil.copy2(kaynak, hedef); return str(hedef)

        def onizle():
            with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot), open(yol, 'rb') as f:
                return client.post('/yonetim/yukle', data={'_csrf_token': 'test-token', 'onizleme': '1',
                                                          'dosya': (f, 'plan.xlsx')})

        yanit = onizle()
        html = yanit.get_data(as_text=True)
        self.assertEqual(yanit.status_code, 200)
        self.assertIn("4 kartta Excel'de DURUM boş, uygulamada ise bir durum var.", html)
        self.assertNotIn('kartta Excel uygulamanın gerisinde.', html)
        self.assertIn("Excel'de DURUM silinmiş görünüyor.", html)
        self.assertIn('Durum uygulamada verilmiş (operatör işlemi veya admin ataması)', html)
        self.assertIn('<option value="" selected>Durumsuz bırak · Excel\'deki (önerilen)</option>', html)
        self.assertIn('>MALZEME TEDARİK</span>', html)
        self.assertIn(f'name="not_temizle" value="{k[5]}"', html)
        self.assertIn('data-not-ozet hidden', html)
        self.assertNotIn('style=', html)
        secimler = dict(re.findall(r'<select name="gerileme_(\d+)"[^>]*>.*?<option value="([^"]*)" selected',
                                   html, re.S))
        self.assertEqual(secimler, {str(k[1]): '', str(k[2]): 'DİZGİDE', str(k[3]): 'PLANA ALINDI', str(k[4]): ''})

        # Önizlemede sunulmayan kartın notu seçilemez; hiçbir şey değişmez.
        once = copy.deepcopy(depo._kartlar)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            yanit = client.post('/yonetim/yukle-onay', data={'_csrf_token': 'test-token', 'not_temizle': [str(k[2])],
                                                             **{f'gerileme_{i}': d for i, d in secimler.items()}})
        self.assertEqual(yanit.status_code, 409)
        self.assertIn('Notları temizlenmek üzere önizlemede sunulmayan kart seçildi', yanit.get_data(as_text=True))
        self.assertEqual(depo._kartlar, once)

        onizle()
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            yanit = client.post('/yonetim/yukle-onay', data={'_csrf_token': 'test-token', 'not_temizle': [str(k[5])],
                                                             **{f'gerileme_{i}': d for i, d in secimler.items()}})
        self.assertEqual(yanit.status_code, 302)
        with client.session_transaction() as oturum:
            mesaj = oturum['_flashes'][-1][1]
        self.assertIn("4 kartta Excel'de DURUM boştu; 2 kartta uygulamadaki durum korundu, 2 kart durumsuz bırakıldı", mesaj)
        self.assertIn('1 kartın notları seçiminizle temizlendi', mesaj)
        self.assertEqual((self.kart(k[1])['durum'], self.kart(k[2])['durum']), (None, 'DİZGİDE'))
        self.assertIsNone(self.kart(k[5])['aciklama'])


if __name__ == '__main__':
    unittest.main()
