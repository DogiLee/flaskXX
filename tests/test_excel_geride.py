"""Excel'in uygulamanın gerisinde kaldığı kartlar: önizlemede admin durum kararı.

Operatör uygulamada ilerlettiği kartı (dizgiye aldı, adet girdi, teslim etti) Excel
henüz güncellenmeden yeniden yüklenirse, kart artık sessizce geri alınmaz: önizleme
kartı "Karar gerekiyor" bölümüne koyar, admin durumu seçer. Geçici klasörde çalışır.
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

EXCEL = {'MAKİNE': [row(1, qty=10, status='PLANA ALINDI', end='2099-12-01'),
                    row(2, qty=4, status='DİZGİDE', end='2099-12-01', stock='B'),
                    row(3, qty=5, status='TESLİM EDİLDİ', actual='2026-09-20', stock='C'),
                    row(4, qty=6, status='DİZGİDE', end='2099-12-01', stock='D')]}


class ExcelGerideTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    load_app = ui.UITests.load_app

    def yukle(self, sayfalar, **kw):
        yol = self.book(sayfalar)
        def snapshot(_):
            hedef = self.root / 'snapshot.xlsx'; shutil.copy2(yol, hedef); return str(hedef)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            return ex.excelden_aktar(str(yol), 'tester', **kw)

    def hazirla(self):
        """1: operatör dizgiye aldı 6/10 (Excel PLANA). 2: operatör teslim etti (Excel DİZGİDE).
        3: Excel'den TESLİM gelmişti, Excel artık DİZGİDE (kasıtlı geri alma). 4: değişmiyor."""
        self.yukle(EXCEL)
        k = {x['sira']: x['id'] for x in depo._kartlar}
        depo.kart_baslat(k[1], 10, 'op', 'operator', 'makine')
        depo.kart_bitir(k[1], 6, 'op', 'operator', 'makine', aciklama='6 bitti')
        depo.kart_bitir(k[2], 4, 'op', 'operator', 'makine')
        depo.kart_teslim_et(k[2], 'op', 'operator', 'makine')
        yeni = copy.deepcopy(EXCEL)
        yeni['MAKİNE'][0][5] = '2099-12-09'                  # 1: Excel plan teslim de değişti
        yeni['MAKİNE'][2] = row(3, qty=5, status='DİZGİDE', end='2099-12-01', stock='C')
        return k, yeni

    def kart(self, kart_id):
        return depo.kart_getir(kart_id)

    def test_onizleme_gerileyen_kartlari_ve_onerileri_gosterir(self):
        k, yeni = self.hazirla()
        onizleme = self.yukle(yeni, onizleme=True)
        self.assertEqual(onizleme['gerileme'], 3)
        g = {d['sira']: d for d in onizleme['degisiklikler'] if d['tur'] == 'gerileme'}
        self.assertEqual(set(g), {1, 2, 3})
        self.assertEqual((g[1]['gerileme']['secenekler'], g[1]['gerileme']['onerilen'], g[1]['gerileme']['uygulamada']),
                         (['PLANA ALINDI', 'DİZGİDE'], 'DİZGİDE', True))
        self.assertEqual((g[2]['gerileme']['secenekler'], g[2]['gerileme']['onerilen']),
                         (['DİZGİDE', 'TESLİM EDİLDİ'], 'TESLİM EDİLDİ'))
        self.assertEqual((g[3]['gerileme']['onerilen'], g[3]['gerileme']['uygulamada']), ('DİZGİDE', False))
        # Önizleme önerilen seçimle hesaplanır: operatörün ilerlemesi görünür şekilde korunur.
        self.assertEqual((g[1]['durum'], g[1]['tamamlanan_adet'], g[1]['plan_teslim']), ('DİZGİDE', 6, '2099-12-09'))
        self.assertEqual(g[2]['durum'], 'TESLİM EDİLDİ')
        self.assertTrue({k[1], k[2], k[3]} <= set(onizleme['sifirlanabilir']))
        self.assertEqual(self.kart(k[1])['durum'], 'DİZGİDE')   # önizleme yazmaz

    def test_secimler_uygulanir(self):
        k, yeni = self.hazirla()
        teslim_tarihi = self.kart(k[2])['gerceklesen_teslim']
        sonuc = self.yukle(yeni, gerileme_secimleri={k[1]: 'DİZGİDE', k[2]: 'TESLİM EDİLDİ', k[3]: 'DİZGİDE'})
        self.assertEqual((sonuc['gerileme'], sonuc['gerileme_korunan']), (3, 2))
        bir, iki, uc = self.kart(k[1]), self.kart(k[2]), self.kart(k[3])
        # 1: durum ve adet korunur, Excel'in diğer alanları (plan teslim) yine gelir.
        self.assertEqual((bir['durum'], bir['tamamlanan_adet'], bir['plan_teslim'], bir['excel_durum']),
                         ('DİZGİDE', 6, '2099-12-09', 'PLANA ALINDI'))
        self.assertIsNotNone(bir['baslama_zamani'])
        self.assertIn('6 bitti', bir['aciklama'])
        # 2: teslim, tarih ve zaman damgasıyla korunur; teslim listesinde kalır.
        self.assertEqual((iki['durum'], iki['tamamlanan_adet'], iki['gerceklesen_teslim']),
                         ('TESLİM EDİLDİ', 4, teslim_tarihi))
        self.assertIsNotNone(iki['teslim_zamani'])
        # 3: Excel'deki kasıtlı geri alma uygulanır.
        self.assertEqual(uc['durum'], 'DİZGİDE')
        kararlar = [l['detay'] for l in depo._loglar if l['islem'] == 'EXCEL GERİDE: DURUM KARARI']
        self.assertEqual(len(kararlar), 3)
        self.assertTrue(any('uygulama DİZGİDE, Excel PLANA ALINDI -> DİZGİDE (uygulamadaki korundu)' in d
                            for d in kararlar))

    def test_excel_e_gore_geri_alma_ve_ara_durum(self):
        k, yeni = self.hazirla()
        self.yukle(yeni, gerileme_secimleri={k[1]: 'PLANA ALINDI', k[2]: 'DİZGİDE', k[3]: 'DİZGİDE'},
                   tamamlanan_sifirla=[k[2]])
        bir, iki = self.kart(k[1]), self.kart(k[2])
        self.assertEqual((bir['durum'], bir['tamamlanan_adet'], bir['baslama_zamani']), ('PLANA ALINDI', 0, None))
        self.assertEqual((iki['durum'], iki['tamamlanan_adet'], iki['gerceklesen_teslim'], iki['teslim_zamani']),
                         ('DİZGİDE', 0, None, None))

    def test_uc_basamak_geri_kalma_secenekleri(self):
        """Uygulamada TESLİM, Excel PLANA: PLANA / DİZGİDE / TESLİM arasından seçilir."""
        self.yukle({'MAKİNE': [row(1, qty=3, status='DİZGİDE', end='2099-12-01')]})
        kid = depo._kartlar[0]['id']
        depo.kart_bitir(kid, 3, 'op', 'operator', 'makine'); depo.kart_teslim_et(kid, 'op', 'operator', 'makine')
        excel = {'MAKİNE': [row(1, qty=3, status='PLANA ALINDI', end='2099-12-01')]}
        g = self.yukle(excel, onizleme=True)['degisiklikler'][0]['gerileme']
        self.assertEqual(g['secenekler'], ['PLANA ALINDI', 'DİZGİDE', 'TESLİM EDİLDİ'])
        self.yukle(excel, gerileme_secimleri={kid: 'DİZGİDE'})
        self.assertEqual((self.kart(kid)['durum'], self.kart(kid)['tamamlanan_adet']), ('DİZGİDE', 3))

    def test_gecersiz_secimler_aktarimi_iptal_eder(self):
        k, yeni = self.hazirla()
        once, disk = copy.deepcopy(depo._kartlar), Path(depo.KARTLAR_DOSYA).read_bytes()
        tam = {k[1]: 'DİZGİDE', k[2]: 'TESLİM EDİLDİ', k[3]: 'DİZGİDE'}
        durumlar = [
            ({k[1]: 'DİZGİDE'}, None, 'geçerli bir durum seçilmedi'),                        # eksik karar
            ({**tam, k[1]: 'TESLİM EDİLDİ'}, None, 'geçerli bir durum seçilmedi'),           # aralık dışı
            ({**tam, k[4]: 'DİZGİDE'}, None, 'Excel.in gerisinde değil'),                   # gerilemeyen kart
            ({**tam, k[1]: 'PLANA ALINDI'}, [k[1]], 'DİZGİDE kalmıyor'),                     # PLANA'da sıfırlama
        ]
        for secim, sifirla, mesaj in durumlar:
            with self.subTest(mesaj=mesaj, secim=secim):
                with self.assertRaisesRegex(depo.IsKuralHatasi, mesaj):
                    self.yukle(yeni, gerileme_secimleri=secim, tamamlanan_sifirla=sifirla)
                self.assertEqual(depo._kartlar, once)
                self.assertEqual(Path(depo.KARTLAR_DOSYA).read_bytes(), disk)

    def test_onizlemesiz_dogrudan_import_eskisi_gibi_excel_i_uygular(self):
        k, yeni = self.hazirla()
        self.yukle(yeni)
        self.assertEqual((self.kart(k[1])['durum'], self.kart(k[1])['tamamlanan_adet']), ('PLANA ALINDI', 0))
        self.assertEqual(self.kart(k[2])['durum'], 'DİZGİDE')

    def test_arayuz_karar_bolumu_ve_onay(self):
        _, client = self.load_app()
        k, yeni = self.hazirla()
        yol = self.book(yeni)

        def snapshot(kaynak):
            hedef = self.root / 'onizleme.xlsx'; shutil.copy2(kaynak, hedef); return str(hedef)

        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot), open(yol, 'rb') as f:
            yanit = client.post('/yonetim/yukle', data={'_csrf_token': 'test-token', 'onizleme': '1',
                                                       'dosya': (f, 'plan.xlsx')})
        html = yanit.get_data(as_text=True)
        self.assertEqual(yanit.status_code, 200)
        self.assertIn('<section class="panel-kutu karar-bolumu" id="karar-gerekiyor">', html)
        self.assertIn('3 kartta Excel uygulamanın gerisinde.', html)
        self.assertIn('Kart uygulamada ilerletilmiş; Excel henüz güncellenmemiş olabilir.', html)
        self.assertIn("Excel'de geri alınmış görünüyor.", html)
        secimler = dict(re.findall(r'<select name="gerileme_(\d+)"[^>]*>.*?<option value="([^"]+)" selected',
                                   html, re.S))
        self.assertEqual(secimler, {str(k[1]): 'DİZGİDE', str(k[2]): 'TESLİM EDİLDİ', str(k[3]): 'DİZGİDE'})
        self.assertIn('Hepsinde uygulamadakini koru', html)
        # Admin önerileri kabul eder, yalnız 1. kartta adedi sıfırlar.
        veri = {'_csrf_token': 'test-token', 'sifirla': [str(k[1])],
                **{f'gerileme_{i}': d for i, d in secimler.items()}}
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            yanit = client.post('/yonetim/yukle-onay', data=veri)
        self.assertEqual(yanit.status_code, 302)
        with client.session_transaction() as oturum:
            self.assertIn('3 kartta Excel uygulamanın gerisindeydi; 2 kartta uygulamadaki durum korundu',
                          oturum['_flashes'][-1][1])
        self.assertEqual((self.kart(k[1])['durum'], self.kart(k[1])['tamamlanan_adet']), ('DİZGİDE', 0))
        self.assertEqual(self.kart(k[2])['durum'], 'TESLİM EDİLDİ')
        teslimler = client.get('/api/panel/teslimler').get_json()['teslim_edilen']
        self.assertIn('B', {t['stok_no'] for t in teslimler})


if __name__ == '__main__':
    unittest.main()
