"""Import önizlemesinde "Tamamlanan adedi sıfırla" seçeneği; geçici klasörde çalışır."""
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


class TamamlananSifirlaTests(unittest.TestCase):
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
        """1: DİZGİDE 3/10 (operatör girdi), 2: TESLİM 5/5, 3: DİZGİDE 0/4, 4: DİZGİDE 2/6 değişmeyecek."""
        self.yukle({'MAKİNE': [row(1, qty=10, status='DİZGİDE', end='2099-12-01'),
                               row(2, qty=5, status='TESLİM EDİLDİ', actual='2026-09-20', stock='B'),
                               row(3, qty=4, status='DİZGİDE', end='2099-12-01', stock='C'),
                               row(4, qty=6, status='DİZGİDE', end='2099-12-01', stock='D')]})
        k = {x['sira']: x for x in depo._kartlar}
        depo.kart_bitir(k[1]['id'], 3, 'op', 'operator', 'makine', aciklama='3 adet bitti')
        depo.kart_bitir(k[4]['id'], 2, 'op', 'operator', 'makine')
        # Yeni Excel: 1 ve 3 tarih değiştirir, 2 DİZGİDE'ye geri döner, 4 aynı kalır.
        yeni = {'MAKİNE': [row(1, qty=10, status='DİZGİDE', end='2099-12-05'),
                           row(2, qty=5, status='DİZGİDE', end='2099-12-05', stock='B'),
                           row(3, qty=4, status='DİZGİDE', end='2099-12-05', stock='C'),
                           row(4, qty=6, status='DİZGİDE', end='2099-12-01', stock='D')]}
        return {x['sira']: x['id'] for x in depo._kartlar}, yeni

    def kart(self, kart_id):
        return depo.kart_getir(kart_id)

    def test_onizleme_yalniz_dizgide_kalan_ve_guncellenen_kartlari_sunar(self):
        idler, yeni = self.hazirla()
        onizleme = self.yukle(yeni, onizleme=True)
        # 3'te tamamlanan 0 (sıfırlanacak bir şey yok), 4 güncellenmiyor.
        self.assertEqual(onizleme['sifirlanabilir'], sorted([idler[1], idler[2]]))
        kartlar = {d['id']: d for d in onizleme['degisiklikler']}
        self.assertEqual((kartlar[idler[1]]['sifirlanabilir'], kartlar[idler[1]]['tamamlanan_adet']), (True, 3))
        self.assertEqual((kartlar[idler[2]]['sifirlanabilir'], kartlar[idler[2]]['tamamlanan_adet']), (True, 5))
        self.assertFalse(kartlar[idler[3]]['sifirlanabilir'])

    def test_secilmezse_korunur_secilirse_sifirlanir(self):
        idler, yeni = self.hazirla()
        once = self.kart(idler[1])
        sonuc = self.yukle(yeni, tamamlanan_sifirla=[idler[2]])
        self.assertEqual(sonuc['sifirlanan'], 1)
        korunan, sifirlanan = self.kart(idler[1]), self.kart(idler[2])
        self.assertEqual(korunan['tamamlanan_adet'], 3)                         # seçilmedi
        self.assertEqual((sifirlanan['durum'], sifirlanan['tamamlanan_adet'], sifirlanan['bitis_zamani'],
                          sifirlanan['baslangic_adet']), ('DİZGİDE', 0, None, 5))
        self.assertTrue(sifirlanan['rozet'].startswith('SÜRESİ İÇİNDE'), sifirlanan['rozet'])  # "ÜRETİM BİTTİ" değil
        self.assertEqual(korunan['rozet'].split(' (')[0], 'SÜRESİ İÇİNDE')
        self.assertEqual(self.kart(idler[4])['tamamlanan_adet'], 2)            # dokunulmadı
        log = [l for l in depo._loglar if l['islem'] == 'EXCEL KART GÜNCELLENDİ' and f"ID={idler[2]} " in l['detay']]
        self.assertIn('"tamamlanan_adet": [5, 0]', log[-1]['detay'])
        self.assertIn('1 kartın tamamlanan adedi admin seçimiyle sıfırlandı', depo._loglar[-1]['detay'])
        # Not ve başlama bilgisi korunur.
        self.assertEqual(korunan['aciklama'], once['aciklama'])

    def test_uygun_olmayan_secim_tum_importu_iptal_eder(self):
        idler, yeni = self.hazirla()
        once, disk = copy.deepcopy(depo._kartlar), Path(depo.KARTLAR_DOSYA).read_bytes()
        for secim in ([idler[4]], [idler[3]], [999]):      # değişmeyen, adedi 0 olan, olmayan kart
            with self.subTest(secim=secim):
                with self.assertRaisesRegex(depo.IsKuralHatasi, 'Aktarım uygulanmadı'):
                    self.yukle(yeni, tamamlanan_sifirla=secim + [idler[1]])
                self.assertEqual(depo._kartlar, once)
                self.assertEqual(Path(depo.KARTLAR_DOSYA).read_bytes(), disk)

    def test_plana_ve_teslimde_adet_kurali_degismez(self):
        idler, _ = self.hazirla()
        self.yukle({'MAKİNE': [row(1, qty=10, status='PLANA ALINDI', end='2099-12-05'),
                               row(2, qty=5, status='TESLİM EDİLDİ', actual='2026-09-21', stock='B'),
                               row(3, qty=4, status='DİZGİDE', end='2099-12-05', stock='C'),
                               row(4, qty=6, status='DİZGİDE', end='2099-12-01', stock='D')]})
        self.assertEqual(self.kart(idler[1])['tamamlanan_adet'], 0)    # PLANA: kendiliğinden 0
        self.assertEqual(self.kart(idler[2])['tamamlanan_adet'], 5)    # TESLİM: toplam

    def test_arayuz_onay_kutusu_ve_onay_akisi(self):
        _, client = self.load_app()
        idler, yeni = self.hazirla()
        yol = self.book(yeni)

        def snapshot(kaynak):
            hedef = self.root / 'onizleme.xlsx'; shutil.copy2(kaynak, hedef); return str(hedef)

        def onizle():
            with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot), open(yol, 'rb') as f:
                yanit = client.post('/yonetim/yukle', data={'_csrf_token': 'test-token', 'onizleme': '1',
                                                           'dosya': (f, 'plan.xlsx')})
            self.assertEqual(yanit.status_code, 200)
            return yanit.get_data(as_text=True)

        def onayla(secim):
            # Kart 2 (Excel'den TESLİM gelmişti, Excel artık DİZGİDE) "Karar gerekiyor" bölümünde;
            # önerilen seçim Excel'in durumu.
            veri = {'_csrf_token': 'test-token', 'sifirla': secim, f'gerileme_{idler[2]}': 'DİZGİDE'}
            with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
                return client.post('/yonetim/yukle-onay', data=veri)

        html = onizle()
        kutular = re.findall(r'<input type="checkbox" name="sifirla" value="(\d+)" form="onay-formu"', html)
        self.assertEqual(sorted(map(int, kutular)), sorted([idler[1], idler[2]]))
        self.assertIn('Şu an 3 / 10 tamamlanmış görünüyor', html)
        self.assertIn('<form id="onay-formu"', html)
        self.assertIn("<option value=\"DİZGİDE\" selected>DİZGİDE · Excel'deki (önerilen)</option>", html)

        # Önizlemede sunulmayan kart seçilemez (form elle değiştirilmiş olsa bile).
        yanit = onayla([str(idler[4])])
        self.assertEqual(yanit.status_code, 409)
        self.assertIn('önizlemede sunulmayan kart', yanit.get_data(as_text=True))
        self.assertEqual(self.kart(idler[4])['tamamlanan_adet'], 2)

        onizle()
        yanit = onayla([str(idler[1])])
        self.assertEqual(yanit.status_code, 302)
        with client.session_transaction() as oturum:
            self.assertIn('1 kartın tamamlanan adedi seçiminizle sıfırlandı', oturum['_flashes'][-1][1])
        self.assertEqual((self.kart(idler[1])['tamamlanan_adet'], self.kart(idler[2])['tamamlanan_adet']), (0, 5))
        self.assertIn('3 adet bitti', self.kart(idler[1])['aciklama'])


if __name__ == '__main__':
    unittest.main()
