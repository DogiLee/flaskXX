"""Denetimdeki yüksek önemli üç bulgunun regresyonları (27.09.2026).

B1: Kayıt dosyaları her işlemde baştan yazılır; yazım doğrusal olmalı, log arşivi
    her commit yolunda (yalnız girişte değil) çalışmalı ve arşiv hatası işlemi bozmamalı.
B2: Uygulamada teslim edilen kartın tarihi, Excel TESLİM EDİLDİ'ye yetişip tarih
    hücresini boş bıraktığında silinmemeli.
B3: Onay formu kendi önizlemesine bağlı olmalı; başka sekmede önizlenen dosya uygulanmamalı.
"""
import copy
import os
import re
import shutil
import unittest
from pathlib import Path
from unittest.mock import patch

import depo
import excel_araclari as ex
from tests import test_excel_sync as fixtures, test_ui as ui
from tests.test_excel_sync import row


class TeslimTarihiTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    cards = fixtures.SyncTests.cards

    def upload(self, sheets, **kw):
        path = self.book(sheets)

        def snapshot(_):
            hedef = self.root / 'snapshot.xlsx'; shutil.copy2(path, hedef); return str(hedef)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            return ex.excelden_aktar(str(path), 'tester', **kw)

    def uygulamada_teslim_et(self):
        self.upload({'MAKİNE': [row(1, status='DİZGİDE', end='2099-12-31')]})
        kid = self.cards()[1]['id']
        depo.kart_bitir(kid, 5, 'makine1', 'operator', 'makine')
        depo.kart_teslim_et(kid, 'makine1', 'operator', 'makine')
        kart = self.cards()[1]
        self.assertEqual(kart['gerceklesen_teslim'], depo.bugun())
        return kart

    def test_excel_tarihsiz_teslime_yetisince_uygulamadaki_tarih_korunur(self):
        once = self.uygulamada_teslim_et()
        sonraki = {'MAKİNE': [row(1, status='TESLİM EDİLDİ', end='2099-12-31', actual=None)]}
        # Önizleme de tarihi silinecek gibi göstermemeli.
        onizleme = self.upload(sonraki, onizleme=True)
        farklar = {f['anahtar'] for d in onizleme['degisiklikler'] for f in d['farklar']}
        self.assertNotIn('gerceklesen_teslim', farklar)
        self.upload(sonraki)
        kart = self.cards()[1]
        self.assertEqual((kart['durum'], kart['gerceklesen_teslim'], kart['teslim_zamani']),
                         ('TESLİM EDİLDİ', once['gerceklesen_teslim'], once['teslim_zamani']))
        self.assertEqual(kart['rozet'], 'ZAMANINDA TESLİM')
        self.assertIsNotNone(kart['sapma'])
        # Sunucu yeniden başlasa da aynı kalır.
        depo.kur()
        self.assertEqual(self.cards()[1]['gerceklesen_teslim'], once['gerceklesen_teslim'])

    def test_excel_tarih_yazarsa_excel_tarihi_gecerli(self):
        self.uygulamada_teslim_et()
        self.upload({'MAKİNE': [row(1, status='TESLİM EDİLDİ', end='2099-12-31', actual='2026-09-30')]})
        kart = self.cards()[1]
        self.assertEqual((kart['gerceklesen_teslim'], kart['teslim_zamani']), ('2026-09-30', None))

    def test_excelden_gelen_tarih_excelde_silinirse_temizlenir(self):
        # Tarih uygulamada değil Excel'de oluştuysa Excel onu geri alabilir.
        self.upload({'MAKİNE': [row(1, status='TESLİM EDİLDİ', actual='2026-09-18')]})
        self.upload({'MAKİNE': [row(1, status='TESLİM EDİLDİ', actual=None)]})
        kart = self.cards()[1]
        self.assertEqual(kart['durum'], 'TESLİM EDİLDİ')
        self.assertIsNone(kart['gerceklesen_teslim'])

    def test_excel_teslimi_geri_alirsa_tarih_silinir(self):
        # Excel kartı DİZGİDE'ye çekiyor ve admin Excel'e uyuyor: kart teslim değil, tarih kalmaz.
        self.uygulamada_teslim_et()
        kid = self.cards()[1]['id']
        self.upload({'MAKİNE': [row(1, status='DİZGİDE', end='2099-12-31')]},
                    gerileme_secimleri={kid: 'DİZGİDE'})
        kart = self.cards()[1]
        self.assertEqual((kart['durum'], kart['gerceklesen_teslim'], kart['teslim_zamani']),
                         ('DİZGİDE', None, None))


class OnizlemeBaglamaTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    upload = fixtures.SyncTests.upload
    load_app = ui.UITests.load_app

    def onizle(self, client, stok):
        yol = self.book({'MAKİNE': [row(1, stock=stok), row(2, stock=stok + '2')]})

        def snapshot(kaynak):
            hedef = self.root / 'onizleme.xlsx'; shutil.copy2(kaynak, hedef); return str(hedef)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot), open(yol, 'rb') as f:
            yanit = client.post('/yonetim/yukle', data={'_csrf_token': 'test-token', 'onizleme': '1',
                                                       'dosya': (f, f'{stok}.xlsx')})
        self.assertEqual(yanit.status_code, 200)
        return ui.onizleme_token(client), yanit.get_data(as_text=True), snapshot

    def onayla(self, client, snapshot, token=None):
        veri = {'_csrf_token': 'test-token'}
        if token is not None:
            veri['onizleme_token'] = token
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot):
            return client.post('/yonetim/yukle-onay', data=veri)

    def stoklar(self):
        return sorted(k['stok_no'] for k in depo.kartlari_getir())

    def test_baska_sekmedeki_onizleme_uygulanmaz(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1, stock='ESKI')]})
        token_a, html_a, snapshot = self.onizle(client, 'A')
        self.assertIn(f'name="onizleme_token" value="{token_a}"', html_a)
        token_b, _, _ = self.onizle(client, 'B')          # ikinci sekme
        self.assertNotEqual(token_a, token_b)

        once, disk = copy.deepcopy(depo._kartlar), Path(depo.KARTLAR_DOSYA).read_bytes()
        yanit = self.onayla(client, snapshot, token_a)    # A'yı gösteren sekmeden onay
        self.assertEqual(yanit.status_code, 409)
        self.assertIn('Bu önizleme sayfası güncel değil', yanit.get_data(as_text=True))
        self.assertEqual(depo._kartlar, once)
        self.assertEqual(Path(depo.KARTLAR_DOSYA).read_bytes(), disk)

        # Reddedilen deneme B'nin önizlemesini tüketmez; B kendi sekmesinden onaylanır.
        self.assertEqual(self.onayla(client, snapshot, token_b).status_code, 302)
        self.assertEqual(self.stoklar(), ['B', 'B2'])

    def test_tokensiz_veya_yanlis_token_ile_onay_reddedilir(self):
        _, client = self.load_app()
        token, _, snapshot = self.onizle(client, 'A')
        for yanlis in (None, '', 'uydurma'):
            self.assertEqual(self.onayla(client, snapshot, yanlis).status_code, 409)
        self.assertEqual(depo._kartlar, [])
        self.assertEqual(self.onayla(client, snapshot, token).status_code, 302)
        self.assertEqual(self.stoklar(), ['A', 'A2'])
        # Kullanılmış önizleme ikinci kez uygulanamaz.
        self.assertEqual(self.onayla(client, snapshot, token).status_code, 409)

    def test_formlar_cift_gonderime_karsi_isaretli(self):
        _, client = self.load_app()
        self.assertIn('data-tek-gonderim data-gonderim-metni="Excel okunuyor…"',
                      client.get('/yonetim').get_data(as_text=True))
        _, html, _ = self.onizle(client, 'A')
        self.assertRegex(html, r'<form id="onay-formu"[^>]*data-tek-gonderim')
        ortak = (fixtures.ROOT / 'static' / 'js' / 'ortak.js').read_text(encoding='utf-8')
        self.assertIn('form.hasAttribute("data-tek-gonderim")', ortak)


class LogYazimiTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    upload = fixtures.SyncTests.upload
    cards = fixtures.SyncTests.cards

    def arsivler(self):
        klasor = Path(depo.YEDEK_KLASORU)
        return sorted(klasor.glob('*_islem_logu_arsiv.xlsx')) if klasor.exists() else []

    def test_cok_satirli_yazim_sirayi_ve_metni_korur(self):
        kayitlar = [depo._log_kaydi('u', 'operator', f'İŞLEM {i}', detay='=SUM(A1:A2)' if i % 2 else 'düz')
                    for i in range(2500)]
        depo._yaz(depo.LOG_DOSYA, depo.LOG_ALANLARI, kayitlar, 'İşlem Logu')
        okunan = depo._oku(depo.LOG_DOSYA, depo.LOG_ALANLARI)
        self.assertEqual([k['islem'] for k in okunan], [k['islem'] for k in kayitlar])
        self.assertEqual(okunan[1]['detay'], '=SUM(A1:A2)')    # formüle dönüşmez

    def test_arsiv_kart_islemlerinde_de_calisir_ve_kayit_kaybolmaz(self):
        self.upload({'MAKİNE': [row(1)]})
        kid = self.cards()[1]['id']
        with patch.object(depo, 'LOG_SINIRI', 10), patch.object(depo, 'LOG_SAKLA', 4):
            for i in range(25):   # yalnız kart işlemi; log_ekle (giriş) hiç çağrılmıyor
                depo.kart_not_guncelle(kid, f'not {i}', 'makine1', 'operator')
            self.assertLessEqual(len(depo._loglar), 10)
            arsivli = [k for yol in self.arsivler() for k in depo._oku(str(yol), depo.LOG_ALANLARI)]
            self.assertGreaterEqual(len(self.arsivler()), 2)
        # Aynı saniyedeki arşivlerin dosya sırası belirsiz: her not tam bir kez bulunmalı.
        notlar = [k['detay'] for k in arsivli + depo._loglar if k['islem'] == 'NOT EKLENDİ']
        self.assertEqual(sorted(int(n.rsplit(' ', 1)[-1]) for n in notlar), list(range(25)))
        diskteki = depo._oku(depo.LOG_DOSYA, depo.LOG_ALANLARI)
        self.assertEqual([k['detay'] for k in diskteki], [k['detay'] for k in depo._loglar])

    def test_arsiv_hatasi_islemi_bozmaz(self):
        self.upload({'MAKİNE': [row(1)]})
        kid = self.cards()[1]['id']
        asil_yaz = depo._yaz
        for bozulan in ('arsiv', 'log'):
            def yaz(dosya, *a, bozulan=bozulan):
                if (bozulan == 'arsiv' and 'arsiv' in os.path.basename(dosya)) or \
                        (bozulan == 'log' and dosya == depo.LOG_DOSYA and len(a[1]) <= 4):
                    raise PermissionError('kilitli')
                return asil_yaz(dosya, *a)
            with patch.object(depo, 'LOG_SINIRI', 10), patch.object(depo, 'LOG_SAKLA', 4), \
                    patch.object(depo, '_yaz', yaz):
                for i in range(12):
                    depo.kart_not_guncelle(kid, f'{bozulan} {i}', 'makine1', 'operator')
                self.assertGreater(len(depo._loglar), 10)      # arşivlenemedi, kayıt kaybolmadı
            self.assertEqual(self.arsivler(), [])               # yarım arşiv dosyası kalmadı
            self.assertEqual(len(depo._oku(depo.LOG_DOSYA, depo.LOG_ALANLARI)), len(depo._loglar))
            self.assertIn(f'{bozulan} 11', depo.kart_getir(kid)['aciklama'])

    def test_not_logu_yalniz_eklenen_satiri_yazar(self):
        self.upload({'MAKİNE': [row(1)]})
        kid = self.cards()[1]['id']
        depo.kart_not_guncelle(kid, 'birinci not', 'makine1', 'operator', isim='Ali')
        depo.kart_not_guncelle(kid, 'ikinci not', 'makine1', 'operator', isim='Ayşe')
        son = depo._loglar[-1]
        self.assertRegex(son['detay'], r'^\[[\d: -]+\] Ayşe: ikinci not$')
        self.assertNotIn('birinci not', son['detay'])
        self.assertIn('birinci not', depo.kart_getir(kid)['aciklama'])

    def test_rapor_kitabi_tum_satirlari_icerir(self):
        self.upload({'MAKİNE': [row(i, stock=f'S{i}') for i in range(1, 6)]})
        wb = ex.calisma_kitabi_uret(depo.kartlari_getir(False), depo.loglari_getir(), [['a', 1]])
        self.assertEqual(wb['Kart Durumları'].max_row, 6)
        self.assertEqual([c.value for c in wb['Kart Durumları']['E'][1:]], [f'S{i}' for i in range(1, 6)])
        wb.close()


if __name__ == '__main__':
    unittest.main()
