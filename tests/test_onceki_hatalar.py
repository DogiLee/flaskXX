"""Sahada yaşanan iki hatanın güncel kodda tekrar etmediğini ekran/API düzeyinde doğrular.

1.1 Aynı talebin 2 adedi elle, 5 adedi makinede: türler/durumlar ters, tarihler
    yanlış, Excel'de teslim edilmiş kart arayüzde "PLANDA".
1.2 Aynı Talep NO + Kart Stok No birden fazla satırda: her satır ayrı kart olmalı.
Geçici klasörde çalışır; üretim verisine dokunmaz.
"""
import re
import unittest

from flask import template_rendered

import depo
import excel_araclari as ex
from tests import test_excel_sync as fixtures, test_ui as ui
from tests.test_excel_sync import row

TALEP, STOK = '1900700', 'AD-7000-0001'


def makine(no=1, **kw):
    ayar = dict(qty=5, status='DİZGİDE', start='2026-09-21', end='2099-12-31', request=TALEP, stock=STOK)
    ayar.update(kw)
    return row(no, **ayar)


def elle(no=1, **kw):
    ayar = dict(qty=2, status='TESLİM EDİLDİ', start='2026-09-14', end='2026-09-18', actual='2026-09-17',
                request=TALEP, stock=STOK)
    ayar.update(kw)
    return row(no, **ayar)


class OncekiHatalarTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    upload = fixtures.SyncTests.upload
    load_app = ui.UITests.load_app

    def kart(self, kod, no):
        return next(k for k in depo._kartlar if k['source_sheet'] == kod and k['source_row_id'] == f'NO:{no}')

    def ekran(self, client):
        """Pano API'si, monitör bağlamı ve teslim listesi: kullanıcının gördüğü veri."""
        baglam = []
        def yakala(sender, template, context, **extra): baglam.append(context)
        with template_rendered.connected_to(yakala, client.application):
            client.get('/monitor')
        monitor = {k['id']: k['rozet'] for k in baglam[-1]['dizgide'] + baglam[-1]['plana_alindi']}
        pano = {k['id']: k for k in client.get('/api/veriler').get_json()['kartlar']}
        teslim = {t['stok_no'] + '/' + t['dizgi_kod']: t
                  for t in client.get('/api/panel/teslimler').get_json()['teslim_edilen']}
        return pano, monitor, teslim

    # --- 1.1 -----------------------------------------------------------------
    def test_ayni_talep_elle_bitti_makine_devam_ediyor(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [makine(1), row(2, stock='DIGER')], 'ELDE DİZGİ': [elle(1)]})
        m, e = self.kart('MAKINE', 1), self.kart('ELLE', 1)
        self.assertNotEqual(m['id'], e['id'])
        self.assertEqual((m['dizgi_tipi'], m['toplam_adet'], m['durum'], m['plan_baslama'], m['plan_teslim'],
                          m['gerceklesen_teslim']),
                         ('MAKİNE', 5, 'DİZGİDE', '2026-09-21', '2099-12-31', None))
        self.assertEqual((e['dizgi_tipi'], e['toplam_adet'], e['durum'], e['tamamlanan_adet'], e['plan_baslama'],
                          e['plan_teslim'], e['gerceklesen_teslim']),
                         ('ELLE DİZGİ', 2, 'TESLİM EDİLDİ', 2, '2026-09-14', '2026-09-18', '2026-09-17'))

        pano, monitor, teslim = self.ekran(client)
        self.assertEqual((pano[m['id']]['dizgi_kod'], pano[m['id']]['durum']), ('MAKINE', 'DİZGİDE'))
        self.assertEqual((pano[e['id']]['dizgi_kod'], pano[e['id']]['durum']), ('ELLE', 'TESLİM EDİLDİ'))
        self.assertIn(m['id'], monitor)                 # makine işi monitörde
        self.assertNotIn(e['id'], monitor)              # bitmiş elle iş monitörde değil
        self.assertEqual(monitor[m['id']], 'SÜRESİ İÇİNDE')
        self.assertEqual(set(teslim), {f'{STOK}/ELLE'})  # teslim listesinde yalnız elle kısmı
        self.assertEqual((teslim[f'{STOK}/ELLE']['toplam_adet'], teslim[f'{STOK}/ELLE']['teslim']),
                         (2, '17.09.2026'))

    def test_eski_hata_yolu_makine_satiri_silinince_elle_kimlik_calmaz(self):
        """Eski kodda: MAKİNE satırı silinince ELDE satırı MAKİNE kartının ID'sini alıyordu."""
        self.upload({'MAKİNE': [makine(1), row(2, stock='DIGER')], 'ELDE DİZGİ': [elle(1)]})
        m_id, e_id = self.kart('MAKINE', 1)['id'], self.kart('ELLE', 1)['id']
        # Sıra değişikliği + araya yeni satır + ELDE tarih değişikliği.
        self.upload({'MAKİNE': [row(3, stock='YENI'), row(2, stock='DIGER'), makine(1, end='2099-12-30')],
                     'ELDE DİZGİ': [elle(1, actual='2026-09-19')]})
        self.assertEqual((self.kart('MAKINE', 1)['id'], self.kart('ELLE', 1)['id']), (m_id, e_id))
        self.assertEqual(self.kart('MAKINE', 1)['plan_teslim'], '2099-12-30')
        self.assertEqual(self.kart('ELLE', 1)['gerceklesen_teslim'], '2026-09-19')
        # Makine satırı silinir: yalnız makine kartı pasifleşir, elle kartı aynen kalır.
        self.upload({'MAKİNE': [row(2, stock='DIGER')], 'ELDE DİZGİ': [elle(1, actual='2026-09-19')]})
        self.assertEqual(self.kart('MAKINE', 1)['source_active'], 0)
        elle_kart = self.kart('ELLE', 1)
        self.assertEqual((elle_kart['id'], elle_kart['dizgi_tipi'], elle_kart['toplam_adet'], elle_kart['durum'],
                          elle_kart['source_active']), (e_id, 'ELLE DİZGİ', 2, 'TESLİM EDİLDİ', 1))
        self.assertEqual(len([k for k in depo._kartlar if k['talep_no'] == TALEP]), 2)
        # Satır geri gelince aynı makine kartı döner; çoğalma yok.
        self.upload({'MAKİNE': [makine(1), row(2, stock='DIGER')], 'ELDE DİZGİ': [elle(1)]})
        self.assertEqual((self.kart('MAKINE', 1)['id'], self.kart('MAKINE', 1)['source_active']), (m_id, 1))
        self.assertEqual(len([k for k in depo._kartlar if k['talep_no'] == TALEP]), 2)

    def test_excelde_teslim_edilen_kart_planda_gorunmez(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [makine(1)]})
        m = self.kart('MAKINE', 1)
        depo.kart_bitir(m['id'], 2, 'op', 'operator', 'makine')       # uygulamada kısmi üretim 2/5
        for durum, tarih in [('TESLİM EDİLDİ', None), ('Teslim Edildi ', '2026-09-24')]:
            with self.subTest(durum=durum, tarih=tarih):
                self.upload({'MAKİNE': [makine(1, status=durum, actual=tarih)]})
                pano, monitor, teslim = self.ekran(client)
                self.assertEqual((pano[m['id']]['durum'], pano[m['id']]['tamamlanan_adet']), ('TESLİM EDİLDİ', 5))
                self.assertNotIn(m['id'], monitor)                     # "PLANDA" rozetiyle görünmez
                self.assertIn(f'{STOK}/MAKINE', teslim)
                self.assertEqual(teslim[f'{STOK}/MAKINE']['teslim'], '24.09.2026' if tarih else '—')

    # --- 1.2 -----------------------------------------------------------------
    def test_ayni_talep_ve_stok_birden_fazla_satir_ayri_kart(self):
        partiler = [makine(1, qty=3, status='TESLİM EDİLDİ', end='2026-09-10', actual='2026-09-10', start='2026-09-01'),
                    makine(2, qty=4, status='DİZGİDE', start='2026-09-21', end='2099-12-31'),
                    makine(3, qty=5, status='PLANA ALINDI', start='2099-10-01', end='2099-10-05')]
        self.upload({'MAKİNE': partiler, 'EÜM': [makine(1, qty=6, status='ÜRETİM PLANA ALINDI', start=None,
                                                         end='2099-10-09', planned='2099-10-06')]})
        kartlar = {(k['source_sheet'], k['source_row_id']): k for k in depo._kartlar}
        self.assertEqual(len(kartlar), 4)
        self.assertEqual({a: (k['toplam_adet'], k['durum']) for a, k in kartlar.items()}, {
            ('MAKINE', 'NO:1'): (3, 'TESLİM EDİLDİ'), ('MAKINE', 'NO:2'): (4, 'DİZGİDE'),
            ('MAKINE', 'NO:3'): (5, 'PLANA ALINDI'), ('EUM', 'NO:1'): (6, 'PLANA ALINDI')})
        idler = {a: k['id'] for a, k in kartlar.items()}
        # Satırlar yer değiştirir, ortadakinin tarihi değişir, ilki silinir: yalnız ilgili kart etkilenir.
        yeni = [makine(3, qty=5, status='PLANA ALINDI', start='2099-10-01', end='2099-10-05'),
                makine(2, qty=4, status='DİZGİDE', start='2026-09-22', end='2099-12-30')]
        self.upload({'MAKİNE': yeni, 'EÜM': [makine(1, qty=6, status='ÜRETİM PLANA ALINDI', start=None,
                                                    end='2099-10-09', planned='2099-10-06')]})
        kartlar = {(k['source_sheet'], k['source_row_id']): k for k in depo._kartlar}
        self.assertEqual({a: k['id'] for a, k in kartlar.items()}, idler)
        self.assertEqual((kartlar[('MAKINE', 'NO:2')]['plan_baslama'], kartlar[('MAKINE', 'NO:2')]['plan_teslim']),
                         ('2026-09-22', '2099-12-30'))
        self.assertEqual((kartlar[('MAKINE', 'NO:3')]['plan_teslim'], kartlar[('MAKINE', 'NO:3')]['toplam_adet']),
                         ('2099-10-05', 5))
        self.assertEqual(kartlar[('MAKINE', 'NO:1')]['source_active'], 0)
        self.assertEqual(sum(k['source_active'] for k in kartlar.values()), 3)


    # --- Aynı belirtileri hâlâ üretebilen yollar (26.09.2026 düzeltmesi) ------------
    def test_no_baska_talebe_verilince_eski_kart_ayrilir(self):
        """Eskiden: NO 7 yeni talebe verilince yeni talep TESLİM EDİLDİ ve eski notla görünüyordu."""
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(7, qty=5, status='TESLİM EDİLDİ', actual='2026-09-10',
                                   request='1111111', stock='AD-ESKI')]})
        eski = self.kart('MAKINE', 7)
        depo.kart_not_guncelle(eski['id'], 'Eski talebin notu', 'op', 'admin')
        yeni_satir = row(7, qty=5, status=None, start='2099-10-01', end='2099-10-05',
                         request='2222222', stock='AD-YENI')

        with self.snapshot_ile() as yukle:
            html = yukle(client, {'MAKİNE': [yeni_satir]})
        self.assertIn("NO'su başka talebe verilmiş kartlar ayrılacak", html)
        self.assertIn('2222222 · AD-YENI', html)

        sonuc = self.upload({'MAKİNE': [yeni_satir]})
        self.assertEqual((sonuc['ayrilan'], sonuc['yeni'], sonuc['guncellenen']), (1, 1, 0))
        yeni = self.kart('MAKINE', 7)
        self.assertNotEqual(yeni['id'], eski['id'])
        self.assertEqual((yeni['talep_no'], yeni['durum'], yeni['tamamlanan_adet'], yeni['aciklama'], yeni['operator']),
                         ('2222222', None, 0, None, None))
        ayrilan = depo.kart_getir(eski['id'])
        self.assertEqual((ayrilan['talep_no'], ayrilan['durum'], ayrilan['source_active'], ayrilan['source_row_id']),
                         ('1111111', 'TESLİM EDİLDİ', 0, f"NO:7~{eski['id']}"))
        self.assertIn('Eski talebin notu', ayrilan['aciklama'])
        self.assertNotIn(eski['id'], {k['id'] for k in depo.kartlari_getir()})   # operasyon ekranında değil
        # Aynı dosya tekrar: yeni ayrılma olmaz; diskten yeniden yükleme de tutarlı.
        self.assertEqual(self.upload({'MAKİNE': [yeni_satir]})['ayrilan'], 0)
        once = [dict(k) for k in depo._kartlar]
        depo._kartlar = []
        depo.kur()
        self.assertEqual([(k['id'], k['source_key'], k['source_active']) for k in depo._kartlar],
                         [(k['id'], k['source_key'], k['source_active']) for k in once])

    def test_ayni_no_yalniz_stok_degisirse_kart_yerinde_guncellenir(self):
        self.upload({'MAKİNE': [makine(1)]})
        kart_id = self.kart('MAKINE', 1)['id']
        sonuc = self.upload({'MAKİNE': [makine(1, stock='AD-7000-0099')]})
        self.assertEqual((sonuc['ayrilan'], sonuc['guncellenen']), (0, 1))
        self.assertEqual((self.kart('MAKINE', 1)['id'], self.kart('MAKINE', 1)['stok_no']), (kart_id, 'AD-7000-0099'))

    def test_taninmayan_durum_importu_durdurur(self):
        self.upload({'MAKİNE': [makine(1)]})
        once = [dict(k) for k in depo._kartlar]
        for durum in ['TESLİM EDİLDİ (KISMİ)', 'TESLIM', 'Bitti']:
            with self.subTest(durum=durum):
                with self.assertRaisesRegex(ex.ExcelAktarimHatasi, rf"satır 2: DURUM '{re.escape(durum)}' tanınmıyor"):
                    self.upload({'MAKİNE': [makine(1, status=durum)]})
                self.assertEqual(depo._kartlar, once)
        # Bilinen yazımlar ve bilerek durumsuz bırakılanlar kabul edilir.
        for durum, beklenen in [('teslim edildi', 'TESLİM EDİLDİ'), ('MALZEME TEDARİK', 'TESLİM EDİLDİ'),
                                ('Üretim Plana Alındı', 'PLANA ALINDI'), ('PDGM ÖNERİ', 'PLANA ALINDI')]:
            with self.subTest(durum=durum):
                self.upload({'MAKİNE': [makine(1, status=durum)]})
                self.assertEqual(self.kart('MAKINE', 1)['durum'], beklenen)

    # --- yardımcı ---
    def snapshot_ile(self):
        import contextlib, shutil
        from unittest.mock import patch

        @contextlib.contextmanager
        def baglam():
            def onizle(client, sayfalar):
                yol = self.book(sayfalar)
                def snapshot(kaynak):
                    hedef = self.root / 'onizleme.xlsx'; shutil.copy2(kaynak, hedef); return str(hedef)
                with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot), open(yol, 'rb') as f:
                    yanit = client.post('/yonetim/yukle', data={'_csrf_token': 'test-token', 'onizleme': '1',
                                                               'dosya': (f, 'plan.xlsx')})
                self.assertEqual(yanit.status_code, 200)
                return yanit.get_data(as_text=True)
            yield onizle
        return baglam()


if __name__ == '__main__':
    unittest.main()
