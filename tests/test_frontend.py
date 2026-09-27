"""Frontend incelemesi (outputs/frontend-inceleme-20260927/RAPOR.md) düzeltmelerinin sözleşmeleri.

Gerçek app modülü izole bir klasörden yüklenir (bkz. test_ui.UITests.load_app); üretim verisine dokunulmaz.
"""
import re
import unittest
from datetime import date, timedelta
from unittest.mock import patch

from flask import template_rendered

import depo
import excel_araclari as ex
from tests.test_excel_sync import row
from tests.test_ui import UITests

HEADERS = {'X-CSRF-Token': 'test-token'}
SATIR_ICI_BETIK = re.compile(r'<script(?![^>]*\bsrc=)[^>]*>', re.I)
SATIR_ICI_STIL = re.compile(r'<style\b|\sstyle="', re.I)
SATIR_ICI_OLAY = re.compile(r'\son[a-z]+="', re.I)


class FrontendTests(unittest.TestCase):
    setUp = UITests.setUp
    book = UITests.book
    upload = UITests.upload
    load_app = UITests.load_app

    def html(self, client, url):
        response = client.get(url)
        self.assertEqual(response.status_code, 200, url)
        return response.get_data(as_text=True)

    def context(self, module, client, url):
        yakalanan = []
        def yakala(sender, template, context, **extra): yakalanan.append(context)
        with template_rendered.connected_to(yakala, module.app):
            client.get(url)
        return yakalanan[-1]

    # --- #1: "Durum Ata" Malzeme Bekliyor işaretini silmemeli ---------------------------------
    def test_durum_ata_malzeme_bekliyor_isaretini_korur(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1, status='MALZEME TEDARİK', stock='MLZ')]})
        kart = depo.kartlari_getir(False)[0]
        self.assertFalse(kart['durum'])
        self.assertTrue(kart['malzeme_bekliyor'])

        html = self.html(client, '/yonetim')
        butonlar = re.findall(r'<button[^>]*data-admin-duzenle[^>]*>', html, re.S)
        self.assertEqual(len(butonlar), 2)  # "Durum Ata" + tablodaki "Düzenle"
        for buton in butonlar:
            self.assertIn('data-malzeme="1"', buton)
        surum = re.search(r'data-surum="([^"]+)"', butonlar[0]).group(1)

        # Kutuya dokunulmadığında arayüz null gönderir; sunucu mevcut değeri korumalı.
        yanit = client.post('/api/admin/duzenle', headers=HEADERS, json={
            'kart_id': kart['id'], 'surum': surum, 'durum': 'PLANA ALINDI', 'toplam_adet': 5,
            'tamamlanan_adet': 0, 'plan_hafta': '', 'plan_baslama': '2026-09-14', 'plan_teslim': '2026-09-20',
            'gerceklesen_teslim': '', 'not': '', 'dizgi_tipi': 'MAKİNE', 'dizgi_sorumlusu': '',
            'malzeme_bekliyor': None})
        self.assertEqual(yanit.status_code, 200, yanit.get_data(as_text=True))
        kart = depo.kartlari_getir(False)[0]
        self.assertEqual(kart['durum'], 'PLANA ALINDI')
        self.assertTrue(kart['malzeme_bekliyor'])

    # --- #22/CSP: şablonlarda satır içi betik, stil ve olay işleyici yok -------------------------
    def test_sayfalar_satir_ici_kod_icermez_ve_csp_gonderir(self):
        module, client = self.load_app()
        self.upload({'MAKİNE': [row(1, status='DİZGİDE'), row(2, status='MALZEME TEDARİK', stock='X2')]})
        sayfalar = {url: client.get(url) for url in ('/panel', '/operator', '/monitor', '/yonetim')}
        sayfalar['/giris'] = module.app.test_client().get('/giris')

        path = self.book({'MAKİNE': [row(1, status='DİZGİDE'), row(3, stock='YENI')]})
        def snapshot(_): return str(path)
        with patch.object(ex, 'excel_deger_snapshot_olustur', snapshot), path.open('rb') as f:
            sayfalar['onizleme'] = client.post('/yonetim/yukle', data={
                '_csrf_token': 'test-token', 'onizleme': '1', 'dosya': (f, 'plan.xlsx')})

        for ad, yanit in sayfalar.items():
            self.assertEqual(yanit.status_code, 200, ad)
            html = yanit.get_data(as_text=True)
            self.assertIsNone(SATIR_ICI_BETIK.search(html), ad)
            self.assertIsNone(SATIR_ICI_STIL.search(html), ad)
            self.assertIsNone(SATIR_ICI_OLAY.search(html), ad)
            csp = yanit.headers.get('Content-Security-Policy', '')
            self.assertIn("script-src 'self'", csp, ad)
            self.assertIn("style-src 'self'", csp, ad)
        self.assertIn('js/onizleme.js', sayfalar['onizleme'].get_data(as_text=True))

    # --- #23: statik dosyalar değişiklik zamanıyla sürümlenir --------------------------------
    def test_statik_adresler_otomatik_surumlu(self):
        module, client = self.load_app()
        for html in (self.html(client, '/panel'), module.app.test_client().get('/giris').get_data(as_text=True)):
            self.assertRegex(html, r'stil\.css\?v=\d+')
        self.assertRegex(self.html(client, '/panel'), r'js/ortak\.js\?v=\d+')

    # --- #6/#11: sürüm ucu değişikliği ve oturum düşmesini bildirir ---------------------------
    def test_api_surum_degisikligi_ve_oturumu_bildirir(self):
        module, client = self.load_app()
        self.upload({'MAKİNE': [row(1)]})
        ilk = client.get('/api/surum').get_json()['surum']
        self.assertEqual(client.get('/api/surum').get_json()['surum'], ilk)
        self.upload({'MAKİNE': [row(1, status='DİZGİDE')]})
        self.assertNotEqual(client.get('/api/surum').get_json()['surum'], ilk)
        self.assertIn(f'data-veri-surumu="{depo.depo_surumu()}"', self.html(client, '/panel'))
        self.assertEqual(module.app.test_client().get('/api/surum').status_code, 401)

    # --- #13: operatör ekranı eski teslimleri çizmez ------------------------------------------
    def test_operator_yalniz_son_teslimleri_cizer(self):
        module, client = self.load_app()
        eski = (date.today() - timedelta(days=60)).isoformat()
        yeni = (date.today() - timedelta(days=3)).isoformat()
        self.upload({'MAKİNE': [
            row(1, status='TESLİM EDİLDİ', start=eski, end=eski, actual=eski, stock='ESKI'),
            row(2, status='TESLİM EDİLDİ', start=yeni, end=yeni, actual=yeni, stock='YENI'),
            row(3, stock='PLAN')]})
        context = self.context(module, client, '/operator')
        self.assertEqual({k['stok_no'] for k in context['kartlar']}, {'YENI', 'PLAN'})
        self.assertEqual(context['eski_teslim'], 1)
        self.assertIn('daha eski 1 kart', self.html(client, '/operator'))

    # --- #11/#12: KPI iş emri sayar; dönem metrikleri sunucuda Türkçe biçimde -------------------
    def test_kpi_ve_donem_metrikleri(self):
        module, client = self.load_app()
        self.upload({'MAKİNE': [row(1, status='DİZGİDE', stock='AYNI'), row(2, status='DİZGİDE', stock='AYNI'),
                                row(3, status='TESLİM EDİLDİ', actual='2026-09-18')]})
        html = self.html(client, '/panel')
        kpi = re.search(r'<button[^>]*data-kpi-filtre="DİZGİDE"[^>]*>', html, re.S).group(0)
        self.assertIn('data-kart-hepsi="2"', kpi)
        self.assertIn('data-stok-hepsi="1"', kpi)
        is_emri = re.search(r'id="donem-metrik-is">([^<]*)<', html).group(1)
        self.assertEqual(is_emri, '1')

        self.assertEqual(module.sayi_filtresi(3.8), '3,8')
        self.assertEqual(module.sayi_filtresi(-1.25), '-1,2')
        self.assertEqual(module.sayi_filtresi(4.0), '4')
        self.assertEqual(module.sayi_filtresi(1234.5), '1.234,5')
        self.assertEqual(module.sayi_filtresi(None), '—')

    # --- #26: her bölümde aynı alanlar aranır (teslimlerde PCB dahil) ---------------------------
    def test_arama_alanlari_bolumler_arasinda_ayni(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1), row(2, status='TESLİM EDİLDİ', actual='2026-09-18', stock='T2')]})
        html = self.html(client, '/panel')
        for durum in ('PLANA ALINDI', 'TESLİM EDİLDİ'):
            satir = re.search(rf'data-durum="{durum}"\s+data-dizgi-tipi="[^"]*" data-arama="([^"]*)"', html).group(1)
            self.assertIn('PCB', satir, durum)
            self.assertIn('Owner', satir, durum)
        teslim = client.get('/api/panel/teslimler').get_json()['teslim_edilen'][0]
        self.assertEqual(teslim['pcb'], 'PCB')

    # --- #7: yükleme geçmişi iç kayıt önekini göstermez ---------------------------------------
    def test_yukleme_adi_oneki_atilir(self):
        module, _ = self.load_app()
        self.assertEqual(module.yukleme_adi_filtresi('20260926_214834_6ecbd7ee_PDGM_Plan.xlsx'), 'PDGM_Plan.xlsx')
        self.assertEqual(module.yukleme_adi_filtresi('rapor.xlsx'), 'rapor.xlsx')
        self.assertEqual(module.yukleme_adi_filtresi(None), '—')

    # --- #3/#14: operatör işlemleri isim alanı taşır, teslim kendi dialogunda -------------------
    def test_operator_dialoglari(self):
        _, client = self.load_app()
        self.upload({'MAKİNE': [row(1, status='DİZGİDE')]})
        html = self.html(client, '/operator')
        for onek in ('baslat', 'bitir', 'teslim', 'not'):
            self.assertRegex(html, rf'id="{onek}-isim"[^>]*data-islem-yapan')
        self.assertIn('id="teslim-dialog"', html)
        self.assertIn('Üretilen Adedi Gir', html)
        self.assertNotIn('Adet Bitir', html)

    # --- 27.09 ek istekler: logo, geciken rozeti yok, KPI başlıkları, açıklayıcı rozet ----------
    def test_logo_kpi_basliklari_ve_rozet_metni(self):
        module, client = self.load_app()
        yakin = (date.today() + timedelta(days=3)).isoformat()
        gecmis = (date.today() - timedelta(days=5)).isoformat()
        self.upload({'MAKİNE': [row(1, status='DİZGİDE', start=gecmis, end=yakin, stock='SUREDE'),
                                row(2, status='DİZGİDE', start=gecmis, end=gecmis, stock='GECIKEN')]})
        panel = self.html(client, '/panel')
        giris = module.app.test_client().get('/giris').get_data(as_text=True)
        for html in (panel, giris):
            self.assertRegex(html, r'<img[^>]+src="/static/logo\.png\?v=\d+"[^>]+alt="PDGM"')
            self.assertNotIn('marka-yazi', html)
        self.assertNotIn('geciken açık kart', panel)
        for baslik in ('Dizgideki İş Emri', 'Plana Alınan İş Emri', 'Teslim Edilen İş Emri'):
            self.assertIn(f'<span>{baslik}</span>', panel)
        self.assertIn('SÜRESİ İÇİNDE (teslime 3 gün kaldı)', panel)
        self.assertIn('>SÜRESİ İÇİNDE<', self.html(client, '/monitor'))

    # --- #15: giriş sayfası geçerli HTML (formlar iç içe değil) --------------------------------
    def test_giris_formlari_ic_ice_degil(self):
        module, _ = self.load_app()
        html = module.app.test_client().get('/giris').get_data(as_text=True)
        ilk_form_sonu = html.index('</form>')
        self.assertLess(ilk_form_sonu, html.index('giris-gozlemci-form'))
        govde = html[html.index('<form'):ilk_form_sonu]
        self.assertEqual(govde.count('<div'), govde.count('</div>'))


if __name__ == '__main__':
    unittest.main()
