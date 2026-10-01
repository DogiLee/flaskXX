"""Regression checks for the September system audit; temporary storage only."""
import copy
import unittest
from pathlib import Path
from unittest.mock import patch
import excel_araclari as ex
from flask import template_rendered
import depo
from tests import test_excel_sync as fixtures, test_ui as ui
from tests.test_excel_sync import row


class AuditFixTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    upload = fixtures.SyncTests.upload
    load_app = ui.UITests.load_app

    def prepare(self):
        self.mod, self.client = self.load_app()
        self.upload({'MAKİNE':[row(1,qty=10)]})
        return depo.kartlari_getir()[0]

    def post(self,path,data):
        return self.client.post(path,json=data,headers={'X-CSRF-Token':'test-token'})

    def test_stale_edit_cannot_replace_new_note_or_quantity(self):
        k=self.prepare()
        self.post('/api/basla',{'kart_id':k['id'],'adet':10})
        self.post('/api/bitir',{'kart_id':k['id'],'adet':3})
        old=depo.kartlari_getir()[0]
        self.post('/api/bitir',{'kart_id':k['id'],'adet':2,'not':'Yeni operatör notu'})
        response=self.post('/api/admin/duzenle',{'kart_id':k['id'],'surum':old.get('surum'),
            'durum':'DİZGİDE','toplam_adet':10,'tamamlanan_adet':3,'not':'Eski not'})
        self.assertEqual(response.status_code,409)
        after=depo.kartlari_getir()[0]
        self.assertEqual(after['tamamlanan_adet'],5)
        self.assertIn('Yeni operatör notu',after['aciklama'])

    def test_fresh_edit_accepted_and_missing_version_rejected(self):
        k=self.prepare()
        payload={'kart_id':k['id'],'durum':'PLANA ALINDI','not':'Yeni açıklama'}
        self.assertEqual(self.post('/api/admin/duzenle',payload).status_code,409)
        payload['surum']=k.get('surum')
        self.assertEqual(self.post('/api/admin/duzenle',payload).status_code,200)
        self.assertEqual(depo.kartlari_getir()[0]['aciklama'],'Yeni açıklama')

    def test_long_note_rejected_without_memory_or_disk_changes(self):
        k=self.prepare()
        self.post('/api/not',{'kart_id':k['id'],'not':'Korunacak not'})
        before=copy.deepcopy(depo._kartlar)
        disk=Path(depo.KARTLAR_DOSYA).read_bytes()
        response=self.post('/api/not',{'kart_id':k['id'],'not':'A'*33000})
        self.assertEqual(response.status_code,400)
        self.assertEqual(depo._kartlar,before)
        self.assertEqual(Path(depo.KARTLAR_DOSYA).read_bytes(),disk)

    def test_fractional_boolean_and_nonfinite_quantities_rejected(self):
        k=self.prepare()
        for value in [1.9,True,False,'1.9','NaN','Infinity']:
            with self.subTest(value=value):
                self.assertGreaterEqual(self.post('/api/basla',{'kart_id':k['id'],'adet':value}).status_code,400)
        self.assertEqual(depo.kartlari_getir()[0]['durum'],'PLANA ALINDI')

    def test_expired_api_session_returns_json_401_without_redirect(self):
        self.prepare()
        c=self.mod.app.test_client()
        for method,url in [('get','/api/veriler'),('post','/api/not')]:
            r=getattr(c,method)(url,json={})
            self.assertEqual(r.status_code,401)
            self.assertTrue(r.is_json)
            self.assertIsNone(r.location)

    def test_all_time_delivery_summary_includes_undated(self):
        self.mod,self.client=self.load_app()
        self.upload({'MAKİNE':[row(1,status='TESLİM EDİLDİ',actual=None),row(2,status='TESLİM EDİLDİ',actual='2026-09-18')]})
        r=self.client.get('/api/panel/teslimler').json
        self.assertEqual(r['ozet']['kart'],2)
        self.assertEqual(len(r['teslim_edilen']),2)

    def test_plan_filter_dataset_does_not_lose_cards_after_twelfth(self):
        self.mod,self.client=self.load_app()
        self.upload({'MAKİNE':[row(n) for n in range(1,15)],'ELDE DİZGİ':[row(1)]})
        contexts=[]
        def capture(sender,template,context,**kwargs): contexts.append(context)
        with template_rendered.connected_to(capture,self.mod.app): self.client.get('/panel')
        self.assertEqual(len(contexts[-1]['plana_alindi']),15)
        self.assertEqual(sum(k['dizgi_kod']=='ELLE' for k in contexts[-1]['plana_alindi']),1)

    def test_import_preview_does_not_write_and_stale_approval_is_rejected(self):
        k=self.prepare()
        self.post('/api/basla',{'kart_id':k['id'],'adet':10})
        self.post('/api/bitir',{'kart_id':k['id'],'adet':3})
        path=self.book({'MAKİNE':[row(1,qty=10)]})
        def snapshot(source):
            import shutil
            target=self.root/'preview.xlsx';shutil.copy2(source,target);return str(target)
        before=copy.deepcopy(depo._kartlar);disk=Path(depo.KARTLAR_DOSYA).read_bytes()
        with patch.object(ex,'excel_deger_snapshot_olustur',snapshot):
            with path.open('rb') as f:
                r=self.client.post('/yonetim/yukle',data={'_csrf_token':'test-token','onizleme':'1','dosya':(f,'plan.xlsx')})
            self.assertEqual(r.status_code,200)
            # Operatör kartı dizgiye alıp 3 adet girdi, Excel hâlâ PLANA ALINDI diyor:
            # önizleme kartı "Karar gerekiyor" bölümüne koyar, önerilen seçim uygulamadaki durum.
            html=r.get_data(as_text=True)
            self.assertIn('id="karar-gerekiyor"',html)
            self.assertRegex(html,rf'<select name="gerileme_{k["id"]}" form="onay-formu"')
            self.assertIn('<option value="DİZGİDE" selected>DİZGİDE · uygulamadaki (önerilen)</option>',html)
            self.assertEqual(depo._kartlar,before)
            self.assertEqual(Path(depo.KARTLAR_DOSYA).read_bytes(),disk)
            self.post('/api/not',{'kart_id':k['id'],'not':'Önizlemeden sonraki not'})
            r=self.client.post('/yonetim/yukle-onay',data={'_csrf_token':'test-token',
                'onizleme_token':ui.onizleme_token(self.client)})
            self.assertEqual(r.status_code,409)
            self.assertEqual(depo.kart_getir(k['id'])['tamamlanan_adet'],3)
            with path.open('rb') as f:
                r=self.client.post('/yonetim/yukle',data={'_csrf_token':'test-token','onizleme':'1','dosya':(f,'plan.xlsx')})
            self.assertEqual(r.status_code,200)
            # Seçim gönderilmezse aktarım uygulanmaz; admin Excel'e göre geri almayı seçerse adet 0 olur.
            self.assertEqual(self.client.post('/yonetim/yukle-onay',data={'_csrf_token':'test-token',
                'onizleme_token':ui.onizleme_token(self.client)}).status_code,409)
            with path.open('rb') as f:
                self.client.post('/yonetim/yukle',data={'_csrf_token':'test-token','onizleme':'1','dosya':(f,'plan.xlsx')})
            self.assertEqual(self.client.post('/yonetim/yukle-onay',data={'_csrf_token':'test-token',
                'onizleme_token':ui.onizleme_token(self.client),
                f'gerileme_{k["id"]}':'PLANA ALINDI'}).status_code,302)
            self.assertEqual(depo.kart_getir(k['id'])['tamamlanan_adet'],0)
            self.assertIn('Önizlemeden sonraki not',depo.kart_getir(k['id'])['aciklama'])
            self.assertEqual(self.client.post('/yonetim/yukle-onay',data={'_csrf_token':'test-token'}).status_code,409)

    def test_changed_calculated_snapshot_requires_new_preview(self):
        self.prepare()
        path=self.book({'MAKİNE':[row(1,qty=10)]})
        calls=0
        def snapshot(source):
            nonlocal calls
            import shutil, openpyxl
            calls+=1
            target=self.root/'calculated.xlsx';shutil.copy2(source,target)
            if calls==2:
                # Models a volatile Excel formula changing on the second COM open.
                wb=openpyxl.load_workbook(target);wb['MAKİNE']['F2']='2026-10-10';wb.save(target);wb.close()
            return str(target)
        before=copy.deepcopy(depo._kartlar)
        with patch.object(ex,'excel_deger_snapshot_olustur',snapshot):
            with path.open('rb') as f:
                self.client.post('/yonetim/yukle',data={'_csrf_token':'test-token','onizleme':'1','dosya':(f,'plan.xlsx')})
            r=self.client.post('/yonetim/yukle-onay',data={'_csrf_token':'test-token',
                'onizleme_token':ui.onizleme_token(self.client)})
        self.assertEqual(r.status_code,409)
        self.assertIn('hesaplanan değerleri önizlemeden sonra değişti',r.get_data(as_text=True))
        self.assertEqual(depo._kartlar,before)


if __name__=='__main__': unittest.main()
