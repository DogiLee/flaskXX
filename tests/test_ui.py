import importlib.util
import json
from pathlib import Path
import sys
from unittest.mock import patch
import unittest

from flask import template_rendered
import depo
from tests import test_excel_sync as fixtures
from tests.test_excel_sync import ROOT, row


class UITests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp
    book = fixtures.SyncTests.book
    upload = fixtures.SyncTests.upload

    def load_app(self):
        # Execute the actual app module from an isolated location: its startup
        # creates secrets/users/logs relative to __file__. No production data touched.
        path = self.root/'app_test.py'
        path.write_bytes((ROOT/'app.py').read_bytes())
        data = self.root/'data'; data.mkdir()
        (data/'kullanicilar.json').write_text(json.dumps({
            'tester': {'rol':'admin','ad':'Tester','aktif':True,'sifre_hash':'unused'}
        }),encoding='utf-8')
        spec = importlib.util.spec_from_file_location('pdgm_test_app',path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name]=module
        self.addCleanup(sys.modules.pop,spec.name,None)
        with patch.object(depo,'process_kilidi_al'):
            spec.loader.exec_module(module)
        module.app.template_folder=str(ROOT/'templates')
        module.app.static_folder=str(ROOT/'static')
        module.app.config['TESTING']=True
        for handler in list(module.app.logger.handlers):
            if getattr(handler,'baseFilename',None):
                self.addCleanup(module.app.logger.removeHandler,handler)
                self.addCleanup(handler.close)
        client=module.app.test_client()
        with client.session_transaction() as session:
            session.update(kullanici='tester',rol='admin',csrf_token='test-token')
        return module,client

    def test_monitor_dataset_and_badge_are_local(self):
        module,client=self.load_app()
        self.upload({'MAKİNE':[row(1,status='DİZGİDE',end='2099-09-20'),row(2,stock='MPLAN')],
                     'ELDE DİZGİ':[row(1,status='DİZGİDE',stock='HAND'),row(2,stock='HPLAN')],
                     'EÜM':[row(1,status='DİZGİDE',stock='EUM'),row(2,stock='EPLAN')]})
        contexts=[]
        def capture(sender,template,context,**extra): contexts.append(context)
        with template_rendered.connected_to(capture,module.app):
            response=client.get('/monitor')
        self.assertEqual(response.status_code,200)
        self.assertEqual({k['dizgi_kod'] for k in contexts[-1]['dizgide']+contexts[-1]['plana_alindi']},{'MAKINE'})
        html=response.get_data(as_text=True)
        self.assertIn('>SÜRESİ İÇİNDE<',html)
        panel=client.get('/panel').get_data(as_text=True)
        self.assertIn('SÜRESİ İÇİNDE (teslime',panel)
        self.assertIn('HAND',panel); self.assertIn('EUM',panel)
        # KPI: ana sayı iş emri, farklı stok sayısı alt bilgi (frontend incelemesi #11).
        self.assertIn('data-kpi-filtre="DİZGİDE"',panel)
        self.assertIn('farklı stok',panel)

    def test_upload_route_updates_api_and_persistence(self):
        module,client=self.load_app()
        import shutil
        import excel_araclari as ex
        def snapshot(source):
            target=self.root/'snapshot.xlsx'; shutil.copy2(source,target); return str(target)
        for status,end in [('PLANA ALINDI','2026-09-20'),('TESLİM EDİLDİ','2026-09-25')]:
            path=self.book({'MAKİNE':[row(1,status=status,end=end)]})
            with patch.object(ex,'excel_deger_snapshot_olustur',snapshot),path.open('rb') as f:
                response=client.post('/yonetim/yukle',data={'_csrf_token':'test-token','dosya':(f,'same.xlsx')})
            self.assertEqual(response.status_code,302)
            cards=client.get('/api/veriler').get_json()['kartlar']
            self.assertEqual(len(cards),1)
            self.assertEqual((cards[0]['durum'],cards[0]['plan_teslim']),(status,end))
        depo.kur()
        self.assertEqual(depo.kartlari_getir()[0]['durum'],'TESLİM EDİLDİ')

    def test_panel_delivery_api_excludes_deleted_source_rows(self):
        _,client=self.load_app()
        self.upload({'MAKİNE':[row(1,status='TESLİM EDİLDİ',actual='2026-09-18')]})
        self.upload({'MAKİNE':[]})
        response=client.get('/api/panel/teslimler')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.get_json()['teslim_edilen'],[])
        self.assertEqual(response.get_json()['ozet']['kart'],0)
        self.assertEqual(len(depo.kartlari_yonetim_getir()),1)


if __name__=='__main__': unittest.main()
