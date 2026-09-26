"""Verify exported fixtures using production parser/merge in temporary storage."""
from pathlib import Path
import copy, json, shutil, sys, tempfile, zipfile
from contextlib import ExitStack
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
import openpyxl
import depo
import excel_araclari as ex
OUT=ROOT/'outputs/pdgm-import-test-20260920'
manifest=json.loads((OUT/'manifest.json').read_text(encoding='utf-8'))
real_com='--com' in sys.argv

# Artifact export coerces ISO strings into UTC serials even under text format.
# Keep all workbook authoring in Artifact Tool; remove only our literal marker
# in the exported XML to preserve the exact timezone string as a text cell.
if not real_com:
    for p in OUT.glob('*.xlsx'):
        with zipfile.ZipFile(p) as z:
            parts=[(i,z.read(i.filename)) for i in z.infolist()]
        hits=sum(data.count(b'ISO_TEXT:') for i,data in parts if i.filename.endswith('.xml'))
        if hits:
            tmp=p.with_suffix('.tmp')
            with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
                for i,data in parts:
                    z.writestr(i,data.replace(b'ISO_TEXT:',b'') if i.filename.endswith('.xml') else data)
            tmp.replace(p)

messages=[]
def check(ok,msg):
    assert ok,msg
    messages.append(msg)
def card(no,sheet='MAKINE'):
    return next(k for k in depo._kartlar if k.get('source_sheet')==sheet and k.get('source_row_id')=='NO:'+str(no))
def active():return [k for k in depo._kartlar if k.get('source_active')]
with tempfile.TemporaryDirectory(prefix='pdgm-fixture-') as td, ExitStack() as stack:
    t=Path(td)
    vals={'VERI_KLASORU':str(t),'YEDEK_KLASORU':str(t/'backups'),
          'KARTLAR_DOSYA':str(t/'kartlar.xlsx'),'LOG_DOSYA':str(t/'log.xlsx'),
          'YUKLEME_DOSYA':str(t/'uploads.xlsx'),'_kartlar':[],'_loglar':[],'_yuklemeler':[]}
    for k,v in vals.items():stack.enter_context(patch.object(depo,k,v))
    depo.kur()
    def upload(name):
        p=OUT/name
        if real_com:return ex.excelden_aktar(str(p),'fixture-test')
        def snapshot(_):
            target=t/'snapshot.xlsx';shutil.copy2(p,target);return str(target)
        with patch.object(ex,'excel_deger_snapshot_olustur',snapshot):return ex.excelden_aktar(str(p),'fixture-test')
    base=manifest['cases'][0]['file']; phase2=manifest['cases'][1]['file']; phase3=manifest['cases'][2]['file']
    result=upload(base)
    check(len(depo._kartlar)==89,'01: 89 kaynak satırı import edildi.')
    ids={k['source_key']:k['id'] for k in depo._kartlar}
    check(card(1001)['toplam_adet']==5 and card(1001,'ELLE')['toplam_adet']==2,'MAKİNE/ELDE aynı talep ve stok bağımsız.')
    check(card('N001')['id']!=card('N1')['id'],'Harfli NO N001 ve N1 ayrı kimlikler.')
    for no,value in [(1001,'2026-09-14'),(3002,'2026-09-14'),(3003,'2026-09-21'),(3004,'2025-12-29'),(3005,'2026-09-21'),(3006,None),(3008,'2026-09-14')]:
        check(card(no,'EUM')['plan_baslama']==value,f'EÜM {no}: Pazartesi/boş tarih = {value}.')
    check(card(1008)['gerceklesen_teslim'] is None,'Tarihsiz teslimde tarih uydurulmadı.')
    for no in (1025,1026,1027):check(card(no)['toplam_adet']==1500,f'{no}: metinsel adet 1500.')
    check(card(1010)['durum'] is None and card(1011)['durum'] is None,'Boş/geçersiz durum yönetimde eksik.')
    result=upload(base)
    check(result['yeni']==0 and len(depo._kartlar)==89,'Aynı dosya tekrarında yeni kart yok.')
    k=card(1012)
    depo.kart_baslat(k['id'],10,'tester','admin','makine',aciklama='BAŞLANGIÇ TEST NOTU')
    depo.kart_bitir(k['id'],3,'tester','admin','makine',aciklama='KISMİ BİTİRME TEST NOTU')
    note=card(1012)['aciklama']
    result=upload(phase2)
    check(result['yeni']==1 and len(active())==89 and len(depo._kartlar)==90,'02: 1 yeni, 1 pasif; 89 aktif kaynak.')
    check(not card(1003)['source_active'],'02: yalnız silinen duplicate 1003 pasif.')
    check(all(k['id']==ids[k['source_key']] for k in depo._kartlar if k['source_key'] in ids),'Satır sıralaması değişince bütün eski ID’ler korundu.')
    check(card(1006)['plan_baslama'] is None and card(1006)['plan_teslim'] is None,'02: silinen plan tarihleri temizlendi.')
    check(card(1005)['plan_baslama']=='2026-09-22' and card(1005)['plan_teslim']=='2026-10-08','02: değişen plan tarihleri güncellendi.')
    check(card(1029)['talep_no']=='TEST-DUZELTILDI','02: aynı kimlikte talep/stok düzeltmesi işlendi.')
    check(card(1012)['tamamlanan_adet']==3 and card(1012)['aciklama']==note,'02: 3 tamamlanan adet ve manuel notlar korundu.')
    check(card(1007)['durum']=='DİZGİDE','02: Excel durumu DİZGİDE oldu.')
    upload(phase3)
    check(len(active())==90 and card(1003)['source_active'],'03: silinen satır aynı ID ile geri geldi.')
    check(card(1007)['durum']=='TESLİM EDİLDİ','03: Excel teslim geçişi uygulandı.')
    check(card(1023)['gerceklesen_teslim'] is None,'03: gerçekleşen teslim tarihi temizlendi.')
    for c in manifest['cases'][5:]:
        if c['kind']=='known_bug':continue
        before=copy.deepcopy((depo._kartlar,depo._loglar,depo._yuklemeler))
        files={p:p.read_bytes() for p in (t/'kartlar.xlsx',t/'log.xlsx',t/'uploads.xlsx')}
        try:upload(c['file'])
        except (ex.ExcelAktarimHatasi,depo.IsKuralHatasi) as err:
            check(before==(depo._kartlar,depo._loglar,depo._yuklemeler),c['file']+': hata sonrası bellek değişmedi.')
            check(all(p.read_bytes()==b for p,b in files.items()),c['file']+': hata sonrası kalıcı dosyalar değişmedi.')
            messages.append('  Ret nedeni: '+str(err))
        else:raise AssertionError('Beklenen ret gerçekleşmedi: '+c['file'])
    upload(manifest['cases'][3]['file'])
    check(len(active())==82 and all(not k['source_active'] for k in depo._kartlar if k['source_sheet']=='ELLE'),'04: yalnız 8 ELDE kartı pasifleşti.')
    upload(phase3)
    manual=depo.admin_kart_ekle('TEST-MANUEL','MANUEL-STOK',5,'tester')
    upload(manifest['cases'][4]['file'])
    check(len(active())==1 and active()[0]['id']==manual['id'],'05: tüm Excel kartları pasif, manuel kart aktif.')
    upload(phase3)
    before=copy.deepcopy(depo._kartlar);depo._kartlar=[];depo.kur()
    before_by_id={k['id']:k for k in before}
    differences=[(k['id'],field,before_by_id[k['id']].get(field),k.get(field)) for k in depo._kartlar for field in set(k)|set(before_by_id[k['id']]) if k.get(field)!=before_by_id[k['id']].get(field)]
    check(all(a in ('',None) and b in ('',None) for _,_,a,b in differences),'Yeniden okumadaki tek temsil farkı: boş metin ile boş hücre.')
    fields=['id','source_key','source_sheet','source_row_id','source_active','plan_baslama','plan_teslim','gerceklesen_teslim','aciklama','durum','tamamlanan_adet']
    check(all(k.get(f)==before_by_id[k['id']].get(f) for k in depo._kartlar for f in fields),'Yeniden okumada ID, tarih, not ve kaynak aktifliği korundu.')
    bug=next(c for c in manifest['cases'] if c['kind']=='known_bug')
    if real_com:
        snapshot=Path(ex.excel_deger_snapshot_olustur(str(OUT/bug['file'])))
        try:
            source=openpyxl.load_workbook(OUT/bug['file'],read_only=True,data_only=True)
            target=openpyxl.load_workbook(snapshot,read_only=True,data_only=True)
            for address in ['A35','A36']:
                a=source['MAKİNE'][address];b=target['MAKİNE'][address]
                messages.append(f'B01 kanıtı {address}: kaynak={a.value!r} ({a.data_type}, {a.number_format}); COM snapshot={b.value!r} ({b.data_type}, {b.number_format}).')
            check(source['MAKİNE']['A35'].value=='001' and target['MAKİNE']['A35'].value=='001','B01: baştaki sıfırlar COM snapshot sınırında korundu.')
            source.close();target.close()
        finally:snapshot.unlink(missing_ok=True)
        upload(bug['file'])
        check(card('001')['id']!=card('1')['id'],'B01: gerçek COM importunda 001 ve 1 ayrı kartlar.')
    else:
        upload(bug['file'])
        check(card('001')['id']!=card('1')['id'],'B01: COM olmadan parser 001 ve 1 kimliklerini ayrı tutuyor.')

report=('GERÇEK EXCEL COM + ' if real_com else 'SNAPSHOT SINIRI TAKLİT EDİLDİ + ')+'GERÇEK PARSER / MERGE / KALICI DEPO\n'
report+='Tüm işlemler geçici veri klasöründe yapıldı. Üretim verisine dokunulmadı.\n'
report+='Bu rapor import senaryolarını doğrular. UI/rol/monitor adımları TEST REHBERİ ile ayrıca denenmelidir.\n\n'
report+='\n'.join(messages)+'\n'
(OUT/('DOGRULAMA_COM.txt' if real_com else 'DOGRULAMA.txt')).write_text(report,encoding='utf-8')
print(report)
