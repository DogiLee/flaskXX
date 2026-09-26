import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const out = path.resolve('outputs/pdgm-import-test-20260920');
const previewDir = path.resolve('artifact_work/pdgm-import/previews');
await fs.mkdir(out,{recursive:true}); await fs.mkdir(previewDir,{recursive:true});
const reference='2026-09-20';
const day=n=>new Date(Date.UTC(2026,8,20+n)).toISOString().slice(0,10);
const dateCols=[6,7,8,9,10];
const headers=['NO','PDGM_ROW_ID','Talep NO','Talep Sahibi','Kart Stok No','Kart Üretim Adet',
 'Planlanan Başlangıç T.','Dizgi Başlama Tarihi','Planlanan Teslim T.','Gerçekleşen Teslim T.',
 'T.planlanan tarih','DURUM','PCB','PDGM Dizgi Sorumlusu','TEST KODU','BEKLENEN DAVRANIŞ'];
const r=(no,test,description,options={})=>({no,id:null,request:`TEST-${no}`,owner:'TEST VERİSİ',stock:`STK-${no}`,
 qty:5,week:'39. hafta',start:day(-2),due:day(7),actual:null,planned:null,status:'PLANA ALINDI',pcb:'VAR',responsible:'',test,description,...options});
const machine=[
 r(1001,'K01','5 MAKİNE dizgide; aynı talep/stoktaki 2 ELDE teslimden bağımsız.',{request:'TEST-ORTAK-7',stock:'ORTAK-ABC',qty:5,status:'DİZGİDE'}),
 r(1002,'K02','Üç duplicate: 2 adet planlı. Tarihi FAZ2’de değişir.',{request:'TEST-DUP-3',stock:'DUP-STOK',qty:2}),
 r(1003,'K03','Üç duplicate: 5 adet dizgide. FAZ2’de yalnız bu satır silinir.',{request:'TEST-DUP-3',stock:'DUP-STOK',qty:5,status:'DİZGİDE'}),
 r(1004,'K04','Üç duplicate: 8 adet teslim. Diğer iki satırdan bağımsız.',{request:'TEST-DUP-3',stock:'DUP-STOK',qty:8,status:'TESLİM EDİLDİ',actual:day(-1)}),
 r(1005,'D01','Plan tarihi değişir: başlangıç ve teslim aynı kartta güncellenir.'),
 r(1006,'D02','FAZ2’de iki plan tarihi boşalır; eski tarih görünmez.'),
 r(1007,'W01','FAZ1 planlı → FAZ2 dizgide → FAZ3 teslim. ID sabit.',{start:day(1)}),
 r(1008,'W02','Teslim tarihi bilinmiyor. Admin yalnız not değiştirince tarih üretilmez.',{status:'TESLİM EDİLDİ',actual:null}),
 r(1009,'W03','HAZIR yalnız yönetimde; Pano/monitor/operasyonda görünmez.',{status:'HAZIR'}),
 r(1010,'W04','Boş DURUM: ilk importta yönetimde durumu eksik.',{status:null}),
 r(1011,'W05','Geçersiz DURUM: ilk importta yönetimde durumu eksik.',{status:'BEKLEMEDE'}),
 r(1012,'W06','Manuel kısmi üretim için: 10 adet başlat, 3 bitir, not ekle; FAZ2 durum boş.',{qty:10}),
 r(1013,'W07','Malzeme bekliyor; admin PLANA ALINDI yapınca operatör malzeme onayı denenir.',{status:'MALZEME TEDARİK',pcb:'YOK'}),
 r(1014,'R01','Referans gününe göre süresi aşılmış açık iş.',{status:'DİZGİDE',start:day(-10),due:day(-3)}),
 r(1015,'R02','Referans gününde SON GÜN rozeti.',{status:'DİZGİDE',due:day(0)}),
 r(1016,'R03','Referans gününden bir gün sonra SON 1 GÜN.',{status:'DİZGİDE',due:day(1)}),
 r(1017,'R04','Pano PLANINDA; monitor PLANDA rozeti.',{status:'DİZGİDE',due:day(15)}),
 r(1018,'R05','Başlangıcı geçmiş planlı iş: BAŞLAMADI.',{start:day(-4)}),
 r(1019,'R06','Referans gününde BUGÜN BAŞLAMALI.',{start:day(0)}),
 r(1020,'R07','Plan tarihsiz açık iş; UI boş tarihi güvenle göstermeli.',{status:'DİZGİDE',start:null,due:null,week:null}),
 r(1021,'R08','Geç teslim; pozitif sapma.',{status:'TESLİM EDİLDİ',start:day(-15),due:day(-5),actual:day(-2)}),
 r(1022,'R09','Zamanında/erken teslim; negatif sapma.',{status:'TESLİM EDİLDİ',start:day(-15),due:day(-2),actual:day(-4)}),
 r(1023,'P01','Teslim tarihi FAZ2’de değişir, FAZ3’te boşalır.',{status:'TESLİM EDİLDİ',actual:day(-1)}),
 r(1024,'A01','Sayısal 400.0 kabul edilir.',{qty:400}),
 r(1025,'A02','Metinsel 1.500 ADET → 1500.',{qty:'1.500 ADET'}),
 r(1026,'A03','Metinsel 1,500 ADET → 1500.',{qty:'1,500 ADET'}),
 r(1027,'A04','Metinsel 1 500 ADET → 1500.',{qty:'1 500 ADET'}),
 r(1028,'T01','Türkçe/boşluk normalizasyonu → DİZGİDE.',{status:'  dizgide  '}),
 r(1029,'K05','Talep/Stok FAZ2’de düzeltilir; NO sabit olduğundan ID korunur.'),
 r(1030,'U01','HTML görünümlü sahip metni ekranda düz metin olmalı.',{owner:'<script>alert("test")</script> Çığ Örnek & "Kart"'}),
 r('N001','K06','Harfli NO=N001 ile NO=N1 iki ayrı kaynak kimliğidir.',{request:'TEST-NO-METIN',stock:'NO-METIN',qty:1}),
 r('N1','K07','Harfli NO=N1, NO=N001 ile birleşmemeli.',{request:'TEST-NO-METIN',stock:'NO-METIN',qty:2}),
 r(null,'K08','NO boş, sabit PDGM_ROW_ID dolu: kabul edilir.',{id:'test-row-uuid-A',request:'TEST-ID',stock:'ID-STOK'}),
 r(1031,'K09','İki kimlik varsa PDGM_ROW_ID öncelikli.',{id:'test-explicit-B'}),
];
for(let i=0;i<25;i++) machine.push(r(1101+i,`M${i+1}`,'Monitor dizgide grubunda çok sayfalı rotasyon.',{status:'DİZGİDE',qty:i+1,due:day(10+i%4)}));
for(let i=0;i<14;i++) machine.push(r(1201+i,`Q${i+1}`,'Monitor planlı grubunda çok sayfalı rotasyon.',{start:day(3+i%3),due:day(20),qty:i+2}));
const hand=[
 r(1001,'E01','2 ELDE teslim; aynı talep/stoktaki 5 MAKİNE dizgideden bağımsız.',{request:'TEST-ORTAK-7',stock:'ORTAK-ABC',qty:2,status:'TESLİM EDİLDİ',actual:day(-1)}),
 r(2002,'E02','DİZGİ İÇİN BEKLİYOR → PLANA ALINDI.',{status:'DİZGİ İÇİN BEKLİYOR',responsible:'Test Elle Operatörü'}),
 r(2003,'E03','DİZGİDE VE MALZEME BEKLİYOR → DİZGİDE + malzeme rozeti.',{status:'DİZGİDE VE MALZEME BEKLİYOR',pcb:'EKSİK'}),
 r(2004,'E04','MALZEME TEDARİK: durumu eksik + malzeme bekliyor.',{status:'MALZEME TEDARİK',pcb:'YOK'}),
 r(2005,'E05','Elle operatörü kısmi üretim / teslim denemesi.',{status:'DİZGİDE',qty:12}),
 r(2006,'E06','Elle backlog HAZIR yalnız yönetimde.',{status:'HAZIR'}),
 r(2007,'E07','Elle planlı; makine operatörü değiştirememeli.'),
 r(2008,'E08','Elle tarihi bilinmeyen teslim.',{status:'TESLİM EDİLDİ'}),
];
const eum=[
 r(1001,'EU01','17.09.2026 Perşembe → başlangıç 14.09.2026 Pazartesi.',{status:'ÜRETİM DEVAM EDİYOR',start:null,planned:'2026-09-17',responsible:'Test EÜM'}),
 r(3002,'EU02','20.09.2026 Pazar → başlangıç 14.09.2026.',{status:'ÜRETİM PLANA ALINDI',start:null,planned:'2026-09-20'}),
 r(3003,'EU03','21.09.2026 Pazartesi → aynı gün.',{status:'EMTD PLANLANDI',start:null,planned:'2026-09-21'}),
 r(3004,'EU04','01.01.2026 → 29.12.2025, yıl geçişi.',{status:'ÜRETİM PLANA ALINDI',start:null,planned:'2026-01-01'}),
 r(3005,'EU05','ISO +03:00 tarih: takvim günü kaymaz, başlangıç 21.09.2026.',{status:'ÜRETİM DEVAM EDİYOR',start:null,planned:'2026-09-21T00:30:00+03:00'}),
 r(3006,'EU06','Planlanan tarih boş: başlangıç boş.',{start:null,planned:null,status:'PDGM ÖNERİ'}),
 r(3007,'EU07','EÜM tamamlanmış üretim.',{status:'TESLİM EDİLDİ',start:null,planned:'2026-09-14',actual:day(-1)}),
 r(3008,'EU08','Metinsel Türkçe tarih 17.09.2026 → 14.09.2026.',{status:'ÜRETİM DEVAM EDİYOR',start:null,planned:'17.09.2026'}),
];
const base={'MAKİNE':machine,'ELDE DİZGİ':hand,'EÜM':eum};
const clone=x=>structuredClone(x);
const find=(s,no)=>s['MAKİNE'].find(r=>r.no===no);
const phase2=clone(base);
Object.assign(find(phase2,1002),{due:day(15)});
phase2['MAKİNE']=phase2['MAKİNE'].filter(x=>x.no!==1003);
Object.assign(find(phase2,1005),{start:day(2),due:day(18)});
Object.assign(find(phase2,1006),{start:null,due:null});
Object.assign(find(phase2,1007),{status:'DİZGİDE'});
Object.assign(find(phase2,1012),{status:null,pcb:'GÜNCELLENDİ',due:day(25)});
Object.assign(find(phase2,1023),{actual:day(0)});
Object.assign(find(phase2,1029),{request:'TEST-DUZELTILDI',stock:'DUZELTILEN-STOK'});
phase2['MAKİNE'].push(r(1900,'Y01','FAZ2’de eklenen tek yeni kart.',{start:day(2)}));
phase2['EÜM'][0].planned='2026-09-24';
for(const rows of Object.values(phase2))rows.reverse();
const phase3=clone(phase2);
Object.assign(find(phase3,1007),{status:'TESLİM EDİLDİ',actual:day(0)});
Object.assign(find(phase3,1023),{actual:null});
phase3['MAKİNE'].push(clone(base['MAKİNE'].find(x=>x.no===1003)));
const cases=[
 {file:'01_ILK_IMPORT.xlsx',title:'İlk import',sheets:base,expect:'Kabul. İlk yükleme ve aynı dosyayı tekrar yükleme.',kind:'base'},
 {file:'02_YENIDEN_IMPORT.xlsx',title:'Tarih, durum ve satır değişiklikleri',sheets:phase2,expect:'Kabul. 1 yeni, 1 pasif; diğer kaynak kimlikleri korunur.',kind:'phase2'},
 {file:'03_TESLIM_VE_GERI_GELEN.xlsx',title:'Teslim ve geri gelen kaynak satırı',sheets:phase3,expect:'Kabul. 1007 teslim; 1003 eski ID ile geri gelir; 1023 teslim tarihi boşalır.',kind:'phase3'},
 {file:'04_ELDE_SAYFASI_YOK.xlsx',title:'Opsiyonel kaynak sayfasının kaldırılması',sheets:{'MAKİNE':clone(phase3['MAKİNE']),'EÜM':clone(phase3['EÜM'])},expect:'Kabul. Yalnız ELDE kaynakları pasifleşir. 03 dosyasıyla geri getirilebilir.',kind:'missing_optional'},
 {file:'05_TUM_KAYNAKLAR_BOS.xlsx',title:'Tüm kaynak satırlarının silinmesi',sheets:{'MAKİNE':[],'ELDE DİZGİ':[],'EÜM':[]},expect:'Kabul. Tüm Excel kartları pasif; manuel kartlar ve audit kayıtları korunur.',kind:'empty'},
];
function reject(file,title,mutate,expected){const sheets=clone(base);const c={file,title,sheets,expect:expected,kind:'reject'};mutate(sheets,c);cases.push(c);}
reject('H01_TEKRAR_NO.xlsx','Aynı sayfada tekrarlanan NO',s=>s['MAKİNE'][1].no=1001,'Reddedilir. Tekrarlanan kaynak kimliği.');
reject('H02_EKSIK_KIMLIK.xlsx','NO ve PDGM_ROW_ID boş',s=>s['MAKİNE'][0].no=null,'Reddedilir. Kalıcı kimlik gerekli.');
reject('H03_EKSIK_STOK.xlsx','Kimliği olan satırda stok boş',s=>s['MAKİNE'][0].stock=null,'Reddedilir. Satır silinmiş sayılmamalı.');
reject('H04_EKSIK_TALEP.xlsx','Kimliği olan satırda talep boş',s=>s['MAKİNE'][0].request=null,'Reddedilir. Satır silinmiş sayılmamalı.');
reject('H05_SIFIR_ADET.xlsx','Üretim adedi sıfır',s=>s['MAKİNE'][0].qty=0,'Reddedilir. Adet en az 1.');
reject('H06_KESIRLI_ADET.xlsx','Üretim adedi kesirli',s=>s['MAKİNE'][0].qty=2.5,'Reddedilir. Adet tam sayı.');
reject('H07_BELIRSIZ_ADET.xlsx','İki sayılı adet metni',s=>s['MAKİNE'][0].qty='400 / 500 ADET','Reddedilir. Adet tek sayı olmalı.');
reject('H08_GECERSIZ_TARIH.xlsx','Geçersiz takvim tarihi',s=>s['MAKİNE'][0].due='31.02.2026','Reddedilir. Geçersiz tarih.');
reject('H09_TERS_TARIHLER.xlsx','Başlangıç teslimden sonra',s=>{s['MAKİNE'][0].start=day(15);s['MAKİNE'][0].due=day(1);},'Reddedilir. Başlangıç teslimden sonra olamaz.');
reject('H10_MAKINE_YOK.xlsx','Zorunlu MAKİNE sayfası yok',s=>delete s['MAKİNE'],'Reddedilir. MAKİNE sayfası gerekli.');
reject('H11_TEKRAR_BASLIK.xlsx','Aynı alana eşlenen iki başlık',(s,c)=>c.duplicateHeader=true,'Reddedilir. İkinci Adet başlığı aynı alana eşlenir.');
reject('H12_ADET_KOLONU_YOK.xlsx','Zorunlu adet başlığı yok',(s,c)=>c.missingQty=true,'Reddedilir. Kart Üretim Adet sütunu gerekli.');
const conflict=clone(base);Object.assign(find(conflict,1012),{qty:2,status:null});
cases.push({file:'H13_MANUEL_ADET_CELISKISI.xlsx',title:'Manuel kısmi üretimle miktar çelişkisi',sheets:conflict,expect:'Önkoşul: 1012 kartında en az 3 tamamlanan adet. Toplam=2, DURUM boş; tüm import reddedilir.',kind:'conditional_reject'});
const comBug=clone(base);comBug['MAKİNE'].find(x=>x.no==='N001').no='001';comBug['MAKİNE'].find(x=>x.no==='N1').no='1';
for(const x of comBug['MAKİNE'].filter(x=>['001','1'].includes(x.no)))x.description='Bilinen COM hatası: metin 001 ve 1 sayısal 1 olur; import tekrar kimlik hatasıyla reddedilir.';
cases.push({file:'B01_BASTAKI_SIFIR_COM_HATASI.xlsx',title:'Bilinen hata: COM metinsel kimliği dönüştürüyor',sheets:comBug,expect:'İstenen: 001 ve 1 ayrı kartlar. Mevcut gerçek COM: NO:1 tekrarı ile ret. Parser tek başına kabul eder.',kind:'known_bug'});

const manual=[
 ['G01','Kurulum','Ayrı test kopyası','Test verisini üretimde yüklemeyin. Import tam kaynak listesidir; mevcut Excel kartları pasifleşir.'],
 ['G02','İlk yükleme','01 → 01','İkinci yüklemede yeni kart=0. Satır ID ve kaynak kimlikleri değişmez.'],
 ['G03','Manuel hazırlık','MAKİNE NO=1012','FAZ2 öncesi 10 adet başlatın, 3 adet bitirin, iki ayrı not ekleyin. FAZ2 DURUM boş olduğundan ilerleme/notlar korunmalı.'],
 ['G04','Senkronizasyon','01 → 02 → 03','Her adımda ID listesini indirin. 02 sıraları ters çevirir, 1003 silinir, 1900 eklenir; 03 ile 1003 aynı ID ile geri gelir.'],
 ['G05','Hatalı import','H01–H12','Her dosyayı ayrı yükleyin. Hata mesajı gelmeli; kartlar, adetler ve yükleme geçmişi değişmemeli.'],
 ['G06','Adet çelişkisi','H13 / NO=1012','Önce G03 hazırlığı gerekir. Tamamlanan≥3 iken toplam=2 importu bütünüyle iptal etmeli. Hazırlıksız denemeyin.'],
 ['G07','Kapsam silme','04 ardından 03','04 ELDE sayfasını kaldırır; sadece ELDE pasifleşir. 03 tekrar aktifleştirir.'],
 ['G08','Tam silme','05, en son','Bütün Excel kartları operasyonlardan kaybolur; yönetim/audit ve manuel kartlar korunur. 03 ile geri dönün.'],
 ['O01','Operatör akışı','MAKİNE NO=1012','Başlat → 3 bitir → kalan bitir → teslim. Tamamı bitmeden teslim ve toplamdan fazla adet engellenmeli. FAZ2 öncesi yalnız 3 bitirin.'],
 ['O02','Elle yetkisi','ELDE NO=2005','Elle operatörü işlem yapabilmeli; makine ve EÜM operatörü yapamamalı.'],
 ['O03','EÜM yetkisi','EÜM NO=1001','EÜM operatörü işlem yapabilmeli; diğer tipler yapamamalı.'],
 ['O04','Gözlemci','Pano ve Monitor','Gözlemci Pano/Monitor görür; yönetim ve operatör aksiyonlarına erişemez.'],
 ['O05','Malzeme onayı','MAKİNE NO=1013','Admin durumunu PLANA ALINDI yapın. Operatörde malzeme uyarısı/onayı; admin geçişi loglanmalı.'],
 ['O06','Pasif kart notu','MAKİNE NO=1003','01’de kartı açık tutun, 02 yükleyin. Eski operatör sekmesinden işlem/not reddedilmeli. Admin audit notu ekleyebilir.'],
 ['A01','Yönetim ekleme','Yeni manuel kart','TEST-MANUEL / MANUEL-STOK ekleyin. 05 importunda manuel kart pasifleşmemeli.'],
 ['A02','Gizle/geri getir','MAKİNE NO=1005','Admin gizleme ve geri getirme deneyin. Gizliyken yeniden import gizlilik flagini korumalı.'],
 ['A03','Eksik durum düzelt','MAKİNE NO=1010/1011','Admin geçerli durum atayabilmeli. Excel boş/geçersiz kalırsa sonraki import bu workflow’u korumalı.'],
 ['A04','Tarihsiz teslim notu','MAKİNE NO=1008','Admin yalnız notu değiştirip kaydedince bilinmeyen tarihler bugüne dönüşmemeli.'],
 ['A05','Yedek dönüş','Yönetim → yedekler','03 sonrası oluşan yedeği not edin; 05 sonrası geri yükleyin. Kart/adet/notların geri geldiğini kontrol edin.'],
 ['A06','Restart / yeniden oku','03 sonrası','Sunucuyu yeniden başlatın veya kartlar.xlsx yeniden oku kullanın. ID, kaynak aktifliği ve tarihler aynı kalmalı.'],
 ['A07','Rapor ve dosyalar','Rapor / kayıt indirme','Kartlar, loglar ve yüklemeler dosyalarını indirin. Rapor dizgi tipi, not, tarih ve kaynak aktifliğini doğru göstermeli.'],
 ['U01','Monitor filtresi','/monitor','DİZGİDE ve PLANA ALINDI bölümlerinde yalnız MAKİNE. ELDE/EÜM yalnız Pano/operator/yönetimde.'],
 ['U02','Slider','/monitor','12 saniye geçiş; 4/6 kartlık sayfalar; son sayfalar da görünmeli. Yenilemede konum korunur.'],
 ['U03','Rozet farkı','MAKİNE NO=1017','Pano PLANINDA (x gün var); monitor PLANDA. Gecikme ve son gün uyarıları korunur.'],
 ['U04','Arama / filtre','Pano ve operatör','Talep, stok, sahip; durum ve dizgi tipi filtreleri. Temizle düğmesi varsayılana döner.'],
 ['U05','Unique stok sayacı','TEST-DUP-3 / ORTAK-ABC','Aynı durumdaki stoklar set olarak sayılır; adet veya satır sayısı değildir.'],
 ['U06','Teslim dönemleri','Pano teslimler','Tümü, hafta, ay, yıl, özel aralık ve dizgi filtresi. Tarihi bilinmeyen teslim Tümü listesinde; tarih aralıklarına dahil değil.'],
 ['U07','Metin kaçışı','MAKİNE NO=1030','Sahip alanındaki script benzeri metin düz yazı olmalı; JavaScript çalışmamalı.'],
 ['X01','Gizli NO','01 dosyasının kopyası','NO sütununu üç kaynak sayfasında gizleyin. Yeniden import aynı kartları bulmalı.'],
 ['X02','Gizli satır','01 dosyasının kopyası','Bir veri satırını gizleyin. Gizli satır yine import edilmeli; silinmiş sayılmamalı.'],
 ['X03','Gizli veri kolonu','01 dosyasının kopyası','PCB kolonunu gizleyin. Import görünür kolonları kullanır; PCB boşalması beklenir. Zorunlu stok kolonunu gizlemek hata verir.'],
 ['X04','Başlık toleransı','EÜM plan başlığı','T.planlanan tarih başlığını Planlanan Başlangıç T. yapmadan önce mevcut Planlanan Başlangıç T. kolonunu kaldırın; tek tarih kolonu ile Monday fallback deneyin.'],
 ['X05','Kimlik sözleşmesi','Tüm NO değerleri','NO değerlerini yeniden numaralandırmayın. Kaynak dosya adı/satır sırası kimlik değildir. NO değişirse yeni kart oluşması beklenir.'],
 ['X06','Kimlik alternatifi','K08/K09','PDGM_ROW_ID olan satırın NO değerini değiştirin; açık ID sabitse aynı entity kalmalı.'],
 ['X07','Tarih sistemleri','Ayrı kopya / gelişmiş','1904 tarih sistemi ve formül cache davranışı ayrıca otomatik COM testinde kapsanır. Rastgele Excel ayarı değiştirmek tarihleri kaydırabilir.'],
 ['S01','Oturum','Giriş/çıkış','Her mevcut rol hesabıyla giriş/çıkış; gözlemci girişi; izinsiz URL erişimi. Parolalar bu pakette bulunmaz.'],
 ['S02','İş kuralı hataları','Operatör formları','0, negatif, kalan adetten fazla ve tam bitmeden teslim: kayıt değişmemeli.'],
 ['S03','CSRF / dosya sınırı','Gelişmiş manuel test','Eksik CSRF ile POST reddedilir. .xlsx/.xlsm dışı yükleme reddedilir; 25 MB sınırı ayrıca denenir.'],
 ['S04','Excel kilidi / rollback','Ayrı test ortamı','kartlar.xlsx kilitliyken importun hata vermesi ve yarım state bırakmaması kontrol edilir. Kalıcı OS/elektrik kesintisi için yedek gerekir.'],
 ['S05','Migration','Eski test kartlar.xlsx','Eski format için otomatik regression testlerini kullanın. Kaynak import Excel’i tek başına eski persistence formatı oluşturmaz.'],
 ['B01','Bilinen COM hatası','B01 dosyası','001 ve 1 metinsel kimlikleri COM snapshotta sayısal 1 olur; mevcut uygulama tekrar kimlik hatası verir. Bu davranış düzeltme gerektirir.'],
];

function sheetStyle(sheet,lastCol,lastRow,widths){
 sheet.showGridLines=false;const all=sheet.getRange(`A1:${lastCol}${lastRow}`);
 all.format.font={name:'Arial',size:10,color:'#243447'};
 all.format.verticalAlignment='center';all.format.rowHeight=30;
 for(let i=0;i<widths.length;i++) sheet.getRangeByIndexes(0,i,lastRow,1).format.columnWidth=widths[i];
 sheet.freezePanes.freezeRows(4);
 sheet.getRange(`A4:${lastCol}4`).format={fill:'#263D54',font:{name:'Arial',bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:44};
 sheet.getRange(`A5:${lastCol}${Math.max(5,lastRow)}`).format.wrapText=true;
 for(let n=5;n<=lastRow;n+=2)sheet.getRange(`A${n}:${lastCol}${n}`).format.fill='#F1F5F9';
}
async function make(c,index){
 const wb=Workbook.create();
 const guide=wb.worksheets.add('TEST REHBERİ');
 const guideRows=[...manual];
 guide.getRange('A1').values=[['PDGM import testleri']];
 guide.getRange('A2').values=[[`${c.file} — ${c.expect}`]];
 guide.getRange('A3').values=[[`Referans gün: ${reference}. Gün bazlı rozetler bilgisayarın bugünün tarihine göre değişir. Tüm kayıtlar sentetiktir.`]];
 guide.getRange('A4:D4').values=[['Kod','Test alanı','Hedef / sıra','İşlem ve beklenen sonuç']];
 guide.getRange(`A5:D${guideRows.length+4}`).values=guideRows;
 sheetStyle(guide,'D',guideRows.length+4,[10,26,35,102]);
 guide.getRange('A1:D3').format.rowHeight=25;guide.getRange('A1').format.font={bold:true,size:16};
 guide.getRange(`A5:D${guideRows.length+4}`).format.rowHeight=52;
 guide.getRange('A2').format.font={color:'#9A3412',bold:true};
 for(const [name,rows] of Object.entries(c.sheets)){
   const sheet=wb.worksheets.add(name);
   const h=[...headers];if(c.missingQty&&name==='MAKİNE')h[5]='Üretim Bilgisi';
   if(c.duplicateHeader&&name==='MAKİNE')h[15]='Adet';
   sheet.getRange('A1').values=[[`PDGM test verisi: ${name}`]];
   sheet.getRange('A2').values=[['Sabit NO / PDGM_ROW_ID korunmalı. TEST KODU ve BEKLENEN DAVRANIŞ uygulamaya aktarılmaz.']];
   sheet.getRange('A3').values=[[`Referans gün ${reference}. ${c.title}.`]];
   sheet.getRange('A4:P4').values=[h];
   const matrix=rows.map(o=>[o.no,o.id,o.request,o.owner,o.stock,o.qty,o.week,o.start,o.due,o.actual,o.planned,o.status,o.pcb,o.responsible,o.test,o.description]);
   if(matrix.length){
     for(const vals of matrix)for(const k of dateCols)if(typeof vals[k]==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(vals[k]))vals[k]=new Date(vals[k]+'T12:00:00Z');
     sheet.getRange(`A5:P${rows.length+4}`).values=matrix;
   }
   sheetStyle(sheet,'P',Math.max(5,rows.length+4),[11,25,23,29,24,19,23,23,23,23,25,30,18,26,13,72]);
   sheet.getRange('A1:P3').format.rowHeight=24;sheet.getRange('A1').format.font={bold:true,size:15};
   sheet.getRange(`A5:E${Math.max(5,rows.length+4)}`).setNumberFormat('@');
   sheet.getRange(`F5:F${Math.max(5,rows.length+4)}`).setNumberFormat('#,##0');
   sheet.getRange(`H5:K${Math.max(5,rows.length+4)}`).setNumberFormat('yyyy-mm-dd');
   for(let n=0;n<rows.length;n++){
     if(typeof rows[n].qty==='number'&&!Number.isInteger(rows[n].qty))sheet.getRange(`F${n+5}`).setNumberFormat('0.0');
     // Export cannot retain timezone ISO strings; preserve a non-date token here.
     // The narrowly scoped package repair replaces it with literal ISO text.
     if(typeof rows[n].planned==='string'&&rows[n].planned.includes('T00:')){
       sheet.getRange(`K${n+5}`).values=[['ISO_TEXT:'+rows[n].planned]];
       sheet.getRange(`K${n+5}`).setNumberFormat('@');
     }
   }
   if(rows.length)sheet.tables.add(`A4:P${rows.length+4}`,true,`Data${index}_${Object.keys(c.sheets).indexOf(name)}`);
 }
 const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:5},maxChars:600});
 if(errors.ndjson.includes('"cell"'))throw Error('Unexpected formula error '+c.file+' '+errors.ndjson);
 if(index===0){
   for(const [name,range,label] of [['TEST REHBERİ','A1:D10','guide'],['MAKİNE','A4:F11','machine'],['ELDE DİZGİ','H4:P9','hand'],['EÜM','H4:P12','eum']]){
     const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});
     await fs.writeFile(path.join(previewDir,label+'.png'),new Uint8Array(await blob.arrayBuffer()));
   }
   console.log((await wb.inspect({kind:'table',range:"'MAKİNE'!A4:F8",include:'values',tableMaxRows:5,tableMaxCols:6,maxChars:1800})).ndjson);
 }
 if(index>0&&index<3){const blob=await wb.render({sheetName:'MAKİNE',range:'H4:L10',scale:1,format:'png'});await fs.writeFile(path.join(previewDir,`phase${index+1}.png`),new Uint8Array(await blob.arrayBuffer()));}
 const xlsx=await SpreadsheetFile.exportXlsx(wb);await xlsx.save(path.join(out,c.file));
 console.log(c.file, Object.values(c.sheets).reduce((sum,rows)=>sum+rows.length,0),'rows');
}
for(let i=0;i<cases.length;i++)await make(cases[i],i);
await fs.writeFile(path.join(out,'manifest.json'),JSON.stringify({reference,headers,cases,manual},null,2));
await fs.writeFile(path.join(out,'BASLAMADAN_ONCE.txt'),`PDGM TEST PAKETİ\nReferans gün: ${reference}\n\nYALNIZ AYRI TEST KOPYASINDA KULLANIN. Import tam kaynak listesi kabul edilir.\nBu dosyalarda bulunmayan mevcut Excel kartları pasifleşir.\n\nSıra: 01'i yükle, aynı dosyayı tekrar yükle. 1012 kartında 10 başlat / 3 bitir / not ekle.\n02'yi yükle, 03'ü yükle. H01–H12 ayrı ayrı reddedilmeli.\nH13 için 1012 kartında 3 tamamlanan adet önkoşulu gerekir.\n04 ELDE sayfasını kaldırır; 05 bütün kaynakları boşaltır. Bunları en son dene.\n03'ü yeniden yükleyerek kaynak kartlarını geri getirebilirsin.\n\nHer Excel'in TEST REHBERİ sekmesinde rol, operatör, yönetim, monitor ve hata testleri var.\nHer kaynak satırının O/P kolonları test kodunu ve beklenen davranışı açıklar.\nBunlar uygulama not alanına aktarılmaz.\nGün bazlı rozetler bilgisayarın gerçek tarihine göre değişir.\nParola/gerçek kullanıcı/üretim verisi bulunmaz.\n\nDOSYALAR\n${cases.map(c=>c.file+'\n  '+c.expect).join('\n')}\n`);
console.log('Created',cases.length,'workbooks');
