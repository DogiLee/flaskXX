# Excel senkronizasyonu — kullanım, uyumluluk ve doğrulama

## Kimlik modeli

Her satır `source_sheet` + `source_row_id` ile tanınır. Sayfa kodları MAKINE,
ELLE ve EUM'dur. `NO=123` için satır kimliği `NO:123` olur; alternatif
`PDGM_ROW_ID=abc` kullanılırsa `ID:abc` olur. Her ikisi varsa PDGM_ROW_ID önceliklidir.
`source_key` ve `anahtar`, örneğin `EXCEL:v2:["MAKINE","NO:123"]` biçimindedir.
JSON kodlaması ayraç içeren kimlikleri de birbirinden ayırır.

Kullanıcı NO'nun sayfa içinde benzersiz ve kalıcı olduğunu doğruladı. NO/PDGM_ROW_ID
sonraki yüklemelerde değişmemeli, silinen bir satırın kimliği başka bir işte tekrar
kullanılmamalıdır. Tarihler, miktar, durum, Talep/Stok, dosya adı ve fiziksel sıra
kimliğe katılmaz. Aynı NO farklı sayfalarda bağımsızdır. Sayfa değişikliği yeni
kaynak satırıdır; eski satır pasifleşir. Gerçek satır transferi ayrıca eşleme gerektirir.

NO olmayan kaynaklar için bir kez kalıcı NO veya PDGM_ROW_ID kolonunu doldurup
aynı değerleri kaynak Excel'de saklayın. Her upload sırasında yeni UUID üretmek
kalıcılık sağlamaz. Eksik/tekrarlanan kimlik tüm importu iptal eder. 26.09.2026'dan
beri bütün gizli sütunlar snapshot'a alınır (bkz. son bölüm); gizlilik yalnız aynı
başlığın birden fazla sütunda geçtiği durumda görünür olanı seçmek için kullanılır.

## Excel sync rules

22.09.2026 güncellemesi: Yönetim yükleme formu önce **Etkiyi Önizle** adımını
gösterir. Gerçek merge mantığı bellek kopyasında çalışır; kart/log/yükleme
dosyaları ve yedekler bu adımda yazılmaz. Önizleme onaylanınca kaynak dosyasının,
hesaplanan Excel değerlerinin ve mevcut kartların değişmediği kontrol edilir.
Değişmişse yeni önizleme istenir. Not ekleme gibi önizleme sonrası operasyonlar
da bu kontrol kapsamındadır. Programatik `excelden_aktar` çağrıları ve eski
doğrudan yükleme POST sözleşmesi korunmuştur; arayüz önizleme kullanır.

Yönetim düzenleme API'si artık kart görünümündeki `surum` değerini gerektirir.
Eski form yeni not/adedi ezemez; 409 yanıtında formdaki metin korunur ve güncel
kayıt kontrolü istenir. Sürüm, alan içeriklerinden hesaplanır; aynı saniyedeki
değişiklikler de fark edilir. Yeni kalıcı kolon gerekmez.

Not ve diğer Excel metinleri için 32.767 karakter sınırı yazmadan doğrulanır.
Sınır aşılırsa işlem reddedilir ve bellek/disk geri alınır; metin kırpılmaz.
Bu sınır tarih/kişi etiketi dahil **toplam not geçmişine** uygulanır. Önceden
kırpılmış geçmişi kendiliğinden geri getirmez. Tamamlanan adet API'leri kesirli
ve boolean değerleri kabul etmez. COM snapshot'ta NO ve PDGM_ROW_ID sütunları
metin biçiminde taşınır; metinsel `001` ve `1` ayrı kalır.

| Alan | Yeni upload kuralı |
|---|---|
| Kaynak kimliği | Sabit; kaynak sayfası + NO/PDGM_ROW_ID. Aynı sayfa+NO farklı Talep NO taşırsa eski kart geçmişiyle ayrılır (pasif, kimlik `NO:x~kartID`), yeni talep için temiz kart açılır. NO daha sonra eski talebine dönerse (ör. yazım hatası düzeltildi) ayrılmış kart notları ve iş akışıyla geri bağlanır; birden fazla aday varsa tahmin yapılmaz (27.09.2026) |
| talep_no, stok_no, sira, talep_sahibi | Aynı kaynak kimliğinde Excel'den güncellenir |
| toplam_adet, adet_metin, plan_hafta, plan_baslama, plan_teslim | Excel authoritative; boş tarih eski tarihi temizler |
| gerceklesen_teslim | Boşaltma dahil Excel authoritative; bugünün tarihi uydurulmaz. İstisna (27.09.2026): kart uygulamada teslim edildiyse (teslim zamanı var), TESLİM EDİLDİ kalıyorsa ve Excel'in tarih hücresi boşsa uygulamadaki tarih korunur; boş hücre bilinen tarihi silmez. Excel bir tarih yazarsa o geçerlidir |
| excel_durum, pcb, dizgi_tipi, dizgi_sorumlusu, malzeme_bekliyor | Kaynaktan güncellenir |
| Geçerli DURUM | Excel authoritative; TESLİM EDİLDİ tarihi eksik olsa da kabul edilir, uyarı sayılır. İstisna (27.09.2026): önizlemeden onaylanan importta Excel'in durumu uygulamadakinin gerisindeyse (HAZIR < PLANA < DİZGİDE < TESLİM) kart "Karar gerekiyor" bölümüne düşer; admin Excel ile uygulama durumu arasından seçer. Öneri: kart uygulamada ilerletildiyse (başlama/teslim zamanı var) uygulamadaki durum, değilse Excel. Uygulamadaki TESLİM korunursa teslim tarihi de korunur. Önizlemesiz doğrudan import eskisi gibi Excel'i uygular |
| Boş DURUM veya MALZEME TEDARİK / PDGM ÖNERİ | Yeni kartın durumu boş kalır. Uygulamada durumu olan kart (27.09.2026): önizlemeden onaylanan importta "Karar gerekiyor" bölümüne düşer; seçenekler "Durumsuz bırak" (durum boş, tamamlanan 0, başlama/bitiş/teslim zamanı silinir, kart Durumu Eksik listesine düşer) ve uygulamadaki durumu korumak. Öneri: durum uygulamada verildiyse (başlama/teslim zamanı var veya kartta saklı önceki Excel DURUM metni bu duruma karşılık gelmiyor, ör. admin Durum Ata) koru, Excel'den geldiyse durumsuz bırak. Karar "EXCEL DURUM BOŞ: DURUM KARARI" olarak loglanır; korunan kartın operatör alanına "Excel" yazılmaz. Önizlemesiz doğrudan import eskisi gibi workflow/adet korur |
| Tanınmayan DURUM yazımı | Tüm import durdurulur; satır ve kabul edilen değerler listelenir |
| Tamamlanan adet | Plan/HAZIR=0; teslim=toplam; DİZGİDE kısmi üretimi korur. Önizlemede, bu importta güncellenen ve DİZGİDE kalan (tamamlanan > 0) kartlar için admin "Tamamlanan adedi sıfırla" seçebilir (varsayılan işaretsiz; sıfırlama bitiş zamanını da temizler, not/başlama korunur; önizlemede sunulmayan kart seçimi importu iptal eder) |
| Çelişkili miktar | DİZGİDE veya workflow korunan kartta toplam, tamamlanandan küçükse tüm import iptal edilir |
| operator, aciklama, admin_gizli | Korunur; boş operatör geçerli Excel durumunda Excel olabilir. Notlar (aciklama) yalnız admin önizlemede "Notları temizle" işaretlerse silinir (27.09.2026): seçenek, bu importta durumu değişen veya durum kararı istenen, notu olan kartlarda çıkar; silinen metin "EXCEL: NOTLAR TEMİZLENDİ" loguna yazılır; önizlemede sunulmayan kart seçimi importu iptal eder |
| Workflow zaman damgaları | Gerçek mevcut kayıtlar korunur veya durum gerilemesinde temizlenir; import zamanı olay tarihi gibi yazılmaz |
| Teslim zaman damgası | Kaynak teslim tarihi değişirse/boşalırsa temizlenir; eski tarih UI'a geri sızmaz (yukarıdaki uygulamada teslim istisnasında tarih ile birlikte korunur) |
| Kaynakta olmayan Excel kartı | source_active=0; tarih/not/ID fiziksel olarak saklanır |
| Manuel kartlar | Excel eşlemesine ve source pasifleştirmesine katılmaz |

Import tam kaynak snapshot'ıdır: zorunlu MAKİNE sayfası ve bulunan opsiyonel
sayfalar birlikte işlenir. Opsiyonel bir sayfanın kaldırılması o kaynağın satırlarını
pasifleştirir. Başlıkları geçerli, bütün kaynak sayfaları boş bir dosya tüm Excel
kartlarını pasifleştirir. Talep NO / Kart Stok No'su eksik ama adet, DURUM,
plan/teslim tarihi, PDGM_ROW_ID veya Talep NO içeren satırlar silinmiş kabul
edilmez; import reddedilir. Yalnız NO/Stok/sorumlu/hafta metni gibi kalıntı
hücreleri içeren satırlar kart sayılmaz ve önizlemede listelenir (26.09.2026).

`workflow_korundu` artık yalnız DURUM boş/geçersiz olduğu için workflow'u korunan
eşleşmeleri sayar. Alan değişiklikleri kart ID'si ve kaynak anahtarıyla, eski/yeni
değerleri içeren audit kaydına yazılır. Pasifleştirme de kart bazında loglanır.

## EÜM ve ekranlar

EÜM'de `T.planlanan tarih`, yoksa `Planlanan Başlangıç T.` kolonunun tarihi
parser'da Pazartesi'ye indirilir: 17.09.2026 → 14.09.2026. Alan mevcut ama boşsa
başlangıç da boştur. Alanların hiçbiri yoksa mevcut açık başlangıç tarihi kullanılır.
Sayısal tarihler workbook'un 1900/1904 epoch'una göre çözülür. ISO saat dilimli
tarihlerde kaynak takvim günü korunur; UTC dönüşümü ile gün kaydırılmaz.

Monitor backend yalnız `dizgi_kod == "MAKINE"` gönderir. "SÜRESİ İÇİNDE (teslime N gün kaldı)"
rozeti (eski adı "PLANINDA (N gün var)") monitor'da kısaca "SÜRESİ İÇİNDE" olur; gecikme ve
son gün uyarıları kalır. Pano/operator
ayrıntılı rozetleri değişmez. Slider 12 saniyedir. Yenileme en az 60 saniyede,
en uzun grubun tüm sayfalarını göstermesine yetecek süre sonunda yapılır; sayfa
konumu sessionStorage'da tutulur. Çok sayfalı monitor'da yenileme 60 saniyeyi aşabilir.

Pano/operator stok etiketleri benzersiz stok sayısını açıkça belirtir. Kaynakta
olmayan kartlar ortak operasyon filtresinde, işlem API'lerinde ve Pano teslimler
API'sinde dışlanır. Yönetim, yedekler ve tarihsel raporlar geçmişi korur.

## Migration / compatibility ve rollback

Eski kartlar.xlsx açılmaya devam eder: dört yeni kolon zorunlu eski kolonlar listesine
eklenmedi. Yeni kolonlar Kaynak Sayfa, Kaynak Satır ID, Kaynak Anahtar, Eski Anahtar.
İlk importta yalnız EXCEL kartları eşlenir:

1. Aynı kaynak tipi + Talep/Stok grubunda sabit NO, eski sayısal Sıra ile tekil
   eşleşiyorsa ID, operatör, not ve gizlilik korunur.
2. Eski Sıra boşsa yalnız eski ve gelen grubun her ikisi de tek satır olduğunda
   eşleme yapılır. Birden çok aday varsa tarihe/miktara bakarak tahmin yapılmaz.
3. Eski Sıra dönüşümünün kaybettiği `001`/`1` gibi ayrımlar otomatik taşınmaz;
   açıklayıcı hata verilir. Yeni modelde bu iki metinsel NO bağımsızdır.
4. Eski anahtar `legacy_anahtar` alanında tutulur. Modern source_key bulunan
   kartlar başka bir kaynak anahtarına otomatik taşınmaz.

Belirsizlikte kaynak Excel ile eski kartların ID/Sıra/dizgi tipi eşleşmesi,
**üretim dosyasının kopyasında** insan tarafından doğrulanmalıdır. Eksik Sıra
alanları doğru NO ile eşlenebilir. Metinsel kimliklerin kayıpsız taşınması için
dört yeni kolonu ve Anahtar'ı doğrulanmış eşlemeye göre doldurmak gerekir;
`depo.kaynak_anahtari(sayfa_kodu, satir_kimligi)` anahtarın doğru biçimini üretir.
Önceki bug nedeniyle zaten başka satırın workflow'unu taşıyan kayıtlar otomatik
olarak güvenilir biçimde onarılamaz; yedek/kaynak kayıtlarla kontrol gerekir.

Her import öncesi kart/log/yüklemeler benzersiz yedek klasörüne kopyalanır; aynı
saniyedeki importlar artık aynı yedeği ezmez. Normal yazma exception'larında hem
bellek hem disk geri alınır. `_import_kilidi` ve depo `_kilit` korunmuştur.

Geri dönüş: sunucuyu durdurun, birlikte alınmış import öncesi kartlar/log/yüklemeler
dosyalarını geri koyun ve gerekiyorsa eski uygulama kodunu geri yükleyin. Sadece
eski kodu yeni verinin üstüne çalıştırmayın: yeni sürüm tarihi bilinmeyen teslimleri
kabul eder, eski validator bunları reddeder. Bu çalışma üretim verisi içermediği
için gerçek kartlar üzerinde migration çalıştırılmadı.

Kodun önceki hali `.investigation/baseline/`, birleşik diff'i
`.investigation/changes.diff` altındadır. Git deposu bulunmadığından commit yoktur.

## Regression tests

Çalıştırma (Windows, proje kökü):

```powershell
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
$env:PDGM_TEST_EXCEL_COM='1'
.\.venv\Scripts\python.exe -B -m unittest tests.test_excel_com -v
node tests/monitor_rotation.cjs
```

Paketler mevcut requirements.txt sürümleriyle yerel `.venv` içine kuruldu.
Testler TemporaryDirectory içinde XLSX üretir; gerçek üretim data klasörünü
okumaz/yazmaz. COM testi isteğe bağlıdır ve kurulu Microsoft Excel gerektirir.

İlk 17 test eski kodda 18 başarısız alt-kontrol ve bir hata gösterdi; 3 test
zaten geçiyordu. Son durumda 39 otomatik Python testi geçti, varsayılan koşuda
COM testi atlandı. COM testi ayrıca gerçek Excel ile geçti. Node slider testi
de geçti. Kapsam: tekrar upload, tarihler, üç duplicate, sıra/ekleme/silme,
MAKİNE–ELDE–EÜM ayrımı, 2+5 senaryosu, durum geçişleri, source pasifleştirme,
disk reload ve ayrı Python process restart, legacy migration, belirsizlikte
iptal, manuel kartların korunması, gerçek Flask upload ve UI/API modelleri,
eşzamanlı commit, ikinci dosya replace hatasında bellek/disk rollback,
timezone/1904/hafta geçişleri, gizli NO ve gerçek COM reupload.

Bağımsız kod incelemesindeki üç bulgu (eksik satırın silinmiş sayılması,
eski teslim zamanının geri gelmesi, başında sıfır olan legacy NO) önce başarısız
testle doğrulandı ve düzeltildi. Son kontrol ayrıca Pano teslim API filtresini düzeltti.

## Remaining risks

- Gerçek üretim Excel'i/kartlar.xlsx sağlanmadı; üretimdeki kolon ve legacy eşleme
  dağılımı doğrulanmadı. NO'nun kalıcılığı kullanıcı beyanına dayanır.
- COM formül sonuçlarını okur; dış bağlantılar ve tam yeniden hesaplama bilerek
  çalıştırılmaz. Kaynak dosyada formül cache'i eskimişse önce Excel'de güncelleyin.
- XLSX çok dosyalı gerçek bir veritabanı transaction'ı değildir. Süreç/power loss
  veya rollback yazmasının da OS tarafından engellenmesi durumunda yedekten
  kurtarma gerekir. Normal exception rollback'i test edildi; crash-journal eklenmedi.
- Kimlik yeniden numaralandırılır/reuse edilirse sistem aynı fiziksel işi kendi
  başına anlayamaz. Belirsiz eski eşleşmeler otomatik düzeltilmez.


## Tekrar analiz — 20.09.2026

Kullanıcı isteğiyle ikinci kez analiz ve bağımsız inceleme yapıldı. İki ek hata,
önce başarısız testle doğrulanıp düzeltildi:
- Aynı durumdaki kartın yalnız notunu düzenleyen admin, bilinmeyen teslim/üretim
  tarihlerine bugünü yazmıyor. Gerçek durum geçişlerindeki tarih ataması korunuyor.
- Kaynakta pasif kartlara operatör not yazamıyor. Admin tarihsel audit notu ekleyebilir.

Son varsayılan test koşusu 40 test: 39 başarılı, 1 isteğe bağlı COM testi atlandı.
COM testi gerçek Microsoft Excel ile ayrıca tekrar çalıştırıldı. Slider ve Python
sözdizimi kontrolleri de başarılı. Son koşu kaydı `.investigation/reanalysis-tests.txt`.

## Sistem bulgularının düzeltilmesi — 22.09.2026

Son doğrulama gerçek Excel COM dahil 50/50 regresyon testi ve 24/24 sistem
testidir. Güncel rapor `outputs/audit-fixes-20260922/DUZELTME_RAPORU.md`.
Önceki tarihli test sayıları ve bilinen hata açıklamaları tarihsel kayıttır.

Yönetim importu artık etki önizlemesi sunar. Uygulamada depo sürümü, kaynak
dosya hash'i ve parse edilmiş satır hash'i yeniden kontrol edilir; değişiklik
varsa import uygulanmaz. Programatik/önizlemesiz eski import çağrıları korunur.
Geçerli Excel DURUM'u ve belgelenmiş adet politikası yetkili olmaya devam eder.
COM kimlik sütunları değer yazılmadan önce metin biçimine alınır; 001 ve 1
gerçek Excel aktarımında ayrı kimlik olarak doğrulandı.

Admin düzenleme API'si güncel kartın `surum` değerini zorunlu tutar; eski form
not veya üretim adedini ezemez. XLSX depolama şeması değiştirilmedi. Not/metin
32.767 karakteri aşıyorsa işlem açık hatayla reddedilir ve eski kayıt korunur;
sınırsız not depolama veya geçmişte kesilen metnin kurtarılması sağlanmaz.


## Gerçek Excel import düzeltmesi — 26.09.2026

Kaynak: proje kökündeki `PDGM_Kart_dizgi_Talepleri_Üretim_Takvimi.xlsx`
(MAKİNE başlık satırı 4, ELDE DİZGİ ve EÜM başlık satırı 2; dosyada formül yok).

- **MAKİNE satır 211 hatası:** satır gizli; yalnız `I: Kart Stok No = AD-0000-0001`
  ve gizli `K: Birim Seviyesi Kullanım` dolu, NO/Talep NO/adet/DURUM/tarih yok.
  20.09 düzeltmesi Talep/Stok/NO'dan biri dolu her eksik satırı reddediyordu; dosyada
  52 MAKİNE + 4 EÜM böyle gizli kalıntı satırı var. Yeni sınıflandırma: tamamen boş
  satır sessizce atlanır; kalıntı satırı kart sayılmaz ve önizleme/bildirim/logda satır
  numarasıyla listelenir; adet, DURUM, plan/teslim tarihi, PDGM_ROW_ID veya Talep NO
  içeren eksik satır hata verir. Satırın gizli olması sınıfı değiştirmez. Satır
  hataları toplanır, tek mesajda raporlanır ve depo'ya hiçbir şey gönderilmez.
- **Gizli sütunlar:** snapshot artık tüm sütunları aynı konumda ve gizlilik
  bilgisiyle taşır. Tekil başlıklı sütun gizli olsa da okunur (önceden gizlenen DURUM
  veya teslim sütunu tüm kartlarda boş sayılırdı). Aynı başlık birden fazla sütunda
  varsa (gerçek dosyada gizli sipariş bloğundaki ikinci `Planlanan Teslim T.`)
  tek görünür olan seçilir; hepsi görünür ya da hepsi gizliyse açık hata verilir.
  EÜM'de gizli `PDGM Dizgi Sorumlusu` artık okunur.
- **EÜM plan başlangıcı:** gerçek sayfada `Planlanan Başlangıç T.` hafta metnidir
  (`23. Hafta (2.06 haftası)`). Tarih gibi okunmaya çalışıldığı için import bir
  sonraki adımda EÜM satır 3'te duracaktı. Hafta metni Pazartesi'ye çözülür; yıl,
  hafta numarası/gün.ay tutarlılığı ve satırın plan/teslim tarihine yakınlıkla
  seçilir. Açık Dizgi Başlama tarihi varsa o önceliklidir. Okunamayan hafta metni
  uyarıdır; `T.planlanan tarih` sütunundaki geçersiz değer hata olmaya devam eder.
- **DURUM:** gerçek ELDE yazımı `DİZGİDE VE MALZEME BEKLENİYOR` eşleşmiyordu; artık
  DİZGİDE + malzeme bekliyor. `PLANA ALINDI` ve `ÜRETİM PLANA ALINDI` (büyük/küçük
  harf, fazla boşluk, ayrışık İ dahil) PLANA ALINDI olur; planlanan kart tamamlanmış
  sayılmaz. `MALZEME TEDARİK` ve `PDGM ÖNERİ` mevcut kurala göre durum boş kalır.
- **COM değer kopyası:** Value2 ile yazılan metinleri Excel yeniden yorumluyordu
  (`001234` → 1234, `1-2` → tarih, `=abc` → formül). Metinler `'` önekiyle yazılır;
  hücre hataları (#N/A vb.) sayı kodu yerine hata metni olarak taşınır ve kayıtta
  hataya, kalıntı satırda atlanmaya yol açar.
- **Tekrar yükleme:** alanı değişmeyen kartın `guncelleme` damgası ve sürümü
  korunur, `degismeyen` olarak sayılır, audit logu yazılmaz. Boş kaynak metni bellekte
  de `None` tutulur; sunucu yeniden başladıktan sonra hayali "" → None farkı oluşmaz.
- **Önizleme ekranı (aynı gün, ikinci adım):** özet satırı ve alan alan uzun tablo yerine
  etki kutuları (yeni / güncellenecek / değişmeyecek / pasifleşecek / uyarı), sayfa bazında
  okunan-atlanan-boş satır kartları, türüne göre gruplanmış uyarılar (ne anlama geldiği ve
  hangi satırlar), aynı hücreleri dolu kalıntı satırlarının desen özeti, kart bazında
  "Şu an / Onaydan sonra" karşılaştırması, pasifleşecek/yeni kart tabloları ve alta
  yapışık onay çubuğu. Aktif Excel kartlarının en az %30'u pasifleşecekse veya bir sayfa
  eksikse kırmızı uyarı çıkar, buton "Yine de uygula" olur. Önizlemede Excel reddedilirse
  Yönetim'e kesilmiş bir bildirimle dönülmez; sorunlar "sayfa · satır" etiketiyle madde
  madde gösterilir (HTTP 422). Stil dosyası sürümü `?v=20260926`.
- **Sahada görülen belirtiler (ters tür, yanlış tarih, teslim edilmiş kartın "PLANDA"
  görünmesi):** eski kök nedenler (tarihten türeyen anahtar, sayfalar arası ortak anahtar,
  tarihsiz TESLİM'in düşürülmesi) güncel kodda tekrar etmiyor; `tests/test_onceki_hatalar.py`
  5 MAKİNE + 2 ELDE ve aynı Talep+Stok çoklu satır senaryolarını ekran/API düzeyinde doğrular.
  Aynı belirtileri hâlâ üreten üç yol kapatıldı: NO'nun başka talebe verilmesi (eski kartın
  durumu/notu yeni talebe geçiyordu → kart ayrılır), tanınmayan DURUM (kart eski durumda
  kalıyordu → import durur), yinelenen "Planlanan Teslim T." sütununun görünürlüğe göre
  seçilmesi (yanlış bloktan tarih → üretim sütunları DURUM'a yakınlıkla seçilir; NO gibi
  kimlik sütunlarında belirsizlik hâlâ hatadır).
- **Rapor özeti:** kaynakta olmayan kartlar pano/teslim API'sinde olduğu gibi rapor
  sayaçlarına girmez; rapordaki kart listesi onları `Kaynakta Aktif=0` ile tutar.

Gerçek dosya sonucu: MAKİNE 15, ELDE DİZGİ 20, EÜM 9 kayıt (toplam 44) → 44 kart;
operasyon ekranlarında 35 (9 kartın kaynak DURUM'u boş/eşlenmeyen). Testler:
`tests/test_gercek_excel.py` (gerçek dosya ve kopyaları; COM sınıfı
`PDGM_TEST_EXCEL_COM=1` ile). Önceki kod `.investigation/import-fix-20260926-before/`,
birleşik diff `.investigation/import-fix-20260926.diff`. Rapor:
`outputs/import-duzeltme-20260926/RAPOR.md`.


## Yüksek öncelikli düzeltmeler — 27.09.2026

- **Kayıt dosyası yazım hızı:** `kartlar.xlsx` ve `islem_logu.xlsx` her işlemde baştan
  yazılır. Satır numarası artık sayaçla verilir (`ws.max_row` her satırda tüm hücreleri
  tarıyordu, yazım karesel büyüyordu) ve gövde hücrelerine tek tek font atanmaz. Log sınırı
  20.000'den 5.000'e indi; aşılınca en yeni 2.000 kayıt kalır, eskiler `yedekler/` altında
  `…_islem_logu_arsiv.xlsx` dosyasına taşınır. Arşiv yalnız girişte değil, log ekleyen her
  commit'ten sonra çalışır; hata verirse işlem bozulmaz, sonraki commit'te yeniden denenir.
  "NOT EKLENDİ" logu tüm not geçmişi yerine yalnız eklenen satırı yazar. Ölçüm (1.000 kart):
  2.000 logla not ekleme 2,55 → 1,03 sn, giriş 0,97 → 0,36 sn; 20.000 satırlık bir log ilk
  işlemde bir kez arşivlenir, sonra işlemler bu düzeyde kalır. Rapor indirmede yalnız
  sayaç düzeltmesi yapıldı, biçim aynı.
- **Teslim tarihi:** yukarıdaki tablodaki gerceklesen_teslim istisnası. Excel TESLİM
  EDİLDİ'ye yetişip tarih hücresini boş bıraktığında uygulamada kaydedilen teslim tarihi
  ve zamanı silinmez (önceden geç teslimler sapma ve dönem sayılarından düşüyordu).
- **Önizleme–onay bağlama:** her önizleme bir token üretir; onay formu bu token'ı gönderir.
  Aynı oturumda başka bir sekmede başka dosya önizlendiyse eski sayfanın onayı 409 ile
  reddedilir ve hiçbir kayıt değişmez; son önizleme tüketilmez, kendi sekmesinden
  onaylanabilir. Yükleme ve onay formları `data-tek-gonderim` ile çift gönderime karşı
  kilitlenir (ortak.js).

Testler: `tests/test_yuksek_oncelik.py` (12 test; 6'sı değişiklik öncesi kodda başarısız).
Python değişiklikleri çalışan sunucuya yeniden başlatmadan yansımaz.


## Veri kaybı ve izlenebilirlik düzeltmeleri — 27.09.2026

- **Karar önerisi:** Excel'den DİZGİDE gelmiş kartta operatörün uygulamada girdiği adet
  (tamamlanan > 0 ve son işlemi yapan "Excel" değil) "uygulamada ilerletildi" sayılır;
  önerilen seçim uygulamadaki durumu korur. Eski bir Excel TESLİM'inden kalan adet sayılmaz.
- **Talep NO geri dönüşü:** yukarıdaki kaynak kimliği satırı. Önizlemede kart "Excel'e geri
  eklendiği için yeniden aktifleşecek" grubunda "NO eski talebine döndü" etiketiyle çıkar;
  "EXCEL NO ESKİ TALEBE DÖNDÜ" loglanır.
- **İşlemi yapan:** operatör hesapları paylaşımlı olduğu için dizgiye alma, adet girişi,
  teslim, not ve malzeme onayında girilen kişi adı kartın operatör alanına ve işlem
  logunun yeni **İşlemi Yapan** sütununa yazılır (hesap Kullanıcı sütununda kalır). Eski log
  dosyaları bu sütun olmadan okunur. Arayüz alanı hesabın adıyla doldurmaz ve hesap adıyla
  gönderimi reddeder; ad gönderilmeyen API çağrılarında hesap adı kullanılır (en fazla 80
  karakter). "Üretilen adet" alanı boş açılır.
- **Admin denetim izi:** "ADMİN DÜZENLEDİ" logu değişen alanların eski/yeni değerini
  yazar; not geçmişi değiştiyse silinen/eklenen satırlar ayrı "ADMİN NOT DÜZENLEDİ" kaydına
  yazılır. Excel hücre sınırını aşan log detayı işlemi engellemez, işaretlenerek kısaltılır.
- **Manuel kart:** aynı Talep NO + Kart Stok No ile listelerde görünen (Excel'den veya elle)
  kart varsa yeni kart açık onayla eklenir; onay logda kalır. Aynı talep ve stok birden
  fazla satırda meşru olabildiği için engellenmez. Yeni manuel kartların anahtarı
  `MANUEL:<kart ID>`; eski manuel kartlar `Talep|Stok` anahtarını korur.

Testler: `tests/test_orta_oncelik.py` (14 test; 12'si değişiklik öncesi kodda başarısız).
