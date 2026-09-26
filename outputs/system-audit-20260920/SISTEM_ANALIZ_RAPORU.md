# PDGM — kapsamlı sistem ve arayüz incelemesi

İnceleme: 20–21 Eylül 2026. Son doğrulama: 21 Eylül 2026.

**Sonuç:** İlk bildirilen satır kimliği, yeniden import, tarih güncelleme, kaynakta silinen satırı pasifleştirme ve MAKİNE/ELDE karışması senaryoları mevcut sürümde testleri geçti. Ancak sistem için “sorunsuz” sonucu verilemez. Yeni not/adet bilgisinin eski yönetim formuyla ezilmesi, uzun notun diskte kırpılması ve COM aşamasında metinsel kimliğin değişmesi öncelikli açık sorunlardır. Arayüzde ayrıca filtre, özet ve oturum hatası geri bildirimi sorunları doğrulandı.

Bu inceleme turunda uygulamanın Python, HTML ve CSS dosyaları değiştirilmedi. Yalnız test yardımcıları ve raporlar eklendi/güncellendi. Testler geçici depoda, sentetik kullanıcılarla ve 89 kartlık Excel paketiyle yürütüldü; gerçek `data/` dosyaları test hedefi yapılmadı. Tarayıcı testi ayrı `127.0.0.1:5017` sunucusunda yapıldı.

## 1. Test sonuçları ve kanıtlar

| Kontrol | Sonuç | Kanıt |
|---|---|---|
| Mevcut Python regresyon paketi | 40 test: 39 geçti, 1 isteğe bağlı COM testi atlandı | `regression-tests.txt` |
| Ek kapsamlı sistem paketi | 24 test: 17 geçti, 7 başarısız, 0 çalışma hatası | `automated-tests.txt`, `results.json` |
| Gerçek Excel COM import paketi | 19 dosya; normal senaryolar ve beklenen retler doğrulandı; B01 bilinen hata yeniden üretildi | `DOGRULAMA_COM.txt` |
| Monitörün gerçek JavaScript kodu | 12 saniye bekleme, sayfa konumunu koruma ve 25 kartın yenilemeden önce gösterilmesi geçti | `monitor-rotation.txt` |
| Frontend hedefli kod deneyleri | Türkçe arama eşleşmesi ve giriş HTML'inin hata sayılması beklentileri karşılanmadı | `frontend-probes.json` |
| Oturumsuz API isteği | 302 giriş yönlendirmesi, takip edildiğinde HTTP 200 HTML | `route-probes.json` |
| Gerçek tarayıcı | Pano, Operatör, Monitör, Yönetim; üretim ve not akışı; masaüstü ve dar görünüm | Bölüm 5 ve `browser-observations.json` |

Yedi başarısız test, hata araştırmak için beklenen doğru davranışı sınar; “beklenen başarısızlık” olarak geçer sayılmadı. İkisi aynı eski form kök nedeninin not ve adet üzerindeki ayrı etkileridir. COM doğrulayıcısının başarılı çıkışı da B01'in düzeldiği anlamına gelmez: bu senaryoda mevcut hatanın yeniden üretildiği kontrol edilir. Frontend deneyleri gerçek şablondan alınan fonksiyonla, kontrollü yanıt/DOM dışı verilerle çalışır; tarayıcı uçtan uca testi olarak sayılmaz.

## 2. İlk sorunların mevcut durumu

| İstek / senaryo | Gözlenen sonuç |
|---|---|
| Aynı dosyayı iki kez import | Yeni mükerrer kart oluşmadı |
| Plan başlangıç/teslim tarihini değiştirme ve temizleme | Aynı kart güncellendi; ID değişmedi; boş tarih eski tarihi temizledi |
| Geçerli Excel DURUM değişikliği | DİZGİDE ve TESLİM EDİLDİ geçişleri uygulandı |
| Satır silme | Yalnız ilgili kaynak kartı pasifleşti |
| Silinen satırın geri gelmesi | Aynı ID ile aktifleşti |
| Aynı Talep/Stok, farklı sayfalar | 2 adet ELDE teslim / 5 adet MAKİNE dizgide bağımsız kaldı |
| Aynı Talep/Stok, aynı sayfada farklı NO | Üç ayrı satır karışmadı; tek duplicate silme doğru kartı etkiledi |
| Fiziksel satır sırasını değiştirme | Eski ID'ler korundu |
| Talep/Stok düzeltme | Sabit kaynak kimliği üzerinden mevcut kart güncellendi |
| Kaydetme ve depoyu yeniden yükleme | Kimlikler, kaynak aktifliği, tarihler ve olağan uzunluktaki notlar korundu |
| EÜM başlangıç tarihi | Pazartesiye normalizasyon; hafta/yıl geçişleri, boş tarih ve saat dilimli kaynak günleri doğrulandı |
| Monitör | Backend'den yalnız MAKİNE; iki grup da doğru; PLANDA etiketi yerel; 12 saniye rotasyon |
| Yanlış/eksik kaynak alanları | H01–H13 reddedildi; bellek ve kalıcı dosyalar korunarak import iptal edildi |
| Opsiyonel sayfanın kaldırılması | ELDE sayfası kaldırılınca yalnız ilgili 8 kart pasifleşti |
| Başlıkları geçerli tamamen boş kaynak | Excel kartları pasifleşti; manuel kart aktif kaldı |

Kimlik modeli `source_sheet + source_row_id` esaslıdır. Örnek: `EXCEL:v2:["MAKINE","NO:123"]`. Tarihler, miktar, Talep/Stok ve fiziksel sıra anahtar değildir. Kullanıcının NO'nun sayfa içinde sabit ve benzersiz olduğu teyidi bu modeli destekler. NO tekrar başka işe verilmemeli; sayfa değişimi ayrı kaynak kimliği anlamına gelir. Parser seviyesinde metinsel `001` ve `1` ayrılabilir; COM yolunun bu ayrımı bozması açık kalmıştır.

Mevcut eski-format/migration regresyonları da pakette çalıştı. Gerçek üretim dosyası üzerinde migration veya geçmişte karışmış kayıtları otomatik onarma yapılmadı. Eski kayıtlardaki belirsiz eşleşmelerin güvenli biçimde reddedilmesi, onların kendiliğinden düzeldiği anlamına gelmez.

## 3. Importun ezme / koruma kuralları

Kod dayanağı: `depo.py:2169`, özellikle `2208–2258`; açıklama: `docs/EXCEL_SYNC.md`.

| Alan | Yeniden import davranışı |
|---|---|
| Kaynak kimliği ve mevcut kart ID'si | Aynı kaynak satırında korunur |
| Talep/Stok, sıra, talep sahibi | Excel'den güncellenir |
| Toplam adet, adet metni, plan haftası, plan başlangıç/teslim | Excel'den güncellenir; boş tarih temizler |
| Gerçekleşen teslim tarihi | Excel'den güncellenir/temizlenir; import günü gerçek teslim günü gibi uydurulmaz |
| Kaynak durum, PCB, dizgi tipi/sorumlusu, malzeme bekliyor | Kaynaktan güncellenir |
| Geçerli DURUM | Excel belirleyicidir |
| Boş/geçersiz DURUM | Var olan workflow ve tamamlanan adet korunur; yeni kart durumu eksik kalır |
| Not geçmişi, yönetici gizliliği | Normal importta korunur |
| Operatör | Korunur; boş değer geçerli kaynak durumunda Excel olarak doldurulabilir |
| Workflow zamanları | Duruma göre korunur/temizlenir; kaynak teslim tarihi değiştiğinde eski teslim zamanının görünmesi engellenir |
| Kaynakta bulunmayan kart | Fiziksel silinmez, pasifleşir |
| Manuel kart | Kaynak pasifleştirmesine katılmaz |

**Somut adet deneyi:** Toplam 10, tamamlanan 3 olan kart yeniden import edildi:

| Excel DURUM | Son tamamlanan adet | Not |
|---|---:|---|
| Boş/geçersiz | 3 | Korundu |
| PLANA ALINDI | 0 | Korundu |
| HAZIR | 0 | Korundu |
| DİZGİDE | 3 | Korundu |
| TESLİM EDİLDİ | 10 | Korundu |

Bu, önceki “geçerli Excel DURUM belirleyici olsun” kuralının mevcut uygulamasıdır. Dolayısıyla normal importun notları koruması, üretim adetlerini her durumda koruduğu anlamına gelmez. Eski PLANA ALINDI durumunu içeren dosyayı tekrar yüklemek ilerlemeyi sıfırlayabilir. DİZGİDE/korunan workflow durumunda toplam adedin tamamlanandan küçük olması ise bütün importu reddeder.

Import **tam kaynak görüntüsü** olarak uygulanıyor. Bir sayfanın veya bütün kaynak satırlarının kaldırılması pasifleştirme yapar. Bu nedenle test paketini canlı üretim deposuna yükleyerek deneme yapmak uygun değildir; mevcut testler ayrı depo kullandı.

## 4. Kanıtlanmış açık sorunlar ve çözüm önerileri

P1: veri bütünlüğü / yetki / yanlış başarı bilgisi; P2: işlevsel doğruluk; P3: görsel kalite. Öncelikler bu incelemenin değerlendirmesidir.

### F01 — P1: Eski yönetim formu yeni notu ve üretim adedini eziyor

**Tekrar üretim:** Yönetim formunun eski kopyası alınır; operatör yeni not ekler ve tamamlanan adedi 3'ten 5'e çıkarır; eski form kaydedilir. HTTP 200 gelir, yeni not kaybolur ve adet tekrar 3 olur. `test_bug_03` ve `test_bug_06` başarısızdır.

**Kök neden:** `templates/yonetim.html:687` eski alanları formda tutar; `754–765` notu ve mutlak tamamlanan adedi birlikte gönderir. `depo.py:1830` içindeki `admin_kart_duzenle` güncel sürüm kontrolü yapmadan `1960` ve `1969` satırlarında bunları değiştirir. `_kilit` aynı anda işlem yürütmeyi engeller; kullanıcıda açık duran eski formu tespit etmez.

**Öneri:** Kart revizyonu ile koşullu güncelleme, çakışmada HTTP 409 ve değişen alanları gösterme. Not eklemeyi geçmişi değiştirmekten ayırma. Sayaç düşürme ve geçmiş düzeltmelerini açık, denetlenebilir yönetici işlemi yapma. Saniye çözünürlüklü zaman damgası tek başına güvenilir revizyon olmayabilir.

### F02 — P1: Uzun not başarıyla kaydedilmiş görünürken diskte kırpılıyor

**Kanıt:** Not ekleme HTTP 200; bellekte 33.049 karakter, depo tekrar açılınca 32.767 karakter. Sondaki `END-MARKER` yok. `test_bug_07` başarısızdır.

**Kök neden:** `depo.py:253` not geçmişini tek metne ekler; `487` Excel hücresine uzunluk sınırı denetlemeden yazar. Hücre yazımı sırasında uzun metin kırpılır. Sorun bir tek uzun notla veya zamanla büyüyen toplam geçmişle oluşabilir.

**Öneri:** Kısa vadede toplam hücre bütçesini kaydetmeden doğrulayıp açıklayıcı hata verin. Kalıcı çözüm olarak notları ayrı satırlarda/ayrı not sayfasında kimliğiyle saklayın. Bunun için Flask veya bütün depoyu değiştirmek gerekmez.

### F03 — P1: Excel COM metinsel kimliklerin başındaki sıfırları kaybediyor

**Kanıt:** `B01_BASTAKI_SIFIR_COM_HATASI.xlsx` A35 kaynak değeri `'001'` (metin, `@`), A36 `'1'`; COM snapshot'ta ikisi de sayısal `1`, `General`. Parser sonra tekrar eden `NO:1` diyerek importu reddeder. Bu örnekte merge uygulanmadı; verinin ezildiği iddia edilmiyor.

**Kök neden:** `excel_araclari.py:183`, `_sayfa_kopyala` içindeki `Value2` ataması metin kimliklerini hedef hücrenin varsayılan biçimine taşır. Sorun parserdan önce oluşur.

**Öneri:** NO/PDGM_ROW_ID sütunlarını snapshot'ta metin olarak kayıpsız taşıyın; kaynak türünü koruyun. Gerçek COM ile `001`, `1`, uzun sayısal görünümlü metin, harfli ve gizli kimlik sütunlarını regresyona bağlayın. Doğrudan openpyxl testinin geçmesi COM yolunun geçtiğini kanıtlamaz.

### F04 — P1: Operatör tipi değişikliği açık oturuma uygulanmıyor

**Kanıt:** Kullanıcı dosyasında tipi MAKİNE'den ELDE'ye değiştirilen kullanıcı, mevcut oturumuyla MAKİNE kartını başlatabildi; HTTP 200. `test_bug_04` başarısızdır. Hesabı pasifleştirme ise oturumu engelledi.

**Kök neden:** `app.py:329` içindeki istek öncesi denetim rol ve adı yeniler, `operator_tipi` değerini yenilemez. İşlem API'leri oturumdaki eski alt tipi kullanır (`app.py:832` ve diğer üretim API'leri).

**Öneri:** Yetkilendirmede güncel kullanıcı kaydını esas alın veya alt tip değişiminde oturumu geçersiz kılın. Genel rol ve alt tip aynı yöntemle güncel tutulmalı.

### F05 — P1: Oturum yokken işlem arayüzü yanlış başarı gösterebilir

**Kanıt:** Oturumsuz `/api/not` POST isteği `/giris` adresine 302 döner; takip edildiğinde HTTP 200 HTML gelir (`route-probes.json`). Gerçek `pdgmFetch` fonksiyonuna bu yanıt verildiğinde hata fırlatmadı; `{hata: HTML}` döndürdü (`frontend-probes.json`). Bu, birleşik route/fonksiyon deneyiyle doğrulandı; gerçek tarayıcıda oturumun zaman aşımına uğraması beklenmedi.

**Kök neden:** `app.py:347` yetki dekoratörü API isteğinde de girişe yönlendirir. `templates/base.html:75` fonksiyonu yalnız `response.ok` kontrol eder; JSON olmayan başarılı yanıtı reddetmez. `templates/operator.html:309` ve form başarı akışları, hata fırlamazsa başarı bildirimi/yenileme yapar.

**Öneri:** API için 401 JSON sözleşmesi, frontend'de oturum yönlendirmesi ve beklenmeyen içerik türü denetimi; kullanıcıya işlemin kaydedilmediğini açıkça gösterme.

### F06 — P2: Kesirli JSON adedi sessizce tam sayıya çevriliyor

**Kanıt:** 3 tamamlanmış karta `adet: 1.9` gönderildi; HTTP 200, sonuç 4. `test_bug_02` başarısızdır. Excel parserı kesirli miktarı doğru reddeder; bu sorun işlem API'sindedir.

**Kök neden:** `depo.py:1561` doğrudan `int(adet)` kullanır. Başlatma (`1474`) ve yönetici adet dönüşümleri (`1868`) benzer yapıdadır; yürütülen somut deney tamamlama yolundadır.

**Öneri:** Ortak katı pozitif tam sayı doğrulaması. Tarayıcının normal number girdisi koruması sunucu doğrulamasının yerini tutmamalı.

### F07 — P2: Pano filtresi bazı gerçek plan kartlarını hiç göremiyor

**Kanıt:** Depoda 36 plan kartı, Pano'ya gönderilende 12 kart. Tarayıcıda ELDE + PLANA ALINDI seçildiğinde üst sayaç 2 iken sonuç bulunamadı. `test_bug_05` başarısızdır.

**Kök neden:** `app.py:735` tüm plan kartlarını filtrelemeden önce `[:12]` ile keser; frontend sadece gönderilmiş satırları filtreleyebilir. Bölüm açıklamasında “ilk 12” denmesi bunu kısmen anlatıyor fakat tip filtresiyle istenen ilk 12 sonucu sağlayamıyor.

**Öneri:** Filtreyi önce uygulayın, sonra sayfalayın; toplam eşleşme ve gösterilen satır sayısını ayırın. Teslim API'sindeki 200 satır sınırını da aynı sayfalama sözleşmesine alın (`app.py:1404` civarı); bu sınırın yüksek hacimli uçtan uca testi yapılmadı.

### F08 — P2: “Tümü” teslim özeti tarihsiz teslimleri dışarıda bırakıyor

**Kanıt:** Başlangıç paketinde liste 8 kart / 40 adet; özet 6 iş emri / 30 adet. `test_bug_01` ve tarayıcı aynı farkı gösterdi.

**Kök neden:** `app.py:1367` civarında “Tümü” özeti 2000–bugün tarih aralığıyla hesaplanır; liste ise tüm teslimleri alır. Tarihsiz teslimler listede kalırken özetten düşer. Aynı yapı tarih aralığı dışındaki teslimleri de özetten dışlar.

**Öneri:** Tümü kart/adet sayısını seçilen bütün teslimler üzerinden hesaplayın; tarih gerektiren performans ölçülerini yalnız ölçülebilen alt kümede hesaplayıp tarihsiz sayısını ayrıca belirtin.

### F09 — P2: Import açıklaması gerçek veri sahipliğiyle çelişiyor

**Kanıt:** Yönetim ekranı “mevcut workflow, tamamlanan adet ve operatör işlemleri korunur; yalnız plan/source alanları güncellenir” diyor (`templates/yonetim.html:38`). Bölüm 3'teki yürütülmüş miktar matrisi bunun koşulsuz doğru olmadığını gösteriyor.

**Öneri:** Alan sahipliği ve geçerli DURUM kurallarını açıklayın. Import önizlemesinde eklenecek/pasifleşecek kartlar, durum gerilemeleri, sıfırlanacak adetler ve temizlenecek tarihler gösterilsin. Büyük pasifleştirme veya ilerleme sıfırlama açık sonuç özeti gerektirir. İş kuralını sessizce tersine çevirmek yerine mevcut kuralı kullanıcıya doğru anlatmak önceliklidir.

### F10 — P2: Özel tarih aralığı yenilemeden sonra kayboluyor

**Kanıt:** Pano'da 18.09.2026–18.09.2026 uygulanınca tek TEST-1021 satırı geldi. Sayfa yenilenince Özel Aralık seçili kaldı, iki tarih boşaldı, başlangıçtaki 8 teslim satırı göründü; performans alanları `—` oldu.

**Kök neden:** `templates/panel.html:404`, `537` dönem adını saklar; tarih değerlerini saklamaz. `546` özel dönem için otomatik yükleme yapmaz.

**Öneri:** Dönem ve iki tarihi birlikte saklayıp açılışta doğrulayarak yükleyin; eksik/geçersiz durumda açıkça Tümü'ne dönün. Ekranda seçili filtre ile gösterilen veri aynı olmalı.

### F11 — P2: Türkçe büyük İ içeren arama eşleşmiyor

**Kanıt:** Operatör ekranında görünen `TEST VERİSİ` aranırken sonuç sıfır oldu. Şablon indeksinde `test veri̇si̇`, sorguda `test verisi` oluşuyor; kod deneyi eşleşmediğini doğruladı.

**Kök neden:** Sunucuda Jinja `|lower`, tarayıcıda `toLocaleLowerCase('tr-TR')` farklı normalizasyon üretir. `templates/operator.html:53,267`; Pano `84,136,218,279`; yönetim `318,626` aynı yapıyı kullanır. Yönetimdeki etki kaynak incelemesidir; somut tarayıcı tekrarı Operatör'de yapıldı.

**Öneri:** İndeks ve sorguya aynı Türkçe/Unicode normalizasyonunu uygulayın. `İ/i`, `I/ı`, birleşik karakterler ve Türkçe stok/sorumlu adlarını test edin.

### F12 — P2: Mobil yönetim sayfası yatay taşıyor

**Kanıt:** 390×844 görünüm isteğinde masaüstü ölçeklendirmesiyle ölçülen CSS viewport 312 piksel; belge genişliği 942 piksel. Sayfa kabı 292,5 pikselken yönetim gridleri yaklaşık 932 piksel, tablo kabı 902,5 piksel. Operatör görünümü aynı dar testte viewport içine sığdı.

**Kök neden:** Grid çocuklarının otomatik minimum içerik genişliği tablo içeriğini üst yerleşime taşıyor. `static/stil.css:187,691,1741,1750,2088`; tablo `min-width:760px`, mobil grid `1fr`. Ölçümde grid çocukları `min-width:auto` kalıyor.

**Öneri:** Grid sütunları/çocuklarında minimum genişliği sıfırlayarak taşmayı tablo kabına sınırlayın; yönetim aksiyonlarını dar ekranlarda erişilebilir tutun. Önerilen CSS değişikliği bu turda uygulanmadı veya çözüm olarak test edilmedi.

### F13 — P3: Logo dosyası eksik

**Kanıt:** DOM'da `/static/pdgm_logo.png` tamamlanmış fakat `naturalWidth=0`; dosya mevcut static klasöründe yok. `templates/base.html:15`, `templates/giris.html:12`, CSS logo arka planı aynı varlığa bağlı.

**Öneri:** Doğru logo varlığını ekleyin veya eksik görsel alanını kaldırın.

## 5. Tamamlanan özellik kontrolleri ve frontend değerlendirmesi

- **Üretim:** MAKİNE, ELDE ve EÜM kullanıcılarıyla başlatma, adet bitirme ve ayrı teslim işlemi API üzerinden geçti. Sıfır/negatif/fazla adet, üretim tamamlanmadan teslim ve tekrar teslim reddedildi. Adet tamamlandığında ayrı teslim adımı gerektiği doğrulandı.
- **Tarayıcı üretim akışı:** TEST-1012 için 10 adet başlatma → 3 adet bitirme → bağımsız çok satırlı Türkçe not → kalan 7 adedi bitirme → teslim onayı. Son görünüm TESLİM EDİLDİ, 10/10, kalan 0; üç not da görünür. Aktif filtresinden çıkması ve teslim filtresinde bulunması doğrulandı.
- **Notlar:** Başlatma/tamamlama/bağımsız not, zaman ve kişi bilgisi, satır sonları, normal import ve yeniden okuma koruması geçti. Aynı süreçte 8 eşzamanlı notun hepsi birer kez korundu. Bu sonuç eski yönetim formu ve uzun not hatalarını ortadan kaldırmıyor.
- **Yetki:** Yönetim erişimi, yönetici API'leri, gözlemci sınırları, üretim tipi ayrımı, CSRF ve hesabın pasifleştirilmesi sınandı. Mevcut oturumda alt tip değişikliği istisnası F04'tür. Bu çalışma tam güvenlik sızma testi değildir.
- **Malzeme:** Önce 409/onay gereksinimi, onaylı başlatma sonrasında üretime geçiş ve bayrağın temizlenmesi geçti.
- **Yönetim:** Manuel kart, gizleme/geri getirme, kaynakta pasif kart, aynı ID ile geri geliş, not/adet içeren yedekten dönüş ve yeniden okuma geçti. Gizlenen kart sonraki importta gizli kaldı.
- **Raporlar:** Rapor, kart, log ve yükleme Excel indirmeleri açılabilir çalışma kitabıydı. Not HTML'i escape edildi; formül görünümlü metin bu testte formül hücresine dönüşmedi.
- **Hata dayanımı:** Simüle yazma kilidi HTTP 423 döndürdü; bellek ve kart dosyası değişmedi. Boş dosya seçimi, yanlış uzantı ve bozuk `.xlsx` reddedildi. Gerçek COM hatalı çalışma kitaplarında da yarım import bırakmadı.
- **Pano/monitör:** Tip ve dönem filtreleri, tarih aralığı doğrulaması, rozet sınırları ve stok sayaçları incelendi. Monitörün gerçek tarayıcıda farklı kartlara döndüğü görüldü; deterministik JS testi 12 saniyeyi ve yenileme sürekliliğini doğruladı. Masaüstü düzeni kullanılabilir; yönetim dar ekranda sorunlu.

Önerilen frontend iyileştirmeleri, yeni bir frontend yazmadan uygulanabilir: görünür filtre özeti ve doğru toplam/gösterilen sayıları; import etkisi önizlemesi; açık oturum sona erdi bildirimi; yönetim düzenlemesinde çakışma uyarısı; not geçmişini ayrı, okunabilir kayıtlar halinde gösterme; mobil yönetim taşmasını düzeltme. Önce F01–F05 veri güvenilirliğini, sonra F06–F12 işlevsel tutarlılığı, son olarak görsel ayrıntıları ele alın.

## 6. Sınırlar ve tekrar çalıştırma

“Bütün özellikler” ana işlev aileleri üzerinden ele alındı; bütün giriş kombinasyonlarının veya üretim ölçeğindeki tüm davranışların kusursuz olduğu iddia edilmiyor. Sentetik veriyle 89 başlangıç kartı kullanıldı. Binlerce kart performansı, çok süreçli dağıtım, elektrik kesilmesi/işletim sistemi çökmesi, gerçek üretim migration'ı ve kapsamlı güvenlik testi yapılmadı. Rol testleri ağırlıkla gerçek Flask route'ları üzerinden, tarayıcı akışı ise sentetik yöneticiyle yapıldı.

Sistem auditinde COM sınırı kontrollü dosya kopyasıyla değiştirilmiştir; parser, merge, depo, route ve şablonlar gerçektir. Gerçek COM ayrı 19 dosyalık doğrulamada çalıştırılmıştır. Tarayıcıdan dosya seçip COM import etme zinciri bu audit turunda tek bir uçtan uca senaryo olarak çalıştırılmadı; route zinciri ve gerçek COM zinciri ayrı doğrulandı.

Rapor rozetleri için 20 Eylül 2026 referans tarihini kullanan deterministik testler içerir. Test Excel'i zaman geçtikçe farklı gecikme rozetleri gösterebilir; bu tek başına regresyon değildir.

Proje kökünde PowerShell komutları:

```powershell
.venv/Scripts/python.exe -B -m unittest discover -s tests -v
.venv/Scripts/python.exe -B artifact_work/pdgm-import/verify.py --com
.venv/Scripts/python.exe -B artifact_work/system-audit/audit.py
node tests/monitor_rotation.cjs
node artifact_work/system-audit/frontend-probes.cjs
.venv/Scripts/python.exe -B artifact_work/system-audit/route-probes.py
```

`audit.py` ve `frontend-probes.cjs`, açık hatalar sürdüğü müddetçe çıkış kodu 1 üretir. Bu bilerek gizlenmeyen başarısız test sonucudur. Uygulama düzeltmesi yapılmış gibi değerlendirilmemelidir.

Test paketi: `outputs/pdgm-import-test-20260920/PDGM_TEST_PAKETI.zip`. Yeni audit yardımcıları: `artifact_work/system-audit/`. Önceki kök neden ve uyumluluk belgeleri: `docs/ROOT_CAUSE_REPORT.md`, `docs/EXCEL_SYNC.md`.
