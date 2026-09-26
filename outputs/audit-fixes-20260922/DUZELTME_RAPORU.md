# PDGM düzeltme ve doğrulama raporu — 22.09.2026

20.09.2026 sistem analizinde tespit edilen 13 bulgu düzeltildi. Son koşuda gerçek Microsoft Excel COM testleri dahil **50/50 regresyon testi** ve **24/24 sistem testi** geçti. Testler geçici veri klasörlerinde çalıştı; üretim kayıtları değiştirilmedi. Bu rapor test edilen kapsamı gösterir; bütün olası girdilerin veya üretim yükünün doğrulandığı anlamına gelmez.

## Bulgular ve sonuçları

| Bulgu | Kök neden ve düzeltme | Doğrulama |
|---|---|---|
| F01 — Eski yönetim formu yeni not/adedi eziyor | Kaydetmede kayıt sürümü kontrol edilmiyordu. Kalıcı kart alanlarının özeti kilit altında karşılaştırılıyor; eski form 409 ile reddediliyor. | Yeni not ve tamamlanan adet korunuyor; iki tarayıcı sekmesiyle eski form reddi görüldü. |
| F02 — Uzun not sessizce kesiliyor | Excel hücre uzunluğu sınırı yazmadan önce denetlenmiyordu. 32.767 karakteri aşan metin reddediliyor; bellek ve dosya geri alınıyor. | 33.000 karakterlik not reddi; önceki dosya baytları ve kayıt korunuyor. |
| F03 — COM kimlik dönüşümü | General biçimli hedef hücre, metinsel sayıları dönüştürüyordu. NO/PDGM ROW ID hedef sütunu Value2 yazımından önce metin yapılıyor. | Gerçek COM: `001`, `1`, uzun sayısal metin, gizli kimlik sütunu ve tekrar import geçti. |
| F04 — Operatör tipi oturumda eski kalıyor | Kullanıcı kontrolü operatör tipini yenilemiyordu. Her istekte güncel kullanıcı kaydı oturuma yansıtılıyor. | Canlı tip değişikliği sonrası eski yetkiyle işlem 409 alıyor. |
| F05 — Oturum kaybı başarılı işlem gibi görünüyor | API giriş sayfasına yönleniyor, istemci HTML 200 yanıtını başarı sayıyordu. API 401 JSON döndürüyor; istemci HTML/yönlendirme yanıtını reddediyor. | Flask 401 kontrolü ve gerçek JS fonksiyonunda HTML/401/yönlendirme ret testleri geçti. |
| F06 — Kesirli adet aşağı yuvarlanıyor | `int()` kesirli sayıyı sessizce dönüştürüyordu. İşlem adetleri katı tam sayı kontrolünden geçiyor. | 1.9, bool, NaN/Infinity gibi girdiler reddediliyor; adet değişmiyor. |
| F07 — Pano filtresi eksik kart arıyor | Sunucu listeleri filtre öncesi 12/20/200 ile kesiyordu. Tüm ilgili kayıtlar gönderiliyor, istemci filtreledikten sonra 12'li sayfalıyor. | 36 planlı kartın tamamı erişilebilir; ikinci sayfa ve iki ELDE kartı tarayıcıda doğrulandı. |
| F08 — Tümü özeti tabloyla uyuşmuyor | Özet 2000–bugün aralığı uygularken tablo tüm teslimleri gösteriyordu. Tümü özeti tarihsiz kayıtları da içeriyor. | 8 tablo kaydı / 8 özet işi eşleşiyor. Sapma ölçülemeyen kayıtlar açıklanıyor. |
| F09 — Import açıklaması yanıltıcı | Yardım metni durum/adetlerin daima korunacağını söylüyordu. Gerçek kurallar açıklanıyor; değişiklik önizlemesi ve uygulama adımı eklendi. | Önizlemede kayıtlar değişmiyor; dosya, hesaplanan değer veya kartlar değişirse uygulama reddediliyor. Tarayıcıdan önizleme ve gerçek COM uygulaması geçti. |
| F10 — Özel tarih aralığı kayboluyor | Dönem seçimi saklanıyor, tarihler saklanmıyordu; yarışan istekler eski sonuç gösterebiliyordu. Tarih çifti saklanıyor, takvim doğrulanıyor, eski yanıtlar yok sayılıyor. | 18.09.2026 tek günlük aralık ve tek sonuç yenilemeden sonra korundu. |
| F11 — Türkçe arama eşleşmiyor | Sunucu ve istemci farklı küçük harf dönüşümü kullanıyordu. İki taraf NFC ve Türkçe küçük harf dönüşümüyle karşılaştırılıyor. | TEST VERİSİ, IŞIK ve ayrık Unicode karakterleri test edildi; tarayıcı araması sonuç döndürüyor. |
| F12 — Mobil yatay taşma | Grid öğelerinin asgari genişliği tabloyu sayfa dışına itiyordu. Grid ve öğe minimumları düzeltildi; tablo kendi alanında kayıyor. | 390×844 viewport ayarında ölçülen belge genişliği ve scrollWidth eşit: 375/375 CSS piksel. |
| F13 — Bozuk logo | Var olmayan PNG dosyasına referanslar vardı. Metinsel PDGM işareti kullanılıyor, sola kaydırma kaldırıldı. | Son Pano kontrolünde bozuk görsel sayısı 0. |

## Import ve veri koruma kuralları

- Sayfa içindeki NO, kullanıcının doğruladığı gibi sabit ve benzersiz kimlik kabul edilir. Yeniden numaralandırma veya kimliğin başka iş için kullanılması desteklenen kimlik sözleşmesini bozar.
- Geçerli Excel DURUM değeri aktarımda yetkilidir. PLANA ALINDI ve HAZIR tamamlanan adedi 0 yapar; TESLİM EDİLDİ toplam adede eşitler. DİZGİDE veya boş/geçersiz DURUM kısmi adedi korur. Notlar importta korunur.
- Toplam adedin korunması gereken tamamlanan adetten küçük olması gibi çelişkiler importu iptal eder. Dosyadan çıkarılan satırlar ve kaldırılan sayfaların kartları pasifleşir; manuel kartlar korunur.
- Yönetim ekranındaki aktarım önce değişiklik önizlemesi gösterir. Onay sırasında depo sürümü, kaynak dosya SHA-256 özeti ve yeniden hesaplanmış/parse edilmiş satır özeti kontrol edilir. Değişiklik varsa yeni önizleme gerekir.
- Önizlemesiz mevcut programatik import ve eski POST kullanımları uyumluluk için korunmuştur. Yönetim arayüzü önizleme kullanır.
- Yönetim düzenleme API'si artık `surum` ister. Harici istemci varsa güncel kart sürümünü göndermelidir. XLSX saklama şeması değişmedi.
- **Not geçmişinin toplam sınırı 32.767 karakterdir**; tarih ve kullanıcı etiketleri de buna dahildir. Sınırı aşan işlem kaydedilmez. Bu düzeltme geçmişte kesilmiş metinleri geri getirmez.

## Test kanıtları

- [Regresyon koşusu](regression-tests.txt): 50 test, gerçek COM dahil, atlanan test yok.
- [Sistem koşusu](automated-tests.txt) ve [ölçümler](results.json): 24 test, hata/başarısızlık yok. İş akışları, tamamlanan adet, notlar, eşzamanlı eski form, roller ve raporlar kapsandı.
- [19 Excel dosyasının gerçek COM doğrulaması](DOGRULAMA_COM.txt): geçerli senaryolar ve beklenen retler geçti; ret sonrası bellek/dosya bütünlüğü kontrol edildi.
- [Frontend sözleşmesi](frontend-tests.txt), [arama/yanıt probları](frontend-probes.json), [API oturum probu](route-probes.json) ve [monitör dönüş testi](monitor-tests.txt) geçti.
- [Tarayıcı gözlemleri](browser-verification.json): gerçek arayüz üzerinden yapılan kontrollerin kayıtlarıdır; otomatik test sayısına eklenmemiştir.
- [Kod farkı](changes.diff) ve [kaynak dosya özetleri](source-hashes.json) bu doğrulanan değişiklikleri tanımlar.

Test komutları: `PDGM_TEST_EXCEL_COM=1` ortamıyla `.venv/Scripts/python.exe -B -m unittest discover -s tests -v`; `.venv/Scripts/python.exe -B artifact_work/system-audit/audit.py`; `node tests/frontend_contract.cjs`; `node tests/monitor_rotation.cjs`. Excel paketi ayrıca `artifact_work/pdgm-import/verify.py --com` ile çalıştırıldı.

## Teslim ve sınırlar

Değişiklikler app.py, depo.py, excel_araclari.py, ilgili HTML şablonları, static/stil.css ve regresyon testlerindedir. Değişiklik öncesi kaynak kopyaları `.investigation/audit-fixes-before/` altındadır. Bu kaynak yedeği üretim veri yedeği değildir.

Son kod farkı yerel olarak incelendi. Ayrı inceleyici kullanım sınırına takıldığı için bu turda bağımsız inceleme sonucu alınamadı.

Üretim sunucusu yeniden başlatılmadı. Python değişikliklerinin çalışan uygulamaya geçmesi için normal uygulama yeniden başlatması gerekir; ardından tarayıcı yenilenmelidir. Test sunucusu geçici depoyla çalıştırılmıştır.

İleride veri hacmi büyürse sunucu tarafında sayfalama; not geçmişi büyürse ayrı not tablosu; süreç/güç kesintisine karşı daha güçlü kalıcılık gerekirse veritabanı değerlendirilmelidir. Bu tur yük testi, kesinti dayanıklılığı sertifikasyonu veya kapsamlı sızma testi değildir. XLSX birden fazla dosyada gerçek veritabanı transaction garantisi sağlamaz; mevcut yedekleme ihtiyacı sürer.
