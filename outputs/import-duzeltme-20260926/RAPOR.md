# Excel import düzeltme raporu — 26.09.2026

Kaynak dosya: `PDGM_Kart_dizgi_Talepleri_Üretim_Takvimi.xlsx` (proje kökü; `C:\gokberk_flask\` yolu mevcut değil).
Üretim `data/` klasörü okunmadı/yazılmadı; testler geçici klasörde, uçtan uca deneme izole kopyada yapıldı.

## 1. Hatalar ve kök nedenler

| # | Belirti | Kök neden |
|---|---|---|
| 1 | `'MAKİNE' satır 211 ... eksik satır silinmiş kabul edilmedi` | Satır 211 gizli; yalnız `I` Kart Stok No (`AD-0000-0001`) ve gizli `K` Birim Seviyesi dolu. 20.09 düzeltmesi Talep/Stok/NO'dan biri dolu her eksik satırı "yarım gerçek kayıt" sayıp reddediyordu. Dosyada 52 MAKİNE + 4 EÜM böyle kalıntı satırı var (hepsi gizli; Talep NO, adet, DURUM, tarih yok). |
| 2 | (211 düzelince) EÜM satır 3'te import duracaktı | EÜM'de `Planlanan Başlangıç T.` hafta metni (`23. Hafta (2.06 haftası)`); parser tarih bekliyordu. |
| 3 | ELDE `DİZGİDE VE MALZEME BEKLENİYOR` → durum eksik, malzeme bayrağı yok | Eşleme yalnız `...BEKLİYOR` yazımını tanıyordu. |
| 4 | Gizlenen DURUM / teslim sütunu tüm kartlarda veriyi boşaltırdı | COM snapshot gizli sütunları tamamen atıyordu (yalnız NO korunuyordu). |
| 5 | Metin Talep/Stok değerleri bozulabilirdi | COM `Value2` yazımında Excel metni yeniden yorumluyor: `001234`→1234, `1-2`→tarih, `=abc`→formül (gerçek Excel ile kanıtlandı). |
| 6 | Hücre hataları (#N/A) sayı gibi okunurdu | pywin32 hata hücrelerini `-2146826246` gibi tamsayı döndürüyor. |
| 7 | Aynı dosyanın tekrar yüklenmesi "N güncellendi" ve sürüm değişikliği | Her eşleşen kartın `guncelleme` damgası yenileniyordu; yeniden başlatma sonrası `""`↔`None` hayali farkları audit logu ve önizleme satırı üretiyordu. |
| 8 | Rapor sayaçlarında Excel'den silinmiş kartlar | `ozet_hesapla` kaynakta pasif kartları sayıyordu (pano/teslim API'si saymıyordu). |

Eski stale veri problemi güncel kodda tekrar üretilemedi: kimlik (sayfa kodu + NO), tarih/durum değişikliğinde aynı kart güncelleniyor; API her istekte bellekten hesaplıyor, `Cache-Control: no-store`.
"Üretim türü ters", "planlanan tarih yanlış", "teslim edilmemiş teslim görünür" belirtileri gerçek dosyanın 44 kaydında karşılaştırmalı testle görülmedi; EÜM başlangıcı ise hiç okunamıyordu (#2).

## 2. Değişen dosyalar

- `excel_araclari.py`: satır sınıflandırma (boş / kalıntı / kanıtlı-eksik / kayıt), satır hatalarının toplanması, başlık adıyla ve gizlilik tercihiyle sütun eşleme, EÜM hafta metni çözümü, DURUM yazım varyantları ve Unicode normalizasyonu, COM snapshot'ta tüm sütunlar + gizlilik + metin koruma + hata değerleri, `kaynak_raporu`.
- `depo.py`: değişmeyen kart sürüm/damga korunur (`degismeyen`), boş metin `None`, yeni kartta tam şema, log özetine kaynak raporu.
- `app.py`: ortak import bildirimi (atlanan satır aralıkları dahil, çerez boyutu sınırlı), rapor sayaçlarında kaynakta pasif kartlar hariç.
- `templates/import_onizleme.html`: sayfa bazlı kayıt/atlanan/boş satır tablosu, atlanan satır ayrıntısı, uyarı listesi.
- `templates/yonetim.html`: yükleme açıklaması.
- `tests/test_gercek_excel.py` (yeni, 23 test; 2 tanesi COM), `tests/test_excel_com.py` (1 yeni test, 1 test yeni sözleşmeye göre güncellendi).
- `docs/EXCEL_SYNC.md`. Önceki kod: `.investigation/import-fix-20260926-before/`, diff: `.investigation/import-fix-20260926.diff`.

## 2b. Önizleme ekranı (ikinci adım)

`templates/import_onizleme.html` yeniden tasarlandı (+ `static/stil.css` bölümü, `base.html` stil sürümü):
etki kutuları, sayfa kartları, türüne göre gruplanmış uyarılar ve açıklamaları, kalıntı satır desenleri,
kart bazında "Şu an / Onaydan sonra" karşılaştırması, pasifleşecek ve yeni kart tabloları, alta yapışık onay çubuğu,
toplu pasifleşme (≥ %30) / eksik sayfa uyarısı, çift tıklama koruması. Excel önizlemede reddedilirse
sorunlar sayfada madde madde gösterilir (422). İlk yüklemede önceki ~600 satırlık alan tablosu yerine
tek satırda açılan 44 kartlık liste var. 375 px genişlikte yatay taşma olmadığı ölçüldü.
`depo.excel_import_onizle` değişiklikleri tür (yeni/güncelleme/pasif/geri) ve kart özetiyle döndürür;
parser uyarıları türlü üretir (`uyari_gruplari`, `atlanan_gruplari`).

## 2c. Sahada yaşanan hatalar (1.1 ve 1.2) — üçüncü adım

- **Eski kök nedenler çözülmüş:** 5 MAKİNE + 2 ELDE aynı Talep/Stok, sıra değişikliği, araya satır ekleme,
  makine satırının silinmesi; Excel'de TESLİM EDİLDİ (tarihli/tarihsiz) kartın monitörde görünmemesi;
  aynı Talep+Stok'un 3 satırının ayrı kart olması `tests/test_onceki_hatalar.py` ile ekran/API düzeyinde PASS.
- **Hâlâ aynı belirtiyi üreten ve bu adımda kapatılan yollar:**
  - NO başka talebe verilince yeni talep eski kartın TESLİM EDİLDİ durumu, adedi ve notuyla görünüyordu
    (yeniden üretildi). Artık eski kart geçmişiyle ayrılır (pasif), yeni talep için temiz kart açılır;
    önizlemede ayrı grup olarak gösterilir. Yalnız Stok/tarih/adet değişirse kart yerinde güncellenir.
  - Tanınmayan DURUM (ör. `TESLİM EDİLDİ (KISMİ)`) kartı eski DİZGİDE durumunda bırakıp monitörde
    "PLANDA" gösteriyordu (yeniden üretildi). Artık import durur; MALZEME TEDARİK / PDGM ÖNERİ bilinen
    durumsuz değerler olarak kalır.
  - Yinelenen "Planlanan Teslim T." sütunlarında görünür olan okunuyordu; üretim sütunu gizlenip sipariş
    sütunu gösterilirse tarihler yanlış bloktan gelirdi (yeniden üretildi). Artık DURUM'a yakın üretim
    sütunu okunur; gizleme/gösterme sonucu değiştirmez.

## 2d. Önizlemede "Tamamlanan adedi sıfırla" seçeneği — dördüncü adım

Mevcut kural: PLANA ALINDI / HAZIR olan kartta tamamlanan adet zorunlu 0, TESLİM EDİLDİ'de zorunlu toplam;
yalnız DİZGİDE kalan kartta operatörün girdiği adet korunur (ör. TESLİM → DİZGİDE dönen kart 5/5 kalır).
Önizlemede, bu importta güncellenen ve sonunda DİZGİDE kalan, tamamlanan adedi > 0 olan her kartın altında
varsayılan işaretsiz "Tamamlanan adedi sıfırla" kutusu, grup başında açıklama ve "Tümünü işaretle", onay
çubuğunda sayaç var. Seçim onay formuyla gönderilir; sunucu yalnız o önizlemede sunulan kartları kabul eder ve
depo uygulama anında uygunluğu yeniden denetler (uygun olmayan seçim tüm importu iptal eder). Sıfırlama
tamamlanan adedi 0 yapar, bitiş zamanını temizler; not, operatör ve başlama bilgisi korunur; audit logda
eski/yeni değerle görünür. Doğrudan (önizlemesiz) import davranışı değişmedi.

## 2e. Excel uygulamanın gerisinde kalınca admin kararı — beşinci adım

Sorun (analiz raporu #2): Excel güncellenmeden yapılan import, operatörün uygulamada ilerlettiği kartı
geri alıyordu (6/10 DİZGİDE → PLANA 0/10; bugün teslim edilen → DİZGİDE, teslim tarihi silinir).
Önizlemede artık en üstte "Karar gerekiyor" bölümü var: her kart için uygulamadaki ve Excel'deki durum,
neden orada olduğu, Excel ile uygulama durumu arasındaki seçenekler (önerilen işaretli) ve yalnız DİZGİDE
seçilince açılan "Tamamlanan adedi sıfırla" kutusu. Öneri: kart uygulamada ilerletildiyse (başlama/teslim
zamanı var) uygulamadaki durum, durum yalnız Excel'den geldiyse (Excel'de kasıtlı geri alma) Excel.
Toplu "Hepsinde uygulamadakini koru" / "Hepsini Excel'e göre ayarla" düğmeleri ve onay çubuğunda özet var.
Seçim eksik, aralık dışı, gerilemeyen karta ait ya da PLANA seçilen kartta sıfırlama ise tüm aktarım iptal
edilir. Her karar işlem loguna "EXCEL GERİDE: DURUM KARARI" olarak yazılır. Önizlemesiz doğrudan import
eskisi gibi Excel'i uygular. Gerçek sunucu + gerçek Excel ile denendi: operatörün 2/5 DİZGİDE ve TESLİM
ettiği kartlar korundu, Excel'de kasıtlı geri alınan kart Excel'e göre güncellendi.

## 3. EÜM entegrasyonu

Sayfa, model, operatör tipi (`eum_dizgi`), filtreler ve rozetler güncel kodda zaten vardı; kırık olan parser katmanıydı.
EÜM kartları `source_sheet=EUM` ile ayrı kimlik uzayında (Talep 1826751 hem MAKİNE NO 11 hem EÜM NO 8'de; iki ayrı kart).
`PLANA ALINDI` / `ÜRETİM PLANA ALINDI` → PLANA ALINDI (tamamlanan 0), `ÜRETİM DEVAM EDİYOR` → DİZGİDE, `PDGM ÖNERİ` → durum boş (mevcut kural).
Plan başlangıcı hafta metninin Pazartesi'si; güncelleme/silme kuralları diğer sayfalarla aynı.

## 4. Test sonuçları (hepsi otomatik, gerçek dosya ve kopyaları)

| Test | Sonuç |
|---|---|
| Orijinal Excel import (44 kayıt, alan alan bağımsız referansla) | PASS |
| Tamamen boş satır (boş + yalnız boşluk) | PASS |
| Gizli boş satır | PASS |
| Gizli dolu satır (mevcut + yeni) | PASS |
| Gizli/boş sütun (sütun ekleme + DURUM/teslim/stok/adet gizleme) | PASS |
| Aynı başlıklı iki sütun ikisi de gizli/görünür → açık hata, veri korunur | PASS |
| Gerçekten eksik zorunlu alan (3 hata birlikte, bellek/disk/yedek değişmez) | PASS |
| Yeni EÜM kayıtları ve PLANA/ÜRETİM PLANA ALINDI varyantları | PASS |
| Tarih ve durum değişikliği (aynı ID, API ve yeniden başlatma) | PASS |
| Yeni kayıt ekleme | PASS |
| Kayıt silme (pasifleşme, satır geri gelince aynı kart) | PASS |
| Kalıntı bırakılarak silinen kayıt (önizlemede raporlanır) | PASS |
| Aynı Excel'i tekrar yükleme (yeniden başlatma öncesi/sonrası 0 değişiklik) | PASS |
| Hatalı Excel / iş kuralı hatası → bellek + disk geri alma | PASS |
| Flask önizleme → onay → API/teslim/filtre tutarlılığı | PASS |
| Gerçek Excel COM: orijinal + değişiklik + tekrar, `=NA()`, metin koruma | PASS |
| Önizleme arayüzü: gruplar, karşılaştırma, toplu pasifleşme uyarısı, 422 hata ekranı | PASS |
| Sahadaki hatalar: 5 MAKİNE + 2 ELDE, teslim/PLANDA, çoklu Talep+Stok, NO yeniden kullanımı, tanınmayan DURUM | PASS |
| Tamamlanan adedi sıfırla: uygunluk, seçilmezse koruma, geçersiz seçimde iptal, arayüz akışı | PASS |
| Excel geride: öneriler, seçimlerin uygulanması, geçersiz seçimde iptal, doğrudan import uyumu, arayüz | PASS |
| Tüm paket: 96 test (`PDGM_TEST_EXCEL_COM=1`) | 96/96 PASS |
| İzole sunucu + gerçek HTTP + gerçek COM uçtan uca (tarayıcıda pano kontrolü) | PASS |
| Node JS testleri (`monitor_rotation.cjs`, `frontend_contract.cjs`) | ÇALIŞTIRILMADI (Node kurulu değil; ilgili dosyalar değişmedi) |

## 5. Sayfa bazlı sayılar (orijinal dosya)

| Sayfa | Talep NO + Stok dolu | Kart sayılmayan kalıntı satır | Boş satır | DB'ye yazılan | Durum dağılımı |
|---|---|---|---|---|---|
| MAKİNE | 15 | 52 (211–225, 229–253, 255–259, 262, 265–268, 270–271) | 219 | 15 | TESLİM 8, DİZGİDE 2, PLANA 1, boş 4 |
| ELDE DİZGİ | 20 | 0 | 50 | 20 | TESLİM 13, DİZGİDE 2, PLANA 2, boş 3 (MALZEME TEDARİK) |
| EÜM | 9 | 4 (10, 205, 208, 211) | 230 | 9 | TESLİM 6, DİZGİDE 1, boş 2 (PDGM ÖNERİ) |
| Toplam | 44 | 56 | — | 44 | Operasyon ekranında 35 |

## 6. Kalan riskler

- Kimlik NO'ya dayanır. NO farklı bir Talep NO'ya verilirse artık kart ayrılır; ancak aynı Talep NO'nun başka bir satırına (başka parti) aynı NO verilirse sistem bunu aynı kart sayar. Talep NO'daki bir yazım düzeltmesi de kartı ayırır (eski kart ve notu Yönetim'de pasif olarak kalır).
- Talep NO, adet, DURUM ve tarihleri temizlenip NO/Stok bırakılan satır "silinmiş" sayılır (yumuşak pasifleşme, geri alınabilir, önizlemede listelenir).
- `MALZEME TEDARİK` ve `PDGM ÖNERİ` iş kuralı gereği durum boş kalır; 5 kart Yönetim'den durum atanana kadar operasyon ekranında görünmez.
- EÜM hafta metninde gün.ay yoksa ve satırda tarih yoksa yıl bugüne göre seçilir (uyarı üretilir).
- Excel yüklemeden önce kaydedilmiş olmalı; XLSX çok dosyalı gerçek transaction değildir (önceki raporlarla aynı sınır).
