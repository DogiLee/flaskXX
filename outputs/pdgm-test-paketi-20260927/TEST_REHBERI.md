# PDGM elle test rehberi

Bu paket 27.09.2026'da gerçek örnek dosyanızdan (`PDGM_Kart_dizgi_Talepleri_Üretim_Takvimi.xlsx`) üretildi.
- Sayfalar, başlıklar, gizli sütunlar ve satırlar ile kalıntı satırlar olduğu gibi korundu.
- Üzerine tüm özellikleri deneyen yeni satırlar eklendi.
- MAKİNE'ye 30 kayıt (NO 16–45), ELDE DİZGİ'ye 15 kayıt (NO 21–35), EÜM'e 14 kayıt (NO 10–23) eklendi.

**Neler otomatik doğrulandı:**
- 1–5. bölümlerdeki sayılar, kart durumları, rozetler ve hata mesajları `tests/test_test_paketi.py` ile doğrulandı. Test aynı dosyaları, bu rehberdeki sırayla ve operatör işlemleriyle birlikte, geçici bir veritabanında çalıştırır.
- Ekran metinleri şablonlardan alındı.
- 6. bölüm, mevcut özelliklerin elle kontrolüdür ve bu pakete özel otomatik testi yoktur.

> Tarihler üretildiği güne göre hesaplanır. "BUGÜN BAŞLAMALI", "SON GÜN" gibi rozetler yalnız o gün geçerlidir.
> Başka bir gün test edecekseniz önce paketi yeniden üretin:
> `.venv\Scripts\python.exe tools\test_paketi_uret.py`

Kısaltmalar: **M** = MAKİNE sayfası, **E** = ELDE DİZGİ, **Ü** = EÜM. "M 20", MAKİNE sayfasında NO sütunu 20 olan kayıttır.
Talep numaraları: MAKİNE 19000xx, ELDE 19100xx, EÜM 19200xx (xx = NO). Pano'da talep numarasıyla arayabilirsiniz.

## 0. Hazırlık (üretim verinize dokunmadan)

- [ ] `TEST_ORTAMI_BASLAT.bat` dosyasını çift tıklayın.
  - Uygulama `..\gokberk_flask_TEST` klasörüne kopyalanır ve **boş veriyle** `http://127.0.0.1:5002` adresinde açılır.
  - Üretim sunucusu (5001) ve üretimdeki `data\` klasörü etkilenmez.
  - Kullanıcı hesapları ve parolalar üretimdekiyle aynıdır.
- [ ] Bir tarayıcıda `admin` ile giriş yapın.
- [ ] İkinci bir tarayıcıda veya gizli pencerede operatör hesaplarıyla giriş yapın: `makine1`, `elle1`, `eum1`.
- Baştan başlamak için:
  1. Test sunucusunun penceresini kapatın.
  2. `gokberk_flask_TEST\data` klasörünü silin.
  3. Bat dosyasını yeniden çalıştırın.

## 1. İlk yükleme: `01_ILK_YUKLEME.xlsx`

Yönetim → **Plan Excel'ini Aktar** → dosyayı seçin → **Etkiyi Önizle**.

Önizlemede:
- [ ] "103 kayıt okundu" yazar.
- [ ] Kutular: Yeni kart **103**, Güncellenecek 0, Değişmeyecek 0, Pasifleşecek 0, Veri uyarısı **16**.
- [ ] Sayfa kartları:
  - **MAKİNE 45** kayıt (1 tanesi gizli satırda); 54 satır kart sayılmadı (liste `…270–271, 313–314` ile biter).
  - **ELDE DİZGİ 35** kayıt.
  - **EÜM 23** kayıt; 4 satır kart sayılmadı: `10, 205, 208, 211`.
- [ ] Kalıntı satırlar listesinde MAKİNE satır 313 (yalnız Kart Stok No) ve 314 (NO + dizgi sorumlusu) görünür.
- [ ] Uyarı grupları:

| Adet | Uyarı | Yer |
|---|---|---|
| 6 | DURUM boş | M satır 273–275, 277, 293 · E satır 86 |
| 4 | DURUM "MALZEME TEDARİK": iş akışı durumu atanmaz | E satır 67, 69, 72, 77 |
| 3 | DURUM "PDGM ÖNERİ": iş akışı durumu atanmaz | Ü satır 244–245, 252 |
| 1 | TESLİM EDİLDİ ama Gerçekleşen Teslim T. boş | M satır 292 (M 28) |
| 1 | Gerçekleşen Teslim T. dolu ama DURUM teslim değil | M satır 294 (M 30) |
| 1 | Planlanan başlangıç tam çözülemedi | Ü satır 254 (Ü 18, hafta metni "belirsiz") |

- [ ] **Değişiklikleri uygula** → Yönetim ekranında yeşil bildirim: "Excel aktarıldı · 103 satır · 103 yeni …".

Uyguladıktan sonra:
- [ ] Pano ve Operatör ekranında durumu olan **90** kart var. Durumu eksik 13 kart bu ekranlarda görünmez.
- [ ] Pano'da Dizgide **24** kart, Plana Alınan İşler **29** kart var.
- [ ] Monitör'de yalnız MAKİNE kartları var: Dizgide **12**, Plana Alınan İşler **14**.
- [ ] Rozetler:

| Kart | Beklenen rozet |
|---|---|
| M 17 | BUGÜN BAŞLAMALI |
| M 18 | BAŞLAMADI (+5 gün) |
| M 20 | PLANINDA (5 gün var); monitörde "PLANDA" |
| M 21 / M 22 | SON GÜN / SON 1 GÜN |
| M 23 | SÜRE AŞILDI (3 gün) |
| M 25 / M 26 | ZAMANINDA TESLİM / GEÇ TESLİM (+3 gün) |
| M 28 | TESLİM EDİLDİ (gerçekleşen tarih boş olduğu için sapma yok) |
| Ü 11 | BAŞLAMADI (+6 gün) |

- [ ] **Geçmiş hata 1.1** (7 adetlik talep: 5 adet makine, 2 adet elle). Talep **1900034** için iki ayrı kart olmalı:
  - M 34: 5 adet, **DİZGİDE**, monitörde görünür.
  - E 28: 2 adet, **TESLİM EDİLDİ**, Teslim Edilenler'de; monitörde görünmez.
- [ ] **Geçmiş hata 1.2** (aynı talep ve stok birden fazla satırda). Talep **1900031** + AD-T031-0001 için **3 ayrı kart** olmalı:
  - 3 adet, TESLİM EDİLDİ
  - 4 adet, DİZGİDE
  - 5 adet, PLANA ALINDI
- [ ] Talep 1900016 için M 16 (makine) ve Ü 21 (stok AD-T016-EUM) **ayrı** kartlardır.
- [ ] Biçim farkları doğru okunur:
  - M 19: DURUM küçük harfle "plana alındı" → PLANA ALINDI.
  - M 27: "Teslim Edildi " (sonunda boşluk) → TESLİM EDİLDİ.
  - M 35: "1.500 ADET" → **1500** adet.
  - M 36: adet hücresinde yalnız sayı var → **12** adet.
  - M 37: gizli satırda olduğu halde kart oluşur.
  - M 38: plan teslim tarihi metin olarak yazılmış ("gg.aa.yyyy") ve okunur.
  - E 35: DURUM küçük harfle "dizgide" → DİZGİDE.
- [ ] EÜM durumları:
  - Ü 10 "PLANA ALINDI", Ü 11 "ÜRETİM PLANA ALINDI", Ü 12 "üretim  plana alındı " ve Ü 14 "EMTD PLANLANDI" → **PLANA ALINDI**.
  - Ü 13 "ÜRETİM DEVAM EDİYOR" → **DİZGİDE**.
- [ ] EÜM plan başlangıcı hafta metninin Pazartesi günüdür:
  - Ü 17'de yalnız "NN. hafta" yazıyor; Ü 19'da hafta sütununda tarih var. İkisi de Pazartesi'ye çevrilir.
  - Ü 18'de başlangıç boş kalır.
- [ ] Malzeme: E 23 "DİZGİDE VE MALZEME BEKLENİYOR" ve E 24 "…BEKLİYOR" → DİZGİDE ve **Malzeme Bekliyor** rozeti.
- [ ] Yönetim → **Durumu Eksik Kartlar** listesinde **13** kart var:
  - MAKİNE 5, aralarında 1900029.
  - Elle 5 (4'ü MALZEME TEDARİK, aralarında 1910025).
  - EÜM 3 (PDGM ÖNERİ, aralarında 1920016).
- [ ] Yönetim → **Rapor İndir**: Excel dosyası iner; kart listesinde 103 kart vardır.

## 2. Operatör işlemleri ve güncelleme: `02_GUNCELLEME.xlsx`

Önce Operatör ekranında şu işlemleri yapın:

| Hesap | Kart | İşlem |
|---|---|---|
| makine1 | M 20 (1900020) | Üretilen Adedi Gir: 3 |
| makine1 | M 24 (1900024) | Üretilen Adedi Gir: 12 → **Teslim Et** |
| makine1 | M 39 (1900039) | Üretilen Adedi Gir: 2 → **Teslim Et** |
| makine1 | M 40 (1900040) | **Dizgiye Al** (10 adet) → Üretilen Adedi Gir: 4 |
| makine1 | M 42 (1900042) | Not Ekle: "Eski talebin notu" |
| elle1 | E 22 (1910022) | **Dizgiye Al** (3 adet) |
| elle1 | E 30 (1910030) | Üretilen Adedi Gir: 2 |
| eum1 | Ü 11 (1920011) | **Dizgiye Al** (3 adet) → Üretilen Adedi Gir: 1 |

- [ ] `makine1` ile Elle ve EÜM kartlarında işlem düğmesi çıkmaz; yerinde "… operatörüne ait" etiketi görünür. Yalnız **Not Ekle** kullanılabilir. `elle1` ve `eum1` için de aynısı geçerlidir.

Sonra admin ile `02_GUNCELLEME.xlsx` → **Etkiyi Önizle**. Bu dosyada iki şey bilerek değiştirildi; ikisi de sonucu bozmamalı:
- MAKİNE sayfasında **DURUM sütunu gizli**.
- Gizli sipariş bloğundaki ikinci "Planlanan Teslim T." sütunu (O) **görünür** ve içinde farklı tarihler var.

- [ ] Üstte sarı bant: "**6 kartta** Excel uygulamanın gerisinde."
- [ ] Kutular: Yeni **5**, Güncellenecek **11**, Değişmeyecek **88**, Pasifleşecek **4**, Veri uyarısı **14**.
- [ ] **KARAR GEREKİYOR · 6 KART** bölümündeki önerilen seçimler:

| Kart | Uygulamada | Excel'de | Önerilen | Gerekçe metni |
|---|---|---|---|---|
| M 24 | TESLİM EDİLDİ | DİZGİDE | TESLİM EDİLDİ | Kart uygulamada ilerletilmiş… |
| M 39 | TESLİM EDİLDİ | DİZGİDE | TESLİM EDİLDİ | Kart uygulamada ilerletilmiş… |
| M 40 | DİZGİDE 4/10 | PLANA ALINDI | DİZGİDE | Kart uygulamada ilerletilmiş… |
| M 41 | TESLİM EDİLDİ | DİZGİDE | **DİZGİDE** | …Excel'de geri alınmış görünüyor. |
| E 22 | DİZGİDE | PLANA ALINDI (Excel metni: DİZGİ İÇİN BEKLİYOR) | DİZGİDE | Kart uygulamada ilerletilmiş… |
| Ü 11 | DİZGİDE 1/3 | PLANA ALINDI (Excel metni: ÜRETİM PLANA ALINDI) | DİZGİDE | Kart uygulamada ilerletilmiş… |

- [ ] M 40'ta durumu **PLANA ALINDI** yapın:
  - Açıklama "Kart plana döner: tamamlanan adet 0 olur…" olarak değişir.
  - "Tamamlanan adedi sıfırla" kutusu devre dışı kalır.
  - Sonra durumu tekrar **DİZGİDE** yapın.
- [ ] **Hepsini Excel'e göre ayarla** ve **Hepsinde uygulamadakini koru** düğmelerini deneyin; alttaki onay çubuğunda sayaç değişmeli. Sonra seçimleri tablodaki önerilere geri getirin.
- [ ] M 41'de **Tamamlanan adedi sıfırla** kutusunu işaretleyin ("Şu an 2 / 2").
- [ ] **Güncellenecek** grubunda (karar kartları hariç) 10 kart var:

| Kart | Değişiklik |
|---|---|
| M 16 | Plan teslim tarihi |
| M 18 | PLANA → DİZGİDE |
| M 20 | Plan teslim tarihi |
| M 23 | DİZGİDE → TESLİM EDİLDİ |
| M 36 | Toplam adet 12 → 15 |
| M 37 | PLANA → DİZGİDE (gizli satırda) |
| E 21 | PLANA → DİZGİDE |
| E 25 | Durum eksik → PLANA ALINDI |
| Ü 10 | PLANA → DİZGİDE |
| Ü 16 | Durum eksik → PLANA ALINDI |

  Tarihler W sütunundan okunur; görünür hale gelen O sütunundaki farklı tarihler okunmaz.
- [ ] M 20 kartının altında "Tamamlanan adedi sıfırla" kutusu var ("Şu an 3 / 10"). Kutuyu **işaretleyin**.
- [ ] **NO'su başka talebe verilmiş** (1 kart): M 42'nin eski talebi 1900042, yenisi **1909042 · AD-T042-YENI**.
- [ ] **Pasifleşecek** (3 kart): M 43 (satır tamamen silindi), M 44 (satırda yalnız NO, sorumlu ve stok kaldı), E 33.
- [ ] Kart sayılmayan satırlar listesine MAKİNE satır **310** (M 44'ün kalıntısı) eklendi.
- [ ] **Yeni** (5 kart): M 42 (1909042), M 46, M 47, M 48, Ü 26.
- [ ] Onay çubuğunda şu yazar: "6 kart için durum kararı: **5** kartta uygulamadaki durum korunacak · **2** kartta tamamlanan adet sıfırlanacak."
- [ ] **Değişiklikleri uygula**. Bildirim şunları içermeli:
  - "1 kartın NO'su Excel'de başka talebe verilmiş…"
  - "6 kartta Excel uygulamanın gerisindeydi; 5 kartta uygulamadaki durum korundu."
  - "2 kartın tamamlanan adedi seçiminizle sıfırlandı."

Uyguladıktan sonra beklenen durumlar:

| Kart | Durum | Tamamlanan / Toplam |
|---|---|---|
| M 16 | PLANA ALINDI (yeni plan teslim tarihiyle) | 0 / 4 |
| M 18, M 37 | DİZGİDE | 0 |
| M 20 | DİZGİDE | **0** / 10 (sıfırlandı) |
| M 23 | TESLİM EDİLDİ, GEÇ TESLİM (+2 gün) | 5 / 5 |
| M 24, M 39 | TESLİM EDİLDİ, ZAMANINDA TESLİM (uygulamadaki bugünkü teslim tarihi korundu) | 12/12, 2/2 |
| M 40 | DİZGİDE | 4 / 10 (korundu) |
| M 41 | DİZGİDE | **0** / 2 (sıfırlandı) |
| M 36 | PLANA ALINDI | 0 / **15** |
| M 42 | yeni kart: PLANA ALINDI, talep 1909042 | 0 / 5 |
| E 21, E 22 | DİZGİDE | 0 / 2, 0 / 3 |
| E 25 | PLANA ALINDI | 0 / 5 |
| Ü 10 | DİZGİDE | 0 / 4 |
| Ü 11 | DİZGİDE | 1 / 3 (korundu) |
| Ü 16 | PLANA ALINDI | 0 / 4 |

- [ ] Yönetim → **Son Excel'de Olmayan Açık Kartlar** listesinde 4 kart var:
  - 1900042 (eski M 42, DİZGİDE 0/3)
  - 1900043, 1900044
  - 1910033
- [ ] Kartlar tablosunda bu 4 kartın altında "Kaynakta yok" etiketi var.
- [ ] Eski 1900042 kartındaki "Eski talebin notu" notu korunmuş.
- [ ] Yönetim → Son İşlemler (veya İşlem Logu dosyası):
  - 6 adet "EXCEL GERİDE: DURUM KARARI"
  - 1 adet "EXCEL NO BAŞKA TALEBE VERİLDİ"

## 3. Aynı dosya tekrar: `03_AYNI_DOSYA_TEKRAR.xlsx`

- [ ] Önizleme kutuları: Yeni 0, Güncellenecek 0, Değişmeyecek **104**, Pasifleşecek 0.
- [ ] Karar gerekiyor bölümünde **5** kart var: M 24, M 39, M 40, E 22, Ü 11.
  - Excel hâlâ geride; öneriler uygulamadaki durumu korur.
  - M 41 listede yok, çünkü artık Excel ile uygulama aynı.
- [ ] Uygulayın. Hiçbir kartın durumu veya adedi değişmez, kopya kart oluşmaz.

## 4. Excel yetişti: `04_EXCEL_YETISTI.xlsx`

Bu dosyada Excel, uygulamada yapılanlara yetiştirildi:
- M 24 ve M 39 bugün tarihli TESLİM EDİLDİ.
- M 40 ve E 22 DİZGİDE.
- Ü 11 "ÜRETİM DEVAM EDİYOR".

- [ ] Önizlemede **Karar gerekiyor bölümü yok**.
- [ ] Güncellenecek **5** kart: M 24, M 39, M 40, E 22, Ü 11. Bunlarda yalnız "Excel Durumu" metni değişir.
- [ ] Uygulayın.

## 5. Reddedilmesi gereken dosyalar (hiçbiri kayıt değiştirmez)

Her dosyayı **Etkiyi Önizle** ile deneyin. Kırmızı **"Excel kabul edilmedi"** sayfası ve madde madde sorun listesi gelmeli.

| Dosya | Beklenen sorun (satır 279 = M 16) |
|---|---|
| H01_EKSIK_TALEP_NO | MAKİNE satır 279: Talep NO boş, ancak satırda kayıt verisi var (…). Eksik satır silinmiş kabul edilmedi. |
| H02_TANINMAYAN_DURUM | MAKİNE satır 279: DURUM 'TESLİM EDİLDİ (KISMİ)' tanınmıyor… Kabul edilen değerler listelenir. |
| H03_TEKRARLANAN_NO | MAKİNE satır 280: Tekrarlanan kaynak kimliği NO:16 (ilk kullanım satır 279). |
| H04_GECERSIZ_ADET | MAKİNE satır 279: Üretim adedi tek bir sayı içermeli: 'on adet' |
| H05_GECERSIZ_TARIH | MAKİNE satır 279: Planlanan Teslim Tarihi okunamadı. Gelen değer: '31.02.2026' |
| H06_BASLAMA_TESLIMDEN_SONRA | MAKİNE satır 279: Dizgi Başlama Tarihi (…) Planlanan Teslim Tarihinden (…) sonra olamaz. |
| H07_UC_HATA_BIRDEN | "**3 sorun bulundu**, hiçbir kayıt değiştirilmedi": MAKİNE satır 279 (Talep NO) · ELDE DİZGİ satır 73 (adet 'on adet') · EÜM satır 246 (DURUM 'Bitti') |
| H08_MAKINE_SAYFASI_YOK | 'MAKİNE' sayfası bulunamadı. Dosyadaki sayfalar: URETIM, ELDE DİZGİ, EÜM |
| H09_IKI_NO_SUTUNU | MAKİNE: 'NO' için birden fazla sütun bulundu (C, V); hangisinin okunacağı belirsiz. |

- [ ] Her denemeden sonra Pano'da hiçbir değişiklik olmamalı.

**Uyarı veren dosyalar** (yalnız önizleyin, sonra **Vazgeç**'e basın):
- [ ] `U01_EUM_SAYFASI_YOK`:
  - Kırmızı bant: "EÜM sayfası dosyada yok; bu sayfadan gelmiş kartların hepsi pasifleşir."
  - Pasifleşecek **24**.
- [ ] `U02_ELDE_SAYFASI_BOSALTILDI`:
  - Kırmızı bant: "Sistemde Excel'den gelen 104 aktif karttan 34 tanesi pasifleşecek."
  - Onay düğmesi kırmızı ve üzerinde **"Yine de uygula"** yazar.

## 6. Diğer ekranlar ve işlemler (elle kontrol)

- [ ] **Malzeme onayı:**
  1. Yönetim → Kartlar → 1910029 (E 29, PLANA ALINDI) → Düzenle → **Malzeme Bekliyor** kutusunu işaretleyip kaydedin.
  2. `elle1` ile bu kartta **Dizgiye Al**'a basın. Malzemenin tedarik edildiğini onaylamanızı isteyen soru gelmeli.
  3. Onaylayın. Kart DİZGİDE olmalı ve işlem kaydında "MALZEME TEDARİK ONAYLANDI" görünmeli.
- [ ] **Durum ata:** Yönetim → Durumu Eksik Kartlar → 1900029 → **Durum Ata** → PLANA ALINDI. Kart listeden çıkar ve Pano'da görünür.
- [ ] **Gizle / geri getir:**
  1. Kartlar tablosunda bir kartı **Gizle**. Kart Pano'dan kalkar.
  2. **Gizlenen Kartlar** → **Geri Getir**. Kart geri gelir.
- [ ] **Yeni kart:** Yönetim → **+ Yeni Kart** (ör. Talep 1999001). Kart PLANA ALINDI olarak oluşur ve Pano'da görünür.
- [ ] **Eşzamanlı düzenleme:**
  1. Admin bir kartın Düzenle formunu açsın.
  2. Başka bir pencerede operatör aynı karta not eklesin.
  3. Admin kaydettiğinde "Kart siz formu açtıktan sonra değişti. Kaydedilmedi." uyarısı gelmeli ve operatörün notu kaybolmamalı.
- [ ] **Yedekten geri yükle:**
  - Her importtan önce bir "Anlık · import oncesi" yedeği alınır.
  - Listeden, 4. adımı uygulamadan hemen önce alınmış yedeği saatine bakarak seçin ve **Geri Yükle**'ye basın. Kart verisi o ana döner.
  - Bu sırada mevcut durum da ayrıca yedeklenir.
- [ ] **Kayıt dosyaları:** Yönetim → Kayıt Dosyaları → Kartlar Excel, İşlem Logu ve Yükleme Geçmişi dosyaları iner.
- [ ] **Monitör:** sayfa sayısı birden fazlaysa sayfalar kendiliğinden döner.

## Bitirince

- Test sunucusunun penceresini kapatın.
- `gokberk_flask_TEST` klasörünü silebilirsiniz; üretimdeki `data\` klasörünüz bu testlerden etkilenmez.
