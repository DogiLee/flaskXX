# PDGM İş Takip — Frontend İncelemesi (27.09.2026)

> **Durum:** 26 maddenin tamamı aynı gün düzeltildi. Ayrıntı: [DUZELTMELER.md](DUZELTMELER.md).

Kapsam: `templates/` altındaki 9 şablon ve satır içi JS'leri (~1.000 satır), `static/stil.css` (2.809 satır).
Kod değiştirilmedi. Uygulamanın bir kopyası, `data/` kopyası ve test kullanıcılarıyla scratchpad'de
`127.0.0.1:5055` üzerinde çalıştırıldı. 5001'deki canlı sunucuya ve üretim verisine dokunulmadı.
Ekranlar 1920×1080 (monitör), 1440×900 (masaüstü) ve 375×812 (telefon) boyutlarında gezildi.

**Kanıt:** *Doğrulandı* = tarayıcıda denendi/ölçüldü. *Koddan* = kod okunarak tespit edildi, denenmedi.
Önceki rapordaki (`outputs/proje-analizi-20260927/ANALIZ_RAPORU.md`) maddeler burada tekrarlanmadı,
yalnız "önceki #N" diye anıldı.

## Özet

| # | Önem | Bulgu | Kanıt |
|---|---|---|---|
| 1 | Yüksek | "Durum Ata" ile kaydetmek kartın **Malzeme Bekliyor** işaretini sessizce siliyor | Doğrulandı: kart 32, 1 → 0 |
| 2 | Yüksek | Operatör işlemlerinden sonra başarı mesajı hiç görünmüyor | Doğrulandı |
| 3 | Yüksek | Paylaşımlı hesaplarda operatörün adı her işlemde sıfırlanıyor, Dizgiye Al / Adet Bitir / Teslim'de hiç sorulmuyor | Doğrulandı + koddan |
| 4 | Yüksek | Tablet ve telefonda (≤820px) arama kutusu 340px yüksekliğe uzuyor | Doğrulandı (Pano, Operatör) |
| 5 | Yüksek | Monitör uzaktan okunamıyor: 9–11px yazılar, kart azken ekran boş | Doğrulandı (1920×1080) |
| 6 | Yüksek | Monitör, sunucuya ulaşamadığı an tarayıcı hata sayfasında kalıyor | Koddan |
| 7 | Orta | Yönetim telefonda yatay taşıyor (375px → 701px) | Doğrulandı |
| 8 | Orta | Klavye odak halkası görünmüyor (açık zeminde 1,36:1, koyu başlıkta sıfır) | Doğrulandı |
| 9 | Orta | Yönetim'de filtre, arama ve kaydırma her kayıttan sonra sıfırlanıyor | Koddan |
| 10 | Orta | Hata mesajları 3,2 sn'de kayboluyor ve kapatılamıyor | Koddan |
| 11 | Orta | Pano "canlı" görünüyor ama kendini yenilemiyor; aynı ekranda iki farklı teslim sayısı | Doğrulandı |
| 12 | Orta | Pano'da teslim tablosu iki kez yükleniyor (sunucu + API), açılışta titriyor | Koddan |
| 13 | Orta | Operatör ekranı tüm teslim edilmiş kartları DOM'a basıyor | Doğrulandı: 36 kartın 26'sı gizli |
| 14 | Orta | Buton adları ve onay akışı tutarsız ("Teslim Edildi", "Adet Bitir", `confirm()`) | Doğrulandı |
| 15 | Orta | Giriş sayfasında bozuk HTML: gözlemci formu giriş formunun içine gömülüyor | Doğrulandı |
| 16 | Düşük | CSS kaskad hatası: 821–1050px'te monitör tek kolona inmiyor | Koddan |
| 17–26 | Düşük | Tekrarlı/ölü CSS, sayı biçimi, küçük yazılar, satır içi JS, önbellek sürümü vb. | Aşağıda |

## Yüksek

### 1. "Durum Ata" Malzeme Bekliyor işaretini siliyor
- **Ne oluyor:** Yönetim > "Durumu Eksik Kartlar" listesindeki **Durum Ata** butonunda `data-malzeme`
  özniteliği yok (`templates/yonetim.html:91-97`). Ana tablodaki **Düzenle** butonunda var (`:408`).
  Dialog açılırken kutu `buton.dataset.malzeme === "1"` ile doldurulduğu için (`:698`) her zaman boş
  geliyor. Kaydedilince `malzeme_bekliyor: false` gönderiliyor, sunucu da bunu 0 olarak yazıyor
  (`depo.py:1945`).
- **Deneme:** Kopyada kart 32'ye "Durum Ata → PLANA ALINDI → Kaydet" yapıldı. Malzeme Bekliyor 1'den 0'a düştü.
  Operatör artık malzeme onayı istenmeden kartı dizgiye alabiliyor.
- **Üretimde etkilenen:** 32 ve 33 numaralı kartlar hem durumu eksik hem de malzeme bekliyor işaretli.
- **Öneri:** Durum Ata butonuna `data-malzeme="{{ '1' if k.malzeme_bekliyor else '0' }}"` ekleyin. Kalıcı
  çözüm: iki butonu tek bir Jinja makrosundan üretin ki öznitelikler bir daha ayrışmasın. Sunucuya yalnız
  değişen alanları göndermek de benzer hataları önler.

### 2. Operatör işlem sonrası geri bildirim almıyor
- **Ne oluyor:** Dizgiye Al, Adet Bitir, Teslim ve Not işlemlerinden sonra `toast(...)` çağrılıyor, ama
  350 ms sonra sayfa yenileniyor (`templates/operator.html:305-309`). Bildirim yenilemeyle birlikte siliniyor.
  Yönetim'de de aynı desen var (`yonetim.html:728, 778, 822`).
- **Deneme:** Not kaydedildikten 2 sn sonra sayfada hiç bildirim yok.
- **Etkisi:** Kullanıcı işlemin kaydedilip kaydedilmediğinden emin olamıyor ve tekrar deneyebiliyor.
  `bitir` yanıtındaki "üretim bitti, teslim edin" gibi yönlendirici mesajlar da okunmuyor.
- **Öneri:** Kısa vadede mesajı `sessionStorage`'a yazıp yenilemeden sonra gösterin. Doğru çözüm: API zaten
  güncel `kart` nesnesini döndürüyor (`app.py:877, 902, 926`). Sayfayı yenilemek yerine yalnız o kartı
  güncelleyin. Kaydırma konumu ve odak da korunur.

### 3. Operatör adı her işlemde sıfırlanıyor
- **Durum:** Operatör hesapları tür başına bir tane (`makine1`, `elle1`, `eum1`), yani büyük olasılıkla birden
  fazla kişi aynı hesabı kullanıyor. Not dialogundaki "Operatör adı" alanı bunun için konmuş.
- **Ne oluyor:**
  - Alan her sayfa yüklemesinde hesap adına dönüyor (`operator.html:219, 409`). Her işlemden sonra sayfa
    yenilendiği için kişi adını her notta yeniden yazmak zorunda kalıyor. *(Doğrulandı: "Ahmet (vardiya 2)"
    yazılıp kaydedildi, yenilemeden sonra alan "Test Admin"e döndü.)*
  - Dizgiye Al, Adet Bitir ve Teslim dialogları isim sormuyor ve `isim` göndermiyor (`operator.html:448-456,
    473-480`). Oysa backend bu üç uçta `isim` alanını zaten kabul ediyor (`app.py:875, 896, 924`). Bu
    işlemlerle eklenen notlar kişinin değil, hesabın adıyla imzalanıyor.
- **Öneri:** Cihaz başına bir "Ben: ____" seçimi (üst çubukta, `localStorage`'da saklanan) ekleyin ve tüm
  işlemlerde `isim` olarak gönderin. Not dialogundaki alan bu değerle dolsun.

### 4. Tablet ve telefonda dev arama kutusu
- **Ne oluyor:** `.operator-sayfa .arama, .panel-arama-alani { flex: 1 1 340px }` (`stil.css:892-896`).
  820px altında araç çubuğu `flex-direction: column` oluyor (`:2120-2124`). Sütun yönünde `flex-basis`
  genişlik değil yükseklik demek.
- **Ölçüm (375px):** Arama alanı 340px, input 181px yüksek. Pano ve Operatör'de aynı. 768px dikey tablette
  de geçerli.
- **Öneri:** 820px medya sorgusuna `.operator-sayfa .arama, .panel-arama-alani { flex: 0 0 auto; max-width: none; }` ekleyin.

### 5. Monitör uzaktan okunamıyor
- **Ölçüm (1920×1080):** Durum rozeti 9px, "TALEP" etiketi 9px, adet açıklaması 10px, tarih etiketi 10px,
  talep sahibi 11px, stok no 17px, talep no 19px. Atölye TV'si birkaç metreden izleniyor; bu boyutlar orada okunmaz.
- **Boş alan:** Liste sabit 2×3 ızgara (`stil.css:1218-1226`). 2 kartlık DİZGİDE bölümünde kartlar
  ekranın ~1/6'sını kaplıyor, geri kalan alan boş.
- **Başka sorunlar:**
  - Üst menü ve **Çıkış** butonu TV'de de görünüyor. Biri dokunursa monitörün oturumu kapanıyor.
  - Rotasyonda kaçıncı sayfada olunduğu gösterilmiyor. İzleyen kişi başka kart olduğunu bilmiyor.
  - Monitör yalnız **Makine** kartlarını gösteriyor (`app.py:784`), ama başlık "Üretim Monitörü".
    Pano'da 7 dizgide kart varken monitörde 2 görünüyor ve nedeni ekranda yazmıyor.
- **Öneri:** Monitör yazılarını `clamp(…, 2vw, …)` ile ekrana göre ölçekleyin (talep no ≥ 32px, rozet
  ≥ 16px). Satır sayısını kart sayısına göre verin (az kartta büyük kart). `monitor-body` için üst çubuğu
  gizleyen kiosk görünümü, "Sayfa 1/3" ve ilerleme çubuğu ekleyin. Başlığı "Makine Dizgi Monitörü" yapın.

### 6. Monitör bağlantı kopunca kendini kurtaramıyor
- **Ne oluyor:** En az 60 sn'de bir koşulsuz `window.location.reload()` (`monitor.html:272`). O anda sunucu
  yeniden başlıyorsa ya da ağ kopmuşsa tarayıcı hata sayfası açılır. Hata sayfasında JS olmadığı için
  monitör, biri elle yenileyene kadar orada kalır.
- **Öneri:** Yenilemeden önce `fetch` ile sunucuyu yoklayın. Yanıt yoksa ekranda "Bağlantı yok, yeniden
  deneniyor" bandı gösterip artan aralıklarla tekrar deneyin. Daha iyisi: `/api/veriler` JSON'unu çekip
  kartları yerinde çizin. Tam sayfa yenilemesine hiç gerek kalmaz.

## Orta

### 7. Yönetim telefonda yatay taşıyor
- **Ölçüm:** 375px ekranda sayfa 701px'e genişliyor. Neden: "Excel Geçmişi" satırlarındaki boşluksuz dosya adları
  (`20260926_214834_6ecbd7ee_PDGM_Kart_dizgi_Talepleri_Uretim_Takvimi_1.xlsx`) kırılmıyor
  (`yonetim.html:448`, `.mini-satir` `stil.css:840-866`).
- Gösterilen ad kullanıcının yüklediği değil, sunucunun iç kayıt adı (zaman + uuid öneki). Ayrıca
  `secure_filename` Türkçe harfleri atıyor ("Üretim" → "Uretim").
- **Öneri:** `.mini-satir > div:first-child { min-width: 0 } .mini-satir strong { overflow-wrap: anywhere }`.
  Yükleme kaydında özgün dosya adını da saklayıp onu gösterin.

### 8. Klavye odağı görünmüyor
- `input:focus, button:focus-visible, a:focus-visible` için `outline: none` ve
  `box-shadow: 0 0 0 3px rgba(32,58,67,.16)` kullanılıyor (`stil.css:440-448`).
- **Ölçüm:** Açık zeminde ~1,36:1 kontrast (WCAG 1.4.11 için en az 3:1 gerekli). Koyu üst çubukta odak
  "Pano" bağlantısındayken görsel bir fark yok (doğrulandı).
- **Öneri:** `:focus-visible { outline: 3px solid <belirgin renk>; outline-offset: 2px; }`. Üst çubuk ve
  koyu onay çubuğu için açık renkli bir varyant tanımlayın.

### 9. Yönetim'de çalışma bağlamı kayboluyor
- `adminAktifFiltre = "HEPSI"` sabit (`yonetim.html:610`), arama ve filtre saklanmıyor. Her
  Düzenle/Gizle/Yeni Kart işleminden sonra sayfa yenilenip en üste dönülüyor. Pano ve Operatör bunu
  `sessionStorage` ile çözmüş, Yönetim çözmemiş.
- Durum ve dizgi tipi tek filtre grubunda. "Dizgide + Elle Dizgi" birlikte seçilemiyor.
- **Öneri:** Pano/Operatör'deki iki gruplu ve durumu saklanan filtre yapısını buraya da taşıyın. Satırı
  yerinde güncellemek (#2) sorunu kökten çözer.

### 10. Hata mesajları çok hızlı kayboluyor
- Tüm bildirimler 3,2 sn sonra kapanıyor (`base.html:113-116`). Bildirim alanı `pointer-events: none`
  (`stil.css:1908`), yani fareyle tutulamıyor ya da kapatılamıyor. "Oturum sona erdi. İşlem kaydedilmedi."
  gibi kritik mesajlar kaçırılabiliyor.
- **Öneri:** Başarı mesajları kendiliğinden kapansın. Hata ve uyarılar kapat düğmesiyle kalıcı olsun,
  `role="alert"` taşısın.

### 11. Pano: "canlı" etiketi ve çelişen sayılar
- Başlıkta "CANLI ÜRETİM PANOSU" ve yeşil noktalı "Güncel pano" yazıyor (`panel.html:9-13`), ama sayfa
  kendini yenilemiyor (önceki #8). Açık kalan bir pano saatlerce eski veriyi "güncel" diye gösterebilir.
- Aynı ekranda **Teslim Edilen Farklı Stok Sayısı: 18** ile **Teslim Edilenler: 26** / **Teslim edilen iş
  emri: 26** yan yana duruyor. Birim farkı ("farklı stok" / "iş emri") ilk bakışta anlaşılmıyor.
- İstatistik kartları üzerine gelince yükseliyor (`stil.css:535-538`), yani tıklanabilir görünüyor ama tıklanmıyor.
- Operatör ekranındaki aynı kartlar dizgi tipi filtresine tepki vermiyor, Pano'dakiler veriyor.
- **Öneri:** Otomatik yenileme gelene kadar etiketi "Son yükleme: 01:15" yapın, N dakika sonra "veriler
  eskimiş olabilir" uyarısı gösterin. Ana sayıyı iş emri yapın, "18 farklı stok"u alt metne alın.
  İstatistik kartlarını durum filtresi olarak tıklanabilir yapın ya da hover efektini kaldırın.

### 12. Teslim tablosu iki kez yükleniyor
- Sunucu tüm teslimleri tabloya basıyor (`panel.html:215-230`). Açılışta `donemSec` → `donemYukle` →
  `donemTemizle` tabloyu siliyor, "Teslim verileri yükleniyor…" gösteriyor ve aynı veriyi API'den çekiyor
  (`:541, 617`). Kullanıcı her açılışta tablonun kaybolup geri geldiğini görüyor.
- **Öneri:** Varsayılan dönem ("Tümü" + "Tümü") için ilk isteği atlayın ya da sunucu tarafında yalnız
  iskelet basın.

### 13. Operatör ekranında gereksiz DOM
- `/operator` tüm kartları gönderiyor (`app.py:801`). Teslim edilmiş kartlar tam kart olarak çizilip
  gizleniyor: **36 kartın 26'sı** gizli teslim kartı. Teslimler biriktikçe sayfa ağırlaşacak.
- **Öneri:** Teslim edilenleri son 30 günle sınırlayın ya da Pano'daki 12'li sayfalamayı kullanın.

### 14. Aksiyon adları ve onay akışı
- **"Teslim Edildi"** butonu (`operator.html:126-129`) bir durum gibi okunuyor, eylem gibi değil. **"Teslim Et"** olmalı.
- **"Adet Bitir"** belirsiz. Dialogdaki **"Tamamlanan adet"** (`:190`) bu seferki adet mi, toplam mı, anlaşılmıyor.
  Öneri: "Üretilen Adedi Gir" / "Bu seferde tamamlanan adet".
- Teslim onayı tarayıcının `window.confirm`'ü ile alınıyor (`:312`). Diğer tüm işlemler özel `<dialog>`
  kullanıyor. Teslim butonu istek sürerken kilitlenmiyor, çift tıklama iki istek gönderiyor.
- Yönetim'de durum PLANA ALINDI/HAZIR seçilince tamamlanan adet sessizce 0 oluyor. Durum geri alınınca
  eski değer dönmüyor. Tamamlanan için istemcide `max = toplam` sınırı yok (`yonetim.html:736-746`).

### 15. Giriş sayfası HTML hatası
- Şifre alanının `<div class="alan">` etiketi kapanmıyor (`giris.html:25-29`). Tarayıcı gözlemci formunu
  bu div'in, yani giriş formunun **içine** yerleştiriyor (doğrulandı: `nestedInForm: true`). Şu an çalışıyor,
  ama iç içe form geçersiz HTML ve ileride form eklenirse beklenmedik gönderimlere yol açabilir.
- Aynı dosyada: stil dosyasında sürüm parametresi yok (`:7`), yani CSS değişiklikleri giriş ekranına
  önbellek yüzünden geç yansıyabilir. `autocomplete="off"` alan düzeyindeki `username` /
  `current-password` ile çelişiyor. Satır içi `style` var. Giriş CSRF'i için bkz. önceki #15.

## Düşük / bakım

16. **Kaskad hatası:** `@media (max-width:1050px) .monitor-grid {1fr}` (`stil.css:1412-1416`) kuralını, daha
    sonra gelen `@media (max-width:1100px) .monitor-grid {1fr 1fr}` (`:2087-2089`) eziyor. 821–1050px
    arasında monitör iki kolonda kalıyor, bölümlerin `min-height: 520px` değeri de işliyor.
17. **Tekrarlanan monitör stilleri:** `monitor.html:7-63` satır içi `<style>` bloğu, `stil.css:1249-1400`
    içindeki kuralları birebir tekrarlıyor ve 24 `!important` ekliyor. Blok tamamen silinebilir.
18. **Ölü CSS:** `.pano-grid`, `.plan-liste/.plan-satir/.plan-sira` (ve 3 medya sorgusundaki karşılıkları),
    `.panel-hazir-kutu`, `.panel-mini-satir`, `.panel-aktif-grid`, `.marka-logo`, `.giris-dipnot`,
    `.giris-logo img`, içeriksiz `.giris-sayfa::before`. Ayrıca yalnız ölü `ozet.html`'in kullandığı
    `.donem-grid`, `.hafta-grafik`, `.sutun*`, `.lejant`, `.ozet-*` (önceki #12). `.bos-durum` iki kez farklı
    değerlerle tanımlı (`:869`, `:2603`). Scrollbar stilleri iki tabloda kopya.
19. **Sayı biçimi:** Pano "Ort. teslim sapması **3.8 gün**" gösteriyor (`panel.html:565`). Türkçede "3,8"
    olmalı: `Number(o.ort_sapma).toLocaleString("tr-TR")`.
20. **Küçük yazılar ve kontrast:** Açıklamalar 11px, rozetler 10px, tablo başlıkları 10px, "Filtreleri
    temizle" 11px. Nötr rozetin kontrastı 4,45:1, sınırın hemen altında. Atölye ekranlarında 12–13px taban
    boyut ve 4,5:1 kontrast önerilir.
21. **Yarım kalmış tasarım token'ları:** 96 farklı hex renk, 161 doğrudan hex kullanımı, yalnız 121 `var()`.
    Aynı turuncunun dört tonu var (`#d18a00`, `#d4a017`, `#9a6500`, `#9a6400`). Durum renklerini, yazı
    boyutlarını ve boşlukları `:root` token'larına toplayın.
22. **Satır içi JS:** ~1.000 satır JS altı şablona gömülü. Filtre, `sessionStorage` ve yenileme mantığı Pano
    ile Operatör'de kopya. `tests/frontend_contract.cjs` kodu şablon metninden string indeksiyle kesiyor
    (kırılgan). `static/js/` altına ortak modüller taşınırsa kod önbelleğe alınır, test edilebilir hâle gelir
    ve CSP (önceki #20) mümkün olur. Bunun için satır içi `onsubmit` (yönetimde 3 yer) ve `style=` öznitelikleri
    de kaldırılmalı.
23. **Elle önbellek sürümü:** `base.html:8` → `stil.css?v=20260927` her değişiklikte elle güncelleniyor,
    giriş sayfasında hiç yok. Dosyanın değişiklik zamanından otomatik sürüm üreten bir `url_for` yardımcısı kullanın.
24. **Korumasız `sessionStorage`:** Pano ve Operatör `sessionStorage`'a `try/catch` olmadan erişiyor. Monitör
    koruyor. Kısıtlı tarayıcı profillerinde tüm sayfa betiği durabilir.
25. **Hareket:** Tıklanmayan öğelerde (istatistik ve operatör kartları) hover'da yükselme var.
    `prefers-reduced-motion` desteği yok.
26. **Metin tutarlılığı:**
    - Arayüzde İngilizce terimler geçiyor: "source of truth", "state", "audit", "Workflow"
      (`yonetim.html:38, 111, 457`).
    - Pano'da bölümler farklı alanlarda arıyor. Arama kutusu "operatör" diyor, ama Plana Alınanlar'da
      operatör/not, Teslimler'de PCB aranmıyor (`panel.html:84, 136, 218`).
    - Tarih girişleri tutarsız: Pano'da maskeli metin kutusu, ortadan düzenlerken imleç sona atlıyor;
      Yönetim'de yerel `type="date"`.

## İyi durumda olanlar

- Yerel `<dialog>` kullanılıyor: odak tuzağı ve Esc ile kapatma kendiliğinden çalışıyor.
- Filtrelerde `aria-pressed`, bildirimlerde `aria-live` var. `lang="tr"` doğru.
- Dinamik içerik her yerde `textContent` ile yazılıyor, `innerHTML` yok. Konsol hatası çıkmadı.
- `pdgmFetch` oturum düşmesini, yönlendirmeyi ve JSON olmayan yanıtı ayırt edip "işlem kaydedilmedi" diyor.
- Arama Türkçe büyük/küçük harfe ve Unicode normalizasyonuna duyarlı ("İ/ı" doğru eşleşiyor).
- Import önizleme ekranı çok iyi: etkiler gruplanmış, yapışkan onay çubuğu var, çift gönderim korumalı.
- Operatör kartları rol/dizgi tipine göre doğru kilitleniyor. Malzeme onayı dialog kapanınca sıfırlanıyor.

## Önerilen sıra

1. **Hızlı düzeltmeler (her biri birkaç satır):** #1 `data-malzeme`, #4 arama yüksekliği, #7 taşma,
   #8 odak halkası, #15 giriş HTML'i, #16 kaskad, #17 tekrar eden stil, #19 sayı biçimi.
2. **Operatör deneyimi:** #3 operatör adı (backend hazır), #2 yerinde güncelleme + kalıcı bildirim,
   #10 hata bildirimleri, #14 buton adları.
3. **Monitör:** #6 bağlantı kurtarma, #5 okunabilirlik ve kiosk görünümü.
4. **Yönetim ve Pano:** #9, #11, #12, #13.
5. **Bakım:** #22 JS'yi dosyalara ayırma (CSP'nin ön koşulu), #21 token'lar, #18 ölü CSS.
