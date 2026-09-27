# Frontend Düzeltmeleri (27.09.2026)

[RAPOR.md](RAPOR.md)'deki 26 maddenin uygulanışı. Test ve tarayıcı doğrulaması: [test-sonuclari.txt](test-sonuclari.txt).
Değişiklik öncesi dosyalar: `.investigation/frontend-fix-20260927-before/`, fark: `.investigation/frontend-fix-20260927.diff`.

**Sonuç:** 111/111 test geçiyor (gerçek Excel COM dahil). Önceden 96 testti; 10 yeni frontend testi eklendi
(`tests/test_frontend.py`), kalan artış aynı gün çalışan import oturumundan. JS testleri tarayıcıda geçti
(bu makinede Node yok).

## Dosya düzeni

| Önce | Sonra |
|---|---|
| ~1.000 satır JS 6 şablona gömülü | `static/js/ortak.js`, `panel.js`, `operator.js`, `monitor.js`, `yonetim.js`, `onizleme.js` |
| Tekrarlanan şablon parçaları | `templates/_makrolar.html` (arama metni, dizgi rozetleri, ilerleme çubuğu, KPI, düzenle butonu) |
| `templates/ozet.html` (hiçbir rotada kullanılmıyordu) | Silindi. `ozet_hesapla()` rapor indirmede kullanıldığı için duruyor. |
| `stil.css` 2.938 satır, 96 farklı renk, 161 doğrudan hex | 2.856 satır, 62 farklı renk, 85 hex, 278 `var()` |

## Maddeler

| # | Ne yapıldı | Nerede |
|---|---|---|
| 1 | Düzenle ve Durum Ata butonları tek makrodan üretiliyor (`data-malzeme` ikisinde de var). Kutuya dokunulmazsa sunucuya `null` gidiyor, sunucu mevcut değeri koruyor. | `_makrolar.html` `duzenle_butonu`, `yonetim.js` |
| 2 | İşlem sonrası bildirim yenilemeden önce saklanıp sonra gösteriliyor. Kaydırma konumu da korunuyor. | `ortak.js` `yenileVeBildir` |
| 3 | Dizgiye Al, Üretilen Adedi Gir, Teslim Et ve Not dialoglarında "İşlemi yapan" alanı var. Ad bu bilgisayarda hesap başına hatırlanıyor ve tüm işlemlerle `isim` olarak gönderiliyor. | `operator.html`, `operator.js` |
| 4 | 820px altında arama alanının `flex` değeri sıfırlandı. | `stil.css` |
| 5 | Monitör: üst menü ve Çıkış yok (kiosk). Yazılar kart boyutuyla ölçekleniyor (container query, desteklenmezse vw). Kart azsa tek sütun geniş kart. "Sayfa x/y" ve ilerleme çubuğu var. Başlık "Makine Dizgi Monitörü". | `monitor.html`, `monitor.js`, `stil.css` |
| 6 | Monitör yenilemeden önce `/api/surum`'u yokluyor. Yanıt yoksa yenilemiyor, bant gösterip artan aralıkla tekrar deniyor. Veri değişince tam turu beklemeden yenileniyor. | `monitor.js`, `app.py` `/api/surum` |
| 7 | Uzun adlar satır içinde kırılıyor. Yükleme geçmişinde iç kayıt öneki gösterilmiyor (tam ad `title`'da). | `stil.css`, `app.py` `yukleme_adi` filtresi |
| 8 | `:focus-visible` için 3px belirgin halka, koyu zeminlerde açık renk. | `stil.css` |
| 9 | Yönetim'de durum ve dizgi tipi iki ayrı filtre grubu. Seçimler, arama, sayfa ve tablo kaydırması yenilemede korunuyor. | `yonetim.html`, `yonetim.js` |
| 10 | Başarı bildirimi 4 sn'de kapanıyor. Hata ve uyarı kapat düğmesiyle kalıcı (`role="alert"`). En fazla 4 bildirim. | `ortak.js`, `stil.css` |
| 11 | Pano, Operatör ve Yönetim 30 sn'de bir `/api/surum` yokluyor. Veri değişince kullanıcı bakmıyorsa (sekme gizli ya da 60 sn işlem yok, açık dialog ya da yazılan alan yok) sayfa kendini yeniliyor, bakıyorsa "Yenile" bandı çıkıyor. Pano'daki "Güncel / Yeni veri var / Bağlantı yok" göstergesi gerçek durumu yansıtıyor. KPI ana sayısı iş emri, farklı stok alt bilgi. KPI kartları tıklanınca o durumu filtreliyor ve dizgi filtresine göre değişiyor (Operatör'de de). | `ortak.js`, `panel.js`, `operator.js`, `_makrolar.html`, `app.py` |
| 12 | Dönem metrikleri sunucuda hesaplanıyor. Varsayılan dönemde API çağrısı yok. Diğer dönemlerde tablo boşaltılmadan güncelleniyor. | `app.py` `panel()`, `panel.html`, `panel.js` |
| 13 | Operatör ekranında son 30 günde teslim edilenler çiziliyor. Eskilerin sayısı ve Pano bağlantısı gösteriliyor. | `app.py` `operator()` |
| 14 | "Teslim Edildi" butonu "Teslim Et", "Adet Bitir" butonu "Üretilen Adedi Gir" oldu. Dialogdaki etiket "Bu seferde tamamlanan adet", altında o ana kadarki ilerleme yazıyor. Teslim, `confirm()` yerine not ve isim alanlı kendi dialogunda yapılıyor ve buton istek sırasında kilitleniyor. Yönetim'de PLANA ALINDI, HAZIR ve TESLİM EDİLDİ seçilince tamamlanan alanı kilitleniyor, DİZGİDE'ye dönülünce eski değer geri geliyor. `max = toplam`. Tüm `confirm()`'ler ortak onay dialoguna taşındı. | `operator.*`, `yonetim.*`, `base.html` |
| 15 | Giriş HTML'i düzeltildi, formlar iç içe değil. `autocomplete="off"` ve satır içi stil kaldırıldı. Flash mesajları artık giriş sayfasında görünüyor (önceden "Hesabınız pasif" gibi mesajlar kayboluyordu). | `giris.html` |
| 16 | Medya sorguları genişten dara sıralandı. 1100px'teki gereksiz `.monitor-grid` kuralı kaldırıldı. | `stil.css` |
| 17 | `monitor.html`'deki satır içi `<style>` bloğu ve 24 `!important` kaldırıldı. | `monitor.html` |
| 18 | Ölü CSS (`pano-grid`, `plan-*`, `panel-hazir-kutu`, `marka-logo`, `giris-dipnot`, özet grafiği vb.) silindi. İki `.bos-durum` tanımı birleşti. Kaydırma çubuğu stilleri tekilleşti. | `stil.css` |
| 19 | `sayi` Jinja filtresi ve `sayiBicimle()` ile Türkçe ondalık: "3,8 gün", "1.234 adet". | `app.py`, `ortak.js` |
| 20 | Taban yazı boyutu 11–13px oldu (açıklama 13, rozet 11, tablo başlığı 11). Nötr rozetin kontrastı 4,45 → 5,49:1. Tüm metin çiftleri AA'yı geçiyor. | `stil.css` |
| 21 | Renk, yazı boyutu ve gölge değerleri `:root` token'larında toplandı (`--dizgide`, `--uyari-cizgi`, `--odak`, `--yazi-sm` vb.). | `stil.css` |
| 22 | Tüm JS `static/js/` altında. Satır içi `onsubmit`, `style=` ve `<script>` kalmadı. **CSP başlığı eklendi** (`script-src 'self'; style-src 'self'` …; önceki #20). Test, 6 sayfada satır içi kod olmadığını denetliyor. | `app.py`, `test_frontend.py` |
| 23 | `url_for('static', …)` adresine dosyanın değişiklik zamanı `?v=` olarak kendiliğinden ekleniyor. Elle sürüm artırma kalktı. | `app.py` `_statik_surum` |
| 24 | `sessionStorage`/`localStorage` erişimi her yerde `try/catch`'li sarmalayıcıdan (`oturumDeposu`, `kaliciDepo`). | `ortak.js` |
| 25 | Hover'da yükselme yalnız tıklanabilir öğelerde. `prefers-reduced-motion` desteği var. `:disabled` imleci `not-allowed`. | `stil.css` |
| 26 | İngilizce terimler Türkçeleşti ("ana kaynağı", "işlem kaydı", "operasyon bilgisi"). Arama her bölümde aynı alanlarda yapılıyor (teslim API'sine `pcb` eklendi). Tarih alanları Pano ve Yönetim'de tek tip gg.aa.yyyy; ortadan düzenlerken imleç yerinde kalıyor. | `_makrolar.html`, `ortak.js`, `yonetim.html` |

## Davranış değişiklikleri (kullanıcıya duyurulmalı)

- **Pano/Operatör KPI:** Büyük sayı artık **iş emri** sayısı. "Farklı stok" altında yazıyor.
- **Operatör ekranı:** Yalnız son 30 günde teslim edilenler görünüyor, eskiler Pano'da (`OPERATOR_TESLIM_GUN`).
- **Operatör dialogları:** "İşlemi yapan" alanı zorunlu. İlk seferde hesap adıyla dolu gelir, değiştirilen ad o bilgisayarda hatırlanır.
- **Monitör:** Üst menü yok. Başka ekrana geçmek için başlıktaki "Pano'ya geç" bağlantısı var. Yalnız makine dizgi kartlarını gösterdiği başlıkta yazıyor.
- **Açık ekranlar kendini yeniliyor:** Veri değişince kimse bakmıyorsa sayfa yenilenir. Bakan kişi ise "Yenile" bandı görür.
- **Yönetim tarih alanları:** Takvim açılır kutusu yerine gg.aa.yyyy yazılan alan (Pano ile aynı; tarayıcı dilinden bağımsız).
- **Pano durum filtresi:** "Hepsi" artık başta (dizgi filtresindeki "Tümü" ile aynı sırada).

## Bilerek yapılmayan / kalanlar

- **Kartı yerinde güncelleme (#2'deki "doğru çözüm"):** Uygulanmadı. Operatör kartının durumu, rengi, rozeti ve butonları
  sunucudaki kurallarla hesaplanıyor; bunları JS'de ikinci kez yazmak iki kaynak demek. Bunun yerine yenileme korundu ve
  kayıp olan her şey (bildirim, kaydırma, filtreler, operatör adı) yenilemeden sağ çıkıyor.
- **Yüklenen dosya adındaki Türkçe harfler:** `secure_filename` "Üretim"i "Uretim" yapıyor. Düzeltmek import rotasına
  dokunmayı gerektirdiği için eşzamanlı çalışan import oturumunun alanında bırakıldı. Yalnız iç önek gizlendi.
- **Önceki rapordan kalanlar:** Giriş formu CSRF'i (önceki #15) ve `depo.py`'deki "Kartı HAZIR durumuna alın" metni
  (önceki #13) bu işin kapsamında değildi.
- **Tarayıcı desteği:** Monitör yazı ölçeklemesi Chrome/Edge 105+, Firefox 110+ ister; daha eskide vw tabanlı yedek boyut kullanılır.
- **CSP kuralı:** Bundan sonra şablonlara satır içi `<script>`, `style="…"` ya da `onclick` eklenirse tarayıcı engeller.
  Yeni JS `static/js/` altına, stil `stil.css`'e yazılmalı. `tests/test_frontend.py` bunu yakalar.
- **Dağıtım:** `app.py` değiştiği için sunucunun yeniden başlatılması gerekir. Statik dosyalar otomatik sürümlü olduğundan
  tarayıcı önbelleği temizlemeye gerek yok.

## Ek istekler (27.09.2026 akşam)

| # | İstek | Yapılan |
|---|---|---|
| 1 | Logo | Üst çubukta (48px, telefonda 40px) ve giriş ekranında (120px) "PDGM" yazısı yerine `static/logo.png`. Not: dosya aslında JPEG (1079×1046), uzantısı `.png`; tarayıcılar içeriğe göre açtığı için sorun değil. |
| 2 | Geciken rozeti | Pano başlığındaki "N geciken açık kart" kaldırıldı. |
| 3 | KPI başlıkları | "Dizgideki İş Emri", "Plana Alınan İş Emri", "Teslim Edilen İş Emri" (16px başlık), altında "X farklı stok". Pano ve Operatör'de aynı. |
| 4 | "20 / 35 kart gösteriliyor" | Hesap doğruydu: ilk sayı ekranda görünen (her bölüm sayfa başına 12 kart), ikincisi filtreye uyan toplam. Metin "35 kart bulundu · ekranda 20, sonraki sayfalarda 15" oldu. Üretim verisinde Tümü/Hepsi: 36 kart = 7 dizgide + 3 plana + 26 teslim; ekranda 7 + 3 + 12 = 22. |
| 5 | "PLANINDA (3 gün var)" | "SÜRESİ İÇİNDE (teslime 3 gün kaldı)" oldu ("SÜRE AŞILDI (N gün)" ile karşıt). Monitörde kısa hâli "PLANDA" yerine "SÜRESİ İÇİNDE". Excel raporuna da yeni metin gider. Eski metni bekleyen testler, `docs/EXCEL_SYNC.md` ve test rehberi güncellendi. |

Test: 112/112 (yeni `test_logo_kpi_basliklari_ve_rozet_metni`).
