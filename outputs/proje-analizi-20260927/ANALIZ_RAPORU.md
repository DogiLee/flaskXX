# PDGM İş Takip — Proje Analizi (27.09.2026)

Kapsam: `app.py`, `depo.py`, `excel_araclari.py`, `kullanici_yonet.py`, 9 şablon + satır içi JS,
`static/stil.css`, testler, `.bat` betikleri, bağımlılıklar. Kod değiştirilmedi; ölçümler ve denemeler
geçici klasörlerde yapıldı. Üretim `data/` klasörü yalnız salt okunur incelendi.

Önem sırası: **Kritik** (kullanımı bozar/veri kaybı), **Yüksek**, **Orta**, **Düşük**.

## Özet

| # | Önem | Bulgu | Kanıt |
|---|---|---|---|
| 1 | Kritik | Log büyüdükçe her işlem saniyeler–dakika sürüyor, bu sırada tüm uygulama kilitli | Ölçüldü: 5.000 log → 4,6 sn, 20.000 log → 55 sn / tıklama |
| 2 | Kritik | Excel güncellenmeden tekrar yüklenirse operatörün uygulamadaki işi sessizce geri alınıyor | Deneme: 6/10 → PLANA 0/10; bugün teslim edilen → DİZGİDE |
| 3 | Yüksek | waitress 3.0.0 ve Werkzeug 3.0.6'da bilinen güvenlik açıkları | CVE-2024-49769 (varsayılan ayarda DoS), CVE-2025-66221 |
| 4 | Yüksek | Proje sürüm kontrolünde değil (git yok) | Değişiklikler elle kopyalanan `.investigation` klasörleriyle izleniyor |
| 5 | Yüksek | Gece yedeği betiği başka bir kullanıcının masaüstüne yazıyor | `HEDEF=C:\Users\ulasarpkocak\Desktop\...` |
| 6 | Orta | Parola ve oturum çerezi ağda şifresiz (HTTP), gözlemci şifresiz | `0.0.0.0:5001`, `SESSION_COOKIE_SECURE` kapalı |
| 7 | Orta | Admin not geçmişini silebiliyor, audit logu not farkını tutmuyor | `admin_kart_duzenle` |
| 8 | Orta | Pano ve Operatör ekranı kendiliğinden yenilenmiyor | Yalnız monitör yenileniyor |
| 9 | Orta | Log arşivleme yalnız girişte tetikleniyor | `_log_arsivle_gerekirse` tek çağrı |
| 10 | Orta | `.env` eksikse ilk açılış çöküyor; ilk parolalar `.env`'de açık metin kalıyor | `generate_password_hash(None)` → AttributeError |
| 11 | Orta | Manuel kartta aynı Talep+Stok ikinci kez eklenemiyor, Excel kartının kopyası eklenebiliyor | `admin_kart_ekle` anahtar kontrolü |
| 12–20 | Düşük | Ölü kod, metin tutarsızlığı, budanmayan yüklemeler, giriş CSRF'i, her yanıtta çerez vb. | Aşağıda |

## Kritik

### 1. İşlem logu büyüdükçe uygulama donuyor
- **Ne oluyor:** Her operatör işlemi (dizgiye al, adet gir, teslim, not) ve her giriş, `kartlar.xlsx` ile
  `islem_logu.xlsx` dosyalarını **baştan** yazıyor. Log satırları, her satır için tüm hücreleri
  tarayan `ws.max_row` ile ekleniyor (`depo.py` `_workbook_uret`), yani süre karesel artıyor.
  Yazma sırasında `_kilit` tutulduğu için pano, operatör ve API istekleri de bekliyor.
- **Ölçüm** (300 kart, geçici klasör):

  | Log satırı | Tek operatör işlemi | Açılış |
  |---|---|---|
  | 1.000 | 0,76 sn | 0,27 sn |
  | 5.000 | 4,58 sn | 0,79 sn |
  | 20.000 (arşiv sınırı) | 55,2 sn | 2,91 sn |

  Profil: 63,7 sn'nin 46 sn'si `ws.max_row` çağrısında. Aynı dosya `ws.append` ile 3,0 sn'de yazıldı.
  Rapor indirme (`excel_araclari._sayfa_yaz`) aynı desende: 10.000 log satırıyla 13,6 sn.
- **Ne zaman sorun olur (tahmin):** Günde ~250 log satırıyla (ör. 5 operatör × 40 işlem, girişler,
  importlar) yaklaşık 20 iş gününde 5.000, 80 iş gününde 20.000 satıra ulaşılır. Şu an üretimde 25 satır var.
- **Öneri:**
  1. Hızlı düzeltme (küçük değişiklik): `_workbook_uret` ve `_sayfa_yaz` içinde satır numarasını
     sayaçla tutun ya da `ws.append` kullanın; hücre bazlı font atamasını kaldırın. `=` ile başlayan
     metinlerin formül olmaması korumasını sürdürün.
  2. Kalıcı çözüm: işlem logunu ekleme yapılan bir biçimde (JSONL/CSV veya SQLite) tutun, Excel'i
     yalnız indirme sırasında üretin. `LOG_SINIRI` değerini düşürün (ör. 5.000).

### 2. Excel importu operatörün uygulamada yaptığı işi geri alıyor

> **Durum (27.09.2026): giderildi.** Önizlemede "Karar gerekiyor" bölümü ve kart bazında durum seçimi eklendi (bkz. `outputs/import-duzeltme-20260926/RAPOR.md` 2e, `tests/test_excel_geride.py`).
- **Deneme:** Operatör 1. kartı dizgiye alıp 6/10 girdi, 2. kartı 4/4 bitirip teslim etti. Excel
  güncellenmeden aynı dosya tekrar yüklendi. Sonuç: 1. kart **PLANA ALINDI 0/10**, 2. kart **DİZGİDE**
  (teslim tarihi silindi). Önizleme bunu yalnız sarı bir "Durum" satırı olarak gösteriyor.
- **Neden:** Kart durumunun iki sahibi var: Excel (import) ve operatör ekranı. Kural "geçerli Excel
  DURUM'u her zaman kazanır". Sahada yaşanan "ters görünme / teslim edilmiş kart geri döndü"
  şikâyetlerinin önemli bir kısmı büyük olasılıkla buradan geliyor.
- **Öneri (iş kararı gerektirir):**
  - Önizlemeye "Uygulamada ilerlemiş, Excel geride kalan kartlar" grubu ekleyin. Excel'in durumu
    uygulamadakinin gerisindeyse (PLANA < DİZGİDE < TESLİM) bu kartlarda varsayılan olarak uygulama
    durumu korunsun; admin kart bazında "Excel'e göre geri al" seçebilsin. Tamamlanan adet için eklenen
    onay kutusu mekanizması doğrudan kullanılabilir.
  - Alternatif: iş akışı ileri yönlü kural (import durumu yalnız ileri taşır, geri almayı yalnız admin yapar).

## Yüksek

### 3. Bağımlılıklardaki bilinen açıklar
- `waitress==3.0.0`: **CVE-2024-49769** (istemci erken kapanınca meşgul döngü / kaynak tükenmesi,
  varsayılan ayarları da etkiler) ve CVE-2024-49768 (yalnız `channel_request_lookahead` açıkken).
  Düzeltme: 3.0.1+.
- `Werkzeug==3.0.6`: **CVE-2025-66221**, `safe_join` Windows aygıt adlarını (CON, AUX, NUL) engellemiyor.
  Düzeltme: 3.1.4+. Bu makinede denendi: `safe_join` bu adları engellemedi ama Flask statik rotası
  404 döndü, asılı kalmadı, yani pratik risk düşük.
- **Öneri:** `waitress` ≥ 3.0.1, `Flask` 3.1.x ile `Werkzeug` ≥ 3.1.4'e birlikte geçin, ardından test
  paketini (`PDGM_TEST_EXCEL_COM=1`) çalıştırın.

### 4. Sürüm kontrolü yok
Proje bir git deposu değil. Değişiklik geçmişi `.investigation/*-before` kopyaları ve diff dosyalarıyla
elle tutuluyor. Geri dönüş ve kimin neyi değiştirdiği güvenilir değil. **Öneri:** `git init` ile ilk commit
atın (`.gitignore` zaten `data/` ve `.env`'i dışlıyor), her değişikliği commit'leyin.

### 5. Gece yedeği bu makinede çalışmaz
`yedek_disari_kopyala.bat`: `HEDEF=C:\Users\ulasarpkocak\Desktop\pdgm_arayuz_yedek`. Bu yol bu
makinede yok, betik "Hedef erişilemiyor" ile çıkar. Ayrıca hedef aynı diskte. Yedek `gizli.key` ve
`kullanicilar.json` dosyalarını da içeriyor. **Öneri:** Hedefi erişimi kısıtlı bir ağ paylaşımına alın,
Görev Zamanlayıcı'da gerçekten çalıştığını ve robocopy logunu kontrol edin.

## Orta

### 6. Ağ güvenliği
Sunucu `0.0.0.0:5001` üzerinde HTTP ile yayında; giriş parolaları ve oturum çerezi ağda açık metin
gidiyor. `SESSION_COOKIE_SECURE` yalnız `PDGM_HTTPS=1` ile açılıyor. Gözlemci girişi şifresiz; ağdaki
herkes pano, notlar ve talep sahibi adlarını görebiliyor (tasarım kararı, onaylanmalı).
**Öneri:** Önüne HTTPS terminasyonu (IIS/Caddy) koyun ve `PDGM_HTTPS=1` yapın; gözlemci erişimini IP ile
kısıtlayın veya parola ekleyin.

### 7. Admin not geçmişini iz bırakmadan değiştirebiliyor
Yönetim "Düzenle" formunda tüm not geçmişi düzenlenebilir bir metin alanında geliyor ve olduğu gibi
kaydediliyor (`admin_kart_duzenle`). Log kaydı yalnız durum/adet yazıyor (`ADMİN DÜZENLEDİ`), not farkını
yazmıyor. Operatör notları silinirse izi kalmaz. **Öneri:** Admin notu da ekleme biçiminde olsun ya da
logda eski/yeni not farkı saklansın.

### 8. Pano ve Operatör ekranı bayat kalabiliyor
Monitör kendini yeniliyor; Pano ve Operatör ekranı yalnız elle "Yenile" ile ya da bir işlemden sonra
yenileniyor. Import sonrası açık kalan ekranlar eski durumu gösterir. Sunucu yanlış işlemi reddettiği
için veri bozulmaz, ama kullanıcı kafa karıştırıcı hata alır. **Öneri:** `depo_surumu()` döndüren
küçük bir uç nokta ve 30–60 sn'de bir yoklayıp "Veriler güncellendi, yenileyin" bandı gösteren kod.

### 9. Log arşivleme yalnız girişte tetikleniyor
`_log_arsivle_gerekirse()` yalnız `log_ekle` içinde çağrılıyor. Operatör işlemleri ve importlar log
eklediği hâlde arşivlemeyi tetiklemiyor; log 20.000'i aşabiliyor (bkz. #1). **Öneri:** Arşiv kontrolünü
ortak commit yoluna (`_kart_log_commit`, import) taşıyın.

### 10. `.env` ve ilk kullanıcılar
`.env`'de bir parola eksikse ilk açılışta `generate_password_hash(None)` AttributeError ile çöküyor
(anlaşılır mesaj yok, uzunluk kontrolü yok). İlk kurulumdan sonra açık metin parolalar `.env`'de kalıyor.
`test_verileri/TEST_SENARYOLARI.txt` ve kod dışa aktarımında `eum1 / eum123` açık yazılı.
**Öneri:** Eksik/kısa parolada anlaşılır hata; kurulumdan sonra `.env`'deki parolaları silin; bu parola
hâlâ geçerliyse değiştirin (`python kullanici_yonet.py`).

### 11. Manuel kart kimliği
`admin_kart_ekle` anahtarı `Talep|Stok`. Aynı Talep+Stok ikinci bir manuel kart olarak eklenemiyor
(1.2'deki "aynı talep/stoktan birden fazla olabilir" kuralıyla çelişir). Buna karşılık Excel'den gelmiş
bir kartın aynısı manuel olarak eklenebiliyor, ekranda çift görünür. **Öneri:** Manuel kartlara benzersiz
kimlik (ör. `MANUEL:<id>`) verin, aynı Talep+Stok'lu aktif bir Excel kartı varsa uyarın.

## Düşük / iyileştirme

12. **Ölü kod:** `templates/ozet.html` hiçbir rotada kullanılmıyor. `depo.kart_guncelle`, `toplu_kaydet`,
    `kart_bul`, `yeni_kimlik`, `yeniden_yukle` hiçbir yerden çağrılmıyor. `depo.kur()` ve
    `process_kilidi_al()` açılışta iki kez çalışıyor (modül yüklemesi + `calistir`).
13. **Metin tutarsızlığı:** Adet bitmiş karta yeniden adet girilince "Kartı HAZIR durumuna alın" deniyor
    (`depo.py:1580`), oysa akış "Teslim Et".
14. **Budanmayan yüklemeler:** Önizleme için yüklenen dosyalar `data/yuklenen_exceller`'e kaydediliyor;
    budama yalnız onaydan sonra çalışıyor. Onaylanmayan önizlemeler birikebilir.
15. **Giriş rotalarında CSRF yok:** Başka bir siteden `/giris/gozlemci`'ye POST ile açık admin oturumu
    gözlemci oturumuna çevrilebilir (zorla çıkış). Düşük risk; giriş formuna da token eklenebilir.
16. **Her yanıtta çerez:** `before_request` her istekte oturuma yazdığı için her yanıt `Set-Cookie`
    taşıyor (denendi). Değer değişmedikçe yazmayın.
17. **Büyük fonksiyonlar:** `excel_import_uygula` yaklaşık 280 satır; `app.py` 1.570, `depo.py`
    2.600 satır. Birleştirme kuralları (ayrılma, sıfırlama, pasifleştirme, önizleme) ayrı fonksiyonlara
    bölünürse bakım kolaylaşır.
18. **Proje kökünde dağınıklık:** `pdgm_codebase_export.*` (~650 KB), `old_codebase.txt`,
    `artifact_work/`, `outputs/`, `.investigation/` çalışma klasöründe duruyor. Dağıtım klasöründen ayırın.
19. **JS testleri çalışmıyor:** `tests/*.cjs` Node gerektiriyor, bu makinede Node yok; CI da yok.
20. **CSP başlığı yok:** Şablonlarda `innerHTML`/`|safe` kullanılmıyor (XSS açısından iyi), ama bir
    `Content-Security-Policy` başlığı ek güvence olur.

## İyi durumda olanlar

- Import önce tüm sayfaları okuyup doğruluyor, sonra tek seferde yazıyor. Bellek ve disk geri alma
  testlerle doğrulanmış, her import öncesi yedek alınıyor.
- Durum değiştiren tüm rotalarda yetki ve CSRF kontrolü var. Parolalar scrypt hash'iyle saklanıyor,
  giriş deneme sınırı var, girişte oturum yenileniyor.
- COM tarafında makrolar ve dış bağlantılar kapalı; testler sonrası yetim `EXCEL.EXE` süreci kalmadı.
- Şablonlarda XSS'e açık yazım yok. İstemci önbelleği `no-store` ile kapalı.
- Test paketi 89 test (gerçek Excel COM dahil) ve hepsi geçiyor.

## Önerilen sıra

1. #1 hızlı düzeltme (küçük değişiklik, büyük etki) ve #9.
2. #2 için iş kararı, ardından önizleme grubu.
3. #3 bağımlılık güncellemesi + tam test koşusu; #4 git.
4. #5 yedek hedefi; #6 HTTPS.
5. Kalanlar fırsat buldukça.
