/* Atölye monitörü: kart rotasyonu, sayfa göstergesi ve bağlantıya dayanıklı yenileme.
   Sayfa bir TV'de gözetimsiz çalışır: sunucuya ulaşılamazsa tarayıcı hata sayfasına
   düşmemek için yenilemeden önce sunucu yoklanır; ulaşılamıyorsa eski veri bir uyarı
   bandıyla ekranda kalır ve artan aralıklarla yeniden denenir. */
"use strict";

(() => {
    const SLAYT_MS = 12000;
    const EN_AZ_YENILEME_MS = 60000;
    const YOKLAMA_MS = 30000;
    const EN_UZUN_BEKLEME_MS = 60000;

    const saat = document.getElementById("monitor-saat");
    const baglanti = document.getElementById("monitor-baglanti");
    const veriZamani = document.getElementById("monitor-veri-zamani");
    const surum = document.body?.dataset?.veriSurumu || "";
    const depo = typeof oturumDeposu !== "undefined" ? oturumDeposu : {al: () => null, yaz() {}};

    function saatiGuncelle() {
        if (!saat) return;
        saat.textContent = new Intl.DateTimeFormat("tr-TR", {
            hour: "2-digit", minute: "2-digit", second: "2-digit",
        }).format(new Date());
    }

    function sayfaBoyutuHesapla() {
        // Laptop / kısa ekran: 4 kart (2x2). Geniş ve yüksek ekran: 6 kart (2x3).
        const kisa = window.matchMedia("(max-height: 900px)").matches;
        const dar = window.matchMedia("(max-width: 1100px)").matches;
        return (kisa || dar) ? 4 : 6;
    }

    const tekSutun = () => window.matchMedia("(max-width: 640px)").matches;

    function rotasyonBaslat(grup) {
        const kartlar = [...grup.querySelectorAll("[data-monitor-kart]")];
        const bilgi = grup.closest?.(".monitor-bolum")?.querySelector("[data-sayfa-bilgi]");
        const konumAnahtari = `pdgm-monitor-${grup.dataset.monitorGrup}`;
        const kayit = Number(depo.al(konumAnahtari));
        let aktifSayfa = Number.isInteger(kayit) && kayit >= 0 ? kayit : 0;
        let sayfaSayisi = 1;

        function uygula() {
            const boyut = sayfaBoyutuHesapla();
            sayfaSayisi = Math.max(1, Math.ceil(kartlar.length / boyut));
            if (aktifSayfa >= sayfaSayisi) aktifSayfa = 0;
            const baslangic = aktifSayfa * boyut;
            kartlar.forEach((kart, index) => {
                kart.classList.toggle("monitor-gizli", !(index >= baslangic && index < baslangic + boyut));
            });

            // Izgara toplam kart sayısına göre kurulur (tüm sayfalarda aynı kalsın):
            // kart bir sütunu dolduracak kadar azsa tek sütun geniş kartlar, değilse 2 sütun.
            // Yazılar kart boyutuyla ölçeklendiği için büyük kart = büyük yazı.
            if (grup.style) {
                const enCokSatir = boyut / 2;
                let sutun = 2;
                let satir = enCokSatir;
                if (kartlar.length <= enCokSatir) {
                    sutun = 1;
                    satir = Math.max(kartlar.length, 2);
                }
                const tek = tekSutun();
                grup.style.gridTemplateColumns = tek ? "" : `repeat(${sutun}, minmax(0, 1fr))`;
                grup.style.gridTemplateRows = tek ? "" : `repeat(${satir}, minmax(0, 1fr))`;
            }

            if (bilgi) {
                bilgi.hidden = sayfaSayisi <= 1;
                bilgi.querySelector("[data-sayfa-metin]").textContent = `Sayfa ${aktifSayfa + 1}/${sayfaSayisi}`;
                const cubuk = bilgi.querySelector("[data-sayfa-ilerleme]");
                cubuk.style.animationDuration = `${SLAYT_MS}ms`;
                cubuk.classList.remove("oynat");
                void cubuk.offsetWidth;  // animasyonu baştan başlat
                cubuk.classList.add("oynat");
            }
            depo.yaz(konumAnahtari, String(aktifSayfa));
        }

        uygula();
        window.setInterval(() => {
            aktifSayfa = (aktifSayfa + 1) % sayfaSayisi;
            uygula();
        }, SLAYT_MS);

        let boyutZamanlayici = null;
        window.addEventListener("resize", () => {
            window.clearTimeout(boyutZamanlayici);
            boyutZamanlayici = window.setTimeout(() => { aktifSayfa = 0; uygula(); }, 200);
        });
    }

    // ------------------------------------------------------------------
    // Sunucu yoklaması ve güvenli yenileme
    // ------------------------------------------------------------------
    let hataSayisi = 0;
    let degisti = false;
    let yenileniyor = false;

    function baglantiyiGoster(ulasildi) {
        if (!baglanti) return;
        if (ulasildi) {
            baglanti.hidden = true;
            return;
        }
        const zaman = veriZamani ? veriZamani.textContent : "";
        baglanti.hidden = false;
        baglanti.textContent = `Sunucuya ulaşılamıyor. Ekrandaki bilgiler ${zaman} itibarıyla; yeniden deneniyor…`;
    }

    async function sunucuyuYokla() {
        const iptal = typeof AbortController === "function" ? new AbortController() : null;
        const zamanAsimi = window.setTimeout(() => iptal?.abort(), 8000);
        try {
            const yanit = await fetch("/api/surum", {credentials: "same-origin", cache: "no-store", signal: iptal?.signal});
            if (yanit.status === 401 || yanit.redirected) return "oturum";
            if (!yanit.ok) throw new Error(String(yanit.status));
            const veri = await yanit.json();
            if (surum && veri.surum && veri.surum !== surum) degisti = true;
            hataSayisi = 0;
            baglantiyiGoster(true);
            return "tamam";
        } catch (_) {
            hataSayisi += 1;
            baglantiyiGoster(false);
            return "hata";
        } finally {
            window.clearTimeout(zamanAsimi);
        }
    }

    async function yenilemeyiDene() {
        if (yenileniyor) return;
        yenileniyor = true;
        const sonuc = await sunucuyuYokla();
        if (sonuc === "hata") {
            // Hata sayfasına düşmemek için yenileme yapılmaz; eski veri bantla ekranda kalır.
            yenileniyor = false;
            window.setTimeout(yenilemeyiDene, Math.min(EN_UZUN_BEKLEME_MS, 15000 * hataSayisi));
            return;
        }
        window.location.reload();
    }

    saatiGuncelle();
    window.setInterval(saatiGuncelle, 1000);
    const gruplar = [...document.querySelectorAll("[data-monitor-grup]")];
    gruplar.forEach(rotasyonBaslat);

    // Ara yoklama: veri değiştiyse hemen, oturum düştüyse giriş için yenile; bağlantı yoksa bant göster.
    window.setInterval(async () => {
        if (yenileniyor) return;
        const sonuc = await sunucuyuYokla();
        if (sonuc === "oturum" || (sonuc === "tamam" && degisti)) yenilemeyiDene();
    }, YOKLAMA_MS);

    // Düzenli yenileme en uzun grubun bütün sayfalarına en az bir gösterim süresi tanır.
    // Sayfa konumu saklandığı için yenilemeden sonra rotasyon kaldığı yerden devam eder.
    const enCokSayfa = Math.max(1, ...gruplar.map(
        (grup) => Math.ceil(grup.querySelectorAll("[data-monitor-kart]").length / sayfaBoyutuHesapla())));
    window.setTimeout(yenilemeyiDene, Math.max(EN_AZ_YENILEME_MS, enCokSayfa * SLAYT_MS));
})();
