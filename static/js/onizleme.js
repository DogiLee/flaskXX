/* Excel aktarım önizlemesi: çift gönderim koruması, "tamamlanan adedi sıfırla" seçimleri
   ve Excel'in geride kaldığı kartlar için durum kararları.
   (import_onizleme.html içindeki satır içi betikten CSP için taşındı; davranış aynı.) */
"use strict";

(() => {
    // Excel COM okuması birkaç saniye sürer; ikinci tıklama kullanılmış önizlemeye çarpmasın.
    document.querySelectorAll("form[data-tek-gonderim]").forEach((form) => {
        form.addEventListener("submit", () => {
            const buton = form.querySelector("button[type=submit]");
            buton.disabled = true;
            buton.textContent = "Uygulanıyor…";
        });
    });

    // Tamamlanan adet sıfırlama ve not temizleme seçimleri: onay çubuğundaki sayaçlar
    // ve grup bazında toplu seçim.
    const sifirlaKutulari = [...document.querySelectorAll("input[data-sifirla]")];
    const notKutulari = [...document.querySelectorAll("input[data-not-temizle]")];
    const sifirlaOzeti = document.querySelector("[data-sifirla-ozet]");
    const notOzeti = document.querySelector("[data-not-ozet]");
    function sayaciGuncelle(ozet, kutular) {
        if (!ozet) return;
        const secili = kutular.filter((kutu) => kutu.checked).length;
        ozet.hidden = secili === 0;
        ozet.querySelector("i").textContent = secili;
    }
    function sifirlaOzetiniGuncelle() {
        sayaciGuncelle(sifirlaOzeti, sifirlaKutulari);
        sayaciGuncelle(notOzeti, notKutulari);
    }
    [...sifirlaKutulari, ...notKutulari].forEach((kutu) => kutu.addEventListener("change", sifirlaOzetiniGuncelle));
    document.querySelectorAll("[data-hepsini-sec]").forEach((buton) => {
        const secici = buton.dataset.hepsiniSec === "not" ? "input[data-not-temizle]" : "input[data-sifirla]";
        buton.addEventListener("click", () => {
            const kutular = [...buton.closest(".etki-grup").querySelectorAll(secici)];
            const hepsiSecili = kutular.every((kutu) => kutu.checked);
            kutular.forEach((kutu) => { kutu.checked = !hepsiSecili; });
            buton.textContent = hepsiSecili ? "Tümünü işaretle" : "İşaretleri kaldır";
            sifirlaOzetiniGuncelle();
        });
    });
    sifirlaOzetiniGuncelle();   // geri tuşuyla dönülünce tarayıcının koruduğu seçimler

    // Excel'in geride kaldığı kartlar: seçilen duruma göre sonuç metni ve sıfırlama kutusu.
    const DURUM_SONUCU = {
        "": "Kart durumsuz kalır: Pano, Operatör ve Monitör'den kalkar, Yönetim'deki Durumu Eksik Kartlar listesine düşer; tamamlanan adet 0 olur, başlama, bitiş ve teslim bilgileri silinir.",
        "HAZIR": "Kart HAZIR'a döner: tamamlanan adet 0 olur, başlama, bitiş ve teslim bilgileri silinir.",
        "PLANA ALINDI": "Kart plana döner: tamamlanan adet 0 olur, başlama, bitiş ve teslim bilgileri silinir.",
        "DİZGİDE": "Kart DİZGİDE olur: tamamlanan adet korunur (isterseniz sıfırlayın), teslim bilgisi silinir.",
        "TESLİM EDİLDİ": "Kart teslim edilmiş kalır: adet toplama eşit, uygulamadaki teslim tarihi korunur.",
    };
    const gerilemeOzeti = document.querySelector("[data-gerileme-ozet]");
    const gerilemeSecimleri = [...document.querySelectorAll("[data-gerileme-secim]")];
    function gerilemeKartiniGuncelle(secim) {
        const kart = secim.closest("[data-gerileme-kart]");
        const kaynak = secim.value === secim.dataset.excel ? "Excel'deki durum uygulanır. " : "Uygulamadaki ilerleme korunur. ";
        kart.querySelector("[data-gerileme-sonuc]").textContent = kaynak + (DURUM_SONUCU[secim.value] ?? "");
        kart.classList.toggle("uygulama-korunuyor", secim.value !== secim.dataset.excel);
        const kutu = kart.querySelector("[data-gerileme-sifirla]");
        if (kutu) {
            const acik = secim.value === "DİZGİDE";
            kutu.disabled = !acik;
            if (!acik) kutu.checked = false;
            kutu.closest(".sifirla-secenek").classList.toggle("pasif", !acik);
        }
        if (gerilemeOzeti) {
            gerilemeOzeti.querySelector("i").textContent =
                gerilemeSecimleri.filter((s) => s.value !== s.dataset.excel).length;
        }
        sifirlaOzetiniGuncelle();
    }
    gerilemeSecimleri.forEach((secim) => {
        secim.addEventListener("change", () => gerilemeKartiniGuncelle(secim));
        gerilemeKartiniGuncelle(secim);
    });
    document.querySelectorAll("[data-gerileme-hepsi]").forEach((buton) => {
        buton.addEventListener("click", () => {
            gerilemeSecimleri.forEach((secim) => {
                secim.value = buton.dataset.gerilemeHepsi === "excel" ? secim.dataset.excel : secim.dataset.uygulama;
                gerilemeKartiniGuncelle(secim);
            });
        });
    });
})();
