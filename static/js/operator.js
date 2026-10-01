/* Operatör ekranı: filtreler, kart işlemleri ve "işlemi yapan" adının hatırlanması. */
"use strict";

(() => {
    const kartAra = document.getElementById("kart-ara");
    const kartlar = [...document.querySelectorAll("[data-kart]")];
    const kpiKartlari = [...document.querySelectorAll("[data-kpi-filtre]")];
    const sonuc = document.getElementById("operator-sonuc");
    const temizle = document.getElementById("operator-temizle");
    const bos = document.getElementById("operator-bos");
    const ARAMA_KEY = "pdgm-op-arama";

    const durumUygunMu = (durum) => {
        const filtre = durumFiltresi.deger;
        if (filtre === "HEPSI") return true;
        if (filtre === "AKTIF") return durum === "PLANA ALINDI" || durum === "DİZGİDE";
        return durum === filtre;
    };
    const dizgiUygunMu = (tip) => dizgiFiltresi.deger === "HEPSI" || tip === dizgiFiltresi.deger;

    function filtrele() {
        const arama = aramaMetni(kartAra.value);
        let gorunen = 0;
        for (const kart of kartlar) {
            const uygun = durumUygunMu(kart.dataset.durum)
                && dizgiUygunMu(kart.dataset.dizgiTipi)
                && (!arama || aramaMetni(kart.dataset.arama).includes(arama));
            kart.hidden = !uygun;
            if (uygun) gorunen += 1;
        }
        sonuc.textContent = `${gorunen} kart gösteriliyor`;
        bos.hidden = gorunen !== 0 || kartlar.length === 0;
        temizle.hidden = !arama && durumFiltresi.deger === "AKTIF" && dizgiFiltresi.deger === "HEPSI";
    }

    const DIZGI_ANAHTARI = {HEPSI: "Hepsi", MAKINE: "Makine", ELLE: "Elle", EUM: "Eum"};

    function kpiGuncelle() {
        const ek = DIZGI_ANAHTARI[dizgiFiltresi.deger] || "Hepsi";
        kpiKartlari.forEach((kutu) => {
            kutu.querySelector("[data-kpi-kart]").textContent = kutu.dataset[`kart${ek}`] ?? "0";
            kutu.querySelector("[data-kpi-stok]").textContent = kutu.dataset[`stok${ek}`] ?? "0";
            kutu.setAttribute("aria-pressed", kutu.dataset.kpiFiltre === durumFiltresi.deger ? "true" : "false");
        });
    }

    const durumFiltresi = filtreGrubuKur({
        butonlar: [...document.querySelectorAll("[data-filtre]")],
        veriAdi: "filtre",
        anahtar: "pdgm-op-filtre",
        varsayilan: "AKTIF",
        degisince: () => { kpiGuncelle(); filtrele(); },
    });
    const dizgiFiltresi = filtreGrubuKur({
        butonlar: [...document.querySelectorAll("[data-dizgi-filtre]")],
        veriAdi: "dizgiFiltre",
        anahtar: "pdgm-op-dizgi-filtre",
        varsayilan: "HEPSI",
        degisince: () => { kpiGuncelle(); filtrele(); },
    });

    kpiKartlari.forEach((kutu) => {
        kutu.addEventListener("click", () => {
            const hedef = kutu.dataset.kpiFiltre;
            durumFiltresi.sec(durumFiltresi.deger === hedef ? "AKTIF" : hedef);
        });
    });

    kartAra.value = oturumDeposu.al(ARAMA_KEY, "");
    kartAra.addEventListener("input", () => {
        oturumDeposu.yaz(ARAMA_KEY, kartAra.value || "");
        filtrele();
    });
    temizle.addEventListener("click", () => {
        kartAra.value = "";
        oturumDeposu.sil(ARAMA_KEY);
        durumFiltresi.sec("AKTIF", {sessiz: true});
        dizgiFiltresi.sec("HEPSI");
        kartAra.focus();
    });

    kpiGuncelle();
    filtrele();

    // ------------------------------------------------------------------
    // İşlemi yapan: operatör hesapları paylaşımlı olduğu için kişinin adı bu
    // bilgisayarda (hesap başına) hatırlanır ve tüm işlemlerle gönderilir; karta ve
    // işlem loguna yazılır. Hesabın adı (ör. "Makine Operatörü") kişi adı sayılmaz:
    // alan onunla doldurulmaz ve onunla gönderilemez.
    // ------------------------------------------------------------------
    const ISIM_KEY = `pdgm-islem-yapan:${document.body.dataset.kullanici || ""}`;
    const hesapAdlari = [document.body.dataset.ad, document.body.dataset.kullanici]
        .filter(Boolean).map(aramaMetni);
    const hesapAdiMi = (isim) => hesapAdlari.includes(aramaMetni(isim));
    const kayitliIsim = () => {
        const isim = kaliciDepo.al(ISIM_KEY, "");
        return isim && !hesapAdiMi(isim) ? isim : "";
    };

    function dialogAc(dialog) {
        const isim = dialog.querySelector("[data-islem-yapan]");
        if (isim) isim.value = kayitliIsim();
        dialog.showModal();
        // Ad boşsa önce ada, değilse ilk alana odaklan.
        const ilk = isim && !isim.value.trim() ? isim : dialog.querySelector("input:not([type=hidden]), textarea");
        ilk?.focus();
    }

    function isimAl(form) {
        const alan = form.querySelector("[data-islem-yapan]");
        const isim = alan.value.trim();
        if (isim && !hesapAdiMi(isim)) kaliciDepo.yaz(ISIM_KEY, isim);
        return isim;
    }

    async function gonder(event, url, govde, basariMesaji) {
        event.preventDefault();
        const form = event.target;
        const submit = event.submitter || form.querySelector("[type=submit]");
        const isim = isimAl(form);
        if (!isim || hesapAdiMi(isim)) {
            toast(isim ? "Hesabın adı değil, işlemi yapan kişinin kendi adını yazın."
                       : "İşlemi yapan kişinin adını yazın.", "uyari");
            form.querySelector("[data-islem-yapan]").focus();
            return;
        }
        submit.disabled = true;
        try {
            const data = await pdgmFetch(url, {method: "POST", body: JSON.stringify({...govde(), isim})});
            dialogKapat(form.closest("dialog"));
            yenileVeBildir(data.mesaj || basariMesaji);
        } catch (hata) {
            hataMesaji(hata);
            submit.disabled = false;
        }
    }

    const el = (id) => document.getElementById(id);

    // --- Dizgiye Al (malzeme bekleyen kartta önce onay) ---
    let malzemeOnayVerildi = false;
    let malzemeOnayBekleyenButon = null;

    function baslatDialogunuAc(buton) {
        const kalan = Number(buton.dataset.kalan || 1);
        el("baslat-id").value = buton.dataset.id;
        el("baslat-adet").value = kalan;
        el("baslat-adet").max = kalan;
        el("baslat-not").value = "";
        el("baslat-kalan").textContent = `En fazla ${kalan} adet`;
        el("baslat-baslik").textContent = buton.dataset.baslik;
        dialogAc(el("baslat-dialog"));
    }

    document.querySelectorAll("[data-baslat]").forEach((buton) => {
        buton.addEventListener("click", () => {
            if (buton.dataset.malzemeBekliyor === "1") {
                malzemeOnayBekleyenButon = buton;
                el("malzeme-baslik").textContent = buton.dataset.baslik;
                el("malzeme-dialog").showModal();
                return;
            }
            malzemeOnayVerildi = false;
            baslatDialogunuAc(buton);
        });
    });

    el("malzeme-onayla").addEventListener("click", () => {
        dialogKapat(el("malzeme-dialog"));
        if (!malzemeOnayBekleyenButon) return;
        malzemeOnayVerildi = true;
        baslatDialogunuAc(malzemeOnayBekleyenButon);
    });

    el("baslat-dialog").addEventListener("close", () => {
        // Dialog hangi sebeple kapanırsa kapansın onay bir sonraki karta sızmasın.
        malzemeOnayVerildi = false;
        malzemeOnayBekleyenButon = null;
    });

    el("baslat-form").addEventListener("submit", (event) => gonder(event, "/api/basla", () => ({
        kart_id: Number(el("baslat-id").value),
        adet: Number(el("baslat-adet").value),
        not: el("baslat-not").value.trim(),
        malzeme_onayi: malzemeOnayVerildi,
    }), "Kart DİZGİDE durumuna alındı."));

    // --- Üretilen adedi gir ---
    document.querySelectorAll("[data-bitir]").forEach((buton) => {
        buton.addEventListener("click", () => {
            const kalan = Number(buton.dataset.kalan || 1);
            el("bitir-id").value = buton.dataset.id;
            // Bilerek boş: dolu gelirse hızlı bir Enter kalan adedin tamamını bitmiş kaydeder.
            el("bitir-adet").value = "";
            el("bitir-adet").placeholder = kalan > 1 ? `1–${kalan}` : "1";
            el("bitir-adet").max = kalan;
            el("bitir-not").value = "";
            el("bitir-kalan").textContent =
                `Şu ana kadar ${buton.dataset.tamamlanan}/${buton.dataset.toplam} tamamlandı · en fazla ${kalan} adet girilebilir`;
            el("bitir-baslik").textContent = buton.dataset.baslik;
            dialogAc(el("bitir-dialog"));
        });
    });

    el("bitir-form").addEventListener("submit", (event) => gonder(event, "/api/bitir", () => ({
        kart_id: Number(el("bitir-id").value),
        adet: Number(el("bitir-adet").value),
        not: el("bitir-not").value.trim(),
    }), "Üretilen adet kaydedildi."));

    // --- Teslim et ---
    document.querySelectorAll("[data-teslim]").forEach((buton) => {
        buton.addEventListener("click", () => {
            el("teslim-id").value = buton.dataset.id;
            el("teslim-not").value = "";
            el("teslim-baslik").textContent = buton.dataset.baslik;
            dialogAc(el("teslim-dialog"));
        });
    });

    el("teslim-form").addEventListener("submit", (event) => gonder(event, "/api/teslim-et", () => ({
        kart_id: Number(el("teslim-id").value),
        not: el("teslim-not").value.trim(),
    }), "Kart TESLİM EDİLDİ olarak kaydedildi."));

    // --- Not ekle ---
    document.querySelectorAll("[data-not]").forEach((buton) => {
        buton.addEventListener("click", () => {
            el("not-id").value = buton.dataset.id;
            el("not-metin").value = "";
            el("not-baslik").textContent = `${buton.dataset.baslik} · Not Ekle`;
            const gecmis = (buton.dataset.notMevcut || "").trim();
            el("not-gecmis").textContent = gecmis;
            el("not-gecmis-kutu").hidden = !gecmis;
            dialogAc(el("not-dialog"));
            el("not-metin").focus();
        });
    });

    el("not-form").addEventListener("submit", (event) => {
        if (!el("not-metin").value.trim()) {
            event.preventDefault();
            toast("Not metni boş olamaz.", "uyari");
            return;
        }
        gonder(event, "/api/not", () => ({
            kart_id: Number(el("not-id").value),
            not: el("not-metin").value.trim(),
        }), "Not eklendi.");
    });
})();
