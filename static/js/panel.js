/* Pano: arama, durum/dizgi filtreleri, sayfalama, dönem özeti ve canlılık göstergesi. */
"use strict";

(() => {
    const SAYFA_BOYUTU = 12;
    const arama = document.getElementById("panel-kart-ara");
    const bolumler = [...document.querySelectorAll("[data-panel-bolum]")];
    const sonuc = document.getElementById("panel-sonuc-sayisi");
    const temizle = document.getElementById("panel-temizle");
    const bos = document.getElementById("panel-arama-bos");
    const kpiKartlari = [...document.querySelectorAll("[data-kpi-filtre]")];
    const ARAMA_KEY = "pdgm-panel-arama";

    const sayfalar = new Map();
    bolumler.forEach((bolum) => {
        const nav = document.createElement("nav");
        nav.className = "sayfalama";
        nav.setAttribute("aria-label", `${bolum.dataset.durum} sayfalama`);
        const onceki = document.createElement("button");
        const sonraki = document.createElement("button");
        const bilgi = document.createElement("span");
        onceki.textContent = "Önceki";
        sonraki.textContent = "Sonraki";
        [onceki, sonraki].forEach((b) => { b.type = "button"; b.className = "buton buton-kucuk buton-hayalet"; });
        nav.append(onceki, bilgi, sonraki);
        bolum.appendChild(nav);
        const durum = {sayfa: 0, anahtar: "", nav, onceki, sonraki, bilgi};
        sayfalar.set(bolum, durum);
        onceki.addEventListener("click", () => { durum.sayfa -= 1; filtrele(); });
        sonraki.addEventListener("click", () => { durum.sayfa += 1; filtrele(); });
    });

    // Teslim tablosu JS ile yeniden kurulabildiği için kart listesi her seferinde canlı sorgulanır.
    const tumKartlar = () => [...document.querySelectorAll("[data-panel-kart]")];
    const durumUygun = (durum) => durumFiltresi.deger === "HEPSI" || durum === durumFiltresi.deger;
    const dizgiUygun = (tip) => dizgiFiltresi.deger === "HEPSI" || tip === dizgiFiltresi.deger;

    function filtrele() {
        const metin = aramaMetni(arama.value);
        let gorunen = 0;

        tumKartlar().forEach((kart) => {
            const uygun = durumUygun(kart.dataset.durum)
                && dizgiUygun(kart.dataset.dizgiTipi)
                && (!metin || aramaMetni(kart.dataset.arama).includes(metin));
            kart.hidden = !uygun;
            if (uygun) gorunen += 1;
        });

        let gosterilen = 0;
        bolumler.forEach((bolum) => {
            const bolumKartlari = [...bolum.querySelectorAll("[data-panel-kart]")];
            const uygunlar = bolumKartlari.filter((k) => !k.hidden);
            const sayac = bolum.querySelector(".panel-baslik .sayi-rozet");
            if (sayac) sayac.textContent = uygunlar.length;
            bolum.hidden = !durumUygun(bolum.dataset.durum) || (bolumKartlari.length > 0 && !uygunlar.length);

            const durum = sayfalar.get(bolum);
            const anahtar = `${metin}|${durumFiltresi.deger}|${dizgiFiltresi.deger}|${uygunlar.length}`;
            if (durum.anahtar !== anahtar) durum.sayfa = 0;
            durum.anahtar = anahtar;
            const sonSayfa = Math.max(0, Math.ceil(uygunlar.length / SAYFA_BOYUTU) - 1);
            durum.sayfa = Math.max(0, Math.min(durum.sayfa, sonSayfa));
            uygunlar.forEach((kart, i) => {
                kart.hidden = Math.floor(i / SAYFA_BOYUTU) !== durum.sayfa;
                if (!kart.hidden) gosterilen += 1;
            });
            durum.nav.hidden = uygunlar.length <= SAYFA_BOYUTU;
            durum.onceki.disabled = durum.sayfa === 0;
            durum.sonraki.disabled = durum.sayfa === sonSayfa;
            durum.bilgi.textContent = `${durum.sayfa + 1} / ${sonSayfa + 1} sayfa · ${uygunlar.length} eşleşme`;
        });

        sonuc.textContent = gorunen ? `${gosterilen} / ${gorunen} kart gösteriliyor` : "Sonuç bulunamadı";
        bos.hidden = gorunen !== 0;
        temizle.hidden = !metin && durumFiltresi.deger === "HEPSI" && dizgiFiltresi.deger === "HEPSI";
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
        butonlar: [...document.querySelectorAll("[data-panel-filtre]")],
        veriAdi: "panelFiltre",
        anahtar: "pdgm-panel-filtre",
        varsayilan: "HEPSI",
        degisince: () => { kpiGuncelle(); filtrele(); },
    });

    const dizgiFiltresi = filtreGrubuKur({
        butonlar: [...document.querySelectorAll("[data-panel-dizgi-filtre]")],
        veriAdi: "panelDizgiFiltre",
        anahtar: "pdgm-panel-dizgi-filtre",
        varsayilan: "HEPSI",
        degisince: () => { kpiGuncelle(); filtrele(); donemYukle(donemFiltresi.deger, true); },
    });

    kpiKartlari.forEach((kutu) => {
        kutu.addEventListener("click", () => {
            const hedef = kutu.dataset.kpiFiltre;
            durumFiltresi.sec(durumFiltresi.deger === hedef ? "HEPSI" : hedef);
            const bolum = bolumler.find((b) => b.dataset.durum === hedef);
            if (bolum && durumFiltresi.deger === hedef) bolum.scrollIntoView({behavior: "smooth", block: "start"});
        });
    });

    arama.value = oturumDeposu.al(ARAMA_KEY, "");
    arama.addEventListener("input", () => {
        oturumDeposu.yaz(ARAMA_KEY, arama.value || "");
        filtrele();
    });

    temizle.addEventListener("click", () => {
        arama.value = "";
        oturumDeposu.sil(ARAMA_KEY);
        durumFiltresi.sec("HEPSI", {sessiz: true});
        dizgiFiltresi.sec("HEPSI");
        arama.focus();
    });

    document.getElementById("panel-yenile").addEventListener("click", () => sayfayiYenile());

    // ------------------------------------------------------------------
    // Canlılık göstergesi: ortak.js'teki veri yoklamasının sonucunu gösterir.
    // ------------------------------------------------------------------
    const canli = document.querySelector("[data-canli]");
    const sonKontrol = document.querySelector("[data-son-kontrol]");
    const CANLI_METIN = {
        guncel: "Güncel",
        degisti: "Yeni veri var",
        "baglanti-yok": "Bağlantı yok",
        oturum: "Oturum sona erdi",
    };
    window.addEventListener("pdgm:veri-durumu", (event) => {
        const {durum, zaman} = event.detail;
        canli.dataset.durum = durum;
        canli.textContent = CANLI_METIN[durum] || "Güncel";
        if (zaman) sonKontrol.textContent = ` · son kontrol ${zaman}`;
    });

    // ------------------------------------------------------------------
    // Dönem filtresi: Tümü / Bu Hafta / Bu Ay / Bu Yıl / Özel Aralık
    // ------------------------------------------------------------------
    const DONEM_KEY = "pdgm-panel-donem";
    const TARIH_KEY = "pdgm-panel-tarihler";
    const donemOzelAlan = document.getElementById("donem-ozel-alan");
    const donemBaslangic = document.getElementById("donem-baslangic");
    const donemBitis = document.getElementById("donem-bitis");
    const donemBolumu = document.getElementById("bolum-teslim");
    const donemTabloGovde = document.getElementById("donem-tablo-govde");
    const donemBos = document.getElementById("donem-bos");
    const donemSayi = document.getElementById("panel-donem-sayi");
    const metrik = {
        is: document.getElementById("donem-metrik-is"),
        adetAlt: document.getElementById("donem-metrik-adet-alt"),
        zamaninda: document.getElementById("donem-metrik-zamaninda"),
        zamanindaAlt: document.getElementById("donem-metrik-zamaninda-alt"),
        gecikme: document.getElementById("donem-metrik-gecikme"),
        gecikmeAlt: document.getElementById("donem-metrik-gecikme-alt"),
        sapma: document.getElementById("donem-metrik-sapma"),
    };
    const donemVeriUyari = document.getElementById("donem-veri-uyari");
    let donemIstek = 0;

    try {
        const tarihler = JSON.parse(oturumDeposu.al(TARIH_KEY, "{}"));
        donemBaslangic.value = tarihler.baslangic || "";
        donemBitis.value = tarihler.bitis || "";
    } catch (_) {
        oturumDeposu.sil(TARIH_KEY);
    }

    const aralikGecerli = () => ggAaYyyyGecerliMi(donemBaslangic.value) && ggAaYyyyGecerliMi(donemBitis.value)
        && isoyaCevir(donemBaslangic.value) <= isoyaCevir(donemBitis.value);

    function uyariGoster(metin) {
        donemVeriUyari.hidden = !metin;
        donemVeriUyari.textContent = metin || "";
    }

    function hucreEkle(satir, deger) {
        const td = document.createElement("td");
        td.textContent = (deger === null || deger === undefined || deger === "") ? "—" : deger;
        satir.appendChild(td);
    }

    function rozet(sinif, metin) {
        const span = document.createElement("span");
        span.className = `durum-rozet ${sinif}`;
        span.textContent = metin;
        return span;
    }

    function satirKur(k) {
        const satir = document.createElement("tr");
        const dizgiKod = k.dizgi_kod || "MAKINE";
        satir.setAttribute("data-panel-kart", "");
        satir.dataset.durum = "TESLİM EDİLDİ";
        satir.dataset.dizgiTipi = dizgiKod;
        satir.dataset.arama = [k.talep_no, k.stok_no, k.talep_sahibi, k.operator, k.aciklama, k.pcb,
            k.dizgi_sorumlusu, k.dizgi_etiket].filter(Boolean).join(" ");

        const talepTd = document.createElement("td");
        const strong = document.createElement("strong");
        strong.textContent = k.talep_no || "—";
        talepTd.appendChild(strong);
        satir.appendChild(talepTd);

        hucreEkle(satir, k.stok_no);
        hucreEkle(satir, k.toplam_adet);
        hucreEkle(satir, k.plan_baslama);
        hucreEkle(satir, k.teslim);

        const rozetTd = document.createElement("td");
        rozetTd.appendChild(rozet(k.renk || "notr", k.rozet || ""));
        if (dizgiKod === "ELLE") rozetTd.appendChild(rozet("elle", "Elle Dizgi"));
        if (dizgiKod === "EUM") rozetTd.appendChild(rozet("eum", "EÜM'de Dizgi"));
        satir.appendChild(rozetTd);
        return satir;
    }

    function metrikleriYaz(o) {
        const adet = (sayi) => sayiBicimle(sayi, 0);
        donemSayi.textContent = o.kart;
        metrik.is.textContent = o.kart;
        metrik.adetAlt.textContent = `${adet(o.adet)} adet kart teslim edildi`;
        metrik.zamaninda.textContent = `%${o.zamaninda_yuzde}`;
        metrik.zamanindaAlt.textContent = `${o.zamaninda} iş emri · ${adet(o.zamaninda_adet)} adet`;
        metrik.gecikme.textContent = o.gecikmeli;
        metrik.gecikmeAlt.textContent = `${adet(o.gecikmeli_adet)} adet`;
        metrik.sapma.textContent = o.sapma_olculen ? `${sayiBicimle(o.ort_sapma)} gün` : "—";
        const olculemeyen = o.kart - o.sapma_olculen;
        uyariGoster(olculemeyen
            ? `${olculemeyen} iş emrinde plan veya gerçekleşen teslim tarihi eksik olduğu için sapma hesaplanamadı; `
              + `yüzde yalnız ${o.sapma_olculen} iş emri üzerinden hesaplandı.`
            : "");
    }

    function metrikleriBosalt() {
        donemTabloGovde.replaceChildren();
        donemBos.hidden = false;
        donemSayi.textContent = "—";
        Object.values(metrik).forEach((e) => { e.textContent = "—"; });
        filtrele();
    }

    async function donemYukle(aralik, sessizce = false) {
        const istek = ++donemIstek;
        const params = new URLSearchParams({aralik, dizgi: dizgiFiltresi.deger});

        if (aralik === "ozel") {
            if (!aralikGecerli()) {
                if (!sessizce) toast("Geçerli bir başlangıç ve bitiş tarihi girin; başlangıç bitişten sonra olamaz.", "uyari");
                metrikleriBosalt();
                uyariGoster("Tarih aralığını girip Uygula'ya basın.");
                return;
            }
            params.set("baslangic", isoyaCevir(donemBaslangic.value));
            params.set("bitis", isoyaCevir(donemBitis.value));
        }

        // Tablo boşaltılmaz: yeni veri gelene kadar eskisi görünür kalır (titreme olmaz).
        donemBolumu.setAttribute("aria-busy", "true");
        try {
            const data = await pdgmFetch(`/api/panel/teslimler?${params.toString()}`);
            if (istek !== donemIstek) return;
            if (aralik === "ozel") {
                oturumDeposu.yaz(TARIH_KEY, JSON.stringify({baslangic: donemBaslangic.value, bitis: donemBitis.value}));
            }
            const liste = data.teslim_edilen || [];
            donemTabloGovde.replaceChildren(...liste.map(satirKur));
            donemBos.hidden = liste.length !== 0;
            metrikleriYaz(data.ozet);
            filtrele();
        } catch (hata) {
            if (istek !== donemIstek) return;
            metrikleriBosalt();
            uyariGoster("Veriler yüklenemedi; tekrar deneyin.");
            hataMesaji(hata);
        } finally {
            if (istek === donemIstek) donemBolumu.removeAttribute("aria-busy");
        }
    }

    const donemFiltresi = filtreGrubuKur({
        butonlar: [...document.querySelectorAll("[data-donem-filtre]")],
        veriAdi: "donemFiltre",
        anahtar: DONEM_KEY,
        varsayilan: "tumu",
        degisince: (deger) => {
            donemOzelAlan.hidden = deger !== "ozel";
            donemYukle(deger, true);
        },
    });

    document.getElementById("donem-uygula").addEventListener("click", () => donemYukle("ozel"));
    [donemBaslangic, donemBitis].forEach((input) => {
        input.addEventListener("input", () => {
            donemIstek += 1;
            uyariGoster("Yeni tarih aralığı için Uygula'ya basın.");
        });
        input.addEventListener("keydown", (event) => {
            if (event.key === "Enter") donemYukle("ozel");
        });
    });

    // Açılış: sunucunun çizdiği tablo ve metrikler "Tümü / Tümü" için zaten doğru;
    // yalnız başka bir dönem ya da dizgi tipi hatırlanıyorsa API'ye gidilir.
    if (donemFiltresi.deger === "ozel" && !aralikGecerli()) donemFiltresi.sec("tumu", {sessiz: true});
    donemOzelAlan.hidden = donemFiltresi.deger !== "ozel";
    kpiGuncelle();
    filtrele();
    if (donemFiltresi.deger !== "tumu" || dizgiFiltresi.deger !== "HEPSI") {
        donemYukle(donemFiltresi.deger, true);
    }
})();
