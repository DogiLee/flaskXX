/* Yönetim ekranı: kart tablosu filtreleri, düzenleme / yeni kart dialogları, gizleme. */
"use strict";

(() => {
    const el = (id) => document.getElementById(id);
    const adminAra = el("admin-kart-ara");
    const satirlar = [...document.querySelectorAll("[data-admin-kart]")];
    const sonuc = el("admin-sonuc");
    const bos = el("admin-kart-bos");
    const temizle = el("admin-temizle");
    const tabloKaydirma = el("admin-tablo-kaydirma");
    const ARAMA_KEY = "pdgm-yon-arama";
    const TABLO_KAYDIRMA_KEY = "pdgm-yon-tablo-kaydirma";

    // ------------------------------------------------------------------
    // Filtreler: durum ve dizgi tipi ayrı gruplar, birlikte uygulanır.
    // Seçimler, arama ve tablo kaydırması kayıttan sonraki yenilemede korunur.
    // ------------------------------------------------------------------
    function durumUygun(satir) {
        const filtre = durumFiltresi.deger;
        if (filtre === "HEPSI") return true;
        if (filtre === "AKTIF") return ["PLANA ALINDI", "DİZGİDE", "HAZIR"].includes(satir.dataset.durum);
        return satir.dataset.durum === filtre;
    }

    function filtrele() {
        const arama = aramaMetni(adminAra.value);
        let gorunen = 0;
        satirlar.forEach((satir) => {
            const uygun = durumUygun(satir)
                && (dizgiFiltresi.deger === "HEPSI" || satir.dataset.dizgiKod === dizgiFiltresi.deger)
                && (!arama || aramaMetni(satir.dataset.arama).includes(arama));
            satir.hidden = !uygun;
            if (uygun) gorunen += 1;
        });
        sonuc.textContent = gorunen === satirlar.length ? `${gorunen} kart` : `${gorunen} / ${satirlar.length} kart`;
        bos.hidden = gorunen !== 0 || satirlar.length === 0;
        temizle.hidden = !arama && durumFiltresi.deger === "HEPSI" && dizgiFiltresi.deger === "HEPSI";
    }

    const durumFiltresi = filtreGrubuKur({
        butonlar: [...document.querySelectorAll("[data-admin-filtre]")],
        veriAdi: "adminFiltre",
        anahtar: "pdgm-yon-durum",
        varsayilan: "HEPSI",
        degisince: filtrele,
    });
    const dizgiFiltresi = filtreGrubuKur({
        butonlar: [...document.querySelectorAll("[data-admin-dizgi-filtre]")],
        veriAdi: "adminDizgiFiltre",
        anahtar: "pdgm-yon-dizgi",
        varsayilan: "HEPSI",
        degisince: filtrele,
    });

    adminAra.value = oturumDeposu.al(ARAMA_KEY, "");
    adminAra.addEventListener("input", () => {
        oturumDeposu.yaz(ARAMA_KEY, adminAra.value);
        filtrele();
    });
    temizle.addEventListener("click", () => {
        adminAra.value = "";
        oturumDeposu.sil(ARAMA_KEY);
        durumFiltresi.sec("HEPSI", {sessiz: true});
        dizgiFiltresi.sec("HEPSI");
        adminAra.focus();
    });
    filtrele();

    window.addEventListener("pdgm:yenilenecek", () => {
        oturumDeposu.yaz(TABLO_KAYDIRMA_KEY, String(tabloKaydirma.scrollTop));
    });
    const tabloKonumu = oturumDeposu.al(TABLO_KAYDIRMA_KEY);
    if (tabloKonumu !== null) {
        oturumDeposu.sil(TABLO_KAYDIRMA_KEY);
        tabloKaydirma.scrollTop = Number(tabloKonumu) || 0;
    }

    // ------------------------------------------------------------------
    // Ortak dialog yardımcıları
    // ------------------------------------------------------------------
    function dizgiAlanlariniGuncelle(selectId, sarmalayiciId) {
        el(sarmalayiciId).hidden = el(selectId).value !== "ELLE DİZGİ";
    }

    function dizgiKoduToDeger(kod) {
        if (kod === "ELLE") return "ELLE DİZGİ";
        if (kod === "EUM") return "EÜM'DE DİZGİ";
        return "MAKİNE";
    }

    function tarihYaz(id, iso) {
        el(id).value = isodanGoster(iso);
        tarihAlaniniDogrula(el(id));
    }

    const tarihOku = (id) => isoyaCevir(el(id).value.trim());

    async function gonder(event, url, govde, mesaj) {
        event.preventDefault();
        const form = event.target;
        const submit = event.submitter || form.querySelector("[type=submit]");
        submit.disabled = true;
        try {
            await pdgmFetch(url, {method: "POST", body: JSON.stringify(govde())});
            dialogKapat(form.closest("dialog"));
            yenileVeBildir(mesaj);
        } catch (hata) {
            hataMesaji(hata);
            submit.disabled = false;
        }
    }

    // ------------------------------------------------------------------
    // Kart düzenleme
    // ------------------------------------------------------------------
    const adminForm = el("admin-form");
    const durumSecimi = el("admin-durum");
    const toplamAlani = el("admin-toplam");
    const tamamlananAlani = el("admin-tamamlanan");
    const malzemeKutusu = el("admin-malzeme-bekliyor");
    let serbestTamamlanan = "0";   // DİZGİDE'ye dönülünce geri gelecek değer
    let malzemeIlk = false;

    // PLANA ALINDI / HAZIR tamamlananı 0, TESLİM EDİLDİ toplamı yapar (sunucu da böyle kaydeder).
    // Bu durumlarda alan kilitlenir; DİZGİDE'ye dönülünce elle girilen değer geri gelir.
    function durumAlanlariniAyarla() {
        const durum = durumSecimi.value;
        const toplam = Number(toplamAlani.value || 0);
        tamamlananAlani.max = toplam > 0 ? String(toplam) : "";
        if (durum === "PLANA ALINDI" || durum === "HAZIR") {
            tamamlananAlani.value = 0;
            tamamlananAlani.readOnly = true;
        } else if (durum === "TESLİM EDİLDİ") {
            tamamlananAlani.value = toplam > 0 ? toplam : tamamlananAlani.value;
            tamamlananAlani.readOnly = true;
        } else {
            tamamlananAlani.readOnly = false;
            tamamlananAlani.value = serbestTamamlanan;
        }
    }

    function duzenlemeDialogunuAc(buton) {
        const v = buton.dataset;
        adminForm.dataset.surum = v.surum;
        el("admin-id").value = v.id;
        durumSecimi.value = v.durum || "";
        toplamAlani.value = v.toplam;
        serbestTamamlanan = v.tamamlanan || "0";
        tamamlananAlani.value = serbestTamamlanan;
        el("admin-plan-hafta").value = v.planHafta || "";
        tarihYaz("admin-plan-baslama", v.planBaslama);
        tarihYaz("admin-plan-teslim", v.planTeslim);
        tarihYaz("admin-gerceklesen-teslim", v.gerceklesenTeslim);
        el("admin-not").value = v.not || "";
        el("admin-dizgi-tipi").value = dizgiKoduToDeger(v.dizgiKod);
        el("admin-dizgi-sorumlusu").value = v.dizgiSorumlusu || "";
        // Öznitelik yoksa işaret "bilinmiyor" sayılır ve dokunulmadıkça sunucuya gönderilmez.
        malzemeIlk = v.malzeme === "1";
        malzemeKutusu.checked = malzemeIlk;
        el("admin-baslik").textContent = `${v.talep || "Kart"} · ${v.stok || ""}`;
        dizgiAlanlariniGuncelle("admin-dizgi-tipi", "admin-elle-alanlari");
        durumAlanlariniAyarla();
        el("admin-dialog").showModal();
    }

    document.querySelectorAll("[data-admin-duzenle]").forEach((buton) => {
        buton.addEventListener("click", () => duzenlemeDialogunuAc(buton));
    });
    durumSecimi.addEventListener("change", durumAlanlariniAyarla);
    toplamAlani.addEventListener("input", durumAlanlariniAyarla);
    tamamlananAlani.addEventListener("input", () => {
        if (!tamamlananAlani.readOnly) serbestTamamlanan = tamamlananAlani.value;
    });
    el("admin-dizgi-tipi").addEventListener("change", () => dizgiAlanlariniGuncelle("admin-dizgi-tipi", "admin-elle-alanlari"));

    adminForm.addEventListener("submit", (event) => gonder(event, "/api/admin/duzenle", () => ({
        kart_id: Number(el("admin-id").value),
        surum: adminForm.dataset.surum,
        durum: durumSecimi.value,
        toplam_adet: Number(toplamAlani.value),
        tamamlanan_adet: Number(tamamlananAlani.value),
        plan_hafta: el("admin-plan-hafta").value,
        plan_baslama: tarihOku("admin-plan-baslama"),
        plan_teslim: tarihOku("admin-plan-teslim"),
        gerceklesen_teslim: tarihOku("admin-gerceklesen-teslim"),
        not: el("admin-not").value,
        dizgi_tipi: el("admin-dizgi-tipi").value,
        dizgi_sorumlusu: el("admin-dizgi-sorumlusu").value,
        // null = değiştirilmedi; sunucu mevcut değeri korur.
        malzeme_bekliyor: malzemeKutusu.checked === malzemeIlk ? null : malzemeKutusu.checked,
    }), "Kart güncellendi."));

    // ------------------------------------------------------------------
    // Yeni kart
    // ------------------------------------------------------------------
    const yeniForm = el("admin-yeni-form");
    el("yeni-dizgi-tipi").addEventListener("change", () => dizgiAlanlariniGuncelle("yeni-dizgi-tipi", "yeni-elle-alanlari"));
    el("admin-yeni-ac").addEventListener("click", () => {
        yeniForm.reset();
        yeniForm.querySelectorAll("input[data-tarih]").forEach((input) => input.setCustomValidity(""));
        el("yeni-dizgi-tipi").value = "MAKİNE";
        dizgiAlanlariniGuncelle("yeni-dizgi-tipi", "yeni-elle-alanlari");
        el("admin-yeni-dialog").showModal();
        el("yeni-talep-no").focus();
    });

    yeniForm.addEventListener("submit", (event) => gonder(event, "/api/admin/kart-ekle", () => {
        const sira = el("yeni-sira").value;
        return {
            sira: sira ? Number(sira) : null,
            talep_no: el("yeni-talep-no").value,
            talep_sahibi: el("yeni-talep-sahibi").value,
            stok_no: el("yeni-stok-no").value,
            toplam_adet: Number(el("yeni-toplam").value),
            plan_hafta: el("yeni-plan-hafta").value,
            plan_baslama: tarihOku("yeni-plan-baslama"),
            plan_teslim: tarihOku("yeni-plan-teslim"),
            pcb: el("yeni-pcb").value,
            dizgi_tipi: el("yeni-dizgi-tipi").value,
            dizgi_sorumlusu: el("yeni-dizgi-sorumlusu").value,
            not: el("yeni-not").value,
        };
    }, "Yeni kart PLANA ALINDI durumunda oluşturuldu."));

    // ------------------------------------------------------------------
    // Gizle
    // ------------------------------------------------------------------
    document.querySelectorAll("[data-admin-gizle]").forEach((buton) => {
        buton.addEventListener("click", async () => {
            const tamam = await onayIste({
                baslik: `${buton.dataset.talep || "Bu kart"} listeden gizlensin mi?`,
                mesaj: "Kart silinmez; kartlar.xlsx içinde tutulur ve Gizlenen Kartlar listesinden geri getirilebilir.",
                evet: "Gizle",
                tehlike: true,
            });
            if (!tamam) return;
            buton.disabled = true;
            try {
                await pdgmFetch("/api/admin/kart-sil", {
                    method: "POST",
                    body: JSON.stringify({kart_id: Number(buton.dataset.id)}),
                });
                yenileVeBildir("Kart listeden gizlendi.");
            } catch (hata) {
                hataMesaji(hata);
                buton.disabled = false;
            }
        });
    });
})();
