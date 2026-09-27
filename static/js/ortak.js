/* PDGM İş Takip — tüm sayfalarda yüklenen ortak yardımcılar.
   Sayfa betikleri (panel.js, operator.js, ...) bu dosyadan SONRA yüklenir. */
"use strict";

function csrfToken() {
    return document.querySelector('meta[name="csrf-token"]')?.content || "";
}

function aramaMetni(deger) {
    return String(deger || "").normalize("NFC").toLocaleLowerCase("tr-TR").normalize("NFC").trim();
}

async function pdgmFetch(url, options = {}) {
    const headers = new Headers(options.headers || {});
    headers.set("X-CSRF-Token", csrfToken());

    if (options.body && !(options.body instanceof FormData) && !headers.has("Content-Type")) {
        headers.set("Content-Type", "application/json");
    }

    const response = await fetch(url, {credentials: "same-origin", ...options, headers});
    const contentType = response.headers.get("content-type") || "";
    if (response.status === 401 || response.redirected) {
        throw new Error("Oturum sona erdi. İşlem kaydedilmedi; tekrar giriş yapın.");
    }
    if (!contentType.includes("application/json")) {
        throw new Error("Sunucudan beklenmeyen yanıt alındı. İşlemin kaydedildiği doğrulanamadı.");
    }
    const data = await response.json();

    if (!response.ok) {
        throw new Error(data.hata || data.detail || "İşlem tamamlanamadı.");
    }
    return data;
}

// ---------------------------------------------------------------------------
// Tarayıcı depolaması: gizli pencere / kısıtlı profilde erişim hata fırlatabilir,
// bu durumda sayfa depolamasız çalışmaya devam eder.
// ---------------------------------------------------------------------------
function depoOlustur(getir) {
    return {
        al(anahtar, varsayilan = null) {
            try {
                const deger = getir().getItem(anahtar);
                return deger === null ? varsayilan : deger;
            } catch (_) {
                return varsayilan;
            }
        },
        yaz(anahtar, deger) {
            try { getir().setItem(anahtar, String(deger)); } catch (_) { /* depolama kapalı */ }
        },
        sil(anahtar) {
            try { getir().removeItem(anahtar); } catch (_) { /* depolama kapalı */ }
        },
    };
}

const oturumDeposu = depoOlustur(() => window.sessionStorage);
const kaliciDepo = depoOlustur(() => window.localStorage);

// ---------------------------------------------------------------------------
// Bildirimler: başarı kendiliğinden kapanır; hata ve uyarı kullanıcı kapatana kadar kalır.
// ---------------------------------------------------------------------------
const TOAST_SINIRI = 4;

function toast(mesaj, tip = "basari", {kalici} = {}) {
    const alan = document.getElementById("toast-alani");
    if (!alan) return;

    const kalsin = kalici ?? (tip === "hata" || tip === "uyari");
    const kutu = document.createElement("div");
    kutu.className = `toast ${tip}`;
    if (tip === "hata") kutu.setAttribute("role", "alert");

    const metin = document.createElement("span");
    metin.textContent = mesaj;
    kutu.appendChild(metin);

    const kapat = () => {
        kutu.classList.remove("goster");
        window.setTimeout(() => kutu.remove(), 200);
    };
    if (kalsin) {
        const buton = document.createElement("button");
        buton.type = "button";
        buton.className = "toast-kapat";
        buton.setAttribute("aria-label", "Bildirimi kapat");
        buton.textContent = "×";
        buton.addEventListener("click", kapat);
        kutu.appendChild(buton);
    } else {
        window.setTimeout(kapat, 4000);
    }

    while (alan.children.length >= TOAST_SINIRI) alan.firstElementChild.remove();
    alan.appendChild(kutu);
    requestAnimationFrame(() => kutu.classList.add("goster"));
}

function hataMesaji(hata) {
    toast(hata?.message || "Beklenmeyen bir hata oluştu.", "hata");
}

// ---------------------------------------------------------------------------
// Sayfayı yenile ama bildirimi ve kaydırma konumunu kaybetme.
// ---------------------------------------------------------------------------
const BEKLEYEN_BILDIRIM = "pdgm-bekleyen-bildirim";
const KAYDIRMA_ANAHTARI = `pdgm-kaydirma:${window.location.pathname}`;

function kaydirmaKaydet() {
    oturumDeposu.yaz(KAYDIRMA_ANAHTARI, String(window.scrollY));
}

function sayfayiYenile() {
    kaydirmaKaydet();
    window.dispatchEvent(new CustomEvent("pdgm:yenilenecek"));
    window.location.reload();
}

function yenileVeBildir(mesaj, tip = "basari") {
    if (mesaj) oturumDeposu.yaz(BEKLEYEN_BILDIRIM, JSON.stringify({mesaj, tip}));
    sayfayiYenile();
}

// DOMContentLoaded, sayfa betikleri filtreleri uyguladıktan SONRA gelir; kaydırma
// konumu ancak o zaman doğru yere denk gelir.
document.addEventListener("DOMContentLoaded", () => {
    const kaydirma = oturumDeposu.al(KAYDIRMA_ANAHTARI);
    if (kaydirma !== null) {
        oturumDeposu.sil(KAYDIRMA_ANAHTARI);
        window.scrollTo(0, Number(kaydirma) || 0);
    }
    const bekleyen = oturumDeposu.al(BEKLEYEN_BILDIRIM);
    if (bekleyen) {
        oturumDeposu.sil(BEKLEYEN_BILDIRIM);
        try {
            const {mesaj, tip} = JSON.parse(bekleyen);
            if (mesaj) toast(mesaj, tip);
        } catch (_) { /* bozuk kayıt: yok say */ }
    }
});

// ---------------------------------------------------------------------------
// Dialoglar ve onay
// ---------------------------------------------------------------------------
function dialogKapat(dialog) {
    if (dialog?.open) dialog.close();
}

document.addEventListener("click", (event) => {
    const kapat = event.target.closest("[data-dialog-kapat]");
    if (kapat) dialogKapat(kapat.closest("dialog"));
});

/** Tarayıcının confirm() kutusu yerine uygulamanın kendi onay dialogu. Promise<boolean>. */
function onayIste({baslik = "Emin misiniz?", mesaj = "", etiket = "ONAY", evet = "Devam et", tehlike = false} = {}) {
    const dialog = document.getElementById("onay-dialog");
    if (!dialog) return Promise.resolve(window.confirm(mesaj || baslik));

    dialog.querySelector("#onay-etiket").textContent = etiket;
    dialog.querySelector("#onay-baslik").textContent = baslik;
    dialog.querySelector("#onay-mesaj").textContent = mesaj;
    const evetButonu = dialog.querySelector("#onay-evet");
    evetButonu.textContent = evet;
    evetButonu.classList.toggle("buton-tehlike", tehlike);
    evetButonu.classList.toggle("buton-ana", !tehlike);

    dialog.returnValue = "";
    dialog.showModal();
    evetButonu.focus();
    return new Promise((resolve) => {
        dialog.addEventListener("close", () => resolve(dialog.returnValue === "evet"), {once: true});
    });
}

// <form data-onay="Mesaj"> gönderilmeden önce onay ister (satır içi onsubmit yerine).
document.addEventListener("submit", async (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || !form.dataset.onay) return;
    if (form.dataset.onayVerildi === "1") return;
    event.preventDefault();
    const submitter = event.submitter;
    const tamam = await onayIste({
        baslik: form.dataset.onayBaslik || "Emin misiniz?",
        mesaj: form.dataset.onay,
        evet: form.dataset.onayEvet || "Devam et",
        tehlike: form.dataset.onayTehlike === "1",
    });
    if (!tamam) return;
    form.dataset.onayVerildi = "1";
    form.requestSubmit(submitter && submitter.form === form ? submitter : undefined);
});

// ---------------------------------------------------------------------------
// Filtre buton grupları (aria-pressed + seçim hatırlama)
// ---------------------------------------------------------------------------
function filtreGrubuKur({butonlar, veriAdi, anahtar, varsayilan, degisince}) {
    let aktif = anahtar ? oturumDeposu.al(anahtar, varsayilan) : varsayilan;
    if (!butonlar.some((buton) => buton.dataset[veriAdi] === aktif)) aktif = varsayilan;

    function sec(deger, {sessiz = false} = {}) {
        aktif = deger;
        if (anahtar) oturumDeposu.yaz(anahtar, deger);
        butonlar.forEach((buton) => {
            const secili = buton.dataset[veriAdi] === deger;
            buton.classList.toggle("aktif", secili);
            buton.setAttribute("aria-pressed", secili ? "true" : "false");
        });
        if (!sessiz && degisince) degisince(deger);
    }

    butonlar.forEach((buton) => buton.addEventListener("click", () => sec(buton.dataset[veriAdi])));
    sec(aktif, {sessiz: true});
    return {
        get deger() { return aktif; },
        sec,
    };
}

// ---------------------------------------------------------------------------
// Tarih alanları: her tarayıcı dilinde aynı gg.aa.yyyy biçimi.
// ---------------------------------------------------------------------------
function ggAaYyyyGecerliMi(deger) {
    if (!/^\d{2}\.\d{2}\.\d{4}$/.test(deger || "")) return false;
    const [g, a, y] = deger.split(".").map(Number);
    const tarih = new Date(y, a - 1, g);
    return tarih.getFullYear() === y && tarih.getMonth() === a - 1 && tarih.getDate() === g;
}

/** "gg.aa.yyyy" → "yyyy-aa-gg"; boş ise "". */
function isoyaCevir(deger) {
    if (!deger) return "";
    const [gg, aa, yyyy] = deger.split(".");
    return `${yyyy}-${aa}-${gg}`;
}

/** "yyyy-aa-gg..." → "gg.aa.yyyy"; boş/geçersiz ise "". */
function isodanGoster(deger) {
    const eslesme = /^(\d{4})-(\d{2})-(\d{2})/.exec(deger || "");
    return eslesme ? `${eslesme[3]}.${eslesme[2]}.${eslesme[1]}` : "";
}

function tarihAlaniniDogrula(input) {
    const deger = input.value.trim();
    input.setCustomValidity(deger && !ggAaYyyyGecerliMi(deger) ? "Geçerli bir tarih girin (gg.aa.yyyy)." : "");
}

function tarihAlaniKur(input) {
    input.addEventListener("input", () => {
        // İmleç ortadayken düzenleme yapılabilsin: biçimlendirmeden sonra imleç,
        // öncesindeki rakam sayısına göre yeniden konumlanır (sona atlamaz).
        const imlec = input.selectionStart ?? input.value.length;
        const oncekiRakam = input.value.slice(0, imlec).replace(/\D/g, "").length;
        const rakam = input.value.replace(/\D/g, "").slice(0, 8);
        input.value = [rakam.slice(0, 2), rakam.slice(2, 4), rakam.slice(4, 8)].filter(Boolean).join(".");

        let konum = 0;
        for (let sayilan = 0; konum < input.value.length && sayilan < oncekiRakam; konum += 1) {
            if (/\d/.test(input.value[konum])) sayilan += 1;
        }
        if (document.activeElement === input) input.setSelectionRange(konum, konum);
        tarihAlaniniDogrula(input);
    });
    input.addEventListener("blur", () => tarihAlaniniDogrula(input));
}

document.querySelectorAll("input[data-tarih]").forEach(tarihAlaniKur);

function sayiBicimle(deger, basamak = 1) {
    const sayi = Number(deger);
    return Number.isFinite(sayi) ? sayi.toLocaleString("tr-TR", {maximumFractionDigits: basamak}) : "—";
}

function saatMetni(tarih = new Date()) {
    return new Intl.DateTimeFormat("tr-TR", {hour: "2-digit", minute: "2-digit", second: "2-digit"}).format(tarih);
}

// ---------------------------------------------------------------------------
// Veri değişikliği yoklaması: başka bir ekranda yapılan değişiklikleri ve sunucu
// bağlantısını izler. <body data-veri-surumu="..."> olan sayfalarda çalışır.
// Kullanıcı bir süredir dokunmuyorsa sayfayı kendisi yeniler; çalışıyorsa bant gösterir.
// ---------------------------------------------------------------------------
const VERI_YOKLAMA_MS = 30000;
const BOSTA_YENILE_MS = 60000;

function veriDurumuYay(durum, ayrinti = {}) {
    window.dispatchEvent(new CustomEvent("pdgm:veri-durumu", {detail: {durum, ...ayrinti}}));
}

function veriDegisiminiIzle() {
    const surum = document.body.dataset.veriSurumu;
    const bant = document.getElementById("veri-bandi");
    if (!surum || !bant) return;

    const bantMetni = bant.querySelector("[data-veri-bandi-metin]");
    let sonEtkilesim = Date.now();
    let degisti = false;
    let hataSayisi = 0;
    let sonBasarili = saatMetni();

    ["pointerdown", "keydown", "input", "wheel", "touchstart", "scroll"].forEach((olay) => {
        window.addEventListener(olay, () => { sonEtkilesim = Date.now(); }, {passive: true, capture: true});
    });
    bant.querySelector("[data-veri-bandi-yenile]").addEventListener("click", () => sayfayiYenile());

    function bantGoster(tip, metin) {
        bant.hidden = false;
        bant.dataset.tip = tip;
        bantMetni.textContent = metin;
    }

    // Sekme görünmüyorsa ya da kullanıcı bir süredir dokunmuyorsa yenilemek hiçbir işi bölmez;
    // açık dialog veya yazılan bir alan varsa kayıtsız girdi kaybolmasın diye yenilenmez.
    function kullaniciBostaMi() {
        const odak = document.activeElement;
        const yaziyor = odak && odak.matches("input:not([type=search]), textarea, select");
        const bakmiyor = document.hidden || Date.now() - sonEtkilesim >= BOSTA_YENILE_MS;
        return bakmiyor && !document.querySelector("dialog[open]") && !yaziyor;
    }

    async function yokla() {
        const iptal = new AbortController();
        const zamanAsimi = window.setTimeout(() => iptal.abort(), 8000);
        try {
            const yanit = await fetch("/api/surum", {credentials: "same-origin", cache: "no-store", signal: iptal.signal});
            if (yanit.status === 401 || yanit.redirected) {
                bantGoster("hata", "Oturum sona erdi. Devam etmek için sayfayı yenileyip tekrar giriş yapın.");
                veriDurumuYay("oturum");
                return;  // yoklamayı durdur
            }
            if (!yanit.ok) throw new Error(String(yanit.status));
            const veri = await yanit.json();
            hataSayisi = 0;
            sonBasarili = veri.zaman || saatMetni();
            degisti = degisti || veri.surum !== surum;
            if (degisti) {
                if (kullaniciBostaMi()) {
                    sayfayiYenile();
                    return;
                }
                bantGoster("degisti", "Veriler başka bir ekranda güncellendi. Güncel hâli görmek için yenileyin.");
                veriDurumuYay("degisti", {zaman: sonBasarili});
            } else {
                bant.hidden = true;
                veriDurumuYay("guncel", {zaman: sonBasarili});
            }
        } catch (_) {
            hataSayisi += 1;
            bantGoster("hata", `Sunucuya ulaşılamıyor. Ekrandaki bilgiler en son ${sonBasarili} itibarıyla doğrulandı; yeniden deneniyor.`);
            veriDurumuYay("baglanti-yok", {zaman: sonBasarili});
        } finally {
            window.clearTimeout(zamanAsimi);
        }
        const bekle = hataSayisi ? Math.min(VERI_YOKLAMA_MS * 2, 10000 * 2 ** (hataSayisi - 1)) : VERI_YOKLAMA_MS;
        window.setTimeout(yokla, bekle);
    }

    window.setTimeout(yokla, VERI_YOKLAMA_MS);
}

veriDegisiminiIzle();
