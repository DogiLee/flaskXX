"""PDGM · Baskı Dizgi Atölyesi İş Takip Sistemi.

Flask + Excel tabanlı intranet uygulaması.

Mimari sınırlar:
- kartlar.xlsx uygulamanın source of truth dosyasıdır.
- Tek Python process + çok thread kullanılır.
- Storage concurrency ve atomic write sorumluluğu depo.py'dedir.
- Workflow: HAZIR -> PLANA ALINDI -> DİZGİDE -> TESLİM EDİLDİ.
"""

from __future__ import annotations

import json
import hashlib
import logging
import logging.handlers
import os
import re
import secrets
from collections import OrderedDict,deque
from datetime import date, datetime, timedelta
from functools import wraps
from urllib.parse import urlsplit
from uuid import uuid4
import threading
import time
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

import depo
import excel_araclari as ex
from dotenv import load_dotenv

KOK = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(KOK, ".env"))

VERI_KLASORU = os.path.join(KOK, "data")
YUKLEME_KLASORU = os.path.join(VERI_KLASORU, "yuklenen_exceller")
KULLANICI_DOSYASI = os.path.join(VERI_KLASORU, "kullanicilar.json")
LOG_DOSYASI = os.path.join(VERI_KLASORU, "uygulama.log")
SUNUCU_PORTU = int(os.environ.get("PDGM_PORT", "5001"))
DINLENEN_ADRES = os.environ.get("PDGM_BIND", "0.0.0.0")

app = Flask(__name__)
depo.process_kilidi_al()



# ---------------------------------------------------------------------------
# Config ve kullanıcı dosyası
# ---------------------------------------------------------------------------

def _anahtar() -> str:
    """Session secret'ını ilk çalıştırmada üretir ve diskte saklar."""
    yol = os.path.join(VERI_KLASORU, "gizli.key")
    os.makedirs(os.path.dirname(yol), exist_ok=True)

    if not os.path.exists(yol):
        with open(yol, "w", encoding="utf-8") as f:
            f.write(secrets.token_hex(32))
        try:
            os.chmod(yol, 0o600)
        except OSError:
            pass

    with open(yol, encoding="utf-8") as f:
        anahtar = f.read().strip()
    if len(anahtar) < 32:
        raise RuntimeError("data/gizli.key geçersiz veya çok kısa.")
    return anahtar


app.secret_key = _anahtar()
app.config.update(
    MAX_CONTENT_LENGTH=25 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("PDGM_HTTPS", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    TEMPLATES_AUTO_RELOAD=True,
)

@app.after_request
def _guvenlik_basliklari(response):
    response.headers.setdefault(
        "X-Content-Type-Options",
        "nosniff",
    )
    response.headers.setdefault(
        "X-Frame-Options",
        "DENY",
    )
    response.headers.setdefault(
        "Referrer-Policy",
        "same-origin",
    )
    response.headers.setdefault(
        "Permissions-Policy",
        "camera=(), microphone=(), geolocation=()",
    )
    # Şablonlarda satır içi <script>, style="..." ve onclick/onsubmit yoktur; tüm JS
    # static/js/ altındadır. Yeni kod da bu kurala uymalı, yoksa tarayıcı engeller.
    response.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; "
        "font-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'self'; "
        "form-action 'self'; frame-ancestors 'none'",
    )

    if not request.path.startswith("/static/"):
        response.headers.setdefault(
            "Cache-Control",
            "no-store",
        )

    return response


@app.url_defaults
def _statik_surum(endpoint, values):
    """Statik dosya adresine dosyanın değişiklik zamanını ekler (?v=...).

    Elle sürüm artırmaya gerek kalmaz; dosya değişince tarayıcı yenisini ister.
    """
    if endpoint != "static" or "v" in values or not values.get("filename"):
        return
    try:
        values["v"] = int(os.stat(os.path.join(app.static_folder, values["filename"])).st_mtime)
    except OSError:
        pass



def _gunluk_dosya_logu_kur():
    os.makedirs(VERI_KLASORU, exist_ok=True)
    if any(isinstance(handler, logging.FileHandler) for handler in app.logger.handlers):
        return

    handler = logging.handlers.RotatingFileHandler(
        LOG_DOSYASI,
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    app.logger.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    logging.getLogger("waitress").setLevel(logging.INFO)


def _kullanicilari_yukle() -> dict[str, dict]:
    """İlk çalışmada .env şifrelerinden hash üretir."""

    os.makedirs(VERI_KLASORU, exist_ok=True)

    if os.path.exists(KULLANICI_DOSYASI):
        with open(KULLANICI_DOSYASI, encoding="utf-8") as f:
            veri = json.load(f)

        if not isinstance(veri, dict) or not veri:
            raise RuntimeError("data/kullanicilar.json geçersiz.")

        return veri


    baslangic = {
        "admin": {
        "rol": "admin",
        "ad": "Sistem Yöneticisi",
        "env": "PDGM_ADMIN_PASSWORD",
        "operator_tipi": None,
    },

    "makine1": {
        "rol": "operator",
        "ad": "Makine Operatörü",
        "env": "PDGM_OPERATOR_PASSWORD",
        "operator_tipi": "makine",
    },

    "elle1": {
        "rol": "operator",
        "ad": "Elle Dizgi Operatörü",
        "env": "PDGM_ELLE_PASSWORD",
        "operator_tipi": "elle_dizgi",
    },

    "eum1": {
        "rol": "operator",
        "ad": "EÜM Dizgi Operatörü",
        "env": "PDGM_EUM_PASSWORD",
        "operator_tipi": "eum_dizgi",
    },

    "gozlemci": {
        "rol": "gozlemci",
        "ad": "Gözlemci",
        "env": "PDGM_VIEWER_PASSWORD",
        "operator_tipi": None,
    },
}

    sonuc = {}

    for kullanici, bilgi in baslangic.items():

        parola = os.getenv(bilgi["env"])

        sonuc[kullanici] = {
            "sifre_hash": generate_password_hash(parola),
            "rol": bilgi["rol"],
            "ad": bilgi["ad"],
            "operator_tipi": bilgi["operator_tipi"],
            "aktif": True,
        }

    gecici = KULLANICI_DOSYASI + ".yeni"

    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(
            sonuc,
            f,
            ensure_ascii=False,
            indent=2,
        )

    os.replace(
        gecici,
        KULLANICI_DOSYASI,
    )


    print("İlk kullanıcılar oluşturuldu.")
    print("Şifreler hash olarak kaydedildi.")

    return sonuc


_kullanici_mtime: float | None = None
_kullanici_onbellek: dict[str, dict] | None = None


def _kullanicilari_al() -> dict[str, dict]:
    """kullanicilar.json değişmişse restart gerektirmeden yeniden okur."""
    global _kullanici_mtime, _kullanici_onbellek

    try:
        mtime = os.path.getmtime(KULLANICI_DOSYASI)
    except OSError:
        if _kullanici_onbellek is None:
            _kullanici_onbellek = _kullanicilari_yukle()
        return _kullanici_onbellek

    if _kullanici_onbellek is None or _kullanici_mtime != mtime:
        try:
            with open(KULLANICI_DOSYASI, encoding="utf-8") as f:
                veri = json.load(f)
            if not isinstance(veri, dict) or not veri:
                raise RuntimeError("data/kullanicilar.json geçersiz.")
            _kullanici_onbellek = veri
            _kullanici_mtime = mtime
        except (json.JSONDecodeError, OSError, RuntimeError) as exc:
            if _kullanici_onbellek is not None:
                app.logger.error(
                    "kullanicilar.json okunamadı, son iyi önbellek kullanılıyor: %s",
                    exc,
                )
                return _kullanici_onbellek
            raise

    return _kullanici_onbellek


_kullanicilari_yukle()
_gunluk_dosya_logu_kur()




def _baslatma_hatasi_bildir(exc: Exception) -> None:
    """Açılış başarısızsa Python traceback yerine anlaşılır talimat üretir."""
    yedekler = []
    try:
        yedekler = depo.yedekleri_getir(5)
    except Exception:
        pass

    satirlar = [
        "=" * 70,
        "  PDGM İŞ TAKİP SİSTEMİ BAŞLATILAMADI",
        "=" * 70,
        "",
        f"  Hata: {exc}",
        "",
        "  Bu genellikle data/kartlar.xlsx dosyasının elle düzenlenmesi",
        "  sırasında oluşan bir veri hatasıdır.",
        "",
        "  YAPILACAKLAR:",
        "  1) data/kartlar.xlsx dosyasını Excel'de KAPATIN.",
        "  2) Yukarıdaki hata mesajında geçen kart ID / sütunu düzeltin,",
        "     VEYA aşağıdaki yedeklerden birini kartlar.xlsx üzerine kopyalayın.",
        "  3) Sunucuyu tekrar başlatın.",
        "",
    ]

    if yedekler:
        satirlar.append("  KULLANILABİLİR YEDEKLER (data/yedekler/ altında):")
        for y in yedekler:
            satirlar.append(f"    - {y['ad']}   ({y['zaman']}, {y['tip']})")
    else:
        satirlar.append("  UYARI: Kullanılabilir yedek bulunamadı.")

    satirlar += ["", "=" * 70, ""]
    metin = "\n".join(satirlar)

    print(metin)
    try:
        with open(
            os.path.join(VERI_KLASORU, "BASLATMA_HATASI.txt"),
            "w",
            encoding="utf-8",
        ) as f:
            f.write(metin)
    except OSError:
        pass


try:
    depo.kur()
except depo.VeriDogrulamaHatasi as _hata:
    app.logger.critical("Açılışta kart dosyası doğrulanamadı: %s", _hata)
    _baslatma_hatasi_bildir(_hata)
    raise SystemExit(1)


# ---------------------------------------------------------------------------
# Yetki, CSRF ve ortak template verisi
# ---------------------------------------------------------------------------

@app.before_request
def _oturum_kullanici_kontrol():
    kullanici = session.get("kullanici")
    if not kullanici:
        return None

    kayit = _kullanicilari_al().get(kullanici)
    if not kayit or not kayit.get("aktif", True):
        session.clear()
        if request.path.startswith("/api/"):
            return jsonify(hata="Oturum sonlandırıldı. Tekrar giriş yapın."), 401
        flash("Hesabınız pasif veya bulunamadı. Tekrar giriş yapın.", "hata")
        return redirect(url_for("giris"))

    session["rol"] = kayit.get("rol") or session.get("rol")
    session["ad"] = kayit.get("ad") or session.get("ad")
    session["operator_tipi"] = kayit.get("operator_tipi")
    return None


def yetki(*roller):
    def sarmalayici(fn):
        @wraps(fn)
        def ic(*args, **kwargs):
            if "kullanici" not in session:
                if request.path.startswith("/api/"):
                    return jsonify(hata="Oturum sona erdi. İşlem kaydedilmedi; tekrar giriş yapın."), 401
                return redirect(url_for("giris", devam=request.path))
            if roller and session.get("rol") not in roller:
                return render_template("yetkisiz.html"), 403
            return fn(*args, **kwargs)

        return ic

    return sarmalayici


def _csrf_token_uret() -> str:
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def csrf_koru(fn):
    @wraps(fn)
    def ic(*args, **kwargs):
        beklenen = session.get("csrf_token")
        gelen = request.headers.get("X-CSRF-Token") or request.form.get("_csrf_token")

        if beklenen and gelen and secrets.compare_digest(beklenen, gelen):
            return fn(*args, **kwargs)

        if request.path.startswith("/api/"):
            return jsonify(hata="Geçersiz veya eksik CSRF token."), 403
        flash("Oturum doğrulaması başarısız. Sayfayı yenileyip tekrar deneyin.", "hata")
        return redirect(request.referrer or url_for("ana"))

    return ic


def _guvenli_devam_hedefi(hedef: str | None) -> bool:
    if not hedef or "\\" in hedef:
        return False

    parca = urlsplit(hedef)

    return (
        not parca.scheme
        and not parca.netloc
        and hedef.startswith("/")
        and not hedef.startswith("//")
    )


@app.context_processor
def genel_degiskenler():
    return {
        "oturum_ad": session.get("ad"),
        "oturum_rol": session.get("rol"),
        "oturum_kullanici": session.get("kullanici"),
        "oturum_operator_tipi": session.get("operator_tipi"),
        "bugun": date.today().strftime("%d.%m.%Y"),
        "csrf_token": _csrf_token_uret() if "kullanici" in session else "",
    }


@app.template_filter("gun")
def gun_filtresi(deger):
    if not deger:
        return "—"
    parcalar = str(deger)[:10].split("-")
    return f"{parcalar[2]}.{parcalar[1]}.{parcalar[0]}" if len(parcalar) == 3 else str(deger)


@app.template_filter("sayi")
def sayi_filtresi(deger, basamak=1):
    """Ondalık sayıyı Türkçe biçimde yazar: 3.8 → 3,8; 4.0 → 4."""
    if deger is None or deger == "":
        return "—"
    try:
        sayi = round(float(deger), basamak)
    except (TypeError, ValueError):
        return str(deger)
    if sayi == int(sayi):
        return f"{int(sayi):,}".replace(",", ".")
    return f"{sayi:,.{basamak}f}".replace(",", " ").replace(".", ",").replace(" ", ".")


@app.template_filter("yukleme_adi")
def yukleme_adi_filtresi(deger):
    """Yükleme klasöründeki kayıt adından zaman/uuid önekini atar."""
    return re.sub(r"^\d{8}_\d{6}_[0-9a-f]{8}_", "", str(deger or "")) or "—"


@app.template_filter("onizleme_deger")
def onizleme_deger_filtresi(deger, alan=None):
    """Import önizlemesindeki eski/yeni değerleri kullanıcı diliyle gösterir."""
    if alan in ("plan_baslama", "plan_teslim", "gerceklesen_teslim"):
        return gun_filtresi(deger)
    if alan == "malzeme_bekliyor":
        return "Evet" if deger == 1 else "Hayır"
    if alan == "source_active":
        return "Excel'de var" if deger == 1 else "Excel'de yok"
    if alan == "durum" and not deger:
        return "Durum yok"
    return "—" if deger in (None, "") else str(deger)


_SORUN_DESENI = re.compile(
    r"^'(?P<sayfa>[^']+)'(?: sayfası)?,?\s*(?:Excel )?(?:satır (?P<satir>\d+)(?P<gizli> \(gizli\))?)?\s*:\s*(?P<mesaj>.+)$",
    re.S,
)


def _sorunlari_ayir(sorunlar):
    """"'MAKİNE' satır 272: ..." -> {"konum": "MAKİNE · satır 272", "mesaj": "..."}"""
    sonuc = []
    for sorun in sorunlar:
        eslesme = _SORUN_DESENI.match(str(sorun))
        if not eslesme:
            sonuc.append({"konum": "", "mesaj": str(sorun)})
            continue
        konum = eslesme["sayfa"]
        if eslesme["satir"]:
            konum += f" · satır {eslesme['satir']}" + (" (gizli)" if eslesme["gizli"] else "")
        sonuc.append({"konum": konum, "mesaj": eslesme["mesaj"].strip()})
    return sonuc


@app.errorhandler(PermissionError)
def _excel_kilitli(_hata):
    mesaj = (
        "Bir kayıt Excel dosyası şu anda Excel'de açık olabilir. "
        "Kartlar, işlem logu ve yükleme geçmişi dosyalarını kapatıp "
        "tekrar deneyin."
    )

    if request.path.startswith("/api/"):
        return jsonify(hata=mesaj), 423

    flash(mesaj, "hata")
    return redirect(url_for("ana"))


# ---------------------------------------------------------------------------
# Giriş / çıkış
# ---------------------------------------------------------------------------
_GIRIS_LIMIT = 8
_GIRIS_PENCERE_SN = 5 * 60

_giris_kilit = threading.Lock()
_giris_basarisiz: dict[str, deque[float]] = {}


def _giris_ip() -> str:
    return request.remote_addr or "bilinmeyen"


def _giris_engelli_mi() -> bool:
    ip = _giris_ip()
    an = time.monotonic()

    with _giris_kilit:
        denemeler = _giris_basarisiz.get(ip)

        if not denemeler:
            return False

        while (
            denemeler
            and an - denemeler[0] >= _GIRIS_PENCERE_SN
        ):
            denemeler.popleft()

        if not denemeler:
            _giris_basarisiz.pop(ip, None)
            return False

        return len(denemeler) >= _GIRIS_LIMIT


def _giris_basarisiz_kaydet():
    ip = _giris_ip()
    an = time.monotonic()

    with _giris_kilit:
        denemeler = _giris_basarisiz.setdefault(
            ip,
            deque(),
        )

        while (
            denemeler
            and an - denemeler[0] >= _GIRIS_PENCERE_SN
        ):
            denemeler.popleft()

        denemeler.append(an)


def _giris_limit_temizle():
    with _giris_kilit:
        _giris_basarisiz.pop(
            _giris_ip(),
            None,
        )

@app.route("/giris", methods=["GET", "POST"])
def giris():
    if request.method == "GET":
        return render_template("giris.html")

    if _giris_engelli_mi():
        return render_template(
            "giris.html",
            hata=(
                "Çok fazla başarısız giriş denemesi yapıldı. "
                "Birkaç dakika sonra tekrar deneyin."
            ),
        ), 429

    kullanici = (
        request.form.get("kullanici") or ""
    ).strip()
    sifre = request.form.get("sifre") or ""

    # Aşırı büyük input'u password hash fonksiyonuna ve Excel loguna sokma.
    if len(kullanici) > 128 or len(sifre) > 512:
        _giris_basarisiz_kaydet()
        return render_template(
            "giris.html",
            hata="Kullanıcı adı veya şifre hatalı.",
        ), 401

    kayit = _kullanicilari_al().get(kullanici)

    if (
        kayit
        and kayit.get("aktif", True)
        and check_password_hash(
            kayit.get("sifre_hash", ""),
            sifre,
        )
    ):
        _giris_limit_temizle()

        session.clear()
        session.permanent = True
        session.update(
            kullanici=kullanici,
            rol=kayit["rol"],
            ad=kayit["ad"],
            operator_tipi = kayit.get("operator_tipi")
        )
        _csrf_token_uret()

        try:
            depo.log_ekle(
                kullanici,
                kayit["rol"],
                "GİRİŞ YAPILDI",
            )
        except Exception:
            # Audit Excel geçici olarak yazılamasa bile başarılı auth
            # bozulmasın. Fallback olarak uygulama loguna yaz.
            app.logger.exception(
                "Giriş audit kaydı Excel'e yazılamadı: %s",
                kullanici,
            )

        hedef = request.args.get("devam")
        return redirect(
            hedef
            if _guvenli_devam_hedefi(hedef)
            else url_for("ana")
        )

    _giris_basarisiz_kaydet()

    app.logger.warning(
        "HATALI GİRİŞ · kullanici=%s · ip=%s",
        kullanici or "-",
        _giris_ip(),
    )

    return render_template(
        "giris.html",
        hata="Kullanıcı adı veya şifre hatalı.",
    ), 401

def _varsayilan_gozlemci_bul():
    """Aktif, 'gozlemci' rolündeki ilk kullanıcıyı döner (kullanıcı adı, kayıt)."""
    for kullanici, kayit in _kullanicilari_al().items():
        if kayit.get("rol") == "gozlemci" and kayit.get("aktif", True):
            return kullanici, kayit
    return None, None


@app.route("/giris/gozlemci", methods=["POST"])
def giris_gozlemci():
    """Gözlemci rolü için şifre istemeden tek tıkla giriş.

    Şifre kontrolü kasıtlı olarak yok: gözlemci salt-okunur bir rol, üretim
    üzerinde hiçbir aksiyon alamıyor (bkz. yetki() decorator'ları). Bu yüzden
    giriş sürtünmesini azaltmak güvenlik riski oluşturmuyor.
    """
    if _giris_engelli_mi():
        return render_template(
            "giris.html",
            hata=(
                "Çok fazla başarısız giriş denemesi yapıldı. "
                "Birkaç dakika sonra tekrar deneyin."
            ),
        ), 429

    kullanici, kayit = _varsayilan_gozlemci_bul()

    if not kullanici:
        return render_template(
            "giris.html",
            hata="Aktif bir gözlemci hesabı bulunamadı. Yöneticinizle iletişime geçin.",
        ), 401

    session.clear()
    session.permanent = True
    session.update(
        kullanici=kullanici,
        rol=kayit["rol"],
        ad=kayit["ad"],
        operator_tipi=kayit.get("operator_tipi"),
    )
    _csrf_token_uret()

    try:
        depo.log_ekle(
            kullanici,
            kayit["rol"],
            "GİRİŞ YAPILDI (ŞİFRESİZ GÖZLEMCİ GİRİŞİ)",
        )
    except Exception:
        app.logger.exception(
            "Gözlemci girişi audit kaydı Excel'e yazılamadı: %s",
            kullanici,
        )

    return redirect(url_for("ana"))

@app.route("/cikis", methods=["POST"])
@yetki("admin", "operator", "gozlemci")
@csrf_koru
def cikis():
    kullanici = session.get("kullanici", "-")
    rol = session.get("rol", "")

    try:
        depo.log_ekle(
            kullanici,
            rol,
            "ÇIKIŞ YAPILDI",
        )
    except Exception:
        app.logger.exception(
            "Çıkış audit kaydı Excel'e yazılamadı: %s",
            kullanici,
        )
    finally:
        session.clear()

    return redirect(url_for("giris"))


@app.route("/")
def ana():
    hedefler = {
        "admin": "yonetim",
        "operator": "operator",
        "gozlemci": "panel",
    }
    endpoint = hedefler.get(session.get("rol"))
    return redirect(url_for(endpoint or "giris"))


# ---------------------------------------------------------------------------
# Ekran verileri
# ---------------------------------------------------------------------------

def _pano_verisi():
    kartlar = depo.kartlari_getir()

    def farkli_stok(durum, dizgi_kod=None):
        """Durum (ve isteğe bağlı dizgi tipi kodu) için benzersiz stok_no sayısı."""
        stoklar = set()
        for k in kartlar:
            if k.get("durum") != durum:
                continue
            if dizgi_kod is not None and k.get("dizgi_kod") != dizgi_kod:
                continue
            stok = str(k.get("stok_no") or "").strip()
            if stok:
                stoklar.add(stok)
        return len(stoklar)

    sayac = {
        "plana_alindi": farkli_stok(depo.PLANA_ALINDI),
        "dizgide": farkli_stok(depo.DIZGIDE),
        "teslim": farkli_stok(depo.TESLIM_EDILDI),
        "gecikme": sum(
            k["renk"] == "kotu" and k["durum"] != depo.TESLIM_EDILDI
            for k in kartlar
        ),
        "plana_alindi_makine": farkli_stok(depo.PLANA_ALINDI, "MAKINE"),
        "plana_alindi_elle": farkli_stok(depo.PLANA_ALINDI, "ELLE"),
        "plana_alindi_eum": farkli_stok(depo.PLANA_ALINDI, "EUM"),
        "dizgide_makine": farkli_stok(depo.DIZGIDE, "MAKINE"),
        "dizgide_elle": farkli_stok(depo.DIZGIDE, "ELLE"),
        "dizgide_eum": farkli_stok(depo.DIZGIDE, "EUM"),
        "teslim_makine": farkli_stok(depo.TESLIM_EDILDI, "MAKINE"),
        "teslim_elle": farkli_stok(depo.TESLIM_EDILDI, "ELLE"),
        "teslim_eum": farkli_stok(depo.TESLIM_EDILDI, "EUM"),
    }

    # Ana sayı iş emri (kart) sayısıdır; farklı stok sayısı yanında alt bilgi olarak gösterilir.
    for anahtar, durum in (("plana_alindi", depo.PLANA_ALINDI), ("dizgide", depo.DIZGIDE),
                           ("teslim", depo.TESLIM_EDILDI)):
        sayac[f"{anahtar}_kart"] = sum(k.get("durum") == durum for k in kartlar)
        for kod in ("MAKINE", "ELLE", "EUM"):
            sayac[f"{anahtar}_kart_{kod.lower()}"] = sum(
                k.get("durum") == durum and k.get("dizgi_kod") == kod for k in kartlar
            )
    return {
        "kartlar": kartlar,
        "sayac": sayac,
        "guncelleme": datetime.now().strftime("%H:%M:%S"),
    }

@app.route("/panel")
@yetki("admin", "operator", "gozlemci")
def panel():
    # Sürüm veriden ÖNCE alınır: arada değişiklik olursa sayfa eski sürümle işaretlenir
    # ve yoklama değişikliği yakalar (tersi değişikliği kaçırırdı).
    surum = depo.depo_surumu()
    veri = _pano_verisi()
    kartlar = veri["kartlar"]

    teslim_edilen = sorted(
        [k for k in kartlar if k["durum"] == depo.TESLIM_EDILDI],
        key=lambda k: k.get("teslim_zamani") or k.get("gerceklesen_teslim") or "",
        reverse=True,
    )

    return render_template(
        "panel.html",
        sayac=veri["sayac"],
        guncelleme=veri["guncelleme"],
        dizgide=[k for k in kartlar if k["durum"] == depo.DIZGIDE],
        plana_alindi=[k for k in kartlar if k["durum"] == depo.PLANA_ALINDI],
        teslim_edilen=teslim_edilen,
        donem_ozet=_donem_ozeti(kartlar),
        veri_surumu=surum,
    )


@app.route("/monitor")
@yetki("admin", "operator", "gozlemci")
def monitor():
    surum = depo.depo_surumu()
    veri = _pano_verisi()
    kartlar = [k for k in veri["kartlar"] if k.get("dizgi_kod") == "MAKINE"]
    for kart in kartlar:
        # Monitörde gün sayısı yerine kısa durum; teslim tarihi kartta zaten yazıyor.
        if kart.get("rozet", "").startswith(f"{depo.SURESI_ICINDE} ("):
            kart["rozet"] = depo.SURESI_ICINDE

    return render_template(
        "monitor.html",
        guncelleme=veri["guncelleme"],
        dizgide=[k for k in kartlar if k["durum"] == depo.DIZGIDE],
        plana_alindi=[k for k in kartlar if k["durum"] == depo.PLANA_ALINDI],
        veri_surumu=surum,
    )


# Operatör ekranı teslim edilmiş kartların yalnız son günlerini çizer; eskiler Pano'da.
OPERATOR_TESLIM_GUN = 30


@app.route("/operator")
@yetki("admin", "operator")
def operator():
    surum = depo.depo_surumu()
    veri = _pano_verisi()
    sinir = (date.today() - timedelta(days=OPERATOR_TESLIM_GUN)).isoformat()
    kartlar, eski_teslim = [], 0
    for kart in veri["kartlar"]:
        teslim = _teslim_tarihi(kart) if kart["durum"] == depo.TESLIM_EDILDI else None
        if teslim and teslim < sinir:
            eski_teslim += 1
            continue
        kartlar.append(kart)
    return render_template(
        "operator.html",
        kartlar=kartlar,
        sayac=veri["sayac"],
        eski_teslim=eski_teslim,
        teslim_gun=OPERATOR_TESLIM_GUN,
        veri_surumu=surum,
    )


@app.route("/yonetim")
@yetki("admin")
def yonetim():
    surum = depo.depo_surumu()
    kartlar = depo.kartlari_yonetim_getir()
    kaynakta_olmayan = [
        k
        for k in kartlar
        if k.get("kaynakta_yok") and k.get("durum") != depo.TESLIM_EDILDI
    ]

    return render_template(
        "yonetim.html",
        kartlar=kartlar,
        kaynakta_olmayan=kaynakta_olmayan,
        durumu_eksik=depo.durumu_eksik_kartlari_getir(),
        gizlenen_kartlar=depo.gizlenen_kartlari_getir(),
        yedekler=depo.yedekleri_getir(12),
        yuklemeler=depo.yuklemeleri_getir(8),
        loglar=depo.loglari_getir(25),
        veri_surumu=surum,
    )


@app.route("/api/surum")
@yetki("admin", "operator", "gozlemci")
def api_surum():
    """Açık ekranların "veri değişti mi / sunucu ayakta mı" yoklaması için hafif uç."""
    return jsonify(surum=depo.depo_surumu(), zaman=datetime.now().strftime("%H:%M:%S"))





@app.route("/api/veriler")
@yetki("admin", "operator", "gozlemci")
def api_veriler():
    return jsonify(_pano_verisi())


# ---------------------------------------------------------------------------
# Operatör API'leri
# ---------------------------------------------------------------------------

def _api_kart_islemi(fn):
    try:
        return fn()
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    except depo.MalzemeBekliyorHatasi as hata:
        return jsonify(hata=str(hata), malzeme_bekliyor_onay_gerekli=True), 409
    except depo.IsKuralHatasi as hata:
        return jsonify(hata=str(hata)), 409
    except (depo.VeriDogrulamaHatasi, ValueError) as hata:
        return jsonify(hata=str(hata)), 400
    except TypeError as hata:
        # ValueError'dan bilinçli olarak ayrıldı: TypeError genelde kullanıcı
        # girdisinden değil, çağıran/çağrılan taraf imza uyuşmazlığından
        # (programlama hatası) kaynaklanır. 400 arkasında sessizce gizlenmemeli.
        app.logger.exception("API katmanında beklenmeyen TypeError")
        return jsonify(
            hata="Sunucu tarafında beklenmeyen bir hata oluştu. Uygulama logu kontrol edilmeli."
        ), 500


@app.route("/api/basla", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_basla():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_baslat(
            kart_id=veri.get("kart_id"),
            adet=veri.get("adet"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            operator_tipi=session.get("operator_tipi"),
            aciklama=(veri.get("not") or "").strip(),
            malzeme_onayi=bool(veri.get("malzeme_onayi")),
            isim=(veri.get("isim") or session.get("ad") or session["kullanici"]),
        )
        return jsonify(tamam=True, mesaj="Kart DİZGİDE durumuna alındı.", kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/bitir", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_bitir():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart, uretim_bitti, mesaj = depo.kart_bitir(
            kart_id=veri.get("kart_id"),
            adet=veri.get("adet"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            operator_tipi=session.get("operator_tipi"),
            aciklama=(veri.get("not") or "").strip(),
            isim=(veri.get("isim") or session.get("ad") or session["kullanici"]),
        )
        return jsonify(
            tamam=True,
            uretim_bitti=uretim_bitti,
            mesaj=mesaj,
            kart=kart,
        )

    return _api_kart_islemi(islem)





@app.route("/api/teslim-et", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_teslim_et():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_teslim_et(
            kart_id=veri.get("kart_id"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            operator_tipi=session.get("operator_tipi"),
            aciklama=(veri.get("not") or "").strip(),
            isim=(veri.get("isim") or session.get("ad") or session["kullanici"]),
        )
        return jsonify(tamam=True, mesaj="Kart TESLİM EDİLDİ olarak kaydedildi.", kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/not", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_not():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.kart_not_guncelle(
            kart_id=veri.get("kart_id"),
            aciklama=(veri.get("not") or "").strip(),
            kullanici=session["kullanici"],
            rol=session["rol"],
            isim=(veri.get("isim") or "").strip()
            or session.get("ad")
            or session["kullanici"],
        )
        return jsonify(tamam=True, kart=kart)

    return _api_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Admin API'leri
# ---------------------------------------------------------------------------

@app.route("/api/admin/kart-ekle", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_kart_ekle():
    veri = request.get_json(silent=True) or {}

    def islem():
        kart = depo.admin_kart_ekle(
            sira=veri.get("sira"),
            talep_no=veri.get("talep_no"),
            talep_sahibi=veri.get("talep_sahibi"),
            stok_no=veri.get("stok_no"),
            toplam_adet=veri.get("toplam_adet"),
            plan_hafta=veri.get("plan_hafta"),
            plan_baslama=veri.get("plan_baslama"),
            plan_teslim=veri.get("plan_teslim"),
            gerceklesen_teslim=veri.get("gerceklesen_teslim"),
            pcb=veri.get("pcb"),
            aciklama=veri.get("not"),
            elle_dizgi=None if veri.get("dizgi_tipi") is not None else bool(veri.get("elle_dizgi")),
            dizgi_tipi=veri.get("dizgi_tipi"),
            dizgi_sorumlusu=veri.get("dizgi_sorumlusu"),
            kullanici=session["kullanici"],
        )
        return jsonify(tamam=True, kart=kart), 201

    return _api_kart_islemi(islem)


@app.route("/api/admin/duzenle", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_duzenle():
    veri = request.get_json(silent=True) or {}

    def islem():
        if not isinstance(veri.get("surum"), str) or not veri["surum"]:
            raise depo.IsKuralHatasi("Formun sürüm bilgisi eksik. Sayfayı yenileyip tekrar deneyin.")
        kart = depo.admin_kart_duzenle(
            beklenen_surum=veri["surum"],
            kart_id=veri.get("kart_id"),
            durum=veri.get("durum"),
            tamamlanan_adet=veri.get("tamamlanan_adet"),
            toplam_adet=veri.get("toplam_adet"),
            plan_hafta=veri.get("plan_hafta"),
            plan_baslama=veri.get("plan_baslama"),
            plan_teslim=veri.get("plan_teslim"),
            gerceklesen_teslim=veri.get("gerceklesen_teslim"),
            aciklama=veri.get("not"),
            elle_dizgi=veri.get("elle_dizgi"),
            dizgi_tipi=veri.get("dizgi_tipi"),
            dizgi_sorumlusu=veri.get("dizgi_sorumlusu"),
            malzeme_bekliyor=veri.get("malzeme_bekliyor"),
            kullanici=session["kullanici"],
        )
        return jsonify(tamam=True, kart=kart)

    return _api_kart_islemi(islem)


@app.route("/api/admin/kart-sil", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_kart_sil():
    veri = request.get_json(silent=True) or {}

    def islem():
        depo.admin_kart_gizle(veri.get("kart_id"), session["kullanici"])
        return jsonify(tamam=True)

    return _api_kart_islemi(islem)


@app.route("/yonetim/kart-geri-getir", methods=["POST"])
@yetki("admin")
@csrf_koru
def kart_geri_getir():
    try:
        kart = depo.admin_kart_geri_getir(
            kart_id=request.form.get("kart_id"),
            kullanici=session["kullanici"],
        )
    except (depo.KartBulunamadi, depo.IsKuralHatasi) as hata:
        flash(str(hata), "hata")
    else:
        flash(
            f"Kart geri getirildi · {kart.get('talep_no') or '—'} · {kart.get('stok_no') or '—'}",
            "basari",
        )
    return redirect(url_for("yonetim"))


@app.route("/yonetim/yedek-geri-yukle", methods=["POST"])
@yetki("admin")
@csrf_koru
def yedek_geri_yukle():
    yedek_adi = (request.form.get("yedek") or "").strip()

    try:
        sonuc = depo.yedekten_geri_yukle(yedek_adi, session["kullanici"])
    except (depo.IsKuralHatasi, depo.VeriDogrulamaHatasi) as hata:
        flash(f"Yedek geri yüklenemedi: {hata}", "hata")
    except OSError as hata:
        app.logger.exception("Yedek geri yükleme dosya hatası")
        flash(f"Yedek geri yüklenirken dosya hatası oluştu: {hata}", "hata")
    else:
        flash(
            f"Yedek geri yüklendi · {sonuc['kart']} kart. "
            f"Önceki durum '{os.path.basename(sonuc['koruma_yedegi'])}' içinde korundu.",
            "basari",
        )
        app.logger.warning("Admin yedek geri yükledi: %s", yedek_adi)

    return redirect(url_for("yonetim"))


# ---------------------------------------------------------------------------
# Excel yükleme / kayıt dosyaları / rapor
# ---------------------------------------------------------------------------

YUKLEME_SAKLA = 20


def _yuklenen_exceleri_buda():
    """En yeni N yüklenmiş Excel'i tutar, gerisini siler."""
    try:
        dosyalar = []
        for ad in os.listdir(YUKLEME_KLASORU):
            yol = os.path.join(YUKLEME_KLASORU, ad)
            if os.path.isfile(yol):
                dosyalar.append((os.path.getmtime(yol), yol))
        dosyalar.sort(reverse=True)
        for _, yol in dosyalar[YUKLEME_SAKLA:]:
            try:
                os.remove(yol)
            except OSError:
                pass
    except OSError:
        pass


def _kisalt(metin, sinir=1500):
    """Flash mesajları session çerezinde taşınır; çerez 4 KB sınırını aşmamalı."""
    metin = str(metin)
    return metin if len(metin) <= sinir else metin[: sinir - 1] + "…"


def _import_mesaji(sonuc):
    mesaj = (
        f"Excel aktarıldı · {sonuc['satir']} satır · {sonuc['yeni']} yeni · "
        f"{sonuc['guncellenen']} güncellendi · {sonuc.get('degismeyen', 0)} değişmedi · "
        f"{sonuc.get('workflow_korundu', 0)} workflow korundu."
    )
    if sonuc.get("pasife_alinan"):
        mesaj += f" · {sonuc['pasife_alinan']} kart kaynak Excel'de artık yok."
    if sonuc.get("ayrilan"):
        mesaj += (f" · {sonuc['ayrilan']} kartın NO'su Excel'de başka talebe verilmiş: eski kart geçmişiyle "
                  "ayrıldı, yeni talep için kart açıldı.")
    if sonuc.get("elle_dizgi_satir"):
        mesaj += f" · {sonuc['elle_dizgi_satir']} kart ELDE DİZGİ sayfasından."
    if sonuc.get("eum_dizgi_satir"):
        mesaj += f" · {sonuc['eum_dizgi_satir']} kart EÜM sayfasından."
    if sonuc.get("gerileme"):
        mesaj += (f" · {sonuc['gerileme']} kartta Excel uygulamanın gerisindeydi; "
                  f"{sonuc.get('gerileme_korunan', 0)} kartta uygulamadaki durum korundu.")
    if sonuc.get("durumsuz"):
        korunan = sonuc.get("durumsuz_korunan", 0)
        mesaj += (f" · {sonuc['durumsuz']} kartta Excel'de DURUM boştu; {korunan} kartta uygulamadaki durum "
                  f"korundu, {sonuc['durumsuz'] - korunan} kart durumsuz bırakıldı (Durumu Eksik Kartlar).")
    if sonuc.get("sifirlanan"):
        mesaj += f" · {sonuc['sifirlanan']} kartın tamamlanan adedi seçiminizle sıfırlandı."
    if sonuc.get("notu_temizlenen"):
        mesaj += (f" · {sonuc['notu_temizlenen']} kartın notları seçiminizle temizlendi; "
                  "silinen notlar işlem logunda.")
    if sonuc.get("durum_iyilesen"):
        mesaj += (
            f" · {sonuc['durum_iyilesen']} kartın durumu Excel'den güncellendi."
        )
    if sonuc.get("uyari"):
        mesaj += (
            f" · {sonuc['uyari']} veri uyarısı var (boş/geçersiz DURUM, eksik teslim tarihi vb.); "
            "bu kartlar Yönetim ekranında kontrol edilebilir."
        )
    rapor = sonuc.get("kaynak_raporu") or {}
    if rapor.get("atlanan_satir"):
        yerler = "; ".join(
            f"{s['sayfa']} satır {s['atlanan_araliklar']}"
            for s in rapor.get("sayfalar", []) if s.get("atlanan")
        )
        mesaj += (
            f" · {rapor['atlanan_satir']} satır Talep NO/Kart Stok No içermediği için kart sayılmadı "
            f"({_kisalt(yerler, 400)})."
        )
    if sonuc.get("yedek"):
        mesaj += f" · Yedek: {os.path.basename(sonuc['yedek'])}"
    return mesaj


@app.route("/yonetim/yukle", methods=["POST"])
@yetki("admin")
@csrf_koru
def yukle():
    dosya = request.files.get("dosya")
    if not dosya or not dosya.filename:
        flash("Dosya seçilmedi.", "hata")
        return redirect(url_for("yonetim"))

    guvenli_ad = secure_filename(dosya.filename)
    if not guvenli_ad.lower().endswith((".xlsx", ".xlsm")):
        flash("Sadece .xlsx veya .xlsm dosyası yükleyin.", "hata")
        return redirect(url_for("yonetim"))

    os.makedirs(YUKLEME_KLASORU, exist_ok=True)
    kayit_adi = f"{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}_{guvenli_ad}"
    yol = os.path.join(YUKLEME_KLASORU, kayit_adi)
    dosya.save(yol)

    try:
        if request.form.get("onizleme") == "1":
            session.pop("import_onizleme", None)
            sonuc = ex.excelden_aktar(yol, session["kullanici"], onizleme=True)
            with open(yol, "rb") as kaynak:
                dosya_hash = hashlib.file_digest(kaynak, "sha256").hexdigest()
            session["import_onizleme"] = {"dosya": kayit_adi, "hash": dosya_hash,
                "surum": sonuc["surum"], "kaynak_surum": sonuc["kaynak_surum"],
                "sifirlanabilir": sonuc.get("sifirlanabilir", []),
                "not_temizlenebilir": sonuc.get("not_temizlenebilir", [])}
            return render_template("import_onizleme.html", sonuc=sonuc, dosya=guvenli_ad)
        sonuc = ex.excelden_aktar(yol, session["kullanici"])
        mesaj = _import_mesaji(sonuc)
        flash(mesaj, "basari")
        app.logger.info("Excel import OK: %s", mesaj)
        _yuklenen_exceleri_buda()
    except (ex.ExcelAktarimHatasi, depo.IsKuralHatasi, depo.VeriDogrulamaHatasi) as hata:
        app.logger.warning("Excel import reddedildi: %s", hata)
        if request.form.get("onizleme") == "1":
            # Sorunlar kesilmiş tek bir bildirim yerine madde madde gösterilir.
            sorunlar = getattr(hata, "sorunlar", None) or [str(hata)]
            return render_template("import_onizleme.html", hata=str(hata), hata_turu="excel",
                                   sorunlar=_sorunlari_ayir(sorunlar), dosya=guvenli_ad), 422
        flash(_kisalt(f"Excel içeriği kabul edilmedi: {hata}"), "hata")
    except Exception:  # noqa: BLE001
        app.logger.exception("Excel import sırasında beklenmeyen hata")
        flash("Excel okunurken beklenmeyen bir hata oluştu. Uygulama logunu kontrol edin.", "hata")

    return redirect(url_for("yonetim"))


@app.route("/yonetim/yukle-onay", methods=["POST"])
@yetki("admin")
@csrf_koru
def yukle_onay():
    bekleyen = session.pop("import_onizleme", None)
    try:
        if not bekleyen or not bekleyen.get("kaynak_surum"):
            raise depo.IsKuralHatasi("Önizleme bulunamadı veya daha önce kullanıldı. Dosyayı yeniden seçin.")
        yol = os.path.join(YUKLEME_KLASORU, os.path.basename(bekleyen["dosya"]))
        with open(yol, "rb") as kaynak:
            if hashlib.file_digest(kaynak, "sha256").hexdigest() != bekleyen["hash"]:
                raise depo.IsKuralHatasi("Dosya önizlemeden sonra değişti. Dosyayı yeniden seçin.")
        # Önizlemede "Tamamlanan adedi sıfırla" işaretlenen kartlar; yalnız o önizlemede
        # sunulan kartlar kabul edilir (depo uygulama anında uygunluğu yeniden denetler).
        try:
            sifirla = sorted({int(deger) for deger in request.form.getlist("sifirla")})
            not_temizle = sorted({int(deger) for deger in request.form.getlist("not_temizle")})
        except ValueError:
            raise depo.IsKuralHatasi("Geçersiz kart seçimi. Dosyayı yeniden seçin.") from None
        disarida = set(sifirla) - set(bekleyen.get("sifirlanabilir") or [])
        if disarida:
            raise depo.IsKuralHatasi(
                "Tamamlanan adedi sıfırlanmak üzere önizlemede sunulmayan kart seçildi "
                f"(ID: {', '.join(map(str, sorted(disarida)))}). Dosyayı yeniden seçin."
            )
        disarida = set(not_temizle) - set(bekleyen.get("not_temizlenebilir") or [])
        if disarida:
            raise depo.IsKuralHatasi(
                "Notları temizlenmek üzere önizlemede sunulmayan kart seçildi "
                f"(ID: {', '.join(map(str, sorted(disarida)))}). Dosyayı yeniden seçin."
            )
        # Excel'in uygulamanın gerisinde kaldığı kartlar için seçilen durumlar. Hangi
        # kartların karar gerektirdiğini ve seçeneklerini depo yeniden hesaplar;
        # eksik, fazla veya aralık dışı seçim tüm aktarımı iptal eder.
        try:
            gerileme_secimleri = {int(ad.split("_", 1)[1]): deger for ad, deger in request.form.items()
                                  if ad.startswith("gerileme_")}
        except ValueError:
            raise depo.IsKuralHatasi("Geçersiz durum seçimi. Dosyayı yeniden seçin.") from None
        sonuc = ex.excelden_aktar(yol, session["kullanici"], beklenen_surum=bekleyen["surum"],
                                 beklenen_kaynak_surum=bekleyen["kaynak_surum"], tamamlanan_sifirla=sifirla,
                                 gerileme_secimleri=gerileme_secimleri, notlari_temizle=not_temizle)
    except (depo.DepoHatasi, ex.ExcelAktarimHatasi, OSError) as hata:
        sorunlar = getattr(hata, "sorunlar", None) or [str(hata)]
        return render_template("import_onizleme.html", hata=str(hata), hata_turu="onay",
                               sorunlar=_sorunlari_ayir(sorunlar),
                               # Kayıt adı "tarih_saat_kimlik_asıl-ad.xlsx" biçiminde.
                               dosya=os.path.basename((bekleyen or {}).get("dosya") or "").split("_", 3)[-1]), 409
    mesaj = _import_mesaji(sonuc)
    flash(mesaj, "basari")
    app.logger.info("Önizleme onaylandı; Excel import OK: %s", mesaj)
    _yuklenen_exceleri_buda()
    return redirect(url_for("yonetim"))


@app.route("/yonetim/yeniden-oku", methods=["POST"])
@yetki("admin")
@csrf_koru
def yeniden_oku():
    try:
        adet = depo.kartlari_diskten_yeniden_yukle()
    except depo.VeriDogrulamaHatasi as hata:
        flash(f"kartlar.xlsx yeniden okunamadı: {hata}", "hata")
        return redirect(url_for("yonetim"))

    depo.log_ekle(
        session["kullanici"],
        "admin",
        "KART DOSYASI YENİDEN OKUNDU",
        detay=f"{adet} kart",
    )
    flash(f"kartlar.xlsx doğrulandı ve yeniden okundu · {adet} kart.", "basari")
    return redirect(url_for("yonetim"))


@app.route("/yonetim/kayit-dosyasi/<hangi>")
@yetki("admin")
def kayit_dosyasi(hangi):
    dosyalar = {
        "kartlar": depo.KARTLAR_DOSYA,
        "log": depo.LOG_DOSYA,
        "yuklemeler": depo.YUKLEME_DOSYA,
    }
    yol = dosyalar.get(hangi)
    if not yol or not os.path.exists(yol):
        flash("Dosya henüz oluşmamış.", "hata")
        return redirect(url_for("yonetim"))

    app.logger.info(
        "KAYIT DOSYASI İNDİRİLDİ · kullanici=%s · hangi=%s",
        session.get("kullanici"),
        hangi,
    )
    return send_file(yol, as_attachment=True, download_name=os.path.basename(yol))


@app.route("/yonetim/rapor")
@yetki("admin")
def rapor_indir():
    ozet = ozet_hesapla()
    ozet_satirlari = [
        ["Rapor tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Toplam kart", ozet["genel"]["toplam"]],
        ["Plana alındı", ozet["genel"]["plana_alindi"]],
        ["Dizgide", ozet["genel"]["dizgide"]],
        ["Hazır", ozet["genel"]["hazir"]],
        ["Teslim edildi", ozet["genel"]["teslim"]],
        ["Süresi aşan açık kart", ozet["genel"]["gecikme"]],
        ["Bu hafta teslim edilen", ozet["donemler"]["Bu hafta"]["kart"]],
        ["Bu ay teslim edilen", ozet["donemler"]["Bu ay"]["kart"]],
        ["Bu yıl teslim edilen", ozet["donemler"]["Bu yıl"]["kart"]],
        ["Bu ay zamanında teslim (%)", ozet["donemler"]["Bu ay"]["zamaninda_yuzde"]],
        ["Bu ay ortalama teslim sapması (gün)", ozet["donemler"]["Bu ay"]["ort_sapma"]],
    ]

    wb = ex.calisma_kitabi_uret(
        depo.kartlari_getir(sadece_gorunen=False),
        depo.loglari_getir(),
        ozet_satirlari,
    )
    app.logger.info(
        "RAPOR İNDİRİLDİ · kullanici=%s",
        session.get("kullanici"),
    )
    return send_file(
        ex.kitap_baytlari(wb),
        as_attachment=True,
        download_name=ex.dosya_adi("PDGM_Rapor"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Özet hesapları
# ---------------------------------------------------------------------------

def _teslim_tarihi(kart) -> str | None:
    deger = kart.get("gerceklesen_teslim") or kart.get("teslim_zamani")
    return str(deger)[:10] if deger else None


def _donem_ozeti(kartlar, baslangic=None, bitis=None):
    alt = baslangic.strftime("%Y-%m-%d") if baslangic else None
    ust = bitis.strftime("%Y-%m-%d") if bitis else None
    secilen = []

    for kart in kartlar:
        if kart.get("durum") != depo.TESLIM_EDILDI:
            continue
        teslim = _teslim_tarihi(kart)
        if alt is None or (teslim and alt <= teslim <= ust):
            secilen.append(kart)

    sapmalar = [kart["sapma"] for kart in secilen if kart["sapma"] is not None]
    zamaninda = sum(sapma <= 0 for sapma in sapmalar)
    zamaninda_adet = sum(
        kart["tamamlanan_adet"]
        for kart in secilen
        if kart["sapma"] is not None and kart["sapma"] <= 0
    )
    gecikmeli_adet = sum(
        kart["tamamlanan_adet"]
        for kart in secilen
        if kart["sapma"] is not None and kart["sapma"] > 0
    )

    return {
        "kart": len(secilen),
        "adet": sum(kart["tamamlanan_adet"] for kart in secilen),
        "zamaninda": zamaninda,
        "gecikmeli": len(sapmalar) - zamaninda,
        "zamaninda_adet": zamaninda_adet,
        "gecikmeli_adet": gecikmeli_adet,
        "sapma_olculen": len(sapmalar),
        "zamaninda_yuzde": round(zamaninda / len(sapmalar) * 100) if sapmalar else 0,
        "ort_sapma": round(sum(sapmalar) / len(sapmalar), 1) if sapmalar else 0,
    }


def ozet_hesapla():
    # Kaynak Excel'den kaldırılmış kartlar pano/teslim API'sinde olduğu gibi
    # sayaçlara girmez; rapordaki kart listesi onları "Kaynakta Aktif=0" ile tutar.
    kartlar = [
        kart
        for kart in depo.kartlari_getir(sadece_gorunen=False)
        if not kart.get("admin_gizli") and kart.get("aktif", 1) == 1
        and not kart.get("kaynakta_yok")
    ]
    bugun_tarih = date.today()

    donemler = OrderedDict()
    donemler["Bu hafta"] = _donem_ozeti(
        kartlar,
        bugun_tarih - timedelta(days=bugun_tarih.weekday()),
        bugun_tarih,
    )
    donemler["Bu ay"] = _donem_ozeti(kartlar, bugun_tarih.replace(day=1), bugun_tarih)
    donemler["Bu yıl"] = _donem_ozeti(
        kartlar,
        bugun_tarih.replace(month=1, day=1),
        bugun_tarih,
    )

    haftalar = []
    for geri in range(7, -1, -1):
        bas = bugun_tarih - timedelta(days=bugun_tarih.weekday() + geri * 7)
        son = min(bas + timedelta(days=6), bugun_tarih)
        alt, ust = bas.strftime("%Y-%m-%d"), son.strftime("%Y-%m-%d")

        planlanan = sum(
            1
            for kart in kartlar
            if kart.get("plan_teslim") and alt <= kart["plan_teslim"] <= ust
        )
        teslim = sum(
            1
            for kart in kartlar
            if kart.get("durum") == depo.TESLIM_EDILDI
            and (teslim_tarihi := _teslim_tarihi(kart))
            and alt <= teslim_tarihi <= ust
        )
        haftalar.append(
            {
                "etiket": bas.strftime("%d.%m"),
                "hafta_no": bas.isocalendar()[1],
                "planlanan": planlanan,
                "teslim": teslim,
                "sapma": teslim - planlanan,
            }
        )

    en_yuksek = max(
        [hafta["planlanan"] for hafta in haftalar]
        + [hafta["teslim"] for hafta in haftalar]
        + [1]
    )

    genel = {
        "toplam": len(kartlar),
        "plana_alindi": sum(k["durum"] == depo.PLANA_ALINDI for k in kartlar),
        "dizgide": sum(k["durum"] == depo.DIZGIDE for k in kartlar),
        "hazir": sum(k["durum"] == depo.HAZIR for k in kartlar),
        "teslim": sum(k["durum"] == depo.TESLIM_EDILDI for k in kartlar),
        "durumu_eksik": sum(not k.get("durum") for k in kartlar),
        "gecikme": sum(
            k.get("gorunur")
            and k.get("renk") == "kotu"
            and k.get("durum") != depo.TESLIM_EDILDI
            for k in kartlar
        ),
    }
    geciken = [
        k
        for k in kartlar
        if k.get("gorunur")
        and k.get("renk") == "kotu"
        and k.get("durum") != depo.TESLIM_EDILDI
    ]

    return {
        "genel": genel,
        "donemler": donemler,
        "haftalar": haftalar,
        "en_yuksek": en_yuksek,
        "geciken_kartlar": geciken,
    }

def _tarih_araligi_coz(aralik, baslangic_param=None, bitis_param=None):
    """Pano dönem filtresi için (baslangic, bitis) date çiftini döndürür.
    'tumu' için (None, None) döner — çağıran taraf bunu 'filtre yok' olarak yorumlar.
    """
    bugun_tarih = date.today()

    if aralik == "hafta":
        return bugun_tarih - timedelta(days=bugun_tarih.weekday()), bugun_tarih
    if aralik == "ay":
        return bugun_tarih.replace(day=1), bugun_tarih
    if aralik == "yil":
        return bugun_tarih.replace(month=1, day=1), bugun_tarih
    if aralik == "ozel":
        try:
            alt = datetime.strptime(baslangic_param, "%Y-%m-%d").date()
            ust = datetime.strptime(bitis_param, "%Y-%m-%d").date()
        except (TypeError, ValueError) as exc:
            raise ValueError("Geçerli bir tarih aralığı seçin.") from exc
        if alt > ust:
            raise ValueError("Başlangıç tarihi bitiş tarihinden sonra olamaz.")
        return alt, ust
    return None, None


@app.route("/api/panel/teslimler")
@yetki("admin", "operator", "gozlemci")
def api_panel_teslimler():
    aralik = (request.args.get("aralik") or "tumu").strip().lower()
    dizgi_filtre = (request.args.get("dizgi") or "HEPSI").strip().upper()

    try:
        alt, ust = _tarih_araligi_coz(
            aralik,
            request.args.get("baslangic"),
            request.args.get("bitis"),
        )
    except ValueError as hata:
        return jsonify(hata=str(hata)), 400

    kartlar = depo.kartlari_getir()

    if dizgi_filtre == "MAKINE":
        kartlar = [k for k in kartlar if k.get("dizgi_kod") == "MAKINE"]
    elif dizgi_filtre == "ELLE":
        kartlar = [k for k in kartlar if k.get("dizgi_kod") == "ELLE"]
    elif dizgi_filtre == "EUM":
        kartlar = [k for k in kartlar if k.get("dizgi_kod") == "EUM"]

    if alt is None:
        donem_ozet = _donem_ozeti(kartlar)
        secilenler = [k for k in kartlar if k["durum"] == depo.TESLIM_EDILDI]
    else:
        donem_ozet = _donem_ozeti(kartlar, alt, ust)
        alt_iso, ust_iso = alt.strftime("%Y-%m-%d"), ust.strftime("%Y-%m-%d")
        secilenler = [
            k
            for k in kartlar
            if k["durum"] == depo.TESLIM_EDILDI
            and (teslim_tarihi := _teslim_tarihi(k))
            and alt_iso <= teslim_tarihi <= ust_iso
        ]

    secilenler.sort(
        key=lambda k: k.get("teslim_zamani") or k.get("gerceklesen_teslim") or "",
        reverse=True,
    )

    return jsonify(
        ozet=donem_ozet,
        teslim_edilen=[
            {
                "talep_no": k.get("talep_no"),
                "talep_sahibi": k.get("talep_sahibi"),
                "aciklama": k.get("aciklama"),
                "stok_no": k.get("stok_no"),
                "pcb": k.get("pcb"),
                "toplam_adet": k.get("toplam_adet"),
                "plan_baslama": gun_filtresi(k.get("plan_baslama")),
                "teslim": gun_filtresi(_teslim_tarihi(k)),
                "operator": k.get("operator"),
                "rozet": k.get("rozet"),
                "renk": k.get("renk"),
                "elle_dizgi_mi": k.get("elle_dizgi_mi", False),
                "eum_dizgi_mi": k.get("eum_dizgi_mi", False),
                "dizgi_kod": k.get("dizgi_kod", "MAKINE"),
                "dizgi_etiket": k.get("dizgi_etiket", "Makine"),
                "dizgi_sorumlusu": k.get("dizgi_sorumlusu"),
            }
            for k in secilenler
        ],
    )



# ---------------------------------------------------------------------------
# Çalıştırma
# ---------------------------------------------------------------------------

def calistir():
    depo.process_kilidi_al()
    try:
        depo.kur()
    except depo.VeriDogrulamaHatasi as hata:
        app.logger.critical("Açılışta kart dosyası doğrulanamadı: %s", hata)
        _baslatma_hatasi_bildir(hata)
        raise SystemExit(1)
    _gunluk_dosya_logu_kur()

    print("\n  PDGM İş Takip Sistemi çalışıyor")
    print(f"  Dinlenen adres : {DINLENEN_ADRES}:{SUNUCU_PORTU}")
    print(f"  Bu bilgisayarda : http://127.0.0.1:{SUNUCU_PORTU}")
    print(f"  Ağdaki diğer PC : http://<sunucunun-ip-adresi>:{SUNUCU_PORTU}")
    print(f"  Kayıtlar        : {VERI_KLASORU}")
    print("  Sunucu modeli   : tek process + çok thread")
    print(f"  Uygulama logu   : {LOG_DOSYASI}")
    print("  Durdurmak için  : Ctrl + C\n")
    app.logger.info("Sunucu başladı port=%s", SUNUCU_PORTU)

    try:
        from waitress import serve

        serve(app, host=DINLENEN_ADRES, port=SUNUCU_PORTU, threads=8)
    except ImportError:
        app.run(host=DINLENEN_ADRES, port=SUNUCU_PORTU, debug=False, threaded=True)


if __name__ == "__main__":
    calistir()
