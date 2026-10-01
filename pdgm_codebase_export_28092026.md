# PDGM Codebase Export

- Üretim zamanı: 2026-09-28T00:02:31+03:00
- Kök dizin: `C:\Users\DOĞUKAN ÇAYAĞZI\Downloads\gokberk_flask`
- Dahil edilen dosya sayısı: 29
- Hariç tutulanlar: testler, yardımcı çalışma araçları, Markdown dokümanları, çalışma verileri, yüklemeler, yedekler, çıktı raporları, inceleme kopyaları, sanal ortamlar, önbellekler ve export dosyaları.

## Dosya listesi

- `.gitignore`
- `app.py`
- `depo.py`
- `excel_araclari.py`
- `kullanici_yonet.py`
- `old_codebase.txt`
- `requirements.txt`
- `run_pdgm.bat`
- `static/js/monitor.js`
- `static/js/onizleme.js`
- `static/js/operator.js`
- `static/js/ortak.js`
- `static/js/panel.js`
- `static/js/yonetim.js`
- `static/stil.css`
- `templates/_makrolar.html`
- `templates/base.html`
- `templates/giris.html`
- `templates/import_onizleme.html`
- `templates/monitor.html`
- `templates/operator.html`
- `templates/panel.html`
- `templates/yetkisiz.html`
- `templates/yonetim.html`
- `test_verileri/mac_import_yardimci.py`
- `test_verileri/TEST_SENARYOLARI.txt`
- `tools/export_codebase.py`
- `tools/test_paketi_uret.py`
- `yedek_disari_kopyala.bat`

## Dosya içerikleri

## `.gitignore`

```text
.env
.venv/
venv/
__pycache__/
*.pyc
*.pyo
.DS_Store

data/
!data/.gitkeep

*.xlsx.bozuk_*
data/BASLATMA_HATASI.txt
```

## `app.py`

```python
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

ISLEM_YAPAN_SINIRI = 80


def _islem_yapan_adi(veri):
    """Operatör hesapları paylaşımlı: işlemi yapan kişinin adı karta ve loga yazılır.

    Arayüz adı zorunlu tutar; ad gönderilmeyen API çağrılarında hesabın adı kullanılır.
    """
    isim = str(veri.get("isim") or "").strip()
    if len(isim) > ISLEM_YAPAN_SINIRI:
        raise ValueError(f"İşlemi yapan adı en fazla {ISLEM_YAPAN_SINIRI} karakter olabilir.")
    return isim or session.get("ad") or session["kullanici"]


def _api_kart_islemi(fn):
    try:
        return fn()
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    except depo.MalzemeBekliyorHatasi as hata:
        return jsonify(hata=str(hata), malzeme_bekliyor_onay_gerekli=True), 409
    except depo.TekrarKartOnayiGerekli as hata:
        return jsonify(hata=str(hata), tekrar_onayi_gerekli=True, mevcut=hata.kartlar), 409
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
            isim=_islem_yapan_adi(veri),
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
            isim=_islem_yapan_adi(veri),
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
            isim=_islem_yapan_adi(veri),
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
            isim=_islem_yapan_adi(veri),
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
            tekrar_onayi=veri.get("tekrar_onayi") is True,
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
    if sonuc.get("geri_baglanan"):
        mesaj += (f" · {sonuc['geri_baglanan']} kartın NO'su eski talebine döndü: ayrılmış kart notları ve "
                  "iş akışıyla geri bağlandı.")
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
            # Oturumda tek önizleme tutulur; token onay formunu bu önizlemeye bağlar. Başka
            # sekmede yeni bir dosya önizlenirse eski sayfanın onayı o dosyayı uygulayamaz.
            token = secrets.token_urlsafe(16)
            session["import_onizleme"] = {"dosya": kayit_adi, "hash": dosya_hash, "token": token,
                "surum": sonuc["surum"], "kaynak_surum": sonuc["kaynak_surum"],
                "sifirlanabilir": sonuc.get("sifirlanabilir", []),
                "not_temizlenebilir": sonuc.get("not_temizlenebilir", [])}
            return render_template("import_onizleme.html", sonuc=sonuc, dosya=guvenli_ad,
                                   onizleme_token=token)
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
    bekleyen = session.get("import_onizleme")
    gelen_token = request.form.get("onizleme_token") or ""
    if bekleyen and not secrets.compare_digest(str(bekleyen.get("token") or ""), gelen_token):
        # Bu sayfa oturumdaki son önizlemeye ait değil (ör. başka sekmede başka dosya
        # önizlendi). Diğer sekmenin önizlemesi tüketilmez; o sekmeden onaylanabilir.
        mesaj = ("Bu önizleme sayfası güncel değil: bu oturumda daha sonra başka bir dosya önizlendi "
                 "(ör. başka bir sekmede). Hiçbir kayıt değişmedi. Son önizlemenin sekmesinden onaylayın "
                 "veya dosyayı yeniden seçin.")
        app.logger.warning("Import onayı reddedildi: önizleme token'ı eşleşmedi")
        return render_template("import_onizleme.html", hata=mesaj, hata_turu="onay",
                               sorunlar=_sorunlari_ayir([mesaj]), dosya=""), 409
    session.pop("import_onizleme", None)
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
```

## `depo.py`

```python
"""PDGM İş Takip Sistemi için Excel tabanlı veri deposu.

Tasarım hedefi:
- kartlar.xlsx uygulamanın source of truth dosyasıdır.
- Tek Python process + çok thread modeli kullanılır.
- Read/modify/write işlemleri RLock ile korunur.
- Yazmalar temp dosya + os.replace ile yapılır.
- Kritik işlemler öncesi yedek alınır.
- Workflow yalnız dört gerçek durumdan oluşur:
    PLANA ALINDI -> DİZGİDE -> HAZIR -> TESLİM EDİLDİ
- Kaynak Excel'de DURUM boşsa kart saklanır fakat operasyon ekranlarında gösterilmez.

Not: Excel transactional database değildir. Aynı data klasörünü birden fazla Python
process'i paylaşmamalıdır. Bu sınır process lock ile açıkça korunur.
"""

from __future__ import annotations

import atexit
import copy
import hashlib
import json
import os
import re
import shutil
import subprocess
import threading
from datetime import date, datetime
from uuid import uuid4

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel


# ---------------------------------------------------------------------------
# Dosyalar ve workflow sabitleri
# ---------------------------------------------------------------------------

KOK = os.path.dirname(os.path.abspath(__file__))
VERI_KLASORU = os.path.join(KOK, "data")
KARTLAR_DOSYA = os.path.join(VERI_KLASORU, "kartlar.xlsx")
LOG_DOSYA = os.path.join(VERI_KLASORU, "islem_logu.xlsx")
YUKLEME_DOSYA = os.path.join(VERI_KLASORU, "yuklemeler.xlsx")
YEDEK_KLASORU = os.path.join(VERI_KLASORU, "yedekler")
PROCESS_KILIT = os.path.join(VERI_KLASORU, "sunucu.lock")

PLANA_ALINDI = "PLANA ALINDI"
DIZGIDE = "DİZGİDE"
HAZIR = "HAZIR"
TESLIM_EDILDI = "TESLİM EDİLDİ"
# Planlanan teslimden önce süren DİZGİDE kartın rozeti; monitör yalnız bu kısa hâlini gösterir.
SURESI_ICINDE = "SÜRESİ İÇİNDE"

# Dizgi tipi: kart hangi kaynak sayfadan geldi (MAKİNE / ELDE DİZGİ / EÜM).
# Bu, workflow'dan (durum) tamamen bağımsız ikinci bir boyuttur; kaynağı yapısal
# olarak (hangi sheet'ten okunduğu) belirlenir, DURUM metninden tahmin edilmez.
DIZGI_TIPI_MAKINE = "MAKİNE"
DIZGI_TIPI_ELLE = "ELLE DİZGİ"
DIZGI_TIPI_EUM = "EÜM'DE DİZGİ"
GECERLI_DIZGI_TIPLERI = {DIZGI_TIPI_MAKINE, DIZGI_TIPI_ELLE, DIZGI_TIPI_EUM}

# UI / filtre kodları ve görünen etiketler
DIZGI_KODLARI = {
    DIZGI_TIPI_MAKINE: "MAKINE",
    DIZGI_TIPI_ELLE: "ELLE",
    DIZGI_TIPI_EUM: "EUM",
}
DIZGI_ETIKETLERI = {
    DIZGI_TIPI_MAKINE: "Makine",
    DIZGI_TIPI_ELLE: "Elle Dizgi",
    DIZGI_TIPI_EUM: "EÜM'de Dizgi",
}
OPERATOR_TIPI_DIZGI = {
    "makine": DIZGI_TIPI_MAKINE,
    "elle_dizgi": DIZGI_TIPI_ELLE,
    "eum_dizgi": DIZGI_TIPI_EUM,
}

AKTIF_DURUMLAR = {PLANA_ALINDI, DIZGIDE}  # HAZIR artık "aktif iş" değil, backlog
GECERLI_DURUMLAR = {PLANA_ALINDI, DIZGIDE, HAZIR, TESLIM_EDILDI}
# Yeni: operasyon ekranlarında (panel/operatör/monitör) görünür durumlar.
# HAZIR bilinçli olarak dışarıda — yalnız admin görsün.
OPERASYONEL_DURUMLAR = {PLANA_ALINDI, DIZGIDE, TESLIM_EDILDI}

# İş akışındaki ilerleme sırası. Excel bir kartı bu sırada geriye çekerse (ör. uygulamada
# DİZGİDE iken Excel hâlâ PLANA ALINDI diyorsa) önizlemede admin kararına bırakılır.
DURUM_ILERLEME = {HAZIR: 0, PLANA_ALINDI: 1, DIZGIDE: 2, TESLIM_EDILDI: 3}

# Admin tablosu sıralaması: aktif üretim > planlanan > backlog > tamamlanan > eksik
SIRALAMA = {
    DIZGIDE: 0,
    PLANA_ALINDI: 1,
    HAZIR: 2,
    TESLIM_EDILDI: 3,
    None: 4,
}

# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------

class DepoHatasi(Exception):
    """Depo katmanının temel exception sınıfı."""


class KartBulunamadi(DepoHatasi):
    pass


class IsKuralHatasi(DepoHatasi):
    pass


class MalzemeBekliyorHatasi(IsKuralHatasi):
    """Malzeme bekleyen bir kartta operatör onayı alınmadan işlem yapılmaya çalışıldığını belirtir.

    IsKuralHatasi'nin bir alt sınıfıdır; ayrı tutulmasının nedeni çağıran tarafın
    (app.py) bunu genel iş kuralı hatalarından ayırt edip frontend'e "onay gerekli"
    bilgisini iletebilmesidir.
    """


class TekrarKartOnayiGerekli(IsKuralHatasi):
    """Aynı Talep NO + Kart Stok No ile görünür kart varken yeni manuel kart için onay gerekir.

    MalzemeBekliyorHatasi gibi ayrı tutulur: app.py bunu "onayla ve tekrar gönder"
    yanıtına çevirir; kartlar, mevcut eşleşmelerin kısa özetidir.
    """

    def __init__(self, mesaj, kartlar):
        super().__init__(mesaj)
        self.kartlar = kartlar


class VeriDogrulamaHatasi(DepoHatasi):
    pass


# ---------------------------------------------------------------------------
# Excel şemaları
# ---------------------------------------------------------------------------

KART_ALANLARI = [
    ("ID", "id"),
    ("Sıra", "sira"),
    ("Talep NO", "talep_no"),
    ("Kart Stok No", "stok_no"),
    ("Talep Sahibi", "talep_sahibi"),
    ("Toplam Adet", "toplam_adet"),
    ("Adet Metni", "adet_metin"),
    ("Plan Haftası", "plan_hafta"),
    ("Plan Başlangıç", "plan_baslama"),
    ("Plan Teslim", "plan_teslim"),
    ("Gerçekleşen Teslim", "gerceklesen_teslim"),
    ("Excel Durumu", "excel_durum"),
    ("PCB", "pcb"),
    ("Dizgi Tipi", "dizgi_tipi"),
    ("Dizgi Sorumlusu", "dizgi_sorumlusu"),
    ("Malzeme Bekliyor", "malzeme_bekliyor"),
    ("Durum", "durum"),
    ("Başlangıç Adedi", "baslangic_adet"),
    ("Tamamlanan Adet", "tamamlanan_adet"),
    ("Başlama Zamanı", "baslama_zamani"),
    ("Üretim Bitiş Zamanı", "bitis_zamani"),
    ("Teslim Zamanı", "teslim_zamani"),
    ("Operatör", "operator"),
    ("Not", "aciklama"),
    ("Son Güncelleme", "guncelleme"),
    ("Listede", "aktif"),
    ("Kaynakta Aktif", "source_active"),
    ("Admin Gizli", "admin_gizli"),
    ("Kaynak", "kaynak"),
    ("Anahtar", "anahtar"),
    ("Kaynak Sayfa", "source_sheet"),
    ("Kaynak Satır ID", "source_row_id"),
    ("Kaynak Anahtar", "source_key"),
    ("Eski Anahtar", "legacy_anahtar"),
]

LOG_ALANLARI = [
    ("Zaman", "zaman"),
    ("Kullanıcı", "kullanici"),
    ("Rol", "rol"),
    ("İşlem", "islem"),
    ("Talep NO", "talep_no"),
    ("Kart Stok No", "stok_no"),
    ("Adet", "adet"),
    ("Detay", "detay"),
    # Operatör hesapları paylaşımlı: hesabın yanında işlemi yapan kişinin adı.
    # Eski log dosyalarında sütun yoktur; _oku başlık adıyla eşlediği için boş okunur.
    ("İşlemi Yapan", "islem_yapan"),
]

YUKLEME_ALANLARI = [
    ("Zaman", "zaman"),
    ("Kullanıcı", "kullanici"),
    ("Dosya", "dosya"),
    ("Okunan Satır", "satir"),
    ("Yeni Kart", "yeni"),
    ("Güncellenen", "guncellenen"),
    ("Kaynakta Olmayan", "pasife_alinan"),
    ("Uyarı", "uyari"),
]

SAYISAL_ALANLAR = {
    "id",
    "sira",
    "toplam_adet",
    "baslangic_adet",
    "tamamlanan_adet",
    "aktif",
    "source_active",
    "admin_gizli",
    "malzeme_bekliyor",
    "adet",
    "satir",
    "yeni",
    "guncellenen",
    "pasife_alinan",
    "uyari",
}

ZORUNLU_KART_ALANLARI = {
    "id",
    "talep_no",
    "stok_no",
    "toplam_adet",
    "tamamlanan_adet",
    "anahtar",
}


# ---------------------------------------------------------------------------
# In-memory state
# ---------------------------------------------------------------------------

_kilit = threading.RLock()
_kartlar: list[dict] = []
_loglar: list[dict] = []
_yuklemeler: list[dict] = []

# islem_logu.xlsx her işlemde baştan yazılır; boyutu her tıklamanın süresini belirler.
# Sınır aşılınca eski kayıtlar yedekler/ altındaki arşiv dosyasına taşınır.
LOG_SINIRI = 5_000
LOG_SAKLA = 2_000

BASLIK_DOLGU = PatternFill("solid", fgColor="0F2027")
BASLIK_YAZI = Font(name="Arial", bold=True, color="FFFFFF", size=11)


# ---------------------------------------------------------------------------
# Genel yardımcılar
# ---------------------------------------------------------------------------

def simdi() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def bugun() -> str:
    return date.today().strftime("%Y-%m-%d")


def _sayi(deger, varsayilan=0) -> int:
    try:
        return int(float(deger))
    except (TypeError, ValueError):
        return varsayilan


def _temiz_metin(deger) -> str:
    return str(deger or "").strip()


def _tam_sayi(deger):
    """İşlem adetlerini yuvarlamaz; bool ve kesirli değerleri reddeder."""
    if isinstance(deger, bool) or not re.fullmatch(r"[+-]?\d+", str(deger).strip()):
        raise ValueError("Adet tam sayı olmalı; kesirli değer kabul edilmez.")
    return int(deger)


def kart_surumu(kart):
    """Aynı saniyedeki değişiklikleri ve yeniden importu da algılar."""
    veri = {alan: (kart.get(alan) if kart.get(alan) != "" else None)
            for _, alan in KART_ALANLARI}
    return hashlib.sha256(json.dumps(veri, sort_keys=True, ensure_ascii=False,
                                     default=str).encode("utf-8")).hexdigest()


def _islem_yapan(isim, kullanici):
    """Paylaşımlı operatör hesabında işlemi yapan kişi; ad verilmezse hesap adı."""
    return _temiz_metin(isim) or _temiz_metin(kullanici)


def _kart_notu_ekle(mevcut, metin, isim):
    """Yeni notu tarih + operatör ismiyle mevcut nota ekler; eski notlar korunur.

    Örnek satır:
      [2026-09-12 21:55:03] Ahmet Yılmaz: PCB eksik geldi
    """
    temiz = re.sub(r"\s+", " ", _temiz_metin(metin)).strip()
    if not temiz:
        return _temiz_metin(mevcut) or None

    kim = _temiz_metin(isim) or "İsimsiz"
    etiket = f"[{simdi()}] {kim}: {temiz}"
    onceki = _temiz_metin(mevcut)
    if onceki:
        return f"{onceki}\n{etiket}"
    return etiket


def _durum_normalize(deger):
    """Dört gerçek workflow durumunu normalize eder; boş değer None olarak kalır."""
    metin = _temiz_metin(deger)
    if not metin:
        return None

    sade = (
        metin.upper()
        .replace("İ", "I")
        .replace("Ğ", "G")
        .replace("Ü", "U")
        .replace("Ş", "S")
        .replace("Ö", "O")
        .replace("Ç", "C")
    )
    sade = re.sub(r"\s+", " ", sade).strip()

    esleme = {
        "PLANA ALINDI": PLANA_ALINDI,
        "DIZGIDE": DIZGIDE,
        "HAZIR": HAZIR,
        "TESLIM EDILDI": TESLIM_EDILDI,
    }
    return esleme.get(sade)


def tarih_coz(deger):
    """Excel veya kullanıcı girdisini YYYY-MM-DD biçimine çevirir."""
    if deger is None:
        return None

    if isinstance(deger, datetime):
        return deger.strftime("%Y-%m-%d")
    if isinstance(deger, date):
        return deger.strftime("%Y-%m-%d")

    if isinstance(deger, (int, float)) and not isinstance(deger, bool):
        try:
            sonuc = from_excel(deger)
            if isinstance(sonuc, (datetime, date)):
                return sonuc.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OverflowError):
            return None

    metin = str(deger).strip()
    if not metin or metin.upper() in {"-", "YOK", "N/A", "NONE"}:
        return None

    # Takvim tarihini koru; saat dilimi dönüşümü bir günü kaydırmamalı.
    if re.match(r"^\d{4}-\d{2}-\d{2}T", metin):
        try:
            return datetime.fromisoformat(metin).date().isoformat()
        except ValueError:
            return None

    if re.fullmatch(r"\d+(?:\.\d+)?", metin):
        try:
            sonuc = from_excel(float(metin))
            if isinstance(sonuc, (datetime, date)):
                return sonuc.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OverflowError):
            pass

    for kalip in (
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(metin, kalip).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def gun_farki(a, b):
    """a - b gün farkını döndürür."""
    if not a or not b:
        return None
    try:
        t1 = datetime.strptime(str(a)[:10], "%Y-%m-%d").date()
        t2 = datetime.strptime(str(b)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return (t1 - t2).days


def _kart_ref(kart_id: int):
    return next((kart for kart in _kartlar if kart.get("id") == kart_id), None)


def _operasyonda_gorunur_mu(kart: dict) -> bool:
    return (
        kart.get("aktif", 1) == 1
        and (kart.get("kaynak") != "EXCEL" or kart.get("source_active", 1) == 1)
        and kart.get("admin_gizli", 0) != 1
        and kart.get("durum") in OPERASYONEL_DURUMLAR
    )

def dizgi_kodu(dizgi_tipi) -> str:
    return DIZGI_KODLARI.get(dizgi_tipi, "MAKINE")


def dizgi_etiketi(dizgi_tipi) -> str:
    return DIZGI_ETIKETLERI.get(dizgi_tipi, DIZGI_ETIKETLERI[DIZGI_TIPI_MAKINE])


def _dizgi_tipi_coz(dizgi_tipi=None, elle_dizgi=None, mevcut=None):
    """Admin API / form değerinden geçerli dizgi_tipi üretir."""
    if dizgi_tipi not in (None, ""):
        tip = _temiz_metin(dizgi_tipi)
        if tip in GECERLI_DIZGI_TIPLERI:
            return tip
        kod = tip.upper().replace("İ", "I").replace("Ü", "U").replace("'", "")
        kod = re.sub(r"\s+", " ", kod).strip()
        if kod in ("MAKINE",):
            return DIZGI_TIPI_MAKINE
        if kod in ("ELLE", "ELLE DIZGI"):
            return DIZGI_TIPI_ELLE
        if kod in ("EUM", "EUMDE DIZGI", "EUM DIZGI"):
            return DIZGI_TIPI_EUM
        raise IsKuralHatasi(
            "Dizgi tipi Makine, Elle Dizgi veya EÜM'de Dizgi olmalı."
        )
    if elle_dizgi is not None:
        return DIZGI_TIPI_ELLE if elle_dizgi else DIZGI_TIPI_MAKINE
    return mevcut or DIZGI_TIPI_MAKINE


def operator_kart_yetkisi_var_mi(kart: dict, operator_tipi: str) -> bool:
    """
    Operatörün kart üzerinde işlem yapıp yapamayacağını kontrol eder.

    Görme yetkisi değildir.
    Sadece operasyon aksiyonları içindir.

    Not: operator_tipi değerleri kullanicilar.json / session ile aynı
    sözlüğü kullanır: "makine", "elle_dizgi", "eum_dizgi". Admin bu
    kontrolün dışındadır; admin yetkisi çağıran taraf (kart_baslat,
    kart_bitir, kart_teslim_et) içinde rol == "admin" kısayoluyla ayrıca
    ele alınır.
    """
    beklenen = OPERATOR_TIPI_DIZGI.get(operator_tipi)
    if not beklenen:
        return False
    return kart.get("dizgi_tipi") == beklenen

def _yonetimde_gorunur_mu(kart: dict) -> bool:
    return kart.get("aktif", 1) == 1 and kart.get("admin_gizli", 0) != 1


# ---------------------------------------------------------------------------
# Dosya okuma/yazma
# ---------------------------------------------------------------------------

def _oku(dosya, alanlar, zorunlu_alanlar=frozenset()):
    if not os.path.exists(dosya):
        return []

    try:
        wb = openpyxl.load_workbook(dosya, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        raise VeriDogrulamaHatasi(
            f"'{os.path.basename(dosya)}' açılamadı: {exc}"
        ) from exc

    try:
        ws = wb[wb.sheetnames[0]]
        satirlar = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not satirlar:
        return []

    basliklar = [str(h or "").strip() for h in satirlar[0]]
    yerlesim = {
        alan: basliklar.index(baslik)
        for baslik, alan in alanlar
        if baslik in basliklar
    }

    eksik = sorted(alan for alan in zorunlu_alanlar if alan not in yerlesim)
    if eksik:
        ters = {alan: baslik for baslik, alan in alanlar}
        eksik_adlar = ", ".join(ters.get(alan, alan) for alan in eksik)
        raise VeriDogrulamaHatasi(
            f"'{os.path.basename(dosya)}' zorunlu sütunları eksik: {eksik_adlar}"
        )

    kayitlar = []
    for satir in satirlar[1:]:
        if not any(hucre not in (None, "") for hucre in satir):
            continue

        kayit = {}
        for _, alan in alanlar:
            index = yerlesim.get(alan)
            deger = satir[index] if index is not None and index < len(satir) else None

            if alan in SAYISAL_ALANLAR:
                kayit[alan] = _sayi(deger, 0) if deger not in (None, "") else None
            elif isinstance(deger, datetime):
                kayit[alan] = deger.strftime("%Y-%m-%d %H:%M:%S")
            else:
                kayit[alan] = str(deger).strip() if deger not in (None, "") else None
        kayitlar.append(kayit)

    return kayitlar

def _excel_hucre_yaz(hucre, deger):
    """'=' ile başlayan kullanıcı metninin Excel formülüne dönüşmesini engeller."""
    if isinstance(deger, str) and len(deger) > 32767:
        raise VeriDogrulamaHatasi(
            "Metin/not geçmişi Excel'in 32.767 karakter sınırını aşıyor. "
            "İşlem kaydedilmedi; mevcut kayıt korundu. Daha kısa bir metin girin."
        )
    hucre.value = deger

    if isinstance(deger, str) and deger.startswith("="):
        hucre.data_type = "s"

def _workbook_uret(alanlar, kayitlar, sayfa_adi):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sayfa_adi
    ws.append([baslik for baslik, _ in alanlar])

    for hucre in ws[1]:
        hucre.fill = BASLIK_DOLGU
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")

    # Bu dosyalar her işlemde baştan yazılır; hız için satır numarası sayaçla verilir
    # (ws.max_row her çağrıda tüm hücreleri tarar, yazım karesel yavaşlardı) ve gövde
    # hücrelerine tek tek font atanmaz. Biçimli çıktı gereken rapor excel_araclari'dadır.
    for satir_no, kayit in enumerate(kayitlar, start=2):
        for sutun_no, (_, alan) in enumerate(alanlar, start=1):
            _excel_hucre_yaz(
                ws.cell(row=satir_no, column=sutun_no),
                kayit.get(alan),
            )

    for sutun, (baslik, alan) in enumerate(alanlar, start=1):
        en = max(
            [len(baslik), 10]
            + [len(str(kayit.get(alan) or "")) for kayit in kayitlar[:300]]
        )
        ws.column_dimensions[get_column_letter(sutun)].width = min(42, en + 3)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return wb


def _diske_zorla(yol):
    """Temp dosya içeriğinin OS page cache'ten diske inmesini zorlar."""
    try:
        fd = os.open(yol, os.O_RDONLY)
    except OSError:
        return
    try:
        os.fsync(fd)
    except OSError:
        pass
    finally:
        os.close(fd)


def _temp_yaz(hedef, alanlar, kayitlar, sayfa_adi):
    os.makedirs(os.path.dirname(hedef), exist_ok=True)
    temp = f"{hedef}.{uuid4().hex}.yeni"
    wb = _workbook_uret(alanlar, kayitlar, sayfa_adi)
    try:
        wb.save(temp)
    finally:
        wb.close()
    _diske_zorla(temp)
    return temp


def _coklu_yaz(dosyalar):
    """Birden fazla Excel dosyasını temp + replace + rollback yaklaşımıyla yazar."""
    os.makedirs(VERI_KLASORU, exist_ok=True)
    temps = []
    backups = {}
    degisen = []
    commit_basarili = False

    try:
        for hedef, alanlar, kayitlar, sayfa_adi in dosyalar:
            temps.append((hedef, _temp_yaz(hedef, alanlar, kayitlar, sayfa_adi)))

        for hedef, _ in temps:
            if os.path.exists(hedef):
                backup = f"{hedef}.{uuid4().hex}.txn.bak"
                shutil.copy2(hedef, backup)
                backups[hedef] = backup
            else:
                backups[hedef] = None

        for hedef, temp in temps:
            os.replace(temp, hedef)
            degisen.append(hedef)

        commit_basarili = True

    except Exception:
        for hedef in reversed(degisen):
            backup = backups.get(hedef)
            try:
                if backup and os.path.exists(backup):
                    os.replace(backup, hedef)
                    backups[hedef] = None
                elif os.path.exists(hedef):
                    os.remove(hedef)
            except OSError:
                pass
        raise
    finally:
        for _, temp in temps:
            if os.path.exists(temp):
                try:
                    os.remove(temp)
                except OSError:
                    pass
        if commit_basarili:
            for backup in backups.values():
                if backup and os.path.exists(backup):
                    try:
                        os.remove(backup)
                    except OSError:
                        pass


def _yaz(dosya, alanlar, kayitlar, sayfa_adi):
    _coklu_yaz([(dosya, alanlar, kayitlar, sayfa_adi)])


def _gunluk_yedek(dosya):
    if not os.path.exists(dosya):
        return
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    hedef = os.path.join(
        YEDEK_KLASORU,
        f"{date.today():%Y%m%d}_{os.path.basename(dosya)}",
    )
    if not os.path.exists(hedef):
        shutil.copy2(dosya, hedef)


ANLIK_YEDEK_SAKLA = 30
GUNLUK_YEDEK_GUN = 90
ANLIK_YEDEK_DESEN = re.compile(r"^\d{8}_\d{6}_")


def yedekleri_buda():
    """Yalnız PDGM naming pattern'li yedekleri sınırlar. Hata olursa sessizce geçer."""
    try:
        if not os.path.isdir(YEDEK_KLASORU):
            return

        anlik = []
        for ad in os.listdir(YEDEK_KLASORU):
            yol = os.path.join(YEDEK_KLASORU, ad)

            if os.path.isdir(yol) and ANLIK_YEDEK_DESEN.match(ad):
                try:
                    anlik.append((os.path.getmtime(yol), yol))
                except OSError:
                    pass
                continue

            if re.fullmatch(r"\d{8}_kartlar\.xlsx", ad, flags=re.IGNORECASE):
                try:
                    yas_gun = (
                        datetime.now() - datetime.fromtimestamp(os.path.getmtime(yol))
                    ).days
                    if yas_gun > GUNLUK_YEDEK_GUN:
                        os.remove(yol)
                except OSError:
                    pass

        anlik.sort(reverse=True)
        for _, yol in anlik[ANLIK_YEDEK_SAKLA:]:
            shutil.rmtree(yol, ignore_errors=True)
    except Exception:  # noqa: BLE001
        pass


def anlik_yedek(etiket: str = "once") -> str:
    """Kart/log/yükleme dosyalarının timestamp'li güvenlik kopyasını alır."""
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    damga = datetime.now().strftime("%Y%m%d_%H%M%S")
    guvenli = re.sub(r"[^0-9A-Za-z_-]+", "_", etiket or "yedek")
    klasor = os.path.join(YEDEK_KLASORU, f"{damga}_{guvenli}_{uuid4().hex[:8]}")
    os.makedirs(klasor, exist_ok=True)

    for dosya in (KARTLAR_DOSYA, LOG_DOSYA, YUKLEME_DOSYA):
        if os.path.exists(dosya):
            shutil.copy2(dosya, os.path.join(klasor, os.path.basename(dosya)))

    yedekleri_buda()
    return klasor


def yedekleri_getir(adet=12):
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    sonuc = []

    for ad in os.listdir(YEDEK_KLASORU):
        yol = os.path.join(YEDEK_KLASORU, ad)

        if os.path.isdir(yol):
            kart_dosyasi = os.path.join(yol, os.path.basename(KARTLAR_DOSYA))
            if not os.path.isfile(kart_dosyasi):
                continue
            try:
                mtime = os.path.getmtime(kart_dosyasi)
                boyut = os.path.getsize(kart_dosyasi)
            except OSError:
                continue

            parcalar = ad.split("_", 2)
            etiket = parcalar[2].replace("_", " ") if len(parcalar) >= 3 else "anlık yedek"
            sonuc.append(
                {
                    "ad": ad,
                    "tip": "Anlık",
                    "etiket": etiket,
                    "zaman": datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S"),
                    "boyut_kb": round(boyut / 1024, 1),
                    "_mtime": mtime,
                }
            )
            continue

        if not os.path.isfile(yol) or not re.fullmatch(
            r"\d{8}_kartlar\.xlsx", ad, flags=re.IGNORECASE
        ):
            continue

        try:
            mtime = os.path.getmtime(yol)
            boyut = os.path.getsize(yol)
        except OSError:
            continue

        sonuc.append(
            {
                "ad": ad,
                "tip": "Günlük",
                "etiket": "günlük otomatik yedek",
                "zaman": datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S"),
                "boyut_kb": round(boyut / 1024, 1),
                "_mtime": mtime,
            }
        )

    sonuc.sort(key=lambda kayit: kayit["_mtime"], reverse=True)
    for kayit in sonuc:
        kayit.pop("_mtime", None)
    return sonuc if adet is None else sonuc[:adet]


def _yedek_kart_dosyasi_bul(yedek_adi):
    yedek_adi = _temiz_metin(yedek_adi)
    if not yedek_adi:
        raise IsKuralHatasi("Yedek seçilmedi.")
    if os.path.basename(yedek_adi) != yedek_adi:
        raise IsKuralHatasi("Geçersiz yedek adı.")

    yedek_kok = os.path.abspath(YEDEK_KLASORU)
    yol = os.path.abspath(os.path.join(yedek_kok, yedek_adi))
    try:
        if os.path.commonpath([yedek_kok, yol]) != yedek_kok:
            raise IsKuralHatasi("Geçersiz yedek yolu.")
    except ValueError as exc:
        raise IsKuralHatasi("Geçersiz yedek yolu.") from exc

    if os.path.isdir(yol):
        aday = os.path.join(yol, os.path.basename(KARTLAR_DOSYA))
        if os.path.isfile(aday):
            return aday
        raise IsKuralHatasi("Seçilen yedekte kartlar.xlsx bulunamadı.")

    if os.path.isfile(yol) and re.fullmatch(
        r"\d{8}_kartlar\.xlsx", yedek_adi, flags=re.IGNORECASE
    ):
        return yol

    raise IsKuralHatasi("Seçilen yedek kullanılamıyor.")


# ---------------------------------------------------------------------------
# Process lock
# ---------------------------------------------------------------------------



def _process_kilit_sahibi():
    if not os.path.exists(PROCESS_KILIT):
        return None

    try:
        with open(PROCESS_KILIT, encoding="utf-8") as f:
            metin = (f.read() or "").strip()
        ilk = metin.split("|", 1)[0].strip()
        return int(ilk) if ilk else None
    except (OSError, ValueError):
        return None


def _pid_calisiyor_mu(pid: int):
    """PID canlı mı? Karar verilemezse True (fail-closed: kilidi koru)."""
    if pid is None or pid <= 0:
        return False

    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        except OSError:
            return True

    try:
        cikti = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return True

    if cikti.returncode != 0:
        return True

    return f'"{pid}"' in (cikti.stdout or "")


def _pid_python_mu(pid: int):
    """PID python/pythonw sürecine mi ait? Karar verilemezse True (fail-closed)."""
    if pid is None or pid <= 0:
        return False

    if os.name != "nt":
        try:
            import sys

            if pid == os.getpid():
                return True
            cmdline_yolu = f"/proc/{pid}/cmdline"
            if os.path.exists(cmdline_yolu):
                with open(cmdline_yolu, "rb") as f:
                    cmd = f.read().decode("utf-8", errors="ignore").lower()
                return "python" in cmd
            return True
        except OSError:
            return True

    try:
        cikti = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return True

    if cikti.returncode != 0:
        return True

    satir = (cikti.stdout or "").strip().lower()
    if f'"{pid}"' not in satir and str(pid) not in satir:
        return False
    return "python.exe" in satir or "pythonw.exe" in satir


def process_kilidi_al():
    """Aynı data klasörünü ikinci PDGM process'inin açmasını atomik olarak engeller.

    Not:
    - Windows'ta os.kill(pid, 0) yerine tasklist kullanılır.
    - Lock dosyası O_EXCL ile atomik oluşturulur.
    - Stale lock: PID ölüyse veya PID canlı ama python değilse (PID reuse) devralınır.
    - Stale reclaim os.replace ile atomik yapılır (remove+O_EXCL yarışı yok).
    - PID canlı ve python ise reddedilir. tasklist başarısızsa fail-closed.
    """
    os.makedirs(VERI_KLASORU, exist_ok=True)
    mevcut_pid = _process_kilit_sahibi()

    if mevcut_pid == os.getpid():
        return

    if os.path.exists(PROCESS_KILIT):
        canli = mevcut_pid is not None and _pid_calisiyor_mu(mevcut_pid)
        python_sureci = canli and _pid_python_mu(mevcut_pid)

        if canli and python_sureci:
            raise RuntimeError(
                f"data/ klasörü başka bir PDGM process tarafından kilitli "
                f"(PID {mevcut_pid}) ve bu process ŞU AN ÇALIŞIYOR. "
                "İkinci sunucu açmayın."
            )

        sahip = mevcut_pid if mevcut_pid is not None else "bilinmiyor"
        if canli and not python_sureci:
            print(
                f"UYARI: data/sunucu.lock PID {sahip} başka bir uygulamaya ait "
                "(PID reuse). Kalıntı kilit devralınıyor."
            )
        else:
            print(
                f"UYARI: data/sunucu.lock artık çalışmayan bir process'e ait "
                f"(PID {sahip}). Kalıntı kilit devralınıyor."
            )

        # Atomik reclaim: remove+O_EXCL yarışını önlemek için stale lock'u
        # benzersiz bir isme taşı. İki process aynı anda denerse yalnız biri
        # os.replace kazanır; diğeri FileNotFoundError alır ve O_EXCL'de kaybeder.
        stale_yol = f"{PROCESS_KILIT}.stale.{uuid4().hex}"
        try:
            os.replace(PROCESS_KILIT, stale_yol)
        except FileNotFoundError:
            stale_yol = None
        except OSError as exc:
            raise RuntimeError(
                "data/sunucu.lock devralınamadı: "
                f"{exc}. Dosyayı manuel silip tekrar deneyin."
            ) from exc
    else:
        stale_yol = None

    bayraklar = os.O_WRONLY | os.O_CREAT | os.O_EXCL

    try:
        fd = os.open(PROCESS_KILIT, bayraklar)
    except FileExistsError as exc:
        if stale_yol:
            try:
                os.remove(stale_yol)
            except OSError:
                pass
        sahip = _process_kilit_sahibi()
        raise RuntimeError(
            f"data/ klasörü başka bir PDGM process tarafından kilitlendi "
            f"(PID {sahip if sahip is not None else 'bilinmiyor'})."
        ) from exc

    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}|{datetime.now():%Y-%m-%d %H:%M:%S}")
            f.flush()
            os.fsync(f.fileno())
    except Exception:
        try:
            os.remove(PROCESS_KILIT)
        except OSError:
            pass
        if stale_yol:
            try:
                os.remove(stale_yol)
            except OSError:
                pass
        raise

    if stale_yol:
        try:
            os.remove(stale_yol)
        except OSError:
            pass

    if getattr(process_kilidi_al, "_atexit_bagli", False):
        return

    def _birak():
        try:
            if _process_kilit_sahibi() == os.getpid():
                os.remove(PROCESS_KILIT)
        except OSError:
            pass

    atexit.register(_birak)
    process_kilidi_al._atexit_bagli = True


# ---------------------------------------------------------------------------
# Kart normalizasyonu ve doğrulama
# ---------------------------------------------------------------------------

def _kart_normalize(kart: dict) -> dict:
    kart = dict(kart)
    kart["id"] = _sayi(kart.get("id"), 0)
    kart["sira"] = _sayi(kart.get("sira"), 0) or None
    kart["toplam_adet"] = _sayi(kart.get("toplam_adet"), 1) or 1
    kart["baslangic_adet"] = _sayi(kart.get("baslangic_adet"), 0)
    kart["tamamlanan_adet"] = _sayi(kart.get("tamamlanan_adet"), 0)
    kart["aktif"] = 0 if kart.get("aktif") == 0 else 1
    kart["source_active"] = 0 if kart.get("source_active") == 0 else 1
    kart["admin_gizli"] = 1 if kart.get("admin_gizli") == 1 else 0
    kart["kaynak"] = _temiz_metin(kart.get("kaynak")) or "EXCEL"
    kart["durum"] = _durum_normalize(kart.get("durum"))

    kart["dizgi_tipi"] = _temiz_metin(kart.get("dizgi_tipi")) or DIZGI_TIPI_MAKINE
    if kart["dizgi_tipi"] not in GECERLI_DIZGI_TIPLERI:
        kart["dizgi_tipi"] = DIZGI_TIPI_MAKINE
    kart["dizgi_sorumlusu"] = _temiz_metin(kart.get("dizgi_sorumlusu")) or None
    kart["malzeme_bekliyor"] = 1 if kart.get("malzeme_bekliyor") == 1 else 0

    for alan in ("plan_baslama", "plan_teslim", "gerceklesen_teslim"):
        ham = kart.get(alan)
        if ham in (None, ""):
            kart[alan] = None
            continue
        cozulmus = tarih_coz(ham)
        if not cozulmus:
            raise VeriDogrulamaHatasi(
                f"Kart {kart.get('id') or '?'}: {alan} geçerli bir tarih değil: {ham!r}"
            )
        kart[alan] = cozulmus

    return kart


def _kart_dogrula(kart: dict):
    kart_id = _sayi(kart.get("id"), -1)
    toplam = _sayi(kart.get("toplam_adet"), 0)
    tamam = _sayi(kart.get("tamamlanan_adet"), 0)
    durum = kart.get("durum")

    if kart_id < 1:
        raise VeriDogrulamaHatasi("Kart ID pozitif tam sayı olmalı.")
    if not _temiz_metin(kart.get("talep_no")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Talep NO boş olamaz.")
    if not _temiz_metin(kart.get("stok_no")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Kart Stok No boş olamaz.")
    if toplam < 1:
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Toplam Adet en az 1 olmalı.")
    if tamam < 0 or tamam > toplam:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: Tamamlanan Adet ({tamam}) 0 ile Toplam Adet ({toplam}) arasında olmalı."
        )
    if durum is not None and durum not in GECERLI_DURUMLAR:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: Geçersiz durum '{durum}'. "
            f"Geçerli durumlar: {', '.join(sorted(GECERLI_DURUMLAR))}."
        )
    if durum in (None, PLANA_ALINDI, HAZIR) and tamam != 0:
        etiket = "Durumu boş" if durum is None else durum
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: {etiket} kartta Tamamlanan Adet 0 olmalı."
        )
    if durum == TESLIM_EDILDI and tamam != toplam:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: {durum} durumunda Tamamlanan Adet Toplam Adet'e eşit olmalı."
        )
    # Kaynak teslim durumunu tarihi bilinmeden de bildirebilir. Tarih uydurulmaz.
    if kart.get("plan_baslama") and kart.get("plan_teslim"):
        if kart["plan_baslama"] > kart["plan_teslim"]:
            raise VeriDogrulamaHatasi(
                f"Kart {kart_id}: Plan başlangıç tarihi plan teslim tarihinden sonra olamaz."
            )
    if not _temiz_metin(kart.get("anahtar")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Anahtar boş olamaz.")
    kimlik = [kart.get(a) for a in ("source_sheet", "source_row_id", "source_key")]
    if any(kimlik):
        if (not all(kimlik) or kart.get("kaynak") != "EXCEL"
                or kimlik[0] not in {"MAKINE", "ELLE", "EUM"}
                or kimlik[2] != kaynak_anahtari(kimlik[0], kimlik[1])
                or kart["anahtar"] != kimlik[2]):
            raise VeriDogrulamaHatasi(f"Kart {kart_id}: Kaynak kimliği tutarsız.")


def _kart_listesi_dogrula(kartlar):
    idler = set()
    anahtarlar = set()

    for kart in kartlar:
        _kart_dogrula(kart)
        if kart["id"] in idler:
            raise VeriDogrulamaHatasi(f"Tekrarlanan kart ID: {kart['id']}")
        if kart["anahtar"] in anahtarlar:
            raise VeriDogrulamaHatasi(f"Tekrarlanan kart anahtarı: {kart['anahtar']}")
        idler.add(kart["id"])
        anahtarlar.add(kart["anahtar"])


# ---------------------------------------------------------------------------
# Başlangıç ve reload
# ---------------------------------------------------------------------------

def _baslatma_uyarisi_yaz(mesaj: str) -> None:
    """Açılış uyarılarını konsola basar ve data/BASLATMA_HATASI.txt'ye ekler."""
    print(mesaj)
    yol = os.path.join(VERI_KLASORU, "BASLATMA_HATASI.txt")
    try:
        with open(yol, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {mesaj}\n")
    except OSError:
        pass


def _bozuk_dosyayi_kenara_al(dosya, alanlar):
    """Okunamayan yardımcı dosyayı yeniden adlandırıp boş liste döner.

    Yalnız islem_logu / yuklemeler için kullanılır. Kart verisi bu yolla sıfırlanmaz.
    """
    try:
        return _oku(dosya, alanlar)
    except VeriDogrulamaHatasi as exc:
        if os.path.exists(dosya):
            damga = datetime.now().strftime("%Y%m%d_%H%M%S")
            bozuk = f"{dosya}.bozuk_{damga}"
            try:
                os.replace(dosya, bozuk)
            except OSError:
                bozuk = dosya
            _baslatma_uyarisi_yaz(
                f"UYARI: '{os.path.basename(dosya)}' okunamadı ({exc}). "
                f"Dosya '{os.path.basename(bozuk)}' olarak kenara alındı; "
                "boş liste ile devam ediliyor. Kart verisi etkilenmedi."
            )
        else:
            _baslatma_uyarisi_yaz(
                f"UYARI: '{os.path.basename(dosya)}' okunamadı ({exc}). "
                "Boş liste ile devam ediliyor."
            )
        return []


def kur():
    global _kartlar, _loglar, _yuklemeler

    with _kilit:
        os.makedirs(VERI_KLASORU, exist_ok=True)

        kartlar = _oku(
            KARTLAR_DOSYA,
            KART_ALANLARI,
            ZORUNLU_KART_ALANLARI if os.path.exists(KARTLAR_DOSYA) else frozenset(),
        )
        kartlar = [_kart_normalize(kart) for kart in kartlar]
        if kartlar:
            _kart_listesi_dogrula(kartlar)

        _kartlar = kartlar
        _loglar = _bozuk_dosyayi_kenara_al(LOG_DOSYA, LOG_ALANLARI)
        _yuklemeler = _bozuk_dosyayi_kenara_al(YUKLEME_DOSYA, YUKLEME_ALANLARI)

        eksikler = []
        if not os.path.exists(KARTLAR_DOSYA):
            eksikler.append((KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"))
        if not os.path.exists(LOG_DOSYA):
            eksikler.append((LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"))
        if not os.path.exists(YUKLEME_DOSYA):
            eksikler.append((YUKLEME_DOSYA, YUKLEME_ALANLARI, _yuklemeler, "Yüklemeler"))
        if eksikler:
            _coklu_yaz(eksikler)


def kartlari_diskten_yeniden_yukle():
    """Manuel Excel müdahalesinden sonra kartlar.xlsx'i validate ederek tekrar yükler."""
    global _kartlar

    with _kilit:
        yeni = _oku(KARTLAR_DOSYA, KART_ALANLARI, ZORUNLU_KART_ALANLARI)
        yeni = [_kart_normalize(kart) for kart in yeni]
        _kart_listesi_dogrula(yeni)
        _kartlar = yeni
        return len(_kartlar)


def yeniden_yukle():
    return kartlari_diskten_yeniden_yukle()


def _kartlari_kaydet():
    _gunluk_yedek(KARTLAR_DOSYA)
    _yaz(KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar")


# ---------------------------------------------------------------------------
# Log ve yükleme geçmişi
# ---------------------------------------------------------------------------

LOG_DETAY_SINIRI = 32_000


def _log_kaydi(kullanici, rol, islem, talep_no="", stok_no="", adet=None, detay="", islem_yapan=None):
    # Kart metni kırpılmaz; log satırı ise Excel hücre sınırını aşıp işlemi engellememeli.
    if isinstance(detay, str) and len(detay) > LOG_DETAY_SINIRI:
        detay = detay[:LOG_DETAY_SINIRI] + " … [log için kısaltıldı; tam metin kart yedeklerinde]"
    return {
        "zaman": simdi(),
        "kullanici": kullanici,
        "rol": rol,
        "islem": islem,
        "talep_no": talep_no,
        "stok_no": stok_no,
        "adet": adet,
        "detay": detay,
        "islem_yapan": _temiz_metin(islem_yapan) or None,
    }


def _log_arsivle_gerekirse():
    """Commit sonrası bakım: log LOG_SINIRI'nı aştıysa eski kayıtları arşive taşır.

    Log ekleyen her commit yolu (log_ekle, kart işlemleri, import, yedekten geri
    yükleme) yazım başarılı olduktan sonra çağırır. İşlem zaten diske yazıldığı için
    arşivleme hata verirse işlem bozulmaz; log olduğu gibi kalır ve bir sonraki
    commit'te yeniden denenir.
    """
    global _loglar

    if len(_loglar) <= LOG_SINIRI:
        return

    arsivlenecek, kalan = _loglar[:-LOG_SAKLA], _loglar[-LOG_SAKLA:]
    arsiv = os.path.join(
        YEDEK_KLASORU,
        f"{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}_islem_logu_arsiv.xlsx",
    )
    try:
        os.makedirs(YEDEK_KLASORU, exist_ok=True)
        _yaz(arsiv, LOG_ALANLARI, arsivlenecek, "İşlem Logu")
        _yaz(LOG_DOSYA, LOG_ALANLARI, kalan, "İşlem Logu")
    except Exception:  # noqa: BLE001
        # Log dosyası kısaltılamadıysa kayıtlar hâlâ onda; arşiv kopyası çift olmasın.
        try:
            os.remove(arsiv)
        except OSError:
            pass
        return
    _loglar = kalan


def log_ekle(kullanici, rol, islem, talep_no="", stok_no="", adet=None, detay=""):
    with _kilit:
        # Log kayıtları değiştirilmez, yalnız eklenir: geri almak için uzunluk yeter.
        log_sayisi = len(_loglar)
        try:
            _loglar.append(
                _log_kaydi(kullanici, rol, islem, talep_no, stok_no, adet, detay)
            )
            _yaz(LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu")
        except Exception:
            del _loglar[log_sayisi:]
            raise
        _log_arsivle_gerekirse()


def loglari_getir(adet=None):
    with _kilit:
        secim = list(reversed(_loglar))
        if adet:
            secim = secim[:adet]
        return copy.deepcopy(secim)


def yuklemeleri_getir(adet=None):
    with _kilit:
        secim = list(reversed(_yuklemeler))
        if adet:
            secim = secim[:adet]
        return copy.deepcopy(secim)


# ---------------------------------------------------------------------------
# Görünüm hesapları
# ---------------------------------------------------------------------------

def durum_bilgisi(kart):
    durum = kart.get("durum")
    plan_baslama = kart.get("plan_baslama")
    plan_teslim = kart.get("plan_teslim")
    bugun_iso = bugun()

    bilgi = {
        "rozet": durum or "DURUMU EKSİK",
        "renk": "notr",
        "sapma": None,
        "kalan": None,
        "zaman_yuzde": 0,
        "plan_gun": gun_farki(plan_teslim, plan_baslama),
    }

    if durum is None:
        bilgi["rozet"] = "DURUMU EKSİK"
        bilgi["renk"] = "uyari"
        return bilgi

    if durum == TESLIM_EDILDI:
        teslim = kart.get("gerceklesen_teslim") or str(kart.get("teslim_zamani") or "")[:10]
        sapma = gun_farki(teslim, plan_teslim)
        bilgi["sapma"] = sapma
        bilgi["zaman_yuzde"] = 100
        if sapma is None:
            bilgi["rozet"], bilgi["renk"] = "TESLİM EDİLDİ", "iyi"
        elif sapma > 0:
            bilgi["rozet"], bilgi["renk"] = f"GEÇ TESLİM (+{sapma} gün)", "kotu"
        else:
            bilgi["rozet"], bilgi["renk"] = "ZAMANINDA TESLİM", "iyi"
        return bilgi

    if durum == HAZIR:
        bilgi["rozet"], bilgi["renk"] = "HAZIR · PLANLANMAYI BEKLİYOR", "notr"
        bilgi["zaman_yuzde"] = 0
        return bilgi

    if durum == DIZGIDE:
        if kart.get("tamamlanan_adet", 0) >= kart.get("toplam_adet", 1):
            bilgi["rozet"], bilgi["renk"] = "ÜRETİM BİTTİ · TESLİME HAZIR", "uyari"
            bilgi["zaman_yuzde"] = 100
            return bilgi

        kalan = gun_farki(plan_teslim, bugun_iso)
        bilgi["kalan"] = kalan
        baslangic = str(kart.get("baslama_zamani") or plan_baslama or bugun_iso)[:10]
        gecen = gun_farki(bugun_iso, baslangic) or 0
        plan_gun = bilgi["plan_gun"]

        if plan_gun and plan_gun > 0:
            bilgi["zaman_yuzde"] = max(0, min(140, round(gecen / plan_gun * 100)))
        else:
            bilgi["zaman_yuzde"] = 100 if kalan is not None and kalan < 0 else 50

        if kalan is None:
            bilgi["rozet"], bilgi["renk"] = "DİZGİDE", "uyari"
        elif kalan < 0:
            bilgi["sapma"] = -kalan
            bilgi["rozet"], bilgi["renk"] = f"SÜRE AŞILDI ({-kalan} gün)", "kotu"
        elif kalan <= 1:
            bilgi["rozet"] = "SON GÜN" if kalan == 0 else "SON 1 GÜN"
            bilgi["renk"] = "uyari"
        else:
            # "SÜRE AŞILDI (N gün)" ile karşıt; parantez sayının neyi anlattığını söyler.
            bilgi["rozet"], bilgi["renk"] = f"{SURESI_ICINDE} (teslime {kalan} gün kaldı)", "iyi"
        return bilgi

    gecikme = gun_farki(bugun_iso, plan_baslama)
    if plan_baslama and gecikme is not None and gecikme > 0:
        bilgi["rozet"], bilgi["renk"], bilgi["sapma"] = (
            f"BAŞLAMADI (+{gecikme} gün)",
            "kotu",
            gecikme,
        )
    elif plan_baslama and gecikme == 0:
        bilgi["rozet"], bilgi["renk"] = "BUGÜN BAŞLAMALI", "uyari"
    else:
        bilgi["rozet"], bilgi["renk"] = PLANA_ALINDI, "notr"
    return bilgi


def kart_gorunumu(kart):
    d = copy.deepcopy(kart)
    d["surum"] = kart_surumu(kart)
    d.update(durum_bilgisi(kart))
    d["toplam_adet"] = d.get("toplam_adet") or 1
    d["tamamlanan_adet"] = d.get("tamamlanan_adet") or 0
    d["baslangic_adet"] = d.get("baslangic_adet") or 0
    d["kalan_adet"] = max(0, d["toplam_adet"] - d["tamamlanan_adet"])
    d["adet_yuzde"] = min(100, round(d["tamamlanan_adet"] / d["toplam_adet"] * 100))
    d["gorunur"] = _operasyonda_gorunur_mu(d)
    d["kaynakta_yok"] = d.get("source_active", 1) != 1
    d["is_durumu"] = d.get("durum") or "DURUMU EKSİK"
    d["kaynak_durumu"] = _temiz_metin(d.get("excel_durum"))
    d["elle_dizgi_mi"] = d.get("dizgi_tipi") == DIZGI_TIPI_ELLE
    d["eum_dizgi_mi"] = d.get("dizgi_tipi") == DIZGI_TIPI_EUM
    d["dizgi_kod"] = dizgi_kodu(d.get("dizgi_tipi"))
    d["dizgi_etiket"] = dizgi_etiketi(d.get("dizgi_tipi"))
    d["malzeme_bekliyor"] = d.get("malzeme_bekliyor") == 1
    return d


def _kartlari_sirala(kartlar):
    kartlar.sort(
        key=lambda kart: (
            SIRALAMA.get(kart.get("durum"), 9),
            kart.get("plan_baslama") or "9999-12-31",
            kart.get("sira") or 999999,
            kart.get("id") or 0,
        )
    )
    return kartlar


def kartlari_getir(sadece_gorunen=True):
    with _kilit:
        secim = [
            kart for kart in _kartlar
            if not sadece_gorunen or _operasyonda_gorunur_mu(kart)
        ]
        kartlar = [kart_gorunumu(kart) for kart in secim]
    return _kartlari_sirala(kartlar)


def kartlari_yonetim_getir():
    """Admin tablosu için, durumu boş kartlar dahil, gizlenmemiş tüm aktif kartlar."""
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if _yonetimde_gorunur_mu(kart)
        ]
    return _kartlari_sirala(kartlar)


def durumu_eksik_kartlari_getir():
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if _yonetimde_gorunur_mu(kart) and not kart.get("durum")
        ]
    return _kartlari_sirala(kartlar)


def gizlenen_kartlari_getir():
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if kart.get("aktif", 1) == 1 and kart.get("admin_gizli", 0) == 1
        ]
    kartlar.sort(key=lambda kart: (kart.get("guncelleme") or "", kart.get("id") or 0), reverse=True)
    return kartlar


def kart_getir(kart_id):
    kart_id = _sayi(kart_id, -1)
    with _kilit:
        kart = _kart_ref(kart_id)
        return kart_gorunumu(kart) if kart else None


def kart_bul(anahtar):
    with _kilit:
        for kart in _kartlar:
            if kart.get("anahtar") == anahtar:
                return copy.deepcopy(kart)
    return None


def yeni_kimlik():
    with _kilit:
        return max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1


# ---------------------------------------------------------------------------
# Kart + log commit yardımcıları
# ---------------------------------------------------------------------------

def _kart_log_commit():
    _gunluk_yedek(KARTLAR_DOSYA)
    _coklu_yaz(
        [
            (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
            (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
        ]
    )


def _atomik_kart_islemi(islem):
    """RAM state rollback kalıbını tek yerde tutar.

    Kartlar in-place mutasyon gördüğü için deepcopy zorunlu.
    Loglara yalnız append yapıldığı için uzunluk + del yeterlidir.
    """
    global _kartlar

    eski_kartlar = copy.deepcopy(_kartlar)
    log_sayisi = len(_loglar)
    try:
        sonuc = islem()
        _kart_listesi_dogrula(_kartlar)
        _kart_log_commit()
    except Exception:
        _kartlar = eski_kartlar
        del _loglar[log_sayisi:]
        raise
    _log_arsivle_gerekirse()
    return sonuc


# ---------------------------------------------------------------------------
# Operatör workflow işlemleri
# ---------------------------------------------------------------------------

def kart_baslat(kart_id, adet, kullanici, rol, operator_tipi, aciklama="", malzeme_onayi=False, isim=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not _operasyonda_gorunur_mu(kart):
            raise IsKuralHatasi("Bu kart operasyon ekranında aktif değil.")
        if rol != "admin" and not operator_kart_yetkisi_var_mi(kart, operator_tipi):
            raise IsKuralHatasi(
                "Bu kart sizin operatör tipiniz için uygun değil."
            )
        
        if kart["durum"] != PLANA_ALINDI:
            raise IsKuralHatasi("Yalnız PLANA ALINDI durumundaki kart DİZGİDE'ye alınabilir.")

        malzeme_bekleniyor_mu = kart.get("malzeme_bekliyor") == 1
        if rol != "admin" and malzeme_bekleniyor_mu and not malzeme_onayi:
            raise MalzemeBekliyorHatasi(
                "Bu kart malzeme bekliyor olarak işaretlenmiş. "
                "Malzemenin tedarik edildiğini onaylıyor musunuz?"
            )

        if adet in (None, ""):
            adet = kart["toplam_adet"]
        try:
            adet = _tam_sayi(adet)
        except (TypeError, ValueError) as exc:
            raise ValueError("Adet sayı olmalı.") from exc
        if adet < 1 or adet > kart["toplam_adet"]:
            raise IsKuralHatasi(f"Başlatılacak adet 1 ile {kart['toplam_adet']} arasında olmalı.")

        def islem():
            yapan = _islem_yapan(isim, kullanici)
            yeni_aciklama = kart.get("aciklama")
            if _temiz_metin(aciklama):
                yeni_aciklama = _kart_notu_ekle(
                    kart.get("aciklama"),
                    aciklama,
                    yapan,
                )

            kart.update(
                durum=DIZGIDE,
                baslangic_adet=adet,
                baslama_zamani=kart.get("baslama_zamani") or simdi(),
                bitis_zamani=None,
                teslim_zamani=None,
                gerceklesen_teslim=None,
                malzeme_bekliyor=0,
                operator=yapan,
                aciklama=yeni_aciklama,
                guncelleme=simdi(),
            )
            if malzeme_bekleniyor_mu:
                if malzeme_onayi:
                    _loglar.append(
                        _log_kaydi(
                            kullanici,
                            rol,
                            "MALZEME TEDARİK ONAYLANDI",
                            kart.get("talep_no") or "",
                            kart.get("stok_no") or "",
                            detay="Operatör malzemenin tedarik edildiğini onayladı; kart dizgiye alındı.",
                            islem_yapan=yapan,
                        )
                    )
                elif rol == "admin":
                    _loglar.append(
                        _log_kaydi(
                            kullanici,
                            rol,
                            "ADMİN MALZEME BEKLEME DURUMUNU ATLADI",
                            kart.get("talep_no") or "",
                            kart.get("stok_no") or "",
                            detay="Admin, malzeme bekleyen kartı onay istemeden doğrudan dizgiye aldı.",
                            islem_yapan=yapan,
                        )
                    )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "DİZGİYE ALINDI",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    adet,
                    aciklama or f"{adet} adet dizgiye alındı",
                    islem_yapan=yapan,
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def kart_bitir(kart_id, adet, kullanici, rol, operator_tipi, aciklama="", isim=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not _operasyonda_gorunur_mu(kart):
            raise IsKuralHatasi("Bu kart operasyon ekranında aktif değil.")
        if rol != "admin" and not operator_kart_yetkisi_var_mi(kart, operator_tipi):
            raise IsKuralHatasi(
                "Bu kart sizin operatör tipiniz için uygun değil."
            )
        if kart.get("durum") != DIZGIDE:
            raise IsKuralHatasi("Tamamlanan adet yalnız DİZGİDE durumundaki karta girilebilir.")

        kalan = kart["toplam_adet"] - kart["tamamlanan_adet"]
        if kalan <= 0:
            raise IsKuralHatasi("Üretim adedi zaten tamamlandı. Kartı HAZIR durumuna alın.")

        try:
            adet = _tam_sayi(adet)
        except (TypeError, ValueError) as exc:
            raise ValueError("Adet sayı olmalı.") from exc
        if adet < 1 or adet > kalan:
            raise IsKuralHatasi(f"Adet 1 ile {kalan} arasında olmalı.")

        def islem():
            yapan = _islem_yapan(isim, kullanici)
            yeni_toplam = kart["tamamlanan_adet"] + adet
            uretim_bitti = yeni_toplam == kart["toplam_adet"]

            yeni_aciklama = kart.get("aciklama")
            if _temiz_metin(aciklama):
                yeni_aciklama = _kart_notu_ekle(
                    kart.get("aciklama"),
                    aciklama,
                    yapan,
                )

            kart.update(
                tamamlanan_adet=yeni_toplam,
                bitis_zamani=simdi() if uretim_bitti else kart.get("bitis_zamani"),
                operator=yapan,
                aciklama=yeni_aciklama,
                guncelleme=simdi(),
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "ÜRETİM ADEDİ TAMAMLANDI" if uretim_bitti else "KISMİ ÜRETİM",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    adet,
                    aciklama or f"{yeni_toplam}/{kart['toplam_adet']} adet tamamlandı",
                    islem_yapan=yapan,
                )
            )

            mesaj = (
                "Üretim adedi tamamlandı. Hazır olduğunuzda kartı Teslim Edildi olarak işaretleyebilirsiniz."
                if uretim_bitti
                else f"{yeni_toplam}/{kart['toplam_adet']} adet tamamlandı."
            )

            return kart_gorunumu(kart), uretim_bitti, mesaj

        return _atomik_kart_islemi(islem)



def kart_teslim_et(kart_id, kullanici, rol, operator_tipi, aciklama="", isim=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not _operasyonda_gorunur_mu(kart):
            raise IsKuralHatasi("Bu kart operasyon ekranında aktif değil.")
        if rol != "admin" and not operator_kart_yetkisi_var_mi(kart, operator_tipi):
            raise IsKuralHatasi(
                "Bu kart sizin operatör tipiniz için uygun değil."
            )
        if kart.get("durum") != DIZGIDE:
            raise IsKuralHatasi("Yalnız DİZGİDE durumundaki kart TESLİM EDİLDİ yapılabilir.")
        if kart.get("tamamlanan_adet", 0) != kart.get("toplam_adet", 0):
            raise IsKuralHatasi(
                "Kart teslim edilmeden önce üretim adedinin tamamı bitirilmelidir."
            )

        def islem():
            yapan = _islem_yapan(isim, kullanici)
            teslim_ani = simdi()
            yeni_aciklama = kart.get("aciklama")
            if _temiz_metin(aciklama):
                yeni_aciklama = _kart_notu_ekle(
                    kart.get("aciklama"),
                    aciklama,
                    yapan,
                )

            kart.update(
                durum=TESLIM_EDILDI,
                gerceklesen_teslim=bugun(),
                teslim_zamani=teslim_ani,
                bitis_zamani=kart.get("bitis_zamani") or teslim_ani,
                operator=yapan,
                aciklama=yeni_aciklama,
                guncelleme=teslim_ani,
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "TESLİM EDİLDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    kart.get("toplam_adet"),
                    aciklama or f"Teslim tarihi: {bugun()}",
                    islem_yapan=yapan,
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)
    



def kart_not_guncelle(kart_id, aciklama, kullanici, rol, isim=""):
    """Yeni not satırını tarih/operatör ismiyle ekler; geçmiş notlar silinmez."""
    kart_id = _sayi(kart_id, -1)
    if not _temiz_metin(aciklama):
        raise IsKuralHatasi("Not boş olamaz.")
    gosterim_adi = _islem_yapan(isim, kullanici)
    if not gosterim_adi:
        raise IsKuralHatasi("Operatör ismi boş olamaz.")

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if rol != "admin" and not _operasyonda_gorunur_mu(kart):
            raise IsKuralHatasi("Bu kart operasyon ekranında aktif değil.")

        def islem():
            kart["aciklama"] = _kart_notu_ekle(
                kart.get("aciklama"),
                aciklama,
                gosterim_adi,
            )
            kart["guncelleme"] = simdi()
            # Logda yalnız eklenen satır: tüm geçmişi her notta tekrar yazmak logu
            # karesel büyütüyordu. Geçmişin tamamı kartta duruyor.
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "NOT EKLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay=(kart.get("aciklama") or "").rsplit("\n", 1)[-1],
                    islem_yapan=gosterim_adi,
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Admin kart işlemleri
# ---------------------------------------------------------------------------

def _tarih_form_degeri(deger, alan_adi):
    if deger in (None, ""):
        return None
    sonuc = tarih_coz(deger)
    if not sonuc:
        raise IsKuralHatasi(f"{alan_adi} geçerli bir tarih olmalı.")
    return sonuc


def admin_kart_ekle(
    talep_no,
    stok_no,
    toplam_adet,
    kullanici,
    sira=None,
    talep_sahibi="",
    plan_hafta="",
    plan_baslama=None,
    plan_teslim=None,
    gerceklesen_teslim=None,
    pcb="",
    aciklama="",
    elle_dizgi=False,
    dizgi_tipi=None,
    dizgi_sorumlusu="",
    tekrar_onayi=False,
):
    """Admin panelinden PLANA ALINDI durumunda manuel kart oluşturur.

    Aynı Talep NO + Kart Stok No ile görünür bir kart varsa tekrar_onayi=True
    verilmedikçe TekrarKartOnayiGerekli fırlatılır.
    """
    global _kartlar, _loglar

    talep_no = _temiz_metin(talep_no)
    stok_no = _temiz_metin(stok_no)
    if not talep_no:
        raise IsKuralHatasi("Talep NO boş olamaz.")
    if not stok_no:
        raise IsKuralHatasi("Kart Stok No boş olamaz.")

    try:
        toplam = _tam_sayi(toplam_adet)
    except (TypeError, ValueError) as exc:
        raise ValueError("Toplam adet sayı olmalı.") from exc
    if toplam < 1:
        raise IsKuralHatasi("Toplam adet en az 1 olmalı.")

    if sira in (None, ""):
        sira_degeri = None
    else:
        try:
            sira_degeri = int(sira)
        except (TypeError, ValueError) as exc:
            raise ValueError("Sıra tam sayı olmalı.") from exc

    plan_baslama_iso = _tarih_form_degeri(plan_baslama, "Dizgi Başlama Tarihi")
    plan_teslim_iso = _tarih_form_degeri(plan_teslim, "Planlanan Teslim Tarihi")
    gerceklesen_iso = _tarih_form_degeri(gerceklesen_teslim, "Gerçekleşen Teslim Tarihi")

    if plan_baslama_iso and plan_teslim_iso and plan_baslama_iso > plan_teslim_iso:
        raise IsKuralHatasi("Dizgi Başlama Tarihi Planlanan Teslim Tarihinden sonra olamaz.")
    if gerceklesen_iso:
        raise IsKuralHatasi(
            "Yeni kart PLANA ALINDI durumunda başlar; Gerçekleşen Teslim Tarihi başlangıçta boş olmalı."
        )

    secilen_dizgi = _dizgi_tipi_coz(dizgi_tipi=dizgi_tipi, elle_dizgi=elle_dizgi)

    with _kilit:
        # Excel'den gelmiş ya da elle eklenmiş, listelerde görünen aynı Talep+Stok kartları.
        # Aynı talep ve stok birden fazla sipariş satırında meşru olabildiği için engellenmez;
        # yanlışlıkla çift kart açılıp sayaçlarda iki kez sayılmasın diye açık onay istenir.
        ayni = [
            kart for kart in _kartlar
            if _yonetimde_gorunur_mu(kart)
            and (kart.get("kaynak") != "EXCEL" or kart.get("source_active", 1) == 1)
            and _temiz_metin(kart.get("talep_no")).casefold() == talep_no.casefold()
            and _temiz_metin(kart.get("stok_no")).casefold() == stok_no.casefold()
        ]
        if ayni and not tekrar_onayi:
            raise TekrarKartOnayiGerekli(
                f"{talep_no} · {stok_no} için listelerde {len(ayni)} kart zaten var. Aynı iş için ikinci "
                "kart açılırsa Pano ve raporlarda iki kez sayılır.",
                [{"id": kart["id"], "kaynak": kart.get("kaynak"), "durum": kart.get("durum") or "DURUMU EKSİK",
                  "dizgi_etiket": dizgi_etiketi(kart.get("dizgi_tipi")), "toplam_adet": kart.get("toplam_adet")}
                 for kart in ayni],
            )

        def islem():
            kart_id = max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1
            # Aynı Talep+Stok'lu birden çok kart olabildiği için anahtar kart kimliğinden üretilir.
            anahtar = f"MANUEL:{kart_id}"
            kayit = {
                "id": kart_id,
                "sira": sira_degeri,
                "talep_no": talep_no,
                "stok_no": stok_no,
                "talep_sahibi": _temiz_metin(talep_sahibi) or None,
                "toplam_adet": toplam,
                "adet_metin": f"{toplam} ADET",
                "plan_hafta": _temiz_metin(plan_hafta) or None,
                "plan_baslama": plan_baslama_iso,
                "plan_teslim": plan_teslim_iso,
                "gerceklesen_teslim": None,
                "excel_durum": "MANUEL",
                "pcb": _temiz_metin(pcb) or None,
                "dizgi_tipi": secilen_dizgi,
                "dizgi_sorumlusu": _temiz_metin(dizgi_sorumlusu) or None,
                "malzeme_bekliyor": 0,
                "durum": PLANA_ALINDI,
                "baslangic_adet": 0,
                "tamamlanan_adet": 0,
                "baslama_zamani": None,
                "bitis_zamani": None,
                "teslim_zamani": None,
                "operator": None,
                "aciklama": _temiz_metin(aciklama) or None,
                "guncelleme": simdi(),
                "aktif": 1,
                "source_active": 1,
                "admin_gizli": 0,
                "kaynak": "MANUEL",
                "anahtar": anahtar,
            }
            _kartlar.append(kayit)
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "MANUEL KART EKLENDİ",
                    talep_no,
                    stok_no,
                    toplam,
                    f"{toplam} adet · Durum: {PLANA_ALINDI}"
                    + f" · {dizgi_etiketi(secilen_dizgi)}"
                    + (f" · aynı Talep+Stok ile {len(ayni)} kart vardı (ID: "
                       f"{', '.join(str(kart['id']) for kart in ayni)}); admin onayıyla eklendi" if ayni else ""),
                )
            )
            return kart_gorunumu(kayit)

        return _atomik_kart_islemi(islem)


# Admin düzenlemesinde denetim loguna eski/yeni değeri yazılan alanlar (not ayrı loglanır).
ADMIN_IZLENEN_ALANLAR = (
    "durum", "toplam_adet", "tamamlanan_adet", "baslangic_adet", "plan_hafta", "plan_baslama",
    "plan_teslim", "gerceklesen_teslim", "baslama_zamani", "bitis_zamani", "teslim_zamani",
    "dizgi_tipi", "dizgi_sorumlusu", "malzeme_bekliyor",
)


def admin_kart_duzenle(
    kart_id,
    durum,
    tamamlanan_adet,
    toplam_adet,
    aciklama,
    kullanici,
    plan_hafta=None,
    plan_baslama=None,
    plan_teslim=None,
    gerceklesen_teslim=None,
    elle_dizgi=None,
    dizgi_tipi=None,
    dizgi_sorumlusu=None,
    malzeme_bekliyor=None,
    beklenen_surum=None,
):
    """Admin workflow ve plan alanlarını kontrollü biçimde düzeltir.

    dizgi_tipi / elle_dizgi / dizgi_sorumlusu / malzeme_bekliyor: None verilirse
    mevcut değer korunur (plan_hafta vb. alanlarla aynı "None = değiştirme"
    kuralı). Excel her zaman otorite olmaya devam eder: bir sonraki Excel içe
    aktarımında bu alanlar (dizgi_tipi, dizgi_sorumlusu, malzeme_bekliyor)
    kaynak Excel'deki değere göre yeniden hesaplanır ve admin'in buradaki
    değişikliği ezilir. Bu bilinçli bir tasarım kararıdır: kaynak Excel tek
    doğruluk kaynağıdır.
    """
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        if beklenen_surum is not None and beklenen_surum != kart_surumu(kart):
            raise IsKuralHatasi(
                "Kart siz formu açtıktan sonra değişti. Kaydedilmedi. "
                "Girdiğiniz metni kopyalayın, sayfayı yenileyip güncel kaydı kontrol edin."
            )
        yeni_durum = _durum_normalize(durum)
        if yeni_durum not in GECERLI_DURUMLAR:
            raise IsKuralHatasi("Durum PLANA ALINDI, DİZGİDE, HAZIR veya TESLİM EDİLDİ olmalı.")

        try:
            toplam = _tam_sayi(kart["toplam_adet"] if toplam_adet in (None, "") else toplam_adet)
            tamamlanan = _tam_sayi(
                kart["tamamlanan_adet"] if tamamlanan_adet in (None, "") else tamamlanan_adet
            )
        except (TypeError, ValueError) as exc:
            raise ValueError("Adetler sayı olmalı.") from exc

        if toplam < 1 or tamamlanan < 0 or tamamlanan > toplam:
            raise IsKuralHatasi("Adet değerleri tutarsız.")

        yeni_plan_hafta = (
            kart.get("plan_hafta")
            if plan_hafta is None
            else (_temiz_metin(plan_hafta) or None)
        )
        yeni_plan_baslama = (
            kart.get("plan_baslama")
            if plan_baslama is None
            else _tarih_form_degeri(plan_baslama, "Dizgi Başlama Tarihi")
        )
        yeni_plan_teslim = (
            kart.get("plan_teslim")
            if plan_teslim is None
            else _tarih_form_degeri(plan_teslim, "Planlanan Teslim Tarihi")
        )
        yeni_gerceklesen = (
            kart.get("gerceklesen_teslim")
            if gerceklesen_teslim is None
            else _tarih_form_degeri(gerceklesen_teslim, "Gerçekleşen Teslim Tarihi")
        )

        if yeni_plan_baslama and yeni_plan_teslim and yeni_plan_baslama > yeni_plan_teslim:
            raise IsKuralHatasi("Dizgi Başlama Tarihi Planlanan Teslim Tarihinden sonra olamaz.")

        if dizgi_tipi is None and elle_dizgi is None:
            yeni_dizgi_tipi = kart.get("dizgi_tipi")
        else:
            yeni_dizgi_tipi = _dizgi_tipi_coz(
                dizgi_tipi=dizgi_tipi,
                elle_dizgi=elle_dizgi,
                mevcut=kart.get("dizgi_tipi"),
            )
        yeni_dizgi_sorumlusu = (
            kart.get("dizgi_sorumlusu")
            if dizgi_sorumlusu is None
            else (_temiz_metin(dizgi_sorumlusu) or None)
        )
        yeni_malzeme_bekliyor = (
            kart.get("malzeme_bekliyor", 0)
            if malzeme_bekliyor is None
            else (1 if malzeme_bekliyor else 0)
        )

        onceki_durum = kart.get("durum") or "DURUMU EKSİK"

        def islem():
            nonlocal tamamlanan, yeni_gerceklesen

            onceki = copy.deepcopy(kart)
            baslama = kart.get("baslama_zamani")
            bitis = kart.get("bitis_zamani")
            teslim_zamani = kart.get("teslim_zamani")
            baslangic_adet = kart.get("baslangic_adet") or 0

            if yeni_durum in (PLANA_ALINDI, HAZIR):
                tamamlanan = 0
                baslangic_adet = 0
                baslama = None
                bitis = None
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == DIZGIDE:
                if onceki_durum != DIZGIDE:
                    baslama = baslama or simdi()
                baslangic_adet = baslangic_adet or toplam
                bitis = bitis if tamamlanan == toplam else None
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == TESLIM_EDILDI:
                tamamlanan = toplam
                if onceki_durum != TESLIM_EDILDI:
                    baslama = baslama or simdi()
                    bitis = bitis or simdi()
                    yeni_gerceklesen = yeni_gerceklesen or bugun()
                    teslim_zamani = teslim_zamani or simdi()
                elif yeni_gerceklesen != kart.get("gerceklesen_teslim"):
                    teslim_zamani = None

            kart.update(
                durum=yeni_durum,
                toplam_adet=toplam,
                tamamlanan_adet=tamamlanan,
                baslangic_adet=baslangic_adet,
                plan_hafta=yeni_plan_hafta,
                plan_baslama=yeni_plan_baslama,
                plan_teslim=yeni_plan_teslim,
                gerceklesen_teslim=yeni_gerceklesen,
                baslama_zamani=baslama,
                bitis_zamani=bitis,
                teslim_zamani=teslim_zamani,
                aciklama=_temiz_metin(aciklama) if aciklama is not None else kart.get("aciklama"),
                dizgi_tipi=yeni_dizgi_tipi,
                dizgi_sorumlusu=yeni_dizgi_sorumlusu,
                malzeme_bekliyor=yeni_malzeme_bekliyor,
                guncelleme=simdi(),
            )

            # Denetim izi: değişen her alanın eski/yeni değeri (import'taki "EXCEL KART
            # GÜNCELLENDİ" gibi). Boş metin ile None aynı sayılır.
            def deger(k, alan):
                return None if k.get(alan) in (None, "") else k.get(alan)

            degisen = {alan: [deger(onceki, alan), deger(kart, alan)] for alan in ADMIN_IZLENEN_ALANLAR
                       if deger(onceki, alan) != deger(kart, alan)}
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "ADMİN DÜZENLEDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    tamamlanan,
                    f"{onceki_durum} → {yeni_durum} · {tamamlanan}/{toplam} adet · ID={kart['id']} · "
                    + (f"değişenler: {json.dumps(degisen, ensure_ascii=False)}" if degisen
                       else "alan değişikliği yok"),
                )
            )
            # Not geçmişi düzenlenebildiği için silinen satırlar ayrı kayıtta saklanır.
            eski_not, yeni_not = _temiz_metin(onceki.get("aciklama")), _temiz_metin(kart.get("aciklama"))
            if eski_not != yeni_not:
                eski_satirlar, yeni_satirlar = eski_not.splitlines(), yeni_not.splitlines()
                eski_kume, yeni_kume = set(eski_satirlar), set(yeni_satirlar)
                silinen = [s for s in eski_satirlar if s not in yeni_kume]
                eklenen = [s for s in yeni_satirlar if s not in eski_kume]
                _loglar.append(
                    _log_kaydi(
                        kullanici,
                        "admin",
                        "ADMİN NOT DÜZENLEDİ",
                        kart.get("talep_no") or "",
                        kart.get("stok_no") or "",
                        detay=(f"ID={kart['id']} · {len(silinen)} satır silindi, {len(eklenen)} satır eklendi"
                               + ("\nSilinen:\n" + "\n".join(silinen) if silinen else "")
                               + ("\nEklenen:\n" + "\n".join(eklenen) if eklenen else "")),
                    )
                )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def admin_kart_gizle(kart_id, kullanici):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        def islem():
            kart["admin_gizli"] = 1
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "KART LİSTEDEN GİZLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay="Kart silinmedi; admin gizli olarak işaretlendi.",
                )
            )

        _atomik_kart_islemi(islem)


def admin_kart_geri_getir(kart_id, kullanici):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if kart.get("admin_gizli", 0) != 1:
            raise IsKuralHatasi("Bu kart zaten görünür durumda.")

        def islem():
            kart["admin_gizli"] = 0
            kart["aktif"] = 1
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "GİZLENEN KART GERİ GETİRİLDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay="Kart tekrar yönetim/operasyon listelerine alındı.",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Yedekten geri yükleme
# ---------------------------------------------------------------------------

def yedekten_geri_yukle(yedek_adi, kullanici):
    global _kartlar, _loglar

    with _kilit:
        aday = _yedek_kart_dosyasi_bul(yedek_adi)
        yeni_kartlar = _oku(aday, KART_ALANLARI, ZORUNLU_KART_ALANLARI)
        yeni_kartlar = [_kart_normalize(kart) for kart in yeni_kartlar]
        if not yeni_kartlar:
            raise VeriDogrulamaHatasi("Seçilen yedekte hiç kart bulunmuyor.")
        _kart_listesi_dogrula(yeni_kartlar)

        koruma_yedegi = anlik_yedek("geri_yukleme_oncesi")
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)

        try:
            _kartlar = yeni_kartlar
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "YEDEKTEN GERİ YÜKLENDİ",
                    detay=(
                        f"Yedek: {yedek_adi} · {len(_kartlar)} kart · "
                        f"koruma: {os.path.basename(koruma_yedegi)}"
                    ),
                )
            )
            _coklu_yaz(
                [
                    (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
                    (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
                ]
            )
        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            raise
        _log_arsivle_gerekirse()

        return {
            "kart": len(_kartlar),
            "yedek": yedek_adi,
            "koruma_yedegi": koruma_yedegi,
        }


# ---------------------------------------------------------------------------
# Excel import commit
# ---------------------------------------------------------------------------

def kaynak_anahtari(source_sheet, source_row_id):
    """Sürümü belirtilmiş, ayraç çakışması olmayan kalıcı kaynak kimliği."""
    return "EXCEL:v2:" + json.dumps([source_sheet, source_row_id], ensure_ascii=False, separators=(",", ":"))


def _import_eslesmelerini_bul(satirlar):
    """Tüm eşleşmeleri mutasyon öncesi çöz; belirsiz legacy satırları tahmin etme."""
    _kart_listesi_dogrula(_kartlar)
    harita = {k["source_key"]: k for k in _kartlar if k.get("source_key")}
    eski_gruplar = {}
    gelen_gruplar = {}
    for k in _kartlar:
        if k.get("kaynak") == "EXCEL" and not k.get("source_key"):
            grup = (dizgi_kodu(k.get("dizgi_tipi")), k.get("talep_no"), k.get("stok_no"))
            eski_gruplar.setdefault(grup, []).append(k)
    gorulen = set()
    for s in satirlar:
        key = s.get("source_key")
        if (not s.get("source_sheet") or not s.get("source_row_id")
                or key != kaynak_anahtari(s["source_sheet"], s["source_row_id"])
                or key != s.get("anahtar") or key in gorulen):
            raise VeriDogrulamaHatasi("Eksik veya tekrarlanan kaynak satır kimliği; import iptal edildi.")
        gorulen.add(key)
        grup = (s["source_sheet"], s["talep_no"], s["stok_no"])
        gelen_gruplar.setdefault(grup, []).append(s)

    kullanilan_idler = set()
    for grup, gelenler in gelen_gruplar.items():
        adaylar = eski_gruplar.get(grup, [])
        # Eski Sıra dönüşümü "001" ve "1" ayrımını kaybetmiştir. Bu bilgi
        # artık geri çıkarılamaz; birine rastgele geçmiş atamak yerine dur.
        for s in gelenler:
            kimlik = s["source_row_id"]
            if kimlik.startswith("NO:") and s["source_key"] not in harita:
                no = kimlik[3:]
                sayi = _sayi(no, None)
                if sayi is not None and str(sayi) != no and any(k.get("sira") == sayi for k in adaylar):
                    raise VeriDogrulamaHatasi(
                        f"{grup}: NO={no!r} eski Sıra={sayi} ile belirsiz eşleşiyor. "
                        "Kaynak kimliğini yedek üzerinde açıkça eşleyin."
                    )
        for s in gelenler:
            key = s["source_key"]
            if key in harita:
                continue
            no = s["source_row_id"].removeprefix("NO:")
            kesin = [k for k in adaylar if k.get("sira") is not None
                     and s["source_row_id"].startswith("NO:") and str(k["sira"]) == no]
            if len(kesin) == 1:
                secilen = kesin[0]
            elif len(kesin) > 1:
                raise VeriDogrulamaHatasi(f"{grup}: Eski kayıtlarda aynı NO birden fazla; migration belirsiz.")
            elif (len(adaylar) == 1 and len(gelenler) == 1
                  and adaylar[0].get("sira") is None):
                secilen = adaylar[0]
            elif any(k.get("sira") is None for k in adaylar):
                raise VeriDogrulamaHatasi(
                    f"{grup}: Eski kartların kaynak satırı belirsiz. Yedek üzerinde "
                    "NO/Sıra eşlemesini doğrulayın; tarih veya adet ile otomatik eşleştirme yapılmadı."
                )
            else:
                continue
            if secilen["id"] in kullanilan_idler:
                raise VeriDogrulamaHatasi(f"{grup}: Bir eski kart birden fazla satıra eşleşiyor.")
            kullanilan_idler.add(secilen["id"])
            harita[key] = secilen
    return harita


def depo_surumu():
    with _kilit:
        return hashlib.sha256("|".join(kart_surumu(k) for k in _kartlar).encode()).hexdigest()


def excel_import_onizle(**veri):
    """Aynı merge algoritmasını yalıtılmış bellek kopyasında, yazmadan çalıştırır."""
    global _kartlar, _loglar, _yuklemeler
    with _kilit:
        asil = (_kartlar, _loglar, _yuklemeler)
        _kartlar, _loglar, _yuklemeler = copy.deepcopy(asil)
        try:
            return excel_import_uygula(**veri, sadece_onizleme=True)
        finally:
            _kartlar, _loglar, _yuklemeler = asil


def _import_plani(plan):
    """Boş kaynak metnini diskten okunan kartlarla aynı biçimde (None) tutar.

    _oku boş hücreyi None döndürür; bellekte "" kalırsa sunucu yeniden
    başladıktan sonra aynı dosyanın her yüklenmesi "" -> None hayali alan
    değişikliği, gereksiz audit logu ve önizleme satırı üretir.
    """
    return {alan: (None if isinstance(deger, str) and not deger.strip() else deger)
            for alan, deger in plan.items()}


def _uygulamada_ilerletildi_mi(kart):
    """Kartın şu anki durumu uygulamada (operatör/admin) mı verildi?

    Import zaman damgası yazmaz; başlama/teslim zamanı yalnız uygulamadaki
    işlemlerle (dizgiye al, teslim et, admin düzenleme) oluşur.

    Excel'den DİZGİDE gelmiş kartta başlama zamanı yoktur; operatör yine de adet
    girmiş olabilir. Tamamlanan adet > 0 ve son işlemi yapan Excel değilse (import
    operatör alanına yalnız "Excel" yazar) ilerleme uygulamadadır; eski bir Excel
    TESLİM'inden kalan adet bu sayılmaz.
    """
    if kart.get("durum") == TESLIM_EDILDI:
        return bool(kart.get("teslim_zamani"))
    if kart.get("durum") == DIZGIDE:
        operatorde_adet = (_sayi(kart.get("tamamlanan_adet"), 0) > 0
                           and _temiz_metin(kart.get("operator")) not in ("", "Excel"))
        return bool(kart.get("baslama_zamani")) or operatorde_adet
    return False


def _karar_sayisi(kararlar, tur, korunan=False):
    """Önizlemedeki durum kararlarından bir türün sayısı; korunan=True ise
    uygulamadaki durumun korunduğu (Excel'deki durumun seçilmediği) kartlar."""
    return sum(1 for k in kararlar.values()
               if k["karar_turu"] == tur and (not korunan or k["secilen"] != k["excel_durum"]))


def excel_import_uygula(dosya_adi, kullanici, satirlar, uyari_sayisi=0,
                       sadece_onizleme=False, beklenen_surum=None, kaynak_ozeti="",
                       tamamlanan_sifirla=(), gerileme_secimleri=None, notlari_temizle=(),
                       kaynak_durum_coz=None):
    """Tamamen parse/validate edilmiş kaynak satırlarını tek transaction-benzeri blokta uygular.

    tamamlanan_sifirla: önizlemede admin'in seçtiği kart ID'leri. Yalnız bu importta
    güncellenen ve sonunda DİZGİDE kalan, tamamlanan adedi 0'dan büyük kartlar
    sıfırlanabilir (PLANA/HAZIR'da adet zaten 0, TESLİM'de toplama eşittir).
    Uygun olmayan bir ID tüm importu iptal eder.

    gerileme_secimleri: Excel'in uygulamanın gerisinde kaldığı kartlar için admin'in
    önizlemede seçtiği durum {kart_id: durum}. Seçenekler Excel'in durumu ile
    uygulamadaki durum arasındadır. None ise (önizlemesiz doğrudan import) eski
    sözleşme geçerlidir: Excel'in durumu uygulanır; önizleme ise önerilen seçimi
    gösterir. Sözlük verilirse her gerileyen kart için geçerli seçim zorunludur.

    Excel'de DURUM boş (veya MALZEME TEDARİK / PDGM ÖNERİ gibi bir iş akışı durumu
    olmayan metin) iken uygulamada durumu olan kart da aynı karar listesine girer:
    seçenekler durumsuz bırakmak (None, form değeri "") ya da uygulamadaki durumu
    korumaktır. Önizlemesiz doğrudan importta eski sözleşme sürer: durum korunur.

    notlari_temizle: önizlemede "Notları temizle" işaretlenen kart ID'leri. Yalnız bu
    aktarımda durumu değişen veya durum kararı istenen, notu olan kartlar seçilebilir.
    Silinen not metni işlem loguna yazılır. Uygun olmayan bir ID importu iptal eder.

    kaynak_durum_coz: kartta saklı ham Excel DURUM metnini iş akışı durumuna çeviren
    fonksiyon (excel_araclari verir). Boş DURUM kararında önerilen seçimi belirler:
    kartın durumu önceki Excel metninden gelmediyse uygulamada verilmiştir.
    """
    global _kartlar, _loglar, _yuklemeler

    with _kilit:
        onceki_surum = depo_surumu()
        if beklenen_surum is not None and beklenen_surum != onceki_surum:
            raise IsKuralHatasi("Önizlemeden sonra kartlar değişti. Güncel etkiyi görmek için dosyayı yeniden seçin.")
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)
        eski_yuklemeler = copy.deepcopy(_yuklemeler)

        try:
            yedek_klasoru = None if sadece_onizleme else anlik_yedek("import_oncesi")
            mevcut_harita = _import_eslesmelerini_bul(satirlar)
            gorulen = set()
            yeni = 0
            guncellenen = 0
            degismeyen = 0
            workflow_korundu = 0
            durum_iyilesen = 0
            elle_dizgi_satir = sum(
                1
                for satir in satirlar
                if satir.get("plan", {}).get("dizgi_tipi") == DIZGI_TIPI_ELLE
            )
            eum_dizgi_satir = sum(
                1
                for satir in satirlar
                if satir.get("plan", {}).get("dizgi_tipi") == DIZGI_TIPI_EUM
            )
            pasife_listesi = []
            ayrilanlar = {}
            geri_baglananlar = set()
            sifirla = {_sayi(kart_id, -1) for kart_id in (tamamlanan_sifirla or ())}
            sifirlanabilir = set()
            sifirlanabilir_olasi = set()
            sifirlanan = 0
            secimler = (None if gerileme_secimleri is None
                        else {_sayi(k, -1): v for k, v in gerileme_secimleri.items()})
            gerilemeler = {}
            not_temizle = {_sayi(kart_id, -1) for kart_id in (notlari_temizle or ())}
            not_temizlenebilir = set()
            notu_temizlenen = 0
            onceki_durumu_coz = kaynak_durum_coz or _durum_normalize
            sonraki_id = max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1

            for satir in satirlar:
                anahtar = satir["anahtar"]
                gorulen.add(anahtar)
                plan = _import_plani(satir["plan"])
                mevcut = mevcut_harita.get(anahtar)

                if (mevcut and mevcut.get("source_key") == anahtar
                        and _temiz_metin(mevcut.get("talep_no")) != satir["talep_no"]):
                    # Aynı sayfa+NO artık başka bir talebi taşıyor. Eski talebin notu,
                    # operatörü ve iş akışı yeni talebe geçmesin: eski kart kendi
                    # geçmişiyle ayrı bir kimliğe alınıp pasifleşir, yeni talep için
                    # aşağıda temiz kart açılır.
                    eski_kimlik = mevcut["source_row_id"]
                    ayrik_kimlik = f"{eski_kimlik}~{mevcut['id']}"
                    ayrik_anahtar = kaynak_anahtari(mevcut["source_sheet"], ayrik_kimlik)
                    ayrilanlar[mevcut["id"]] = {"talep_no": satir["talep_no"], "stok_no": satir["stok_no"]}
                    _loglar.append(_log_kaydi(
                        kullanici, "admin", "EXCEL NO BAŞKA TALEBE VERİLDİ", mevcut["talep_no"], mevcut["stok_no"],
                        detay=(f"ID={mevcut['id']} kaynak={anahtar}: {mevcut['talep_no']}/{mevcut['stok_no']} -> "
                               f"{satir['talep_no']}/{satir['stok_no']}; eski kart pasifleşti, yeni kart açıldı"),
                    ))
                    mevcut.update(source_row_id=ayrik_kimlik, source_key=ayrik_anahtar,
                                  anahtar=ayrik_anahtar, source_active=0, guncelleme=simdi())
                    del mevcut_harita[anahtar]
                    mevcut = None

                if mevcut is None:
                    # NO önceki bir yüklemede başka talebe verilip ayrıldıysa ve şimdi eski
                    # talebine döndüyse (ör. Talep NO yazım hatası düzeltildi) ayrılmış kart
                    # notları, operatörü ve iş akışıyla geri bağlanır; yeni kart açılmaz.
                    # Birden fazla aday varsa tahmin yapılmaz, eskisi gibi yeni kart açılır.
                    onek = f"{satir['source_row_id']}~"
                    adaylar = [k for k in _kartlar
                               if k.get("kaynak") == "EXCEL" and k.get("source_active", 1) == 0
                               and k.get("source_sheet") == satir["source_sheet"]
                               and str(k.get("source_row_id") or "").startswith(onek)
                               and _temiz_metin(k.get("talep_no")) == satir["talep_no"]]
                    if len(adaylar) == 1:
                        mevcut = adaylar[0]
                        geri_baglananlar.add(mevcut["id"])
                        _loglar.append(_log_kaydi(
                            kullanici, "admin", "EXCEL NO ESKİ TALEBE DÖNDÜ", mevcut["talep_no"], mevcut["stok_no"],
                            detay=(f"ID={mevcut['id']} {mevcut['source_row_id']} -> {anahtar}: ayrılmış kart "
                                   "notları ve iş akışıyla geri bağlandı"),
                        ))
                        mevcut.update(source_row_id=satir["source_row_id"], source_key=anahtar, anahtar=anahtar)
                        mevcut_harita[anahtar] = mevcut

                if mevcut:
                    onceki = copy.deepcopy(mevcut)
                    yeni_toplam = int(plan["toplam_adet"])
                    tamamlanan = int(mevcut.get("tamamlanan_adet") or 0)
                    yeni_durum = satir.get("ilk_durum")
                    onceki_durum = mevcut.get("durum")
                    etkin_durum = yeni_durum
                    gerileme = None
                    karar_turu = None
                    if (yeni_durum in DURUM_ILERLEME and onceki_durum in DURUM_ILERLEME
                            and DURUM_ILERLEME[yeni_durum] < DURUM_ILERLEME[onceki_durum]):
                        karar_turu = "geride"
                        secenekler = [durum for durum, sira in sorted(DURUM_ILERLEME.items(), key=lambda x: x[1])
                                      if DURUM_ILERLEME[yeni_durum] <= sira <= DURUM_ILERLEME[onceki_durum]]
                        uygulamada = _uygulamada_ilerletildi_mi(mevcut)
                    elif (yeni_durum is None and onceki_durum is not None
                          and (secimler is not None or sadece_onizleme)):
                        # Excel bu kart için bir iş akışı durumu söylemiyor (DURUM boş,
                        # MALZEME TEDARİK, PDGM ÖNERİ). Önizlemeli akışta durum sessizce
                        # korunmaz: admin kartı durumsuz bırakmakla uygulamadaki durumu
                        # korumak arasında seçer. Önizlemesiz doğrudan import eski
                        # sözleşmeyi sürdürür: karar yok, iş akışı olduğu gibi korunur.
                        karar_turu = "durumsuz"
                        secenekler = [None, onceki_durum]
                        uygulamada = (_uygulamada_ilerletildi_mi(mevcut)
                                      or onceki_durumu_coz(mevcut.get("excel_durum")) != onceki_durum)
                    if karar_turu:
                        onerilen = onceki_durum if uygulamada else yeni_durum
                        if secimler is None:
                            etkin_durum = onerilen if sadece_onizleme else yeni_durum
                        else:
                            etkin_durum = _durum_normalize(secimler.get(mevcut["id"]))
                            if mevcut["id"] not in secimler or etkin_durum not in secenekler:
                                raise IsKuralHatasi(
                                    f"Kart {mevcut['id']} ({mevcut.get('talep_no')}): "
                                    + ("Excel'in geride kaldığı" if karar_turu == "geride" else "Excel'de DURUM'u boş olan")
                                    + " kart için geçerli bir durum seçilmedi. Aktarım uygulanmadı; yeni önizleme oluşturun."
                                )
                        gerileme = {
                            "karar_turu": karar_turu,
                            "excel_durum": yeni_durum, "uygulama_durum": onceki_durum, "secenekler": secenekler,
                            "onerilen": onerilen, "secilen": etkin_durum, "uygulamada": uygulamada,
                            "onceki_excel_durum": mevcut.get("excel_durum"),
                            "onceki_tamamlanan": tamamlanan, "baslama_zamani": mevcut.get("baslama_zamani"),
                            "teslim_zamani": mevcut.get("teslim_zamani"),
                            "onceki_gerceklesen": mevcut.get("gerceklesen_teslim"), "operator": mevcut.get("operator"),
                        }
                        gerilemeler[mevcut["id"]] = gerileme
                        if DIZGIDE in secenekler and tamamlanan > 0:
                            sifirlanabilir_olasi.add(mevcut["id"])
                    # Karar olmadan durum gelmiyorsa (doğrudan import) iş akışı olduğu gibi
                    # korunur ve adet kuralları o iş akışına göre denetlenir. Admin kararıyla
                    # durumsuz kalan kartın adedi aşağıda 0'a iner, dolayısıyla denetim gerekmez.
                    is_akisi_korunur = etkin_durum is None and karar_turu != "durumsuz"
                    if (is_akisi_korunur or etkin_durum == DIZGIDE) and yeni_toplam < tamamlanan:
                        raise IsKuralHatasi(
                            f"{anahtar}: Excel toplam adedi ({yeni_toplam}), tamamlanan "
                            f"adetten ({tamamlanan}) küçük. Import iptal edildi."
                        )
                    if is_akisi_korunur and onceki_durum == TESLIM_EDILDI and yeni_toplam != tamamlanan:
                        raise IsKuralHatasi(f"{anahtar}: Teslim edilmiş kartın adedi için geçerli Excel DURUM gerekli.")

                    if not mevcut.get("source_key"):
                        mevcut["legacy_anahtar"] = mevcut["anahtar"]
                    mevcut.update({a: satir[a] for a in ("source_key", "source_sheet", "source_row_id")})
                    mevcut.update(plan)
                    mevcut.update(
                        anahtar=anahtar, talep_no=satir["talep_no"], stok_no=satir["stok_no"],
                        gerceklesen_teslim=satir.get("gerceklesen_teslim"),
                        source_active=1, aktif=1, kaynak="EXCEL", guncelleme=simdi(),
                    )
                    # Boş/geçersiz DURUM manuel workflow'u korusa da kaynak tarihi
                    # değiştiğinde eski zaman damgası onu geri getirmemeli.
                    if onceki.get("gerceklesen_teslim") != mevcut.get("gerceklesen_teslim"):
                        mevcut["teslim_zamani"] = None
                    if (etkin_durum == onceki_durum == TESLIM_EDILDI
                            and not satir.get("gerceklesen_teslim")
                            and (gerileme or onceki.get("teslim_zamani"))):
                        # Kart teslim edilmiş kalıyor ve Excel'in tarih hücresi boş: boş
                        # hücre "tarih bilinmiyor" demektir, bilinen tarihi silmez. İki yol:
                        # admin kararla uygulamadaki teslimi korudu, ya da Excel TESLİM'e
                        # yetişti ama tarihi yazılmadı ve teslim uygulamada kaydedilmişti
                        # (teslim_zamani yalnız Teslim Et / admin düzenlemesiyle oluşur).
                        # Excel bir tarih yazarsa o tarih geçerli olmaya devam eder.
                        mevcut.update(gerceklesen_teslim=onceki.get("gerceklesen_teslim"),
                                      teslim_zamani=onceki.get("teslim_zamani"))
                    if etkin_durum is None:
                        if karar_turu == "durumsuz":
                            # Admin kararıyla Excel'deki gibi durumsuz: operasyon ekranlarından
                            # kalkar, Durumu Eksik listesine düşer. Durumsuz kartta adet 0'dır.
                            mevcut.update(durum=None, tamamlanan_adet=0, baslangic_adet=0,
                                          baslama_zamani=None, bitis_zamani=None, teslim_zamani=None)
                            durum_iyilesen += 1
                        else:
                            workflow_korundu += 1
                    else:
                        if karar_turu == "durumsuz":
                            workflow_korundu += 1
                        mevcut["durum"] = etkin_durum
                        durum_iyilesen += int(onceki_durum != etkin_durum)
                        if etkin_durum in (PLANA_ALINDI, HAZIR):
                            mevcut.update(tamamlanan_adet=0, baslangic_adet=0,
                                          baslama_zamani=None, bitis_zamani=None, teslim_zamani=None)
                        elif etkin_durum == DIZGIDE:
                            mevcut["baslangic_adet"] = min(mevcut.get("baslangic_adet") or yeni_toplam, yeni_toplam)
                            if tamamlanan != yeni_toplam:
                                mevcut["bitis_zamani"] = None
                            mevcut["teslim_zamani"] = None
                        elif etkin_durum == TESLIM_EDILDI:
                            mevcut["tamamlanan_adet"] = yeni_toplam
                            mevcut["baslangic_adet"] = min(mevcut.get("baslangic_adet") or yeni_toplam, yeni_toplam)
                        # Excel import zamanı gerçek üretim/teslim zamanı değildir. Uygulamadaki
                        # durumu korunan durumsuz kartta durum Excel'den gelmez; operatör alanı kalır.
                        if karar_turu != "durumsuz":
                            mevcut["operator"] = mevcut.get("operator") or "Excel"
                    _kart_dogrula(mevcut)
                    # Durumu değişen (veya durum kararı istenen) kartın eski notları yeni
                    # duruma taşınmasın isteniyorsa admin önizlemede temizletir.
                    if _temiz_metin(onceki.get("aciklama")) and (gerileme or mevcut.get("durum") != onceki_durum):
                        not_temizlenebilir.add(mevcut["id"])
                        if mevcut["id"] in not_temizle:
                            mevcut["aciklama"] = None
                            notu_temizlenen += 1
                            _loglar.append(_log_kaydi(
                                kullanici, "admin", "EXCEL: NOTLAR TEMİZLENDİ", mevcut["talep_no"], mevcut["stok_no"],
                                detay=(f"ID={mevcut['id']} kaynak={anahtar} durum {onceki_durum or 'yok'} -> "
                                       f"{mevcut.get('durum') or 'yok'}. Silinen notlar:\n{onceki.get('aciklama')}"),
                            ))
                    degisiklik = {a: [onceki.get(a), mevcut.get(a)] for a in mevcut
                                 if a != "guncelleme" and onceki.get(a) != mevcut.get(a)}
                    # Güncellenen ve DİZGİDE kalan kartta operatörün girdiği tamamlanan adet
                    # korunur; admin önizlemede isterse sıfırlatabilir.
                    if ((degisiklik or gerileme) and mevcut.get("durum") == DIZGIDE
                            and _sayi(mevcut.get("tamamlanan_adet"), 0) > 0):
                        sifirlanabilir.add(mevcut["id"])
                        if mevcut["id"] in sifirla:
                            mevcut.update(tamamlanan_adet=0, bitis_zamani=None)
                            _kart_dogrula(mevcut)
                            degisiklik = {a: [onceki.get(a), mevcut.get(a)] for a in mevcut
                                         if a != "guncelleme" and onceki.get(a) != mevcut.get(a)}
                            sifirlanan += 1
                    if gerileme and karar_turu == "geride":
                        _loglar.append(_log_kaydi(
                            kullanici, "admin", "EXCEL GERİDE: DURUM KARARI", mevcut["talep_no"], mevcut["stok_no"],
                            detay=(f"ID={mevcut['id']} kaynak={anahtar}: uygulama {onceki_durum}, Excel {yeni_durum} -> "
                                   f"{etkin_durum}" + (" (uygulamadaki korundu)" if etkin_durum != yeni_durum else "")),
                        ))
                    elif gerileme:
                        excel_metni = _temiz_metin(plan.get("excel_durum"))
                        _loglar.append(_log_kaydi(
                            kullanici, "admin", "EXCEL DURUM BOŞ: DURUM KARARI", mevcut["talep_no"], mevcut["stok_no"],
                            detay=(f"ID={mevcut['id']} kaynak={anahtar}: uygulama {onceki_durum}, Excel DURUM "
                                   + (f"'{excel_metni}' (iş akışı durumu değil)" if excel_metni else "boş") + " -> "
                                   + (f"{etkin_durum} (uygulamadaki korundu)" if etkin_durum else "durumsuz bırakıldı")),
                        ))
                    if degisiklik:
                        guncellenen += 1
                        _loglar.append(_log_kaydi(
                            kullanici, "admin", "EXCEL KART GÜNCELLENDİ", mevcut["talep_no"], mevcut["stok_no"],
                            detay=f"ID={mevcut['id']} kaynak={anahtar} " + json.dumps(degisiklik, ensure_ascii=False),
                        ))
                    else:
                        # Aynı dosyanın tekrar yüklenmesi kart sürümünü değiştirmemeli;
                        # aksi halde açık admin formları gereksiz yere 409 alır.
                        mevcut["guncelleme"] = onceki.get("guncelleme")
                        degismeyen += 1
                    continue

                toplam = int(plan["toplam_adet"])
                durum = satir.get("ilk_durum")
                gerceklesen = satir.get("gerceklesen_teslim")

                tamamlanan = toplam if durum == TESLIM_EDILDI else 0
                baslangic_adet = toplam if durum in (DIZGIDE, TESLIM_EDILDI) else 0

                kayit = {
                    "id": sonraki_id,
                    "anahtar": anahtar,
                    "source_key": satir["source_key"],
                    "source_sheet": satir["source_sheet"],
                    "source_row_id": satir["source_row_id"],
                    "talep_no": satir["talep_no"],
                    "stok_no": satir["stok_no"],
                    "gerceklesen_teslim": gerceklesen,
                    "durum": durum,
                    "baslangic_adet": baslangic_adet,
                    "tamamlanan_adet": tamamlanan,
                    "baslama_zamani": None,
                    "bitis_zamani": None,
                    "teslim_zamani": None,
                    "operator": "Excel" if durum in GECERLI_DURUMLAR else None,
                    "aciklama": None,
                    "guncelleme": simdi(),
                    "aktif": 1,
                    "source_active": 1,
                    "admin_gizli": 0,
                    "kaynak": "EXCEL",
                    "legacy_anahtar": None,
                }
                kayit.update(plan)
                _kart_dogrula(kayit)
                _kartlar.append(kayit)
                mevcut_harita[anahtar] = kayit
                sonraki_id += 1
                yeni += 1

            if secimler:
                fazla = set(secimler) - set(gerilemeler)
                if fazla:
                    raise IsKuralHatasi(
                        "Durum kararı verilen kart bu aktarımda Excel'in gerisinde değil "
                        f"(ID: {', '.join(map(str, sorted(fazla)))}). Aktarım uygulanmadı; yeni önizleme oluşturun."
                    )
            gecersiz_not = not_temizle - not_temizlenebilir
            if gecersiz_not:
                raise IsKuralHatasi(
                    "Notları temizlenmek üzere seçilen kartın bu aktarımda durumu değişmiyor veya notu yok "
                    f"(ID: {', '.join(map(str, sorted(gecersiz_not)))}). "
                    "Aktarım uygulanmadı; dosyayı yeniden seçip yeni önizleme oluşturun."
                )
            gecersiz_secim = sifirla - sifirlanabilir
            if gecersiz_secim:
                raise IsKuralHatasi(
                    "Tamamlanan adedi sıfırlanmak üzere seçilen kart bu aktarımda güncellenmiyor veya "
                    f"DİZGİDE kalmıyor (ID: {', '.join(map(str, sorted(gecersiz_secim)))}). "
                    "Aktarım uygulanmadı; dosyayı yeniden seçip yeni önizleme oluşturun."
                )

            pasife_alinan = 0
            for kart in _kartlar:
                if kart.get("kaynak") != "EXCEL" or kart.get("anahtar") in gorulen:
                    continue
                if kart.get("source_active", 1) == 1:
                    kart["source_active"] = 0
                    kart["guncelleme"] = simdi()
                    pasife_alinan += 1
                    pasife_listesi.append(
                        f"ID={kart['id']} {kart.get('source_key') or kart['anahtar']}"
                    )
                    _loglar.append(_log_kaydi(
                        kullanici, "admin", "EXCEL KAYNAKTA YOK", kart.get("talep_no"), kart.get("stok_no"),
                        detay=f"ID={kart['id']} kaynak={kart.get('source_key') or kart['anahtar']}",
                    ))

            _kart_listesi_dogrula(_kartlar)

            if sadece_onizleme:
                oncekiler = {k["id"]: k for k in eski_kartlar}
                alanlar = [(baslik, alan) for baslik, alan in KART_ALANLARI if alan in {
                    "talep_no", "stok_no", "talep_sahibi", "toplam_adet", "tamamlanan_adet",
                    "durum", "plan_baslama", "plan_teslim", "gerceklesen_teslim", "source_active",
                    "pcb", "dizgi_tipi", "dizgi_sorumlusu", "malzeme_bekliyor", "plan_hafta", "excel_durum"}]
                degisiklikler = []
                for kart in _kartlar:
                    onceki = oncekiler.get(kart["id"], {})
                    farklar = [{"alan": baslik, "anahtar": alan, "eski": onceki.get(alan), "yeni": kart.get(alan)}
                               for baslik, alan in alanlar if onceki.get(alan) != kart.get(alan)]
                    if farklar or kart["id"] in gerilemeler:
                        # Önizleme ekranı kartları ne olacağına göre gruplar.
                        if not onceki:
                            tur = "yeni"
                        elif kart["id"] in ayrilanlar:
                            tur = "ayrildi"
                        elif kart["id"] in gerilemeler:
                            tur = "gerileme"
                        elif onceki.get("source_active", 1) == 1 and kart.get("source_active") == 0:
                            tur = "pasif"
                        elif onceki.get("source_active", 1) == 0 and kart.get("source_active") == 1:
                            tur = "geri"
                        else:
                            tur = "guncelleme"
                        degisiklikler.append({
                            "id": kart["id"], "talep_no": kart.get("talep_no"),
                            "stok_no": kart.get("stok_no"), "yeni": not onceki, "farklar": farklar,
                            "tur": tur, "sayfa": dizgi_kodu(kart.get("dizgi_tipi")),
                            "dizgi_etiket": dizgi_etiketi(kart.get("dizgi_tipi")),
                            **{alan: kart.get(alan) for alan in (
                                "sira", "talep_sahibi", "toplam_adet", "durum", "excel_durum",
                                "plan_baslama", "plan_teslim", "gerceklesen_teslim")},
                            "yerine": ayrilanlar.get(kart["id"]),
                            "geri_baglandi": kart["id"] in geri_baglananlar,
                            "tamamlanan_adet": kart.get("tamamlanan_adet"),
                            "sifirlanabilir": kart["id"] in sifirlanabilir,
                            "gerileme": gerilemeler.get(kart["id"]),
                            "aciklama": kart.get("aciklama"),
                            "not_temizlenebilir": kart["id"] in not_temizlenebilir,
                        })
                mevcut_aktif = sum(1 for k in eski_kartlar
                                   if k.get("kaynak") == "EXCEL" and k.get("source_active", 1) == 1)
                return {"surum": onceki_surum, "satir": len(satirlar), "yeni": yeni,
                        "guncellenen": guncellenen, "degismeyen": degismeyen,
                        "ayrilan": len(ayrilanlar), "geri_baglanan": len(geri_baglananlar),
                        "sifirlanabilir": sorted(sifirlanabilir | sifirlanabilir_olasi),
                        "not_temizlenebilir": sorted(not_temizlenebilir),
                        "gerileme": _karar_sayisi(gerilemeler, "geride"),
                        "durumsuz": _karar_sayisi(gerilemeler, "durumsuz"),
                        "pasife_alinan": pasife_alinan, "pasife_listesi": pasife_listesi,
                        "mevcut_aktif": mevcut_aktif,
                        "degisiklikler": degisiklikler, "uyari": uyari_sayisi}

            _yuklemeler.append(
                {
                    "zaman": simdi(),
                    "kullanici": kullanici,
                    "dosya": dosya_adi,
                    "satir": len(satirlar),
                    "yeni": yeni,
                    "guncellenen": guncellenen,
                    "pasife_alinan": pasife_alinan,
                    "uyari": uyari_sayisi,
                }
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "EXCEL YÜKLENDİ",
                    detay=(
                        f"{len(satirlar)} satır · {yeni} yeni · {guncellenen} güncellendi · "
                        f"{degismeyen} değişmedi · "
                        f"{workflow_korundu} workflow korundu · {durum_iyilesen} durum Excel'den değişti · "
                        f"{pasife_alinan} kaynakta yok · {len(ayrilanlar)} NO başka talebe verildi · "
                        f"{len(geri_baglananlar)} ayrılmış kart eski talebine geri bağlandı · "
                        f"{sifirlanan} kartın tamamlanan adedi admin seçimiyle sıfırlandı · "
                        f"{notu_temizlenen} kartın notları admin seçimiyle temizlendi · "
                        f"{_karar_sayisi(gerilemeler, 'geride')} kartta Excel geride "
                        f"({_karar_sayisi(gerilemeler, 'geride', korunan=True)} uygulamadaki korundu) · "
                        f"{_karar_sayisi(gerilemeler, 'durumsuz')} kartta Excel DURUM boş "
                        f"({_karar_sayisi(gerilemeler, 'durumsuz', korunan=True)} uygulamadaki korundu) · "
                        f"{uyari_sayisi} uyarı · yedek={os.path.basename(yedek_klasoru)}"
                        + (f" · {kaynak_ozeti}" if kaynak_ozeti else "")
                    ),
                )
            )

            _gunluk_yedek(KARTLAR_DOSYA)
            _coklu_yaz(
                [
                    (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
                    (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
                    (YUKLEME_DOSYA, YUKLEME_ALANLARI, _yuklemeler, "Yüklemeler"),
                ]
            )
            # Hata fırlatmaz; import yazıldıktan sonraki bakım adımıdır.
            _log_arsivle_gerekirse()

            return {
                "satir": len(satirlar),
                "yeni": yeni,
                "guncellenen": guncellenen,
                "degismeyen": degismeyen,
                "ayrilan": len(ayrilanlar),
                "geri_baglanan": len(geri_baglananlar),
                "sifirlanan": sifirlanan,
                "notu_temizlenen": notu_temizlenen,
                "gerileme": _karar_sayisi(gerilemeler, "geride"),
                "gerileme_korunan": _karar_sayisi(gerilemeler, "geride", korunan=True),
                "durumsuz": _karar_sayisi(gerilemeler, "durumsuz"),
                "durumsuz_korunan": _karar_sayisi(gerilemeler, "durumsuz", korunan=True),
                "pasife_alinan": pasife_alinan,
                "pasife_listesi": pasife_listesi[:20],
                "workflow_korundu": workflow_korundu,
                "durum_iyilesen": durum_iyilesen,
                "uyari": uyari_sayisi,
                "elle_dizgi_satir": elle_dizgi_satir,
                "eum_dizgi_satir": eum_dizgi_satir,
                "yedek": yedek_klasoru,
            }

        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            _yuklemeler = eski_yuklemeler
            raise


# ---------------------------------------------------------------------------
# Legacy küçük yardımcılar
# ---------------------------------------------------------------------------

def kart_guncelle(kart_id, **alanlar):
    global _kartlar

    kart_id = _sayi(kart_id, -1)

    with _kilit:
        index = next(
            (
                i
                for i, kart in enumerate(_kartlar)
                if kart.get("id") == kart_id
            ),
            None,
        )

        if index is None:
            return None

        eski = copy.deepcopy(_kartlar)

        try:
            yeni = copy.deepcopy(_kartlar[index])
            yeni.update(alanlar)
            yeni["guncelleme"] = simdi()

            if not yeni.get("source_key") and ("talep_no" in alanlar or "stok_no" in alanlar):
                yeni["anahtar"] = (
                    f"{_temiz_metin(yeni.get('talep_no'))}|"
                    f"{_temiz_metin(yeni.get('stok_no'))}"
                )

            yeni = _kart_normalize(yeni)
            _kartlar[index] = yeni

            _kart_listesi_dogrula(_kartlar)
            _kartlari_kaydet()

            return kart_gorunumu(yeni)

        except Exception:
            _kartlar = eski
            raise


def toplu_kaydet(yeni_kartlar=(), degisen=True):
    global _kartlar

    with _kilit:
        eski = copy.deepcopy(_kartlar)
        try:
            _kartlar.extend(_kart_normalize(kart) for kart in yeni_kartlar)
            _kart_listesi_dogrula(_kartlar)
            if degisen:
                _kartlari_kaydet()
        except Exception:
            _kartlar = eski
            raise
```

## `excel_araclari.py`

```python
"""PDGM kaynak Excel importu ve rapor üretimi.

Kaynak import kuralları:
- MAKİNE sayfası (zorunlu) + ELDE DİZGİ sayfası (opsiyonel) + EÜM sayfası
  (opsiyonel) kullanılır. Sayfalar TEK import işleminde birlikte işlenir:
  aksi halde bir sayfayı import etmek diğer sayfadan gelen kartları
  "kaynakta yok" (pasif) yapardı.
- Gizli satır ve sütunlar da okunur. Sütunlar başlık adıyla eşlenir; gizlilik
  yalnız aynı başlık birden fazla sütunda geçtiğinde görünür olanı seçmek için
  kullanılır (ör. gizli sipariş bloğundaki ikinci "Planlanan Teslim T.").
- Microsoft Excel COM ile values-only snapshot oluşturulur.
- External link güncellemesi ve full recalculation yapılmaz.
- Talep NO + Kart Stok No olmayan satırlar kart sayılmaz. Tamamen boş satırlar
  sessizce atlanır; yalnız şablon/kalıntı hücreleri (NO, Stok, sorumlu, hafta
  metni...) içeren satırlar atlanır ve önizlemede satır numarasıyla raporlanır.
  Adet, DURUM veya plan/teslim tarihi içeren ama Talep NO / Kart Stok No'su
  eksik satır gerçek kayıt sayılır ve tüm import reddedilir.
- Her satır kaynak sayfası + kalıcı NO (veya PDGM_ROW_ID) ile tanınır.
  Tarihler, adetler ve fiziksel satır sırası kimliğe katılmaz.
- Workflow yalnız şu dört durumdan oluşur:
    PLANA ALINDI, DİZGİDE, HAZIR, TESLİM EDİLDİ
  ELDE DİZGİ ve EÜM sayfalarındaki özel durum metinleri bu dört duruma
  indirgenir; ham metin excel_durum alanında saklanır ve kaybolmaz.
- Kartın MAKİNE / ELDE DİZGİ / EÜM sayfasından hangisinden geldiği
  "dizgi_tipi" alanında yapısal olarak (hangi sheet'ten okunduğuna göre)
  damgalanır; DURUM metninden tahmin edilmez.
- DURUM boş veya farklı bir değer ise kart kaybolmaz; durum None olur ve
  operasyon ekranında gösterilmez. Admin Yönetim ekranında düzeltebilir.
"""

from __future__ import annotations

import gc
import hashlib
import io
import json
import os
import re
import tempfile
import threading
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import from_excel, CALENDAR_WINDOWS_1900

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

import depo

_import_kilidi = threading.Lock()


class ExcelAktarimHatasi(Exception):
    # Satır bazlı sorunların tek tek listesi (önizleme ekranı madde madde gösterir).
    sorunlar = ()


KAYNAK_SAYFA_ADI = "MAKİNE"
KAYNAK_SAYFA_ADI_ELLE = "ELDE DİZGİ"
KAYNAK_SAYFA_ADI_EUM = "EÜM"
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# (sayfa adı, dizgi_tipi, zorunlu mu) — import bu listedeki sayfaları sırayla dener.
# MAKİNE zorunludur (yoksa hata); ELDE DİZGİ ve EÜM opsiyoneldir (eski/sade
# Excel dosyalarıyla geriye dönük uyumluluk için — yoksa sessizce atlanır).
KAYNAK_SAYFALARI = (
    (KAYNAK_SAYFA_ADI, depo.DIZGI_TIPI_MAKINE, True),
    (KAYNAK_SAYFA_ADI_ELLE, depo.DIZGI_TIPI_ELLE, False),
    (KAYNAK_SAYFA_ADI_EUM, depo.DIZGI_TIPI_EUM, False),
)


TURKCE_HARFLER = str.maketrans(
    {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
        "â": "a",
        "Â": "A",
    }
)


def _sadelestir(deger):
    metin = str(deger or "").translate(TURKCE_HARFLER)
    # Ayrışık yazılmış "I + birleşik nokta" gibi biçimler de aynı metne insin.
    metin = "".join(c for c in unicodedata.normalize("NFKD", metin) if not unicodedata.combining(c))
    metin = metin.upper()
    metin = metin.replace(".", " ").replace("_", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", metin).strip()


BASLIK_ESLESME = {
    _sadelestir(baslik): alan
    for baslik, alan in {
        "NO": "sira",
        "PDGM_ROW_ID": "source_row_id",
        "Talep NO": "talep_no",
        "Talep Sahibi": "talep_sahibi",
        "Kart Stok No": "stok_no",
        "Kart Üretim Adet": "adet_metin",
        "Planlanan Başlangıç T.": "plan_hafta",
        "T.planlanan tarih": "eum_planlanan_tarih",
        "Dizgi Başlama Tarihi": "plan_baslama",
        "Planlanan Teslim T.": "plan_teslim",
        "Gerçekleşen Teslim T.": "gerceklesen_teslim",
        "DURUM": "excel_durum",
        "PCB": "pcb",
        # Küçük format toleransı
        "Sıra": "sira",
        "Üretim Adet": "adet_metin",
        "Adet": "adet_metin",
        # ELDE DİZGİ YENİ sayfasına özel kolonlar / farklı adlandırmalar.
        # Aynı sözlük her iki sayfa için de kullanılır; bir sayfanın başlık
        # satırında bulunmayan bir eşleşme basitçe kolonlar sözlüğüne girmez.
        "PDGM Dizgi Sorumlusu": "dizgi_sorumlusu",
        "Dizgi Başlama T.": "plan_baslama",
        "MALZEME/PCB": "pcb",
    }.items()
}

DURUM_ESLESME = {
    "PLANA ALINDI": depo.PLANA_ALINDI,
    "DIZGIDE": depo.DIZGIDE,
    "HAZIR": depo.HAZIR,
    "TESLIM EDILDI": depo.TESLIM_EDILDI,
    # ELDE DİZGİ sayfasının beş durumu, dört gerçek workflow durumuna
    # indirgenir. Ham metin her zaman excel_durum alanında ayrıca saklanır,
    # bu yüzden bu indirgeme bilgi kaybı yaratmaz.
    "DIZGI ICIN BEKLIYOR": depo.PLANA_ALINDI,
    "DIZGIDE VE MALZEME BEKLIYOR": depo.DIZGIDE,
    # Gerçek ELDE DİZGİ sayfası aynı durumu "BEKLENİYOR" yazımıyla kullanıyor.
    "DIZGIDE VE MALZEME BEKLENIYOR": depo.DIZGIDE,
    # EÜM sayfası durumları. Planlanmış üretim hiçbir zaman DİZGİDE/TESLİM
    # sayılmaz: iki planlama yazımı da PLANA ALINDI'ya iner.
    "URETIM DEVAM EDIYOR": depo.DIZGIDE,
    "URETIM PLANA ALINDI": depo.PLANA_ALINDI,
    "EMTD PLANLANDI": depo.PLANA_ALINDI,
}

# Bu ham durumlar, kart bir workflow durumuna girmiş olsa bile fiilen
# malzeme eksikliğinden ilerleyemediğini ifade eder. Operatör ekranı bu
# bayrağa bakarak "Dizgiye Al" aksiyonunu gizler / uyarı gösterir.
MALZEME_BEKLEYEN_DURUMLAR = {
    "MALZEME TEDARIK",
    "DIZGIDE VE MALZEME BEKLIYOR",
    "DIZGIDE VE MALZEME BEKLENIYOR",
}

# Bilerek iş akışı durumuna eşlenmeyen kaynak durumları: kart durumsuz kalır
# (durumu olan mevcut kart için önizlemede admin kararı istenir). Bunların dışındaki tanınmayan her DURUM
# importu durdurur; aksi halde kart eski durumunda kalıp yanlış görünür
# (ör. Excel "TESLİM EDİLDİ (KISMİ)" derken monitörde "PLANDA").
DURUMSUZ_DURUMLAR = {"MALZEME TEDARIK", "PDGM ONERI"}
KABUL_EDILEN_DURUMLAR_METNI = (
    "PLANA ALINDI, DİZGİDE, HAZIR, TESLİM EDİLDİ, DİZGİ İÇİN BEKLİYOR, DİZGİDE VE MALZEME "
    "BEKLİYOR/BEKLENİYOR, ÜRETİM PLANA ALINDI, ÜRETİM DEVAM EDİYOR, EMTD PLANLANDI; durumsuz: "
    "MALZEME TEDARİK, PDGM ÖNERİ"
)

# pywin32 Range.Value2, hücre hatalarını (#N/A vb.) HRESULT tamsayısı olarak
# döndürür. Snapshot'a sayı olarak yazılırsa ör. Talep NO "-2146826246" olur.
EXCEL_HATA_KODLARI = {
    -2146826288: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826245: "#GETTING_DATA",
    -2146826243: "#SPILL!",
    -2146826238: "#CALC!",
}
EXCEL_HATA_METINLERI = set(EXCEL_HATA_KODLARI.values()) | {
    "#HATA!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#UNKNOWN!", "#BUSY!", "#PYTHON!",
}
XL_CELL_TYPE_VISIBLE = 12

# Talep NO / Kart Stok No eksik bir satırda bu alanlardan biri doluysa satır
# yarım kalmış gerçek bir kayıttır; sessizce atlanırsa kart "silinmiş" sayılırdı.
# NO, Stok, sorumlu, talep sahibi, PCB ve hafta metni tek başına kanıt değildir:
# gerçek dosyada gizli kalıntı/şablon satırları tam olarak bunları içeriyor.
KAYIT_KANITI_ALANLARI = (
    "talep_no",
    "adet_metin",
    "excel_durum",
    "plan_baslama",
    "plan_teslim",
    "gerceklesen_teslim",
    "eum_planlanan_tarih",
    "source_row_id",
)
ALAN_ADLARI = {
    "sira": "NO",
    "source_row_id": "PDGM_ROW_ID",
    "talep_no": "Talep NO",
    "talep_sahibi": "Talep Sahibi",
    "stok_no": "Kart Stok No",
    "adet_metin": "Kart Üretim Adet",
    "plan_hafta": "Planlanan Başlangıç T.",
    "eum_planlanan_tarih": "T.planlanan tarih",
    "plan_baslama": "Dizgi Başlama Tarihi",
    "plan_teslim": "Planlanan Teslim T.",
    "gerceklesen_teslim": "Gerçekleşen Teslim T.",
    "excel_durum": "DURUM",
    "pcb": "PCB",
    "dizgi_sorumlusu": "PDGM Dizgi Sorumlusu",
}


# ---------------------------------------------------------------------------
# COM snapshot
# ---------------------------------------------------------------------------

def _com_hata_kodu_mu(deger):
    # 0x800A07D0 (#NULL!) ve sonrası: Excel CVErr aralığı; gerçek veri bu değerleri almaz.
    return (isinstance(deger, int) and not isinstance(deger, bool)
            and -2146826288 <= deger <= -2146826200)


def _com_hucre(deger):
    if _com_hata_kodu_mu(deger):
        return "'" + EXCEL_HATA_KODLARI.get(deger, "#HATA!")
    # Value2 ile yazılan metni Excel yeniden yorumlar: "001234" -> 1234,
    # "1-2" -> tarih, "=abc" -> formül. Baştaki ' öneki metni aynen saklar.
    if isinstance(deger, str) and deger:
        return "'" + deger
    return deger


def _com_degerleri(deger):
    """Value2 sonucunu tuple-of-tuples yapar; metni ve hücre hatalarını metin olarak korur."""
    if not isinstance(deger, tuple):
        deger = ((deger,),)
    return tuple(tuple(_com_hucre(h) for h in satir) for satir in deger)


def _gizli_satir_bloklari(kaynak_ws, ilk_satir, son_satir):
    """Gizli satırları [(ilk, son), ...] blokları olarak döndürür (yalnız raporlama için)."""
    blok = kaynak_ws.Range(kaynak_ws.Cells(ilk_satir, 1), kaynak_ws.Cells(son_satir, 1)).EntireRow
    gorunur = set()
    try:
        alan = blok.SpecialCells(XL_CELL_TYPE_VISIBLE)
    except Exception:  # noqa: BLE001
        # Bilgi alınamadıysa "hepsi gizli" demek yanıltıcı olur; gizlilik raporlanmaz.
        return []
    for i in range(1, alan.Areas.Count + 1):
        parca = alan.Areas(i)
        gorunur.update(range(parca.Row, parca.Row + parca.Rows.Count))
    bloklar = []
    for satir in range(ilk_satir, son_satir + 1):
        if satir in gorunur:
            continue
        if bloklar and bloklar[-1][1] == satir - 1:
            bloklar[-1][1] = satir
        else:
            bloklar.append([satir, satir])
    return bloklar


def _sayfa_kopyala(kaynak_ws, hedef_ws):
    """Kaynak sayfanın kullanılan alanını değer olarak aynı satır/sütun konumlarına kopyalar.

    Gizli satır ve sütunlar da kopyalanır ve hedefte yine gizli işaretlenir.
    Böylece parser sütunları gerçek başlıklarla eşler: gizlenen DURUM veya
    teslim sütunu verinin boş sanılmasına yol açmaz; gizlilik yalnız aynı
    başlığın birden fazla sütunda geçtiği durumda görünür olanı seçmek için
    kullanılır.
    """
    used = kaynak_ws.UsedRange
    ilk_satir = used.Row
    son_satir = used.Row + used.Rows.Count - 1
    ilk_sutun = used.Column
    son_sutun = used.Column + used.Columns.Count - 1
    gizli_sutunlar = []

    for sutun in range(ilk_sutun, son_sutun + 1):
        kimlik_sutunu = any(
            _sadelestir(kaynak_ws.Cells(r, sutun).Value2) in {"NO", "PDGM ROW ID"}
            for r in range(ilk_satir, min(son_satir, 10) + 1)
        )
        kaynak_aralik = kaynak_ws.Range(
            kaynak_ws.Cells(ilk_satir, sutun),
            kaynak_ws.Cells(son_satir, sutun),
        )
        hedef_aralik = hedef_ws.Range(
            hedef_ws.Cells(ilk_satir, sutun),
            hedef_ws.Cells(son_satir, sutun),
        )
        # General biçimi metinsel '001' ve uzun sayısal kimlikleri dönüştürür.
        # Değerleri yazmadan önce yalnız kimlik sütununun hedefini metin yap.
        if kimlik_sutunu:
            hedef_aralik.NumberFormat = "@"
        hedef_aralik.Value2 = _com_degerleri(kaynak_aralik.Value2)
        if bool(kaynak_ws.Columns(sutun).Hidden):
            gizli_sutunlar.append(sutun)

    for sutun in gizli_sutunlar:
        hedef_ws.Columns(sutun).Hidden = True
    try:
        for bas, son in _gizli_satir_bloklari(kaynak_ws, ilk_satir, son_satir):
            hedef_ws.Rows(f"{bas}:{son}").Hidden = True
    except Exception:  # noqa: BLE001 - satır gizliliği yalnız rapor bilgisidir
        pass


def excel_deger_snapshot_olustur(dosya_yolu):
    """MAKİNE (zorunlu), ELDE DİZGİ ve EÜM (varsa) sayfalarını values-only geçici xlsx'e kopyalar.

    Her sayfa tek snapshot dosyasında, kaynak adlarıyla aynı isimde ayrı sheet
    olarak ve aynı hücre konumlarında yer alır; gizli satır/sütun bilgisi
    korunur. Böylece _excelden_aktar tek workbook üzerinden hepsini okuyup
    tek import işleminde birleştirebilir.
    """
    if pythoncom is None or win32com is None:
        raise ExcelAktarimHatasi(
            "Excel COM aktarımı yalnızca Windows + Microsoft Excel ortamında çalışır "
            "(pywin32 gerekli)."
        )

    kaynak = Path(dosya_yolu).resolve()
    temp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    temp.close()
    hedef = Path(temp.name).resolve()

    pythoncom.CoInitialize()
    excel = None
    kaynak_wb = None
    hedef_wb = None
    basarili = False

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        kaynak_wb = excel.Workbooks.Open(
            str(kaynak),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        def _sayfa_bul(sayfa_adi):
            return next(
                (
                    ws
                    for ws in kaynak_wb.Worksheets
                    if _sadelestir(ws.Name) == _sadelestir(sayfa_adi)
                ),
                None,
            )

        hedef_wb = excel.Workbooks.Add()
        hedef_wb.Date1904 = kaynak_wb.Date1904
        while hedef_wb.Worksheets.Count > 1:
            hedef_wb.Worksheets(hedef_wb.Worksheets.Count).Delete()

        ilk_sayfa_kullanildi = False
        for sayfa_adi, _dizgi_tipi, zorunlu in KAYNAK_SAYFALARI:
            kaynak_ws = _sayfa_bul(sayfa_adi)
            if kaynak_ws is None:
                if zorunlu:
                    sayfalar = [ws.Name for ws in kaynak_wb.Worksheets]
                    raise ExcelAktarimHatasi(
                        f"'{sayfa_adi}' sayfası bulunamadı. "
                        f"Dosyadaki sayfalar: {', '.join(sayfalar)}"
                    )
                continue  # opsiyonel sayfa yoksa sessizce atla

            if not ilk_sayfa_kullanildi:
                hedef_ws = hedef_wb.Worksheets(1)
                hedef_ws.Name = sayfa_adi
                ilk_sayfa_kullanildi = True
            else:
                hedef_ws = hedef_wb.Worksheets.Add(
                    After=hedef_wb.Worksheets(hedef_wb.Worksheets.Count)
                )
                hedef_ws.Name = sayfa_adi

            _sayfa_kopyala(kaynak_ws, hedef_ws)

        hedef_wb.SaveAs(str(hedef), FileFormat=51)
        basarili = True
        return str(hedef)

    finally:
        if hedef_wb is not None:
            try:
                hedef_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if kaynak_wb is not None:
            try:
                kaynak_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        kaynak_wb = None
        hedef_wb = None
        excel = None
        gc.collect()

        pythoncom.CoUninitialize()

        if not basarili:
            try:
                os.remove(str(hedef))
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Parser helpers
# ---------------------------------------------------------------------------

def durum_coz(deger):
    """Kaynak durumunu dört gerçek workflow durumundan birine çevirir."""
    temiz = _sadelestir(deger)
    if not temiz:
        return None, True
    durum = DURUM_ESLESME.get(temiz)
    return durum, durum is None


def malzeme_bekliyor_mu(deger):
    """Ham DURUM metni, kartın malzeme eksikliğinden ilerleyemediğini mi söylüyor?"""
    return _sadelestir(deger) in MALZEME_BEKLEYEN_DURUMLAR


def adet_coz(deger):
    """Üretim adedini pozitif tam sayıya çevirir.

    Kabul:
    - 400
    - 400.0
    - "400 ADET"
    - "1.500 ADET"
    - "1,500 ADET"
    - "1 500 ADET"

    Red:
    - 400.5
    - "400.5 ADET"
    - "400,5 ADET"
    - "-5 ADET"
    - "400 / 500 ADET"
    """
    if deger is None or str(deger).strip() == "":
        raise ExcelAktarimHatasi("Üretim adedi boş olamaz.")

    if isinstance(deger, bool):
        raise ExcelAktarimHatasi("Üretim adedi sayı olmalı.")

    if isinstance(deger, (int, float)):
        if isinstance(deger, float) and not deger.is_integer():
            raise ExcelAktarimHatasi(
                f"Üretim adedi tam sayı olmalı: {deger}"
            )
        adet = int(deger)

    else:
        metin = str(deger).strip()

        eslesmeler = [
            eslesme.strip()
            for eslesme in re.findall(
                r"[-+]?\d[\d.,\s]*",
                metin,
            )
            if eslesme.strip()
        ]

        if len(eslesmeler) != 1:
            raise ExcelAktarimHatasi(
                f"Üretim adedi tek bir sayı içermeli: '{metin}'"
            )

        token = re.sub(r"\s+", "", eslesmeler[0])

        if token.startswith(("+", "-")):
            raise ExcelAktarimHatasi(
                f"Üretim adedi pozitif tam sayı olmalı: '{metin}'"
            )

        if token.isdigit():
            adet = int(token)

        elif re.fullmatch(
            r"\d{1,3}(?:[.,]\d{3})+",
            token,
        ):
            adet = int(re.sub(r"[.,]", "", token))

        else:
            raise ExcelAktarimHatasi(
                f"Üretim adedi tam sayı olmalı: '{metin}'"
            )

    if adet < 1:
        raise ExcelAktarimHatasi(
            f"Üretim adedi en az 1 olmalı: {deger!r}"
        )

    return adet


def _baslik_satiri_bul(ws):
    for satir_no, satir in enumerate(
        ws.iter_rows(min_row=1, max_row=10, values_only=True),
        start=1,
    ):
        temiz = [_sadelestir(hucre) for hucre in satir]
        if "TALEP NO" in temiz or "KART STOK NO" in temiz:
            return satir_no, temiz
    return None, None


def _bos_mu(deger):
    """Boş hücre: None, "" veya yalnız boşluk (formülün "" sonucu dahil)."""
    return deger is None or (isinstance(deger, str) and not deger.strip())


def _excel_hatasi_mi(deger):
    return isinstance(deger, str) and deger.strip().upper() in EXCEL_HATA_METINLERI


def _hucre_al(satir, kolonlar, alan):
    index = kolonlar.get(alan)
    if index is None or index >= len(satir):
        return None
    deger = satir[index]
    return None if _bos_mu(deger) else deger


def _gizli_sutunlar(ws):
    """1 tabanlı gizli sütun numaraları. openpyxl bitişik sütunları tek boyutta (min..max) tutar."""
    gizli = set()
    for boyut in (getattr(ws, "column_dimensions", None) or {}).values():
        if not boyut.hidden:
            continue
        bas = boyut.min or column_index_from_string(boyut.index)
        gizli.update(range(bas, (boyut.max or bas) + 1))
    return gizli


def _satir_gizli_mi(ws, satir_no):
    boyutlar = getattr(ws, "row_dimensions", None)
    boyut = boyutlar.get(satir_no) if boyutlar is not None else None
    return bool(boyut is not None and boyut.hidden)


def satir_araliklari(satirlar):
    """[211, 212, 213, 229] -> '211–213, 229'"""
    parcalar = []
    for satir in sorted(satirlar):
        if parcalar and parcalar[-1][1] == satir - 1:
            parcalar[-1][1] = satir
        else:
            parcalar.append([satir, satir])
    return ", ".join(str(a) if a == b else f"{a}–{b}" for a, b in parcalar)


# Üretim planı/teslim sütunları DURUM'un yanındaki üretim bloğundadır; aynı başlıklar
# soldaki sipariş bloğunda (gizli) tekrar eder. Tekrarda doğru sütunu gizlilik değil
# DURUM'a yakınlık belirler: sütunu gizlemek/göstermek okunan tarihi değiştirmemeli.
URETIM_BLOGU_ALANLARI = {"plan_hafta", "eum_planlanan_tarih", "plan_baslama", "plan_teslim", "gerceklesen_teslim"}


def _kolonlari_esle(ws, sayfa_adi, basliklar):
    """Başlık adına göre alan -> 0 tabanlı sütun indeksi.

    Aynı alan birden fazla sütunda geçiyorsa:
    - üretim bloğu alanlarında (plan/teslim tarihleri) DURUM sütununa en yakın
      sütun seçilir; gerçek dosyada gizli sipariş bloğunda ikinci bir
      "Planlanan Teslim T." var ve hangisinin gizli olduğu sonucu değiştirmez;
    - diğer alanlarda (ör. NO) yalnız biri görünürse o seçilir.
    Karar verilemezse tahmin yerine hata verilir.
    "__kimlik_no__" yalnız başlığı tam olarak NO olan kimlik sütunudur.
    """
    gizli = _gizli_sutunlar(ws)
    adaylar = {}
    for index, baslik in enumerate(basliklar):
        alan = BASLIK_ESLESME.get(baslik)
        if alan:
            adaylar.setdefault(alan, []).append(index)
        if baslik == "NO":
            adaylar.setdefault("__kimlik_no__", []).append(index)
    durum_sutunlari = adaylar.get("excel_durum", [])
    durum_sutunu = durum_sutunlari[0] if len(durum_sutunlari) == 1 else None

    kolonlar = {}
    for alan, indexler in adaylar.items():
        if len(indexler) > 1:
            secilen = None
            if alan in URETIM_BLOGU_ALANLARI and durum_sutunu is not None:
                uzaklik = sorted((abs(i - durum_sutunu), i) for i in indexler)
                if uzaklik[0][0] != uzaklik[1][0]:
                    secilen = uzaklik[0][1]
            if secilen is None:
                gorunur = [i for i in indexler if i + 1 not in gizli]
                if len(gorunur) == 1:
                    secilen = gorunur[0]
            if secilen is None:
                harfler = ", ".join(get_column_letter(i + 1) for i in indexler)
                ad = "NO" if alan == "__kimlik_no__" else ALAN_ADLARI.get(alan, alan)
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' sayfası: '{ad}' için birden fazla sütun bulundu ({harfler}); hangisinin "
                    "okunacağı belirsiz. Okunması gereken sütunu görünür, diğerlerini gizli bırakın "
                    "veya başlığını değiştirin."
                )
            indexler = [secilen]
        kolonlar[alan] = indexler[0]
    return kolonlar


def _tarih_coz_ve_dogrula(deger, alan_adi, excel_satir_no, epoch=CALENDAR_WINDOWS_1900):
    # Value2 sayısal tarihleri workbook'un tarih sistemine göre yorumlanır.
    sonuc = None
    sayisal = isinstance(deger, (int, float)) and not isinstance(deger, bool)
    if isinstance(deger, str) and re.fullmatch(r"\d+(?:\.\d+)?", deger.strip()):
        sayisal = True
    if sayisal:
        try:
            tarih = from_excel(float(deger), epoch=epoch)
            if isinstance(tarih, (date, datetime)):
                sonuc = tarih.strftime("%Y-%m-%d")
        except (ValueError, TypeError, OverflowError):
            pass
    else:
        sonuc = depo.tarih_coz(deger)
    if deger not in (None, "") and not sonuc:
        raise ExcelAktarimHatasi(
            f"Excel satır {excel_satir_no}: {alan_adi} okunamadı. Gelen değer: {deger!r}"
        )
    return sonuc


def _sira_coz(deger):
    if deger in (None, ""):
        return None, False
    try:
        return int(float(deger)), False
    except (TypeError, ValueError):
        return None, True


def _kaynak_satir_kimligi(no, acik_id, sayfa_adi, satir_no):
    deger = acik_id if acik_id not in (None, "") else no
    if isinstance(deger, bool) or deger in (None, ""):
        raise ExcelAktarimHatasi(
            f"'{sayfa_adi}' satır {satir_no}: Kalıcı NO veya PDGM_ROW_ID gerekli. "
            "Her satıra benzersiz ve sonraki yüklemelerde değişmeyen bir kimlik verin."
        )
    if isinstance(deger, float) and deger.is_integer():
        deger = int(deger)
    metin = str(deger).strip()
    if not metin:
        raise ExcelAktarimHatasi(f"'{sayfa_adi}' satır {satir_no}: Kaynak kimliği boş.")
    return ("ID:" if acik_id not in (None, "") else "NO:") + metin


# Tam tarih gibi görünen metinler (17.09.2026, 2026-09-21T00:30, 46147) katı
# tarih olarak çözülür; "2.06" gibi yılsız gün.ay hafta metni sayılır.
TARIH_BENZERI = re.compile(
    r"^\s*(\d{4,6}(?:\.\d+)?|\d{1,4}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{1,4}(?:[T\s]\d{1,2}:\d{2}\S*)?)\s*$"
)
# Önünde rakam/nokta olmamalı: "2.06 haftası" -> 6. hafta, "2026 haftası" -> 26. hafta değil.
HAFTA_NO_DESENI = re.compile(r"(?<![\d./])(\d{1,2})\s*\.?\s*HAFTA")
GUN_AY_DESENI = re.compile(r"(\d{1,2})\s*[./]\s*(\d{1,2})(?:\s*[./]\s*(\d{2,4}))?")
BOS_TARIH_METINLERI = {"-", "YOK", "N/A", "NONE"}


def _pazartesi(gun):
    return gun - timedelta(days=gun.weekday())


def _hafta_metni_coz(metin, referans_iso=None):
    """ "23. Hafta (2.06 haftası)" -> ('2025-06-02', uyarı veya None).

    Gerçek EÜM sayfası planlanan başlangıcı yılsız hafta metni olarak tutar.
    Yıl, satırın plan/teslim tarihine (yoksa bugüne) en yakın ve hafta numarası
    ile gün.ay bilgisinin tutarlı olduğu yıldır; tarih haftanın Pazartesi'sidir.
    """
    referans = date.fromisoformat(referans_iso) if referans_iso else date.today()
    temiz = str(metin).translate(TURKCE_HARFLER).upper()
    eslesme = HAFTA_NO_DESENI.search(temiz)
    hafta = int(eslesme.group(1)) if eslesme and 1 <= int(eslesme.group(1)) <= 53 else None
    kalan = temiz[eslesme.end():] if eslesme else temiz

    def en_yakin(adaylar):
        return min(adaylar, key=lambda gun: (abs((gun - referans).days), gun))

    gun_ay = GUN_AY_DESENI.search(kalan)
    if gun_ay:
        gun, ay, yil = int(gun_ay.group(1)), int(gun_ay.group(2)), gun_ay.group(3)
        if yil:
            yillar = [int(yil) + 2000 if int(yil) < 100 else int(yil)]
        else:
            yillar = [referans.year - 1, referans.year, referans.year + 1]
        adaylar = []
        for aday_yil in yillar:
            try:
                adaylar.append(date(aday_yil, ay, gun))
            except ValueError:
                pass
        if adaylar:
            uyumlu = [a for a in adaylar if hafta is None or a.isocalendar()[1] == hafta]
            # "(2.06 haftası)" haftanın Pazartesi'sidir; yıl belirsizliğini bu çözer.
            secim = en_yakin([a for a in uyumlu if a.weekday() == 0] or uyumlu
                             or [a for a in adaylar if a.weekday() == 0] or adaylar)
            uyari = None
            if hafta is not None and not uyumlu:
                uyari = f"hafta numarası ({hafta}) ile {secim:%d.%m.%Y} uyuşmuyor; gün.ay esas alındı"
            return _pazartesi(secim).isoformat(), uyari

    if hafta is not None:
        adaylar = []
        for aday_yil in (referans.year - 1, referans.year, referans.year + 1):
            try:
                adaylar.append(date.fromisocalendar(aday_yil, hafta, 1))
            except ValueError:
                pass
        if adaylar:
            uyari = None if referans_iso else "yıl bilgisi yok; bugüne en yakın yıl seçildi"
            return en_yakin(adaylar).isoformat(), uyari
    return None, None


def _planlanan_baslangic_coz(deger, alan_adi, satir_no, epoch, referans_iso, kati):
    """EÜM plan başlangıcı: tarih ise o haftanın Pazartesi'si, hafta metni ise çözülmüş Pazartesi.

    Döner (iso veya None, uyarı metni veya None). Tarih gibi görünüp geçersiz
    değer her zaman hata verir; tanınmayan serbest metin katı (açık tarih)
    sütununda hata, hafta metni sütununda uyarıdır.
    """
    if _bos_mu(deger) or (isinstance(deger, str) and deger.strip().upper() in BOS_TARIH_METINLERI):
        return None, None
    if not isinstance(deger, str) or TARIH_BENZERI.match(deger):
        tarih = _tarih_coz_ve_dogrula(deger, alan_adi, satir_no, epoch)
        return _pazartesi(date.fromisoformat(tarih)).isoformat(), None
    sonuc, uyari = _hafta_metni_coz(deger, referans_iso)
    if sonuc:
        return sonuc, (f"{alan_adi} '{deger.strip()}': {uyari}." if uyari else None)
    if kati:
        raise ExcelAktarimHatasi(
            f"Excel satır {satir_no}: {alan_adi} okunamadı. Gelen değer: {deger!r}"
        )
    return None, (f"{alan_adi} '{deger.strip()}' tarih veya 'NN. hafta' olarak okunamadı; "
                  "plan başlangıcı boş bırakıldı.")


def _hata_ozeti(hatalar, sinir=10):
    if len(hatalar) == 1:
        return hatalar[0]
    ozet = f"{len(hatalar)} sorun bulundu, hiçbir kayıt değiştirilmedi: " + " | ".join(hatalar[:sinir])
    if len(hatalar) > sinir:
        ozet += f" | … ve {len(hatalar) - sinir} sorun daha"
    return ozet


def _sorun_hatasi(hatalar):
    hata = ExcelAktarimHatasi(_hata_ozeti(hatalar))
    hata.sorunlar = tuple(hatalar)
    return hata


# Önizlemede uyarılar türüne göre gruplanır; açıklama kartlara ne olacağını söyler.
UYARI_TURLERI = {
    "durum_bos": (
        "DURUM boş",
        "Yeni kart durumsuz oluşturulur ve Pano, Operatör ve Monitör ekranlarında görünmez; "
        "Yönetim ekranındaki \"Durumu eksik kartlar\" bölümünden durum atanabilir. Kart sistemde "
        "bir durumla duruyorsa \"Karar gerekiyor\" bölümünde durumsuz bırakmak ile mevcut durumu "
        "korumak arasında seçim yaparsınız.",
    ),
    "durum_durumsuz": (
        "DURUM \"{deger}\": iş akışı durumu atanmaz",
        "Bu aşama (malzeme tedariği, PDGM önerisi) bilerek PLANA ALINDI / DİZGİDE / TESLİM EDİLDİ'ye "
        "eşlenmez; sonuç DURUM boş ile aynıdır (durumu olan kart için karar istenir). Excel'deki "
        "metin kartta kaynak durumu olarak saklanır.",
    ),
    "teslim_tarihsiz": (
        "TESLİM EDİLDİ ama Gerçekleşen Teslim T. boş",
        "Kart teslim edilmiş sayılır. Kart uygulamada teslim edildiyse uygulamadaki teslim "
        "tarihi korunur. Değilse tarih bilinmiyor olarak kalır, zamanında/geç teslim hesabına "
        "girmez; sistem tarih uydurmaz.",
    ),
    "teslim_tarihi_durum_disi": (
        "Gerçekleşen Teslim T. dolu ama DURUM teslim değil",
        "DURUM esas alınır: kart teslim edilmiş sayılmaz. Tarih kartta bilgi olarak saklanır.",
    ),
    "no_sayisal_degil": (
        "NO sayısal değil",
        "Kart kimliği NO metniyle tutulmaya devam eder; yalnız sıralama alanı boş kalır.",
    ),
    "plan_baslangic": (
        "Planlanan başlangıç tam çözülemedi",
        "EÜM hafta metni tarih veya \"NN. hafta\" olarak okunamadı ya da hafta numarasıyla tarih "
        "uyuşmadı. Satır ayrıntısında kullanılan değer yazıyor.",
    ),
}


def _uyari_gruplari(sayfalar):
    """Uyarıları (tür, değer) bazında toplar: kaç satır, hangi sayfa/satırlar."""
    gruplar = {}
    for sayfa in sayfalar:
        for kayit in sayfa.get("uyari_kayitlari", []):
            anahtar = (kayit["tur"], kayit.get("deger"))
            grup = gruplar.setdefault(anahtar, {"satirlar": {}, "mesajlar": []})
            grup["satirlar"].setdefault(sayfa["sayfa"], []).append(kayit["satir"])
            grup["mesajlar"].append(kayit["mesaj"])
    sonuc = []
    for (tur, deger), grup in gruplar.items():
        baslik, aciklama = UYARI_TURLERI.get(tur, (tur, ""))
        sonuc.append({
            "tur": tur,
            "baslik": baslik.format(deger=deger or ""),
            "aciklama": aciklama,
            "adet": len(grup["mesajlar"]),
            "yerler": [{"sayfa": ad, "satirlar": satir_araliklari(s)} for ad, s in grup["satirlar"].items()],
            "mesajlar": grup["mesajlar"],
        })
    return sonuc


def _hucre_etiketi(dolu):
    """'I:Kart Stok No' -> 'Kart Stok No (I)'; başlıksız sütun -> 'V sütunu'."""
    harf, _, baslik = dolu.partition(":")
    return f"{baslik} ({harf})" if baslik else f"{harf} sütunu"


def kaynak_raporu_ozeti(rapor):
    """Log ve bildirim için tek satırlık sayfa özeti."""
    parcalar = []
    for sayfa in (rapor or {}).get("sayfalar", []):
        if sayfa.get("bulunamadi"):
            parcalar.append(f"{sayfa['sayfa']}: sayfa yok")
            continue
        metin = f"{sayfa['sayfa']}: {sayfa['kayit']} kayıt"
        if sayfa.get("atlanan"):
            metin += (f", {len(sayfa['atlanan'])} satır Talep NO/Kart Stok No olmadığı için "
                      f"kart sayılmadı ({sayfa['atlanan_araliklar']})")
        parcalar.append(metin)
    return "; ".join(parcalar)


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

def _sayfa_satirlarini_coz(ws, sayfa_adi, dizgi_tipi, anahtar_gruplari, parsed_liste,
                           rapor=None, hatalar=None):
    """Kaynak satırlarını doğrular; sayfa kapsamında kalıcı kimlik üretir.

    Satır sınıfları:
    - Tamamen boş (formülün "" sonucu dahil): sayılır, sessizce atlanır.
    - Talep NO + Kart Stok No dolu: kayıt; kimlik, adet ve tarihler doğrulanır.
    - Talep NO / Kart Stok No eksik ama adet, DURUM, plan/teslim tarihi,
      PDGM_ROW_ID veya Talep NO içeren satır: yarım gerçek kayıt -> hata.
    - Diğerleri (yalnız NO, Stok, sorumlu, hafta metni gibi kalıntılar): kart
      sayılmaz, satır numarasıyla rapora yazılır.
    Satırın gizli olması sınıfı değiştirmez.

    rapor verilirse sayfa özeti eklenir. hatalar listesi verilirse satır
    hataları oraya toplanır; verilmezse sayfa sonunda tek ExcelAktarimHatasi
    fırlatılır. Hatalı satır hiçbir durumda parsed_liste'ye eklenmez.
    """
    kendi_hatalari = hatalar is None
    if kendi_hatalari:
        hatalar = []
    sayfa_raporu = {
        "sayfa": sayfa_adi,
        "kod": depo.dizgi_kodu(dizgi_tipi),
        "bulunamadi": False,
        "baslik_satiri": None,
        "kayit": 0,
        "gizli_kayit": 0,
        "bos_satir": 0,
        "atlanan": [],
        "atlanan_araliklar": "",
        "atlanan_gruplari": [],
        "uyarilar": [],
        "uyari_kayitlari": [],
    }
    if rapor is not None:
        rapor.setdefault("sayfalar", []).append(sayfa_raporu)
    uyari_sayisi = 0

    def bitir():
        atlanan = sayfa_raporu["atlanan"]
        sayfa_raporu["atlanan_araliklar"] = satir_araliklari(a["satir"] for a in atlanan)
        # Aynı hücreleri dolu kalıntı satırları tek satırda özetlenir.
        gruplar = {}
        for a in atlanan:
            gruplar.setdefault(tuple(a["dolu"]), []).append(a)
        sayfa_raporu["atlanan_gruplari"] = [
            {"hucreler": [_hucre_etiketi(d) for d in dolu], "adet": len(satirlar),
             "gizli": sum(1 for a in satirlar if a["gizli"]),
             "satirlar": satir_araliklari(a["satir"] for a in satirlar)}
            for dolu, satirlar in gruplar.items()
        ]
        if kendi_hatalari and hatalar:
            raise _sorun_hatasi(hatalar)
        return uyari_sayisi

    try:
        baslik_no, basliklar = _baslik_satiri_bul(ws)
        if not baslik_no:
            raise ExcelAktarimHatasi(
                f"'{sayfa_adi}' sayfasında başlık satırı bulunamadı. "
                "Dosyada 'Talep NO' veya 'Kart Stok No' sütunu olmalı."
            )
        kolonlar = _kolonlari_esle(ws, sayfa_adi, basliklar)
        for gerekli in ("talep_no", "stok_no", "adet_metin"):
            if gerekli not in kolonlar:
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' sayfası: Zorunlu '{ALAN_ADLARI[gerekli]}' sütunu bulunamadı."
                )
    except ExcelAktarimHatasi as exc:
        hatalar.append(str(exc))
        return bitir()

    sayfa_raporu["baslik_satiri"] = baslik_no
    ham_basliklar = next(ws.iter_rows(min_row=baslik_no, max_row=baslik_no, values_only=True))
    epoch = ws.parent.epoch

    def sutun_adi(index):
        baslik = ham_basliklar[index] if index < len(ham_basliklar) else None
        harf = get_column_letter(index + 1)
        return f"{harf}:{str(baslik).strip()}" if not _bos_mu(baslik) else harf

    for excel_satir_no, satir in enumerate(
        ws.iter_rows(min_row=baslik_no + 1, values_only=True),
        start=baslik_no + 1,
    ):
        dolu = [i for i, hucre in enumerate(satir) if not _bos_mu(hucre)]
        if not dolu:
            sayfa_raporu["bos_satir"] += 1
            continue
        gizli = _satir_gizli_mi(ws, excel_satir_no)

        def deger(alan):
            ham = _hucre_al(satir, kolonlar, alan)
            return None if _excel_hatasi_mi(ham) else ham

        talep_no = str(deger("talep_no") or "").strip()
        stok_no = str(deger("stok_no") or "").strip()
        if not talep_no or not stok_no:
            kanit = [ALAN_ADLARI[a] for a in KAYIT_KANITI_ALANLARI if deger(a) is not None]
            if kanit:
                eksik = " ve ".join(ad for ad, v in (("Talep NO", talep_no), ("Kart Stok No", stok_no)) if not v)
                hatalar.append(
                    f"'{sayfa_adi}' satır {excel_satir_no}{' (gizli)' if gizli else ''}: {eksik} boş, "
                    f"ancak satırda kayıt verisi var ({', '.join(kanit)}). Eksik alanı doldurun; "
                    "kayıt kaldırıldıysa satırı tamamen temizleyin. Eksik satır silinmiş kabul edilmedi."
                )
            else:
                sayfa_raporu["atlanan"].append({
                    "satir": excel_satir_no,
                    "gizli": gizli,
                    "dolu": [sutun_adi(i) for i in dolu],
                })
            continue

        satir_uyarilari = []

        def uyar(tur, mesaj, ek=None):
            satir_uyarilari.append((tur, ek, f"'{sayfa_adi}' satır {excel_satir_no}: {mesaj}"))

        try:
            hatali_hucreler = [
                ALAN_ADLARI.get(alan, alan) for alan, index in kolonlar.items()
                if not alan.startswith("__") and index < len(satir) and _excel_hatasi_mi(satir[index])
            ]
            if hatali_hucreler:
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' satır {excel_satir_no}: Excel hata değeri içeren hücre "
                    f"({', '.join(hatali_hucreler)}). Formülü düzeltip dosyayı kaydedin."
                )

            try:
                toplam_adet = adet_coz(_hucre_al(satir, kolonlar, "adet_metin"))
            except ExcelAktarimHatasi as exc:
                raise ExcelAktarimHatasi(f"'{sayfa_adi}' sayfası, satır {excel_satir_no}: {exc}") from exc

            plan_baslama = _tarih_coz_ve_dogrula(
                deger("plan_baslama"), "Dizgi Başlama Tarihi", excel_satir_no, epoch
            )
            plan_teslim = _tarih_coz_ve_dogrula(
                deger("plan_teslim"), "Planlanan Teslim Tarihi", excel_satir_no, epoch
            )
            gerceklesen = _tarih_coz_ve_dogrula(
                deger("gerceklesen_teslim"), "Gerçekleşen Teslim Tarihi", excel_satir_no, epoch
            )

            if dizgi_tipi == depo.DIZGI_TIPI_EUM and not plan_baslama:
                # EÜM'de ayrı Dizgi Başlama sütunu yok: planlanan tarih/hafta
                # metninin Pazartesi'si plan başlangıcıdır. Açık "T.planlanan
                # tarih" sütunu varsa o esastır (boşsa başlangıç da boş).
                referans = plan_teslim or gerceklesen
                if "eum_planlanan_tarih" in kolonlar:
                    plan_baslama, not_ = _planlanan_baslangic_coz(
                        deger("eum_planlanan_tarih"), "EÜM planlanan tarih", excel_satir_no,
                        epoch, referans, kati=True,
                    )
                else:
                    plan_baslama, not_ = _planlanan_baslangic_coz(
                        deger("plan_hafta"), "Planlanan Başlangıç T.", excel_satir_no,
                        epoch, referans, kati=False,
                    )
                if not_:
                    uyar("plan_baslangic", not_)

            if plan_baslama and plan_teslim and plan_baslama > plan_teslim:
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' sayfası, satır {excel_satir_no}: Dizgi Başlama Tarihi ({plan_baslama}) "
                    f"Planlanan Teslim Tarihinden ({plan_teslim}) sonra olamaz."
                )

            no_index = kolonlar.get("__kimlik_no__")
            no = satir[no_index] if no_index is not None and no_index < len(satir) else None
            source_row_id = _kaynak_satir_kimligi(
                None if _bos_mu(no) else no,
                deger("source_row_id"), sayfa_adi, excel_satir_no,
            )
            source_sheet = depo.dizgi_kodu(dizgi_tipi)
            anahtar = depo.kaynak_anahtari(source_sheet, source_row_id)
            if anahtar in anahtar_gruplari:
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' satır {excel_satir_no}: Tekrarlanan kaynak kimliği {source_row_id} "
                    f"(ilk kullanım satır {anahtar_gruplari[anahtar]})."
                )

            durum_ham = deger("excel_durum")
            if (not _bos_mu(durum_ham) and durum_coz(durum_ham)[0] is None
                    and _sadelestir(durum_ham) not in DURUMSUZ_DURUMLAR):
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' satır {excel_satir_no}: DURUM '{str(durum_ham).strip()}' tanınmıyor; "
                    "kart yanlış durumda görünmesin diye aktarım durduruldu. Excel'deki yazımı "
                    f"düzeltin. Kabul edilen değerler: {KABUL_EDILEN_DURUMLAR_METNI}."
                )
        except ExcelAktarimHatasi as exc:
            mesaj = str(exc)
            if not mesaj.startswith("'"):
                mesaj = f"'{sayfa_adi}' {mesaj}"
            hatalar.append(mesaj)
            continue
        anahtar_gruplari[anahtar] = excel_satir_no

        excel_durum_raw = deger("excel_durum")
        ilk_durum, durum_uyarisi = durum_coz(excel_durum_raw)
        if durum_uyarisi:
            if _bos_mu(excel_durum_raw):
                uyar("durum_bos", "DURUM boş; yeni kartın durumu boş kalır, durumu olan mevcut kart için karar istenir.")
            else:
                uyar("durum_durumsuz",
                     f"DURUM '{str(excel_durum_raw).strip()}' bilerek bir iş akışı durumuna eşlenmez; "
                     "yeni kartın durumu boş kalır, durumu olan mevcut kart için karar istenir.",
                     ek=str(excel_durum_raw).strip())

        if ilk_durum == depo.TESLIM_EDILDI and not gerceklesen:
            uyar("teslim_tarihsiz", "TESLİM EDİLDİ ama Gerçekleşen Teslim T. boş; uygulamada teslim edilen kartın "
                                    "tarihi korunur, diğerlerinde tarih bilinmiyor olarak kalır.")
        elif ilk_durum != depo.TESLIM_EDILDI and gerceklesen:
            uyar("teslim_tarihi_durum_disi", f"DURUM {ilk_durum or 'boş'} iken Gerçekleşen Teslim T. dolu; kart teslim edilmiş sayılmaz.")

        sira, sira_uyarisi = _sira_coz(deger("sira"))
        if sira_uyarisi:
            uyar("no_sayisal_degil", "NO sayısal değil; Sıra alanı boş bırakıldı.")

        plan = {
            "sira": sira,
            "talep_sahibi": str(deger("talep_sahibi") or "").strip(),
            "toplam_adet": toplam_adet,
            "adet_metin": str(deger("adet_metin") or "").strip(),
            "plan_hafta": str(deger("plan_hafta") or "").strip(),
            "plan_baslama": plan_baslama,
            "plan_teslim": plan_teslim,
            "excel_durum": str(excel_durum_raw or "").strip(),
            "pcb": str(deger("pcb") or "").strip(),
            "dizgi_tipi": dizgi_tipi,
            "dizgi_sorumlusu": str(deger("dizgi_sorumlusu") or "").strip(),
            "malzeme_bekliyor": 1 if malzeme_bekliyor_mu(excel_durum_raw) else 0,
        }

        parsed_liste.append(
            {
                "anahtar": anahtar,
                "source_key": anahtar,
                "source_sheet": source_sheet,
                "source_row_id": source_row_id,
                "talep_no": talep_no,
                "stok_no": stok_no,
                "plan": plan,
                "gerceklesen_teslim": gerceklesen,
                "ilk_durum": ilk_durum,
            }
        )
        sayfa_raporu["kayit"] += 1
        sayfa_raporu["gizli_kayit"] += int(gizli)
        for tur, ek, mesaj in satir_uyarilari:
            sayfa_raporu["uyarilar"].append(mesaj)
            sayfa_raporu["uyari_kayitlari"].append(
                {"tur": tur, "deger": ek, "satir": excel_satir_no, "mesaj": mesaj}
            )
        uyari_sayisi += len(satir_uyarilari)

    return bitir()


def excelden_aktar(dosya_yolu, kullanici, onizleme=False, beklenen_surum=None, beklenen_kaynak_surum=None,
                   tamamlanan_sifirla=None, gerileme_secimleri=None, notlari_temizle=None):
    with _import_kilidi:
        return _excelden_aktar(dosya_yolu, kullanici, onizleme, beklenen_surum, beklenen_kaynak_surum,
                               tamamlanan_sifirla, gerileme_secimleri, notlari_temizle)


def _sayfa_bul(wb, sayfa_adi):
    """COM snapshot sayfaları kanonik adla yazar; doğrudan okunan dosyada adlar serbesttir."""
    hedef = _sadelestir(sayfa_adi)
    return next((wb[ad] for ad in wb.sheetnames if _sadelestir(ad) == hedef), None)


def _excelden_aktar(dosya_yolu, kullanici, onizleme=False, beklenen_surum=None, beklenen_kaynak_surum=None,
                    tamamlanan_sifirla=None, gerileme_secimleri=None, notlari_temizle=None):
    snapshot_yolu = None
    wb = None

    try:
        snapshot_yolu = excel_deger_snapshot_olustur(dosya_yolu)
        # read_only değil: satır/sütun gizlilik bilgisi yalnız normal modda okunur.
        wb = openpyxl.load_workbook(snapshot_yolu, data_only=True)
    except ExcelAktarimHatasi:
        raise
    except Exception as exc:
        raise ExcelAktarimHatasi(
            f"Dosya geçerli bir Excel çalışma kitabı değil: {exc}"
        ) from exc

    # Önce bütün sayfalar okunup doğrulanır; tek satır hatası bile varsa
    # depo'ya hiçbir şey gönderilmez (kısmi/eksik okunmuş Excel ile ezme yok).
    rapor = {"sayfalar": []}
    try:
        parsed = []
        anahtar_gruplari = {}
        uyari_sayisi = 0
        hatalar = []

        for sayfa_adi, dizgi_tipi, zorunlu in KAYNAK_SAYFALARI:
            ws = _sayfa_bul(wb, sayfa_adi)
            if ws is None:
                if zorunlu:
                    raise ExcelAktarimHatasi(
                        f"'{sayfa_adi}' sayfası bulunamadı. "
                        f"Dosyadaki sayfalar: {', '.join(wb.sheetnames)}"
                    )
                # Opsiyonel sayfanın yokluğu tam snapshot sözleşmesinde o
                # kaynağın kartlarını pasifleştirir; önizlemede görünür olsun.
                rapor["sayfalar"].append({
                    "sayfa": sayfa_adi, "kod": depo.dizgi_kodu(dizgi_tipi), "bulunamadi": True,
                    "kayit": 0, "gizli_kayit": 0, "bos_satir": 0, "atlanan": [],
                    "atlanan_araliklar": "", "atlanan_gruplari": [], "uyarilar": [],
                    "uyari_kayitlari": [],
                })
                continue

            uyari_sayisi += _sayfa_satirlarini_coz(
                ws, sayfa_adi, dizgi_tipi, anahtar_gruplari, parsed, rapor=rapor, hatalar=hatalar
            )

        if hatalar:
            raise _sorun_hatasi(hatalar)

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if snapshot_yolu:
            try:
                os.remove(snapshot_yolu)
            except OSError:
                pass

    rapor["uyarilar"] = [u for sayfa in rapor["sayfalar"] for u in sayfa["uyarilar"]]
    rapor["uyari_gruplari"] = _uyari_gruplari(rapor["sayfalar"])
    rapor["atlanan_satir"] = sum(len(sayfa["atlanan"]) for sayfa in rapor["sayfalar"])
    rapor["kayit"] = len(parsed)

    # Dosya baytları aynı kalsa bile Excel formülleri yeniden hesaplanabilir.
    kaynak_surum = hashlib.sha256(json.dumps(parsed, sort_keys=True, ensure_ascii=False,
                                            default=str).encode("utf-8")).hexdigest()
    if beklenen_kaynak_surum is not None and kaynak_surum != beklenen_kaynak_surum:
        raise ExcelAktarimHatasi("Excel'in hesaplanan değerleri önizlemeden sonra değişti. Yeni önizleme oluşturun.")
    uygula = depo.excel_import_onizle if onizleme else depo.excel_import_uygula
    secenekler = {"beklenen_surum": beklenen_surum} if beklenen_surum is not None else {}
    if tamamlanan_sifirla:
        secenekler["tamamlanan_sifirla"] = tamamlanan_sifirla
    if gerileme_secimleri is not None:
        secenekler["gerileme_secimleri"] = gerileme_secimleri
    if notlari_temizle:
        secenekler["notlari_temizle"] = notlari_temizle
    # Kartta saklı eski DURUM metninin iş akışı karşılığı: boş DURUM kararında
    # kartın durumunun Excel'den mi uygulamadan mı geldiğini ayırt eder.
    secenekler["kaynak_durum_coz"] = lambda metin: durum_coz(metin)[0]
    sonuc = uygula(
        dosya_adi=os.path.basename(dosya_yolu),
        kullanici=kullanici,
        satirlar=parsed,
        uyari_sayisi=uyari_sayisi,
        kaynak_ozeti=kaynak_raporu_ozeti(rapor),
        **secenekler,
    )
    sonuc["kaynak_raporu"] = rapor
    if onizleme:
        sonuc["kaynak_surum"] = kaynak_surum
    return sonuc


# ---------------------------------------------------------------------------
# Rapor üretimi
# ---------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="0F2027")
BASLIK_YAZI = Font(name="Arial", bold=True, color="FFFFFF", size=11)
GOVDE_YAZI = Font(name="Arial", size=10)

def _excel_hucre_yaz(hucre, deger):
    hucre.value = deger

    if isinstance(deger, str) and deger.startswith("="):
        hucre.data_type = "s"

def _sayfa_yaz(ws, basliklar, satirlar):
    ws.append(basliklar)
    for hucre in ws[1]:
        hucre.fill = BASLIK_DOLGU
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")

    # ws.max_row her çağrıda tüm hücreleri tarar; satır numarası sayaçla verilir.
    for satir_no, satir in enumerate(satirlar, start=2):
        for sutun_no, deger in enumerate(satir,start=1):
            _excel_hucre_yaz(
                ws.cell(row=satir_no,column=sutun_no,),deger)


    for sutun in range(1, len(basliklar) + 1):
        en = max(
            [len(str(basliklar[sutun - 1])), 10]
            + [len(str(satir[sutun - 1])) for satir in satirlar[:400]]
        )
        ws.column_dimensions[get_column_letter(sutun)].width = min(40, en + 3)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for satir in ws.iter_rows(min_row=2):
        for hucre in satir:
            hucre.font = GOVDE_YAZI


def calisma_kitabi_uret(kartlar, loglar, ozet):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Kart Durumları"
    _sayfa_yaz(
        ws,
        [
            "ID",
            "Sıra",
            "Talep NO",
            "Talep Sahibi",
            "Kart Stok No",
            "Toplam Adet",
            "Tamamlanan",
            "Kalan",
            "Durum",
            "Kaynak Durumu",
            "Değerlendirme",
            "Sapma (gün)",
            "Plan Başlangıç",
            "Plan Teslim",
            "Üretim Başlangıç",
            "Üretim Bitiş",
            "Teslim Zamanı",
            "Gerçekleşen Teslim",
            "Operatör",
            "Not",
            "PCB",
            "Dizgi Tipi",
            "Dizgi Sorumlusu",
            "Malzeme Bekliyor",
            "Kaynakta Aktif",
            "Admin Gizli",
        ],
        [
            [
                k.get("id"),
                k.get("sira"),
                k.get("talep_no"),
                k.get("talep_sahibi"),
                k.get("stok_no"),
                k.get("toplam_adet"),
                k.get("tamamlanan_adet"),
                k.get("kalan_adet"),
                k.get("durum") or "DURUMU EKSİK",
                k.get("excel_durum") or "",
                k.get("rozet"),
                k.get("sapma") if k.get("sapma") is not None else "",
                k.get("plan_baslama") or "",
                k.get("plan_teslim") or "",
                k.get("baslama_zamani") or "",
                k.get("bitis_zamani") or "",
                k.get("teslim_zamani") or "",
                k.get("gerceklesen_teslim") or "",
                k.get("operator") or "",
                k.get("aciklama") or "",
                k.get("pcb") or "",
                k.get("dizgi_tipi") or depo.DIZGI_TIPI_MAKINE,
                k.get("dizgi_sorumlusu") or "",
                "EVET" if k.get("malzeme_bekliyor") else "",
                k.get("source_active", 1),
                k.get("admin_gizli", 0),
            ]
            for k in kartlar
        ],
    )

    ws2 = wb.create_sheet("İşlem Logu")
    _sayfa_yaz(
        ws2,
        ["Zaman", "Kullanıcı", "Rol", "İşlem", "Talep NO", "Kart Stok No", "Adet", "Detay", "İşlemi Yapan"],
        [
            [
                l.get("zaman"),
                l.get("kullanici"),
                l.get("rol"),
                l.get("islem"),
                l.get("talep_no") or "",
                l.get("stok_no") or "",
                l.get("adet") if l.get("adet") is not None else "",
                l.get("detay") or "",
                l.get("islem_yapan") or "",
            ]
            for l in loglar
        ],
    )

    ws3 = wb.create_sheet("Özet")
    _sayfa_yaz(ws3, ["Başlık", "Değer"], ozet)
    return wb


def kitap_baytlari(wb):
    tampon = io.BytesIO()
    wb.save(tampon)
    tampon.seek(0)
    return tampon


def dosya_adi(on_ek):
    return f"{on_ek}_{datetime.now():%Y%m%d_%H%M}.xlsx"
```

## `kullanici_yonet.py`

```python
"""PDGM kullanıcı yönetimi için küçük CLI aracı.

Parolalar plaintext saklanmaz; yalnız Werkzeug hash'i kullanicilar.json'a yazılır.
"""

from __future__ import annotations

import getpass
import json
import os
import sys

from werkzeug.security import generate_password_hash

KOK = os.path.dirname(os.path.abspath(__file__))
DOSYA = os.path.join(KOK, "data", "kullanicilar.json")
ROLLER = {"admin", "operator", "gozlemci"}
OPERATOR_TIPLERI = {"makine", "elle_dizgi", "eum_dizgi"}


def _operator_tipi_sor():
    while True:
        secim = input("Operatör tipi (makine / elle_dizgi / eum_dizgi): ").strip().lower()
        if secim in OPERATOR_TIPLERI:
            return secim
        print(f"Geçersiz değer. Şu değerlerden biri olmalı: {', '.join(sorted(OPERATOR_TIPLERI))}")


def oku():
    if not os.path.exists(DOSYA):
        raise SystemExit("Önce uygulamayı bir kez çalıştırın; data/kullanicilar.json oluşturulsun.")
    with open(DOSYA, encoding="utf-8") as f:
        veri = json.load(f)
    if not isinstance(veri, dict):
        raise SystemExit("kullanicilar.json geçersiz.")
    return veri


def yaz(veri):
    gecici = DOSYA + ".yeni"
    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(gecici, DOSYA)


def _parola_sor():
    parola = getpass.getpass("Yeni parola: ")
    parola_tekrar = getpass.getpass("Yeni parola tekrar: ")
    if parola != parola_tekrar:
        raise SystemExit("Parolalar eşleşmiyor.")
    if len(parola) < 8:
        raise SystemExit("Parola en az 8 karakter olmalı.")
    return parola


def _aktif_admin_sayisi(veri, haric=None):
    return sum(
        1
        for k, b in veri.items()
        if k != haric and b.get("rol") == "admin" and b.get("aktif", True)
    )


def ekle(kullanici, rol, gorunen_ad):
    kullanici = kullanici.strip()
    rol = rol.strip().lower()
    gorunen_ad = gorunen_ad.strip()

    if not kullanici:
        raise SystemExit("Kullanıcı adı boş olamaz.")
    if rol not in ROLLER:
        raise SystemExit(f"Rol şu değerlerden biri olmalı: {', '.join(sorted(ROLLER))}")

    veri = oku()
    if kullanici in veri:
        raise SystemExit("Bu kullanıcı adı zaten var.")

    parola = getpass.getpass("Parola: ")
    parola_tekrar = getpass.getpass("Parola tekrar: ")
    if parola != parola_tekrar:
        raise SystemExit("Parolalar eşleşmiyor.")
    if len(parola) < 8:
        raise SystemExit("Parola en az 8 karakter olmalı.")

    operator_tipi = _operator_tipi_sor() if rol == "operator" else None

    veri[kullanici] = {
        "sifre_hash": generate_password_hash(parola),
        "rol": rol,
        "ad": gorunen_ad or kullanici,
        "operator_tipi": operator_tipi,
        "aktif": True,
    }
    yaz(veri)
    print(f"Kullanıcı eklendi: {kullanici} ({rol})" + (f" · operator_tipi={operator_tipi}" if operator_tipi else ""))


def parola_degistir(kullanici):
    kullanici = kullanici.strip()
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    veri[kullanici]["sifre_hash"] = generate_password_hash(_parola_sor())
    yaz(veri)
    print(f"{kullanici}: parola güncellendi.")
    print("Sunucuyu yeniden başlatmanıza gerek yok.")


def rol_degistir(kullanici, yeni_rol):
    kullanici = kullanici.strip()
    yeni_rol = yeni_rol.strip().lower()

    if yeni_rol not in ROLLER:
        raise SystemExit(f"Rol şu değerlerden biri olmalı: {', '.join(sorted(ROLLER))}")

    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    eski_rol = veri[kullanici].get("rol", "-")

    if eski_rol == "admin" and yeni_rol != "admin":
        if _aktif_admin_sayisi(veri, haric=kullanici) == 0:
            raise SystemExit(
                "Bu son aktif admin. Rolü düşürürseniz sistemi yönetemezsiniz. "
                "Önce başka bir admin oluşturun."
            )

    veri[kullanici]["rol"] = yeni_rol

    if yeni_rol == "operator":
        if veri[kullanici].get("operator_tipi") not in OPERATOR_TIPLERI:
            veri[kullanici]["operator_tipi"] = _operator_tipi_sor()
        else:
            print(f"Mevcut operator_tipi korunuyor: {veri[kullanici]['operator_tipi']}")
    else:
        veri[kullanici]["operator_tipi"] = None

    yaz(veri)
    print(f"{kullanici}: {eski_rol} -> {yeni_rol}")
    print("Değişiklik ilk request'te etkili olur; yeniden başlatma gerekmez.")


def aktiflik(kullanici, aktif):
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    if not aktif and veri[kullanici].get("rol") == "admin":
        if _aktif_admin_sayisi(veri, haric=kullanici) == 0:
            raise SystemExit(
                "Bu son aktif admin. Pasife alırsanız sisteme admin olarak "
                "giremezsiniz. Önce başka bir admin oluşturun."
            )

    veri[kullanici]["aktif"] = aktif
    yaz(veri)
    print(f"{kullanici}: {'aktif' if aktif else 'pasif'}")


def listele():
    for kullanici, bilgi in oku().items():
        print(
            f"{kullanici:20} "
            f"{bilgi.get('rol', '-'):10} "
            f"{bilgi.get('ad', '-'):30} "
            f"aktif={bilgi.get('aktif', True)}"
        )


def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Kullanım:\n"
            "  python kullanici_yonet.py listele\n"
            "  python kullanici_yonet.py ekle <kullanici> <admin|operator|gozlemci> <Görünen Ad>\n"
            "  python kullanici_yonet.py parola <kullanici>\n"
            "  python kullanici_yonet.py rol <kullanici> <admin|operator|gozlemci>\n"
            "  python kullanici_yonet.py pasif <kullanici>\n"
            "  python kullanici_yonet.py aktif <kullanici>"
        )

    komut = sys.argv[1].lower()
    if komut == "listele":
        listele()
    elif komut == "ekle" and len(sys.argv) >= 5:
        ekle(sys.argv[2], sys.argv[3], " ".join(sys.argv[4:]))
    elif komut == "parola" and len(sys.argv) == 3:
        parola_degistir(sys.argv[2])
    elif komut == "rol" and len(sys.argv) == 4:
        rol_degistir(sys.argv[2], sys.argv[3])
    elif komut == "pasif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], False)
    elif komut == "aktif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], True)
    else:
        raise SystemExit("Geçersiz komut veya eksik argüman.")


if __name__ == "__main__":
    main()
```

## `old_codebase.txt`

````text
# Repository Codebase Dump

Olusturulma tarihi: 2026-08-24 12:17:00

Repository: `pdgm_flask_updated`

Toplam dosya: **21**


---


## Dosya Listesi

```text
PYTHON
├── app.py
├── depo.py
├── excel_araclari.py
├── kullanici_yonet.py

TEMPLATES
├── templates/base.html
├── templates/giris.html
├── templates/monitor.html
├── templates/operator.html
├── templates/ozet.html
├── templates/panel.html
├── templates/yetkisiz.html
├── templates/yonetim.html

STATIC
├── static/stil.css

SCRIPTS
├── run_pdgm.bat
├── yedek_disari_kopyala.bat

CONFIG / DIGER
├── .env.example
├── .gitignore
├── requirements.txt

DOCS
├── ELDE_DIZGI_DEGISIKLIK_OZETI.md
├── PDGM_Codebase_Teknik_Rehber.md
├── pdgm_kod_incelemesi.md

```


---


# 1. PYTHON


## `app.py`


```python
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
import logging
import logging.handlers
import os
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

    if not request.path.startswith("/static/"):
        response.headers.setdefault(
            "Cache-Control",
            "no-store",
        )

    return response



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
    return None


def yetki(*roller):
    def sarmalayici(fn):
        @wraps(fn)
        def ic(*args, **kwargs):
            if "kullanici" not in session:
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
        "bugun": date.today().strftime("%d.%m.%Y"),
        "csrf_token": _csrf_token_uret() if "kullanici" in session else "",
    }


@app.template_filter("gun")
def gun_filtresi(deger):
    if not deger:
        return "—"
    parcalar = str(deger)[:10].split("-")
    return f"{parcalar[2]}.{parcalar[1]}.{parcalar[0]}" if len(parcalar) == 3 else str(deger)


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
    sayac = {
        "plana_alindi": sum(k["durum"] == depo.PLANA_ALINDI for k in kartlar),
        "dizgide": sum(k["durum"] == depo.DIZGIDE for k in kartlar),
        "teslim": sum(k["durum"] == depo.TESLIM_EDILDI for k in kartlar),
        "gecikme": sum(
            k["renk"] == "kotu" and k["durum"] != depo.TESLIM_EDILDI
            for k in kartlar
        ),
    }
    return {
        "kartlar": kartlar,
        "sayac": sayac,
        "guncelleme": datetime.now().strftime("%H:%M:%S"),
    }

def kart_islem_yetkisi_var_mi(kart):

    if session.get("rol") == "admin":
        return True

    if session.get("rol") != "operator":
        return False

    operator_tipi = session.get("operator_tipi")

    if operator_tipi == "makine":
        return kart.get("dizgi_tipi") == depo.DIZGI_TIPI_MAKINE

    if operator_tipi == "elle_dizgi":
        return kart.get("dizgi_tipi") == depo.DIZGI_TIPI_ELLE

    return False
@app.route("/panel")
@yetki("admin", "operator", "gozlemci")
def panel():
    veri = _pano_verisi()
    kartlar = veri["kartlar"]

    teslim_edilen = sorted(
        [k for k in kartlar if k["durum"] == depo.TESLIM_EDILDI],
        key=lambda k: k.get("teslim_zamani") or k.get("gerceklesen_teslim") or "",
        reverse=True,
    )[:20]

    return render_template(
        "panel.html",
        sayac=veri["sayac"],
        guncelleme=veri["guncelleme"],
        dizgide=[k for k in kartlar if k["durum"] == depo.DIZGIDE],
        plana_alindi=[k for k in kartlar if k["durum"] == depo.PLANA_ALINDI][:12],
        teslim_edilen=teslim_edilen,
    )


@app.route("/monitor")
@yetki("admin", "operator", "gozlemci")
def monitor():
    veri = _pano_verisi()
    kartlar = veri["kartlar"]

    return render_template(
        "monitor.html",
        guncelleme=veri["guncelleme"],
        dizgide=[k for k in kartlar if k["durum"] == depo.DIZGIDE],
        plana_alindi=[k for k in kartlar if k["durum"] == depo.PLANA_ALINDI],
    )


@app.route("/operator")
@yetki("admin", "operator")
def operator():
    veri = _pano_verisi()
    return render_template("operator.html", kartlar=veri["kartlar"], sayac=veri["sayac"])


@app.route("/yonetim")
@yetki("admin")
def yonetim():
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
    )





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
    except depo.IsKuralHatasi as hata:
        return jsonify(hata=str(hata)), 409
    except (depo.VeriDogrulamaHatasi, TypeError, ValueError) as hata:
        return jsonify(hata=str(hata)), 400


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
            aciklama=(veri.get("not") or "").strip(),
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
            aciklama=(veri.get("not") or "").strip(),
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
            aciklama=(veri.get("not") or "").strip(),
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
            elle_dizgi=bool(veri.get("elle_dizgi")),
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
        kart = depo.admin_kart_duzenle(
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
            dizgi_sorumlusu=veri.get("dizgi_sorumlusu"),
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
        sonuc = ex.excelden_aktar(yol, session["kullanici"])
        mesaj = (
            f"Excel aktarıldı · {sonuc['satir']} satır · {sonuc['yeni']} yeni · "
            f"{sonuc['guncellenen']} güncellendi · {sonuc['workflow_korundu']} workflow korundu."
        )
        if sonuc.get("pasife_alinan"):
            mesaj += f" · {sonuc['pasife_alinan']} kart kaynak Excel'de artık yok."
        if sonuc.get("elle_dizgi_satir"):
            mesaj += f" · {sonuc['elle_dizgi_satir']} kart ELDE DİZGİ YENİ sayfasından."
        if sonuc.get("uyari"):
            mesaj += (
                f" · {sonuc['uyari']} satırda DURUM boş/geçersiz veya başka bir küçük veri uyarısı var; "
                "bu kartlar Yönetim ekranında kontrol edilebilir."
            )
        if sonuc.get("yedek"):
            mesaj += f" · Yedek: {os.path.basename(sonuc['yedek'])}"

        flash(mesaj, "basari")
        app.logger.info("Excel import OK: %s", mesaj)
        _yuklenen_exceleri_buda()
    except (ex.ExcelAktarimHatasi, depo.IsKuralHatasi, depo.VeriDogrulamaHatasi) as hata:
        flash(f"Excel içeriği kabul edilmedi: {hata}", "hata")
        app.logger.warning("Excel import reddedildi: %s", hata)
    except Exception:  # noqa: BLE001
        app.logger.exception("Excel import sırasında beklenmeyen hata")
        flash("Excel okunurken beklenmeyen bir hata oluştu. Uygulama logunu kontrol edin.", "hata")

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


def _donem_ozeti(kartlar, baslangic, bitis):
    alt = baslangic.strftime("%Y-%m-%d")
    ust = bitis.strftime("%Y-%m-%d")
    secilen = []

    for kart in kartlar:
        if kart.get("durum") != depo.TESLIM_EDILDI:
            continue
        teslim = _teslim_tarihi(kart)
        if teslim and alt <= teslim <= ust:
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
    kartlar = [
        kart
        for kart in depo.kartlari_getir(sadece_gorunen=False)
        if not kart.get("admin_gizli") and kart.get("aktif", 1) == 1
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

    try:
        alt, ust = _tarih_araligi_coz(
            aralik,
            request.args.get("baslangic"),
            request.args.get("bitis"),
        )
    except ValueError as hata:
        return jsonify(hata=str(hata)), 400

    kartlar = [
        kart
        for kart in depo.kartlari_getir(sadece_gorunen=False)
        if not kart.get("admin_gizli") and kart.get("aktif", 1) == 1
    ]

    if alt is None:
        donem_ozet = _donem_ozeti(kartlar, date(2000, 1, 1), date.today())
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
                "stok_no": k.get("stok_no"),
                "toplam_adet": k.get("toplam_adet"),
                "teslim": gun_filtresi(_teslim_tarihi(k)),
                "operator": k.get("operator"),
                "rozet": k.get("rozet"),
                "renk": k.get("renk"),
                "elle_dizgi_mi": k.get("elle_dizgi_mi", False),
                "dizgi_sorumlusu": k.get("dizgi_sorumlusu"),
            }
            for k in secilenler[:200]
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
```


## `depo.py`


```python
"""PDGM İş Takip Sistemi için Excel tabanlı veri deposu.

Tasarım hedefi:
- kartlar.xlsx uygulamanın source of truth dosyasıdır.
- Tek Python process + çok thread modeli kullanılır.
- Read/modify/write işlemleri RLock ile korunur.
- Yazmalar temp dosya + os.replace ile yapılır.
- Kritik işlemler öncesi yedek alınır.
- Workflow yalnız dört gerçek durumdan oluşur:
    PLANA ALINDI -> DİZGİDE -> HAZIR -> TESLİM EDİLDİ
- Kaynak Excel'de DURUM boşsa kart saklanır fakat operasyon ekranlarında gösterilmez.

Not: Excel transactional database değildir. Aynı data klasörünü birden fazla Python
process'i paylaşmamalıdır. Bu sınır process lock ile açıkça korunur.
"""

from __future__ import annotations

import atexit
import copy
import os
import re
import shutil
import subprocess
import threading
from datetime import date, datetime
from uuid import uuid4

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel


# ---------------------------------------------------------------------------
# Dosyalar ve workflow sabitleri
# ---------------------------------------------------------------------------

KOK = os.path.dirname(os.path.abspath(__file__))
VERI_KLASORU = os.path.join(KOK, "data")
KARTLAR_DOSYA = os.path.join(VERI_KLASORU, "kartlar.xlsx")
LOG_DOSYA = os.path.join(VERI_KLASORU, "islem_logu.xlsx")
YUKLEME_DOSYA = os.path.join(VERI_KLASORU, "yuklemeler.xlsx")
YEDEK_KLASORU = os.path.join(VERI_KLASORU, "yedekler")
PROCESS_KILIT = os.path.join(VERI_KLASORU, "sunucu.lock")

PLANA_ALINDI = "PLANA ALINDI"
DIZGIDE = "DİZGİDE"
HAZIR = "HAZIR"
TESLIM_EDILDI = "TESLİM EDİLDİ"

# Dizgi tipi: kart MAKİNE sayfasından mı yoksa ELDE DİZGİ YENİ sayfasından mı geldi.
# Bu, workflow'dan (durum) tamamen bağımsız ikinci bir boyuttur; kaynağı yapısal
# olarak (hangi sheet'ten okunduğu) belirlenir, DURUM metninden tahmin edilmez.
DIZGI_TIPI_MAKINE = "MAKİNE"
DIZGI_TIPI_ELLE = "ELLE DİZGİ"
GECERLI_DIZGI_TIPLERI = {DIZGI_TIPI_MAKINE, DIZGI_TIPI_ELLE}

AKTIF_DURUMLAR = {PLANA_ALINDI, DIZGIDE}  # HAZIR artık "aktif iş" değil, backlog
GECERLI_DURUMLAR = {PLANA_ALINDI, DIZGIDE, HAZIR, TESLIM_EDILDI}
# Yeni: operasyon ekranlarında (panel/operatör/monitör) görünür durumlar.
# HAZIR bilinçli olarak dışarıda — yalnız admin görsün.
OPERASYONEL_DURUMLAR = {PLANA_ALINDI, DIZGIDE, TESLIM_EDILDI}

# Admin tablosu sıralaması: aktif üretim > planlanan > backlog > tamamlanan > eksik
SIRALAMA = {
    DIZGIDE: 0,
    PLANA_ALINDI: 1,
    HAZIR: 2,
    TESLIM_EDILDI: 3,
    None: 4,
}

# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------

class DepoHatasi(Exception):
    """Depo katmanının temel exception sınıfı."""


class KartBulunamadi(DepoHatasi):
    pass


class IsKuralHatasi(DepoHatasi):
    pass


class VeriDogrulamaHatasi(DepoHatasi):
    pass


# ---------------------------------------------------------------------------
# Excel şemaları
# ---------------------------------------------------------------------------

KART_ALANLARI = [
    ("ID", "id"),
    ("Sıra", "sira"),
    ("Talep NO", "talep_no"),
    ("Kart Stok No", "stok_no"),
    ("Talep Sahibi", "talep_sahibi"),
    ("Toplam Adet", "toplam_adet"),
    ("Adet Metni", "adet_metin"),
    ("Plan Haftası", "plan_hafta"),
    ("Plan Başlangıç", "plan_baslama"),
    ("Plan Teslim", "plan_teslim"),
    ("Gerçekleşen Teslim", "gerceklesen_teslim"),
    ("Excel Durumu", "excel_durum"),
    ("PCB", "pcb"),
    ("Dizgi Tipi", "dizgi_tipi"),
    ("Dizgi Sorumlusu", "dizgi_sorumlusu"),
    ("Malzeme Bekliyor", "malzeme_bekliyor"),
    ("Durum", "durum"),
    ("Başlangıç Adedi", "baslangic_adet"),
    ("Tamamlanan Adet", "tamamlanan_adet"),
    ("Başlama Zamanı", "baslama_zamani"),
    ("Üretim Bitiş Zamanı", "bitis_zamani"),
    ("Teslim Zamanı", "teslim_zamani"),
    ("Operatör", "operator"),
    ("Not", "aciklama"),
    ("Son Güncelleme", "guncelleme"),
    ("Listede", "aktif"),
    ("Kaynakta Aktif", "source_active"),
    ("Admin Gizli", "admin_gizli"),
    ("Kaynak", "kaynak"),
    ("Anahtar", "anahtar"),
]

LOG_ALANLARI = [
    ("Zaman", "zaman"),
    ("Kullanıcı", "kullanici"),
    ("Rol", "rol"),
    ("İşlem", "islem"),
    ("Talep NO", "talep_no"),
    ("Kart Stok No", "stok_no"),
    ("Adet", "adet"),
    ("Detay", "detay"),
]

YUKLEME_ALANLARI = [
    ("Zaman", "zaman"),
    ("Kullanıcı", "kullanici"),
    ("Dosya", "dosya"),
    ("Okunan Satır", "satir"),
    ("Yeni Kart", "yeni"),
    ("Güncellenen", "guncellenen"),
    ("Kaynakta Olmayan", "pasife_alinan"),
    ("Uyarı", "uyari"),
]

SAYISAL_ALANLAR = {
    "id",
    "sira",
    "toplam_adet",
    "baslangic_adet",
    "tamamlanan_adet",
    "aktif",
    "source_active",
    "admin_gizli",
    "malzeme_bekliyor",
    "adet",
    "satir",
    "yeni",
    "guncellenen",
    "pasife_alinan",
    "uyari",
}

ZORUNLU_KART_ALANLARI = {
    "id",
    "talep_no",
    "stok_no",
    "toplam_adet",
    "tamamlanan_adet",
    "anahtar",
}


# ---------------------------------------------------------------------------
# In-memory state
# ---------------------------------------------------------------------------

_kilit = threading.RLock()
_kartlar: list[dict] = []
_loglar: list[dict] = []
_yuklemeler: list[dict] = []

LOG_SINIRI = 20_000
LOG_SAKLA = 5_000

BASLIK_DOLGU = PatternFill("solid", fgColor="0F2027")
BASLIK_YAZI = Font(name="Arial", bold=True, color="FFFFFF", size=11)
GOVDE_YAZI = Font(name="Arial", size=10)


# ---------------------------------------------------------------------------
# Genel yardımcılar
# ---------------------------------------------------------------------------

def simdi() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def bugun() -> str:
    return date.today().strftime("%Y-%m-%d")


def _sayi(deger, varsayilan=0) -> int:
    try:
        return int(float(deger))
    except (TypeError, ValueError):
        return varsayilan


def _temiz_metin(deger) -> str:
    return str(deger or "").strip()


def _durum_normalize(deger):
    """Dört gerçek workflow durumunu normalize eder; boş değer None olarak kalır."""
    metin = _temiz_metin(deger)
    if not metin:
        return None

    sade = (
        metin.upper()
        .replace("İ", "I")
        .replace("Ğ", "G")
        .replace("Ü", "U")
        .replace("Ş", "S")
        .replace("Ö", "O")
        .replace("Ç", "C")
    )
    sade = re.sub(r"\s+", " ", sade).strip()

    esleme = {
        "PLANA ALINDI": PLANA_ALINDI,
        "DIZGIDE": DIZGIDE,
        "HAZIR": HAZIR,
        "TESLIM EDILDI": TESLIM_EDILDI,
    }
    return esleme.get(sade)


def tarih_coz(deger):
    """Excel veya kullanıcı girdisini YYYY-MM-DD biçimine çevirir."""
    if deger is None:
        return None

    if isinstance(deger, datetime):
        return deger.strftime("%Y-%m-%d")
    if isinstance(deger, date):
        return deger.strftime("%Y-%m-%d")

    if isinstance(deger, (int, float)) and not isinstance(deger, bool):
        try:
            sonuc = from_excel(deger)
            if isinstance(sonuc, (datetime, date)):
                return sonuc.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OverflowError):
            return None

    metin = str(deger).strip()
    if not metin or metin.upper() in {"-", "YOK", "N/A", "NONE"}:
        return None

    if re.fullmatch(r"\d+(?:\.\d+)?", metin):
        try:
            sonuc = from_excel(float(metin))
            if isinstance(sonuc, (datetime, date)):
                return sonuc.strftime("%Y-%m-%d")
        except (TypeError, ValueError, OverflowError):
            pass

    for kalip in (
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(metin, kalip).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def gun_farki(a, b):
    """a - b gün farkını döndürür."""
    if not a or not b:
        return None
    try:
        t1 = datetime.strptime(str(a)[:10], "%Y-%m-%d").date()
        t2 = datetime.strptime(str(b)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return (t1 - t2).days


def _kart_ref(kart_id: int):
    return next((kart for kart in _kartlar if kart.get("id") == kart_id), None)


def _operasyonda_gorunur_mu(kart: dict) -> bool:
    return (
        kart.get("aktif", 1) == 1
        and kart.get("admin_gizli", 0) != 1
        and kart.get("durum") in OPERASYONEL_DURUMLAR
    )

def operator_kart_yetkisi_var_mi(kart: dict, operator_tipi: str) -> bool:
    """
    Operatörün kart üzerinde işlem yapıp yapamayacağını kontrol eder.

    Görme yetkisi değildir.
    Sadece operasyon aksiyonları içindir.
    """

    if operator_tipi == "makine_operator":
        return kart.get("dizgi_tipi") == DIZGI_TIPI_MAKINE

    if operator_tipi == "elle_dizgi_operatoru":
        return kart.get("dizgi_tipi") == DIZGI_TIPI_ELLE

    return False

def _yonetimde_gorunur_mu(kart: dict) -> bool:
    return kart.get("aktif", 1) == 1 and kart.get("admin_gizli", 0) != 1


# ---------------------------------------------------------------------------
# Dosya okuma/yazma
# ---------------------------------------------------------------------------

def _oku(dosya, alanlar, zorunlu_alanlar=frozenset()):
    if not os.path.exists(dosya):
        return []

    try:
        wb = openpyxl.load_workbook(dosya, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        raise VeriDogrulamaHatasi(
            f"'{os.path.basename(dosya)}' açılamadı: {exc}"
        ) from exc

    try:
        ws = wb[wb.sheetnames[0]]
        satirlar = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not satirlar:
        return []

    basliklar = [str(h or "").strip() for h in satirlar[0]]
    yerlesim = {
        alan: basliklar.index(baslik)
        for baslik, alan in alanlar
        if baslik in basliklar
    }

    eksik = sorted(alan for alan in zorunlu_alanlar if alan not in yerlesim)
    if eksik:
        ters = {alan: baslik for baslik, alan in alanlar}
        eksik_adlar = ", ".join(ters.get(alan, alan) for alan in eksik)
        raise VeriDogrulamaHatasi(
            f"'{os.path.basename(dosya)}' zorunlu sütunları eksik: {eksik_adlar}"
        )

    kayitlar = []
    for satir in satirlar[1:]:
        if not any(hucre not in (None, "") for hucre in satir):
            continue

        kayit = {}
        for _, alan in alanlar:
            index = yerlesim.get(alan)
            deger = satir[index] if index is not None and index < len(satir) else None

            if alan in SAYISAL_ALANLAR:
                kayit[alan] = _sayi(deger, 0) if deger not in (None, "") else None
            elif isinstance(deger, datetime):
                kayit[alan] = deger.strftime("%Y-%m-%d %H:%M:%S")
            else:
                kayit[alan] = str(deger).strip() if deger not in (None, "") else None
        kayitlar.append(kayit)

    return kayitlar

def _excel_hucre_yaz(hucre, deger):
    """'=' ile başlayan kullanıcı metninin Excel formülüne dönüşmesini engeller."""
    hucre.value = deger

    if isinstance(deger, str) and deger.startswith("="):
        hucre.data_type = "s"

def _workbook_uret(alanlar, kayitlar, sayfa_adi):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sayfa_adi
    ws.append([baslik for baslik, _ in alanlar])

    for hucre in ws[1]:
        hucre.fill = BASLIK_DOLGU
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")

    for kayit in kayitlar:
        satir_no = ws.max_row + 1

        for sutun_no, (_, alan) in enumerate(alanlar, start=1):
            _excel_hucre_yaz(
                ws.cell(row=satir_no, column=sutun_no),
                kayit.get(alan),
            )

    for satir in ws.iter_rows(min_row=2):
        for hucre in satir:
            hucre.font = GOVDE_YAZI

    for sutun, (baslik, alan) in enumerate(alanlar, start=1):
        en = max(
            [len(baslik), 10]
            + [len(str(kayit.get(alan) or "")) for kayit in kayitlar[:300]]
        )
        ws.column_dimensions[get_column_letter(sutun)].width = min(42, en + 3)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return wb


def _diske_zorla(yol):
    """Temp dosya içeriğinin OS page cache'ten diske inmesini zorlar."""
    try:
        fd = os.open(yol, os.O_RDONLY)
    except OSError:
        return
    try:
        os.fsync(fd)
    except OSError:
        pass
    finally:
        os.close(fd)


def _temp_yaz(hedef, alanlar, kayitlar, sayfa_adi):
    os.makedirs(os.path.dirname(hedef), exist_ok=True)
    temp = f"{hedef}.{uuid4().hex}.yeni"
    wb = _workbook_uret(alanlar, kayitlar, sayfa_adi)
    try:
        wb.save(temp)
    finally:
        wb.close()
    _diske_zorla(temp)
    return temp


def _coklu_yaz(dosyalar):
    """Birden fazla Excel dosyasını temp + replace + rollback yaklaşımıyla yazar."""
    os.makedirs(VERI_KLASORU, exist_ok=True)
    temps = []
    backups = {}
    degisen = []
    commit_basarili = False

    try:
        for hedef, alanlar, kayitlar, sayfa_adi in dosyalar:
            temps.append((hedef, _temp_yaz(hedef, alanlar, kayitlar, sayfa_adi)))

        for hedef, _ in temps:
            if os.path.exists(hedef):
                backup = f"{hedef}.{uuid4().hex}.txn.bak"
                shutil.copy2(hedef, backup)
                backups[hedef] = backup
            else:
                backups[hedef] = None

        for hedef, temp in temps:
            os.replace(temp, hedef)
            degisen.append(hedef)

        commit_basarili = True

    except Exception:
        for hedef in reversed(degisen):
            backup = backups.get(hedef)
            try:
                if backup and os.path.exists(backup):
                    os.replace(backup, hedef)
                    backups[hedef] = None
                elif os.path.exists(hedef):
                    os.remove(hedef)
            except OSError:
                pass
        raise
    finally:
        for _, temp in temps:
            if os.path.exists(temp):
                try:
                    os.remove(temp)
                except OSError:
                    pass
        if commit_basarili:
            for backup in backups.values():
                if backup and os.path.exists(backup):
                    try:
                        os.remove(backup)
                    except OSError:
                        pass


def _yaz(dosya, alanlar, kayitlar, sayfa_adi):
    _coklu_yaz([(dosya, alanlar, kayitlar, sayfa_adi)])


def _gunluk_yedek(dosya):
    if not os.path.exists(dosya):
        return
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    hedef = os.path.join(
        YEDEK_KLASORU,
        f"{date.today():%Y%m%d}_{os.path.basename(dosya)}",
    )
    if not os.path.exists(hedef):
        shutil.copy2(dosya, hedef)


ANLIK_YEDEK_SAKLA = 30
GUNLUK_YEDEK_GUN = 90
ANLIK_YEDEK_DESEN = re.compile(r"^\d{8}_\d{6}_")


def yedekleri_buda():
    """Yalnız PDGM naming pattern'li yedekleri sınırlar. Hata olursa sessizce geçer."""
    try:
        if not os.path.isdir(YEDEK_KLASORU):
            return

        anlik = []
        for ad in os.listdir(YEDEK_KLASORU):
            yol = os.path.join(YEDEK_KLASORU, ad)

            if os.path.isdir(yol) and ANLIK_YEDEK_DESEN.match(ad):
                try:
                    anlik.append((os.path.getmtime(yol), yol))
                except OSError:
                    pass
                continue

            if re.fullmatch(r"\d{8}_kartlar\.xlsx", ad, flags=re.IGNORECASE):
                try:
                    yas_gun = (
                        datetime.now() - datetime.fromtimestamp(os.path.getmtime(yol))
                    ).days
                    if yas_gun > GUNLUK_YEDEK_GUN:
                        os.remove(yol)
                except OSError:
                    pass

        anlik.sort(reverse=True)
        for _, yol in anlik[ANLIK_YEDEK_SAKLA:]:
            shutil.rmtree(yol, ignore_errors=True)
    except Exception:  # noqa: BLE001
        pass


def anlik_yedek(etiket: str = "once") -> str:
    """Kart/log/yükleme dosyalarının timestamp'li güvenlik kopyasını alır."""
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    damga = datetime.now().strftime("%Y%m%d_%H%M%S")
    guvenli = re.sub(r"[^0-9A-Za-z_-]+", "_", etiket or "yedek")
    klasor = os.path.join(YEDEK_KLASORU, f"{damga}_{guvenli}")
    os.makedirs(klasor, exist_ok=True)

    for dosya in (KARTLAR_DOSYA, LOG_DOSYA, YUKLEME_DOSYA):
        if os.path.exists(dosya):
            shutil.copy2(dosya, os.path.join(klasor, os.path.basename(dosya)))

    yedekleri_buda()
    return klasor


def yedekleri_getir(adet=12):
    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    sonuc = []

    for ad in os.listdir(YEDEK_KLASORU):
        yol = os.path.join(YEDEK_KLASORU, ad)

        if os.path.isdir(yol):
            kart_dosyasi = os.path.join(yol, os.path.basename(KARTLAR_DOSYA))
            if not os.path.isfile(kart_dosyasi):
                continue
            try:
                mtime = os.path.getmtime(kart_dosyasi)
                boyut = os.path.getsize(kart_dosyasi)
            except OSError:
                continue

            parcalar = ad.split("_", 2)
            etiket = parcalar[2].replace("_", " ") if len(parcalar) >= 3 else "anlık yedek"
            sonuc.append(
                {
                    "ad": ad,
                    "tip": "Anlık",
                    "etiket": etiket,
                    "zaman": datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S"),
                    "boyut_kb": round(boyut / 1024, 1),
                    "_mtime": mtime,
                }
            )
            continue

        if not os.path.isfile(yol) or not re.fullmatch(
            r"\d{8}_kartlar\.xlsx", ad, flags=re.IGNORECASE
        ):
            continue

        try:
            mtime = os.path.getmtime(yol)
            boyut = os.path.getsize(yol)
        except OSError:
            continue

        sonuc.append(
            {
                "ad": ad,
                "tip": "Günlük",
                "etiket": "günlük otomatik yedek",
                "zaman": datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S"),
                "boyut_kb": round(boyut / 1024, 1),
                "_mtime": mtime,
            }
        )

    sonuc.sort(key=lambda kayit: kayit["_mtime"], reverse=True)
    for kayit in sonuc:
        kayit.pop("_mtime", None)
    return sonuc if adet is None else sonuc[:adet]


def _yedek_kart_dosyasi_bul(yedek_adi):
    yedek_adi = _temiz_metin(yedek_adi)
    if not yedek_adi:
        raise IsKuralHatasi("Yedek seçilmedi.")
    if os.path.basename(yedek_adi) != yedek_adi:
        raise IsKuralHatasi("Geçersiz yedek adı.")

    yedek_kok = os.path.abspath(YEDEK_KLASORU)
    yol = os.path.abspath(os.path.join(yedek_kok, yedek_adi))
    try:
        if os.path.commonpath([yedek_kok, yol]) != yedek_kok:
            raise IsKuralHatasi("Geçersiz yedek yolu.")
    except ValueError as exc:
        raise IsKuralHatasi("Geçersiz yedek yolu.") from exc

    if os.path.isdir(yol):
        aday = os.path.join(yol, os.path.basename(KARTLAR_DOSYA))
        if os.path.isfile(aday):
            return aday
        raise IsKuralHatasi("Seçilen yedekte kartlar.xlsx bulunamadı.")

    if os.path.isfile(yol) and re.fullmatch(
        r"\d{8}_kartlar\.xlsx", yedek_adi, flags=re.IGNORECASE
    ):
        return yol

    raise IsKuralHatasi("Seçilen yedek kullanılamıyor.")


# ---------------------------------------------------------------------------
# Process lock
# ---------------------------------------------------------------------------



def _process_kilit_sahibi():
    if not os.path.exists(PROCESS_KILIT):
        return None

    try:
        with open(PROCESS_KILIT, encoding="utf-8") as f:
            metin = (f.read() or "").strip()
        ilk = metin.split("|", 1)[0].strip()
        return int(ilk) if ilk else None
    except (OSError, ValueError):
        return None


def _pid_calisiyor_mu(pid: int):
    """PID canlı mı? Karar verilemezse True (fail-closed: kilidi koru)."""
    if pid is None or pid <= 0:
        return False

    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        except OSError:
            return True

    try:
        cikti = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return True

    if cikti.returncode != 0:
        return True

    return f'"{pid}"' in (cikti.stdout or "")


def _pid_python_mu(pid: int):
    """PID python/pythonw sürecine mi ait? Karar verilemezse True (fail-closed)."""
    if pid is None or pid <= 0:
        return False

    if os.name != "nt":
        try:
            import sys

            if pid == os.getpid():
                return True
            cmdline_yolu = f"/proc/{pid}/cmdline"
            if os.path.exists(cmdline_yolu):
                with open(cmdline_yolu, "rb") as f:
                    cmd = f.read().decode("utf-8", errors="ignore").lower()
                return "python" in cmd
            return True
        except OSError:
            return True

    try:
        cikti = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return True

    if cikti.returncode != 0:
        return True

    satir = (cikti.stdout or "").strip().lower()
    if f'"{pid}"' not in satir and str(pid) not in satir:
        return False
    return "python.exe" in satir or "pythonw.exe" in satir


def process_kilidi_al():
    """Aynı data klasörünü ikinci PDGM process'inin açmasını atomik olarak engeller.

    Not:
    - Windows'ta os.kill(pid, 0) yerine tasklist kullanılır.
    - Lock dosyası O_EXCL ile atomik oluşturulur.
    - Stale lock: PID ölüyse veya PID canlı ama python değilse (PID reuse) devralınır.
    - Stale reclaim os.replace ile atomik yapılır (remove+O_EXCL yarışı yok).
    - PID canlı ve python ise reddedilir. tasklist başarısızsa fail-closed.
    """
    os.makedirs(VERI_KLASORU, exist_ok=True)
    mevcut_pid = _process_kilit_sahibi()

    if mevcut_pid == os.getpid():
        return

    if os.path.exists(PROCESS_KILIT):
        canli = mevcut_pid is not None and _pid_calisiyor_mu(mevcut_pid)
        python_sureci = canli and _pid_python_mu(mevcut_pid)

        if canli and python_sureci:
            raise RuntimeError(
                f"data/ klasörü başka bir PDGM process tarafından kilitli "
                f"(PID {mevcut_pid}) ve bu process ŞU AN ÇALIŞIYOR. "
                "İkinci sunucu açmayın."
            )

        sahip = mevcut_pid if mevcut_pid is not None else "bilinmiyor"
        if canli and not python_sureci:
            print(
                f"UYARI: data/sunucu.lock PID {sahip} başka bir uygulamaya ait "
                "(PID reuse). Kalıntı kilit devralınıyor."
            )
        else:
            print(
                f"UYARI: data/sunucu.lock artık çalışmayan bir process'e ait "
                f"(PID {sahip}). Kalıntı kilit devralınıyor."
            )

        # Atomik reclaim: remove+O_EXCL yarışını önlemek için stale lock'u
        # benzersiz bir isme taşı. İki process aynı anda denerse yalnız biri
        # os.replace kazanır; diğeri FileNotFoundError alır ve O_EXCL'de kaybeder.
        stale_yol = f"{PROCESS_KILIT}.stale.{uuid4().hex}"
        try:
            os.replace(PROCESS_KILIT, stale_yol)
        except FileNotFoundError:
            stale_yol = None
        except OSError as exc:
            raise RuntimeError(
                "data/sunucu.lock devralınamadı: "
                f"{exc}. Dosyayı manuel silip tekrar deneyin."
            ) from exc
    else:
        stale_yol = None

    bayraklar = os.O_WRONLY | os.O_CREAT | os.O_EXCL

    try:
        fd = os.open(PROCESS_KILIT, bayraklar)
    except FileExistsError as exc:
        if stale_yol:
            try:
                os.remove(stale_yol)
            except OSError:
                pass
        sahip = _process_kilit_sahibi()
        raise RuntimeError(
            f"data/ klasörü başka bir PDGM process tarafından kilitlendi "
            f"(PID {sahip if sahip is not None else 'bilinmiyor'})."
        ) from exc

    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}|{datetime.now():%Y-%m-%d %H:%M:%S}")
            f.flush()
            os.fsync(f.fileno())
    except Exception:
        try:
            os.remove(PROCESS_KILIT)
        except OSError:
            pass
        if stale_yol:
            try:
                os.remove(stale_yol)
            except OSError:
                pass
        raise

    if stale_yol:
        try:
            os.remove(stale_yol)
        except OSError:
            pass

    if getattr(process_kilidi_al, "_atexit_bagli", False):
        return

    def _birak():
        try:
            if _process_kilit_sahibi() == os.getpid():
                os.remove(PROCESS_KILIT)
        except OSError:
            pass

    atexit.register(_birak)
    process_kilidi_al._atexit_bagli = True


# ---------------------------------------------------------------------------
# Kart normalizasyonu ve doğrulama
# ---------------------------------------------------------------------------

def _kart_normalize(kart: dict) -> dict:
    kart = dict(kart)
    kart["id"] = _sayi(kart.get("id"), 0)
    kart["sira"] = _sayi(kart.get("sira"), 0) or None
    kart["toplam_adet"] = _sayi(kart.get("toplam_adet"), 1) or 1
    kart["baslangic_adet"] = _sayi(kart.get("baslangic_adet"), 0)
    kart["tamamlanan_adet"] = _sayi(kart.get("tamamlanan_adet"), 0)
    kart["aktif"] = 0 if kart.get("aktif") == 0 else 1
    kart["source_active"] = 0 if kart.get("source_active") == 0 else 1
    kart["admin_gizli"] = 1 if kart.get("admin_gizli") == 1 else 0
    kart["kaynak"] = _temiz_metin(kart.get("kaynak")) or "EXCEL"
    kart["durum"] = _durum_normalize(kart.get("durum"))

    kart["dizgi_tipi"] = _temiz_metin(kart.get("dizgi_tipi")) or DIZGI_TIPI_MAKINE
    if kart["dizgi_tipi"] not in GECERLI_DIZGI_TIPLERI:
        kart["dizgi_tipi"] = DIZGI_TIPI_MAKINE
    kart["dizgi_sorumlusu"] = _temiz_metin(kart.get("dizgi_sorumlusu")) or None
    kart["malzeme_bekliyor"] = 1 if kart.get("malzeme_bekliyor") == 1 else 0

    for alan in ("plan_baslama", "plan_teslim", "gerceklesen_teslim"):
        ham = kart.get(alan)
        if ham in (None, ""):
            kart[alan] = None
            continue
        cozulmus = tarih_coz(ham)
        if not cozulmus:
            raise VeriDogrulamaHatasi(
                f"Kart {kart.get('id') or '?'}: {alan} geçerli bir tarih değil: {ham!r}"
            )
        kart[alan] = cozulmus

    return kart


def _kart_dogrula(kart: dict):
    kart_id = _sayi(kart.get("id"), -1)
    toplam = _sayi(kart.get("toplam_adet"), 0)
    tamam = _sayi(kart.get("tamamlanan_adet"), 0)
    durum = kart.get("durum")

    if kart_id < 1:
        raise VeriDogrulamaHatasi("Kart ID pozitif tam sayı olmalı.")
    if not _temiz_metin(kart.get("talep_no")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Talep NO boş olamaz.")
    if not _temiz_metin(kart.get("stok_no")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Kart Stok No boş olamaz.")
    if toplam < 1:
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Toplam Adet en az 1 olmalı.")
    if tamam < 0 or tamam > toplam:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: Tamamlanan Adet ({tamam}) 0 ile Toplam Adet ({toplam}) arasında olmalı."
        )
    if durum is not None and durum not in GECERLI_DURUMLAR:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: Geçersiz durum '{durum}'. "
            f"Geçerli durumlar: {', '.join(sorted(GECERLI_DURUMLAR))}."
        )
    if durum in (None, PLANA_ALINDI, HAZIR) and tamam != 0:
        etiket = "Durumu boş" if durum is None else durum
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: {etiket} kartta Tamamlanan Adet 0 olmalı."
        )
    if durum == TESLIM_EDILDI and tamam != toplam:
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: {durum} durumunda Tamamlanan Adet Toplam Adet'e eşit olmalı."
        )
    if durum == TESLIM_EDILDI and not kart.get("gerceklesen_teslim"):
        raise VeriDogrulamaHatasi(
            f"Kart {kart_id}: TESLİM EDİLDİ durumunda Gerçekleşen Teslim tarihi zorunlu."
        )
    if kart.get("plan_baslama") and kart.get("plan_teslim"):
        if kart["plan_baslama"] > kart["plan_teslim"]:
            raise VeriDogrulamaHatasi(
                f"Kart {kart_id}: Plan başlangıç tarihi plan teslim tarihinden sonra olamaz."
            )
    if not _temiz_metin(kart.get("anahtar")):
        raise VeriDogrulamaHatasi(f"Kart {kart_id}: Anahtar boş olamaz.")


def _kart_listesi_dogrula(kartlar):
    idler = set()
    anahtarlar = set()

    for kart in kartlar:
        _kart_dogrula(kart)
        if kart["id"] in idler:
            raise VeriDogrulamaHatasi(f"Tekrarlanan kart ID: {kart['id']}")
        if kart["anahtar"] in anahtarlar:
            raise VeriDogrulamaHatasi(f"Tekrarlanan kart anahtarı: {kart['anahtar']}")
        idler.add(kart["id"])
        anahtarlar.add(kart["anahtar"])


# ---------------------------------------------------------------------------
# Başlangıç ve reload
# ---------------------------------------------------------------------------

def _baslatma_uyarisi_yaz(mesaj: str) -> None:
    """Açılış uyarılarını konsola basar ve data/BASLATMA_HATASI.txt'ye ekler."""
    print(mesaj)
    yol = os.path.join(VERI_KLASORU, "BASLATMA_HATASI.txt")
    try:
        with open(yol, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {mesaj}\n")
    except OSError:
        pass


def _bozuk_dosyayi_kenara_al(dosya, alanlar):
    """Okunamayan yardımcı dosyayı yeniden adlandırıp boş liste döner.

    Yalnız islem_logu / yuklemeler için kullanılır. Kart verisi bu yolla sıfırlanmaz.
    """
    try:
        return _oku(dosya, alanlar)
    except VeriDogrulamaHatasi as exc:
        if os.path.exists(dosya):
            damga = datetime.now().strftime("%Y%m%d_%H%M%S")
            bozuk = f"{dosya}.bozuk_{damga}"
            try:
                os.replace(dosya, bozuk)
            except OSError:
                bozuk = dosya
            _baslatma_uyarisi_yaz(
                f"UYARI: '{os.path.basename(dosya)}' okunamadı ({exc}). "
                f"Dosya '{os.path.basename(bozuk)}' olarak kenara alındı; "
                "boş liste ile devam ediliyor. Kart verisi etkilenmedi."
            )
        else:
            _baslatma_uyarisi_yaz(
                f"UYARI: '{os.path.basename(dosya)}' okunamadı ({exc}). "
                "Boş liste ile devam ediliyor."
            )
        return []


def kur():
    global _kartlar, _loglar, _yuklemeler

    with _kilit:
        os.makedirs(VERI_KLASORU, exist_ok=True)

        kartlar = _oku(
            KARTLAR_DOSYA,
            KART_ALANLARI,
            ZORUNLU_KART_ALANLARI if os.path.exists(KARTLAR_DOSYA) else frozenset(),
        )
        kartlar = [_kart_normalize(kart) for kart in kartlar]
        if kartlar:
            _kart_listesi_dogrula(kartlar)

        _kartlar = kartlar
        _loglar = _bozuk_dosyayi_kenara_al(LOG_DOSYA, LOG_ALANLARI)
        _yuklemeler = _bozuk_dosyayi_kenara_al(YUKLEME_DOSYA, YUKLEME_ALANLARI)

        eksikler = []
        if not os.path.exists(KARTLAR_DOSYA):
            eksikler.append((KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"))
        if not os.path.exists(LOG_DOSYA):
            eksikler.append((LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"))
        if not os.path.exists(YUKLEME_DOSYA):
            eksikler.append((YUKLEME_DOSYA, YUKLEME_ALANLARI, _yuklemeler, "Yüklemeler"))
        if eksikler:
            _coklu_yaz(eksikler)


def kartlari_diskten_yeniden_yukle():
    """Manuel Excel müdahalesinden sonra kartlar.xlsx'i validate ederek tekrar yükler."""
    global _kartlar

    with _kilit:
        yeni = _oku(KARTLAR_DOSYA, KART_ALANLARI, ZORUNLU_KART_ALANLARI)
        yeni = [_kart_normalize(kart) for kart in yeni]
        _kart_listesi_dogrula(yeni)
        _kartlar = yeni
        return len(_kartlar)


def yeniden_yukle():
    return kartlari_diskten_yeniden_yukle()


def _kartlari_kaydet():
    _gunluk_yedek(KARTLAR_DOSYA)
    _yaz(KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar")


# ---------------------------------------------------------------------------
# Log ve yükleme geçmişi
# ---------------------------------------------------------------------------

def _log_kaydi(kullanici, rol, islem, talep_no="", stok_no="", adet=None, detay=""):
    return {
        "zaman": simdi(),
        "kullanici": kullanici,
        "rol": rol,
        "islem": islem,
        "talep_no": talep_no,
        "stok_no": stok_no,
        "adet": adet,
        "detay": detay,
    }


def _log_arsivle_gerekirse():
    global _loglar

    if len(_loglar) <= LOG_SINIRI:
        return

    os.makedirs(YEDEK_KLASORU, exist_ok=True)
    arsivlenecek = _loglar[:-LOG_SAKLA]
    arsiv = os.path.join(
        YEDEK_KLASORU,
        f"{datetime.now():%Y%m%d_%H%M%S}_islem_logu_arsiv.xlsx",
    )
    _yaz(arsiv, LOG_ALANLARI, arsivlenecek, "İşlem Logu")
    _loglar = _loglar[-LOG_SAKLA:]


def log_ekle(kullanici, rol, islem, talep_no="", stok_no="", adet=None, detay=""):
    global _loglar

    with _kilit:
        eski = copy.deepcopy(_loglar)
        try:
            _loglar.append(
                _log_kaydi(kullanici, rol, islem, talep_no, stok_no, adet, detay)
            )
            _log_arsivle_gerekirse()
            _yaz(LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu")
        except Exception:
            _loglar = eski
            raise


def loglari_getir(adet=None):
    with _kilit:
        secim = list(reversed(_loglar))
        if adet:
            secim = secim[:adet]
        return copy.deepcopy(secim)


def yuklemeleri_getir(adet=None):
    with _kilit:
        secim = list(reversed(_yuklemeler))
        if adet:
            secim = secim[:adet]
        return copy.deepcopy(secim)


# ---------------------------------------------------------------------------
# Görünüm hesapları
# ---------------------------------------------------------------------------

def durum_bilgisi(kart):
    durum = kart.get("durum")
    plan_baslama = kart.get("plan_baslama")
    plan_teslim = kart.get("plan_teslim")
    bugun_iso = bugun()

    bilgi = {
        "rozet": durum or "DURUMU EKSİK",
        "renk": "notr",
        "sapma": None,
        "kalan": None,
        "zaman_yuzde": 0,
        "plan_gun": gun_farki(plan_teslim, plan_baslama),
    }

    if durum is None:
        bilgi["rozet"] = "DURUMU EKSİK"
        bilgi["renk"] = "uyari"
        return bilgi

    if durum == TESLIM_EDILDI:
        teslim = kart.get("gerceklesen_teslim") or str(kart.get("teslim_zamani") or "")[:10]
        sapma = gun_farki(teslim, plan_teslim)
        bilgi["sapma"] = sapma
        bilgi["zaman_yuzde"] = 100
        if sapma is None:
            bilgi["rozet"], bilgi["renk"] = "TESLİM EDİLDİ", "iyi"
        elif sapma > 0:
            bilgi["rozet"], bilgi["renk"] = f"GEÇ TESLİM (+{sapma} gün)", "kotu"
        else:
            bilgi["rozet"], bilgi["renk"] = "ZAMANINDA TESLİM", "iyi"
        return bilgi

    if durum == HAZIR:
        bilgi["rozet"], bilgi["renk"] = "HAZIR · PLANLANMAYI BEKLİYOR", "notr"
        bilgi["zaman_yuzde"] = 0
        return bilgi

    if durum == DIZGIDE:
        if kart.get("tamamlanan_adet", 0) >= kart.get("toplam_adet", 1):
            bilgi["rozet"], bilgi["renk"] = "ÜRETİM BİTTİ · TESLİME HAZIR", "uyari"
            bilgi["zaman_yuzde"] = 100
            return bilgi

        kalan = gun_farki(plan_teslim, bugun_iso)
        bilgi["kalan"] = kalan
        baslangic = str(kart.get("baslama_zamani") or plan_baslama or bugun_iso)[:10]
        gecen = gun_farki(bugun_iso, baslangic) or 0
        plan_gun = bilgi["plan_gun"]

        if plan_gun and plan_gun > 0:
            bilgi["zaman_yuzde"] = max(0, min(140, round(gecen / plan_gun * 100)))
        else:
            bilgi["zaman_yuzde"] = 100 if kalan is not None and kalan < 0 else 50

        if kalan is None:
            bilgi["rozet"], bilgi["renk"] = "DİZGİDE", "uyari"
        elif kalan < 0:
            bilgi["sapma"] = -kalan
            bilgi["rozet"], bilgi["renk"] = f"SÜRE AŞILDI ({-kalan} gün)", "kotu"
        elif kalan <= 1:
            bilgi["rozet"] = "SON GÜN" if kalan == 0 else "SON 1 GÜN"
            bilgi["renk"] = "uyari"
        else:
            bilgi["rozet"], bilgi["renk"] = f"PLANINDA ({kalan} gün var)", "iyi"
        return bilgi

    gecikme = gun_farki(bugun_iso, plan_baslama)
    if plan_baslama and gecikme is not None and gecikme > 0:
        bilgi["rozet"], bilgi["renk"], bilgi["sapma"] = (
            f"BAŞLAMADI (+{gecikme} gün)",
            "kotu",
            gecikme,
        )
    elif plan_baslama and gecikme == 0:
        bilgi["rozet"], bilgi["renk"] = "BUGÜN BAŞLAMALI", "uyari"
    else:
        bilgi["rozet"], bilgi["renk"] = PLANA_ALINDI, "notr"
    return bilgi


def kart_gorunumu(kart):
    d = copy.deepcopy(kart)
    d.update(durum_bilgisi(kart))
    d["toplam_adet"] = d.get("toplam_adet") or 1
    d["tamamlanan_adet"] = d.get("tamamlanan_adet") or 0
    d["baslangic_adet"] = d.get("baslangic_adet") or 0
    d["kalan_adet"] = max(0, d["toplam_adet"] - d["tamamlanan_adet"])
    d["adet_yuzde"] = min(100, round(d["tamamlanan_adet"] / d["toplam_adet"] * 100))
    d["gorunur"] = _operasyonda_gorunur_mu(d)
    d["kaynakta_yok"] = d.get("source_active", 1) != 1
    d["is_durumu"] = d.get("durum") or "DURUMU EKSİK"
    d["kaynak_durumu"] = _temiz_metin(d.get("excel_durum"))
    d["elle_dizgi_mi"] = d.get("dizgi_tipi") == DIZGI_TIPI_ELLE
    d["malzeme_bekliyor"] = d.get("malzeme_bekliyor") == 1
    return d


def _kartlari_sirala(kartlar):
    kartlar.sort(
        key=lambda kart: (
            SIRALAMA.get(kart.get("durum"), 9),
            kart.get("plan_baslama") or "9999-12-31",
            kart.get("sira") or 999999,
            kart.get("id") or 0,
        )
    )
    return kartlar


def kartlari_getir(sadece_gorunen=True):
    with _kilit:
        secim = [
            kart for kart in _kartlar
            if not sadece_gorunen or _operasyonda_gorunur_mu(kart)
        ]
        kartlar = [kart_gorunumu(kart) for kart in secim]
    return _kartlari_sirala(kartlar)


def kartlari_yonetim_getir():
    """Admin tablosu için, durumu boş kartlar dahil, gizlenmemiş tüm aktif kartlar."""
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if _yonetimde_gorunur_mu(kart)
        ]
    return _kartlari_sirala(kartlar)


def durumu_eksik_kartlari_getir():
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if _yonetimde_gorunur_mu(kart) and not kart.get("durum")
        ]
    return _kartlari_sirala(kartlar)


def gizlenen_kartlari_getir():
    with _kilit:
        kartlar = [
            kart_gorunumu(kart)
            for kart in _kartlar
            if kart.get("aktif", 1) == 1 and kart.get("admin_gizli", 0) == 1
        ]
    kartlar.sort(key=lambda kart: (kart.get("guncelleme") or "", kart.get("id") or 0), reverse=True)
    return kartlar


def kart_getir(kart_id):
    kart_id = _sayi(kart_id, -1)
    with _kilit:
        kart = _kart_ref(kart_id)
        return kart_gorunumu(kart) if kart else None


def kart_bul(anahtar):
    with _kilit:
        for kart in _kartlar:
            if kart.get("anahtar") == anahtar:
                return copy.deepcopy(kart)
    return None


def yeni_kimlik():
    with _kilit:
        return max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1


# ---------------------------------------------------------------------------
# Kart + log commit yardımcıları
# ---------------------------------------------------------------------------

def _kart_log_commit():
    _gunluk_yedek(KARTLAR_DOSYA)
    _coklu_yaz(
        [
            (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
            (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
        ]
    )


def _atomik_kart_islemi(islem):
    """RAM state rollback kalıbını tek yerde tutar.

    Kartlar in-place mutasyon gördüğü için deepcopy zorunlu.
    Loglara yalnız append yapıldığı için uzunluk + del yeterlidir.
    """
    global _kartlar

    eski_kartlar = copy.deepcopy(_kartlar)
    log_sayisi = len(_loglar)
    try:
        sonuc = islem()
        _kart_listesi_dogrula(_kartlar)
        _kart_log_commit()
        return sonuc
    except Exception:
        _kartlar = eski_kartlar
        del _loglar[log_sayisi:]
        raise


# ---------------------------------------------------------------------------
# Operatör workflow işlemleri
# ---------------------------------------------------------------------------

def kart_baslat(kart_id, adet, kullanici, rol, operator_tipi, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not _operasyonda_gorunur_mu(kart):
            raise IsKuralHatasi("Bu kart operasyon ekranında aktif değil.")
        if not operator_kart_yetkisi_var_mi(kart, operator_tipi):
            raise IsKuralHatasi(
                "Bu kart sizin operatör tipiniz için uygun değil."
            )
        
        if kart["durum"] != PLANA_ALINDI:
            raise IsKuralHatasi("Yalnız PLANA ALINDI durumundaki kart DİZGİDE'ye alınabilir.")

        if adet in (None, ""):
            adet = kart["toplam_adet"]
        try:
            adet = int(adet)
        except (TypeError, ValueError) as exc:
            raise ValueError("Adet sayı olmalı.") from exc
        if adet < 1 or adet > kart["toplam_adet"]:
            raise IsKuralHatasi(f"Başlatılacak adet 1 ile {kart['toplam_adet']} arasında olmalı.")

        def islem():
            kart.update(
                durum=DIZGIDE,
                baslangic_adet=adet,
                baslama_zamani=kart.get("baslama_zamani") or simdi(),
                bitis_zamani=None,
                teslim_zamani=None,
                gerceklesen_teslim=None,
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=simdi(),
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "DİZGİYE ALINDI",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    adet,
                    aciklama or f"{adet} adet dizgiye alındı",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def kart_bitir(kart_id, adet, kullanici, rol, operator_tipi, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not operator_kart_yetkisi_var_mi(kart, operator_tipi):
            raise IsKuralHatasi(
                "Bu kart sizin operatör tipiniz için uygun değil."
            )
        if kart.get("durum") != DIZGIDE:
            raise IsKuralHatasi("Tamamlanan adet yalnız DİZGİDE durumundaki karta girilebilir.")

        kalan = kart["toplam_adet"] - kart["tamamlanan_adet"]
        if kalan <= 0:
            raise IsKuralHatasi("Üretim adedi zaten tamamlandı. Kartı HAZIR durumuna alın.")

        try:
            adet = int(adet)
        except (TypeError, ValueError) as exc:
            raise ValueError("Adet sayı olmalı.") from exc
        if adet < 1 or adet > kalan:
            raise IsKuralHatasi(f"Adet 1 ile {kalan} arasında olmalı.")

        def islem():
            yeni_toplam = kart["tamamlanan_adet"] + adet
            uretim_bitti = yeni_toplam == kart["toplam_adet"]

            kart.update(
                tamamlanan_adet=yeni_toplam,
                bitis_zamani=simdi() if uretim_bitti else kart.get("bitis_zamani"),
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=simdi(),
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "ÜRETİM ADEDİ TAMAMLANDI" if uretim_bitti else "KISMİ ÜRETİM",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    adet,
                    aciklama or f"{yeni_toplam}/{kart['toplam_adet']} adet tamamlandı",
                )
            )

            mesaj = (
                "Üretim adedi tamamlandı. Hazır olduğunuzda kartı Teslim Edildi olarak işaretleyebilirsiniz."
                if uretim_bitti
                else f"{yeni_toplam}/{kart['toplam_adet']} adet tamamlandı."
            )

            return kart_gorunumu(kart), uretim_bitti, mesaj

        return _atomik_kart_islemi(islem)



def kart_teslim_et(kart_id, kullanici, rol, operator_tipi, aciklama=""):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if not operator_kart_yetkisi_var_mi(kart, operator_tipi):
            raise IsKuralHatasi(
                "Bu kart sizin operatör tipiniz için uygun değil."
            )
        if kart.get("durum") != DIZGIDE:
            raise IsKuralHatasi("Yalnız DİZGİDE durumundaki kart TESLİM EDİLDİ yapılabilir.")
        if kart.get("tamamlanan_adet", 0) != kart.get("toplam_adet", 0):
            raise IsKuralHatasi(
                "Kart teslim edilmeden önce üretim adedinin tamamı bitirilmelidir."
            )

        def islem():
            teslim_ani = simdi()
            kart.update(
                durum=TESLIM_EDILDI,
                gerceklesen_teslim=bugun(),
                teslim_zamani=teslim_ani,
                bitis_zamani=kart.get("bitis_zamani") or teslim_ani,
                operator=kullanici,
                aciklama=aciklama or kart.get("aciklama"),
                guncelleme=teslim_ani,
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "TESLİM EDİLDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    kart.get("toplam_adet"),
                    aciklama or f"Teslim tarihi: {bugun()}",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)
    



def kart_not_guncelle(kart_id, aciklama, kullanici, rol):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        def islem():
            kart["aciklama"] = _temiz_metin(aciklama) or None
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    rol,
                    "NOT GÜNCELLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay=kart.get("aciklama") or "Not temizlendi",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Admin kart işlemleri
# ---------------------------------------------------------------------------

def _tarih_form_degeri(deger, alan_adi):
    if deger in (None, ""):
        return None
    sonuc = tarih_coz(deger)
    if not sonuc:
        raise IsKuralHatasi(f"{alan_adi} geçerli bir tarih olmalı.")
    return sonuc


def admin_kart_ekle(
    talep_no,
    stok_no,
    toplam_adet,
    kullanici,
    sira=None,
    talep_sahibi="",
    plan_hafta="",
    plan_baslama=None,
    plan_teslim=None,
    gerceklesen_teslim=None,
    pcb="",
    aciklama="",
    elle_dizgi=False,
    dizgi_sorumlusu="",
):
    """Admin panelinden PLANA ALINDI durumunda manuel kart oluşturur."""
    global _kartlar, _loglar

    talep_no = _temiz_metin(talep_no)
    stok_no = _temiz_metin(stok_no)
    if not talep_no:
        raise IsKuralHatasi("Talep NO boş olamaz.")
    if not stok_no:
        raise IsKuralHatasi("Kart Stok No boş olamaz.")

    try:
        toplam = int(toplam_adet)
    except (TypeError, ValueError) as exc:
        raise ValueError("Toplam adet sayı olmalı.") from exc
    if toplam < 1:
        raise IsKuralHatasi("Toplam adet en az 1 olmalı.")

    if sira in (None, ""):
        sira_degeri = None
    else:
        try:
            sira_degeri = int(sira)
        except (TypeError, ValueError) as exc:
            raise ValueError("Sıra tam sayı olmalı.") from exc

    plan_baslama_iso = _tarih_form_degeri(plan_baslama, "Dizgi Başlama Tarihi")
    plan_teslim_iso = _tarih_form_degeri(plan_teslim, "Planlanan Teslim Tarihi")
    gerceklesen_iso = _tarih_form_degeri(gerceklesen_teslim, "Gerçekleşen Teslim Tarihi")

    if plan_baslama_iso and plan_teslim_iso and plan_baslama_iso > plan_teslim_iso:
        raise IsKuralHatasi("Dizgi Başlama Tarihi Planlanan Teslim Tarihinden sonra olamaz.")
    if gerceklesen_iso:
        raise IsKuralHatasi(
            "Yeni kart PLANA ALINDI durumunda başlar; Gerçekleşen Teslim Tarihi başlangıçta boş olmalı."
        )

    anahtar = f"{talep_no}|{stok_no}"

    with _kilit:
        if any(kart.get("anahtar") == anahtar for kart in _kartlar):
            raise IsKuralHatasi(f"Bu Talep NO + Kart Stok No zaten mevcut: {anahtar}")

        def islem():
            kart_id = max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1
            kayit = {
                "id": kart_id,
                "sira": sira_degeri,
                "talep_no": talep_no,
                "stok_no": stok_no,
                "talep_sahibi": _temiz_metin(talep_sahibi) or None,
                "toplam_adet": toplam,
                "adet_metin": f"{toplam} ADET",
                "plan_hafta": _temiz_metin(plan_hafta) or None,
                "plan_baslama": plan_baslama_iso,
                "plan_teslim": plan_teslim_iso,
                "gerceklesen_teslim": None,
                "excel_durum": "MANUEL",
                "pcb": _temiz_metin(pcb) or None,
                "dizgi_tipi": DIZGI_TIPI_ELLE if elle_dizgi else DIZGI_TIPI_MAKINE,
                "dizgi_sorumlusu": _temiz_metin(dizgi_sorumlusu) or None,
                "malzeme_bekliyor": 0,
                "durum": PLANA_ALINDI,
                "baslangic_adet": 0,
                "tamamlanan_adet": 0,
                "baslama_zamani": None,
                "bitis_zamani": None,
                "teslim_zamani": None,
                "operator": None,
                "aciklama": _temiz_metin(aciklama) or None,
                "guncelleme": simdi(),
                "aktif": 1,
                "source_active": 1,
                "admin_gizli": 0,
                "kaynak": "MANUEL",
                "anahtar": anahtar,
            }
            _kartlar.append(kayit)
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "MANUEL KART EKLENDİ",
                    talep_no,
                    stok_no,
                    toplam,
                    f"{toplam} adet · Durum: {PLANA_ALINDI}"
                    + (" · Elle Dizgi" if elle_dizgi else ""),
                )
            )
            return kart_gorunumu(kayit)

        return _atomik_kart_islemi(islem)


def admin_kart_duzenle(
    kart_id,
    durum,
    tamamlanan_adet,
    toplam_adet,
    aciklama,
    kullanici,
    plan_hafta=None,
    plan_baslama=None,
    plan_teslim=None,
    gerceklesen_teslim=None,
    elle_dizgi=None,
    dizgi_sorumlusu=None,
):
    """Admin workflow ve plan alanlarını kontrollü biçimde düzeltir.

    elle_dizgi / dizgi_sorumlusu: None verilirse mevcut değer korunur
    (plan_hafta vb. alanlarla aynı "None = değiştirme" kuralı).
    """
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        yeni_durum = _durum_normalize(durum)
        if yeni_durum not in GECERLI_DURUMLAR:
            raise IsKuralHatasi("Durum PLANA ALINDI, DİZGİDE, HAZIR veya TESLİM EDİLDİ olmalı.")

        try:
            toplam = int(kart["toplam_adet"] if toplam_adet in (None, "") else toplam_adet)
            tamamlanan = int(
                kart["tamamlanan_adet"] if tamamlanan_adet in (None, "") else tamamlanan_adet
            )
        except (TypeError, ValueError) as exc:
            raise ValueError("Adetler sayı olmalı.") from exc

        if toplam < 1 or tamamlanan < 0 or tamamlanan > toplam:
            raise IsKuralHatasi("Adet değerleri tutarsız.")

        yeni_plan_hafta = (
            kart.get("plan_hafta")
            if plan_hafta is None
            else (_temiz_metin(plan_hafta) or None)
        )
        yeni_plan_baslama = (
            kart.get("plan_baslama")
            if plan_baslama is None
            else _tarih_form_degeri(plan_baslama, "Dizgi Başlama Tarihi")
        )
        yeni_plan_teslim = (
            kart.get("plan_teslim")
            if plan_teslim is None
            else _tarih_form_degeri(plan_teslim, "Planlanan Teslim Tarihi")
        )
        yeni_gerceklesen = (
            kart.get("gerceklesen_teslim")
            if gerceklesen_teslim is None
            else _tarih_form_degeri(gerceklesen_teslim, "Gerçekleşen Teslim Tarihi")
        )

        if yeni_plan_baslama and yeni_plan_teslim and yeni_plan_baslama > yeni_plan_teslim:
            raise IsKuralHatasi("Dizgi Başlama Tarihi Planlanan Teslim Tarihinden sonra olamaz.")

        yeni_dizgi_tipi = (
            kart.get("dizgi_tipi")
            if elle_dizgi is None
            else (DIZGI_TIPI_ELLE if elle_dizgi else DIZGI_TIPI_MAKINE)
        )
        yeni_dizgi_sorumlusu = (
            kart.get("dizgi_sorumlusu")
            if dizgi_sorumlusu is None
            else (_temiz_metin(dizgi_sorumlusu) or None)
        )

        onceki_durum = kart.get("durum") or "DURUMU EKSİK"

        def islem():
            nonlocal tamamlanan, yeni_gerceklesen

            baslama = kart.get("baslama_zamani")
            bitis = kart.get("bitis_zamani")
            teslim_zamani = kart.get("teslim_zamani")
            baslangic_adet = kart.get("baslangic_adet") or 0

            if yeni_durum in (PLANA_ALINDI, HAZIR):
                tamamlanan = 0
                baslangic_adet = 0
                baslama = None
                bitis = None
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == DIZGIDE:
                baslama = baslama or simdi()
                baslangic_adet = baslangic_adet or toplam
                bitis = bitis if tamamlanan == toplam else None
                teslim_zamani = None
                yeni_gerceklesen = None

            elif yeni_durum == TESLIM_EDILDI:
                tamamlanan = toplam
                baslama = baslama or simdi()
                bitis = bitis or simdi()
                yeni_gerceklesen = yeni_gerceklesen or bugun()
                teslim_zamani = teslim_zamani or simdi()

            kart.update(
                durum=yeni_durum,
                toplam_adet=toplam,
                tamamlanan_adet=tamamlanan,
                baslangic_adet=baslangic_adet,
                plan_hafta=yeni_plan_hafta,
                plan_baslama=yeni_plan_baslama,
                plan_teslim=yeni_plan_teslim,
                gerceklesen_teslim=yeni_gerceklesen,
                baslama_zamani=baslama,
                bitis_zamani=bitis,
                teslim_zamani=teslim_zamani,
                aciklama=_temiz_metin(aciklama) if aciklama is not None else kart.get("aciklama"),
                dizgi_tipi=yeni_dizgi_tipi,
                dizgi_sorumlusu=yeni_dizgi_sorumlusu,
                guncelleme=simdi(),
            )

            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "ADMİN DÜZENLEDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    tamamlanan,
                    f"{onceki_durum} → {yeni_durum} · {tamamlanan}/{toplam} adet",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


def admin_kart_gizle(kart_id, kullanici):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")

        def islem():
            kart["admin_gizli"] = 1
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "KART LİSTEDEN GİZLENDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay="Kart silinmedi; admin gizli olarak işaretlendi.",
                )
            )

        _atomik_kart_islemi(islem)


def admin_kart_geri_getir(kart_id, kullanici):
    kart_id = _sayi(kart_id, -1)

    with _kilit:
        kart = _kart_ref(kart_id)
        if not kart:
            raise KartBulunamadi("Kart bulunamadı.")
        if kart.get("admin_gizli", 0) != 1:
            raise IsKuralHatasi("Bu kart zaten görünür durumda.")

        def islem():
            kart["admin_gizli"] = 0
            kart["aktif"] = 1
            kart["guncelleme"] = simdi()
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "GİZLENEN KART GERİ GETİRİLDİ",
                    kart.get("talep_no") or "",
                    kart.get("stok_no") or "",
                    detay="Kart tekrar yönetim/operasyon listelerine alındı.",
                )
            )
            return kart_gorunumu(kart)

        return _atomik_kart_islemi(islem)


# ---------------------------------------------------------------------------
# Yedekten geri yükleme
# ---------------------------------------------------------------------------

def yedekten_geri_yukle(yedek_adi, kullanici):
    global _kartlar, _loglar

    with _kilit:
        aday = _yedek_kart_dosyasi_bul(yedek_adi)
        yeni_kartlar = _oku(aday, KART_ALANLARI, ZORUNLU_KART_ALANLARI)
        yeni_kartlar = [_kart_normalize(kart) for kart in yeni_kartlar]
        if not yeni_kartlar:
            raise VeriDogrulamaHatasi("Seçilen yedekte hiç kart bulunmuyor.")
        _kart_listesi_dogrula(yeni_kartlar)

        koruma_yedegi = anlik_yedek("geri_yukleme_oncesi")
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)

        try:
            _kartlar = yeni_kartlar
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "YEDEKTEN GERİ YÜKLENDİ",
                    detay=(
                        f"Yedek: {yedek_adi} · {len(_kartlar)} kart · "
                        f"koruma: {os.path.basename(koruma_yedegi)}"
                    ),
                )
            )
            _coklu_yaz(
                [
                    (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
                    (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
                ]
            )
        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            raise

        return {
            "kart": len(_kartlar),
            "yedek": yedek_adi,
            "koruma_yedegi": koruma_yedegi,
        }


# ---------------------------------------------------------------------------
# Excel import commit
# ---------------------------------------------------------------------------

def excel_import_uygula(dosya_adi, kullanici, satirlar, uyari_sayisi=0):
    """Tamamen parse/validate edilmiş kaynak satırlarını tek transaction-benzeri blokta uygular."""
    global _kartlar, _loglar, _yuklemeler

    with _kilit:
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)
        eski_yuklemeler = copy.deepcopy(_yuklemeler)

        try:
            yedek_klasoru = anlik_yedek("import_oncesi")
            mevcut_harita = {kart.get("anahtar"): kart for kart in _kartlar}
            gorulen = set()
            yeni = 0
            guncellenen = 0
            workflow_korundu = 0
            elle_dizgi_satir = sum(
                1
                for satir in satirlar
                if satir.get("plan", {}).get("dizgi_tipi") == DIZGI_TIPI_ELLE
            )
            pasife_listesi = []
            sonraki_id = max([_sayi(kart.get("id"), 0) for kart in _kartlar] or [0]) + 1

            for satir in satirlar:
                anahtar = satir["anahtar"]
                gorulen.add(anahtar)
                plan = satir["plan"]
                mevcut = mevcut_harita.get(anahtar)

                if mevcut:
                    yeni_toplam = int(plan["toplam_adet"])
                    tamamlanan = int(mevcut.get("tamamlanan_adet") or 0)
                    if yeni_toplam < tamamlanan:
                        raise IsKuralHatasi(
                            f"{anahtar}: Excel toplam adedi ({yeni_toplam}), sistemde tamamlanan "
                            f"adetten ({tamamlanan}) küçük olamaz. Import iptal edildi."
                        )
                    if mevcut.get("durum") == TESLIM_EDILDI and yeni_toplam != tamamlanan:
                        raise IsKuralHatasi(
                            f"{anahtar}: Kart {mevcut['durum']} ve {tamamlanan} adet tamamlanmış. "
                            f"Yeni Excel toplam adedi {yeni_toplam}; admin kontrolü gerekir."
                        )

                    workflow = {
                        "durum": mevcut.get("durum"),
                        "baslangic_adet": mevcut.get("baslangic_adet"),
                        "tamamlanan_adet": tamamlanan,
                        "baslama_zamani": mevcut.get("baslama_zamani"),
                        "bitis_zamani": mevcut.get("bitis_zamani"),
                        "teslim_zamani": mevcut.get("teslim_zamani"),
                        "operator": mevcut.get("operator"),
                        "aciklama": mevcut.get("aciklama"),
                        "gerceklesen_teslim": mevcut.get("gerceklesen_teslim"),
                    }

                    mevcut.update(plan)
                    mevcut.update(workflow)
                    mevcut["source_active"] = 1
                    mevcut["aktif"] = 1
                    mevcut["kaynak"] = "EXCEL"
                    mevcut["guncelleme"] = simdi()
                    _kart_dogrula(mevcut)
                    guncellenen += 1
                    workflow_korundu += 1
                    continue

                toplam = int(plan["toplam_adet"])
                durum = satir.get("ilk_durum")
                gerceklesen = satir.get("gerceklesen_teslim")

                tamamlanan = toplam if durum == TESLIM_EDILDI else 0
                baslangic_adet = toplam if durum in (DIZGIDE, TESLIM_EDILDI) else 0

                baslama = None
                bitis = None
                teslim_zamani = None

                kayit = {
                    "id": sonraki_id,
                    "anahtar": anahtar,
                    "talep_no": satir["talep_no"],
                    "stok_no": satir["stok_no"],
                    "gerceklesen_teslim": (
                        gerceklesen
                        if durum == TESLIM_EDILDI
                        else None
                    ),
                    "durum": durum,
                    "baslangic_adet": baslangic_adet,
                    "tamamlanan_adet": tamamlanan,
                    "baslama_zamani": None,
                    "bitis_zamani": None,
                    "teslim_zamani": None,
                    "operator": "Excel" if durum in GECERLI_DURUMLAR else None,
                    "aciklama": None,
                    "guncelleme": simdi(),
                    "aktif": 1,
                    "source_active": 1,
                    "admin_gizli": 0,
                    "kaynak": "EXCEL",
                }
                kayit.update(plan)
                _kart_dogrula(kayit)
                _kartlar.append(kayit)
                mevcut_harita[anahtar] = kayit
                sonraki_id += 1
                yeni += 1

            pasife_alinan = 0
            for kart in _kartlar:
                if kart.get("kaynak") != "EXCEL" or kart.get("anahtar") in gorulen:
                    continue
                if kart.get("source_active", 1) == 1:
                    kart["source_active"] = 0
                    kart["guncelleme"] = simdi()
                    pasife_alinan += 1
                    pasife_listesi.append(
                        f"{kart.get('talep_no') or ''}|{kart.get('stok_no') or ''}"
                    )

            _kart_listesi_dogrula(_kartlar)

            _yuklemeler.append(
                {
                    "zaman": simdi(),
                    "kullanici": kullanici,
                    "dosya": dosya_adi,
                    "satir": len(satirlar),
                    "yeni": yeni,
                    "guncellenen": guncellenen,
                    "pasife_alinan": pasife_alinan,
                    "uyari": uyari_sayisi,
                }
            )
            _loglar.append(
                _log_kaydi(
                    kullanici,
                    "admin",
                    "EXCEL YÜKLENDİ",
                    detay=(
                        f"{len(satirlar)} satır · {yeni} yeni · {guncellenen} güncellendi · "
                        f"{workflow_korundu} workflow korundu · {pasife_alinan} kaynakta yok · "
                        f"{uyari_sayisi} uyarı · yedek={os.path.basename(yedek_klasoru)}"
                    ),
                )
            )

            _gunluk_yedek(KARTLAR_DOSYA)
            _coklu_yaz(
                [
                    (KARTLAR_DOSYA, KART_ALANLARI, _kartlar, "Kartlar"),
                    (LOG_DOSYA, LOG_ALANLARI, _loglar, "İşlem Logu"),
                    (YUKLEME_DOSYA, YUKLEME_ALANLARI, _yuklemeler, "Yüklemeler"),
                ]
            )

            return {
                "satir": len(satirlar),
                "yeni": yeni,
                "guncellenen": guncellenen,
                "pasife_alinan": pasife_alinan,
                "pasife_listesi": pasife_listesi[:20],
                "workflow_korundu": workflow_korundu,
                "uyari": uyari_sayisi,
                "elle_dizgi_satir": elle_dizgi_satir,
                "yedek": yedek_klasoru,
            }

        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            _yuklemeler = eski_yuklemeler
            raise


# ---------------------------------------------------------------------------
# Legacy küçük yardımcılar
# ---------------------------------------------------------------------------

def kart_guncelle(kart_id, **alanlar):
    global _kartlar

    kart_id = _sayi(kart_id, -1)

    with _kilit:
        index = next(
            (
                i
                for i, kart in enumerate(_kartlar)
                if kart.get("id") == kart_id
            ),
            None,
        )

        if index is None:
            return None

        eski = copy.deepcopy(_kartlar)

        try:
            yeni = copy.deepcopy(_kartlar[index])
            yeni.update(alanlar)
            yeni["guncelleme"] = simdi()

            if "talep_no" in alanlar or "stok_no" in alanlar:
                yeni["anahtar"] = (
                    f"{_temiz_metin(yeni.get('talep_no'))}|"
                    f"{_temiz_metin(yeni.get('stok_no'))}"
                )

            yeni = _kart_normalize(yeni)
            _kartlar[index] = yeni

            _kart_listesi_dogrula(_kartlar)
            _kartlari_kaydet()

            return kart_gorunumu(yeni)

        except Exception:
            _kartlar = eski
            raise


def toplu_kaydet(yeni_kartlar=(), degisen=True):
    global _kartlar

    with _kilit:
        eski = copy.deepcopy(_kartlar)
        try:
            _kartlar.extend(_kart_normalize(kart) for kart in yeni_kartlar)
            _kart_listesi_dogrula(_kartlar)
            if degisen:
                _kartlari_kaydet()
        except Exception:
            _kartlar = eski
            raise
```


## `excel_araclari.py`


```python
"""PDGM kaynak Excel importu ve rapor üretimi.

Kaynak import kuralları:
- MAKİNE sayfası (zorunlu) + ELDE DİZGİ YENİ sayfası (opsiyonel) kullanılır.
  İki sayfa TEK import işleminde birlikte işlenir: aksi halde bir sayfayı
  import etmek diğer sayfadan gelen kartları "kaynakta yok" (pasif) yapardı.
- Hidden kolonlar alınmaz; hidden satırlar korunur.
- Microsoft Excel COM ile values-only snapshot oluşturulur.
- External link güncellemesi ve full recalculation yapılmaz.
- Talep NO + Kart Stok No olmayan satırlar kart sayılmaz.
- Aynı Talep NO + Kart Stok No birden fazla satırda olabilir (aynı talebin
  farklı tarihte teslim edilen ikinci/üçüncü partisi gibi). İlk görülen satır
  eski anahtar biçimini (talep_no|stok_no) korur; sonraki partiler o partiyi
  ayıran tarihe göre (mümkünse) okunabilir bir ek alır. Bkz. _parti_anahtari_uret.
- Workflow yalnız şu dört durumdan oluşur:
    PLANA ALINDI, DİZGİDE, HAZIR, TESLİM EDİLDİ
  ELDE DİZGİ YENİ sayfasındaki beş durum (MALZEME TEDARİK, DİZGİ İÇİN
  BEKLİYOR, DİZGİDE, DİZGİDE VE MALZEME BEKLİYOR, TESLİM EDİLDİ) bu dört
  duruma indirgenir; ham metin excel_durum alanında saklanır ve kaybolmaz.
- Kartın MAKİNE mi yoksa ELDE DİZGİ YENİ sayfasından mı geldiği "dizgi_tipi"
  alanında yapısal olarak (hangi sheet'ten okunduğuna göre) damgalanır;
  DURUM metninden tahmin edilmez.
- DURUM boş veya farklı bir değer ise kart kaybolmaz; durum None olur ve
  operasyon ekranında gösterilmez. Admin Yönetim ekranında düzeltebilir.
"""

from __future__ import annotations

import gc
import io
import os
import re
import tempfile
import threading
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

import depo

_import_kilidi = threading.Lock()


class ExcelAktarimHatasi(Exception):
    pass


KAYNAK_SAYFA_ADI = "MAKİNE"
KAYNAK_SAYFA_ADI_ELLE = "ELDE DİZGİ YENİ"
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# (sayfa adı, dizgi_tipi, zorunlu mu) — import bu listedeki sayfaları sırayla dener.
# MAKİNE zorunludur (yoksa hata); ELDE DİZGİ YENİ opsiyoneldir (eski/sade
# Excel dosyalarıyla geriye dönük uyumluluk için — yoksa sessizce atlanır).
KAYNAK_SAYFALARI = (
    (KAYNAK_SAYFA_ADI, depo.DIZGI_TIPI_MAKINE, True),
    (KAYNAK_SAYFA_ADI_ELLE, depo.DIZGI_TIPI_ELLE, False),
)


TURKCE_HARFLER = str.maketrans(
    {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
        "â": "a",
        "Â": "A",
    }
)


def _sadelestir(deger):
    metin = str(deger or "").translate(TURKCE_HARFLER).upper()
    metin = metin.replace(".", " ").replace("_", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", metin).strip()


BASLIK_ESLESME = {
    _sadelestir(baslik): alan
    for baslik, alan in {
        "NO": "sira",
        "Talep NO": "talep_no",
        "Talep Sahibi": "talep_sahibi",
        "Kart Stok No": "stok_no",
        "Kart Üretim Adet": "adet_metin",
        "Planlanan Başlangıç T.": "plan_hafta",
        "Dizgi Başlama Tarihi": "plan_baslama",
        "Planlanan Teslim T.": "plan_teslim",
        "Gerçekleşen Teslim T.": "gerceklesen_teslim",
        "DURUM": "excel_durum",
        "PCB": "pcb",
        # Küçük format toleransı
        "Sıra": "sira",
        "Üretim Adet": "adet_metin",
        "Adet": "adet_metin",
        # ELDE DİZGİ YENİ sayfasına özel kolonlar / farklı adlandırmalar.
        # Aynı sözlük her iki sayfa için de kullanılır; bir sayfanın başlık
        # satırında bulunmayan bir eşleşme basitçe kolonlar sözlüğüne girmez.
        "PDGM Dizgi Sorumlusu": "dizgi_sorumlusu",
        "Dizgi Başlama T.": "plan_baslama",
        "MALZEME/PCB": "pcb",
    }.items()
}

DURUM_ESLESME = {
    "PLANA ALINDI": depo.PLANA_ALINDI,
    "DIZGIDE": depo.DIZGIDE,
    "HAZIR": depo.HAZIR,
    "TESLIM EDILDI": depo.TESLIM_EDILDI,
    # ELDE DİZGİ YENİ sayfasının beş durumu, dört gerçek workflow durumuna
    # indirgenir. Ham metin her zaman excel_durum alanında ayrıca saklanır,
    # bu yüzden bu indirgeme bilgi kaybı yaratmaz.
    "MALZEME TEDARIK": depo.PLANA_ALINDI,
    "DIZGI ICIN BEKLIYOR": depo.PLANA_ALINDI,
    "DIZGIDE VE MALZEME BEKLIYOR": depo.DIZGIDE,
}

# Bu iki ham durum, kart bir workflow durumuna girmiş olsa bile fiilen
# malzeme eksikliğinden ilerleyemediğini ifade eder. Operatör ekranı bu
# bayrağa bakarak "Dizgiye Al" aksiyonunu gizler / uyarı gösterir.
MALZEME_BEKLEYEN_DURUMLAR = {
    "MALZEME TEDARIK",
    "DIZGIDE VE MALZEME BEKLIYOR",
}


# ---------------------------------------------------------------------------
# COM snapshot
# ---------------------------------------------------------------------------

def _sayfa_kopyala(kaynak_ws, hedef_ws):
    """Bir kaynak sheet'in hidden olmayan kolonlarını, hidden satırları koruyarak hedefe kopyalar."""
    used = kaynak_ws.UsedRange
    ilk_satir = used.Row
    son_satir = used.Row + used.Rows.Count - 1
    ilk_sutun = used.Column
    son_sutun = used.Column + used.Columns.Count - 1
    hedef_sutun = 1

    for kaynak_sutun in range(ilk_sutun, son_sutun + 1):
        if bool(kaynak_ws.Columns(kaynak_sutun).Hidden):
            continue

        kaynak_aralik = kaynak_ws.Range(
            kaynak_ws.Cells(ilk_satir, kaynak_sutun),
            kaynak_ws.Cells(son_satir, kaynak_sutun),
        )
        hedef_aralik = hedef_ws.Range(
            hedef_ws.Cells(ilk_satir, hedef_sutun),
            hedef_ws.Cells(son_satir, hedef_sutun),
        )
        hedef_aralik.Value2 = kaynak_aralik.Value2
        hedef_sutun += 1


def excel_deger_snapshot_olustur(dosya_yolu):
    """MAKİNE (zorunlu) ve ELDE DİZGİ YENİ (varsa) sheet'lerinin visible

    kolonlarını values-only geçici xlsx'e kopyalar. Her iki sayfa da tek
    snapshot dosyasında, kaynak adlarıyla aynı isimde ayrı sheet olarak yer
    alır; böylece _excelden_aktar tek workbook üzerinden ikisini de okuyup
    tek import işleminde birleştirebilir.
    """
    if pythoncom is None or win32com is None:
        raise ExcelAktarimHatasi(
            "Excel COM aktarımı yalnızca Windows + Microsoft Excel ortamında çalışır "
            "(pywin32 gerekli)."
        )

    kaynak = Path(dosya_yolu).resolve()
    temp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    temp.close()
    hedef = Path(temp.name).resolve()

    pythoncom.CoInitialize()
    excel = None
    kaynak_wb = None
    hedef_wb = None
    basarili = False

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        kaynak_wb = excel.Workbooks.Open(
            str(kaynak),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        def _sayfa_bul(sayfa_adi):
            return next(
                (
                    ws
                    for ws in kaynak_wb.Worksheets
                    if _sadelestir(ws.Name) == _sadelestir(sayfa_adi)
                ),
                None,
            )

        hedef_wb = excel.Workbooks.Add()
        while hedef_wb.Worksheets.Count > 1:
            hedef_wb.Worksheets(hedef_wb.Worksheets.Count).Delete()

        ilk_sayfa_kullanildi = False
        for sayfa_adi, _dizgi_tipi, zorunlu in KAYNAK_SAYFALARI:
            kaynak_ws = _sayfa_bul(sayfa_adi)
            if kaynak_ws is None:
                if zorunlu:
                    sayfalar = [ws.Name for ws in kaynak_wb.Worksheets]
                    raise ExcelAktarimHatasi(
                        f"'{sayfa_adi}' sayfası bulunamadı. "
                        f"Dosyadaki sayfalar: {', '.join(sayfalar)}"
                    )
                continue  # opsiyonel sayfa yoksa sessizce atla

            if not ilk_sayfa_kullanildi:
                hedef_ws = hedef_wb.Worksheets(1)
                hedef_ws.Name = sayfa_adi
                ilk_sayfa_kullanildi = True
            else:
                hedef_ws = hedef_wb.Worksheets.Add(
                    After=hedef_wb.Worksheets(hedef_wb.Worksheets.Count)
                )
                hedef_ws.Name = sayfa_adi

            _sayfa_kopyala(kaynak_ws, hedef_ws)

        hedef_wb.SaveAs(str(hedef), FileFormat=51)
        basarili = True
        return str(hedef)

    finally:
        if hedef_wb is not None:
            try:
                hedef_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if kaynak_wb is not None:
            try:
                kaynak_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        kaynak_wb = None
        hedef_wb = None
        excel = None
        gc.collect()

        pythoncom.CoUninitialize()

        if not basarili:
            try:
                os.remove(str(hedef))
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Parser helpers
# ---------------------------------------------------------------------------

def durum_coz(deger):
    """Kaynak durumunu dört gerçek workflow durumundan birine çevirir."""
    temiz = _sadelestir(deger)
    if not temiz:
        return None, True
    durum = DURUM_ESLESME.get(temiz)
    return durum, durum is None


def malzeme_bekliyor_mu(deger):
    """Ham DURUM metni, kartın malzeme eksikliğinden ilerleyemediğini mi söylüyor?"""
    return _sadelestir(deger) in MALZEME_BEKLEYEN_DURUMLAR


def adet_coz(deger):
    """Üretim adedini pozitif tam sayıya çevirir.

    Kabul:
    - 400
    - 400.0
    - "400 ADET"
    - "1.500 ADET"
    - "1,500 ADET"
    - "1 500 ADET"

    Red:
    - 400.5
    - "400.5 ADET"
    - "400,5 ADET"
    - "-5 ADET"
    - "400 / 500 ADET"
    """
    if deger is None or str(deger).strip() == "":
        raise ExcelAktarimHatasi("Üretim adedi boş olamaz.")

    if isinstance(deger, bool):
        raise ExcelAktarimHatasi("Üretim adedi sayı olmalı.")

    if isinstance(deger, (int, float)):
        if isinstance(deger, float) and not deger.is_integer():
            raise ExcelAktarimHatasi(
                f"Üretim adedi tam sayı olmalı: {deger}"
            )
        adet = int(deger)

    else:
        metin = str(deger).strip()

        eslesmeler = [
            eslesme.strip()
            for eslesme in re.findall(
                r"[-+]?\d[\d.,\s]*",
                metin,
            )
            if eslesme.strip()
        ]

        if len(eslesmeler) != 1:
            raise ExcelAktarimHatasi(
                f"Üretim adedi tek bir sayı içermeli: '{metin}'"
            )

        token = re.sub(r"\s+", "", eslesmeler[0])

        if token.startswith(("+", "-")):
            raise ExcelAktarimHatasi(
                f"Üretim adedi pozitif tam sayı olmalı: '{metin}'"
            )

        if token.isdigit():
            adet = int(token)

        elif re.fullmatch(
            r"\d{1,3}(?:[.,]\d{3})+",
            token,
        ):
            adet = int(re.sub(r"[.,]", "", token))

        else:
            raise ExcelAktarimHatasi(
                f"Üretim adedi tam sayı olmalı: '{metin}'"
            )

    if adet < 1:
        raise ExcelAktarimHatasi(
            f"Üretim adedi en az 1 olmalı: {deger!r}"
        )

    return adet


def _baslik_satiri_bul(ws):
    for satir_no, satir in enumerate(
        ws.iter_rows(min_row=1, max_row=10, values_only=True),
        start=1,
    ):
        temiz = [_sadelestir(hucre) for hucre in satir]
        if "TALEP NO" in temiz or "KART STOK NO" in temiz:
            return satir_no, temiz
    return None, None


def _hucre_al(satir, kolonlar, alan):
    index = kolonlar.get(alan)
    if index is None or index >= len(satir):
        return None
    deger = satir[index]
    return None if deger == "" else deger


def _tarih_coz_ve_dogrula(deger, alan_adi, excel_satir_no):
    sonuc = depo.tarih_coz(deger)
    if deger not in (None, "") and not sonuc:
        raise ExcelAktarimHatasi(
            f"Excel satır {excel_satir_no}: {alan_adi} okunamadı. Gelen değer: {deger!r}"
        )
    return sonuc


def _sira_coz(deger):
    if deger in (None, ""):
        return None, False
    try:
        return int(float(deger)), False
    except (TypeError, ValueError):
        return None, True


def _parti_anahtari_uret(taban_anahtar, plan_teslim, gerceklesen_teslim, plan_baslama, anahtar_gruplari):
    """Aynı Talep NO + Kart Stok No'ya sahip birden fazla satır (parti) olabilir

    — ör. aynı talebin bir kısmı bugün, kalanı farklı bir tarihte teslim
    edilecek ayrı bir parti olarak Excel'e ayrı satır halinde girilmiş olabilir.

    Kural:
    - İlk görülüşte taban anahtar (talep_no|stok_no) aynen kullanılır. Bu,
      tek-partili (bugüne kadarki tüm) kartlarla ve halihazırda diskte kayıtlı
      kartlarla eşleşmenin kesintiye uğramaması için ZORUNLUDUR.
    - Sonraki partiler için, mümkünse o partiyi asıl ayıran bilgiye göre —
      Planlanan Teslim T. > Gerçekleşen Teslim T. > Dizgi Başlama T. sırasıyla
      ilk doluya bakılır — okunabilir/kararlı bir ek üretilir
      (ör. "TLP-1|STK-1#2026-09-15"). Bu, Excel'de satırların yeri değişse
      (araya yeni bir parti satırı eklense) bile aynı partiyi sonraki
      import'ta yine aynı karta eşler; salt satır sırasına dayanan bir sayaç
      bu garantiyi veremezdi.
    - Ayırt edici tarih de yoksa, ya da iki parti aynı tarihi paylaşıyorsa,
      son çare olarak parti sırasına göre numaralanır (#2, #3, ...).

    Döner: (anahtar, coklu_parti_mi)
    """
    kullanilanlar = anahtar_gruplari.setdefault(taban_anahtar, set())

    if not kullanilanlar:
        anahtar = taban_anahtar
    else:
        ayrac = plan_teslim or gerceklesen_teslim or plan_baslama
        aday = f"{taban_anahtar}#{ayrac}" if ayrac else None
        if aday and aday not in kullanilanlar:
            anahtar = aday
        else:
            sira = len(kullanilanlar) + 1
            while f"{taban_anahtar}#{sira}" in kullanilanlar:
                sira += 1
            anahtar = f"{taban_anahtar}#{sira}"

    kullanilanlar.add(anahtar)
    return anahtar, len(kullanilanlar) > 1


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

def _sayfa_satirlarini_coz(ws, sayfa_adi, dizgi_tipi, anahtar_gruplari, parsed_liste):
    """Tek bir worksheet'i (MAKİNE ya da ELDE DİZGİ YENİ) parse eder.

    Sonuçları parsed_liste'ye ekler, anahtar_gruplari sözlüğünü günceller
    (iki sayfa arasında ve aynı sayfa içinde Talep NO + Kart Stok No aynı
    olan çoklu-parti satırları doğru şekilde ayırt etmek için bu sözlük
    çağıran tarafından paylaşılır — bkz. _parti_anahtari_uret) ve bu sayfada
    üretilen uyarı sayısını döner.
    """
    baslik_no, basliklar = _baslik_satiri_bul(ws)
    if not baslik_no:
        raise ExcelAktarimHatasi(
            f"'{sayfa_adi}' sayfasında başlık satırı bulunamadı. "
            "Dosyada 'Talep NO' veya 'Kart Stok No' sütunu olmalı."
        )

    kolonlar = {}
    for index, baslik in enumerate(basliklar):
        alan = BASLIK_ESLESME.get(baslik)
        if not alan:
            continue
        if alan in kolonlar:
            raise ExcelAktarimHatasi(
                f"'{sayfa_adi}' sayfası: '{baslik}' için birden fazla görünür Excel sütunu bulundu."
            )
        kolonlar[alan] = index

    for gerekli in ("talep_no", "stok_no", "adet_metin"):
        if gerekli not in kolonlar:
            ad = {
                "talep_no": "Talep NO",
                "stok_no": "Kart Stok No",
                "adet_metin": "Kart Üretim Adet",
            }[gerekli]
            raise ExcelAktarimHatasi(f"'{sayfa_adi}' sayfası: Zorunlu '{ad}' sütunu bulunamadı.")

    uyari_sayisi = 0

    for excel_satir_no, satir in enumerate(
        ws.iter_rows(min_row=baslik_no + 1, values_only=True),
        start=baslik_no + 1,
    ):
        if not any(hucre not in (None, "") for hucre in satir):
            continue

        talep_no = str(_hucre_al(satir, kolonlar, "talep_no") or "").strip()
        stok_no = str(_hucre_al(satir, kolonlar, "stok_no") or "").strip()
        if not talep_no or not stok_no:
            continue

        try:
            toplam_adet = adet_coz(_hucre_al(satir, kolonlar, "adet_metin"))
        except ExcelAktarimHatasi as exc:
            raise ExcelAktarimHatasi(f"'{sayfa_adi}' sayfası, satır {excel_satir_no}: {exc}") from exc

        plan_baslama_ham = _hucre_al(satir, kolonlar, "plan_baslama")
        plan_teslim_ham = _hucre_al(satir, kolonlar, "plan_teslim")
        gerceklesen_ham = _hucre_al(satir, kolonlar, "gerceklesen_teslim")

        plan_baslama = _tarih_coz_ve_dogrula(
            plan_baslama_ham, "Dizgi Başlama Tarihi", excel_satir_no
        )
        plan_teslim = _tarih_coz_ve_dogrula(
            plan_teslim_ham, "Planlanan Teslim Tarihi", excel_satir_no
        )
        gerceklesen = _tarih_coz_ve_dogrula(
            gerceklesen_ham, "Gerçekleşen Teslim Tarihi", excel_satir_no
        )

        if plan_baslama and plan_teslim and plan_baslama > plan_teslim:
            raise ExcelAktarimHatasi(
                f"'{sayfa_adi}' sayfası, satır {excel_satir_no}: Dizgi Başlama Tarihi ({plan_baslama}) "
                f"Planlanan Teslim Tarihinden ({plan_teslim}) sonra olamaz."
            )

        taban_anahtar = f"{talep_no}|{stok_no}"
        anahtar, coklu_parti = _parti_anahtari_uret(
            taban_anahtar, plan_teslim, gerceklesen, plan_baslama, anahtar_gruplari
        )
        if coklu_parti:
            uyari_sayisi += 1

        excel_durum_raw = _hucre_al(satir, kolonlar, "excel_durum")
        ilk_durum, durum_uyarisi = durum_coz(excel_durum_raw)
        if durum_uyarisi:
            uyari_sayisi += 1

        if (ilk_durum == depo.TESLIM_EDILDI and not gerceklesen):
            ilk_durum = None
            uyari_sayisi += 1

        elif (ilk_durum != depo.TESLIM_EDILDI and gerceklesen):
            uyari_sayisi += 1

        sira, sira_uyarisi = _sira_coz(_hucre_al(satir, kolonlar, "sira"))
        if sira_uyarisi:
            uyari_sayisi += 1

        plan = {
            "sira": sira,
            "talep_sahibi": str(_hucre_al(satir, kolonlar, "talep_sahibi") or "").strip(),
            "toplam_adet": toplam_adet,
            "adet_metin": str(_hucre_al(satir, kolonlar, "adet_metin") or "").strip(),
            "plan_hafta": str(_hucre_al(satir, kolonlar, "plan_hafta") or "").strip(),
            "plan_baslama": plan_baslama,
            "plan_teslim": plan_teslim,
            "excel_durum": str(excel_durum_raw or "").strip(),
            "pcb": str(_hucre_al(satir, kolonlar, "pcb") or "").strip(),
            "dizgi_tipi": dizgi_tipi,
            "dizgi_sorumlusu": str(_hucre_al(satir, kolonlar, "dizgi_sorumlusu") or "").strip(),
            "malzeme_bekliyor": 1 if malzeme_bekliyor_mu(excel_durum_raw) else 0,
        }

        parsed_liste.append(
            {
                "anahtar": anahtar,
                "talep_no": talep_no,
                "stok_no": stok_no,
                "plan": plan,
                "gerceklesen_teslim": gerceklesen,
                "ilk_durum": ilk_durum,
            }
        )

    return uyari_sayisi


def excelden_aktar(dosya_yolu, kullanici):
    with _import_kilidi:
        return _excelden_aktar(dosya_yolu, kullanici)


def _excelden_aktar(dosya_yolu, kullanici):
    snapshot_yolu = None
    wb = None

    try:
        snapshot_yolu = excel_deger_snapshot_olustur(dosya_yolu)
        wb = openpyxl.load_workbook(snapshot_yolu, data_only=True, read_only=True)
    except ExcelAktarimHatasi:
        raise
    except Exception as exc:
        raise ExcelAktarimHatasi(
            f"Dosya geçerli bir Excel çalışma kitabı değil: {exc}"
        ) from exc

    try:
        parsed = []
        anahtar_gruplari = {}
        uyari_sayisi = 0
        sayfa_bulundu = False

        for sayfa_adi, dizgi_tipi, zorunlu in KAYNAK_SAYFALARI:
            if sayfa_adi not in wb.sheetnames:
                if zorunlu:
                    raise ExcelAktarimHatasi(
                        f"'{sayfa_adi}' sayfası bulunamadı. "
                        f"Dosyadaki sayfalar: {', '.join(wb.sheetnames)}"
                    )
                continue  # opsiyonel sayfa (ör. ELDE DİZGİ YENİ) yoksa sessizce atla

            sayfa_bulundu = True
            ws = wb[sayfa_adi]
            uyari_sayisi += _sayfa_satirlarini_coz(
                ws, sayfa_adi, dizgi_tipi, anahtar_gruplari, parsed
            )

        if not sayfa_bulundu or not parsed:
            raise ExcelAktarimHatasi("Excel'de işlenecek kart satırı bulunamadı.")

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if snapshot_yolu:
            try:
                os.remove(snapshot_yolu)
            except OSError:
                pass

    return depo.excel_import_uygula(
        dosya_adi=os.path.basename(dosya_yolu),
        kullanici=kullanici,
        satirlar=parsed,
        uyari_sayisi=uyari_sayisi,
    )


# ---------------------------------------------------------------------------
# Rapor üretimi
# ---------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="0F2027")
BASLIK_YAZI = Font(name="Arial", bold=True, color="FFFFFF", size=11)
GOVDE_YAZI = Font(name="Arial", size=10)

def _excel_hucre_yaz(hucre, deger):
    hucre.value = deger

    if isinstance(deger, str) and deger.startswith("="):
        hucre.data_type = "s"

def _sayfa_yaz(ws, basliklar, satirlar):
    ws.append(basliklar)
    for hucre in ws[1]:
        hucre.fill = BASLIK_DOLGU
        hucre.font = BASLIK_YAZI
        hucre.alignment = Alignment(horizontal="center", vertical="center")

    for satir in satirlar:
        satir_no = ws.max_row + 1

        for sutun_no, deger in enumerate(satir,start=1):
            _excel_hucre_yaz(
                ws.cell(row=satir_no,column=sutun_no,),deger)


    for sutun in range(1, len(basliklar) + 1):
        en = max(
            [len(str(basliklar[sutun - 1])), 10]
            + [len(str(satir[sutun - 1])) for satir in satirlar[:400]]
        )
        ws.column_dimensions[get_column_letter(sutun)].width = min(40, en + 3)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for satir in ws.iter_rows(min_row=2):
        for hucre in satir:
            hucre.font = GOVDE_YAZI


def calisma_kitabi_uret(kartlar, loglar, ozet):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Kart Durumları"
    _sayfa_yaz(
        ws,
        [
            "ID",
            "Sıra",
            "Talep NO",
            "Talep Sahibi",
            "Kart Stok No",
            "Toplam Adet",
            "Tamamlanan",
            "Kalan",
            "Durum",
            "Kaynak Durumu",
            "Değerlendirme",
            "Sapma (gün)",
            "Plan Başlangıç",
            "Plan Teslim",
            "Üretim Başlangıç",
            "Üretim Bitiş",
            "Teslim Zamanı",
            "Gerçekleşen Teslim",
            "Operatör",
            "Not",
            "PCB",
            "Dizgi Tipi",
            "Dizgi Sorumlusu",
            "Malzeme Bekliyor",
            "Kaynakta Aktif",
            "Admin Gizli",
        ],
        [
            [
                k.get("id"),
                k.get("sira"),
                k.get("talep_no"),
                k.get("talep_sahibi"),
                k.get("stok_no"),
                k.get("toplam_adet"),
                k.get("tamamlanan_adet"),
                k.get("kalan_adet"),
                k.get("durum") or "DURUMU EKSİK",
                k.get("excel_durum") or "",
                k.get("rozet"),
                k.get("sapma") if k.get("sapma") is not None else "",
                k.get("plan_baslama") or "",
                k.get("plan_teslim") or "",
                k.get("baslama_zamani") or "",
                k.get("bitis_zamani") or "",
                k.get("teslim_zamani") or "",
                k.get("gerceklesen_teslim") or "",
                k.get("operator") or "",
                k.get("aciklama") or "",
                k.get("pcb") or "",
                k.get("dizgi_tipi") or depo.DIZGI_TIPI_MAKINE,
                k.get("dizgi_sorumlusu") or "",
                "EVET" if k.get("malzeme_bekliyor") else "",
                k.get("source_active", 1),
                k.get("admin_gizli", 0),
            ]
            for k in kartlar
        ],
    )

    ws2 = wb.create_sheet("İşlem Logu")
    _sayfa_yaz(
        ws2,
        ["Zaman", "Kullanıcı", "Rol", "İşlem", "Talep NO", "Kart Stok No", "Adet", "Detay"],
        [
            [
                l.get("zaman"),
                l.get("kullanici"),
                l.get("rol"),
                l.get("islem"),
                l.get("talep_no") or "",
                l.get("stok_no") or "",
                l.get("adet") if l.get("adet") is not None else "",
                l.get("detay") or "",
            ]
            for l in loglar
        ],
    )

    ws3 = wb.create_sheet("Özet")
    _sayfa_yaz(ws3, ["Başlık", "Değer"], ozet)
    return wb


def kitap_baytlari(wb):
    tampon = io.BytesIO()
    wb.save(tampon)
    tampon.seek(0)
    return tampon


def dosya_adi(on_ek):
    return f"{on_ek}_{datetime.now():%Y%m%d_%H%M}.xlsx"
```


## `kullanici_yonet.py`


```python
"""PDGM kullanıcı yönetimi için küçük CLI aracı.

Parolalar plaintext saklanmaz; yalnız Werkzeug hash'i kullanicilar.json'a yazılır.
"""

from __future__ import annotations

import getpass
import json
import os
import sys

from werkzeug.security import generate_password_hash

KOK = os.path.dirname(os.path.abspath(__file__))
DOSYA = os.path.join(KOK, "data", "kullanicilar.json")
ROLLER = {"admin", "operator", "gozlemci"}


def oku():
    if not os.path.exists(DOSYA):
        raise SystemExit("Önce uygulamayı bir kez çalıştırın; data/kullanicilar.json oluşturulsun.")
    with open(DOSYA, encoding="utf-8") as f:
        veri = json.load(f)
    if not isinstance(veri, dict):
        raise SystemExit("kullanicilar.json geçersiz.")
    return veri


def yaz(veri):
    gecici = DOSYA + ".yeni"
    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(gecici, DOSYA)


def _parola_sor():
    parola = getpass.getpass("Yeni parola: ")
    parola_tekrar = getpass.getpass("Yeni parola tekrar: ")
    if parola != parola_tekrar:
        raise SystemExit("Parolalar eşleşmiyor.")
    if len(parola) < 8:
        raise SystemExit("Parola en az 8 karakter olmalı.")
    return parola


def _aktif_admin_sayisi(veri, haric=None):
    return sum(
        1
        for k, b in veri.items()
        if k != haric and b.get("rol") == "admin" and b.get("aktif", True)
    )


def ekle(kullanici, rol, gorunen_ad):
    kullanici = kullanici.strip()
    rol = rol.strip().lower()
    gorunen_ad = gorunen_ad.strip()

    if not kullanici:
        raise SystemExit("Kullanıcı adı boş olamaz.")
    if rol not in ROLLER:
        raise SystemExit(f"Rol şu değerlerden biri olmalı: {', '.join(sorted(ROLLER))}")

    veri = oku()
    if kullanici in veri:
        raise SystemExit("Bu kullanıcı adı zaten var.")

    parola = getpass.getpass("Parola: ")
    parola_tekrar = getpass.getpass("Parola tekrar: ")
    if parola != parola_tekrar:
        raise SystemExit("Parolalar eşleşmiyor.")
    if len(parola) < 8:
        raise SystemExit("Parola en az 8 karakter olmalı.")

    veri[kullanici] = {
        "sifre_hash": generate_password_hash(parola),
        "rol": rol,
        "ad": gorunen_ad or kullanici,
        "aktif": True,
    }
    yaz(veri)
    print(f"Kullanıcı eklendi: {kullanici} ({rol})")


def parola_degistir(kullanici):
    kullanici = kullanici.strip()
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    veri[kullanici]["sifre_hash"] = generate_password_hash(_parola_sor())
    yaz(veri)
    print(f"{kullanici}: parola güncellendi.")
    print("Sunucuyu yeniden başlatmanıza gerek yok.")


def rol_degistir(kullanici, yeni_rol):
    kullanici = kullanici.strip()
    yeni_rol = yeni_rol.strip().lower()

    if yeni_rol not in ROLLER:
        raise SystemExit(f"Rol şu değerlerden biri olmalı: {', '.join(sorted(ROLLER))}")

    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    eski_rol = veri[kullanici].get("rol", "-")

    if eski_rol == "admin" and yeni_rol != "admin":
        if _aktif_admin_sayisi(veri, haric=kullanici) == 0:
            raise SystemExit(
                "Bu son aktif admin. Rolü düşürürseniz sistemi yönetemezsiniz. "
                "Önce başka bir admin oluşturun."
            )

    veri[kullanici]["rol"] = yeni_rol
    yaz(veri)
    print(f"{kullanici}: {eski_rol} -> {yeni_rol}")
    print("Değişiklik ilk request'te etkili olur; yeniden başlatma gerekmez.")


def aktiflik(kullanici, aktif):
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")

    if not aktif and veri[kullanici].get("rol") == "admin":
        if _aktif_admin_sayisi(veri, haric=kullanici) == 0:
            raise SystemExit(
                "Bu son aktif admin. Pasife alırsanız sisteme admin olarak "
                "giremezsiniz. Önce başka bir admin oluşturun."
            )

    veri[kullanici]["aktif"] = aktif
    yaz(veri)
    print(f"{kullanici}: {'aktif' if aktif else 'pasif'}")


def listele():
    for kullanici, bilgi in oku().items():
        print(
            f"{kullanici:20} "
            f"{bilgi.get('rol', '-'):10} "
            f"{bilgi.get('ad', '-'):30} "
            f"aktif={bilgi.get('aktif', True)}"
        )


def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Kullanım:\n"
            "  python kullanici_yonet.py listele\n"
            "  python kullanici_yonet.py ekle <kullanici> <admin|operator|gozlemci> <Görünen Ad>\n"
            "  python kullanici_yonet.py parola <kullanici>\n"
            "  python kullanici_yonet.py rol <kullanici> <admin|operator|gozlemci>\n"
            "  python kullanici_yonet.py pasif <kullanici>\n"
            "  python kullanici_yonet.py aktif <kullanici>"
        )

    komut = sys.argv[1].lower()
    if komut == "listele":
        listele()
    elif komut == "ekle" and len(sys.argv) >= 5:
        ekle(sys.argv[2], sys.argv[3], " ".join(sys.argv[4:]))
    elif komut == "parola" and len(sys.argv) == 3:
        parola_degistir(sys.argv[2])
    elif komut == "rol" and len(sys.argv) == 4:
        rol_degistir(sys.argv[2], sys.argv[3])
    elif komut == "pasif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], False)
    elif komut == "aktif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], True)
    else:
        raise SystemExit("Geçersiz komut veya eksik argüman.")


if __name__ == "__main__":
    main()
```


# 2. TEMPLATES


## `templates/base.html`


```html
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="csrf-token" content="{{ csrf_token }}">
    <title>{% block title %}PDGM İş Takip{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='stil.css') }}">
    {% block head %}{% endblock %}
</head>
<body class="{% block body_class %}{% endblock %}">
<header class="ust-cubuk">
    <a href="{{ url_for('ana') }}" class="marka-link">
        <span class="marka-isaret">
            <img src="{{ url_for('static', filename='pdgm_logo.png') }}" alt="PDGM" class="marka-logo">
        </span>
        <span class="marka-metin">
            <strong>İş Takip Sistemi</strong>
            <small>Prototip Kart Dizgi Atölyesi</small>
        </span>
    </a>

    <nav class="ana-nav" aria-label="Ana menü">
        {% if oturum_rol in ["admin", "operator", "gozlemci"] %}
            <a class="{% if request.endpoint == 'panel' %}aktif{% endif %}" href="{{ url_for('panel') }}">Pano</a>
        {% endif %}
    
        {% if oturum_rol in ["admin", "operator"] %}
            <a class="{% if request.endpoint == 'operator' %}aktif{% endif %}" href="{{ url_for('operator') }}">Operatör</a>
        {% endif %}
    
        {% if oturum_rol in ["admin", "operator", "gozlemci"] %}
            <a class="{% if request.endpoint == 'monitor' %}aktif{% endif %}" href="{{ url_for('monitor') }}">Monitör</a>
        {% endif %}
    
        {% if oturum_rol == "admin" %}
            <a class="{% if request.endpoint == 'yonetim' %}aktif{% endif %}" href="{{ url_for('yonetim') }}">Yönetim</a>
        {% endif %}
    </nav>

    <div class="oturum">
        <div class="oturum-metin">
            <strong>{{ oturum_ad or oturum_kullanici }}</strong>
            <small>{{ oturum_rol|upper }}</small>
        </div>
        {% if oturum_kullanici %}
        <form method="post" action="{{ url_for('cikis') }}">
            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
            <button class="buton buton-hayalet buton-kucuk" type="submit">Çıkış</button>
        </form>
        {% endif %}
    </div>
</header>

<main class="sayfa">
    {% with mesajlar = get_flashed_messages(with_categories=true) %}
        {% if mesajlar %}
        <section class="bildirimler" aria-live="polite">
            {% for kategori, mesaj in mesajlar %}
                <div class="bildirim {{ kategori }}">{{ mesaj }}</div>
            {% endfor %}
        </section>
        {% endif %}
    {% endwith %}

    {% block content %}{% endblock %}
</main>

<div id="toast-alani" class="toast-alani" aria-live="polite"></div>

<script>
function csrfToken() {
    return document.querySelector('meta[name="csrf-token"]')?.content || "";
}

async function pdgmFetch(url, options = {}) {
    const headers = new Headers(options.headers || {});
    headers.set("X-CSRF-Token", csrfToken());

    if (options.body && !(options.body instanceof FormData) && !headers.has("Content-Type")) {
        headers.set("Content-Type", "application/json");
    }

    const response = await fetch(url, {credentials: "same-origin", ...options, headers});
    const contentType = response.headers.get("content-type") || "";
    const data = contentType.includes("application/json")
        ? await response.json()
        : {hata: await response.text()};

    if (!response.ok) {
        throw new Error(data.hata || data.detail || "İşlem tamamlanamadı.");
    }
    return data;
}

function toast(mesaj, tip = "basari") {
    const alan = document.getElementById("toast-alani");
    if (!alan) return;

    const kutu = document.createElement("div");
    kutu.className = `toast ${tip}`;
    kutu.textContent = mesaj;
    alan.appendChild(kutu);

    requestAnimationFrame(() => kutu.classList.add("goster"));
    window.setTimeout(() => {
        kutu.classList.remove("goster");
        window.setTimeout(() => kutu.remove(), 200);
    }, 3200);
}

function hataMesaji(hata) {
    toast(hata?.message || "Beklenmeyen bir hata oluştu.", "hata");
}

function dialogKapat(dialog) {
    if (dialog?.open) dialog.close();
}

document.addEventListener("click", (event) => {
    const kapat = event.target.closest("[data-dialog-kapat]");
    if (kapat) dialogKapat(kapat.closest("dialog"));
});
</script>
{% block scripts %}{% endblock %}
</body>
</html>
```


## `templates/giris.html`


```html
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Giriş · PDGM İş Takip</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='stil.css') }}">
</head>
<body class="giris-sayfa">
<main class="giris-kutu">
    <div class="giris-logo">
        <img src="{{ url_for('static', filename='pdgm_logo.png') }}" alt="PDGM Logo">
    </div>

    <h1>PDGM İŞ TAKİP</h1>
    <p class="alt">Baskı Dizgi Atölyesi</p>

    {% if hata %}<div class="bildirim hata">{{ hata }}</div>{% endif %}

    <form method="post" autocomplete="off" class="giris-form">
        <div class="alan">
            <label for="kullanici">Kullanıcı adı</label>
            <input type="text" id="kullanici" name="kullanici" autocomplete="username" autofocus required>
        </div>
        <div class="alan">
            <label for="sifre">Şifre</label>
            <input type="password" id="sifre" name="sifre" autocomplete="current-password" required>
        </div>
        <button class="buton buton-ana tam-genislik" type="submit">Giriş Yap</button>
    </form>

    <p class="giris-dipnot">Bu sistem yalnızca yetkili PDGM personeli içindir.</p>
</main>
</body>
</html>
```


## `templates/monitor.html`


```html
{% extends "base.html" %}

{% block title %}Monitör · PDGM İş Takip{% endblock %}
{% block body_class %}monitor-body{% endblock %}

{% block content %}
<div class="monitor-sayfa">

    <section class="monitor-ust">
        <div>
            <p class="ust-etiket">ATÖLYE CANLI GÖRÜNÜMÜ</p>
            <h1>Üretim Monitörü</h1>
            <p class="monitor-aciklama">
                Aktif üretim ve üretim sırası.
            </p>
        </div>

        <div class="monitor-zaman">
            <span id="monitor-saat">{{ guncelleme }}</span>
            <small>
                {{ bugun }} · Son veri yenileme {{ guncelleme }}
            </small>
        </div>
    </section>

    <section class="monitor-grid">

        <!-- ============================================================
             DİZGİDE
        ============================================================ -->

        <article class="panel-kutu monitor-bolum monitor-dizgide">

            <div class="panel-baslik monitor-bolum-baslik">
                <div>
                    <p class="ust-etiket">AKTİF ÜRETİM</p>
                    <h2>DİZGİDE</h2>
                </div>

                <span class="sayi-rozet">
                    {{ dizgide|length }}
                </span>
            </div>

            {% if dizgide %}

            <div class="monitor-liste monitor-dizgide-liste {% if dizgide|length >= 3 %}uc-dizgi{% endif %}"
                data-monitor-grup="dizgide"
                data-sayfa-boyutu="3">

                {% for k in dizgide %}

                <article class="monitor-kart {{ k.renk }}" data-monitor-kart>

                    <div class="monitor-kart-ust">

                        <div class="monitor-talep">
                            <span>TALEP</span>
                            <strong>
                                {{ k.talep_no or "—" }}
                            </strong>
                        </div>

                        <span class="durum-rozet {{ k.renk }}">
                            {{ k.rozet }}
                        </span>

                    </div>

                    {% if k.elle_dizgi_mi or k.malzeme_bekliyor %}
                    <div class="monitor-ek-rozetler">
                        {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                        {% if k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif %}
                    </div>
                    {% endif %}

                    <h3>
                        {{ k.stok_no or "Stok no yok" }}
                    </h3>

                    <p class="monitor-sahip">
                        {{ k.talep_sahibi or "Talep sahibi belirtilmemiş" }}
                    </p>

                    <div class="monitor-adet">
                        <strong>
                            {{ k.tamamlanan_adet }}
                            /
                            {{ k.toplam_adet }}
                        </strong>

                        <span>
                            adet tamamlandı
                        </span>
                    </div>

                    <div class="ilerleme monitor-ilerleme">

                        <div class="ilerleme-ust">
                            <span>
                                Üretim ilerlemesi
                            </span>

                            <strong>
                                %{{ k.adet_yuzde }}
                            </strong>
                        </div>

                        <div class="ilerleme-ray">
                            <div
                                class="ilerleme-dolgu"
                                style="width: {{ [k.adet_yuzde, 100]|min }}%">
                            </div>
                        </div>

                    </div>

                    <div class="monitor-meta">

                        <span>
                            <small>Plan teslim</small>
                            <strong>
                                {{ k.plan_teslim|gun }}
                            </strong>
                        </span>

                    </div>

                </article>

                {% endfor %}

            </div>

            {% else %}

            <div class="bos-durum monitor-bos">
                Şu anda dizgide kart bulunmuyor.
            </div>

            {% endif %}

        </article>


        <!-- ============================================================
             PLANA ALINDI
        ============================================================ -->

        <article class="panel-kutu monitor-bolum monitor-plana">

            <div class="panel-baslik monitor-bolum-baslik">

                <div>
                    <p class="ust-etiket">ÜRETİM KUYRUĞU</p>
                    <h2>PLANA ALINDI</h2>
                </div>

                <span class="sayi-rozet">
                    {{ plana_alindi|length }}
                </span>

            </div>

            {% if plana_alindi %}

            <div class="monitor-liste monitor-plan-liste"
                data-monitor-grup="plan"
                data-sayfa-boyutu="6">

                {% for k in plana_alindi %}

                <article class="monitor-plan-kart {{ k.renk }}" data-monitor-kart>

                    <div class="monitor-plan-sira">
                        {{ k.sira or loop.index }}
                    </div>

                    <div class="monitor-plan-icerik">

                        <div class="monitor-plan-ust">

                            <div>
                                <span>TALEP</span>
                                <strong>
                                    {{ k.talep_no or "—" }}
                                </strong>
                            </div>

                            <span class="durum-rozet {{ k.renk }}">
                                {{ k.rozet }}
                            </span>

                        </div>

                        {% if k.elle_dizgi_mi or k.malzeme_bekliyor %}
                        <div class="monitor-ek-rozetler">
                            {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                            {% if k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif %}
                        </div>
                        {% endif %}

                        <h3>
                            {{ k.stok_no or "Stok no yok" }}
                        </h3>

                        <p>
                            {{ k.talep_sahibi or "Talep sahibi belirtilmemiş" }}
                        </p>

                        <div class="monitor-plan-meta">

                            <span>
                                <small>Adet</small>
                                <strong>
                                    {{ k.toplam_adet }}
                                </strong>
                            </span>

                            <span>
                                <small>Başlangıç</small>
                                <strong>
                                    {{ k.plan_baslama|gun }}
                                </strong>
                            </span>

                            <span>
                                <small>Teslim</small>
                                <strong>
                                    {{ k.plan_teslim|gun }}
                                </strong>
                            </span>

                        </div>

                    </div>

                </article>

                {% endfor %}

            </div>

            {% else %}

            <div class="bos-durum monitor-bos">
                Plana alınmış bekleyen kart yok.
            </div>

            {% endif %}

        </article>

    </section>

</div>
{% endblock %}


{% block scripts %}
<script>
(() => {
    const saat = document.getElementById("monitor-saat");

    function saatiGuncelle() {
        if (!saat) return;

        saat.textContent = new Intl.DateTimeFormat("tr-TR", {
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit"
        }).format(new Date());
    }

    function rotasyonBaslat(grup) {
        const kartlar = [...grup.querySelectorAll("[data-monitor-kart]")];
        const sayfaBoyutu = Number(grup.dataset.sayfaBoyutu || 0);

        if (!sayfaBoyutu || kartlar.length <= sayfaBoyutu) {
            return;
        }

        const sayfaSayisi = Math.ceil(
            kartlar.length / sayfaBoyutu
        );

        let aktifSayfa = 0;

        function sayfayiGoster() {
            const baslangic = aktifSayfa * sayfaBoyutu;
            const bitis = baslangic + sayfaBoyutu;

            kartlar.forEach((kart, index) => {
                kart.hidden = !(
                    index >= baslangic &&
                    index < bitis
                );
            });
        }

        sayfayiGoster();

        window.setInterval(() => {
            aktifSayfa = (aktifSayfa + 1) % sayfaSayisi;
            sayfayiGoster();
        }, 5000);
    }

    saatiGuncelle();

    window.setInterval(
        saatiGuncelle,
        1000
    );

    document
        .querySelectorAll("[data-monitor-grup]")
        .forEach(rotasyonBaslat);

    window.setTimeout(
        () => window.location.reload(),
        30000
    );
})();
</script>
{% endblock %}
```


## `templates/operator.html`


```html
{% extends "base.html" %}
{% block title %}Operatör · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell operator-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">OPERATÖR EKRANI</p>
            <h1>Kart İşlemleri</h1>
            <p class="soluk">Üretimi başlatın, tamamlanan adedi kaydedin, üretim bittiğinde TESLİM EDİLDİ olarak işaretleyin.</p>
        </div>
    </section>

    <section class="istatistik-grid istatistik-grid-3" aria-label="Durum özeti">
        <article class="istatistik kart-dizgide"><span>Dizgide</span><strong>{{ sayac.dizgide }}</strong></article>
        <article class="istatistik kart-plana"><span>Plana Alındı</span><strong>{{ sayac.plana_alindi }}</strong></article>
        <article class="istatistik kart-teslim"><span>Teslim Edildi</span><strong>{{ sayac.teslim }}</strong></article>
    </section>

    <section class="arac-cubugu panel-kutu operator-arac-cubugu">
        <div class="arama">
            <label for="kart-ara">Kart ara</label>
            <input id="kart-ara" type="search" placeholder="Talep no, stok no, talep sahibi, PCB, dizgi sorumlusu..." autocomplete="off">
        </div>

        <div class="operator-filtre-alani">
            <div class="filtreler" role="group" aria-label="Durum filtresi">
                <button class="filtre aktif" data-filtre="AKTIF" type="button">Aktif</button>
                <button class="filtre" data-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                <button class="filtre" data-filtre="DİZGİDE" type="button">Dizgide</button>
                <button class="filtre" data-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
                <button class="filtre" data-filtre="HEPSI" type="button">Hepsi</button>
            </div>
            <div class="operator-filtre-meta">
                <span id="operator-sonuc"></span>
                <button id="operator-temizle" class="operator-temizle" type="button" hidden>Filtreleri temizle</button>
            </div>
        </div>
    </section>

    <section id="operator-kartlari" class="operator-grid">
    {% for k in kartlar %}
        <article class="operator-kart {{ k.renk }}"
                 data-kart
                 data-durum="{{ k.durum }}"
                 data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.pcb or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or '') ~ ' ' ~ (k.dizgi_sorumlusu or ''))|lower }}">

            <div class="operator-kart-ust">
                <div>
                    <span class="ust-etiket">{{ k.talep_no or "TALEP YOK" }}</span>
                    <h2>{{ k.stok_no or "Stok no yok" }}</h2>
                    <p>{{ k.talep_sahibi or "Talep sahibi yok" }}</p>
                </div>
                <div class="operator-rozetler">
                    <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                    {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                    {% if k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif %}
                    {% if k.kaynakta_yok %}<span class="durum-rozet uyari">Kaynak Excel'de yok</span>{% endif %}
                </div>
            </div>

            <div class="durum-satiri">
                <span><b>Durum:</b> {{ k.durum }}</span>
                <span><b>Kaynak:</b> {{ k.kaynak_durumu or "—" }}</span>
            </div>

            <div class="bilgi-grid">
                <div><span>Toplam</span><strong>{{ k.toplam_adet }}</strong></div>
                <div><span>Tamamlanan</span><strong>{{ k.tamamlanan_adet }}</strong></div>
                <div><span>Kalan</span><strong>{{ k.kalan_adet }}</strong></div>
                <div><span>Plan Teslim</span><strong>{{ k.plan_teslim|gun }}</strong></div>
            </div>

            <div class="ilerleme">
                <div class="ilerleme-ust">
                    <span>Üretim ilerlemesi</span>
                    <strong>%{{ k.adet_yuzde }}</strong>
                </div>
                <div class="ilerleme-ray">
                    <div class="ilerleme-dolgu" style="width: {{ [k.adet_yuzde, 100]|min }}%"></div>
                </div>
            </div>

            {% if k.aciklama %}<p class="kart-not"><b>Not:</b> {{ k.aciklama }}</p>{% endif %}

            <div class="kart-ek-bilgi">
                {% if k.pcb %}<span>{{ "Malzeme/PCB" if k.elle_dizgi_mi else "PCB" }}: {{ k.pcb }}</span>{% endif %}
                {% if k.dizgi_sorumlusu %}<span>Dizgi Sorumlusu: {{ k.dizgi_sorumlusu }}</span>{% endif %}
                {% if k.operator %}<span>Son işlem: {{ k.operator }}</span>{% endif %}
                {% if k.bitis_zamani %}<span>Üretim bitiş: {{ k.bitis_zamani|gun }}</span>{% endif %}
            </div>

            <div class="kart-aksiyonlar">
                {% if k.durum == "PLANA ALINDI" and k.malzeme_bekliyor %}
                    <button class="buton buton-hayalet" type="button" disabled
                            title="Kaynak Excel'de malzeme bekliyor olarak işaretli; dizgiye alınamaz.">
                        Malzeme Bekleniyor
                    </button>
                {% elif k.durum == "PLANA ALINDI" %}
                    <button class="buton buton-ana" type="button" data-baslat
                            data-id="{{ k.id }}" data-kalan="{{ k.kalan_adet }}"
                            data-talep="{{ k.talep_no or '' }}" data-stok="{{ k.stok_no or '' }}">
                        Dizgiye Al
                    </button>
                {% elif k.durum == "DİZGİDE" and k.kalan_adet > 0 %}
                    <button class="buton buton-ana" type="button" data-bitir
                            data-id="{{ k.id }}" data-kalan="{{ k.kalan_adet }}"
                            data-talep="{{ k.talep_no or '' }}" data-stok="{{ k.stok_no or '' }}">
                        Adet Bitir
                    </button>
                {% elif k.durum == "DİZGİDE" and k.kalan_adet == 0 %}
                    <button class="buton buton-basari" type="button" data-teslim
                            data-id="{{ k.id }}" data-talep="{{ k.talep_no or '' }}">
                        Teslim Edildi
                    </button>
                {% endif %}

                <button class="buton buton-hayalet" type="button" data-not
                        data-id="{{ k.id }}" data-not-mevcut="{{ k.aciklama or '' }}"
                        data-talep="{{ k.talep_no or '' }}">Not</button>
            </div>
        </article>
    {% else %}
        <div class="bos-durum tam-satir">Görüntülenecek kart yok.</div>
    {% endfor %}
    </section>

    <div id="operator-bos" class="bos-durum" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
</div>

<dialog id="baslat-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="baslat-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">DİZGİYE AL</p><h2 id="baslat-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="baslat-id">
        <div class="alan">
            <label for="baslat-adet">Dizgiye alınacak adet</label>
            <input type="number" id="baslat-adet" min="1" required>
            <small id="baslat-kalan"></small>
        </div>
        <div class="alan">
            <label for="baslat-not">Not <span class="soluk">(opsiyonel)</span></label>
            <textarea id="baslat-not" rows="3"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Dizgiye Al</button>
        </div>
    </form>
</dialog>

<dialog id="bitir-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="bitir-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">ÜRETİM ADEDİ</p><h2 id="bitir-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="bitir-id">
        <div class="alan">
            <label for="bitir-adet">Tamamlanan adet</label>
            <input type="number" id="bitir-adet" min="1" required>
            <small id="bitir-kalan"></small>
        </div>
        <div class="alan">
            <label for="bitir-not">Not <span class="soluk">(opsiyonel)</span></label>
            <textarea id="bitir-not" rows="3"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>

<dialog id="not-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="not-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">KART NOTU</p><h2 id="not-baslik">Not Düzenle</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="not-id">
        <div class="alan">
            <label for="not-metin">Not</label>
            <textarea id="not-metin" rows="5"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>
{% endblock %}

{% block scripts %}
<script>
(() => {
    const kartAra = document.getElementById("kart-ara");
    const filtreButonlari = [...document.querySelectorAll("[data-filtre]")];
    const kartlar = [...document.querySelectorAll("[data-kart]")];
    const sonuc = document.getElementById("operator-sonuc");
    const temizle = document.getElementById("operator-temizle");
    const bos = document.getElementById("operator-bos");

    const FILTRE_KEY = "pdgm-op-filtre";
    const ARAMA_KEY = "pdgm-op-arama";
    const SCROLL_KEY = "pdgm-op-scroll";
    let aktifFiltre = sessionStorage.getItem(FILTRE_KEY) || "AKTIF";

    function aktifDurum(durum) {
        return ["PLANA ALINDI", "DİZGİDE"].includes(durum);
    }

    function durumUygunMu(durum) {
        if (aktifFiltre === "HEPSI") return true;
        if (aktifFiltre === "AKTIF") return aktifDurum(durum);
        return durum === aktifFiltre;
    }

    function filtrele() {
        const arama = (kartAra.value || "").trim().toLocaleLowerCase("tr-TR");
        let gorunen = 0;

        for (const kart of kartlar) {
            const uygun = durumUygunMu(kart.dataset.durum) && (!arama || kart.dataset.arama.includes(arama));
            kart.hidden = !uygun;
            if (uygun) gorunen += 1;
        }

        sonuc.textContent = `${gorunen} kart gösteriliyor`;
        bos.hidden = gorunen !== 0;
        temizle.hidden = !arama && aktifFiltre === "AKTIF";
    }

    function filtreSec(deger) {
        aktifFiltre = deger;
        sessionStorage.setItem(FILTRE_KEY, deger);
        filtreButonlari.forEach((buton) => {
            const secili = buton.dataset.filtre === deger;
            buton.classList.toggle("aktif", secili);
            buton.setAttribute("aria-pressed", secili ? "true" : "false");
        });
        filtrele();
    }

    function durumuSaklaVeYenile() {
        sessionStorage.setItem(ARAMA_KEY, kartAra.value || "");
        sessionStorage.setItem(SCROLL_KEY, String(window.scrollY));
        window.setTimeout(() => window.location.reload(), 350);
    }

    async function kartAksiyonu(url, kartId, onayMesaji, basariMesaji) {
        if (onayMesaji && !window.confirm(onayMesaji)) return false;
        const data = await pdgmFetch(url, {
            method: "POST",
            body: JSON.stringify({kart_id: Number(kartId)})
        });
        toast(data.mesaj || basariMesaji);
        return true;
    }

    kartAra.value = sessionStorage.getItem(ARAMA_KEY) || "";
    kartAra.addEventListener("input", () => {
        sessionStorage.setItem(ARAMA_KEY, kartAra.value || "");
        filtrele();
    });

    filtreButonlari.forEach((buton) => buton.addEventListener("click", () => filtreSec(buton.dataset.filtre)));
    temizle.addEventListener("click", () => {
        kartAra.value = "";
        sessionStorage.removeItem(ARAMA_KEY);
        filtreSec("AKTIF");
        kartAra.focus();
    });

    const gecerliFiltre = filtreButonlari.some((buton) => buton.dataset.filtre === aktifFiltre);
    filtreSec(gecerliFiltre ? aktifFiltre : "AKTIF");

    const kayitliScroll = sessionStorage.getItem(SCROLL_KEY);
    if (kayitliScroll !== null) {
        sessionStorage.removeItem(SCROLL_KEY);
        window.scrollTo(0, Number(kayitliScroll) || 0);
    }

    document.querySelectorAll("[data-baslat]").forEach((buton) => {
        buton.addEventListener("click", () => {
            const kalan = Number(buton.dataset.kalan || 1);
            document.getElementById("baslat-id").value = buton.dataset.id;
            document.getElementById("baslat-adet").value = kalan;
            document.getElementById("baslat-adet").max = kalan;
            document.getElementById("baslat-not").value = "";
            document.getElementById("baslat-kalan").textContent = `En fazla ${kalan} adet`;
            document.getElementById("baslat-baslik").textContent = `${buton.dataset.talep || "Kart"} · ${buton.dataset.stok || ""}`;
            document.getElementById("baslat-dialog").showModal();
        });
    });

    document.querySelectorAll("[data-bitir]").forEach((buton) => {
        buton.addEventListener("click", () => {
            const kalan = Number(buton.dataset.kalan || 1);
            document.getElementById("bitir-id").value = buton.dataset.id;
            document.getElementById("bitir-adet").value = kalan;
            document.getElementById("bitir-adet").max = kalan;
            document.getElementById("bitir-not").value = "";
            document.getElementById("bitir-kalan").textContent = `Kalan ${kalan} adet`;
            document.getElementById("bitir-baslik").textContent = `${buton.dataset.talep || "Kart"} · ${buton.dataset.stok || ""}`;
            document.getElementById("bitir-dialog").showModal();
        });
    });

    document.querySelectorAll("[data-not]").forEach((buton) => {
        buton.addEventListener("click", () => {
            document.getElementById("not-id").value = buton.dataset.id;
            document.getElementById("not-metin").value = buton.dataset.notMevcut || "";
            document.getElementById("not-baslik").textContent = `${buton.dataset.talep || "Kart"} · Not`;
            document.getElementById("not-dialog").showModal();
        });
    });

    document.querySelectorAll("[data-teslim]").forEach((buton) => {
        buton.addEventListener("click", async () => {
            try {
                const tamam = await kartAksiyonu(
                    "/api/teslim-et",
                    buton.dataset.id,
                    `${buton.dataset.talep || "Bu kart"} fiziksel olarak teslim edildi mi?`,
                    "Teslim kaydedildi."
                );
                if (tamam) durumuSaklaVeYenile();
            } catch (hata) {
                hataMesaji(hata);
            }
        });
    });

    document.getElementById("baslat-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;
        try {
            const data = await pdgmFetch("/api/basla", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: Number(document.getElementById("baslat-id").value),
                    adet: Number(document.getElementById("baslat-adet").value),
                    not: document.getElementById("baslat-not").value.trim()
                })
            });
            toast(data.mesaj || "Kart DİZGİDE durumuna alındı.");
            dialogKapat(document.getElementById("baslat-dialog"));
            durumuSaklaVeYenile();
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });

    document.getElementById("bitir-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;
        try {
            const kartId = Number(document.getElementById("bitir-id").value);
            const data = await pdgmFetch("/api/bitir", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: kartId,
                    adet: Number(document.getElementById("bitir-adet").value),
                    not: document.getElementById("bitir-not").value.trim()
                })
            });
            dialogKapat(document.getElementById("bitir-dialog"));

            if (data.uretim_bitti) {
                toast(data.mesaj, "basari");
            } else {
                toast(data.mesaj);
            }
            durumuSaklaVeYenile();
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });

    document.getElementById("not-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;
        try {
            await pdgmFetch("/api/not", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: Number(document.getElementById("not-id").value),
                    not: document.getElementById("not-metin").value.trim()
                })
            });
            toast("Not kaydedildi.");
            dialogKapat(document.getElementById("not-dialog"));
            durumuSaklaVeYenile();
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });
})();
</script>
{% endblock %}
```


## `templates/ozet.html`


```html
{% extends "base.html" %}
{% block title %}Özet · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell ozet-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">PERFORMANS ÖZETİ</p>
            <h1>Üretim ve Teslim Özeti</h1>
            <p class="soluk">Workflow dağılımı, zamanında teslim oranı ve son sekiz haftalık plan/teslim görünümü.</p>
        </div>
    </section>

    <section class="istatistik-grid">
        <article class="istatistik kart-plana"><span>Plana Alındı</span><strong>{{ genel.plana_alindi }}</strong></article>
        <article class="istatistik kart-dizgide"><span>Dizgide</span><strong>{{ genel.dizgide }}</strong></article>
        <article class="istatistik kart-hazir"><span>Hazır</span><strong>{{ genel.hazir }}</strong></article>
        <article class="istatistik kart-teslim"><span>Teslim Edildi</span><strong>{{ genel.teslim }}</strong></article>
    </section>

    {% if genel.gecikme or genel.durumu_eksik %}
    <section class="ozet-uyari-satiri">
        {% if genel.gecikme %}<span class="durum-rozet kotu">{{ genel.gecikme }} geciken açık kart</span>{% endif %}
        {% if genel.durumu_eksik %}<span class="durum-rozet uyari">{{ genel.durumu_eksik }} durumu eksik kart</span>{% endif %}
    </section>
    {% endif %}

    <section class="donem-grid">
    {% for ad, d in donemler.items() %}
        <article class="panel-kutu ozet-donem-karti">
            <p class="ust-etiket">{{ ad|upper }}</p>
            <h2>{{ d.kart }} teslim</h2>
            <div class="donem-metrik"><span>Teslim edilen adet</span><strong>{{ d.adet }}</strong></div>
            <div class="donem-metrik"><span>Zamanında teslim</span><strong>%{{ d.zamaninda_yuzde }}</strong></div>
            <div class="donem-metrik"><span>Ort. teslim sapması</span><strong>{{ d.ort_sapma }} gün</strong></div>
        </article>
    {% endfor %}
    </section>

    <section class="panel-kutu ozet-grafik-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">8 HAFTALIK GÖRÜNÜM</p>
                <h2>Planlanan Teslim / Gerçekleşen Teslim</h2>
            </div>
        </div>
        <div class="grafik-aciklama">
            <span><i class="lejant plan"></i>Planlanan</span>
            <span><i class="lejant teslim"></i>Teslim Edilen</span>
        </div>
        <div class="hafta-grafik" role="img" aria-label="Son sekiz hafta planlanan ve teslim edilen kart grafiği">
        {% for h in haftalar %}
            {% set plan_yuzde = (h.planlanan / en_yuksek * 100) if en_yuksek else 0 %}
            {% set teslim_yuzde = (h.teslim / en_yuksek * 100) if en_yuksek else 0 %}
            <div class="hafta-sutun">
                <div class="sutun-alani">
                    <div class="sutun plan" style="height: {{ plan_yuzde }}%"><span>{{ h.planlanan }}</span></div>
                    <div class="sutun teslim" style="height: {{ teslim_yuzde }}%"><span>{{ h.teslim }}</span></div>
                </div>
                <strong>H{{ h.hafta_no }}</strong>
                <small>{{ h.etiket }}</small>
            </div>
        {% endfor %}
        </div>
    </section>

    <section class="panel-kutu ozet-geciken-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">DİKKAT GEREKTİREN</p>
                <h2>Geciken Açık Kartlar</h2>
            </div>
            <span class="sayi-rozet">{{ geciken_kartlar|length }}</span>
        </div>

        {% if geciken_kartlar %}
        <div class="tablo-kapsayici">
            <table>
                <thead><tr><th>Talep NO</th><th>Stok No</th><th>Durum</th><th>Plan Başlangıç</th><th>Plan Teslim</th><th>İlerleme</th><th>Değerlendirme</th></tr></thead>
                <tbody>
                {% for k in geciken_kartlar %}
                    <tr>
                        <td>{{ k.talep_no or "—" }}</td>
                        <td>{{ k.stok_no or "—" }}</td>
                        <td>{{ k.durum }}</td>
                        <td>{{ k.plan_baslama|gun }}</td>
                        <td>{{ k.plan_teslim|gun }}</td>
                        <td>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }}</td>
                        <td><span class="durum-rozet kotu">{{ k.rozet }}</span></td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
        <div class="bos-durum">Şu anda geciken açık kart bulunmuyor.</div>
        {% endif %}
    </section>
</div>
{% endblock %}
```


## `templates/panel.html`


```html
{% extends "base.html" %}
{% block title %}Pano · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell panel-sayfa">

    <section class="sayfa-baslik sayfa-hero panel-hero">
        <div>
            <p class="ust-etiket">CANLI ÜRETİM PANOSU</p>
            <h1>İş Durumu</h1>
            <div class="panel-canli-satir">
                <span class="panel-canli">Güncel pano</span>
                <span class="panel-guncelleme">Son görüntüleme: {{ guncelleme }} · {{ bugun }}</span>
                {% if sayac.gecikme %}<span class="durum-rozet kotu">{{ sayac.gecikme }} geciken açık kart</span>{% endif %}
            </div>
        </div>
        <button id="panel-yenile" class="buton buton-hayalet" type="button">Yenile</button>
    </section>

    <section class="istatistik-grid istatistik-grid-3" aria-label="İş durumu özeti">
        <article class="istatistik kart-dizgide"><span>Dizgide</span><strong>{{ sayac.dizgide }}</strong></article>
        <article class="istatistik kart-plana"><span>Plana Alındı</span><strong>{{ sayac.plana_alindi }}</strong></article>
        <article class="istatistik kart-teslim"><span>Teslim Edildi</span><strong>{{ sayac.teslim }}</strong></article>
    </section>

    <section class="arac-cubugu panel-kutu panel-arac-cubugu">
        <div class="arama panel-arama-alani">
            <label for="panel-kart-ara">Kart ara</label>
            <input id="panel-kart-ara" type="search" placeholder="Talep no, stok no, talep sahibi, operatör..." autocomplete="off">
        </div>
        <div class="panel-filtre-alani">
            <div class="filtreler" role="group" aria-label="Durum filtresi">
                <button class="filtre" data-panel-filtre="DİZGİDE" type="button">Dizgide</button>
                <button class="filtre" data-panel-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                <button class="filtre" data-panel-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
                <button class="filtre aktif" data-panel-filtre="HEPSI" type="button">Hepsi</button>
            </div>
            <div class="panel-filtre-alt">
                <span id="panel-sonuc-sayisi"></span>
                <button id="panel-temizle" class="metin-buton" type="button" hidden>Filtreleri temizle</button>
            </div>
        </div>
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="DİZGİDE">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">AKTİF ÜRETİM</p>
                <h2>Dizgide</h2>
                <p class="panel-aciklama">Üretimi devam eden kartlar.</p>
            </div>
            <span class="sayi-rozet">{{ dizgide|length }}</span>
        </div>

        {% if dizgide %}
        <div class="kart-listesi">
            {% for k in dizgide %}
            <article class="is-karti panel-is-karti {{ k.renk }}" data-panel-kart data-durum="DİZGİDE"
                     data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or '') ~ ' ' ~ (k.pcb or '') ~ ' ' ~ (k.dizgi_sorumlusu or ''))|lower }}">
                <div class="is-karti-ust">
                    <div>
                        <strong>{{ k.talep_no or "Talep yok" }}</strong>
                        <span>{{ k.stok_no or "Stok no yok" }}</span>
                    </div>
                    <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                </div>
                {% if k.elle_dizgi_mi or k.malzeme_bekliyor %}
                <div class="satir-alt-rozet">
                    {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                    {% if k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif %}
                </div>
                {% endif %}
                <div class="kart-bilgiler">
                    <span><b>Sahibi:</b> {{ k.talep_sahibi or "—" }}</span>
                    <span><b>Plan teslim:</b> {{ k.plan_teslim|gun }}</span>
                    <span><b>Operatör:</b> {{ k.operator or "—" }}</span>
                </div>
                <div class="ilerleme">
                    <div class="ilerleme-ust"><span>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span><strong>%{{ k.adet_yuzde }}</strong></div>
                    <div class="ilerleme-ray"><div class="ilerleme-dolgu" style="width: {{ [k.adet_yuzde, 100]|min }}%"></div></div>
                </div>
                {% if k.aciklama %}<p class="kart-not">{{ k.aciklama }}</p>{% endif %}
            </article>
            {% endfor %}
        </div>
        {% else %}
        <div class="bos-durum">Şu anda dizgide iş yok.</div>
        {% endif %}
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="PLANA ALINDI">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">ÜRETİM SIRASI</p>
                <h2>Plana Alınan İşler</h2>
                <p class="panel-aciklama">Dizgiye alınmayı bekleyen ilk 12 kart.</p>
            </div>
            <span class="sayi-rozet">{{ plana_alindi|length }}</span>
        </div>

        {% if plana_alindi %}
        <div class="tablo-kapsayici panel-tablo-scroll">
            <table>
                <thead><tr><th>Talep No</th><th>Stok No</th><th>Talep Sahibi</th><th>Adet</th><th>Başlangıç</th><th>Teslim</th><th>Değerlendirme</th></tr></thead>
                <tbody>
                {% for k in plana_alindi %}
                    <tr data-panel-kart data-durum="PLANA ALINDI"
                        data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.pcb or '') ~ ' ' ~ (k.dizgi_sorumlusu or ''))|lower }}">
                        <td><strong>{{ k.talep_no or "—" }}</strong></td>
                        <td>{{ k.stok_no or "—" }}</td>
                        <td>{{ k.talep_sahibi or "—" }}</td>
                        <td>{{ k.toplam_adet }}</td>
                        <td>{{ k.plan_baslama|gun }}</td>
                        <td>{{ k.plan_teslim|gun }}</td>
                        <td>
                            <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                            {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                            {% if k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif %}
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
        <div class="bos-durum">Plana alınmış bekleyen kart yok.</div>
        {% endif %}
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="TESLİM EDİLDİ">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">TESLİM GEÇMİŞİ</p>
                <h2>Teslim Edilenler</h2>
                <p class="panel-aciklama">Seçilen döneme göre teslim edilen kartlar ve performans özeti.</p>
            </div>
            <span class="sayi-rozet" id="panel-donem-sayi">{{ teslim_edilen|length }}</span>
        </div>

        <div class="donem-filtre-cubugu">
            <div class="filtreler" role="group" aria-label="Dönem filtresi">
                <button class="filtre aktif" data-donem-filtre="tumu" type="button">Tümü</button>
                <button class="filtre" data-donem-filtre="hafta" type="button">Bu Hafta</button>
                <button class="filtre" data-donem-filtre="ay" type="button">Bu Ay</button>
                <button class="filtre" data-donem-filtre="yil" type="button">Bu Yıl</button>
                <button class="filtre" data-donem-filtre="ozel" type="button">Özel Aralık</button>
            </div>
            <div id="donem-ozel-alan" class="donem-ozel-alan" hidden>
                <input type="text" id="donem-baslangic" inputmode="numeric"
                       placeholder="gg.aa.yyyy" maxlength="10" aria-label="Başlangıç tarihi">
                <span>—</span>
                <input type="text" id="donem-bitis" inputmode="numeric"
                       placeholder="gg.aa.yyyy" maxlength="10" aria-label="Bitiş tarihi">
                <button id="donem-uygula" class="buton buton-kucuk buton-ana" type="button">Uygula</button>
            </div>
        </div>

        <div class="donem-metrik-grid donem-metrik-grid-4">
            <div class="donem-metrik-kart">
                <span>Teslim edilen iş emri</span>
                <strong id="donem-metrik-is">{{ sayac.teslim }}</strong>
                <small id="donem-metrik-adet-alt">— adet kart teslim edildi</small>
            </div>
            <div class="donem-metrik-kart">
                <span>Zamanında teslim</span>
                <strong id="donem-metrik-zamaninda">—</strong>
                <small id="donem-metrik-zamaninda-alt">—</small>
            </div>
            <div class="donem-metrik-kart">
                <span>Geciken teslim</span>
                <strong id="donem-metrik-gecikme">—</strong>
                <small id="donem-metrik-gecikme-alt">—</small>
            </div>
            <div class="donem-metrik-kart">
                <span>Ort. teslim sapması</span>
                <strong id="donem-metrik-sapma">—</strong>
                <small>+ geç · − erken</small>
            </div>
        </div>
        <div id="donem-veri-uyari" class="donem-veri-uyari" hidden></div>

        <div class="tablo-kapsayici panel-tablo-scroll">
            <table>
                <thead><tr><th>Talep No</th><th>Stok No</th><th>Adet</th><th>Teslim</th><th>Değerlendirme</th></tr></thead>
                <tbody id="donem-tablo-govde">
                {% for k in teslim_edilen %}
                <tr data-panel-kart data-durum="TESLİM EDİLDİ"
                    data-arama="{{ ((k.talep_no or '') ~ ' ' ~ (k.stok_no or '') ~ ' ' ~ (k.talep_sahibi or '') ~ ' ' ~ (k.operator or '') ~ ' ' ~ (k.aciklama or '') ~ ' ' ~ (k.dizgi_sorumlusu or ''))|lower }}">
                    <td><strong>{{ k.talep_no or "—" }}</strong></td>
                    <td>{{ k.stok_no or "—" }}</td>
                    <td>{{ k.toplam_adet }}</td>
                    <td>{{ (k.gerceklesen_teslim or k.teslim_zamani)|gun }}</td>
                    <td>
                        <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                        {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                    </td>
                </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        <div id="donem-bos" class="bos-durum" hidden>Bu dönemde teslim edilen kart bulunmuyor.</div>
    </section>

    <div id="panel-arama-bos" class="bos-durum" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
</div>
{% endblock %}

{% block scripts %}
<script>
(() => {
    const arama = document.getElementById("panel-kart-ara");
    const filtreler = [...document.querySelectorAll("[data-panel-filtre]")];
    const bolumler = [...document.querySelectorAll("[data-panel-bolum]")];
    const sonuc = document.getElementById("panel-sonuc-sayisi");
    const temizle = document.getElementById("panel-temizle");
    const bos = document.getElementById("panel-arama-bos");

    const FILTRE_KEY = "pdgm-panel-filtre";
    const ARAMA_KEY = "pdgm-panel-arama";
    const SCROLL_KEY = "pdgm-panel-scroll";
    const DONEM_KEY = "pdgm-panel-donem";

    let aktifFiltre = sessionStorage.getItem(FILTRE_KEY) || "HEPSI";

    // Teslim tablosu JS ile yeniden kurulduğu için kart listesi her seferinde
    // canlı olarak sorgulanır; sabit bir dizi tutulursa DOM'dan kopmuş
    // satırlara referans kalır ve sayaç yanlış çalışır.
    function tumKartlar() {
        return [...document.querySelectorAll("[data-panel-kart]")];
    }

    function durumUygun(durum) {
        if (aktifFiltre === "HEPSI") return true;
        return durum === aktifFiltre;
    }

    function filtrele() {
        const metin = (arama.value || "").trim().toLocaleLowerCase("tr-TR");
        let gorunen = 0;

        tumKartlar().forEach((kart) => {
            const uygun = durumUygun(kart.dataset.durum)
                && (!metin || (kart.dataset.arama || "").includes(metin));
            kart.hidden = !uygun;
            if (uygun) gorunen += 1;
        });

        bolumler.forEach((bolum) => {
            const bolumKartlari = [...bolum.querySelectorAll("[data-panel-kart]")];
            const bolumDurumUygun = durumUygun(bolum.dataset.durum);
            bolum.hidden = !bolumDurumUygun
                || (bolumKartlari.length > 0 && !bolumKartlari.some((kart) => !kart.hidden));
        });

        sonuc.textContent = gorunen ? `${gorunen} kart gösteriliyor` : "Sonuç bulunamadı";
        bos.hidden = gorunen !== 0;
        temizle.hidden = !metin && aktifFiltre === "HEPSI";
    }

    function filtreSec(deger) {
        aktifFiltre = deger;
        sessionStorage.setItem(FILTRE_KEY, deger);
        filtreler.forEach((buton) => {
            const secili = buton.dataset.panelFiltre === deger;
            buton.classList.toggle("aktif", secili);
            buton.setAttribute("aria-pressed", secili ? "true" : "false");
        });
        filtrele();
    }

    function scrollKaydet() {
        sessionStorage.setItem(SCROLL_KEY, String(window.scrollY));
    }

    arama.value = sessionStorage.getItem(ARAMA_KEY) || "";
    arama.addEventListener("input", () => {
        sessionStorage.setItem(ARAMA_KEY, arama.value || "");
        filtrele();
    });

    filtreler.forEach((buton) => {
        buton.addEventListener("click", () => filtreSec(buton.dataset.panelFiltre));
    });

    temizle.addEventListener("click", () => {
        arama.value = "";
        sessionStorage.removeItem(ARAMA_KEY);
        filtreSec("HEPSI");
        arama.focus();
    });

    document.getElementById("panel-yenile").addEventListener("click", () => {
        scrollKaydet();
        window.location.reload();
    });

    const kayitliScroll = sessionStorage.getItem(SCROLL_KEY);
    if (kayitliScroll !== null) {
        sessionStorage.removeItem(SCROLL_KEY);
        window.scrollTo(0, Number(kayitliScroll) || 0);
    }

    filtreSec(filtreler.some((b) => b.dataset.panelFiltre === aktifFiltre) ? aktifFiltre : "HEPSI");

    // ------------------------------------------------------------------
    // Dönem filtresi: Tümü / Bu Hafta / Bu Ay / Bu Yıl / Özel Aralık
    // ------------------------------------------------------------------
    const donemButonlar = [...document.querySelectorAll("[data-donem-filtre]")];
    const donemOzelAlan = document.getElementById("donem-ozel-alan");
    const donemBaslangic = document.getElementById("donem-baslangic");
    const donemBitis = document.getElementById("donem-bitis");
    const donemUygula = document.getElementById("donem-uygula");
    const donemTabloGovde = document.getElementById("donem-tablo-govde");
    const donemBos = document.getElementById("donem-bos");
    const donemSayi = document.getElementById("panel-donem-sayi");
    const donemMetrikIs = document.getElementById("donem-metrik-is");
    const donemMetrikAdetAlt = document.getElementById("donem-metrik-adet-alt");
    const donemMetrikZamaninda = document.getElementById("donem-metrik-zamaninda");
    const donemMetrikZamanindaAlt = document.getElementById("donem-metrik-zamaninda-alt");
    const donemMetrikGecikme = document.getElementById("donem-metrik-gecikme");
    const donemMetrikGecikmeAlt = document.getElementById("donem-metrik-gecikme-alt");
    const donemMetrikSapma = document.getElementById("donem-metrik-sapma");
    const donemVeriUyari = document.getElementById("donem-veri-uyari");

    let aktifDonem = sessionStorage.getItem(DONEM_KEY) || "tumu";

    // -- gg.aa.yyyy maskeleme yardımcıları -------------------------------
    function tarihMaskele(input) {
        input.addEventListener("input", () => {
            let rakam = input.value.replace(/\D/g, "").slice(0, 8);
            let parcalar = [];
            if (rakam.length > 0) parcalar.push(rakam.slice(0, 2));
            if (rakam.length > 2) parcalar.push(rakam.slice(2, 4));
            if (rakam.length > 4) parcalar.push(rakam.slice(4, 8));
            input.value = parcalar.join(".");
        });
    }

    function ggAaYyyyGecerliMi(deger) {
        return /^\d{2}\.\d{2}\.\d{4}$/.test(deger || "");
    }

    function isoyaCevir(deger) {
        const [gg, aa, yyyy] = deger.split(".");
        return `${yyyy}-${aa}-${gg}`;
    }

    tarihMaskele(donemBaslangic);
    tarihMaskele(donemBitis);
    // ---------------------------------------------------------------------

    function hucreEkle(satir, deger) {
        const td = document.createElement("td");
        td.textContent = (deger === null || deger === undefined || deger === "") ? "—" : deger;
        satir.appendChild(td);
    }

    function rozetEkle(satir, renk, metin) {
        const td = document.createElement("td");
        const span = document.createElement("span");
        span.className = `durum-rozet ${renk || "notr"}`;
        span.textContent = metin || "";
        td.appendChild(span);
        satir.appendChild(td);
    }

    function satirKur(k) {
    const satir = document.createElement("tr");
    satir.setAttribute("data-panel-kart", "");
    satir.dataset.durum = "TESLİM EDİLDİ";
    satir.dataset.arama = [k.talep_no, k.stok_no, k.operator, k.dizgi_sorumlusu]
        .filter(Boolean)
        .join(" ")
        .toLocaleLowerCase("tr-TR");

    const talepTd = document.createElement("td");
    const strong = document.createElement("strong");
    strong.textContent = k.talep_no || "—";
    talepTd.appendChild(strong);
    satir.appendChild(talepTd);

    hucreEkle(satir, k.stok_no);
    hucreEkle(satir, k.toplam_adet);
    hucreEkle(satir, k.teslim);

    const rozetTd = document.createElement("td");
    const anaRozet = document.createElement("span");
    anaRozet.className = `durum-rozet ${k.renk || "notr"}`;
    anaRozet.textContent = k.rozet || "";
    rozetTd.appendChild(anaRozet);
    if (k.elle_dizgi_mi) {
        const elleRozet = document.createElement("span");
        elleRozet.className = "durum-rozet elle";
        elleRozet.textContent = "Elle Dizgi";
        rozetTd.appendChild(elleRozet);
    }
    satir.appendChild(rozetTd);

    return satir;
}

    async function donemYukle(aralik) {
        const params = new URLSearchParams({aralik});

        if (aralik === "ozel") {
            if (!ggAaYyyyGecerliMi(donemBaslangic.value) || !ggAaYyyyGecerliMi(donemBitis.value)) {
                toast("Tarihleri gg.aa.yyyy biçiminde girin (örn. 05.03.2026).", "uyari");
                return;
            }
            params.set("baslangic", isoyaCevir(donemBaslangic.value));
            params.set("bitis", isoyaCevir(donemBitis.value));
        }

        try {
            const data = await pdgmFetch(`/api/panel/teslimler?${params.toString()}`);
            const liste = data.teslim_edilen || [];

            donemTabloGovde.replaceChildren();
            liste.forEach((k) => donemTabloGovde.appendChild(satirKur(k)));

            donemBos.hidden = liste.length !== 0;

            const o = data.ozet;
            donemSayi.textContent = o.kart;
            donemMetrikIs.textContent = o.kart;
            donemMetrikAdetAlt.textContent = `${o.adet.toLocaleString("tr-TR")} adet kart teslim edildi`;

            donemMetrikZamaninda.textContent = `%${o.zamaninda_yuzde}`;
            donemMetrikZamanindaAlt.textContent =
                `${o.zamaninda} iş emri · ${o.zamaninda_adet.toLocaleString("tr-TR")} adet`;

            donemMetrikGecikme.textContent = o.gecikmeli;
            donemMetrikGecikmeAlt.textContent = `${o.gecikmeli_adet.toLocaleString("tr-TR")} adet`;

            donemMetrikSapma.textContent = o.sapma_olculen ? `${o.ort_sapma} gün` : "—";

            const olculemeyen = o.kart - o.sapma_olculen;
            donemVeriUyari.hidden = !olculemeyen;
            donemVeriUyari.textContent = olculemeyen
                ? `${olculemeyen} iş emrinde plan teslim tarihi girilmediği için sapma hesaplanamadı; ` +
                  `yüzde yalnız ${o.sapma_olculen} iş emri üzerinden hesaplandı.`
                : "";

            filtrele();
        } catch (hata) {
            hataMesaji(hata);
        }
    }

    function donemSec(deger) {
        aktifDonem = deger;
        sessionStorage.setItem(DONEM_KEY, deger);

        donemButonlar.forEach((buton) => {
            const secili = buton.dataset.donemFiltre === deger;
            buton.classList.toggle("aktif", secili);
            buton.setAttribute("aria-pressed", secili ? "true" : "false");
        });

        donemOzelAlan.hidden = deger !== "ozel";
        if (deger !== "ozel") donemYukle(deger);
    }

    donemButonlar.forEach((buton) => {
        buton.addEventListener("click", () => donemSec(buton.dataset.donemFiltre));
    });
    donemUygula.addEventListener("click", () => donemYukle("ozel"));

    donemSec(donemButonlar.some((b) => b.dataset.donemFiltre === aktifDonem) ? aktifDonem : "tumu");
})();
</script>
{% endblock %}
```


## `templates/yetkisiz.html`


```html
{% extends "base.html" %}
{% block title %}Yetkisiz · PDGM İş Takip{% endblock %}
{% block content %}
<div class="yetkisiz-sayfa">
    <section class="durum-sayfasi">
        <p class="ust-etiket">ERİŞİM KISITLI</p>
        <div class="durum-ikon">403</div>
        <h1>Bu sayfaya erişim yetkiniz yok.</h1>
        <p>Mevcut hesabınız bu işlemi gerçekleştirmek için gerekli role sahip değil.</p>
        <a class="buton buton-ana" href="{{ url_for('ana') }}">Ana Sayfaya Dön</a>
    </section>
</div>
{% endblock %}
```


## `templates/yonetim.html`


```html
{% extends "base.html" %}
{% block title %}Yönetim · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell yonetim-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">YÖNETİCİ PANELİ</p>
            <h1>Sistem Yönetimi</h1>
            <p class="soluk">İlk Excel aktarımı, manuel kart yönetimi, kartlar.xlsx bakımı, yedekler ve audit geçmişi.</p>
        </div>
        <div class="baslik-aksiyon">
            <a class="buton buton-hayalet" href="{{ url_for('rapor_indir') }}">Rapor İndir</a>
            <a class="buton buton-ana" href="{{ url_for('panel') }}">Canlı Panoyu Aç</a>
        </div>
    </section>

    <section class="yonetim-grid">
        <article class="panel-kutu yonetim-yukleme-karti">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">VERİ AKTARIMI</p>
                    <h2>Plan Excel'ini Aktar</h2>
                    <p class="panel-aciklama">MAKİNE ve ELDE DİZGİ YENİ sayfaları okunur; hidden kolonlar atlanır. ELDE DİZGİ YENİ sayfası opsiyoneldir.</p>
                </div>
            </div>
            <form method="post" action="{{ url_for('yukle') }}" enctype="multipart/form-data" class="yukleme-form">
                <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                <label class="dosya-sec">
                    <span>Excel dosyası seçin</span>
                    <small>.xlsx veya .xlsm</small>
                    <input type="file" name="dosya" accept=".xlsx,.xlsm" required>
                </label>
                <button class="buton buton-ana" type="submit">Excel'i Aktar</button>
            </form>
            <p class="yardim">
                Kaynak Excel ilk kurulumda kartları oluşturmak için kullanılabilir. Sonrasında günlük operasyonun source of truth'u
                <code>data/kartlar.xlsx</code> olur. Tekrar import yapılırsa mevcut workflow, tamamlanan adet ve operatör işlemleri korunur;
                yalnız plan/source alanları güncellenir.
            </p>
        </article>

        <article class="panel-kutu yonetim-bakim-karti">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">BAKIM</p>
                    <h2>Kayıt Dosyaları</h2>
                    <p class="panel-aciklama">Excel'e doğrudan müdahale gerektiğinde kullanılacak dosyalar.</p>
                </div>
            </div>
            <div class="buton-grup dikey">
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='kartlar') }}">Kartlar Excel</a>
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='log') }}">İşlem Logu</a>
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='yuklemeler') }}">Yükleme Geçmişi</a>
                <form method="post" action="{{ url_for('yeniden_oku') }}"
                      onsubmit="return confirm('Diskteki kartlar.xlsx doğrulanarak yeniden okunacak. Devam edilsin mi?')">
                    <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                    <button class="buton buton-uyari tam-genislik" type="submit">Kart Dosyasını Yeniden Oku</button>
                </form>
            </div>
            <p class="yardim">Manuel Excel düzenlemesi yaparken önce dosyayı kaydedip Excel'de kapatın; ardından "Kart Dosyasını Yeniden Oku" kullanın.</p>
        </article>
    </section>

    {% if durumu_eksik %}
    <section class="panel-kutu yonetim-durum-eksik">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KONTROL GEREKİYOR</p>
                <h2>Durumu Eksik Kartlar</h2>
                <p class="panel-aciklama">Kaynak Excel'de DURUM boş veya geçersiz olduğu için Pano ve Operatör ekranında gösterilmezler.</p>
            </div>
            <span class="sayi-rozet">{{ durumu_eksik|length }}</span>
        </div>

        <div class="mini-liste yonetim-durum-eksik-liste">
            {% for k in durumu_eksik %}
            <div class="mini-satir">
                <div>
                    <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                    <span>{{ k.talep_sahibi or "Talep sahibi yok" }} · {{ k.toplam_adet }} adet</span>
                    <small>Kaynak DURUM: {{ k.excel_durum or "boş" }}</small>
                </div>
                <div class="mini-sag">
                    <span class="durum-rozet uyari">DURUMU EKSİK</span>
                    {% if k.elle_dizgi_mi %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                    <button class="buton buton-kucuk buton-ana" type="button" data-admin-duzenle
                            data-id="{{ k.id }}" data-talep="{{ k.talep_no or '' }}" data-stok="{{ k.stok_no or '' }}"
                            data-durum="" data-toplam="{{ k.toplam_adet }}" data-tamamlanan="{{ k.tamamlanan_adet }}"
                            data-plan-hafta="{{ k.plan_hafta or '' }}" data-plan-baslama="{{ k.plan_baslama or '' }}"
                            data-plan-teslim="{{ k.plan_teslim or '' }}" data-gerceklesen-teslim="{{ k.gerceklesen_teslim or '' }}"
                            data-elle="{{ '1' if k.elle_dizgi_mi else '0' }}" data-dizgi-sorumlusu="{{ k.dizgi_sorumlusu or '' }}"
                            data-not="{{ k.aciklama or '' }}">Durum Ata</button>
                </div>
            </div>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    {% if kaynakta_olmayan %}
    <section class="panel-kutu yonetim-uyari-kutu">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KAYNAK KONTROLÜ</p>
                <h2>Son Excel'de Olmayan Açık Kartlar</h2>
                <p class="panel-aciklama">Kartlar silinmez; mevcut operasyon state'i korunur.</p>
            </div>
            <span class="sayi-rozet">{{ kaynakta_olmayan|length }}</span>
        </div>
        <div class="mini-liste">
            {% for k in kaynakta_olmayan %}
            <div class="mini-satir">
                <div>
                    <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                    <span>{{ k.is_durumu }} · {{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span>
                </div>
                <div class="mini-sag"><span class="durum-rozet uyari">Kaynak Excel'de yok</span></div>
            </div>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    <section class="yonetim-grid">
        <article class="panel-kutu">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">KART YÖNETİMİ</p>
                    <h2>Gizlenen Kartlar</h2>
                    <p class="panel-aciklama">Gizlemek silme işlemi değildir; kart kartlar.xlsx içinde kalır.</p>
                </div>
                <span class="sayi-rozet">{{ gizlenen_kartlar|length }}</span>
            </div>

            {% if gizlenen_kartlar %}
            <div class="mini-liste">
                {% for k in gizlenen_kartlar %}
                <div class="mini-satir">
                    <div>
                        <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                        <span>{{ k.is_durumu }} · {{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span>
                    </div>
                    <div class="mini-sag">
                        <span class="durum-rozet">Gizli</span>
                        <form method="post" action="{{ url_for('kart_geri_getir') }}"
                              onsubmit="return confirm('Bu kart tekrar aktif listelere alınsın mı?')">
                            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                            <input type="hidden" name="kart_id" value="{{ k.id }}">
                            <button type="submit" class="buton buton-kucuk buton-basari">Geri Getir</button>
                        </form>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Gizlenmiş kart bulunmuyor.</div>
            {% endif %}
        </article>

        <article class="panel-kutu">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">KURTARMA</p>
                    <h2>Yedekten Geri Yükle</h2>
                    <p class="panel-aciklama">Yalnız kart verisi geri alınır; audit logu geriye sarılmaz.</p>
                </div>
            </div>

            {% if yedekler %}
            <div class="mini-liste">
                {% for y in yedekler %}
                <div class="mini-satir">
                    <div>
                        <strong>{{ y.zaman }}</strong>
                        <span>{{ y.tip }} · {{ y.etiket }}</span>
                        <small>{{ y.boyut_kb }} KB</small>
                    </div>
                    <div class="mini-sag">
                        <form method="post" action="{{ url_for('yedek_geri_yukle') }}"
                              onsubmit="return confirm('Mevcut kart verileri seçilen yedekle değiştirilecek. Mevcut durum önce ayrıca yedeklenecek. Devam edilsin mi?')">
                            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                            <input type="hidden" name="yedek" value="{{ y.ad }}">
                            <button type="submit" class="buton buton-kucuk buton-uyari">Geri Yükle</button>
                        </form>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Henüz kullanılabilir yedek bulunmuyor.</div>
            {% endif %}
        </article>
    </section>

    <section class="panel-kutu yonetim-kart-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KART YÖNETİMİ</p>
                <h2>Kartlar</h2>
                <p class="panel-aciklama">
                    Gizlenmemiş tüm aktif kayıtlar; durumu eksik kayıtlar dahil.
                </p>
            </div>

            <div class="baslik-aksiyon">
                <button id="admin-yeni-ac"
                        class="buton buton-ana"
                        type="button">
                    + Yeni Kart
                </button>

                <div class="arama kucuk">
                    <label class="sr-only" for="admin-kart-ara">Kart ara</label>
                    <input id="admin-kart-ara"
                           type="search"
                           placeholder="Talep, stok, kişi..."
                           autocomplete="off">
                </div>
            </div>
        </div>

        <div class="admin-filtre-cubugu">
            <div class="filtreler"
                 role="group"
                 aria-label="Yönetim kart durum filtresi">

                <button class="filtre aktif"
                        data-admin-filtre="HEPSI"
                        type="button">
                    Hepsi
                </button>

                <button class="filtre"
                        data-admin-filtre="AKTIF"
                        type="button">
                    Açık İşler
                </button>

                <button class="filtre"
                        data-admin-filtre="PLANA ALINDI"
                        type="button">
                    Plana Alındı
                </button>

                <button class="filtre"
                        data-admin-filtre="DİZGİDE"
                        type="button">
                    Dizgide
                </button>

                <button class="filtre"
                        data-admin-filtre="HAZIR"
                        type="button">
                    Hazır
                </button>

                <button class="filtre"
                        data-admin-filtre="TESLİM EDİLDİ"
                        type="button">
                    Teslim Edildi
                </button>

                <button class="filtre"
                        data-admin-filtre="DURUMU EKSİK"
                        type="button">
                    Durumu Eksik
                </button>

                <button class="filtre"
                        data-admin-filtre="ELLE DİZGİ"
                        type="button">
                    Elle Dizgi
                </button>
            </div>

            <span id="admin-sonuc" class="admin-filtre-sonuc">
                {{ kartlar|length }} kart
            </span>
        </div>

        <div class="tablo-kapsayici yonetim-kart-scroll">
            <table id="admin-kart-tablosu">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Talep NO</th>
                        <th>Stok No</th>
                        <th>Talep Sahibi</th>
                        <th>Adet</th>
                        <th>Durum</th>
                        <th>Kaynak Durumu</th>
                        <th>Dizgi</th>
                        <th>Plan Teslim</th>
                        <th>Operatör</th>
                        <th>İşlem</th>
                    </tr>
                </thead>

                <tbody>
                {% for k in kartlar %}
                    <tr data-admin-kart
                        data-durum="{{ k.durum or 'DURUMU EKSİK' }}"
                        data-elle="{{ '1' if k.elle_dizgi_mi else '0' }}"
                        data-arama="{{ (
                            (k.id|string)
                            ~ ' ' ~ (k.talep_no or '')
                            ~ ' ' ~ (k.stok_no or '')
                            ~ ' ' ~ (k.talep_sahibi or '')
                            ~ ' ' ~ (k.durum or 'DURUMU EKSİK')
                            ~ ' ' ~ (k.excel_durum or '')
                            ~ ' ' ~ (k.dizgi_sorumlusu or '')
                        )|lower }}">

                        <td>{{ k.id }}</td>

                        <td>
                            <strong>{{ k.talep_no or "—" }}</strong>
                        </td>

                        <td>{{ k.stok_no or "—" }}</td>

                        <td>{{ k.talep_sahibi or "—" }}</td>

                        <td>
                            {{ k.tamamlanan_adet }}/{{ k.toplam_adet }}
                        </td>

                        <td>
                            {% if k.durum %}
                                <span class="durum-rozet {{ k.renk }}">
                                    {{ k.durum }}
                                </span>
                            {% else %}
                                <span class="durum-rozet uyari">
                                    DURUMU EKSİK
                                </span>
                            {% endif %}

                            {% if k.kaynakta_yok %}
                                <div class="satir-alt-rozet">
                                    <span class="durum-rozet uyari">
                                        Kaynakta yok
                                    </span>
                                </div>
                            {% endif %}
                        </td>

                        <td>{{ k.kaynak_durumu or "—" }}</td>

                        <td>
                            {% if k.elle_dizgi_mi %}
                                <span class="durum-rozet elle">Elle Dizgi</span>
                                {% if k.dizgi_sorumlusu %}
                                    <div class="satir-alt-rozet"><small>{{ k.dizgi_sorumlusu }}</small></div>
                                {% endif %}
                            {% else %}
                                <span class="durum-rozet">Makine</span>
                            {% endif %}
                            {% if k.malzeme_bekliyor %}
                                <div class="satir-alt-rozet">
                                    <span class="durum-rozet uyari">Malzeme Bekliyor</span>
                                </div>
                            {% endif %}
                        </td>

                        <td>{{ k.plan_teslim|gun }}</td>

                        <td>{{ k.operator or "—" }}</td>

                        <td>
                            <div class="satir-aksiyon">
                                <button class="buton buton-kucuk buton-hayalet"
                                        type="button"
                                        data-admin-duzenle
                                        data-id="{{ k.id }}"
                                        data-talep="{{ k.talep_no or '' }}"
                                        data-stok="{{ k.stok_no or '' }}"
                                        data-durum="{{ k.durum or '' }}"
                                        data-toplam="{{ k.toplam_adet }}"
                                        data-tamamlanan="{{ k.tamamlanan_adet }}"
                                        data-plan-hafta="{{ k.plan_hafta or '' }}"
                                        data-plan-baslama="{{ k.plan_baslama or '' }}"
                                        data-plan-teslim="{{ k.plan_teslim or '' }}"
                                        data-gerceklesen-teslim="{{ k.gerceklesen_teslim or '' }}"
                                        data-elle="{{ '1' if k.elle_dizgi_mi else '0' }}"
                                        data-dizgi-sorumlusu="{{ k.dizgi_sorumlusu or '' }}"
                                        data-not="{{ k.aciklama or '' }}">
                                    Düzenle
                                </button>

                                <button class="buton buton-kucuk buton-tehlike"
                                        type="button"
                                        data-admin-gizle
                                        data-id="{{ k.id }}"
                                        data-talep="{{ k.talep_no or '' }}">
                                    Gizle
                                </button>
                            </div>
                        </td>
                    </tr>
                {% else %}
                    <tr>
                        <td colspan="11" class="bos-durum">
                            Kart bulunmuyor.
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <div id="admin-kart-bos"
             class="bos-durum admin-kart-bos"
             hidden>
            Arama ve filtreye uyan kart bulunamadı.
        </div>
    </section>

    <section class="yonetim-grid yonetim-gecmis">
        <article class="panel-kutu">
            <div class="panel-baslik"><div><p class="ust-etiket">SON YÜKLEMELER</p><h2>Excel Geçmişi</h2></div></div>
            {% if yuklemeler %}
            <div class="mini-liste">
                {% for y in yuklemeler %}
                <div class="mini-satir">
                    <div><strong>{{ y.dosya or "—" }}</strong><span>{{ y.zaman or "—" }} · {{ y.kullanici or "—" }}</span></div>
                    <div class="mini-sag"><strong>{{ y.satir or 0 }} satır</strong><small>{{ y.yeni or 0 }} yeni · {{ y.guncellenen or 0 }} güncel · {{ y.uyari or 0 }} uyarı</small></div>
                </div>
                {% endfor %}
            </div>
            {% else %}<div class="bos-durum">Henüz Excel yüklenmemiş.</div>{% endif %}
        </article>

        <article class="panel-kutu">
            <div class="panel-baslik"><div><p class="ust-etiket">AUDIT</p><h2>Son İşlemler</h2></div></div>
            {% if loglar %}
            <div class="log-liste">
                {% for l in loglar %}
                <div class="log-satir">
                    <div><strong>{{ l.islem or "İşlem" }}</strong><span>{{ l.kullanici or "—" }} · {{ l.zaman or "—" }}</span></div>
                    <small>{% if l.talep_no %}{{ l.talep_no }}{% endif %}{% if l.stok_no %} · {{ l.stok_no }}{% endif %}{% if l.detay %} · {{ l.detay }}{% endif %}</small>
                </div>
                {% endfor %}
            </div>
            {% else %}<div class="bos-durum">Henüz işlem kaydı yok.</div>{% endif %}
        </article>
    </section>
</div>

<dialog id="admin-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="admin-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">ADMİN MÜDAHALESİ</p><h2 id="admin-baslik">Kart Düzenle</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="admin-id">

        <div class="iki-kolon">
            <div class="alan">
                <label for="admin-durum">Durum</label>
                <select id="admin-durum" required>
                    <option value="" disabled>Durum seçin</option>
                    <option value="HAZIR">HAZIR</option>
                    <option value="PLANA ALINDI">PLANA ALINDI</option>
                    <option value="DİZGİDE">DİZGİDE</option>
                    <option value="TESLİM EDİLDİ">TESLİM EDİLDİ</option>
                </select>
            </div>
            <div class="alan">
                <label for="admin-toplam">Toplam adet</label>
                <input type="number" id="admin-toplam" min="1" required>
            </div>
        </div>

        <div class="alan">
            <label for="admin-tamamlanan">Tamamlanan adet</label>
            <input type="number" id="admin-tamamlanan" min="0" required>
            <small>HAZIR ve TESLİM EDİLDİ seçildiğinde sistem otomatik olarak toplam adede eşitler.</small>
        </div>

        <div class="alan">
            <label for="admin-plan-hafta">Plan Haftası</label>
            <input type="text" id="admin-plan-hafta" placeholder="Örn. 34. hafta (17.08 haftası)">
        </div>

        <div class="iki-kolon">
            <div class="alan"><label for="admin-plan-baslama">Dizgi Başlama Tarihi</label><input type="date" id="admin-plan-baslama"></div>
            <div class="alan"><label for="admin-plan-teslim">Planlanan Teslim Tarihi</label><input type="date" id="admin-plan-teslim"></div>
        </div>

        <div class="alan">
            <label for="admin-gerceklesen-teslim">Gerçekleşen Teslim Tarihi</label>
            <input type="date" id="admin-gerceklesen-teslim">
            <small>TESLİM EDİLDİ durumunda boş bırakılırsa bugünün tarihi kullanılır.</small>
        </div>

        <div class="alan"><label for="admin-not">Not</label><textarea id="admin-not" rows="4"></textarea></div>

        <div class="alan alan-checkbox">
            <label><input type="checkbox" id="admin-elle-dizgi"> Elle Dizgi</label>
        </div>
        <div class="alan">
            <label for="admin-dizgi-sorumlusu">PDGM Dizgi Sorumlusu</label>
            <input type="text" id="admin-dizgi-sorumlusu" placeholder="Yalnız Elle Dizgi kartlarda kullanılır">
        </div>

        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>

<dialog id="admin-yeni-dialog" class="modal">
    <form method="dialog" class="modal-kutu" id="admin-yeni-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">KART YÖNETİMİ</p><h2>Yeni Kart</h2><p class="soluk">Yeni kart PLANA ALINDI durumunda oluşturulur.</p></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-sira">NO / Sıra</label><input type="number" id="yeni-sira" step="1"></div>
            <div class="alan"><label for="yeni-talep-no">Talep NO *</label><input type="text" id="yeni-talep-no" required></div>
        </div>

        <div class="alan"><label for="yeni-talep-sahibi">Talep Sahibi</label><input type="text" id="yeni-talep-sahibi"></div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-stok-no">Kart Stok No *</label><input type="text" id="yeni-stok-no" required></div>
            <div class="alan"><label for="yeni-toplam">Toplam Adet *</label><input type="number" id="yeni-toplam" min="1" required></div>
        </div>

        <div class="alan"><label for="yeni-plan-hafta">Plan Haftası</label><input type="text" id="yeni-plan-hafta" placeholder="Örn. 34. hafta (17.08 haftası)"></div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-plan-baslama">Dizgi Başlama Tarihi</label><input type="date" id="yeni-plan-baslama"></div>
            <div class="alan"><label for="yeni-plan-teslim">Planlanan Teslim Tarihi</label><input type="date" id="yeni-plan-teslim"></div>
        </div>

        <div class="alan"><label for="yeni-pcb">PCB</label><input type="text" id="yeni-pcb" placeholder="Örn. HBT"></div>

        <div class="alan alan-checkbox">
            <label><input type="checkbox" id="yeni-elle-dizgi"> Elle Dizgi</label>
        </div>
        <div class="alan">
            <label for="yeni-dizgi-sorumlusu">PDGM Dizgi Sorumlusu</label>
            <input type="text" id="yeni-dizgi-sorumlusu" placeholder="Yalnız Elle Dizgi kartlarda kullanılır">
        </div>

        <div class="alan"><label for="yeni-not">Not</label><textarea id="yeni-not" rows="3"></textarea></div>

        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kartı Oluştur</button>
        </div>
    </form>
</dialog>
{% endblock %}

{% block scripts %}
<script>
(() => {
    const adminAra = document.getElementById("admin-kart-ara");
    const adminSatirlar = [...document.querySelectorAll("[data-admin-kart]")];
    const adminFiltreler = [...document.querySelectorAll("[data-admin-filtre]")];
    const adminSonuc = document.getElementById("admin-sonuc");
    const adminBos = document.getElementById("admin-kart-bos");

    const adminDialog = document.getElementById("admin-dialog");
    const adminForm = document.getElementById("admin-form");
    const yeniDialog = document.getElementById("admin-yeni-dialog");
    const yeniForm = document.getElementById("admin-yeni-form");

    let adminAktifFiltre = "HEPSI";

    function adminDurumUygun(satir) {
        if (adminAktifFiltre === "HEPSI") {
            return true;
        }

        if (adminAktifFiltre === "ELLE DİZGİ") {
            return satir.dataset.elle === "1";
        }

        if (adminAktifFiltre === "AKTIF") {
            return ["PLANA ALINDI", "DİZGİDE", "HAZIR"].includes(satir.dataset.durum);
        }

        return satir.dataset.durum === adminAktifFiltre;
    }

    function adminKartlariFiltrele() {
        const arama = adminAra.value
            .trim()
            .toLocaleLowerCase("tr-TR");

        let gorunen = 0;

        adminSatirlar.forEach((satir) => {
            const durumUygun = adminDurumUygun(satir);
            const aramaUygun = !arama || satir.dataset.arama.includes(arama);
            const uygun = durumUygun && aramaUygun;

            satir.hidden = !uygun;

            if (uygun) {
                gorunen += 1;
            }
        });

        adminSonuc.textContent =
            gorunen === adminSatirlar.length
                ? `${gorunen} kart`
                : `${gorunen} / ${adminSatirlar.length} kart`;

        adminBos.hidden = gorunen !== 0;
    }

    adminAra.addEventListener("input", adminKartlariFiltrele);

    adminFiltreler.forEach((buton) => {
        buton.addEventListener("click", () => {
            adminAktifFiltre = buton.dataset.adminFiltre;

            adminFiltreler.forEach((filtre) => {
                const aktif = filtre === buton;
                filtre.classList.toggle("aktif", aktif);
                filtre.setAttribute(
                    "aria-pressed",
                    aktif ? "true" : "false"
                );
            });

            adminKartlariFiltrele();
        });
    });

adminKartlariFiltrele();

    function duzenlemeDialogunuAc(buton) {
        document.getElementById("admin-id").value = buton.dataset.id;
        document.getElementById("admin-durum").value = buton.dataset.durum || "";
        document.getElementById("admin-toplam").value = buton.dataset.toplam;
        document.getElementById("admin-tamamlanan").value = buton.dataset.tamamlanan;
        document.getElementById("admin-plan-hafta").value = buton.dataset.planHafta || "";
        document.getElementById("admin-plan-baslama").value = buton.dataset.planBaslama || "";
        document.getElementById("admin-plan-teslim").value = buton.dataset.planTeslim || "";
        document.getElementById("admin-gerceklesen-teslim").value = buton.dataset.gerceklesenTeslim || "";
        document.getElementById("admin-not").value = buton.dataset.not || "";
        document.getElementById("admin-elle-dizgi").checked = buton.dataset.elle === "1";
        document.getElementById("admin-dizgi-sorumlusu").value = buton.dataset.dizgiSorumlusu || "";
        document.getElementById("admin-baslik").textContent = `${buton.dataset.talep || "Kart"} · ${buton.dataset.stok || ""}`;
        adminDialog.showModal();
    }

    document.querySelectorAll("[data-admin-duzenle]").forEach((buton) => {
        buton.addEventListener("click", () => duzenlemeDialogunuAc(buton));
    });

    document.querySelectorAll("[data-admin-gizle]").forEach((buton) => {
        buton.addEventListener("click", async () => {
            if (!confirm(`${buton.dataset.talep || "Bu kart"} listeden gizlensin mi?\n\nKart silinmez; kartlar.xlsx içinde tutulur.`)) return;

            buton.disabled = true;

            try {
                await pdgmFetch("/api/admin/kart-sil", {
                    method: "POST",
                    body: JSON.stringify({kart_id: Number(buton.dataset.id)})
                });
                toast("Kart listeden gizlendi.");
                window.setTimeout(() => window.location.reload(), 350);
            } catch (hata) {
                hataMesaji(hata);
                buton.disabled = false;
            }
        });
    });

    function durumAlanlariniAyarla() {
    const durum = document.getElementById("admin-durum").value;
    const toplam = Number(document.getElementById("admin-toplam").value || 0);
    const tamamlanan = document.getElementById("admin-tamamlanan");

    if (["PLANA ALINDI", "HAZIR"].includes(durum)) {
        tamamlanan.value = 0;
    } else if (durum === "TESLİM EDİLDİ" && toplam > 0) {
        tamamlanan.value = toplam;
    }
}

    document.getElementById("admin-durum").addEventListener("change", durumAlanlariniAyarla);
    document.getElementById("admin-toplam").addEventListener("input", durumAlanlariniAyarla);

    adminForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;

        try {
            await pdgmFetch("/api/admin/duzenle", {
                method: "POST",
                body: JSON.stringify({
                    kart_id: Number(document.getElementById("admin-id").value),
                    durum: document.getElementById("admin-durum").value,
                    toplam_adet: Number(document.getElementById("admin-toplam").value),
                    tamamlanan_adet: Number(document.getElementById("admin-tamamlanan").value),
                    plan_hafta: document.getElementById("admin-plan-hafta").value,
                    plan_baslama: document.getElementById("admin-plan-baslama").value,
                    plan_teslim: document.getElementById("admin-plan-teslim").value,
                    gerceklesen_teslim: document.getElementById("admin-gerceklesen-teslim").value,
                    not: document.getElementById("admin-not").value,
                    elle_dizgi: document.getElementById("admin-elle-dizgi").checked,
                    dizgi_sorumlusu: document.getElementById("admin-dizgi-sorumlusu").value
                })
            });

            toast("Kart güncellendi.");
            dialogKapat(adminDialog);
            window.setTimeout(() => window.location.reload(), 350);
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });

    document.getElementById("admin-yeni-ac").addEventListener("click", () => {
        yeniForm.reset();
        yeniDialog.showModal();
        window.setTimeout(() => document.getElementById("yeni-talep-no").focus(), 50);
    });

    yeniForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter;
        submit.disabled = true;

        try {
            const sira = document.getElementById("yeni-sira").value;

            await pdgmFetch("/api/admin/kart-ekle", {
                method: "POST",
                body: JSON.stringify({
                    sira: sira ? Number(sira) : null,
                    talep_no: document.getElementById("yeni-talep-no").value,
                    talep_sahibi: document.getElementById("yeni-talep-sahibi").value,
                    stok_no: document.getElementById("yeni-stok-no").value,
                    toplam_adet: Number(document.getElementById("yeni-toplam").value),
                    plan_hafta: document.getElementById("yeni-plan-hafta").value,
                    plan_baslama: document.getElementById("yeni-plan-baslama").value,
                    plan_teslim: document.getElementById("yeni-plan-teslim").value,
                    pcb: document.getElementById("yeni-pcb").value,
                    elle_dizgi: document.getElementById("yeni-elle-dizgi").checked,
                    dizgi_sorumlusu: document.getElementById("yeni-dizgi-sorumlusu").value,
                    not: document.getElementById("yeni-not").value
                })
            });

            toast("Yeni kart PLANA ALINDI durumunda oluşturuldu.");
            dialogKapat(yeniDialog);
            window.setTimeout(() => window.location.reload(), 350);
        } catch (hata) {
            hataMesaji(hata);
        } finally {
            submit.disabled = false;
        }
    });
})();
</script>
{% endblock %}
```


# 3. STATIC


## `static/stil.css`


```css
:root {
    --zemin: #f4f6f8;
    --kart: #ffffff;
    --yazi: #172027;
    --soluk: #66717a;
    --cizgi: #dfe5e8;
    --ana: #0f2027;
    --ana-2: #203a43;
    --iyi: #237a57;
    --iyi-bg: #e9f7f0;
    --uyari: #9a6500;
    --uyari-bg: #fff5dc;
    --kotu: #b33a3a;
    --kotu-bg: #fdecec;
    --notr: #64717a;
    --notr-bg: #eef2f4;
    --hazir: #216b78;
    --hazir-bg: #e8f5f7;
    --elle: #6b3fa0;
    --elle-bg: #f2ecfa;
    --golge: 0 12px 34px rgba(15, 32, 39, .08);
    --radius: 16px;
}

* {
    box-sizing: border-box;
}

html {
    color-scheme: light;
}

body {
    position: relative;
    min-height: 100vh;
    margin: 0;
    background: linear-gradient(180deg, #f8fafb 0%, var(--zemin) 100%);
    color: var(--yazi);
    font-family: Arial, Helvetica, sans-serif;
}

button,
input,
select,
textarea {
    font: inherit;
}

button,
a {
    -webkit-tap-highlight-color: transparent;
}

a {
    color: inherit;
}

[hidden] {
    display: none !important;
}

.sr-only {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    white-space: nowrap;
    border: 0;
}

.ust-cubuk,
.sayfa {
    position: relative;
    z-index: 1;
}

/* --------------------------------------------------------------------------
   Header
---------------------------------------------------------------------------- */

.ust-cubuk {
    position: sticky;
    top: 0;
    z-index: 50;
    min-height: 72px;
    padding: 10px clamp(18px, 4vw, 56px);
    display: grid;
    grid-template-columns: minmax(240px, 1fr) auto minmax(220px, 1fr);
    align-items: center;
    gap: 24px;
    background: rgba(15, 32, 39, .97);
    color: #fff;
    box-shadow: 0 8px 30px rgba(0, 0, 0, .13);
}

.marka-link {
    display: inline-flex;
    align-items: center;
    gap: 12px;
    width: fit-content;
    transform: translateX(-35px);
    text-decoration: none;
}

.marka-isaret {
    width: 72px;
    height: 52px;
    padding: 3px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}

.marka-logo {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: contain;
}

.marka-metin {
    display: grid;
    gap: 2px;
}

.marka-metin small {
    color: #b9c7cd;
}

.ana-nav {
    display: flex;
    align-items: center;
    gap: 4px;
}

.ana-nav a {
    padding: 10px 13px;
    border-radius: 10px;
    color: #d8e1e5;
    text-decoration: none;
    font-size: 14px;
    font-weight: 700;
}

.ana-nav a:hover,
.ana-nav a.aktif {
    background: rgba(255, 255, 255, .10);
    color: #fff;
}

.oturum {
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 12px;
}

.oturum-metin {
    display: grid;
    gap: 2px;
    text-align: right;
    font-size: 13px;
}

.oturum-metin small {
    color: #aebdc3;
    font-size: 10px;
    letter-spacing: .7px;
}

/* --------------------------------------------------------------------------
   Genel layout
---------------------------------------------------------------------------- */

.sayfa {
    width: min(1520px, calc(100% - 36px));
    margin: 0 auto;
    padding: 30px 0 60px;
}

.sayfa-shell {
    display: grid;
    gap: 20px;
}

.sayfa-baslik {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 24px;
    margin-bottom: 22px;
}

.sayfa-shell > .sayfa-baslik {
    margin-bottom: 0;
}

.sayfa-hero {
    align-items: center;
}

.sayfa-baslik h1,
.panel-baslik h2,
.operator-kart h2,
.ozet-donem-karti h2 {
    margin: 4px 0 6px;
}

.sayfa-baslik h1 {
    font-size: clamp(28px, 4vw, 42px);
    letter-spacing: -.8px;
}

.sayfa-hero .soluk {
    max-width: 760px;
    margin: 4px 0 0;
    line-height: 1.55;
}

.ust-etiket {
    margin: 0;
    color: #66757e;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.2px;
}

.soluk {
    color: var(--soluk);
}

.baslik-aksiyon,
.buton-grup,
.satir-aksiyon,
.kart-aksiyonlar {
    display: flex;
    align-items: center;
    gap: 9px;
    flex-wrap: wrap;
}

.buton-grup.dikey {
    display: grid;
}

.panel-kutu {
    margin: 0;
    padding: 20px;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-radius: var(--radius);
    box-shadow: var(--golge);
}

.panel-baslik {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
    padding-bottom: 13px;
    margin-bottom: 15px;
    border-bottom: 1px solid #edf1f2;
}

.panel-baslik h2 {
    font-size: 20px;
}

.panel-aciklama {
    margin: 2px 0 0;
    color: var(--soluk);
    font-size: 11px;
    line-height: 1.5;
}

.sayi-rozet {
    min-width: 34px;
    height: 34px;
    padding: 0 8px;
    display: grid;
    place-items: center;
    border-radius: 999px;
    background: var(--notr-bg);
    font-weight: 800;
}

/* --------------------------------------------------------------------------
   Butonlar ve formlar
---------------------------------------------------------------------------- */

.buton {
    min-height: 40px;
    padding: 10px 15px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    border: 1px solid transparent;
    border-radius: 10px;
    cursor: pointer;
    text-decoration: none;
    font-weight: 700;
    transition: transform .12s ease, opacity .12s ease, background .12s ease;
}

.buton:hover {
    transform: translateY(-1px);
}

.buton:disabled {
    cursor: wait;
    opacity: .55;
    transform: none;
}

.buton-ana {
    background: var(--ana);
    color: #fff;
}

.buton-basari {
    background: var(--iyi);
    color: #fff;
}

.buton-uyari {
    background: var(--uyari-bg);
    color: #785000;
    border-color: #e8c777;
}

.buton-tehlike {
    background: var(--kotu-bg);
    color: var(--kotu);
    border-color: #efb7b7;
}

.buton-hayalet {
    background: #fff;
    color: var(--yazi);
    border-color: var(--cizgi);
}

.ust-cubuk .buton-hayalet {
    background: rgba(255, 255, 255, .08);
    color: #fff;
    border-color: rgba(255, 255, 255, .14);
}

.buton-kucuk {
    min-height: 34px;
    padding: 7px 10px;
    font-size: 12px;
}

.tam-genislik {
    width: 100%;
}

.alan,
.arama {
    display: grid;
    gap: 7px;
}

.alan label,
.arama label {
    font-size: 12px;
    font-weight: 700;
}

.alan small,
.dosya-sec small {
    color: var(--soluk);
    font-size: 11px;
}

.alan-checkbox label {
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 600;
    cursor: pointer;
}

.alan-checkbox input[type="checkbox"] {
    width: 17px;
    min-height: 0;
    height: 17px;
    accent-color: var(--elle);
}

input,
select,
textarea {
    width: 100%;
    min-height: 42px;
    padding: 9px 11px;
    border: 1px solid #cfd8dc;
    border-radius: 9px;
    background: #fff;
    color: var(--yazi);
    outline: 0;
}

textarea {
    resize: vertical;
}

input:focus,
select:focus,
textarea:focus,
button:focus-visible,
a:focus-visible {
    outline: none;
    border-color: #5c7885;
    box-shadow: 0 0 0 3px rgba(32, 58, 67, .16);
}

.arac-cubugu {
    display: flex;
    justify-content: space-between;
    align-items: end;
    gap: 18px;
}

.arama {
    min-width: min(360px, 100%);
}

.arama.kucuk {
    min-width: 250px;
}

.filtreler {
    display: flex;
    gap: 7px;
    flex-wrap: wrap;
}

.filtre {
    min-height: 40px;
    padding: 8px 12px;
    display: inline-flex;
    align-items: center;
    border: 1px solid var(--cizgi);
    border-radius: 999px;
    background: #fff;
    color: #59666d;
    cursor: pointer;
    font-weight: 700;
    font-size: 12px;
    transition: transform .14s ease, background .14s ease, border-color .14s ease;
}

.filtre:hover {
    transform: translateY(-1px);
}

.filtre.aktif {
    background: var(--ana);
    color: #fff;
    border-color: var(--ana);
}

.metin-buton,
.operator-temizle {
    padding: 0;
    border: 0;
    background: transparent;
    color: var(--ana-2);
    cursor: pointer;
    font-size: 11px;
    font-weight: 800;
}

.metin-buton:hover,
.operator-temizle:hover {
    text-decoration: underline;
}

/* --------------------------------------------------------------------------
   Durum renkleri ve KPI
---------------------------------------------------------------------------- */

.istatistik-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
}

.istatistik {
    min-height: 108px;
    padding: 18px;
    display: grid;
    gap: 8px;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-top: 4px solid var(--notr);
    border-radius: 14px;
    box-shadow: var(--golge);
    transition: transform .15s ease, box-shadow .15s ease;
}

.istatistik:hover {
    transform: translateY(-2px);
    box-shadow: 0 15px 34px rgba(15, 32, 39, .10);
}

.istatistik span {
    color: var(--soluk);
    font-size: 13px;
    font-weight: 700;
}

.istatistik strong {
    font-size: 32px;
    line-height: 1;
}

.istatistik-grid-3 {
    grid-template-columns: repeat(3, minmax(0, 1fr));
}

.donem-filtre-cubugu {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 12px;
    margin-bottom: 14px;
}

.donem-ozel-alan {
    display: flex;
    align-items: center;
    gap: 8px;
}

.donem-ozel-alan input[type="date"] {
    min-height: 40px;
    width: auto;
}

.donem-metrik-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 10px;
    margin-bottom: 16px;
}

.donem-metrik-kart {
    padding: 12px 14px;
    display: grid;
    gap: 4px;
    border: 1px solid var(--cizgi);
    border-radius: 10px;
    background: #fafbfc;
}


.donem-metrik-grid-4 {
    grid-template-columns: repeat(4, minmax(0, 1fr));
}

.donem-metrik-kart small {
    display: block;
    margin-top: 4px;
    color: var(--soluk);
    font-size: 14px;
}

.donem-veri-uyari {
    margin-top: 10px;
    padding: 8px 11px;
    border: 1px solid #ecd08a;
    border-radius: 9px;
    background: var(--uyari-bg);
    color: #775200;
    font-size: 11px;
    line-height: 1.5;
}

@media (max-width: 820px) {
    .donem-metrik-grid-4 {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }
}



.donem-metrik-kart span {
    color: var(--soluk);
    font-size: 11px;
    font-weight: 700;
}

.donem-metrik-kart strong {
    color: var(--ana);
    font-size: 20px;
}

@media (max-width: 520px) {
    .donem-metrik-grid {
        grid-template-columns: 1fr 1fr;
    }
    .donem-ozel-alan {
        width: 100%;
        flex-wrap: wrap;
    }
}

.kart-plana {
    border-top-color: #788890;
}

.kart-dizgide {
    border-top-color: #d18a00;
}

.kart-hazir {
    border-top-color: var(--hazir);
}

.kart-teslim {
    border-top-color: var(--iyi);
}

.durum-rozet {
    width: fit-content;
    max-width: 100%;
    padding: 5px 9px;
    display: inline-flex;
    align-items: center;
    border-radius: 999px;
    background: var(--notr-bg);
    color: var(--notr);
    font-size: 10px;
    line-height: 1.2;
    font-weight: 800;
    white-space: nowrap;
}

.durum-rozet.iyi {
    background: var(--iyi-bg);
    color: var(--iyi);
}

.durum-rozet.uyari {
    background: var(--uyari-bg);
    color: var(--uyari);
}

.durum-rozet.kotu {
    background: var(--kotu-bg);
    color: var(--kotu);
}

.durum-rozet.elle {
    background: var(--elle-bg);
    color: var(--elle);
}

.satir-alt-rozet {
    margin-top: 5px;
}

/* --------------------------------------------------------------------------
   Kartlar ve listeler
---------------------------------------------------------------------------- */

.pano-grid,
.yonetim-grid {
    display: grid;
    grid-template-columns: 1.3fr .9fr;
    gap: 20px;
}

.kart-listesi,
.mini-liste,
.log-liste {
    display: grid;
    gap: 10px;
}

.is-karti,
.operator-kart {
    background: #fff;
    border: 1px solid var(--cizgi);
    border-left: 5px solid var(--notr);
    border-radius: 13px;
}

.is-karti {
    padding: 15px;
}

.is-karti.iyi,
.operator-kart.iyi {
    border-left-color: var(--iyi);
}

.is-karti.uyari,
.operator-kart.uyari {
    border-left-color: var(--uyari);
}

.is-karti.kotu,
.operator-kart.kotu {
    border-left-color: var(--kotu);
}

.is-karti-ust,
.operator-kart-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 16px;
}

.is-karti-ust > div {
    display: grid;
    gap: 4px;
}

.is-karti-ust span:not(.durum-rozet) {
    color: var(--soluk);
}

.kart-bilgiler,
.kart-ek-bilgi {
    display: flex;
    flex-wrap: wrap;
    gap: 7px 15px;
    margin-top: 12px;
    color: var(--soluk);
    font-size: 12px;
}

.operator-rozetler {
    display: grid;
    gap: 6px;
    justify-items: end;
}

.durum-satiri {
    display: flex;
    flex-wrap: wrap;
    gap: 10px 18px;
    margin: 12px 0 4px;
    color: #44525a;
    font-size: 13px;
}

.ilerleme {
    margin-top: 14px;
}

.ilerleme-ust {
    display: flex;
    justify-content: space-between;
    gap: 10px;
    margin-bottom: 7px;
    font-size: 12px;
}

.ilerleme-ray {
    height: 8px;
    overflow: hidden;
    border-radius: 999px;
    background: #e8edef;
}

.ilerleme-dolgu {
    height: 100%;
    border-radius: inherit;
    background: linear-gradient(90deg, var(--ana-2), #2c5364);
}

.kart-not {
    margin: 12px 0 0;
    padding: 9px 11px;
    border-left: 3px solid #d7e0e4;
    border-radius: 9px;
    background: #f7f9fa;
    color: #4f5b62;
    font-size: 12px;
}

.mini-satir {
    padding: 12px 0;
    display: flex;
    justify-content: space-between;
    gap: 16px;
    border-bottom: 1px solid #edf0f2;
}

.mini-satir:last-child {
    border-bottom: 0;
}

.mini-satir > div:first-child,
.mini-sag {
    display: grid;
    gap: 4px;
}

.mini-satir span,
.mini-satir small {
    color: var(--soluk);
    font-size: 11px;
}

.mini-sag {
    justify-items: end;
    text-align: right;
}

.bos-durum {
    padding: 28px 16px;
    text-align: center;
    color: var(--soluk);
    border: 1px dashed #ccd5da;
    border-radius: 12px;
    background: rgba(255, 255, 255, .72);
}

.tam-satir {
    grid-column: 1 / -1;
}

/* --------------------------------------------------------------------------
   Operator
---------------------------------------------------------------------------- */

.operator-arac-cubugu,
.panel-arac-cubugu {
    padding: 17px 18px;
    background: linear-gradient(135deg, rgba(255,255,255,.99), rgba(247,250,251,.96));
}

.operator-sayfa .arama,
.panel-arama-alani {
    flex: 1 1 340px;
    max-width: 540px;
}

.operator-filtre-alani,
.panel-filtre-alani {
    display: grid;
    gap: 7px;
    justify-items: end;
}

.operator-filtre-meta,
.panel-filtre-alt {
    min-height: 20px;
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 10px;
    color: var(--soluk);
    font-size: 11px;
}

.operator-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 16px;
    align-items: start;
}

.operator-kart {
    padding: 18px;
    box-shadow: var(--golge);
    transition: transform .16s ease, box-shadow .16s ease;
}

.operator-kart:hover {
    transform: translateY(-2px);
    box-shadow: 0 15px 35px rgba(15, 32, 39, .10);
}

.operator-kart-ust {
    padding-bottom: 13px;
    border-bottom: 1px solid #edf1f2;
}

.operator-kart-ust h2 {
    margin-top: 5px;
    margin-bottom: 4px;
    font-size: 19px;
}

.operator-kart-ust p {
    margin: 0;
    color: var(--soluk);
    font-size: 12px;
}

.bilgi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin-top: 16px;
}

.bilgi-grid > div {
    padding: 10px;
    display: grid;
    gap: 4px;
    border: 1px solid #edf1f2;
    border-radius: 10px;
    background: #fafbfc;
}

.bilgi-grid span {
    color: var(--soluk);
    font-size: 10px;
    font-weight: 700;
}

.bilgi-grid strong {
    font-size: 14px;
}

.kart-aksiyonlar {
    margin-top: 15px;
    padding-top: 14px;
    border-top: 1px solid #edf1f2;
}

.kart-aksiyonlar .buton {
    min-height: 44px;
    flex: 1 1 120px;
    padding: 12px 16px;
}

/* --------------------------------------------------------------------------
   Panel
---------------------------------------------------------------------------- */

.panel-canli-satir {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px 12px;
    margin-top: 5px;
}

.panel-canli {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    color: #526169;
    font-size: 12px;
    font-weight: 700;
}

.panel-canli::before {
    content: "";
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--iyi);
    box-shadow: 0 0 0 4px rgba(35, 122, 87, .11);
}

.panel-guncelleme {
    color: var(--soluk);
    font-size: 12px;
}

.panel-aktif-grid {
    align-items: start;
}

.panel-is-karti {
    transition: transform .14s ease, box-shadow .14s ease;
}

.panel-is-karti:hover {
    transform: translateY(-1px);
    box-shadow: 0 9px 22px rgba(15, 32, 39, .07);
}

.panel-hazir-kutu {
    background: linear-gradient(180deg, #fff, #fbfefe);
}

.panel-mini-satir {
    margin: 0 -4px;
    padding: 12px 6px;
    border-radius: 9px;
}

.panel-mini-satir:hover {
    background: #f7f9fa;
}

.plan-liste {
    display: grid;
}

.plan-satir {
    display: grid;
    grid-template-columns: 44px minmax(180px, 1.1fr) minmax(160px, 1fr) 100px 120px minmax(150px, auto);
    align-items: center;
    gap: 14px;
    padding: 12px 4px;
    border-bottom: 1px solid #edf0f2;
    font-size: 12px;
}

.plan-satir:last-child {
    border-bottom: 0;
}

.plan-satir > div:nth-child(2) {
    display: grid;
    gap: 3px;
}

.plan-satir > div:nth-child(2) span,
.plan-satir > span:not(.durum-rozet) {
    color: var(--soluk);
}

.plan-sira {
    width: 32px;
    height: 32px;
    display: grid;
    place-items: center;
    border-radius: 9px;
    background: var(--notr-bg);
    font-weight: 800;
}

/* --------------------------------------------------------------------------
   Monitör
---------------------------------------------------------------------------- */
.monitor-body {
    height: 100vh;
    overflow: hidden;
}

.monitor-body .sayfa {
    width: min(1900px, calc(100% - 28px));
    height: calc(100vh - 72px);
    min-height: 0;
    padding-top: 12px;
    padding-bottom: 12px;
    overflow: hidden;
}


.monitor-sayfa {
    height: 100%;
    min-height: 0;
    display: grid;
    grid-template-rows: auto minmax(0, 1fr);
    gap: 10px;
}

.monitor-ust {
    min-height: 58px;
    padding: 0 3px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 24px;
}

.monitor-ust h1 {
    margin: 2px 0;
    font-size: clamp(26px, 2.2vw, 36px);
}
.monitor-aciklama {
    margin: 0;
    color: var(--soluk);
    font-size: 12px;
}

.monitor-zaman {
    display: grid;
    justify-items: end;
    gap: 2px;
    text-align: right;
}

.monitor-zaman span {
    color: var(--ana);
    font-size: clamp(30px, 2.7vw, 42px);
    line-height: 1;
    font-weight: 800;
    letter-spacing: -1px;
}

.monitor-zaman small {
    color: var(--soluk);
    font-size: 10px;
}

.monitor-grid {
    min-height: 0;
    height: 100%;
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1.15fr);
    gap: 14px;
    align-items: stretch;
}

.monitor-bolum {
    min-width: 0;
    min-height: 0;
    height: 100%;
    padding: 12px;
    display: grid;
    grid-template-rows: auto minmax(0, 1fr);
    align-content: stretch;
    overflow: hidden;
    border-top-width: 4px;
    box-shadow: 0 8px 22px rgba(15, 32, 39, .06);
}

.monitor-dizgide {
    border-top-color: #d18a00;
}

.monitor-hazir {
    border-top-color: var(--hazir);
}

.monitor-plana {
    border-top-color: #788890;
}

.monitor-bolum-baslik {
    margin-bottom: 9px;
    padding-bottom: 10px;
}

.monitor-bolum-baslik h2 {
    margin: 2px 0 0;
    font-size: clamp(20px, 1.5vw, 27px);
    letter-spacing: -.4px;
}

.monitor-bolum-baslik .sayi-rozet {
    min-width: 34px;
    height: 34px;
    font-size: 14px;
}

/*
   İç scroll kaldırıldı.
   Kartlar section içinde doğal olarak yerleşir.
*/
.monitor-liste {
    min-height: 0;
    height: 100%;
    padding: 0;
    display: grid;
    align-content: start;
    gap: 8px;
    overflow: hidden;
}

.monitor-dizgide-liste {
    min-height: 0;
    height: 100%;
    display: grid;
    gap: 8px;
    align-content: start;
}

/* Tam 3 kart varsa mevcut alanı üç eşit parçaya böl */
.monitor-dizgide-liste.uc-dizgi {
    grid-template-rows: repeat(3, minmax(0, 1fr));
    align-content: stretch;
}

.monitor-dizgide-liste.uc-dizgi .monitor-kart {
    min-height: 0;
    height: 100%;
    padding: 10px 13px;
    display: grid;
    align-content: space-between;
    overflow: hidden;
}

.monitor-dizgide-liste.uc-dizgi .monitor-kart-ust {
    gap: 6px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-talep strong {
    font-size: clamp(19px, 1.25vw, 24px);
}

.monitor-dizgide-liste.uc-dizgi .monitor-kart h3 {
    margin: 6px 0 2px;
    font-size: clamp(17px, 1.1vw, 22px);
}

.monitor-dizgide-liste.uc-dizgi .monitor-sahip {
    font-size: 11px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-adet {
    margin-top: 5px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-adet strong {
    font-size: clamp(19px, 1.25vw, 24px);
}

.monitor-dizgide-liste.uc-dizgi .monitor-adet span {
    font-size: 9px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-ilerleme {
    margin-top: 5px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-ilerleme .ilerleme-ust {
    margin-bottom: 3px;
    font-size: 9px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-ilerleme .ilerleme-ray {
    height: 6px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-meta {
    margin-top: 5px;
    padding-top: 5px;
}

.monitor-dizgide-liste.uc-dizgi .monitor-meta strong {
    font-size: 12px;
}



/* --------------------------------------------------------------------------
   DİZGİDE / HAZIR
---------------------------------------------------------------------------- */

.monitor-kart {
    padding: 15px 16px;
    background: #fff;
    border: 1px solid #dde4e7;
    border-left: 4px solid #d18a00;
    border-radius: 11px;
}

.monitor-hazir .monitor-kart {
    border-left-color: var(--hazir);
}

.monitor-kart.kotu,
.monitor-plan-kart.kotu {
    border-left-color: var(--kotu);
}

.monitor-kart-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 9px;
}

.monitor-ek-rozetler {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-top: 6px;
}

.monitor-talep {
    display: grid;
    gap: 1px;
}

.monitor-talep span,
.monitor-plan-ust > div span {
    color: var(--soluk);
    font-size: 9px;
    font-weight: 800;
    letter-spacing: .9px;
}

.monitor-talep strong {
    color: var(--ana);
    font-size: clamp(21px, 1.45vw, 28px);
    line-height: 1;
}

.monitor-kart .durum-rozet {
    padding: 4px 7px;
    font-size: 9px;
}

.monitor-kart h3 {
    margin: 11px 0 3px;
    overflow-wrap: anywhere;
    font-size: clamp(19px, 1.35vw, 26px);
    line-height: 1.08;
}

.monitor-sahip {
    margin: 0;
    color: var(--soluk);
    font-size: 12px;
    font-weight: 700;
}

.monitor-adet {
    margin-top: 11px;
    display: flex;
    align-items: baseline;
    gap: 6px;
}

.monitor-adet strong {
    font-size: clamp(22px, 1.5vw, 28px);
}

.monitor-adet span {
    color: var(--soluk);
    font-size: 10px;
    font-weight: 700;
}

.monitor-ilerleme {
    margin-top: 9px;
}

.monitor-ilerleme .ilerleme-ust {
    margin-bottom: 5px;
    font-size: 10px;
}

.monitor-ilerleme .ilerleme-ray {
    height: 8px;
}

.monitor-meta {
    margin-top: 10px;
    padding-top: 9px;
    display: flex;
    gap: 15px;
    border-top: 1px solid #edf1f2;
}

.monitor-meta span,
.monitor-hazir-bilgi span,
.monitor-plan-meta span {
    display: grid;
    gap: 2px;
}

.monitor-meta small,
.monitor-hazir-bilgi small,
.monitor-plan-meta small {
    color: var(--soluk);
    font-size: 8px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .45px;
}

.monitor-meta strong {
    font-size: 14px;
}

/* --------------------------------------------------------------------------
   HAZIR
---------------------------------------------------------------------------- */

.monitor-hazir-mesaj {
    margin-top: 11px;
    padding: 8px 10px;
    border: 1px solid #c9e0e4;
    border-radius: 8px;
    background: var(--hazir-bg);
    color: #175965;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: .25px;
}

.monitor-hazir-bilgi {
    margin-top: 9px;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 7px;
}

.monitor-hazir-bilgi span {
    padding: 8px 9px;
    border: 1px solid #edf1f2;
    border-radius: 8px;
    background: #fafcfc;
}

.monitor-hazir-bilgi strong {
    font-size: 14px;
}

/* --------------------------------------------------------------------------
   PLANA ALINDI
---------------------------------------------------------------------------- */

/*
   Plan kuyruğu özellikle daha yoğun.
   Büyük ekranda iki mini kart yan yana.
*/
.monitor-plan-liste {
    grid-template-columns: repeat(2, minmax(0, 1fr));
    grid-auto-rows: max-content;
    gap: 12px;
}
.monitor-plan-kart {
    min-width: 0;
    padding: 14px 15px;
    display: grid;
    grid-template-columns: 38px minmax(0, 1fr);
    gap: 12px;
    background: #fff;
    border: 1px solid #dfe5e8;
    border-left: 4px solid #788890;
    border-radius: 11px;
}

.monitor-plan-sira {
    width: 36px;
    height: 36px;
    display: grid;
    place-items: center;
    border-radius: 9px;
    background: var(--notr-bg);
    color: #4e5b62;
    font-size: 13px;
    font-weight: 800;
}

.monitor-plan-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 8px;
}

.monitor-plan-ust > div {
    min-width: 0;
    display: grid;
    gap: 2px;
}

.monitor-plan-ust > div span {
    font-size: 10px;
}

.monitor-plan-ust strong {
    color: var(--ana);
    font-size: 19px;
    line-height: 1.05;
}

.monitor-plan-kart .durum-rozet {
    padding: 4px 8px;
    font-size: 9px;
}

.monitor-plan-kart h3 {
    margin: 9px 0 4px;
    overflow-wrap: anywhere;
    font-size: 17px;
    line-height: 1.08;
}

.monitor-plan-kart p {
    margin: 0;
    overflow: hidden;
    color: var(--soluk);
    font-size: 12px;
    font-weight: 700;
    line-height: 1.25;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.monitor-plan-meta {
    margin-top: 10px;
    padding-top: 9px;
    display: grid;
    grid-template-columns: .6fr 1fr 1fr;
    gap: 8px;
    border-top: 1px solid #edf1f2;
}

.monitor-plan-meta small {
    font-size: 9px;
}

.monitor-plan-meta strong {
    font-size: 12px;
    line-height: 1.2;
}

/* --------------------------------------------------------------------------
   Empty state
---------------------------------------------------------------------------- */

.monitor-bos {
    min-height: 110px;
    display: grid;
    place-items: center;
    align-self: start;
    padding: 18px;
    font-size: 12px;
    line-height: 1.45;
}

/* --------------------------------------------------------------------------
   Monitör responsive
---------------------------------------------------------------------------- */

@media (max-width: 1050px) {
    .monitor-grid {
        grid-template-columns: 1fr;
    }

    .monitor-plan-liste {
        grid-template-columns: repeat(3, minmax(0, 1fr));
    }
}


@media (max-width: 820px) {
    .monitor-grid {
        grid-template-columns: 1fr;
    }

    .monitor-plan-liste {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .monitor-ust {
        align-items: flex-start;
    }

    .monitor-zaman {
        flex-shrink: 0;
    }
}

@media (max-width: 520px) {
    .monitor-body .sayfa {
        width: min(100% - 20px, 1520px);
        padding-top: 14px;
    }

    .monitor-ust {
        display: grid;
        gap: 10px;
    }

    .monitor-zaman {
        justify-items: start;
        text-align: left;
    }

    .monitor-zaman span {
        font-size: 30px;
    }

    .monitor-bolum {
        padding: 12px;
    }

    .monitor-plan-liste {
        grid-template-columns: 1fr;
    }

    .monitor-hazir-bilgi {
        grid-template-columns: 1fr 1fr;
    }
}

/* --------------------------------------------------------------------------
   Yönetim
---------------------------------------------------------------------------- */

.yonetim-yukleme-karti {
    background: linear-gradient(145deg, #fff, #f9fbfc);
}

.yukleme-form {
    display: grid;
    gap: 12px;
}

.dosya-sec {
    min-height: 115px;
    padding: 16px;
    display: grid;
    place-content: center;
    gap: 9px;
    border: 2px dashed #aab9c0;
    border-radius: 12px;
    background: #f8fafb;
    cursor: pointer;
    transition: border-color .15s ease, background .15s ease;
}

.dosya-sec:hover {
    border-color: #708890;
    background: #f3f7f8;
}

.dosya-sec span {
    font-size: 14px;
    font-weight: 700;
}

.dosya-sec input {
    padding: 7px;
    background: #fff;
}

.yardim {
    margin-bottom: 0;
    color: var(--soluk);
    font-size: 11px;
    line-height: 1.65;
}

.yardim code {
    padding: 1px 6px;
    border-radius: 6px;
    background: var(--notr-bg);
    font-size: 12px;
}

.yonetim-durum-eksik {
    border-color: #e9ca7c;
    background: linear-gradient(135deg, #fffdf6, #fff9e9);
}

.yonetim-durum-eksik-liste {
    max-height: 320px;
    padding-right: 7px;
    overflow-y: auto;
    overscroll-behavior: contain;
}

.yonetim-durum-eksik-liste .mini-satir:first-child {
    padding-top: 4px;
}

.yonetim-durum-eksik-liste .mini-satir:last-child {
    padding-bottom: 4px;
}

.yonetim-uyari-kutu {
    border-color: #ecd497;
    background: linear-gradient(135deg, #fffdf6, #fffaf0);
}

.yonetim-kart-panel .panel-baslik {
    align-items: end;
}

.yonetim-gecmis .mini-satir,
.yonetim-gecmis .log-satir {
    padding-left: 6px;
    padding-right: 6px;
    border-radius: 8px;
}

.yonetim-gecmis .mini-satir:hover,
.yonetim-gecmis .log-satir:hover {
    background: #f7f9fa;
}

.log-satir {
    padding: 10px 0;
    display: grid;
    gap: 4px;
    border-bottom: 1px solid #edf0f2;
}

.log-satir:last-child {
    border-bottom: 0;
}

.log-satir > div {
    display: flex;
    justify-content: space-between;
    gap: 12px;
}

.log-satir span,
.log-satir small {
    color: var(--soluk);
    font-size: 11px;
}

.admin-filtre-cubugu {
    margin: -2px 0 13px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 14px;
}

.admin-filtre-sonuc {
    flex-shrink: 0;
    color: var(--soluk);
    font-size: 11px;
    font-weight: 700;
}

.yonetim-kart-scroll {
    max-height: 470px;
    overflow: auto;
    overscroll-behavior: contain;
    scrollbar-width: thin;
    scrollbar-color: #b8c3c8 transparent;
}

/*
   tablo-kapsayici overflow oluşturduğu için mevcut sticky th
   burada header'ı sabit tutar.
*/
.yonetim-kart-scroll th {
    top: 0;
    z-index: 2;
}

.yonetim-kart-scroll::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

.yonetim-kart-scroll::-webkit-scrollbar-track {
    background: transparent;
}

.yonetim-kart-scroll::-webkit-scrollbar-thumb {
    border-radius: 999px;
    background: #b8c3c8;
}

.yonetim-kart-scroll::-webkit-scrollbar-thumb:hover {
    background: #929fa5;
}

.admin-kart-bos {
    margin-top: 12px;
}

/* --------------------------------------------------------------------------
   Özet ve grafik
---------------------------------------------------------------------------- */

.ozet-uyari-satiri {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}

.donem-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
}

.ozet-donem-karti {
    transition: transform .15s ease, box-shadow .15s ease;
}

.ozet-donem-karti:hover {
    transform: translateY(-2px);
}

.ozet-donem-karti h2 {
    margin-bottom: 17px;
}

.donem-metrik {
    min-height: 42px;
    padding: 10px 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    border-top: 1px solid #edf0f2;
    font-size: 12px;
}

.donem-metrik span {
    color: var(--soluk);
}

.donem-metrik strong {
    color: var(--ana);
    font-size: 14px;
}

.grafik-aciklama {
    display: flex;
    gap: 18px;
    margin-bottom: 14px;
    color: var(--soluk);
    font-size: 11px;
}

.grafik-aciklama span {
    display: flex;
    align-items: center;
    gap: 6px;
}

.lejant {
    width: 10px;
    height: 10px;
    display: inline-block;
    border-radius: 3px;
}

.lejant.plan,
.sutun.plan {
    background: #aab7bd;
}

.lejant.teslim,
.sutun.teslim {
    background: #264653;
}

.hafta-grafik {
    min-height: 310px;
    padding: 20px 8px 0;
    display: grid;
    grid-template-columns: repeat(8, minmax(62px, 1fr));
    align-items: end;
    gap: 12px;
    overflow-x: auto;
}

.hafta-sutun {
    height: 270px;
    display: grid;
    grid-template-rows: 1fr auto auto;
    gap: 5px;
    text-align: center;
    font-size: 11px;
}

.sutun-alani {
    min-height: 220px;
    display: flex;
    align-items: end;
    justify-content: center;
    gap: 4px;
    border-bottom: 1px solid #ccd5da;
}

.sutun {
    position: relative;
    width: min(28px, 42%);
    min-height: 2px;
    border-radius: 6px 6px 0 0;
}

.sutun span {
    position: absolute;
    top: -18px;
    left: 50%;
    transform: translateX(-50%);
    font-size: 10px;
    font-weight: 700;
}

.hafta-sutun small {
    color: var(--soluk);
}

/* --------------------------------------------------------------------------
   Tablolar
---------------------------------------------------------------------------- */

.tablo-kapsayici {
    max-width: 100%;
    overflow: auto;
    border: 1px solid #edf1f2;
    border-radius: 11px;
}

table {
    width: 100%;
    min-width: 760px;
    border-collapse: collapse;
}

th,
td {
    padding: 11px 12px;
    text-align: left;
    vertical-align: middle;
    border-bottom: 1px solid #e9edef;
    font-size: 12px;
}

th {
    position: sticky;
    top: 0;
    z-index: 1;
    background: #f6f8f9;
    color: #59666d;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: .5px;
}

tbody tr:hover {
    background: #fafbfb;
}

/* --------------------------------------------------------------------------
   Modal
---------------------------------------------------------------------------- */

.modal {
    width: min(640px, calc(100% - 24px));
    max-width: none;
    max-height: calc(100vh - 32px);
    padding: 0;
    overflow: visible;
    border: 0;
    background: transparent;
}

.modal[open] {
    display: block;
}

.modal::backdrop {
    background: rgba(9, 19, 24, .62);
    backdrop-filter: blur(2px);
}

.modal-kutu {
    width: 100%;
    max-height: calc(100vh - 40px);
    padding: 20px;
    overflow-y: auto;
    border-radius: 16px;
    background: #fff;
    box-shadow: 0 28px 70px rgba(0, 0, 0, .24);
}

.modal-baslik {
    display: flex;
    justify-content: space-between;
    gap: 15px;
    margin-bottom: 18px;
}

.modal-baslik h2 {
    margin: 4px 0 0;
}

.ikon-buton {
    width: 36px;
    height: 36px;
    border: 1px solid var(--cizgi);
    border-radius: 9px;
    background: #fff;
    cursor: pointer;
    font-size: 22px;
}

.modal-aksiyon {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 18px;
}

.iki-kolon {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
}

.modal .alan + .alan,
.modal .iki-kolon + .alan,
.modal .alan + .iki-kolon,
.modal .iki-kolon + .iki-kolon {
    margin-top: 12px;
}

/* --------------------------------------------------------------------------
   Bildirim / toast
---------------------------------------------------------------------------- */

.bildirimler {
    margin-bottom: 16px;
}

.bildirim {
    padding: 11px 13px;
    border: 1px solid transparent;
    border-radius: 10px;
    font-size: 13px;
}

.bildirim.basari {
    background: var(--iyi-bg);
    border-color: #b7dfca;
    color: #175d41;
}

.bildirim.hata {
    background: var(--kotu-bg);
    border-color: #efb9b9;
    color: #922f2f;
}

.bildirim.uyari {
    background: var(--uyari-bg);
    border-color: #ecd08a;
    color: #775200;
}

.toast-alani {
    position: fixed;
    right: 20px;
    bottom: 20px;
    z-index: 200;
    display: grid;
    gap: 8px;
    pointer-events: none;
}

.toast {
    max-width: 380px;
    padding: 12px 14px;
    border-radius: 11px;
    background: var(--ana);
    color: #fff;
    box-shadow: 0 14px 40px rgba(0, 0, 0, .16);
    opacity: 0;
    transform: translateY(8px);
    transition: .18s ease;
    font-size: 13px;
}

.toast.goster {
    opacity: 1;
    transform: translateY(0);
}

.toast.hata {
    background: var(--kotu);
}

.toast.uyari {
    background: #8b5f06;
}

/* --------------------------------------------------------------------------
   Login / 403
---------------------------------------------------------------------------- */

.giris-sayfa {
    position: relative;
    min-height: 100vh;
    padding: 24px;
    display: grid;
    place-items: center;
    overflow: hidden;
    background:
        radial-gradient(circle at 20% 20%, rgba(44, 83, 100, .18), transparent 30%),
        linear-gradient(145deg, #0f2027, #203a43 50%, #2c5364);
}

.giris-sayfa::before {
    content: "";
    position: fixed;
    top: 50%;
    left: 50%;
    width: min(900px, 78vw);
    height: min(900px, 78vw);
    transform: translate(-50%, -50%) rotate(-6deg);
    background: url("/static/pdgm_logo.png") center / contain no-repeat;
    opacity: .025;
    filter: grayscale(100%);
    pointer-events: none;
}

.giris-kutu {
    position: relative;
    z-index: 1;
    width: min(420px, 100%);
    padding: 34px;
    border: 1px solid rgba(255,255,255,.55);
    border-radius: 20px;
    background: rgba(255, 255, 255, .98);
    box-shadow: 0 30px 80px rgba(0, 0, 0, .28);
    backdrop-filter: blur(4px);
}

.giris-logo {
    width: 170px;
    height: 92px;
    margin: 0 auto 16px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.giris-logo img {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: contain;
}

.giris-kutu h1 {
    margin: 0;
    text-align: center;
    letter-spacing: -.5px;
    font-size: 27px;
}

.giris-kutu .alt {
    margin: 7px 0 24px;
    text-align: center;
    color: var(--soluk);
}

.giris-form {
    display: grid;
    gap: 15px;
}

.giris-form input,
.giris-form .buton {
    min-height: 46px;
}

.giris-dipnot {
    margin: 20px 0 0;
    text-align: center;
    color: #89949a;
    font-size: 10px;
}

.yetkisiz-sayfa {
    min-height: calc(100vh - 170px);
    display: grid;
    place-items: center;
}

.durum-sayfasi {
    width: min(580px, 100%);
    padding: clamp(32px, 6vw, 54px);
    text-align: center;
    border: 1px solid var(--cizgi);
    border-radius: 18px;
    background: linear-gradient(145deg, rgba(255,255,255,.99), rgba(248,250,251,.98));
    box-shadow: var(--golge);
}

.durum-ikon {
    margin: 12px 0;
    color: #d7e0e4;
    font-size: clamp(60px, 11vw, 90px);
    line-height: 1;
    font-weight: 900;
}

.durum-sayfasi h1 {
    margin: 5px 0 10px;
    font-size: clamp(24px, 4vw, 32px);
}

.durum-sayfasi p:not(.ust-etiket) {
    max-width: 440px;
    margin: 0 auto 24px;
    color: var(--soluk);
    line-height: 1.6;
}

/* --------------------------------------------------------------------------
   Responsive
---------------------------------------------------------------------------- */

@media (max-width: 1100px) {
    .operator-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .ust-cubuk {
        grid-template-columns: 1fr auto;
    }

    .ana-nav {
        grid-column: 1 / -1;
        order: 3;
        overflow-x: auto;
    }

    .plan-satir {
        grid-template-columns: 44px minmax(180px, 1fr) 110px 120px minmax(150px, auto);
    }

    .plan-satir > span:nth-of-type(1) {
        display: none;
    }

    .monitor-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .monitor-plana {
        grid-column: 1 / -1;
    }


}

@media (max-width: 820px) {
    .sayfa-shell {
        gap: 15px;
    }

    .istatistik-grid {
        grid-template-columns: repeat(2, 1fr);
    }

    .pano-grid,
    .yonetim-grid,
    .donem-grid,
    .operator-grid,
    .monitor-grid {
        grid-template-columns: 1fr;
    }

    .monitor-plana {
        grid-column: auto;
    }

    .monitor-ust {
        align-items: flex-start;
    }

    .monitor-zaman {
        flex-shrink: 0;
    }



    .sayfa-baslik,
    .arac-cubugu {
        align-items: stretch;
        flex-direction: column;
    }

    .baslik-aksiyon {
        width: 100%;
    }

    .baslik-aksiyon .buton {
        flex: 1;
    }

    .operator-filtre-alani,
    .panel-filtre-alani {
        width: 100%;
        justify-items: stretch;
    }

    .operator-filtre-meta,
    .panel-filtre-alt {
        justify-content: space-between;
    }

    .filtreler {
        width: 100%;
    }

    .filtre {
        flex: 1 1 auto;
        justify-content: center;
    }

    .bilgi-grid {
        grid-template-columns: repeat(2, 1fr);
    }

    .oturum-metin {
        display: none;
    }

    .ana-nav {
        width: 100%;
        gap: 2px;
    }

    .ana-nav a {
        flex: 1 1 auto;
        padding: 12px 10px;
        text-align: center;
    }

    .plan-satir {
        grid-template-columns: 38px minmax(160px, 1fr) 100px minmax(140px, auto);
    }

    .plan-satir > span:nth-of-type(1),
    .plan-satir > span:nth-of-type(2) {
        display: none;
    }
    .admin-filtre-cubugu {
        align-items: stretch;
        flex-direction: column;
    }

    .admin-filtre-cubugu .filtreler {
        width: 100%;
    }

    .admin-filtre-cubugu .filtre {
        flex: 1 1 auto;
        justify-content: center;
    }

    .admin-filtre-sonuc {
        text-align: right;
    }
}

@media (max-width: 520px) {
    .sayfa,
    .monitor-body .sayfa {
        width: min(100% - 20px, 1520px);
        padding-top: 18px;
    }

    .ust-cubuk {
        padding: 9px 12px;
        gap: 10px;
    }

    .marka-metin small {
        display: none;
    }

    .marka-isaret {
        width: 44px;
        height: 40px;
    }

    .panel-kutu,
    .operator-arac-cubugu,
    .panel-arac-cubugu {
        padding: 14px;
    }

    .istatistik-grid {
        gap: 8px;
    }

    .istatistik {
        min-height: 94px;
        padding: 14px;
    }

    .istatistik strong {
        font-size: 28px;
    }

    .iki-kolon {
        grid-template-columns: 1fr;
    }

    .modal-kutu {
        padding: 16px;
    }

    .is-karti-ust,
    .operator-kart-ust {
        display: grid;
    }

    .operator-rozetler {
        justify-items: start;
    }

    .durum-rozet {
        white-space: normal;
    }

    .operator-filtre-alani .filtre,
    .panel-filtre-alani .filtre {
        flex: 1 1 calc(50% - 7px);
    }

    .giris-kutu {
        padding: 28px 22px;
    }

    .giris-logo {
        width: 145px;
        height: 80px;
    }

    .plan-satir {
        grid-template-columns: 34px 1fr;
        gap: 10px;
    }

    .plan-satir > span {
        display: none !important;
    }

    .monitor-ust {
        display: grid;
        gap: 14px;
    }

    .monitor-zaman {
        justify-items: start;
        text-align: left;
    }

    .monitor-zaman span {
        font-size: 32px;
    }

    .monitor-bolum {
        padding: 14px;
    }

    .monitor-liste {
        max-height: none;
        padding-right: 0;
        overflow: visible;
    }

    .monitor-kart-ust,
    .monitor-plan-ust {
        display: grid;
    }

    .monitor-hazir-bilgi {
        grid-template-columns: 1fr;
    }

    .monitor-plan-kart {
        grid-template-columns: 36px minmax(0, 1fr);
        gap: 10px;
    }

    .monitor-plan-sira {
        width: 34px;
        height: 34px;
    }

    .monitor-plan-meta {
        grid-template-columns: 1fr 1fr;
    }

    .monitor-plan-meta span:first-child {
        grid-column: 1 / -1;
    }

    .yonetim-durum-eksik-liste {
        max-height: 360px;
    }

    .yonetim-durum-eksik-liste .mini-satir {
        display: grid;
    }

    .yonetim-durum-eksik-liste .mini-sag {
        justify-items: start;
        text-align: left;
    }
    .admin-filtre-cubugu .filtre {
        flex: 1 1 calc(50% - 7px);
    }

    .yonetim-kart-scroll {
        max-height: 430px;
    }
    
    .panel-tablo-scroll {
        max-height: 350px;
    }
}

/* --------------------------------------------------------------------------
   Panel tablo scroll optimizasyonu
---------------------------------------------------------------------------- */

/*
   Plana alınan ve teslim edilen tablolar büyüyünce
   tüm pano uzamasın. Sadece tablo alanı scroll olsun.
*/

.panel-tablo-scroll {
    max-height: 420px;
    overflow-y: auto;
    overflow-x: auto;
    overscroll-behavior: contain;
}


/* Header sabit kalsın */
.panel-tablo-scroll table th {
    position: sticky;
    top: 0;
    z-index: 3;
}


/* scrollbar */
.panel-tablo-scroll::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

.panel-tablo-scroll::-webkit-scrollbar-track {
    background: transparent;
}

.panel-tablo-scroll::-webkit-scrollbar-thumb {
    background: #b8c3c8;
    border-radius: 999px;
}

.panel-tablo-scroll::-webkit-scrollbar-thumb:hover {
    background: #929fa5;
}


/* Firefox */
.panel-tablo-scroll {
    scrollbar-width: thin;
    scrollbar-color: #b8c3c8 transparent;
}
```
````

## `requirements.txt`

```text
Flask==3.0.3
Werkzeug==3.0.6
openpyxl==3.1.5
waitress==3.0.0
python-dotenv==1.0.1
pywin32==308; sys_platform == "win32"
```

## `run_pdgm.bat`

```bat
@echo off
setlocal

rem Betigin bulundugu klasore gec (nereden calistirilirsa calistirilsin dogru calissin)
cd /d "%~dp0"

rem Sanal ortam kontrolu
if not exist ".venv\Scripts\python.exe" (
    echo HATA: .venv sanal ortami bulunamadi.
    echo Once su komutlari calistirin:
    echo   uv venv
    echo   uv pip install -r requirements.txt
    pause
    exit /b 1
)

rem .env kontrolu
if not exist ".env" (
    echo HATA: .env dosyasi bulunamadi.
    echo .env.example dosyasini kopyalayip .env olarak duzenleyin.
    pause
    exit /b 1
)

echo ============================================
echo   PDGM Is Takip Sistemi baslatiliyor...
echo ============================================
echo.

".venv\Scripts\python.exe" app.py

set HATA_KODU=%ERRORLEVEL%

echo.
if %HATA_KODU% NEQ 0 (
    echo Sunucu HATA ile kapandi ^(kod: %HATA_KODU%^). Yukaridaki mesaji kontrol edin.
    pause
) else (
    echo Sunucu normal sekilde kapandi.
)

endlocal
```

## `static/js/monitor.js`

```javascript
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
```

## `static/js/onizleme.js`

```javascript
/* Excel aktarım önizlemesi: "tamamlanan adedi sıfırla" seçimleri ve Excel'in geride
   kaldığı kartlar için durum kararları. Onay formunun çift gönderim koruması ortak.js'te
   (form[data-tek-gonderim]).
   (import_onizleme.html içindeki satır içi betikten CSP için taşındı; davranış aynı.) */
"use strict";

(() => {
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
```

## `static/js/operator.js`

```javascript
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
```

## `static/js/ortak.js`

```javascript
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
        // Yanıt gövdesi hataya eklenir: çağıran taraf "onay gerekli" gibi bayrakları okuyabilir.
        const hata = new Error(data.hata || data.detail || "İşlem tamamlanamadı.");
        hata.veri = data;
        throw hata;
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

// <form data-tek-gonderim>: Excel okuması saniyeler sürer; form ikinci kez gönderilmesin
// (ikinci önizleme ilkini geçersiz kılar, kullanılmış onay reddedilir).
// data-gonderim-metni verilirse düğme beklerken o metni gösterir.
document.addEventListener("submit", (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || !form.hasAttribute("data-tek-gonderim")) return;
    if (form.dataset.gonderildi === "1") {
        event.preventDefault();
        return;
    }
    if (event.defaultPrevented) return;
    form.dataset.gonderildi = "1";
    const buton = event.submitter || form.querySelector("button[type=submit]");
    if (buton) {
        buton.dataset.ilkMetin = buton.textContent;
        buton.disabled = true;
        if (form.dataset.gonderimMetni) buton.textContent = form.dataset.gonderimMetni;
    }
});

// Geri tuşuyla önbellekten dönülen sayfada düğme kilitli kalmasın.
window.addEventListener("pageshow", (event) => {
    if (!event.persisted) return;
    document.querySelectorAll("form[data-tek-gonderim]").forEach((form) => {
        delete form.dataset.gonderildi;
        form.querySelectorAll("button[type=submit]").forEach((buton) => {
            buton.disabled = false;
            if (buton.dataset.ilkMetin) buton.textContent = buton.dataset.ilkMetin;
        });
    });
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
```

## `static/js/panel.js`

```javascript
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

        // Her bölüm sayfa başına 12 kart gösterir; toplam eşleşme ile o an ekranda olan ayrı yazılır.
        const sayfada = gorunen - gosterilen;
        sonuc.textContent = !gorunen
            ? "Sonuç bulunamadı"
            : sayfada
                ? `${gorunen} kart bulundu · ekranda ${gosterilen}, sonraki sayfalarda ${sayfada}`
                : `${gorunen} kart bulundu`;
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
```

## `static/js/yonetim.js`

```javascript
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

    function yeniKartGovdesi() {
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
    }

    // Aynı Talep NO + Kart Stok No ile kart varsa sunucu onay ister (çift sayımı önlemek için);
    // admin ayrı bir iş olduğunu onaylarsa istek tekrar_onayi ile yeniden gönderilir.
    yeniForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        const submit = event.submitter || yeniForm.querySelector("[type=submit]");
        const govde = yeniKartGovdesi();
        const gonderKart = (ek = {}) => pdgmFetch("/api/admin/kart-ekle", {
            method: "POST", body: JSON.stringify({...govde, ...ek}),
        });
        submit.disabled = true;
        try {
            try {
                await gonderKart();
            } catch (hata) {
                if (!hata.veri?.tekrar_onayi_gerekli) throw hata;
                const mevcut = (hata.veri.mevcut || []).map((k) =>
                    `#${k.id} ${k.dizgi_etiket} · ${k.durum} · ${k.toplam_adet} adet (${k.kaynak === "EXCEL" ? "Excel" : "manuel"})`);
                const tamam = await onayIste({
                    etiket: "AYNI TALEP + STOK",
                    baslik: "Bu iş için zaten kart var",
                    mesaj: `${hata.message} Mevcut: ${mevcut.join("; ")}. Ayrı bir iş olduğundan eminseniz devam edin.`,
                    evet: "Yine de oluştur",
                    tehlike: true,
                });
                if (!tamam) {
                    submit.disabled = false;
                    return;
                }
                await gonderKart({tekrar_onayi: true});
            }
            dialogKapat(yeniForm.closest("dialog"));
            yenileVeBildir("Yeni kart PLANA ALINDI durumunda oluşturuldu.");
        } catch (hata) {
            hataMesaji(hata);
            submit.disabled = false;
        }
    });

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
```

## `static/stil.css`

```css
/* ==========================================================================
   PDGM İş Takip — stil dosyası
   Renk, yazı boyutu ve gölge değerleri :root token'larından gelir; yeni
   kurallarda doğrudan hex yerine token kullanın.
   ========================================================================== */

:root {
    /* Zemin ve yüzey */
    --zemin: #f4f6f8;
    --zemin-2: #f7f9fa;
    --zemin-3: #eef2f4;
    --kart: #ffffff;

    /* Yazı */
    --yazi: #172027;
    --yazi-2: #44525a;
    --soluk: #5d6970;

    /* Çizgi */
    --cizgi: #dfe5e8;
    --cizgi-acik: #edf1f2;
    --cizgi-giris: #c5d0d5;

    /* Marka / koyu yüzey */
    --ana: #0f2027;
    --ana-2: #203a43;
    --ana-3: #2c5364;
    --koyu-yazi: #d8e1e5;
    --koyu-soluk: #aebdc3;

    /* Durum renkleri: yazı / zemin / kenar */
    --iyi: #237a57;
    --iyi-bg: #e9f7f0;
    --iyi-cizgi: #b7dfca;
    --uyari: #8a5a00;
    --uyari-bg: #fff5dc;
    --uyari-cizgi: #ecd08a;
    --uyari-koyu: #6b4c00;
    --kotu: #b33a3a;
    --kotu-bg: #fdecec;
    --kotu-cizgi: #efb9b9;
    --kotu-koyu: #922f2f;
    --notr: #56636b;
    --notr-bg: #eef2f4;
    --hazir: #216b78;
    --hazir-bg: #e8f5f7;
    --elle: #6b3fa0;
    --elle-bg: #f2ecfa;
    --eum: #1f5f8b;
    --eum-bg: #e8f2f8;

    /* Aşama vurgu renkleri (yalnız kenar/çubuk; yazı rengi olarak kullanmayın) */
    --dizgide: #d18a00;
    --plana: #6b828c;

    /* Klavye odağı */
    --odak: #1a66d2;
    --odak-koyu-zemin: #9cc2ff;

    /* Yazı boyutları */
    --yazi-xs: 11px;
    --yazi-sm: 12px;
    --yazi-md: 13px;
    --yazi-lg: 14px;

    --golge: 0 12px 34px rgba(15, 32, 39, .08);
    --golge-kuvvetli: 0 15px 34px rgba(15, 32, 39, .12);
    --radius: 16px;
}

* {
    box-sizing: border-box;
}

html {
    color-scheme: light;
}

body {
    position: relative;
    min-height: 100vh;
    margin: 0;
    background: linear-gradient(180deg, #f8fafb 0%, var(--zemin) 100%);
    color: var(--yazi);
    font-family: Arial, Helvetica, sans-serif;
}

button,
input,
select,
textarea {
    font: inherit;
}

button,
a {
    -webkit-tap-highlight-color: transparent;
}

a {
    color: inherit;
}

[hidden] {
    display: none !important;
}

.sr-only {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    white-space: nowrap;
    border: 0;
}

/* --------------------------------------------------------------------------
   Klavye odağı: her zeminde en az 3:1 kontrastlı, belirgin halka.
---------------------------------------------------------------------------- */

:focus-visible {
    outline: 3px solid var(--odak);
    outline-offset: 2px;
}

input:focus-visible,
select:focus-visible,
textarea:focus-visible {
    outline-width: 2px;
    outline-offset: 1px;
    border-color: var(--odak);
}

.ust-cubuk :focus-visible,
.onizleme-aksiyon :focus-visible,
.toast :focus-visible {
    outline-color: var(--odak-koyu-zemin);
}

/* --------------------------------------------------------------------------
   Üst çubuk
---------------------------------------------------------------------------- */

.ust-cubuk {
    position: sticky;
    top: 0;
    z-index: 50;
    min-height: 72px;
    padding: 10px clamp(18px, 4vw, 56px);
    display: grid;
    grid-template-columns: minmax(240px, 1fr) auto minmax(220px, 1fr);
    align-items: center;
    gap: 10px 24px;
    background: rgba(15, 32, 39, .97);
    color: #fff;
    box-shadow: 0 8px 30px rgba(0, 0, 0, .13);
}

.marka-link {
    display: inline-flex;
    align-items: center;
    gap: 12px;
    width: fit-content;
    text-decoration: none;
}

/* Logo kare değil (1079×1046); kutuyu doldurur, taşan kenar kırpılır. */
.marka-logo {
    display: block;
    width: 48px;
    height: 48px;
    flex-shrink: 0;
    object-fit: cover;
    border-radius: 10px;
    box-shadow: 0 0 0 1px rgba(255, 255, 255, .14);
}

.marka-metin {
    display: grid;
    gap: 2px;
}

.ana-nav {
    display: flex;
    align-items: center;
    gap: 4px;
}

.ana-nav a {
    padding: 10px 13px;
    border-radius: 10px;
    color: var(--koyu-yazi);
    text-decoration: none;
    font-size: var(--yazi-lg);
    font-weight: 700;
}

.ana-nav a:hover,
.ana-nav a.aktif {
    background: rgba(255, 255, 255, .10);
    color: #fff;
}

.oturum {
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 12px;
}

.oturum-metin {
    display: grid;
    gap: 2px;
    text-align: right;
    font-size: var(--yazi-md);
}

.oturum-metin small {
    color: var(--koyu-soluk);
    font-size: var(--yazi-xs);
    letter-spacing: .7px;
}

/* Başka ekranda yapılan değişiklik / bağlantı kopması bandı (ortak.js). */
.veri-bandi {
    grid-column: 1 / -1;
    padding: 8px 12px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 8px 14px;
    border-radius: 10px;
    background: var(--uyari-bg);
    color: var(--uyari-koyu);
    font-size: var(--yazi-md);
    font-weight: 700;
}

.veri-bandi[data-tip="hata"] {
    background: var(--kotu-bg);
    color: var(--kotu-koyu);
}

.veri-bandi .buton {
    background: var(--ana);
    color: #fff;
}

/* --------------------------------------------------------------------------
   Genel yerleşim
---------------------------------------------------------------------------- */

.sayfa {
    width: min(1520px, calc(100% - 36px));
    margin: 0 auto;
    padding: 30px 0 60px;
}

.sayfa-shell {
    display: grid;
    grid-template-columns: minmax(0, 1fr);
    gap: 20px;
}

.sayfa-shell > *,
.yonetim-grid > * {
    min-width: 0;
}

.sayfalama {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 12px;
    margin-top: 16px;
    flex-wrap: wrap;
    font-size: var(--yazi-md);
}

.sayfa-baslik {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 24px;
    margin-bottom: 22px;
}

.sayfa-shell > .sayfa-baslik {
    margin-bottom: 0;
}

.sayfa-hero {
    align-items: center;
}

.sayfa-baslik h1,
.panel-baslik h2,
.operator-kart h2 {
    margin: 4px 0 6px;
}

.sayfa-baslik h1 {
    font-size: clamp(28px, 4vw, 42px);
    letter-spacing: -.8px;
}

.sayfa-hero .soluk {
    max-width: 760px;
    margin: 4px 0 0;
    line-height: 1.55;
}

.ust-etiket {
    margin: 0;
    color: var(--soluk);
    font-size: var(--yazi-xs);
    font-weight: 800;
    letter-spacing: 1.2px;
}

.soluk {
    color: var(--soluk);
}

.baslik-aksiyon,
.buton-grup,
.satir-aksiyon,
.kart-aksiyonlar {
    display: flex;
    align-items: center;
    gap: 9px;
    flex-wrap: wrap;
}

.buton-grup.dikey {
    display: grid;
}

.panel-kutu {
    margin: 0;
    padding: 20px;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-radius: var(--radius);
    box-shadow: var(--golge);
}

.panel-baslik {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
    padding-bottom: 13px;
    margin-bottom: 15px;
    border-bottom: 1px solid var(--cizgi-acik);
}

.panel-baslik h2 {
    font-size: 20px;
}

.panel-aciklama {
    margin: 2px 0 0;
    color: var(--soluk);
    font-size: var(--yazi-md);
    line-height: 1.5;
}

.sayi-rozet {
    min-width: 34px;
    height: 34px;
    padding: 0 8px;
    display: grid;
    place-items: center;
    border-radius: 999px;
    background: var(--notr-bg);
    font-weight: 800;
}

/* --------------------------------------------------------------------------
   Butonlar ve formlar
---------------------------------------------------------------------------- */

.buton {
    min-height: 40px;
    padding: 10px 15px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    border: 1px solid transparent;
    border-radius: 10px;
    cursor: pointer;
    text-decoration: none;
    font-weight: 700;
    transition: transform .12s ease, opacity .12s ease, background .12s ease;
}

.buton:hover:not(:disabled) {
    transform: translateY(-1px);
}

.buton:disabled {
    cursor: not-allowed;
    opacity: .55;
}

.buton-ana {
    background: var(--ana);
    color: #fff;
}

.buton-basari {
    background: var(--iyi);
    color: #fff;
}

.buton-uyari {
    background: var(--uyari-bg);
    color: var(--uyari-koyu);
    border-color: var(--uyari-cizgi);
}

.buton-tehlike {
    background: var(--kotu-bg);
    color: var(--kotu);
    border-color: var(--kotu-cizgi);
}

.buton-hayalet {
    background: var(--kart);
    color: var(--yazi);
    border-color: var(--cizgi);
}

.ust-cubuk .buton-hayalet {
    background: rgba(255, 255, 255, .08);
    color: #fff;
    border-color: rgba(255, 255, 255, .14);
}

.buton-kucuk {
    min-height: 34px;
    padding: 7px 10px;
    font-size: var(--yazi-md);
}

.tam-genislik {
    width: 100%;
}

.alan,
.arama {
    display: grid;
    gap: 7px;
}

/* .alan 'display: grid' tanımladığı için [hidden] ile aynı özgüllükte çakışır;
   dizgi tipine göre alan gösterme/gizleme buna dayandığı için açık kural gerekli. */
.alan[hidden] {
    display: none;
}

.alan label,
.alan-etiket,
.arama label {
    font-size: var(--yazi-md);
    font-weight: 700;
}

.alan small,
.dosya-sec small {
    color: var(--soluk);
    font-size: var(--yazi-sm);
}

.alan-checkbox label {
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 600;
    cursor: pointer;
}

.alan-checkbox input[type="checkbox"] {
    width: 17px;
    min-height: 0;
    height: 17px;
    accent-color: var(--elle);
}

input,
select,
textarea {
    width: 100%;
    min-height: 42px;
    padding: 9px 11px;
    border: 1px solid var(--cizgi-giris);
    border-radius: 9px;
    background: var(--kart);
    color: var(--yazi);
}

input[readonly] {
    background: var(--zemin-2);
    color: var(--soluk);
}

textarea {
    resize: vertical;
}

.arac-cubugu {
    display: flex;
    justify-content: space-between;
    align-items: end;
    gap: 18px;
}

.arama {
    min-width: min(360px, 100%);
}

.arama.kucuk {
    min-width: min(250px, 100%);
}

.filtreler {
    display: flex;
    gap: 7px;
    flex-wrap: wrap;
}

.filtre {
    min-height: 40px;
    padding: 8px 12px;
    display: inline-flex;
    align-items: center;
    border: 1px solid var(--cizgi);
    border-radius: 999px;
    background: var(--kart);
    color: var(--yazi-2);
    cursor: pointer;
    font-weight: 700;
    font-size: var(--yazi-md);
}

.filtre:hover {
    border-color: var(--plana);
}

.filtre.aktif {
    background: var(--ana);
    color: #fff;
    border-color: var(--ana);
}

.metin-buton {
    padding: 0;
    border: 0;
    background: transparent;
    color: var(--ana-2);
    cursor: pointer;
    font-size: var(--yazi-md);
    font-weight: 800;
}

.metin-buton:hover {
    text-decoration: underline;
}

/* --------------------------------------------------------------------------
   Durum özet kartları (KPI)
---------------------------------------------------------------------------- */

.istatistik-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
}

.istatistik-grid-3 {
    grid-template-columns: repeat(3, minmax(0, 1fr));
}

.istatistik {
    min-height: 108px;
    padding: 18px;
    display: grid;
    gap: 8px;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-top: 4px solid var(--notr);
    border-radius: 14px;
    box-shadow: var(--golge);
}

.istatistik span {
    color: var(--soluk);
    font-size: var(--yazi-md);
    font-weight: 700;
}

.istatistik strong {
    font-size: 32px;
    line-height: 1;
}

.istatistik small {
    color: var(--soluk);
    font-size: var(--yazi-sm);
}

/* Tıklanınca o durumu filtreleyen KPI butonu. */
.istatistik-buton {
    width: 100%;
    color: inherit;
    font: inherit;
    text-align: left;
    cursor: pointer;
    transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}

.istatistik-buton:hover {
    transform: translateY(-2px);
    box-shadow: var(--golge-kuvvetli);
}

.istatistik-buton > span {
    color: var(--yazi);
    font-size: 16px;
    font-weight: 800;
}

.istatistik-buton > small {
    font-size: var(--yazi-md);
}

.istatistik-buton[aria-pressed="true"] {
    border-color: var(--ana-2);
    box-shadow: 0 0 0 2px var(--ana-2), var(--golge);
}

.kart-plana { border-top-color: var(--plana); }
.kart-dizgide { border-top-color: var(--dizgide); }
.kart-teslim { border-top-color: var(--iyi); }

/* --------------------------------------------------------------------------
   Pano dönem özeti
---------------------------------------------------------------------------- */

.donem-filtre-cubugu {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 12px;
    margin-bottom: 14px;
}

.donem-ozel-alan {
    display: flex;
    align-items: center;
    gap: 8px;
}

.donem-ozel-alan input {
    width: 9.5em;
    min-height: 40px;
}

.donem-metrik-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 10px;
    margin-bottom: 16px;
}

.donem-metrik-grid-4 {
    grid-template-columns: repeat(4, minmax(0, 1fr));
}

.donem-metrik-kart {
    padding: 12px 14px;
    display: grid;
    gap: 4px;
    border: 1px solid var(--cizgi);
    border-radius: 10px;
    background: var(--zemin-2);
}

.donem-metrik-kart span {
    color: var(--soluk);
    font-size: var(--yazi-sm);
    font-weight: 700;
}

.donem-metrik-kart strong {
    color: var(--ana);
    font-size: 20px;
}

.donem-metrik-kart small {
    display: block;
    margin-top: 4px;
    color: var(--soluk);
    font-size: var(--yazi-md);
}

.donem-veri-uyari {
    margin-top: 10px;
    padding: 8px 11px;
    border: 1px solid var(--uyari-cizgi);
    border-radius: 9px;
    background: var(--uyari-bg);
    color: var(--uyari-koyu);
    font-size: var(--yazi-sm);
    line-height: 1.5;
}

[aria-busy="true"] .donem-metrik-grid,
[aria-busy="true"] .panel-tablo-scroll {
    opacity: .6;
    transition: opacity .15s ease;
}

/* --------------------------------------------------------------------------
   Rozetler
---------------------------------------------------------------------------- */

.durum-rozet {
    width: fit-content;
    max-width: 100%;
    padding: 5px 9px;
    display: inline-flex;
    align-items: center;
    border-radius: 999px;
    background: var(--notr-bg);
    color: var(--notr);
    font-size: var(--yazi-xs);
    line-height: 1.2;
    font-weight: 800;
    white-space: nowrap;
}

td .durum-rozet + .durum-rozet {
    margin-left: 4px;
}

.durum-rozet.iyi { background: var(--iyi-bg); color: var(--iyi); }
.durum-rozet.uyari { background: var(--uyari-bg); color: var(--uyari); }
.durum-rozet.kotu { background: var(--kotu-bg); color: var(--kotu); }
.durum-rozet.elle { background: var(--elle-bg); color: var(--elle); }
.durum-rozet.eum { background: var(--eum-bg); color: var(--eum); }

.satir-alt-rozet {
    margin-top: 5px;
    display: flex;
    flex-wrap: wrap;
    gap: 5px;
}

/* --------------------------------------------------------------------------
   Kartlar ve listeler
---------------------------------------------------------------------------- */

.yonetim-grid {
    display: grid;
    grid-template-columns: 1.3fr .9fr;
    gap: 20px;
}

.kart-listesi,
.mini-liste,
.log-liste {
    display: grid;
    gap: 10px;
}

.is-karti,
.operator-kart {
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-left: 5px solid var(--notr);
    border-radius: 13px;
}

.is-karti {
    padding: 15px;
}

.is-karti.iyi,
.operator-kart.iyi { border-left-color: var(--iyi); }
.is-karti.uyari,
.operator-kart.uyari { border-left-color: var(--dizgide); }
.is-karti.kotu,
.operator-kart.kotu { border-left-color: var(--kotu); }

.is-karti-ust,
.operator-kart-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 16px;
}

.is-karti-ust > div {
    display: grid;
    gap: 4px;
}

.is-karti-ust span:not(.durum-rozet) {
    color: var(--soluk);
}

.kart-bilgiler,
.kart-ek-bilgi {
    display: flex;
    flex-wrap: wrap;
    gap: 7px 15px;
    margin-top: 12px;
    color: var(--soluk);
    font-size: var(--yazi-md);
}

.operator-rozetler {
    display: grid;
    gap: 6px;
    justify-items: end;
}

.durum-satiri {
    display: flex;
    flex-wrap: wrap;
    gap: 10px 18px;
    margin: 12px 0 4px;
    color: var(--yazi-2);
    font-size: var(--yazi-md);
}

.ilerleme {
    margin-top: 14px;
}

.ilerleme-ust {
    display: flex;
    justify-content: space-between;
    gap: 10px;
    margin-bottom: 7px;
    font-size: var(--yazi-md);
}

/* <progress> ilerleme çubuğu (satır içi style="width" yerine; CSP uyumlu). */
.ilerleme-cubugu {
    display: block;
    width: 100%;
    height: 8px;
    border: 0;
    border-radius: 999px;
    overflow: hidden;
    background: #e3e9ec;
    color: var(--ana-2);
    -webkit-appearance: none;
    appearance: none;
}

.ilerleme-cubugu::-webkit-progress-bar {
    border-radius: 999px;
    background: #e3e9ec;
}

.ilerleme-cubugu::-webkit-progress-value {
    border-radius: 999px;
    background: linear-gradient(90deg, var(--ana-2), var(--ana-3));
}

.ilerleme-cubugu::-moz-progress-bar {
    border-radius: 999px;
    background: linear-gradient(90deg, var(--ana-2), var(--ana-3));
}

.kart-not {
    margin: 12px 0 0;
    padding: 9px 11px;
    border-left: 3px solid var(--cizgi);
    border-radius: 9px;
    background: var(--zemin-2);
    color: var(--yazi-2);
    font-size: var(--yazi-md);
    white-space: pre-wrap;
    word-break: break-word;
}

.not-gecmis {
    margin: 0;
    padding: 10px 12px;
    max-height: 160px;
    overflow: auto;
    border: 1px solid var(--cizgi-acik);
    border-radius: 10px;
    background: var(--zemin-2);
    color: var(--yazi-2);
    font: 12px/1.45 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    white-space: pre-wrap;
    word-break: break-word;
}

.mini-satir {
    padding: 12px 0;
    display: flex;
    justify-content: space-between;
    gap: 16px;
    border-bottom: 1px solid var(--cizgi-acik);
}

.mini-satir:last-child {
    border-bottom: 0;
}

.mini-satir > div:first-child,
.mini-sag {
    display: grid;
    gap: 4px;
}

/* Boşluksuz uzun dosya/talep adları dar ekranda sayfayı genişletmesin. */
.mini-satir > div:first-child {
    min-width: 0;
    overflow-wrap: anywhere;
}

.mini-satir span,
.mini-satir small {
    color: var(--soluk);
    font-size: var(--yazi-sm);
}

.mini-sag {
    flex-shrink: 0;
    justify-items: end;
    text-align: right;
}

.bos-durum {
    margin: 0;
    padding: 24px 16px;
    text-align: center;
    color: var(--soluk);
    border: 1px dashed var(--cizgi-giris);
    border-radius: 12px;
    background: rgba(255, 255, 255, .72);
}

.tam-satir {
    grid-column: 1 / -1;
}

/* --------------------------------------------------------------------------
   Operatör
---------------------------------------------------------------------------- */

.operator-arac-cubugu,
.panel-arac-cubugu {
    padding: 17px 18px;
    background: linear-gradient(135deg, rgba(255, 255, 255, .99), rgba(247, 250, 251, .96));
}

.operator-sayfa .arama,
.panel-arama-alani {
    flex: 1 1 340px;
    max-width: 540px;
}

.operator-filtre-alani,
.panel-filtre-alani {
    display: grid;
    gap: 7px;
    justify-items: end;
}

.operator-filtre-meta,
.panel-filtre-alt {
    min-height: 20px;
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 10px;
    color: var(--soluk);
    font-size: var(--yazi-sm);
}

.operator-teslim-notu {
    margin: 0;
    color: var(--soluk);
    font-size: var(--yazi-sm);
    text-align: right;
}

.operator-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 16px;
    align-items: start;
}

.operator-kart {
    padding: 18px;
    box-shadow: var(--golge);
}

.operator-kart-ust {
    padding-bottom: 13px;
    border-bottom: 1px solid var(--cizgi-acik);
}

.operator-kart-ust h2 {
    margin-top: 5px;
    margin-bottom: 4px;
    font-size: 19px;
}

.operator-kart-ust p {
    margin: 0;
    color: var(--soluk);
    font-size: var(--yazi-md);
}

.bilgi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin-top: 16px;
}

.bilgi-grid > div {
    padding: 10px;
    display: grid;
    gap: 4px;
    border: 1px solid var(--cizgi-acik);
    border-radius: 10px;
    background: var(--zemin-2);
}

.bilgi-grid span {
    color: var(--soluk);
    font-size: var(--yazi-xs);
    font-weight: 700;
}

.bilgi-grid strong {
    font-size: var(--yazi-lg);
}

.kart-aksiyonlar {
    margin-top: 15px;
    padding-top: 14px;
    border-top: 1px solid var(--cizgi-acik);
}

.kart-aksiyonlar .buton {
    min-height: 44px;
    flex: 1 1 120px;
    padding: 12px 16px;
}

/* --------------------------------------------------------------------------
   Pano
---------------------------------------------------------------------------- */

.panel-canli-satir {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px 12px;
    margin-top: 5px;
}

.panel-canli {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    color: var(--yazi-2);
    font-size: var(--yazi-md);
    font-weight: 700;
}

.panel-canli::before {
    content: "";
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--iyi);
    box-shadow: 0 0 0 4px rgba(35, 122, 87, .11);
}

.panel-canli[data-durum="degisti"] { color: var(--uyari); }
.panel-canli[data-durum="degisti"]::before { background: var(--dizgide); box-shadow: 0 0 0 4px rgba(209, 138, 0, .15); }
.panel-canli[data-durum="baglanti-yok"],
.panel-canli[data-durum="oturum"] { color: var(--kotu); }
.panel-canli[data-durum="baglanti-yok"]::before,
.panel-canli[data-durum="oturum"]::before { background: var(--kotu); box-shadow: 0 0 0 4px rgba(179, 58, 58, .13); }

.panel-guncelleme {
    color: var(--soluk);
    font-size: var(--yazi-md);
}

.panel-bolum {
    scroll-margin-top: 90px;
}

/* --------------------------------------------------------------------------
   Monitör (atölye TV'si, kiosk)
   Yazı boyutları kartın kendi boyutuna göre ölçeklenir (container query
   birimleri); desteklemeyen tarayıcı önceki satırdaki vw değerini kullanır.
---------------------------------------------------------------------------- */

.monitor-body {
    height: 100vh;
    overflow: hidden;
    background: var(--zemin-3);
}

.monitor-body .sayfa {
    width: min(1900px, calc(100% - 28px));
    height: 100vh;
    min-height: 0;
    padding: 14px 0;
    overflow: hidden;
}

.monitor-sayfa {
    height: 100%;
    min-height: 0;
    display: flex;
    flex-direction: column;
    gap: 12px;
}

.monitor-ust {
    flex: 0 0 auto;
    padding: 0 4px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 24px;
}

.monitor-ust h1 {
    margin: 2px 0;
    font-size: clamp(26px, 2.2vw, 44px);
}

.monitor-ust .ust-etiket {
    font-size: clamp(11px, .7vw, 14px);
}

.monitor-link {
    color: var(--soluk);
}

.monitor-aciklama {
    margin: 0;
    color: var(--soluk);
    font-size: clamp(13px, .9vw, 18px);
}

.monitor-zaman {
    display: grid;
    justify-items: end;
    gap: 4px;
    text-align: right;
}

.monitor-zaman span {
    color: var(--ana);
    font-size: clamp(32px, 3vw, 64px);
    line-height: 1;
    font-weight: 800;
    letter-spacing: -1px;
    font-variant-numeric: tabular-nums;
}

.monitor-zaman small {
    color: var(--soluk);
    font-size: clamp(12px, .8vw, 16px);
}

.monitor-baglanti {
    flex: 0 0 auto;
    padding: 12px 16px;
    border-radius: 12px;
    background: var(--kotu);
    color: #fff;
    font-size: clamp(16px, 1.2vw, 24px);
    font-weight: 800;
}

.monitor-grid {
    flex: 1 1 auto;
    min-height: 0;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
}

.monitor-bolum {
    min-width: 0;
    min-height: 0;
    padding: 14px;
    display: grid;
    grid-template-rows: auto minmax(0, 1fr);
    overflow: hidden;
    border-radius: 16px;
    border: 1px solid var(--cizgi);
    background: var(--kart);
    box-shadow: 0 6px 18px rgba(15, 32, 39, .05);
}

.monitor-dizgide {
    border-top: 5px solid var(--dizgide);
    background: linear-gradient(180deg, #fffaf0 0%, #ffffff 22%);
}

.monitor-plana {
    border-top: 5px solid var(--plana);
    background: linear-gradient(180deg, #f5f8fa 0%, #ffffff 22%);
}

.monitor-bolum-baslik {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    margin-bottom: 12px;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--cizgi-acik);
}

.monitor-bolum-baslik .ust-etiket {
    font-size: clamp(11px, .7vw, 14px);
}

.monitor-bolum-baslik h2 {
    margin: 2px 0 0;
    font-size: clamp(22px, 1.8vw, 38px);
    letter-spacing: -.4px;
}

.monitor-dizgide .monitor-bolum-baslik h2 { color: var(--uyari); }
.monitor-plana .monitor-bolum-baslik h2 { color: var(--yazi-2); }

.monitor-bolum-sag {
    display: flex;
    align-items: center;
    gap: 14px;
}

.monitor-bolum-baslik .sayi-rozet {
    min-width: clamp(36px, 2.6vw, 56px);
    height: clamp(36px, 2.6vw, 56px);
    font-size: clamp(16px, 1.3vw, 26px);
}

.monitor-sayfa-bilgi {
    display: grid;
    justify-items: end;
    gap: 5px;
    color: var(--soluk);
    font-size: clamp(13px, .95vw, 20px);
    font-weight: 800;
}

.monitor-sayfa-cubugu {
    width: clamp(80px, 7vw, 150px);
    height: 5px;
    overflow: hidden;
    border-radius: 999px;
    background: var(--cizgi);
}

.monitor-sayfa-cubugu i {
    display: block;
    width: 0;
    height: 100%;
    background: var(--ana-2);
}

.monitor-sayfa-cubugu i.oynat {
    animation: monitor-sayfa-ilerle linear forwards;
}

@keyframes monitor-sayfa-ilerle {
    from { width: 0; }
    to { width: 100%; }
}

.monitor-liste {
    min-height: 0;
    height: 100%;
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    grid-template-rows: repeat(3, minmax(0, 1fr));
    gap: 12px;
    overflow: hidden;
}

.monitor-kart {
    container-type: size;
    min-width: 0;
    min-height: 0;
    height: 100%;
    padding: 12px 14px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    gap: 8px;
    overflow: hidden;
    background: var(--kart);
    border: 1px solid var(--cizgi);
    border-left: 5px solid var(--dizgide);
    border-radius: 12px;
}

.monitor-kart.monitor-gizli {
    display: none;
}

.monitor-kart-plan { border-left-color: var(--plana); }
.monitor-kart.kotu { border-left-color: var(--kotu); }
.monitor-kart.iyi { border-left-color: var(--iyi); }
.monitor-kart.uyari { border-left-color: var(--dizgide); }

.monitor-kart-govde {
    min-height: 0;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    gap: 4px;
}

.monitor-kart-ust {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 8px;
    flex-shrink: 0;
}

.monitor-talep {
    min-width: 0;
    display: grid;
    gap: 1px;
}

.monitor-talep span {
    color: var(--soluk);
    font-size: clamp(11px, .7vw, 16px);
    font-size: clamp(11px, min(2.8cqi, 4.4cqh), 18px);
    font-weight: 800;
    letter-spacing: .8px;
}

.monitor-talep strong {
    color: var(--ana);
    font-size: clamp(18px, 1.5vw, 32px);
    font-size: clamp(18px, min(7cqi, 10.5cqh), 56px);
    line-height: 1.05;
}

.monitor-kart .durum-rozet {
    max-width: 50%;
    padding: .35em .7em;
    font-size: clamp(11px, .75vw, 16px);
    font-size: clamp(11px, min(3cqi, 4.8cqh), 22px);
    white-space: normal;
    text-align: right;
}

.monitor-ek-rozetler {
    display: flex;
    flex-wrap: wrap;
    gap: 5px;
    flex-shrink: 0;
}

.monitor-ek-rozetler .durum-rozet {
    max-width: 100%;
}

.monitor-kart h3 {
    margin: 0;
    color: var(--ana);
    overflow-wrap: anywhere;
    font-size: clamp(16px, 1.3vw, 28px);
    font-size: clamp(16px, min(5.8cqi, 8.5cqh), 48px);
    line-height: 1.1;
    flex-shrink: 0;
}

.monitor-sahip {
    margin: 0;
    overflow: hidden;
    color: var(--yazi-2);
    font-size: clamp(12px, .9vw, 20px);
    font-size: clamp(12px, min(3.8cqi, 5.8cqh), 28px);
    font-weight: 700;
    text-overflow: ellipsis;
    white-space: nowrap;
    flex-shrink: 0;
}

.monitor-adet-satir {
    display: flex;
    align-items: baseline;
    gap: 6px;
    flex-shrink: 0;
}

.monitor-adet-satir strong {
    color: var(--ana);
    font-size: clamp(16px, 1.3vw, 28px);
    font-size: clamp(16px, min(5.8cqi, 8.5cqh), 48px);
}

.monitor-adet-satir span {
    color: var(--soluk);
    font-size: clamp(11px, .75vw, 16px);
    font-size: clamp(11px, min(3cqi, 4.6cqh), 22px);
    font-weight: 700;
}

.monitor-ilerleme {
    flex-shrink: 0;
    height: 8px;
    height: clamp(6px, 2.4cqh, 14px);
}

.monitor-tarihler {
    flex: 0 0 auto;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    padding: 8px 10px;
    border-radius: 10px;
    border: 1px solid var(--cizgi);
    background: var(--zemin-2);
}

.monitor-tarih-etiket {
    color: var(--soluk);
    font-size: clamp(10px, .65vw, 14px);
    font-size: clamp(10px, min(2.6cqi, 4cqh), 17px);
    font-weight: 700;
    letter-spacing: .35px;
    text-transform: uppercase;
    line-height: 1.2;
}

.monitor-tarih-deger {
    color: var(--ana);
    font-size: clamp(14px, 1.1vw, 24px);
    font-size: clamp(14px, min(4.8cqi, 7cqh), 36px);
    font-weight: 800;
    line-height: 1.25;
    white-space: nowrap;
    font-variant-numeric: tabular-nums;
}

.monitor-bos {
    min-height: 0;
    height: 100%;
    display: grid;
    place-items: center;
    padding: 18px;
    color: var(--soluk);
    font-size: clamp(14px, 1.1vw, 22px);
}

/* --------------------------------------------------------------------------
   Yönetim
---------------------------------------------------------------------------- */

.yonetim-yukleme-karti {
    background: linear-gradient(145deg, #fff, #f9fbfc);
}

.yukleme-form {
    display: grid;
    gap: 12px;
}

.dosya-sec {
    min-height: 115px;
    padding: 16px;
    display: grid;
    place-content: center;
    gap: 9px;
    border: 2px dashed #aab9c0;
    border-radius: 12px;
    background: var(--zemin-2);
    cursor: pointer;
    transition: border-color .15s ease, background .15s ease;
}

.dosya-sec:hover {
    border-color: var(--plana);
    background: #f3f7f8;
}

.dosya-sec span {
    font-size: var(--yazi-lg);
    font-weight: 700;
}

.dosya-sec input {
    padding: 7px;
    background: var(--kart);
}

.yardim {
    margin-bottom: 0;
    color: var(--soluk);
    font-size: var(--yazi-md);
    line-height: 1.65;
}

.yardim p {
    margin: 0;
}

.yardim code {
    padding: 1px 6px;
    border-radius: 6px;
    background: var(--notr-bg);
    font-size: var(--yazi-sm);
}

.yonetim-durum-eksik {
    border-color: var(--uyari-cizgi);
    background: linear-gradient(135deg, #fffdf6, #fff9e9);
}

.yonetim-durum-eksik-liste {
    max-height: 320px;
    padding-right: 7px;
    overflow-y: auto;
    overscroll-behavior: contain;
}

.yonetim-durum-eksik-liste .mini-satir:first-child {
    padding-top: 4px;
}

.yonetim-durum-eksik-liste .mini-satir:last-child {
    padding-bottom: 4px;
}

.yonetim-uyari-kutu {
    border-color: var(--uyari-cizgi);
    background: linear-gradient(135deg, #fffdf6, #fffaf0);
}

.yonetim-kart-panel .panel-baslik {
    align-items: end;
}

.yonetim-gecmis .mini-satir,
.yonetim-gecmis .log-satir {
    padding-left: 6px;
    padding-right: 6px;
    border-radius: 8px;
}

.yonetim-gecmis .mini-satir:hover,
.yonetim-gecmis .log-satir:hover {
    background: var(--zemin-2);
}

.log-satir {
    padding: 10px 0;
    display: grid;
    gap: 4px;
    border-bottom: 1px solid var(--cizgi-acik);
    overflow-wrap: anywhere;
}

.log-satir:last-child {
    border-bottom: 0;
}

.log-satir > div {
    display: flex;
    justify-content: space-between;
    gap: 12px;
}

.log-satir span,
.log-satir small {
    color: var(--soluk);
    font-size: var(--yazi-sm);
}

.admin-filtre-cubugu {
    margin: -2px 0 13px;
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    gap: 14px;
}

.admin-filtre-gruplari {
    display: grid;
    gap: 7px;
}

.admin-filtre-meta {
    display: grid;
    justify-items: end;
    gap: 4px;
    flex-shrink: 0;
}

.admin-filtre-sonuc {
    color: var(--soluk);
    font-size: var(--yazi-sm);
    font-weight: 700;
}

.yonetim-kart-scroll {
    max-height: 470px;
}

.admin-kart-bos {
    margin-top: 12px;
}

/* --------------------------------------------------------------------------
   Tablolar
---------------------------------------------------------------------------- */

.tablo-kapsayici {
    max-width: 100%;
    overflow: auto;
    border: 1px solid var(--cizgi-acik);
    border-radius: 11px;
}

/* Kendi içinde kayan tablolar: sayfa uzamasın, başlık sabit kalsın. */
.panel-tablo-scroll {
    max-height: 420px;
}

.panel-tablo-scroll,
.yonetim-kart-scroll {
    overscroll-behavior: contain;
    scrollbar-width: thin;
    scrollbar-color: #b8c3c8 transparent;
}

.panel-tablo-scroll::-webkit-scrollbar,
.yonetim-kart-scroll::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

.panel-tablo-scroll::-webkit-scrollbar-track,
.yonetim-kart-scroll::-webkit-scrollbar-track {
    background: transparent;
}

.panel-tablo-scroll::-webkit-scrollbar-thumb,
.yonetim-kart-scroll::-webkit-scrollbar-thumb {
    border-radius: 999px;
    background: #b8c3c8;
}

.panel-tablo-scroll::-webkit-scrollbar-thumb:hover,
.yonetim-kart-scroll::-webkit-scrollbar-thumb:hover {
    background: #929fa5;
}

table {
    width: 100%;
    min-width: 760px;
    border-collapse: collapse;
}

th,
td {
    padding: 11px 12px;
    text-align: left;
    vertical-align: middle;
    border-bottom: 1px solid #e9edef;
    font-size: var(--yazi-md);
}

th {
    position: sticky;
    top: 0;
    z-index: 2;
    background: #f6f8f9;
    color: var(--yazi-2);
    font-size: var(--yazi-xs);
    text-transform: uppercase;
    letter-spacing: .5px;
}

tbody tr:hover {
    background: #fafbfb;
}

/* --------------------------------------------------------------------------
   Modal
---------------------------------------------------------------------------- */

.modal {
    width: min(640px, calc(100% - 24px));
    max-width: none;
    max-height: calc(100vh - 32px);
    padding: 0;
    overflow: visible;
    border: 0;
    background: transparent;
}

.modal-dar {
    width: min(480px, calc(100% - 24px));
}

.modal[open] {
    display: block;
}

.modal::backdrop {
    background: rgba(9, 19, 24, .62);
    backdrop-filter: blur(2px);
}

.modal-kutu {
    width: 100%;
    max-height: calc(100vh - 40px);
    padding: 20px;
    overflow-y: auto;
    border-radius: 16px;
    background: var(--kart);
    box-shadow: 0 28px 70px rgba(0, 0, 0, .24);
}

.modal-baslik {
    display: flex;
    justify-content: space-between;
    gap: 15px;
    margin-bottom: 18px;
}

.modal-baslik h2 {
    margin: 4px 0 0;
}

.onay-mesaj {
    margin: 0 0 4px;
    color: var(--yazi-2);
    line-height: 1.55;
    white-space: pre-line;
}

.ikon-buton {
    width: 36px;
    height: 36px;
    flex-shrink: 0;
    border: 1px solid var(--cizgi);
    border-radius: 9px;
    background: var(--kart);
    cursor: pointer;
    font-size: 22px;
}

.modal-aksiyon {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 18px;
}

.iki-kolon {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
}

.modal .alan + .alan,
.modal .iki-kolon + .alan,
.modal .alan + .iki-kolon,
.modal .iki-kolon + .iki-kolon,
.modal .onay-mesaj + .alan {
    margin-top: 12px;
}

/* --------------------------------------------------------------------------
   Bildirim / toast
---------------------------------------------------------------------------- */

.bildirimler {
    margin-bottom: 16px;
    display: grid;
    gap: 8px;
}

.bildirim {
    padding: 11px 13px;
    border: 1px solid transparent;
    border-radius: 10px;
    font-size: var(--yazi-md);
    overflow-wrap: anywhere;
}

.bildirim.basari {
    background: var(--iyi-bg);
    border-color: var(--iyi-cizgi);
    color: #175d41;
}

.bildirim.hata {
    background: var(--kotu-bg);
    border-color: var(--kotu-cizgi);
    color: var(--kotu-koyu);
}

.bildirim.uyari {
    background: var(--uyari-bg);
    border-color: var(--uyari-cizgi);
    color: var(--uyari-koyu);
}

.toast-alani {
    position: fixed;
    right: 20px;
    bottom: 20px;
    z-index: 200;
    display: grid;
    gap: 8px;
    justify-items: end;
    max-width: calc(100% - 40px);
    pointer-events: none;
}

.toast {
    max-width: 420px;
    padding: 12px 14px;
    display: flex;
    align-items: flex-start;
    gap: 12px;
    border-radius: 11px;
    background: var(--ana);
    color: #fff;
    box-shadow: 0 14px 40px rgba(0, 0, 0, .16);
    opacity: 0;
    transform: translateY(8px);
    transition: opacity .18s ease, transform .18s ease;
    font-size: var(--yazi-lg);
    line-height: 1.45;
    pointer-events: auto;
}

.toast.goster {
    opacity: 1;
    transform: translateY(0);
}

.toast.hata {
    background: var(--kotu);
}

.toast.uyari {
    background: var(--uyari-koyu);
}

.toast-kapat {
    flex-shrink: 0;
    width: 26px;
    height: 26px;
    margin: -3px -4px -3px 0;
    border: 0;
    border-radius: 7px;
    background: rgba(255, 255, 255, .16);
    color: inherit;
    cursor: pointer;
    font-size: 18px;
    line-height: 1;
}

.toast-kapat:hover {
    background: rgba(255, 255, 255, .28);
}

/* --------------------------------------------------------------------------
   Giriş / 403
---------------------------------------------------------------------------- */

.giris-sayfa {
    min-height: 100vh;
    padding: 24px;
    display: grid;
    place-items: center;
    background:
        radial-gradient(circle at 20% 20%, rgba(44, 83, 100, .18), transparent 30%),
        linear-gradient(145deg, var(--ana), var(--ana-2) 50%, var(--ana-3));
}

.giris-kutu {
    width: min(420px, 100%);
    padding: 34px;
    border: 1px solid rgba(255, 255, 255, .55);
    border-radius: 20px;
    background: rgba(255, 255, 255, .98);
    box-shadow: 0 30px 80px rgba(0, 0, 0, .28);
}

.giris-kutu > .bildirim {
    margin-bottom: 14px;
}

.giris-logo {
    width: 120px;
    height: 120px;
    margin: 0 auto 18px;
}

.giris-logo img {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: cover;
    border-radius: 24px;
    box-shadow: 0 12px 30px rgba(15, 32, 39, .28);
}

.giris-kutu h1 {
    margin: 0;
    text-align: center;
    letter-spacing: -.5px;
    font-size: 27px;
}

.giris-kutu .alt {
    margin: 7px 0 24px;
    text-align: center;
    color: var(--soluk);
}

.giris-form {
    display: grid;
    gap: 15px;
}

.giris-form input,
.giris-form .buton,
.giris-gozlemci-form .buton {
    min-height: 46px;
}

.giris-gozlemci-form {
    margin-top: 12px;
}

.yetkisiz-sayfa {
    min-height: calc(100vh - 170px);
    display: grid;
    place-items: center;
}

.durum-sayfasi {
    width: min(580px, 100%);
    padding: clamp(32px, 6vw, 54px);
    text-align: center;
    border: 1px solid var(--cizgi);
    border-radius: 18px;
    background: linear-gradient(145deg, rgba(255, 255, 255, .99), rgba(248, 250, 251, .98));
    box-shadow: var(--golge);
}

.durum-ikon {
    margin: 12px 0;
    color: #c3ced3;
    font-size: clamp(60px, 11vw, 90px);
    line-height: 1;
    font-weight: 900;
}

.durum-sayfasi h1 {
    margin: 5px 0 10px;
    font-size: clamp(24px, 4vw, 32px);
}

.durum-sayfasi p:not(.ust-etiket) {
    max-width: 440px;
    margin: 0 auto 24px;
    color: var(--soluk);
    line-height: 1.6;
}

/* --------------------------------------------------------------------------
   Duyarlı yerleşim (geniş → dar sırayla; aynı özellik iki sorguda
   tanımlanıyorsa dar olan sonra gelir ki kazansın)
---------------------------------------------------------------------------- */

@media (max-height: 900px) {
    .monitor-liste {
        grid-template-rows: repeat(2, minmax(0, 1fr));
    }
}

@media (max-width: 1100px) {
    .operator-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .ust-cubuk {
        grid-template-columns: 1fr auto;
    }

    .ana-nav {
        grid-column: 1 / -1;
        order: 3;
        overflow-x: auto;
    }

    .veri-bandi {
        order: 4;
    }
}

@media (max-width: 1050px) {
    .monitor-body,
    .monitor-body .sayfa {
        height: auto;
        overflow: visible;
    }

    .monitor-body .sayfa {
        min-height: 100vh;
    }

    .monitor-grid {
        grid-template-columns: 1fr;
        gap: 14px;
    }

    .monitor-bolum {
        min-height: 520px;
    }
}

@media (max-width: 820px) {
    .sayfa-shell {
        gap: 15px;
    }

    .istatistik-grid,
    .donem-metrik-grid-4 {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .yonetim-grid,
    .operator-grid {
        grid-template-columns: 1fr;
    }

    .monitor-ust {
        align-items: flex-start;
    }

    .monitor-zaman {
        flex-shrink: 0;
    }

    .sayfa-baslik,
    .arac-cubugu {
        align-items: stretch;
        flex-direction: column;
    }

    /* Dikey yönde flex-basis yükseklik demek: arama alanı 340px'e uzamasın. */
    .operator-sayfa .arama,
    .panel-arama-alani {
        flex: 0 0 auto;
        max-width: none;
    }

    .baslik-aksiyon {
        width: 100%;
    }

    .baslik-aksiyon .buton {
        flex: 1;
    }

    .operator-filtre-alani,
    .panel-filtre-alani {
        width: 100%;
        justify-items: stretch;
    }

    .operator-filtre-meta,
    .panel-filtre-alt {
        justify-content: space-between;
    }

    .operator-teslim-notu {
        text-align: left;
    }

    .filtreler {
        width: 100%;
    }

    .filtre {
        flex: 1 1 auto;
        justify-content: center;
    }

    .bilgi-grid {
        grid-template-columns: repeat(2, 1fr);
    }

    .oturum-metin {
        display: none;
    }

    .ana-nav {
        width: 100%;
        gap: 2px;
    }

    .ana-nav a {
        flex: 1 1 auto;
        padding: 12px 10px;
        text-align: center;
    }

    .admin-filtre-cubugu {
        align-items: stretch;
        flex-direction: column;
    }

    .admin-filtre-meta {
        justify-items: start;
    }

    .yonetim-kart-panel .panel-baslik {
        align-items: stretch;
        flex-direction: column;
    }
}

@media (max-width: 640px) {
    .monitor-liste {
        grid-template-columns: 1fr;
        grid-template-rows: none;
        grid-auto-rows: minmax(180px, auto);
        overflow: auto;
    }

    .monitor-kart {
        container-type: normal;
        height: auto;
    }

    .monitor-ust {
        flex-direction: column;
    }

    .monitor-zaman {
        justify-items: start;
        text-align: left;
    }
}

@media (max-width: 520px) {
    .sayfa,
    .monitor-body .sayfa {
        width: min(100% - 20px, 1520px);
        padding-top: 18px;
    }

    .ust-cubuk {
        padding: 9px 12px;
        gap: 10px;
    }

    .marka-logo {
        width: 40px;
        height: 40px;
    }

    .panel-kutu,
    .operator-arac-cubugu,
    .panel-arac-cubugu {
        padding: 14px;
    }

    .istatistik-grid {
        gap: 8px;
    }

    .istatistik {
        min-height: 94px;
        padding: 14px;
    }

    .istatistik strong {
        font-size: 28px;
    }

    .donem-metrik-grid {
        grid-template-columns: 1fr 1fr;
    }

    .donem-ozel-alan {
        width: 100%;
        flex-wrap: wrap;
    }

    .iki-kolon {
        grid-template-columns: 1fr;
    }

    .modal-kutu {
        padding: 16px;
    }

    .is-karti-ust,
    .operator-kart-ust {
        display: grid;
    }

    .operator-rozetler {
        justify-items: start;
    }

    .durum-rozet {
        white-space: normal;
    }

    .operator-filtre-alani .filtre,
    .panel-filtre-alani .filtre,
    .admin-filtre-cubugu .filtre {
        flex: 1 1 calc(50% - 7px);
    }

    .giris-kutu {
        padding: 28px 22px;
    }

    .giris-logo {
        width: 96px;
        height: 96px;
    }

    .monitor-zaman span {
        font-size: 32px;
    }

    .yonetim-durum-eksik-liste {
        max-height: 360px;
    }

    .mini-satir {
        display: grid;
    }

    .mini-sag {
        justify-items: start;
        text-align: left;
    }

    .yonetim-kart-scroll {
        max-height: 430px;
    }

    .panel-tablo-scroll {
        max-height: 350px;
    }

    .toast-alani {
        right: 10px;
        bottom: 10px;
        left: 10px;
        max-width: none;
    }
}

/* Hareket azaltma tercihi: kayma/yükselme animasyonları kapanır. */
@media (prefers-reduced-motion: reduce) {
    *,
    *::before,
    *::after {
        transition-duration: .01ms !important;
        animation-duration: .01ms !important;
        scroll-behavior: auto !important;
    }

    .buton:hover:not(:disabled),
    .istatistik-buton:hover {
        transform: none;
    }

    .monitor-sayfa-cubugu {
        display: none;
    }
}

/* --------------------------------------------------------------------------
   Import önizleme
---------------------------------------------------------------------------- */

.onizleme-dikkat {
    line-height: 1.55;
}

/* Bu sayfanın işi açıklamak: yardımcı metinler genel 11px yerine okunur boyutta. */
.onizleme-sayfa .panel-aciklama {
    max-width: 920px;
    font-size: 13px;
}

/* Boşluksuz uzun dosya adı / hücre metni dar ekranda sayfayı genişletmesin. */
.onizleme-sayfa .sayfa-hero p,
.sorun-listesi li,
.kontrol-madde,
.fark-karti header {
    overflow-wrap: anywhere;
}

.onizleme-dikkat p {
    margin: 4px 0 0;
}

.onizleme-ozet {
    display: grid;
    grid-template-columns: repeat(5, minmax(0, 1fr));
    gap: 14px;
}

.onizleme-ozet .istatistik {
    min-height: 0;
    align-content: start;
}

.onizleme-ozet .istatistik small {
    color: var(--soluk);
    font-size: 12px;
    line-height: 1.45;
}

.onizleme-ozet .istatistik.sifir strong {
    color: #a3adb2;
}

.ozet-yeni { border-top-color: var(--iyi); }
.ozet-guncel { border-top-color: var(--dizgide); }
.ozet-ayni { border-top-color: #b8c2c7; }
.ozet-pasif { border-top-color: var(--kotu); }
.ozet-uyari { border-top-color: var(--uyari); }

.sayfa-ozet-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 12px;
}

.sayfa-ozet-kart {
    padding: 14px 16px;
    border: 1px solid var(--cizgi);
    border-radius: 12px;
    background: var(--zemin-2);
}

.sayfa-ozet-kart header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 8px;
}

.sayfa-ozet-kart header strong {
    font-size: 15px;
}

.sayfa-ozet-kart.eksik {
    background: var(--kotu-bg);
    border-color: var(--kotu-cizgi);
}

.sayfa-ozet-sayi {
    margin: 10px 0 8px;
    color: var(--soluk);
    font-size: 13px;
}

.sayfa-ozet-sayi strong {
    margin-right: 4px;
    color: var(--ana);
    font-size: 28px;
}

.sayfa-ozet-kart ul {
    margin: 0;
    padding-left: 18px;
    color: var(--soluk);
    font-size: 13px;
    line-height: 1.6;
}

.kontrol-listesi {
    margin: 0;
    padding: 0;
    display: grid;
    gap: 12px;
    list-style: none;
}

.kontrol-madde {
    padding: 14px;
    display: grid;
    grid-template-columns: 44px minmax(0, 1fr);
    gap: 14px;
    border: 1px solid var(--cizgi);
    border-radius: 12px;
}

.kontrol-sayi {
    width: 44px;
    height: 44px;
    display: grid;
    place-items: center;
    border-radius: 12px;
    background: var(--notr-bg);
    color: var(--notr);
    font-size: 16px;
    font-weight: 800;
}

.kontrol-sayi.uyari {
    background: var(--uyari-bg);
    color: var(--uyari);
}

.kontrol-madde strong {
    font-size: 14px;
}

.kontrol-madde p {
    margin: 4px 0 0;
    max-width: 920px;
    color: var(--soluk);
    font-size: 13px;
    line-height: 1.55;
}

.kontrol-madde p.kontrol-yer {
    color: var(--yazi);
    font-weight: 700;
}

.acilir {
    margin-top: 10px;
}

.acilir > summary {
    width: fit-content;
    cursor: pointer;
    color: var(--ana-2);
    font-size: 13px;
    font-weight: 700;
}

.acilir[open] > summary {
    margin-bottom: 10px;
}

.acilir ul {
    margin: 0;
    padding-left: 18px;
    color: var(--soluk);
    font-size: 12px;
    line-height: 1.6;
}

.tablo-dar table,
.fark-karti table {
    min-width: 0;
}

.kaynak-metni {
    display: block;
    margin-top: 3px;
    color: var(--soluk);
    font-size: var(--yazi-xs);
    font-weight: 400;
}

.etki-grup + .etki-grup {
    margin-top: 22px;
    padding-top: 18px;
    border-top: 1px solid var(--cizgi-acik);
}

.etki-grup h3 {
    margin: 0 0 4px;
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 16px;
}

.etki-aciklama {
    max-width: 920px;
    margin: 0 0 12px;
    color: var(--soluk);
    font-size: 13px;
    line-height: 1.55;
}

.fark-listesi {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
    gap: 12px;
}

.fark-karti {
    overflow: hidden;
    border: 1px solid var(--cizgi);
    border-radius: 12px;
}

.fark-karti header {
    padding: 10px 12px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 6px 10px;
    background: #f6f8f9;
    font-size: 13px;
}

.fark-karti header small {
    color: var(--soluk);
}

.fark-karti td.eski {
    color: var(--soluk);
}

.fark-karti td.yeni {
    font-weight: 700;
}

.fark-karti tr.onemli td {
    background: var(--uyari-bg);
}

.sifirla-bilgi {
    margin: 0 0 12px;
    padding: 12px 14px;
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 10px 16px;
    border: 1px solid var(--uyari-cizgi);
    border-radius: 10px;
    background: var(--uyari-bg);
}

.sifirla-bilgi p {
    max-width: 900px;
    margin: 0;
    color: var(--uyari-koyu);
    font-size: 13px;
    line-height: 1.55;
}

.sifirla-secenek {
    padding: 10px 12px;
    display: flex;
    align-items: flex-start;
    gap: 10px;
    border-top: 1px solid #e9edef;
    cursor: pointer;
}

.sifirla-secenek input {
    width: 18px;
    height: 18px;
    margin: 1px 0 0;
    flex: none;
    accent-color: var(--kotu);
}

.sifirla-secenek span {
    display: grid;
    gap: 2px;
    font-size: 13px;
}

.sifirla-secenek small {
    color: var(--soluk);
    font-size: 12px;
}

.sifirla-secenek:has(input:checked) {
    background: var(--kotu-bg);
}

.sifirla-secenek:has(input:checked) strong {
    color: var(--kotu);
}

.onizleme-aksiyon p b {
    color: #ffd48a;
    font-weight: 700;
}

.onizleme-aksiyon p b i {
    font-style: normal;
}

.onizleme-aksiyon p b a {
    color: inherit;
}

.karar-bandi {
    display: block;
    text-decoration: none;
    line-height: 1.55;
}

.karar-bolumu {
    border-color: var(--uyari-cizgi);
    box-shadow: 0 0 0 3px var(--uyari-bg), var(--golge);
    scroll-margin-top: 90px;
}

.karar-listesi {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
    gap: 12px;
}

.karar-karti {
    overflow: hidden;
    display: grid;
    border: 1px solid var(--cizgi);
    border-radius: 12px;
    background: var(--kart);
}

.karar-karti.uygulama-korunuyor {
    border-color: #9fd3b9;
}

.karar-karti header {
    padding: 10px 12px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 6px 10px;
    background: #f6f8f9;
    font-size: 13px;
    overflow-wrap: anywhere;
}

.karar-karti header small {
    color: var(--soluk);
}

.karar-karsilastirma {
    padding: 12px;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
}

.karar-karsilastirma > div {
    display: grid;
    align-content: start;
    gap: 5px;
}

.karar-karsilastirma > div > span {
    color: var(--soluk);
    font-size: 11px;
    font-weight: 800;
    letter-spacing: .5px;
    text-transform: uppercase;
}

.karar-karsilastirma small {
    color: var(--soluk);
    font-size: 12px;
    line-height: 1.45;
}

.karar-neden {
    margin: 0 12px;
    color: var(--uyari-koyu);
    font-size: 13px;
    line-height: 1.5;
}

.karar-secim {
    margin: 12px 12px 0;
    display: grid;
    gap: 6px;
    font-size: 13px;
    font-weight: 700;
}

.karar-secim select {
    min-height: 40px;
    padding: 8px 10px;
    border: 1px solid var(--cizgi);
    border-radius: 10px;
    background: #fff;
    font-weight: 700;
}

.karar-sonuc {
    margin: 8px 12px 12px;
    color: var(--soluk);
    font-size: 12px;
    line-height: 1.5;
}

.karar-diger {
    margin: 0;
    padding: 10px 12px;
    border-top: 1px solid #e9edef;
}

.karar-diger table {
    min-width: 0;
}

.sifirla-secenek.pasif {
    opacity: .5;
    cursor: not-allowed;
}

@media (max-width: 640px) {
    .karar-listesi,
    .karar-karsilastirma {
        grid-template-columns: 1fr;
    }
}

.onizleme-bilgi ul {
    margin: 8px 0 0;
    padding-left: 18px;
    color: var(--soluk);
    font-size: 13px;
    line-height: 1.7;
}

.sorun-listesi {
    margin: 0;
    padding: 0;
    display: grid;
    gap: 8px;
    list-style: none;
}

.sorun-listesi li {
    padding: 11px 13px;
    display: grid;
    grid-template-columns: minmax(160px, max-content) minmax(0, 1fr);
    gap: 14px;
    border: 1px solid var(--kotu-cizgi);
    border-radius: 10px;
    background: var(--kotu-bg);
    color: #6f2424;
    font-size: 13px;
    line-height: 1.5;
}

.sorun-konum {
    color: var(--kotu);
    font-weight: 800;
    white-space: nowrap;
}

.onizleme-aksiyon {
    position: sticky;
    bottom: 12px;
    z-index: 20;
    padding: 14px 18px;
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 14px;
    border-radius: 14px;
    background: rgba(15, 32, 39, .97);
    color: #fff;
    box-shadow: 0 14px 40px rgba(0, 0, 0, .18);
}

.onizleme-aksiyon p {
    margin: 0;
    display: grid;
    gap: 2px;
}

.onizleme-aksiyon p span {
    color: #b9c7cd;
    font-size: 12px;
}

.onizleme-aksiyon form {
    margin: 0;
}

.onizleme-aksiyon .buton-hayalet {
    background: rgba(255, 255, 255, .08);
    color: #fff;
    border-color: rgba(255, 255, 255, .18);
}

/* Durumu değişen kartta "Notları temizle" seçeneği ve temizlenecek not metni */
.not-bilgi {
    border-color: var(--cizgi);
    background: var(--notr-bg);
}

.not-bilgi p {
    color: var(--notr);
}

.not-secenek {
    border-top: 1px solid var(--cizgi-acik);
}

.not-secenek .sifirla-secenek {
    border-top: 0;
}

.not-metni {
    max-height: 7.5em;
    overflow: auto;
    margin: 0 12px 10px 40px;
    padding: 6px 10px;
    border-left: 3px solid var(--cizgi);
    background: var(--zemin-2);
    color: var(--soluk);
    font-size: var(--yazi-xs);
    line-height: 1.5;
    white-space: pre-line;
    overflow-wrap: anywhere;
}

.not-secenek:has(input:checked) .not-metni {
    text-decoration: line-through;
}

@media (max-width: 1100px) {
    .onizleme-ozet {
        grid-template-columns: repeat(3, minmax(0, 1fr));
    }
}

@media (max-width: 640px) {
    .onizleme-ozet {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .kontrol-madde,
    .sorun-listesi li,
    .fark-listesi {
        grid-template-columns: 1fr;
    }

    .onizleme-aksiyon {
        bottom: 0;
        border-radius: 12px 12px 0 0;
    }
}
```

## `templates/_makrolar.html`

```html
{# Birden fazla ekranda kullanılan parçalar. Tek kaynaktan üretilir ki ekranlar arasında kaymasın. #}

{# Aramada her ekranda aynı alanlar taranır. #}
{% macro arama_metni(k) -%}
{{ [k.id, k.talep_no, k.stok_no, k.talep_sahibi, k.operator, k.aciklama, k.pcb, k.dizgi_sorumlusu, k.dizgi_etiket]|select|join(" ") }}
{%- endmacro %}

{% macro dizgi_rozetleri(k, malzeme=True) -%}
    {%- if k.dizgi_kod == 'ELLE' %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif -%}
    {%- if k.dizgi_kod == 'EUM' %}<span class="durum-rozet eum">EÜM'de Dizgi</span>{% endif -%}
    {%- if malzeme and k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif -%}
{%- endmacro %}

{# Satır içi style="width:%" yerine <progress>: CSP ile uyumlu ve ekran okuyucuya değer bildirir. #}
{% macro ilerleme_cubugu(k, sinif="") -%}
<progress class="ilerleme-cubugu {{ sinif }}" value="{{ [k.adet_yuzde or 0, 100]|min }}" max="100"
          aria-label="Üretim ilerlemesi: {{ k.tamamlanan_adet }} / {{ k.toplam_adet }} adet">%{{ k.adet_yuzde }}</progress>
{%- endmacro %}

{# Durum özet kartı. Ana sayı iş emri (kart) sayısı; farklı stok sayısı alt bilgi.
   Tıklanınca o durumun filtresini seçer. Dizgi tipi filtresine göre sayılar JS ile değişir. #}
{% macro kpi_karti(sayac, anahtar, durum, sinif, baslik) -%}
<button type="button" class="istatistik istatistik-buton {{ sinif }}" data-kpi-filtre="{{ durum }}"
        aria-pressed="false" title="Yalnız {{ durum }} kartlarını listele"
        data-kart-hepsi="{{ sayac[anahtar ~ '_kart'] }}"
        data-kart-makine="{{ sayac[anahtar ~ '_kart_makine'] }}"
        data-kart-elle="{{ sayac[anahtar ~ '_kart_elle'] }}"
        data-kart-eum="{{ sayac[anahtar ~ '_kart_eum'] }}"
        data-stok-hepsi="{{ sayac[anahtar] }}"
        data-stok-makine="{{ sayac[anahtar ~ '_makine'] }}"
        data-stok-elle="{{ sayac[anahtar ~ '_elle'] }}"
        data-stok-eum="{{ sayac[anahtar ~ '_eum'] }}">
    <span>{{ baslik }}</span>
    <strong data-kpi-kart>{{ sayac[anahtar ~ '_kart'] }}</strong>
    <small><b data-kpi-stok>{{ sayac[anahtar] }}</b> farklı stok</small>
</button>
{%- endmacro %}

{% macro kpi_satiri(sayac) -%}
<section class="istatistik-grid istatistik-grid-3" aria-label="İş durumu özeti">
    {{ kpi_karti(sayac, "dizgide", "DİZGİDE", "kart-dizgide", "Dizgideki İş Emri") }}
    {{ kpi_karti(sayac, "plana_alindi", "PLANA ALINDI", "kart-plana", "Plana Alınan İş Emri") }}
    {{ kpi_karti(sayac, "teslim", "TESLİM EDİLDİ", "kart-teslim", "Teslim Edilen İş Emri") }}
</section>
{%- endmacro %}

{# Yönetim "Düzenle" / "Durum Ata" butonları aynı öznitelikleri taşımalı (bkz. Malzeme Bekliyor hatası). #}
{% macro duzenle_butonu(k, etiket="Düzenle", sinif="buton-hayalet") -%}
<button class="buton buton-kucuk {{ sinif }}" type="button" data-admin-duzenle
        data-id="{{ k.id }}"
        data-surum="{{ k.surum }}"
        data-talep="{{ k.talep_no or '' }}"
        data-stok="{{ k.stok_no or '' }}"
        data-durum="{{ k.durum or '' }}"
        data-toplam="{{ k.toplam_adet if k.toplam_adet is not none else '' }}"
        data-tamamlanan="{{ k.tamamlanan_adet if k.tamamlanan_adet is not none else 0 }}"
        data-plan-hafta="{{ k.plan_hafta or '' }}"
        data-plan-baslama="{{ k.plan_baslama or '' }}"
        data-plan-teslim="{{ k.plan_teslim or '' }}"
        data-gerceklesen-teslim="{{ k.gerceklesen_teslim or '' }}"
        data-dizgi-kod="{{ k.dizgi_kod }}"
        data-dizgi-sorumlusu="{{ k.dizgi_sorumlusu or '' }}"
        data-malzeme="{{ '1' if k.malzeme_bekliyor else '0' }}"
        data-not="{{ k.aciklama or '' }}">{{ etiket }}</button>
{%- endmacro %}
```

## `templates/base.html`

```html
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="csrf-token" content="{{ csrf_token }}">
    <title>{% block title %}PDGM İş Takip{% endblock %}</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='stil.css') }}">
    {% block head %}{% endblock %}
</head>
<body class="{% block body_class %}{% endblock %}"
      data-kullanici="{{ oturum_kullanici or '' }}"
      data-ad="{{ oturum_ad or oturum_kullanici or '' }}"
      {% if veri_surumu %}data-veri-surumu="{{ veri_surumu }}"{% endif %}>
{% block ust_cubuk %}
<header class="ust-cubuk">
    <a href="{{ url_for('ana') }}" class="marka-link">
        <img class="marka-logo" src="{{ url_for('static', filename='logo.png') }}" alt="PDGM" width="48" height="48">
        <span class="marka-metin">
            <strong>Prototip Kart Dizgi Atölyesi</strong>
        </span>
    </a>

    <nav class="ana-nav" aria-label="Ana menü">
        {% if oturum_rol in ["admin", "operator", "gozlemci"] %}
            <a class="{% if request.endpoint == 'panel' %}aktif{% endif %}" href="{{ url_for('panel') }}"{% if request.endpoint == 'panel' %} aria-current="page"{% endif %}>Pano</a>
        {% endif %}
        {% if oturum_rol in ["admin", "operator"] %}
            <a class="{% if request.endpoint == 'operator' %}aktif{% endif %}" href="{{ url_for('operator') }}"{% if request.endpoint == 'operator' %} aria-current="page"{% endif %}>Operatör</a>
        {% endif %}
        {% if oturum_rol in ["admin", "operator", "gozlemci"] %}
            <a class="{% if request.endpoint == 'monitor' %}aktif{% endif %}" href="{{ url_for('monitor') }}"{% if request.endpoint == 'monitor' %} aria-current="page"{% endif %}>Monitör</a>
        {% endif %}
        {% if oturum_rol == "admin" %}
            <a class="{% if request.endpoint == 'yonetim' %}aktif{% endif %}" href="{{ url_for('yonetim') }}"{% if request.endpoint == 'yonetim' %} aria-current="page"{% endif %}>Yönetim</a>
        {% endif %}
    </nav>

    <div class="oturum">
        <div class="oturum-metin">
            <strong>{{ oturum_ad or oturum_kullanici }}</strong>
            <small>{{ {"admin": "YÖNETİCİ", "operator": "OPERATÖR", "gozlemci": "GÖZLEMCİ"}.get(oturum_rol, (oturum_rol or "")|upper) }}</small>
        </div>
        {% if oturum_kullanici %}
        <form method="post" action="{{ url_for('cikis') }}">
            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
            <button class="buton buton-hayalet buton-kucuk" type="submit">Çıkış</button>
        </form>
        {% endif %}
    </div>

    <div id="veri-bandi" class="veri-bandi" role="status" hidden>
        <span data-veri-bandi-metin></span>
        <button type="button" class="buton buton-kucuk" data-veri-bandi-yenile>Yenile</button>
    </div>
</header>
{% endblock %}

<main class="sayfa">
    {% with mesajlar = get_flashed_messages(with_categories=true) %}
        {% if mesajlar %}
        <section class="bildirimler" aria-live="polite">
            {% for kategori, mesaj in mesajlar %}
                <div class="bildirim {{ kategori }}"{% if kategori == 'hata' %} role="alert"{% endif %}>{{ mesaj }}</div>
            {% endfor %}
        </section>
        {% endif %}
    {% endwith %}

    {% block content %}{% endblock %}
</main>

<div id="toast-alani" class="toast-alani" aria-live="polite"></div>

<dialog id="onay-dialog" class="modal modal-dar" aria-labelledby="onay-baslik">
    <form method="dialog" class="modal-kutu">
        <div class="modal-baslik">
            <div><p class="ust-etiket" id="onay-etiket">ONAY</p><h2 id="onay-baslik">Emin misiniz?</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <p id="onay-mesaj" class="onay-mesaj"></p>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana" id="onay-evet" value="evet">Devam et</button>
        </div>
    </form>
</dialog>

<script src="{{ url_for('static', filename='js/ortak.js') }}"></script>
{% block scripts %}{% endblock %}
</body>
</html>
```

## `templates/giris.html`

```html
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Giriş · PDGM İş Takip</title>
    <link rel="stylesheet" href="{{ url_for('static', filename='stil.css') }}">
</head>
<body class="giris-sayfa">
<main class="giris-kutu">
    <div class="giris-logo">
        <img src="{{ url_for('static', filename='logo.png') }}" alt="PDGM" width="120" height="120">
    </div>

    <h1>PDGM İŞ TAKİP</h1>
    <p class="alt">Prototip Kart Dizgi Atölyesi</p>

    {% for kategori, mesaj in get_flashed_messages(with_categories=true) %}
        <div class="bildirim {{ kategori }}"{% if kategori == 'hata' %} role="alert"{% endif %}>{{ mesaj }}</div>
    {% endfor %}
    {% if hata %}<div class="bildirim hata" role="alert">{{ hata }}</div>{% endif %}

    <form method="post" class="giris-form">
        <div class="alan">
            <label for="kullanici">Kullanıcı adı</label>
            <input type="text" id="kullanici" name="kullanici" autocomplete="username" autocapitalize="none" autofocus required>
        </div>
        <div class="alan">
            <label for="sifre">Şifre</label>
            <input type="password" id="sifre" name="sifre" autocomplete="current-password" required>
        </div>
        <button class="buton buton-ana tam-genislik" type="submit">Giriş Yap</button>
    </form>

    <form method="post" action="{{ url_for('giris_gozlemci') }}" class="giris-gozlemci-form">
        <button class="buton buton-hayalet tam-genislik" type="submit">Gözlemci olarak giriş yap</button>
    </form>
</main>
</body>
</html>
```

## `templates/import_onizleme.html`

```html
{% extends "base.html" %}
{% block title %}Import Önizleme · PDGM İş Takip{% endblock %}

{% macro tur_rozeti(kod) -%}
    {%- if kod == "ELLE" -%}<span class="durum-rozet elle">Elle Dizgi</span>
    {%- elif kod == "EUM" -%}<span class="durum-rozet eum">EÜM'de Dizgi</span>
    {%- else -%}<span class="durum-rozet">Makine</span>{%- endif -%}
{%- endmacro %}

{% macro durum_rozeti(durum, excel_durum=None) -%}
    {%- if durum == "TESLİM EDİLDİ" -%}<span class="durum-rozet iyi">{{ durum }}</span>
    {%- elif durum == "DİZGİDE" -%}<span class="durum-rozet uyari">{{ durum }}</span>
    {%- elif durum -%}<span class="durum-rozet">{{ durum }}</span>
    {%- else -%}<span class="durum-rozet kotu">Durum yok</span>{%- endif -%}
    {%- if excel_durum and excel_durum != durum -%}<small class="kaynak-metni">Excel: {{ excel_durum }}</small>{%- endif -%}
{%- endmacro %}

{% macro not_temizle_secenegi(k) -%}
<div class="not-secenek">
    <label class="sifirla-secenek">
        <input type="checkbox" name="not_temizle" value="{{ k.id }}" form="onay-formu" data-not-temizle>
        <span>
            <strong>Notları temizle</strong>
            <small>Eski notlar yeni duruma taşınmasın istiyorsanız işaretleyin; işaretlemezseniz korunur. Silinen notlar işlem loguna yazılır.</small>
        </span>
    </label>
    <p class="not-metni" aria-label="Kartın mevcut notları">{{ k.aciklama }}</p>
</div>
{%- endmacro %}

{% block content %}
<div class="sayfa-shell onizleme-sayfa">
{% if hata %}
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">EXCEL AKTARIMI</p>
            <h1>{{ "Excel kabul edilmedi" if hata_turu == "excel" else "Aktarım uygulanmadı" }}</h1>
            <p class="soluk">{% if dosya %}<strong>{{ dosya }}</strong> · {% endif %}Hiçbir kart, not veya yedek değişmedi.</p>
        </div>
        <div class="baslik-aksiyon">
            <a class="buton buton-hayalet" href="{{ url_for('yonetim') }}">Yönetime Dön</a>
        </div>
    </section>

    <section class="panel-kutu" role="alert">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">{{ sorunlar|length }} SORUN</p>
                {% if hata_turu == "excel" %}
                <h2>Dosyada düzeltilmesi gerekenler</h2>
                <p class="panel-aciklama">Bu satırları Excel'de düzeltip dosyayı kaydedin, ardından Yönetim ekranından yeniden yükleyin. Tek bir satır hatalıyken diğer satırlar da uygulanmaz; böylece yarım okunmuş bir dosya kartları ezemez.</p>
                {% else %}
                <h2>Neden uygulanmadı</h2>
                <p class="panel-aciklama">Önizleme ile onay arasında dosya, Excel'in hesapladığı değerler veya sistemdeki kartlar değişti, bu önizleme zaten kullanıldı ya da bu oturumda daha sonra başka bir dosya önizlendi. Dosyayı yeniden seçip yeni bir önizleme oluşturun.</p>
                {% endif %}
            </div>
        </div>
        <ol class="sorun-listesi">
            {% for s in sorunlar %}
            <li>
                {% if s.konum %}<span class="sorun-konum">{{ s.konum }}</span>{% endif %}
                <span>{{ s.mesaj }}</span>
            </li>
            {% endfor %}
        </ol>
    </section>
{% else %}
    {% set r = sonuc.kaynak_raporu or {} %}
    {% set d = sonuc.degisiklikler %}
    {% set yeniler = d|selectattr("tur", "equalto", "yeni")|list %}
    {% set guncellenecek = d|selectattr("tur", "equalto", "guncelleme")|list %}
    {% set pasifler = d|selectattr("tur", "equalto", "pasif")|list %}
    {% set geri_gelenler = d|selectattr("tur", "equalto", "geri")|list %}
    {% set ayrilanlar = d|selectattr("tur", "equalto", "ayrildi")|list %}
    {% set gerilenler = d|selectattr("tur", "equalto", "gerileme")|list %}
    {% set durumsuzlar = gerilenler|selectattr("gerileme.karar_turu", "equalto", "durumsuz")|list %}
    {% set geride_kalanlar = gerilenler|rejectattr("gerileme.karar_turu", "equalto", "durumsuz")|list %}
    {% set not_temizlenebilirler = d|selectattr("not_temizlenebilir")|list %}
    {% set pasiflesecek = (sonuc.pasife_alinan or 0) + (sonuc.ayrilan or 0) %}
    {% set eksik_sayfalar = (r.sayfalar or [])|selectattr("bulunamadi")|list %}
    {% set toplu_pasif = sonuc.pasife_alinan and sonuc.mevcut_aktif and sonuc.pasife_alinan * 10 >= sonuc.mevcut_aktif * 3 %}
    {% set guncellenen = sonuc.guncellenen or 0 %}
    {% set degismeyen = sonuc.degismeyen or 0 %}
    {% set atlanan = r.atlanan_satir or 0 %}

    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">ÖNİZLEME · HENÜZ HİÇBİR KAYIT DEĞİŞMEDİ</p>
            <h1>Excel aktarımını kontrol edin</h1>
            <p class="soluk"><strong>{{ dosya }}</strong> dosyasından {{ sonuc.satir }} kayıt okundu ve sistemdeki kartlarla karşılaştırıldı. Aşağıdakiler, onaylarsanız ne olacağını gösterir.</p>
        </div>
    </section>

    {% if gerilenler %}
    <a class="bildirim uyari karar-bandi" href="#karar-gerekiyor">
        {% if geride_kalanlar %}<strong>{{ geride_kalanlar|length }} kartta Excel uygulamanın gerisinde.</strong>{% endif %}
        {% if durumsuzlar %}<strong>{{ durumsuzlar|length }} kartta Excel'de DURUM boş, uygulamada ise bir durum var.</strong>{% endif %}
        Onaylamadan önce bu kartların hangi durumda kalacağını seçin ↓
    </a>
    {% endif %}

    {% if toplu_pasif or eksik_sayfalar %}
    <div class="bildirim hata onizleme-dikkat" role="alert">
        <strong>Onaylamadan önce kontrol edin.</strong>
        {% if eksik_sayfalar %}
        <p>{{ eksik_sayfalar|map(attribute="sayfa")|join(", ") }} sayfası dosyada yok; bu sayfadan gelmiş kartların hepsi pasifleşir.</p>
        {% endif %}
        {% if toplu_pasif %}
        <p>Sistemde Excel'den gelen {{ sonuc.mevcut_aktif }} aktif karttan {{ sonuc.pasife_alinan }} tanesi pasifleşecek. Yanlış veya eksik bir dosya seçilmiş olabilir.</p>
        {% endif %}
    </div>
    {% endif %}

    <section class="onizleme-ozet" aria-label="Aktarımın kartlara etkisi">
        <article class="istatistik ozet-yeni{% if not sonuc.yeni %} sifir{% endif %}">
            <span>Yeni kart</span><strong>{{ sonuc.yeni }}</strong>
            <small>Excel'de olup sistemde olmayan kayıtlar oluşturulur.</small>
        </article>
        <article class="istatistik ozet-guncel{% if not guncellenen %} sifir{% endif %}">
            <span>Güncellenecek</span><strong>{{ guncellenen }}</strong>
            <small>Tarih, adet veya durum Excel'deki değere çekilir.</small>
        </article>
        <article class="istatistik ozet-ayni{% if not degismeyen %} sifir{% endif %}">
            <span>Değişmeyecek</span><strong>{{ degismeyen }}</strong>
            <small>Excel ile sistem zaten aynı; bu kartlara dokunulmaz.</small>
        </article>
        <article class="istatistik ozet-pasif{% if not pasiflesecek %} sifir{% endif %}">
            <span>Pasifleşecek</span><strong>{{ pasiflesecek }}</strong>
            <small>Excel'de artık yok veya NO'su başka talebe verilmiş; silinmez, ekranlardan kalkar.</small>
        </article>
        <article class="istatistik ozet-uyari{% if not sonuc.uyari %} sifir{% endif %}">
            <span>Veri uyarısı</span><strong>{{ sonuc.uyari }}</strong>
            <small>Aktarımı engellemez; aşağıda neyi etkilediği yazıyor.</small>
        </article>
    </section>

    {% if gerilenler %}
    <section class="panel-kutu karar-bolumu" id="karar-gerekiyor">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KARAR GEREKİYOR · {{ gerilenler|length }} KART</p>
                <h2>{% if geride_kalanlar and durumsuzlar %}Excel'in geride kaldığı veya DURUM'u boş bıraktığı kartlar{% elif durumsuzlar %}Excel'de DURUM'u boş olan kartlar{% else %}Excel'in uygulamanın gerisinde kaldığı kartlar{% endif %}</h2>
                <p class="panel-aciklama">
                    {% if geride_kalanlar %}Excel, uygulamadaki durumdan daha geri bir durum söylüyor olabilir (ör. operatör kartı dizgiye almış ama Excel hâlâ PLANA ALINDI diyor).{% endif %}
                    {% if durumsuzlar %}Excel'de DURUM boş ya da MALZEME TEDARİK / PDGM ÖNERİ gibi bir iş akışı durumu olmayan bir metin olabilir; kart durumsuz bırakılırsa Pano, Operatör ve Monitör'den kalkar ve Durumu Eksik Kartlar listesine düşer.{% endif %}
                    Her kart için onaydan sonraki durumu seçin. Önerilen seçim: durum uygulamada verildiyse (operatör ilerletti veya admin atadı) uygulamadaki durum korunur; durum Excel'den geldiyse Excel'e uyulur.
                </p>
            </div>
            {% if gerilenler|length > 1 %}
            <div class="buton-grup">
                <button type="button" class="buton buton-hayalet buton-kucuk" data-gerileme-hepsi="uygulama">Hepsinde uygulamadakini koru</button>
                <button type="button" class="buton buton-hayalet buton-kucuk" data-gerileme-hepsi="excel">Hepsini Excel'e göre ayarla</button>
            </div>
            {% endif %}
        </div>
        <div class="karar-listesi">
            {% for k in gerilenler %}{% set g = k.gerileme %}
            <article class="karar-karti" data-gerileme-kart>
                <header>
                    <strong>{{ k.talep_no }} · {{ k.stok_no }}</strong>
                    {{ tur_rozeti(k.sayfa) }}
                    <small>NO {{ k.sira or "—" }} · Kart #{{ k.id }}</small>
                </header>
                <div class="karar-karsilastirma">
                    <div>
                        <span>Uygulamada</span>
                        {{ durum_rozeti(g.uygulama_durum) }}
                        <small>{{ g.onceki_tamamlanan }} / {{ k.toplam_adet }} adet tamamlandı{% if g.baslama_zamani %} · başlama {{ g.baslama_zamani|gun }}{% endif %}{% if g.onceki_gerceklesen %} · teslim {{ g.onceki_gerceklesen|gun }}{% endif %}{% if g.operator and g.operator != "Excel" %} · {{ g.operator }}{% endif %}</small>
                    </div>
                    <div>
                        <span>Excel'de</span>
                        {% if g.karar_turu == "durumsuz" %}
                        <span class="durum-rozet kotu">{{ k.excel_durum or "DURUM boş" }}</span>
                        <small>İş akışı durumu değil.</small>
                        {% else %}
                        {{ durum_rozeti(g.excel_durum, k.excel_durum) }}
                        {% endif %}
                        {% if g.onceki_excel_durum %}<small>Önceki yüklemede Excel: {{ g.onceki_excel_durum }}</small>{% endif %}
                    </div>
                </div>
                <p class="karar-neden">
                    {%- if g.karar_turu == "durumsuz" -%}
                        {% if g.uygulamada %}Durum uygulamada verilmiş (operatör işlemi veya admin ataması); Excel bu kart için durum söylemiyor.{% else %}Bu durum daha önce Excel'den gelmişti; {% if k.excel_durum %}Excel'deki yeni DURUM bir iş akışı durumu değil.{% else %}Excel'de DURUM silinmiş görünüyor.{% endif %}{% endif %}
                    {%- elif g.uygulamada -%}Kart uygulamada ilerletilmiş; Excel henüz güncellenmemiş olabilir.
                    {%- else -%}Bu durum daha önce Excel'den gelmişti; Excel'de geri alınmış görünüyor.{%- endif -%}
                </p>
                <label class="karar-secim">
                    <span>Onaydan sonraki durum</span>
                    <select name="gerileme_{{ k.id }}" form="onay-formu" data-gerileme-secim
                            data-excel="{{ g.excel_durum or '' }}" data-uygulama="{{ g.uygulama_durum }}">
                        {% for s in g.secenekler %}
                        <option value="{{ s or '' }}"{% if s == g.onerilen %} selected{% endif %}>{{ s or "Durumsuz bırak" }}{% if s == g.excel_durum %} · Excel'deki{% elif s == g.uygulama_durum %} · uygulamadaki{% endif %}{% if s == g.onerilen %} (önerilen){% endif %}</option>
                        {% endfor %}
                    </select>
                </label>
                <p class="karar-sonuc" data-gerileme-sonuc aria-live="polite"></p>
                {% if "DİZGİDE" in g.secenekler and g.onceki_tamamlanan > 0 %}
                <label class="sifirla-secenek">
                    <input type="checkbox" name="sifirla" value="{{ k.id }}" form="onay-formu" data-sifirla data-gerileme-sifirla>
                    <span>
                        <strong>Tamamlanan adedi sıfırla</strong>
                        <small>Şu an {{ g.onceki_tamamlanan }} / {{ k.toplam_adet }}. Yalnız DİZGİDE seçiliyken kullanılabilir; işaretlemezseniz korunur.</small>
                    </span>
                </label>
                {% endif %}
                {% if k.not_temizlenebilir %}{{ not_temizle_secenegi(k) }}{% endif %}
                {% set diger = k.farklar|rejectattr("anahtar", "in", ["source_active", "durum", "tamamlanan_adet"])|list %}
                {% if diger %}
                <details class="acilir karar-diger">
                    <summary>Excel'den gelen diğer değişiklikler ({{ diger|length }})</summary>
                    <table>
                        <thead><tr><th>Alan</th><th>Şu an</th><th>Onaydan sonra</th></tr></thead>
                        <tbody>
                        {% for f in diger %}
                            <tr><td>{{ f.alan }}</td><td class="eski">{{ f.eski|onizleme_deger(f.anahtar) }}</td><td class="yeni">{{ f.yeni|onizleme_deger(f.anahtar) }}</td></tr>
                        {% endfor %}
                        </tbody>
                    </table>
                </details>
                {% endif %}
            </article>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    <section class="panel-kutu">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">1 · EXCEL'DEN OKUNANLAR</p>
                <h2>Sayfa bazında kayıtlar</h2>
                <p class="panel-aciklama">Talep NO ve Kart Stok No'su dolu her satır bir karttır. Gizli satır ve sütunlar da okunur; tamamen boş satırlar yok sayılır.</p>
            </div>
        </div>
        <div class="sayfa-ozet-grid">
            {% for s in r.sayfalar or [] %}
            <article class="sayfa-ozet-kart{% if s.bulunamadi %} eksik{% endif %}">
                <header><strong>{{ s.sayfa }}</strong>{{ tur_rozeti(s.kod) }}</header>
                {% if s.bulunamadi %}
                <p class="sayfa-ozet-sayi">Sayfa dosyada yok.</p>
                {% else %}
                <p class="sayfa-ozet-sayi"><strong>{{ s.kayit }}</strong> kart okundu</p>
                <ul>
                    {% if s.gizli_kayit %}<li>{{ s.gizli_kayit }} tanesi gizli satırda; normal işlendi.</li>{% endif %}
                    <li>{% if s.atlanan %}{{ s.atlanan|length }} satır kart sayılmadı: {{ s.atlanan_araliklar }}{% else %}Kart sayılmayan satır yok.{% endif %}</li>
                    <li>{{ s.bos_satir }} boş satır atlandı.</li>
                </ul>
                {% endif %}
            </article>
            {% endfor %}
        </div>
    </section>

    {% if r.uyari_gruplari or atlanan %}
    <section class="panel-kutu">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">2 · KONTROL EDİN</p>
                <h2>Onaylamadan önce bakılması gerekenler</h2>
                <p class="panel-aciklama">Bunlar aktarımı durdurmaz, ama ilgili kartların ekranlarda nasıl görüneceğini etkiler. Excel'de düzeltip yeniden yüklerseniz uyarılar kalkar.</p>
            </div>
        </div>
        <ul class="kontrol-listesi">
            {% for g in r.uyari_gruplari or [] %}
            <li class="kontrol-madde">
                <span class="kontrol-sayi uyari">{{ g.adet }}</span>
                <div>
                    <strong>{{ g.baslik }}</strong>
                    <p>{{ g.aciklama }}</p>
                    <p class="kontrol-yer">{% for y in g.yerler %}{{ y.sayfa }}: satır {{ y.satirlar }}{% if not loop.last %} · {% endif %}{% endfor %}</p>
                    {% if g.tur == "plan_baslangic" %}
                    <details class="acilir">
                        <summary>Satır ayrıntısı</summary>
                        <ul>{% for m in g.mesajlar %}<li>{{ m }}</li>{% endfor %}</ul>
                    </details>
                    {% endif %}
                </div>
            </li>
            {% endfor %}
            {% if atlanan %}
            <li class="kontrol-madde">
                <span class="kontrol-sayi">{{ atlanan }}</span>
                <div>
                    <strong>Kart sayılmayan satırlar</strong>
                    <p>Bu satırlarda Talep NO veya Kart Stok No yok; adet, DURUM ve plan/teslim tarihi de bulunmadığı için talep kaydı değil, eski veya şablon kalıntısı kabul edildi. Kart olması gerekiyorsa Excel'de Talep NO ve Kart Stok No'yu doldurun.</p>
                    <details class="acilir">
                        <summary>Hangi satırlar, hangi hücreler dolu?</summary>
                        <div class="tablo-kapsayici tablo-dar">
                            <table>
                                <thead><tr><th>Sayfa</th><th>Satırlar</th><th>Dolu hücreler</th><th>Gizli</th></tr></thead>
                                <tbody>
                                {% for s in r.sayfalar or [] %}{% for a in s.atlanan_gruplari or [] %}
                                    <tr>
                                        <td>{{ s.sayfa }}</td>
                                        <td>{{ a.satirlar }} <small class="kaynak-metni">{{ a.adet }} satır</small></td>
                                        <td>{{ a.hucreler|join(", ") }}</td>
                                        <td>{% if a.gizli == a.adet %}Hepsi{% elif a.gizli %}{{ a.gizli }} / {{ a.adet }}{% else %}Hayır{% endif %}</td>
                                    </tr>
                                {% endfor %}{% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </details>
                </div>
            </li>
            {% endif %}
        </ul>
    </section>
    {% endif %}

    <section class="panel-kutu">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">{{ 3 if (r.uyari_gruplari or atlanan) else 2 }} · KARTLARA ETKİSİ</p>
                <h2>Onaylarsanız değişecek kartlar</h2>
                <p class="panel-aciklama">Excel ana kaynaktır: tarih, adet ve geçerli DURUM Excel'deki değere çekilir. Gizlenen ve elle eklenen kartlar korunur; operatör notları da siz "Notları temizle" demedikçe korunur.</p>
            </div>
        </div>

        {% if not d %}
        <p class="bos-durum">Kartlarda değişiklik yok: Excel ve sistem aynı. Onaylamak yalnız yükleme kaydı ve yedek oluşturur.</p>
        {% endif %}

        {% if ayrilanlar %}
        <div class="etki-grup">
            <h3><span class="durum-rozet kotu">{{ ayrilanlar|length }}</span> NO'su başka talebe verilmiş kartlar ayrılacak</h3>
            <p class="etki-aciklama">Excel'de bu NO artık farklı bir Talep NO taşıyor. Eski talebin kartı notları, operatörü ve durumuyla birlikte pasifleşir ve Yönetim ekranında kalır; yeni talep için Excel verisiyle temiz bir kart açılır (yeni kartlar listesinde). Böylece eski talebin durumu veya notu yeni talebe geçmez. Talep NO'yu yanlışlıkla değiştirdiyseniz sorun değil: Excel'i düzeltip yeniden yüklediğinizde eski kart notları ve iş akışıyla geri bağlanır.</p>
            <div class="tablo-kapsayici">
                <table>
                    <thead><tr><th>Tür</th><th>NO</th><th>Eski talep</th><th>Eski durum</th><th>Excel'deki yeni talep</th></tr></thead>
                    <tbody>
                    {% for k in ayrilanlar %}
                        <tr>
                            <td>{{ tur_rozeti(k.sayfa) }}</td><td>{{ k.sira or "—" }}</td>
                            <td>{{ k.talep_no }} · {{ k.stok_no }}</td>
                            <td>{{ durum_rozeti(k.durum) }}</td>
                            <td><strong>{{ k.yerine.talep_no }} · {{ k.yerine.stok_no }}</strong></td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        {% endif %}

        {% if pasifler %}
        <div class="etki-grup">
            <h3><span class="durum-rozet kotu">{{ pasifler|length }}</span> Excel'de artık olmayan kartlar pasifleşecek</h3>
            <p class="etki-aciklama">Kart silinmez: "kaynakta yok" olarak işaretlenir, Pano, Operatör ve Monitör ekranlarından kalkar, Yönetim ekranında listelenir. Satır Excel'e geri eklenirse aynı kart notları ve geçmişiyle geri gelir.</p>
            <div class="tablo-kapsayici">
                <table>
                    <thead><tr><th>Tür</th><th>NO</th><th>Talep NO</th><th>Kart Stok No</th><th>Talep sahibi</th><th>Son durum</th></tr></thead>
                    <tbody>
                    {% for k in pasifler %}
                        <tr>
                            <td>{{ tur_rozeti(k.sayfa) }}</td><td>{{ k.sira or "—" }}</td>
                            <td>{{ k.talep_no }}</td><td>{{ k.stok_no }}</td><td>{{ k.talep_sahibi or "—" }}</td>
                            <td>{{ durum_rozeti(k.durum) }}</td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        {% endif %}

        {% for baslik, aciklama, grup, rozet in [
            ("güncellenecek", "Yalnız değişen alanlar gösteriliyor. Sarı satırlar kartın iş durumunu veya tamamlanan adedini değiştirir.", guncellenecek, "uyari"),
            ("Excel'e geri eklendiği için yeniden aktifleşecek", "Bu kartlar daha önce Excel'de yoktu ya da NO'ları başka talebe verilip eski talebine döndü; aynı kart notları ve geçmişiyle geri gelir.", geri_gelenler, "iyi")] if grup %}
        <div class="etki-grup">
            <h3><span class="durum-rozet {{ rozet }}">{{ grup|length }}</span> Kart {{ baslik }}</h3>
            <p class="etki-aciklama">{{ aciklama }}</p>
            {% set sifirlanabilirler = grup|selectattr("sifirlanabilir")|list %}
            {% if sifirlanabilirler %}
            <div class="sifirla-bilgi">
                <p><strong>Tamamlanan adet korunacak:</strong> {{ sifirlanabilirler|length }} kart onaydan sonra DİZGİDE kalıyor ve operatörün girdiği tamamlanan adet olduğu gibi korunur. Excel'e göre üretime baştan başlanacaksa kartın altındaki kutuyu işaretleyin: tamamlanan adet 0 olur ve bitiş zamanı temizlenir; notlar ve başlama bilgisi korunur.</p>
                {% if sifirlanabilirler|length > 1 %}
                <button type="button" class="buton buton-hayalet buton-kucuk" data-hepsini-sec>Tümünü işaretle</button>
                {% endif %}
            </div>
            {% endif %}
            {% set notlular = grup|selectattr("not_temizlenebilir")|list %}
            {% if notlular %}
            <div class="sifirla-bilgi not-bilgi">
                <p><strong>Notlar korunacak:</strong> {{ notlular|length }} kartın durumu bu aktarımla değişiyor ve kartta operatör/admin notu var. Notlar eski duruma aitse kartın altındaki "Notları temizle" kutusunu işaretleyin; silinen not metni işlem loguna yazılır.</p>
                {% if notlular|length > 1 %}
                <button type="button" class="buton buton-hayalet buton-kucuk" data-hepsini-sec="not">Tümünü işaretle</button>
                {% endif %}
            </div>
            {% endif %}
            <div class="fark-listesi">
                {% for k in grup %}
                <article class="fark-karti">
                    <header>
                        <strong>{{ k.talep_no }} · {{ k.stok_no }}</strong>
                        {{ tur_rozeti(k.sayfa) }}
                        <small>NO {{ k.sira or "—" }} · Kart #{{ k.id }}</small>
                        {% if k.geri_baglandi %}<span class="durum-rozet iyi">NO eski talebine döndü · geçmişiyle geri bağlanır</span>{% endif %}
                    </header>
                    <table>
                        <thead><tr><th>Alan</th><th>Şu an</th><th>Onaydan sonra</th></tr></thead>
                        <tbody>
                        {% for f in k.farklar if f.anahtar != "source_active" %}
                            <tr{% if f.anahtar in ("durum", "tamamlanan_adet") %} class="onemli"{% endif %}>
                                <td>{{ f.alan }}</td>
                                <td class="eski">{{ f.eski|onizleme_deger(f.anahtar) }}</td>
                                <td class="yeni">{{ f.yeni|onizleme_deger(f.anahtar) }}</td>
                            </tr>
                        {% else %}
                            <tr><td colspan="3">Alan değişikliği yok; yalnız yeniden aktifleşir.</td></tr>
                        {% endfor %}
                        </tbody>
                    </table>
                    {% if k.sifirlanabilir %}
                    <label class="sifirla-secenek">
                        <input type="checkbox" name="sifirla" value="{{ k.id }}" form="onay-formu" data-sifirla>
                        <span>
                            <strong>Tamamlanan adedi sıfırla</strong>
                            <small>Şu an {{ k.tamamlanan_adet }} / {{ k.toplam_adet }} tamamlanmış görünüyor. İşaretlemezseniz korunur.</small>
                        </span>
                    </label>
                    {% endif %}
                    {% if k.not_temizlenebilir %}{{ not_temizle_secenegi(k) }}{% endif %}
                </article>
                {% endfor %}
            </div>
        </div>
        {% endfor %}

        {% if yeniler %}
        <div class="etki-grup">
            <h3><span class="durum-rozet iyi">{{ yeniler|length }}</span> Yeni kart oluşturulacak</h3>
            <p class="etki-aciklama">Bu kayıtlar sistemde yok. "Durum yok" olanlar oluşturulur ama DURUM atanana kadar operasyon ekranlarında görünmez. Excel'deki durum metni sistemdeki karşılığından farklıysa altında yazar.</p>
            {% if yeniler|length > 12 %}
            <details class="acilir">
                <summary>{{ yeniler|length }} yeni kartın listesini göster</summary>
            {% endif %}
            <div class="tablo-kapsayici">
                <table>
                    <thead><tr><th>Tür</th><th>NO</th><th>Talep NO</th><th>Kart Stok No</th><th>Talep sahibi</th><th>Adet</th><th>Durum</th><th>Plan başlangıç</th><th>Plan teslim</th><th>Gerçekleşen teslim</th></tr></thead>
                    <tbody>
                    {% for k in yeniler %}
                        <tr>
                            <td>{{ tur_rozeti(k.sayfa) }}</td><td>{{ k.sira or "—" }}</td>
                            <td>{{ k.talep_no }}</td><td>{{ k.stok_no }}</td><td>{{ k.talep_sahibi or "—" }}</td>
                            <td>{{ k.toplam_adet }}</td><td>{{ durum_rozeti(k.durum, k.excel_durum) }}</td>
                            <td>{{ k.plan_baslama|gun }}</td><td>{{ k.plan_teslim|gun }}</td><td>{{ k.gerceklesen_teslim|gun }}</td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
            {% if yeniler|length > 12 %}</details>{% endif %}
        </div>
        {% endif %}
    </section>

    <section class="panel-kutu onizleme-bilgi">
        <p class="ust-etiket">ONAYLARSANIZ</p>
        <ul>
            <li>Önce kart, işlem logu ve yükleme geçmişi dosyalarının yedeği alınır.</li>
            <li>Değişiklikler tek seferde uygulanır; bir hata olursa hiçbiri uygulanmaz.</li>
            <li>Excel'de boş bırakılan tarih, kartta da temizlenir. DURUM boş ya da MALZEME TEDARİK / PDGM ÖNERİ ise yeni kart durumsuz oluşur; durumu olan kartta yukarıdaki kararınız uygulanır. Tanınmayan DURUM aktarımı durdurur.</li>
            <li>Tamamlanan adet: PLANA ALINDI olan veya durumsuz bırakılan kartta 0'a, TESLİM EDİLDİ olan kartta toplam adede kendiliğinden eşitlenir. DİZGİDE kalan kartta korunur; yalnız işaretlediğiniz kartlarda sıfırlanır.</li>
            <li>Notlar korunur; yalnız "Notları temizle" işaretlediğiniz kartlarda silinir ve silinen metin işlem loguna yazılır.</li>
            <li>Bu sayfa açıkken dosya veya kartlar değişirse aktarım uygulanmaz ve yeni önizleme istenir.</li>
        </ul>
    </section>

    <div class="onizleme-aksiyon">
        <p>
            <strong>{{ sonuc.yeni }} yeni · {{ guncellenen }} güncelleme · {{ pasiflesecek }} pasifleşme</strong>
            <span>Onaylayana kadar hiçbir kayıt değişmez.{% if gerilenler %}<b data-gerileme-ozet> · <a href="#karar-gerekiyor">{{ gerilenler|length }} kart için durum kararı</a>: <i>0</i> kartta uygulamadaki durum korunacak.</b>{% endif %}<b data-sifirla-ozet hidden> · <i>0</i> kartta tamamlanan adet sıfırlanacak.</b>{% if not_temizlenebilirler %}<b data-not-ozet hidden> · <i>0</i> kartta notlar temizlenecek.</b>{% endif %}</span>
        </p>
        <div class="buton-grup">
            <a class="buton buton-hayalet" href="{{ url_for('yonetim') }}">Vazgeç</a>
            <form id="onay-formu" action="{{ url_for('yukle_onay') }}" method="post"
                  data-tek-gonderim data-gonderim-metni="Uygulanıyor…">
                <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                <input type="hidden" name="onizleme_token" value="{{ onizleme_token or '' }}">
                <button class="buton {{ 'buton-tehlike' if toplu_pasif else 'buton-basari' }}" type="submit">{{ "Yine de uygula" if toplu_pasif else "Değişiklikleri uygula" }}</button>
            </form>
        </div>
    </div>
{% endif %}
</div>
{% endblock %}

{% block scripts %}
{% if not hata %}<script src="{{ url_for('static', filename='js/onizleme.js') }}"></script>{% endif %}
{% endblock %}
```

## `templates/monitor.html`

```html
{% extends "base.html" %}
{% from "_makrolar.html" import ilerleme_cubugu %}

{% block title %}Monitör · PDGM İş Takip{% endblock %}
{% block body_class %}monitor-body{% endblock %}

{# Atölye TV'si: menü ve Çıkış butonu gösterilmez (kiosk). #}
{% block ust_cubuk %}{% endblock %}

{% macro monitor_karti(k, plan=False) -%}
<article class="monitor-kart {{ 'monitor-kart-plan' if plan }} {{ k.renk }}" data-monitor-kart>
    <div class="monitor-kart-govde">
        <div class="monitor-kart-ust">
            <div class="monitor-talep">
                <span>TALEP</span>
                <strong>{{ k.talep_no or "—" }}</strong>
            </div>
            <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
        </div>
        {% if k.malzeme_bekliyor %}
        <div class="monitor-ek-rozetler"><span class="durum-rozet uyari">Malzeme Bekliyor</span></div>
        {% endif %}
        <h3>{{ k.stok_no or "Stok no yok" }}</h3>
        <p class="monitor-sahip">{{ k.talep_sahibi or "—" }}</p>
        {% if plan %}
        <div class="monitor-adet-satir">
            <strong>{{ k.toplam_adet }}</strong>
            <span>adet planlandı</span>
        </div>
        {% else %}
        <div class="monitor-adet-satir">
            <strong>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }}</strong>
            <span>adet · %{{ k.adet_yuzde }}</span>
        </div>
        {{ ilerleme_cubugu(k, "monitor-ilerleme") }}
        {% endif %}
    </div>
    <div class="monitor-tarihler">
        <div>
            <div class="monitor-tarih-etiket">Başlangıç</div>
            <div class="monitor-tarih-deger">{{ k.plan_baslama|gun }}</div>
        </div>
        <div>
            <div class="monitor-tarih-etiket">Teslim</div>
            <div class="monitor-tarih-deger">{{ k.plan_teslim|gun }}</div>
        </div>
    </div>
</article>
{%- endmacro %}

{% macro sayfa_bilgisi() -%}
<div class="monitor-sayfa-bilgi" data-sayfa-bilgi hidden>
    <span data-sayfa-metin>Sayfa 1/1</span>
    <span class="monitor-sayfa-cubugu" aria-hidden="true"><i data-sayfa-ilerleme></i></span>
</div>
{%- endmacro %}

{% block content %}
<div class="monitor-sayfa">

    <section class="monitor-ust">
        <div>
            <p class="ust-etiket">ATÖLYE CANLI GÖRÜNÜMÜ · <a class="monitor-link" href="{{ url_for('panel') }}">Pano'ya geç</a></p>
            <h1>Makine Dizgi Monitörü</h1>
            <p class="monitor-aciklama">Yalnız makine dizgi kartları. Elle ve EÜM'de dizgi kartları Pano'da.</p>
        </div>
        <div class="monitor-zaman">
            <span id="monitor-saat">{{ guncelleme }}</span>
            <small>{{ bugun }} · Veri: <b id="monitor-veri-zamani">{{ guncelleme }}</b></small>
        </div>
    </section>

    <div id="monitor-baglanti" class="monitor-baglanti" role="alert" hidden></div>

    <section class="monitor-grid">

        <article class="monitor-bolum monitor-dizgide">
            <div class="monitor-bolum-baslik">
                <div>
                    <p class="ust-etiket">AKTİF ÜRETİM</p>
                    <h2>DİZGİDE</h2>
                </div>
                <div class="monitor-bolum-sag">
                    {{ sayfa_bilgisi() }}
                    <span class="sayi-rozet">{{ dizgide|length }}</span>
                </div>
            </div>

            {% if dizgide %}
            <div class="monitor-liste" data-monitor-grup="dizgide">
                {% for k in dizgide %}{{ monitor_karti(k) }}{% endfor %}
            </div>
            {% else %}
            <div class="bos-durum monitor-bos">Şu anda dizgide kart bulunmuyor.</div>
            {% endif %}
        </article>

        <article class="monitor-bolum monitor-plana">
            <div class="monitor-bolum-baslik">
                <div>
                    <p class="ust-etiket">ÜRETİM KUYRUĞU</p>
                    <h2>PLANA ALINDI</h2>
                </div>
                <div class="monitor-bolum-sag">
                    {{ sayfa_bilgisi() }}
                    <span class="sayi-rozet">{{ plana_alindi|length }}</span>
                </div>
            </div>

            {% if plana_alindi %}
            <div class="monitor-liste" data-monitor-grup="plan">
                {% for k in plana_alindi %}{{ monitor_karti(k, plan=True) }}{% endfor %}
            </div>
            {% else %}
            <div class="bos-durum monitor-bos">Plana alınmış bekleyen kart yok.</div>
            {% endif %}
        </article>

    </section>
</div>
{% endblock %}

{% block scripts %}
<script src="{{ url_for('static', filename='js/monitor.js') }}"></script>
{% endblock %}
```

## `templates/operator.html`

```html
{% extends "base.html" %}
{% from "_makrolar.html" import arama_metni, dizgi_rozetleri, ilerleme_cubugu, kpi_satiri %}
{% block title %}Operatör · PDGM İş Takip{% endblock %}

{% macro isim_alani(onek) -%}
<div class="alan">
    <label for="{{ onek }}-isim">İşlemi yapan</label>
    <input id="{{ onek }}-isim" type="text" maxlength="80" data-islem-yapan required
           placeholder="Adınız soyadınız" autocomplete="name">
    <small>Kartta ve işlem kaydında bu ad görünür. Bu bilgisayarda hatırlanır.</small>
</div>
{%- endmacro %}

{% block content %}
<div class="sayfa-shell operator-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">OPERATÖR EKRANI</p>
            <h1>Kart İşlemleri</h1>
            <p class="soluk">Kartı dizgiye alın, üretilen adedi girin, üretim bittiğinde kartı teslim edin.</p>
        </div>
    </section>

    {{ kpi_satiri(sayac) }}

    <section class="arac-cubugu panel-kutu operator-arac-cubugu">
        <div class="arama">
            <label for="kart-ara">Kart ara</label>
            <input id="kart-ara" type="search" placeholder="Talep no, stok no, talep sahibi, PCB, dizgi sorumlusu..." autocomplete="off">
        </div>

        <div class="operator-filtre-alani">
            <div class="filtreler" role="group" aria-label="Durum filtresi">
                <button class="filtre aktif" data-filtre="AKTIF" type="button">Aktif</button>
                <button class="filtre" data-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                <button class="filtre" data-filtre="DİZGİDE" type="button">Dizgide</button>
                <button class="filtre" data-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
                <button class="filtre" data-filtre="HEPSI" type="button">Hepsi</button>
            </div>
            <div class="filtreler" role="group" aria-label="Dizgi tipi filtresi">
                <button class="filtre aktif" data-dizgi-filtre="HEPSI" type="button">Tümü</button>
                <button class="filtre" data-dizgi-filtre="MAKINE" type="button">Makine</button>
                <button class="filtre" data-dizgi-filtre="ELLE" type="button">Elle Dizgi</button>
                <button class="filtre" data-dizgi-filtre="EUM" type="button">EÜM'de Dizgi</button>
            </div>
            <div class="operator-filtre-meta">
                <span id="operator-sonuc" aria-live="polite"></span>
                <button id="operator-temizle" class="metin-buton" type="button" hidden>Filtreleri temizle</button>
            </div>
            {% if eski_teslim %}
            <p class="operator-teslim-notu">Son {{ teslim_gun }} günde teslim edilen kartlar gösteriliyor; daha eski {{ eski_teslim }} kart <a href="{{ url_for('panel') }}">Pano</a>'da.</p>
            {% endif %}
        </div>
    </section>

    <section id="operator-kartlari" class="operator-grid">
    {% for k in kartlar %}
        <article class="operator-kart {{ k.renk }}"
                 data-kart
                 data-durum="{{ k.durum }}"
                 data-dizgi-tipi="{{ k.dizgi_kod }}"
                 data-arama="{{ arama_metni(k) }}">

            <div class="operator-kart-ust">
                <div>
                    <span class="ust-etiket">{{ k.talep_no or "TALEP YOK" }}</span>
                    <h2>{{ k.stok_no or "Stok no yok" }}</h2>
                    <p>{{ k.talep_sahibi or "Talep sahibi yok" }}</p>
                </div>
                <div class="operator-rozetler">
                    <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                    {{ dizgi_rozetleri(k) }}
                    {% if k.kaynakta_yok %}<span class="durum-rozet uyari">Kaynak Excel'de yok</span>{% endif %}
                </div>
            </div>

            <div class="durum-satiri">
                <span><b>Durum:</b> {{ k.durum }}</span>
                <span><b>Excel'deki durum:</b> {{ k.kaynak_durumu or "—" }}</span>
            </div>

            <div class="bilgi-grid">
                <div><span>Toplam</span><strong>{{ k.toplam_adet }}</strong></div>
                <div><span>Tamamlanan</span><strong>{{ k.tamamlanan_adet }}</strong></div>
                <div><span>Kalan</span><strong>{{ k.kalan_adet }}</strong></div>
                <div><span>Plan Teslim</span><strong>{{ k.plan_teslim|gun }}</strong></div>
            </div>

            <div class="ilerleme">
                <div class="ilerleme-ust">
                    <span>Üretim ilerlemesi</span>
                    <strong>%{{ k.adet_yuzde }}</strong>
                </div>
                {{ ilerleme_cubugu(k) }}
            </div>

            {% if k.aciklama %}<p class="kart-not"><b>Notlar:</b><br>{{ k.aciklama }}</p>{% endif %}

            <div class="kart-ek-bilgi">
                {% if k.pcb %}<span>{{ "Malzeme/PCB" if k.dizgi_kod != "MAKINE" else "PCB" }}: {{ k.pcb }}</span>{% endif %}
                {% if k.dizgi_sorumlusu %}<span>Dizgi Sorumlusu: {{ k.dizgi_sorumlusu }}</span>{% endif %}
                {% if k.operator %}<span>Son işlem: {{ k.operator }}</span>{% endif %}
                {% if k.bitis_zamani %}<span>Üretim bitiş: {{ k.bitis_zamani|gun }}</span>{% endif %}
            </div>

            {% set kart_izinli = (oturum_rol == "admin") or
                (oturum_operator_tipi == "makine" and k.dizgi_kod == "MAKINE") or
                (oturum_operator_tipi == "elle_dizgi" and k.dizgi_kod == "ELLE") or
                (oturum_operator_tipi == "eum_dizgi" and k.dizgi_kod == "EUM") %}
            {% set baslik = (k.talep_no or "Kart") ~ " · " ~ (k.stok_no or "") %}
            <div class="kart-aksiyonlar">
                {% if not kart_izinli %}
                    <span class="durum-rozet notr"
                          title="Bu kart {{ k.dizgi_etiket }} operatörüne ait; işlem yapamazsınız.">
                        {{ k.dizgi_etiket }} operatörüne ait
                    </span>
                {% elif k.durum == "PLANA ALINDI" %}
                    <button class="buton buton-ana" type="button" data-baslat
                            data-malzeme-bekliyor="{{ '1' if k.malzeme_bekliyor else '0' }}"
                            data-id="{{ k.id }}" data-kalan="{{ k.kalan_adet }}" data-baslik="{{ baslik }}">
                        Dizgiye Al
                    </button>
                {% elif k.durum == "DİZGİDE" and k.kalan_adet > 0 %}
                    <button class="buton buton-ana" type="button" data-bitir
                            data-id="{{ k.id }}" data-kalan="{{ k.kalan_adet }}"
                            data-tamamlanan="{{ k.tamamlanan_adet }}" data-toplam="{{ k.toplam_adet }}"
                            data-baslik="{{ baslik }}">
                        Üretilen Adedi Gir
                    </button>
                {% elif k.durum == "DİZGİDE" and k.kalan_adet == 0 %}
                    <button class="buton buton-basari" type="button" data-teslim
                            data-id="{{ k.id }}" data-baslik="{{ baslik }}" data-toplam="{{ k.toplam_adet }}">
                        Teslim Et
                    </button>
                {% endif %}

                <button class="buton buton-hayalet" type="button" data-not
                        data-id="{{ k.id }}" data-not-mevcut="{{ k.aciklama or '' }}"
                        data-baslik="{{ baslik }}">Not Ekle</button>
            </div>
        </article>
    {% else %}
        <div class="bos-durum tam-satir">Görüntülenecek kart yok.</div>
    {% endfor %}
    </section>

    <div id="operator-bos" class="bos-durum" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
</div>

<dialog id="malzeme-dialog" class="modal modal-dar" aria-labelledby="malzeme-baslik">
    <div class="modal-kutu">
        <div class="modal-baslik">
            <div><p class="ust-etiket">MALZEME BEKLİYOR</p><h2 id="malzeme-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <p>Bu kart malzeme bekliyor olarak işaretlenmiş.<br>Malzemenin tedarik edildiğini onaylıyor musunuz?</p>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="button" class="buton buton-ana" id="malzeme-onayla">Malzeme geldi, devam et</button>
        </div>
    </div>
</dialog>

<dialog id="baslat-dialog" class="modal" aria-labelledby="baslat-baslik">
    <form method="dialog" class="modal-kutu" id="baslat-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">DİZGİYE AL</p><h2 id="baslat-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="baslat-id">
        <div class="alan">
            <label for="baslat-adet">Dizgiye alınacak adet</label>
            <input type="number" id="baslat-adet" min="1" required>
            <small id="baslat-kalan"></small>
        </div>
        {{ isim_alani("baslat") }}
        <div class="alan">
            <label for="baslat-not">Not <span class="soluk">(isteğe bağlı)</span></label>
            <textarea id="baslat-not" rows="3"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Dizgiye Al</button>
        </div>
    </form>
</dialog>

<dialog id="bitir-dialog" class="modal" aria-labelledby="bitir-baslik">
    <form method="dialog" class="modal-kutu" id="bitir-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">ÜRETİLEN ADET</p><h2 id="bitir-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="bitir-id">
        <div class="alan">
            <label for="bitir-adet">Bu seferde tamamlanan adet</label>
            <input type="number" id="bitir-adet" min="1" required>
            <small id="bitir-kalan"></small>
        </div>
        {{ isim_alani("bitir") }}
        <div class="alan">
            <label for="bitir-not">Not <span class="soluk">(isteğe bağlı)</span></label>
            <textarea id="bitir-not" rows="3"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>

<dialog id="teslim-dialog" class="modal" aria-labelledby="teslim-baslik">
    <form method="dialog" class="modal-kutu" id="teslim-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">TESLİM ET</p><h2 id="teslim-baslik">Kart</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="teslim-id">
        <p id="teslim-aciklama" class="onay-mesaj">Kart fiziksel olarak teslim edildi mi? Onaylarsanız kart TESLİM EDİLDİ olur ve bugünün tarihi teslim tarihi olarak yazılır.</p>
        {{ isim_alani("teslim") }}
        <div class="alan">
            <label for="teslim-not">Not <span class="soluk">(isteğe bağlı)</span></label>
            <textarea id="teslim-not" rows="2"></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-basari">Teslim Et</button>
        </div>
    </form>
</dialog>

<dialog id="not-dialog" class="modal" aria-labelledby="not-baslik">
    <form method="dialog" class="modal-kutu" id="not-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">KART NOTU</p><h2 id="not-baslik">Not Ekle</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="not-id">
        <div id="not-gecmis-kutu" class="alan" hidden>
            <span class="alan-etiket">Önceki notlar</span>
            <pre id="not-gecmis" class="not-gecmis"></pre>
        </div>
        {{ isim_alani("not") }}
        <div class="alan">
            <label for="not-metin">Yeni not</label>
            <textarea id="not-metin" rows="4" placeholder="Not metni. Tarih ve adınız otomatik eklenir." required></textarea>
        </div>
        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Ekle</button>
        </div>
    </form>
</dialog>
{% endblock %}

{% block scripts %}
<script src="{{ url_for('static', filename='js/operator.js') }}"></script>
{% endblock %}
```

## `templates/panel.html`

```html
{% extends "base.html" %}
{% from "_makrolar.html" import arama_metni, dizgi_rozetleri, ilerleme_cubugu, kpi_satiri %}
{% block title %}Pano · PDGM İş Takip{% endblock %}

{% block content %}
<div class="sayfa-shell panel-sayfa">

    <section class="sayfa-baslik sayfa-hero panel-hero">
        <div>
            <p class="ust-etiket">CANLI ÜRETİM PANOSU</p>
            <h1>İş Durumu</h1>
            <div class="panel-canli-satir">
                <span class="panel-canli" data-canli data-durum="guncel">Güncel</span>
                <span class="panel-guncelleme">Yüklendi: {{ guncelleme }} · {{ bugun }}<span data-son-kontrol></span></span>
            </div>
        </div>
        <button id="panel-yenile" class="buton buton-hayalet" type="button">Yenile</button>
    </section>

    {{ kpi_satiri(sayac) }}

    <section class="arac-cubugu panel-kutu panel-arac-cubugu">
        <div class="arama panel-arama-alani">
            <label for="panel-kart-ara">Kart ara</label>
            <input id="panel-kart-ara" type="search" placeholder="Talep no, stok no, talep sahibi, operatör, PCB, not..." autocomplete="off">
        </div>
        <div class="panel-filtre-alani">
            <div class="filtreler" role="group" aria-label="Durum filtresi">
                <button class="filtre aktif" data-panel-filtre="HEPSI" type="button">Hepsi</button>
                <button class="filtre" data-panel-filtre="DİZGİDE" type="button">Dizgide</button>
                <button class="filtre" data-panel-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                <button class="filtre" data-panel-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
            </div>
            <div class="filtreler" role="group" aria-label="Dizgi tipi filtresi">
                <button class="filtre aktif" data-panel-dizgi-filtre="HEPSI" type="button">Tümü</button>
                <button class="filtre" data-panel-dizgi-filtre="MAKINE" type="button">Makine</button>
                <button class="filtre" data-panel-dizgi-filtre="ELLE" type="button">Elle Dizgi</button>
                <button class="filtre" data-panel-dizgi-filtre="EUM" type="button">EÜM'de Dizgi</button>
            </div>
            <div class="panel-filtre-alt">
                <span id="panel-sonuc-sayisi" aria-live="polite"></span>
                <button id="panel-temizle" class="metin-buton" type="button" hidden>Filtreleri temizle</button>
            </div>
        </div>
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="DİZGİDE" id="bolum-dizgide">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">AKTİF ÜRETİM</p>
                <h2>Dizgide</h2>
                <p class="panel-aciklama">Üretimi devam eden kartlar.</p>
            </div>
            <span class="sayi-rozet">{{ dizgide|length }}</span>
        </div>

        {% if dizgide %}
        <div class="kart-listesi">
            {% for k in dizgide %}
            <article class="is-karti panel-is-karti {{ k.renk }}" data-panel-kart data-durum="DİZGİDE"
                     data-dizgi-tipi="{{ k.dizgi_kod }}" data-arama="{{ arama_metni(k) }}">
                <div class="is-karti-ust">
                    <div>
                        <strong>{{ k.talep_no or "Talep yok" }}</strong>
                        <span>{{ k.stok_no or "Stok no yok" }}</span>
                    </div>
                    <span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>
                </div>
                {% if k.dizgi_kod != 'MAKINE' or k.malzeme_bekliyor %}
                <div class="satir-alt-rozet">{{ dizgi_rozetleri(k) }}</div>
                {% endif %}
                <div class="kart-bilgiler">
                    <span><b>Sahibi:</b> {{ k.talep_sahibi or "—" }}</span>
                    <span><b>Dizgi başlangıç tarihi:</b> {{ k.plan_baslama|gun }}</span>
                    <span><b>Planlanan teslim tarihi:</b> {{ k.plan_teslim|gun }}</span>
                </div>
                <div class="ilerleme">
                    <div class="ilerleme-ust"><span>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span><strong>%{{ k.adet_yuzde }}</strong></div>
                    {{ ilerleme_cubugu(k) }}
                </div>
                {% if k.aciklama %}<p class="kart-not">{{ k.aciklama }}</p>{% endif %}
            </article>
            {% endfor %}
        </div>
        {% else %}
        <div class="bos-durum">Şu anda dizgide iş yok.</div>
        {% endif %}
    </section>

    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="PLANA ALINDI" id="bolum-plana">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">ÜRETİM SIRASI</p>
                <h2>Plana Alınan İşler</h2>
                <p class="panel-aciklama">Dizgiye alınmayı bekleyen kartlar; arama ve filtreler tüm kayıtları kapsar.</p>
            </div>
            <span class="sayi-rozet">{{ plana_alindi|length }}</span>
        </div>

        {% if plana_alindi %}
        <div class="tablo-kapsayici panel-tablo-scroll">
            <table>
                <thead><tr><th>Talep No</th><th>Stok No</th><th>Talep Sahibi</th><th>Adet</th><th>Dizgi Başlangıç T.</th><th>Planlanan Teslim T.</th><th>Değerlendirme</th></tr></thead>
                <tbody>
                {% for k in plana_alindi %}
                    <tr data-panel-kart data-durum="PLANA ALINDI"
                        data-dizgi-tipi="{{ k.dizgi_kod }}" data-arama="{{ arama_metni(k) }}">
                        <td><strong>{{ k.talep_no or "—" }}</strong></td>
                        <td>{{ k.stok_no or "—" }}</td>
                        <td>{{ k.talep_sahibi or "—" }}</td>
                        <td>{{ k.toplam_adet }}</td>
                        <td>{{ k.plan_baslama|gun }}</td>
                        <td>{{ k.plan_teslim|gun }}</td>
                        <td><span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>{{ dizgi_rozetleri(k) }}</td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
        <div class="bos-durum">Plana alınmış bekleyen kart yok.</div>
        {% endif %}
    </section>

    {% set o = donem_ozet %}
    <section class="panel-kutu panel-bolum" data-panel-bolum data-durum="TESLİM EDİLDİ" id="bolum-teslim">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">TESLİM GEÇMİŞİ</p>
                <h2>Teslim Edilenler</h2>
                <p class="panel-aciklama">Seçilen döneme göre teslim edilen kartlar ve performans özeti.</p>
            </div>
            <span class="sayi-rozet" id="panel-donem-sayi">{{ teslim_edilen|length }}</span>
        </div>

        <div class="donem-filtre-cubugu">
            <div class="filtreler" role="group" aria-label="Dönem filtresi">
                <button class="filtre aktif" data-donem-filtre="tumu" type="button">Tümü</button>
                <button class="filtre" data-donem-filtre="hafta" type="button">Bu Hafta</button>
                <button class="filtre" data-donem-filtre="ay" type="button">Bu Ay</button>
                <button class="filtre" data-donem-filtre="yil" type="button">Bu Yıl</button>
                <button class="filtre" data-donem-filtre="ozel" type="button">Özel Aralık</button>
            </div>
            <div id="donem-ozel-alan" class="donem-ozel-alan" hidden>
                <input type="text" id="donem-baslangic" data-tarih inputmode="numeric"
                       placeholder="gg.aa.yyyy" maxlength="10" aria-label="Başlangıç tarihi">
                <span aria-hidden="true">—</span>
                <input type="text" id="donem-bitis" data-tarih inputmode="numeric"
                       placeholder="gg.aa.yyyy" maxlength="10" aria-label="Bitiş tarihi">
                <button id="donem-uygula" class="buton buton-kucuk buton-ana" type="button">Uygula</button>
            </div>
        </div>

        <div class="donem-metrik-grid donem-metrik-grid-4" id="donem-metrikler" aria-live="polite">
            <div class="donem-metrik-kart">
                <span>Teslim edilen iş emri</span>
                <strong id="donem-metrik-is">{{ o.kart }}</strong>
                <small id="donem-metrik-adet-alt">{{ o.adet|sayi(0) }} adet kart teslim edildi</small>
            </div>
            <div class="donem-metrik-kart">
                <span>Zamanında teslim</span>
                <strong id="donem-metrik-zamaninda">%{{ o.zamaninda_yuzde }}</strong>
                <small id="donem-metrik-zamaninda-alt">{{ o.zamaninda }} iş emri · {{ o.zamaninda_adet|sayi(0) }} adet</small>
            </div>
            <div class="donem-metrik-kart">
                <span>Geciken teslim</span>
                <strong id="donem-metrik-gecikme">{{ o.gecikmeli }}</strong>
                <small id="donem-metrik-gecikme-alt">{{ o.gecikmeli_adet|sayi(0) }} adet</small>
            </div>
            <div class="donem-metrik-kart">
                <span>Ort. teslim sapması</span>
                <strong id="donem-metrik-sapma">{% if o.sapma_olculen %}{{ o.ort_sapma|sayi }} gün{% else %}—{% endif %}</strong>
                <small>+ geç · − erken</small>
            </div>
        </div>
        {% set olculemeyen = o.kart - o.sapma_olculen %}
        <div id="donem-veri-uyari" class="donem-veri-uyari" {% if not olculemeyen %}hidden{% endif %}>
            {%- if olculemeyen %}{{ olculemeyen }} iş emrinde plan veya gerçekleşen teslim tarihi eksik olduğu için sapma hesaplanamadı; yüzde yalnız {{ o.sapma_olculen }} iş emri üzerinden hesaplandı.{% endif -%}
        </div>

        <div class="tablo-kapsayici panel-tablo-scroll">
            <table>
                <thead><tr><th>Talep No</th><th>Stok No</th><th>Adet</th><th>Dizgi Başlangıç T.</th><th>Gerçekleşen Teslim T.</th><th>Değerlendirme</th></tr></thead>
                <tbody id="donem-tablo-govde">
                {% for k in teslim_edilen %}
                <tr data-panel-kart data-durum="TESLİM EDİLDİ"
                    data-dizgi-tipi="{{ k.dizgi_kod }}" data-arama="{{ arama_metni(k) }}">
                    <td><strong>{{ k.talep_no or "—" }}</strong></td>
                    <td>{{ k.stok_no or "—" }}</td>
                    <td>{{ k.toplam_adet }}</td>
                    <td>{{ k.plan_baslama|gun }}</td>
                    <td>{{ (k.gerceklesen_teslim or k.teslim_zamani)|gun }}</td>
                    <td><span class="durum-rozet {{ k.renk }}">{{ k.rozet }}</span>{{ dizgi_rozetleri(k, malzeme=False) }}</td>
                </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        <div id="donem-bos" class="bos-durum" {% if teslim_edilen %}hidden{% endif %}>Bu dönemde teslim edilen kart bulunmuyor.</div>
    </section>

    <div id="panel-arama-bos" class="bos-durum" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
</div>
{% endblock %}

{% block scripts %}
<script src="{{ url_for('static', filename='js/panel.js') }}"></script>
{% endblock %}
```

## `templates/yetkisiz.html`

```html
{% extends "base.html" %}
{% block title %}Yetkisiz · PDGM İş Takip{% endblock %}
{% block content %}
<div class="yetkisiz-sayfa">
    <section class="durum-sayfasi">
        <p class="ust-etiket">ERİŞİM KISITLI</p>
        <div class="durum-ikon">403</div>
        <h1>Bu sayfaya erişim yetkiniz yok.</h1>
        <p>Mevcut hesabınız bu işlemi gerçekleştirmek için gerekli role sahip değil.</p>
        <a class="buton buton-ana" href="{{ url_for('ana') }}">Ana Sayfaya Dön</a>
    </section>
</div>
{% endblock %}
```

## `templates/yonetim.html`

```html
{% extends "base.html" %}
{% from "_makrolar.html" import arama_metni, duzenle_butonu %}
{% block title %}Yönetim · PDGM İş Takip{% endblock %}

{% macro tarih_alani(id, etiket) -%}
<div class="alan">
    <label for="{{ id }}">{{ etiket }}</label>
    <input type="text" id="{{ id }}" data-tarih inputmode="numeric" placeholder="gg.aa.yyyy" maxlength="10" autocomplete="off">
</div>
{%- endmacro %}

{% block content %}
<div class="sayfa-shell yonetim-sayfa">
    <section class="sayfa-baslik sayfa-hero">
        <div>
            <p class="ust-etiket">YÖNETİCİ PANELİ</p>
            <h1>Sistem Yönetimi</h1>
            <p class="soluk">Excel aktarımı, manuel kart yönetimi, kartlar.xlsx bakımı, yedekler ve işlem geçmişi.</p>
        </div>
        <div class="baslik-aksiyon">
            <a class="buton buton-hayalet" href="{{ url_for('rapor_indir') }}">Rapor İndir</a>
            <a class="buton buton-ana" href="{{ url_for('panel') }}">Canlı Panoyu Aç</a>
        </div>
    </section>

    <section class="yonetim-grid">
        <article class="panel-kutu yonetim-yukleme-karti">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">VERİ AKTARIMI</p>
                    <h2>Plan Excel'ini Aktar</h2>
                    <p class="panel-aciklama">MAKİNE (zorunlu), ELDE DİZGİ ve EÜM sayfaları okunur; ELDE DİZGİ ve EÜM isteğe bağlıdır. Gizli satır ve sütunlar da okunur, sütunlar başlık adıyla eşlenir. Talep NO ve Kart Stok No olmayan satırlar kart sayılmaz ve önizlemede listelenir.</p>
                </div>
            </div>
            <form method="post" action="{{ url_for('yukle') }}" enctype="multipart/form-data" class="yukleme-form"
                  data-tek-gonderim data-gonderim-metni="Excel okunuyor…">
                <input type="hidden" name="onizleme" value="1">
                <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                <label class="dosya-sec">
                    <span>Excel dosyası seçin</span>
                    <small>.xlsx veya .xlsm</small>
                    <input type="file" name="dosya" accept=".xlsx,.xlsm" required>
                </label>
                <button class="buton buton-ana" type="submit">Etkiyi Önizle</button>
            </form>
            <details class="acilir yardim">
                <summary>Aktarım nasıl çalışır?</summary>
                <p>
                    Kaynak Excel ilk kurulumda kartları oluşturmak için kullanılır. Sonrasında günlük işin ana kaynağı
                    <code>data/kartlar.xlsx</code> olur. Yeniden aktarımda kaynak alanları ve geçerli DURUM Excel'den güncellenir.
                    PLANA ALINDI / HAZIR tamamlanan adedi sıfırlar; TESLİM EDİLDİ toplam adede eşitler.
                    DİZGİDE veya boş/geçersiz DURUM kısmi adedi korur. Notlar korunur.
                    Dosyada bulunmayan satırlar ve kaldırılan sayfaların kartları pasifleşir.
                </p>
            </details>
        </article>

        <article class="panel-kutu yonetim-bakim-karti">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">BAKIM</p>
                    <h2>Kayıt Dosyaları</h2>
                    <p class="panel-aciklama">Excel'e doğrudan müdahale gerektiğinde kullanılacak dosyalar.</p>
                </div>
            </div>
            <div class="buton-grup dikey">
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='kartlar') }}">Kartlar Excel</a>
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='log') }}">İşlem Logu</a>
                <a class="buton buton-hayalet" href="{{ url_for('kayit_dosyasi', hangi='yuklemeler') }}">Yükleme Geçmişi</a>
                <form method="post" action="{{ url_for('yeniden_oku') }}"
                      data-onay-baslik="Kart dosyası yeniden okunsun mu?"
                      data-onay="Diskteki kartlar.xlsx doğrulanarak yeniden okunacak. Excel'de açıksa önce kaydedip kapatın."
                      data-onay-evet="Yeniden oku">
                    <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                    <button class="buton buton-uyari tam-genislik" type="submit">Kart Dosyasını Yeniden Oku</button>
                </form>
            </div>
            <p class="yardim">Manuel Excel düzenlemesi yaparken önce dosyayı kaydedip Excel'de kapatın; ardından "Kart Dosyasını Yeniden Oku" kullanın.</p>
        </article>
    </section>

    {% if durumu_eksik %}
    <section class="panel-kutu yonetim-durum-eksik">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KONTROL GEREKİYOR</p>
                <h2>Durumu Eksik Kartlar</h2>
                <p class="panel-aciklama">Kaynak Excel'de DURUM boş veya geçersiz olduğu için Pano ve Operatör ekranında gösterilmezler.</p>
            </div>
            <span class="sayi-rozet">{{ durumu_eksik|length }}</span>
        </div>

        <div class="mini-liste yonetim-durum-eksik-liste">
            {% for k in durumu_eksik %}
            <div class="mini-satir">
                <div>
                    <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                    <span>{{ k.talep_sahibi or "Talep sahibi yok" }} · {{ k.toplam_adet }} adet</span>
                    <small>Kaynak DURUM: {{ k.excel_durum or "boş" }}</small>
                </div>
                <div class="mini-sag">
                    <span class="durum-rozet uyari">DURUMU EKSİK</span>
                    {% if k.dizgi_kod == 'ELLE' %}<span class="durum-rozet elle">Elle Dizgi</span>{% endif %}
                    {% if k.dizgi_kod == 'EUM' %}<span class="durum-rozet eum">EÜM'de Dizgi</span>{% endif %}
                    {% if k.malzeme_bekliyor %}<span class="durum-rozet uyari">Malzeme Bekliyor</span>{% endif %}
                    {{ duzenle_butonu(k, "Durum Ata", "buton-ana") }}
                </div>
            </div>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    {% if kaynakta_olmayan %}
    <section class="panel-kutu yonetim-uyari-kutu">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KAYNAK KONTROLÜ</p>
                <h2>Son Excel'de Olmayan Açık Kartlar</h2>
                <p class="panel-aciklama">Kartlar silinmez; üzerlerindeki operasyon bilgisi korunur.</p>
            </div>
            <span class="sayi-rozet">{{ kaynakta_olmayan|length }}</span>
        </div>
        <div class="mini-liste">
            {% for k in kaynakta_olmayan %}
            <div class="mini-satir">
                <div>
                    <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                    <span>{{ k.is_durumu }} · {{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span>
                </div>
                <div class="mini-sag"><span class="durum-rozet uyari">Kaynak Excel'de yok</span></div>
            </div>
            {% endfor %}
        </div>
    </section>
    {% endif %}

    <section class="yonetim-grid">
        <article class="panel-kutu">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">KART YÖNETİMİ</p>
                    <h2>Gizlenen Kartlar</h2>
                    <p class="panel-aciklama">Gizlemek silme işlemi değildir; kart kartlar.xlsx içinde kalır.</p>
                </div>
                <span class="sayi-rozet">{{ gizlenen_kartlar|length }}</span>
            </div>

            {% if gizlenen_kartlar %}
            <div class="mini-liste">
                {% for k in gizlenen_kartlar %}
                <div class="mini-satir">
                    <div>
                        <strong>{{ k.talep_no or "—" }} · {{ k.stok_no or "—" }}</strong>
                        <span>{{ k.is_durumu }} · {{ k.tamamlanan_adet }}/{{ k.toplam_adet }} adet</span>
                    </div>
                    <div class="mini-sag">
                        <span class="durum-rozet">Gizli</span>
                        <form method="post" action="{{ url_for('kart_geri_getir') }}"
                              data-onay-baslik="Kart geri getirilsin mi?"
                              data-onay="{{ k.talep_no or 'Bu kart' }} tekrar aktif listelere alınacak."
                              data-onay-evet="Geri getir">
                            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                            <input type="hidden" name="kart_id" value="{{ k.id }}">
                            <button type="submit" class="buton buton-kucuk buton-basari">Geri Getir</button>
                        </form>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Gizlenmiş kart bulunmuyor.</div>
            {% endif %}
        </article>

        <article class="panel-kutu">
            <div class="panel-baslik">
                <div>
                    <p class="ust-etiket">KURTARMA</p>
                    <h2>Yedekten Geri Yükle</h2>
                    <p class="panel-aciklama">Yalnız kart verisi geri alınır; işlem kaydı geriye sarılmaz.</p>
                </div>
            </div>

            {% if yedekler %}
            <div class="mini-liste">
                {% for y in yedekler %}
                <div class="mini-satir">
                    <div>
                        <strong>{{ y.zaman }}</strong>
                        <span>{{ y.tip }} · {{ y.etiket }}</span>
                        <small>{{ y.boyut_kb }} KB</small>
                    </div>
                    <div class="mini-sag">
                        <form method="post" action="{{ url_for('yedek_geri_yukle') }}"
                              data-onay-baslik="Yedek geri yüklensin mi?"
                              data-onay="Mevcut kart verileri {{ y.zaman }} yedeğiyle değiştirilecek. Mevcut durum önce ayrıca yedeklenir."
                              data-onay-evet="Geri yükle" data-onay-tehlike="1">
                            <input type="hidden" name="_csrf_token" value="{{ csrf_token }}">
                            <input type="hidden" name="yedek" value="{{ y.ad }}">
                            <button type="submit" class="buton buton-kucuk buton-uyari">Geri Yükle</button>
                        </form>
                    </div>
                </div>
                {% endfor %}
            </div>
            {% else %}
            <div class="bos-durum">Henüz kullanılabilir yedek bulunmuyor.</div>
            {% endif %}
        </article>
    </section>

    <section class="panel-kutu yonetim-kart-panel">
        <div class="panel-baslik">
            <div>
                <p class="ust-etiket">KART YÖNETİMİ</p>
                <h2>Kartlar</h2>
                <p class="panel-aciklama">Gizlenmemiş tüm aktif kayıtlar; durumu eksik kayıtlar dahil.</p>
            </div>

            <div class="baslik-aksiyon">
                <button id="admin-yeni-ac" class="buton buton-ana" type="button">+ Yeni Kart</button>
                <div class="arama kucuk">
                    <label class="sr-only" for="admin-kart-ara">Kart ara</label>
                    <input id="admin-kart-ara" type="search" placeholder="Talep, stok, kişi..." autocomplete="off">
                </div>
            </div>
        </div>

        <div class="admin-filtre-cubugu">
            <div class="admin-filtre-gruplari">
                <div class="filtreler" role="group" aria-label="Durum filtresi">
                    <button class="filtre aktif" data-admin-filtre="HEPSI" type="button">Hepsi</button>
                    <button class="filtre" data-admin-filtre="AKTIF" type="button">Açık İşler</button>
                    <button class="filtre" data-admin-filtre="PLANA ALINDI" type="button">Plana Alındı</button>
                    <button class="filtre" data-admin-filtre="DİZGİDE" type="button">Dizgide</button>
                    <button class="filtre" data-admin-filtre="HAZIR" type="button">Hazır</button>
                    <button class="filtre" data-admin-filtre="TESLİM EDİLDİ" type="button">Teslim Edildi</button>
                    <button class="filtre" data-admin-filtre="DURUMU EKSİK" type="button">Durumu Eksik</button>
                </div>
                <div class="filtreler" role="group" aria-label="Dizgi tipi filtresi">
                    <button class="filtre aktif" data-admin-dizgi-filtre="HEPSI" type="button">Tümü</button>
                    <button class="filtre" data-admin-dizgi-filtre="MAKINE" type="button">Makine</button>
                    <button class="filtre" data-admin-dizgi-filtre="ELLE" type="button">Elle Dizgi</button>
                    <button class="filtre" data-admin-dizgi-filtre="EUM" type="button">EÜM'de Dizgi</button>
                </div>
            </div>
            <div class="admin-filtre-meta">
                <span id="admin-sonuc" class="admin-filtre-sonuc" aria-live="polite">{{ kartlar|length }} kart</span>
                <button id="admin-temizle" class="metin-buton" type="button" hidden>Filtreleri temizle</button>
            </div>
        </div>

        <div class="tablo-kapsayici yonetim-kart-scroll" id="admin-tablo-kaydirma">
            <table id="admin-kart-tablosu">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Talep NO</th>
                        <th>Stok No</th>
                        <th>Talep Sahibi</th>
                        <th>Adet</th>
                        <th>Durum</th>
                        <th>Kaynak Durumu</th>
                        <th>Dizgi</th>
                        <th>Plan Teslim</th>
                        <th>Operatör</th>
                        <th>İşlem</th>
                    </tr>
                </thead>
                <tbody>
                {% for k in kartlar %}
                    <tr data-admin-kart
                        data-durum="{{ k.durum or 'DURUMU EKSİK' }}"
                        data-dizgi-kod="{{ k.dizgi_kod }}"
                        data-arama="{{ arama_metni(k) }} {{ k.durum or 'DURUMU EKSİK' }} {{ k.excel_durum or '' }}">
                        <td>{{ k.id }}</td>
                        <td><strong>{{ k.talep_no or "—" }}</strong></td>
                        <td>{{ k.stok_no or "—" }}</td>
                        <td>{{ k.talep_sahibi or "—" }}</td>
                        <td>{{ k.tamamlanan_adet }}/{{ k.toplam_adet }}</td>
                        <td>
                            {% if k.durum %}
                                <span class="durum-rozet {{ k.renk }}">{{ k.durum }}</span>
                            {% else %}
                                <span class="durum-rozet uyari">DURUMU EKSİK</span>
                            {% endif %}
                            {% if k.kaynakta_yok %}
                                <div class="satir-alt-rozet"><span class="durum-rozet uyari">Kaynakta yok</span></div>
                            {% endif %}
                        </td>
                        <td>{{ k.kaynak_durumu or "—" }}</td>
                        <td>
                            {% if k.dizgi_kod == 'ELLE' %}
                                <span class="durum-rozet elle">Elle Dizgi</span>
                                {% if k.dizgi_sorumlusu %}<div class="satir-alt-rozet"><small>{{ k.dizgi_sorumlusu }}</small></div>{% endif %}
                            {% elif k.dizgi_kod == 'EUM' %}
                                <span class="durum-rozet eum">EÜM'de Dizgi</span>
                            {% else %}
                                <span class="durum-rozet">Makine</span>
                            {% endif %}
                            {% if k.malzeme_bekliyor %}
                                <div class="satir-alt-rozet"><span class="durum-rozet uyari">Malzeme Bekliyor</span></div>
                            {% endif %}
                        </td>
                        <td>{{ k.plan_teslim|gun }}</td>
                        <td>{{ k.operator or "—" }}</td>
                        <td>
                            <div class="satir-aksiyon">
                                {{ duzenle_butonu(k) }}
                                <button class="buton buton-kucuk buton-tehlike" type="button" data-admin-gizle
                                        data-id="{{ k.id }}" data-talep="{{ k.talep_no or '' }}">Gizle</button>
                            </div>
                        </td>
                    </tr>
                {% else %}
                    <tr><td colspan="11" class="bos-durum">Kart bulunmuyor.</td></tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <div id="admin-kart-bos" class="bos-durum admin-kart-bos" hidden>Arama ve filtreye uyan kart bulunamadı.</div>
    </section>

    <section class="yonetim-grid yonetim-gecmis">
        <article class="panel-kutu">
            <div class="panel-baslik"><div><p class="ust-etiket">SON YÜKLEMELER</p><h2>Excel Geçmişi</h2></div></div>
            {% if yuklemeler %}
            <div class="mini-liste">
                {% for y in yuklemeler %}
                <div class="mini-satir">
                    <div><strong title="{{ y.dosya or '' }}">{{ y.dosya|yukleme_adi }}</strong><span>{{ y.zaman or "—" }} · {{ y.kullanici or "—" }}</span></div>
                    <div class="mini-sag"><strong>{{ y.satir or 0 }} satır</strong><small>{{ y.yeni or 0 }} yeni · {{ y.guncellenen or 0 }} güncel · {{ y.uyari or 0 }} uyarı</small></div>
                </div>
                {% endfor %}
            </div>
            {% else %}<div class="bos-durum">Henüz Excel yüklenmemiş.</div>{% endif %}
        </article>

        <article class="panel-kutu">
            <div class="panel-baslik"><div><p class="ust-etiket">İŞLEM KAYDI</p><h2>Son İşlemler</h2></div></div>
            {% if loglar %}
            <div class="log-liste">
                {% for l in loglar %}
                <div class="log-satir">
                    <div><strong>{{ l.islem or "İşlem" }}</strong><span>{% if l.islem_yapan and l.islem_yapan != l.kullanici %}{{ l.islem_yapan }} ({{ l.kullanici or "—" }}){% else %}{{ l.kullanici or "—" }}{% endif %} · {{ l.zaman or "—" }}</span></div>
                    <small>{% if l.talep_no %}{{ l.talep_no }}{% endif %}{% if l.stok_no %} · {{ l.stok_no }}{% endif %}{% if l.detay %} · {{ l.detay }}{% endif %}</small>
                </div>
                {% endfor %}
            </div>
            {% else %}<div class="bos-durum">Henüz işlem kaydı yok.</div>{% endif %}
        </article>
    </section>
</div>

<dialog id="admin-dialog" class="modal" aria-labelledby="admin-baslik">
    <form method="dialog" class="modal-kutu" id="admin-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">ADMİN MÜDAHALESİ</p><h2 id="admin-baslik">Kart Düzenle</h2></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>
        <input type="hidden" id="admin-id">

        <div class="iki-kolon">
            <div class="alan">
                <label for="admin-durum">Durum</label>
                <select id="admin-durum" required>
                    <option value="" disabled>Durum seçin</option>
                    <option value="HAZIR">HAZIR</option>
                    <option value="PLANA ALINDI">PLANA ALINDI</option>
                    <option value="DİZGİDE">DİZGİDE</option>
                    <option value="TESLİM EDİLDİ">TESLİM EDİLDİ</option>
                </select>
            </div>
            <div class="alan">
                <label for="admin-toplam">Toplam adet</label>
                <input type="number" id="admin-toplam" min="1" required>
            </div>
        </div>

        <div class="alan">
            <label for="admin-tamamlanan">Tamamlanan adet</label>
            <input type="number" id="admin-tamamlanan" min="0" required>
            <small id="admin-tamamlanan-yardim">PLANA ALINDI ve HAZIR tamamlanan adedi sıfırlar; TESLİM EDİLDİ toplam adede eşitler.</small>
        </div>

        <div class="alan">
            <label for="admin-plan-hafta">Plan Haftası</label>
            <input type="text" id="admin-plan-hafta" placeholder="Örn. 34. hafta (17.08 haftası)">
        </div>

        <div class="iki-kolon">
            {{ tarih_alani("admin-plan-baslama", "Dizgi Başlama Tarihi") }}
            {{ tarih_alani("admin-plan-teslim", "Planlanan Teslim Tarihi") }}
        </div>

        <div class="alan">
            <label for="admin-gerceklesen-teslim">Gerçekleşen Teslim Tarihi</label>
            <input type="text" id="admin-gerceklesen-teslim" data-tarih inputmode="numeric" placeholder="gg.aa.yyyy" maxlength="10" autocomplete="off">
            <small>Yeni bir TESLİM EDİLDİ geçişinde boş tarih bugün olur. Zaten teslim edilmiş kartta tarih uydurulmaz.</small>
        </div>

        <div class="alan"><label for="admin-not">Not</label><textarea id="admin-not" rows="4"></textarea><small>Burada mevcut not geçmişini düzenlersiniz; silinen ve eklenen satırlar işlem loguna yazılır. Toplam sınır 32.767 karakterdir; sınırı aşan işlem kaydedilmez.</small></div>

        <div class="alan">
            <label for="admin-dizgi-tipi">Dizgi Tipi</label>
            <select id="admin-dizgi-tipi">
                <option value="MAKİNE">Makine</option>
                <option value="ELLE DİZGİ">Elle Dizgi</option>
                <option value="EÜM'DE DİZGİ">EÜM'de Dizgi</option>
            </select>
        </div>
        <div class="alan" id="admin-elle-alanlari">
            <label for="admin-dizgi-sorumlusu">PDGM Dizgi Sorumlusu</label>
            <input type="text" id="admin-dizgi-sorumlusu" placeholder="Elle Dizgi kartlarda kullanılır">
        </div>
        <div class="alan alan-checkbox">
            <label><input type="checkbox" id="admin-malzeme-bekliyor"> Malzeme Bekliyor</label>
            <small>İşaretliyse operatör bu kartı dizgiye alırken önce onay vermek zorunda kalır. Bir sonraki Excel içe aktarımında kaynak veriye göre yeniden hesaplanır.</small>
        </div>

        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kaydet</button>
        </div>
    </form>
</dialog>

<dialog id="admin-yeni-dialog" class="modal" aria-labelledby="admin-yeni-baslik">
    <form method="dialog" class="modal-kutu" id="admin-yeni-form">
        <div class="modal-baslik">
            <div><p class="ust-etiket">KART YÖNETİMİ</p><h2 id="admin-yeni-baslik">Yeni Kart</h2><p class="soluk">Yeni kart PLANA ALINDI durumunda oluşturulur.</p></div>
            <button class="ikon-buton" type="button" data-dialog-kapat aria-label="Kapat">×</button>
        </div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-sira">NO / Sıra</label><input type="number" id="yeni-sira" min="1" step="1"></div>
            <div class="alan"><label for="yeni-talep-no">Talep NO *</label><input type="text" id="yeni-talep-no" required></div>
        </div>

        <div class="alan"><label for="yeni-talep-sahibi">Talep Sahibi</label><input type="text" id="yeni-talep-sahibi"></div>

        <div class="iki-kolon">
            <div class="alan"><label for="yeni-stok-no">Kart Stok No *</label><input type="text" id="yeni-stok-no" required></div>
            <div class="alan"><label for="yeni-toplam">Toplam Adet *</label><input type="number" id="yeni-toplam" min="1" required></div>
        </div>

        <div class="alan"><label for="yeni-plan-hafta">Plan Haftası</label><input type="text" id="yeni-plan-hafta" placeholder="Örn. 34. hafta (17.08 haftası)"></div>

        <div class="iki-kolon">
            {{ tarih_alani("yeni-plan-baslama", "Dizgi Başlama Tarihi") }}
            {{ tarih_alani("yeni-plan-teslim", "Planlanan Teslim Tarihi") }}
        </div>

        <div class="alan"><label for="yeni-pcb">PCB</label><input type="text" id="yeni-pcb" placeholder="Örn. HBT"></div>

        <div class="alan">
            <label for="yeni-dizgi-tipi">Dizgi Tipi</label>
            <select id="yeni-dizgi-tipi">
                <option value="MAKİNE">Makine</option>
                <option value="ELLE DİZGİ">Elle Dizgi</option>
                <option value="EÜM'DE DİZGİ">EÜM'de Dizgi</option>
            </select>
        </div>
        <div class="alan" id="yeni-elle-alanlari" hidden>
            <label for="yeni-dizgi-sorumlusu">PDGM Dizgi Sorumlusu</label>
            <input type="text" id="yeni-dizgi-sorumlusu" placeholder="Elle Dizgi kartlarda kullanılır">
        </div>

        <div class="alan"><label for="yeni-not">Not</label><textarea id="yeni-not" rows="3"></textarea></div>

        <div class="modal-aksiyon">
            <button type="button" class="buton buton-hayalet" data-dialog-kapat>Vazgeç</button>
            <button type="submit" class="buton buton-ana">Kartı Oluştur</button>
        </div>
    </form>
</dialog>
{% endblock %}

{% block scripts %}
<script src="{{ url_for('static', filename='js/yonetim.js') }}"></script>
{% endblock %}
```

## `test_verileri/mac_import_yardimci.py`

```python
#!/usr/bin/env python3
"""macOS / COM yokken test Excel'ini ayni parser + depo.import ile uygular.

Uretim yolu (Windows COM snapshot) degil; sadece ortak test icin.
Kullanim (proje kokunden):

  .venv/bin/python test_verileri/mac_import_yardimci.py test_verileri/PDGM_TEST_INPUT.xlsx
"""

from __future__ import annotations

import os
import sys

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)

import openpyxl

import depo
import excel_araclari as ex


def aktar_com_siz(dosya_yolu: str, kullanici: str = "admin"):
    wb = openpyxl.load_workbook(dosya_yolu, data_only=True)
    parsed = []
    anahtar_gruplari = {}
    uyari_sayisi = 0
    sayfa_bulundu = False

    try:
        for sayfa_adi, dizgi_tipi, zorunlu in ex.KAYNAK_SAYFALARI:
            if sayfa_adi not in wb.sheetnames:
                if zorunlu:
                    raise ex.ExcelAktarimHatasi(
                        f"'{sayfa_adi}' sayfası bulunamadı. "
                        f"Dosyadaki sayfalar: {', '.join(wb.sheetnames)}"
                    )
                continue
            sayfa_bulundu = True
            uyari_sayisi += ex._sayfa_satirlarini_coz(
                wb[sayfa_adi], sayfa_adi, dizgi_tipi, anahtar_gruplari, parsed
            )
        if not sayfa_bulundu:
            raise ex.ExcelAktarimHatasi("Excel'de işlenecek kart satırı bulunamadı.")
    finally:
        wb.close()

    return depo.excel_import_uygula(
        dosya_adi=os.path.basename(dosya_yolu),
        kullanici=kullanici,
        satirlar=parsed,
        uyari_sayisi=uyari_sayisi,
    )


def main():
    if len(sys.argv) < 2:
        print("Kullanim: python test_verileri/mac_import_yardimci.py <xlsx>")
        sys.exit(2)

    yol = os.path.abspath(sys.argv[1])
    if not os.path.isfile(yol):
        print(f"Dosya yok: {yol}")
        sys.exit(1)

    depo.kur()
    sonuc = aktar_com_siz(yol)
    print("Import OK")
    for k, v in sonuc.items():
        if k == "pasife_listesi":
            print(f"  {k}: {len(v)} ornek")
            continue
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
```

## `test_verileri/TEST_SENARYOLARI.txt`

```text
GÜNCELLEME — 19.09.2026
======================
Aşağıdaki eski manuel senaryolardaki kimlik/durum beklentileri yerine
../docs/EXCEL_SYNC.md ve otomatik tests/ paketi esas alınmalıdır.
Her kaynak satırında sabit, sayfa içinde benzersiz NO veya PDGM_ROW_ID gereklidir.
TESLİM EDİLDİ + boş tarih artık durum eksik yapmaz; teslim tarihi bilinmiyor kalır.
Geçerli Excel DURUM authoritative'dir; yalnız boş DURUM ile MALZEME TEDARİK / PDGM ÖNERİ
manuel workflow'u korur. 26.09.2026'dan beri tanınmayan DURUM (ör. BEKLEMEDE) importu durdurur.
Tarih anahtar değildir. Kaynakta olmayan kartlar operasyonlarda gösterilmez.
Teslim adedi geçerli TESLİM EDİLDİ durumuyla yeniden yüklenirse yeni toplama eşitlenir.

PDGM — Ortak Test Senaryoları (örnek Excel formatı)
===================================================
MAKİNE sayfası: Downloads/ornek_uretim_takip_random.xlsx ile aynı kolon düzeni
  Talep NO | Talep Sahibi | Kart Stok No | Kart Üretim Adet | Planlanan Başlangıç T.
  | Dizgi Başlama Tarihi | Planlanan Teslim T. | Gerçekleşen Teslim T. | DURUM | PCB | | NOT

EÜM sayfası: Downloads/EUM_Uretim_Takip.xlsx ile aynı kolon düzeni
  NO | Talep NO | Talep Sahibi | Revizyon | Kart Stok No | Kart Üretim Adet
  | Planlanan Başlangıç T. | Planlanan Teslim T. | Gerçekleşen Teslim T. | DURUM | MALZEME/PCB

Dosyalar:
  1) PDGM_TEST_INPUT.xlsx
  2) PDGM_TEST_INPUT_FAZ2_yeniden_import.xlsx
  3) PDGM_TEST_INPUT_CONFLICT_iptal.xlsx

NOT: Import için sayfa adı 'MAKİNE' olmalı (örnek dosyadaki 'ÜRETİM TAKİP' değil).
     Opsiyonel sayfalar: 'ELDE DİZGİ', 'EÜM'.


FAZ 0 — Giriş
-------------
[ ] admin / operator (makine) / elle / eum / gozlemci
    eum1 / eum123  →  operator_tipi=eum_dizgi


FAZ 1 — İlk yükleme (PDGM_TEST_INPUT.xlsx)
-----------------------------------------
Beklenen: ~34 örnek satır + birkaç edge + ELDE satırları + EÜM satırları
Uyarılar normal: DURUM boş satırlar, edge testler, çoklu parti, PDGM ÖNERİ

Öne çıkan kartlar:

Talep     | Stok         | Durum        | Ne test edilir
----------|--------------|--------------|---------------------------
1634701   | AD-5155-3484 | PLANA ALINDI | Dizgiye Al mutlu yol (200 ADET)
1674905   | AD-4040-9198 | PLANA ALINDI | İkinci plana kartı
1745738   | AD-2664-6066 | DİZGİDE      | Kısmi Adet Bitir (25, PCB=VAR)
1637435   | AD-8751-4011 | DİZGİDE      | İkinci dizgi kartı
1658369   | AD-4012-3658 | TESLİM       | Teslim listesi / özet KPI
1704286   | AD-2646-8838 | (boş)        | Admin durum ata
1658369   | AD-4012-3658 | PLANA (#tarih)| Çoklu parti (42. hafta satırı)
1999001   | AD-9999-0001 | HAZIR        | Operatörde GÖRÜNMEZ
1999002   | AD-9999-0002 | BEKLEMEDE    | Durum eksik
1999003   | AD-9999-0003 | TESLİM/tarih yok | Uyarı → durum eksik

Elle:
1900101 EL-2201-1101  DİZGİ İÇİN BEKLİYOR → PLANA
1900102 EL-2201-1102  MALZEME TEDARİK → eksik + malzeme_bekliyor
1900103 EL-2201-1103  DİZGİDE + malzeme
1900104 EL-2201-1104  DİZGİDE normal

EÜM (örnek + test; dizgi_tipi = EÜM'de Dizgi):
2773908 AD-9999-0000  TESLİM EDİLDİ          (örnek veri)
2773909 AD-9999-0001  TESLİM EDİLDİ          (örnek veri)
2773910 AD-9999-0002  TESLİM EDİLDİ          (örnek veri)
2773911 AD-9999-0003  TESLİM EDİLDİ          (örnek veri)
2800101 EU-3301-2001  ÜRETİM PLANA ALINDI → PLANA
2800102 EU-3301-2002  ÜRETİM DEVAM EDİYOR → DİZGİDE
2800103 EU-3301-2003  EMTD PLANLANDI → PLANA
2800104 EU-3301-2004  PDGM ÖNERİ → durum eksik (admin ata)
2800105 EU-3301-2005  ÜRETİM DEVAM EDİYOR → DİZGİDE
2800106 EU-3301-2006  TESLİM EDİLDİ


FAZ 1b — Operatör
-----------------
Makine hesabı:
[ ] 1634701 Dizgiye Al
[ ] 1745738 Adet Bitir (örn. 10) → not ekle (tarihli)
[ ] Elle / EÜM kartta işlem YAPAMAZ

Elle hesabı:
[ ] 1900101 / 1900104 işlem yapabilir
[ ] Makine / EÜM kartında işlem YAPAMAZ
[ ] 1900102 malzeme onayı

EÜM hesabı (eum1):
[ ] 2800101 Dizgiye Al
[ ] 2800102 Adet Bitir
[ ] Makine / Elle kartında işlem YAPAMAZ
[ ] Rozet: EÜM'de Dizgi
[ ] Pano filtresi "EÜM'de Dizgi" yalnız bu kartları gösterir


FAZ 1c — Notlar
---------------
[ ] Aynı karta iki kullanıcı not eklesin → iki tarihli satır görünsün


FAZ 2 — Yeniden import
----------------------
[ ] 1745738 PCB güncellenir, durum/adet korunur
[ ] 1999100 yeni kart eklenir
[ ] FAZ1'de olup FAZ2'de olmayanlar → kaynakta yok
[ ] Edge HAZIR (1999001) pasife düşer
[ ] EÜM 2800102 PCB → PCB-E2-GUNCELLENDI
[ ] EÜM 2800199 yeni kart eklenir
[ ] EÜM 2800106 FAZ2'de yok → kaynakta yok


FAZ 2b — Conflict
-----------------
1658369 TESLİM (6 adet) iken CONFLICT dosyası (99 ADET):
[ ] Tüm import iptal
[ ] Diğer kartlar değişmez


FAZ 3 — Yönetim / yedek / rapor
------------------------------
[ ] Durum ata (1704286)
[ ] Admin filtre: EÜM'de Dizgi
[ ] Manuel kart ekle → Dizgi Tipi = EÜM'de Dizgi
[ ] Gizle / geri getir
[ ] Yedek al / geri yükle
[ ] Rapor Excel'de Dizgi Tipi kolonu EÜM'DE DİZGİ gösterir
```

## `tools/export_codebase.py`

```python
"""Export the reviewable PDGM source tree as one Markdown document."""

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path


INCLUDED_SUFFIXES = {
    ".py", ".html", ".css", ".js", ".txt", ".json", ".toml",
    ".yaml", ".yml", ".ini", ".cfg", ".bat", ".ps1", ".sh",
}
INCLUDED_NAMES = {".gitignore", "requirements.txt"}
EXCLUDED_DIRECTORIES = {
    ".git", ".idea", ".investigation", ".pytest_cache", ".venv", "venv", "__pycache__",
    "artifact_work", "data", "node_modules", "outputs", "tests", "uploads", "yedekler",
}
LANGUAGES = {
    ".py": "python", ".html": "html", ".css": "css", ".js": "javascript",
    ".json": "json", ".ps1": "powershell", ".bat": "bat",
    ".sh": "bash", ".yaml": "yaml", ".yml": "yaml", ".toml": "toml",
}


def dahil_mi(path: Path, root: Path, output: Path) -> bool:
    """Return whether a single source/documentation file belongs in the export."""
    if path.resolve() == output:
        return False
    relative = path.relative_to(root)
    if any(part in EXCLUDED_DIRECTORIES for part in relative.parts[:-1]):
        return False
    if path.name.startswith("pdgm_codebase_export."):
        return False
    if path.name.startswith(".env"):
        return False
    return path.suffix.lower() in INCLUDED_SUFFIXES or path.name in INCLUDED_NAMES


def kod_blogu(metin: str) -> str:
    """Pick a Markdown fence that cannot be closed by source content."""
    uzunluk = max((len(parca) for parca in re.findall(r"`+", metin)), default=0)
    return "`" * max(3, uzunluk + 1)


def dil(path: Path) -> str:
    return LANGUAGES.get(path.suffix.lower(), "text")


def export_et(root: Path, output: Path) -> int:
    root = root.resolve()
    output = output.resolve()
    dosyalar = sorted(
        (path for path in root.rglob("*") if path.is_file() and dahil_mi(path, root, output)),
        key=lambda path: path.relative_to(root).as_posix().lower(),
    )

    satirlar = [
        "# PDGM Codebase Export",
        "",
        f"- Üretim zamanı: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"- Kök dizin: `{root}`",
        f"- Dahil edilen dosya sayısı: {len(dosyalar)}",
        "- Hariç tutulanlar: testler, yardımcı çalışma araçları, Markdown dokümanları, çalışma verileri, yüklemeler, yedekler, çıktı raporları, inceleme kopyaları, sanal ortamlar, önbellekler ve export dosyaları.",
        "",
        "## Dosya listesi",
        "",
    ]
    satirlar.extend(f"- `{path.relative_to(root).as_posix()}`" for path in dosyalar)
    satirlar.extend(["", "## Dosya içerikleri", ""])

    for path in dosyalar:
        relative = path.relative_to(root).as_posix()
        try:
            metin = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            satirlar.extend([f"## `{relative}`", "", "_UTF-8 olmayan/binary dosya: içerik dışa aktarılmadı._", ""])
            continue
        fence = kod_blogu(metin)
        satirlar.extend([f"## `{relative}`", "", f"{fence}{dil(path)}", metin.rstrip("\n"), fence, ""])

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(satirlar), encoding="utf-8")
    return len(dosyalar)


def main() -> None:
    varsayilan_kok = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description="PDGM kaynak ağacını tek Markdown dosyasına aktarır.")
    parser.add_argument("--root", type=Path, default=varsayilan_kok, help="Aktarılacak proje kökü")
    parser.add_argument("--output", type=Path, help="Üretilecek Markdown yolu")
    args = parser.parse_args()

    root = args.root.resolve()
    if not root.is_dir():
        parser.error(f"Kök dizin bulunamadı: {root}")
    output = (args.output or root / "pdgm_codebase_export.md").resolve()
    sayi = export_et(root, output)
    print(f"{sayi} dosya dışa aktarıldı: {output}")


if __name__ == "__main__":
    main()
```

## `tools/test_paketi_uret.py`

```python
"""Gerçek PDGM örnek Excel'inden kapsamlı elle test paketi üretir.

Kullanım (proje kökünden):
    .venv\\Scripts\\python.exe tools\\test_paketi_uret.py [hedef_klasor]

Kaynak dosyanın yapısı (sayfalar, başlıklar, gizli sütun/satırlar, kalıntı
satırlar, hücre stilleri) korunur; yeni satırlar mevcut kayıt satırlarının
stiliyle eklenir. Tarihler üretildiği güne göre hesaplanır; rozet beklentileri
(BUGÜN BAŞLAMALI, SON GÜN, SÜRE AŞILDI...) o gün için geçerlidir. Başka bir gün
test edilecekse paketi o gün yeniden üretin.

Dosyalar ve sıra için bkz. TEST_REHBERI.md (bu betik aynı klasöre yazar).
"""
from __future__ import annotations

import copy
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

KOK = Path(__file__).resolve().parents[1]
KAYNAK = KOK / "PDGM_Kart_dizgi_Talepleri_Üretim_Takvimi.xlsx"

SAHIPLER = ["Gökhan Uz", "Sarper Van", "Doğu Çan", "Ali Tek", "Ayşe Al", "Volkan Batı", "Tümer Ak",
            "Ayla Bil", "Emre Kibar", "Yiğit Alp", "Mustafa As", "Orkun Kan", "Burak Tan", "Deniz Var"]
SORUMLULAR = ["Ahmet Düzgün", "Mehmet Büyük"]


def _gun(bugun, fark):
    return datetime.combine(bugun + timedelta(days=fark), datetime.min.time()) if fark is not None else None


def _hafta_metni(gun: date) -> str:
    pazartesi = gun - timedelta(days=gun.weekday())
    return f"{pazartesi.isocalendar()[1]}. hafta ({pazartesi.day}.{pazartesi.month:02d} haftası)"


# --- Kayıt tanımları ---------------------------------------------------------
# bas/tes/ger: bugüne göre gün farkı (None = boş). adet metin veya sayı.
# "not": test rehberinde bu satırın neyi sınadığı.
MAKINE = [
    dict(no=16, talep=1900016, adet="4 ADET", bas=7, tes=10, durum="PLANA ALINDI", not_="Plana alındı, ileri tarihli"),
    dict(no=17, talep=1900017, adet="3 ADET", bas=0, tes=3, durum="PLANA ALINDI", not_="BUGÜN BAŞLAMALI rozeti"),
    dict(no=18, talep=1900018, adet="6 ADET", bas=-5, tes=2, durum="PLANA ALINDI", not_="BAŞLAMADI (+5 gün) rozeti"),
    dict(no=19, talep=1900019, adet="2 ADET", bas=14, tes=18, durum="plana alındı", not_="küçük harf DURUM normalizasyonu"),
    dict(no=20, talep=1900020, adet="10 ADET", bas=-3, tes=5, durum="DİZGİDE", not_="SÜRESİ İÇİNDE; operatör kısmi adet girecek"),
    dict(no=21, talep=1900021, adet="2 ADET", bas=-2, tes=0, durum="DİZGİDE", not_="SON GÜN rozeti"),
    dict(no=22, talep=1900022, adet="3 ADET", bas=-4, tes=1, durum="DİZGİDE", not_="SON 1 GÜN rozeti"),
    dict(no=23, talep=1900023, adet="5 ADET", bas=-10, tes=-3, durum="DİZGİDE", not_="SÜRE AŞILDI (3 gün)"),
    dict(no=24, talep=1900024, adet="12 ADET", bas=-1, tes=6, durum="DİZGİDE", not_="operatör bitirip teslim edecek"),
    dict(no=25, talep=1900025, adet="4 ADET", bas=-20, tes=-15, ger=-15, durum="TESLİM EDİLDİ", not_="ZAMANINDA TESLİM"),
    dict(no=26, talep=1900026, adet="4 ADET", bas=-20, tes=-15, ger=-12, durum="TESLİM EDİLDİ", not_="GEÇ TESLİM (+3 gün)"),
    dict(no=27, talep=1900027, adet="1 ADET", bas=-10, tes=-8, ger=-8, durum="Teslim Edildi ", not_="yazım farkı + boşluk"),
    dict(no=28, talep=1900028, adet="2 ADET", bas=-9, tes=-6, durum="TESLİM EDİLDİ", not_="teslim tarihi boş uyarısı"),
    dict(no=29, talep=1900029, adet="8 ADET", bas=None, tes=None, durum=None, not_="DURUM boş; 2. adımdan önce admin durum atar, Excel boş kalır"),
    dict(no=30, talep=1900030, adet="3 ADET", bas=5, tes=9, ger=-1, durum="PLANA ALINDI", not_="tarih dolu ama teslim değil uyarısı"),
    dict(no=31, talep=1900031, stok="AD-T031-0001", adet="3 ADET", bas=-15, tes=-10, ger=-10, durum="TESLİM EDİLDİ", not_="aynı Talep+Stok 1. parti"),
    dict(no=32, talep=1900031, stok="AD-T031-0001", adet="4 ADET", bas=-2, tes=4, durum="DİZGİDE", not_="aynı Talep+Stok 2. parti"),
    dict(no=33, talep=1900031, stok="AD-T031-0001", adet="5 ADET", bas=10, tes=14, durum="PLANA ALINDI", not_="aynı Talep+Stok 3. parti"),
    dict(no=34, talep=1900034, stok="AD-T034-0001", adet="5 ADET", bas=-2, tes=4, durum="DİZGİDE", not_="7 adet: 5 makine (bu) + 2 elle (ELDE NO 28)"),
    dict(no=35, talep=1900035, adet="1.500 ADET", bas=20, tes=30, durum="PLANA ALINDI", not_="binlik ayraçlı adet"),
    dict(no=36, talep=1900036, adet=12, bas=21, tes=25, durum="PLANA ALINDI", not_="sayısal adet"),
    dict(no=37, talep=1900037, adet="6 ADET", bas=2, tes=6, durum="PLANA ALINDI", gizli=True, not_="gizli satırdaki kayıt"),
    dict(no=38, talep=1900038, adet="2 ADET", bas=3, tes="metin", durum="PLANA ALINDI", not_="metin tarih (gg.aa.yyyy)"),
    dict(no=39, talep=1900039, adet="2 ADET", bas=-1, tes=3, durum="DİZGİDE", not_="operatör teslim edecek; Excel geride kalacak"),
    dict(no=40, talep=1900040, adet="10 ADET", bas=0, tes=9, durum="PLANA ALINDI", not_="operatör dizgiye alacak; Excel geride kalacak"),
    dict(no=41, talep=1900041, adet="2 ADET", bas=-8, tes=-4, ger=-4, durum="TESLİM EDİLDİ", not_="2. adımda Excel'de DİZGİDE'ye geri alınacak"),
    dict(no=42, talep=1900042, adet="3 ADET", bas=-1, tes=5, durum="DİZGİDE", not_="2. adımda NO başka talebe verilecek"),
    dict(no=43, talep=1900043, adet="2 ADET", bas=8, tes=12, durum="PLANA ALINDI", not_="2. adımda satır silinecek"),
    dict(no=44, talep=1900044, adet="2 ADET", bas=9, tes=13, durum="PLANA ALINDI", not_="2. adımda kalıntı bırakılarak silinecek"),
    dict(no=45, talep=1900045, adet="7 ADET", bas=-3, tes=8, durum="DİZGİDE", pcb="VAR", not_="PCB bilgisi; 2. adımda Excel'de DURUM silinecek"),
]
ELDE = [
    dict(no=21, talep=1910021, adet="2 ADET", bas=2, tes=5, durum="DİZGİ İÇİN BEKLİYOR", not_="PLANA'ya iner; 2. adımda DİZGİDE"),
    dict(no=22, talep=1910022, adet="3 ADET", bas=0, tes=2, durum="DİZGİ İÇİN BEKLİYOR", not_="elle operatörü dizgiye alacak"),
    dict(no=23, talep=1910023, adet="4 ADET", bas=-3, tes=3, durum="DİZGİDE VE MALZEME BEKLENİYOR", not_="DİZGİDE + malzeme bekliyor"),
    dict(no=24, talep=1910024, adet="2 ADET", bas=-2, tes=4, durum="DİZGİDE VE MALZEME BEKLİYOR", not_="diğer yazım"),
    dict(no=25, talep=1910025, adet="5 ADET", bas=None, tes=None, durum="MALZEME TEDARİK", not_="durumsuz + malzeme; 2. adımda PLANA"),
    dict(no=26, talep=1910026, adet="6 ADET", bas=-2, tes=3, durum="DİZGİDE", not_="elle dizgide; 2. adımda Excel'de MALZEME TEDARİK"),
    dict(no=27, talep=1910027, adet="2 ADET", bas=-14, tes=-10, ger=-10, durum="TESLİM EDİLDİ", not_="elle teslim"),
    dict(no=28, talep=1900034, stok="AD-T034-0001", adet="2 ADET", bas=-6, tes=-3, ger=-3, durum="TESLİM EDİLDİ", not_="7 adetin 2 elle kısmı (bitmiş)"),
    dict(no=29, talep=1910029, adet="4 ADET", bas=1, tes=4, durum="DİZGİ İÇİN BEKLİYOR", not_="PLANA"),
    dict(no=30, talep=1910030, adet="6 ADET", bas=-1, tes=5, durum="DİZGİDE", not_="elle operatörü kısmi adet girecek"),
    dict(no=31, talep=1910031, adet="3 ADET", bas=-12, tes=-9, ger=-6, durum="TESLİM EDİLDİ", not_="GEÇ TESLİM"),
    dict(no=32, talep=1910032, adet="2 ADET", bas=-9, tes=-2, durum="DİZGİDE", not_="SÜRE AŞILDI"),
    dict(no=33, talep=1910033, adet="1 ADET", bas=4, tes=7, durum="DİZGİ İÇİN BEKLİYOR", not_="2. adımda silinecek"),
    dict(no=34, talep=1910034, adet="2 ADET", bas=None, tes=None, durum=None, not_="DURUM boş"),
    dict(no=35, talep=1910035, adet="3 ADET", bas=-1, tes=4, durum="dizgide", not_="küçük harf"),
]
EUM = [
    dict(no=10, talep=1920010, adet="4 ADET", hafta=7, tes=11, durum="PLANA ALINDI", not_="yeni EÜM değeri; 2. adımda DİZGİDE"),
    dict(no=11, talep=1920011, adet="3 ADET", hafta=0, tes=4, durum="ÜRETİM PLANA ALINDI", not_="EÜM operatörü dizgiye alacak"),
    dict(no=12, talep=1920012, adet="2 ADET", hafta=14, tes=18, durum="üretim  plana alındı ", not_="küçük harf + fazla boşluk"),
    dict(no=13, talep=1920013, adet="5 ADET", hafta=-3, tes=4, durum="ÜRETİM DEVAM EDİYOR", not_="DİZGİDE"),
    dict(no=14, talep=1920014, adet="2 ADET", hafta=10, tes=14, durum="EMTD PLANLANDI", not_="PLANA"),
    dict(no=15, talep=1920015, adet="3 ADET", hafta=-12, tes=-8, ger=-7, durum="TESLİM EDİLDİ", not_="EÜM teslim"),
    dict(no=16, talep=1920016, adet="4 ADET", hafta=21, tes=None, durum="PDGM ÖNERİ", not_="durumsuz; 2. adımda PLANA"),
    dict(no=17, talep=1920017, adet="2 ADET", hafta=28, tes=33, durum="PLANA ALINDI", hafta_bicim="yalniz_hafta", not_="yalnız 'NN. hafta'"),
    dict(no=18, talep=1920018, adet="2 ADET", hafta="belirsiz", tes=40, durum="PLANA ALINDI", not_="okunamayan hafta metni uyarısı"),
    dict(no=19, talep=1920019, adet="1 ADET", hafta=35, tes=40, durum="PLANA ALINDI", hafta_bicim="tarih", not_="hafta sütununda tarih"),
    dict(no=20, talep=1920020, adet="3 ADET", hafta=-9, tes=-2, durum="ÜRETİM DEVAM EDİYOR", not_="SÜRE AŞILDI"),
    dict(no=21, talep=1900016, stok="AD-T016-EUM", adet="2 ADET", hafta=7, tes=12, durum="PLANA ALINDI", not_="MAKİNE NO 16 ile aynı Talep, ayrı kart"),
    dict(no=22, talep=1920022, adet="6 ADET", hafta=-2, tes=5, durum="ÜRETİM DEVAM EDİYOR", not_="EÜM dizgide"),
    dict(no=23, talep=1920023, adet="2 ADET", hafta=3, tes=8, durum="ÜRETİM PLANA ALINDI", not_="EÜM plan"),
]
YENI_MAKINE = [  # 2. adımda eklenecek
    dict(no=46, talep=1900046, adet="2 ADET", bas=6, tes=9, durum="PLANA ALINDI", not_="yeni kayıt"),
    dict(no=47, talep=1900047, adet="3 ADET", bas=-1, tes=4, durum="DİZGİDE", not_="yeni kayıt"),
    dict(no=48, talep=1900048, adet="1 ADET", bas=-6, tes=-2, ger=-2, durum="TESLİM EDİLDİ", not_="yeni kayıt"),
]


# --- Yazma yardımcıları ---------------------------------------------------------
def _stil_kopyala(ws, kaynak_satir, hedef_satir, son_sutun):
    for sutun in range(1, son_sutun + 1):
        ws.cell(hedef_satir, sutun)._style = copy.copy(ws.cell(kaynak_satir, sutun)._style)


def _makine_yaz(ws, r, k, bugun, sablon=272):
    _stil_kopyala(ws, sablon, r, 26)
    bas = _gun(bugun, k.get("bas"))
    tes = (_gun(bugun, k["tes"]) if k.get("tes") != "metin"
           else (bugun + timedelta(days=10)).strftime("%d.%m.%Y"))
    stok = k.get("stok") or f"AD-T{k['no']:03d}-0001"
    sira = k["no"]
    degerler = {
        "C": k["no"], "D": k["talep"], "E": SAHIPLER[sira % len(SAHIPLER)], "F": _gun(bugun, -30),
        "G": SORUMLULAR[sira % 2], "H": "01", "I": stok, "J": k["adet"],
        "K": f"MP-T{sira:03d}", "L": "01", "M": "HBT", "N": _gun(bugun, -60),
        # Gizli sipariş bloğundaki aynı adlı "Planlanan Teslim T." kasıtlı olarak farklı:
        # yanlış sütun okunursa tarihler hemen fark edilir.
        "O": _gun(bugun, -45), "P": _gun(bugun, -44), "Q": "MAKİNE", "R": _gun(bugun, -25), "S": _gun(bugun, -25),
        "T": _hafta_metni(bas.date()) if bas else None, "U": bas, "W": tes,
        "X": _gun(bugun, k.get("ger")), "Y": k.get("durum"), "Z": k.get("pcb"),
    }
    for harf, deger in degerler.items():
        ws[f"{harf}{r}"] = deger
    ws.row_dimensions[r].hidden = bool(k.get("gizli"))


def _elde_yaz(ws, r, k, bugun, sablon=70):
    _stil_kopyala(ws, sablon, r, 23)
    stok = k.get("stok") or f"EL-T{k['no']:03d}-0001"
    degerler = {
        "B": k["no"], "C": k["talep"], "D": SAHIPLER[(k["no"] + 3) % len(SAHIPLER)], "E": _gun(bugun, -30),
        "F": SORUMLULAR[k["no"] % 2], "G": "01", "H": stok, "I": k["adet"],
        "L": "HBT", "M": _gun(bugun, -60), "N": _gun(bugun, -45), "O": _gun(bugun, -44), "P": "MANUEL",
        "Q": _gun(bugun, -25), "R": _gun(bugun, -20),
        "S": _gun(bugun, k.get("bas")), "T": _gun(bugun, k.get("tes")), "U": _gun(bugun, k.get("ger")),
        "V": k.get("durum"), "W": k.get("malzeme"),
    }
    for harf, deger in degerler.items():
        ws[f"{harf}{r}"] = deger
    ws.row_dimensions[r].hidden = False


def _eum_yaz(ws, r, k, bugun, sablon=244):
    _stil_kopyala(ws, sablon, r, 24)
    stok = k.get("stok") or f"EU-T{k['no']:03d}-0001"
    if k["hafta"] == "belirsiz":
        hafta = "belirsiz"
    else:
        gun = bugun + timedelta(days=k["hafta"])
        bicim = k.get("hafta_bicim")
        if bicim == "yalniz_hafta":
            hafta = f"{gun.isocalendar()[1]}. hafta"
        elif bicim == "tarih":
            hafta = gun.strftime("%d.%m.%Y")
        else:
            hafta = _hafta_metni(gun)
    degerler = {
        "B": k["no"], "C": k["talep"], "D": SAHIPLER[(k["no"] + 7) % len(SAHIPLER)], "E": _gun(bugun, -30),
        "F": SORUMLULAR[k["no"] % 2], "G": "01", "H": stok, "I": k["adet"], "L": "YURTDIŞI",
        "M": _gun(bugun, -60), "N": _gun(bugun, -45), "O": _gun(bugun, -44), "P": "ÜRETİMDE",
        "Q": _gun(bugun, -20), "R": _gun(bugun, -18), "S": _gun(bugun, -15),
        "T": hafta, "U": _gun(bugun, k.get("tes")), "V": _gun(bugun, k.get("ger")),
        "W": k.get("durum"), "X": k.get("malzeme"),
    }
    for harf, deger in degerler.items():
        ws[f"{harf}{r}"] = deger
    ws.row_dimensions[r].hidden = False


def _satir_bul(ws, no_sutunu, no, baslangic):
    for r in range(baslangic, ws.max_row + 1):
        if ws[f"{no_sutunu}{r}"].value == no and ws[f"{'D' if no_sutunu == 'C' else 'C'}{r}"].value:
            return r
    raise KeyError(f"{ws.title} NO {no}")


# --- Dosya setleri ---------------------------------------------------------------
def ilk_yukleme(bugun):
    wb = openpyxl.load_workbook(KAYNAK)
    ws = wb["MAKİNE"]
    r = 279
    for i, k in enumerate(MAKINE):
        _makine_yaz(ws, r, k, bugun)
        r += 1
        if i == 9:          # tamamen boş satır
            r += 1
        if i == 19:         # yalnız boşluk içeren satır ve gizli boş satır
            for harf in "DEIJY":
                ws[f"{harf}{r}"] = "   "
            ws.row_dimensions[r + 1].hidden = True
            r += 2
    ws[f"I{r + 1}"] = "AD-KALINTI-0001"        # görünür kalıntı: yalnız stok
    ws[f"C{r + 2}"], ws[f"G{r + 2}"] = 99, "Ahmet Düzgün"   # kalıntı: NO + sorumlu

    ws = wb["ELDE DİZGİ"]
    for i, k in enumerate(ELDE):
        _elde_yaz(ws, 73 + i, k, bugun)

    ws = wb["EÜM"]
    for i, k in enumerate(EUM):
        _eum_yaz(ws, 246 + i, k, bugun)
    return wb


def guncelleme(bugun):
    """2. adım: tarih/durum/adet değişiklikleri, yeni/silinen satırlar, NO'nun başka talebe
    verilmesi, Excel'de geri alma, Excel'de DURUM'un boşaltılması, gizli DURUM sütunu,
    görünür yinelenen başlık."""
    wb = ilk_yukleme(bugun)
    ws = wb["MAKİNE"]
    bul = lambda no: _satir_bul(ws, "C", no, 205)
    ws[f"W{bul(16)}"] = _gun(bugun, 12)                                   # tarih değişikliği
    ws[f"Y{bul(18)}"] = "DİZGİDE"                                         # Excel ilerletti
    ws[f"W{bul(20)}"] = _gun(bugun, 7)                                    # DİZGİDE kalır, tarih değişti
    ws[f"Y{bul(23)}"], ws[f"X{bul(23)}"] = "TESLİM EDİLDİ", _gun(bugun, -1)
    ws[f"J{bul(36)}"] = 15                                                 # adet değişikliği
    ws[f"Y{bul(37)}"] = "DİZGİDE"                                         # gizli satırda güncelleme
    ws[f"Y{bul(41)}"], ws[f"X{bul(41)}"] = "DİZGİDE", None                 # Excel'de geri alma
    ws[f"Y{bul(45)}"] = None                                               # Excel'de DURUM silindi
    r42 = bul(42)                                                         # NO 42 yeni talebe verildi
    ws[f"D{r42}"], ws[f"I{r42}"], ws[f"J{r42}"], ws[f"Y{r42}"] = 1909042, "AD-T042-YENI", "5 ADET", "PLANA ALINDI"
    ws[f"U{r42}"], ws[f"W{r42}"] = _gun(bugun, 3), _gun(bugun, 8)
    for hucre in ws[bul(43)]:                                            # satır tamamen silindi
        hucre.value = None
    r44 = bul(44)                                                         # kalıntı bırakılarak silindi
    for harf in "DEFHJKLMNOPQRSTUWXY":
        ws[f"{harf}{r44}"] = None
    son = max(r for r in range(279, ws.max_row + 1) if ws[f"C{r}"].value)
    for i, k in enumerate(YENI_MAKINE):
        _makine_yaz(ws, son + 3 + i, k, bugun)
    ws.column_dimensions["Y"].hidden = True       # DURUM sütunu gizli: yine okunmalı
    ws.column_dimensions["O"].hidden = False      # aynı adlı ikinci "Planlanan Teslim T." görünür

    ws = wb["ELDE DİZGİ"]
    bul = lambda no: _satir_bul(ws, "B", no, 3)
    ws[f"V{bul(21)}"] = "DİZGİDE"
    ws[f"V{bul(25)}"], ws[f"S{bul(25)}"], ws[f"T{bul(25)}"] = "DİZGİ İÇİN BEKLİYOR", _gun(bugun, 3), _gun(bugun, 6)
    ws[f"V{bul(26)}"] = "MALZEME TEDARİK"                                  # iş akışı durumu olmayan metin
    for hucre in ws[bul(33)]:
        hucre.value = None

    ws = wb["EÜM"]
    bul = lambda no: _satir_bul(ws, "B", no, 3)
    ws[f"W{bul(10)}"] = "ÜRETİM DEVAM EDİYOR"
    ws[f"W{bul(16)}"], ws[f"U{bul(16)}"] = "PLANA ALINDI", _gun(bugun, 26)
    _eum_yaz(ws, 246 + len(EUM), dict(no=26, talep=1920026, adet="3 ADET", hafta=5, tes=10,
                                      durum="ÜRETİM PLANA ALINDI"), bugun)
    return wb


def excel_guncel(bugun):
    """4. adım: Excel operatörün uygulamada yaptıklarına yetişti; karar gerekmez."""
    wb = guncelleme(bugun)
    ws = wb["MAKİNE"]
    bul = lambda no: _satir_bul(ws, "C", no, 205)
    ws[f"Y{bul(24)}"], ws[f"X{bul(24)}"] = "TESLİM EDİLDİ", _gun(bugun, 0)
    ws[f"Y{bul(39)}"], ws[f"X{bul(39)}"] = "TESLİM EDİLDİ", _gun(bugun, 0)
    ws[f"Y{bul(40)}"] = "DİZGİDE"
    ws[f"Y{bul(29)}"] = "PLANA ALINDI"          # admin'in atadığı durum Excel'e de yazıldı
    wb["ELDE DİZGİ"][f"V{_satir_bul(wb['ELDE DİZGİ'], 'B', 22, 3)}"] = "DİZGİDE"
    wb["EÜM"][f"W{_satir_bul(wb['EÜM'], 'B', 11, 3)}"] = "ÜRETİM DEVAM EDİYOR"
    return wb


def hatali_dosyalar(bugun):
    """Reddedilmesi gereken (H) ve uyarı gösterip vazgeçilecek (U) dosyalar."""
    def temel():
        return excel_guncel(bugun)

    def makine_satiri(wb, no):
        return _satir_bul(wb["MAKİNE"], "C", no, 205)

    dosyalar = {}
    wb = temel(); wb["MAKİNE"][f"D{makine_satiri(wb, 16)}"] = None
    dosyalar["H01_EKSIK_TALEP_NO"] = wb
    wb = temel(); wb["MAKİNE"][f"Y{makine_satiri(wb, 16)}"] = "TESLİM EDİLDİ (KISMİ)"
    dosyalar["H02_TANINMAYAN_DURUM"] = wb
    wb = temel(); wb["MAKİNE"][f"C{makine_satiri(wb, 17)}"] = 16
    dosyalar["H03_TEKRARLANAN_NO"] = wb
    wb = temel(); wb["MAKİNE"][f"J{makine_satiri(wb, 16)}"] = "on adet"
    dosyalar["H04_GECERSIZ_ADET"] = wb
    wb = temel(); wb["MAKİNE"][f"W{makine_satiri(wb, 16)}"] = "31.02.2026"
    dosyalar["H05_GECERSIZ_TARIH"] = wb
    wb = temel(); r = makine_satiri(wb, 16); wb["MAKİNE"][f"U{r}"] = _gun(bugun, 30)
    dosyalar["H06_BASLAMA_TESLIMDEN_SONRA"] = wb
    wb = temel()
    wb["MAKİNE"][f"D{makine_satiri(wb, 16)}"] = None
    wb["ELDE DİZGİ"][f"I{_satir_bul(wb['ELDE DİZGİ'], 'B', 21, 3)}"] = "on adet"
    wb["EÜM"][f"W{_satir_bul(wb['EÜM'], 'B', 10, 3)}"] = "Bitti"
    dosyalar["H07_UC_HATA_BIRDEN"] = wb
    wb = temel(); wb["MAKİNE"].title = "URETIM"
    dosyalar["H08_MAKINE_SAYFASI_YOK"] = wb
    wb = temel(); wb["MAKİNE"]["V4"] = "NO"; wb["MAKİNE"].column_dimensions["V"].hidden = False
    dosyalar["H09_IKI_NO_SUTUNU"] = wb
    wb = temel(); del wb["EÜM"]
    dosyalar["U01_EUM_SAYFASI_YOK"] = wb
    wb = temel(); ws = wb["ELDE DİZGİ"]
    for r in range(3, ws.max_row + 1):
        for hucre in ws[r]:
            hucre.value = None
    dosyalar["U02_ELDE_SAYFASI_BOSALTILDI"] = wb
    return dosyalar


def uret(hedef: Path, bugun: date | None = None) -> dict[str, Path]:
    bugun = bugun or date.today()
    hedef.mkdir(parents=True, exist_ok=True)
    sirali = {
        "01_ILK_YUKLEME": ilk_yukleme(bugun),
        "02_GUNCELLEME": guncelleme(bugun),
        "03_AYNI_DOSYA_TEKRAR": guncelleme(bugun),
        "04_EXCEL_YETISTI": excel_guncel(bugun),
        **hatali_dosyalar(bugun),
    }
    yollar = {}
    for ad, wb in sirali.items():
        yol = hedef / f"{ad}.xlsx"
        wb.save(yol)
        wb.close()
        yollar[ad] = yol
    return yollar


if __name__ == "__main__":
    hedef = Path(sys.argv[1]) if len(sys.argv) > 1 else KOK / "outputs" / f"pdgm-test-paketi-{date.today():%Y%m%d}"
    for ad, yol in uret(hedef).items():
        print(f"{ad:32} {yol}")
```

## `yedek_disari_kopyala.bat`

```bat
@echo off
REM ============================================================
REM  PDGM - Gecelik yedek kopyalama
REM  Task Scheduler ile her gece calistirin.
REM  HEDEF yolunu kendi ag paylasiminizla degistirin.
REM  Robocopy: 0-7 basari/uyari, 8+ hata.
REM ============================================================

set KAYNAK=%~dp0data
set HEDEF=C:\Users\ulasarpkocak\Desktop\pdgm_arayuz_yedek

if not exist "%HEDEF%" (
    echo HATA: Hedef erisilemiyor: %HEDEF%
    exit /b 1
)

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set BUGUN=%%I

robocopy "%KAYNAK%" "%HEDEF%\%BUGUN%" /E /R:2 /W:5 /NFL /NDL /LOG+:"%HEDEF%\robocopy.log"
set RC=%ERRORLEVEL%

powershell -NoProfile -Command ^
  "Get-ChildItem -Path '%HEDEF%' -Directory | Where-Object { $_.Name -match '^\d{8}$' -and $_.LastWriteTime -lt (Get-Date).AddDays(-60) } | Remove-Item -Recurse -Force"

if %RC% GEQ 8 (
    echo HATA: robocopy basarisiz, kod=%RC%
    exit /b %RC%
)

exit /b 0
```
