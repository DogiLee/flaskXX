"""PDGM kaynak Excel importu ve rapor üretimi.

Kaynak import kuralları:
- MAKİNE sayfası (zorunlu) + ELDE DİZGİ sayfası (opsiyonel) + EÜM sayfası
  (opsiyonel) kullanılır. Sayfalar TEK import işleminde birlikte işlenir:
  aksi halde bir sayfayı import etmek diğer sayfadan gelen kartları
  "kaynakta yok" (pasif) yapardı.
- Hidden kolonlar alınmaz; hidden satırlar korunur.
- Microsoft Excel COM ile values-only snapshot oluşturulur.
- External link güncellemesi ve full recalculation yapılmaz.
- Talep NO + Kart Stok No olmayan satırlar kart sayılmaz.
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
import io
import os
import re
import tempfile
import threading
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
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
    pass


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
    metin = str(deger or "").translate(TURKCE_HARFLER).upper()
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
    # EÜM sayfası durumları
    "URETIM DEVAM EDIYOR": depo.DIZGIDE,
    "URETIM PLANA ALINDI": depo.PLANA_ALINDI,
    "EMTD PLANLANDI": depo.PLANA_ALINDI,
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
            # Kalıcı kimlik gizli olsa da snapshot'tan düşürülmez.
            kimlik_basliklari = {"NO", "PDGM ROW ID"}
            if not any(
                _sadelestir(kaynak_ws.Cells(r, kaynak_sutun).Value2) in kimlik_basliklari
                for r in range(ilk_satir, min(son_satir, 10) + 1)
            ):
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


def _hucre_al(satir, kolonlar, alan):
    index = kolonlar.get(alan)
    if index is None or index >= len(satir):
        return None
    deger = satir[index]
    return None if deger == "" else deger


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


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

def _sayfa_satirlarini_coz(ws, sayfa_adi, dizgi_tipi, anahtar_gruplari, parsed_liste):
    """Kaynak satırlarını doğrular; sayfa kapsamında kalıcı kimlik üretir."""
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
            if (talep_no or stok_no or _hucre_al(satir, kolonlar, "sira") is not None
                    or _hucre_al(satir, kolonlar, "source_row_id") is not None):
                raise ExcelAktarimHatasi(
                    f"'{sayfa_adi}' satır {excel_satir_no}: Talep NO ve Kart Stok No gerekli; "
                    "eksik satır silinmiş kabul edilmedi."
                )
            continue

        try:
            toplam_adet = adet_coz(_hucre_al(satir, kolonlar, "adet_metin"))
        except ExcelAktarimHatasi as exc:
            raise ExcelAktarimHatasi(f"'{sayfa_adi}' sayfası, satır {excel_satir_no}: {exc}") from exc

        plan_baslama_ham = _hucre_al(satir, kolonlar, "plan_baslama")
        plan_teslim_ham = _hucre_al(satir, kolonlar, "plan_teslim")
        gerceklesen_ham = _hucre_al(satir, kolonlar, "gerceklesen_teslim")

        plan_baslama = _tarih_coz_ve_dogrula(
            plan_baslama_ham, "Dizgi Başlama Tarihi", excel_satir_no, ws.parent.epoch
        )
        plan_teslim = _tarih_coz_ve_dogrula(
            plan_teslim_ham, "Planlanan Teslim Tarihi", excel_satir_no, ws.parent.epoch
        )
        gerceklesen = _tarih_coz_ve_dogrula(
            gerceklesen_ham, "Gerçekleşen Teslim Tarihi", excel_satir_no, ws.parent.epoch
        )

        if dizgi_tipi == depo.DIZGI_TIPI_EUM:
            planlanan = _hucre_al(satir, kolonlar, "eum_planlanan_tarih")
            if "eum_planlanan_tarih" not in kolonlar:
                planlanan = _hucre_al(satir, kolonlar, "plan_hafta")
            if planlanan not in (None, ""):
                tarih = _tarih_coz_ve_dogrula(
                    planlanan, "EÜM planlanan tarih", excel_satir_no, ws.parent.epoch
                )
                gun = date.fromisoformat(tarih)
                plan_baslama = (gun - timedelta(days=gun.weekday())).isoformat()
            elif "eum_planlanan_tarih" in kolonlar or "plan_hafta" in kolonlar:
                plan_baslama = None

        if plan_baslama and plan_teslim and plan_baslama > plan_teslim:
            raise ExcelAktarimHatasi(
                f"'{sayfa_adi}' sayfası, satır {excel_satir_no}: Dizgi Başlama Tarihi ({plan_baslama}) "
                f"Planlanan Teslim Tarihinden ({plan_teslim}) sonra olamaz."
            )

        no_index = basliklar.index("NO") if "NO" in basliklar else None
        no = satir[no_index] if no_index is not None and no_index < len(satir) else None
        source_row_id = _kaynak_satir_kimligi(
            no, _hucre_al(satir, kolonlar, "source_row_id"), sayfa_adi, excel_satir_no
        )
        source_sheet = depo.dizgi_kodu(dizgi_tipi)
        anahtar = depo.kaynak_anahtari(source_sheet, source_row_id)
        if anahtar in anahtar_gruplari:
            raise ExcelAktarimHatasi(
                f"'{sayfa_adi}' satır {excel_satir_no}: Tekrarlanan kaynak kimliği {source_row_id}."
            )
        anahtar_gruplari[anahtar] = excel_satir_no

        excel_durum_raw = _hucre_al(satir, kolonlar, "excel_durum")
        ilk_durum, durum_uyarisi = durum_coz(excel_durum_raw)
        if durum_uyarisi:
            uyari_sayisi += 1

        if (ilk_durum == depo.TESLIM_EDILDI and not gerceklesen):
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

        if not sayfa_bulundu:
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
