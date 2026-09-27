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
# (mevcut kartta iş akışı korunur). Bunların dışındaki tanınmayan her DURUM
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
        "Yönetim ekranındaki \"Durumu eksik kartlar\" bölümünden durum atanabilir. Kart zaten "
        "sistemdeyse mevcut iş akışı ve adetleri korunur.",
    ),
    "durum_durumsuz": (
        "DURUM \"{deger}\": iş akışı durumu atanmaz",
        "Bu aşama (malzeme tedariği, PDGM önerisi) bilerek PLANA ALINDI / DİZGİDE / TESLİM EDİLDİ'ye "
        "eşlenmez; sonuç DURUM boş ile aynıdır. Excel'deki metin kartta kaynak durumu olarak saklanır.",
    ),
    "teslim_tarihsiz": (
        "TESLİM EDİLDİ ama Gerçekleşen Teslim T. boş",
        "Kart teslim edilmiş sayılır. Teslim tarihi bilinmediği için zamanında/geç teslim "
        "hesabına girmez; sistem tarih uydurmaz.",
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
                uyar("durum_bos", "DURUM boş; yeni kartın durumu boş kalır, mevcut kartın iş akışı korunur.")
            else:
                uyar("durum_durumsuz",
                     f"DURUM '{str(excel_durum_raw).strip()}' bilerek bir iş akışı durumuna eşlenmez; "
                     "yeni kartın durumu boş kalır, mevcut kartın iş akışı korunur.",
                     ek=str(excel_durum_raw).strip())

        if ilk_durum == depo.TESLIM_EDILDI and not gerceklesen:
            uyar("teslim_tarihsiz", "TESLİM EDİLDİ ama Gerçekleşen Teslim T. boş; teslim tarihi bilinmiyor olarak kalır.")
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
                   tamamlanan_sifirla=None, gerileme_secimleri=None):
    with _import_kilidi:
        return _excelden_aktar(dosya_yolu, kullanici, onizleme, beklenen_surum, beklenen_kaynak_surum,
                               tamamlanan_sifirla, gerileme_secimleri)


def _sayfa_bul(wb, sayfa_adi):
    """COM snapshot sayfaları kanonik adla yazar; doğrudan okunan dosyada adlar serbesttir."""
    hedef = _sadelestir(sayfa_adi)
    return next((wb[ad] for ad in wb.sheetnames if _sadelestir(ad) == hedef), None)


def _excelden_aktar(dosya_yolu, kullanici, onizleme=False, beklenen_surum=None, beklenen_kaynak_surum=None,
                    tamamlanan_sifirla=None, gerileme_secimleri=None):
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
