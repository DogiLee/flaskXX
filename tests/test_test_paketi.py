"""Elle test paketi (tools/test_paketi_uret.py) ve TEST_REHBERI.md'deki beklenen sonuçlar.

Rehberdeki her sayı bu testte doğrulanır: paket bugünün tarihiyle üretilir, adımlar
rehberdeki sırayla (operatör işlemleri dahil) geçici bir veritabanında uygulanır.
PDGM_TEST_EXCEL_COM=1 ile ilk iki adım gerçek Microsoft Excel üzerinden de koşar.
"""
import os
import shutil
import sys
import unittest
from collections import Counter
from datetime import date, timedelta
from unittest.mock import patch

import depo
import excel_araclari as ex
from tests import test_excel_sync as fixtures
from tests.test_excel_sync import ROOT

sys.path.insert(0, str(ROOT / "tools"))
import test_paketi_uret as paket_uret  # noqa: E402


def _rapor(sonuc):
    r = sonuc["kaynak_raporu"]
    return ({s["kod"]: (s["kayit"], len(s["atlanan"])) for s in r["sayfalar"]},
            {g["baslik"]: g["adet"] for g in r["uyari_gruplari"]})


@unittest.skipUnless(paket_uret.KAYNAK.exists(), "Gerçek PDGM Excel dosyası proje kökünde yok")
class TestPaketiTests(unittest.TestCase):
    setUp = fixtures.SyncTests.setUp

    def yukle(self, yol, com=False, **kw):
        if com:
            return ex.excelden_aktar(str(yol), "admin", **kw)

        def snapshot(_):
            hedef = self.root / "snapshot.xlsx"; shutil.copy2(yol, hedef); return str(hedef)
        with patch.object(ex, "excel_deger_snapshot_olustur", snapshot):
            return ex.excelden_aktar(str(yol), "admin", **kw)

    def kid(self, kod, no):
        return next(k["id"] for k in depo._kartlar if k["source_sheet"] == kod and k["source_row_id"] == f"NO:{no}")

    def kart(self, kod, no):
        return depo.kart_getir(self.kid(kod, no))

    def operator_islemleri(self):
        """TEST_REHBERI.md 2. adımdaki operatör işlemleri."""
        depo.kart_bitir(self.kid("MAKINE", 20), 3, "makine1", "operator", "makine")
        for no, adet in ((24, 12), (39, 2)):
            depo.kart_bitir(self.kid("MAKINE", no), adet, "makine1", "operator", "makine")
            depo.kart_teslim_et(self.kid("MAKINE", no), "makine1", "operator", "makine")
        depo.kart_baslat(self.kid("MAKINE", 40), 10, "makine1", "operator", "makine")
        depo.kart_bitir(self.kid("MAKINE", 40), 4, "makine1", "operator", "makine")
        depo.kart_not_guncelle(self.kid("MAKINE", 42), "Eski talebin notu", "makine1", "operator")
        depo.kart_baslat(self.kid("ELLE", 22), 3, "elle1", "operator", "elle_dizgi")
        depo.kart_bitir(self.kid("ELLE", 30), 2, "elle1", "operator", "elle_dizgi")
        depo.kart_baslat(self.kid("EUM", 11), 3, "eum1", "operator", "eum_dizgi")
        depo.kart_bitir(self.kid("EUM", 11), 1, "eum1", "operator", "eum_dizgi")

    def test_rehberdeki_akis(self):
        paket = paket_uret.uret(self.root / "paket")

        # 1. adım: ilk yükleme
        sonuc = self.yukle(paket["01_ILK_YUKLEME"])
        sayfalar, uyarilar = _rapor(sonuc)
        self.assertEqual((sonuc["yeni"], sonuc["uyari"]), (103, 16))
        self.assertEqual(sayfalar, {"MAKINE": (45, 54), "ELLE": (35, 0), "EUM": (23, 4)})
        self.assertEqual(uyarilar, {"DURUM boş": 6, "TESLİM EDİLDİ ama Gerçekleşen Teslim T. boş": 1,
                                    "Gerçekleşen Teslim T. dolu ama DURUM teslim değil": 1,
                                    'DURUM "MALZEME TEDARİK": iş akışı durumu atanmaz': 4,
                                    'DURUM "PDGM ÖNERİ": iş akışı durumu atanmaz': 3,
                                    "Planlanan başlangıç tam çözülemedi": 1})
        gorunen = depo.kartlari_getir()
        self.assertEqual(len(gorunen), 90)
        self.assertEqual(Counter((k["dizgi_kod"], k["durum"]) for k in gorunen if k["dizgi_kod"] == "MAKINE"),
                         Counter({("MAKINE", "PLANA ALINDI"): 14, ("MAKINE", "DİZGİDE"): 12, ("MAKINE", "TESLİM EDİLDİ"): 14}))
        rozetler = {no: self.kart("MAKINE", no)["rozet"] for no in (17, 18, 20, 21, 22, 23, 25, 26, 28)}
        self.assertEqual(rozetler, {17: "BUGÜN BAŞLAMALI", 18: "BAŞLAMADI (+5 gün)", 20: "PLANINDA (5 gün var)",
                                    21: "SON GÜN", 22: "SON 1 GÜN", 23: "SÜRE AŞILDI (3 gün)",
                                    25: "ZAMANINDA TESLİM", 26: "GEÇ TESLİM (+3 gün)", 28: "TESLİM EDİLDİ"})
        self.assertEqual((self.kart("MAKINE", 35)["toplam_adet"], self.kart("MAKINE", 36)["toplam_adet"]), (1500, 12))
        metin_tarih = (date.today() + timedelta(days=10)).isoformat()
        self.assertEqual(self.kart("MAKINE", 38)["plan_teslim"], metin_tarih)   # "gg.aa.yyyy" metni okundu
        self.assertEqual(self.kart("MAKINE", 37)["durum"], "PLANA ALINDI")      # gizli satırdaki kayıt
        self.assertTrue(self.kart("ELLE", 23)["malzeme_bekliyor"] and self.kart("ELLE", 24)["malzeme_bekliyor"])
        self.assertEqual([self.kart(k, n)["durum"] for k, n in (("MAKINE", 34), ("ELLE", 28))], ["DİZGİDE", "TESLİM EDİLDİ"])
        self.assertIsNone(self.kart("EUM", 18)["plan_baslama"])                  # "belirsiz" hafta metni
        self.assertNotEqual(self.kid("MAKINE", 16), self.kid("EUM", 21))         # aynı Talep, ayrı kart

        # 2. adım: operatör işlemleri + güncelleme önizlemesi ve onay
        self.operator_islemleri()
        onizleme = self.yukle(paket["02_GUNCELLEME"], onizleme=True)
        self.assertEqual((onizleme["yeni"], onizleme["guncellenen"], onizleme["pasife_alinan"],
                          onizleme["ayrilan"], onizleme["gerileme"]), (5, 11, 3, 1, 6))
        kararlar = {(d["sayfa"], d["sira"]): d["gerileme"]["onerilen"]
                    for d in onizleme["degisiklikler"] if d["tur"] == "gerileme"}
        self.assertEqual(kararlar, {("MAKINE", 24): "TESLİM EDİLDİ", ("MAKINE", 39): "TESLİM EDİLDİ",
                                    ("MAKINE", 40): "DİZGİDE", ("MAKINE", 41): "DİZGİDE",
                                    ("ELLE", 22): "DİZGİDE", ("EUM", 11): "DİZGİDE"})
        turler = Counter(d["tur"] for d in onizleme["degisiklikler"])
        self.assertEqual((turler["pasif"], turler["ayrildi"], turler["yeni"]), (3, 1, 5))
        sifirlanabilir = {(k["source_sheet"], k["source_row_id"]) for k in depo._kartlar
                          if k["id"] in onizleme["sifirlanabilir"]}
        self.assertIn(("MAKINE", "NO:20"), sifirlanabilir)
        secim = {self.kid(kod, no): durum for (kod, no), durum in kararlar.items()}
        self.yukle(paket["02_GUNCELLEME"], gerileme_secimleri=secim,
                   tamamlanan_sifirla=[self.kid("MAKINE", 20), self.kid("MAKINE", 41)])
        beklenen = {("MAKINE", 16): ("PLANA ALINDI", 0), ("MAKINE", 18): ("DİZGİDE", 0), ("MAKINE", 20): ("DİZGİDE", 0),
                    ("MAKINE", 23): ("TESLİM EDİLDİ", 5), ("MAKINE", 24): ("TESLİM EDİLDİ", 12),
                    ("MAKINE", 36): ("PLANA ALINDI", 0), ("MAKINE", 37): ("DİZGİDE", 0),
                    ("MAKINE", 39): ("TESLİM EDİLDİ", 2), ("MAKINE", 40): ("DİZGİDE", 4), ("MAKINE", 41): ("DİZGİDE", 0),
                    ("MAKINE", 42): ("PLANA ALINDI", 0), ("ELLE", 21): ("DİZGİDE", 0), ("ELLE", 22): ("DİZGİDE", 0),
                    ("ELLE", 25): ("PLANA ALINDI", 0), ("EUM", 10): ("DİZGİDE", 0), ("EUM", 11): ("DİZGİDE", 1),
                    ("EUM", 16): ("PLANA ALINDI", 0)}
        self.assertEqual({a: (self.kart(*a)["durum"], self.kart(*a)["tamamlanan_adet"]) for a in beklenen}, beklenen)
        self.assertEqual((self.kart("MAKINE", 36)["toplam_adet"], self.kart("MAKINE", 42)["talep_no"]), (15, "1909042"))
        eski42 = [k for k in depo._kartlar if k["source_row_id"].startswith("NO:42~")]
        self.assertEqual(len(eski42), 1)
        self.assertEqual((eski42[0]["talep_no"], eski42[0]["source_active"]), ("1900042", 0))
        self.assertIn("Eski talebin notu", eski42[0]["aciklama"])
        for kod, no in (("MAKINE", 43), ("MAKINE", 44), ("ELLE", 33)):
            self.assertEqual(self.kart(kod, no)["source_active"], 0)

        # 3. adım: aynı dosya tekrar: yalnız hâlâ Excel'den ileride olan kartlar karar ister
        onizleme = self.yukle(paket["03_AYNI_DOSYA_TEKRAR"], onizleme=True)
        self.assertEqual((onizleme["yeni"], onizleme["guncellenen"], onizleme["pasife_alinan"], onizleme["gerileme"]),
                         (0, 0, 0, 5))
        secim = {d["id"]: d["gerileme"]["onerilen"] for d in onizleme["degisiklikler"] if d["tur"] == "gerileme"}
        self.yukle(paket["03_AYNI_DOSYA_TEKRAR"], gerileme_secimleri=secim)

        # 4. adım: Excel yetişti: karar yok, 5 kartın yalnız Excel durum metni değişir
        onizleme = self.yukle(paket["04_EXCEL_YETISTI"], onizleme=True)
        self.assertEqual((onizleme["gerileme"], onizleme["guncellenen"], onizleme["yeni"], onizleme["pasife_alinan"]),
                         (0, 5, 0, 0))
        self.assertEqual({d["sira"] for d in onizleme["degisiklikler"]}, {24, 39, 40, 22, 11})
        self.yukle(paket["04_EXCEL_YETISTI"], gerileme_secimleri={})
        once = [dict(k) for k in depo._kartlar]

        # 5. adım: hatalı dosyalar reddedilir, uyarı dosyaları yalnız önizlenir
        hatalar = {"H01_EKSIK_TALEP_NO": "Talep NO boş, ancak satırda kayıt verisi var",
                   "H02_TANINMAYAN_DURUM": "DURUM 'TESLİM EDİLDİ (KISMİ)' tanınmıyor",
                   "H03_TEKRARLANAN_NO": "Tekrarlanan kaynak kimliği NO:16",
                   "H04_GECERSIZ_ADET": "Üretim adedi tek bir sayı içermeli",
                   "H05_GECERSIZ_TARIH": "Planlanan Teslim Tarihi okunamadı",
                   "H06_BASLAMA_TESLIMDEN_SONRA": "Planlanan Teslim Tarihinden",
                   "H07_UC_HATA_BIRDEN": "3 sorun bulundu",
                   "H08_MAKINE_SAYFASI_YOK": "'MAKİNE' sayfası bulunamadı",
                   "H09_IKI_NO_SUTUNU": "'NO' için birden fazla sütun bulundu"}
        for ad, mesaj in hatalar.items():
            with self.subTest(ad):
                with self.assertRaises(ex.ExcelAktarimHatasi) as hata:
                    self.yukle(paket[ad], onizleme=True)
                self.assertIn(mesaj, str(hata.exception))
        u01 = self.yukle(paket["U01_EUM_SAYFASI_YOK"], onizleme=True)
        self.assertEqual((u01["pasife_alinan"], [s["sayfa"] for s in u01["kaynak_raporu"]["sayfalar"] if s["bulunamadi"]]),
                         (24, ["EÜM"]))
        u02 = self.yukle(paket["U02_ELDE_SAYFASI_BOSALTILDI"], onizleme=True)
        self.assertEqual((u02["pasife_alinan"], u02["mevcut_aktif"]), (34, 104))
        self.assertGreaterEqual(u02["pasife_alinan"] * 10, u02["mevcut_aktif"] * 3)   # toplu pasifleşme uyarısı
        self.assertEqual([dict(k) for k in depo._kartlar], once)

    @unittest.skipUnless(os.environ.get("PDGM_TEST_EXCEL_COM") == "1", "Set PDGM_TEST_EXCEL_COM=1 for real Excel COM")
    def test_gercek_excel_com_ile_ilk_iki_adim(self):
        paket = paket_uret.uret(self.root / "paket")
        sayfalar, _ = _rapor(self.yukle(paket["01_ILK_YUKLEME"], com=True))
        self.assertEqual(sayfalar, {"MAKINE": (45, 54), "ELLE": (35, 0), "EUM": (23, 4)})
        self.assertEqual(self.kart("MAKINE", 38)["plan_teslim"], (date.today() + timedelta(days=10)).isoformat())
        self.assertEqual(self.kart("MAKINE", 16)["plan_teslim"], (date.today() + timedelta(days=10)).isoformat())
        self.operator_islemleri()
        onizleme = self.yukle(paket["02_GUNCELLEME"], com=True, onizleme=True)
        self.assertEqual((onizleme["yeni"], onizleme["pasife_alinan"], onizleme["ayrilan"], onizleme["gerileme"]),
                         (5, 3, 1, 6))


if __name__ == "__main__":
    unittest.main()
