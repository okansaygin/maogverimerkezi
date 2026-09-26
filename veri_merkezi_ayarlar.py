# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - AYARLAR (veri_merkezi_ayarlar.py)
===================================================
Kod açmadan değiştirilebilen bütün ayarlar TEK bir dosyada tutulur:
    veri_merkezi_ayarlar.json   (uygulama klasöründe)

Dosya yoksa ya da bir değer eksik/bozuksa varsayılan kullanılır; uygulama asla
ayar yüzünden açılmamazlık etmez. Ayarlar sayfasından (Sistem > Ayarlar)
kaydedilen değerler doğrulanır, dosyaya güvenli biçimde (önce geçici dosyaya,
sonra tek adımda yerine) yazılır ve çalışan uygulamaya hemen uygulanır.

Kullanım:
    import veri_merkezi_ayarlar as ayarlar
    ayarlar.al("yillik_fm_siniri")          -> 270.0
    ayarlar.kaydet({"proje_adi": "..."})    -> (True, {})  |  (False, {anahtar: hata})

NOT - standart iş günü (9 saat) BİLEREK ayar değildir: içe aktarılan günlük
toplam saat, içe aktarma anında bu değere göre normal/fazla mesaiye bölünüp
kalıcı olarak yazılır. Sonradan değiştirilirse eski ve yeni kayıtlar farklı
kurala göre bölünmüş olur ve raporlar tutarsızlaşır.
"""

import json
import logging
import os
import tempfile
import threading
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
AYAR_DOSYASI = SCRIPT_DIR / "veri_merkezi_ayarlar.json"

gunluk = logging.getLogger("veri_merkezi")

# (anahtar, tür, varsayılan, en_az, en_çok, bölüm, etiket, açıklama)
TANIMLAR = [
    ("proje_adi", str, "MAOG Projesi", 1, 80, "genel", "Proje adı",
     "Günlük mesai formunun (PDF) başlığında yazar."),
    ("marka_alt_baslik", str, "Hızlı Tren Projesi · Puantaj", 1, 60, "genel", "Menü alt başlığı",
     "Sol menüde 'VERİ MERKEZİ' yazısının altında görünür."),

    ("yillik_fm_siniri", float, 270.0, 1, 2000, "uyum", "Yıllık fazla mesai sınırı (saat)",
     "4857 s. İş Kanunu md. 41. Verimlilik Analizi'ndeki uyum kontrolü ve Excel raporu kullanır; "
     "sınırın %85'i 'yaklaşan' sayılır."),
    ("gunluk_azami_saat", float, 11.0, 1, 24, "uyum", "Günlük azami çalışma (saat)",
     "4857 s. İş Kanunu md. 63. Normal + fazla mesai bu değeri aşan günler risk listesine girer."),
    ("kesintisiz_gun_esigi", int, 7, 2, 31, "uyum", "Kesintisiz çalışma eşiği (gün)",
     "Hafta tatili kullanmadan bu kadar gün ve üzeri çalışanlar işaretlenir (md. 46)."),
    ("devamsizlik_esigi", int, 3, 1, 31, "uyum", "Devamsızlık eşiği (gün)",
     "Dönemde bu kadar ve üzeri mazeretsiz devamsızlığı olanlar risk sayılır."),
    ("aylik_fm_esigi", float, 52.0, 0, 500, "uyum", "Aylık fazla mesai eşiği (saat)",
     "Dönem uzunluğuna oranlanır (ör. 15 günlük dönemde yarısı). Önceden Puantaj Suite "
     "ayarından okunuyordu."),

    ("fm_katsayi_haftaici", float, 1.5, 1, 5, "ucret", "Fazla mesai katsayısı · pazartesi-cumartesi",
     "Fazla mesai saatleri bununla çarpılır (hak edilmemiş pazarın fazla mesaisi dahil)."),
    ("fm_katsayi_pazar", float, 2.5, 1, 5, "ucret", "Pazar katsayısı",
     "Haftanın 6 günü tamamlanmışsa pazar günü çalışılan NORMAL + FAZLA saatlerin tamamı bununla çarpılır. "
     "Haftada ücretsiz izin / devamsızlık varsa pazar sıradan iş günü sayılır."),

    ("mesai_formu_bos_satir", int, 3, 0, 10, "form", "Mesai formu · boş satır (varsayılan)",
     "Listede olmayanlar için son sayfaya eklenen boş satır sayısının başlangıç değeri."),

    ("acilis_yedegi", bool, True, None, None, "yedek", "Günlük açılış yedeği",
     "Uygulama her gün ilk açıldığında veritabanının yedeği alınır."),
    ("yedek_saklama", int, 30, 3, 365, "yedek", "Saklanacak yedek sayısı",
     "Açılış yedeklerinden ve işlem öncesi yedeklerden her biri için en fazla bu kadarı tutulur; "
     "eskiler silinir. Elle alınan yedekler hiçbir zaman otomatik silinmez."),
]

BOLUMLER = [
    ("genel", "Genel"),
    ("uyum", "Uyum eşikleri"),
    ("ucret", "Ücret katsayıları"),
    ("form", "Mesai formu"),
    ("yedek", "Yedekleme"),
]

TANIM = {t[0]: t for t in TANIMLAR}
VARSAYILANLAR = {t[0]: t[2] for t in TANIMLAR}

_kilit = threading.Lock()
_onbellek = {"anahtar": None, "degerler": None}


class AyarHatasi(Exception):
    pass


# ---------------------------------------------------------------- doğrulama

def dogrula_deger(anahtar, deger):
    """Tek değeri doğrular ve doğru türe çevirir. Döner: (değer, None) | (None, hata_mesajı)."""
    if anahtar not in TANIM:
        return None, "bilinmeyen ayar"
    _a, tur, _v, en_az, en_cok, _b, _e, _ac = TANIM[anahtar]
    if tur is bool:
        if isinstance(deger, bool):
            return deger, None
        if isinstance(deger, (int, float)) and deger in (0, 1):
            return bool(deger), None
        if isinstance(deger, str) and deger.strip().lower() in ("true", "false", "1", "0", "evet", "hayır"):
            return deger.strip().lower() in ("true", "1", "evet"), None
        return None, "evet/hayır olmalı"
    if tur is str:
        if deger is None:
            return None, "boş olamaz"
        s = " ".join(str(deger).split())
        if len(s) < en_az:
            return None, "boş olamaz"
        if len(s) > en_cok:
            return None, f"en fazla {en_cok} karakter olabilir"
        return s, None
    # sayılar
    if isinstance(deger, bool) or deger is None or (isinstance(deger, str) and not deger.strip()):
        return None, "bir sayı girin"
    try:
        sayi = float(str(deger).strip().replace(",", ".")) if isinstance(deger, str) else float(deger)
    except (TypeError, ValueError):
        return None, "bir sayı girin"
    if sayi != sayi or sayi in (float("inf"), float("-inf")):
        return None, "bir sayı girin"
    if tur is int:
        if not float(sayi).is_integer():
            return None, "tam sayı olmalı"
        sayi = int(sayi)
    if en_az is not None and sayi < en_az:
        return None, f"en az {en_az} olabilir"
    if en_cok is not None and sayi > en_cok:
        return None, f"en fazla {en_cok} olabilir"
    return (int(sayi) if tur is int else float(sayi)), None


# ---------------------------------------------------------------- okuma

def _dosya_imzasi(yol):
    try:
        st = yol.stat()
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _oku(yol):
    """Dosyadaki geçerli değerler + varsayılanlar. Bozuk değer/dosya uygulamayı durdurmaz, günlüğe yazılır."""
    degerler = dict(VARSAYILANLAR)
    if not yol.exists():
        return degerler
    try:
        ham = json.loads(yol.read_text(encoding="utf-8"))
        if not isinstance(ham, dict):
            raise ValueError("kök öğe bir sözlük değil")
    except (OSError, ValueError) as e:
        gunluk.warning("Ayar dosyası okunamadı, varsayılanlar kullanılıyor (%s): %s", yol, e)
        return degerler
    for anahtar, deger in ham.items():
        if anahtar.startswith("_"):
            continue
        if anahtar not in TANIM:
            gunluk.warning("Ayar dosyasında bilinmeyen anahtar yok sayıldı: %s", anahtar)
            continue
        dogru, hata = dogrula_deger(anahtar, deger)
        if hata:
            gunluk.warning("Ayar '%s' geçersiz (%r: %s), varsayılan kullanılıyor", anahtar, deger, hata)
            continue
        degerler[anahtar] = dogru
    return degerler


def hepsi(yol=None):
    """Bütün ayarların güncel değerleri (dosya değişince kendiliğinden yeniden okunur)."""
    yol = Path(yol) if yol else AYAR_DOSYASI
    imza = (str(yol), _dosya_imzasi(yol))
    with _kilit:
        if _onbellek["anahtar"] != imza:
            _onbellek["degerler"] = _oku(yol)
            _onbellek["anahtar"] = imza
        return dict(_onbellek["degerler"])


def al(anahtar, yol=None):
    if anahtar not in TANIM:
        raise AyarHatasi(f"Bilinmeyen ayar: {anahtar}")
    return hepsi(yol)[anahtar]


# ---------------------------------------------------------------- yazma

def kaydet(yeni_degerler, yol=None, uygula_=True):
    """Verilen ayarları doğrular; HEPSİ geçerliyse dosyaya yazar ve uygular.
    Döner: (True, {}) ya da (False, {anahtar: hata_mesajı}) - hata varsa HİÇBİR ŞEY yazılmaz."""
    yol = Path(yol) if yol else AYAR_DOSYASI
    hatalar, temiz = {}, {}
    for anahtar, deger in (yeni_degerler or {}).items():
        dogru, hata = dogrula_deger(anahtar, deger)
        if hata:
            hatalar[anahtar] = hata
        else:
            temiz[anahtar] = dogru
    if hatalar:
        return False, hatalar

    birlesik = hepsi(yol)
    birlesik.update(temiz)
    icerik = {"_aciklama": "Veri Merkezi ayarları. Uygulamadaki Ayarlar sayfasından değiştirmeniz önerilir."}
    icerik.update({t[0]: birlesik[t[0]] for t in TANIMLAR})

    yol.parent.mkdir(parents=True, exist_ok=True)
    fd, gecici = tempfile.mkstemp(prefix=".ayarlar_", suffix=".tmp", dir=str(yol.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(icerik, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(gecici, yol)
    except OSError as e:
        try:
            os.unlink(gecici)
        except OSError:
            pass
        raise AyarHatasi(f"Ayar dosyası yazılamadı: {e}") from e

    with _kilit:
        _onbellek["anahtar"] = None
    gunluk.info("Ayarlar kaydedildi: %s", ", ".join(f"{k}={v}" for k, v in temiz.items()))
    if uygula_:
        uygula(yol)
    return True, {}


# ---------------------------------------------------------------- uygulama

def uygula(yol=None):
    """Ayarları çalışan modüllere aktarır. Modüller bu değerleri kendi sabitleri olarak
    (vm.YILLIK_FM_SINIRI gibi) kullanır; her hesaplama anında okudukları için yeni
    değer bir sonraki hesaplamadan itibaren geçerlidir. Analiz önbelleği temizlenir."""
    d = hepsi(yol)
    import veri_merkezi_sorgu as sorgu
    import veri_merkezi_verimlilik as vm
    import veri_merkezi_mesai_formu as mf

    sorgu.FAZLA_MESAI_KATSAYI_HAFTAICI = d["fm_katsayi_haftaici"]
    sorgu.FAZLA_MESAI_KATSAYI_PAZAR = d["fm_katsayi_pazar"]
    vm.K_HI = d["fm_katsayi_haftaici"]
    vm.K_PZ = d["fm_katsayi_pazar"]
    vm.YILLIK_FM_SINIRI = d["yillik_fm_siniri"]
    vm.GUNLUK_AZAMI_SAAT = d["gunluk_azami_saat"]
    vm.KESINTISIZ_GUN_ESIGI = d["kesintisiz_gun_esigi"]
    vm.DEVAMSIZLIK_ESIGI = d["devamsizlik_esigi"]
    mf.PROJE_ADI = d["proje_adi"]
    mf.VARSAYILAN_BOS_SATIR = d["mesai_formu_bos_satir"]
    try:
        vm._ONBELLEK.clear()
    except AttributeError:
        pass
    return d
