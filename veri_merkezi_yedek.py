# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - VERİTABANI YEDEKLEME (veri_merkezi_yedek.py)
==============================================================
Bütün veri tek bir dosyada (personel_veritabani.db). Bu modül o dosyanın
tutarlı kopyalarını alır, eskilerini temizler ve gerektiğinde geri yükler.

Yedekler:  <uygulama klasörü>/veritabani_yedekleri/
           personel_veritabani__2026-09-25_08-07-12__acilis.db

Ne zaman alınır:
  - acilis        : uygulama her gün İLK açıldığında (şema yükseltmesinden ÖNCE)
  - ice_aktarma   : her içe aktarma yazılmadan hemen önce
  - geri_alma     : bir içe aktarma geri alınmadan önce
  - geri_yukleme  : bir yedek geri yüklenmeden önce (geri yüklemenin de geri dönüşü olsun diye)
  - sema          : uygulama yeni bir şema sürümüne yükseltilmeden önce (ör. v4 -> v5)
  - elle          : Ayarlar sayfasındaki "Şimdi yedek al" düğmesi

Neden dosya kopyalama DEĞİL: veritabanı WAL kipinde çalışır; son değişiklikler bir
süre -wal dosyasında durur. Dosyayı düz kopyalamak bu değişiklikleri kaçırabilir ya
da yarım yazılmış bir sayfayı kopyalayabilir. SQLite'ın yedekleme API'si (backup)
uygulama çalışırken bile TUTARLI bir anlık görüntü üretir. Üretilen yedek tek bir
.db dosyasıdır (WAL kapalı), doğrudan açılabilir, taşınabilir.

Saklama: 'otomatik' (açılış) ve 'işlem' (içe aktarma / geri alma / geri yükleme)
yedeklerinin her birinden en yeni N tanesi tutulur (Ayarlar > yedek_saklama).
Elle alınan yedekler asla otomatik silinmez.
"""

import datetime
import logging
import re
import sqlite3
from pathlib import Path

import veri_merkezi_ayarlar as ayarlar
import veri_merkezi_sema as sema

gunluk = logging.getLogger("veri_merkezi")

KLASOR_ADI = "veritabani_yedekleri"

# tür -> (etiket, saklama kategorisi)
TURLER = {
    "acilis": ("Açılış", "otomatik"),
    "ice_aktarma": ("İçe aktarma öncesi", "islem"),
    "geri_alma": ("Geri alma öncesi", "islem"),
    "geri_yukleme": ("Geri yükleme öncesi", "islem"),
    "sema": ("Şema yükseltme öncesi", "islem"),
    "elle": ("Elle alınan", "elle"),
}
_AD_DESENI = re.compile(r"^(?P<kok>.+)__(?P<zaman>\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})(?:-(?P<sira>\d+))?__"
                        r"(?P<tur>[a-z_]+)\.db$")


class YedekHatasi(Exception):
    pass


def _db_yolu(db_path=None):
    return Path(db_path) if db_path else Path(sema.DB_PATH)


def yedek_klasoru(db_path=None):
    return _db_yolu(db_path).parent / KLASOR_ADI


def _baglan_salt_okunur(yol):
    return sqlite3.connect(f"file:{Path(yol).as_posix()}?mode=ro", uri=True, timeout=15)


# ---------------------------------------------------------------- yedek alma

def yedek_al(tur="elle", db_path=None):
    """Tutarlı bir yedek alır. Döner: yedek bilgisi sözlüğü. Hata olursa YedekHatasi."""
    if tur not in TURLER:
        raise YedekHatasi(f"Bilinmeyen yedek türü: {tur}")
    kaynak_yol = _db_yolu(db_path)
    if not kaynak_yol.exists():
        raise YedekHatasi(f"Veritabanı bulunamadı: {kaynak_yol}")

    klasor = yedek_klasoru(db_path)
    klasor.mkdir(parents=True, exist_ok=True)
    zaman = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    ad = f"{kaynak_yol.stem}__{zaman}__{tur}.db"
    sira = 2
    while (klasor / ad).exists():
        ad = f"{kaynak_yol.stem}__{zaman}-{sira}__{tur}.db"
        sira += 1
    hedef = klasor / ad
    gecici = klasor / (ad + ".yaziliyor")

    kaynak = sqlite3.connect(str(kaynak_yol), timeout=15)
    try:
        if gecici.exists():
            gecici.unlink()
        hedef_baglanti = sqlite3.connect(str(gecici))
        try:
            kaynak.backup(hedef_baglanti)
            # Yedek tek başına açılabilen, kendi -wal dosyası olmayan tek bir dosya olsun.
            hedef_baglanti.execute("PRAGMA journal_mode=DELETE")
            sonuc = hedef_baglanti.execute("PRAGMA quick_check").fetchone()[0]
            if sonuc != "ok":
                raise YedekHatasi(f"Yedek doğrulanamadı: {sonuc}")
        finally:
            hedef_baglanti.close()
        gecici.replace(hedef)
    except sqlite3.Error as e:
        raise YedekHatasi(f"Yedek alınamadı: {e}") from e
    finally:
        kaynak.close()
        if gecici.exists():
            try:
                gecici.unlink()
            except OSError:
                pass

    bilgi = _dosya_bilgisi(hedef)
    gunluk.info("yedek alındı: %s (%s)", hedef.name, _boyut_metni(bilgi["boyut"]))
    try:
        temizle(db_path)
    except Exception:
        gunluk.exception("eski yedekler temizlenemedi")
    return bilgi


def islem_oncesi_yedek(tur, db_path=None):
    """İçe aktarma / geri alma öncesi yedek. Başarısız olursa işlemi DURDURMAZ (değişiklik geçmişi
    zaten geri almaya izin verir); günlüğe yazar ve None döner."""
    try:
        return yedek_al(tur, db_path)
    except Exception as e:
        gunluk.warning("%s öncesi yedek alınamadı: %s", tur, e)
        return None


def acilis_yedegi(db_path=None, bugun=None):
    """Bugün için açılış yedeği yoksa alır (ayar kapalıysa ya da veritabanı yoksa/boşsa almaz).
    Döner: yedek bilgisi | None. Hiçbir durumda istisna fırlatmaz: uygulamanın açılmasını engellememeli."""
    try:
        if not ayarlar.al("acilis_yedegi"):
            return None
        yol = _db_yolu(db_path)
        if not yol.exists() or yol.stat().st_size == 0:
            return None
        bugun = bugun or datetime.date.today()
        if any(y["tur"] == "acilis" and y["zaman"].date() == bugun for y in yedekleri_listele(db_path)):
            return None
        return yedek_al("acilis", db_path)
    except Exception as e:
        gunluk.warning("açılış yedeği alınamadı: %s", e)
        return None


# ---------------------------------------------------------------- listeleme / temizlik

def _boyut_metni(n):
    for birim in ("B", "KB", "MB", "GB"):
        if n < 1024 or birim == "GB":
            return f"{n:.0f} {birim}" if birim == "B" else f"{n:.1f} {birim}".replace(".", ",")
        n /= 1024.0
    return str(n)


def _dosya_bilgisi(yol):
    m = _AD_DESENI.match(yol.name)
    if not m:
        return None
    try:
        zaman = datetime.datetime.strptime(m.group("zaman"), "%Y-%m-%d_%H-%M-%S")
        boyut = yol.stat().st_size
    except (ValueError, OSError):
        return None
    tur = m.group("tur")
    etiket, kategori = TURLER.get(tur, (tur, "islem"))
    return {"ad": yol.name, "yol": str(yol), "tur": tur, "tur_etiket": etiket, "kategori": kategori,
            "zaman": zaman, "sira": int(m.group("sira") or 1), "boyut": boyut, "boyut_metni": _boyut_metni(boyut)}


def yedekleri_listele(db_path=None):
    """En yeniden eskiye yedek listesi."""
    klasor = yedek_klasoru(db_path)
    if not klasor.is_dir():
        return []
    liste = [b for b in (_dosya_bilgisi(p) for p in klasor.glob("*.db")) if b]
    liste.sort(key=lambda b: (b["zaman"], b["sira"], b["ad"]), reverse=True)
    return liste


def temizle(db_path=None, saklama=None):
    """Her saklama kategorisinden (otomatik, işlem) en yeni `saklama` kadarını tutar, fazlasını siler.
    Elle alınan yedeklere dokunmaz. Döner: silinen dosya adları."""
    saklama = int(saklama if saklama is not None else ayarlar.al("yedek_saklama"))
    silinen = []
    for kategori in ("otomatik", "islem"):
        grup = [y for y in yedekleri_listele(db_path) if y["kategori"] == kategori]
        for y in grup[saklama:]:
            try:
                Path(y["yol"]).unlink()
                silinen.append(y["ad"])
            except OSError as e:
                gunluk.warning("eski yedek silinemedi: %s (%s)", y["ad"], e)
    if silinen:
        gunluk.info("eski yedekler silindi: %d dosya", len(silinen))
    return silinen


def yedek_bul(ad, db_path=None):
    """Yalnızca yedek klasöründeki, adı desene uyan bir dosyayı döner (yol dışına çıkılamaz)."""
    if not ad or "/" in ad or "\\" in ad or ad.startswith("."):
        return None
    yol = yedek_klasoru(db_path) / ad
    if not yol.is_file() or not _AD_DESENI.match(ad):
        return None
    return yol


# ---------------------------------------------------------------- doğrulama / geri yükleme

def yedek_incele(yol):
    """Yedeği SALT OKUNUR açar, bütünlüğünü ve içeriğini kontrol eder.
    Döner: {"gecerli": bool, "hata": str|None, "personel": n, "gunluk_kayit": n, "ice_aktarma": n,
            "sema": n, "son_tarih": "YYYY-MM-DD"|None}"""
    sonuc = {"gecerli": False, "hata": None, "personel": 0, "gunluk_kayit": 0, "ice_aktarma": 0,
             "sema": 0, "son_tarih": None}
    try:
        conn = _baglan_salt_okunur(yol)
    except sqlite3.Error as e:
        sonuc["hata"] = f"Dosya açılamadı: {e}"
        return sonuc
    try:
        kontrol = conn.execute("PRAGMA integrity_check").fetchone()[0]
        if kontrol != "ok":
            sonuc["hata"] = f"Bütünlük kontrolü başarısız: {kontrol}"
            return sonuc
        tablolar = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if "personnel" not in tablolar:
            sonuc["hata"] = "Bu dosya bir Veri Merkezi veritabanı değil (personnel tablosu yok)."
            return sonuc
        sonuc["personel"] = conn.execute("SELECT COUNT(*) FROM personnel").fetchone()[0]
        if "gunluk_puantaj" in tablolar:
            sonuc["gunluk_kayit"] = conn.execute("SELECT COUNT(*) FROM gunluk_puantaj").fetchone()[0]
            sonuc["son_tarih"] = conn.execute("SELECT MAX(tarih) FROM gunluk_puantaj").fetchone()[0]
        if "import_batches" in tablolar:
            sonuc["ice_aktarma"] = conn.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0]
        if "schema_meta" in tablolar:
            r = conn.execute("SELECT value FROM schema_meta WHERE key='version'").fetchone()
            sonuc["sema"] = int(r[0]) if r and str(r[0]).isdigit() else 0
        sonuc["gecerli"] = True
        return sonuc
    except sqlite3.Error as e:
        sonuc["hata"] = f"Dosya okunamadı: {e}"
        return sonuc
    finally:
        conn.close()


def geri_yukle(yedek_yolu, db_path=None):
    """Seçilen yedeği çalışan veritabanının YERİNE koyar.
    1) Yedek salt okunur açılıp doğrulanır (bozuksa hiçbir şey yapılmaz).
    2) Mevcut veritabanının 'geri_yukleme' yedeği alınır - bu ALINAMAZSA işlem durur.
    3) SQLite yedekleme API'siyle yedeğin içeriği mevcut veritabanına tek seferde yazılır.
    4) Şema güncel sürüme yükseltilir (eski bir yedekten dönülmüş olabilir).
    Döner: {"onceki_yedek": ad, "inceleme": {...}}"""
    yedek_yolu = Path(yedek_yolu)
    hedef_yol = _db_yolu(db_path)
    if yedek_yolu.resolve() == hedef_yol.resolve():
        raise YedekHatasi("Çalışan veritabanı kendi üzerine geri yüklenemez.")
    inceleme = yedek_incele(yedek_yolu)
    if not inceleme["gecerli"]:
        raise YedekHatasi(inceleme["hata"] or "Yedek geçersiz.")
    if inceleme["sema"] > sema.SCHEMA_VERSION:
        raise YedekHatasi(f"Bu yedek daha yeni bir sürüme ait (şema v{inceleme['sema']}); "
                          f"bu uygulama en fazla v{sema.SCHEMA_VERSION}'i tanır.")

    onceki = None
    if hedef_yol.exists():
        onceki = yedek_al("geri_yukleme", db_path)   # alınamazsa YedekHatasi -> geri yükleme yapılmaz

    kaynak = _baglan_salt_okunur(yedek_yolu)
    hedef = sqlite3.connect(str(hedef_yol), timeout=30)
    try:
        kaynak.backup(hedef)
    except sqlite3.Error as e:
        raise YedekHatasi(f"Geri yükleme başarısız: {e}. Mevcut veritabanı değişmedi"
                          f"{' (önceki hâlinin yedeği: ' + onceki['ad'] + ')' if onceki else ''}.") from e
    finally:
        hedef.close()
        kaynak.close()

    import personnel_db as pdb
    pdb.init_db(db_path)
    sema.migrate(db_path)
    gunluk.warning("YEDEK GERİ YÜKLENDİ: %s (önceki hâl: %s)", yedek_yolu.name, onceki["ad"] if onceki else "-")
    return {"onceki_yedek": onceki["ad"] if onceki else None, "inceleme": inceleme}


def sema_yukseltme_yedegi(db_path=None):
    """Kayıtlı şema sürümü uygulamanınkinden eskiyse, yükseltmeden ÖNCE yedek alır.
    İstisna fırlatmaz; döner: yedek bilgisi | None."""
    try:
        yol = _db_yolu(db_path)
        if not yol.exists() or yol.stat().st_size == 0:
            return None
        conn = _baglan_salt_okunur(yol)
        try:
            tablo = conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='schema_meta'").fetchone()
            r = conn.execute("SELECT value FROM schema_meta WHERE key='version'").fetchone() if tablo else None
        finally:
            conn.close()
        kayitli = int(r[0]) if r and str(r[0]).isdigit() else 0
        if kayitli >= sema.SCHEMA_VERSION:
            return None
        return yedek_al("sema", db_path)
    except Exception as e:
        gunluk.warning("şema yükseltme öncesi yedek alınamadı: %s", e)
        return None
