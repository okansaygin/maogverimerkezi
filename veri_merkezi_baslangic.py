# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - AÇILIŞ ÖNCESİ KONTROL (veri_merkezi_baslangic.py)
===================================================================
Uygulama, bulunduğu klasöre yazabilmelidir: veritabanı (+ WAL dosyaları), yedekler,
günlük dosyası ve ayarlar oraya yazılır. Klasör yazmaya kapalıysa (Windows Güvenliği
"Denetimli klasör erişimi", şirket güvenlik yazılımı, salt okunur kopya, eşitleme
klasörü...) eskiden onlarca satırlık anlaşılmaz hata çıkıyordu. Bu modül bunu açılışta
TEK bir denemeyle yakalar ve ne yapılacağını Türkçe söyler.

veri_merkezi_main.py uygulamayı yüklemeden ÖNCE çağırır.
"""

import os
import shutil
import sqlite3
import sys
from pathlib import Path


def _klasor_yazilabilir_mi(klasor):
    """Klasörde alt klasör + SQLite (WAL) veritabanı oluşturmayı dener. Döner: hata metni | None."""
    test = klasor / f".vm_yazma_testi_{os.getpid()}"
    try:
        test.mkdir(exist_ok=True)
        conn = sqlite3.connect(str(test / "test.db"), timeout=5)
        try:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("CREATE TABLE t (x)")
            conn.execute("INSERT INTO t VALUES (1)")
            conn.commit()
        finally:
            conn.close()
        with open(test / "test.txt", "w", encoding="utf-8") as f:
            f.write("ok")
            f.flush()
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"
    finally:
        shutil.rmtree(test, ignore_errors=True)


def _veritabani_yazilabilir_mi(db_yolu):
    """Var olan veritabanına yazma kilidi alınabiliyor mu (salt okunur dosya, başka program kilidi)."""
    if not db_yolu.exists():
        return None
    if not os.access(db_yolu, os.W_OK):
        return "dosya salt okunur (Özellikler > 'Salt okunur' işaretli)"
    try:
        conn = sqlite3.connect(str(db_yolu), timeout=5)
        try:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("ROLLBACK")
        finally:
            conn.close()
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def kontrol(klasor, db_yolu):
    """Döner: None (sorun yok) ya da kullanıcıya gösterilecek açıklama metni."""
    klasor, db_yolu = Path(klasor), Path(db_yolu)
    hata = _klasor_yazilabilir_mi(klasor)
    if hata:
        return (
            "Veri Merkezi bu klasöre YAZAMIYOR, bu yüzden açılamadı:\n"
            f"    {klasor}\n\n"
            f"Teknik ayrıntı: {hata}\n\n"
            "En olası neden: Windows Güvenliği'nin 'Denetimli klasör erişimi' (fidye yazılımı koruması) ya da "
            "şirketin güvenlik yazılımı, Belgeler / Masaüstü gibi korunan klasörlerde Python'un dosya "
            "oluşturmasını engelliyor. Klasör salt okunur bir kopya ya da eşitleme (OneDrive, yedekleme) "
            "klasörü de olabilir.\n\n"
            "Çözüm (birini seçin):\n"
            "  1) Uygulama klasörünü korunmayan bir yere taşıyın, ör. C:\\VeriMerkezi  (önerilen)\n"
            "  2) Windows Güvenliği > Virüs ve tehdit koruması > Fidye yazılımına karşı koruma >\n"
            "     'Denetimli klasör erişimi üzerinden bir uygulamaya izin ver' > python.exe'yi ekleyin\n"
            "     (şirket bilgisayarında BT izni gerekebilir).\n\n"
            "Not: Bu bir yedek klasörüyse, uygulamayı yedekten değil asıl klasöründen çalıştırın; yedekten "
            "çalıştırılan uygulama yedeğin içindeki veritabanını değiştirir."
        )
    hata = _veritabani_yazilabilir_mi(db_yolu)
    if hata:
        return (
            "Veri Merkezi veritabanına YAZAMIYOR, bu yüzden açılamadı:\n"
            f"    {db_yolu}\n\n"
            f"Teknik ayrıntı: {hata}\n\n"
            "Olası nedenler: dosya salt okunur işaretli (sağ tık > Özellikler > 'Salt okunur' işaretini kaldırın), "
            "başka bir program (ör. açık kalmış eski bir Veri Merkezi ya da Puantaj Suite penceresi, bir "
            "veritabanı görüntüleyici) dosyayı kilitlemiş, ya da dosya ağ / eşitleme klasöründe."
        )
    return None


def hata_goster(metin):
    """Metni konsola yazar; Windows'ta ayrıca bir uyarı penceresi açar (çift tıkla başlatıldıysa görünsün)."""
    try:
        print("\n" + "=" * 78 + "\n" + metin + "\n" + "=" * 78 + "\n", file=sys.stderr)
    except Exception:
        pass
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, metin, "Veri Merkezi açılamadı", 0x10)
        except Exception:
            pass
