# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - İK (İNSAN KAYNAKLARI) BİLGİLERİ (veri_merkezi_ik.py)
======================================================================
İK sisteminden (SAP "PERSONEL LİSTESİ" dışa aktarımı) gelen, puantajda olmayan
personel bilgileri: kimlik, iletişim, acil durum, eğitim, pozisyon, taşeron,
konaklama, İSG tarihleri (tekrar eğitim / muayene, MYK) ve evrak durumu.

Tablo: personel_ik_bilgileri (TC başına tek satır). Personel Kartı bu tabloyu gösterir.
Kural: personnel tablosunda yalnız BOŞ işe giriş tarihi doldurulur (SGK giriş tarihinden);
grup / alt ekip İK'dan alınmaz (İK'nın Bölüm / Lokasyon sınıflandırması farklıdır).

Yeni bir İK listesini içe aktarmak için (uygulama kapalıyken):
    python veri_merkezi_ik.py "PERSONEL LİSTESİ.xlsx"
Veritabanında bulunmayan personel atlanır. Önce veritabanının yedeği alınır.
"""

import collections
import datetime
import json
import re
import sqlite3
from pathlib import Path

import veri_merkezi_sema as sema

TABLO = "personel_ik_bilgileri"

# Excel başlığı -> (sütun adı, tür)   tür: t=metin, d=tarih, b=true/false, x=evrak işareti (X), tel, kan
ESLEME = [
    ("Personel No", "personel_no", "t"), ("Onay Belge No", "onay_belge_no", "t"), ("Onay Durumu", "onay_durumu", "t"),
    ("Çalışan Statüsü", "calisan_statusu", "t"), ("Adı", "adi", "t"), ("Soyadı", "soyadi", "t"),
    ("Doğum Tarihi", "dogum_tarihi", "d"), ("Baba Adı", "baba_adi", "t"), ("Anne Adı", "anne_adi", "t"),
    ("Doğum Yeri", "dogum_yeri", "t"), ("Adres/Şehir", "adres_sehir", "t"), ("Ayrıntılı Adres", "adres", "t"),
    ("Cinsiyet", "cinsiyet", "t"), ("Medeni Durum", "medeni_durum", "t"), ("Telefon No", "telefon", "tel"),
    ("Acil Ad Soyad", "acil_kisi", "t"), ("Acil Telefon No", "acil_telefon", "tel"),
    ("Okul Türü", "egitim_kodu", "t"), ("Okul Türü Metni", "egitim_durumu", "t"),
    ("Eğitim dalı 1", "egitim_dali_kodu", "t"), ("Eğitim dalı metni", "egitim_dali", "t"),
    ("Kan Grubu", "kan_grubu", "kan"), ("Mail Adresi", "eposta", "t"),
    ("Meslek Kodu(SGK)", "sgk_meslek_kodu", "t"), ("Meslek Metni", "sgk_meslek", "t"),
    ("Pozisyon", "pozisyon_kodu", "t"), ("Pozisyon Tanım", "pozisyon", "t"), ("Görev", "gorev", "t"),
    ("Çalışan Grubu Tanımı", "calisan_grubu", "t"), ("Bölüm", "bolum", "t"),
    ("Normal Kadro", "normal_kadro", "b"), ("Destek İşçilik/Hizmet", "destek_iscilik", "b"),
    ("Geçici Görevlendirme", "gecici_gorevlendirme", "b"), ("Geçici Görev Bitiş Tarihi", "gecici_gorev_bitis", "d"),
    ("SGK Giriş Tarihi", "sgk_giris_tarihi", "d"), ("İşe giriş trh.", "ise_giris_tarihi", "d"),
    ("Form Oluşturma Tarihi", "form_olusturma_tarihi", "d"),
    ("Taşeron Firma", "taseron_kodu", "t"), ("Taşeron Firma Adı", "taseron_firma", "t"),
    ("Lokasyon Kodu", "lokasyon_kodu", "t"), ("Lokasyon", "lokasyon", "t"), ("Çalışan Tipi", "calisan_tipi", "t"),
    ("Pers.aln.metni", "personel_alani", "t"), ("Pers.alt alanı", "personel_alt_alani", "t"),
    ("Ekip", "ik_ekip_kodu", "t"), ("Alt alan metni", "alt_alan", "t"), ("Görev Yeri", "gorev_yeri", "t"),
    ("Konaklama Seçenekleri", "konaklama", "t"), ("Lojman oda ve Bina no", "lojman_bilgisi", "t"),
    ("MYK Başlangıç Tarihi", "myk_baslangic", "d"), ("MYK Bitiş Tarihi", "myk_bitis", "d"),
    ("Tekrar Eğitim Tarihi", "tekrar_egitim_tarihi", "d"), ("Tekrar Muayene Tarihi", "tekrar_muayene_tarihi", "d"),
    ("Oluşturan", "olusturan", "t"),
    ("Diğer Dökümanlar", "evrak_diger", "x"), ("SGK İşe Giriş Bildirgesi", "evrak_sgk_bildirge", "x"),
    ("Diploma ve Sertifikalar", "evrak_diploma", "x"), ("Sözleşme", "evrak_sozlesme", "x"),
    ("İşe Giriş ve Periyodik Muayene Raporu", "evrak_muayene", "x"), ("İSG-Ç Evrakları", "evrak_isg", "x"),
    ("İSG-Ç Periyodik Eğitim Evrakları", "evrak_isg_egitim", "x"), ("Yeniden Giriş", "yeniden_giris_tarihi", "d"),
    ("Ulaşım Seçenekleri", "ulasim", "t"), ("Yaka Kartı", "yaka_karti", "x"), ("Karakol Bildirimi", "karakol_bildirimi", "b"),
]
SAHTE_TEL = re.compile(r"^0?(5?0+|1+)$")      # 05000000000, 0500000000, 01111111111 gibi yer tutucular


def cevir(v, tur):
    if v is None or (isinstance(v, str) and not v.strip()):
        return 0 if tur == "x" else None        # evrak işareti yoksa evrak yok
    if tur == "d":
        if isinstance(v, datetime.datetime):
            return v.date().isoformat()
        if isinstance(v, datetime.date):
            return v.isoformat()
        return str(v).strip()[:10]
    if tur == "b":
        return 1 if str(v).strip().lower() == "true" else 0
    if tur == "x":
        return 1 if str(v).strip().upper() == "X" else 0
    s = " ".join(str(v).split())
    if tur == "tel":
        rakam = re.sub(r"\D", "", s)
        return None if SAHTE_TEL.match(rakam) else s
    if tur == "kan":
        return s.replace("RH", "Rh").replace("ABRh", "AB Rh").replace("0Rh", "0 Rh")
    return s




def tablo_olustur(conn):
    """Tabloyu (yoksa) oluşturur; sonradan eklenen sütunları tamamlar. sema.migrate() çağırır."""
    sutunlar = ",\n    ".join(f"{s} {'INTEGER' if t in ('b', 'x') else 'TEXT'}" for _h, s, t in ESLEME)
    conn.executescript(f"""
CREATE TABLE IF NOT EXISTS {TABLO} (
    tc TEXT PRIMARY KEY,
    {sutunlar},
    kayit_sayisi INTEGER,
    onceki_kayitlar TEXT,
    kaynak_dosya TEXT,
    aktarma_zamani TEXT
);""")
    mevcut = {r[1] for r in conn.execute(f"PRAGMA table_info({TABLO})")}
    for _h, s, t in ESLEME:
        if s not in mevcut:
            conn.execute(f"ALTER TABLE {TABLO} ADD COLUMN {s} {'INTEGER' if t in ('b', 'x') else 'TEXT'}")


def getir(tc, db_path=None):
    """Bir kişinin İK kaydı (sözlük) ya da None. onceki_kayitlar listeye çevrilir."""
    conn = sema.get_connection(db_path)
    try:
        if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (TABLO,)).fetchone():
            return None
        r = conn.execute(f"SELECT * FROM {TABLO} WHERE tc=?", (tc,)).fetchone()
    finally:
        conn.close()
    if not r:
        return None
    d = dict(r)
    try:
        d["onceki_kayitlar"] = json.loads(d["onceki_kayitlar"]) if d.get("onceki_kayitlar") else []
    except ValueError:
        d["onceki_kayitlar"] = []
    return d


def sil(tc, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (TABLO,)).fetchone():
            conn.execute(f"DELETE FROM {TABLO} WHERE tc=?", (tc,))
            conn.commit()
    finally:
        conn.close()


def _esas(kayitlar):
    """Aynı TC'nin birden çok kaydından 'Etkin' olan (yoksa en yenisi) esas alınır."""
    etkin = [k for k in kayitlar if k["calisan_statusu"] == "Etkin"]
    return max(etkin or kayitlar, key=lambda k: (k["ise_giris_tarihi"] or "", k["form_olusturma_tarihi"] or "",
                                                  k["onay_belge_no"] or ""))


def ice_aktar(xlsx_yolu, db_path=None):
    """İK listesini içe aktarır. Döner: özet sözlüğü. Tek transaction: hata olursa hiçbir şey yazılmaz."""
    import openpyxl
    import puantaj_engine as engine
    import veri_merkezi_yedek as yedek

    xlsx_yolu = Path(xlsx_yolu)
    wb = openpyxl.load_workbook(xlsx_yolu, read_only=True, data_only=True)
    satirlar = list(wb.active.iter_rows(values_only=True))
    ix = {h: k for k, h in enumerate(satirlar[0])}
    if "TC Kimlik No" not in ix:
        raise ValueError("Dosyada 'TC Kimlik No' başlığı yok; İK personel listesi değil gibi görünüyor.")
    ik = collections.defaultdict(list)
    for r in satirlar[1:]:
        tc = engine.normalize_tc(r[ix["TC Kimlik No"]])
        if tc:
            ik[tc].append({s: (cevir(r[ix[h]], t) if h in ix else None) for h, s, t in ESLEME})

    sema.migrate(db_path)
    yedek.islem_oncesi_yedek("ice_aktarma", db_path)
    simdi = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    kaynak = f"ik_listesi:{xlsx_yolu.name}"
    ozet = {"ik_kayit": len(ik), "eslesen": 0, "atlanan": 0, "giris_doldurulan": 0}
    conn = sema.get_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        for p in conn.execute("SELECT tc, ise_giris_tarihi FROM personnel").fetchall():
            tc = p["tc"]
            if tc not in ik:
                ozet["atlanan"] += 1
                continue
            k = _esas(ik[tc])
            onceki = [{a: x[a] for a in ("calisan_statusu", "ise_giris_tarihi", "sgk_giris_tarihi", "gorev", "bolum",
                                         "lokasyon", "onay_belge_no", "form_olusturma_tarihi")}
                      for x in ik[tc] if x is not k]
            deger = dict(k, tc=tc, kayit_sayisi=len(ik[tc]),
                         onceki_kayitlar=json.dumps(onceki, ensure_ascii=False) if onceki else None,
                         kaynak_dosya=xlsx_yolu.name, aktarma_zamani=simdi)
            conn.execute(f"INSERT OR REPLACE INTO {TABLO} ({','.join(deger)}) VALUES ({','.join('?' * len(deger))})",
                         list(deger.values()))
            ozet["eslesen"] += 1
            yeni = k["sgk_giris_tarihi"] or k["ise_giris_tarihi"]
            if not (p["ise_giris_tarihi"] or "").strip() and yeni:
                conn.execute("UPDATE personnel SET ise_giris_tarihi=?, updated_at=? WHERE tc=?", (yeni, simdi, tc))
                conn.execute("INSERT INTO personnel_history (tc, field, old_value, new_value, changed_at, source) "
                             "VALUES (?,?,?,?,?,?)", (tc, "ise_giris_tarihi", p["ise_giris_tarihi"], yeni, simdi, kaynak))
                ozet["giris_doldurulan"] += 1
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return ozet


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print('Kullanım: python veri_merkezi_ik.py "PERSONEL LİSTESİ.xlsx"')
        sys.exit(2)
    o = ice_aktar(sys.argv[1])
    print(f"İK listesinde {o['ik_kayit']} kişi · veritabanında eşleşen {o['eslesen']} · İK'da bulunmayan "
          f"{o['atlanan']} (dokunulmadı) · doldurulan işe giriş tarihi {o['giris_doldurulan']}")
