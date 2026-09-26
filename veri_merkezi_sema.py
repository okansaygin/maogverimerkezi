# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - ŞEMA (veri_merkezi_sema.py)
=============================================
Kalıcı veri deposunun (personel_veritabani.db - personnel_db.py ile AYNI dosya)
şema tanımı ve sürüm yönetimi.

Mevcut personnel_db.py'nin "personnel" / "personnel_history" tablolarına DOKUNMAZ,
sadece yeni tablolar ekler (idempotent - tekrar tekrar çalıştırılabilir, veri
kaybına yol açmaz):

  - gunluk_puantaj    kişi + gün başına TEK satır: o günün normal_saat VE
                       fazla_saat'i AYNI satırda (sürüm 3, bkz. aşağı). Bu,
                       Veri Merkezi'nin kalbidir: bir kez buraya yazılan veri
                       bir daha ASLA ham Excel'den okunmaz.
  - daily_timesheet   ESKİ (sürüm 2) yapı: kişi + gün + tip (normal/fazla)
                       başına AYRI satır. Sürüm 3 ile gunluk_puantaj'a
                       taşındı; artık YAZILMIYOR ama var olan veri ASLA
                       silinmez (arşiv/iz olarak durur).
  - import_batches    her içe aktarmanın denetim kaydı (hangi dosya, ne zaman,
                       kaç yeni/güncellenen personel, kaç uyarı, hangi durumda).
  - quality_alerts    tespit edilen TÜM veri kalitesi sorunlarının kalıcı listesi.
                       Excel'de bir sayfada görünüp kaybolmaz - çözülene kadar
                       burada durur, "çözüldü" olarak işaretlenebilir.
  - schema_meta       şema sürüm numarası (ileride yeni alan eklerken mevcut
                       veriyi kaybetmeden yükseltme yapabilmek için).

Sürüm 5 (uygulama v5.0.0):
  - puantaj_history          gunluk_puantaj'daki HER değişikliğin izi (ekle / güncelle /
                              sil): eski ve yeni saat + durum, hangi içe aktarma yaptı.
                              Bir içe aktarma bir günün değerini değiştirdiğinde eski
                              değer artık KAYBOLMAZ; içe aktarma geri alınabilir.
  - personel_degisiklik_izi   personel bilgileri içe aktarmasının yaptığı ekleme ve alan
                              değişiklikleri (geri alma için). personnel_history
                              (personnel_db.py) aynen yazılmaya devam eder.
  - import_batches            + izli (1 = bu içe aktarmanın izi tutuldu, geri alınabilir),
                              + geri_alma_zamani, geri_alma_notu, geri_alma_ozeti.
                              durum artık 'geri_alindi' de olabilir.
  - Veritabanı WAL kipine alınır (okuma ve yazma birbirini beklemez; "database is
    locked" hatası pratikte ortadan kalkar). Bağlantılar 15 sn bekleme süresiyle açılır.
"""

__version__ = "2026-09-25.2"


import logging
import sqlite3
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
DB_PATH = SCRIPT_DIR / "personel_veritabani.db"

SCHEMA_VERSION = 5
BAGLANTI_BEKLEME_SN = 15   # başka bir bağlantı yazarken en fazla bu kadar beklenir

gunluk = logging.getLogger("veri_merkezi")

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS schema_meta (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS daily_timesheet (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tc TEXT NOT NULL,
    tarih TEXT NOT NULL,              -- ISO: YYYY-MM-DD
    tip TEXT NOT NULL,                -- 'normal' | 'fazla'
    saat REAL NOT NULL,
    kaynak_satir INTEGER,             -- ham exceldeki satır no (izlenebilirlik)
    import_batch_id INTEGER NOT NULL,
    UNIQUE(tc, tarih, tip)
);
CREATE INDEX IF NOT EXISTS idx_timesheet_tc ON daily_timesheet(tc);
CREATE INDEX IF NOT EXISTS idx_timesheet_tarih ON daily_timesheet(tarih);
CREATE INDEX IF NOT EXISTS idx_timesheet_batch ON daily_timesheet(import_batch_id);

-- Sürüm 3: kişi + gün başına TEK satır (normal_saat + fazla_saat aynı
-- satırda). "SAP Mesai" günlük dosyasından (bkz. veri_merkezi_mesai_ice_aktarma.py)
-- beslenir - her satır zaten bir kişinin BİR günü olduğu için burada da
-- doğal olarak tek satır kalır (2 satır - normal/fazla ayrı - YOK).
CREATE TABLE IF NOT EXISTS gunluk_puantaj (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tc TEXT NOT NULL,
    tarih TEXT NOT NULL,              -- ISO: YYYY-MM-DD
    normal_saat REAL NOT NULL DEFAULT 0,
    fazla_saat REAL NOT NULL DEFAULT 0,
    durum TEXT,                        -- 'ÇALIŞTI' / 'ÜCRETSİZ İZİN' / 'SAĞLIK RAPORU' vb.
    kaynak_satir INTEGER,
    import_batch_id INTEGER,
    UNIQUE(tc, tarih)
);
CREATE INDEX IF NOT EXISTS idx_gp_tc ON gunluk_puantaj(tc);
CREATE INDEX IF NOT EXISTS idx_gp_tarih ON gunluk_puantaj(tarih);

CREATE TABLE IF NOT EXISTS import_batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dosya_adi TEXT NOT NULL,
    dosya_hash TEXT NOT NULL,
    donem_baslangic TEXT,
    donem_bitis TEXT,
    baslama_zamani TEXT NOT NULL,
    tamamlanma_zamani TEXT,
    durum TEXT NOT NULL DEFAULT 'basladi',   -- basladi | tamamlandi | iptal | hata
    taranan_satir INTEGER DEFAULT 0,
    yeni_personel INTEGER DEFAULT 0,
    guncellenen_personel INTEGER DEFAULT 0,
    yazilan_gun_kaydi INTEGER DEFAULT 0,
    engelleyici_sayisi INTEGER DEFAULT 0,
    uyari_sayisi INTEGER DEFAULT 0,
    hata_mesaji TEXT
);
CREATE INDEX IF NOT EXISTS idx_import_hash ON import_batches(dosya_hash);

CREATE TABLE IF NOT EXISTS quality_alerts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER,
    tip TEXT NOT NULL,                -- kural adı (bkz. veri_merkezi_dogrulama.py)
    onem TEXT NOT NULL DEFAULT 'uyari',   -- 'engelleyici' | 'uyari'
    tc TEXT,
    detay TEXT NOT NULL,
    olusturma_zamani TEXT NOT NULL,
    cozuldu INTEGER NOT NULL DEFAULT 0,
    cozum_notu TEXT,
    cozulme_zamani TEXT
);
CREATE INDEX IF NOT EXISTS idx_alerts_cozuldu ON quality_alerts(cozuldu);
CREATE INDEX IF NOT EXISTS idx_alerts_tc ON quality_alerts(tc);
CREATE INDEX IF NOT EXISTS idx_alerts_batch ON quality_alerts(import_batch_id);

-- Sürüm 5: gunluk_puantaj değişiklik izi. Her satır bir (tc, tarih) kaydındaki TEK değişikliktir.
--   islem      'ekle' | 'guncelle' | 'sil'
--   eski_*     değişiklikten ÖNCEKİ değerler ('ekle'de NULL)
--   yeni_*     değişiklikten SONRAKİ değerler ('sil'de NULL)
--   import_batch_id   değişikliği yapan içe aktarma (geri almada: geri alınan içe aktarma)
--   geri_alma  1 = bu satır bir GERİ ALMA işleminin kaydıdır
CREATE TABLE IF NOT EXISTS puantaj_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tc TEXT NOT NULL,
    tarih TEXT NOT NULL,
    islem TEXT NOT NULL,
    eski_normal REAL, eski_fazla REAL, eski_durum TEXT, eski_kaynak_satir INTEGER, eski_batch_id INTEGER,
    yeni_normal REAL, yeni_fazla REAL, yeni_durum TEXT, yeni_kaynak_satir INTEGER,
    import_batch_id INTEGER NOT NULL,
    geri_alma INTEGER NOT NULL DEFAULT 0,
    zaman TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_ph_tc_tarih ON puantaj_history(tc, tarih);
CREATE INDEX IF NOT EXISTS idx_ph_batch ON puantaj_history(import_batch_id);

-- Sürüm 5: personel bilgileri içe aktarmasının izi (geri alma için).
--   islem  'ekle' (kişi bu içe aktarmayla eklendi) | 'guncelle' (bir alan değişti)
--          | 'sil' / 'geri_yukle' (geri alma kayıtları, geri_alma=1)
CREATE TABLE IF NOT EXISTS personel_degisiklik_izi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER NOT NULL,
    tc TEXT NOT NULL,
    islem TEXT NOT NULL,
    alan TEXT,
    eski_deger TEXT,
    yeni_deger TEXT,
    eski_kaynak TEXT,
    geri_alma INTEGER NOT NULL DEFAULT 0,
    zaman TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_pdi_batch ON personel_degisiklik_izi(import_batch_id);
CREATE INDEX IF NOT EXISTS idx_pdi_tc ON personel_degisiklik_izi(tc);
"""


def get_connection(db_path=None):
    db_path = db_path or DB_PATH
    conn = sqlite3.connect(str(db_path), timeout=BAGLANTI_BEKLEME_SN)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def wal_kipine_al(conn):
    """Veritabanını WAL kipine alır (kalıcıdır, dosyaya yazılır). Zaten WAL ise hiçbir şey yapmaz.
    Başarısız olursa (ör. başka bir program dosyayı kilitli tutuyorsa) uygulama eski kipte çalışmaya
    devam eder; bir sonraki açılışta yeniden denenir. Döner: geçerli kip ('wal', 'delete' ...)."""
    try:
        kip = conn.execute("PRAGMA journal_mode").fetchone()[0]
        if str(kip).lower() != "wal":
            kip = conn.execute("PRAGMA journal_mode=WAL").fetchone()[0]
        return str(kip).lower()
    except sqlite3.Error as e:
        gunluk.warning("WAL kipine geçilemedi: %s", e)
        return "?"


def gunluk_kipi(db_path=None):
    conn = get_connection(db_path)
    try:
        return str(conn.execute("PRAGMA journal_mode").fetchone()[0]).lower()
    finally:
        conn.close()


def wal_birlestir(db_path=None):
    """WAL dosyasındaki değişiklikleri ana .db dosyasına aktarıp -wal dosyasını boşaltır.
    Uygulama kapanırken çağrılır: böylece .db dosyası tek başına güncel olur."""
    try:
        conn = get_connection(db_path)
        try:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        finally:
            conn.close()
    except sqlite3.Error as e:
        gunluk.warning("WAL birleştirilemedi: %s", e)


def _ensure_columns(conn, table, columns):
    """SQLite'ta 'ADD COLUMN IF NOT EXISTS' yoktur - önce mevcut sütunları
    kontrol edip sadece EKSİK olanları ekler. Var olan veriyi ASLA silmez,
    mevcut satırlarda yeni sütunlar NULL olarak başlar."""
    existing = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    for col_name, col_type in columns:
        if col_name not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")


def migrate(db_path=None):
    """Şemayı en güncel sürüme getirir. Var olan tabloları/veriyi ASLA silmez -
    sadece eksik tabloları/sütunları/indeksleri ekler. Her uygulama başlangıcında
    çağrılması güvenlidir (idempotent)."""
    conn = get_connection(db_path)
    try:
        wal_kipine_al(conn)
        conn.executescript(SCHEMA_SQL)
        # Sürüm 2: quality_alerts'e yapılandırılmış "ad_soyad" ve "kaynak_satir"
        # sütunları eklendi (önceden sadece serbest metin "detay" vardı) -
        # Veri Kalitesi sayfasında düzgün tablo sütunları göstermek için.
        _ensure_columns(conn, "quality_alerts", [("ad_soyad", "TEXT"), ("kaynak_satir", "INTEGER")])

        # Sürüm 4: import_batches'e "tur" ('kimlik' | 'mesai') sütunu eklendi -
        # İçe Aktarma Geçmişi ekranında iki içe aktarma türünü ayırt etmek için.
        # Eski satırlarda NULL kalır; sorgu katmanı bunları sayılara bakarak tahmin eder.
        _ensure_columns(conn, "import_batches", [("tur", "TEXT")])

        # Sürüm 5: geri alma. "izli" = bu içe aktarmanın değişiklik izi tutuldu (v5 ve sonrası);
        # v5 öncesi içe aktarmalarda NULL kalır ve bunlar geri alınamaz (eski değerler kayıtlı değil).
        _ensure_columns(conn, "import_batches", [("izli", "INTEGER"), ("geri_alma_zamani", "TEXT"),
                                                 ("geri_alma_notu", "TEXT"), ("geri_alma_ozeti", "TEXT")])

        # Sürüm 5.2: İK bilgileri tablosu (veri_merkezi_ik) - Personel Kartı'ndaki İK sekmeleri
        import veri_merkezi_ik as ik
        ik.tablo_olustur(conn)

        # Sürüm 3: eski daily_timesheet (kişi+gün+tip başına 2 satır) verisini,
        # bir kereliğine, yeni gunluk_puantaj (kişi+gün başına 1 satır) yapısına
        # AKTAR. daily_timesheet SİLİNMEZ (arşiv olarak durur). UNIQUE(tc,tarih)
        # + INSERT OR IGNORE sayesinde bu adım tamamen idempotent: tekrar tekrar
        # çağrılması veya elle silinmiş bir satırın yeniden oluşması sorun değil.
        # PERFORMANS: bu kopyalama bir kez (sürüm 3'e geçerken) gereklidir. Önceden her
        # migrate() çağrısında (ör. Ekip Listesi her açıldığında) eski arşivin TAMAMI
        # yeniden gruplanıp yazılıyordu. Artık kayıtlı sürüm >= 3 ise atlanır.
        kayitli = conn.execute("SELECT value FROM schema_meta WHERE key='version'").fetchone()
        kayitli_surum = int(kayitli["value"]) if kayitli and str(kayitli["value"]).isdigit() else 0
        eski_var = kayitli_surum < 3 and conn.execute("SELECT 1 FROM daily_timesheet LIMIT 1").fetchone()
        if eski_var:
            gruplu = conn.execute(
                "SELECT tc, tarih, "
                "SUM(CASE WHEN tip='normal' THEN saat ELSE 0 END) AS normal_saat, "
                "SUM(CASE WHEN tip='fazla' THEN saat ELSE 0 END) AS fazla_saat, "
                "MAX(import_batch_id) AS import_batch_id "
                "FROM daily_timesheet GROUP BY tc, tarih"
            ).fetchall()
            for row in gruplu:
                conn.execute(
                    "INSERT OR IGNORE INTO gunluk_puantaj (tc, tarih, normal_saat, fazla_saat, "
                    "import_batch_id) VALUES (?,?,?,?,?)",
                    (row["tc"], row["tarih"], row["normal_saat"], row["fazla_saat"], row["import_batch_id"]),
                )

        cur = conn.execute("SELECT value FROM schema_meta WHERE key='version'")
        row = cur.fetchone()
        if row is None:
            conn.execute("INSERT INTO schema_meta (key, value) VALUES ('version', ?)",
                         (str(SCHEMA_VERSION),))
        elif not str(row["value"]).isdigit() or int(row["value"]) < SCHEMA_VERSION:
            # Yalnızca değiştiğinde yazılır: her migrate() çağrısında yazmak dosyanın değişiklik zamanını
            # günceller ve Verimlilik Analizi önbelleğini boş yere geçersiz kılardı.
            conn.execute("UPDATE schema_meta SET value=? WHERE key='version'", (str(SCHEMA_VERSION),))
        conn.commit()
    finally:
        conn.close()


def get_schema_version(db_path=None):
    conn = get_connection(db_path)
    try:
        cur = conn.execute("SELECT value FROM schema_meta WHERE key='version'")
        row = cur.fetchone()
        return int(row["value"]) if row else 0
    finally:
        conn.close()
