# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - SORGU KATMANI (veri_merkezi_sorgu.py)
========================================================
Raporlama motorlarının ve Dash arayüzünün kullandığı TEK temiz API.
Bu katmanın altındaki hiçbir çağıran taraf artık ham Excel dosyasını
açmaz - her şey burada, kalıcı depodan (personel_veritabani.db) gelir.
"""

__version__ = "2026-09-25.1"


import datetime

import veri_merkezi_ayarlar as ayarlar
import veri_merkezi_sema as sema

# Katsayılar Ayarlar sayfasından değiştirilebilir (veri_merkezi_ayarlar.uygula bu değerleri günceller).
FAZLA_MESAI_KATSAYI_HAFTAICI = ayarlar.al("fm_katsayi_haftaici")   # varsayılan 1,5
FAZLA_MESAI_KATSAYI_PAZAR = ayarlar.al("fm_katsayi_pazar")          # varsayılan 2,5
STANDART_GUN_SAATI = 9.0   # BİLEREK sabit: içe aktarmada normal/fazla ayrımı buna göre yapılır (bkz. ayarlar)


def _daterange(d1, d2):
    for n in range((d2 - d1).days + 1):
        yield d1 + datetime.timedelta(days=n)


# ---------------------------------------------------------------- genel bakış

def get_dashboard_summary(donem_baslangic=None, donem_bitis=None, db_path=None, saatler=True):
    """Dashboard'un üst KPI kartları için özet: toplam/aktif personel,
    seçili dönemde toplam normal/fazla/yevmiye ve kritik uyarı sayısı.
    saatler=False: yalnız personel ve uyarı sayıları (bütün puantajı taramaz; Genel Bakış bunu kullanır)."""
    conn = sema.get_connection(db_path)
    try:
        toplam = conn.execute("SELECT COUNT(*) FROM personnel").fetchone()[0]
        aktif = conn.execute(
            "SELECT COUNT(*) FROM personnel WHERE isten_cikis_tarihi IS NULL OR isten_cikis_tarihi=''"
        ).fetchone()[0]

        params = []
        where = "1=1"
        if donem_baslangic:
            where += " AND tarih >= ?"
            params.append(donem_baslangic)
        if donem_bitis:
            where += " AND tarih <= ?"
            params.append(donem_bitis)

        # v5.1: yevmiye hafta tatili / pazar kuralıyla hesaplanır (veri_merkezi_hafta_kurali)
        import veri_merkezi_hafta_kurali as hk
        if saatler:
            t = hk.db_toplami(donem_baslangic, donem_bitis, db_path=db_path)
            normal, fazla, yevmiye = t["normal"], t["fazla"], t["yevmiye"]
        else:
            normal = fazla = yevmiye = 0.0

        kritik_uyari = conn.execute(
            "SELECT COUNT(*) FROM quality_alerts WHERE cozuldu=0 AND onem='engelleyici'"
        ).fetchone()[0]
        acik_uyari = conn.execute("SELECT COUNT(*) FROM quality_alerts WHERE cozuldu=0").fetchone()[0]

        return {
            "toplam_personel": toplam, "aktif_personel": aktif,
            "normal_saat": normal, "fazla_saat": fazla, "toplam_saat": normal + fazla,
            "yevmiye_gunu": round(yevmiye, 1),
            "kritik_uyari": kritik_uyari, "acik_uyari": acik_uyari,
        }
    finally:
        conn.close()


def get_monthly_trend(ay_sayisi=12, db_path=None):
    """Son N ayın normal/fazla mesai toplamlarını döner (trend grafiği için)."""
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT strftime('%Y-%m', tarih) AS ay, SUM(normal_saat) AS normal, "
            "SUM(fazla_saat) AS fazla FROM gunluk_puantaj GROUP BY ay ORDER BY ay"
        ).fetchall()
        return [dict(r) for r in rows][-ay_sayisi:]
    finally:
        conn.close()


def get_group_breakdown(donem_baslangic=None, donem_bitis=None, db_path=None):
    """Grup bazında normal/fazla/toplam mesai (dashboard çubuk grafiği için)."""
    conn = sema.get_connection(db_path)
    try:
        params = []
        where = "1=1"
        if donem_baslangic:
            where += " AND gp.tarih >= ?"
            params.append(donem_baslangic)
        if donem_bitis:
            where += " AND gp.tarih <= ?"
            params.append(donem_bitis)
        rows = conn.execute(
            f"SELECT p.grup AS grup, SUM(gp.normal_saat) AS normal, SUM(gp.fazla_saat) AS fazla, "
            f"COUNT(DISTINCT p.tc) AS personel_sayisi "
            f"FROM gunluk_puantaj gp JOIN personnel p ON p.tc = gp.tc "
            f"WHERE {where} GROUP BY p.grup", params
        ).fetchall()
        by_group = {}
        for r in rows:
            g = by_group.setdefault(r["grup"] or "(Grup belirtilmemiş)",
                                     {"normal": 0.0, "fazla": 0.0, "personel_sayisi": 0})
            g["normal"] = r["normal"] or 0.0
            g["fazla"] = r["fazla"] or 0.0
            g["personel_sayisi"] = r["personel_sayisi"]
        return [{"grup": g, **v} for g, v in sorted(by_group.items(), key=lambda kv: -(kv[1]["normal"] + kv[1]["fazla"]))]
    finally:
        conn.close()


# ---------------------------------------------------------------- puantaj geçmişi sorgusu

def get_daily_hours(tc=None, grup=None, donem_baslangic=None, donem_bitis=None, db_path=None):
    """Kişi/grup ve tarih aralığına göre günlük saat kayıtlarını döner -
    Excel açmadan anında cevap."""
    conn = sema.get_connection(db_path)
    try:
        params = []
        where = "1=1"
        if tc:
            where += " AND gp.tc = ?"
            params.append(tc)
        if grup:
            where += " AND p.grup = ?"
            params.append(grup)
        if donem_baslangic:
            where += " AND gp.tarih >= ?"
            params.append(donem_baslangic)
        if donem_bitis:
            where += " AND gp.tarih <= ?"
            params.append(donem_bitis)
        rows = conn.execute(
            f"SELECT gp.tc, COALESCE(p.ad_soyad, '(personel kaydı yok)') AS ad_soyad, p.grup, "
            f"p.bagli_oldugu_ekip, p.gorevi, gp.tarih, gp.durum, gp.normal_saat, gp.fazla_saat, "
            f"(gp.normal_saat + gp.fazla_saat) AS toplam_saat "
            f"FROM gunluk_puantaj gp LEFT JOIN personnel p ON p.tc = gp.tc "
            f"WHERE {where} ORDER BY gp.tarih, ad_soyad",
            params,
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_person_period_totals(tc, donem_baslangic, donem_bitis, db_path=None):
    """Bir kişinin dönem içindeki normal/fazla/yevmiye toplamını döner (hafta tatili / pazar kuralıyla)."""
    import veri_merkezi_hafta_kurali as hk
    t = hk.db_toplami(donem_baslangic, donem_bitis, tc=tc, db_path=db_path)
    return {"normal": t["normal"], "fazla": t["fazla"], "yevmiye_gunu": round(t["yevmiye"], 1),
            "pazar_katli": t["pazar_katli"]}


# ---------------------------------------------------------------- veri kalitesi

def get_quality_alerts(cozuldu=None, onem=None, tip=None, limit=500, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        params = []
        where = "1=1"
        if cozuldu is not None:
            where += " AND cozuldu = ?"
            params.append(1 if cozuldu else 0)
        if onem:
            where += " AND onem = ?"
            params.append(onem)
        if tip:
            where += " AND tip = ?"
            params.append(tip)
        params.append(limit)
        # NOT: qa.ad_soyad (yeni, doğrulama anında yakalanan isim) ile
        # p.ad_soyad (personnel tablosundaki güncel isim) AYNI ADI taşıdığı
        # için "qa.*" ile birlikte seçilmez - açıkça COALESCE edilir (önce
        # doğrulama anındaki isim, o da yoksa personnel tablosundaki isim).
        rows = conn.execute(
            f"SELECT qa.id, qa.import_batch_id, qa.tip, qa.onem, qa.tc, "
            f"COALESCE(qa.ad_soyad, p.ad_soyad) AS ad_soyad, qa.kaynak_satir, qa.detay, "
            f"qa.olusturma_zamani, qa.cozuldu, qa.cozum_notu, qa.cozulme_zamani "
            f"FROM quality_alerts qa LEFT JOIN personnel p ON p.tc = qa.tc "
            f"WHERE {where} ORDER BY qa.olusturma_zamani DESC LIMIT ?", params
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def resolve_alert(alert_id, cozum_notu="", db_path=None):
    conn = sema.get_connection(db_path)
    try:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "UPDATE quality_alerts SET cozuldu=1, cozum_notu=?, cozulme_zamani=? WHERE id=?",
            (cozum_notu, now, alert_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_alert_type_counts(db_path=None):
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT tip, onem, COUNT(*) AS adet FROM quality_alerts WHERE cozuldu=0 "
            "GROUP BY tip, onem ORDER BY adet DESC"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------- içe aktarma geçmişi

def get_import_history(limit=50, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT * FROM import_batches ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_covered_periods(db_path=None):
    """Hangi ayların verisi var, hangileri eksik - Genel Bakış'taki takvim
    zaman çizelgesi için."""
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT DISTINCT strftime('%Y-%m', tarih) AS ay FROM gunluk_puantaj ORDER BY ay"
        ).fetchall()
        return [r["ay"] for r in rows]
    finally:
        conn.close()


# =====================================================================
# SÜRÜM 2026-09-24 — YENİ ARAYÜZ (tasarım yenilemesi) İÇİN EK SORGULAR
# Aşağıdaki fonksiyonlar yukarıdakileri DEĞİŞTİRMEZ, sadece ekler.
# =====================================================================

# Veri Kalitesi ekranındaki önem dereceleri. Veritabanında "onem" sadece
# 'engelleyici' | 'uyari' olarak tutulur (engelleyiciler zaten hiç yazılmaz,
# içe aktarmayı durdurur). Arayüz, uyarıları etkilerine göre üç seviyede gösterir:
#   kritik : veri yanlış kişiye yazılmış / tamamen kaybolmuş olabilir
#   uyari  : kontrol edilmeli
#   bilgi  : eksik ama zararsız
KRITIK_TIPLER = {"ayni_tc_farkli_isim", "tcsiz_saat_ozeti", "dosya_tekrar", "baslik_eslesmedi"}
BILGI_TIPLER = {"giris_cikis_tutarsiz", "kimlik_bilgisi_eksik", "personel_sayisi_anomalisi"}

UYARI_ETIKET = {
    "tc_eksik": "TC eksik",
    "ayni_tc_farkli_isim": "Aynı TC, farklı isim",
    "olasi_farkli_tc": "Olası farklı TC (aynı isim)",
    "tc_format_hatali": "TC format hatalı",
    "kimlik_bilgisi_eksik": "Kimlik bilgisi eksik",
    "tcsiz_saat_ozeti": "TC'siz kaybolan saatler",
    "gunluk_kayit_cakismasi": "Günlük kayıt çakışması",
    "anormal_saat": "Anormal saat değeri",
    "giris_cikis_tutarsiz": "Giriş/çıkış tutarsız",
    "zorunlu_alan_eksik": "Zorunlu alan eksik",
    "personel_sayisi_anomalisi": "Personel sayısı anomalisi",
    "personel_kaydi_yok": "Personel kaydı yok",
    "tarihsiz_satir": "Tarihsiz satır",
    "baslik_eslesmedi": "Başlık eşleşmedi",
    "dosya_tekrar": "Dosya tekrarı",
}

UYARI_ACIKLAMA = {
    "ayni_tc_farkli_isim": "Aynı TC iki kişiye aitse Ekip Listesi'nde tek satır görünür ve saatler yanlış kişiye yazılabilir.",
    "tcsiz_saat_ozeti": "TC'si olmayan satırların saatleri hiçbir personele yazılamadı — raporlarda eksik görünür.",
    "personel_kaydi_yok": "Saatler yazıldı ama kişinin grup/ekip/görev bilgisi yok; raporlarda '(Grup belirtilmemiş)' altında kalır.",
    "olasi_farkli_tc": "Aynı isim farklı TC ile gelmiş; TC'lerden biri hatalı yazılmış olabilir.",
    "anormal_saat": "Günlük saat 0–24 aralığının dışında; kaynakta yazım hatası olabilir.",
    "tc_eksik": "Bu satırlardan kimlik bilgisi alınamadı.",
    "tc_format_hatali": "TC 11 hane değil ya da sağlaması tutmuyor.",
    "gunluk_kayit_cakismasi": "Aynı kişi aynı gün için dosyada birden fazla kez geçiyor; son değer kullanıldı.",
}


def uyari_seviyesi(tip, onem=None):
    if onem == "engelleyici" or tip in KRITIK_TIPLER:
        return "kritik"
    if tip in BILGI_TIPLER:
        return "bilgi"
    return "uyari"


def yevmiye_hesapla(normal, fazla, katli_normal=0.0, katli_fazla=0.0):
    """v5.1: normal/fazla = TÜM saatler; katli_* = bunların ×2,5 ödenen (hak edilmiş) pazar kısmı.
    Hangi pazarın ×2,5 olduğuna veri_merkezi_hafta_kurali karar verir."""
    import veri_merkezi_hafta_kurali as hk
    return hk.yevmiye(normal, fazla, katli_normal, katli_fazla)


def ay_araligi(ay):
    """'2026-09' -> ('2026-09-01', '2026-09-30')"""
    y, m = (int(x) for x in ay.split("-"))
    bas = datetime.date(y, m, 1)
    son = (datetime.date(y + (m == 12), m % 12 + 1, 1) - datetime.timedelta(days=1))
    return bas.isoformat(), son.isoformat()


_TUR_SQL = ("COALESCE(tur, CASE WHEN yazilan_gun_kaydi > 0 THEN 'mesai' "
            "WHEN yeni_personel > 0 OR guncellenen_personel > 0 THEN 'kimlik' ELSE NULL END)")


# ---------------------------------------------------------------- genel bakış (yeni)

def get_period_summary(bas, bit, db_path=None):
    """Seçili dönem için normal / fazla (hafta içi + pazar) / yevmiye ve kişi sayısı."""
    import veri_merkezi_hafta_kurali as hk
    conn = sema.get_connection(db_path)
    try:
        r = conn.execute("SELECT COUNT(DISTINCT tc) AS kisi, COUNT(*) AS kayit FROM gunluk_puantaj "
                         "WHERE tarih BETWEEN ? AND ?", (bas, bit)).fetchone()
    finally:
        conn.close()
    t = hk.db_toplami(bas, bit, db_path=db_path)
    return {"normal": t["normal"], "fazla": t["fazla"], "pazar_katli": t["pazar_katli"],
            "toplam": t["normal"] + t["fazla"], "yevmiye": t["yevmiye"], "kisi": r["kisi"], "kayit": r["kayit"]}


def get_daily_coverage(bas, bit, db_path=None):
    """{tarih_iso: kayıt_sayısı} - kapsama takvimi için."""
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute("SELECT tarih, COUNT(*) AS n FROM gunluk_puantaj WHERE tarih BETWEEN ? AND ? "
                            "GROUP BY tarih", (bas, bit)).fetchall()
        return {r["tarih"]: r["n"] for r in rows}
    finally:
        conn.close()


def gun_kapsamasi(bas, bit, bugun=None, db_path=None, kismi_oran=0.8):
    """Her gün için durum: 'tam' | 'kismi' | 'eksik' | 'gelecek'. 'Tam' eşiği,
    o dönemdeki kayıtlı günlerin medyan kayıt sayısının %80'idir (hafta tatilleri
    de SAP dosyasında satır olarak geldiği için gün bazında karşılaştırılabilir)."""
    bugun = bugun or datetime.date.today()
    sayilar = get_daily_coverage(bas, bit, db_path)
    dolu = sorted(v for v in sayilar.values() if v)
    medyan = dolu[len(dolu) // 2] if dolu else 0
    d1, d2 = datetime.date.fromisoformat(bas), datetime.date.fromisoformat(bit)
    sonuc = []
    for d in _daterange(d1, d2):
        n = sayilar.get(d.isoformat(), 0)
        if n:
            durum = "tam" if n >= medyan * kismi_oran else "kismi"
        elif d >= bugun:
            durum = "gelecek"
        else:
            durum = "eksik"
        sonuc.append({"tarih": d, "kayit": n, "durum": durum})
    return sonuc


def get_last_import(tur, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        r = conn.execute(f"SELECT * FROM import_batches WHERE durum='tamamlandi' AND {_TUR_SQL}=? "
                         f"ORDER BY id DESC LIMIT 1", (tur,)).fetchone()
        return dict(r) if r else None
    finally:
        conn.close()


def get_alert_counts(cozuldu=False, db_path=None):
    """[{tip, etiket, seviye, adet}] - seviye ve adede göre sıralı."""
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute("SELECT tip, MAX(onem) AS onem, COUNT(*) AS adet FROM quality_alerts "
                            "WHERE cozuldu=? GROUP BY tip", (1 if cozuldu else 0,)).fetchall()
    finally:
        conn.close()
    sira = {"kritik": 0, "uyari": 1, "bilgi": 2}
    out = [{"tip": r["tip"], "etiket": UYARI_ETIKET.get(r["tip"], r["tip"].replace("_", " ").capitalize()),
            "seviye": uyari_seviyesi(r["tip"], r["onem"]), "adet": r["adet"]} for r in rows]
    return sorted(out, key=lambda a: (sira[a["seviye"]], -a["adet"]))


def get_alert_totals(db_path=None):
    counts = get_alert_counts(False, db_path)
    return {"acik": sum(a["adet"] for a in counts),
            "kritik": sum(a["adet"] for a in counts if a["seviye"] == "kritik")}


def get_alerts(tip=None, cozuldu=False, batch_id=None, tc=None, limit=2000, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        where, params = ["qa.cozuldu = ?"], [1 if cozuldu else 0]
        if tip:
            where.append("qa.tip = ?"); params.append(tip)
        if batch_id:
            where.append("qa.import_batch_id = ?"); params.append(batch_id)
        if tc:
            where.append("qa.tc = ?"); params.append(tc)
        params.append(limit)
        rows = conn.execute(
            "SELECT qa.id, qa.import_batch_id, qa.tip, qa.onem, qa.tc, "
            "COALESCE(qa.ad_soyad, p.ad_soyad) AS ad_soyad, qa.kaynak_satir, qa.detay, "
            "qa.olusturma_zamani, qa.cozuldu, qa.cozum_notu, qa.cozulme_zamani, "
            "ib.dosya_adi, (p.tc IS NOT NULL) AS personel_var "
            "FROM quality_alerts qa LEFT JOIN personnel p ON p.tc = qa.tc "
            "LEFT JOIN import_batches ib ON ib.id = qa.import_batch_id "
            f"WHERE {' AND '.join(where)} ORDER BY qa.id DESC LIMIT ?", params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_alert(alert_id, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        r = conn.execute(
            "SELECT qa.*, COALESCE(qa.ad_soyad, p.ad_soyad) AS ad_soyad_goster, p.grup, p.bagli_oldugu_ekip, "
            "p.ad_soyad AS kayitli_isim, (p.tc IS NOT NULL) AS personel_var, ib.dosya_adi "
            "FROM quality_alerts qa LEFT JOIN personnel p ON p.tc = qa.tc "
            "LEFT JOIN import_batches ib ON ib.id = qa.import_batch_id WHERE qa.id=?", (alert_id,)).fetchone()
        return dict(r) if r else None
    finally:
        conn.close()


def resolve_alerts(alert_ids, cozum_notu="", db_path=None):
    """Birden fazla uyarıyı TEK işlemde 'çözüldü' işaretler. Döner: işaretlenen sayısı."""
    ids = [int(i) for i in alert_ids if i is not None]
    if not ids:
        return 0
    conn = sema.get_connection(db_path)
    try:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = conn.execute(
            f"UPDATE quality_alerts SET cozuldu=1, cozum_notu=?, cozulme_zamani=? "
            f"WHERE cozuldu=0 AND id IN ({','.join('?' * len(ids))})", [cozum_notu, now] + ids)
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def reopen_alert(alert_id, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        conn.execute("UPDATE quality_alerts SET cozuldu=0, cozum_notu=NULL, cozulme_zamani=NULL WHERE id=?",
                     (alert_id,))
        conn.commit()
    finally:
        conn.close()


def get_alert_type_trend(tip, n=6, db_path=None):
    """Son n tamamlanmış içe aktarmada bu tipten kaç uyarı üretildiği (eskiden yeniye)."""
    conn = sema.get_connection(db_path)
    try:
        batches = conn.execute("SELECT id, baslama_zamani FROM import_batches WHERE durum='tamamlandi' "
                               "ORDER BY id DESC LIMIT ?", (n,)).fetchall()
        out = []
        for b in reversed(batches):
            adet = conn.execute("SELECT COUNT(*) FROM quality_alerts WHERE import_batch_id=? AND tip=?",
                                (b["id"], tip)).fetchone()[0]
            out.append({"batch_id": b["id"], "tarih": (b["baslama_zamani"] or "")[:10], "adet": adet})
        return out
    finally:
        conn.close()


def get_open_alert_tcs(db_path=None):
    conn = sema.get_connection(db_path)
    try:
        return {r[0] for r in conn.execute("SELECT DISTINCT tc FROM quality_alerts WHERE cozuldu=0 AND tc IS NOT NULL")}
    finally:
        conn.close()


# ---------------------------------------------------------------- içe aktarma geçmişi (yeni)

def get_import_history_v2(limit=200, tur=None, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        sql = f"SELECT *, {_TUR_SQL} AS tur_hesap FROM import_batches"
        params = []
        if tur:
            sql += f" WHERE {_TUR_SQL} = ?"
            params.append(tur)
        sql += " ORDER BY id DESC LIMIT ?"
        params.append(limit)
        rows = [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()
    for r in rows:
        r["sure_sn"] = None
        try:
            a = datetime.datetime.strptime(r["baslama_zamani"], "%Y-%m-%d %H:%M:%S")
            b = datetime.datetime.strptime(r["tamamlanma_zamani"], "%Y-%m-%d %H:%M:%S")
            r["sure_sn"] = (b - a).total_seconds()
        except (TypeError, ValueError):
            pass
    return rows


def get_import_stats(gun=30, db_path=None):
    esik = (datetime.datetime.now() - datetime.timedelta(days=gun)).strftime("%Y-%m-%d %H:%M:%S")
    conn = sema.get_connection(db_path)
    try:
        r = conn.execute(
            "SELECT COUNT(*) AS toplam, SUM(durum='tamamlandi') AS basarili, SUM(durum='hata') AS hata, "
            "SUM(durum='geri_alindi') AS geri_alinan, "
            "COALESCE(SUM(CASE WHEN durum='tamamlandi' THEN yazilan_gun_kaydi ELSE 0 END),0) AS gun_kaydi "
            "FROM import_batches WHERE baslama_zamani >= ?",
            (esik,)).fetchone()
        return {k: (r[k] or 0) for k in ("toplam", "basarili", "hata", "geri_alinan", "gun_kaydi")}
    finally:
        conn.close()


def get_batch_alert_breakdown(batch_id, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute("SELECT tip, MAX(onem) AS onem, COUNT(*) AS adet FROM quality_alerts "
                            "WHERE import_batch_id=? GROUP BY tip ORDER BY adet DESC", (batch_id,)).fetchall()
        return [{"tip": r["tip"], "etiket": UYARI_ETIKET.get(r["tip"], r["tip"]),
                 "seviye": uyari_seviyesi(r["tip"], r["onem"]), "adet": r["adet"]} for r in rows]
    finally:
        conn.close()


def get_batch_affected_tcs(batch_id, db_path=None):
    """Bu içe aktarmada kaydı değişen/eklenen personel (personnel_history + gunluk_puantaj)."""
    conn = sema.get_connection(db_path)
    try:
        b = conn.execute("SELECT dosya_adi FROM import_batches WHERE id=?", (batch_id,)).fetchone()
        tcs = set()
        if b:
            etiket = f"veri_merkezi:{b['dosya_adi']}"
            tcs |= {r[0] for r in conn.execute("SELECT DISTINCT tc FROM personnel_history WHERE source=?", (etiket,))}
            tcs |= {r[0] for r in conn.execute("SELECT tc FROM personnel WHERE kaynak=?", (etiket,))}
        tcs |= {r[0] for r in conn.execute("SELECT DISTINCT tc FROM gunluk_puantaj WHERE import_batch_id=?", (batch_id,))}
        # v5: değişiklik izi (geri alınmış içe aktarmalarda da etkilenenleri gösterir)
        tcs |= {r[0] for r in conn.execute("SELECT DISTINCT tc FROM puantaj_history WHERE import_batch_id=?", (batch_id,))}
        tcs |= {r[0] for r in conn.execute("SELECT DISTINCT tc FROM personel_degisiklik_izi WHERE import_batch_id=?",
                                           (batch_id,))}
        return tcs
    finally:
        conn.close()


# ---------------------------------------------------------------- personel (yeni)

def _personel_where(search=None, durumlar=None, gruplar=None, tc_set=None):
    where, params = ["1=1"], []
    if durumlar and len(durumlar) == 1:
        if "aktif" in durumlar:
            where.append("(p.isten_cikis_tarihi IS NULL OR p.isten_cikis_tarihi='')")
        else:
            where.append("(p.isten_cikis_tarihi IS NOT NULL AND p.isten_cikis_tarihi<>'')")
    if gruplar:
        g = [x for x in gruplar if x != "__bos__"]
        parts = []
        if g:
            parts.append(f"p.grup IN ({','.join('?' * len(g))})")
            params += g
        if "__bos__" in gruplar:
            parts.append("(p.grup IS NULL OR p.grup='')")
        where.append("(" + " OR ".join(parts) + ")")
    if search:
        like = f"%{search.strip()}%"
        where.append("(p.ad_soyad LIKE ? OR p.tc LIKE ? OR p.bagli_oldugu_ekip LIKE ? OR p.gorevi LIKE ?)")
        params += [like] * 4
    if tc_set is not None:
        tcs = list(tc_set) or ["__yok__"]
        where.append(f"p.tc IN ({','.join('?' * len(tcs))})")
        params += tcs
    return " AND ".join(where), params


def get_personnel_page(search=None, durumlar=None, gruplar=None, sadece_uyarili=False,
                       sadece_puantajsiz=False, limit=25, offset=0, db_path=None):
    tc_set = None
    if sadece_uyarili:
        tc_set = get_open_alert_tcs(db_path)
    where, params = _personel_where(search, durumlar, gruplar, tc_set)
    if sadece_puantajsiz:
        where += " AND NOT EXISTS (SELECT 1 FROM gunluk_puantaj g WHERE g.tc = p.tc)"
    conn = sema.get_connection(db_path)
    try:
        toplam = conn.execute(f"SELECT COUNT(*) FROM personnel p WHERE {where}", params).fetchone()[0]
        rows = conn.execute(f"SELECT p.* FROM personnel p WHERE {where} "
                            f"ORDER BY p.ad_soyad COLLATE NOCASE LIMIT ? OFFSET ?",
                            params + [limit, offset]).fetchall()
        return {"toplam": toplam, "kayitlar": [dict(r) for r in rows]}
    finally:
        conn.close()


def get_personnel_facets(db_path=None):
    conn = sema.get_connection(db_path)
    try:
        aktif_sql = "(isten_cikis_tarihi IS NULL OR isten_cikis_tarihi='')"
        aktif = conn.execute(f"SELECT COUNT(*) FROM personnel WHERE {aktif_sql}").fetchone()[0]
        toplam = conn.execute("SELECT COUNT(*) FROM personnel").fetchone()[0]
        gruplar = [dict(r) for r in conn.execute(
            "SELECT COALESCE(NULLIF(grup,''),'__bos__') AS grup, COUNT(*) AS adet FROM personnel "
            "GROUP BY COALESCE(NULLIF(grup,''),'__bos__') ORDER BY adet DESC").fetchall()]
        uyarili = conn.execute("SELECT COUNT(DISTINCT q.tc) FROM quality_alerts q JOIN personnel p ON p.tc=q.tc "
                               "WHERE q.cozuldu=0").fetchone()[0]
        puantajsiz = conn.execute("SELECT COUNT(*) FROM personnel p WHERE NOT EXISTS "
                                  "(SELECT 1 FROM gunluk_puantaj g WHERE g.tc=p.tc)").fetchone()[0]
        return {"toplam": toplam, "aktif": aktif, "ayrilmis": toplam - aktif, "gruplar": gruplar,
                "uyarili": uyarili, "puantajsiz": puantajsiz}
    finally:
        conn.close()


def get_all_personnel_brief(db_path=None):
    """Üst çubuktaki genel arama için hafif liste."""
    conn = sema.get_connection(db_path)
    try:
        return [dict(r) for r in conn.execute(
            "SELECT tc, ad_soyad, grup, gorevi FROM personnel ORDER BY ad_soyad COLLATE NOCASE").fetchall()]
    finally:
        conn.close()


def get_person_months(tc, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        return [r[0] for r in conn.execute("SELECT DISTINCT strftime('%Y-%m', tarih) FROM gunluk_puantaj "
                                           "WHERE tc=? ORDER BY 1", (tc,))]
    finally:
        conn.close()


def get_person_daily(tc, bas, bit, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        return [dict(r) for r in conn.execute(
            "SELECT tarih, normal_saat, fazla_saat, durum, kaynak_satir, import_batch_id FROM gunluk_puantaj "
            "WHERE tc=? AND tarih BETWEEN ? AND ? ORDER BY tarih", (tc, bas, bit)).fetchall()]
    finally:
        conn.close()


def get_person_year_fazla(tc, yil, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        return conn.execute("SELECT COALESCE(SUM(fazla_saat),0) FROM gunluk_puantaj WHERE tc=? "
                            "AND strftime('%Y', tarih)=?", (tc, str(yil))).fetchone()[0]
    finally:
        conn.close()


# ---------------------------------------------------------------- puantaj geçmişi (yeni)

def get_daily_hours_filtered(gruplar=None, ekipler=None, gorevler=None, tc=None,
                             bas=None, bit=None, limit=None, db_path=None):
    """get_daily_hours'un çoklu seçimli hali (grup/alt ekip/görev listeleri)."""
    conn = sema.get_connection(db_path)
    try:
        where, params = ["1=1"], []

        def _in(col, values, bos_etiket):
            vals = [v for v in values if v != bos_etiket]
            parts = []
            if vals:
                parts.append(f"{col} IN ({','.join('?' * len(vals))})")
                params.extend(vals)
            if bos_etiket in values:
                parts.append(f"({col} IS NULL OR {col}='')")
            where.append("(" + " OR ".join(parts) + ")")

        if gruplar:
            _in("p.grup", gruplar, "(Grup belirtilmemiş)")
        if ekipler:
            _in("p.bagli_oldugu_ekip", ekipler, "(Ekip belirtilmemiş)")
        if gorevler:
            _in("p.gorevi", gorevler, "(Görev belirtilmemiş)")
        if tc:
            where.append("gp.tc = ?"); params.append(tc)
        if bas:
            where.append("gp.tarih >= ?"); params.append(bas)
        if bit:
            where.append("gp.tarih <= ?"); params.append(bit)
        sql = (f"SELECT gp.tc, COALESCE(p.ad_soyad, '(personel kaydı yok)') AS ad_soyad, p.grup, "
               f"p.bagli_oldugu_ekip, p.gorevi, gp.tarih, gp.durum, gp.normal_saat, gp.fazla_saat, "
               f"(gp.normal_saat + gp.fazla_saat) AS toplam_saat "
               f"FROM gunluk_puantaj gp LEFT JOIN personnel p ON p.tc = gp.tc "
               f"WHERE {' AND '.join(where)} ORDER BY ad_soyad COLLATE NOCASE, gp.tarih")
        if limit:
            sql += f" LIMIT {int(limit)}"
        return [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()


# ---------------------------------------------------------------- v5: puantaj değişiklik geçmişi

def get_puantaj_history(tc, limit=500, db_path=None):
    """Bir kişinin günlük puantaj kayıtlarında olan değişiklikler (en yeni önce), içe aktarma dosya adıyla."""
    conn = sema.get_connection(db_path)
    try:
        return [dict(r) for r in conn.execute(
            "SELECT h.*, b.dosya_adi, b.durum AS batch_durum FROM puantaj_history h "
            "LEFT JOIN import_batches b ON b.id = h.import_batch_id "
            "WHERE h.tc=? ORDER BY h.id DESC LIMIT ?", (tc, int(limit))).fetchall()]
    finally:
        conn.close()


def get_puantaj_history_count(tc, db_path=None):
    conn = sema.get_connection(db_path)
    try:
        return conn.execute("SELECT COUNT(*) FROM puantaj_history WHERE tc=?", (tc,)).fetchone()[0]
    finally:
        conn.close()


def get_db_bilgisi(db_path=None):
    """Ayarlar > Sistem bilgisi için: dosya yolu, boyutlar, günlük kipi, kayıt sayıları."""
    from pathlib import Path
    yol = Path(db_path) if db_path else Path(sema.DB_PATH)
    wal = yol.with_name(yol.name + "-wal")
    conn = sema.get_connection(db_path)
    try:
        say = lambda t: conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        return {"yol": str(yol), "boyut": yol.stat().st_size if yol.exists() else 0,
                "wal_boyut": wal.stat().st_size if wal.exists() else 0,
                "kip": str(conn.execute("PRAGMA journal_mode").fetchone()[0]).lower(),
                "personel": say("personnel"), "gunluk_kayit": say("gunluk_puantaj"),
                "ice_aktarma": say("import_batches"), "puantaj_izi": say("puantaj_history")}
    finally:
        conn.close()


def get_person_calisma_ozeti(tc, db_path=None):
    """Personel Kartı için: ilk / son kayıt ve çalışılan gün, toplam saatler (bütün zamanlar)."""
    conn = sema.get_connection(db_path)
    try:
        r = conn.execute(
            "SELECT MIN(tarih) AS ilk_kayit, MAX(tarih) AS son_kayit, "
            "MIN(CASE WHEN normal_saat + fazla_saat > 0 THEN tarih END) AS ilk_calisma, "
            "MAX(CASE WHEN normal_saat + fazla_saat > 0 THEN tarih END) AS son_calisma, "
            "SUM(CASE WHEN normal_saat + fazla_saat > 0 THEN 1 ELSE 0 END) AS calisilan_gun, "
            "COALESCE(SUM(normal_saat), 0) AS normal, COALESCE(SUM(fazla_saat), 0) AS fazla, COUNT(*) AS kayit "
            "FROM gunluk_puantaj WHERE tc=?", (tc,)).fetchone()
        return dict(r)
    finally:
        conn.close()
