# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - İÇE AKTARMAYI GERİ ALMA (veri_merkezi_geri_alma.py)
=====================================================================
Yanlış dosya yüklendiğinde elle temizlik yerine, içe aktarmanın yaptığı değişiklikleri
tek işlemde geri döndürür. v5 ile her içe aktarma yaptığı değişikliklerin izini tutar:
  - mesai     : puantaj_history        (eklenen / eski değeri değiştirilen günler)
  - kimlik    : personel_degisiklik_izi (eklenen kişiler / değişen alanlar)

İki adım:
  onizle(batch_id)  -> SADECE okur: neyin silineceğini, neyin eski değerine döneceğini,
                       neyin korunacağını ve geri almayı engelleyen bir şey olup olmadığını söyler.
  geri_al(batch_id) -> Önce veritabanının yedeğini alır, sonra aynı planı TEK bir
                       transaction içinde yeniden hesaplayıp uygular (ya hep ya hiç).

GÜVENLİK KURALLARI (verinin tutarlı kalması için):
  1) Sonradan yapılmış bir içe aktarma AYNI günleri / AYNI alanları değiştirmişse geri alma
     ENGELLENİR ve önce o (daha yeni) içe aktarmanın geri alınması istenir. Böylece geri alma
     her zaman "en yeniden eskiye" doğru yapılır ve hiçbir değer yanlış bir ara duruma dönmez.
  2) Personel bilgisi sonradan ELLE değiştirilmişse (Personel Düzenle) o alan KORUNUR;
     elle yapılan düzeltme geri almayla kaybolmaz.
  3) İçe aktarmayla eklenmiş bir kişi ancak (a) puantaj kaydı yoksa ve (b) sonradan elle
     düzenlenmemişse silinir; aksi hâlde kayıt korunur ve özet bunu söyler.
  4) Bu içe aktarmanın ürettiği açık veri kalitesi uyarıları "çözüldü" olarak kapatılır
     (silinmez; not: "İçe aktarma geri alındı").
  5) İçe aktarma kaydı silinmez; durumu 'geri_alindi' olur, zamanı, notu ve özeti saklanır.
     Aynı dosya daha sonra yeniden içe aktarılabilir.
  6) v5 öncesi içe aktarmaların izi tutulmadığı için bunlar geri alınamaz.
"""

import datetime
import json
import logging

import veri_merkezi_sema as sema

gunluk = logging.getLogger("veri_merkezi")

ORNEK_SAYISI = 8   # önizlemede gösterilecek örnek satır sayısı


class GeriAlmaHatasi(Exception):
    pass


def _simdi():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _batch(conn, batch_id):
    r = conn.execute("SELECT * FROM import_batches WHERE id=?", (batch_id,)).fetchone()
    return dict(r) if r else None


def _tur(b):
    if b.get("tur"):
        return b["tur"]
    if (b.get("yazilan_gun_kaydi") or 0) > 0:
        return "mesai"
    if (b.get("yeni_personel") or 0) > 0 or (b.get("guncellenen_personel") or 0) > 0:
        return "kimlik"
    return None


def _batch_ozetleri(conn, idler):
    """Engelleyen içe aktarmaların kısa bilgisi (en yeni önce)."""
    idler = sorted({i for i in idler if i is not None}, reverse=True)
    if not idler:
        return []
    rows = conn.execute(f"SELECT id, dosya_adi, baslama_zamani, durum FROM import_batches "
                        f"WHERE id IN ({','.join('?' * len(idler))}) ORDER BY id DESC", idler).fetchall()
    return [dict(r) for r in rows]


def _isimler(conn, tcs):
    tcs = list({t for t in tcs if t})
    sonuc = {}
    for i in range(0, len(tcs), 500):
        parca = tcs[i:i + 500]
        for r in conn.execute(f"SELECT tc, ad_soyad FROM personnel WHERE tc IN ({','.join('?' * len(parca))})", parca):
            sonuc[r["tc"]] = r["ad_soyad"]
    return sonuc


# ---------------------------------------------------------------- plan: mesai

def _mesai_plani(conn, batch_id):
    rows = conn.execute(
        "SELECT h.*, g.id AS g_id, g.normal_saat AS c_normal, g.fazla_saat AS c_fazla, g.durum AS c_durum, "
        "g.import_batch_id AS c_batch "
        "FROM puantaj_history h LEFT JOIN gunluk_puantaj g ON g.tc = h.tc AND g.tarih = h.tarih "
        "WHERE h.import_batch_id=? AND h.geri_alma=0 ORDER BY h.tarih, h.tc", (batch_id,)).fetchall()
    plan = {"silinecek": [], "geri_yuklenecek": [], "zaten_yok": [], "engelleyen_idler": set(),
            "cakisan": []}
    for r in rows:
        r = dict(r)
        if r["g_id"] is None:
            plan["zaten_yok"].append(r)          # kayıt artık yok: yapılacak bir şey kalmamış
        elif r["c_batch"] != batch_id:
            plan["engelleyen_idler"].add(r["c_batch"])
            plan["cakisan"].append(r)
        elif r["islem"] == "ekle":
            plan["silinecek"].append(r)
        else:
            plan["geri_yuklenecek"].append(r)
    return plan


# ---------------------------------------------------------------- plan: kimlik

def _kimlik_plani(conn, batch_id, dosya_adi):
    izler = [dict(r) for r in conn.execute(
        "SELECT * FROM personel_degisiklik_izi WHERE import_batch_id=? AND geri_alma=0 ORDER BY id",
        (batch_id,)).fetchall()]
    plan = {"alan_geri_yuklenecek": [], "alan_elle_degismis": [], "silinecek_kisi": [], "korunan_kisi": [],
            "zaten_yok": [], "engelleyen_idler": set(), "cakisan": []}

    def sonraki_batchler(tc, alan=None):
        sql = ("SELECT DISTINCT i.import_batch_id FROM personel_degisiklik_izi i "
               "JOIN import_batches b ON b.id = i.import_batch_id "
               "WHERE i.tc=? AND i.import_batch_id > ? AND i.geri_alma=0 AND b.durum='tamamlandi'")
        params = [tc, batch_id]
        if alan is not None:
            sql += " AND (i.alan=? OR i.islem='ekle')"
            params.append(alan)
        return {r[0] for r in conn.execute(sql, params)}

    for iz in izler:
        tc = iz["tc"]
        kisi = conn.execute("SELECT * FROM personnel WHERE tc=?", (tc,)).fetchone()
        if iz["islem"] == "guncelle":
            sonraki = sonraki_batchler(tc, iz["alan"])
            if sonraki:
                plan["engelleyen_idler"] |= sonraki
                plan["cakisan"].append(iz)
            elif kisi is None:
                plan["zaten_yok"].append(iz)
            elif (kisi[iz["alan"]] or "") == (iz["yeni_deger"] or ""):
                iz["simdiki"] = kisi[iz["alan"]]
                plan["alan_geri_yuklenecek"].append(iz)
            else:
                iz["simdiki"] = kisi[iz["alan"]]
                plan["alan_elle_degismis"].append(iz)
        elif iz["islem"] == "ekle":
            sonraki = sonraki_batchler(tc)
            if sonraki:
                plan["engelleyen_idler"] |= sonraki
                plan["cakisan"].append(iz)
                continue
            if kisi is None:
                plan["zaten_yok"].append(iz)
                continue
            puantaj_n = conn.execute("SELECT COUNT(*) FROM gunluk_puantaj WHERE tc=?", (tc,)).fetchone()[0]
            # "Elle" = Veri Merkezi içe aktarmaları DIŞINDAN gelen değişiklik (Personel Düzenle, Puantaj Suite).
            # Başka bir (etkin) içe aktarmanın değişikliği zaten yukarıda "engelleyen" olarak yakalandı;
            # geri alınmış içe aktarmaların ve geri almaların kayıtları elle sayılmaz.
            elle_n = conn.execute(
                "SELECT COUNT(*) FROM personnel_history WHERE tc=? AND "
                "(COALESCE(source,'') NOT LIKE 'veri_merkezi:%' OR source='veri_merkezi:manuel')",
                (tc,)).fetchone()[0]
            if puantaj_n:
                iz["neden"] = f"{puantaj_n} günlük puantaj kaydı var"
                plan["korunan_kisi"].append(iz)
            elif elle_n:
                iz["neden"] = "sonradan elle düzenlenmiş"
                plan["korunan_kisi"].append(iz)
            else:
                plan["silinecek_kisi"].append(iz)
    return plan


# ---------------------------------------------------------------- önizleme

def _plan(conn, batch_id):
    b = _batch(conn, batch_id)
    sonuc = {"batch_id": batch_id, "batch": b, "tur": None, "izin": False, "neden": None,
             "engelleyenler": [], "sayilar": {}, "ornekler": {}, "_plan": None}
    if b is None:
        sonuc["neden"] = "İçe aktarma kaydı bulunamadı."
        return sonuc
    tur = _tur(b)
    sonuc["tur"] = tur
    if b["durum"] == "geri_alindi":
        sonuc["neden"] = f"Bu içe aktarma zaten geri alındı ({b.get('geri_alma_zamani') or '?'})."
        return sonuc
    if b["durum"] == "hata":
        sonuc["neden"] = "Bu içe aktarma hatayla sonuçlanmıştı; veri setine hiçbir şey yazılmadı, geri alınacak bir şey yok."
        return sonuc
    if b["durum"] != "tamamlandi":
        sonuc["neden"] = "Bu içe aktarma tamamlanmamış (yarım kalmış); veri setine hiçbir şey yazılmadı."
        return sonuc
    if not b.get("izli"):
        sonuc["neden"] = ("Bu içe aktarma v5'ten önce yapıldığı için değiştirdiği eski değerler kayıtlı değil; "
                          "güvenle geri alınamaz. Gerekirse Ayarlar > Veritabanı yedekleri'nden o tarihten "
                          "önceki bir yedeği geri yükleyebilirsiniz.")
        return sonuc
    if tur not in ("mesai", "kimlik"):
        sonuc["neden"] = "İçe aktarma türü anlaşılamadı."
        return sonuc

    if tur == "mesai":
        plan = _mesai_plani(conn, batch_id)
        sonuc["sayilar"] = {"silinecek": len(plan["silinecek"]), "geri_yuklenecek": len(plan["geri_yuklenecek"]),
                            "zaten_yok": len(plan["zaten_yok"]), "cakisan": len(plan["cakisan"])}
        isim = _isimler(conn, [r["tc"] for k in ("silinecek", "geri_yuklenecek", "cakisan") for r in plan[k][:ORNEK_SAYISI]])
        sonuc["ornekler"] = {k: [dict(r, ad_soyad=isim.get(r["tc"])) for r in plan[k][:ORNEK_SAYISI]]
                             for k in ("silinecek", "geri_yuklenecek", "cakisan")}
        sonuc["kisi_sayisi"] = len({r["tc"] for k in ("silinecek", "geri_yuklenecek") for r in plan[k]})
    else:
        plan = _kimlik_plani(conn, batch_id, b["dosya_adi"])
        sonuc["sayilar"] = {"alan_geri_yuklenecek": len(plan["alan_geri_yuklenecek"]),
                            "alan_elle_degismis": len(plan["alan_elle_degismis"]),
                            "silinecek_kisi": len(plan["silinecek_kisi"]),
                            "korunan_kisi": len(plan["korunan_kisi"]),
                            "zaten_yok": len(plan["zaten_yok"]), "cakisan": len(plan["cakisan"])}
        anahtarlar = ("alan_geri_yuklenecek", "alan_elle_degismis", "silinecek_kisi", "korunan_kisi", "cakisan")
        isim = _isimler(conn, [r["tc"] for k in anahtarlar for r in plan[k][:ORNEK_SAYISI]])
        sonuc["ornekler"] = {k: [dict(r, ad_soyad=isim.get(r["tc"])) for r in plan[k][:ORNEK_SAYISI]]
                             for k in anahtarlar}
        sonuc["kisi_sayisi"] = len({r["tc"] for k in ("alan_geri_yuklenecek", "silinecek_kisi") for r in plan[k]})

    sonuc["acik_uyari"] = conn.execute("SELECT COUNT(*) FROM quality_alerts WHERE import_batch_id=? AND cozuldu=0",
                                       (batch_id,)).fetchone()[0]
    if plan["engelleyen_idler"]:
        sonuc["engelleyenler"] = _batch_ozetleri(conn, plan["engelleyen_idler"])
        idler = ", ".join(f"#{e['id']}" for e in sonuc["engelleyenler"]) or "(bilinmeyen)"
        sonuc["neden"] = (f"Bu içe aktarmadan SONRA yapılan içe aktarma(lar) aynı kayıtları değiştirmiş: {idler}. "
                          f"Verinin tutarlı kalması için önce onları (en yeniden başlayarak) geri alın.")
        return sonuc
    sonuc["izin"] = True
    sonuc["_plan"] = plan
    return sonuc


def onizle(batch_id, db_path=None):
    """Geri almanın ne yapacağını SADECE okuyarak hesaplar. Veritabanına hiçbir şey yazmaz."""
    sema.migrate(db_path)
    conn = sema.get_connection(db_path)
    try:
        sonuc = _plan(conn, batch_id)
    finally:
        conn.close()
    sonuc.pop("_plan", None)
    return sonuc


# ---------------------------------------------------------------- uygulama

def _mesai_uygula(conn, batch_id, plan, now):
    silinen = geri_yuklenen = 0
    for r in plan["silinecek"]:
        cur = conn.execute("DELETE FROM gunluk_puantaj WHERE tc=? AND tarih=? AND import_batch_id=?",
                           (r["tc"], r["tarih"], batch_id))
        if cur.rowcount != 1:
            raise GeriAlmaHatasi(f"{r['tc']} / {r['tarih']} kaydı beklenmedik biçimde değişmiş; işlem iptal edildi.")
        conn.execute(
            "INSERT INTO puantaj_history (tc, tarih, islem, eski_normal, eski_fazla, eski_durum, eski_batch_id, "
            "import_batch_id, geri_alma, zaman) VALUES (?,?, 'sil', ?,?,?,?,?, 1, ?)",
            (r["tc"], r["tarih"], r["c_normal"], r["c_fazla"], r["c_durum"], batch_id, batch_id, now))
        silinen += 1
    for r in plan["geri_yuklenecek"]:
        cur = conn.execute(
            "UPDATE gunluk_puantaj SET normal_saat=?, fazla_saat=?, durum=?, kaynak_satir=?, import_batch_id=? "
            "WHERE tc=? AND tarih=? AND import_batch_id=?",
            (r["eski_normal"] or 0, r["eski_fazla"] or 0, r["eski_durum"], r["eski_kaynak_satir"], r["eski_batch_id"],
             r["tc"], r["tarih"], batch_id))
        if cur.rowcount != 1:
            raise GeriAlmaHatasi(f"{r['tc']} / {r['tarih']} kaydı beklenmedik biçimde değişmiş; işlem iptal edildi.")
        conn.execute(
            "INSERT INTO puantaj_history (tc, tarih, islem, eski_normal, eski_fazla, eski_durum, eski_batch_id, "
            "yeni_normal, yeni_fazla, yeni_durum, yeni_kaynak_satir, import_batch_id, geri_alma, zaman) "
            "VALUES (?,?, 'guncelle', ?,?,?,?,?,?,?,?,?, 1, ?)",
            (r["tc"], r["tarih"], r["c_normal"], r["c_fazla"], r["c_durum"], batch_id, r["eski_normal"] or 0,
             r["eski_fazla"] or 0, r["eski_durum"], r["eski_kaynak_satir"], batch_id, now))
        geri_yuklenen += 1
    return {"silinen_gun": silinen, "geri_yuklenen_gun": geri_yuklenen}


def _kimlik_uygula(conn, batch_id, dosya_adi, plan, now):
    etiket = f"veri_merkezi:{dosya_adi}"
    kaynak_geri = f"veri_merkezi:geri_alma#{batch_id}"
    alan_n = 0
    kisiler = {}
    for iz in plan["alan_geri_yuklenecek"]:
        alan = iz["alan"]
        if alan not in ("ad_soyad", "grup", "bagli_oldugu_ekip", "gorevi", "ise_giris_tarihi", "isten_cikis_tarihi"):
            raise GeriAlmaHatasi(f"Beklenmeyen alan adı: {alan}")
        if alan == "ad_soyad" and not (iz["eski_deger"] or "").strip():
            raise GeriAlmaHatasi(f"{iz['tc']}: ad soyad boş bir değere döndürülemez.")
        conn.execute(f"UPDATE personnel SET {alan}=? WHERE tc=?", (iz["eski_deger"], iz["tc"]))
        conn.execute("INSERT INTO personnel_history (tc, field, old_value, new_value, changed_at, source) "
                     "VALUES (?,?,?,?,?,?)", (iz["tc"], alan, iz["simdiki"], iz["eski_deger"], now, kaynak_geri))
        conn.execute("INSERT INTO personel_degisiklik_izi (import_batch_id, tc, islem, alan, eski_deger, yeni_deger, "
                     "geri_alma, zaman) VALUES (?,?, 'geri_yukle', ?,?,?, 1, ?)",
                     (batch_id, iz["tc"], alan, iz["simdiki"], iz["eski_deger"], now))
        kisiler.setdefault(iz["tc"], iz["eski_kaynak"])
        alan_n += 1
    for tc, eski_kaynak in kisiler.items():
        conn.execute("UPDATE personnel SET updated_at=?, kaynak=CASE WHEN kaynak=? THEN ? ELSE kaynak END WHERE tc=?",
                     (now, etiket, eski_kaynak, tc))

    silinen_kisi = 0
    for iz in plan["silinecek_kisi"]:
        kisi = conn.execute("SELECT * FROM personnel WHERE tc=?", (iz["tc"],)).fetchone()
        if kisi is None:
            continue
        conn.execute("INSERT INTO personel_degisiklik_izi (import_batch_id, tc, islem, eski_deger, geri_alma, zaman) "
                     "VALUES (?,?, 'sil', ?, 1, ?)",
                     (batch_id, iz["tc"], json.dumps(dict(kisi), ensure_ascii=False), now))
        # personnel_db.delete_personnel ile aynı davranış: kişi silinince alan geçmişi de silinir
        # (kişinin tam kaydı yukarıda personel_degisiklik_izi'ne yazıldı, iz kaybolmaz).
        conn.execute("DELETE FROM personnel_history WHERE tc=?", (iz["tc"],))
        conn.execute("DELETE FROM personnel WHERE tc=?", (iz["tc"],))
        silinen_kisi += 1
    return {"geri_yuklenen_alan": alan_n, "guncellenen_kisi": len(kisiler), "silinen_kisi": silinen_kisi,
            "korunan_kisi": len(plan["korunan_kisi"]), "elle_degismis_alan": len(plan["alan_elle_degismis"])}


def geri_al(batch_id, not_="", db_path=None):
    """İçe aktarmayı geri alır. Başarısız olursa hiçbir şey değişmez (GeriAlmaHatasi).
    Döner: {"batch_id", "tur", "ozet": {...}, "yedek": ad|None}"""
    sema.migrate(db_path)
    on = onizle(batch_id, db_path)
    if not on["izin"]:
        raise GeriAlmaHatasi(on["neden"] or "Bu içe aktarma geri alınamaz.")

    import veri_merkezi_yedek as yedek
    yedek_bilgisi = yedek.islem_oncesi_yedek("geri_alma", db_path)

    conn = sema.get_connection(db_path)
    now = _simdi()
    try:
        conn.execute("BEGIN IMMEDIATE")          # başka bir yazma araya giremez; plan bu kilit altında yeniden hesaplanır
        sonuc = _plan(conn, batch_id)
        if not sonuc["izin"]:
            raise GeriAlmaHatasi(sonuc["neden"] or "Bu içe aktarma geri alınamaz.")
        b, plan, tur = sonuc["batch"], sonuc["_plan"], sonuc["tur"]
        if tur == "mesai":
            ozet = _mesai_uygula(conn, batch_id, plan, now)
            ozet["degismemis_atlanan"] = len(plan["zaten_yok"])
        else:
            ozet = _kimlik_uygula(conn, batch_id, b["dosya_adi"], plan, now)
        not_ = " ".join(str(not_ or "").split())[:500]
        cozum = "İçe aktarma geri alındı" + (f": {not_}" if not_ else "")
        cur = conn.execute("UPDATE quality_alerts SET cozuldu=1, cozum_notu=?, cozulme_zamani=? "
                           "WHERE import_batch_id=? AND cozuldu=0", (cozum, now, batch_id))
        ozet["kapatilan_uyari"] = cur.rowcount
        conn.execute("UPDATE import_batches SET durum='geri_alindi', geri_alma_zamani=?, geri_alma_notu=?, "
                     "geri_alma_ozeti=? WHERE id=? AND durum='tamamlandi'",
                     (now, not_ or None, json.dumps(ozet, ensure_ascii=False), batch_id))
        conn.commit()
    except Exception:
        conn.rollback()
        gunluk.exception("içe aktarma #%s geri alınamadı", batch_id)
        raise
    finally:
        conn.close()
    gunluk.warning("içe aktarma #%s GERİ ALINDI (%s): %s", batch_id, tur, ozet)
    return {"batch_id": batch_id, "tur": tur, "ozet": ozet, "yedek": yedek_bilgisi["ad"] if yedek_bilgisi else None}


def ozet_metni(tur, ozet):
    """geri_alma_ozeti sözlüğünü tek satırlık Türkçe metne çevirir (Geçmiş ekranı için)."""
    if not ozet:
        return ""
    if tur == "mesai":
        parca = [f"{ozet.get('silinen_gun', 0)} gün kaydı silindi",
                 f"{ozet.get('geri_yuklenen_gun', 0)} gün önceki değerine döndü"]
    else:
        parca = [f"{ozet.get('silinen_kisi', 0)} kişi silindi",
                 f"{ozet.get('geri_yuklenen_alan', 0)} alan önceki değerine döndü"]
        if ozet.get("korunan_kisi"):
            parca.append(f"{ozet['korunan_kisi']} kişi korundu")
        if ozet.get("elle_degismis_alan"):
            parca.append(f"elle değiştirilmiş {ozet['elle_degismis_alan']} alan korundu")
    if ozet.get("kapatilan_uyari"):
        parca.append(f"{ozet['kapatilan_uyari']} uyarı kapatıldı")
    return ", ".join(parca)
