# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - PERSONEL BİLGİLERİ İÇE AKTARMA (veri_merkezi_ice_aktarma.py)
===============================================================================
Ham, düzensiz kaynak excel'i okuyup doğrulanmış, kalıcı veri deposuna
(personel_veritabani.db > personnel) yazan yer. İki aşamalı çalışır:

  1) preview_import(dosya_yolu)   -> SADECE OKUR, hiçbir şey yazmaz.
     Kaç yeni/güncellenen personel, hangi uyarılar/engelleyiciler olduğunu
     döner. Arayüz bu özeti kullanıcıya gösterir.

  2) commit_import(preview_sonucu) -> preview'da üretilen veriyi TEK BİR
     SQLite transaction'ında yazar. Ya TAMAMEN başarılı olur ya da hiçbir
     şey yazılmaz (yarım kalmış/bozuk bir içe aktarma asla iz bırakmaz).

SADECE KİMLİK: ad soyad, TC, grup, bağlı olduğu ekip, görev, işe giriş/çıkış
tarihi. Kaynaktaki her TC iki satır halinde (NORMAL MESAİ + FAZLA MESAİ)
görünse de, kimlik bilgisi bu iki satırda AYNI olduğundan TEK SEFER (ilk
görülen satırdan) alınır - kişi başına TEK kayıt.

ÇALIŞMA SAATİ (mesai) BU DOSYADAN ARTIK ÇEKİLMEZ - o, ayrı bir günlük SAP
roster dosyasından gelir (bkz. veri_merkezi_mesai_ice_aktarma.py). Bu ayrım
sayesinde bu modül artık gün sütunlarını hiç taramaz.

Sütunlar SABİT harf/satır numarasıyla DEĞİL, başlık metni taranarak tespit
edilir (veri_merkezi_dogrulama.py). Bu, "dosya büyüdü, sütun kaydı" türü
hataların kökten önüne geçer.

v5: yazmadan önce veritabanının yedeği alınır; eklenen kişiler ve değişen alanlar
personel_degisiklik_izi tablosuna da yazılır, böylece içe aktarma geri alınabilir
(veri_merkezi_geri_alma).
"""

__version__ = "2026-09-25.1"


import datetime
import hashlib
from pathlib import Path

import openpyxl

import puantaj_engine as engine
import veri_merkezi_sema as sema
import veri_merkezi_dogrulama as dogrulama


class ImportError_(Exception):
    """İçe aktarma sırasında ENGELLEYİCİ bir sorun (baslik_eslesmedi, dosya_tekrar vb.)."""
    pass


def _file_hash(dosya_yolu):
    h = hashlib.sha256()
    with open(dosya_yolu, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _parse_date_cell(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def _pick_sheet(wb, sheet_name=None):
    """Grafik sayfaları (Chartsheet) veri içermez ve max_column/max_row gibi
    temel özellikleri bile yoktur - otomatik seçimde daima ATLANIR. Belirli
    bir isim verilmişse ve o bir çalışma sayfasıysa (Chartsheet değilse) o
    kullanılır; verilmemişse ilk GERÇEK çalışma sayfası seçilir."""
    from openpyxl.worksheet.worksheet import Worksheet

    if sheet_name and sheet_name in wb.sheetnames:
        candidate = wb[sheet_name]
        if isinstance(candidate, Worksheet):
            return candidate
    for name in wb.sheetnames:
        candidate = wb[name]
        if isinstance(candidate, Worksheet):
            return candidate
    raise ImportError_("Dosyada okunabilir bir çalışma sayfası (grafik sayfası değil) bulunamadı.")


def _existing_file_hashes(db_path=None):
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute("SELECT DISTINCT dosya_hash FROM import_batches WHERE durum='tamamlandi'").fetchall()
        return {r["dosya_hash"] for r in rows}
    finally:
        conn.close()


def _last_completed_count(db_path=None):
    conn = sema.get_connection(db_path)
    try:
        row = conn.execute(
            "SELECT yeni_personel + guncellenen_personel AS n FROM import_batches "
            "WHERE durum='tamamlandi' ORDER BY id DESC LIMIT 1"
        ).fetchone()
        return row["n"] if row else None
    finally:
        conn.close()


def preview_import(dosya_yolu, sheet_name=None, db_path=None, max_saat=24, colmap_override=None):
    """Dosyayı okur, doğrular, SADECE bir önizleme sözlüğü döner. Veritabanına
    hiçbir şey yazmaz."""
    sema.migrate(db_path)
    import personnel_db as pdb
    pdb.init_db(db_path)  # "personnel" tablosu yoksa oluşturur (sadece YAPI, satır yazmaz)
    dosya_yolu = Path(dosya_yolu)
    dosya_hash = _file_hash(dosya_yolu)

    wb = openpyxl.load_workbook(dosya_yolu, data_only=True)
    ws = _pick_sheet(wb, sheet_name)

    colmap = dogrulama.build_column_map(ws)
    if colmap_override:            # kullanıcının elle seçtiği sütunlar (alan -> sütun no) başlık eşleşmesine üstün gelir
        colmap.update({k: v for k, v in colmap_override.items() if v})
    engelleyiciler = list(dogrulama.kural_baslik_eslesmesi(colmap))

    dosya_tekrar_uyarilari = dogrulama.kural_dosya_tekrar(dosya_hash, _existing_file_hashes(db_path))
    engelleyiciler.extend(a for a in dosya_tekrar_uyarilari if a["onem"] == "engelleyici")

    if engelleyiciler:
        return {
            "durum": "engellendi",
            "dosya_adi": dosya_yolu.name,
            "dosya_hash": dosya_hash,
            "sheet_name": ws.title,
            "engelleyiciler": engelleyiciler,
            "uyarilar": [],
            "personel_kayitlari": [],
            "yeni_personel": 0,
            "guncellenen_personel": 0,
            "donem_baslangic": None,
            "donem_bitis": None,
        }

    data_start, data_end = dogrulama.detect_data_row_range(ws, colmap["tc"])

    uyarilar = []
    all_rows = []          # kural_ayni_tc_farkli_isim için
    personel_map = {}      # tc -> {ad_soyad, grup, alt_ekip, gorevi, giris, cikis}
    taranan_satir = 0
    giris_gorulen = set()  # ilk görülen kaynak satırı (izlenebilirlik için)

    for r in range(data_start, data_end + 1):
        tc = engine.normalize_tc(ws.cell(row=r, column=colmap["tc"]).value) if "tc" in colmap else None
        name = ws.cell(row=r, column=colmap["ad_soyad"]).value if "ad_soyad" in colmap else None
        name = str(name).strip() if name else None
        grup = ws.cell(row=r, column=colmap["grup"]).value if "grup" in colmap else None
        grup = str(grup).strip() if grup else None
        alt_ekip = ws.cell(row=r, column=colmap.get("alt_ekip", 0)).value if "alt_ekip" in colmap else None
        alt_ekip = str(alt_ekip).strip() if alt_ekip else None
        gorevi = ws.cell(row=r, column=colmap.get("gorevi", 0)).value if "gorevi" in colmap else None
        gorevi = str(gorevi).strip() if gorevi else None

        # Tamamen boş satır (ne TC ne isim) - veri değil, sessizce atla.
        if not tc and not name:
            continue
        taranan_satir += 1

        giris = _parse_date_cell(ws.cell(row=r, column=colmap["giris_tarihi"]).value) if "giris_tarihi" in colmap else None
        cikis = _parse_date_cell(ws.cell(row=r, column=colmap["cikis_tarihi"]).value) if "cikis_tarihi" in colmap else None

        row_ctx = {
            "tc": tc, "ad_soyad": name, "grup": grup,
            "kaynak_satir": r, "giris_tarihi": giris, "cikis_tarihi": cikis,
        }

        if not tc:
            uyarilar.append({"tip": "tc_eksik", "onem": "uyari", "tc": None, "ad_soyad": name,
                              "kaynak_satir": r,
                              "detay": f"Satır {r}: {name or '(isim de belirtilmemiş)'} - TC boş, "
                                       f"bu satırdan kimlik bilgisi alınamadı."})
            continue

        for kural in (dogrulama.kural_eksik_kimlik_bilgisi, dogrulama.kural_tc_format,
                      dogrulama.kural_giris_cikis_tutarsiz):
            sonuc = kural(row_ctx)
            if sonuc:
                uyarilar.append(sonuc)

        # HER satır kural_ayni_tc_farkli_isim için izlenir (aynı TC, 2 satırda -
        # normal/fazla - AYNI isimle geliyorsa zararsız tekrar; FARKLI isimle
        # geliyorsa bu, o kuralın yakalaması gereken gerçek bir çakışmadır).
        # Kimlik ise TEK SEFER, İLK görülen satırdan personel_map'e alınır.
        all_rows.append(row_ctx)
        if tc not in personel_map:
            personel_map[tc] = {"ad_soyad": name or "(İsimsiz)", "grup": grup, "alt_ekip": alt_ekip,
                                 "gorevi": gorevi, "giris_tarihi": giris, "cikis_tarihi": cikis}

    uyarilar.extend(dogrulama.kural_ayni_tc_farkli_isim(all_rows))
    onceki_sayi = _last_completed_count(db_path)
    uyarilar.extend(dogrulama.kural_personel_sayisi_anomalisi(len(personel_map), onceki_sayi))

    # --- Aynı kişi, farklı ay dosyalarında farklı bilgilerle kayıtlı olabilir mi? ---
    # (Bu içe aktarmadaki her TC, veritabanında ZATEN var olan kayıtla karşılaştırılır.
    # Farklıysa OTOMATİK üzerine YAZILMAZ - kullanıcıya sorulacak "field_conflicts"
    # listesine eklenir. Bkz. commit_import(field_resolutions=...).)
    field_conflicts = []
    CONFLICT_FIELDS = [("ad_soyad", "Ad Soyad"), ("grup", "Grup"),
                        ("bagli_oldugu_ekip", "Bağlı Olduğu Ekip"), ("gorevi", "Görevi")]
    if personel_map:
        conn = sema.get_connection(db_path)
        try:
            tcs = list(personel_map.keys())
            placeholders = ",".join("?" * len(tcs))
            existing_rows = conn.execute(
                f"SELECT * FROM personnel WHERE tc IN ({placeholders})", tcs
            ).fetchall()
            existing_by_tc = {r["tc"]: dict(r) for r in existing_rows}
        finally:
            conn.close()

        for tc, bilgi in personel_map.items():
            existing = existing_by_tc.get(tc)
            if not existing:
                continue
            yeni_degerler = {"ad_soyad": bilgi["ad_soyad"], "grup": bilgi["grup"],
                              "bagli_oldugu_ekip": bilgi["alt_ekip"], "gorevi": bilgi["gorevi"]}
            for db_alan, etiket in CONFLICT_FIELDS:
                eski = (existing.get(db_alan) or "").strip()
                yeni = (yeni_degerler.get(db_alan) or "").strip()
                if yeni and eski and yeni != eski:
                    field_conflicts.append({
                        "tc": tc, "alan": db_alan, "alan_etiket": etiket,
                        "ad_soyad_baglam": bilgi["ad_soyad"] or existing.get("ad_soyad") or tc,
                        "eski_deger": existing.get(db_alan),
                        "eski_kaynak": existing.get("kaynak") or "(bilinmiyor)",
                        "yeni_deger": yeni_degerler.get(db_alan),
                        "yeni_kaynak": f"Bu içe aktarma: {dosya_yolu.name}",
                    })

        # --- Olası "aynı kişi, farklı TC" durumu (bilgilendirme amaçlı, otomatik
        # birleştirilmez - TC farklıysa hangi kaydın doğru/güncel olduğunu yazılım
        # güvenle tahmin edemez, insan kararı gerekir). ---
        conn = sema.get_connection(db_path)
        try:
            tum_mevcut = conn.execute("SELECT tc, ad_soyad FROM personnel").fetchall()
        finally:
            conn.close()
        mevcut_isim_to_tc = {}
        for r in tum_mevcut:
            if r["ad_soyad"]:
                mevcut_isim_to_tc.setdefault(engine.normalize_name(r["ad_soyad"]), set()).add(r["tc"])

        for tc, bilgi in personel_map.items():
            if tc in existing_by_tc or not bilgi["ad_soyad"]:
                continue
            benzer_tcler = mevcut_isim_to_tc.get(engine.normalize_name(bilgi["ad_soyad"]), set()) - {tc}
            if benzer_tcler:
                uyarilar.append({
                    "tip": "olasi_farkli_tc", "onem": "uyari", "tc": tc,
                    "ad_soyad": bilgi["ad_soyad"], "kaynak_satir": None,
                    "detay": f"{bilgi['ad_soyad']}: bu içe aktarmada TC {tc} ile geliyor, ama veri "
                             f"setinde AYNI İSİMLE farklı TC ile kayıtlı kişi(ler) var: "
                             f"{', '.join(sorted(benzer_tcler))}. Aynı kişi olup TC'lerden biri "
                             f"hatalı yazılmış olabilir - otomatik birleştirilmedi, elle kontrol gerekir."
                })

    # Bu artık kimlik-only bir içe aktarma - "dönem" kavramı yok (o, mesai
    # içe aktarımına ait). İzlenebilirlik için işe GİRİŞ tarihi aralığı bilgi
    # amaçlı tutulur (kayıt tarihi değil, en erken/en geç işe giriş tarihi).
    giris_tarihleri = [b["giris_tarihi"] for b in personel_map.values() if b["giris_tarihi"]]

    return {
        "durum": "hazir",
        "dosya_adi": dosya_yolu.name,
        "dosya_hash": dosya_hash,
        "sheet_name": ws.title,
        "engelleyiciler": [],
        "uyarilar": uyarilar,
        "personel_kayitlari": personel_map,
        "field_conflicts": field_conflicts,
        "yeni_personel": None,   # commit sırasında personnel_db ile karşılaştırılınca hesaplanır
        "guncellenen_personel": None,
        "taranan_satir": taranan_satir,
        "donem_baslangic": min(giris_tarihleri).isoformat() if giris_tarihleri else None,
        "donem_bitis": max(giris_tarihleri).isoformat() if giris_tarihleri else None,
    }


def commit_import(preview, db_path=None, field_resolutions=None):
    """preview_import()'un ürettiği veriyi TEK bir transaction'da kalıcı
    depoya yazar. preview["durum"] != "hazir" ise çağrılmamalıdır.

    field_resolutions: preview["field_conflicts"]'te listelenen her çakışma
    için kullanıcının kararı - {(tc, alan): "existing"|"new"}. "existing"
    seçilirse o alan İÇİN mevcut değer korunur, bu içe aktarmadaki değer
    YOK SAYILIR. Bir çakışma için karar VERİLMEMİŞSE (dict'te yoksa),
    geriye dönük uyumluluk için varsayılan "new" (bu içe aktarmanın değeri
    kazanır) - arayüz katmanı HER çakışma için bir karar göndermelidir.

    NOT: personnel_db.upsert_personnel() KENDİ bağlantısını açıp anında commit
    ettiği için burada KULLANILMAZ - bu, "ya hep ya hiç" garantisini bozar
    (personel güncellenir ama günlük kayıt satırı hata verirse personel
    değişikliği geri alınamaz kalırdı). Bunun yerine personel upsert mantığı,
    BU fonksiyonun tek transaction'ı içinde, doğrudan aynı bağlantıyla
    tekrarlanır."""
    if preview["durum"] != "hazir":
        raise ImportError_("Bu önizleme içe aktarılamaz durumda (engelleyici sorunlar var).")

    field_resolutions = field_resolutions or {}
    sema.migrate(db_path)
    import personnel_db as pdb
    pdb.init_db(db_path)  # "personnel"/"personnel_history" tabloları yoksa oluşturur

    import veri_merkezi_yedek as yedek
    yedek_bilgisi = yedek.islem_oncesi_yedek("ice_aktarma", db_path)
    conn = sema.get_connection(db_path)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    kaynak_etiketi = f"veri_merkezi:{preview['dosya_adi']}"
    try:
        cur = conn.execute(
            "INSERT INTO import_batches (dosya_adi, dosya_hash, donem_baslangic, donem_bitis, "
            "baslama_zamani, durum, tur, izli) VALUES (?,?,?,?,?, 'basladi', 'kimlik', 1)",
            (preview["dosya_adi"], preview["dosya_hash"], preview["donem_baslangic"],
             preview["donem_bitis"], now),
        )
        batch_id = cur.lastrowid

        yeni = guncellenen = 0
        for tc, bilgi in preview["personel_kayitlari"].items():
            row = conn.execute("SELECT * FROM personnel WHERE tc=?", (tc,)).fetchone()
            yeni_alanlar = {
                "ad_soyad": bilgi["ad_soyad"], "grup": bilgi["grup"],
                "bagli_oldugu_ekip": bilgi["alt_ekip"], "gorevi": bilgi["gorevi"],
                "ise_giris_tarihi": bilgi["giris_tarihi"].isoformat() if bilgi["giris_tarihi"] else None,
                "isten_cikis_tarihi": bilgi["cikis_tarihi"].isoformat() if bilgi["cikis_tarihi"] else None,
            }
            if row is None:
                conn.execute(
                    "INSERT INTO personnel (tc, ad_soyad, grup, bagli_oldugu_ekip, gorevi, "
                    "ise_giris_tarihi, isten_cikis_tarihi, notlar, kaynak, created_at, updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    (tc, yeni_alanlar["ad_soyad"], yeni_alanlar["grup"], yeni_alanlar["bagli_oldugu_ekip"],
                     yeni_alanlar["gorevi"], yeni_alanlar["ise_giris_tarihi"], yeni_alanlar["isten_cikis_tarihi"],
                     "", kaynak_etiketi, now, now),
                )
                conn.execute(
                    "INSERT INTO personel_degisiklik_izi (import_batch_id, tc, islem, zaman) "
                    "VALUES (?,?, 'ekle', ?)", (batch_id, tc, now),
                )
                yeni += 1
            else:
                degisti = False
                for alan, yeni_deger in yeni_alanlar.items():
                    if not yeni_deger:
                        continue
                    eski_deger = row[alan]
                    if (eski_deger or "") != (yeni_deger or ""):
                        secim = field_resolutions.get((tc, alan), "new")
                        if secim == "existing":
                            continue  # kullanıcı mevcut değeri korumayı seçti - hiçbir şey değişmez
                        conn.execute(f"UPDATE personnel SET {alan}=? WHERE tc=?", (yeni_deger, tc))
                        conn.execute(
                            "INSERT INTO personnel_history (tc, field, old_value, new_value, changed_at, source) "
                            "VALUES (?,?,?,?,?,?)", (tc, alan, eski_deger, yeni_deger, now, kaynak_etiketi),
                        )
                        conn.execute(
                            "INSERT INTO personel_degisiklik_izi (import_batch_id, tc, islem, alan, eski_deger, "
                            "yeni_deger, eski_kaynak, zaman) VALUES (?,?, 'guncelle', ?,?,?,?,?)",
                            (batch_id, tc, alan, eski_deger, yeni_deger, row["kaynak"], now),
                        )
                        degisti = True
                if degisti:
                    conn.execute("UPDATE personnel SET updated_at=?, kaynak=? WHERE tc=?", (now, kaynak_etiketi, tc))
                    guncellenen += 1

        engelleyici_sayisi = sum(1 for a in preview["uyarilar"] if a["onem"] == "engelleyici")
        for alert in preview["uyarilar"]:
            conn.execute(
                "INSERT INTO quality_alerts (import_batch_id, tip, onem, tc, ad_soyad, kaynak_satir, "
                "detay, olusturma_zamani) VALUES (?,?,?,?,?,?,?,?)",
                (batch_id, alert["tip"], alert["onem"], alert.get("tc"), alert.get("ad_soyad"),
                 alert.get("kaynak_satir"), alert["detay"], now),
            )

        conn.execute(
            "UPDATE import_batches SET tamamlanma_zamani=?, durum='tamamlandi', taranan_satir=?, "
            "yeni_personel=?, guncellenen_personel=?, yazilan_gun_kaydi=0, "
            "engelleyici_sayisi=?, uyari_sayisi=? WHERE id=?",
            (now, preview.get("taranan_satir", 0), yeni, guncellenen,
             engelleyici_sayisi, len(preview["uyarilar"]), batch_id),
        )
        conn.commit()
        return {
            "batch_id": batch_id, "yeni_personel": yeni, "guncellenen_personel": guncellenen,
            "uyari_sayisi": len(preview["uyarilar"]),
            "yedek": yedek_bilgisi["ad"] if yedek_bilgisi else None,
        }
    except Exception as e:
        conn.rollback()
        conn.execute(
            "INSERT INTO import_batches (dosya_adi, dosya_hash, baslama_zamani, tamamlanma_zamani, "
            "durum, hata_mesaji, tur) VALUES (?,?,?,?, 'hata', ?, 'kimlik')",
            (preview["dosya_adi"], preview["dosya_hash"], now, now, str(e)),
        )
        conn.commit()
        raise
    finally:
        conn.close()
