# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - DOĞRULAMA KURALLARI KATALOĞU (veri_merkezi_dogrulama.py)
===========================================================================
Her kural bağımsız, tek başına test edilebilir bir fonksiyondur. İçe aktarma
motoru (veri_merkezi_ice_aktarma.py) bu kataloğu sırayla çalıştırır.

Yeni bir kural eklemek = bu dosyaya bir fonksiyon eklemek ve RULES listesine
eklemek; başka hiçbir yeri değiştirmeye gerek yoktur.

Her satır-bazlı kural şu şekli döner: dict | None
    {"tip": str, "onem": "engelleyici"|"uyari", "tc": str|None, "detay": str}
Toplu (import geneli) kurallar bir LİSTE alert döner.
"""

__version__ = "2026-09-20.4"


import re
import datetime

import puantaj_engine as engine

# ---------------------------------------------------------------- başlık sözlüğü / sütun tespiti

# Kaynak exceldeki alan -> başlık metninde aranacak alt dizeler.
# SABİT SÜTUN HARFİ KULLANILMAZ - başlık kayarsa/değişirse otomatik bulunur.
FIELD_HEADER_RULES = [
    ("tc", ["TC KİMLİK NO"], []),
    ("ad_soyad", ["AD SOYAD", "ADI SOYADI"], []),
    ("grup", ["GRUPLAR"], []),
    ("alt_ekip", ["BAĞLI OLDUĞU EKİP"], []),
    ("gorevi", ["GÖREVİ"], []),
    ("tip", ["ÇALIŞMA"], ["PERSONEL KIRILIMI"]),
    ("giris_tarihi", ["GİRİŞ TARİHİ"], []),
    ("cikis_tarihi", ["ÇIKIŞ TARİHİ"], []),
]

# NOT: "tip" (NORMAL/FAZLA MESAİ) artık ZORUNLU DEĞİL - bu içe aktarma SADECE
# kimlik bilgisi (ad soyad, grup, alt ekip, görev) çekiyor, saat ÇEKMİYOR (bkz.
# veri_merkezi_mesai_ice_aktarma.py - saatler artık AYRI bir dosyadan gelir).
# Kaynaktaki her TC için kaç satır olursa olsun (normal/fazla ayrımı önemsiz),
# kimlik bilgisi TEK SEFER, İLK görülen satırdan alınır.
REQUIRED_FIELDS = ["tc", "ad_soyad", "grup"]


def normalize_header(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\n", " ")).strip().upper()


def build_column_map(ws, header_rows=(3, 2, 1, 4)):
    """Başlık metnine göre {alan: sütun_no} eşlemesi çıkarır. Birden fazla
    satır taranır çünkü bazı kaynak dosyalarda başlık 2 satıra yayılabilir."""
    texts = {}
    for r in header_rows:
        for c in range(1, ws.max_column + 1):
            if c not in texts:
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    texts[c] = normalize_header(v)

    colmap = {}
    claimed = set()
    for key, any_of, none_of in FIELD_HEADER_RULES:
        for c in sorted(texts):
            if c in claimed:
                continue
            t = texts[c]
            if any(a in t for a in any_of) and not any(n in t for n in none_of):
                colmap[key] = c
                claimed.add(c)
                break
    return colmap


def detect_day_columns(ws, header_row=3):
    """Gün sütunlarını SABİT harf aralığı yerine, başlık satırında GERÇEK
    TARİH DEĞERİ bulunan hücreleri tarayarak tespit eder. Sütunlar kayarsa,
    eklenirse veya silinirse bile otomatik doğru bulunur."""
    cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, datetime.datetime):
            v = v.date()
        if isinstance(v, datetime.date):
            cols.append((c, v))
    return cols


def detect_data_row_range(ws, tc_col, start_hint=4):
    """Veri başlangıç/bitiş satırlarını SABİT sayı yerine, TC sütununda
    gerçek veri biten yeri tarayarak tespit eder. Dosya büyüdükçe (yeni
    personel eklendikçe) elle güncelleme ihtiyacını ORTADAN KALDIRIR -
    Ağustos ayı raporunda tam olarak bu yüzden 20 personel kaybolmuştu."""
    last_row = start_hint - 1
    for r in range(start_hint, ws.max_row + 1):
        if ws.cell(row=r, column=tc_col).value not in (None, ""):
            last_row = r
    return start_hint, max(last_row, ws.max_row)


# ---------------------------------------------------------------- satır bazlı kurallar

def kural_zorunlu_alan(row):
    """Bir günlük saat kaydı yazılabilmesi için SADECE 'tc' ve 'tip' zorunludur -
    TC olmadan bu saatlerin kime ait olduğu belirlenemez. İsim/grup gibi diğer
    alanlar eksikse satır YİNE DE YAZILIR (saat verisi kaybolmaz), sadece eksik
    alan '(bilinmiyor)' olarak işaretlenip ayrı bir uyarı üretilir."""
    if not row.get("tc"):
        isim = row.get("ad_soyad") or "(isim de belirtilmemiş)"
        return {"tip": "tc_eksik", "onem": "uyari", "tc": None,
                 "ad_soyad": row.get("ad_soyad"), "kaynak_satir": row.get("kaynak_satir"),
                 "detay": f"Satır {row.get('kaynak_satir')}: {isim} - TC boş, bu satırın saatleri "
                          f"kimseye atfedilemedi (bkz. 'TC'siz Kaybolan Saatler' özeti)."}
    if not row.get("tip"):
        return {"tip": "zorunlu_alan_eksik", "onem": "uyari", "tc": row.get("tc"),
                 "ad_soyad": row.get("ad_soyad"), "kaynak_satir": row.get("kaynak_satir"),
                 "detay": f"Satır {row.get('kaynak_satir')}: 'tip' (NORMAL/FAZLA MESAİ) alanı boş, satır atlandı."}
    return None


def kural_eksik_kimlik_bilgisi(row):
    """TC var ama isim/grup gibi kimlik alanları eksikse - saatler YİNE DE
    kaydedilir (kaybolmaz), ama bu durum ayrıca uyarı olarak işaretlenir."""
    eksikler = [f for f in ("ad_soyad", "grup") if not row.get(f)]
    if eksikler:
        return {"tip": "kimlik_bilgisi_eksik", "onem": "uyari", "tc": row.get("tc"),
                 "ad_soyad": row.get("ad_soyad"), "kaynak_satir": row.get("kaynak_satir"),
                 "detay": f"Satır {row.get('kaynak_satir')} (TC {row.get('tc')}): "
                          f"{', '.join(eksikler)} eksik, saatler yine de kaydedildi."}
    return None


def kural_tc_format(row):
    """TC formatı/checksum'ı hatalıysa uyarı (engellemez, veri yine de yazılır -
    çünkü sağlaması tutmayan bir TC'nin sağlaması tutan başka bir TC olduğunu
    varsaymak riskli bir tahmindir; insan kararı gerekir)."""
    tc = row.get("tc")
    if not tc:
        return None
    issue = engine.check_tc(tc, "kaynak", row.get("kaynak_satir"), row.get("ad_soyad"))
    if issue:
        return {"tip": "tc_format_hatali", "onem": "uyari", "tc": tc,
                 "ad_soyad": row.get("ad_soyad"), "kaynak_satir": row.get("kaynak_satir"),
                 "detay": f"Satır {issue['row']}: TC '{issue['tc']}' ({issue['name']}) -> {issue['reason']}"}
    return None


def kural_anormal_saat(row, max_saat=24):
    """Bir günlük saat değeri negatifse ya da mantıksız derecede yüksekse
    (24 saatten fazla) uyarı verir."""
    saat = row.get("saat")
    if saat is None:
        return None
    if saat < 0 or saat > max_saat:
        return {"tip": "anormal_saat", "onem": "uyari", "tc": row.get("tc"),
                 "ad_soyad": row.get("ad_soyad"), "kaynak_satir": row.get("kaynak_satir"),
                 "detay": f"Satır {row.get('kaynak_satir')}: {row.get('ad_soyad') or '(isimsiz)'} - "
                          f"{row.get('tarih')} tarihinde {saat} saat (beklenen aralık: 0-{max_saat})."}
    return None


def kural_giris_cikis_tutarsiz(row):
    """İşe giriş tarihi işten çıkış tarihinden SONRA ise (mantık hatası) uyarır."""
    giris, cikis = row.get("giris_tarihi"), row.get("cikis_tarihi")
    if giris and cikis and giris > cikis:
        return {"tip": "giris_cikis_tutarsiz", "onem": "uyari", "tc": row.get("tc"),
                 "ad_soyad": row.get("ad_soyad"), "kaynak_satir": row.get("kaynak_satir"),
                 "detay": f"{row.get('ad_soyad')} (TC {row.get('tc')}): işe giriş ({giris}) "
                          f"işten çıkıştan ({cikis}) sonra görünüyor."}
    return None


SATIR_KURALLARI = [kural_zorunlu_alan, kural_eksik_kimlik_bilgisi, kural_tc_format,
                    kural_anormal_saat, kural_giris_cikis_tutarsiz]


# ---------------------------------------------------------------- toplu (import geneli) kurallar

def kural_ayni_tc_farkli_isim(all_rows):
    """Aynı TC, farklı isimlerde geçiyorsa (Ağustos ayı raporunda bulduğumuz
    gerçek hata: 2 farklı kişi 1 TC paylaşıyordu) - OTOMATİK BİRLEŞTİRİLMEZ,
    her ikisi de ayrı ayrı kaydedilir ve bu uyarı üretilir. Her ismin İLK
    görüldüğü satır numarası da izlenir (hangi ismin hangi satırdan geldiği)."""
    tc_names = {}       # tc -> {isim: ilk_satir}
    for row in all_rows:
        tc, name = row.get("tc"), row.get("ad_soyad")
        if tc and name:
            d = tc_names.setdefault(tc, {})
            if name not in d:
                d[name] = row.get("kaynak_satir")

    alerts = []
    for tc, isim_satir in tc_names.items():
        if len(isim_satir) > 1:
            isimler_str = " / ".join(f"{isim} (satır {satir})" for isim, satir in sorted(isim_satir.items()))
            alerts.append({"tip": "ayni_tc_farkli_isim", "onem": "uyari", "tc": tc,
                             "ad_soyad": " / ".join(sorted(isim_satir.keys())), "kaynak_satir": None,
                             "detay": f"TC {tc}: {isimler_str} - aynı TC birden fazla farklı isimle "
                                      f"kayıtlı, gerçek kişi çakışması olabilir."})
    return alerts


def kural_dosya_tekrar(dosya_hash, gecmis_hashler):
    """Bu dosya (bayt bazında birebir aynı) daha önce içe aktarılmışsa
    ENGELLEYİCİ uyarı verir - kullanıcı onayı olmadan tekrar yazılmaz."""
    if dosya_hash in gecmis_hashler:
        return [{"tip": "dosya_tekrar", "onem": "engelleyici", "tc": None,
                   "detay": "Bu dosya (birebir aynı içerikle) daha önce içe aktarılmış."}]
    return []


def kural_personel_sayisi_anomalisi(yeni_sayi, onceki_sayi, esik_oran=0.3):
    """Personel sayısı önceki içe aktarmaya göre aniden çok değiştiyse
    (ör. dosya eksik okunmuş olabilir) uyarır."""
    if not onceki_sayi:
        return []
    fark_oran = abs(yeni_sayi - onceki_sayi) / onceki_sayi
    if fark_oran >= esik_oran:
        yon = "azaldı" if yeni_sayi < onceki_sayi else "arttı"
        return [{"tip": "personel_sayisi_anomalisi", "onem": "uyari", "tc": None,
                   "detay": f"Personel sayısı önceki içe aktarmaya göre %{fark_oran*100:.0f} {yon} "
                            f"({onceki_sayi} -> {yeni_sayi}). Dosya eksik okunmuş olabilir, kontrol edin."}]
    return []


def kural_baslik_eslesmesi(colmap):
    """ENGELLEYİCİ: kaynak dosyada beklenen sütun başlıkları bulunamazsa içe
    aktarma tamamen durur - yanlış sütundan sessizce okumak, hiç okumamaktan
    daha tehlikelidir."""
    eksikler = [f for f in REQUIRED_FIELDS if f not in colmap]
    if eksikler:
        return [{"tip": "baslik_eslesmedi", "onem": "engelleyici", "tc": None,
                   "detay": f"Kaynak dosyada şu alanların başlığı bulunamadı: {', '.join(eksikler)}. "
                            f"Dosya yapısı beklenenden farklı olabilir."}]
    return []
