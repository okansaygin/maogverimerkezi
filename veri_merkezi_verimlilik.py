# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - VERİMLİLİK ANALİZİ (veri_merkezi_verimlilik.py)
=================================================================
"Fazla Mesai Raporu"nun yerini alan yönetim raporu. Kalıcı depodaki günlük
puantajdan (gunluk_puantaj + personnel) iş gücü verimliliğini her açıdan ölçer:

  * Devam        : planlanan iş günlerinin ne kadarında çalışıldı
  * Kapasite     : planlanan normal saatin (gün × 9 sa) ne kadarı kullanıldı
  * Kayıp gün    : devamsızlık + rapor + ücretsiz izin
  * Fazla mesai  : saat, oran, pazar payı ve yevmiye (maliyet) etkisi
  * Uyum / risk  : yıllık 270 sa sınırı, günlük 11 saati aşan gün, 7+ gün
                   kesintisiz çalışma, aylık fazla mesai eşiği
  * Hareket      : dönemde işe giren / çıkan personel
  * Özet         : bütün bunlardan otomatik üretilen "yönetici özeti"

Ekran (veri_merkezi_sayfa_raporlar) ve Excel (veri_merkezi_verimlilik_excel)
AYNI sözlüğü kullanır: ekranda görülen her sayı Excel'de de aynıdır.

Katsayılar ve 9 saatlik standart gün veri_merkezi_sorgu ile aynıdır.
"""

__version__ = "2026-09-25.3"

import datetime
import functools
from pathlib import Path

import puantaj_engine as engine
import team_report as report
import veri_merkezi_ayarlar as ayarlar
import veri_merkezi_hafta_kurali as hk
import veri_merkezi_sema as sema
import veri_merkezi_sorgu as sorgu

# Eşikler ve katsayılar Ayarlar sayfasından değiştirilebilir; veri_merkezi_ayarlar.uygula() kaydedilince
# bu modül sabitlerini günceller ve analiz önbelleğini temizler.
STANDART_GUN = sorgu.STANDART_GUN_SAATI                  # 9 sa (sabit)
K_HI = sorgu.FAZLA_MESAI_KATSAYI_HAFTAICI                # varsayılan 1,5
K_PZ = sorgu.FAZLA_MESAI_KATSAYI_PAZAR                   # varsayılan 2,5
YILLIK_FM_SINIRI = ayarlar.al("yillik_fm_siniri")        # 270 - 4857 s. İş Kanunu md. 41
GUNLUK_AZAMI_SAAT = ayarlar.al("gunluk_azami_saat")      # 11 - md. 63 (günlük çalışma 11 saati aşamaz)
KESINTISIZ_GUN_ESIGI = ayarlar.al("kesintisiz_gun_esigi")  # 7 - md. 46: 7 günde en az 24 saat kesintisiz dinlenme
DEVAMSIZLIK_ESIGI = ayarlar.al("devamsizlik_esigi")      # 3 - dönemde bu kadar ve üzeri devamsızlık "risk"

NO_GRUP = "(Grup belirtilmemiş)"
NO_EKIP = report.NO_SUBTEAM
NO_GOREV = report.NO_ROLE

GUN_ADLARI = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
GUN_KISA = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]
AY_ADLARI = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım",
             "Aralık"]
AY_KISA = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"]

# gün kategorileri: (anahtar, etiket, renk, Excel kodu)
KATEGORILER = [
    ("calisti", "Çalıştı", "1F4E78", ""),
    ("ucretli_izin", "Ücretli izin", "5B3386", "İ"),
    ("ucretsiz_izin", "Ücretsiz izin", "8C99A6", "Ü"),
    ("rapor", "Rapor", "E8762C", "R"),
    ("devamsiz", "Devamsızlık", "B42318", "D"),
    ("hafta_tatili", "Hafta tatili", "C9D6E3", "HT"),
    ("resmi_tatil", "Resmî tatil", "8FAACB", "RT"),
    ("belirsiz", "Kayıt boş / diğer", "E3DED5", "?"),
]
KAT_ETIKET = {k: e for k, e, _r, _c in KATEGORILER}
KAT_RENK = {k: r for k, _e, r, _c in KATEGORILER}
KAT_KOD = {k: c for k, _e, _r, c in KATEGORILER}
KAYIP = ("devamsiz", "rapor", "ucretsiz_izin")           # kayıp iş günü
PLAN_DISI = ("hafta_tatili", "resmi_tatil", "belirsiz")  # planlanan iş gününe sayılmaz

PUAN_BANTLARI = [(95, "A", "Çok iyi"), (85, "B", "İyi"), (70, "C", "Orta"), (0, "D", "Zayıf")]


def _sa(x):
    """Etiketler için saat biçimi: 270 -> '270', 229.5 -> '229,5'."""
    x = float(x)
    return str(int(x)) if x.is_integer() else f"{x:.1f}".replace(".", ",")


def etiket_yillik():
    """'270' - yıllık fazla mesai sınırı (Ayarlar)."""
    return _sa(YILLIK_FM_SINIRI)


def etiket_yakin():
    """'229,5' - yıllık sınıra 'yaklaşan' eşiği (%85)."""
    return _sa(round(YILLIK_FM_SINIRI * 0.85, 1))


def etiket_gunluk():
    """'11' - günlük azami çalışma (Ayarlar)."""
    return _sa(GUNLUK_AZAMI_SAAT)


def esik_notu(deger, yasal):
    """Ayardaki eşik yasal değerden farklıysa yasa metninin yanına eklenen not."""
    return "" if float(deger) == float(yasal) else f" · uygulanan eşik {_sa(deger)} sa"


def katsayi_metni():
    """'(Normal + FM × 1,5; hak edilmiş pazar: (normal + FM) × 2,5) ÷ 9'"""
    return (f"(Normal + FM × {_sa(K_HI)}; hak edilmiş pazarda (normal + FM) × {_sa(K_PZ)}) ÷ {_sa(STANDART_GUN)}")


def katsayi_farki_metni():
    """'(FM × 0,5 + hak edilmiş pazar saatleri × 1,5) ÷ 9'"""
    return (f"(FM × {_sa(K_HI - 1)} + hak edilmiş pazar saatleri (normal + FM) × {_sa(K_PZ - 1)}) ÷ "
            f"{_sa(STANDART_GUN)}")


def pazar_kurali_metni():
    """Yöntem açıklaması: hafta tatili / pazar kuralı."""
    return ("Pazartesi-cumartesi fazla mesai ×1,5. Pazar, haftanın 6 günü tamamlandıysa (normal + FM) ×2,5; "
            "haftada ücretsiz izin ya da devamsızlık varsa pazar sıradan iş günü sayılır (normal ×1, FM ×1,5) ve "
            "tek eksik gün hafta tatiline döner. Eksik gün olup pazar da çalışılmadıysa pazar ücretsiz izin sayılır. "
            "Rapor, yıllık / babalık / vefat izni ve resmî tatil çalışılmış gün sayılır.").replace(
        "×1,5", f"×{_sa(K_HI)}").replace("×2,5", f"×{_sa(K_PZ)}")


def aylik_esik():
    """Aylık fazla mesai eşiği (saat). v5: Ayarlar sayfasından gelir (önceden Puantaj Suite ayarıydı)."""
    try:
        return float(ayarlar.al("aylik_fm_esigi"))
    except Exception:
        return 52.0


def kategori(durum, saat):
    """SAP 'Çalışma Durumu' serbest metnini + saati standart kategoriye çevirir.
    Saat girilmişse gün her zaman 'çalıştı' sayılır (hafta tatilinde çalışma dahil)."""
    if (saat or 0) > 0:
        return "calisti"
    return _durum_kategorisi(durum)


@functools.lru_cache(maxsize=512)
def _durum_kategorisi(durum):
    n = engine.normalize_name(durum)
    if not n:
        return "belirsiz"
    if "HAFTA TAT" in n or n == "HT":
        return "hafta_tatili"
    if "RESMI" in n or "BAYRAM" in n or "GENEL TATIL" in n:
        return "resmi_tatil"
    if "UCRETSIZ" in n:
        return "ucretsiz_izin"
    if "RAPOR" in n or "HASTA" in n or "KAZA" in n:
        return "rapor"
    if "DEVAMSIZ" in n or "GELMEDI" in n or "MAZERETSIZ" in n:
        return "devamsiz"
    if "IZIN" in n or "IZNI" in n:
        return "ucretli_izin"
    if "CALIS" in n:
        return "calisti"
    return "belirsiz"


def puan_bandi(puan):
    for alt, harf, ad in PUAN_BANTLARI:
        if puan >= alt:
            return harf, ad
    return "D", "Zayıf"


# ---------------------------------------------------------------- toplayıcı

class _T:
    """Bir kişi / ekip / görev / periyot / gün için ham toplamlar.
    katli_normal / katli_fazla: hafta tatili kuralına göre ×2,5 ödenen (hak edilmiş) pazar saatleri."""
    __slots__ = ("kayit", "gun", "normal", "fazla", "katli_normal", "katli_fazla", "pazar_gun", "uzun_gun",
                 "kisiler", "tarihler")

    def __init__(self):
        self.kayit = 0
        self.gun = dict.fromkeys(KAT_ETIKET, 0)
        self.normal = self.fazla = self.katli_normal = self.katli_fazla = 0.0
        self.pazar_gun = self.uzun_gun = 0
        self.kisiler = set()
        self.tarihler = set()

    def ekle(self, tc, d, kat, n, f, katli=False):
        self.kayit += 1
        self.gun[kat] += 1
        self.normal += n
        self.fazla += f
        if katli:
            self.katli_normal += n
            self.katli_fazla += f
        if d.weekday() == 6 and kat == "calisti":
            self.pazar_gun += 1
        if n + f > GUNLUK_AZAMI_SAAT:
            self.uzun_gun += 1
        self.kisiler.add(tc)
        self.tarihler.add(d)

    def birlestir(self, o):
        """Başka bir toplayıcıyı bu toplayıcıya ekler (satırları yeniden dolaşmadan üst toplam üretmek için)."""
        self.kayit += o.kayit
        for k, v in o.gun.items():
            self.gun[k] += v
        self.normal += o.normal
        self.fazla += o.fazla
        self.katli_normal += o.katli_normal
        self.katli_fazla += o.katli_fazla
        self.pazar_gun += o.pazar_gun
        self.uzun_gun += o.uzun_gun
        self.kisiler |= o.kisiler
        self.tarihler |= o.tarihler
        return self


def _birlestir(hedef, anahtar, t):
    h = hedef.get(anahtar)
    if h is None:
        h = hedef[anahtar] = _T()
    h.birlestir(t)


def _pct(a, b):
    return (a / b * 100.0) if b else 0.0


def metrikler(t):
    """Ham toplamlardan türetilmiş göstergeler (ekran + Excel aynı formüller)."""
    planlanan = t.kayit - sum(t.gun[k] for k in PLAN_DISI)
    calisilan = t.gun["calisti"]
    kayip = sum(t.gun[k] for k in KAYIP)
    toplam = t.normal + t.fazla
    yevmiye = hk.yevmiye(t.normal, t.fazla, t.katli_normal, t.katli_fazla)
    kapasite = _pct(t.normal, planlanan * STANDART_GUN)
    devam = _pct(calisilan, planlanan)
    kisi = len(t.kisiler)
    gun_sayisi = len(t.tarihler)
    puan = 0.6 * devam + 0.4 * min(kapasite, 100.0) if planlanan else 0.0
    return {
        "kisi": kisi, "kayit": t.kayit, "planlanan": planlanan, "calisilan": calisilan,
        "gun": dict(t.gun), "kayip": kayip,
        "devam": devam, "kayip_oran": _pct(kayip, planlanan), "devamsiz_oran": _pct(t.gun["devamsiz"], planlanan),
        "normal": t.normal, "fazla": t.fazla, "pazar_katli": t.katli_normal + t.katli_fazla, "toplam": toplam,
        "ort_saat": (toplam / calisilan) if calisilan else 0.0,
        "fm_oran": _pct(t.fazla, toplam),
        "kapasite": kapasite,
        "kayip_saat": max(0.0, planlanan * STANDART_GUN - t.normal),
        "yevmiye": yevmiye,
        # FM + pazar yevmiyesi: yevmiyenin, hafta içi normal saatlerin (×1) dışında kalan kısmı
        "fm_yevmiye": yevmiye - (t.normal - t.katli_normal) / STANDART_GUN,
        # katsayı farkı: aynı saatler ×1 ödenseydi ödenmeyecek kısım
        "fm_prim": yevmiye - toplam / STANDART_GUN,
        "kisi_basi_fm": (t.fazla / kisi) if kisi else 0.0,
        "ort_calisan": (calisilan / gun_sayisi) if gun_sayisi else 0.0,
        "pazar_gun": t.pazar_gun, "uzun_gun": t.uzun_gun, "gun_sayisi": gun_sayisi,
        "puan": puan,
    }


# ---------------------------------------------------------------- yardımcılar

def _periyot(d, periyot):
    if periyot == "gun":
        return d
    if periyot == "ay":
        return datetime.date(d.year, d.month, 1)
    return d - datetime.timedelta(days=d.weekday())


def _periyot_etiketi(key, periyot, d1, d2):
    if periyot == "gun":
        return f"{key:%d.%m}"
    if periyot == "ay":
        return f"{AY_KISA[key.month - 1]} {key.year}"
    lo, hi = max(key, d1), min(key + datetime.timedelta(days=6), d2)
    return f"{lo:%d}–{hi:%d.%m}" if lo.month == hi.month else f"{lo:%d.%m}–{hi:%d.%m}"


def _tarih(v):
    s = str(v or "").strip()[:10]
    try:
        return datetime.date.fromisoformat(s)
    except ValueError:
        return None


def _bos(v, etiket):
    v = (v or "").strip() if isinstance(v, str) else v
    return v or etiket


def _filtre_uyar(k, gruplar, ekipler, gorevler):
    return ((not gruplar or k["grup"] in gruplar) and (not ekipler or k["ekip"] in ekipler)
            and (not gorevler or k["gorev"] in gorevler))


def _satirlar(bas, bit, gruplar, ekipler, gorevler, db_path):
    """Analiz için günlük satırlar. sorgu.get_daily_hours_filtered ile AYNI süzme kuralı (boş alan etiketleri
    dahil), ama sıralama (ORDER BY ad) ve sözlüğe çevirme yok: analiz sıraya ihtiyaç duymaz; büyük dönemde
    sorgu süresi yarıya iner."""
    where, params = ["gp.tarih >= ?", "gp.tarih <= ?"], [bas, bit]

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
        _in("p.grup", gruplar, NO_GRUP)
    if ekipler:
        _in("p.bagli_oldugu_ekip", ekipler, NO_EKIP)
    if gorevler:
        _in("p.gorevi", gorevler, NO_GOREV)
    conn = sema.get_connection(db_path)
    try:
        return conn.execute(
            "SELECT gp.tc, COALESCE(p.ad_soyad, '(personel kaydı yok)') AS ad_soyad, p.grup, p.bagli_oldugu_ekip, "
            "p.gorevi, gp.tarih, gp.durum, gp.normal_saat, gp.fazla_saat "
            f"FROM gunluk_puantaj gp LEFT JOIN personnel p ON p.tc = gp.tc WHERE {' AND '.join(where)}",
            params).fetchall()
    finally:
        conn.close()


def _kurali_satirlar(d1, d2, gruplar, ekipler, gorevler, db_path):
    """Dönemin satırları + hafta tatili / pazar kuralı. Kural, dönem sınırındaki haftaların TAMAMINA bakılarak
    uygulanır (satırlar tam haftalar için okunur), ama dönen satırlar yalnız [d1, d2] içindedir.
    Döner: (satırlar, {(tc, tarih_iso): (kategori, pazar_katli)})"""
    gb, gs = hk.genislet(d1.isoformat(), d2.isoformat())
    tum = _satirlar(gb, gs, gruplar, ekipler, gorevler, db_path)
    kural = hk.kural_uygula((r["tc"], r["tarih"], r["durum"], r["normal_saat"], r["fazla_saat"]) for r in tum)
    bas, bit = d1.isoformat(), d2.isoformat()
    return [r for r in tum if bas <= r["tarih"] <= bit], kural


def _yillik_fm(tcs, yil, bit, db_path):
    out = {}
    if not tcs:
        return out
    conn = sema.get_connection(db_path)
    try:
        tcs = list(tcs)
        for i in range(0, len(tcs), 900):
            parca = tcs[i:i + 900]
            for r in conn.execute(
                    f"SELECT tc, COALESCE(SUM(fazla_saat),0) AS f FROM gunluk_puantaj WHERE tarih BETWEEN ? AND ? "
                    f"AND tc IN ({','.join('?' * len(parca))}) GROUP BY tc", [f"{yil}-01-01", bit] + parca):
                out[r["tc"]] = r["f"]
    finally:
        conn.close()
    return out


def _hareketler(d1, d2, gruplar, ekipler, gorevler, db_path):
    """Dönemde işe giren / işten çıkan personel ve dönem başı / sonu mevcut."""
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute("SELECT tc, ad_soyad, grup, bagli_oldugu_ekip, gorevi, ise_giris_tarihi, "
                            "isten_cikis_tarihi FROM personnel").fetchall()
    finally:
        conn.close()
    g, e, o = set(gruplar or []), set(ekipler or []), set(gorevler or [])
    girenler, cikanlar = [], []
    bas_mevcut = son_mevcut = 0
    for r in rows:
        k = {"tc": r["tc"], "ad_soyad": (r["ad_soyad"] or "").strip(), "grup": _bos(r["grup"], NO_GRUP),
             "ekip": _bos(r["bagli_oldugu_ekip"], NO_EKIP), "gorev": _bos(r["gorevi"], NO_GOREV),
             "giris": _tarih(r["ise_giris_tarihi"]), "cikis": _tarih(r["isten_cikis_tarihi"])}
        if not _filtre_uyar(k, g, e, o):
            continue
        if k["giris"] and d1 <= k["giris"] <= d2:
            girenler.append(k)
        if k["cikis"] and d1 <= k["cikis"] <= d2:
            cikanlar.append(k)
        if (not k["giris"] or k["giris"] <= d1) and (not k["cikis"] or k["cikis"] >= d1):
            bas_mevcut += 1
        if (not k["giris"] or k["giris"] <= d2) and (not k["cikis"] or k["cikis"] >= d2):
            son_mevcut += 1
    girenler.sort(key=lambda k: (k["giris"], k["ad_soyad"]))
    cikanlar.sort(key=lambda k: (k["cikis"], k["ad_soyad"]))
    ort = (bas_mevcut + son_mevcut) / 2
    return {"girenler": girenler, "cikanlar": cikanlar, "bas_mevcut": bas_mevcut, "son_mevcut": son_mevcut,
            "devir_orani": _pct(len(cikanlar), ort)}


def _seriler(gunler):
    """gunler: tarihe göre sıralı [(tarih, kategori)].
    Döner: (en uzun kesintisiz çalışma serisi, [(başlangıç, bitiş, gün)] - eşiği aşan seriler).
    Kaydı olmayan gün seriyi böler (veri yoksa çalışıldığı varsayılmaz)."""
    seriler, en_uzun, bas, onceki, uz = [], 0, None, None, 0
    for d, kat, *_ in gunler:
        if kat == "calisti":
            if onceki is not None and (d - onceki).days == 1:
                uz += 1
            else:
                if uz >= KESINTISIZ_GUN_ESIGI:
                    seriler.append((bas, onceki, uz))
                bas, uz = d, 1
            onceki = d
            en_uzun = max(en_uzun, uz)
        else:
            if uz >= KESINTISIZ_GUN_ESIGI:
                seriler.append((bas, onceki, uz))
            bas, onceki, uz = None, None, 0
    if uz >= KESINTISIZ_GUN_ESIGI:
        seriler.append((bas, onceki, uz))
    return en_uzun, seriler


# ---------------------------------------------------------------- ana analiz

def _topla(rows, periyot, d1, d2, kirilim, tam=True, kural=None):
    """Satırlar yalnız İKİ temel toplayıcıya eklenir: kişi ve (alt ekip, gün). Genel, kırılım, görev, periyot,
    gün, haftanın günü ve ısı haritası toplamları bunların birleştirilmesiyle üretilir (satır başına 9 ekleme
    yerine 2 - büyük veride analiz ~3 kat hızlanır).
    tam=False: yalnız genel / kırılım / görev toplamları (önceki dönem karşılaştırması için)."""
    kisi_t, hucre_t = {}, {}                  # tc -> _T ; (isi_ad, tarih) -> _T
    kisi_bilgi, kisi_gunler, uzun_gunler = {}, {}, []
    tarih_onbellek = {}
    for r in rows:
        ts = r["tarih"]
        d = tarih_onbellek.get(ts)
        if d is None:
            d = tarih_onbellek[ts] = datetime.date.fromisoformat(ts)
        n, f = float(r["normal_saat"] or 0.0), float(r["fazla_saat"] or 0.0)
        tc = r["tc"]
        # Hafta tatili / pazar kuralı (veri_merkezi_hafta_kurali): kategori düzeltilmiş olabilir
        # (ör. pazar çalışılan haftada cuma ücretsiz izni -> hafta tatili) ve pazar ×2,5 bilgisi.
        kat, katli = kural.get((tc, ts)) if kural and (tc, ts) in kural else (kategori(r["durum"], n + f), False)
        b = kisi_bilgi.get(tc)
        if b is None:
            grup, ekip = _bos(r["grup"], NO_GRUP), _bos(r["bagli_oldugu_ekip"], NO_EKIP)
            b = kisi_bilgi[tc] = {"tc": tc, "ad_soyad": r["ad_soyad"], "grup": grup, "ekip": ekip,
                                  "gorev": _bos(r["gorevi"], NO_GOREV),
                                  "kirilim": ekip if kirilim == "ekip" else grup,
                                  "isi_ad": ekip if kirilim == "ekip" else f"{grup} › {ekip}"}
            kisi_t[tc] = _T()
        kisi_t[tc].ekle(tc, d, kat, n, f, katli)
        if not tam:
            continue
        anahtar = (b["isi_ad"], d)
        h = hucre_t.get(anahtar)
        if h is None:
            h = hucre_t[anahtar] = _T()
        h.ekle(tc, d, kat, n, f, katli)
        kisi_gunler.setdefault(tc, []).append((d, kat, f))
        if n + f > GUNLUK_AZAMI_SAAT:
            uzun_gunler.append({"tc": tc, "ad_soyad": r["ad_soyad"], "ekip": b["ekip"], "gorev": b["gorev"],
                                "tarih": d, "normal": n, "fazla": f, "toplam": n + f})

    genel, kir_t, gor_t = _T(), {}, {}
    for tc, t in kisi_t.items():
        b = kisi_bilgi[tc]
        genel.birlestir(t)
        _birlestir(kir_t, b["kirilim"], t)
        _birlestir(gor_t, b["gorev"], t)
    per_t, gun_t, hg_t, isi_t = {}, {}, {}, {}
    if tam:
        for (isi_ad, d), t in hucre_t.items():
            isi_t[(isi_ad, d)] = t
            _birlestir(isi_t, (isi_ad, None), t)
            _birlestir(gun_t, d, t)
        for d, t in gun_t.items():
            _birlestir(per_t, _periyot(d, periyot), t)
            _birlestir(hg_t, d.weekday(), t)
    for b in kisi_bilgi.values():
        b.pop("isi_ad", None)
    return genel, kisi_t, kir_t, gor_t, per_t, gun_t, hg_t, isi_t, kisi_bilgi, kisi_gunler, uzun_gunler


# ---- sonuç önbelleği: aynı filtrelerle sayfaya dönünce ya da Excel alırken analiz tekrarlanmaz.
# Anahtar veritabanı dosyasının değişiklik zamanını da içerir: içe aktarma / personel düzenleme olunca
# önbellek kendiliğinden geçersiz olur.
_ONBELLEK = {}
_ONBELLEK_BOYUT = 6


def _veri_surumu(db_path=None):
    try:
        yol = Path(db_path) if db_path else Path(sema.DB_PATH)
        st = yol.stat()
        wal = yol.with_name(yol.name + "-wal")
        return (st.st_mtime_ns, st.st_size, wal.stat().st_mtime_ns if wal.exists() else 0, datetime.date.today())
    except OSError:
        return None


def analiz(bas, bit, periyot="hafta", gruplar=None, ekipler=None, gorevler=None, db_path=None):
    """bas/bit: ISO tarih. periyot: 'gun' | 'hafta' | 'ay'. Döner: ekran + Excel için tek sözlük.
    Sonuç önbelleğe alınır (bkz. _ONBELLEK); dönen sözlük değiştirilmemelidir."""
    surum = _veri_surumu(db_path)
    anahtar = (bas[:10], bit[:10], periyot, tuple(gruplar or ()), tuple(ekipler or ()), tuple(gorevler or ()),
               str(db_path), surum)
    if surum is not None and anahtar in _ONBELLEK:
        d = dict(_ONBELLEK[anahtar])
        d["olusturma"] = datetime.datetime.now()
        return d
    d = _analiz(bas, bit, periyot, gruplar, ekipler, gorevler, db_path)
    if surum is not None:
        if len(_ONBELLEK) >= _ONBELLEK_BOYUT:
            _ONBELLEK.pop(next(iter(_ONBELLEK)))
        _ONBELLEK[anahtar] = d
    return d


def _analiz(bas, bit, periyot="hafta", gruplar=None, ekipler=None, gorevler=None, db_path=None):
    d1, d2 = datetime.date.fromisoformat(bas[:10]), datetime.date.fromisoformat(bit[:10])
    if d2 < d1:
        d1, d2 = d2, d1
    gun_sayisi = (d2 - d1).days + 1
    kirilim = "ekip" if (gruplar and len(gruplar) == 1) else "grup"
    rows, kural = _kurali_satirlar(d1, d2, gruplar, ekipler, gorevler, db_path)
    (genel, kisi_t, kir_t, gor_t, per_t, gun_t, hg_t, isi_t, kisi_bilgi, kisi_gunler,
     uzun_gunler) = _topla(rows, periyot, d1, d2, kirilim, kural=kural)

    # ---- önceki eşit uzunluktaki dönem
    o_bit = d1 - datetime.timedelta(days=1)
    o_bas = o_bit - datetime.timedelta(days=gun_sayisi - 1)
    o_rows, o_kural = _kurali_satirlar(o_bas, o_bit, gruplar, ekipler, gorevler, db_path)
    o_genel, _ok, o_kir, o_gor, *_ = _topla(o_rows, periyot, o_bas, o_bit, kirilim, tam=False, kural=o_kural)
    onceki = metrikler(o_genel) if o_rows else None
    o_kir_m = {k: metrikler(t) for k, t in o_kir.items()}
    o_gor_m = {k: metrikler(t) for k, t in o_gor.items()}

    g = metrikler(genel)
    esik_aylik = aylik_esik()
    esik_donem = esik_aylik * gun_sayisi / 30.0
    yillik = _yillik_fm(kisi_t.keys(), d2.year, d2.isoformat(), db_path)

    # ---- personel karnesi
    kisiler = []
    for tc, t in kisi_t.items():
        m = metrikler(t)
        m.update(kisi_bilgi[tc])
        gunler = sorted(kisi_gunler[tc])
        m["en_uzun_seri"], m["seriler"] = _seriler(gunler)
        m["yillik_fm"] = yillik.get(tc, m["fazla"])
        m["yillik_kalan"] = max(0.0, YILLIK_FM_SINIRI - m["yillik_fm"])
        m["esik_ustu"] = m["fazla"] > esik_donem
        m["harf"], m["bant"] = puan_bandi(m["puan"])
        m["periyot_fm"] = {}
        riskler = []
        if m["yillik_fm"] >= YILLIK_FM_SINIRI:
            riskler.append(f"Yıllık {etiket_yillik()} sa aşıldı")
        elif m["yillik_fm"] >= YILLIK_FM_SINIRI * 0.85:
            riskler.append(f"Yıllık {etiket_yillik()} sa sınırına yakın")
        if m["esik_ustu"]:
            riskler.append("FM eşiği üstü")
        if m["uzun_gun"]:
            riskler.append(f"{m['uzun_gun']} gün {etiket_gunluk()} sa üstü")
        if m["seriler"]:
            riskler.append(f"{m['en_uzun_seri']} gün kesintisiz")
        if m["gun"]["devamsiz"] >= DEVAMSIZLIK_ESIGI:
            riskler.append(f"{m['gun']['devamsiz']} gün devamsızlık")
        m["riskler"] = riskler
        kisiler.append(m)
    # periyot bazlı FM (kişi)
    per_anahtarlar = sorted(per_t)
    for k in kisiler:
        pf = k["periyot_fm"]
        for d, _kat, f in kisi_gunler[k["tc"]]:
            if f:
                pk = _periyot(d, periyot)
                pf[pk] = pf.get(pk, 0.0) + f
    for k in kisiler:
        k["periyot_fm_liste"] = [k["periyot_fm"].get(p, 0.0) for p in per_anahtarlar]
        del k["periyot_fm"]
    kisiler.sort(key=lambda k: (-k["fazla"], engine.normalize_name(k["ad_soyad"])))

    # ---- periyotlar
    periyotlar = []
    for key in per_anahtarlar:
        m = metrikler(per_t[key])
        m.update(anahtar=key, etiket=_periyot_etiketi(key, periyot, d1, d2))
        periyotlar.append(m)

    # ---- günlük seri (ısı haritası ve günlük tablo için)
    gunluk = []
    for d in sorted(gun_t):
        m = metrikler(gun_t[d])
        m.update(tarih=d, gun_adi=GUN_KISA[d.weekday()])
        gunluk.append(m)

    # ---- haftanın günleri
    haftanin_gunleri = []
    for wd in range(7):
        if wd in hg_t:
            m = metrikler(hg_t[wd])
            m.update(wd=wd, ad=GUN_ADLARI[wd], kisa=GUN_KISA[wd])
            haftanin_gunleri.append(m)

    # ---- kırılım + görev tabloları
    def _tablo(src, onceki_m):
        out = []
        for ad, t in src.items():
            m = metrikler(t)
            m["ad"] = ad
            om = onceki_m.get(ad)
            m["onceki_fazla"] = om["fazla"] if om else None
            m["onceki_devam"] = om["devam"] if om else None
            m["fm_degisim"] = _pct(m["fazla"] - om["fazla"], om["fazla"]) if om and om["fazla"] else None
            m["devam_fark"] = (m["devam"] - om["devam"]) if om and om["planlanan"] else None
            m["harf"], m["bant"] = puan_bandi(m["puan"])
            out.append(m)
        out.sort(key=lambda x: -x["kisi"])
        for i, m in enumerate(sorted(out, key=lambda x: -x["puan"]), start=1):
            m["sira"] = i
        return out
    kirilimlar = _tablo(kir_t, o_kir_m)
    gorevler_t = _tablo(gor_t, o_gor_m)

    # ---- ısı haritası (kırılım × gün, devam oranı)
    tarihler = sorted(gun_t)
    isi = []
    satir_adlari = sorted({a for a, t in isi_t if t is None},
                          key=lambda a: (a.split(" › ")[0], -len(isi_t[(a, None)].kisiler), a))
    for ad in satir_adlari:
        top = metrikler(isi_t[(ad, None)])
        hucreler = []
        for d in tarihler:
            t = isi_t.get((ad, d))
            if t is None:
                hucreler.append(None)
            else:
                m = metrikler(t)
                hucreler.append({"devam": m["devam"] if m["planlanan"] else None, "calisan": m["calisilan"],
                                 "planlanan": m["planlanan"], "saat": m["toplam"]})
        isi.append({"ad": ad, "hucreler": hucreler, "devam": top["devam"], "kisi": top["kisi"]})

    # ---- uyum ve risk listeleri
    risk = {
        "yillik_asan": sorted([k for k in kisiler if k["yillik_fm"] >= YILLIK_FM_SINIRI], key=lambda k: -k["yillik_fm"]),
        "yillik_yakin": sorted([k for k in kisiler if YILLIK_FM_SINIRI * 0.85 <= k["yillik_fm"] < YILLIK_FM_SINIRI],
                               key=lambda k: -k["yillik_fm"]),
        "esik_ustu": sorted([k for k in kisiler if k["esik_ustu"]], key=lambda k: -k["fazla"]),
        "uzun_gunler": sorted(uzun_gunler, key=lambda x: (-x["toplam"], x["tarih"])),
        "kesintisiz": sorted([k for k in kisiler if k["seriler"]], key=lambda k: -k["en_uzun_seri"]),
        "devamsiz": sorted([k for k in kisiler if k["gun"]["devamsiz"] >= DEVAMSIZLIK_ESIGI],
                           key=lambda k: -k["gun"]["devamsiz"]),
    }
    uzun_kisi = {}
    for u in uzun_gunler:
        x = uzun_kisi.setdefault(u["tc"], {"tc": u["tc"], "ad_soyad": u["ad_soyad"], "ekip": u["ekip"],
                                           "gorev": u["gorev"], "gun": 0, "en_yuksek": 0.0, "tarihler": []})
        x["gun"] += 1
        x["en_yuksek"] = max(x["en_yuksek"], u["toplam"])
        x["tarihler"].append(u["tarih"])
    for x in uzun_kisi.values():
        x["tarihler"].sort()
    risk["uzun_kisi"] = sorted(uzun_kisi.values(), key=lambda x: (-x["gun"], -x["en_yuksek"]))
    risk["riskli_kisi"] = sum(1 for k in kisiler if k["riskler"])

    hareket = _hareketler(d1, d2, gruplar, ekipler, gorevler, db_path)
    try:
        kapsama = sorgu.gun_kapsamasi(d1.isoformat(), d2.isoformat(), db_path=db_path)
    except Exception:
        kapsama = []
    eksik_gun = [c["tarih"] for c in kapsama if c["durum"] == "eksik"]
    kismi_gun = [c["tarih"] for c in kapsama if c["durum"] == "kismi"]

    data = {
        "bas": d1, "bit": d2, "gun_sayisi": gun_sayisi, "periyot": periyot, "kirilim": kirilim,
        "gruplar": gruplar or [], "ekipler": ekipler or [], "gorevler": gorevler or [],
        "genel": g, "onceki": onceki, "onceki_aralik": (o_bas, o_bit),
        "periyotlar": periyotlar, "gunluk": gunluk, "haftanin_gunleri": haftanin_gunleri,
        "kirilimlar": kirilimlar, "gorev_tablosu": gorevler_t, "kisiler": kisiler, "isi": isi, "isi_tarihler": tarihler,
        "risk": risk, "hareket": hareket, "eksik_gun": eksik_gun, "kismi_gun": kismi_gun,
        "esik_aylik": esik_aylik, "esik_donem": esik_donem, "yil": d2.year, "kayit_sayisi": len(rows),
        "olusturma": datetime.datetime.now(),
        # Devam Takvimi (Excel) için: gün bazında kurala göre düzeltilmiş kategori ve pazar ×2,5 bilgisi
        "gun_kurali": {k: v for k, v in kural.items() if d1.isoformat() <= k[1] <= d2.isoformat()},
    }
    data["ozet"] = yonetici_ozeti(data)
    return data


# ---------------------------------------------------------------- yönetici özeti

def _s(x, ondalik=0):
    """Türkçe sayı: 1.234,5"""
    return report.format_tr(x, ondalik)


def _fark_metni(simdi, once, birim="%", puan=False):
    if once is None:
        return ""
    fark = simdi - once
    if abs(fark) < 0.05:
        return " (önceki dönemle aynı)"
    yon = "arttı" if fark > 0 else "azaldı"
    if puan:
        return f" (önceki döneme göre {_s(abs(fark), 1)} puan {'yükseldi' if fark > 0 else 'düştü'})"
    if once:
        return f" (önceki döneme göre %{_s(abs(fark) / once * 100, 1)} {yon})"
    return ""


def yonetici_ozeti(d):
    """Veriden otomatik 'öne çıkanlar'. Döner: [{"ton": iyi|uyari|kritik|bilgi, "baslik", "metin"}]"""
    g, o = d["genel"], d["onceki"]
    out = []
    if not g["kayit"]:
        return out

    out.append({"ton": "bilgi", "baslik": "İş gücü",
                "metin": f"Dönemde günde ortalama {_s(g['ort_calisan'], 0)} kişi çalıştı; toplam "
                         f"{_s(g['calisilan'])} adam-gün ve {_s(g['toplam'])} saat çalışma yapıldı "
                         f"({_s(g['kisi'])} farklı personel)."})

    ton = "iyi" if g["devam"] >= 95 else ("uyari" if g["devam"] >= 88 else "kritik")
    out.append({"ton": ton, "baslik": "Devam",
                "metin": f"Devam oranı %{_s(g['devam'], 1)}"
                         f"{_fark_metni(g['devam'], o['devam'] if o else None, puan=True)}. "
                         f"Planlanan {_s(g['planlanan'])} iş gününün {_s(g['kayip'])}'i kayıp: "
                         f"{_s(g['gun']['devamsiz'])} devamsızlık, {_s(g['gun']['rapor'])} rapor, "
                         f"{_s(g['gun']['ucretsiz_izin'])} ücretsiz izin."})

    en_fm = max(d["kirilimlar"], key=lambda k: k["fm_oran"], default=None)
    ton = "iyi" if g["fm_oran"] < 10 else ("uyari" if g["fm_oran"] < 18 else "kritik")
    metin = (f"Fazla mesai {_s(g['fazla'])} saat, toplam çalışmanın %{_s(g['fm_oran'], 1)}'i"
             f"{_fark_metni(g['fazla'], o['fazla'] if o else None)}. ")
    if g["fazla"]:
        metin += f"Hak edilmiş pazarlarda {_s(g['pazar_katli'])} saat (normal + FM) ×2,5 ödendi. "
    if en_fm and len(d["kirilimlar"]) > 1 and en_fm["fazla"]:
        metin += f"En yüksek FM oranı: {en_fm['ad']} (%{_s(en_fm['fm_oran'], 1)})."
    out.append({"ton": ton, "baslik": "Fazla mesai", "metin": metin.strip()})

    if g["fazla"]:
        out.append({"ton": "bilgi", "baslik": "Maliyet etkisi",
                    "metin": f"Fazla mesai {_s(g['fm_yevmiye'], 1)} yevmiyeye karşılık geliyor; bunun "
                             f"{_s(g['fm_prim'], 1)} yevmiyesi katsayı farkından (aynı saatler normal mesaide "
                             f"çalışılsaydı ödenmeyecek ek tutar). Toplam yevmiye karşılığı {_s(g['yevmiye'], 1)} gün."})

    if len(d["kirilimlar"]) > 1:
        sirali = sorted([k for k in d["kirilimlar"] if k["planlanan"] >= 5], key=lambda k: -k["puan"])
        if sirali:
            iyi, zayif = sirali[0], sirali[-1]
            etiket = "Alt ekip" if d["kirilim"] == "ekip" else "Grup"
            out.append({"ton": "bilgi", "baslik": f"{etiket} karşılaştırması",
                        "metin": f"En yüksek verimlilik puanı {iyi['ad']} ({_s(iyi['puan'], 1)}, devam "
                                 f"%{_s(iyi['devam'], 1)}); en düşük {zayif['ad']} ({_s(zayif['puan'], 1)}, devam "
                                 f"%{_s(zayif['devam'], 1)})."})

    r = d["risk"]
    parca = []
    if r["yillik_asan"]:
        parca.append(f"{len(r['yillik_asan'])} kişi yıllık {etiket_yillik()} saat FM sınırını aştı")
    if r["yillik_yakin"]:
        parca.append(f"{len(r['yillik_yakin'])} kişi sınıra yakın (≥ %85)")
    if r["uzun_gunler"]:
        parca.append(f"{len(r['uzun_gunler'])} kez günlük {etiket_gunluk()} saat aşıldı")
    if r["kesintisiz"]:
        parca.append(f"{len(r['kesintisiz'])} kişi {KESINTISIZ_GUN_ESIGI}+ gün hafta tatili kullanmadan çalıştı")
    if r["devamsiz"]:
        parca.append(f"{len(r['devamsiz'])} kişinin {DEVAMSIZLIK_ESIGI}+ gün devamsızlığı var")
    if parca:
        ton = "kritik" if (r["yillik_asan"] or r["uzun_gunler"]) else "uyari"
        out.append({"ton": ton, "baslik": "Uyum ve risk", "metin": "; ".join(parca) + "."})
    else:
        out.append({"ton": "iyi", "baslik": "Uyum ve risk",
                    "metin": f"Yasal sınırlar (yıllık {etiket_yillik()} sa, günlük {etiket_gunluk()} sa, haftalık dinlenme) "
                            f"açısından risk görünmüyor."})

    h = d["hareket"]
    if h["girenler"] or h["cikanlar"]:
        out.append({"ton": "bilgi", "baslik": "Personel hareketi",
                    "metin": f"{len(h['girenler'])} kişi işe başladı, {len(h['cikanlar'])} kişi ayrıldı "
                             f"(mevcut {_s(h['bas_mevcut'])} → {_s(h['son_mevcut'])}; devir oranı "
                             f"%{_s(h['devir_orani'], 1)})."})

    if d["eksik_gun"]:
        out.append({"ton": "uyari", "baslik": "Veri kapsaması",
                    "metin": f"{len(d['eksik_gun'])} günün SAP mesai verisi eksik; bu günler analize girmedi. "
                             f"Sonuçlar bu günler yüklendikçe değişir."})
    return out


# ---------------------------------------------------------------- adlandırma

def kapsam_adi(data):
    if data["ekipler"]:
        liste, birim = data["ekipler"], "ALT EKİP"
    elif data["gruplar"]:
        liste, birim = data["gruplar"], "GRUP"
    else:
        return "TÜM GRUPLAR"
    ad = " / ".join(liste) if len(liste) <= 3 else f"{len(liste)} {birim}"
    if data["gorevler"]:
        ad += f" · {len(data['gorevler'])} görev"
    return ad


def donem_adi(data):
    d1, d2 = data["bas"], data["bit"]
    ay_son = (datetime.date(d1.year + (d1.month == 12), d1.month % 12 + 1, 1) - datetime.timedelta(days=1))
    if d1.day == 1 and d2 == ay_son:
        return f"{AY_ADLARI[d1.month - 1]} {d1.year}"
    return f"{d1:%d.%m.%Y} – {d2:%d.%m.%Y}"
