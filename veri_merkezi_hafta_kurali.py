# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - HAFTA TATİLİ VE PAZAR ÇARPANI KURALI (veri_merkezi_hafta_kurali.py)
=====================================================================================
Yevmiye (ücret günü) hesabının TEK kaynağı. Genel Bakış, Puantaj Geçmişi, Personel
Kartı, Verimlilik Analizi ekranı ve Excel raporu hepsi bu modülü kullanır.

ÇARPANLAR
  Pazartesi - Cumartesi : normal saat × 1, fazla mesai × 1,5
  Pazar (hak edilmişse) : (normal + fazla mesai) × 2,5
  Pazar (hak edilmemişse): o gün sıradan bir iş günü gibi: normal × 1, fazla × 1,5

HAFTA (Pazartesi-Pazar) BAZINDA KURAL - her kişi için ayrı:
  E = Pazartesi-Cumartesi arasındaki MAZERETSİZ günler: ücretsiz izin, devamsızlık.
  Mazeretli günler (sağlık raporu, yıllık / babalık / vefat vb. ücretli izinler, resmî
  tatil) çalışılmış sayılır; kurala etki etmez.

  1) E yok  -> 6 gün tamamlanmış. Pazar çalışıldıysa pazar saatleri × 2,5.
  2) E var ve PAZAR ÇALIŞILDI -> pazar ×2,5 DEĞİL, sıradan iş günü sayılır.
       E tek günse (ör. cuma ücretsiz izin) o gün HAFTA TATİLİNE döner: kişi pazarla
       birlikte 6 gün çalışmış, hafta tatilini hak etmiştir.
       E iki gün ve üzeriyse hiçbiri dönmez (pazarla birlikte bile 6 gün tamamlanmamış).
  3) E var ve PAZAR ÇALIŞILMADI -> kişi hafta tatilini hak etmemiştir: pazar günü
       (hafta tatili / boş kaydı) ÜCRETSİZ İZİN sayılır. E günleri aynen kalır.

  Ek: Pazartesi-Cumartesi arasında bir gün HAFTA TATİLİ kodluysa kişinin hafta tatili o
  güne kaydırılmıştır; pazar çalışması sıradan iş günü sayılır (×2,5 değil).

  Veri olmayan günler (kayıt yok: işe giriş öncesi, çıkış sonrası, eksik yükleme) kuralı
  TETİKLEMEZ; yalnızca açıkça ücretsiz izin / devamsızlık yazan günler tetikler. Böylece
  eksik bir içe aktarma kimsenin pazarını haksız yere düşürmez.

DÖNEM SINIRI: bir dönem haftanın ortasında başlasa/bitse de kural HER ZAMAN o haftanın
tamamına bakılarak uygulanır (ör. 1 Ekim çarşamba ise pazartesi-salı da okunur); ama
toplamlara yalnızca dönem içindeki günler girer.
"""

import datetime

import veri_merkezi_sema as sema

MAZERETSIZ = ("ucretsiz_izin", "devamsiz")


def hafta_basi(d):
    return d - datetime.timedelta(days=d.weekday())


def genislet(bas, bit):
    """Dönemi tam haftalara genişletir. bas/bit: ISO metin ya da None. Döner: ISO metinler (None korunur)."""
    b = hafta_basi(datetime.date.fromisoformat(str(bas)[:10])).isoformat() if bas else None
    s = None
    if bit:
        d = datetime.date.fromisoformat(str(bit)[:10])
        s = (d + datetime.timedelta(days=6 - d.weekday())).isoformat()
    return b, s


def _kategori(durum, saat):
    import veri_merkezi_verimlilik as vm   # döngüsel içe aktarmayı önlemek için geç yüklenir
    return vm.kategori(durum, saat)


def kural_uygula(kayitlar):
    """kayitlar: (tc, tarih_iso, durum, normal, fazla) demetleri - en az kuralın uygulanacağı haftaların
    tamamını içermeli. Döner: {(tc, tarih_iso): (kategori, pazar_katli)} - HER kayıt için.
    kategori: kurala göre DÜZELTİLMİŞ kategori (ör. cuma ücretsiz izin -> hafta_tatili).
    pazar_katli: bu günün saatleri (normal + fazla) × pazar katsayısıyla mı ödenir."""
    sonuc = {}
    haftalar = {}
    for tc, ts, durum, n, f in kayitlar:
        d = datetime.date.fromisoformat(str(ts)[:10])
        kat = _kategori(durum, float(n or 0) + float(f or 0))
        sonuc[(tc, ts)] = (kat, False)
        haftalar.setdefault((tc, hafta_basi(d)), {})[d.weekday()] = (ts, kat)

    for (tc, _pzt), gunler in haftalar.items():
        eksik = [gunler[w][0] for w in range(6) if w in gunler and gunler[w][1] in MAZERETSIZ]
        kaydirilmis_ht = any(w in gunler and gunler[w][1] == "hafta_tatili" for w in range(6))
        pazar = gunler.get(6)
        if pazar and pazar[1] == "calisti":
            if not eksik and not kaydirilmis_ht:
                sonuc[(tc, pazar[0])] = ("calisti", True)
            elif len(eksik) == 1 and not kaydirilmis_ht:
                sonuc[(tc, eksik[0])] = ("hafta_tatili", False)
        elif eksik and pazar and pazar[1] in ("hafta_tatili", "belirsiz"):
            sonuc[(tc, pazar[0])] = ("ucretsiz_izin", False)
    return sonuc


def yevmiye(normal, fazla, katli_normal=0.0, katli_fazla=0.0):
    """Ücret günü karşılığı. normal/fazla: TÜM saatler; katli_*: bunların ×2,5 ödenen pazar kısmı."""
    import veri_merkezi_sorgu as sorgu
    k_hi, k_pz, gun = sorgu.FAZLA_MESAI_KATSAYI_HAFTAICI, sorgu.FAZLA_MESAI_KATSAYI_PAZAR, sorgu.STANDART_GUN_SAATI
    return ((normal - katli_normal) + (fazla - katli_fazla) * k_hi + (katli_normal + katli_fazla) * k_pz) / gun


def _bos_toplam():
    return {"normal": 0.0, "fazla": 0.0, "katli_normal": 0.0, "katli_fazla": 0.0, "katli_gun": 0, "kayit": 0}


def topla(kayitlar, bas=None, bit=None, kural=None):
    """kayitlar: (tc, tarih_iso, durum, normal, fazla). Kural tam haftalar üzerinden uygulanır, toplama yalnız
    [bas, bit] içindeki günler girer. Döner: normal, fazla, katli_normal, katli_fazla, katli_gun, kayit,
    pazar_katli (saat), yevmiye."""
    kayitlar = list(kayitlar)
    kural = kural if kural is not None else kural_uygula(kayitlar)
    t = _bos_toplam()
    for tc, ts, _durum, n, f in kayitlar:
        if (bas and ts < bas) or (bit and ts > bit):
            continue
        n, f = float(n or 0), float(f or 0)
        t["normal"] += n
        t["fazla"] += f
        t["kayit"] += 1
        if kural.get((tc, ts), (None, False))[1]:
            t["katli_normal"] += n
            t["katli_fazla"] += f
            t["katli_gun"] += 1
    t["pazar_katli"] = t["katli_normal"] + t["katli_fazla"]
    t["yevmiye"] = yevmiye(t["normal"], t["fazla"], t["katli_normal"], t["katli_fazla"])
    return t


def db_kayitlari(bas=None, bit=None, tc=None, tcs=None, db_path=None):
    """Kural için gereken kayıtlar: dönem tam haftalara genişletilerek okunur."""
    gb, gs = genislet(bas, bit)
    where, params = ["1=1"], []
    if gb:
        where.append("tarih >= ?")
        params.append(gb)
    if gs:
        where.append("tarih <= ?")
        params.append(gs)
    if tc:
        where.append("tc = ?")
        params.append(tc)
    conn = sema.get_connection(db_path)
    try:
        if tcs is not None:
            tcs = list(tcs)
            out = []
            for i in range(0, len(tcs), 900):
                parca = tcs[i:i + 900]
                out += [tuple(r) for r in conn.execute(
                    f"SELECT tc, tarih, durum, normal_saat, fazla_saat FROM gunluk_puantaj WHERE {' AND '.join(where)} "
                    f"AND tc IN ({','.join('?' * len(parca))})", params + parca)]
            return out
        return [tuple(r) for r in conn.execute(
            f"SELECT tc, tarih, durum, normal_saat, fazla_saat FROM gunluk_puantaj WHERE {' AND '.join(where)}", params)]
    finally:
        conn.close()


def db_toplami(bas=None, bit=None, tc=None, tcs=None, db_path=None):
    """Veritabanından okuyup kurala göre toplar (bkz. topla)."""
    if tcs is not None and not tcs:
        t = _bos_toplam()
        t["pazar_katli"], t["yevmiye"] = 0.0, 0.0
        return t
    return topla(db_kayitlari(bas, bit, tc, tcs, db_path), bas, bit)
