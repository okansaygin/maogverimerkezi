# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - GÜNLÜK MESAİ FORMU (PDF) (veri_merkezi_mesai_formu.py)
========================================================================
Ekip Listesi'nde seçilen kapsam için, sahada elle doldurulacak A4 YATAY
günlük mesai formunu PDF olarak üretir.

Kurallar (tasarımda onaylanan):
  - Her ALT EKİP kendi formunu alır (kendi başlığı ve sayfaları).
  - Başlık: GRUP / ALT EKİP / PERSONEL / TARİH. Alt kısımda iki imza kutusu:
    Hazırlayan ve Kontrol eden (v5.1: formen alanı ve "Puantaja işleyen" kutusu kaldırıldı).
    Birden fazla alt ekip seçildiyse hepsi TEK PDF'te, arka arkaya çıkar.
  - Sayfa başına en fazla 20 kişi; fazlası sonraki sayfaya geçer, sıra
    numarası sayfalar boyunca devam eder.
  - Satır yüksekliği ve yazı boyutu kişi sayısına göre otomatik: az kişilik
    ekipte satırlar büyür (en fazla ~15 mm), 20 kişide ~6,5 mm. Çok sayfalı
    ekipte bütün sayfalar aynı (20 satırlık) ölçüyü kullanır.
  - Tarih elle yazılır (basılı tarih yok). Son sayfaya, listede olmayanlar
    için boş satırlar eklenir.
  - Kişi sırası Ekip Listesi ile aynıdır (team_report.sort_roster:
    Formen > Ekip Başı > diğer görevler > Bayrakçı).

İzin kodları puantaj_engine.LEAVE_CODE_MAP ile aynıdır (DV = Devamsızlık
dahil) - formdaki kod aktarımda otomatik tanınır.

Gereken paket: reportlab  (pip install reportlab)
Yazı tipleri: fonts/ klasöründeki DejaVu Sans (Türkçe karakterler için);
bulunamazsa sistemdeki DejaVu / Liberation / Arial denenir.
"""

__version__ = "2026-09-25.1"

import datetime
from pathlib import Path

import team_report as report
import puantaj_engine as engine

PROJE_ADI = "MAOG Projesi"
SAYFA_MAX_KISI = 20
VARSAYILAN_BOS_SATIR = 3

SCRIPT_DIR = Path(__file__).resolve().parent
FONT_DIR = SCRIPT_DIR / "fonts"

# px (tasarım, 96 dpi) -> pt (PDF, 72 dpi)
PX = 0.75

# Renkler (tasarımla aynı)
MUREKKEP = "#16202B"
LACIVERT = "#16344F"
GRI_METIN = "#4A5460"
GRI_CIZGI = "#9AA1A9"
ACIK_CIZGI = "#C9CDD2"
BASLIK_ZEMIN = "#ECE8E1"

IZIN_KODLARI = [("Yİ", "Yıllık izin"), ("RP", "Sağlık raporu"), ("ÜCS", "Ücretsiz izin"), ("Bİ", "Babalık izni"),
                ("Vİ", "Vefat izni"), ("HT", "Hafta tatili"), ("DV", "Devamsızlık")]


class MesaiFormuHatasi(Exception):
    pass


# ---------------------------------------------------------------- yazı tipleri

_FONT_ADAYLARI = {
    "normal": [FONT_DIR / "DejaVuSansCondensed.ttf",
               "/usr/share/fonts/dejavu-sans-fonts/DejaVuSansCondensed.ttf",
               "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
               "/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf",
               "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
               "C:/Windows/Fonts/arialn.ttf", "C:/Windows/Fonts/arial.ttf"],
    "kalin": [FONT_DIR / "DejaVuSansCondensed-Bold.ttf",
              "/usr/share/fonts/dejavu-sans-fonts/DejaVuSansCondensed-Bold.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed-Bold.ttf",
              "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf",
              "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
              "C:/Windows/Fonts/arialnb.ttf", "C:/Windows/Fonts/arialbd.ttf"],
    "mono": [FONT_DIR / "DejaVuSansMono.ttf",
             "/usr/share/fonts/dejavu-sans-mono-fonts/DejaVuSansMono.ttf",
             "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
             "/usr/share/fonts/liberation-mono/LiberationMono-Regular.ttf",
             "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
             "C:/Windows/Fonts/consola.ttf", "C:/Windows/Fonts/cour.ttf"],
}
_FONTLAR = {}


def _fontlari_yukle():
    if _FONTLAR:
        return _FONTLAR
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError as e:
        raise MesaiFormuHatasi("PDF için 'reportlab' paketi gerekli: pip install reportlab") from e
    for tur, adaylar in _FONT_ADAYLARI.items():
        ad = f"VM-{tur}"
        for yol in adaylar:
            yol = Path(yol)
            if yol.exists():
                try:
                    pdfmetrics.registerFont(TTFont(ad, str(yol)))
                    _FONTLAR[tur] = ad
                    break
                except Exception:
                    continue
    eksik = [t for t in _FONT_ADAYLARI if t not in _FONTLAR]
    if eksik:
        raise MesaiFormuHatasi(
            "Türkçe karakter destekleyen yazı tipi bulunamadı. Uygulama klasöründeki 'fonts' klasörünün "
            "(DejaVuSansCondensed.ttf, DejaVuSansCondensed-Bold.ttf, DejaVuSansMono.ttf) yerinde olduğundan emin olun.")
    return _FONTLAR


# ---------------------------------------------------------------- veri hazırlığı

def formen_bul(kayitlar):
    for anahtar in (report._FORMEN_KEY, report._EKIPBASI_KEY):
        for r in kayitlar:
            if anahtar in report._role_norm_nospace(r["role"]):
                return r["name"]
    return ""


def ekiplere_bol(kayitlar):
    """[(grup, alt_ekip, sıralı_kayıtlar)] - her alt ekip ayrı form."""
    gruplar = {}
    for r in kayitlar:
        gruplar.setdefault((r["team"], r["subteam"]), []).append(r)
    sonuc = []
    for (grup, alt), recs in sorted(gruplar.items(), key=lambda kv: (engine.normalize_name(kv[0][0]),
                                                                     engine.normalize_name(kv[0][1]))):
        sirali, _ = report.sort_roster(recs)
        sonuc.append((grup, alt, sirali))
    return sonuc


def sayfa_plani(kisi_sayisi, bos_satir=VARSAYILAN_BOS_SATIR):
    """Döner: [(ilk_index, son_index_haric, bos_satir_sayisi)], satır_yüksekliği_px, yazı_px"""
    sayfa_sayisi = max(1, -(-kisi_sayisi // SAYFA_MAX_KISI))
    sayfalar = []
    for s in range(sayfa_sayisi):
        a = s * SAYFA_MAX_KISI
        b = min(kisi_sayisi, a + SAYFA_MAX_KISI)
        bos = min(bos_satir, SAYFA_MAX_KISI - (b - a)) if s == sayfa_sayisi - 1 else 0
        sayfalar.append((a, b, bos))
    olcu = SAYFA_MAX_KISI if sayfa_sayisi > 1 else (sayfalar[0][1] - sayfalar[0][0] + sayfalar[0][2])
    alan = 492
    h = max(22, min(56, alan // max(olcu, 1)))
    yazi = 18 if h >= 44 else 17 if h >= 36 else 16 if h >= 30 else 15 if h >= 26 else 14
    return sayfalar, h, yazi


def form_ozeti(kayitlar, bos_satir=VARSAYILAN_BOS_SATIR):
    """Arayüz için: kaç ekip, kaç sayfa."""
    ekipler = ekiplere_bol(kayitlar)
    sayfa = sum(len(sayfa_plani(len(k), bos_satir)[0]) for _g, _a, k in ekipler)
    return {"ekip": len(ekipler), "sayfa": sayfa}


# ---------------------------------------------------------------- çizim yardımcıları

def _renk(hex_):
    from reportlab.lib.colors import HexColor
    return HexColor(hex_)


def _sigdir(c, metin, font, boyut, genislik):
    """Metni verilen genişliğe sığdırır (gerekirse '…' ile kısaltır)."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    metin = str(metin or "")
    if stringWidth(metin, font, boyut) <= genislik:
        return metin
    while metin and stringWidth(metin + "…", font, boyut) > genislik:
        metin = metin[:-1]
    return metin + "…"


def _sar(metin, font, boyut, genislik):
    """Basit kelime kaydırma - satır listesi döner."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    satirlar, simdiki = [], ""
    for kelime in metin.split(" "):
        deneme = (simdiki + " " + kelime).strip()
        if stringWidth(deneme, font, boyut) <= genislik:
            simdiki = deneme
        else:
            if simdiki:
                satirlar.append(simdiki)
            simdiki = kelime
    if simdiki:
        satirlar.append(simdiki)
    return satirlar


# ---------------------------------------------------------------- sayfa çizimi

def _sayfa_ciz(c, F, grup, alt_ekip, kisi_toplam, satirlar, h_px, yazi_px, sayfa_no, sayfa_sayisi, tc_goster):
    from reportlab.lib.pagesizes import A4, landscape
    W, H = landscape(A4)                       # 841.9 x 595.3 pt
    M = 40 * PX                                # 30 pt kenar boşluğu
    x0, x1 = M, W - M
    genislik = x1 - x0
    y = H - M                                  # üstten aşağı doğru çiziyoruz

    # ---- başlık bloğu
    bas_h = 48 * PX
    c.setFillColor(_renk(LACIVERT))
    c.setFont(F["kalin"], 17.5)
    baslik = "GÜNLÜK MESAİ FORMU"
    c.drawString(x0, y - 20, baslik)
    from reportlab.pdfbase.pdfmetrics import stringWidth
    baslik_w = stringWidth(baslik, F["kalin"], 17.5)
    c.setFillColor(_renk(GRI_METIN))
    c.setFont(F["kalin"], 10)
    c.drawString(x0, y - 33, PROJE_ADI)

    tarih_w = 200 * PX
    kutu_x = x0 + baslik_w + 16
    kutu_w = x1 - tarih_w - 9 - kutu_x
    c.setStrokeColor(_renk(MUREKKEP))
    c.setLineWidth(1.1)
    c.rect(kutu_x, y - bas_h, kutu_w, bas_h)
    c.rect(x1 - tarih_w, y - bas_h, tarih_w, bas_h)
    alanlar = [("GRUP", grup, 1.0), ("ALT EKİP", alt_ekip, 1.3), ("PERSONEL", f"{kisi_toplam} kişi", 0.8)]
    toplam_oran = sum(a[2] for a in alanlar)
    fx = kutu_x
    c.setLineWidth(0.75)
    for i, (etiket, deger, oran) in enumerate(alanlar):
        w = kutu_w * oran / toplam_oran
        if i:
            c.line(fx, y - bas_h, fx, y)
        c.setFillColor(_renk(GRI_METIN))
        c.setFont(F["kalin"], 8)
        c.drawString(fx + 6, y - 12, etiket)
        c.setFillColor(_renk(MUREKKEP))
        c.setFont(F["kalin"], 11.5)
        c.drawString(fx + 6, y - 27, _sigdir(c, deger or "—", F["kalin"], 11.5, w - 12))
        fx += w
    c.setFillColor(_renk(GRI_METIN))
    c.setFont(F["kalin"], 8)
    c.drawString(x1 - tarih_w + 6, y - 12, "TARİH")
    c.setFillColor(_renk(GRI_CIZGI))
    c.setFont(F["mono"], 11.5)
    c.drawString(x1 - tarih_w + 6, y - bas_h + 8, "___ / ___ / 20___")
    y -= bas_h + 12 * PX

    # ---- tablo
    sabit = {0: 30, 2: 112, 3: 120, 4: 62, 5: 62, 6: 58, 7: 58, 8: 64, 10: 84}
    esnek = (genislik / PX) - sum(sabit.values())
    gen_px = [sabit.get(i, None) for i in range(11)]
    gen_px[1] = esnek * 1.4 / 2.4
    gen_px[9] = esnek * 1.0 / 2.4
    gen = [g * PX for g in gen_px]
    xs = [x0]
    for g in gen:
        xs.append(xs[-1] + g)
    xs[-1] = x1

    th1, th2 = 24 * PX, 20 * PX
    satir_h = h_px * PX
    tablo_ust = y
    tablo_alt = y - th1 - th2 - satir_h * len(satirlar)

    # başlık zemini
    c.setFillColor(_renk(BASLIK_ZEMIN))
    c.rect(x0, y - th1 - th2, genislik, th1 + th2, stroke=0, fill=1)
    c.setFillColor(_renk(MUREKKEP))

    def orta(metin, xa, xb, yy, font, boyut, renk=MUREKKEP):
        c.setFillColor(_renk(renk))
        c.setFont(font, boyut)
        c.drawCentredString((xa + xb) / 2, yy, metin)

    bas_y_tek = y - (th1 + th2) / 2 - 3.2
    tek = {0: "Sıra", 2: "TC Kimlik No", 3: "Görevi", 9: "Açıklama", 10: "İmza"}
    for i, m in tek.items():
        orta(m, xs[i], xs[i + 1], bas_y_tek, F["kalin"], 9)
    c.setFont(F["kalin"], 9)
    c.drawString(xs[1] + 5, bas_y_tek, "Ad Soyad")
    orta("İzin", xs[8], xs[9], bas_y_tek + 5, F["kalin"], 9)
    orta("kodu", xs[8], xs[9], bas_y_tek - 5, F["kalin"], 9)
    orta("Çalışma saati", xs[4], xs[6], y - th1 / 2 - 3.2, F["kalin"], 9)
    orta("Süre (saat)", xs[6], xs[8], y - th1 / 2 - 3.2, F["kalin"], 9)
    alt_y = y - th1 - th2 / 2 - 3
    for i, m in ((4, "Başlama"), (5, "Bitiş"), (6, "Normal"), (7, "FM")):
        orta(m, xs[i], xs[i + 1], alt_y, F["kalin"], 8)

    # satırlar
    yazi_pt = yazi_px * PX
    tc_pt = min(13, yazi_px - 2) * PX
    gorev_pt = (yazi_px - 1) * PX
    ry = y - th1 - th2
    for r in satirlar:
        orta_y = ry - satir_h / 2 - yazi_pt * 0.35
        orta(r["no"], xs[0], xs[1], ry - satir_h / 2 - 3, F["mono"], 9, GRI_METIN)
        if r["ad"]:
            c.setFillColor(_renk(MUREKKEP))
            c.setFont(F["kalin"], yazi_pt)
            c.drawString(xs[1] + 5, orta_y, _sigdir(c, r["ad"], F["kalin"], yazi_pt, gen[1] - 10))
            tc = r["tc"] if tc_goster else (r["tc"][:3] + "•••••" + r["tc"][-3:] if len(r["tc"]) > 6 else r["tc"])
            orta(tc, xs[2], xs[3], ry - satir_h / 2 - tc_pt * 0.35, F["mono"], tc_pt)
            c.setFillColor(_renk(MUREKKEP))
            c.setFont(F["normal"], gorev_pt)
            c.drawString(xs[3] + 5, ry - satir_h / 2 - gorev_pt * 0.35,
                         _sigdir(c, r["gorev"], F["normal"], gorev_pt, gen[3] - 10))
        orta("__:__", xs[4], xs[5], ry - satir_h / 2 - 3, F["mono"], 9, GRI_CIZGI)
        orta("__:__", xs[5], xs[6], ry - satir_h / 2 - 3, F["mono"], 9, GRI_CIZGI)
        ry -= satir_h
        # yatay ince çizgi
        c.setStrokeColor(_renk(GRI_CIZGI))
        c.setLineWidth(0.5)
        c.line(x0, ry, x1, ry)

    # dikey çizgiler (ince: alt bölümler arası)
    ince = {5, 7}
    for i in range(1, 11):
        c.setStrokeColor(_renk(GRI_CIZGI if i in ince else MUREKKEP))
        c.setLineWidth(0.5 if i in ince else 0.75)
        ust = y - th1 if i in ince else y
        c.line(xs[i], tablo_alt, xs[i], ust)
    # başlık ayırıcıları
    c.setStrokeColor(_renk(MUREKKEP))
    c.setLineWidth(0.75)
    c.line(xs[4], y - th1, xs[8], y - th1)
    c.line(x0, y - th1 - th2, x1, y - th1 - th2)
    # dış çerçeve
    c.setLineWidth(1.1)
    c.rect(x0, tablo_alt, genislik, tablo_ust - tablo_alt)

    # ---- alt bölüm (sayfanın dibine yaslı)
    alt_h = 92 * PX
    ay = M + 20                                   # en alttaki sayfa satırı için yer
    sol_w = genislik * 1.6 / 4.6 - 7
    imza_kutulari = ("Hazırlayan", "Kontrol eden")
    kutu_w2 = (genislik - sol_w - len(imza_kutulari) * 10 * PX) / len(imza_kutulari)
    # izin kodları
    c.setFillColor(_renk(GRI_METIN))
    c.setFont(F["kalin"], 8)
    c.drawString(x0, ay + alt_h - 8, "İZİN KODLARI — İZİN KODU SÜTUNUNA YAZINIZ")
    ty = ay + alt_h - 20
    # kodları öğe öğe satıra yerleştir (bir kod ile açıklaması asla bölünmez)
    satirlar_k, simdiki, simdiki_w = [], [], 0.0
    ayrac_w = stringWidth(" · ", F["normal"], 8.5)
    for kod, ad in IZIN_KODLARI:
        w = stringWidth(kod + " ", F["kalin"], 8.5) + stringWidth(ad, F["normal"], 8.5)
        ek = w + (ayrac_w if simdiki else 0)
        if simdiki and simdiki_w + ek > sol_w:
            satirlar_k.append(simdiki)
            simdiki, simdiki_w, ek = [], 0.0, w
        simdiki.append((kod, ad))
        simdiki_w += ek
    if simdiki:
        satirlar_k.append(simdiki)
    for ogeler in satirlar_k:
        xx = x0
        for j, (kod, ad) in enumerate(ogeler):
            c.setFillColor(_renk(MUREKKEP))
            c.setFont(F["kalin"], 8.5)
            c.drawString(xx, ty, kod)
            xx += stringWidth(kod + " ", F["kalin"], 8.5)
            c.setFont(F["normal"], 8.5)
            c.drawString(xx, ty, ad)
            xx += stringWidth(ad, F["normal"], 8.5)
            if j < len(ogeler) - 1:
                c.drawString(xx, ty, " · ")
                xx += ayrac_w
        ty -= 11
    ty -= 2
    c.setFillColor(_renk(GRI_METIN))
    aciklama = ("Saatleri rakamla yazınız (ör. 9 · 2,5). Normal mesai en fazla 9 saattir; fazlası FM sütununa "
                "yazılır. Listede olmayan personeli boş satırlara TC'si ile ekleyiniz.")
    for satir in _sar(aciklama, F["normal"], 8.5, sol_w):
        c.drawString(x0, ty, satir)
        ty -= 11

    # imza kutuları
    kx = x0 + sol_w + 10 * PX
    for baslik_ in imza_kutulari:
        c.setStrokeColor(_renk(MUREKKEP))
        c.setLineWidth(0.75)
        c.setFillColor(_renk(BASLIK_ZEMIN))
        c.rect(kx, ay + alt_h - 16, kutu_w2, 16, stroke=0, fill=1)
        c.rect(kx, ay, kutu_w2, alt_h)
        c.line(kx, ay + alt_h - 16, kx + kutu_w2, ay + alt_h - 16)
        c.setFillColor(_renk(MUREKKEP))
        c.setFont(F["kalin"], 9)
        c.drawString(kx + 6, ay + alt_h - 11.5, baslik_)
        # Ad Soyad / Tarih / İmza satırları
        c.setStrokeColor(_renk(ACIK_CIZGI))
        c.setLineWidth(0.5)
        yy = ay + alt_h - 16
        for i, (etiket, yuk) in enumerate((("Ad Soyad", 15), ("Tarih", 15), ("İmza", 23))):
            yy -= yuk
            c.setFillColor(_renk(GRI_METIN))
            c.setFont(F["normal"], 8)
            c.drawString(kx + 6, yy + 3, etiket)
            if i < 2:
                c.line(kx, yy, kx + kutu_w2, yy)
        kx += kutu_w2 + 10 * PX

    # ---- en alt satır
    c.setFillColor(_renk(GRI_METIN))
    c.setFont(F["normal"], 8)
    c.drawString(x0, M - 2, f"{grup} — {alt_ekip}")
    if sayfa_sayisi > 1:
        c.setFillColor(_renk(MUREKKEP))
        c.setFont(F["kalin"], 8.5)
        c.drawRightString(x1, M - 2, f"Sayfa {sayfa_no} / {sayfa_sayisi}")


def pdf_olustur(kayitlar, output_path, bos_satir=VARSAYILAN_BOS_SATIR, tc_goster=True):
    """kayitlar: veri_merkezi_ekip_listesi.build_roster()['records'] biçiminde
    [{tc, name, role, team, subteam}]. Döner: {"yol", "sayfa", "ekip"}"""
    if not kayitlar:
        raise MesaiFormuHatasi("Seçilen filtrelerle eşleşen personel yok.")
    F = _fontlari_yukle()
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=landscape(A4))
    c.setTitle("Günlük Mesai Formu")
    c.setAuthor("Veri Merkezi")
    c.setSubject(PROJE_ADI)

    toplam_sayfa = 0
    ekipler = ekiplere_bol(kayitlar)
    for grup, alt, recs in ekipler:
        sayfalar, h, yazi = sayfa_plani(len(recs), bos_satir)
        for s_no, (a, b, bos) in enumerate(sayfalar, start=1):
            satirlar = [{"no": str(i + 1), "ad": recs[i]["name"], "tc": recs[i]["tc"], "gorev": recs[i]["role"]}
                        for i in range(a, b)]
            satirlar += [{"no": str(b + j + 1), "ad": "", "tc": "", "gorev": ""} for j in range(bos)]
            _sayfa_ciz(c, F, grup, alt, len(recs), satirlar, h, yazi, s_no, len(sayfalar), tc_goster)
            c.showPage()
            toplam_sayfa += 1
    c.save()
    return {"yol": output_path, "sayfa": toplam_sayfa, "ekip": len(ekipler)}


def dosya_adi(kapsam):
    return f"{report._slugify(kapsam)}_gunluk_mesai_formu_{datetime.date.today():%Y%m%d}.pdf"
