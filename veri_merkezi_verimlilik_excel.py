# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - VERİMLİLİK ANALİZİ EXCEL RAPORU (veri_merkezi_verimlilik_excel.py)
===================================================================================
veri_merkezi_verimlilik.analiz() sözlüğünden yönetime sunulabilir Excel üretir.
Ekrandaki "Verimlilik Analizi" sayfasıyla aynı bölümler:

  1 Özet              Verimlilik puanı, 10 gösterge (önceki döneme göre), öne
                      çıkanlar, dönem karşılaştırması, gün dağılımı grafiği
  2 Trend             Periyot tablosu + normal/fazla mesai ve devam grafiği
  3 Günlük            Günlük iş gücü tablosu + yığılmış grafik
  4 Ekip              Ekip / grup karşılaştırması (canlı formüller) + puan grafiği
  5 Görev             Görev karşılaştırması
  6 Devam Haritası    Ekip × gün devam oranı ısı haritası
  7 Personel Karnesi  Herkes için tüm göstergeler (süzme / sıralama hazır)
  8 Devam Takvimi     Kişi × gün: saat ya da izin kodu, renkli
  9 Uyum ve Risk      yıllık / günlük sınır, kesintisiz çalışma, FM eşiği, devamsızlık listeleri (eşikler: Ayarlar)
 10 Personel Hareketi İşe giren / ayrılan
 11 Yöntem            Tanımlar ve formüller

Tablolardaki oranlar, toplamlar, yevmiye ve puan HÜCRE FORMÜLÜ olarak yazılır:
bir saat düzeltilirse satırın bütün göstergeleri Excel'de kendiliğinden güncellenir.
"""

__version__ = "2026-09-25.1"

import datetime
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

import team_report as report
import veri_merkezi_verimlilik as vm

# ---------------------------------------------------------------- görsel dil (uygulama ile aynı palet)
NAVY, NAVY_D, NAVY_L, NAVY_BG = "1F4E78", "16344F", "D6E2EE", "E6ECF3"
ORANGE, ORANGE_T, ORANGE_BG = "E8762C", "9A3A0B", "FDF3EA"
GREEN, GREEN_BG, RED, RED_BG = "2F6B3A", "E3EFE5", "B42318", "F9E1DE"
PURPLE, INK, INK2, MUTED = "5B3386", "16202B", "3E4751", "5A6572"
GROUND, GROUND2, LINE, LINE_SOFT = "F4F2EE", "FAF8F5", "DDD8CF", "EFEBE4"
FONT = "Calibri"
PROJE = "MAOG Projesi"

TON_RENK = {"iyi": GREEN, "uyari": ORANGE, "kritik": RED, "bilgi": NAVY}
TON_ZEMIN = {"iyi": GREEN_BG, "uyari": ORANGE_BG, "kritik": RED_BG, "bilgi": NAVY_BG}
TON_ETIKET = {"iyi": "OLUMLU", "uyari": "DİKKAT", "kritik": "KRİTİK", "bilgi": "BİLGİ"}


def _fill(c):
    return PatternFill(start_color=c, end_color=c, fill_type="solid")


def _font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


ALT_CIZGI = Border(bottom=Side(style="thin", color=LINE_SOFT))
KUTU = Border(left=Side(style="thin", color=LINE), right=Side(style="thin", color=LINE),
              top=Side(style="thin", color=LINE), bottom=Side(style="thin", color=LINE))
UST_VURGU = Side(style="thick", color=ORANGE)

F_INT, F_1, F_2 = "#,##0", "#,##0.0", "#,##0.00"
F_PCT = "0.0%"
F_TL = '#,##0" TL"'
F_TARIH = "DD.MM.YYYY"
F_DELTA = '+0.0%;-0.0%;0.0%'


# ---------------------------------------------------------------- ortak yardımcılar

def _yaz(ws, r, c, v, fmt=None, font=None, fill=None, align=None, border=None):
    cell = ws.cell(row=r, column=c, value=v)
    if fmt:
        cell.number_format = fmt
    cell.font = font or _font()
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    if border:
        cell.border = border
    return cell


def _sayfa_basligi(ws, son_kol, baslik, d, alt_ek=None):
    """1. satır: koyu lacivert başlık bandı. 2. satır: kapsam · dönem · oluşturma. 3. satır: turuncu ince çizgi."""
    ws.sheet_view.showGridLines = False
    for c in range(1, son_kol + 1):
        ws.cell(row=1, column=c).fill = _fill(NAVY_D)
        ws.cell(row=2, column=c).fill = _fill(GROUND)
        ws.cell(row=3, column=c).fill = _fill(ORANGE)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=son_kol)
    _yaz(ws, 1, 1, f"  {baslik}", font=_font(16, True, "FFFFFF"), fill=_fill(NAVY_D),
         align=Alignment(vertical="center"))
    ws.row_dimensions[1].height = 34
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=son_kol)
    alt = (f"  {PROJE}  ·  {vm.kapsam_adi(d)}  ·  "
           f"{vm.donem_adi(d)} ({d['gun_sayisi']} gün)  ·  oluşturma {d['olusturma']:%d.%m.%Y %H:%M}")
    if alt_ek:
        alt += f"  ·  {alt_ek}"
    _yaz(ws, 2, 1, alt, font=_font(10, False, MUTED, True), fill=_fill(GROUND), align=Alignment(vertical="center"))
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 3


def _bolum(ws, r, c1, c2, metin, alt=None):
    """Bölüm başlığı: kalın lacivert yazı + alt çizgi."""
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    _yaz(ws, r, c1, metin.upper(), font=_font(11, True, NAVY_D), align=Alignment(vertical="bottom"))
    for c in range(c1, c2 + 1):
        ws.cell(row=r, column=c).border = Border(bottom=Side(style="medium", color=NAVY))
    ws.row_dimensions[r].height = 22
    if alt:
        ws.merge_cells(start_row=r + 1, start_column=c1, end_row=r + 1, end_column=c2)
        _yaz(ws, r + 1, c1, alt, font=_font(9, False, MUTED, True))
        return r + 2
    return r + 1


def _baslik_satiri(ws, r, c1, basliklar, yukseklik=32):
    for i, h in enumerate(basliklar):
        _yaz(ws, r, c1 + i, h, font=_font(9.5, True, "FFFFFF"), fill=_fill(NAVY),
             align=Alignment(horizontal="center", vertical="center", wrap_text=True),
             border=Border(left=Side(style="thin", color="3A6A94"), right=Side(style="thin", color="3A6A94")))
    ws.row_dimensions[r].height = yukseklik


def _veri_satiri(ws, r, c1, c2, zebra=False, fill=None):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=r, column=c)
        cell.border = ALT_CIZGI
        if fill is not None:
            cell.fill = fill
        elif zebra:
            cell.fill = _fill(GROUND2)


def _toplam_satiri(ws, r, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=r, column=c)
        cell.fill = _fill(NAVY_D)
        cell.font = _font(10, True, "FFFFFF")
        cell.border = Border(top=Side(style="medium", color=ORANGE))
    ws.row_dimensions[r].height = 20


def _genislik(ws, genislikler):
    for i, w in enumerate(genislikler, start=1):
        ws.column_dimensions[L(i)].width = w


def _yazdir(ws, yatay=True, basliklar=None):
    ws.page_setup.orientation = "landscape" if yatay else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top, ws.page_margins.bottom = 0.6, 0.6
    ws.oddHeader.left.text = "&8MAOG Projesi — Verimlilik Analizi"
    ws.oddHeader.right.text = "&8&A"
    ws.oddFooter.left.text = "&8Veri Merkezi"
    ws.oddFooter.center.text = "&8Sayfa &P / &N"
    ws.oddFooter.right.text = "&8&D"
    if basliklar:
        ws.print_title_rows = basliklar


def _devam_olcegi(ws, aralik):
    ws.conditional_formatting.add(aralik, ColorScaleRule(start_type="num", start_value=0.8, start_color="F4B6AE",
                                                         mid_type="num", mid_value=0.92, mid_color="FBE7B5",
                                                         end_type="num", end_value=1, end_color="A9D3AF"))


def _puan_kurallari(ws, puan_aralik, not_aralik):
    ws.conditional_formatting.add(puan_aralik, ColorScaleRule(start_type="num", start_value=70, start_color="F4B6AE",
                                                              mid_type="num", mid_value=88, mid_color="FBE7B5",
                                                              end_type="num", end_value=98, end_color="A9D3AF"))
    for harf, zemin, renk in (("A", GREEN_BG, GREEN), ("B", NAVY_BG, NAVY_D), ("C", ORANGE_BG, ORANGE_T),
                              ("D", RED_BG, RED)):
        ws.conditional_formatting.add(not_aralik, CellIsRule(operator="equal", formula=[f'"{harf}"'], fill=_fill(zemin),
                                                             font=Font(name=FONT, bold=True, color=renk)))


def _delta_metni(simdi, once, iyi_yon, tip):
    """(metin, renk) - KPI kutuları için önceki döneme göre değişim."""
    if once is None:
        return "önceki dönem verisi yok", MUTED
    if tip == "oran":
        if not once:
            return "—", MUTED
        fark = (simdi - once) / once * 100
        metin = f"%{report.format_tr(abs(fark), 1)}"
    else:
        fark = simdi - once
        metin = f"{report.format_tr(abs(fark), 1)} puan"
    if abs(fark) < 0.05:
        return "önceki dönemle aynı", MUTED
    ok = "▲" if fark > 0 else "▼"
    renk = MUTED if not iyi_yon else (GREEN if (fark > 0) == (iyi_yon > 0) else RED)
    return f"{ok} {metin}  önceki döneme göre", renk


# ================================================================ 1) ÖZET

def _ozet(wb, d, ucret):
    ws = wb.active
    ws.title = "Özet"
    ws.sheet_properties.tabColor = NAVY_D
    SON = 11
    _genislik(ws, [2.5] + [13.5] * 10)
    _sayfa_basligi(ws, SON, "VERİMLİLİK ANALİZİ — YÖNETİCİ ÖZETİ", d)
    g, o = d["genel"], d["onceki"]
    o1, o2 = d["onceki_aralik"]

    # ---- puan bandı (satır 5-8)
    harf, bant = vm.puan_bandi(g["puan"])
    renk_zemin = {"A": GREEN_BG, "B": NAVY_BG, "C": ORANGE_BG, "D": RED_BG}[harf]
    renk_yazi = {"A": GREEN, "B": NAVY_D, "C": ORANGE_T, "D": RED}[harf]
    for r in range(5, 9):
        for c in range(2, SON + 1):
            ws.cell(row=r, column=c).fill = _fill(renk_zemin)
    ws.merge_cells("B5:D5")
    _yaz(ws, 5, 2, "VERİMLİLİK PUANI", font=_font(9, True, MUTED), fill=_fill(renk_zemin))
    ws.merge_cells("B6:C8")
    _yaz(ws, 6, 2, round(g["puan"], 1), F_1, _font(36, True, renk_yazi), _fill(renk_zemin),
         Alignment(horizontal="left", vertical="center", indent=1))
    ws.merge_cells("D6:D8")
    _yaz(ws, 6, 4, f"{harf}\n{bant}", None, _font(16, True, renk_yazi), _fill(renk_zemin),
         Alignment(horizontal="center", vertical="center", wrap_text=True))
    dm, dr = _delta_metni(g["puan"], o["puan"] if o else None, 1, "puan")
    ws.merge_cells("E5:K5")
    _yaz(ws, 5, 5, "Puan = devam oranı × 0,6 + kapasite kullanımı × 0,4   ·   A ≥ 95 · B 85–95 · C 70–85 · D < 70",
         font=_font(9, False, MUTED, True), fill=_fill(renk_zemin))
    satirlar = [("Devam oranı", g["devam"] / 100, F_PCT), ("Kapasite kullanımı", min(g["kapasite"], 100) / 100, F_PCT)]
    for i, (e, v, f) in enumerate(satirlar):
        ws.merge_cells(start_row=6 + i, start_column=5, end_row=6 + i, end_column=7)
        _yaz(ws, 6 + i, 5, e, font=_font(10, True, INK2), fill=_fill(renk_zemin))
        _yaz(ws, 6 + i, 8, v, f, _font(11, True, INK), _fill(renk_zemin), Alignment(horizontal="right"))
    ws.merge_cells("E8:K8")
    _yaz(ws, 8, 5, dm, font=_font(10, True, dr), fill=_fill(renk_zemin))
    ws.merge_cells("I6:K7")
    _yaz(ws, 6, 9, f"Karşılaştırma dönemi\n{o1:%d.%m.%Y} – {o2:%d.%m.%Y}", font=_font(9, False, MUTED, True),
         fill=_fill(renk_zemin), align=Alignment(horizontal="right", vertical="center", wrap_text=True))
    for r in (5, 6, 7, 8):
        ws.row_dimensions[r].height = 20

    # ---- 10 gösterge kutusu (2 satır × 5 kutu, her kutu 2 sütun × 3 satır)
    og = (lambda k: o[k] if o else None)
    kutular = [
        ("GÜNLÜK ORT. ÇALIŞAN", g["ort_calisan"], F_INT, f"{report.format_tr(g['kisi'], 0)} farklı personel",
         _delta_metni(g["ort_calisan"], og("ort_calisan"), 0, "oran"), NAVY),
        ("DEVAM ORANI", g["devam"] / 100, F_PCT, f"{report.format_tr(g['calisilan'], 0)} / "
                                                 f"{report.format_tr(g['planlanan'], 0)} iş günü",
         _delta_metni(g["devam"], og("devam"), 1, "puan"), NAVY),
        ("KAPASİTE KULLANIMI", g["kapasite"] / 100, F_PCT, f"kayıp {report.format_tr(g['kayip_saat'], 0)} saat",
         _delta_metni(g["kapasite"], og("kapasite"), 1, "puan"), NAVY),
        ("KİŞİ BAŞI GÜNLÜK SAAT", g["ort_saat"], F_2, "normal + fazla, çalışılan gün başı",
         _delta_metni(g["ort_saat"], og("ort_saat"), 0, "oran"), NAVY),
        ("KAYIP İŞ GÜNÜ", g["kayip"], F_INT, f"%{report.format_tr(g['kayip_oran'], 1)} · devamsızlık "
                                            f"{g['gun']['devamsiz']} · rapor {g['gun']['rapor']}",
         _delta_metni(g["kayip_oran"], og("kayip_oran"), -1, "puan"), RED),
        ("FAZLA MESAİ (SAAT)", g["fazla"], F_INT, f"pazar ×2,5 {report.format_tr(g['pazar_katli'], 0)} sa · kişi başı "
                                                 f"{report.format_tr(g['kisi_basi_fm'], 1)} sa",
         _delta_metni(g["fazla"], og("fazla"), -1, "oran"), ORANGE),
        ("FM ORANI", g["fm_oran"] / 100, F_PCT, "toplam çalışma saati içinde",
         _delta_metni(g["fm_oran"], og("fm_oran"), -1, "puan"), ORANGE),
        ("YEVMİYE KARŞILIĞI", g["yevmiye"], F_1,
         f"≈ {report.format_tr(g['yevmiye'] * ucret, 0)} TL" if ucret else f"(normal + FM × katsayı) ÷ {vm._sa(vm.STANDART_GUN)}",
         _delta_metni(g["yevmiye"], og("yevmiye"), 0, "oran"), NAVY),
        ("FM YEVMİYE KARŞILIĞI", g["fm_yevmiye"], F_1,
         (f"≈ {report.format_tr(g['fm_yevmiye'] * ucret, 0)} TL · ek {report.format_tr(g['fm_prim'] * ucret, 0)} TL"
          if ucret else f"katsayı farkı {report.format_tr(g['fm_prim'], 1)} yevmiye"),
         _delta_metni(g["fm_yevmiye"], og("fm_yevmiye"), -1, "oran"), ORANGE),
        ("UYUM RİSKİ OLAN PERSONEL", d["risk"]["riskli_kisi"], F_INT,
         f"{len(d['risk']['uzun_gunler'])} kez {vm.etiket_gunluk()} sa aşımı", ("ayrıntı: Uyum ve Risk sayfası", MUTED), RED),
    ]
    for i, (etiket, deger, fmt, alt, (dm, dr), vurgu) in enumerate(kutular):
        satir = 10 + (i // 5) * 4
        kol = 2 + (i % 5) * 2
        for r in range(satir, satir + 3):
            for c in (kol, kol + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = _fill("FFFFFF")
                cell.border = Border(left=Side(style="thin", color=LINE) if c == kol else None,
                                     right=Side(style="thin", color=LINE) if c == kol + 1 else None,
                                     top=Side(style="thick", color=vurgu) if r == satir else None,
                                     bottom=Side(style="thin", color=LINE) if r == satir + 2 else None)
        ws.merge_cells(start_row=satir, start_column=kol, end_row=satir, end_column=kol + 1)
        _yaz(ws, satir, kol, etiket, font=_font(8.5, True, MUTED), align=Alignment(vertical="bottom", indent=1))
        ws.cell(row=satir, column=kol).border = Border(left=Side(style="thin", color=LINE), top=Side(style="thick", color=vurgu))
        ws.merge_cells(start_row=satir + 1, start_column=kol, end_row=satir + 1, end_column=kol + 1)
        _yaz(ws, satir + 1, kol, round(deger, 4) if isinstance(deger, float) else deger, fmt,
             _font(18, True, INK), align=Alignment(vertical="center", indent=1))
        ws.cell(row=satir + 1, column=kol).border = Border(left=Side(style="thin", color=LINE))
        ws.merge_cells(start_row=satir + 2, start_column=kol, end_row=satir + 2, end_column=kol + 1)
        _yaz(ws, satir + 2, kol, None, align=Alignment(vertical="top", wrap_text=True, indent=1))
        ws.cell(row=satir + 2, column=kol).value = CellRichText(
            TextBlock(InlineFont(rFont=FONT, sz=8, b=True, color=dr), f"{dm}\n"),
            TextBlock(InlineFont(rFont=FONT, sz=8, color=MUTED), alt))
        ws.cell(row=satir + 2, column=kol).border = Border(left=Side(style="thin", color=LINE),
                                                           bottom=Side(style="thin", color=LINE))
        ws.row_dimensions[satir].height = 18
        ws.row_dimensions[satir + 1].height = 28
        ws.row_dimensions[satir + 2].height = 26

    # ---- öne çıkanlar
    r = _bolum(ws, 19, 2, SON, "Öne çıkanlar", "Veriden otomatik üretilmiştir")
    for oz in d["ozet"]:
        _yaz(ws, r, 2, TON_ETIKET[oz["ton"]], font=_font(8, True, "FFFFFF"), fill=_fill(TON_RENK[oz["ton"]]),
             align=Alignment(horizontal="center", vertical="center"))
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=SON)
        _yaz(ws, r, 3, None)
        ws.cell(row=r, column=3).value = CellRichText(
            TextBlock(InlineFont(rFont=FONT, sz=10, b=True, color=INK), f"{oz['baslik']}:  "),
            TextBlock(InlineFont(rFont=FONT, sz=10, color=INK2), oz["metin"]))
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        for c in range(3, SON + 1):
            ws.cell(row=r, column=c).fill = _fill(TON_ZEMIN[oz["ton"]])
        satir_say = max(1, -(-(len(oz["metin"]) + len(oz["baslik"]) + 3) // 120))
        ws.row_dimensions[r].height = 15 * satir_say + 8
        r += 1

    # ---- dönem karşılaştırması
    r = _bolum(ws, r + 1, 2, SON, "Dönem karşılaştırması",
               f"Önceki dönem: {o1:%d.%m.%Y} – {o2:%d.%m.%Y} (aynı uzunlukta)")
    bas_r = r
    _baslik_satiri(ws, r, 2, ["Gösterge", "", "", "Bu dönem", "Önceki dönem", "Değişim", "Yorum", "", "", ""], 22)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=SON)
    r += 1
    gostergeler = [
        ("Günlük ortalama çalışan (kişi)", "ort_calisan", F_1, 0), ("Farklı personel", "kisi", F_INT, 0),
        ("Çalışılan adam-gün", "calisilan", F_INT, 0), ("Planlanan iş günü", "planlanan", F_INT, 0),
        ("Devam oranı", "devam", F_PCT, 1), ("Kayıp iş günü", "kayip", F_INT, -1),
        ("  · devamsızlık (gün)", ("gun", "devamsiz"), F_INT, -1), ("  · rapor (gün)", ("gun", "rapor"), F_INT, -1),
        ("  · ücretsiz izin (gün)", ("gun", "ucretsiz_izin"), F_INT, -1),
        ("Kapasite kullanımı", "kapasite", F_PCT, 1), ("Normal mesai (saat)", "normal", F_INT, 0),
        ("Fazla mesai (saat)", "fazla", F_INT, -1), ("  · pazar ×2,5 (saat, normal + FM)", "pazar_katli", F_INT, -1),
        ("FM oranı", "fm_oran", F_PCT, -1), ("Kişi başı günlük saat", "ort_saat", F_2, 0),
        ("Yevmiye karşılığı", "yevmiye", F_1, 0), ("FM yevmiye karşılığı", "fm_yevmiye", F_1, -1),
        ("Verimlilik puanı", "puan", F_1, 1),
    ]
    yuzde = {"devam", "kapasite", "fm_oran"}
    for i, (etiket, anahtar, fmt, yon) in enumerate(gostergeler):
        def al(m):
            if m is None:
                return None
            v = m[anahtar[0]][anahtar[1]] if isinstance(anahtar, tuple) else m[anahtar]
            return v / 100 if anahtar in yuzde else v
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _yaz(ws, r, 2, etiket, font=_font(10, not etiket.startswith("  "), INK if not etiket.startswith("  ") else MUTED),
             align=Alignment(indent=1))
        _yaz(ws, r, 5, al(g), fmt, _font(10, True), align=Alignment(horizontal="right"))
        _yaz(ws, r, 6, al(o), fmt, _font(10, False, MUTED), align=Alignment(horizontal="right"))
        if o is not None:
            if anahtar in yuzde or anahtar == "puan":
                formul = f"=(E{r}-F{r})*100" if anahtar in yuzde else f"=E{r}-F{r}"
                _yaz(ws, r, 7, formul, '+0.0" puan";-0.0" puan";0.0" puan"', _font(10, True),
                     align=Alignment(horizontal="right"))
            else:
                _yaz(ws, r, 7, f'=IF(F{r}=0,"",E{r}/F{r}-1)', F_DELTA, _font(10, True), align=Alignment(horizontal="right"))
            if yon:
                ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=SON)
                iyi = f'=IF(OR(F{r}="",F{r}=0),"",IF(ROUND(E{r}-F{r},4)=0,"değişmedi",IF((E{r}-F{r})*{yon}>0,"iyileşti","kötüleşti")))'
                _yaz(ws, r, 8, iyi, font=_font(9, False, MUTED, True))
        _veri_satiri(ws, r, 2, SON, zebra=(i % 2 == 1))
        r += 1
    son_r = r - 1
    ws.conditional_formatting.add(f"H{bas_r + 1}:H{son_r}", FormulaRule(formula=[f'H{bas_r + 1}="iyileşti"'],
                                                                           font=Font(name=FONT, color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"H{bas_r + 1}:H{son_r}", FormulaRule(formula=[f'H{bas_r + 1}="kötüleşti"'],
                                                                           font=Font(name=FONT, color=RED, bold=True)))

    # ---- gün dağılımı + halka grafik
    r = _bolum(ws, r + 1, 2, SON, "Gün dağılımı", "Kayıtlı bütün kişi-günlerin durumlara göre dağılımı")
    _baslik_satiri(ws, r, 2, ["Durum", "Gün", "Pay"], 22)
    tbl_bas = r + 1
    toplam = sum(g["gun"].values()) or 1
    r += 1
    renkler = []
    for k, e, renk, _kod in vm.KATEGORILER:
        n = g["gun"][k]
        if not n:
            continue
        _yaz(ws, r, 2, e, font=_font(10, True, INK))
        _yaz(ws, r, 3, n, F_INT, align=Alignment(horizontal="right"))
        _yaz(ws, r, 4, f"=C{r}/SUM(C${tbl_bas}:C${tbl_bas + sum(1 for x in g['gun'].values() if x) - 1})", F_PCT,
             align=Alignment(horizontal="right"))
        ws.cell(row=r, column=2).border = Border(left=Side(style="thick", color=renk), bottom=Side(style="thin", color=LINE_SOFT))
        _veri_satiri(ws, r, 3, 4)
        renkler.append(renk)
        r += 1
    tbl_son = r - 1
    ch = DoughnutChart(holeSize=55)
    ch.add_data(Reference(ws, min_col=3, min_row=tbl_bas, max_row=tbl_son), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=2, min_row=tbl_bas, max_row=tbl_son))
    ch.title = None
    s = ch.series[0]
    for i, renk in enumerate(renkler):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = renk
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        s.dPt.append(pt)
    ch.legend.position = "r"
    ch.height, ch.width = 6.4, 13
    ws.add_chart(ch, f"F{tbl_bas - 1}")
    ws.row_dimensions[tbl_son + 1].height = 15
    ws.freeze_panes = "A4"
    _yazdir(ws, yatay=False)
    ws.page_setup.fitToHeight = 1          # özet tek sayfa
    ws.print_area = f"A1:K{max(tbl_son, tbl_bas + 12) + 1}"


# ================================================================ ortak karşılaştırma tablosu (formüllü)

KARS_BASLIK = ["Sıra", "{AD}", "Kişi", "Planlanan gün", "Çalışılan gün", "Devam %", "Kayıp gün", "Kayıp %",
               "Normal sa", "Fazla sa", "Toplam sa", "FM %", "Sa / gün", "Kapasite %", "Pazar ×2,5 sa", "Yevmiye",
               "FM yevmiye", "Önceki dönem FM", "FM değişimi", "Puan", "Not"]


def _kars_satir_formulleri(ws, r, onceki_var):
    """C..U sütunları: E/D devam, G/D kayıp, I+J toplam, J/K FM, K/E sa/gün, I/(D*9) kapasite, puan, not.
    Yevmiye (P) ve FM yevmiye (Q) hafta tatili / pazar kuralına göre Python'da hesaplanıp DEĞER olarak yazılır
    (hangi pazarın ×2,5 olduğu hücre formülüyle çıkarılamaz)."""
    ws[f"F{r}"] = f'=IF(D{r}=0,"",E{r}/D{r})'
    ws[f"H{r}"] = f'=IF(D{r}=0,"",G{r}/D{r})'
    ws[f"K{r}"] = f"=I{r}+J{r}"
    ws[f"L{r}"] = f'=IF(K{r}=0,"",J{r}/K{r})'
    ws[f"M{r}"] = f'=IF(E{r}=0,"",K{r}/E{r})'
    ws[f"N{r}"] = f'=IF(D{r}=0,"",I{r}/(D{r}*{vm.STANDART_GUN}))'
    ws[f"S{r}"] = f'=IF(OR(R{r}="",R{r}=0),"",J{r}/R{r}-1)' if onceki_var else ""
    ws[f"T{r}"] = f'=IF(D{r}=0,"",(0.6*F{r}+0.4*MIN(N{r},1))*100)'
    ws[f"U{r}"] = f'=IF(T{r}="","",IF(T{r}>=95,"A",IF(T{r}>=85,"B",IF(T{r}>=70,"C","D"))))'


def _kars_tablosu(ws, r0, satirlar, ad_basligi, onceki_var):
    basliklar = [h.replace("{AD}", ad_basligi) for h in KARS_BASLIK]
    son = len(basliklar)
    _baslik_satiri(ws, r0, 1, basliklar, 34)
    r = r0 + 1
    ilk = r
    formatlar = {1: F_INT, 3: F_INT, 4: F_INT, 5: F_INT, 6: F_PCT, 7: F_INT, 8: F_PCT, 9: F_INT, 10: F_INT, 11: F_INT,
                 12: F_PCT, 13: F_2, 14: F_PCT, 15: F_INT, 16: F_1, 17: F_1, 18: F_INT, 19: F_DELTA, 20: F_1}
    for i, k in enumerate(sorted(satirlar, key=lambda x: x["sira"])):
        _yaz(ws, r, 1, k["sira"], align=Alignment(horizontal="center"))
        _yaz(ws, r, 2, k["ad"], font=_font(10, True))
        for c, v in ((3, k["kisi"]), (4, k["planlanan"]), (5, k["calisilan"]), (7, k["kayip"]), (9, round(k["normal"], 2)),
                     (10, round(k["fazla"], 2)), (15, round(k["pazar_katli"], 2)), (16, round(k["yevmiye"], 2)),
                     (17, round(k["fm_yevmiye"], 2))):
            _yaz(ws, r, c, v)
        if onceki_var and k["onceki_fazla"] is not None:
            _yaz(ws, r, 18, round(k["onceki_fazla"], 2))
        _kars_satir_formulleri(ws, r, onceki_var)
        for c, f in formatlar.items():
            ws.cell(row=r, column=c).number_format = f
            if c not in (1, 2):
                ws.cell(row=r, column=c).alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=c).font = _font(10, c in (2, 6, 20))
        ws.cell(row=r, column=21).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=21).font = _font(10, True)
        _veri_satiri(ws, r, 1, son, zebra=(i % 2 == 1))
        r += 1
    sonr = r - 1
    # toplam satırı
    _yaz(ws, r, 2, "TOPLAM")
    for c in (3, 4, 5, 7, 9, 10, 15, 16, 17, 18):
        if c == 18 and not onceki_var:
            continue
        ws.cell(row=r, column=c, value=f"=SUM({L(c)}{ilk}:{L(c)}{sonr})")
    _kars_satir_formulleri(ws, r, onceki_var)
    _toplam_satiri(ws, r, 1, son)
    for c, f in formatlar.items():
        ws.cell(row=r, column=c).number_format = f
        if c > 2:
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right")
    ws.cell(row=r, column=1).value = None
    ws.cell(row=r, column=21).alignment = Alignment(horizontal="center")
    # koşullu biçim
    if sonr >= ilk:
        _devam_olcegi(ws, f"F{ilk}:F{sonr}")
        ws.conditional_formatting.add(f"L{ilk}:L{sonr}", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                                      color=ORANGE, showValue=True))
        ws.conditional_formatting.add(f"H{ilk}:H{sonr}", ColorScaleRule(start_type="min", start_color="FFFFFF",
                                                                         end_type="max", end_color="F4B6AE"))
        ws.conditional_formatting.add(f"S{ilk}:S{sonr}", CellIsRule(operator="greaterThan", formula=["0.05"],
                                                                     font=Font(name=FONT, color=RED, bold=True)))
        ws.conditional_formatting.add(f"S{ilk}:S{sonr}", CellIsRule(operator="lessThan", formula=["-0.05"],
                                                                     font=Font(name=FONT, color=GREEN, bold=True)))
        _puan_kurallari(ws, f"T{ilk}:T{sonr}", f"U{ilk}:U{sonr}")
    return ilk, sonr, r


def _kars_sayfasi(wb, d, baslik, sekme, satirlar, ad_basligi, renk):
    ws = wb.create_sheet(sekme)
    ws.sheet_properties.tabColor = renk
    son = len(KARS_BASLIK)
    _genislik(ws, [6, 28, 7, 10, 10, 9, 8, 8, 10, 9, 10, 8, 8, 10, 9, 10, 10, 10, 10, 8, 6])
    _sayfa_basligi(ws, son, baslik, d)
    r = _bolum(ws, 5, 1, son, f"{ad_basligi} bazında göstergeler",
               "Sıra verimlilik puanına göredir. Oranlar ve puan hücre formülüdür. Yevmiye, hafta tatili / pazar "
               "kuralına göre hesaplanmış değerdir (Yöntem sayfası).")
    ilk, sonr, top_r = _kars_tablosu(ws, r, satirlar, ad_basligi, d["onceki"] is not None)
    ws.freeze_panes = ws.cell(row=r + 1, column=3)
    ws.auto_filter.ref = f"A{r}:{L(son)}{sonr}"
    # puan grafiği (yatay çubuk)
    if sonr >= ilk:
        ch = BarChart()
        ch.type = "bar"
        ch.style = 10
        ch.title = f"{ad_basligi} bazında devam oranı ve kapasite"
        ch.add_data(Reference(ws, min_col=6, min_row=r, max_row=sonr), titles_from_data=True)
        ch.add_data(Reference(ws, min_col=14, min_row=r, max_row=sonr), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=2, min_row=ilk, max_row=sonr))
        ch.series[0].graphicalProperties.solidFill = NAVY
        ch.series[1].graphicalProperties.solidFill = "8FAACB"
        ch.y_axis.number_format = "0%"
        ch.y_axis.scaling.min = 0.5
        ch.y_axis.scaling.max = 1.05
        ch.y_axis.majorGridlines = None
        ch.x_axis.scaling.orientation = "maxMin"
        ch.legend.position = "b"
        ch.height = max(6.5, 0.75 * (sonr - ilk + 1) + 3)
        ch.width = 22
        ws.add_chart(ch, f"B{top_r + 3}")
    _yazdir(ws, basliklar=f"{r}:{r}")
    return ws


# ================================================================ 2) TREND + 3) GÜNLÜK

def _trend(wb, d):
    ws = wb.create_sheet("Trend")
    ws.sheet_properties.tabColor = NAVY
    basliklar = ["Periyot", "Başlangıç", "Bitiş", "Ort. çalışan", "Planlanan gün", "Çalışılan gün", "Devam %",
                 "Kayıp gün", "Normal sa", "Fazla sa", "Toplam sa", "FM %", "Sa / gün", "Kapasite %", "Pazar ×2,5 sa",
                 "Yevmiye"]
    son = len(basliklar)
    _genislik(ws, [15, 11, 11, 10, 10, 10, 9, 9, 11, 10, 11, 8, 8, 10, 10, 10])
    periyot_adi = {"gun": "GÜNLÜK", "hafta": "HAFTALIK", "ay": "AYLIK"}[d["periyot"]]
    _sayfa_basligi(ws, son, f"{periyot_adi} TREND", d)
    r = _bolum(ws, 5, 1, son, f"{periyot_adi.capitalize()} kırılım", "Oranlar ve toplamlar formüldür.")
    _baslik_satiri(ws, r, 1, basliklar)
    hr = r
    r += 1
    ilk = r
    for i, p in enumerate(d["periyotlar"]):
        bas = max(p["anahtar"], d["bas"])
        bit = min({"gun": p["anahtar"], "hafta": p["anahtar"] + datetime.timedelta(days=6)}.get(
            d["periyot"], (datetime.date(p["anahtar"].year + (p["anahtar"].month == 12), p["anahtar"].month % 12 + 1, 1)
                           - datetime.timedelta(days=1))), d["bit"])
        vals = [p["etiket"], bas, bit, round(p["ort_calisan"], 1), p["planlanan"], p["calisilan"], None, p["kayip"],
                round(p["normal"], 2), round(p["fazla"], 2), None, None, None, None, round(p["pazar_katli"], 2),
                round(p["yevmiye"], 2)]
        for c, v in enumerate(vals, start=1):
            _yaz(ws, r, c, v)
        ws[f"G{r}"] = f'=IF(E{r}=0,"",F{r}/E{r})'
        ws[f"K{r}"] = f"=I{r}+J{r}"
        ws[f"L{r}"] = f'=IF(K{r}=0,"",J{r}/K{r})'
        ws[f"M{r}"] = f'=IF(F{r}=0,"",K{r}/F{r})'
        ws[f"N{r}"] = f'=IF(E{r}=0,"",I{r}/(E{r}*{vm.STANDART_GUN}))'
        _veri_satiri(ws, r, 1, son, zebra=(i % 2 == 1))
        r += 1
    sonr = r - 1
    _yaz(ws, r, 1, "TOPLAM")
    ws[f"D{r}"] = round(d["genel"]["ort_calisan"], 1)
    for c in (5, 6, 8, 9, 10, 15, 16):
        ws.cell(row=r, column=c, value=f"=SUM({L(c)}{ilk}:{L(c)}{sonr})")
    ws[f"G{r}"] = f'=IF(E{r}=0,"",F{r}/E{r})'
    ws[f"K{r}"] = f"=I{r}+J{r}"
    ws[f"L{r}"] = f'=IF(K{r}=0,"",J{r}/K{r})'
    ws[f"M{r}"] = f'=IF(F{r}=0,"",K{r}/F{r})'
    ws[f"N{r}"] = f'=IF(E{r}=0,"",I{r}/(E{r}*{vm.STANDART_GUN}))'
    _toplam_satiri(ws, r, 1, son)
    fmt = {2: F_TARIH, 3: F_TARIH, 4: F_1, 5: F_INT, 6: F_INT, 7: F_PCT, 8: F_INT, 9: F_INT, 10: F_INT, 11: F_INT,
           12: F_PCT, 13: F_2, 14: F_PCT, 15: F_INT, 16: F_1}
    for rr in range(ilk, r + 1):
        for c, f in fmt.items():
            ws.cell(row=rr, column=c).number_format = f
            ws.cell(row=rr, column=c).alignment = Alignment(horizontal="right" if c > 3 else "center")
        if rr < r:
            ws.cell(row=rr, column=1).font = _font(10, True)
    if sonr >= ilk:
        _devam_olcegi(ws, f"G{ilk}:G{sonr}")
        ws.conditional_formatting.add(f"L{ilk}:L{sonr}", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                                      color=ORANGE, showValue=True))
        # grafik: yığılmış normal + fazla, ikincil eksende devam oranı
        bar = BarChart()
        bar.type, bar.grouping, bar.overlap = "col", "stacked", 100
        bar.title = "Çalışma saatleri ve devam oranı"
        bar.add_data(Reference(ws, min_col=9, max_col=10, min_row=hr, max_row=sonr), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=ilk, max_row=sonr))
        bar.series[0].graphicalProperties.solidFill = NAVY
        bar.series[1].graphicalProperties.solidFill = ORANGE
        bar.gapWidth = 60
        bar.y_axis.title = "saat"
        bar.y_axis.number_format = "#,##0"
        bar.y_axis.majorGridlines.spPr = None
        line = LineChart()
        line.add_data(Reference(ws, min_col=7, min_row=hr, max_row=sonr), titles_from_data=True)
        line.series[0].graphicalProperties.line.solidFill = GREEN
        line.series[0].graphicalProperties.line.width = 28000
        line.series[0].marker.symbol = "circle"
        line.series[0].marker.size = 6
        line.series[0].marker.graphicalProperties.solidFill = GREEN
        line.series[0].marker.graphicalProperties.line.solidFill = GREEN
        line.y_axis.axId = 200
        line.y_axis.title = "devam"
        line.y_axis.number_format = "0%"
        line.y_axis.scaling.min = 0.7
        line.y_axis.scaling.max = 1.0
        line.y_axis.majorGridlines = None
        line.y_axis.crosses = "max"
        bar += line
        bar.legend.position = "b"
        bar.height, bar.width = 9.5, 26
        ws.add_chart(bar, f"A{r + 3}")
    ws.freeze_panes = ws.cell(row=hr + 1, column=2)
    _yazdir(ws, basliklar=f"{hr}:{hr}")


def _gunluk(wb, d):
    ws = wb.create_sheet("Günlük")
    ws.sheet_properties.tabColor = NAVY
    katlar = ["calisti", "ucretli_izin", "ucretsiz_izin", "rapor", "devamsiz", "hafta_tatili", "resmi_tatil", "belirsiz"]
    basliklar = (["Tarih", "Gün"] + [vm.KAT_ETIKET[k] for k in katlar] +
                 ["Planlanan", "Devam %", "Normal sa", "Fazla sa", "Toplam sa", f"{vm.etiket_gunluk()} sa üstü"])
    son = len(basliklar)
    _genislik(ws, [11, 6] + [9.5] * 8 + [10, 9, 10, 9, 10, 9])
    _sayfa_basligi(ws, son, "GÜNLÜK İŞ GÜCÜ", d)
    r = _bolum(ws, 5, 1, son, "Günlük dağılım", "Kişi sayıları; pazar günleri turuncu zeminli. Devam % = çalışan ÷ "
                                                  "planlanan.")
    _baslik_satiri(ws, r, 1, basliklar)
    hr = r
    r += 1
    ilk = r
    for g in d["gunluk"]:
        pz = g["tarih"].weekday() == 6
        _yaz(ws, r, 1, g["tarih"], F_TARIH, align=Alignment(horizontal="center"))
        _yaz(ws, r, 2, g["gun_adi"], align=Alignment(horizontal="center"), font=_font(10, pz, ORANGE_T if pz else INK))
        for j, k in enumerate(katlar):
            _yaz(ws, r, 3 + j, g["gun"][k] or None, F_INT)
        c_pl = 3 + len(katlar)
        ws.cell(row=r, column=c_pl, value=f"=SUM(C{r}:G{r})").number_format = F_INT
        ws.cell(row=r, column=c_pl + 1, value=f'=IF({L(c_pl)}{r}=0,"",C{r}/{L(c_pl)}{r})').number_format = F_PCT
        _yaz(ws, r, c_pl + 2, round(g["normal"], 2), F_INT)
        _yaz(ws, r, c_pl + 3, round(g["fazla"], 2), F_INT)
        ws.cell(row=r, column=c_pl + 4, value=f"={L(c_pl + 2)}{r}+{L(c_pl + 3)}{r}").number_format = F_INT
        _yaz(ws, r, c_pl + 5, g["uzun_gun"] or None, F_INT)
        _veri_satiri(ws, r, 1, son, fill=_fill(ORANGE_BG) if pz else None)
        r += 1
    sonr = r - 1
    _yaz(ws, r, 1, "TOPLAM")
    for c in range(3, son + 1):
        if c == 3 + len(katlar) + 1:
            ws.cell(row=r, column=c, value=f'=IF({L(c - 1)}{r}=0,"",C{r}/{L(c - 1)}{r})').number_format = F_PCT
        else:
            ws.cell(row=r, column=c, value=f"=SUM({L(c)}{ilk}:{L(c)}{sonr})").number_format = F_INT
    _toplam_satiri(ws, r, 1, son)
    if sonr >= ilk:
        _devam_olcegi(ws, f"{L(4 + len(katlar))}{ilk}:{L(4 + len(katlar))}{sonr}")
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
        ch.title = "Günlük iş gücü (kişi)"
        ch.add_data(Reference(ws, min_col=3, max_col=7, min_row=hr, max_row=sonr), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=ilk, max_row=sonr))
        for s, k in zip(ch.series, katlar[:5]):
            s.graphicalProperties.solidFill = vm.KAT_RENK[k]
            s.graphicalProperties.line.noFill = True
        ch.gapWidth = 20
        ch.x_axis.number_format = "dd.mm"
        ch.y_axis.majorGridlines.spPr = None
        ch.legend.position = "b"
        ch.height, ch.width = 9, 28
        ws.add_chart(ch, f"A{r + 3}")
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)
    _yazdir(ws, basliklar=f"{hr}:{hr}")


# ================================================================ 6) DEVAM HARİTASI

def _isi(wb, d):
    ws = wb.create_sheet("Devam Haritası")
    ws.sheet_properties.tabColor = GREEN
    tarihler = d["isi_tarihler"]
    son = 2 + len(tarihler)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 8
    for i in range(len(tarihler)):
        ws.column_dimensions[L(3 + i)].width = 4.2
    _sayfa_basligi(ws, max(son, 12), "DEVAM HARİTASI", d)
    etiket = "Alt ekip" if d["kirilim"] == "ekip" else "Grup › alt ekip"
    ws.column_dimensions["A"].width = 34
    r = _bolum(ws, 5, 1, max(son, 12), "Alt ekip × gün devam oranı",
               "Hücre = o gün çalışan ÷ planlanan. Boş = kayıt yok, gri = tatil / plan yok. Yeşil iyi, kırmızı düşük.")
    _yaz(ws, r, 1, etiket, font=_font(9, True, "FFFFFF"), fill=_fill(NAVY), align=Alignment(vertical="center", indent=1))
    _yaz(ws, r, 2, "Dönem", font=_font(9, True, "FFFFFF"), fill=_fill(NAVY), align=Alignment(horizontal="center",
                                                                                             vertical="center"))
    for i, t in enumerate(tarihler):
        pz = t.weekday() == 6
        _yaz(ws, r, 3 + i, f"{t:%d}\n{vm.GUN_KISA[t.weekday()][:2]}", font=_font(8, True, "FFFFFF"),
             fill=_fill(ORANGE if pz else NAVY), align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[r].height = 28
    hr = r
    r += 1
    ilk = r
    for s in d["isi"]:
        _yaz(ws, r, 1, s["ad"], font=_font(9.5, True), border=ALT_CIZGI)
        _yaz(ws, r, 2, round(s["devam"] / 100, 4), F_PCT, _font(9.5, True), border=ALT_CIZGI,
             align=Alignment(horizontal="center"))
        for i, h in enumerate(s["hucreler"]):
            c = ws.cell(row=r, column=3 + i)
            if h is None:
                continue
            if h["devam"] is None:
                c.fill = _fill("ECE9E3")
            else:
                c.value = round(h["devam"] / 100, 4)
                c.number_format = '0%'
            c.font = _font(7, False, INK2)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"),
                              bottom=Side(style="thin", color="FFFFFF"))
        ws.row_dimensions[r].height = 18
        r += 1
    sonr = r - 1
    if sonr >= ilk and tarihler:
        rule = ColorScaleRule(start_type="num", start_value=0.8, start_color="D9534F", mid_type="num", mid_value=0.93,
                              mid_color="F7E3A1", end_type="num", end_value=1, end_color="3C8D4A")
        ws.conditional_formatting.add(f"C{ilk}:{L(son)}{sonr}", rule)
        _devam_olcegi(ws, f"B{ilk}:B{sonr}")
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)
    _yazdir(ws, basliklar=f"{hr}:{hr}")


# ================================================================ 7) PERSONEL KARNESİ

def _karne(wb, d, ucret):
    # Sütun başlıkları eşiklere göre (Ayarlar) üretilir; aşağıda sözlük anahtarı olarak da kullanılır.
    KALAN = f"{vm.etiket_yillik()} sa sınırına kalan"
    UZUN = f"{vm.etiket_gunluk()} sa üstü gün"
    ws = wb.create_sheet("Personel Karnesi")
    ws.sheet_properties.tabColor = ORANGE
    basliklar = ["Sıra", "Ad Soyad", "TC Kimlik No", "Grup", "Alt ekip", "Görevi", "Planlanan gün", "Çalışılan gün",
                 "Devam %", "Ücretli izin", "Ücretsiz izin", "Rapor", "Devamsız", "Kayıp %", "Normal sa", "Fazla sa",
                 "Pazar ×2,5 sa", "FM %", "Sa / gün", "Kapasite %", "Yevmiye"]
    if ucret:
        basliklar.append("Yevmiye tutarı (TL)")
    basliklar += [f"{d['yil']} FM sa", KALAN, "En uzun seri (gün)", UZUN, "Puan", "Not", "Uyarılar"]
    son = len(basliklar)
    k = {h: i + 1 for i, h in enumerate(basliklar)}
    C = {h: L(i) for h, i in k.items()}
    _genislik(ws, [6, 26, 13, 14, 18, 16] + [9] * (son - 7) + [40])
    _sayfa_basligi(ws, son, "PERSONEL KARNESİ", d, alt_ek=f"{len(d['kisiler'])} kişi")
    r = _bolum(ws, 5, 1, son, "Kişi bazında bütün göstergeler",
               "Sütun başlıklarındaki oklarla süzebilir / sıralayabilirsiniz. Oranlar ve puan formüldür; yevmiye hafta "
               "tatili / pazar kuralına göre hesaplanmış değerdir.")
    _baslik_satiri(ws, r, 1, basliklar, 40)
    hr = r
    r += 1
    ilk = r
    for i, p in enumerate(d["kisiler"], start=1):
        v = {"Sıra": i, "Ad Soyad": p["ad_soyad"], "TC Kimlik No": p["tc"], "Grup": p["grup"], "Alt ekip": p["ekip"],
             "Görevi": p["gorev"], "Planlanan gün": p["planlanan"], "Çalışılan gün": p["calisilan"],
             "Ücretli izin": p["gun"]["ucretli_izin"] or None, "Ücretsiz izin": p["gun"]["ucretsiz_izin"] or None,
             "Rapor": p["gun"]["rapor"] or None, "Devamsız": p["gun"]["devamsiz"] or None,
             "Normal sa": round(p["normal"], 2), "Fazla sa": round(p["fazla"], 2), "Pazar ×2,5 sa": round(p["pazar_katli"], 2),
             "Yevmiye": round(p["yevmiye"], 2),
             f"{d['yil']} FM sa": round(p["yillik_fm"], 2), "En uzun seri (gün)": p["en_uzun_seri"],
             UZUN: p["uzun_gun"] or None, "Uyarılar": " · ".join(p["riskler"])}
        for h, val in v.items():
            _yaz(ws, r, k[h], val)
        pl, cl = C["Planlanan gün"], C["Çalışılan gün"]
        ws[f"{C['Devam %']}{r}"] = f'=IF({pl}{r}=0,"",{cl}{r}/{pl}{r})'
        ws[f"{C['Kayıp %']}{r}"] = (f'=IF({pl}{r}=0,"",({C["Ücretsiz izin"]}{r}+{C["Rapor"]}{r}+{C["Devamsız"]}{r})'
                                    f'/{pl}{r})')
        n, f = C["Normal sa"], C["Fazla sa"]
        ws[f"{C['FM %']}{r}"] = f'=IF({n}{r}+{f}{r}=0,"",{f}{r}/({n}{r}+{f}{r}))'
        ws[f"{C['Sa / gün']}{r}"] = f'=IF({cl}{r}=0,"",({n}{r}+{f}{r})/{cl}{r})'
        ws[f"{C['Kapasite %']}{r}"] = f'=IF({pl}{r}=0,"",{n}{r}/({pl}{r}*{vm.STANDART_GUN}))'
        if ucret:
            ws[f"{C['Yevmiye tutarı (TL)']}{r}"] = f"={C['Yevmiye']}{r}*{ucret}"
        yfm = C[f"{d['yil']} FM sa"]
        ws[f"{C[KALAN]}{r}"] = f"=MAX(0,{vm.YILLIK_FM_SINIRI}-{yfm}{r})"
        ws[f"{C['Puan']}{r}"] = (f'=IF({pl}{r}=0,"",(0.6*{C["Devam %"]}{r}+0.4*MIN({C["Kapasite %"]}{r},1))*100)')
        pu = C["Puan"]
        ws[f"{C['Not']}{r}"] = f'=IF({pu}{r}="","",IF({pu}{r}>=95,"A",IF({pu}{r}>=85,"B",IF({pu}{r}>=70,"C","D"))))'
        _veri_satiri(ws, r, 1, son, zebra=(i % 2 == 0))
        r += 1
    sonr = r - 1
    fmt = {"Sıra": F_INT, "Planlanan gün": F_INT, "Çalışılan gün": F_INT, "Devam %": F_PCT, "Ücretli izin": F_INT,
           "Ücretsiz izin": F_INT, "Rapor": F_INT, "Devamsız": F_INT, "Kayıp %": F_PCT, "Normal sa": F_INT,
           "Fazla sa": F_1, "Pazar ×2,5 sa": F_1, "FM %": F_PCT, "Sa / gün": F_2, "Kapasite %": F_PCT, "Yevmiye": F_1,
           "Yevmiye tutarı (TL)": F_TL, f"{d['yil']} FM sa": F_1, KALAN: F_1, "En uzun seri (gün)": F_INT,
           UZUN: F_INT, "Puan": F_1}
    for rr in range(ilk, sonr + 1):
        for h, f in fmt.items():
            if h in k:
                cell = ws.cell(row=rr, column=k[h])
                cell.number_format = f
                cell.alignment = Alignment(horizontal="right")
        ws.cell(row=rr, column=k["Ad Soyad"]).font = _font(10, True)
        ws.cell(row=rr, column=k["Not"]).alignment = Alignment(horizontal="center")
        ws.cell(row=rr, column=k["Uyarılar"]).font = _font(9, False, RED)
        ws.cell(row=rr, column=k["TC Kimlik No"]).font = _font(10, False, INK2)
    # toplam
    _yaz(ws, r, 2, f"TOPLAM ({len(d['kisiler'])} kişi)")
    for h in ("Planlanan gün", "Çalışılan gün", "Ücretli izin", "Ücretsiz izin", "Rapor", "Devamsız", "Normal sa", "Fazla sa",
              "Pazar ×2,5 sa", "Yevmiye", "Yevmiye tutarı (TL)", UZUN):
        if h in k:
            ws.cell(row=r, column=k[h], value=f"=SUBTOTAL(9,{C[h]}{ilk}:{C[h]}{sonr})").number_format = fmt[h]
    pl, cl = C["Planlanan gün"], C["Çalışılan gün"]
    ws[f"{C['Devam %']}{r}"] = f'=IF({pl}{r}=0,"",{cl}{r}/{pl}{r})'
    ws[f"{C['Kapasite %']}{r}"] = f'=IF({pl}{r}=0,"",{C["Normal sa"]}{r}/({pl}{r}*{vm.STANDART_GUN}))'
    ws[f"{C['FM %']}{r}"] = f'=IF({C["Normal sa"]}{r}+{C["Fazla sa"]}{r}=0,"",{C["Fazla sa"]}{r}/({C["Normal sa"]}{r}+{C["Fazla sa"]}{r}))'
    for h in ("Devam %", "Kapasite %", "FM %"):
        ws[f"{C[h]}{r}"].number_format = F_PCT
    _toplam_satiri(ws, r, 1, son)
    for c in range(3, son + 1):
        ws.cell(row=r, column=c).alignment = Alignment(horizontal="right")
    _yaz(ws, r + 1, 2, "Toplam satırı süzülen (görünen) satırları toplar.", font=_font(8.5, False, MUTED, True))
    if sonr >= ilk:
        _devam_olcegi(ws, f"{C['Devam %']}{ilk}:{C['Devam %']}{sonr}")
        _devam_olcegi(ws, f"{C['Kapasite %']}{ilk}:{C['Kapasite %']}{sonr}")
        ws.conditional_formatting.add(f"{C['Fazla sa']}{ilk}:{C['Fazla sa']}{sonr}",
                                      DataBarRule(start_type="num", start_value=0, end_type="max", color=ORANGE))
        ws.conditional_formatting.add(f"{C['Devamsız']}{ilk}:{C['Devamsız']}{sonr}",
                                      CellIsRule(operator="greaterThanOrEqual", formula=[str(vm.DEVAMSIZLIK_ESIGI)],
                                                 fill=_fill(RED_BG), font=Font(name=FONT, bold=True, color=RED)))
        ws.conditional_formatting.add(f"{C[str(d['yil']) + ' FM sa']}{ilk}:{C[str(d['yil']) + ' FM sa']}{sonr}",
                                      CellIsRule(operator="greaterThanOrEqual", formula=[str(vm.YILLIK_FM_SINIRI * 0.85)],
                                                 fill=_fill(RED_BG), font=Font(name=FONT, bold=True, color=RED)))
        ws.conditional_formatting.add(f"{C['En uzun seri (gün)']}{ilk}:{C['En uzun seri (gün)']}{sonr}",
                                      CellIsRule(operator="greaterThanOrEqual", formula=[str(vm.KESINTISIZ_GUN_ESIGI)],
                                                 fill=_fill(ORANGE_BG), font=Font(name=FONT, bold=True, color=ORANGE_T)))
        _puan_kurallari(ws, f"{C['Puan']}{ilk}:{C['Puan']}{sonr}", f"{C['Not']}{ilk}:{C['Not']}{sonr}")
    ws.auto_filter.ref = f"A{hr}:{L(son)}{sonr}"
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)
    _yazdir(ws, basliklar=f"{hr}:{hr}")
    ws.page_setup.paperSize = ws.PAPERSIZE_A3


# ================================================================ 8) DEVAM TAKVİMİ

def _takvim(wb, d, db_path=None):
    ws = wb.create_sheet("Devam Takvimi")
    ws.sheet_properties.tabColor = PURPLE
    tarihler = [d["bas"] + datetime.timedelta(days=i) for i in range(d["gun_sayisi"])]
    son = 3 + len(tarihler) + 3
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    for i in range(len(tarihler)):
        ws.column_dimensions[L(4 + i)].width = 4.3
    for j, w in enumerate((8, 9, 8)):
        ws.column_dimensions[L(4 + len(tarihler) + j)].width = w
    _sayfa_basligi(ws, min(son, 40), "DEVAM TAKVİMİ", d)
    # lejant: renkli örnek hücreler + açıklama
    r = 5
    lej = [("9", NAVY_BG, "çalışılan saat"), ("11", "FBE3CF", "fazla mesaili gün"), ("8", "F8C9A3", "pazar ×2,5"),
           ("13", "FFFFFF", f"{vm.etiket_gunluk()} saat üstü (kırmızı)"), ("İ", "E7DDF1", "ücretli izin"), ("Ü", "E4E7EA", "ücretsiz izin"),
           ("R", "FCE3CF", "rapor"), ("D", RED_BG, "devamsızlık"), ("HT", "F1EEE8", "hafta tatili"),
           ("RT", "DDE6F0", "resmî tatil")]
    _yaz(ws, r, 1, "Hücre içeriği:", font=_font(9, True, INK2))
    parcalar = []
    for i, (kod, zemin, ad) in enumerate(lej):
        c = ws.cell(row=r, column=4 + i * 3)
        _yaz(ws, r, 4 + i * 3, kod, font=_font(8, True, RED if kod == "13" else (ORANGE_T if zemin in ("FBE3CF", "F8C9A3") else INK)),
             fill=_fill(zemin), align=Alignment(horizontal="center"), border=KUTU)
        ws.merge_cells(start_row=r, start_column=5 + i * 3, end_row=r, end_column=6 + i * 3)
        _yaz(ws, r, 5 + i * 3, ad, font=_font(7.5, False, MUTED), align=Alignment(vertical="center", indent=0))
    r += 2
    basl = ["Ad Soyad", "Görevi", "Alt ekip"]
    for i, h in enumerate(basl):
        _yaz(ws, r, 1 + i, h, font=_font(9, True, "FFFFFF"), fill=_fill(NAVY), align=Alignment(vertical="center"))
    for i, t in enumerate(tarihler):
        pz = t.weekday() == 6
        _yaz(ws, r, 4 + i, f"{t:%d}\n{vm.GUN_KISA[t.weekday()][:2]}", font=_font(8, True, "FFFFFF"),
             fill=_fill(ORANGE if pz else NAVY), align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    for j, h in enumerate(("Çalışılan", "Toplam sa", "Devam")):
        _yaz(ws, r, 4 + len(tarihler) + j, h, font=_font(8.5, True, "FFFFFF"), fill=_fill(NAVY_D),
             align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[r].height = 28
    hr = r
    r += 1
    # kişi-gün verisi
    import veri_merkezi_sorgu as sorgu
    rows = sorgu.get_daily_hours_filtered(d["gruplar"] or None, d["ekipler"] or None, d["gorevler"] or None, None,
                                          d["bas"].isoformat(), d["bit"].isoformat(), db_path=db_path)
    gunler = {}
    for x in rows:
        gunler[(x["tc"], x["tarih"])] = x
    kural = d.get("gun_kurali") or {}
    idx = {t: i for i, t in enumerate(tarihler)}
    zemin_kod = {"ucretli_izin": "E7DDF1", "ucretsiz_izin": "E4E7EA", "rapor": "FCE3CF", "devamsiz": RED_BG,
                 "hafta_tatili": "F1EEE8", "resmi_tatil": "DDE6F0", "belirsiz": "FFFFFF"}
    kisiler = sorted(d["kisiler"], key=lambda p: (p["kirilim"], p["ekip"], p["gorev"], p["ad_soyad"]))
    onceki_kir = None
    for p in kisiler:
        if p["kirilim"] != onceki_kir:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            _yaz(ws, r, 1, p["kirilim"], font=_font(9.5, True, NAVY_D), fill=_fill(NAVY_BG))
            for cc in range(1, son + 1):
                ws.cell(row=r, column=cc).fill = _fill(NAVY_BG)
            onceki_kir = p["kirilim"]
            r += 1
        _yaz(ws, r, 1, p["ad_soyad"], font=_font(9, True), border=ALT_CIZGI)
        _yaz(ws, r, 2, p["gorev"], font=_font(8.5, False, MUTED), border=ALT_CIZGI)
        _yaz(ws, r, 3, p["ekip"], font=_font(8.5, False, MUTED), border=ALT_CIZGI)
        for t in tarihler:
            x = gunler.get((p["tc"], t.isoformat()))
            c = ws.cell(row=r, column=4 + idx[t])
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"),
                              bottom=Side(style="thin", color=LINE_SOFT))
            if x is None:
                continue
            n, f = float(x["normal_saat"] or 0), float(x["fazla_saat"] or 0)
            # Hafta tatili / pazar kuralına göre düzeltilmiş kategori (ör. cuma ücretsiz izin -> HT) ve pazar ×2,5
            kat, katli = kural.get((p["tc"], x["tarih"]), (vm.kategori(x["durum"], n + f), False))
            if kat == "calisti":
                c.value = round(n + f, 1) if (n + f) % 1 else int(n + f)
                c.font = _font(8, f > 0, ORANGE_T if f > 0 else INK)
                # pazar zemini yalnız ×2,5 ödenen (hak edilmiş) pazarlarda; hak edilmemiş pazar sıradan gün gibi
                c.fill = _fill("F8C9A3" if katli else ("FBE3CF" if f > 0 else NAVY_BG))
                if n + f > vm.GUNLUK_AZAMI_SAAT:
                    c.font = _font(8, True, RED)
            else:
                c.value = vm.KAT_KOD[kat] or None
                c.font = _font(7.5, True, RED if kat == "devamsiz" else INK2)
                c.fill = _fill(zemin_kod[kat])
        c0 = 4 + len(tarihler)
        _yaz(ws, r, c0, p["calisilan"], F_INT, _font(9, True), align=Alignment(horizontal="right"))
        _yaz(ws, r, c0 + 1, round(p["toplam"], 1), F_INT, _font(9), align=Alignment(horizontal="right"))
        _yaz(ws, r, c0 + 2, round(p["devam"] / 100, 4), F_PCT, _font(9, True), align=Alignment(horizontal="right"))
        ws.row_dimensions[r].height = 15
        r += 1
    _devam_olcegi(ws, f"{L(4 + len(tarihler) + 2)}{hr + 1}:{L(4 + len(tarihler) + 2)}{r - 1}")
    ws.freeze_panes = ws.cell(row=hr + 1, column=4)
    _yazdir(ws, basliklar=f"{hr}:{hr}")
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    # uzun dönemde tek sayfa genişliğine sıkıştırma okunmaz: her ~35 gün bir sayfa, ilk 3 sütun her sayfada
    ws.page_setup.fitToWidth = max(1, -(-len(tarihler) // 35))
    ws.print_title_cols = "A:C"


# ================================================================ 9) UYUM VE RİSK

def _risk(wb, d):
    ws = wb.create_sheet("Uyum ve Risk")
    ws.sheet_properties.tabColor = RED
    SON = 8
    _genislik(ws, [6, 26, 14, 18, 18, 14, 14, 30])
    _sayfa_basligi(ws, SON, "UYUM VE RİSK", d)
    rk = d["risk"]
    r = _bolum(ws, 5, 1, SON, "Özet", "4857 sayılı İş Kanunu sınırları ve Veri Merkezi ayarlarındaki fazla mesai eşiğine göre")
    _baslik_satiri(ws, r, 1, ["#", "Kontrol", "Kayıt", "Dayanak", "", "", "", "Açıklama"], 22)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    r += 1
    kontroller = [
        (f"Yıllık {vm.etiket_yillik()} sa FM sınırını aşan", len(rk["yillik_asan"]), "md. 41", f"{d['yil']} başından dönem sonuna"),
        ("Yıllık sınıra yaklaşan (≥ %85)", len(rk["yillik_yakin"]), "md. 41", f"{vm.etiket_yakin()} saat ve üzeri"),
        ("Dönem FM eşiği üstü", len(rk["esik_ustu"]), "Ayarlar",
         f"aylık {report.format_tr(d['esik_aylik'], 0)} sa → dönem {report.format_tr(d['esik_donem'], 1)} sa"),
        (f"Günlük {vm.etiket_gunluk()} saat aşımı (kayıt)", len(rk["uzun_gunler"]), "md. 63",
         f"normal + fazla > {vm.etiket_gunluk()} sa"),
        (f"{vm.KESINTISIZ_GUN_ESIGI}+ gün kesintisiz çalışan", len(rk["kesintisiz"]), "md. 46",
         "7 günde en az 24 sa dinlenme"),
        (f"{vm.DEVAMSIZLIK_ESIGI}+ gün devamsızlık", len(rk["devamsiz"]), "iç kural", "mazeretsiz devamsızlık"),
    ]
    for i, (a, n, dayanak, acik) in enumerate(kontroller, start=1):
        _yaz(ws, r, 1, i, align=Alignment(horizontal="center"))
        _yaz(ws, r, 2, a, font=_font(10, True))
        _yaz(ws, r, 3, n, F_INT, _font(11, True, RED if n else GREEN), align=Alignment(horizontal="center"))
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        _yaz(ws, r, 4, dayanak, font=_font(10, False, INK2))
        _yaz(ws, r, 8, acik, font=_font(9, False, MUTED, True))
        _veri_satiri(ws, r, 1, SON, zebra=(i % 2 == 0))
        ws.cell(row=r, column=3).fill = _fill(RED_BG if n else GREEN_BG)
        r += 1

    def liste(r, baslik, alt, basliklar, satirlar, fmts):
        r = _bolum(ws, r + 1, 1, SON, f"{baslik} ({len(satirlar)})", alt)
        _baslik_satiri(ws, r, 1, basliklar, 26)
        r += 1
        if not satirlar:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=SON)
            _yaz(ws, r, 1, "Bu dönemde kayıt yok.", font=_font(10, False, GREEN, True), fill=_fill(GREEN_BG))
            return r + 1
        for i, vals in enumerate(satirlar, start=1):
            for c, v in enumerate([i] + vals, start=1):
                cell = _yaz(ws, r, c, v, fmts.get(c))
                if c > 2 and isinstance(v, (int, float)):
                    cell.alignment = Alignment(horizontal="right")
            ws.cell(row=r, column=2).font = _font(10, True)
            _veri_satiri(ws, r, 1, SON, zebra=(i % 2 == 0))
            r += 1
        return r

    yil = rk["yillik_asan"] + rk["yillik_yakin"]
    r = liste(r, "Yıllık fazla mesai sınırı", "4857 s. İş Kanunu md. 41: yılda en fazla 270 saat"
              + vm.esik_notu(vm.YILLIK_FM_SINIRI, 270),
              ["#", "Ad Soyad", "TC", "Alt ekip", "Görevi", "Dönem FM sa", f"{d['yil']} FM sa", "Kalan / aşım (sa)"],
              [[k["ad_soyad"], k["tc"], k["ekip"], k["gorev"], round(k["fazla"], 1), round(k["yillik_fm"], 1),
                round(vm.YILLIK_FM_SINIRI - k["yillik_fm"], 1)] for k in yil], {6: F_1, 7: F_1, 8: '+#,##0.0;-#,##0.0'})
    r = liste(r, "Dönem fazla mesai eşiğini aşanlar",
              f"Aylık {report.format_tr(d['esik_aylik'], 0)} sa eşiği dönem uzunluğuna oranlanır: "
              f"{report.format_tr(d['esik_donem'], 1)} sa",
              ["#", "Ad Soyad", "TC", "Alt ekip", "Görevi", "Dönem FM sa", "Pazar ×2,5 sa", "Eşik farkı (sa)"],
              [[k["ad_soyad"], k["tc"], k["ekip"], k["gorev"], round(k["fazla"], 1), round(k["pazar_katli"], 1),
                round(k["fazla"] - d["esik_donem"], 1)] for k in rk["esik_ustu"]], {6: F_1, 7: F_1, 8: '+#,##0.0'})
    r = liste(r, f"Günlük {vm.etiket_gunluk()} saati aşan çalışma — kişi bazında",
              f"4857 s. İş Kanunu md. 63{vm.esik_notu(vm.GUNLUK_AZAMI_SAAT, 11)} · toplam {len(rk['uzun_gunler'])} kişi-gün; günlerin tamamı Devam Takvimi'nde "
              f"kırmızı yazılıdır",
              ["#", "Ad Soyad", "TC", "Alt ekip", "Görevi", "Gün sayısı", "En yüksek (sa)", "Tarihler"],
              [[u["ad_soyad"], u["tc"], u["ekip"], u["gorev"], u["gun"], u["en_yuksek"],
                ", ".join(f"{t:%d.%m}" for t in u["tarihler"][:8]) + (" …" if len(u["tarihler"]) > 8 else "")]
               for u in rk["uzun_kisi"]], {6: F_INT, 7: F_1})
    seriler = []
    for k in rk["kesintisiz"]:
        for a, b, n in k["seriler"]:
            seriler.append([k["ad_soyad"], k["tc"], k["ekip"], a, b, n, k["gorev"]])
    r = liste(r, f"{vm.KESINTISIZ_GUN_ESIGI} gün ve üzeri kesintisiz çalışma", "4857 s. İş Kanunu md. 46 (hafta tatili)",
              ["#", "Ad Soyad", "TC", "Alt ekip", "Başlangıç", "Bitiş", "Gün", "Görevi"], seriler,
              {5: F_TARIH, 6: F_TARIH, 7: F_INT})
    r = liste(r, f"{vm.DEVAMSIZLIK_ESIGI} gün ve üzeri devamsızlık", "Mazeretsiz devamsızlık (DV)",
              ["#", "Ad Soyad", "TC", "Alt ekip", "Görevi", "Devamsız gün", "Rapor gün", "Devam %"],
              [[k["ad_soyad"], k["tc"], k["ekip"], k["gorev"], k["gun"]["devamsiz"], k["gun"]["rapor"],
                round(k["devam"] / 100, 4)] for k in rk["devamsiz"]], {6: F_INT, 7: F_INT, 8: F_PCT})
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=SON)
    _yaz(ws, r + 1, 1, "Bu liste bilgi amaçlıdır; hukuki değerlendirme için İK / hukuk birimine danışın. Hesaplar SAP "
                       "mesai kayıtlarına dayanır.", font=_font(9, False, MUTED, True),
         align=Alignment(wrap_text=True))
    ws.row_dimensions[r + 1].height = 28
    _yazdir(ws, yatay=False)


# ================================================================ 10) PERSONEL HAREKETİ

def _hareket(wb, d):
    ws = wb.create_sheet("Personel Hareketi")
    ws.sheet_properties.tabColor = GREEN
    SON = 7
    _genislik(ws, [6, 26, 14, 18, 20, 18, 13])
    _sayfa_basligi(ws, SON, "PERSONEL HAREKETİ", d)
    h = d["hareket"]
    r = _bolum(ws, 5, 1, SON, "Özet", "Kimlik kayıtlarındaki işe giriş / işten çıkış tarihlerine göre")
    for i, (e, v, f) in enumerate([("Dönem başı mevcut", h["bas_mevcut"], F_INT), ("İşe başlayan", len(h["girenler"]), F_INT),
                                   ("Ayrılan", len(h["cikanlar"]), F_INT), ("Dönem sonu mevcut", h["son_mevcut"], F_INT),
                                   ("Net değişim", h["son_mevcut"] - h["bas_mevcut"], '+#,##0;-#,##0;0'),
                                   ("Devir oranı (ayrılan ÷ ortalama mevcut)", h["devir_orani"] / 100, F_PCT)]):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _yaz(ws, r, 2, e, font=_font(10, True, INK2))
        _yaz(ws, r, 5, v, f, _font(11, True), align=Alignment(horizontal="right"))
        _veri_satiri(ws, r, 2, 5, zebra=(i % 2 == 1))
        r += 1
    for baslik, liste, alan in (("İşe başlayanlar", h["girenler"], "giris"), ("Ayrılanlar", h["cikanlar"], "cikis")):
        r = _bolum(ws, r + 1, 1, SON, f"{baslik} ({len(liste)})")
        _baslik_satiri(ws, r, 1, ["#", "Ad Soyad", "TC", "Grup", "Alt ekip", "Görevi",
                                  "Giriş tarihi" if alan == "giris" else "Çıkış tarihi"], 22)
        r += 1
        if not liste:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=SON)
            _yaz(ws, r, 1, "Bu dönemde kayıt yok.", font=_font(10, False, MUTED, True))
            r += 1
        for i, k in enumerate(liste, start=1):
            for c, v in enumerate([i, k["ad_soyad"], k["tc"], k["grup"], k["ekip"], k["gorev"], k[alan]], start=1):
                _yaz(ws, r, c, v, F_TARIH if c == 7 else None, _font(10, c == 2))
            _veri_satiri(ws, r, 1, SON, zebra=(i % 2 == 0))
            r += 1
    _yazdir(ws, yatay=False)


# ================================================================ 11) YÖNTEM

def _yontem(wb, d):
    ws = wb.create_sheet("Yöntem")
    ws.sheet_properties.tabColor = "8C99A6"
    _genislik(ws, [3, 30, 90])
    _sayfa_basligi(ws, 3, "YÖNTEM VE TANIMLAR", d)
    r = _bolum(ws, 5, 2, 3, "Göstergeler")
    tanimlar = [
        ("Planlanan iş günü", "Kişinin kaydı olan günler − hafta tatili − resmî tatil − boş kayıt."),
        ("Çalışılan gün", "Saat girilmiş (normal + fazla > 0) ya da durumu ÇALIŞTI olan gün."),
        ("Devam oranı", "Çalışılan gün ÷ planlanan iş günü."),
        ("Kayıp iş günü", "Devamsızlık + sağlık raporu + ücretsiz izin. Yıllık / ücretli izin kayıp sayılmaz."),
        ("Kapasite kullanımı", f"Normal mesai saati ÷ (planlanan iş günü × {report.format_tr(vm.STANDART_GUN, 0)} sa). "
                               "Fazla mesai katılmaz; devamsızlık ve kısa günler oranı düşürür."),
        ("Kişi başı günlük saat", "(Normal + fazla mesai) ÷ çalışılan gün."),
        ("FM oranı", "Fazla mesai ÷ (normal + fazla mesai)."),
        ("Yevmiye karşılığı", f"{vm.katsayi_metni()}."),
        ("Pazar ve hafta tatili", vm.pazar_kurali_metni()),
        ("FM yevmiye karşılığı", "Yevmiyenin, hafta içi normal saatlerin (×1) dışında kalan kısmı: fazla mesai ve "
                                 "hak edilmiş pazar saatleri."),
        ("FM katsayı farkı", f"Aynı saatler normal mesaide çalışılsaydı ödenmeyecek kısım: {vm.katsayi_farki_metni()} "
                             f"yevmiye."),
        ("Verimlilik puanı", "Devam oranı × 0,6 + kapasite kullanımı (en fazla %100) × 0,4; 0–100 arası. "
                             "Not: A ≥ 95 · B 85–95 · C 70–85 · D < 70."),
        ("Günlük ortalama çalışan", "Çalışılan adam-gün ÷ verisi olan gün sayısı."),
        ("Önceki dönem", "Seçilen dönemle aynı uzunlukta, hemen önceki dönem."),
        ("Devir oranı", "Dönemde ayrılan ÷ (dönem başı + dönem sonu mevcut) / 2."),
    ]
    for i, (a, b) in enumerate(tanimlar):
        _yaz(ws, r, 2, a, font=_font(10, True), align=Alignment(vertical="top"))
        _yaz(ws, r, 3, b, font=_font(10, False, INK2), align=Alignment(wrap_text=True, vertical="top"))
        _veri_satiri(ws, r, 2, 3, zebra=(i % 2 == 1))
        ws.row_dimensions[r].height = 17 if len(b) <= 95 else 14 * (len(b) // 95 + 1) + 2
        r += 1
    r = _bolum(ws, r + 1, 2, 3, "Çalışma durumu sınıflandırması")
    siniflar = [("Çalıştı", "Saat girilmiş her gün (hafta tatilinde çalışma dahil) ya da 'ÇALIŞ…' içeren durum"),
                ("Hafta tatili (HT)", "'HAFTA TATİLİ'"), ("Resmî tatil (RT)", "'RESMİ', 'BAYRAM', 'GENEL TATİL'"),
                ("Ücretsiz izin (Ü)", "'ÜCRETSİZ' içeren"), ("Rapor (R)", "'RAPOR', 'HASTA', 'KAZA' içeren"),
                ("Devamsızlık (D)", "'DEVAMSIZ', 'GELMEDİ', 'MAZERETSİZ' içeren"),
                ("Ücretli izin (İ)", "diğer 'İZİN / İZNİ' içerenler (yıllık, babalık, vefat …)"),
                ("Kayıt boş / diğer (?)", "durum boş ve saat 0 — planlanan güne sayılmaz")]
    for i, (a, b) in enumerate(siniflar):
        _yaz(ws, r, 2, a, font=_font(10, True))
        _yaz(ws, r, 3, b, font=_font(10, False, INK2))
        _veri_satiri(ws, r, 2, 3, zebra=(i % 2 == 1))
        r += 1
    r = _bolum(ws, r + 1, 2, 3, "Yasal dayanaklar ve eşikler")
    for i, (a, b) in enumerate([("Yıllık fazla mesai", "4857 s. İş Kanunu md. 41 — yılda en fazla 270 saat"
                                 + vm.esik_notu(vm.YILLIK_FM_SINIRI, 270)),
                                ("Günlük çalışma", "md. 63 — günlük çalışma 11 saati aşamaz"
                                 + vm.esik_notu(vm.GUNLUK_AZAMI_SAAT, 11)),
                                ("Hafta tatili", "md. 46 — 7 günlük süre içinde en az 24 saat kesintisiz dinlenme"),
                                ("Aylık FM eşiği", f"Veri Merkezi ayarı: {report.format_tr(d['esik_aylik'], 0)} saat")]):
        _yaz(ws, r, 2, a, font=_font(10, True))
        _yaz(ws, r, 3, b, font=_font(10, False, INK2))
        _veri_satiri(ws, r, 2, 3, zebra=(i % 2 == 1))
        r += 1
    r = _bolum(ws, r + 1, 2, 3, "Kaynak")
    _yaz(ws, r, 2, "Veri", font=_font(10, True))
    _yaz(ws, r, 3, f"Veri Merkezi kalıcı deposu — SAP günlük mesai ({report.format_tr(d['kayit_sayisi'], 0)} kişi-gün "
                   f"kaydı) ve personel kimlik kayıtları", font=_font(10, False, INK2))
    if d["eksik_gun"]:
        r += 1
        _yaz(ws, r, 2, "Eksik günler", font=_font(10, True, RED))
        _yaz(ws, r, 3, ", ".join(f"{t:%d.%m.%Y}" for t in d["eksik_gun"]), font=_font(10, False, RED),
             align=Alignment(wrap_text=True))
    _yazdir(ws, yatay=False)


# ================================================================ ana fonksiyon

def dosya_adi(d):
    return f"{report._slugify(vm.kapsam_adi(d))}_verimlilik_{d['bas']:%Y%m%d}_{d['bit']:%Y%m%d}.xlsx"


def excel_yaz(d, output_path=None, ucret=None, db_path=None):
    if not d["kayit_sayisi"]:
        raise ValueError("Seçilen dönem ve kapsamda puantaj kaydı yok.")
    if output_path is None:
        report.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = report.OUTPUT_DIR / dosya_adi(d)
    wb = openpyxl.Workbook()
    _ozet(wb, d, ucret)
    _trend(wb, d)
    _gunluk(wb, d)
    etiket = "Alt ekip" if d["kirilim"] == "ekip" else "Grup"
    _kars_sayfasi(wb, d, f"{etiket.upper()} KARŞILAŞTIRMASI", "Ekip" if d["kirilim"] == "ekip" else "Grup",
                  d["kirilimlar"], etiket, NAVY)
    _kars_sayfasi(wb, d, "GÖREV KARŞILAŞTIRMASI", "Görev", d["gorev_tablosu"], "Görevi", NAVY)
    _isi(wb, d)
    _karne(wb, d, ucret)
    _takvim(wb, d, db_path)
    _risk(wb, d)
    _hareket(wb, d)
    _yontem(wb, d)
    wb.properties.title = f"Verimlilik Analizi — {vm.kapsam_adi(d)} — {vm.donem_adi(d)}"
    wb.properties.creator = "Veri Merkezi"
    wb.properties.subject = "MAOG Projesi iş gücü verimliliği"
    wb.active = 0
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
