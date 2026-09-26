# -*- coding: utf-8 -*-
"""
PERSONEL MALİYET RAPORU MOTORU (cost_report.py)
==================================================
Kaynak puantaj excelinden (Fazla Mesai Raporu / Ekip Listesi ile aynı
mekanizma) seçilen grup/alt ekip/meslek kapsamındaki personeli TC Kimlik
No üzerinden hakediş (bordro/maliyet) exceliyle eşleştirir ve detaylı bir
personel maliyet raporu üretir.

Hakediş dosyası genellikle çok sayıda sayfa içerir (aylar, Türk/yabancı
personel ayrımı vb.) ve sütun düzeni sayfadan sayfaya kayabilir (ör. bir
sayfada bir sütun eksik/fazla olabilir). Bu yüzden sütunlar SABİT harflerle
değil, başlık METNİYLE eşleştirilir (find_payroll_sheets / build_column_map) -
böylece dosya yapısı hafifçe değişse bile araç çalışmaya devam eder.

KULLANIM: cost_report_gui bölümünden (puantaj_suite.py), ya da doğrudan:
    generate_cost_report(CONFIG, source_path=..., hakedis_path=..., ...)
"""

import re
import datetime
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import puantaj_engine as engine
import team_report as report

SCRIPT_DIR = Path(__file__).resolve().parent
HAKEDIS_DIR = SCRIPT_DIR / "HAKEDIS"
OUTPUT_DIR = report.OUTPUT_DIR

# ============================== CONFIG ==============================
CONFIG = {
    # Kaynak puantaj (personel seçimi) - team_report.CONFIG ile aynı alanlar
    "source_sheet": report.CONFIG["source_sheet"],
    "source_group_col": report.CONFIG["source_group_col"],
    "source_subteam_col": report.CONFIG["source_subteam_col"],
    "source_role_col": report.CONFIG["source_role_col"],
    "source_name_col": report.CONFIG["source_name_col"],
    "source_type_col": report.CONFIG["source_type_col"],
    "source_tc_col": report.CONFIG["source_tc_col"],
    "source_date_header_row": report.CONFIG["source_date_header_row"],
    "source_data_start_row": report.CONFIG["source_data_start_row"],
    "source_data_end_row": report.CONFIG["source_data_end_row"],
    "source_date_col_start": report.CONFIG["source_date_col_start"],
    "source_date_col_end": report.CONFIG["source_date_col_end"],

    # Hakediş dosyası tarama ayarları
    "hakedis_header_row": 8,
    "hakedis_search_rows": (8, 7, 6, 5),
    "hakedis_data_start_row": 9,

    "cost_output_name": "Personel_Maliyet_Raporu.xlsx",
    # Bu değeri aşan Yevmiye Bedeli / Saatlik Net gibi birim değerler
    # "anormal / kontrol edilmeli" olarak işaretlenir (veri hatası belirtisi).
    "unit_rate_sanity_ceiling": 10000.0,
}
# ======================================================================

# --- Tasarım sabitleri: team_report ile birebir aynı görsel dil ---
NAVY_DARK = report.NAVY_DARK
HOURS_FMT = report.HOURS_FMT
INT_FMT = report.INT_FMT
CURRENCY_FMT = '#,##0.00 "TL"'
WARN_FILL = report.WARN_FILL
ZEBRA_FILL = report.ZEBRA_FILL

# --- Hakediş excelindeki alan -> başlık eşleştirme kuralları ---
# (anahtar, tüm bu alt dizeleri içermeli, hiçbirini içermemeli)
FIELD_RULES = [
    ("tc", ["T.C."], []),
    ("ad_soyad", ["ADI SOYADI"], []),
    ("gorevi", ["GÖREVİ"], []),
    ("makine", ["MAKİNE"], []),
    ("brut_aylik", ["BRÜT ÜCRET", "AYLIK"], []),
    ("brut_gunluk", ["BRÜT ÜCRET", "GÜNLÜK"], []),
    ("normal_gun", ["NORMAL ÇALIŞMA"], []),
    ("reel_fm_saat", ["REEL FAZLA MESAİ"], []),
    ("fiziki_calisilan_gun", ["FİZİKİ ÇALIŞILAN GÜN"], []),
    ("sigorta_gun", ["SİGORTA GÜN SAYISI"], []),
    ("yevmiye_gun", ["HAKEDİŞE ESAS YEVMİYE"], []),
    ("odenecek_mesai_saat", ["ÖDENECEK TOPLAM MESAİ"], []),
    ("pazar_calisma_gun", ["PAZAR ÇALIŞMASI"], []),
    ("yevmiye_bedeli", ["YEVMİYE BEDELİ"], ["PERSONELE"]),
    ("saatlik_net", ["SAATLİK NET ÜCRET"], []),
    ("sgk_maliyeti", ["SGK MALİYETİ"], []),
    ("toplam_z9", ["TOPLAM", "Z9"], []),
    ("odenecek_yevmiye_tutar", ["PERSONELE", "YEVMİYE"], ["RAPORLU"]),
    ("odenecek_fm_tutar", ["PERSONELE", "FAZLA MESAİ"], []),
    ("odenecek_raporlu_gun", ["PERSONELE", "RAPORLU"], []),
    ("sgk_odemesi_tutar", ["SGK ÖDEMESİ"], []),
    ("firma_payi_tutar", ["FİRMA PAYI"], []),
    ("toplam_odenecek", ["TOPLAM", "ÖDENECEK", "Z17"], []),
    ("personele_odenecek_toplam", ["PERSONELE", "TOPLAM TUTAR"], ["YEVMİYE", "FAZLA MESAİ", "RAPORLU"]),
]

MONEY_FIELDS = ["brut_aylik", "brut_gunluk", "yevmiye_bedeli", "saatlik_net", "sgk_maliyeti",
                 "toplam_z9", "odenecek_yevmiye_tutar", "odenecek_fm_tutar", "sgk_odemesi_tutar",
                 "firma_payi_tutar", "toplam_odenecek", "personele_odenecek_toplam"]


class CostReportError(engine.TransferError):
    pass


def normalize_header(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\n", " ")).strip().upper()


def build_column_map(ws, header_row=8, search_rows=(8, 7, 6, 5)):
    """Başlık metnine göre {alan_adı: sütun_no} eşlemesi çıkarır.
    Sütun sırası dosyadan dosyaya kaysa bile (ör. Özbek sayfası 1 sütun
    kayık) doğru sütunu bulur."""
    texts = {}
    for r in search_rows:
        for c in range(1, ws.max_column + 1):
            if c not in texts:
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    texts[c] = normalize_header(v)

    colmap = {}
    claimed = set()
    for key, all_of, none_of in FIELD_RULES:
        for c in sorted(texts):
            if c in claimed:
                continue
            t = texts[c]
            if all(a in t for a in all_of) and not any(n in t for n in none_of):
                colmap[key] = c
                claimed.add(c)
                break
    return colmap


def find_payroll_sheets(wb, header_row=8, search_rows=(8, 7, 6, 5)):
    """Çalışma kitabındaki hangi sayfaların bir 'hakediş/bordro' sayfası
    olduğunu (T.C. + ADI SOYADI başlığı bulunarak) tespit eder."""
    found = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        try:
            colmap = build_column_map(ws, header_row, search_rows)
        except Exception:
            continue
        if "tc" in colmap and "ad_soyad" in colmap:
            found.append(sn)
    return found


def read_hakedis_records(cfg, hakedis_path, sheet_names):
    """TC -> hakediş alanları sözlüğü. Aynı TC birden fazla sayfada
    geçerse SONUNCU bulunan sayfa esas alınır (uyarı listesine eklenir)."""
    wb = openpyxl.load_workbook(hakedis_path, data_only=True)
    header_row = cfg.get("hakedis_header_row", 8)
    search_rows = cfg.get("hakedis_search_rows", (8, 7, 6, 5))
    data_start_row = cfg.get("hakedis_data_start_row", 9)

    records = {}
    duplicate_tc = []
    per_sheet_info = {}

    for sn in sheet_names:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        colmap = build_column_map(ws, header_row, search_rows)
        if "tc" not in colmap:
            per_sheet_info[sn] = 0
            continue

        count = 0
        for r in range(data_start_row, ws.max_row + 1):
            tc_raw = ws.cell(row=r, column=colmap["tc"]).value
            tc = engine.normalize_tc(tc_raw)
            if not tc:
                continue
            rec = {"_sheet": sn}
            for key, col in colmap.items():
                if key == "tc":
                    continue
                rec[key] = ws.cell(row=r, column=col).value
            if tc in records and records[tc]["_sheet"] != sn:
                duplicate_tc.append((tc, records[tc]["_sheet"], sn))
            records[tc] = rec
            count += 1
        per_sheet_info[sn] = count

    return records, per_sheet_info, duplicate_tc


def build_cost_data(cfg, source_path, hakedis_path, hakedis_sheets,
                     team_filter=None, subteam_filter=None, role_filter=None):
    roster = report.build_employee_roster(cfg, source_path, team_filter, subteam_filter, role_filter)
    hakedis_records, per_sheet_info, duplicate_tc = read_hakedis_records(cfg, hakedis_path, hakedis_sheets)

    ceiling = cfg.get("unit_rate_sanity_ceiling", 10000.0)
    merged, not_found, quality_flags = [], [], []

    for rec in roster["records"]:
        tc = rec["tc"]
        h = hakedis_records.get(tc)
        if not h:
            not_found.append(rec)
            continue

        row = dict(rec)
        for key in [k for k, _, _ in FIELD_RULES if k != "tc"]:
            row[key] = h.get(key)
        row["_hakedis_sheet"] = h.get("_sheet")

        flags = []
        personele_toplam = row.get("personele_odenecek_toplam") or 0
        if personele_toplam and not row.get("sgk_odemesi_tutar"):
            flags.append("SGK ödemesi hesaplanmamış")
        if personele_toplam and not row.get("firma_payi_tutar"):
            flags.append("Firma payı hesaplanmamış")
        yb = row.get("yevmiye_bedeli") or 0
        if isinstance(yb, (int, float)) and yb > ceiling:
            flags.append(f"Yevmiye bedeli anormal yüksek ({report.format_tr(yb)} TL)")
        if flags:
            quality_flags.append((tc, row.get("name"), flags))
        row["_flags"] = flags

        merged.append(row)

    return {
        "merged": merged,
        "not_found": not_found,
        "quality_flags": quality_flags,
        "duplicate_tc": duplicate_tc,
        "hakedis_sheet_counts": per_sheet_info,
        "warning": roster["warning"],
        "team_filter": roster["team_filter"],
        "subteam_filter": roster["subteam_filter"],
        "role_filter": roster["role_filter"],
    }


# ---------------------------------------------------------------- rapor yazma

def _scope_str(data):
    if data["subteam_filter"]:
        s = " / ".join(data["subteam_filter"])
        if data["team_filter"]:
            s = f"{' / '.join(data['team_filter'])} — {s}"
    elif data["team_filter"]:
        s = " / ".join(data["team_filter"])
    else:
        s = "TÜM GRUPLAR"
    if data.get("role_filter"):
        roles_txt = " / ".join(data["role_filter"])
        if len(roles_txt) > 60:
            roles_txt = f"{len(data['role_filter'])} meslek grubu"
        s = f"{s}  ·  {roles_txt}"
    return s


def _sum(records, key):
    return sum((r.get(key) or 0) for r in records if isinstance(r.get(key), (int, float)))


def write_cost_report(cfg, data, output_path):
    merged = data["merged"]
    if not merged:
        raise CostReportError("Seçilen personelden hiçbiri hakediş dosyasında bulunamadı.")

    scope_str = _scope_str(data)
    grand_toplam_odenecek = _sum(merged, "toplam_odenecek")
    grand_personele_odenecek = _sum(merged, "personele_odenecek_toplam")
    grand_sgk = _sum(merged, "sgk_odemesi_tutar")
    grand_firma = _sum(merged, "firma_payi_tutar")
    grand_fm_tutar = _sum(merged, "odenecek_fm_tutar")

    wb = openpyxl.Workbook()

    # ==================================================================
    # PERSONEL MALİYET LİSTESİ
    # ==================================================================
    ws = wb.active
    ws.title = "Personel Maliyet Listesi"
    ws.sheet_view.showGridLines = False

    headers = ["Ekip", "Bağlı Olduğu Ekip", "Ad Soyad", "TC Kimlik No", "Görevi",
               "Brüt Aylık (TL)", "Yevmiye Bedeli (TL)", "Hakedişe Esas Yevmiye (gün)",
               "Reel Fazla Mesai (saat)", "Ödenecek Fazla Mesai (saat)",
               "Personele Ödenen - Yevmiye (TL)", "Personele Ödenen - Fazla Mesai (TL)",
               "Personele Ödenen TOPLAM (TL)", "SGK Ödemesi (TL)", "Firma Payı (TL)",
               "TOPLAM MALİYET (TL)", "Uyarı"]
    last_col = len(headers)

    report.banner(ws, f"A1:{get_column_letter(last_col)}1",
                   f"{scope_str} — PERSONEL MALİYET LİSTESİ", height=28)

    header_row = 3
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    report.style_header_row(ws, header_row, 1, last_col)
    ws.row_dimensions[header_row].height = 34

    money_cols = {6, 7, 11, 12, 13, 14, 15, 16}
    hour_cols = {9, 10}
    int_cols = {8}

    records_sorted = sorted(merged, key=lambda r: -(r.get("toplam_odenecek") or 0))
    f = header_row + 1
    for i, r in enumerate(records_sorted):
        row = f + i
        vals = [
            r.get("team"), r.get("subteam"), r.get("name"), r.get("tc"), r.get("role"),
            r.get("brut_aylik"), r.get("yevmiye_bedeli"), r.get("yevmiye_gun"),
            r.get("reel_fm_saat"), r.get("odenecek_mesai_saat"),
            r.get("odenecek_yevmiye_tutar"), r.get("odenecek_fm_tutar"),
            r.get("personele_odenecek_toplam"), r.get("sgk_odemesi_tutar"),
            r.get("firma_payi_tutar"), r.get("toplam_odenecek"),
            "; ".join(r.get("_flags") or []),
        ]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=row, column=c, value=v)
            if c in money_cols:
                cell.number_format = CURRENCY_FMT
            elif c in hour_cols:
                cell.number_format = HOURS_FMT
            elif c in int_cols:
                cell.number_format = HOURS_FMT
        fill = WARN_FILL if r.get("_flags") else (ZEBRA_FILL if i % 2 == 1 else None)
        report.style_data_row(ws, row, 1, last_col, fill=fill)
    last_data_row = f + len(records_sorted) - 1

    total_row = last_data_row + 1
    ws.cell(row=total_row, column=3, value="GENEL TOPLAM")
    ws.cell(row=total_row, column=11, value=round(_sum(merged, "odenecek_yevmiye_tutar"), 2)).number_format = CURRENCY_FMT
    ws.cell(row=total_row, column=12, value=round(grand_fm_tutar, 2)).number_format = CURRENCY_FMT
    ws.cell(row=total_row, column=13, value=round(grand_personele_odenecek, 2)).number_format = CURRENCY_FMT
    ws.cell(row=total_row, column=14, value=round(grand_sgk, 2)).number_format = CURRENCY_FMT
    ws.cell(row=total_row, column=15, value=round(grand_firma, 2)).number_format = CURRENCY_FMT
    ws.cell(row=total_row, column=16, value=round(grand_toplam_odenecek, 2)).number_format = CURRENCY_FMT
    report.style_total_row(ws, total_row, 1, last_col)

    if last_data_row >= f:
        ws.conditional_formatting.add(f"P{f}:P{last_data_row}",
                                        DataBarRule(start_type="min", end_type="max", color="5B9BD5"))

    widths = [22, 22, 26, 15, 26, 15, 15, 12, 12, 14, 16, 16, 16, 14, 14, 16, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{f}"
    report.setup_print(ws)

    # ==================================================================
    # EKİP BAZLI MALİYET
    # ==================================================================
    row_dim_key = "subteam" if (data["team_filter"] and len(data["team_filter"]) == 1) else "team"
    row_dim_label = "Bağlı Olduğu Ekip" if row_dim_key == "subteam" else "Ekip"

    groups = {}
    for r in merged:
        key = r.get(row_dim_key) or "-"
        groups.setdefault(key, []).append(r)

    ws2 = wb.create_sheet("Ekip Bazlı Maliyet")
    ws2.sheet_view.showGridLines = False
    report.banner(ws2, "A1:F1", f"{scope_str} — EKİP BAZLI MALİYET", height=28)

    h2 = 3
    headers2 = [row_dim_label, "Personel Sayısı", "Personele Ödenen Toplam (TL)",
                "SGK Ödemesi (TL)", "Firma Payı (TL)", "TOPLAM MALİYET (TL)"]
    for c, h in enumerate(headers2, start=1):
        ws2.cell(row=h2, column=c, value=h)
    report.style_header_row(ws2, h2, 1, 6)
    ws2.row_dimensions[h2].height = 28

    groups_sorted = sorted(groups.items(), key=lambda kv: -_sum(kv[1], "toplam_odenecek"))
    f2 = h2 + 1
    for i, (name, recs) in enumerate(groups_sorted):
        row = f2 + i
        ws2.cell(row=row, column=1, value=name)
        ws2.cell(row=row, column=2, value=len(recs)).number_format = INT_FMT
        ws2.cell(row=row, column=3, value=round(_sum(recs, "personele_odenecek_toplam"), 2)).number_format = CURRENCY_FMT
        ws2.cell(row=row, column=4, value=round(_sum(recs, "sgk_odemesi_tutar"), 2)).number_format = CURRENCY_FMT
        ws2.cell(row=row, column=5, value=round(_sum(recs, "firma_payi_tutar"), 2)).number_format = CURRENCY_FMT
        t = ws2.cell(row=row, column=6, value=round(_sum(recs, "toplam_odenecek"), 2))
        t.number_format = CURRENCY_FMT
        t.font = Font(bold=True)
        report.style_data_row(ws2, row, 1, 6, zebra=(i % 2 == 1))
    l2 = f2 + len(groups_sorted) - 1

    tot2 = l2 + 1
    ws2.cell(row=tot2, column=1, value="GENEL TOPLAM")
    ws2.cell(row=tot2, column=2, value=len(merged)).number_format = INT_FMT
    ws2.cell(row=tot2, column=3, value=round(grand_personele_odenecek, 2)).number_format = CURRENCY_FMT
    ws2.cell(row=tot2, column=4, value=round(grand_sgk, 2)).number_format = CURRENCY_FMT
    ws2.cell(row=tot2, column=5, value=round(grand_firma, 2)).number_format = CURRENCY_FMT
    ws2.cell(row=tot2, column=6, value=round(grand_toplam_odenecek, 2)).number_format = CURRENCY_FMT
    report.style_total_row(ws2, tot2, 1, 6)

    for col, w in zip("ABCDEF", (26, 16, 22, 18, 16, 20)):
        ws2.column_dimensions[col].width = w
    report.setup_print(ws2)

    if l2 >= f2:
        bar = BarChart()
        bar.type = "bar"
        bar.title = f"{row_dim_label} Bazında Toplam Maliyet"
        bar.x_axis.title = "TL"
        bar.height = max(7, min(20, 0.6 * len(groups_sorted) + 3))
        bar.width = 19
        bar.visible_cells_only = False
        bar.add_data(Reference(ws2, min_col=6, max_col=6, min_row=h2, max_row=l2), titles_from_data=True)
        bar.set_categories(Reference(ws2, min_col=1, max_col=1, min_row=f2, max_row=l2))
        bar.legend = None
        ws2.add_chart(bar, "H3")

        pie = PieChart()
        pie.title = "Genel Maliyet Dağılımı (Personel / SGK / Firma Payı)"
        pie.height = 8
        pie.width = 12
        cat_ws = wb.create_sheet("_yardimci")
        cat_ws.sheet_state = "hidden"
        cat_ws["A1"] = "Kalem"
        cat_ws["B1"] = "Tutar"
        cat_ws["A2"] = "Personele Ödenen"
        cat_ws["B2"] = round(grand_personele_odenecek, 2)
        cat_ws["A3"] = "SGK Ödemesi"
        cat_ws["B3"] = round(grand_sgk, 2)
        cat_ws["A4"] = "Firma Payı"
        cat_ws["B4"] = round(grand_firma, 2)
        pie.add_data(Reference(cat_ws, min_col=2, max_col=2, min_row=1, max_row=4), titles_from_data=True)
        pie.set_categories(Reference(cat_ws, min_col=1, max_col=1, min_row=2, max_row=4))
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        ws2.add_chart(pie, "H24")

    # ==================================================================
    # GÖREV BAZLI MALİYET
    # ==================================================================
    role_groups = {}
    for r in merged:
        role_groups.setdefault(r.get("role") or "(Görev belirtilmemiş)", []).append(r)

    ws3 = wb.create_sheet("Görev Bazlı Maliyet")
    ws3.sheet_view.showGridLines = False
    report.banner(ws3, "A1:E1", f"{scope_str} — GÖREV BAZLI MALİYET", height=28)

    h3 = 3
    for c, h in enumerate(["Görevi", "Personel Sayısı", "Personele Ödenen Toplam (TL)",
                            "TOPLAM MALİYET (TL)", "Kişi Başı Maliyet (TL)"], start=1):
        ws3.cell(row=h3, column=c, value=h)
    report.style_header_row(ws3, h3, 1, 5)
    ws3.row_dimensions[h3].height = 28

    role_sorted = sorted(role_groups.items(), key=lambda kv: -_sum(kv[1], "toplam_odenecek"))
    f3 = h3 + 1
    for i, (role, recs) in enumerate(role_sorted):
        row = f3 + i
        cnt = len(recs)
        tot = _sum(recs, "toplam_odenecek")
        ws3.cell(row=row, column=1, value=role)
        ws3.cell(row=row, column=2, value=cnt).number_format = INT_FMT
        ws3.cell(row=row, column=3, value=round(_sum(recs, "personele_odenecek_toplam"), 2)).number_format = CURRENCY_FMT
        ws3.cell(row=row, column=4, value=round(tot, 2)).number_format = CURRENCY_FMT
        ws3.cell(row=row, column=5, value=round(tot / cnt if cnt else 0, 2)).number_format = CURRENCY_FMT
        report.style_data_row(ws3, row, 1, 5, zebra=(i % 2 == 1))
    l3 = f3 + len(role_sorted) - 1

    tot3 = l3 + 1
    ws3.cell(row=tot3, column=1, value="GENEL TOPLAM")
    ws3.cell(row=tot3, column=2, value=len(merged)).number_format = INT_FMT
    ws3.cell(row=tot3, column=3, value=round(grand_personele_odenecek, 2)).number_format = CURRENCY_FMT
    ws3.cell(row=tot3, column=4, value=round(grand_toplam_odenecek, 2)).number_format = CURRENCY_FMT
    ws3.cell(row=tot3, column=5,
             value=round(grand_toplam_odenecek / len(merged) if merged else 0, 2)).number_format = CURRENCY_FMT
    report.style_total_row(ws3, tot3, 1, 5)

    if l3 >= f3:
        ws3.conditional_formatting.add(f"D{f3}:D{l3}",
                                         DataBarRule(start_type="min", end_type="max", color="7E57C2"))
        rchart = BarChart()
        rchart.type = "bar"
        rchart.title = "Göreve Göre Toplam Maliyet"
        rchart.x_axis.title = "TL"
        rchart.height = max(7, min(22, 0.6 * len(role_sorted) + 3))
        rchart.width = 19
        rchart.visible_cells_only = False
        rchart.add_data(Reference(ws3, min_col=4, max_col=4, min_row=h3, max_row=l3), titles_from_data=True)
        rchart.set_categories(Reference(ws3, min_col=1, max_col=1, min_row=f3, max_row=l3))
        rchart.legend = None
        ws3.add_chart(rchart, "G3")

    for col, w in zip("ABCDE", (34, 16, 22, 18, 18)):
        ws3.column_dimensions[col].width = w
    report.setup_print(ws3)

    # ==================================================================
    # EŞLEŞMEYEN PERSONEL (varsa)
    # ==================================================================
    if data["not_found"]:
        ws4 = wb.create_sheet("Eşleşmeyen Personel")
        ws4.sheet_view.showGridLines = False
        report.banner(ws4, "A1:D1", f"{scope_str} — HAKEDİŞTE BULUNAMAYAN PERSONEL", height=28)
        ws4.cell(row=2, column=1,
                 value="Bu kişiler kaynak puantajda seçili kapsamda bulundu, "
                       "ancak hakediş dosyasındaki taranan sayfalarda aynı TC ile eşleşmedi.").font = report.SUBTITLE_FONT
        ws4.merge_cells("A2:D2")

        h4 = 4
        for c, h in enumerate(["Ekip", "Bağlı Olduğu Ekip", "Ad Soyad", "TC Kimlik No"], start=1):
            ws4.cell(row=h4, column=c, value=h)
        report.style_header_row(ws4, h4, 1, 4)

        f4 = h4 + 1
        for i, rec in enumerate(data["not_found"]):
            row = f4 + i
            ws4.cell(row=row, column=1, value=rec.get("team"))
            ws4.cell(row=row, column=2, value=rec.get("subteam"))
            ws4.cell(row=row, column=3, value=rec.get("name"))
            ws4.cell(row=row, column=4, value=rec.get("tc"))
            report.style_data_row(ws4, row, 1, 4, zebra=(i % 2 == 1))

        for col, w in zip("ABCD", (24, 24, 28, 16)):
            ws4.column_dimensions[col].width = w
        report.setup_print(ws4)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------- giriş noktası

def generate_cost_report(cfg, source_path=None, hakedis_path=None, hakedis_sheets=None,
                          team_filter=None, subteam_filter=None, role_filter=None,
                          output_name=None):
    source_path = source_path or engine.find_single_xlsx(report.SOURCE_DIR, "SOURCE")
    if hakedis_path is None:
        hakedis_path = engine.find_single_xlsx(HAKEDIS_DIR, "HAKEDIS")
    cfg = dict(cfg)

    if hakedis_sheets is None:
        wb = openpyxl.load_workbook(hakedis_path, data_only=True)
        hakedis_sheets = find_payroll_sheets(wb, cfg.get("hakedis_header_row", 8),
                                              cfg.get("hakedis_search_rows", (8, 7, 6, 5)))
        if not hakedis_sheets:
            raise CostReportError("Hakediş dosyasında T.C./Adı Soyadı başlıklı bir sayfa bulunamadı.")

    data = build_cost_data(cfg, source_path, hakedis_path, hakedis_sheets,
                            team_filter, subteam_filter, role_filter)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / (output_name or cfg.get("cost_output_name", "Personel_Maliyet_Raporu.xlsx"))
    saved_path = write_cost_report(cfg, data, output_path)
    return saved_path, data, _scope_str(data)

