# -*- coding: utf-8 -*-
"""
EKİP / PERSONEL FAZLA MESAİ RAPORU (team_report.py) - v4 (detaylı analiz)
============================================================================
Kaynak aylık puantaj excelinden, seçilen ekip(ler)/bağlı olduğu ekip(ler) ve
seçilen tarih aralığında detaylı, grafikli bir fazla mesai raporu üretir.

v4'te eklenenler:
  - Başlıklar seçili filtreyi (ekip + bağlı olduğu ekip) doğru yansıtır
  - Tek ekip/alt ekip seçildiğinde KPI'lar ve sayfalar kişi bazlı analize döner
  - Kısa tarih aralıklarında (<= 10 gün) otomatik GÜNLÜK kırılım
  - Günlük Kırılım sayfası (hangi gün ne kadar, hafta sonu vurgusu)
  - Personel Analizi: dağılım grafiği, istatistik özeti (ort/medyan/min/maks)
  - Görev bazlı gruplama sayfası
  - Eşik uyarı sistemi (belirlenen saati aşanlar kırmızı işaretlenir)
  - Önceki eşit uzunluktaki dönemle karşılaştırma (değişim %)
  - Kapakta özet grafik, iş günü/gün sayısı bilgisi

KULLANIM: team_report_gui.py üzerinden, ya da:
    python team_report.py
"""

import datetime
import statistics
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.page import PageMargins

import puantaj_engine as engine

SCRIPT_DIR = Path(__file__).resolve().parent
SOURCE_DIR = SCRIPT_DIR / "SOURCE"
OUTPUT_DIR = SCRIPT_DIR / "OUTPUT"

# ============================== CONFIG ==============================
CONFIG = {
    "source_sheet": "PERSONEL.0526",
    "source_group_col": "B",        # GRUPLAR (ekip)
    "source_subteam_col": "K",      # BAĞLI OLDUĞU EKİP
    "source_role_col": "L",         # GÖREVİ
    "source_name_col": "I",         # AD SOYAD
    "source_type_col": "R",         # NORMAL MESAİ / FAZLA MESAİ
    "source_tc_col": "J",
    "source_date_header_row": 3,
    "source_data_start_row": 4,
    "source_data_end_row": 800,
    "source_date_col_start": "S",
    "source_date_col_end": "AW",
    "source_giris_col": "O",         # İŞE GİRİŞ TARİHİ
    "source_cikis_col": "P",         # İŞTEN ÇIKIŞ TARİHİ

    "output_name": "Ekip_Mesai_Raporu.xlsx",
    "overtime_threshold": 52.0,      # AYLIK bazda: bu saati aşan personelin fazla mesaisi kritik sayılır
                                      # (rapor ay dışında bir aralık kapsıyorsa gün sayısına göre orantılanır)
    "standard_workday_hours": 9.0,   # yevmiye günü hesaplamasında 1 günlük karşılık (saat)
    "auto_daily_max_days": 10,      # bu kadar veya daha kısa aralıkta günlük kırılıma geç
}
# ======================================================================

# Fazla mesai ücret katsayıları (yevmiye günü hesaplamasında kullanılır):
# Pazartesi-Cumartesi %50 zamlı (1.5x), Pazar günü 2.5 kat.
FAZLA_MESAI_KATSAYI_HAFTAICI = 1.5
FAZLA_MESAI_KATSAYI_PAZAR = 2.5

GUN_ADLARI = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]

TURKCE_AYLAR = {1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN", 5: "MAYIS", 6: "HAZİRAN",
                7: "TEMMUZ", 8: "AĞUSTOS", 9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"}


def ay_adi(d):
    """Verilen tarihin Türkçe ay adını büyük harfle döner (rapor başlıklarında kullanılır)."""
    return TURKCE_AYLAR.get(d.month, "")


def _effective_data_end_row(cfg, ws):
    """cfg["source_data_end_row"] her ay elle güncellenmesi gereken sabit bir
    değerdir; personel sayısı arttıkça kaynak dosya büyür ve bu değer eskiyip
    kalabilir (satırların sessizce atlanmasına yol açar). Güvenlik için,
    yapılandırılan değer ile sayfanın gerçek son satırından BÜYÜK olanı kullanılır -
    böylece dosya configden daha büyürse veri asla sessizce kaybolmaz."""
    configured = cfg.get("source_data_end_row", 800)
    actual = ws.max_row or configured
    return max(configured, actual)

# Kaynakta boş bırakılmış alt ekip / görev hücreleri için görünen etiketler
NO_SUBTEAM = "(Ekip belirtilmemiş)"
NO_ROLE = "(Görev belirtilmemiş)"

# Kademe 4: personnel_db'deki alan adlarının rapor sayfasında gösterilecek
# Türkçe karşılıkları (team_report.py, döngüsel içe aktarmayı önlemek için
# personnel_db'ye bağımlı değildir - bu eşleme burada bağımsız tutulur).
FIELD_LABEL_TR = {
    "ad_soyad": "Ad Soyad",
    "grup": "Grup",
    "bagli_oldugu_ekip": "Bağlı Olduğu Ekip",
    "gorevi": "Görevi",
    "ise_giris_tarihi": "İşe Giriş Tarihi",
    "isten_cikis_tarihi": "İşten Çıkış Tarihi",
    "notlar": "Notlar",
}

# --- Tasarım sabitleri -------------------------------------------------
NAVY = "1F4E78"
NAVY_DARK = "16344F"
ACCENT = "2E7D32"
LIGHT_GRAY = "F2F5F7"
BORDER_COLOR = "C7D0D6"
WARN_FILL_COLOR = "FFC7CE"
WEEKEND_FILL_COLOR = "FFF3E0"

HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, color=NAVY_DARK, size=18)
SUBTITLE_FONT = Font(color="546E7A", size=11, italic=True)
KPI_LABEL_FONT = Font(color="546E7A", size=10, bold=True)
TOTAL_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
ZEBRA_FILL = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
WARN_FILL = PatternFill(start_color=WARN_FILL_COLOR, end_color=WARN_FILL_COLOR, fill_type="solid")
WEEKEND_FILL = PatternFill(start_color=WEEKEND_FILL_COLOR, end_color=WEEKEND_FILL_COLOR, fill_type="solid")
THIN_SIDE = Side(style="thin", color=BORDER_COLOR)
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HOURS_FMT = "#,##0.0"
INT_FMT = "#,##0"
PCT_FMT = '+0.0"%";-0.0"%";"—"'


def format_tr(value, decimals=1):
    s = f"{value:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_date(d):
    return d.strftime("%d.%m.%Y")


def period_block_of(date_val, date_start, period_days):
    if period_days is None:
        return 1
    return (date_val - date_start).days // period_days + 1


# ---------------------------------------------------------------- dosya tarama (GUI için)

def scan_filter_options(cfg, source_path):
    wb = openpyxl.load_workbook(source_path, data_only=True)
    sheet_name, warning = engine.resolve_sheet(wb, cfg["source_sheet"], "Kaynak")
    ws = wb[sheet_name]

    group_col = column_index_from_string(cfg["source_group_col"])
    subteam_col = column_index_from_string(cfg["source_subteam_col"])
    role_col = column_index_from_string(cfg["source_role_col"])
    date_start_col = column_index_from_string(cfg["source_date_col_start"])
    date_end_col = column_index_from_string(cfg["source_date_col_end"])

    teams, subteams, roles = set(), set(), set()
    # hiyerarşi: grup -> alt ekip -> meslek grupları
    hierarchy = {}

    for r in range(cfg["source_data_start_row"], _effective_data_end_row(cfg, ws) + 1):
        t = ws.cell(row=r, column=group_col).value
        if not t:
            continue
        team = str(t).strip()
        teams.add(team)

        st = ws.cell(row=r, column=subteam_col).value
        subteam = str(st).strip() if st else NO_SUBTEAM
        subteams.add(subteam)

        rl = ws.cell(row=r, column=role_col).value
        role = str(rl).strip() if rl else NO_ROLE
        roles.add(role)

        hierarchy.setdefault(team, {}).setdefault(subteam, set()).add(role)

    dates = []
    for c in range(date_start_col, date_end_col + 1):
        v = ws.cell(row=cfg["source_date_header_row"], column=c).value
        if isinstance(v, datetime.datetime):
            v = v.date()
        if isinstance(v, datetime.date):
            dates.append(v)

    # setleri GUI'nin rahat kullanması için sıralı listelere çevir
    hierarchy_sorted = {
        team: {sub: sorted(rs) for sub, rs in sorted(subs.items())}
        for team, subs in sorted(hierarchy.items())
    }

    return {
        "teams": sorted(teams),
        "subteams": sorted(subteams),
        "roles": sorted(roles),
        "hierarchy": hierarchy_sorted,
        "date_min": min(dates) if dates else None,
        "date_max": max(dates) if dates else None,
        "warning": warning,
    }


def subteams_for(hierarchy, selected_teams=None):
    """Seçili grup(lar)a bağlı alt ekiplerin sıralı listesi.
    selected_teams boşsa tüm alt ekipler döner."""
    result = set()
    for team, subs in hierarchy.items():
        if selected_teams and team not in selected_teams:
            continue
        result.update(subs.keys())
    return sorted(result)


def roles_for(hierarchy, selected_teams=None, selected_subteams=None):
    """Seçili grup ve alt ekip(ler)e ait meslek gruplarının sıralı listesi."""
    result = set()
    for team, subs in hierarchy.items():
        if selected_teams and team not in selected_teams:
            continue
        for sub, rs in subs.items():
            if selected_subteams and sub not in selected_subteams:
                continue
            result.update(rs)
    return sorted(result)


# ---------------------------------------------------------------- veri toplama

def _collect(cfg, ws, day_columns, team_set, subteam_set, want_details, role_set=None,
              breakdown_by="team", tc_filter=None, dataset_lookup=None):
    """Verilen gün sütunları için NORMAL + FAZLA MESAİ toplamlarını ayrı ayrı,
    ayrıca çalışılan gün sayısını (yevmiye günü) çıkarır.
    breakdown_by: "team" -> grup bazında kırılım, "subteam" -> alt ekip bazında.
    tc_filter: verilirse (Veri Seti modu), sadece bu TC kümesindeki personel
    işlenir - kaynaktaki GRUPLAR/ALT EKİP/GÖREV filtreleri bu modda uygulanmaz.
    dataset_lookup: {tc: {"grup":.., "bagli_oldugu_ekip":.., "gorevi":..}} verilirse,
    kaynaktaki (muhtemelen bozuk/eski) Grup/Alt Ekip/Görev bilgisi yerine veri
    setindeki güncel bilgi kullanılır (Kademe 1: geçmiş ay verisi düzeltme).
    Filtreleme de bu DÜZELTİLMİŞ değerlere göre yapılır."""
    group_col = column_index_from_string(cfg["source_group_col"])
    subteam_col = column_index_from_string(cfg["source_subteam_col"])
    role_col = column_index_from_string(cfg["source_role_col"])
    name_col = column_index_from_string(cfg["source_name_col"])
    type_col = column_index_from_string(cfg["source_type_col"])
    tc_col = column_index_from_string(cfg["source_tc_col"])

    team_period_normal = {}    # bucket -> {period: normal saat}
    team_period_fazla = {}     # bucket -> {period: fazla saat}
    team_worked_days = {}      # bucket -> toplam FİİLEN çalışılan gün sayısı (doluluk oranı için)
    team_unattributed_normal = {}  # bucket -> TC'si boş/geçersiz olduğu için kişiye atfedilemeyen normal saat
    team_unattributed_fazla = {}   # bucket -> TC'si boş/geçersiz olduğu için kişiye atfedilemeyen fazla saat
    team_fazla_haftaici = {}   # bucket -> Pazartesi-Cumartesi fazla mesai saati (x1.5)
    team_fazla_pazar = {}      # bucket -> Pazar fazla mesai saati (x2.5)
    team_employees = {}
    employee_records = {}
    employee_normal_totals = {}   # tc -> toplam normal saat
    employee_fazla_totals = {}    # tc -> toplam fazla saat (ham, ağırlıksız)
    employee_worked_days = {}     # tc -> fiilen çalışılan gün sayısı (doluluk oranı için)
    employee_fazla_haftaici = {}  # tc -> Pazartesi-Cumartesi fazla mesai saati
    employee_fazla_pazar = {}     # tc -> Pazar fazla mesai saati
    employee_day_totals = {}      # tc -> {date: fazla saat}
    employee_day_normal = {}      # tc -> {date: normal saat}
    day_totals = {}               # date -> toplam fazla saat
    day_totals_normal = {}        # date -> toplam normal saat

    # Kademe 2: veri kalitesi tespiti
    missing_subteam_tcs = set()
    missing_role_tcs = set()
    tc_issues = []              # [(row, tc, name), ...] format/checksum sorunlu
    dataset_mismatches = []     # [(tc, name, field, kaynak_deger, veriseti_degeri), ...]
    tc_name_variants = {}       # tc -> {isim1, isim2, ...} - aynı TC farklı isimlerde görülüyorsa (paylaşılan/hatalı TC)
    seen_tc_for_quality = set()

    for r in range(cfg["source_data_start_row"], _effective_data_end_row(cfg, ws) + 1):
        tip = ws.cell(row=r, column=type_col).value
        if not tip:
            continue
        team_raw = ws.cell(row=r, column=group_col).value
        if not team_raw:
            continue
        team = str(team_raw).strip()

        tc = engine.normalize_tc(ws.cell(row=r, column=tc_col).value)

        subteam_raw = ws.cell(row=r, column=subteam_col).value
        subteam = str(subteam_raw).strip() if subteam_raw else NO_SUBTEAM

        role_raw = ws.cell(row=r, column=role_col).value
        role = str(role_raw).strip() if role_raw else NO_ROLE

        name_val = str(ws.cell(row=r, column=name_col).value or "").strip()
        pre_override_subteam = subteam
        pre_override_role = role

        # --- Kademe 1: Veri Seti ile düzeltme (override) ---
        pending_mismatches = []
        if dataset_lookup and tc in dataset_lookup:
            dl = dataset_lookup[tc]
            new_team = (dl.get("grup") or "").strip()
            new_subteam = (dl.get("bagli_oldugu_ekip") or "").strip()
            new_role = (dl.get("gorevi") or "").strip()
            if new_team and new_team != team:
                pending_mismatches.append((tc, name_val, "Grup", team, new_team))
                team = new_team
            if new_subteam and new_subteam != subteam:
                pending_mismatches.append((tc, name_val, "Bağlı Olduğu Ekip", subteam, new_subteam))
                subteam = new_subteam
            elif not subteam_raw and new_subteam:
                subteam = new_subteam
            if new_role and new_role != role:
                pending_mismatches.append((tc, name_val, "Görevi", role, new_role))
                role = new_role
            elif not role_raw and new_role:
                role = new_role

        # --- filtreleme (düzeltilmiş değerlere göre) ---
        if tc_filter is not None:
            if tc not in tc_filter:
                continue
        else:
            if team_set and team not in team_set:
                continue
            if subteam_set and subteam not in subteam_set:
                continue
            if role_set and role not in role_set:
                continue

        # --- Kademe 2: kaynak veri kalitesi (sadece kapsamdaki personel için,
        # düzeltmeden ÖNCEki ham hale bakılır) ---
        if tc and name_val:
            tc_name_variants.setdefault(tc, set()).add(name_val)
        if tc and tc not in seen_tc_for_quality:
            seen_tc_for_quality.add(tc)
            if pre_override_subteam == NO_SUBTEAM:
                missing_subteam_tcs.add(tc)
            if pre_override_role == NO_ROLE:
                missing_role_tcs.add(tc)
            issue = engine.check_tc(tc, "kaynak", r, name_val)
            if issue:
                tc_issues.append(issue)
        dataset_mismatches.extend(pending_mismatches)

        # raporun satır kırılımı: grup ya da alt ekip
        bucket = subteam if breakdown_by == "subteam" else team

        if tc:
            team_employees.setdefault(bucket, set()).add(tc)

        tip_upper = str(tip).upper()
        is_normal = "NORMAL" in tip_upper
        is_fazla = "FAZLA" in tip_upper

        if want_details and is_normal and tc and tc not in employee_records:
            employee_records[tc] = {
                "team": team,
                "subteam": subteam,
                "name": name_val,
                "role": role,
            }

        if is_normal:
            periods = team_period_normal.setdefault(bucket, {})
            row_total = 0.0
            worked_days = 0
            for col, date_val, p in day_columns:
                val = ws.cell(row=r, column=col).value
                if isinstance(val, (int, float)):
                    periods[p] = periods.get(p, 0.0) + val
                    row_total += val
                    worked_days += 1
                    day_totals_normal[date_val] = day_totals_normal.get(date_val, 0.0) + val
                    if tc:
                        employee_day_normal.setdefault(tc, {})[date_val] = \
                            employee_day_normal.setdefault(tc, {}).get(date_val, 0.0) + val
            if tc:
                employee_normal_totals[tc] = employee_normal_totals.get(tc, 0.0) + row_total
                employee_worked_days[tc] = employee_worked_days.get(tc, 0) + worked_days
                team_worked_days[bucket] = team_worked_days.get(bucket, 0) + worked_days
            else:
                team_unattributed_normal[bucket] = team_unattributed_normal.get(bucket, 0.0) + row_total

        if is_fazla:
            periods = team_period_fazla.setdefault(bucket, {})
            row_total = 0.0
            for col, date_val, p in day_columns:
                val = ws.cell(row=r, column=col).value
                if isinstance(val, (int, float)):
                    periods[p] = periods.get(p, 0.0) + val
                    row_total += val
                    day_totals[date_val] = day_totals.get(date_val, 0.0) + val
                    if tc:
                        employee_day_totals.setdefault(tc, {})[date_val] = \
                            employee_day_totals.setdefault(tc, {}).get(date_val, 0.0) + val
                    # Pazar (haftanın 7. günü) %250, diğer günler %150 zamlı ödenir
                    if date_val.weekday() == 6:
                        day_bucket_fazla, day_emp_fazla = team_fazla_pazar, employee_fazla_pazar
                    else:
                        day_bucket_fazla, day_emp_fazla = team_fazla_haftaici, employee_fazla_haftaici
                    day_bucket_fazla[bucket] = day_bucket_fazla.get(bucket, 0.0) + val
                    if tc:
                        day_emp_fazla[tc] = day_emp_fazla.get(tc, 0.0) + val
            if tc:
                employee_fazla_totals[tc] = employee_fazla_totals.get(tc, 0.0) + row_total
            else:
                team_unattributed_fazla[bucket] = team_unattributed_fazla.get(bucket, 0.0) + row_total

    # --- Yevmiye Günü: normal saat + ücret katsayılı fazla mesai saatinin
    # standart iş günü saatine oranı (Pazartesi-Cumartesi x1.5, Pazar x2.5) ---
    std_hours = cfg.get("standard_workday_hours", 9.0) or 9.0

    employee_yevmiye_days = {}
    for tc in set(employee_normal_totals) | set(employee_fazla_haftaici) | set(employee_fazla_pazar):
        normal_h = employee_normal_totals.get(tc, 0.0)
        fw = employee_fazla_haftaici.get(tc, 0.0)
        fs = employee_fazla_pazar.get(tc, 0.0)
        employee_yevmiye_days[tc] = (normal_h + fw * FAZLA_MESAI_KATSAYI_HAFTAICI +
                                       fs * FAZLA_MESAI_KATSAYI_PAZAR) / std_hours

    team_yevmiye_days = {}
    for bucket in set(team_period_normal) | set(team_fazla_haftaici) | set(team_fazla_pazar):
        normal_h = sum(team_period_normal.get(bucket, {}).values())
        fw = team_fazla_haftaici.get(bucket, 0.0)
        fs = team_fazla_pazar.get(bucket, 0.0)
        team_yevmiye_days[bucket] = (normal_h + fw * FAZLA_MESAI_KATSAYI_HAFTAICI +
                                       fs * FAZLA_MESAI_KATSAYI_PAZAR) / std_hours

    return {
        "team_period_normal": team_period_normal,
        "team_period_fazla": team_period_fazla,
        "team_yevmiye_days": team_yevmiye_days,
        "team_worked_days": team_worked_days,
        "team_unattributed_normal": team_unattributed_normal,
        "team_unattributed_fazla": team_unattributed_fazla,
        "team_employees": team_employees,
        "employee_records": employee_records,
        "employee_normal_totals": employee_normal_totals,
        "employee_fazla_totals": employee_fazla_totals,
        "employee_fazla_haftaici": employee_fazla_haftaici,
        "employee_fazla_pazar": employee_fazla_pazar,
        "employee_yevmiye_days": employee_yevmiye_days,
        "employee_worked_days": employee_worked_days,
        "employee_day_totals": employee_day_totals,
        "employee_day_normal": employee_day_normal,
        "day_totals": day_totals,
        "day_totals_normal": day_totals_normal,
        "missing_subteam_tcs": missing_subteam_tcs,
        "missing_role_tcs": missing_role_tcs,
        "tc_issues": tc_issues,
        "dataset_mismatches": dataset_mismatches,
        "tc_name_variants": tc_name_variants,
    }


def _day_columns_for(cfg, ws, date_start, date_end, period_days):
    date_col_start = column_index_from_string(cfg["source_date_col_start"])
    date_col_end = column_index_from_string(cfg["source_date_col_end"])
    cols = []
    for c in range(date_col_start, date_col_end + 1):
        v = ws.cell(row=cfg["source_date_header_row"], column=c).value
        if isinstance(v, datetime.datetime):
            v = v.date()
        if isinstance(v, datetime.date) and date_start <= v <= date_end:
            cols.append((c, v, period_block_of(v, date_start, period_days)))
    return cols


def _parse_date_cell(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _read_giris_cikis(cfg, ws, tc_set):
    """Verilen TC kümesi için İşe Giriş / İşten Çıkış Tarihi'ni okur
    (Kademe 3: kısmi ay / doluluk oranı hesaplaması için)."""
    tc_col = column_index_from_string(cfg["source_tc_col"])
    giris_col = column_index_from_string(cfg.get("source_giris_col", "O"))
    cikis_col = column_index_from_string(cfg.get("source_cikis_col", "P"))
    result = {}
    for r in range(cfg["source_data_start_row"], _effective_data_end_row(cfg, ws) + 1):
        tc = engine.normalize_tc(ws.cell(row=r, column=tc_col).value)
        if tc and tc in tc_set and tc not in result:
            result[tc] = {
                "giris": _parse_date_cell(ws.cell(row=r, column=giris_col).value),
                "cikis": _parse_date_cell(ws.cell(row=r, column=cikis_col).value),
            }
    return result


def _active_days_in_range(giris, cikis, dates_in_range):
    """Kişinin [giriş, çıkış] aralığı ile rapor dönemi kesişiminde,
    rapor dönemindeki gerçek tarih sütunlarına göre kaç gün aktif olduğunu sayar."""
    if not dates_in_range:
        return 0
    lo = giris if giris else dates_in_range[0]
    hi = cikis if cikis else dates_in_range[-1]
    return sum(1 for d in dates_in_range if lo <= d <= hi)


def build_report_data(cfg, source_path, date_start, date_end, period_days=7,
                       team_filter=None, subteam_filter=None, role_filter=None,
                       include_employee_list=False, compare_previous=True, tc_filter=None,
                       dataset_lookup=None):
    wb = openpyxl.load_workbook(source_path, data_only=True)
    sheet_name, warning = engine.resolve_sheet(wb, cfg["source_sheet"], "Kaynak")
    ws = wb[sheet_name]

    team_set = {t.strip() for t in team_filter} if team_filter else None
    subteam_set = {t.strip() for t in subteam_filter} if subteam_filter else None
    role_set = {t.strip() for t in role_filter} if role_filter else None

    # Tek grup seçiliyse rapor satırlarını o grubun ALT EKİPLERİ bazında kır;
    # birden fazla grup (veya hiç filtre, ya da Veri Seti modu) varsa GRUP bazında kır.
    breakdown_by = "subteam" if (team_set and len(team_set) == 1 and tc_filter is None) else "team"

    span_days = (date_end - date_start).days + 1
    auto_daily = False
    if period_days is not None and span_days <= cfg.get("auto_daily_max_days", 10):
        period_days = 1
        auto_daily = True

    day_columns = _day_columns_for(cfg, ws, date_start, date_end, period_days)
    if not day_columns:
        raise engine.TransferError("Seçilen tarih aralığında kaynak excelde veri sütunu bulunamadı.")

    period_blocks = sorted({p for _, _, p in day_columns})
    period_labels = {}
    for p in period_blocks:
        dates_in_block = [d for _, d, pb in day_columns if pb == p]
        lo, hi = min(dates_in_block), max(dates_in_block)
        if lo == hi:
            period_labels[p] = f"{lo:%d.%m} {GUN_ADLARI[lo.weekday()]}" if auto_daily else f"{lo:%d.%m}"
        else:
            period_labels[p] = f"{lo:%d.%m}-{hi:%d.%m}"

    current = _collect(cfg, ws, day_columns, team_set, subteam_set, want_details=True,
                        role_set=role_set, breakdown_by=breakdown_by, tc_filter=tc_filter,
                        dataset_lookup=dataset_lookup)

    # tüm ekipler için toplamları eksik dönemlerle doldur
    all_buckets = set(current["team_period_normal"]) | set(current["team_period_fazla"])
    for bucket in all_buckets:
        for p in period_blocks:
            current["team_period_normal"].setdefault(bucket, {}).setdefault(p, 0.0)
            current["team_period_fazla"].setdefault(bucket, {}).setdefault(p, 0.0)

    # --- önceki eşit uzunluktaki dönem (karşılaştırma) ---
    previous_total = None
    previous_range = None
    if compare_previous:
        prev_end = date_start - datetime.timedelta(days=1)
        prev_start = prev_end - datetime.timedelta(days=span_days - 1)
        prev_cols = _day_columns_for(cfg, ws, prev_start, prev_end, None)
        if prev_cols:
            prev = _collect(cfg, ws, prev_cols, team_set, subteam_set, want_details=False,
                             role_set=role_set, breakdown_by=breakdown_by, tc_filter=tc_filter,
                             dataset_lookup=dataset_lookup)
            prev_normal = sum(sum(v.values()) for v in prev["team_period_normal"].values())
            prev_fazla = sum(sum(v.values()) for v in prev["team_period_fazla"].values())
            previous_total = prev_normal + prev_fazla
            previous_range = (prev_start, prev_end)

    dates_in_range = sorted({d for _, d, _p in day_columns})

    # --- Kademe 3: kısmi ay / doluluk oranı ---
    all_tcs = {tc for s in current["team_employees"].values() for tc in s}
    giris_cikis = _read_giris_cikis(cfg, ws, all_tcs)
    employee_active_days = {}
    employee_doluluk = {}
    employee_durum = {}
    for tc in all_tcs:
        gc = giris_cikis.get(tc, {})
        giris, cikis = gc.get("giris"), gc.get("cikis")
        active = _active_days_in_range(giris, cikis, dates_in_range)
        employee_active_days[tc] = active
        worked = current["employee_worked_days"].get(tc, 0)
        employee_doluluk[tc] = round(worked / active * 100, 1) if active else 0.0
        parts = []
        if giris and giris > date_start:
            parts.append(f"Giriş: {giris:%d.%m.%Y}")
        if cikis and cikis < date_end:
            parts.append(f"Çıkış: {cikis:%d.%m.%Y}")
        employee_durum[tc] = f"Kısmi ({', '.join(parts)})" if parts else "Tam Dönem"

    return {
        "period_blocks": period_blocks,
        "period_labels": period_labels,
        "team_period_normal": current["team_period_normal"],
        "team_period_fazla": current["team_period_fazla"],
        "team_yevmiye_days": current["team_yevmiye_days"],
        "team_unattributed_normal": current["team_unattributed_normal"],
        "team_unattributed_fazla": current["team_unattributed_fazla"],
        "team_employees": current["team_employees"],
        "employee_records": current["employee_records"] if include_employee_list else {},
        "employee_normal_totals": current["employee_normal_totals"],
        "employee_fazla_totals": current["employee_fazla_totals"],
        "employee_yevmiye_days": current["employee_yevmiye_days"],
        "employee_day_totals": current["employee_day_totals"],
        "employee_day_normal": current["employee_day_normal"],
        "employee_active_days": employee_active_days,
        "employee_doluluk": employee_doluluk,
        "employee_durum": employee_durum,
        "day_totals": current["day_totals"],
        "day_totals_normal": current["day_totals_normal"],
        "dates_in_range": dates_in_range,
        "source_warning": warning,
        "date_start": date_start,
        "date_end": date_end,
        "span_days": span_days,
        "auto_daily": auto_daily,
        "team_filter": sorted(team_set) if team_set else None,
        "subteam_filter": sorted(subteam_set) if subteam_set else None,
        "role_filter": sorted(role_set) if role_set else None,
        "tc_filter_used": tc_filter is not None,
        "dataset_override_used": dataset_lookup is not None,
        "breakdown_by": breakdown_by,
        "previous_total": previous_total,
        "previous_range": previous_range,
        "threshold": cfg.get("overtime_threshold", 52.0) * span_days / 30.0,
        "threshold_monthly": cfg.get("overtime_threshold", 52.0),
        # veri kalitesi (Kademe 2)
        "missing_subteam_count": len(current["missing_subteam_tcs"]),
        "missing_role_count": len(current["missing_role_tcs"]),
        "tc_issues": current["tc_issues"],
        "dataset_mismatches": current["dataset_mismatches"],
        "duplicate_tc_names": {tc: sorted(names) for tc, names in current["tc_name_variants"].items() if len(names) > 1},
        # personel listesi kapalıyken bile kişi analizi için gerekli olabilir
        "all_employee_records": current["employee_records"],
    }


# ---------------------------------------------------------------- biçimlendirme yardımcıları

def style_header_row(ws, row, first_col, last_col):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_data_row(ws, row, first_col, last_col, zebra=False, fill=None):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = THIN_BORDER
        if fill is not None:
            cell.fill = fill
        elif zebra:
            cell.fill = ZEBRA_FILL


def style_total_row(ws, row, first_col, last_col):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL
        cell.border = THIN_BORDER


def setup_print(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6, header=0.3, footer=0.3)
    ws.oddHeader.center.text = "Mesai Raporu"
    ws.oddFooter.center.text = "Sayfa &P / &N"


def banner(ws, cell_range, text, height=26, size=14):
    ws.merge_cells(cell_range)
    first = cell_range.split(":")[0]
    cell = ws[first]
    cell.value = text
    cell.font = Font(bold=True, size=size, color="FFFFFF")
    cell.fill = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[int("".join(ch for ch in first if ch.isdigit()))].height = height


# ---------------------------------------------------------------- rapor yazma

def write_report(cfg, data, output_path, history_changes=None):
    period_blocks = data["period_blocks"]
    period_labels = data["period_labels"]
    team_period_normal = data["team_period_normal"]
    team_period_fazla = data["team_period_fazla"]
    team_yevmiye_days = data["team_yevmiye_days"]
    team_employees = data["team_employees"]
    employee_normal_totals = data["employee_normal_totals"]
    employee_fazla_totals = data["employee_fazla_totals"]
    employee_yevmiye_days = data["employee_yevmiye_days"]
    all_records = data["all_employee_records"]
    threshold = data["threshold"]
    threshold_monthly = data.get("threshold_monthly", threshold)
    date_range_str = f"{fmt_date(data['date_start'])} - {fmt_date(data['date_end'])}"

    # --- kapsam etiketi: ekip + alt ekip filtresini birlikte yansıtır ---
    if data.get("tc_filter_used"):
        scope_str = "VERİ SETİNDEN SEÇİLİ PERSONEL"
    elif data["subteam_filter"]:
        scope_str = " / ".join(data["subteam_filter"])
        if data["team_filter"]:
            scope_str = f"{' / '.join(data['team_filter'])} — {scope_str}"
    elif data["team_filter"]:
        scope_str = " / ".join(data["team_filter"])
    else:
        scope_str = "TÜM GRUPLAR"

    if data.get("role_filter"):
        roles_txt = " / ".join(data["role_filter"])
        if len(roles_txt) > 60:
            roles_txt = f"{len(data['role_filter'])} meslek grubu"
        scope_str = f"{scope_str}  ·  {roles_txt}"

    def bucket_total(bucket):
        return sum(team_period_normal.get(bucket, {}).values()) + sum(team_period_fazla.get(bucket, {}).values())

    all_buckets = set(team_period_normal) | set(team_period_fazla)
    teams_sorted = sorted(all_buckets, key=bucket_total, reverse=True)
    if not teams_sorted:
        raise engine.TransferError("Seçilen filtrelerle eşleşen mesai verisi bulunamadı.")

    grand_normal = sum(sum(v.values()) for v in team_period_normal.values())
    grand_fazla = sum(sum(v.values()) for v in team_period_fazla.values())
    grand_total = grand_normal + grand_fazla
    grand_yevmiye = sum(team_yevmiye_days.get(b, 0) for b in teams_sorted)
    grand_employees = sum(len(team_employees.get(t, set())) for t in teams_sorted)
    grand_avg = grand_total / grand_employees if grand_employees else 0
    single_scope = len(teams_sorted) == 1 or bool(data["subteam_filter"]) or data.get("tc_filter_used")
    # satırların neyi temsil ettiği: grup mu, alt ekip mi
    row_dim = "Bağlı Olduğu Ekip" if data.get("breakdown_by") == "subteam" else "Ekip"

    # kişi bazlı istatistikler (normal + fazla + toplam + yevmiye günü)
    person_stats = {}
    for tc in {tc for s in team_employees.values() for tc in s}:
        normal_h = employee_normal_totals.get(tc, 0.0)
        fazla_h = employee_fazla_totals.get(tc, 0.0)
        person_stats[tc] = {
            "normal": normal_h, "fazla": fazla_h, "toplam": normal_h + fazla_h,
            "yevmiye": employee_yevmiye_days.get(tc, 0),
        }
    toplam_only = [p["toplam"] for p in person_stats.values()] or [0.0]
    stat_avg = statistics.mean(toplam_only)
    stat_median = statistics.median(toplam_only)
    stat_max = max(toplam_only)
    stat_min = min(toplam_only)
    stat_std = statistics.pstdev(toplam_only) if len(toplam_only) > 1 else 0.0
    over_threshold = [(tc, p["fazla"]) for tc, p in person_stats.items() if p["fazla"] > threshold]

    top_person_tc, top_person = (max(person_stats.items(), key=lambda kv: kv[1]["toplam"])
                                  if person_stats else (None, {"toplam": 0}))
    top_person_name = (all_records.get(top_person_tc, {}) or {}).get("name", top_person_tc or "-")

    workdays = len(data["dates_in_range"])
    weekend_days = sum(1 for d in data["dates_in_range"] if d.weekday() >= 5)
    daily_avg = grand_total / workdays if workdays else 0

    change_pct = None
    if data["previous_total"]:
        change_pct = (grand_total - data["previous_total"]) / data["previous_total"] * 100

    wb = openpyxl.Workbook()

    # ==================================================================
    # KAPAK
    # ==================================================================
    cover = wb.active
    cover.title = "Kapak"
    cover.sheet_view.showGridLines = False

    cover.merge_cells("B2:H3")
    tc_cell = cover.cell(row=2, column=2, value=f"{scope_str} ÇALIŞMA ANALİZİ RAPORU")
    tc_cell.font = TITLE_FONT
    tc_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    cover.merge_cells("B4:H4")
    cover.cell(row=4, column=2,
                value=f"Dönem: {date_range_str}  ({workdays} gün, {weekend_days} hafta sonu)   |   "
                      f"Kaynak: {Path(cfg['source_path']).name}   |   "
                      f"Oluşturma: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}").font = SUBTITLE_FONT

    if single_scope:
        kpis = [
            ("TOPLAM MESAİ (Normal + Fazla)", f"{format_tr(grand_total)} sa", ACCENT),
            ("NORMAL MESAİ", f"{format_tr(grand_normal)} sa", NAVY),
            ("FAZLA MESAİ", f"{format_tr(grand_fazla)} sa", NAVY),
            ("YEVMİYE GÜNÜ", f"{format_tr(grand_yevmiye)} gün", "00695C"),
            ("PERSONEL SAYISI", f"{grand_employees}", NAVY),
            ("KİŞİ BAŞI ORTALAMA (Toplam)", f"{format_tr(stat_avg)} sa", NAVY),
        ]
    else:
        kpis = [
            ("TOPLAM MESAİ (Normal + Fazla)", f"{format_tr(grand_total)} sa", ACCENT),
            ("NORMAL MESAİ", f"{format_tr(grand_normal)} sa", NAVY),
            ("FAZLA MESAİ", f"{format_tr(grand_fazla)} sa", NAVY),
            ("YEVMİYE GÜNÜ", f"{format_tr(grand_yevmiye)} gün", "00695C"),
            ("EKİP SAYISI", f"{len(teams_sorted)}", NAVY),
            ("TOPLAM PERSONEL", f"{grand_employees}", NAVY),
            ("KİŞİ BAŞI ORTALAMA (Toplam)", f"{format_tr(grand_avg)} sa", NAVY),
        ]

    if change_pct is not None:
        prev_s, prev_e = data["previous_range"]
        arrow = "▲" if change_pct > 0 else ("▼" if change_pct < 0 else "=")
        color = "B71C1C" if change_pct > 0 else ACCENT
        kpis.append((f"ÖNCEKİ DÖNEME GÖRE ({prev_s:%d.%m}-{prev_e:%d.%m}: {format_tr(data['previous_total'], 0)} sa)",
                      f"{arrow} %{format_tr(abs(change_pct))}", color))

    start_row = 6
    for i, (label, value, color) in enumerate(kpis):
        r = start_row + i * 2
        cover.merge_cells(f"B{r}:H{r}")
        cover.cell(row=r, column=2, value=label).font = KPI_LABEL_FONT
        cover.merge_cells(f"B{r+1}:H{r+1}")
        cover.cell(row=r + 1, column=2, value=value).font = Font(bold=True, size=15, color=color)

    cover.column_dimensions["A"].width = 2
    for col in "BCDEFGH":
        cover.column_dimensions[col].width = 14

    warn_row = start_row + len(kpis) * 2 + 1
    if data["source_warning"]:
        cover.cell(row=warn_row, column=2, value=f"UYARI: {data['source_warning']}").font = Font(color="B71C1C")

    # kapak için gizli veri + özet grafik (ekip bazında toplam mesai)
    helper_start = warn_row + 3
    cover.cell(row=helper_start, column=11, value=row_dim)
    cover.cell(row=helper_start, column=12, value="Normal")
    cover.cell(row=helper_start, column=13, value="Fazla")
    for i, b in enumerate(teams_sorted):
        cover.cell(row=helper_start + 1 + i, column=11, value=b)
        cover.cell(row=helper_start + 1 + i, column=12, value=round(sum(team_period_normal.get(b, {}).values()), 2))
        cover.cell(row=helper_start + 1 + i, column=13, value=round(sum(team_period_fazla.get(b, {}).values()), 2))
    for col in ("K", "L", "M"):
        cover.column_dimensions[col].hidden = True

    cover_chart = BarChart()
    cover_chart.type = "col"
    cover_chart.grouping = "stacked"
    cover_chart.overlap = 100
    cover_chart.title = f"{row_dim} Bazında Normal + Fazla Mesai"
    cover_chart.y_axis.title = "Saat"
    cover_chart.height = 9
    cover_chart.width = 20
    cover_chart.visible_cells_only = False
    cc_data = Reference(cover, min_col=12, max_col=13, min_row=helper_start, max_row=helper_start + len(teams_sorted))
    cc_cats = Reference(cover, min_col=11, max_col=11, min_row=helper_start + 1, max_row=helper_start + len(teams_sorted))
    cover_chart.add_data(cc_data, titles_from_data=True)
    cover_chart.set_categories(cc_cats)
    cover.add_chart(cover_chart, f"B{warn_row + 2}")
    setup_print(cover, landscape=False)

    # ==================================================================
    # EKİP-DÖNEM ÖZETİ
    # ==================================================================
    ws = wb.create_sheet("Ekip-Dönem Özeti")
    ws.sheet_view.showGridLines = False

    n_periods = len(period_blocks)
    normal_col = 3 + n_periods
    fazla_col = normal_col + 1
    toplam_col = fazla_col + 1
    yevmiye_col = toplam_col + 1
    avg_col = yevmiye_col + 1
    pay_col = avg_col + 1
    last_col = pay_col
    period_col_start, period_col_end = 3, 2 + n_periods

    label_kind = "GÜNLÜK" if data["auto_daily"] else "DÖNEM BAZLI"
    banner(ws, f"A1:{get_column_letter(last_col)}1",
           f"{scope_str} {ay_adi(data['date_start'])} AYI {label_kind} MESAİ ({date_range_str})")

    header_row = 3
    headers = ([row_dim, "Personel Sayısı"] + [period_labels[p] for p in period_blocks] +
               ["Normal Mesai (sa)", "Fazla Mesai (sa)", "Toplam Mesai (sa)", "Yevmiye Günü", "Kişi Başı Ort. (sa)", "Payı (%)"])
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    style_header_row(ws, header_row, 1, last_col)
    ws.row_dimensions[header_row].height = 32

    data_first_row = header_row + 1
    for i, b in enumerate(teams_sorted):
        r = data_first_row + i
        normals = team_period_normal.get(b, {})
        fazlas = team_period_fazla.get(b, {})
        emp = len(team_employees.get(b, set()))
        normal_total = sum(normals.values())
        fazla_total = sum(fazlas.values())
        period_total = normal_total + fazla_total
        yevmiye = team_yevmiye_days.get(b, 0)

        ws.cell(row=r, column=1, value=b)
        ws.cell(row=r, column=2, value=emp).number_format = INT_FMT
        for j, p in enumerate(period_blocks):
            combined = normals.get(p, 0.0) + fazlas.get(p, 0.0)
            ws.cell(row=r, column=3 + j, value=round(combined, 2)).number_format = HOURS_FMT
        ws.cell(row=r, column=normal_col, value=round(normal_total, 2)).number_format = HOURS_FMT
        ws.cell(row=r, column=fazla_col, value=round(fazla_total, 2)).number_format = HOURS_FMT
        t = ws.cell(row=r, column=toplam_col, value=round(period_total, 2))
        t.number_format = HOURS_FMT
        t.font = Font(bold=True)
        ws.cell(row=r, column=yevmiye_col, value=round(yevmiye, 1)).number_format = HOURS_FMT
        ws.cell(row=r, column=avg_col,
                value=round(period_total / emp if emp else 0, 2)).number_format = HOURS_FMT
        ws.cell(row=r, column=pay_col,
                value=round(period_total / grand_total if grand_total else 0, 4)).number_format = '0.0%'
        style_data_row(ws, r, 1, last_col, zebra=(i % 2 == 1))
    data_last_row = data_first_row + len(teams_sorted) - 1

    total_row = data_last_row + 1
    ws.cell(row=total_row, column=1, value="GENEL TOPLAM")
    ws.cell(row=total_row, column=2, value=grand_employees).number_format = INT_FMT
    for j, p in enumerate(period_blocks):
        combined = (sum(team_period_normal.get(b, {}).get(p, 0.0) for b in teams_sorted) +
                    sum(team_period_fazla.get(b, {}).get(p, 0.0) for b in teams_sorted))
        ws.cell(row=total_row, column=3 + j, value=round(combined, 2)).number_format = HOURS_FMT
    ws.cell(row=total_row, column=normal_col, value=round(grand_normal, 2)).number_format = HOURS_FMT
    ws.cell(row=total_row, column=fazla_col, value=round(grand_fazla, 2)).number_format = HOURS_FMT
    ws.cell(row=total_row, column=toplam_col, value=round(grand_total, 2)).number_format = HOURS_FMT
    ws.cell(row=total_row, column=yevmiye_col, value=round(grand_yevmiye, 1)).number_format = HOURS_FMT
    ws.cell(row=total_row, column=avg_col, value=round(grand_avg, 2)).number_format = HOURS_FMT
    ws.cell(row=total_row, column=pay_col, value=1.0).number_format = '0.0%'
    style_total_row(ws, total_row, 1, last_col)

    if data_last_row >= data_first_row and n_periods >= 1:
        ws.conditional_formatting.add(
            f"{get_column_letter(period_col_start)}{data_first_row}:{get_column_letter(period_col_end)}{data_last_row}",
            ColorScaleRule(start_type="min", start_color="C8E6C9",
                            mid_type="percentile", mid_value=50, mid_color="FFF59D",
                            end_type="max", end_color="EF9A9A"))
        ws.conditional_formatting.add(
            f"{get_column_letter(avg_col)}{data_first_row}:{get_column_letter(avg_col)}{data_last_row}",
            DataBarRule(start_type="min", end_type="max", color="5B9BD5"))

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 15
    for j in range(n_periods):
        ws.column_dimensions[get_column_letter(3 + j)].width = 18
    for col in (normal_col, fazla_col, toplam_col, avg_col):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions[get_column_letter(yevmiye_col)].width = 13
    ws.column_dimensions[get_column_letter(pay_col)].width = 12
    ws.freeze_panes = f"A{data_first_row}"
    setup_print(ws)

    if n_periods > 1:
        line = LineChart()
        line.title = f"{row_dim} Bazında Dönemsel Toplam Mesai Trendi"
        line.y_axis.title = "Mesai (saat)"
        line.x_axis.title = "Dönem"
        line.height = 11
        line.width = 26
        line.style = 2
        for i, b in enumerate(teams_sorted):
            r = data_first_row + i
            ref = Reference(ws, min_col=period_col_start, max_col=period_col_end, min_row=r, max_row=r)
            s = openpyxl.chart.Series(ref, title=b)
            s.marker = Marker(symbol="circle", size=6)
            s.smooth = False
            line.series.append(s)
        line.set_categories(Reference(ws, min_col=period_col_start, max_col=period_col_end,
                                       min_row=header_row, max_row=header_row))
        ws.add_chart(line, f"A{total_row + 3}")

    # ==================================================================
    # GÜNLÜK KIRILIM
    # ==================================================================
    wsd = wb.create_sheet("Günlük Kırılım")
    wsd.sheet_view.showGridLines = False
    banner(wsd, "A1:E1", f"{scope_str} {ay_adi(data['date_start'])} AYI GÜNLÜK MESAİ DAĞILIMI ({date_range_str})")

    hd = 2
    for c, h in enumerate(["Tarih", "Gün", "Normal Mesai (sa)", "Fazla Mesai (sa)", "Toplam Mesai (sa)"], start=1):
        wsd.cell(row=hd, column=c, value=h)
    style_header_row(wsd, hd, 1, 5)

    df_first = hd + 1
    for i, d in enumerate(data["dates_in_range"]):
        r = df_first + i
        normal_h = data["day_totals_normal"].get(d, 0.0)
        fazla_h = data["day_totals"].get(d, 0.0)
        wsd.cell(row=r, column=1, value=fmt_date(d))
        wsd.cell(row=r, column=2, value=GUN_ADLARI[d.weekday()])
        wsd.cell(row=r, column=3, value=round(normal_h, 2)).number_format = HOURS_FMT
        wsd.cell(row=r, column=4, value=round(fazla_h, 2)).number_format = HOURS_FMT
        t = wsd.cell(row=r, column=5, value=round(normal_h + fazla_h, 2))
        t.number_format = HOURS_FMT
        t.font = Font(bold=True)
        is_weekend = d.weekday() >= 5
        style_data_row(wsd, r, 1, 5, zebra=(i % 2 == 1),
                       fill=WEEKEND_FILL if is_weekend else None)
    df_last = df_first + len(data["dates_in_range"]) - 1

    dtot = df_last + 1
    wsd.cell(row=dtot, column=1, value="TOPLAM")
    wsd.cell(row=dtot, column=3, value=round(grand_normal, 2)).number_format = HOURS_FMT
    wsd.cell(row=dtot, column=4, value=round(grand_fazla, 2)).number_format = HOURS_FMT
    wsd.cell(row=dtot, column=5, value=round(grand_total, 2)).number_format = HOURS_FMT
    style_total_row(wsd, dtot, 1, 5)

    for col, w in zip("ABCDE", (14, 8, 16, 16, 16)):
        wsd.column_dimensions[col].width = w
    wsd.freeze_panes = f"A{df_first}"
    setup_print(wsd, landscape=False)

    if df_last >= df_first:
        dchart = BarChart()
        dchart.type = "col"
        dchart.grouping = "stacked"
        dchart.overlap = 100
        dchart.title = "Günlere Göre Normal + Fazla Mesai"
        dchart.y_axis.title = "Saat"
        dchart.height = 9
        dchart.width = 24
        dchart.visible_cells_only = False
        dchart.add_data(Reference(wsd, min_col=3, max_col=4, min_row=hd, max_row=df_last), titles_from_data=True)
        dchart.set_categories(Reference(wsd, min_col=1, max_col=1, min_row=df_first, max_row=df_last))
        wsd.add_chart(dchart, "G4")

    # ==================================================================
    # EKİP TOPLAM / PAY (sadece birden fazla ekip varsa anlamlı)
    # ==================================================================
    if len(teams_sorted) > 1:
        ws2 = wb.create_sheet("Ekip Toplam")
        ws2.sheet_view.showGridLines = False
        banner(ws2, "A1:F1", f"{scope_str} {ay_adi(data['date_start'])} AYI DÖNEM TOPLAMI VE PAY DAĞILIMI")

        h_row = 3
        for c, h in enumerate([row_dim, "Normal Mesai (sa)", "Fazla Mesai (sa)", "Toplam Mesai (sa)",
                                "Yevmiye Günü", "Payı (%)"], start=1):
            ws2.cell(row=h_row, column=c, value=h)
        style_header_row(ws2, h_row, 1, 6)

        d_first = h_row + 1
        for i, b in enumerate(teams_sorted):
            r = d_first + i
            nt = sum(team_period_normal.get(b, {}).values())
            ft = sum(team_period_fazla.get(b, {}).values())
            pt = nt + ft
            ws2.cell(row=r, column=1, value=b)
            ws2.cell(row=r, column=2, value=round(nt, 2)).number_format = HOURS_FMT
            ws2.cell(row=r, column=3, value=round(ft, 2)).number_format = HOURS_FMT
            tc_ = ws2.cell(row=r, column=4, value=round(pt, 2))
            tc_.number_format = HOURS_FMT
            tc_.font = Font(bold=True)
            ws2.cell(row=r, column=5, value=round(team_yevmiye_days.get(b, 0.0), 1)).number_format = HOURS_FMT
            ws2.cell(row=r, column=6,
                     value=round(pt / grand_total if grand_total else 0, 4)).number_format = '0.0%'
            style_data_row(ws2, r, 1, 6, zebra=(i % 2 == 1))
        d_last = d_first + len(teams_sorted) - 1

        tr2 = d_last + 1
        ws2.cell(row=tr2, column=1, value="GENEL TOPLAM")
        ws2.cell(row=tr2, column=2, value=round(grand_normal, 2)).number_format = HOURS_FMT
        ws2.cell(row=tr2, column=3, value=round(grand_fazla, 2)).number_format = HOURS_FMT
        ws2.cell(row=tr2, column=4, value=round(grand_total, 2)).number_format = HOURS_FMT
        ws2.cell(row=tr2, column=5, value=round(grand_yevmiye, 1)).number_format = HOURS_FMT
        ws2.cell(row=tr2, column=6, value=1.0).number_format = '0.0%'
        style_total_row(ws2, tr2, 1, 6)

        for col, w in zip("ABCDEF", (24, 16, 16, 16, 13, 12)):
            ws2.column_dimensions[col].width = w
        setup_print(ws2)

        bar2 = BarChart()
        bar2.type = "bar"
        bar2.grouping = "stacked"
        bar2.overlap = 100
        bar2.title = f"{row_dim} Bazında Normal + Fazla Mesai"
        bar2.height = 10
        bar2.width = 18
        bar2.visible_cells_only = False
        bar2.add_data(Reference(ws2, min_col=2, max_col=3, min_row=h_row, max_row=d_last), titles_from_data=True)
        bar2.set_categories(Reference(ws2, min_col=1, max_col=1, min_row=d_first, max_row=d_last))
        ws2.add_chart(bar2, "H3")

        pie = PieChart()
        pie.title = f"{row_dim} Bazında Toplam Mesai Payı"
        pie.height = 10
        pie.width = 14
        pie.visible_cells_only = False
        pie.add_data(Reference(ws2, min_col=4, max_col=4, min_row=h_row, max_row=d_last), titles_from_data=True)
        pie.set_categories(Reference(ws2, min_col=1, max_col=1, min_row=d_first, max_row=d_last))
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showCatName = False
        pie.dataLabels.showSerName = False
        pie.dataLabels.showVal = False
        pie.dataLabels.showLegendKey = False
        ws2.add_chart(pie, "H23")

        # ------------------------------------------------------------
        # Ekip bazında GÖREV DAĞILIMI detay tabloları (aynı sayfada, altta)
        # ------------------------------------------------------------
        if all_records:
            detail_cursor = tr2 + 3

            for b in teams_sorted:
                bucket_roles = {}
                for tc in team_employees.get(b, set()):
                    st = person_stats.get(tc, {"normal": 0.0, "fazla": 0.0, "yevmiye": 0.0})
                    role = (all_records.get(tc, {}) or {}).get("role") or "(Görev belirtilmemiş)"
                    br = bucket_roles.setdefault(role, {"count": 0, "normal": 0.0, "fazla": 0.0, "yevmiye": 0.0})
                    br["count"] += 1
                    br["normal"] += st["normal"]
                    br["fazla"] += st["fazla"]
                    br["yevmiye"] += st["yevmiye"]

                # TC'si boş/geçersiz olduğu için kimseye atfedilemeyen ama ekip
                # toplamına dahil olan saatler - şeffaflık için ayrı satırda gösterilir
                unattr_normal = data.get("team_unattributed_normal", {}).get(b, 0.0)
                unattr_fazla = data.get("team_unattributed_fazla", {}).get(b, 0.0)
                if unattr_normal or unattr_fazla:
                    std_h = cfg.get("standard_workday_hours", 9.0) or 9.0
                    unattr_yevmiye = (unattr_normal + unattr_fazla * FAZLA_MESAI_KATSAYI_HAFTAICI) / std_h
                    br = bucket_roles.setdefault("(TC Bilgisi Eksik)", {"count": 0, "normal": 0.0, "fazla": 0.0, "yevmiye": 0.0})
                    br["normal"] += unattr_normal
                    br["fazla"] += unattr_fazla
                    br["yevmiye"] += unattr_yevmiye

                if not bucket_roles:
                    continue

                bucket_toplam = sum(v["normal"] + v["fazla"] for v in bucket_roles.values())

                title_row = detail_cursor
                ws2.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=7)
                tcell = ws2.cell(row=title_row, column=1, value=f"{b} — GÖREV DAĞILIMI")
                tcell.font = Font(bold=True, size=12, color="FFFFFF")
                tcell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
                tcell.alignment = Alignment(vertical="center", indent=1)
                ws2.row_dimensions[title_row].height = 22

                dh = title_row + 1
                for c, h in enumerate(["Görevi", "Personel Sayısı", "Normal Mesai (sa)", "Fazla Mesai (sa)",
                                        "Toplam Mesai (sa)", "Yevmiye Günü", "Payı (%)"], start=1):
                    ws2.cell(row=dh, column=c, value=h)
                style_header_row(ws2, dh, 1, 7)

                roles_sorted = sorted(bucket_roles.items(),
                                       key=lambda kv: kv[1]["normal"] + kv[1]["fazla"], reverse=True)
                df = dh + 1
                for i, (role, br) in enumerate(roles_sorted):
                    r = df + i
                    role_toplam = br["normal"] + br["fazla"]
                    ws2.cell(row=r, column=1, value=role)
                    ws2.cell(row=r, column=2, value=br["count"]).number_format = INT_FMT
                    ws2.cell(row=r, column=3, value=round(br["normal"], 2)).number_format = HOURS_FMT
                    ws2.cell(row=r, column=4, value=round(br["fazla"], 2)).number_format = HOURS_FMT
                    rt = ws2.cell(row=r, column=5, value=round(role_toplam, 2))
                    rt.number_format = HOURS_FMT
                    rt.font = Font(bold=True)
                    ws2.cell(row=r, column=6, value=round(br["yevmiye"], 1)).number_format = HOURS_FMT
                    ws2.cell(row=r, column=7,
                             value=round(role_toplam / bucket_toplam if bucket_toplam else 0, 4)).number_format = '0.0%'
                    style_data_row(ws2, r, 1, 7, zebra=(i % 2 == 1))
                dl = df + len(roles_sorted) - 1

                dtot = dl + 1
                ws2.cell(row=dtot, column=1, value="TOPLAM")
                ws2.cell(row=dtot, column=2, value=sum(br["count"] for _r, br in roles_sorted)).number_format = INT_FMT
                ws2.cell(row=dtot, column=3,
                         value=round(sum(br["normal"] for _r, br in roles_sorted), 2)).number_format = HOURS_FMT
                ws2.cell(row=dtot, column=4,
                         value=round(sum(br["fazla"] for _r, br in roles_sorted), 2)).number_format = HOURS_FMT
                ws2.cell(row=dtot, column=5, value=round(bucket_toplam, 2)).number_format = HOURS_FMT
                ws2.cell(row=dtot, column=6,
                         value=round(sum(br["yevmiye"] for _r, br in roles_sorted), 1)).number_format = HOURS_FMT
                ws2.cell(row=dtot, column=7, value=1.0).number_format = '0.0%'
                style_total_row(ws2, dtot, 1, 7)

                if dl >= df:
                    ws2.conditional_formatting.add(f"E{df}:E{dl}",
                                                    DataBarRule(start_type="min", end_type="max", color="90A4AE"))

                detail_cursor = dtot + 2

        # Kişi Başı Yük (ekipler arası kıyas)
        ws3 = wb.create_sheet("Kişi Başı Yük")
        ws3.sheet_view.showGridLines = False
        banner(ws3, "A1:E1", f"{scope_str} — {row_dim.upper()} BAZINDA KİŞİ BAŞI YÜK")
        ws3.cell(row=2, column=1,
                 value="Toplamda küçük görünse de az personelle çalışan, kişi başına yüksek mesai düşen ekipleri gösterir.").font = SUBTITLE_FONT
        ws3.merge_cells("A2:E2")

        h3 = 4
        for c, h in enumerate([row_dim, "Personel Sayısı", "Toplam Mesai (sa)", "Yevmiye Günü", "Kişi Başı Ort. (sa)"], start=1):
            ws3.cell(row=h3, column=c, value=h)
        style_header_row(ws3, h3, 1, 5)

        teams_by_avg = sorted(teams_sorted, key=lambda b: (
            bucket_total(b) / len(team_employees[b]) if team_employees.get(b) else 0), reverse=True)
        f3 = h3 + 1
        for i, b in enumerate(teams_by_avg):
            r = f3 + i
            emp = len(team_employees.get(b, set()))
            pt = bucket_total(b)
            ws3.cell(row=r, column=1, value=b)
            ws3.cell(row=r, column=2, value=emp).number_format = INT_FMT
            ws3.cell(row=r, column=3, value=round(pt, 2)).number_format = HOURS_FMT
            ws3.cell(row=r, column=4, value=round(team_yevmiye_days.get(b, 0.0), 1)).number_format = HOURS_FMT
            ws3.cell(row=r, column=5, value=round(pt / emp if emp else 0, 2)).number_format = HOURS_FMT
            style_data_row(ws3, r, 1, 5, zebra=(i % 2 == 1))
        l3 = f3 + len(teams_by_avg) - 1
        ws3.conditional_formatting.add(f"E{f3}:E{l3}",
                                        DataBarRule(start_type="min", end_type="max", color="EF5350"))
        for col, w in zip("ABCDE", (24, 16, 18, 13, 18)):
            ws3.column_dimensions[col].width = w
        setup_print(ws3)

    # ==================================================================
    # PERSONEL ANALİZİ (istatistik + dağılım) - kişi verisi varsa
    # ==================================================================
    if all_records:
        wsa = wb.create_sheet("Personel Analizi")
        wsa.sheet_view.showGridLines = False
        banner(wsa, "A1:E1", f"{scope_str} — PERSONEL DAĞILIMI VE İSTATİSTİKLER ({date_range_str})")

        stats = [
            ("Personel sayısı", grand_employees, INT_FMT),
            ("Toplam mesai (normal + fazla, saat)", round(grand_total, 2), HOURS_FMT),
            ("Toplam normal mesai (saat)", round(grand_normal, 2), HOURS_FMT),
            ("Toplam fazla mesai (saat)", round(grand_fazla, 2), HOURS_FMT),
            ("Toplam yevmiye günü", round(grand_yevmiye, 1), HOURS_FMT),
            ("Ortalama (toplam saat/kişi)", round(stat_avg, 2), HOURS_FMT),
            ("Medyan (toplam saat/kişi)", round(stat_median, 2), HOURS_FMT),
            ("En yüksek (toplam saat)", round(stat_max, 2), HOURS_FMT),
            ("En düşük (toplam saat)", round(stat_min, 2), HOURS_FMT),
            ("Standart sapma (toplam saat)", round(stat_std, 2), HOURS_FMT),
            (f"Kritik eşiği aşan personel (aylık {format_tr(threshold_monthly, 0)} sa → bu dönem {format_tr(threshold, 1)} sa)", len(over_threshold), INT_FMT),
        ]
        sr = 3
        for i, (label, value, fmt) in enumerate(stats):
            r = sr + i
            wsa.cell(row=r, column=1, value=label).font = Font(bold=True)
            c = wsa.cell(row=r, column=2, value=value)
            c.number_format = fmt
            style_data_row(wsa, r, 1, 2, zebra=(i % 2 == 1))
        stats_last = sr + len(stats) - 1

        ph_sorted = sorted(person_stats.items(), key=lambda kv: kv[1]["toplam"], reverse=True)
        lh = stats_last + 2
        headers_p = ["Ad Soyad", "Görevi", "Normal (sa)", "Fazla (sa)", "Toplam (sa)", "Yevmiye Günü", "Eşik Durumu"]
        for c, h in enumerate(headers_p, start=1):
            wsa.cell(row=lh, column=c, value=h)
        style_header_row(wsa, lh, 1, len(headers_p))

        lf = lh + 1
        for i, (tc, st) in enumerate(ph_sorted):
            r = lf + i
            rec = all_records.get(tc, {})
            over = st["fazla"] > threshold
            wsa.cell(row=r, column=1, value=rec.get("name") or tc)
            wsa.cell(row=r, column=2, value=rec.get("role", ""))
            wsa.cell(row=r, column=3, value=round(st["normal"], 2)).number_format = HOURS_FMT
            wsa.cell(row=r, column=4, value=round(st["fazla"], 2)).number_format = HOURS_FMT
            wsa.cell(row=r, column=5, value=round(st["toplam"], 2)).number_format = HOURS_FMT
            wsa.cell(row=r, column=6, value=round(st["yevmiye"], 1)).number_format = HOURS_FMT
            wsa.cell(row=r, column=7, value="FAZLA MESAİ EŞİK ÜSTÜ" if over else "")
            style_data_row(wsa, r, 1, len(headers_p), zebra=(i % 2 == 1), fill=WARN_FILL if over else None)
        ll = lf + len(ph_sorted) - 1

        if ll >= lf:
            wsa.conditional_formatting.add(f"E{lf}:E{ll}",
                                            DataBarRule(start_type="min", end_type="max", color="5B9BD5"))
            active_sorted = [(tc, s) for tc, s in ph_sorted if s["toplam"] > 0][:20]
            if active_sorted:
                gh = lh
                gcol = 9
                wsa.cell(row=gh, column=gcol, value="Ad Soyad")
                wsa.cell(row=gh, column=gcol + 1, value="Normal")
                wsa.cell(row=gh, column=gcol + 2, value="Fazla")
                for k, (tc, st) in enumerate(active_sorted):
                    wsa.cell(row=gh + 1 + k, column=gcol, value=(all_records.get(tc, {}) or {}).get("name") or tc)
                    wsa.cell(row=gh + 1 + k, column=gcol + 1, value=round(st["normal"], 2))
                    wsa.cell(row=gh + 1 + k, column=gcol + 2, value=round(st["fazla"], 2))
                wsa.column_dimensions[get_column_letter(gcol)].width = 26

                pchart = BarChart()
                pchart.type = "bar"
                pchart.grouping = "stacked"
                pchart.overlap = 100
                pchart.title = f"En Çok Çalışan {len(active_sorted)} Personel (Normal + Fazla)"
                pchart.x_axis.title = "Saat"
                pchart.height = max(8, 0.6 * len(active_sorted) + 3)
                pchart.width = 19
                pchart.visible_cells_only = False
                pchart.add_data(Reference(wsa, min_col=gcol + 1, max_col=gcol + 2,
                                           min_row=gh, max_row=gh + len(active_sorted)), titles_from_data=True)
                pchart.set_categories(Reference(wsa, min_col=gcol, max_col=gcol,
                                                 min_row=gh + 1, max_row=gh + len(active_sorted)))
                wsa.add_chart(pchart, "L3")

        for col, w in zip("ABCDEFG", (28, 26, 13, 13, 14, 13, 20)):
            wsa.column_dimensions[col].width = w
        wsa.freeze_panes = f"A{lf}"
        setup_print(wsa)

        # ==============================================================
        # GÖREV BAZLI DAĞILIM
        # ==============================================================
        role_normal = {}
        role_fazla = {}
        role_yevmiye = {}
        role_counts = {}
        for tc, st in person_stats.items():
            role = (all_records.get(tc, {}) or {}).get("role") or "(Görev belirtilmemiş)"
            role_normal[role] = role_normal.get(role, 0.0) + st["normal"]
            role_fazla[role] = role_fazla.get(role, 0.0) + st["fazla"]
            role_yevmiye[role] = role_yevmiye.get(role, 0) + st["yevmiye"]
            role_counts[role] = role_counts.get(role, 0) + 1

        wsr = wb.create_sheet("Görev Bazlı")
        wsr.sheet_view.showGridLines = False
        banner(wsr, "A1:F1", f"{scope_str} {ay_adi(data['date_start'])} AYI GÖREV BAZLI MESAİ DAĞILIMI")

        hr = 3
        headers_r = ["Görevi", "Personel Sayısı", "Normal (sa)", "Fazla (sa)", "Toplam (sa)",
                     "Yevmiye Günü", "Kişi Başı Ort. (sa)"]
        for c, h in enumerate(headers_r, start=1):
            wsr.cell(row=hr, column=c, value=h)
        style_header_row(wsr, hr, 1, len(headers_r))

        roles_sorted = sorted(role_counts.keys(), key=lambda rl: role_normal[rl] + role_fazla[rl], reverse=True)
        rf = hr + 1
        for i, role in enumerate(roles_sorted):
            r = rf + i
            cnt = role_counts[role]
            nt = role_normal[role]
            ft = role_fazla[role]
            tot = nt + ft
            wsr.cell(row=r, column=1, value=role)
            wsr.cell(row=r, column=2, value=cnt).number_format = INT_FMT
            wsr.cell(row=r, column=3, value=round(nt, 2)).number_format = HOURS_FMT
            wsr.cell(row=r, column=4, value=round(ft, 2)).number_format = HOURS_FMT
            wsr.cell(row=r, column=5, value=round(tot, 2)).number_format = HOURS_FMT
            wsr.cell(row=r, column=6, value=round(role_yevmiye[role], 1)).number_format = HOURS_FMT
            wsr.cell(row=r, column=7, value=round(tot / cnt if cnt else 0, 2)).number_format = HOURS_FMT
            style_data_row(wsr, r, 1, len(headers_r), zebra=(i % 2 == 1))
        rl_ = rf + len(roles_sorted) - 1

        rtot = rl_ + 1
        wsr.cell(row=rtot, column=1, value="GENEL TOPLAM")
        wsr.cell(row=rtot, column=2, value=grand_employees).number_format = INT_FMT
        wsr.cell(row=rtot, column=3, value=round(grand_normal, 2)).number_format = HOURS_FMT
        wsr.cell(row=rtot, column=4, value=round(grand_fazla, 2)).number_format = HOURS_FMT
        wsr.cell(row=rtot, column=5, value=round(grand_total, 2)).number_format = HOURS_FMT
        wsr.cell(row=rtot, column=6, value=round(grand_yevmiye, 1)).number_format = HOURS_FMT
        wsr.cell(row=rtot, column=7, value=round(stat_avg, 2)).number_format = HOURS_FMT
        style_total_row(wsr, rtot, 1, len(headers_r))

        if rl_ >= rf:
            wsr.conditional_formatting.add(f"E{rf}:E{rl_}",
                                            DataBarRule(start_type="min", end_type="max", color="7E57C2"))
            rchart = BarChart()
            rchart.type = "bar"
            rchart.grouping = "stacked"
            rchart.overlap = 100
            rchart.title = "Göreve Göre Normal + Fazla Mesai"
            rchart.x_axis.title = "Saat"
            rchart.height = max(7, min(20, 0.5 * len(roles_sorted) + 3))
            rchart.width = 20
            rchart.visible_cells_only = False
            rchart.add_data(Reference(wsr, min_col=3, max_col=4, min_row=hr, max_row=rl_), titles_from_data=True)
            rchart.set_categories(Reference(wsr, min_col=1, max_col=1, min_row=rf, max_row=rl_))
            wsr.add_chart(rchart, "I3")

        for col, w in zip("ABCDEFG", (32, 16, 14, 14, 14, 13, 16)):
            wsr.column_dimensions[col].width = w
        setup_print(wsr)

    # ==================================================================
    # PERSONEL LİSTESİ (opsiyonel, günlük kırılımlı - toplam mesai)
    # ==================================================================
    if data["employee_records"]:
        ws4 = wb.create_sheet("Personel Listesi")
        ws4.sheet_view.showGridLines = False

        dates = data["dates_in_range"]
        base_cols = 11
        last4 = base_cols + len(dates)
        banner(ws4, f"A1:{get_column_letter(last4)}1",
               f"{scope_str} — PERSONEL LİSTESİ VE GÜNLÜK MESAİ ({date_range_str})")
        ws4.cell(row=2, column=1,
                 value=f"Kritik fazla mesai eşiğini (aylık {format_tr(threshold_monthly, 0)} saat, bu dönem için "
                       f"{format_tr(threshold, 1)} saat) aşan personel satırları kırmızı "
                       f"işaretlenmiştir. Günlük sütunlar Normal+Fazla toplamını gösterir. Doluluk Oranı, "
                       f"yevmiye günü sayısının kişinin bu dönemde aktif olduğu gün sayısına oranıdır "
                       f"(ay içi işe giriş/çıkış yapanlar için düşük çıkması normaldir).").font = SUBTITLE_FONT
        ws4.merge_cells(f"A2:{get_column_letter(last4)}2")

        h4 = 4
        headers4 = (["Ekip", "Bağlı Olduğu Ekip", "Ad Soyad", "TC Kimlik No", "Görevi",
                     "Normal Mesai (sa)", "Fazla Mesai (sa)", "Toplam Mesai (sa)", "Yevmiye Günü",
                     "Doluluk Oranı (%)", "Durum"] +
                    [f"{d:%d.%m}\n{GUN_ADLARI[d.weekday()]}" for d in dates])
        for c, h in enumerate(headers4, start=1):
            ws4.cell(row=h4, column=c, value=h)
        style_header_row(ws4, h4, 1, last4)

        records = sorted(data["employee_records"].items(),
                          key=lambda kv: (-(employee_normal_totals.get(kv[0], 0.0) + employee_fazla_totals.get(kv[0], 0.0)),
                                          kv[1]["name"]))
        f4 = h4 + 1
        for i, (tc, rec) in enumerate(records):
            r = f4 + i
            normal_h = employee_normal_totals.get(tc, 0.0)
            fazla_h = employee_fazla_totals.get(tc, 0.0)
            over = fazla_h > threshold
            durum = data["employee_durum"].get(tc, "Tam Dönem")
            is_kismi = durum != "Tam Dönem"
            ws4.cell(row=r, column=1, value=rec["team"])
            ws4.cell(row=r, column=2, value=rec["subteam"])
            ws4.cell(row=r, column=3, value=rec["name"])
            ws4.cell(row=r, column=4, value=tc)
            ws4.cell(row=r, column=5, value=rec["role"])
            ws4.cell(row=r, column=6, value=round(normal_h, 2)).number_format = HOURS_FMT
            ws4.cell(row=r, column=7, value=round(fazla_h, 2)).number_format = HOURS_FMT
            tcell = ws4.cell(row=r, column=8, value=round(normal_h + fazla_h, 2))
            tcell.number_format = HOURS_FMT
            tcell.font = Font(bold=True)
            ws4.cell(row=r, column=9, value=round(employee_yevmiye_days.get(tc, 0.0), 1)).number_format = HOURS_FMT
            ws4.cell(row=r, column=10, value=data["employee_doluluk"].get(tc, 0.0) / 100).number_format = '0.0%'
            ws4.cell(row=r, column=11, value=durum)
            normal_day_map = data["employee_day_normal"].get(tc, {})
            fazla_day_map = data["employee_day_totals"].get(tc, {})
            for j, d in enumerate(dates):
                v = normal_day_map.get(d, 0.0) + fazla_day_map.get(d, 0.0)
                ws4.cell(row=r, column=base_cols + 1 + j, value=round(v, 2) if v else None).number_format = HOURS_FMT
            fill = WARN_FILL if over else (WEEKEND_FILL if is_kismi else None)
            style_data_row(ws4, r, 1, last4, zebra=(i % 2 == 1), fill=fill)
        l4 = f4 + len(records) - 1

        ttl4 = l4 + 1
        ws4.cell(row=ttl4, column=1, value="TOPLAM")
        ws4.cell(row=ttl4, column=6, value=round(grand_normal, 2)).number_format = HOURS_FMT
        ws4.cell(row=ttl4, column=7, value=round(grand_fazla, 2)).number_format = HOURS_FMT
        ws4.cell(row=ttl4, column=8, value=round(grand_total, 2)).number_format = HOURS_FMT
        ws4.cell(row=ttl4, column=9, value=round(grand_yevmiye, 1)).number_format = HOURS_FMT
        for j, d in enumerate(dates):
            combined = data["day_totals_normal"].get(d, 0.0) + data["day_totals"].get(d, 0.0)
            ws4.cell(row=ttl4, column=base_cols + 1 + j, value=round(combined, 2)).number_format = HOURS_FMT
        style_total_row(ws4, ttl4, 1, last4)

        if l4 >= f4:
            ws4.conditional_formatting.add(f"H{f4}:H{l4}",
                                            DataBarRule(start_type="min", end_type="max", color="5B9BD5"))
            ws4.conditional_formatting.add(
                f"{get_column_letter(base_cols + 1)}{f4}:{get_column_letter(last4)}{l4}",
                ColorScaleRule(start_type="min", start_color="FFFFFF",
                                mid_type="percentile", mid_value=50, mid_color="FFF59D",
                                end_type="max", end_color="EF9A9A"))

        for col, w in zip("ABCDEFGHIJK", (20, 20, 26, 15, 26, 14, 14, 14, 12, 14, 24)):
            ws4.column_dimensions[col].width = w
        for j in range(len(dates)):
            ws4.column_dimensions[get_column_letter(base_cols + 1 + j)].width = 8
        ws4.freeze_panes = f"F{f4}"
        setup_print(ws4)

    # ==================================================================
    # VERİ KALİTESİ KONTROLÜ (Kademe 2)
    # ==================================================================
    tc_issues = data.get("tc_issues") or []
    dataset_mismatches = data.get("dataset_mismatches") or []
    duplicate_tc_names = data.get("duplicate_tc_names") or {}
    missing_subteam = data.get("missing_subteam_count", 0)
    missing_role = data.get("missing_role_count", 0)
    kismi_ay_list = [(tc, all_records.get(tc, {}).get("name", tc), durum)
                      for tc, durum in data["employee_durum"].items() if durum != "Tam Dönem"]

    if tc_issues or dataset_mismatches or duplicate_tc_names or missing_subteam or missing_role or kismi_ay_list:
        wsq = wb.create_sheet("Veri Kalitesi Kontrolü")
        wsq.sheet_view.showGridLines = False
        banner(wsq, "A1:D1", f"{scope_str} — VERİ KALİTESİ KONTROLÜ", height=28)
        wsq.cell(row=2, column=1,
                 value="Bu sayfa kaynak dosyadaki eksik/hatalı veya Veri Seti ile çelişen kayıtları listeler. "
                       "'Aynı TC, Farklı İsim' bulguları gerçek kişi sayısını ve o kişilerin saatlerini "
                       "etkiler (iki kişi tek kayıt gibi toplanır) - MUTLAKA incelenmelidir.").font = SUBTITLE_FONT
        wsq.merge_cells("A2:D2")

        row_cursor = 4
        summary_rows = [
            ("Aynı TC, farklı isimde kullanılıyor (gerçek kişi çakışması)", len(duplicate_tc_names)),
            ("Alt ekip bilgisi eksik personel", missing_subteam),
            ("Görev bilgisi eksik personel", missing_role),
            ("TC format/checksum sorunlu kayıt", len(tc_issues)),
            ("Veri Seti ile çelişen (düzeltilen) alan", len(dataset_mismatches)),
            ("Kısmi dönemde çalışan (ay içi giriş/çıkış) personel", len(kismi_ay_list)),
        ]
        for label, value in summary_rows:
            wsq.cell(row=row_cursor, column=1, value=label).font = Font(bold=True)
            c = wsq.cell(row=row_cursor, column=2, value=value)
            c.number_format = INT_FMT
            if value:
                c.font = Font(bold=True, color="B71C1C")
            row_cursor += 1
        row_cursor += 1

        if duplicate_tc_names:
            wsq.cell(row=row_cursor, column=1, value="Aynı TC, Farklı İsim (GERÇEK KİŞİ ÇAKIŞMASI)").font = Font(bold=True, size=12, color="B71C1C")
            row_cursor += 1
            for c, h in enumerate(["TC Kimlik No", "Kayıtlı İsimler"], start=1):
                wsq.cell(row=row_cursor, column=c, value=h)
            style_header_row(wsq, row_cursor, 1, 2)
            row_cursor += 1
            for tc, names in duplicate_tc_names.items():
                wsq.cell(row=row_cursor, column=1, value=tc)
                wsq.cell(row=row_cursor, column=2, value=" / ".join(names))
                style_data_row(wsq, row_cursor, 1, 2, fill=WARN_FILL)
                row_cursor += 1
            row_cursor += 1

        if tc_issues:
            wsq.cell(row=row_cursor, column=1, value="TC Format/Checksum Sorunları").font = Font(bold=True, size=12)
            row_cursor += 1
            for c, h in enumerate(["Satır", "TC", "Ad Soyad", "Sorun"], start=1):
                wsq.cell(row=row_cursor, column=c, value=h)
            style_header_row(wsq, row_cursor, 1, 4)
            row_cursor += 1
            for issue in tc_issues[:100]:
                wsq.cell(row=row_cursor, column=1, value=issue["row"])
                wsq.cell(row=row_cursor, column=2, value=issue["tc"])
                wsq.cell(row=row_cursor, column=3, value=issue["name"])
                wsq.cell(row=row_cursor, column=4, value=issue["reason"])
                style_data_row(wsq, row_cursor, 1, 4, fill=WARN_FILL)
                row_cursor += 1
            row_cursor += 1

        if dataset_mismatches:
            wsq.cell(row=row_cursor, column=1, value="Veri Seti İle Çelişen / Düzeltilen Alanlar").font = Font(bold=True, size=12)
            row_cursor += 1
            for c, h in enumerate(["TC", "Ad Soyad", "Alan", "Kaynaktaki Değer", "Veri Setindeki Değer"], start=1):
                wsq.cell(row=row_cursor, column=c, value=h)
            style_header_row(wsq, row_cursor, 1, 5)
            row_cursor += 1
            for tc, name, field, old_val, new_val in dataset_mismatches[:200]:
                wsq.cell(row=row_cursor, column=1, value=tc)
                wsq.cell(row=row_cursor, column=2, value=name)
                wsq.cell(row=row_cursor, column=3, value=field)
                wsq.cell(row=row_cursor, column=4, value=old_val)
                wsq.cell(row=row_cursor, column=5, value=new_val)
                style_data_row(wsq, row_cursor, 1, 5, fill=WEEKEND_FILL)
                row_cursor += 1
            row_cursor += 1

        if kismi_ay_list:
            wsq.cell(row=row_cursor, column=1, value="Kısmi Dönem Çalışanları").font = Font(bold=True, size=12)
            row_cursor += 1
            for c, h in enumerate(["TC", "Ad Soyad", "Durum"], start=1):
                wsq.cell(row=row_cursor, column=c, value=h)
            style_header_row(wsq, row_cursor, 1, 3)
            row_cursor += 1
            for tc, name, durum in sorted(kismi_ay_list, key=lambda x: x[1])[:200]:
                wsq.cell(row=row_cursor, column=1, value=tc)
                wsq.cell(row=row_cursor, column=2, value=name)
                wsq.cell(row=row_cursor, column=3, value=durum)
                style_data_row(wsq, row_cursor, 1, 3)
                row_cursor += 1

        for col, w in zip("ABCDE", (24, 28, 22, 26, 26)):
            wsq.column_dimensions[col].width = w
        setup_print(wsq, landscape=False)

    # ==================================================================
    # PERSONEL DEĞİŞİKLİKLERİ (Kademe 4 - Veri Seti geçmişi, çağıran taraf sağlar)
    # ==================================================================
    if history_changes:
        wsh = wb.create_sheet("Personel Değişiklikleri")
        wsh.sheet_view.showGridLines = False
        banner(wsh, "A1:E1", f"{scope_str} — PERSONEL DEĞİŞİKLİKLERİ (Veri Seti Geçmişi)", height=28)
        wsh.cell(row=2, column=1,
                 value="Veri Setinde kayıtlı, bu rapor kapsamındaki personelin geçmiş bilgi değişiklikleri.").font = SUBTITLE_FONT
        wsh.merge_cells("A2:E2")

        hh = 4
        for c, h in enumerate(["Ad Soyad", "TC", "Alan", "Önceki Değer", "Yeni Değer", "Tarih", "Kaynak"], start=1):
            wsh.cell(row=hh, column=c, value=h)
        style_header_row(wsh, hh, 1, 7)
        fh = hh + 1
        for i, ch in enumerate(history_changes[:500]):
            r = fh + i
            wsh.cell(row=r, column=1, value=ch.get("name") or ch.get("tc"))
            wsh.cell(row=r, column=2, value=ch.get("tc"))
            wsh.cell(row=r, column=3, value=FIELD_LABEL_TR.get(ch.get("field"), ch.get("field")))
            wsh.cell(row=r, column=4, value=ch.get("old_value"))
            wsh.cell(row=r, column=5, value=ch.get("new_value"))
            wsh.cell(row=r, column=6, value=ch.get("changed_at"))
            wsh.cell(row=r, column=7, value=ch.get("source"))
            style_data_row(wsh, r, 1, 7, zebra=(i % 2 == 1))
        for col, w in zip("ABCDEFG", (26, 15, 20, 22, 22, 18, 26)):
            wsh.column_dimensions[col].width = w
        wsh.freeze_panes = f"A{fh}"
        setup_print(wsh)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


# ==================================================================
# EKİP LİSTESİ (personel listesi - tarih bağımsız)
# ==================================================================

# Öncelik sırası: FORMEN > EKİP BAŞI > (diğer görevler, personel sayısına göre) > BAYRAKÇI
# Gerçek veride "EKİPBAŞI / TOPRAK İŞLERİ" gibi ek metinli varyantlar da görülüyor,
# bu yüzden boşluksuz normalize edilmiş metin üzerinde "içerir mi" kontrolü yapılıyor.
_FORMEN_KEY = "FORMEN"
_EKIPBASI_KEY = "EKIPBASI"
_BAYRAKCI_KEY = "BAYRAKCI"


def _role_norm_nospace(role):
    return engine.normalize_name(role).replace(" ", "")


def build_employee_roster(cfg, source_path, team_filter=None, subteam_filter=None, role_filter=None,
                           exclude_exited=False):
    """Tarih bağımsız personel listesi: seçili grup/alt ekip/meslek için
    kaynaktaki NORMAL MESAİ satırlarından benzersiz personelleri (Ad Soyad,
    TC, Görevi, Ekip, Alt Ekip) toplar.
    exclude_exited=True ise, İşten Çıkış Tarihi dolu olan (ayrılmış) personel
    listeye dahil edilmez - SADECE Ekip Listesi (generate_roster) için kullanılır;
    cost_report ve personnel_db'nin ayrılan personeli de görebilmesi gerektiğinden
    varsayılan olarak False'tur."""
    wb = openpyxl.load_workbook(source_path, data_only=True)
    sheet_name, warning = engine.resolve_sheet(wb, cfg["source_sheet"], "Kaynak")
    ws = wb[sheet_name]

    group_col = column_index_from_string(cfg["source_group_col"])
    subteam_col = column_index_from_string(cfg["source_subteam_col"])
    role_col = column_index_from_string(cfg["source_role_col"])
    name_col = column_index_from_string(cfg["source_name_col"])
    type_col = column_index_from_string(cfg["source_type_col"])
    tc_col = column_index_from_string(cfg["source_tc_col"])
    cikis_col = column_index_from_string(cfg.get("source_cikis_col", "P"))

    team_set = {t.strip() for t in team_filter} if team_filter else None
    subteam_set = {t.strip() for t in subteam_filter} if subteam_filter else None
    role_set = {t.strip() for t in role_filter} if role_filter else None

    records = {}
    excluded_exited = 0
    for r in range(cfg["source_data_start_row"], _effective_data_end_row(cfg, ws) + 1):
        tip = ws.cell(row=r, column=type_col).value
        if not tip or "NORMAL" not in str(tip).upper():
            continue

        team_raw = ws.cell(row=r, column=group_col).value
        if not team_raw:
            continue
        team = str(team_raw).strip()
        if team_set and team not in team_set:
            continue

        subteam_raw = ws.cell(row=r, column=subteam_col).value
        subteam = str(subteam_raw).strip() if subteam_raw else NO_SUBTEAM
        if subteam_set and subteam not in subteam_set:
            continue

        role_raw = ws.cell(row=r, column=role_col).value
        role = str(role_raw).strip() if role_raw else NO_ROLE
        if role_set and role not in role_set:
            continue

        tc = engine.normalize_tc(ws.cell(row=r, column=tc_col).value)
        if not tc or tc in records:
            continue

        # İşten çıkış tarihi doluysa (personel ayrılmış), istenirse ekip listesine dahil etme
        if exclude_exited and _parse_date_cell(ws.cell(row=r, column=cikis_col).value):
            excluded_exited += 1
            continue

        name = str(ws.cell(row=r, column=name_col).value or "").strip()
        records[tc] = {"tc": tc, "name": name, "role": role, "team": team, "subteam": subteam}

    return {
        "records": list(records.values()),
        "warning": warning,
        "excluded_exited": excluded_exited,
        "team_filter": sorted(team_set) if team_set else None,
        "subteam_filter": sorted(subteam_set) if subteam_set else None,
        "role_filter": sorted(role_set) if role_set else None,
    }


def sort_roster(records):
    """İstenen öncelik sırasına göre personel listesini sıralar:
    1) FORMEN, 2) EKİP BAŞI, 3) diğer görevler (personel sayısına göre çoktan aza),
    4) BAYRAKÇI (her zaman en sonda)."""
    role_counts = {}
    for rec in records:
        role_counts[rec["role"]] = role_counts.get(rec["role"], 0) + 1

    def role_rank(role):
        key = _role_norm_nospace(role)
        if _FORMEN_KEY in key:
            return (0, 0)
        if _EKIPBASI_KEY in key:
            return (0, 1)
        if _BAYRAKCI_KEY in key:
            return (2, 0)
        return (1, -role_counts.get(role, 0))

    def sort_key(rec):
        role = rec["role"]
        return (role_rank(role), role, engine.normalize_name(rec["name"]), rec["name"])

    return sorted(records, key=sort_key), role_counts


def write_roster_report(cfg, roster_data, output_path, scope_str):
    records_sorted, role_counts = sort_roster(roster_data["records"])
    if not records_sorted:
        raise engine.TransferError("Seçilen filtrelerle eşleşen personel bulunamadı.")

    wb = openpyxl.Workbook()

    # ---------------------------------------------------------- Personel Listesi
    ws = wb.active
    ws.title = "Personel Listesi"
    ws.sheet_view.showGridLines = False

    last_col = 4
    banner(ws, f"A1:{get_column_letter(last_col)}1", f"{scope_str} EKİP LİSTESİ", height=28)
    sub_text = f"Toplam {len(records_sorted)} personel"
    if roster_data.get("excluded_exited"):
        sub_text += f"   ·   {roster_data['excluded_exited']} ayrılmış personel listeye dahil edilmedi"
    sub = ws.cell(row=2, column=1, value=sub_text)
    sub.font = SUBTITLE_FONT
    ws.merge_cells(f"A2:{get_column_letter(last_col)}2")

    header_row = 4
    headers = ["Sıra No", "Ad Soyad", "TC Kimlik No", "Görevi"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    style_header_row(ws, header_row, 1, last_col)

    ROLE_BANNER_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    ROLE_BANNER_FONT = Font(bold=True, color=NAVY_DARK, size=10)

    r = header_row + 1
    sira = 1
    prev_role = None
    for rec in records_sorted:
        if rec["role"] != prev_role:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell = ws.cell(row=r, column=1,
                            value=f"{rec['role']}  ({role_counts[rec['role']]} kişi)")
            cell.font = ROLE_BANNER_FONT
            cell.fill = ROLE_BANNER_FILL
            cell.alignment = Alignment(vertical="center", indent=1)
            for c in range(1, last_col + 1):
                ws.cell(row=r, column=c).border = THIN_BORDER
                ws.cell(row=r, column=c).fill = ROLE_BANNER_FILL
            prev_role = rec["role"]
            r += 1

        ws.cell(row=r, column=1, value=sira).number_format = INT_FMT
        ws.cell(row=r, column=2, value=rec["name"])
        ws.cell(row=r, column=3, value=rec["tc"])
        ws.cell(row=r, column=4, value=rec["role"])
        style_data_row(ws, r, 1, last_col, zebra=(sira % 2 == 0))
        sira += 1
        r += 1

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 34
    ws.freeze_panes = f"A{header_row + 1}"
    setup_print(ws, landscape=False)

    # ---------------------------------------------------------- Görev Özeti
    ws2 = wb.create_sheet("Görev Özeti")
    ws2.sheet_view.showGridLines = False
    banner(ws2, "A1:B1", f"{scope_str} — GÖREVE GÖRE PERSONEL SAYISI", height=26)

    h2 = 3
    ws2.cell(row=h2, column=1, value="Görevi")
    ws2.cell(row=h2, column=2, value="Personel Sayısı")
    style_header_row(ws2, h2, 1, 2)

    # özet tabloda da aynı öncelik sırasını koru (görev bazında tekilleştirip aynı sort_key mantığını uygula)
    seen = []
    seen_roles = set()
    for rec in records_sorted:
        if rec["role"] not in seen_roles:
            seen.append(rec["role"])
            seen_roles.add(rec["role"])

    f2 = h2 + 1
    for i, role in enumerate(seen):
        r = f2 + i
        ws2.cell(row=r, column=1, value=role)
        ws2.cell(row=r, column=2, value=role_counts[role]).number_format = INT_FMT
        style_data_row(ws2, r, 1, 2, zebra=(i % 2 == 1))
    l2 = f2 + len(seen) - 1

    tot2 = l2 + 1
    ws2.cell(row=tot2, column=1, value="TOPLAM")
    ws2.cell(row=tot2, column=2, value=len(records_sorted)).number_format = INT_FMT
    style_total_row(ws2, tot2, 1, 2)

    ws2.column_dimensions["A"].width = 34
    ws2.column_dimensions["B"].width = 18
    setup_print(ws2, landscape=False)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def _slugify(text):
    """Dosya adı için güvenli kısaltma: Türkçe karakterleri sadeleştirir,
    boşlukları tamamen kaldırır, küçük harfe çevirir. 'KM 28' -> 'km28'."""
    ascii_text = engine.normalize_name(text).replace(" ", "")
    return ascii_text.lower() or "ekip"


def generate_roster(cfg, source_path=None, team_filter=None, subteam_filter=None,
                     role_filter=None, output_name=None):
    source_path = source_path or engine.find_single_xlsx(SOURCE_DIR, "SOURCE")
    cfg = dict(cfg)
    cfg["source_path"] = str(source_path)

    roster_data = build_employee_roster(cfg, source_path, team_filter, subteam_filter, role_filter,
                                         exclude_exited=True)

    # Başlık ve dosya adı SADECE ekip/alt ekip seçimini yansıtır (meslek filtresi
    # başlığa dahil edilmez - "tümünü seç" ile tüm meslekler işaretlense bile
    # başlık sade kalır, örn. "KM 28 EKİP LİSTESİ").
    if roster_data["subteam_filter"]:
        scope_name = " / ".join(roster_data["subteam_filter"])
    elif roster_data["team_filter"]:
        scope_name = " / ".join(roster_data["team_filter"])
    else:
        scope_name = "TÜM GRUPLAR"

    # Ayrıca meslek filtresi de seçiliyse (ve tüm meslekler değil, belirli birkaçıysa)
    # kısa bir not olarak ekle - başlığı kalabalıklaştırmadan bilgi vermek için.
    scope_str = scope_name

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output_name:
        filename = output_name
    else:
        filename = f"{_slugify(scope_name)}_ekip_listesi.xlsx"
    output_path = OUTPUT_DIR / filename
    saved_path = write_roster_report(cfg, roster_data, output_path, scope_str)
    return saved_path, roster_data, scope_str


# ---------------------------------------------------------------- giriş noktaları

def generate_report(cfg, source_path=None, date_start=None, date_end=None, period_days=7,
                     team_filter=None, subteam_filter=None, role_filter=None, include_employee_list=False,
                     output_name=None, threshold=None, compare_previous=True, tc_filter=None,
                     dataset_lookup=None, history_changes=None):
    source_path = source_path or engine.find_single_xlsx(SOURCE_DIR, "SOURCE")
    cfg = dict(cfg)
    cfg["source_path"] = str(source_path)
    if threshold is not None:
        cfg["overtime_threshold"] = float(threshold)

    if date_start is None or date_end is None:
        opts = scan_filter_options(cfg, source_path)
        date_start = date_start or opts["date_min"]
        date_end = date_end or opts["date_max"]

    data = build_report_data(cfg, source_path, date_start, date_end, period_days,
                              team_filter=team_filter, subteam_filter=subteam_filter,
                              role_filter=role_filter,
                              include_employee_list=include_employee_list,
                              compare_previous=compare_previous, tc_filter=tc_filter,
                              dataset_lookup=dataset_lookup)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / (output_name or cfg.get("output_name", "Ekip_Mesai_Raporu.xlsx"))
    return write_report(cfg, data, output_path, history_changes=history_changes)


if __name__ == "__main__":
    out = generate_report(CONFIG, period_days=7, include_employee_list=True)
    print(f"Rapor oluşturuldu: {out}")
