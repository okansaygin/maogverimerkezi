# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - PERSONEL MALİYET RAPORU (veri_merkezi_maliyet.py)
====================================================================
Puantaj Suite'teki "Personel Maliyet Raporu"nun (cost_report.py) Veri
Merkezi karşılığı.

FARK: Suite personel kapsamını kaynak puantaj Excel'inden okur. Bu modül
kapsamı KALICI PERSONEL KAYDINDAN (personnel tablosu) alır; hakediş
dosyası yine Excel'dir (TC ile eşleştirilir).

MOTOR TEKRARLANMAZ: hakediş sayfası tespiti, başlık->alan eşleştirmesi,
hakediş okuma ve Excel çıktısı doğrudan cost_report'tan çağrılır. Sadece
personel listesinin nereden geldiği değişir. cost_report.py'ye DOKUNULMAZ.
"""

__version__ = "2026-09-24.1"

from pathlib import Path

import openpyxl

import cost_report as cost
import personnel_db as pdb
import team_report as report
import veri_merkezi_sema as sema

NO_GRUP = "(Grup belirtilmemiş)"


def sayfalari_tara(hakedis_path):
    """Her sayfa için: bordro sayfası mı, kaç kişi var, sütun düzeni ilk bordro
    sayfasından farklı mı. Döner: [{"sayfa", "bordro", "kisi", "kayik"}]"""
    cfg = cost.CONFIG
    wb = openpyxl.load_workbook(hakedis_path, data_only=True)
    bordro = cost.find_payroll_sheets(wb, cfg["hakedis_header_row"], cfg["hakedis_search_rows"])
    _rec, per_sheet, _dup = cost.read_hakedis_records(cfg, hakedis_path, bordro)
    ilk_map = None
    sonuc = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        if sn in bordro:
            cmap = cost.build_column_map(ws, cfg["hakedis_header_row"], cfg["hakedis_search_rows"])
            kayik = False
            if ilk_map is None:
                ilk_map = cmap
            else:
                ortak = set(cmap) & set(ilk_map)
                kayik = any(cmap[k] != ilk_map[k] for k in ortak)
            sonuc.append({"sayfa": sn, "bordro": True, "kisi": per_sheet.get(sn, 0), "kayik": kayik})
        else:
            sonuc.append({"sayfa": sn, "bordro": False, "kisi": 0, "kayik": False})
    return sonuc


def _roster(gruplar=None, ekipler=None, gorevler=None, db_path=None):
    """cost_report'un beklediği kayıt biçiminde (team/subteam/name/tc/role) personel.
    Maliyet bir DÖNEME ait olduğu için ayrılmış personel de dahildir."""
    pdb.init_db(db_path)
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute("SELECT tc, ad_soyad, grup, bagli_oldugu_ekip, gorevi FROM personnel").fetchall()
    finally:
        conn.close()
    out = []
    for r in rows:
        rec = {"tc": r["tc"], "name": (r["ad_soyad"] or "").strip(),
               "team": (r["grup"] or "").strip() or NO_GRUP,
               "subteam": (r["bagli_oldugu_ekip"] or "").strip() or report.NO_SUBTEAM,
               "role": (r["gorevi"] or "").strip() or report.NO_ROLE}
        if gruplar and rec["team"] not in gruplar:
            continue
        if ekipler and rec["subteam"] not in ekipler:
            continue
        if gorevler and rec["role"] not in gorevler:
            continue
        out.append(rec)
    return out


def maliyet_verisi(hakedis_path, sayfalar, gruplar=None, ekipler=None, gorevler=None, db_path=None):
    """cost_report.build_cost_data ile AYNI çıktı sözlüğü (write_cost_report'a
    doğrudan verilebilir) + ekran için özet alanlar."""
    cfg = cost.CONFIG
    roster = _roster(gruplar, ekipler, gorevler, db_path)
    hakedis, per_sheet, dup = cost.read_hakedis_records(cfg, hakedis_path, sayfalar)
    ceiling = cfg.get("unit_rate_sanity_ceiling", 10000.0)
    alanlar = [k for k, _a, _n in cost.FIELD_RULES if k != "tc"]

    merged, not_found, flags_all = [], [], []
    for rec in roster:
        h = hakedis.get(rec["tc"])
        if not h:
            continue   # kapsamdaki herkes hakedişte olmak zorunda değil; aşağıda ayrıca sayılır
        row = dict(rec)
        for k in alanlar:
            row[k] = h.get(k)
        row["_hakedis_sheet"] = h.get("_sheet")
        flags = []
        pt = row.get("personele_odenecek_toplam") or 0
        if pt and not row.get("sgk_odemesi_tutar"):
            flags.append("SGK ödemesi hesaplanmamış")
        if pt and not row.get("firma_payi_tutar"):
            flags.append("Firma payı hesaplanmamış")
        yb = row.get("yevmiye_bedeli") or 0
        if isinstance(yb, (int, float)) and yb > ceiling:
            flags.append(f"Yevmiye bedeli anormal yüksek ({report.format_tr(yb)} TL)")
        row["_flags"] = flags
        if flags:
            flags_all.append((rec["tc"], rec["name"], flags))
        merged.append(row)

    # Hakedişte olup personel kaydında (kapsamda) karşılığı olmayanlar ve
    # kapsamda olup puantajı olan ama hakedişte bulunmayanlar
    kapsam_tc = {r["tc"] for r in roster}
    conn = sema.get_connection(db_path)
    try:
        puantajli = {r[0] for r in conn.execute("SELECT DISTINCT tc FROM gunluk_puantaj")}
    finally:
        conn.close()
    for rec in roster:
        if rec["tc"] not in hakedis and rec["tc"] in puantajli:
            not_found.append(rec)
    kayitsiz = [(tc, h.get("ad_soyad")) for tc, h in hakedis.items() if tc not in kapsam_tc] if not (gruplar or ekipler or gorevler) else []

    data = {
        "merged": merged, "not_found": not_found, "quality_flags": flags_all, "duplicate_tc": dup,
        "hakedis_sheet_counts": per_sheet, "warning": None,
        "team_filter": sorted(gruplar) if gruplar else None,
        "subteam_filter": sorted(ekipler) if ekipler else None,
        "role_filter": sorted(gorevler) if gorevler else None,
    }

    def s(recs, k):
        return sum((r.get(k) or 0) for r in recs if isinstance(r.get(k), (int, float)))

    kir = "subteam" if (gruplar and len(gruplar) == 1) else "team"
    gruplu = {}
    for r in merged:
        gruplu.setdefault(r[kir], []).append(r)
    ekip_ozet = sorted(({"ad": ad, "kisi": len(v), "personel": s(v, "personele_odenecek_toplam"),
                         "sgk": s(v, "sgk_odemesi_tutar"), "firma": s(v, "firma_payi_tutar"),
                         "toplam": s(v, "toplam_odenecek")} for ad, v in gruplu.items()), key=lambda x: -x["toplam"])
    gorevli = {}
    for r in merged:
        gorevli.setdefault(r["role"], []).append(r)
    gorev_ozet = sorted(({"ad": ad, "kisi": len(v), "toplam": s(v, "toplam_odenecek"),
                          "kisi_basi": s(v, "toplam_odenecek") / len(v)} for ad, v in gorevli.items()),
                        key=lambda x: -x["toplam"])
    flag_say = {}
    for _tc, _n, fl in flags_all:
        for f in fl:
            anahtar = "Yevmiye bedeli anormal" if f.startswith("Yevmiye") else f
            flag_say[anahtar] = flag_say.get(anahtar, 0) + 1
    data["ozet"] = {
        "toplam": s(merged, "toplam_odenecek"), "personel": s(merged, "personele_odenecek_toplam"),
        "sgk": s(merged, "sgk_odemesi_tutar"), "firma": s(merged, "firma_payi_tutar"),
        "fm_tutar": s(merged, "odenecek_fm_tutar"),
        "eslesen": len(merged), "kapsam": len(roster), "bulunamayan": len(not_found),
        "kayitsiz_hakedis": len(kayitsiz), "tekrar_tc": len(dup),
        "kirilim": "Alt ekip" if kir == "subteam" else "Grup",
        "ekipler": ekip_ozet, "gorevler": gorev_ozet, "flag_sayilari": flag_say, "esik": ceiling,
    }
    return data


def kapsam_adi(data):
    return cost._scope_str(data)


def excel_yaz(data, output_name=None):
    report.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    name = output_name or f"{report._slugify(kapsam_adi(data))}_personel_maliyet.xlsx"
    return cost.write_cost_report(cost.CONFIG, data, report.OUTPUT_DIR / name)
