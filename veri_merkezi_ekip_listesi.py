# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - EKİP LİSTESİ MOTORU (veri_merkezi_ekip_listesi.py)
=====================================================================
Puantaj Suite'teki "Ekip Listesi" modülünün Veri Merkezi karşılığı.

FARK: Suite her seferinde kaynak Excel'i yeniden okur
(team_report.build_employee_roster). Bu modül aynı listeyi ham Excel'e HİÇ
dokunmadan, kalıcı personel kaydından (personel_veritabani.db > personnel)
üretir.

MOTOR TEKRARLANMAZ: sıralama (Formen > Ekip Başı > diğer görevler > Bayrakçı)
ve Excel çıktısı doğrudan team_report.sort_roster / team_report.write_roster_report
ile yapılır. Böylece Suite ile Veri Merkezi'nden alınan liste birebir aynı
görünür ve bakım tek yerde yapılır. team_report.py'ye DOKUNULMAZ.

İş kuralları (Suite ile aynı):
  - İşten çıkış tarihi dolu olan personel listeye ALINMAZ.
  - Ekip Listesi tarih bağımsızdır (dönem/saat verisi kullanılmaz).
  - Grup / alt ekip / görev filtreleri bağımsız kümelerdir ve VE ile birleşir;
    boş seçim = o alanda filtre yok.

BİLİNÇLİ FARK: Grubu boş olan personel Suite'te (Excel motoru) sessizce
düşüyordu; burada "(Grup belirtilmemiş)" başlığı altında listelenir
(Veri Merkezi'nin "hiçbir şey sessizce kaybolmaz" ilkesi).
"""

__version__ = "2026-09-21.1"

import os
import subprocess
import sys

import personnel_db as pdb
import puantaj_engine as engine
import team_report as report
import veri_merkezi_sema as sema

NO_GRUP = "(Grup belirtilmemiş)"
NO_EKIP = report.NO_SUBTEAM
NO_GOREV = report.NO_ROLE


# ---------------------------------------------------------------- okuma

def _bos_ise(deger, etiket):
    deger = (deger or "").strip()
    return deger or etiket


def _tum_kayitlar(db_path=None):
    """personnel tablosundaki HERKES (aktif + ayrılmış), boş alanlar etiketlenmiş."""
    pdb.init_db(db_path)  # tablo yoksa sadece YAPIYI oluşturur
    conn = sema.get_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT tc, ad_soyad, grup, bagli_oldugu_ekip, gorevi, isten_cikis_tarihi FROM personnel"
        ).fetchall()
    finally:
        conn.close()
    return [{
        "tc": r["tc"],
        "name": (r["ad_soyad"] or "").strip(),
        "team": _bos_ise(r["grup"], NO_GRUP),
        "subteam": _bos_ise(r["bagli_oldugu_ekip"], NO_EKIP),
        "role": _bos_ise(r["gorevi"], NO_GOREV),
        "exited": bool((r["isten_cikis_tarihi"] or "").strip()),
    } for r in rows]


def _sirala_tr(degerler):
    return sorted(degerler, key=lambda x: (engine.normalize_name(x), x))


def get_hierarchy(db_path=None):
    """Filtre seçenekleri için {grup: {alt_ekip: [görevler]}} - SADECE aktif
    personelden (ayrılmış personelin görevi/ekibi seçenek olarak görünmez)."""
    agac = {}
    for k in _tum_kayitlar(db_path):
        if k["exited"]:
            continue
        agac.setdefault(k["team"], {}).setdefault(k["subteam"], set()).add(k["role"])
    return {
        g: {e: _sirala_tr(roller) for e, roller in sorted(ekipler.items(), key=lambda kv: engine.normalize_name(kv[0]))}
        for g, ekipler in sorted(agac.items(), key=lambda kv: engine.normalize_name(kv[0]))
    }


def build_roster(grup_filter=None, ekip_filter=None, gorev_filter=None, db_path=None):
    """team_report.build_employee_roster(exclude_exited=True) ile AYNI şekilde
    bir sözlük döner (records / warning / excluded_exited / *_filter), böylece
    team_report.sort_roster ve write_roster_report doğrudan kullanılabilir.
    excluded_exited: seçili kapsamdaki ayrılmış (listeye alınmayan) personel sayısı."""
    team_set = {t.strip() for t in grup_filter} if grup_filter else None
    subteam_set = {t.strip() for t in ekip_filter} if ekip_filter else None
    role_set = {t.strip() for t in gorev_filter} if gorev_filter else None

    records = []
    excluded_exited = 0
    for k in _tum_kayitlar(db_path):
        if team_set and k["team"] not in team_set:
            continue
        if subteam_set and k["subteam"] not in subteam_set:
            continue
        if role_set and k["role"] not in role_set:
            continue
        if k["exited"]:
            excluded_exited += 1
            continue
        records.append({"tc": k["tc"], "name": k["name"], "role": k["role"],
                        "team": k["team"], "subteam": k["subteam"]})

    return {
        "records": records,
        "warning": None,
        "excluded_exited": excluded_exited,
        "team_filter": sorted(team_set) if team_set else None,
        "subteam_filter": sorted(subteam_set) if subteam_set else None,
        "role_filter": sorted(role_set) if role_set else None,
    }


def sirala_ve_dagilim(records):
    """(sıralı_kayıtlar, [(görev, kişi_sayısı), ...]) - görev dağılımı da aynı
    öncelik sırasında (Excel'deki 'Görev Özeti' sayfasıyla aynı)."""
    sirali, sayilar = report.sort_roster(records)
    gorevler = []
    gorulen = set()
    for rec in sirali:
        if rec["role"] not in gorulen:
            gorulen.add(rec["role"])
            gorevler.append(rec["role"])
    return sirali, [(g, sayilar[g]) for g in gorevler]


def kapsam_adi(roster_data, en_fazla=3):
    """Başlık/dosya adı SADECE grup/alt ekip seçimini yansıtır (Suite ile aynı;
    görev filtresi başlığa girmez). 3'ten fazla seçimde '5 ALT EKİP' gibi kısaltılır
    ki başlık ve dosya adı taşmasın."""
    if roster_data["subteam_filter"]:
        liste, birim = roster_data["subteam_filter"], "ALT EKİP"
    elif roster_data["team_filter"]:
        liste, birim = roster_data["team_filter"], "GRUP"
    else:
        return "TÜM GRUPLAR"
    return " / ".join(liste) if len(liste) <= en_fazla else f"{len(liste)} {birim}"


def get_kalite_notu(db_path=None):
    """Listenin neden eksik görünebileceğini açıklayan AÇIK uyarı sayıları.
    Bu liste TC'si olan personelden üretilir (TC = birincil anahtar):
      - tcsiz_isim : TC'siz satırlarda kalan farklı isim sayısı (listede YOK)
      - cakisan_tc : birden fazla isimle gelen TC sayısı (listede TEK satır)"""
    sema.migrate(db_path)
    conn = sema.get_connection(db_path)
    try:
        tcsiz = conn.execute(
            "SELECT COUNT(DISTINCT ad_soyad) FROM quality_alerts "
            "WHERE cozuldu=0 AND tip='tc_eksik' AND ad_soyad IS NOT NULL AND ad_soyad<>''"
        ).fetchone()[0]
        cakisan = conn.execute(
            "SELECT COUNT(DISTINCT tc) FROM quality_alerts "
            "WHERE cozuldu=0 AND tip='ayni_tc_farkli_isim' AND tc IS NOT NULL"
        ).fetchone()[0]
    finally:
        conn.close()
    return {"tcsiz_isim": tcsiz, "cakisan_tc": cakisan}


# ---------------------------------------------------------------- Excel çıktısı

def excel_olustur(grup_filter=None, ekip_filter=None, gorev_filter=None, db_path=None, output_name=None):
    """Suite'teki generate_roster ile aynı çıktıyı (Personel Listesi + Görev Özeti
    sayfaları) OUTPUT klasörüne yazar. Döner: (dosya_yolu, roster_data, kapsam).
    Eşleşen personel yoksa engine.TransferError fırlatır."""
    roster = build_roster(grup_filter, ekip_filter, gorev_filter, db_path)
    kapsam = kapsam_adi(roster)
    dosya_adi = output_name or f"{report._slugify(kapsam)}_ekip_listesi.xlsx"
    yol = report.write_roster_report(report.CONFIG, roster, report.OUTPUT_DIR / dosya_adi, kapsam)
    return yol, roster, kapsam


def cikti_klasoru():
    report.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return report.OUTPUT_DIR


def open_path(path):
    """Dosya/klasörü işletim sisteminin varsayılan uygulamasıyla açar. Veri Merkezi
    yerel bir masaüstü uygulaması olduğundan (sunucu = kullanıcının kendi bilgisayarı)
    sunucu tarafında açmak doğrudur; pywebview'in indirme kısıtlarına da takılmaz."""
    path = str(path)
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])
