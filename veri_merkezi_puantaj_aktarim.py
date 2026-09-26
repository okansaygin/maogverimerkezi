# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - SAP PUANTAJ AKTARIM (veri_merkezi_puantaj_aktarim.py)
========================================================================
puantaj_suite.py'deki "Puantaj Aktarımı" modülünün (puantaj_engine.py) ince
bir sarmalayıcısı - MOTOR TEKRARLANMAZ, doğrudan engine.run_transfer()
çağrılır. Amaç: AYLIK kaynak excel'den, seçilen bir GÜN için, o günün
normal+fazla toplamını TC eşleştirerek GÜNLÜK hedef roster excel(ler)ine
yazmak (Çalışma Durumu ve açıklama sütunları dahil).

ÖNEMLİ - bu modül Veri Merkezi'nin kalıcı deposunu (personel_veritabani.db)
OKUMAZ/YAZMAZ: bu, kimlik/mesai içe aktarımından TAMAMEN BAĞIMSIZ bir
"transfer" aracıdır. Ürettiği günlük hedef roster excel'i (ki uploaded örnek
dosyayla AYNI yapıdadır: A=TC, L=Çalışma Durumu, M=Rönesans Çalışma Saati
Girişi) daha sonra "İçe Aktar - Çalışma Bilgileri (SAP Mesai)" sayfasından
Veri Merkezi'ne geri yüklenerek kalıcı hale getirilir - iki özellik birlikte
tam iş akışını oluşturur:
    1) SAP Puantaj Aktarım  : aylık kaynak -> günlük hedef excel (bu dosya)
    2) İçe Aktar - Çalışma Bilgileri : günlük hedef excel -> kalıcı depo

AYARLAR - puantaj_suite.py'deki "Puantaj Aktarımı" modülüyle AYNI ayar
dosyasını (puantaj_settings.json) okur, böylece kullanıcının Suite'te zaten
ayarladığı sütun/satır yapılandırması (ör. kaynağın gerçek satır aralığı)
buraya da otomatik yansır - iki yerde ayrı ayrı bakım gerekmez. Ayar dosyası
her transferden HEMEN ÖNCE yeniden okunur (statik önbelleklenmez) - Suite
açıkken ayarları değiştirse bile Veri Merkezi güncel değeri kullanır.
"""

__version__ = "2026-09-21.2"

import datetime
import json
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string

import puantaj_engine as engine

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "OUTPUT"
PUANTAJ_SETTINGS_PATH = SCRIPT_DIR / "puantaj_settings.json"


def get_config():
    """puantaj_suite.py'nin PuantajModule'ü ile BİREBİR AYNI mantık: önce
    engine.DEFAULT_CONFIG, üzerine (varsa) puantaj_settings.json'daki
    kaydedilmiş değerler bindirilir. Her çağrıda diskten TAZE okunur."""
    cfg = dict(engine.DEFAULT_CONFIG)
    if PUANTAJ_SETTINGS_PATH.exists():
        try:
            with open(PUANTAJ_SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            cfg.update({k: v for k, v in saved.items() if k in engine.DEFAULT_CONFIG})
        except Exception:
            pass  # bozuk/okunamayan ayar dosyası -> sessizce varsayılana dön
    return cfg


# Geriye dönük uyumluluk için modül düzeyinde bir isim de tutulur, ama BU
# İÇE AKTARIMDA DEĞİL, HER TRANSFERDEN ÖNCE get_config() ile TAZELENMELİDİR -
# aşağıdaki fonksiyonlar zaten kendi içlerinde get_config() çağırır.
CONFIG = get_config()


def list_available_dates(cfg, source_path, sheet_name=None):
    """Kaynak excelin tarih başlık satırında GERÇEK tarih değeri bulunan
    sütunları tarar, sıralı [datetime.date, ...] döner - tarih seçici bu
    listeyle sınırlandırılır (kaynakta olmayan bir tarih zaten aktarılamaz)."""
    wb = openpyxl.load_workbook(source_path, data_only=True)
    sheet, _w = engine.resolve_sheet(wb, cfg["source_sheet"], "Kaynak")
    ws = wb[sheet]
    start = column_index_from_string(cfg["source_date_col_start"])
    end = column_index_from_string(cfg["source_date_col_end"])
    tarihler = []
    for c in range(start, end + 1):
        v = ws.cell(row=cfg["source_date_header_row"], column=c).value
        if isinstance(v, datetime.datetime):
            v = v.date()
        if isinstance(v, datetime.date):
            tarihler.append(v)
    return sorted(set(tarihler))


def onizleme(cfg, source_path, target_path, target_date_str):
    """Hiçbir şey yazmadan (dry_run) tek bir hedef dosya için sonucu döner."""
    return engine.run_transfer(cfg, source_path, target_path, output_path=None,
                                target_date_str=target_date_str, dry_run=True)


def aktar(cfg, source_path, target_path, target_date_str, manual_mapping=None, output_name=None):
    """Gerçek aktarımı yapar, OUTPUT klasörüne yeni bir dosya olarak kaydeder
    (hedef dosyanın KENDİSİ asla değiştirilmez - engine.run_transfer zaten
    ayrı bir output_path'e save eder). Çıktı dosya adı, VARSAYILAN olarak
    yüklenen hedef dosyanın adıyla BİREBİR AYNIDIR (kullanıcı isteği: tarih/ek
    etiket eklenmez) - aynı isimle art arda birden fazla aktarım yapılırsa
    OUTPUT klasöründeki önceki çıktının üzerine yazılır."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    target_path = Path(target_path)
    if output_name is None:
        output_name = target_path.name
    output_path = OUTPUT_DIR / output_name
    return engine.run_transfer(cfg, source_path, target_path, output_path,
                                target_date_str, manual_mapping=manual_mapping, dry_run=False)


def ozet_metrikleri(result):
    """Bir run_transfer sonucundan Dash KPI kartları için kısa özet."""
    return {
        "eslesen": len(result["matched"]),
        "eslesmeyen": len(result["unmatched"]),
        "izinli": len(result["on_leave"]),
        "anormal": len(result["abnormal"]),
        "tc_sorunlu": len(result["tc_issues"]),
        "kaynakta_var_hedefte_yok": len(result["source_not_in_target"]),
    }
