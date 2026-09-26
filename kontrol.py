# -*- coding: utf-8 -*-
"""Veri Merkezi kurulum kontrolü: python kontrol.py
Uygulamanın ihtiyaç duyduğu dosyaların bu klasörde olup olmadığını ve paketlerin kurulu olup olmadığını söyler."""
import importlib.util
import sys
from pathlib import Path

KOK = Path(__file__).resolve().parent
DOSYALAR = """personnel_db puantaj_engine team_report cost_report veri_merkezi_app veri_merkezi_main veri_merkezi_ayarlar
veri_merkezi_baslangic veri_merkezi_dogrulama veri_merkezi_dosya_tespit veri_merkezi_hafta_kurali veri_merkezi_ekip_excel veri_merkezi_ekip_listesi veri_merkezi_geri_alma
veri_merkezi_ice_aktarma veri_merkezi_ik veri_merkezi_maliyet veri_merkezi_personel_karti veri_merkezi_mesai_formu veri_merkezi_mesai_ice_aktarma
veri_merkezi_puantaj_aktarim veri_merkezi_sayfa_ayarlar veri_merkezi_sayfa_ekip veri_merkezi_sayfa_gecmis
veri_merkezi_sayfa_genel veri_merkezi_sayfa_ice_aktar veri_merkezi_sayfa_kalite veri_merkezi_sayfa_personel
veri_merkezi_sayfa_puantaj veri_merkezi_sayfa_raporlar veri_merkezi_sayfa_sap veri_merkezi_sayfa_verimlilik
veri_merkezi_sema veri_merkezi_sorgu veri_merkezi_surum veri_merkezi_ui veri_merkezi_varliklar veri_merkezi_verimlilik
veri_merkezi_verimlilik_excel veri_merkezi_yedek""".split()
PAKETLER = ["dash", "dash_bootstrap_components", "plotly", "webview", "openpyxl", "reportlab", "flask"]

eksik = [f"{d}.py" for d in DOSYALAR if not (KOK / f"{d}.py").is_file()]
eksik += [f"assets/{a}" for a in ("veri_merkezi.css", "tani.js", "kisayollar.js") if not (KOK / "assets" / a).is_file()]
paket_eksik = [p for p in PAKETLER if importlib.util.find_spec(p) is None]
alt = [p.name for p in KOK.iterdir() if p.is_dir() and (p / "veri_merkezi_app.py").exists()]

print(f"Klasör: {KOK}\nPython: {sys.version.split()[0]} ({sys.executable})")
if alt:
    print(f"\n! Alt klasörde uygulama dosyaları var: {', '.join(alt)}  -> içindekileri bu klasöre taşıyın.")
if eksik:
    print("\nEKSİK DOSYALAR:\n  " + "\n  ".join(eksik))
if paket_eksik:
    ad = {"webview": "pywebview", "dash_bootstrap_components": "dash-bootstrap-components"}
    print("\nEKSİK PAKETLER:\n  pip install " + " ".join(ad.get(p, p) for p in paket_eksik))
if not (KOK / "fonts").is_dir():
    print("\nNot: fonts/ klasörü yok - Günlük Mesai Formu (PDF) için DejaVu yazı tipleri gerekir; eski kurulumdan kopyalayın.")
print("\nTAMAM: uygulama başlatılabilir." if not (eksik or paket_eksik) else "\nYukarıdakiler düzeltilmeden uygulama açılmaz.")
