# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - DASH ARAYÜZÜ (veri_merkezi_app.py)
=====================================================
Uygulama kabuğu: sol menü, sayfa yönlendirme, genel arama. Her sayfa kendi
modülündedir (veri_merkezi_sayfa_*.py) ve callback'lerini dash.callback ile
kaydeder. Görünüm: assets/veri_merkezi.css (Dash otomatik yükler).

Açılışta sırasıyla: şema yükseltme / açılış yedeği -> tablolar + şema (WAL) -> ayarlar (v5).
pywebview ile masaüstü penceresinde gösterilir (bkz. veri_merkezi_main.py);
doğrudan `python veri_merkezi_app.py` ile tarayıcıda da açılabilir.

Gereken paketler:
    pip install dash dash-bootstrap-components plotly pywebview openpyxl
(Dash >= 2.9 gerekir: allow_duplicate özelliği kullanılıyor.)
"""

import re
import time

import dash
from dash import dcc, html, Input, Output, ALL, no_update, callback
import dash_bootstrap_components as dbc

import personnel_db as pdb
import veri_merkezi_ayarlar as ayarlar
import veri_merkezi_sema as sema
import veri_merkezi_sorgu as sorgu
import veri_merkezi_surum as surum
import veri_merkezi_ui as ui
import veri_merkezi_varliklar as varliklar
import veri_merkezi_yedek as yedek

import veri_merkezi_sayfa_genel as s_genel
import veri_merkezi_sayfa_ice_aktar as s_ice
import veri_merkezi_sayfa_sap as s_sap
import veri_merkezi_sayfa_kalite as s_kalite
import veri_merkezi_sayfa_gecmis as s_gecmis
import veri_merkezi_sayfa_personel as s_personel
import veri_merkezi_sayfa_puantaj as s_puantaj
import veri_merkezi_sayfa_ekip as s_ekip
import veri_merkezi_sayfa_raporlar as s_rapor
import veri_merkezi_sayfa_verimlilik as s_verim
import veri_merkezi_sayfa_ayarlar as s_ayarlar

# ---- tanı günlüğü: hangi sayfa ne zaman, kaç saniyede açıldı / hata verdi (sorun bildirirken işe yarar).
# Dosya uygulama klasöründe, en fazla ~2 x 256 KB.
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path as _Path

gunluk = logging.getLogger("veri_merkezi")
# Günlük dosyası yazılamazsa (izin, kilit...) uygulama bundan etkilenmemeli: logging kendi hatalarını
# konsola "--- Logging error ---" diye dökmesin.
logging.raiseExceptions = False
if not gunluk.handlers:
    try:
        _h = RotatingFileHandler(_Path(__file__).resolve().parent / "veri_merkezi_gunluk.log", maxBytes=256_000,
                                 backupCount=1, encoding="utf-8")
        _h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        _h.stream.write("")
        _h.stream.flush()      # açılabilen ama yazılamayan dosyayı burada yakala
        gunluk.addHandler(_h)
        gunluk.setLevel(logging.INFO)
    except (OSError, ValueError):
        pass

__version__ = surum.SURUM   # tek sürüm numarası: veri_merkezi_surum.py

gunluk.info("Veri Merkezi %s başlıyor (veritabanı: %s)", surum.etiket(), sema.DB_PATH)
# SIRA ÖNEMLİ: yedekler, veritabanına dokunan her şeyden (tablo oluşturma, şema yükseltme) ÖNCE alınır.
yedek.sema_yukseltme_yedegi()   # kayıtlı şema eskiyse (ör. v4 -> v5) yükseltmeden önce
yedek.acilis_yedegi()           # günde bir açılış yedeği (Ayarlar'dan kapatılabilir)
pdb.init_db()      # personnel / personnel_history tabloları (yoksa)
sema.migrate()     # Veri Merkezi tabloları + sürüm yükseltmeleri + WAL kipi
ayarlar.uygula()   # veri_merkezi_ayarlar.json -> eşikler, katsayılar, proje adı

app = dash.Dash(
    __name__,
    # Bootstrap, ikonlar ve yazı tipleri: yerel_varliklar/ klasöründe varsa oradan (internet gerekmez),
    # yoksa CDN'den. Bkz. veri_merkezi_varliklar.py ve Ayarlar > Çevrimdışı çalışma.
    external_stylesheets=varliklar.stil_listesi(),
    suppress_callback_exceptions=True,
    title="Veri Merkezi",
    update_title=None,
)
server = app.server
varliklar.rota_kaydet(server)

# (bölüm başlığı, [(anahtar, etiket, ikon, yol, rozet_id)])
MENU = [
    (None, [("genel", "Genel Bakış", "house", "/", None)]),
    ("Veri Girişi", [("ice", "İçe Aktar", "upload", "/ice-aktar", None),
                     ("sap", "SAP Puantaj Aktarım", "arrow-left-right", "/sap-puantaj-aktarim", None)]),
    ("Denetim", [("kalite", "Veri Kalitesi", "shield-check", "/kalite", "nav-uyari-sayi"),
                 ("gecmis", "İçe Aktarma Geçmişi", "clock-history", "/gecmis", None)]),
    ("Kayıtlar", [("personel", "Personel", "people", "/personel", None),
                  ("puantaj", "Puantaj Geçmişi", "calendar3", "/puantaj", None),
                  ("ekip", "Ekip Listesi", "list-ul", "/ekip-listesi", None)]),
    ("Raporlar", [("fm", "Verimlilik Analizi", "graph-up-arrow", "/verimlilik", None),
                  ("maliyet", "Maliyet Raporu", "cash-stack", "/maliyet", None)]),
    ("Sistem", [("ayarlar", "Ayarlar", "gear", "/ayarlar", None)]),
]
MENU_KEYS = [item[0] for _s, items in MENU for item in items]


def _aktif_anahtar(path):
    path = path or "/"
    kurallar = [("/ice-aktar", "ice"), ("/sap-puantaj", "sap"), ("/kalite", "kalite"), ("/gecmis", "gecmis"),
                ("/personel", "personel"), ("/kayitlar", "personel"), ("/puantaj", "puantaj"),
                ("/ekip-listesi", "ekip"), ("/verimlilik", "fm"), ("/fazla-mesai", "fm"), ("/maliyet", "maliyet"),
                ("/ayarlar", "ayarlar")]
    for onek, k in kurallar:
        if path.startswith(onek):
            return k
    return "genel"


def sol_menu():
    gruplar = []
    for baslik, items in MENU:
        linkler = []
        for key, etiket, ik, href, rozet in items:
            parcalar = [ui.ikon(ik), html.Span(etiket, className="vm-nav-label")]
            if rozet:
                parcalar.append(html.Span(id=rozet, className="vm-nav-badge"))
            linkler.append(dcc.Link(parcalar, href=href, id={"type": "nav", "key": key}, className="vm-nav-link"))
        gruplar.append(html.Div(([html.P(baslik, className="vm-navsec")] if baslik else []) + linkler,
                                className="vm-navgrp"))
    try:
        sema_surum = sema.get_schema_version()
    except Exception:
        sema_surum = "?"
    return html.Aside([
        dcc.Link([
            html.Span(ui.ikon("database"), className="vm-brand-mark"),
            html.Div([html.Div("VERİ MERKEZİ", className="vm-brand-name"),
                      html.Div(ayarlar.al("marka_alt_baslik"), id="vm-brand-sub", className="vm-brand-sub")]),
        ], href="/", className="vm-brand"),
        html.Nav(gruplar, className="vm-nav"),
        html.Div([html.Span(sema.DB_PATH.name, className="db"),
                  html.Span(f"şema v{sema_surum} · sürüm {surum.etiket()}", className="ver")], className="vm-side-foot"),
    ], className="vm-side")


app.layout = html.Div([
    dcc.Location(id="url"),
    # Genel arama bu AYRI Location ile yönlendirir (refresh="callback-nav": sayfa yenilenmeden adres değişir,
    # "url" bunu tarayıcı olayıyla görür). "url" hiçbir callback'in ÇIKTISI değildir - bkz. genel_arama_git.
    dcc.Location(id="arama-git", refresh="callback-nav"),
    sol_menu(),
    html.Main(id="page-content", className="vm-main"),
], className="vm-shell")


# ============================================================ yönlendirme

ROTALAR = [
    (r"^/$", s_genel.sayfa),
    (r"^/ice-aktar(-personel|-mesai)?/?$", s_ice.sayfa),
    (r"^/sap-puantaj-aktarim/?$", s_sap.sayfa),
    (r"^/kalite(/(?P<tip>[^/]+))?/?$", s_kalite.sayfa),
    (r"^/gecmis/?$", s_gecmis.sayfa),
    (r"^/(personel|kayitlar)/?$", s_personel.sayfa_liste),
    (r"^/personel/(?P<tc>[^/]+)/?$", s_personel.sayfa_kart),
    (r"^/personel-duzenle(/(?P<tc>[^/]+))?/?$", s_personel.sayfa_duzenle),
    (r"^/puantaj/?$", s_puantaj.sayfa),
    (r"^/ekip-listesi/?$", s_ekip.sayfa),
    (r"^/(verimlilik|fazla-mesai)/?$", s_verim.sayfa),   # eski adres de çalışır
    (r"^/maliyet/?$", s_rapor.sayfa_maliyet),
    (r"^/ayarlar/?$", s_ayarlar.sayfa),
]


def sayfa_uret(pathname):
    pathname = pathname or "/"
    for desen, fn in ROTALAR:
        m = re.match(desen, pathname)
        if m:
            kw = {k: v for k, v in m.groupdict().items() if v}
            return fn(**kw)
    # Bilinmeyen adres: sessizce Genel Bakış'a DÖNME (neden olduğu anlaşılmaz); açıkça söyle.
    return ui.sayfa("Sayfa bulunamadı", "Veri Merkezi", [
        ui.banner([html.Strong("Bu adreste bir sayfa yok: "), html.Code(pathname), ". Soldaki menüden devam edin."],
                  "warn"),
        dcc.Link([ui.ikon("house"), "Genel Bakış"], href="/", className="vm-btn", style={"alignSelf": "start"}),
    ])


@app.server.route("/_tani", methods=["POST"])
def tarayici_tanisi():
    """assets/tani.js'in gönderdiği tarayıcı hatalarını ve tam sayfa yenilemelerini günlüğe yazar."""
    from flask import request
    try:
        v = request.get_json(force=True, silent=True) or {}
        gunluk.warning("TARAYICI %s [%s] %s", v.get("tur"), v.get("yol"), str(v.get("mesaj", ""))[:2000])
    except Exception:
        pass
    return "", 204


@callback(Output("page-content", "children"), Input("url", "pathname"))
def render_page(pathname):
    t0 = time.perf_counter()
    try:
        sonuc = sayfa_uret(pathname)
        gunluk.info("sayfa %s (%.2f sn)", pathname, time.perf_counter() - t0)
        return sonuc
    except Exception as e:   # bir sayfa hata verirse uygulamanın geri kalanı çalışmaya devam etsin
        gunluk.exception("sayfa %s yüklenemedi", pathname)
        return ui.sayfa("Hata", "Veri Merkezi", [ui.banner(f"Sayfa yüklenemedi: {e}", "err")])


@callback(
    Output({"type": "nav", "key": ALL}, "className"),
    Output("nav-uyari-sayi", "children"),
    Output("vm-brand-sub", "children"),
    Input("url", "pathname"),
)
def menu_durumu(pathname):
    aktif = _aktif_anahtar(pathname)
    try:
        acik = sorgu.get_alert_totals()["acik"]
    except Exception:
        acik = 0
    try:
        alt_baslik = ayarlar.al("marka_alt_baslik")   # Ayarlar'da değişince bir sonraki gezinmede görünür
    except Exception:
        alt_baslik = no_update
    return ([("vm-nav-link on" if k == aktif else "vm-nav-link") for k in MENU_KEYS],
            ui.sayi(acik) if acik else "", alt_baslik)


@callback(
    Output("arama-git", "pathname"),
    Output("arama-git", "search"),
    Input("global-arama", "value"),
    prevent_initial_call=True,
)
def genel_arama_git(tc):
    """ÖNEMLİ: Çıktı "url" DEĞİL, ayrı "arama-git" Location'ı.
    Üst çubuktaki arama kutusu her sayfayla yeniden çizildiği için bu callback HER sayfa açılışında tetiklenir.
    Çıktısı "url" iken: (1) yeni sayfanın yüklenmesi bu callback bitene kadar bekletiliyordu (sunucu Verimlilik
    analizi gibi ağır bir işle meşgulken sayfa geç açılıyordu); (2) "url" bileşeni bu sırada eski adresle yeniden
    çizilirse refresh=True olduğu için tarayıcıyı o eski adrese ("/" = Genel Bakış) tam yeniliyordu.
    search'e zaman damgası: aynı kişi tekrar arandığında da yönlendirme olsun diye."""
    if not tc:
        return no_update, no_update
    return f"/personel/{tc}", f"?a={int(time.time() * 1000)}"


@callback(
    Output("global-arama", "options"),
    Input("global-arama", "search_value"),
    prevent_initial_call=True,
)
def genel_arama_secenekleri(aranan):
    """Yazdıkça en fazla 25 eşleşme döner (ad, TC, görev içinde; Türkçe harf farkı gözetmeksizin)."""
    aranan = (aranan or "").strip()
    if len(aranan) < 2:
        return no_update
    import puantaj_engine as engine
    q = engine.normalize_name(aranan)
    sonuc = []
    for k in sorgu.get_all_personnel_brief():
        metin = engine.normalize_name(f"{k['ad_soyad']} {k['tc']} {k.get('gorevi') or ''}")
        if q in metin:
            sonuc.append({"label": f"{k['ad_soyad']} · {k['tc']}" + (f" · {k['gorevi']}" if k.get("gorevi") else ""),
                          "value": k["tc"],
                          # Dropdown seçenekleri tarayıcıda da süzer; yazılan metni ekleyerek
                          # "sahin" -> "ŞAHİN" gibi Türkçe harf farklarında eşleşmenin kaybolmasını önleriz.
                          "search": f"{aranan} {metin}"})
            if len(sonuc) >= 25:
                break
    return sonuc


if __name__ == "__main__":
    app.run(debug=True, port=8765)
