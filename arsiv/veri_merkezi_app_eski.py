# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - DASH ARAYÜZÜ (veri_merkezi_app.py)
=====================================================
Profesyonel, tek-ekran dashboard. pywebview ile masaüstü penceresinde
gösterilmek üzere tasarlanmıştır (bkz. veri_merkezi_main.py).

NOT: Bu dosya bu ortamda (internet erişimi olmadığı için dash/plotly/
dash-bootstrap-components kurulamadı) ÇALIŞTIRILARAK test edilemedi -
sadece dikkatle yazıldı ve Python söz dizimi olarak doğrulandı. Backend
katmanı (şema/doğrulama/içe aktarma/sorgu) gerçek veriyle eksiksiz test
edildi. Kurulumdan sonra küçük görsel/callback düzeltmeleri gerekebilir.
"""

import base64
import datetime
import tempfile
from pathlib import Path

import dash
from dash import dcc, html, Input, Output, State, dash_table, no_update, ALL, ctx
import dash_bootstrap_components as dbc
import plotly.graph_objects as go

import puantaj_engine as engine
import veri_merkezi_sema as sema
import veri_merkezi_sorgu as sorgu
import veri_merkezi_ice_aktarma as ie
import veri_merkezi_mesai_ice_aktarma as mie
import veri_merkezi_puantaj_aktarim as pa
import veri_merkezi_dogrulama as dogrulama
import veri_merkezi_ekip_listesi as el
import personnel_db as pdb

__version__ = "2026-09-21.2"

sema.migrate()

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY, dbc.icons.BOOTSTRAP],
    suppress_callback_exceptions=True,
    title="Veri Merkezi",
)
server = app.server  # pywebview / olası bir WSGI sarmalayıcı için

# Basit, tek-kullanıcılı sunucu-içi önbellek: önizleme sonuçları büyük olabileceği
# (on binlerce günlük kayıt) için istemci tarafına (dcc.Store/JSON) taşınmaz,
# burada tutulur; istemciye sadece hafif bir özet + bir "preview_id" gönderilir.
_PENDING_PREVIEW = {}
_PREVIEW_COUNTER = {"n": 0}
_PENDING_MESAI_PREVIEW = {}
_MESAI_PREVIEW_COUNTER = {"n": 0}
# SAP Puantaj Aktarım: yüklenen kaynak/hedef dosyaların GEÇİCİ yolları ve son
# önizleme sonucu - tek kullanıcılı masaüstü uygulaması olduğu için sunucu
# belleğinde tutmak yeterli (bkz. içe aktarma sayfalarındaki AYNI desen).
_PENDING_SAP = {}
_SAP_COUNTER = {"n": 0}


# ============================================================ küçük bileşenler

def kpi_card(baslik, deger, renk="primary"):
    return dbc.Card(
        dbc.CardBody([
            html.Div(baslik, className="text-muted small fw-semibold text-uppercase"),
            html.H3(deger, className=f"text-{renk} fw-bold mb-0 mt-1"),
        ]),
        className="shadow-sm border-0 h-100",
    )


def sidebar():
    links = [
        ("📊  Genel Bakış", "/"),
        ("📥  İçe Aktar - Personel Bilgileri", "/ice-aktar-personel"),
        ("📥  İçe Aktar - Çalışma Bilgileri", "/ice-aktar-mesai"),
        ("🚚  SAP Puantaj Aktarım", "/sap-puantaj-aktarim"),
        ("🔍  Veri Kalitesi", "/kalite"),
        ("📅  Puantaj Geçmişi", "/puantaj"),
        ("👥  Personel Kayıtları", "/kayitlar"),
        ("✏️  Personel Düzenle", "/personel-duzenle"),
        ("🧾  Ekip Listesi", "/ekip-listesi"),
        ("🕓  İçe Aktarma Geçmişi", "/gecmis"),
    ]
    return dbc.Nav(
        [dbc.NavLink(label, href=href, active="exact", className="mb-1 fw-semibold") for label, href in links],
        vertical=True, pills=True, className="p-2",
    )


app.layout = html.Div([
    dcc.Location(id="url"),
    dcc.Store(id="preview-id-store"),
    dcc.Store(id="refresh-signal", data=0),
    html.Div([
        html.H4("🗂️ Veri Merkezi", className="fw-bold mt-3 mb-3 px-3"),
        sidebar(),
        html.Div(f"sürüm {__version__}", className="text-muted small px-3 mt-3"),
    ], style={"width": "230px", "position": "fixed", "top": 0, "bottom": 0, "left": 0,
              "backgroundColor": "#ffffff", "borderRight": "1px solid #e2e2e2", "overflowY": "auto"}),
    html.Div(id="page-content", style={"marginLeft": "230px", "padding": "28px"}),
], style={"backgroundColor": "#f4f6f8", "minHeight": "100vh"})


# ============================================================ SAYFA: genel bakış

def sayfa_genel_bakis():
    ozet = sorgu.get_dashboard_summary()
    gruplar = sorgu.get_group_breakdown()
    uyari_tipleri = sorgu.get_alert_type_counts()
    donemler = sorgu.get_covered_periods()
    trend = sorgu.get_monthly_trend()

    grup_fig = go.Figure()
    if gruplar:
        grup_fig.add_bar(name="Normal", x=[g["grup"] for g in gruplar],
                          y=[g["normal"] for g in gruplar], marker_color="#2E86C1")
        grup_fig.add_bar(name="Fazla", x=[g["grup"] for g in gruplar],
                          y=[g["fazla"] for g in gruplar], marker_color="#E74C3C")
    grup_fig.update_layout(barmode="stack", title="Grup Bazında Mesai (Normal + Fazla)",
                             template="plotly_white", margin=dict(t=50, b=30, l=30, r=20), height=380,
                             legend=dict(orientation="h", y=1.1))

    trend_fig = go.Figure()
    if trend:
        trend_fig.add_scatter(x=[t["ay"] for t in trend], y=[t["normal"] for t in trend],
                               name="Normal", mode="lines+markers", line=dict(color="#2E86C1"))
        trend_fig.add_scatter(x=[t["ay"] for t in trend], y=[t["fazla"] for t in trend],
                               name="Fazla", mode="lines+markers", line=dict(color="#E74C3C"))
    trend_fig.update_layout(title="Aylık Mesai Trendi", template="plotly_white",
                              margin=dict(t=50, b=30, l=30, r=20), height=380,
                              legend=dict(orientation="h", y=1.1))

    if uyari_tipleri:
        alert_rows = [
            dbc.ListGroupItem([
                html.Span(a["tip"].replace("_", " ").title(), className="fw-semibold"),
                dbc.Badge(str(a["adet"]), color="danger" if a["onem"] == "engelleyici" else "warning",
                           className="ms-2 float-end"),
            ]) for a in uyari_tipleri
        ]
    else:
        alert_rows = [dbc.ListGroupItem("Açık veri kalitesi uyarısı yok ✓", color="success")]

    donem_badges = ([dbc.Badge(d, color="light", text_color="dark", className="me-2 mb-2 p-2 border")
                       for d in donemler] if donemler
                      else [html.Span("Henüz hiç içe aktarma yapılmadı.", className="text-muted")])

    return html.Div([
        html.H3("Genel Bakış", className="fw-bold mb-4"),
        dbc.Row([
            dbc.Col(kpi_card("Toplam Personel", ozet["toplam_personel"], "primary"), md=2),
            dbc.Col(kpi_card("Aktif Personel", ozet["aktif_personel"], "success"), md=2),
            dbc.Col(kpi_card("Toplam Mesai (sa)", f"{ozet['toplam_saat']:,.0f}"), md=2),
            dbc.Col(kpi_card("Yevmiye Günü", f"{ozet['yevmiye_gunu']:,.1f}", "secondary"), md=2),
            dbc.Col(kpi_card("Açık Uyarı", ozet["acik_uyari"], "warning" if ozet["acik_uyari"] else "success"), md=2),
            dbc.Col(kpi_card("Kritik Uyarı", ozet["kritik_uyari"], "danger" if ozet["kritik_uyari"] else "success"), md=2),
        ], className="g-3 mb-4"),
        dbc.Row([
            dbc.Col(dbc.Card(dbc.CardBody(
                dcc.Graph(figure=grup_fig, config={"displayModeBar": False}) if gruplar
                else html.Div("Henüz veri yok. 'İçe Aktar' sayfasından başlayın.", className="text-muted py-5 text-center")
            ), className="shadow-sm border-0"), md=7),
            dbc.Col(dbc.Card(dbc.CardBody(
                dcc.Graph(figure=trend_fig, config={"displayModeBar": False}) if trend
                else html.Div("Trend için en az 1 ay veri gerekli.", className="text-muted py-5 text-center")
            ), className="shadow-sm border-0"), md=5),
        ], className="g-3 mb-4"),
        dbc.Row([
            dbc.Col([
                html.H5("Veri Kalitesi Uyarıları", className="fw-semibold mb-2"),
                dbc.ListGroup(alert_rows),
            ], md=4),
            dbc.Col([
                html.H5("Verisi Bulunan Aylar", className="fw-semibold mb-2"),
                html.Div(donem_badges),
            ], md=8),
        ]),
    ])


# ============================================================ SAYFA: içe aktar

def sayfa_ice_aktar_personel():
    return html.Div([
        html.H3("İçe Aktar — Personel Bilgileri", className="fw-bold mb-1"),
        html.P("Aylık kaynak excel'den SADECE kimlik bilgisi çekilir: ad soyad, TC, grup, "
               "bağlı olduğu ekip, görev, işe giriş/çıkış tarihi. Çalışma saatleri BURADAN "
               "DEĞİL, 'İçe Aktar - Çalışma Bilgileri' sayfasından gelir.", className="text-muted mb-3"),
        dbc.Card(dbc.CardBody([
            dcc.Upload(
                id="upload-excel-personel",
                children=html.Div([
                    html.I(className="bi bi-cloud-upload fs-1 text-primary d-block mb-2"),
                    "Kaynak puantaj excel'ini buraya sürükle veya ",
                    html.A("dosya seç"),
                ]),
                style={
                    "border": "2px dashed #adb5bd", "borderRadius": "10px",
                    "padding": "40px", "textAlign": "center", "cursor": "pointer",
                },
                multiple=False,
            ),
        ]), className="shadow-sm border-0 mb-4"),
        dcc.Loading(html.Div(id="preview-personel-sonucu")),
        html.Div(id="commit-personel-sonucu", className="mt-3"),
    ])


def _cakisma_karti(conflicts):
    """Her alan çakışması için bir radyo düğmesi grubu: kullanıcı hangi
    ayın/dosyanın değerini baz alacağını AÇIKÇA seçer - hiçbir şey
    otomatik/sessizce üzerine yazılmaz."""
    if not conflicts:
        return None

    items = []
    for c in conflicts:
        radio_id = {"type": "conflict-choice", "tc": c["tc"], "field": c["alan"]}
        items.append(
            dbc.ListGroupItem([
                html.Div([
                    html.Strong(c["ad_soyad_baglam"]), f" — {c['alan_etiket']}",
                ], className="mb-2"),
                dbc.RadioItems(
                    id=radio_id,
                    options=[
                        {"label": f"Mevcut değeri koru: \u201c{c['eski_deger']}\u201d  ({c['eski_kaynak']})",
                         "value": "existing"},
                        {"label": f"Bu içe aktarmayı kullan: \u201c{c['yeni_deger']}\u201d  ({c['yeni_kaynak']})",
                         "value": "new"},
                    ],
                    value="existing",  # varsayılan: mevcut korunur, hiçbir şey sessizce değişmez
                    inline=False,
                ),
            ])
        )

    return dbc.Card(dbc.CardBody([
        html.H5(f"⚠ {len(conflicts)} Bilgi Çakışması Bulundu", className="text-warning fw-bold"),
        html.P("Aşağıdaki kişilerin bilgileri, veri setinde zaten kayıtlı olandan farklı. "
               "Her biri için hangi kaynağı baz almak istediğini seç.", className="text-muted"),
        dbc.ListGroup(items, className="mb-3", style={"maxHeight": "420px", "overflowY": "auto"}),
    ]), className="shadow-sm border-warning mb-3")


def _preview_personel_ozet_karti(preview_id, preview):
    if preview["durum"] == "engellendi":
        return dbc.Alert([
            html.H5("İçe aktarma durduruldu", className="alert-heading"),
            html.Ul([html.Li(a["detay"]) for a in preview["engelleyiciler"]]),
        ], color="danger")

    uyarilar = preview["uyarilar"]
    onemli_tipler = {}
    for a in uyarilar:
        onemli_tipler[a["tip"]] = onemli_tipler.get(a["tip"], 0) + 1

    uyari_liste = [
        dbc.ListGroupItem(f"{ALERT_TIP_ETIKET.get(tip, tip.replace('_', ' ').title())}: {adet} adet")
        for tip, adet in sorted(onemli_tipler.items(), key=lambda kv: -kv[1])
    ] or [dbc.ListGroupItem("Uyarı yok ✓", color="success")]

    conflicts = preview.get("field_conflicts") or []
    cakisma_karti = _cakisma_karti(conflicts)

    return html.Div([
        dbc.Alert(f"'{preview['dosya_adi']}' okundu — {len(preview['personel_kayitlari'])} personel bulundu.",
                    color="info"),
        dbc.Row([
            dbc.Col(kpi_card("Taranan Satır", preview["taranan_satir"]), md=4),
            dbc.Col(kpi_card("Personel", len(preview["personel_kayitlari"])), md=4),
            dbc.Col(kpi_card("Uyarı", len(uyarilar), "warning" if uyarilar else "success"), md=4),
        ], className="g-3 mb-3"),
        html.H6("Uyarı Özeti", className="fw-semibold"),
        dbc.ListGroup(uyari_liste, className="mb-3"),
        cakisma_karti if cakisma_karti else html.Div(),
        dbc.Button("✓ Onayla ve Kalıcı Depoya Yaz", id="commit-personel-btn", color="success", size="lg"),
        dcc.Store(id="active-personel-preview-id", data=preview_id),
    ])


@app.callback(
    Output("preview-personel-sonucu", "children"),
    Input("upload-excel-personel", "contents"),
    State("upload-excel-personel", "filename"),
    prevent_initial_call=True,
)
def on_upload_personel(contents, filename):
    if not contents:
        return no_update
    _content_type, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)

    tmp_dir = Path(tempfile.gettempdir()) / "veri_merkezi_yuklemeler"
    tmp_dir.mkdir(exist_ok=True)
    tmp_path = tmp_dir / filename
    with open(tmp_path, "wb") as f:
        f.write(decoded)

    try:
        preview = ie.preview_import(tmp_path)
    except Exception as e:
        return dbc.Alert(f"Dosya okunamadı: {e}", color="danger")

    _PREVIEW_COUNTER["n"] += 1
    preview_id = _PREVIEW_COUNTER["n"]
    _PENDING_PREVIEW[preview_id] = preview
    return _preview_personel_ozet_karti(preview_id, preview)


@app.callback(
    Output("commit-personel-sonucu", "children"),
    Input("commit-personel-btn", "n_clicks"),
    State("active-personel-preview-id", "data"),
    State({"type": "conflict-choice", "tc": ALL, "field": ALL}, "value"),
    State({"type": "conflict-choice", "tc": ALL, "field": ALL}, "id"),
    prevent_initial_call=True,
)
def on_commit_personel(n_clicks, preview_id, secim_degerleri, secim_idleri):
    preview = _PENDING_PREVIEW.get(preview_id)
    if preview is None:
        return dbc.Alert("Önizleme süresi doldu, lütfen dosyayı tekrar yükleyin.", color="warning")

    field_resolutions = {
        (id_dict["tc"], id_dict["field"]): deger
        for id_dict, deger in zip(secim_idleri, secim_degerleri)
    }

    try:
        sonuc = ie.commit_import(preview, field_resolutions=field_resolutions)
    except Exception as e:
        return dbc.Alert(f"İçe aktarma sırasında hata oluştu, HİÇBİR ŞEY YAZILMADI: {e}", color="danger")
    finally:
        _PENDING_PREVIEW.pop(preview_id, None)

    return dbc.Alert(
        f"✓ Başarıyla içe aktarıldı — {sonuc['yeni_personel']} yeni personel, "
        f"{sonuc['guncellenen_personel']} güncellendi, "
        f"{sonuc['uyari_sayisi']} uyarı 'Veri Kalitesi' sayfasına eklendi. "
        f"Güncel rakamları görmek için 'Genel Bakış' sayfasına gidin.",
        color="success",
    )


# ============================================================ SAYFA: içe aktar - çalışma bilgileri (SAP Mesai)

def sayfa_ice_aktar_mesai():
    return html.Div([
        html.H3("İçe Aktar — Çalışma Bilgileri (SAP Mesai)", className="fw-bold mb-1"),
        html.P("GÜNLÜK SAP roster excel'inden (TC, Çalışma Durumu, Rönesans Çalışma Saati Girişi) "
               "o günün normal+fazla mesai toplamı çekilir. Kaynaktaki tek 'toplam saat' değeri, "
               "standart 9 saatlik iş gününe göre normal/fazla olarak bölünür.", className="text-muted mb-3"),
        dbc.Card(dbc.CardBody([
            dcc.Upload(
                id="upload-excel-mesai",
                children=html.Div([
                    html.I(className="bi bi-cloud-upload fs-1 text-success d-block mb-2"),
                    "Günlük SAP mesai excel'ini buraya sürükle veya ",
                    html.A("dosya seç"),
                ]),
                style={
                    "border": "2px dashed #adb5bd", "borderRadius": "10px",
                    "padding": "40px", "textAlign": "center", "cursor": "pointer",
                },
                multiple=False,
            ),
        ]), className="shadow-sm border-0 mb-4"),
        dcc.Loading(html.Div(id="preview-mesai-sonucu")),
        html.Div(id="commit-mesai-sonucu", className="mt-3"),
    ])


def _preview_mesai_ozet_karti(preview_id, preview):
    if preview["durum"] == "engellendi":
        return dbc.Alert([
            html.H5("İçe aktarma durduruldu", className="alert-heading"),
            html.Ul([html.Li(a["detay"]) for a in preview["engelleyiciler"]]),
        ], color="danger")

    uyarilar = preview["uyarilar"]
    onemli_tipler = {}
    for a in uyarilar:
        onemli_tipler[a["tip"]] = onemli_tipler.get(a["tip"], 0) + 1

    uyari_liste = [
        dbc.ListGroupItem(f"{ALERT_TIP_ETIKET.get(tip, tip.replace('_', ' ').title())}: {adet} adet")
        for tip, adet in sorted(onemli_tipler.items(), key=lambda kv: -kv[1])
    ] or [dbc.ListGroupItem("Uyarı yok ✓", color="success")]

    donem = preview["donem_baslangic"]
    if preview["donem_baslangic"] != preview["donem_bitis"]:
        donem = f"{preview['donem_baslangic']} — {preview['donem_bitis']}"

    return html.Div([
        dbc.Alert(f"'{preview['dosya_adi']}' okundu — Tarih: {donem}", color="info"),
        dbc.Row([
            dbc.Col(kpi_card("Taranan Satır", preview["taranan_satir"]), md=4),
            dbc.Col(kpi_card("Günlük Kayıt", len(preview["gunluk_kayitlar"])), md=4),
            dbc.Col(kpi_card("Uyarı", len(uyarilar), "warning" if uyarilar else "success"), md=4),
        ], className="g-3 mb-3"),
        html.H6("Uyarı Özeti", className="fw-semibold"),
        dbc.ListGroup(uyari_liste, className="mb-3"),
        dbc.Button("✓ Onayla ve Kalıcı Depoya Yaz", id="commit-mesai-btn", color="success", size="lg"),
        dcc.Store(id="active-mesai-preview-id", data=preview_id),
    ])


@app.callback(
    Output("preview-mesai-sonucu", "children"),
    Input("upload-excel-mesai", "contents"),
    State("upload-excel-mesai", "filename"),
    prevent_initial_call=True,
)
def on_upload_mesai(contents, filename):
    if not contents:
        return no_update
    _content_type, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)

    tmp_dir = Path(tempfile.gettempdir()) / "veri_merkezi_yuklemeler"
    tmp_dir.mkdir(exist_ok=True)
    tmp_path = tmp_dir / filename
    with open(tmp_path, "wb") as f:
        f.write(decoded)

    try:
        preview = mie.preview_mesai_import(tmp_path)
    except Exception as e:
        return dbc.Alert(f"Dosya okunamadı: {e}", color="danger")

    _MESAI_PREVIEW_COUNTER["n"] += 1
    preview_id = _MESAI_PREVIEW_COUNTER["n"]
    _PENDING_MESAI_PREVIEW[preview_id] = preview
    return _preview_mesai_ozet_karti(preview_id, preview)


@app.callback(
    Output("commit-mesai-sonucu", "children"),
    Input("commit-mesai-btn", "n_clicks"),
    State("active-mesai-preview-id", "data"),
    prevent_initial_call=True,
)
def on_commit_mesai(n_clicks, preview_id):
    preview = _PENDING_MESAI_PREVIEW.get(preview_id)
    if preview is None:
        return dbc.Alert("Önizleme süresi doldu, lütfen dosyayı tekrar yükleyin.", color="warning")

    try:
        sonuc = mie.commit_mesai_import(preview)
    except Exception as e:
        return dbc.Alert(f"İçe aktarma sırasında hata oluştu, HİÇBİR ŞEY YAZILMADI: {e}", color="danger")
    finally:
        _PENDING_MESAI_PREVIEW.pop(preview_id, None)

    return dbc.Alert(
        f"✓ Başarıyla içe aktarıldı — {sonuc['yeni_kayit']} yeni günlük kayıt, "
        f"{sonuc['guncellenen_kayit']} güncellendi, {sonuc['degismeyen_kayit']} değişmedi, "
        f"{sonuc['uyari_sayisi']} uyarı 'Veri Kalitesi' sayfasına eklendi. "
        f"Güncel rakamları görmek için 'Genel Bakış' sayfasına gidin.",
        color="success",
    )


# ============================================================ SAYFA: SAP Puantaj Aktarım
# puantaj_suite.py'deki "Puantaj Aktarımı" modülünün (puantaj_engine.py) Dash
# karşılığı - MOTOR TEKRARLANMAZ (bkz. veri_merkezi_puantaj_aktarim.py). Bu
# sayfa Veri Merkezi'nin kalıcı deposunu OKUMAZ/YAZMAZ; tamamen bağımsız bir
# "aylık kaynak -> günlük hedef roster" transfer aracıdır. Ürettiği günlük
# hedef dosya, ayrıca "İçe Aktar - Çalışma Bilgileri" sayfasından Veri
# Merkezi'ne geri yüklenebilir.

_SAP_STATE = {"kaynak": None, "hedefler": []}


def sayfa_sap_aktarim():
    return html.Div([
        html.H3("SAP Puantaj Aktarım", className="fw-bold mb-1"),
        html.P("Aylık kaynak excel'den, seçilen GÜN için normal+fazla toplamını TC eşleştirerek "
               "günlük hedef roster excel(ler)ine yazar. Hedef dosyanın kendisi DEĞİŞTİRİLMEZ - "
               "sonuç yeni bir dosya olarak kaydedilir.", className="text-muted mb-3"),
        dbc.Row([
            dbc.Col(dbc.Card(dbc.CardBody([
                html.H6("1. Aylık Kaynak Excel", className="fw-semibold"),
                dcc.Upload(id="sap-kaynak-upload",
                           children=html.Div([html.I(className="bi bi-cloud-upload me-2"), "Kaynak seç"]),
                           style={"border": "2px dashed #adb5bd", "borderRadius": "8px",
                                  "padding": "16px", "textAlign": "center", "cursor": "pointer"},
                           multiple=False),
                html.Div(id="sap-kaynak-info", className="small text-muted mt-2"),
            ]), className="shadow-sm border-0 h-100"), md=6),
            dbc.Col(dbc.Card(dbc.CardBody([
                html.H6("2. Günlük Hedef Roster Excel(ler)i", className="fw-semibold"),
                dcc.Upload(id="sap-hedef-upload",
                           children=html.Div([html.I(className="bi bi-cloud-upload me-2"), "Hedef(ler) seç"]),
                           style={"border": "2px dashed #adb5bd", "borderRadius": "8px",
                                  "padding": "16px", "textAlign": "center", "cursor": "pointer"},
                           multiple=True),
                html.Div(id="sap-hedef-info", className="small text-muted mt-2"),
            ]), className="shadow-sm border-0 h-100"), md=6),
        ], className="g-3 mb-3"),

        dbc.Row([
            dbc.Col(dcc.Dropdown(id="sap-tarih", placeholder="Önce kaynak excel yükle..."), md=4),
            dbc.Col(dbc.Button("🔍 Önizle", id="sap-onizle-btn", color="primary", className="me-2"), md="auto"),
            dbc.Col(dbc.Button("✓ Aktarımı Yap ve Kaydet", id="sap-aktar-btn", color="success"), md="auto"),
        ], className="g-2 mb-3 align-items-center"),

        dcc.Loading(html.Div(id="sap-onizleme")),
        html.Div(id="sap-sonuc", className="mt-3"),
        dcc.Store(id="sap-dummy"),
    ])


@app.callback(
    Output("sap-kaynak-info", "children"),
    Output("sap-tarih", "options"),
    Output("sap-tarih", "value"),
    Input("sap-kaynak-upload", "contents"),
    State("sap-kaynak-upload", "filename"),
    prevent_initial_call=True,
)
def on_sap_kaynak_upload(contents, filename):
    if not contents:
        return no_update, no_update, no_update
    _ctype, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)
    tmp_dir = Path(tempfile.gettempdir()) / "veri_merkezi_yuklemeler"
    tmp_dir.mkdir(exist_ok=True)
    tmp_path = tmp_dir / filename
    with open(tmp_path, "wb") as f:
        f.write(decoded)
    _SAP_STATE["kaynak"] = tmp_path

    try:
        tarihler = pa.list_available_dates(pa.get_config(), tmp_path)
    except Exception as e:
        return dbc.Alert(f"Kaynak okunamadı: {e}", color="danger"), [], None
    if not tarihler:
        return dbc.Alert("Kaynakta hiçbir tarih sütunu bulunamadı.", color="warning"), [], None

    options = [{"label": t.strftime("%d.%m.%Y"), "value": t.isoformat()} for t in tarihler]
    return (f"✓ {filename} yüklendi — {len(tarihler)} gün mevcut.",
            options, tarihler[-1].isoformat())


@app.callback(
    Output("sap-hedef-info", "children"),
    Input("sap-hedef-upload", "contents"),
    State("sap-hedef-upload", "filename"),
    prevent_initial_call=True,
)
def on_sap_hedef_upload(contents_list, filenames):
    if not contents_list:
        return no_update
    tmp_dir = Path(tempfile.gettempdir()) / "veri_merkezi_yuklemeler"
    tmp_dir.mkdir(exist_ok=True)
    yollar = []
    for contents, filename in zip(contents_list, filenames):
        _ctype, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)
        tmp_path = tmp_dir / filename
        with open(tmp_path, "wb") as f:
            f.write(decoded)
        yollar.append(tmp_path)
    _SAP_STATE["hedefler"] = yollar
    return f"✓ {len(yollar)} hedef dosya yüklendi: {', '.join(p.name for p in yollar)}"


def _sap_ozet_karti(hedef_adi, sonuc):
    ozet = pa.ozet_metrikleri(sonuc)
    kartlar = dbc.Row([
        dbc.Col(kpi_card("Eşleşen", ozet["eslesen"], "success"), md=2),
        dbc.Col(kpi_card("Eşleşmeyen", ozet["eslesmeyen"], "danger" if ozet["eslesmeyen"] else "success"), md=2),
        dbc.Col(kpi_card("İzinli/Raporlu", ozet["izinli"], "secondary"), md=2),
        dbc.Col(kpi_card("Anormal Değer", ozet["anormal"], "warning" if ozet["anormal"] else "success"), md=2),
        dbc.Col(kpi_card("TC Sorunlu", ozet["tc_sorunlu"], "warning" if ozet["tc_sorunlu"] else "success"), md=2),
        dbc.Col(kpi_card("Kaynakta Var, Hedefte Yok", ozet["kaynakta_var_hedefte_yok"]), md=2),
    ], className="g-2 mb-2")

    detay = []
    if sonuc["unmatched"]:
        detay.append(html.Div([
            html.Strong("Eşleşmeyenler (hedefte TC var, kaynakta yok): "),
            ", ".join(f"satır {r} ({ad or tc})" for r, tc, ad in sonuc["unmatched"][:15]),
            " ..." if len(sonuc["unmatched"]) > 15 else "",
        ], className="small text-danger mb-1"))
    if sonuc["tc_issues"]:
        detay.append(html.Div([
            html.Strong("TC format/checksum sorunlu: "),
            ", ".join(f"{i['tc']} ({i['side']}, satır {i['row']})" for i in sonuc["tc_issues"][:15]),
        ], className="small text-warning mb-1"))

    return dbc.Card(dbc.CardBody([html.H6(hedef_adi, className="fw-semibold"), kartlar, *detay]),
                     className="shadow-sm border-0 mb-3")


@app.callback(
    Output("sap-onizleme", "children"),
    Input("sap-onizle-btn", "n_clicks"),
    State("sap-tarih", "value"),
    prevent_initial_call=True,
)
def on_sap_onizle(_n, tarih):
    kaynak = _SAP_STATE.get("kaynak")
    hedefler = _SAP_STATE.get("hedefler") or []
    if not kaynak:
        return dbc.Alert("Önce bir kaynak excel yükle.", color="warning")
    if not hedefler:
        return dbc.Alert("Önce en az bir hedef roster excel yükle.", color="warning")
    if not tarih:
        return dbc.Alert("Bir tarih seç.", color="warning")

    kartlar = []
    for hedef in hedefler:
        try:
            sonuc = pa.onizleme(pa.get_config(), kaynak, hedef, tarih)
        except engine.TransferError as e:
            kartlar.append(dbc.Alert(f"{hedef.name}: {e}", color="danger"))
            continue
        kartlar.append(_sap_ozet_karti(hedef.name, sonuc))
    return html.Div(kartlar)


@app.callback(
    Output("sap-sonuc", "children"),
    Input("sap-aktar-btn", "n_clicks"),
    State("sap-tarih", "value"),
    prevent_initial_call=True,
)
def on_sap_aktar(_n, tarih):
    kaynak = _SAP_STATE.get("kaynak")
    hedefler = _SAP_STATE.get("hedefler") or []
    if not kaynak or not hedefler or not tarih:
        return dbc.Alert("Önce kaynak, hedef(ler) ve tarih seçilmeli.", color="warning")

    basarili, hatali = [], []
    for hedef in hedefler:
        try:
            sonuc = pa.aktar(pa.get_config(), kaynak, hedef, tarih)
            basarili.append(sonuc["output_path"])
        except Exception as e:
            hatali.append(f"{hedef.name}: {e}")

    icerik = []
    if basarili:
        icerik.append(dbc.Alert([
            html.Strong(f"✓ {len(basarili)} dosya aktarıldı:"),
            html.Ul([html.Li(str(p)) for p in basarili]),
            dbc.Button("Çıktı Klasörünü Aç", id="sap-ac-klasor", color="secondary", outline=True, size="sm"),
        ], color="success"))
    for h in hatali:
        icerik.append(dbc.Alert(h, color="danger"))
    return html.Div(icerik)


@app.callback(
    Output("sap-dummy", "data"),
    Input("sap-ac-klasor", "n_clicks"),
    prevent_initial_call=True,
)
def on_sap_ac_klasor(_n):
    el.open_path(pa.OUTPUT_DIR)
    return no_update


# ============================================================ SAYFA: veri kalitesi

ALERT_TIP_ETIKET = {
    "tc_eksik": "TC Eksik",
    "ayni_tc_farkli_isim": "Aynı TC, Farklı İsim",
    "olasi_farkli_tc": "Olası Farklı TC (Aynı İsim)",
    "tc_format_hatali": "TC Format Hatalı",
    "kimlik_bilgisi_eksik": "Kimlik Bilgisi Eksik",
    "tcsiz_saat_ozeti": "TC'siz Kaybolan Saatler",
    "gunluk_kayit_cakismasi": "Günlük Kayıt Çakışması",
    "anormal_saat": "Anormal Saat Değeri",
    "giris_cikis_tutarsiz": "Giriş/Çıkış Tutarsız",
    "zorunlu_alan_eksik": "Zorunlu Alan Eksik",
    "personel_sayisi_anomalisi": "Personel Sayısı Anomalisi",
    "personel_kaydi_yok": "Personel Kaydı Yok",
    "tarihsiz_satir": "Tarihsiz Satır",
    "baslik_eslesmedi": "Başlık Eşleşmedi",
}

# Her uyarı tipi için tabloda gösterilecek EN ANLAMLI sütunlar (genel/gereksiz
# sütunlar yerine, o tipe özel netlik). "detay" her zaman en sona eklenir.
ALERT_TIP_SUTUNLARI = {
    "tc_eksik": [("kaynak_satir", "Satır"), ("ad_soyad", "Ad Soyad")],
    "ayni_tc_farkli_isim": [("tc", "TC"), ("ad_soyad", "Çakışan İsimler")],
    "olasi_farkli_tc": [("tc", "Yeni TC"), ("ad_soyad", "Ad Soyad")],
    "tc_format_hatali": [("kaynak_satir", "Satır"), ("tc", "Hatalı TC"), ("ad_soyad", "Ad Soyad")],
    "kimlik_bilgisi_eksik": [("kaynak_satir", "Satır"), ("tc", "TC"), ("ad_soyad", "Ad Soyad")],
    "tcsiz_saat_ozeti": [("ad_soyad", "Etkilenen İsimler")],
    "gunluk_kayit_cakismasi": [("kaynak_satir", "Satır"), ("tc", "TC"), ("ad_soyad", "Ad Soyad")],
    "anormal_saat": [("kaynak_satir", "Satır"), ("tc", "TC"), ("ad_soyad", "Ad Soyad")],
    "giris_cikis_tutarsiz": [("tc", "TC"), ("ad_soyad", "Ad Soyad")],
    "personel_kaydi_yok": [("tc", "TC"), ("ad_soyad", "Ad Soyad")],
    "tarihsiz_satir": [("kaynak_satir", "Satır"), ("tc", "TC")],
}
_VARSAYILAN_SUTUNLAR = [("kaynak_satir", "Satır"), ("tc", "TC"), ("ad_soyad", "Ad Soyad")]


def _kalite_tablosu(alerts_of_type, tip):
    kolonlar = ALERT_TIP_SUTUNLARI.get(tip, _VARSAYILAN_SUTUNLAR) + [("detay", "Detay")]
    return dash_table.DataTable(
        columns=[{"name": ad, "id": key} for key, ad in kolonlar],
        data=alerts_of_type,
        page_size=15,
        filter_action="native",
        sort_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"textAlign": "left", "padding": "8px", "fontFamily": "Segoe UI", "whiteSpace": "normal"},
        style_header={"backgroundColor": "#1F4E78", "color": "white", "fontWeight": "bold"},
        style_cell_conditional=[{"if": {"column_id": "detay"}, "minWidth": "320px"}],
    )


def sayfa_kalite():
    tum_uyarilar = sorgu.get_quality_alerts(cozuldu=False)
    if not tum_uyarilar:
        return html.Div([
            html.H3("Veri Kalitesi Merkezi", className="fw-bold mb-4"),
            dbc.Alert("Açık veri kalitesi uyarısı yok ✓", color="success"),
        ])

    by_type = {}
    for a in tum_uyarilar:
        by_type.setdefault(a["tip"], []).append(a)
    # en çok uyarısı olan tip en üstte
    siralı_tipler = sorted(by_type.keys(), key=lambda t: -len(by_type[t]))

    tabs = [
        dbc.Tab(
            _kalite_tablosu(by_type[tip], tip),
            label=f"{ALERT_TIP_ETIKET.get(tip, tip.replace('_', ' ').title())} ({len(by_type[tip])})",
            tab_id=tip,
        )
        for tip in siralı_tipler
    ]

    return html.Div([
        html.H3("Veri Kalitesi Merkezi", className="fw-bold mb-2"),
        html.P(f"Toplam {len(tum_uyarilar)} açık uyarı, {len(siralı_tipler)} kategoride. "
               f"Her sekme kendi tipine özel, isim/satır bazında detaylı bilgi gösterir.",
               className="text-muted mb-3"),
        dbc.Tabs(tabs, active_tab=siralı_tipler[0]),
    ])


# ============================================================ SAYFA: puantaj geçmişi

def sayfa_puantaj():
    conn = sema.get_connection()
    gruplar = [r[0] for r in conn.execute(
        "SELECT DISTINCT grup FROM personnel WHERE grup IS NOT NULL ORDER BY grup").fetchall()]
    conn.close()

    return html.Div([
        html.H3("Puantaj Geçmişi Sorgusu", className="fw-bold mb-1"),
        html.P("Kişi başına GÜN başına TEK satır (normal, fazla ve toplam aynı satırda).",
               className="text-muted mb-3"),
        dbc.Row([
            dbc.Col(dcc.Dropdown(id="pt-grup", options=[{"label": g, "value": g} for g in gruplar],
                                   placeholder="Grup seç (opsiyonel)"), md=3),
            dbc.Col(dcc.DatePickerRange(id="pt-tarih", display_format="DD.MM.YYYY"), md=5),
            dbc.Col(dbc.Button("Sorgula", id="pt-sorgula", color="primary"), md=2),
        ], className="g-2 mb-4"),
        dcc.Loading(html.Div(id="pt-sonuc")),
    ])


@app.callback(
    Output("pt-sonuc", "children"),
    Input("pt-sorgula", "n_clicks"),
    State("pt-grup", "value"),
    State("pt-tarih", "start_date"),
    State("pt-tarih", "end_date"),
    prevent_initial_call=True,
)
def on_puantaj_sorgula(_n, grup, start_date, end_date):
    rows = sorgu.get_daily_hours(grup=grup, donem_baslangic=start_date, donem_bitis=end_date)
    if not rows:
        return dbc.Alert("Bu filtrelerle kayıt bulunamadı.", color="warning")
    columns = [{"name": n, "id": k} for k, n in [
        ("ad_soyad", "Ad Soyad"), ("grup", "Grup"), ("bagli_oldugu_ekip", "Alt Ekip"),
        ("gorevi", "Görevi"), ("tarih", "Tarih"), ("durum", "Çalışma Durumu"),
        ("normal_saat", "Normal"), ("fazla_saat", "Fazla"), ("toplam_saat", "Toplam"),
    ]]
    return dash_table.DataTable(
        columns=columns, data=rows[:2000], page_size=25, filter_action="native", sort_action="native",
        style_table={"overflowX": "auto"}, style_cell={"textAlign": "left", "padding": "8px"},
        style_header={"backgroundColor": "#1F4E78", "color": "white", "fontWeight": "bold"},
    )


# ============================================================ SAYFA: personel kayıtları

def sayfa_kayitlar():
    records = pdb.query_personnel()
    columns = [{"name": n, "id": k} for k, n in [
        ("ad_soyad", "Ad Soyad"), ("tc", "TC"), ("grup", "Grup"),
        ("bagli_oldugu_ekip", "Alt Ekip"), ("gorevi", "Görevi"),
        ("ise_giris_tarihi", "Giriş"), ("isten_cikis_tarihi", "Çıkış"),
    ]]
    return html.Div([
        dbc.Row([
            dbc.Col(html.H3("Personel Kayıtları", className="fw-bold mb-0"), width="auto"),
            dbc.Col(dbc.Button("✏️ Düzenle / Yeni Ekle", href="/personel-duzenle", color="primary",
                                 outline=True, size="sm"), width="auto", className="ms-auto"),
        ], className="mb-4 align-items-center"),
        dash_table.DataTable(
            columns=columns, data=records, page_size=25, filter_action="native", sort_action="native",
            style_table={"overflowX": "auto"}, style_cell={"textAlign": "left", "padding": "8px"},
            style_header={"backgroundColor": "#1F4E78", "color": "white", "fontWeight": "bold"},
        ),
    ])


# ============================================================ SAYFA: ekip listesi
# Suite'teki "Ekip Listesi" modülünün karşılığı. Kaynak Excel AÇILMAZ: liste kalıcı
# personel kaydından üretilir. Sıralama ve Excel çıktısı team_report'taki aynı
# fonksiyonlarla yapılır (bkz. veri_merkezi_ekip_listesi.py).

def _secenekler(degerler):
    return [{"label": d, "value": d} for d in degerler]


def _ekip_secenekleri(hiyerarsi, grup_sec):
    return el._sirala_tr({e for g, ekipler in hiyerarsi.items()
                           if not grup_sec or g in grup_sec for e in ekipler})


def _gorev_secenekleri(hiyerarsi, grup_sec, ekip_sec):
    return el._sirala_tr({r for g, ekipler in hiyerarsi.items() if not grup_sec or g in grup_sec
                           for e, roller in ekipler.items() if not ekip_sec or e in ekip_sec
                           for r in roller})


def _filtre_alani(etiket, dropdown_id, secenekler):
    return dbc.Col([
        html.Label(etiket, className="small fw-semibold text-muted mb-1"),
        dcc.Dropdown(id=dropdown_id, options=_secenekler(secenekler), multi=True, placeholder="Tümü"),
    ], md=4)


def sayfa_ekip_listesi():
    hiyerarsi = el.get_hierarchy()
    kalite = el.get_kalite_notu()

    kalite_notu = html.Div()
    if kalite["tcsiz_isim"] or kalite["cakisan_tc"]:
        parcalar = []
        if kalite["tcsiz_isim"]:
            parcalar.append(f"{kalite['tcsiz_isim']} isim TC'siz satırlarda kalıyor (listede yok)")
        if kalite["cakisan_tc"]:
            parcalar.append(f"{kalite['cakisan_tc']} TC birden fazla isimle geliyor (listede tek satır)")
        kalite_notu = dbc.Alert([
            html.Strong("Liste eksik görünebilir: "), "; ".join(parcalar), ". ",
            dcc.Link("Veri Kalitesi'ne git →", href="/kalite"),
        ], color="warning", className="py-2")

    return html.Div([
        html.H3("Ekip Listesi", className="fw-bold mb-1"),
        html.P("Güncel (işten çıkmamış) personel listesi, tarih bağımsız. Sıralama: Formen > Ekip Başı > "
               "diğer görevler (kişi sayısına göre) > Bayrakçı.", className="text-muted mb-3"),
        kalite_notu,
        dbc.Card(dbc.CardBody(dbc.Row([
            _filtre_alani("Grup", "el-grup", sorted(hiyerarsi, key=engine.normalize_name)),
            _filtre_alani("Alt Ekip", "el-ekip", _ekip_secenekleri(hiyerarsi, None)),
            _filtre_alani("Görevi", "el-gorev", _gorev_secenekleri(hiyerarsi, None, None)),
        ], className="g-3")), className="shadow-sm border-0 mb-3"),
        dcc.Loading(html.Div(id="el-onizleme")),
        html.Div([
            dbc.Button("📄 Excel Listesini Oluştur", id="el-olustur", color="success", className="me-2"),
            dbc.Button("Excel'de Aç", id="el-ac-dosya", color="secondary", outline=True,
                       disabled=True, className="me-2"),
            dbc.Button("Çıktı Klasörünü Aç", id="el-ac-klasor", color="secondary", outline=True),
        ], className="mt-3"),
        html.Div(id="el-sonuc", className="mt-3"),
        dcc.Store(id="el-son-dosya"),
        dcc.Store(id="el-dummy"),
    ])


@app.callback(
    Output("el-ekip", "options"),
    Output("el-ekip", "value"),
    Input("el-grup", "value"),
    State("el-ekip", "value"),
)
def on_el_grup_degisti(grup_sec, ekip_sec):
    ekipler = _ekip_secenekleri(el.get_hierarchy(), grup_sec)
    # grup değişince artık geçerli olmayan alt ekip seçimlerini temizle
    return _secenekler(ekipler), [e for e in (ekip_sec or []) if e in ekipler]


@app.callback(
    Output("el-gorev", "options"),
    Output("el-gorev", "value"),
    Input("el-grup", "value"),
    Input("el-ekip", "value"),
    State("el-gorev", "value"),
)
def on_el_ekip_degisti(grup_sec, ekip_sec, gorev_sec):
    gorevler = _gorev_secenekleri(el.get_hierarchy(), grup_sec, ekip_sec)
    return _secenekler(gorevler), [g for g in (gorev_sec or []) if g in gorevler]


@app.callback(
    Output("el-onizleme", "children"),
    Input("el-grup", "value"),
    Input("el-ekip", "value"),
    Input("el-gorev", "value"),
)
def on_el_onizleme(grup_sec, ekip_sec, gorev_sec):
    roster = el.build_roster(grup_sec, ekip_sec, gorev_sec)
    sirali, dagilim = el.sirala_ve_dagilim(roster["records"])
    if not sirali:
        return dbc.Alert("Seçilen filtrelerle eşleşen personel yok.", color="warning")

    satirlar = [{"sira": i, "ad_soyad": r["name"], "tc": r["tc"], "gorevi": r["role"],
                 "grup": r["team"], "alt_ekip": r["subteam"]} for i, r in enumerate(sirali, start=1)]
    tablo = dash_table.DataTable(
        columns=[{"name": n, "id": k} for k, n in [
            ("sira", "Sıra"), ("ad_soyad", "Ad Soyad"), ("tc", "TC Kimlik No"),
            ("gorevi", "Görevi"), ("grup", "Grup"), ("alt_ekip", "Alt Ekip")]],
        data=satirlar, page_size=20, filter_action="native", sort_action="native",
        style_table={"overflowX": "auto"}, style_cell={"textAlign": "left", "padding": "8px"},
        style_header={"backgroundColor": "#1F4E78", "color": "white", "fontWeight": "bold"},
        style_cell_conditional=[{"if": {"column_id": "sira"}, "width": "60px"}],
    )
    dagilim_listesi = dbc.ListGroup([
        dbc.ListGroupItem([html.Span(g), dbc.Badge(str(n), color="primary", className="ms-2 float-end")])
        for g, n in dagilim
    ], style={"maxHeight": "520px", "overflowY": "auto"})

    return html.Div([
        dbc.Row([
            dbc.Col(kpi_card("Toplam Personel", len(sirali), "success"), md=3),
            dbc.Col(kpi_card("Görev Sayısı", len(dagilim), "primary"), md=3),
            dbc.Col(kpi_card("Listeye Alınmayan (Ayrılmış)", roster["excluded_exited"], "secondary"), md=3),
            dbc.Col(kpi_card("Kapsam", el.kapsam_adi(roster), "dark"), md=3),
        ], className="g-3 mb-3"),
        dbc.Row([
            dbc.Col(tablo, md=8),
            dbc.Col([html.H6("Göreve Göre Dağılım", className="fw-semibold"), dagilim_listesi], md=4),
        ], className="g-3"),
    ])


@app.callback(
    Output("el-sonuc", "children"),
    Output("el-son-dosya", "data"),
    Output("el-ac-dosya", "disabled"),
    Input("el-olustur", "n_clicks"),
    State("el-grup", "value"),
    State("el-ekip", "value"),
    State("el-gorev", "value"),
    prevent_initial_call=True,
)
def on_el_olustur(_n, grup_sec, ekip_sec, gorev_sec):
    try:
        yol, roster, kapsam = el.excel_olustur(grup_sec, ekip_sec, gorev_sec)
    except engine.TransferError as e:
        return dbc.Alert(str(e), color="warning"), no_update, no_update
    except PermissionError:
        return (dbc.Alert("Dosya yazılamadı: aynı adlı Excel şu an açık olabilir. Kapatıp tekrar dene.",
                            color="danger"), no_update, no_update)
    except Exception as e:
        return dbc.Alert(f"Liste oluşturulamadı: {e}", color="danger"), no_update, no_update

    return (
        dbc.Alert([
            html.Strong("✓ Ekip listesi oluşturuldu"),
            f" — {len(roster['records'])} personel, kapsam: {kapsam}", html.Br(),
            html.Code(str(yol)),
        ], color="success"),
        str(yol),
        False,
    )


@app.callback(
    Output("el-dummy", "data"),
    Input("el-ac-dosya", "n_clicks"),
    Input("el-ac-klasor", "n_clicks"),
    State("el-son-dosya", "data"),
    prevent_initial_call=True,
)
def on_el_ac(_n_dosya, _n_klasor, son_dosya):
    # Sunucu = kullanıcının kendi bilgisayarı, o yüzden dosya/klasör burada açılır
    # (pywebview'de tarayıcı indirmesine güvenmek yerine).
    if ctx.triggered_id == "el-ac-dosya" and son_dosya:
        el.open_path(son_dosya)
    elif ctx.triggered_id == "el-ac-klasor":
        el.open_path(el.cikti_klasoru())
    return no_update


# ============================================================ SAYFA: personel düzenle

def _personel_secenekleri():
    kayitlar = pdb.query_personnel()
    kayitlar.sort(key=lambda r: engine.normalize_name(r["ad_soyad"]))
    return [{"label": f"{r['ad_soyad']} — {r['tc']}", "value": r["tc"]} for r in kayitlar]


def sayfa_personel_duzenle():
    return html.Div([
        html.H3("Personel Düzenle", className="fw-bold mb-1"),
        html.P("Mevcut bir personeli arayıp tüm bilgilerini düzenleyebilir, ya da "
               "'Yeni Personel Ekle' ile sıfırdan kayıt açabilirsin.", className="text-muted mb-3"),
        dbc.Row([
            dbc.Col(dcc.Dropdown(id="pd-tc-sec", options=_personel_secenekleri(),
                                   placeholder="Düzenlenecek personeli ada veya TC'ye göre ara...",
                                   searchable=True), md=8),
            dbc.Col(dbc.Button("➕ Yeni Personel Ekle", id="pd-yeni", color="secondary", outline=True,
                                 className="w-100"), md=4),
        ], className="g-2 mb-4"),

        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col([html.Label("TC Kimlik No", className="small fw-semibold text-muted"),
                         dcc.Input(id="pd-tc", type="text", className="form-control")], md=4),
                dbc.Col([html.Label("Ad Soyad", className="small fw-semibold text-muted"),
                         dcc.Input(id="pd-ad", type="text", className="form-control")], md=8),
            ], className="g-3 mb-3"),
            dbc.Row([
                dbc.Col([html.Label("Grup", className="small fw-semibold text-muted"),
                         dcc.Input(id="pd-grup", type="text", className="form-control")], md=4),
                dbc.Col([html.Label("Bağlı Olduğu Ekip", className="small fw-semibold text-muted"),
                         dcc.Input(id="pd-ekip", type="text", className="form-control")], md=4),
                dbc.Col([html.Label("Görevi", className="small fw-semibold text-muted"),
                         dcc.Input(id="pd-gorev", type="text", className="form-control")], md=4),
            ], className="g-3 mb-3"),
            dbc.Row([
                dbc.Col([html.Label("İşe Giriş Tarihi", className="small fw-semibold text-muted d-block"),
                         dcc.DatePickerSingle(id="pd-giris", display_format="DD.MM.YYYY",
                                                clearable=True, placeholder="seçilmedi")], md=4),
                dbc.Col([html.Label("İşten Çıkış Tarihi", className="small fw-semibold text-muted d-block"),
                         dcc.DatePickerSingle(id="pd-cikis", display_format="DD.MM.YYYY",
                                                clearable=True, placeholder="hâlâ çalışıyor")], md=4),
            ], className="g-3 mb-3"),
            dbc.Row([
                dbc.Col([html.Label("Notlar", className="small fw-semibold text-muted"),
                         dcc.Textarea(id="pd-notlar", className="form-control", style={"height": "80px"})], md=12),
            ], className="g-3 mb-3"),
            html.Div([
                dbc.Button("💾 Kaydet", id="pd-kaydet", color="success", className="me-2"),
                dbc.Button("🗑 Sil", id="pd-sil", color="danger", outline=True, disabled=True, className="me-2"),
                dbc.Button("Formu Temizle", id="pd-temizle", color="secondary", outline=True),
            ]),
        ]), className="shadow-sm border-0"),

        html.Div(id="pd-sonuc", className="mt-3"),
        dcc.ConfirmDialog(id="pd-sil-onay",
                           message="Bu personeli ve tüm değişiklik geçmişini kalıcı olarak silmek "
                                    "istediğinize emin misiniz? Bu işlem geri alınamaz."),
        dcc.Store(id="pd-mod-store", data="yeni"),
    ])


@app.callback(
    Output("pd-tc", "value"), Output("pd-tc", "disabled"),
    Output("pd-ad", "value"), Output("pd-grup", "value"), Output("pd-ekip", "value"),
    Output("pd-gorev", "value"), Output("pd-giris", "date"), Output("pd-cikis", "date"),
    Output("pd-notlar", "value"), Output("pd-mod-store", "data"),
    Output("pd-tc-sec", "value", allow_duplicate=True),
    Input("pd-tc-sec", "value"),
    Input("pd-yeni", "n_clicks"),
    Input("pd-temizle", "n_clicks"),
    prevent_initial_call="initial_duplicate",
)
def on_pd_form_yukle(tc_secilen, _yeni_click, _temizle_click):
    tetikleyen = ctx.triggered_id
    if tetikleyen in ("pd-yeni", "pd-temizle") or tetikleyen is None:
        # Yeni personel modu (sayfa ilk açıldığında da varsayılan budur): form
        # boş, TC alanı DÜZENLENEBİLİR (yeni kayıt için gerekli), arama kutusu
        # da temizlenir ki eski seçim ekranda asılı kalmasın.
        return "", False, "", "", "", "", None, None, "", "yeni", None

    if not tc_secilen:
        return (no_update,) * 11

    p = pdb.get_personnel(tc_secilen)
    if not p:
        return (no_update,) * 11
    return (p["tc"], True, p["ad_soyad"] or "", p["grup"] or "", p["bagli_oldugu_ekip"] or "",
            p["gorevi"] or "", p["ise_giris_tarihi"] or None, p["isten_cikis_tarihi"] or None,
            p["notlar"] or "", "duzenle", no_update)


@app.callback(
    Output("pd-sil", "disabled"),
    Input("pd-mod-store", "data"),
)
def on_pd_mod_degisti(mod):
    return mod != "duzenle"


@app.callback(
    Output("pd-sonuc", "children"),
    Output("pd-tc-sec", "options"),
    Input("pd-kaydet", "n_clicks"),
    State("pd-mod-store", "data"),
    State("pd-tc", "value"), State("pd-ad", "value"), State("pd-grup", "value"),
    State("pd-ekip", "value"), State("pd-gorev", "value"),
    State("pd-giris", "date"), State("pd-cikis", "date"), State("pd-notlar", "value"),
    prevent_initial_call=True,
)
def on_pd_kaydet(_n, mod, tc, ad, grup, ekip, gorev, giris, cikis, notlar):
    if mod == "yeni":
        gecerli, mesaj = pdb.validate_tc(tc)
        if not gecerli:
            return dbc.Alert(f"TC Kimlik No geçersiz: {mesaj}", color="danger"), no_update
        if pdb.get_personnel(tc):
            return dbc.Alert("Bu TC ile zaten bir personel kayıtlı - onu 'Personel Düzenle' "
                               "arama kutusundan seçip düzenleyebilirsin.", color="warning"), no_update

    if not (ad or "").strip():
        return dbc.Alert("Ad Soyad boş olamaz.", color="danger"), no_update

    kayit = {
        "tc": tc, "ad_soyad": ad, "grup": grup or "", "bagli_oldugu_ekip": ekip or "",
        "gorevi": gorev or "", "ise_giris_tarihi": giris or "", "isten_cikis_tarihi": cikis or "",
        "notlar": notlar or "",
    }
    try:
        sonuc = pdb.upsert_personnel(kayit, source="veri_merkezi:manuel", overwrite_blanks=True)
    except pdb.DatasetError as e:
        return dbc.Alert(str(e), color="danger"), no_update

    etiket = {"inserted": "✓ Yeni personel eklendi.", "updated": "✓ Değişiklikler kaydedildi.",
              "unchanged": "Değişiklik yok - kayıt zaten güncel."}[sonuc]
    return dbc.Alert(etiket, color="success"), _personel_secenekleri()


@app.callback(
    Output("pd-sil-onay", "displayed"),
    Input("pd-sil", "n_clicks"),
    prevent_initial_call=True,
)
def on_pd_sil_iste(_n):
    return True


@app.callback(
    Output("pd-sonuc", "children", allow_duplicate=True),
    Output("pd-tc-sec", "options", allow_duplicate=True),
    Output("pd-tc-sec", "value", allow_duplicate=True),
    Input("pd-sil-onay", "submit_n_clicks"),
    State("pd-tc", "value"),
    prevent_initial_call=True,
)
def on_pd_sil_onaylandi(_n, tc):
    silindi = pdb.delete_personnel(tc)
    mesaj = "✓ Personel silindi." if silindi else "Silinecek kayıt bulunamadı."
    return dbc.Alert(mesaj, color="success" if silindi else "warning"), _personel_secenekleri(), None


# ============================================================ SAYFA: içe aktarma geçmişi

def sayfa_gecmis():
    rows = sorgu.get_import_history()
    columns = [{"name": n, "id": k} for k, n in [
        ("dosya_adi", "Dosya"), ("durum", "Durum"), ("baslama_zamani", "Başlama"),
        ("donem_baslangic", "Dönem Başlangıç"), ("donem_bitis", "Dönem Bitiş"),
        ("yeni_personel", "Yeni Personel"), ("guncellenen_personel", "Güncellenen"),
        ("yazilan_gun_kaydi", "Günlük Kayıt"), ("uyari_sayisi", "Uyarı"),
    ]]
    return html.Div([
        html.H3("İçe Aktarma Geçmişi", className="fw-bold mb-4"),
        dash_table.DataTable(
            columns=columns, data=rows, page_size=20,
            style_table={"overflowX": "auto"}, style_cell={"textAlign": "left", "padding": "8px"},
            style_header={"backgroundColor": "#1F4E78", "color": "white", "fontWeight": "bold"},
            style_data_conditional=[
                {"if": {"filter_query": "{durum} = hata"}, "backgroundColor": "#f8d7da"},
                {"if": {"filter_query": "{durum} = tamamlandi"}, "backgroundColor": "#d4edda"},
            ],
        ),
    ])


# ============================================================ yönlendirme (router)

@app.callback(Output("page-content", "children"), Input("url", "pathname"))
def render_page(pathname):
    pages = {
        "/": sayfa_genel_bakis,
        "/ice-aktar-personel": sayfa_ice_aktar_personel,
        "/ice-aktar-mesai": sayfa_ice_aktar_mesai,
        "/sap-puantaj-aktarim": sayfa_sap_aktarim,
        "/kalite": sayfa_kalite,
        "/puantaj": sayfa_puantaj,
        "/kayitlar": sayfa_kayitlar,
        "/personel-duzenle": sayfa_personel_duzenle,
        "/ekip-listesi": sayfa_ekip_listesi,
        "/gecmis": sayfa_gecmis,
    }
    sayfa_fn = pages.get(pathname, sayfa_genel_bakis)
    return sayfa_fn()


if __name__ == "__main__":
    app.run(debug=True, port=8765)
