# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: Personel Maliyet Raporu.

Puantaj Suite'teki karşılığının Veri Merkezi sürümüdür: personel kapsamı kalıcı
depodan okunur (veri_merkezi_maliyet). Excel çıktısı cost_report biçimindedir.
(Eski "Fazla Mesai Raporu" artık veri_merkezi_sayfa_verimlilik.py: Verimlilik Analizi.)"""

import datetime

from dash import dcc, html, Input, Output, State, no_update, callback
import dash_bootstrap_components as dbc
import plotly.graph_objects as go

import personnel_db as pdb
import veri_merkezi_ekip_listesi as el
import veri_merkezi_maliyet as ml
import veri_merkezi_sorgu as sorgu
import veri_merkezi_ui as ui
from veri_merkezi_sayfa_ice_aktar import gecici_kaydet
from veri_merkezi_sayfa_puantaj import _varsayilan_aralik, _hiyerarsi, _ekipler, _gorevler

FONT = dict(family="IBM Plex Sans, Segoe UI, sans-serif", size=12, color="#5A6572")


def _kapsam_filtreleri(onek):
    h = _hiyerarsi()
    return [
        ui.alan("Grup", ui.dropdown(f"{onek}-grup", sorted(h, key=ui.tr_upper), [], coklu=True)),
        ui.alan("Alt ekip", ui.dropdown(f"{onek}-ekip", _ekipler(h, None), [], coklu=True)),
        ui.alan("Görevi", ui.dropdown(f"{onek}-gorev", _gorevler(h, None, None), [], coklu=True)),
    ]


def _zincir_kaydet(onek):
    @callback(
        Output(f"{onek}-ekip", "options"), Output(f"{onek}-ekip", "value"),
        Output(f"{onek}-gorev", "options"), Output(f"{onek}-gorev", "value"),
        Input(f"{onek}-grup", "value"), Input(f"{onek}-ekip", "value"),
        State(f"{onek}-gorev", "value"),
        prevent_initial_call=True,
    )
    def _zincir(gruplar, ekipler, gorevler):
        h = _hiyerarsi()
        eks = _ekipler(h, gruplar)
        ekipler = [e for e in (ekipler or []) if e in eks]
        gors = _gorevler(h, gruplar, ekipler)
        return ([{"label": e, "value": e} for e in eks], ekipler,
                [{"label": g, "value": g} for g in gors], [g for g in (gorevler or []) if g in gors])
    return _zincir


# ============================================================ MALİYET RAPORU

_ML = {"yol": None, "sayfalar": [], "veri": None}


def _sayfa_listesi():
    if not _ML["sayfalar"]:
        return html.Div([html.Span("Hakediş dosyası yükleyince bordro sayfaları otomatik bulunur.", className="small-12 muted"),
                         dbc.Checklist(id="ml-sayfalar", options=[], value=[], className="vm-facet")])
    opts, secili = [], []
    for s in _ML["sayfalar"]:
        etiket = html.Span([
            html.Span([html.Span(s["sayfa"], className="mono", style={"fontWeight": 600, "fontSize": "13px"}),
                       html.Span("sütun düzeni farklı — başlıktan eşleştirildi", className="small-12 c-orange")
                       if s["kayik"] else None], className="col", style={"gap": 0}),
            html.Span(f"{ui.sayi(s['kisi'])} kişi" if s["bordro"] else "bordro değil", className="fc-n"),
        ], style={"display": "flex", "width": "100%", "gap": "8px"})
        opts.append({"label": etiket, "value": s["sayfa"], "disabled": not s["bordro"]})
        if s["bordro"]:
            secili.append(s["sayfa"])
    return dbc.Checklist(id="ml-sayfalar", options=opts, value=secili, className="vm-facet")


def sayfa_maliyet():
    dosya = _ML["yol"]
    return ui.sayfa("Personel Maliyet Raporu", "Raporlar / Maliyet", html.Div([
        html.Div([
            ui.lbl("Hakediş / bordro dosyası"),
            ui.yukleme_alani("ml-upload", "Hakediş Excel'ini seç", kucuk=True),
            html.Div(_dosya_bilgi(), id="ml-dosya"),
            ui.lbl("Taranacak sayfalar"),
            html.Div(_sayfa_listesi(), id="ml-sayfa-kutu"),
            ui.lbl("Personel kapsamı"),
            *_kapsam_filtreleri("ml"),
            html.Div([ui.bilgi_satiri("Anormal birim ücret eşiği", ui.tl(ml.cost.CONFIG["unit_rate_sanity_ceiling"]), True),
                      ui.bilgi_satiri("Eşleştirme anahtarı", "TC Kimlik No", True)], className="col", style={"gap": "4px"}),
            html.Div([
                html.Button([ui.ikon("calculator"), "Maliyeti hesapla"], id="ml-hesapla", className="vm-btn pri", n_clicks=0),
                html.Button([ui.ikon("file-earmark-excel"), "Excel raporunu oluştur"], id="ml-excel", className="vm-btn",
                            n_clicks=0),
            ], className="col", style={"gap": "8px", "marginTop": "auto"}),
        ], className="vm-pane-l", style={"gap": "12px"}),
        html.Div([html.Div(id="ml-mesaj"),
                  dcc.Loading(html.Div(_maliyet_gorunumu(_ML["veri"]) if _ML["veri"] else ui.bos(
                      "Hakediş dosyasını yükleyip ‘Maliyeti hesapla’ya basın. Kapsam boş bırakılırsa veri setindeki "
                      "herkes TC ile eşleştirilir."), id="ml-sonuc", className="col", style={"gap": "16px"}),
                              type="default", color="#1F4E78")], className="vm-pane-c"),
    ], className="vm-split", style={"gridTemplateColumns": "340px minmax(0,1fr)"}), govde_cls=None)


_zincir_kaydet("ml")


def _dosya_bilgi():
    if not _ML["yol"]:
        return None
    bordro = sum(1 for s in _ML["sayfalar"] if s["bordro"])
    return html.Div([
        html.Span(ui.ikon("file-earmark-spreadsheet"), className="vm-pipe-ic pi-navy",
                  style={"width": "34px", "height": "34px", "borderRadius": "8px"}),
        html.Div([html.Span(_ML["yol"].name, className="mono", style={"fontSize": "13px", "fontWeight": 600, "wordBreak": "break-all"}),
                  html.Span(f"{len(_ML['sayfalar'])} sayfa · {bordro}'ü bordro sayfası", className="small-12 muted")],
                 className="col", style={"gap": 0}),
    ], className="row", style={"gap": "10px", "padding": "10px 12px", "border": "1px solid #DDD8CF", "borderRadius": "8px",
                               "background": "#fff"})


@callback(
    Output("ml-dosya", "children"),
    Output("ml-sayfa-kutu", "children"),
    Output("ml-mesaj", "children", allow_duplicate=True),
    Input("ml-upload", "contents"),
    State("ml-upload", "filename"),
    prevent_initial_call=True,
)
def ml_yuklendi(contents, filename):
    if not contents:
        return no_update, no_update, no_update
    yol = gecici_kaydet(contents, filename)
    try:
        sayfalar = ml.sayfalari_tara(yol)
    except Exception as e:
        return None, _sayfa_listesi(), ui.banner(f"Hakediş dosyası okunamadı: {e}", "err")
    _ML.update(yol=yol, sayfalar=sayfalar, veri=None)
    uyari = None if any(s["bordro"] for s in sayfalar) else \
        ui.banner("Dosyada T.C. ve ADI SOYADI başlıklı bir bordro sayfası bulunamadı.", "warn")
    return _dosya_bilgi(), _sayfa_listesi(), uyari


def _maliyet_gorunumu(d):
    o = d["ozet"]
    top = o["toplam"] or 1
    kpiler = html.Section([
        ui.kpi("Toplam maliyet", ui.sayi(o["toplam"]), None, "TL"),
        ui.kpi("Personele ödenen", ui.sayi(o["personel"]), "navy", "TL"),
        ui.kpi("SGK ödemesi", ui.sayi(o["sgk"]), "orange", "TL"),
        ui.kpi("Firma payı", ui.sayi(o["firma"]), "purple", "TL"),
        ui.kpi("Eşleşen personel", f"{ui.sayi(o['eslesen'])}", None,
               f"{ui.sayi(o['bulunamayan'])} puantajlı kişi hakedişte yok"),
    ], className="g g-5 gap-12")

    pie = go.Figure(go.Pie(labels=["Personele ödenen", "SGK ödemesi", "Firma payı"],
                           values=[o["personel"], o["sgk"], o["firma"]], hole=0.62, sort=False,
                           marker=dict(colors=["#1F4E78", "#E8762C", "#7B4FA8"]), textinfo="percent",
                           hovertemplate="%{label}: %{value:,.0f} TL<extra></extra>"))
    pie.update_layout(height=230, margin=dict(t=0, b=0, l=0, r=0), showlegend=True, font=FONT, separators=",.",
                      legend=dict(orientation="h", y=-0.05), paper_bgcolor="rgba(0,0,0,0)")
    dagilim = ui.kart([ui.lbl("Maliyet dağılımı"), dcc.Graph(figure=pie, config={"displayModeBar": False})],
                      "pad col", style={"gap": "8px"})

    mx = max([e["toplam"] for e in o["ekipler"]] + [1])
    ekip = ui.kart([
        html.Div([ui.h2(f"{o['kirilim']} bazlı maliyet"), ui.lejant("#1F4E78", "Personel"), ui.lejant("#E8762C", "SGK"),
                  ui.lejant("#7B4FA8", "Firma payı")], className="vm-card-head"),
        *[html.Div([html.Span(e["ad"], style={"fontSize": "13px", "fontWeight": 600, "overflow": "hidden",
                                              "textOverflow": "ellipsis", "whiteSpace": "nowrap"}, title=e["ad"]),
                    ui.bar([(e["personel"] / mx * 100, "bg-navy"), (e["sgk"] / mx * 100, "bg-orange"),
                            (e["firma"] / mx * 100, "bg-purple")], "mid"),
                    html.Span(ui.tl(e["toplam"]), className="mono", style={"fontSize": "12px", "textAlign": "right"})],
                   className="g", style={"gridTemplateColumns": "140px minmax(0,1fr) 130px", "gap": "12px",
                                         "alignItems": "center"}) for e in o["ekipler"][:12]],
    ], "pad-l col", style={"gap": "10px"})

    gmx = max([g["kisi_basi"] for g in o["gorevler"]] + [1])
    gorev = html.Section([
        html.Div([ui.h2("Görev bazlı maliyet")], className="vm-card-head pad"),
        ui.tablo(["Görevi", ("Kişi", "r"), ("Toplam maliyet", "r"), ("Kişi başı", "r"), ""], [[
            html.Td(g["ad"], style={"fontWeight": 600, "fontSize": "13px"}),
            html.Td(ui.sayi(g["kisi"]), className="mono r"),
            html.Td(ui.tl(g["toplam"]), className="mono r"),
            html.Td(ui.tl(g["kisi_basi"]), className="mono r", style={"fontWeight": 600}),
            html.Td(ui.bar([(g["kisi_basi"] / gmx * 100, "bg-navy")], "thin"), style={"width": "90px", "minWidth": "60px"}),
        ] for g in o["gorevler"][:15]]),
    ], className="vm-card clip")

    kontroller = [("Hakedişte bulunamayan (puantajı olan) personel", o["bulunamayan"], "kritik")]
    for ad, n in o["flag_sayilari"].items():
        kontroller.append((f"{ad}" + (f" (> {ui.tl(o['esik'])})" if ad.startswith("Yevmiye") else ""), n, "uyari"))
    kontroller.append(("Aynı TC birden fazla sayfada", o["tekrar_tc"], "bilgi"))
    if o["kayitsiz_hakedis"]:
        kontroller.append(("Hakedişte olup veri setinde kaydı olmayan", o["kayitsiz_hakedis"], "bilgi"))
    arka = {"kritik": "#FDF6F5", "uyari": "#FDF3EA", "bilgi": "#F4F2EE"}
    kontrol = ui.kart([ui.lbl("Kontrol listesi"),
                       *[html.Div([ui.dot(s if n else "bilgi"), html.Span(l, className="grow"),
                                   html.Span(ui.sayi(n), className="mono", style={"fontWeight": 600})],
                                  className="row", style={"gap": "10px", "padding": "9px 10px", "borderRadius": "8px",
                                                          "background": arka[s] if n else "#F4F2EE", "fontSize": "13px"})
                         for l, n, s in kontroller],
                       html.Span("Uyarılı satırlar Excel'de işaretlenir; bulunamayanlar ayrı sayfada listelenir.",
                                 className="small-12 muted")], "pad col", style={"gap": "8px"})

    return [kpiler,
            html.Div([dagilim, ekip], className="g", style={"gridTemplateColumns": "300px minmax(0,1fr)"}),
            html.Div([gorev, kontrol], className="g", style={"gridTemplateColumns": "minmax(0,1fr) 340px"})]


def _ml_hesapla(sayfalar, grup, ekip, gorev):
    if not _ML["yol"]:
        raise ValueError("Önce hakediş dosyasını yükleyin.")
    if not sayfalar:
        raise ValueError("En az bir bordro sayfası seçin.")
    d = ml.maliyet_verisi(_ML["yol"], sayfalar, grup or None, ekip or None, gorev or None)
    if not d["merged"]:
        raise ValueError("Seçilen kapsamdaki personelden hiçbiri hakediş dosyasında (TC ile) bulunamadı.")
    _ML["veri"] = d
    return d


@callback(
    Output("ml-sonuc", "children"),
    Output("ml-mesaj", "children"),
    Input("ml-hesapla", "n_clicks"),
    State("ml-sayfalar", "value"), State("ml-grup", "value"), State("ml-ekip", "value"), State("ml-gorev", "value"),
    prevent_initial_call=True,
)
def ml_goster(n, sayfalar, grup, ekip, gorev):
    if not n:
        return no_update, no_update
    try:
        d = _ml_hesapla(sayfalar, grup, ekip, gorev)
    except Exception as e:
        return no_update, ui.banner(str(e), "warn")
    return _maliyet_gorunumu(d), None


@callback(
    Output("ml-mesaj", "children", allow_duplicate=True),
    Input("ml-excel", "n_clicks"),
    State("ml-sayfalar", "value"), State("ml-grup", "value"), State("ml-ekip", "value"), State("ml-gorev", "value"),
    prevent_initial_call=True,
)
def ml_excel(n, sayfalar, grup, ekip, gorev):
    if not n:
        return no_update
    try:
        d = _ml_hesapla(sayfalar, grup, ekip, gorev)
        yol = ml.excel_yaz(d)
    except PermissionError:
        return ui.banner("Dosya yazılamadı: aynı adlı Excel şu an açık olabilir.", "err")
    except Exception as e:
        return ui.banner(f"Rapor oluşturulamadı: {e}", "err")
    el.open_path(yol)
    return ui.banner(["Rapor oluşturuldu: ", html.Code(str(yol))], "ok", "check2-circle")
