# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: Genel Bakış (komuta ekranı)."""

import datetime

from dash import dcc, html, Input, Output, callback
import plotly.graph_objects as go

import veri_merkezi_sorgu as sorgu
import veri_merkezi_ui as ui


def _ay_secenekleri():
    aylar = sorgu.get_covered_periods()
    bu_ay = datetime.date.today().strftime("%Y-%m")
    if bu_ay not in aylar:
        aylar = aylar + [bu_ay]
    aylar = sorted(set(aylar), reverse=True)
    return [{"label": ui.ay_etiketi(a), "value": a} for a in aylar]


def _varsayilan_ay():
    aylar = sorgu.get_covered_periods()
    bu_ay = datetime.date.today().strftime("%Y-%m")
    if not aylar or bu_ay in aylar:
        return bu_ay
    return aylar[-1]


def sayfa():
    ay = _varsayilan_ay()
    secici = html.Div(dcc.Dropdown(id="gb-ay", options=_ay_secenekleri(), value=ay, clearable=False,
                                   searchable=False, style={"width": "170px"}), className="vm-dd")
    return ui.sayfa("Genel Bakış", "Veri Merkezi", html.Div(icerik(ay), id="gb-icerik",
                                                          style={"display": "flex", "flexDirection": "column", "gap": "16px"}),
                    sag=[secici])


@callback(Output("gb-icerik", "children"), Input("gb-ay", "value"), prevent_initial_call=True)
def ay_degisti(ay):
    return icerik(ay or _varsayilan_ay())


# ---------------------------------------------------------------- bölümler

def _akis_seridi(ay, kapsama, toplamlar):
    bugun = datetime.date.today()
    kimlik = sorgu.get_last_import("kimlik")
    mesai = sorgu.get_last_import("mesai")
    eksik = [g["tarih"] for g in kapsama if g["durum"] == "eksik"]

    def hucre(no, baslik, durum_ikon, ikon_cls, satir1, satir2, uyari=False):
        return html.Div([
            html.Span(ui.ikon(durum_ikon), className=f"vm-pipe-ic {ikon_cls}"),
            html.Div([html.Span(f"{no} · {baslik}", className="vm-lbl"), satir1, satir2], className="col",
                     style={"gap": "2px"}),
        ], className="warn" if uyari else "")

    if kimlik:
        k1 = html.Span(f"Güncel — {ui.tarih_tr(kimlik['baslama_zamani'])}", className="vm-pipe-t")
        k2 = html.Span(f"{kimlik['dosya_adi']} · {kimlik['yeni_personel'] or 0} yeni", className="vm-sub")
        c1 = hucre(1, "Kimlik verisi", "check-lg", "pi-ok", k1, k2)
    else:
        c1 = hucre(1, "Kimlik verisi", "exclamation-lg", "pi-warn", html.Span("Henüz yüklenmedi", className="vm-pipe-t c-orange"),
                   dcc.Link("Personel dosyası yükle →", href="/ice-aktar", className="small-12"), True)

    if eksik:
        gosterilen = " · ".join(d.strftime("%d.%m") for d in eksik[:4]) + (" …" if len(eksik) > 4 else "")
        c2 = hucre(2, "Mesai verisi", "exclamation-lg", "pi-warn",
                   html.Span(f"{len(eksik)} gün eksik", className="vm-pipe-t c-orange"),
                   html.Span(gosterilen, className="vm-sub"), True)
    elif mesai:
        c2 = hucre(2, "Mesai verisi", "check-lg", "pi-ok", html.Span("Eksik gün yok", className="vm-pipe-t"),
                   html.Span(f"son: {ui.tarih_tr(mesai['donem_bitis'] or mesai['baslama_zamani'])}", className="vm-sub"))
    else:
        c2 = hucre(2, "Mesai verisi", "exclamation-lg", "pi-warn", html.Span("Henüz yüklenmedi", className="vm-pipe-t c-orange"),
                   dcc.Link("SAP mesai dosyası yükle →", href="/ice-aktar", className="small-12"), True)

    c3 = hucre(3, "Veri kalitesi", "shield-exclamation" if toplamlar["acik"] else "shield-check",
               "pi-red" if toplamlar["kritik"] else ("pi-warn" if toplamlar["acik"] else "pi-ok"),
               html.Span([f"{ui.sayi(toplamlar['acik'])} açık · ",
                          html.Span(f"{toplamlar['kritik']} kritik", className="c-red" if toplamlar["kritik"] else "")],
                         className="vm-pipe-t"),
               dcc.Link("Kontrol et →", href="/kalite", className="small-12", style={"fontWeight": 600}))

    hazir = not eksik and not toplamlar["kritik"] and mesai is not None
    c4 = hucre(4, "Raporlar", "bar-chart", "pi-navy",
               html.Span("Hazır" if hazir else "Eksikler kapanınca hazır", className="vm-pipe-t"),
               html.Span([dcc.Link("Verimlilik", href="/verimlilik"), " · ", dcc.Link("Maliyet", href="/maliyet"),
                          " · ", dcc.Link("Ekip Listesi", href="/ekip-listesi")], className="small-12"))
    return html.Section([c1, c2, c3, c4], className="vm-card vm-pipe")


def _kpiler(ay, ozet, genel, toplamlar):
    ay_adi = ui.ay_etiketi(ay).split()[0]
    oran = (ozet["fazla"] / ozet["toplam"] * 100) if ozet["toplam"] else 0
    return html.Section([
        ui.kpi("Toplam personel", ui.sayi(genel["toplam_personel"]), None, "veri setinde kayıtlı", True),
        ui.kpi("Aktif personel", ui.sayi(genel["aktif_personel"]), "green",
               f"{ui.sayi(genel['toplam_personel'] - genel['aktif_personel'])} ayrılmış", True),
        ui.kpi("Toplam mesai", ui.sayi(ozet["toplam"]), None, f"saat · {ay_adi}", True),
        ui.kpi("Fazla mesai", ui.sayi(ozet["fazla"]), "orange", f"saat · %{ui.sayi(oran, 1)}", True),
        ui.kpi("Yevmiye günü", ui.sayi(ozet["yevmiye"], 1), "navy", "FM ×1,5 · hak edilen pazar ×2,5", True),
        ui.kpi("Kritik uyarı", ui.sayi(toplamlar["kritik"]), "red" if toplamlar["kritik"] else "green",
               f"{ui.sayi(toplamlar['acik'] - toplamlar['kritik'])} diğer uyarı", True),
    ], className="g g-6 gap-12")


def _grup_cubuklari(ay, bas, bit):
    gruplar = sorgu.get_group_breakdown(bas, bit)
    if not gruplar:
        govde = ui.bos("Bu ay için puantaj kaydı yok. 'İçe Aktar' sayfasından SAP mesai dosyası yükleyin.")
    else:
        mx = max(g["normal"] + g["fazla"] for g in gruplar) or 1
        satirlar = []
        for g in gruplar:
            top = g["normal"] + g["fazla"]
            satirlar.append(html.Div([
                html.Span(g["grup"], style={"fontSize": "13px", "fontWeight": 600}),
                ui.bar([(g["normal"] / mx * 100, "bg-navy"), (g["fazla"] / mx * 100, "bg-orange")]),
                html.Span([f"{ui.sayi(top)} sa · ",
                           html.Span(f"%{ui.sayi(g['fazla'] / top * 100 if top else 0)} FM", className="c-orange")],
                          className="mono", style={"fontSize": "12px", "textAlign": "right"}),
            ], className="vm-barrow"))
        govde = html.Div(satirlar, className="col", style={"gap": "10px"})
    return ui.kart([
        html.Div([ui.h2(f"Grup bazında mesai · {ui.ay_etiketi(ay).split()[0]}", buyuk=True),
                  ui.lejant("#1F4E78", "Normal"), ui.lejant("#E8762C", "Fazla")], className="vm-card-head"),
        govde,
    ], "pad-l span-8 col", style={"gap": "14px"})


def _kapsama_takvimi(ay, kapsama):
    ilk = kapsama[0]["tarih"]
    bugun = datetime.date.today()
    hucreler = [html.Span(g, className="vm-cal-h") for g in ui.GUN_KISA]
    hucreler += [html.Div(className="vm-cal-d cd-bos") for _ in range(ilk.weekday())]
    for g in kapsama:
        cls = f"vm-cal-d cd-{g['durum']}" + (" cd-bugun" if g["tarih"] == bugun else "")
        baslik = f"{g['tarih']:%d.%m.%Y} · {ui.sayi(g['kayit'])} kayıt"
        hucreler.append(html.Div(str(g["tarih"].day), className=cls, title=baslik))
    eksik = sum(1 for g in kapsama if g["durum"] == "eksik")
    return ui.kart([
        html.Div([ui.h2("Veri kapsama takvimi", buyuk=True), html.Span(ui.ay_etiketi(ay), className="small-12 muted")],
                 className="vm-card-head"),
        html.Div(hucreler, className="vm-cal"),
        html.Div([ui.lejant("#1F4E78", "Tam"), ui.lejant("#8FAACB", "Kısmi"),
                  html.Span([html.I(style={"background": "#fff", "border": "1.5px dashed #B8410C"}), "Eksik gün"],
                            className="vm-legend")], className="row", style={"flexWrap": "wrap", "gap": "10px 14px"}),
        dcc.Link([ui.ikon("upload"), f"Eksik {eksik} günü içe aktar"], href="/ice-aktar", className="vm-btn sig")
        if eksik else ui.banner("Bu ayda geçmiş günlerin hepsi için mesai verisi var.", "ok", "check2-circle"),
    ], "pad-l span-4 col", style={"gap": "12px"})


def _trend():
    trend = sorgu.get_monthly_trend(12)
    if not trend:
        govde = ui.bos("Trend için en az bir ay veri gerekli.")
    else:
        x = [f"{ui.AY_KISA[int(t['ay'][5:]) - 1]} {t['ay'][2:4]}" for t in trend]
        fig = go.Figure()
        fig.add_scatter(x=x, y=[t["normal"] or 0 for t in trend], name="Normal", mode="lines+markers",
                        line=dict(color="#1F4E78", width=2.5), marker=dict(size=6))
        fig.add_scatter(x=x, y=[t["fazla"] or 0 for t in trend], name="Fazla", mode="lines+markers",
                        line=dict(color="#E8762C", width=2.5, dash="dash"), marker=dict(size=6))
        fig.update_layout(template="plotly_white", height=230, margin=dict(t=10, b=30, l=50, r=10),
                          font=dict(family="IBM Plex Sans, Segoe UI, sans-serif", size=12, color="#5A6572"),
                          legend=dict(orientation="h", y=1.12, x=0), hovermode="x unified",
                          separators=",.", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        fig.update_yaxes(gridcolor="#EFEBE4", ticksuffix=" sa")
        govde = dcc.Graph(figure=fig, config={"displayModeBar": False})
    return ui.kart([
        html.Div([ui.h2("Aylık mesai trendi", buyuk=True), html.Span("son 12 ay", className="small-12 muted")],
                 className="vm-card-head"),
        govde,
    ], "pad-l span-7 col", style={"gap": "10px"})


def _uyarilar(sayimlar, toplam):
    satirlar = [dcc.Link([
        ui.dot(a["seviye"]), html.Span(a["etiket"], className="grow"),
        html.Span(ui.sayi(a["adet"]), className="mono", style={"fontWeight": 600}),
    ], href=f"/kalite/{a['tip']}", className="row", style={
        "padding": "10px 20px", "borderTop": "1px solid #EFEBE4", "textDecoration": "none", "color": "#16202B"})
        for a in sayimlar[:6]]
    if not satirlar:
        satirlar = [html.Div(ui.banner("Açık veri kalitesi uyarısı yok.", "ok", "check2-circle"),
                             style={"padding": "0 20px 16px"})]
    return html.Section([
        html.Div([ui.h2("Açık uyarılar", buyuk=True), dcc.Link(f"Tümü ({ui.sayi(toplam)}) →", href="/kalite",
                                                             style={"fontSize": "13px", "fontWeight": 600})],
                 className="vm-card-head", style={"padding": "16px 20px 10px"}),
        *satirlar,
    ], className="vm-card span-5", style={"display": "flex", "flexDirection": "column"})


def icerik(ay):
    bas, bit = sorgu.ay_araligi(ay)
    ozet = sorgu.get_period_summary(bas, bit)
    genel = sorgu.get_dashboard_summary(saatler=False)   # yalnız personel sayıları kullanılıyor
    sayimlar = sorgu.get_alert_counts(False)
    toplamlar = {"acik": sum(a["adet"] for a in sayimlar),
                 "kritik": sum(a["adet"] for a in sayimlar if a["seviye"] == "kritik")}
    kapsama = sorgu.gun_kapsamasi(bas, bit)
    return [
        _akis_seridi(ay, kapsama, toplamlar),
        _kpiler(ay, ozet, genel, toplamlar),
        html.Div([_grup_cubuklari(ay, bas, bit), _kapsama_takvimi(ay, kapsama)], className="g-12"),
        html.Div([_trend(), _uyarilar(sayimlar, toplamlar["acik"])], className="g-12"),
    ]
