# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: Ekip Listesi.

Liste kalıcı personel kaydından üretilir (veri_merkezi_ekip_listesi);
sıralama ve Excel çıktısı team_report ile aynıdır. Ekranda Excel çıktısının
birebir önizlemesi (görev gruplarına bölünmüş) gösterilir.

Aynı kapsam için sahada doldurulacak A4 yatay GÜNLÜK MESAİ FORMU da buradan
PDF olarak üretilir (veri_merkezi_mesai_formu): her alt ekip kendi formunu
alır, sayfa başına en fazla 20 kişi."""

from dash import dcc, html, Input, Output, State, ALL, no_update, callback, ctx

import dash_bootstrap_components as dbc

import puantaj_engine as engine
import team_report as report
import veri_merkezi_ekip_listesi as el
import veri_merkezi_ekip_excel as ex
import veri_merkezi_mesai_formu as mf
import veri_merkezi_ui as ui

ONIZLEME_ILK = 60      # ilk açılışta gösterilen satır (PERFORMANS: 1000+ satır tarayıcıyı yavaşlatır)
ONIZLEME_MAX = 1500    # "Tümünü göster" ile en fazla


def _ekip_opts(h, grup):
    return el._sirala_tr({e for g, ek in h.items() if not grup or g in grup for e in ek})


def _gorev_opts(h, grup, ekip):
    return el._sirala_tr({r for g, ek in h.items() if not grup or g in grup
                          for e, roller in ek.items() if not ekip or e in ekip for r in roller})


def sayfa():
    h = el.get_hierarchy()
    k = el.get_kalite_notu()
    uyari = None
    if k["tcsiz_isim"] or k["cakisan_tc"]:
        parca = []
        if k["tcsiz_isim"]:
            parca.append(f"{k['tcsiz_isim']} isim TC'siz satırlarda kalıyor (listede yok)")
        if k["cakisan_tc"]:
            parca.append(f"{k['cakisan_tc']} TC birden fazla isimle geliyor (listede tek satır)")
        uyari = ui.banner([html.Strong("Liste eksik görünebilir: "), "; ".join(parca), ". ",
                           dcc.Link("Veri Kalitesi'ne git →", href="/kalite", style={"fontWeight": 600})], "warn")
    return ui.sayfa("Ekip Listesi", "Kayıtlar / Ekip Listesi", [
        ui.kart(html.Div([
            ui.alan("Grup", ui.dropdown("el-grup", sorted(h, key=engine.normalize_name), [], coklu=True)),
            ui.alan("Alt ekip", ui.dropdown("el-ekip", _ekip_opts(h, None), [], coklu=True)),
            ui.alan("Görevi", ui.dropdown("el-gorev", _gorev_opts(h, None, None), [], coklu=True)),
        ], className="g g-3", style={"gap": "14px"}), "pad"),
        uyari,
        dcc.Loading(html.Div(id="el-onizleme"), type="default", color="#1F4E78"),
        dcc.Store(id="el-son-dosya"),
        dcc.Store(id="el-form-dosya"),
        dcc.Store(id="el-tumu", data=False),
        # Gizlenen personelin TC listesi. Sadece bu ekranda geçerli: sayfadan çıkınca sıfırlanır,
        # veritabanına yazılmaz. Gizlenenler Excel listesine ve mesai formuna girmez.
        dcc.Store(id="el-gizli", data=[]),
        dcc.Store(id="el-dummy"),
    ])


@callback(
    Output("el-ekip", "options"), Output("el-ekip", "value"),
    Output("el-gorev", "options"), Output("el-gorev", "value"),
    Input("el-grup", "value"), Input("el-ekip", "value"),
    State("el-gorev", "value"),
    prevent_initial_call=True,
)
def zincir(grup, ekip, gorev):
    h = el.get_hierarchy()
    eks = _ekip_opts(h, grup)
    ekip = [e for e in (ekip or []) if e in eks]
    gors = _gorev_opts(h, grup, ekip)
    return ([{"label": e, "value": e} for e in eks], ekip, [{"label": g, "value": g} for g in gors],
            [g for g in (gorev or []) if g in gors])


@callback(
    Output("el-onizleme", "children"),
    Input("el-grup", "value"), Input("el-ekip", "value"), Input("el-gorev", "value"),
    Input("el-son-dosya", "data"),
    Input("el-tumu", "data"),
    Input("el-gizli", "data"),
)
def onizleme(grup, ekip, gorev, son_dosya, tumu=False, gizli=None):
    roster = el.build_roster(grup, ekip, gorev)
    tum_sirali, _ = el.sirala_ve_dagilim(roster["records"])
    if not tum_sirali:
        return ui.banner("Seçilen filtrelerle eşleşen (işten çıkmamış) personel yok.", "warn")
    gizli_set = set(gizli or [])
    sirali, dagilim = el.sirala_ve_dagilim(_gorunenler(roster["records"], gizli_set))
    gizli_burada = [r for r in tum_sirali if r["tc"] in gizli_set]
    kapsam = el.kapsam_adi(roster)

    satirlar = [html.Div([html.Span(b) for b in ["Sıra", "Ad soyad", "TC kimlik no", "Görevi", "Alt ekip", ""]],
                         className="vm-roster-row h")]
    sayim = dict(dagilim)
    onceki = None
    sinir = ONIZLEME_MAX if tumu else ONIZLEME_ILK
    sira = 0
    for r in tum_sirali[:sinir]:
        if r["role"] != onceki:
            satirlar.append(html.Div(f"{r['role']} · {sayim.get(r['role'], 0)}", className="vm-roster-grp"))
            onceki = r["role"]
        saklandi = r["tc"] in gizli_set
        if not saklandi:
            sira += 1
        satirlar.append(html.Div([
            html.Span("—" if saklandi else str(sira), className="mono muted"),
            html.Span(r["name"], style={"fontWeight": 600}),
            html.Span(ui.maske_tc(r["tc"]), className="mono", style={"fontSize": "13px"}),
            html.Span(r["role"], style={"fontSize": "13px"}), html.Span(r["subteam"], style={"fontSize": "13px"}),
            html.Button(ui.ikon("eye-slash" if saklandi else "eye"), id={"type": "el-goz", "tc": r["tc"]}, n_clicks=0,
                        className="vm-eyebtn", title="Listeye geri al" if saklandi else "Listeden gizle",
                        **{"aria-label": f"{r['name']} — " + ("göster" if saklandi else "gizle")}),
        ], className="vm-roster-row gizli" if saklandi else "vm-roster-row"))
    if len(tum_sirali) > sinir:
        kalan = len(tum_sirali) - sinir
        satirlar.append(html.Div([
            html.Span(f"… {ui.sayi(kalan)} personel daha (Excel ve mesai formunda tamamı var)", className="muted"),
            html.Button(f"Tümünü göster ({ui.sayi(len(tum_sirali))})", id="el-tumu-btn", n_clicks=0, className="vm-btn sm")
            if not tumu and len(tum_sirali) <= ONIZLEME_MAX else None,
        ], className="row", style={"justifyContent": "center", "padding": "10px", "fontSize": "13px", "gap": "12px"}))

    mx = max([n for _g, n in dagilim] + [1])
    dagilim_satir = [html.Div([html.Span(g, style={"fontSize": "13px", "overflow": "hidden", "textOverflow": "ellipsis",
                                                    "whiteSpace": "nowrap"}),
                               ui.bar([(n / mx * 100, "bg-navy")], "thin"),
                               html.Span(ui.sayi(n), className="mono", style={"textAlign": "right", "fontSize": "13px"})],
                              className="g", style={"gridTemplateColumns": "120px minmax(0,1fr) 34px", "gap": "10px",
                                                    "alignItems": "center"})
                     for g, n in dagilim]

    kagit = html.Section([
        html.Div([html.Div([ui.lbl("Excel önizleme · Personel Listesi"),
                            html.Span(f"{kapsam} — EKİP LİSTESİ", className="disp",
                                      style={"fontSize": "20px", "fontWeight": 800, "color": "#16344F"})],
                           className="col grow", style={"gap": "2px"}),
                  html.Span(f"{ui.sayi(len(sirali))} personel", className="vm-sub")], className="vm-paper-head"),
        ui.banner([html.Strong(f"{len(gizli_burada)} kişi gizlendi. "),
                   "Excel listesine ve mesai formuna girmeyecek. ",
                   html.Button("Hepsini geri al", id="el-gizli-sifirla", n_clicks=0, className="vm-btn xs",
                               style={"marginLeft": "6px"})], "info", "eye-slash") if gizli_burada else None,
        html.Div(satirlar, style={"maxHeight": "760px", "overflowY": "auto"}),
    ], className="vm-card vm-paper")

    sag = html.Div([
        html.Section([ui.kpi("Listedeki personel", ui.sayi(len(sirali)), "green",
                             f"{len(gizli_burada)} gizli" if gizli_burada else None),
                      ui.kpi("Görev sayısı", ui.sayi(len(dagilim)), "navy"),
                      ui.kpi("Ayrılmış (hariç)", ui.sayi(roster["excluded_exited"]), "muted"),
                      ui.kpi("Kapsam", kapsam if len(kapsam) < 18 else kapsam[:16] + "…")], className="g g-2 gap-10"),
        ui.kart([ui.lbl("Göreve göre dağılım · liste sırası"), *dagilim_satir,
                 html.Span("Sıralama: Formen › Ekip başı › diğer görevler (kişi sayısına göre) › Bayrakçı",
                           className="small-12 muted")], "pad col", style={"gap": "10px"}),
        ui.kart([
            html.Div([
                html.Label("Çıkış yapanlar", htmlFor="el-cikis-ay", className="vm-lbl", style={"whiteSpace": "nowrap"}),
                html.Div(dcc.Dropdown(id="el-cikis-ay", options=[{"label": ex.ay_adi(a), "value": a}
                                                                  for a in ex.son_aylar(12)],
                                      value=ex.son_aylar(1)[0], clearable=False, searchable=False,
                                      persistence=True, persistence_type="memory"),
                         className="vm-dd grow"),
            ], className="row", style={"gap": "8px"}),
            html.Span(id="el-cikis-bilgi", className="small-12 muted"),
            html.Button([ui.ikon("file-earmark-excel"), "Excel listesini oluştur"], id="el-olustur", className="vm-btn pri",
                        n_clicks=0),
            html.Div([html.Button("Excel'de aç", id="el-ac-dosya", className="vm-btn", n_clicks=0, disabled=not son_dosya),
                      html.Button("Klasörü aç", id="el-ac-klasor", className="vm-btn", n_clicks=0)], className="g g-2 gap-8"),
            html.Span(f"Son: {son_dosya}" if son_dosya else "Henüz oluşturulmadı.", className="vm-sub",
                      style={"fontSize": "11px", "wordBreak": "break-all"}),
            html.Div(id="el-sonuc"),
        ], "pad col", style={"gap": "8px"}),
        _form_karti(sirali),
    ], className="col", style={"gap": "14px"})

    return html.Div([kagit, sag], className="g", style={"gridTemplateColumns": "minmax(0,1fr) 360px"})


@callback(
    Output("el-cikis-bilgi", "children"),
    Input("el-cikis-ay", "value"),
    State("el-grup", "value"), State("el-ekip", "value"), State("el-gorev", "value"),
)
def cikis_bilgi(ay, grup, ekip, gorev):
    if not ay:
        return ""
    n = len(ex.cikis_yapanlar(ay, grup, ekip, gorev))
    return (f"{ex.ay_adi(ay)} ayında bu kapsamda {n} kişi çıkış yaptı — Excel'de görev özetinin altında listelenir."
            if n else f"{ex.ay_adi(ay)} ayında bu kapsamda çıkış yapan yok.")


def _gorunenler(kayitlar, gizli):
    gizli = set(gizli or [])
    return [r for r in kayitlar if r["tc"] not in gizli]


# NOT: Göz düğmeleri ile "Hepsini geri al" düğmesi AYRI callback'lerde. Dash bir callback'i
# ancak tüm (joker olmayan) girdileri sayfada varken çalıştırır; "Hepsini geri al" sadece
# gizlenen kişi varken çizildiği için aynı callback'te durunca göz düğmeleri hiç çalışmıyordu.
@callback(
    Output("el-gizli", "data"),
    Input({"type": "el-goz", "tc": ALL}, "n_clicks"),
    State("el-gizli", "data"),
    prevent_initial_call=True,
)
def gizle_goster(_goz, gizli):
    # Önizleme yeniden çizilince düğmeler n_clicks=0 ile yeniden oluşur; bu tetiklemeleri yok say.
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update
    tc = ctx.triggered_id["tc"]
    gizli = list(gizli or [])
    if tc in gizli:
        gizli.remove(tc)
    else:
        gizli.append(tc)
    return gizli


@callback(
    Output("el-gizli", "data", allow_duplicate=True),
    Input("el-gizli-sifirla", "n_clicks"),
    prevent_initial_call=True,
)
def gizlileri_sifirla(n):
    return [] if n else no_update


def _form_karti(kayitlar):
    ozet = mf.form_ozeti(kayitlar)
    return ui.kart([
        html.Div([html.Span(ui.ikon("printer"), className="vm-pipe-ic pi-navy"),
                  html.Div([html.Span("Günlük mesai formu", style={"fontWeight": 700, "fontSize": "15px"}),
                            html.Span("A4 yatay · sahada elle doldurulur", className="small-12 muted")],
                           className="col", style={"gap": 0})], className="row", style={"gap": "10px"}),
        html.Div([
            ui.bilgi_satiri("Alt ekip (ayrı form)", ui.sayi(ozet["ekip"]), True),
            ui.bilgi_satiri("Toplam sayfa", ui.sayi(ozet["sayfa"]), True),
            ui.bilgi_satiri("Sayfa başına", f"en fazla {mf.SAYFA_MAX_KISI} kişi", False),
        ], className="col", style={"gap": "4px"}),
        html.Div([
            html.Label("Boş satır", htmlFor="el-form-bos", className="vm-lbl", style={"whiteSpace": "nowrap"}),
            html.Div(dcc.Dropdown(id="el-form-bos", options=[{"label": str(i), "value": i} for i in range(0, 11)],
                                  value=mf.VARSAYILAN_BOS_SATIR, clearable=False, searchable=False,
                                  style={"width": "80px"}), className="vm-dd"),
            html.Span(className="spacer"),
            dbc.Checkbox(id="el-form-tc", label="TC tam görünsün", value=True),
        ], className="row", style={"gap": "8px"}),
        html.Button([ui.ikon("file-earmark-pdf"), "Mesai formu oluştur (PDF)"], id="el-form", className="vm-btn pri",
                    n_clicks=0),
        html.Button([ui.ikon("box-arrow-up-right"), "Son PDF'i aç"], id="el-form-ac", className="vm-btn", n_clicks=0,
                    disabled=True),
        html.Div(id="el-form-sonuc"),
    ], "pad col", style={"gap": "10px"})


@callback(
    Output("el-form-sonuc", "children"),
    Output("el-form-dosya", "data"),
    Output("el-form-ac", "disabled"),
    Input("el-form", "n_clicks"),
    State("el-grup", "value"), State("el-ekip", "value"), State("el-gorev", "value"),
    State("el-form-bos", "value"), State("el-form-tc", "value"),
    State("el-gizli", "data"),
    prevent_initial_call=True,
)
def form_olustur(n, grup, ekip, gorev, bos, tc, gizli):
    if not n:
        return no_update, no_update, no_update
    roster = el.build_roster(grup, ekip, gorev)
    roster["records"] = _gorunenler(roster["records"], gizli)
    if not roster["records"]:
        return ui.banner("Listede kimse kalmadı: seçilen filtrelerle eşleşen personel yok ya da hepsi gizlendi.", "warn"), no_update, no_update
    yol = report.OUTPUT_DIR / mf.dosya_adi(el.kapsam_adi(roster))
    try:
        sonuc = mf.pdf_olustur(roster["records"], yol, bos_satir=int(bos if bos is not None else 3), tc_goster=bool(tc))
    except PermissionError:
        return ui.banner("PDF yazılamadı: aynı adlı dosya şu an açık olabilir. Kapatıp tekrar deneyin.", "err"), \
            no_update, no_update
    except mf.MesaiFormuHatasi as e:
        return ui.banner(str(e), "err"), no_update, no_update
    except Exception as e:
        return ui.banner(f"Mesai formu oluşturulamadı: {e}", "err"), no_update, no_update
    el.open_path(sonuc["yol"])
    return (ui.banner([html.Strong(f"{sonuc['sayfa']} sayfa, {sonuc['ekip']} alt ekip. "), html.Code(str(sonuc["yol"]))],
                      "ok", "check2-circle"), str(sonuc["yol"]), False)


@callback(
    Output("el-dummy", "data", allow_duplicate=True),
    Input("el-form-ac", "n_clicks"),
    State("el-form-dosya", "data"),
    prevent_initial_call=True,
)
def form_ac(n, yol):
    if n and yol:
        el.open_path(yol)
    return no_update


@callback(Output("el-tumu", "data"), Input("el-tumu-btn", "n_clicks"), prevent_initial_call=True)
def tumunu_goster(n):
    return True if n else no_update


@callback(
    Output("el-tumu", "data", allow_duplicate=True),
    Input("el-grup", "value"), Input("el-ekip", "value"), Input("el-gorev", "value"),
    prevent_initial_call=True,
)
def filtre_degisti(_grup, _ekip, _gorev):
    return False     # filtre değişince yine ilk 60 satır


@callback(
    Output("el-son-dosya", "data"),
    Output("el-dummy", "data", allow_duplicate=True),
    Input("el-olustur", "n_clicks"),
    State("el-grup", "value"), State("el-ekip", "value"), State("el-gorev", "value"),
    State("el-gizli", "data"),
    State("el-cikis-ay", "value"),
    prevent_initial_call=True,
)
def olustur(n, grup, ekip, gorev, gizli, ay=None):
    if not n:
        return no_update, no_update
    try:
        # el.excel_olustur ile aynı; tek fark gizlenen personelin çıkarılması
        roster = el.build_roster(grup, ekip, gorev)
        roster["records"] = _gorunenler(roster["records"], gizli)
        kapsam = el.kapsam_adi(roster)
        ay = ay or ex.son_aylar(1)[0]
        cikanlar = ex.cikis_yapanlar(ay, grup, ekip, gorev)
        yol = ex.yaz(roster, kapsam, cikanlar, ay,
                     report.OUTPUT_DIR / f"{report._slugify(kapsam)}_ekip_listesi.xlsx")
    except PermissionError:
        return no_update, "Dosya yazılamadı: aynı adlı Excel açık olabilir."
    except Exception as e:
        return no_update, f"Liste oluşturulamadı: {e}"
    return str(yol), None


@callback(
    Output("el-dummy", "data"),
    Input("el-ac-dosya", "n_clicks"),
    Input("el-ac-klasor", "n_clicks"),
    State("el-son-dosya", "data"),
    prevent_initial_call=True,
)
def ac(_d, _k, son):
    if not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update
    if ctx.triggered_id == "el-ac-dosya" and son:
        el.open_path(son)
    elif ctx.triggered_id == "el-ac-klasor":
        el.open_path(el.cikti_klasoru())
    return no_update


@callback(Output("el-sonuc", "children"), Input("el-dummy", "data"), prevent_initial_call=True)
def hata_goster(mesaj):
    return ui.banner(mesaj, "err") if mesaj else None
