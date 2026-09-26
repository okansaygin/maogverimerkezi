# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: SAP Puantaj Aktarım.

Aylık kaynak -> günlük hedef roster(lar). Motor: veri_merkezi_puantaj_aktarim
(puantaj_engine.run_transfer'in ince sarmalayıcısı) - TEKRARLANMAZ.
Yeni: eşleşmeyen satırlar için isim benzerliği önerisi ve tek tıkla atama
(Suite'teki manuel eşleştirmenin karşılığı)."""

import datetime
from pathlib import Path

from dash import dcc, html, Input, Output, State, ALL, MATCH, no_update, callback, ctx
import dash_bootstrap_components as dbc

import puantaj_engine as engine
import veri_merkezi_puantaj_aktarim as pa
import veri_merkezi_ekip_listesi as el
import veri_merkezi_ui as ui
from veri_merkezi_sayfa_ice_aktar import gecici_kaydet

# Tek kullanıcılı masaüstü uygulaması - durum sunucu belleğinde.
_SAP = {"kaynak": None, "tarihler": [], "hedefler": [], "sonuclar": {}, "eslestirme": {}}


def _gun_izgarasi(tarihler, secili):
    if not tarihler:
        return html.Span("Önce aylık kaynak dosyasını yükleyin.", className="small-12 muted")
    hucreler = [html.Span(g, className="vm-cal-h") for g in ui.GUN_KISA]
    ilk = tarihler[0]
    hucreler += [html.Button(className="vm-daybtn ph", disabled=True) for _ in range(ilk.weekday())]
    mevcut = set(tarihler)
    d = ilk
    while d <= tarihler[-1]:
        iso = d.isoformat()
        hucreler.append(html.Button(str(d.day), id={"type": "sap-gun-btn", "v": iso}, n_clicks=0,
                                    disabled=d not in mevcut, title=d.strftime("%d.%m.%Y"),
                                    className="vm-daybtn on" if iso == secili else "vm-daybtn"))
        d += datetime.timedelta(days=1)
    return html.Div(hucreler, className="vm-cal", style={"gap": "4px"})


def _hedef_listesi():
    if not _SAP["hedefler"]:
        return html.Span("Henüz hedef dosya yok. Birden fazla dosya seçebilirsiniz.", className="small-12 muted")
    kartlar = []
    for p in _SAP["hedefler"]:
        s = _SAP["sonuclar"].get(p.name)
        cipler = []
        if s and not isinstance(s, str):
            atanan = len(_SAP["eslestirme"].get(p.name, {}))
            kalan = len(s["unmatched"]) - atanan
            cipler = [ui.chip(f"{ui.sayi(len(s['matched']))} eşleşti", "green"),
                      ui.chip(f"{kalan} eşleşmedi", "red" if kalan else "gray")]
        elif isinstance(s, str):
            cipler = [ui.chip("hata", "red")]
        kartlar.append(html.Div([
            html.Span(ui.ikon("file-earmark-spreadsheet"), className="vm-pipe-ic",
                      style={"width": "34px", "height": "34px", "borderRadius": "8px", "background": "#F4F2EE"}),
            html.Span(p.name, className="mono grow", style={"fontSize": "13px", "fontWeight": 600, "wordBreak": "break-all"}),
            *cipler,
        ], className="row", style={"padding": "10px 12px", "border": "1px solid #E3DED5", "borderRadius": "8px", "gap": "10px"}))
    return html.Div(kartlar, className="col", style={"gap": "8px"})


def _ayar_karti():
    cfg = pa.get_config()
    return ui.kart([
        ui.lbl("3 · Aktarım ayarları"),
        html.Div([
            ui.bilgi_satiri("Kaynak sayfa", cfg["source_sheet"], True),
            ui.bilgi_satiri("TC sütunu (kaynak → hedef)", f"{cfg['source_tc_col']} → {cfg['target_tc_col']}", True),
            ui.bilgi_satiri("Saat · not · durum", f"{cfg['target_write_col']} · {cfg['target_note_col']} · {cfg['target_status_col']}", True),
            ui.bilgi_satiri("Veri satırları (kaynak)", f"{cfg['source_data_start_row']}–{cfg['source_data_end_row']}", True),
            ui.bilgi_satiri("Anormal değer eşiği", f"{cfg['max_reasonable_hours']} sa", True),
        ], className="col", style={"gap": "7px"}),
        dbc.Checkbox(id="sap-ozet", label="Çıktıya Özet sayfası ekle", value=bool(cfg.get("write_summary_sheet", True))),
        html.Div("Puantaj Suite ile ortak ayar dosyası (puantaj_settings.json) — her aktarımdan önce yeniden okunur. "
                 "Sütun/satır ayarlarını Suite'teki Gelişmiş Ayarlar'dan değiştirebilirsiniz.", className="vm-note",
                 style={"marginTop": "auto"}),
    ], "pad col", style={"gap": "10px"})


def sayfa():
    secili = _SAP.get("gun")
    return ui.sayfa("SAP Puantaj Aktarım", "Veri Girişi / SAP Puantaj Aktarım", html.Div([
        html.Div([
            html.Div([
                ui.kart([
                    ui.lbl("1 · Aylık kaynak"),
                    ui.yukleme_alani("sap-kaynak-upload", "Aylık puantaj Excel'i seç", kucuk=True),
                    html.Div(_kaynak_bilgi(), id="sap-kaynak-info"),
                    ui.lbl("Aktarılacak gün"),
                    html.Div(_gun_izgarasi(_SAP["tarihler"], secili), id="sap-gunler"),
                ], "pad col", style={"gap": "10px"}),
                ui.kart([
                    ui.lbl("2 · Günlük hedef roster(lar)"),
                    ui.yukleme_alani("sap-hedef-upload", "Hedef dosya(lar) ekle", kucuk=True, coklu=True),
                    html.Div(_hedef_listesi(), id="sap-hedef-liste"),
                    html.Span("Hedef dosyalar değiştirilmez; sonuçlar OUTPUT klasörüne aynı adla kaydedilir.",
                              className="small-12 muted"),
                ], "pad col", style={"gap": "10px"}),
                _ayar_karti(),
            ], className="g", style={"gridTemplateColumns": "330px minmax(0,1fr) 330px"}),
            dcc.Loading(html.Div(id="sap-onizleme"), type="default", color="#1F4E78"),
            html.Div(id="sap-sonuc"),
        ], className="vm-page"),
        html.Footer([
            html.Span(["Sonraki adım: oluşan günlük roster'ı ",
                       dcc.Link("Çalışma Bilgileri olarak içe aktar →", href="/ice-aktar", style={"fontWeight": 600})],
                      className="grow", style={"fontSize": "13px", "color": "#5A6572"}),
            html.Button([ui.ikon("search"), "Önizle"], id="sap-onizle-btn", className="vm-btn", n_clicks=0),
            html.Button([ui.ikon("arrow-left-right"), "Aktar ve kaydet"], id="sap-aktar-btn", className="vm-btn pri", n_clicks=0),
        ], className="vm-footbar"),
        dcc.Store(id="sap-gun", data=secili),
        dcc.Store(id="sap-dummy"),
    ], style={"display": "flex", "flexDirection": "column", "flexGrow": 1}), govde_cls=None)


def _kaynak_bilgi():
    k = _SAP.get("kaynak")
    if not k:
        return None
    return html.Div([
        html.Span(ui.ikon("file-earmark-spreadsheet"), className="vm-pipe-ic pi-navy",
                  style={"width": "34px", "height": "34px", "borderRadius": "8px"}),
        html.Div([html.Span(k.name, className="mono", style={"fontSize": "13px", "fontWeight": 600, "wordBreak": "break-all"}),
                  html.Span(f"{len(_SAP['tarihler'])} gün sütunu bulundu", className="small-12 muted")],
                 className="col", style={"gap": 0}),
    ], className="row", style={"gap": "10px"})


# ---------------------------------------------------------------- önizleme görünümü

def _hedef_paneli(idx, ad, s):
    if isinstance(s, str):
        return ui.banner(f"{ad}: {s}", "err", "x-octagon")
    atanan = _SAP["eslestirme"].get(ad, {})
    izin_say = {}
    for _r, _tc, _n, lab in s.get("leave_marked", []):
        izin_say[lab] = izin_say.get(lab, 0) + 1

    satirlar = []
    for r, tc, isim in s["unmatched"]:
        oneriler = s["name_suggestions"].get(r) or []
        if oneriler:
            stc, sad, skor = oneriler[0]
            diger = f" (+{len(oneriler) - 1} öneri)" if len(oneriler) > 1 else ""
            secili = atanan.get(r) == stc
            satirlar.append(html.Tr([
                html.Td(html.Div([html.Span(isim or "(isimsiz)", style={"fontWeight": 600}),
                                  html.Span(f"satır {r}", className="vm-sub")], className="col", style={"gap": 0})),
                html.Td(ui.maske_tc(tc), className="mono", style={"fontSize": "13px"}),
                html.Td(html.Div([html.Span(sad + diger), html.Span(ui.maske_tc(stc), className="vm-sub")],
                                 className="col", style={"gap": 0})),
                html.Td(html.Div([ui.bar([(skor * 100, "bg-navy")], "thin"),
                                  html.Span(f"%{int(skor * 100)}", className="mono small-12")],
                                 className="row", style={"gap": "8px", "width": "120px"})),
                html.Td(html.Button("Atandı" if secili else "Ata",
                                    id={"type": "sap-ata", "h": idx, "row": r, "tc": stc}, n_clicks=0,
                                    className="vm-btn xs ok" if secili else "vm-btn xs")),
            ]))
        else:
            satirlar.append(html.Tr([
                html.Td(html.Div([html.Span(isim or "(isimsiz)", style={"fontWeight": 600}),
                                  html.Span(f"satır {r}", className="vm-sub")], className="col", style={"gap": 0})),
                html.Td(ui.maske_tc(tc), className="mono"), html.Td("Benzer isim bulunamadı", className="muted", colSpan=3),
            ]))

    kaynakta = s["source_not_in_target"]
    sag = [ui.lbl("İzinli olarak işaretlenecek (sarı)"),
           html.Div([ui.chip(f"{lab} · {n}", "purple") for lab, n in sorted(izin_say.items(), key=lambda kv: -kv[1])]
                    or [html.Span("Yok", className="small-12 muted")], className="row", style={"flexWrap": "wrap", "gap": "6px"}),
           ui.lbl("Kaynakta var, hedefte yok", style={"marginTop": "6px"})]
    for tc, isim, val, izin in kaynakta[:6]:
        deger = izin[0][1] if (izin and not val) else f"{ui.saat(val)} sa"
        sag.append(html.Div([html.Span(isim, className="grow"), html.Span(str(deger), className="mono")],
                            className="row", style={"fontSize": "13px"}))
    if len(kaynakta) > 6:
        sag.append(html.Span(f"+{len(kaynakta) - 6} kişi daha", className="small-12 muted"))
    if not kaynakta:
        sag.append(html.Span("Yok", className="small-12 muted"))
    if s["abnormal"]:
        sag.append(ui.lbl("Anormal değer", style={"marginTop": "6px"}))
        for r, tc, isim, v in s["abnormal"][:5]:
            sag.append(html.Div([html.Span(isim or tc, className="grow"), html.Span(f"{v} sa", className="mono c-orange")],
                                className="row", style={"fontSize": "13px"}))

    uyarilar = list(s.get("warnings", []))
    if s.get("odd_row_counts"):
        uyarilar.append(f"{len(s['odd_row_counts'])} TC kaynakta beklenmeyen sayıda satırda geçiyor (1 ya da 2 bekleniyordu).")

    kalan = len(s["unmatched"]) - len(atanan)
    return html.Div([
        html.Div([ui.banner(u, "warn") for u in uyarilar], className="col", style={"gap": "8px"}) if uyarilar else None,
        html.Div([
            ui.kpi("Eşleşen", ui.sayi(len(s["matched"])), "green"),
            ui.kpi("Eşleşmeyen", ui.sayi(kalan), "red" if kalan else "green"),
            ui.kpi("İzinli / raporlu", ui.sayi(len(s.get("leave_marked", []))), "purple"),
            ui.kpi("Anormal değer", ui.sayi(len(s["abnormal"])), "orange" if s["abnormal"] else None),
            ui.kpi("TC sorunlu", ui.sayi(len(s["tc_issues"])), "orange" if s["tc_issues"] else None),
            ui.kpi("Kaynakta var, hedefte yok", ui.sayi(len(kaynakta))),
        ], className="g g-6 gap-10"),
        html.Div([
            html.Div([
                html.Div([ui.h2("Eşleşmeyen satırlar — isimle eşleştirme önerisi"),
                          html.Span("atanmayanlar kırmızıya boyanır", className="small-12 muted")], className="vm-card-head pad"),
                ui.tablo(["Hedef satır", "Hedefteki TC", "En iyi öneri (kaynak)", "Benzerlik", ""], satirlar)
                if satirlar else html.Div(ui.banner("Hedefteki tüm TC'ler kaynakta bulundu.", "ok", "check2-circle"),
                                          style={"padding": "0 16px 16px"}),
            ], className="col", style={"gap": 0, "minWidth": 0}),
            html.Div(sag, className="col", style={"gap": "8px", "padding": "14px 16px", "borderLeft": "1px solid #EFEBE4"}),
        ], className="vm-card clip g", style={"gridTemplateColumns": "minmax(0,1fr) 340px", "gap": 0}),
    ], className="col", style={"gap": "12px", "paddingTop": "14px"})


def _onizleme_gorunumu(tarih_iso):
    if not _SAP["hedefler"]:
        return None
    parcalar = [(f"t{i}", p.stem, _hedef_paneli(i, p.name, _SAP["sonuclar"].get(p.name, "önizlenmedi")))
                for i, p in enumerate(_SAP["hedefler"])]
    return html.Section([
        html.Div([ui.h2("Önizleme", f"{ui.tarih_tr(tarih_iso)} · hiçbir dosya değiştirilmedi")], className="vm-card-head"),
        ui.sekmeler("sap", parcalar, aktif="t0"),
    ], className="vm-card pad-l")


# ---------------------------------------------------------------- callback'ler

@callback(
    Output("sap-kaynak-info", "children"),
    Output("sap-gunler", "children"),
    Output("sap-gun", "data"),
    Input("sap-kaynak-upload", "contents"),
    State("sap-kaynak-upload", "filename"),
    prevent_initial_call=True,
)
def kaynak_yuklendi(contents, filename):
    if not contents:
        return no_update, no_update, no_update
    yol = gecici_kaydet(contents, filename)
    try:
        tarihler = pa.list_available_dates(pa.get_config(), yol)
    except Exception as e:
        return ui.banner(f"Kaynak okunamadı: {e}", "err"), _gun_izgarasi([], None), None
    if not tarihler:
        return ui.banner("Kaynakta hiçbir tarih sütunu bulunamadı. Suite'teki kaynak ayarlarını kontrol edin.", "warn"), \
            _gun_izgarasi([], None), None
    _SAP.update(kaynak=yol, tarihler=tarihler, sonuclar={}, eslestirme={})
    bugun = datetime.date.today()
    gecmis = [t for t in tarihler if t < bugun]
    secili = (gecmis[-1] if gecmis else tarihler[-1]).isoformat()
    _SAP["gun"] = secili
    return _kaynak_bilgi(), _gun_izgarasi(tarihler, secili), secili


@callback(
    Output("sap-gunler", "children", allow_duplicate=True),
    Output("sap-gun", "data", allow_duplicate=True),
    Input({"type": "sap-gun-btn", "v": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def gun_secildi(_n):
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update, no_update
    secili = ctx.triggered_id["v"]
    _SAP["gun"] = secili
    return _gun_izgarasi(_SAP["tarihler"], secili), secili


@callback(
    Output("sap-hedef-liste", "children"),
    Output("sap-onizleme", "children", allow_duplicate=True),
    Input("sap-hedef-upload", "contents"),
    State("sap-hedef-upload", "filename"),
    prevent_initial_call=True,
)
def hedef_yuklendi(icerikler, adlar):
    if not icerikler:
        return no_update, no_update
    yollar = [gecici_kaydet(c, n) for c, n in zip(icerikler, adlar)]
    _SAP.update(hedefler=yollar, sonuclar={}, eslestirme={})
    return _hedef_listesi(), None


@callback(
    Output("sap-onizleme", "children"),
    Output("sap-hedef-liste", "children", allow_duplicate=True),
    Output("sap-sonuc", "children", allow_duplicate=True),
    Input("sap-onizle-btn", "n_clicks"),
    State("sap-gun", "data"),
    prevent_initial_call=True,
)
def onizle(n, tarih):
    if not n:
        return no_update, no_update, no_update
    if not _SAP["kaynak"]:
        return ui.banner("Önce aylık kaynak Excel'ini yükleyin.", "warn"), no_update, None
    if not _SAP["hedefler"]:
        return ui.banner("Önce en az bir günlük hedef roster yükleyin.", "warn"), no_update, None
    if not tarih:
        return ui.banner("Aktarılacak günü seçin.", "warn"), no_update, None
    cfg = pa.get_config()
    _SAP["sonuclar"], _SAP["eslestirme"] = {}, {}
    for p in _SAP["hedefler"]:
        try:
            _SAP["sonuclar"][p.name] = pa.onizleme(cfg, _SAP["kaynak"], p, tarih)
        except engine.TransferError as e:
            _SAP["sonuclar"][p.name] = str(e)
        except Exception as e:
            _SAP["sonuclar"][p.name] = f"Dosya işlenemedi: {e}"
    return _onizleme_gorunumu(tarih), _hedef_listesi(), None


@callback(
    Output({"type": "sap-ata", "h": MATCH, "row": MATCH, "tc": MATCH}, "children"),
    Output({"type": "sap-ata", "h": MATCH, "row": MATCH, "tc": MATCH}, "className"),
    Input({"type": "sap-ata", "h": MATCH, "row": MATCH, "tc": MATCH}, "n_clicks"),
    State({"type": "sap-ata", "h": MATCH, "row": MATCH, "tc": MATCH}, "id"),
    prevent_initial_call=True,
)
def ata(n, id_):
    if not n:
        return no_update, no_update
    ad = _SAP["hedefler"][id_["h"]].name
    esl = _SAP["eslestirme"].setdefault(ad, {})
    if esl.get(id_["row"]) == id_["tc"]:
        esl.pop(id_["row"], None)
        return "Ata", "vm-btn xs"
    esl[id_["row"]] = id_["tc"]
    return "Atandı", "vm-btn xs ok"


@callback(
    Output("sap-sonuc", "children"),
    Input("sap-aktar-btn", "n_clicks"),
    State("sap-gun", "data"),
    State("sap-ozet", "value"),
    prevent_initial_call=True,
)
def aktar(n, tarih, ozet):
    if not n:
        return no_update
    if not _SAP["kaynak"] or not _SAP["hedefler"] or not tarih:
        return ui.banner("Önce kaynak, hedef(ler) ve gün seçilmeli.", "warn")
    cfg = pa.get_config()
    cfg["write_summary_sheet"] = bool(ozet)
    basarili, hatali = [], []
    for p in _SAP["hedefler"]:
        try:
            s = pa.aktar(cfg, _SAP["kaynak"], p, tarih, manual_mapping=_SAP["eslestirme"].get(p.name) or None)
            basarili.append((p.name, s))
        except Exception as e:
            hatali.append(f"{p.name}: {e}")
    parcalar = []
    if basarili:
        parcalar.append(ui.kart([
            html.Div([html.Span(ui.ikon("check-lg"), className="vm-pipe-ic pi-ok"),
                      ui.h2(f"{len(basarili)} dosya aktarıldı · {ui.tarih_tr(tarih)}")], className="row"),
            ui.tablo(["Dosya", ("Eşleşen", "r"), ("Manuel", "r"), ("Kırmızı", "r"), ("İzinli", "r")], [
                [html.Td(str(s["output_path"]), className="mono", style={"fontSize": "12px"}),
                 html.Td(ui.sayi(len(s["matched"])), className="mono r"),
                 html.Td(ui.sayi(len(s["manual_matched"])), className="mono r"),
                 html.Td(ui.sayi(len(s["flagged"])), className="mono r c-red"),
                 html.Td(ui.sayi(len(s.get("leave_marked", []))), className="mono r")] for _a, s in basarili]),
            html.Div([html.Button([ui.ikon("folder2-open"), "Çıktı klasörünü aç"], id="sap-ac-klasor", className="vm-btn",
                                  n_clicks=0),
                      dcc.Link([ui.ikon("upload"), "Çalışma Bilgileri olarak içe aktar"], href="/ice-aktar",
                               className="vm-btn pri")], className="row"),
        ], "pad-l col", style={"gap": "12px"}))
    for h in hatali:
        parcalar.append(ui.banner(h, "err", "x-octagon"))
    return html.Div(parcalar, className="col", style={"gap": "12px"})


@callback(Output("sap-dummy", "data"), Input("sap-ac-klasor", "n_clicks"), prevent_initial_call=True)
def klasor_ac(n):
    if n:
        pa.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        el.open_path(pa.OUTPUT_DIR)
    return no_update
