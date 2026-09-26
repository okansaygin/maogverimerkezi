# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: İçe Aktarma Geçmişi (import_batches denetim kaydı).

v5: seçilen içe aktarma, detay panelinden GERİ ALINABİLİR. Akış iki adımlıdır:
"Bu içe aktarmayı geri al" -> ne olacağının önizlemesi (silinecek / eski değerine dönecek /
korunacak kayıtlar, engelleyen daha yeni içe aktarmalar) -> not + "Geri almayı onayla".
Asıl iş veri_merkezi_geri_alma'dadır; bu sayfa yalnızca gösterir ve onay alır.
"""

import json

from dash import dcc, html, Input, Output, State, ALL, no_update, callback, ctx

import veri_merkezi_geri_alma as ga
import veri_merkezi_sema as sema
import veri_merkezi_sorgu as sorgu
import veri_merkezi_ui as ui

TUR = {"kimlik": ("Personel", "navy-l", "Personel Bilgileri (kimlik)"),
       "mesai": ("Mesai", "orange", "Çalışma Bilgileri (SAP mesai)"),
       None: ("—", "gray", "Bilinmiyor")}
DURUM = {"tamamlandi": ("tamamlandı", "green"), "hata": ("hata", "red"), "basladi": ("yarım kaldı", "orange"),
         "iptal": ("iptal", "gray"), "geri_alindi": ("geri alındı", "purple")}


def sayfa():
    st = sorgu.get_import_stats(30)
    return ui.sayfa("İçe Aktarma Geçmişi", "Denetim / İçe Aktarma Geçmişi", html.Div([
        html.Div([
            html.Section([
                ui.kpi("Son 30 gün", ui.sayi(st["toplam"]), None, "içe aktarma"),
                ui.kpi("Başarılı", ui.sayi(st["basarili"]), "green", "tek işlemde yazıldı"),
                ui.kpi("Hatalı / geri alınan", f"{ui.sayi(st['hata'])} / {ui.sayi(st['geri_alinan'])}",
                       "red" if st["hata"] else None, "hatalıda hiçbir şey yazılmadı"),
                ui.kpi("Yazılan günlük kayıt", ui.sayi(st["gun_kaydi"]), "navy", "son 30 gün"),
            ], className="g g-4 gap-12"),
            html.Div([
                ui.segment("gc-tur", [("", "Tümü"), ("kimlik", "Personel Bilgileri"), ("mesai", "Çalışma Bilgileri")], ""),
                html.Span(className="spacer"),
                html.Div(ui.dropdown("gc-durum", [{"label": "Tüm durumlar", "value": ""},
                                                  {"label": "Tamamlandı", "value": "tamamlandi"},
                                                  {"label": "Hata", "value": "hata"},
                                                  {"label": "Geri alındı", "value": "geri_alindi"},
                                                  {"label": "Yarım kaldı", "value": "basladi"}], "", clearable=False,
                                     searchable=False), style={"width": "180px"}),
            ], className="row"),
            html.Section(id="gc-tablo", className="vm-card clip"),
        ], className="vm-pane-c"),
        html.Aside(id="gc-detay", className="vm-pane-r"),
        dcc.Store(id="gc-secili"),
        dcc.Store(id="gc-yenile", data=0),
    ], className="vm-split", style={"gridTemplateColumns": "minmax(0,1fr) 380px"}), govde_cls=None)


@callback(
    Output("gc-tablo", "children"),
    Output("gc-secili", "data"),
    Input("gc-tur", "value"),
    Input("gc-durum", "value"),
    Input("gc-secili", "data"),
    Input("gc-yenile", "data"),
)
def tablo(tur, durum, secili, _yenile):
    kayitlar = sorgu.get_import_history_v2(limit=300, tur=tur or None)
    if durum:
        kayitlar = [k for k in kayitlar if k["durum"] == durum]
    if not kayitlar:
        return ui.bos("Bu filtrelerle içe aktarma kaydı yok."), None
    if secili not in [k["id"] for k in kayitlar]:
        secili = kayitlar[0]["id"]
    satirlar = []
    for k in kayitlar:
        t = TUR.get(k["tur_hesap"], TUR[None])
        d = DURUM.get(k["durum"], (k["durum"], "gray"))
        donem = "—"
        if k["donem_baslangic"]:
            donem = ui.tarih_tr(k["donem_baslangic"])
            if k["donem_bitis"] and k["donem_bitis"] != k["donem_baslangic"]:
                donem += f" – {ui.tarih_tr(k['donem_bitis'])}"
        pers = f"{ui.sayi(k['yeni_personel'])} / {ui.sayi(k['guncellenen_personel'])}" if k["tur_hesap"] == "kimlik" else "—"
        gun = ui.sayi(k["yazilan_gun_kaydi"]) if k["tur_hesap"] == "mesai" else "—"
        cls = "click " + ("sel" if k["id"] == secili else ("err" if k["durum"] == "hata" else ""))
        satirlar.append(html.Tr([
            html.Td(str(k["id"]), className="mono muted"),
            html.Td(ui.chip(t[0], t[1])),
            html.Td(k["dosya_adi"], className="mono", style={"fontSize": "13px", "fontWeight": 500}),
            html.Td(ui.chip(d[0], d[1])),
            html.Td(ui.tarih_tr(k["baslama_zamani"], saatli=True), className="mono small-12"),
            html.Td(donem, className="mono small-12 muted"),
            html.Td(pers, className="mono r"),
            html.Td(gun, className="mono r"),
            html.Td(ui.sayi(k["uyari_sayisi"]), className="mono r", style={"fontWeight": 600}),
        ], id={"type": "gc-satir", "id": k["id"]}, n_clicks=0, className=cls))
    return ui.tablo(["#", "Tür", "Dosya", "Durum", "Zaman", "Dönem", ("Yeni / güncel.", "r"), ("Günlük kayıt", "r"),
                     ("Uyarı", "r")], satirlar), secili


@callback(
    Output("gc-secili", "data", allow_duplicate=True),
    Input({"type": "gc-satir", "id": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def satir_tiklandi(_n):
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update
    return ctx.triggered_id["id"]


@callback(Output("gc-detay", "children"), Input("gc-secili", "data"))
def detay(bid):
    return _detay_icerik(bid)


def _detay_icerik(bid, ust=None):
    """Detay paneli. ust: panelin en üstüne konacak bileşen (ör. geri alma sonucu)."""
    if not bid:
        return html.Span("Ayrıntı için listeden bir içe aktarma seçin.", className="muted")
    kayitlar = [k for k in sorgu.get_import_history_v2(limit=1000) if k["id"] == bid]
    if not kayitlar:
        return html.Span("Kayıt bulunamadı.", className="muted")
    k = kayitlar[0]
    t = TUR.get(k["tur_hesap"], TUR[None])
    d = DURUM.get(k["durum"], (k["durum"], "gray"))
    sure = f"{ui.sayi(k['sure_sn'], 1)} sn" if k["sure_sn"] is not None else "—"
    dagilim = sorgu.get_batch_alert_breakdown(bid)
    mx = max([a["adet"] for a in dagilim] + [1])
    renk = {"kritik": "bg-red", "uyari": "bg-orange", "bilgi": "bg-navy"}
    parca = [
        ust,
        html.Div([ui.chip(d[0], d[1]), html.Span(f"içe aktarma #{bid}", className="vm-sub")], className="row", style={"gap": "8px"}),
        html.H3(k["dosya_adi"], className="vm-h3", style={"wordBreak": "break-all"}),
        html.Div([
            ui.bilgi_satiri("Tür", t[2]),
            ui.bilgi_satiri("Başlama · süre", f"{ui.tarih_tr(k['baslama_zamani'], saatli=True)} · {sure}", True),
            ui.bilgi_satiri("Taranan satır", ui.sayi(k["taranan_satir"]), True),
            ui.bilgi_satiri("Yeni / güncellenen personel", f"{ui.sayi(k['yeni_personel'])} / {ui.sayi(k['guncellenen_personel'])}", True)
            if k["tur_hesap"] == "kimlik" else ui.bilgi_satiri("Yazılan günlük kayıt", ui.sayi(k["yazilan_gun_kaydi"]), True),
            ui.bilgi_satiri("SHA-256", f"{k['dosya_hash'][:4]}…{k['dosya_hash'][-4:]}" if k["dosya_hash"] else "—", True),
        ], className="col", style={"gap": "6px"}),
    ]
    if k["durum"] == "hata":
        parca.append(html.Div([html.Strong("Hata mesajı"), html.Span(k["hata_mesaji"] or "—", className="mono small-12"),
                               html.Span("İşlem geri alındı — veri setinde hiçbir değişiklik yok. Dosyayı tekrar "
                                         "yükleyebilirsiniz.")], className="col",
                              style={"gap": "4px", "padding": "12px", "borderRadius": "8px", "background": "#F9E1DE",
                                     "color": "#8F1B12", "fontSize": "13px"}))
    if k["durum"] == "geri_alindi":
        try:
            ozet = json.loads(k.get("geri_alma_ozeti") or "{}")
        except ValueError:
            ozet = {}
        parca.append(ui.banner([
            html.Strong(f"Geri alındı · {ui.tarih_tr(k.get('geri_alma_zamani'), saatli=True)}"), html.Br(),
            ga.ozet_metni(k["tur_hesap"], ozet) or "—",
            html.Br() if k.get("geri_alma_notu") else None,
            html.Span(f"Not: {k['geri_alma_notu']}", className="muted") if k.get("geri_alma_notu") else None,
        ], "info", "arrow-counterclockwise"))
    parca.append(ui.lbl("Üretilen uyarılar"))
    if dagilim:
        for a in dagilim:
            parca.append(html.Div([html.Span(a["etiket"], className="grow"),
                                   html.Div(ui.bar([(a["adet"] / mx * 100, renk[a["seviye"]])], "thin"), style={"width": "90px"}),
                                   html.Span(ui.sayi(a["adet"]), className="mono", style={"width": "36px", "textAlign": "right"})],
                                  className="row", style={"fontSize": "13px", "gap": "10px"}))
    else:
        parca.append(html.Span("Uyarı üretilmedi.", className="small-12 muted"))
    geri_alma = []
    if k["durum"] == "tamamlandi":
        if k.get("izli"):
            geri_alma = [html.Button([ui.ikon("arrow-counterclockwise"), "Bu içe aktarmayı geri al"], id="gc-geri-al-btn",
                                     n_clicks=0, className="vm-btn danger"),
                         html.Div(id="gc-geri-al-alan")]
        else:
            geri_alma = [html.Span("Bu içe aktarma v5'ten önce yapıldığı için geri alınamaz (eski değerler kayıtlı değil).",
                                   className="small-12 muted")]
    parca.append(html.Div([
        *geri_alma,
        dcc.Link(["Veri Kalitesi'nde incele", ui.ikon("chevron-right")], href="/kalite", className="vm-btn between"),
        html.Button(["Etkilenen personeli listele", ui.ikon("chevron-down")], id="gc-etkilenen-btn", n_clicks=0,
                    className="vm-btn between"),
        html.Div(id="gc-etkilenen"),
    ], className="col", style={"gap": "8px", "marginTop": "auto"}))
    return parca


@callback(
    Output("gc-etkilenen", "children"),
    Input("gc-etkilenen-btn", "n_clicks"),
    State("gc-secili", "data"),
    prevent_initial_call=True,
)
def etkilenen(n, bid):
    if not n or not bid:
        return no_update
    tcs = sorted(sorgu.get_batch_affected_tcs(bid))
    if not tcs:
        return html.Span("Bu içe aktarmada kaydı değişen personel yok.", className="small-12 muted")
    conn = sema.get_connection()
    try:
        isimler = {r["tc"]: r["ad_soyad"] for r in conn.execute(
            f"SELECT tc, ad_soyad FROM personnel WHERE tc IN ({','.join('?' * len(tcs[:200]))})", tcs[:200])}
    finally:
        conn.close()
    ogeler = [dcc.Link(isimler.get(tc) or tc, href=f"/personel/{tc}", style={"fontSize": "13px"}) for tc in tcs[:40]]
    return html.Div([html.Span(f"{ui.sayi(len(tcs))} kişi", className="vm-lbl"), *ogeler,
                     html.Span(f"+{len(tcs) - 40} kişi daha", className="small-12 muted") if len(tcs) > 40 else None],
                    className="col", style={"gap": "4px", "maxHeight": "260px", "overflowY": "auto"})


# ---------------------------------------------------------------- v5: geri alma

def _ornek_tablosu(baslik, satirlar, toplam, tur):
    if not satirlar:
        return None
    if tur == "mesai":
        def deger(n, f, d):
            if n is None and f is None:
                return "—"
            return f"{ui.saat((n or 0) + (f or 0))} sa" + (f" · {d}" if d else "")
        govde = [[html.Td(r.get("ad_soyad") or r["tc"], style={"fontSize": "12px"}),
                  html.Td(ui.tarih_tr(r["tarih"]), className="mono small-12"),
                  html.Td(deger(r.get("c_normal"), r.get("c_fazla"), r.get("c_durum")), className="small-12"),
                  html.Td(deger(r.get("eski_normal"), r.get("eski_fazla"), r.get("eski_durum")) if r["islem"] == "guncelle"
                          else "silinecek", className="small-12")] for r in satirlar]
        tablo = ui.tablo(["Kişi", "Gün", "Şimdi", "Sonra"], govde, "compact")
    else:
        govde = [[html.Td(r.get("ad_soyad") or r["tc"], style={"fontSize": "12px"}),
                  html.Td((r.get("alan") or "kayıt").replace("_", " "), className="small-12"),
                  html.Td(r.get("simdiki") if r.get("alan") else "", className="small-12"),
                  html.Td(r.get("eski_deger") if r.get("alan") else (r.get("neden") or "silinecek"),
                          className="small-12")] for r in satirlar]
        tablo = ui.tablo(["Kişi", "Alan", "Şimdi", "Sonra"], govde, "compact")
    fazla = toplam - len(satirlar)
    return html.Details([html.Summary(f"{baslik} · {ui.sayi(toplam)}", className="small-12", style={"cursor": "pointer"}),
                         html.Div(tablo, style={"overflowX": "auto", "marginTop": "6px"}),
                         html.Span(f"… ve {ui.sayi(fazla)} kayıt daha", className="small-12 muted") if fazla > 0 else None])


def _plan_gorunumu(on):
    """Geri alma önizlemesi + onay alanı."""
    if not on["izin"]:
        icerik = [ui.banner(on["neden"] or "Bu içe aktarma geri alınamaz.", "err", "x-octagon")]
        if on.get("engelleyenler"):
            icerik.append(html.Div([ui.lbl("Önce bunları geri alın (en yeniden başlayarak)")] + [
                html.Button(f"#{e['id']} · {e['dosya_adi']} · {ui.tarih_tr(e['baslama_zamani'], saatli=True)}",
                            id={"type": "gc-git", "id": e["id"]}, n_clicks=0, className="vm-btn sm between",
                            title="Bu içe aktarmayı seç")
                for e in on["engelleyenler"]], className="col", style={"gap": "6px"}))
        icerik.append(html.Button("Kapat", id="gc-geri-al-vazgec", n_clicks=0, className="vm-btn sm"))
        return html.Div(icerik, className="col vm-card pad", style={"gap": "10px"})

    s, o, tur = on["sayilar"], on["ornekler"], on["tur"]
    if tur == "mesai":
        maddeler = [(s["silinecek"], "günlük kayıt silinecek (bu içe aktarmayla eklenmişti)"),
                    (s["geri_yuklenecek"], "günlük kayıt önceki değerine dönecek")]
        ornekler = [_ornek_tablosu("Silinecek günler", o["silinecek"], s["silinecek"], tur),
                    _ornek_tablosu("Önceki değerine dönecek günler", o["geri_yuklenecek"], s["geri_yuklenecek"], tur)]
    else:
        maddeler = [(s["silinecek_kisi"], "kişi silinecek (bu içe aktarmayla eklenmişti)"),
                    (s["alan_geri_yuklenecek"], "alan önceki değerine dönecek"),
                    (s["korunan_kisi"], "kişi KORUNACAK (puantaj kaydı var ya da elle düzenlenmiş)"),
                    (s["alan_elle_degismis"], "alan KORUNACAK (sonradan elle değiştirilmiş)")]
        ornekler = [_ornek_tablosu("Silinecek kişiler", o["silinecek_kisi"], s["silinecek_kisi"], tur),
                    _ornek_tablosu("Önceki değerine dönecek alanlar", o["alan_geri_yuklenecek"], s["alan_geri_yuklenecek"], tur),
                    _ornek_tablosu("Korunacak kişiler", o["korunan_kisi"], s["korunan_kisi"], tur),
                    _ornek_tablosu("Korunacak (elle değişmiş) alanlar", o["alan_elle_degismis"], s["alan_elle_degismis"], tur)]
    if on.get("acik_uyari"):
        maddeler.append((on["acik_uyari"], "açık uyarı 'çözüldü' olarak kapatılacak"))
    degisiklik = sum(n for n, _m in maddeler[:2])
    return html.Div([
        html.Strong("Geri alınırsa:", style={"fontSize": "13px"}),
        html.Ul([html.Li([html.Strong(ui.sayi(n), className="mono"), f" {m}"]) for n, m in maddeler if n],
                style={"margin": 0, "paddingLeft": "18px", "fontSize": "13px"})
        if any(n for n, _m in maddeler) else html.Span("Veri setinde değişecek kayıt yok; yalnızca durum güncellenecek.",
                                                       className="small-12 muted"),
        *[x for x in ornekler if x is not None],
        html.Span("Önce veritabanının yedeği alınır. İşlem tek adımda yapılır: yarıda kalırsa hiçbir şey değişmez. "
                  "İçe aktarma kaydı silinmez, 'geri alındı' olarak işaretlenir; aynı dosya sonra yeniden yüklenebilir.",
                  className="small-12 muted"),
        ui.alan("Not (isteğe bağlı)", dcc.Textarea(id="gc-geri-al-not", className="vm-input",
                                                   placeholder="ör. yanlış ayın dosyası yüklenmişti",
                                                   style={"height": "60px"})),
        html.Div([html.Button([ui.ikon("arrow-counterclockwise"),
                               "Geri almayı onayla" if degisiklik else "Geri alındı olarak işaretle"],
                              id="gc-geri-al-onay", n_clicks=0, className="vm-btn sig sm"),
                  html.Button("Vazgeç", id="gc-geri-al-vazgec", n_clicks=0, className="vm-btn sm")],
                 className="row", style={"gap": "8px"}),
    ], className="col vm-card pad", style={"gap": "10px", "borderColor": "#F0C9A6"})


@callback(
    Output("gc-geri-al-alan", "children"),
    Input("gc-geri-al-btn", "n_clicks"),
    State("gc-secili", "data"),
    prevent_initial_call=True,
)
def geri_al_onizle(n, bid):
    if not n or not bid:
        return no_update
    try:
        return _plan_gorunumu(ga.onizle(bid))
    except Exception as e:
        import logging
        logging.getLogger("veri_merkezi").exception("geri alma önizlemesi #%s", bid)
        return ui.banner(f"Önizleme hazırlanamadı: {e}", "err")


@callback(
    Output("gc-geri-al-alan", "children", allow_duplicate=True),
    Input("gc-geri-al-vazgec", "n_clicks"),
    prevent_initial_call=True,
)
def geri_al_vazgec(n):
    return None if n else no_update


@callback(
    Output("gc-detay", "children", allow_duplicate=True),
    Output("gc-yenile", "data"),
    Input("gc-geri-al-onay", "n_clicks"),
    State("gc-geri-al-not", "value"),
    State("gc-secili", "data"),
    State("gc-yenile", "data"),
    prevent_initial_call=True,
)
def geri_al_onayla(n, not_, bid, yenile):
    if not n or not bid:
        return no_update, no_update
    try:
        sonuc = ga.geri_al(bid, not_ or "")
    except Exception as e:
        return _detay_icerik(bid, ui.banner([html.Strong("Geri alınamadı, hiçbir şey değişmedi. "), str(e)], "err",
                                            "x-octagon")), no_update
    ust = ui.banner([html.Strong(f"İçe aktarma #{bid} geri alındı. "), ga.ozet_metni(sonuc["tur"], sonuc["ozet"]) + ".",
                     html.Br(), html.Span(f"Öncesinin yedeği: {sonuc['yedek']}" if sonuc["yedek"]
                                          else "Uyarı: işlem öncesi yedek alınamadı (ayrıntı günlük dosyasında).",
                                          className="small-12")], "ok", "check2-circle")
    return _detay_icerik(bid, ust), (yenile or 0) + 1


@callback(
    Output("gc-secili", "data", allow_duplicate=True),
    Input({"type": "gc-git", "id": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def engelleyene_git(_n):
    """Engelleyen içe aktarma düğmesine basılınca onu seçer (detay paneli ona geçer)."""
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update
    return ctx.triggered_id["id"]
