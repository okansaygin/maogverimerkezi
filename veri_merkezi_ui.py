# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - ORTAK ARAYÜZ BİLEŞENLERİ (veri_merkezi_ui.py)
================================================================
Yeni tasarımın tekrar kullanılan parçaları: üst çubuk, gösterge kartı,
çip, tablo, çubuk, segment düğmesi, biçimlendiriciler. Görünüm
assets/veri_merkezi.css'te tanımlıdır; burada sadece sınıf adları verilir.
"""

__version__ = "2026-09-24.1"

import datetime
import hashlib

from dash import dcc, html, callback, ctx, no_update, Input, Output, MATCH, ALL
import dash_bootstrap_components as dbc

import veri_merkezi_sorgu as sorgu

GUN_KISA = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]
AY_ADI = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
AY_KISA = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"]
SEVIYE_CHIP = {"kritik": ("KRİTİK", "ch-red-s"), "uyari": ("UYARI", "ch-orange"), "bilgi": ("BİLGİ", "ch-gray")}
AVATAR_RENK = ["#1F4E78", "#16344F", "#3E4751", "#9A3A0B", "#2F6B3A"]


# ---------------------------------------------------------------- biçimlendirme

def sayi(x, ondalik=0):
    """Türkçe sayı biçimi: 12.345,6"""
    if x is None:
        return "—"
    try:
        s = f"{float(x):,.{ondalik}f}"
    except (TypeError, ValueError):
        return str(x)
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def saat(x):
    """Saat değeri: tam sayıysa ondalıksız, değilse 1 ondalık."""
    if x is None:
        return "—"
    return sayi(x, 0 if float(x).is_integer() else 1)


def tl(x):
    return f"{sayi(x, 0)} TL"


def tarih_tr(v, saatli=False):
    if not v:
        return "—"
    if isinstance(v, (datetime.date, datetime.datetime)):
        d = v
    else:
        s = str(v)
        try:
            d = datetime.datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S") if len(s) > 10 else datetime.date.fromisoformat(s[:10])
        except ValueError:
            return s
    if saatli and isinstance(d, datetime.datetime):
        return d.strftime("%d.%m %H:%M")
    return d.strftime("%d.%m.%Y")


def ay_etiketi(ay):
    y, m = ay.split("-")
    return f"{AY_ADI[int(m) - 1]} {y}"


def maske_tc(tc):
    tc = str(tc or "")
    if len(tc) < 7:
        return tc or "—"
    return f"{tc[:3]}•••••{tc[-3:]}"


def bas_harf(isim):
    parcalar = [p for p in str(isim or "?").split() if p]
    return "".join(p[0] for p in parcalar[:2]).upper() or "?"


def avatar_renk(anahtar):
    h = int(hashlib.md5(str(anahtar).encode("utf-8")).hexdigest(), 16)
    return AVATAR_RENK[h % len(AVATAR_RENK)]


def tr_upper(s):
    return (s or "").replace("i", "İ").replace("ı", "I").upper()


# ---------------------------------------------------------------- küçük bileşenler

def ikon(ad, className="", **kw):
    return html.I(className=f"bi bi-{ad} {className}".strip(), **kw)


def kart(children, cls="pad", **kw):
    return html.Div(children, className=f"vm-card {cls}".strip(), **kw)


def kpi(etiket, deger, renk=None, not_=None, buyuk=False):
    return html.Div([
        html.Span(etiket, className="vm-lbl"),
        html.Span(deger, className=f"vm-kpi-val {('c-' + renk) if renk else ''}"),
        html.Span(not_, className="vm-kpi-note") if not_ else None,
    ], className=f"vm-card vm-kpi {'big' if buyuk else ''}")


def h2(metin, alt=None, buyuk=False):
    return html.H2([metin, html.Span(f" · {alt}", className="sub") if alt else None],
                   className=f"vm-h2 {'l' if buyuk else ''}")


def chip(metin, ton="gray", **kw):
    return html.Span(metin, className=f"vm-chip ch-{ton}", **kw)


def seviye_chip(seviye):
    t, c = SEVIYE_CHIP.get(seviye, SEVIYE_CHIP["uyari"])
    return html.Span(t, className=f"vm-chip {c}")


def dot(seviye):
    return html.Span(className=f"vm-dot dot-{seviye}")


def lejant(renk, metin):
    return html.Span([html.I(style={"background": renk}), metin], className="vm-legend")


def lbl(metin, **kw):
    return html.Span(metin, className="vm-lbl", **kw)


def bar(parcalar, cls=""):
    """parcalar: [(yüzde, css_sınıfı)]"""
    return html.Div([html.Div(className=c, style={"width": f"{max(0.0, min(100.0, p)):.1f}%"}) for p, c in parcalar],
                    className=f"vm-bar {cls}")


def bilgi_satiri(etiket, deger, mono=False):
    return html.Div([html.Span(etiket, className="muted grow"),
                     html.Span(deger, className="mono" if mono else "")], className="row",
                    style={"gap": "8px", "fontSize": "13px"})


def banner(icerik, tur="warn", ikon_adi="exclamation-triangle"):
    return html.Div([ikon(ikon_adi), html.Div(icerik, className="grow")], className=f"vm-banner {tur}")


def bos(metin):
    return html.Div(metin, className="vm-empty")


def tablo(basliklar, satirlar, cls=""):
    """basliklar: [str | (str, 'r'|'c')]; satirlar: [html.Tr] ya da [[hücre,...]]"""
    th = []
    for b in basliklar:
        if isinstance(b, tuple):
            th.append(html.Th(b[0], className=b[1]))
        else:
            th.append(html.Th(b))
    govde = []
    for s in satirlar:
        govde.append(s if isinstance(s, html.Tr) else html.Tr([c if isinstance(c, html.Td) else html.Td(c) for c in s]))
    return html.Table([html.Thead(html.Tr(th)), html.Tbody(govde)], className=f"vm-tbl {cls}")


def segment(id_, secenekler, deger, blok=False):
    """Segment düğme grubu (tek seçim). secenekler: [(değer, etiket)]"""
    return dbc.RadioItems(
        id=id_, options=[{"label": e, "value": v} for v, e in secenekler], value=deger,
        className=f"vm-seg {'block' if blok else ''}", inputClassName="form-check-input",
        labelClassName="vm-seg-btn", labelCheckedClassName="on", inline=True,
    )


def dropdown(id_, secenekler, deger=None, coklu=False, yer_tutucu="Tümü", **kw):
    opts = [s if isinstance(s, dict) else {"label": s, "value": s} for s in secenekler]
    return html.Div(dcc.Dropdown(id=id_, options=opts, value=deger, multi=coklu, placeholder=yer_tutucu, **kw),
                    className="vm-dd")


def alan(etiket, bilesen, for_=None):
    return html.Div([html.Label(etiket, className="vm-lbl", htmlFor=for_), bilesen], className="vm-field")


def kisi_link(tc, isim, href=None, renk_anahtari=None):
    return dcc.Link([
        html.Span(bas_harf(isim), className="vm-av", style={"background": avatar_renk(renk_anahtari or tc)}),
        html.Span(isim or "(isimsiz)", style={"fontWeight": 600}),
    ], href=href or f"/personel/{tc}", className="vm-person")


def yukleme_alani(id_, metin="Excel dosyasını buraya sürükle veya ", kucuk=False, coklu=False):
    icerik = html.Div([ikon("cloud-arrow-up"), metin, html.A("dosya seç")]) if not kucuk else \
        html.Div([ikon("upload"), metin])
    return dcc.Upload(id=id_, children=icerik, className=f"vm-drop {'sm' if kucuk else ''}", multiple=coklu,
                      accept=".xlsx,.xlsm")


# ---------------------------------------------------------------- üst çubuk

def ust_cubuk(baslik, yol, sag=None):
    """Her sayfanın üst çubuğu: yol + başlık, genel personel araması (Ctrl K) ve sayfaya özel sağ öğeler."""
    # PERFORMANS: tüm personel listesi (1000+ seçenek) her sayfada tarayıcıya gönderilmez;
    # seçenekler yazdıkça sunucudan gelir (veri_merkezi_app.genel_arama_secenekleri).
    arama = html.Div(dcc.Dropdown(id="global-arama", options=[], placeholder="Personel, TC ara…   Ctrl K",
                                  clearable=True, searchable=True, optionHeight=36),
                     className="vm-dd vm-search")
    return html.Header([
        html.Div([html.Span(yol, className="vm-crumb"), html.H1(baslik)], className="vm-top-title"),
        arama,
        *(sag or []),
        dcc.Link([ikon("upload"), "İçe Aktar"], href="/ice-aktar", className="vm-btn pri"),
    ], className="vm-top")


def sayfa(baslik, yol, govde, sag=None, govde_cls="vm-page"):
    """govde_cls=None ise govde olduğu gibi yerleştirilir (bölmeli sayfalar kendi ızgarasını kurar)."""
    icerik = govde if govde_cls is None else html.Div(govde, className=govde_cls)
    return html.Div([ust_cubuk(baslik, yol, sag), icerik],
                    style={"display": "flex", "flexDirection": "column", "flexGrow": 1, "minHeight": "100vh"})


# ---------------------------------------------------------------- sekmeler (dbc.Tabs yerine)

def sekmeler(grup, parcalar, aktif=None):
    """Düz HTML sekmeler. parcalar: [(anahtar, etiket, içerik)].

    NEDEN dbc.Tabs DEĞİL: dbc.Tabs, içeriği callback'le sonradan doldurulan sekmelerde tarayıcıda çizim
    hatası verebiliyor. Dash (2.x) böyle bir hatada bütün sayfa durumunu bir adım GERİ ALIR ("revert");
    geri alınan durumda adres de önceki sayfa olduğu için dcc.Location sayfayı oraya yeniler. Verimlilik
    Analizi'nin açılır açılmaz Genel Bakış'a dönmesinin sebebi buydu. Bu bileşen yalnızca html.Button ve
    html.Div kullanır; sekme değişimi tek bir callback'le yalnızca görünürlüğü (display) değiştirir.
    Bütün paneller sayfada durur, bu yüzden içlerine yazan callback'ler her zaman çalışır.
    grup: sayfa içinde benzersiz kısa ad (ör. "va")."""
    aktif = aktif if aktif is not None else (parcalar[0][0] if parcalar else None)
    dugmeler = [html.Li(html.Button(etiket, id={"type": "vm-sekme-dugme", "grup": grup, "key": k}, n_clicks=0,
                                    type="button", className="nav-link active" if k == aktif else "nav-link"),
                        className="nav-item") for k, etiket, _ic in parcalar]
    paneller = [html.Div(ic, id={"type": "vm-sekme-panel", "grup": grup, "key": k},
                         style={} if k == aktif else {"display": "none"}) for k, _e, ic in parcalar]
    return html.Div([html.Ul(dugmeler, className="nav nav-tabs", role="tablist"),
                     html.Div(paneller, className="tab-content")], className="vm-tabs")


@callback(
    Output({"type": "vm-sekme-panel", "grup": MATCH, "key": ALL}, "style"),
    Output({"type": "vm-sekme-dugme", "grup": MATCH, "key": ALL}, "className"),
    Input({"type": "vm-sekme-dugme", "grup": MATCH, "key": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def _sekme_sec(_tiklar):
    secilen = ctx.triggered_id
    idler = ctx.outputs_list[1] if ctx.outputs_list else []
    if not secilen or not (ctx.triggered and ctx.triggered[0]["value"]):
        return [no_update] * len(idler), [no_update] * len(idler)
    anahtar = secilen["key"]
    return ([{} if i["id"]["key"] == anahtar else {"display": "none"} for i in idler],
            ["nav-link active" if i["id"]["key"] == anahtar else "nav-link" for i in idler])
