# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: Sistem › Ayarlar (v5).

  - Ayarlar       : proje adı, uyum eşikleri, fazla mesai katsayıları, mesai formu, yedekleme.
                    Kaydedilince doğrulanır, veri_merkezi_ayarlar.json'a yazılır ve hemen uygulanır.
  - Yedekler      : veritabanı yedeklerinin listesi, "Şimdi yedek al", seçilen yedeği geri yükleme
                    (önce yedek incelenir, onay istenir, mevcut hâlin de yedeği alınır).
  - Çevrimdışı    : Bootstrap / ikon / yazı tipi dosyalarının yerelde olup olmadığı ve indirme.
  - Sistem bilgisi: sürüm, şema, veritabanı yolu ve boyutu, günlük kipi (WAL), kayıt sayıları.
"""

import logging
import platform
import sys

from dash import dcc, html, Input, Output, State, ALL, no_update, callback, ctx

import veri_merkezi_ayarlar as ayarlar
import veri_merkezi_sema as sema
import veri_merkezi_sorgu as sorgu
import veri_merkezi_surum as surum
import veri_merkezi_ui as ui
import veri_merkezi_varliklar as varliklar
import veri_merkezi_yedek as yedek

gunluk = logging.getLogger("veri_merkezi")
ACIK = "acik"
YEDEK_TUR_TON = {"acilis": "navy-l", "ice_aktarma": "orange", "geri_alma": "purple", "geri_yukleme": "red",
                 "elle": "green"}


# ---------------------------------------------------------------- ayar formu

def _girdi(anahtar, deger):
    _a, tur, _v, en_az, en_cok, _b, _e, _ac = ayarlar.TANIM[anahtar]
    kimlik = {"type": "ay-alan", "k": anahtar}
    if tur is bool:
        return dcc.Checklist(id=kimlik, options=[{"label": " Açık", "value": ACIK}], value=[ACIK] if deger else [],
                             className="vm-check", inputStyle={"marginRight": "6px"})
    if tur is str:
        return dcc.Input(id=kimlik, type="text", value=deger, maxLength=en_cok, className="vm-input")
    return dcc.Input(id=kimlik, type="number", value=deger, min=en_az, max=en_cok,
                     step=1 if tur is int else 0.1, className="vm-input")


def _form_degeri(anahtar, deger):
    """Arayüzdeki ham değeri ayar doğrulamasına uygun hâle getirir (Checklist -> bool)."""
    if ayarlar.TANIM[anahtar][1] is bool:
        return ACIK in (deger or [])
    return deger


def _varsayilan_metni(v):
    if isinstance(v, bool):
        return "açık" if v else "kapalı"
    if isinstance(v, str):
        return v
    return ui.sayi(v, 0 if float(v).is_integer() else 1)


def _ayar_bolumleri(degerler):
    kartlar = []
    for bolum, baslik in ayarlar.BOLUMLER:
        alanlar = []
        for anahtar, _t, varsayilan, *_x, _b, etiket, aciklama in [t for t in ayarlar.TANIMLAR if t[5] == bolum]:
            vars_metin = _varsayilan_metni(varsayilan)
            alanlar.append(html.Div([
                html.Label(etiket, className="vm-lbl"),
                _girdi(anahtar, degerler[anahtar]),
                html.Span(id={"type": "ay-hata", "k": anahtar}, className="vm-field-msg c-red"),
                html.Span([aciklama, html.Span(f" Varsayılan: {vars_metin}.", className="muted")],
                          className="small-12", style={"color": "#5A6572"}),
            ], className="vm-field"))
        if bolum == "uyum":
            alanlar.append(html.Div([
                html.Label("Standart iş günü (saat)", className="vm-lbl"),
                dcc.Input(type="text", value=ui.sayi(sorgu.STANDART_GUN_SAATI), disabled=True, className="vm-input"),
                html.Span("Sabittir: içe aktarılan günlük toplam saat, içe aktarma anında bu değere göre normal / fazla "
                          "mesaiye bölünüp kalıcı yazılır. Değiştirilirse eski ve yeni kayıtlar farklı kurala göre "
                          "bölünmüş olur.", className="small-12", style={"color": "#5A6572"}),
            ], className="vm-field"))
        kartlar.append(ui.kart([ui.h2(baslik), html.Div(alanlar, className="vm-ay-grid")], "pad-l col",
                               style={"gap": "14px"}))
    return kartlar


# ---------------------------------------------------------------- yedekler

def _yedek_listesi():
    liste = yedek.yedekleri_listele()
    if not liste:
        return ui.bos("Henüz yedek yok. İlk yedek uygulama açılışında ya da ilk içe aktarmada alınır.")
    satirlar = []
    for y in liste[:200]:
        satirlar.append([
            html.Td(y["zaman"].strftime("%d.%m.%Y %H:%M:%S"), className="mono small-12"),
            html.Td(ui.chip(y["tur_etiket"], YEDEK_TUR_TON.get(y["tur"], "gray"))),
            html.Td(y["boyut_metni"], className="mono r small-12"),
            html.Td(y["ad"], className="mono small-12 muted", style={"wordBreak": "break-all"}),
            html.Td(html.Button("Geri yükle", id={"type": "ay-geri-yukle", "ad": y["ad"]}, n_clicks=0,
                                className="vm-btn xs"), className="r"),
        ])
    kat = {}
    for y in liste:
        kat[y["kategori"]] = kat.get(y["kategori"], 0) + 1
    ozet = (f"{ui.sayi(len(liste))} yedek · açılış {kat.get('otomatik', 0)}, işlem öncesi {kat.get('islem', 0)}, "
            f"elle {kat.get('elle', 0)} · toplam {yedek._boyut_metni(sum(y['boyut'] for y in liste))}")
    return html.Div([html.Span(ozet, className="small-12 muted"),
                     html.Div(html.Section(ui.tablo(["Zaman", "Tür", ("Boyut", "r"), "Dosya", ""], satirlar),
                                           className="vm-card clip"),
                              style={"maxHeight": "420px", "overflowY": "auto"})], className="col", style={"gap": "8px"})


def _yedek_karti():
    return ui.kart([
        html.Div([ui.h2("Veritabanı yedekleri", "veritabani_yedekleri/"),
                  html.Button([ui.ikon("folder2-open"), "Klasörü aç"], id="ay-yedek-klasor", n_clicks=0, className="vm-btn sm"),
                  html.Button([ui.ikon("database-down"), "Şimdi yedek al"], id="ay-yedek-al", n_clicks=0,
                              className="vm-btn sm pri")], className="vm-card-head", style={"gap": "8px"}),
        html.Span("Yedekler uygulama açıkken bile tutarlı alınır. Açılışta (günde bir), her içe aktarmadan ve her geri "
                  "almadan önce otomatik alınır. Elle alınan yedekler hiçbir zaman otomatik silinmez.",
                  className="small-12 muted"),
        html.Div(id="ay-yedek-mesaj"),
        html.Div(id="ay-geri-yukle-alan"),
        html.Div(_yedek_listesi(), id="ay-yedek-liste"),
        dcc.Store(id="ay-secili-yedek"),
    ], "pad-l col", style={"gap": "12px"})


# ---------------------------------------------------------------- çevrimdışı dosyalar

def _varlik_durumu():
    d = varliklar.durum()
    yerel = sum(1 for x in d if x["yerel"])
    satirlar = [[html.Td(x["etiket"]),
                 html.Td(ui.chip("yerel", "green") if x["yerel"] else ui.chip("internetten", "orange")),
                 html.Td(yedek._boyut_metni(x["boyut"]) if x["yerel"] else "—", className="mono r small-12")] for x in d]
    if yerel == len(d):
        bant = ui.banner("Bütün arayüz dosyaları yerelde: uygulama internet olmadan da tam görünür.", "ok", "check2-circle")
    elif yerel:
        bant = ui.banner("Bazı dosyalar hâlâ internetten yükleniyor. Tekrar indirmeyi deneyin.", "warn")
    else:
        bant = ui.banner("Arayüz dosyaları internetten yükleniyor. İnternet yokken ikonlar görünmez ve düzen bozulabilir. "
                         "Bir kez indirmeniz yeterli.", "warn")
    return html.Div([bant, ui.tablo(["Dosya", "Kaynak", ("Boyut", "r")], satirlar, "compact")], className="col",
                    style={"gap": "10px"})


def _varlik_karti():
    return ui.kart([
        html.Div([ui.h2("Çevrimdışı çalışma"),
                  html.Button([ui.ikon("cloud-download"), "Dosyaları indir"], id="ay-varlik-indir", n_clicks=0,
                              className="vm-btn sm")], className="vm-card-head"),
        html.Div(_varlik_durumu(), id="ay-varlik-durum"),
        dcc.Loading(html.Div(id="ay-varlik-mesaj"), type="circle"),
        html.Span(["Komut satırından da indirilebilir: ", html.Code("python veri_merkezi_varliklar.py"),
                   ". İndirmeden sonra uygulamayı yeniden başlatın."], className="small-12 muted"),
    ], "pad-l col", style={"gap": "12px"})


# ---------------------------------------------------------------- sistem bilgisi

def _sistem_karti():
    try:
        b = sorgu.get_db_bilgisi()
    except Exception as e:
        return ui.kart([ui.h2("Sistem bilgisi"), ui.banner(f"Veritabanı bilgisi okunamadı: {e}", "err")], "pad-l col")
    try:
        import dash
        dash_surum = dash.__version__
    except Exception:
        dash_surum = "?"
    kip = b["kip"].upper()
    satirlar = [
        ("Uygulama sürümü", f"{surum.etiket()} · {surum.SURUM_ADI}"),
        ("Veritabanı şeması", f"v{sema.get_schema_version()}"),
        ("Veritabanı", b["yol"]),
        ("Boyut", yedek._boyut_metni(b["boyut"]) + (f" + WAL {yedek._boyut_metni(b['wal_boyut'])}" if b["wal_boyut"] else "")),
        ("Günlük kipi", kip + (" (okuma ve yazma birbirini beklemez)" if kip == "WAL" else
                               " — WAL'a geçilemedi, ayrıntı günlük dosyasında")),
        ("Kayıtlar", f"{ui.sayi(b['personel'])} personel · {ui.sayi(b['gunluk_kayit'])} günlük kayıt · "
                     f"{ui.sayi(b['ice_aktarma'])} içe aktarma · {ui.sayi(b['puantaj_izi'])} değişiklik izi"),
        ("Ayar dosyası", str(ayarlar.AYAR_DOSYASI)),
        ("Python · Dash", f"{platform.python_version()} · {dash_surum} · {sys.platform}"),
    ]
    return ui.kart([ui.h2("Sistem bilgisi"),
                    html.Div([ui.bilgi_satiri(a, d, mono=a in ("Veritabanı", "Ayar dosyası")) for a, d in satirlar],
                             className="col", style={"gap": "6px", "wordBreak": "break-all"}),
                    html.Span("Veritabanını elle kopyalayacaksanız uygulamayı kapatın ya da yukarıdaki 'Şimdi yedek al'ı "
                              "kullanın: uygulama açıkken son değişiklikler bir süre -wal dosyasında durur.",
                              className="small-12 muted")], "pad-l col", style={"gap": "12px"})


# ---------------------------------------------------------------- sayfa

def sayfa():
    degerler = ayarlar.hepsi()
    return ui.sayfa("Ayarlar", "Sistem / Ayarlar", [
        html.Div([
            html.Div([
                *_ayar_bolumleri(degerler),
                html.Div([html.Button([ui.ikon("check2"), "Ayarları kaydet"], id="ay-kaydet", n_clicks=0, className="vm-btn pri"),
                          html.Button("Varsayılanları doldur", id="ay-varsayilan", n_clicks=0, className="vm-btn",
                                      title="Formu varsayılan değerlerle doldurur; kaydetmeden hiçbir şey değişmez."),
                          html.Div(id="ay-mesaj", className="grow")], className="row", style={"gap": "8px"}),
            ], className="col", style={"gap": "16px", "minWidth": 0}),
            html.Div([_yedek_karti(), _varlik_karti(), _sistem_karti()], className="col",
                     style={"gap": "16px", "minWidth": 0}),
        ], className="g", style={"gridTemplateColumns": "minmax(0,1fr) minmax(0,1fr)", "gap": "16px",
                                 "alignItems": "start"}),
    ])


# ---------------------------------------------------------------- callback'ler: ayarlar

@callback(
    Output("ay-mesaj", "children"),
    Output({"type": "ay-hata", "k": ALL}, "children"),
    Input("ay-kaydet", "n_clicks"),
    State({"type": "ay-alan", "k": ALL}, "value"),
    State({"type": "ay-alan", "k": ALL}, "id"),
    State({"type": "ay-hata", "k": ALL}, "id"),
    prevent_initial_call=True,
)
def kaydet(n, degerler, idler, hata_idleri):
    bos = [""] * len(hata_idleri)
    if not n:
        return no_update, [no_update] * len(hata_idleri)
    yeni = {i["k"]: _form_degeri(i["k"], v) for i, v in zip(idler, degerler)}
    try:
        ok, hatalar = ayarlar.kaydet(yeni)
    except ayarlar.AyarHatasi as e:
        return ui.banner(str(e), "err"), bos
    if not ok:
        return (ui.banner(f"{len(hatalar)} alan hatalı; hiçbir ayar değiştirilmedi.", "err"),
                [hatalar.get(h["k"], "") for h in hata_idleri])
    return ui.banner("Ayarlar kaydedildi ve uygulandı. Raporlar bir sonraki hesaplamadan itibaren yeni değerleri kullanır.",
                     "ok", "check2-circle"), bos


@callback(
    Output({"type": "ay-alan", "k": ALL}, "value"),
    Input("ay-varsayilan", "n_clicks"),
    State({"type": "ay-alan", "k": ALL}, "id"),
    prevent_initial_call=True,
)
def varsayilanlar(n, idler):
    if not n:
        return [no_update] * len(idler)
    sonuc = []
    for i in idler:
        v = ayarlar.VARSAYILANLAR[i["k"]]
        sonuc.append(([ACIK] if v else []) if isinstance(v, bool) else v)
    return sonuc


# ---------------------------------------------------------------- callback'ler: yedekler

@callback(
    Output("ay-yedek-mesaj", "children"),
    Output("ay-yedek-liste", "children"),
    Input("ay-yedek-al", "n_clicks"),
    prevent_initial_call=True,
)
def simdi_yedek_al(n):
    if not n:
        return no_update, no_update
    try:
        y = yedek.yedek_al("elle")
    except Exception as e:
        return ui.banner(f"Yedek alınamadı: {e}", "err"), no_update
    return ui.banner(f"Yedek alındı: {y['ad']} ({y['boyut_metni']})", "ok", "check2-circle"), _yedek_listesi()


@callback(
    Output("ay-yedek-mesaj", "children", allow_duplicate=True),
    Input("ay-yedek-klasor", "n_clicks"),
    prevent_initial_call=True,
)
def klasoru_ac(n):
    if not n:
        return no_update
    import veri_merkezi_ekip_listesi as el
    klasor = yedek.yedek_klasoru()
    try:
        klasor.mkdir(parents=True, exist_ok=True)
        el.open_path(klasor)
        return None
    except Exception as e:
        return ui.banner([f"Klasör açılamadı ({e}). Konum: ", html.Code(str(klasor))], "warn")


@callback(
    Output("ay-geri-yukle-alan", "children"),
    Output("ay-secili-yedek", "data"),
    Input({"type": "ay-geri-yukle", "ad": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def geri_yukle_sec(_n):
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update, no_update
    ad = ctx.triggered_id["ad"]
    yol = yedek.yedek_bul(ad)
    if yol is None:
        return ui.banner("Yedek dosyası bulunamadı (silinmiş olabilir).", "err"), None
    inc = yedek.yedek_incele(yol)
    if not inc["gecerli"]:
        return ui.banner([html.Strong("Bu yedek geri yüklenemez: "), inc["hata"] or "geçersiz dosya"], "err", "x-octagon"), None
    try:
        simdi = sorgu.get_db_bilgisi()
    except Exception:
        simdi = {"personel": "?", "gunluk_kayit": "?", "ice_aktarma": "?"}
    tablo = ui.tablo(["", ("Şimdiki", "r"), ("Yedekteki", "r")], [
        ["Personel", html.Td(ui.sayi(simdi["personel"]), className="mono r"), html.Td(ui.sayi(inc["personel"]), className="mono r")],
        ["Günlük kayıt", html.Td(ui.sayi(simdi["gunluk_kayit"]), className="mono r"),
         html.Td(ui.sayi(inc["gunluk_kayit"]), className="mono r")],
        ["İçe aktarma", html.Td(ui.sayi(simdi["ice_aktarma"]), className="mono r"),
         html.Td(ui.sayi(inc["ice_aktarma"]), className="mono r")],
    ], "compact")
    return html.Div([
        html.Strong(f"Geri yüklenecek yedek: {ad}", style={"fontSize": "13px", "wordBreak": "break-all"}),
        html.Span(f"Şema v{inc['sema'] or '?'} · son puantaj günü {ui.tarih_tr(inc['son_tarih'])} · bütünlük kontrolü: tamam",
                  className="small-12 muted"),
        tablo,
        ui.banner("Çalışan veritabanının TAMAMI bu yedekteki hâline döner; yedekten sonra yapılan bütün içe aktarmalar "
                  "ve düzenlemeler kaybolur. Geri yüklemeden önce şimdiki hâlin yedeği otomatik alınır, yani bu işlem de "
                  "geri alınabilir.", "warn"),
        html.Div([html.Button([ui.ikon("arrow-counterclockwise"), "Geri yüklemeyi onayla"], id="ay-geri-yukle-onay",
                              n_clicks=0, className="vm-btn sig sm"),
                  html.Button("Vazgeç", id="ay-geri-yukle-vazgec", n_clicks=0, className="vm-btn sm")],
                 className="row", style={"gap": "8px"}),
    ], className="col vm-card pad", style={"gap": "10px", "borderColor": "#F0C9A6"}), ad


@callback(
    Output("ay-geri-yukle-alan", "children", allow_duplicate=True),
    Output("ay-secili-yedek", "data", allow_duplicate=True),
    Input("ay-geri-yukle-vazgec", "n_clicks"),
    prevent_initial_call=True,
)
def geri_yukle_vazgec(n):
    return (None, None) if n else (no_update, no_update)


@callback(
    Output("ay-geri-yukle-alan", "children", allow_duplicate=True),
    Output("ay-yedek-liste", "children", allow_duplicate=True),
    Output("ay-secili-yedek", "data", allow_duplicate=True),
    Input("ay-geri-yukle-onay", "n_clicks"),
    State("ay-secili-yedek", "data"),
    prevent_initial_call=True,
)
def geri_yukle_onayla(n, ad):
    if not n or not ad:
        return no_update, no_update, no_update
    yol = yedek.yedek_bul(ad)
    if yol is None:
        return ui.banner("Yedek dosyası bulunamadı.", "err"), no_update, None
    try:
        sonuc = yedek.geri_yukle(yol)
    except Exception as e:
        gunluk.exception("yedek geri yüklenemedi: %s", ad)
        return ui.banner([html.Strong("Geri yüklenemedi, veritabanı değişmedi. "), str(e)], "err", "x-octagon"), \
            _yedek_listesi(), None
    return ui.banner([html.Strong("Yedek geri yüklendi. "),
                      f"Geri yüklemeden önceki hâlin yedeği: {sonuc['onceki_yedek']}. " if sonuc["onceki_yedek"] else "",
                      "Diğer sayfalar yeni veriyle açılır; açık bir önizleme varsa yeniden yükleyin."],
                     "ok", "check2-circle"), _yedek_listesi(), None


# ---------------------------------------------------------------- callback'ler: çevrimdışı

@callback(
    Output("ay-varlik-mesaj", "children"),
    Output("ay-varlik-durum", "children"),
    Input("ay-varlik-indir", "n_clicks"),
    prevent_initial_call=True,
)
def varliklari_indir(n):
    if not n:
        return no_update, no_update
    rapor = varliklar.indir()
    basarili = [r for r in rapor if r["ok"]]
    hatali = [r for r in rapor if not r["ok"]]
    if not hatali:
        m = ui.banner("İndirildi. Değişikliğin geçerli olması için uygulamayı kapatıp yeniden açın.", "ok", "check2-circle")
    else:
        m = ui.banner([html.Strong(f"{len(hatali)} dosya indirilemedi"),
                       " (internet bağlantısını kontrol edin; mevcut dosyalar değişmedi): ",
                       "; ".join(f"{r['etiket']}: {r['hata']}" for r in hatali),
                       f". {len(basarili)} dosya indirildi." if basarili else ""], "err")
    return m, _varlik_durumu()
