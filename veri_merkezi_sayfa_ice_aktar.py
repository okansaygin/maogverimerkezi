# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: İçe Aktarma Sihirbazı.

Eski iki sayfa ("İçe Aktar - Personel Bilgileri" ve "İçe Aktar - Çalışma
Bilgileri") tek akışta birleşti: dosya türü başlıklardan otomatik anlaşılır,
sonra ilgili motor (veri_merkezi_ice_aktarma / veri_merkezi_mesai_ice_aktarma)
AYNEN çalıştırılır. Önizleme -> onay -> tek transaction'da yazma mantığı
değişmedi."""

import base64
import datetime
import tempfile
from pathlib import Path

from dash import dcc, html, Input, Output, State, ALL, no_update, callback, ctx

import puantaj_engine as engine
import veri_merkezi_ice_aktarma as ie
import veri_merkezi_mesai_ice_aktarma as mie
import veri_merkezi_dosya_tespit as tespit
import veri_merkezi_sorgu as sorgu
import veri_merkezi_ui as ui

# Tek kullanıcılı masaüstü uygulaması: önizleme sonuçları (büyük olabilir)
# istemciye taşınmaz, sunucu belleğinde tutulur.
_BEKLEYEN = {}
_SAYAC = {"n": 0}

TMP_DIR = Path(tempfile.gettempdir()) / "veri_merkezi_yuklemeler"


def gecici_kaydet(contents, filename):
    _ctype, veri = contents.split(",", 1)
    TMP_DIR.mkdir(exist_ok=True)
    yol = TMP_DIR / Path(filename).name
    with open(yol, "wb") as f:
        f.write(base64.b64decode(veri))
    return yol


# ---------------------------------------------------------------- parçalar

def adimlar(simdiki, ucuncu="Önizleme & çakışmalar"):
    etiketler = ["Dosya", "Sütun eşleştirme", ucuncu, "Onay"]
    parca = []
    for i, e in enumerate(etiketler, start=1):
        durum = "done" if i < simdiki else ("now" if i == simdiki else "todo")
        nokta = ui.ikon("check-lg") if durum == "done" else str(i)
        parca.append(html.Div([html.Span(nokta, className="vm-step-dot"), html.Span(e)], className=f"vm-step {durum}"))
        if i < len(etiketler):
            parca.append(html.Span(className="vm-step-line"))
    return html.Div(parca, className="vm-card vm-steps")


def _dosya_karti(dosya_adi, t, preview=None):
    tur_chip = {"kimlik": ui.chip("Personel Bilgileri", "navy"), "mesai": ui.chip("Çalışma Bilgileri (SAP mesai)", "navy"),
                None: ui.chip("Tanınmadı", "red")}[t["tur"]]
    aciklama = {
        "kimlik": "Başlıklara bakılarak aylık kimlik dosyası olduğu anlaşıldı. Sadece kimlik bilgisi (ad, TC, grup, "
                  "ekip, görev, giriş/çıkış) alınır; çalışma saatleri alınmaz.",
        "mesai": "Başlıklara bakılarak günlük SAP mesai dosyası olduğu anlaşıldı. Her satır bir kişinin bir günüdür; "
                 "toplam saat 9 saatlik güne göre normal/fazla olarak bölünür.",
        None: "Ne kimlik ne de SAP mesai dosyasının zorunlu başlıkları bulundu. Aşağıda eksik başlıklar işaretli.",
    }[t["tur"]]
    satirlar = [ui.bilgi_satiri("Veri satırları", f"{ui.sayi(t['son_veri_satiri'])} satıra kadar", True)]
    if preview:
        satirlar.insert(0, ui.bilgi_satiri("Dosya özeti (SHA-256)", f"{preview['dosya_hash'][:4]}…{preview['dosya_hash'][-4:]}", True))
        satirlar.insert(1, ui.bilgi_satiri("Daha önce aktarıldı mı?", html.Span("Hayır", className="c-green",
                                                                              style={"fontWeight": 600})))
        if preview.get("donem_baslangic") and t["tur"] == "mesai":
            d = ui.tarih_tr(preview["donem_baslangic"])
            if preview["donem_bitis"] != preview["donem_baslangic"]:
                d += f" – {ui.tarih_tr(preview['donem_bitis'])}"
            satirlar.append(ui.bilgi_satiri("Tarih", d, True))
    return ui.kart([
        html.Div([
            html.Span(ui.ikon("file-earmark-spreadsheet"), className="vm-pipe-ic pi-navy",
                      style={"width": "40px", "height": "40px", "borderRadius": "8px", "fontSize": "20px"}),
            html.Div([html.Span(dosya_adi, className="mono", style={"fontWeight": 600, "wordBreak": "break-all"}),
                      html.Span(f"Sayfa: {t['sayfa']} · {ui.sayi(t['boyut_kb'])} KB", className="small-12 muted")],
                     className="col", style={"gap": "0"}),
        ], className="row"),
        html.Div([ui.lbl("Algılanan dosya türü"),
                  html.Div([tur_chip, html.Span(f"{t['bulunan']}/{t['beklenen']} başlık eşleşti", className="small-12 muted")],
                           className="row", style={"gap": "8px"}),
                  html.Span(aciklama, className="small-12 muted")],
                 className="col", style={"padding": "12px", "borderRadius": "8px", "background": "#F4F2EE"}),
        html.Div(satirlar, className="col", style={"gap": "4px"}),
    ], "pad col", style={"gap": "12px"})


def _sutun_tablosu(t):
    """Sütun eşleştirme tablosu. Her alanın sütun harfi DÜZENLENEBİLİR: başlık tanınmadıysa ya da yanlış
    sütun seçildiyse kullanıcı doğru harfi yazıp "Eşleştirmeyi uygula" ile yeniden dener."""
    satirlar = []
    for e in t["eslesme"]:
        if e["bulundu"]:
            durum = ui.ikon("pencil-fill" if e["elle"] else "check-lg", className="c-navy" if e["elle"] else "c-green",
                            title="Elle seçildi" if e["elle"] else "Başlıktan bulundu")
        else:
            durum = ui.ikon("x-lg", className="c-red" if e["zorunlu"] else "c-muted",
                            title="Zorunlu alan bulunamadı" if e["zorunlu"] else "Bulunamadı (isteğe bağlı)")
        eksik_zorunlu = e["zorunlu"] and not e["bulundu"]
        satirlar.append(html.Tr([
            html.Td([html.Div(e["etiket"], style={"fontWeight": 600, "fontSize": "13px", "minWidth": "104px"}),
                     html.Span("ZORUNLU", style={"fontSize": "10px", "fontWeight": 700, "color": "#B42318" if eksik_zorunlu
                                                 else "#1F4E78"}) if e["zorunlu"] else None]),
            html.Td(e["baslik"], className="mono", style={"fontSize": "11.5px", "maxWidth": "92px", "overflow": "hidden",
                                                         "textOverflow": "ellipsis", "whiteSpace": "nowrap"},
                    title=e["baslik"]),
            # NOT: dcc.Input "aria-*" / "data-*" gibi joker özellikleri KABUL ETMEZ (sadece html.* kabul eder);
            # verilirse TypeError olur ve yükleme sessizce durur. Erişilebilir ad için Td'nin title'ı kullanılır.
            html.Td(dcc.Input(id={"type": "ia-sutun", "alan": e["alan"]}, value=e["sutun"], type="text", maxLength=3,
                              placeholder="?", className="vm-input mono vm-colinput"),
                    title=f"{e['etiket']} sütun harfi"),
            html.Td(e["ornek"] or "—", className="mono muted", style={"fontSize": "11.5px", "maxWidth": "92px",
                                                                     "overflow": "hidden", "textOverflow": "ellipsis",
                                                                     "whiteSpace": "nowrap"}, title=e["ornek"]),
            html.Td(durum),
        ], className="err" if eksik_zorunlu else ""))
    return html.Section([
        html.Div([ui.h2("Sütun eşleştirme"), html.Span("harfi değiştirebilirsiniz", className="small-12 muted")],
                 className="vm-card-head pad"),
        html.Div([
            html.Span("Dosya türü", className="vm-lbl"),
            ui.segment("ia-tur", [("kimlik", "Personel Bilgileri"), ("mesai", "SAP Mesai")], t.get("tahmin") or "kimlik",
                       blok=True),
        ], className="col", style={"gap": "6px", "padding": "0 18px 10px"}),
        ui.tablo(["Alan", "Excel başlığı", ("Sütun", "c"), "Örnek değer", ""], satirlar, cls="vm-colmap"),
        html.Div([
            html.Span("Başlık tanınmadıysa ya da yanlış sütun seçildiyse doğru sütun harfini (ör. B) yazın. "
                      "Örnek değer sütununda o sütunun ilk verisi görünür.", className="small-12 muted"),
            html.Button([ui.ikon("arrow-repeat"), "Eşleştirmeyi uygula ve yeniden dene"], id="ia-yeniden",
                        className="vm-btn pri block", n_clicks=0),
        ], className="col", style={"gap": "8px", "padding": "12px 18px 16px", "borderTop": "1px solid #EFEBE4"}),
    ], className="vm-card clip")


def _uyari_cipleri(uyarilar):
    say = {}
    for a in uyarilar:
        say[a["tip"]] = say.get(a["tip"], 0) + 1
    ton = {"kritik": "red", "uyari": "orange", "bilgi": "gray"}
    cipler = [ui.chip(f"{sorgu.UYARI_ETIKET.get(t, t)} · {n}", ton[sorgu.uyari_seviyesi(t)])
              for t, n in sorted(say.items(), key=lambda kv: -kv[1])]
    return ui.kart([
        html.Div([ui.h2(f"Uyarılar · {len(uyarilar)}"),
                  html.Span("Onaydan sonra Veri Kalitesi'ne kaydedilir", className="small-12 muted")], className="vm-card-head"),
        html.Div(cipler, className="row", style={"flexWrap": "wrap", "gap": "8px"}) if cipler
        else ui.banner("Uyarı yok.", "ok", "check2-circle"),
    ], "pad col", style={"gap": "10px"})


def _cakisma_karti(conflicts):
    if not conflicts:
        return ui.kart(ui.banner("Veri setindeki kayıtlarla çakışan bilgi yok — mevcut kişiler için sadece boş alanlar "
                                 "doldurulacak ya da hiçbir şey değişmeyecek.", "ok", "check2-circle"))
    satirlar = []
    for c in conflicts:
        satirlar.append([
            html.Td(html.Div([html.Span(c["ad_soyad_baglam"], style={"fontWeight": 600}),
                              html.Span(ui.maske_tc(c["tc"]), className="vm-sub")], className="col", style={"gap": 0})),
            html.Td(c["alan_etiket"], className="muted"),
            html.Td(html.Div([html.Span(c["eski_deger"] or "—", style={"fontWeight": 600}),
                              html.Span(str(c["eski_kaynak"]).replace("veri_merkezi:", "").replace("excel:", ""), className="vm-sub", style={"fontSize": "11px"})],
                             className="col", style={"gap": 0})),
            html.Td(html.Span(c["yeni_deger"] or "—", className="c-navy", style={"fontWeight": 600})),
            html.Td(ui.segment({"type": "conflict-choice", "tc": c["tc"], "field": c["alan"]},
                               [("existing", "Mevcut"), ("new", "Yeni")], "existing")),
        ])
    return html.Section([
        html.Div([
            ui.h2("Bilgi çakışmaları", "veri setindeki kayıttan farklı gelenler"),
            html.Button("Tümünde mevcudu koru", id="ia-hepsi-mevcut", className="vm-btn sm", n_clicks=0),
            html.Button("Tümünde yeniyi kullan", id="ia-hepsi-yeni", className="vm-btn sm", n_clicks=0),
        ], className="vm-card-head line"),
        html.Div(ui.tablo(["Personel", "Alan", "Veri setindeki (mevcut)", "Bu dosyadaki (yeni)", "Karar"], satirlar),
                 style={"maxHeight": "440px", "overflow": "auto"}),
    ], className="vm-card clip")


def _alt_cubuk(etiket, aktif=True):
    return html.Footer([
        html.Span([ui.ikon("database"), " Tek işlemde yazılır — bir hata olursa hiçbir şey yazılmaz, içe aktarma "
                                        "geçmişine “hata” olarak düşer."], className="grow small-12 muted",
                  style={"fontSize": "13px"}),
        dcc.Link("Başka dosya yükle", href="/ice-aktar", className="vm-btn", refresh=True),
        html.Button([ui.ikon("check2"), etiket], id="ia-onay", className="vm-btn pri", n_clicks=0, disabled=not aktif),
    ], className="vm-footbar")


# ---------------------------------------------------------------- sayfa

def sayfa():
    return ui.sayfa("İçe Aktarma Sihirbazı", "Veri Girişi / İçe Aktar", html.Div([
        dcc.Loading(html.Div(ilk_gorunum(), id="ia-govde", style={"display": "flex", "flexDirection": "column", "flexGrow": 1}),
                    type="default", color="#1F4E78", parent_style={"display": "flex", "flexDirection": "column", "flexGrow": 1}),
        dcc.Store(id="ia-pid"),
        dcc.Store(id="ia-dosya"),
    ], style={"display": "flex", "flexDirection": "column", "flexGrow": 1}), govde_cls=None)


def ilk_gorunum():
    tur_kart = lambda ik, b, a: ui.kart([
        html.Div([html.Span(ui.ikon(ik), className="vm-pipe-ic pi-navy"), html.Span(b, style={"fontWeight": 600})],
                 className="row", style={"gap": "10px"}),
        html.Span(a, className="small-12 muted")], "pad col", style={"gap": "8px"})
    return html.Div([
        adimlar(1),
        ui.kart(ui.yukleme_alani("ia-upload", "Personel veya SAP mesai Excel'ini buraya sürükle ya da "), "pad"),
        html.Div([
            tur_kart("person-vcard", "Personel Bilgileri (aylık kimlik dosyası)",
                     "Ad soyad, TC, grup, bağlı olduğu ekip, görev, işe giriş/çıkış tarihi. Başlıklar 1–4. satırlarda "
                     "aranır; zorunlu: TC KİMLİK NO, ADI SOYADI, GRUPLAR."),
            tur_kart("calendar-check", "Çalışma Bilgileri (günlük SAP mesai dosyası)",
                     "Her satır bir kişinin bir günü. Zorunlu: TC Kimlik No, Tarih, Rönesans Çalışma Saati Girişi. "
                     "SAP Puantaj Aktarım'ın ürettiği dosya doğrudan yüklenebilir."),
        ], className="g g-2"),
    ], className="vm-page")


def _hata_gorunumu(dosya_adi, t, mesajlar):
    return html.Div([
        adimlar(2),
        html.Div([
            html.Div([_dosya_karti(dosya_adi, t), _sutun_tablosu(t)], className="col", style={"gap": "16px"}),
            html.Div([ui.banner([html.Strong("İçe aktarma durduruldu. "), html.Ul([html.Li(m) for m in mesajlar],
                                                                                   style={"margin": "6px 0 0"})], "err",
                                "x-octagon"),
                      ui.banner([html.Strong("Sütunlar eşleşmediyse: "),
                                 "soldaki Sütun eşleştirme tablosunda eksik alanların sütun harfini yazın (gerekirse dosya "
                                 "türünü de seçin) ve ", html.Strong("Eşleştirmeyi uygula ve yeniden dene"),
                                 "'ye basın. Dosyayı tekrar yüklemeniz gerekmez."], "info", "info-circle"),
                      ui.kart(ui.yukleme_alani("ia-upload", "ya da düzeltilmiş dosyayı yükle — ", kucuk=True))],
                     className="col", style={"gap": "16px"}),
        ], className="g", style={"gridTemplateColumns": "460px minmax(0,1fr)"}),
    ], className="vm-page")


def _kimlik_gorunumu(dosya_adi, t, preview):
    sayi = tespit.personel_onizleme_sayilari(preview)
    conflicts = preview.get("field_conflicts") or []
    yazilacak = sayi["yeni"] + sayi["guncellenecek"]
    return [
        html.Div([
            adimlar(3),
            html.Div([
                html.Div([_dosya_karti(dosya_adi, t, preview), _sutun_tablosu(t)], className="col", style={"gap": "16px"}),
                html.Div([
                    html.Section([
                        ui.kpi("Taranan satır", ui.sayi(preview["taranan_satir"])),
                        ui.kpi("Yeni personel", f"+{ui.sayi(sayi['yeni'])}", "green"),
                        ui.kpi("Güncellenecek", ui.sayi(sayi["guncellenecek"]), "navy"),
                        ui.kpi("Değişmeyen", ui.sayi(sayi["degismeyen"]), "muted"),
                        html.Div([ui.lbl("Karar bekleyen", style={"color": "#9A3A0B"}),
                                  html.Span(ui.sayi(len(conflicts)), className="vm-kpi-val c-orange")],
                                 className="vm-card vm-kpi warn" if conflicts else "vm-card vm-kpi"),
                    ], className="g g-5 gap-12"),
                    _cakisma_karti(conflicts),
                    _uyari_cipleri(preview["uyarilar"]),
                ], className="col", style={"gap": "16px", "minWidth": 0}),
            ], className="g", style={"gridTemplateColumns": "460px minmax(0,1fr)"}),
        ], className="vm-page"),
        _alt_cubuk(f"Onayla ve kalıcı depoya yaz · {ui.sayi(yazilacak)} kişi"),
    ]


def _mesai_gorunumu(dosya_adi, t, preview):
    sayi = tespit.mesai_onizleme_sayilari(preview)
    gunler = {}
    for k in preview["gunluk_kayitlar"]:
        g = gunler.setdefault(k["tarih"], {"n": 0, "fazla": 0.0})
        g["n"] += 1
        g["fazla"] += k["fazla_saat"]
    gun_satirlari = [[html.Td(ui.tarih_tr(d), className="mono"),
                      html.Td(ui.GUN_KISA[datetime.date.fromisoformat(d).weekday()], className="muted"),
                      html.Td(ui.sayi(v["n"]), className="mono r"), html.Td(ui.saat(v["fazla"]), className="mono r c-orange")]
                     for d, v in sorted(gunler.items())]
    kayitsiz = sum(1 for a in preview["uyarilar"] if a["tip"] == "personel_kaydi_yok")
    return [
        html.Div([
            adimlar(3, "Önizleme"),
            html.Div([
                html.Div([_dosya_karti(dosya_adi, t, preview), _sutun_tablosu(t)], className="col", style={"gap": "16px"}),
                html.Div([
                    html.Section([
                        ui.kpi("Taranan satır", ui.sayi(preview["taranan_satir"])),
                        ui.kpi("Kişi", ui.sayi(sayi["kisi"])),
                        ui.kpi("Yeni günlük kayıt", f"+{ui.sayi(sayi['yeni'])}", "green"),
                        ui.kpi("Güncellenecek", ui.sayi(sayi["guncellenecek"]), "navy"),
                        ui.kpi("Değişmeyen", ui.sayi(sayi["degismeyen"]), "muted"),
                    ], className="g g-5 gap-12"),
                    ui.banner([html.Strong(f"{kayitsiz} kişinin personel kaydı yok. "),
                               "Saatleri yine de yazılır; grup/ekip/görev boş görünür. Önce Personel Bilgileri "
                               "dosyasını yüklemeniz önerilir."], "warn") if kayitsiz else None,
                    html.Section([
                        html.Div([ui.h2("Gün bazında özet")], className="vm-card-head pad"),
                        html.Div(ui.tablo(["Tarih", "Gün", ("Kayıt", "r"), ("Fazla mesai (sa)", "r")], gun_satirlari),
                                 style={"maxHeight": "360px", "overflowY": "auto"}),
                    ], className="vm-card clip"),
                    _uyari_cipleri(preview["uyarilar"]),
                ], className="col", style={"gap": "16px", "minWidth": 0}),
            ], className="g", style={"gridTemplateColumns": "460px minmax(0,1fr)"}),
        ], className="vm-page"),
        _alt_cubuk(f"Onayla ve kalıcı depoya yaz · {ui.sayi(sayi['yeni'] + sayi['guncellenecek'])} kayıt",
                   aktif=bool(preview["gunluk_kayitlar"])),
    ]


# ---------------------------------------------------------------- callback'ler

def _beklenmeyen_hata(e):
    """Callback içinde yakalanmayan hata Dash'te SESSİZCE yutulur (ekran hiç değişmez). Bunun yerine hatayı
    ekranda ve konsolda göster ki kullanıcı ne olduğunu görebilsin."""
    import traceback
    traceback.print_exc()
    return [html.Div([adimlar(1),
                      ui.banner([html.Strong("Beklenmeyen bir hata oluştu. "), f"{type(e).__name__}: {e}",
                                 html.Br(), "Bu mesajı geliştiriciye iletin."], "err", "x-octagon"),
                      ui.kart(ui.yukleme_alani("ia-upload", "Tekrar dene — ", kucuk=True))],
                     className="vm-page")]


def _isle(yol, filename, tur=None, override=None, ek_mesaj=None):
    """Tespit -> önizleme. Döner: (görünüm, pid). tur/override: kullanıcının elle seçtikleri."""
    try:
        t = tespit.dosya_turu_tespit(yol, tur_zorla=tur, override=override)
    except Exception as e:
        return [html.Div([adimlar(1), ui.banner(f"Dosya okunamadı: {e}", "err", "x-octagon"),
                          ui.kart(ui.yukleme_alani("ia-upload", "Başka bir dosya dene — ", kucuk=True))],
                         className="vm-page")], None
    try:
        if ek_mesaj:
            return _hata_gorunumu(filename, t, ek_mesaj), None
        if t["tur"] is None:
            eksik = [f"'{e['etiket']}' sütunu bulunamadı" for e in t["eslesme"] if e["zorunlu"] and not e["bulundu"]]
            return _hata_gorunumu(filename, t, eksik or ["Dosya türü anlaşılamadı."]), None
    except Exception as e:
        return _beklenmeyen_hata(e), None
    try:
        cm = t["override"] or None
        preview = (ie.preview_import(yol, colmap_override=cm) if t["tur"] == "kimlik"
                   else mie.preview_mesai_import(yol, colmap_override=cm))
    except Exception as e:
        return _hata_gorunumu(filename, t, [f"Dosya işlenemedi: {e}. Seçilen sütunları kontrol edin."]), None
    if preview["durum"] == "engellendi":
        return _hata_gorunumu(filename, t, [a["detay"] for a in preview["engelleyiciler"]]), None
    try:
        gorunum = (_kimlik_gorunumu(filename, t, preview) if t["tur"] == "kimlik"
                   else _mesai_gorunumu(filename, t, preview))
    except Exception as e:
        return _beklenmeyen_hata(e), None
    _SAYAC["n"] += 1
    pid = _SAYAC["n"]
    _BEKLEYEN[pid] = (t["tur"], preview, filename)
    return gorunum, pid


@callback(
    Output("ia-govde", "children"),
    Output("ia-pid", "data"),
    Output("ia-dosya", "data"),
    Input("ia-upload", "contents"),
    State("ia-upload", "filename"),
    prevent_initial_call=True,
)
def dosya_yuklendi(contents, filename):
    if not contents:
        return no_update, no_update, no_update
    try:
        yol = gecici_kaydet(contents, filename)
    except Exception as e:
        return [html.Div([adimlar(1), ui.banner(f"Dosya kaydedilemedi: {e}", "err", "x-octagon")],
                         className="vm-page")], None, None
    gorunum, pid = _isle(yol, filename)
    return gorunum, pid, {"yol": str(yol), "ad": filename}


@callback(
    Output("ia-govde", "children", allow_duplicate=True),
    Output("ia-pid", "data", allow_duplicate=True),
    Input("ia-yeniden", "n_clicks"),
    State("ia-dosya", "data"),
    State("ia-tur", "value"),
    State({"type": "ia-sutun", "alan": ALL}, "value"),
    State({"type": "ia-sutun", "alan": ALL}, "id"),
    State("ia-pid", "data"),
    prevent_initial_call=True,
)
def yeniden_dene(n, dosya, tur, harfler, idler, eski_pid):
    """Kullanıcının girdiği sütun harfleriyle aynı dosyayı yeniden tespit + önizleme."""
    if not n or not dosya:
        return no_update, no_update
    _BEKLEYEN.pop(eski_pid, None)          # eski önizleme artık geçersiz
    yol, ad = Path(dosya["yol"]), dosya["ad"]
    if not yol.exists():
        return [html.Div([adimlar(1), ui.banner("Geçici dosya bulunamadı, lütfen dosyayı tekrar yükleyin.", "warn"),
                          ui.kart(ui.yukleme_alani("ia-upload", "Dosyayı yükle — ", kucuk=True))],
                         className="vm-page")], None
    override, hatalar, kullanilan = {}, [], {}
    for id_, harf in zip(idler or [], harfler or []):
        try:
            col = tespit.sutun_harfi_coz(harf)
        except ValueError as e:
            hatalar.append(f"{tespit.ALAN_ETIKET.get(id_['alan'], id_['alan'])}: {e}")
            continue
        if col:
            if col in kullanilan:
                hatalar.append(f"'{(harf or '').strip().upper()}' sütunu iki alana birden verilmiş: "
                               f"{tespit.ALAN_ETIKET.get(kullanilan[col], kullanilan[col])} ve "
                               f"{tespit.ALAN_ETIKET.get(id_['alan'], id_['alan'])}")
            kullanilan[col] = id_["alan"]
            override[id_["alan"]] = col
    return _isle(yol, ad, tur=tur, override=override, ek_mesaj=hatalar or None)


@callback(
    Output({"type": "conflict-choice", "tc": ALL, "field": ALL}, "value"),
    Input("ia-hepsi-mevcut", "n_clicks"),
    Input("ia-hepsi-yeni", "n_clicks"),
    State({"type": "conflict-choice", "tc": ALL, "field": ALL}, "value"),
    prevent_initial_call=True,
)
def toplu_karar(_m, _y, degerler):
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return [no_update] * len(degerler)
    secim = "existing" if ctx.triggered_id == "ia-hepsi-mevcut" else "new"
    return [secim] * len(degerler)


@callback(
    Output("ia-govde", "children", allow_duplicate=True),
    Input("ia-onay", "n_clicks"),
    State("ia-pid", "data"),
    State({"type": "conflict-choice", "tc": ALL, "field": ALL}, "value"),
    State({"type": "conflict-choice", "tc": ALL, "field": ALL}, "id"),
    prevent_initial_call=True,
)
def onayla(n, pid, degerler, idler):
    if not n:
        return no_update
    kayit = _BEKLEYEN.pop(pid, None)
    if kayit is None:
        return html.Div([ui.banner("Önizleme süresi doldu, lütfen dosyayı tekrar yükleyin.", "warn"),
                         dcc.Link("Tekrar yükle", href="/ice-aktar", className="vm-btn", refresh=True)],
                        className="vm-page")
    tur, preview, dosya_adi = kayit
    try:
        if tur == "kimlik":
            kararlar = {(i["tc"], i["field"]): v for i, v in zip(idler or [], degerler or [])}
            sonuc = ie.commit_import(preview, field_resolutions=kararlar)
            kutular = [ui.kpi("Yeni personel", f"+{ui.sayi(sonuc['yeni_personel'])}", "green"),
                       ui.kpi("Güncellenen", ui.sayi(sonuc["guncellenen_personel"]), "navy"),
                       ui.kpi("Uyarı", ui.sayi(sonuc["uyari_sayisi"]), "orange" if sonuc["uyari_sayisi"] else None)]
        else:
            sonuc = mie.commit_mesai_import(preview)
            kutular = [ui.kpi("Yeni günlük kayıt", f"+{ui.sayi(sonuc['yeni_kayit'])}", "green"),
                       ui.kpi("Güncellenen", ui.sayi(sonuc["guncellenen_kayit"]), "navy"),
                       ui.kpi("Değişmeyen", ui.sayi(sonuc["degismeyen_kayit"]), "muted"),
                       ui.kpi("Uyarı", ui.sayi(sonuc["uyari_sayisi"]), "orange" if sonuc["uyari_sayisi"] else None)]
    except Exception as e:
        return html.Div([adimlar(4), ui.banner([html.Strong("İçe aktarma sırasında hata oluştu, HİÇBİR ŞEY YAZILMADI. "),
                                                str(e)], "err", "x-octagon"),
                         dcc.Link("Tekrar dene", href="/ice-aktar", className="vm-btn", refresh=True)], className="vm-page")
    return html.Div([
        adimlar(5),
        ui.kart([
            html.Div([html.Span(ui.ikon("check-lg"), className="vm-pipe-ic pi-ok", style={"width": "40px", "height": "40px",
                                                                                         "fontSize": "20px"}),
                      html.Div([html.H2("İçe aktarma tamamlandı", className="vm-h3"),
                                html.Span(f"{dosya_adi} · içe aktarma #{sonuc['batch_id']}", className="vm-sub")],
                               className="col", style={"gap": "2px"})], className="row"),
            html.Div(kutular, className=f"g g-{len(kutular)} gap-12"),
            html.Span([ui.ikon("arrow-counterclockwise"),
                       " Yanlış dosya yüklendiyse Denetim › İçe Aktarma Geçmişi'nden bu içe aktarmayı geri alabilirsiniz. ",
                       f"İçe aktarma öncesinin yedeği: {sonuc['yedek']}." if sonuc.get("yedek")
                       else "Uyarı: içe aktarma öncesi yedek alınamadı (ayrıntı günlük dosyasında)."],
                      className="small-12 muted"),
            html.Div([
                dcc.Link([ui.ikon("house"), "Genel Bakış"], href="/", className="vm-btn"),
                dcc.Link([ui.ikon("shield-check"), "Veri Kalitesi"], href="/kalite", className="vm-btn"),
                dcc.Link([ui.ikon("clock-history"), "İçe aktarma geçmişi"], href="/gecmis", className="vm-btn"),
                html.Span(className="spacer"),
                dcc.Link([ui.ikon("upload"), "Başka dosya yükle"], href="/ice-aktar", className="vm-btn pri", refresh=True),
            ], className="row", style={"flexWrap": "wrap"}),
        ], "pad-l col", style={"gap": "16px"}),
    ], className="vm-page")
