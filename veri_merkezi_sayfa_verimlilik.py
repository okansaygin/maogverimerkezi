# -*- coding: utf-8 -*-
"""VERİ MERKEZİ - SAYFA: Verimlilik Analizi (eski adı: Fazla Mesai Raporu).

Analiz veri_merkezi_verimlilik.analiz() ile, Excel veri_merkezi_verimlilik_excel
ile üretilir; ekran ve Excel aynı sözlüğü kullandığı için sayılar birebir aynıdır.

Grafikler HTML/CSS ile çizilir (sütun, halka, ısı haritası): plotly gerektirmez,
yazdırılınca da aynı görünür. Her grafik öğesinin üzerine gelince (title) değeri görünür.
"""

import datetime
import logging
import time

from dash import dcc, html, dash_table, Input, Output, State, no_update, callback, ctx
import dash_bootstrap_components as dbc

import veri_merkezi_ekip_listesi as el
import veri_merkezi_ui as ui
import veri_merkezi_verimlilik as vm
import veri_merkezi_verimlilik_excel as vx
from veri_merkezi_sayfa_puantaj import _hiyerarsi, _ekipler, _gorevler

ONEK = "va"
gunluk = logging.getLogger("veri_merkezi")
TON_RENK = {"iyi": "#2F6B3A", "uyari": "#B8410C", "kritik": "#B42318", "bilgi": "#1F4E78"}
TON_ZEMIN = {"iyi": "#E3EFE5", "uyari": "#FDF3EA", "kritik": "#F9E1DE", "bilgi": "#E6ECF3"}
TON_IKON = {"iyi": "check-circle-fill", "uyari": "exclamation-circle-fill", "kritik": "exclamation-octagon-fill",
            "bilgi": "info-circle-fill"}
BANT_TON = {"A": "green", "B": "navy", "C": "orange", "D": "red"}


# ============================================================ yardımcılar

def _s(x, ondalik=0):
    return ui.sayi(x, ondalik)


def _pc(x, ondalik=1):
    return f"%{ui.sayi(x, ondalik)}"


def _varsayilan_aralik():
    """Varsayılan: içinde bulunulan ayın başı - bugün (ayın ilk 3 günündeysek geçen ay)."""
    bugun = datetime.date.today()
    if bugun.day <= 3:
        son = bugun.replace(day=1) - datetime.timedelta(days=1)
        return son.replace(day=1).isoformat(), son.isoformat()
    return bugun.replace(day=1).isoformat(), bugun.isoformat()


def _hizli_aralik(anahtar, bugun=None):
    bugun = bugun or datetime.date.today()
    if anahtar == "bu-ay":
        return bugun.replace(day=1), bugun
    if anahtar == "gecen-ay":
        son = bugun.replace(day=1) - datetime.timedelta(days=1)
        return son.replace(day=1), son
    if anahtar == "son-30":
        return bugun - datetime.timedelta(days=29), bugun
    if anahtar == "son-3-ay":
        y, m = bugun.year, bugun.month - 2
        if m <= 0:
            y, m = y - 1, m + 12
        return datetime.date(y, m, 1), bugun
    return datetime.date(bugun.year, 1, 1), bugun          # bu-yil


def _delta(simdi, once, iyi_yon=1, tip="puan", ondalik=1):
    """Önceki döneme göre değişim çipi. tip='puan' fark (yüzde puanı), 'oran' göreli %.
    iyi_yon: +1 artış iyi, -1 azalış iyi, 0 nötr."""
    if once is None:
        return None
    if tip == "oran":
        if not once:
            return None
        fark = (simdi - once) / once * 100
        metin = f"%{ui.sayi(abs(fark), ondalik)}"
    else:
        fark = simdi - once
        metin = f"{ui.sayi(abs(fark), ondalik)} puan"
    if abs(fark) < 0.05:
        return html.Span("değişmedi", className="va-delta notr")
    iyi = (fark > 0) == (iyi_yon > 0) if iyi_yon else None
    cls = "iyi" if iyi else ("kotu" if iyi is not None else "notr")
    return html.Span([ui.ikon("arrow-up-short" if fark > 0 else "arrow-down-short"), metin], className=f"va-delta {cls}")


def _kpi(etiket, deger, alt=None, delta=None, renk=None, ipucu=None):
    return html.Div([
        html.Span(etiket, className="vm-lbl"),
        html.Span(deger, className=f"vm-kpi-val {('c-' + renk) if renk else ''}"),
        html.Div([delta, html.Span(alt, className="vm-kpi-note")], className="row", style={"gap": "6px",
                                                                                              "flexWrap": "wrap"}),
    ], className="vm-card vm-kpi", title=ipucu)


def _kart(baslik, govde, alt=None, sag=None, cls="pad-l", **kw):
    return html.Section([html.Div([ui.h2(baslik, alt), sag], className="vm-card-head"), govde],
                        className=f"vm-card {cls} col", style={"gap": "12px", **kw.pop("style", {})}, **kw)


def _ton_rengi(devam):
    if devam is None:
        return "va-h0"
    if devam >= 97:
        return "va-h5"
    if devam >= 93:
        return "va-h4"
    if devam >= 88:
        return "va-h3"
    if devam >= 80:
        return "va-h2"
    return "va-h1"


def _tl(x):
    return f"{ui.sayi(x)} TL"


# ============================================================ sayfa

def sayfa():
    bas, bit = _varsayilan_aralik()
    h = _hiyerarsi()
    hizli = [("bu-ay", "Bu ay"), ("gecen-ay", "Geçen ay"), ("son-30", "Son 30 gün"), ("son-3-ay", "Son 3 ay"),
             ("bu-yil", "Bu yıl")]
    filtre = ui.kart([
        html.Div([
            ui.alan("Dönem", html.Div(dcc.DatePickerRange(id="va-tarih", start_date=bas, end_date=bit,
                                                          display_format="DD.MM.YYYY", first_day_of_week=1,
                                                          minimum_nights=0), className="vm-dd")),
            html.Div([ui.lbl("Hızlı seçim"),
                      html.Div([html.Button(e, id=f"va-hizli-{k}", n_clicks=0, className="vm-btn sm") for k, e in hizli],
                               className="row", style={"gap": "6px", "flexWrap": "wrap"})],
                     className="col", style={"gap": "6px"}),
            ui.alan("Grafik kırılımı", ui.segment("va-periyot", [("gun", "Günlük"), ("hafta", "Haftalık"),
                                                                 ("ay", "Aylık")], "hafta")),
        ], className="row", style={"gap": "18px", "alignItems": "end", "flexWrap": "wrap"}),
        html.Div([
            ui.alan("Grup", ui.dropdown("va-grup", sorted(h, key=ui.tr_upper), [], coklu=True)),
            ui.alan("Alt ekip", ui.dropdown("va-ekip", _ekipler(h, None), [], coklu=True)),
            ui.alan("Görevi", ui.dropdown("va-gorev", _gorevler(h, None, None), [], coklu=True)),
            ui.alan("Günlük yevmiye ücreti (TL)",
                    dcc.Input(id="va-ucret", type="number", min=0, step=50, placeholder="isteğe bağlı",
                              className="vm-input", debounce=True, persistence=True, persistence_type="local")),
            html.Button([ui.ikon("file-earmark-excel"), "Excel raporu"], id="va-excel", className="vm-btn pri",
                        n_clicks=0, style={"alignSelf": "end", "height": "38px"}),
        ], className="g", style={"gridTemplateColumns": "repeat(3, minmax(0,1fr)) 170px auto", "gap": "12px",
                                 "alignItems": "end"}),
    ], "pad col", style={"gap": "14px"})

    sekmeler = ui.sekmeler("va", [
        ("genel", "Genel görünüm", html.Div(id="va-t-genel", className="col", style={"gap": "16px"})),
        ("ekip", "Ekip ve görev", html.Div(id="va-t-ekip", className="col", style={"gap": "16px"})),
        ("personel", "Personel karnesi", html.Div(id="va-t-personel", className="col", style={"gap": "16px"})),
        ("risk", "Uyum ve risk", html.Div(id="va-t-risk", className="col", style={"gap": "16px"})),
        ("yontem", "Yöntem", _yontem()),
    ], aktif="genel")

    return ui.sayfa("Verimlilik Analizi", "Raporlar / Verimlilik Analizi", [
        filtre,
        html.Div(id="va-mesaj"),
        dcc.Loading(html.Div([html.Div(id="va-ozet", className="col", style={"gap": "16px"}), sekmeler],
                             className="col", style={"gap": "16px"}), type="default", color="#1F4E78"),
    ])


# ============================================================ özet (üst bölüm)

def _puan_karti(d):
    g = d["genel"]
    harf, bant = vm.puan_bandi(g["puan"])
    o = d["onceki"]
    esikler = html.Div([html.Span(style={"left": f"{x}%"}, className="va-tick") for x in (70, 85, 95)],
                       className="va-ticks")
    return html.Section([
        html.Span("Verimlilik puanı", className="vm-lbl"),
        html.Div([html.Span(_s(g["puan"], 1), className="va-puan"),
                  html.Span(harf, className=f"va-harf ch-{BANT_TON[harf]}"),
                  html.Span(bant, className="va-bant")], className="row", style={"gap": "10px", "alignItems": "baseline"}),
        html.Div([ui.bar([(g["puan"], f"bg-{'navy' if harf in 'AB' else ('orange' if harf == 'C' else 'red')}")], "mid"),
                  esikler], style={"position": "relative"}),
        html.Div([html.Span("D < 70", className="small-12 muted"), html.Span("C 70–85", className="small-12 muted"),
                  html.Span("B 85–95", className="small-12 muted"), html.Span("A ≥ 95", className="small-12 muted")],
                 className="row", style={"justifyContent": "space-between"}),
        html.Div([_delta(g["puan"], o["puan"] if o else None, 1),
                  html.Span("önceki döneme göre" if o else "önceki dönemde veri yok", className="small-12 muted")],
                 className="row", style={"gap": "6px"}),
        html.Div([ui.bilgi_satiri("Devam oranı × 0,6", _pc(g["devam"]), True),
                  ui.bilgi_satiri("Kapasite kullanımı × 0,4", _pc(min(g["kapasite"], 100)), True)],
                 className="col", style={"gap": "4px", "borderTop": "1px solid #EFEBE4", "paddingTop": "10px"}),
        _puan_seyri(d["periyotlar"]),
    ], className="vm-card pad col va-puankart", style={"gap": "10px"})


def _puan_seyri(periyotlar):
    """Periyot bazında puan: 60-100 aralığında mini sütunlar."""
    pp = [p for p in periyotlar if p["planlanan"]][-12:]
    if len(pp) < 2:
        return None
    sutun = []
    for p in pp:
        harf, _b = vm.puan_bandi(p["puan"])
        h = max(4.0, min(100.0, (p["puan"] - 60) / 40 * 100))
        sutun.append(html.Div(html.Div(className=f"bg-{'navy' if harf in 'AB' else ('orange' if harf == 'C' else 'red')}",
                                       style={"height": f"{h:.0f}%", "borderRadius": "3px 3px 0 0"}),
                              className="va-ps", title=f"{p['etiket']}: {ui.sayi(p['puan'], 1)} ({harf})"))
    return html.Div([ui.lbl("Periyot bazında puan"), html.Div(sutun, className="va-ps-row"),
                     html.Div([html.Span(pp[0]["etiket"]), html.Span(pp[-1]["etiket"])], className="row small-12 muted",
                              style={"justifyContent": "space-between"})],
                    className="col", style={"gap": "4px", "marginTop": "auto"})


def _ozet_listesi(d):
    satirlar = [html.Div([
        html.Span(ui.ikon(TON_IKON[o["ton"]]), className="va-oz-ic",
                  style={"color": TON_RENK[o["ton"]], "background": TON_ZEMIN[o["ton"]]}),
        html.Div([html.Span(o["baslik"], className="va-oz-b"), html.Span(o["metin"], className="va-oz-m")],
                 className="col", style={"gap": "2px"}),
    ], className="va-oz") for o in d["ozet"]]
    o1, o2 = d["onceki_aralik"]
    return html.Section([
        html.Div([ui.h2("Yönetici özeti", vm.donem_adi(d)),
                  ui.chip(vm.kapsam_adi(d), "navy-l"),
                  html.Span(f"karşılaştırma: {o1:%d.%m} – {o2:%d.%m.%Y}", className="small-12 muted")],
                 className="vm-card-head", style={"flexWrap": "wrap"}),
        html.Div(satirlar, className="va-oz-grid"),
    ], className="vm-card pad-l col", style={"gap": "12px"})


def _kpi_seridi(d, ucret):
    g, o = d["genel"], d["onceki"]
    og = (lambda k: o[k] if o else None)
    yev_alt = f"≈ {_tl(g['yevmiye'] * ucret)}" if ucret else f"{_s(g['calisilan'])} adam-gün çalışıldı"
    fm_yev_alt = (f"≈ {_tl(g['fm_yevmiye'] * ucret)} · ek {_tl(g['fm_prim'] * ucret)}" if ucret
                  else f"katsayı farkı {_s(g['fm_prim'], 1)} yevmiye")
    return html.Section([
        _kpi("Günlük ortalama çalışan", _s(g["ort_calisan"]), f"{_s(g['kisi'])} farklı personel",
             _delta(g["ort_calisan"], og("ort_calisan"), 0, "oran"),
             ipucu="Çalışılan adam-gün ÷ verisi olan gün sayısı"),
        _kpi("Devam oranı", _pc(g["devam"]), f"{_s(g['calisilan'])} / {_s(g['planlanan'])} iş günü",
             _delta(g["devam"], og("devam"), 1), "navy", ipucu="Çalışılan gün ÷ planlanan iş günü (hafta tatili, "
                                                               "resmî tatil hariç)"),
        _kpi("Kapasite kullanımı", _pc(g["kapasite"]), f"kayıp {_s(g['kayip_saat'])} sa",
             _delta(g["kapasite"], og("kapasite"), 1), ipucu="Normal mesai saati ÷ (planlanan gün × 9 sa)"),
        _kpi("Kişi başı günlük çalışma", f"{ui.sayi(g['ort_saat'], 2)} sa", "normal + fazla, çalışılan gün başına",
             _delta(g["ort_saat"], og("ort_saat"), 0, "oran")),
        _kpi("Fazla mesai", f"{_s(g['fazla'])} sa", f"pazar ×2,5: {_s(g['pazar_katli'])} sa · kişi başı "
                                                    f"{ui.sayi(g['kisi_basi_fm'], 1)} sa",
             _delta(g["fazla"], og("fazla"), -1, "oran"), "orange"),
        _kpi("FM oranı", _pc(g["fm_oran"]), "toplam çalışma içinde", _delta(g["fm_oran"], og("fm_oran"), -1)),
        _kpi("Kayıp iş günü", _s(g["kayip"]), f"{_pc(g['kayip_oran'])} · devamsızlık {_s(g['gun']['devamsiz'])} gün",
             _delta(g["kayip_oran"], og("kayip_oran"), -1), "red" if g["kayip_oran"] >= 6 else None,
             ipucu="Devamsızlık + rapor + ücretsiz izin"),
        _kpi("Yevmiye karşılığı", _s(g["yevmiye"], 1), yev_alt, _delta(g["yevmiye"], og("yevmiye"), 0, "oran"),
             ipucu=vm.katsayi_metni()),
        _kpi("FM yevmiye karşılığı", _s(g["fm_yevmiye"], 1), fm_yev_alt,
             _delta(g["fm_yevmiye"], og("fm_yevmiye"), -1, "oran"), "orange"),
        _kpi("Uyum riski olan personel", _s(d["risk"]["riskli_kisi"]),
             f"{_s(len(d['risk']['uzun_gunler']))} kez {vm.etiket_gunluk()} sa aşımı",
             None, "red" if d["risk"]["riskli_kisi"] else "green", ipucu="Ayrıntı: Uyum ve risk sekmesi"),
    ], className="g g-5 gap-12")


def ust_bolum(d, ucret):
    return [html.Div([_puan_karti(d), _ozet_listesi(d)], className="g",
                     style={"gridTemplateColumns": "300px minmax(0,1fr)"}),
            _kpi_seridi(d, ucret)]


# ============================================================ grafikler (HTML/CSS)

def _sutun_grafigi(periyotlar):
    mx = max([p["toplam"] for p in periyotlar] + [1])
    n = len(periyotlar)
    az = n <= 16                        # sütun üstü FM değerleri
    dv_goster = n <= 12                 # alt satırda devam oranı
    adim = max(1, -(-n // 10))          # en fazla ~10 x etiketi
    sutunlar, etiketler, devam = [], [], []
    for i, p in enumerate(periyotlar):
        h = p["toplam"] / mx * 86          # üstte değer etiketi için pay
        fz = (p["fazla"] / p["toplam"] * 100) if p["toplam"] else 0
        ipucu = (f"{p['etiket']}\nNormal {ui.saat(p['normal'])} sa · Fazla {ui.saat(p['fazla'])} sa "
                 f"(%{ui.sayi(p['fm_oran'], 1)})\nDevam %{ui.sayi(p['devam'], 1)} · günlük ort. "
                 f"{ui.sayi(p['ort_calisan'])} kişi")
        sutunlar.append(html.Div([
            html.Span(_s(p["fazla"]) if az and p["fazla"] else "", className="va-col-v"),
            html.Div([html.Div(className="f", style={"height": f"{fz:.1f}%"}), html.Div(className="n")],
                     className="va-col-bar", style={"height": f"{max(h, 1):.1f}%"}),
        ], className="va-col", title=ipucu))
        etiketler.append(html.Span(p["etiket"] if i % adim == 0 else "", title=p["etiket"]))
        devam.append(html.Span(_pc(p["devam"]), className=f"va-dv {_ton_rengi(p['devam'] if p['planlanan'] else None)}"))
    return html.Div([
        html.Div(sutunlar, className=f"va-cols {'sik' if n > 24 else ''}"),
        html.Div(etiketler, className=f"va-xlab {'sik' if n > 24 else ''}"),
        html.Div(devam, className="va-xlab dv") if dv_goster else None,
    ], className="col", style={"gap": "4px"})


def _halka(g):
    toplam = sum(g["gun"].values()) or 1
    bas, parcalar, lejant = 0.0, [], []
    for k, e, renk, _kod in vm.KATEGORILER:
        n = g["gun"][k]
        if not n:
            continue
        pay = n / toplam * 100
        parcalar.append(f"#{renk} {bas:.2f}% {bas + pay:.2f}%")
        bas += pay
        lejant.append(html.Div([html.I(style={"background": f"#{renk}"}), html.Span(e, className="grow"),
                                html.Span(_s(n), className="mono"), html.Span(_pc(pay), className="mono muted",
                                                                               style={"width": "52px", "textAlign": "right"})],
                               className="va-lej"))
    return html.Div([
        html.Div(html.Div([html.Span(_pc(g["devam"]), className="va-halka-v"), html.Span("devam", className="small-12 muted")],
                          className="va-halka-ic"),
                 className="va-halka", style={"background": f"conic-gradient({', '.join(parcalar)})"}),
        html.Div(lejant, className="col", style={"gap": "6px", "flexGrow": 1}),
    ], className="row", style={"gap": "22px", "alignItems": "center"})


def _gun_yigini(gunluk):
    """Günlük iş gücü: her gün için kategori (çalıştı / izin / rapor / devamsız / ücretsiz) yığını."""
    katlar = ["calisti", "ucretli_izin", "ucretsiz_izin", "rapor", "devamsiz"]
    mx = max([sum(g["gun"][k] for k in katlar) for g in gunluk] + [1])
    sutun, etiket = [], []
    adim = max(1, len(gunluk) // 12)
    for i, g in enumerate(gunluk):
        toplam = sum(g["gun"][k] for k in katlar)
        parca = [html.Div(className=f"va-k-{k}", style={"height": f"{g['gun'][k] / toplam * 100:.1f}%"} if toplam else {})
                 for k in reversed(katlar) if g["gun"][k]]
        ipucu = f"{g['tarih']:%d.%m.%Y} {g['gun_adi']}\n" + "\n".join(
            f"{vm.KAT_ETIKET[k]}: {g['gun'][k]}" for k in katlar if g["gun"][k])
        sutun.append(html.Div(html.Div(parca, className="va-col-bar yigin", style={"height": f"{toplam / mx * 100:.1f}%"}),
                              className=f"va-col {'pz' if g['tarih'].weekday() == 6 else ''}", title=ipucu))
        etiket.append(html.Span(f"{g['tarih']:%d.%m}" if i % adim == 0 else ""))
    lej = [ui.lejant(f"#{vm.KAT_RENK[k]}", vm.KAT_ETIKET[k]) for k in katlar]
    sik = "sik" if len(gunluk) > 45 else ""
    return html.Div([html.Div(lej, className="row", style={"gap": "12px", "flexWrap": "wrap"}),
                     html.Div(sutun, className=f"va-cols ince {sik}"), html.Div(etiket, className=f"va-xlab ince {sik}")],
                    className="col", style={"gap": "6px"})


def _haftanin_gunleri(hg):
    mx = max([h["ort_calisan"] for h in hg] + [1])
    satirlar = [html.Div([
        html.Span(h["ad"], style={"fontWeight": 600, "fontSize": "13px"}),
        ui.bar([(h["ort_calisan"] / mx * 100, "bg-orange" if h["wd"] == 6 else "bg-navy")], "mid"),
        html.Span(_s(h["ort_calisan"]), className="mono r"),
        html.Span(f"{ui.sayi(h['ort_saat'], 1)} sa", className="mono r muted"),
        html.Span(_pc(h["fm_oran"]), className="mono r c-orange"),
    ], className="va-hg") for h in hg]
    return html.Div([html.Div([html.Span("Gün"), html.Span("Ortalama çalışan kişi"), html.Span("kişi", className="r"),
                               html.Span("sa / kişi", className="r"), html.Span("FM oranı", className="r")],
                              className="va-hg h"), *satirlar], className="col", style={"gap": "8px"})


def _hareket_karti(h):
    def liste(kayitlar, alan, bos):
        if not kayitlar:
            return html.Span(bos, className="small-12 muted")
        return html.Div([html.Div([ui.kisi_link(k["tc"], k["ad_soyad"]),
                                   html.Span(k["gorev"], className="small-12 muted grow"),
                                   html.Span(f"{k[alan]:%d.%m}", className="mono small-12")],
                                  className="row", style={"gap": "8px"}) for k in kayitlar[:6]]
                        + ([html.Span(f"+{len(kayitlar) - 6} kişi daha (Excel'de tam liste)", className="small-12 muted")]
                           if len(kayitlar) > 6 else []), className="col", style={"gap": "6px"})
    return _kart("Personel hareketi", html.Div([
        html.Div([
            html.Div([ui.lbl("Dönem başı"), html.Span(_s(h["bas_mevcut"]), className="vm-kpi-val")], className="col"),
            html.Span(ui.ikon("arrow-right"), className="muted", style={"fontSize": "20px"}),
            html.Div([ui.lbl("Dönem sonu"), html.Span(_s(h["son_mevcut"]), className="vm-kpi-val")], className="col"),
            html.Div([ui.lbl("Giren"), html.Span(f"+{_s(len(h['girenler']))}", className="vm-kpi-val c-green")], className="col"),
            html.Div([ui.lbl("Çıkan"), html.Span(f"−{_s(len(h['cikanlar']))}", className="vm-kpi-val c-red")], className="col"),
            html.Div([ui.lbl("Devir oranı"), html.Span(_pc(h["devir_orani"]), className="vm-kpi-val")], className="col"),
        ], className="row", style={"gap": "22px", "flexWrap": "wrap"}),
        html.Div([html.Div([ui.lbl("İşe başlayanlar"), liste(h["girenler"], "giris", "Dönemde işe başlayan yok.")],
                           className="col", style={"gap": "8px"}),
                  html.Div([ui.lbl("Ayrılanlar"), liste(h["cikanlar"], "cikis", "Dönemde ayrılan yok.")],
                           className="col", style={"gap": "8px"})], className="g g-2"),
    ], className="col", style={"gap": "16px"}), alt="kimlik kayıtlarındaki giriş / çıkış tarihlerinden")


def sekme_genel(d):
    g = d["genel"]
    periyot_adi = {"gun": "Günlük", "hafta": "Haftalık", "ay": "Aylık"}[d["periyot"]]
    trend = _kart(f"{periyot_adi} çalışma ve fazla mesai", _sutun_grafigi(d["periyotlar"]),
                  sag=html.Div([ui.lejant("#1F4E78", "Normal"), ui.lejant("#E8762C", "Fazla mesai"),
                                html.Span("alt satır: devam oranı", className="small-12 muted")],
                               className="row", style={"gap": "12px"}), cls="pad-l span-8")
    dagilim = _kart("Gün dağılımı", _halka(g), alt="adam-gün", cls="pad-l span-4")
    isgucu = _kart("Günlük iş gücü", _gun_yigini(d["gunluk"]), alt="kişi sayısı · turuncu zemin = pazar",
                   cls="pad-l span-7")
    haftalik = _kart("Haftanın günleri", _haftanin_gunleri(d["haftanin_gunleri"]), alt="gün başına ortalama",
                     cls="pad-l span-5")
    eksik = None
    if d["eksik_gun"]:
        eksik = ui.banner([html.Strong(f"{len(d['eksik_gun'])} günün SAP mesai verisi eksik: "),
                           ", ".join(f"{t:%d.%m}" for t in d["eksik_gun"][:15]) + (" …" if len(d["eksik_gun"]) > 15 else ""),
                           ". Bu günler hesaplara girmedi."], "warn")
    return [eksik, html.Div([trend, dagilim], className="g-12"), html.Div([isgucu, haftalik], className="g-12"),
            _hareket_karti(d["hareket"])]


# ============================================================ ekip ve görev

def _karsilastirma_tablosu(satirlar, baslik_ad, onceki_var):
    mx_fm = max([k["fm_oran"] for k in satirlar] + [1])
    govde = []
    for k in sorted(satirlar, key=lambda x: x["sira"]):
        dv = _delta(k["devam"], k["onceki_devam"], 1) if onceki_var else None
        fm = _delta(k["fazla"], k["onceki_fazla"], -1, "oran") if onceki_var else None
        govde.append(html.Tr([
            html.Td(k["sira"], className="mono c muted"),
            html.Td(k["ad"], style={"fontWeight": 600, "fontSize": "13px"}),
            html.Td(_s(k["kisi"]), className="mono r"),
            html.Td(_s(k["calisilan"]), className="mono r"),
            html.Td(html.Div([html.Span(_pc(k["devam"]), className=f"va-pill {_ton_rengi(k['devam'])}"), dv],
                             className="row", style={"gap": "4px", "justifyContent": "flex-end"}), className="r"),
            html.Td(_pc(k["kapasite"]), className="mono r"),
            html.Td(ui.sayi(k["ort_saat"], 2), className="mono r"),
            html.Td(_s(k["fazla"]), className="mono r"),
            html.Td(html.Div([ui.bar([(k["fm_oran"] / mx_fm * 100, "bg-orange")], "thin"),
                              html.Span(_pc(k["fm_oran"]), className="mono", style={"width": "48px", "textAlign": "right"})],
                             className="row", style={"gap": "6px"}), style={"minWidth": "120px"}),
            html.Td(fm, className="r"),
            html.Td(_s(k["kayip"]), className="mono r", title=f"devamsızlık {k['gun']['devamsiz']} · rapor "
                                                                f"{k['gun']['rapor']} · ücretsiz {k['gun']['ucretsiz_izin']}"),
            html.Td(ui.sayi(k["yevmiye"], 1), className="mono r"),
            html.Td(html.Span([html.B(ui.sayi(k["puan"], 1)), f" {k['harf']}"], className=f"vm-chip ch-{BANT_TON[k['harf']]}"),
                    className="r"),
        ]))
    return ui.tablo([("#", "c"), baslik_ad, ("Kişi", "r"), ("Adam-gün", "r"), ("Devam", "r"), ("Kapasite", "r"),
                     ("Sa/gün", "r"), ("FM sa", "r"), "FM oranı", ("FM Δ", "r"), ("Kayıp gün", "r"), ("Yevmiye", "r"),
                     ("Puan", "r")], govde, cls="va-tbl")


def _isi_haritasi(d):
    tarihler = d["isi_tarihler"]
    if not tarihler:
        return ui.bos("Veri yok.")
    kol = f"210px repeat({len(tarihler)}, minmax(9px, 1fr))"
    adim = 1 if len(tarihler) <= 31 else 7
    bas = [html.Span("")] + [html.Span(f"{t:%d}" if (i % adim == 0 or t.day == 1) else "",
                                       className=f"va-isi-t {'pz' if t.weekday() == 6 else ''}")
                             for i, t in enumerate(tarihler)]
    satirlar = [html.Div(bas, className="va-isi", style={"gridTemplateColumns": kol})]
    for s in d["isi"][:30]:
        hucre = [html.Span(s["ad"], className="va-isi-ad", title=s["ad"])]
        for t, h in zip(tarihler, s["hucreler"]):
            if h is None:
                hucre.append(html.Span(className="va-isi-c va-hx", title=f"{t:%d.%m} · kayıt yok"))
            else:
                ipucu = (f"{s['ad']} · {t:%d.%m.%Y}\nÇalışan {h['calisan']} / planlanan {h['planlanan']}"
                         + (f" (%{ui.sayi(h['devam'], 1)})" if h["devam"] is not None else " (tatil)")
                         + f"\n{ui.saat(h['saat'])} sa")
                hucre.append(html.Span(className=f"va-isi-c {_ton_rengi(h['devam'])}", title=ipucu))
        satirlar.append(html.Div(hucre, className="va-isi", style={"gridTemplateColumns": kol}))
    lej = [html.Span([html.I(className=c), t], className="vm-legend va-lg") for c, t in
           [("va-h5", "≥ %97"), ("va-h4", "%93–97"), ("va-h3", "%88–93"), ("va-h2", "%80–88"), ("va-h1", "< %80"),
            ("va-h0", "tatil / plan yok"), ("va-hx", "kayıt yok")]]
    return html.Div([html.Div(lej, className="row", style={"gap": "12px", "flexWrap": "wrap"}),
                     html.Div(satirlar, className="col", style={"gap": "3px", "overflowX": "auto"})],
                    className="col", style={"gap": "10px"})


def sekme_ekip(d):
    etiket = "Alt ekip" if d["kirilim"] == "ekip" else "Grup"
    onceki_var = d["onceki"] is not None
    ipucu = "Sıra verimlilik puanına göredir. FM Δ: önceki döneme göre fazla mesai değişimi."
    return [
        html.Section([html.Div([ui.h2(f"{etiket} karşılaştırması", f"{len(d['kirilimlar'])} {etiket.lower()}"),
                                html.Span(ipucu, className="small-12 muted")], className="vm-card-head pad"),
                      html.Div(_karsilastirma_tablosu(d["kirilimlar"], etiket, onceki_var), style={"overflowX": "auto"})],
                     className="vm-card clip"),
        _kart("Günlük devam haritası · alt ekip bazında", _isi_haritasi(d),
              alt="her kare bir gün; üzerine gelince ayrıntı"),
        html.Section([html.Div([ui.h2("Görev karşılaştırması", f"{len(d['gorev_tablosu'])} görev")],
                               className="vm-card-head pad"),
                      html.Div(_karsilastirma_tablosu(d["gorev_tablosu"], "Görevi", onceki_var), style={"overflowX": "auto"})],
                     className="vm-card clip"),
    ]


# ============================================================ personel karnesi

def _mini_liste(baslik, kisiler, deger, alt, ikon, renk):
    return _kart(baslik, html.Div([
        html.Div([html.Span(str(i), className="va-sira"), ui.kisi_link(k["tc"], k["ad_soyad"]),
                  html.Span(alt(k), className="small-12 muted grow", style={"overflow": "hidden", "textOverflow": "ellipsis",
                                                                           "whiteSpace": "nowrap"}),
                  html.Span(deger(k), className=f"mono c-{renk}", style={"fontWeight": 600, "whiteSpace": "nowrap"})],
                 className="row", style={"gap": "8px"}) for i, k in enumerate(kisiler, start=1)]
        or [html.Span("Kayıt yok.", className="small-12 muted")], className="col", style={"gap": "8px"}),
        sag=ui.ikon(ikon, className=f"c-{renk}"), cls="pad")


def sekme_personel(d):
    ks = d["kisiler"]
    yeterli = [k for k in ks if k["planlanan"] >= 5]
    en_iyi = sorted(yeterli, key=lambda k: (-k["puan"], -k["calisilan"]))[:5]
    en_dusuk = sorted(yeterli, key=lambda k: (k["devam"], -k["kayip"]))[:5]
    en_fm = sorted(ks, key=lambda k: -k["fazla"])[:5]
    en_kayip = sorted([k for k in ks if k["kayip"]], key=lambda k: (-k["kayip"], -k["gun"]["devamsiz"]))[:5]
    listeler = html.Div([
        _mini_liste("En yüksek puan", en_iyi, lambda k: ui.sayi(k["puan"], 1), lambda k: k["gorev"], "trophy", "green"),
        _mini_liste("En düşük devam", en_dusuk, lambda k: _pc(k["devam"]), lambda k: f"{k['kayip']} gün kayıp",
                    "person-dash", "red"),
        _mini_liste("En çok fazla mesai", en_fm, lambda k: f"{ui.saat(k['fazla'])} sa",
                    lambda k: f"yıllık {ui.saat(k['yillik_fm'])} sa", "clock-history", "orange"),
        _mini_liste("En çok kayıp gün", en_kayip, lambda k: f"{k['kayip']} gün",
                    lambda k: f"D {k['gun']['devamsiz']} · R {k['gun']['rapor']} · Ü {k['gun']['ucretsiz_izin']}",
                    "calendar-x", "red"),
    ], className="g g-2")

    kolonlar = [("ad_soyad", "Ad Soyad", "text"), ("gorev", "Görevi", "text"), ("ekip", "Alt ekip", "text"),
                ("calisilan", "Çalışılan gün", "numeric"), ("devam", "Devam %", "numeric"),
                ("kapasite", "Kapasite %", "numeric"), ("ort_saat", "Sa/gün", "numeric"),
                ("fazla", "FM sa", "numeric"), ("pazar_katli", "Pazar ×2,5 sa", "numeric"), ("fm_oran", "FM %", "numeric"),
                ("devamsiz", "Devamsız", "numeric"), ("rapor", "Rapor", "numeric"), ("kayip", "Kayıp gün", "numeric"),
                ("yillik_fm", f"{d['yil']} FM sa", "numeric"), ("puan", "Puan", "numeric"), ("harf", "Not", "text"),
                ("risk", "Uyarılar", "text")]
    veri = [{"ad_soyad": k["ad_soyad"], "gorev": k["gorev"], "ekip": k["ekip"], "calisilan": k["calisilan"],
             "devam": round(k["devam"], 1), "kapasite": round(k["kapasite"], 1), "ort_saat": round(k["ort_saat"], 2),
             "fazla": round(k["fazla"], 1), "pazar_katli": round(k["pazar_katli"], 1), "fm_oran": round(k["fm_oran"], 1),
             "devamsiz": k["gun"]["devamsiz"], "rapor": k["gun"]["rapor"], "kayip": k["kayip"],
             "yillik_fm": round(k["yillik_fm"], 1), "puan": round(k["puan"], 1), "harf": k["harf"],
             "risk": " · ".join(k["riskler"])} for k in ks]
    tablo = html.Section([
        html.Div([ui.h2("Personel karnesi", f"{len(ks)} kişi"),
                  html.Span("Sütun başlığına tıklayarak sıralayın, altındaki kutuya yazarak süzün (ör. > 10).",
                            className="small-12 muted")], className="vm-card-head pad"),
        html.Div(dash_table.DataTable(
            id="va-karne", columns=[{"name": a, "id": k, "type": t} for k, a, t in kolonlar], data=veri,
            page_size=20, sort_action="native", filter_action="native", style_table={"overflowX": "auto"},
            style_cell={"whiteSpace": "normal", "height": "auto", "fontSize": "13px"},
            style_cell_conditional=[{"if": {"column_id": "ad_soyad"}, "minWidth": "160px", "fontWeight": 600},
                                    {"if": {"column_id": "risk"}, "minWidth": "200px"}],
            style_data_conditional=[
                {"if": {"filter_query": "{devam} < 88", "column_id": "devam"}, "color": "#B42318", "fontWeight": 600},
                {"if": {"filter_query": "{devam} >= 97", "column_id": "devam"}, "color": "#2F6B3A", "fontWeight": 600},
                {"if": {"filter_query": "{fazla} > 0", "column_id": "fazla"}, "color": "#9A3A0B", "fontWeight": 600},
                {"if": {"filter_query": "{devamsiz} >= 3", "column_id": "devamsiz"}, "backgroundColor": "#F9E1DE"},
                {"if": {"filter_query": '{harf} = "A"', "column_id": "harf"}, "backgroundColor": "#E3EFE5"},
                {"if": {"filter_query": '{harf} = "C"', "column_id": "harf"}, "backgroundColor": "#FDF3EA"},
                {"if": {"filter_query": '{harf} = "D"', "column_id": "harf"}, "backgroundColor": "#F9E1DE"},
                {"if": {"filter_query": "!({risk} is blank)", "column_id": "risk"}, "color": "#B42318"},
                {"if": {"state": "active"}, "backgroundColor": "#EEF3F8", "border": "none"},
            ],
        ), className="vm-dt"),
    ], className="vm-card clip")
    return [listeler, tablo]


# ============================================================ uyum ve risk

def _risk_tablosu(baslik, aciklama, basliklar, satirlar, toplam, ton="red"):
    govde = satirlar[:15]
    return html.Section([
        html.Div([html.Div([ui.h2(baslik), html.Span(aciklama, className="small-12 muted")], className="col",
                           style={"gap": "2px", "flexGrow": 1}),
                  ui.chip(f"{toplam}", ton if toplam else "green")], className="vm-card-head pad"),
        ui.tablo(basliklar, govde) if govde else html.Div(ui.banner("Bu dönemde kayıt yok.", "ok", "check2-circle"),
                                                          style={"padding": "0 18px 16px"}),
        html.Div(f"+{toplam - 15} kayıt daha — tam liste Excel raporunda", className="small-12 muted",
                 style={"padding": "8px 18px 12px"}) if toplam > 15 else None,
    ], className="vm-card clip")


def sekme_risk(d):
    r = d["risk"]
    ozet = html.Section([
        _kpi(f"Yıllık {vm.etiket_yillik()} sa sınırı", f"{len(r['yillik_asan'])} / {len(r['yillik_yakin'])}", "aşan / yaklaşan (≥ %85)",
             renk="red" if r["yillik_asan"] else ("orange" if r["yillik_yakin"] else "green")),
        _kpi("Dönem FM eşiği üstü", _s(len(r["esik_ustu"])),
             f"eşik {ui.sayi(d['esik_donem'], 1)} sa (aylık {ui.sayi(d['esik_aylik'])} sa)",
             renk="orange" if r["esik_ustu"] else "green"),
        _kpi(f"Günlük {vm.etiket_gunluk()} saat aşımı", _s(len(r["uzun_gunler"])),
             f"{_s(len({u['tc'] for u in r['uzun_gunler']}))} kişide", renk="red" if r["uzun_gunler"] else "green"),
        _kpi(f"{vm.KESINTISIZ_GUN_ESIGI}+ gün kesintisiz", _s(len(r["kesintisiz"])), "hafta tatili kullanılmadan",
             renk="orange" if r["kesintisiz"] else "green"),
        _kpi(f"{vm.DEVAMSIZLIK_ESIGI}+ gün devamsızlık", _s(len(r["devamsiz"])), "mazeretsiz",
             renk="red" if r["devamsiz"] else "green"),
    ], className="g g-5 gap-12")

    def kisi_td(k):
        return html.Td(ui.kisi_link(k["tc"], k["ad_soyad"]))

    yillik = r["yillik_asan"] + r["yillik_yakin"]
    t1 = _risk_tablosu("Yıllık fazla mesai sınırı", "4857 s. İş Kanunu md. 41: yılda en fazla 270 saat"
                       + vm.esik_notu(vm.YILLIK_FM_SINIRI, 270),
                       ["Personel", "Görevi", ("Dönem FM", "r"), (f"{d['yil']} toplam", "r"), ("Kalan", "r"), "Doluluk"],
                       [[kisi_td(k), html.Td(k["gorev"], style={"fontSize": "13px"}),
                         html.Td(f"{ui.saat(k['fazla'])} sa", className="mono r"),
                         html.Td(f"{ui.saat(k['yillik_fm'])} sa", className="mono r", style={"fontWeight": 600}),
                         html.Td(f"{ui.saat(k['yillik_kalan'])} sa", className="mono r"),
                         html.Td(ui.bar([(min(100, k["yillik_fm"] / vm.YILLIK_FM_SINIRI * 100),
                                          "bg-red" if k["yillik_fm"] >= vm.YILLIK_FM_SINIRI else "bg-orange")], "thin"),
                                 style={"width": "120px"})] for k in yillik], len(yillik))
    t2 = _risk_tablosu(f"Günlük {vm.etiket_gunluk()} saati aşan çalışma",
                       f"md. 63: günlük çalışma 11 saati aşamaz{vm.esik_notu(vm.GUNLUK_AZAMI_SAAT, 11)} · "
                       f"{len(r['uzun_gunler'])} kişi-gün, kişi bazında",
                       ["Personel", "Alt ekip", ("Gün", "r"), ("En yüksek", "r"), "Tarihler"],
                       [[kisi_td(u), html.Td(u["ekip"], style={"fontSize": "13px"}),
                         html.Td(f"{u['gun']} gün", className="mono r c-red", style={"fontWeight": 600}),
                         html.Td(f"{ui.saat(u['en_yuksek'])} sa", className="mono r"),
                         html.Td(", ".join(f"{t:%d.%m}" for t in u["tarihler"][:8]) + (" …" if len(u["tarihler"]) > 8 else ""),
                                 className="mono small-12")]
                        for u in r["uzun_kisi"]], len(r["uzun_kisi"]))
    t3 = _risk_tablosu(f"{vm.KESINTISIZ_GUN_ESIGI} gün ve üzeri kesintisiz çalışma",
                       "md. 46: 7 günlük süre içinde en az 24 saat kesintisiz dinlenme (hafta tatili)",
                       ["Personel", "Alt ekip", ("En uzun seri", "r"), "Seriler"],
                       [[kisi_td(k), html.Td(k["ekip"], style={"fontSize": "13px"}),
                         html.Td(f"{k['en_uzun_seri']} gün", className="mono r c-orange", style={"fontWeight": 600}),
                         html.Td(", ".join(f"{a:%d.%m}–{b:%d.%m}" for a, b, _n in k["seriler"][:3]), className="mono small-12")]
                        for k in r["kesintisiz"]], len(r["kesintisiz"]), "orange")
    t4 = _risk_tablosu("Dönem fazla mesai eşiğini aşanlar",
                       f"Ayarlar'daki aylık {ui.sayi(d['esik_aylik'])} sa eşiği, dönem uzunluğuna oranlanır",
                       ["Personel", "Görevi", ("Dönem FM", "r"), ("Pazar ×2,5", "r"), ("Eşik", "r")],
                       [[kisi_td(k), html.Td(k["gorev"], style={"fontSize": "13px"}),
                         html.Td(f"{ui.saat(k['fazla'])} sa", className="mono r c-orange", style={"fontWeight": 600}),
                         html.Td(f"{ui.saat(k['pazar_katli'])} sa", className="mono r"),
                         html.Td(f"{ui.sayi(d['esik_donem'], 1)} sa", className="mono r muted")]
                        for k in r["esik_ustu"]], len(r["esik_ustu"]), "orange")
    t5 = _risk_tablosu(f"{vm.DEVAMSIZLIK_ESIGI} gün ve üzeri devamsızlık", "mazeretsiz devamsızlık (DV)",
                       ["Personel", "Alt ekip", ("Devamsız", "r"), ("Rapor", "r"), ("Devam", "r")],
                       [[kisi_td(k), html.Td(k["ekip"], style={"fontSize": "13px"}),
                         html.Td(f"{k['gun']['devamsiz']} gün", className="mono r c-red", style={"fontWeight": 600}),
                         html.Td(f"{k['gun']['rapor']} gün", className="mono r"),
                         html.Td(_pc(k["devam"]), className="mono r")] for k in r["devamsiz"]], len(r["devamsiz"]))
    return [ozet, html.Div([t1, t4], className="g g-2"), t2, html.Div([t3, t5], className="g g-2"),
            html.Div("Bu liste bilgi amaçlıdır; hukuki değerlendirme için İK / hukuk birimine danışın. Hesaplar SAP "
                     "mesai kayıtlarına dayanır; eksik ya da hatalı gün kaydı sonucu etkiler.", className="vm-note")]


# ============================================================ yöntem

def _yontem():
    tanimlar = [
        ("Planlanan iş günü", "Kişinin kaydı olan günler − hafta tatili − resmî tatil − boş kayıt."),
        ("Devam oranı", "Çalışılan gün ÷ planlanan iş günü. Saat girilmiş her gün 'çalıştı' sayılır."),
        ("Kayıp iş günü", "Devamsızlık + sağlık raporu + ücretsiz izin. Yıllık/ücretli izin kayıp sayılmaz. "
                          "Pazar çalışılan haftadaki tek ücretsiz izin hafta tatiline döner; eksik günlü haftada "
                          "çalışılmayan pazar ücretsiz izin sayılır."),
        ("Kapasite kullanımı", f"Normal mesai saati ÷ (planlanan iş günü × {vm._sa(vm.STANDART_GUN)} saat). Devamsızlık ve eksik günler "
                               "kapasiteyi düşürür; fazla mesai bu orana katılmaz."),
        ("Kişi başı günlük çalışma", "(Normal + fazla mesai) ÷ çalışılan gün."),
        ("FM oranı", "Fazla mesai ÷ (normal + fazla mesai)."),
        ("Yevmiye karşılığı", f"{vm.katsayi_metni()}. Ücret girilirse TL'ye çevrilir."),
        ("Pazar ve hafta tatili", vm.pazar_kurali_metni()),
        ("FM yevmiye karşılığı", "Yevmiyenin, hafta içi normal saatlerin (×1) dışında kalan kısmı: fazla mesai "
                                 "ve hak edilmiş pazar saatleri."),
        ("FM katsayı farkı", f"Aynı saatler ×1 ödenseydi ödenmeyecek kısım: {vm.katsayi_farki_metni()} yevmiye."),
        ("Verimlilik puanı", "Devam oranı × 0,6 + kapasite kullanımı (en fazla 100) × 0,4. Not: A ≥ 95, B 85–95, "
                             "C 70–85, D < 70."),
        ("Önceki dönem", "Seçilen dönemle aynı uzunlukta, hemen önceki dönem."),
        ("Günlük ortalama çalışan", "Çalışılan adam-gün ÷ verisi olan gün sayısı."),
        ("Uyum kontrolleri", f"Yıllık {vm.etiket_yillik()} sa FM (md. 41), günlük {vm.etiket_gunluk()} sa (md. 63), "
                             f"{vm.KESINTISIZ_GUN_ESIGI}+ gün kesintisiz çalışma (md. 46), Ayarlar'daki aylık FM eşiği. "
                             f"Eşikler Sistem › Ayarlar'dan değiştirilebilir."),
    ]
    return html.Div([
        _kart("Göstergeler nasıl hesaplanır?", html.Div([
            html.Div([html.Span(a, className="va-ym-a"), html.Span(b, className="va-ym-b")], className="va-ym")
            for a, b in tanimlar], className="col", style={"gap": "0"}),
            alt="ekran ve Excel aynı formülleri kullanır"),
        html.Div("Kaynak: Veri Merkezi kalıcı deposu (SAP günlük mesai + personel kimlik kayıtları). Çalışma durumu "
                 "metni (ÇALIŞTI, HAFTA TATİLİ, YILLIK İZİN, SAĞLIK RAPORU, DEVAMSIZLIK, ÜCRETSİZ İZİN …) otomatik "
                 "sınıflandırılır.", className="vm-note"),
    ], className="col", style={"gap": "16px"})


# ============================================================ callback'ler

@callback(
    Output("va-ekip", "options"), Output("va-ekip", "value"),
    Output("va-gorev", "options"), Output("va-gorev", "value"),
    Input("va-grup", "value"), Input("va-ekip", "value"),
    State("va-gorev", "value"),
    prevent_initial_call=True,
)
def va_zincir(gruplar, ekipler, gorevler):
    h = _hiyerarsi()
    eks = _ekipler(h, gruplar)
    ekipler = [e for e in (ekipler or []) if e in eks]
    gors = _gorevler(h, gruplar, ekipler)
    return ([{"label": e, "value": e} for e in eks], ekipler,
            [{"label": g, "value": g} for g in gors], [g for g in (gorevler or []) if g in gors])


_HIZLI = ["bu-ay", "gecen-ay", "son-30", "son-3-ay", "bu-yil"]


@callback(
    Output("va-tarih", "start_date"), Output("va-tarih", "end_date"),
    Input("va-hizli-bu-ay", "n_clicks"), Input("va-hizli-gecen-ay", "n_clicks"), Input("va-hizli-son-30", "n_clicks"),
    Input("va-hizli-son-3-ay", "n_clicks"), Input("va-hizli-bu-yil", "n_clicks"),
    prevent_initial_call=True,
)
def va_hizli(n1, n2, n3, n4, n5):
    if not ctx.triggered_id or not (ctx.triggered and ctx.triggered[0]["value"]):
        return no_update, no_update
    a, b = _hizli_aralik(ctx.triggered_id.replace("va-hizli-", ""))
    return a.isoformat(), b.isoformat()


def _veri(bas, bit, periyot, grup, ekip, gorev):
    return vm.analiz(bas[:10], bit[:10], periyot or "hafta", grup or None, ekip or None, gorev or None)


def _ucret(u):
    try:
        u = float(u)
        return u if u > 0 else None
    except (TypeError, ValueError):
        return None


@callback(
    Output("va-ozet", "children"), Output("va-t-genel", "children"), Output("va-t-ekip", "children"),
    Output("va-t-personel", "children"), Output("va-t-risk", "children"),
    Input("va-tarih", "start_date"), Input("va-tarih", "end_date"), Input("va-periyot", "value"),
    Input("va-grup", "value"), Input("va-ekip", "value"), Input("va-gorev", "value"), Input("va-ucret", "value"),
)
def va_goster(bas, bit, periyot, grup, ekip, gorev, ucret):
    if not bas or not bit:
        return ui.banner("Dönem seçin.", "warn"), None, None, None, None
    t0 = time.perf_counter()
    try:
        d = _veri(bas, bit, periyot, grup, ekip, gorev)
        if not d["kayit_sayisi"]:
            return ui.bos("Seçilen dönem ve kapsamda puantaj kaydı yok."), None, None, None, None
        sonuc = (ust_bolum(d, _ucret(ucret)), sekme_genel(d), sekme_ekip(d), sekme_personel(d), sekme_risk(d))
    except Exception as e:
        gunluk.exception("verimlilik analizi hata verdi (%s – %s)", bas, bit)
        return ui.banner(f"Analiz yapılamadı: {type(e).__name__}: {e}", "err", "x-octagon"), None, None, None, None
    gunluk.info("verimlilik %s – %s, %s kayıt (%.2f sn)", bas[:10], bit[:10], d["kayit_sayisi"],
                time.perf_counter() - t0)
    return sonuc


@callback(
    Output("va-mesaj", "children"),
    Input("va-excel", "n_clicks"),
    State("va-tarih", "start_date"), State("va-tarih", "end_date"), State("va-periyot", "value"),
    State("va-grup", "value"), State("va-ekip", "value"), State("va-gorev", "value"), State("va-ucret", "value"),
    prevent_initial_call=True,
)
def va_excel(n, bas, bit, periyot, grup, ekip, gorev, ucret):
    if not n or not bas or not bit:
        return no_update
    try:
        d = _veri(bas, bit, periyot, grup, ekip, gorev)
        yol = vx.excel_yaz(d, ucret=_ucret(ucret))
    except PermissionError:
        return ui.banner("Dosya yazılamadı: aynı adlı Excel şu an açık olabilir. Kapatıp tekrar deneyin.", "err")
    except Exception as e:
        return ui.banner(f"Rapor oluşturulamadı: {e}", "err")
    el.open_path(yol)
    return ui.banner(["Verimlilik raporu oluşturuldu: ", html.Code(str(yol))], "ok", "check2-circle")
