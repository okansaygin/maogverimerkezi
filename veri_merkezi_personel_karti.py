# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - PERSONEL KARTI BİLEŞENLERİ (veri_merkezi_personel_karti.py)
=============================================================================
Personel Kartı'nın (sayfa_personel.sayfa_kart) yerleşim parçaları. Burada callback yok;
yalnız veriyi alıp bileşen üretir. İki kaynak birleştirilir:
  - Veri Merkezi kaydı (personnel + gunluk_puantaj): grup, alt ekip, görev, puantaj
  - İK kaydı (personel_ik_bilgileri, veri_merkezi_ik): kimlik, iletişim, acil durum,
    eğitim, pozisyon, taşeron, konaklama, İSG tarihleri, evrak

Tasarım: üstte şantiye yaka kartı gibi bir kimlik bandı (sahada ilk bakılacak bilgi:
kan grubu ve acil durumda aranacak kişi sağda, kırmızı çerçevede), altında İSG durum
şeridi (tekrar muayene / tekrar eğitim / MYK / evrak - süresi geçen kırmızı, 30 gün
içinde dolan turuncu), sonra sekmeler.
"""

import datetime

from dash import dcc, html

import veri_merkezi_sorgu as sorgu
import veri_merkezi_ui as ui
import veri_merkezi_verimlilik as vm

YAKLASMA_GUN = 30          # İSG tarihine bu kadar gün kala "yaklaşıyor"

# (sütun, etiket, zorunlu mu) - zorunlu evraklardan biri eksikse şerit turuncu olur
EVRAKLAR = [
    ("evrak_sgk_bildirge", "SGK işe giriş bildirgesi", True),
    ("evrak_sozlesme", "Sözleşme", True),
    ("evrak_muayene", "İşe giriş / periyodik muayene raporu", True),
    ("evrak_isg_egitim", "İSG periyodik eğitim evrakı", True),
    ("evrak_diploma", "Diploma ve sertifikalar", False),
    ("evrak_isg", "İSG-Ç evrakları", False),
    ("evrak_diger", "Diğer dökümanlar", False),
    ("yaka_karti", "Yaka kartı", False),
]


# ---------------------------------------------------------------- yardımcılar

def _bos(v):
    return v is None or (isinstance(v, str) and not v.strip())


def _d(iso):
    try:
        return datetime.date.fromisoformat(str(iso)[:10])
    except (TypeError, ValueError):
        return None


def yas(dogum, bugun=None):
    d = _d(dogum)
    if not d:
        return None
    b = bugun or datetime.date.today()
    return b.year - d.year - ((b.month, b.day) < (d.month, d.day))


def sure_metni(bas, bugun=None):
    """'2 yıl 3 ay' / '5 ay' / '12 gün'"""
    d = _d(bas)
    if not d:
        return None
    b = bugun or datetime.date.today()
    if d > b:
        return "henüz başlamadı"
    ay = (b.year - d.year) * 12 + b.month - d.month - (b.day < d.day)
    if ay <= 0:
        return f"{(b - d).days} gün"
    yil, ay = divmod(ay, 12)
    return " ".join(x for x in (f"{yil} yıl" if yil else "", f"{ay} ay" if ay else "") if x)


def tarih_durumu(iso, bugun=None):
    """İSG tarihleri için: (durum_kodu, etiket, ayrıntı). durum: gecti | yakin | tamam | suresiz | yok"""
    if _bos(iso):
        return "yok", "Kayıt yok", ""
    if str(iso).startswith("9999"):
        return "suresiz", "Süresiz", ""
    d = _d(iso)
    if not d:
        return "yok", str(iso), ""
    kalan = (d - (bugun or datetime.date.today())).days
    if kalan < 0:
        return "gecti", "Süresi geçti", f"{-kalan} gün önce doldu"
    if kalan <= YAKLASMA_GUN:
        return "yakin", "Yaklaşıyor", "bugün doluyor" if kalan == 0 else f"{kalan} gün kaldı"
    return "tamam", "Geçerli", f"{kalan} gün kaldı"


def bilgi(etiket, deger, mono=False, ek=None):
    """Tanım listesi satırı: solda etiket, sağda değer (boşsa '—')."""
    bos = deger is None or (isinstance(deger, str) and not deger.strip())
    return html.Div([
        html.Dt(etiket),
        html.Dd([html.Span("—", className="muted") if bos else deger, ek] if ek else
                (html.Span("—", className="muted") if bos else deger), className="mono" if mono and not bos else None),
    ], className="pk-dl-row")


def bilgi_karti(baslik, satirlar, alt=None, cls=""):
    return html.Section([
        html.Div([html.H3(baslik, className="pk-kart-baslik"), html.Span(alt, className="pk-kart-alt") if alt else None],
                 className="pk-kart-bas"),
        html.Dl(satirlar, className="pk-dl"),
    ], className=f"vm-card pad-l pk-kart {cls}".strip())


def evet_hayir(v):
    if v is None:
        return None
    return "Evet" if v else "Hayır"


def _tel(v):
    """05321234567 -> 0532 123 45 67 (okunur); başka biçimleri olduğu gibi bırakır."""
    if _bos(v):
        return None
    rakam = "".join(ch for ch in str(v) if ch.isdigit())
    if len(rakam) == 11 and rakam.startswith("0"):
        return f"{rakam[:4]} {rakam[4:7]} {rakam[7:9]} {rakam[9:]}"
    return str(v)


def _acil_kisi(v):
    """'EŞE AYAR- ANNESİ' -> ('EŞE AYAR', 'annesi')"""
    if _bos(v):
        return None, None
    s = " ".join(str(v).split())
    if "-" in s:
        ad, yakinlik = s.rsplit("-", 1)
        if yakinlik.strip() and ad.strip():
            return ad.strip(), tr_kucuk(yakinlik.strip())
    parcalar = s.split(" ")
    if len(parcalar) > 1 and ui.tr_upper(parcalar[-1]) in YAKINLIKLAR:      # 'ASAD FAYZILLOEV ABİSİ'
        return " ".join(parcalar[:-1]), tr_kucuk(parcalar[-1])
    return s, None


YAKINLIKLAR = {"ANNESİ", "BABASI", "EŞİ", "KARDEŞİ", "ABİSİ", "ABLASI", "OĞLU", "KIZI", "ARKADAŞI", "AMCASI",
               "DAYISI", "TEYZESİ", "HALASI", "DEDESİ", "NİNESİ", "AKRABASI", "KAYINBABASI", "ENİŞTESİ", "YEĞENİ"}


def tr_kucuk(s):
    """Türkçe küçük harf: 'ANNESİ' -> 'annesi', 'KARDEŞI' -> 'kardeşı' değil 'kardeşi' için önce I/İ çevrilir."""
    return (s or "").replace("İ", "i").replace("I", "ı").lower()


# ---------------------------------------------------------------- kimlik bandı (yaka kartı)

def kimlik_bandi(p, ikb, durum_chip, kritik_uyari):
    tc = p["tc"]
    ik_ad = f"{ikb.get('adi') or ''} {ikb.get('soyadi') or ''}".strip() if ikb else ""
    import puantaj_engine as engine
    farkli_ad = ik_ad and engine.normalize_name(ik_ad) != engine.normalize_name(p["ad_soyad"])
    yas_ = yas(ikb.get("dogum_tarihi")) if ikb else None
    kidem = sure_metni(p.get("ise_giris_tarihi") or (ikb or {}).get("sgk_giris_tarihi"))

    cipler = [durum_chip]
    if ikb:
        st = ikb.get("calisan_statusu")
        if st:
            cipler.append(ui.chip(f"İK: {st}", "green" if st == "Etkin" else "red"))
        if ikb.get("calisan_tipi") and ikb["calisan_tipi"] != "Normal":
            cipler.append(ui.chip(ikb["calisan_tipi"], "purple"))
        if ikb.get("gecici_gorevlendirme"):
            cipler.append(ui.chip("Geçici görevlendirme", "orange"))
    if kritik_uyari:
        cipler.append(ui.chip(f"{kritik_uyari} kritik uyarı", "red"))

    kunye = [
        html.Span([html.Span(ui.maske_tc(tc), id="pk-tc", className="mono"),
                   html.Button(ui.ikon("eye"), id="pk-tc-goster", n_clicks=0, className="vm-btn xs",
                               title="TC numarasını göster", **{"aria-label": "TC numarasını göster"})],
                  className="pk-kunye-oge row", style={"gap": "6px"}),
    ]
    if yas_ is not None:
        kunye.append(html.Span([html.Span(str(yas_), className="pk-kunye-deger"), " yaşında"], className="pk-kunye-oge"))
    if kidem:
        kunye.append(html.Span([html.Span(kidem, className="pk-kunye-deger"), " kıdem"], className="pk-kunye-oge",
                               title=f"İşe giriş: {ui.tarih_tr(p.get('ise_giris_tarihi'))}"))
    if ikb and ikb.get("egitim_durumu"):
        kunye.append(html.Span(ikb["egitim_durumu"], className="pk-kunye-oge"))

    sol = html.Div([
        html.Span(ui.bas_harf(p["ad_soyad"]), className="pk-bant-av"),
        html.Span("Personel no", className="pk-bant-lbl"),
        html.Span(ikb.get("personel_no") if ikb and ikb.get("personel_no") else "—", className="pk-bant-no mono"),
    ], className="pk-bant")

    orta = html.Div([
        html.Div(cipler, className="row", style={"gap": "6px", "flexWrap": "wrap"}),
        html.H2(p["ad_soyad"], className="pk-ad"),
        html.Span(["İK kaydındaki resmî ad: ", html.Strong(ik_ad)], className="pk-ad-not") if farkli_ad else None,
        html.Div([html.Span(p.get("grup") or "Grup yok"), ui.ikon("chevron-right", className="muted"),
                  html.Span(p.get("bagli_oldugu_ekip") or "Alt ekip yok"), ui.ikon("chevron-right", className="muted"),
                  html.Strong(p.get("gorevi") or "—")], className="pk-yol"),
        html.Div(kunye, className="pk-kunye"),
        html.Div([
            dcc.Link([ui.ikon("pencil"), "Düzenle"], href=f"/personel-duzenle/{tc}", className="vm-btn sm"),
            html.Button([ui.ikon("file-earmark-excel"), "Kartı Excel'e aktar"], id="pk-excel", className="vm-btn sm",
                        n_clicks=0, title="Kişi bilgileri + bütün puantaj"),
        ], className="row", style={"gap": "8px", "marginTop": "auto"}),
    ], className="pk-orta")

    if ikb:
        acil_ad, yakinlik = _acil_kisi(ikb.get("acil_kisi"))
        kan = ikb.get("kan_grubu")
        sag = html.Div([
            html.Span("Sahada acil durumda", className="pk-acil-baslik"),
            html.Div([
                html.Div([html.Span("Kan grubu", className="pk-acil-lbl"),
                          html.Span(kan or "—", className="pk-kan" + ("" if kan else " bos"))], className="pk-kan-kutu"),
                html.Div([
                    html.Span("Aranacak kişi", className="pk-acil-lbl"),
                    html.Strong(acil_ad or "Kayıt yok"),
                    html.Span(yakinlik, className="muted small-12") if yakinlik else None,
                    html.Span(_tel(ikb.get("acil_telefon")) or "telefon yok", className="mono pk-acil-tel"),
                ], className="col", style={"gap": "2px", "minWidth": 0}),
            ], className="row", style={"gap": "14px", "alignItems": "stretch"}),
            html.Div([html.Span("Kendi telefonu", className="pk-acil-lbl"),
                      html.Span(_tel(ikb.get("telefon")) or "—", className="mono")], className="row pk-acil-kendi"),
        ], className="pk-acil")
    else:
        sag = html.Div([html.Span("Sahada acil durumda", className="pk-acil-baslik"),
                        html.Span("Bu kişi İK listesinde bulunmadığı için kan grubu ve acil durum kişisi bilinmiyor.",
                                  className="small-12 muted")], className="pk-acil bos")
    return html.Div([sol, orta, sag], className="vm-card pk-kimlik")


# ---------------------------------------------------------------- İSG şeridi

def isg_seridi(ikb):
    if not ikb:
        return ui.banner(["Bu kişi İK listesinde bulunamadı; kimlik, iletişim, İSG ve evrak bilgileri yok. ",
                          "İK listesi içe aktarıldığında burada görünür (", html.Code("python veri_merkezi_ik.py"), ")."],
                         "info", "info-circle")
    hucreler = []
    for etiket, alan in (("Tekrar muayene", "tekrar_muayene_tarihi"), ("Tekrar İSG eğitimi", "tekrar_egitim_tarihi")):
        kod, durum, ayr = tarih_durumu(ikb.get(alan))
        hucreler.append(_serit_hucre(etiket, ui.tarih_tr(ikb.get(alan)) if kod != "yok" else "—", durum, ayr, kod))
    if _bos(ikb.get("myk_bitis")):
        hucreler.append(_serit_hucre("MYK belgesi", "—", "Belge kaydı yok", "", "yok"))
    else:
        kod, durum, ayr = tarih_durumu(ikb.get("myk_bitis"))
        deger = "Süresiz" if kod == "suresiz" else ui.tarih_tr(ikb.get("myk_bitis"))
        hucreler.append(_serit_hucre("MYK belgesi", deger, durum, ayr, kod))
    var = sum(1 for a, _e, _z in EVRAKLAR if ikb.get(a))
    eksik_zorunlu = [e for a, e, z in EVRAKLAR if z and not ikb.get(a)]
    hucreler.append(_serit_hucre("Evrak", f"{var} / {len(EVRAKLAR)}",
                                 "Zorunlu evrak eksik" if eksik_zorunlu else "Zorunlu evrak tamam",
                                 ", ".join(eksik_zorunlu), "yakin" if eksik_zorunlu else "tamam"))
    return html.Div(hucreler, className="pk-serit")


def _serit_hucre(etiket, deger, durum, ayrinti, kod):
    return html.Div([
        html.Span(etiket, className="pk-serit-lbl"),
        html.Span(deger, className="pk-serit-deger mono"),
        html.Span([html.Strong(durum), f" · {ayrinti}" if ayrinti else ""], className="pk-serit-durum"),
    ], className=f"pk-serit-hucre d-{kod}", title=ayrinti or durum)


# ---------------------------------------------------------------- özet sekmesinin sağ sütunu

def calisma_ozeti(p, ikb):
    o = sorgu.get_person_calisma_ozeti(p["tc"])
    yil = datetime.date.today().year
    yfm = sorgu.get_person_year_fazla(p["tc"], yil) or 0.0
    sinir = vm.YILLIK_FM_SINIRI
    oran = min(100.0, yfm / sinir * 100) if sinir else 0
    renk = "bg-red" if yfm >= sinir else ("bg-orange" if yfm >= sinir * 0.85 else "bg-navy")
    satirlar = [
        bilgi("İşe giriş", ui.tarih_tr(p.get("ise_giris_tarihi")) if p.get("ise_giris_tarihi") else None, True,
              html.Span(f"  {sure_metni(p['ise_giris_tarihi'])}", className="muted small-12")
              if p.get("ise_giris_tarihi") else None),
        bilgi("İşten çıkış", ui.tarih_tr(p.get("isten_cikis_tarihi")) if p.get("isten_cikis_tarihi") else None, True),
        bilgi("İlk / son çalışma", f"{ui.tarih_tr(o['ilk_calisma'])} – {ui.tarih_tr(o['son_calisma'])}"
              if o["ilk_calisma"] else None, True),
        bilgi("Çalışılan gün", ui.sayi(o["calisilan_gun"] or 0), True),
        bilgi("Toplam saat", f"{ui.saat(o['normal'])} normal + {ui.saat(o['fazla'])} fazla", True),
    ]
    return html.Section([
        html.Div(html.H3("Çalışma özeti", className="pk-kart-baslik"), className="pk-kart-bas"),
        html.Dl(satirlar, className="pk-dl"),
        html.Div([
            html.Div([html.Span(f"{yil} fazla mesaisi", className="small-12"), html.Span(className="spacer"),
                      html.Span([html.Strong(ui.saat(yfm)), f" / {vm.etiket_yillik()} sa"], className="mono small-12")],
                     className="row"),
            ui.bar([(oran, renk)], "thin"),
            html.Span(f"Yıllık sınıra {ui.saat(max(0.0, sinir - yfm))} saat kaldı" if yfm < sinir
                      else f"Yıllık sınır {ui.saat(yfm - sinir)} saat aşıldı", className="small-12 " +
                      ("c-red" if yfm >= sinir else "muted")),
        ], className="col", style={"gap": "6px", "marginTop": "12px"}),
    ], className="vm-card pad-l pk-kart")


# ---------------------------------------------------------------- sekmeler

def _ik_yok():
    return ui.bos("Bu kişi İK listesinde bulunamadı. İK listesi içe aktarıldığında bu bilgiler burada görünür.")


def sekme_kisisel(p, ikb):
    if not ikb:
        return _ik_yok()
    acil_ad, yakinlik = _acil_kisi(ikb.get("acil_kisi"))
    kimlik = bilgi_karti("Kimlik", [
        bilgi("Adı", ikb.get("adi")), bilgi("Soyadı", ikb.get("soyadi")),
        bilgi("Doğum tarihi", ui.tarih_tr(ikb.get("dogum_tarihi")) if ikb.get("dogum_tarihi") else None, True,
              html.Span(f"  {yas(ikb['dogum_tarihi'])} yaşında", className="muted small-12") if ikb.get("dogum_tarihi") else None),
        bilgi("Doğum yeri", ikb.get("dogum_yeri")), bilgi("Cinsiyet", ikb.get("cinsiyet")),
        bilgi("Medeni durum", ikb.get("medeni_durum")),
        bilgi("Baba adı", ikb.get("baba_adi")), bilgi("Anne adı", ikb.get("anne_adi")),
    ])
    iletisim = bilgi_karti("İletişim", [
        bilgi("Telefon", _tel(ikb.get("telefon")), True), bilgi("E-posta", ikb.get("eposta")),
        bilgi("Şehir", ikb.get("adres_sehir")), bilgi("Adres", ikb.get("adres")),
    ])
    acil = bilgi_karti("Acil durum", [
        bilgi("Aranacak kişi", acil_ad), bilgi("Yakınlığı", yakinlik),
        bilgi("Telefonu", _tel(ikb.get("acil_telefon")), True), bilgi("Kan grubu", ikb.get("kan_grubu")),
    ], cls="pk-kart-acil")
    egitim = bilgi_karti("Eğitim", [
        bilgi("Öğrenim durumu", ikb.get("egitim_durumu")), bilgi("Bölüm / dal", ikb.get("egitim_dali")),
    ])
    return html.Div([html.Div([kimlik, egitim], className="col", style={"gap": "16px"}),
                     html.Div([acil, iletisim], className="col", style={"gap": "16px"})], className="pk-iki")


def sekme_is(p, ikb):
    import puantaj_engine as engine
    vm_kart = bilgi_karti("Veri Merkezi kaydı", [
        bilgi("Grup", p.get("grup")), bilgi("Alt ekip", p.get("bagli_oldugu_ekip")), bilgi("Görevi", p.get("gorevi")),
        bilgi("İşe giriş", ui.tarih_tr(p.get("ise_giris_tarihi")) if p.get("ise_giris_tarihi") else None, True),
        bilgi("İşten çıkış", ui.tarih_tr(p.get("isten_cikis_tarihi")) if p.get("isten_cikis_tarihi") else None, True),
        bilgi("Notlar", p.get("notlar")),
        bilgi("Kaynak", html.Span(p.get("kaynak") or "—", className="small-12 muted")),
    ], alt="puantaj ve raporlarda kullanılan")
    if not ikb:
        return html.Div([vm_kart, _ik_yok()], className="pk-iki")

    gorev_farkli = (ikb.get("gorev") and p.get("gorevi")
                    and engine.normalize_name(ikb["gorev"]) != engine.normalize_name(p["gorevi"]))
    ik_kart = bilgi_karti("İK kaydı", [
        bilgi("Durum", ikb.get("calisan_statusu")),
        bilgi("Pozisyon", ikb.get("pozisyon"), ek=html.Span(f"  {ikb['pozisyon_kodu']}", className="muted small-12 mono")
              if ikb.get("pozisyon_kodu") else None),
        bilgi("Görev", ikb.get("gorev"), ek=ui.chip("Veri Merkezi'nden farklı", "orange") if gorev_farkli else None),
        bilgi("Bölüm", ikb.get("bolum")), bilgi("Çalışan grubu", ikb.get("calisan_grubu")),
        bilgi("Çalışan tipi", ikb.get("calisan_tipi")),
        bilgi("SGK meslek", ikb.get("sgk_meslek"), ek=html.Span(f"  {ikb['sgk_meslek_kodu']}", className="muted small-12 mono")
              if ikb.get("sgk_meslek_kodu") else None),
        bilgi("Lokasyon", ikb.get("lokasyon")), bilgi("İK ekip kodu", ikb.get("ik_ekip_kodu"), True),
        bilgi("Görev yeri", ikb.get("gorev_yeri")),
    ], alt="İK sistemi")
    sozlesme = bilgi_karti("Giriş ve sözleşme", [
        bilgi("SGK giriş", ui.tarih_tr(ikb.get("sgk_giris_tarihi")) if ikb.get("sgk_giris_tarihi") else None, True),
        bilgi("İşe giriş (İK)", ui.tarih_tr(ikb.get("ise_giris_tarihi")) if ikb.get("ise_giris_tarihi") else None, True),
        bilgi("Yeniden giriş", ui.tarih_tr(ikb.get("yeniden_giris_tarihi")) if ikb.get("yeniden_giris_tarihi") else None, True),
        bilgi("Taşeron firma", ikb.get("taseron_firma")),
        bilgi("Normal kadro", evet_hayir(ikb.get("normal_kadro"))),
        bilgi("Destek işçilik", evet_hayir(ikb.get("destek_iscilik"))),
        bilgi("Geçici görevlendirme", evet_hayir(ikb.get("gecici_gorevlendirme")),
              ek=html.Span(f"  bitiş {ui.tarih_tr(ikb['gecici_gorev_bitis'])}", className="muted small-12")
              if ikb.get("gecici_gorevlendirme") and ikb.get("gecici_gorev_bitis") else None),
        bilgi("Onay belge no", ikb.get("onay_belge_no"), True),
        bilgi("Form tarihi", ui.tarih_tr(ikb.get("form_olusturma_tarihi")) if ikb.get("form_olusturma_tarihi") else None, True),
    ])
    yasam = bilgi_karti("Konaklama ve ulaşım", [
        bilgi("Konaklama", ikb.get("konaklama")), bilgi("Lojman / ek bilgi", ikb.get("lojman_bilgisi"), True),
        bilgi("Ulaşım", ikb.get("ulasim")),
    ])
    return html.Div([
        html.Div([vm_kart, sozlesme, yasam], className="col", style={"gap": "16px"}),
        html.Div([ik_kart, _ik_gecmisi(ikb)], className="col", style={"gap": "16px"}),
    ], className="pk-iki")


def _ik_gecmisi(ikb):
    kayitlar = [dict(ikb, _guncel=True)] + [dict(k, _guncel=False) for k in (ikb.get("onceki_kayitlar") or [])]
    kayitlar.sort(key=lambda k: (k.get("ise_giris_tarihi") or "", k.get("form_olusturma_tarihi") or ""), reverse=True)
    satirlar = [[html.Td(ui.tarih_tr(k.get("ise_giris_tarihi")), className="mono small-12"),
                 html.Td(ui.chip(k.get("calisan_statusu") or "—", "green" if k.get("calisan_statusu") == "Etkin" else "gray")),
                 html.Td(k.get("gorev") or "—", className="small-12"),
                 html.Td(k.get("lokasyon") or "—", className="small-12"),
                 html.Td("güncel" if k["_guncel"] else "", className="small-12 muted")] for k in kayitlar]
    return html.Section([
        html.Div([html.H3("İK kayıt geçmişi", className="pk-kart-baslik"),
                  html.Span(f"{len(kayitlar)} kayıt", className="pk-kart-alt")], className="pk-kart-bas"),
        ui.tablo(["İşe giriş", "Durum", "Görev", "Lokasyon", ""], satirlar, "compact"),
        html.Span(f"Kaynak: {ikb.get('kaynak_dosya') or '—'}, aktarma {ui.tarih_tr(ikb.get('aktarma_zamani'), saatli=True)}",
                  className="small-12 muted"),
    ], className="vm-card pad-l pk-kart col", style={"gap": "10px"})


def sekme_isg(ikb):
    if not ikb:
        return _ik_yok()
    satirlar = []
    for etiket, alan in (("Tekrar muayene", "tekrar_muayene_tarihi"), ("Tekrar İSG eğitimi", "tekrar_egitim_tarihi"),
                         ("MYK belgesi başlangıç", "myk_baslangic"), ("MYK belgesi bitiş", "myk_bitis")):
        v = ikb.get(alan)
        if alan == "myk_baslangic":
            satirlar.append([etiket, html.Td(ui.tarih_tr(v) if v else "—", className="mono"), "", ""])
            continue
        kod, durum, ayr = tarih_durumu(v)
        ton = {"gecti": "red", "yakin": "orange", "tamam": "green", "suresiz": "green"}.get(kod, "gray")
        satirlar.append([etiket, html.Td("Süresiz" if kod == "suresiz" else (ui.tarih_tr(v) if v else "—"), className="mono"),
                         html.Td(ui.chip(durum, ton)), html.Td(ayr, className="small-12 muted")])
    tarihler = html.Section([
        html.Div(html.H3("Tarihler", className="pk-kart-baslik"), className="pk-kart-bas"),
        ui.tablo(["", "Tarih", "Durum", ""], satirlar, "compact"),
        html.Span(f"Süresi {YAKLASMA_GUN} gün içinde dolanlar 'yaklaşıyor' olarak işaretlenir.", className="small-12 muted"),
    ], className="vm-card pad-l pk-kart col", style={"gap": "10px"})
    evrak = html.Section([
        html.Div([html.H3("Evraklar", className="pk-kart-baslik"),
                  html.Span(f"{sum(1 for a, _e, _z in EVRAKLAR if ikb.get(a))} / {len(EVRAKLAR)} tamam",
                            className="pk-kart-alt")], className="pk-kart-bas"),
        html.Ul([html.Li([ui.ikon("check-circle-fill" if ikb.get(a) else ("x-circle-fill" if z else "dash-circle"),
                                  className="pk-evrak-ikon " + ("var" if ikb.get(a) else ("eksik" if z else "yok"))),
                          html.Span(e), html.Span("zorunlu", className="small-12 muted") if z else None],
                         className="pk-evrak") for a, e, z in EVRAKLAR], className="pk-evrak-liste"),
        html.Span("Evrak bilgisi İK listesindeki işaretlerden gelir; belgenin kendisi burada tutulmaz.",
                  className="small-12 muted"),
    ], className="vm-card pad-l pk-kart col", style={"gap": "10px"})
    return html.Div([tarihler, evrak], className="pk-iki")


# ---------------------------------------------------------------- Excel için düz liste

def excel_satirlari(p, ikb):
    """Kartın Excel'e aktarılan 'Kişi bilgileri' sayfası için (başlık, değer) listesi."""
    s = [("Ad soyad", p["ad_soyad"]), ("TC Kimlik No", p["tc"]), ("Grup", p.get("grup")),
         ("Alt ekip", p.get("bagli_oldugu_ekip")), ("Görevi", p.get("gorevi")),
         ("İşe giriş", p.get("ise_giris_tarihi")), ("İşten çıkış", p.get("isten_cikis_tarihi")), ("Notlar", p.get("notlar"))]
    if ikb:
        import veri_merkezi_ik as ik
        s.append(("", None))
        s.append(("İK BİLGİLERİ", None))
        for baslik, sutun, tur in ik.ESLEME:
            v = ikb.get(sutun)
            if tur in ("b", "x"):
                v = None if v is None else ("Evet" if v else "Hayır")
            s.append((baslik, v))
    return s
