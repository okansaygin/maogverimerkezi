# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - ÇEVRİMDIŞI ARAYÜZ DOSYALARI (veri_merkezi_varliklar.py)
=========================================================================
Arayüz üç dış stil dosyasına dayanır: Bootstrap (düzen), Bootstrap ikonları ve
yazı tipleri (Archivo, IBM Plex). Bunlar normalde internetten (CDN) yüklenir;
internet yoksa ikonlar kaybolur, Bootstrap gelmezse düzen bozulabilir.

Bu modül üçünü BİR KEZ indirip uygulama klasöründeki yerel_varliklar/ altına koyar.
Dosyalar varsa uygulama onları kendi içinden sunar (internet gerekmez); yoksa eskisi
gibi CDN'den yükler. Yani indirme yapılmasa da hiçbir şey bozulmaz.

İndirme:
  - Uygulamadan: Sistem > Ayarlar > "Çevrimdışı çalışma" > Dosyaları indir
  - Komut satırından:  python veri_merkezi_varliklar.py
İndirmeden sonra uygulamayı yeniden başlatın.

Güvenlik: her kaynak önce geçici bir klasöre indirilir; yalnızca TAMAMI başarıyla
inince eskisinin yerine konur. Yarım indirme mevcut dosyaları bozmaz.
"""

import logging
import re
import shutil
import urllib.parse
import urllib.request
from pathlib import Path

import dash_bootstrap_components as dbc

gunluk = logging.getLogger("veri_merkezi")

SCRIPT_DIR = Path(__file__).resolve().parent
VARLIK_KLASORU = SCRIPT_DIR / "yerel_varliklar"
URL_ONEKI = "/yerel"

FONTLAR = ("https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800"
           "&family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap")

# (klasör adı, CDN adresi) - sıra önemlidir: sayfaya bu sırayla eklenir.
KAYNAKLAR = [
    ("bootstrap", dbc.themes.BOOTSTRAP),
    ("ikonlar", dbc.icons.BOOTSTRAP),
    ("yazitipleri", FONTLAR),
]
ETIKETLER = {"bootstrap": "Bootstrap (düzen)", "ikonlar": "Bootstrap ikonları", "yazitipleri": "Yazı tipleri"}

# Google Fonts, tarayıcıya göre farklı biçim döner; modern bir tarayıcı gibi görünüp woff2 isteriz.
TARAYICI = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36")
_URL_DESENI = re.compile(r"""url\(\s*(['"]?)([^'")]+?)\1\s*\)""")


def _varsayilan_getir(url, zaman_asimi):
    istek = urllib.request.Request(url, headers={"User-Agent": TARAYICI})
    with urllib.request.urlopen(istek, timeout=zaman_asimi) as yanit:
        return yanit.read()


def _guvenli_ad(url, kullanilan):
    ad = Path(urllib.parse.urlparse(url).path).name or "dosya"
    ad = re.sub(r"[^A-Za-z0-9._-]", "_", ad)[:80]
    aday, i = ad, 2
    while aday in kullanilan:
        aday = f"{i}_{ad}"
        i += 1
    kullanilan.add(aday)
    return aday


def _kaynak_indir(ad, url, gecici, getir, zaman_asimi):
    """Bir stil dosyasını ve içinde başvurulan bütün dosyaları (yazı tipleri) indirir,
    başvuruları yerel yollara çevirip gecici/stil.css olarak yazar. Döner: indirilen dosya sayısı."""
    css = getir(url, zaman_asimi).decode("utf-8")
    (gecici / "dosya").mkdir(parents=True, exist_ok=True)
    eslesme, kullanilan = {}, set()

    def degistir(m):
        ref = m.group(2).strip()
        if ref.startswith("data:") or ref.startswith("#"):
            return m.group(0)
        tam = urllib.parse.urljoin(url, ref)
        if tam not in eslesme:
            dosya_adi = _guvenli_ad(tam, kullanilan)
            (gecici / "dosya" / dosya_adi).write_bytes(getir(tam, zaman_asimi))
            eslesme[tam] = dosya_adi
        return f'url("dosya/{eslesme[tam]}")'

    yeni = _URL_DESENI.sub(degistir, css)
    (gecici / "stil.css").write_text(yeni, encoding="utf-8")
    return len(eslesme)


def indir(hedef=None, getir=None, kaynaklar=None, zaman_asimi=30):
    """Bütün kaynakları indirir. Döner: [{"ad", "etiket", "ok", "hata", "dosya"}].
    Başarısız olan kaynağın eski (varsa) yerel kopyasına dokunulmaz."""
    hedef = Path(hedef) if hedef else VARLIK_KLASORU
    getir = getir or _varsayilan_getir
    kaynaklar = kaynaklar or KAYNAKLAR
    hedef.mkdir(parents=True, exist_ok=True)
    rapor = []
    for ad, url in kaynaklar:
        gecici = hedef / f".{ad}.indiriliyor"
        eski = hedef / f".{ad}.eski"
        shutil.rmtree(gecici, ignore_errors=True)
        try:
            n = _kaynak_indir(ad, url, gecici, getir, zaman_asimi)
            shutil.rmtree(eski, ignore_errors=True)
            if (hedef / ad).exists():
                (hedef / ad).rename(eski)
            gecici.rename(hedef / ad)
            shutil.rmtree(eski, ignore_errors=True)
            rapor.append({"ad": ad, "etiket": ETIKETLER.get(ad, ad), "ok": True, "hata": None, "dosya": n + 1})
            gunluk.info("çevrimdışı dosya indirildi: %s (%d dosya)", ad, n + 1)
        except Exception as e:   # ağ, izin, kod çözme... - hiçbiri uygulamayı durdurmamalı
            shutil.rmtree(gecici, ignore_errors=True)
            if eski.exists() and not (hedef / ad).exists():
                eski.rename(hedef / ad)
            rapor.append({"ad": ad, "etiket": ETIKETLER.get(ad, ad), "ok": False, "hata": str(e), "dosya": 0})
            gunluk.warning("çevrimdışı dosya indirilemedi: %s (%s)", ad, e)
    return rapor


def _yerel_var(hedef, ad):
    return (hedef / ad / "stil.css").is_file()


def durum(hedef=None, kaynaklar=None):
    hedef = Path(hedef) if hedef else VARLIK_KLASORU
    sonuc = []
    for ad, url in (kaynaklar or KAYNAKLAR):
        yol = hedef / ad
        yerel = _yerel_var(hedef, ad)
        boyut = sum(p.stat().st_size for p in yol.rglob("*") if p.is_file()) if yerel else 0
        sonuc.append({"ad": ad, "etiket": ETIKETLER.get(ad, ad), "yerel": yerel, "boyut": boyut, "cdn": url})
    return sonuc


def stil_listesi(hedef=None, kaynaklar=None):
    """Dash'in external_stylesheets listesi: yerel kopyası olan kaynak yerelden, olmayan CDN'den."""
    hedef = Path(hedef) if hedef else VARLIK_KLASORU
    return [f"{URL_ONEKI}/{ad}/stil.css" if _yerel_var(hedef, ad) else url for ad, url in (kaynaklar or KAYNAKLAR)]


def rota_kaydet(server, hedef=None):
    """Flask sunucusuna /yerel/<yol> adresini ekler (yalnızca yerel_varliklar klasöründen dosya verir)."""
    from flask import abort, send_from_directory
    klasor = Path(hedef) if hedef else VARLIK_KLASORU

    @server.route(f"{URL_ONEKI}/<path:yol>")
    def yerel_varlik(yol):
        if ".." in Path(yol).parts or yol.startswith("."):
            abort(404)
        return send_from_directory(str(klasor), yol, max_age=31536000)
    return yerel_varlik


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    print("Çevrimdışı arayüz dosyaları indiriliyor...")
    for r in indir():
        print(f"  {'✓' if r['ok'] else '✗'} {r['etiket']}: " + (f"{r['dosya']} dosya" if r["ok"] else r["hata"]))
    print(f"Klasör: {VARLIK_KLASORU}\nUygulamayı yeniden başlatın.")
