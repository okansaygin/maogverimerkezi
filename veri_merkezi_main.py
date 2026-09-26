# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - BAŞLATICI (veri_merkezi_main.py)
===================================================
Çift tıkla çalıştırılacak TEK giriş noktası. Dash sunucusunu arka planda
sessizce başlatıp, pywebview ile kendi masaüstü penceresinde açar -
tarayıcı adres çubuğu görünmez, native bir masaüstü uygulaması gibi
hissettirir.

Kurulum (bir defalık, komut satırında):
    pip install dash dash-bootstrap-components plotly pywebview openpyxl reportlab

Çalıştırma:
    python veri_merkezi_main.py

Tek dosya .exe'ye dönüştürmek isterseniz (opsiyonel):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name VeriMerkezi veri_merkezi_main.py
"""

import socket
import threading
import time

import logging
import sys
from pathlib import Path

# Uygulamayı yüklemeden ÖNCE: klasöre ve veritabanına yazılabiliyor mu? (yazılamıyorsa anlaşılır bir mesajla çık)
import veri_merkezi_baslangic as baslangic

_KLASOR = Path(__file__).resolve().parent
_sorun = baslangic.kontrol(_KLASOR, _KLASOR / "personel_veritabani.db")
if _sorun:
    baslangic.hata_goster(_sorun)
    sys.exit(1)

import webview

from veri_merkezi_app import app
import veri_merkezi_sema as sema

gunluk = logging.getLogger("veri_merkezi")

HOST = "127.0.0.1"


def _find_free_port():
    """İşletim sisteminden boş bir port ister - port çakışması asla olmaz."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind((HOST, 0))
        return s.getsockname()[1]


def _run_dash(port):
    # use_reloader=False ŞART: aksi halde Flask'ın otomatik yeniden yükleyicisi
    # ikinci bir alt süreç başlatır ve pencere/sunucu senkronizasyonu bozulur.
    app.run(host=HOST, port=port, debug=False, use_reloader=False)


def _sunucuyu_bekle(port, en_fazla_sn=30.0):
    """Sunucu bağlantı kabul edene kadar bekler (sabit bekleme yerine). Yavaş bir makinede
    pencerenin boş/hatalı açılmasını önler. Döner: sunucu hazır mı."""
    bitis = time.monotonic() + en_fazla_sn
    while time.monotonic() < bitis:
        try:
            with socket.create_connection((HOST, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.1)
    return False


def main():
    port = _find_free_port()
    server_thread = threading.Thread(target=_run_dash, args=(port,), daemon=True)
    server_thread.start()

    if not _sunucuyu_bekle(port):
        gunluk.error("Sunucu 30 sn içinde başlamadı (port %s)", port)

    webview.create_window(
        "Veri Merkezi",
        f"http://{HOST}:{port}",
        width=1440,
        height=900,
        min_size=(1100, 700),
    )
    webview.start()

    # Pencere kapandı: WAL dosyasındaki değişiklikleri ana .db dosyasına aktar (dosya tek başına güncel kalsın).
    sema.wal_birlestir()
    gunluk.info("Veri Merkezi kapatıldı")


if __name__ == "__main__":
    main()
