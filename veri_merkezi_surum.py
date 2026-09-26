# -*- coding: utf-8 -*-
"""
VERİ MERKEZİ - SÜRÜM (veri_merkezi_surum.py)
===============================================
Uygulamanın TEK sürüm numarası. Sol menünün altında, Ayarlar > Sistem bilgisi'nde
ve günlük dosyasında (veri_merkezi_gunluk.log) bu değer görünür. Hata bildirirken
hangi sürümün çalıştığı buradan anlaşılır.

Dosyaların başındaki eski `__version__` satırları yalnızca o dosyanın son
değişiklik tarihini gösterir; uygulamanın sürümü BUDUR.
"""

SURUM = "5.2.0"
SURUM_TARIHI = "2026-09-25"
SURUM_ADI = "v5.2 · Personel Kartı"


def etiket():
    """Kısa gösterim: 'v5.0.0 (2026-09-25)'"""
    return f"v{SURUM} ({SURUM_TARIHI})"
