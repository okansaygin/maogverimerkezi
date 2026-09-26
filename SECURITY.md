# Güvenlik ve Kişisel Veri Politikası

Veri Merkezi; TC kimlik numarası, iletişim, sağlık (kan grubu, muayene), puantaj ve ücretle ilgili
**kişisel veriler** işler. Bu nedenle (KVKK kapsamında):

- **Veritabanı (`personel_veritabani.db`), yedekler (`veritabani_yedekleri/`), günlük dosyası
  (`veri_merkezi_gunluk.log`), `OUTPUT/` çıktıları ve kaynak Excel dosyaları asla repoya eklenmez.**
  `.gitignore` bunları dışarıda tutar; CI'daki *Kişisel veri kontrolü* adımı da yanlışlıkla eklenirse hatayı yakalar.
- Hata bildirirken log satırlarını paylaşmadan önce TC numaralarını ve adları maskeleyin
  (log, `/personel/<TC>` biçiminde adresler içerir).
- Uygulama yalnızca `127.0.0.1` üzerinde dinler; ağa açılmamalıdır.
- Veritabanı ağ sürücüsünde (SMB/NFS) tutulmamalıdır (WAL kipi güvenilir çalışmaz).

Bir güvenlik açığı ya da yanlışlıkla paylaşılmış kişisel veri fark ederseniz herkese açık issue açmak yerine
doğrudan repo sahibine bildirin. Git geçmişine girmiş bir veri dosyası yalnızca silinerek temizlenmez;
geçmişin yeniden yazılması gerekir.
