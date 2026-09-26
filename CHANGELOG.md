# Değişiklik Günlüğü

Bu dosya Veri Merkezi'nin sürüm notlarını içerir (en yeni en üstte).
Sürüm numarasının tek kaynağı [`veri_merkezi_surum.py`](veri_merkezi_surum.py) dosyasıdır.

---

## Veri Merkezi — Sürüm 5.2.0 (2026-09-25) · Yeni Personel Kartı ve İK bilgileri

### Personel Kartı (yeniden tasarlandı)
- **Kimlik bandı (yaka kartı):** solda lacivert şerit (baş harfler, personel no), ortada ad, İK'daki resmî ad farklıysa
  uyarı, grup › alt ekip › görev, TC (gizli, göz düğmesiyle), yaş, kıdem, öğrenim; durum çipleri (Aktif/Ayrılmış,
  İK durumu, Emekli/Engelli, geçici görevlendirme, kritik uyarı). Sağda kırmızı çerçevede **sahada acil durum**:
  kan grubu, aranacak kişi + yakınlığı + telefonu, kendi telefonu.
- **İSG şeridi:** tekrar muayene, tekrar İSG eğitimi, MYK belgesi, evrak. Süresi geçen kırmızı, 30 gün içinde dolan
  turuncu, geçerli yeşil.
- **Sekmeler:** Özet (aylık takvim + çalışma özeti: işe giriş/kıdem, ilk-son çalışma, toplam saat, yıllık FM sınırına
  kalan) · Puantaj · **Kişisel bilgiler** (kimlik, eğitim, acil durum, iletişim) · **İş ve sözleşme** (Veri Merkezi
  kaydı yan yana İK kaydı; görev farklıysa işaretli; giriş/sözleşme; konaklama/ulaşım; İK kayıt geçmişi — işten çıkıp
  yeniden girenler dahil) · **İSG ve evrak** (tarihler + durum, 8 evrakın listesi) · Değişiklik geçmişi · Puantaj
  değişiklikleri · Uyarılar.
- "Kartı Excel'e aktar": iki sayfa — Kişi bilgileri (Veri Merkezi + bütün İK alanları) ve Puantaj.
- İK listesinde olmayan kişide bu bölümler boş durum mesajı gösterir; sayfa çalışmaya devam eder.

### İK bilgileri (`veri_merkezi_ik.py`)
- `personel_ik_bilgileri` tablosu artık uygulamanın şemasının parçası (yeni veritabanlarında da oluşur).
- Yeni bir İK listesi geldiğinde (uygulama kapalıyken): `python veri_merkezi_ik.py "PERSONEL LİSTESİ.xlsx"`.
  Önce yedek alınır; veritabanında olmayan kişi atlanır; personnel'de yalnız boş işe giriş tarihi (SGK girişten)
  doldurulur; grup / alt ekip İK'dan alınmaz.
- Personel silinince İK kaydı da silinir.

---

## Veri Merkezi — Sürüm 5.1.1 (2026-09-25) · Açılış kontrolü

- Yeni `veri_merkezi_baslangic.py`: uygulama açılmadan önce klasöre ve veritabanına yazılabildiğini dener. Yazılamıyorsa
  (Windows "Denetimli klasör erişimi", güvenlik yazılımı, salt okunur dosya, kilitli veritabanı) onlarca satırlık hata
  yerine tek bir Türkçe açıklama ve uyarı penceresi gösterip kapanır.
- Günlük dosyası yazılamazsa konsola "--- Logging error ---" dökülmez; uygulama etkilenmez.

---

## Veri Merkezi — Sürüm 5.1.0 (2026-09-25) · Hafta tatili ve pazar kuralı

### 1. Yevmiye hesabı: hafta tatili / pazar kuralı
Yeni modül: `veri_merkezi_hafta_kurali.py` — yevmiyenin TEK kaynağı. Genel Bakış, Puantaj Geçmişi, Personel Kartı,
Verimlilik Analizi (ekran + Excel) hepsi bunu kullanır.

- **Pazartesi–cumartesi:** normal ×1, fazla mesai ×1,5.
- **Pazar, haftanın 6 günü tamamlandıysa:** (normal + fazla mesai) ×2,5. (Önceden yalnız pazar FM'si ×2,5'ti.)
- **Haftada ücretsiz izin / devamsızlık varsa ve pazar çalışıldıysa:** pazar sıradan iş günü (normal ×1, FM ×1,5);
  tek eksik gün hafta tatiline döner (ör. cuma ücretsiz izin → HT).
- **Haftada ücretsiz izin / devamsızlık var ve pazar çalışılmadıysa:** pazar da ücretsiz izin sayılır.
- **Mazeretli günler** (sağlık raporu, yıllık / babalık / vefat vb. ücretli izin, resmî tatil) çalışılmış sayılır.
- Hafta pazartesi–pazar. Dönem haftanın ortasında başlasa / bitse de kural haftanın tamamına bakar.

Kararlaştırılması gereken iki uç durum şöyle uygulandı (farklı olmasını isterseniz tek satırlık değişiklik):
- Pazartesi–cumartesi arasında **HAFTA TATİLİ** kodlu bir gün varsa hafta tatili o güne kaydırılmış sayılır; pazar
  çalışması ×2,5 değil, sıradan iş günü.
- **Kaydı hiç olmayan** gün (işe giriş öncesi, eksik yükleme) kuralı tetiklemez; yalnız açıkça ücretsiz izin /
  devamsızlık yazan gün tetikler.

Ekranlarda: Personel Kartı takviminde "PAZAR ×2,5" / "PAZAR ×1", "HT'YE DÖNDÜ", "ÜCRETSİZ SAYILDI" etiketleri (üzerine
gelince nedeni yazar). Verimlilik ve Excel'de "Pazar FM sa" sütunu **"Pazar ×2,5 sa"** oldu (×2,5 ödenen normal + FM
saatleri). Excel'de Yevmiye / FM yevmiye artık hücre formülü değil, kurala göre hesaplanmış değerdir (hangi pazarın
×2,5 olduğu formülle çıkarılamaz); toplam satırları yine formüldür. Devam Takviminde dönüşen günler HT, pazar zemini yalnız
×2,5 ödenen pazarlarda. Ayarlar'daki "Pazar katsayısı" artık normal + FM'ye uygulanır.

### 2. Günlük mesai formu (PDF)
- Başlık: GRUP / ALT EKİP / PERSONEL / TARİH (FORMEN alanı kaldırıldı).
- "Puantaja işleyen" kutusu kaldırıldı; Hazırlayan ve Kontrol eden kaldı, sağ köşelerindeki "Formen" /
  "Şantiye Şefi" yazıları kaldırıldı. Kalan iki kutu, boşalan yeri dolduracak şekilde genişledi.

### Testler
`python -m unittest discover -s testler -v` → 46 test (12'si yeni kural için: kullanıcı senaryoları dahil).

---

## Veri Merkezi — Sürüm 5.0.0 (2026-09-25) · Güvenilirlik

### Kurulum
1. **Önce uygulamayı kapatın ve uygulama klasörünün tamamını yedekleyin** (ör. `cp -r VeriMerkezi VeriMerkezi_v4_yedek`).
2. Paketteki dosyaları uygulama klasörünün üzerine kopyalayın. `assets/veri_merkezi.css` → `assets/` klasörüne, `testler/` klasörü olduğu gibi.
3. Artık kullanılmayan iki dosyayı arşive taşıyın (silmek de olur; hiçbir yerden çağrılmıyorlar):
   ```
   mkdir -p arsiv
   mv veri_merkezi_app_eski.py veri_merkezi_fazla_mesai.py arsiv/
   ```
4. Yeni paket yok. Her zamanki gibi başlatın: `python veri_merkezi_main.py`
5. İsteğe bağlı: `python -m unittest discover -s testler -v` (34 test; gerçek veritabanınıza dokunmaz).
6. İsteğe bağlı, internet varken bir kez: Sistem › Ayarlar › Çevrimdışı çalışma › **Dosyaları indir**, sonra uygulamayı yeniden başlatın.

İlk açılışta: veritabanının şema v4 hâli `veritabani_yedekleri/` klasörüne yedeklenir, sonra şema v5'e yükseltilir ve WAL kipine alınır. Mevcut veri silinmez ya da değiştirilmez; yalnızca yeni tablolar ve sütunlar eklenir.

### Yenilikler
- **Puantaj değişiklik geçmişi.** Bir içe aktarma bir günün değerini değiştirdiğinde eski değer artık kaybolmaz (`puantaj_history`). Personel Kartı'nda yeni **Puantaj değişiklikleri** sekmesi: gün, işlem, önce → sonra, hangi içe aktarma, ne zaman.
- **İçe aktarmayı geri alma.** Denetim › İçe Aktarma Geçmişi › bir içe aktarma seç › **Bu içe aktarmayı geri al**. Önce ne olacağı gösterilir (silinecek / önceki değerine dönecek / korunacak kayıtlar, örnek satırlar), not yazılıp onaylanır.
  - Önce veritabanının yedeği alınır; işlem tek adımda yapılır, yarıda kalırsa hiçbir şey değişmez.
  - Aynı kayıtları sonradan değiştiren daha yeni bir içe aktarma varsa geri alma engellenir; engelleyenler listelenir, tıklayınca onlara geçilir. Geri alma her zaman en yeniden eskiye doğru yapılır.
  - Personel Düzenle'den elle yapılmış düzeltmeler korunur. Puantaj kaydı olan ya da elle düzenlenmiş kişi silinmez.
  - İçe aktarmanın açık uyarıları "İçe aktarma geri alındı" notuyla kapatılır. Kayıt silinmez, durumu **geri alındı** olur; aynı dosya yeniden yüklenebilir.
  - v5'ten önceki içe aktarmalar geri alınamaz (eski değerleri kayıtlı değil); bunlar için yedekten geri yükleme kullanılabilir.
- **Otomatik yedek** (`veritabani_yedekleri/`): günde bir açılış yedeği, her içe aktarma / geri alma / geri yükleme / şema yükseltme öncesi yedek. Her türden en yeni 30 tanesi tutulur (ayarlanabilir); elle alınanlar silinmez. Yedekler SQLite'ın yedekleme yöntemiyle, uygulama çalışırken bile tutarlı alınır; her biri tek başına açılabilen bir `.db` dosyasıdır.
- **Yedekten geri yükleme** (Ayarlar › Veritabanı yedekleri): yedek önce bütünlük kontrolünden geçer, şimdiki ve yedekteki kayıt sayıları yan yana gösterilir, onay istenir ve geri yüklemeden önce şimdiki hâlin yedeği alınır.
- **WAL kipi.** Okuma ve yazma birbirini beklemez; "database is locked" hatası pratikte ortadan kalkar. Bağlantılar 15 sn bekleyebilir. Kapanışta değişiklikler ana dosyaya birleştirilir.
- **Ayarlar sayfası** (yeni menü: Sistem › Ayarlar) ve `veri_merkezi_ayarlar.json`: proje adı, menü alt başlığı, yıllık FM sınırı, günlük azami saat, kesintisiz çalışma ve devamsızlık eşikleri, aylık FM eşiği (önceden Puantaj Suite'ten okunuyordu), FM katsayıları, mesai formu boş satır sayısı, yedekleme. Değerler kaydederken doğrulanır; hatalı alan varsa hiçbir şey yazılmaz. Verimlilik Analizi ekranı ve Excel raporundaki bütün etiketler ("270 sa", "11 sa", katsayılar) artık bu ayarlardan gelir; eşik yasal değerden farklıysa yasa metninin yanında "uygulanan eşik" yazar.
  - Standart iş günü (9 saat) bilerek ayar değildir: içe aktarmada normal / fazla ayrımı buna göre yapılıp kalıcı yazılır.
- **Çevrimdışı çalışma** (`veri_merkezi_varliklar.py`): Bootstrap, ikonlar ve yazı tipleri bir kez indirilip `yerel_varliklar/`'dan sunulur. İndirilmemişse eskisi gibi internetten yüklenir. Yazı tipi yedeklerine Fedora/Linux yazı tipleri (Noto, DejaVu, Liberation) eklendi.
- **Tek sürüm numarası** (`veri_merkezi_surum.py`): sol menünün altında ve Ayarlar › Sistem bilgisi'nde.
- **Başlatıcı:** sabit 1,5 sn bekleme yerine sunucunun hazır olması beklenir.

### Düzeltmeler
- `migrate()` her çağrıda şema sürümünü yeniden yazıyordu; bu, Ekip Listesi'ni açmak gibi işlemlerde bile Verimlilik Analizi önbelleğini boşa silerek sayfayı yavaşlatıyordu. Artık yalnızca sürüm değiştiğinde yazıyor. Daha eski bir uygulama da şema numarasını geriye çekemiyor.
- İçe Aktarma Geçmişi'ndeki "Yazılan günlük kayıt" sayısı yalnızca geçerli (tamamlanmış) içe aktarmaları sayıyor; "Etkilenen personeli listele" geri alınmış içe aktarmalarda da çalışıyor.

### Bilinmesi gerekenler
- **Veritabanını elle kopyalamayın** (ya da önce uygulamayı kapatın): uygulama açıkken son değişiklikler bir süre `personel_veritabani.db-wal` dosyasında durur. Ayarlar › **Şimdi yedek al** her zaman güvenlidir.
- **Veritabanı ağ sürücüsünde olmamalı:** WAL kipi ağ paylaşımlarında (SMB/NFS) güvenilir çalışmaz. Yerel diskte sorun yok.
- Puantaj Suite aynı veritabanını kullanmaya devam eder; WAL kipiyle uyumludur, Suite'te değişiklik yok.
- Yeni dosyalar/klasörler: `veri_merkezi_ayarlar.json` (ilk kayıtta oluşur), `veritabani_yedekleri/`, `yerel_varliklar/` (indirilince).
- Geliştirme ortamında gerçek Dash kurulamadığı için arayüz, sahte bir Dash katmanıyla bütün sayfalar açılarak ve yeni callback'lerin tamamı çalıştırılarak denendi; ayrıca gerçek bir v4 veritabanının v5'e yükseltilmesi denendi. İlk gerçek çalıştırmada küçük görsel düzeltmeler gerekebilir; `veri_merkezi_gunluk.log` dosyasını iletmeniz yeterli.

### Dosyalar
| Dosya | Durum |
|---|---|
| `veri_merkezi_surum.py`, `veri_merkezi_ayarlar.py`, `veri_merkezi_yedek.py`, `veri_merkezi_geri_alma.py`, `veri_merkezi_varliklar.py`, `veri_merkezi_sayfa_ayarlar.py` | YENİ |
| `testler/test_v5.py` | YENİ: 34 otomatik test |
| `veri_merkezi_sema.py` | Şema v5, WAL, bekleme süresi, gereksiz yazma düzeltmesi |
| `veri_merkezi_ice_aktarma.py`, `veri_merkezi_mesai_ice_aktarma.py` | İşlem öncesi yedek + değişiklik izi |
| `veri_merkezi_sayfa_gecmis.py` | Geri alma akışı, "geri alındı" durumu |
| `veri_merkezi_sayfa_personel.py` | Puantaj değişiklikleri sekmesi |
| `veri_merkezi_sorgu.py`, `veri_merkezi_verimlilik.py`, `veri_merkezi_verimlilik_excel.py`, `veri_merkezi_sayfa_verimlilik.py` | Eşik ve katsayılar ayarlardan; etiketler dinamik |
| `veri_merkezi_app.py`, `veri_merkezi_main.py`, `veri_merkezi_sayfa_ice_aktar.py`, `personnel_db.py`, `assets/veri_merkezi.css` | Bağlantılar, menü, başlatma sırası, küçük eklemeler |
| `veri_merkezi_app_eski.py`, `veri_merkezi_fazla_mesai.py` | Kullanılmıyor → `arsiv/` |

---

## Veri Merkezi — Yeni Arayüz (sürüm 2026-09-24.1)

### Kurulum
1. Bu klasördeki dosyaları mevcut Veri Merkezi klasörünün üzerine kopyalayın (önce klasörün yedeğini alın).
   `assets/` klasörü de `veri_merkezi_app.py` ile **aynı klasörde** olmalı; Dash oradaki `.css` / `.js` dosyalarını otomatik yükler.
2. Paketler: `pip install dash dash-bootstrap-components plotly pywebview openpyxl reportlab`
   (`reportlab` günlük mesai formu PDF'i için yeni eklendi.) `fonts/` klasörü de uygulama klasöründe olmalı.
   Dash 2.9 veya üstü gerekir.
3. Her zamanki gibi başlatın: `python veri_merkezi_main.py`

Veritabanı ilk açılışta kendiliğinden şema v4'e yükseltilir. Sadece `import_batches` tablosuna `tur` sütunu eklenir; mevcut veri silinmez ya da değiştirilmez.

### Dosyalar
| Dosya | Durum |
|---|---|
| `veri_merkezi_app.py` | Yeniden yazıldı: sadece kabuk (sol menü, yönlendirme, genel arama). Eski sürümü `yedek/` klasöründe |
| `veri_merkezi_ui.py` | YENİ: ortak bileşenler (kart, gösterge, çip, tablo, segment düğme…) |
| `veri_merkezi_sayfa_*.py` | YENİ: her ekran ayrı modül (genel, ice_aktar, sap, kalite, gecmis, personel, puantaj, ekip, raporlar) |
| `assets/veri_merkezi.css`, `assets/kisayollar.js` | YENİ: tasarım sistemi ve Ctrl+K kısayolu |
| `veri_merkezi_sorgu.py` | Sona yeni sorgular eklendi; mevcut fonksiyonlar değişmedi |
| `veri_merkezi_dosya_tespit.py` | YENİ: yüklenen Excel'in kimlik mi SAP mesai mi olduğunu başlıklardan anlar |
| `veri_merkezi_fazla_mesai.py` | Artık kullanılmıyor; yerini Verimlilik Analizi aldı (aşağıya bakın). Silinebilir |
| `veri_merkezi_maliyet.py` | YENİ: Maliyet Raporu, personel kapsamı depodan gelir; hakediş okuma ve Excel çıktısı cost_report ile yapılır |
| `veri_merkezi_sema.py` | Şema v4 (`import_batches.tur`) |
| `veri_merkezi_ice_aktarma.py`, `veri_merkezi_mesai_ice_aktarma.py` | Tek değişiklik: içe aktarma kaydına `tur` yazılıyor |
| Diğerleri (engine, team_report, cost_report, personnel_db, suite…) | Değişmedi |

### Ekranlardaki yenilikler
- **Genel Bakış:** Ay seçici var. Üstte 4 adımlı durum şeridi, eksik günleri gösteren kapsama takvimi ve kategoriye göre açık uyarılar.
- **İçe Aktar:** İki içe aktarma sayfası tek sihirbazda birleşti. Dosya türü otomatik anlaşılıyor ve sütun eşleştirmesi tabloda görünüyor. Önizlemede yeni / güncellenecek / değişmeyen sayıları var. Çakışmalar tek tek ya da toplu olarak (mevcudu koru / yeniyi kullan) karara bağlanıyor.
- **SAP Puantaj Aktarım:** Gün, takvimden seçiliyor. Hedef dosyalar sekmelerde önizleniyor. Eşleşmeyen satırlar için isim benzerliğine göre öneri çıkıyor ve **Ata** ile eşleştiriliyor (motorun `manual_mapping` özelliği).
- **Veri Kalitesi:** Uyarılar tek tek ya da toplu olarak, çözüm notuyla **çözüldü** işaretlenebiliyor. Çözülmüş uyarılar ayrı görüntülenebiliyor, gerekirse yeniden açılabiliyor. İçe aktarmaya göre süzme, Excel'e aktarma ve son içe aktarmalardaki eğilim de var.
- **İçe Aktarma Geçmişi:** Tür ve durum filtresi var. Detay panelinde süre, hata mesajı, uyarı dağılımı ve etkilenen personel görünüyor.
- **Personel:** Grup, durum ve uyarı filtreleri; sayfalama; Excel'e aktarma. **Personel Kartı** yeni: günlük puantaj takvimi, dönem toplamları, değişiklik geçmişi ve uyarılar. **Düzenle** ekranında TC anında kontrol ediliyor, mevcut değerler öneriliyor ve kaydetmeden önce eski → yeni farkı gösteriliyor.
- **Puantaj Geçmişi:** Çoklu grup / alt ekip / görev seçimi, hızlı tarih aralıkları, Matris ve Liste görünümü, Excel'e aktarma.
- **Ekip Listesi:** Excel çıktısının görev gruplarına bölünmüş önizlemesi.
- **Fazla Mesai Raporu (yeni):** Günlük / haftalık / aylık kırılım, önceki dönemle karşılaştırma ve alt ekiplerin fazla mesai oranı. Kişi bazında Suite'teki aylık eşik ve yıllık 270 saatlik yasal sınır kontrol ediliyor. Excel çıktısı da var.
- **Maliyet Raporu (yeni):** Hakediş dosyasındaki bordro sayfaları otomatik bulunuyor; sütunu kayık sayfa ayrıca işaretleniyor. Maliyet dağılımı, ekip ve görev bazlı maliyet ve bir kontrol listesi var. Excel çıktısı `cost_report` biçiminde.

### Bilinmesi gerekenler
- Yazı tipleri (Archivo, IBM Plex) ve ikonlar internetten yüklenir. İnternet yoksa sistem yazı tipi (Segoe UI) kullanılır ve ikonlar görünmez, ama uygulama çalışır.
- Geliştirme ortamında Dash kurulamadığı için arayüz gerçek Dash'le çalıştırılamadı. Bunun yerine test amaçlı sahte bir Dash katmanıyla tüm sayfalar ve callback'lerin büyük çoğunluğu örnek veriyle çalıştırıldı. Bu testlerde içe aktarma, tekrar engelleme, çakışma kararı, uyarı çözme, personel kaydetme ve silme, SAP aktarımı ve rapor Excel'leri doğrulandı. İlk gerçek çalıştırmada küçük görsel düzeltmeler gerekebilir; hata mesajını iletmeniz yeterli.

### Güncelleme — Günlük Mesai Formu (PDF)
- **Nerede:** Ekip Listesi sayfasındaki "Günlük mesai formu" kartı. Grup, alt ekip ve görev filtresini seçip **Mesai formu oluştur (PDF)** düğmesine basın. PDF `OUTPUT/` klasörüne kaydedilir ve otomatik açılır.
- **Kurallar:**
  - Her alt ekip kendi formunu alır: kendi başlığı, formeni ve sayfaları var.
  - Sayfa başına en fazla 20 kişi sığar; fazlası sonraki sayfaya geçer ve sıra numarası sayfalar boyunca devam eder.
  - Satır yüksekliği ve yazı boyutu kişi sayısına göre otomatik ayarlanır.
  - Tarih elle yazılır.
  - Son sayfaya, listede olmayanlar için boş satırlar eklenir. Sayısı ayarlanabilir, varsayılan 3.
- **İzin kodları:** Formda Yİ, RP, ÜCS, Bİ, Vİ, HT ve yeni **DV (Devamsızlık)** yazılıdır. `puantaj_engine.py` içindeki `LEAVE_CODE_MAP` tablosuna `"DV": "DEVAMSIZLIK"` eklendi; bu, dosyadaki tek değişiklik. Puantaj Suite de bu kodu artık tanıyor.
- **Dosyalar:**
  - `veri_merkezi_mesai_formu.py`: PDF motoru (yeni).
  - `veri_merkezi_sayfa_ekip.py`: form kartı eklendi.
  - `fonts/`: DejaVu Sans yazı tipleri, Türkçe karakterler için. Lisansı `DejaVu_LICENSE.txt`'de, serbestçe dağıtılabilir.
- **Proje adı:** Formda "MAOG Projesi" yazar. Değiştirmek için `veri_merkezi_mesai_formu.py` içindeki `PROJE_ADI` değerini düzenleyin.
- **.exe'ye çevirirken:** `fonts` klasörünü de pakete ekleyin (`--add-data "fonts:fonts"`). Eklenmezse uygulama sistemdeki DejaVu, Liberation ya da Arial yazı tipini kullanmayı dener.

### Güncelleme — Performans (2026-09-25)
- `veri_merkezi_sema.migrate()`: Eski `daily_timesheet` arşivinin `gunluk_puantaj`'a kopyalanması artık yalnızca şema sürümü 3'ten küçükken bir kez yapılıyor. Önceden Ekip Listesi her açıldığında bu kopyalama tekrarlanıyordu; sayfanın geç yüklenmesinin asıl sebebi buydu.
- Ekip Listesi önizlemesi ilk açılışta 60 satır gösteriyor; **Tümünü göster** ile tamamı açılıyor. Excel ve mesai formu her zaman tam listeyle üretiliyor.
- Üst çubuktaki personel araması her sayfada bütün personel listesini tarayıcıya göndermiyor; en az 2 harf yazınca sunucudan en fazla 25 eşleşme geliyor. Türkçe harf farkları eşleşmeyi bozmuyor ("sahin" yazınca "ŞAHİN" bulunuyor).

### Güncelleme — Ekip Listesi'nde göster / gizle
- Önizlemedeki her personel satırının en sağında göz ikonu var. Tıklanan kişi gizlenir: satırı soluk görünür ve sıra numarası verilmez.
- Gizlenen kişiler **Excel ekip listesine** ve **günlük mesai formuna (PDF)** girmez. Sıra numaraları, kişi sayıları, görev dağılımı ve formun sayfa sayısı buna göre yeniden hesaplanır.
- Önizlemenin üstündeki bantta kaç kişinin gizli olduğu yazar. **Hepsini geri al** ile tamamı listeye döner; göz ikonuna tekrar tıklamak da tek kişiyi geri alır.
- Gizleme sadece o ekranda geçerlidir: veritabanında hiçbir şey değişmez. Sayfadan çıkınca ya da uygulamayı kapatınca sıfırlanır.

#### Düzeltme (2026-09-25)
- **Hata:** Göz düğmesine basınca kişi gizlenmiyordu.
- **Sebep:** Göz düğmeleri ile "Hepsini geri al" düğmesi aynı callback'teydi. Dash bir callback'i ancak bütün girdileri sayfada varken çalıştırır. "Hepsini geri al" düğmesi sadece gizli kişi varken çizildiği için callback hiç tetiklenmiyordu.
- **Çözüm:** İki düğme ayrı callback'lere ayrıldı. Aynı türde başka bir hata var mı diye tüm sayfalar tarandı; başka yok.

### Güncelleme — Ekip Listesi Excel'i tek sayfa
- Veri Merkezi'nden alınan ekip listesi Excel'i artık **tek bir sayfada**, yukarıdan aşağıya şu bölümleri içeriyor:
  1. Personel listesi (görev gruplarına bölünmüş, gizlenenler hariç)
  2. **Göreve göre personel sayısı** tablosu ve toplam
  3. **Seçilen ayda çıkış yapan personel:** ad soyad, TC, görev, çıkış tarihi
- Ay, Excel kartındaki "Çıkış yapanlar" listesinden seçiliyor; varsayılan içinde bulunulan ay. Kartta o ay kaç kişinin çıkış yaptığı önceden yazıyor. Çıkış yapanlara da grup, alt ekip ve görev filtresi uygulanıyor.
- Yeni dosya: `veri_merkezi_ekip_excel.py`. Puantaj Suite'in kullandığı `team_report.write_roster_report` (iki sayfalı çıktı) değişmedi.

### Güncelleme — İçe aktarmada elle sütun seçimi
- Sütunlar başlıktan tanınmadığında (örneğin TC No A sütunu yerine B sütununda olduğunda ya da başlık farklı yazıldığında) içe aktarma durur. Soldaki **Sütun eşleştirme** tablosunda her alanın yanında bir **Sütun** kutusu var.
- Eksik alanlar kırmızıyla işaretlenir. Kutuya doğru sütun harfini yazın (A, B, … AA), sonra **Eşleştirmeyi uygula ve yeniden dene**'ye basın. Dosyayı tekrar yüklemeniz gerekmez.
- **Örnek değer** sütunu, seçilen sütundaki ilk veriyi gösterir; doğru sütunu seçip seçmediğinizi buradan görebilirsiniz. Elle girilen sütunlar "elle" olarak işaretlenir.
- Dosya türü yanlış anlaşıldıysa tablonun üstünden **Personel bilgileri** ya da **SAP mesai** seçilebilir.
- Geçersiz harf ya da iki alana aynı sütunun verilmesi uyarıyla engellenir. Boş kutu, o alan için otomatik tespit kullanılır demektir.
- Önizleme ekranında da tablo düzenlenebilir: eşleştirme yanlışsa düzeltip yeniden deneyebilirsiniz.
- Değişen dosyalar:
  - `veri_merkezi_dosya_tespit.py`
  - `veri_merkezi_sayfa_ice_aktar.py`
  - `veri_merkezi_ice_aktarma.py` ve `veri_merkezi_mesai_ice_aktarma.py`: önizleme fonksiyonlarına isteğe bağlı `colmap_override` parametresi eklendi. Parametre verilmezse davranış aynı kalır.

#### Düzeltme (2026-09-25): dosya seçildikten sonra ilerlememe
- **Hata:** İçe Aktar'da dosya seçildikten sonra sihirbaz ilerlemiyor, dosya seçim ekranı olduğu gibi kalıyordu.
- **Sebep:** Sütun eşleştirme tablosundaki harf kutularına (`dcc.Input`) `aria-label` özelliği verilmişti. Dash bu tür özellikleri yalnızca `html.*` bileşenlerinde kabul ediyor. Bu yüzden ekran oluşturulurken hata çıkıyor ve Dash bu hatayı ekranda göstermeden yutuyordu.
- **Çözüm:** Özellik kaldırıldı; açıklama hücrenin `title` özelliğine taşındı.
- **Ek önlem:** İçe aktarmada beklenmeyen bir hata olursa ekran artık donmuyor; hata mesajı hem ekranda hem konsolda görünüyor.
- **Tarama:** Diğer sayfalarda aynı türde hata olmadığı kontrol edildi. `aria-label` kullanılan diğer iki yer (Ekip Listesi'ndeki göz düğmesi ve Personel Kartı'ndaki TC düğmesi) `html.Button`; orada bu özellik geçerli.

### Güncelleme — Verimlilik Analizi (eski adıyla Fazla Mesai Raporu)
Menüdeki **Fazla Mesai Raporu** yerine **Verimlilik Analizi** geldi. Adresi `/verimlilik`; eski `/fazla-mesai` adresi de aynı sayfayı açar.

**Sayfa**
- **Filtreler:**
  - Dönem; hızlı seçim düğmeleri: Bu ay, Geçen ay, Son 30 gün, Son 3 ay, Bu yıl.
  - Grafik kırılımı: günlük, haftalık ya da aylık.
  - Grup, alt ekip ve görev.
  - İsteğe bağlı **günlük yevmiye ücreti (TL)**. Girilirse yevmiye ve fazla mesai TL olarak da gösterilir. Değer bu bilgisayarda hatırlanır.
- **Üst bölüm:**
  - **Verimlilik puanı:** A–D notu ve periyot bazında seyri.
  - **Yönetici özeti:** veriden otomatik yazılır; iş gücü, devam, fazla mesai, maliyet etkisi, en iyi ve en zayıf ekip, uyum riskleri, personel hareketi ve eksik veri başlıklarını kapsar.
  - **10 gösterge:** her biri önceki döneme göre değişimle birlikte; iyileşme yeşil, kötüleşme kırmızı gösterilir.
- **Sekmeler:**
  - **Genel görünüm:** çalışma ve fazla mesai grafiği (alt satırda devam oranı), gün dağılımı halkası, günlük iş gücü (çalıştı, izin, rapor, devamsızlık), haftanın günleri ve personel hareketi.
  - **Ekip ve görev:** ekip/grup ve görev karşılaştırma tabloları (sıra, devam, kapasite, fazla mesai ve değişimi, kayıp gün, yevmiye, puan) ve alt ekip × gün **devam haritası**.
  - **Personel karnesi:** en iyi ve en zayıf beşli listeler, tüm personel için sıralanıp süzülebilen tablo.
  - **Uyum ve risk:** yıllık 270 sa (md. 41), günlük 11 sa (md. 63), 7+ gün kesintisiz çalışma (md. 46), dönem fazla mesai eşiği ve 3+ gün devamsızlık.
  - **Yöntem:** her göstergenin nasıl hesaplandığı.
- Grafikler HTML/CSS ile çizilir; plotly gerekmez. Üzerine gelince ayrıntı görünür.

**Excel raporu (11 sayfa)**

| Sayfa | İçerik |
|---|---|
| Özet | Puan, 10 gösterge, öne çıkanlar, dönem karşılaştırması, gün dağılımı grafiği. A4'e tek sayfa basılır |
| Trend | Periyot tablosu ve grafik (normal + fazla mesai, ikinci eksende devam oranı) |
| Günlük | Günlük iş gücü tablosu ve grafiği |
| Grup / Ekip, Görev | Karşılaştırma tabloları ve grafik |
| Devam Haritası | Alt ekip × gün devam oranı |
| Personel Karnesi | Herkes, bütün göstergeler; süzme hazır. A3 |
| Devam Takvimi | Kişi × gün: saat ya da izin kodu, renkli. A3 |
| Uyum ve Risk | Risk listeleri |
| Personel Hareketi | İşe giren ve ayrılan personel |
| Yöntem | Tanımlar ve formüller |

- Oranlar, toplamlar, yevmiye, puan ve not **hücre formülüdür**: bir saat düzeltilirse satır kendiliğinden güncellenir. Formüllerin sonuçları ekrandaki değerlerle karşılaştırıldı; tümü aynı çıktı.
- Dosya adı: `OUTPUT/<kapsam>_verimlilik_<başlangıç>_<bitiş>.xlsx`

**Puan ve tanımlar**
- **Verimlilik puanı** = devam oranı × 0,6 + kapasite kullanımı (en fazla %100) × 0,4
- **Not:** A ≥ 95 · B 85–95 · C 70–85 · D < 70
- **Devam oranı** = çalışılan gün ÷ planlanan iş günü. Hafta tatili ve resmî tatil planlanan güne sayılmaz.
- **Kapasite kullanımı** = normal saat ÷ (planlanan gün × 9)
- **Kayıp gün** = devamsızlık + rapor + ücretsiz izin
- **FM katsayı farkı** = aynı saatler normal mesaide çalışılsaydı ödenmeyecek ek yevmiye
- Çalışma durumu metni (ÇALIŞTI, HAFTA TATİLİ, SAĞLIK RAPORU, DEVAMSIZLIK …) otomatik sınıflandırılır. Kurallar `veri_merkezi_verimlilik.kategori()` içinde. Eşikler dosyanın başında: `DEVAMSIZLIK_ESIGI`, `KESINTISIZ_GUN_ESIGI` vb.

**Dosyalar**
- `veri_merkezi_verimlilik.py`: analiz motoru (yeni)
- `veri_merkezi_verimlilik_excel.py`: Excel raporu (yeni)
- `veri_merkezi_sayfa_verimlilik.py`: sayfa (yeni)
- `veri_merkezi_sayfa_raporlar.py`: artık yalnız Maliyet Raporu
- `veri_merkezi_app.py`, `veri_merkezi_sayfa_genel.py`: menü ve bağlantı adı
- `assets/veri_merkezi.css`: yeni stiller

#### Düzeltme (2026-09-25): Verimlilik Analizi geç açılıp Genel Bakış'a dönüyordu
- **Sebep:**
  - Üst çubuktaki personel arama kutusu her sayfayla yeniden çiziliyor. Bu yüzden ona bağlı `genel_arama_git` callback'i her sayfa açılışında çalışıyordu.
  - Bu callback'in çıktısı ana adres bileşeni `url` idi. Dash, bir callback'in çıktısı olan bileşen beklemedeyken ona bağlı sayfa yönlendirmesini bekletir. Sunucu Verimlilik analizi gibi ağır bir işle meşgulken yeni sayfa bu yüzden geç açılıyordu.
  - `url` bileşeni `refresh=True` ile çalışır. Bu bekleme sırasında eski adresle yeniden çizilirse tarayıcıyı o eski adrese (`/` = Genel Bakış) tamamen yeniler; sayfa açıldıktan hemen sonra Genel Bakış'a dönülmesinin sebebi bu.
  - Bekleme, gerçek Dash 2.18, 3.4 ve 4.4 ile tarayıcıda tekrarlanarak doğrulandı.
- **Çözüm:**
  - Genel arama artık ayrı bir `dcc.Location(id="arama-git", refresh="callback-nav")` üzerinden yönlendiriyor. `url` hiçbir callback'in çıktısı değil; sayfa değişimi hiçbir şeyi beklemiyor ve eski adrese yenileme olamıyor.
  - Aynı kişi ikinci kez arandığında da yönlendirme çalışıyor; adrese `?a=...` eklenir, zararsızdır.
- **Hız:**
  - Verimlilik analizi yeniden düzenlendi: her satır 9 yerine 2 toplayıcıya ekleniyor ve gereksiz sıralı sorgu kaldırıldı. 700 kişi × 9 ay 3,6 sn'den 2,0 sn'ye, bir ay 0,8 sn'den 0,55 sn'ye indi. Sonuçlar eskisiyle birebir aynı.
  - Sonuçlar önbelleğe alınıyor: aynı filtrelerle sayfaya dönmek ya da Excel almak anında oluyor. Veritabanı değişince (içe aktarma, düzenleme) önbellek kendiliğinden yenilenir.
- **Diğer:**
  - Bilinmeyen bir adres artık sessizce Genel Bakış'ı açmıyor; "Sayfa bulunamadı" gösteriyor.
  - Uygulama klasörüne `veri_merkezi_gunluk.log` yazılıyor: hangi sayfa ne zaman, kaç saniyede açıldı, analiz süreleri ve hatalar. Bir sorun olursa bu dosyayı iletmen yeterli.

#### Düzeltme 2 (2026-09-25): Verimlilik Analizi hâlâ Genel Bakış'a dönüyordu — asıl sebep
- **Asıl sebep:** Personel karnesi tablosundaki bir renklendirme kuralı (`{risk} is nonblank`) DataTable için geçersiz bir sorgu. Tablo tarayıcıda çizilirken hata veriyordu.
- **Neden başka sayfaya atıyordu:** Dash, bir bileşen çizim hatası verince sayfa durumunu bir adım geri alır. Geri alınan durumda adres de önceki sayfa (`/` = Genel Bakış) olduğu için `dcc.Location` sayfayı oraya yeniliyordu. Hata sessiz olduğu için yalnızca "geri atıldım" gibi görünüyordu.
- **Neden önceki testlerde çıkmadı:** Tablo "Personel karnesi" sekmesinde. Test ortamımdaki dbc.Tabs yerine geçen bileşen sadece açık sekmeyi çiziyordu; gerçek dbc.Tabs bütün sekmeleri arka planda çizer, bu yüzden hata sizde hemen oluşuyordu.
- **Düzeltmeler:**
  - Kural geçerli sözdizimine çevrildi: `!({risk} is blank)`. Uyarısı olan kişilerin "Uyarılar" hücresi kırmızı yazılıyor.
  - `assets/veri_merkezi.css`'teki tablo hücresi rengi `!important` olduğu için bütün koşullu renkleri eziyordu. Buna Puantaj Geçmişi'ndeki turuncu fazla mesai de dahil. Kaldırıldı; koşullu renkler artık görünüyor.
  - `dbc.Tabs` bütün sayfalarda (Verimlilik, Personel Kartı, SAP önizleme) düz HTML sekmelerle (`ui.sekmeler`) değiştirildi. Tek bir callback yalnızca görünürlüğü değiştiriyor. Daha az dış bağımlılık, daha az sessiz hata riski.
  - **Tarayıcı tanı kaydı (yeni):** `assets/tani.js` tarayıcıdaki hataları ve tam sayfa yenilemelerini `veri_merkezi_gunluk.log` dosyasına yazar. Bu hatayı bulan da bu kayıt oldu. Benzer bir durumda bu dosyayı göndermen yeterli.
- **Test:** Gerçek Dash 2.18, 3.4 ve 4.4 ile tarayıcıda şu gezinme denendi: Genel → Verimlilik → sekmeler → Personel → Verimlilik → Ekip Listesi → Verimlilik. Hiçbirinde geri dönme olmadı ve tarayıcı hata kaydı boş.
