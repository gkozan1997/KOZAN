# GökhanOS E-Ticaret Depo & Sipariş Toplama Sistemi - Proje Hafızası ve Kuralları

Bu doküman, projenin mimari yapısını, pazaryeri sipariş formatlarını, iş kurallarını ve geliştirme standartlarını kalıcı olarak hafızaya kaydeder.

---

## 1. Proje Mimarisi ve Çalışma Prensibi

* **Web Arayüzü & Sunucu:** `app.py` (Flask tabanlı, 127.0.0.1:5000 portunda çalışır).
* **Çekirdek Motor:** `core_engine.py` (Tüm Excel okuma, temizleme, marka tespiti, Beko ayrıştırma ve A4 oluşturma mantığını barındırır).
* **Konsol / CLI:** `process_orders.py` ve `order_processor.py`.
* **Canlı Bulut Yayını (7/24 Aktif):** [kozan-rtki.vercel.app](https://kozan-rtki.vercel.app) (Vercel + GitHub `gkozan1997/KOZAN` entegrasyonu ile cep telefonu ve her yerden erişilebilir).
* **Başlatıcılar:**
  * Proje içindeki `baslat.bat`.
  * Masaüstündeki `Siparis_Toplama_Programi.bat`.

---

## 2. Çoklu Excel Dosyası Yükleme ve Çapraz Konsolidasyon

* **Çoklu Yükleme:**
  * Kullanıcı tek seferde birden fazla `.xls` veya `.xlsx` dosyasını sürükleyip bırakabilir veya seçebilir (`<input type="file" multiple>`).
  * Backend (`/api/upload`) tüm dosyaları tek seferde alır ve `parse_and_process_multiple_files` fonksiyonuna iletir.
  * İndirilenler klasöründeki güncel sipariş dosyaları tek tıkla `/api/process-all-recent` ile birleştirilebilir.
* **Mükerrer Birleştirme (Akıllı Konsolidasyon):**
  * **Stok Kodu (SKU) Öncelikli Eşleştirme:** Pazaryerleri aynı ürüne farklı ekler eklese dahi (örn. Hepsiburada'nın `[SSD Kapasitesi:..., Ram (...):...]` veya Trendyol'un `[Renk:Beyaz]`, `, one size` gibi etiketleri), aynı marka ve geçerli bir Stok Kodu taşıyan tüm siparişler otomatik olarak **tek bir satırda toplanır ve adetleri birleştirilir**.
  * **Akıllı İsim Temizleme (`clean_product_name`):** Köşeli parantezli pazar yeri varyant etiketleri temizlenir.
  * **En Uygun Başlık Seçimi (`pick_best_name`):** Farklı platformlardan gelen varyasyonlar arasından marka adını içeren, anlamsız sistem kodları taşımayan en temiz ve okunaklı ürün başlığı seçilir.
  * **Stok Kodu Olmayanlar:** Temizlenmiş ve normalize edilmiş ürün adı üzerinden birleştirilir.
* **Geçici Dosya Çakışması Koruması:** Yüklenen her dosya geçici klasöre **benzersiz ön ekli** adla yazılır (`{secrets.token_hex(4)}_{filename}` — `core_engine.py`, `app.py` oliz rotaları). Böylece farklı pazaryerlerden gelen **aynı adlı** dosyalar (örn. iki adet `Siparis.xlsx`) birbirini sessizce ezmez.

---

## 3. Dinamik Sütun ve Başlık Tespiti

Pazaryeri ve entegrasyon dosyalarında başlıklar değişkenlik gösterebilir:
* **Başlık Satırı:** İlk 5 satır taranır; yasal uyarı (Trendyol disclaimer) olsa dahi asıl başlık satırı otomatik bulunur.
* **Sütun Eşleşmeleri (`find_column_index`):**
  * **Ürün Adı:** `Ürün`, `Ürün Adı`, `Urun Adi`, `Urun`, `Ürün İsmi`, `Product Name`.
  * **Stok Kodu:** `Stok Kodu`, `Ürün Kodu`, `Urun Kodu`, `Stok`, `SKU`, `Model Kodu`.
  * **Miktar:** `Adet`, `Miktar`, `Miktar (Adet)`, `Sipariş Adedi`, `Quantity`, `Qty`.
  * **Varyant:** `Varyant`, `Variant`, `Seçenek`, `Özellik`. (Eğer varyant ürün adında zaten geçiyorsa mükerrer eklenmez).
  * **Sipariş No:** `Sipariş No`, `Sipariş Numarası`, `Siparis No`, `Paket No`.
  * **Barkod:** `Gtin (Barkod)`, `Barkod`, `GTIN`, `Barcode`, `EAN`.
  * **Mağaza:** `Mağaza`, `Magaza`, `Satıcı`, `Store`.
  * **Pazaryeri:** `Pazaryeri`, `Pazar Yeri`, `Platform`, `Marketplace`.

---

## 4. Marka Tespiti ve Beko Entegrasyonu

* **Mağazalar:**
  * `Ahidenal`: Küçük ev aletleri, elektronik (Tefal, Babyliss, Philips, Braun, Lenovo, Bissell, Teka, Laurastar, Ariete, iPhone vb.).
  * `Bingöl Ticaret - Beko` & `Bingöl Ticaret`: Beko bayisi siparişleri.
* **Beko Tanıma Kriterleri:**
  * Ürün adı veya ilk kelimesi `BEKO` içeriyorsa,
  * Mağaza adı `beko` içeriyorsa (`Bingöl Ticaret - Beko`),
  * Veya ürün adı doğrudan Beko model kodlarıyla başlıyorsa (`KMX`, `7053MB`, `CM `, `CMX`, `B 600`, `B 710`, `BKK`, `BK RHC`, `BFC`, `BDE`, `TKM`, `BEU`, `KMB`, `9704`, `9705`, `31825`, `74826`, `FR 8374`, `FRA ` vb.) ya da net Beko model regex kalıbı taşıyorsa otomatik olarak **BEKO** markası olarak tanımlanır. (Not: Metin içindeki santimetre "CM" veya renk kodu "BK" gibi kısaltmalar ya da diğer markalar asla Beko ile karıştırılmaz).
* **Beko, Grundig, Lenovo ve Sony Listeleme Davranışı (2. Sayfa ve 2. Sekme Standardı):**
  * **IdeaPad ve PlayStation Tanıma Kriteri:**
    - Ürün adında veya marka sütununda `IdeaPad` geçen veya `IdeaPad` ile başlayan tüm ürünler (başında Lenovo yazmasa dahi, örn: 'IdeaPad Slim 3...') otomatik olarak **LENOVO** markası olarak tanımlanır ve doğrudan **2. Sayfaya** basılır.
    - Ürün adında veya markasında `Sony`, `PlayStation`, `PS5`, `PS4` geçen ürünler otomatik olarak **SONY** markası olarak tanımlanır ve doğrudan **2. Sayfaya** basılır.
  * **Genel Liste (1. Sekme - `Ürün Toplama Listesi`):**
    - İlk olarak diğer tüm markalar (Delonghi, Thor, Tefal, Babyliss, Philips, Braun, Teka, Bissell vb.) listelenir.
    - Diğer markaların bittiği satıra dikey sayfa sonu (`Break(id=other_items_count + 1)`) eklenir; böylece **BEKO, GRUNDIG, LENOVO (IdeaPad dahil) ve SONY (PlayStation dahil) markalı ürünler doğrudan 2. Sayfaya** basılır.
    - 2. sayfada da `Miktar`, `Ürün Adı`, `Stok Kodu` başlıklarının en üstte tekrarlanması için `ws.print_title_rows = '1:1'` ve `fitToHeight = 2` kullanılır.
  * **Excel Çalışma Kitabı Sekmeleri:**
    - 1. Sekme: `Ürün Toplama Listesi` (Sayfa 1: Diğer Markalar, Sayfa 2: Beko, Grundig, Lenovo & Sony).
    - 2. Sekme: Özel **`BEKO, GRUNDIG, LENOVO & SONY`** sekmesi (tüm Beko, Grundig, Lenovo ve Sony siparişleri burada toplanır).
    - 3+ Sekmeler: Diğer markaların alfabetik/adet sıralı özel sekmeleri (`TEFAL`, `BABYLISS`, `DELONGHI`, `THOR`, `PHILIPS` vb.).
  * **Web Arayüzü Sekmeleri:**
    - 1. Sekme: `Tümü (Ana Liste)`
    - 2. Sekme: `🛡️ BEKO, GRUNDIG, LENOVO & SONY (X Adet)`
    - 3+ Sekmeler: `DELONGHI`, `THOR`, `TEFAL`, `BABYLISS`, vb.
  * **Bağımsız Dosya:** `Beko_Urun_Toplama_Listesi_A4.xlsx` eş zamanlı olarak tüm Beko, Grundig, Lenovo & Sony siparişlerini bağımsız A4 listesi olarak üretir.
  * **Seçenekli Ayrıştırma (`filter_beko=True`):** Kullanıcı arayüzdeki "Beko, Grundig, Lenovo & Sony siparişlerini ana listeden ayır" onay kutusunu işaretlerse, bu ürünler ana listeden çıkarılır ve yalnızca bağımsız listede yer alır.
* **Diğer Markalar (1. Sayfa Standardı):**
  * `DELONGHI`, `THOR`, `TEFAL`, `BABYLISS`, `BISSELL`, `PHILIPS`, `BRAUN`, `WMF`, `KENWOOD`, `ARIETE`, `LAURASTAR`, `TEKA`, `IPHONE/APPLE`, `FAKIR`, `ARZUM`, `KARACA`, `KORKMAZ`, `NESPRESSO`, `KRUPS`, `MELITTA`, `SAGE`, `SIMFER`, `KUMTEL`, `LUXELL`, `SINBO`, `KARCHER`, `ROWENTA` vb.
  * Delonghi, Thor ve benzeri tüm markalar doğrudan **1. Sayfaya** basılır. Asla 2. Sayfaya (Beko, Grundig, Lenovo & Sony) kaymaz.
  * Her marka için Excel çalışma kitabında otomatik olarak özel A4 sekmesi oluşturulur.
* **Büyük Beyaz Eşya, Ankastre, TV ve Garanti Filtrelemesi (`is_major_appliance_or_warranty` + `is_conditional_appliance`):**
  * Gerçek filtreleme `consolidate_and_build` (`core_engine.py`) içindeki `is_item_excluded()` ile yapılır; `is_major_appliance_or_warranty` her satır için doğrudan uygulanır; `is_conditional_appliance` ise **Beko Ocaklar** ve **Beko ADP 61420 Davlumbaz serisi** için uygulanır (müşteri tek aldıysa listeye eklenir, yanında büyük eşya/ürün varsa elenir).
  * Bu filtreleme **YALNIZCA BEKO markalı ürünlere** uygulanır. Diğer tüm markalar (Lenovo, Tefal, Babyliss, Philips, Braun, Teka, Bissell, Laurastar, WMF, Kenwood vb.) doğrudan listelenir.
  * **Beko İçin Filtrelenen Ürünler (Listeden Hariç Tutulanlar):**
    - Davlumbaz / Duvar Tipi Davlumbaz (Beko ADE 62540 B, BDE 6062 G vb. - ADP 61420 tek başına alımları hariç)
    - Aspiratör / Ankastre Sürgülü Aspiratör (Beko P 38, P 41, P 27 vb.)
    - Çamaşır Kurutma Makinesi, Kurutma Makinesi
    - Bulaşık Makinesi
    - Buzdolabı (Mini Buzdolabı dahil)
    - Ankastre Fırın, Mini Fırın, Solo Fırın, Buhar Destekli Fırın, Ocaklı Fırın
    - Çamaşır Makinesi
    - Ek Garanti, Garanti Uzatma Paketleri (Beko vb. ek garantiler - Lenovo Garanti hariç)
    - TV / Televizyon (QLED, OLED, Smart LED vb.)
    - Derin Dondurucu (Çekmeceli, Sandık tipi vb.)
    - Klima (Split, Inverter, Salon tipi vb.)
    - Termosifon, Şofben, Kombi, Boyler ve Ani Su Isıtıcıları (Beko BKT 500 E BS Dijital Termosifon vb. büyük su ısıtma ve ısıtma cihazları)
    - Yazarkasa / POS Cihazları (Beko X30 TR Yazarkasa POS, 300 TR vb. mali cihazlar)
  * **Listede KORUNAN (Hariç Tutulmayan) İstisnalar:**
    - Beko Ocaklar (Beko BOCD, BOMD, BOI, Cam Tablalı Ocak, Gazlı Ocak vb. - Müşteri siparişte tek başına aldıysa listeye eklenir; yanında büyük eşyalar varsa elenir)
    - Beko ADP61420S Duvar Tipi Davlumbaz ve Diğer Renkleri (Beko ADP 61420 S / B / W / G vb. - Müşteri siparişte tek başına aldıysa listeye eklenir; yanında büyük ürünler varsa elenir)
    - Lenovo Garanti ve Bilgisayarlar (Lenovo 1 Yıl Garanti Uzatma Paketi, IdeaPad vb.)
    - Mikrodalga Fırın (Beko BMD vb.)
    - Su Isıtıcı / Su Isıtıcısı (Kettle - Mutfak tipi tezgah üstü kettle cihazları küçük ev aletidir, listede toplanmaya devam eder)
    - Saç Kurutma Makinesi (BaByliss, Grundig vb.)
    - Tüm Küçük Ev Aletleri (Kahve/Çay Makinesi, Blender, Ütü, Fritöz, Süpürge, Tost Makinesi vb.), Telefon ve Bilgisayarlar.

---

## 5. Çıktı Standartları (A4 Dikey)

* **Sayfa Düzeni ve A4 Sığdırma Standardı:**
  * **Genel Liste (1. Sekme):** Toplam **tam 2 sayfa A4**. `fitToWidth=1`, `fitToHeight=2` ve aradaki `Break` sayesinde:
    - **1. Sayfa:** Diğer tüm markalar tam 1 sayfaya sığar.
    - **2. Sayfa:** Beko, Grundig & Lenovo ürünleri tam 2. sayfaya sığar.
    - `ws.print_title_rows = '1:1'` ile her sayfada başlık satırı tekrarlanır.
  * **Tek Sayfalık Sekmeler (2+ Sekmeler ve Beko Dosyası):** `fitToWidth=1`, `fitToHeight=1` ile her marka sekmesi tam 1 sayfaya sığdırılır.
  * **Dinamik Satır ve Font Optimizasyonu:** Sayfa başına düşen kalem > 35 ise satır yüksekliği 20pt, Segoe UI 9.5pt font ve kompakt kenar boşlukları (0.35/0.4 inç) kullanılır; <= 35 kalem için 24pt satır yüksekliği ve 10pt font kullanılır.
  * **Web A4 Yazdır Standardı (`@media print`):** `thead { display: table-header-group; }` ve `tr.print-page-break { break-before: page; }` kuralları ile tarayıcıdan A4 yazdırıldığında da 1. sayfa diğer markaları, 2. sayfa Beko, Grundig & Lenovo ürünlerini tam 2 A4 sayfasına sığdırır.
  * **TEK `@page` Kuralı (Zorunlu):** Proje genelinde **yalnızca bir** `@page` tanımı vardır ve bu **JS tarafından dinamik olarak** `<style id="dynamic-page-rule">` içine yazılır: `applyPrintPageSize(tab)` (`templates/index.html`). Sebep: CSS'te `@page` hiçbir seçiciyle **kapsanamaz** (`body.active-tab-barkod` görünemez) ve aynı özgüllükte yazan kurallardan **kaynakta en son yazan kazanır**. Bu yüzden `index.html`, `barcode_style.css` ve `oliz_style.css` içinde `@page` yazmak **yasaktır**. Değerler: `siparis`/`oliz` → `A4 portrait, 8mm 10mm`; `barkod` → `100mm 100mm, margin 0`.
* **İçeriğe Göre Otomatik Sütun Genişliği & Hafif Gri Zebra:**
  1. `Miktar` (İçeriğe göre dinamik genişlik, ortalı, kalın font)
  2. `Ürün Adı` (Kalan genişliğin tamamı, sola dayalı, kelime kaydırma aktif)
  3. `Stok Kodu` (İçeriğe göre dinamik genişlik, ortalı, monospace/koyu font)
* **Stil:** Koyu başlık satırı (`#1E293B`, beyaz yazı), satırların biri diğerine göre hafif gri zebra desenli (`#F1F5F9`), ince kenarlıklar.
* **Dosyalar:**
  1. `Urun_Toplama_Listesi_A4_Cikti.xlsx` (Ana Liste + her marka için ayrı sekme).
  2. `Beko_Urun_Toplama_Listesi_A4.xlsx` (Beko & Grundig ürünleri toplama listesi).
* **Kayıt Konumları:**
  * `projem/YYYY-AA-GG/`
  * `Desktop/Siparis/YYYY-AA-GG/`
  * `Downloads/YYYY-AA-GG/`
* **Açık Dosya Koruması (`save_workbook_safely`):** Eğer Excel dosyası personelde açıksa program çökmez, otomatik olarak `_1.xlsx`, `_2.xlsx` adıyla kaydeder.

---

## 6. Port ve Süreç Yönetimi

* Flask varsayılan olarak `127.0.0.1:5000` portunu kullanır.
* Olası donma veya eski açık kalmış `SiparisToplama.exe` süreçlerinin portu kilitlemesini önlemek için:
  * Tüm `.bat` başlatıcıları açılışta 5000 portunu dinleyen süreçleri otomatik temizler (`taskkill /F /PID %%a`).

---

## 7. Sayfa İmzası & Kapanış Metni Durumu

* Kullanıcı talebi doğrultusunda hem web arayüzünden (ekran tablosu ve `@media print` A4 baskı önizlemesi), hem de Excel çıktılarından (`core_engine.py`) **"Saygılarımla,"** yazısı ve **Gökhan B.** imza görseli tamamen kaldırılmıştır.
* A4 baskı ve Excel listeleri sade, temiz ve yalnızca sipariş toplama kalemlerine odaklı olarak üretilmektedir.
* Web arayüzünden 4 adet 46KB'lık gömülü base64 veri bloğu temizlenerek sayfa yüklenme hızı ve kaynak verimliliği maksimuma çıkarılmıştır.

---

## 8. Otomatik GitHub & Vercel Senkronizasyonu Standardı

* **Zorunlu Güncelleme Kuralı:** Projede yapılan her yenilik, kural güncellemesi, özellik eklemesi veya hata düzeltmesi sonrasında proje MUTLAKA GitHub'a commit ve push (`git push origin main`) edilir.
* **Canlı Sistem Senkronizasyonu:** Bu sayede [kozan-rtki.vercel.app](https://kozan-rtki.vercel.app) üzerindeki canlı bulut yayını masaüstü sistemiyle daima %100 senkronize ve güncel kalır.

---

## 9. 1. Sayfa Özel Not Kutusu Standardı

* **Özel Not Girişi (`page1-note-input`):**
  * Web arayüzünde tablo araçlarının üzerinde yer alan özel not kutusu (`.page1-note-card`) sayesinde kullanıcı serbest metin notu girebilir.
  * Not metni `localStorage` ile tarayıcıda kalıcı olarak saklanır ve tek tıkla temizlenebilir (`clearPage1Note`).
* **1. Sayfa Sonuna Eklenme Davranışı:**
  * Girilen not, Ana Toplama Listesinde (1. Sekme) **yalnızca 1. sayfa listesinin en alt sırasına** (`<tr>` olarak) eklenir; 2. sayfaya (Beko, Grundig & Lenovo) asla taşmaz veya sirayet etmez.
  * Miktar sütununda belirgin **`NOT`** rozeti (`.note-badge`), Ürün Adı ve Stok Kodu sütunlarını kapsayan birleşik alanda (`colspan="2"`) **`Not: [Metin]`** şeklinde basılır.
  * Hem web ekranında hafif sıcak vurguyla, hem de A4 baskı ve PDF çıktısında (`@media print`) net kenarlıklar ve Segoe UI font standardıyla kusursuz basılır.
  * Excel çalışma kitabında da (`core_engine.py`) 1. sekmenin 1. sayfa sonuna (sayfa sonu kırılımından hemen önce) otomatik eklenir ve B-C sütunları birleştirilir.

---

## 10. Ürün Ailesi ve Renk/Varyant Sıralama Standardı (Alt Alta Gruplama)

* **Aynı Modelin Renk Varyantları:** Farklı renk veya varyanta sahip aynı ürün/model kalemleri (örnek: `Beko CM 5964 R Floral Çay Makinesi`, `Beko CM 5964 B Floral Çay Makinesi`, `Beko TKM 2341 Keyf-i Bol Beyaz / Siyah`, `Beko BMD 200 B / G / S Mikrodalga Fırın`) araya başka ürün girmeden **doğrudan alt alta** basılır.
* **Akıllı Ürün Ailesi & Renk Ayrıştırma (`extract_product_sort_keys`):**
  * Ürün başlığındaki model numarası ve renk eki (R, B, S, G, I, M, K, EB, TB, DS vb.) veya Türkçe/İngilizce renk kelimeleri (Beyaz, Siyah, Kırmızı, Mavi, Gri vb.) ayrıştırılarak ortak `base` gövde anahtarı üretilir.
  * Ölçü birimleri (W, V, Kg, Gr, Cm, Mm, Gb, Tb, Btu vb.) renk ekiyle karıştırılmaz.
* **Grup İçi & Genel Sıralama (`sort_items_by_family_and_color`):**
  * Ürün ailesi içindeki varyantlar kendi içinde renk koduna göre sıralanır.
  * Ürün aileleri ise depoda toplama kolaylığı sağlamak adına öncelikle en yüksek sipariş adedine (`max_qty`), adetler eşitse ürün adına göre alfabetik olarak sıralanır.
  * Bu sıralama hem Web canlı ekranında ve A4 baskıda (`grouped_items`), hem Excel genel listesinde, hem Beko bağımsız dosyasında hem de her markanın kendi sekmesinde geçerlidir.

---

## 11. Toplu Barkod & Kargo Etiketi Modülü Entegrasyonu

* **Toplu Barkod Çekirdeği (`core_barcode.py`):**
  - Excel/CSV otomatik sütun algılama (Alıcı, Adres, İlçe, İl, Telefon, Barkod/Takip No, Ürün Adı, Adet).
  - Manuel dinamik sütun eşleştirme çubuğu (`#mapping-bar`).
  - JsBarcode ile istemci tarafında SVG Code128 canlı etiket önizleme.
  - ReportLab ile 100mm x 100mm yüksek kaliteli PDF üretimi (`/api/barcode/generate-pdf`).
* **İzole Çift Baskı Standardı (`@media print`):**
  - Sipariş Toplama sekmesinde iken standart A4 dikey baskı kuralları çalışır.
  - Barkod sekmesinde iken `body.active-tab-barkod` devreye girerek `@page { size: 100mm 100mm; margin: 0; }` termal yazıcı formatında sayfa sayfa baskı verir. İki baskı sistemi asla birbirine karışmaz. (Bu `@page` kuralı da §5'teki tek dinamik kuraldan enjekte edilir.)
* **Oturum Bazlı Veri Önbelleği (`_barcode_cache_get/set`):**
  - Yüklenen kargo dosyasının başlık/satır/eşleme verisi **global dict yerine** `barcode_sid` çereziyle oturuma bağlı `BARCODE_SESSION_CACHE` sözlüğünde tutulur (TTL 1 saat, en fazla 20 oturum).
  - Amaç: Bulutta eşzamanlı iki kullanıcının birbirinin müşteri/adres/telefon verisini görmesini engellemek. `/api/barcode/generate-pdf` yalnızca **kendi oturumunun** önbelleğine düşebilir.

---

## 12. Sol Menü Dashboard Mimarisi & 3 Sayfa Entegrasyonu (Sol Menü Standardı)

* **Sol Menü (Sidebar) Navigasyon Paneli (`aside.app-sidebar`):**
  - Tüm platformun sol tarafında sabit (sticky), modern cam efektli (`rgba(11, 16, 28, 0.96)`) dikey sidebar yer alır.
  - Sol tarafta sıralanan **3 Ana Modül Sayfası:**
    1. **`📦 Sipariş & Toplama` (`#side-nav-siparis`):** Pazaryeri sipariş konsolidasyonu, Beko/Grundig/Lenovo 2. sayfa standardı, özel not ve tam A4 dikey baskı.
    2. **`🏷️ Toplu Barkod & Kargo` (`#side-nav-barkod`):** 10x10cm termal rulo etiket, otomatik sütun eşleme, SVG barkod ve ReportLab PDF.
    3. **`🎁 Oliz Kampanya` (`#side-nav-oliz`):** Beko & Arçelik 1, 2, 3, 4 ürünlü sepet analizi, paket fırsatları, toptan kuponlar ve katılım payları.
* **Dinamik Geçiş ve URL Senkronizasyonu (`switchMainTab`):**
  - Sayfa yenilenmeden tek tıkla sekmeler arası geçiş yapılır.
  - URL hash `#siparis`, `#barkod`, `#oliz` otomatik güncellenir.
  - Bağımsız rotalar: `/barkod` ve `/oliz` doğrudan ilgili sayfayı açar.
* **Baskı İzolasyonu:** `@media print` sırasında sol menü gizlenir (`display: none !important`), baskı alanı sayfa genişliğini tam doldurur.

---

## 13. Mobil Uyumluluk ve Responsive Sistem Standardı (Mobile-First Dashboard Standardı)

* **Mobil Navigasyon Mimarisi (`<= 768px`):**
  - **Masaüstü Sidebar İzolasyonu:** Geniş sol menü (`.app-sidebar`) mobil ekranlarda gizlenir (`display: none !important;`) ve ekranda dikey yer işgal etmesi önlenir.
  - **Kompakt Mobil Üst Çubuk (`.mobile-top-bar`):** Ekranın en üstünde yapışkan (sticky), modern cam efektli logo, çevrimiçi göstergesi ve anlık aktif modülü belirten dinamik rozet (`#mobile-module-badge`) yer alır.
  - **Mobil Üst Sekme Kartları (`.mobile-tabs-container`):** Ekranın üst kısmında ana görünümün hemen başında 3 ana modülü (Sipariş & Toplama [A4], Toplu Barkod [10x10], Oliz Kampanya [Fırsat]) gösteren, renk vurgulu, dokunmatik 3'lü kart switcher konumlanır.
  - **Mobil Sabit Alt Menü Dok'u (`.mobile-bottom-nav`):** Ekranın en altında sabit (fixed), başparmakla tek dokunuşla erişilebilir 3 ana modül butonu (`Sipariş`, `Barkod`, `Oliz`) konumlanır. CSS özgüllük (specificity) çakışması önlenerek masaüstü (`min-width: 769px`) ve mobil (`max-width: 768px`) kuralları tam izole edilmiştir.
  - **Çok Yönlü Durum Senkronizasyonu (`switchMainTab`):** Sekme geçişlerinde masaüstü sol menü, mobil üst sekme kartları, mobil alt dock ve URL hash'i eş zamanlı aktifleşir.
  - **Güvenli Alan (Safe Area Inset):** iPhone çentiği ve alt ev çubuğu için `viewport-fit=cover` ve `padding-bottom: max(4px, env(safe-area-inset-bottom))` tam uyumludur.
* **Modül Bazlı Mobil Uyarlamalar:**
  - **Sipariş & Toplama Modülü (Mobil Tablo Standardı):**
    - Çift sarmal container temizlendi; padding mobil ekranlar için optimize edildi.
    - KPI özet kartları kompakt 2x2 grid yapısına dönüştürüldü.
    - **`table-layout: fixed !important` Standardı:** 1. Tablo (Diğer Markalar) ile 2. Tablo (Beko, Grundig & Lenovo) kolonları mobilde `50px` (Miktar), `auto` (Ürün Adı) ve `96px` (Stok Kodu) olarak kilitlendi; dikey sütun çizgilerinin kayması ve hizasızlıklar tamamen önlendi.
    - **Stok Kodu Okunabilirlik Standardı:** `word-break: break-all` kaldırıldı; stok kodlarının harf ve rakamlarının dikeyde bölünmesi engellendi, `code-badge` içinde tek parça ve okunabilir monospace format sağlandı.
    - **Dokunmatik Toplama Kontrolü (`toggleRowCollected`):** Depo personeli telefonda ürünü aldığında satıra dokunarak yeşil vurgu ve üstü çizili olarak işaretleyebilir; baskıda bu durum asla görünmez.
    - **Baskı İzolasyonu (`@media screen and (max-width: 768px)`):** Mobil ekran kuralları `@media print` A4 dikey baskı motoruna asla sızmaz, cep telefonundan dahi yazdırılsa tam 2 sayfa A4 standardı kusursuz korunur.
    - **Mobil A4 Yazdırma & Boş Sayfa İzolasyonu:** `.print-page-block` üzerindeki `break-inside: avoid` kaldırıldı; tarayıcının 1. sayfayı boş geçerek içeriği 2. sayfaya fırlatması engellendi. CSS'in en sonuna eklenen nihai `@media print` bloğu ile `.mobile-bottom-nav` ve `.mobile-top-bar` yazdırılan A4 çıktısından kesin olarak silindi; siyah mobil barın tablonun ortasına yapışması tamamen önlendi.
    - Araç çubuğu (toolbar) butonları 2 sütunlu dokunmatik grid olarak düzenlendi.
    - Bildirim pencereleri (toast) alt menü dok'unun hemen üzerinde görüntülenecek şekilde konumlandırıldı.
  - **Toplu Barkod & Kargo Modülü:**
    - Üst aksiyon butonları ve sütun eşleme seçicileri 2 sütunlu dokunmatik ızgaraya dönüştürüldü.
    - 10cm x 10cm etiket önizleme kartları mobil ekranlarda tekli tam kart genişliğinde ortalanarak kusursuz kare oranını ve SVG netliğini korur.
  - **Oliz Kampanya Modülü:**
    - İstatistik kartları 2x2 mobil grid formatına uyarlandı.
    - Hızlı deneme butonları yatay kaydırılabilir dokunmatik çip şeridi haline getirildi.
    - 4 ürün giriş kutusu ve aksiyon butonları dikey akışa optimize edildi.
* **Masaüstü ve Baskı Koruma Garantisi:**
  - Masaüstü görünüm (`> 768px` ve laptop ekranları dahil): `.app-sidebar` her zaman sol tarafta dikey (sticky, 260px genişlik) olarak kalır, asla yatay çubuğa dönüşmez. `.mobile-top-bar` ve `.mobile-bottom-nav` masaüstünde daima gizlidir (`display: none !important;`).
  - A4 / 100x100mm termal baskı kuralları (`@media print`) bu geliştirmelerden bağımsız olarak %100 korunmaktadır.

---

## 14. KOD & GÖRSEL DENETİM RAPORU VE DÜZELTMELERİ (2026-10-03 / 2026-10-04)

Statik + canlı denetimde tespit edilen **19 doğrulanmış hata** bulunmuş, tamamı 2026-10-04 tarihinde düzeltilmiş ve regresyon testleriyle kanıtlanmıştır. **Tüm maddeler "açık hata" olmaktan çıkarılmıştır.** Aşağıda her maddenin kök nedeni ve uygulanan çözüm sabitlenmiştir; **bu çözümler geri alınmamalıdır.**

### 14.1 KRİTİK

| # | Hata | Kök Neden | Çözüm |
|---|------|-----------|--------|
| H-01 | Termal barkod baskısı bozuk (`@page` kaskad çakışması) | `index.html`, `barcode_style.css` ve `oliz_style.css` üçünde de **kapsamsız** `@page` vardı; `body.active-tab-*` ile kapsanamayacağı için kaynakta en son yazan (`oliz` → A4) kazanıyordu | **Tek** `@page` kuralı JS ile dinamik enjekte edildi: `PRINT_PAGE_RULES` + `applyPrintPageSize(tab)` → `<style id="dynamic-page-rule">`. Üç CSS/HTML kaynağından `@page` **tamamen kaldırıldı**. `switchMainTab` ve `beforeprint` yöneticisine bağlandı. (bkz. §5) |
| H-02 | Aynı adlı iki dosyada sessiz veri kaybı | Geçici dosya yolu `os.path.join(temp_dir, filename)` idi; iki `Siparis.xlsx` birbirini eziyordu | `parse_single_file_rows` geçici adı `{secrets.token_hex(4)}_{filename}` yapıyor; `oliz_upload_endpoint` da aynı ön eki aldı (bkz. §2) |
| H-03 | Geçersiz marka adı Excel'i çökertiyor (HTTP 500) | `ws.create_sheet(title=brand[:30])`; `ARCELIK*`, `BEKO/ARCELIK`, `BEKO: 5` geçersiz karakter/şema hatası veriyordu | Yeni `safe_sheet_title(title, used_titles)`: `[]:*?/\` temizliği + 31 karakter kısaltma + `_2`, `_3` benzersizlik sırası |
| H-04 | KPI çift sayımı | `filter_beko=False` iken `grouped_items` Beko/Grundig/Lenovo'yu da içeriyor, `non_beko_*` tüm listeyi sayıyordu | `other_page_items = grouped_items[:other_items_count]` üzerinden hesaplanıyor; `non_beko_count == other_items_count` garantisi test edildi |

### 14.2 ORTA

| # | Hata | Çözüm |
|---|------|--------|
| H-05 | Sahte marka = model numarası (`6715DE` sekmesi) | `detect_brand` yedeği artık **rakam içermeyen, ≥3 harfli** ilk kelimeyi kabul eder; aksi halde `DİĞER` döner |
| H-06 | `.xls` / `.xlsx` müşteri sütunu sırası farklıydı | İki dal **tek ve aynı** `['üye adı soyadı', 'fatura - müşteri', …]` listesini kullanır (`Üye Adı Soyadı` öncelikli) |
| H-07 | `Marketplace` başlığı algılanmıyordu | `platform` listesine `marketplace` eklendi (her iki dal) |
| H-08 | `/api/oliz/autocomplete?limit=abc` → 500 | `try/except` ile güvenli `int()` + `1..100` aralığa kırpma (`/api/autocomplete` alias'ı dahil) |
| H-09 | A4 kenar boşluğu dokümanla çelişiyordu | Tek dinamik kural §5'teki standarda sabitlendi: `A4 portrait, 8mm 10mm` |
| H-10 | Barkod önbelleği global → bulutta veri sızıntısı | `BARCODE_SESSION_CACHE` artık `barcode_sid` çerezine bağlı **oturum sözlüğü** (`_barcode_sid`, `_barcode_cache_get/set`, `@app.after_request`). TTL 1 saat, en fazla 20 oturum (bkz. §11) |
| H-11 | Excel çıktısında gridlines basılıyordu | `ws.print_options.gridLines = False` (hücre kenarlıkları zaten çizili) |
| H-12 | `/api/open-folder` yol doğrulaması yoktu | `os.path.isdir` + `_is_allowed_local_path()` kontrolü; geçersiz yol `BASE_DIR`'e düşer |

### 14.3 DÜŞÜK — Ölü Kod / Tutarsızlık

| # | Madde | Çözüm |
|---|-------|--------|
| H-13 | `is_excluded_product()` ölü koddu, §4 yanlış fonksiyonu gösteriyordu | §4 gerçek fonksiyonlara (`is_major_appliance_or_warranty` / `is_conditional_appliance` / `is_item_excluded`) güncellendi |
| H-14 | Kullanılmayan `font_header/font_regular/font_qty/font_code` | `consolidate_and_build` içinden kaldırıldı (yerel `f_*` değişkenleri zaten var) |
| H-15 | Başlık satırı taraması ikiye kopyalanmıştı | Tek `HEADER_HINTS` + `detect_header_row(rows)` yardımcısı; her iki biçim aynı mantığı kullanıyor |
| H-16 | `secure_filename` ikinci kez import edilmişti | Yinelenen import satırı kaldırıldı |
| H-17 | Kullanılmayan import/değişkenler | `core_barcode.datetime`, `except ... as e`, `core_campaign.valid_skus`, `order_processor.os/unicodedata/get_base_dirs`, `process_orders.os` temizlendi |
| H-18 | `.main-tabs-nav` ölü CSS seçicisi | `barcode_style.css` baskı bloğundan kaldırıldı (HTML'de zaten yok) |
| H-19 | Denetim raporu commit edilmemişti | Bu bölüm + tüm düzeltmeler tek commit ile `main` dalına gönderildi (§8) |

### 14.4 Doğrulama Kanıtları (2026-10-04)

* `python -m py_compile` tüm `.py` dosyalarında geçti; `python -m pyflakes` **sıfır** bulgu.
* `_regression_test.py`: masaüstü ve bulut yolu çıktıları **birebir aynı** (11 örnek dosya → 553 sipariş, 45 ana kalem, 34 Beko kalemi, 13 marka).
* Düzeltme doğrulama paketi (25/25 geçti): aynı adlı çift yükleme, geçersiz sekme adları, KPI toplamları, sahte marka engeli, `Marketplace` algılama, gridlines kapalılığı, Beko'ya özel filtreleme.
* API & baskı doğrulama paketi (24/24 geçti): `limit=abc/-5/99999` → 200, açık klasör yol sızıntısı engeli, **çok kullanıcılı barkod oturum izolasyonu** (B oturumu A'nın verisini göremiyor, PDF üretemiyor), tek `@page` kuralı, `/` `/barkod` `/oliz` → 200.
* `_api_test.py` (canlı sunucu): `/api/upload` 200, disk çıktısı yazıldı, yol sızma denemesi 404.

### 14.5 Denetimde DOĞRU ÇIKAN Noktalar (regresyon koruması)

* Beko büyük eşya filtresi **yalnızca BEKO'ya** uygulanıyor: `Grundig 40 Klasik LED TV` ve `Lenovo 1 Yıl Garanti Uzatma Paketi` listede **korunuyor** (§4 ile uyumlu).
* Aynı siparişteki `Beko B 710 Bulaşık Makinesi` + `Beko P 38 Aspiratör` doğru elendi (`excluded_qty=2`).
* Renk varyantı gruplama (`CM 5964 R` / `CM 5964 B` alt alta) çalışıyor; ölçü birimleri (`W`, `V`, `Kg`) renk kodundan ayrılıyor.
* Excel sayfa sonu kırılımı doğru: `Break(id=…)`, `fitToHeight=2`, `print_title_rows='$1:$1'`.
* `templates/index.html`: tekrar eden `id` yok; `<style>`/`<script>` bloklarında süslü parantez dengesi tam.
* Font dosyaları Vercel'e gönderiliyor (`fonts/arial.ttf` + `arialbd.ttf`) → bulutta Türkçe karakter kaybı yok.


---

## 15. Her Sayfada Liste Oluşturulma Tarihi Standardı (Sol Üst / Miktar Sütunu Üzeri)

* **Konum ve Yerleşim:**
  - Ana Toplama Listesinde (Sayfa 1: Diğer Markalar, Sayfa 2: Beko, Grundig, Lenovo & Sony) ve tekil marka sekmelerinde, tablonun sol üst köşesinde, doğrudan `Miktar` sütununun hemen üzerinde `.print-page-date` (`#print-page-date-1`, `#print-page-date-2`) yer alır.
  - Ekranda modern ve sade fontla (`11.5px`, `var(--text-muted)`), boşken yer kaplamaz (`:empty { display: none !important; }`).
  - `@media print` A4 dikey baskıda: `display: block !important; font-size: 8.5pt !important; font-weight: 700 !important; color: #1e293b; line-height: 1.1; margin-bottom: 2px;` kompakt yapısıyla A4 dikey 2 sayfa sığdırma düzenini asla bozmaz.
  - Sayfa 2 blok kırılımında (`break-before: page`) 2. sayfanın da en başında otomatik olarak yer alır.
* **Tarih Formatı:**
  - Liste Excel'den aktarılıp oluşturulduğunda güncel tarih (`GG.AA.YYYY`, örn: `06.10.2026`) dinamik olarak basılır (`created_date`).
* **Excel Çıktısı Entegrasyonu:**
  - `core_engine.py` içindeki `setup_a4_sheet` fonksiyonunda `ws.oddHeader.left.text = display_date` olarak atanır; böylece Excel'den yazdırıldığında da sayfanın sol üst köşesinde liste tarihi görüntülenir ve 1. satır tablo başlığı (`Miktar`, `Ürün Adı`, `Stok Kodu`) yapısı korunur.
