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
  * Derlenmiş bağımsız paket: `dist/SiparisToplama/SiparisToplama.exe` ve `C:\Users\ugurk\OneDrive\Masaüstü\Siparis`.
  * Taşınabilir zip: `C:\Users\ugurk\OneDrive\Masaüstü\Siparis_Toplama_Programi_Paketi.zip`.

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
* **Beko, Grundig ve Lenovo Listeleme Davranışı (2. Sayfa ve 2. Sekme Standardı):**
  * **IdeaPad Tanıma Kriteri:**
    - Ürün adında veya marka sütununda `IdeaPad` geçen veya `IdeaPad` ile başlayan tüm ürünler (başında Lenovo yazmasa dahi, örn: 'IdeaPad Slim 3...') otomatik olarak **LENOVO** markası olarak tanımlanır ve doğrudan **2. Sayfaya** basılır.
  * **Genel Liste (1. Sekme - `Ürün Toplama Listesi`):**
    - İlk olarak diğer tüm markalar (Delonghi, Thor, Tefal, Babyliss, Philips, Braun, Teka, Bissell vb.) listelenir.
    - Diğer markaların bittiği satıra dikey sayfa sonu (`Break(id=other_items_count + 1)`) eklenir; böylece **BEKO, GRUNDIG ve LENOVO (IdeaPad dahil) markalı ürünler doğrudan 2. Sayfaya** basılır.
    - 2. sayfada da `Miktar`, `Ürün Adı`, `Stok Kodu` başlıklarının en üstte tekrarlanması için `ws.print_title_rows = '1:1'` ve `fitToHeight = 2` kullanılır.
  * **Excel Çalışma Kitabı Sekmeleri:**
    - 1. Sekme: `Ürün Toplama Listesi` (Sayfa 1: Diğer Markalar, Sayfa 2: Beko, Grundig & Lenovo).
    - 2. Sekme: Özel **`BEKO, GRUNDIG & LENOVO`** sekmesi (tüm Beko, Grundig ve Lenovo siparişleri burada toplanır).
    - 3+ Sekmeler: Diğer markaların alfabetik/adet sıralı özel sekmeleri (`TEFAL`, `BABYLISS`, `DELONGHI`, `THOR`, `PHILIPS` vb.).
  * **Web Arayüzü Sekmeleri:**
    - 1. Sekme: `Tümü (Ana Liste)`
    - 2. Sekme: `🛡️ BEKO, GRUNDIG & LENOVO (X Adet)`
    - 3+ Sekmeler: `DELONGHI`, `THOR`, `TEFAL`, `BABYLISS`, vb.
  * **Bağımsız Dosya:** `Beko_Urun_Toplama_Listesi_A4.xlsx` eş zamanlı olarak tüm Beko, Grundig & Lenovo siparişlerini bağımsız A4 listesi olarak üretir.
  * **Seçenekli Ayrıştırma (`filter_beko=True`):** Kullanıcı arayüzdeki "Beko, Grundig & Lenovo siparişlerini ana listeden ayır" onay kutusunu işaretlerse, bu ürünler ana listeden çıkarılır ve yalnızca bağımsız listede yer alır.
* **Diğer Markalar (1. Sayfa Standardı):**
  * `DELONGHI`, `THOR`, `TEFAL`, `BABYLISS`, `BISSELL`, `PHILIPS`, `BRAUN`, `WMF`, `KENWOOD`, `ARIETE`, `LAURASTAR`, `TEKA`, `IPHONE/APPLE`, `FAKIR`, `ARZUM`, `KARACA`, `KORKMAZ`, `NESPRESSO`, `KRUPS`, `MELITTA`, `SAGE`, `SIMFER`, `KUMTEL`, `LUXELL`, `SINBO`, `KARCHER`, `ROWENTA` vb.
  * Delonghi, Thor ve benzeri tüm markalar doğrudan **1. Sayfaya** basılır. Asla 2. Sayfaya (Beko, Grundig & Lenovo) kaymaz.
  * Her marka için Excel çalışma kitabında otomatik olarak özel A4 sekmesi oluşturulur.
* **Büyük Beyaz Eşya, TV ve Garanti Filtrelemesi (`is_excluded_product`):**
  * Bu filtreleme **YALNIZCA BEKO markalı ürünlere** uygulanır. Diğer tüm markalar (Lenovo, Tefal, Babyliss, Philips, Braun, Teka, Bissell, Laurastar, WMF, Kenwood vb.) doğrudan listelenir.
  * **Beko İçin Filtrelenen Ürünler (Listeden Hariç Tutulanlar):**
    - Çamaşır Kurutma Makinesi, Kurutma Makinesi
    - Bulaşık Makinesi
    - Buzdolabı (Mini Buzdolabı dahil)
    - Ankastre Fırın, Mini Fırın, Solo Fırın, Buhar Destekli Fırın
    - Çamaşır Makinesi
    - Ek Garanti, Garanti Uzatma Paketleri (Beko vb. ek garantiler - Lenovo Garanti hariç)
    - TV / Televizyon (QLED, OLED, Smart LED vb.)
    - Derin Dondurucu (Çekmeceli, Sandık tipi vb.)
    - Klima (Split, Inverter, Salon tipi vb.)
    - Termosifon, Şofben, Kombi, Boyler ve Ani Su Isıtıcıları (Beko BKT 500 E BS Dijital Termosifon vb. büyük su ısıtma ve ısıtma cihazları)
    - Yazarkasa / POS Cihazları (Beko X30 TR Yazarkasa POS, 300 TR vb. mali cihazlar)
  * **Listede KORUNAN (Hariç Tutulmayan) İstisnalar:**
    - Lenovo Garanti ve Bilgisayarlar (Lenovo 1 Yıl Garanti Uzatma Paketi, IdeaPad vb.)
    - Mikrodalga Fırın (Beko BMD vb.)
    - Davlumbaz / Duvar Tipi Davlumbaz (Beko ADP61420B vb. - Tek başına alındıysa toplanır; Çamaşır, Kurutma, Bulaşık, Buzdolabı, Fırın, TV, Dondurucu, Klima veya Termosifon gibi büyük beyaz eşyalarla birlikte alındıysa hariç tutulur)
    - Ankastre Ocak / Ocaklar (Beko BOCD, BOI vb. - Tek başına alındıysa toplanır; Çamaşır, Kurutma, Bulaşık, Buzdolabı, Fırın, TV, Dondurucu, Klima veya Termosifon ile alındıysa hariç tutulur)
    - Aspiratör / Ankastre Sürgülü Aspiratör (Beko P 38 vb. - Tek başına alındıysa toplanır; büyük beyaz eşyalarla birlikte alındıysa hariç tutulur)
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
  * **Web A4 Yazdır Standardı (`@media print`):** `@page { size: A4 portrait; margin: 8mm 10mm; }`, `thead { display: table-header-group; }` ve `tr.print-page-break { break-before: page; }` kuralları ile tarayıcıdan A4 yazdırıldığında da 1. sayfa diğer markaları, 2. sayfa Beko, Grundig & Lenovo ürünlerini tam 2 A4 sayfasına sığdırır.
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



