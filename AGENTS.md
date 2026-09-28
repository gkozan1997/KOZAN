# GökhanOS E-Ticaret Depo & Sipariş Toplama Sistemi - Proje Hafızası ve Kuralları

Bu doküman, projenin mimari yapısını, pazaryeri sipariş formatlarını, iş kurallarını ve geliştirme standartlarını kalıcı olarak hafızaya kaydeder.

---

## 1. Proje Mimarisi ve Çalışma Prensibi

* **Web Arayüzü & Sunucu:** `app.py` (Flask tabanlı, 127.0.0.1:5000 portunda çalışır).
* **Çekirdek Motor:** `core_engine.py` (Tüm Excel okuma, temizleme, marka tespiti, Beko ayrıştırma ve A4 oluşturma mantığını barındırır).
* **Konsol / CLI:** `process_orders.py` ve `order_processor.py`.
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
  * Veya ürün Beko model kodlarıyla başlıyorsa (`KMX`, `7053MB`, `CM`, `CMX`, `B 600`, `B 710`, `BKK`, `BFC`, `BDE`, `TKM`, `BEU`, `KMB`, `9704`, `9705`, `31825`, `74826`, `FR 8374` vb.) otomatik olarak **BEKO** markası olarak tanımlanır.
* **Beko ve Grundig Listeleme Davranışı (2. Sayfa ve 2. Sekme Standardı):**
  * **Genel Liste (1. Sekme - `Ürün Toplama Listesi`):**
    - İlk olarak diğer tüm markalar (Tefal, Babyliss, Philips, Lenovo vb.) listelenir.
    - Diğer markaların bittiği satıra dikey sayfa sonu (`Break(id=other_items_count + 1)`) eklenir; böylece **BEKO ve GRUNDIG markalı ürünler doğrudan 2. Sayfaya** basılır.
    - 2. sayfada da `Miktar`, `Ürün Adı`, `Stok Kodu` başlıklarının en üstte tekrarlanması için `ws.print_title_rows = '1:1'` ve `fitToHeight = 0` (sıkıştırmadan çok sayfalı dikey yazdırma) kullanılır.
  * **Excel Çalışma Kitabı Sekmeleri:**
    - 1. Sekme: `Ürün Toplama Listesi` (Sayfa 1: Diğer Markalar, Sayfa 2: Beko & Grundig).
    - 2. Sekme: Özel **`BEKO & GRUNDIG`** sekmesi (tüm Beko ve Grundig siparişleri burada toplanır).
    - 3+ Sekmeler: Diğer markaların alfabetik/adet sıralı özel sekmeleri (`TEFAL`, `BABYLISS`, `PHILIPS`, `LENOVO` vb.).
  * **Web Arayüzü Sekmeleri:**
    - 1. Sekme: `Tümü (Ana Liste)`
    - 2. Sekme: `🛡️ BEKO & GRUNDIG (X Adet)`
    - 3+ Sekmeler: `TEFAL`, `BABYLISS`, vb.
  * **Bağımsız Dosya:** `Beko_Urun_Toplama_Listesi_A4.xlsx` eş zamanlı olarak tüm Beko & Grundig siparişlerini bağımsız A4 listesi olarak üretir.
  * **Seçenekli Ayrıştırma (`filter_beko=True`):** Kullanıcı arayüzdeki "Beko & Grundig siparişlerini ana listeden ayır" onay kutusunu işaretlerse, Beko ve Grundig ürünleri ana listeden çıkarılır ve yalnızca bağımsız listede yer alır.
* **Diğer Markalar:**
  * `TEFAL`, `BABYLISS`, `LENOVO`, `BISSELL`, `PHILIPS`, `BRAUN`, `WMF`, `KENWOOD`, `ARIETE`, `LAURASTAR`, `TEKA`, `IPHONE/APPLE`, `FAKIR`, `ARZUM`, `KARACA`, `KORKMAZ` vb.
  * Her marka için Excel çalışma kitabında otomatik olarak özel A4 sekmesi oluşturulur.
* **Büyük Beyaz Eşya, TV ve Garanti Filtrelemesi (`is_excluded_product`):**
  * Bu filtreleme **YALNIZCA BEKO markalı ürünlere** uygulanır. Diğer tüm markalar (Tefal, Babyliss, Philips, Braun, Lenovo, Teka, Bissell, Laurastar, WMF, Kenwood vb.) doğrudan listelenir.
  * **Beko İçin Filtrelenen Ürünler (Listeden Hariç Tutulanlar):**
    - Çamaşır Kurutma Makinesi, Kurutma Makinesi
    - Bulaşık Makinesi
    - Buzdolabı (Mini Buzdolabı dahil)
    - Ankastre Fırın, Mini Fırın, Solo Fırın, Buhar Destekli Fırın
    - Çamaşır Makinesi
    - Ek Garanti, Garanti Uzatma Paketleri (Beko vb. ek garantiler - Lenovo Garanti hariç)
    - Davlumbaz
    - TV / Televizyon (QLED, OLED, Smart LED vb.)
    - Derin Dondurucu (Çekmeceli, Sandık tipi vb.)
    - Klima (Split, Inverter, Salon tipi vb.)
  * **Listede KORUNAN (Hariç Tutulmayan) İstisnalar:**
    - Lenovo Garanti (Lenovo 1 Yıl Garanti Uzatma Paketi vb.)
    - Mikrodalga Fırın (Beko BMD vb.)
    - Ankastre Ocak / Ocaklar (Beko BOCD, BOI vb. - Tek başına alındıysa toplanır; Çamaşır, Kurutma, Bulaşık, Buzdolabı, Fırın, Davlumbaz, TV, Dondurucu veya Klima ile alındıysa hariç tutulur)
    - Aspiratör / Ankastre Sürgülü Aspiratör (Beko P 38 vb. - Tek başına alındıysa toplanır; büyük beyaz eşyalarla birlikte alındıysa hariç tutulur)
    - Saç Kurutma Makinesi (BaByliss, Grundig vb.)
    - Tüm Küçük Ev Aletleri (Kahve/Çay Makinesi, Blender, Ütü, Fritöz, Süpürge, Tost Makinesi vb.), Telefon ve Bilgisayarlar.

---

## 5. Çıktı Standartları (A4 Dikey)

* **Sayfa Düzeni ve A4 Sığdırma Standardı:**
  * **Genel Liste (1. Sekme):** Toplam **tam 2 sayfa A4**. `fitToWidth=1`, `fitToHeight=2` ve aradaki `Break` sayesinde:
    - **1. Sayfa:** Diğer tüm markalar tam 1 sayfaya sığar.
    - **2. Sayfa:** Beko & Grundig ürünleri tam 2. sayfaya sığar.
    - `ws.print_title_rows = '1:1'` ile her sayfada başlık satırı tekrarlanır.
  * **Tek Sayfalık Sekmeler (2+ Sekmeler ve Beko Dosyası):** `fitToWidth=1`, `fitToHeight=1` ile her marka sekmesi tam 1 sayfaya sığdırılır.
  * **Dinamik Satır ve Font Optimizasyonu:** Sayfa başına düşen kalem > 35 ise satır yüksekliği 20pt, Segoe UI 9.5pt font ve kompakt kenar boşlukları (0.35/0.4 inç) kullanılır; <= 35 kalem için 24pt satır yüksekliği ve 10pt font kullanılır.
  * **Web A4 Yazdır Standardı (`@media print`):** `@page { size: A4 portrait; margin: 8mm 10mm; }`, `thead { display: table-header-group; }` ve `tr.print-page-break { break-before: page; }` kuralları ile tarayıcıdan A4 yazdırıldığında da 1. sayfa diğer markaları, 2. sayfa Beko & Grundig ürünlerini tam 2 A4 sayfasına sığdırır.
* **3 Sütun Yapısı:**
  1. `Miktar` (Genişlik 10-12, ortalı, kalın font)
  2. `Ürün Adı` (Genişlik 56-60, sola dayalı, kelime kaydırma aktif)
  3. `Stok Kodu` (Genişlik 16-18, ortalı, monospace/koyu font)
* **Stil:** Koyu başlık satırı (`#1E293B`, beyaz yazı), zebra desenli satırlar (`#F8FAFC`), ince gri kenarlıklar.
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
