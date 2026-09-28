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
* **Mükerrer Birleştirme (Konsolidasyon):**
  * Birden fazla dosyadan gelen aynı ürünler `(Ürün Adı, Stok Kodu)` anahtarına göre gruplanır.
  * Miktarlar (`qty`) toplanır; böylece depodaki personel aynı ürünü tek bir satırda toplam adet olarak görür.

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
* **Beko Listeleme Davranışı:**
  * **Varsayılan (`filter_beko=False`):** Beko ürünleri **Ana Toplama Listesi**ne (Tümü) dahil edilir. `Urun_Toplama_Listesi_A4_Cikti.xlsx` içinde hem 1. sayfadaki genel toplama listesinde hem de özel **BEKO** sekmesinde yer alır. Aynı zamanda bağımsız `Beko_Urun_Toplama_Listesi_A4.xlsx` dosyası da eş zamanlı üretilerek istendiğinde tek tıkla açılabilir.
  * **Seçenekli Ayrıştırma (`filter_beko=True`):** Kullanıcı arayüzdeki "Beko siparişlerini ana listeden ayır" onay kutusunu işaretlerse, Beko ürünleri ana listeden çıkarılır ve yalnızca bağımsız Beko listesinde yer alır.
* **Diğer Markalar:**
  * `TEFAL`, `BABYLISS`, `LENOVO`, `GRUNDIG`, `BISSELL`, `PHILIPS`, `BRAUN`, `WMF`, `KENWOOD`, `ARIETE`, `LAURASTAR`, `TEKA`, `IPHONE/APPLE`, `FAKIR`, `ARZUM`, `KARACA`, `KORKMAZ` vb.
  * Her marka için Excel çalışma kitabında otomatik olarak özel A4 sekmesi oluşturulur.
* **Büyük Beyaz Eşya, TV ve Garanti Filtrelemesi (`is_excluded_product`):**
  * Sipariş toplama listelerinde yalnızca depodan fiziki toplanacak küçük ev aletleri ve elektronik ürünler yer alır.
  * **Filtrelenen Ürünler (Listeden Hariç Tutulanlar):**
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
    - Ankastre Ocak / Ocaklar (Beko BOCD, BOI vb.)
    - Aspiratör / Ankastre Sürgülü Aspiratör (Beko P 38 vb.)
    - Saç Kurutma Makinesi (BaByliss, Grundig vb.)
    - Tüm Küçük Ev Aletleri (Kahve/Çay Makinesi, Blender, Ütü, Fritöz, Süpürge, Tost Makinesi vb.), Telefon ve Bilgisayarlar.

---

## 5. Çıktı Standartları (A4 Dikey)

* **Sayfa Düzeni:** A4 Dikey, `fitToWidth=1`, `fitToHeight=1` (yazdırırken sayfaya tam sığdırma).
* **3 Sütun Yapısı:**
  1. `Miktar` (Genişlik 12, ortalı, kalın font)
  2. `Ürün Adı` (Genişlik 56, sola dayalı, kelime kaydırma aktif)
  3. `Stok Kodu` (Genişlik 18, ortalı, monospace/koyu font)
* **Stil:** Koyu başlık satırı (`#1E293B`, beyaz yazı), zebra desenli satırlar (`#F8FAFC`), ince gri kenarlıklar.
* **Dosyalar:**
  1. `Urun_Toplama_Listesi_A4_Cikti.xlsx` (Ana Liste + her marka için ayrı sekme).
  2. `Beko_Urun_Toplama_Listesi_A4.xlsx` (Beko ürünleri toplama listesi).
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
