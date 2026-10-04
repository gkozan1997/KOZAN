import sys
from core_engine import get_latest_download_file
from order_processor import process_excel_orders

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

if __name__ == '__main__':
    if len(sys.argv) > 1:
        target_file = sys.argv[1]
    else:
        target_file = get_latest_download_file()
        if not target_file:
            print("İndirilenler klasöründe uygun sipariş Excel dosyası bulunamadı.")
            sys.exit(1)
            
    print(f"İşleniyor: {target_file}")
    res = process_excel_orders(target_file)
    print("\nİşlem Başarıyla Tamamlandı:")
    print(f"  Toplam Sipariş Satırı: {res['total_orders']}")
    print(f"  Ana Liste Ürün Çeşidi: {res['non_beko_count']} (Toplam {res['non_beko_qty']} adet)")
    print(f"  Beko Ürün Çeşidi: {res['beko_count']} (Toplam {res['beko_qty']} adet)")
    print(f"  Markalar: {[b['brand'] + ': ' + str(b['qty']) for b in res['brands']]}")
    print(f"  Kaydedilen Klasör: {res['folder_path']}")
