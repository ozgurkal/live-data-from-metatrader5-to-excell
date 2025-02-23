import MetaTrader5 as mt5
import time
import os
import sys  # Sistem çıkışı için ekleme yapıldı

# Çalıştırılan dosyanın bulunduğu dizini al

if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)  # Exe modunda çalışıyorsa
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))  # Normal Python dosyası için
symbols_file = os.path.join(current_dir, "semboller.txt")  # İzlenecek semboller
data_file = os.path.join(current_dir, "veri.dat")  # Verilerin yazılacağı dosya

# MetaTrader 5'e bağlan
if not mt5.initialize():
    print("MT5 bağlantısı başarısız!")
    input("Devam etmek için bir tuşa basın...")  # Çıkmadan önce bekle
    sys.exit()  # quit() yerine sys.exit() kullanıldı

try:
    # İzlenecek sembolleri dosyadan oku
    if not os.path.exists(symbols_file):
        print(f"{symbols_file} bulunamadı! Lütfen dosyayı oluşturun ve içine sembolleri yazın.")
        input("Devam etmek için bir tuşa basın...")  # Bekletme ekledik
        sys.exit()

    with open(symbols_file, "r") as f:
        symbols = [line.strip() for line in f.readlines() if line.strip()]

    if not symbols:
        print("Sembol listesi boş! 'semboller.txt' dosyasına semboller ekleyin.")
        input("Devam etmek için bir tuşa basın...")
        sys.exit()

    print(f"İzlenen semboller: {symbols}")

    # Sonsuz döngü ile sürekli veri güncelle
    while True:
        data_lines = []
        for symbol in symbols:
            tick = mt5.symbol_info_tick(symbol)
            if tick:
                data_lines.append(f"{symbol},{tick.bid},{tick.ask}")
            else:
                print(f"{symbol} için veri alınamadı!")

        # Veriyi dosyaya yaz
        with open(data_file, "w") as f:
            f.write("\n".join(data_lines))

        time.sleep(1)  # 1 saniyede bir güncelle

except Exception as e:
    print(f"Hata oluştu: {e}")

finally:
    mt5.shutdown()  # MT5 bağlantısını kapat
