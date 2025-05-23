#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Bu script gerekli paketlerin kurulumunu otomatik olarak yapar.
Kullanım:
    python kurulum.py
"""

import subprocess
import sys
import os

def install_requirements():
    print("Excel Karşılaştırma Uygulaması - Paket Kurulumu")
    print("-" * 50)
    
    # Python versiyonunu kontrol et
    python_version = sys.version_info
    if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 6):
        print("Hata: Bu uygulama Python 3.6 veya daha yüksek bir versiyon gerektirir.")
        print(f"Şu anki Python versiyonu: {sys.version}")
        sys.exit(1)
    
    print(f"Python versiyonu uyumlu: {sys.version}")
    
    # Paketleri kur
    try:
        # requirements.txt dosyasının varlığını kontrol et
        req_file = "requirements.txt"
        if not os.path.exists(req_file):
            # Dosya yoksa, gerekli paketleri içeren bir requirements.txt oluştur
            with open(req_file, "w", encoding="utf-8") as f:
                f.write("pandas>=1.3.0\n")
                f.write("openpyxl>=3.0.9\n")
                f.write("xlrd>=2.0.1\n")
            print(f"'{req_file}' dosyası oluşturuldu.")
        
        print("\nGerekli paketler kuruluyor...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file])
        
        print("\nKurulum başarıyla tamamlandı!")
    
    except subprocess.CalledProcessError as e:
        print(f"\nHata: Paket kurulumu sırasında bir sorun oluştu: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\nHata: {e}")
        sys.exit(1)

if __name__ == "__main__":
    install_requirements()