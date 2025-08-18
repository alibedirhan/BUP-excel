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
import platform

def check_python_version():
    """Python versiyonunu kontrol et"""
    python_version = sys.version_info
    if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 7):
        print("HATA: Bu uygulama Python 3.7 veya daha yüksek bir versiyon gerektirir.")
        print(f"Şu anki Python versiyonu: {sys.version}")
        print("Lütfen Python'u güncelleyin: https://www.python.org/downloads/")
        return False
    
    print(f"✓ Python versiyonu uyumlu: {sys.version}")
    return True

def create_requirements_file():
    """requirements.txt dosyasını oluştur veya güncelle"""
    req_file = "requirements.txt"
    
    # Gerekli paketler listesi - main.py'daki importlara göre
    required_packages = [
        "pandas>=2.0.0",  # Güncellendi: pandas 2.x için
        "openpyxl>=3.0.9", 
        "xlrd>=2.0.1",
        "matplotlib>=3.5.0",
        "tkinterdnd2>=0.3.0"  # Opsiyonel ama faydalı
    ]
    
    try:
        # Mevcut requirements.txt'i oku (varsa)
        existing_packages = set()
        if os.path.exists(req_file):
            with open(req_file, "r", encoding="utf-8") as f:
                existing_packages = set(line.strip() for line in f if line.strip() and not line.startswith('#'))
        
        # Yeni paketleri ekle
        all_packages = list(existing_packages) + required_packages
        
        # Duplicates'leri kaldır (paket adına göre)
        unique_packages = {}
        for package in all_packages:
            pkg_name = package.split('>=')[0].split('==')[0].strip()
            unique_packages[pkg_name] = package
        
        # Dosyayı yaz
        with open(req_file, "w", encoding="utf-8") as f:
            f.write("# Excel Karşılaştırma Uygulaması - Gerekli Paketler\n")
            f.write("# Python 3.7+ gereklidir\n")
            f.write("# Pandas 2.x uyumluluğu için güncellenmiştir\n\n")
            
            # Ana paketler
            f.write("# Ana bağımlılıklar\n")
            for package in sorted(unique_packages.values()):
                f.write(f"{package}\n")
            
            f.write("\n# Opsiyonel - Drag & Drop desteği için\n")
            f.write("# tkinterdnd2 kurulumu başarısız olursa normal gözat butonları kullanılır\n")
        
        print(f"✓ '{req_file}' dosyası güncellendi.")
        return True
        
    except Exception as e:
        print(f"HATA: requirements.txt oluşturulurken hata: {e}")
        return False

def install_package(package_name, optional=False):
    """Tek bir paketi kur"""
    try:
        print(f"  → {package_name} kuruluyor...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", package_name],
            capture_output=True,
            text=True,
            timeout=300  # 5 dakika timeout
        )
        
        if result.returncode == 0:
            print(f"  ✓ {package_name} başarıyla kuruldu")
            return True
        else:
            if optional:
                print(f"  ⚠ {package_name} kurulamadı (opsiyonel): {result.stderr.strip()}")
                return True  # Opsiyonel paketler için başarılı say
            else:
                print(f"  ✗ {package_name} kurulumu başarısız: {result.stderr.strip()}")
                return False
            
    except subprocess.TimeoutExpired:
        print(f"  ✗ {package_name} kurulumu zaman aşımına uğradı")
        return False if not optional else True
    except Exception as e:
        print(f"  ✗ {package_name} kurulumu sırasında hata: {e}")
        return False if not optional else True

def check_pandas_version():
    """Pandas versiyonunu kontrol et ve uyarı ver"""
    try:
        import pandas as pd
        version = pd.__version__
        major_version = int(version.split('.')[0])
        
        print(f"✓ Pandas versiyonu: {version}")
        
        if major_version >= 2:
            print("  ✓ Pandas 2.x - Modern API desteği mevcut")
            return True
        else:
            print("  ⚠ Pandas 1.x - Bazı fonksiyonlar deprecated olabilir")
            print("  → Pandas 2.x'e güncellemek önerilir: pip install --upgrade pandas")
            return True
            
    except ImportError:
        print("  ✗ Pandas henüz kurulmamış")
        return False
    except Exception as e:
        print(f"  ⚠ Pandas versiyonu kontrol edilemedi: {e}")
        return True

def install_requirements():
    """Gerekli paketleri kur"""
    print("Excel Karşılaştırma Uygulaması - Paket Kurulumu")
    print("=" * 55)
    
    # Python versiyonunu kontrol et
    if not check_python_version():
        return False
    
    # Platform bilgisi
    print(f"✓ İşletim sistemi: {platform.system()} {platform.release()}")
    
    # requirements.txt dosyasını oluştur/güncelle
    if not create_requirements_file():
        return False
    
    try:
        # pip'i güncelle
        print("\n📦 pip güncelleniyor...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--upgrade", "pip"
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("✓ pip güncellendi")
        
        # Ana paketleri kur - Pandas 2.x uyumlu
        print("\n📦 Ana paketler kuruluyor...")
        main_packages = ["pandas>=2.0.0", "openpyxl>=3.0.9", "xlrd>=2.0.1", "matplotlib>=3.5.0"]
        
        failed_packages = []
        for package in main_packages:
            if not install_package(package):
                failed_packages.append(package)
        
        # Pandas versiyonunu kontrol et
        if not failed_packages:
            print("\n🔍 Pandas versiyonu kontrol ediliyor...")
            check_pandas_version()
        
        # Opsiyonel paketi kur
        print("\n📦 Opsiyonel paketler kuruluyor...")
        install_package("tkinterdnd2>=0.3.0", optional=True)
        
        # Sonuçları değerlendir
        if failed_packages:
            print(f"\n❌ Bazı ana paketler kurulamadı: {', '.join(failed_packages)}")
            print("Manuel kurulum deneyin:")
            for package in failed_packages:
                print(f"  pip install {package}")
            return False
        else:
            print("\n✅ Tüm paketler başarıyla kuruldu!")
            print("\nUygulamayı başlatmak için:")
            print("  python main.py")
            print("  veya")
            print("  python3 main.py")
            return True
    
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Paket kurulumu sırasında pip hatası: {e}")
        print("Çözüm önerileri:")
        print("1. Internet bağlantınızı kontrol edin")
        print("2. pip'i manuel güncelleyin: python -m pip install --upgrade pip")
        print("3. Yönetici olarak çalıştırmayı deneyin")
        print("4. Ubuntu'da: sudo apt update && sudo apt install python3-pip")
        return False
    except KeyboardInterrupt:
        print("\n⚠ Kurulum kullanıcı tarafından iptal edildi")
        return False
    except Exception as e:
        print(f"\n❌ Beklenmeyen hata: {e}")
        print("Çözüm için:")
        print("1. Python kurulumunuzu kontrol edin")
        print("2. Paketleri manuel kurmayı deneyin")
        print("3. Ubuntu'da: sudo apt install python3-tk python3-pip")
        return False

def verify_installation():
    """Kurulumu doğrula"""
    print("\n🔍 Kurulum doğrulanıyor...")
    
    required_modules = [
        ("pandas", "Veri işleme"),
        ("openpyxl", "Excel dosya desteği"),
        ("xlrd", "Eski Excel formatları"),
        ("matplotlib", "Grafik oluşturma"),
        ("tkinter", "Kullanıcı arayüzü")
    ]
    
    optional_modules = [
        ("tkinterdnd2", "Drag & Drop desteği")
    ]
    
    missing_modules = []
    
    # Ana modülleri kontrol et
    for module, description in required_modules:
        try:
            if module == "pandas":
                import pandas as pd
                print(f"  ✓ {module} ({pd.__version__}) - {description}")
            elif module == "matplotlib":
                import matplotlib
                print(f"  ✓ {module} ({matplotlib.__version__}) - {description}")
            elif module == "openpyxl":
                import openpyxl
                print(f"  ✓ {module} ({openpyxl.__version__}) - {description}")
            else:
                __import__(module)
                print(f"  ✓ {module} - {description}")
        except ImportError:
            print(f"  ✗ {module} - {description} (EKSİK)")
            missing_modules.append(module)
        except Exception as e:
            print(f"  ⚠ {module} - {description} (versiyon kontrol edilemedi)")
    
    # Opsiyonel modülleri kontrol et
    for module, description in optional_modules:
        try:
            __import__(module)
            print(f"  ✓ {module} - {description}")
        except ImportError:
            print(f"  ⚠ {module} - {description} (opsiyonel - eksik)")
    
    if missing_modules:
        print(f"\n❌ Eksik modüller: {', '.join(missing_modules)}")
        print("\nManuel kurulum:")
        if "tkinter" in missing_modules:
            print("  Ubuntu/Debian: sudo apt install python3-tk")
            print("  CentOS/RHEL: sudo yum install tkinter")
        print("  Diğer paketler: pip install " + " ".join(missing_modules))
        return False
    else:
        print("\n✅ Tüm gerekli modüller mevcut!")
        return True

def show_system_info():
    """Sistem bilgilerini göster"""
    print("\n📋 Sistem Bilgileri:")
    print(f"  • Python: {sys.version}")
    print(f"  • Platform: {platform.platform()}")
    print(f"  • İşlemci: {platform.processor()}")
    
    # pip versiyonu
    try:
        result = subprocess.run([sys.executable, "-m", "pip", "--version"], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print(f"  • pip: {result.stdout.strip()}")
    except:
        print("  • pip: Versiyon alınamadı")

def main():
    """Ana fonksiyon"""
    try:
        # Sistem bilgilerini göster
        show_system_info()
        
        # Kurulumu başlat
        if install_requirements():
            # Kurulumu doğrula
            if verify_installation():
                print("\n🎉 Kurulum başarıyla tamamlandı!")
                print("\n📝 Kullanım:")
                print("  python main.py        # Windows")
                print("  python3 main.py       # Linux/Mac")
                print("\n💡 İpucu: Deprecation warning'ları göz ardı edilebilir,")
                print("   program normal çalışır. Gelecek güncellemelerde düzeltilecektir.")
                
                response = input("\nUygulamayı şimdi başlatmak ister misiniz? (y/n): ")
                if response.lower() in ['y', 'yes', 'evet', 'e']:
                    print("\n🚀 Uygulama başlatılıyor...")
                    try:
                        subprocess.run([sys.executable, "main.py"])
                    except FileNotFoundError:
                        print("❌ main.py dosyası bulunamadı!")
                    except KeyboardInterrupt:
                        print("\n⚠ Uygulama kullanıcı tarafından sonlandırıldı")
                else:
                    print("\n✅ Kurulum tamamlandı. İyi kullanımlar!")
                    
            else:
                print("\n⚠ Kurulum tamamlandı ancak bazı modüller eksik olabilir.")
                print("Yukarıdaki talimatları takip ederek eksik modülleri kurun.")
                input("\nÇıkmak için Enter tuşuna basın...")
        else:
            print("\n❌ Kurulum başarısız!")
            print("\n🔧 Manuel kurulum adımları:")
            print("  1. pip install --upgrade pip")
            print("  2. pip install pandas>=2.0.0 openpyxl xlrd matplotlib")
            print("  3. pip install tkinterdnd2  # Opsiyonel")
            print("\n🐧 Ubuntu/Debian için:")
            print("  sudo apt update")
            print("  sudo apt install python3-pip python3-tk")
            print("  pip3 install pandas openpyxl xlrd matplotlib tkinterdnd2")
            input("\nÇıkmak için Enter tuşuna basın...")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n⚠ Program kullanıcı tarafından sonlandırıldı.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Kritik hata: {e}")
        print("\n🔧 Sorun giderme:")
        print("  1. Python kurulumunuzu kontrol edin")
        print("  2. Terminal/Command Prompt'u yönetici olarak çalıştırın")
        print("  3. Internet bağlantınızı kontrol edin")
        sys.exit(1)

if __name__ == "__main__":
    main()