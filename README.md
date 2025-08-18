# CAL - Excel Karşılaştırma Uygulaması

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![Pandas](https://img.shields.io/badge/Pandas-2.0+-green.svg)](https://pandas.pydata.org/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Build Status](https://github.com/alibedirhan/CAL-excel/workflows/Build%20Windows%20EXE/badge.svg)](https://github.com/alibedirhan/CAL-excel/actions)

Excel dosyalarında cari ünvan karşılaştırması yapan modern, kullanıcı dostu masaüstü uygulaması. Araç-plasiyer eşleştirme sistemi ile otomatik dosya adlandırma ve çoklu format desteği sunar.

## 🚀 Özellikler

### 📊 Temel Fonksiyonlar
- **Excel Karşılaştırma**: İki Excel dosyasında cari ünvan farklılıklarını tespit eder
- **Akıllı Başlık Algılama**: "Cari Ünvan" sütununu otomatik bulur
- **Veri Temizleme**: Boşlukları ve duplikaları otomatik temizler
- **Büyük/Küçük Harf Kontrolü**: Opsiyonel duyarlılık ayarı

### 🚗 Araç-Plasiyer Sistemi
- **Otomatik Araç Tespiti**: Depo kartından araç numarası çıkarır
- **Plasiyer Eşleştirme**: Araç numaralarını plasiyer adlarıyla eşleştirir
- **Akıllı Dosya Adlandırma**: "Araç 01 Ahmet ALTILI" formatında dosya adları
- **JSON Konfigürasyon**: Esnek ve düzenlenebilir ayar sistemi

### 💾 Çıktı Formatları
- **Excel (.xlsx)**: Detaylı rapor ve başlık bilgisi ile
- **PNG Resim**: Görsel tablo formatında
- **Çoklu Kaydetme**: Her iki formatı aynı anda kaydetme

### 🎨 Modern Arayüz
- **Drag & Drop**: Dosyaları sürükleyip bırakma desteği
- **Responsive Tasarım**: Tüm ekran boyutlarına uyumlu
- **Progress Bar**: İşlem durumu göstergesi
- **Uniform Font**: Tutarlı Segoe UI 8px font sistemi
- **Thread-Safe**: Donmayan kullanıcı arayüzü

## 📦 Kurulum

### Gereksinimler
- Python 3.7 veya üzeri
- Windows/Linux/macOS desteği

### Otomatik Kurulum
```bash
# Repository'yi klonlayın
git clone https://github.com/alibedirhan/CAL-excel.git
cd CAL-excel

# Otomatik kurulum scripti çalıştırın
python kurulum.py
```

### Manuel Kurulum
```bash
# Gerekli paketleri yükleyin
pip install pandas>=2.0.0 openpyxl>=3.0.9 xlrd>=2.0.1 matplotlib>=3.5.0

# Opsiyonel: Drag & Drop desteği için
pip install tkinterdnd2>=0.3.0
```

## 🔧 Kullanım

### Hızlı Başlangıç
1. **Uygulamayı Başlatın**:
   ```bash
   python main.py
   ```

2. **Dosyaları Seçin**:
   - Eski tarihli Excel dosyasını seçin/sürükleyin
   - Yeni tarihli Excel dosyasını seçin/sürükleyin

3. **Karşılaştır**:
   - "🔍 Karşılaştır" butonuna tıklayın
   - Sonuçlar otomatik olarak kaydedilir

### Araç-Plasiyer Ayarları

İlk çalıştırmada araç-plasiyer eşleştirmesi yapın:

```json
{
    "vehicle_drivers": {
        "01": "Ahmet ALTILI",
        "02": "Erhan AYDOĞDU",
        "04": "Soner TANAY",
        "05": "Süleyman TANAY",
        "06": "Hakan YILMAZ"
    }
}
```

### Desteklenen Excel Formatları
- Excel 2007-2019 (.xlsx)
- Excel 97-2003 (.xls)
- Maximum dosya boyutu: 100MB

## 📁 Proje Yapısı

```
CAL-excel/
├── main.py              # Ana uygulama dosyası
├── ui.py                # Modern kullanıcı arayüzü
├── kurulum.py           # Otomatik kurulum scripti
├── requirements.txt     # Python bağımlılıkları
├── config.json          # Araç-plasiyer konfigürasyonu
├── .gitignore          # Git ignore kuralları
├── README.md           # Bu dosya
└── .github/
    └── workflows/
        ├── build.yml        # CI/CD pipeline
        └── build-release.yml # Release automation
```

## 🛠️ Geliştirme

### Pandas 2.x Uyumluluğu
Bu proje Pandas 2.x ile tam uyumludur:
- `applymap` yerine `map` kullanımı
- Deprecation warning'lar düzeltildi
- Modern API desteği

### Katkıda Bulunma
1. Fork yapın
2. Feature branch oluşturun (`git checkout -b feature/amazing-feature`)
3. Commit yapın (`git commit -m 'Add amazing feature'`)
4. Push yapın (`git push origin feature/amazing-feature`)
5. Pull Request oluşturun

### Test Etme
```bash
# Uygulamayı test edin
python main.py

# Kurulum doğrulaması
python -c "import pandas, openpyxl, matplotlib; print('Tüm modüller başarıyla yüklendi!')"
```

## 📊 Örnek Kullanım Senaryoları

### Senaryo 1: Müşteri Listesi Karşılaştırma
- **Durum**: Eski ve yeni müşteri listelerini karşılaştırma
- **Çıktı**: Yeni listede olmayan müşteriler
- **Format**: Excel + görsel rapor

### Senaryo 2: Araç Bazlı Raporlama
- **Durum**: Araç 01'e ait müşteri değişiklikleri
- **Çıktı**: "Araç 01 Ahmet YILMAZ.xlsx"
- **Avantaj**: Otomatik dosya adlandırma

### Senaryo 3: Toplu Analiz
- **Durum**: Birden fazla araç dosyası analizi
- **Çıktı**: Her araç için ayrı rapor
- **Özellik**: Batch processing desteği

## 🔧 Sorun Giderme

### Yaygın Sorunlar

**1. tkinterdnd2 Yükleme Hatası**
```bash
pip install tkinterdnd2
# veya drag & drop olmadan kullanın
```

**2. Pandas Uyumluluk Uyarıları**
```bash
pip install --upgrade pandas>=2.0.0
```

**3. Excel Dosyası Açılmıyor**
- Dosya boyutunu kontrol edin (max 100MB)
- Dosya izinlerini kontrol edin
- Excel dosyasının bozuk olmadığından emin olun

**4. "Cari Ünvan" Sütunu Bulunamıyor**
- Excel dosyasında "Cari Ünvan" başlığının olduğundan emin olun
- Başlık satırının doğru konumda olduğunu kontrol edin

### Log Dosyaları
Hata durumunda `app.log` dosyasını kontrol edin:
```bash
tail -f app.log
```

## 📄 Lisans

Bu proje MIT lisansı altında dağıtılmaktadır. Detaylar için [LICENSE](LICENSE) dosyasına bakın.

## 👨‍💻 Geliştirici

**Ali Bedirhan**
- GitHub: [@alibedirhan](https://github.com/alibedirhan)
- Email: [İletişim bilgileri]

## 🙏 Teşekkürler

- Pandas takımına veri işleme kütüphanesi için
- OpenPyXL geliştiricilerine Excel desteği için
- Matplotlib takımına görselleştirme için
- tkinterdnd2 geliştiricilerine drag & drop desteği için

## 📈 Versiyon Geçmişi

- **v3.0** - Pandas 2.x uyumluluğu, uniform font sistemi
- **v2.5** - Araç-plasiyer otomatik eşleştirme
- **v2.0** - Modern UI, drag & drop desteği
- **v1.0** - Temel Excel karşılaştırma fonksiyonu

---

**⭐ Bu projeyi beğendiyseniz yıldız vermeyi unutmayın!**
