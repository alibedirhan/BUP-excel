# CAL - Excel Karşılaştırma Aracı

Modern arayüzlü Excel dosyalarını karşılaştırma uygulaması.

## Özellikler

- 📊 İki Excel dosyasını karşılaştırma
- 🚚 Araç-Plasiyer eşleştirme sistemi
- 🎯 Drag & Drop desteği (isteğe bağlı)
- 💾 Excel ve PNG formatında sonuç kaydetme
- 🔤 Büyük/küçük harf duyarlılık seçeneği
- 🎨 Modern kullanıcı arayüzü

## Kurulum

### Gereksinimler

Python 3.7+ gereklidir.

```bash
pip install -r requirements.txt
```

### İsteğe Bağlı (Drag & Drop için)

```bash
pip install tkinterdnd2
```

## Kullanım

### Ubuntu/Linux
```bash
python3 main.py
```

### Windows
```bash
python main.py
```

## Dosya Yapısı

```
.
├── main.py           # Ana uygulama dosyası
├── ui.py             # Kullanıcı arayüzü
├── config.json       # Araç-plasiyer eşleştirmeleri (otomatik oluşur)
├── requirements.txt  # Python bağımlılıkları
└── app.log          # Uygulama logları (otomatik oluşur)
```

## Araç-Plasiyer Eşleştirmesi

İlk çalıştırmada araç-plasiyer eşleştirmesi yapmak için dialog açılır. Bu eşleştirmeler `config.json` dosyasında saklanır ve istediğiniz zaman "Araç-Plasiyer Ayarları" butonuyla düzenleyebilirsiniz.

## Sorun Giderme

### tkinterdnd2 Hatası
Eğer `ModuleNotFoundError: No module named 'tkinterdnd2'` hatası alırsanız:

```bash
pip install tkinterdnd2
```

Modül yüklenemezse program normal gözat butonlarıyla çalışmaya devam eder.

### Font Sorunları
Farklı işletim sistemlerinde font boyutları değişebilir. Kod içinde font ayarları mevcuttur.

## Lisans

Bu proje açık kaynak kodludur.