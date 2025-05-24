import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import os
import re
import json
from datetime import datetime
import matplotlib
matplotlib.use('Agg')  # Backend'i GUI uygulaması için ayarla
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from tkinterdnd2 import TkinterDnD  # Drag & Drop desteği
from ui import ModernExcelComparisonUI
import threading
import logging

# Logging sistemi kurulumu
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class VehicleDriverSetupDialog:
    """Araç-Plasiyer Eşleştirme Dialog'u - Uniform Font"""
    def __init__(self, parent, existing_data=None):
        self.parent = parent
        self.existing_data = existing_data or {}
        self.result = {}
        self.dialog = None
        self.entries = {}  # entries'i __init__'de tanımla
        
    def show_setup_dialog(self):
        """Araç-plasiyer eşleştirme dialog'unu göster - Uniform Font"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Araç-Plasiyer Eşleştirmesi")
        self.dialog.geometry("500x600")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Dialog'u ortala
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (600 // 2)
        self.dialog.geometry(f"500x600+{x}+{y}")
        
        # Ana frame
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Başlık - uniform font
        title_label = tk.Label(
            main_frame,
            text="Araç-Plasiyer Eşleştirmesi",
            font=('Segoe UI', 8, 'bold')  # Arial 14 → Segoe UI 8
        )
        title_label.pack(pady=(0, 10))
        
        # Açıklama - uniform font
        desc_label = tk.Label(
            main_frame,
            text="Lütfen her araç numarası için plasiyer adını girin.\nBoş bırakılan araçlar kullanılmayacak.",
            font=('Segoe UI', 8),  # Arial 10 → Segoe UI 8
            justify=tk.LEFT
        )
        desc_label.pack(pady=(0, 15))
        
        # Scrollable frame
        canvas = tk.Canvas(main_frame)
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Entry'ler için dict
        self.entries = {}
        
        # 20 araç için entry oluştur - uniform font
        for i in range(1, 21):
            vehicle_num = f"{i:02d}"
            
            # Frame for each vehicle
            vehicle_frame = tk.Frame(scrollable_frame)
            vehicle_frame.pack(fill=tk.X, pady=2)
            
            # Label - uniform font
            label = tk.Label(
                vehicle_frame,
                text=f"Araç {vehicle_num}:",
                width=10,
                anchor='w',
                font=('Segoe UI', 8)  # Uniform font eklendi
            )
            label.pack(side=tk.LEFT)
            
            # Entry - uniform font
            entry = tk.Entry(
                vehicle_frame, 
                width=30,
                font=('Segoe UI', 8)  # Uniform font eklendi
            )
            entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
            self.entries[vehicle_num] = entry
            
            # Mevcut veriyi yükle (varsa)
            if vehicle_num in self.existing_data:
                entry.insert(0, self.existing_data[vehicle_num])
            
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Butonlar - uniform font
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=15)
        
        # Kaydet butonu - uniform font
        save_btn = tk.Button(
            button_frame,
            text="Kaydet",
            command=self.save_config,
            bg="#4CAF50",
            fg="white",
            font=('Segoe UI', 8, 'bold'),  # Arial 10 → Segoe UI 8
            padx=20
        )
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        # İptal butonu - uniform font
        cancel_btn = tk.Button(
            button_frame,
            text="İptal",
            command=self.cancel,
            bg="#f44336",
            fg="white",
            font=('Segoe UI', 8),  # Arial 10 → Segoe UI 8
            padx=20
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Örnek veriler yükle - uniform font
        load_sample_btn = tk.Button(
            button_frame,
            text="Örnek Veriler",
            command=self.load_sample_data,
            bg="#2196F3",
            fg="white",
            font=('Segoe UI', 8),  # Arial 10 → Segoe UI 8
            padx=20
        )
        load_sample_btn.pack(side=tk.LEFT, padx=5)
        
        # Dialog'u modal yap
        self.dialog.wait_window()
        return self.result
    
    def load_sample_data(self):
        """Örnek verileri yükle"""
        sample_data = {
            "01": "Ahmet ALTILI",
            "02": "Erhan AYDOĞDU", 
            "04": "Soner TANAY",
            "05": "Süleyman TANAY",
            "06": "Hakan YILMAZ"
        }
        
        for vehicle_num, driver_name in sample_data.items():
            if vehicle_num in self.entries:
                self.entries[vehicle_num].delete(0, tk.END)
                self.entries[vehicle_num].insert(0, driver_name)
    
    def save_config(self):
        """Konfigürasyonu kaydet"""
        # Entry'lerden verileri al
        vehicle_drivers = {}
        for vehicle_num, entry in self.entries.items():
            driver_name = entry.get().strip()
            if driver_name:  # Boş olmayan girişleri ekle
                vehicle_drivers[vehicle_num] = driver_name
        
        if not vehicle_drivers:
            messagebox.showwarning("Uyarı", "En az bir araç-plasiyer eşleştirmesi yapmalısınız!")
            return
        
        # Config dict'i oluştur
        config = {"vehicle_drivers": vehicle_drivers}
        
        try:
            # config.json dosyasına kaydet
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            self.result = vehicle_drivers
            messagebox.showinfo("Başarılı", f"{len(vehicle_drivers)} araç-plasiyer eşleştirmesi kaydedildi!")
            self.dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme hatası: {str(e)}")
    
    def cancel(self):
        """İptal et"""
        self.result = None
        self.dialog.destroy()


class ExcelComparisonLogic:
    """Excel karşılaştırma iş mantığı"""
    def __init__(self):
        # Dosya yolları
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # Seçenekler
        self.case_sensitive = tk.BooleanVar(value=False)
        self.save_format = tk.StringVar(value="excel")
        
        # Varsayılan çıktı dosyası adı - boş başlar
        self.output_path.set("")
        
        # UI referansı
        self.ui = None
        
        # Maximum dosya boyutu (MB)
        self.max_file_size_mb = 100
        
        # Araç-Şoför eşleştirme tablosunu yükle
        self.vehicle_drivers = self.load_vehicle_drivers()
        
    def load_vehicle_drivers(self):
        """Araç-plasiyer eşleştirmesini dosyadan yükler"""
        try:
            # Önce config.json dosyasını dene
            config_files = ['config.json', 'vehicle_config.json', 'drivers.json']
            
            for config_file in config_files:
                if os.path.exists(config_file):
                    with open(config_file, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                        vehicle_drivers = config.get('vehicle_drivers', {})
                        if vehicle_drivers:
                            logging.info(f"Araç-plasiyer konfigürasyonu yüklendi: {config_file}")
                            return vehicle_drivers
            
            # Hiçbir config dosyası bulunamazsa çevre değişkenlerini kontrol et
            env_config = {}
            for i in range(1, 21):  # 01-20 arası araç numaralarını kontrol et
                key = f"DRIVER_{i:02d}"
                if key in os.environ:
                    env_config[f"{i:02d}"] = os.environ[key]
            
            if env_config:
                logging.info("Araç-plasiyer konfigürasyonu çevre değişkenlerinden yüklendi")
                return env_config
                
            # Hiçbiri bulunamazsa uyarı ver ve boş dict döndür
            logging.warning("Araç-plasiyer konfigürasyonu bulunamadı!")
            return {}
            
        except json.JSONDecodeError:
            logging.error("Config dosyası geçersiz JSON formatında!")
            return {}
        except Exception as e:
            logging.error(f"Config yükleme hatası: {str(e)}")
            return {}
    
    def show_vehicle_setup_dialog(self):
        """Araç-plasiyer eşleştirme dialog'unu göster"""
        if not self.ui or not self.ui.root:
            return False
            
        dialog = VehicleDriverSetupDialog(self.ui.root, self.vehicle_drivers)
        result = dialog.show_setup_dialog()
        
        if result:
            self.vehicle_drivers = result
            return True
        return False
    
    def set_ui(self, ui):
        """UI referansını ayarla"""
        self.ui = ui
        
        # Eğer config yüklenemişse kullanıcıdan eşleştirme isteyin
        if not self.vehicle_drivers and ui:
            response = messagebox.askyesno(
                "Araç-Plasiyer Eşleştirmesi",
                "Araç-plasiyer eşleştirmesi bulunamadı.\n\n"
                "Şimdi eşleştirme yapmak ister misiniz?\n\n"
                "Evet: Eşleştirme ekranını aç\n"
                "Hayır: Varsayılan örnek config oluştur"
            )
            
            if response:
                # Dialog'u göster
                if self.show_vehicle_setup_dialog():
                    logging.info("Kullanıcı araç-plasiyer eşleştirmesi yaptı")
                else:
                    logging.info("Kullanıcı araç-plasiyer eşleştirmesini iptal etti")
            else:
                # Varsayılan config oluştur
                self.create_default_config()
                self.vehicle_drivers = self.load_vehicle_drivers()
    
    def create_default_config(self):
        """Varsayılan config dosyasını oluşturur"""
        default_config = {
            "vehicle_drivers": {
                "01": "Plasiyer 1",
                "02": "Plasiyer 2",
                "03": "Plasiyer 3",
                "04": "Plasiyer 4",
                "05": "Plasiyer 5"
            }
        }
        
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            
            logging.info("Varsayılan config.json dosyası oluşturuldu")
            
            if self.ui:
                self.ui.show_info("Bilgi", 
                    "Varsayılan config.json dosyası oluşturuldu.\n"
                    "Daha sonra Ayarlar menüsünden düzenleyebilirsiniz.")
                    
        except Exception as e:
            logging.error(f"Varsayılan config oluşturma hatası: {str(e)}")
            if self.ui:
                self.ui.show_error("Hata", f"Config dosyası oluşturulamadı: {str(e)}")
    
    def edit_vehicle_drivers(self):
        """Araç-plasiyer eşleştirmesini düzenle"""
        if not self.ui:
            return
        
        # Mevcut verileri dialog'a geç
        dialog = VehicleDriverSetupDialog(self.ui.root, self.vehicle_drivers)
        result = dialog.show_setup_dialog()
        
        if result:
            self.vehicle_drivers = result
            if self.ui:
                self.ui.show_info("Başarılı", "Araç-plasiyer eşleştirmesi güncellendi!")
        
    def validate_file_size(self, file_path):
        """Dosya boyutunu kontrol et"""
        try:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > self.max_file_size_mb:
                return False, f"Dosya boyutu çok büyük ({file_size_mb:.1f}MB). Maximum {self.max_file_size_mb}MB destekleniyor."
            return True, ""
        except Exception as e:
            return False, f"Dosya boyutu kontrol edilemedi: {str(e)}"
    
    def find_header_row(self, df):
        """DataFrame içinde başlık satırını bulur - Optimize edilmiş versiyon"""
        try:
            # DataFrame'i string'e çevir ve "Cari Ünvan" içeren satırları bul
            df_str = df.astype(str)
            mask = df_str.applymap(lambda x: "Cari Ünvan" in x if isinstance(x, str) else False).any(axis=1)
            
            if mask.any():
                return mask.idxmax()
            return -1
        except Exception as e:
            logging.error(f"Başlık satırı bulma hatası: {str(e)}")
            # Fallback - eski yöntem
            for i, row in df.iterrows():
                for value in row.values:
                    if isinstance(value, str) and "Cari Ünvan" in value:
                        return i
            return -1
    
    def extract_vehicle_number(self, depo_text):
        """Depo kartı metninden araç numarasını çıkarır"""
        try:
            if not isinstance(depo_text, str):
                return None
                
            # Farklı formatları dene
            patterns = [
                r'[Aa]raç\s*(\d{1,2})',           # "Araç 01", "araç 1" 
                r'[Aa]rac\s*(\d{1,2})',           # "Arac 01" (typo)
                r'[Vv]ehicle\s*(\d{1,2})',        # "Vehicle 01"
                r'(\d{1,2})\s*[Nn]o',             # "01 No", "1 No"
                r'(\d{1,2})\s*[Nn]olu',           # "01 Nolu"
                r'\b(\d{1,2})\b'                  # Sadece rakam (en son denenir)
            ]
            
            for pattern in patterns:
                match = re.search(pattern, depo_text)
                if match:
                    vehicle_num = match.group(1).zfill(2)  # 2 haneli yap (01, 02, etc.)
                    if vehicle_num in self.vehicle_drivers:
                        return vehicle_num
                        
            return None
            
        except Exception as e:
            logging.error(f"Araç numarası çıkarma hatası: {str(e)}")
            return None
    
    def create_filename_with_driver(self, depo_text):
        """Depo kartından araç numarası çıkarıp plasiyer adıyla dosya adı oluşturur"""
        try:
            # Araç numarasını çıkar
            vehicle_num = self.extract_vehicle_number(depo_text)
            
            if vehicle_num and vehicle_num in self.vehicle_drivers:
                driver_name = self.vehicle_drivers[vehicle_num]
                # "Araç 01 Ahmet ALTILI" formatında dosya adı oluştur
                filename = f"Araç {vehicle_num} {driver_name}"
                # Dosya adını güvenli hale getir
                return self.sanitize_filename(filename)
            else:
                # Araç numarası bulunamadıysa eski yöntemi kullan
                return self.sanitize_filename(depo_text) if depo_text else ""
                
        except Exception as e:
            logging.error(f"Plasiyerli dosya adı oluşturma hatası: {str(e)}")
            # Hata durumunda güvenli bir ad döndür
            return self.sanitize_filename(depo_text) if depo_text else ""
    
    def sanitize_filename(self, filename):
        """Dosya adını güvenli hale getirir"""
        # Windows'da yasak karakterleri kaldır
        invalid_chars = r'[\\/*?:"<>|]'
        safe_name = re.sub(invalid_chars, '', filename)
        
        # Türkçe karakterleri korur, sadece yasak olanları kaldırır
        safe_name = safe_name.strip()
        
        # Çok uzun ise kısalt
        if len(safe_name) > 100:
            safe_name = safe_name[:100]
            
        return safe_name if safe_name else "output"
    
    def update_output_filename(self, file_path):
        """Seçilen dosyaya göre çıktı dosya adını günceller"""
        try:
            # Dosya boyutunu kontrol et
            is_valid, error_msg = self.validate_file_size(file_path)
            if not is_valid:
                if self.ui:
                    self.ui.show_warning("Uyarı", error_msg)
                return
                
            # Excel dosyasını oku (sadece ilk 10 satır)
            df = pd.read_excel(file_path, header=None, nrows=10)
            
            # Depo adını bul
            depo_name = None
            for i in range(min(10, len(df))):
                row_str = str(df.iloc[i, 0]) if len(df.columns) > 0 else None
                if isinstance(row_str, str) and "Depo Kartı" in row_str:
                    # [xxxx] DEPO ADI formatını bul
                    match = re.search(r'\[(.*?)\]\s*(.*?)(?=\n|\r\n|$)', row_str)
                    if match and match.group(2):
                        depo_name = match.group(2).strip()
                        break
            
            # Depo adı bulunduysa çıktı dosya adını güncelle (araç-plasiyer eşleştirmeli)
            if depo_name:
                # Yeni fonksiyonu kullan - araç numarasına göre plasiyer adı ekle
                filename_with_driver = self.create_filename_with_driver(depo_name)
                self.output_path.set(filename_with_driver)
            else:
                # Depo adı bulunamadıysa boş bırak
                self.output_path.set("")
                
        except PermissionError:
            if self.ui:
                self.ui.show_error("Hata", "Dosyaya erişim izni yok!")
        except pd.errors.EmptyDataError:
            if self.ui:
                self.ui.show_error("Hata", "Excel dosyası boş veya bozuk!")
        except Exception as e:
            # Hata olursa boş bırak
            logging.error(f"Dosya adı güncelleme hatası: {str(e)}")
            if self.ui:
                self.ui.show_warning("Uyarı", f"Dosya adı güncellenemedi: {str(e)}")
    
    def save_results_as_image(self, unique_cari_unvan_list, output_path, depo_name=None):
        """Sonuçları resim dosyası olarak kaydeder - Düzeltilmiş versiyon"""
        try:
            # Matplotlib figürü oluştur
            plt.figure(figsize=(12, 8), dpi=150)
            
            # Türkçe karakter desteği için font ayarla
            plt.rcParams['font.family'] = 'DejaVu Sans'
            
            # Başlık için araç-plasiyer bilgisini kullan
            if depo_name:
                # Araç numarasını çıkar ve plasiyer adını ekle
                vehicle_num = self.extract_vehicle_number(depo_name)
                if vehicle_num and vehicle_num in self.vehicle_drivers:
                    driver_name = self.vehicle_drivers[vehicle_num]
                    title = f"Araç {vehicle_num} - {driver_name}"
                else:
                    title = depo_name
                plt.suptitle(title, fontsize=16, fontweight='bold')
            else:
                plt.suptitle("Eksik Cari Ünvanlar", fontsize=16, fontweight='bold')
                
            # Tablo verilerini oluştur
            cell_text = []
            for i, unvan in enumerate(unique_cari_unvan_list, 1):
                # Çok uzun ünvanları kısalt
                display_unvan = unvan if len(str(unvan)) <= 80 else str(unvan)[:77] + "..."
                cell_text.append([i, display_unvan])
                
            # Eğer liste boşsa
            if not cell_text:
                cell_text = [["", "Tüm cari ünvanlar her iki dosyada da mevcut."]]
                
            # Tabloyu oluştur
            plt.axis('off')  # Eksen çizgilerini kapat
            table = plt.table(
                cellText=cell_text,
                colLabels=["#", "Cari Ünvan"],
                loc='center',
                cellLoc='left',
                colWidths=[0.1, 0.9]
            )
            
            # Tablo stilini ayarla
            table.auto_set_font_size(False)
            table.set_fontsize(9)
            table.scale(1, 1.5)  # Tablo yüksekliğini ayarla
            
            # Başlık satırını kalın göster
            for (i, j), cell in table.get_celld().items():
                if i == 0:  # Başlık satırı
                    cell.set_text_props(fontweight='bold')
                    cell.set_facecolor('#e6e6e6')
                else:
                    # Veri satırları için zebra efekti
                    if i % 2 == 0:
                        cell.set_facecolor('#f9f9f9')
            
            # Dosyayı kaydet
            plt.savefig(output_path, bbox_inches='tight', dpi=150, 
                       facecolor='white', edgecolor='none')
            plt.close()  # Bellek sızıntısını önle
            
            return True
            
        except PermissionError:
            logging.error(f"Resim kaydetme izin hatası: {output_path}")
            return False, "Dosya kaydetme izni yok!"
        except Exception as e:
            logging.error(f"Resim kaydetme hatası: {str(e)}")
            return False, f"Resim kaydetme hatası: {str(e)}"
    
    def validate_excel_file(self, file_path):
        """Excel dosyasının geçerli olup olmadığını kontrol eder"""
        try:
            # Dosya var mı?
            if not os.path.exists(file_path):
                return False, "Dosya bulunamadı!"
                
            # Dosya boyutu kontrolü
            is_valid, error_msg = self.validate_file_size(file_path)
            if not is_valid:
                return False, error_msg
                
            # Excel dosyasını okumaya çalış
            pd.read_excel(file_path, nrows=1)
            return True, ""
            
        except PermissionError:
            return False, "Dosyaya erişim izni yok!"
        except pd.errors.EmptyDataError:
            return False, "Excel dosyası boş!"
        except Exception as e:
            return False, f"Geçersiz Excel dosyası: {str(e)}"
    
    def compare_files_thread(self):
        """Dosya karşılaştırmayı ayrı thread'de çalıştır"""
        try:
            self.compare_files_internal()
        except Exception as e:
            logging.error(f"Thread hatası: {str(e)}")
            if self.ui:
                self.ui.show_error("Hata", f"İşlem sırasında beklenmeyen hata: {str(e)}")
    
    def compare_files(self):
        """Excel dosyalarını karşılaştırır - Ana fonksiyon"""
        # Thread'de çalıştır
        thread = threading.Thread(target=self.compare_files_thread)
        thread.daemon = True
        thread.start()
    
    def compare_files_internal(self):
        """Excel dosylarını karşılaştırır ve sonuçları gösterir - İç fonksiyon"""
        # Önce girdileri kontrol et
        file1_path = self.file1_path.get()
        file2_path = self.file2_path.get()
        output_path = self.output_path.get()
        
        if not file1_path or not file2_path:
            if self.ui:
                self.ui.show_error("Hata", "Lütfen her iki Excel dosyasını da seçin!")
            return
        
        # Dosyaları doğrula
        is_valid1, error1 = self.validate_excel_file(file1_path)
        if not is_valid1:
            if self.ui:
                self.ui.show_error("Hata", f"Eski tarihli dosya hatası: {error1}")
            return
            
        is_valid2, error2 = self.validate_excel_file(file2_path)
        if not is_valid2:
            if self.ui:
                self.ui.show_error("Hata", f"Yeni tarihli dosya hatası: {error2}")
            return
        
        # Sonuç listesini temizle
        self.clear_results()
        
        try:
            # Excel dosyalarını oku (başlık satırını bilmediğimiz için tüm içeriği okuyoruz)
            logging.info(f"Dosyalar okunuyor: {file1_path}, {file2_path}")
            
            df1_full = pd.read_excel(file1_path, header=None)
            df2_full = pd.read_excel(file2_path, header=None)
            
            # Depo Kartı bilgisini çıkar
            depo_name = None
            for i in range(min(10, len(df1_full))):
                row_str = str(df1_full.iloc[i, 0]) if len(df1_full.columns) > 0 else None
                if isinstance(row_str, str) and "Depo Kartı" in row_str:
                    # [xxxx] DEPO ADI formatını bul
                    match = re.search(r'\[(.*?)\]\s*(.*?)(?=\n|\r\n|$)', row_str)
                    if match and match.group(2):
                        depo_name = match.group(2).strip()
                        break
            
            # Başlık satırlarını bul
            header_row1 = self.find_header_row(df1_full)
            header_row2 = self.find_header_row(df2_full)
            
            if header_row1 == -1 or header_row2 == -1:
                if self.ui:
                    self.ui.show_error("Hata", "Excel dosyalarında 'Cari Ünvan' başlığı bulunamadı!")
                return
            
            # Başlık satırlarını kullanarak tekrar oku
            df1 = pd.read_excel(file1_path, header=header_row1)
            df2 = pd.read_excel(file2_path, header=header_row2)
            
            # Sütun isimlerini temizle (baştaki ve sondaki boşlukları kaldır)
            df1.columns = [col.strip() if isinstance(col, str) else col for col in df1.columns]
            df2.columns = [col.strip() if isinstance(col, str) else col for col in df2.columns]
            
            # İki dosyada da "Cari Ünvan" sütunu var mı kontrol et
            cari_unvan_col1 = next((col for col in df1.columns if isinstance(col, str) and "Cari Ünvan" in col), None)
            cari_unvan_col2 = next((col for col in df2.columns if isinstance(col, str) and "Cari Ünvan" in col), None)
            
            if not cari_unvan_col1 or not cari_unvan_col2:
                if self.ui:
                    self.ui.show_error("Hata", "Bir veya daha fazla Excel dosyasında 'Cari Ünvan' sütunu bulunamadı.")
                return
            
            # Cari ünvanları liste haline getir ve boşlukları temizle
            cari_unvan_list1 = df1[cari_unvan_col1].dropna().apply(
                lambda x: x.strip() if isinstance(x, str) else str(x).strip()
            ).tolist()
            cari_unvan_list2 = df2[cari_unvan_col2].dropna().apply(
                lambda x: x.strip() if isinstance(x, str) else str(x).strip()
            ).tolist()
            
            # Boş değerleri filtrele
            cari_unvan_list1 = [x for x in cari_unvan_list1 if x and x.strip()]
            cari_unvan_list2 = [x for x in cari_unvan_list2 if x and x.strip()]
            
            # Büyük/küçük harf duyarlılığını devre dışı bırakma seçeneği
            if not self.case_sensitive.get():
                cari_unvan_list1_upper = [unvan.upper() for unvan in cari_unvan_list1]
                cari_unvan_list2_upper = [unvan.upper() for unvan in cari_unvan_list2]
                
                # Birinci dosyada olup ikinci dosyada olmayan cari ünvanları bul (büyük/küçük harf duyarsız)
                unique_cari_unvan_list = [
                    cari_unvan_list1[i] for i, unvan in enumerate(cari_unvan_list1_upper) 
                    if unvan not in cari_unvan_list2_upper
                ]
            else:
                # Birinci dosyada olup ikinci dosyada olmayan cari ünvanları bul (büyük/küçük harf duyarlı)
                unique_cari_unvan_list = [unvan for unvan in cari_unvan_list1 if unvan not in cari_unvan_list2]
            
            # Duplicates'leri kaldır (sırayı koruyarak)
            seen = set()
            unique_cari_unvan_list = [x for x in unique_cari_unvan_list if not (x in seen or seen.add(x))]
            
            # Sonuçları UI'ya gönder
            status_text = f"Toplam {len(cari_unvan_list1)} cari ünvandan {len(unique_cari_unvan_list)} tanesi yeni dosyada bulunmuyor."
            if self.ui:
                self.ui.update_results(unique_cari_unvan_list, status_text)
            
            logging.info(f"Karşılaştırma tamamlandı. {len(unique_cari_unvan_list)} farklılık bulundu.")
            
            # Sonuçları kaydet
            self.save_results(unique_cari_unvan_list, output_path, depo_name)
        
        except MemoryError:
            if self.ui:
                self.ui.show_error("Hata", "Dosyalar çok büyük, bellek yetersiz!")
        except pd.errors.EmptyDataError:
            if self.ui:
                self.ui.show_error("Hata", "Excel dosyalarından biri boş veya bozuk!")
        except PermissionError:
            if self.ui:
                self.ui.show_error("Hata", "Dosyalara erişim izni yok!")
        except Exception as e:
            logging.error(f"Karşılaştırma hatası: {str(e)}")
            if self.ui:
                self.ui.show_error("Hata", f"İşlem sırasında bir hata oluştu: {str(e)}")
    
    def save_results(self, unique_cari_unvan_list, output_path, depo_name):
        """Sonuçları kaydet"""
        if not output_path:
            return
            
        try:
            # UI'dan checkbox değerlerini al
            save_excel = self.ui.save_excel.get() if self.ui else True
            save_image = self.ui.save_image.get() if self.ui else False
            
            saved_files = []
            
            # Excel formatında kaydet
            if save_excel:
                excel_path = output_path + ".xlsx"
                
                try:
                    # Excel dosyasını oluştur
                    result_df = pd.DataFrame({"Cari Ünvan": unique_cari_unvan_list})
                    
                    # Context manager kullanarak güvenli kaydetme
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        # Depo kartı bilgisini üst satıra yaz - araç ve plasiyer bilgisiyle
                        if depo_name:
                            # Araç numarasını çıkar ve plasiyer adını ekle
                            vehicle_num = self.extract_vehicle_number(depo_name)
                            if vehicle_num and vehicle_num in self.vehicle_drivers:
                                driver_name = self.vehicle_drivers[vehicle_num]
                                header_text = f"Araç {vehicle_num} - {driver_name}"
                            else:
                                header_text = depo_name
                                
                            # Yeni bir DataFrame oluştur - başlık bilgisini içeren
                            header_df = pd.DataFrame({" ": [header_text]})
                            header_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                            
                            # Cari ünvan listesini 2. satırdan başlatarak yaz
                            result_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)
                        else:
                            # Depo adı yoksa normal şekilde yaz
                            result_df.to_excel(writer, sheet_name='Sheet1', index=False)
                    
                    saved_files.append(f"Excel: {excel_path}")
                    logging.info(f"Excel dosyası kaydedildi: {excel_path}")
                    
                except PermissionError:
                    if self.ui:
                        self.ui.show_error("Hata", f"Excel dosyası kaydetme izni yok: {excel_path}")
                except Exception as e:
                    if self.ui:
                        self.ui.show_error("Hata", f"Excel dosyası kaydedilemedi: {str(e)}")
            
            # Resim formatında kaydet
            if save_image:
                image_path = output_path + ".png"
                
                result = self.save_results_as_image(unique_cari_unvan_list, image_path, depo_name)
                if result is True:
                    saved_files.append(f"Resim: {image_path}")
                    logging.info(f"Resim dosyası kaydedildi: {image_path}")
                elif isinstance(result, tuple) and not result[0]:
                    if self.ui:
                        self.ui.show_error("Hata", result[1])
                else:
                    if self.ui:
                        self.ui.show_error("Hata", "Resim dosyası kaydedilirken bir hata oluştu.")
            
            # Başarı mesajı
            if saved_files:
                message = "Sonuçlar kaydedildi:\n" + "\n".join(saved_files)
                if self.ui:
                    self.ui.show_info("Bilgi", message)
            
            # Hiçbiri seçili değilse uyarı ver
            if not save_excel and not save_image:
                if self.ui:
                    self.ui.show_warning("Uyarı", "Lütfen en az bir kaydetme formatı seçin!")
                    
        except Exception as e:
            logging.error(f"Sonuç kaydetme hatası: {str(e)}")
            if self.ui:
                self.ui.show_error("Hata", f"Sonuçlar kaydedilirken hata oluştu: {str(e)}")
    
    def clear_results(self):
        """Sonuç listesini temizler"""
        if self.ui:
            self.ui.clear_results()


class ExcelComparisonApp:
    """Ana uygulama sınıfı - Drag & Drop desteği ile"""
    def __init__(self, root):
        self.root = root
        
        # İş mantığını oluştur
        self.logic = ExcelComparisonLogic()
        
        # Modern UI'ı oluştur
        self.ui = ModernExcelComparisonUI(root, self.logic)
        
        # UI referansını logic'e ver
        self.logic.set_ui(self.ui)


def main():
    """Ana program fonksiyonu - Drag & Drop desteği ile"""
    try:
        # Önce tkinterdnd2'yi dene
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
        
        # Uygulamayı başlat
        app = ExcelComparisonApp(root)
        
        # Ana döngüyü başlat
        root.mainloop()
        
    except ImportError:
        # tkinterdnd2 yoksa normal tkinter kullan
        messagebox.showwarning(
            "Bilgi", 
            "Drag & Drop özelliği için 'tkinterdnd2' kütüphanesini yükleyin:\n\npip install tkinterdnd2\n\nŞimdilik normal gözat butonlarıyla devam ediliyor."
        )
        
        # Normal tkinter ile çalıştır
        root = tk.Tk()
        app = ExcelComparisonApp(root)
        root.mainloop()
        
    except Exception as e:
        logging.critical(f"Uygulama başlatma hatası: {str(e)}")
        messagebox.showerror("Kritik Hata", f"Uygulama başlatılırken hata oluştu: {str(e)}")


if __name__ == "__main__":
    main()