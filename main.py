import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import os
import re
import json
import sys
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import threading
import logging

# UI import kontrolü
try:
    from ui import ModernExcelComparisonUI
except ImportError as e:
    print("HATA: ui.py dosyası bulunamadı veya import edilemedi!")
    print(f"Detay: {str(e)}")
    print("Lütfen ui.py dosyasının aynı dizinde olduğundan emin olun.")
    sys.exit(1)

# Logging sistemi kurulumu
try:
    logging.basicConfig(
        filename='app.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='utf-8'
    )
except Exception:
    logging.basicConfig(
        filename='app.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

class VehicleDriverSetupDialog:
    """Araç-Plasiyer Eşleştirme Dialog'u"""
    def __init__(self, parent, existing_data=None):
        self.parent = parent
        self.existing_data = existing_data or {}
        self.result = {}
        self.dialog = None
        self.entries = {}
        
    def show_setup_dialog(self):
        """Araç-plasiyer eşleştirme dialog'unu göster"""
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
        
        # Başlık
        title_label = tk.Label(
            main_frame,
            text="Araç-Plasiyer Eşleştirmesi",
            font=('Segoe UI', 8, 'bold')
        )
        title_label.pack(pady=(0, 10))
        
        # Açıklama
        desc_label = tk.Label(
            main_frame,
            text="Lütfen her araç numarası için plasiyer adını girin.\nBoş bırakılan araçlar kullanılmayacak.",
            font=('Segoe UI', 8),
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
        
        # 20 araç için entry oluştur
        for i in range(1, 21):
            vehicle_num = f"{i:02d}"
            
            vehicle_frame = tk.Frame(scrollable_frame)
            vehicle_frame.pack(fill=tk.X, pady=2)
            
            label = tk.Label(
                vehicle_frame,
                text=f"Araç {vehicle_num}:",
                width=10,
                anchor='w',
                font=('Segoe UI', 8)
            )
            label.pack(side=tk.LEFT)
            
            entry = tk.Entry(
                vehicle_frame, 
                width=30,
                font=('Segoe UI', 8)
            )
            entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
            self.entries[vehicle_num] = entry
            
            # Mevcut veriyi yükle (varsa)
            if vehicle_num in self.existing_data:
                entry.insert(0, self.existing_data[vehicle_num])
            
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Butonlar
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=15)
        
        # Kaydet butonu
        save_btn = tk.Button(
            button_frame,
            text="Kaydet",
            command=self.save_config,
            bg="#4CAF50",
            fg="white",
            font=('Segoe UI', 8, 'bold'),
            padx=20
        )
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        # İptal butonu
        cancel_btn = tk.Button(
            button_frame,
            text="İptal",
            command=self.cancel,
            bg="#f44336",
            fg="white",
            font=('Segoe UI', 8),
            padx=20
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Örnek veriler yükle
        load_sample_btn = tk.Button(
            button_frame,
            text="Örnek Veriler",
            command=self.load_sample_data,
            bg="#2196F3",
            fg="white",
            font=('Segoe UI', 8),
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
        vehicle_drivers = {}
        for vehicle_num, entry in self.entries.items():
            driver_name = entry.get().strip()
            if driver_name:
                vehicle_drivers[vehicle_num] = driver_name
        
        if not vehicle_drivers:
            messagebox.showwarning("Uyarı", "En az bir araç-plasiyer eşleştirmesi yapmalısınız!")
            return
        
        config = {"vehicle_drivers": vehicle_drivers}
        
        try:
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
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.case_sensitive = tk.BooleanVar(value=False)
        self.save_format = tk.StringVar(value="excel")
        self.output_path.set("")
        self.ui = None
        self.max_file_size_mb = 100
        self.vehicle_drivers = self.load_vehicle_drivers()
        
    def load_vehicle_drivers(self):
        """Araç-plasiyer eşleştirmesini dosyadan yükler"""
        try:
            config_files = ['config.json', 'vehicle_config.json', 'drivers.json']
            
            for config_file in config_files:
                if os.path.exists(config_file):
                    with open(config_file, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                        vehicle_drivers = config.get('vehicle_drivers', {})
                        if vehicle_drivers:
                            logging.info(f"Araç-plasiyer konfigürasyonu yüklendi: {config_file}")
                            return vehicle_drivers
            
            env_config = {}
            for i in range(1, 21):
                key = f"DRIVER_{i:02d}"
                if key in os.environ:
                    env_config[f"{i:02d}"] = os.environ[key]
            
            if env_config:
                logging.info("Araç-plasiyer konfigürasyonu çevre değişkenlerinden yüklendi")
                return env_config
                
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
        if not self.ui or not hasattr(self.ui, 'root') or not self.ui.root:
            logging.error("UI referansı bulunamadı veya root widget yok")
            return False
            
        try:
            dialog = VehicleDriverSetupDialog(self.ui.root, self.vehicle_drivers)
            result = dialog.show_setup_dialog()
            
            if result:
                self.vehicle_drivers = result
                return True
            return False
        except Exception as e:
            logging.error(f"Vehicle setup dialog hatası: {str(e)}")
            return False
    
    def set_ui(self, ui):
        """UI referansını ayarla"""
        self.ui = ui
        
        if not self.vehicle_drivers and ui:
            response = messagebox.askyesno(
                "Araç-Plasiyer Eşleştirmesi",
                "Araç-plasiyer eşleştirmesi bulunamadı.\n\n"
                "Şimdi eşleştirme yapmak ister misiniz?\n\n"
                "Evet: Eşleştirme ekranını aç\n"
                "Hayır: Varsayılan örnek config oluştur"
            )
            
            if response:
                if self.show_vehicle_setup_dialog():
                    logging.info("Kullanıcı araç-plasiyer eşleştirmesi yaptı")
                else:
                    logging.info("Kullanıcı araç-plasiyer eşleştirmesini iptal etti")
            else:
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
        """DataFrame içinde başlık satırını bulur"""
        try:
            df_str = df.astype(str)
            mask = df_str.map(lambda x: "Cari Ünvan" in x if isinstance(x, str) else False).any(axis=1)
            
            if mask.any():
                return mask.idxmax()
            return -1
        except Exception as e:
            logging.error(f"Başlık satırı bulma hatası: {str(e)}")
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
                
            logging.info(f"Araç numarası çıkarma denemesi: '{depo_text}'")
                
            patterns = [
                r'[İI][Zz][Mm][İi][Rr]\s+[Aa][Rr][Aa][Çç]\s+(\d{1,2})',
                r'[İI][Zz][Mm][İi][Rr]\s+[Aa][Rr][Aa][Cc]\s+(\d{1,2})',
                r'[Aa]raç\s*(\d{1,2})',
                r'[Aa]rac\s*(\d{1,2})',
                r'[Vv]ehicle\s*(\d{1,2})',
                r'(\d{1,2})\s*[Nn]o',
                r'(\d{1,2})\s*[Nn]olu',
                r'\b(\d{1,2})\b'
            ]
            
            for i, pattern in enumerate(patterns):
                match = re.search(pattern, depo_text)
                if match:
                    vehicle_num = match.group(1).zfill(2)
                    logging.info(f"Pattern {i+1} ile eşleşti. Araç numarası: {vehicle_num}")
                    
                    if vehicle_num in self.vehicle_drivers:
                        driver_name = self.vehicle_drivers[vehicle_num]
                        logging.info(f"Araç {vehicle_num} → Plasiyer: {driver_name}")
                        return vehicle_num
                    else:
                        logging.warning(f"Araç {vehicle_num} config'de bulunamadı")
                        
            logging.warning(f"Hiçbir pattern eşleşmedi: '{depo_text}'")
            return None
            
        except Exception as e:
            logging.error(f"Araç numarası çıkarma hatası: {str(e)}")
            return None
    
    def create_filename_with_driver(self, depo_text):
        """Depo kartından araç numarası çıkarıp plasiyer adıyla dosya adı oluşturur"""
        try:
            vehicle_num = self.extract_vehicle_number(depo_text)
            
            if vehicle_num and vehicle_num in self.vehicle_drivers:
                driver_name = self.vehicle_drivers[vehicle_num]
                filename = f"Arac_{vehicle_num}_{driver_name}"
                return self.sanitize_filename(filename)
            else:
                return self.sanitize_filename(depo_text) if depo_text else ""
                
        except Exception as e:
            logging.error(f"Plasiyerli dosya adı oluşturma hatası: {str(e)}")
            return self.sanitize_filename(depo_text) if depo_text else ""
    
    def sanitize_filename(self, filename):
        """Dosya adını güvenli hale getirir"""
        invalid_chars = r'[\\/*?:"<>|]'
        safe_name = re.sub(invalid_chars, '', filename)
        safe_name = safe_name.strip()
        
        if len(safe_name) > 100:
            safe_name = safe_name[:100]
            
        return safe_name if safe_name else "output"
    
    def update_output_filename(self, file_path):
        """Seçilen dosyaya göre çıktı dosya adını günceller"""
        logging.info(f"🔍 update_output_filename çağrıldı: {file_path}")
        
        try:
            is_valid, error_msg = self.validate_file_size(file_path)
            if not is_valid:
                logging.warning(f"🔍 Dosya boyutu hatası: {error_msg}")
                if self.ui:
                    self.ui.show_warning("Uyarı", error_msg)
                default_name = f"output_{datetime.now().strftime('%H%M%S')}"
                self.output_path.set(default_name)
                return
                
            df = pd.read_excel(file_path, header=None, nrows=10)
            logging.info(f"🔍 Excel dosyası okundu, {len(df)} satır")
            
            depo_name = None
            for i in range(min(10, len(df))):
                row_str = str(df.iloc[i, 0]) if len(df.columns) > 0 else None
                logging.info(f"🔍 Satır {i}: '{row_str}'")
                
                if isinstance(row_str, str) and "Cari Kategori 3" in row_str:
                    logging.info(f"🔍 CARİ KATEGORİ 3 BULUNDU: {row_str}")
                    match = re.search(r'\[(.*?)\]\s*(.*?)(?:\n|\r\n|$)', row_str)
                    if match and match.group(2):
                        depo_name = match.group(2).strip()
                        logging.info(f"🔍 Araç adı çıkarıldı: '{depo_name}'")
                        break
                    else:
                        logging.warning(f"🔍 Regex eşleşmedi, ham metin: {repr(row_str)}")
            
            if not depo_name:
                logging.warning("🔍 Hiçbir satırda 'Cari Kategori 3' bulunamadı")
            
            if depo_name:
                logging.info(f"🔍 create_filename_with_driver çağrılıyor: '{depo_name}'")
                filename_with_driver = self.create_filename_with_driver(depo_name)
                logging.info(f"🔍 Oluşturulan dosya adı: '{filename_with_driver}'")
                self.output_path.set(filename_with_driver)
            else:
                default_name = f"karsilastirma_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                logging.info(f"🔍 Varsayılan ad kullanılıyor: '{default_name}'")
                self.output_path.set(default_name)
                
        except PermissionError:
            logging.error(f"🔍 İzin hatası: {file_path}")
            if self.ui:
                self.ui.show_error("Hata", "Dosyaya erişim izni yok!")
            default_name = f"output_{datetime.now().strftime('%H%M%S')}"
            self.output_path.set(default_name)
        except pd.errors.EmptyDataError:
            logging.error(f"🔍 Boş dosya hatası: {file_path}")
            if self.ui:
                self.ui.show_error("Hata", "Excel dosyası boş veya bozuk!")
            default_name = f"output_{datetime.now().strftime('%H%M%S')}"
            self.output_path.set(default_name)
        except Exception as e:
            default_name = f"karsilastirma_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.output_path.set(default_name)
            logging.error(f"🔍 Genel hata: {str(e)}")
            if self.ui:
                self.ui.show_warning("Uyarı", f"Dosya adı güncellenemedi, varsayılan ad kullanılıyor: {default_name}")
    
    def save_results_as_image(self, unique_cari_unvan_list, output_path, depo_name=None):
        """Sonuçları resim dosyası olarak kaydeder"""
        try:
            plt.figure(figsize=(12, 8), dpi=150)
            plt.rcParams['font.family'] = 'DejaVu Sans'
            
            if depo_name:
                vehicle_num = self.extract_vehicle_number(depo_name)
                if vehicle_num and vehicle_num in self.vehicle_drivers:
                    driver_name = self.vehicle_drivers[vehicle_num]
                    title = f"Araç {vehicle_num} - {driver_name}"
                else:
                    title = depo_name
                plt.suptitle(title, fontsize=16, fontweight='bold')
            else:
                plt.suptitle("Eksik Cari Ünvanlar", fontsize=16, fontweight='bold')
                
            cell_text = []
            for i, unvan in enumerate(unique_cari_unvan_list, 1):
                display_unvan = unvan if len(str(unvan)) <= 80 else str(unvan)[:77] + "..."
                cell_text.append([i, display_unvan])
                
            if not cell_text:
                cell_text = [["", "Tüm cari ünvanlar her iki dosyada da mevcut."]]
                
            plt.axis('off')
            table = plt.table(
                cellText=cell_text,
                colLabels=["#", "Cari Ünvan"],
                loc='center',
                cellLoc='left',
                colWidths=[0.1, 0.9]
            )
            
            table.auto_set_font_size(False)
            table.set_fontsize(9)
            table.scale(1, 1.5)
            
            for (i, j), cell in table.get_celld().items():
                if i == 0:
                    cell.set_text_props(fontweight='bold')
                    cell.set_facecolor('#e6e6e6')
                else:
                    if i % 2 == 0:
                        cell.set_facecolor('#f9f9f9')
            
            full_output_path = os.path.abspath(output_path)
            
            output_dir = os.path.dirname(full_output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            plt.savefig(full_output_path, bbox_inches='tight', dpi=150, 
                       facecolor='white', edgecolor='none')
            plt.close()
            
            return True, full_output_path
            
        except PermissionError:
            error_msg = f"Resim kaydetme izin hatası: {output_path}"
            logging.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"Resim kaydetme hatası: {str(e)}"
            logging.error(error_msg)
            return False, error_msg
    
    def validate_excel_file(self, file_path):
        """Excel dosyasının geçerli olup olmadığını kontrol eder"""
        try:
            if not os.path.exists(file_path):
                return False, "Dosya bulunamadı!"
                
            is_valid, error_msg = self.validate_file_size(file_path)
            if not is_valid:
                return False, error_msg
                
            pd.read_excel(file_path, nrows=1)
            return True, ""
            
        except PermissionError:
            return False, "Dosyaya erişim izni yok!"
        except pd.errors.EmptyDataError:
            return False, "Excel dosyası boş!"
        except Exception as e:
            return False, f"Geçersiz Excel dosyası: {str(e)}"
    
    def compare_files_thread(self):
        """Dosya karşılaştırmasını ayrı thread'de çalıştır"""
        try:
            self.compare_files_internal()
        except Exception as e:
            logging.error(f"Thread hatası: {str(e)}")
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", f"İşlem sırasında beklenmeyen hata: {str(e)}"))
    
    def compare_files(self):
        """Excel dosyalarını karşılaştırır - Ana fonksiyon"""
        thread = threading.Thread(target=self.compare_files_thread)
        thread.daemon = True
        thread.start()
    
    def compare_files_internal(self):
        """Excel dosylarını karşılaştırır ve sonuçları gösterir"""
        file1_path = self.file1_path.get()
        file2_path = self.file2_path.get()
        output_path = self.output_path.get()
        
        if not file1_path or not file2_path:
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", "Lütfen her iki Excel dosyasını da seçin!"))
            return
        
        is_valid1, error1 = self.validate_excel_file(file1_path)
        if not is_valid1:
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", f"Eski tarihli dosya hatası: {error1}"))
            return
            
        is_valid2, error2 = self.validate_excel_file(file2_path)
        if not is_valid2:
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", f"Yeni tarihli dosya hatası: {error2}"))
            return
        
        self.clear_results()
        
        try:
            logging.info(f"Dosyalar okunuyor: {file1_path}, {file2_path}")
            
            df1_full = pd.read_excel(file1_path, header=None)
            df2_full = pd.read_excel(file2_path, header=None)
            
            depo_name = None
            for i in range(min(10, len(df1_full))):
                row_str = str(df1_full.iloc[i, 0]) if len(df1_full.columns) > 0 else None
                if isinstance(row_str, str) and "Cari Kategori 3" in row_str:
                    match = re.search(r'\[(.*?)\]\s*(.*?)(?:\n|\r\n|$)', row_str)
                    if match and match.group(2):
                        depo_name = match.group(2).strip()
                        logging.info(f"🔍 Karşılaştırmada araç adı bulundu: '{depo_name}'")
                        break
            
            header_row1 = self.find_header_row(df1_full)
            header_row2 = self.find_header_row(df2_full)
            
            if header_row1 == -1 or header_row2 == -1:
                if self.ui:
                    self.ui.root.after(0, lambda: self.ui.show_error("Hata", "Excel dosyalarında 'Cari Ünvan' başlığı bulunamadı!"))
                return
            
            df1 = pd.read_excel(file1_path, header=header_row1)
            df2 = pd.read_excel(file2_path, header=header_row2)
            
            df1.columns = [col.strip() if isinstance(col, str) else col for col in df1.columns]
            df2.columns = [col.strip() if isinstance(col, str) else col for col in df2.columns]
            
            cari_unvan_col1 = next((col for col in df1.columns if isinstance(col, str) and "Cari Ünvan" in col), None)
            cari_unvan_col2 = next((col for col in df2.columns if isinstance(col, str) and "Cari Ünvan" in col), None)
            
            if not cari_unvan_col1 or not cari_unvan_col2:
                if self.ui:
                    self.ui.root.after(0, lambda: self.ui.show_error("Hata", "Bir veya daha fazla Excel dosyasında 'Cari Ünvan' sütunu bulunamadı."))
                return
            
            cari_unvan_list1 = df1[cari_unvan_col1].dropna().apply(
                lambda x: x.strip() if isinstance(x, str) else str(x).strip()
            ).tolist()
            cari_unvan_list2 = df2[cari_unvan_col2].dropna().apply(
                lambda x: x.strip() if isinstance(x, str) else str(x).strip()
            ).tolist()
            
            cari_unvan_list1 = [x for x in cari_unvan_list1 if x and x.strip()]
            cari_unvan_list2 = [x for x in cari_unvan_list2 if x and x.strip()]
            
            if not self.case_sensitive.get():
                cari_unvan_list1_upper = [unvan.upper() for unvan in cari_unvan_list1]
                cari_unvan_list2_upper = [unvan.upper() for unvan in cari_unvan_list2]
                
                unique_cari_unvan_list = [
                    cari_unvan_list1[i] for i, unvan in enumerate(cari_unvan_list1_upper) 
                    if unvan not in cari_unvan_list2_upper
                ]
            else:
                unique_cari_unvan_list = [unvan for unvan in cari_unvan_list1 if unvan not in cari_unvan_list2]
            
            seen = set()
            unique_cari_unvan_list = [x for x in unique_cari_unvan_list if not (x in seen or seen.add(x))]
            
            status_text = f"Toplam {len(cari_unvan_list1)} cari ünvandan {len(unique_cari_unvan_list)} tanesi yeni dosyada bulunmuyor."
            if self.ui:
                self.ui.update_results(unique_cari_unvan_list, status_text)
            
            logging.info(f"Karşılaştırma tamamlandı. {len(unique_cari_unvan_list)} farklılık bulundu.")
            
            self.save_results(unique_cari_unvan_list, output_path, depo_name)
        
        except MemoryError:
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", "Dosyalar çok büyük, bellek yetersiz!"))
        except pd.errors.EmptyDataError:
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", "Excel dosyalarından biri boş veya bozuk!"))
        except PermissionError:
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", "Dosyalara erişim izni yok!"))
        except Exception as e:
            logging.error(f"Karşılaştırma hatası: {str(e)}")
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", f"İşlem sırasında bir hata oluştu: {str(e)}"))
    
    def save_results(self, unique_cari_unvan_list, output_path, depo_name):
        """Sonuçları kaydet"""
        if not output_path or output_path.strip() == "":
            output_path = f"karsilastirma_sonucu_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            logging.warning(f"Output path boş, varsayılan oluşturuldu: {output_path}")
            
        current_dir = os.getcwd()
        logging.info(f"Çalışma dizini: {current_dir}")
        logging.info(f"Output path: {output_path}")
        logging.info(f"Sonuç listesi uzunluğu: {len(unique_cari_unvan_list)}")
            
        try:
            save_excel = self.ui.save_excel.get() if self.ui else True
            save_image = self.ui.save_image.get() if self.ui else False
            
            logging.info(f"Save Excel: {save_excel}, Save Image: {save_image}")
            
            saved_files = []
            
            if save_excel:
                excel_path = os.path.join(current_dir, output_path + ".xlsx")
                logging.info(f"Excel dosyası kaydediliyor: {excel_path}")
                
                try:
                    excel_dir = os.path.dirname(excel_path)
                    if excel_dir and not os.path.exists(excel_dir):
                        os.makedirs(excel_dir, exist_ok=True)
                        logging.info(f"Dizin oluşturuldu: {excel_dir}")
                    
                    table_data = []
                    for i, unvan in enumerate(unique_cari_unvan_list, 1):
                        table_data.append([i, unvan])
                    
                    result_df = pd.DataFrame(table_data, columns=["#", "Cari Ünvan"])
                    
                    try:
                        from openpyxl.styles import Font, Border, Side, Alignment
                        
                        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                            if depo_name:
                                vehicle_num = self.extract_vehicle_number(depo_name)
                                if vehicle_num and vehicle_num in self.vehicle_drivers:
                                    driver_name = self.vehicle_drivers[vehicle_num]
                                    header_text = f"Araç {vehicle_num} - {driver_name}"
                                else:
                                    header_text = depo_name
                                
                                header_df = pd.DataFrame({"A": [header_text], "B": [""]})
                                header_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0)
                                
                                result_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=2)
                                
                                workbook = writer.book
                                worksheet = writer.sheets['Sheet1']
                                
                                bold_font = Font(bold=True, color="000000", size=12)
                                header_font = Font(bold=True, color="000000", size=10)
                                normal_font = Font(color="000000", size=10)
                                
                                thin_border = Border(
                                    left=Side(style='thin'),
                                    right=Side(style='thin'),
                                    top=Side(style='thin'),
                                    bottom=Side(style='thin')
                                )
                                
                                center_alignment = Alignment(horizontal='center', vertical='center')
                                left_alignment = Alignment(horizontal='left', vertical='center')
                                
                                worksheet.merge_cells('A1:B1')
                                worksheet['A1'] = header_text
                                worksheet['A1'].font = bold_font
                                worksheet['A1'].alignment = center_alignment
                                worksheet['A1'].border = thin_border
                                
                                worksheet['A3'].font = header_font
                                worksheet['A3'].alignment = center_alignment
                                worksheet['A3'].border = thin_border
                                
                                worksheet['B3'].font = header_font
                                worksheet['B3'].alignment = center_alignment
                                worksheet['B3'].border = thin_border
                                
                                for row in range(4, len(unique_cari_unvan_list) + 4):
                                    worksheet[f'A{row}'].font = normal_font
                                    worksheet[f'A{row}'].alignment = center_alignment
                                    worksheet[f'A{row}'].border = thin_border
                                    
                                    worksheet[f'B{row}'].font = normal_font
                                    worksheet[f'B{row}'].alignment = left_alignment
                                    worksheet[f'B{row}'].border = thin_border
                                
                                worksheet.column_dimensions['A'].width = 8
                                worksheet.column_dimensions['B'].width = 60
                                
                            else:
                                result_df.to_excel(writer, sheet_name='Sheet1', index=False)
                    
                    except ImportError:
                        logging.warning("openpyxl.styles import edilemedi, basit format kullanılıyor")
                        result_df.to_excel(excel_path, index=False)
                    
                    saved_files.append(f"Excel: {excel_path}")
                    logging.info(f"Excel dosyası başarıyla kaydedildi: {excel_path}")
                    
                except PermissionError as e:
                    error_msg = f"Excel dosyası kaydetme izni yok: {excel_path}\nHata: {str(e)}"
                    logging.error(error_msg)
                    if self.ui:
                        self.ui.root.after(0, lambda: self.ui.show_error("Hata", error_msg))
                except Exception as e:
                    error_msg = f"Excel dosyası kaydedilemedi: {str(e)}"
                    logging.error(error_msg)
                    if self.ui:
                        self.ui.root.after(0, lambda: self.ui.show_error("Hata", error_msg))
            
            if save_image:
                image_path = os.path.join(current_dir, output_path + ".png")
                logging.info(f"Resim dosyası kaydediliyor: {image_path}")
                
                success, result_msg = self.save_results_as_image(unique_cari_unvan_list, image_path, depo_name)
                if success:
                    saved_files.append(f"Resim: {result_msg}")
                    logging.info(f"Resim dosyası başarıyla kaydedildi: {result_msg}")
                else:
                    error_msg = f"Resim dosyası kaydedilemedi: {result_msg}"
                    logging.error(error_msg)
                    if self.ui:
                        self.ui.root.after(0, lambda: self.ui.show_error("Hata", error_msg))
            
            if saved_files:
                success_message = "Sonuçlar başarıyla kaydedildi:\n\n" + "\n".join(saved_files)
                logging.info(success_message)
                if self.ui:
                    self.ui.root.after(0, lambda: self.ui.show_info("Başarılı", success_message))
            elif not save_excel and not save_image:
                warning_msg = "Lütfen en az bir kaydetme formatı seçin (Excel veya Resim)!"
                logging.warning(warning_msg)
                if self.ui:
                    self.ui.root.after(0, lambda: self.ui.show_warning("Uyarı", warning_msg))
            else:
                error_msg = "Hiçbir dosya kaydedilemedi. Lütfen log dosyasını kontrol edin."
                logging.error(error_msg)
                if self.ui:
                    self.ui.root.after(0, lambda: self.ui.show_error("Hata", error_msg))
                    
        except Exception as e:
            error_msg = f"Sonuç kaydetme genel hatası: {str(e)}"
            logging.error(error_msg)
            if self.ui:
                self.ui.root.after(0, lambda: self.ui.show_error("Hata", error_msg))
    
    def clear_results(self):
        """Sonuç listesini temizler"""
        if self.ui:
            self.ui.clear_results()


class ExcelComparisonApp:
    """Ana uygulama sınıfı"""
    def __init__(self, root):
        self.root = root
        self.logic = ExcelComparisonLogic()
        self.ui = ModernExcelComparisonUI(root, self.logic)
        self.logic.set_ui(self.ui)


def main():
    """Ana program fonksiyonu"""
    try:
        has_dnd = False
        try:
            from tkinterdnd2 import TkinterDnD
            root = TkinterDnD.Tk()
            has_dnd = True
            logging.info("tkinterdnd2 başarıyla yüklendi - Drag & Drop aktif")
        except ImportError:
            root = tk.Tk()
            has_dnd = False
            logging.warning("tkinterdnd2 bulunamadı - Normal mod aktif")
            messagebox.showwarning(
                "Bilgi", 
                "Drag & Drop özelliği için 'tkinterdnd2' kütüphanesini yükleyin:\n\n"
                "pip install tkinterdnd2\n\n"
                "Şimdilik normal gözat butonlarıyla devam ediliyor."
            )
        
        if not root:
            raise RuntimeError("Tkinter root window oluşturulamadı")
        
        try:
            app = ExcelComparisonApp(root)
            if not app.logic or not app.ui:
                raise RuntimeError("Uygulama bileşenleri başlatılamadı")
        except Exception as e:
            logging.error(f"Uygulama başlatma hatası: {str(e)}")
            messagebox.showerror(
                "Başlatma Hatası", 
                f"Uygulama başlatılamadı:\n{str(e)}\n\n"
                "Lütfen tüm dosyaların mevcut olduğundan emin olun."
            )
            return
        
        logging.info("Uygulama başarıyla başlatıldı")
        
        try:
            root.mainloop()
        except KeyboardInterrupt:
            logging.info("Uygulama kullanıcı tarafından sonlandırıldı")
        except Exception as e:
            logging.error(f"Ana döngü hatası: {str(e)}")
            messagebox.showerror("Çalışma Hatası", f"Uygulama çalışırken hata oluştu: {str(e)}")
        finally:
            try:
                if root:
                    root.quit()
                    root.destroy()
            except:
                pass
        
    except Exception as e:
        error_msg = f"Kritik uygulama hatası: {str(e)}"
        logging.critical(error_msg)
        try:
            messagebox.showerror("Kritik Hata", error_msg)
        except:
            print(error_msg)
        sys.exit(1)


if __name__ == "__main__":
    main()