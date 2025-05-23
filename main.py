import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
import re
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
        
        # Varsayılan çıktı dosyası adı
        self.output_path.set("eksik_cari_unvanlar")
        
        # UI referansı
        self.ui = None
        
        # Maximum dosya boyutu (MB)
        self.max_file_size_mb = 100
        
    def set_ui(self, ui):
        """UI referansını ayarla"""
        self.ui = ui
        
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
            
            # Depo adı bulunduysa çıktı dosya adını güncelle
            if depo_name:
                # Dosya adını güvenli hale getir
                safe_depo_name = self.sanitize_filename(depo_name)
                self.output_path.set(safe_depo_name)
                
        except PermissionError:
            if self.ui:
                self.ui.show_error("Hata", "Dosyaya erişim izni yok!")
        except pd.errors.EmptyDataError:
            if self.ui:
                self.ui.show_error("Hata", "Excel dosyası boş veya bozuk!")
        except Exception as e:
            # Hata olursa varsayılan dosya adını kullan
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
            
            # Sadece depo adını başlık olarak ekle
            if depo_name:
                plt.suptitle(f"{depo_name}", fontsize=16, fontweight='bold')
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
        """Excel dosyalarını karşılaştırır ve sonuçları gösterir - İç fonksiyon"""
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
                        # Depo kartı bilgisini üst satıra yaz
                        if depo_name:
                            # Yeni bir DataFrame oluştur - sadece depo adını içeren
                            depo_df = pd.DataFrame({" ": [depo_name]})
                            depo_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                            
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