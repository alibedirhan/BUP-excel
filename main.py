import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
import re
from datetime import datetime
import matplotlib.pyplot as plt
from ui import ModernExcelComparisonUI

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
        
    def set_ui(self, ui):
        """UI referansını ayarla"""
        self.ui = ui
        
    def find_header_row(self, df):
        """DataFrame içinde başlık satırını bulur"""
        for i, row in df.iterrows():
            for value in row.values:
                if isinstance(value, str) and "Cari Ünvan" in value:
                    return i
        return -1
    
    def update_output_filename(self, file_path):
        """Seçilen dosyaya göre çıktı dosya adını günceller"""
        try:
            # Excel dosyasını oku
            df = pd.read_excel(file_path, header=None, nrows=10)
            
            # Depo adını bul
            depo_name = None
            for i in range(min(10, len(df))):
                row_str = df.iloc[i, 0] if len(df.columns) > 0 else None
                if isinstance(row_str, str) and "Depo Kartı" in row_str:
                    # [xxxx] DEPO ADI formatını bul
                    match = re.search(r'\[(.*?)\]\s*(.*?)(?=\n|\r\n|$)', row_str)
                    if match and match.group(2):
                        depo_name = match.group(2).strip()
                        break
            
            # Depo adı bulunduysa çıktı dosya adını güncelle
            if depo_name:
                # Dosya adında kullanılamayacak karakterleri temizle
                safe_depo_name = re.sub(r'[\\/*?:"<>|]', '', depo_name)
                
                # Uzantı olmadan dosya adı
                output_filename = safe_depo_name
                self.output_path.set(output_filename)
        except Exception as e:
            # Hata olursa varsayılan dosya adını kullan
            print(f"Dosya adı güncelleme hatası: {str(e)}")
    
    def save_results_as_image(self, unique_cari_unvan_list, output_path, depo_name=None):
        """Sonuçları resim dosyası olarak kaydeder"""
        try:
            # Matplotlib figürü oluştur
            plt.figure(figsize=(10, 8))
            
            # Sadece depo adını başlık olarak ekle
            if depo_name:
                plt.suptitle(f"{depo_name}", fontsize=16, fontweight='bold')
            else:
                plt.suptitle("Eksik Cari Ünvanlar", fontsize=16, fontweight='bold')
                
            # Tablo verilerini oluştur
            cell_text = []
            for i, unvan in enumerate(unique_cari_unvan_list, 1):
                cell_text.append([i, unvan])
                
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
            
            # Dosyayı kaydet
            plt.savefig(output_path, bbox_inches='tight', dpi=150)
            plt.close()
            
            return True
        except Exception as e:
            print(f"Resim kaydetme hatası: {str(e)}")
            return False
    
    def compare_files(self):
        """Excel dosyalarını karşılaştırır ve sonuçları gösterir"""
        # Önce girdileri kontrol et
        file1_path = self.file1_path.get()
        file2_path = self.file2_path.get()
        output_path = self.output_path.get()
        
        if not file1_path or not file2_path:
            if self.ui:
                self.ui.show_error("Hata", "Lütfen her iki Excel dosyasını da seçin!")
            return
        
        if not os.path.exists(file1_path):
            if self.ui:
                self.ui.show_error("Hata", f"'{file1_path}' dosyası bulunamadı!")
            return
            
        if not os.path.exists(file2_path):
            if self.ui:
                self.ui.show_error("Hata", f"'{file2_path}' dosyası bulunamadı!")
            return
        
        # Sonuç listesini temizle
        self.clear_results()
        
        try:
            # Excel dosyalarını oku (başlık satırını bilmediğimiz için tüm içeriği okuyoruz)
            df1_full = pd.read_excel(file1_path, header=None)
            df2_full = pd.read_excel(file2_path, header=None)
            
            # Depo Kartı bilgisini çıkar
            depo_name = None
            for i in range(min(10, len(df1_full))):
                row_str = df1_full.iloc[i, 0] if len(df1_full.columns) > 0 else None
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
            cari_unvan_list1 = df1[cari_unvan_col1].dropna().apply(lambda x: x.strip() if isinstance(x, str) else x).tolist()
            cari_unvan_list2 = df2[cari_unvan_col2].dropna().apply(lambda x: x.strip() if isinstance(x, str) else x).tolist()
            
            # Büyük/küçük harf duyarlılığını devre dışı bırakma seçeneği
            if not self.case_sensitive.get():
                cari_unvan_list1_upper = [unvan.upper() if isinstance(unvan, str) else unvan for unvan in cari_unvan_list1]
                cari_unvan_list2_upper = [unvan.upper() if isinstance(unvan, str) else unvan for unvan in cari_unvan_list2]
                
                # Birinci dosyada olup ikinci dosyada olmayan cari ünvanları bul (büyük/küçük harf duyarsız)
                unique_indices = [i for i, unvan in enumerate(cari_unvan_list1_upper) 
                              if unvan not in cari_unvan_list2_upper]
                
                # Orijinal metinleri al
                unique_cari_unvan_list = [cari_unvan_list1[i] for i in unique_indices]
            else:
                # Birinci dosyada olup ikinci dosyada olmayan cari ünvanları bul (büyük/küçük harf duyarlı)
                unique_cari_unvan_list = [unvan for unvan in cari_unvan_list1 if unvan not in cari_unvan_list2]
            
            # Sonuçları UI'ya gönder
            status_text = f"Toplam {len(cari_unvan_list1)} cari ünvandan {len(unique_cari_unvan_list)} tanesi yeni dosyada bulunmuyor."
            if self.ui:
                self.ui.update_results(unique_cari_unvan_list, status_text)
            
            # Sonuçları kaydet
            if output_path:
                # UI'dan checkbox değerlerini al
                save_excel = self.ui.save_excel.get() if self.ui else True
                save_image = self.ui.save_image.get() if self.ui else False
                
                # Excel formatında kaydet
                if save_excel:
                    excel_path = output_path + ".xlsx"
                    
                    # Excel dosyasını oluştur
                    result_df = pd.DataFrame({"Cari Ünvan": unique_cari_unvan_list})
                    
                    # Excel dosyasını oluştur fakat hemen kaydetme
                    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
                    
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
                    
                    # Değişiklikleri kaydet
                    writer.close()
                    
                    if self.ui:
                        self.ui.show_info("Bilgi", f"Sonuçlar '{excel_path}' Excel dosyasına kaydedildi.")
                
                # Resim formatında kaydet
                if save_image:
                    image_path = output_path + ".png"
                    
                    # Resim dosyasını oluştur ve kaydet
                    if self.save_results_as_image(unique_cari_unvan_list, image_path, depo_name):
                        if self.ui:
                            self.ui.show_info("Bilgi", f"Sonuçlar '{image_path}' resim dosyasına kaydedildi.")
                    else:
                        if self.ui:
                            self.ui.show_error("Hata", "Resim dosyası kaydedilirken bir hata oluştu.")
                
                # Hiçbiri seçili değilse uyarı ver
                if not save_excel and not save_image:
                    if self.ui:
                        self.ui.show_warning("Uyarı", "Lütfen en az bir kaydetme formatı seçin!")
        
        except Exception as e:
            if self.ui:
                self.ui.show_error("Hata", f"İşlem sırasında bir hata oluştu: {str(e)}")
    
    def clear_results(self):
        """Sonuç listesini temizler"""
        if self.ui:
            self.ui.clear_results()


class ExcelComparisonApp:
    """Ana uygulama sınıfı"""
    def __init__(self, root):
        self.root = root
        
        # İş mantığını oluştur
        self.logic = ExcelComparisonLogic()
        
        # Modern UI'ı oluştur
        self.ui = ModernExcelComparisonUI(root, self.logic)
        
        # UI referansını logic'e ver
        self.logic.set_ui(self.ui)


def main():
    """Ana program fonksiyonu"""
    try:
        # Ana pencereyi oluştur
        root = tk.Tk()
        
        # Uygulamayı başlat
        app = ExcelComparisonApp(root)
        
        # Ana döngüyü başlat
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Kritik Hata", f"Uygulama başlatılırken hata oluştu: {str(e)}")


if __name__ == "__main__":
    main()