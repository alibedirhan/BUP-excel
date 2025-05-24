import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading

class ModernExcelComparisonUI:
    def __init__(self, root, app_logic):
        self.root = root
        self.app_logic = app_logic
        
        # Modern tema ayarları
        self.setup_modern_theme()
        
        # Ana pencere ayarları
        self.setup_main_window()
        
        # Arayüz bileşenlerini oluştur
        self.create_modern_interface()
        
        # Progress bar için değişken
        self.progress_var = tk.BooleanVar()
        
    def setup_modern_theme(self):
        """Modern tema ve renkler"""
        self.colors = {
            'primary': '#2563eb',      # Modern mavi
            'primary_hover': '#1d4ed8', # Koyu mavi
            'secondary': '#64748b',     # Gri-mavi
            'success': '#10b981',       # Yeşil
            'warning': '#f59e0b',       # Turuncu
            'danger': '#ef4444',        # Kırmızı
            'background': '#f8fafc',    # Açık gri arka plan
            'card': '#ffffff',          # Beyaz kart
            'border': '#e2e8f0',       # Açık gri kenarlık
            'text': '#1e293b',         # Koyu gri metin
            'text_light': '#64748b',   # Açık gri metin
            'entry_normal': '#ffffff', # Normal entry arka planı
            'entry_hover': '#f0f9ff',  # Entry hover arka planı
            'entry_drop': '#e0f2fe'    # Entry drop arka planı
        }
        
        # Modern stil tanımlamaları
        style = ttk.Style()
        
        # Tema seç (mevcut temalar arasından en uygun olanı)
        available_themes = style.theme_names()
        if 'clam' in available_themes:
            style.theme_use('clam')
        elif 'alt' in available_themes:
            style.theme_use('alt')
        
        # Özel stil tanımlamaları
        self.configure_modern_styles(style)
        
    def configure_modern_styles(self, style):
        """Modern stil konfigürasyonları - TÜM FONTLAR 8px"""
        # Ana buton stili
        style.configure(
            'Modern.TButton',
            font=('Segoe UI', 8),  # 10 → 8
            padding=(20, 12),
            relief='flat',
            borderwidth=0
        )
        
        # Vurgulu buton stili
        style.configure(
            'Accent.TButton',
            font=('Segoe UI', 8, 'bold'),  # 10 → 8
            padding=(25, 15),
            relief='flat',
            borderwidth=0
        )
        
        # İkincil buton stili
        style.configure(
            'Secondary.TButton',
            font=('Segoe UI', 8),  # 10 → 8
            padding=(20, 12),
            relief='flat',
            borderwidth=1
        )
        
        # Küçük buton stili (Gözat butonları için)
        style.configure(
            'Small.TButton',
            font=('Segoe UI', 8),
            padding=(12, 8),
            relief='flat',
            borderwidth=1
        )
        
        # Küçük checkbutton stili (Seçenekler için)
        style.configure(
            'Small.TCheckbutton',
            font=('Segoe UI', 8),
            padding=(0, 3)
        )
        
        # LabelFrame stili
        style.configure(
            'Modern.TLabelframe',
            relief='flat',
            borderwidth=1,
            padding=20
        )
        
        style.configure(
            'Modern.TLabelframe.Label',
            font=('Segoe UI', 8, 'bold'),  # 11 → 8
            padding=(0, 5)
        )
        
        # Entry stili
        style.configure(
            'Modern.TEntry',
            font=('Segoe UI', 8),  # 10 → 8
            padding=10,
            relief='flat',
            borderwidth=1
        )
        
        # Drag & Drop Entry stili
        style.configure(
            'DragDrop.TEntry',
            font=('Segoe UI', 8),
            padding=8,
            relief='solid',
            borderwidth=2
        )
        
        # Checkbutton stili
        style.configure(
            'Modern.TCheckbutton',
            font=('Segoe UI', 8),  # 10 → 8
            padding=(0, 5)
        )
        
        # Radiobutton stili
        style.configure(
            'Modern.TRadiobutton',
            font=('Segoe UI', 8),  # 10 → 8
            padding=(0, 3)
        )
        
    def setup_main_window(self):
        """Ana pencere ayarları - Özel başlık ile"""
        # Pencere başlığını CAL olarak ayarla
        self.root.title("CAL")
        self.root.geometry("900x650")
        self.root.minsize(850, 600)
        
        # Pencereyi ortala
        self.center_window()
        
        # Arka plan rengi
        self.root.configure(bg=self.colors['background'])
        
        # Pencere kapatma protokolü
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # İkon ayarla (varsa)
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
            
    def on_closing(self):
        """Pencere kapatılırken çalışır"""
        # Aktif thread'ler varsa uyar
        active_threads = threading.active_count()
        if active_threads > 1:  # Ana thread + aktif thread'ler
            result = messagebox.askyesno(
                "Çıkış", 
                "İşlem devam ediyor. Çıkmak istediğinizden emin misiniz?"
            )
            if not result:
                return
        
        self.root.destroy()
            
    def center_window(self):
        """Pencereyi ekranın ortasına yerleştir"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_modern_interface(self):
        """Modern arayüz bileşenlerini oluştur - Özel başlık ile"""
        # Ana konteyner
        main_container = tk.Frame(self.root, bg=self.colors['background'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        
        # Özel başlık bölümü oluştur
        self.create_custom_header(main_container)
        
        # İçerik container'ı
        content_frame = tk.Frame(main_container, bg=self.colors['background'])
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # Sol panel (Dosya seçimi ve seçenekler)
        left_panel = tk.Frame(content_frame, bg=self.colors['background'])
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # Sağ panel (Sonuçlar)
        right_panel = tk.Frame(content_frame, bg=self.colors['background'])
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Sol paneli oluştur
        self.create_left_panel(left_panel)
        
        # Sağ paneli oluştur
        self.create_right_panel(right_panel)
        
    def create_custom_header(self, parent):
        """Özel başlık bölümü oluştur - Sadece Excel Karşılaştırma ortada"""
        header_frame = tk.Frame(parent, bg=self.colors['card'], height=50, relief='solid', bd=1)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        header_frame.pack_propagate(False)
        
        # Orta - Excel Karşılaştırma (tek başına ortada)
        center_frame = tk.Frame(header_frame, bg=self.colors['card'])
        center_frame.pack(expand=True, fill=tk.BOTH)
        
        title_label = tk.Label(
            center_frame,
            text="Excel Karşılaştırma",
            font=('Segoe UI', 14, 'bold'),  # Font boyutu 12 → 14 (biraz daha büyütüldü)
            bg=self.colors['card'],
            fg=self.colors['text']
        )
        title_label.pack(expand=True)
        
    def create_left_panel(self, parent):
        """Sol panel (Dosya seçimi ve seçenekler) - Uniform font"""
        # Sol panel için sabit genişlik ayarla
        parent.config(width=380)
        parent.pack_propagate(False)
        
        # Dosya seçimi kartı
        self.create_file_selection_card(parent)
        
        # Minimal boşluk
        tk.Frame(parent, bg=self.colors['background'], height=2).pack()
        
        # Seçenekler kartı
        self.create_options_card(parent)
        
        # Minimal boşluk
        tk.Frame(parent, bg=self.colors['background'], height=2).pack()
        
        # İşlem butonları
        self.create_action_buttons(parent)
        
    def create_dragdrop_file_input(self, parent, label_text, text_var, browse_command, icon):
        """Drag & Drop destekli dosya input grubu - Uniform font"""
        # Ana container
        container = tk.Frame(parent, bg=self.colors['card'])
        container.pack(fill=tk.X, pady=(0, 4))
        
        # Label - uniform font
        label = tk.Label(
            container,
            text=f"{icon} {label_text}",
            font=('Segoe UI', 8, 'bold'),  # Uniform font
            bg=self.colors['card'],
            fg=self.colors['text']
        )
        label.pack(anchor=tk.W, pady=(0, 1))
        
        # İpucu metni - uniform font
        hint_label = tk.Label(
            container,
            text="Sürükle veya Gözat",
            font=('Segoe UI', 8, 'italic'),  # Uniform font (7 → 8)
            bg=self.colors['card'],
            fg=self.colors['text_light']
        )
        hint_label.pack(anchor=tk.W, pady=(0, 1))
        
        # Input ve buton frame
        input_frame = tk.Frame(container, bg=self.colors['card'])
        input_frame.pack(fill=tk.X)
        
        # Entry - uniform font
        entry = ttk.Entry(
            input_frame,
            textvariable=text_var,
            font=('Segoe UI', 8),  # Uniform font
            style='DragDrop.TEntry'
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        
        # Gözat butonu
        browse_btn = ttk.Button(
            input_frame,
            text="Gözat",
            command=browse_command,
            style='Small.TButton'
        )
        browse_btn.pack(side=tk.RIGHT)
        
        # Drag & Drop fonksiyonalitesi ekle
        self.setup_drag_drop_for_entry(entry, text_var, browse_command)
        
        return entry
    
    def setup_drag_drop_for_entry(self, entry_widget, text_var, browse_command):
        """Entry widget'ına drag & drop fonksiyonalitesi ekle"""
        
        def on_drop(event):
            # Dosya yollarını al
            try:
                files = self.root.tk.splitlist(event.data)
                if files:
                    file_path = files[0]  # İlk dosyayı al
                    
                    # Dosyayı doğrula ve ayarla
                    if self.validate_dropped_file(file_path):
                        text_var.set(file_path)
                        
                        # Visual başarı feedback
                        self.show_entry_success(entry_widget)
                        
                        # Eğer file1 ise output filename'i güncelle
                        if text_var == self.app_logic.file1_path:
                            self.app_logic.update_output_filename(file_path)
                    else:
                        # Visual hata feedback
                        self.show_entry_error(entry_widget)
            except Exception as e:
                print(f"Drop işlemi hatası: {str(e)}")
                self.show_entry_error(entry_widget)
        
        def on_drag_enter(event):
            # Entry'ye hover efekti
            entry_widget.configure(style='DragDrop.TEntry')
            # Arka plan rengini değiştir (görsel feedback)
            self.root.after(10, lambda: self.set_entry_bg(entry_widget, self.colors['entry_hover']))
        
        def on_drag_leave(event):
            # Normal duruma dön
            self.root.after(10, lambda: self.set_entry_bg(entry_widget, self.colors['entry_normal']))
        
        # Event'leri bağla
        try:
            from tkinterdnd2 import DND_FILES
            entry_widget.drop_target_register(DND_FILES)
            entry_widget.dnd_bind('<<Drop>>', on_drop)
            entry_widget.dnd_bind('<<DragEnter>>', on_drag_enter)
            entry_widget.dnd_bind('<<DragLeave>>', on_drag_leave)
        except ImportError:
            # tkinterdnd2 yoksa basit tıklama ile dosya seçimi
            entry_widget.bind('<Double-Button-1>', lambda e: browse_command())
    
    def set_entry_bg(self, entry_widget, color):
        """Entry widget'ının arka plan rengini değiştir"""
        try:
            # ttk Entry için style kullanarak renk değiştirme
            style = ttk.Style()
            style.configure('DragDrop.TEntry', fieldbackground=color)
        except:
            pass
    
    def show_entry_success(self, entry_widget):
        """Entry için başarı visual feedback"""
        try:
            # Geçici olarak yeşil arka plan
            self.set_entry_bg(entry_widget, self.colors['success'])
            self.root.after(500, lambda: self.set_entry_bg(entry_widget, self.colors['entry_normal']))
        except:
            pass
    
    def show_entry_error(self, entry_widget):
        """Entry için hata visual feedback"""
        try:
            # Geçici olarak kırmızı arka plan
            self.set_entry_bg(entry_widget, self.colors['danger'])
            self.root.after(800, lambda: self.set_entry_bg(entry_widget, self.colors['entry_normal']))
        except:
            pass
    
    def validate_dropped_file(self, file_path):
        """Sürüklenen dosyayı doğrula"""
        try:
            # Dosya var mı?
            if not os.path.exists(file_path):
                self.show_error("Hata", "Dosya bulunamadı!")
                return False
                
            # Excel dosyası mı?
            valid_extensions = ['.xlsx', '.xls']
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext not in valid_extensions:
                self.show_error("Hata", f"Geçersiz dosya formatı!\nDesteklenen formatlar: {', '.join(valid_extensions)}")
                return False
                
            return True
            
        except Exception as e:
            self.show_error("Hata", f"Dosya kontrolü hatası: {str(e)}")
            return False
        
    def create_file_selection_card(self, parent):
        """Dosya seçimi kartı - Uniform font"""
        card_frame = tk.Frame(parent, bg=self.colors['card'], relief='solid', bd=1)
        card_frame.pack(fill=tk.X, pady=2)
        
        # Kart başlığı - uniform font
        header = tk.Label(
            card_frame,
            text="📁 Dosya Seçimi",
            font=('Segoe UI', 8, 'bold'),  # Uniform font
            bg=self.colors['card'],
            fg=self.colors['text'],
            pady=4
        )
        header.pack(anchor=tk.W, padx=12)
        
        # Dosya seçimi içeriği
        content_frame = tk.Frame(card_frame, bg=self.colors['card'])
        content_frame.pack(fill=tk.X, padx=8, pady=(0, 6))
        
        # Eski dosya - Drag & Drop destekli
        self.create_dragdrop_file_input(
            content_frame,
            "Eski Tarihli Excel",
            self.app_logic.file1_path,
            self.browse_file1,
            "📄"
        )
        
        # Yeni dosya - Drag & Drop destekli
        self.create_dragdrop_file_input(
            content_frame,
            "Yeni Tarihli Excel",
            self.app_logic.file2_path,
            self.browse_file2,
            "📄"
        )
        
        # Çıktı dosyası (sadece gösterim - gözat butonu yok)
        self.create_display_input(
            content_frame,
            "Sonuç Dosyası",
            self.app_logic.output_path,
            "💾"
        )
        
    def create_display_input(self, parent, label_text, text_var, icon):
        """Sadece gösterim amaçlı dosya input grubu - Uniform font"""
        # Ana container
        container = tk.Frame(parent, bg=self.colors['card'])
        container.pack(fill=tk.X, pady=(0, 4))
        
        # Label - uniform font
        label = tk.Label(
            container,
            text=f"{icon} {label_text}",
            font=('Segoe UI', 8, 'bold'),  # Uniform font
            bg=self.colors['card'],
            fg=self.colors['text']
        )
        label.pack(anchor=tk.W, pady=(0, 1))
        
        # İpucu metni - uniform font
        hint_label = tk.Label(
            container,
            text="Otomatik oluşturulur",
            font=('Segoe UI', 8, 'italic'),  # Uniform font (7 → 8)
            bg=self.colors['card'],
            fg=self.colors['text_light']
        )
        hint_label.pack(anchor=tk.W, pady=(0, 1))
        
        # Sadece Entry - uniform font
        entry = ttk.Entry(
            container,
            textvariable=text_var,
            font=('Segoe UI', 8),  # Uniform font
            style='Modern.TEntry',
            state='readonly'
        )
        entry.pack(fill=tk.X)
        
        return entry
        
    def create_options_card(self, parent):
        """Seçenekler kartı - Uniform font"""
        card_frame = tk.Frame(parent, bg=self.colors['card'], relief='solid', bd=1)
        card_frame.pack(fill=tk.X, pady=2)
        
        # Kart başlığı - uniform font
        header = tk.Label(
            card_frame,
            text="⚙️ Seçenekler",
            font=('Segoe UI', 8, 'bold'),  # Uniform font
            bg=self.colors['card'],
            fg=self.colors['text'],
            pady=4
        )
        header.pack(anchor=tk.W, padx=12)
        
        # Seçenekler içeriği
        content_frame = tk.Frame(card_frame, bg=self.colors['card'])
        content_frame.pack(fill=tk.X, padx=8, pady=(0, 6))
        
        # Büyük/küçük harf seçeneği - uniform font
        case_check = ttk.Checkbutton(
            content_frame,
            text="Büyük/Küçük Harf Duyarlı",
            variable=self.app_logic.case_sensitive,
            style='Small.TCheckbutton'
        )
        case_check.pack(anchor=tk.W, pady=2)
        
        # Kaydetme formatı - uniform font
        tk.Label(
            content_frame,
            text="💾 Kaydetme Formatı:",
            font=('Segoe UI', 8, 'bold'),  # Uniform font
            bg=self.colors['card'],
            fg=self.colors['text']
        ).pack(anchor=tk.W, pady=(6, 2))
        
        # Format seçenekleri - uniform font
        format_frame = tk.Frame(content_frame, bg=self.colors['card'])
        format_frame.pack(anchor=tk.W, padx=6)
        
        # Excel checkbox
        self.save_excel = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            format_frame,
            text="Excel (.xlsx)",
            variable=self.save_excel,
            style='Small.TCheckbutton'
        ).pack(anchor=tk.W, pady=1)
        
        # Resim checkbox
        self.save_image = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            format_frame,
            text="Resim (.png)",
            variable=self.save_image,
            style='Small.TCheckbutton'
        ).pack(anchor=tk.W, pady=1)
        
    def create_action_buttons(self, parent):
        """İşlem butonları - Uniform font"""
        button_frame = tk.Frame(parent, bg=self.colors['background'])
        button_frame.pack(fill=tk.X, pady=6)
        
        # Progress bar (başlangıçta gizli)
        self.progress = ttk.Progressbar(
            button_frame,
            mode='indeterminate',
            length=350
        )
        
        # Karşılaştır butonu (Ana buton) - uniform font
        self.compare_btn = ttk.Button(
            button_frame,
            text="🔍 Karşılaştır",
            command=self.safe_compare_files,
            style='Accent.TButton'
        )
        self.compare_btn.pack(fill=tk.X, pady=(0, 6))
        
        # Araç-Plasiyer Ayarları butonu - uniform font
        settings_btn = ttk.Button(
            button_frame,
            text="⚙️ Araç-Plasiyer Ayarları",
            command=self.edit_vehicle_settings,
            style='Small.TButton'
        )
        settings_btn.pack(fill=tk.X, pady=(0, 6))
        
        # Temizle butonu - uniform font
        clear_btn = ttk.Button(
            button_frame,
            text="🗑️ Temizle",
            command=self.app_logic.clear_results,
            style='Small.TButton'
        )
        clear_btn.pack(fill=tk.X)
    
    def edit_vehicle_settings(self):
        """Araç-plasiyer ayarları düzenleme"""
        self.app_logic.edit_vehicle_drivers()
        
    def safe_compare_files(self):
        """Güvenli dosya karşılaştırma - Progress bar ile"""
        try:
            # Buton durumunu değiştir
            self.compare_btn.configure(text="⏳ İşleniyor...", state='disabled')
            
            # Progress bar'ı göster
            self.progress.pack(fill=tk.X, pady=5)
            self.progress.start(10)
            
            # Ana thread'de UI güncellemelerini yap
            self.root.update()
            
            # Karşılaştırmayı başlat
            self.app_logic.compare_files()
            
            # 2 saniye sonra progress bar'ı gizle (thread tamamlandıktan sonra)
            self.root.after(2000, self.reset_ui)
            
        except Exception as e:
            self.show_error("Hata", f"Karşılaştırma başlatılamadı: {str(e)}")
            self.reset_ui()
    
    def reset_ui(self):
        """UI'ı sıfırla"""
        try:
            # Progress bar'ı durdur ve gizle
            self.progress.stop()
            self.progress.pack_forget()
            
            # Buton durumunu eski haline getir
            self.compare_btn.configure(text="🔍 Karşılaştır", state='normal')
        except:
            pass
        
    def create_right_panel(self, parent):
        """Sağ panel (Sonuçlar) - Uniform font"""
        # Sonuçlar kartı
        card_frame = tk.Frame(parent, bg=self.colors['card'], relief='solid', bd=1)
        card_frame.pack(fill=tk.BOTH, expand=True, pady=2)
        
        # Kart başlığı - uniform font
        header_frame = tk.Frame(card_frame, bg=self.colors['card'])
        header_frame.pack(fill=tk.X, padx=12, pady=6)
        
        tk.Label(
            header_frame,
            text="📊 Sonuçlar",
            font=('Segoe UI', 8, 'bold'),  # Uniform font (10 → 8)
            bg=self.colors['card'],
            fg=self.colors['text']
        ).pack(side=tk.LEFT)
        
        # Sonuç tablosu frame
        table_frame = tk.Frame(card_frame, bg=self.colors['card'])
        table_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        
        # Treeview ve scrollbar
        tree_frame = tk.Frame(table_frame, bg=self.colors['card'])
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ("no", "unvan")
        self.result_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            yscrollcommand=scrollbar.set,
            height=12
        )
        
        # Başlıkları ayarla
        self.result_tree.heading("no", text="#")
        self.result_tree.heading("unvan", text="Cari Ünvan")
        
        # Sütun genişlikleri
        self.result_tree.column("no", width=50, anchor=tk.CENTER)
        self.result_tree.column("unvan", width=350)
        
        self.result_tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.result_tree.yview)
        
        # Durum bilgisi - uniform font
        status_frame = tk.Frame(card_frame, bg=self.colors['card'])
        status_frame.pack(fill=tk.X, padx=12, pady=(0, 8))
        
        self.status_var = tk.StringVar(value="Henüz karşılaştırma yapılmadı.")
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=('Segoe UI', 8, 'italic'),  # Uniform font
            bg=self.colors['card'],
            fg=self.colors['text_light'],
            wraplength=350
        )
        status_label.pack(anchor=tk.W)
    
    def validate_file_selection(self, file_path, file_type):
        """Dosya seçimini doğrula"""
        if not file_path:
            return False, f"Lütfen {file_type} dosyasını seçin!"
            
        if not os.path.exists(file_path):
            return False, f"{file_type} dosyası bulunamadı!"
            
    def validate_file_selection(self, file_path, file_type):
        """Dosya seçimini doğrula"""
        if not file_path:
            return False, f"Lütfen {file_type} dosyasını seçin!"
            
        if not os.path.exists(file_path):
            return False, f"{file_type} dosyası bulunamadı!"
            
        # Dosya uzantısını kontrol et
        valid_extensions = ['.xlsx', '.xls']
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext not in valid_extensions:
            return False, f"Geçersiz dosya formatı! Desteklenen formatlar: {', '.join(valid_extensions)}"
            
        return True, ""
        
    def browse_file1(self):
        """Eski Excel dosyasını seç"""
        file_path = filedialog.askopenfilename(
            title="Eski Tarihli Excel Dosyasını Seç",
            filetypes=[
                ("Excel Dosyaları", "*.xlsx *.xls"), 
                ("Excel 2007-2019", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Tüm Dosyalar", "*.*")
            ],
            initialdir=os.path.expanduser("~")
        )
        
        if file_path:
            # Dosyayı doğrula
            is_valid, error_msg = self.validate_file_selection(file_path, "Eski tarihli Excel")
            
            if is_valid:
                self.app_logic.file1_path.set(file_path)
                self.app_logic.update_output_filename(file_path)
            else:
                self.show_error("Dosya Seçim Hatası", error_msg)
            
    def browse_file2(self):
        """Yeni Excel dosyasını seç"""
        file_path = filedialog.askopenfilename(
            title="Yeni Tarihli Excel Dosyasını Seç",
            filetypes=[
                ("Excel Dosyaları", "*.xlsx *.xls"), 
                ("Excel 2007-2019", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Tüm Dosyalar", "*.*")
            ],
            initialdir=os.path.expanduser("~")
        )
        
        if file_path:
            # Dosyayı doğrula
            is_valid, error_msg = self.validate_file_selection(file_path, "Yeni tarihli Excel")
            
            if is_valid:
                self.app_logic.file2_path.set(file_path)
            else:
                self.show_error("Dosya Seçim Hatası", error_msg)
            
    def update_results(self, results, status_text):
        """Sonuçları güncelle - Thread-safe"""
        def _update():
            try:
                # Mevcut sonuçları temizle
                for item in self.result_tree.get_children():
                    self.result_tree.delete(item)
                    
                # Yeni sonuçları ekle
                for i, unvan in enumerate(results, 1):
                    # Çok uzun ünvanları kısalt
                    display_unvan = unvan if len(str(unvan)) <= 50 else str(unvan)[:47] + "..."
                    self.result_tree.insert("", tk.END, values=(i, display_unvan))
                    
                # Durum metnini güncelle
                self.status_var.set(status_text)
                
                # UI'ı sıfırla
                self.reset_ui()
                
            except Exception as e:
                print(f"UI güncelleme hatası: {str(e)}")
        
        # Ana thread'de çalıştır
        self.root.after(0, _update)
        
    def clear_results(self):
        """Sonuçları temizle"""
        try:
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            self.status_var.set("Henüz karşılaştırma yapılmadı.")
        except Exception as e:
            print(f"Sonuç temizleme hatası: {str(e)}")
        
    def show_info(self, title, message):
        """Bilgi mesajı göster - Thread-safe"""
        def _show():
            messagebox.showinfo(title, message)
        self.root.after(0, _show)
        
    def show_error(self, title, message):
        """Hata mesajı göster - Thread-safe"""
        def _show():
            messagebox.showerror(title, message)
        self.root.after(0, _show)
        
    def show_warning(self, title, message):
        """Uyarı mesajı göster - Thread-safe"""
        def _show():
            messagebox.showwarning(title, message)
        self.root.after(0, _show)