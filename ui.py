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
        """Modern stil konfigürasyonları"""
        # Ana buton stili
        style.configure(
            'Modern.TButton',
            font=('Segoe UI', 10),
            padding=(20, 12),
            relief='flat',
            borderwidth=0
        )
        
        # Vurgulu buton stili
        style.configure(
            'Accent.TButton',
            font=('Segoe UI', 10, 'bold'),
            padding=(25, 15),
            relief='flat',
            borderwidth=0
        )
        
        # İkincil buton stili
        style.configure(
            'Secondary.TButton',
            font=('Segoe UI', 10),
            padding=(20, 12),
            relief='flat',
            borderwidth=1
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
            font=('Segoe UI', 11, 'bold'),
            padding=(0, 5)
        )
        
        # Entry stili
        style.configure(
            'Modern.TEntry',
            font=('Segoe UI', 10),
            padding=10,
            relief='flat',
            borderwidth=1
        )
        
        # Drag & Drop Entry stili
        style.configure(
            'DragDrop.TEntry',
            font=('Segoe UI', 10),
            padding=12,
            relief='solid',
            borderwidth=2
        )
        
        # Checkbutton stili
        style.configure(
            'Modern.TCheckbutton',
            font=('Segoe UI', 10),
            padding=(0, 5)
        )
        
        # Radiobutton stili
        style.configure(
            'Modern.TRadiobutton',
            font=('Segoe UI', 10),
            padding=(0, 3)
        )
        
    def setup_main_window(self):
        """Ana pencere ayarları"""
        self.root.title("CAL Excel Cari Karşılaştırma")
        self.root.geometry("1000x850")
        self.root.minsize(950, 800)
        
        # Pencereyi ortala
        self.center_window()
        
        # Arka plan rengi
        self.root.configure(bg=self.colors['background'])
        
        # Pencere kapatma protokolü
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # İkon ayarla (varsa)
        try:
            # İkon dosyası varsa kullan
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
        """Modern arayüz bileşenlerini oluştur"""
        # Ana konteyner - padding'i azalt
        main_container = tk.Frame(self.root, bg=self.colors['background'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Başlık bölümü
        self.create_header(main_container)
        
        # İçerik container'ı - üst padding'i azalt
        content_frame = tk.Frame(main_container, bg=self.colors['background'])
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Sol panel (Dosya seçimi ve seçenekler)
        left_panel = tk.Frame(content_frame, bg=self.colors['background'])
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        
        # Sağ panel (Sonuçlar)
        right_panel = tk.Frame(content_frame, bg=self.colors['background'])
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Sol paneli oluştur
        self.create_left_panel(left_panel)
        
        # Sağ paneli oluştur
        self.create_right_panel(right_panel)
        
    def create_header(self, parent):
        """Başlık bölümünü oluştur"""
        header_frame = tk.Frame(parent, bg=self.colors['card'], height=80)
        header_frame.pack(fill=tk.X, pady=(0, 8))
        header_frame.pack_propagate(False)
        
        # Başlık metni
        title_label = tk.Label(
            header_frame,
            text="Excel Cari Ünvan Karşılaştırma",
            font=('Segoe UI', 16, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text'],
            pady=8
        )
        title_label.pack()
        
        # Alt başlık
        subtitle_label = tk.Label(
            header_frame,
            text="İki Excel dosyasındaki cari ünvanları karşılaştırır ve farklılıkları tespit eder.",
            font=('Segoe UI', 9),
            bg=self.colors['card'],
            fg=self.colors['text_light']
        )
        subtitle_label.pack()
        
    def create_left_panel(self, parent):
        """Sol panel (Dosya seçimi ve seçenekler)"""
        # Sol panel için sabit genişlik ayarla
        parent.config(width=400)
        parent.pack_propagate(False)
        
        # Dosya seçimi kartı
        self.create_file_selection_card(parent)
        
        # Boşluk - Minimal
        tk.Frame(parent, bg=self.colors['background'], height=3).pack()
        
        # Seçenekler kartı
        self.create_options_card(parent)
        
        # Boşluk - Minimal
        tk.Frame(parent, bg=self.colors['background'], height=3).pack()
        
        # İşlem butonları
        self.create_action_buttons(parent)
        
    def create_dragdrop_file_input(self, parent, label_text, text_var, browse_command, icon):
        """Drag & Drop destekli dosya input grubu"""
        # Ana container
        container = tk.Frame(parent, bg=self.colors['card'])
        container.pack(fill=tk.X, pady=(0, 8))
        
        # Label
        label = tk.Label(
            container,
            text=f"{icon} {label_text}",
            font=('Segoe UI', 9, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text']
        )
        label.pack(anchor=tk.W, pady=(0, 3))
        
        # İpucu metni
        hint_label = tk.Label(
            container,
            text="Dosyayı aşağıdaki alana sürükleyip bırakın veya Gözat butonuna tıklayın",
            font=('Segoe UI', 8, 'italic'),
            bg=self.colors['card'],
            fg=self.colors['text_light']
        )
        hint_label.pack(anchor=tk.W, pady=(0, 3))
        
        # Input ve buton frame
        input_frame = tk.Frame(container, bg=self.colors['card'])
        input_frame.pack(fill=tk.X)
        
        # Entry - Drag & Drop destekli
        entry = ttk.Entry(
            input_frame,
            textvariable=text_var,
            font=('Segoe UI', 10),
            style='DragDrop.TEntry'
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        
        # Gözat butonu
        browse_btn = ttk.Button(
            input_frame,
            text="Gözat",
            command=browse_command,
            style='Secondary.TButton'
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
        """Dosya seçimi kartı - Drag & Drop destekli"""
        card_frame = tk.Frame(parent, bg=self.colors['card'], relief='solid', bd=1)
        card_frame.pack(fill=tk.X, pady=3)
        
        # Kart başlığı
        header = tk.Label(
            card_frame,
            text="📁 Dosya Seçimi",
            font=('Segoe UI', 10, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text'],
            pady=6
        )
        header.pack(anchor=tk.W, padx=15)
        
        # Dosya seçimi içeriği
        content_frame = tk.Frame(card_frame, bg=self.colors['card'])
        content_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        
        # Eski dosya - Drag & Drop destekli
        self.create_dragdrop_file_input(
            content_frame,
            "Eski Tarihli Excel Dosyası",
            self.app_logic.file1_path,
            self.browse_file1,
            "📄"
        )
        
        # Yeni dosya - Drag & Drop destekli
        self.create_dragdrop_file_input(
            content_frame,
            "Yeni Tarihli Excel Dosyası",
            self.app_logic.file2_path,
            self.browse_file2,
            "📄"
        )
        
        # Çıktı dosyası (normal)
        self.create_file_input(
            content_frame,
            "Sonuç Dosyası",
            self.app_logic.output_path,
            self.browse_output,
            "💾"
        )
        
    def create_file_input(self, parent, label_text, text_var, browse_command, icon):
        """Normal dosya seçimi input grubu"""
        # Ana container
        container = tk.Frame(parent, bg=self.colors['card'])
        container.pack(fill=tk.X, pady=(0, 8))
        
        # Label
        label = tk.Label(
            container,
            text=f"{icon} {label_text}",
            font=('Segoe UI', 9, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text']
        )
        label.pack(anchor=tk.W, pady=(0, 3))
        
        # Input ve buton frame
        input_frame = tk.Frame(container, bg=self.colors['card'])
        input_frame.pack(fill=tk.X)
        
        # Entry
        entry = ttk.Entry(
            input_frame,
            textvariable=text_var,
            font=('Segoe UI', 10),
            style='Modern.TEntry'
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        
        # Gözat butonu
        browse_btn = ttk.Button(
            input_frame,
            text="Gözat",
            command=browse_command,
            style='Secondary.TButton'
        )
        browse_btn.pack(side=tk.RIGHT)
        
    def create_options_card(self, parent):
        """Seçenekler kartı"""
        card_frame = tk.Frame(parent, bg=self.colors['card'], relief='solid', bd=1)
        card_frame.pack(fill=tk.X, pady=3)
        
        # Kart başlığı
        header = tk.Label(
            card_frame,
            text="⚙️ Seçenekler",
            font=('Segoe UI', 10, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text'],
            pady=6
        )
        header.pack(anchor=tk.W, padx=15)
        
        # Seçenekler içeriği
        content_frame = tk.Frame(card_frame, bg=self.colors['card'])
        content_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        
        # Büyük/küçük harf seçeneği
        case_check = ttk.Checkbutton(
            content_frame,
            text="Büyük/Küçük Harf Duyarlı Karşılaştırma",
            variable=self.app_logic.case_sensitive,
            style='Modern.TCheckbutton'
        )
        case_check.pack(anchor=tk.W, pady=3)
        
        # Kaydetme formatı
        tk.Label(
            content_frame,
            text="💾 Kaydetme Formatı:",
            font=('Segoe UI', 9, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text']
        ).pack(anchor=tk.W, pady=(10, 3))
        
        # Format seçenekleri - Checkbox'lar
        format_frame = tk.Frame(content_frame, bg=self.colors['card'])
        format_frame.pack(anchor=tk.W, padx=8)
        
        # Excel checkbox
        self.save_excel = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            format_frame,
            text="Excel (.xlsx)",
            variable=self.save_excel,
            style='Modern.TCheckbutton'
        ).pack(anchor=tk.W, pady=1)
        
        # Resim checkbox
        self.save_image = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            format_frame,
            text="Resim (.png)",
            variable=self.save_image,
            style='Modern.TCheckbutton'
        ).pack(anchor=tk.W, pady=1)
        
    def create_action_buttons(self, parent):
        """İşlem butonları"""
        button_frame = tk.Frame(parent, bg=self.colors['background'])
        button_frame.pack(fill=tk.X, pady=8)
        
        # Progress bar (başlangıçta gizli)
        self.progress = ttk.Progressbar(
            button_frame,
            mode='indeterminate',
            length=350
        )
        
        # Karşılaştır butonu (Ana buton)
        self.compare_btn = ttk.Button(
            button_frame,
            text="🔍 Karşılaştır",
            command=self.safe_compare_files,
            style='Accent.TButton'
        )
        self.compare_btn.pack(fill=tk.X, pady=(0, 8))
        
        # Temizle butonu
        clear_btn = ttk.Button(
            button_frame,
            text="🗑️ Temizle",
            command=self.app_logic.clear_results,
            style='Secondary.TButton'
        )
        clear_btn.pack(fill=tk.X)
        
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
        """Sağ panel (Sonuçlar)"""
        # Sonuçlar kartı
        card_frame = tk.Frame(parent, bg=self.colors['card'], relief='solid', bd=1)
        card_frame.pack(fill=tk.BOTH, expand=True, pady=3)
        
        # Kart başlığı
        header_frame = tk.Frame(card_frame, bg=self.colors['card'])
        header_frame.pack(fill=tk.X, padx=15, pady=10)
        
        tk.Label(
            header_frame,
            text="📊 Sonuçlar",
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['card'],
            fg=self.colors['text']
        ).pack(side=tk.LEFT)
        
        # Sonuç tablosu frame
        table_frame = tk.Frame(card_frame, bg=self.colors['card'])
        table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
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
            height=15
        )
        
        # Başlıkları ayarla
        self.result_tree.heading("no", text="#")
        self.result_tree.heading("unvan", text="Cari Ünvan")
        
        # Sütun genişlikleri
        self.result_tree.column("no", width=50, anchor=tk.CENTER)
        self.result_tree.column("unvan", width=400)
        
        self.result_tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.result_tree.yview)
        
        # Durum bilgisi
        status_frame = tk.Frame(card_frame, bg=self.colors['card'])
        status_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        self.status_var = tk.StringVar(value="Henüz karşılaştırma yapılmadı.")
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=('Segoe UI', 9, 'italic'),
            bg=self.colors['card'],
            fg=self.colors['text_light'],
            wraplength=400
        )
        status_label.pack(anchor=tk.W)
    
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
            
    def browse_output(self):
        """Sonuç dosyasını kaydet"""
        # Hangi formatların seçili olduğunu kontrol et
        excel_selected = self.save_excel.get()
        image_selected = self.save_image.get()
        
        if not excel_selected and not image_selected:
            self.show_warning("Uyarı", "Lütfen en az bir kaydetme formatı seçin!")
            return
        
        if excel_selected and image_selected:
            # Her iki format da seçili
            filetypes = [("Tüm Dosyalar", "*.*")]
            defaultextension = ""
            title = "Sonuç Dosyalarını Kaydet (uzantı olmadan)"
        elif excel_selected:
            # Sadece Excel
            filetypes = [("Excel Dosyaları", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
            defaultextension = ".xlsx"
            title = "Excel Sonuç Dosyasını Kaydet"
        elif image_selected:
            # Sadece Resim
            filetypes = [("PNG Dosyaları", "*.png"), ("Tüm Dosyalar", "*.*")]
            defaultextension = ".png"
            title = "Resim Sonuç Dosyasını Kaydet"
        else:
            # Hiçbiri seçili değil (bu duruma normalde gelmemeli)
            filetypes = [("Tüm Dosyalar", "*.*")]
            defaultextension = ""
            title = "Sonuç Dosyasını Kaydet"
            
        file_path = filedialog.asksaveasfilename(
            title=title,
            defaultextension=defaultextension,
            filetypes=filetypes,
            initialdir=os.path.expanduser("~")
        )
        
        if file_path:
            # Uzantıyı kaldır (çünkü program kendi uzantılarını ekleyecek)
            base_name = os.path.splitext(file_path)[0] 
            self.app_logic.output_path.set(base_name)
            
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
                    display_unvan = unvan if len(str(unvan)) <= 60 else str(unvan)[:57] + "..."
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