import tkinter as tk
from tkinter import filedialog, ttk
import os
from amazon_scraper import AmazonScraper
import threading
import asyncio
import nest_asyncio
from tkinter import font
from PIL import Image, ImageTk
import webbrowser

# Apply nest_asyncio to handle nested event loops
nest_asyncio.apply()

class ModernButton(tk.Button):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.config(
            background='#2962ff',
            foreground='white',
            font=('Helvetica', 10, 'bold'),
            borderwidth=0,
            padx=20,
            pady=10,
            cursor='hand2'
        )
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)

    def on_enter(self, e):
        self['background'] = '#1565c0'

    def on_leave(self, e):
        self['background'] = '#2962ff'

class AmazonScraperGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Enterprise Amazon Product Scraper")
        self.window.geometry("800x600")
        self.window.configure(bg='#1a1a1a')
        
        # Set minimum window size
        self.window.minsize(800, 600)
        
        # Custom fonts
        self.title_font = font.Font(family='Helvetica', size=24, weight='bold')
        self.subtitle_font = font.Font(family='Helvetica', size=12)
        self.text_font = font.Font(family='Helvetica', size=10)
        
        # Create gradient frame
        self.create_gradient_frame()
        
        # Create main content
        self.create_main_content()
        
        # Create footer
        self.create_footer()
        
        # Create directories
        os.makedirs("temp", exist_ok=True)
        os.makedirs("output", exist_ok=True)

    def create_gradient_frame(self):
        # Main container with gradient effect
        self.gradient_frame = tk.Frame(self.window, bg='#1a1a1a')
        self.gradient_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header frame
        self.header_frame = tk.Frame(self.gradient_frame, bg='#1a1a1a')
        self.header_frame.pack(fill=tk.X, padx=20, pady=(30, 0))
        
        # Title with modern styling
        title_label = tk.Label(
            self.header_frame,
            text="Enterprise Amazon Product Scraper",
            font=self.title_font,
            bg='#1a1a1a',
            fg='#ffffff'
        )
        title_label.pack(pady=(0, 10))
        
        # Subtitle
        subtitle_label = tk.Label(
            self.header_frame,
            text="Professional Data Extraction Solution",
            font=self.subtitle_font,
            bg='#1a1a1a',
            fg='#bbbbbb'
        )
        subtitle_label.pack()

    def create_main_content(self):
        # Main content frame
        self.main_frame = tk.Frame(self.gradient_frame, bg='#1a1a1a')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)
        
        # File selection frame
        file_frame = tk.Frame(self.main_frame, bg='#1a1a1a')
        file_frame.pack(fill=tk.X, pady=20)
        
        # File selection button with modern styling
        self.file_button = ModernButton(
            file_frame,
            text="Select Excel File",
            command=self.select_file
        )
        self.file_button.pack(pady=10)
        
        # File label with modern styling
        self.file_label = tk.Label(
            file_frame,
            text="No file selected",
            font=self.text_font,
            bg='#1a1a1a',
            fg='#bbbbbb'
        )
        self.file_label.pack(pady=5)
        
        # Start button with modern styling
        self.start_button = ModernButton(
            file_frame,
            text="Start Scraping",
            command=self.start_scraping
        )
        self.start_button.pack(pady=20)
        
        # Progress frame
        progress_frame = tk.Frame(self.main_frame, bg='#1a1a1a')
        progress_frame.pack(fill=tk.X, pady=10)
        
        # Progress label with modern styling
        self.progress_label = tk.Label(
            progress_frame,
            text="",
            font=self.text_font,
            bg='#1a1a1a',
            fg='#bbbbbb'
        )
        self.progress_label.pack(pady=5)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='indeterminate',
            length=300
        )
        self.progress_bar.pack(pady=10)
        
        # Results label with modern styling
        self.result_label = tk.Label(
            progress_frame,
            text="",
            font=self.text_font,
            bg='#1a1a1a',
            fg='#bbbbbb',
            wraplength=500
        )
        self.result_label.pack(pady=5)

    def create_footer(self):
        # Footer frame
        footer_frame = tk.Frame(self.gradient_frame, bg='#1a1a1a')
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=20)
        
        # Copyright text
        copyright_label = tk.Label(
            footer_frame,
            text="© 2024 Enterprise Amazon Scraper. All rights reserved.",
            font=self.text_font,
            bg='#1a1a1a',
            fg='#666666'
        )
        copyright_label.pack(side=tk.BOTTOM)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.file_label.config(text=os.path.basename(filename))
            self.selected_file = filename
            self.file_label.config(fg='#4CAF50')  # Green color for success

    def run_scraper(self):
        try:
            # Initialize scraper
            scraper = AmazonScraper(output_dir="output")
            
            # Create and set event loop
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            # Run the scraper
            output_file = loop.run_until_complete(scraper.process_excel(self.selected_file))
            
            # Update GUI with results
            self.window.after(0, self.show_success, output_file)
            
        except Exception as e:
            self.window.after(0, self.show_error, str(e))
        finally:
            loop.close()

    def start_scraping(self):
        if not hasattr(self, 'selected_file'):
            self.result_label.config(text="Please select a file first!", fg='#ff3d00')
            return
            
        self.progress_label.config(text="Processing... Please wait...")
        self.progress_bar.start()
        self.start_button.config(state='disabled')
        self.file_button.config(state='disabled')
        
        # Run scraper in separate thread
        thread = threading.Thread(target=self.run_scraper)
        thread.daemon = True
        thread.start()

    def show_success(self, output_file):
        self.progress_bar.stop()
        self.progress_label.config(text="")
        self.start_button.config(state='normal')
        self.file_button.config(state='normal')
        
        if output_file and os.path.exists(output_file):
            self.result_label.config(
                text=f"✅ Scraping completed successfully!\nResults saved to: {output_file}",
                fg='#4CAF50'  # Green color
            )
            # Open output directory
            os.startfile(os.path.dirname(output_file))
        else:
            self.result_label.config(
                text="❌ Error: Output file not generated",
                fg='#ff3d00'  # Red color
            )

    def show_error(self, error_message):
        self.progress_bar.stop()
        self.progress_label.config(text="")
        self.start_button.config(state='normal')
        self.file_button.config(state='normal')
        self.result_label.config(
            text=f"❌ Error: {error_message}",
            fg='#ff3d00'  # Red color
        )

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = AmazonScraperGUI()
    app.run() 