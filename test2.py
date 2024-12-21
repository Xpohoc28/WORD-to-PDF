import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import threading
from concurrent.futures import ThreadPoolExecutor
import queue
from docx2pdf import convert
import win32com.client
import pythoncom
from utils import *

class FileConverter:
    """Класс для конвертации файлов"""
    @staticmethod
    def docx_to_pdf(input_file, output_file):
        convert(input_file, output_file)
        
    @staticmethod
    def pdf_to_docx(input_file, output_file):
        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            doc = word.Documents.Open(input_file)
            doc.SaveAs(output_file, FileFormat=16)
            doc.Close()
            word.Quit()
        finally:
            pythoncom.CoUninitialize()

class ConverterGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.setup_threading()
        self.create_widgets()
        
    def setup_window(self):
        """Настройка основного окна"""
        self.root.title("Конвертер DOCX ⟷ PDF")
        self.root.geometry("800x600")
        try:
            self.root.iconbitmap(resource_path('icon.ico'))
        except:
            pass
            
    def setup_variables(self):
        """Инициализация переменных"""
        self.file_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.output_directory = os.path.expanduser("~")
        self.conversion_mode = tk.StringVar(value=DOCX_TO_PDF)
        self.total_files = 0
        self.processed_files = 0
        self.is_converting = False
        
    def setup_threading(self):
        """Настройка многопоточности"""
        max_workers = max(4, os.cpu_count() or 4)
        self.thread_pool = ThreadPoolExecutor(max_workers=max_workers)
        self.status_lock = threading.Lock()
        
    def create_widgets(self):
        """Создание элементов интерфейса"""
        self._create_mode_selector()
        self._create_buttons()
        self._create_progress_bar()
        self._create_file_list()
        
    def _create_mode_selector(self):
        """Создание селектора режима конвертации"""
        mode_frame = ttk.LabelFrame(self.root, text="Режим конвертации")
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        
        for mode, text in [(DOCX_TO_PDF, "DOCX → PDF"), (PDF_TO_DOCX, "PDF → DOCX")]:
            ttk.Radiobutton(
                mode_frame,
                text=text,
                variable=self.conversion_mode,
                value=mode
            ).pack(side=tk.LEFT, padx=5)
            
    # ... (остальные методы создания виджетов)

    def convert_single_file(self, input_file):
        """Конвертация одного файла"""
        try:
            file_name = os.path.basename(input_file)
            mode = self.conversion_mode.get()
            output_name = os.path.splitext(file_name)[0]
            output_name += ".pdf" if mode == DOCX_TO_PDF else ".docx"
            output_file = os.path.join(self.output_directory, output_name)
            
            if os.path.exists(output_file):
                if not self._handle_existing_file(input_file, output_file):
                    return False
                    
            converter = FileConverter.docx_to_pdf if mode == DOCX_TO_PDF else FileConverter.pdf_to_docx
            converter(input_file, output_file)
            
            self.result_queue.put((input_file, STATUS_DONE))
            return True
            
        except Exception as e:
            self.result_queue.put((input_file, STATUS_ERROR))
            self.root.after(0, lambda: self._show_error(input_file, str(e)))
            return False
            
    def _show_error(self, file, error_msg):
        """Отображение ошибки"""
        messagebox.showerror(
            "Ошибка конвертации", 
            f"Ошибка при конвертации {file}:\n{error_msg}"
        )

    # ... (остальные вспомогательные методы)

def main():
    root = tk.Tk()
    app = ConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 