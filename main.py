import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from threading import Thread
import queue
import sys
from docx2pdf import convert

class DocxToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер DOCX в PDF")
        self.root.geometry("800x600")
        
        # Устанавливаем иконку окна
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # Очередь для обработки файлов
        self.file_queue = queue.Queue()
        
        # Директория для сохранения по умолчанию
        self.output_directory = os.path.expanduser("~")
        
        # Создаем основной интерфейс
        self.create_widgets()
        
        self.is_converting = False
        
    def create_widgets(self):
        # Фрейм для кнопок выбора
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Кнопка выбора файлов
        self.select_btn = ttk.Button(
            self.button_frame, 
            text="Выбрать файлы Word", 
            command=self.select_files
        )
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        # Кнопка выбора директории
        self.dir_btn = ttk.Button(
            self.button_frame,
            text="Выбрать папку сохранения",
            command=self.select_output_directory
        )
        self.dir_btn.pack(side=tk.LEFT, padx=5)
        
        # Метка с текущей директорией
        self.dir_label = ttk.Label(
            self.root,
            text=f"Папка сохранения: {self.output_directory}"
        )
        self.dir_label.pack(pady=5)
        
        # Список файлов
        self.file_frame = ttk.Frame(self.root)
        self.file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Создаем Treeview для отображения файлов
        self.file_tree = ttk.Treeview(
            self.file_frame,
            columns=("Файл", "Статус"),
            show="headings"
        )
        self.file_tree.heading("Файл", text="Файл")
        self.file_tree.heading("Статус", text="Статус")
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Добавляем скроллбар
        scrollbar = ttk.Scrollbar(self.file_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        # Кнопка конвертации
        self.convert_btn = ttk.Button(
            self.root,
            text="Конвертировать",
            command=self.start_conversion
        )
        self.convert_btn.pack(pady=10)
        
        # Кнопка отмены (добавить после кнопки конвертации)
        self.cancel_btn = ttk.Button(
            self.root,
            text="Отменить",
            command=self.stop_conversion,
            state="disabled"  # Изначально кнопка неактивна
        )
        self.cancel_btn.pack(pady=5)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(
            self.root,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        self.progress.pack(pady=10)
    
        

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите файлы Word",
            filetypes=[("Word files", "*.docx")]
        )
        
        # Нужно добавить проверку
        if not files:
            return
        
        # Очищаем текущий список
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
            
        # Добавляем новые файлы
        for file in files:
            self.file_tree.insert("", tk.END, values=(file, "Ожидает"))
            self.file_queue.put(file)
            
    def select_output_directory(self):
        directory = filedialog.askdirectory(
            title="Выерите папку для сохранения PDF файлов",
            initialdir=self.output_directory
        )
        if not directory:
            return
        if directory:
            self.output_directory = directory
            self.dir_label.config(text=f"Папка сохранения: {self.output_directory}")

    def convert_file(self, input_file):
        try:
            file_name = os.path.basename(input_file)
            output_name = os.path.splitext(file_name)[0] + ".pdf"
            output_file = os.path.join(self.output_directory, output_name)
            
            if os.path.exists(output_file):
                if not messagebox.askyesno("Файл существует", 
                    "Файл уже существует. Перезаписать?"):
                    return False
            
            convert(input_file, output_file)
            return True
        except Exception as e:
            # Показываем ошибку пользователю
            messagebox.showerror("Ошибка конвертации", 
                f"Ошибка при конвертации {input_file}:\n{str(e)}")
            print(f"Ошибка при конвертации {input_file}: {str(e)}")  # Для отладки
            return False
            
    def start_conversion(self):
        self.is_converting = True
        
        if self.file_queue.empty():
            messagebox.showwarning("Прдупреждение", "Выберите файлы для конвертации!")
            return
            
        self.convert_btn.configure(state="disabled")
        self.select_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")  # Активируем кнопку отмены
        
        # Запускаем конвертацию в отдельном потоке
        conversion_thread = Thread(target=self.process_queue)
        conversion_thread.daemon = True
        conversion_thread.start()
        
    def process_queue(self):
        try:
            total_files = self.file_queue.qsize()
            processed_files = 0
            
            while not self.file_queue.empty() and self.is_converting:  # Добавляем проверку флага
                current_file = self.file_queue.get()
                
                # Находим элемент в дереве
                for item in self.file_tree.get_children():
                    if self.file_tree.item(item)['values'][0] == current_file:
                        self.file_tree.set(item, "Статус", "Конвертация...")
                        
                        # Конвертируем файл
                        success = self.convert_file(current_file)
                        
                        # Обновляем статус
                        status = "Готово" if success else "Ошибка"
                        self.file_tree.set(item, "Статус", status)
                        break
                        
                processed_files += 1
                self.progress['value'] = (processed_files / total_files) * 100
                
            if not self.is_converting:
                # Если конвертация была отменена
                self.root.after(0, lambda: messagebox.showinfo("Отменено", "Конвертация была отменена!"))
            else:
                self.root.after(0, self.conversion_completed)
                
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.is_converting = False
            self.root.after(0, lambda: self.cancel_btn.configure(state="disabled"))
        
    def conversion_completed(self):
        self.convert_btn.configure(state="normal")
        self.select_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.progress['value'] = 0
        messagebox.showinfo("Готово", "Конвертаци завершена!")

    def stop_conversion(self):
        self.is_converting = False
        self.convert_btn.configure(state="normal")
        self.select_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.progress['value'] = 0

def resource_path(relative_path):
    """ Получить абсолютный путь к ресурсу """
    try:
        # PyInstaller создает временную папку и хранит путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfConverter(root)
    root.mainloop()
