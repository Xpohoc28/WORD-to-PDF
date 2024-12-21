import os
import sys

# Константы для режимов конвертации
DOCX_TO_PDF = "docx_to_pdf"
PDF_TO_DOCX = "pdf_to_docx"

# Константы для форматов файлов
FILE_TYPES = {
    DOCX_TO_PDF: [("Word files", "*.docx")],
    PDF_TO_DOCX: [("PDF files", "*.pdf")]
}

# Константы для заголовков
TITLES = {
    DOCX_TO_PDF: "Выберите файлы Word",
    PDF_TO_DOCX: "Выберите PDF файлы"
}

# Константы для статусов
STATUS_WAITING = "Ожидает"
STATUS_CONVERTING = "Конвертация..."
STATUS_DONE = "Готово"
STATUS_ERROR = "Ошибка"

def resource_path(relative_path):
    """Получить абсолютный путь к ресурсу"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path) 