from tkinter import filedialog, messagebox
import os

def select_excel_file(title: str = "Выберите Excel файл") -> str:
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls"), ("Все файлы", "*.*")]
    )

def save_excel_file(title: str = "Сохранить файл") -> str:
    return filedialog.asksaveasfilename(
        title=title,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

def show_success(filename: str, rows: int):
    messagebox.showinfo("Успех!", f"Файл сохранен:\n{filename}\n\nСтрок: {rows}")

def show_error(message: str):
    messagebox.showerror("Ошибка", message)

def show_file_error(filename: str):
    messagebox.showerror("Неверный формат!",
                       f"Файл '{os.path.basename(filename)}' не является Excel файлом!\n\n"
                       "Пожалуйста, выберите файл с расширением .xlsx или .xls")
