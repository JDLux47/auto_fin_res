from datetime import datetime
import customtkinter as ctk
import os
from tkinter import messagebox, filedialog
from models.excel_models import ExcelModel
from utils.dialogs import select_excel_file, show_file_error

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class FileSelectorFrame(ctk.CTkFrame):
    def __init__(self, parent, on_file_select):
        super().__init__(parent)
        self.on_file_select = on_file_select
        self.file_labels = []
        self.create_widgets()

    def create_widgets(self):
        files = ["Файл с ЗП сотрудников", "Файл с отчётом по марже", "Файл с отчётом по реализации нарядов спецами"]
        for i in range(3):
            label = ctk.CTkLabel(self, text=f"{files[i]}: Не выбран",
                                 font=ctk.CTkFont(size=14))
            label.pack(pady=8, padx=20, anchor="w")
            self.file_labels.append(label)

            btn = ctk.CTkButton(self, text="Выбрать Excel файл",
                                width=220, height=35,
                                command=lambda idx=i: self.select_file(idx))
            btn.pack(pady=5, padx=20)

    def select_file(self, index):
        file_path = select_excel_file(f"Выберите Excel файл {index + 1}")
        if not file_path:
            return

        filename = os.path.basename(file_path)
        success = self.on_file_select(index, file_path)

        if success:
            self.file_labels[index].configure(
                text=f"Файл {index + 1}: {filename}",
                text_color="green"
            )
        else:
            self.file_labels[index].configure(
                text=f"Файл {index + 1}: Ошибка формата",
                text_color="red"
            )
            show_file_error(filename)


class ExcelMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Автоматизация отчётности")
        self.geometry("900x750")
        self.model = ExcelModel()

        self.create_widgets()
        self.update_merge_button()
        self._load_params_to_fields()

    def _load_params_to_fields(self):
        """Загружает значения из params.json в поля"""
        try:
            fot_tax = round(self.model.fot_tax_pct * 100, 1)
            revenue_tax = round(self.model.revenue_tax_pct * 100, 1)
            fixed_costs = int(self.model.fixed_costs)

            self.fot_tax_entry.delete(0, ctk.END)
            self.fot_tax_entry.insert(0, str(fot_tax))

            self.revenue_tax_entry.delete(0, ctk.END)
            self.revenue_tax_entry.insert(0, str(revenue_tax))

            self.fixed_costs_entry.delete(0, ctk.END)
            self.fixed_costs_entry.insert(0, str(fixed_costs))

            print(f"Поля заполнены из params.json: {fot_tax}%, {revenue_tax}%, {fixed_costs:,}")
        except Exception as e:
            print(f"Ошибка загрузки параметров: {e}")

    def create_widgets(self):
        title = ctk.CTkLabel(self, text="Загрузите Excel файлы (.xlsx, .xls, .xlsm)",
                             font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=20)

        self.selector_frame = FileSelectorFrame(self, self.on_file_select)
        self.selector_frame.pack(pady=10, padx=20, fill="x")

        params_frame = ctk.CTkFrame(self)
        params_frame.pack(pady=20, padx=20, fill="x")

        params_title = ctk.CTkLabel(params_frame, text="Параметры расчётов:",
                                    font=ctk.CTkFont(size=16, weight="bold"))
        params_title.pack(pady=(15, 10))

        # Фрейм для полей
        fields_frame = ctk.CTkFrame(params_frame)
        fields_frame.pack(fill="x", padx=20, pady=5)

        # Налог на ФОТ %
        ctk.CTkLabel(fields_frame, text="Налог на ФОТ %:").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        self.fot_tax_entry = ctk.CTkEntry(fields_frame, placeholder_text="30.0", width=120)
        self.fot_tax_entry.grid(row=0, column=1, padx=10, pady=8, sticky="e")
        self.fot_tax_entry.insert(0, "30.0")  # Значение по умолчанию

        # Налог на выручку %
        ctk.CTkLabel(fields_frame, text="Налог на выручку %:").grid(row=1, column=0, sticky="w", padx=10, pady=8)
        self.revenue_tax_entry = ctk.CTkEntry(fields_frame, placeholder_text="20.0", width=120)
        self.revenue_tax_entry.grid(row=1, column=1, padx=10, pady=8, sticky="e")
        self.revenue_tax_entry.insert(0, "20.0")

        # Постоянные расходы
        ctk.CTkLabel(fields_frame, text="Постоянные расходы:").grid(row=2, column=0, sticky="w", padx=10, pady=8)
        self.fixed_costs_entry = ctk.CTkEntry(fields_frame, placeholder_text="500000", width=120)
        self.fixed_costs_entry.grid(row=2, column=1, padx=10, pady=8, sticky="e")
        self.fixed_costs_entry.insert(0, "500000")

        self.merge_btn = ctk.CTkButton(
            self,
            text="Сформировать отчёт",
            width=320,
            height=50,
            font=ctk.CTkFont(size=18, weight="bold"),
            command=self.merge_files,
            fg_color="green",
            hover_color="darkgreen",
            state="disabled"
        )
        self.merge_btn.pack(pady=30)

        self.status_label = ctk.CTkLabel(
            self,
            text="Загрузите 3 валидных Excel файла",
            text_color="gray",
            font=ctk.CTkFont(size=14)
        )
        self.status_label.pack(pady=10)

    def on_file_select(self, index: int, file_path: str) -> bool:
        """Только логика модели — НИКАКИХ лейблов!"""
        success = self.model.add_file(file_path, index)
        self.update_merge_button()
        return success

    def update_merge_button(self):
        """Обновляет состояние кнопки 'Сформировать отчёт'"""
        valid_count = self.model.get_valid_count()

        if valid_count == 3 and self.model.all_files_valid():
            self.merge_btn.configure(
                state="normal",
                text="Сформировать отчёт",
                fg_color="green"
            )
        else:
            self.merge_btn.configure(
                state="disabled",
                text=f"Загрузите {3 - valid_count} Excel файлов",
                fg_color="gray"
            )

    def save_report_to_desktop(self, excel_data, filename):
        """Диалог сохранения с выбором папки"""
        # Диалог "Сохранить как"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=filename,
            title="Сохранить отчет"
        )

        if filepath:  # Пользователь нажал OK
            try:
                with open(filepath, 'wb') as f:
                    f.write(excel_data)
                messagebox.showinfo("Готово!", f"Сохранено:\n{filepath}")
                os.startfile(filepath)  # Открыть Excel
            except PermissionError:
                messagebox.showerror("Ошибка", "Нет прав на запись в эту папку или перезаписываемый файл в данный момент открыт!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка сохранения:\n{str(e)}")

    def generate_report(self, employees, specialists, managers):
        result_list = self.model.create_result(employees, specialists, managers)  # Ваш список person

        excel_data = self.model.create_report(result_list)

        timestamp = datetime.now().strftime("%Y-%m-%d")
        filename = f"Отчет_Сотрудники_{timestamp}.xlsx"

        # Сохранение файла на рабочий стол
        self.save_report_to_desktop(excel_data, filename)

    def merge_files(self):
        """Обрабатывает все файлы"""
        if not self.model.all_files_valid():
            messagebox.showerror("Ошибка", "Загрузите все 3 Excel файла!")
            return

        try:
            fot_tax = float(self.fot_tax_entry.get())
            revenue_tax = float(self.revenue_tax_entry.get())
            fixed_costs = float(self.fixed_costs_entry.get())

            # Сохраняем в модель (и в params.json автоматически)
            self.model.set_parameters(fot_tax, revenue_tax, fixed_costs)

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числа в поля параметров!")
            return

        # Получаем список сотрудников из первого файла
        employees = self.model.get_employees()
        specialists = self.model.get_specialists()
        managers = self.model.get_managers()

        # print("==========Сотрудники==========")
        # for i, emp in enumerate(employees, 1):
        #     print(f"{i:2d}. {emp['name']:<25} {emp['salary']}")
        #
        # print("==========Спецы==========")
        # for i, emp in enumerate(specialists, 1):
        #     print(f"{i:2d}. {emp['name']:<25} {emp['sum']}")

        # print("==========Менеджеры==========")
        # for i, manager in enumerate(managers, 1):
        #     print(f"{i:2d}. {manager['name']} | " f"{manager['total_price']} | " f"{manager['total_cost']}")
        #
        #     for cat in manager['categories']:
        #         print(f"    └─ {cat['name']} | {cat['price']} | {cat['cost']}")
        #     print()

        self.generate_report(employees, specialists, managers)

        result_list = self.model.create_result(employees, specialists, managers)

        for i, res in enumerate(result_list, 1):
            print(f"{res['name']} {res['categories']} {res['sum']} {res['cost_price']} {res['margin']} {res['salary']} {res['salary_tax']} {res['sum_tax']} {res['reg_costs']} {res['other_costs_margin']} {res['res_costs']} {res['profit_month']} {res['profit_prc']}")
