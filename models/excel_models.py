import io
import json
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from models.nomenclature import nomenclature_list


class ExcelModel:
    def __init__(self):
        self.file_paths = []
        self.dataframes = []
        self.is_valid = []

        self.fot_tax_pct = 0.0
        self.revenue_tax_pct = 0.0
        self.fixed_costs = 0.0

        self.params_file = "models/params.json"
        self._load_params()

    def add_file(self, file_path: str, index: int) -> bool:
        """Универсальное чтение Excel"""
        while len(self.dataframes) <= index:
            self.dataframes.append(None)
            self.file_paths.append("")
            self.is_valid.append(False)

        if not self._is_excel_file(file_path):
            self.is_valid[index] = False
            return False

        try:
            df = pd.read_excel(file_path, engine=None)
            df = df.dropna(how='all').dropna(axis=1, how='all')

            # Заменяем по индексу
            self.dataframes[index] = df
            self.file_paths[index] = file_path
            self.is_valid[index] = True

            print(f"Файл {index + 1}: {len(df)} строк")
            return True

        except Exception as e:
            print(f"Файл {index + 1}: {e}")
            self.is_valid[index] = False
            return False

    def clear_files(self):
        self.dataframes.clear()
        self.file_paths.clear()
        self.is_valid.clear()

    def _is_excel_file(self, file_path: str) -> bool:
        ext = os.path.splitext(file_path.lower())[1]
        return ext in ['.xlsx', '.xls', '.xlsm']

    def _is_valid_fio(self, text: str) -> bool:
        """ровно 3 слова, каждое начинается с большой буквы"""
        words = text.strip().split()

        # Ровно 3 слова
        if len(words) != 3:
            return False

        # Каждое слово: Заглавная + маленькие (Title Case)
        for word in words:
            if not word or not (
                    word[0].isupper() and  # Первая буква большая
                    len(word) > 1 and  # Более 1 символа
                    all(c.islower() for c in word[1:])  # Остальные маленькие
            ):
                return False

        return True

    def _is_category(self, text: str) -> bool:
        """
        Проверка на категорию
        """
        text_clean = text.strip()
        return text_clean in nomenclature_list

    def all_files_valid(self) -> bool:
        """Все ли 3 файла Excel валидные"""
        return len(self.is_valid) == 3 and all(self.is_valid)

    def get_valid_count(self) -> int:
        """Считает валидные Excel файлы"""
        return sum(1 for valid in self.is_valid if valid)

    def get_employees(self, file_index: int = 0) -> list:
        """
        Извлекает сотрудников и их ЗП
        """
        df = self.dataframes[file_index]

        employees = []

        # Фиксированные столбцы
        name_col = df.columns[0]
        salary_col = df.columns[3]

        for idx, row in df.iterrows():
            name_raw = str(row[name_col]).strip()

            # проверка ФИО ровно 3 слова, каждое с большой буквы
            if self._is_valid_fio(name_raw):
                salary = self._parse_salary(row[salary_col])

                employees.append({
                    'name': name_raw,
                    'salary': salary
                })

        print(f"Найдено {len(employees)} сотрудников (3 слова, большая буква)")
        return employees

    def get_specialists(self, file_index: int = 2) -> list:
        """
        Извлекает спецов по ФИО и сумму их реализации
        """
        df = self.dataframes[file_index].copy()

        specialists = []

        for idx in range(len(df) - 1):
            name_raw = str(df.iloc[idx, 0]).strip()
            next_row_idx = idx + 1
            next_row_col5 = df.iloc[next_row_idx, 4] if next_row_idx < len(df) else None

            is_next_col5_empty = (pd.isna(next_row_col5) or str(next_row_col5).strip() == '')

            # ФИО 3 слова с большой буквы
            if name_raw and len(name_raw.split()) >= 2 and is_next_col5_empty and self._is_valid_fio(name_raw):
                sum_raw = self._parse_salary(df.iloc[idx, 10])

                specialists.append({
                    'name': name_raw,
                    'sum': sum_raw if pd.notna(sum_raw) else 0,
                })

        print(f"Файл {file_index + 1}: Найдено {len(specialists)} сотрудников с услугами")
        return specialists

    def get_managers(self, file_index: int = 1) -> list:
        """
        Извлекает менеджеров по ФИО, услуги, себестоимость и стоимость без ндс
        """
        df = self.dataframes[file_index].copy()

        name_col = df.columns[0]
        price_col = df.columns[2]
        cost_col = df.columns[3]

        managers = []
        current_manager = None

        for idx, row in df.iterrows():
            name_raw = str(row[name_col]).strip()
            if pd.isna(name_raw) or not name_raw:
                continue

            price = self._parse_salary(row[price_col]) or 0
            cost = self._parse_salary(row[cost_col]) or 0

            if self._is_valid_fio(name_raw):
                if current_manager:
                    managers.append(current_manager)

                current_manager = {
                    'name': name_raw,
                    'total_price': price if pd.notna(price) else 0,
                    'total_cost': cost if pd.notna(cost) else 0,
                    'categories': []
                }

            elif current_manager and self._is_category(name_raw):

                service = {
                    'name': name_raw,
                    'price': price if pd.notna(price) else 0,
                    'cost': cost if pd.notna(cost) else 0
                }
                current_manager['categories'].append(service)

        # Последний менеджер
        if current_manager:
            managers.append(current_manager)

        print(f"Найдено {len(managers)} менеджеров")
        for m in managers:
            print(f"   {m['name']}: {len(m['categories'])} услуг из списка")
        return managers

    def _parse_salary(self, salary_raw: str) -> float:
        """Парсит ЗП"""
        salary_raw = str(salary_raw).replace(' ', '').replace(',', '.').replace('₽', '')
        try:
            return float(salary_raw)
        except:
            return 0.0

    def _load_params(self):
        """Загружает ваш params.json"""
        if os.path.exists(self.params_file):
            try:
                with open(self.params_file, 'r', encoding='utf-8') as f:
                    params = json.load(f)
                    self.fot_tax_pct = params.get("fot_tax") / 100
                    self.revenue_tax_pct = params.get("revenue_tax") / 100
                    self.fixed_costs = params.get("fixed_costs")
            except Exception as error:
                print(error)

    def set_parameters(self, fot_tax: float, revenue_tax: float, fixed_costs: float):
        """Сохраняет параметры в JSON (ваш формат)"""
        self.fot_tax_pct = fot_tax / 100
        self.revenue_tax_pct = revenue_tax / 100
        self.fixed_costs = fixed_costs

        params = {
            "fot_tax": fot_tax,
            "revenue_tax": revenue_tax,
            "fixed_costs": fixed_costs
        }

        try:
            with open(self.params_file, 'w', encoding='utf-8') as f:
                json.dump(params, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Ошибка сохранения params.json: {e}")

    def create_result(self, employees: list, specialists: list, managers: list) -> list:
        """
        Объединяет списки в один общий
        """
        result_list = []

        for emp in employees:
            person = {
                'name': emp['name'],
                'categories': [],
                'sum': 0,
                'cost_price':0,
                'margin':0,
                'salary': emp['salary'],
                'salary_tax': emp['salary'] * self.fot_tax_pct,
                'sum_tax': None,
                'reg_costs': self.fixed_costs,
                'other_costs_margin': None,
                'res_costs': None,
                'profit_month':None,
                'profit_prc':None
            }

            match_spec = next((s for s in specialists if s['name'] == emp['name']), None)
            match_manager = next((s for s in managers if s['name'] == emp['name']), None)

            if match_manager:
                person['sum'] = match_manager['total_price']
                person['cost_price'] = match_manager['total_cost']
                person['categories'] = match_manager['categories']

            if match_spec:
                person['sum'] = match_spec['sum']
                person['cost_price'] = 0
                person['categories'] = []

            person['margin'] = person['sum'] - person['cost_price']
            person['sum_tax'] = round(person['sum'] * self.revenue_tax_pct, 1)
            person['res_costs'] = person['salary'] + person['salary_tax'] + person['sum_tax'] + person['reg_costs']
            person['profit_month'] = round(person['margin'] - person['res_costs'], 2)
            person['profit_prc'] = round(person['profit_month'] / person['sum'] * 100, 1) if person['sum'] != 0 else None

            result_list.append(person)

            result_list.sort(key=lambda x: (len(x['categories']) == 0, x['name']))

        return result_list

    def create_report(self, result_list):
        """Создает Excel с сотрудниками, заголовками и форматированием"""
        output = io.BytesIO()

        wb = Workbook()
        ws = wb.active
        ws.title = "Финансовая модель ДИТ. Отчёт"

        # Заголовки
        headers = [
            "Сотрудник", "Выручка (реализации)", "Себестоимость", "Маржа",
            "ФОТ", "Налоги на ФОТ %", "Налоги на выручку %", "Постоянные расходы",
            "Итого затраты", "Рент-сть текущего месяца", "% рент-сти"
        ]

        # Названия колонок
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = PatternFill(start_color="D9D9D9", fill_type="solid")

        ws.cell(row=2, column=6, value=self.fot_tax_pct).number_format = '0.0%'
        ws.cell(row=2, column=7, value=self.revenue_tax_pct).number_format = '0.0%'
        ws.cell(row=2, column=8, value=self.fixed_costs)

        row_num = 3

        for person in result_list:
            # Строка ФИО
            fio_row = row_num

            # Данные строки
            data_row = [
                person['name'],
                person.get('sum', 0),
                person.get('cost_price', 0),
                '',
                person.get('salary', 0),
                '',
                '',
                '',
                '',
                '',
                ''
            ]

            # Заполняем строку ФИО
            for col, value in enumerate(data_row, 1):
                cell = ws.cell(row=fio_row, column=col, value=value)

                # Бледно-жёлтый для всей строки ФИО
                if col == 1:  # ФИО - жирный
                    cell.font = Font(bold=True, size=11)
                    cell.alignment = Alignment(horizontal="center")
                else:  # Числа
                    cell.number_format = '#,##0.0'

                # Бледно-жёлтый фон для всех ячеек строки
                cell.fill = PatternFill(start_color="FFF2CC", fill_type="solid")

                # Маржа = B - C
                ws.cell(row=row_num, column=4).value = f"=B{row_num}-C{row_num}"
                # Налоги ФОТ = E * $F$2
                ws.cell(row=row_num, column=6).value = f"=E{row_num}*$F$2"
                # Налоги выручка = B * $G$2
                ws.cell(row=row_num, column=7).value = f"=B{row_num}*$G$2"
                # Постоянные расходы = $H$2
                ws.cell(row=row_num, column=8).value = "=$H$2"
                # Итого затраты = E+F+G+H
                ws.cell(row=row_num, column=9).value = f"=SUM(E{row_num}:H{row_num})"
                # Рентабельность = D - I
                ws.cell(row=row_num, column=10).value = f"=D{row_num}-I{row_num}"
                # % рент. = безопасная формула
                ws.cell(row=row_num, column=11).value = f"=IF(B{row_num}=0,0,J{row_num}/B{row_num}*100)"

            # Categories под ФИО
            if person.get('categories') and len(person['categories']) > 0:
                row_num += 1  # Пустая строка
                for cat in person['categories']:
                    cat_row = row_num

                    # Категория: только первые 4 колонки
                    ws.cell(row=cat_row, column=1, value=f"{cat['name']}")
                    ws.cell(row=cat_row, column=2, value=cat.get('price', 0) or 0)
                    ws.cell(row=cat_row, column=3, value=cat.get('cost', 0) or 0)
                    ws.cell(row=cat_row, column=4, value=(cat.get('price', 0) - cat.get('cost', 0)) or 0)

                    # Форматирование чисел категорий
                    for col in [2, 3, 4]:
                        cell = ws.cell(row=cat_row, column=col)
                        cell.number_format = '#,##0'

                    row_num += 1
            else:
                row_num += 1

        # Табличные границы
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Применяем границы ко всем ячейкам с данными
        for row in ws.iter_rows(min_row=1, max_row=row_num - 1, min_col=1, max_col=11):
            for cell in row:
                cell.border = thin_border

        # Автоподгонка ширины колонок
        for column_cells in ws.columns:
            length = max(len(str(cell.value or "")) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 25)

        wb.save(output)
        output.seek(0)
        return output.getvalue()
