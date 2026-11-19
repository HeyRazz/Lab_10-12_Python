import os
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


class ExcelDataManager:

  def __init__(self, file_path):
    self.file_path = file_path
    self.df = None

  def load_data(self):
    if not os.path.exists(self.file_path):
      raise FileNotFoundError(
          f"Помилка: Файл '{self.file_path}' відсутній у директорії програми."
      )

    try:
      if self.file_path.endswith(".csv"):
        self.df = pd.read_csv(self.file_path, sep=";")
      elif self.file_path.endswith((".xlsx", ".xls")):
        self.df = pd.read_excel(self.file_path)
      else:
        raise ValueError("Непідтримуваний формат файлу.")

      print(f"Дані успішно завантажено з файлу: {self.file_path}")
    except Exception as e:
      raise RuntimeError(f"Помилка під час читання файлу: {e}")

  def save_styled_excel(self, data_to_save, output_path):
    try:
      data_to_save.to_excel(output_path, index=False, sheet_name="Звіт")

      from openpyxl import load_workbook

      wb = load_workbook(output_path)
      ws = wb.active

      header_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
      header_fill = PatternFill(
          start_color="366092", end_color="366092", fill_type="solid"
      )
      align_center = Alignment(horizontal="center", vertical="center")
      thin_border = Border(
          left=Side(style="thin", color="D9D9D9"),
          right=Side(style="thin", color="D9D9D9"),
          top=Side(style="thin", color="D9D9D9"),
          bottom=Side(style="thin", color="D9D9D9"),
      )

      for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

      for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
          cell = ws.cell(row=row, column=col)
          cell.border = thin_border
          if row > 1 and isinstance(cell.value, (int, float)):
            cell.alignment = Alignment(horizontal="right", vertical="center")
          elif row > 1:
            cell.alignment = Alignment(horizontal="left", vertical="center")

          if cell.value:
            max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

      wb.save(output_path)
      print(f"Стилізований звіт успішно збережено у файл: {output_path}")
    except Exception as e:
      print(f"Помилка, не вдалося зберегти стилізований Excel: {e}")


class DataAnalyzer:

  def __init__(self, dataframe):
    self.df = dataframe

  def filter_data(self, column_name, threshold):
    try:
      filtered_df = self.df[self.df[column_name] > threshold]
      return filtered_df
    except Exception as e:
      raise ValueError(f"Помилка фільтрації: {e}")

  def sort_data(self, column_name, ascending=True):
    try:
      sorted_df = self.df.sort_values(by=column_name, ascending=ascending)
      return sorted_df
    except Exception as e:
      raise ValueError(f"Помилка сортування: {e}")

  def get_statistics(self, column_name):
    try:
      col_data = self.df[column_name]
      stats = {
          "Середнє": col_data.mean(),
          "Мінімум": col_data.min(),
          "Максимум": col_data.max(),
      }
      return stats
    except Exception as e:
      raise ValueError(f"Помилка розрахунку статистики: {e}")


def main():
  input_file = "results_full.csv"
  output_file = "macro_analysis_report.xlsx"

  try:
    manager = ExcelDataManager(input_file)
    manager.load_data()
  except FileNotFoundError as fnf_err:
    print(f"Помилка {fnf_err}")
    print("Будь ласка, переконайтеся, що файл знаходиться у папці з програмою.")
    return
  except Exception as err:
    print(f"Помилка {err}")
    return

  analyzer = DataAnalyzer(manager.df)

  while True:
    print("\nМеню")
    print("1. Вивести перші рядки даних на екран")
    print("2. Показати доступні стовпці таблиці")
    print("3. Фільтрувати дані за числовим показником")
    print("4. Сортувати дані за стовпцем")
    print("5. Розрахувати статистику по стовпцю")
    print("6. Зберегти результати у новий стилізований Excel-файл (.xlsx)")
    print("7. Вихід")

    choice = input("Оберіть номер дії (1-7): ").strip()

    if choice == "1":
      print("\nПерші рядки таблиці даних")
      print(manager.df.head(10).to_string(index=False))

    elif choice == "2":
      print("\nДоступні стовпці у файлі")
      for col in manager.df.columns:
        print(f" - {col}")

    elif choice == "3":
      print("\nДоступні стовпці:", list(manager.df.columns))
      col = input("Введіть точну назву стовпця для фільтрації: ").strip()
      try:
        val = float(input("Введіть порогове значення: "))
        res_df = analyzer.filter_data(col, val)
        print(f"\nРезультат фільтрації ({col} > {val})")
        print(
            res_df.head(15).to_string(index=False)
            if not res_df.empty
            else "Записи за такими критеріями не знайдено."
        )
      except ValueError as ex:
        print(f"Помилка вводу {ex}")

    elif choice == "4":
      print("\nДоступні стовпці для сортування:", list(manager.df.columns))
      col = input("Введіть назву стовпця: ").strip()
      order = (
          input("Сортувати за зростанням? (так/ні): ").strip().lower() == "так"
      )
      try:
        res_df = analyzer.sort_data(col, ascending=order)
        print("\n Відсортовані дані")
        print(res_df.head(15).to_string(index=False))
      except Exception as ex:
        print(f"Помилка {ex}")

    elif choice == "5":
      print("\nДоступні стовпці:", list(manager.df.columns))
      col = input("Введіть числовий стовпець для аналізу: ").strip()
      try:
        stats = analyzer.get_statistics(col)
        print(f"\nСтатистика для стовпця '{col}' ")
        for k, v in stats.items():
          print(f"{k}: {v:.2f}" if isinstance(v, float) else f"{k}: {v}")
      except Exception as ex:
        print(f"Помилка {ex}")

    elif choice == "6":
      manager.save_styled_excel(manager.df, output_file)

    elif choice == "7":
      print("Роботу програми завершено.")
      break
    else:
      print("Помилка, невірний вибір. Будь ласка, оберіть від 1 до 7.")


if __name__ == "__main__":
  main()