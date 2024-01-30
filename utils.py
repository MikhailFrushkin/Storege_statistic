from loguru import logger
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def df_in_xlsx(df, filename, max_width=50):
    try:
        workbook = Workbook()
        sheet = workbook.active
        for row in dataframe_to_rows(df, index=False, header=True):
            sheet.append(row)
        for column in sheet.columns:
            column_letter = column[0].column_letter
            max_length = max(len(str(cell.value)) for cell in column)
            adjusted_width = min(max_length + 2, max_width)
            sheet.column_dimensions[column_letter].width = adjusted_width
        workbook.save(f"{filename}.xlsx")
    except Exception as ex:
        logger.error(f"Ошибка при записи файла {filename} {ex}")


class SearchProgress:
    def __init__(self, total_folders, progress_bar, current_folder=0):
        self.current_folder = current_folder
        self.total_folders = total_folders
        self.progress_bar = progress_bar

    def update_progress(self):
        self.current_folder += 1
        self.progress_bar.update_progress(self.current_folder, self.total_folders)

    def __str__(self):
        return str(self.current_folder)
