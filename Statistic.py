import datetime
import os
import subprocess
import time
from pathlib import Path
from pprint import pprint
from peewee import fn
import pandas as pd

from bd import Operations
from utils import df_in_xlsx, SearchProgress
import qdarkstyle
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import (
    QFileDialog, QCheckBox, QProgressBar, QMessageBox
)
from loguru import logger


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(853, 689)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_2.addWidget(self.lineEdit)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout_2.addWidget(self.pushButton)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.gridLayout_3.addLayout(self.verticalLayout, 0, 0, 1, 1)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.verticalLayout_4.addWidget(self.label_4)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_3.addWidget(self.lineEdit_2)
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_4.setObjectName("pushButton_4")
        self.horizontalLayout_3.addWidget(self.pushButton_4)
        self.verticalLayout_4.addLayout(self.horizontalLayout_3)
        self.gridLayout_3.addLayout(self.verticalLayout_4, 1, 0, 1, 1)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout_3.addItem(spacerItem, 2, 0, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout_3.addWidget(self.pushButton_2, 3, 0, 1, 1)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Myanmar Text")
        font.setPointSize(12)
        font.setBold(True)
        font.setItalic(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        font.setKerning(True)
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        self.label_2.setFont(font)
        self.label_2.setMouseTracking(False)
        self.label_2.setTabletTracking(False)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2)
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")

        self.verticalLayout_2.addLayout(self.gridLayout_2)
        self.gridLayout_3.addLayout(self.verticalLayout_2, 4, 0, 1, 1)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Myanmar Text")
        font.setPointSize(12)
        font.setBold(True)
        font.setItalic(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        font.setKerning(True)
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        self.label_3.setFont(font)
        self.label_3.setMouseTracking(False)
        self.label_3.setTabletTracking(False)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_3.addWidget(self.label_3)
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")

        self.verticalLayout_3.addLayout(self.gridLayout)
        self.gridLayout_3.addLayout(self.verticalLayout_3, 5, 0, 1, 1)
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout_3.addWidget(self.pushButton_3, 6, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Загрузите файл \"Отслеживания заданий ТСД\""))
        self.pushButton.setText(_translate("MainWindow", "Загрузить файл"))
        self.label_4.setText(_translate("MainWindow", "Загрузите файлы \"Перемещений и локалок\""))
        self.pushButton_4.setText(_translate("MainWindow", "Загрузить файлы"))
        self.pushButton_2.setText(_translate("MainWindow", "Создать базу"))
        self.label_2.setText(_translate("MainWindow", "Выберите операции для отображения"))

        self.label_3.setText(_translate("MainWindow", "Выберите сотрудников для отображения"))

        self.pushButton_3.setText(_translate("MainWindow", "Сформировать статистику"))


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.current_dir = Path.cwd()
        self.move(550, 100)
        self.column_counter_user = 0
        self.count_user = 0

        self.column_counter_oper = 0
        self.count_oper = 0

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(10, 10, 100, 25)
        self.progress_bar.setMaximum(100)
        self.statusbar.addWidget(self.progress_bar, 1)

        self.pushButton.clicked.connect(self.evt_btn_open_file_clicked)
        self.pushButton_4.clicked.connect(self.evt_btn_open_files_clicked)
        self.pushButton_2.clicked.connect(self.evt_btn_update_db)
        self.pushButton_3.clicked.connect(self.getCheckboxList)
        if Operations.table_exists():
            self.old_db()

    def evt_btn_open_file_clicked(self):
        """Ивент на кнопку загрузить файл"""
        res, _ = QFileDialog.getOpenFileName(self, 'Загрузить файл', str(self.current_dir), 'Лист XLSX (*.xlsx)')
        if res:
            self.lineEdit.setText(res)

    def evt_btn_open_files_clicked(self):
        """Ивент на кнопку загрузить файлы"""
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_names, _ = QFileDialog.getOpenFileNames(self, 'Загрузить файлы', str(self.current_dir),
                                                     'Лист XLSX (*.xlsx)', options=options)

        if file_names:
            for file_path in file_names:
                print(f"Выбранный файл: {file_path}")

            self.lineEdit_2.setText(";".join(file_names))  # Показывает пути разделенные символом ";"
            union_df(file_names)

    def old_db(self):
        self.label_2.setVisible(True)
        self.label_3.setVisible(True)

        try:
            user_list = Operations.get_unique_user()
            for user in user_list:
                self.addUserCheckbox(str(user))
        except Exception as ex:
            logger.error(ex)

        try:
            oper_list = Operations.get_unique_type_oper()
            for oper in oper_list:
                self.addOperCheckbox(str(oper))
        except Exception as ex:
            logger.error(ex)

    def evt_btn_update_db(self):
        if self.lineEdit.text():
            self.label_2.setVisible(False)
            self.label_3.setVisible(False)
            widget = self.verticalLayout_3
            while self.gridLayout.count():
                item = self.gridLayout.takeAt(0)
                widget.layout().removeItem(item)
                widget.layout().removeWidget(item.widget())
            widget.update()

            widget = self.verticalLayout_2
            while self.gridLayout_2.count():
                item = self.gridLayout_2.takeAt(0)
                widget.layout().removeItem(item)
                widget.layout().removeWidget(item.widget())
            widget.update()
            self.column_counter_user = 0
            self.count_user = 0

            self.column_counter_oper = 0
            self.count_oper = 0
            time.sleep(1)

            if created_db(self.lineEdit.text(), self):
                self.label_2.setVisible(True)
                self.label_3.setVisible(True)

                try:
                    user_list = Operations.get_unique_user()
                    for user in user_list:
                        self.addUserCheckbox(str(user))
                except Exception as ex:
                    logger.error(ex)

                try:
                    oper_list = Operations.get_unique_type_oper()
                    for oper in oper_list:
                        self.addOperCheckbox(str(oper))
                except Exception as ex:
                    logger.error(ex)
        else:
            QMessageBox.information(self, 'Инфо', 'Загрузите файл')

    def addOperCheckbox(self, oper):
        checkbox = QCheckBox(oper, self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        checkbox.setFont(font)
        checkbox.setObjectName(oper)
        checkbox.setChecked(True)
        self.gridLayout_2.addWidget(checkbox, self.column_counter_oper, self.count_oper)
        self.count_oper += 1
        if self.count_oper == 2:
            self.column_counter_oper += 1
            self.count_oper = 0

    def addUserCheckbox(self, user):
        checkbox = QCheckBox(user, self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        checkbox.setFont(font)
        checkbox.setObjectName(user)
        checkbox.setChecked(True)
        self.gridLayout.addWidget(checkbox, self.column_counter_user, self.count_user)
        self.count_user += 1
        if self.count_user == 7:
            self.column_counter_user += 1
            self.count_user = 0

    def update_progress(self, current_value, total_value):
        progress = int(current_value / total_value * 100)
        self.progress_bar.setValue(progress)
        QApplication.processEvents()

    def getCheckboxList(self):
        checkbox_list_user = []
        for row in range(self.gridLayout.rowCount()):
            for column in range(self.gridLayout.columnCount()):
                item = self.gridLayout.itemAtPosition(row, column)
                if item is not None:
                    widget = item.widget()
                    if isinstance(widget, QCheckBox) and widget.isChecked():
                        checkbox_list_user.append(int(widget.text()))

        checkbox_list_oper = []
        for row in range(self.gridLayout_2.rowCount()):
            for column in range(self.gridLayout_2.columnCount()):
                item = self.gridLayout_2.itemAtPosition(row, column)
                if item is not None:
                    widget = item.widget()
                    if isinstance(widget, QCheckBox) and widget.isChecked():
                        checkbox_list_oper.append(widget.text())
        if len(checkbox_list_user) == 0:
            QMessageBox.warning(self, 'Галочки', f'Не выбранны сотрудники')
        elif len(checkbox_list_oper) == 0:
            QMessageBox.warning(self, 'Галочки', f'Не выбранны операции')
        else:
            try:
                filename = filter_db(checkbox_list_user, checkbox_list_oper)
            except Exception as ex:
                logger.error(ex)
            open_file(f'{self.current_dir}/{filename}')
            return checkbox_list_user

    def restart(self):
        os.execl(sys.executable, os.path.abspath(__file__), *sys.argv)


def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


def good_df(df):
    columns_list = ['ИД документа', 'Название документа', 'Тип документа', 'Склад', 'Объём задания',
                    'Количество строк в задании', 'Время создания', 'Время завершения', 'Статус', 'Исполнитель']
    data_types = {
        'Название документа': str,
        'Тип документа': str,
        'Склад': int,
        'Объём задания': int,
        'Количество строк в задании': int,
        'Время создания': 'datetime64[ns]',
        'Время завершения': 'datetime64[ns]',
        'Статус': str,
        'Исполнитель': str
    }
    columns_list_file = df.columns.tolist()
    bad_columns = [i for i in columns_list_file if i not in columns_list]
    good_list = [i for i in columns_list if i not in columns_list_file]
    for i in good_list:
        df[i] = None
    df = df.drop(columns=bad_columns, axis=True)
    df = df[(df['Статус'] == 'Завершено') & ~df['Исполнитель'].isna() & ~df['Исполнитель'].isnull()]
    try:
        df = df.astype(data_types)
    except Exception as ex:
        logger.error(f"Ошибка при применении типов {ex}")
    df = df[(df['Исполнитель'].str.len() > 5)]

    try:
        df['Время создания'] = pd.to_datetime(df['Время создания'])
        df['Время завершения'] = pd.to_datetime(df['Время завершения'])

        # Вычисление разницы между столбцами и запись в новый столбец "Время выполнения"
        df['Время выполнения'] = df['Время завершения'] - df['Время создания']

        # Перемещение столбца "Время выполнения" после столбца "Время завершения"
        column_name = 'Время выполнения'
        position = df.columns.get_loc('Время завершения') + 1
        df.insert(position, column_name, df.pop(column_name))
    except Exception as ex:
        logger.error(f"Ошибка при вычисления времени выполнения операции {ex}")

    df.loc[(~df['Название документа'].str.contains('шт')) &
           (~df['Название документа'].str.contains('->')) &
           (df['Тип документа'] == 'Внутрискладское перемещение'),
    'Тип документа'] = 'Перенос'

    df_in_xlsx(df, '!После обработки')

    return df


def created_db(filename='DYNB450.tmp.xlsx', self=None):
    try:
        df = good_df(pd.read_excel(filename).fillna(0))
    except Exception as ex:
        logger.error(f'Ошибка при чтение Excel {ex}')
        if self:
            QMessageBox.warning(self, 'Ошибка при чтение Excel', f'Ошибка при чтение Excel {ex}')
        return
    Operations.drop_table()
    if not Operations.table_exists():
        Operations.create_table(safe=True)
    logger.success(df.columns)
    logger.success(len(df))
    if self:
        progress = SearchProgress(len(df), self)
    for index, row in df.iterrows():
        if self:
            progress.update_progress()
        new_record = Operations(
            id_doc=row['ИД документа'],
            name=row['Название документа'],
            type_oper=row['Тип документа'],
            storage=row['Склад'],
            v=row['Объём задания'],
            count=row['Количество строк в задании'],
            created_at=row['Время создания'].to_pydatetime(),
            finish_at=row['Время завершения'].to_pydatetime(),
            lead_time=int(row['Время выполнения'].total_seconds()),
            user=row['Исполнитель']
        )
        new_record.save()
    return True


def filter_db(checkbox_list_user=None, checkbox_list_oper=None):
    min_date = Operations.select(fn.MIN(Operations.finish_at)).where(Operations.finish_at.year > 2022).scalar()
    max_date = Operations.select(fn.MAX(Operations.finish_at)).scalar()
    min_date = min_date.date()
    max_date = max_date.date()
    # Создание Excel-файла и объекта писателя (writer)
    writer = pd.ExcelWriter(f'!Статистика с {min_date} по {max_date}.xlsx', engine='xlsxwriter')
    workbook = writer.book
    wrap_format = workbook.add_format({'text_wrap': True, 'font_name': 'Arial', 'font_size': 14, 'border': 1,
                                       'bg_color': '#abcef5', 'align': 'center'})
    cell_format = workbook.add_format(
        {'align': 'center', 'valign': 'top', 'font_size': 12, 'text_wrap': True, 'bold': True})

    query = Operations.select(Operations.user, Operations.type_oper, fn.COUNT(Operations.id).alias('operation_count'),
                              fn.SUM(Operations.v).alias('v_sum'),
                              fn.SUM(Operations.count).alias('count_sum')).where(
        (Operations.type_oper << checkbox_list_oper) &
        (Operations.user << checkbox_list_user)).group_by(Operations.user,
                                                          Operations.type_oper)
    results = query.dicts()
    df = pd.DataFrame(results)
    pivot_df = df.pivot(index='user', columns='type_oper', values='operation_count').fillna(0)
    pivot_df['Всего'] = pivot_df.sum(axis=1)
    pivot_df['Объем'] = df.groupby('user')['v_sum'].sum()
    pivot_df['Всего строк'] = df.groupby('user')['count_sum'].sum()
    pivot_df_main = pivot_df.reset_index().rename(columns={'user': 'Сотрудник'})
    worksheet = workbook.add_worksheet('Главный')
    max_col = len(pivot_df_main.columns.tolist())
    max_row = len(pivot_df_main)
    print(max_col, max_row)
    pivot_table = table_df_create(pivot_df_main)
    set_column2(worksheet=worksheet, cell_format=cell_format, max_col=max_col)
    pivot_table.to_excel(writer, sheet_name='Главный', startrow=0, startcol=0, header=True, index=False, na_rep='')
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    query = (Operations
             .select(Operations.user, Operations.type_oper, fn.COUNT(Operations.id).alias('total_operations'),
                     fn.Sum(Operations.v).alias('total_v'), fn.Sum(Operations.count).alias('total_count'),
                     fn.AVG(Operations.v).alias('average_v'), fn.AVG(Operations.count).alias('average_count'))
             .where((Operations.type_oper << checkbox_list_oper) &
                    (Operations.user << checkbox_list_user))
             .group_by(Operations.user, Operations.type_oper))

    result = []
    for row in query:
        result.append({
            'Сотрудник': row.user,
            'Тип операции': row.type_oper,
            'Количество операций': row.total_operations,
            'Сумма по объему': row.total_v,
            'Сумма по количеству строк': row.total_count,
            'Среднее значение по объему': round(row.average_v, 2),
            'Среднее значение по количеству строк': round(row.average_count, 2)
        })

    df = pd.DataFrame(result)
    df_in_xlsx(df, '!Общий файл статистики')
    len_columns = len(df['Тип операции'].unique().tolist())

    # Список столбцов, для которых нужно создать отдельные таблицы
    columns = ['Количество операций', 'Сумма по объему', 'Сумма по количеству строк', 'Среднее значение по объему',
               'Среднее значение по количеству строк']
    # Запись каждой таблицы на отдельный лист в файл Excel
    row_num = 0
    worksheet = workbook.add_worksheet('Расширенная')

    for column in columns:
        worksheet.merge_range(row_num, 0, row_num, len_columns + 1, column, wrap_format)
        pivot_table = df.pivot_table(index='Сотрудник', columns='Тип операции', values=column, fill_value=0)
        if row_num == 0:
            header_row = ['Сотрудник']
            header_row.extend(pivot_table.columns.tolist())
            header_row.append('Всего')
        row_num += 1

        if column in ['Среднее значение по объему', 'Среднее значение по количеству строк']:
            pivot_table['Всего'] = pivot_table.mean(axis=1)
        else:
            pivot_table['Всего'] = pivot_table.sum(axis=1)
        max_col = len(pivot_table.columns.tolist())

        pivot_table = pivot_table.reset_index()
        pivot_table_rows = pivot_table.shape[0]

        pivot_table = table_df_create(pivot_table)
        set_column2(worksheet=worksheet, cell_format=cell_format, max_col=max_col)
        pivot_table.to_excel(writer, sheet_name='Расширенная', startrow=row_num, startcol=0, header=True, index=False,
                             na_rep='')
        worksheet.write_row(f'A{row_num + 1}', header_row, cell_format)

        row_num += pivot_table_rows + 2
    writer.close()
    return f'!Статистика с {min_date} по {max_date}.xlsx'


def table_df_create(df):
    df_sort = df.sort_values(by='Всего', ascending=False)
    subset = [i for i in df_sort.columns.tolist() if i != 'Сотрудник']

    df_formatted = df_sort.style.set_properties(
        **{
            "text-align": "center",
            "font-weight": "bold",
            "font-size": "18px",
            "border": "1px solid black",
            'text_wrap': True
        })

    # Apply background gradient to the DataFrame
    df_colors = df_formatted.background_gradient(axis=0,
                                                 subset=subset,
                                                 cmap='YlGn')
    return df_colors


def set_column2(worksheet, cell_format=None, max_col=0):
    worksheet.set_column(0, 1, 30, cell_format)
    worksheet.set_column(1, max_col, 20, cell_format)


def union_df(files, self=None):
    try:
        union_df = pd.DataFrame()
        for file in files:
            print(file)
            try:
                temp_df = pd.read_excel(file,
                                        usecols=['Код номенклатуры', 'Количество факт', 'Код пользователя']).fillna('')
                selected_columns = ['Код номенклатуры', 'Количество факт', 'Код пользователя']
                temp_df[selected_columns] = temp_df[selected_columns].astype(str)
                union_df = pd.concat([union_df, temp_df], ignore_index=True)
            except Exception as ex:
                print(ex)
                QMessageBox.warning(self, 'Ошибка при чтение Excel', f'Ошибка при чтение Excel {ex}')

        # Теперь функция df_in_xlsx должна быть вызвана за пределами цикла, после объединения всех DataFrame
        union_df = union_df[union_df['Код пользователя'] != '']
        union_df['Код пользователя'] = union_df['Код пользователя'].apply(lambda x: int(x.split('.')[0]))
        selected_columns = ['Код номенклатуры', 'Количество факт', 'Код пользователя']
        union_df[selected_columns] = union_df[selected_columns].astype(int)
        df_in_xlsx(union_df, 'Объедененные файлы перемещений и локалок')
        # Группировка по столбцу "Код пользователя" и вывод результатов
        grouped_df = union_df.groupby('Код пользователя').agg({'Количество факт': ['count', 'sum']})

        # Переименовываем столбцы, чтобы получить понятные названия
        grouped_df.columns = ['Количество строк', 'Сумма по "Количество факт"']

        # Выводим результаты
        print(grouped_df)

        # Далее, вы можете сохранить grouped_df в файл Excel, если это необходимо
        df_in_xlsx(grouped_df.reset_index(), 'Общая статистика по перемещениям и локалкам')
    except Exception as ex:
        print(ex)
        QMessageBox.warning(self, 'Ошибка', f'Произошла ошибка: {ex}')
    open_file('Общая статистика по перемещениям и локалкам.xlsx')


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
