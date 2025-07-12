import sys
import json
import pandas as pd
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
from PyQt5.QtGui import QTextDocument


from PyQt5.QtWidgets import (
    QApplication, 
    QWidget,
    QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem,
    QLabel, QLineEdit, QPushButton,
    QColorDialog, QFileDialog,
    QTextEdit, QDialog
)
from PyQt5.QtGui import QColor

class Main(QWidget):
    def __init__(self):
        super().__init__() #так как мы наследуемся, вызываем главный конструктор от QWidget

        self.setWindowTitle("Программа выполнения")
        self.resize(800, 600)

        self.tasks = ["Abbruch", "Trockenbau", "Heizung", "Elektro", "Malerputz", "Fliesen", "Estrich", "Bodenbelag", "Schreiner", 
        "Brandschutzmassnamen", "Endreinigung"]
        self.days = list(range(1,31))
        self.input = {}
        self.colors = {}

        layout = QVBoxLayout()

        layout.addWidget(QLabel("Введите значение от 1 до 30"))


        self.setLayout(layout) # добавляет все созданные обьекты с своиствами в вертикальном расположении и показывает в окне


        for index, task in enumerate(self.tasks):  # переменная task будет хранить значения из списка self.tasks. 
            input_line = QLineEdit() 
            input_line.setPlaceholderText(task) # точка после input_line - так как переменная input_line хранит в себе обьект QLineEdit(),
            #поэтому и ставится точка, и input_line следует рассматривать как обьект. 
            self.input[task] = input_line # Положи в словарь self.input по ключу task значение input_line
            """ Ты создал набор полей ввода, каждое из которых: - имеет своё имя (как ключ в словаре), может быть потом найдено по этому имени,
            и использовано: читать текст, вставить текст, заблокировать, изменить цвет и т.п. """

            self.colors[task] = QColor("red")
            color_button = QPushButton("Color")
            name_button = QPushButton("Name")

            name_button.clicked.connect(lambda _, i = index: self.choose_name(i))
            color_button.clicked.connect(lambda _, t = task: self.choose_color(t))
            """ Это лямбда-функция (анонимная функция, без имени). Создай такую функцию, которая ничего не делает с аргументами,
            но когда её вызовут — она вызовет choose_color(...) с нужной задачей task 
            lambda _:  "есть аргумент, но я его не использую".
            lambda _, t=task: self.choose_color(t) -  сохраняет текущее значение task как значение переменной t внутри каждой лямбды.
            То есть каждая кнопка "запоминает", к какой задаче она относится.
             """ 
            
            
            row_layout = QHBoxLayout()
            row_layout.addWidget(input_line)
            row_layout.addWidget(color_button)
            row_layout.addWidget(name_button)
            
            

            layout.addLayout(row_layout) # добавляем в вертикальный layout горизонтальный layout

        # layout.addWidget(input_line) - так нельзя ставить, так как находится вне цикла - иначе покажет только последний, нужно внутри цикла.

        self.table = QTableWidget(len(self.tasks), len(self.days))
        self.table.setHorizontalHeaderLabels(str(day) for day in self.days)
        self.table.setVerticalHeaderLabels(self.tasks)
        layout.addWidget(self.table)

        # PyQt5: создаём кнопки действий

        btn_layout = QHBoxLayout()

        btn_restart = QPushButton("restart")
        btn_restart.clicked.connect(self.clear_table)
        btn_layout.addWidget(btn_restart)

        btn_update = QPushButton("update")
        btn_update.clicked.connect(self.update_table)
        btn_layout.addWidget(btn_update)

        btn_save = QPushButton("save")
        btn_save.clicked.connect(self.safe_table)
        btn_layout.addWidget(btn_save)

        btn_load = QPushButton("load")
        btn_load.clicked.connect(self.load_table)
        btn_layout.addWidget(btn_load)

        btn_export = QPushButton("export to excel")
        btn_export.clicked.connect(self.export_table)
        btn_layout.addWidget(btn_export)
        
        btn_print = QPushButton("print")
        btn_print.clicked.connect(self.print_table)
        btn_layout.addWidget(btn_print)

        layout.addLayout(btn_layout)

        # PyQt5: создаём функции к кнопкам

    def choose_name(self, index):
        old_name = self.tasks[index]
        
        dialog = QDialog()
        dialog.setWindowTitle("Изменить имя")
        dialog.resize(400, 100)

        layout = QVBoxLayout()
        text_edit = QTextEdit()
        text_edit.setPlaceholderText("Введите новое имя")

        ok_button = QPushButton("OK")
        layout.addWidget(text_edit)
        layout.addWidget(ok_button)
        dialog.setLayout(layout)

        def on_ok():
            new_name = text_edit.toPlainText().strip()
            if new_name:
            # Обновляем имя задачи
                self.tasks[index] = new_name
            # Обновляем словарь input
                self.input[new_name] = self.input.pop(old_name)
                self.input[new_name].setPlaceholderText(new_name)

            
                self.table.setVerticalHeaderLabels(self.tasks)  #  Обновляем заголовки таблицы
                dialog.accept()

        ok_button.clicked.connect(on_ok)
        dialog.exec()  # Показываем диалог (модально)

    def choose_color(self, task):
        color = QColorDialog.getColor()
        if color.isValid():
            self.colors[task] = color
    

    def clear_table(self):
        for input_line in self.input.values():     # if self.table == input_line - input line чего, здесь не ясно с каким лайном работать, 
            input_line.clear()                     # for table in self.table: - self.table это виджет, а не массив, то есть не итерабл. 
                                                   
    def update_table(self):
        for row, task in enumerate(self.tasks):
            input_text = self.input[task].text() # из словаря self.inputs по ключу(task) берем текст(к примеру ключ Abbruch, value - текст)
            days_to_colors = set() # можно было создать список вместо множества set, но с set не будет дубликатов

            for part in input_text.split(','): # Без split(',') ты не сможешь обработать ввод по частям.
                part = part.strip() # удаляем пробелы итд.
                if '-' in part:
                    try:
                        start, end = map(int, part.split('-'))
                        days_to_colors.update(range(start, end + 1))
                    except ValueError:
                        continue
                elif part.isdigit():
                    days_to_colors.add(int(part))

            for col, day in enumerate(self.days):
                item = QTableWidgetItem()
                if day in days_to_colors:
                    item.setBackground(self.colors.get(task, QColor("red")))
                self.table.setItem(row, col, item) # закрашенную ячейку вставляем в таблицу

    def safe_table(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить данные", "", "JSON Files(*.json)")
        if file_path:
            data = {}
            for task in self.tasks:
                data[task] = {
                    "text": self.input[task].text(),
                    "color": self.choose_color[task].name()
                }
            with open(file_path, "w", encoding= "utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

    def load_table(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Загрузить данные", "", "JSON Files (*.json)")  # PyQt5
        if file_path:
            with open(file_path, "r", encoding="utf-8") as f:  # Python
                data = json.load(f)  # Python
                for task in self.tasks:
                    if task in data:
                        self.input[task].setText(data[task]["text"])  # PyQt5
                        self.colors[task] = QColor(data[task]["color"])  # PyQt5
            self.update_table()

    def export_table(self):  # Python + pandas: экспорт таблицы в Excel
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как Excel", "", "Excel Files (*.xlsx)")  # PyQt5
        if file_path:
            data = {}
            for row, task in enumerate(self.tasks):
                row_data = []
                for col, day in enumerate(self.days):
                    item = self.table.item(row, col)  # PyQt5
                    if item and item.background().color().isValid():  # PyQt5
                        row_data.append("✓")  # Python
                    else:
                        row_data.append("")  # Python
                data[task] = row_data  # Python
            df = pd.DataFrame(data, index=[str(day) for day in self.days]).T  # pandas
            df.to_excel(file_path, index=True)  # pandas

    def print_table(self):
    # Создаём объект принтера
        printer = QPrinter(QPrinter.HighResolution)

    # Предпросмотр
        preview = QPrintPreviewDialog(printer, self)
        preview.setWindowTitle("Предпросмотр печати")
        preview.resize(800, 600)
        preview.paintRequested.connect(lambda p: self.handle_paint_request(p))
        preview.exec_()
        
    def handle_paint_request(self, printer):
            doc = QTextDocument()
            # th, td, tr в HTML - это теги - строка, заголовок и ячейка
            html = """
            <style>
                table {
                    border-collapse: collapse;
                    width: 100%;
                    font-family: Arial;
                    font-size: 10pt;
                }
                th, td {
                    border: 1px solid #888;
                    padding: 4px;
                    text-align: center;
                }
                th {
                    background-color: #f0f0f0;
                }
            </style>
            """

            html += "<h2 align='center'>Расписание задач</h2>" #html += — прибавляем эту строку к нашему HTML.
            html += "<table>"
            html += "<tr><th>Задача</th>"

            for day in self.days:
                html += f"<th>{day}</th>"
            html += "</tr>"

            for row, task in enumerate(self.tasks):
                html += f"<tr><td>{task}</td>"
                for col in range(len(self.days)):
                    item = self.table.item(row, col)
                    content = "✓" if item and item.background().color().isValid() else ""
                    html += f"<td>{content}</td>"
                html += "</tr>"

            html += "</table>"
            doc.setHtml(html)
            doc.print_(printer)
            

# Запуск приложения
app = QApplication(sys.argv)
Window = Main()
Window.show()
sys.exit(app.exec_())

