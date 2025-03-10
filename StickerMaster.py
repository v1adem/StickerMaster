import sys
import os
import fitz  # PyMuPDF
import re
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, \
    QMessageBox, QGroupBox, QGridLayout, QTabWidget, QComboBox, QCheckBox
from PySide6.QtGui import QPixmap, Qt
from PIL import ImageQt, Image


class StickerGeneratorApp(QWidget):
    def __init__(self):

        self.font_mapping = {
            "MyriadPro-Regular": "fonts/MyriadPro-Regular.ttf",
            "MyriadPro-Bold": "fonts/MyriadPro-Bold.ttf",
            "Arial-BoldMT": "fonts/arial-mt-bold.ttf",
            "YVUNJSå¼«Arial-BoldMT": "fonts/4068-font.ttf",  #
            "ArialMT": "fonts/ArialMT-Light.ttf",
            "Bahnschrift": "fonts/bahnschrift.ttf",
            "STGONDå¼«Bahnschrift": "fonts/bahnschrift.ttf",
            "JIZLJBå¼«Bahnschrift": "fonts/bahnschrift.ttf",
            "RHTNZFå¼«Bahnschrift": "fonts/bahnschrift.ttf",
            "WHCMRVå¼«Bahnschrift": "fonts/bahnschrift.ttf",
            "IMEGDXå¼«Bahnschrift": "fonts/bahnschrift.ttf",
            "MyriadPro-SemiboldCond": "fonts/MyriadPro-SemiboldCond.otf",
            "MyriadPro-BlackSemiExt": "fonts/Myriad Pro Black SemiExtended.otf",
            "MyanmarText-Bold": "fonts/Myanmar Text Bold.TTF",
            "MinionPro-Bold": "fonts/Minion Pro Bold.ttf",
            "RJWHAAå¼«MyriadPro-Semib": "fonts/MyriadPro-SemiboldCond.otf",
            "JFXDXWå¼«MyriadPro-Bold": "fonts/MyriadPro-Bold.ttf",
            "LCJXNXå¼«MinionPro-Bold": "fonts/Minion Pro Bold.ttf",
            "ZPRMHKå¼«MyriadPro-Regul": "fonts/MyriadPro-Regular.ttf",
        }

        super().__init__()
        self.templates_dir = "templates"  # Папка з шаблонами
        self.load_templates()
        self.initUI()
        self.template_pixmap = None
        self.preview_pixmap = None

    def load_templates(self):
        self.standard_templates = {}
        self.box_templates = {}

        try:
            for filename in os.listdir(self.templates_dir):
                if filename.endswith(".pdf"):
                    template_name = filename[:-4]  # Видаляємо розширення .pdf
                    file_path = os.path.join(self.templates_dir, filename)
                    if template_name.startswith("standard_"):  # Розрізняємо за префіксом
                        self.standard_templates[template_name[9:]] = file_path  # Видаляємо префікс "standard_"
                    elif template_name.startswith("box_"):
                        self.box_templates[template_name[4:]] = file_path  # Видаляємо префікс "box_"
        except FileNotFoundError:
            QMessageBox.critical(self, "Помилка", f"Папка з шаблонами '{self.templates_dir}' не знайдена!")
            return

    def initUI(self):
        main_layout = QGridLayout()  # Головний layout - QGridLayout

        # Створюємо TabWidget
        self.tabs = QTabWidget()

        self.create_IME_standard_form_elements()
        self.create_IME_box_form_elements()

        # Створюємо вкладку "IME Standart"
        self.standard_tab = QWidget()
        self.create_IME_standard_tab()  # Функція для створення вмісту вкладки
        self.tabs.addTab(self.standard_tab, "IME Standard")

        # Створюємо вкладку "IME Коробки"
        self.boxes_tab = QWidget()
        self.create_IME_box_tab()  # Функція для створення вмісту вкладки
        self.tabs.addTab(self.boxes_tab, "IME Box")

        main_layout.addWidget(self.tabs, 0, 0)  # row, col
        main_layout.setAlignment(self.tabs, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        # Створюємо області попереднього перегляду
        self.template_preview_label = QLabel()
        self.preview_preview_label = QLabel()

        # Додаємо області попереднього перегляду до головного layout
        main_layout.addWidget(self.template_preview_label, 0, 1)  # row, col
        main_layout.addWidget(self.preview_preview_label, 1, 1)  # row, col

        # Вирівнюємо області попереднього перегляду по верхньому лівому куту
        main_layout.setAlignment(self.template_preview_label, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        main_layout.setAlignment(self.preview_preview_label, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        self.setLayout(main_layout)
        self.setWindowTitle("Sticker Master IME")
        self.template_path = ""
        self.temp_preview_path = ""

    def create_IME_standard_form_elements(self):
        self.IME_standard_template_label = QLabel("PDF-шаблон:")
        self.IME_standard_template_combo = QComboBox()
        self.IME_standard_template_combo.addItem("- Шаблон не обрано -")  # Пункт за замовчуванням
        for template_name in sorted(self.standard_templates.keys()):  # Сортуємо за алфавітом
            self.IME_standard_template_combo.addItem(template_name)
        self.IME_standard_template_combo.currentIndexChanged.connect(
            self.select_standard_template)  # Connect to new function

        self.prefix_IME_standard_label = QLabel("Перші 6 цифр серійного номеру:")
        self.prefix_IME_standard_input = QLineEdit()

        self.short_prefix_IME_standard_label = QLabel("Серія:")
        self.short_prefix_IME_standard_input = QLineEdit()

        self.art_seria_IME_standard_label = QLabel("Серія в артикулі\nДо номіналу:")
        self.art_seria_IME_standard_input = QLineEdit()

        self.nominal_IME_standard_label = QLabel("Номінал:")
        self.nominal_IME_standard_input = QLineEdit()

        self.va_IME_standard_label = QLabel("VA:")
        self.va_IME_standard_input = QLineEdit()

        self.count_IME_standard_label = QLabel("Кількість:")
        self.count_IME_standard_input = QLineEdit()

        self.year_IME_standard_label = QLabel("Рік:")
        self.year_IME_standard_input = QLineEdit()

        self.week_IME_standard_label = QLabel("Тиждень:")
        self.week_IME_standard_input = QLineEdit()

        self.va_cl_02s_IME_standard_label = QLabel("VA (CL 0.2S):")
        self.va_cl_02_IME_standard_label = QLabel("VA (CL 0.2):")
        self.va_cl_05s_IME_standard_label = QLabel("VA (CL 0.5S):")

        # Додаємо нові поля VA
        self.va_cl_02s_IME_standard_input = QLineEdit()
        self.va_cl_02_IME_standard_input = QLineEdit()
        self.va_cl_05s_IME_standard_input = QLineEdit()

        self.add_3_checkbox = QCheckBox("Додати 3 в кінці артикула")

        self.generate_1_standard_sticker_button = QPushButton("Згенерувати 1 стікер")
        self.generate_1_standard_sticker_button.clicked.connect(self.generate_one_IME_standard_pdfs)
        self.generate_1_standard_sticker_input = QLineEdit()
        self.generate_1_standard_sticker_input.setPlaceholderText("Всі 4 цифри номеру")

        # Встановлюємо рамку для QLineEdit
        self.art_seria_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.prefix_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_cl_02s_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_cl_02_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_cl_05s_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.generate_1_standard_sticker_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        self.generate_IME_standard_button = QPushButton("Згенерувати PDF")
        self.generate_IME_standard_button.clicked.connect(self.generate_IME_standard_pdfs)

        self.preview_IME_standard_button = QPushButton("Попередній перегляд")
        self.preview_IME_standard_button.clicked.connect(self.show_IME_standard_preview)

    def create_IME_box_form_elements(self):
        self.IME_box_template_label = QLabel("PDF-шаблон:")
        self.IME_box_template_combo = QComboBox()
        self.IME_box_template_combo.addItem("- Шаблон не обрано -")  # Пункт за замовчуванням
        for template_name in sorted(self.box_templates.keys()):  # Сортуємо за алфавітом
            self.IME_box_template_combo.addItem(template_name)
        self.IME_box_template_combo.currentIndexChanged.connect(self.select_box_template)  # Connect to new function

        self.seria_IME_box_label = QLabel("Серія:")
        self.seria_IME_box_input = QLineEdit()

        self.art_seria_IME_box_label = QLabel("Серія в артикулі\nДо номіналу:")
        self.art_seria_IME_box_input = QLineEdit()

        self.nominal_IME_box_label = QLabel("Номінал:")
        self.nominal_IME_box_input = QLineEdit()

        self.box_count_IME_box_label = QLabel("Кількість:")
        self.box_count_IME_box_input = QLineEdit()

        self.year_IME_box_label = QLabel("Рік:")
        self.year_IME_box_input = QLineEdit()

        self.week_IME_box_label = QLabel("Тиждень:")
        self.week_IME_box_input = QLineEdit()

        self.add_3_checkbox_box = QCheckBox("Додати 3 в кінці артикула")

        # Встановлюємо рамку для QLineEdit
        self.art_seria_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.seria_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.box_count_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        self.generate_IME_box_button = QPushButton("Згенерувати PDF")
        self.generate_IME_box_button.clicked.connect(self.generate_IME_box_pdfs)

        self.preview_IME_box_button = QPushButton("Попередній перегляд")
        self.preview_IME_box_button.clicked.connect(self.show_IME_box_preview)

    def create_IME_standard_tab(self):
        # Група для параметрів генерації
        input_group = QGroupBox("Параметри генерації")  # Створюємо групу
        input_layout = QGridLayout()

        input_layout.setHorizontalSpacing(5)  # Зменшуємо горизонтальний відступ
        input_layout.setVerticalSpacing(5)  # Зменшуємо вертикальний відступ

        # Встановлюємо рамку для групи
        input_group.setStyleSheet("QGroupBox { border: 1px solid gray; }")

        # Розміщуємо віджети в QGridLayout
        input_layout.addWidget(self.IME_standard_template_label, 0, 0)  # row, col
        input_layout.addWidget(self.IME_standard_template_combo, 0, 1)

        input_layout.addWidget(self.prefix_IME_standard_label, 1, 0)
        input_layout.addWidget(self.prefix_IME_standard_input, 1, 1)

        input_layout.addWidget(self.short_prefix_IME_standard_label, 2, 0)
        input_layout.addWidget(self.short_prefix_IME_standard_input, 2, 1)

        input_layout.addWidget(self.art_seria_IME_standard_label, 3, 0)
        input_layout.addWidget(self.art_seria_IME_standard_input, 3, 1)

        input_layout.addWidget(self.nominal_IME_standard_label, 4, 0)
        input_layout.addWidget(self.nominal_IME_standard_input, 4, 1)

        input_layout.addWidget(self.va_IME_standard_label, 5, 0)
        input_layout.addWidget(self.va_IME_standard_input, 5, 1)

        input_layout.addWidget(self.count_IME_standard_label, 6, 0)
        input_layout.addWidget(self.count_IME_standard_input, 6, 1)

        input_layout.addWidget(self.year_IME_standard_label, 7, 0)
        input_layout.addWidget(self.year_IME_standard_input, 7, 1)

        input_layout.addWidget(self.week_IME_standard_label, 8, 0)
        input_layout.addWidget(self.week_IME_standard_input, 8, 1)

        # Встановлюємо рамку для QLineEdit
        self.prefix_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.short_prefix_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        input_layout.addWidget(self.generate_IME_standard_button, 9, 0, 1, 2)
        input_layout.addWidget(self.preview_IME_standard_button, 10, 0, 1, 2)

        input_layout.addWidget(self.va_cl_02s_IME_standard_label, 11, 0)
        input_layout.addWidget(self.va_cl_02s_IME_standard_input, 11, 1)
        input_layout.addWidget(self.va_cl_02_IME_standard_label, 12, 0)
        input_layout.addWidget(self.va_cl_02_IME_standard_input, 12, 1)
        input_layout.addWidget(self.va_cl_05s_IME_standard_label, 13, 0)
        input_layout.addWidget(self.va_cl_05s_IME_standard_input, 13, 1)

        input_layout.addWidget(self.add_3_checkbox, 14, 0, 1, 2)

        input_layout.addWidget(self.generate_1_standard_sticker_input, 15, 0)
        input_layout.addWidget(self.generate_1_standard_sticker_button, 15, 1)

        self.va_cl_02s_IME_standard_label.setVisible(False)
        self.va_cl_02s_IME_standard_input.setVisible(False)
        self.va_cl_02_IME_standard_label.setVisible(False)
        self.va_cl_02_IME_standard_input.setVisible(False)
        self.va_cl_05s_IME_standard_label.setVisible(False)
        self.va_cl_05s_IME_standard_input.setVisible(False)
        self.add_3_checkbox.setVisible(False)

        self.standard_tab.setLayout(input_layout)  # Встановлюємо layout для вкладки

    def create_IME_box_tab(self):
        # Група для параметрів генерації
        input_group = QGroupBox("Параметри генерації")  # Створюємо групу
        input_layout = QGridLayout()

        input_layout.setHorizontalSpacing(5)  # Зменшуємо горизонтальний відступ
        input_layout.setVerticalSpacing(5)  # Зменшуємо вертикальний відступ

        # Встановлюємо рамку для групи
        input_group.setStyleSheet("QGroupBox { border: 1px solid gray; }")

        input_layout.addWidget(self.IME_box_template_label, 0, 0)  # row, col
        input_layout.addWidget(self.IME_box_template_combo, 0, 1)

        input_layout.addWidget(self.seria_IME_box_label, 1, 0)
        input_layout.addWidget(self.seria_IME_box_input, 1, 1)

        input_layout.addWidget(self.art_seria_IME_box_label, 2, 0)
        input_layout.addWidget(self.art_seria_IME_box_input, 2, 1)

        input_layout.addWidget(self.nominal_IME_box_label, 3, 0)
        input_layout.addWidget(self.nominal_IME_box_input, 3, 1)

        input_layout.addWidget(self.box_count_IME_box_label, 4, 0)
        input_layout.addWidget(self.box_count_IME_box_input, 4, 1)

        input_layout.addWidget(self.year_IME_box_label, 5, 0)
        input_layout.addWidget(self.year_IME_box_input, 5, 1)

        input_layout.addWidget(self.week_IME_box_label, 6, 0)
        input_layout.addWidget(self.week_IME_box_input, 6, 1)

        input_layout.addWidget(self.generate_IME_box_button, 7, 0, 1, 2)  # span columns
        input_layout.addWidget(self.preview_IME_box_button, 8, 0, 1, 2)  # span columns

        input_layout.addWidget(self.add_3_checkbox_box, 9, 0, 1, 2)
        self.add_3_checkbox_box.setVisible(False)

        self.boxes_tab.setLayout(input_layout)  # Встановлюємо layout для вкладки

    def select_standard_template(self):
        selected_template = self.IME_standard_template_combo.currentText()
        if selected_template != "- Шаблон не обрано -":
            self.template_path = self.standard_templates[selected_template]
            self.display_template_preview(self.template_path)

            # Відображаємо або приховуємо поля VA залежно від шаблону
            is_special_template = selected_template.endswith("_special_1")
            self.va_IME_standard_label.setVisible(not is_special_template)
            self.va_IME_standard_input.setVisible(not is_special_template)
            self.va_cl_02s_IME_standard_label.setVisible(is_special_template)
            self.va_cl_02s_IME_standard_input.setVisible(is_special_template)
            self.va_cl_02_IME_standard_label.setVisible(is_special_template)
            self.va_cl_02_IME_standard_input.setVisible(is_special_template)
            self.va_cl_05s_IME_standard_label.setVisible(is_special_template)
            self.va_cl_05s_IME_standard_input.setVisible(is_special_template)
            self.add_3_checkbox.setVisible(is_special_template)
        else:
            self.template_path = ""
            self.template_pixmap = None
            self.template_preview_label.clear()

            # Приховуємо поля VA, якщо шаблон не обрано
            self.va_IME_standard_label.setVisible(True)
            self.va_IME_standard_input.setVisible(True)
            self.va_cl_02s_IME_standard_label.setVisible(False)
            self.va_cl_02s_IME_standard_input.setVisible(False)
            self.va_cl_02_IME_standard_label.setVisible(False)
            self.va_cl_02_IME_standard_input.setVisible(False)
            self.va_cl_05s_IME_standard_label.setVisible(False)
            self.va_cl_05s_IME_standard_input.setVisible(False)

    def select_box_template(self):
        selected_template = self.IME_box_template_combo.currentText()
        if selected_template != "- Шаблон не обрано -":
            self.template_path = self.box_templates[selected_template]
            self.display_template_preview(self.template_path)
            # Додаємо визначення чи шаблон спецільний
            self.is_box_special_template = selected_template.endswith("_box_special_1")
            self.add_3_checkbox_box.setVisible(self.is_box_special_template)
        else:
            self.template_path = ""
            self.template_pixmap = None
            self.template_preview_label.clear()
            self.is_box_special_template = False

    def display_template_preview(self, file_path):
        try:
            pixmap = self.pdf_to_pixmap(file_path)
            if pixmap:
                self.template_pixmap = pixmap
                self.template_preview_label.setPixmap(pixmap)  # Без масштабування
                self.template_preview_label.resize(pixmap.size())  # Автоматичний розмір
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Помилка завантаження прев'ю шаблону: {e}")

    def pdf_to_pixmap(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]

            # Змінюємо матрицю для збільшення розміру
            matrix = fitz.Matrix(3, 3)  # Збільшення в 4 рази (400%)
            pix = page.get_pixmap(matrix=matrix)

            # Конвертуємо за допомогою Pillow (якщо потрібно)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            qimage = ImageQt.ImageQt(img)
            pixmap = QPixmap.fromImage(qimage)

            doc.close()
            return pixmap
        except Exception as e:
            print(f"Error converting PDF to pixmap: {e}")
            return None

    def show_IME_standard_preview(self):
        if not self.template_path:
            return
        short_prefix = self.short_prefix_IME_standard_input.text()
        prefix = self.prefix_IME_standard_input.text()
        year = self.year_IME_standard_input.text()
        week = self.week_IME_standard_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_standard_input.text()
        va = self.va_IME_standard_input.text()

        self.temp_preview_path = os.path.join(os.getcwd(), "tmp_preview.pdf")
        doc_output = fitz.open()
        if "special_1" in self.template_path:
            va_cl_02s = self.va_cl_02s_IME_standard_input.text()
            va_cl_02 = self.va_cl_02_IME_standard_input.text()
            va_cl_05s = self.va_cl_05s_IME_standard_input.text()
            self.modify_IME_standart_pdf(self.template_path, doc_output, f"{prefix}0001", date_code, nominal,
                                         va_cl_02s, short_prefix, va_cl_02, va_cl_05s)
        else:
            self.modify_IME_standart_pdf(self.template_path, doc_output, f"{prefix}0001", date_code, nominal, va,
                                         short_prefix)
        doc_output.save(self.temp_preview_path)
        doc_output.close()

        # Перевірка існування файлу
        if os.path.exists(self.temp_preview_path):
            print(f"Тимчасовий файл створено: {self.temp_preview_path}")  # Для дебагу
            try:
                pixmap = self.pdf_to_pixmap(self.temp_preview_path)
                if pixmap:
                    self.preview_pixmap = pixmap
                    self.preview_preview_label.setPixmap(pixmap)  # Без масштабування
                    self.preview_preview_label.resize(pixmap.size())  # Автоматичний розмір
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Помилка генерації прев'ю: {e}")
        else:
            print("Помилка: Тимчасовий файл не створено!")  # Для дебагу
            QMessageBox.critical(self, "Помилка", "Помилка створення тимчасового файлу прев'ю.")

    def show_IME_box_preview(self):
        if not self.template_path:
            return
        short_prefix = self.seria_IME_box_input.text()
        year = self.year_IME_box_input.text()
        week = self.week_IME_box_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_box_input.text()

        self.temp_preview_path = os.path.join(os.getcwd(), "tmp_preview.pdf")
        doc_output = fitz.open()
        # Додаємо перевірку на спецільний шаблон
        if self.is_box_special_template:
            self.modify_IME_special_box_pdf(self.template_path, doc_output, nominal, short_prefix)
        else:
            self.modify_IME_box_pdf(self.template_path, doc_output, date_code, nominal, short_prefix)
        doc_output.save(self.temp_preview_path)
        doc_output.close()

        # Перевірка існування файлу
        if os.path.exists(self.temp_preview_path):
            print(f"Тимчасовий файл створено: {self.temp_preview_path}")  # Для дебагу
            try:
                pixmap = self.pdf_to_pixmap(self.temp_preview_path)
                if pixmap:
                    self.preview_pixmap = pixmap
                    self.preview_preview_label.setPixmap(pixmap)  # Без масштабування
                    self.preview_preview_label.resize(pixmap.size())  # Автоматичний розмір
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Помилка генерації прев'ю: {e}")
        else:
            print("Помилка: Тимчасовий файл не створено!")  # Для дебагу
            QMessageBox.critical(self, "Помилка", "Помилка створення тимчасового файлу прев'ю.")

    def generate_IME_standard_pdfs(self):
        if not self.template_path:
            return

        prefix = self.prefix_IME_standard_input.text()
        short_prefix = self.short_prefix_IME_standard_input.text()
        count = int(self.count_IME_standard_input.text())
        year = self.year_IME_standard_input.text()
        week = self.week_IME_standard_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_standard_input.text()
        va = self.va_IME_standard_input.text()

        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження")
        if not folder:
            return

        # Формуємо назву файлу
        output_pdf = os.path.join(folder, f"{short_prefix} {nominal}A {date_code}-{count}шт.pdf")
        doc_output = fitz.open()  # Новий PDF документ для збереження всіх сторінок

        if "special_1" in self.template_path:
            va_cl_02s = self.va_cl_02s_IME_standard_input.text()
            va_cl_02 = self.va_cl_02_IME_standard_input.text()
            va_cl_05s = self.va_cl_05s_IME_standard_input.text()
            for i in range(1, count + 1):
                serial_number = f"{prefix}{i:04}"
                self.modify_IME_standart_pdf(self.template_path, doc_output, serial_number, date_code, nominal,
                                             va_cl_02s, short_prefix, va_cl_02, va_cl_05s)
        else:
            for i in range(1, count + 1):
                serial_number = f"{prefix}{i:04}"
                self.modify_IME_standart_pdf(self.template_path, doc_output, serial_number, date_code, nominal, va,
                                             short_prefix)

        # Зберегти фінальний PDF файл з усіма сторінками
        doc_output.save(output_pdf)
        doc_output.close()

    def generate_one_IME_standard_pdfs(self):
        if not self.template_path:
            return

        prefix = self.prefix_IME_standard_input.text()
        short_prefix = self.short_prefix_IME_standard_input.text()
        year = self.year_IME_standard_input.text()
        week = self.week_IME_standard_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_standard_input.text()
        va = self.va_IME_standard_input.text()
        i = self.generate_1_standard_sticker_input.text()

        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження")
        if not folder:
            return

        # Формуємо назву файлу
        output_pdf = os.path.join(folder, f"{short_prefix} {nominal}A {date_code} #{i}-1шт.pdf")
        doc_output = fitz.open()  # Новий PDF документ для збереження всіх сторінок
        serial_number = f"{prefix}{i:04}"
        if "special_1" in self.template_path:
            va_cl_02s = self.va_cl_02s_IME_standard_input.text()
            va_cl_02 = self.va_cl_02_IME_standard_input.text()
            va_cl_05s = self.va_cl_05s_IME_standard_input.text()
            self.modify_IME_standart_pdf(self.template_path, doc_output, serial_number, date_code, nominal,
                                         va_cl_02s, short_prefix, va_cl_02, va_cl_05s)
        else:
            self.modify_IME_standart_pdf(self.template_path, doc_output, serial_number, date_code, nominal, va,
                                         short_prefix)

        # Зберегти фінальний PDF файл з усіма сторінками
        doc_output.save(output_pdf)
        doc_output.close()

    def generate_IME_box_pdfs(self):
        if not self.template_path:
            return
        seria = self.seria_IME_box_input.text()
        count = int(self.box_count_IME_box_input.text())
        year = self.year_IME_box_input.text()
        week = self.week_IME_box_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_box_input.text()

        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження")
        if not folder:
            return

        # Формуємо назву файлу
        output_pdf = os.path.join(folder, f"{seria} {nominal}A {date_code}-КОРОБКА-{count}шт.pdf")
        doc_output = fitz.open()  # Новий PDF документ для збереження всіх сторінок

        for i in range(1, count + 1):
            # Додаємо перевірку на спецільний шаблон
            if self.is_box_special_template:
                self.modify_IME_special_box_pdf(self.template_path, doc_output, nominal, seria)
            else:
                self.modify_IME_box_pdf(self.template_path, doc_output, date_code, nominal, seria)

        # Зберегти фінальний PDF файл з усіма сторінками
        doc_output.save(output_pdf)
        doc_output.close()

    def modify_IME_standart_pdf(self, input_pdf, doc_output, serial_number, date_code, nominal, va, short_prefix,
                                va_cl_02=None, va_cl_05s=None):
        doc = fitz.open(input_pdf)
        date_pattern = r"\d{2}W\d{2}"  # Шаблон для року і тижня
        serial_pattern = r"^\s*\d{10}\s*$"  # Шаблон для серійного номера
        ipr_pattern = r"Ipr\d+A"  # Шаблон для пошуку "Ipr"
        d_ipr_pattern = r"Ipr $"  # Шаблон для пошуку "Ipr" перед числом
        blank_ipr_pattern = r"Ipr \d+A"  # Шаблон для пошуку "Ipr "
        article_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+[C|B|D]\d+SE$"  # Шаблон для артикулу
        seria_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+B?$"  # Шаблон для серії

        special_article_pattern = r"TASL50C6003S\+1,2/6KV\s{6}TAS65"
        is_special_template = input_pdf.endswith("_special_1.pdf")

        combined_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+[C|B|D]\d+SE\s+(TA|TT|TAS|TASS|TASO)\d+B?"

        for page in doc:
            new_page = doc_output.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, doc, page.number)

            text_blocks = page.get_text("dict")["blocks"]

            for block in text_blocks:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_text = span["text"]  # Поточний текст
                        bbox = span["bbox"]  # Позиція тексту
                        font_size = span["size"]  # Розмір шрифту
                        font_name = span["font"]  # Назва шрифту

                        font_path = self.font_mapping.get(font_name)

                        print(span_text)

                        x0, y0, a, b = bbox
                        y0 += 0.1
                        bbox = x0, y0, a, b

                        # Заміна серійного номера
                        if re.match(serial_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            if serial_number is not None:
                                new_page.insert_text((x0, y0 + font_size), serial_number, fontsize=font_size,
                                                     color=(0, 0, 0),
                                                     fontfile=font_path, fontname=font_name)

                        # Заміна дати
                        if re.match(date_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), date_code, fontsize=font_size, color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        # Заміна номіналу в тексті після "Ipr" та перед "A"
                        if re.match(ipr_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), f"Ipr {nominal}A", fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        # Заміна номіналу в тексті після "Ipr" та перед "A"
                        if re.match(d_ipr_pattern, span_text):
                            bbox = x0, y0, a + 15, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), f"Ipr {nominal}A", fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        if re.match(blank_ipr_pattern, span_text):
                            bbox = x0, y0, a + 15, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), f"Ipr {nominal}A", fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        # Заміна артикулу на кожній сторінці
                        if re.fullmatch(article_pattern, span_text):
                            local_nominal = nominal
                            letter = "B"
                            try:
                                nominal_int = int(nominal)  # Перетворюємо nominal на ціле число
                                if 100 <= nominal_int <= 999:  # Перевіряємо, чи nominal 3-значне
                                    letter = "C"
                                elif 1000 <= nominal_int <= 9999:  # Перевіряємо, чи nominal 4-значне
                                    letter = "D"
                                    local_nominal = nominal_int + 3  # Обчислюємо local_nominal
                                    x0 -= 2
                                else:
                                    # Обробка ситуації, коли nominal не 3-значне і не 4-значне
                                    print("nominal має бути 3- або 4-значним числом.")
                            except ValueError:
                                # Обробка помилки, якщо nominal не можна перетворити на ціле число
                                print("Помилка: nominal має бути цілим числом.")

                            new_text = f"{self.art_seria_IME_standard_input.text()}{letter}{local_nominal}SE"
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            if len(span_text) < len(new_text):
                                x0 = x0 - ((len(new_text) - len(span_text)) * 3)
                            elif len(span_text) > len(new_text):
                                x0 = x0 + ((len(span_text) - len(new_text)) * 3)
                            new_page.insert_text((x0, y0 + font_size), new_text,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if re.match(seria_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            if len(span_text) < len(short_prefix):
                                x0 = x0 - ((len(short_prefix) - len(span_text)) * 3)
                            elif len(span_text) > len(short_prefix):
                                x0 = x0 + ((len(span_text) - len(short_prefix)) * 3)
                            new_page.insert_text((x0, y0 + font_size), short_prefix,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        combined_match = re.search(combined_pattern, span_text)
                        if combined_match:
                            local_nominal = nominal
                            letter = "B"
                            try:
                                nominal_int = int(nominal)
                                if 100 <= nominal_int <= 999:
                                    letter = "C"
                                elif 1000 <= nominal_int <= 9999:
                                    letter = "D"
                                    local_nominal = nominal_int + 3
                                else:
                                    print("nominal має бути 3- або 4-значним числом.")
                            except ValueError:
                                print("Помилка: nominal має бути цілим числом.")

                            new_text = f"{self.art_seria_IME_standard_input.text()}{letter}{local_nominal}SE     {short_prefix}"

                            new_span_text = span_text.replace(span_text[combined_match.start():combined_match.end()],
                                                              new_text)

                            x0, y0, a, b = bbox
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), new_span_text,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if is_special_template:
                            if re.fullmatch(special_article_pattern, span_text):
                                local_nominal = nominal
                                letter = "B"
                                try:
                                    nominal_int = int(nominal)
                                    if 100 <= nominal_int <= 999:
                                        letter = "C"
                                    elif 1000 <= nominal_int <= 9999:
                                        letter = "D"
                                        local_nominal = nominal_int + 3
                                    else:
                                        print("nominal має бути 3- або 4-значним числом.")
                                except ValueError:
                                    print("Помилка: nominal має бути цілим числом.")

                                three = "3" if self.add_3_checkbox.isChecked() else ""

                                new_text = f"{self.art_seria_IME_standard_input.text()}{letter}{local_nominal}{three}S+1,2/6KV      {short_prefix}"
                                x0, y0, a, b = bbox
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size), new_text,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                            if re.match(r"^1$", span_text):
                                bbox = x0, y0 + 1, a, b
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size - 0.1), va,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                            if re.match(r"^3$", span_text):
                                bbox = x0, y0 + 1, a, b
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size - 0.1), va_cl_02,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                            if re.match(r"^5$", span_text):
                                bbox = x0, y0 + 1, a, b
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size - 0.1), va_cl_05s,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                        else:
                            if re.match(r"^(15|10)$", span_text):
                                bbox = x0, y0 + 1, a, b
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size - 0.1), va,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                            if re.match(r"^(5|3)$", span_text):
                                bbox = x0, y0 + 1, a, b
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                va_num = int(va)
                                minus = 0 if va_num < 9 else 2
                                new_page.insert_text((x0 - minus, y0 + font_size - 0.1), va,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

        doc.close()

    def modify_IME_box_pdf(self, input_pdf, doc_output, date_code, nominal, seria):
        doc = fitz.open(input_pdf)
        date_pattern = r"\d{2}W\d{2}$"  # Шаблон для року і тижня
        article_pattern = r"(TA|TT|TAS|TASS|TASO)\d+[C|B|D]\d+SE"  # Шаблон для артикулу
        start_pattern = r"\s*(TA|TT|TAS|TASS|TASO)\d+B?$"
        date_article_combined_pattern = r"(\d{2}W\d{2})\s+(TAS\d+[A-Z]?[0-9]*)"

        for page in doc:
            new_page = doc_output.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, doc, page.number)

            text_blocks = page.get_text("dict")["blocks"]

            for block in text_blocks:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_text = span["text"]  # Поточний текст
                        bbox = span["bbox"]  # Позиція тексту
                        font_size = span["size"]  # Розмір шрифту
                        font_name = span["font"]  # Назва шрифту

                        font_path = self.font_mapping.get(font_name)

                        print(span_text)

                        # Заміна дати
                        if re.match(date_pattern, span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0 - 1, y0 + font_size), date_code, fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        nominal_match = re.search(r"(\d{3,4})/5A", span_text)
                        if nominal_match:
                            extracted_nominal = nominal_match.group(1)
                            try:
                                extracted_nominal_int = int(extracted_nominal)
                                if 100 <= extracted_nominal_int <= 9999:
                                    # Замінюємо текст у span_text
                                    new_span_text = span_text.replace(extracted_nominal, nominal)
                                    x0, y0, a, b = bbox
                                    bbox = x0, y0, a, b
                                    new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                    new_page.apply_redactions()

                                    new_page.insert_text((x0, y0 + font_size), new_span_text,
                                                         # Використовуємо new_span_text
                                                         fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                         fontname=font_name)
                            except ValueError:
                                print(f"Помилка: '{extracted_nominal}' не є дійсним числом.")

                        # Заміна артикулу на кожній сторінці
                        if re.fullmatch(article_pattern, span_text):
                            local_nominal = nominal
                            letter = "B"
                            try:
                                nominal_int = int(nominal)  # Перетворюємо nominal на ціле число
                                if 100 <= nominal_int <= 999:  # Перевіряємо, чи nominal 3-значне
                                    letter = "C"
                                elif 1000 <= nominal_int <= 9999:  # Перевіряємо, чи nominal 4-значне
                                    letter = "D"
                                    local_nominal = nominal_int + 3  # Обчислюємо local_nominal
                                    x0 -= 2
                                else:
                                    # Обробка ситуації, коли nominal не 3-значне і не 4-значне
                                    print("nominal має бути 3- або 4-значним числом.")
                            except ValueError:
                                # Обробка помилки, якщо nominal не можна перетворити на ціле число
                                print("Помилка: nominal має бути цілим числом.")

                            new_text = f"{self.art_seria_IME_box_input.text()}{letter}{local_nominal}SE"
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, a, b = bbox
                            new_page.insert_text((x0, y0 + font_size), new_text,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if re.match(start_pattern, span_text):
                            x0, y0, a, b = bbox
                            x0 += 1
                            y0 += 1
                            a -= 1
                            b -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0 - 3, y0 + font_size), seria,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        date_article_match = re.search(date_article_combined_pattern, span_text)
                        if date_article_match:
                            combined_x0, combined_y0, combined_x1, combined_y1 = bbox

                            extracted_date = date_article_match.group(1)
                            extracted_seria = date_article_match.group(2)

                            date_width = len(extracted_date) * 3  # Приблизна ширина дати
                            seria_width = len(extracted_seria) * 3  # Приблизна ширина розділювача

                            # Замінюємо дату
                            new_page.add_redact_annot(
                                (combined_x0, combined_y0, combined_x0 + date_width - 0.2, combined_y1 - 0.3),
                                fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((combined_x0 - 4, combined_y0 + font_size), date_code,
                                                 fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                            # Замінюємо серію
                            seria_x0 = combined_x0 + date_width + 2
                            new_page.add_redact_annot(
                                (seria_x0, combined_y0, seria_x0 + seria_width, combined_y1 - 0.3),
                                fill=[255, 255, 255])  # Приблизна ширина для серії
                            new_page.apply_redactions()
                            new_page.insert_text((seria_x0 + 1, combined_y0 + font_size), seria,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

        doc.close()

    def modify_IME_special_box_pdf(self, input_pdf, doc_output, nominal, seria):
        doc = fitz.open(input_pdf)
        year = self.year_IME_box_input.text()
        week = self.week_IME_box_input.text()

        for page in doc:
            new_page = doc_output.new_page(width=page.rect.width, height=page.rect.height)
            new_page.show_pdf_page(new_page.rect, doc, page.number)

            text_blocks = page.get_text("dict")["blocks"]

            for block in text_blocks:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_text = span["text"]  # Поточний текст
                        bbox = span["bbox"]  # Позиція тексту
                        font_size = span["size"]  # Розмір шрифту
                        font_name = span["font"]  # Назва шрифту

                        font_path = self.font_mapping.get(font_name)

                        print(span_text)

                        if re.match(r"TASL50C6003S", span_text):
                            x0, y0, a, b = bbox
                            local_nominal = nominal
                            letter = "B"
                            try:
                                nominal_int = int(nominal)  # Перетворюємо nominal на ціле число
                                if 100 <= nominal_int <= 999:  # Перевіряємо, чи nominal 3-значне
                                    letter = "C"
                                elif 1000 <= nominal_int <= 9999:  # Перевіряємо, чи nominal 4-значне
                                    letter = "D"
                                    local_nominal = nominal_int + 3  # Обчислюємо local_nominal
                                    x0 -= 2
                                else:
                                    # Обробка ситуації, коли nominal не 3-значне і не 4-значне
                                    print("nominal має бути 3- або 4-значним числом.")
                            except ValueError:
                                # Обробка помилки, якщо nominal не можна перетворити на ціле число
                                print("Помилка: nominal має бути цілим числом.")
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            three = "3" if self.add_3_checkbox_box.isChecked() else ""
                            new_page.insert_text((x0 + 1, y0 + font_size),
                                                 f"{seria}{letter}{local_nominal}{three}S+0,8/1,2кВ",
                                                 fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        if re.match(r"21$", span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()

                        if re.match(r"W$", span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()

                        if re.match(r"37$", span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x_minus = 11
                            if year == 11 and week == 11:
                                x_minus = 6
                            new_page.insert_text((x0 - x_minus, y0 + font_size), f"{year}W{week}",
                                                 fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        if re.match(r"T$", span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()

                        if re.match(r"AS$", span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()

                        if re.match(r"65$", span_text):
                            x0, y0, a, b = bbox
                            a -= 1
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0 - 9, y0 + font_size), seria,
                                                 fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        if re.match(r"32x65mm 6$", span_text):
                            x0, y0, a, b = bbox
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), "32x65mm ",
                                                 fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        if re.match(r"00/5AM.L$", span_text):
                            x0, y0, a, b = bbox
                            a += 20
                            bbox = x0, y0, a, b
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0 - 2, y0 + font_size), f"{nominal}/5AM.L LUNGO",
                                                 fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

    def closeEvent(self, event):
        if os.path.exists(self.temp_preview_path):
            os.remove(self.temp_preview_path)
            print(f"Тимчасовий файл видалено: {self.temp_preview_path}")
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StickerGeneratorApp()
    window.show()
    sys.exit(app.exec())
