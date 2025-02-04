import sys
import os
import fitz  # PyMuPDF
import re
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, \
    QMessageBox, QGroupBox, QGridLayout, QTabWidget
from PySide6.QtGui import QPixmap, Qt
from PIL import ImageQt, Image


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class StickerGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.template_pixmap = None
        self.preview_pixmap = None

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
        self.IME_standard_template_button = QPushButton("Вибрати")
        self.IME_standard_template_button.clicked.connect(self.select_template)

        self.prefix_IME_standard_label = QLabel("Перші 6 цифр серійного номеру:")
        self.prefix_IME_standard_input = QLineEdit()

        self.short_prefix_IME_standard_label = QLabel("Короткий артикул:")
        self.short_prefix_IME_standard_input = QLineEdit()

        self.nominal_IME_standard_label = QLabel("Номінал:")
        self.nominal_IME_standard_input = QLineEdit()

        self.va_IME_standard_label = QLabel("Навантаження VA:")
        self.va_IME_standard_input = QLineEdit()

        self.count_IME_standard_label = QLabel("Кількість:")
        self.count_IME_standard_input = QLineEdit()

        self.year_IME_standard_label = QLabel("Рік:")
        self.year_IME_standard_input = QLineEdit()

        self.week_IME_standard_label = QLabel("Тиждень:")
        self.week_IME_standard_input = QLineEdit()

        # Встановлюємо рамку для QLineEdit
        self.prefix_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        self.generate_IME_standard_button = QPushButton("Згенерувати PDF")
        self.generate_IME_standard_button.clicked.connect(self.generate_IME_standard_pdfs)

        self.preview_IME_standard_button = QPushButton("Попередній перегляд")
        self.preview_IME_standard_button.clicked.connect(self.show_IME_standard_preview)

    def create_IME_box_form_elements(self):
        self.IME_box_template_label = QLabel("PDF-шаблон:")
        self.IME_box_template_button = QPushButton("Вибрати")
        self.IME_box_template_button.clicked.connect(self.select_template)

        self.IME_box_serial_pattern_label = QLabel("Оригінальний серійний номер\nОБОВ'ЯЗКОВО!")
        self.IME_box_serial_pattern_label.setStyleSheet("font-weight: bold; color: red;")
        self.IME_box_serial_pattern_input = QLineEdit()

        self.serial_IME_box_label = QLabel("Серійний номер:")
        self.serial_IME_box_input = QLineEdit()

        self.short_prefix_IME_box_label = QLabel("Короткий артикул:")
        self.short_prefix_IME_box_input = QLineEdit()

        self.voltage_IME_box_label = QLabel("Напруга:")
        self.voltage_IME_box_input = QLineEdit()

        self.current_IME_box_label = QLabel("Струм A:")
        self.current_IME_box_input = QLineEdit()

        self.count_IME_box_label = QLabel("Кількість в коробці:")
        self.count_IME_box_input = QLineEdit()

        self.box_count_IME_box_label = QLabel("Кількість коробок:")
        self.box_count_IME_box_input = QLineEdit()

        self.year_IME_box_label = QLabel("Рік:")
        self.year_IME_box_input = QLineEdit()

        self.week_IME_box_label = QLabel("Тиждень:")
        self.week_IME_box_input = QLineEdit()

        # Встановлюємо рамку для QLineEdit
        self.IME_box_serial_pattern_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.serial_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.short_prefix_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.voltage_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.current_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.box_count_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_IME_box_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        self.generate_IME_box_button = QPushButton("Згенерувати PDF")
        self.generate_IME_box_button.clicked.connect(self.generate_IME_standard_pdfs)

        self.preview_IME_box_button = QPushButton("Попередній перегляд")
        self.preview_IME_box_button.clicked.connect(self.show_IME_standard_preview)

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
        input_layout.addWidget(self.IME_standard_template_button, 0, 1)

        input_layout.addWidget(self.prefix_IME_standard_label, 1, 0)
        input_layout.addWidget(self.prefix_IME_standard_input, 1, 1)

        input_layout.addWidget(self.short_prefix_IME_standard_label, 2, 0)
        input_layout.addWidget(self.short_prefix_IME_standard_input, 2, 1)

        input_layout.addWidget(self.nominal_IME_standard_label, 3, 0)
        input_layout.addWidget(self.nominal_IME_standard_input, 3, 1)

        input_layout.addWidget(self.va_IME_standard_label, 4, 0)
        input_layout.addWidget(self.va_IME_standard_input, 4, 1)

        input_layout.addWidget(self.count_IME_standard_label, 5, 0)
        input_layout.addWidget(self.count_IME_standard_input, 5, 1)

        input_layout.addWidget(self.year_IME_standard_label, 6, 0)
        input_layout.addWidget(self.year_IME_standard_input, 6, 1)

        input_layout.addWidget(self.week_IME_standard_label, 7, 0)
        input_layout.addWidget(self.week_IME_standard_input, 7, 1)

        # Встановлюємо рамку для QLineEdit
        self.prefix_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.short_prefix_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_IME_standard_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        input_layout.addWidget(self.generate_IME_standard_button, 8, 0, 1, 2)

        input_layout.addWidget(self.preview_IME_standard_button, 9, 0, 1, 2)

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
        input_layout.addWidget(self.IME_box_template_button, 0, 1)

        input_layout.addWidget(self.short_prefix_IME_box_label, 1, 0)
        input_layout.addWidget(self.short_prefix_IME_box_input, 1, 1)

        input_layout.addWidget(self.IME_box_serial_pattern_label, 2, 0)
        input_layout.addWidget(self.IME_box_serial_pattern_input, 2, 1)

        input_layout.addWidget(self.serial_IME_box_label, 3, 0)
        input_layout.addWidget(self.serial_IME_box_input, 3, 1)

        input_layout.addWidget(self.voltage_IME_box_label, 4, 0)
        input_layout.addWidget(self.voltage_IME_box_input, 4, 1)

        input_layout.addWidget(self.current_IME_box_label, 5, 0)
        input_layout.addWidget(self.current_IME_box_input, 5, 1)

        input_layout.addWidget(self.count_IME_box_label, 6, 0)
        input_layout.addWidget(self.count_IME_box_input, 6, 1)

        input_layout.addWidget(self.year_IME_box_label, 7, 0)
        input_layout.addWidget(self.year_IME_box_input, 7, 1)

        input_layout.addWidget(self.week_IME_box_label, 8, 0)
        input_layout.addWidget(self.week_IME_box_input, 8, 1)

        input_layout.addWidget(self.generate_IME_box_button, 9, 0, 1, 2)  # span columns
        input_layout.addWidget(self.preview_IME_box_button, 10, 0, 1, 2)  # span columns

        self.boxes_tab.setLayout(input_layout)  # Встановлюємо layout для вкладки

    def select_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Виберіть PDF-шаблон", "", "PDF Files (*.pdf)")
        if file_path:
            self.template_path = file_path
            self.display_template_preview(file_path)

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

        article_code = self.extract_article_code(self.template_path, nominal)

        if not article_code:
            return

        self.temp_preview_path = os.path.join(os.getcwd(), "tmp_preview.pdf")
        doc_output = fitz.open()
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
        short_prefix = self.short_prefix_IME_box_input.text()
        count = int(self.count_IME_box_input.text())
        year = self.year_IME_box_input.text()
        week = self.week_IME_box_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_standard_input.text()
        va = self.va_IME_standard_input.text()

        self.temp_preview_path = os.path.join(os.getcwd(), "tmp_preview.pdf")
        doc_output = fitz.open()
        self.modify_IME_box_pdf(self.template_path, doc_output, date_code, nominal, va, short_prefix)
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

        # Зчитуємо артикул з шаблону (припускаємо, що артикул міститься в першій сторінці шаблону)
        article_code = self.extract_article_code(self.template_path, nominal)

        if not article_code:
            return  # Якщо артикул не знайдений, зупиняємо виконання

        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження")
        if not folder:
            return

        # Формуємо назву файлу
        output_pdf = os.path.join(folder, f"{short_prefix} {date_code}-{count}шт.pdf")
        doc_output = fitz.open()  # Новий PDF документ для збереження всіх сторінок

        for i in range(1, count + 1):
            serial_number = f"{prefix}{i:04}"
            self.modify_IME_standart_pdf(self.template_path, doc_output, serial_number, date_code, nominal, va,
                                         short_prefix)

        # Зберегти фінальний PDF файл з усіма сторінками
        doc_output.save(output_pdf)
        doc_output.close()

    def generate_IME_box_pdfs(self):
        if not self.template_path:
            return
        serail = self.serial_IME_box_input.text()
        short_prefix = self.short_prefix_IME_box_input.text()
        count = int(self.count_IME_box_input.text())
        year = self.year_IME_box_input.text()
        week = self.week_IME_box_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_IME_standard_input.text()
        va = self.va_IME_standard_input.text()

        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження")
        if not folder:
            return

        # Формуємо назву файлу
        output_pdf = os.path.join(folder, f"{short_prefix} {nominal} {date_code} КОРОБКА.pdf")
        doc_output = fitz.open()  # Новий PDF документ для збереження всіх сторінок

        for i in range(1, count + 1):
            serial_number = None
            self.modify_IME_box_pdf(self.template_path, doc_output, serail, date_code, nominal, va, short_prefix)

        # Зберегти фінальний PDF файл з усіма сторінками
        doc_output.save(output_pdf)
        doc_output.close()

    def extract_article_code(self, input_pdf, nominal):
        """Функція для витягування артикулу з шаблону PDF та зміни номіналу"""
        doc = fitz.open(input_pdf)
        article_code_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+[C|B|D]\d+SE$"
        article_code = None

        # Перевірка на наявність артикулу в тексті першої сторінки
        page = doc[0]  # Першу сторінку PDF
        text = page.get_text("text")
        match = re.search(article_code_pattern, text, re.MULTILINE)

        if match:
            # Заміна номіналу в артикулі
            article_code = match.group(0).replace("C600", f"C{nominal}")

        doc.close()
        return article_code

    def modify_IME_standart_pdf(self, input_pdf, doc_output, serial_number, date_code, nominal, va, short_prefix):
        doc = fitz.open(input_pdf)
        date_pattern = r"\d{2}W\d{2}"  # Шаблон для року і тижня
        serial_pattern = r"\d{10}"  # Шаблон для серійного номера
        ipr_pattern = r"Ipr \d+A"  # Шаблон для пошуку "Ipr" та "A"
        article_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+[C|B|D]\d+SE$"  # Шаблон для артикулу
        start_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+$"  # Шаблон для початку заміни

        font_mapping = {
            "MyriadPro-Regular": resource_path("fonts/MyriadPro-Regular.ttf"),
            "MyriadPro-Bold": resource_path("fonts/MyriadPro-Regular.ttf"),
            "Arial-BoldMT": resource_path("fonts/arial-mt-bold.ttf"),
            "ArialMT": resource_path("fonts/ArialMT-Light.ttf")
        }

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

                        font_path = font_mapping.get(font_name)

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
                            print(font_name)
                            new_page.insert_text((x0, y0 + font_size), date_code, fontsize=font_size, color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        # Заміна номіналу в тексті після "Ipr" та перед "A"
                        if re.match(ipr_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), f"Ipr {nominal}A", fontsize=font_size,
                                                 color=(0, 0, 0))

                            # Заміна артикулу на кожній сторінці
                        if re.fullmatch(article_pattern, span_text):
                            match = re.search(r"\d[ABC]\d\d\d", span_text)
                            if match:
                                num_before = match.group(0)[0]
                                letter = match.group(0)[1]  # Отримуємо першу літеру
                                new_text = span_text.replace(match.group(0),
                                                             f"{num_before}{letter}{nominal}")  # Замінюємо літеру + цифри на літеру + nominal
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size), new_text,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                        if re.match(start_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), short_prefix,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if re.match(r"15$", span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size - 0.1), va,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

        doc.close()

    def modify_IME_box_pdf(self, input_pdf, doc_output, serial, date_code, nominal, va, short_prefix):
        doc = fitz.open(input_pdf)
        date_pattern = r"\d{2}W\d{2}"  # Шаблон для року і тижня
        serial_pattern = fr"{self.IME_box_serial_pattern_input.text()}$"  # Шаблон для серійного номера
        ipr_pattern = r"Ipr \d+A"  # Шаблон для пошуку "Ipr" та "A"
        article_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+[C|B|D]\d+SE$"  # Шаблон для артикулу
        start_pattern = r"^(TA|TT|TAS|TASS|TASO)\d+$"  # Шаблон для початку заміни

        font_mapping = {
            "MyriadPro-Regular": resource_path("fonts/MyriadPro-Regular.ttf"),
            "MyriadPro-Bold": resource_path("fonts/MyriadPro-Regular.ttf"),
            "Arial-BoldMT": resource_path("fonts/arial-mt-bold.ttf"),
            "ArialMT": resource_path("fonts/ArialMT-Light.ttf")
        }

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

                        font_path = font_mapping.get(font_name)

                        x0, y0, a, b = bbox
                        y0 += 0.1
                        bbox = x0, y0, a, b

                        # Заміна серійного номера
                        if re.match(serial_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            if serial is not None:
                                new_page.insert_text((x0, y0 + font_size), serial, fontsize=font_size,
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
                            # new_page.insert_text((x0, y0 + font_size), f"Ipr {nominal}A", fontsize=font_size,
                            #                     color=(0, 0, 0))

                            # Заміна артикулу на кожній сторінці
                        if re.fullmatch(article_pattern, span_text):
                            match = re.search(r"\d[ABC]\d\d\d", span_text)
                            if match:
                                num_before = match.group(0)[0]
                                letter = match.group(0)[1]  # Отримуємо першу літеру
                                new_text = span_text.replace(match.group(0),
                                                             f"{num_before}{letter}{nominal}")  # Замінюємо літеру + цифри на літеру + nominal
                                new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                                new_page.apply_redactions()
                                new_page.insert_text((x0, y0 + font_size), new_text,
                                                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                     fontname=font_name)

                        if re.match(start_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            new_page.insert_text((x0, y0 + font_size), short_prefix,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if re.match(r"15$", span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            # new_page.insert_text((x0, y0 + font_size - 0.1), va,
                            #                     fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                            #                     fontname=font_name)

        doc.close()

    def closeEvent(self, event):
        # Перевірка та видалення тимчасового файлу
        if os.path.exists(self.temp_preview_path):
            os.remove(self.temp_preview_path)
            print(f"Тимчасовий файл видалено: {self.temp_preview_path}")  # Для дебагу
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StickerGeneratorApp()
    window.show()
    sys.exit(app.exec())
