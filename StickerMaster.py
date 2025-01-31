import sys
import os
import fitz  # PyMuPDF
import re
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, \
    QHBoxLayout, \
    QMessageBox, QGroupBox, QGridLayout, QTabWidget
from PySide6.QtGui import QPixmap
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
        main_layout = QHBoxLayout()

        # Створюємо TabWidget
        self.tabs = QTabWidget()

        # Створюємо вкладку "IME Standart"
        self.standard_tab = QWidget()
        self.create_standard_tab()  # Функція для створення вмісту вкладки
        self.tabs.addTab(self.standard_tab, "IME Standard")

        # Створюємо вкладку "IME Коробки"
        self.boxes_tab = QWidget()
        self.create_boxes_tab()  # Функція для створення вмісту вкладки
        self.tabs.addTab(self.boxes_tab, "IME Коробки")

        main_layout.addWidget(self.tabs)  # Додаємо TabWidget до головного layout

        preview_layout = QVBoxLayout()
        self.template_preview_label = QLabel()
        self.preview_preview_label = QLabel()
        preview_layout.addWidget(self.template_preview_label)
        preview_layout.addWidget(self.preview_preview_label)
        main_layout.addLayout(preview_layout)

        self.setLayout(main_layout)
        self.setWindowTitle("Sticker Master IME")
        self.resize(400, 300)
        self.template_path = ""
        self.temp_preview_path = ""

    def create_standard_tab(self):
        # Група для параметрів генерації
        input_group = QGroupBox("Параметри генерації")  # Створюємо групу
        input_layout = QGridLayout()

        input_layout.setHorizontalSpacing(5)  # Зменшуємо горизонтальний відступ
        input_layout.setVerticalSpacing(5)  # Зменшуємо вертикальний відступ

        # Встановлюємо рамку для групи
        input_group.setStyleSheet("QGroupBox { border: 1px solid gray; }")

        # Розміщуємо віджети в QGridLayout
        self.template_label = QLabel("PDF-шаблон:")
        self.template_button = QPushButton("Вибрати")
        self.template_button.clicked.connect(self.select_template)
        input_layout.addWidget(self.template_label, 0, 0)  # row, col
        input_layout.addWidget(self.template_button, 0, 1)

        self.prefix_label = QLabel("Перші 6 цифр серійного номеру:")
        self.prefix_input = QLineEdit()
        input_layout.addWidget(self.prefix_label, 1, 0)
        input_layout.addWidget(self.prefix_input, 1, 1)

        self.nominal_label = QLabel("Номінал:")
        self.nominal_input = QLineEdit()
        input_layout.addWidget(self.nominal_label, 2, 0)
        input_layout.addWidget(self.nominal_input, 2, 1)

        self.va_label = QLabel("Навантаження VA:")
        self.va_input = QLineEdit()
        input_layout.addWidget(self.va_label, 3, 0)
        input_layout.addWidget(self.va_input, 3, 1)

        self.count_label = QLabel("Кількість:")
        self.count_input = QLineEdit()
        input_layout.addWidget(self.count_label, 4, 0)
        input_layout.addWidget(self.count_input, 4, 1)

        self.year_label = QLabel("Рік:")
        self.year_input = QLineEdit()
        input_layout.addWidget(self.year_label, 5, 0)
        input_layout.addWidget(self.year_input, 5, 1)

        self.week_label = QLabel("Тиждень:")
        self.week_input = QLineEdit()
        input_layout.addWidget(self.week_label, 6, 0)
        input_layout.addWidget(self.week_input, 6, 1)

        # Встановлюємо рамку для QLineEdit
        self.prefix_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        self.generate_button = QPushButton("Згенерувати PDF")
        self.generate_button.clicked.connect(self.generate_pdfs)
        input_layout.addWidget(self.generate_button, 7, 0, 1, 2)  # span columns

        self.preview_button = QPushButton("Попередній перегляд")
        self.preview_button.clicked.connect(self.show_preview)
        input_layout.addWidget(self.preview_button, 8, 0, 1, 2)  # span columns

        self.standard_tab.setLayout(input_layout)  # Встановлюємо layout для вкладки

    def create_boxes_tab(self):
        # Група для параметрів генерації
        input_group = QGroupBox("Параметри генерації")  # Створюємо групу
        input_layout = QGridLayout()

        input_layout.setHorizontalSpacing(5)  # Зменшуємо горизонтальний відступ
        input_layout.setVerticalSpacing(5)  # Зменшуємо вертикальний відступ

        # Встановлюємо рамку для групи
        input_group.setStyleSheet("QGroupBox { border: 1px solid gray; }")

        # Розміщуємо віджети в QGridLayout
        self.template_label = QLabel("PDF-шаблон:")
        self.template_button = QPushButton("Вибрати")
        self.template_button.clicked.connect(self.select_template)
        input_layout.addWidget(self.template_label, 0, 0)  # row, col
        input_layout.addWidget(self.template_button, 0, 1)

        self.nominal_label = QLabel("Номінал:")
        self.nominal_input = QLineEdit()
        input_layout.addWidget(self.nominal_label, 1, 0)
        input_layout.addWidget(self.nominal_input, 1, 1)

        self.va_label = QLabel("Навантаження VA:")
        self.va_input = QLineEdit()
        input_layout.addWidget(self.va_label, 2, 0)
        input_layout.addWidget(self.va_input, 2, 1)

        self.count_label = QLabel("Кількість:")
        self.count_input = QLineEdit()
        input_layout.addWidget(self.count_label, 3, 0)
        input_layout.addWidget(self.count_input, 3, 1)

        self.year_label = QLabel("Рік:")
        self.year_input = QLineEdit()
        input_layout.addWidget(self.year_label, 4, 0)
        input_layout.addWidget(self.year_input, 4, 1)

        self.week_label = QLabel("Тиждень:")
        self.week_input = QLineEdit()
        input_layout.addWidget(self.week_label, 5, 0)
        input_layout.addWidget(self.week_input, 5, 1)

        # Встановлюємо рамку для QLineEdit
        self.prefix_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.nominal_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.va_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.count_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.year_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")
        self.week_input.setStyleSheet("QLineEdit { border: 1px solid gray; }")

        self.generate_button = QPushButton("Згенерувати PDF")
        self.generate_button.clicked.connect(self.generate_pdfs)
        input_layout.addWidget(self.generate_button, 6, 0, 1, 2)  # span columns

        self.preview_button = QPushButton("Попередній перегляд")
        self.preview_button.clicked.connect(self.show_preview)
        input_layout.addWidget(self.preview_button, 7, 0, 1, 2)  # span columns

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
            matrix = fitz.Matrix(4, 4)  # Збільшення в 4 рази (400%)
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

    def show_preview(self):
        if not self.template_path:
            return

        prefix = self.prefix_input.text()
        date_code = self.date_input.text()
        nominal = self.nominal_input.text()
        va = self.va_input.text()

        article_code = self.extract_article_code(self.template_path, nominal)

        if not article_code:
            return

        self.temp_preview_path = os.path.join(os.getcwd(), "tmp_preview.pdf")
        doc_output = fitz.open()
        self.modify_pdf(self.template_path, doc_output, f"{prefix}0001", date_code, nominal, va)
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

    def generate_pdfs(self):
        if not self.template_path:
            return

        prefix = self.prefix_input.text()
        count = int(self.count_input.text())
        year = self.year_input.text()
        week = self.week_input.text()
        date_code = f"{year}W{week}"
        nominal = self.nominal_input.text()
        va = self.va_input.text()

        # Зчитуємо артикул з шаблону (припускаємо, що артикул міститься в першій сторінці шаблону)
        article_code = self.extract_article_code(self.template_path, nominal)

        if not article_code:
            return  # Якщо артикул не знайдений, зупиняємо виконання

        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження")
        if not folder:
            return

        # Формуємо назву файлу
        output_pdf = os.path.join(folder, f"{article_code} {date_code}-{count}шт.pdf")
        doc_output = fitz.open()  # Новий PDF документ для збереження всіх сторінок

        for i in range(1, count + 1):
            serial_number = f"{prefix}{i:04}"
            self.modify_pdf(self.template_path, doc_output, serial_number, date_code, nominal, va)

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

    def modify_pdf(self, input_pdf, doc_output, serial_number, date_code, nominal, va):
        doc = fitz.open(input_pdf)
        date_pattern = r"\d{2}W\d{2}"  # Шаблон для року і тижня
        serial_pattern = r"\d{10}"  # Шаблон для серійного номера
        ipr_pattern = r"Ipr \d+A"  # Шаблон для пошуку "Ipr" та "A"
        article_pattern = r"TA32750C600SE$"  # Шаблон для артикулу
        start_pattern = r"TA327$"  # Шаблон для початку заміни

        font_mapping = {
            "MyriadPro-Regular": resource_path("fonts/MyriadPro-Regular.ttf"),
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

                        # Заміна серійного номера
                        if re.match(serial_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), serial_number, fontsize=font_size,
                                                 color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        # Заміна дати
                        if re.match(date_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), date_code, fontsize=font_size, color=(0, 0, 0),
                                                 fontfile=font_path, fontname=font_name)

                        # Заміна номіналу в тексті після "Ipr" та перед "A"
                        if re.match(ipr_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), f"Ipr {nominal}A", fontsize=font_size,
                                                 color=(0, 0, 0))

                            # Заміна артикулу на кожній сторінці
                        if re.fullmatch(article_pattern, span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), span_text.replace("C600", f"C{nominal}"),
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if re.match(start_pattern, span_text):
                            print(font_path)
                            print(font_name)
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), "TA327",
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

                        if re.match(r"15$", span_text):
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), va,
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name)

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
