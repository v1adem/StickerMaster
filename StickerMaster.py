import sys
import os
import fitz  # PyMuPDF
import re
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog


def resource_path(relative_path):
    """Отримує шлях до вбудованих ресурсів для PyInstaller"""
    if getattr(sys, 'frozen', False):  # Якщо програма запущена як .exe
        base_path = sys._MEIPASS
    else:  # Якщо це середовище розробки
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class StickerGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.template_label = QLabel("Оберіть PDF-шаблон:")
        self.template_button = QPushButton("Вибрати файл")
        self.template_button.clicked.connect(self.select_template)
        layout.addWidget(self.template_label)
        layout.addWidget(self.template_button)

        self.prefix_label = QLabel("Перші 6 цифр серійного номера:")
        self.prefix_input = QLineEdit()
        layout.addWidget(self.prefix_label)
        layout.addWidget(self.prefix_input)

        self.nominal_label = QLabel("Номінал:")
        self.nominal_input = QLineEdit()
        layout.addWidget(self.nominal_label)
        layout.addWidget(self.nominal_input)

        self.count_label = QLabel("Кількість пристроїв у партії:")
        self.count_input = QLineEdit()
        layout.addWidget(self.count_label)
        layout.addWidget(self.count_input)

        self.date_label = QLabel("Дата виготовлення (РікWТиждень, напр. 24W18):")
        self.date_input = QLineEdit()
        layout.addWidget(self.date_label)
        layout.addWidget(self.date_input)

        self.generate_button = QPushButton("Згенерувати PDF")
        self.generate_button.clicked.connect(self.generate_pdfs)
        layout.addWidget(self.generate_button)

        self.setLayout(layout)
        self.setWindowTitle("Генератор стікерів")
        self.resize(400, 250)
        self.template_path = ""

    def select_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Виберіть PDF-шаблон", "", "PDF Files (*.pdf)")
        if file_path:
            self.template_path = file_path

    def generate_pdfs(self):
        if not self.template_path:
            return

        prefix = self.prefix_input.text()
        count = int(self.count_input.text())
        date_code = self.date_input.text()
        nominal = self.nominal_input.text()

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
            self.modify_pdf(self.template_path, doc_output, serial_number, date_code, nominal)

        # Зберегти фінальний PDF файл з усіма сторінками
        doc_output.save(output_pdf)
        doc_output.close()

    def extract_article_code(self, input_pdf, nominal):
        """Функція для витягування артикулу з шаблону PDF та зміни номіналу"""
        doc = fitz.open(input_pdf)
        article_code_pattern = r"TA32750C600SE"  # Приклад артикулу, можна адаптувати під інші варіанти
        article_code = None

        # Перевірка на наявність артикулу в тексті першої сторінки
        page = doc[0]  # Першу сторінку PDF
        text = page.get_text("text")
        match = re.search(article_code_pattern, text)

        if match:
            # Заміна номіналу в артикулі
            article_code = match.group(0).replace("C600", f"C{nominal}")

        doc.close()
        return article_code

    def modify_pdf(self, input_pdf, doc_output, serial_number, date_code, nominal):
        doc = fitz.open(input_pdf)
        date_pattern = r"\d{2}W\d{2}"  # Шаблон для року і тижня
        serial_pattern = r"\d{10}"  # Шаблон для серійного номера
        ipr_pattern = r"Ipr \d+A"  # Шаблон для пошуку "Ipr" та "A"
        article_pattern = r"TA32750C600SE$"  # Шаблон для артикулу
        start_pattern = r"TA327$"  # Шаблон для початку заміни

        font_mapping = {
            "MyriadPro-Regular": resource_path("MyriadPro-Regular.ttf"),
            "Arial-BoldMT": resource_path("arial-mt-bold.ttf"),
            "ArialMT": resource_path("ArialMT-Light.ttf")
        }
        print(font_mapping)

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
                            new_page.add_redact_annot(bbox, fill=[255, 255, 255])
                            new_page.apply_redactions()
                            x0, y0, _, _ = bbox
                            new_page.insert_text((x0, y0 + font_size), span_text.replace("TA327", f"TA{nominal}"),
                                                 fontsize=font_size, color=(0, 0, 0), fontfile=font_path,
                                                 fontname=font_name, border_width=1.5)
                            print(span)

        doc.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StickerGeneratorApp()
    window.show()
    sys.exit(app.exec())


