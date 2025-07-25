import sys
import csv
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTableWidget,
    QTableWidgetItem, QPushButton, QVBoxLayout, QWidget,
    QMessageBox, QLabel, QHBoxLayout
)
from parser import parse_file
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


class FileParserApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Universal File Parser")
        self.is_dark_mode = True
        self.setStyleSheet(self.load_light_stylesheet())
        self.setGeometry(100, 100, 800, 600)
        self.parsed_data = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.label = QLabel("Upload a file (CSV, Excel, TXT, XML, DOCX):")
        layout.addWidget(self.label)

        self.mode_toggle_button = QPushButton("Toggle Dark Mode")
        self.mode_toggle_button.clicked.connect(self.toggle_theme)
        layout.addWidget(self.mode_toggle_button)        
        
        self.upload_button = QPushButton("Upload File")
        self.upload_button.clicked.connect(self.upload_file)
        layout.addWidget(self.upload_button)

        self.result_table = QTableWidget()
        layout.addWidget(self.result_table)

        download_layout = QHBoxLayout()

        self.download_csv_btn = QPushButton("Download as CSV")
        self.download_csv_btn.clicked.connect(lambda: self.download_file("csv"))
        download_layout.addWidget(self.download_csv_btn)


        self.download_pdf_btn = QPushButton("Download as PDF")
        self.download_pdf_btn.clicked.connect(self.download_pdf)
        download_layout.addWidget(self.download_pdf_btn)

        layout.addLayout(download_layout)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        
    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        if self.is_dark_mode:
            self.setStyleSheet(self.load_dark_stylesheet())
            self.mode_toggle_button.setText("Switch to Light Mode")
        else:
            self.setStyleSheet(self.load_light_stylesheet())
            self.mode_toggle_button.setText("Switch to Dark Mode")
        
        
    def load_light_stylesheet(self):
        return """
            QWidget {
                background-color: #ffffff;
                font-family: 'Roboto';
                font-size: 14px;
            }

            QLabel {
                color: #333;
                font-size: 16px;
            }

            QPushButton {
                background-color: #1949b0;
                color: white;
                border: none;
                padding: 10px 18px;
                border-radius: 6px;
            }

            QPushButton:hover {
                background-color: #1e91e3;
            }

            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #ccc;
                gridline-color: #ccc;
            }   

            QHeaderView::section {
                background-color: #1949b0;
                color: white;
                padding: 6px;
                border: none;
            }
        """


    def load_dark_stylesheet(self):
        return """
            QWidget {
                background-color: #1e1e1e;
                font-family: 'Segoe UI';
                font-size: 14px;
                color: #dcdcdc;
            }
            QLabel {
                color: #f0f0f0;
                font-size: 16px;
            }
            QPushButton {
                background-color: #3a3a3a;
                color: white;
                border: none;
                padding: 10px 18px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #555;
            }
            QTableWidget {
                background-color: #2b2b2b;
                color: #f0f0f0;
                border: 1px solid #555;
                gridline-color: #666;
            }
            QHeaderView::section {
                background-color: #444;
                color: white;
                padding: 6px;
                border: none;
            }
        """




    def upload_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select a file",
            "",
            "All Supported Files (*.csv *.xlsx *.xls *.txt *.xml *.docx);;"
            "CSV Files (*.csv);;Excel Files (*.xlsx *.xls);;Text Files (*.txt);;XML Files (*.xml);;Word Files (*.docx)",
            options=options,
        )
        if file_path:
            self.parsed_data = parse_file(file_path)
            self.populate_table()

    def populate_table(self):
        data = self.parsed_data

        if not data or not isinstance(data, list):
            self.result_table.clear()
            self.result_table.setRowCount(0)
            self.result_table.setColumnCount(0)
            QMessageBox.warning(self, "Warning", "No data to display or unsupported format.")
            return

        if isinstance(data[0], dict):
            headers = list(data[0].keys())
            self.result_table.setColumnCount(len(headers))
            self.result_table.setHorizontalHeaderLabels(headers)
            self.result_table.setRowCount(len(data))

            for row_idx, row_data in enumerate(data):
                for col_idx, header in enumerate(headers):
                    val = row_data.get(header, "")
                    item = QTableWidgetItem(str(val))
                    self.result_table.setItem(row_idx, col_idx, item)
        else:
            self.result_table.setColumnCount(1)
            self.result_table.setHorizontalHeaderLabels(["Line"])
            self.result_table.setRowCount(len(data))
            for row_idx, line in enumerate(data):
                item = QTableWidgetItem(str(line))
                self.result_table.setItem(row_idx, 0, item)

        self.result_table.resizeColumnsToContents()

    def download_file(self, file_type):
        if not self.parsed_data:
            QMessageBox.warning(self, "Warning", "No data to save.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, f"Save {file_type.upper()}", "", f"{file_type.upper()} Files (*.{file_type})", options=options)
        if not file_path:
            return

        if not file_path.lower().endswith(f".{file_type}"):
            file_path += f".{file_type}"

        try:
            if file_type == "csv":
                with open(file_path, "w", newline="", encoding="utf-8") as csvfile:
                    if isinstance(self.parsed_data[0], dict):
                        fieldnames = self.parsed_data[0].keys()
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        for row in self.parsed_data:
                            writer.writerow(row)
                    else:
                        writer = csv.writer(csvfile)
                        for row in self.parsed_data:
                            writer.writerow([str(row)])
            elif file_type == "txt":
                with open(file_path, "w", encoding="utf-8") as txtfile:
                    for row in self.parsed_data:
                        if isinstance(row, dict):
                            txtfile.write("\t".join(str(v) for v in row.values()) + "\n")
                        else:
                            txtfile.write(str(row) + "\n")

            QMessageBox.information(self, "Success", f"File saved as {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error saving file: {e}")

    def download_pdf(self):
        if not self.parsed_data:
            QMessageBox.warning(self, "Warning", "No data to save.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)", options=options)
        if not file_path:
            return

        if not file_path.lower().endswith(".pdf"):
            file_path += ".pdf"

        try:
            doc = SimpleDocTemplate(file_path, pagesize=landscape(letter),
                                    leftMargin=20, rightMargin=20,
                                    topMargin=20, bottomMargin=20)
            elements = []
            styles = getSampleStyleSheet()

            data = self.parsed_data

            if isinstance(data[0], dict):
                headers = list(data[0].keys())
                table_data = [headers]
                for row in data:
                    table_data.append([str(row.get(h, "")) for h in headers])
            else:
                headers = ["Line"]
                table_data = [headers]
                for row in data:
                    table_data.append([str(row)])

            # Wrap cells using Paragraph
            wrapped_data = [table_data[0]]  # header
            normal_style = styles["Normal"]
            for row in table_data[1:]:
                wrapped_row = [Paragraph(cell.replace("\n", "<br/>"), normal_style) for cell in row]
                wrapped_data.append(wrapped_row)

            # Adjust column widths based on total width
            total_width = landscape(letter)[0] - doc.leftMargin - doc.rightMargin
            num_cols = len(headers)
            col_widths = [total_width / num_cols] * num_cols

            table = Table(wrapped_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 6),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]))

            elements.append(table)
            doc.build(elements)

            QMessageBox.information(self, "Success", "PDF exported successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error saving PDF: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileParserApp()
    window.show()
    sys.exit(app.exec_())
