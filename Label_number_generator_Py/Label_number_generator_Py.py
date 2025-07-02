import sys
import os
import pandas as pd
import re
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox
)

class QRGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Label number generator")
        self.setGeometry(300, 200, 900, 600)
        self.setWindowIcon(QIcon("Nelson.ico"))
        layout = QVBoxLayout()

        self.sn_start = QLineEdit()
        self.sn_end = QLineEdit()
        self.sn_per_box = QLineEdit()
        self.box_code = QLineEdit()

        export_xlsx_button = QPushButton("匯出為 Excel (.xlsx)")
        export_xlsx_button.clicked.connect(self.export_to_xlsx)

        export_csv_button = QPushButton("匯出為 CSV (.csv)")
        export_csv_button.clicked.connect(self.export_to_csv)

        form_layout = QVBoxLayout()
        form_layout.addWidget(QLabel("起始序號："))
        form_layout.addWidget(self.sn_start)
        form_layout.addWidget(QLabel("結束序號："))
        form_layout.addWidget(self.sn_end)
        form_layout.addWidget(QLabel("每箱序號數量："))
        form_layout.addWidget(self.sn_per_box)
        form_layout.addWidget(QLabel("箱號格式："))
        form_layout.addWidget(self.box_code)

        layout.addLayout(form_layout)
        layout.addWidget(export_xlsx_button)
        layout.addWidget(export_csv_button)

        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.setLayout(layout)

    def generate_data(self):
        sn_start = int(self.sn_start.text())
        sn_end = int(self.sn_end.text())
        sn_per_box = int(self.sn_per_box.text())
        box_format = self.box_code.text()

        match = re.search(r'(.*?)(\d+)$', box_format)
        if not match:
            raise ValueError("箱號格式錯誤，必須以數字結尾。")
        prefix = match.group(1)
        numeric_part = match.group(2)
        box_no_numeric = int(numeric_part)
        numeric_len = len(numeric_part)

        all_data = []

        while sn_start <= sn_end:
            row_data = {}
            box_no_text = f"C/NO.{prefix}{str(box_no_numeric).zfill(numeric_len)}"
            row_data["BoxNo"] = box_no_text
            qr_text = ""

            for i in range(sn_per_box):
                current_sn = sn_start + i
                if current_sn > sn_end:
                    break
                serial = str(current_sn).zfill(12)
                row_data[f"Serial{i+1}"] = serial
                qr_text += serial + "\n"

            row_data["QRCodeContent"] = qr_text.strip()
            all_data.append(row_data)
            sn_start += sn_per_box
            box_no_numeric += 1

        df = pd.DataFrame(all_data)
        df = df.fillna("")
        return df

    def export_to_xlsx(self):
        try:
            df = self.generate_data()
            save_path, _ = QFileDialog.getSaveFileName(self, "另存為 Excel 檔", "", "Excel Files (*.xlsx)")
            if save_path:
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"
                df.to_excel(save_path, index=False)
                self.populate_table(df)
                QMessageBox.information(self, "完成", f"Excel 匯出成功！\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def export_to_csv(self):
        try:
            df = self.generate_data()
            save_path, _ = QFileDialog.getSaveFileName(self, "另存為 CSV 檔", "", "CSV Files (*.csv)")
            if save_path:
                if not save_path.endswith(".csv"):
                    save_path += ".csv"
                df.to_csv(save_path, index=False, encoding="utf-8-sig")
                self.populate_table(df)
                QMessageBox.information(self, "完成", f"CSV 匯出成功！\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def populate_table(self, df):
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.tolist())

        for row in range(len(df)):
            for col in range(len(df.columns)):
                self.table.setItem(row, col, QTableWidgetItem(str(df.iat[row, col])))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = QRGeneratorApp()
    window.show()
    sys.exit(app.exec())