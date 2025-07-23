import sys
import os
import pandas as pd
import re
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox, QRadioButton
)

class QRGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Label number generator")
        self.setGeometry(300, 200, 900, 600)
        self.setWindowIcon(QIcon("Nelson.ico"))
        layout = QVBoxLayout()

        self.mode_box = QRadioButton("外箱")
        self.mode_pallet = QRadioButton("棧板")
        self.mode_box.setChecked(True)
        self.mode_box.toggled.connect(self.update_mode_ui)
        self.mode_pallet.toggled.connect(self.update_mode_ui)
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(self.mode_box)
        mode_layout.addWidget(self.mode_pallet)
        layout.addLayout(mode_layout)

        self.sn_start = QLineEdit()
        self.sn_end = QLineEdit()
        self.sn_per_box = QLineEdit()
        self.box_code = QLineEdit()

        self.plt_box_start = QLineEdit()
        self.plt_box_end = QLineEdit()
        self.boxes_per_pallet = QLineEdit()
        self.plt_start_code = QLineEdit()

        export_xlsx_button = QPushButton("匯出為 Excel (.xlsx)")
        export_xlsx_button.clicked.connect(self.export_to_xlsx)

        export_csv_button = QPushButton("匯出為 CSV (.csv)")
        export_csv_button.clicked.connect(self.export_to_csv)

        self.form_layout = QVBoxLayout()
        self.form_layout.addWidget(QLabel("起始序號："))
        self.form_layout.addWidget(self.sn_start)
        self.form_layout.addWidget(QLabel("結束序號："))
        self.form_layout.addWidget(self.sn_end)
        self.form_layout.addWidget(QLabel("每箱序號數量："))
        self.form_layout.addWidget(self.sn_per_box)
        self.form_layout.addWidget(QLabel("箱號格式："))
        self.form_layout.addWidget(self.box_code)

        self.form_layout.addWidget(QLabel("箱號起始編號："))
        self.form_layout.addWidget(self.plt_box_start)
        self.form_layout.addWidget(QLabel("箱號結束編號："))
        self.form_layout.addWidget(self.plt_box_end)
        self.form_layout.addWidget(QLabel("每棧板箱數："))
        self.form_layout.addWidget(self.boxes_per_pallet)
        self.form_layout.addWidget(QLabel("棧板起始編號："))
        self.form_layout.addWidget(self.plt_start_code)

        layout.addLayout(self.form_layout)
        layout.addWidget(export_xlsx_button)
        layout.addWidget(export_csv_button)

        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.setLayout(layout)
        self.update_mode_ui()

    def update_mode_ui(self):
        """根據所選模式顯示對應欄位"""
        is_box = self.mode_box.isChecked()
        for i in range(self.form_layout.count()):
            widget = self.form_layout.itemAt(i).widget()
            if isinstance(widget, QLabel):
                text = widget.text()
                if "序號" in text or "箱號格式" in text:
                    widget.setVisible(is_box)
                elif "箱號" in text or "棧板" in text:
                    widget.setVisible(not is_box)
            if isinstance(widget, QLineEdit):
                if widget in [self.sn_start, self.sn_end, self.sn_per_box, self.box_code]:
                    widget.setVisible(is_box)
                elif widget in [self.plt_box_start, self.plt_box_end, self.boxes_per_pallet, self.plt_start_code]:
                    widget.setVisible(not is_box)

    def generate_data(self):
        """根據模式產生資料表格"""
        if self.mode_box.isChecked():
            sn_start = int(self.sn_start.text())
            sn_end = int(self.sn_end.text())
            sn_per_box = self.sn_per_box.text()
            if not sn_per_box.isdigit():
                raise ValueError("每箱序號數量必須是數字。")
            sn_per_box = int(sn_per_box)
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
        else:
            box_start_code = self.plt_box_start.text()
            box_end_code = self.plt_box_end.text()
            boxes_per_pallet = self.boxes_per_pallet.text()
            if not boxes_per_pallet.isdigit():
                raise ValueError("每棧板箱數必須是數字。")
            boxes_per_pallet = int(boxes_per_pallet)

            pallet_start_code = self.plt_start_code.text().strip()
            if not re.match(r'^[A-Za-z]*\d+$', pallet_start_code):
                raise ValueError("棧板起始編號格式錯誤，需為英文+數字組合，例如 P001。")
            m_pallet = re.match(r'([A-Za-z]*)(\d+)', pallet_start_code)
            pallet_prefix = m_pallet.group(1)
            pallet_no = int(m_pallet.group(2))
            pallet_num_len = len(m_pallet.group(2))

            def parse_code(code):
                m = re.match(r'([A-Za-z]*)(\d+)', code)
                if not m:
                    raise ValueError(f"無法解析箱號格式：{code}")
                return m.group(1), int(m.group(2)), len(m.group(2))

            prefix, box_start, num_len = parse_code(box_start_code)
            _, box_end, _ = parse_code(box_end_code)

            current_box = box_start
            all_data = []
            while current_box <= box_end:
                row_data = {}
                row_data["PalletCode"] = f"PLT NO.:{pallet_prefix}{str(pallet_no).zfill(pallet_num_len)}"
                qr_lines = []
                for _ in range(boxes_per_pallet):
                    if current_box > box_end:
                        break
                    box_code = f"C/NO.{prefix}{str(current_box).zfill(num_len)}"
                    qr_lines.append(box_code)
                    current_box += 1
                row_data["QRCodeContent"] = "\n".join(qr_lines)
                all_data.append(row_data)
                pallet_no += 1

            df = pd.DataFrame(all_data)
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
