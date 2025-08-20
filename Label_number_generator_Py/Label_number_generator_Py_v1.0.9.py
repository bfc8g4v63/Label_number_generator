import sys
import os
import re

import pandas as pd
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

        export_xls_button = QPushButton("匯出為 Excel 97-2003 (.xls)")
        export_xls_button.clicked.connect(self.export_to_xls)

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
        layout.addWidget(export_xls_button)

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
        """外箱/棧板統一支援：英數混合 + 尾端數字，尾碼遞增"""

        def parse_tail(code: str):
            m = re.search(r'(.*?)(\d+)$', code.strip())
            if not m:
                raise ValueError(f"格式錯誤：{code}（需以數字結尾，例如 ABC001 或 HFAH220F201）")
            return m.group(1), int(m.group(2)), len(m.group(2))

        if self.mode_box.isChecked():

            sn_start_raw = self.sn_start.text().strip()
            sn_end_raw = self.sn_end.text().strip()
            sn_per_box = self.sn_per_box.text().strip()
            if not sn_per_box.isdigit():
                raise ValueError("每箱序號數量必須是數字。")
            sn_per_box = int(sn_per_box)

            s_prefix, s_no, s_len = parse_tail(sn_start_raw)
            e_prefix, e_no, _ = parse_tail(sn_end_raw)
            if s_prefix != e_prefix:
                raise ValueError("起始與結束序號的前綴必須相同。")

            box_format = self.box_code.text().strip()
            m = re.search(r'(.*?)(\d+)$', box_format)
            if not m:
                raise ValueError("箱號格式錯誤，需以數字結尾，例如 BX001。")
            box_prefix, box_no, box_len = m.group(1), int(m.group(2)), len(m.group(2))

            current = s_no
            last = e_no

            all_data = []
            while current <= last:
                row = {}
                row["BoxNo"] = f"C/NO.{box_prefix}{str(box_no).zfill(box_len)}"
                qr_lines = []
                for i in range(sn_per_box):
                    if current > last:
                        break
                    serial = f"{s_prefix}{str(current).zfill(s_len)}"
                    row[f"Serial{i+1}"] = serial
                    qr_lines.append(serial)
                    current += 1
                row["QRCodeContent"] = "\n".join(qr_lines)
                all_data.append(row)
                box_no += 1
            return pd.DataFrame(all_data).fillna("")

        else:

            box_start_code = self.plt_box_start.text().strip()
            box_end_code = self.plt_box_end.text().strip()
            boxes_per_pallet = self.boxes_per_pallet.text().strip()
            if not boxes_per_pallet.isdigit():
                raise ValueError("每棧板箱數必須是數字。")
            boxes_per_pallet = int(boxes_per_pallet)

            pallet_start_code = self.plt_start_code.text().strip()
            p_prefix, p_no, p_len = parse_tail(pallet_start_code)

            b_prefix, b_start, b_len = parse_tail(box_start_code)
            e_prefix, b_end, _ = parse_tail(box_end_code)
            if b_prefix != e_prefix:
                raise ValueError("箱號起訖的前綴必須相同。")

            current_box = b_start
            all_data = []
            while current_box <= b_end:
                row = {}
                row["PalletCode"] = f"PLT NO.:{p_prefix}{str(p_no).zfill(p_len)}"
                qr_lines = []
                for _ in range(boxes_per_pallet):
                    if current_box > b_end:
                        break
                    code_text = f"{b_prefix}{str(current_box).zfill(b_len)}"
                    qr_lines.append(code_text)
                    current_box += 1
                row["QRCodeContent"] = "\n".join(qr_lines)
                all_data.append(row)
                p_no += 1
            return pd.DataFrame(all_data)

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

    def export_to_xls(self):
        try:
            df = self.generate_data()

            max_rows, max_cols = 65536, 256
            if len(df) > max_rows or len(df.columns) > max_cols:
                QMessageBox.warning(
                    self, "超出 .xls 限制",
                    f".xls 最多 {max_rows} 列、{max_cols} 欄。\n"
                    f"目前：{len(df)} 列、{len(df.columns)} 欄。\n請改用 .xlsx 匯出。"
                )
                return

            save_path, _ = QFileDialog.getSaveFileName(
                self, "另存為 Excel 97-2003 檔", "", "Excel 97-2003 (*.xls)"
            )
            if not save_path:
                return
            if not save_path.endswith(".xls"):
                save_path += ".xls"

            try:
                import xlwt
            except ImportError:
                QMessageBox.critical(self, "缺少套件", "匯出 .xls 需要安裝 xlwt\n\n請執行：pip install xlwt")
                return

            wb = xlwt.Workbook()
            ws = wb.add_sheet("Sheet1")

            for c, col_name in enumerate(df.columns):
                ws.write(0, c, str(col_name))

            for r in range(len(df)):
                for c in range(len(df.columns)):
                    val = df.iat[r, c]
                    ws.write(r + 1, c, "" if pd.isna(val) else str(val))

            wb.save(save_path)

            self.populate_table(df)
            QMessageBox.information(self, "完成", f"Excel 97-2003 匯出成功！\n{save_path}")

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