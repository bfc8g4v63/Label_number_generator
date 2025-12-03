import sys
import os
import re

import pandas as pd
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QRadioButton,
    QGroupBox,
)
from PyQt6.QtCore import Qt


class QRGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Label number generator")
        self.setGeometry(300, 200, 900, 600)
        self.setWindowIcon(QIcon("Nelson.ico"))
        self.setAcceptDrops(True)

        self.current_df = None
        self.import_source_df = None

        main_layout = QVBoxLayout()

        self.mode_box = QRadioButton("外箱")
        self.mode_pallet = QRadioButton("棧板")
        self.mode_box.setChecked(True)
        self.mode_box.toggled.connect(self.update_mode_ui)
        self.mode_pallet.toggled.connect(self.update_mode_ui)
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(self.mode_box)
        mode_layout.addWidget(self.mode_pallet)
        mode_layout.addStretch()
        main_layout.addLayout(mode_layout)

        self.sn_start = QLineEdit()
        self.sn_end = QLineEdit()
        self.sn_per_box = QLineEdit()
        self.box_code = QLineEdit()

        self.plt_box_start = QLineEdit()
        self.plt_box_end = QLineEdit()
        self.boxes_per_pallet = QLineEdit()
        self.plt_start_code = QLineEdit()

        for w in [
            self.sn_start,
            self.sn_end,
            self.sn_per_box,
            self.box_code,
            self.plt_box_start,
            self.plt_box_end,
            self.boxes_per_pallet,
            self.plt_start_code,
        ]:
            w.setMaximumWidth(260)

        left_form_layout = QVBoxLayout()
        left_form_layout.addWidget(QLabel("起始序號："))
        left_form_layout.addWidget(self.sn_start)
        left_form_layout.addWidget(QLabel("結束序號："))
        left_form_layout.addWidget(self.sn_end)
        left_form_layout.addWidget(QLabel("每箱序號數量："))
        left_form_layout.addWidget(self.sn_per_box)
        left_form_layout.addWidget(QLabel("箱號格式："))
        left_form_layout.addWidget(self.box_code)

        left_form_layout.addWidget(QLabel("箱號起始編號："))
        left_form_layout.addWidget(self.plt_box_start)
        left_form_layout.addWidget(QLabel("箱號結束編號："))
        left_form_layout.addWidget(self.plt_box_end)
        left_form_layout.addWidget(QLabel("每棧板箱數："))
        left_form_layout.addWidget(self.boxes_per_pallet)
        left_form_layout.addWidget(QLabel("棧板起始編號："))
        left_form_layout.addWidget(self.plt_start_code)

        left_form_group = QGroupBox("參數設定")
        left_form_group.setLayout(left_form_layout)

        self.scan_input = QLineEdit()
        self.scan_result_label = QLabel("尚未查詢")
        scan_font = self.scan_result_label.font()
        scan_font.setPointSize(40)
        self.scan_result_label.setFont(scan_font)
        self.scan_result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scan_result_label.setWordWrap(True)
        self.scan_result_label.setStyleSheet("color: #d32f2f; font-weight: bold;")

        self.scan_mode_continuous = QRadioButton("連續查詢模式")
        self.scan_mode_continuous.setChecked(True)

        self.scan_input.returnPressed.connect(self.lookup_serial)

        scan_layout = QVBoxLayout()
        scan_layout.addWidget(QLabel("掃描序號查詢"))
        scan_layout.addWidget(self.scan_input)
        scan_layout.addWidget(self.scan_mode_continuous)
        scan_layout.addWidget(QLabel("查詢結果："))
        scan_layout.addWidget(self.scan_result_label)
        scan_layout.addStretch()

        scan_group = QGroupBox("序號所在箱查詢")
        scan_group.setLayout(scan_layout)

        top_layout = QHBoxLayout()
        top_layout.addWidget(left_form_group)
        top_layout.addWidget(scan_group)
        main_layout.addLayout(top_layout)

        import_button = QPushButton("匯入檔案並重排")
        import_button.clicked.connect(lambda: self.import_from_file())

        export_xlsx_button = QPushButton("匯出為 Excel (.xlsx)")
        export_xlsx_button.clicked.connect(self.export_to_xlsx)

        export_csv_button = QPushButton("匯出為 CSV (.csv)")
        export_csv_button.clicked.connect(self.export_to_csv)

        export_xls_button = QPushButton("匯出為 Excel 97-2003 (.xls)")
        export_xls_button.clicked.connect(self.export_to_xls)

        main_layout.addWidget(import_button)
        main_layout.addWidget(export_xlsx_button)
        main_layout.addWidget(export_csv_button)
        main_layout.addWidget(export_xls_button)

        self.table = QTableWidget()
        main_layout.addWidget(self.table)

        self.setLayout(main_layout)
        self.update_mode_ui()

    @staticmethod
    def _parse_tail(code: str):
        m = re.match(r"^(.*?)(\d+)(\D*)$", code.strip())
        if not m:
            raise ValueError(
                f"格式錯誤：{code}（需在中間有連續數字作為流水號，例如 R5209-2508003 20250819 10578010 00000001 SC）"
            )
        return m.group(1), int(m.group(2)), len(m.group(2)), m.group(3)

    def update_mode_ui(self):
        is_box = self.mode_box.isChecked()
        labels_to_show_box = [
            "起始序號：",
            "結束序號：",
            "每箱序號數量：",
            "箱號格式：",
        ]
        labels_to_show_pallet = [
            "箱號起始編號：",
            "箱號結束編號：",
            "每棧板箱數：",
            "棧板起始編號：",
        ]

        for group in self.findChildren(QGroupBox):
            if group.title() == "參數設定":
                layout = group.layout()
                for i in range(layout.count()):
                    widget = layout.itemAt(i).widget()
                    if isinstance(widget, QLabel):
                        text = widget.text()
                        if text in labels_to_show_box:
                            widget.setVisible(is_box)
                        elif text in labels_to_show_pallet:
                            widget.setVisible(not is_box)
                    if isinstance(widget, QLineEdit):
                        if widget in [self.sn_start, self.sn_end, self.sn_per_box, self.box_code]:
                            widget.setVisible(is_box)
                        elif widget in [self.plt_box_start, self.plt_box_end, self.boxes_per_pallet, self.plt_start_code]:
                            widget.setVisible(not is_box)

    def generate_data(self):
        if self.mode_box.isChecked():
            sn_start_raw = self.sn_start.text().strip()
            sn_end_raw = self.sn_end.text().strip()
            sn_per_box = self.sn_per_box.text().strip()
            if not sn_per_box.isdigit():
                raise ValueError("每箱序號數量必須是數字。")
            sn_per_box = int(sn_per_box)

            s_prefix, s_no, s_len, s_suffix = self._parse_tail(sn_start_raw)
            e_prefix, e_no, _, e_suffix = self._parse_tail(sn_end_raw)
            if s_prefix != e_prefix or s_suffix != e_suffix:
                raise ValueError("起始與結束序號的前綴與尾碼必須相同。")

            box_format = self.box_code.text().strip()
            box_prefix, box_no, box_len, box_suffix = self._parse_tail(box_format)

            current = s_no
            last = e_no

            all_data = []
            while current <= last:
                row = {}
                row["BoxNo"] = f"C/NO.{box_prefix}{str(box_no).zfill(box_len)}{box_suffix}"
                qr_lines = []
                for i in range(sn_per_box):
                    if current > last:
                        break
                    serial = f"{s_prefix}{str(current).zfill(s_len)}{s_suffix}"
                    row[f"Serial{i + 1}"] = serial
                    qr_lines.append(serial)
                    current += 1
                row["QRCodeContent"] = "\n".join(qr_lines)
                all_data.append(row)
                box_no += 1
            df = pd.DataFrame(all_data).fillna("")
            self.current_df = df
            return df

        box_start_code = self.plt_box_start.text().strip()
        box_end_code = self.plt_box_end.text().strip()
        boxes_per_pallet = self.boxes_per_pallet.text().strip()
        if not boxes_per_pallet.isdigit():
            raise ValueError("每棧板箱數必須是數字。")
        boxes_per_pallet = int(boxes_per_pallet)

        pallet_start_code = self.plt_start_code.text().strip()
        p_prefix, p_no, p_len, p_suffix = self._parse_tail(pallet_start_code)

        b_prefix, b_start, b_len, b_suffix = self._parse_tail(box_start_code)
        e_prefix, b_end, _, e_suffix = self._parse_tail(box_end_code)
        if b_prefix != e_prefix or b_suffix != e_suffix:
            raise ValueError("箱號起訖的前綴與尾碼必須相同。")

        current_box = b_start
        all_data = []
        while current_box <= b_end:
            row = {}
            row["PalletCode"] = f"PLT NO.:{p_prefix}{str(p_no).zfill(p_len)}{p_suffix}"
            qr_lines = []
            for _ in range(boxes_per_pallet):
                if current_box > b_end:
                    break
                code_text = f"{b_prefix}{str(current_box).zfill(b_len)}{b_suffix}"
                qr_lines.append(code_text)
                current_box += 1
            row["QRCodeContent"] = "\n".join(qr_lines)
            all_data.append(row)
            p_no += 1
        df = pd.DataFrame(all_data)
        self.current_df = df
        return df

    def _read_source_dataframe(self, file_path: str) -> pd.DataFrame:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(file_path)
        elif ext == ".csv":
            df = pd.read_csv(file_path)
        else:
            raise ValueError("只支援 xls、xlsx、csv 檔案。")
        if df.shape[1] < 3:
            raise ValueError("匯入檔案至少需要三欄，例如：棧板號、箱號、客戶SN。")
        return df

    def _rebox_serials_from_dataframe(self, src_df: pd.DataFrame, sn_per_box: int) -> pd.DataFrame:
        sn_series = src_df.iloc[:, 2].astype(str)
        serials = [s for s in sn_series if s.strip()]
        if not serials:
            raise ValueError("檔案中找不到任何客戶SN。")

        box_code_text = self.box_code.text().strip()
        if box_code_text:
            b_prefix, b_no, b_len, b_suffix = self._parse_tail(box_code_text)
        else:
            first_box = str(src_df.iloc[0, 1])
            b_prefix, b_no, b_len, b_suffix = self._parse_tail(first_box)

        all_data = []
        idx = 0
        while idx < len(serials):
            row = {}
            row["BoxNo"] = f"C/NO.{b_prefix}{str(b_no).zfill(b_len)}{b_suffix}"
            qr_lines = []
            for i in range(sn_per_box):
                if idx >= len(serials):
                    break
                serial = serials[idx]
                row[f"Serial{i + 1}"] = serial
                qr_lines.append(serial)
                idx += 1
            row["QRCodeContent"] = "\n".join(qr_lines)
            all_data.append(row)
            b_no += 1

        df = pd.DataFrame(all_data).fillna("")
        return df

    def import_from_file(self, file_path: str | None = None):
        try:
            if file_path is None:
                file_path, _ = QFileDialog.getOpenFileName(
                    self,
                    "選擇要匯入的檔案",
                    "",
                    "Excel/CSV Files (*.xlsx *.xls *.csv)",
                )
                if not file_path:
                    return

            sn_per_box_text = self.sn_per_box.text().strip()
            if not sn_per_box_text.isdigit():
                raise ValueError("請先在「每箱序號數量」輸入要重排的每箱數量（必須是數字）。")
            sn_per_box = int(sn_per_box_text)
            if sn_per_box <= 0:
                raise ValueError("每箱序號數量必須大於 0。")

            src_df = self._read_source_dataframe(file_path)
            self.import_source_df = src_df.copy()

            df = self._rebox_serials_from_dataframe(src_df, sn_per_box)
            self.current_df = df
            self.populate_table(df)

            QMessageBox.information(
                self,
                "完成",
                f"匯入並重排完成，共 {len(df)} 箱。",
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "錯誤",
                "請將 xls / xlsx / csv 檔案拖曳至下方白色表格（白框）處，"
                "或使用「匯入檔案並重排」按鈕選擇檔案。\n\n詳細錯誤資訊："
                + str(e)
            )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile()
                ext = os.path.splitext(path)[1].lower()
                if ext in (".xlsx", ".xls", ".csv"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        file_path = urls[0].toLocalFile()
        if file_path:
            self.import_from_file(file_path)

    def lookup_serial(self):
        serial = self.scan_input.text().strip()
        if not serial:
            self.scan_result_label.setText("請輸入序號")
            if self.scan_mode_continuous.isChecked():
                self.scan_input.setFocus()
                self.scan_input.selectAll()
            return

        if self.current_df is None or self.current_df.empty:
            self.scan_result_label.setText("尚未產生資料\n請先匯出或預覽")
            if self.scan_mode_continuous.isChecked():
                self.scan_input.setFocus()
                self.scan_input.selectAll()
            return

        df = self.current_df
        target_col_name = None
        if "BoxNo" in df.columns:
            target_col_name = "BoxNo"
        elif "PalletCode" in df.columns:
            target_col_name = "PalletCode"

        found_row_index = None

        for idx in range(len(df)):
            row = df.iloc[idx]
            for col_name in df.columns:
                if col_name == target_col_name:
                    continue
                if str(row[col_name]) == serial:
                    found_row_index = idx
                    break
            if found_row_index is not None:
                break

        if found_row_index is None:
            self.scan_result_label.setText(f"序號{serial}\n未在任何箱中找到")
        else:
            row_number = found_row_index + 1
            if target_col_name == "PalletCode":
                self.scan_result_label.setText(f"序號{serial}\n第{row_number}棧板")
            else:
                self.scan_result_label.setText(f"序號{serial}\n第{row_number}箱")

        if self.scan_mode_continuous.isChecked():
            self.scan_input.setFocus()
            self.scan_input.selectAll()

    def export_to_xlsx(self):
        try:
            df = self.generate_data()
            save_path, _ = QFileDialog.getSaveFileName(self, "另存為 Excel 檔", "", "Excel Files (*.xlsx)")
            if save_path:
                if not save_path.endswith(".xlsx"):
                    save_path += ".xlsx"
                df.to_excel(save_path, index=False)
                self.populate_table(df)
                QMessageBox.information(self, "完成", f"Excel 匯出成功\n{save_path}")
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
                QMessageBox.information(self, "完成", f"CSV 匯出成功\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def export_to_xls(self):
        try:
            df = self.generate_data()

            max_rows, max_cols = 65536, 256
            if len(df) > max_rows or len(df.columns) > max_cols:
                QMessageBox.warning(
                    self,
                    "超出 .xls 限制",
                    f".xls 最多 {max_rows} 列、{max_cols} 欄。\n"
                    f"目前 {len(df)} 列、{len(df.columns)} 欄。\n請改用 .xlsx 匯出。",
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
                QMessageBox.critical(self, "缺少套件", "匯出 .xls 需要安裝 xlwt\n請執行 pip install xlwt")
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
            QMessageBox.information(self, "完成", f"Excel 97-2003 匯出成功\n{save_path}")

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