import os
import sys
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt


class ConvertToPDFApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word & PPT 批量转 PDF")
        self.resize(400, 200)

        self.folder_path = ""

        # UI Elements
        self.label = QLabel("未选择文件夹", self)
        self.label.setWordWrap(True)
        self.label.setAlignment(Qt.AlignCenter)

        self.select_button = QPushButton("选择文件夹", self)
        self.convert_button = QPushButton("转换为 PDF", self)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.select_button)
        layout.addWidget(self.convert_button)
        self.setLayout(layout)

        # Connect buttons
        self.select_button.clicked.connect(self.select_folder)
        self.convert_button.clicked.connect(self.convert_files)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.folder_path = folder
            self.label.setText(folder)

    def convert_files(self):
        if not self.folder_path:
            QMessageBox.warning(self, "错误", "请先选择一个文件夹！")
            return

        word_files = [f for f in os.listdir(self.folder_path) if f.endswith(('.doc', '.docx'))]
        ppt_files = [f for f in os.listdir(self.folder_path) if f.endswith(('.ppt', '.pptx'))]

        total_files = word_files + ppt_files
        if not total_files:
            QMessageBox.information(self, "提示", "没有找到可转换的 Word 或 PPT 文件。")
            return

        try:
            self._convert_word_files(word_files)
            self._convert_ppt_files(ppt_files)
            QMessageBox.information(self, "完成", "所有文件已成功转换为 PDF！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"转换过程中出错：{str(e)}")

    def _convert_word_files(self, files):
        word = win32com.client.Dispatch("Word.Application")
        # word.SetDisplayAlerts(False)
        # word.Visible = False
        for filename in files:
            try:
                file_path = os.path.join(self.folder_path, filename)
                doc = word.Documents.Open(file_path)
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17  # wdExportFormatPDF
                )
                doc.Close()
            except:
                continue
        word.Quit()

    def _convert_ppt_files(self, files):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # powerpoint.SetDisplayAlerts(False)
        # powerpoint.Visible = False
        for filename in files:
            try:
                file_path = os.path.join(self.folder_path, filename)
                presentation = powerpoint.Presentations.Open(file_path)
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                presentation.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    FixedFormatType=2  # ppFixedFormatTypePDF
                )
                presentation.Close()
            except:
                continue
        powerpoint.Quit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConvertToPDFApp()
    window.show()
    sys.exit(app.exec_())