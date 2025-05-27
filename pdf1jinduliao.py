import os
import sys
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QFileDialog,
    QMessageBox, QListWidget, QHBoxLayout, QComboBox, QAction, QMenu, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from PyPDF2 import PdfReader, PdfWriter


class ConversionWorker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int, int, str)  # current, total, filename

    def __init__(self, folder_path, word_files, ppt_files):
        super().__init__()
        self.folder_path = folder_path
        self.word_files = word_files
        self.ppt_files = ppt_files

    def run(self):
        total = len(self.word_files) + len(self.ppt_files)
        if total == 0:
            self.finished.emit()
            return

        count = 0
        for filename in self.word_files:
            self._convert_word_single(filename)
            count += 1
            self.progress.emit(count, total, filename)

        for filename in self.ppt_files:
            self._convert_ppt_single(filename)
            count += 1
            self.progress.emit(count, total, filename)

        self.finished.emit()

    def _convert_word_single(self, filename):
        try:
            file_path = os.path.join(self.folder_path, filename)
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(file_path)
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17  # wdExportFormatPDF
            )
            doc.Close()
            word.Quit()
        except Exception as e:
            print(f"Word 转换失败: {filename}, 错误: {e}")

    def _convert_ppt_single(self, filename):
        try:
            file_path = os.path.join(self.folder_path, filename)
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(file_path)
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            presentation.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                FixedFormatType=2  # ppFixedFormatTypePDF
            )
            presentation.Close()
            powerpoint.Quit()
        except Exception as e:
            print(f"PPT 转换失败: {filename}, 错误: {e}")


class OfficeToPDFConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("办公文档转PDF & PDF合并工具")
        self.resize(700, 650)

        # --- Word/PPT 转换部分 ---
        self.label_folder = QLabel("未选择文件夹", self)
        self.label_folder.setWordWrap(True)
        self.label_folder.setAlignment(Qt.AlignCenter)

        self.btn_select_folder = QPushButton("选择文件夹（Word/PPT）")
        self.btn_convert = QPushButton("转换为 PDF")

        self.current_file_label = QLabel("当前文件：无")
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        convert_layout = QVBoxLayout()
        convert_layout.addWidget(QLabel("【Word & PPT 批量转 PDF】"))
        convert_layout.addWidget(self.label_folder)
        convert_layout.addWidget(self.btn_select_folder)
        convert_layout.addWidget(self.btn_convert)
        convert_layout.addWidget(self.current_file_label)
        convert_layout.addWidget(self.progress_bar)

        # --- PDF 合并部分 ---
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.setDragDropMode(QListWidget.InternalMove)
        self.file_list.setAcceptDrops(True)
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_context_menu)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["按名称排序", "按创建时间排序"])

        self.btn_add_pdf = QPushButton("添加 PDF 文件")
        self.btn_merge_pdf = QPushButton("合并 PDF")

        merge_layout = QVBoxLayout()
        merge_layout.addWidget(QLabel("【PDF 合并器】"))
        merge_layout.addWidget(QLabel("选择排序方式："))
        merge_layout.addWidget(self.sort_combo)
        merge_layout.addWidget(QLabel("拖动可调整顺序（右键可移除）："))
        merge_layout.addWidget(self.file_list)
        merge_layout.addWidget(self.btn_add_pdf)
        merge_layout.addWidget(self.btn_merge_pdf)

        # --- 快捷键支持 ---
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        QShortcut(QKeySequence("Ctrl+A"), self, self.select_all_items)
        QShortcut(QKeySequence("Delete"), self, self.remove_selected_items)

        # --- 主布局 ---
        main_layout = QVBoxLayout()
        main_layout.addLayout(convert_layout)
        main_layout.addSpacing(20)
        main_layout.addLayout(merge_layout)

        self.setLayout(main_layout)

        # --- 绑定事件 ---
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.btn_convert.clicked.connect(self.convert_files)

        self.btn_add_pdf.clicked.connect(self.add_pdfs)
        self.btn_merge_pdf.clicked.connect(self.merge_pdfs)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.folder_path = folder
            self.label_folder.setText(folder)

    def convert_files(self):
        if not hasattr(self, 'folder_path'):
            QMessageBox.warning(self, "错误", "请先选择一个文件夹！")
            return

        word_files = [f for f in os.listdir(self.folder_path) if f.endswith(('.doc', '.docx'))]
        ppt_files = [f for f in os.listdir(self.folder_path) if f.endswith(('.ppt', '.pptx'))]

        total_files = word_files + ppt_files
        if not total_files:
            QMessageBox.information(self, "提示", "没有找到可转换的 Word 或 PPT 文件。")
            return

        self.btn_convert.setEnabled(False)
        self.progress_bar.setRange(0, len(total_files))
        self.progress_bar.setValue(0)
        self.current_file_label.setText("当前文件：开始转换...")

        # 创建线程和 worker
        self.worker = ConversionWorker(self.folder_path, word_files, ppt_files)
        self.thread = QThread()
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def update_progress(self, current, total, filename):
        self.progress_bar.setValue(current)
        self.current_file_label.setText(f"当前文件：{filename}")

    def conversion_finished(self):
        self.btn_convert.setEnabled(True)
        self.progress_bar.setValue(0)
        self.current_file_label.setText("当前文件：无")
        QMessageBox.information(self, "完成", "所有文件已成功转换为 PDF！")

    def show_context_menu(self, position):
        menu = QMenu(self)
        remove_action = QAction("移除选中文件", self)
        remove_action.triggered.connect(self.remove_selected_items)
        menu.addAction(remove_action)
        menu.exec_(self.file_list.mapToGlobal(position))

    def remove_selected_items(self):
        selected_items = self.file_list.selectedItems()
        for item in selected_items:
            self.file_list.takeItem(self.file_list.row(item))

    def add_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择 PDF 文件", "", "PDF 文件 (*.pdf)")
        if not files:
            return

        sort_mode = self.sort_combo.currentText()

        if sort_mode == "按名称排序":
            sorted_files = sorted(files, key=lambda x: os.path.basename(x).lower())
        elif sort_mode == "按创建时间排序":
            sorted_files = sorted(files, key=lambda x: os.path.getctime(x))
        else:
            sorted_files = files

        self.file_list.clear()
        for file in sorted_files:
            item = QListWidgetItem()
            item.setText(os.path.basename(file))
            item.setData(Qt.UserRole, file)
            item.setToolTip(file)
            self.file_list.addItem(item)

    def merge_pdfs(self):
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "错误", "请先添加至少一个 PDF 文件。")
            return

        output_path, _ = QFileDialog.getSaveFileName(self, "保存合并后的 PDF", "", "PDF 文件 (*.pdf)")
        if not output_path:
            return

        try:
            pdf_writer = PdfWriter()
            current_page = 0
            bookmarks = []

            for i in range(self.file_list.count()):
                item = self.file_list.item(i)
                pdf_path = item.data(Qt.UserRole)
                reader = PdfReader(pdf_path)

                outlines = reader.outline if reader.outline else []
                bookmarks.append({
                    'filename': item.text(),
                    'pages': len(reader.pages),
                    'outline': outlines
                })

                for page in reader.pages:
                    pdf_writer.add_page(page)

                bookmarks[-1]['start_page'] = current_page
                current_page += len(reader.pages)

            current_page = 0
            for bm in bookmarks:
                filename = bm['filename']
                start_page = bm['start_page']

                top_level = pdf_writer.add_outline_item(filename, start_page)
                self.add_outline(pdf_writer, bm['outline'], parent=top_level, offset=current_page)
                current_page += bm['pages']

            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)

            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle("完成")
            msg_box.setText(f"PDF 已成功合并到：\n{output_path}")
            open_button = msg_box.addButton("打开文件所在位置", QMessageBox.ActionRole)
            ok_button = msg_box.addButton(QMessageBox.Ok)
            msg_box.exec_()

            if msg_box.clickedButton() == open_button:
                self.open_file_location(output_path)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并过程中出错：{str(e)}")

    def add_outline(self, writer, outline, parent=None, offset=0):
        if isinstance(outline, list):
            for item in outline:
                self.add_outline(writer, item, parent, offset)
        elif hasattr(outline, 'dest'):
            dest = outline.dest
            page_number = writer.get_destination_page_number(dest) + offset
            writer.add_outline_item(outline.title, page_number, parent=parent)
        elif hasattr(outline, '__dict__'):
            title = getattr(outline, 'title', 'Untitled')
            try:
                page_ref = getattr(outline, 'page_reference', None)
                if page_ref:
                    page_number = writer.page_references.index(page_ref) + offset
                    writer.add_outline_item(title, page_number, parent=parent)
                else:
                    pass
            except Exception:
                pass

    def open_file_location(self, file_path):
        directory = os.path.dirname(file_path)
        try:
            if sys.platform == 'win32':
                os.startfile(f'explorer /select,"{file_path}"')
            elif sys.platform == 'darwin':
                os.system(f'open -R "{file_path}"')
            else:
                os.system(f'xdg-open "{directory}"')
        except Exception as e:
            QMessageBox.warning(self, "警告", f"无法打开文件位置：{e}")

    def select_all_items(self):
        self.file_list.selectAll()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OfficeToPDFConverter()
    window.show()
    sys.exit(app.exec_())