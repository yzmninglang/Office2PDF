import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QListWidget, QVBoxLayout,
    QFileDialog, QHBoxLayout, QMessageBox, QLabel, QComboBox,
    QAction, QMenu,QListWidgetItem
)
from PyQt5.QtCore import Qt
from PyPDF2 import PdfReader, PdfWriter

from PyQt5.QtWidgets import QShortcut
from PyQt5.QtGui import QKeySequence


class PDFMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 合并工具（支持书签）")
        self.resize(600, 400)

        # UI 元素
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.setDragDropMode(QListWidget.InternalMove)
        self.file_list.setSelectionBehavior(QListWidget.SelectRows)
        self.file_list.setAcceptDrops(True)
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_context_menu)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["按名称排序", "按创建时间排序"])

        self.add_button = QPushButton("添加 PDF 文件")
        self.merge_button = QPushButton("合并 PDF")

        # 布局
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.merge_button)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("选择排序方式："))
        layout.addWidget(self.sort_combo)
        layout.addWidget(QLabel("拖动可调整顺序（右键可移除）："))
        layout.addWidget(self.file_list)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        # 事件绑定
        self.add_button.clicked.connect(self.add_pdfs)
        self.merge_button.clicked.connect(self.merge_pdfs)


        # 添加 Ctrl+A 快捷键：全选
        self.shortcut_select_all = QShortcut(QKeySequence("Ctrl+A"), self)
        self.shortcut_select_all.activated.connect(self.select_all_items)

        # 可选：添加 Delete 键删除
        self.shortcut_delete = QShortcut(QKeySequence("Delete"), self)
        self.shortcut_delete.activated.connect(self.remove_selected_items)

    def select_all_items(self):
        self.file_list.selectAll()


    def open_file_location(self, file_path):
        """ 打开文件所在目录并选中该文件 """
        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)

        try:
            if sys.platform == 'win32':
                # Windows: explorer /select,"文件路径"
                os.startfile(f'{file_path.rsplit("/", 1)[0]}')
            elif sys.platform == 'darwin':
                # macOS: open -R 文件路径
                os.system(f'open -R "{file_path}"')
            else:
                # Linux: xdg-open 文件夹路径
                os.system(f'xdg-open "{directory}"')
        except Exception as e:
            QMessageBox.warning(self, "警告", f"无法打开文件位置：{e}")

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

        # 排序逻辑
        if sort_mode == "按名称排序":
            sorted_files = sorted(files, key=lambda x: os.path.basename(x).lower())
        elif sort_mode == "按创建时间排序":
            sorted_files = sorted(files, key=lambda x: os.path.getctime(x))
        else:
            sorted_files = files

        # 清空并添加
        self.file_list.clear()
        for file in sorted_files:
            item = QListWidgetItem()
            item.setText(os.path.basename(file))      # 显示文件名
            item.setData(Qt.UserRole, file)           # 存储完整路径
            item.setToolTip(file)                     # 鼠标悬停显示完整路径
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

                # 获取书签
                outlines = reader.outline if reader.outline else []
                bookmarks.append({
                    'filename': item.text(),
                    'pages': len(reader.pages),
                    'outline': outlines
                })

                # 添加页面
                for page in reader.pages:
                    pdf_writer.add_page(page)

                # 记录当前插入位置
                bookmarks[-1]['start_page'] = current_page
                current_page += len(reader.pages)

            # 写入页面后再添加书签
            current_page = 0
            for bm in bookmarks:
                filename = bm['filename']
                start_page = bm['start_page']

                # 添加主书签（文件名）
                top_level = pdf_writer.add_outline_item(filename, start_page)

                # 添加原书签
                self.add_outline(pdf_writer, bm['outline'], parent=top_level, offset=current_page)

                current_page += bm['pages']

            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)

                    # 显示完成消息，并添加打开文件夹按钮
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle("完成")
            msg_box.setText(f"PDF 已成功合并到：\n{output_path}")
            # QMessageBox.information(self, "完成", f"PDF 已成功合并并添加书签到：\n{output_path}")
            # 添加按钮
            open_folder_button = msg_box.addButton("打开文件所在位置", QMessageBox.ActionRole)
            ok_button = msg_box.addButton(QMessageBox.Ok)

            msg_box.exec_()

            if msg_box.clickedButton() == open_folder_button:
                self.open_file_location(output_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并过程中出错：{str(e)}")

    def add_outline(self, writer, outline, parent=None, offset=0):
        """递归添加书签"""
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFMergerApp()
    window.show()
    sys.exit(app.exec_())