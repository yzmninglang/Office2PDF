import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                            QHBoxLayout, QFileDialog, QWidget, QCheckBox, QLabel, 
                            QProgressBar, QMessageBox, QGroupBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
import fitz
from tqdm import tqdm

class CropThread(QThread):
    progress_updated = pyqtSignal(int)
    task_completed = pyqtSignal(bool, str)
    
    def __init__(self, input_path, output_path, trim_vertical):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.trim_vertical = trim_vertical
        
    def detect_content_bbox(self, page, threshold=0.1, trim_vertical=True):
        """检测页面内容的边界框"""
        mediabox = page.rect
        pix = page.get_pixmap()
        
        width = pix.width
        height = pix.height
        
        left = width
        right = 0
        top = height
        bottom = 0
        
        for y in range(height):
            for x in range(width):
                r, g, b = pix.pixel(x, y)
                gray = (r + g + b) // 3
                if gray < 255 * (1 - threshold):
                    left = min(left, x)
                    right = max(right, x)
                    top = min(top, y)
                    bottom = max(bottom, y)
        
        if left > right or top > bottom:
            return mediabox
        
        x0 = mediabox.x0 + left * (mediabox.x1 - mediabox.x0) / width
        y0 = mediabox.y0 + top * (mediabox.y1 - mediabox.y0) / height
        x1 = mediabox.x0 + right * (mediabox.x1 - mediabox.x0) / width
        y1 = mediabox.y0 + bottom * (mediabox.y1 - mediabox.y0) / height
        
        safety_margin = 15  # 磅
        x0 = max(mediabox.x0, x0 - safety_margin)
        x1 = min(mediabox.x1, x1 + safety_margin)
        
        if trim_vertical:
            y0 = max(mediabox.y0, y0 - safety_margin)
            y1 = min(mediabox.y1, y1 + safety_margin)
        else:
            # 不裁剪垂直方向，使用原始页面的上下边界
            y0 = mediabox.y0
            y1 = mediabox.y1
            
        return fitz.Rect(x0, y0, x1, y1)
    
    def run(self):
        try:
            doc = fitz.open(self.input_path)
            output_doc = fitz.open()
            
            # 映射原页面到新页面的索引
            page_mapping = {}
            
            total_pages = len(doc)
            for i, page in enumerate(doc):
                content_bbox = self.detect_content_bbox(page, trim_vertical=self.trim_vertical)
                new_page = output_doc.new_page(width=content_bbox.width, height=content_bbox.height)
                new_page.show_pdf_page(new_page.rect, doc, i, clip=content_bbox)
                
                # 保存页面映射关系（原页面索引 -> 新页面索引）
                page_mapping[i] = new_page.number
                
                progress = int((i + 1) / total_pages * 100)
                self.progress_updated.emit(progress)
            
            # 复制书签并调整目标页面
            bookmarks = doc.get_toc()
            if bookmarks:
                new_bookmarks = []
                for bm in bookmarks:
                    level, title, page_num = bm[:3]
                    # 页面索引从0开始
                    if page_num - 1 in page_mapping:
                        new_page_num = page_mapping[page_num - 1] + 1  # 转换为1-based索引
                        new_bookmark = [level, title, new_page_num]
                        # 添加可选的坐标信息（如果存在）
                        if len(bm) > 3:
                            new_bookmark.extend(bm[3:])
                        new_bookmarks.append(new_bookmark)
                # 设置新文档的书签
                output_doc.set_toc(new_bookmarks)
            
            # 复制元数据
            output_doc.set_metadata(doc.metadata)
            
            output_doc.save(self.output_path)
            output_doc.close()
            doc.close()
            
            original_size = os.path.getsize(self.input_path)
            cropped_size = os.path.getsize(self.output_path)
            reduction = (1 - cropped_size / original_size) * 100
            
            self.task_completed.emit(True, f"处理完成！文件大小减少: {reduction:.2f}%")
        except Exception as e:
            self.task_completed.emit(False, f"处理失败: {str(e)}")

class PDFTrimmer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowIcon(QIcon("trim.ico"))
        self.setWindowTitle('PDF白边裁剪工具')
        self.setGeometry(300, 300, 600, 400)
        
        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 文件选择部分
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        file_group.setLayout(file_layout)
        
        self.file_path_label = QLabel("未选择文件")
        file_layout.addWidget(self.file_path_label)
        
        file_button_layout = QHBoxLayout()
        self.select_file_button = QPushButton("选择PDF文件")
        self.select_file_button.clicked.connect(self.select_file)
        file_button_layout.addWidget(self.select_file_button)
        
        file_layout.addLayout(file_button_layout)
        main_layout.addWidget(file_group)
        
        # 裁剪选项部分
        options_group = QGroupBox("裁剪选项")
        options_layout = QVBoxLayout()
        options_group.setLayout(options_layout)
        
        self.trim_vertical_checkbox = QCheckBox("裁剪上下白边")
        self.trim_vertical_checkbox.setChecked(False)
        options_layout.addWidget(self.trim_vertical_checkbox)
        
        main_layout.addWidget(options_group)
        
        # 处理按钮
        self.process_button = QPushButton("去除白边")
        self.process_button.clicked.connect(self.process_pdf)
        self.process_button.setEnabled(False)
        main_layout.addWidget(self.process_button)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        main_layout.addWidget(self.status_label)
        
        # 添加拉伸以调整布局
        main_layout.addStretch(1)
        
        self.show()
        
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PDF文件", "", "PDF Files (*.pdf);;All Files (*)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(file_path)
            self.process_button.setEnabled(True)
            
    def process_pdf(self):
        if not hasattr(self, 'file_path'):
            QMessageBox.warning(self, "警告", "请先选择PDF文件")
            return
            
        self.process_button.setEnabled(False)
        self.status_label.setText("处理中...")
        
        base_name, ext = os.path.splitext(self.file_path)
        output_path = f"{base_name}_trim{ext}"
        
        self.crop_thread = CropThread(
            self.file_path, 
            output_path, 
            self.trim_vertical_checkbox.isChecked()
        )
        
        self.crop_thread.progress_updated.connect(self.update_progress)
        self.crop_thread.task_completed.connect(self.on_task_completed)
        self.crop_thread.start()
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def on_task_completed(self, success, message):
        self.process_button.setEnabled(True)
        
        if success:
            self.status_label.setText("处理完成")
            QMessageBox.information(self, "成功", message)
        else:
            self.status_label.setText("处理失败")
            QMessageBox.critical(self, "错误", message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格，跨平台一致性更好
    window = PDFTrimmer()
    sys.exit(app.exec_())    