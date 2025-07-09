import argparse
import os
from decimal import Decimal
from PyPDF2 import PdfReader, PdfWriter, PageObject
from PyPDF2.generic import DictionaryObject, ArrayObject, NameObject, NumberObject, TextStringObject


import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                            QHBoxLayout, QFileDialog, QLabel, QWidget, QMessageBox)

def main():
    parser = argparse.ArgumentParser(description='处理PDF文件，确保所有页面宽度一致')
    parser.add_argument('input_pdf', help='输入PDF文件路径')
    parser.add_argument('output_pdf', help='输出PDF文件路径')
    args = parser.parse_args()

    if not os.path.exists(args.input_pdf):
        print(f"错误：输入文件 '{args.input_pdf}' 不存在")
        return

    try:
        process_pdf(args.input_pdf, args.output_pdf)
        print(f"处理完成，输出文件: {args.output_pdf}")
    except Exception as e:
        print(f"处理PDF时出错: {e}")




class PDFProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # 设置窗口标题和大小
        self.setWindowTitle('PDF页面宽度统一工具')
        self.setGeometry(300, 300, 600, 200)
        
        # 创建主布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        # main_layout.setContentsMargins(20, 20, 20, 20)
        # 文件选择区域
        file_layout = QHBoxLayout()
        self.file_label = QLabel('未选择文件')
        self.file_label.setStyleSheet("color: #333; font-size: 26px;")
        select_button = QPushButton('选择PDF文件')
        select_button.setStyleSheet("background-color: #2196F3; color: white; font-size: 16px;")
        select_button.setMinimumHeight(40)
        select_button.clicked.connect(self.select_file)
        
        file_layout.addWidget(select_button)
        file_layout.addWidget(self.file_label)
        main_layout.addLayout(file_layout)
        
        # 处理按钮
        process_button = QPushButton('Crop')
        process_button.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px;")
        process_button.setMinimumHeight(40)
        process_button.clicked.connect(self.process_pdf)
        main_layout.addWidget(process_button)
        
        # 状态标签
        self.status_label = QLabel('就绪')
        self.status_label.setStyleSheet("color: #555; font-size: 12px;")
        main_layout.addWidget(self.status_label)
        
        # 记录选择的文件路径
        self.selected_file = None
        
    def select_file(self):
        """打开文件选择对话框"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PDF文件", "", "PDF Files (*.pdf);;All Files (*)"
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.status_label.setText('就绪')
            
    def process_pdf(self):
        """处理PDF文件"""
        if not self.selected_file:
            QMessageBox.warning(self, "警告", "请先选择一个PDF文件")
            return
            
        try:
            # 获取输出文件路径（原文件名+_Crop.pdf）
            base_dir = os.path.dirname(self.selected_file)
            base_name = os.path.basename(self.selected_file)
            file_name, ext = os.path.splitext(base_name)
            output_file = os.path.join(base_dir, f"{file_name}_Crop{ext}")
            
            # 更新状态
            self.status_label.setText(f"正在处理: {base_name}...")
            QApplication.processEvents()  # 刷新界面
            
            # 调用处理函数（需确保此函数在作用域内）
            process_pdf(self.selected_file, output_file)
            
            # 显示成功消息
            self.status_label.setText(f"处理完成: {os.path.basename(output_file)}")
            QMessageBox.information(self, "成功", f"PDF处理完成!\n输出文件: {output_file}")
            
        except Exception as e:
            self.status_label.setText("处理失败")
            QMessageBox.critical(self, "错误", f"处理PDF时出错:\n{str(e)}")




def process_pdf(input_path, output_path):
    with open(input_path, 'rb') as input_file:
        reader = PdfReader(input_file)
        writer = PdfWriter()

        # 获取第一页的宽度作为目标宽度
        if not reader.pages:
            print("PDF文件为空，没有页面可处理")
            return

        first_page = reader.pages[0]
        # 确保使用float类型进行计算
        target_width = float(first_page.mediabox[2] - first_page.mediabox[0])
        print(f"目标宽度: {first_page.mediabox}")
        first_page_height = float(first_page.mediabox[3] - first_page.mediabox[0])

        # 处理每一页
        for i, page in enumerate(reader.pages):
            # 确保使用float类型进行计算
            current_width = float(page.mediabox[2] - page.mediabox[0])
            current_height = float(page.mediabox[3] - page.mediabox[0])

            if i == 0:
                # 第一页保持原样
                writer.add_page(page)
            else:
                if abs(current_width - target_width) > 0.1:  # 考虑到浮点数精度问题
                    # 需要缩放页面
                    scale_factor = target_width / current_width
                    print(f"缩放页面 {i + 1}：当前宽度 {current_width}, 目标宽度 {target_width}, 缩放因子 {scale_factor}")
                    new_height = current_height * scale_factor
                    # 创建新的页面对象
                    new_page = page
                    new_page.scale(scale_factor, scale_factor)
                    writer.add_page(new_page)
                else:
                    # 宽度匹配，直接添加
                    writer.add_page(page)

        # 复制书签
        if reader.outline:
            # print(f"复制书签...{reader.outline}个书签")
            copy_bookmarks(reader, writer, reader.outline)


        # 写入输出文件
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)

def copy_bookmarks(reader, writer, outlines, parent=None):
    """递归复制书签，保持原有结构"""
    for item in outlines:
        if isinstance(item, dict):  # 书签项
            # 创建新书签
            title = item.get('/Title', '')
            if not title:
                continue
                
            # 获取目标页面引用
            page_ref = item.get('/Page')
            if not page_ref:
                continue
                
            try:
                # 直接从IndirectObject中获取页面编号
                page_id = page_ref.idnum
                page_index = next(
                    (i for i, page in enumerate(reader.pages) 
                     if page.indirect_reference.idnum == page_id), 
                    None
                )
                
                if page_index is not None and 0 <= page_index < len(writer.pages):
                    print(f"复制书签: {title}, 页面索引: {page_index}")
                    
                    # 创建书签并设置页面
                    new_bookmark = writer.add_outline_item(
                        title,
                        page_index,
                        parent=parent,
                        
                    )
                    # 复制其他属性
            except Exception as e:
                print(f"复制书签时出错: {e}")
                continue
                
        elif isinstance(item, list):  # 子书签列表
            # 递归处理子书签
            copy_bookmarks(reader, writer, item, parent)

if __name__ == '__main__':
    # 确保中文显示正常
    os.environ["QT_FONT_DPI"] = "96"
    
    app = QApplication(sys.argv)
    window = PDFProcessorGUI()
    window.show()
    sys.exit(app.exec_())    