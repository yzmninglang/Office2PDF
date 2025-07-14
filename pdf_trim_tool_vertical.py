import fitz
import os
import argparse
from tqdm import tqdm

def detect_content_bbox(page, threshold=0.1, trim_vertical=True):
    """检测页面内容的边界框，可选择是否裁剪垂直方向白边"""
    # 获取页面尺寸
    mediabox = page.rect
    pix = page.get_pixmap()
    
    width = pix.width
    height = pix.height
    
    # 初始化边界
    left = width
    right = 0
    top = height
    bottom = 0
    
    # 转换为灰度图像进行检测
    for y in range(height):
        for x in range(width):
            r, g, b = pix.pixel(x, y)
            # 计算灰度值
            gray = (r + g + b) // 3
            # 如果不是白色(接近255)，则认为是内容
            if gray < 255 * (1 - threshold):
                left = min(left, x)
                right = max(right, x)
                top = min(top, y)
                bottom = max(bottom, y)
    
    # 如果没有检测到内容，返回原始页面边界
    if left > right or top > bottom:
        return mediabox
    
    # 将像素坐标转换为PDF坐标
    x0 = mediabox.x0 + left * (mediabox.x1 - mediabox.x0) / width
    y0 = mediabox.y0 + top * (mediabox.y1 - mediabox.y0) / height
    x1 = mediabox.x0 + right * (mediabox.x1 - mediabox.x0) / width
    y1 = mediabox.y0 + bottom * (mediabox.y1 - mediabox.y0) / height
    
    # 添加安全边距
    safety_margin = 10  # 磅
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

def crop_pdf(input_path, output_path, threshold=0.1, margin=0, trim_vertical=True):
    """裁剪PDF文件的每一页"""
    try:
        # 打开输入PDF
        doc = fitz.open(input_path)
        
        # 创建输出PDF
        output_doc = fitz.open()
        
        # 处理每一页
        for page_num in tqdm(range(len(doc)), desc="裁剪页面"):
            page = doc.load_page(page_num)
            
            # 检测内容边界
            content_bbox = detect_content_bbox(page, threshold, trim_vertical)
            
            # 创建新页面并复制内容
            new_page = output_doc.new_page(width=content_bbox.width, 
                                         height=content_bbox.height)
            
            # 将原页面内容映射到新页面
            new_page.show_pdf_page(new_page.rect, doc, page_num, 
                                  clip=content_bbox)
        
        # 保存输出PDF
        output_doc.save(output_path)
        output_doc.close()
        doc.close()
        
        print(f"PDF裁剪完成，已保存至: {output_path}")
        return True
    
    except Exception as e:
        print(f"处理PDF时出错: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description='PDF自动裁剪工具 - 移除页面白边')
    parser.add_argument('-i', '--input', required=True, help='输入PDF文件路径')
    parser.add_argument('-o', '--output', help='输出PDF文件路径，默认为input_cropped.pdf')
    parser.add_argument('-t', '--threshold', type=float, default=0.1, 
                        help='内容检测阈值(0-1)，值越小越严格，默认为0.1')
    parser.add_argument('-m', '--margin', type=int, default=10, 
                        help='保留的边距(磅)，默认为10')
    parser.add_argument('--trim_vertical', action='store_true', 
                        help='是否裁剪上下空白部分，默认启用')
    parser.add_argument('--no-trim_vertical', action='store_false', 
                        help='禁用裁剪上下空白部分')
    parser.set_defaults(trim_vertical=False)
    
    args = parser.parse_args()
    
    # 确定输出路径
    if not args.output:
        base_name, ext = os.path.splitext(args.input)
        args.output = f"{base_name}_cropped{ext}"
    
    # 裁剪PDF
    success = crop_pdf(args.input, args.output, args.threshold, args.margin, args.trim_vertical)
    
    if success:
        # 计算压缩率
        original_size = os.path.getsize(args.input)
        cropped_size = os.path.getsize(args.output)
        reduction = (1 - cropped_size / original_size) * 100
        print(f"文件大小减少: {reduction:.2f}%")

if __name__ == "__main__":
    main()    