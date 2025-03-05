import os
import re
import sys
from PyPDF2 import PdfReader
import fitz  # PyMuPDF

sys.setrecursionlimit(1000000)  # 防止复杂PDF解析崩溃
def extract_title(filepath):
    # 尝试从PDF元数据中获取标题
    try:
        with open(filepath, 'rb') as f:
            reader = PdfReader(f)
            metadata = reader.metadata
            title = metadata.get('/Title', '')
            if title and title.strip():
                if not title.strip().isdigit():
                    return title.strip()
    except Exception as e:
        print(f"读取 {filepath} 元数据时出错：{e}")
    
    # 使用PyMuPDF分析页面内容提取标题
    try:
        doc = fitz.open(filepath)
        if len(doc) == 0:
            return None
        first_page = doc[0]
        blocks = first_page.get_text("dict")["blocks"]
        max_font_size = 0
        title_candidate = ""
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        current_font_size = span["size"]
                        current_text = span["text"].strip()
                        if current_text:
                            if current_font_size > max_font_size:
                                max_font_size = current_font_size
                                title_candidate = current_text
                            elif current_font_size == max_font_size:
                                title_candidate += " " + current_text
        if title_candidate:
            return title_candidate
        
        # 如果未找到，提取第一页的第一行非空文本
        text = first_page.get_text()
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        return lines[0] if lines else None
    except Exception as e:
        print(f"处理 {filepath} 时出错：{e}")
        return None
    finally:
        if 'doc' in locals():
            doc.close()

def sanitize_filename(title):
    illegal_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(illegal_chars, '_', title)
    return sanitized[:255]

def main():
    folder_path = sys.argv[1] if len(sys.argv) > 1 else input("请输入文件夹路径：")
    
    if not os.path.isdir(folder_path):
        print("错误：路径无效或不是文件夹。")
        return
    
    for filename in os.listdir(folder_path):
        if re.match(r'^\d+\.pdf$', filename):
            filepath = os.path.join(folder_path, filename)
            print(f"处理文件：{filename}")
            try:
                title = extract_title(filepath)
                if not title:
                    print(f"无法提取标题：{filename}")
                    continue
                clean_title = sanitize_filename(title)
                new_filename = f"{clean_title}.pdf"
                new_path = os.path.join(folder_path, new_filename)
                
                # 处理重复文件名
                if os.path.exists(new_path):
                    base, ext = os.path.splitext(new_filename)
                    counter = 1
                    while os.path.exists(new_path):
                        new_filename = f"{base}_{counter}{ext}"
                        new_path = os.path.join(folder_path, new_filename)
                        counter += 1
                
                os.rename(filepath, new_path)
                print(f"重命名成功：{filename} -> {new_filename}")
            except Exception as e:
                print(f"处理 {filename} 时出错：{e}")

if __name__ == "__main__":
    main()
