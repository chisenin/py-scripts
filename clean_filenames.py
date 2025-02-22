import os
import re
import sys
import io
import csv
from datetime import datetime

# 设置控制台编码为UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def clean_filename(name, bad_strings):
    """清理文件名并规范化空格"""
    cleaned = name
    for s in bad_strings:
        cleaned = cleaned.replace(s.strip(), '')
    # 清理连续空格并去除首尾空格
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned

def handle_rename(src_path, new_name):
    """处理文件重命名冲突"""
    if not new_name:
        new_name = "cleaned_file"
    
    dst_path = os.path.join(os.path.dirname(src_path), new_name)
    counter = 1
    base, ext = os.path.splitext(new_name)
    
    while os.path.exists(dst_path):
        new_name = f"{base}_{counter}{ext}"
        dst_path = os.path.join(os.path.dirname(src_path), new_name)
        counter += 1
    
    os.rename(src_path, dst_path)
    return new_name

def process_directory(folder_path, bad_strings):
    """递归处理目录并记录日志"""
    processed = 0
    current = 0
    total = 0
    log_path = os.path.join(folder_path, "rename_log.csv")
    
    # 预计算总数用于进度显示
    for root, dirs, files in os.walk(folder_path):
        total += len(files) + len(dirs)
    
    # 初始化日志文件
    with open(log_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["原路径", "新路径", "修改时间"])

    for root, dirs, files in os.walk(folder_path, topdown=False):
        # 处理文件
        for filename in files:
            current += 1
            if current % 50 == 0 or current == total:
                print(f"处理进度：{current}/{total} ({current/total:.1%})")
                
            original_path = os.path.join(root, filename)
            cleaned_name = clean_filename(filename, bad_strings)
            
            if cleaned_name != filename:
                new_name = handle_rename(original_path, cleaned_name)
                new_path = os.path.join(os.path.dirname(original_path), new_name)
                
                # 记录日志
                with open(log_path, 'a', encoding='utf-8', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        original_path,
                        new_path,
                        datetime.now().isoformat()
                    ])
                print(f"文件重命名: {filename} → {new_name}")
                processed += 1

        # 处理目录
        for dirname in dirs:
            current += 1
            if current % 50 == 0 or current == total:
                print(f"处理进度：{current}/{total} ({current/total:.1%})")
                
            original_path = os.path.join(root, dirname)
            cleaned_name = clean_filename(dirname, bad_strings)
            
            if cleaned_name != dirname:
                new_name = handle_rename(original_path, cleaned_name)
                new_path = os.path.join(os.path.dirname(original_path), new_name)
                
                with open(log_path, 'a', encoding='utf-8', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        original_path,
                        new_path,
                        datetime.now().isoformat()
                    ])
                print(f"目录重命名: {dirname} → {new_name}")
                processed += 1
                
    return processed

def get_target_folder():
    """获取并验证目标文件夹路径"""
    while True:
        path = input("请输入或拖拽文件夹路径：").strip('" ')
        if os.path.isdir(path):
            return os.path.abspath(path)
        print(f"错误：路径 '{path}' 不存在或不是目录")

def get_bad_strings():
    """交互式获取清理字符串"""
    print("\n请输入要清理的字符串（逐个输入，回车结束）：")
    bad_strings = []
    while True:
        s = input(f"第 {len(bad_strings)+1} 个字符串：").strip()
        if not s:
            if len(bad_strings) == 0:
                print("至少需要输入一个清理字符串！")
                continue
            return bad_strings
        bad_strings.append(s)

def main():
    print("=== 智能文件清理工具 ===")
    print("功能特性：")
    print("- 支持中英文路径")
    print("- 自动处理重名冲突")
    print("- 实时进度显示")
    print("- 详细操作日志\n")
    
    target_folder = get_target_folder()
    bad_strings = get_bad_strings()
    
    print("\n配置摘要：")
    print(f"目标目录：{target_folder}")
    print("清理字符串列表：")
    for i, s in enumerate(bad_strings, 1):
        print(f" {i}. {s}")
    
    confirm = input("\n确认开始清理？(y/n) ").lower()
    if confirm != 'y':
        print("操作已取消")
        return
    
    try:
        total = process_directory(target_folder, bad_strings)
        print(f"\n操作完成，共清理 {total} 个项目")
        if total > 0:
            print(f"日志文件已保存至：{os.path.join(target_folder, 'rename_log.csv')}")
            print("建议使用Excel打开日志文件进行复核")
    except Exception as e:
        print(f"\n错误发生：{str(e)}")
        print("排查建议：")
        print("1. 检查文件是否被其他程序占用")
        print("2. 确认有足够的权限")
        print("3. 尝试缩短路径长度")

if __name__ == "__main__":
    main()
