import os
import subprocess
from pathlib import Path

CONVERT_TYPES = {
    1: (".doc", ".docx"),
    2: (".ppt", ".pptx"),
    3: (".xls", ".xlsx"),
    4: (".doc", ".docx"),
    5: (".ppt", ".pptx"),
    6: (".xls", ".xlsx")
}

def show_menu():
    """显示转换类型菜单"""
    print("\n请选择要转换的文件类型（输入数字，多个用逗号分隔）:")
    print("1. Word文档 (.doc → .docx)")
    print("2. PowerPoint演示 (.ppt → .pptx)")
    print("3. Excel表格 (.xls → .xlsx)")
    print("4. 所有类型")
    while True:
        choice = input("请输入选择（例如 1 或 1,2,3）: ")
        choices = [int(c.strip()) for c in choice.split(',') if c.strip().isdigit()]
        if choices:
            return choices
        print("错误: 输入无效，请重新输入。")

def get_extensions(choices):
    """根据用户选择获取扩展名列表"""
    if 4 in choices:
        return [(".doc", ".docx"), (".ppt", ".pptx"), (".xls", ".xlsx")]
    return list(set(CONVERT_TYPES.get(c, ()) for c in choices if c in CONVERT_TYPES))

def convert_file(input_file, output_ext):
    """使用 LibreOffice 转换单个文件"""
    try:
        output_dir = os.path.dirname(input_file)
        subprocess.run(
            ['soffice', '--headless', '--convert-to', output_ext[1:], '--outdir', output_dir, input_file],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE
        )
        print(f"✓ 转换成功: {Path(input_file).name} → {Path(input_file).stem}{output_ext}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ 转换失败: {Path(input_file).name}\n   错误信息: {e.stderr.decode().strip()}")
        return False
    except FileNotFoundError:
        print("错误: 未找到 LibreOffice，请确保已安装并配置环境变量")
        return False

def process_files(folder_path, extensions, process_subfolders):
    """处理文件夹中的文件"""
    folder = Path(folder_path)
    if not folder.exists():
        print(f"错误: 路径不存在 - {folder_path}")
        return

    total_files = 0
    converted_files = 0

    for root, dirs, files in os.walk(folder):
        # 如果不处理子文件夹，清空后续目录列表
        if not process_subfolders:
            dirs[:] = []

        for file in files:
            file_path = Path(root) / file
            current_ext = file_path.suffix.lower()
            for src_ext, dst_ext in extensions:
                if current_ext == src_ext:
                    total_files += 1
                    output_file = file_path.with_suffix(dst_ext)

                    if output_file.exists():
                        print(f"⚠ 跳过: {file} → 目标文件已存在")
                        continue

                    print(f"⌛ 正在转换: {file}")
                    if convert_file(str(file_path), dst_ext):
                        converted_files += 1
                    break

    print(f"\n转换完成: 共找到 {total_files} 个文件，成功转换 {converted_files} 个")

def main():
    """主程序"""
    print("="*50)
    print("Office 文档批量转换工具".center(50))
    print("="*50)
    
    # 获取文件夹路径
    while True:
        folder_path = input("\n请输入文件夹路径: ").strip()
        if Path(folder_path).exists():
            break
        print(f"错误: 路径不存在，请重新输入。")

    # 选择是否处理子文件夹
    while True:
        process_sub = input("是否处理子文件夹？(y/n): ").strip().lower()
        if process_sub in ['y', 'n']:
            process_subfolders = process_sub == 'y'
            break
        print("错误: 请输入 y 或 n。")

    # 选择转换类型
    choices = show_menu()
    extensions = get_extensions(choices)
    if not extensions:
        print("错误: 未选择有效类型")
        return

    print("\n" + "="*50)
    print(f"转换类型: {', '.join([f'{src}→{dst}' for src, dst in extensions])}")
    print(f"处理子文件夹: {'是' if process_subfolders else '否'}")
    print("="*50 + "\n")
    
    process_files(folder_path, extensions, process_subfolders)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n操作已取消")
