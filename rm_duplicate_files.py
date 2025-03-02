import os
import hashlib
import shutil
from tqdm import tqdm

def calculate_hash(file_path, block_size=65536):
    """计算文件的哈希值"""
    hasher = hashlib.sha256()
    with open(file_path, 'rb') as f:
        buf = f.read(block_size)
        while buf:
            hasher.update(buf)
            buf = f.read(block_size)
    return hasher.hexdigest()

def find_duplicate_files(directory):
    """查找重复文件"""
    hashes = {}
    total_files = sum(len(files) for _, _, files in os.walk(directory))
    with tqdm(total=total_files, desc="查找文件", unit="file") as pbar:
        for root, _, files in os.walk(directory):
            for filename in files:
                file_path = os.path.join(root, filename)
                file_hash = calculate_hash(file_path)
                if file_hash in hashes:
                    hashes[file_hash].append(file_path)
                else:
                    hashes[file_hash] = [file_path]
                pbar.update(1)
    return {k: v for k, v in hashes.items() if len(v) > 1}

def delete_duplicates(duplicates):
    """删除重复文件，只保留一份"""
    for file_group in duplicates.values():
        # 保留第一个文件，删除其余文件
        for file_path in file_group[1:]:
            os.remove(file_path)
            print(f"Deleted: {file_path}")

def main():
    directory = input("请输入要遍历的文件夹路径: ")
    if not os.path.isdir(directory):
        print("指定的路径不是一个有效的文件夹。")
        return

    print("正在查找重复文件...")
    duplicates = find_duplicate_files(directory)
    print("查找完成。")

    if not duplicates:
        print("没有找到重复文件。")
        return

    print("找到的重复文件组：")
    for file_group in duplicates.values():
        for file_path in file_group:
            print(file_path)
        print()

    confirm = input("是否删除多余的重复文件？(y/n): ")
    if confirm.lower() == 'y':
        delete_duplicates(duplicates)
        print("删除完成。")
    else:
        print("未删除任何文件。")

if __name__ == "__main__":
    main()
