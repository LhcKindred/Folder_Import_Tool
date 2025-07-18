import os
import re
from pathlib import Path
from docx import Document

# --- 配置变量 ---
INPUT_BASE_DIR = Path("input_images")
TEMPLATE_DOCX_PATH = "模板.docx"
OUTPUT_DOCX_PATH = "生成.docx"
IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.cr2', '.nef', '.dng')

def process_folders_and_update_word(input_base_dir, template_path, output_path):
    folder_data = []

    # 阶段1：提取数据
    print(f"开始从 {input_base_dir} 提取数据...")
    for folder_path in input_base_dir.iterdir():
        if folder_path.is_dir():
            print(f"处理文件夹: {folder_path.name}")

            image_files = [item.name for item in folder_path.iterdir()
                           if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS]

            if not image_files:
                print(f"  警告: 在 '{folder_path.name}' 中未找到有效影像文件。跳过此文件夹。")
                continue

            def natural_sort_key(s):
                # 把字符串分割成数字和非数字部分组成的列表，数字转成int
                return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

            # 对影像文件进行自然排序    
            sorted_image_files = sorted(image_files, key=natural_sort_key)

            first_file = Path(sorted_image_files[0]).stem
            second_file = ""
            last_file = ""

            if len(sorted_image_files) >= 2:
                second_file = Path(sorted_image_files[1]).stem
                last_file = Path(sorted_image_files[-1]).stem
            else:
                print(f"  警告: '{folder_path.name}' 中影像文件少于2个。无法确定第二个/最后一个文件进行连接。")

            second_to_last_combined = f"{second_file}-{last_file}" if second_file and last_file else ""
            file_count_second_to_last = max(0, len(sorted_image_files) - 1)

            total_folder_size_bytes = sum(
                item.stat().st_size for item in folder_path.iterdir()
                if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS
            )
            total_folder_size_gb = total_folder_size_bytes / (1024 ** 3)

            folder_data.append({
                "first_file": first_file,
                "second_to_last_combined": second_to_last_combined,
                "file_count": file_count_second_to_last,
                "folder_size_gb": total_folder_size_gb,
                "processor_name": "愣头青"
            })
        else:
            print(f"跳过非目录项: {folder_path.name}")

    if not folder_data:
        print("未收集到任何文件夹数据。请检查输入目录和文件。")
        return

    # 阶段2：填充Word表格
    print(f"\n加载模板文档: {template_path}...")
    try:
        document = Document(template_path)
    except Exception as e:
        print(f"错误: 无法加载Word文档模板 '{template_path}'。错误信息: {e}")
        return

    target_table = None
    if len(document.tables) == 0:
        print("错误: 文档中没有找到任何表格。")
        return
    target_table = document.tables[0]

    # 检查表格行数是否大于等于8，如果不够，添加空行
    start_row_idx = 7
    if len(target_table.rows) <= start_row_idx:
        for _ in range(start_row_idx + 1 - len(target_table.rows)):
            target_table.add_row()

    print("开始填充数据，从表格第8行开始（索引7）...")

    # 获取第七行格式
    row_cells = target_table.rows[start_row_idx].cells
    row_format = [cell for cell in row_cells]

    # 填充数据并复制格式
    for i, data_entry in enumerate(folder_data):
        row_idx = start_row_idx + i
        if row_idx < len(target_table.rows):
            row_cells = target_table.rows[row_idx].cells
        else:
            row_cells = target_table.add_row().cells

        # 保持单元格格式
        copy_row_format(row_format, row_cells)

        row_cells[0].text = data_entry["first_file"]
        row_cells[1].text = data_entry["second_to_last_combined"]
        row_cells[2].text = str(data_entry["file_count"])
        row_cells[3].text = f"{data_entry['folder_size_gb']:.2f} GB"
        row_cells[4].text = data_entry.get("processor_name", "")

    print(f"准备保存文件: {output_path}")

    # 保存文件并检查错误
    try:
        document.save(output_path)
        print(f"\n报告已成功生成: {output_path}")
    except Exception as e:
        print(f"错误: 无法保存Word文档 '{output_path}'。错误信息: {e}")

def copy_row_format(src_row, target_row):
    for src_cell, target_cell in zip(src_row, target_row):
        if is_merged(src_cell):
            target_cell._element.getparent().append(src_cell._element)

def is_merged(cell):
    try:
        return cell._tc != cell._element.getparent()
    except Exception:
        return False

if __name__ == "__main__":
    if not INPUT_BASE_DIR.exists():
        print(f"错误: 输入目录 '{INPUT_BASE_DIR}' 不存在。请创建目录并放入影像文件夹。")
    elif not TEMPLATE_DOCX_PATH:
        print("错误: 模板文档路径未指定。")
    else:
        process_folders_and_update_word(INPUT_BASE_DIR, TEMPLATE_DOCX_PATH, OUTPUT_DOCX_PATH)
