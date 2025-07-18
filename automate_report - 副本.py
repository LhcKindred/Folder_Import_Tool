import os
from pathlib import Path
from docx import Document

# --- 配置变量 ---
# 定义影像文件夹所在的根目录。
INPUT_BASE_DIR = Path("input_images")
# 指定现有Word文档模板的路径。
TEMPLATE_DOCX_PATH = "介休源神庙内业记录表.docx"
# 定义要创建的新更新Word文档的名称。
OUTPUT_DOCX_PATH = "介休源神庙内业记录表_updated.docx"
# 常见影像文件扩展名列表，用于过滤文件。确保一致性。
IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.cr2', '.nef', '.dng')

def process_folders_and_update_word(input_base_dir, template_path, output_path):
    """
    处理指定目录下的所有影像文件夹，提取所需数据，并更新Word文档表格。

    Args:
        input_base_dir (Path): 包含影像文件夹的根目录路径。
        template_path (str): Word文档模板的路径。
        output_path (str): 生成的更新Word文档的保存路径。
    """
    folder_data = []

    # --- 阶段1：从影像文件夹中提取数据 ---
    print(f"开始从 {input_base_dir} 提取数据...")
    for folder_path in input_base_dir.iterdir():
        if folder_path.is_dir():
            print(f"处理文件夹: {folder_path.name}")
            
            image_files = []
            for item in folder_path.iterdir():
                if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS:
                    image_files.append(item.name)

            if not image_files:
                print(f"  警告: 在 '{folder_path.name}' 中未找到有效影像文件。跳过此文件夹。")
                continue

            sorted_image_files = sorted(image_files)

            first_file = sorted_image_files[0]

            second_file = ""
            last_file = ""

            if len(sorted_image_files) >= 2:
                second_file = sorted_image_files[1] # 修正：使用索引 [1] 获取第二个文件
                last_file = sorted_image_files[-1]
            else:
                print(f"  警告: '{folder_path.name}' 中影像文件少于2个。无法确定第二个/最后一个文件进行连接。")
            
            second_to_last_combined = f"{second_file}-{last_file}" if second_file and last_file else ""

            file_count_second_to_last = max(0, len(sorted_image_files) - 1)

            total_folder_size_bytes = 0
            for item in folder_path.iterdir():
                if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS:
                    total_folder_size_bytes += item.stat().st_size

            total_folder_size_gb = total_folder_size_bytes / (1024**3)

            folder_data.append({
                "first_file": first_file,
                "second_to_last_combined": second_to_last_combined,
                "file_count": file_count_second_to_last,
                "folder_size_gb": total_folder_size_gb,
                "processor_name": "王烁宁" # 示例值，可根据需要修改或从其他来源获取
            })
        else:
            print(f"跳过非目录项: {folder_path.name}")

    if not folder_data:
        print("未收集到任何文件夹数据。请检查输入目录和文件。")
        return

    # --- 阶段2：填充Word文档表格 ---
    print(f"\n加载模板文档: {template_path}...")
    try:
        document = Document(template_path)
    except Exception as e:
        print(f"错误: 无法加载Word文档模板 '{template_path}'。请确保文件存在且格式正确。错误信息: {e}")
        return

    target_table = None
    header_text_marker = "色卡照片编号"

    for table in document.tables:
        # 检查表格是否至少有一行和第一行中的一个单元格。
        # 修正：确保在访问 cells 之前检查 rows 是否存在且有元素
        if len(table.rows) > 0:
            if any(header_text_marker in cell.text for cell in table.rows[0].cells):
                target_table = table
                break


    if not target_table:
        print(f"错误: 在文档中找不到带有标题 '{header_text_marker}' 的目标表格。请确保模板正确。")
        return

    print("已找到目标表格。开始填充数据...")
    for data_entry in folder_data:
        row = target_table.add_row()
        row_cells = row.cells
        row_cells[0].text = data_entry["first_file"]
        row_cells[1].text = data_entry["second_to_last_combined"]
        row_cells[2].text = str(data_entry["file_count"])
        row_cells[3].text = f"{data_entry['folder_size_gb']:.2f} GB"
        if len(row_cells) > 7:
            row_cells[7].text = data_entry.get("processor_name", "")



    # --- 保存更新的文档 ---
    try:
        document.save(output_path)
        print(f"\n报告已成功生成: {output_path}")
    except Exception as e:
        print(f"错误: 无法保存更新的Word文档 '{output_path}'。请检查文件权限或路径。错误信息: {e}")

if __name__ == "__main__":
    # 确保输入目录存在
    if not INPUT_BASE_DIR.exists():
        print(f"错误: 输入目录 '{INPUT_BASE_DIR}' 不存在。请创建该目录并将影像文件夹放入其中。")
    elif not TEMPLATE_DOCX_PATH:
        print(f"错误: 模板文档路径未指定。请在脚本中设置 TEMPLATE_DOCX_PATH。")
    else:
        process_folders_and_update_word(INPUT_BASE_DIR, TEMPLATE_DOCX_PATH, OUTPUT_DOCX_PATH)