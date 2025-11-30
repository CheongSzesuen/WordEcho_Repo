import xlrd
import json
import re

def extract_data_from_xls(filename):
    """
    从 xls 文件中提取数据
    """
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)  # 假设数据在第一个 sheet
    data = []
    # 假设表头在第一行，数据从第二行开始 (索引 1)
    for row_idx in range(1, sheet.nrows):
        row = sheet.row_values(row_idx)
        # 假设列顺序为：课本、单元、英文、中文、音标
        if len(row) >= 5:
            raw_book = row[0]
            raw_unit = row[1] # 单元号
            raw_word = row[2]
            raw_trans = row[3]
            raw_symbol = row[4]
            
            data.append((raw_book, raw_unit, raw_word, raw_trans, raw_symbol))
        else:
            print(f"警告：第 {row_idx + 1} 行数据不足5列，已跳过: {row}")
    return data

def process_pep_s_data(data):
    """
    处理提取的数据，生成 JSON 对象列表
    按 课本-单元 分组，为每组内的单词编号
    """
    json_objects = []
    # 使用字典来跟踪每个 课本-单元 组合的序号
    unit_counters = {}

    for item_data in data: # 修复了这里
        raw_book, raw_unit, raw_word, raw_trans, raw_symbol = item_data

        # 处理课本
        book_str = str(raw_book).strip()
        if book_str.startswith("必修"):
            book_code = "C" + book_str[2:] # 例如 "必修1" -> "C1"
        elif book_str.startswith("选修"):
            book_code = "E" + book_str[2:] # 例如 "选修3" -> "E3"
        else:
            # 如果课本名称不符合预期，可以选择保留原名或使用默认值
            book_code = book_str # 或者使用 "U" + book_str 表示未知
            print(f"警告：未识别的课本名称 '{book_str}'，使用原名。")

        # 处理单元 (raw_unit)
        # 确保 raw_unit 是数字，然后转换为整数字符串
        try:
            # xlrd 读取数字单元格时可能是 float 类型
            unit_num = int(float(raw_unit))
            unit_str = str(unit_num)
        except (ValueError, TypeError):
            # 如果无法转换为数字，则保留原始字符串
            unit_str = str(raw_unit).strip()
            print(f"警告：单元号 '{raw_unit}' 不是数字，使用原始字符串 '{unit_str}'。")

        # 创建组合键用于计数
        unit_key = (book_code, unit_str)

        # 更新该单元的计数器
        if unit_key in unit_counters:
            unit_counters[unit_key] += 1
        else:
            unit_counters[unit_key] = 1

        # 生成 num: PEP-S-{课本缩写}-{单元}-{单元内序号}
        current_num = f"PEP-S-{book_code}-{unit_str}-{unit_counters[unit_key]}"

        # 处理 word
        word = str(raw_word).strip()

        # 处理 trans
        trans = str(raw_trans).strip()

        # 处理 symbol
        symbol_str = str(raw_symbol).strip()
        if symbol_str:
            # 检查是否已经以 / 开头和结尾
            if not (symbol_str.startswith('/') and symbol_str.endswith('/')):
                # 如果没有，则添加
                symbol = f"/{symbol_str}/"
            else:
                # 如果已经包含，则保留原样
                symbol = symbol_str
        else:
            # 如果原始音标为空，则符号也为空字符串
            symbol = ""

        item = {
            "num": current_num,
            "word": word,
            "symbol": symbol,
            "trans": trans
        }
        json_objects.append(item)
    return json_objects

# --- 主执行流程 ---
if __name__ == "__main__":
    input_filename = 'PEP-S.xls'

    try:
        # 1. 从 xls 提取数据
        print(f"正在从 {input_filename} 提取数据...")
        raw_data = extract_data_from_xls(input_filename)
        print(f"提取到 {len(raw_data)} 行数据。")

        # 2. 处理数据并生成 JSON 对象
        print("正在处理数据并生成 JSON...")
        json_list = process_pep_s_data(raw_data)

        # 3. 输出 JSON
        if json_list:
            # 输出为 JSON 格式字符串（每项一行，便于复制粘贴）
            json_lines = [json.dumps(item, ensure_ascii=False) for item in json_list]
            output = ",\n".join(json_lines)

            # 打印结果 (可以选择只打印前几行以供预览)
            print("[")
            # print("\n".join(json_lines[:10])) # 预览前10行
            # print("...") if len(json_lines) > 10 else None
            # print(json_lines[-1] if len(json_lines) > 10 else "") # 如果超过10行，打印最后一行
            print(output) # 打印全部
            print("]")

            # 可选：保存到文件
            output_filename = 'pep_s_output_fixed_unit.json'
            with open(output_filename, 'w', encoding='utf-8') as f:
                f.write("[\n")
                f.write(",\n".join(json_lines))
                f.write("\n]")
            print(f"\n结果已保存到 {output_filename}")
        else:
            print("处理后没有生成任何有效的 JSON 对象。")

    except FileNotFoundError:
        print(f"错误：找不到文件 {input_filename}")
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
