import xlrd
import json

# 打开 xls 文件
workbook = xlrd.open_workbook('CEE3500.xls')
sheet = workbook.sheet_by_index(0)  # 假设数据在第一个 sheet

data_list = []

# 假设表头在第一行，数据从第二行开始
for row_idx in range(1, sheet.nrows):
    row = sheet.row_values(row_idx)
    
    # 假设列顺序为：序号、单词、音标、释义
    raw_num = row[0] if len(row) > 0 else ""
    raw_word = row[1] if len(row) > 1 else ""
    raw_symbol = row[2] if len(row) > 2 else ""
    raw_trans = row[3] if len(row) > 3 else ""
    
    # 确保值是字符串类型再进行处理
    # 对于 num，我们使用原始值转换为字符串（去掉 .0），然后拼接前缀
    num_value = str(raw_num).strip()
    # 如果是 float 类型的整数，xlrd 会显示为 1.0, 2.0 等，需要处理掉 .0
    if num_value.endswith('.0'):
        num_value = num_value[:-2]
    num = f"CEE3500-{num_value}" 
    
    word = str(raw_word).strip()
    symbol = str(raw_symbol).strip()
    trans = str(raw_trans).strip()
    
    item = {
        "num": num,
        "word": word,
        "symbol": symbol,
        "trans": trans
    }
    data_list.append(item)

# 输出为 JSON 格式字符串（每项一行，便于复制粘贴）
json_lines = [json.dumps(item, ensure_ascii=False) for item in data_list]
output = ",\n".join(json_lines)

# 打印结果
print("[")
print(output)
print("]")

# 可选：保存到文件
with open('output.json', 'w', encoding='utf-8') as f:
    f.write("[\n")
    f.write(",\n".join(json_lines))
    f.write("\n]")