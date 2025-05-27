import pdfplumber
import os
import re
import pandas as pd
import logging
from typing import List, Dict, Tuple, Optional, Any

# 配置常量
REPORTS_DIR = "报告(2)"
OUTPUT_DIR = "提取结果"
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

# 预编译正则表达式
PART_PATTERN = re.compile(r'(?:零件名|零件|部件名|部件)[:：]?\s*(.+)')
REVISION_PATTERN = re.compile(r'修订号[:：]\s*(.+)')
SERIAL_PATTERN = re.compile(r'(?:序列号|序号|编号)[:：]?\s*(.+)')
COUNT_PATTERN = re.compile(r'统计计数[:：]\s*(.+)')
NOTE_PATTERN = re.compile(r'以下(.+?)(?:，|。|$)')
DIM_FEATURE_PATTERN = re.compile(r'DIM\s+(\S+)=\s*(.*?)(?:\s+单位=|$)')
DIM_SIMPLE_PATTERN = re.compile(r'DIM\s+(.*?)(?:\s+单位=|$)')
CHINESE_CHAR_PATTERN = re.compile(r'[\u4e00-\u9fa5]')
DIGIT_PATTERN = re.compile(r'\d+')
OP_CODE_PATTERN = re.compile(r'OP\d+[A-Za-z0-9]*')

def setup_logging():
    """设置日志配置"""
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
    return logging.getLogger(__name__)

def ensure_directories(reports_dir: str, output_dir: str) -> Tuple[str, str]:
    """确保必要的目录存在"""
    if not os.path.exists(reports_dir):
        logging.warning(f"警告：目录 '{reports_dir}' 不存在，将使用当前目录")
        reports_dir = "."
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    return reports_dir, output_dir

def get_pdf_files(reports_dir: str) -> List[str]:
    """获取目录中的所有PDF文件"""
    pdf_files = [f for f in os.listdir(reports_dir) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        logging.error(f"错误：在 '{reports_dir}' 中未找到PDF文件")
    
    return pdf_files

def extract_pdf_text(pdf_path: str) -> List[str]:
    """从PDF文件中提取所有文本行"""
    all_lines = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            logging.info(f"开始处理PDF文件: {pdf_path}")
            
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_lines.extend(text.split('\n'))
        
        return all_lines
    except Exception as e:
        logging.error(f"提取PDF文本时出错: {e}")
        return []

def extract_header_info(all_lines: List[str], pdf_path: str) -> Dict[str, str]:
    """从PDF内容中提取标题信息"""
    info = {
        "part_name": "",
        "revision": "",
        "serial_number": "",
        "count": ""
    }
    
    # 从内容提取信息
    for i, line in enumerate(all_lines):
        # 提取零件名
        if not info["part_name"]:
            part_match = PART_PATTERN.search(line)
            if part_match:
                raw_part_name = part_match.group(1).strip()
                # 查找下一行的型号信息
                if i + 1 < len(all_lines):
                    model_line = all_lines[i + 1].strip()
                    if model_line.startswith(':'):
                        model_info = model_line[1:].strip()
                        # 组合零件名和型号
                        part_name = f"{raw_part_name}{model_info}".replace(' ', '')
                        # 在汉字和字母/数字之间添加空格
                        info["part_name"] = re.sub(r'([\u4e00-\u9fa5])(?!$)([A-Za-z0-9])', r'\1 \2', part_name)
        
        # 提取修订号
        if not info["revision"]:
            revision_match = REVISION_PATTERN.search(line)
            if revision_match:
                info["revision"] = revision_match.group(1).strip()
        
        # 提取序列号
        if not info["serial_number"]:
            serial_match = SERIAL_PATTERN.search(line)
            if serial_match:
                raw_serial = serial_match.group(1).strip()
                if i + 1 < len(all_lines):
                    shift_info = all_lines[i + 1].strip()
                    if shift_info.startswith(':'):
                        info["serial_number"] = f"{raw_serial} {shift_info[1:].strip()}"
                    else:
                        info["serial_number"] = raw_serial
                else:
                    info["serial_number"] = raw_serial
        
        # 提取计数
        if not info["count"]:
            count_match = COUNT_PATTERN.search(line)
            if count_match:
                info["count"] = count_match.group(1).strip()
    
    # 如果从内容中找不到信息，尝试从文件名提取
    if not info["part_name"]:
        base_name = os.path.basename(pdf_path)
        name_parts = base_name.split("  ", 1)  # 使用两个空格分割
        
        if len(name_parts) >= 2:
            info["part_name"] = name_parts[0]
            if not info["serial_number"]:
                info["serial_number"] = name_parts[1].replace(".pdf", "").replace(".PDF", "")
        else:
            info["part_name"] = os.path.splitext(base_name)[0]
    
    return info

def find_notes_in_text(all_lines: List[str], dim_indices: List[int]) -> Dict[int, str]:
    """在文本中查找所有备注信息"""
    notes = {}
    
    for i, line in enumerate(all_lines):
        match = NOTE_PATTERN.search(line)
        if match:
            # 找到紧随其后的第一个DIM索引
            next_dim_index = next((idx for idx in dim_indices if idx > i), None)
            if next_dim_index is not None:
                # 提取匹配的内容部分
                content = match.group(1).strip()
                
                # 检查下一行是否为操作代码
                op_code = ""
                if i+1 < len(all_lines):
                    op_match = OP_CODE_PATTERN.search(all_lines[i+1])
                    if op_match:
                        op_code = op_match.group(0)
                
                # 构建备注信息
                if "仅供测量员参考" in line:
                    if op_code:
                        modified_note = f"从此处开始为{op_code}控制图项目尺寸"
                    else:
                        modified_note = f"从此处开始为{content}控制图项目尺寸"
                elif op_code:
                    modified_note = f"从此处开始为{op_code}{content}"
                else:
                    modified_note = line.replace("以下", "从此处开始")
                
                notes[next_dim_index] = modified_note
    
    return notes

def merge_dim_with_previous_line(all_lines: List[str], dim_indices: List[int]) -> None:
    """合并DIM行与前一行的描述信息"""
    for idx in dim_indices:
        if idx > 0:  # 确保不是第一行
            current_line = all_lines[idx-1].strip()
            next_line = all_lines[idx].strip()
            
            # 判断当前行是否包含足够多的汉字
            chinese_chars = CHINESE_CHAR_PATTERN.findall(current_line)
            
            # 只有当汉字数量达到5个或更多时才进行合并
            if len(chinese_chars) >= 5:
                # 将当前行和下一行按空格分割
                current_parts = re.split(r'(\s+)', current_line)
                next_parts = re.split(r'(\s+)', next_line)
                
                # 检查当前行是否包含大量数字
                digit_count = len(DIGIT_PATTERN.findall(current_line))
                if digit_count > 3:  # 如果数字太多，跳过合并
                    continue
                
                # 删除DIM前缀
                try:
                    dim_index = next_parts.index('DIM')
                    next_parts = next_parts[dim_index + 1:]  # 去掉'DIM'和前面的内容
                    
                    # 创建新行，以DIM开头
                    merged_parts = ['DIM']
                    
                    # 交替合并两行的内容
                    desc_index = 0
                    data_index = 0
                    
                    while desc_index < len(current_parts):
                        # 添加描述部分
                        if desc_index < len(current_parts):
                            part = current_parts[desc_index]
                            if part.strip():  # 如果不是空白
                                merged_parts.append(' ' + part)
                            else:
                                merged_parts.append(part)  # 保留空格
                            desc_index += 1
                        
                        # 添加数据部分
                        if data_index < len(next_parts):
                            part = next_parts[data_index]
                            if part.strip():  # 如果不是空白
                                merged_parts.append(part)
                            data_index += 1
                    
                    # 将剩余的数据部分添加到末尾
                    while data_index < len(next_parts):
                        merged_parts.append(next_parts[data_index])
                        data_index += 1
                    
                    # 合并所有部分成一行
                    all_lines[idx] = ''.join(merged_parts).strip()
                except ValueError:
                    # 如果没有找到'DIM'，继续处理下一行
                    continue

def extract_feature_name(dim_line: str) -> str:
    """从DIM行中提取特征名称"""
    feature_match = DIM_FEATURE_PATTERN.search(dim_line)
    if feature_match:
        dim_id = feature_match.group(1).strip()
        feature_desc = feature_match.group(2).strip()
        return f"{dim_id}= {feature_desc}"
    
    feature_match = DIM_SIMPLE_PATTERN.search(dim_line)
    if feature_match:
        return feature_match.group(1).strip()
    
    # 如果上述正则都不匹配，使用简单的字符串分割
    return dim_line[4:].split(' 单位=')[0].strip()

def find_data_lines(segment_lines: List[str]) -> List[Tuple[str, str, int, List[int]]]:
    """在DIM段落中查找所有可能的数据行"""
    candidate_lines = []
    i = 0
    
    while i < len(segment_lines):
        line = segment_lines[i]
        # 跳过空行、页眉页脚和表头
        if not line.strip() or 'PART NUMBER=' in line or line.startswith('AX '):
            i += 1
            continue
        
        parts = line.split()
        # 检查是否是只有轴信息的行
        if len(parts) == 1 and len(parts[0]) <= 5 and i+1 < len(segment_lines):
            next_line = segment_lines[i+1]
            next_parts = next_line.split()
            # 检查下一行是否是数值行
            if next_line.startswith(' ') and len(next_parts) >= 6:
                try:
                    # 尝试将下一行的第一个非空元素转换为浮点数
                    float(next_parts[0].strip())
                    # 组合当前行（轴）和下一行（数据）
                    combined_line = parts[0] + " " + next_line.lstrip()
                    line = combined_line
                    parts = line.split()
                    i += 2  # 跳过下一行
                except ValueError:
                    i += 1
                    continue
            else:
                i += 1
                continue
        
        # 常规数据行处理逻辑
        if len(parts) >= 7 and len(parts[0]) <= 5:
            float_count = 0
            float_positions = []
            non_float_parts = []
            
            for j, part in enumerate(parts[1:], 1):
                try:
                    value = part.strip('=')
                    float(value)
                    float_count += 1
                    float_positions.append(j)
                except ValueError:
                    # 记录非浮点数部分
                    non_float_parts.append((j, part))
                    continue
            
            if float_count >= 1:
                candidate_lines.append((line, parts[0], float_count, float_positions))
                # 输出含有非浮点数的数据行信息
                if non_float_parts and logging.getLogger().level <= logging.DEBUG:
                    logging.debug(f"发现含非浮点数的数据行: {line}")
                    logging.debug(f"轴: {parts[0]}, 非浮点数部分: {non_float_parts}")
                    logging.debug(f"分割后的所有部分: {parts}")
        i += 1
    
    return candidate_lines

def process_dim_segments(all_lines: List[str], dim_indices: List[int], notes: Dict[int, str]) -> List[List[Any]]:
    """处理所有DIM段落并提取数据"""
    data_rows = []
    row_count = 1
    
    for i in range(len(dim_indices)):
        start_idx = dim_indices[i]
        # 如果是最后一个DIM，则结束索引是文件末尾
        end_idx = dim_indices[i+1] if i < len(dim_indices) - 1 else len(all_lines)
        
        # 获取特征名称
        dim_line = all_lines[start_idx]
        feature_name = extract_feature_name(dim_line)
        
        # 获取当前DIM段落的所有行
        segment_lines = all_lines[start_idx:end_idx]
        
        # 查找所有可能的数据行
        candidate_lines = find_data_lines(segment_lines)
        
        # 从候选行中选择浮点数最多的行
        if candidate_lines:
            best_line = max(candidate_lines, key=lambda x: x[2])
            data_line = best_line[0]
            
            # 分析数据行
            parts = data_line.split()
            
            # 输出最终选择的数据行信息
            if logging.getLogger().level <= logging.DEBUG:
                logging.debug(f"选择的数据行: {data_line}")
                non_float_parts = []
                for j, part in enumerate(parts):
                    try:
                        value = part.strip('=')
                        float(value)
                    except ValueError:
                        if j > 0:  # 忽略轴信息
                            non_float_parts.append((j, part))
                
                if non_float_parts:
                    logging.debug(f"最终数据行中的非浮点数部分: {non_float_parts}")
                    logging.debug(f"最终分割后的所有部分: {parts}")
            
            # 确保有足够的列
            if len(parts) >= 7:  # 轴信息 + 6个数值
                # 当前顺序：名义值、上公差、下公差、实际值、偏差、超差
                current_values = [parts[1], parts[2], parts[3], parts[4], parts[5], parts[6]]
                # 调整顺序为：实际值, 名义值, 上公差, 下公差, 偏差, 超差
                reordered_values = [current_values[3], current_values[0], current_values[1],
                                    current_values[2], current_values[4], current_values[5]]
                
                # 获取备注信息
                note = notes.get(start_idx, "")
                
                # 组装行数据
                row_data = [row_count, feature_name] + reordered_values + [note]
                data_rows.append(row_data)
                row_count += 1
    
    return data_rows

def save_to_excel(data_rows: List[List[Any]], excel_path: str) -> None:
    """将数据保存到Excel文件"""
    columns = ['序号', '特征名称', '实际值', '名义值', '上公差', '下公差', '偏差', '超差', '备注']
    df = pd.DataFrame(data_rows, columns=columns)
    
    try:
        df.to_excel(excel_path, index=False)
        logging.info(f"数据已提取并保存到Excel文件：{excel_path}")
        logging.info(f"总共提取了 {len(data_rows)} 条数据记录")
    except Exception as e:
        logging.error(f"保存Excel文件时出错: {e}")

def process_pdf_file(pdf_path: str, output_dir: str) -> None:
    """处理单个PDF文件并提取数据"""
    try:
        # 提取所有文本行
        all_lines = extract_pdf_text(pdf_path)
        if not all_lines:
            return
        
        # 提取标题信息
        info = extract_header_info(all_lines, pdf_path)
        
        # 打印提取的信息
        logging.info(f"零件名: {info['part_name']}")
        logging.info(f"修订号: {info['revision']}")
        logging.info(f"序列号: {info['serial_number']}")
        logging.info(f"统计计数: {info['count']}")
        
        # 查找所有DIM开头的行的索引
        dim_indices = [i for i, line in enumerate(all_lines) if line.startswith('DIM ')]
        
        # 查找备注信息
        notes = find_notes_in_text(all_lines, dim_indices)
        
        # 合并DIM上一行的文本
        merge_dim_with_previous_line(all_lines, dim_indices)
        
        # 处理所有DIM段落
        data_rows = process_dim_segments(all_lines, dim_indices, notes)
        
        # 输出文件名
        excel_filename = f"{info['part_name']} {info['serial_number']}.xlsx" if info['serial_number'] else f"{info['part_name']}.xlsx"
        excel_path = os.path.join(output_dir, excel_filename)
        
        # 保存到Excel
        save_to_excel(data_rows, excel_path)
        
    except Exception as e:
        logging.error(f"处理PDF时发生错误: {e}")
        import traceback
        traceback.print_exc()

def main():
    """主函数"""
    # 设置日志
    logger = setup_logging()
    
    try:
        # 确保目录存在
        reports_dir, output_dir = ensure_directories(REPORTS_DIR, OUTPUT_DIR)
        
        # 获取PDF文件列表
        pdf_files = get_pdf_files(reports_dir)
        if not pdf_files:
            return
        
        # 处理每个PDF文件
        for pdf_filename in pdf_files:
            pdf_path = os.path.join(reports_dir, pdf_filename)
            
            # 检查文件是否存在
            if not os.path.exists(pdf_path):
                logger.error(f"错误：找不到文件 '{pdf_path}'")
                continue
            
            process_pdf_file(pdf_path, output_dir)
    
    except Exception as e:
        logger.error(f"程序执行时发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
