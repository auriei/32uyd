import pdfplumber
import os
import re
import pandas as pd
import logging
from typing import List, Dict, Tuple, Optional, Any
import statistics

# --- Constants ---
REPORTS_DIR = "."
OUTPUT_DIR = "./excel_output"
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
LINE_TOLERANCE = 2
MIN_X_COORD = 5.0
MAX_X_COORD = 700.0 
DATA_ROW_NUMERIC_THRESHOLD = 0.5
DATA_ROW_MIN_COLUMNS = 2 # Minimum plausible columns for a data row
HEADER_ROW_CANDIDATE_LINES = 15
HEADER_MIN_WORDS = 3 # Slightly more flexible for header
HEADER_MAX_WORDS = 15
HEADER_NON_NUMERIC_RATIO = 0.5 # Slightly more flexible

KAN_HEADERS_TEMPLATE = ['轴', '标称值', '正公差', '负公差', '测定', '最大值', '最小值', '偏差', '超差']
EXCEL_OUTPUT_COLUMNS = ['序号', '特征名称', '实际值', '名义值', '上公差', '下公差', '偏差', '超差', '备注']

# --- Utility Functions ---
def setup_logging():
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
    # logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT)
    return logging.getLogger(__name__)

def ensure_directories(reports_dir: str, output_dir: str) -> Tuple[str, str]:
    if not os.path.exists(reports_dir):
        logging.warning(f"报告目录 '{reports_dir}' 不存在，将使用当前目录。")
        reports_dir = "."
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logging.info(f"创建目录: {output_dir}")
    return reports_dir, output_dir

def get_pdf_files(reports_dir: str) -> List[str]:
    try:
        files = [f for f in os.listdir(reports_dir) if f.lower().endswith('.pdf')]
        if not files: logging.warning(f"在 '{reports_dir}' 中未找到PDF文件。")
        return files
    except FileNotFoundError:
        logging.error(f"错误：报告目录 '{reports_dir}' 不存在。")
        return []

def sanitize_filename_component(text: str, default_if_empty: str = "comp") -> str:
    """Aggressively sanitizes a string component for use in a filename."""
    if not text or not text.strip():
        return default_if_empty
    
    # Remove leading/trailing whitespace
    text = text.strip()
    
    # Replace problematic characters (anything not alphanumeric, underscore, or dot) with underscore
    # Keep hyphens but ensure they are not leading/trailing after this step.
    text = re.sub(r'[^\w.-]', '_', text)
    
    # Remove leading hyphens or dots
    if text.startswith("-") or text.startswith("."):
        text = "_" + text[1:]
        
    # Consolidate multiple underscores
    text = re.sub(r'_+', '_', text)
    
    # Remove leading/trailing underscores that might have been created
    text = text.strip('_')

    if not text: # If all chars were problematic and removed
        return default_if_empty
    return text


def is_numeric(text: str) -> bool:
    if text is None: return False
    text = text.strip()
    if not text: return False
    try:
        float(text.replace(',', ''))
        return True
    except ValueError:
        if re.match(r"^-?(\d*\.\d+|\d+\.\d*)$", text): return True
    return False

# --- PDF Parsing and Grid Generation ---
def extract_words_and_reconstruct_lines(pdf_path: str, line_tolerance: float = LINE_TOLERANCE) -> List[List[Dict[str, Any]]]:
    all_reconstructed_lines: List[List[Dict[str, Any]]] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_words_from_pdf: List[Dict[str, Any]] = []
            if not pdf.pages:
                logging.warning(f"PDF '{pdf_path}'没有任何页面.")
                return []
            for i, page in enumerate(pdf.pages):
                words_on_page = page.extract_words(
                    keep_blank_chars=True, use_text_flow=True, 
                    horizontal_ltr=True, vertical_ttb=True,
                    extra_attrs=["fontname", "size"]
                )
                if words_on_page: all_words_from_pdf.extend(words_on_page)
            if not all_words_from_pdf:
                logging.info(f"在 '{pdf_path}' 中未提取到任何单词。")
                return []
            all_words_from_pdf.sort(key=lambda word: (float(word['top']), float(word['x0'])))
            current_line: List[Dict[str, Any]] = []
            if all_words_from_pdf:
                current_line.append(all_words_from_pdf[0])
                for i in range(1, len(all_words_from_pdf)):
                    prev_word_top = float(current_line[0]['top'])
                    current_word = all_words_from_pdf[i]
                    if abs(float(current_word['top']) - prev_word_top) <= line_tolerance:
                        current_line.append(current_word)
                    else:
                        current_line.sort(key=lambda word: float(word['x0']))
                        all_reconstructed_lines.append(current_line)
                        current_line = [current_word]
                if current_line:
                    current_line.sort(key=lambda word: float(word['x0']))
                    all_reconstructed_lines.append(current_line)
            return all_reconstructed_lines
    except Exception as e:
        logging.error(f"提取单词或重建行时出错 '{pdf_path}': {e}", exc_info=True)
        return []

def identify_data_rows_and_header_ref(
    reconstructed_lines: List[List[Dict[str, Any]]],
) -> Tuple[List[int], Optional[int], int]:
    data_row_indices: List[int] = []
    header_row_idx: Optional[int] = None
    num_header_columns = 0
    
    for i, line_words in enumerate(reconstructed_lines):
        if not line_words or len(line_words) < DATA_ROW_MIN_COLUMNS: continue # Skip very short lines for data
        first_word_text = line_words[0].get('text', '').strip().upper()
        if first_word_text in ['M', 'X', 'Y', 'Z', 'L', 'AX', 'PT'] or re.match(r"^[A-Z]$", first_word_text):
            data_row_indices.append(i)
            continue
        if len(line_words) > 1:
            numeric_word_count = sum(1 for w in line_words[1:] if is_numeric(w.get('text', '')))
            if len(line_words) > 2 and numeric_word_count >= (DATA_ROW_MIN_COLUMNS -1) and \
               (numeric_word_count / (len(line_words) -1)) >= DATA_ROW_NUMERIC_THRESHOLD :
                data_row_indices.append(i)

    best_header_candidate_idx: Optional[int] = None
    max_heuristic_score = -1

    for i, line_words in enumerate(reconstructed_lines[:HEADER_ROW_CANDIDATE_LINES]):
        if not line_words or not (HEADER_MIN_WORDS <= len(line_words) <= HEADER_MAX_WORDS):
            continue
        
        non_numeric_count = sum(1 for w in line_words if not is_numeric(w.get('text','')))
        total_words = len(line_words)
        current_score = 0
        
        if total_words > 0 and (non_numeric_count / total_words) >= HEADER_NON_NUMERIC_RATIO:
            current_score += non_numeric_count # Base score on number of text items
            if data_row_indices and i == data_row_indices[0] -1: # Bonus if just before first data row
                current_score += total_words # Boost score significantly
            
            if current_score > max_heuristic_score:
                max_heuristic_score = current_score
                best_header_candidate_idx = i
    
    if best_header_candidate_idx is not None:
        header_row_idx = best_header_candidate_idx
        num_header_columns = len(reconstructed_lines[header_row_idx])
        # logging.info(f"选定的头行索引: {header_row_idx}，有 {num_header_columns} 列。")
    elif reconstructed_lines: 
        # Fallback: take the first line that has at least HEADER_MIN_WORDS
        for i, line_words in enumerate(reconstructed_lines[:HEADER_ROW_CANDIDATE_LINES]):
            if len(line_words) >= HEADER_MIN_WORDS:
                header_row_idx = i
                num_header_columns = len(line_words)
                logging.warning(f"无法通过启发式找到清晰的头行，使用索引 {i}，有 {num_header_columns} 列。")
                break
        if header_row_idx is None: # Absolute fallback to first line
            header_row_idx = 0
            num_header_columns = len(reconstructed_lines[0]) if reconstructed_lines[0] else 0
            logging.warning(f"绝对回退：使用索引0作为头行，有 {num_header_columns} 列。")
    else:
        header_row_idx = 0
        num_header_columns = 0
        logging.error("PDF中未重建任何行，无法确定头行。")
    return data_row_indices, header_row_idx, num_header_columns

def define_adaptive_column_boundaries(
    reconstructed_lines: List[List[Dict[str, Any]]],
    data_row_indices: List[int],
    header_row_idx: int, 
    num_expected_columns: int,
    min_x: float = MIN_X_COORD, max_x: float = MAX_X_COORD
) -> List[Tuple[float, float]]:
    if num_expected_columns == 0: return [(min_x, max_x)]
    column_x0_candidates: List[List[float]] = [[] for _ in range(num_expected_columns)]
    source_rows_for_x0s = data_row_indices
    if not data_row_indices:
        if header_row_idx < len(reconstructed_lines): source_rows_for_x0s = [header_row_idx]
        else: return [(min_x, max_x)] 
    for row_idx in source_rows_for_x0s:
        if row_idx >= len(reconstructed_lines): continue
        line_words = reconstructed_lines[row_idx]
        for word_idx, word in enumerate(line_words):
            if word_idx < num_expected_columns: column_x0_candidates[word_idx].append(float(word['x0']))
    avg_x0_starts: List[Optional[float]] = [None] * num_expected_columns
    ref_header_words = reconstructed_lines[header_row_idx] if header_row_idx < len(reconstructed_lines) else []
    for i in range(num_expected_columns):
        if column_x0_candidates[i]: avg_x0_starts[i] = statistics.mean(column_x0_candidates[i])
        else: 
            if i < len(ref_header_words): avg_x0_starts[i] = float(ref_header_words[i]['x0'])
            elif i > 0 and avg_x0_starts[i-1] is not None: avg_x0_starts[i] = avg_x0_starts[i-1] + 50 
            else: avg_x0_starts[i] = min_x + i * ((max_x - min_x) / num_expected_columns if num_expected_columns > 0 else 50)
    effective_col_starts = sorted([x for x in avg_x0_starts if x is not None and x <= max_x])
    if not effective_col_starts: return [(min_x, max_x)]
    boundaries: List[Tuple[float, float]] = []
    current_col_start = min_x
    for i in range(len(effective_col_starts)):
        col_text_start = effective_col_starts[i]
        if i + 1 < len(effective_col_starts):
            next_col_text_start = effective_col_starts[i+1]
            col_end = (col_text_start + next_col_text_start) / 2.0 
            if col_end <= current_col_start + 1: col_end = current_col_start + 10 
        else: col_end = max_x
        actual_start = min(current_col_start, col_text_start if col_text_start > current_col_start else current_col_start )
        if actual_start >= col_end : actual_start = col_end -10 if col_end > 10 else 0
        boundaries.append((actual_start, col_end))
        current_col_start = col_end if col_end > actual_start else actual_start + 10 
    if boundaries and boundaries[-1][1] < max_x : 
        last_start, _ = boundaries[-1]
        if last_start < max_x: boundaries[-1] = (last_start, max_x)
        else: boundaries[-1] = (max_x-10, max_x) 
    elif not boundaries and num_expected_columns > 0: boundaries.append((min_x,max_x))
    return boundaries

def map_lines_to_grid(reconstructed_lines: List[List[Dict[str, Any]]], 
                      column_boundaries: List[Tuple[float, float]]) -> List[List[str]]:
    grid: List[List[str]] = []
    if not column_boundaries :
        for line in reconstructed_lines: grid.append([" ".join(w.get('text','') for w in line).strip()])
        return grid
    num_columns = len(column_boundaries)
    for line_idx, line_words in enumerate(reconstructed_lines):
        row_cells: List[str] = [""] * num_columns
        for word in line_words:
            word_text = word.get('text', '').strip(); w_x0, w_x1 = float(word['x0']), float(word['x1'])
            if not word_text: continue
            best_col_idx, max_overlap_ratio = -1, 0.0 # Use ratio for better assignment
            word_width = w_x1 - w_x0 if w_x1 > w_x0 else 1.0

            for col_idx, (csx, cex) in enumerate(column_boundaries):
                if csx >= cex : continue 
                overlap = max(0, min(w_x1, cex) - max(w_x0, csx))
                overlap_ratio = overlap / word_width
                
                if overlap_ratio > max_overlap_ratio:
                    max_overlap_ratio = overlap_ratio
                    best_col_idx = col_idx
                # If overlap is small, but it's the only option so far, take it.
                elif overlap > 0 and best_col_idx == -1 : 
                    best_col_idx = col_idx

            if best_col_idx != -1:
                if row_cells[best_col_idx]: row_cells[best_col_idx] += " "
                row_cells[best_col_idx] += word_text
        grid.append([cell.strip() for cell in row_cells])
    return grid

# --- Data Structuring and Excel Export ---
def extract_top_level_info_from_grid(grid_data: List[List[str]], pdf_filename: str) -> Dict[str, str]:
    info = {"part_name": sanitize_filename_component(os.path.splitext(pdf_filename)[0], "UnknownPart"), "serial_number": ""}
    for row_idx, row in enumerate(grid_data[:5]):
        row_text = " ".join(cell for cell in row if cell).replace("：", ":") # Join cells before regex
        part_match = re.search(r'(?:零件名|零件|部件名|部件)\s*:\s*([^\s]+(?:[\s_][^\s]+)*)', row_text, re.IGNORECASE)
        serial_match = re.search(r'(?:序列号|序号|编号)\s*:\s*([^\s]+(?:[\s_][^\s]+)*)', row_text, re.IGNORECASE)
        if part_match:
            extracted_part_name = part_match.group(1).strip()
            info["part_name"] = sanitize_filename_component(extracted_part_name, f"PartRow{row_idx}")
        if serial_match:
            extracted_serial = serial_match.group(1).strip()
            # Try to capture subsequent cells if they seem part of the serial
            if row_idx < len(grid_data) and len(row) > ( (row_text.index(extracted_serial) + len(extracted_serial)) // 10 ) : # crude check if serial is in first cell
                 try:
                    # Find which cell contained the serial_match.group(1)
                    cell_with_serial_val = ""
                    start_char_idx_of_serial_val = row_text.find(extracted_serial)
                    current_char_count = 0
                    serial_cell_idx = -1
                    for idx, cell_content in enumerate(row):
                        if current_char_count <= start_char_idx_of_serial_val < current_char_count + len(cell_content) +1: # +1 for space
                            serial_cell_idx = idx
                            break
                        current_char_count += len(cell_content) + 1
                    
                    if serial_cell_idx != -1 and serial_cell_idx + 1 < len(row) and row[serial_cell_idx+1]:
                        extracted_serial += " " + row[serial_cell_idx+1]
                 except Exception: pass # Ignore errors in this complex heuristic for now

            info["serial_number"] = sanitize_filename_component(extracted_serial, f"SerialRow{row_idx}")
            # No break, might find better one later, though unlikely for these fields
    if not info["serial_number"]: info["serial_number"] = "NoSerial" # Default if not found
    return info


def structure_data_for_excel(
    grid_data: List[List[str]], 
    data_row_indices: List[int],
    num_actual_columns_in_grid: int
) -> List[List[Any]]:
    final_data_for_excel: List[List[Any]] = []
    current_feature_name = "N/A"
    item_serial_number = 1
    effective_kan_headers = KAN_HEADERS_TEMPLATE[:num_actual_columns_in_grid]
    kan_header_to_idx: Dict[str, int] = {name: i for i, name in enumerate(effective_kan_headers)}
    idx_actual = kan_header_to_idx.get('测定'); idx_nominal = kan_header_to_idx.get('标称值')
    idx_pos_tol = kan_header_to_idx.get('正公差'); idx_neg_tol = kan_header_to_idx.get('负公差')
    idx_dev = kan_header_to_idx.get('偏差'); idx_oot = kan_header_to_idx.get('超差')
    idx_axis = kan_header_to_idx.get('轴')

    for i, row_cells in enumerate(grid_data):
        if not any(row_cells): continue
        first_cell_text = row_cells[0].strip()
        if re.match(r"^\s*DIM", first_cell_text, re.IGNORECASE): # More robust DIM check
            current_feature_name = " ".join(cell for cell in row_cells if cell).strip()
            continue
        if i in data_row_indices:
            excel_row = [""] * len(EXCEL_OUTPUT_COLUMNS)
            excel_row[0] = item_serial_number; excel_row[1] = current_feature_name
            if idx_actual is not None and idx_actual < len(row_cells): excel_row[2] = row_cells[idx_actual]
            if idx_nominal is not None and idx_nominal < len(row_cells): excel_row[3] = row_cells[idx_nominal]
            if idx_pos_tol is not None and idx_pos_tol < len(row_cells): excel_row[4] = row_cells[idx_pos_tol]
            if idx_neg_tol is not None and idx_neg_tol < len(row_cells): excel_row[5] = row_cells[idx_neg_tol]
            if idx_dev is not None and idx_dev < len(row_cells): excel_row[6] = row_cells[idx_dev]
            if idx_oot is not None and idx_oot < len(row_cells): excel_row[7] = row_cells[idx_oot]
            remarks = []
            if idx_axis is not None and idx_axis < len(row_cells) and row_cells[idx_axis]:
                 remarks.append(f"轴: {row_cells[idx_axis]}")
            other_kan_values = []
            direct_map_kan_indices = {idx_actual, idx_nominal, idx_pos_tol, idx_neg_tol, idx_dev, idx_oot, idx_axis}
            for k_idx, k_name in enumerate(effective_kan_headers):
                if k_idx not in direct_map_kan_indices and k_idx < len(row_cells) and row_cells[k_idx]:
                    other_kan_values.append(f"{k_name}: {row_cells[k_idx]}")
            if other_kan_values: remarks.append("; ".join(other_kan_values))
            excel_row[8] = "; ".join(remarks).strip()
            final_data_for_excel.append(excel_row)
            item_serial_number += 1
    return final_data_for_excel

def save_to_excel(data_rows: List[List[Any]], excel_path: str) -> None:
    if not data_rows:
        logging.warning(f"没有数据可保存到Excel: {excel_path}")
        return
    df = pd.DataFrame(data_rows, columns=EXCEL_OUTPUT_COLUMNS)
    try:
        df.to_excel(excel_path, index=False)
        # Consolidated logging is now in process_pdf_file
    except Exception as e:
        logging.error(f"保存Excel文件时出错 '{excel_path}': {e}", exc_info=True)
        raise # Re-raise to be caught by process_pdf_file for failure tracking

# --- Main Processing Logic ---
def process_pdf_file(pdf_path: str, output_dir: str) -> bool:
    pdf_filename = os.path.basename(pdf_path)
    logging.info(f"开始处理PDF文件: {pdf_filename}")
    try:
        reconstructed_lines = extract_words_and_reconstruct_lines(pdf_path)
        if not reconstructed_lines: 
            logging.warning(f"'{pdf_filename}' 未重建任何行。跳过。")
            return False

        data_row_indices, header_row_idx, num_header_cols = identify_data_rows_and_header_ref(reconstructed_lines)
        
        if header_row_idx is None or header_row_idx >= len(reconstructed_lines):
            logging.error(f"未能确定 '{pdf_filename}' 的有效头行索引。跳过。")
            return False
        # num_header_cols is now determined more reliably in identify_data_rows_and_header_ref
        if num_header_cols == 0 and not data_row_indices : # If no cols and no data, probably useless
             logging.warning(f"'{pdf_filename}' 头行列数为0且无数据行。可能无法正确处理。")
             # return False # Decide if this is a hard fail

        column_boundaries = define_adaptive_column_boundaries(
            reconstructed_lines, data_row_indices, header_row_idx, num_header_cols
        )
        grid_data = map_lines_to_grid(reconstructed_lines, column_boundaries)
        
        top_level_info = extract_top_level_info_from_grid(grid_data, pdf_filename)
        
        final_data_for_excel = structure_data_for_excel(
            grid_data, data_row_indices, len(column_boundaries)
        )

        if final_data_for_excel:
            # Sanitize components before joining for filename
            s_part_name = sanitize_filename_component(top_level_info['part_name'], "UnknownPart")
            s_serial_num = sanitize_filename_component(top_level_info['serial_number'], "NoSerial")
            
            excel_filename_base = f"{s_part_name}_{s_serial_num}"
            excel_filename = f"{excel_filename_base}.xlsx"
            excel_path = os.path.join(output_dir, excel_filename)
            
            # Handle potential filename collision by appending a number
            counter = 1
            while os.path.exists(excel_path):
                excel_filename = f"{excel_filename_base}_{counter}.xlsx"
                excel_path = os.path.join(output_dir, excel_filename)
                counter += 1
            
            save_to_excel(final_data_for_excel, excel_path)
            logging.info(f"成功处理PDF '{pdf_filename}' 并保存到 '{excel_path}' ({len(final_data_for_excel)} 条记录).")
            return True
        else:
            logging.warning(f"未能从 '{pdf_filename}' 提取到最终数据用于Excel。")
            return False
            
    except Exception as e:
        logging.error(f"处理PDF '{pdf_filename}' 时发生未捕获错误: {e}", exc_info=True)
        return False

def main():
    logger = setup_logging()
    successful_pdfs: List[str] = []
    failed_pdfs: Dict[str, str] = {} 

    try:
        reports_dir, output_dir = ensure_directories(REPORTS_DIR, OUTPUT_DIR)
        pdf_files = get_pdf_files(reports_dir)
        if not pdf_files: return
        logger.info(f"找到 {len(pdf_files)} 个PDF文件进行处理。")

        for pdf_filename in pdf_files:
            pdf_path = os.path.join(reports_dir, pdf_filename)
            # Skip already processed files if they exist (simple check, could be more robust)
            # This check is basic; assumes default naming. If sanitization changes names a lot, this might not work.
            # For now, let's re-process all to test sanitization.
            # s_part_name_guess = sanitize_filename_component(os.path.splitext(pdf_filename)[0], "UnknownPart")
            # potential_excel_name = f"{s_part_name_guess}_NoSerial.xlsx" # A common default if serial is missing
            # if os.path.exists(os.path.join(output_dir, potential_excel_name)):
            #     logger.info(f"跳过 '{pdf_filename}', 可能已处理: '{potential_excel_name}' 存在。")
            #     successful_pdfs.append(f"{pdf_filename} (skipped, assumed processed)")
            #     continue

            try:
                if process_pdf_file(pdf_path, output_dir):
                    successful_pdfs.append(pdf_filename)
                else:
                    failed_pdfs[pdf_filename] = "Processing returned False."
            except Exception as e: 
                logging.error(f"处理 '{pdf_filename}' 时发生顶层意外错误: {e}", exc_info=True)
                failed_pdfs[pdf_filename] = f"Unexpected top-level error: {str(e)}"
    finally:
        logger.info("\n--- 处理总结 ---")
        logger.info(f"成功处理的PDF ({len(successful_pdfs)}):")
        for fname in successful_pdfs: logger.info(f"  - {fname}")
        logger.info(f"\n处理失败的PDF ({len(failed_pdfs)}):")
        for fname, err_summary in failed_pdfs.items(): logger.info(f"  - {fname}: {err_summary}")

if __name__ == "__main__":
    main()
