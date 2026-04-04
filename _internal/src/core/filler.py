"""
Word 表格填充逻辑
- 根据映射关系将 Excel 数据填充到 Word 表格
- 支持行数动态调整
- 支持合并单元格处理
"""

from lxml import etree as ET
import pandas as pd
import numpy as np
import os
import copy
import shutil
import tempfile

from .word_parser import (
    NAMESPACES, extract_docx, pack_docx, register_namespaces_from_file,
    capture_xml_wrapper, write_xml_preserving_wrapper,
    flatten_body_elements, get_paragraph_style,
    get_cell_text, is_merged_cell, detect_header_rows,
    get_row_first_cell_text, parse_styles, DEFAULT_SKIP_KEYWORDS
)


def format_value(value):
    """格式化单元格的值"""
    if pd.isna(value):
        return ''

    if isinstance(value, str):
        stripped = value.strip()
        if stripped == '-':
            return '-'
        try:
            num = float(stripped)
            if num == 0:
                return '-'
            return f"{num:,.2f}"
        except ValueError:
            return value

    if isinstance(value, (int, float, np.integer, np.floating)):
        if value == 0:
            return '-'
        return f"{value:,.2f}"

    return str(value)


def clear_cell_content(cell):
    """清空单元格的文本内容"""
    for t in cell.findall('.//w:t', NAMESPACES):
        t.text = ''


def set_cell_text(cell, text):
    """设置单元格的文本内容"""
    clear_cell_content(cell)

    p = cell.find('.//w:p', NAMESPACES)
    if p is None:
        return False

    run = p.find('.//w:r', NAMESPACES)
    w_ns = '{' + NAMESPACES['w'] + '}'

    if run is None:
        run = ET.SubElement(p, f'{w_ns}r')

    t = run.find('.//w:t', NAMESPACES)
    if t is None:
        t = ET.SubElement(run, f'{w_ns}t')

    t.text = format_value(text)
    return True


def fill_cell(cell, text, overwrite=False):
    """填充单元格"""
    existing_text = get_cell_text(cell).strip()

    if existing_text and not overwrite:
        return False

    if existing_text and overwrite:
        return set_cell_text(cell, text)

    p = cell.find('.//w:p', NAMESPACES)
    if p is None:
        return False

    w_ns = '{' + NAMESPACES['w'] + '}'
    run = ET.SubElement(p, f'{w_ns}r')
    t = ET.SubElement(run, f'{w_ns}t')
    t.text = format_value(text)

    return True


def copy_row(row):
    """深拷贝一个表格行"""
    return copy.deepcopy(row)


def clear_row_content(row):
    """清空行中所有单元格的文本内容"""
    cells = row.findall('.//w:tc', NAMESPACES)
    for cell in cells:
        clear_cell_content(cell)


def normalize_text(text):
    """标准化文本用于比较"""
    if not text:
        return ''
    return text.replace(' ', '').replace('：', ':').replace(':', '').strip()


def texts_match(excel_text, word_text):
    """判断 Excel 文本和 Word 文本是否匹配"""
    if not excel_text and not word_text:
        return True
    if not excel_text or not word_text:
        return False

    e_norm = normalize_text(excel_text)
    w_norm = normalize_text(word_text)

    if e_norm == w_norm:
        return True
    if w_norm.startswith(e_norm) or e_norm.startswith(w_norm):
        return True

    return False


def detect_excel_header_rows(df, word_header_texts):
    """
    通过与 Word 表头内容比对来检测 Excel 表头行数
    """
    excel_header_count = 0

    for i in range(min(len(word_header_texts), len(df))):
        excel_text = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
        word_text = word_header_texts[i] if i < len(word_header_texts) else ''

        if texts_match(excel_text, word_text):
            excel_header_count = i + 1
        else:
            break

    return max(1, excel_header_count)


def fill_table_from_dataframe(table, df):
    """
    用 DataFrame 的数据填充表格
    返回: (filled_count, added_rows, removed_rows)
    """
    rows = table.findall('.//w:tr', NAMESPACES)
    if not rows:
        return 0, 0, 0

    # 检测 Word 表头行数
    word_header_count = detect_header_rows(rows)

    # 获取 Word 表头各行的第一列文本
    word_header_texts = []
    for i in range(word_header_count):
        text = get_row_first_cell_text(rows[i])
        word_header_texts.append(text)

    # 检测 Excel 表头行数
    excel_header_count = detect_excel_header_rows(df, word_header_texts)

    excel_data_count = len(df) - excel_header_count
    if excel_data_count <= 0:
        return 0, 0, 0

    excel_first_col = []
    for i in range(excel_header_count, len(df)):
        val = df.iloc[i, 0]
        excel_first_col.append(str(val).strip() if pd.notna(val) else '')

    word_data_rows = rows[word_header_count:]
    word_first_col = [get_row_first_cell_text(row) for row in word_data_rows]

    matched_word_indices = []
    word_search_start = 0

    for excel_idx, excel_text in enumerate(excel_first_col):
        found_word_idx = None
        for word_idx in range(word_search_start, len(word_first_col)):
            word_text = word_first_col[word_idx]
            if texts_match(excel_text, word_text):
                found_word_idx = word_idx
                word_search_start = word_idx + 1
                break
        matched_word_indices.append(found_word_idx)

    word_indices_to_keep = set(idx for idx in matched_word_indices if idx is not None)

    # 选择模板行
    template_row = None
    max_cells = 0
    for row in word_data_rows:
        cells = row.findall('.//w:tc', NAMESPACES)
        if len(cells) > max_cells:
            max_cells = len(cells)
            template_row = row

    added_rows = 0
    removed_rows = 0

    rows_to_delete = [i for i in range(len(word_data_rows)) if i not in word_indices_to_keep]
    for idx in reversed(rows_to_delete):
        actual_row = rows[word_header_count + idx]
        table.remove(actual_row)
        removed_rows += 1

    rows = table.findall('.//w:tr', NAMESPACES)
    word_data_rows_remaining = rows[word_header_count:]

    kept_indices = sorted(word_indices_to_keep)
    old_to_new_idx = {old_idx: new_idx for new_idx, old_idx in enumerate(kept_indices)}

    final_rows = []
    for excel_idx, matched_word_idx in enumerate(matched_word_indices):
        if matched_word_idx is not None and matched_word_idx in old_to_new_idx:
            new_idx = old_to_new_idx[matched_word_idx]
            final_rows.append(('existing', new_idx))
        else:
            final_rows.append(('new', None))
            added_rows += 1

    current_data_rows = table.findall('.//w:tr', NAMESPACES)[word_header_count:]
    for row in current_data_rows:
        table.remove(row)

    rows = table.findall('.//w:tr', NAMESPACES)
    for row_type, idx in final_rows:
        if row_type == 'existing' and idx is not None and idx < len(word_data_rows_remaining):
            table.append(word_data_rows_remaining[idx])
        else:
            if template_row is not None:
                new_row = copy_row(template_row)
                clear_row_content(new_row)
                table.append(new_row)

    rows = table.findall('.//w:tr', NAMESPACES)
    word_data_rows = rows[word_header_count:]

    filled_count = 0
    for row_idx, row in enumerate(word_data_rows):
        excel_row_idx = excel_header_count + row_idx
        if excel_row_idx >= len(df):
            break

        cells = row.findall('.//w:tc', NAMESPACES)
        df_row = df.iloc[excel_row_idx]

        for col_idx, cell in enumerate(cells):
            if is_merged_cell(cell):
                continue
            if col_idx < len(df.columns):
                value = df_row.iloc[col_idx]
                if fill_cell(cell, value, overwrite=True):
                    filled_count += 1

    return filled_count, added_rows, removed_rows


def check_empty_header_cells(table):
    """
    检查表格标题行是否有空单元格
    返回: [(行号, 列号), ...]
    """
    rows = table.findall('.//w:tr', NAMESPACES)
    if not rows:
        return []

    header_row_count = detect_header_rows(rows)
    empty_cells = []

    for row_idx in range(header_row_count):
        row = rows[row_idx]
        cells = row.findall('.//w:tc', NAMESPACES)

        for col_idx, cell in enumerate(cells):
            if is_merged_cell(cell):
                continue

            text = get_cell_text(cell).strip()
            if not text:
                empty_cells.append((row_idx + 1, col_idx + 1))

    return empty_cells


class DocumentFiller:
    """文档填充器"""

    def __init__(self, docx_path, excel_source, output_path=None, data_start_row=1, data_start_col=1, single_file_mode=False, skip_keywords=None):
        """
        初始化填充器

        参数:
            docx_path: Word 模板路径（.docx 文件或解压目录）
            excel_source: Excel 数据源（文件夹路径或单个文件路径）
            output_path: 输出文件路径（默认在原文件同目录下生成 _filled 文件）
            data_start_row: Excel 数据起始行（从1开始计数）
            data_start_col: Excel 数据起始列（从1开始计数）
            single_file_mode: 是否为单文件模式
            skip_keywords: 忽略关键词列表
        """
        self.docx_path = docx_path
        self.excel_source = excel_source
        self.data_start_row = max(1, data_start_row)
        self.data_start_col = max(1, data_start_col)
        self.single_file_mode = single_file_mode
        self.skip_keywords = skip_keywords if skip_keywords is not None else DEFAULT_SKIP_KEYWORDS

        if output_path is None:
            base, ext = os.path.splitext(docx_path)
            output_path = f"{base}_filled{ext}"
        self.output_path = output_path

        self.work_dir = None
        self.root = None
        self.tables_dict = {}
        self.style_levels = {}  # 样式到层级的映射
        self.document_xml_declaration = None
        self.document_root_start_tag = None
        self.document_root_end_tag = None

    def prepare(self):
        """准备工作目录"""
        # 解压或复制模板
        if os.path.isfile(self.docx_path) and self.docx_path.endswith('.docx'):
            self.work_dir = tempfile.mkdtemp(prefix='word_filler_')
            extract_docx(self.docx_path, self.work_dir)
        else:
            self.work_dir = tempfile.mkdtemp(prefix='word_filler_')
            shutil.copytree(self.docx_path, self.work_dir, dirs_exist_ok=True)

        # 解析样式层级
        self.style_levels = parse_styles(self.work_dir)

        # 解析 XML（先注册模板中的所有命名空间）
        xml_path = os.path.join(self.work_dir, 'word', 'document.xml')
        (
            self.document_xml_declaration,
            self.document_root_start_tag,
            self.document_root_end_tag,
        ) = capture_xml_wrapper(xml_path)
        register_namespaces_from_file(xml_path)
        tree = ET.parse(xml_path)
        self.root = tree.getroot()

        # 获取表格
        self._build_tables_dict()

    def _build_tables_dict(self):
        """构建表格字典（使用与 get_tables_with_paths 相同的逻辑）"""
        body = self.root.find('.//w:body', NAMESPACES)
        if body is None:
            return

        w_p_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
        w_tbl_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl'

        # 使用字典存储各级标题，支持任意层级（outlineLvl 0-8）
        level_titles = {}
        prev_text = None

        elements = flatten_body_elements(body)

        for child in elements:
            if child.tag == w_p_tag:
                texts = child.findall('.//w:t', NAMESPACES)
                text = ''.join(t.text or '' for t in texts).strip()

                if not text:
                    continue

                # 检查是否包含忽略关键词
                is_skip = any(kw in text for kw in self.skip_keywords)
                if is_skip:
                    continue

                style = get_paragraph_style(child)

                # 检查样式的大纲级别
                if style and style in self.style_levels:
                    outline_lvl = self.style_levels[style]
                    # 设置当前级别的标题
                    level_titles[outline_lvl] = text
                    # 清除所有更低级别的标题（数字更大的级别）
                    for lvl in list(level_titles.keys()):
                        if lvl > outline_lvl:
                            del level_titles[lvl]

                prev_text = text

            elif child.tag == w_tbl_tag:
                if prev_text:
                    # 按级别顺序构建路径（从高到低，即数字从小到大）
                    path_parts = []
                    for lvl in sorted(level_titles.keys()):
                        title = level_titles[lvl]
                        if title:
                            path_parts.append(title)

                    # 如果 prev_text 不在已有路径中，添加到末尾
                    if prev_text and prev_text not in path_parts:
                        path_parts.append(prev_text)

                    if path_parts:
                        full_path = '/'.join(path_parts)
                        self.tables_dict[full_path] = child

                prev_text = None

    def fill(self, mappings, progress_callback=None):
        """
        执行填充

        参数:
            mappings: [(Word路径, Excel文件名, Sheet名), ...]
            progress_callback: 进度回调函数 callback(current, total, message)

        返回: {
            'success': int,
            'failed': int,
            'total_filled': int,
            'total_added': int,
            'total_removed': int,
            'empty_headers': [(Word路径, [(行号, 列号), ...]), ...],
            'errors': [(Word路径, error_message), ...]
        }
        """
        if self.root is None:
            self.prepare()

        result = {
            'success': 0,
            'failed': 0,
            'total_filled': 0,
            'total_added': 0,
            'total_removed': 0,
            'empty_headers': [],
            'errors': []
        }

        total = len(mappings)

        for idx, (word_path, excel_file, sheet_name) in enumerate(mappings):
            if progress_callback:
                progress_callback(idx + 1, total, f"填充: {word_path}")

            # 查找 Word 表格
            table = self.tables_dict.get(word_path)
            if table is None:
                result['failed'] += 1
                result['errors'].append((word_path, 'Word 表格未找到'))
                continue

            # 读取 Excel 数据
            if self.single_file_mode:
                # 单文件模式：excel_source 就是完整的文件路径
                excel_path = self.excel_source
            else:
                # 多文件模式：拼接文件夹和文件名
                excel_path = os.path.join(self.excel_source, excel_file)

            if not os.path.exists(excel_path):
                result['failed'] += 1
                result['errors'].append((word_path, f'Excel 文件未找到: {excel_file}'))
                continue

            try:
                # 计算要跳过的行数（从0开始，所以 start_row=1 时跳过0行）
                skiprows = self.data_start_row - 1 if self.data_start_row > 1 else 0

                # 读取 Excel
                df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, skiprows=skiprows)

                # 如果起始列不是第1列，截取从指定列开始的数据
                if self.data_start_col > 1:
                    start_col_idx = self.data_start_col - 1
                    if start_col_idx < len(df.columns):
                        df = df.iloc[:, start_col_idx:]
                        df.columns = range(len(df.columns))  # 重置列索引
            except Exception as e:
                result['failed'] += 1
                result['errors'].append((word_path, f'读取 Excel 失败: {e}'))
                continue

            # 填充
            try:
                filled, added, removed = fill_table_from_dataframe(table, df)
                result['success'] += 1
                result['total_filled'] += filled
                result['total_added'] += added
                result['total_removed'] += removed

                # 检查空标题
                empty_cells = check_empty_header_cells(table)
                if empty_cells:
                    result['empty_headers'].append((word_path, empty_cells))

            except Exception as e:
                result['failed'] += 1
                result['errors'].append((word_path, f'填充失败: {e}'))

        return result

    def save(self):
        """保存结果"""
        if self.work_dir is None:
            raise RuntimeError("请先调用 prepare() 方法")

        # 保存 XML
        xml_path = os.path.join(self.work_dir, 'word', 'document.xml')
        write_xml_preserving_wrapper(
            self.root,
            xml_path,
            xml_declaration=self.document_xml_declaration,
            root_start_tag=self.document_root_start_tag,
            root_end_tag=self.document_root_end_tag,
        )

        # 打包
        pack_docx(self.work_dir, self.output_path)

        return self.output_path

    def cleanup(self):
        """清理临时文件"""
        if self.work_dir and os.path.exists(self.work_dir):
            shutil.rmtree(self.work_dir, ignore_errors=True)
