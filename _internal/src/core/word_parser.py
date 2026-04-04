"""
Word XML 解析工具
- 解析 .docx 文件
- 提取表格信息
- 支持合并单元格、SDT 等复杂结构
"""

import xml.etree.ElementTree as ET
import os
import re
import shutil
import tempfile
from zipfile import ZipFile

# Word XML 命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

# 注册主命名空间（XPath 查询用）
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


XML_DECLARATION_RE = re.compile(r'^\ufeff?\s*(<\?xml[^>]*\?>)', re.S)
ROOT_START_TAG_RE = re.compile(r'^\ufeff?\s*(?:<\?xml[^>]*\?>\s*)?(<[^!?/][^>]*>)', re.S)
ROOT_END_TAG_RE = re.compile(r'(</[^>]+>\s*)$', re.S)


def register_namespaces_from_file(xml_path):
    """从 XML 文件中扫描所有命名空间声明并注册，避免输出 ns0, ns1 等前缀"""
    for _, (prefix, uri) in ET.iterparse(xml_path, events=['start-ns']):
        try:
            ET.register_namespace(prefix or '', uri)
        except ValueError:
            # 跳过 xml 等 ElementTree 内部保留前缀
            continue


def capture_xml_wrapper(xml_path):
    """
    读取原始 XML 头、根节点起始标签、根节点结束标签。

    用于在保存时把模板 document.xml 的根包装层原样带回去，
    避免 ElementTree 重序列化后调整 xmlns 声明、顺序或丢失未使用前缀。
    """
    with open(xml_path, 'r', encoding='utf-8') as f:
        xml_text = f.read()

    xml_declaration_match = XML_DECLARATION_RE.search(xml_text)
    root_start_tag_match = ROOT_START_TAG_RE.search(xml_text)
    root_end_tag_match = ROOT_END_TAG_RE.search(xml_text)

    xml_declaration = (
        xml_declaration_match.group(1)
        if xml_declaration_match else
        '<?xml version="1.0" encoding="UTF-8"?>'
    )
    root_start_tag = root_start_tag_match.group(1) if root_start_tag_match else None
    root_end_tag = root_end_tag_match.group(1).strip() if root_end_tag_match else None

    return xml_declaration, root_start_tag, root_end_tag


def write_xml_preserving_wrapper(root, xml_path, xml_declaration=None, root_start_tag=None, root_end_tag=None):
    """
    写回 XML，并保留模板原始根节点包装标签。

    说明：
    - ElementTree 负责生成修改后的 XML 主体；
    - 根节点起始/结束标签直接替换为模板原文，以保留 xmlns 声明、顺序和未使用前缀。
    """
    xml_body = ET.tostring(root, encoding='unicode')

    if root_start_tag:
        xml_body = re.sub(r'^<[^>]+>', root_start_tag, xml_body, count=1)
    if root_end_tag:
        xml_body = re.sub(r'</[^>]+>\s*$', root_end_tag, xml_body, count=1)

    xml_declaration = (xml_declaration or '<?xml version="1.0" encoding="UTF-8"?>')

    with open(xml_path, 'w', encoding='utf-8', newline='') as f:
        f.write(xml_declaration)
        f.write('\n')
        f.write(xml_body)


def extract_docx(docx_path, extract_dir=None):
    """
    解压 .docx 文件
    如果未指定 extract_dir，则创建临时目录
    返回: 解压目录路径
    """
    if extract_dir is None:
        extract_dir = tempfile.mkdtemp(prefix='word_filler_')

    with ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    return extract_dir


def pack_docx(source_dir, output_path):
    """
    将目录打包为 .docx 文件
    """
    if os.path.exists(output_path):
        os.remove(output_path)

    with ZipFile(output_path, 'w') as zipf:
        for root_dir, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, source_dir)
                zipf.write(file_path, arcname)


def get_paragraph_style(p):
    """获取段落的样式"""
    pPr = p.find('.//w:pPr', NAMESPACES)
    if pPr is not None:
        pStyle = pPr.find('.//w:pStyle', NAMESPACES)
        if pStyle is not None:
            return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
    return None


def flatten_body_elements(body):
    """
    展平 body 中的元素，包括 <w:sdt> 内部的段落和表格
    返回: [element, ...] 按文档顺序排列
    """
    w_p_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    w_tbl_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl'
    w_sdt_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'
    w_sdtContent_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent'

    result = []

    def extract_elements(parent):
        for child in parent:
            if child.tag == w_p_tag or child.tag == w_tbl_tag:
                result.append(child)
            elif child.tag == w_sdt_tag:
                sdt_content = child.find(w_sdtContent_tag)
                if sdt_content is not None:
                    extract_elements(sdt_content)
                else:
                    extract_elements(child)

    extract_elements(body)
    return result


def get_cell_text(cell):
    """获取单元格的文本内容"""
    texts = cell.findall('.//w:t', NAMESPACES)
    return ''.join(t.text or '' for t in texts)


def is_merged_cell(cell):
    """检查单元格是否是被合并的延续单元格"""
    tcPr = cell.find('.//w:tcPr', NAMESPACES)
    if tcPr is not None:
        vMerge = tcPr.find('.//w:vMerge', NAMESPACES)
        if vMerge is not None:
            val = vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if val is None or val == 'continue':
                return True
    return False


def detect_header_rows(rows):
    """检测表格的表头行数"""
    header_rows = 0

    for row_idx, row in enumerate(rows):
        cells = row.findall('.//w:tc', NAMESPACES)
        has_vmerge_continue = False

        for cell in cells:
            tcPr = cell.find('.//w:tcPr', NAMESPACES)
            if tcPr is not None:
                vMerge = tcPr.find('.//w:vMerge', NAMESPACES)
                if vMerge is not None:
                    val = vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    if val is None or val == 'continue':
                        has_vmerge_continue = True
                        break

        if has_vmerge_continue:
            header_rows = row_idx + 1
        elif row_idx == 0:
            header_rows = 1

    return header_rows


def get_row_first_cell_text(row):
    """获取行的第一个单元格的文本"""
    cells = row.findall('.//w:tc', NAMESPACES)
    if cells:
        return get_cell_text(cells[0]).strip()
    return ''


def get_table_info(table):
    """
    获取表格的结构信息
    返回: {
        'header_rows': int,
        'data_rows': int,
        'columns': int,
        'first_column': [str, ...]
    }
    """
    rows = table.findall('.//w:tr', NAMESPACES)
    if not rows:
        return {'header_rows': 0, 'data_rows': 0, 'columns': 0, 'first_column': []}

    header_row_count = detect_header_rows(rows)
    data_rows = rows[header_row_count:]

    # 获取列数（取第一行的单元格数）
    first_row_cells = rows[0].findall('.//w:tc', NAMESPACES)
    columns = len(first_row_cells)

    # 获取第一列内容
    first_column = []
    for row in data_rows:
        text = get_row_first_cell_text(row)
        first_column.append(text)

    return {
        'header_rows': header_row_count,
        'data_rows': len(data_rows),
        'columns': columns,
        'first_column': first_column
    }


def parse_styles(extract_dir):
    """
    解析 styles.xml，获取样式 ID 到大纲级别的映射

    返回: {styleId: outlineLvl, ...}
    - outlineLvl 是 0-based 的整数（0=一级标题, 1=二级标题, ...）

    层级识别策略（按优先级）：
    1. 样式定义中的 outlineLvl 属性
    2. 如果样式名是纯数字（如 "3", "4", "5"），使用该数字作为层级
    """
    styles_path = os.path.join(extract_dir, 'word', 'styles.xml')
    if not os.path.exists(styles_path):
        return {}

    try:
        tree = ET.parse(styles_path)
        root = tree.getroot()
    except Exception:
        return {}

    style_to_level = {}
    w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    # 查找所有样式定义
    for style in root.findall(f'.//{w_ns}style'):
        style_id = style.get(f'{w_ns}styleId')
        if not style_id:
            continue

        # 策略1：查找样式中的 outlineLvl
        pPr = style.find(f'{w_ns}pPr')
        if pPr is not None:
            outline_lvl = pPr.find(f'{w_ns}outlineLvl')
            if outline_lvl is not None:
                val = outline_lvl.get(f'{w_ns}val')
                if val is not None:
                    try:
                        style_to_level[style_id] = int(val)
                        continue  # 已找到 outlineLvl，跳过后续策略
                    except ValueError:
                        pass

        # 策略2：如果样式 ID 是纯数字，使用该数字作为层级（0-based）
        # 例如样式 ID "3" 对应层级 2（三级标题，0-based）
        if style_id.isdigit():
            level = int(style_id) - 1  # 转为 0-based
            if 0 <= level <= 8:
                style_to_level[style_id] = level

    return style_to_level


def parse_document(docx_path_or_dir):
    """
    解析 Word 文档
    接受 .docx 文件路径或已解压的目录路径
    返回: (root, xml_path, extract_dir, style_levels)
    - root: XML 根元素
    - xml_path: document.xml 路径
    - extract_dir: 解压目录（如果输入是 .docx 则为临时目录，否则为 None）
    - style_levels: 样式 ID 到大纲级别的映射
    """
    if os.path.isfile(docx_path_or_dir) and docx_path_or_dir.endswith('.docx'):
        extract_dir = extract_docx(docx_path_or_dir)
        xml_path = os.path.join(extract_dir, 'word', 'document.xml')
    else:
        extract_dir = None
        xml_path = os.path.join(docx_path_or_dir, 'word', 'document.xml')

    register_namespaces_from_file(xml_path)
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # 解析样式到层级的映射
    base_dir = extract_dir if extract_dir else docx_path_or_dir
    style_levels = parse_styles(base_dir)

    return root, xml_path, extract_dir, style_levels


DEFAULT_SKIP_KEYWORDS = ['单位：', '单位:', '币种：', '币种:', '其他说明']


def get_tables_with_paths(docx_path_or_dir, skip_keywords=None):
    """
    获取文档中所有表格的层级路径

    参数:
        docx_path_or_dir: .docx 文件路径或已解压的目录路径
        skip_keywords: 要忽略的关键词列表，包含这些关键词的行不会作为表格名称

    返回: [(序号, 路径, table_element), ...]

    层级识别逻辑:
        - 从 styles.xml 读取样式的 outlineLvl（大纲级别）
        - outlineLvl 0-8 对应 1-9 级标题
        - 根据大纲级别构建层级路径
    """
    if skip_keywords is None:
        skip_keywords = DEFAULT_SKIP_KEYWORDS

    root, xml_path, extract_dir, style_levels = parse_document(docx_path_or_dir)

    body = root.find('.//w:body', NAMESPACES)
    if body is None:
        return []

    result = []
    w_p_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    w_tbl_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl'

    # 使用字典存储各级标题，支持任意层级（outlineLvl 0-8）
    # key: outlineLvl (0-8), value: 标题文本
    level_titles = {}
    prev_text = None
    table_index = 0

    elements = flatten_body_elements(body)

    for child in elements:
        if child.tag == w_p_tag:
            texts = child.findall('.//w:t', NAMESPACES)
            text = ''.join(t.text or '' for t in texts).strip()

            if not text:
                continue

            is_skip = any(kw in text for kw in skip_keywords)
            if is_skip:
                continue

            style = get_paragraph_style(child)

            # 检查样式的大纲级别
            if style and style in style_levels:
                outline_lvl = style_levels[style]
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
                    table_index += 1
                    full_path = '/'.join(path_parts)
                    result.append((table_index, full_path, child))

            prev_text = None

    return result
