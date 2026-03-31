"""
核心模块
- word_parser: Word XML 解析
- mapping: 映射表生成与管理
- filler: 表格填充逻辑
"""

from .word_parser import (
    extract_docx,
    pack_docx,
    get_tables_with_paths,
    get_table_info,
)

from .mapping import (
    get_word_tables,
    get_word_tables_with_info,
    get_excel_sheets,
    get_excel_files,
    get_sheets_of_file,
    preview_excel_data,
    generate_mapping,
    save_mapping_to_excel,
    load_mapping_from_excel,
    get_valid_mappings,
)

from .filler import (
    DocumentFiller,
    fill_table_from_dataframe,
)

__all__ = [
    'extract_docx',
    'pack_docx',
    'get_tables_with_paths',
    'get_table_info',
    'get_word_tables',
    'get_word_tables_with_info',
    'get_excel_sheets',
    'get_excel_files',
    'get_sheets_of_file',
    'preview_excel_data',
    'generate_mapping',
    'save_mapping_to_excel',
    'load_mapping_from_excel',
    'get_valid_mappings',
    'DocumentFiller',
    'fill_table_from_dataframe',
]
