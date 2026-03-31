"""
步骤 2：编辑映射（类 Excel 表格编辑）
- 使用 tksheet 实现类 Excel 的编辑体验
- 支持拖拽、复制粘贴、删除行等操作
"""

import customtkinter as ctk
from tkinter import messagebox
import threading
import sys

try:
    from tksheet import Sheet
except ImportError:
    Sheet = None



from ..core.mapping import (
    get_word_tables,
    get_excel_sheets,
    generate_mapping,
    get_valid_mappings,
)
from ..utils.config import save_mapping, load_mapping, delete_mapping


class Step3MappingFrame(ctk.CTkFrame):
    """步骤 2：映射编辑（类 Excel 表格）"""

    def __init__(self, master, wizard, **kwargs):
        super().__init__(master, **kwargs)
        self.wizard = wizard
        self.configure(fg_color="transparent")

        self.mapping_data = []
        self.excel_folder = ''
        self.sheet = None
        self._column_widths_initialized = False  # 标记列宽是否已初始化

        # 可调节的字体大小和行高
        # Windows 默认使用第3大的设置（字体18，行高44）
        if sys.platform == 'win32':
            self._font_size = 18
            self._row_height = 44
        else:
            self._font_size = 11
            self._row_height = 25
        self._font_name = "Microsoft YaHei" if sys.platform == 'win32' else "Arial"

        self._create_widgets()

    def _create_widgets(self):
        """创建组件"""
        # 顶部工具栏
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            toolbar,
            text="匹配表（类 Excel 编辑）",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left")

        # 提示
        tip_label = ctk.CTkLabel(
            toolbar,
            text="提示：拖拽单元格内容修改匹配 | Ctrl+C/V 复制粘贴 | Delete 删除行 | 双击编辑",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray60")
        )
        tip_label.pack(side="left", padx=20)

        # 工具按钮
        btn_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_frame.pack(side="right")

        ctk.CTkButton(
            btn_frame,
            text="删除选中行",
            command=self._delete_selected_rows,
            width=90,
            fg_color="gray"
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_frame,
            text="清除选中匹配",
            command=self._clear_selected_mapping,
            width=100,
            fg_color="gray"
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_frame,
            text="自动匹配",
            command=self._auto_match,
            width=80,
            fg_color="green"
        ).pack(side="left", padx=2)

        # 搜索和筛选栏
        filter_frame = ctk.CTkFrame(self, fg_color="transparent")
        filter_frame.pack(fill="x", pady=(0, 5))

        # Windows 上添加缩放按钮（解决 DPI 问题）
        if sys.platform == 'win32':
            ctk.CTkButton(
                filter_frame,
                text="A-",
                command=self._zoom_out,
                width=38,
                fg_color="#2563eb",
                hover_color="#1d4ed8"
            ).pack(side="left", padx=(0, 2))

            ctk.CTkButton(
                filter_frame,
                text="A+",
                command=self._zoom_in,
                width=38,
                fg_color="#2563eb",
                hover_color="#1d4ed8"
            ).pack(side="left", padx=(0, 15))

        ctk.CTkLabel(filter_frame, text="搜索:").pack(side="left")
        self.search_entry = ctk.CTkEntry(filter_frame, width=150)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", self._on_search)

        ctk.CTkButton(
            filter_frame, text="搜索", command=self._on_search, width=60
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            filter_frame, text="重置", command=self._reset_filter, width=60
        ).pack(side="left", padx=2)

        # 筛选
        self.filter_var = ctk.StringVar(value="all")
        ctk.CTkLabel(filter_frame, text="    筛选:").pack(side="left", padx=(20, 5))

        for text, value in [("全部", "all"), ("已匹配", "matched"), ("未匹配", "unmatched"), ("Excel未用", "unused")]:
            ctk.CTkRadioButton(
                filter_frame, text=text, variable=self.filter_var,
                value=value, command=self._apply_filter
            ).pack(side="left", padx=3)

        # 重置匹配按钮
        ctk.CTkButton(
            filter_frame, text="清空全部匹配", command=self._clear_all_mapping,
            width=100, fg_color="gray", hover_color="darkgray"
        ).pack(side="right", padx=5)

        ctk.CTkButton(
            filter_frame, text="重置匹配", command=self._reset_mapping,
            width=80, fg_color="gray", hover_color="darkgray"
        ).pack(side="right", padx=5)

        # 表格区域
        self.sheet_frame = ctk.CTkFrame(self)
        self.sheet_frame.pack(fill="both", expand=True, pady=5)

        # 创建 tksheet
        # 注意：第5列(索引4)是隐藏的完整路径列，用于保存原始Word路径
        if Sheet:
            self.sheet = Sheet(
                self.sheet_frame,
                headers=["Word表格路径", "Excel文件名", "Sheet名", "匹配状态", "_full_path"],
                show_row_index=True,
                row_index_width=50,
                font=(self._font_name, self._font_size, "normal"),
                header_font=(self._font_name, self._font_size, "bold"),
                default_row_height=self._row_height,
            )
            self.sheet.pack(fill="both", expand=True)

            # 设置列宽
            self._set_column_widths()

            # 启用功能
            self.sheet.enable_bindings((
                "single_select",
                "row_select",
                "column_select",
                "drag_select",
                "select_all",
                "copy",
                "paste",
                "delete",
                "undo",
                "edit_cell",
                "row_drag_and_drop",
                "column_width_resize",
                "row_height_resize",
                "arrowkeys",
            ))

            # 绑定事件
            self.sheet.bind("<<SheetModified>>", self._on_sheet_modified)
        else:
            # 如果没有 tksheet，显示提示
            ctk.CTkLabel(
                self.sheet_frame,
                text="请安装 tksheet: pip install tksheet",
                font=ctk.CTkFont(size=14)
            ).pack(expand=True)

        # 底部统计
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(fill="x", pady=(5, 0))

        self.stats_label = ctk.CTkLabel(
            bottom_frame,
            text="",
            font=ctk.CTkFont(size=12)
        )
        self.stats_label.pack(side="left")

        # 加载提示
        self.loading_label = ctk.CTkLabel(
            self,
            text="正在加载...",
            font=ctk.CTkFont(size=16)
        )

    def _set_column_widths(self):
        """设置默认列宽（仅首次）"""
        if not self.sheet or self._column_widths_initialized:
            return

        # Windows 下使用更大的列宽（配合更大的字体）
        if sys.platform == 'win32':
            self.sheet.column_width(column=0, width=500)
            self.sheet.column_width(column=1, width=260)
            self.sheet.column_width(column=2, width=260)
            self.sheet.column_width(column=3, width=140)
        else:
            self.sheet.column_width(column=0, width=350)
            self.sheet.column_width(column=1, width=180)
            self.sheet.column_width(column=2, width=180)
            self.sheet.column_width(column=3, width=100)
        # 隐藏第5列（完整路径列）- 设置宽度为0
        self.sheet.column_width(column=4, width=0)
        self._column_widths_initialized = True

    def on_show(self):
        """显示时调用"""
        state = self.wizard.state

        # 检查文件路径和选项是否改变
        current_docx = state.get('docx_path', '')
        current_excel = state.get('excel_folder', '')
        current_filter = state.get('use_color_filter', True)
        current_mode = state.get('excel_mode', 'folder')
        current_separator = state.get('prefix_separator', '')
        current_skip_keywords = state.get('word_skip_keywords', [])

        last_docx = state.get('_last_docx_path', '')
        last_excel = state.get('_last_excel_folder', '')
        last_filter = state.get('_last_color_filter', True)
        last_mode = state.get('_last_excel_mode', 'folder')
        last_separator = state.get('_last_separator', '')
        last_skip_keywords = state.get('_last_skip_keywords', [])

        paths_changed = (current_docx != last_docx) or (current_excel != last_excel)
        filter_changed = (current_filter != last_filter)
        mode_changed = (current_mode != last_mode)
        rules_changed = (current_separator != last_separator) or (current_skip_keywords != last_skip_keywords)

        # 只在以下情况重新加载：没有映射数据、路径改变、筛选选项改变、模式改变、规则改变、或明确要求刷新
        need_reload = (not state.get('mapping_data') or paths_changed or filter_changed or
                       mode_changed or rules_changed or state.get('_mapping_needs_refresh', False))

        if need_reload:
            state['_last_docx_path'] = current_docx
            state['_last_excel_folder'] = current_excel
            state['_last_color_filter'] = current_filter
            state['_last_excel_mode'] = current_mode
            state['_last_separator'] = current_separator
            state['_last_skip_keywords'] = current_skip_keywords
            # 如果规则改变，需要强制重新生成映射（不使用保存的映射）
            state['_force_regenerate'] = rules_changed
            self._load_data()
        else:
            self.mapping_data = state.get('mapping_data', [])
            self.excel_folder = state.get('excel_folder', '')
            self._refresh_sheet()

    def _load_data(self):
        """加载数据"""
        self.loading_label.place(relx=0.5, rely=0.5, anchor="center")
        self.wizard.set_next_enabled(False)

        def load_task():
            state = self.wizard.state
            docx_path = state['docx_path']
            excel_folder = state['excel_folder']
            use_color_filter = state.get('use_color_filter', True)
            excel_mode = state.get('excel_mode', 'folder')
            single_file_mode = (excel_mode == 'single')
            prefix_separator = state.get('prefix_separator', '')
            skip_keywords = state.get('word_skip_keywords', None)

            word_tables = get_word_tables(docx_path, skip_keywords=skip_keywords)
            excel_sheets = get_excel_sheets(excel_folder, use_color_filter=use_color_filter)

            # 检查是否需要强制重新生成（规则改变时）
            force_regenerate = state.get('_force_regenerate', False)
            state['_force_regenerate'] = False  # 重置标志

            # 尝试加载已保存的映射（除非强制重新生成）
            saved_mapping = None if force_regenerate else load_mapping(docx_path, excel_folder)

            if saved_mapping:
                # 使用已保存的映射
                mapping_data = saved_mapping
            else:
                # 生成新映射
                word_start_level = state.get('word_start_level', 1)
                mapping_data = generate_mapping(
                    word_tables, excel_sheets,
                    single_file_mode=single_file_mode,
                    prefix_separator=prefix_separator,
                    word_start_level=word_start_level
                )

            state['word_tables'] = word_tables
            state['excel_sheets'] = excel_sheets
            state['mapping_data'] = mapping_data
            state['_mapping_needs_refresh'] = False

            self.after(0, lambda: self._on_data_loaded(mapping_data, excel_folder))

        threading.Thread(target=load_task, daemon=True).start()

    def _on_data_loaded(self, mapping_data, excel_folder):
        """数据加载完成"""
        self.loading_label.place_forget()
        self.mapping_data = mapping_data
        self.excel_folder = excel_folder
        self._refresh_sheet()
        self.wizard.set_next_enabled(True)

    def _refresh_sheet(self, filter_text='', filter_status='all'):
        """刷新表格"""
        if not self.sheet:
            return

        # 获取 Step 2 设置的起始层级
        word_start_level = self.wizard.state.get('word_start_level', 1)

        # 构建表格数据
        rows = []
        current_levels = []  # 当前各级类别 [一级, 二级, ...]

        for item in self.mapping_data:
            word_path = item.get('Word表格路径', '')
            excel_file = item.get('Excel文件名', '')
            sheet_name = item.get('Sheet名', '')
            status = item.get('匹配状态', '')

            # 获取层级路径（应用起始层级后的路径）
            if word_path and word_path != '(无对应Word表格)':
                # 应用起始层级截取路径
                path_parts = word_path.split('/')
                skip = word_start_level - 1
                if skip > 0 and skip < len(path_parts):
                    path_parts = path_parts[skip:]

                # 第一级作为分类，剩余作为表格名
                if len(path_parts) > 1:
                    categories = [path_parts[0]]  # 第一级作为分类
                    display_name = '/'.join(path_parts[1:])  # 剩余部分作为表格名
                else:
                    categories = []
                    display_name = path_parts[0] if path_parts else word_path
            elif excel_file:
                categories = [excel_file.replace('.xlsx', '')]
                display_name = word_path if word_path else '(无对应Word表格)'
            else:
                categories = ['其他']
                display_name = word_path if word_path else '(无对应Word表格)'

            # 搜索过滤
            if filter_text:
                filter_lower = filter_text.lower()
                if (filter_lower not in word_path.lower() and
                    filter_lower not in excel_file.lower() and
                    filter_lower not in sheet_name.lower()):
                    continue

            # 状态过滤
            is_matched = excel_file and sheet_name and word_path != '(无对应Word表格)'
            is_unused = word_path == '(无对应Word表格)'

            if filter_status == 'matched' and not is_matched:
                continue
            if filter_status == 'unmatched' and (is_matched or is_unused):
                continue
            if filter_status == 'unused' and not is_unused:
                continue

            # 添加类别分隔行（根据选择的层级深度）
            for level, category in enumerate(categories):
                # 检查这一级是否变化了
                if level >= len(current_levels) or current_levels[level] != category:
                    # 更新当前层级
                    current_levels = current_levels[:level]  # 截断到当前级别
                    current_levels.append(category)
                    # 添加分隔行，用缩进表示层级（第5列也留空）
                    indent = "  " * level
                    rows.append([f"━━━ {indent}{category} ━━━", "", "", "", ""])

            # 添加数据行（第5列存储完整路径，用于后续恢复）
            rows.append([display_name, excel_file, sheet_name, status, word_path])

        # 更新表格（不重置列宽）
        self.sheet.set_sheet_data(rows, reset_col_positions=False)
        self._set_column_widths()  # 只在首次设置默认列宽

        self._update_stats()

    def _update_stats(self):
        """更新统计"""
        matched = 0
        unmatched = 0
        unused = 0

        for item in self.mapping_data:
            word_path = item.get('Word表格路径', '')
            excel_file = item.get('Excel文件名', '')
            sheet_name = item.get('Sheet名', '')

            if word_path == '(无对应Word表格)':
                unused += 1
            elif excel_file and sheet_name:
                matched += 1
            else:
                unmatched += 1

        self.stats_label.configure(
            text=f"Word表格: {matched + unmatched} | 已匹配: {matched} | 未匹配: {unmatched} | Excel未用: {unused}"
        )

    def _on_search(self, event=None):
        """搜索"""
        self._apply_filter()

    def _reset_filter(self):
        """重置筛选"""
        self.search_entry.delete(0, "end")
        self.filter_var.set("all")
        self._refresh_sheet()

    def _reset_mapping(self):
        """重置匹配（清空手工匹配，重新自动生成）"""
        result = messagebox.askyesno(
            "确认重置",
            "确定要重置匹配吗？\n\n这将清空所有手工匹配，重新使用自动匹配生成匹配表。"
        )

        if not result:
            return

        state = self.wizard.state
        docx_path = state.get('docx_path', '')
        excel_folder = state.get('excel_folder', '')

        if not docx_path or not excel_folder:
            return

        # 删除保存的映射文件
        delete_mapping(docx_path, excel_folder)

        # 重新生成映射
        word_tables = state.get('word_tables', [])
        excel_sheets = state.get('excel_sheets', [])
        excel_mode = state.get('excel_mode', 'folder')
        single_file_mode = (excel_mode == 'single')
        prefix_separator = state.get('prefix_separator', '')
        word_start_level = state.get('word_start_level', 1)

        if word_tables and excel_sheets:
            mapping_data = generate_mapping(
                word_tables, excel_sheets,
                single_file_mode=single_file_mode,
                prefix_separator=prefix_separator,
                word_start_level=word_start_level
            )
            state['mapping_data'] = mapping_data
            self.mapping_data = mapping_data
            self._refresh_sheet()
            messagebox.showinfo("完成", "匹配已重置为自动匹配结果。")

    def _clear_all_mapping(self):
        """清空全部匹配（将所有 Excel 映射清空，Excel 列为未使用状态，按大类分组）"""
        result = messagebox.askyesno(
            "确认清空",
            "确定要清空全部匹配吗？\n\n这将清空所有匹配关系，Word 表格变为未匹配，Excel 工作表变为未使用。"
        )

        if not result:
            return

        state = self.wizard.state
        docx_path = state.get('docx_path', '')
        excel_folder = state.get('excel_folder', '')
        excel_sheets = state.get('excel_sheets', [])

        # 删除保存的映射文件
        if docx_path and excel_folder:
            delete_mapping(docx_path, excel_folder)

        # 按大类分组
        import os
        groups = {}  # {大类: {'word_items': [], 'excel_items': []}}
        level1_order = []

        # 1. 收集 Word 表格，按大类分组
        for item in self.mapping_data:
            word_path = item.get('Word表格路径', '')
            if word_path and word_path != '(无对应Word表格)':
                level1 = word_path.split('/')[0] if '/' in word_path else word_path
                if level1 not in groups:
                    groups[level1] = {'word_items': [], 'excel_items': []}
                    level1_order.append(level1)
                groups[level1]['word_items'].append({
                    '序号': item.get('序号', ''),
                    'Word表格路径': word_path,
                    'Excel文件名': '',
                    'Sheet名': '',
                    '匹配分数': 0,
                    '匹配状态': '未匹配'
                })

        # 2. 收集 Excel Sheet，按文件名对应的大类分组
        for excel_file, sheet_name in excel_sheets:
            # Excel 文件名（去掉扩展名）作为大类
            level1 = os.path.splitext(excel_file)[0]
            if level1 not in groups:
                groups[level1] = {'word_items': [], 'excel_items': []}
                level1_order.append(level1)
            groups[level1]['excel_items'].append({
                '序号': '',
                'Word表格路径': '(无对应Word表格)',
                'Excel文件名': excel_file,
                'Sheet名': sheet_name,
                '匹配分数': 0,
                '匹配状态': 'Excel未使用'
            })

        # 3. 按大类顺序组合结果
        new_mapping_data = []
        for level1 in level1_order:
            if level1 in groups:
                new_mapping_data.extend(groups[level1]['word_items'])
                new_mapping_data.extend(groups[level1]['excel_items'])

        self.mapping_data = new_mapping_data
        state['mapping_data'] = new_mapping_data
        self._refresh_sheet()
        messagebox.showinfo("完成", "已清空全部匹配。")

    def _apply_filter(self):
        """应用筛选"""
        filter_text = self.search_entry.get()
        filter_status = self.filter_var.get()
        self._refresh_sheet(filter_text, filter_status)

    def _on_sheet_modified(self, event=None):
        """表格被修改时"""
        # 从表格同步数据回 mapping_data
        self._sync_from_sheet()

    def _sync_from_sheet(self):
        """从表格同步数据"""
        if not self.sheet:
            return

        data = self.sheet.get_sheet_data()

        # 重建 mapping_data
        new_mapping = []
        idx = 1

        for row_idx, row in enumerate(data):
            if len(row) < 4:
                continue

            display_name = str(row[0]).strip()
            excel_file = str(row[1]).strip()
            sheet_name = str(row[2]).strip()
            status = str(row[3]).strip()

            # 跳过类别行
            if display_name.startswith("━━━"):
                continue

            # 从第5列恢复完整路径（如果有）
            if len(row) >= 5 and row[4]:
                word_path = str(row[4]).strip()
            else:
                # 如果没有第5列，使用显示名称（兼容旧数据）
                word_path = display_name

            # 确定状态
            if word_path == '(无对应Word表格)':
                status = 'Excel未使用'
            elif excel_file and sheet_name:
                if status not in ['已匹配', '手工匹配']:
                    status = '手工匹配'
            else:
                status = '未匹配'

            new_mapping.append({
                '序号': idx,
                'Word表格路径': word_path,
                'Excel文件名': excel_file,
                'Sheet名': sheet_name,
                '匹配状态': status
            })
            idx += 1

        self.mapping_data = new_mapping
        self._update_stats()

    def _delete_selected_rows(self):
        """删除选中的行"""
        if not self.sheet:
            return

        selected = self.sheet.get_selected_rows()
        if not selected:
            messagebox.showinfo("提示", "请先选中要删除的行")
            return

        # 获取当前数据
        data = self.sheet.get_sheet_data()

        # 检查是否要删除类别行
        rows_to_delete = sorted(selected, reverse=True)
        for row_idx in rows_to_delete:
            if row_idx < len(data):
                word_path = data[row_idx][0] if data[row_idx] else ""
                if word_path.startswith("━━━"):
                    messagebox.showwarning("警告", "不能删除类别分隔行")
                    return

        # 删除行
        for row_idx in rows_to_delete:
            if row_idx < len(data):
                data.pop(row_idx)

        self.sheet.set_sheet_data(data, reset_col_positions=False)
        self._sync_from_sheet()

    def _clear_selected_mapping(self):
        """清除选中行的映射（保留行，只清空 Excel 列）"""
        if not self.sheet:
            return

        selected = self.sheet.get_selected_rows()
        if not selected:
            messagebox.showinfo("提示", "请先选中要清除匹配的行")
            return

        data = self.sheet.get_sheet_data()

        for row_idx in selected:
            if row_idx < len(data):
                word_path = data[row_idx][0] if len(data[row_idx]) > 0 else ""
                # 跳过类别行和 Excel 未用行
                if word_path.startswith("━━━") or word_path == '(无对应Word表格)':
                    continue
                # 清空 Excel 文件名和 Sheet 名
                data[row_idx][1] = ""
                data[row_idx][2] = ""
                data[row_idx][3] = "未匹配"

        self.sheet.set_sheet_data(data, reset_col_positions=False)
        self._sync_from_sheet()

    def _auto_match(self):
        """重新自动匹配"""
        state = self.wizard.state
        word_tables = state.get('word_tables', [])
        excel_sheets = state.get('excel_sheets', [])
        excel_mode = state.get('excel_mode', 'folder')
        single_file_mode = (excel_mode == 'single')
        prefix_separator = state.get('prefix_separator', '')
        word_start_level = state.get('word_start_level', 1)

        # 保留手工映射
        manual_mappings = {}
        for item in self.mapping_data:
            if item.get('匹配状态') == '手工匹配':
                word_path = item.get('Word表格路径', '')
                excel_file = item.get('Excel文件名', '')
                sheet_name = item.get('Sheet名', '')
                if word_path and excel_file and sheet_name and word_path != '(无对应Word表格)':
                    manual_mappings[word_path] = (excel_file, sheet_name)

        # 重新生成
        self.mapping_data = generate_mapping(
            word_tables, excel_sheets, existing_mappings=manual_mappings,
            single_file_mode=single_file_mode,
            prefix_separator=prefix_separator,
            word_start_level=word_start_level
        )

        state['mapping_data'] = self.mapping_data
        self._refresh_sheet()

        matched = sum(1 for item in self.mapping_data
                     if item.get('Excel文件名') and item.get('Sheet名')
                     and item.get('Word表格路径') != '(无对应Word表格)')
        messagebox.showinfo("完成", f"自动匹配完成，已匹配 {matched} 个表格")

    def validate(self):
        """验证"""
        self._sync_from_sheet()
        valid_mappings = get_valid_mappings(self.mapping_data)
        if not valid_mappings:
            messagebox.showwarning("警告", "没有有效的匹配关系，请至少配置一个匹配")
            return False
        return True

    def save_state(self):
        """保存状态"""
        self._sync_from_sheet()
        self.wizard.state['mapping_data'] = self.mapping_data

        # 持久化保存映射到文件
        state = self.wizard.state
        docx_path = state.get('docx_path', '')
        excel_folder = state.get('excel_folder', '')
        if docx_path and excel_folder:
            save_mapping(docx_path, excel_folder, self.mapping_data)

    def _zoom_in(self):
        """放大表格（增大字体和行高）"""
        if not self.sheet:
            return
        self._font_size = min(self._font_size + 1, 20)  # 最大 20
        self._row_height = min(self._row_height + 3, 50)  # 最大 50
        self._apply_zoom()

    def _zoom_out(self):
        """缩小表格（减小字体和行高）"""
        if not self.sheet:
            return
        self._font_size = max(self._font_size - 1, 8)  # 最小 8
        self._row_height = max(self._row_height - 3, 16)  # 最小 16
        self._apply_zoom()

    def _apply_zoom(self):
        """应用缩放设置"""
        if not self.sheet:
            return
        # 更新字体
        self.sheet.font((self._font_name, self._font_size, "normal"))
        self.sheet.header_font((self._font_name, self._font_size, "bold"))
        # 更新所有行高
        self.sheet.set_all_row_heights(self._row_height)
        # 刷新显示
        self.sheet.redraw()
