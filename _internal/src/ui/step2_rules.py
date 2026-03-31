"""
步骤 2：映射规则
- 显示 Word 表格路径和 Excel 文件名对比
- 让用户配置忽略前缀的分隔符
- 自动检测 Word 表格路径的公共前缀层级
- 实时预览处理效果
"""

import customtkinter as ctk
from tkinter import StringVar

from ..core.mapping import get_word_tables, get_excel_sheets, normalize_for_match
from ..utils.config import load_config, save_config


# 常用分隔符
COMMON_SEPARATORS = ["、", ".", "_", " ", "-", "）", ")"]


def remove_prefix_by_separator(text, separator):
    """
    根据分隔符移除前缀

    参数:
        text: 原始文本
        separator: 分隔符

    返回:
        处理后的文本
    """
    if not text or not separator:
        return text

    # 找到第一个分隔符的位置
    pos = text.find(separator)
    if pos != -1:
        # 返回分隔符之后的内容
        return text[pos + len(separator):].strip()

    return text


def detect_separator(names):
    """
    自动检测常用分隔符
    返回检测到的分隔符，如果没有检测到返回空字符串
    """
    if not names:
        return ''

    # 统计每个分隔符出现的次数
    separator_counts = {}
    for sep in COMMON_SEPARATORS:
        count = 0
        for name in names:
            # 检查分隔符是否出现在名称中（且不是在末尾，避免匹配扩展名的点）
            pos = name.find(sep)
            if pos > 0 and pos < len(name) - 1:  # 分隔符在中间
                count += 1
        if count > 0:
            separator_counts[sep] = count

    if not separator_counts:
        return ''

    # 返回出现次数最多且超过50%的分隔符
    best_sep = max(separator_counts, key=separator_counts.get)
    if separator_counts[best_sep] >= len(names) * 0.5:
        return best_sep

    return ''


class Step2RulesFrame(ctk.CTkFrame):
    """步骤 2：映射规则配置"""

    def __init__(self, master, wizard, **kwargs):
        super().__init__(master, **kwargs)
        self.wizard = wizard
        self.configure(fg_color="transparent")

        # 加载配置
        config = load_config()
        self.separator_var = StringVar(value=config.get('prefix_separator', ''))

        # Word 忽略关键词（默认值）
        default_skip = '单位：,单位:,币种：,币种:,其他说明'
        self.skip_keywords_var = StringVar(value=config.get('word_skip_keywords', default_skip))

        # Word 表格路径起始层级（自动检测）
        self._detected_start_level = 1

        self.word_paths = []  # Word 完整路径列表
        self.excel_paths = []  # Excel 完整路径列表
        self._raw_excel_names = []  # 原始 Excel 名称（未应用分隔符）
        self._max_word_level = 1  # Word 路径的最大层级数
        self._refresh_job = None  # 延迟刷新任务

        # 用于检测路径变化
        self._last_docx_path = ''
        self._last_excel_source = ''
        self._last_excel_mode = ''
        self._last_color_filter = True

        self._create_widgets()

    def _create_widgets(self):
        """创建组件"""
        # 说明
        intro_label = ctk.CTkLabel(
            self,
            text="配置命名规则，帮助自动匹配 Word 表格和 Excel 数据",
            font=ctk.CTkFont(size=16)
        )
        intro_label.pack(pady=(10, 10))

        # 主内容区域（可滚动）
        content = ctk.CTkScrollableFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=30)

        # ========== 第一部分：Word 表格名称提取规则 ==========
        word_rules_frame = ctk.CTkFrame(content)
        word_rules_frame.pack(fill="x", pady=(0, 10))

        # 标题行（包含红色提示）
        title_row = ctk.CTkFrame(word_rules_frame, fg_color="transparent")
        title_row.pack(anchor="w", padx=15, pady=(10, 5))

        ctk.CTkLabel(
            title_row,
            text="Word 表格名称提取规则：",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        ctk.CTkLabel(
            title_row,
            text="根据表格上一行段落的文字确定",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#CC0000", "#FF4444")  # 红色
        ).pack(side="left")

        # 规则说明
        word_tip = ctk.CTkLabel(
            word_rules_frame,
            text="如果上方文字包含以下关键词，则会被忽略（跳到更上一行）",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray60"),
            wraplength=700,
            justify="left"
        )
        word_tip.pack(anchor="w", padx=15, pady=(0, 5))

        # 忽略关键词输入行
        skip_row = ctk.CTkFrame(word_rules_frame, fg_color="transparent")
        skip_row.pack(fill="x", padx=15, pady=5)

        ctk.CTkLabel(
            skip_row,
            text="忽略包含这些关键词的行（多个关键词间请用英文逗号「,」分隔）:",
            anchor="w"
        ).pack(side="left")

        self.skip_entry = ctk.CTkEntry(
            skip_row,
            width=350,
            textvariable=self.skip_keywords_var
        )
        self.skip_entry.pack(side="left", padx=10)
        self.skip_entry.bind("<KeyRelease>", self._on_rule_change)

        # 恢复默认按钮
        ctk.CTkButton(
            skip_row,
            text="恢复默认",
            command=self._reset_skip_keywords,
            width=80,
            fg_color="gray"
        ).pack(side="left", padx=5)

        # 显示默认值提示
        default_tip = ctk.CTkLabel(
            word_rules_frame,
            text="默认关键词: 单位：, 单位:, 币种：, 币种:, 其他说明",
            font=ctk.CTkFont(size=10),
            text_color=("gray50", "gray50")
        )
        default_tip.pack(anchor="w", padx=15, pady=(0, 10))

        # ========== 第二部分：Excel 前缀分隔符配置 ==========
        rules_frame = ctk.CTkFrame(content)
        rules_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            rules_frame,
            text="Excel 前缀分隔符配置",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=15, pady=(10, 5))

        # 分隔符配置行
        sep_row = ctk.CTkFrame(rules_frame, fg_color="transparent")
        sep_row.pack(fill="x", padx=15, pady=5)

        ctk.CTkLabel(
            sep_row,
            text="忽略此分隔符之前的内容:",
            width=180,
            anchor="e"
        ).pack(side="left")

        self.sep_entry = ctk.CTkEntry(
            sep_row,
            width=60,
            textvariable=self.separator_var,
            placeholder_text="如 、"
        )
        self.sep_entry.pack(side="left", padx=10)
        self.sep_entry.bind("<KeyRelease>", self._on_rule_change)

        # 常用分隔符按钮
        ctk.CTkLabel(sep_row, text="常用:").pack(side="left", padx=(20, 5))

        for sep in COMMON_SEPARATORS:
            display = sep if sep != " " else "空格"
            btn = ctk.CTkButton(
                sep_row,
                text=display,
                command=lambda s=sep: self._set_separator(s),
                width=45,
                height=28,
                fg_color="gray"
            )
            btn.pack(side="left", padx=2)

        # 自动检测按钮
        self.detect_btn = ctk.CTkButton(
            sep_row,
            text="自动检测",
            command=self._auto_detect_separator,
            width=80
        )
        self.detect_btn.pack(side="left", padx=(15, 5))

        # 清空按钮
        ctk.CTkButton(
            sep_row,
            text="清空",
            command=lambda: self._set_separator(''),
            width=50,
            fg_color="gray"
        ).pack(side="left", padx=2)

        # 说明
        tip_label = ctk.CTkLabel(
            rules_frame,
            text="例如：Excel 名称为「1、货币资金」，设置分隔符为「、」后，将匹配 Word 中的「货币资金」",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray60")
        )
        tip_label.pack(anchor="w", padx=15, pady=(5, 10))

        # ========== 第三部分：预览区域 ==========
        preview_frame = ctk.CTkFrame(content)
        preview_frame.pack(fill="both", expand=True, pady=10)

        ctk.CTkLabel(
            preview_frame,
            text="预览效果（根据上方配置实时更新）",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=15, pady=(10, 5))

        # 名称对比
        compare_frame = ctk.CTkFrame(preview_frame, fg_color="transparent")
        compare_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Word 名称列表
        word_frame = ctk.CTkFrame(compare_frame)
        word_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.word_list_label = ctk.CTkLabel(
            word_frame,
            text="Word 表格路径",
            font=ctk.CTkFont(size=12)
        )
        self.word_list_label.pack(pady=(5, 5))

        self.word_listbox = ctk.CTkTextbox(word_frame, width=300, height=150)
        self.word_listbox.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        # Excel 名称列表
        excel_frame = ctk.CTkFrame(compare_frame)
        excel_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.excel_list_label = ctk.CTkLabel(
            excel_frame,
            text="Excel 路径",
            font=ctk.CTkFont(size=12)
        )
        self.excel_list_label.pack(pady=(5, 5))

        self.excel_listbox = ctk.CTkTextbox(excel_frame, width=300, height=150)
        self.excel_listbox.pack(fill="both", expand=True, padx=5, pady=(0, 5))

    def _set_separator(self, sep):
        """设置分隔符"""
        self.separator_var.set(sep)
        self._refresh_preview()

    def _reset_skip_keywords(self):
        """恢复默认忽略关键词"""
        default_skip = '单位：,单位:,币种：,币种:,其他说明'
        self.skip_keywords_var.set(default_skip)
        self._refresh_preview()

    def _auto_detect_separator(self):
        """自动检测分隔符"""
        detected = detect_separator(self._raw_excel_names)
        if detected:
            self.separator_var.set(detected)
            self._refresh_preview()

    def _on_rule_change(self, event=None):
        """规则变化时更新预览"""
        # 使用 after 延迟执行，避免频繁刷新
        if hasattr(self, '_refresh_job') and self._refresh_job:
            self.after_cancel(self._refresh_job)
        self._refresh_job = self.after(300, self._refresh_preview)

    def _refresh_preview(self):
        """刷新预览（重新加载数据）"""
        self._load_names()

    def _detect_start_level(self, word_paths):
        """
        自动检测 Word 路径的起始层级

        逻辑：找出所有路径的公共前缀层数，跳过这些公共层

        例如：
        - 合并财务报表项目注释/货币资金
        - 合并财务报表项目注释/交易性金融资产
        - 合并财务报表项目注释/其他应收款/应收利息

        第1层全部相同("合并财务报表项目注释") → 公共前缀，跳过
        第2层不同 → 起始层级 = 2
        """
        if not word_paths:
            return 1, []

        # 将所有路径拆分为层级列表
        all_parts = [path.split('/') for path in word_paths]
        min_depth = min(len(parts) for parts in all_parts)

        # 找出公共前缀层数
        common_prefix_levels = []
        for level in range(min_depth):
            # 获取这一层的所有值
            level_values = set(parts[level] for parts in all_parts if len(parts) > level)
            if len(level_values) == 1:
                # 这一层全部相同，是公共前缀
                common_prefix_levels.append(list(level_values)[0])
            else:
                # 这一层不同，停止
                break

        # 起始层级 = 公共前缀层数 + 1
        start_level = len(common_prefix_levels) + 1
        return start_level, common_prefix_levels

    def _get_start_level(self):
        """获取当前的起始层级（1-based）"""
        return self._detected_start_level if hasattr(self, '_detected_start_level') else 1

    def _apply_start_level(self, path, start_level):
        """根据起始层级截取路径"""
        if not path:
            return path
        parts = path.split('/')
        # start_level 是 1-based，所以跳过前 (start_level - 1) 层
        skip = start_level - 1
        if skip >= len(parts):
            return parts[-1]  # 如果层级超出，返回最后一层
        return '/'.join(parts[skip:])

    def _load_names(self):
        """加载 Word 和 Excel 名称并应用规则"""
        state = self.wizard.state
        docx_path = state.get('docx_path', '')
        excel_source = state.get('excel_folder', '')
        use_color_filter = state.get('use_color_filter', True)
        excel_mode = state.get('excel_mode', 'folder')

        if not docx_path or not excel_source:
            self.word_listbox.delete("1.0", "end")
            self.word_listbox.insert("1.0", "(请先在步骤1选择文件)")
            self.excel_listbox.delete("1.0", "end")
            self.excel_listbox.insert("1.0", "(请先在步骤1选择文件)")
            return

        # 获取当前配置的忽略关键词
        skip_keywords_str = self.skip_keywords_var.get()
        skip_keywords = [kw.strip() for kw in skip_keywords_str.split(',') if kw.strip()]

        # 获取 Word 表格完整路径
        word_tables = get_word_tables(docx_path, skip_keywords=skip_keywords)
        self.word_paths = []
        self._max_word_level = 1
        seen = set()
        for _, path in word_tables:
            if path not in seen:
                self.word_paths.append(path)
                seen.add(path)
                # 计算最大层级数
                level = len(path.split('/'))
                if level > self._max_word_level:
                    self._max_word_level = level

        # 自动检测起始层级（公共前缀）
        self._detected_start_level, common_prefix = self._detect_start_level(self.word_paths)

        # 获取起始层级
        start_level = self._detected_start_level

        # 获取 Excel 路径
        excel_sheets = get_excel_sheets(excel_source, use_color_filter=use_color_filter)
        separator = self.separator_var.get()

        self._raw_excel_names = []
        self.excel_paths = []
        seen = set()

        for file_name, sheet_name in excel_sheets:
            # 构建原始名称（用于分隔符检测）
            if excel_mode == 'single':
                raw_name = sheet_name
            else:
                raw_name = file_name.replace('.xlsx', '').replace('.xls', '')

            if raw_name not in seen:
                self._raw_excel_names.append(raw_name)
                seen.add(raw_name)

            # 构建 Excel 完整路径
            file_base = file_name.replace('.xlsx', '').replace('.xls', '')
            file_base_cleaned = remove_prefix_by_separator(file_base, separator)
            sheet_cleaned = remove_prefix_by_separator(sheet_name, separator)

            if excel_mode == 'single':
                # 单文件模式：只用 Sheet 名
                excel_path = sheet_cleaned
            else:
                # 多文件模式：文件名/Sheet名（如果不同）
                if sheet_name == file_base or sheet_cleaned == file_base_cleaned:
                    excel_path = file_base_cleaned
                else:
                    # 使用 normalize_for_match 的逻辑，- 会被转为 /
                    excel_path = f"{file_base_cleaned}/{sheet_cleaned}"

            # 标准化路径（- 转为 /）
            excel_path = normalize_for_match(excel_path)

            if excel_path not in [p for p in self.excel_paths]:
                self.excel_paths.append(excel_path)

        # 更新标签文字
        self.word_list_label.configure(text=f"Word 表格路径（从{start_level}级标题开始）")
        self.excel_list_label.configure(text="Excel 路径（应用分隔符后）")

        # 更新 Word 列表显示（应用层级过滤）
        self.word_listbox.delete("1.0", "end")
        word_display = []
        for path in self.word_paths:
            processed = self._apply_start_level(path, start_level)
            word_display.append(processed)
        self.word_listbox.insert("1.0", "\n".join(word_display))

        # 更新 Excel 列表显示
        self.excel_listbox.delete("1.0", "end")
        self.excel_listbox.insert("1.0", "\n".join(self.excel_paths))

    def on_show(self):
        """显示时调用"""
        state = self.wizard.state

        # 检测路径和选项是否变化
        current_docx = state.get('docx_path', '')
        current_excel = state.get('excel_folder', '')
        current_mode = state.get('excel_mode', 'folder')
        current_filter = state.get('use_color_filter', True)

        paths_changed = (
            current_docx != self._last_docx_path or
            current_excel != self._last_excel_source or
            current_mode != self._last_excel_mode or
            current_filter != self._last_color_filter
        )

        # 如果路径变化或首次加载，重新加载数据
        if paths_changed or not self.word_paths:
            self._last_docx_path = current_docx
            self._last_excel_source = current_excel
            self._last_excel_mode = current_mode
            self._last_color_filter = current_filter

        # 每次显示都刷新预览（应用最新的规则配置）
        self._load_names()

    def validate(self):
        """验证"""
        return True

    def save_state(self):
        """保存状态"""
        separator = self.separator_var.get()
        skip_keywords_str = self.skip_keywords_var.get()
        word_start_level = self._detected_start_level

        # 解析忽略关键词（逗号分隔）
        skip_keywords = [kw.strip() for kw in skip_keywords_str.split(',') if kw.strip()]

        # 保存到向导状态
        self.wizard.state['prefix_separator'] = separator
        self.wizard.state['word_skip_keywords'] = skip_keywords
        self.wizard.state['word_start_level'] = word_start_level  # 数字形式，方便后续使用

        # 保存到配置文件
        config = load_config()
        config['prefix_separator'] = separator
        config['word_skip_keywords'] = skip_keywords_str
        save_config(config)
