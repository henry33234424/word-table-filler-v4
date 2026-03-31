"""
步骤 1：选择文件
- 选择 Word 模板文件
- 选择 Excel 数据文件夹
- 附加功能：根据 Word 导出 Excel 空表
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar
import os
import threading

from ..utils.config import get_recent_paths, save_recent_paths, load_config, save_config
from ..core.word_exporter import export_word_tables_to_excel


class Step1FilesFrame(ctk.CTkFrame):
    """步骤 1：文件选择"""

    def __init__(self, master, wizard, **kwargs):
        super().__init__(master, **kwargs)
        self.wizard = wizard
        self.configure(fg_color="transparent")

        # 先加载配置
        recent = get_recent_paths()
        config = load_config()

        # 创建 StringVar 并设置初始值
        self.word_var = StringVar(value=recent.get('docx_path', ''))
        self.excel_var = StringVar(value=recent.get('excel_folder', ''))
        self.output_var = StringVar(value=recent.get('output_path', ''))

        # Excel 模式（"folder" = 多文件模式，"single" = 单文件模式）
        self.excel_mode_var = StringVar(value=config.get('excel_mode', 'folder'))

        # 工作表筛选选项（"green" = 绿色标签优先，"all" = 所有Sheet）
        saved_filter = config.get('sheet_filter_mode', 'green')
        # 兼容旧配置
        if config.get('use_color_filter', True) is False:
            saved_filter = 'all'
        self.sheet_filter_var = StringVar(value=saved_filter)

        # 数据区域起始位置（默认第1行第1列）
        self.start_row_var = StringVar(value=str(config.get('data_start_row', 1)))
        self.start_col_var = StringVar(value=str(config.get('data_start_col', 1)))

        self._create_widgets()

    def _create_widgets(self):
        """创建组件"""
        # 主要内容区域（可滚动，用于低分辨率屏幕）
        content = ctk.CTkScrollableFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=50)

        # 说明文字
        intro_label = ctk.CTkLabel(
            content,
            text="请选择 Word 模板文件和 Excel 数据文件夹",
            font=ctk.CTkFont(size=16)
        )
        intro_label.pack(pady=(20, 30))

        # Word 文件选择
        word_frame = ctk.CTkFrame(content)
        word_frame.pack(fill="x", pady=10)

        word_label = ctk.CTkLabel(
            word_frame,
            text="Word 模板文件 (.docx)：",
            font=ctk.CTkFont(size=14),
            width=180,
            anchor="w"
        )
        word_label.pack(side="left", padx=10)

        self.word_entry = ctk.CTkEntry(word_frame, width=500, textvariable=self.word_var)
        self.word_entry.pack(side="left", padx=5, fill="x", expand=True)

        word_button = ctk.CTkButton(
            word_frame,
            text="浏览...",
            command=self._select_word_file,
            width=100
        )
        word_button.pack(side="left", padx=10)

        # Excel 模式选择
        excel_mode_frame = ctk.CTkFrame(content)
        excel_mode_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            excel_mode_frame,
            text="Excel 数据源模式：",
            font=ctk.CTkFont(size=14),
            width=180,
            anchor="w"
        ).pack(side="left", padx=10)

        ctk.CTkRadioButton(
            excel_mode_frame,
            text="多个Excel文件（选择文件夹）",
            variable=self.excel_mode_var,
            value="folder",
            font=ctk.CTkFont(size=13),
            command=self._on_excel_mode_change
        ).pack(side="left", padx=5)

        ctk.CTkRadioButton(
            excel_mode_frame,
            text="单个Excel工作簿（选择文件）",
            variable=self.excel_mode_var,
            value="single",
            font=ctk.CTkFont(size=13),
            command=self._on_excel_mode_change
        ).pack(side="left", padx=15)

        # Excel 文件/文件夹选择
        excel_frame = ctk.CTkFrame(content)
        excel_frame.pack(fill="x", pady=10)

        self.excel_label = ctk.CTkLabel(
            excel_frame,
            text="Excel 数据文件夹：",
            font=ctk.CTkFont(size=14),
            width=180,
            anchor="w"
        )
        self.excel_label.pack(side="left", padx=10)

        self.excel_entry = ctk.CTkEntry(excel_frame, width=500, textvariable=self.excel_var)
        self.excel_entry.pack(side="left", padx=5, fill="x", expand=True)

        self.excel_button = ctk.CTkButton(
            excel_frame,
            text="浏览...",
            command=self._select_excel_source,
            width=100
        )
        self.excel_button.pack(side="left", padx=10)

        # 输出文件选择
        output_frame = ctk.CTkFrame(content)
        output_frame.pack(fill="x", pady=10)

        output_label = ctk.CTkLabel(
            output_frame,
            text="输出文件路径：",
            font=ctk.CTkFont(size=14),
            width=180,
            anchor="w"
        )
        output_label.pack(side="left", padx=10)

        self.output_entry = ctk.CTkEntry(output_frame, width=500, textvariable=self.output_var)
        self.output_entry.pack(side="left", padx=5, fill="x", expand=True)

        output_button = ctk.CTkButton(
            output_frame,
            text="浏览...",
            command=self._select_output_file,
            width=100
        )
        output_button.pack(side="left", padx=10)

        # 选项区域
        options_frame = ctk.CTkFrame(content)
        options_frame.pack(fill="x", pady=(20, 10))

        options_label = ctk.CTkLabel(
            options_frame,
            text="如何选择待填充Excel工作表：",
            font=ctk.CTkFont(size=14),
            width=220,
            anchor="w"
        )
        options_label.pack(side="left", padx=10)

        ctk.CTkRadioButton(
            options_frame,
            text="带有绿色标签的Sheet（如无，则选择所有Sheet）",
            variable=self.sheet_filter_var,
            value="green",
            font=ctk.CTkFont(size=13)
        ).pack(side="left", padx=5)

        ctk.CTkRadioButton(
            options_frame,
            text="所有Sheet",
            variable=self.sheet_filter_var,
            value="all",
            font=ctk.CTkFont(size=13)
        ).pack(side="left", padx=15)

        # 数据区域设置
        data_range_frame = ctk.CTkFrame(content)
        data_range_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            data_range_frame,
            text="Excel数据区域起始位置：",
            font=ctk.CTkFont(size=14),
            width=220,
            anchor="w"
        ).pack(side="left", padx=10)

        ctk.CTkLabel(
            data_range_frame,
            text="第",
            font=ctk.CTkFont(size=13)
        ).pack(side="left")

        self.start_row_entry = ctk.CTkEntry(
            data_range_frame,
            width=50,
            textvariable=self.start_row_var
        )
        self.start_row_entry.pack(side="left", padx=2)

        ctk.CTkLabel(
            data_range_frame,
            text="行、第",
            font=ctk.CTkFont(size=13)
        ).pack(side="left")

        self.start_col_entry = ctk.CTkEntry(
            data_range_frame,
            width=50,
            textvariable=self.start_col_var
        )
        self.start_col_entry.pack(side="left", padx=2)

        ctk.CTkLabel(
            data_range_frame,
            text="列开始",
            font=ctk.CTkFont(size=13)
        ).pack(side="left")

        ctk.CTkLabel(
            data_range_frame,
            text="（默认从第1行第1列开始读取）",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray60")
        ).pack(side="left", padx=10)

        # 提示信息
        tip_frame = ctk.CTkFrame(content, fg_color=("gray90", "gray20"))
        tip_frame.pack(fill="x", pady=(10, 5), padx=10)

        tip_label = ctk.CTkLabel(
            tip_frame,
            text="提示：• Word 模板应为 .docx 文件  • 多文件模式：Excel 文件夹应包含各财务科目的数据文件  • 单文件模式：所有数据放在一个工作簿的不同 Sheet 中  • 绿色标签筛选：将要填充的 Sheet 标签设为绿色",
            font=ctk.CTkFont(size=11),
            justify="left",
            anchor="w",
            wraplength=800
        )
        tip_label.pack(padx=15, pady=8, fill="x")

        # ========== 附：根据 Word 模板导出 Excel 空表 ==========
        export_frame = ctk.CTkFrame(content)
        export_frame.pack(fill="x", pady=(10, 5))

        # 标题行
        export_title_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        export_title_row.pack(fill="x", padx=15, pady=(8, 3))

        ctk.CTkLabel(
            export_title_row,
            text="附：根据 Word 模板导出 Excel 空表",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        ctk.CTkLabel(
            export_title_row,
            text="（用于生成待填写的 Excel 数据文件）",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray60")
        ).pack(side="left", padx=10)

        # Word 模板选择行（独立的）
        export_word_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        export_word_row.pack(fill="x", padx=15, pady=2)

        ctk.CTkLabel(
            export_word_row,
            text="Word 模板：",
            font=ctk.CTkFont(size=12),
            width=100,
            anchor="w"
        ).pack(side="left")

        self.export_word_var = StringVar()
        self.export_word_entry = ctk.CTkEntry(
            export_word_row,
            width=400,
            textvariable=self.export_word_var
        )
        self.export_word_entry.pack(side="left", padx=5, fill="x", expand=True)

        ctk.CTkButton(
            export_word_row,
            text="浏览...",
            command=self._select_export_word_file,
            width=80
        ).pack(side="left", padx=5)

        # 导出路径选择行
        export_path_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        export_path_row.pack(fill="x", padx=15, pady=2)

        self.export_path_label = ctk.CTkLabel(
            export_path_row,
            text="导出文件夹：",
            font=ctk.CTkFont(size=12),
            width=100,
            anchor="w"
        )
        self.export_path_label.pack(side="left")

        self.export_path_var = StringVar()
        self.export_path_entry = ctk.CTkEntry(
            export_path_row,
            width=400,
            textvariable=self.export_path_var
        )
        self.export_path_entry.pack(side="left", padx=5, fill="x", expand=True)

        ctk.CTkButton(
            export_path_row,
            text="浏览...",
            command=self._select_export_path,
            width=80
        ).pack(side="left", padx=5)

        # 导出选项行
        export_options_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        export_options_row.pack(fill="x", padx=15, pady=(2, 8))

        self.export_mode_var = StringVar(value="folder")  # 默认为多个工作簿

        ctk.CTkRadioButton(
            export_options_row,
            text="导出为多个工作簿（每个表格一个文件）",
            variable=self.export_mode_var,
            value="folder",
            font=ctk.CTkFont(size=12),
            command=self._on_export_mode_change
        ).pack(side="left", padx=5)

        ctk.CTkRadioButton(
            export_options_row,
            text="导出为单个工作簿（每个表格一个 Sheet）",
            variable=self.export_mode_var,
            value="single",
            font=ctk.CTkFont(size=12),
            command=self._on_export_mode_change
        ).pack(side="left", padx=20)

        # 导出按钮
        ctk.CTkButton(
            export_options_row,
            text="导出 Excel",
            command=self._export_excel_from_word,
            width=100,
            fg_color="green"
        ).pack(side="right", padx=10)

    def _on_excel_mode_change(self):
        """Excel 模式切换时更新界面"""
        mode = self.excel_mode_var.get()
        if mode == "folder":
            self.excel_label.configure(text="Excel 数据文件夹：")
        else:
            self.excel_label.configure(text="Excel 工作簿文件：")
        # 清空已选择的路径
        self.excel_var.set("")

    def _select_word_file(self):
        """选择 Word 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 Word 模板文件",
            filetypes=[
                ("Word 文档", "*.docx"),
                ("所有文件", "*.*")
            ]
        )

        if file_path:
            self.word_var.set(file_path)
            self._auto_set_output_path(file_path)

    def _select_excel_source(self):
        """选择 Excel 数据源（文件夹或单个文件）"""
        mode = self.excel_mode_var.get()
        if mode == "folder":
            path = filedialog.askdirectory(
                title="选择 Excel 数据文件夹"
            )
        else:
            path = filedialog.askopenfilename(
                title="选择 Excel 工作簿",
                filetypes=[
                    ("Excel 文件", "*.xlsx"),
                    ("所有文件", "*.*")
                ]
            )

        if path:
            self.excel_var.set(path)

    def _select_output_file(self):
        """选择输出文件"""
        file_path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".docx",
            filetypes=[
                ("Word 文档", "*.docx"),
                ("所有文件", "*.*")
            ]
        )

        if file_path:
            self.output_var.set(file_path)

    def _auto_set_output_path(self, word_path):
        """自动设置输出路径"""
        if not self.output_var.get():
            base, ext = os.path.splitext(word_path)
            output_path = f"{base}_filled{ext}"
            self.output_var.set(output_path)

    def on_show(self):
        """显示时调用"""
        # 根据模式更新标签文字
        self._on_excel_mode_change_silent()

    def _on_excel_mode_change_silent(self):
        """更新界面但不清空路径（用于初始化）"""
        mode = self.excel_mode_var.get()
        if mode == "folder":
            self.excel_label.configure(text="Excel 数据文件夹：")
        else:
            self.excel_label.configure(text="Excel 工作簿文件：")

    def validate(self):
        """验证输入"""
        word_path = self.word_var.get().strip()
        excel_source = self.excel_var.get().strip()
        excel_mode = self.excel_mode_var.get()

        if not word_path:
            messagebox.showerror("错误", "请选择 Word 模板文件")
            return False

        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"Word 模板文件不存在：{word_path}")
            return False

        if not excel_source:
            if excel_mode == "folder":
                messagebox.showerror("错误", "请选择 Excel 数据文件夹")
            else:
                messagebox.showerror("错误", "请选择 Excel 工作簿文件")
            return False

        if excel_mode == "folder":
            if not os.path.isdir(excel_source):
                messagebox.showerror("错误", f"Excel 文件夹不存在：{excel_source}")
                return False
        else:
            if not os.path.isfile(excel_source):
                messagebox.showerror("错误", f"Excel 文件不存在：{excel_source}")
                return False

        return True

    def save_state(self):
        """保存状态"""
        docx_path = self.word_var.get().strip()
        excel_source = self.excel_var.get().strip()
        output_path = self.output_var.get().strip()
        excel_mode = self.excel_mode_var.get()
        sheet_filter_mode = self.sheet_filter_var.get()
        use_color_filter = (sheet_filter_mode == "green")

        # 获取数据区域起始位置
        try:
            start_row = int(self.start_row_var.get().strip())
            if start_row < 1:
                start_row = 1
        except ValueError:
            start_row = 1

        try:
            start_col = int(self.start_col_var.get().strip())
            if start_col < 1:
                start_col = 1
        except ValueError:
            start_col = 1

        # 更新向导状态
        self.wizard.state['docx_path'] = docx_path
        self.wizard.state['excel_folder'] = excel_source  # 保持兼容性，仍用 excel_folder
        self.wizard.state['excel_mode'] = excel_mode
        self.wizard.state['output_path'] = output_path
        self.wizard.state['use_color_filter'] = use_color_filter
        self.wizard.state['data_start_row'] = start_row
        self.wizard.state['data_start_col'] = start_col

        # 保存到配置
        save_recent_paths(
            docx_path=docx_path,
            excel_folder=excel_source,
            output_path=output_path
        )

        # 保存选项
        config = load_config()
        config['excel_mode'] = excel_mode
        config['sheet_filter_mode'] = sheet_filter_mode
        config['data_start_row'] = start_row
        config['data_start_col'] = start_col
        save_config(config)

    def _select_export_word_file(self):
        """选择导出功能的 Word 模板文件"""
        file_path = filedialog.askopenfilename(
            title="选择 Word 模板文件",
            filetypes=[
                ("Word 文档", "*.docx"),
                ("所有文件", "*.*")
            ]
        )

        if file_path:
            self.export_word_var.set(file_path)
            self._auto_set_export_path()

    def _on_export_mode_change(self):
        """导出模式切换时更新界面"""
        mode = self.export_mode_var.get()
        if mode == "folder":
            self.export_path_label.configure(text="导出文件夹：")
        else:
            self.export_path_label.configure(text="导出文件：")
        # 更新默认路径
        self._auto_set_export_path()

    def _auto_set_export_path(self):
        """自动设置导出路径"""
        word_path = self.export_word_var.get().strip()
        if not word_path:
            return

        word_dir = os.path.dirname(word_path)
        word_name = os.path.splitext(os.path.basename(word_path))[0]
        mode = self.export_mode_var.get()

        if mode == "folder":
            # 多文件模式：默认在 Word 模板同级目录下创建文件夹
            export_path = os.path.join(word_dir, f"{word_name}_导出Excel")
        else:
            # 单文件模式：默认在 Word 模板同级目录下创建文件
            export_path = os.path.join(word_dir, f"{word_name}_导出Excel.xlsx")

        self.export_path_var.set(export_path)

    def _select_export_path(self):
        """选择导出路径"""
        mode = self.export_mode_var.get()
        word_path = self.export_word_var.get().strip()
        current_path = self.export_path_var.get().strip()

        # 确定初始目录
        initial_dir = ""
        if current_path:
            if mode == "single":
                initial_dir = os.path.dirname(current_path)
            else:
                initial_dir = os.path.dirname(current_path) if not os.path.isdir(current_path) else current_path
        elif word_path:
            initial_dir = os.path.dirname(word_path)

        if mode == "folder":
            # 选择文件夹
            path = filedialog.askdirectory(
                title="选择导出文件夹",
                initialdir=initial_dir
            )
        else:
            # 选择文件
            default_name = ""
            if word_path:
                default_name = os.path.splitext(os.path.basename(word_path))[0] + "_导出Excel.xlsx"
            path = filedialog.asksaveasfilename(
                title="保存 Excel 文件",
                defaultextension=".xlsx",
                initialfile=default_name,
                initialdir=initial_dir,
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )

        if path:
            self.export_path_var.set(path)

    def _export_excel_from_word(self):
        """根据 Word 模板导出 Excel 空表"""
        word_path = self.export_word_var.get().strip()

        if not word_path:
            messagebox.showerror("错误", "请先选择 Word 模板文件")
            return

        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"Word 模板文件不存在：{word_path}")
            return

        export_mode = self.export_mode_var.get()
        output_path = self.export_path_var.get().strip()

        if not output_path:
            messagebox.showerror("错误", "请选择导出路径")
            return

        # 创建进度对话框
        self._show_export_progress_dialog(word_path, output_path, export_mode)

    def _show_export_progress_dialog(self, word_path, output_path, export_mode):
        """显示导出进度对话框"""
        # 创建模态对话框
        dialog = ctk.CTkToplevel(self)
        dialog.title("导出 Excel")
        dialog.geometry("450x180")
        dialog.resizable(False, False)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = self.winfo_toplevel().winfo_x() + (self.winfo_toplevel().winfo_width() - 450) // 2
        y = self.winfo_toplevel().winfo_y() + (self.winfo_toplevel().winfo_height() - 180) // 2
        dialog.geometry(f"+{x}+{y}")

        # 状态标签
        status_label = ctk.CTkLabel(
            dialog,
            text="正在导出，请稍候...",
            font=ctk.CTkFont(size=14)
        )
        status_label.pack(pady=(25, 10))

        # 进度条
        progress_bar = ctk.CTkProgressBar(dialog, width=350)
        progress_bar.pack(pady=10)
        progress_bar.set(0)
        progress_bar.configure(mode="indeterminate")
        progress_bar.start()

        # 结果标签（初始隐藏）
        result_label = ctk.CTkLabel(
            dialog,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60")
        )
        result_label.pack(pady=5)

        # 关闭按钮（初始隐藏）
        close_btn = ctk.CTkButton(
            dialog,
            text="关闭",
            command=dialog.destroy,
            width=100
        )

        # 在后台线程执行导出
        def do_export():
            try:
                result = export_word_tables_to_excel(
                    word_path,
                    output_path,
                    single_file=(export_mode == "single")
                )
                # 在主线程更新界面
                dialog.after(0, lambda: on_export_complete(result))
            except Exception as e:
                dialog.after(0, lambda: on_export_error(str(e)))

        def on_export_complete(result):
            progress_bar.stop()
            progress_bar.configure(mode="determinate")
            progress_bar.set(1)

            if result['success']:
                status_label.configure(text="导出完成！", text_color=("green", "lightgreen"))
                if export_mode == "single":
                    result_label.configure(text=f"共导出 {result['table_count']} 个表格")
                else:
                    result_label.configure(text=f"共导出 {result['table_count']} 个表格，生成 {result['file_count']} 个文件")
            else:
                status_label.configure(text="导出失败", text_color=("red", "lightcoral"))
                result_label.configure(text=result.get('error', '未知错误'))

            close_btn.pack(pady=15)

        def on_export_error(error_msg):
            progress_bar.stop()
            progress_bar.configure(mode="determinate")
            progress_bar.set(0)
            status_label.configure(text="导出失败", text_color=("red", "lightcoral"))
            result_label.configure(text=error_msg)
            close_btn.pack(pady=15)

        # 启动导出线程
        threading.Thread(target=do_export, daemon=True).start()

    def _show_export_result(self, result, output_path, export_mode):
        """显示导出结果（保留兼容）"""
        if result['success']:
            if export_mode == "single":
                msg = f"导出成功！\n\n共导出 {result['table_count']} 个表格到：\n{output_path}"
            else:
                msg = f"导出成功！\n\n共导出 {result['table_count']} 个表格到文件夹：\n{output_path}\n\n生成了 {result['file_count']} 个 Excel 文件"
            messagebox.showinfo("导出完成", msg)
        else:
            messagebox.showerror("导出失败", f"导出失败：{result.get('error', '未知错误')}")
