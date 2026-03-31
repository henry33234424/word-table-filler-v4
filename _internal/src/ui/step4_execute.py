"""
步骤 3：执行填充
- 显示进度
- 显示结果
- 提供打开文件的选项
"""

import customtkinter as ctk
from tkinter import messagebox
import threading
import subprocess
import platform
import os
import json

from ..core.filler import DocumentFiller
from ..core.mapping import get_valid_mappings


class Step4ExecuteFrame(ctk.CTkFrame):
    """步骤 3：执行填充"""

    def __init__(self, master, wizard, **kwargs):
        super().__init__(master, **kwargs)
        self.wizard = wizard
        self.configure(fg_color="transparent")
        self.is_running = False
        self.is_complete = False
        self._last_mapping_hash = None  # 用于检测映射是否已更改

        self._create_widgets()

    def _create_widgets(self):
        """创建组件"""
        # 主要内容区域（可滚动，用于低分辨率屏幕）
        self.content = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.content.pack(fill="both", expand=True, padx=50)

        # 标题
        self.title_label = ctk.CTkLabel(
            self.content,
            text="准备开始填充",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.title_label.pack(pady=(20, 10))

        # 摘要信息
        self.summary_frame = ctk.CTkFrame(self.content)
        self.summary_frame.pack(fill="x", pady=10)

        self.summary_label = ctk.CTkLabel(
            self.summary_frame,
            text="",
            font=ctk.CTkFont(size=12),
            justify="left",
            anchor="w"
        )
        self.summary_label.pack(padx=20, pady=15, fill="x")

        # 进度条
        self.progress_frame = ctk.CTkFrame(self.content, fg_color="transparent")
        self.progress_frame.pack(fill="x", pady=20)

        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="",
            font=ctk.CTkFont(size=12)
        )
        self.progress_label.pack(anchor="w")

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, width=600)
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.set(0)

        # 按钮区域（放在日志上面，确保可见）
        self.result_frame = ctk.CTkFrame(self.content, fg_color="transparent")
        self.result_frame.pack(fill="x", pady=10)

        # 开始按钮
        self.start_button = ctk.CTkButton(
            self.result_frame,
            text="开始填充",
            command=self._start_filling,
            width=150,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.start_button.pack(pady=10)

        # 完成后的操作按钮（初始隐藏）
        self.action_frame = ctk.CTkFrame(self.result_frame, fg_color="transparent")

        self.open_file_btn = ctk.CTkButton(
            self.action_frame,
            text="打开文件",
            command=self._open_output_file,
            width=120
        )
        self.open_file_btn.pack(side="left", padx=10)

        self.open_folder_btn = ctk.CTkButton(
            self.action_frame,
            text="打开所在文件夹",
            command=self._open_output_folder,
            width=140
        )
        self.open_folder_btn.pack(side="left", padx=10)

        # 日志区域（放在按钮下面）
        self.log_frame = ctk.CTkFrame(self.content)
        self.log_frame.pack(fill="both", expand=True, pady=(0, 10))

        ctk.CTkLabel(
            self.log_frame,
            text="执行日志：",
            font=ctk.CTkFont(size=12)
        ).pack(anchor="w", padx=10, pady=(10, 0))

        self.log_text = ctk.CTkTextbox(self.log_frame, height=200)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

    def on_show(self):
        """显示时调用"""
        self._update_summary()
        self.wizard.set_next_enabled(False)

        # 确保上一步按钮可用（除非正在运行）
        self.wizard.prev_button.configure(state="disabled" if self.is_running else "normal")

        # 检查映射是否已更改
        current_mapping_hash = self._get_mapping_hash()
        if self._last_mapping_hash is not None and current_mapping_hash != self._last_mapping_hash:
            # 映射已更改，重置完成状态
            self.is_complete = False
            self.title_label.configure(text="准备开始填充")

        if not self.is_complete:
            self.progress_bar.set(0)
            self.log_text.delete("1.0", "end")
            self.action_frame.pack_forget()
            self.start_button.configure(state="normal", text="开始填充")
            self.start_button.pack(pady=10)

    def _get_mapping_hash(self):
        """获取当前映射数据的哈希值，用于检测变化"""
        mapping_data = self.wizard.state.get('mapping_data', [])
        try:
            return hash(json.dumps(mapping_data, sort_keys=True))
        except Exception:
            return None

    def _update_summary(self):
        """更新摘要信息"""
        state = self.wizard.state
        mapping_data = state.get('mapping_data', [])
        valid_mappings = get_valid_mappings(mapping_data)

        docx_path = state.get('docx_path', '')
        excel_folder = state.get('excel_folder', '')
        output_path = state.get('output_path', '')
        excel_mode = state.get('excel_mode', 'folder')
        data_start_row = state.get('data_start_row', 1)
        data_start_col = state.get('data_start_col', 1)

        # 根据模式显示不同的标签
        excel_label = "Excel 数据源" if excel_mode == 'single' else "Excel 文件夹"

        summary = (
            f"Word 模板：{os.path.basename(docx_path)}\n"
            f"{excel_label}：{os.path.basename(excel_folder)}\n"
            f"输出文件：{os.path.basename(output_path)}\n"
            f"数据区域：从第 {data_start_row} 行、第 {data_start_col} 列开始\n"
            f"\n"
            f"将填充 {len(valid_mappings)} 个表格"
        )

        self.summary_label.configure(text=summary)

    def _log(self, message):
        """添加日志"""
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def _start_filling(self):
        """开始填充"""
        if self.is_running:
            return

        self.is_running = True
        self.is_complete = False
        self.start_button.configure(state="disabled", text="正在填充...")
        self.log_text.delete("1.0", "end")
        self.progress_bar.set(0)

        # 禁用上一步按钮，防止填充过程中切换
        self.wizard.prev_button.configure(state="disabled")

        # 在后台线程执行
        threading.Thread(target=self._do_fill, daemon=True).start()

    def _do_fill(self):
        """执行填充（后台线程）"""
        try:
            state = self.wizard.state
            docx_path = state['docx_path']
            excel_folder = state['excel_folder']
            output_path = state['output_path']
            mapping_data = state['mapping_data']

            # 获取有效映射
            mappings = get_valid_mappings(mapping_data)

            # 获取数据区域设置
            data_start_row = state.get('data_start_row', 1)
            data_start_col = state.get('data_start_col', 1)
            excel_mode = state.get('excel_mode', 'folder')
            single_file_mode = (excel_mode == 'single')
            skip_keywords = state.get('word_skip_keywords', None)

            self.after(0, lambda: self._log(f"开始填充，共 {len(mappings)} 个表格..."))
            if data_start_row > 1 or data_start_col > 1:
                self.after(0, lambda: self._log(f"数据区域从第 {data_start_row} 行、第 {data_start_col} 列开始"))

            # 创建填充器
            filler = DocumentFiller(
                docx_path, excel_folder, output_path,
                data_start_row=data_start_row,
                data_start_col=data_start_col,
                single_file_mode=single_file_mode,
                skip_keywords=skip_keywords
            )
            filler.prepare()

            self.after(0, lambda: self._log("模板已加载"))

            # 进度回调
            def progress_callback(current, total, message):
                progress = current / total
                self.after(0, lambda p=progress, m=message: self._update_progress(p, m))

            # 执行填充
            result = filler.fill(mappings, progress_callback)

            # 保存
            output_file = filler.save()
            filler.cleanup()

            self.after(0, lambda: self._on_fill_complete(result, output_file))

        except Exception as e:
            self.after(0, lambda: self._on_fill_error(str(e)))

    def _update_progress(self, progress, message):
        """更新进度"""
        self.progress_bar.set(progress)
        self.progress_label.configure(text=message)
        self._log(message)

    def _on_fill_complete(self, result, output_file):
        """填充完成"""
        self.is_running = False
        self.is_complete = True

        # 记录完成时的映射哈希值
        self._last_mapping_hash = self._get_mapping_hash()

        self.progress_bar.set(1)
        self.title_label.configure(text="填充完成！")
        self.start_button.pack_forget()
        self.action_frame.pack(pady=10)

        # 重新启用上一步按钮
        self.wizard.prev_button.configure(state="normal")

        # 启用"完成"按钮
        self.wizard.set_next_enabled(True)

        # 显示结果
        self._log("")
        self._log("=" * 50)
        self._log("填充完成！")
        self._log(f"  成功: {result['success']}")
        self._log(f"  失败: {result['failed']}")
        self._log(f"  共填充: {result['total_filled']} 个单元格")
        if result['total_added'] > 0:
            self._log(f"  共增加: {result['total_added']} 行")
        if result['total_removed'] > 0:
            self._log(f"  共删除: {result['total_removed']} 行")

        # 显示错误
        if result['errors']:
            self._log("")
            self._log("错误列表：")
            for word_path, error_msg in result['errors']:
                self._log(f"  ✗ {word_path}: {error_msg}")

        # 显示空标题警告
        if result['empty_headers']:
            self._log("")
            self._log("⚠ 以下表格的标题行存在空单元格，请手工补充：")
            for word_path, empty_cells in result['empty_headers']:
                positions = [f"第{r}行第{c}列" for r, c in empty_cells]
                self._log(f"  • {word_path}")
                self._log(f"    空单元格: {', '.join(positions)}")

        self._log("")
        self._log(f"输出文件: {output_file}")

        # 保存输出路径
        self.output_file = output_file

    def _on_fill_error(self, error_msg):
        """填充出错"""
        self.is_running = False
        self.title_label.configure(text="填充失败")
        self.start_button.configure(state="normal", text="重试")

        # 重新启用上一步按钮
        self.wizard.prev_button.configure(state="normal")

        self._log("")
        self._log(f"错误: {error_msg}")

        messagebox.showerror("错误", f"填充失败：{error_msg}")

    def _open_output_file(self):
        """打开输出文件"""
        if hasattr(self, 'output_file') and os.path.exists(self.output_file):
            if platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', self.output_file])
            elif platform.system() == 'Windows':
                os.startfile(self.output_file)
            else:  # Linux
                subprocess.run(['xdg-open', self.output_file])

    def _open_output_folder(self):
        """打开输出文件夹"""
        if hasattr(self, 'output_file'):
            folder = os.path.dirname(self.output_file)
            if os.path.exists(folder):
                if platform.system() == 'Darwin':
                    subprocess.run(['open', folder])
                elif platform.system() == 'Windows':
                    os.startfile(folder)
                else:
                    subprocess.run(['xdg-open', folder])

    def validate(self):
        """验证"""
        return True

    def save_state(self):
        """保存状态"""
        pass
