"""
向导框架
- 管理多步骤向导流程
- 提供步骤导航
"""

import customtkinter as ctk


class StepIndicator(ctk.CTkFrame):
    """步骤指示器"""

    def __init__(self, master, steps, **kwargs):
        super().__init__(master, **kwargs)

        self.steps = steps
        self.current_step = 0
        self.step_labels = []
        self.step_circles = []

        self._create_widgets()

    def _create_widgets(self):
        """创建组件"""
        self.configure(fg_color="transparent")

        for i, step_name in enumerate(self.steps):
            # 步骤容器
            step_frame = ctk.CTkFrame(self, fg_color="transparent")
            step_frame.pack(side="left", padx=20)

            # 圆圈
            circle = ctk.CTkLabel(
                step_frame,
                text=str(i + 1),
                width=36,
                height=36,
                corner_radius=18,
                fg_color=("gray70", "gray30"),
                text_color=("gray30", "gray70"),
                font=ctk.CTkFont(size=14, weight="bold")
            )
            circle.pack()
            self.step_circles.append(circle)

            # 步骤名称
            label = ctk.CTkLabel(
                step_frame,
                text=step_name,
                font=ctk.CTkFont(size=12),
                text_color=("gray50", "gray50")
            )
            label.pack(pady=(5, 0))
            self.step_labels.append(label)

            # 连接线（除了最后一个）
            if i < len(self.steps) - 1:
                line = ctk.CTkFrame(self, height=2, width=60, fg_color=("gray70", "gray30"))
                line.pack(side="left", pady=(0, 20))

        self.update_step(0)

    def update_step(self, step_index):
        """更新当前步骤"""
        self.current_step = step_index

        for i, (circle, label) in enumerate(zip(self.step_circles, self.step_labels)):
            if i < step_index:
                # 已完成
                circle.configure(
                    fg_color=("green", "#2E7D32"),
                    text_color="white",
                    text="✓"
                )
                label.configure(text_color=("gray30", "gray70"))
            elif i == step_index:
                # 当前
                circle.configure(
                    fg_color=("#1976D2", "#1976D2"),
                    text_color="white",
                    text=str(i + 1)
                )
                label.configure(text_color=("gray10", "gray90"))
            else:
                # 未完成
                circle.configure(
                    fg_color=("gray70", "gray30"),
                    text_color=("gray30", "gray70"),
                    text=str(i + 1)
                )
                label.configure(text_color=("gray50", "gray50"))


class WizardFrame(ctk.CTkFrame):
    """向导主框架"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.configure(fg_color="transparent")

        # 应用状态
        self.state = {
            'docx_path': '',
            'excel_folder': '',
            'output_path': '',
            'mapping_data': [],
            'word_tables': [],
            'excel_sheets': [],
        }

        self.current_step = 0
        self.step_frames = []

        self._create_widgets()
        self._create_steps()
        self.show_step(0)

    def _create_widgets(self):
        """创建组件"""
        # 标题
        title_label = ctk.CTkLabel(
            self,
            text="Excel-to-Word Template Filler -- Word表格模板填报工具",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 10))

        # 步骤指示器
        self.step_indicator = StepIndicator(
            self,
            steps=["选择文件", "匹配规则", "编辑匹配", "执行填充"]
        )
        self.step_indicator.pack(pady=20)

        # 导航按钮区域（先 pack，确保始终可见）
        self.nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.nav_frame.pack(side="bottom", fill="x", pady=(10, 0))

        self.prev_button = ctk.CTkButton(
            self.nav_frame,
            text="上一步",
            command=self.prev_step,
            width=120
        )
        self.prev_button.pack(side="left")

        self.next_button = ctk.CTkButton(
            self.nav_frame,
            text="下一步",
            command=self.next_step,
            width=120
        )
        self.next_button.pack(side="right")

        # 内容区域（最后 pack，使用剩余空间）
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, pady=10)

    def _create_steps(self):
        """创建步骤页面"""
        from .step1_files import Step1FilesFrame
        from .step2_rules import Step2RulesFrame
        from .step3_mapping import Step3MappingFrame
        from .step4_execute import Step4ExecuteFrame

        step1 = Step1FilesFrame(self.content_frame, self)
        step2 = Step2RulesFrame(self.content_frame, self)
        step3 = Step3MappingFrame(self.content_frame, self)
        step4 = Step4ExecuteFrame(self.content_frame, self)

        self.step_frames = [step1, step2, step3, step4]

    def show_step(self, step_index):
        """显示指定步骤"""
        # 隐藏所有步骤
        for frame in self.step_frames:
            frame.pack_forget()

        # 显示当前步骤
        self.current_step = step_index
        self.step_frames[step_index].pack(fill="both", expand=True)
        self.step_frames[step_index].on_show()

        # 更新指示器
        self.step_indicator.update_step(step_index)

        # 更新导航按钮
        self.prev_button.configure(state="normal" if step_index > 0 else "disabled")

        if step_index == len(self.step_frames) - 1:
            self.next_button.configure(text="完成", state="disabled")
        else:
            self.next_button.configure(text="下一步", state="normal")

    def prev_step(self):
        """上一步"""
        if self.current_step > 0:
            self.show_step(self.current_step - 1)

    def next_step(self):
        """下一步"""
        # 验证当前步骤
        current_frame = self.step_frames[self.current_step]
        if hasattr(current_frame, 'validate') and not current_frame.validate():
            return

        # 保存当前步骤数据
        if hasattr(current_frame, 'save_state'):
            current_frame.save_state()

        if self.current_step < len(self.step_frames) - 1:
            self.show_step(self.current_step + 1)
        else:
            # 最后一步，弹出询问框
            self._on_finish()

    def _on_finish(self):
        """完成按钮点击后的处理"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("完成")
        dialog.geometry("360x80")
        dialog.resizable(False, False)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = self.winfo_toplevel().winfo_x() + (self.winfo_toplevel().winfo_width() - 360) // 2
        y = self.winfo_toplevel().winfo_y() + (self.winfo_toplevel().winfo_height() - 80) // 2
        dialog.geometry(f"+{x}+{y}")

        # 按钮区域
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(expand=True, pady=20)

        def on_new_task():
            dialog.destroy()
            self._reset_for_new_task()

        def on_exit():
            dialog.destroy()
            self.winfo_toplevel().destroy()

        def on_cancel():
            dialog.destroy()

        ctk.CTkButton(btn_frame, text="开始新任务", command=on_new_task, width=100).pack(side="left", padx=8)
        ctk.CTkButton(btn_frame, text="关闭程序", command=on_exit, width=100).pack(side="left", padx=8)
        ctk.CTkButton(btn_frame, text="取消", command=on_cancel, width=80, fg_color="gray").pack(side="left", padx=8)

    def _reset_for_new_task(self):
        """重置状态，开始新任务"""
        # 清空映射相关的状态（保留路径配置）
        self.state['mapping_data'] = []
        self.state['word_tables'] = []
        self.state['excel_sheets'] = []
        self.state['_mapping_needs_refresh'] = True

        # 清除追踪变量
        keys_to_remove = [k for k in self.state.keys() if k.startswith('_last_')]
        for key in keys_to_remove:
            del self.state[key]

        # 重置 step4 (执行填充) 的完成状态
        step4 = self.step_frames[3]
        step4.is_complete = False
        step4._last_mapping_hash = None

        # 回到 step1
        self.show_step(0)

    def set_next_enabled(self, enabled):
        """设置下一步按钮状态"""
        self.next_button.configure(state="normal" if enabled else "disabled")
