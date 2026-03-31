"""
主应用窗口
"""

import customtkinter as ctk
from .ui.wizard import WizardFrame


class App(ctk.CTk):
    """主应用程序"""

    def __init__(self):
        super().__init__()

        # 窗口设置
        self.title("Excel-to-Word Template Filler -- Word表格模板填报工具")
        self.geometry("1200x900")
        self.minsize(900, 700)

        # 设置主题
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # 居中显示
        self.center_window()

        # 创建向导框架
        self.wizard = WizardFrame(self)
        self.wizard.pack(fill="both", expand=True, padx=20, pady=20)

    def center_window(self):
        """窗口居中"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')

    def run(self):
        """运行应用"""
        self.mainloop()
