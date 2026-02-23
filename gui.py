# "gui.py"
# 源代码文件地址
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
from core.processor import HomeworkProcessor
from core.config_manager import ConfigManager


class HomeworkCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("作业检查与重命名系统")
        self.root.geometry("800x600")

        self.processor = HomeworkProcessor()
        self.config_manager = ConfigManager()

        self.setup_ui()
        self.load_config()

    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 花名册选择
        ttk.Label(main_frame, text="花名册文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.roster_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.roster_var).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_roster).grid(row=0, column=2, padx=5)

        # 作业文件夹选择
        ttk.Label(main_frame, text="作业文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.homework_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.homework_var).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_homework).grid(row=1, column=2, padx=5)

        # 输出目录选择
        ttk.Label(main_frame, text="输出目录:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.output_var).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.browse_output).grid(row=2, column=2, padx=5)

        # 重命名格式配置
        ttk.Label(main_frame, text="重命名格式:").grid(row=3, column=0, sticky=tk.W, pady=5)
        format_frame = ttk.Frame(main_frame)
        format_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        format_frame.columnconfigure(0, weight=1)

        self.format_var = tk.StringVar()
        self.format_combo = ttk.Combobox(format_frame, textvariable=self.format_var, state="readonly")
        self.format_combo.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        self.format_combo['values'] = self.config_manager.get_format_names()
        self.format_combo.bind('<<ComboboxSelected>>', self.on_format_selected)

        ttk.Button(format_frame, text="管理格式", command=self.manage_formats).grid(row=0, column=1)

        # 操作按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)

        ttk.Button(button_frame, text="开始检查", command=self.start_check).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="仅重命名文件", command=self.rename_only).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        # 添加【快速配置新花名册】按钮
        ttk.Button(button_frame, text="快速配置新花名册", command=self.quick_setup).pack(side=tk.LEFT, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # 日志文本框
        ttk.Label(main_frame, text="处理日志:").grid(row=6, column=0, sticky=tk.W, pady=(10, 0))
        self.log_text = tk.Text(main_frame, height=15, width=80)
        self.log_text.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # 配置滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=7, column=3, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # === 新增：批量检查框架 ===
        ttk.Separator(main_frame, orient='horizontal').grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15)

        batch_frame = ttk.LabelFrame(main_frame, text="🚀 批量检查汇总（多实验文件夹）", padding="10")
        batch_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # 母文件夹选择
        ttk.Label(batch_frame, text="母文件夹（包含多个实验子文件夹）:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.batch_parent_var = tk.StringVar()
        ttk.Entry(batch_frame, textvariable=self.batch_parent_var, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E),
                                                                                  padx=(5, 0), pady=5)
        ttk.Button(batch_frame, text="浏览", command=self.browse_batch_parent, width=8).grid(row=0, column=2,
                                                                                             padx=(5, 5))

        # 在“浏览”按钮下方添加一个“选择子文件夹”按钮
        ttk.Label(batch_frame, text="选择要扫描的子文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Button(batch_frame, text="选择子文件夹...",
                   command=self.select_subfolders, width=15).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=5)

        self.selected_folders_var = tk.StringVar(value="未选择")
        ttk.Label(batch_frame, textvariable=self.selected_folders_var,
                  foreground="blue", wraplength=400).grid(row=1, column=2, sticky=tk.W, padx=(10, 0), pady=5)

        # 批量检查按钮
        ttk.Button(batch_frame, text="开始批量检查汇总", command=self.batch_check,
                   style="Accent.TButton").grid(row=2, column=1, pady=15)

        # 批量检查说明
        help_label = ttk.Label(batch_frame,
                               text="💡 功能：自动扫描‘母文件夹’下所有子文件夹，生成一份汇总Excel，显示每个学生在每个实验的提交情况。",
                               foreground="gray")
        help_label.grid(row=3, column=0, columnspan=3, sticky=tk.W)

        # 配置网格权重
        batch_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(9, weight=0)

        # 配置主框架行权重
        main_frame.rowconfigure(7, weight=1)

        # === 新增：底部 GitHub 链接 ===
        github_frame = ttk.Frame(main_frame)
        github_frame.grid(row=10, column=0, columnspan=3, pady=10, sticky=tk.EW)
        github_frame.columnconfigure(0, weight=1)  # 让左侧空白区域可扩展，使按钮靠右

        github_label = ttk.Label(github_frame, text="如果这个项目对你有帮助，欢迎前往 GitHub 为我点亮 Star ⭐️",
                                 font=('TkDefaultFont', 9), foreground='gray')
        github_label.pack(side=tk.LEFT, padx=10)

        def open_github():
            import webbrowser
            webbrowser.open("https://github.com/Mlukth/-Batch-Check-Unsubmitted-Homework-Standardize-File-Naming-")

        github_btn = ttk.Button(github_frame, text="跳转", command=open_github, width=8)
        github_btn.pack(side=tk.RIGHT, padx=10)

    def browse_roster(self):
        filename = filedialog.askopenfilename(
            title="选择花名册文件",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.roster_var.set(filename)

    def browse_homework(self):
        directory = filedialog.askdirectory(title="选择作业文件夹")
        if directory:
            self.homework_var.set(directory)

    def browse_output(self):
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_var.set(directory)

    def on_format_selected(self, event):
        format_name = self.format_var.get()
        format_config = self.config_manager.get_format_config(format_name)
        if format_config:
            self.log(f"已选择格式: {format_name}")

    def manage_formats(self):
        """打开格式管理窗口（已改造为纯界面点选模式）"""
        # 在打开管理窗口前，必须确保已有花名册（这样才能知道有哪些变量）
        if not self.roster_var.get():
            messagebox.showwarning("警告", "请先通过【快速配置新花名册】或手动选择导入花名册文件。")
            return

        format_window = tk.Toplevel(self.root)
        format_window.title("管理重命名格式 - 点选模式")
        format_window.geometry("650x500")

        # 将当前花名册的列信息传递给管理窗口
        FormatManagerWindow(format_window, self.config_manager, self.refresh_formats)

    def refresh_formats(self):
        self.format_combo['values'] = self.config_manager.get_format_names()

    def start_check(self):
        if not self.validate_inputs():
            return

        self.progress.start()
        self.log("开始检查作业...")

        try:
            # 获取选中的格式配置
            format_name = self.format_var.get()
            format_config = self.config_manager.get_format_config(format_name)

            if not format_config:
                messagebox.showerror("错误", "请选择有效的重命名格式")
                return

            self.processor.process_homework(
                roster_path=self.roster_var.get(),
                homework_dir=self.homework_var.get(),
                output_dir=self.output_var.get(),
                rename_format=format_config,
                log_callback=self.log
            )

            self.log("处理完成！")
            messagebox.showinfo("完成", "作业检查处理完成！")

        except Exception as e:
            self.log(f"处理失败: {str(e)}")
            messagebox.showerror("错误", f"处理失败: {str(e)}")
        finally:
            self.progress.stop()

    def rename_only(self):
        if not self.validate_inputs():
            return

        self.progress.start()
        self.log("开始重命名文件...")

        try:
            format_name = self.format_var.get()
            format_config = self.config_manager.get_format_config(format_name)

            if not format_config:
                messagebox.showerror("错误", "请选择有效的重命名格式")
                return

            count = self.processor.rename_files_only(
                roster_path=self.roster_var.get(),
                homework_dir=self.homework_var.get(),
                rename_format=format_config,
                log_callback=self.log
            )

            self.log(f"重命名完成，共处理 {count} 个文件")
            messagebox.showinfo("完成", f"文件重命名完成！共处理 {count} 个文件")

        except Exception as e:
            self.log(f"重命名失败: {str(e)}")
            messagebox.showerror("错误", f"重命名失败: {str(e)}")
        finally:
            self.progress.stop()

    def browse_batch_parent(self):
        """浏览选择母文件夹"""
        directory = filedialog.askdirectory(title="选择母文件夹（它包含实验一、实验二等子文件夹）")
        if directory:
            self.batch_parent_var.set(directory)

    def select_subfolders(self):
        """打开子文件夹选择和管理窗口"""
        parent_dir = self.batch_parent_var.get()
        if not parent_dir or not os.path.exists(parent_dir):
            messagebox.showwarning("提示", "请先选择有效的母文件夹。")
            return

        # 获取母文件夹下的所有子文件夹
        all_subfolders = []
        try:
            for item in os.listdir(parent_dir):
                item_path = os.path.join(parent_dir, item)
                if os.path.isdir(item_path) and not item.startswith('.'):
                    all_subfolders.append(item)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件夹失败: {str(e)}")
            return

        if not all_subfolders:
            messagebox.showinfo("提示", "该文件夹下没有找到子文件夹。")
            return

        # 打开文件夹选择窗口
        select_window = tk.Toplevel(self.root)
        select_window.title("选择并排序子文件夹")
        select_window.geometry("500x500")

        FolderSelectorWindow(select_window, parent_dir, all_subfolders,
                             self.config_manager, self.update_selected_folders)

    def update_selected_folders(self, selected_folders):
        """更新显示已选择的文件夹"""
        if selected_folders:
            text = f"已选择 {len(selected_folders)} 个: " + ", ".join(selected_folders[:3])
            if len(selected_folders) > 3:
                text += f" 等{len(selected_folders)}个"
            self.selected_folders_var.set(text)
        else:
            self.selected_folders_var.set("未选择")

    def batch_check(self):
        """执行批量检查汇总"""
        if not self.roster_var.get():
            messagebox.showerror("错误", "请先选择或导入花名册文件。")
            return
        if not self.batch_parent_var.get():
            messagebox.showerror("错误", "请选择包含多个实验子文件夹的‘母文件夹’。")
            return
        if not os.path.exists(self.batch_parent_var.get()):
            messagebox.showerror("错误", "选择的母文件夹不存在，请重新选择。")
            return

        self.progress.start()
        self.log("\n" + "=" * 60)
        self.log("开始批量检查汇总...")
        self.log(f"母文件夹: {self.batch_parent_var.get()}")

        try:
            # 获取当前选中的格式（用于重命名，可选）
            format_name = self.format_var.get()
            format_config = self.config_manager.get_format_config(format_name) if format_name else None

            # 获取选择的子文件夹配置
            folder_config = self.config_manager.load_folder_config(self.batch_parent_var.get())
            selected_folders = None
            if folder_config and 'selected_folders' in folder_config:
                selected_folders = folder_config['selected_folders']
                self.log(f"使用预设文件夹选择: {len(selected_folders)} 个文件夹")

            # 调用处理器的批量检查方法
            output_path = self.processor.batch_check_submissions(
                roster_path=self.roster_var.get(),
                parent_dir=self.batch_parent_var.get(),
                rename_format=format_config,  # 可以为None，表示不重命名
                selected_folders=selected_folders,  # 传递选择的文件夹
                log_callback=self.log
            )

            self.log(f"✅ 批量汇总完成！报告已生成: {output_path}")
            messagebox.showinfo("批量检查完成",
                                f"汇总报告已生成！\n\n"
                                f"文件位置: {output_path}\n\n"
                                f"报告包含所有子文件夹的提交状态统计。")

        except Exception as e:
            self.log(f"❌ 批量检查失败: {str(e)}")
            messagebox.showerror("错误", f"批量检查失败:\n{str(e)}")
        finally:
            self.progress.stop()

    def validate_inputs(self):
        if not self.roster_var.get():
            messagebox.showerror("错误", "请选择花名册文件")
            return False
        if not self.homework_var.get():
            messagebox.showerror("错误", "请选择作业文件夹")
            return False
        if not self.output_var.get():
            messagebox.showerror("错误", "请选择输出目录")
            return False
        if not self.format_var.get():
            messagebox.showerror("错误", "请选择重命名格式")
            return False
        return True

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def load_config(self):
        config = self.config_manager.load_app_config()
        if config:
            self.roster_var.set(config.get('roster_path', ''))
            self.homework_var.set(config.get('homework_dir', ''))
            self.output_var.set(config.get('output_dir', ''))
            self.format_var.set(config.get('format_name', ''))

    def save_config(self):
        config = {
            'roster_path': self.roster_var.get(),
            'homework_dir': self.homework_var.get(),
            'output_dir': self.output_var.get(),
            'format_name': self.format_var.get()
        }
        self.config_manager.save_app_config(config)
        messagebox.showinfo("成功", "配置已保存！")

    def quick_setup(self):
        """
        【终极轮椅模式】一键配置。
        选择花名册 -> 自动分析列 -> 创建基础格式 -> 自动填入路径。
        """
        roster_path = filedialog.askopenfilename(
            title="选择你的花名册Excel文件",
            filetypes=[("Excel文件", "*.xls *.xlsx")]
        )
        if not roster_path:
            return

        try:
            # 读取并分析花名册
            df = pd.read_excel(roster_path, dtype={'学号': str})
            columns = df.columns.tolist()

            if '学号' not in columns or '姓名' not in columns:
                messagebox.showerror("错误", f"花名册必须包含‘学号’和‘姓名’列！\n当前列：{', '.join(columns)}")
                return

            # 1. 自动将花名册路径设置到主界面
            self.roster_var.set(roster_path)

            # 2. 将花名册的列信息传递给配置管理器，保存为“当前可用变量”
            self.config_manager.set_current_roster_columns(columns)

            # 3. 自动创建并保存几个最基础的格式（如果尚未存在）
            base_formats = {
                "标准格式(文件)": {"template": "{学号} {姓名}{扩展名}", "is_folder": False},
                "标准格式(文件夹)": {"template": "{学号} {姓名}", "is_folder": True},
            }
            # 可选：如果花名册有“班级”列，额外创建一个格式
            if '班级' in columns:
                base_formats["含班级格式"] = {"template": "{姓名}_{班级}{扩展名}", "is_folder": False}

            for name, config in base_formats.items():
                self.config_manager.save_format(name, config)

            # 4. 更新主界面的格式下拉框，并选中第一个
            self.refresh_formats()
            self.format_var.set("标准格式(文件)")

            # 5. 保存此次快速配置的状态（主要是花名册路径）
            self.save_config()

            messagebox.showinfo("配置成功",
                                f"花名册【{os.path.basename(roster_path)}】已载入！\n"
                                f"系统已识别可用变量：{', '.join(columns)}\n"
                                f"已为您创建了基础格式，可直接使用或点击‘管理格式’进行编辑。")

        except Exception as e:
            messagebox.showerror("配置失败", f"读取花名册时出错：{str(e)}")


class FormatManagerWindow:
    """格式管理窗口（新版：无需导入，直接点选）"""

    def __init__(self, parent, config_manager, refresh_callback):
        self.parent = parent
        self.config_manager = config_manager
        self.refresh_callback = refresh_callback
        self.available_vars = self.config_manager.get_current_roster_columns()  # 获取当前变量

        if not self.available_vars:
            messagebox.showerror("错误", "未检测到可用的花名册变量。请先导入花名册。")
            self.parent.destroy()
            return

        self.setup_ui()
        self.load_formats()

    def setup_ui(self):
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 左侧：现有格式列表 ---
        list_frame = ttk.LabelFrame(main_frame, text="现有格式列表")
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))

        self.format_listbox = tk.Listbox(list_frame, height=15, width=25)
        self.format_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)

        ttk.Button(list_btn_frame, text="添加", command=self.add_format).pack(fill=tk.X, pady=2)
        ttk.Button(list_btn_frame, text="编辑", command=self.edit_format).pack(fill=tk.X, pady=2)
        ttk.Button(list_btn_frame, text="删除", command=self.delete_format).pack(fill=tk.X, pady=2)

        # --- 右侧：变量与说明 ---
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 可用变量展示区
        var_frame = ttk.LabelFrame(right_frame, text=f"可用变量 (来自当前花名册)")
        var_frame.pack(fill=tk.X, pady=(0, 10))

        var_text = "， ".join([f"{{{col}}}" for col in self.available_vars])
        var_label = ttk.Label(var_frame, text=var_text, wraplength=400, justify=tk.LEFT)
        var_label.pack(padx=5, pady=5)

        # 使用说明
        help_frame = ttk.LabelFrame(right_frame, text="使用说明")
        help_frame.pack(fill=tk.BOTH, expand=True)

        help_content = """
        ✨ 【点选模式】使用指南 ✨

        1. 【添加】或【编辑】格式时，会打开编辑窗口。
        2. 在编辑窗口中，您可以从上方点击【变量按钮】来插入变量。
        3. 【常用片段】按钮可以帮助您快速组合基础格式。
        4. 手动在输入框中调整顺序或添加固定文字（如“_”、“-”）。

        📝 格式示例：
        • 文件：{学号} {姓名}{扩展名}
        • 文件：{姓名}_{班级}_作业{扩展名}
        • 文件夹：{学号} {姓名}
        • 文件夹：{项目组}_{姓名}

        ⚠️ 注意：
        • 文件格式必须包含 {扩展名}
        • 文件夹格式不能包含 {扩展名}
        """
        help_label = ttk.Label(help_frame, text=help_content, justify=tk.LEFT, wraplength=400)
        help_label.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)

    def load_formats(self):
        self.format_listbox.delete(0, tk.END)
        for name in self.config_manager.get_format_names():
            self.format_listbox.insert(tk.END, name)

    def add_format(self):
        self.edit_format_window(is_new=True)

    def edit_format(self):
        selection = self.format_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先在左侧列表中选择一个要编辑的格式。")
            return
        format_name = self.format_listbox.get(selection[0])
        self.edit_format_window(is_new=False, old_name=format_name)

    def edit_format_window(self, is_new=True, old_name=None):
        edit_window = tk.Toplevel(self.parent)
        edit_window.title("添加新格式" if is_new else f"编辑格式: {old_name}")
        edit_window.geometry("600x450")

        # 传递可用变量和配置管理器
        EditFormatWindow(edit_window, self.config_manager, self.available_vars,
                         old_name, is_new, self.load_formats, self.refresh_callback)

    def delete_format(self):
        selection = self.format_listbox.curselection()
        if not selection:
            return
        format_name = self.format_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除格式【{format_name}】吗？"):
            self.config_manager.delete_format(format_name)
            self.load_formats()
            self.refresh_callback()


class EditFormatWindow:
    """格式编辑窗口（新版：核心点选界面）"""

    def __init__(self, parent, config_manager, available_vars,
                 old_name, is_new, load_formats_callback, refresh_main_callback):
        self.parent = parent
        self.config_manager = config_manager
        self.available_vars = available_vars  # 可用变量列表
        self.old_name = old_name
        self.is_new = is_new
        self.load_formats_callback = load_formats_callback
        self.refresh_main_callback = refresh_main_callback

        # 如果是在编辑，则加载原有配置
        self.original_config = None if is_new else self.config_manager.get_format_config(old_name)

        self.setup_ui()
        if not is_new and self.original_config:
            self.load_existing_data()

    def setup_ui(self):
        main_frame = ttk.Frame(self.parent, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 格式名称
        ttk.Label(main_frame, text="格式名称：").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        self.name_var = tk.StringVar(value="新格式" if self.is_new else self.old_name)
        ttk.Entry(main_frame, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E),
                                                                         pady=(0, 10))

        # 2. 【变量按钮区】
        ttk.Label(main_frame, text="点击插入变量：").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        var_button_frame = ttk.Frame(main_frame)
        var_button_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

        # 创建变量按钮（每行最多5个）
        for i, var_name in enumerate(self.available_vars):
            btn = ttk.Button(var_button_frame, text=f"{{{var_name}}}",
                             command=lambda v=var_name: self.insert_text(f"{{{v}}}"), width=10)
            btn.grid(row=i // 5, column=i % 5, padx=2, pady=2)

        # 3. 【常用片段按钮区】
        ttk.Label(main_frame, text="常用片段：").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        snippet_frame = ttk.Frame(main_frame)
        snippet_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

        snippets = [
            ("学号+姓名", "{学号} {姓名}"),
            ("姓名+学号", "{姓名}_{学号}"),
            ("学号+姓名+班级", "{学号}_{姓名}_{班级}"),
        ]
        for i, (label, snippet) in enumerate(snippets):
            # 只展示当前花名册变量存在的片段
            if all(('{' + key + '}') in snippet for key in ['学号', '姓名']):  # 基础检查
                btn = ttk.Button(snippet_frame, text=label,
                                 command=lambda s=snippet: self.insert_text(s), width=15)
                btn.grid(row=0, column=i, padx=2, pady=2)

        # 4. 格式模板输入框
        ttk.Label(main_frame, text="格式模板：").grid(row=3, column=0, sticky=tk.W, pady=(0, 5))
        self.template_var = tk.StringVar()
        self.template_entry = tk.Text(main_frame, height=4, width=50)
        self.template_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

        # 5. 文件夹项目选项
        self.is_folder_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(main_frame, text="这是文件夹项目（不包含文件扩展名）",
                        variable=self.is_folder_var,
                        command=self.on_folder_toggle).grid(row=4, column=1, sticky=tk.W, pady=(0, 15))

        # 6. 操作按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="保存格式", command=self.save_format,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.parent.destroy).pack(side=tk.LEFT, padx=5)

        main_frame.columnconfigure(1, weight=1)
        # 绑定文件夹类型切换事件
        self.is_folder_var.trace('w', lambda *args: self.on_folder_toggle())

    def insert_text(self, text_to_insert):
        """向模板输入框的光标位置插入文本"""
        self.template_entry.focus_set()
        self.template_entry.insert(tk.INSERT, text_to_insert)

    def on_folder_toggle(self):
        """当切换文件夹类型时，自动处理{扩展名}"""
        current_text = self.template_entry.get("1.0", tk.END).strip()
        if self.is_folder_var.get():
            # 如果是文件夹，移除所有 {扩展名}
            new_text = current_text.replace("{扩展名}", "")
            self.template_entry.delete("1.0", tk.END)
            self.template_entry.insert("1.0", new_text)
        else:
            # 如果是文件，检查末尾是否有{扩展名}，没有则提示
            if not current_text.endswith("{扩展名}"):
                # 可以选择不自动添加，仅提示
                pass

    def load_existing_data(self):
        """加载已有格式的数据"""
        if self.original_config:
            self.template_entry.delete("1.0", tk.END)
            self.template_entry.insert("1.0", self.original_config.get('template', ''))
            self.is_folder_var.set(self.original_config.get('is_folder', False))

    def save_format(self):
        """保存格式"""
        name = self.name_var.get().strip()
        template = self.template_entry.get("1.0", tk.END).strip()
        is_folder = self.is_folder_var.get()

        if not name:
            messagebox.showerror("错误", "格式名称不能为空！")
            return
        if not template:
            messagebox.showerror("错误", "格式模板不能为空！")
            return

        # 基本验证
        if not is_folder and "{扩展名}" not in template:
            if not messagebox.askyesno("提示",
                                       "这是一个文件格式，但模板中没有包含 {扩展名} 变量。\n文件可能无法正常打开。确定继续吗？"):
                return
        if is_folder and "{扩展名}" in template:
            messagebox.showerror("错误", "文件夹格式不能包含 {扩展名} 变量！")
            return

        # 保存配置
        format_config = {
            'template': template,
            'is_folder': is_folder
        }

        # 如果是编辑且改名了，删除旧格式
        if not self.is_new and self.old_name != name:
            self.config_manager.delete_format(self.old_name)

        self.config_manager.save_format(name, format_config)

        # 刷新回调
        self.load_formats_callback()
        self.refresh_main_callback()

        self.parent.destroy()
        messagebox.showinfo("成功", f"格式【{name}】已保存！")


class FolderSelectorWindow:
    """子文件夹选择与排序窗口 (改进版：固定按钮+防撞车微调)"""

    def __init__(self, parent, parent_dir, all_folders, config_manager, update_callback):
        self.parent = parent
        self.parent_dir = parent_dir
        self.all_folders = sorted(all_folders)  # 初始按名称排序
        self.config_manager = config_manager
        self.update_callback = update_callback
        self.max_folders = len(all_folders)

        # 加载配置
        self.saved_config = self.config_manager.load_folder_config(self.parent_dir)
        if self.saved_config:
            self.selected_folders = self.saved_config.get('selected_folders', [])
            saved_order = self.saved_config.get('folder_order', [])
            # 重建序号映射
            self.order_mapping = {}  # folder -> order_num
            current_num = 1
            for folder in saved_order:
                if folder in self.all_folders:
                    self.order_mapping[folder] = current_num
                    current_num += 1
            # 为未排序的文件夹分配后续序号
            for folder in self.all_folders:
                if folder not in self.order_mapping:
                    self.order_mapping[folder] = current_num
                    current_num += 1
        else:
            self.selected_folders = self.all_folders.copy()
            self.order_mapping = {folder: i + 1 for i, folder in enumerate(self.all_folders)}

        # 新增：存储Spinbox变量引用
        self.order_vars = {}
        self.check_vars = {}
        self.spinboxes = {}

        self.setup_ui()

    def setup_ui(self):
        # 主窗口设置
        self.parent.title(f"选择并排序子文件夹 - {os.path.basename(self.parent_dir)}")
        self.parent.geometry("650x550")

        # 配置自定义样式
        style = ttk.Style()
        # 定义一个用于错误高亮的 Spinbox 样式
        style.configure('Error.TSpinbox', fieldbackground='#ffcccc')  # 浅红色背景

        # 主框架（分为上中下三部分，底部固定）
        main_container = ttk.Frame(self.parent)
        main_container.pack(fill=tk.BOTH, expand=True)

        # ===== 顶部：说明区域 =====
        top_frame = ttk.LabelFrame(main_container, text="操作说明", padding=(10, 5))
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        instructions = (
            "1. 在左侧勾选需要扫描的文件夹。\n"
            "2. 使用右侧的“▲/▼”按钮或直接输入数字调整排序序号（1~{max}）。\n"
            "3. 调整某个序号时，系统会自动处理重复的序号。\n"
            "4. 点击底部【应用选择】确认并关闭窗口。"
        ).format(max=self.max_folders)
        ttk.Label(top_frame, text=instructions, justify=tk.LEFT).pack(anchor=tk.W)

        # ===== 中部：可滚动的列表区域 =====
        middle_frame = ttk.Frame(main_container)
        middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 列表标题
        header_frame = ttk.Frame(middle_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header_frame, text="勾选", width=8, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="文件夹名称", width=35, anchor="w").pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="排序序号", width=12, anchor="center").pack(side=tk.LEFT, padx=2)

        # 带滚动条的Canvas
        list_canvas = tk.Canvas(middle_frame, highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(middle_frame, orient="vertical", command=list_canvas.yview)
        self.list_inner_frame = ttk.Frame(list_canvas)

        self.list_inner_frame.bind(
            "<Configure>",
            lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        )

        list_canvas.create_window((0, 0), window=self.list_inner_frame, anchor="nw")
        list_canvas.configure(yscrollcommand=v_scrollbar.set)

        # 布局滚动区域
        list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建列表行
        for folder in self.all_folders:
            self._create_list_row(folder)

        # ===== 底部：固定的操作按钮区域 =====
        bottom_frame = ttk.Frame(main_container)
        bottom_frame.pack(fill=tk.X, padx=10, pady=(5, 10))

        # 左侧批量操作按钮
        batch_btn_frame = ttk.Frame(bottom_frame)
        batch_btn_frame.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Button(batch_btn_frame, text="全选", width=8,
                   command=self._select_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame, text="清空", width=8,
                   command=self._select_none).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame, text="自动编号", width=10,
                   command=self._auto_number).pack(side=tk.LEFT)

        # 右侧主操作按钮（突出显示）
        action_btn_frame = ttk.Frame(bottom_frame)
        action_btn_frame.pack(side=tk.RIGHT)

        ttk.Button(action_btn_frame, text="取消",
                   command=self.parent.destroy, width=10).pack(side=tk.LEFT, padx=(0, 10))

        # 主要的“应用选择”按钮
        apply_btn = ttk.Button(action_btn_frame, text="应用选择",
                               command=self._apply_selection,
                               style="Accent.TButton", width=12)
        apply_btn.pack(side=tk.LEFT)

        # 配置突出按钮样式
        style.configure("Accent.TButton", font=('TkDefaultFont', 10, 'bold'))

        # 状态标签
        self.status_label = ttk.Label(bottom_frame, text="就绪", foreground="grey")
        self.status_label.pack(side=tk.LEFT, padx=(20, 0))

        # 绑定鼠标滚轮
        list_canvas.bind_all("<MouseWheel>",
                             lambda e: list_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # 初始聚焦
        self.parent.after(100, lambda: apply_btn.focus_set())

    def _create_list_row(self, folder):
        """为单个文件夹创建一行控件"""
        row_frame = ttk.Frame(self.list_inner_frame)
        row_frame.pack(fill=tk.X, padx=2, pady=1)

        # 复选框 - 绑定事件，勾选状态变化时更新编号逻辑
        is_selected = folder in self.selected_folders
        check_var = tk.BooleanVar(value=is_selected)
        self.check_vars[folder] = check_var
        check_btn = ttk.Checkbutton(row_frame, variable=check_var, width=6,
                                    command=lambda f=folder: self._on_checkbox_toggle(f))
        check_btn.pack(side=tk.LEFT)

        # 文件夹名标签
        ttk.Label(row_frame, text=folder, anchor="w", width=35).pack(side=tk.LEFT, padx=5)

        # Spinbox - 注意：初始只对已勾选的文件夹分配有效序号
        initial_value = str(self.order_mapping.get(folder, 0)) if is_selected else ""
        order_var = tk.StringVar(value=initial_value)

        # 自定义验证函数：允许空值或1~max_folders的数字
        vcmd = (self.parent.register(self._validate_spinbox_input), '%P')
        spinbox = ttk.Spinbox(row_frame, from_=1, to=self.max_folders,
                              textvariable=order_var, width=8,
                              validate='key', validatecommand=vcmd)
        spinbox.pack(side=tk.LEFT)

        # 存储引用
        self.spinboxes[folder] = spinbox
        self.order_vars[folder] = order_var  # 新增：存储变量引用以监听变化

        # 绑定事件：当焦点离开或按下回车时，进行最终处理
        spinbox.bind('<FocusOut>', lambda e, f=folder: self._finalize_order_change(f))
        spinbox.bind('<Return>', lambda e, f=folder: self._finalize_order_change(f))

        # 初始颜色设置
        self._update_spinbox_style(folder)

    def _validate_spinbox_input(self, new_value):
        """验证Spinbox输入是否有效 (允许空值)"""
        if new_value == "":
            return True
        if not new_value.isdigit():
            return False
        num = int(new_value)
        return 1 <= num <= self.max_folders

    def _on_checkbox_toggle(self, folder):
        """当复选框状态改变时的处理"""
        is_checked = self.check_vars[folder].get()

        if is_checked:
            # 被勾选：分配一个可用的最小序号
            used_numbers = {int(self.order_vars[f].get()) for f in self.check_vars
                            if self.check_vars[f].get() and self.order_vars[f].get().isdigit()}
            available = 1
            while available in used_numbers:
                available += 1
            self.order_vars[folder].set(str(available))
            self.order_mapping[folder] = available
        else:
            # 被取消勾选：清空序号
            self.order_vars[folder].set("")
            if folder in self.order_mapping:
                del self.order_mapping[folder]

        # 更新所有Spinbox的样式（检查重复）
        self._refresh_all_spinbox_styles()
        self.status_label.config(text=f"已{'勾选' if is_checked else '取消'} {folder}")

    def _on_spinbox_change(self, folder):
        """当通过微调按钮改变数值时的处理 - 仅标记，不解决冲突"""
        # 此方法现在只更新内部映射和检查重复（高亮显示），不触发重排
        current_value = self.spinboxes[folder].get().strip()

        if current_value == "":
            # 清空的情况
            if folder in self.order_mapping:
                del self.order_mapping[folder]
            self._refresh_all_spinbox_styles()
            return

        new_order = int(current_value)
        self.order_mapping[folder] = new_order
        # 立即更新显示样式（检查重复）
        self._refresh_all_spinbox_styles()
        # 状态栏可以给出提示，但不自动重排
        self.status_label.config(text=f"{folder} 序号改为 {new_order}。如有重复，请调整或点击“解决冲突”。")

    def _finalize_order_change(self, folder):
        """
        当用户完成对一个序号框的编辑（焦点离开或回车）时，
        检查并解决编号冲突，进行最终结算。
        """
        if not self.check_vars[folder].get():
            # 如果此项未被勾选，忽略
            return

        current_value = self.order_vars[folder].get().strip()
        if not current_value or not current_value.isdigit():
            return

        new_order = int(current_value)
        old_order = self.order_mapping.get(folder)

        if old_order == new_order:
            return  # 序号未变

        # 1. 找出所有序号冲突（重复项）
        order_to_folders = {}
        for f, var in self.order_vars.items():
            if not self.check_vars[f].get():
                continue
            val = var.get().strip()
            if val and val.isdigit():
                num = int(val)
                order_to_folders.setdefault(num, []).append(f)

        duplicates = {num: folders for num, folders in order_to_folders.items() if len(folders) > 1}

        if new_order not in duplicates:
            # 没有冲突，直接更新
            self.order_mapping[folder] = new_order
            self._refresh_all_spinbox_styles()
            return

        # 2. 解决冲突：当前文件夹获得该序号，其他重复项需要重新分配
        conflict_folders = [f for f in duplicates[new_order] if f != folder]

        # 为每个冲突文件夹寻找新的可用序号
        used_numbers = set(self.order_mapping.values())
        for conflict_folder in conflict_folders:
            available = 1
            while available in used_numbers:
                available += 1
            # 更新冲突文件夹的序号
            self.order_mapping[conflict_folder] = available
            self.order_vars[conflict_folder].set(str(available))
            used_numbers.add(available)
            self.status_label.config(text=f"已为 {conflict_folder} 重新分配序号 {available}")

        # 3. 最后更新当前文件夹的映射
        self.order_mapping[folder] = new_order
        self._refresh_all_spinbox_styles()

    def _refresh_all_spinbox_styles(self):
        """更新所有Spinbox的样式，用颜色高亮重复的序号"""
        # 统计当前所有已输入的序号
        order_count = {}
        for folder, var in self.order_vars.items():
            if not self.check_vars[folder].get():
                continue  # 只统计已勾选的
            val = var.get().strip()
            if val and val.isdigit():
                num = int(val)
                order_count[num] = order_count.get(num, 0) + 1

        # 根据重复状态设置样式
        for folder, spinbox in self.spinboxes.items():
            val = self.order_vars[folder].get().strip()
            if not val or not val.isdigit():
                spinbox.configure(style='TSpinbox')  # 默认样式
                continue

            num = int(val)
            if order_count.get(num, 0) > 1:
                # 重复序号：红色背景警示
                spinbox.configure(style='Error.TSpinbox')
            else:
                # 唯一序号：正常样式
                spinbox.configure(style='TSpinbox')

    def _update_spinbox_style(self, folder):
        """更新单个Spinbox的样式"""
        val = self.order_vars[folder].get().strip()
        if not val or not val.isdigit():
            self.spinboxes[folder].configure(style='TSpinbox')
            return

        # 统计该序号出现的次数
        num = int(val)
        count = 0
        for f, var in self.order_vars.items():
            if self.check_vars[f].get() and var.get().strip() == str(num):
                count += 1

        if count > 1:
            self.spinboxes[folder].configure(style='Error.TSpinbox')
        else:
            self.spinboxes[folder].configure(style='TSpinbox')

    def _select_all(self):
        """全选"""
        for var in self.check_vars.values():
            var.set(True)
            # 触发勾选事件以分配序号
            folder = next(f for f, v in self.check_vars.items() if v == var)
            self._on_checkbox_toggle(folder)
        self.status_label.config(text="已全选所有文件夹")

    def _select_none(self):
        """清空选择"""
        for var in self.check_vars.values():
            var.set(False)
            # 触发取消勾选事件以清空序号
            folder = next(f for f, v in self.check_vars.items() if v == var)
            self._on_checkbox_toggle(folder)
        self.status_label.config(text="已清空选择")

    def _auto_number(self):
        """为已勾选的文件夹自动编号（从1开始连续），并解决所有冲突"""
        selected_folders = [f for f, var in self.check_vars.items() if var.get()]

        if not selected_folders:
            self.status_label.config(text="请先勾选文件夹", foreground="orange")
            return

        # 直接分配连续序号
        for index, folder in enumerate(selected_folders, start=1):
            self.order_vars[folder].set(str(index))
            self.order_mapping[folder] = index

        self._refresh_all_spinbox_styles()
        self.status_label.config(text=f"已为 {len(selected_folders)} 个勾选文件夹分配连续序号")

    def _apply_selection(self):
        """应用选择并关闭窗口"""
        # 收集最终选择
        final_selected = [f for f, var in self.check_vars.items() if var.get()]

        if not final_selected:
            if not messagebox.askyesno("确认", "未选择任何文件夹，确定要继续吗？"):
                return

        # 收集勾选文件夹的序号映射
        final_order_mapping = {}
        for folder in self.all_folders:
            if self.check_vars[folder].get():
                val = self.order_vars[folder].get().strip()
                if val and val.isdigit():
                    final_order_mapping[folder] = int(val)

        # 按序号升序排序（关键修改）
        sorted_items = sorted(final_order_mapping.items(), key=lambda x: x[1])
        final_ordered_folders = [f for f, _ in sorted_items]

        # 保存配置
        config = {
            'selected_folders': final_ordered_folders,  # 已经是排序后的
            'folder_order': final_ordered_folders,
            'order_mapping': final_order_mapping,
            'total_folders': len(self.all_folders)
        }

        try:
            self.config_manager.save_folder_config(self.parent_dir, config)

            # 更新主界面显示
            self.update_callback(final_ordered_folders)

            # 显示顺序确认
            order_info = "\n".join([f"{i + 1}. {folder}" for i, folder in enumerate(final_ordered_folders[:10])])
            if len(final_ordered_folders) > 10:
                order_info += f"\n... 等 {len(final_ordered_folders)} 个文件夹"

            messagebox.showinfo("选择已保存",
                                f"✅ 已保存 {len(final_ordered_folders)} 个文件夹\n\n"
                                f"Excel中的列顺序为：\n{order_info}")
            self.parent.destroy()

        except Exception as e:
            messagebox.showerror("保存失败", f"保存配置时出错:\n{str(e)}")