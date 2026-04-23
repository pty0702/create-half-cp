import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd


class DictionaryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("科目字典检索系统 v4.2")
        self.geometry("600x550")  # 调整为主流的竖向舒适比例

        self.data_file = "local_dict.json"
        self.subject_dict = {}

        self.load_data()
        self.init_menu()
        self.init_main_ui()

    # ================= 数据持久化 =================
    def load_data(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.subject_dict = json.load(f)
            except Exception as e:
                messagebox.showerror("读取错误", f"无法读取本地字典文件：\n{e}")

    def save_data(self):
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.subject_dict, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("保存错误", f"无法保存数据到本地字典文件：\n{e}")

    # ================= 菜单栏初始化 =================
    def init_menu(self):
        menubar = tk.Menu(self)

        # 1. 文件菜单 (保留备用)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="退出程序", command=self.quit)
        menubar.add_cascade(label="文件", menu=file_menu)

        # 2. 字典设置菜单
        dict_menu = tk.Menu(menubar, tearoff=0)
        dict_menu.add_command(label="字典数据管理 (增删改)", command=self.open_manager_window)
        dict_menu.add_separator()
        dict_menu.add_command(label="生成字典模板", command=self.generate_template)
        dict_menu.add_command(label="导入模板 (全量同步)", command=self.import_template)
        menubar.add_cascade(label="字典设置", menu=dict_menu)

        self.config(menu=menubar)

    # ================= 主界面初始化 (上下排版) =================
    def init_main_ui(self):
        # 搜索区 (上下排列)
        search_frame = ttk.LabelFrame(self, text="快速检索", padding=15)
        search_frame.pack(fill=tk.X, padx=15, pady=10)

        # 第一排：关键字输入
        ttk.Label(search_frame, text="输入关键字:").grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_main)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, padx=5, pady=8, sticky=tk.EW)

        # 第二排：匹配结果下拉框
        ttk.Label(search_frame, text="匹配结果:").grid(row=1, column=0, padx=5, pady=8, sticky=tk.W)
        self.result_combo = ttk.Combobox(search_frame, state="readonly")
        self.result_combo.grid(row=1, column=1, padx=5, pady=8, sticky=tk.EW)
        self.result_combo.bind('<<ComboboxSelected>>', self.on_combo_select)

        # 让第二列（输入框和下拉框）自动拉伸填满剩余宽度
        search_frame.columnconfigure(1, weight=1)

        # 详细信息展示区
        detail_frame = ttk.LabelFrame(self, text="科目详细信息与提取", padding=15)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        self.fields = {
            "name": {"label": "科目名称", "var": tk.StringVar()},
            "full_code": {"label": "完整科目编码", "var": tk.StringVar()},
            "cat_code": {"label": "分类代码(7-12位)", "var": tk.StringVar()},
            "spec_code": {"label": "规格代码(13-15位)", "var": tk.StringVar()}
        }

        row = 0
        for key, field_info in self.fields.items():
            ttk.Label(detail_frame, text=field_info["label"] + ":", font=('Microsoft YaHei', 10, 'bold')).grid(row=row,
                                                                                                               column=0,
                                                                                                               sticky=tk.W,
                                                                                                               pady=10)
            entry = ttk.Entry(detail_frame, textvariable=field_info["var"], width=35, state='readonly',
                              font=('Microsoft YaHei', 11))
            entry.grid(row=row, column=1, padx=15, pady=10, sticky=tk.EW)
            btn = ttk.Button(detail_frame, text="提取复制",
                             command=lambda v=field_info["var"]: self.copy_to_clipboard(v.get()))
            btn.grid(row=row, column=2, padx=5, pady=10)
            row += 1

        detail_frame.columnconfigure(1, weight=1)

        # 底部状态栏
        self.status_var = tk.StringVar()
        self.refresh_status()
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.on_search_main()

    def refresh_status(self):
        self.status_var.set(f"就绪 - 当前字典条目数: {len(self.subject_dict)}")

    def on_search_main(self, *args):
        search_term = self.search_var.get().strip().lower()
        matches = []
        for name, code in self.subject_dict.items():
            # 修改点：只匹配科目名称，不再匹配编码
            if search_term in name.lower():
                matches.append(f"{name} | {code}")

        self.result_combo['values'] = matches
        if matches:
            self.result_combo.current(0)
            self.on_combo_select()
        else:
            self.result_combo.set("无匹配结果")
            self.clear_main_inputs()

    def on_combo_select(self, event=None):
        selected = self.result_combo.get()
        if not selected or " | " not in selected:
            return

        name, code = selected.split(" | ")
        self.fields["name"]["var"].set(name)
        self.fields["full_code"]["var"].set(code)

        # 长度不足时直接置空
        if len(code) >= 12:
            self.fields["cat_code"]["var"].set(code[6:12])
        else:
            self.fields["cat_code"]["var"].set("")

        if len(code) >= 15:
            self.fields["spec_code"]["var"].set(code[12:15])
        else:
            self.fields["spec_code"]["var"].set("")

    def clear_main_inputs(self):
        for field in self.fields.values():
            field["var"].set("")

    def copy_to_clipboard(self, text):
        if text:  # 只有文本非空时才复制
            self.clipboard_clear()
            self.clipboard_append(text)
            self.status_var.set(f"已成功复制: {text}")
        else:
            self.status_var.set("当前字段为空，无内容可复制")

    # ================= 字典数据管理器 (上下排版) =================
    def open_manager_window(self):
        manager_win = tk.Toplevel(self)
        manager_win.title("字典数据管理")
        manager_win.geometry("600x650")  # 设置为较修长的窗口
        manager_win.transient(self)
        manager_win.grab_set()

        mgr_search_var = tk.StringVar()
        mgr_name_var = tk.StringVar()
        mgr_code_var = tk.StringVar()
        self.current_selected_name_in_mgr = None

        # 上下分栏
        paned = ttk.PanedWindow(manager_win, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # ====== 上半部：列表与搜索区 ======
        top_frame = ttk.Frame(paned)
        paned.add(top_frame, weight=3)  # 列表区占据更多空间

        search_entry = ttk.Entry(top_frame, textvariable=mgr_search_var)
        search_entry.pack(fill=tk.X, pady=(0, 5))
        search_entry.insert(0, "在这里搜索...")
        search_entry.bind("<FocusIn>", lambda args: search_entry.delete('0', 'end') if search_entry.get() == "在这里搜索..." else None)

        scrollbar = ttk.Scrollbar(top_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox = tk.Listbox(top_frame, yscrollcommand=scrollbar.set, font=('Microsoft YaHei', 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        # ====== 下半部：映射编辑区 ======
        bottom_frame = ttk.LabelFrame(paned, text="映射编辑区", padding=15)
        paned.add(bottom_frame, weight=1)

        ttk.Label(bottom_frame, text="科目名称:").grid(row=0, column=0, pady=8, sticky=tk.W)
        name_entry = ttk.Entry(bottom_frame, textvariable=mgr_name_var)
        name_entry.grid(row=0, column=1, pady=8, padx=10, sticky=tk.EW)

        ttk.Label(bottom_frame, text="科目编码:").grid(row=1, column=0, pady=8, sticky=tk.W)
        code_entry = ttk.Entry(bottom_frame, textvariable=mgr_code_var)
        code_entry.grid(row=1, column=1, pady=8, padx=10, sticky=tk.EW)

        bottom_frame.columnconfigure(1, weight=1)  # 让输入框填满横向空间

        # 操作按钮区
        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(15, 0), sticky=tk.EW)

        ttk.Button(btn_frame, text="新增映射", command=lambda: action_add()).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        ttk.Button(btn_frame, text="修改保存", command=lambda: action_update()).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        ttk.Button(btn_frame, text="删除选中", command=lambda: action_delete()).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # 按钮功能函数
        def refresh_listbox(*args):
            term = mgr_search_var.get().strip().lower()
            if term == "在这里搜索...": term = ""
            listbox.delete(0, tk.END)
            for n, c in self.subject_dict.items():
                # 修改点：字典管理器中的搜索也只匹配科目名称
                if term in n.lower():
                    listbox.insert(tk.END, f"{n} | {c}")

        def on_mgr_select(event):
            selection = listbox.curselection()
            if not selection: return
            item_text = listbox.get(selection[0])
            try:
                n, c = item_text.split(" | ")
                self.current_selected_name_in_mgr = n
                mgr_name_var.set(n)
                mgr_code_var.set(c)
            except:
                pass

        def action_add():
            n, c = mgr_name_var.get().strip(), mgr_code_var.get().strip()
            if not n or not c: return messagebox.showwarning("提示", "名称和编码不能为空", parent=manager_win)
            if n in self.subject_dict: return messagebox.showwarning("提示", "科目已存在，请使用修改", parent=manager_win)
            self.subject_dict[n] = c
            save_and_sync()
            messagebox.showinfo("成功", "新增成功", parent=manager_win)

        def action_update():
            new_n, new_c = mgr_name_var.get().strip(), mgr_code_var.get().strip()
            if not self.current_selected_name_in_mgr: return messagebox.showwarning("提示", "请先从左侧选择", parent=manager_win)
            if not new_n or not new_c: return messagebox.showwarning("提示", "名称和编码不能为空", parent=manager_win)

            if self.current_selected_name_in_mgr != new_n:
                if new_n in self.subject_dict and not messagebox.askyesno("冲突", "目标名称已存在，是否覆盖？", parent=manager_win):
                    return
                del self.subject_dict[self.current_selected_name_in_mgr]

            self.subject_dict[new_n] = new_c
            self.current_selected_name_in_mgr = new_n
            save_and_sync()
            messagebox.showinfo("成功", "更新成功", parent=manager_win)

        def action_delete():
            if not self.current_selected_name_in_mgr: return messagebox.showwarning("提示", "请先从左侧选择", parent=manager_win)
            if messagebox.askyesno("确认", f"删除 '{self.current_selected_name_in_mgr}'？", parent=manager_win):
                del self.subject_dict[self.current_selected_name_in_mgr]
                mgr_name_var.set("")
                mgr_code_var.set("")
                self.current_selected_name_in_mgr = None
                save_and_sync()

        def save_and_sync():
            self.save_data()
            refresh_listbox()
            self.on_search_main()
            self.refresh_status()

        mgr_search_var.trace_add("write", refresh_listbox)
        listbox.bind('<<ListboxSelect>>', on_mgr_select)
        refresh_listbox()

    # ================= Excel 导入导出模板功能 =================
    def generate_template(self):
        file_path = os.path.join(os.getcwd(), "科目字典模板.xlsx")
        try:
            df = pd.DataFrame({"科目名称": ["示例：吡唑酯．精甲霜．甲维2．9％"], "科目编码": ["140402010104520"]})
            df.to_excel(file_path, index=False)
            messagebox.showinfo("成功", f"模板已生成：\n{file_path}")
        except Exception as e:
            messagebox.showerror("失败", f"生成出错：\n{e}")

    def import_template(self):
        file_path = filedialog.askopenfilename(title="选择模板", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if not file_path: return
        try:
            df = pd.read_excel(file_path, dtype=str)
        except Exception as e:
            return messagebox.showerror("导入失败", f"读取失败：\n{e}")

        if "科目名称" not in df.columns or "科目编码" not in df.columns:
            return messagebox.showerror("格式错误", "必须包含“科目名称”和“科目编码”表头。")

        df = df.dropna(subset=["科目名称", "科目编码"])
        imported_names = set(df["科目名称"].str.strip().tolist())

        new_added = []
        conflicts = []
        code_to_name = {v: k for k, v in self.subject_dict.items()}

        for _, row in df.iterrows():
            name, code = row["科目名称"].strip(), row["科目编码"].strip()
            if name in self.subject_dict:
                if self.subject_dict[name] != code: conflicts.append((name, name, code, "新编码"))
            elif code in code_to_name:
                if code_to_name[code] != name: conflicts.append((code_to_name[code], name, code, "新名称"))
            else:
                new_added.append((name, code))
                self.subject_dict[name] = code
                code_to_name[code] = name

        conflict_resolved_count = 0
        if conflicts and messagebox.askyesno("更新提示", f"发现 {len(conflicts)} 条修改记录，是否更新覆盖？"):
            for old_n, new_n, new_c, reason in conflicts:
                if reason == "新名称" and old_n in self.subject_dict: del self.subject_dict[old_n]
                self.subject_dict[new_n] = new_c
                conflict_resolved_count += 1

        to_delete = [n for n in self.subject_dict if n not in imported_names]
        deleted_count = 0
        if to_delete and messagebox.askyesno("同步删除", f"Excel中删除了 {len(to_delete)} 个现存条目，是否同步删除？"):
            for name in to_delete: del self.subject_dict[name]
            deleted_count = len(to_delete)

        self.save_data()
        self.on_search_main()
        self.refresh_status()

        msg = f"新增: {len(new_added)} 条\n更新: {conflict_resolved_count} 条\n同步删除: {deleted_count} 条"
        messagebox.showinfo("导入完成",
                            msg if (new_added or conflict_resolved_count or deleted_count) else "字典未发生变化。")


if __name__ == "__main__":
    app = DictionaryApp()
    app.mainloop()