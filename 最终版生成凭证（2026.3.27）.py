# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Menu, simpledialog
from datetime import datetime, timedelta
import json
import os
import pandas as pd
import csv
import re

# 尝试导入拖拽库，如果未安装则静默降级
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except ImportError:
    HAS_DND = False

BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


class VoucherFStringApp:
    def __init__(self, root):
        self.root = root
        self.root.title(
            "北农凭证标准生成器 - 最终修正版" + (" (支持拖拽解析)" if HAS_DND else " (未安装tkinterdnd2, 不支持拖拽)"))
        self.root.geometry("1100x850")

        self.dict_file = "subject_dict.json"
        self.state_file = "unsaved_state.json"
        self.draft_file = "draft_box.json"

        self.subject_dict = self.load_dict()
        self.line_counter = 1
        self._search_timer = None

        self.mode_var = tk.StringVar(value="transfer")

        self.setup_ui()
        self.check_restore_state()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.handle_drop)

    # --- 1. 字典管理 (增删改查 UI) ---
    def load_dict(self):
        if os.path.exists(self.dict_file):
            try:
                with open(self.dict_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        else:
            return {"11%晶格嘧菌酯": "1403001"}

    def save_dict(self):
        try:
            with open(self.dict_file, 'w', encoding='utf-8') as f:
                json.dump(self.subject_dict, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"字典保存失败: {e}")

    def open_dict_manager(self):
        mgr = tk.Toplevel(self.root)
        mgr.title("字典映射管理器")
        mgr.geometry("500x500")

        tree_frame = ttk.Frame(mgr)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("name", "code")
        dtree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        dtree.heading("name", text="科目名称")
        dtree.heading("code", text="科目编码")
        dtree.column("name", width=250)
        dtree.column("code", width=150)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=dtree.yview)
        dtree.configure(yscrollcommand=scrollbar.set)
        dtree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def refresh_list():
            for i in dtree.get_children(): dtree.delete(i)
            for k, v in self.subject_dict.items():
                dtree.insert("", "end", values=(k, v))

        refresh_list()

        edit_frame = ttk.Frame(mgr)
        edit_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(edit_frame, text="名称:").grid(row=0, column=0, padx=5)
        ent_name = ttk.Entry(edit_frame, width=20)
        ent_name.grid(row=0, column=1, padx=5)

        ttk.Label(edit_frame, text="编码:").grid(row=0, column=2, padx=5)
        ent_code = ttk.Entry(edit_frame, width=15)
        ent_code.grid(row=0, column=3, padx=5)

        def on_select(e):
            sel = dtree.selection()
            if not sel: return
            vals = dtree.item(sel[0], "values")
            ent_name.delete(0, tk.END);
            ent_name.insert(0, vals[0])
            ent_code.delete(0, tk.END);
            ent_code.insert(0, vals[1])

        dtree.bind("<<TreeviewSelect>>", on_select)

        def save_mapping():
            n, c = ent_name.get().strip(), ent_code.get().strip()
            if not n or not c: return messagebox.showwarning("提示", "名称和编码不能为空")
            self.subject_dict[n] = c
            self.save_dict()
            refresh_list()
            ent_name.delete(0, tk.END);
            ent_code.delete(0, tk.END)

        def delete_mapping():
            sel = dtree.selection()
            if not sel: return
            vals = dtree.item(sel[0], "values")
            if vals[0] in self.subject_dict:
                del self.subject_dict[vals[0]]
                self.save_dict()
                refresh_list()
                ent_name.delete(0, tk.END);
                ent_code.delete(0, tk.END)

        btn_frame = ttk.Frame(mgr)
        btn_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(btn_frame, text="保存/更新选中", command=save_mapping).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="删除选中", command=delete_mapping).pack(side="left", padx=5)

    # --- 2. 界面构建 ---
    def setup_ui(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        dict_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="字典管理", menu=dict_menu)
        dict_menu.add_command(label="打开字典管理器...", command=self.open_dict_manager)
        dict_menu.add_separator()
        dict_menu.add_command(label="导出字典 (Excel)", command=self.export_excel_dict)
        dict_menu.add_command(label="导入字典 (Excel)", command=self.import_excel_dict)

        cfg = ttk.LabelFrame(self.root, text="公共信息 & 模式选择", padding=10)
        cfg.pack(fill="x", padx=15, pady=10)

        tk.Label(cfg, text="当前业务模式:", font=("微软雅黑", 10, "bold"), fg="blue").grid(row=0, column=0, sticky="e")
        rb_frame = ttk.Frame(cfg)
        rb_frame.grid(row=0, column=1, columnspan=3, sticky="w", padx=5)
        ttk.Radiobutton(rb_frame, text="半成品入库 (结转)", variable=self.mode_var, value="transfer").pack(side="left",
                                                                                                           padx=10)
        ttk.Radiobutton(rb_frame, text="半成品领用 (使用)", variable=self.mode_var, value="usage").pack(side="left",
                                                                                                        padx=10)

        ttk.Label(cfg, text="凭证日期:").grid(row=1, column=0, pady=5)
        self.ent_f1 = ttk.Entry(cfg, width=15)
        self.ent_f1.insert(0, datetime.now().strftime("%Y%m%d"))
        self.ent_f1.grid(row=1, column=1, padx=5)
        self.ent_f1.bind("<FocusOut>", self.auto_format_f1)
        self.ent_f1.bind("<Return>", self.auto_format_f1)

        ttk.Label(cfg, text="期间(F52):").grid(row=1, column=2, padx=10)
        self.ent_period = ttk.Entry(cfg, width=8)
        self.ent_period.grid(row=1, column=3, padx=5)

        ttk.Label(cfg, text="制单人:").grid(row=1, column=4, padx=10)
        self.ent_maker = ttk.Entry(cfg, width=12)
        self.ent_maker.insert(0, "潘天宇")
        self.ent_maker.grid(row=1, column=5)

        ttk.Label(cfg, text="摘要日期:").grid(row=2, column=0, pady=5)
        self.ent_sum_date = ttk.Entry(cfg, width=18)
        self.ent_sum_date.grid(row=2, column=1)
        self.ent_sum_date.bind("<FocusOut>", self.auto_format_sum_date)
        self.ent_sum_date.bind("<Return>", self.auto_format_sum_date)

        ttk.Label(cfg, text="账套:").grid(row=2, column=2, padx=10)
        self.ent_set = ttk.Entry(cfg, width=10)
        self.ent_set.insert(0, "001")
        self.ent_set.grid(row=2, column=3)

        input_f = ttk.LabelFrame(self.root, text="分录录入 (支持模糊搜索)", padding=10)
        input_f.pack(fill="x", padx=15, pady=5)

        ttk.Label(input_f, text="科目名称(搜):").grid(row=0, column=0)
        self.name_var = tk.StringVar()
        self.name_in = ttk.Entry(input_f, textvariable=self.name_var, width=20)
        self.name_in.grid(row=0, column=1, padx=5)
        self.name_in.bind('<KeyRelease>', self.schedule_search)

        ttk.Label(input_f, text="科目编码:").grid(row=0, column=2)
        self.code_in = ttk.Entry(input_f, width=15)
        self.code_in.grid(row=0, column=3, padx=5)

        ttk.Label(input_f, text="数量:").grid(row=0, column=4)
        self.qty_in = ttk.Entry(input_f, width=10)
        self.qty_in.grid(row=0, column=5, padx=5)

        ttk.Label(input_f, text="单价:").grid(row=0, column=6)
        self.price_in = ttk.Entry(input_f, width=10)
        self.price_in.grid(row=0, column=7, padx=5)

        ttk.Button(input_f, text="添加此分录对", command=self.add_to_tree).grid(row=0, column=8, padx=15)
        self.price_in.bind('<Return>', lambda e: self.add_to_tree())

        self.suggestion_list = tk.Listbox(self.root, height=5)
        self.suggestion_list.bind('<<ListboxSelect>>', self.on_select_suggestion)

        # --- 表格列顺序调整 ---
        self.tree = ttk.Treeview(self.root, columns=("line", "sum", "code", "name", "qty", "price", "db", "cr"),
                                 show="headings")
        columns = ["行号", "摘要", "科目编码", "科目名称", "数量", "单价", "借方", "贷方"]
        widths = [40, 240, 100, 140, 60, 80, 80, 80]
        for c, n, w in zip(self.tree["columns"], columns, widths):
            self.tree.heading(c, text=n)
            self.tree.column(c, width=w, anchor="center" if c == "line" else "w")
        self.tree.pack(fill="both", expand=True, padx=15, pady=10)

        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="修改此分录 (双击)", command=self.edit_selected_entry)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="仅修改摘要文字", command=self.modify_summary_text_only)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected_entry())

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=15, pady=10)

        ttk.Button(btn_frame, text="新建下一张 (清空前提示)", command=self.reset_voucher).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="暂存至草稿箱", command=self.suspend_voucher).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="打开草稿箱", command=self.open_draft_box).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="读取TXT还原 (支持拖拽)", command=self.verify_txt_file).pack(side="left", padx=15)

        ttk.Button(btn_frame, text="导出 TXT", padding=10, command=self.export_txt).pack(side="right", padx=5)

        self.auto_format_f1()

    # --- 3. 草稿箱 ---
    def get_drafts(self):
        if os.path.exists(self.draft_file):
            try:
                with open(self.draft_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}

    def save_drafts(self, drafts):
        try:
            with open(self.draft_file, 'w', encoding='utf-8') as f:
                json.dump(drafts, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("保存失败", f"草稿箱保存失败: {e}")

    def suspend_voucher(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("提示", "当前没有录入任何分录，无需暂存。")
            return

        default_name = f"{self.ent_f1.get()} 暂存的凭证"
        draft_name = simpledialog.askstring("暂存凭证", "请为这份草稿起个名字以便后续查找：", initialvalue=default_name)

        if not draft_name: return

        state = {
            "mode": self.mode_var.get(),
            "f1": self.ent_f1.get(),
            "period": self.ent_period.get(),
            "sum_date": self.ent_sum_date.get(),
            "maker": self.ent_maker.get(),
            "set": self.ent_set.get(),
            "items": [self.tree.item(i, "values") for i in items]
        }

        drafts = self.get_drafts()
        draft_id = datetime.now().strftime("%Y%m%d%H%M%S")
        drafts[draft_id] = {"name": draft_name, "state": state}
        self.save_drafts(drafts)

        for i in self.tree.get_children(): self.tree.delete(i)
        self.line_counter = 1
        self.clear_inputs()

        messagebox.showinfo("成功", f"凭证已暂存至草稿箱：【{draft_name}】\n界面已清空，您可以开始做下一张凭证了。")

    def open_draft_box(self):
        drafts = self.get_drafts()
        if not drafts:
            messagebox.showinfo("草稿箱", "当前草稿箱是空的，快去存几个试试吧。")
            return

        dbx = tk.Toplevel(self.root)
        dbx.title("草稿箱管理")
        dbx.geometry("400x350")

        tk.Label(dbx, text="双击或选中后点击下方按钮恢复:", fg="gray").pack(pady=5)

        lb = tk.Listbox(dbx, selectmode=tk.SINGLE, font=("微软雅黑", 10))
        lb.pack(fill="both", expand=True, padx=15, pady=5)

        draft_ids = list(drafts.keys())
        for did in draft_ids:
            lb.insert(tk.END, f" 📝 {drafts[did]['name']}")

        def load_selected(event=None):
            sel = lb.curselection()
            if not sel: return

            if self.tree.get_children():
                if not messagebox.askyesno("覆盖警告",
                                           "当前界面有未导出的内容！\n强制恢复草稿将【清空并替换】当前界面的内容。确定继续吗？"):
                    return

            did = draft_ids[sel[0]]
            state = drafts[did]["state"]

            for i in self.tree.get_children(): self.tree.delete(i)
            self.mode_var.set(state.get("mode", "transfer"))
            self.ent_f1.delete(0, tk.END);
            self.ent_f1.insert(0, state.get("f1", ""))
            self.ent_period.delete(0, tk.END);
            self.ent_period.insert(0, state.get("period", ""))
            self.ent_sum_date.delete(0, tk.END);
            self.ent_sum_date.insert(0, state.get("sum_date", ""))

            for val in state.get("items", []):
                self.tree.insert("", "end", values=val)
            self.renumber_rows()

            if messagebox.askyesno("加载成功",
                                   "草稿已成功恢复到主界面。\n是否将该凭证从草稿箱中彻底【删除】？(选是清理，选否保留)"):
                del drafts[did]
                self.save_drafts(drafts)

            dbx.destroy()

        def delete_selected():
            sel = lb.curselection()
            if not sel: return
            did = draft_ids[sel[0]]
            draft_name = drafts[did]['name']

            if messagebox.askyesno("删除确认", f"确定要彻底删除草稿【{draft_name}】吗？\n删除后不可恢复！"):
                del drafts[did]
                self.save_drafts(drafts)
                dbx.destroy()
                self.open_draft_box()

        lb.bind("<Double-1>", load_selected)

        bf = ttk.Frame(dbx)
        bf.pack(fill="x", padx=15, pady=10)
        ttk.Button(bf, text="恢复选中草稿", command=load_selected).pack(side="left", padx=5)
        ttk.Button(bf, text="删除废弃草稿", command=delete_selected).pack(side="right", padx=5)

    # --- 原有核心逻辑 (更新了录入数组的顺序: 行号, 摘要, 编码, 名称, 数量, 单价, 借方, 贷方) ---
    def add_to_tree(self):
        try:
            name = self.name_var.get().strip()
            dynamic_code = self.code_in.get().strip()
            qty = float(self.qty_in.get() or 0)
            price = float(self.price_in.get() or 0)
            amt = round(qty * price, 2)

            s_date = self.ent_sum_date.get()
            f1_raw = self.ent_f1.get()
            try:
                dt_tmp = datetime.strptime(f1_raw, "%Y-%m-%d")
                vid = dt_tmp.strftime("%Y%m%d")
            except:
                vid = f1_raw.replace("-", "").zfill(8)

            if name and dynamic_code:
                if name not in self.subject_dict or self.subject_dict[name] != dynamic_code:
                    self.subject_dict[name] = dynamic_code
                    self.save_dict()

            mode = self.mode_var.get()

            # 新增时调整了插入数组的顺序: 0:line, 1:sum, 2:code, 3:name, 4:qty, 5:price, 6:db, 7:cr
            if mode == "transfer":
                sum_db = f"{s_date}统计本期在产品(单号:{vid})"
                self.tree.insert("", "end", values=(
                    self.line_counter, sum_db, dynamic_code, name,
                    self.clean_num(qty), self.clean_num(price), self.clean_num(amt), 0
                ))
                self.line_counter += 1

                sum_cr = f"{s_date}预计本期在产品成本(单号:{vid})"
                calc_code = f"500101{dynamic_code[-6:]}998"
                self.tree.insert("", "end", values=(
                    self.line_counter, sum_cr, calc_code, "自制半成品成本",
                    self.clean_num(qty), "-", 0, self.clean_num(amt)
                ))
                self.line_counter += 1

            elif mode == "usage":
                sum_db = f"{s_date}产品生产使用半产品成本(单号:{vid})"
                fixed_code = "5001019907"
                self.tree.insert("", "end", values=(
                    self.line_counter, sum_db, fixed_code, "半产品成本",
                    0, "-", self.clean_num(amt), 0
                ))
                self.line_counter += 1

                sum_cr = f"{s_date}车间使用半成品(单号:{vid})"
                self.tree.insert("", "end", values=(
                    self.line_counter, sum_cr, dynamic_code, name,
                    self.clean_num(qty), self.clean_num(price), 0, self.clean_num(amt)
                ))
                self.line_counter += 1

            self.renumber_rows()
            self.clear_inputs()

        except Exception as e:
            messagebox.showerror("录入错误", f"数值格式有误或缺失: {e}")

    def on_closing(self):
        items = self.tree.get_children()
        if items:
            state = {
                "mode": self.mode_var.get(),
                "f1": self.ent_f1.get(),
                "period": self.ent_period.get(),
                "sum_date": self.ent_sum_date.get(),
                "maker": self.ent_maker.get(),
                "set": self.ent_set.get(),
                "items": [self.tree.item(i, "values") for i in items]
            }
            try:
                with open(self.state_file, 'w', encoding='utf-8') as f:
                    json.dump(state, f, ensure_ascii=False)
            except:
                pass
        else:
            if os.path.exists(self.state_file):
                try:
                    os.remove(self.state_file)
                except:
                    pass
        self.root.destroy()

    def check_restore_state(self):
        if not os.path.exists(self.state_file): return
        try:
            with open(self.state_file, 'r', encoding='utf-8') as f:
                state = json.load(f)
            if state.get("items"):
                if messagebox.askyesno("恢复提示", "发现上次退出时有未保存的凭证内容，是否继续生成？"):
                    self.mode_var.set(state.get("mode", "transfer"))
                    self.ent_f1.delete(0, tk.END);
                    self.ent_f1.insert(0, state.get("f1", ""))
                    self.ent_period.delete(0, tk.END);
                    self.ent_period.insert(0, state.get("period", ""))
                    self.ent_sum_date.delete(0, tk.END);
                    self.ent_sum_date.insert(0, state.get("sum_date", ""))
                    for val in state.get("items", []):
                        self.tree.insert("", "end", values=val)
                    self.renumber_rows()
                else:
                    os.remove(self.state_file)
        except:
            pass

    def export_txt(self):
        items = self.tree.get_children()
        if not items: return False

        f1_val = self.ent_f1.get().strip()
        try:
            dt_obj = datetime.strptime(f1_val, "%Y-%m-%d")
            vid = dt_obj.strftime("%Y%m%d")
            year_str = str(dt_obj.year)
            month_str = str(dt_obj.month)
        except:
            vid = datetime.now().strftime("%Y%m%d")
            year_str = datetime.now().strftime("%Y")
            month_str = str(datetime.now().month)

        mode = self.mode_var.get()
        mode_name = "半成品成本结转凭证" if mode == "transfer" else "半成品使用凭证"
        filename = f"{vid}{mode_name}.txt"

        template = '"{f1}","记","9001","0","{summary}","{code}",{debit},{credit},{qty},0,.00,"{maker}",,,,,,,,,,,,,,,,,,,,,,,,,,,1,1,1,1,1,,1,1,1,0,,"{acc_set}","北农（海利）涿州种衣剂有限公司",{year},{month},1,,,,0,,{line},,0,,,,,,,,,,,,,,0,0,0,,,0,0,,,, , ,,,,,,0,.00'
        header_template = '"凭证输出","V800","001","北农（海利）涿州种衣剂有限公司","{year}","F1日期F2类别","F3凭证号F4附单据数","F5摘要F6科目编码","F7借方F8贷方","F9数量F10外币","F11汇率","F12制单人","F13结算方式","F14票号","F15发生日期","F16部门编码","F17个人编码","F18客户编码","F19供应商编码","F20业务员","F21项目编码","F22自定义项1","F23自定义项2","F24自由项1","F25自由项2","F26自由项3","F27自由项4","F28自由项5","F29自由项6","F30自由项7","F31自由项8","F32自由项9","F33自由项10","F34外部系统标识","F35业务类型","F36单据类型","F37单据日期","F38单据号","F39凭证是否可改","F40分录是否可增删","F41合计金额是否保值","F42数值是否可改", "F43科目是否可改","F44受控科目","F45往来是否可改","F46部门是否可改","F47项目是否可改","F48往来项是否必输","F49账套号","F50核算单位","F51会计年度","F52会计期间","F53类别顺序号","F54凭证号","F55审核人","F56记账人","F57是否记账","F58出纳人","F59行号","F60外币名称","F61单价","F62科目名称","F63部门名称","F64个人名称","F65客户简称","F66供应商简称","F67项目名称","F68项目大类编码","F69项目大类名称","F70对方科目","F71银行两清标志","F72往来两清标志","F73银行核销标志","F74外部系统名称","F75外部账套号","F76外部会计年度","F77外部会计期间","F78外部制单日期","F79外部系统版本","F80凭证标识","F81分录自动编号","F82唯一标识"\n'

        try:
            with open(filename, "w", encoding="gbk") as f:
                f.write(header_template.format(year=year_str))
                for tid in items:
                    v = self.tree.item(tid)["values"]
                    # 导出索引更新: 4:qty, 6:db, 7:cr
                    line_str = template.format(
                        f1=f1_val, summary=v[1], code=v[2],
                        debit=v[6], credit=v[7], qty=v[4],
                        maker=self.ent_maker.get().strip(),
                        acc_set=self.ent_set.get().strip(),
                        year=year_str, month=month_str, line=v[0]
                    )
                    f.write(line_str + "\n")
            messagebox.showinfo("成功", f"文件已生成：\n{filename}")
            return True
        except Exception as e:
            messagebox.showerror("导出错误", str(e))
            return False

    def reset_voucher(self):
        if self.tree.get_children():
            ans = messagebox.askyesnocancel("保存提示", "当前凭证尚未清空，是否先将其【保存(导出)】再生成下一份？")
            if ans is None: return
            if ans is True:
                if not self.export_txt(): return
        for i in self.tree.get_children(): self.tree.delete(i)
        self.line_counter = 1
        self.clear_inputs()
        self.ent_set.delete(0, tk.END)
        self.ent_set.insert(0, "001")

    def handle_drop(self, event):
        file_path = event.data.strip('{}')
        self.verify_txt_file(file_path)

    def verify_txt_file(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("凭证TXT", "*.txt")])
        if not file_path: return

        if self.tree.get_children():
            if not messagebox.askyesno("覆盖警告", "解析新的TXT将清空当前表格中的未保存内容，确定要继续吗？"):
                return

        try:
            with open(file_path, 'r', encoding='gbk') as f:
                reader = csv.reader(f)
                lines = list(reader)
            if not lines or len(lines) < 2:
                raise ValueError("文件内容为空或格式不符")

            for i in self.tree.get_children(): self.tree.delete(i)
            self.line_counter = 1

            is_usage = any("5001019907" in row[5] for row in lines[1:] if len(row) > 10)
            self.mode_var.set("usage" if is_usage else "transfer")

            first_data = lines[1]
            if len(first_data) >= 49:
                f1_val = first_data[0]
                self.ent_f1.delete(0, tk.END);
                self.ent_f1.insert(0, f1_val)
                self.ent_maker.delete(0, tk.END);
                self.ent_maker.insert(0, first_data[11])
                self.ent_set.delete(0, tk.END);
                self.ent_set.insert(0, first_data[48])
                self.auto_format_f1()

                match = re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", first_data[4])
                if match:
                    self.ent_sum_date.delete(0, tk.END);
                    self.ent_sum_date.insert(0, match.group(1))

            for row in lines[1:]:
                if len(row) > 10:
                    code, summary = row[5], row[4]
                    db, cr, qty = float(row[6]), float(row[7]), float(row[8])
                    total = db if db > 0 else cr
                    price = round(total / qty, 4) if qty else 0

                    name_found = "未知科目(请检查字典)"
                    if is_usage and code == "5001019907":
                        name_found = "半产品成本"
                    elif not is_usage and code.startswith("500101") and code.endswith("998"):
                        name_found = "自制半成品成本"
                    else:
                        for k, v in self.subject_dict.items():
                            if v == code: name_found = k; break

                    # 还原索引更新: 4:qty, 5:price, 6:db, 7:cr
                    self.tree.insert("", "end", values=(
                        self.line_counter, summary, code, name_found,
                        self.clean_num(qty) if qty else "0",
                        self.clean_num(price) if price else "-",
                        self.clean_num(db), self.clean_num(cr)
                    ))
                    self.line_counter += 1

            messagebox.showinfo("解析成功", "凭证已加载到原录入区，现在您可以像平时一样去修改价格、增减分录并重新导出了。")

        except Exception as e:
            messagebox.showerror("读取错误", f"无法解析该凭证文件: {e}")

    def schedule_search(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"): return
        if self._search_timer: self.root.after_cancel(self._search_timer)
        self._search_timer = self.root.after(300, self.perform_search)

    def perform_search(self):
        val = self.name_var.get().strip()
        if not val:
            self.suggestion_list.place_forget();
            return
        matches = [k for k in self.subject_dict.keys() if val.lower() in k.lower()]
        if matches:
            self.suggestion_list.delete(0, tk.END)
            for m in matches: self.suggestion_list.insert(tk.END, m)
            x = self.name_in.winfo_rootx() - self.root.winfo_rootx()
            y = self.name_in.winfo_rooty() - self.root.winfo_rooty() + self.name_in.winfo_height()
            self.suggestion_list.place(x=x, y=y, width=self.name_in.winfo_width())
            self.suggestion_list.lift()
        else:
            self.suggestion_list.place_forget()

    def on_select_suggestion(self, event):
        if not self.suggestion_list.curselection(): return
        name = self.suggestion_list.get(self.suggestion_list.curselection())
        self.name_var.set(name)
        self.code_in.delete(0, tk.END);
        self.code_in.insert(0, self.subject_dict.get(name, ""))
        self.suggestion_list.place_forget()
        self.qty_in.focus()

    def clean_num(self, val):
        v = float(val)
        return str(int(v)) if v == int(v) else str(v)

    def clear_inputs(self):
        self.name_var.set("")
        self.code_in.delete(0, tk.END)
        self.qty_in.delete(0, tk.END)
        self.price_in.delete(0, tk.END)
        self.suggestion_list.place_forget()
        self.name_in.focus()

    def edit_selected_entry(self):
        sel = self.tree.selection()
        if not sel: return
        item = sel[0]
        vals = self.tree.item(item, "values")

        # 回填索引更新: 2:code, 6:db, 7:cr, 4:qty
        code, db, cr, qty = vals[2], float(vals[6]), float(vals[7]), float(vals[4])
        amt = db if db > 0 else cr
        price = round(amt / qty, 4) if qty else 0

        self.code_in.delete(0, tk.END);
        self.code_in.insert(0, code)
        self.qty_in.delete(0, tk.END);
        self.qty_in.insert(0, self.clean_num(qty))
        self.price_in.delete(0, tk.END);
        self.price_in.insert(0, self.clean_num(price))

        fname = ""
        for k, v in self.subject_dict.items():
            if v == code: fname = k; break
        self.name_var.set(fname)

        all_items = self.tree.get_children()
        idx = all_items.index(item)
        to_del = [item]
        pair_idx = idx + 1 if idx % 2 == 0 else idx - 1
        if 0 <= pair_idx < len(all_items):
            pair_item = all_items[pair_idx]
            p_vals = self.tree.item(pair_item, "values")
            # 匹配对应行索引也是6/7
            if float(p_vals[6]) == amt or float(p_vals[7]) == amt:
                to_del.append(pair_item)

        for i in to_del: self.tree.delete(i)
        self.renumber_rows()
        messagebox.showinfo("修改", "数据已回填，请修改后重新点击添加。\n注意：请确保当前【业务模式】选择正确！")

    def renumber_rows(self):
        for index, item_id in enumerate(self.tree.get_children()):
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = index + 1
            self.tree.item(item_id, values=vals)
        self.line_counter = len(self.tree.get_children()) + 1

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            if len(self.tree.selection()) == 0: self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def modify_summary_text_only(self):
        sel = self.tree.selection()
        if not sel: return
        item = sel[0]
        vals = list(self.tree.item(item, "values"))
        d = tk.Toplevel(self.root)
        d.title("修改摘要");
        d.geometry("500x200")
        tk.Label(d, text="编辑摘要:").pack(pady=5)
        t = tk.Text(d, height=5, width=50);
        t.pack(pady=5)
        t.insert("1.0", vals[1])

        def save():
            vals[1] = t.get("1.0", "end-1c").strip().replace("\n", "")
            self.tree.item(item, values=vals)
            d.destroy()

        tk.Button(d, text="确认", command=save).pack()

    def auto_format_f1(self, e=None):
        val = self.ent_f1.get().strip()
        if len(val) == 8 and val.isdigit():
            dt = datetime.strptime(val, "%Y%m%d")
            self.ent_f1.delete(0, tk.END);
            self.ent_f1.insert(0, f"{dt.year}-{dt.month}-{dt.day}")
            self.ent_period.delete(0, tk.END);
            self.ent_period.insert(0, str(dt.month))
            self.ent_sum_date.delete(0, tk.END);
            self.ent_sum_date.insert(0, (dt - timedelta(days=1)).strftime("%Y%m%d"))
            self.auto_format_sum_date()

    def auto_format_sum_date(self, e=None):
        val = self.ent_sum_date.get().strip()
        if len(val) == 8 and val.isdigit():
            dt = datetime.strptime(val, "%Y%m%d")
            self.ent_sum_date.delete(0, tk.END);
            self.ent_sum_date.insert(0, dt.strftime("%Y年%m月%d日"))

    def export_excel_dict(self):
        try:
            pd.DataFrame(list(self.subject_dict.items()), columns=["科目名称", "科目编码"]).to_excel(
                f"字典备份_{datetime.now().strftime('%H%M%S')}.xlsx", index=False)
            messagebox.showinfo("OK", "已导出")
        except Exception as e:
            messagebox.showerror("Err", str(e))

    def import_excel_dict(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not fp: return
        try:
            df = pd.read_excel(fp)
            df.drop_duplicates(subset=["科目名称"], keep="last", inplace=True)

            new_dict = {}
            for _, row in df.iterrows():
                new_name = str(row["科目名称"]).strip()
                new_code = str(row["科目编码"]).strip()

                old_code = self.subject_dict.get(new_name)
                old_name = next((k for k, v in self.subject_dict.items() if v == new_code), None)

                if old_code and old_code != new_code:
                    if messagebox.askyesno("编码冲突",
                                           f"科目【{new_name}】的编码已从 {old_code} 变更为 {new_code}。\n是否更新？(选否则保留旧编码)"):
                        new_dict[new_name] = new_code
                    else:
                        new_dict[new_name] = old_code
                elif old_name and old_name != new_name:
                    if messagebox.askyesno("名称冲突",
                                           f"编码【{new_code}】对应的科目已从 {old_name} 变更为 {new_name}。\n是否更新？(选否则保留原科目名映射)"):
                        new_dict[new_name] = new_code
                    else:
                        new_dict[old_name] = new_code
                else:
                    new_dict[new_name] = new_code

            self.subject_dict = new_dict
            self.save_dict()
            messagebox.showinfo("完成", "字典导入并更新完成！未包含在 Excel 中的旧映射已被清理。")

        except Exception as e:
            messagebox.showerror("导入错误", str(e))


if __name__ == "__main__":
    root = BaseTk()
    app = VoucherFStringApp(root)
    root.mainloop()