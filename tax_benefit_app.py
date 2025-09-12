import tkinter as tk
from tkinter import ttk, messagebox
import re
import sys
from tkinter import filedialog
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


class TaxBenefitApp:

    def __init__(self, root):
        """初始化应用程序"""
        self.root = root
        self.root.title("注销风险监测器")

        # 窗口大小
        if sys.platform == "darwin":  # macOS
            self.root.geometry("900x700+100+100")
        else:
            self.root.geometry("900x700+100+100")

        # 强制浅色
        self.root.configure(bg='white')
        self.root.minsize(900, 600)

        # ====== 顶层：用一个容器 + grid，保证底部按钮永远可见 ======
        self.root_container = tk.Frame(self.root, bg='white')
        self.root_container.pack(fill=tk.BOTH, expand=True)

        # 两行：0=内容可伸缩，1=固定底部按钮
        self.root_container.grid_rowconfigure(0, weight=1)
        self.root_container.grid_rowconfigure(1, weight=0)
        self.root_container.grid_columnconfigure(0, weight=1)

        # 内容区（用 pack 布局，不与 root 混用）
        self.main_frame = tk.Frame(self.root_container, bg='white', padx=20, pady=20, relief=tk.RAISED, bd=2)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 0))

        # 固定底部按钮区（单独一行）
        self.footer_frame = tk.Frame(self.root_container, bg='white', padx=10, pady=10, relief=tk.GROOVE, bd=1)
        self.footer_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        # 数据
        self.tax_items = ["制造业缓税", "交通运输减免", "小微企业减免", "高新技术企业减免", "环保设备减免", "研发费用加计扣除"]
        self.entries = {}
        self.rule_inputs = {}

        # 界面
        self.create_widgets()
        self.create_buttons_section()  # 渲染在 footer_frame

        # 置顶闪现
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(self.root.attributes, '-topmost', False)

    def create_widgets(self):
        """创建所有界面组件（标题、企业信息、规则区、表格区）"""
        title_label = tk.Label(self.main_frame, text="注销风险监测器",
                               font=('Arial', 18, 'bold'), bg='white', fg='#2c3e50')
        title_label.pack(pady=(0, 20))

        self.create_header_section()
        self.create_rules_section()

        separator_frame = tk.Frame(self.main_frame, height=2, bg='#bdc3c7')
        separator_frame.pack(fill=tk.X, pady=20)

        self.create_tax_items_section()

    # ---------- 顶部企业信息 ----------
    def create_header_section(self):
        header_frame = tk.Frame(self.main_frame, bg='white', relief=tk.GROOVE, bd=2)
        header_frame.pack(fill=tk.X, pady=(0, 10), padx=5)

        code_frame = tk.Frame(header_frame, bg='white')
        code_frame.pack(fill=tk.X, pady=10, padx=10)

        tk.Label(code_frame, text="社会信用代码:", font=('Arial', 14, 'bold'),
                 bg='black', fg='white', width=15, anchor='w').pack(side=tk.LEFT)
        self.credit_code_entry = tk.Entry(code_frame, font=('Arial', 12), width=40,
                                          relief=tk.SUNKEN, bd=2, fg='black', bg='white', insertbackground='black')
        self.credit_code_entry.pack(side=tk.LEFT, padx=(10, 0))

        name_frame = tk.Frame(header_frame, bg='white')
        name_frame.pack(fill=tk.X, pady=10, padx=10)

        tk.Label(name_frame, text="企业名称:", font=('Arial', 14, 'bold'),
                 bg='black', fg='white', width=15, anchor='w').pack(side=tk.LEFT)
        self.company_name_entry = tk.Entry(name_frame, font=('Arial', 12), width=40,
                                           relief=tk.SUNKEN, bd=2, fg='black', bg='white', insertbackground='black')
        self.company_name_entry.pack(side=tk.LEFT, padx=(10, 0))

    # ---------- 疑点规则输入 ----------
    def create_rules_section(self):
        block = tk.LabelFrame(self.main_frame, text="疑点规则检查（Excel 同款）",
                              bg='white', fg='#2c3e50', font=('Arial', 12, 'bold'))
        block.pack(fill=tk.X, padx=5, pady=10)

        # ttk.Combobox 强制黑字白底（macOS 深色模式兼容）
        style = ttk.Style()
        style.configure("White.TCombobox", foreground='black', fieldbackground='white', background='white')

        fields = [
            {"name": "主营行业 (A)", "key": "A_主营行业", "type": "select",
             "choices": ["批发零售", "制造", "建筑安装", "交通运输", "其他"]},
            {"name": "年度 (B)", "key": "B_年度", "type": "number"},
            {"name": "营业收入 (C)", "key": "C_营业收入", "type": "number"},
            {"name": "销售收入 (D)", "key": "D_销售收入", "type": "number"},
            {"name": "成本 (E)", "key": "E_成本", "type": "number"},
            {"name": "销售费用 (F)", "key": "F_销售费用", "type": "number"},
            {"name": "管理费用 (G)", "key": "G_管理费用", "type": "number"},
            {"name": "财务费用 (H)", "key": "H_财务费用", "type": "number"},
            {"name": "工资薪金 (I)", "key": "I_工资薪金", "type": "number"},
            {"name": "个税扣缴工资总额 (J)", "key": "J_个税扣缴工资总额", "type": "number"},
            {"name": "计提折旧 (K)", "key": "K_计提折旧", "type": "number"},
            {"name": "当期开票额度 (L)", "key": "L_当期开票额度", "type": "number"},
            {"name": "当期受票额度 (M)", "key": "M_当期受票额度", "type": "number"},
            {"name": "印花税计税依据 (N)", "key": "N_印花税计税依据", "type": "number"},
        ]

        left = tk.Frame(block, bg='white'); right = tk.Frame(block, bg='white')
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        half = (len(fields) + 1) // 2
        cols = (fields[:half], fields[half:])
        self.rule_inputs.clear()

        for col_fields, parent in zip(cols, (left, right)):
            for f in col_fields:
                row = tk.Frame(parent, bg='white'); row.pack(fill=tk.X, pady=4)
                tk.Label(row, text=f["name"], bg='white', fg='#2c3e50',
                         font=('Arial', 11), width=22, anchor='w').pack(side=tk.LEFT)

                if f["type"] == "select":
                    cb = ttk.Combobox(row, values=f["choices"], state="readonly",
                                      width=22, font=('Arial', 11), style="White.TCombobox")
                    cb.pack(side=tk.LEFT, padx=6)
                    cb.set(f["choices"][0])
                    self.rule_inputs[f["key"]] = cb
                else:
                    ent = tk.Entry(row, font=('Arial', 11), width=25,
                                   relief=tk.SUNKEN, bd=2, justify='center',
                                   fg='black', bg='white', insertbackground='black')
                    ent.pack(side=tk.LEFT, padx=6)
                    ent.bind('<KeyRelease>', self._rule_number_hint)
                    self.rule_inputs[f["key"]] = ent

    def _rule_number_hint(self, event):
        w = event.widget
        val = w.get().strip()
        if not val:
            w.config(bg='white', fg='black', insertbackground='black'); return
        if re.match(r'^\d+(\.\d{1,2})?$', val):
            w.config(bg='white', fg='black', insertbackground='black')
        else:
            w.config(bg='#ffecec', fg='black', insertbackground='black')

    def _get_dec(self, key) -> Decimal:
        widget = self.rule_inputs[key]
        if isinstance(widget, ttk.Combobox):
            raise ValueError("请求的字段是下拉框（非数字）")
        txt = widget.get().strip()
        if not txt:
            return Decimal("0")
        if not re.match(r'^\d+(\.\d{1,2})?$', txt):
            raise ValueError(f"{key} 数字格式不正确")
        try:
            return Decimal(txt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            raise ValueError(f"{key} 数字解析失败")

    # ---------- 税收项目表 ----------
    def create_tax_items_section(self):
        table_frame = tk.Frame(self.main_frame, bg='white', height=400, relief=tk.GROOVE, bd=2)
        table_frame.pack(fill=tk.X, pady=10, padx=5)
        table_frame.pack_propagate(False)

        header_frame = tk.Frame(table_frame, bg='#e8f4fd', relief=tk.RAISED, bd=2)
        header_frame.pack(fill=tk.X, pady=(5, 5), padx=5)

        headers = ["税收项目", "应享优惠(元)", "已享优惠(元)", "未享优惠(元)"]
        for i, header in enumerate(headers):
            label = tk.Label(header_frame, text=header, font=('Arial', 12, 'bold'),
                             bg='#e8f4fd', fg='#2c3e50', relief=tk.RIDGE, bd=1, padx=10, pady=8)
            label.grid(row=0, column=0+i, sticky='ew', padx=1, pady=1)
            header_frame.grid_columnconfigure(i, weight=1)

        data_frame = tk.Frame(table_frame, bg='white')
        data_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for tax_item in self.tax_items:
            row_frame = tk.Frame(data_frame, bg='#f8f9fa', relief=tk.RIDGE, bd=1)
            row_frame.pack(fill=tk.X, pady=2)
            for j in range(4):
                row_frame.grid_columnconfigure(j, weight=1)

            tk.Label(row_frame, text=tax_item, font=('Arial', 11),
                     bg='#f8f9fa', fg='#2c3e50', relief=tk.FLAT, padx=8, pady=8, anchor='w')\
                .grid(row=0, column=0, sticky='ew', padx=2, pady=2)

            should_entry = tk.Entry(row_frame, font=('Arial', 11), relief=tk.SUNKEN, bd=2,
                                    justify='center', bg='white', fg='black', insertbackground='black')
            should_entry.grid(row=0, column=1, sticky='ew', padx=2, pady=4)
            should_entry.bind('<KeyRelease>', self.validate_number_input)

            enjoyed_entry = tk.Entry(row_frame, font=('Arial', 11), relief=tk.SUNKEN, bd=2,
                                     justify='center', bg='white', fg='black', insertbackground='black')
            enjoyed_entry.grid(row=0, column=2, sticky='ew', padx=2, pady=4)
            enjoyed_entry.bind('<KeyRelease>', self.validate_number_input)

            not_enjoyed_entry = tk.Entry(row_frame, font=('Arial', 11), state='readonly',
                                         relief=tk.SUNKEN, bd=2, readonlybackground='#fff3cd',
                                         justify='center', fg='#856404')
            not_enjoyed_entry.grid(row=0, column=3, sticky='ew', padx=2, pady=4)

            self.entries[tax_item] = {'should': should_entry, 'enjoyed': enjoyed_entry, 'not_enjoyed': not_enjoyed_entry}

    # ---------- 底部固定按钮 ----------
    def create_buttons_section(self):
        button_frame = tk.Frame(self.footer_frame, bg='white')
        button_frame.pack(side=tk.TOP)

        button_config = {'font': ('Arial', 13, 'bold'), 'width': 12, 'height': 2,
                         'relief': tk.RAISED, 'bd': 3, 'cursor': 'hand2'}

        query_btn = tk.Button(button_frame, text="查询", command=self.calculate_benefits,
                              bg='#007bff', fg='black', activebackground='#0056b3',
                              activeforeground='white', **button_config)
        query_btn.pack(side=tk.LEFT, padx=12)

        check_btn = tk.Button(button_frame, text="规则检查", command=self.run_rule_checks,
                              bg='#17a2b8', fg='black', activebackground='#117a8b',
                              activeforeground='white', **button_config)
        check_btn.pack(side=tk.LEFT, padx=12)

        reset_btn = tk.Button(button_frame, text="重置", command=self.reset_form,
                              bg='#dc3545', fg='black', activebackground='#c82333',
                              activeforeground='white', **button_config)
        reset_btn.pack(side=tk.LEFT, padx=12)

        exit_btn = tk.Button(button_frame, text="退出", command=self.exit_app,
                             bg='#6c757d', fg='black', activebackground='#5a6268',
                             activeforeground='white', **button_config)
        exit_btn.pack(side=tk.LEFT, padx=12)

    # ---------- 校验/计算 ----------
    def validate_number_input(self, event):
        entry = event.widget
        value = entry.get()
        if value and not re.match(r'^\d*\.?\d*$', value):
            entry.delete(len(value) - 1, tk.END)

    def _set_readonly_value(self, entry, value, bg=None, fg=None):
        entry.config(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, value)
        if bg: entry.config(readonlybackground=bg)
        if fg: entry.config(fg=fg)
        entry.config(state='readonly')

    def calculate_benefits(self):
        try:
            if not self.credit_code_entry.get().strip() or not self.company_name_entry.get().strip():
                messagebox.showwarning("警告", "请填写社会信用代码和企业名称！")
                return

            total_not_enjoyed = 0.0
            for tax_item in self.tax_items:
                refs = self.entries[tax_item]
                should_value = refs['should'].get().strip()
                enjoyed_value = refs['enjoyed'].get().strip()
                should_amount = float(should_value) if should_value else 0.0
                enjoyed_amount = float(enjoyed_value) if enjoyed_value else 0.0
                not_enjoyed_amount = max(0.0, should_amount - enjoyed_amount)
                total_not_enjoyed += not_enjoyed_amount

                self._set_readonly_value(
                    refs['not_enjoyed'],
                    f"{not_enjoyed_amount:.2f}",
                    bg='#ffebee' if not_enjoyed_amount > 0 else '#e8f5e8',
                    fg='#d32f2f' if not_enjoyed_amount > 0 else '#2e7d32'
                )

            messagebox.showinfo("计算完成", f"计算完成！\n总未享优惠金额：{total_not_enjoyed:.2f}元")

        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
        except Exception as e:
            messagebox.showerror("错误", f"计算过程中发生错误：{e}")

    # ---------- Excel 等价规则检查 ----------
    def run_rule_checks(self):
        try:
            A = self.rule_inputs["A_主营行业"].get().strip()
            C = self._get_dec("C_营业收入")
            D = self._get_dec("D_销售收入")
            E = self._get_dec("E_成本")
            F = self._get_dec("F_销售费用")
            G = self._get_dec("G_管理费用")
            H = self._get_dec("H_财务费用")
            I = self._get_dec("I_工资薪金")
            J = self._get_dec("J_个税扣缴工资总额")
            K = self._get_dec("K_计提折旧")
            L = self._get_dec("L_当期开票额度")
            M = self._get_dec("M_当期受票额度")
            N = self._get_dec("N_印花税计税依据")  # 当前 5 条规则未直接用

            issues = []

            # 1) |D - C| > 100
            diff = (D - C)
            if diff > Decimal("100") or diff < Decimal("-100"):
                issues.append("同期申报的增值税收入与企业所得税收入有差异")

            # 2) 成本/费用占比异常
            if C > 0 and A in {"批发零售", "制造"}:
                total_ratio = (E + F + G + H) / C
                if total_ratio >= Decimal("0.70"):
                    if E / C <= Decimal("0.50"):
                        issues.append("工商业纳税人成本费用占比异常")
                    if (F + G + H) / C >= Decimal("0.50"):
                        issues.append("工商业纳税人费用占比异常")

            # 3) E+F+G+H - I - K - M > 20000
            if (E + F + G + H - I - K - M) > Decimal("20000"):
                issues.append("疑似未取得合法有效凭证列支")

            # 4) I≥500000 且 I-J≥100000
            if I >= Decimal("500000") and (I - J) >= Decimal("100000"):
                issues.append("疑似多列工资薪金支出或少扣缴个人所得税")

            # 5) 行业在集合 且 (C+E+F+G - I)≥1,000,000
            base = (C + E + F + G - I)
            if A in {"批发零售", "制造", "建筑安装", "交通运输", "其他"} and base >= Decimal("1000000"):
                issues.append("疑似少计印花税计税依据")

            if issues:
                msg = "发现以下疑点：\n\n" + "\n".join(f"• {s}" for s in issues)
                messagebox.showwarning("规则检查结果", msg)
            else:
                messagebox.showinfo("规则检查结果", "未发现疑点。")

        except ValueError as e:
            messagebox.showerror("输入错误", f"请检查规则区输入：{e}")
        except Exception as e:
            messagebox.showerror("错误", f"规则检查过程发生异常：{e}")

    # ---------- 重置/退出 ----------
    def reset_form(self):
        if messagebox.askyesno("确认重置", "确定要清空所有输入吗？"):
            self.credit_code_entry.delete(0, tk.END)
            self.company_name_entry.delete(0, tk.END)

            # 规则区
            for key, widget in self.rule_inputs.items():
                if isinstance(widget, ttk.Combobox):
                    widget.set(widget['values'][0])
                else:
                    widget.delete(0, tk.END)
                    widget.config(bg='white', fg='black', insertbackground='black')

            # 税收项目区
            for tax_item in self.tax_items:
                refs = self.entries[tax_item]
                refs['should'].delete(0, tk.END)
                refs['enjoyed'].delete(0, tk.END)
                refs['not_enjoyed'].config(state='normal')
                refs['not_enjoyed'].delete(0, tk.END)
                refs['not_enjoyed'].config(state='readonly', readonlybackground='#f8f9fa', fg='black')

            messagebox.showinfo("重置完成", "所有输入已清空！")

    def exit_app(self):
        if messagebox.askyesno("确认退出", "确定要退出程序吗？"):
            self.root.quit()
            self.root.destroy()


def main():
    try:
        root = tk.Tk()
        app = TaxBenefitApp(root)
        root.protocol("WM_DELETE_WINDOW", app.exit_app)
        root.mainloop()
    except Exception as e:
        print(f"程序启动失败：{str(e)}")
        if 'messagebox' in globals():
            messagebox.showerror("启动错误", f"程序启动失败：{e}")


if __name__ == "__main__":
    main()
