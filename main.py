"""
久益检验报告更新器
==================
根据发货计划自动更新检验报告的出货数量和检验日期，并支持一键打印。

工作流:
1. 读取发货计划 (.xlsx)，提取 料件编号 + 本周要求数量
2. 在检验报告目录中按料件编号匹配对应的 .xls 报告文件
3. 通过 VBScript + COM 自动化操作 WPS/Excel 修改报告
4. 一键打印所有已更新的报告

目标平台: Windows (需安装 WPS 或 Microsoft Excel)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import subprocess
import tempfile
import threading
from collections import OrderedDict

import openpyxl


# ---------------------------------------------------------------------------
# 发货计划读取
# ---------------------------------------------------------------------------

def read_shipping_plan(filepath):
    """读取送货单，提取 (料件编号, 数量, 行号)

    新送货单格式：列B=产品料号，列E=数量，第8行起为数据行。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    items = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=8, values_only=True), start=8):
        if len(row) < 5:
            continue
        part_no = row[1]   # 列 B: 产品料号
        quantity = row[4]  # 列 E: 数量
        # 跳过合计行和空行
        if not part_no or not str(part_no).strip():
            continue
        pn = str(part_no).strip()
        if pn in ('合计', '总计'):
            continue
        if quantity is None:
            continue
        try:
            qty = int(float(quantity))
        except (ValueError, TypeError):
            continue
        items.append({
            'row': row_idx,
            'part_no': pn,
            'quantity': qty,
        })
    wb.close()
    return items


def find_report_file(report_dir, part_no):
    """按文件名前缀匹配检验报告文件"""
    for fname in os.listdir(report_dir):
        if fname.upper().startswith(part_no.upper()) and (
            fname.lower().endswith('.xls') or fname.lower().endswith('.xlsx')
        ):
            return os.path.join(report_dir, fname)
    return None


def group_by_part_no(items):
    """按料件编号分组，保持顺序"""
    groups = OrderedDict()
    for item in items:
        pn = item['part_no']
        if pn not in groups:
            groups[pn] = []
        groups[pn].append(item)
    return groups


# ---------------------------------------------------------------------------
# VBScript 生成器
# ---------------------------------------------------------------------------

class VBScriptGenerator:
    """动态生成 VBScript 脚本，用于通过 COM 操作 WPS/Excel"""

    @staticmethod
    def _header():
        """VBS 脚本头部：连接 WPS 或 Excel"""
        return r'''On Error Resume Next

Dim app
Set app = CreateObject("Excel.Application")
If app Is Nothing Then
    Set app = CreateObject("KET.Application")
End If
If app Is Nothing Then
    Set app = CreateObject("ET.Application")
End If
If app Is Nothing Then
    WScript.Echo "FATAL:无法启动 WPS 或 Excel，请确认已安装"
    WScript.Quit 1
End If

On Error GoTo 0
app.Visible = False
app.DisplayAlerts = False

Dim wb, ws
'''

    @staticmethod
    def _footer():
        return '''
app.Quit
WScript.Echo "DONE"
'''

    @staticmethod
    def generate_update_script(tasks):
        """生成更新报告的 VBS 脚本

        tasks: [(part_no, filepath, final_quantity), ...]
        日期由报告内 =TODAY() 公式自动更新，无需手动写入。
        """
        lines = [VBScriptGenerator._header()]
        lines.append('')
        lines.append('On Error Resume Next')
        lines.append('')

        for i, (part_no, filepath, quantity) in enumerate(tasks):
            fp = filepath.replace('"', '""')
            lines.append(f'Set wb = app.Workbooks.Open("{fp}")')
            lines.append('If Err.Number <> 0 Then')
            lines.append(f'    WScript.Echo "ERROR:{part_no}:无法打开文件"')
            lines.append('    Err.Clear')
            lines.append('Else')
            lines.append('    Set ws = wb.Sheets(1)')
            lines.append('    Err.Clear')
            lines.append(f'    ws.Cells(2, 14).Value = {quantity}')
            lines.append('    If Err.Number <> 0 Then')
            lines.append(f'        WScript.Echo "ERROR:{part_no}:更新失败"')
            lines.append('        Err.Clear')
            lines.append('    Else')
            lines.append('        wb.Save')
            lines.append(f'        WScript.Echo "OK:{part_no}:{quantity}"')
            lines.append('    End If')
            lines.append('    wb.Close False')
            lines.append('End If')
            lines.append(f'WScript.Echo "PROGRESS:{i + 1}"')
            lines.append('')

        lines.append(VBScriptGenerator._footer())
        return '\n'.join(lines)

    @staticmethod
    def generate_export_pdf_script(tasks, pdf_dir):
        """生成将报告导出为 PDF 的 VBS 脚本

        tasks: [(part_no, filepath, [qty1, qty2, ...]), ...]
        pdf_dir: PDF 输出目录
        返回: (vbs_content, expected_pdf_files)
        expected_pdf_files 按导出顺序排列
        """
        lines = [VBScriptGenerator._header()]
        lines.append('')
        lines.append('On Error Resume Next')
        lines.append('')

        pdf_files = []  # 按导出顺序记录文件名
        export_idx = 0
        for part_no, filepath, quantities in tasks:
            fp = filepath.replace('"', '""')
            lines.append(f'Set wb = app.Workbooks.Open("{fp}")')
            lines.append('If Err.Number <> 0 Then')
            lines.append(f'    WScript.Echo "ERROR:{part_no}:无法打开文件"')
            lines.append('    Err.Clear')
            lines.append('Else')

            if len(quantities) == 1:
                pdf_name = f"{export_idx:04d}_{part_no}.pdf"
                pdf_path = os.path.join(pdf_dir, pdf_name).replace('"', '""')
                lines.append(f'    wb.ExportAsFixedFormat 0, "{pdf_path}"')
                lines.append('    If Err.Number <> 0 Then')
                lines.append(f'        WScript.Echo "EXPORT_ERROR:{part_no}:导出PDF失败"')
                lines.append('        Err.Clear')
                lines.append('    Else')
                lines.append(f'        WScript.Echo "EXPORTED:{part_no}:{quantities[0]}"')
                lines.append('    End If')
                lines.append('    wb.Close False')
                pdf_files.append(pdf_name)
                export_idx += 1
            else:
                lines.append('    Set ws = wb.Sheets(1)')
                for qi, qty in enumerate(quantities):
                    pdf_name = f"{export_idx:04d}_{part_no}_{qi}.pdf"
                    pdf_path = os.path.join(
                        pdf_dir, pdf_name).replace('"', '""')
                    lines.append(f'    ws.Cells(2, 14).Value = {qty}')
                    lines.append('    Err.Clear')
                    lines.append(f'    wb.ExportAsFixedFormat 0, "{pdf_path}"')
                    lines.append('    If Err.Number <> 0 Then')
                    lines.append(f'        WScript.Echo "EXPORT_ERROR:{part_no}:导出PDF失败(数量={qty})"')
                    lines.append('        Err.Clear')
                    lines.append('    Else')
                    lines.append(f'        WScript.Echo "EXPORTED:{part_no}:{qty}"')
                    lines.append('    End If')
                    pdf_files.append(pdf_name)
                    export_idx += 1
                lines.append('    wb.Save')
                lines.append('    wb.Close False')

            lines.append('End If')
            lines.append(f'WScript.Echo "PROGRESS:{export_idx}"')
            lines.append('')

        lines.append(VBScriptGenerator._footer())
        return '\n'.join(lines), pdf_files


# ---------------------------------------------------------------------------
# VBS 执行器
# ---------------------------------------------------------------------------

def run_vbs_script(vbs_content, on_line=None):
    """将 VBS 内容写入临时文件并通过 cscript 执行

    on_line: 回调函数，每读到一行 stdout 时调用
    返回: (success: bool, error_msg: str)
    """
    enc = 'mbcs' if sys.platform == 'win32' else 'utf-8'

    tmp = tempfile.NamedTemporaryFile(
        suffix='.vbs', delete=False, mode='w', encoding=enc
    )
    try:
        tmp.write(vbs_content)
        tmp.close()

        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE

        proc = subprocess.Popen(
            ['cscript', '//nologo', tmp.name],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding=enc,
            startupinfo=startupinfo,
        )

        for line in proc.stdout:
            line = line.strip()
            if line and on_line:
                on_line(line)

        proc.wait(timeout=600)

        if proc.returncode != 0:
            err = proc.stderr.read()
            return False, f"VBS 执行失败 (code {proc.returncode}): {err}"

        return True, ""

    except subprocess.TimeoutExpired:
        proc.kill()
        return False, "执行超时（10分钟）"
    except Exception as e:
        return False, str(e)
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# 重复料件对话框
# ---------------------------------------------------------------------------

class DuplicateDialog(tk.Toplevel):
    """处理重复料件编号的模态对话框"""

    def __init__(self, parent, part_no, items):
        super().__init__(parent)
        self.title(f"重复料件 - {part_no}")
        self.result = None
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        pad = {'padx': 15}
        total = sum(it['quantity'] for it in items)

        ttk.Label(
            self,
            text=f"料件编号 {part_no} 出现了 {len(items)} 次：",
            font=("Microsoft YaHei", 11, "bold"),
        ).pack(pady=(15, 8), **pad, anchor='w')

        for it in items:
            ttk.Label(
                self,
                text=f"    第 {it['row']} 行：数量 = {it['quantity']:,}",
                font=("Microsoft YaHei", 10),
            ).pack(padx=20, anchor='w', pady=1)

        ttk.Separator(self, orient='horizontal').pack(fill='x', padx=15, pady=12)

        ttk.Label(
            self, text="请选择处理方式：", font=("Microsoft YaHei", 10)
        ).pack(**pad, anchor='w')

        bf = ttk.Frame(self)
        bf.pack(pady=15)
        ttk.Button(bf, text=f"合并（总计 {total:,}）",
                   command=lambda: self._choose('merge')).pack(side='left', padx=8)
        ttk.Button(bf, text="分开处理",
                   command=lambda: self._choose('split')).pack(side='left', padx=8)
        ttk.Button(bf, text="跳过",
                   command=lambda: self._choose('skip')).pack(side='left', padx=8)

        self.protocol("WM_DELETE_WINDOW", lambda: self._choose('skip'))

        self.update_idletasks()
        w = self.winfo_reqwidth() + 40
        h = self.winfo_reqheight() + 20
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.wait_window()

    def _choose(self, choice):
        self.result = choice
        self.destroy()


# ---------------------------------------------------------------------------
# 状态常量
# ---------------------------------------------------------------------------

STATUS_PENDING = "待处理"
STATUS_UPDATING = "更新中"
STATUS_OK = "完成"
STATUS_FAIL = "失败"
STATUS_NO_FILE = "无报告"
STATUS_SKIP = "已跳过"
STATUS_EXPORTING = "导出中"
STATUS_EXPORTED = "已导出"
STATUS_EXPORT_FAIL = "导出失败"
STATUS_MERGING = "合并中"
STATUS_PRINTING = "打印中"
STATUS_PRINTED = "已打印"
STATUS_PRINT_FAIL = "打印失败"

TAG_COLORS = {
    STATUS_PENDING: "#999999",
    STATUS_UPDATING: "#2196F3",
    STATUS_OK: "#4CAF50",
    STATUS_FAIL: "#F44336",
    STATUS_NO_FILE: "#F44336",
    STATUS_SKIP: "#9E9E9E",
    STATUS_EXPORTING: "#2196F3",
    STATUS_EXPORTED: "#4CAF50",
    STATUS_EXPORT_FAIL: "#F44336",
    STATUS_MERGING: "#FF9800",
    STATUS_PRINTING: "#2196F3",
    STATUS_PRINTED: "#4CAF50",
    STATUS_PRINT_FAIL: "#F44336",
}


# ---------------------------------------------------------------------------
# 主应用 GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("久益 · 检验报告更新器")
        self.root.geometry("780x620")
        self.root.minsize(680, 520)

        if sys.platform == 'win32':
            try:
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        self.plan_paths = []  # 多个送货单路径
        self.print_tasks = []
        self.processing = False
        self.tree_items = {}  # part_no -> tree item id

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill='both', expand=True)

        # 标题
        ttk.Label(
            main, text="久益 · 检验报告更新器",
            font=("Microsoft YaHei", 18, "bold"),
        ).pack(pady=(0, 12))

        # 文件选择区
        sel = ttk.LabelFrame(main, text=" 文件选择 ", padding=10)
        sel.pack(fill='x', pady=(0, 8))

        r1 = ttk.Frame(sel)
        r1.pack(fill='x', pady=3)
        ttk.Label(r1, text="送货单：", width=10,
                  font=("Microsoft YaHei", 10)).pack(side='left')
        self.plan_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.plan_var,
                  font=("Microsoft YaHei", 9)).pack(
            side='left', fill='x', expand=True, padx=5)
        ttk.Button(r1, text="选择", command=self._pick_plan).pack(side='right')

        r2 = ttk.Frame(sel)
        r2.pack(fill='x', pady=3)
        ttk.Label(r2, text="报告目录：", width=10,
                  font=("Microsoft YaHei", 10)).pack(side='left')
        self.dir_var = tk.StringVar()
        ttk.Entry(r2, textvariable=self.dir_var,
                  font=("Microsoft YaHei", 9)).pack(
            side='left', fill='x', expand=True, padx=5)
        ttk.Button(r2, text="选择", command=self._pick_dir).pack(side='right')

        # 操作按钮
        bf = ttk.Frame(main)
        bf.pack(fill='x', pady=8)
        self.btn_update = ttk.Button(
            bf, text="  ① 更新报告  ", command=self._on_update)
        self.btn_update.pack(side='left', expand=True, fill='x',
                             padx=(0, 5), ipady=10)
        self.btn_print = ttk.Button(
            bf, text="  ② 导出并打印  ", command=self._on_print,
            state='disabled')
        self.btn_print.pack(side='right', expand=True, fill='x',
                            padx=(5, 0), ipady=10)

        # 进度条 + 状态摘要
        pf = ttk.Frame(main)
        pf.pack(fill='x', pady=(0, 3))
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(pf, variable=self.progress_var,
                        maximum=100).pack(fill='x')
        self.status_frame = tk.Frame(main)
        self.status_frame.pack(fill='x', anchor='w')

        # 列表看板
        lf = ttk.LabelFrame(main, text=" 处理列表 ", padding=5)
        lf.pack(fill='both', expand=True, pady=(5, 0))

        cols = ("part_no", "quantity", "prints", "status", "detail")
        self.tree = ttk.Treeview(
            lf, columns=cols, show='headings', selectmode='none')
        self.tree.heading("part_no", text="料件编号")
        self.tree.heading("quantity", text="数量")
        self.tree.heading("prints", text="打印份数")
        self.tree.heading("status", text="状态")
        self.tree.heading("detail", text="详情")

        self.tree.column("part_no", width=130, minwidth=100)
        self.tree.column("quantity", width=80, minwidth=60, anchor='center')
        self.tree.column("prints", width=70, minwidth=50, anchor='center')
        self.tree.column("status", width=80, minwidth=60, anchor='center')
        self.tree.column("detail", width=260, minwidth=120)

        vsb = ttk.Scrollbar(lf, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        # 配置标签颜色
        for status, color in TAG_COLORS.items():
            tag = f"tag_{status}"
            self.tree.tag_configure(tag, foreground=color)

    # ---- 列表操作 ----

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_items = {}

    def _add_item(self, part_no, quantity, status, detail="", prints=""):
        tag = f"tag_{status}"
        iid = self.tree.insert(
            '', 'end',
            values=(part_no, quantity, prints, status, detail),
            tags=(tag,))
        self.tree_items[part_no] = iid
        return iid

    def _update_item(self, part_no, status=None, detail=None,
                     quantity=None, prints=None):
        iid = self.tree_items.get(part_no)
        if not iid:
            return
        current = self.tree.item(iid, 'values')
        new_vals = list(current)
        if quantity is not None:
            new_vals[1] = quantity
        if prints is not None:
            new_vals[2] = prints
        if status is not None:
            new_vals[3] = status
        if detail is not None:
            new_vals[4] = detail
        tag = f"tag_{new_vals[3]}"
        self.tree.item(iid, values=new_vals, tags=(tag,))
        self.tree.see(iid)

    def _update_item_safe(self, part_no, **kwargs):
        """线程安全的列表更新"""
        self.root.after(0, lambda: self._update_item(part_no, **kwargs))

    def _set_status(self, segments):
        """设置状态栏，支持多段不同颜色
        segments: [(text, color), ...] 或 str
        """
        def _do():
            for w in self.status_frame.winfo_children():
                w.destroy()
            if isinstance(segments, str):
                tk.Label(self.status_frame, text=segments,
                         font=("Microsoft YaHei", 9)).pack(side='left')
            else:
                for text, color in segments:
                    tk.Label(self.status_frame, text=text, fg=color,
                             font=("Microsoft YaHei", 9)).pack(side='left')
        self.root.after(0, _do)

    def _set_progress(self, value, text=None):
        self.root.after(0, lambda: self.progress_var.set(value))
        if text:
            self._set_status(text)

    def _set_buttons(self, update=None, print_=None):
        if update is not None:
            self.root.after(0, lambda: self.btn_update.configure(
                state='normal' if update else 'disabled'))
        if print_ is not None:
            self.root.after(0, lambda: self.btn_print.configure(
                state='normal' if print_ else 'disabled'))

    # ---- 文件选择 ----

    def _pick_plan(self):
        paths = filedialog.askopenfilenames(
            title="选择送货单（可多选）",
            filetypes=[("Excel 文件", "*.xlsx;*.xls"), ("所有文件", "*.*")])
        if paths:
            self.plan_paths = list(paths)
            if len(paths) == 1:
                self.plan_var.set(paths[0])
            else:
                self.plan_var.set(f"已选择 {len(paths)} 个送货单")

    def _pick_dir(self):
        d = filedialog.askdirectory(title="选择检验报告目录")
        if d:
            self.dir_var.set(d)

    # ---- 更新报告 ----

    def _on_update(self):
        if not self.plan_paths:
            messagebox.showerror("错误", "请选择送货单文件")
            return
        for p in self.plan_paths:
            if not os.path.isfile(p):
                messagebox.showerror("错误", f"文件不存在：{p}")
                return
        report_dir = self.dir_var.get().strip()
        if not report_dir or not os.path.isdir(report_dir):
            messagebox.showerror("错误", "请选择有效的检验报告目录")
            return

        self._set_buttons(update=False, print_=False)
        self.print_tasks = []
        self._clear_tree()
        self.progress_var.set(0)

        threading.Thread(
            target=self._do_update, args=(self.plan_paths, report_dir),
            daemon=True
        ).start()

    def _do_update(self, plan_paths, report_dir):
        try:
            self._set_progress(0, "正在读取送货单...")
            items = []
            for p in plan_paths:
                items.extend(read_shipping_plan(p))
            groups = group_by_part_no(items)

            # 构建任务列表 + 填充列表看板
            update_tasks = []
            self.print_tasks = []
            update_part_nos = []
            no_file_count = 0

            for part_no, group_items in groups.items():
                report_file = find_report_file(report_dir, part_no)

                if not report_file:
                    no_file_count += 1
                    qty_str = ", ".join(str(i['quantity']) for i in group_items)
                    self.root.after(0, self._add_item, part_no, qty_str,
                                    STATUS_NO_FILE, "未找到报告文件", "-")
                    continue

                if len(group_items) == 1:
                    qty = group_items[0]['quantity']
                    self.root.after(0, self._add_item, part_no, str(qty),
                                    STATUS_PENDING, "", "1")
                    update_tasks.append((part_no, report_file, qty))
                    self.print_tasks.append((part_no, report_file, [qty]))
                    update_part_nos.append(part_no)
                else:
                    # 重复料件：在主线程显示对话框
                    result_holder = [None]
                    event = threading.Event()

                    def show_dlg(pn=part_no, gi=group_items,
                                 rh=result_holder, ev=event):
                        dlg = DuplicateDialog(self.root, pn, gi)
                        rh[0] = dlg.result
                        ev.set()

                    self.root.after(0, show_dlg)
                    event.wait()

                    choice = result_holder[0]
                    if choice == 'merge':
                        total = sum(i['quantity'] for i in group_items)
                        self.root.after(0, self._add_item, part_no,
                                        str(total), STATUS_PENDING,
                                        f"合并 {len(group_items)} 条", "1")
                        update_tasks.append((part_no, report_file, total))
                        self.print_tasks.append(
                            (part_no, report_file, [total]))
                        update_part_nos.append(part_no)
                    elif choice == 'split':
                        qtys = [i['quantity'] for i in group_items]
                        qty_str = "+".join(str(q) for q in qtys)
                        self.root.after(0, self._add_item, part_no,
                                        qty_str, STATUS_PENDING,
                                        f"分开 {len(qtys)} 份",
                                        str(len(qtys)))
                        update_tasks.append(
                            (part_no, report_file, qtys[-1]))
                        self.print_tasks.append(
                            (part_no, report_file, qtys))
                        update_part_nos.append(part_no)
                    else:
                        qty_str = ", ".join(
                            str(i['quantity']) for i in group_items)
                        self.root.after(0, self._add_item, part_no,
                                        qty_str, STATUS_SKIP, "用户跳过",
                                        "-")
                        continue

            if not update_tasks:
                self._set_progress(100, "没有需要处理的报告")
                self._set_buttons(update=True)
                return

            # 标记所有待处理项为"更新中"
            import time
            time.sleep(0.1)  # let GUI catch up
            for pn in update_part_nos:
                self._update_item_safe(pn, status=STATUS_UPDATING)

            self._set_progress(0, f"准备更新 {len(update_tasks)} 个报告...")
            vbs = VBScriptGenerator.generate_update_script(update_tasks)
            total = len(update_tasks)
            ok_count = 0
            fail_count = 0

            def on_vbs_line(line):
                nonlocal ok_count, fail_count
                if line.startswith("OK:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    qty = parts[2] if len(parts) > 2 else ""
                    ok_count += 1
                    self._update_item_safe(
                        pn, status=STATUS_OK,
                        detail=f"出货数量={qty}")
                elif line.startswith("ERROR:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    err = parts[2] if len(parts) > 2 else "未知错误"
                    fail_count += 1
                    self._update_item_safe(
                        pn, status=STATUS_FAIL, detail=err)
                elif line.startswith("PROGRESS:"):
                    try:
                        n = int(line.split(":")[1])
                        pct = n / total * 100
                        self._set_progress(pct, f"更新中: {n}/{total}")
                    except ValueError:
                        pass
                elif line.startswith("FATAL:"):
                    msg = line.split(":", 1)[1]
                    self.root.after(
                        0, messagebox.showerror, "错误", msg)

            success, err_msg = run_vbs_script(vbs, on_vbs_line)

            if success:
                total_prints = sum(len(t[2]) for t in self.print_tasks)
                self.root.after(0, lambda: self.progress_var.set(100))
                segments = [
                    (f"成功 {ok_count}", "#4CAF50"),
                    (f"  |  待打印 {total_prints} 份", "#333333"),
                ]
                if fail_count:
                    segments.append((f"  |  更新失败 {fail_count}", "#F44336"))
                if no_file_count:
                    segments.append(
                        (f"  |  无报告 {no_file_count}", "#F44336"))
                self._set_status(segments)
                if ok_count > 0:
                    self._set_buttons(update=True, print_=True)
                else:
                    self._set_buttons(update=True)
            else:
                self._set_progress(0, f"执行出错: {err_msg}")
                self._set_buttons(update=True)

        except Exception as e:
            self._set_progress(0, f"出错: {e}")
            import traceback
            traceback.print_exc()
            self._set_buttons(update=True)

    # ---- 一键打印 ----

    def _on_print(self):
        if not self.print_tasks:
            messagebox.showinfo("提示", "没有需要打印的报告")
            return

        total_prints = sum(len(t[2]) for t in self.print_tasks)
        if not messagebox.askyesno(
            "确认打印",
            f"即将导出并打印 {total_prints} 份检验报告，确定吗？"
        ):
            return

        self._set_buttons(update=False, print_=False)
        self.progress_var.set(0)

        threading.Thread(target=self._do_print, daemon=True).start()

    def _do_print(self):
        try:
            from pypdf import PdfReader, PdfWriter

            total_prints = sum(len(t[2]) for t in self.print_tasks)

            # 创建临时目录存放导出的 PDF
            pdf_dir = tempfile.mkdtemp(prefix="jiuyi_pdf_")

            # 标记所有为导出中
            for pn, _, _ in self.print_tasks:
                self._update_item_safe(pn, status=STATUS_EXPORTING)

            self._set_progress(0, f"导出PDF: 0/{total_prints}")

            # 第一步：VBS 导出 PDF
            vbs, expected_pdfs = VBScriptGenerator.generate_export_pdf_script(
                self.print_tasks, pdf_dir)
            export_ok = 0
            export_fail = 0

            def on_vbs_line(line):
                nonlocal export_ok, export_fail
                if line.startswith("EXPORTED:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    qty = parts[2] if len(parts) > 2 else ""
                    export_ok += 1
                    self._update_item_safe(
                        pn, status=STATUS_EXPORTED,
                        detail=f"已导出 数量={qty}")
                elif line.startswith("EXPORT_ERROR:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    err = parts[2] if len(parts) > 2 else "导出失败"
                    export_fail += 1
                    self._update_item_safe(
                        pn, status=STATUS_EXPORT_FAIL, detail=err)
                elif line.startswith("ERROR:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    err = parts[2] if len(parts) > 2 else ""
                    export_fail += 1
                    self._update_item_safe(
                        pn, status=STATUS_EXPORT_FAIL, detail=err)
                elif line.startswith("PROGRESS:"):
                    try:
                        n = int(line.split(":")[1])
                        pct = n / total_prints * 50  # 导出占 50%
                        self._set_progress(
                            pct, f"导出PDF: {n}/{total_prints}")
                    except ValueError:
                        pass
                elif line.startswith("FATAL:"):
                    msg = line.split(":", 1)[1]
                    self.root.after(
                        0, messagebox.showerror, "错误", msg)

            success, err_msg = run_vbs_script(vbs, on_vbs_line)

            if not success:
                self._set_progress(0, f"导出出错: {err_msg}")
                self._set_buttons(update=True, print_=True)
                return

            if export_ok == 0:
                self._set_progress(0, "所有报告导出失败")
                self._set_buttons(update=True, print_=True)
                return

            # 第二步：倒序合并 PDF
            self._set_progress(55, "正在合并PDF...")
            for pn, _, _ in self.print_tasks:
                self._update_item_safe(pn, status=STATUS_MERGING)

            writer = PdfWriter()
            # 收集实际存在的 PDF 文件（按导出顺序）
            existing_pdfs = []
            for pdf_name in expected_pdfs:
                pdf_path = os.path.join(pdf_dir, pdf_name)
                if os.path.isfile(pdf_path):
                    existing_pdfs.append(pdf_path)

            # 倒序合并：打印机输出后纸张顺序为正序
            for pdf_path in reversed(existing_pdfs):
                reader = PdfReader(pdf_path)
                for page in reader.pages:
                    writer.add_page(page)

            merged_path = os.path.join(pdf_dir, "合并打印.pdf")
            with open(merged_path, 'wb') as f:
                writer.write(f)

            self._set_progress(80, "PDF合并完成，正在打印...")

            # 第三步：自动打印
            for pn, _, _ in self.print_tasks:
                self._update_item_safe(pn, status=STATUS_PRINTING)

            os.startfile(merged_path, 'print')

            # 完成
            for pn, _, _ in self.print_tasks:
                self._update_item_safe(pn, status=STATUS_PRINTED)

            self._set_progress(100, [
                (f"打印完成：共 {len(existing_pdfs)} 份", "#4CAF50"),
                (f"  |  合并PDF: {merged_path}", "#333333"),
            ])

        except Exception as e:
            self._set_progress(0, f"打印出错: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self._set_buttons(update=True, print_=True)

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
