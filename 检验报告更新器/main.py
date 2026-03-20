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
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import subprocess
import tempfile
import threading
from datetime import date
from collections import OrderedDict

import openpyxl


# ---------------------------------------------------------------------------
# 发货计划读取
# ---------------------------------------------------------------------------

def read_shipping_plan(filepath):
    """读取发货计划，提取 (料件编号, 本周要求数量, 行号)"""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    items = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) < 7:
            continue
        part_no = row[3]   # 列 D: 料件编号
        quantity = row[6]  # 列 G: 本周要求数量
        if part_no and quantity is not None:
            try:
                qty = int(float(quantity))
            except (ValueError, TypeError):
                continue
            items.append({
                'row': row_idx,
                'part_no': str(part_no).strip(),
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
    def generate_update_script(tasks, today):
        """生成更新报告的 VBS 脚本

        tasks: [(part_no, filepath, final_quantity), ...]
        today: date object
        """
        lines = [VBScriptGenerator._header()]
        lines.append(f'Dim d')
        lines.append(f'd = DateSerial({today.year}, {today.month}, {today.day})')
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
            lines.append(f'    ws.Cells(2, 14).Value = {quantity}')
            lines.append(f'    ws.Cells(3, 14).Value = d')
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
    def generate_print_script(tasks, today):
        """生成打印报告的 VBS 脚本

        tasks: [(part_no, filepath, [qty1, qty2, ...]), ...]
        对于只有一个数量的任务，直接打印。
        对于多个数量（split），逐个设值后打印，最后保存最后一个值。
        """
        lines = [VBScriptGenerator._header()]
        lines.append(f'd = DateSerial({today.year}, {today.month}, {today.day})')
        lines.append('')
        lines.append('On Error Resume Next')
        lines.append('')

        print_idx = 0
        for part_no, filepath, quantities in tasks:
            fp = filepath.replace('"', '""')
            lines.append(f'Set wb = app.Workbooks.Open("{fp}")')
            lines.append('If Err.Number <> 0 Then')
            lines.append(f'    WScript.Echo "ERROR:{part_no}:无法打开文件"')
            lines.append('    Err.Clear')
            lines.append('Else')

            if len(quantities) == 1:
                # 单数量：直接打印
                lines.append('    wb.PrintOut')
                lines.append('    If Err.Number <> 0 Then')
                lines.append(f'        WScript.Echo "PRINT_ERROR:{part_no}:打印失败"')
                lines.append('        Err.Clear')
                lines.append('    Else')
                lines.append(f'        WScript.Echo "PRINTED:{part_no}:{quantities[0]}"')
                lines.append('    End If')
                lines.append('    wb.Close False')
                print_idx += 1
            else:
                # 多数量(split)：逐个设值后打印
                lines.append('    Set ws = wb.Sheets(1)')
                for qty in quantities:
                    lines.append(f'    ws.Cells(2, 14).Value = {qty}')
                    lines.append(f'    ws.Cells(3, 14).Value = d')
                    lines.append('    wb.PrintOut')
                    lines.append('    If Err.Number <> 0 Then')
                    lines.append(f'        WScript.Echo "PRINT_ERROR:{part_no}:打印失败(数量={qty})"')
                    lines.append('        Err.Clear')
                    lines.append('    Else')
                    lines.append(f'        WScript.Echo "PRINTED:{part_no}:{qty}"')
                    lines.append('    End If')
                    print_idx += 1
                # 保存最后一个值
                lines.append('    wb.Save')
                lines.append('    wb.Close False')

            lines.append('End If')
            lines.append(f'WScript.Echo "PROGRESS:{print_idx}"')
            lines.append('')

        lines.append(VBScriptGenerator._footer())
        return '\n'.join(lines)


# ---------------------------------------------------------------------------
# VBS 执行器
# ---------------------------------------------------------------------------

def run_vbs_script(vbs_content, on_line=None):
    """将 VBS 内容写入临时文件并通过 cscript 执行

    on_line: 回调函数，每读到一行 stdout 时调用
    返回: (success: bool, error_msg: str)
    """
    # Windows 使用 mbcs (系统 ANSI 编码), 其他平台用 utf-8
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

        # 居中于父窗口
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
# 主应用 GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("久益 · 检验报告更新器")
        self.root.geometry("720x580")
        self.root.minsize(620, 480)

        # Windows DPI 适配
        if sys.platform == 'win32':
            try:
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        self.print_tasks = []   # [(part_no, filepath, [qty1, ...])]
        self.processing = False

        self._build_ui()

    # ---- UI 构建 ----

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
        ttk.Label(r1, text="发货计划：", width=10,
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
            bf, text="  ② 一键打印  ", command=self._on_print,
            state='disabled')
        self.btn_print.pack(side='right', expand=True, fill='x',
                            padx=(5, 0), ipady=10)

        # 进度条
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main, variable=self.progress_var,
                        maximum=100).pack(fill='x', pady=(0, 3))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main, textvariable=self.status_var,
                  font=("Microsoft YaHei", 9)).pack(anchor='w')

        # 日志区
        lf = ttk.LabelFrame(main, text=" 处理日志 ", padding=5)
        lf.pack(fill='both', expand=True, pady=(5, 0))
        self.log = scrolledtext.ScrolledText(
            lf, height=12, font=("Consolas", 9),
            state='disabled', wrap='word')
        self.log.pack(fill='both', expand=True)

    # ---- 工具方法 ----

    def _append_log(self, msg):
        self.log.configure(state='normal')
        self.log.insert('end', msg + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _log(self, msg):
        """线程安全日志"""
        self.root.after(0, self._append_log, msg)

    def _set_progress(self, value, text=None):
        self.root.after(0, lambda: self.progress_var.set(value))
        if text:
            self.root.after(0, lambda: self.status_var.set(text))

    def _set_buttons(self, update=None, print_=None):
        if update is not None:
            self.root.after(0, lambda: self.btn_update.configure(
                state='normal' if update else 'disabled'))
        if print_ is not None:
            self.root.after(0, lambda: self.btn_print.configure(
                state='normal' if print_ else 'disabled'))

    def _clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    # ---- 文件选择 ----

    def _pick_plan(self):
        p = filedialog.askopenfilename(
            title="选择发货计划",
            filetypes=[("Excel 文件", "*.xlsx;*.xls"), ("所有文件", "*.*")])
        if p:
            self.plan_var.set(p)

    def _pick_dir(self):
        d = filedialog.askdirectory(title="选择检验报告目录")
        if d:
            self.dir_var.set(d)

    # ---- 更新报告 ----

    def _on_update(self):
        plan_path = self.plan_var.get().strip()
        report_dir = self.dir_var.get().strip()

        if not plan_path or not os.path.isfile(plan_path):
            messagebox.showerror("错误", "请选择有效的发货计划文件")
            return
        if not report_dir or not os.path.isdir(report_dir):
            messagebox.showerror("错误", "请选择有效的检验报告目录")
            return

        self._set_buttons(update=False, print_=False)
        self.print_tasks = []
        self._clear_log()
        self.progress_var.set(0)

        threading.Thread(
            target=self._do_update, args=(plan_path, report_dir),
            daemon=True
        ).start()

    def _do_update(self, plan_path, report_dir):
        try:
            # 1. 读取发货计划
            self._log("正在读取发货计划...")
            items = read_shipping_plan(plan_path)
            self._log(f"共 {len(items)} 条记录")

            # 2. 按料件编号分组
            groups = group_by_part_no(items)
            self._log(f"共 {len(groups)} 个不同料件编号\n")

            # 3. 构建任务列表 + 处理重复项
            update_tasks = []   # [(part_no, filepath, final_qty)]
            self.print_tasks = []

            for part_no, group_items in groups.items():
                report_file = find_report_file(report_dir, part_no)
                if not report_file:
                    self._log(f"  [跳过] {part_no} — 未找到报告文件")
                    continue

                if len(group_items) == 1:
                    qty = group_items[0]['quantity']
                    update_tasks.append((part_no, report_file, qty))
                    self.print_tasks.append((part_no, report_file, [qty]))
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
                        update_tasks.append((part_no, report_file, total))
                        self.print_tasks.append(
                            (part_no, report_file, [total]))
                        self._log(f"  [合并] {part_no} — 总计 {total:,}")
                    elif choice == 'split':
                        qtys = [i['quantity'] for i in group_items]
                        # 更新时用最后一个数量
                        update_tasks.append(
                            (part_no, report_file, qtys[-1]))
                        self.print_tasks.append(
                            (part_no, report_file, qtys))
                        self._log(
                            f"  [分开] {part_no} — {len(qtys)} 份")
                    else:
                        self._log(f"  [跳过] {part_no}")
                        continue

            if not update_tasks:
                self._log("\n没有需要处理的报告")
                self._set_buttons(update=True)
                return

            # 4. 生成并执行 VBS 更新脚本
            self._log(f"\n准备更新 {len(update_tasks)} 个报告文件...")
            today = date.today()
            vbs = VBScriptGenerator.generate_update_script(
                update_tasks, today)
            total = len(update_tasks)

            def on_vbs_line(line):
                if line.startswith("OK:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    qty = parts[2] if len(parts) > 2 else ""
                    self._log(f"  [完成] {pn} — 出货数量={qty}")
                elif line.startswith("ERROR:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    err = parts[2] if len(parts) > 2 else "未知错误"
                    self._log(f"  [失败] {pn} — {err}")
                elif line.startswith("PROGRESS:"):
                    try:
                        n = int(line.split(":")[1])
                        pct = n / total * 100
                        self._set_progress(pct, f"更新中: {n}/{total}")
                    except ValueError:
                        pass
                elif line.startswith("FATAL:"):
                    msg = line.split(":", 1)[1]
                    self._log(f"\n[错误] {msg}")
                    self.root.after(
                        0, messagebox.showerror, "错误", msg)

            success, err_msg = run_vbs_script(vbs, on_vbs_line)

            if success:
                self._log(
                    f"\n更新完成！共处理 {len(update_tasks)} 个报告")
                self._set_progress(100, "更新完成，可点击「一键打印」")
                self._set_buttons(update=True, print_=True)
            else:
                self._log(f"\n执行出错: {err_msg}")
                self._set_buttons(update=True)

        except Exception as e:
            self._log(f"\n处理出错: {e}")
            import traceback
            self._log(traceback.format_exc())
            self._set_buttons(update=True)

    # ---- 一键打印 ----

    def _on_print(self):
        if not self.print_tasks:
            messagebox.showinfo("提示", "没有需要打印的报告")
            return

        total_prints = sum(len(t[2]) for t in self.print_tasks)
        if not messagebox.askyesno(
            "确认打印",
            f"即将打印 {total_prints} 份检验报告，确定吗？"
        ):
            return

        self._set_buttons(update=False, print_=False)
        self.progress_var.set(0)

        threading.Thread(target=self._do_print, daemon=True).start()

    def _do_print(self):
        try:
            today = date.today()
            total_prints = sum(len(t[2]) for t in self.print_tasks)

            self._log(f"\n开始打印 {total_prints} 份报告...\n")

            vbs = VBScriptGenerator.generate_print_script(
                self.print_tasks, today)

            def on_vbs_line(line):
                if line.startswith("PRINTED:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    qty = parts[2] if len(parts) > 2 else ""
                    self._log(f"  [打印] {pn} — 数量={qty}")
                elif line.startswith("PRINT_ERROR:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    err = parts[2] if len(parts) > 2 else ""
                    self._log(f"  [打印失败] {pn} — {err}")
                elif line.startswith("ERROR:"):
                    parts = line.split(":", 2)
                    pn = parts[1] if len(parts) > 1 else ""
                    err = parts[2] if len(parts) > 2 else ""
                    self._log(f"  [失败] {pn} — {err}")
                elif line.startswith("PROGRESS:"):
                    try:
                        n = int(line.split(":")[1])
                        pct = n / total_prints * 100
                        self._set_progress(pct, f"打印中: {n}/{total_prints}")
                    except ValueError:
                        pass
                elif line.startswith("FATAL:"):
                    msg = line.split(":", 1)[1]
                    self._log(f"\n[错误] {msg}")
                    self.root.after(
                        0, messagebox.showerror, "错误", msg)

            success, err_msg = run_vbs_script(vbs, on_vbs_line)

            if success:
                self._log(f"\n打印完成！")
                self._set_progress(100, "打印完成")
            else:
                self._log(f"\n打印出错: {err_msg}")

        except Exception as e:
            self._log(f"\n打印出错: {e}")
        finally:
            self._set_buttons(update=True, print_=True)

    # ---- 启动 ----

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
