#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import glob
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

# ----------------------------------------------------------------------
# 公用函数
# ----------------------------------------------------------------------

ATOMIC_NUMBER_TO_SYMBOL = {
    1: 'H', 3: 'Li', 4: 'Be', 5: 'B', 6: 'C', 7: 'N', 8: 'O', 9: 'F',
    14: 'Si', 15: 'P', 16: 'S', 17: 'Cl', 34: 'Se', 35: 'Br', 53: 'I'
}

PRESET_RESOURCES = {
    "hachimi单并行": {"nproc": "10", "mem": "40GB"},
    "hachimi四并行": {"nproc": "4", "mem": "10GB"},
    "Tomori八队列": {"nproc": "12", "mem": "12GB"},
    "students/zstoffice": {"nproc": "8", "mem": "20GB"},
    "zst106": {"nproc": "24", "mem": "180GB"}
}


def modify_gjf_content(input_path, output_path, mem, nprocshared, keyword, charge, mult, chk_name=None):
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    if chk_name is None:
        chk_filename = f"{output_path}.chk"
    else:
        chk_filename = os.path.basename(chk_name)

    new_lines = []
    i = 0
    chk_replaced = False
    mem_replaced = False
    nproc_replaced = False
    keyword_replaced = False
    charge_mult_replaced = False
    title_found = False

    while i < len(lines):
        line = lines[i]

        if re.match(r'^%chk\s*=', line, re.IGNORECASE):
            new_lines.append(f"%chk={chk_filename}\n")
            chk_replaced = True
            i += 1
            continue

        elif re.match(r'^%mem\s*=', line, re.IGNORECASE):
            new_lines.append(f"%mem={mem}\n")
            mem_replaced = True
            i += 1
            continue

        elif re.match(r'^%nproc(shared)?\s*=', line, re.IGNORECASE):
            new_lines.append(f"%nprocshared={nprocshared}\n")
            nproc_replaced = True
            i += 1
            continue

        elif re.match(r'^\s*#', line) and not keyword_replaced:
            new_lines.append(f"{keyword}\n")
            keyword_replaced = True
            i += 1
            while i < len(lines):
                nxt = lines[i]
                if nxt.strip() and (nxt.strip().startswith('#') or nxt[0] == ' '):
                    i += 1
                else:
                    break
            continue

        elif title_found and not charge_mult_replaced and re.match(r'^\s*[-+]?\d+\s+[-+]?\d+\s*$', line):
            new_lines.append(f"{charge} {mult}\n")
            charge_mult_replaced = True
            i += 1
            continue

        elif not title_found and not re.match(r'^\s*%', line) and not re.match(r'^\s*#', line) and line.strip():
            new_lines.append(line)
            title_found = True
            i += 1
            continue

        else:
            new_lines.append(line)
            i += 1

    if not chk_replaced:
        new_lines.insert(0, f"%chk={chk_filename}\n")
    if not mem_replaced:
        insert_pos = 1 if not chk_replaced else 1
        new_lines.insert(insert_pos, f"%mem={mem}\n")
    if not nproc_replaced:
        insert_pos = 0
        if not chk_replaced:
            insert_pos += 1
        if not mem_replaced:
            insert_pos += 1
        new_lines.insert(insert_pos, f"%nprocshared={nprocshared}\n")
    if not keyword_replaced:
        insert_idx = 0
        for idx, ln in enumerate(new_lines):
            if not ln.startswith('%'):
                insert_idx = idx
                break
        else:
            insert_idx = len(new_lines)
        new_lines.insert(insert_idx, f"{keyword}\n")

    return new_lines


def parse_log_last_structure(filename):
    try:
        with open(f"{filename}.log", 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
    except FileNotFoundError:
        return None, None

    atomic_numbers = []
    coordinates = []
    re_std = re.compile(r'^\s*Standard\s+orientation\s*:', re.IGNORECASE)
    re_coord = re.compile(
        r'^\s*(\d+)\s+(\d+)\s+(\d+)\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)'
    )
    in_std = False
    current_atomic = []
    current_coords = []

    for line in lines:
        if re_std.search(line):
            in_std = True
            current_atomic = []
            current_coords = []
            continue

        if in_std:
            if line.strip().startswith('---'):
                if current_atomic:
                    atomic_numbers = current_atomic
                    coordinates = current_coords
                    in_std = False
                continue
            m = re_coord.match(line)
            if m:
                an = int(m.group(2))
                x = float(m.group(4))
                y = float(m.group(5))
                z = float(m.group(6))
                current_atomic.append(an)
                current_coords.append((x, y, z))

    return atomic_numbers, coordinates


def write_gjf_from_coords(output_path, mem, nprocshared, keyword, charge, mult,
                          atomic_numbers, coordinates, title="Generated from log"):
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(f"%chk={os.path.basename(output_path).replace('.gjf', '.chk')}\n")
        f.write(f"%mem={mem}\n")
        f.write(f"%nprocshared={nprocshared}\n")
        f.write(f"{keyword}\n")
        f.write("\n")
        f.write(f"{title}\n")
        f.write("\n")
        f.write(f"{charge} {mult}\n")
        for an, (x, y, z) in zip(atomic_numbers, coordinates):
            sym = ATOMIC_NUMBER_TO_SYMBOL.get(an, f"X{an}")
            f.write(f" {sym:<2s}  {x:12.6f} {y:12.6f} {z:12.6f}\n")
        f.write("\n")


def extract_scan_header_info(lines):
    route_lines = []
    title_lines = []
    charge = 0
    mult = 1
    for line in lines:
        if 'Charge =' in line and 'Multiplicity =' in line:
            m = re.search(r'Charge\s*=\s*(-?\d+)\s+Multiplicity\s*=\s*(\d+)', line)
            if m:
                charge = int(m.group(1))
                mult = int(m.group(2))
                break
    if not route_lines:
        for line in lines:
            if line.strip().startswith('#'):
                route_lines.append(line.strip())
                break
    if not title_lines:
        for line in lines:
            s = line.strip()
            if s and not s.startswith('#') and not s.startswith('%'):
                if 'Entering' not in s and 'Link' not in s:
                    title_lines = [s]
                    break
    route_str = '\n'.join(route_lines) if route_lines else '#P B3LYP/6-31G(d)'
    title_str = ' '.join(title_lines) if title_lines else 'Scan Point'
    return route_str, title_str, charge, mult


def parse_standard_orientation_at(lines, start):
    for i in range(start, len(lines)):
        if 'Standard orientation:' in lines[i]:
            j = i + 5
            atomic_numbers = []
            coordinates = []
            while j < len(lines):
                line = lines[j].strip()
                if '----' in line or line == '':
                    break
                parts = line.split()
                if len(parts) >= 6:
                    try:
                        num = int(parts[1])
                        x = float(parts[3])
                        y = float(parts[4])
                        z = float(parts[5])
                        atomic_numbers.append(num)
                        coordinates.append((x, y, z))
                    except ValueError:
                        break
                j += 1
            if atomic_numbers:
                return atomic_numbers, coordinates
    return None, None


def extract_modredundant_scan_steps(lines):
    steps = []
    current_scan_point = None
    last_std_orient_atoms = None
    last_std_orient_coords = None
    converged_atoms = None
    converged_coords = None
    scan_point_pattern = re.compile(r'on scan point\s+(\d+)\s+out of\s+\d+')

    for i, line in enumerate(lines):
        m = scan_point_pattern.search(line)
        if m:
            new_scan_point = int(m.group(1))
            if current_scan_point is not None and new_scan_point != current_scan_point:
                if converged_atoms is not None:
                    steps.append((current_scan_point, converged_atoms, converged_coords))
                converged_atoms = None
                converged_coords = None
                last_std_orient_atoms = None
                last_std_orient_coords = None
            current_scan_point = new_scan_point

        if 'Standard orientation:' in line:
            atoms, coords = parse_standard_orientation_at(lines, i)
            if atoms is not None:
                last_std_orient_atoms = atoms
                last_std_orient_coords = coords

        if 'Optimization completed.' in line:
            if last_std_orient_atoms is not None:
                converged_atoms = last_std_orient_atoms
                converged_coords = last_std_orient_coords

    if current_scan_point is not None and converged_atoms is not None:
        steps.append((current_scan_point, converged_atoms, converged_coords))

    return steps


# ----------------------------------------------------------------------
# 图形界面
# ----------------------------------------------------------------------

class GaussianToolGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gaussian 辅助工具")
        self.geometry("850x700")
        self.resizable(True, True)

        mode_frame = ttk.LabelFrame(self, text="操作模式", padding=5)
        mode_frame.pack(fill=tk.X, padx=5, pady=5)

        self.mode_var = tk.StringVar(value="modify_gjf")
        modes = [
            ("修改 GJF 参数", "modify_gjf"),
            ("从 LOG 生成 GJF", "log_to_gjf"),
            ("提取扫描构象", "extract_scan")
        ]
        for text, value in modes:
            ttk.Radiobutton(mode_frame, text=text, variable=self.mode_var,
                            value=value, command=self.on_mode_change).pack(side=tk.LEFT, padx=10)

        self.param_frame = None
        log_frame = ttk.LabelFrame(self, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        self.run_btn = ttk.Button(btn_frame, text="运行", command=self.start_task)
        self.run_btn.pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=5)

        self.current_widgets = []
        self.on_mode_change()

    def on_mode_change(self):
        if self.param_frame is not None:
            self.param_frame.destroy()
        self.param_frame = ttk.Frame(self)
        self.param_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.current_widgets.clear()

        mode = self.mode_var.get()
        if mode == "modify_gjf":
            self.create_modify_gjf_widgets()
        elif mode == "log_to_gjf":
            self.create_log_to_gjf_widgets()
        else:
            self.create_extract_scan_widgets()

    # ---------- 辅助方法：添加上拉菜单资源预设 ----------
    def add_resource_preset_ui(self, parent, mem_var, nproc_var):
        """在父容器中添加上拉菜单（资源预设），选择后自动填充mem和nproc"""
        row = len(parent.grid_slaves())
        ttk.Label(parent, text="资源预设:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)

        # 当前选中的预设文本，默认改为 students/zstoffice
        self.current_preset = tk.StringVar(value="students/zstoffice")

        # 按钮显示当前预设
        preset_btn = ttk.Button(parent, textvariable=self.current_preset, width=20)
        preset_btn.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)

        # 创建弹出菜单（上拉位置）
        def show_menu(event):
            x = preset_btn.winfo_rootx()
            y = preset_btn.winfo_rooty()
            menu_height = len(PRESET_RESOURCES) * 22
            menu = tk.Menu(parent, tearoff=0)
            for name, conf in PRESET_RESOURCES.items():
                menu.add_command(label=name, command=lambda n=name, c=conf: (
                    nproc_var.set(c["nproc"]),
                    mem_var.set(c["mem"]),
                    self.current_preset.set(n)
                ))
            # 在按钮上方弹出
            menu.post(x, y - menu_height)

        preset_btn.bind("<Button-1>", show_menu)

        self.current_widgets.append(preset_btn)
        # 初始化默认值改为 students/zstoffice
        default = PRESET_RESOURCES["students/zstoffice"]
        nproc_var.set(default["nproc"])
        mem_var.set(default["mem"])
        self.current_preset.set("students/zstoffice")

        return preset_btn

    # ---------- 修改 GJF 参数模式 ----------
    def create_modify_gjf_widgets(self):
        row = 0

        self.input_folder_var = tk.StringVar()
        ttk.Label(self.param_frame, text="输入文件夹 (包含 .gjf):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.input_folder_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览", command=lambda: self.select_folder(self.input_folder_var)).grid(
            row=row, column=2, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        self.output_folder_var = tk.StringVar()
        ttk.Label(self.param_frame, text="输出文件夹:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.output_folder_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览", command=lambda: self.select_folder(self.output_folder_var)).grid(
            row=row, column=2, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        self.file_prefix_var = tk.StringVar()
        ttk.Label(self.param_frame, text="文件名前缀:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.file_prefix_var, width=30)
        entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        self.keyword_var = tk.StringVar(value="#p opt b3lpy/6-31g(d,p)")
        ttk.Label(self.param_frame, text="关键词行:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        kw_entry = ttk.Entry(self.param_frame, textvariable=self.keyword_var, width=60)
        kw_entry.grid(row=row, column=1, padx=5, pady=2, sticky=tk.W)
        self.current_widgets.extend([kw_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # 电荷/自旋多重度
        self.charge_var = tk.StringVar(value="0")
        self.mult_var = tk.StringVar(value="1")
        ttk.Label(self.param_frame, text="电荷 / 自旋:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        cm_frame = ttk.Frame(self.param_frame)
        cm_frame.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(cm_frame, text="电荷:").pack(side=tk.LEFT, padx=2)
        e1 = ttk.Entry(cm_frame, textvariable=self.charge_var, width=5)
        e1.pack(side=tk.LEFT, padx=2)
        ttk.Label(cm_frame, text="自旋:").pack(side=tk.LEFT, padx=2)
        e2 = ttk.Entry(cm_frame, textvariable=self.mult_var, width=5)
        e2.pack(side=tk.LEFT, padx=2)
        self.current_widgets.extend([cm_frame, e1, e2, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # %mem
        self.mem_var = tk.StringVar(value="40GB")
        ttk.Label(self.param_frame, text="%mem:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        mem_entry = ttk.Entry(self.param_frame, textvariable=self.mem_var, width=20)
        mem_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([mem_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # %nprocshared
        self.nproc_var = tk.StringVar(value="10")
        ttk.Label(self.param_frame, text="%nprocshared:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        nproc_entry = ttk.Entry(self.param_frame, textvariable=self.nproc_var, width=20)
        nproc_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([nproc_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # 资源预设（上拉菜单）
        self.add_resource_preset_ui(self.param_frame, self.mem_var, self.nproc_var)
        row += 1

        self.param_frame.grid_rowconfigure(row, weight=1)

    # ---------- 从 LOG 生成 GJF 模式 ----------
    def create_log_to_gjf_widgets(self):
        row = 0

        self.input_folder_var = tk.StringVar()
        ttk.Label(self.param_frame, text="输入文件夹 (包含 .log):").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.input_folder_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览", command=lambda: self.select_folder(self.input_folder_var)).grid(
            row=row, column=2, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        self.output_folder_var = tk.StringVar()
        ttk.Label(self.param_frame, text="输出文件夹:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.output_folder_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览", command=lambda: self.select_folder(self.output_folder_var)).grid(
            row=row, column=2, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        self.file_prefix_var = tk.StringVar()
        ttk.Label(self.param_frame, text="文件名前缀:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.file_prefix_var, width=30)
        entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        self.keyword_var = tk.StringVar(value="#p opt b3lpy/6-31g(d,p)")
        ttk.Label(self.param_frame, text="关键词行:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        kw_entry = ttk.Entry(self.param_frame, textvariable=self.keyword_var, width=60)
        kw_entry.grid(row=row, column=1, padx=5, pady=2, sticky=tk.W)
        self.current_widgets.extend([kw_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # 电荷/自旋多重度
        self.charge_var = tk.StringVar(value="0")
        self.mult_var = tk.StringVar(value="1")
        ttk.Label(self.param_frame, text="电荷 / 自旋:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        cm_frame = ttk.Frame(self.param_frame)
        cm_frame.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(cm_frame, text="电荷:").pack(side=tk.LEFT, padx=2)
        e1 = ttk.Entry(cm_frame, textvariable=self.charge_var, width=5)
        e1.pack(side=tk.LEFT, padx=2)
        ttk.Label(cm_frame, text="自旋:").pack(side=tk.LEFT, padx=2)
        e2 = ttk.Entry(cm_frame, textvariable=self.mult_var, width=5)
        e2.pack(side=tk.LEFT, padx=2)
        self.current_widgets.extend([cm_frame, e1, e2, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # %mem
        self.mem_var = tk.StringVar(value="40GB")
        ttk.Label(self.param_frame, text="%mem:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        mem_entry = ttk.Entry(self.param_frame, textvariable=self.mem_var, width=20)
        mem_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([mem_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # %nprocshared
        self.nproc_var = tk.StringVar(value="10")
        ttk.Label(self.param_frame, text="%nprocshared:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        nproc_entry = ttk.Entry(self.param_frame, textvariable=self.nproc_var, width=20)
        nproc_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([nproc_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # 资源预设（上拉菜单）
        self.add_resource_preset_ui(self.param_frame, self.mem_var, self.nproc_var)
        row += 1

        self.param_frame.grid_rowconfigure(row, weight=1)

    # ---------- 提取扫描构象模式 ----------
    def create_extract_scan_widgets(self):
        row = 0

        self.input_file_var = tk.StringVar()
        ttk.Label(self.param_frame, text="输入 LOG 文件:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.input_file_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览", command=lambda: self.select_file(self.input_file_var)).grid(row=row,
                                                                                                              column=2,
                                                                                                              padx=5,
                                                                                                              pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        self.output_dir_var = tk.StringVar()
        ttk.Label(self.param_frame, text="输出目录:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.output_dir_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览", command=lambda: self.select_folder(self.output_dir_var)).grid(row=row,
                                                                                                                column=2,
                                                                                                                padx=5,
                                                                                                                pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        self.override_route_var = tk.BooleanVar(value=False)
        self.add_resources_var = tk.BooleanVar(value=True)

        cb1 = ttk.Checkbutton(self.param_frame, text="覆盖关键词行", variable=self.override_route_var,
                              command=self.toggle_route_entry)
        cb1.grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.append(cb1)
        self.route_entry = ttk.Entry(self.param_frame, width=60, state='disabled')
        self.route_entry.grid(row=row, column=1, padx=5, pady=2, sticky=tk.W)
        self.current_widgets.append(self.route_entry)
        row += 1

        self.override_cm_var = tk.BooleanVar(value=False)
        cb2 = ttk.Checkbutton(self.param_frame, text="覆盖电荷/自旋多重度", variable=self.override_cm_var,
                              command=self.toggle_cm_entries)
        cb2.grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.append(cb2)

        cm_frame = ttk.Frame(self.param_frame)
        cm_frame.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(cm_frame, text="电荷:").pack(side=tk.LEFT, padx=2)
        self.charge_scan_var = tk.StringVar(value="0")
        e1 = ttk.Entry(cm_frame, textvariable=self.charge_scan_var, width=5, state='disabled')
        e1.pack(side=tk.LEFT, padx=2)
        ttk.Label(cm_frame, text="自旋:").pack(side=tk.LEFT, padx=2)
        self.mult_scan_var = tk.StringVar(value="1")
        e2 = ttk.Entry(cm_frame, textvariable=self.mult_scan_var, width=5, state='disabled')
        e2.pack(side=tk.LEFT, padx=2)
        self.current_widgets.extend([cm_frame, e1, e2])
        row += 1

        cb3 = ttk.Checkbutton(self.param_frame, text="添加 %mem / %nprocshared 行", variable=self.add_resources_var)
        cb3.grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.append(cb3)
        row += 1

        # %mem
        self.mem_scan_var = tk.StringVar(value="40GB")
        ttk.Label(self.param_frame, text="%mem:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        mem_entry = ttk.Entry(self.param_frame, textvariable=self.mem_scan_var, width=20)
        mem_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([mem_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # %nprocshared
        self.nproc_scan_var = tk.StringVar(value="10")
        ttk.Label(self.param_frame, text="%nprocshared:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        nproc_entry = ttk.Entry(self.param_frame, textvariable=self.nproc_scan_var, width=20)
        nproc_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([nproc_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        # 资源预设（上拉菜单）
        self.add_resource_preset_ui(self.param_frame, self.mem_scan_var, self.nproc_scan_var)
        row += 1

        self.scan_charge_entry = e1
        self.scan_mult_entry = e2
        self.param_frame.grid_rowconfigure(row, weight=1)

    def toggle_route_entry(self):
        self.route_entry.config(state='normal' if self.override_route_var.get() else 'disabled')

    def toggle_cm_entries(self):
        state = 'normal' if self.override_cm_var.get() else 'disabled'
        self.scan_charge_entry.config(state=state)
        self.scan_mult_entry.config(state=state)

    def select_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)

    def select_file(self, var):
        fname = filedialog.askopenfilename(filetypes=[("LOG files", "*.log")])
        if fname:
            var.set(fname)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def start_task(self):
        self.run_btn.config(state='disabled')
        threading.Thread(target=self.run_task, daemon=True).start()

    def run_task(self):
        try:
            mode = self.mode_var.get()
            if mode == "modify_gjf":
                self.run_modify_gjf()
            elif mode == "log_to_gjf":
                self.run_log_to_gjf()
            else:
                self.run_extract_scan()
        except Exception as e:
            self.log(f"错误: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.run_btn.config(state='normal')

    def run_modify_gjf(self):
        input_folder = self.input_folder_var.get().strip()
        if not input_folder:
            self.log("请选择输入文件夹")
            return
        output_folder = self.output_folder_var.get().strip()
        if not output_folder:
            self.log("请选择输出文件夹")
            return
        prefix = self.file_prefix_var.get().strip()
        if not prefix:
            self.log("请填写文件名前缀")
            return

        mem = self.mem_var.get().strip()
        nproc = self.nproc_var.get().strip()
        keyword = self.keyword_var.get().strip()
        charge = self.charge_var.get().strip()
        mult = self.mult_var.get().strip()

        gjf_files = glob.glob(os.path.join(input_folder, "*.gjf"))
        if not gjf_files:
            self.log("未找到 .gjf 文件")
            return

        os.makedirs(output_folder, exist_ok=True)
        self.log(f"找到 {len(gjf_files)} 个 gjf 文件，输出至 {output_folder}")
        self.log(f"文件名前缀: {prefix}")

        for gjf in gjf_files:
            basename = os.path.basename(gjf)
            name, _ = os.path.splitext(basename)
            output_base = os.path.join(output_folder, f"{prefix}{name}")
            chk_filename = f"{prefix}{name}.chk"
            try:
                new_content = modify_gjf_content(gjf, output_base, mem, nproc, keyword, charge, mult,
                                                 chk_name=chk_filename)
                output_path = output_base + ".gjf"
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.writelines(new_content)
                self.log(f"已生成: {output_path} (chk: {chk_filename})")
            except Exception as e:
                self.log(f"处理 {gjf} 失败: {e}")
        self.log("修改 GJF 参数任务完成。")

    def run_log_to_gjf(self):
        input_folder = self.input_folder_var.get().strip()
        if not input_folder:
            self.log("请选择输入文件夹")
            return
        output_folder = self.output_folder_var.get().strip()
        if not output_folder:
            self.log("请选择输出文件夹")
            return
        prefix = self.file_prefix_var.get().strip()
        if not prefix:
            self.log("请填写文件名前缀")
            return

        mem = self.mem_var.get().strip()
        nproc = self.nproc_var.get().strip()
        keyword = self.keyword_var.get().strip()
        charge = self.charge_var.get().strip()
        mult = self.mult_var.get().strip()

        log_files = glob.glob(os.path.join(input_folder, "*.log"))
        if not log_files:
            self.log("未找到 .log 文件")
            return

        os.makedirs(output_folder, exist_ok=True)
        self.log(f"找到 {len(log_files)} 个 log 文件，输出至 {output_folder}")
        self.log(f"文件名前缀: {prefix}")

        for logf in log_files:
            basename = os.path.basename(logf)
            name, _ = os.path.splitext(basename)
            base_full = os.path.splitext(logf)[0]
            atomic_numbers, coords = parse_log_last_structure(base_full)
            if not atomic_numbers:
                self.log(f"无法从 {logf} 提取结构，跳过")
                continue
            output_name = f"{prefix}{name}"
            output_path = os.path.join(output_folder, f"{output_name}.gjf")
            try:
                write_gjf_from_coords(output_path, mem, nproc, keyword, charge, mult,
                                      atomic_numbers, coords, title=f"From {name}")
                self.log(f"已生成: {output_path}")
            except Exception as e:
                self.log(f"生成 {output_name} 失败: {e}")
        self.log("从 LOG 生成 GJF 任务完成。")

    def run_extract_scan(self):
        input_log = self.input_file_var.get().strip()
        if not input_log or not os.path.exists(input_log):
            self.log("请选择有效的 LOG 文件")
            return
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            output_dir = os.path.dirname(input_log)
        os.makedirs(output_dir, exist_ok=True)

        with open(input_log, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        orig_route, orig_title, orig_charge, orig_mult = extract_scan_header_info(lines)
        if self.override_route_var.get() and self.route_entry.get().strip():
            route = self.route_entry.get().strip()
        else:
            route = orig_route
        if self.override_cm_var.get():
            try:
                charge = int(self.charge_scan_var.get())
                mult = int(self.mult_scan_var.get())
            except:
                charge, mult = orig_charge, orig_mult
        else:
            charge, mult = orig_charge, orig_mult

        steps = extract_modredundant_scan_steps(lines)
        if not steps:
            self.log("未找到任何扫描步结构。")
            return

        base_name = os.path.splitext(os.path.basename(input_log))[0]
        self.log(f"共找到 {len(steps)} 个扫描构象，输出至 {output_dir}")

        for step_id, atomic_numbers, coords in steps:
            gjf_name = f"{base_name}_ScanPoint{step_id}.gjf"
            out_path = os.path.join(output_dir, gjf_name)
            try:
                if self.add_resources_var.get():
                    mem = self.mem_scan_var.get().strip()
                    nproc = self.nproc_scan_var.get().strip()
                    with open(out_path, 'w', encoding='utf-8') as f:
                        f.write(f"%chk={gjf_name.replace('.gjf', '.chk')}\n")
                        f.write(f"%mem={mem}\n")
                        f.write(f"%nprocshared={nproc}\n")
                        f.write(f"{route}\n")
                        f.write("\n")
                        f.write(f"{orig_title} ScanPoint {step_id}\n")
                        f.write("\n")
                        f.write(f"{charge} {mult}\n")
                        for an, (x, y, z) in zip(atomic_numbers, coords):
                            sym = ATOMIC_NUMBER_TO_SYMBOL.get(an, f"X{an}")
                            f.write(f" {sym:<2s}  {x:12.6f} {y:12.6f} {z:12.6f}\n")
                        f.write("\n")
                else:
                    with open(out_path, 'w', encoding='utf-8') as f:
                        f.write(f"{route}\n\n")
                        f.write(f"{orig_title} ScanPoint {step_id}\n\n")
                        f.write(f"{charge} {mult}\n")
                        for an, (x, y, z) in zip(atomic_numbers, coords):
                            sym = ATOMIC_NUMBER_TO_SYMBOL.get(an, f"X{an}")
                            f.write(f" {sym:<2s}  {x:12.6f} {y:12.6f} {z:12.6f}\n")
                        f.write("\n")
                self.log(f"已生成: {out_path}")
            except Exception as e:
                self.log(f"生成 {gjf_name} 失败: {e}")

        self.log("提取扫描构象任务完成。")


if __name__ == "__main__":
    app = GaussianToolGUI()
    app.mainloop()