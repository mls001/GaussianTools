#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os
import re
import glob
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from PIL import Image, ImageTk
from openpyxl import Workbook

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


def parse_orbital_energies_advanced(log_path):
    """
    从高斯 log 文件末尾提取最后一次 SCF 收敛后的完整轨道能量。
    占据轨道和虚轨道均向上收集连续的同类型行，确保完整性和正确顺序。
    """
    with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    def find_last_title(pattern):
        for i in range(len(lines)-1, -1, -1):
            if re.search(pattern, lines[i], re.I):
                return i
        return None

    def extract_energies_from_line(line):
        nums = re.findall(r'[-+]?\d*\.?\d+(?:[DdEe][-+]?\d+)?', line)
        energies = []
        for num_str in nums:
            num_str = num_str.replace('D', 'E').replace('d', 'E')
            try:
                energies.append(float(num_str))
            except ValueError:
                continue
        return energies

    def collect_continuous_upward(start_idx, pattern):
        """从 start_idx 开始向上收集连续匹配 pattern 的行，返回按文件顺序排列的行索引列表"""
        rows = []
        i = start_idx
        while i >= 0 and re.search(pattern, lines[i], re.I):
            rows.append(i)
            i -= 1
        rows.reverse()  # 变为从上到下的顺序
        return rows

    # Alpha 占据轨道
    alpha_occ_pos = find_last_title(r'Alpha\s+occ\.?\s+eigenvalues\s*--')
    alpha_occ_energies = []
    if alpha_occ_pos is not None:
        rows = collect_continuous_upward(alpha_occ_pos, r'Alpha\s+occ\.?\s+eigenvalues\s*--')
        for idx in rows:
            alpha_occ_energies.extend(extract_energies_from_line(lines[idx]))

    # Alpha 虚轨道（同样向上收集连续行）
    alpha_virt_pos = find_last_title(r'Alpha\s+virt\.?\s+eigenvalues\s*--')
    alpha_virt_energies = []
    if alpha_virt_pos is not None:
        rows = collect_continuous_upward(alpha_virt_pos, r'Alpha\s+virt\.?\s+eigenvalues\s*--')
        for idx in rows:
            alpha_virt_energies.extend(extract_energies_from_line(lines[idx]))

    # Beta 占据轨道
    beta_occ_pos = find_last_title(r'Beta\s+occ\.?\s+eigenvalues\s*--')
    beta_occ_energies = []
    if beta_occ_pos is not None:
        rows = collect_continuous_upward(beta_occ_pos, r'Beta\s+occ\.?\s+eigenvalues\s*--')
        for idx in rows:
            beta_occ_energies.extend(extract_energies_from_line(lines[idx]))

    # Beta 虚轨道
    beta_virt_pos = find_last_title(r'Beta\s+virt\.?\s+eigenvalues\s*--')
    beta_virt_energies = []
    if beta_virt_pos is not None:
        rows = collect_continuous_upward(beta_virt_pos, r'Beta\s+virt\.?\s+eigenvalues\s*--')
        for idx in rows:
            beta_virt_energies.extend(extract_energies_from_line(lines[idx]))

    # RHF 后备（如果没有 Alpha 但有RHF）
    if not alpha_occ_energies:
        rhf_occ_pos = find_last_title(r'Occupied\s*\(RHF\)\s*--')
        if rhf_occ_pos is not None:
            rows = collect_continuous_upward(rhf_occ_pos, r'Occupied\s*\(RHF\)\s*--')
            for idx in rows:
                alpha_occ_energies.extend(extract_energies_from_line(lines[idx]))
    if not alpha_virt_energies:
        rhf_virt_pos = find_last_title(r'Virtual\s*\(RHF\)\s*--')
        if rhf_virt_pos is not None:
            rows = collect_continuous_upward(rhf_virt_pos, r'Virtual\s*\(RHF\)\s*--')
            for idx in rows:
                alpha_virt_energies.extend(extract_energies_from_line(lines[idx]))

    # 分配轨道序号（占据轨道从1开始，虚轨道接续）
    alpha_occ = [(i+1, eng) for i, eng in enumerate(alpha_occ_energies)]
    alpha_virt = [(i+1+len(alpha_occ_energies), eng) for i, eng in enumerate(alpha_virt_energies)]
    beta_occ = [(i+1, eng) for i, eng in enumerate(beta_occ_energies)]
    beta_virt = [(i+1+len(beta_occ_energies), eng) for i, eng in enumerate(beta_virt_energies)]

    # HOMO / LUMO
    homo_alpha = alpha_occ[-1][0] if alpha_occ else None
    lumo_alpha = alpha_virt[0][0] if alpha_virt else None
    homo_beta = beta_occ[-1][0] if beta_occ else None
    lumo_beta = beta_virt[0][0] if beta_virt else None

    return {
        'filename': os.path.basename(log_path),
        'alpha_occ': alpha_occ,
        'alpha_virt': alpha_virt,
        'beta_occ': beta_occ,
        'beta_virt': beta_virt,
        'homo_alpha': homo_alpha,
        'lumo_alpha': lumo_alpha,
        'homo_beta': homo_beta,
        'lumo_beta': lumo_beta,
    }


def resource_path(relative_path):
    """获取资源的绝对路径，兼容开发环境和 PyInstaller 打包后的环境"""
    try:
        base_path = sys._MEIPASS  # PyInstaller 创建的临时文件夹
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


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
    route_str = '\n'.join(route_lines) if route_lines else '#p b3lpy/6-31G(d,p)'
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


def parse_td_data(log_path):
    """
    解析高斯log文件中的TD激发态信息，并获取轨道能量。
    返回字典，包含轨道能量映射和激发态列表。
    """
    # 首先获取轨道能量（复用现有函数）
    orbital_data = parse_orbital_energies_advanced(log_path)
    # 构建全局轨道序号 -> 能量(Ha) 的映射
    orb_energy_map = {}
    # Alpha 占据 (idx, eng)
    for idx, eng in orbital_data.get('alpha_occ', []):
        orb_energy_map[idx] = eng
    # Alpha 虚
    for idx, eng in orbital_data.get('alpha_virt', []):
        orb_energy_map[idx] = eng
    # Beta 占据
    for idx, eng in orbital_data.get('beta_occ', []):
        orb_energy_map[idx] = eng
    # Beta 虚
    for idx, eng in orbital_data.get('beta_virt', []):
        orb_energy_map[idx] = eng

    with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # 正则表达式匹配激发态块
    state_pattern = re.compile(
        r'^\s*Excited\s+State\s+(\d+):\s+(\S+)\s+([\d\.]+)\s+eV\s+([\d\.]+)\s+nm\s+f=([\d\.Ee+-]+)',
        re.M
    )
    # 匹配跃迁行: "     40 -> 41       0.66499" 或 "     39 -> 41      -0.15291"
    trans_pattern = re.compile(r'^\s*(\d+)\s*->\s*(\d+)\s+([-+]?[\d\.Ee+-]+)')

    states = []
    # 按激发态分割内容：每个激发态从"Excited State"开始，直到下一个"Excited State"或文件结束
    blocks = re.split(r'\n(?=\s*Excited\s+State)', content)
    for block in blocks[1:]:  # 第一个块是空或无用
        lines = block.splitlines()
        if not lines:
            continue
        first_line = lines[0]
        m = state_pattern.match(first_line)
        if not m:
            continue
        state_num = int(m.group(1))
        mult_type = m.group(2)      # 如 Singlet-A, Triplet-A
        energy_eV = float(m.group(3))
        wavelength_nm = float(m.group(4))
        osc_strength = float(m.group(5))

        # 解析跃迁
        transitions = []
        for line in lines[1:]:
            t = trans_pattern.match(line)
            if t:
                from_orb = int(t.group(1))
                to_orb = int(t.group(2))
                coeff = float(t.group(3))
                percent = (coeff ** 2) * 100 * 2
                from_energy = orb_energy_map.get(from_orb, None)
                to_energy = orb_energy_map.get(to_orb, None)
                delta_energy = None
                if from_energy is not None and to_energy is not None:
                    delta_energy = to_energy - from_energy
                transitions.append({
                    'from': from_orb,
                    'to': to_orb,
                    'coeff': coeff,
                    'percent': percent,
                    'from_energy': from_energy,
                    'to_energy': to_energy,
                    'delta_energy': delta_energy,
                })
        states.append({
            'state_num': state_num,
            'mult_type': mult_type,
            'energy_eV': energy_eV,
            'wavelength_nm': wavelength_nm,
            'osc_strength': osc_strength,
            'transitions': transitions,
        })
    return {'orbital_map': orb_energy_map, 'states': states}


# ----------------------------------------------------------------------
# 图形界面
# ----------------------------------------------------------------------

class GaussianToolGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("mls V0.0.2")
        self.geometry("850x700")
        self.resizable(True, True)
        mode_frame = ttk.LabelFrame(self, text="操作模式", padding=5)
        mode_frame.pack(fill=tk.X, padx=5, pady=5)
        # 在 __init__ 方法中，创建窗口后
        icon_path = resource_path("md.ico")
        try:
            self.iconbitmap(icon_path)  # 仅 Windows 支持 .ico
        except Exception as e:
            print(f"图标加载失败: {e}")
        self.mode_var = tk.StringVar(value="modify_gjf")
        modes = [
            ("修改 GJF 参数", "modify_gjf"),
            ("从 LOG 生成 GJF", "log_to_gjf"),
            ("提取扫描构象", "extract_scan"),
            ("批量提取轨道能量", "batch_orbital"),
            ("提取TD信息", "extract_td")
        ]
        for text, value in modes:
            ttk.Radiobutton(mode_frame, text=text, variable=self.mode_var,
                            value=value, command=self.on_mode_change).pack(side=tk.LEFT, padx=10)

        self.param_frame = None
        log_frame = ttk.LabelFrame(self, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 底部按钮框架
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        # 先添加清空和运行按钮（从右向左排列）
        ttk.Button(btn_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=5)
        self.run_btn = ttk.Button(btn_frame, text="运行", command=self.start_task)
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # 在右下角添加自定义图片（放在最右边）
        self.add_logo(btn_frame)

        self.current_widgets = []
        self.on_mode_change()

    def add_logo(self, parent_frame):
        # 图片路径：可修改为你的图片文件名，支持相对路径
        logo_path = os.path.join(os.path.dirname(__file__), "md.jpg")
        # 如果当前目录没有，尝试 exe 同级目录（打包后适用）
        if not os.path.exists(logo_path):
            logo_path = "md.jpg"
        if not os.path.exists(logo_path):
            return

        try:
            img = Image.open(logo_path)
            # 缩放图片高度为 32 像素，宽度按比例
            img.thumbnail((64, 64), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)  # 保持引用
            logo_label = ttk.Label(parent_frame, image=self.logo_img)
            logo_label.pack(side=tk.RIGHT, padx=5)
        except Exception as e:
            print(f"加载图片失败: {e}")

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
        elif mode == "batch_orbital":
            self.create_batch_orbital_widgets()
        elif mode == "extract_td":
            self.create_extract_td_widgets()
        else:
            self.create_extract_scan_widgets()

    # ---------- 辅助方法：添加上拉菜单资源预设 ----------
    def add_resource_preset_ui(self, parent, mem_var, nproc_var):
        row = len(parent.grid_slaves())
        ttk.Label(parent, text="资源预设:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)

        self.current_preset = tk.StringVar(value="students/zstoffice")
        preset_btn = ttk.Button(parent, textvariable=self.current_preset, width=20)
        preset_btn.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)

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
            menu.post(x, y - menu_height)

        preset_btn.bind("<Button-1>", show_menu)
        self.current_widgets.append(preset_btn)

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

        self.keyword_var = tk.StringVar(value="#p opt b3lpy/6-31G(d,p)")
        ttk.Label(self.param_frame, text="关键词行:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        kw_entry = ttk.Entry(self.param_frame, textvariable=self.keyword_var, width=60)
        kw_entry.grid(row=row, column=1, padx=5, pady=2, sticky=tk.W)
        self.current_widgets.extend([kw_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

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

        self.mem_var = tk.StringVar(value="20GB")
        ttk.Label(self.param_frame, text="%mem:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        mem_entry = ttk.Entry(self.param_frame, textvariable=self.mem_var, width=20)
        mem_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([mem_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        self.nproc_var = tk.StringVar(value="8")
        ttk.Label(self.param_frame, text="%nprocshared:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        nproc_entry = ttk.Entry(self.param_frame, textvariable=self.nproc_var, width=20)
        nproc_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([nproc_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

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

        self.keyword_var = tk.StringVar(value="#p opt b3lpy/6-31G(d,p)")
        ttk.Label(self.param_frame, text="关键词行:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        kw_entry = ttk.Entry(self.param_frame, textvariable=self.keyword_var, width=60)
        kw_entry.grid(row=row, column=1, padx=5, pady=2, sticky=tk.W)
        self.current_widgets.extend([kw_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

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

        self.mem_var = tk.StringVar(value="20GB")
        ttk.Label(self.param_frame, text="%mem:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        mem_entry = ttk.Entry(self.param_frame, textvariable=self.mem_var, width=20)
        mem_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([mem_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        self.nproc_var = tk.StringVar(value="8")
        ttk.Label(self.param_frame, text="%nprocshared:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        nproc_entry = ttk.Entry(self.param_frame, textvariable=self.nproc_var, width=20)
        nproc_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([nproc_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

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

        self.mem_scan_var = tk.StringVar(value="20GB")
        ttk.Label(self.param_frame, text="%mem:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        mem_entry = ttk.Entry(self.param_frame, textvariable=self.mem_scan_var, width=20)
        mem_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([mem_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

        self.nproc_scan_var = tk.StringVar(value="8")
        ttk.Label(self.param_frame, text="%nprocshared:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
        nproc_entry = ttk.Entry(self.param_frame, textvariable=self.nproc_scan_var, width=20)
        nproc_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=2)
        self.current_widgets.extend([nproc_entry, self.param_frame.grid_slaves(row=row, column=0)[0]])
        row += 1

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
            elif mode == "batch_orbital":
                # 注意：start_batch_orbital 会自己处理，不需要额外线程？但为了避免界面卡顿，建议也放到线程中
                # 但 Treeview 更新必须在主线程，所以实际上需要自定义线程+主线程回调。为了简化，不放入 run_task，
                # 而是直接调用 start_batch_orbital（已在按钮事件中）。因此 run_task 中可以忽略。
                pass
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

    def create_batch_orbital_widgets(self):
        row = 0

        self.batch_orbital_folder_var = tk.StringVar()
        ttk.Label(self.param_frame, text="选择包含 LOG 文件的文件夹:").grid(row=row, column=0, sticky=tk.W, padx=5,
                                                                            pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.batch_orbital_folder_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览",
                   command=lambda: self.select_folder(self.batch_orbital_folder_var)).grid(row=row, column=2, padx=5,
                                                                                           pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        btn_inner = ttk.Frame(self.param_frame)
        btn_inner.grid(row=row, column=0, columnspan=3, pady=5)
        ttk.Button(btn_inner, text="提取并显示", command=self.start_batch_orbital).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_inner, text="导出CSV（合并）", command=self.export_batch_orbital_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_inner, text="导出Excel（多sheet）", command=self.export_batch_orbital_excel).pack(side=tk.LEFT,
                                                                                                       padx=5)
        self.current_widgets.append(btn_inner)
        row += 1

        # 创建 Notebook 容器
        self.batch_notebook = ttk.Notebook(self.param_frame)
        self.batch_notebook.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=5)
        self.param_frame.grid_rowconfigure(row, weight=1)
        self.param_frame.grid_columnconfigure(1, weight=1)
        self.current_widgets.append(self.batch_notebook)
        row += 1

        # 添加能隙计算窗格
        gap_frame = ttk.LabelFrame(self.param_frame, text="计算能隙（当前文件）", padding=5)
        gap_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=5)
        self.param_frame.grid_columnconfigure(0, weight=1)
        self.current_widgets.append(gap_frame)

        ttk.Label(gap_frame, text="轨道序号 A:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.gap_idx1 = tk.StringVar()
        entry1 = ttk.Entry(gap_frame, textvariable=self.gap_idx1, width=10)
        entry1.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(gap_frame, text="轨道序号 B:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.gap_idx2 = tk.StringVar()
        entry2 = ttk.Entry(gap_frame, textvariable=self.gap_idx2, width=10)
        entry2.grid(row=0, column=3, padx=5, pady=2)

        ttk.Button(gap_frame, text="计算能隙", command=self.calc_gap).grid(row=0, column=4, padx=10, pady=2)

        self.gap_result_var = tk.StringVar(value="")
        result_label = ttk.Label(gap_frame, textvariable=self.gap_result_var, foreground="blue")
        result_label.grid(row=0, column=5, padx=5, pady=2, sticky=tk.W)

        # 存储所有文件的完整轨道数据，用于导出
        self.batch_orbital_data = []  # 每个元素为 {'filename': , 'tracks': [(spin, type, idx, eng_ha), ...]}

    def calc_gap(self):
        """计算当前选中的文件中两个指定轨道序号之间的能隙（eV 和 Hartree）"""
        if not self.batch_orbital_data:
            self.gap_result_var.set("请先提取数据")
            return
        current_tab = self.batch_notebook.select()
        if not current_tab:
            self.gap_result_var.set("未选中任何文件")
            return
        tab_text = self.batch_notebook.tab(current_tab, "text")
        target_data = None
        for item in self.batch_orbital_data:
            if item['filename'].replace('.log', '') == tab_text:
                target_data = item
                break
        if not target_data:
            self.gap_result_var.set(f"未找到文件 {tab_text} 的数据")
            return

        try:
            idx1 = int(self.gap_idx1.get().strip())
            idx2 = int(self.gap_idx2.get().strip())
        except ValueError:
            self.gap_result_var.set("请输入有效的整数序号")
            return

        eng1 = None
        eng2 = None
        for spin, typ, idx, eng_ha in target_data['tracks']:
            if idx == idx1:
                eng1 = eng_ha
            if idx == idx2:
                eng2 = eng_ha
        if eng1 is None:
            self.gap_result_var.set(f"轨道序号 {idx1} 不存在")
            return
        if eng2 is None:
            self.gap_result_var.set(f"轨道序号 {idx2} 不存在")
            return

        gap_ha = abs(eng1 - eng2)
        gap_ev = gap_ha * 27.211386245988
        self.gap_result_var.set(f"能隙 = {gap_ev:.4f} eV = {gap_ha:.6f} Hartree (|{idx1} - {idx2}|)")

    def start_batch_orbital(self):
        folder = self.batch_orbital_folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self.log("请选择有效的文件夹")
            return

        log_files = glob.glob(os.path.join(folder, "*.log"))
        if not log_files:
            self.log("未找到 .log 文件")
            return

        self.log(f"开始处理 {len(log_files)} 个 LOG 文件...")
        # 清空 Notebook 和数据存储
        for tab in self.batch_notebook.tabs():
            self.batch_notebook.forget(tab)
        self.batch_orbital_data.clear()

        for logf in log_files:
            try:
                data = parse_orbital_energies_advanced(logf)
                filename = data['filename']
                # 收集所有轨道数据，用于导出和显示
                tracks = []  # (spin, type, idx, eng_ha)
                # Alpha 占据
                for idx, eng in data['alpha_occ']:
                    tracks.append(('Alpha', 'Occ', idx, eng))
                # Alpha 虚
                for idx, eng in data['alpha_virt']:
                    tracks.append(('Alpha', 'Vir', idx, eng))
                # Beta 占据
                for idx, eng in data['beta_occ']:
                    tracks.append(('Beta', 'Occ', idx, eng))
                # Beta 虚
                for idx, eng in data['beta_virt']:
                    tracks.append(('Beta', 'Vir', idx, eng))
                # 按轨道序号排序
                tracks.sort(key=lambda x: x[2])
                self.batch_orbital_data.append({'filename': filename, 'tracks': tracks})

                # 创建 Tab
                tab_frame = ttk.Frame(self.batch_notebook)
                self.batch_notebook.add(tab_frame, text=filename.replace('.log', ''))

                # 在 Tab 内创建 Treeview 显示所有轨道
                tree_frame = ttk.Frame(tab_frame)
                tree_frame.pack(fill=tk.BOTH, expand=True)
                scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
                scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
                tree = ttk.Treeview(tree_frame,
                                    columns=('spin', 'type', 'index', 'energy_ha', 'energy_ev'),
                                    show='headings',
                                    yscrollcommand=scroll_y.set,
                                    xscrollcommand=scroll_x.set)
                scroll_y.config(command=tree.yview)
                scroll_x.config(command=tree.xview)

                tree.heading('spin', text='自旋')
                tree.heading('type', text='类型')
                tree.heading('index', text='轨道序号')
                tree.heading('energy_ha', text='能量 (Ha)')
                tree.heading('energy_ev', text='能量 (eV)')

                tree.column('spin', width=60)
                tree.column('type', width=80)
                tree.column('index', width=80)
                tree.column('energy_ha', width=120)
                tree.column('energy_ev', width=120)

                tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
                scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

                # 填充轨道数据，并高亮 HOMO/LUMO
                homo_alpha = data.get('homo_alpha')
                lumo_alpha = data.get('lumo_alpha')
                homo_beta = data.get('homo_beta')
                lumo_beta = data.get('lumo_beta')

                tree.tag_configure('homo', background='lightgreen')
                tree.tag_configure('lumo', background='lightblue')

                homo_item = None  # 用于记录 HOMO 行的 Treeview item ID
                for spin, typ, idx, eng_ha in tracks:
                    eng_ev = eng_ha * 27.211386245988
                    tag = ''
                    if spin == 'Alpha' and typ == 'Occ' and idx == homo_alpha:
                        tag = 'homo'
                    elif spin == 'Alpha' and typ == 'Vir' and idx == lumo_alpha:
                        tag = 'lumo'
                    elif spin == 'Beta' and typ == 'Occ' and idx == homo_beta:
                        tag = 'homo'
                    elif spin == 'Beta' and typ == 'Vir' and idx == lumo_beta:
                        tag = 'lumo'
                    item = tree.insert('', tk.END, values=(spin, typ, idx, f"{eng_ha:.6f}", f"{eng_ev:.4f}"),
                                       tags=(tag,))
                    if tag == 'homo':
                        homo_item = item  # 记录最后一个 HOMO 行（通常只有一行，因为 HOMO 唯一）

                # 自动滚动到 HOMO 行
                if homo_item:
                    tree.see(homo_item)
                    # 可选：同时选中该行，方便定位
                    tree.selection_set(homo_item)
                    tree.focus(homo_item)

                self.log(f"已处理: {filename} (共 {len(tracks)} 条轨道)")
            except Exception as e:
                self.log(f"处理 {logf} 时出错: {e}")
        self.log("提取完成，请在选项卡中查看各文件的轨道能量（已自动定位到 HOMO 行）")

    def export_batch_orbital_csv(self):
        if not self.batch_orbital_data:
            self.log("没有数据可导出，请先执行提取")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not save_path:
            return
        try:
            with open(save_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
                import csv
                writer = csv.writer(csvfile)
                writer.writerow(["文件", "自旋", "类型", "轨道序号", "能量(Ha)", "能量(eV)"])
                for item in self.batch_orbital_data:
                    filename = item['filename']
                    for spin, typ, idx, eng_ha in item['tracks']:
                        eng_ev = eng_ha * 27.211386245988
                        writer.writerow([filename, spin, typ, idx, f"{eng_ha:.6f}", f"{eng_ev:.4f}"])
            self.log(f"CSV 文件已保存至: {save_path}")
        except Exception as e:
            self.log(f"导出失败: {e}")

    def _get_homo_idx(self, filename, spin):
        """根据文件名和自旋获取 HOMO 序号"""
        for data in self.batch_orbital_data:
            if data['filename'] == filename:
                if spin == "Alpha":
                    return data['homo_alpha']
                else:
                    return data['homo_beta']
        return None

    def _get_lumo_idx(self, filename, spin):
        for data in self.batch_orbital_data:
            if data['filename'] == filename:
                if spin == "Alpha":
                    return data['lumo_alpha']
                else:
                    return data['lumo_beta']
        return None

    def create_extract_td_widgets(self):
        """TD信息批量提取模式的界面"""
        row = 0
        self.td_folder_var = tk.StringVar()
        ttk.Label(self.param_frame, text="选择包含 LOG 文件的文件夹:").grid(row=row, column=0, sticky=tk.W, padx=5,
                                                                            pady=2)
        entry = ttk.Entry(self.param_frame, textvariable=self.td_folder_var, width=50)
        entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Button(self.param_frame, text="浏览",
                   command=lambda: self.select_folder(self.td_folder_var)).grid(row=row, column=2, padx=5, pady=2)
        self.current_widgets.extend([entry, self.param_frame.grid_slaves(row=row, column=0)[0],
                                     self.param_frame.grid_slaves(row=row, column=2)[0]])
        row += 1

        btn_frame = ttk.Frame(self.param_frame)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=5)
        ttk.Button(btn_frame, text="批量解析TD", command=self.start_batch_td).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导出CSV（合并）", command=self.export_batch_td_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导出Excel（多sheet）", command=self.export_batch_td_excel).pack(side=tk.LEFT, padx=5)
        self.current_widgets.append(btn_frame)
        row += 1

        # 创建 Notebook 容器用于显示各文件的TD结果
        self.td_notebook = ttk.Notebook(self.param_frame)
        self.td_notebook.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=5)
        self.param_frame.grid_rowconfigure(row, weight=1)
        self.param_frame.grid_columnconfigure(1, weight=1)
        self.current_widgets.append(self.td_notebook)

        # 存储所有文件的TD数据，用于导出
        self.batch_td_data = []  # 每个元素: {'filename': , 'states': [...]}

    def start_batch_td(self):
        folder = self.td_folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self.log("请选择有效的文件夹")
            return

        log_files = glob.glob(os.path.join(folder, "*.log"))
        if not log_files:
            self.log("未找到 .log 文件")
            return

        self.log(f"开始批量解析TD信息，共 {len(log_files)} 个文件...")
        # 清空Notebook和数据
        for tab in self.td_notebook.tabs():
            self.td_notebook.forget(tab)
        self.batch_td_data.clear()

        # 使用线程避免界面卡顿
        threading.Thread(target=self._batch_td_worker, args=(log_files,), daemon=True).start()

    def _batch_td_worker(self, log_files):
        """后台批量解析TD信息"""
        for logf in log_files:
            try:
                self.log(f"正在解析: {os.path.basename(logf)}")
                td_data = parse_td_data(logf)  # 注意：parse_td_data 需要能处理没有TD输出的文件，返回空states
                if not td_data['states']:
                    self.log(f"警告: {os.path.basename(logf)} 未发现TD激发态信息")
                    continue
                self.batch_td_data.append({
                    'filename': os.path.basename(logf),
                    'states': td_data['states'],
                    'orbital_map': td_data.get('orbital_map', {})
                })
                # 在主线程中更新GUI
                self.after(0, self._add_td_tab, os.path.basename(logf), td_data['states'])
            except Exception as e:
                self.log(f"解析 {logf} 失败: {e}")
                import traceback
                self.log(traceback.format_exc())
        self.after(0, lambda: self.log("批量TD解析完成"))

    def _add_td_tab(self, filename, states):
        """为每个文件创建一个选项卡显示TD结果"""
        tab_frame = ttk.Frame(self.td_notebook)
        self.td_notebook.add(tab_frame, text=filename.replace('.log', ''))

        # 创建文本框和滚动条
        text_widget = scrolledtext.ScrolledText(tab_frame, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True)

        # 格式化显示
        text_widget.insert(tk.END, "=" * 80 + "\n")
        text_widget.insert(tk.END, f"文件: {filename}\n")
        text_widget.insert(tk.END, f"共 {len(states)} 个激发态\n")
        text_widget.insert(tk.END, "=" * 80 + "\n\n")

        for state in states:
            text_widget.insert(tk.END, f"激发态 {state['state_num']:3d}: {state['mult_type']}\n")
            text_widget.insert(tk.END,
                               f"  能量: {state['energy_eV']:.4f} eV, 波长: {state['wavelength_nm']:.2f} nm, 振子强度: {state['osc_strength']:.6f}\n")
            text_widget.insert(tk.END, f"  主要跃迁贡献 (占比 > 2%):\n")
            sorted_trans = sorted(state['transitions'], key=lambda x: x['percent'], reverse=True)
            for trans in sorted_trans:
                if trans['percent'] < 2.0:
                    continue
                text_widget.insert(tk.END,
                                   f"    {trans['from']:3d} → {trans['to']:3d} : 系数 {trans['coeff']:8.5f} (占比 {trans['percent']:.2f}%)\n")
                if trans['from_energy'] is not None and trans['to_energy'] is not None:
                    delta = trans['delta_energy']
                    text_widget.insert(tk.END,
                                       f"        轨道能量: {trans['from_energy']:8.6f} Ha → {trans['to_energy']:8.6f} Ha, 能量差: {delta:8.6f} Ha ({delta * 27.211386:8.4f} eV)\n")
            text_widget.insert(tk.END, "\n")
        text_widget.see("1.0")  # 滚动到第一行

    def export_batch_td_csv(self):
        if not self.batch_td_data:
            self.log("没有数据可导出，请先执行批量解析")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not save_path:
            return
        try:
            with open(save_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
                import csv
                writer = csv.writer(csvfile)
                writer.writerow(["文件", "激发态序号", "多重度", "能量(eV)", "波长(nm)", "振子强度",
                                 "起始轨道", "目标轨道", "系数", "占比(%)",
                                 "起始轨道能量(Ha)", "目标轨道能量(Ha)", "轨道能量差(Ha)", "轨道能量差(eV)"])
                for item in self.batch_td_data:
                    filename = item['filename']
                    for state in item['states']:
                        for trans in state['transitions']:
                            delta_ev = trans['delta_energy'] * 27.211386 if trans['delta_energy'] is not None else ""
                            writer.writerow([
                                filename, state['state_num'], state['mult_type'],
                                f"{state['energy_eV']:.6f}", f"{state['wavelength_nm']:.2f}",
                                f"{state['osc_strength']:.8f}",
                                trans['from'], trans['to'], f"{trans['coeff']:.8f}", f"{trans['percent']:.4f}",
                                f"{trans['from_energy']:.8f}" if trans['from_energy'] is not None else "",
                                f"{trans['to_energy']:.8f}" if trans['to_energy'] is not None else "",
                                f"{trans['delta_energy']:.8f}" if trans['delta_energy'] is not None else "",
                                f"{delta_ev:.6f}" if delta_ev != "" else ""
                            ])
            self.log(f"合并CSV已保存至: {save_path}")
        except Exception as e:
            self.log(f"导出失败: {e}")

    def run_extract_td(self):
        log_file = self.td_file_var.get().strip()
        if not log_file or not os.path.exists(log_file):
            self.log("请选择有效的LOG文件")
            return
        self.log(f"开始解析TD信息: {log_file}")
        try:
            td_data = parse_td_data(log_file)
            self.current_td_data = td_data
            self.display_td_results(td_data)
            self.log("解析完成")
        except Exception as e:
            self.log(f"解析失败: {str(e)}")
            import traceback
            self.log(traceback.format_exc())

    def display_td_results(self, td_data):
        """在文本区域中格式化显示TD结果"""
        self.td_text.delete(1.0, tk.END)
        orb_map = td_data['orbital_map']
        states = td_data['states']
        self.td_text.insert(tk.END, "=" * 80 + "\n")
        self.td_text.insert(tk.END, f"轨道能量信息 (Ha):\n")
        # 显示前10个轨道作为示例，可以全部显示但会很长，这里简要显示
        sorted_orbs = sorted(orb_map.keys())
        for idx in sorted_orbs[:20]:
            self.td_text.insert(tk.END, f"  轨道 {idx:3d}: {orb_map[idx]:10.6f} Ha\n")
        if len(sorted_orbs) > 20:
            self.td_text.insert(tk.END, f"  ... (共{len(sorted_orbs)}个轨道)\n")
        self.td_text.insert(tk.END, "=" * 80 + "\n\n")

        for state in states:
            self.td_text.insert(tk.END, f"激发态 {state['state_num']:3d}: {state['mult_type']}\n")
            self.td_text.insert(tk.END,
                                f"  能量: {state['energy_eV']:.4f} eV, 波长: {state['wavelength_nm']:.2f} nm, 振子强度: {state['osc_strength']:.6f}\n")
            self.td_text.insert(tk.END, f"  主要跃迁贡献 (占比 > 2%):\n")
            # 按占比排序取前5
            sorted_trans = sorted(state['transitions'], key=lambda x: x['percent'], reverse=True)
            for trans in sorted_trans:
                if trans['percent'] < 2.0:
                    continue
                from_eng = trans['from_energy']
                to_eng = trans['to_energy']
                delta = trans['delta_energy']
                self.td_text.insert(tk.END,
                                    f"    {trans['from']:3d} → {trans['to']:3d} : 系数 {trans['coeff']:8.5f} (占比 {trans['percent']:.2f}%)\n")
                if from_eng is not None and to_eng is not None:
                    self.td_text.insert(tk.END,
                                        f"        轨道能量: {from_eng:8.6f} Ha → {to_eng:8.6f} Ha, 能量差: {delta:8.6f} Ha ({delta * 27.211386:8.4f} eV)\n")
            self.td_text.insert(tk.END, "\n")
        self.td_text.see(tk.END)

    def export_to_excel(self, data_list, sheet_name_func, headers_func, rows_func, default_filename="output.xlsx"):
        """将多个数据集导出到同一个Excel文件的不同sheet"""
        if not data_list:
            self.log("没有数据可导出")
            return
        from tkinter import filedialog
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return
        try:
            wb = Workbook()
            # 删除默认创建的空白sheet
            default_sheet = wb.active
            wb.remove(default_sheet)
            for data in data_list:
                sheet_name = sheet_name_func(data)
                # 处理非法字符并限制长度
                invalid_chars = r'[]:*?/\\'
                for ch in invalid_chars:
                    sheet_name = sheet_name.replace(ch, '_')
                sheet_name = sheet_name[:31]
                ws = wb.create_sheet(title=sheet_name)
                headers = headers_func()
                ws.append(headers)
                rows = rows_func(data)
                for row in rows:
                    ws.append(row)
            wb.save(save_path)
            self.log(f"Excel文件已保存至: {save_path} (包含 {len(data_list)} 个sheet)")
        except Exception as e:
            self.log(f"导出Excel失败: {e}")

    def export_batch_orbital_excel(self):
        if not self.batch_orbital_data:
            self.log("没有数据可导出，请先执行提取")
            return

        def sheet_name_func(data):
            return data['filename'].replace('.log', '')

        def headers_func():
            return ["自旋", "类型", "轨道序号", "能量(Ha)", "能量(eV)"]

        def rows_func(data):
            rows = []
            for spin, typ, idx, eng_ha in data['tracks']:
                eng_ev = eng_ha * 27.211386245988
                rows.append([spin, typ, idx, f"{eng_ha:.6f}", f"{eng_ev:.4f}"])
            return rows

        self.export_to_excel(self.batch_orbital_data, sheet_name_func, headers_func, rows_func, "orbital_energies.xlsx")

    def export_batch_td_excel(self):
        if not self.batch_td_data:
            self.log("没有数据可导出，请先执行批量解析")
            return

        def sheet_name_func(data):
            return data['filename'].replace('.log', '')

        def headers_func():
            return ["激发态序号", "多重度", "能量(eV)", "波长(nm)", "振子强度",
                    "起始轨道", "目标轨道", "系数", "占比(%)",
                    "起始轨道能量(Ha)", "目标轨道能量(Ha)", "轨道能量差(Ha)", "轨道能量差(eV)"]

        def rows_func(data):
            rows = []
            for state in data['states']:
                for trans in state['transitions']:
                    delta_ev = trans['delta_energy'] * 27.211386 if trans['delta_energy'] is not None else ""
                    rows.append([
                        state['state_num'], state['mult_type'],
                        f"{state['energy_eV']:.6f}", f"{state['wavelength_nm']:.2f}", f"{state['osc_strength']:.8f}",
                        trans['from'], trans['to'], f"{trans['coeff']:.8f}", f"{trans['percent']:.4f}",
                        f"{trans['from_energy']:.8f}" if trans['from_energy'] is not None else "",
                        f"{trans['to_energy']:.8f}" if trans['to_energy'] is not None else "",
                        f"{trans['delta_energy']:.8f}" if trans['delta_energy'] is not None else "",
                        f"{delta_ev:.6f}" if delta_ev != "" else ""
                    ])
            return rows

        self.export_to_excel(self.batch_td_data, sheet_name_func, headers_func, rows_func, "td_results.xlsx")


if __name__ == "__main__":
    app = GaussianToolGUI()
    app.mainloop()
