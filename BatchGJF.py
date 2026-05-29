import argparse
import os
import re
import glob


def modify_gjf_content(input_path, output_path, mem, nprocshared, KEYWORD, CHARGE, MULT):
    """
    读取 input_path 的 .gjf 文件，返回修改后的内容列表。
    生成的 %chk 名称基于 output_path 的文件名（不含目录部分），
    这样新文件中 chk 名与新文件名一致。
    """
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 根据输出文件名生成 chk 名（不带路径）
    chk_filename = f"{output_path}.chk"

    new_lines = []
    i = 0

    # 标记是否已替换相关行
    chk_replaced = False
    mem_replaced = False
    nproc_replaced = False
    keyword_replaced = False
    charge_mult_replaced = False
    title_found = False  # 是否已处理到标题行

    while i < len(lines):
        line = lines[i]

        # ---- 1. %chk 行 ----
        if re.match(r'^%chk\s*=', line, re.IGNORECASE):
            new_lines.append(f"%chk={chk_filename}\n")
            chk_replaced = True
            i += 1
            continue

        # ---- 2. %mem 行 ----
        elif re.match(r'^%mem\s*=', line, re.IGNORECASE):
            new_lines.append(f"%mem={mem}\n")
            mem_replaced = True
            i += 1
            continue

        # ---- 3. %nprocshared (或 %nproc) 行 ----
        elif re.match(r'^%nproc(shared)?\s*=', line, re.IGNORECASE):
            new_lines.append(f"%nprocshared={nprocshared}\n")
            nproc_replaced = True
            i += 1
            continue

        # ---- 4. 关键词行 (# 开头，且尚未替换) ----
        elif re.match(r'^\s*#', line) and not keyword_replaced:
            new_lines.append(f"{KEYWORD}\n")
            keyword_replaced = True
            i += 1
            # 跳过关键词的续行（空格开头或另一个以 # 开头的行）
            while i < len(lines):
                nxt = lines[i]
                if nxt.strip() and (nxt.strip().startswith('#') or nxt[0] == ' '):
                    i += 1
                else:
                    break
            continue

        # ---- 5. 电荷与自旋多重度行（在标题之后） ----
        elif title_found and not charge_mult_replaced \
                and re.match(r'^\s*[-+]?\d+\s+[-+]?\d+\s*$', line):
            new_lines.append(f"{CHARGE} {MULT}\n")
            charge_mult_replaced = True
            i += 1
            continue

        # ---- 6. 标题行（第一个既不是 % 也不是 # 的非空行） ----
        elif not title_found and not re.match(r'^\s*%', line) \
                and not re.match(r'^\s*#', line) and line.strip():
            new_lines.append(line)  # 保留原标题
            title_found = True
            i += 1
            continue

        # ---- 其他所有行原样保留 ----
        else:
            new_lines.append(line)
            i += 1

    # ---- 如果某些配置在原文件中不存在，则在适当位置插入默认值 ----
    # 插入顺序：%chk, %mem, %nprocshared
    if not chk_replaced:
        new_lines.insert(0, f"%chk={chk_filename}\n")
    if not mem_replaced:
        # 插在 %chk 之后（如果它也被插入，此时位置 0 是 %chk）
        insert_pos = 1 if not chk_replaced else 1
        new_lines.insert(insert_pos, f"%mem={mem}\n")
    if not nproc_replaced:
        # 插在已有两个或一个 % 行之后
        insert_pos = 0
        if not chk_replaced:
            insert_pos += 1
        if not mem_replaced:
            insert_pos += 1
        new_lines.insert(insert_pos, f"%nprocshared={nprocshared}\n")

    # 如果关键词行未找到，在所有 % 行之后、标题行之前插入
    if not keyword_replaced:
        # 找到第一个不是 % 开头的行的位置
        insert_idx = 0
        for idx, ln in enumerate(new_lines):
            if not ln.startswith('%'):
                insert_idx = idx
                break
        else:
            insert_idx = len(new_lines)
        new_lines.insert(insert_idx, f"{KEYWORD}\n")

    return new_lines


def parse_log(filename):
    """
    从 Gaussian 输出文件中提取：最后一个收敛的几何构型（原子序数+坐标）。
    返回: (charge, mult, atomic_numbers, coordinates)
    coordinates 是 list of (x,y,z) 浮点数元组。
    若无法提取到有效原子信息，返回 None。
    """
    with open(f"{filename}.log", 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    atomic_numbers = []  # 存储最后找到的构型的原子序数
    coordinates = []  # 存储最后找到的构型的坐标 (x,y,z)

    re_std = re.compile(r'^\s*Standard\s+orientation\s*:', re.IGNORECASE)
    # 坐标行：Center  AtomicNum  Type  X  Y  Z
    re_coord = re.compile(
        r'^\s*(\d+)\s+(\d+)\s+(\d+)\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)'
    )

    in_std = False
    current_atomic = []
    current_coords = []

    for line in lines:

        # ---- 检测 Standard orientation 段开始 ----
        if re_std.search(line):
            in_std = True
            current_atomic = []
            current_coords = []
            continue

        if in_std:
            # 遇到分隔线（----）
            if line.strip().startswith('---'):
                if current_atomic:  # 已经有坐标 → 才是真正的段落结尾
                    atomic_numbers = current_atomic
                    coordinates = current_coords
                    in_std = False
                # 如果 current_atomic 为空，说明只是表头前面的横线，忽略，继续读取后面的坐标
                continue

            # 尝试匹配坐标行
            m = re_coord.match(line)
            if m:
                an = int(m.group(2))
                x = float(m.group(4))
                y = float(m.group(5))
                z = float(m.group(6))
                current_atomic.append(an)
                current_coords.append((x, y, z))
            # 其他行忽略，继续保持 in_std = True

    return atomic_numbers, coordinates


def write_gjf(og, name, charge, mult, mem, nprocshared, KEYWORD, atomic_numbers, coordinates):

    with open(f"{og}/{name}.gjf", 'w', encoding='utf-8') as f:
        f.write(f"%chk={name}.chk\n")
        f.write(f"%mem={mem}\n")
        f.write(f"%nprocshared={nprocshared}\n")
        f.write(f"{KEYWORD}\n")
        f.write("\n")  # 空行
        f.write(f"Generated from log file\n")  # 标题行，可自行修改
        f.write("\n")  # 空行
        f.write(f"{charge} {mult}\n")
        for an, (x, y, z) in zip(atomic_numbers, coordinates):
            f.write(f"{an:2d}   {x:12.6f}   {y:12.6f}   {z:12.6f}\n")
        f.write("\n")  # 文件末尾空行
        f.write("\n")  # 文件末尾空行


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='批量修改gjf配置信息脚本')
    parser.add_argument('--mode', default='log', help='gjf格式，log格式')
    parser.add_argument('--o', required=True, help='前缀')
    parser.add_argument('--mem', required=True, help='内存数')
    parser.add_argument('--npr', required=True, help='核心数')
    parser.add_argument('--k', required=True, help='关键词')
    parser.add_argument('--charge', required=True, help='电荷数')
    parser.add_argument('--mult', required=True, help='自旋多重度')
    args = parser.parse_args()

    OUTPUT_NAME = args.o  # 前缀
    mem = args.mem
    nprocshared = args.npr
    KEYWORD = args.k
    CHARGE = args.charge
    MULT = args.mult

    if args.mode == 'gjf':
        pattern = os.path.join("./", "*.gjf")
        gjf_files = glob.glob(pattern)
        basenames = [os.path.basename(f) for f in gjf_files]
        names = []
        for name in basenames:
            name_temp, _ = os.path.splitext(name)
            names.append(name_temp)
        try:
            os.mkdir(OUTPUT_NAME)
        except FileExistsError:
            print(f"文件夹 '{OUTPUT_NAME}' 已存在，跳过创建。")
        print(f"找到 {len(gjf_files)} 个 .gjf 文件，将生成名为 '{OUTPUT_NAME}+旧文件名' 的新文件...")
        for name in names:
            new_content = modify_gjf_content(f'{name}.gjf',
                                             f'{OUTPUT_NAME}{name}',
                                             mem, nprocshared, KEYWORD, CHARGE, MULT)
            path = os.path.join(f"./{OUTPUT_NAME}", f"{OUTPUT_NAME}{name}.gjf")
            with open(path, 'w', encoding='utf-8') as f:
                f.writelines(new_content)
            print(f"  已生成: {OUTPUT_NAME}{name}.gjf")

    elif args.mode == 'log':
        pattern = os.path.join("./", "*.log")
        log_files = glob.glob(pattern)
        basenames = [os.path.basename(f) for f in log_files]
        names = []
        for name in basenames:
            name_temp, _ = os.path.splitext(name)
            names.append(name_temp)
        try:
            os.mkdir(OUTPUT_NAME)
        except FileExistsError:
            print(f"文件夹 '{OUTPUT_NAME}' 已存在，跳过创建。")

        print(f"找到 {len(log_files)} 个 .log 文件，开始生成 .gjf ...")
        for name in names:
            atomic_numbers, coordinates = parse_log(name)
            write_gjf(OUTPUT_NAME, f'{OUTPUT_NAME}{name}',
                      CHARGE, MULT, mem, nprocshared, KEYWORD,
                      atomic_numbers, coordinates)
            print(f"  已生成: {OUTPUT_NAME}{name}.gjf")
