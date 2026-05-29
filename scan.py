#!/usr/bin/env python3

import re
import os
import sys

ATOMIC_NUMBER_TO_SYMBOL = {
    1: 'H', 3: 'Li', 4: 'Be', 5: 'B', 6: 'C', 7: 'N', 8: 'O', 9: 'F',
    14: 'Si', 15: 'P', 16: 'S', 17: 'Cl', 34: 'Se', 35: 'Br', 53: 'I'
}


def read_output(filename):
    with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
        return f.readlines()


def extract_header_info(lines):
    route_lines = []
    title_lines = []
    charge = 0
    multiplicity = 1

    # Charge = , Multiplicity =
    for line in lines:
        if 'Charge =' in line and 'Multiplicity =' in line:
            m = re.search(r'Charge\s*=\s*(-?\d+)\s+Multiplicity\s*=\s*(\d+)', line)
            if m:
                charge = int(m.group(1))
                multiplicity = int(m.group(2))
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
    return route_str, title_str, charge, multiplicity


def parse_standard_orientation(lines, start):
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
                    steps.append((current_scan_point,converged_atoms, converged_coords))
                converged_atoms = None
                converged_coords = None
                last_std_orient_atoms = None
                last_std_orient_coords = None
            current_scan_point = new_scan_point

        if 'Standard orientation:' in line:
            atoms, coords = parse_standard_orientation(lines, i)
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


def write_gjf(filename, route, title, charge, mult, step_label, step_id, atomic_numbers, coordinates):
    symbols = [ATOMIC_NUMBER_TO_SYMBOL.get(z, f'X{z}') for z in atomic_numbers]
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(route + '\n\n')
        f.write(f'{title} {step_label} {step_id}\n\n')
        f.write(f'{charge} {mult}\n')
        for sym, (x, y, z) in zip(symbols, coordinates):
            f.write(f' {sym:<2s}  {x:12.6f} {y:12.6f} {z:12.6f}\n')
        f.write('\n')


def main():
    if len(sys.argv) < 2:
        print("用法: python extract_scan_structures.py <Gaussian输出文件>")
        sys.exit(1)

    out_file = sys.argv[1]
    if not os.path.exists(out_file):
        print(f"错误: 文件 {out_file} 不存在")
        sys.exit(1)

    lines = read_output(out_file)
    route, title, charge, mult = extract_header_info(lines)

    steps = extract_modredundant_scan_steps(lines)
    step_label = "ScanPoint"

    base_name = os.path.splitext(os.path.basename(out_file))[0]
    for step_id, atomic_numbers, coords in steps:
        gjf_name = f'{base_name}_{step_label}{step_id}.gjf'
        write_gjf(gjf_name, route, title, charge, mult, step_label, step_id, atomic_numbers, coords)
        print(f'已生成: {gjf_name} ({step_label} {step_id})')

    print(f'共导出 {len(steps)} 个构象。')


if __name__ == '__main__':
    main()
