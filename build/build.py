#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能美高择校系统 · 自动构建脚本
触发条件：data/学校信息表.xlsx 有更新时，由 GitHub Actions 自动运行
流程：读取 Excel → 清洗数据 → 合并坐标库（新学校用州中心近似兜底）→ 注入模板 → 输出 index.html
"""
import json, os, sys
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH   = os.path.join(ROOT, 'data', '学校信息表.xlsx')
TEMPLATE_PATH= os.path.join(ROOT, 'build', 'template.html')
COORDS_PATH  = os.path.join(ROOT, 'build', 'school_coords.json')
OUTPUT_PATH  = os.path.join(ROOT, 'index.html')
NAME_KEY     = '学校名称(School Name）'
STATE_KEY    = '州 State'

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from state_centers import STATE_CENTERS


def load_excel_records(path):
    """读取 Excel，清洗 \xa0，转换为 dict 列表（与 Claude 手动流程完全一致）"""
    df = pd.read_excel(path)
    df2 = df[df[NAME_KEY].notna()].copy()
    exclude_cols = {'data_id'}
    cols = [c for c in df.columns if c not in exclude_cols]

    records = []
    for _, row in df2.iterrows():
        r = {}
        for col in cols:
            val = row[col]
            if pd.isna(val):
                r[col] = ''
            elif isinstance(val, float) and val == int(val):
                r[col] = int(val)
            else:
                r[col] = str(val).replace('\xa0', ' ').strip()
        r[NAME_KEY] = r[NAME_KEY].replace('\xa0', ' ').strip()
        records.append(r)
    return records


def merge_coords(records, coords_path):
    """合并坐标库：已有精确坐标直接用；新学校缺失坐标时用州中心近似，并记录待人工核对清单"""
    with open(coords_path, 'r', encoding='utf-8') as f:
        coords = json.load(f)

    missing = []
    for r in records:
        name = r[NAME_KEY]
        if name not in coords:
            state = str(r.get(STATE_KEY, '')).strip().upper()
            approx = STATE_CENTERS.get(state)
            if approx:
                coords[name] = approx
                missing.append((name, state, 'state_center_fallback'))
            else:
                # 连州信息都没有，给一个全美中心兜底，避免程序崩溃
                coords[name] = [39.8283, -98.5795]
                missing.append((name, state or '未知', 'us_center_fallback'))

    # 保存合并后的坐标库（含新增的近似坐标），供下次构建复用，也便于人工日后手动精修
    with open(coords_path, 'w', encoding='utf-8') as f:
        json.dump(coords, f, ensure_ascii=False, indent=0)

    return coords, missing


def inject_template(records, coords, template_path, output_path):
    """精确定位注入 SCHOOL_COORDS_DB 和 SCHOOLS（brace counting，与手动流程一致，不用正则）"""
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    marker = 'const SCHOOL_COORDS_DB = '
    idx_s = html.index(marker) + len(marker)
    depth = 0
    idx_e = idx_s
    for i, ch in enumerate(html[idx_s:], idx_s):
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                idx_e = i + 1
                break
    html = html[:idx_s] + json.dumps(coords, ensure_ascii=False) + html[idx_e:]

    start = html.index('const SCHOOLS = ') + len('const SCHOOLS = ')
    end = html.index(';\n', start)
    html = html[:start] + json.dumps(records, ensure_ascii=False) + html[end:]

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    if not os.path.exists(EXCEL_PATH):
        print(f'❌ 找不到 Excel 文件: {EXCEL_PATH}')
        sys.exit(1)

    print(f'📖 读取 Excel: {EXCEL_PATH}')
    records = load_excel_records(EXCEL_PATH)
    print(f'✓ 共 {len(records)} 所学校, {len(records[0]) if records else 0} 列')

    print('📍 合并坐标库…')
    coords, missing = merge_coords(records, COORDS_PATH)
    if missing:
        print(f'⚠️  {len(missing)} 所新学校使用了近似坐标（州中心兜底），建议人工核实精确坐标：')
        for name, state, kind in missing:
            print(f'   - {name} ({state}) [{kind}]')
    else:
        print('✓ 所有学校均有精确坐标')

    print(f'🏗️  注入模板 → {OUTPUT_PATH}')
    inject_template(records, coords, TEMPLATE_PATH, OUTPUT_PATH)

    size_kb = os.path.getsize(OUTPUT_PATH) / 1024
    print(f'✅ 完成！{len(records)} 所学校, {size_kb:.0f}KB')

    # 输出供 workflow 使用的摘要（GitHub Actions 会读取这个环境文件写入 commit message）
    summary_path = os.environ.get('GITHUB_STEP_SUMMARY')
    if summary_path:
        with open(summary_path, 'a', encoding='utf-8') as f:
            f.write(f'## 构建完成\n- 学校数量: {len(records)}\n- 输出大小: {size_kb:.0f}KB\n')
            if missing:
                f.write(f'- ⚠️ {len(missing)} 所新学校坐标为近似值，建议人工核实：\n')
                for name, state, kind in missing:
                    f.write(f'  - {name} ({state})\n')


if __name__ == '__main__':
    main()
