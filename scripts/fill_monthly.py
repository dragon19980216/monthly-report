"""汇总周报为月度工作汇报表。

用法：
    python fill_monthly.py --month 2026-09 [--weeks 目录] [--force] [--insert-rows]

行为：
    1. 从本技能 assets/月度工作汇报表模板.xlsx 复制模板到 月报 目录，命名 月度工作汇报YYYYMM.xlsx。
    2. B4 写汇报日期（当月最后一个工作日，与模板示例一致），O4 写汇报时间段（YYYY/M/第一个工作日-最后一天）。
    3. 汇总 周报 目录中该月各周报文件：
       - 上月工作汇报区：合并各周"本周工作"列(D)+完成效果列(G)，去除跨周重复/递进项，
         保留最终状态；未完成原因、遇到的困难两列留空。
       - 下月工作安排区：取各周"下周工作安排"中尚未完成、仍适用于下月的条目，
         去重合并（具体工作安排列D + 完成效果列J）。
   4. 只写值，不修改任何模板样式（字体、边框、对齐、合并、数值格式均保持原样）；
      数据区行高按内容行数显式写入，表头/标题行保持模板高度；
      条目超出固定行数（工作 14 / 计划 6）时默认报错退出，须合并同类条目适配模板；
      仅当显式加 --insert-rows 时才插入行（新行复制行样式并按新行数重建合并区）。
    5. 若目标文件已存在且未加 --force，则拒绝覆盖并退出。
"""
import argparse
import calendar
import json
import copy
import os
import re
import shutil
import sys
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Alignment

# 估算行高时使用的字符像素度量（仅用于行高，不改动任何单元格样式）
_LINE_MIN = 16.0


def _font_char_px(font) -> float:
    """估算某字体/字号下单个字符宽度（Excel 宽度单位），用于行高估算。

    中文字符在 Excel 宽度单位下约等于字号/7 左右；英文/数字约其一半。
    这里按内容实际字符计算，避免统一系数导致行高偏差。
    """
    size = font.size or 11.0
    return size / 7.0 * 2  # 中文字符宽度（宽度单位）≈ size/7*2? 用经验值


def est_lines(text: str, font, col_width: float) -> int:
    """估算文本在给定列宽（Excel 宽度单位）下占几行（按字符实际宽度）。"""
    total = 0
    size = font.size or 11.0
    cjk = size / 7.0 * 2.0  # 中文字符宽度（宽度单位）≈ size/7*2? 用经验值
    latin = size / 7.0  # 英文/数字/标点宽度约为中文一半
    for seg in str(text).split("\n"):
        cap = max(1.0, col_width - 0.5)  # 单元格左右内边距近似
        need = 0.0
        for ch in seg:
            if ch == "\n":
                continue
            need += cjk if ord(ch) > 0x2E80 else latin
            if need >= cap:
                break
        lines = 1 + int(need // cap + (1 if need % cap > 1e-6 else 0))
        total += max(1, lines)
    return total


def cell_lines(ws, row: int, col: int, value) -> int:
    """按单元格字体与列宽估算 value 占几行。"""
    if value is None:
        return 1
    cell = ws.cell(row, col)
    from openpyxl.utils import get_column_letter
    col_letter = get_column_letter(col)
    dim = ws.column_dimensions[col_letter]
    width = dim.width if dim.width else 8.43
    return est_lines(value, cell.font, width)


def auto_row_heights(ws, rows: list, cols: list):
    """仅设置行高（按内容估算），不改动任何单元格样式。"""
    for r in rows:
        n = 1
        for c in cols:
            v = ws.cell(r, c).value
            if v:
                n = max(n, cell_lines(ws, r, c, v))
        ws.row_dimensions[r].height = max(_LINE_MIN, n * 14 + 6)


def snapshot_row_style(ws, row: int) -> dict:
    """快照某行所有列的样式（字体/边框/填充/对齐/数值格式）。"""
    snap = {}
    for c in range(1, ws.max_column + 1):
        s = ws.cell(row, c)
        snap[c] = {
            "font": copy.copy(s.font),
            "border": copy.copy(s.border),
            "fill": copy.copy(s.fill),
            "align": copy.copy(s.alignment),
            "numfmt": s.number_format,
        }
    return snap


def apply_row_style(ws, row: int, snap: dict):
    """把快照样式应用到某行（仅样式，不改值与合并）。"""
    for c, st in snap.items():
        d = ws.cell(row, c)
        d.font = copy.copy(st["font"])
        d.border = copy.copy(st["border"])
        d.fill = copy.copy(st["fill"])
        d.alignment = copy.copy(st["align"])
        d.number_format = st["numfmt"]


def copy_row_style(ws, src_row: int, dst_row: int):
    """把 src_row 整行样式（值除外）复制到 dst_row。"""
    apply_row_style(ws, dst_row, snapshot_row_style(ws, src_row))

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(SKILL_ROOT, "assets", "月度工作汇报表模板.xlsx")
BASE_DIR = r"D:/桌面/桌面文件/周报"
WEEKS_DIR = os.path.join(BASE_DIR, "周报")  # 默认周报目录（实际周报表所在目录）
OUT_DIR = os.path.join(BASE_DIR, "月报")


def last_workday(year: int, month: int) -> date:
    """返回当月最后一个工作日（周一~周五）。"""
    last_day = calendar.monthrange(year, month)[1]
    d = date(year, month, last_day)
    while d.weekday() >= 5:
        d = date.fromordinal(d.toordinal() - 1)
    return d

def first_workday(year: int, month: int) -> date:
    """返回当月第一个工作日（周一~周五）。"""
    d = date(year, month, 1)
    while d.weekday() >= 5:
        d = date.fromordinal(d.toordinal() + 1)
    return d


def find_week_files(month: str, weeks_dir: str) -> list:
    """找出该月（YYYY-MM）下的周报文件，按文件名日期排序。"""
    pat = re.compile(r"(\d{4})(\d{2})(\d{2})")
    files = []
    for name in os.listdir(weeks_dir):
        if not name.endswith(".xlsx") or name.startswith("~$"):
            continue
        m = pat.search(name)
        if not m:
            continue
        if m.group(1) + m.group(2) == month.replace("-", ""):
            files.append((m.group(2) + m.group(3), os.path.join(weeks_dir, name)))
    files.sort()
    return [p for _, p in files]


def read_zone(path: str, hdr_row: int, key_row_range: range, col_text: int, col_eff: int):
    """读取周报某个区块的(工作描述, 完成效果)列表，跳过空行。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sheet1"]
    items = []
    for r in key_row_range:
        text = ws.cell(r, col_text).value
        effect = ws.cell(r, col_eff).value
        if text is not None and str(text).strip():
            items.append((str(text).strip(), str(effect).strip() if effect is not None else ""))
    wb.close()
    return items


def norm(s: str) -> str:
    return re.sub(r"\s+", "", s)


def merge_work(all_items: list) -> list:
    """合并跨周工作项：同一事项在多周出现（递进/完成），取最终状态。

    启发式：以归一化前缀匹配判重，后周的同名条目覆盖前周（保留最终完成效果）。
    无法匹配的新条目追加。
    """
    merged: list = []
    for text, effect in all_items:
        key = norm(text)[:12] or norm(text)
        found = False
        for i, (mt, me) in enumerate(merged):
            if norm(mt)[:12] in key or key in norm(mt)[:12] or norm(mt)[:12] == key:
                merged[i] = (text, effect)
                found = True
                break
        if not found:
            merged.append((text, effect))
    return merged


def main():
    ap = argparse.ArgumentParser(description="汇总周报为月度工作汇报表")
    ap.add_argument("--month", required=True, help="汇报月，格式 YYYY-MM，如 2026-09")
    ap.add_argument("--weeks", default=None, help="周报文件所在目录，默认 D:/桌面/桌面文件/周报")
    ap.add_argument("--work-items", default=None, help="JSON 文件：上月工作汇报合并结果 [[text, effect], ...]")
    ap.add_argument("--plan-items", default=None, help="JSON 文件：下月工作安排合并结果 [[text, effect], ...]")
    ap.add_argument("--force", action="store_true", help="允许覆盖已存在的目标文件")
    ap.add_argument("--insert-rows", action="store_true",
                    help="条目超出模板固定行数时允许插入行（默认禁止，模板结构不可改动）")
    args = ap.parse_args()

    m = re.fullmatch(r"(\d{4})-(\d{2})", args.month)
    if not m:
        sys.exit("--month 必须为 YYYY-MM 格式")
    year, month = int(m.group(1)), int(m.group(2))
    ym = args.month.replace("-", "")

    weeks_dir = args.weeks or WEEKS_DIR
    if not os.path.isdir(weeks_dir):
        sys.exit(f"周报目录不存在: {weeks_dir}")

    out_dir = OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"月度工作汇报{ym}.xlsx")
    if os.path.exists(out_path) and not args.force:
        sys.exit(f"目标文件已存在: {out_path}\n如需覆盖请加 --force，或删除旧文件后重跑。")

    # 汇总工作项
    if args.work_items and os.path.exists(args.work_items):
        with open(args.work_items, encoding="utf-8") as f:
            work = [(str(a).strip(), str(b).strip()) for a, b in json.load(f)]
    else:
        files = find_week_files(ym, weeks_dir)
        if not files:
            sys.exit(f"未找到 {args.month} 的周报文件（目录: {weeks_dir}）")
        print(f"汇总周报文件: {[os.path.basename(p) for p in files]}")
        all_items = []
        for p in files:
            all_items.extend(read_zone(p, 6, range(7, 17), 4, 7))  # 本周工作/完成效果
        work = merge_work(all_items)

    if args.plan_items and os.path.exists(args.plan_items):
        with open(args.plan_items, encoding="utf-8") as f:
            plan = [(str(a).strip(), str(b).strip()) for a, b in json.load(f)]
    else:
        plan = []
        if args.work_items is None:
            seen = set()
            for p in files:
                for text, effect in read_zone(p, 17, range(18, 26), 4, 10):
                    if norm(text) not in seen:
                        seen.add(norm(text))
                        plan.append((text, effect))

    # 模板固定区（行号基于原始模板）：
    #   工作区数据行 7-20（14 行），"下月工作安排"表头行 21，计划区数据行 22-27（6 行）
    WORK_START, WORK_BASE = 7, 14
    PLAN_START, PLAN_BASE = 22, 6
    PLAN_HDR_ROW = 21
    PLAN_TAIL_ROW = 28            # 原计划区末行（27）之后，计划区扩展的插入锚点
    WORK_MID_SRC, WORK_LAST_SRC = 10, 20   # 新工作行样式源：中间行 / 末行（含底部细边框）
    PLAN_MID_SRC, PLAN_LAST_SRC = 24, 27   # 新计划行样式源：中间行 / 末行（含底部细边框）
    # 模板各数据行的横向合并组：工作区 D:F、G:I、J:L、M:O；计划区 D:I、J:O
    WORK_HMERGES = [(4, 6), (7, 9), (10, 12), (13, 15)]
    PLAN_HMERGES = [(4, 9), (10, 15)]

    # 未完成的项放最后（已完成在前）
    def done_eff(e):
        e = e.strip()
        return e.startswith("已完成") or e in ("", "供注册发补使用") or e.endswith("已完成")

    work_sorted = [it for it in work if done_eff(it[1])] + [it for it in work if not done_eff(it[1])]

    n_work, n_plan = len(work_sorted), len(plan)

    # 模板结构不可改动：超行默认拒绝，须合并同类条目适配固定行数
    if (n_work > WORK_BASE or n_plan > PLAN_BASE) and not args.insert_rows:
        sys.exit(
            f"条目数超出模板固定行数（工作 {n_work}/{WORK_BASE}，计划 {n_plan}/{PLAN_BASE}）。\n"
            "模板结构不可改动：请把同类/同主题条目合并至行数内后重跑，"
            "或用 --work-items / --plan-items 提供合并后的 JSON；"
            "确需加行时必须显式加 --insert-rows。"
        )

    # 校验全部通过后才写目标文件：任何失败都不碰已有成品
    if not os.path.exists(TEMPLATE):
        sys.exit(f"模板文件缺失: {TEMPLATE}")
    shutil.copyfile(TEMPLATE, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb["Sheet1"]

    # 日期与时间段
    start = first_workday(year, month)
    end = last_workday(year, month)
    ws["B4"] = datetime(end.year, end.month, end.day)
    ws["O4"] = f"{year}/{start.month}/{start.day}-{end.day}"

    def expand_rows():
        """条目超出模板固定行数时插入行，严格保持模板样式：

        - 插入前先解绑 B 列纵向合并区，插入完成后按新行数重建（始终为合并单元格）
        - 新数据行样式复制自模板同角色行：中间行 ← 中间样式源，末行 ← 末行样式源
          （末行样式源含底部细边框，与模板区末行一致）
        - 新数据行补齐模板的横向合并组（工作区 D:F/G:I/J:L/M:O；计划区 D:I/J:O）
        - "下月工作安排"表头行随 insert_rows 平移，原样式（字体/边框/粗体）自动保留
        - 不写行高，保持模板默认，由 Excel 按内容自动调整
        """
        plan_extra = max(0, n_plan - PLAN_BASE)
        work_extra = max(0, n_work - WORK_BASE)
        if not (plan_extra or work_extra):
            return work_extra, plan_extra

        # openpyxl 的 insert_rows 不会平移已存在的合并区，因此先解绑数据区
        # （第 6 行起）全部合并（B 列纵向 + 横向合并组），插入完行后在新位置重建。
        b_merges = [m for m in ws.merged_cells.ranges if m.min_col == 2 and m.max_col == 2 and m.min_row >= 6]
        h_merges = [m for m in ws.merged_cells.ranges
                    if (m.min_col, m.max_col) in WORK_HMERGES + PLAN_HMERGES]
        for mr in b_merges + h_merges:
            ws.unmerge_cells(str(mr))

        # 先扩展计划区（锚点 28 之后，不影响工作区与表头行），再扩展工作区
        if plan_extra:
            ws.insert_rows(PLAN_TAIL_ROW, plan_extra)
            for i in range(plan_extra):
                r = PLAN_TAIL_ROW + i
                src = PLAN_LAST_SRC if i == plan_extra - 1 else PLAN_MID_SRC
                copy_row_style(ws, src, r)
            print(f"提示：下月工作安排 条目超出模板行数，已在第 {PLAN_TAIL_ROW} 行前插入 {plan_extra} 行"
                  f"（样式复制自模板，合并区已同步调整）。")

        if work_extra:
            ws.insert_rows(PLAN_HDR_ROW, work_extra)
            for i in range(work_extra):
                r = PLAN_HDR_ROW + i
                src = WORK_LAST_SRC if i == work_extra - 1 else WORK_MID_SRC
                copy_row_style(ws, src, r)
            print(f"提示：上月工作汇报 条目超出模板行数，已在第 {PLAN_HDR_ROW} 行前插入 {work_extra} 行"
                  f"（样式复制自模板，合并区已同步调整）。")

        plan_hdr_new = PLAN_HDR_ROW + work_extra

        # 重建工作区横向合并组（模板原有行 + 新插入行，行号 7 起共 n_work 行）
        for i in range(n_work):
            r = WORK_START + i
            for c1, c2 in WORK_HMERGES:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        # 重建计划区横向合并组（表头行 + 数据行）
        for r in [plan_hdr_new] + [plan_hdr_new + 1 + i for i in range(n_plan)]:
            for c1, c2 in PLAN_HMERGES:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        # 重建 B 列纵向合并区（标签行起，到各区末行止，始终为合并单元格）
        ws.merge_cells(start_row=6, end_row=WORK_START + n_work - 1, start_column=2, end_column=2)
        if n_plan:
            ws.merge_cells(start_row=plan_hdr_new, end_row=plan_hdr_new + n_plan,
                           start_column=2, end_column=2)
        return work_extra, plan_extra

    if args.insert_rows:
        work_extra, plan_extra = expand_rows()
    else:
        work_extra, plan_extra = 0, 0

    # 写入数据值（行号已含插入平移）：工作 7 起；计划数据新起点 = 22 + work_extra
    plan_start = PLAN_START + work_extra
    for i, (text, effect) in enumerate(work_sorted):
        r = WORK_START + i
        ws.cell(r, 4).value = text
        if effect:
            ws.cell(r, 7).value = effect
    for i, (text, effect) in enumerate(plan):
        r = plan_start + i
        ws.cell(r, 4).value = text
        if effect:
            ws.cell(r, 10).value = effect

    # 行高自适应：数据单元格开启自动换行（仅数据区，不动模板其它样式），
    # 再按内容估算每行所需行数写入行高，使行高随内容自动调整。
    data_rows = list(range(WORK_START, WORK_START + n_work)) + list(range(plan_start, plan_start + n_plan))
    data_cols = [4, 7, 10]
    for r in data_rows:
        for c in data_cols:
            cell = ws.cell(r, c)
            if cell.alignment is not None:
                cell.alignment = Alignment(horizontal=cell.alignment.horizontal,
                                           vertical=cell.alignment.vertical,
                                           wrap_text=True)
            else:
                cell.alignment = Alignment(wrap_text=True)
        auto_row_heights(ws, [r], data_cols)

    # 序号：按各区实际行数重排
    for i in range(n_work):
        ws.cell(WORK_START + i, 3).value = i + 1
    for i in range(n_plan):
        ws.cell(plan_start + i, 3).value = i + 1

    wb.save(out_path)
    print(f"已保存: {out_path}")
    print(f"上月工作汇报 {len(work_sorted)} 条；下月工作安排 {len(plan)} 条；汇报时间段 {year}/{start.month}/{start.day}-{end.day}")


if __name__ == "__main__":
    main()
