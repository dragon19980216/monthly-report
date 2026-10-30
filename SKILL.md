---
name: monthly-report
description: Fill the 月度工作汇报表 (monthly work report) by merging the week's 周报 Excel files into the monthly template. Use when the user asks to generate/update a 月度工作汇报表 from 周报 files, or mentions 月报、月度汇报、monthly report from weekly reports.
---

# 月度工作汇报表 (Monthly Work Report)

Merge 周报 (weekly report) files into the 月度工作汇报表 template.

## Inputs

- 周报 files: `D:/桌面/桌面文件/周报/YYYYMMDD周报文件.xlsx` (the 周报 folder; ignore the 周报\2024 / 周报\2025 subfolders unless told otherwise).
- Template: `assets/月度工作汇报表模板.xlsx` (in this skill folder). Never modify the original template at `D:/桌面/桌面文件/周报/月报/【模板】月度工作汇报表+部门.xlsx`.
- Output: `D:/桌面/桌面文件/周报/月报/月度工作汇报YYYYMM.xlsx` (user's naming convention, e.g. `月度工作汇报202609`). If it already exists, ask before overwriting, or use `--force`.

## Procedure

1. Identify which 周报 files belong to the target month (report date in the month). Confirm the list with the user when it is ambiguous.
2. Read each 周报: 本周工作 (rows 7-16, columns D/G) and 下周工作安排 (rows 18-25, columns D/J).
3. Copy the template to the output path, then run the fill script:
   `python scripts/fill_monthly.py --month YYYY-MM [--weeks <周报文件夹>] [--force]`
   The script finds the week files for that month itself (default folder `D:/桌面/桌面文件/周报`), auto-merges/dedupes items, and can take optional `--work-items` / `--plan-items` JSON files when you have decided the merged content yourself (format: `[[text, effect], ...]`).
   The script writes 上月工作汇报 rows 7-20 and 下月工作安排 rows 22-27 of Sheet1, sets 汇报日期 (B4) to the **last workday** of the month (matching the template's own example `2025-06-27`), 汇报时间段 (O4) to `YYYY/M/<first workday>-<last day>` (e.g. `2026/9/1-30`), and leaves 未完成原因 / 遇到的困难 blank.
4. Before finishing, verify the filled values, and merge/dedupe repeated work items across weeks following the rules in `references/fill-rules.md`.

## Rules (see references/fill-rules.md for details)

- 上月工作汇报 column D: numbered items `1、...`；column G: final completion effect (latest state).
- Merge duplicate/progressive items across weeks into one line with the final effect; drop items that appear only in 下周工作安排 of later weeks but were already completed.
- 下月工作安排: take the next-week plan of the last week (rows 18-25), deduped.

## Fixed row capacity (strict, added 2026-10)

The template's table layout is fixed and must never grow. Only the date fields and the data cells may change:

- 上月工作汇报: data rows 7-20 = **exactly 14 rows**.
- 下月工作安排: data rows 22-27 = **exactly 6 rows**.

If extraction yields more items than the capacity, **merge same-theme items until they fit**; never insert rows and never truncate silently:

- merge same-product / same-topic items into one line (e.g. 设计开发资料整理 + 体考答疑 → one line; 打标图更新 + 打标图、组件图、包装图整理 → one line);
- merge small same-category maintenance items (e.g. several IT/office-equipment fixes → one line);
- combine their completion effects into the effect cell, joined with `；`.

The fill script enforces this: it exits with an error when items exceed capacity. `--insert-rows` exists only for the rare case where the user explicitly asks for extra rows.

## Template fidelity (strict)

The fill script copies the template file and writes **values only**. It must never alter the template appearance:

**Template files are read-only. The original template (`D:/桌面/桌面文件/周报/月报/【模板】月度工作汇报表+部门.xlsx`) and the skill asset template (`assets/月度工作汇报表模板.xlsx`) must never be opened for editing, styled, saved over, or otherwise changed in any way. All work happens on the freshly copied output file only. Verify after generating that the two template files are untouched (unchanged file size / last-modified time).**

- **No style changes**: do not modify fonts, borders, fills, alignment, or number formats of any existing template cell.
- **No structural changes**: never insert/delete rows and never add/remove/modify merged cells. The only allowed edits are: the date values in B4/O4, data-zone cell values (C/D/G in rows 7-20; C/D/J in rows 22-27), wrap-text on those data cells, and content-estimated row heights of the data rows. Row insertion happens only with the explicit `--insert-rows` opt-in.
- **Keep original borders**: each new row inherits the exact thin-border pattern of the template row it is based on (middle data row vs. last row with a bottom border). Never add, remove, or redraw borders.
- **B-column merged cells always preserved**: the 上月工作汇报 B-cell (B6 downward) and 下月工作安排 B-cell (B21 downward) must always stay as **merged cells**. On row expansion the script unbinds the old B merge, inserts, then rebinds the B merge across the full new range; never split it.
- **Row heights auto-fit**: the script enables wrap-text on the data cells (cols 4/7/10 in the two data zones) and writes a content-estimated row height so each row's height fits its text. Font/borders/fills/alignment(H/V) of template cells stay untouched; only the data-zone alignment gains `wrap_text=True` and the row heights become content-based. Header/title rows keep their template fixed heights.
- **Never edit the source template** at D:/桌面/桌面文件/周报/月报/【模板】月度工作汇报表+部门.xlsx; work only on the copy.

After generating, open the output and confirm: fonts/borders/merges match the template, B cells in both sections are merged, and row 6+ and 21+ blocks look unchanged.
