"""扫描 Excel 文件，生成字段结构统计与优先级清单，供 finance-sop-kickoff skill 使用。

依赖：pandas, openpyxl（读 .xlsx）、xlrd（读旧版 .xls）

用法：
    python scan_fields.py "path/to/file.xlsx"
    python scan_fields.py "path/to/file.xlsx" --sheet "Sheet1"
    python scan_fields.py "path/to/file.xlsx" --null-threshold 0.05 --max-categories 20

输出：按优先级（高/中/低）分组的 Markdown 表格，打印到 stdout。
只做结构统计，不做业务判断——优先级和统计值仅用于生成"待确认清单"，
真正的口径、正常范围、异常阈值必须由访谈确认后写入 03 文档。
"""

import argparse
import sys

import pandas as pd

# Windows 终端默认代码页常导致中文输出乱码；强制 stdout 用 UTF-8。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass


def _looks_like_code(name: str) -> bool:
    name = str(name)
    if len(name) <= 6 and any(c.isdigit() for c in name):
        return True
    letters = [c for c in name if c.isalpha()]
    if letters and sum(c.isupper() for c in letters) == len(letters) and len(name) <= 10:
        return True
    return False


def _value_summary(non_null: pd.Series, is_numeric: bool, n_unique: int, max_categories: int) -> str:
    if len(non_null) == 0:
        return "(全空)"
    if is_numeric:
        return (
            f"min={non_null.min():.4g}, max={non_null.max():.4g}, "
            f"mean={non_null.mean():.4g}, median={non_null.median():.4g}"
        )
    if n_unique <= max_categories:
        counts = non_null.value_counts().head(max_categories)
        return "; ".join(f"{k}({v})" for k, v in counts.items())
    sample = non_null.astype(str).unique()[:3]
    return f"样例: {', '.join(sample)} ...(共{n_unique}个唯一值)"


def classify_field(series: pd.Series, name: str, null_threshold: float, max_categories: int) -> dict:
    non_null = series.dropna()
    n = len(series)
    null_rate = 1 - len(non_null) / n if n else 0
    n_unique = non_null.nunique()

    is_numeric = pd.api.types.is_numeric_dtype(series)
    is_free_text = False
    if not is_numeric and len(non_null) > 0:
        avg_len = non_null.astype(str).str.len().mean()
        uniqueness_ratio = n_unique / len(non_null) if len(non_null) else 0
        is_free_text = uniqueness_ratio > 0.5 and avg_len > 10

    if is_free_text:
        field_type = "自由文本"
    elif is_numeric:
        field_type = "数值/度量"
    elif n_unique <= max_categories:
        field_type = "分类/代码"
    else:
        field_type = "疑似主键/ID"

    reasons = []
    priority = "低"

    if field_type != "自由文本":
        if null_rate > null_threshold:
            reasons.append(f"空值率{null_rate:.1%}超过阈值{null_threshold:.0%}")
            priority = "高"
        if is_numeric and len(non_null) >= 4:
            q1, q3 = non_null.quantile(0.25), non_null.quantile(0.75)
            iqr = q3 - q1
            if iqr > 0:
                lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
                outliers = non_null[(non_null < lo) | (non_null > hi)]
                if len(outliers) > 0:
                    reasons.append(f"存在{len(outliers)}个IQR外的极端值")
                    priority = "高"
        if _looks_like_code(name):
            reasons.append("字段名疑似缩写/代码，含义不直观")
            priority = "高"
        if field_type in ("分类/代码", "疑似主键/ID") and priority != "高":
            priority = "中"

    summary = _value_summary(non_null, is_numeric, n_unique, max_categories)

    return {
        "字段名": name,
        "类型": field_type,
        "空值率": f"{null_rate:.1%}",
        "唯一值数": n_unique,
        "取值/范围概览": summary,
        "优先级": priority,
        "触发原因": "；".join(reasons) if reasons else "-",
    }


def scan_sheet(df: pd.DataFrame, sheet_name: str, null_threshold: float, max_categories: int) -> str:
    rows = [classify_field(df[col], col, null_threshold, max_categories) for col in df.columns]
    order = {"高": 0, "中": 1, "低": 2}
    rows.sort(key=lambda r: order[r["优先级"]])

    lines = [f"\n## Sheet: {sheet_name}（{len(df)} 行 × {len(df.columns)} 列）\n"]
    for level in ("高", "中", "低"):
        level_rows = [r for r in rows if r["优先级"] == level]
        if not level_rows:
            continue
        lines.append(f"\n### 优先级：{level}\n")
        lines.append("| 字段名 | 类型 | 空值率 | 唯一值数 | 取值/范围概览 | 触发原因 |")
        lines.append("|---|---|---|---|---|---|")
        for r in level_rows:
            lines.append(
                f"| {r['字段名']} | {r['类型']} | {r['空值率']} | {r['唯一值数']} | "
                f"{r['取值/范围概览']} | {r['触发原因']} |"
            )
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="扫描Excel字段结构，生成分级提问清单")
    parser.add_argument("file", help="Excel文件路径")
    parser.add_argument("--sheet", default=None, help="只扫描指定sheet，默认扫描全部sheet")
    parser.add_argument("--null-threshold", type=float, default=0.05, help="空值率高优先级阈值，默认0.05")
    parser.add_argument("--max-categories", type=int, default=20, help="分类型字段唯一值上限，默认20")
    parser.add_argument(
        "--output",
        default=None,
        help="将结果写入该 UTF-8 文件，而不是打印到终端（Windows下建议使用，避免控制台代码页把中文转乱码）",
    )
    args = parser.parse_args()

    try:
        xl = pd.ExcelFile(args.file)
    except Exception as exc:
        print(f"读取文件失败：{exc}", file=sys.stderr)
        sys.exit(1)

    sheets = [args.sheet] if args.sheet else xl.sheet_names
    output_parts = [f"# 字段扫描结果：{args.file}"]

    for sheet in sheets:
        try:
            df = xl.parse(sheet)
        except Exception as exc:
            print(f"读取 sheet「{sheet}」失败：{exc}", file=sys.stderr)
            continue
        output_parts.append(scan_sheet(df, sheet, args.null_threshold, args.max_categories))

    result = "\n".join(output_parts)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"已写入：{args.output}")
    else:
        print(result)


if __name__ == "__main__":
    main()
