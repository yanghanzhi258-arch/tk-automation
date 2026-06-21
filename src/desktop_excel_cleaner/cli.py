"""Command-line interface for the desktop Excel cleaner."""

from __future__ import annotations

import argparse
from pathlib import Path

from .cleaner import build_output_path, convert_excel_to_csv
from .desktop import get_desktop_path, list_excel_files, resolve_desktop_file


def parse_sheet(value: str) -> str | int:
    """Parse a sheet name or zero-based sheet index."""

    try:
        return int(value)
    except ValueError:
        return value


def choose_input_interactively() -> Path:
    """Ask the user to choose an Excel file when no CLI input is provided."""

    desktop = get_desktop_path()
    files = list_excel_files(desktop)
    print(f"桌面目录：{desktop}")

    if files:
        print("发现以下 Excel 文件：")
        for index, file_path in enumerate(files, start=1):
            print(f"  {index}. {file_path.name}")
        answer = input("请输入编号，或输入 Excel 完整路径：").strip()
        if answer.isdigit() and 1 <= int(answer) <= len(files):
            return files[int(answer) - 1]
        return resolve_desktop_file(answer, desktop)

    answer = input("桌面未发现 Excel 文件，请输入 Excel 文件名或完整路径：").strip()
    return resolve_desktop_file(answer, desktop)


def build_parser() -> argparse.ArgumentParser:
    """Build the command-line parser."""

    parser = argparse.ArgumentParser(
        description="读取桌面 Excel 文件，自动整理数据，并导出 CSV 到桌面。"
    )
    parser.add_argument("--input", help="Excel 文件名或完整路径；只写文件名时默认从桌面读取。")
    parser.add_argument("--sheet", default="0", help="工作表名称或索引，默认 0 表示第一个工作表。")
    parser.add_argument("--output", help="CSV 输出文件名或完整路径；只写文件名时默认保存到桌面。")
    parser.add_argument("--sort-by", help="按指定列排序。")
    parser.add_argument(
        "--keep-duplicates",
        action="store_true",
        help="保留完全重复行；默认会删除重复行。",
    )
    return parser


def main() -> int:
    """Run the CLI and return a process exit code."""

    parser = build_parser()
    args = parser.parse_args()

    input_path = resolve_desktop_file(args.input) if args.input else choose_input_interactively()
    output_path = build_output_path(input_path, args.output)

    try:
        result = convert_excel_to_csv(
            input_path,
            sheet_name=parse_sheet(args.sheet),
            output_path=output_path,
            drop_duplicates=not args.keep_duplicates,
            sort_by=args.sort_by,
        )
    except Exception as exc:  # noqa: BLE001 - show a friendly CLI error to end users.
        parser.exit(status=1, message=f"处理失败：{exc}\n")

    print(f"处理完成，CSV 已保存到：{result}")
    return 0
