from pathlib import Path

from desktop_excel_cleaner.desktop import list_excel_files, resolve_desktop_file


def test_list_excel_files_returns_supported_files_only(tmp_path):
    (tmp_path / "a.xlsx").write_text("", encoding="utf-8")
    (tmp_path / "b.xlsm").write_text("", encoding="utf-8")
    (tmp_path / "c.csv").write_text("", encoding="utf-8")

    assert [path.name for path in list_excel_files(tmp_path)] == ["a.xlsx", "b.xlsm"]


def test_resolve_desktop_file_uses_desktop_for_bare_file_name(tmp_path):
    assert resolve_desktop_file("data.xlsx", tmp_path) == tmp_path / "data.xlsx"


def test_resolve_desktop_file_keeps_paths_with_folders(tmp_path):
    input_path = Path("folder/data.xlsx")

    assert resolve_desktop_file(str(input_path), tmp_path) == input_path
