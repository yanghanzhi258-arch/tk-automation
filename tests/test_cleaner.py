from desktop_excel_cleaner.cleaner import clean_rows


def test_clean_rows_trims_columns_values_and_duplicates():
    headers = [" Name ", "Amount", "Empty"]
    rows = [[" Alice ", 10, None], ["Alice", 10, None], [None, None, None]]

    cleaned_headers, cleaned = clean_rows(headers, rows)

    assert cleaned_headers == ["Name", "Amount"]
    assert cleaned == [{"Name": "Alice", "Amount": 10}]


def test_clean_rows_sorts_by_column():
    headers = ["Name", "Amount"]
    rows = [["B", 2], ["A", 1]]

    _, cleaned = clean_rows(headers, rows, sort_by="Name")

    assert [row["Name"] for row in cleaned] == ["A", "B"]


def test_clean_rows_generates_names_for_blank_headers():
    headers = [None, "  "]
    rows = [["first", "second"]]

    cleaned_headers, cleaned = clean_rows(headers, rows)

    assert cleaned_headers == ["column_1", "column_2"]
    assert cleaned == [{"column_1": "first", "column_2": "second"}]
