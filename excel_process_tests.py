import pytest
from unittest.mock import MagicMock
import excel_process_2

@pytest.fixture
def processor():
    xlsx_test_file = 'test_1.xlsx'
    json_test_file = 'test_1.txt'
    processor = excel_process_2.ExcelProcessor(xlsx_test_file, json_test_file)
    return processor

def test_init(processor):
    """Test the initialization of the ExcelProcessor."""
    assert processor.wb is not None
    assert processor.ws is not None
    assert processor.json_file == 'test_1.txt'

# def test_read_json(processor, mocker):
#     """Test reading JSON file."""
#     mock_json_load = mocker.patch('json.load')
#     mock_json_load.return_value = {"table1": {"Columns": ["A", "B"], "Row Names": ["Cairo", ""], "Last Row": ["A"]}}
#     processor.read_json()
#     assert processor.json_data is not None
#     assert "table1" in processor.json_data

def test_populate_cells(processor):
    """Test populating cells with dummy data."""
    processor.populate_cells()
    assert len(processor.matrix) == processor.max_row
    assert len(processor.matrix[0]) == processor.max_col
    assert processor.matrix[0][0].value == "0"

def test_get_table_data(processor):
    """Test extracting table data from JSON."""
    test_data = {"Columns": ["A", "B"], "Row Names": ["Cairo", ""], "Last Row": [", Dolphin"]}
    columns, row_names, last_row = processor.get_table_data(test_data)
    assert columns == ["A", "B"]
    assert row_names == ["Cairo", ""]
    assert last_row == [", Dolphin"]

def test_pre_pass(processor):
    """Test the pre-pass method."""
    columns = ["A", "B"]
    row_names = ["Cairo", ""]
    last_row = [", Dolphin"]
    local_column_occurrences, local_row_name_occurrences, local_last_row_occurrences = processor.pre_pass(columns, row_names, last_row)
    assert local_column_occurrences == [(0, 0)]
    assert local_row_name_occurrences == [(1, 1)]
    assert local_last_row_occurrences == [(2, 2)]

# def test_display_table_coordinates(processor, mocker):
#     """Test display of table coordinates."""
#     mock_xy_to_excel = mocker.patch('new_matrix_method.ExcelProcessor.xy_to_excel')
#     mock_xy_to_excel.side_effect = lambda x, y: f"{x}{y}"
#     processor.tables = [MagicMock()]
#     processor.tables[0].column_headers = [(0, 0)]
#     processor.tables[0].row_names = [(1, 1)]
#     processor.tables[0].last_row = [(2, 2)]
#     mocked_print = mocker.patch('builtins.print')
#     processor.display_table_coordinates()
#     mocked_print.assert_called()

# def test_process_tables(processor, mocker):
#     """Test the overall processing of tables."""
#     mocker.patch.object(processor, 'read_json')
#     mocker.patch.object(processor, 'populate_cells')
#     mocker.patch.object(processor, 'display_table_coordinates')
#     processor.process_tables()
#     processor.read_json.assert_called_once()
#     processor.populate_cells.assert_called_once()
#     processor.display_table_coordinates.assert_called_once()