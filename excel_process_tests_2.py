import pytest
from excel_process_2 import ExcelProcessor

# Define test scenarios and expected values
test_cases = [
    {
        "test_name": "large_test_1",
        "expected_values": [
            "CATEGORY", "TASK", "LABOR", "MATERIALS", "FIXED COST", "BUDGET", "ACTUAL", "UNDER/OVER",
            "SUMMARY", "INCOME", "EXPENSES", "OPERATING EXPENSE", "PAYROLL", "OFFICE", "ENTERTAINMENT", "HEALTH", "VACATION/HOLIDAY"
        ]
    },
    {
        "test_name": "large_test_2",
        "expected_values": [
            "REORDER (auto-fill)", "ITEM NO.", "DATE OF LAST ORDER", "ITEM NAME", "VENDOR", "STOCK LOCATION", "DESCRIPTION", 
            "COST PER ITEM", "STOCK QUANTITY", "TOTAL VALUE", "REORDER LEVEL", "DAYS PER REORDER", "ITEM REORDER QUANTITY", "ITEM DISCONTINUED?"
        ]
    },
    {
        "test_name": "large_test_3",
        "expected_values": ["TASKS", "ASSIGNED TO", "PRIORITY", "STATUS"]
    },
    {
        "test_name": "large_test_4",
        "expected_values": [
            "COMPANY NAME", "YEARS REPRESENTED", "2022", "2021", "REVENUE", "EXPENSES", "NET INCOME BEFORE TAXES"
        ]
    },
    {
        "test_name": "large_test_5",
        "expected_values": ["Command", "Syntax", "Description"]
    },
    {
        "test_name": "simple_test_1",
        "expected_values": ["A", "B", "C"]
    },
    {
        "test_name": "simple_test_2",
        "expected_values": ["A", "B", "C"]
    },
    {
        "test_name": "simple_test_3",
        "expected_values": ["Animal", "Breed", "Age", "Pet", "Name"]
    },
    {
        "test_name": "simple_test_4",
        "expected_values": ["Alpha", "Beta", "Gamma"]
    },
    {
        "test_name": "simple_test_5",
        "expected_values": ["Alpha", "Beta", "Gamma"]
    }
]

@pytest.mark.parametrize("test_case", test_cases)
def test_specific_values_in_output(test_case):
    test_name = test_case["test_name"]
    expected_values = test_case["expected_values"]
    
    # Define the paths for the input Excel and JSON files
    excel_file = f"{test_name}/{test_name}.xlsx"
    json_file = f"{test_name}/{test_name}_json.txt"
    
    # Initialize the processor
    processor = ExcelProcessor(excel_file, json_file)
    processor.process_tables()
    
    # Collect all values from the output for verification
    # Assuming the output is stored in a list of tables or similar structure
    output_values = []
    for table in processor.tables:
        output_values.extend([processor.matrix[cell[0]][cell[1]].value for cell in table.column_headers])
        output_values.extend([processor.matrix[cell[0]][cell[1]].value for cell in table.row_names])
        output_values.extend([processor.matrix[cell[0]][cell[1]].value for cell in table.last_row])
    
    # Check if all expected values are in the output
    for value in expected_values:
        assert value in output_values, f"Expected value '{value}' not found in output for {test_name}"