import pytest
import sys
sys.path.append('../')
from src.excel_process_2 import ExcelProcessor

# Define test scenarios and expected values
test_cases = [
    {
        "test_name": "large_test_1",
        "expected_values_table_1": [
            "TASK", "TOTAL", "UNDER/OVER", "None"],
        "expected_values_table_2": [
            "SUMMARY", "TOTAL EXPENSES", "UNDER/OVER", "$173"
        ],
        "expected_values_table_3": [
            "INCOME", "NONE", "TOTAL", "NONE"
        ],
        "expected_values_table_4": [
            "OPERATING EXPENSE", "Web Hosting & Domains", "None", "None"
        ],
        "expected_values_table_5": [
            "PAYROLL", "Contractor Wages", "None", "None"
        ],
        "expected_values_table_6": [
            "OFFICE", "Pet Supplies", "None", "None"
        ],
        "expected_values_table_7": [
            "ENTERTAINMENT", "Outdoor Recreation", "None", "None"
        ],
        "expected_values_table_8": [
            "HEALTH", "Life Insurance", "None", "None"
        ],
        "expected_values_table_9": [
            "VACATION/HOLIDAY", "Rental Car", "None", "None"
        ]
    },
    {
        "test_name": "large_test_2",
        "expected_values_table_1": [
            "REORDER (auto-fill)", "OK", "ITEM DISCONTINUED?", "None"
        ]
    },
    {
        "test_name": "large_test_3",
        "expected_values_table_1": ["TASKS", "Launch", "STATUS", "None"]
    },
    {
        "test_name": "large_test_4",
        "expected_values_table_1": [
            "COMPANY NAME", "None", "YEARS REPRESENTED", "2021"
        ],
        "expected_values_table_2": [
            "REVENUE", "TOTAL REVENUE", "2021", "None"
        ],
        "expected_values_table_3": [
            "EXPENSES", "TOTAL EXPENSES", "2021", "None"
        ],
        "expected_values_table_4": [
            "NET INCOME BEFORE TAXES", "INCOME TAX EXPENSE", "None", "None"
        ]
    },
    {
        "test_name": "large_test_5",
        "expected_values_table_1": [
            "Command", "nano", "Description", "Opens the nano text editor, which allows you to edit files in the terminal. If <file> is specified, nano will open it for editing."
        ]
    },
    {
        "test_name": "simple_test_1",
        "expected_values_table_1": ["A", "Cairo", "B", "Abra"]
    },
    {
        "test_name": "simple_test_2",
        "expected_values_table_1": ["A", "Cairo", "C", "Barracuda"],
        "expected_values_table_2": ["A", "Barracuda"]
    },
    {
        "test_name": "simple_test_3",
        "expected_values_table_1": ["Animal", "Cat", "Age", "6"],
        "expected_values_table_2": ["Breed", "Lab", "Animal", "Dog"],
        "expected_values_table_3": ["Pet", "Husky", "Name", "Charlie"]
    },
    {
        "test_name": "simple_test_4",
        "expected_values_table_1": ["Alpha", "G", "Gamma", "None"]
    },
    {
        "test_name": "simple_test_5",
        "expected_values_table_1": ["Alpha", "Alpha", "Gamma", "None"]
    }
]

@pytest.mark.parametrize("test_case", test_cases)
def test_coordinates(test_case):
    test_name = test_case["test_name"]
    for i in range(1, len(test_case)):
        expected_values = test_case[f"expected_values_table_{i}"]
    
        # Define the paths for the input Excel and JSON files
        excel_file = f"{test_name}/{test_name}.xlsx"
        json_file = f"{test_name}/{test_name}_json.txt"
        
        # Initialize the processor
        processor = ExcelProcessor(excel_file, json_file)
        processor.process_tables()
        
        # Collect all values from the output for verification
        # Assuming the output is stored in a list of tables or similar structure
        output_values = []
        output_values.extend(processor.corner_coordinates)

        # Check if all expected values are in the output
        for table in output_values:
            for value in expected_values:
                assert value in table, f"Expected value '{value}' not found in output for {test_name}"
                break

