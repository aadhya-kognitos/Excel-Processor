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
        "test_name": "large_test_6",
        "expected_values_table_1": [
            "Income 1", "Total monthly income", "$4,300.00", "$4,600.00"
        ],
        "expected_values_table_2": [
            "Income 1", "Total monthly income", "$4,300.00", "$4,300.00"
        ],
        "expected_values_table_3": [
            "PROJECTED BALANCE", "(Actual minus projected)", "$3,405.00", "($341.00)"
        ],
        "expected_values_table_4": [
            "HOUSING", "Subtotal", "Difference", "-$41.00"
        ],
        "expected_values_table_5": [
            "TRANSPORTATION", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_6": [
            "INSURANCE", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_7": [
            "FOOD", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_8": [
            "PETS", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_9": [
            "PERSONAL CARE", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_10": [
            "ENTERTAINMENT", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_11": [
            "LOANS", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_12": [
            "TAXES", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_13": [
            "SAVINGS OR INVESTMENTS", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_14": [
            "GIFTS AND DONATIONS", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_15": [
            "LEGAL", "Subtotal", "Difference", "$0.00"
        ],
        "expected_values_table_16": [
            "TOTAL PROJECTED COST", "TOTAL INFERENCE", "$1,195.00", "($41.00)"
        ]
    },
    {
        "test_name": "large_test_7",
        "expected_values_table_1": [
            "Task", "Gather feedback", "END", "7/19/24"
        ]
    },
    {
        "test_name": "large_test_8",
        "expected_values_table_1": [
            "BUDGET TOTALS", "Balance (income minus expenses)", "DIFFERENCE", "985"
        ],
        "expected_values_table_2": [
            "ACTUAL EXPENSES", "Total", "15% REDUCTION", "2895.00"
        ]
    },
    {
        "test_name": "large_test_9",
        "expected_values_table_1": [
            "% Done", "0 %", "Notes", "Meeting"
        ]
    },
    {
        "test_name": "large_test_10",
        "expected_values_table_1": [
            "SUMMARY", "Total budget", "OVER/UNDER", "None"
        ],
        "expected_values_table_2": [
            "APPAREL", "Total for apparel", "OVER/UNDER", "None"
        ],
        "expected_values_table_3": [
            "DECORATIONS", "Total for decorations", "OVER/UNDER", "None"
        ]
    },
    {
        "test_name": "large_test_11",
        "expected_values_table_1": [
            "Budget", "Remaining cash", "$1,500.00", "$1,135.00"
        ],
        "expected_values_table_2": [
            "Item", "Winter jacket", "Total cost", "-"
        ],
    },
    {
        "test_name": "large_test_12",
        "expected_values_table_1": [
            "Expense", "Total expenses", "Difference (%)", "-"
        ],
    },
    {
        "test_name": "large_test_13",
        "expected_values_table_1": [
            "Expenses", "Total", "Year", "$None"
        ]
    },
    {
        "test_name": "large_test_14",
        "expected_values_table_1": [
            "None", "Cash short/extra", "Year", "$0"
        ],
        "expected_values_table_2": [
            "Category", "Total", "None", "$0"
        ]
    },
    {
        "test_name": "large_test_15",
        "expected_values_table_1": [
            "Loan No.", "Totals", "Annual Interest Rate", "None"
        ],
        "expected_values_table_2": [
            "Beginning Date", "5/1/2024", "10%", "5/1/2034"
        ],
        "expected_values_table_3": [
            "Current Monthly Payment", "$190.92", "Annual Payment", "$2,291.02"
        ]
    },
    {
        "test_name": "large_test_16",
        "expected_values_table_1": [
            "CATEGORY", "Sale of items", "SUB-TOTALS", "$0.00"
        ]
    },
    {
        "test_name": "large_test_17",
        "expected_values_table_1": [
            "COURSE ID", "PSY 101", "DURATION", "1:00"
        ]
    },
    {
        "test_name": "large_test_18",
        "expected_values_table_1": [
            "Student name", "Teacher name", "Details", "None"
        ],
        "expected_values_table_2": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_3": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_4": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_5": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_6": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_7": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_8": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_9": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
        ],
        "expected_values_table_10": [
            "Assignment #", "Marks scored", "Subject total (%)", "None"
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

