# Project README

## Overview

With this program, you can extract all table coordinates present in an excel spreadsheet.

## Prerequisites

- Python 3.9
- Virtualenv

## Getting Started

### 1. Create a Python Virtual Environment and Install Requirements

First, create a virtual environment and install the necessary packages listed in `requirements.txt`.

```bash
# Create a virtual environment
python3 -m venv venv

# Activate the virtual environment
# On Windows
venv\Scripts\activate

# On Unix or MacOS
source venv/bin/activate

# Install the required packages
pip install -r requirements.txt
```

### 2. Available Tests

The project includes the following tests:

- **Large Tests**
  - large_test_1
  - large_test_2
  - large_test_3
  - large_test_4
  - large_test_5

- **Simple Tests**
  - simple_test_1
  - simple_test_2
  - simple_test_3
  - simple_test_4
  - simple_test_5

To add your own tests, create a new subdirectory entitled `large_test_(# of test)` or `small_test_(# of test)` and add an excel file entitled `large_test_(# of test).xlsx` or `small_test_(# of test).xlsx`. Also add a png screenshot of your spreadsheet entitled `large_test_(# of test).png` or `small_test_(# of test).png`.

**Ex.** To create a new test called large_test_6, create a new subdirectory entitled `large_test_6` and add an excel file entitled `large_test_6.xlsx`. Also add a png screenshot of your spreadsheet entitled `large_test_6.png`.

### 3. Running the Tests

Navigate to the `bin` folder to find the shell scripts for running the tests. Here are the available scripts and their usage:

**Make sure to run the tests in the `bin` directory.**
**Perform a `chmod +x <script_name>` to make the script executable.**

- **all_tests.sh**
  - Runs all the tests available. Terminal output will indicate if a test fails.
  - Usage: `./bin/all_tests.sh`

- **one_test.sh**
  - Runs a specific test. Terminal output will indicate if a test fails.
  - Takes one argument: the name of the test to be run (e.g., `large_test_#` or `simple_test_#`).
  - Usage: `./bin/one_test.sh <test_name>`

- **debug.sh**
  - Allows debugging of a specific test.
  - Takes one argument: the name of the test to be debugged.
  - Usage: `./bin/debug.sh <test_name>`
  - You will need to set `pdb.set_trace()` in the python programs for debugging purposes.

**Output for all tests will be in the corresponding subdirectory for a test, in the output txt file.**

Ex. for large_test_1, the output will be in large_test_1/large_test_1_output.txt.

**I am currently working on developing tests through `pytest`. The file `excel_process_tests_1.py` and `excel_process_tests_2.py` contains the tests that are currently being developed using `pytest`.**

## Current Issues

For `large_test_1`, the program does not return the first subtable found in the spreadsheet. This is due to its inability to process numerical values containing a $ or other symbol.

For `simple_test_2`, the program returns a table encompassing both of the tables in the spreadsheet. This is due to the second table having a column and row of the same value as the column + row values above it.

For `simple_test_5` the program does not return a table. This is because there is a row value with the same value as a column header.


## Error Log

**Large Test 1**
Total Tables: 10,
Error(s): First table missing (CURRENT BUDGET)

**Large Test 2**
Total Tables: 1,
Error(s): None

**Large Test 3**
Total Tables: 1,
Error(s): None

**Large Test 4**
Total Tables: 4,
Error(s): 
Table 1 (Company Name) not represented accurately (column 1 is not present)
Only 1st column present in final table

**Large Test 5**
Total Tables: 1,
Error(s): None

**Large Test 6**
Total Tables: 16,
Error(s): 
First table only extracts middle column
2nd table only extracts middle column
PROJECTED BALANCE SHEET NOT PRESENT
Entertainment table bottem left corner is wrong coordinate (but has correct value)
Table 9-11 inaccurately represented (wrong coordinates)
Gifts and Donations table has bottom left coordinate represented incorrectly
Legal table missing
Total Projected Cost only has one column

**Large Test 7**
Total Tables: 3,
Error(s):
Not split into subtables

**Large Test 8**
Total Tables: 2,
Error(s): None

**Large Test 9**
Total Tables: 1,
Error(s):
% Done column coordinate not present

**Large Test 10**
Total Tables: 3,
Error(s): None

**Large Test 11**
Total Tables: 2,
Error(s):
Missing first table

**Large Test 12**
Total Tables: 1,
Error(s): Table not found

**Large Test 13**
Total Tables: 1,
Error(s): 
No Output

**Large Test 14**
Total Tables: 2,
Error(s): 
No Output

**Large Test 15**
Total Tables: 3,
Error(s):
2nd table's first column is not showing
3rd table not showing

**Large Test 16**
Total Tables: 3,
Error(s):
Agenda table not present
Event Checklist not present
Event Categories not present

**Large Test 17**
Total Tables: 1,
Error(s): None

**Large Test 18**
Total Tables: 11,
Error(s):
All Tables have error in 2nd and 4th coordinate display


## Next Steps

I am currently in the testing stages of a new method which enables the user to upload just a spreadsheet. This eliminates the need for them to manually take screenshots. I will update here as necessary.

## Conclusion

Follow the steps above to set up your environment and run the tests.