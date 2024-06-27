#!/bin/bash

# Check for the correct number of arguments
if [ "$#" -ne 1 ]; then
    echo "Usage: $0 <test_name>"
    exit 1
fi

TEST_NAME="$1"

# Define the required filenames
EXCEL_FILENAME="tests/$TEST_NAME/$TEST_NAME.xlsx"
CSV_FILENAME="tests/$TEST_NAME/$TEST_NAME.csv"
TXT_FILENAME="tests/$TEST_NAME/$TEST_NAME.txt"
IMAGE_PATH="tests/$TEST_NAME/$TEST_NAME.png"

# Check if TXT_FILENAME exists, generate if it does not
if [ ! -f "$TXT_FILENAME" ]; then
    echo "File $TXT_FILENAME does not exist. Generating..."
    python3 src/excel_process_1.py $EXCEL_FILENAME $CSV_FILENAME $TXT_FILENAME $IMAGE_PATH > $TXT_FILENAME
    if [ $? -ne 0 ]; then
        echo "excel_process_1.py encountered an error."
        exit 1
    fi
else
    echo "File $TXT_FILENAME already exists. Skipping generation."
fi

# Continue with the rest of the script

# Run the second Python script with the necessary arguments
echo "Running json_processor.py..."
python3 src/json_processor.py $TXT_FILENAME > tests/$TEST_NAME/$TEST_NAME\_json.txt 

# Check if the second script ran successfully
if [ $? -ne 0 ]; then
    echo "json_processor.py encountered an error."
    exit 1
fi

# Run the third Python script with the necessary arguments
echo "Running excel_process_2.py..."
python3 src/debug_excel_process_2.py $EXCEL_FILENAME tests/$TEST_NAME/$TEST_NAME\_json.txt > tests/$TEST_NAME/$TEST_NAME\_output.txt

# Check if the third script ran successfully
if [ $? -ne 0 ]; then
    echo "excel_process_2.py encountered an error."
    exit 1
fi

echo "All scripts ran successfully."