#!/bin/bash
TEST_NAME=simple_test_4

# Define the required filenames
EXCEL_FILENAME="$TEST_NAME/$TEST_NAME.xlsx"
CSV_FILENAME="$TEST_NAME/$TEST_NAME.csv"
TXT_FILENAME="$TEST_NAME/$TEST_NAME.txt"
IMAGE_PATH="$TEST_NAME/$TEST_NAME.png"

# Check if TXT_FILENAME exists, generate if it does not
if [ ! -f "$TXT_FILENAME" ]; then
    echo "File $TXT_FILENAME does not exist. Generating..."
    python3 excel_process_1.py $EXCEL_FILENAME $CSV_FILENAME $TXT_FILENAME $IMAGE_PATH > $TXT_FILENAME
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
python3 json_processor.py $TXT_FILENAME > $TEST_NAME/$TEST_NAME\_json.txt

# Check if the second script ran successfully
if [ $? -ne 0 ]; then
    echo "json_processor.py encountered an error."
    exit 1
fi

# Run the third Python script with the necessary arguments
echo "Running excel_process_2.py..."
python3 excel_process_2.py $EXCEL_FILENAME $TEST_NAME/$TEST_NAME\_json.txt #> $TEST_NAME/$TEST_NAME\_output.txt

# Check if the third script ran successfully
if [ $? -ne 0 ]; then
    echo "excel_process_2.py encountered an error."
    exit 1
fi

echo "All scripts ran successfully."