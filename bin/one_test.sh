#!/bin/bash

TEST_NAME="simple_test_4"

echo "Running test $TEST_NAME..."

# Define the required filenames
EXCEL_FILENAME="$TEST_NAME/$TEST_NAME.xlsx"
CSV_FILENAME="$TEST_NAME/$TEST_NAME.csv"
TXT_FILENAME="$TEST_NAME/$TEST_NAME.txt"
IMAGE_PATH="$TEST_NAME/$TEST_NAME.png"

# Run the first Python script with the necessary arguments
echo "Running excel_process_1.py..."
python3 excel_process_1.py $EXCEL_FILENAME $CSV_FILENAME $TXT_FILENAME $IMAGE_PATH > $TXT_FILENAME

# Check if the first script ran successfully
if [ $? -ne 0 ]; then
    echo "excel_process_1.py encountered an error for $TEST_NAME"
    exit 1
fi

# Run the second Python script with the necessary arguments
echo "Running json_processor.py..."
    python3 json_processor.py $TXT_FILENAME > $TEST_NAME/$TEST_NAME\_json.txt

    rm -f $TXT_FILENAME

# Check if the second script ran successfully
if [ $? -ne 0 ]; then
    echo "json_processor.py encountered an error for $TEST_NAME"
    exit 1
fi

# Run the third Python script with the necessary arguments
echo "Running excel_process_2.py..."
python3 excel_process_2.py $EXCEL_FILENAME $TEST_NAME/$TEST_NAME\_json.txt > $TEST_NAME/$TEST_NAME\_output.txt

# Check if the third script ran successfully
if [ $? -ne 0 ]; then
    echo "excel_process_2.py encountered an error for $TEST_NAME"
    exit 1
        fi

echo "$TEST_NAME ran successfully."


