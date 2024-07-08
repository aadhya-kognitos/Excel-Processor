#!/bin/bash

# Check for the correct number of arguments
if [ "$#" -ne 1 ]; then
    echo "Usage: $0 <test_name>"
    exit 1
fi

TEST_NAME="$1"

# Remove all files in the directory apart from <test_name>.png, <test_name>.xlsx, and user_query.txt
find "user_tests/$TEST_NAME" -type f ! -name "$TEST_NAME.png" ! -name "$TEST_NAME.xlsx" ! -name "user_query.txt" -exec rm -f {} +


# Define the required filenames
EXCEL_FILENAME="user_tests/$TEST_NAME/$TEST_NAME.xlsx"
CSV_FILENAME="user_tests/$TEST_NAME/$TEST_NAME.csv"
TXT_FILENAME="user_tests/$TEST_NAME/$TEST_NAME.txt"
IMAGE_PATH="user_tests/$TEST_NAME/$TEST_NAME.png"
USER_QUERY_TXT="user_tests/$TEST_NAME/user_query.txt"
USER_QUERY_OUTPUT="user_tests/$TEST_NAME/user_query_output.txt"

# Run the table one finder
echo "Running find_one_table.py..."
python3 user_src/find_one_table.py $EXCEL_FILENAME $CSV_FILENAME $TXT_FILENAME $IMAGE_PATH $USER_QUERY_TXT > $USER_QUERY_OUTPUT

# Check if the first script ran successfully
if [ $? -ne 0 ]; then
    echo "excel_process_1.py encountered an error for $TEST_NAME"
    echo "Error details:"
    cat $USER_QUERY_OUTPUT
    exit 1
fi

# Run the error handler script with the necessary arguments
echo "Running handle_error.py..."
python3 user_src/handle_error.py $USER_QUERY_OUTPUT > user_tests/$TEST_NAME/user_query_error.txt

#Get the content of user query error and figure out what to do next
USER_QUERY_ERROR=$(cat user_tests/$TEST_NAME/user_query_error.txt)

if [ "$USER_QUERY_ERROR" == "find_all_tables.py" ]; then
    echo "Running find_all_tables.py..."
    python3 user_src/find_all_tables.py $EXCEL_FILENAME $CSV_FILENAME $TXT_FILENAME $IMAGE_PATH > $TXT_FILENAME
elif [ "$USER_QUERY_ERROR" == "json_processor.py" ]; then
    cat $USER_QUERY_OUTPUT > $TXT_FILENAME
fi

# Run the json_processor
echo "Running json_processor.py..."
    python3 user_src/json_processor.py $TXT_FILENAME > user_tests/$TEST_NAME/$TEST_NAME\_json.txt

    rm -f $TXT_FILENAME

# Check if the json_processor ran successfully
if [ $? -ne 0 ]; then
    echo "json_processor.py encountered an error for $TEST_NAME"
    exit 1
fi

# Run the table processor
echo "Running process_tables.py..."
python3 user_src/exp_process_tables.py $EXCEL_FILENAME user_tests/$TEST_NAME/$TEST_NAME\_json.txt > user_tests/$TEST_NAME/$TEST_NAME\_output.txt

# Check if the third script ran successfully
if [ $? -ne 0 ]; then
    echo "process_tables.py encountered an error for $TEST_NAME"
    exit 1
        fi

echo "$TEST_NAME ran successfully."


