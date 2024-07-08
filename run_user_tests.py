import os
import sys
import subprocess

def run_command(command):
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running command: {command}")
        print(f"Error details: {result.stderr}")
        sys.exit(1)
    return result.stdout

def run_test(test_name):
    test_dir = f"user_tests/{test_name}"

    # Remove all files in the directory apart from <test_name>.png, <test_name>.xlsx, and user_query.txt
    for root, dirs, files in os.walk(test_dir):
        for file in files:
            if file not in [f"{test_name}.png", f"{test_name}.xlsx", "user_query.txt"]:
                os.remove(os.path.join(root, file))

    # Define the required filenames
    excel_filename = f"{test_dir}/{test_name}.xlsx"
    csv_filename = f"{test_dir}/{test_name}.csv"
    txt_filename = f"{test_dir}/{test_name}.txt"
    image_path = f"{test_dir}/{test_name}.png"
    user_query_txt = f"{test_dir}/user_query.txt"
    user_query_output = f"{test_dir}/user_query_output.txt"

    # Run the table one finder
    print("Running find_one_table.py...")
    run_command(f"python3 user_src/find_one_table.py {excel_filename} {csv_filename} {txt_filename} {image_path} {user_query_txt} > {user_query_output}")

    # Run the error handler script with the necessary arguments
    print("Running handle_error.py...")
    run_command(f"python3 user_src/handle_error.py {user_query_output} > {test_dir}/user_query_error.txt")

    # Get the content of user query error and figure out what to do next
    with open(f"{test_dir}/user_query_error.txt", 'r') as file:
        user_query_error = file.read().strip()

    if user_query_error == "find_all_tables.py":
        print("Running find_all_tables.py...")
        run_command(f"python3 user_src/find_all_tables.py {excel_filename} {csv_filename} {txt_filename} {image_path} > {txt_filename}")
        print("Running json_processor.py...")
        run_command(f"python3 user_src/json_processor.py {txt_filename} > {test_dir}/{test_name}_json.txt")
        os.remove(txt_filename)
        print("Running process_tables.py...")
        run_command(f"python3 user_src/process_tables.py {excel_filename} {test_dir}/{test_name}_json.txt > {test_dir}/{test_name}_output.txt")
    elif user_query_error == "json_processor.py":
        with open(txt_filename, 'w') as file:
            with open(user_query_output, 'r') as output_file:
                file.write(output_file.read())
        print("Running json_processor.py...")
        run_command(f"python3 user_src/json_processor.py {txt_filename} > {test_dir}/{test_name}_json.txt")
        os.remove(txt_filename)
        print("Running process_tables.py...")
        run_command(f"python3 user_src/process_tables.py {excel_filename} {test_dir}/{test_name}_json.txt > {test_dir}/{test_name}_output.txt")

    print(f"{test_name} ran successfully.")

def main():
    for i in range(1, 30):
        test_name = f"test_{i}"
        print(f"Running {test_name}...")
        run_test(test_name)

if __name__ == "__main__":
    main()