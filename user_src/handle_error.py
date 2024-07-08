import sys

def main():
    if len(sys.argv) < 2:
        print("No file provided.")
        return

    file_path = sys.argv[1]

    try:
        with open(file_path, 'r') as file:
            first_line = file.readline().strip()
            first_word = first_line.split()[0] if first_line else ""

            if first_word == "Error:":
                print("find_all_tables.py")
            else:
                print("json_processor.py")
    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()