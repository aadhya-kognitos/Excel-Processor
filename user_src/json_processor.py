from sys import argv
import pdb

def read_json(json_file):
    """ Read contents of txt file with 
    JSON format and stores in JSON object."""
    open_bracket_count = 0
    close_bracket_count = 0
    with open(json_file, 'r') as json_file:
        string = json_file.read()
    json_filtered_string = ""
    for char in string:
        if char == '{':
            open_bracket_count += 1
        elif char == '}':
            close_bracket_count += 1
        if open_bracket_count == close_bracket_count and open_bracket_count > 1:
            json_filtered_string += char
            return json_filtered_string
        json_filtered_string += char

def main():
    json_file = argv[1]
    json_object = read_json(json_file)
    print(json_object)

if __name__ == "__main__":
    main()

