import openpyxl
import openai
import base64
import csv
import sys
from dotenv import load_dotenv

class ExcelPreProcessor:
    def __init__(self, excel_filename, csv_filename, txt_filename, image_path):
        self.excel_filename = excel_filename
        self.csv_filename = csv_filename
        self.txt_filename = txt_filename
        self.txte = None
        self.image_path = image_path
        load_dotenv()
        self.client = openai.OpenAI()

    def process_excel_to_csv(self):
        wb = openpyxl.load_workbook(self.excel_filename, data_only=True)
        ws = wb.active
        with open(self.csv_filename, 'w', newline="") as csvfile:
            csvwriter = csv.writer(csvfile)
            for row in ws.iter_rows(values_only=True):
                csvwriter.writerow(row)

    def write_csv_to_txt(self):
        with open(self.txt_filename, 'w') as f:
            with open(self.csv_filename) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=',')
                for row in csv_reader:
                    f.write(" ".join(row))
                    f.write("\n")


    def encode_image(self):
        with open(self.image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    def run(self):
        self.process_excel_to_csv()
        self.write_csv_to_txt()
        system_prompt = self.system_prompt()
        user_prompt = self.user_prompt(self.txt_filename)
        prompt_1 = self.query_openai(system_prompt, user_prompt)
        add_on = f"""Given this initial response from GPT: {prompt_1}, I want you to go in and take a look at
        the csv again: {self.txt_filename}. 
        When looking at the csv, here are the steps you should follow
        1. For each column header list, row name list, or rows list, find the corresponding row in the CSV
        that is a close match or match for that list.
        2. Leave values that are exact matches in that row. If you see a value in the original response JSON
        that does not match a CSV value, please update that value with the CSV value.
        3. Keep going till you reach the end of the table."""
        prompt_2 = self.query_openai(system_prompt, add_on)
        print(prompt_2)

    def system_prompt(self):
        return f"""
    The following image contains tables in a spreadsheet. For each table you see,
    provide me the data in each table in JSON format. I want you to extract the column headers,
    row names, and last row of row values in each table separately.

    For example if you see the following tables (note the following tables are just examples, not the actual tables in the image.
    There may be more or fewer tables in the image than the tables provided below):

    TABLE 1
    INVENTORY | QUANTITY | PRICE | TOTAL
    ----------|----------|-------|--------
    Apples    | 10       | $2.00  | $20.00
    Oranges   | 15       | $1.50  | $22.50

            TABLE 2
            TRIATHLETE | Run | Swim | Cycle
            ----------|----------|-------|--------
    Athlete 1 Aaron     | 20       | 2.00 | 20.00
    Athlete 2 Mark      | 20       | 1.50 | 6.00

    TABLE 3
    Company | Stock
    ----------|----------|
    Google   | $1000
    Nvidia   | $2000


    the output you would give is simply as follows:

    {{"Table 1":{{
    "Columns": ["INVENTORY", "QUANTITY", "PRICE", "TOTAL"],
    "Row Names: [],
    "Last Row": ["Oranges", "15", "$1.50", "22.50"]
    }},
    "Table 2": {{
    "Columns": ["TRIATHLETE", "Run", "Swim", "Cycle"],
    "Row Names: ["Athlete 1", "Athlete 2"],
    "Last Row": ["Mark", "20", "1.50", "6.00"]
    }},
    "Table 3": {{
    "Columns": ["Company", "Stock"],
    "Row Names: [],
    "Last Row": ["Nvidia", "$2000"]
    }}
    }}
    
    No need to place the ```json in the output.
    Please use only curly braces, as in standard JSON format. Don't use square brackets.
    When processing tables, please look at the CSV values I provide first, then look at the image.
    The CSV values will tell you the exact value of data at a certain location.

    If the last row has some kind of string value, please include at least one of that rather than numerical values.
    If a cell value is empty, no need to provide it in the JSON representation.
    The Last Row should always have values, not be empty. If it appears that the last row is empty in the image,
    use the closest row above that that is not empty belonging to the same table.
    
    Here are some steps to identify a table/tables
    A table is a section of data that is formatted in a way that is easy to read and understand. If there are multiple tables in a spreadsheet,
    each table will have its own section, which will most likely be colored differently than the background of the spreadsheet. Tables also will have
    a relatively close density of values/colors while empty space will not contain this density.

    Make sure to read the tables carefully, line by line, from left to right and accurately extract the given values from each table.

    Don't add any additional text apart from the desired values from each table.
    Make sure to defer to the csv values for each table, rather than the image values.
    Please make sure to capture ALL the tables in the spreadsheet, and ensure each table has its own JSON representation.
    However, be sure to check that you are not repeating the same table more than once.
    Note that sometimes column headers can span multiple rows.
    Make sure the JSON is correctly formatted, and ONLY provide the JSON. Don't add any additional text.
"""
    
    def user_prompt(self, csv_text):
        return f"""
        Consider the following txt representation of the spreadsheet in csv format before parsing the image: ${csv_text}.
        The csv text generally contains the true values present in the spreadsheet. For each table, use the csv values
        rather than the image values.
        For instance, if the csv contains "1000" for a cell, and the image contains "$1,000", use "1000"
        For instance, if the csv contains "850" for a 
        cell, and the image contains "850.00", use "850"
        Image is provided in base64 format: """

    def query_openai(self, system_prompt, user_prompt):
        completion = self.client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", 
                 "content": [
                    {"type": "text", 
                     "text": system_prompt}]
                },
                {"role": "user", 
                 "content": [
                     {"type": "text", 
                      "text": user_prompt},
                      {"type": "image_url", 
                      "image_url": 
                      {
                          "url": f"data:image/jpeg;base64,{self.encode_image()}"
                      }
                      }
                    
                    ]
                 }
            ],
            temperature=0
        )
        returned_text = completion.choices[0].message.content
        return returned_text

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Usage: python excel_process_1.py <excel_filename> <csv_filename> <txt_filename> <image_path>")
        sys.exit(1)

    processor = ExcelPreProcessor(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    processor.run()