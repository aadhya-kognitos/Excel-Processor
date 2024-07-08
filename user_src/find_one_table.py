import openpyxl
import openai
import base64
import csv
import sys
from dotenv import load_dotenv
import pdb

class ExcelPreProcessor:
    def __init__(self, excel_filename, csv_filename, txt_filename, image_path, user_query_txt):
        self.excel_filename = excel_filename
        self.csv_filename = csv_filename
        self.txt_filename = txt_filename
        self.txt_in_csv = None
        self.image_path = image_path
        self.user_query_txt = user_query_txt
        self.user_query = None
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
        with open(self.txt_filename, 'r') as f:
            self.txt_in_csv = f.readlines()
    
    def write_user_query_txt(self):
        with open(self.user_query_txt, 'r') as f:
            self.user_query = f.read()


    def encode_image(self):
        with open(self.image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    def run(self):
        self.process_excel_to_csv()
        self.write_csv_to_txt()
        system_prompt = self.system_prompt()
        self.write_user_query_txt()
        user_prompt = self.user_prompt(self.txt_in_csv, self.user_query)
        prompt_1 = self.query_openai(system_prompt, user_prompt)
        print(prompt_1)

    def system_prompt(self):
        return f"""
    The following image contains tables in a spreadsheet. For the table that the user specifies,
    provide me the data in the single table in JSON format. I want you to extract the column headers,
    row names, and last row of row values in the table separately.

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

    TABLE 4
            Gross Revenue | Net Spending | Net Profit
    Tech  
    Sports
    Arts
    Entertainment

  And the user prompt is "give me the inventory table"
    the output you would give is simply as follows:

    {{"Table 1":{{
    "Columns": ["INVENTORY", "QUANTITY", "PRICE", "TOTAL"],
    "Row Names: [],
    "Last Row": ["Oranges", "15", "$1.50", "22.50"]
    }}}},
Or if the user prompt is "give me the table with triathlete and swim columns"
the output you would give is simply as follows:

    "Table 2": {{
    "Columns": ["TRIATHLETE", "Run", "Swim", "Cycle"],
    "Row Names: ["Athlete 1", "Athlete 2"],
    "Last Row": ["Mark", "20", "1.50", "6.00"]
    }},

    Or if the user prompt is "provide me the table with car data"
    Since no table has car data or car-related columns, you will simply say
    "Error: Table not found." along with the entire user prompt provided.

    No need to place the ```json in the output.
    Please use only curly braces, as in standard JSON format. Don't use square brackets.

    If a cell value is empty, no need to provide it in the JSON representation.
    The Last Row should always have values, not be empty. If it appears that the last row is empty in the image,
    use the closest row above that that is not empty belonging to the same table.

    When providing the row names, do not include the last row values as part
    of the row names. Make sure to double check that the last row values and the row names are distinct.

    
    Here are some steps to identify a table/tables
    A table is a section of data that is formatted in a way that is easy to read and understand. If there are multiple tables in a spreadsheet,
    each table will have its own section, which will most likely be colored differently than the background of the spreadsheet. Tables also will have
    a relatively close density of values/colors while empty space will not contain this density.

    Assume the column headers are all on one row, there are not above or below each other.
    In cases where column headers span multiple, please provide only the columns on the bottom-most row you see.

    Make sure to read the tables carefully, line by line, from left to right and accurately extract the given values from each table.

    Don't add any additional text apart from the desired values from each table.
    Make sure to defer to the csv values for each table, rather than the image values.
    Be sure to only provide ONE table.
    If you cannot find the desired table, throw the Error message specified above.
    Please do not select a table if it does not contain the desired keyword the user is looking for.
    Note that sometimes column headers can span multiple rows.
    Make sure the JSON is correctly formatted, and ONLY provide the JSON. Don't add any additional text.

   The user will always provide you an image and csv file to work with.
"""
    
    def user_prompt(self, csv_text, user_query):
        return f"""Heres the csv: ${csv_text}. The user query is: {user_query}. Here's an image:"""

    def query_openai(self, system_prompt, user_prompt):
        #pdb.set_trace()
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
            top_p=0.0000001
        )
        returned_text = completion.choices[0].message.content
        return returned_text

if __name__ == "__main__":
    if len(sys.argv) != 6:
        print("Usage: python find_one_table.py <excel_filename> <csv_filename> <txt_filename> <image_path> <user_query>")
        sys.exit(1)

    processor = ExcelPreProcessor(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    processor.run()

