This script will be used to process the excel file and extract all the tables from it.
Here's how it will work.

1. The script will first run and directly extract all the row names, rows, and column headers from the excel file and give back JSON.
2. If there is only one table present, it will run with the original method and return the table.
3. If there are multiple tables to extract, it will run with the following method:
    - First, get all the column headers for each table from the file
        - Use a BFS to find the strongest connected sequence with all column headers
    - Then, form the table with just the column headers. Add coordinates of column headers to global coords found list.
        - Store the coordinates of the column headers, the names of the column headers, and the names of rows/row names.
    - Loop through each table starting from the first. Do a cell-wise search (only within the boundaries of the column headers)
      to find the last occurring value for a row/row name that belongs to the table.
        - Start from the bottom-most column headers row, and move downwards until you reach the boundary row of the 
        next table's column header. Record the overall last-occurring value for the table and form the corners of the table using
        the row of this last-occurring value.
4. Ask the user what table they want to extract
    - If they want to extract one table, have GPT generate the JSON of the table's column headers, row names and rows.
    - Find the closest matching table (based on keyword match) across all tables extracted (based only on column headers)
    and return to user.
    - Ask user if this is the right table, if not then go to "Else" below.
    - Elif users says yes but row identifiers/names are missing, go back to table and extract row names. Return table adjusted with boundaries.
    - Else (if table not found),
        - Give user all the tables
    