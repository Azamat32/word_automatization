import os
from docx import Document
from openpyxl import load_workbook
import psycopg2  # Import the library for working with PostgreSQL
from docx.shared import Pt
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
# Set up a connection to the PostgreSQL database
conn = psycopg2.connect(
    dbname="gdp_test",
    user="postgres",
    password="daniyarfgh",
    host="localhost",
    port="5432"
)

# Path to the directory where Excel files are stored
tables_directory = "./tables/"

# Path to the output document
output_document_path = "./word/test.docx"

# Create a new Word document
new_document = Document()

# Create a cursor for executing SQL queries
cursor = conn.cursor()

# Get topics from the database
cursor.execute("SELECT * FROM api_topic")
topics = cursor.fetchall()

# Iterate through topics
for topic in topics:
    topic_id, topic_name = topic
    new_document.add_heading(topic_name, level=1)

    # Get economic indexes for this topic
    cursor.execute("SELECT * FROM api_economic_index WHERE macro_topic_id = %s", (topic_id,))
    economic_indexes = cursor.fetchall()

    # Iterate through economic indexes
    for economic_index in economic_indexes:
        economic_index_id, economic_index_name, *some = economic_index
        new_document.add_heading(economic_index_name, level=2)

        # Get table names for this economic index
        cursor.execute("SELECT path FROM api_table WHERE macro_economic_index_id = %s", (economic_index_id,))
        table_names = cursor.fetchall()

        # Iterate through table names
        for table_name in table_names:
            table_name = table_name[0]  # Extract the table name from the tuple
            excel_file_path = os.path.join(tables_directory, f"{table_name}.xlsx")
            print(table_name)
            try:
                # Load the Excel file
                excel_workbook = load_workbook(excel_file_path, data_only=True)
                excel_sheet = excel_workbook.active

                # Check if the Excel table is not empty
                if excel_sheet.max_row > 0 and excel_sheet.max_column > 0:
                    # Create a new table in the Word document
                    num_rows = excel_sheet.max_row
                    num_cols = excel_sheet.max_column
                    table = new_document.add_table(rows=num_rows, cols=num_cols)
                    
                    # Apply cell formatting and add borders to paragraphs (use your provided code)
                    for row in table.rows:
                        for cell in row.cells:
                            cell.paragraphs[0].style.font.size = Pt(8)
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.size = Pt(8)
                            header_shading_elm = parse_xml(r'<w:shd {} w:fill="006FC0"/>'.format(nsdecls('w')))
                            cell._tc.get_or_add_tcPr().append(header_shading_elm)
                    
                    # Fill the table with data from Excel
                    for row_idx, row in enumerate(excel_sheet.iter_rows()):
                        for col_idx, cell in enumerate(row):
                            table.cell(row_idx, col_idx).text = str(cell.value)

                    # Add an empty paragraph after the table for separation
                    new_document.add_paragraph()
                else:
                    # If the Excel table is empty, add a corresponding message
                    new_document.add_paragraph("Empty table: " + table_name)
                    new_document.add_paragraph()  # Add an empty paragraph after the message

            except FileNotFoundError:
                # If the Excel file is not found, add a corresponding message
                new_document.add_paragraph("Excel file not found: " + table_name)
                new_document.add_paragraph()  # Add an empty paragraph after the message

            except Exception as e:
                # Handle other exceptions if necessary
                print(f"Error processing table {table_name}: {str(e)}")

# Save the Word document
new_document.save(output_document_path)

# Close the database connection
cursor.close()
conn.close()
