import os
from docx import Document
from openpyxl import load_workbook
import psycopg2
from docx.shared import Pt
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def set_text_color_to_white(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 255, 255)

def process_table(new_document, table_name, tables_directory):
    excel_file_path = os.path.join(tables_directory, f"{table_name}.xlsx")
    print(table_name)
    try:
        # Load the Excel file
        excel_workbook = load_workbook(excel_file_path, data_only=True)
        excel_sheet = excel_workbook.active
        if excel_sheet.max_row > 0 and excel_sheet.max_column > 0:
            num_rows = excel_sheet.max_row
            num_cols = excel_sheet.max_column
            table = new_document.add_table(rows=num_rows, cols=num_cols)
            table.style = 'Table Grid'
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    if row_idx == 0:
                        # Set text color to white for the first row
                        set_text_color_to_white(cell)
                    shading_color = "#006FC0" if row_idx == 0 else "#fff"
                    cell.paragraphs[0].style.font.size = Pt(8)
                    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{shading_color}" />')
                    cell._tc.get_or_add_tcPr().append(shading_elm)
            for row_idx, row in enumerate(excel_sheet.iter_rows()):
                for col_idx, cell in enumerate(row):
                    cell_value = cell.value if cell.value is not None else ""
                    table.cell(row_idx, col_idx).text = str(cell_value)
                    if row_idx == 0:
            # Set text color to white for the first row
                        for paragraph in table.rows[row_idx].cells[col_idx].paragraphs:
                            if paragraph.runs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(255, 255, 255)

            new_document.add_paragraph()
        else:
            new_document.add_paragraph("Empty table: " + table_name)
            new_document.add_paragraph()
    except FileNotFoundError:
        new_document.add_paragraph("Excel file not found: " + table_name)
        new_document.add_paragraph()
    except Exception as e:
        print(f"Error processing table {table_name}: {str(e)}")

def main():
    # Set up a connection to the PostgreSQL database
    conn = psycopg2.connect(
        dbname="gdp_test",
        user="postgres",
        password="daniyarfgh",
        host="localhost",
        port="5432"
    )

    # Path to the directory where Excel files are stored
    tables_directory = "./tables_test/"

    # Path to the output document
    output_document_path = "./word/BulletTestTemplate.docx"

    # Create a new Word document
    new_document = Document(output_document_path)

    # Create a cursor for executing SQL queries
    cursor = conn.cursor()

    # Fetch topics from the database
    cursor.execute("SELECT * FROM api_topic")
    topics = cursor.fetchall()

    for i, topic in enumerate(topics):
        topic_id, topic_name = topic
        topic_heading = f"{i + 1}. {topic_name}"
        topic_paragraph = new_document.add_paragraph()
        topic_run = topic_paragraph.add_run(topic_heading)
        topic_font = topic_run.font
        topic_font.name = 'Arial'
        topic_font.size = Pt(16)
        topic_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        topic_font.bold = True
        new_document.add_paragraph()
        cursor.execute("SELECT * FROM api_economic_index WHERE macro_topic_id = %s", (topic_id,))
        economic_indexes = cursor.fetchall()

        for k, economic_index in enumerate(economic_indexes):
            economic_index_id, economic_index_name, *some = economic_index
            economic_index_heading = f"{i + 1}.{k+1}.  {economic_index_name}"
            economic_index_paragraph = new_document.add_paragraph()
            economic_index_run = economic_index_paragraph.add_run(economic_index_heading)
            economic_index_font = economic_index_run.font
            economic_index_font.name = 'Arial'
            economic_index_font.size = Pt(12)
            economic_index_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            economic_index_font.color.rgb = RGBColor(0, 0, 139)

            cursor.execute("SELECT path FROM api_table WHERE macro_economic_index_id = %s", (economic_index_id,))
            table_names = cursor.fetchall()

            for table_name in table_names:
                process_table(new_document, table_name[0], tables_directory)

    new_document.save("./word/test1.docx")

    cursor.close()
    conn.close()

if __name__ == "__main__":
    main()
