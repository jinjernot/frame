from app.routes.qs_tool.core.blocks.paragraph import *
from app.routes.qs_tool.core.blocks.title import *
from app.routes.qs_tool.core.format.hr import *

from docx.enum.text import WD_BREAK
from docx.shared import RGBColor
import pandas as pd
import docx

def processors_section(doc, file):
    """Processors techspecs section"""

    try:
        # Load Excel file and check if "Processors" or "Processor" sheet exists
        xls = pd.ExcelFile(file.stream, engine='openpyxl')
        sheet_name_to_use = None
        
        if "Processors" in xls.sheet_names:
            sheet_name_to_use = "Processors"
        elif "Processor" in xls.sheet_names:
            sheet_name_to_use = "Processor"
        else:
            raise ValueError("Sheet 'Processors' or 'Processor' not found in the Excel file.")

        # Read the sheet
        df = pd.read_excel(file.stream, sheet_name=sheet_name_to_use, engine='openpyxl')

        # Add title
        insert_title(doc, "Processors")

        # Dynamically fetch the column names from the third row (index 2) that have data
        third_row = df.iloc[3]  # Selecting the third row (index 2, but 1-based)

        # Filter to keep columns that have data in the third row
        filtered_df = df.loc[:, ~third_row.isna()]

        # Remove the rows
        filtered_df = filtered_df.iloc[3:]

        # Replace "NaN" string values with an empty string
        filtered_df = filtered_df.fillna('')

        # Convert filtered dataframe to a list of lists (data) for the table
        data = filtered_df.values.tolist()

        # Add the data as a table to the document
        table = doc.add_table(rows=1, cols=len(data[0]))
        table.autofit = True

        # Add table data
        for row in data:
            # Check if the row is empty
            if 'Footnotes' in row:
                break  # Exit the loop if footnote row is reached
            
            row_cells = table.add_row().cells
            for i, cell in enumerate(row):
                row_cells[i].text = str(cell)
                
                # Bold the text if the cell contains "Processor Family"
                if str(cell) == "Processor Family":
                    for paragraph in row_cells[i].paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True

        # Remove the first row
        if len(table.rows) > 1:  # Ensure there are rows to delete
            table.rows[0]._element.getparent().remove(table.rows[0]._element)

        for cell in table.rows[0].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                        
        doc.add_paragraph()

        # --- UPDATED FOOTNOTE LOGIC ---
        # Process Footnotes if available
        footnotes_index = df[df.eq('Footnotes').any(axis=1)].index.tolist()
        if footnotes_index:
            footnotes_index = footnotes_index[0]  
            footnotes_data = df.iloc[footnotes_index + 1:]  
            footnotes_data = footnotes_data.dropna(how='all')  
            
            # Create a single paragraph for all footnotes
            footnote_paragraph = doc.add_paragraph()
            first_footnote = True

            # Iterate over rows of footnotes_data and add them to the document
            for _, row in footnotes_data.iterrows():
                row_values = row.dropna().tolist()
                
                if not row_values:
                    continue # Skip empty rows

                if not first_footnote:
                    # Add a line break before this new footnote
                    footnote_paragraph.add_run().add_break(WD_BREAK.LINE)
                
                first_footnote = False

                # Logic adapted from insert_list
                first_cell = str(row_values[0])
                
                if 'footnote' in first_cell.lower() and len(row_values) > 1:
                    # Case: "Footnote 02", "Text..."
                    footnote_number_str = ''.join(filter(str.isdigit, first_cell))
                    if footnote_number_str:
                        footnote_number = int(footnote_number_str)
                        footnote_text = str(row_values[1])
                        
                        run = footnote_paragraph.add_run(f"{footnote_number}. {footnote_text}")
                        run.font.color.rgb = RGBColor(0, 0, 153)
                    else:
                        # Fallback: "Footnote", "Text..."
                        run = footnote_paragraph.add_run(" - ".join(map(str, row_values)))
                        run.font.color.rgb = RGBColor(0, 0, 153)

                elif len(row_values) == 1:
                    # Case: (empty), "Multicore text..."
                    run = footnote_paragraph.add_run(str(row_values[0]))
                    run.font.color.rgb = RGBColor(0, 0, 153)
                
                else:
                    # Fallback for other formats
                    run = footnote_paragraph.add_run(" - ".join(map(str, row_values)))
                    run.font.color.rgb = RGBColor(0, 0, 153)
        # --- END OF UPDATED LOGIC ---

        # Insert Horizontal Line
        insert_horizontal_line(doc.add_paragraph(), thickness=3)

        # Insert Page Break
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    except Exception as e:
        error_msg = f"An error occurred: {e}"
        print(error_msg)  # Log error to console

        # Add error message to Word document in red bold text
        error_paragraph = doc.add_paragraph()
        error_run = error_paragraph.add_run(error_msg)
        error_run.bold = True
        error_run.font.color.rgb = RGBColor(255, 0, 0) 

        return str(e)