from app.routes.qs_tool.core.format.table import table_column_widths
from app.routes.qs_tool.core.blocks.paragraph import *
from app.routes.qs_tool.core.blocks.title import *
from app.routes.qs_tool.core.format.hr import *
from docx.shared import Inches, Pt, RGBColor

from docx.enum.text import WD_BREAK
import pandas as pd

def options_section(doc, file):
    """Options QS Only Section"""

    try:
        # Load xlsx
        df = pd.read_excel(file.stream, sheet_name='QS-Only Options', engine='openpyxl')

        # Add title: Options
        insert_title(doc, "OPTIONS")
        
        # Removed the "Privacy panel is only available on select models." disclaimer
        
        start_col_idx = 0
        end_col_idx = 2
        start_row_idx = 2  # Changed from 3 to 2 to include the header row
        end_row_idx = 299

        data_range = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1]
        data_range = data_range.dropna(how='all')

        num_rows, num_cols = data_range.shape
        table = doc.add_table(rows=0, cols=num_cols) 

        table_row_index = 0

        for row_idx in range(num_rows):
            row = data_range.iloc[row_idx]
            if row.isna().all(): 
                break
            is_section_header = (
                row_idx > 0 and 
                not pd.isna(row[0]) and 
                pd.isna(row[1]) and 
                pd.isna(row[2])
            )

            if is_section_header:
                table.add_row()
                table_row_index += 1

            table.add_row()

            for col_idx in range(num_cols):
                value = row[col_idx]
                
                cell = table.cell(table_row_index, col_idx) 
                
                if not pd.isna(value):
                    run = cell.paragraphs[0].add_run(str(value))
                    
                    if row_idx == 0 or col_idx == 0:
                        run.font.bold = True
            
            table_row_index += 1
                            
        table_column_widths(table, (Inches(2), Inches(4), Inches(2)))

        # Insert HR
        insert_horizontal_line(doc.add_paragraph(), thickness=3)

        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)

