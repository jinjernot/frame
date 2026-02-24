from app.routes.qs_tool.core.blocks.paragraph import *
from app.routes.qs_tool.core.blocks.title import *
from app.routes.qs_tool.core.format.hr import *

from docx.enum.text import WD_BREAK
from docx.shared import RGBColor, Pt
import pandas as pd
import re


superscript_pattern = re.compile(r"\[(\d+)\]")

def add_formatted_run(paragraph, text, bold=False, color=None, size=None):
    """
    Adds text to a paragraph, processing [n] as superscript,
    and applies optional formatting.
    """
    text = str(text)
    parts = superscript_pattern.split(text)
    
    for i, part in enumerate(parts):
        if not part: 
            continue
            
        run = paragraph.add_run(part)
        
        if i % 2 == 1:
            run.font.superscript = True
        
        if bold:
            run.font.bold = True
        if color:
            run.font.color.rgb = color
        if size:
            run.font.size = size


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
        elif "QS-Only Processors" in xls.sheet_names:
            sheet_name_to_use = "QS-Only Processors"
        else:
            raise ValueError("Sheet 'Processors' or 'Processor' or 'QS-Only Processors' not found in the Excel file.")

        # Read the sheet with header=None to preserve all rows
        df = pd.read_excel(file.stream, sheet_name=sheet_name_to_use, engine='openpyxl', header=None)

        # Add title
        insert_title(doc, "Processors")

        # Find the main header row by looking for "Cores" keyword (typically row 5, index 4)
        main_header_idx = None
        for idx in range(min(10, len(df))):  # Search first 10 rows
            row_values = df.iloc[idx].astype(str).str.lower()
            if any('cores' in str(val) and 'p-cores' not in str(val) and 'e-cores' not in str(val) and 'lp e-cores' not in str(val) for val in row_values):
                main_header_idx = idx
                break
        
        if main_header_idx is None:
            # Fallback to row 5 (index 4)
            main_header_idx = 4
        
        # Get main header row and check if there's a sub-header row below it
        main_header_row = df.iloc[main_header_idx]
        sub_header_idx = main_header_idx + 1
        sub_header_row = df.iloc[sub_header_idx] if sub_header_idx < len(df) else None
        
        # Build structure to track main headers, sub-headers, and columns
        header_structure = []
        
        for col_idx in range(len(main_header_row)):
            main_val = str(main_header_row.iloc[col_idx]) if not pd.isna(main_header_row.iloc[col_idx]) else ""
            sub_val = str(sub_header_row.iloc[col_idx]) if sub_header_row is not None and not pd.isna(sub_header_row.iloc[col_idx]) else ""
            
            # Track both main and sub headers
            header_structure.append({
                'main': main_val if main_val != "nan" else "",
                'sub': sub_val if sub_val != "nan" else "",
                'col_idx': col_idx
            })
        
        # Filter out completely empty columns (no main or sub header)
        header_structure = [h for h in header_structure if h['main'] or h['sub']]
        non_empty_cols = [h['col_idx'] for h in header_structure]
        
        # Extract data rows starting after sub-header row
        data_start_idx = sub_header_idx + 1 if sub_header_row is not None else main_header_idx + 1
        data_df = df.iloc[data_start_idx:, non_empty_cols]
        
        # Replace NaN with empty string
        data_df = data_df.fillna('')

        # Add the data as a table to the document
        table = doc.add_table(rows=0, cols=len(header_structure))
        table.autofit = True

        # Create two-row header structure
        # First header row (main headers)
        main_header_cells = table.add_row().cells
        
        # Track which columns need merging for "Max Turbo Frequency"
        i = 0
        while i < len(header_structure):
            main_header = header_structure[i]['main']
            sub_header = header_structure[i]['sub']
            
            # If this column has a sub-header, it's part of "Max Turbo Frequency"
            if sub_header:
                # Find consecutive columns with sub-headers to merge
                merge_start = i
                while i < len(header_structure) and header_structure[i]['sub']:
                    i += 1
                merge_end = i - 1
                
                # Merge cells for "Max Turbo Frequency"
                if merge_end > merge_start:
                    merged_cell = main_header_cells[merge_start].merge(main_header_cells[merge_end])
                    merged_cell.text = ""
                    paragraph = merged_cell.paragraphs[0]
                    add_formatted_run(paragraph, "Max Turbo Frequency", bold=True)
                else:
                    # Single cell with main header
                    cell = main_header_cells[merge_start]
                    cell.text = ""
                    paragraph = cell.paragraphs[0]
                    add_formatted_run(paragraph, main_header if main_header else "Max Turbo Frequency", bold=True)
            else:
                # No sub-header, use main header
                cell = main_header_cells[i]
                cell.text = ""
                paragraph = cell.paragraphs[0]
                add_formatted_run(paragraph, main_header, bold=True)
                i += 1
        
        # Second header row (sub-headers)
        sub_header_cells = table.add_row().cells
        for i, h in enumerate(header_structure):
            cell = sub_header_cells[i]
            cell.text = ""
            paragraph = cell.paragraphs[0]
            # Use sub-header if it exists, otherwise leave empty (will be merged vertically with main header)
            if h['sub']:
                add_formatted_run(paragraph, h['sub'], bold=True)
            else:
                # Merge this cell vertically with the cell above
                cell.merge(main_header_cells[i])

        # Add data rows
        for row_data in data_df.values.tolist():
            # Check if the row contains 'Footnotes' - if so, stop
            if any('Footnotes' in str(cell) for cell in row_data):
                break
            
            # Skip completely empty rows
            if all(str(cell).strip() == '' for cell in row_data):
                continue
                
            row_cells = table.add_row().cells
            for i, cell_data in enumerate(row_data):
                cell = row_cells[i]
                cell_text = str(cell_data)
                
                # Clear existing paragraph content
                cell.text = "" 
                paragraph = cell.paragraphs[0]
                
                # Check if this specific cell should be bold (in first column, might be processor name)
                is_bold = False  # Can add logic here if needed
                
                # Use the helper function to add text with superscript support
                add_formatted_run(paragraph, cell_text, bold=is_bold)

                    
        doc.add_paragraph()
        # Process Footnotes if available
        # Find the row containing "Footnotes" text
        footnotes_row_idx = None
        for idx in range(len(df)):
            row_values = df.iloc[idx].astype(str)
            if any('footnote' in val.lower() for val in row_values):
                footnotes_row_idx = idx
                break
        
        if footnotes_row_idx is not None:
            footnotes_data = df.iloc[footnotes_row_idx + 1:]  
            footnotes_data = footnotes_data.dropna(how='all')  
            
            formatted_footnotes = []
            
            # Define the blue color
            footnote_color = RGBColor(0, 0, 153)

            # Iterate over rows of footnotes_data and add them to the document
            for _, row in footnotes_data.iterrows():
                # Get first two columns (column B typically has the footnote text)
                if len(row) < 2:
                    continue
                    
                col_a = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
                col_b = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else ""
                
                # Skip if both columns are empty
                if not col_a and not col_b:
                    continue
                
                # Skip the "Footnote" header row
                if col_a.lower() == 'footnote' or col_a.lower() == 'footnotes':
                    continue
                
                # Check if col_a contains "footnote" with a number
                if 'footnote' in col_a.lower() and col_b:
                    # Extract number from "Footnote 2", "Footnote 3", etc.
                    footnote_number_str = ''.join(filter(str.isdigit, col_a))
                    if footnote_number_str:
                        footnote_number = int(footnote_number_str)
                        # Format as "2. Text from column B"
                        formatted_text = f"{footnote_number}. {col_b}"
                        formatted_footnotes.append(formatted_text)
            
            # Add all footnotes to a single paragraph with line breaks
            if formatted_footnotes:
                footnote_paragraph = doc.add_paragraph()
                for index, footnote_text in enumerate(formatted_footnotes):
                    # Process the text for superscripts
                    add_formatted_run(footnote_paragraph, footnote_text, color=footnote_color)
                    
                    # Add line break between footnotes (but not after the last one)
                    if index < len(formatted_footnotes) - 1:
                        footnote_paragraph.add_run().add_break(WD_BREAK.LINE)


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