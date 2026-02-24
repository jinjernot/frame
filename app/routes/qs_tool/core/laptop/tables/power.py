from app.routes.qs_tool.core.blocks.paragraph import *
from app.routes.qs_tool.core.blocks.title import *
from app.routes.qs_tool.core.blocks.table import *
from app.routes.qs_tool.core.format.hr import *

from docx.enum.text import WD_BREAK
import pandas as pd

def power_section(doc, file):
    """Power QS Only Section"""

    try:
        # Load xlsx
        df = pd.read_excel(file.stream, sheet_name='QS-Only Power', engine='openpyxl')

        # Add title: Power
        insert_title(doc, "POWER")
        
        paragraph = doc.add_paragraph()
        run = paragraph.add_run("Power supply availability may vary by country.")
        run = paragraph.add_run("Battery is internal and replaceable by customer. Serviceable by warranty.  ")
        run.font.color.rgb = RGBColor(0, 0, 153) 
        paragraph.add_run().add_break(WD_BREAK.LINE)

        # Add table
        insert_table(doc, df)

        # Insert HR
        insert_horizontal_line(doc.add_paragraph(), thickness=3)

        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)