from app.routes.qs_tool.core.blocks.paragraph import *
from app.routes.qs_tool.core.blocks.title import *
from app.routes.qs_tool.core.blocks.table import *
from app.routes.qs_tool.core.format.hr import *

from docx.enum.text import WD_BREAK
import pandas as pd

NO_BOLD_POWER_LABELS = {
    "Weight(DC Cable Included)",
    "Input",
    "Input Efficiency",
    "Input frequency range",
    "Input AC current",
    "Output power",
    "DC output",
    "Hold-up time",
    "Output Over Current",
    "Protection",
    "AC Inlet Type",
    "DC Cable Connector",
    "DC Cable Material",
    "Connector",
    "Operating temperature",
    "Non - operating(storage)",
    "temperature",
    "Altitude",
    "Humidity",
    "Storage Humidity",
    "EMI and Safety",
    "Certifications",
}


def unbold_power_labels(doc, table_count_before):
    """Remove bold styling from labels in all tables added by the power section."""
    for table in doc.tables[table_count_before:]:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip() in NO_BOLD_POWER_LABELS:
                            run.font.bold = False


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
        table_count_before = len(doc.tables)
        insert_table(doc, df)
        unbold_power_labels(doc, table_count_before)

        # Insert HR
        insert_horizontal_line(doc.add_paragraph(), thickness=3)

        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)