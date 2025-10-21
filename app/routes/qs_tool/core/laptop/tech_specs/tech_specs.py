from app.routes.qs_tool.core.laptop.tech_specs.product_name import product_name_section
from app.routes.qs_tool.core.laptop.tech_specs.operating_systems import operating_systems_section
from app.routes.qs_tool.core.laptop.tech_specs.processors import processors_section
from app.routes.qs_tool.core.laptop.tech_specs.graphics import graphics_section
from app.routes.qs_tool.core.laptop.tech_specs.display import display_section
from app.routes.qs_tool.core.laptop.tech_specs.docking import docking_section
from app.routes.qs_tool.core.laptop.tech_specs.storage import storage_section
from app.routes.qs_tool.core.laptop.tech_specs.memory import memory_section
from app.routes.qs_tool.core.laptop.tech_specs.networking import networking_section
from app.routes.qs_tool.core.laptop.tech_specs.audio import audio_section
from app.routes.qs_tool.core.laptop.tech_specs.keyboard import keyboard_section
from app.routes.qs_tool.core.laptop.tech_specs.software import software_section
from app.routes.qs_tool.core.laptop.tech_specs.power import power_section
from app.routes.qs_tool.core.laptop.tech_specs.dimensions import dimensions_section
from app.routes.qs_tool.core.laptop.tech_specs.ports import ports_section
from app.routes.qs_tool.core.laptop.tech_specs.service import service_section
from app.routes.qs_tool.core.laptop.tech_specs.certification_compliance import certification_section 
from config import QS_TECHSPECS_PATH


from docx.shared import RGBColor

import pandas as pd

def tech_specs_section(doc, file):
    """TechSpecs Section"""

    try:
        # Load available sheet names
        xls = pd.ExcelFile(file.stream, engine='openpyxl')
        available_sheets = xls.sheet_names  # List of sheets in the file

        # Possible sheet names
        sheet_names = ["Tech Specs & QS Features", "QS Features"]
        df = None

        # Find the first matching sheet
        for sheet in sheet_names:
            if sheet in available_sheets:
                df = pd.read_excel(xls, sheet_name=sheet)
                break  # Exit loop once a valid sheet is found

        if df is None:
            raise ValueError(f"None of the specified sheets {sheet_names} were found. Available sheets: {available_sheets}")

        # Remove extra spaces from the end of each value and convert all columns to strings
        df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

        values_to_drop = ["General content", "General content 01", "General content 02"]
        
        if not df.empty:
            # Check if there are at least 2 columns before trying to access index 1
            if len(df.columns) > 1:
                df = df[~df[df.columns[1]].isin(values_to_drop)]

        # Filter out rows where the "Value" column (assumed to be column 2, index 1) is empty
        df_filtered = df.dropna(subset=[df.columns[1]])

        # Save the filtered DataFrame to a new Excel file
        output_file = QS_TECHSPECS_PATH
        df_filtered.to_excel(output_file, index=False)

        # Read the filtered DataFrame
        df = pd.read_excel(output_file, sheet_name='Sheet1', engine='openpyxl')
        df = df.astype(str)

        # Run the functions to build the tech specs section
        product_name_section(doc, file)
        operating_systems_section(doc, df)
        processors_section(doc, file)
        graphics_section(doc, df)
        display_section(doc, df)
        docking_section(doc, df)
        storage_section(doc, df)
        memory_section(doc, df)
        networking_section(doc, df)
        audio_section(doc, df)
        keyboard_section(doc, df)
        software_section(doc, df)
        power_section(doc, df)
        dimensions_section(doc, df)
        ports_section(doc, df)
        service_section(doc, df)
        certification_section(doc, df)
        
        
    except Exception as e:
        error_message = f"An error occurred: {e}"
        print(error_message)

        # Add error message to the document in red
        para = doc.add_paragraph()
        run = para.add_run(error_message)
        run.font.color.rgb = RGBColor(255, 0, 0)  # Red color

        return error_message