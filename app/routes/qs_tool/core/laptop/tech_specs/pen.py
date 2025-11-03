from app.routes.qs_tool.core.blocks.paragraph import insert_list

def digital_pen_section(doc, df):
    """Digital Pen techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, df, "Digital Pen")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)