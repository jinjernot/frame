from app.routes.qs_tool.core.blocks.paragraph import insert_list

def docking_section(doc, df):
    """Docking techspecs section"""
    
    try:
        # Function to insert the list of values
        insert_list(doc, df, "Docking (sold separately)")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)