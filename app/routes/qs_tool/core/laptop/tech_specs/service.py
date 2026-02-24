from app.routes.qs_tool.core.blocks.paragraph import insert_list

def service_section(doc, df):
    """Service and support techspecs section"""

    try:
        # Check for both variations of the section name
        section_name = None
        if "Service and Support [45]" in df.iloc[:, 1].tolist():
            section_name = "Service and Support [45]"
        elif "Service and Support" in df.iloc[:, 1].tolist():
            section_name = "Service and Support"
        
        if section_name:
            insert_list(doc, df, section_name)
        else:
            # Insert error if neither variation is found
            from app.routes.qs_tool.core.blocks.paragraph import insert_error
            insert_error(doc, "'Service and Support' (with or without [45]) not found in DataFrame.")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)