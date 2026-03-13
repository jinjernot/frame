from app.routes.qs_tool.core.blocks.paragraph import insert_list

def service_section(doc, df):
    """Service and support techspecs section"""

    try:
        # Match any value that starts with "Service and Support"
        section_name = None
        for val in df.iloc[:, 1].tolist():
            if str(val).startswith("Service and Support"):
                section_name = val
                break
        
        if section_name:
            insert_list(doc, df, section_name)
        else:
            # Insert error if neither variation is found
            from app.routes.qs_tool.core.blocks.paragraph import insert_error
            insert_error(doc, "'Service and Support' not found in DataFrame.")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)