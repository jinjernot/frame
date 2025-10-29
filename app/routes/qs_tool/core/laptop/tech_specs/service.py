from app.routes.qs_tool.core.blocks.paragraph import insert_list

def service_section(doc, df):
    """Service and support techspecs section"""

    try:
        key_to_use = "Service and Support[45]" if "Service and Support[45]" in df.columns else "Service and Support"
        insert_list(doc, df, key_to_use)
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)