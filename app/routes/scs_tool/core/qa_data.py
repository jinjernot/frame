import pandas as pd
import json
import os
from config import *
from io import BytesIO

from app.routes.scs_tool.core.process_data import process_data, process_data_granular
from app.routes.scs_tool.core.qa_av import av_check
from app.routes.scs_tool.core.format_data import format_data, format_data_granular
from app.routes.scs_tool.core.product_line import pl_check
from app.routes.scs_tool.core.check_missing_fields import check_missing_fields
from app.routes.scs_tool.core.npu_check import npu_check


def clean_report(file):
    """
    Processes a standard report using the new restructured JSON data.
    """
    try:
        # --- 1. Initial Setup & Cleaning ---
        file_content = file.read()
        file_buffer = BytesIO(file_content)

        df = pd.read_excel(file_buffer, engine='openpyxl')
        df = df.drop(SCS_COLS_TO_ADD, axis=1, errors='ignore')
        df[SCS_COLS_TO_ADD] = ''

        pl_check(df)

        # Basic data cleaning
        df = df[df['ContainerValue'] != '[BLANK]'].dropna(
            subset=['ContainerValue', 'ContainerName'])
        df.replace('\u00A0', ' ', regex=True, inplace=True)
        df.loc[df['ContainerValue'].str.endswith(
            ';', na=False), 'ContainerValue'] = df['ContainerValue'].str.slice(stop=-1)
        df['PhwebDescription'] = df['PhwebDescription'].str.lstrip()
        df['ContainerValue'] = df['ContainerValue'].astype(str)

        # --- 2. Component Group Filtering ---
        with open(SCS_COMPONENT_GROUPS_PATH, 'r', encoding='utf-8') as json_file:
            json_data = json.load(json_file)
        groups = json_data['Groups']

        filtered_rows = df[df.apply(lambda row: any(row['ComponentGroup'] == group['ComponentGroup']
                                    and row['ContainerName'] in group['ContainerName'] for group in groups), axis=1)]
        df = filtered_rows.copy()

        # --- 3. Main Data Processing Loop ---
        json_files = [f for f in os.listdir(
            SCS_JSON_PATH) if f.endswith('.json')]

        for json_file in json_files:
            container_name = os.path.splitext(json_file)[0]
            json_file_path = os.path.join(SCS_JSON_PATH, json_file)
            df = process_data(json_file_path, container_name, df)

        # --- 4. NPU Validation Step ---
        df = npu_check(df, NPU_JSON_PATH)

        # --- 5. Save Output ---
        file_buffer.seek(0)
        with pd.ExcelFile(file_buffer, engine='openpyxl') as excel_file:
            if "ms4" in excel_file.sheet_names or "MS4" in excel_file.sheet_names:
                file_buffer.seek(0)
                df_final = av_check(file_buffer)
                with pd.ExcelWriter(SCS_REGULAR_FILE_PATH) as writer:
                    df.to_excel(writer, sheet_name='qa', index=False)
                    df_final.to_excel(
                        writer, sheet_name='duplicated', index=False)
            else:
                df.to_excel(SCS_REGULAR_FILE_PATH, index=False)

        format_data()
        return df

    except Exception as e:
        print(f"An error occurred in clean_report: {e}")
        return None


async def clean_report_granular(file):
    """
    Processes a granular report asynchronously using the new restructured JSON data.
    """
    try:
        # --- 1. Initial Setup & Cleaning ---
        file_content = file.read()
        file_buffer = BytesIO(file_content)

        df_g = pd.read_excel(file_buffer, engine='openpyxl')
        df_g = df_g.drop(SCS_COLS_TO_DROP_GRANULAR, axis=1, errors='ignore')
        df_g[SCS_COLS_TO_ADD] = ''

        df_g = df_g.dropna(
            subset=['Granular Container Value', 'Granular Container Tag'])
        df_g.replace('\u00A0', ' ', regex=True, inplace=True)
        df_g.loc[df_g['Granular Container Value'].str.endswith(
            ';', na=False), 'Granular Container Value'] = df_g['Granular Container Value'].str.slice(stop=-1)
        df_g['Granular Container Value'] = df_g['Granular Container Value'].astype(
            str)

        # --- 2. Data Processing Loop ---
        json_files = [f for f in os.listdir(
            SCS_JSON_GRANULAR_PATH) if f.endswith('.json')]

        with open(SCS_GRANULAR_COMPONENT_GROUPS_PATH, 'r', encoding='utf-8') as json_file:
            json_data = json.load(json_file)
        groups = json_data['Groups']

        filtered_rows = df_g[df_g.apply(lambda row: any(row['SCSGroup'] == group['ComponentGroup']
                                        and row['Granular Container Tag'] in group['ContainerName'] for group in groups), axis=1)]
        df_g = filtered_rows.copy()

        for json_file in json_files:
            container_name = os.path.splitext(json_file)[0]
            json_file_path = os.path.join(SCS_JSON_GRANULAR_PATH, json_file)

            df_g = process_data_granular(json_file_path, container_name, df_g)

        # --- 3. Final Checks and Save ---
        df_g = check_missing_fields(df_g, SCS_GRANULAR_COMPONENT_GROUPS_PATH)

        file_buffer.seek(0)
        with pd.ExcelFile(file_buffer, engine='openpyxl') as excel_file:
            print(f"Available sheets: {excel_file.sheet_names}")
            if "ms4" in excel_file.sheet_names:
                file_buffer.seek(0)
                df_final = av_check(file_buffer)
                with pd.ExcelWriter(SCS_GRANULAR_FILE_PATH) as writer:
                    df_g.to_excel(writer, sheet_name='qa', index=False)
                    df_final.to_excel(
                        writer, sheet_name='duplicated', index=False)
            else:
                df_g.to_excel(SCS_GRANULAR_FILE_PATH, index=False)

        format_data_granular()
        return df_g

    except Exception as e:
        print(f"An error occurred in clean_report_granular: {e}")
        return None
