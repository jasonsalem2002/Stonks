# report_generator.py
import pandas as pd
import logging
import os

def convert_col_num_to_letter(col_num):
    """
    Converts a zero-indexed column number to an Excel column letter.
    E.g., 0 -> 'A', 1 -> 'B', ..., 25 -> 'Z', 26 -> 'AA', etc.
    """
    string = ""
    while col_num >= 0:
        string = chr(col_num % 26 + 65) + string
        col_num = col_num // 26 - 1
    return string

def add_df_to_sheet(writer, df, sheet_name, conditional_format_cols=None):
    """
    Adds a DataFrame to an Excel sheet with optional conditional formatting.
    """
    try:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Apply header styling
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4F81BD',
            'font_color': '#FFFFFF',
            'border': 1
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            # Set column widths
            worksheet.set_column(col_num, col_num, 15)
        
        # Apply conditional formatting if specified
        if conditional_format_cols:
            start_col, end_col = conditional_format_cols
            for col in range(start_col, end_col + 1):
                col_letter = convert_col_num_to_letter(col - 1)  # Adjust if col is 1-based
                worksheet.conditional_format(f"{col_letter}2:{col_letter}{len(df)+1}", {
                    'type': 'cell',
                    'criteria': '>',
                    'value': 0,
                    'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
                })
                worksheet.conditional_format(f"{col_letter}2:{col_letter}{len(df)+1}", {
                    'type': 'cell',
                    'criteria': '<',
                    'value': 0,
                    'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                })
    except Exception as e:
        logging.error(f"Error adding DataFrame to sheet '{sheet_name}': {e}")

def add_individual_symbol_sheets(writer, symbol, changes_df):
    """
    Creates individual sheets for each symbol with daily, weekly, and monthly changes.
    """
    try:
        sheet_name = f"{symbol} Changes"
        sheet_name = sheet_name[:31]  # Excel sheet name max length is 31
        changes_df.to_excel(writer, sheet_name=sheet_name, index=False)
        logging.info(f"Added sheet '{sheet_name}' with {len(changes_df)} rows.")
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Apply header formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4F81BD',
            'font_color': '#FFFFFF',
            'border': 1
        })
        
        for col_num, value in enumerate(changes_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            # Set column widths
            worksheet.set_column(col_num, col_num, 18)
        
        # Apply conditional formatting
        for col in ['B', 'C', 'D']:
            worksheet.conditional_format(f"{col}2:{col}{len(changes_df)+1}", {
                'type': 'cell',
                'criteria': '>',
                'value': 0,
                'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            })
            worksheet.conditional_format(f"{col}2:{col}{len(changes_df)+1}", {
                'type': 'cell',
                'criteria': '<',
                'value': 0,
                'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            })
    except Exception as e:
        logging.error(f"Error creating individual sheet for {symbol}: {e}")
        
def insert_charts_into_workbook(writer, charts, df_overview):
    """
    Inserts chart images into the 'Stock Overview' sheet.
    """
    try:
        workbook = writer.book
        worksheet = writer.sheets['Stock Overview']
        row_position = len(df_overview) + 5
        for chart in charts:
            if os.path.exists(chart):
                worksheet.insert_image(f"A{row_position}", chart, {'x_scale': 0.5, 'y_scale': 0.5})
                row_position += 30  # Adjust the row position for next image
    except Exception as e:
        logging.error(f"Error inserting charts into workbook: {e}")

def cleanup_charts(charts):
    """
    Deletes chart image files after embedding them into the Excel workbook.
    """
    for chart in charts:
        if os.path.exists(chart):
            try:
                os.remove(chart)
                logging.info(f"Deleted chart image: {chart}")
            except Exception as e:
                logging.error(f"Error deleting chart image {chart}: {e}")
