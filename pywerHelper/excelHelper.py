# This script handles the excel file operations for the pywerHelper package.
# It includes functions to read from and write to Excel files using the pandas library.
# Author: Dylan Pope!
# Date: 2025-11-07
# Version: 1.0.0

import os
import logging

# Set up module logger
logger = logging.getLogger(__name__)

# Try to import pandas and openpyxl with helpful error messages
try:
    import pandas as pd
except ImportError:
    logger.error("pandas is not installed. Install with: pip install pandas")
    raise

try:
    import openpyxl
except ImportError:
    logger.error("openpyxl is not installed. Install with: pip install openpyxl")
    raise


def testFunction():
    """Test function to verify module import."""
    return "This is a test function from excelHelper.py"


def create_workbook(output_filename="power_data.xlsx"):
    """
    Create a new Excel workbook with an empty DataFrame.
    
    Args:
        output_filename (str): The filename for the new Excel workbook
        
    Returns:
        str: Path to the created workbook, or None if failed
    """
    try:
        # Create an empty DataFrame
        df = pd.DataFrame()
        
        # Write to Excel
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Summary', index=False)
        
        logger.info(f"Created new workbook: {output_filename}")
        print(f"Created new workbook: {output_filename}")
        return output_filename
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        print(f"Error creating workbook: {e}")
        return None


def import_data_to_workbook(data_file, workbook_filename="power_data.xlsx", sheet_name=None):
    """
    Import data from a text file into an Excel workbook as a new sheet.
    Each line from the data file becomes a row in the worksheet.
    
    Args:
        data_file (str): Path to the data file to import
        workbook_filename (str): Path to the Excel workbook
        sheet_name (str): Name for the sheet. If None, uses the data file name without extension
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Check if data file exists
        if not os.path.exists(data_file):
            logger.warning(f"Data file not found: {data_file}")
            print(f"Data file not found: {data_file}")
            return False
        
        # Read data from text file
        with open(data_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Strip newlines and create DataFrame
        data = [line.strip() for line in lines if line.strip()]
        
        if not data:
            logger.warning(f"No data found in file: {data_file}")
            print(f"No data found in file: {data_file}")
            return False
        
        df = pd.DataFrame({'Power Data': data})
        
        # Determine sheet name
        if sheet_name is None:
            sheet_name = os.path.splitext(os.path.basename(data_file))[0]
        
        # Sanitize sheet name (Excel has 31 char limit and special char restrictions)
        sheet_name = sheet_name[:31].replace('[', '').replace(']', '').replace('*', '').replace('/', '').replace('\\', '').replace('?', '').replace(':', '')
        
        # Check if workbook exists
        if os.path.exists(workbook_filename):
            # Append to existing workbook
            with pd.ExcelWriter(workbook_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # Create new workbook
            with pd.ExcelWriter(workbook_filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        logger.info(f"Imported {len(data)} rows from {data_file} to sheet '{sheet_name}'")
        print(f"Imported {len(data)} rows from {data_file} to sheet '{sheet_name}'")
        return True
    except Exception as e:
        logger.error(f"Error importing data from {data_file}: {e}")
        print(f"Error importing data: {e}")
        return False


def import_multiple_files(data_files, workbook_filename="power_data.xlsx"):
    """
    Import multiple data files into a single Excel workbook, each as a separate sheet.
    
    Args:
        data_files (list): List of data file paths to import
        workbook_filename (str): Path to the Excel workbook
        
    Returns:
        int: Number of files successfully imported
    """
    if not data_files:
        logger.warning("No data files provided to import")
        print("No data files provided to import")
        return 0
    
    success_count = 0
    
    for data_file in data_files:
        if import_data_to_workbook(data_file, workbook_filename):
            success_count += 1
    
    logger.info(f"Imported {success_count} out of {len(data_files)} files to {workbook_filename}")
    print(f"\nImported {success_count} out of {len(data_files)} files to {workbook_filename}")
    return success_count


# Module-level exports
__all__ = ['create_workbook', 'import_data_to_workbook', 'import_multiple_files', 'testFunction']