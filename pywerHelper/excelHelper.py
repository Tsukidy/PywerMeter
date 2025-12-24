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


def write_test_row_to_excel(test_header, samples, workbook_filename="power_measurements.xlsx", sheet_name="Power Data"):
    """
    Write a test's data as a new column in an Excel file.
    
    Args:
        test_header (str): Name of the test (goes in column header)
        samples (list): List of sample values for this test
        workbook_filename (str): Path to the Excel workbook
        sheet_name (str): Name for the sheet
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Convert samples to numeric values (handles text strings from serial data)
        numeric_samples = pd.to_numeric(samples, errors='coerce')
        
        # Check if workbook exists
        if os.path.exists(workbook_filename):
            # Read existing data
            try:
                existing_df = pd.read_excel(workbook_filename, sheet_name=sheet_name)
                
                # Ensure DataFrame has enough rows for new samples
                if len(samples) > len(existing_df):
                    # Add empty rows
                    rows_to_add = len(samples) - len(existing_df)
                    empty_rows = pd.DataFrame([['' for _ in existing_df.columns] for _ in range(rows_to_add)], 
                                             columns=existing_df.columns)
                    existing_df = pd.concat([existing_df, empty_rows], ignore_index=True)
                
                # Add new column with test data (as numeric values)
                existing_df[test_header] = numeric_samples
                updated_df = existing_df
                
            except Exception:
                # Sheet doesn't exist or can't be read, create new DataFrame
                updated_df = pd.DataFrame({test_header: numeric_samples})
        else:
            # Create new workbook with first column
            updated_df = pd.DataFrame({test_header: numeric_samples})
        
        # Write to Excel without index and without default header names
        with pd.ExcelWriter(workbook_filename, engine='openpyxl', mode='w') as writer:
            updated_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        logger.info(f"Wrote test '{test_header}' with {len(samples)} samples to {workbook_filename}")
        print(f"Wrote test '{test_header}' to Excel: {len(samples)} samples")
        return True
        
    except Exception as e:
        logger.error(f"Error writing test column to Excel: {e}")
        print(f"Error writing to Excel: {e}")
        return False

class PowerCalc:
    """
    A class to handle power calculations and Excel formula operations.
    """
    
    def __init__(self, workbook_filename="power_measurements.xlsx", sheet_name="Power Data"):
        """
        Initialize the PowerCalc class.
        
        Args:
            workbook_filename (str): Path to the Excel workbook
            sheet_name (str): Name of the sheet to work with
        """
        self.workbook_filename = workbook_filename
        self.sheet_name = sheet_name
        self.wb = None
        self.ws = None
        self.df = None
        self.last_data_row = None
    
    def _load_workbook(self):
        """Load the Excel workbook and worksheet."""
        if not os.path.exists(self.workbook_filename):
            logger.warning(f"Workbook not found: {self.workbook_filename}")
            print(f"Workbook not found: {self.workbook_filename}")
            return False
        
        # Read the Excel file with pandas
        self.df = pd.read_excel(self.workbook_filename, sheet_name=self.sheet_name)
        
        if self.df.empty:
            logger.warning("No data found in workbook")
            print("No data found in workbook")
            return False
        
        # Load workbook with openpyxl
        from openpyxl import load_workbook
        self.wb = load_workbook(self.workbook_filename)
        self.ws = self.wb[self.sheet_name]
        
        # Find the last row with data
        self.last_data_row = len(self.df) + 1  # +1 for header row
        
        return True
    
    def _save_and_close(self):
        """Save and close the workbook."""
        if self.wb:
            self.wb.save(self.workbook_filename)
            self.wb.close()
    
    def add_averages(self):
        """
        Add average calculations to the bottom of each column in the Excel file.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            if not self._load_workbook():
                return False
            
            # Add average formula for each column
            from openpyxl.utils import get_column_letter
            
            for col_idx, col_name in enumerate(self.df.columns, start=1):
                # Get column letter
                col_letter = get_column_letter(col_idx)
                
                # Add "Average" label in first column of the average row
                if col_idx == 1:
                    self.ws[f'A{self.last_data_row + 2}'] = 'Average'
                
                # Add AVERAGE formula
                formula_cell = self.ws[f'{col_letter}{self.last_data_row + 2}']
                formula_cell.value = f'=AVERAGE({col_letter}2:{col_letter}{self.last_data_row})'
            
            # Save the workbook
            self._save_and_close()
            
            logger.info(f"Added average calculations to {self.workbook_filename}")
            print(f"Added average calculations to Excel file")
            return True
            
        except Exception as e:
            logger.error(f"Error adding calculations: {e}")
            print(f"Error adding calculations: {e}")
            return False
    
    def totalAnnualPower(self):
        """
        Calculate and add Total Annual Power based on the formula:
        totalAnnualPower = 8760/1000*(off*0.15 + shortidle*0.45 + longidle*0.1 + sleep*0.3)
        
        Adds a "Total Annual Power" column to the left of the "sleep" column with the calculated value.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            if not self._load_workbook():
                return False
            
            from openpyxl.utils import get_column_letter
            
            # Find columns by header name (case-insensitive)
            columns_needed = ['off', 'shortidle', 'longidle', 'sleep']
            column_positions = {}
            
            for col_idx, col_name in enumerate(self.df.columns, start=1):
                col_name_lower = str(col_name).lower()
                if col_name_lower in columns_needed:
                    column_positions[col_name_lower] = col_idx
            
            # Check if all required columns exist
            missing_columns = [col for col in columns_needed if col not in column_positions]
            if missing_columns:
                logger.warning(f"Missing required columns: {missing_columns}")
                print(f"Warning: Missing required columns: {missing_columns}")
                return False
            
            # Find the sleep column position
            sleep_col_idx = column_positions['sleep']
            
            # Insert a new column to the left of sleep
            self.ws.insert_cols(sleep_col_idx)
            
            # Add header for the new column
            new_col_letter = get_column_letter(sleep_col_idx)
            self.ws[f'{new_col_letter}1'] = 'Total Annual Power'
            
            # Get column letters for the formula (adjust for insertion)
            off_letter = get_column_letter(column_positions['off'])
            shortidle_letter = get_column_letter(column_positions['shortidle'] if column_positions['shortidle'] < sleep_col_idx else column_positions['shortidle'] + 1)
            longidle_letter = get_column_letter(column_positions['longidle'] if column_positions['longidle'] < sleep_col_idx else column_positions['longidle'] + 1)
            sleep_letter = get_column_letter(sleep_col_idx + 1)  # Sleep moved one column to the right
            
            # Add the formula in the Average row
            # The Average row is at last_data_row + 2
            avg_row = self.last_data_row + 2
            formula = f'=8760/1000*({off_letter}{avg_row}*0.15+{shortidle_letter}{avg_row}*0.45+{longidle_letter}{avg_row}*0.1+{sleep_letter}{avg_row}*0.3)'
            self.ws[f'{new_col_letter}{avg_row}'] = formula
            
            # Save the workbook
            self._save_and_close()
            
            logger.info(f"Added Total Annual Power calculation to {self.workbook_filename}")
            print(f"Added Total Annual Power calculation to Excel file")
            return True
            
        except Exception as e:
            logger.error(f"Error adding Total Annual Power calculation: {e}")
            print(f"Error adding Total Annual Power calculation: {e}")
            return False


# Backward compatibility wrapper function
def powerCalc(workbook_filename="power_measurements.xlsx", sheet_name="Power Data"):
    """
    Legacy function wrapper for PowerCalc class.
    Add average calculations to the bottom of each column in the Excel file.
    
    Args:
        workbook_filename (str): Path to the Excel workbook
        sheet_name (str): Name of the sheet to add calculations to
        
    Returns:
        bool: True if successful, False otherwise
    """
    calc = PowerCalc(workbook_filename, sheet_name)
    return calc.add_averages()


# Module-level exports
__all__ = ['create_workbook', 'import_data_to_workbook', 'import_multiple_files', 'write_test_row_to_excel', 'powerCalc', 'PowerCalc']