"""Tests for excelStructure module."""
import pytest
import tempfile
import os
from openpyxl import Workbook, load_workbook
from pywerHelper.excelStructure import ExcelFileStructure, get_column_letter_cached


class TestGetColumnLetterCached:
    """Test get_column_letter_cached function."""
    
    def test_single_letter_columns(self):
        """Test conversion for single letter columns."""
        assert get_column_letter_cached(1) == "A"
        assert get_column_letter_cached(5) == "E"
        assert get_column_letter_cached(26) == "Z"
    
    def test_double_letter_columns(self):
        """Test conversion for double letter columns."""
        assert get_column_letter_cached(27) == "AA"
        assert get_column_letter_cached(52) == "AZ"
        assert get_column_letter_cached(53) == "BA"
    
    def test_caching(self):
        """Test that caching works (same object returned)."""
        result1 = get_column_letter_cached(10)
        result2 = get_column_letter_cached(10)
        assert result1 == result2
        assert result1 == "J"


@pytest.fixture
def temp_excel_file():
    """Create a temporary Excel file."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        temp_path = f.name
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Test Sheet"
    wb.save(temp_path)
    wb.close()
    
    yield temp_path
    
    # Cleanup
    if os.path.exists(temp_path):
        os.unlink(temp_path)


@pytest.fixture
def excel_with_averages(temp_excel_file):
    """Create Excel file with averages structure."""
    wb = load_workbook(temp_excel_file)
    ws = wb.active
    
    ws['A1'] = "Averages"
    ws['A3'] = "Start Times"
    ws['A5'] = "Time (min)"
    
    wb.save(temp_excel_file)
    wb.close()
    
    return temp_excel_file


@pytest.fixture
def excel_with_start_times_only(temp_excel_file):
    """Create Excel file with start times only."""
    wb = load_workbook(temp_excel_file)
    ws = wb.active
    
    ws['A3'] = "Start Times"
    ws['A4'] = "Time (min)"
    
    wb.save(temp_excel_file)
    wb.close()
    
    return temp_excel_file


@pytest.fixture
def excel_regular(temp_excel_file):
    """Create Excel file with regular structure (no special rows)."""
    wb = load_workbook(temp_excel_file)
    ws = wb.active
    
    ws['A1'] = "Time (min)"
    
    wb.save(temp_excel_file)
    wb.close()
    
    return temp_excel_file


class TestExcelFileStructure:
    """Test ExcelFileStructure class."""
    
    def test_detect_averages_and_start_times(self, excel_with_averages):
        """Test detection of file with both averages and start times."""
        structure = ExcelFileStructure(excel_with_averages)
        
        assert structure.has_averages_row is True
        assert structure.has_start_times_row is True
        assert structure.averages_row_num == 1
        assert structure.start_times_row_num == 3
        assert structure.headers_row_num == 5
        assert structure.data_start_row == 6
    
    def test_detect_start_times_only(self, excel_with_start_times_only):
        """Test detection of file with start times only."""
        structure = ExcelFileStructure(excel_with_start_times_only)
        
        assert structure.has_averages_row is False
        assert structure.has_start_times_row is True
        assert structure.averages_row_num is None
        assert structure.start_times_row_num == 3
        assert structure.headers_row_num == 4
        assert structure.data_start_row == 5
    
    def test_detect_regular_structure(self, excel_regular):
        """Test detection of regular file (no special rows)."""
        structure = ExcelFileStructure(excel_regular)
        
        assert structure.has_averages_row is False
        assert structure.has_start_times_row is False
        assert structure.averages_row_num is None
        assert structure.start_times_row_num is None
        assert structure.headers_row_num == 1
        assert structure.data_start_row == 2
    
    def test_get_headers(self, excel_with_averages):
        """Test getting headers from Excel file."""
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        # Add some headers
        ws['B5'] = "Column1"
        ws['C5'] = "Column2"
        ws['D5'] = "Column3"
        
        wb.save(excel_with_averages)
        wb.close()
        
        structure = ExcelFileStructure(excel_with_averages)
        headers = structure.get_headers()
        
        assert "Column1" in headers
        assert "Column2" in headers
        assert "Column3" in headers
    
    def test_find_column_index(self, excel_with_averages):
        """Test finding column index by header name."""
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        ws['B5'] = "Test Column"
        ws['C5'] = "Another Column"
        
        wb.save(excel_with_averages)
        wb.close()
        
        structure = ExcelFileStructure(excel_with_averages)
        
        assert structure.find_column_index("Test Column") == 2
        assert structure.find_column_index("Another Column") == 3
        assert structure.find_column_index("Nonexistent") is None
    
    def test_get_last_data_row(self, excel_with_averages):
        """Test getting last data row with content."""
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        # Add headers
        ws['B5'] = "Data"
        
        # Add some data
        ws['B6'] = 1.0
        ws['B7'] = 2.0
        ws['B8'] = 3.0
        
        wb.save(excel_with_averages)
        wb.close()
        
        structure = ExcelFileStructure(excel_with_averages)
        last_row = structure.get_last_data_row(2)  # Column B (index 2)
        
        assert last_row == 8
    
    def test_get_last_data_row_empty_column(self, excel_with_averages):
        """Test getting last data row for empty column."""
        structure = ExcelFileStructure(excel_with_averages)
        last_row = structure.get_last_data_row(5)  # Empty column
        
        # Should return the row before data_start_row (headers row)
        assert last_row == structure.headers_row_num
    
    def test_clear_column_data(self, excel_with_averages):
        """Test clearing data from a column."""
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        # Add data to column B
        ws['B6'] = 1.0
        ws['B7'] = 2.0
        ws['B8'] = 3.0
        
        wb.save(excel_with_averages)
        wb.close()
        
        structure = ExcelFileStructure(excel_with_averages)
        structure.clear_column_data(2)  # Column B
        
        # Verify data is cleared
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        assert ws['B6'].value is None
        assert ws['B7'].value is None
        assert ws['B8'].value is None
        
        wb.close()
    
    def test_write_data_to_column(self, excel_with_averages):
        """Test writing data to a column."""
        structure = ExcelFileStructure(excel_with_averages)
        
        test_data = [10.5, 20.3, 30.7]
        structure.write_data_to_column(2, test_data)  # Column B
        
        # Verify data is written
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        assert ws['B6'].value == 10.5
        assert ws['B7'].value == 20.3
        assert ws['B8'].value == 30.7
        
        wb.close()
    
    def test_update_average_formula(self, excel_with_averages):
        """Test updating average formula."""
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        # Add some data
        ws['B6'] = 10.0
        ws['B7'] = 20.0
        ws['B8'] = 30.0
        
        wb.save(excel_with_averages)
        wb.close()
        
        structure = ExcelFileStructure(excel_with_averages)
        structure.update_average_formula(2, 8)  # Column B, last row 8
        
        # Verify formula is created
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        formula = ws['B2'].value
        assert formula is not None
        assert "AVERAGE" in str(formula).upper()
        assert "B6:B8" in str(formula)
        
        wb.close()
    
    def test_set_start_time(self, excel_with_averages):
        """Test setting start time."""
        structure = ExcelFileStructure(excel_with_averages)
        structure.set_start_time(2, "5:30 / 5.50 min")  # Column B
        
        # Verify start time is set
        wb = load_workbook(excel_with_averages)
        ws = wb.active
        
        assert ws['B4'].value == "5:30 / 5.50 min"
        
        wb.close()
    
    def test_no_start_times_row(self, excel_regular):
        """Test set_start_time when file has no start times row."""
        structure = ExcelFileStructure(excel_regular)
        
        # Should not raise error, just skip
        structure.set_start_time(2, "Test")
        
        # No start times row should be created
        wb = load_workbook(excel_regular)
        ws = wb.active
        assert ws['A3'].value != "Start Times"
        wb.close()
