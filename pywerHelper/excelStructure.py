"""Excel file structure detection and management for pywerMeter."""
import logging
from typing import List, Tuple, Optional
from functools import lru_cache

logger = logging.getLogger(__name__)


@lru_cache(maxsize=32)
def get_column_letter_cached(col_idx: int) -> str:
    """
    Cached version of get_column_letter for performance.
    
    Args:
        col_idx: Column index (1-based)
        
    Returns:
        Column letter (e.g., 'A', 'AB')
    """
    from openpyxl.utils import get_column_letter
    return get_column_letter(col_idx)


class ExcelFileStructure:
    """Helper class to detect and manage Excel file structure."""
    
    def __init__(self, ws):
        """
        Initialize structure detector.
        
        Args:
            ws: openpyxl worksheet object
        """
        self.ws = ws
        self.has_averages = False
        self.has_start_times = False
        self.has_both = False
        self.header_row = 1
        self.data_start_row = 2
        self.start_time_row = None
        self.average_row = None
        
        self._detect_structure()
    
    def _detect_structure(self):
        """Detect the Excel file structure based on row 1 and row 3."""
        cell_a1 = self.ws['A1'].value
        cell_a3 = self.ws['A3'].value if self.ws.max_row >= 3 else None
        
        if cell_a1 == 'Averages':
            self.has_averages = True
            self.average_row = 2
            
            if cell_a3 == 'Test Start Times':
                # Full structure: Averages + Start Times
                self.has_both = True
                self.has_start_times = True
                self.header_row = 5
                self.data_start_row = 6
                self.start_time_row = 4
                logger.debug("Detected structure: Averages + Start Times (rows 1-5)")
            else:
                # Just averages
                self.header_row = 3
                self.data_start_row = 4
                self.start_time_row = None
                logger.debug("Detected structure: Averages only (rows 1-3)")
                
        elif cell_a1 == 'Test Start Times':
            # Just start times
            self.has_start_times = True
            self.header_row = 3
            self.data_start_row = 4
            self.start_time_row = 2
            logger.debug("Detected structure: Start Times only (rows 1-3)")
        else:
            # Regular file
            self.header_row = 1
            self.data_start_row = 2
            logger.debug("Detected structure: Regular file (no special headers)")
    
    def get_headers(self) -> List[Tuple[int, str]]:
        """
        Get non-empty headers from header row.
        
        Returns:
            List of (column_index, header_value) tuples
        """
        headers = []
        for idx, cell in enumerate(self.ws[self.header_row], start=1):
            if cell.value and str(cell.value).strip():
                headers.append((idx, str(cell.value).strip()))
            elif idx > 1 and idx > len(headers) + 5:
                # Stop if we've gone 5 columns past the last header
                break
        return headers
    
    def find_column_index(self, test_header: str) -> Optional[int]:
        """
        Find column index for a given test header.
        
        Args:
            test_header: Header name to search for
            
        Returns:
            Column index (1-based) or None if not found
        """
        headers = self.get_headers()
        for idx, header in headers:
            if header == test_header:
                logger.debug(f"Found '{test_header}' at column {idx}")
                return idx
        logger.debug(f"Header '{test_header}' not found")
        return None
    
    def get_last_data_row(self, col_idx: int) -> int:
        """
        Find the last non-empty row in a specific column.
        
        Args:
            col_idx: Column index (1-based)
            
        Returns:
            Last row number with data, or data_start_row - 1 if empty
        """
        col_letter = get_column_letter_cached(col_idx)
        last_row = self.data_start_row - 1
        
        for row_idx in range(self.data_start_row, self.ws.max_row + 1):
            if self.ws[f'{col_letter}{row_idx}'].value is not None:
                last_row = row_idx
        
        return last_row
    
    def get_next_available_column(self) -> int:
        """
        Get the next available column index after existing data.
        
        Returns:
            Next available column index (1-based)
        """
        headers = self.get_headers()
        if headers:
            return max(idx for idx, _ in headers) + 1
        return 1
    
    def clear_column_data(self, col_idx: int):
        """
        Clear all data in a column from data_start_row onwards.
        
        Args:
            col_idx: Column index (1-based)
        """
        col_letter = get_column_letter_cached(col_idx)
        max_row = self.ws.max_row
        
        for row_idx in range(self.data_start_row, max_row + 1):
            self.ws[f'{col_letter}{row_idx}'].value = None
        
        logger.debug(f"Cleared column {col_letter} from row {self.data_start_row}")
    
    def write_data_to_column(self, col_idx: int, data: list, start_row: Optional[int] = None):
        """
        Write data to a column starting at specified row.
        
        Args:
            col_idx: Column index (1-based)
            data: List of values to write
            start_row: Starting row (defaults to data_start_row)
        """
        if start_row is None:
            start_row = self.data_start_row
        
        col_letter = get_column_letter_cached(col_idx)
        
        for i, value in enumerate(data, start=start_row):
            self.ws[f'{col_letter}{i}'].value = value
        
        logger.debug(f"Wrote {len(data)} values to column {col_letter} starting at row {start_row}")
    
    def update_average_formula(self, col_idx: int, last_data_row: int):
        """
        Update average formula for a column.
        
        Args:
            col_idx: Column index (1-based)
            last_data_row: Last row containing data
        """
        if not self.has_averages or self.average_row is None:
            return
        
        col_letter = get_column_letter_cached(col_idx)
        formula = f'=AVERAGE({col_letter}{self.data_start_row}:{col_letter}{last_data_row})'
        self.ws[f'{col_letter}{self.average_row}'].value = formula
        
        logger.debug(f"Updated average formula for column {col_letter}: {formula}")
    
    def set_start_time(self, col_idx: int, start_time_str: str):
        """
        Set start time for a column.
        
        Args:
            col_idx: Column index (1-based)
            start_time_str: Start time string
        """
        if self.start_time_row is None:
            return
        
        col_letter = get_column_letter_cached(col_idx)
        self.ws[f'{col_letter}{self.start_time_row}'].value = start_time_str
        
        logger.debug(f"Set start time for column {col_letter}: {start_time_str}")


__all__ = ['ExcelFileStructure', 'get_column_letter_cached']
