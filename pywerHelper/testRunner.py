"""Test execution management for pywerMeter."""
import time
import logging
from typing import Dict, Any, Optional, Callable
from pywerHelper.timeUtils import get_formatted_start_time

logger = logging.getLogger(__name__)


class TestRunner:
    """Manages test execution with timing and data collection."""
    
    def __init__(self, logger_instance: logging.Logger, global_start_time: float):
        """
        Initialize test runner.
        
        Args:
            logger_instance: Logger for test execution
            global_start_time: Global timer start time (time.time())
        """
        self.logger = logger_instance
        self.global_start_time = global_start_time
        self._pause_handler: Optional[Callable] = None
    
    def set_pause_handler(self, handler: Callable):
        """
        Set callback for handling pauses.
        
        Args:
            handler: Function to call during pause
        """
        self._pause_handler = handler
    
    def get_elapsed_time(self) -> float:
        """
        Get elapsed time from global timer start.
        
        Returns:
            Elapsed time in minutes
        """
        return (time.time() - self.global_start_time) / 60
    
    def wait_until(self, target_time_minutes: float, test_name: str = "test") -> float:
        """
        Wait until global timer reaches target time.
        
        Args:
            target_time_minutes: Target time in minutes
            test_name: Name of test for display
            
        Returns:
            Elapsed time when target is reached
        """
        while True:
            elapsed = self.get_elapsed_time()
            
            if elapsed >= target_time_minutes:
                return elapsed
            
            remaining = target_time_minutes - elapsed
            print(f"\rGlobal Timer: {elapsed:.2f} min | Waiting for {test_name} "
                  f"(starts at {target_time_minutes:.2f} min, {remaining:.2f} min remaining)...", 
                  end="", flush=True)
            time.sleep(1)
    
    def run_test(
        self, 
        test_config: Dict[str, Any],
        data_collector_func: Callable,
        **collector_kwargs
    ) -> tuple:
        """
        Run a single test with the given configuration.
        
        Args:
            test_config: Dictionary with test configuration:
                - header: Test name
                - start_time: Start time in minutes
                - duration: Test duration in minutes
                - pause_after: Whether to pause after test
            data_collector_func: Function to call for data collection
            **collector_kwargs: Additional kwargs for data collector
            
        Returns:
            Tuple of (samples, start_time_str, elapsed_time)
        """
        test_header = test_config['header']
        start_time = test_config['start_time']
        duration = test_config['duration']
        pause_after = test_config.get('pause_after', False)
        
        # Wait until start time
        elapsed = self.wait_until(start_time, test_header)
        
        print(f"\n\n=== Starting Test: {test_header} at {elapsed:.2f} minutes ===")
        print(f"Duration: {duration:.2f} minutes")
        self.logger.info(f"Starting test: {test_header} for {duration} min (at {elapsed:.2f} min)")
        
        # Get start time string
        start_time_str = get_formatted_start_time(elapsed)
        
        # Collect data
        samples = data_collector_func(
            self.logger,
            minutes=duration,
            global_timer_start=self.global_start_time,
            test_header=test_header,
            **collector_kwargs
        )
        
        # Handle pause
        if pause_after:
            elapsed = self.get_elapsed_time()
            print(f"\nGlobal timer paused at {elapsed:.2f} minutes.")
            self.logger.info(f"Timer paused at {elapsed:.2f} min")
            
            if self._pause_handler:
                self._pause_handler()
            else:
                input("Press Enter to continue to the next test...")
            
            # Adjust global start time to account for pause
            self.global_start_time = time.time() - (elapsed * 60)
            print("Global timer resumed.\n")
            self.logger.info("Timer resumed")
        
        final_elapsed = self.get_elapsed_time()
        print(f"=== Test Complete: {test_header} ===\n")
        
        return samples, start_time_str, final_elapsed
    
    def adjust_global_timer(self, elapsed_minutes: float):
        """
        Adjust global timer (useful for manual pauses).
        
        Args:
            elapsed_minutes: Current elapsed time in minutes
        """
        self.global_start_time = time.time() - (elapsed_minutes * 60)
        self.logger.debug(f"Adjusted global timer to {elapsed_minutes:.2f} min")


__all__ = ['TestRunner']
