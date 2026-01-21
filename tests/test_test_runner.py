"""Tests for testRunner module."""
import pytest
import time
from unittest.mock import Mock, patch, MagicMock
from pywerHelper.testRunner import TestRunner


class TestTestRunner:
    """Test TestRunner class."""
    
    def test_create_test_runner(self):
        """Test creating a TestRunner."""
        runner = TestRunner()
        
        assert runner.start_time == 0
        assert runner.pause_duration == 0
        assert runner.pause_handler is None
    
    def test_set_pause_handler(self):
        """Test setting pause handler."""
        runner = TestRunner()
        handler = Mock()
        
        runner.set_pause_handler(handler)
        
        assert runner.pause_handler is handler
    
    def test_get_elapsed_time_no_pause(self):
        """Test getting elapsed time without pauses."""
        runner = TestRunner()
        runner.start_time = time.time() - 10  # 10 seconds ago
        
        elapsed = runner.get_elapsed_time()
        
        assert 9.9 < elapsed < 10.1  # Allow small timing variance
    
    def test_get_elapsed_time_with_pause(self):
        """Test getting elapsed time with pause duration."""
        runner = TestRunner()
        runner.start_time = time.time() - 20
        runner.pause_duration = 5  # 5 seconds paused
        
        elapsed = runner.get_elapsed_time()
        
        # Should be approximately 15 seconds (20 - 5)
        assert 14.9 < elapsed < 15.1
    
    def test_wait_until_already_passed(self):
        """Test wait_until when target time already passed."""
        runner = TestRunner()
        runner.start_time = time.time() - 10  # Started 10 seconds ago
        
        # Wait until 5 minutes (target already passed)
        start = time.time()
        runner.wait_until(5)
        end = time.time()
        
        # Should return immediately
        assert (end - start) < 0.1
    
    def test_wait_until_short_duration(self):
        """Test wait_until with short duration."""
        runner = TestRunner()
        runner.start_time = time.time()
        
        # Wait 0.2 seconds
        start = time.time()
        runner.wait_until(0.2 / 60)  # Convert to minutes
        end = time.time()
        
        elapsed = end - start
        assert 0.15 < elapsed < 0.3  # Allow some variance
    
    def test_wait_until_with_pause_handler(self):
        """Test wait_until calls pause handler."""
        runner = TestRunner()
        runner.start_time = time.time()
        
        pause_handler = Mock(return_value=0)  # No pause
        runner.set_pause_handler(pause_handler)
        
        runner.wait_until(0.1 / 60)  # 0.1 seconds in minutes
        
        # Pause handler should be called at least once
        assert pause_handler.call_count >= 1
    
    def test_wait_until_accumulates_pause(self):
        """Test that wait_until accumulates pause duration."""
        runner = TestRunner()
        runner.start_time = time.time()
        
        # Mock pause handler that returns 0.1 seconds pause
        def mock_pause_handler():
            time.sleep(0.05)  # Simulate pause
            return 0.05
        
        runner.set_pause_handler(mock_pause_handler)
        runner.wait_until(0.2 / 60)  # 0.2 seconds
        
        # Pause duration should be accumulated
        assert runner.pause_duration > 0
    
    @patch('pywerHelper.testRunner.time')
    def test_run_test_basic(self, mock_time):
        """Test running a basic test."""
        # Mock time to control test execution
        mock_time.time.side_effect = [0, 1, 2, 3]  # Simulate time progression
        mock_time.sleep = Mock()
        
        runner = TestRunner()
        
        test_func = Mock()
        cleanup_func = Mock()
        
        runner.run_test(
            test_name="Test 1",
            duration=0.05,  # 3 seconds in minutes
            test_func=test_func,
            cleanup_func=cleanup_func
        )
        
        test_func.assert_called_once()
        cleanup_func.assert_called_once()
    
    def test_run_test_returns_data(self):
        """Test that run_test returns data from test_func."""
        runner = TestRunner()
        
        def test_func():
            return ["data1", "data2", "data3"]
        
        result = runner.run_test(
            test_name="Test",
            duration=0.01,  # Very short
            test_func=test_func,
            cleanup_func=lambda: None
        )
        
        assert result == ["data1", "data2", "data3"]
    
    def test_run_test_no_cleanup(self):
        """Test running test without cleanup function."""
        runner = TestRunner()
        
        test_func = Mock(return_value=[])
        
        result = runner.run_test(
            test_name="Test",
            duration=0.01,
            test_func=test_func
        )
        
        test_func.assert_called_once()
        assert result == []
    
    def test_adjust_global_timer(self):
        """Test adjusting global timer."""
        runner = TestRunner()
        runner.start_time = 1000.0
        
        # Adjust by adding 5 minutes (300 seconds)
        runner.adjust_global_timer(5.0)
        
        assert runner.start_time == 1000.0 - 300.0
    
    def test_adjust_global_timer_negative(self):
        """Test adjusting global timer with negative value."""
        runner = TestRunner()
        runner.start_time = 1000.0
        
        # Adjust by subtracting 3 minutes
        runner.adjust_global_timer(-3.0)
        
        assert runner.start_time == 1000.0 + 180.0
    
    def test_adjust_global_timer_zero(self):
        """Test adjusting global timer with zero."""
        runner = TestRunner()
        original_start = runner.start_time = 1000.0
        
        runner.adjust_global_timer(0)
        
        assert runner.start_time == original_start


class TestTestRunnerIntegration:
    """Integration tests for TestRunner."""
    
    def test_full_test_workflow(self):
        """Test complete workflow: wait, run test, get elapsed time."""
        runner = TestRunner()
        runner.start_time = time.time()
        
        # Wait a tiny amount
        runner.wait_until(0.05 / 60)  # 0.05 seconds
        
        # Run test
        data_collected = []
        
        def test_func():
            data_collected.append("sample1")
            data_collected.append("sample2")
            return data_collected
        
        result = runner.run_test(
            test_name="Integration Test",
            duration=0.05 / 60,
            test_func=test_func,
            cleanup_func=lambda: None
        )
        
        assert result == ["sample1", "sample2"]
        
        # Get elapsed time
        elapsed = runner.get_elapsed_time()
        assert elapsed > 0
    
    def test_multiple_tests_sequence(self):
        """Test running multiple tests in sequence."""
        runner = TestRunner()
        runner.start_time = time.time()
        
        results = []
        
        for i in range(3):
            def test_func(index=i):
                return [f"data_{index}"]
            
            result = runner.run_test(
                test_name=f"Test {i+1}",
                duration=0.02 / 60,  # Very short
                test_func=test_func,
                cleanup_func=lambda: None
            )
            results.append(result)
        
        assert len(results) == 3
        assert results[0] == ["data_0"]
        assert results[1] == ["data_1"]
        assert results[2] == ["data_2"]
