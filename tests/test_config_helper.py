"""Tests for configHelper module."""
import pytest
import os
import tempfile
import yaml
from pywerHelper.configHelper import ConfigManager


@pytest.fixture
def sample_config():
    """Create a sample config dictionary."""
    return {
        "log_file_location": "logs/",
        "test_1_start_time": 0,
        "test_1_duration": 1.5,
        "connection_settings": {
            "port": "COM3",
            "baudrate": 9600,
            "timeout": 1
        },
        "log_settings": {
            "location": "logs/",
            "level": "DEBUG"
        },
        "command_settings": {
            "send_interval": 1,
            "read_delay": 0.1
        },
        "test_settings": {
            "test_1": {
                "start_time": 0,
                "duration": 1.5
            }
        }
    }


@pytest.fixture
def temp_config_file(sample_config):
    """Create a temporary config file."""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.yaml', delete=False) as f:
        yaml.dump(sample_config, f)
        temp_path = f.name
    
    yield temp_path
    
    # Cleanup
    if os.path.exists(temp_path):
        os.unlink(temp_path)


class TestConfigManager:
    """Test ConfigManager class."""
    
    def test_init_with_custom_path(self, temp_config_file):
        """Test initialization with custom config path."""
        manager = ConfigManager(temp_config_file)
        assert manager.config_path == temp_config_file
        assert manager._config is None  # Not loaded yet (lazy loading)
    
    def test_lazy_loading(self, temp_config_file):
        """Test that config is loaded lazily."""
        manager = ConfigManager(temp_config_file)
        assert manager._config is None
        
        # Access config property triggers loading
        config = manager.config
        assert manager._config is not None
        assert isinstance(config, dict)
    
    def test_get_value(self, temp_config_file):
        """Test getting configuration values."""
        manager = ConfigManager(temp_config_file)
        
        assert manager.get("test_1_start_time") == 0
        assert manager.get("test_1_duration") == 1.5
        assert manager.get("log_file_location") == "logs/"
    
    def test_get_with_default(self, temp_config_file):
        """Test getting value with default fallback."""
        manager = ConfigManager(temp_config_file)
        
        assert manager.get("nonexistent_key", "default") == "default"
        assert manager.get("test_1_start_time", 999) == 0  # Should get actual value
    
    def test_get_test_settings(self, temp_config_file):
        """Test retrieving test settings."""
        manager = ConfigManager(temp_config_file)
        
        settings = manager.get_test_settings()
        assert isinstance(settings, dict)
        assert "test_1" in settings
    
    def test_get_serial_settings(self, temp_config_file):
        """Test retrieving serial/connection settings."""
        manager = ConfigManager(temp_config_file)
        
        settings = manager.get_serial_settings()
        assert settings["port"] == "COM3"
        assert settings["baudrate"] == 9600
        assert settings["timeout"] == 1
    
    def test_get_log_settings(self, temp_config_file):
        """Test retrieving log settings."""
        manager = ConfigManager(temp_config_file)
        
        settings = manager.get_log_settings()
        assert settings["location"] == "logs/"
        assert settings["level"] == "DEBUG"
    
    def test_get_command_settings(self, temp_config_file):
        """Test retrieving command settings."""
        manager = ConfigManager(temp_config_file)
        
        settings = manager.get_command_settings()
        assert settings["send_interval"] == 1
        assert settings["read_delay"] == 0.1
    
    def test_config_reload_prevention(self, temp_config_file):
        """Test that config is not reloaded on subsequent access."""
        manager = ConfigManager(temp_config_file)
        
        config1 = manager.config
        config2 = manager.config
        
        # Should be the same object (not reloaded)
        assert config1 is config2
    
    def test_get_missing_section_returns_empty_dict(self, temp_config_file):
        """Test that getting missing section returns empty dict."""
        manager = ConfigManager(temp_config_file)
        
        # Request non-existent section
        result = manager.get("nonexistent_section", {})
        assert result == {}
