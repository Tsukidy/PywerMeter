"""Configuration management for pywerMeter."""
import os
import yaml
import sys
import logging
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


class ConfigManager:
    """Centralized configuration management with lazy loading."""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration manager.
        
        Args:
            config_path: Path to config.yaml. If None, looks in script's parent directory.
        """
        if config_path is None:
            script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(script_dir, "config.yaml")
        
        self.config_path = config_path
        self._config: Optional[Dict[str, Any]] = None
    
    @property
    def config(self) -> Dict[str, Any]:
        """Lazy load configuration."""
        if self._config is None:
            self._config = self._load_config()
        return self._config
    
    def _load_config(self) -> Dict[str, Any]:
        """
        Load configuration from YAML file.
        
        Returns:
            Configuration dictionary
            
        Raises:
            SystemExit: If config file not found or invalid
        """
        try:
            with open(self.config_path, 'r') as file:
                config = yaml.safe_load(file)
                logger.info(f"Configuration loaded from {self.config_path}")
                return config
        except FileNotFoundError:
            print(f"ERROR: Configuration file not found: {self.config_path}")
            logger.error(f"Configuration file not found: {self.config_path}")
            sys.exit(1)
        except yaml.YAMLError as e:
            print(f"ERROR: Invalid YAML syntax: {e}")
            logger.error(f"YAML parsing error: {e}", exc_info=True)
            sys.exit(1)
        except PermissionError:
            print(f"ERROR: Permission denied reading: {self.config_path}")
            logger.error(f"Permission denied: {self.config_path}")
            sys.exit(1)
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value by key."""
        return self.config.get(key, default)
    
    def get_test_settings(self) -> Dict[str, Any]:
        """Get test settings dictionary."""
        return self.get('test_settings', {})
    
    def get_serial_settings(self) -> Dict[str, Any]:
        """Get serial connection settings."""
        return self.get('connection_settings', {})
    
    def get_log_settings(self) -> Dict[str, Any]:
        """Get logging settings."""
        return self.get('log_settings', {})
    
    def get_command_settings(self) -> Dict[str, Any]:
        """Get command settings."""
        return self.get('command_settings', {})


# Maintain backward compatibility
def create_config(file_path: str, settings: Dict[str, Dict[str, str]]):
    """
    Legacy function for creating INI config files.
    Kept for backward compatibility but not used in main application.
    """
    import configparser
    config = configparser.ConfigParser()
    for section, options in settings.items():
        config[section] = options
    with open(file_path, 'w') as configfile:
        config.write(configfile)


__all__ = ['ConfigManager', 'create_config']
