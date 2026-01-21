# This module handles menu and UI display for the pywerHelper package.
# It includes functions for ASCII art display and menu interactions.
# Author: Dylan Pope
# Date: 2024-12-24
# Version: 1.0.0

from dataclasses import dataclass
from typing import Callable, Optional, Dict


def display_ascii_art():
    """Display ASCII art for pywerMeter."""
    art = r"""
    ____                          __  ___      __           
   / __ \__  ___      _____  ____/  |/  /__  / /____  _____
  / /_/ / / / / | /| / / _ \/ __/ /|_/ / _ \/ __/ _ \/ ___/
 / ____/ /_/ /| |/ |/ /  __/ / / /  / /  __/ /_/  __/ /    
/_/    \__, / |__/|__/\___/_/ /_/  /_/\___/\__/\___/_/     
      /____/                                                
    """
    print(art)


def display_menu(menu_options):
    """
    Display a menu with numbered options.
    
    Args:
        menu_options (dict): Dictionary with keys as option numbers and values as option descriptions
        
    Returns:
        str: The selected option key
    """
    print("\n" + "="*60)
    for key, description in menu_options.items():
        print(f"[{key}] {description}")
    print("="*60)
    
    while True:
        choice = input("\nSelect an option: ").strip()
        if choice in menu_options:
            return choice
        else:
            print(f"Invalid option. Please select from {list(menu_options.keys())}")


@dataclass
class MenuItem:
    """Represents a menu item with key, description, and optional action."""
    key: str
    description: str
    action: Optional[Callable] = None
    
    def display(self) -> str:
        """Display the menu item formatted."""
        return f"[{self.key}] {self.description}"


class MenuSystem:
    """Enhanced menu system with action handlers and better organization."""
    
    def __init__(self, title: str = "Menu"):
        """
        Initialize menu system.
        
        Args:
            title: Menu title to display
        """
        self.title = title
        self.items: Dict[str, MenuItem] = {}
    
    def add_item(self, item: MenuItem):
        """
        Add a menu item.
        
        Args:
            item: MenuItem to add
        """
        self.items[item.key] = item
    
    def add_items_from_dict(self, items_dict: Dict[str, str]):
        """
        Add menu items from a dictionary.
        
        Args:
            items_dict: Dictionary of {key: description}
        """
        for key, description in items_dict.items():
            self.add_item(MenuItem(key=key, description=description))
    
    def display(self) -> str:
        """
        Display all menu items and get user choice.
        
        Returns:
            User's choice (validated key)
        """
        print(f"\n{self.title}")
        print("=" * 60)
        
        for item in self.items.values():
            print(item.display())
        
        print("=" * 60)
        
        while True:
            choice = input("\nSelect an option: ").strip()
            if choice in self.items:
                return choice
            
            valid_keys = list(self.items.keys())
            print(f"Invalid option. Please select from: {', '.join(valid_keys)}")
    
    def execute(self, choice: str) -> Optional[any]:
        """
        Execute the action for a given choice.
        
        Args:
            choice: Menu item key to execute
            
        Returns:
            Result of action if defined, None otherwise
        """
        item = self.items.get(choice)
        if item and item.action:
            return item.action()
        return None


# Module-level exports
__all__ = ['display_ascii_art', 'display_menu', 'MenuItem', 'MenuSystem']

