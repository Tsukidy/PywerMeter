# This module handles menu and UI display for the pywerHelper package.
# It includes functions for ASCII art display and menu interactions.
# Author: Dylan Pope
# Date: 2024-12-24
# Version: 1.0.0


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


# Module-level exports
__all__ = ['display_ascii_art', 'display_menu']
