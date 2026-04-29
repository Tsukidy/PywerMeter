"""Tests for menuHelper module."""
import pytest
from io import StringIO
from unittest.mock import patch
from pywerHelper.menuHelper import MenuItem, MenuSystem, display_ascii_art, display_menu


class TestMenuItem:
    """Test MenuItem dataclass."""
    
    def test_create_menu_item(self):
        """Test creating a MenuItem."""
        def dummy_action():
            return "executed"
        
        item = MenuItem(key="1", description="Test Item", action=dummy_action)
        
        assert item.key == "1"
        assert item.description == "Test Item"
        assert item.action == dummy_action
        assert item.action() == "executed"
    
    def test_menu_item_without_action(self):
        """Test creating a MenuItem without action."""
        item = MenuItem(key="x", description="Exit", action=None)
        
        assert item.key == "x"
        assert item.description == "Exit"
        assert item.action is None
    
    def test_menu_item_display(self):
        """Test MenuItem display method."""
        item = MenuItem(key="1", description="Option 1")
        
        assert item.display() == "[1] Option 1"


class TestMenuSystem:
    """Test MenuSystem class."""
    
    def test_create_menu_system(self):
        """Test creating a MenuSystem."""
        menu = MenuSystem("Test Menu")
        
        assert menu.title == "Test Menu"
        assert len(menu.items) == 0
    
    def test_add_single_item(self):
        """Test adding a single menu item."""
        menu = MenuSystem("Test Menu")
        
        def action1():
            return "action1"
        
        item = MenuItem("1", "Option 1", action1)
        menu.add_item(item)
        
        assert len(menu.items) == 1
        assert "1" in menu.items
        assert menu.items["1"].description == "Option 1"
    
    def test_add_multiple_items(self):
        """Test adding multiple menu items."""
        menu = MenuSystem("Test Menu")
        
        menu.add_item(MenuItem("1", "Option 1", lambda: "a"))
        menu.add_item(MenuItem("2", "Option 2", lambda: "b"))
        menu.add_item(MenuItem("x", "Exit", None))
        
        assert len(menu.items) == 3
    
    def test_add_items_from_dict(self):
        """Test adding items from dictionary."""
        menu = MenuSystem("Test Menu")
        
        items_dict = {
            "1": "First Option",
            "2": "Second Option"
        }
        
        menu.add_items_from_dict(items_dict)
        
        assert len(menu.items) == 2
        assert menu.items["1"].description == "First Option"
        assert menu.items["2"].description == "Second Option"
    
    def test_execute_valid_choice(self):
        """Test executing a valid menu choice."""
        menu = MenuSystem("Test Menu")
        
        result_holder = {"value": None}
        
        def action1():
            result_holder["value"] = "executed"
            return "success"
        
        menu.add_item(MenuItem("1", "Option 1", action1))
        
        result = menu.execute("1")
        
        assert result == "success"
        assert result_holder["value"] == "executed"
    
    def test_execute_invalid_choice(self):
        """Test executing an invalid menu choice."""
        menu = MenuSystem("Test Menu")
        menu.add_item(MenuItem("1", "Option 1", lambda: "success"))
        
        result = menu.execute("9")  # Invalid choice
        
        assert result is None
    
    def test_execute_no_action(self):
        """Test executing item with no action (like exit)."""
        menu = MenuSystem("Test Menu")
        menu.add_item(MenuItem("x", "Exit", None))
        
        result = menu.execute("x")
        
        assert result is None
    
    def test_menu_with_callable_actions(self):
        """Test menu with various callable actions."""
        menu = MenuSystem("Test Menu")
        
        # Regular function
        def func1():
            return "func1"
        
        # Lambda
        lambda1 = lambda: "lambda1"
        
        # Class method
        class TestClass:
            def method(self):
                return "method"
        
        obj = TestClass()
        
        menu.add_item(MenuItem("1", "Function", func1))
        menu.add_item(MenuItem("2", "Lambda", lambda1))
        menu.add_item(MenuItem("3", "Method", obj.method))
        
        assert menu.execute("1") == "func1"
        assert menu.execute("2") == "lambda1"
        assert menu.execute("3") == "method"


class TestDisplayMenu:
    """Test display_menu function."""
    
    @patch('builtins.input', side_effect=['1'])
    def test_display_menu_valid_choice(self, mock_input):
        """Test display_menu with valid choice."""
        menu_options = {
            "1": "Option 1",
            "2": "Option 2"
        }
        
        choice = display_menu(menu_options)
        
        assert choice == "1"
    
    @patch('builtins.input', side_effect=['invalid', '2'])
    def test_display_menu_invalid_then_valid(self, mock_input):
        """Test display_menu with invalid then valid choice."""
        menu_options = {
            "1": "Option 1",
            "2": "Option 2"
        }
        
        choice = display_menu(menu_options)
        
        assert choice == "2"
        assert mock_input.call_count == 2


class TestDisplayAsciiArt:
    """Test display_ascii_art function."""
    
    def test_ascii_art_displays(self, capsys):
        """Test that ASCII art displays without error."""
        display_ascii_art()
        
        captured = capsys.readouterr()
        output = captured.out
        
        # Check that something was printed
        assert len(output) > 0
        # Check for part of the ASCII art
        assert "pywerMeter" in output or "____" in output
