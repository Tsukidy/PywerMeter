"""Tests for menuHelper module."""
import pytest
from pywerHelper.menuHelper import MenuItem, MenuSystem


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
        
        menu.add_item("1", "Option 1", action1)
        
        assert len(menu.items) == 1
        assert menu.items[0].key == "1"
        assert menu.items[0].description == "Option 1"
    
    def test_add_multiple_items(self):
        """Test adding multiple menu items."""
        menu = MenuSystem("Test Menu")
        
        menu.add_item("1", "Option 1", lambda: "a")
        menu.add_item("2", "Option 2", lambda: "b")
        menu.add_item("x", "Exit", None)
        
        assert len(menu.items) == 3
    
    def test_add_items_from_dict(self):
        """Test adding items from dictionary."""
        menu = MenuSystem("Test Menu")
        
        actions = {
            "1": lambda: "action1",
            "2": lambda: "action2"
        }
        
        descriptions = {
            "1": "First Option",
            "2": "Second Option"
        }
        
        menu.add_items_from_dict(actions, descriptions)
        
        assert len(menu.items) == 2
        assert menu.items[0].key == "1"
        assert menu.items[0].description == "First Option"
        assert menu.items[1].key == "2"
        assert menu.items[1].description == "Second Option"
    
    def test_add_items_from_dict_missing_descriptions(self):
        """Test adding items when some descriptions are missing."""
        menu = MenuSystem("Test Menu")
        
        actions = {
            "1": lambda: "action1",
            "2": lambda: "action2",
            "3": lambda: "action3"
        }
        
        descriptions = {
            "1": "First Option"
            # Missing descriptions for "2" and "3"
        }
        
        menu.add_items_from_dict(actions, descriptions)
        
        assert len(menu.items) == 3
        assert menu.items[0].description == "First Option"
        assert menu.items[1].description == ""  # Default empty string
        assert menu.items[2].description == ""
    
    def test_execute_valid_choice(self):
        """Test executing a valid menu choice."""
        menu = MenuSystem("Test Menu")
        
        result_holder = {"value": None}
        
        def action1():
            result_holder["value"] = "executed"
            return "success"
        
        menu.add_item("1", "Option 1", action1)
        
        result = menu.execute("1")
        
        assert result == "success"
        assert result_holder["value"] == "executed"
    
    def test_execute_invalid_choice(self):
        """Test executing an invalid menu choice."""
        menu = MenuSystem("Test Menu")
        menu.add_item("1", "Option 1", lambda: "success")
        
        result = menu.execute("9")  # Invalid choice
        
        assert result is None
    
    def test_execute_no_action(self):
        """Test executing item with no action (like exit)."""
        menu = MenuSystem("Test Menu")
        menu.add_item("x", "Exit", None)
        
        result = menu.execute("x")
        
        assert result is None
    
    def test_display_format(self, capsys):
        """Test that display shows correct format."""
        menu = MenuSystem("Test Menu")
        menu.add_item("1", "First Option", lambda: None)
        menu.add_item("2", "Second Option", lambda: None)
        menu.add_item("x", "Exit", None)
        
        menu.display()
        
        captured = capsys.readouterr()
        output = captured.out
        
        # Check that title and items are displayed
        assert "Test Menu" in output
        assert "[1]" in output
        assert "First Option" in output
        assert "[2]" in output
        assert "Second Option" in output
        assert "[x]" in output
        assert "Exit" in output
    
    def test_empty_menu_display(self, capsys):
        """Test displaying empty menu."""
        menu = MenuSystem("Empty Menu")
        menu.display()
        
        captured = capsys.readouterr()
        output = captured.out
        
        assert "Empty Menu" in output
    
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
        
        menu.add_item("1", "Function", func1)
        menu.add_item("2", "Lambda", lambda1)
        menu.add_item("3", "Method", obj.method)
        
        assert menu.execute("1") == "func1"
        assert menu.execute("2") == "lambda1"
        assert menu.execute("3") == "method"
    
    def test_menu_case_sensitivity(self):
        """Test menu choice case sensitivity."""
        menu = MenuSystem("Test Menu")
        menu.add_item("a", "Option A", lambda: "success")
        
        # Should find exact match
        assert menu.execute("a") == "success"
        
        # Different case should not match
        assert menu.execute("A") is None
