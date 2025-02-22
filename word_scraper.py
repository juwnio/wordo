import win32gui
import win32api
import win32con
from collections import defaultdict
import re

class WordScraper:
    def __init__(self):
        self.elements = defaultdict(list)
        self.window_handle = None
        
    def find_word_window(self):
        """Find the main Word window handle"""
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if "Word" in window_text or ".docx" in window_text.lower():
                    hwnds.append(hwnd)
            return True
            
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None

    def get_control_type(self, hwnd):
        """Determine the type of control based on window class and style"""
        class_name = win32gui.GetClassName(hwnd)
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        
        if "Button" in class_name:
            if style & win32con.BS_CHECKBOX:
                return "checkbox"
            elif style & win32con.BS_RADIOBUTTON:
                return "radio"
            else:
                return "button"
        elif "Edit" in class_name:
            return "textbox"
        elif "Combobox" in class_name:
            return "combobox"
        elif "ListBox" in class_name:
            return "listbox"
        elif "Static" in class_name:
            return "label"
        elif "ScrollBar" in class_name:
            return "scrollbar"
        elif "Toolbar" in class_name:
            return "toolbar"
        else:
            return "unknown"

    def get_element_info(self, hwnd):
        """Get detailed information about a window element"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            control_type = self.get_control_type(hwnd)
            
            return {
                'handle': hwnd,
                'text': text,
                'class': class_name,
                'type': control_type,
                'position': {
                    'left': rect[0],
                    'top': rect[1],
                    'right': rect[2],
                    'bottom': rect[3],
                    'width': rect[2] - rect[0],
                    'height': rect[3] - rect[1]
                },
                'visible': win32gui.IsWindowVisible(hwnd)
            }
        except Exception as e:
            return None

    def enum_child_windows(self, hwnd):
        """Enumerate all child windows recursively"""
        def callback(child_hwnd, param):
            info = self.get_element_info(child_hwnd)
            if info:
                self.elements[info['type']].append(info)
            win32gui.EnumChildWindows(child_hwnd, callback, None)
            return True
            
        win32gui.EnumChildWindows(hwnd, callback, None)

    def scrap_word_window(self):
        """Main method to scrap the Word window"""
        self.elements.clear()
        self.window_handle = self.find_word_window()
        
        if not self.window_handle:
            return "No Word window found"
            
        # Get main window info
        main_info = self.get_element_info(self.window_handle)
        if main_info:
            self.elements['window'] = [main_info]
            
        # Get all child windows
        self.enum_child_windows(self.window_handle)
        
        return self.format_results()

    def format_results(self):
        """Format the scraping results for display"""
        result = []
        result.append("\n=== Word Window UI Elements ===\n")
        
        # First display main window info
        if 'window' in self.elements:
            window = self.elements['window'][0]
            result.append(f"Main Window:")
            result.append(f"  Title: {window['text']}")
            result.append(f"  Position: {window['position']}")
            result.append("")
            
        # Then display all other elements grouped by type
        for element_type, elements in self.elements.items():
            if element_type != 'window':
                result.append(f"{element_type.upper()} Elements ({len(elements)}):")
                for elem in elements:
                    if elem['text']:  # Only show elements with text
                        result.append(f"  - {elem['text']}")
                        result.append(f"    Position: (x:{elem['position']['left']}, y:{elem['position']['top']}, " 
                                    f"w:{elem['position']['width']}, h:{elem['position']['height']})")
                result.append("")
                
        return "\n".join(result)

    def get_element_at_position(self, x, y):
        """Find UI element at given coordinates"""
        for elements in self.elements.values():
            for elem in elements:
                pos = elem['position']
                if (pos['left'] <= x <= pos['right'] and 
                    pos['top'] <= y <= pos['bottom']):
                    return elem
        return None