import win32gui
import win32ui
from win32api import GetSystemMetrics
import win32con
import json

class WordScraper:
    def __init__(self, word_app):
        self.word = word_app
        
    def get_window_info(self):
        """Get information about the Word window"""
        try:
            window_handle = win32gui.FindWindow("OpusApp", None)
            if not window_handle:
                return "Word window not found"
                
            # Get window rect
            left, top, right, bottom = win32gui.GetWindowRect(window_handle)
            
            window_info = {
                "handle": window_handle,
                "title": win32gui.GetWindowText(window_handle),
                "position": {
                    "left": left,
                    "top": top,
                    "right": right,
                    "bottom": bottom
                },
                "size": {
                    "width": right - left,
                    "height": bottom - top
                }
            }
            return window_info
        except Exception as e:
            return f"Error getting window info: {str(e)}"
            
    def get_ui_elements(self):
        """Get information about UI elements"""
        try:
            ui_elements = {
                "ribbon": self._get_ribbon_info(),
                "statusbar": self._get_statusbar_info(),
                "document": self._get_document_info()
            }
            return ui_elements
        except Exception as e:
            return f"Error getting UI elements: {str(e)}"
            
    def _get_ribbon_info(self):
        """Get information about the ribbon"""
        try:
            ribbon = self.word.CommandBars
            ribbon_info = {
                "visible": ribbon.Visible,
                "count": ribbon.Count,
                "height": ribbon.Height
            }
            return ribbon_info
        except:
            return "Ribbon information not available"
            
    def _get_statusbar_info(self):
        """Get information about the status bar"""
        try:
            statusbar = self.word.Application.StatusBar
            return {"text": statusbar}
        except:
            return "Status bar information not available"
            
    def _get_document_info(self):
        """Get information about the active document"""
        try:
            doc = self.word.ActiveDocument
            if doc:
                doc_info = {
                    "name": doc.Name,
                    "path": doc.Path,
                    "pages": doc.ComputeStatistics(2),
                    "words": doc.Words.Count,
                    "characters": doc.Characters.Count,
                    "paragraphs": doc.Paragraphs.Count
                }
                return doc_info
            return "No active document"
        except:
            return "Document information not available"
            
    def scrap_word_window(self):
        """Main scraping function that collects all available information"""
        try:
            result = {
                "window": self.get_window_info(),
                "ui_elements": self.get_ui_elements(),
                "screen_resolution": {
                    "width": GetSystemMetrics(0),
                    "height": GetSystemMetrics(1)
                }
            }
            
            # Save results to a JSON file
            with open('word_scrape_results.json', 'w') as f:
                json.dump(result, f, indent=4)
                
            return "Scraping completed. Results saved to word_scrape_results.json"
        except Exception as e:
            return f"Error during scraping: {str(e)}"
