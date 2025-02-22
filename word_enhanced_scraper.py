import win32gui
import win32api
import win32con
import win32com.client
from collections import defaultdict
import json
import re

class WordEnhancedScraper:
    def __init__(self, word_app=None):
        self.elements = defaultdict(list)
        self.window_handle = None
        self.word_app = word_app
        self.document_state = {}
        
    def get_document_state(self):
        """Get detailed state of the active document"""
        if not self.word_app:
            return "Word application not provided"
            
        doc = self.word_app.ActiveDocument
        selection = self.word_app.Selection
        
        state = {
            "document": {
                "name": doc.Name,
                "path": doc.Path,
                "pages": doc.ComputeStatistics(2),  # wdStatisticPages
                "paragraphs": doc.Paragraphs.Count,
                "words": doc.Words.Count,
                "characters": doc.Characters.Count,
                "sections": doc.Sections.Count,
                "saved": doc.Saved,
                "track_changes": doc.TrackRevisions,
                "protection_type": doc.ProtectionType,
            },
            "view": {
                "zoom_percentage": self.word_app.ActiveWindow.View.Zoom.Percentage,
                "draft_view": self.word_app.ActiveWindow.View.Draft,
                "show_revisions": self.word_app.ActiveWindow.View.ShowRevisionsAndComments,
                "show_comments": self.word_app.ActiveWindow.View.ShowComments,
                "show_formatting": self.word_app.ActiveWindow.View.ShowFormatting,
            },
            "selection": {
                "start": selection.Start,
                "end": selection.End,
                "text": selection.Text,
                "type": selection.Type,
                "font": {
                    "name": selection.Font.Name,
                    "size": selection.Font.Size,
                    "bold": selection.Font.Bold,
                    "italic": selection.Font.Italic,
                    "underline": selection.Font.Underline,
                    "color": selection.Font.Color,
                },
                "paragraph": {
                    "alignment": selection.ParagraphFormat.Alignment,
                    "line_spacing": selection.ParagraphFormat.LineSpacing,
                    "space_before": selection.ParagraphFormat.SpaceBefore,
                    "space_after": selection.ParagraphFormat.SpaceAfter,
                    "first_line_indent": selection.ParagraphFormat.FirstLineIndent,
                }
            }
        }
        
        # Get styles in use
        styles_in_use = set()
        for para in doc.Paragraphs:
            styles_in_use.add(para.Style.NameLocal)
        state["styles_in_use"] = list(styles_in_use)
        
        # Get headers and footers
        headers_footers = {
            "headers": [],
            "footers": []
        }
        for section in doc.Sections:
            for header in section.Headers:
                if header.Exists:
                    headers_footers["headers"].append({
                        "type": header.Index,  # 1=Primary, 2=FirstPage, 3=EvenPages
                        "text": header.Range.Text
                    })
            for footer in section.Footers:
                if footer.Exists:
                    headers_footers["footers"].append({
                        "type": footer.Index,
                        "text": footer.Range.Text
                    })
        state["headers_footers"] = headers_footers
        
        # Get table information
        tables = []
        for i in range(doc.Tables.Count):
            table = doc.Tables[i + 1]
            tables.append({
                "rows": table.Rows.Count,
                "columns": table.Columns.Count,
                "style": table.Style,
                "uniform_style": table.Uniform,
                "allow_autofit": table.AllowAutoFit
            })
        state["tables"] = tables
        
        # Get shapes and images
        shapes = []
        for shape in doc.Shapes:
            shapes.append({
                "type": shape.Type,
                "name": shape.Name,
                "width": shape.Width,
                "height": shape.Height,
                "position_x": shape.Left,
                "position_y": shape.Top
            })
        state["shapes"] = shapes
        
        # Get hyperlinks
        hyperlinks = []
        for link in doc.Hyperlinks:
            hyperlinks.append({
                "text": link.TextToDisplay,
                "address": link.Address,
                "target": link.SubAddress
            })
        state["hyperlinks"] = hyperlinks
        
        # Get comments
        comments = []
        if doc.Comments.Count > 0:
            for comment in doc.Comments:
                comments.append({
                    "author": comment.Author,
                    "date": str(comment.Date),
                    "text": comment.Range.Text
                })
        state["comments"] = comments
        
        # Get custom document properties
        custom_props = {}
        for prop in doc.CustomDocumentProperties:
            try:
                custom_props[prop.Name] = prop.Value
            except:
                pass
        state["custom_properties"] = custom_props
        
        return state

    def get_ribbon_state(self):
        """Get state of the Word ribbon"""
        if not self.word_app:
            return "Word application not provided"
            
        ribbon_state = {
            "visible": self.word_app.CommandBars.AdaptiveMenus,
            "minimized": self.word_app.CommandBars.Minimized,
            "active_tab": self.word_app.CommandBars.ActiveMenuBar.Name if self.word_app.CommandBars.ActiveMenuBar else None,
        }
        
        return ribbon_state

    def get_application_state(self):
        """Get state of the Word application"""
        if not self.word_app:
            return "Word application not provided"
            
        app_state = {
            "version": self.word_app.Version,
            "build": self.word_app.Build,
            "caption": self.word_app.Caption,
            "path": self.word_app.Path,
            "startup_path": self.word_app.StartupPath,
            "user_name": self.word_app.UserName,
            "user_initials": self.word_app.UserInitials,
            "user_address": self.word_app.UserAddress,
            "default_save_format": self.word_app.DefaultSaveFormat,
            "system": {
                "operating_system": self.word_app.System.OperatingSystem,
                "country": self.word_app.System.Country,
                "language": self.word_app.System.LanguageDesignation,
                "memory_free": self.word_app.System.FreeDiskSpace,
                "memory_total": self.word_app.System.TotalPhysicalMemory,
            }
        }
        
        return app_state

    def get_full_state(self):
        """Get complete state of Word"""
        full_state = {
            "ui_elements": self.elements,
            "document": self.get_document_state(),
            "ribbon": self.get_ribbon_state(),
            "application": self.get_application_state()
        }
        
        return json.dumps(full_state, indent=2)

    def get_accessibility_info(self):
        """Get accessibility information about UI elements"""
        if not self.window_handle:
            return "No Word window found"
            
        accessibility_info = {}
        
        for element_type, elements in self.elements.items():
            for element in elements:
                if element['visible']:
                    # Get role and state information
                    hwnd = element['handle']
                    role = win32gui.GetClassName(hwnd)
                    state = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    
                    accessibility_info[element['text'] or f"Element_{hwnd}"] = {
                        "role": role,
                        "state": {
                            "enabled": bool(state & win32con.WS_ENABLED),
                            "visible": bool(state & win32con.WS_VISIBLE),
                            "focused": hwnd == win32gui.GetFocus(),
                        },
                        "location": element['position'],
                    }
        
        return accessibility_info