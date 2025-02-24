import win32com.client
from word_scraper import WordScraper
from button_controller import ButtonController
import json

class WordController:
    def __init__(self):
        self.word = None
        self.doc = None
        self.scraper = None
        self.button_controller = None

    def start_word(self):
        if not self.word:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = True
            self.scraper = WordScraper(self.word)
            self.button_controller = ButtonController(self.word)
            return "Microsoft Word started"
        return "Word is already running"

    def check_status(self):
        return "Word is running" if self.word else "Word is not running"

    def minimize_word(self):
        if self.word:
            self.word.WindowState = 2  # Minimized
            return "Word minimized"
        return "Word is not running"

    def maximize_word(self):
        if self.word:
            self.word.WindowState = 1  # Maximized
            return "Word maximized"
        return "Word is not running"

    def hide_word(self):
        if self.word:
            self.word.Visible = False
            return "Word hidden"
        return "Word is not running"

    def show_word(self):
        if self.word:
            self.word.Visible = True
            return "Word shown"
        return "Word is not running"

    def create_document(self):
        self.doc = self.word.Documents.Add()
        return "New document created"

    def open_document(self, path):
        self.doc = self.word.Documents.Open(path)
        return f"Opened document: {path}"

    def save_document(self, path=None):
        if path:
            self.doc.SaveAs(path)
            return f"Document saved as: {path}"
        self.doc.Save()
        return "Document saved"

    def quit_word(self):
        self.word.Quit()
        return "Word closed"

    def write_text(self, text):
        self.word.Selection.TypeText(text)
        return "Text written"

    def add_heading(self, text, level):
        self.word.Selection.TypeText(text)
        self.word.Selection.Style = f"Heading {level}"
        return f"Heading {level} added"

    def find_text(self, text):
        self.word.Selection.Find.Text = text
        self.word.Selection.Find.Execute()
        return f"Found text: {text}"

    def replace_text(self, find, replace):
        self.word.Selection.Find.Text = find
        self.word.Selection.Find.Replacement.Text = replace
        self.word.Selection.Find.Execute(Replace=2)
        return f"Replaced '{find}' with '{replace}'"

    def select_text(self, start, end):
        range = self.doc.Range(start, end)
        range.Select()
        return f"Selected text from {start} to {end}"

    def move_cursor(self, direction, unit, count):
        units = {
            'character': 1,
            'word': 2,
            'sentence': 3,
            'paragraph': 4,
            'line': 5
        }
        if direction == 'left':
            self.word.Selection.MoveLeft(Unit=units[unit], Count=count)
        elif direction == 'right':
            self.word.Selection.MoveRight(Unit=units[unit], Count=count)
        elif direction == 'up':
            self.word.Selection.MoveUp(Unit=units[unit], Count=count)
        elif direction == 'down':
            self.word.Selection.MoveDown(Unit=units[unit], Count=count)
        return f"Moved cursor {direction} by {count} {unit}(s)"

    def format_text(self, format_type, value):
        if format_type == 'bold':
            self.word.Selection.Font.Bold = value
        elif format_type == 'italic':
            self.word.Selection.Font.Italic = value
        elif format_type == 'underline':
            self.word.Selection.Font.Underline = value
        elif format_type == 'size':
            self.word.Selection.Font.Size = value
        elif format_type == 'font':
            self.word.Selection.Font.Name = value
        elif format_type == 'color':
            self.word.Selection.Font.Color = value
        return f"Applied {format_type} formatting"

    def insert_table(self, rows, cols):
        self.word.Selection.Tables.Add(self.word.Selection.Range, rows, cols)
        return f"Inserted table with {rows} rows and {cols} columns"

    def add_picture(self, path):
        self.word.Selection.InlineShapes.AddPicture(path)
        return f"Inserted picture from {path}"

    def print_document(self):
        if self.doc:
            self.doc.PrintOut()
            return "Document sent to printer"
        return "No document open"

    def print_preview(self):
        if self.doc:
            self.doc.PrintPreview()
            return "Showing print preview"
        return "No document open"

    def export_pdf(self, path):
        if self.doc:
            self.doc.SaveAs(path, FileFormat=17)
            return f"Exported as PDF to {path}"
        return "No document open"

    def count_words(self):
        if self.doc:
            return f"Word count: {self.doc.Words.Count}"
        return "No document open"

    def count_pages(self):
        if self.doc:
            return f"Page count: {self.doc.ComputeStatistics(2)}"
        return "No document open"

    def clear_formatting(self):
        if self.word:
            self.word.Selection.ClearFormatting()
            return "Formatting cleared"
        return "Word is not running"

    def select_all(self):
        if self.word:
            self.word.Selection.WholeStory()
            return "All text selected"
        return "Word is not running"

    def align_text(self, alignment):
        if self.word:
            align_values = {
                'left': 0,
                'center': 1,
                'right': 2,
                'justify': 3
            }
            self.word.Selection.ParagraphFormat.Alignment = align_values[alignment]
            return f"Text aligned {alignment}"
        return "Word is not running"

    def add_bullet(self):
        if self.word:
            self.word.Selection.Range.ListFormat.ApplyBulletDefault()
            return "Bullet added"
        return "Word is not running"

    def set_spacing(self, value):
        if self.word:
            self.word.Selection.ParagraphFormat.LineSpacing = value
            return f"Line spacing set to {value}"
        return "Word is not running"

    def scrap(self):
        """Execute scraping of Word window"""
        if self.word and self.scraper:
            return self.scraper.scrap_word_window()
        return "Word is not running"

    def list_buttons(self):
        """List all available buttons"""
        if self.word and self.button_controller:
            buttons = self.button_controller.get_available_buttons()
            return json.dumps(buttons, indent=2)
        return "Word is not running"

    def click_button(self, caption):
        """Click a button by its caption"""
        if self.word and self.button_controller:
            return self.button_controller.click_button(caption)
        return "Word is not running"

    def click_button_by_id(self, button_id):
        """Click a button by its ID"""
        if self.word and self.button_controller:
            return self.button_controller.click_button_by_id(button_id)
        return "Word is not running"
