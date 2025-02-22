import os
import json
from command_parser import CommandParser
from word_controller import WordController
from word_scraper import WordScraper
from word_enhanced_scraper import WordEnhancedScraper

class Wordo:
    def __init__(self):
        # Initialize the Word controller and scrapers
        self.controller = WordController()
        self.scraper = WordScraper()
        self.enhanced_scraper = WordEnhancedScraper(self.controller.word)
        
        # Create commands dictionary from word_commands.txt
        self.commands = self._parse_commands_file()
        
        # Create commands.json for the parser
        with open('commands.json', 'w') as f:
            json.dump(self.commands, f, indent=4)
            
        # Initialize the command parser
        self.parser = CommandParser('commands.json')
        
    def _parse_commands_file(self):
        commands = {
            'Analysis Operations': {
                'scrap': 'scrap'  # Explicitly add the scrap command
            }
        }
        current_category = None
        
        with open('word_commands.txt', 'r') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('!!'):  # Skip empty lines and examples
                    continue
                    
                if line.startswith('! '):  # Category marker
                    current_category = line[2:].strip(':')
                    commands[current_category] = {}
                elif line.startswith('!'):  # Skip commented commands
                    continue
                else:  # Command line
                    if current_category:
                        # Split command and extract base command name
                        parts = line.strip().split()
                        if parts:
                            command = parts[0]
                            # Create template with full command
                            commands[current_category][command] = line.strip()
        
        return commands

    def execute_command(self, command_str):
        if not command_str:
            return "Please enter a command"
            
        if command_str.lower() == 'help':
            return self.parser.get_help()
            
        if command_str.lower() == 'exit':
            return None
            
        # Parse the command
        template, success = self.parser.parse_command(command_str)
        
        if not success:
            return template  # Error message
            
        # Map command to controller method
        try:
            # Extract base command name
            base_command = command_str.split()[0]
            
            # Print parsed command for debugging
            print(f"Debug - Command received: {command_str}")
            print(f"Debug - Base command: {base_command}")
            
            # Map commands to controller methods
            command_mapping = {
                'start_word': self.controller.start_word,
                'check_status': self.controller.check_status,
                'minimize': self.controller.minimize_word,
                'maximize': self.controller.maximize_word,
                'hide': self.controller.hide_word,
                'show': self.controller.show_word,
                'create_new': self.controller.create_document,
                'open': self.controller.open_document,
                'save': self.controller.save_document,
                'save_current': lambda: self.controller.save_document(),
                'close': self.controller.quit_word,
                'write': lambda text: self.controller.write_text(text),
                'add_heading': lambda text, level: self.controller.add_heading(text, int(level)),
                'find': lambda text: self.controller.find_text(text),
                'replace': lambda find, replace: self.controller.replace_text(find, replace),
                'select': lambda start, end: self.controller.select_text(int(start), int(end)),
                'clear_formatting': self.controller.clear_formatting,
                'select_all': self.controller.select_all,
                'align_left': lambda: self.controller.align_text('left'),
                'align_right': lambda: self.controller.align_text('right'),
                'align_center': lambda: self.controller.align_text('center'),
                'align_justify': lambda: self.controller.align_text('justify'),
                'add_bullet': self.controller.add_bullet,
                'set_spacing': lambda value: self.controller.set_spacing(float(value)),
                'move_left': lambda unit, count: self.controller.move_cursor('left', unit, int(count)),
                'move_right': lambda unit, count: self.controller.move_cursor('right', unit, int(count)),
                'move_up': lambda unit, count: self.controller.move_cursor('up', unit, int(count)),
                'move_down': lambda unit, count: self.controller.move_cursor('down', unit, int(count)),
                'bold': lambda value: self.controller.format_text('bold', value.lower() == 'true'),
                'italic': lambda value: self.controller.format_text('italic', value.lower() == 'true'),
                'underline': lambda value: self.controller.format_text('underline', value.lower() == 'true'),
                'font_size': lambda size: self.controller.format_text('size', float(size)),
                'font_name': lambda name: self.controller.format_text('font', name),
                'font_color': lambda color: self.controller.format_text('color', int(color)),
                'insert_table': lambda rows, cols: self.controller.insert_table(int(rows), int(cols)),
                'insert_picture': lambda path: self.controller.add_picture(path),
                'print_doc': self.controller.print_document,
                'print_preview': self.controller.print_preview,
                'export_pdf': self.controller.export_pdf,
                'count_words': self.controller.count_words,
                'count_pages': self.controller.count_pages,
                'scrap': lambda: self.scraper.scrap_word_window(),
                'scrap_document': lambda: self.enhanced_scraper.get_document_state(),
                'scrap_ribbon': lambda: self.enhanced_scraper.get_ribbon_state(),
                'scrap_application': lambda: self.enhanced_scraper.get_application_state(),
                'scrap_full': lambda: self.enhanced_scraper.get_full_state(),
                'scrap_accessibility': lambda: self.enhanced_scraper.get_accessibility_info(),
            }
            
            if base_command not in command_mapping:
                print(f"Debug - Available commands: {list(command_mapping.keys())}")
                return f"Command '{base_command}' not implemented"
                
            # Parse arguments from the command template
            args = self.parser._parse_args(command_str)
            
            # Execute the command with its arguments
            method = command_mapping[base_command]
            if args:
                return method(**args)
            else:
                return method()
                
        except Exception as e:
            return f"Error executing command: {str(e)}"

    def run(self):
        print("Welcome to Wordo - Word Automation Tool")
        print("Type 'help' for available commands or 'exit' to quit")
        
        while True:
            try:
                command = input("\nwordo> ").strip()
                result = self.execute_command(command)
                
                if result is None:  # Exit command
                    print("Goodbye!")
                    break
                    
                print(result)
                
            except KeyboardInterrupt:
                print("\nInterrupted by user. Exiting...")
                break
            except Exception as e:
                print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    wordo = Wordo()
    wordo.run()