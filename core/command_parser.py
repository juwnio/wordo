import json
import re

class CommandParser:
    def __init__(self, commands_file):
        with open(commands_file, 'r') as f:
            self.commands = json.load(f)

    def parse_command(self, user_input):
        parts = user_input.split()
        if not parts:
            return None, None
        
        command = parts[0]
        args = ' '.join(parts[1:])

        # Search for command in all categories
        for category in self.commands:
            if command in self.commands[category]:
                template = self.commands[category][command]
                try:
                    # Extract parameters from command template
                    params = re.findall(r'\{(\w+)\}', template)
                    # Parse arguments based on parameters
                    if params:
                        arg_values = self._parse_args(args)
                        # Replace parameters with actual values
                        for param in params:
                            placeholder = '{' + param + '}'
                            if param in arg_values:
                                template = template.replace(placeholder, str(arg_values[param]))
                    return template, True
                except Exception as e:
                    return f"Error parsing command: {str(e)}", False
        
        return "Command not found", False

    def _parse_args(self, args_string):
        args_dict = {}
        # Match patterns like key=value or "key=value with spaces"
        pattern = r'(\w+)=(?:"([^"]*)"|([\w\/\\\:\.]+))'
        matches = re.finditer(pattern, args_string)
        
        for match in matches:
            key = match.group(1)
            # Use quoted value if present, otherwise use non-quoted value
            value = match.group(2) if match.group(2) is not None else match.group(3)
            args_dict[key] = value
            
        return args_dict

    def get_help(self):
        help_text = "Available commands:\n"
        for category in self.commands:
            help_text += f"\n{category.upper()}:\n"
            for command in self.commands[category]:
                help_text += f"  {command}\n"
        return help_text
