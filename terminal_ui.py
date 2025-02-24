import os
import time
from colorama import init, Fore, Style

init()  # Initialize colorama

class TerminalUI:
    LOGO = f"""{Fore.CYAN}
╭──────────────────────────────────────╮
│                                      │
│   ▄▄▌ ▐ ▄▌      ▄▄▄  ·▄▄▄▄        ·│
│   ██· █▌▐█ ▄█▀▄ ▀▄ █·██▪ ██  ▄█▀▄  │
│   ██▪▐█▐▐▌▐█▌.▐▌▐▀▀▄ ▐█· ▐█▌▐█▌.▐▌ │
│   ▐█▌██▐█▌▐█▌.▐▌▐█•█▌██. ██ ▐█▌.▐▌ │
│    ▀▀▀▀ ▀▪ ▀█▄▀▪.▀  ▀▀▀▀▀▀•  ▀█▄▀▪ │
│                                      │
│        Word Automation Tool          │
╰──────────────────────────────────────╯{Style.RESET_ALL}"""

    PROMPT = f"{Fore.GREEN}wordo>{Style.RESET_ALL} "
    
    @staticmethod
    def clear_screen():
        os.system('cls' if os.name == 'nt' else 'clear')
    
    @staticmethod
    def show_logo():
        TerminalUI.clear_screen()
        print(TerminalUI.LOGO)
        print(f"\n{Fore.YELLOW}Type 'help' for available commands or 'exit' to quit{Style.RESET_ALL}\n")
    
    @staticmethod
    def format_button_list(buttons_data):
        if isinstance(buttons_data, str):  # Error message
            return f"{Fore.RED}{buttons_data}{Style.RESET_ALL}"
            
        output = [f"\n{Fore.CYAN}Available Buttons:{Style.RESET_ALL}\n"]
        
        for bar_name, controls in buttons_data.items():
            output.append(f"\n{Fore.YELLOW}■ {bar_name}{Style.RESET_ALL}")
            for button in controls:
                status = f"{Fore.GREEN}●{Style.RESET_ALL}" if button['enabled'] else f"{Fore.RED}●{Style.RESET_ALL}"
                output.append(f"  {status} {button['caption']} (ID: {button['id']})")
        
        return "\n".join(output)
    
    @staticmethod
    def show_error(message):
        return f"{Fore.RED}Error: {message}{Style.RESET_ALL}"
    
    @staticmethod
    def show_success(message):
        return f"{Fore.GREEN}{message}{Style.RESET_ALL}"
    
    @staticmethod
    def show_info(message):
        return f"{Fore.CYAN}{message}{Style.RESET_ALL}"
