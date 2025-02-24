class ButtonController:
    def __init__(self, word_app):
        self.word = word_app
        
    def get_available_buttons(self):
        """List all available command bars and their controls"""
        buttons = {}
        try:
            for bar in self.word.CommandBars:
                controls = []
                for control in bar.Controls:
                    if control.Type == 1:  # Type 1 is button
                        controls.append({
                            "id": control.ID,
                            "caption": control.Caption,
                            "enabled": control.Enabled
                        })
                if controls:
                    buttons[bar.Name] = controls
            return buttons
        except Exception as e:
            return f"Error getting buttons: {str(e)}"

    def click_button(self, button_caption):
        """Click a button by its caption"""
        try:
            for bar in self.word.CommandBars:
                for control in bar.Controls:
                    if control.Type == 1 and control.Caption.lower() == button_caption.lower():
                        if control.Enabled:
                            control.Execute()
                            return f"Clicked button: {button_caption}"
                        else:
                            return f"Button '{button_caption}' is disabled"
            return f"Button '{button_caption}' not found"
        except Exception as e:
            return f"Error clicking button: {str(e)}"

    def click_button_by_id(self, button_id):
        """Click a button by its ID"""
        try:
            for bar in self.word.CommandBars:
                for control in bar.Controls:
                    if control.Type == 1 and control.ID == int(button_id):
                        if control.Enabled:
                            control.Execute()
                            return f"Clicked button with ID: {button_id}"
                        else:
                            return f"Button with ID '{button_id}' is disabled"
            return f"Button with ID '{button_id}' not found"
        except Exception as e:
            return f"Error clicking button: {str(e)}"
