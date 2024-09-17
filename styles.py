class Style:
    def __init__(self):
        self.bg_color = "#FFE5B4"
        self.light_bg_color = "#FDF3EC"

        self.entry_background_color = "#FDF3EC"
        self.entry_border_color = "#FFFFFF"
        self.active_entry_border_color = "#F89880"
        self.entry_border_size = 3
        self.disabled_background_entry = "#FDF3EC"  # "#E35335"
        self.disabled_foreground_entry = "#9f9f9f"  # "#FF9F81"

        self.drop_background_color = "#FDF3EC"
        self.highlight_drop_color = "#ffffff"
        self.drop_border_size = 3
        self.active_drop_border_color = "#EDE3DC"
        self.disabled_foreground_drop_color = "#9f9f9f"

        self.button_background = "#CD7F32"
        self.button_foreground = "#ffffff"
        self.button_border_size = 2
        self.button_highlight_background = "#DD8F42"
        self.button_highlight_foreground = "#000000"
        self.button_disabled_foreground = "#bfbfbf"
        self.button_active_background = "#DD8F42"
        self.button_active_foreground = "#000000"
        self.bold_font = ('Helvetica', 10, 'bold')


if __name__ == "__main__":
    Style()
