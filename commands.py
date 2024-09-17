class Commands:
    def __init__(self):
        self.enableCommand = "E\n"
        self.disableCommand = "D\n"
        self.clockwiseCommand = "C\n"
        self.antiClockwiseCommand = "A\n"
        self.pulseCommand = "P\n"
        self.homeCommand = "H\n"
        self.endCommand = "N\n"
        self.locateCommand = "L\n"
        self.stopCommand = "S\n"
        self.enable_ok = "e\r\n"
        self.disable_ok = "d\r\n"
        self.clockwise_ok = "c\r\n"
        self.antiClockwise_ok = "a\r\n"
        self.pulse_ok = "p\r\n"
        self.home_ok = "h\r\n"
        self.end_ok = "n\r\n"
        self.stop_ok = "s\r\n"
        self.move_ok = "m\r\n"


if __name__ == "__main__":
    Commands()
