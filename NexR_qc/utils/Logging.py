import logging
import os
import sys
from datetime import datetime


# Logger 정의
class Logger:
    def __init__(self, file=sys.argv[0], proc_name=None, log_folder_path=None, save=True):
        self.today = datetime.today().strftime(format="%Y%m%d")
        self.created_time = datetime.now().strftime(format="%Y%m%d_%H%M%S")
        self.file_name = os.path.basename(file)
        self.colorSetting = {"grey": "\x1b[38;20m", "blue": "\033[34m", "green": "\033[32m", "yellow": "\x1b[33;20m", "red": "\x1b[31;20m", "bold_red": "\x1b[31;1m", "reset": "\x1b[0m"}

        if not proc_name:
            self.logger = logging.getLogger(f"{self.file_name}")
            self.log_path = os.path.join(log_folder_path, f"{self.file_name}_{self.created_time}.log")
        else:
            self.logger = logging.getLogger(f"{proc_name}")
            self.log_path = os.path.join(log_folder_path, f"{proc_name}_{self.created_time}.log")

        if len(self.logger.handlers) == 0:
            formatter = logging.Formatter("(%(asctime)s) [%(levelname)s] %(message)s")

            # StreamHandler
            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(formatter)
            self.logger.addHandler(stream_handler)

            # FileHandler
            if save == True:
                file_handler = logging.FileHandler(self.log_path)
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)

            self.logger.setLevel(logging.INFO)

    def info(self, value):
        self.logger.info(f"{str(value)} (at {self.file_name})")

    def error(self, value):
        self.logger.error(f"{self.colorSetting['red']}{str(value)}{self.colorSetting['reset']} (at {self.file_name})")
