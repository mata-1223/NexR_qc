import os
import sys
import logging
from datetime import datetime

# Logger 정의
class Logger:
    def __init__(self, file = sys.argv[0], proc_name = None, log_folder_path = None):
        self.today = datetime.today().strftime(format = '%Y%m%d')
        self.created_time=datetime.now().strftime(format='%Y%m%d_%H%M%S')
        self.file_name = os.path.basename(file)
        
        if not proc_name:    
            self.logger = logging.getLogger(f"{self.file_name}")
            self.log_path = os.path.join(log_folder_path, f"{self.file_name}_{self.created_time}.log")
        else:
            self.logger = logging.getLogger(f"{proc_name}")
            self.log_path = os.path.join(log_folder_path, f"{proc_name}_{self.created_time}.log")
            
        if len(self.logger.handlers) == 0:
            formatter = logging.Formatter(u'(%(asctime)s) [%(levelname)s] %(message)s')

            # StreamHandler
            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(formatter)
            
            # FileHandler
            file_handler = logging.FileHandler(self.log_path)
            file_handler.setFormatter(formatter)

            self.logger.addHandler(stream_handler)
            self.logger.addHandler(file_handler)
            self.logger.setLevel(logging.INFO)

    def info(self, value):
        self.logger.info("%s (at %s)" % (str(value), self.file_name))

    def error(self, value):
        self.logger.error("%s (at %s)" % (str(value), self.file_name))
