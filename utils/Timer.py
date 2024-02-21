import time


class TimerError(Exception):
    """A custom exception used to report errors in use of Timer class"""


class Timer:
    def __init__(self, logger=None):
        self._start_time = None
        if logger:
            self.logger = logger

    def start(self):
        """Start a new timer"""
        if self._start_time is not None:
            raise TimerError(f"Timer is running. Use .stop() to stop it")

        self._start_time = time.perf_counter()

    def stop(self):
        """Stop the timer, and report the elapsed time"""
        if self._start_time is None:
            raise TimerError(f"Timer is not running. Use .start() to start it")

        elapsed_time = time.strftime("%H:%M:%S", time.gmtime(round(time.perf_counter() - self._start_time, 4)))

        self._start_time = None
        if self.logger:
            self.logger.info(f"소요시간: {elapsed_time}")
        else:
            print(f"소요시간: {elapsed_time}")
