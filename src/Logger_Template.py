import logging 
from functools import wraps

def main():

    logger = MyLogger
    logger.logger.info("This is a test.")

class MyLogger:

    _instance = None # Class-level attribute to hold the singleton instance

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls, *args, **kwargs)
            cls._instance._initialize_logger()
        return cls._instance

    def _initialize_logger(self):
        """
        Configures and returns a logger instance with handlers for both terminal
        and file output.
        """
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)

        # File handler for storing logs
        file_handler = logging.FileHandler( r'C:\Users\urban\Documents\GitHub\CTS_Creation\logs\test.log', mode="w") # TODO Replace the str here with your log file path
        file_formatter = logging.Formatter(
            "%(levelname)s:%(name)s:%(asctime)s:%(funcName)s(): line %(lineno)d: %(message)s"
        )
        file_handler.setFormatter(file_formatter)

        # Stream handler for terminal output
        stream_handler = logging.StreamHandler()
        stream_formatter = logging.Formatter("%(levelname)s: %(message)s")
        stream_handler.setFormatter(stream_formatter)

        # Add both handlers to the logger
        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)

        self.logger = logger
    
    def get_logger(self):
        return self.logger
    
def get_default_logger():
    return MyLogger().get_logger()

def log_function_call(func):
    """
    Decorator to log the entry, exit, and execution details of a function
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        logger = get_default_logger()
        logger.info(f"Running {func.__name__} with args: {args}, kwargs: {kwargs}")
        try:
            result = func(*args, **kwargs)
            logger.info(f"Finished {func.__name__} and returning {result}")
        except Exception as e:
            logger.error(f"Error occurred in {func.__name__}: {e}")
            raise
        else:
            return result
        
    return wrapper

if __name__ == "__main__":
    
    main()