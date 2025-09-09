import sys
sys.path.append(r"C:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)")
from typing import Any
from src.logs.logger import logging

class CustomException(Exception):
    def __init__(self, error_message: Any, error_details: Any = sys):
        # initialize base Exception with message
        super().__init__(str(error_message))
        self.error_message = error_message

        # get traceback info safely
        exc_info = error_details.exc_info() if hasattr(error_details, "exc_info") else (None, None, None)
        _, _, exc_tb = exc_info

        if exc_tb is not None:
            self.lineno = exc_tb.tb_lineno
            self.file_name = exc_tb.tb_frame.f_code.co_filename
        else:
            self.lineno = None
            self.file_name = None

    def __str__(self):
        return (
            "Error occurred in python script name [{0}] line number [{1}] "
            "error message [{2}]".format(self.file_name, self.lineno, str(self.error_message))
        )

if __name__ == '__main__':
    try:
        logging.info("Enter the try block")
        a = 1 / 0
        print("This will not be printed", a)
    except Exception as e:
        # log full exception (stack trace) before raising custom exception
        logging.exception("An exception occurred")
        raise CustomException(e, sys)
