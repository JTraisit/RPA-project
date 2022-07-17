import logging 

#log error
def logsetup():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    format_log = logging.Formatter('%(asctime)s - %(levelname)s - %(module)s - line call :%(lineno)d - %(message)s')
    file_handler = logging.FileHandler(filename = r"C:\Users\admin\Desktop\Internship\Project python version\app\logs\mainprogress.log")
    file_handler.setFormatter(format_log)
    logger.addHandler(file_handler)
    return logger