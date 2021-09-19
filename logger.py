import logging
from pythonjsonlogger import jsonlogger

logger = logging.getLogger()
logHandler = logging.StreamHandler()
formatter = jsonlogger.JsonFormatter("%(levelname)s %(message)s %(funcName)s %(module)s")
logHandler.setFormatter(formatter)
logger.setLevel(logging.INFO)
logger.addHandler(logHandler)
