from customLogging import setup_logging
import logging

setup_logging()
logger = logging.getLogger(__name__)

logger.error("ola")
logger.info("ASD")