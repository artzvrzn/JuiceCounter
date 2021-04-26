import logging.config

logger = logging.getLogger('counter')
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)

formatter = logging.Formatter(fmt='%(asctime)s %(message)s', datefmt='%H:%M:%S')

console_handler.setFormatter(formatter)
logger.addHandler(console_handler)
