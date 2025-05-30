import logging
import os

def setup_logger(name, log_file, level=logging.INFO):
    """Function to setup as many loggers as you want"""
    
    # Ensure log directory exists
    log_dir = os.path.dirname(log_file)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir)

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Handler for file
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)

    # Handler for console
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Avoid adding handlers multiple times
    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)

    return logger

if __name__ == '__main__':
    # Example usage:
    # Ensure the logs directory exists if you run this directly for testing
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    logger1 = setup_logger('my_app_logger', 'logs/app.log')
    logger1.info('This is an info message from logger1.')
    logger1.error('This is an error message from logger1.')

    logger2 = setup_logger('my_module_logger', 'logs/module.log', level=logging.DEBUG)
    logger2.debug('This is a debug message from logger2.')
    logger2.info('This is an info message from logger2.')
    
    # Test that handlers are not added again
    logger1_again = setup_logger('my_app_logger', 'logs/app.log')
    logger1_again.info('Another message from logger1, should not duplicate handlers.')
