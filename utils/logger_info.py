#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: logger_info.py
Author: Me
Email: yourname@email.com
Github: https://github.com/yourname
Description: 
"""
import logging
import logging.config
from config.var_config import project_dir

# read the log configuration file
logging.config.fileConfig(project_dir + '\config\Logger.conf')
# select the log formatter
logging.getLogger("example01")


class MyLogger(object):
    def __init__(self):
        pass

    def debug(self, message):
        return self.logger.debug(message)

    def info(self, message):
        return self.logger.info(message)

    def warning(self, message):
        return self.logger.warning(message)


if __name__ == '__main__':
    Logger = MyLogger()
    Logger.debug("This is one debug message!")
    Logger.info("This is one info message!")
    Logger.warning("This is one warning message!")


