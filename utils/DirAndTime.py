#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: DirAndTime.py
Author: Gene Jiang
Email: zhengrong.jiang@chiefclouds.com
Description: Get the current date, time 
"""

import time
import os
from datetime import datetime
from config.var_config import screenPicturesDir


def get_current_date():
    """
    get the current date
    :return: current date
    """
    time_tup = time.localtime()
    current_date = str(time_tup.tm_year) + "-" + \
        str(time_tup.tm_mon) + "-" + str(time_tup.tm_mday)

    return current_date


def get_current_time():
    """
    get the current time
    :return: current time, hour-minutes-seconds
    """
    time_str = datetime.now()
    now_time = time_str.strftime('%H:%M:%S')

    return now_time


def create_current_date_dir():
    """
    create the directory based on the current date
    :return: current directory
    """

    dir_name = os.path.join(screenPicturesDir, get_current_date())
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)

    return dir_name


if __name__ == "__main__":
    print(get_current_date())
    print(get_current_time())
    print(create_current_date_dir())