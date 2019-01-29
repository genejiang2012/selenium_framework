#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File: ClipBoard.py
Author: Gene Jiang
Email: zhengrong.jiang@chiefclouds.com
Github: https://github.com/yourname
Description: use the clipbaord function
"""

import win32clipboard as w
import win32con


class ClipBoard(object):
    """Simulate the clipboard function"""
    
    @staticmethod
    def get_text():
        """
        Get the text from clipboard
        :returns: the clipboard content
        """
        w.OpenClipBoard()
        d = w.GetClipboardData(win32con.CF_TEXT)
        w.CloseClipboard()

        return d

    @staticmethod
    def set_text(one_string):
        """Send the text to clipboard

        :one_string: send the string to the clipboard
        :returns: None 

        """
        w.OpenClipBoard()
        w.EmptyClipBoard()
        w.SetClipboardData(win32con.UNICODETEXT, one_string)
        w.CloseClipBoard()


if __name__ == "__main__":
    pass

