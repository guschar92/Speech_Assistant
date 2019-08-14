import sys
import os

def resource_path(relative_path):
    """ Get absolute path to resource; works for dev and for PyInstaller """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))  # PyInstaller creates a temp folder and stores path in _MEIPASS
    res_path = os.path.join(base_path, "resources", relative_path)
    if os.path.exists(res_path):
        return res_path
    return os.path.join(base_path, relative_path)
