# student_card_tool/utils/path_utils.py
import os
import sys

def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的路径
        base_path = os.path.dirname(sys.executable)
    else:
        # 从 main.py 所在目录开始，始终指向 student_card_tool/
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)