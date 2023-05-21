import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name = "Hysteresis Calculator",
    version = "0.1",
    description = "The program allows performing a mathematical analysis of hysteresis behavior detected in voltage gating of large beta-barrel transmembrane ion channels",
    executables = [Executable("frontend.py", base=base)]
)
