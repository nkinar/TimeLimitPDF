import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": [], "excludes": [], "optimize": 2, "include_msvcr": True}
base = None
setup(
    name="TimeLimitPDF",
    version="0.1",
    description="A simple utility for closing a PDF and hiding the document when time expires.",
    options={"build_exe": build_exe_options},
    executables=[Executable("TimeLimitPDF.py", base=base)],
)