import cx_Freeze

executables = [cx_Freeze.Executable("main.py")]

cx_Freeze.setup(
    name="Email_app",
    version="0.1",
    description="App for managing email data",
    author="Nicolas A. Schcolnicov",
    options={"build_exe": {"packages":["tkinter","pandas","datetime","os","openpyxl"], "icon": "C://Users//ns39399//email_tracker//offline//app_dev//email-action-at.ico", "base": "Win32GUI"}},
    executables = executables
)
