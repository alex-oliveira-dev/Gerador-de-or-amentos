import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["tkinter", "docx", "docx2pdf","tkinter","threading","docx","docx2pdf","subprocess", "platform"],
    "include_files": [("templates/orcamento.docx")]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Gerador de orçamentos",
    version="1.5",
    description="Gera orçamentos de manutenções!",
    options={"build_exe": build_exe_options},
    executables=[Executable("gerador_de_orcamentos.py", base=base)]
)
