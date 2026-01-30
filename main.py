import flet as ft


from src.db.db_controler import Database
from src.db.paswords import Passwords
from src.core.elaborador import Interfaces

import atexit
import shutil
from pathlib import Path

class Main:
    def __init__(self,page:ft.Page):
        self.page = page
        self.db = Database()
        self.encrypt = Passwords()
        self.Interfaces = Interfaces(self.page)
        self.Interfaces.build_interface()
        
@atexit.register
def Limpar():
	for d in Path(".").rglob("__pycache__"):
		shutil.rmtree(d)

ft.app(target=Main)