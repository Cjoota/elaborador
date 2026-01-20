import flet as ft
import os
from pathlib import Path

from src.db.db_controler import Database
from src.db.paswords import Passwords
from src.functions.digitalize import Digitalize
from src.functions.locker import Locker


import logging
logger = logging.Logger("Encrypting")

class Functions:
    def __init__(self,page: ft.Page):
        self.page = page
        self.db = Database()


    def load_companies_registred(self,dropdown:ft.Dropdown):
        companies = self.db.get_all_companies()
        for enterprise in companies:
            if Passwords().get_companiesKey(enterprise.razao) != None:
                dropdown.options.append(
                    ft.DropdownOption(
                        text=f"{enterprise.razao}"
                    )
                )

    def load_companies(self,dropdown:ft.Dropdown):
        companies = self.db.get_all_companies()
        for enterprise in companies:
            dropdown.options.append(
                ft.DropdownOption(
                    text=f"{enterprise.razao}"
                )
            )
    
    def open_menu(self,dialog):
        self.page.show_dialog(dialog)

    def singUp_companies(self,companies_choices,key):
        try:
            Passwords().set_companies(companies_choices,key)
            return True
        except Exception as e:
            logger.info(f"Erro ao cadastrar chave: {str(e)}")
            return f"Erro ao cadastrar chave: {str(e)}"
        
    def snack_bar(self,message,color):
        snack_bar = ft.SnackBar(content=message,bgcolor=color)
        self.page.show_dialog(snack_bar)

    def digitalize(self,colaborator:str,description:str):
        os.makedirs(fr"{Digitalize().scan_path}\{colaborator}",exist_ok=True)
        archive_name = Digitalize().digitalizar(nome_arquivo=description,sub_path=colaborator)
        if not archive_name:
            return False
        archive_digitalize = Path(archive_name)
        return archive_digitalize
    
    def limpar_input_cpf(self, e):
        valor_atual = e.control.value
        
        novo_valor = ''.join(c for c in valor_atual if c.isdigit() or c in ".-")
        
        if valor_atual != novo_valor:
            e.control.value = novo_valor
            e.control.update()
                
    def createPDF(self, enterprise: str, examsList_path: list, colaborator: str, cpf: str = ""):
        Msg = Locker().create_exam(
            enterprise=enterprise, 
            colaborator=colaborator, 
            exams=examsList_path, 
            passw=Passwords().get_companiesKey(enterprise),
            cpf=cpf
        )
        return Msg
    