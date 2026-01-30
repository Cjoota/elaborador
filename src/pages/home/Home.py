import flet as ft
import traceback
from time import sleep
from src.pages.home.functions import Functions

class Home:
    def __init__(self,page: ft.Page):
        self.page = page
        self.functions = Functions(self.page)
        self.buildWidgetsToView()
        self.buildRegisterMenu()
        
    def buildWidgetsToView(self):
        self.companies_registred = ft.Dropdown(width=350,label="Empresa",menu_width=350,enable_filter=True,enable_search=True,editable=True)
        self.companies_list = ft.Dropdown(width=350,border_radius=10,label="Empresa",enable_filter=True,enable_search=True,editable=True)
        self.colaborator_name = ft.TextField(label="Nome do Colaborador",width=300, border_radius=10)
        self.colaborator_cpf = ft.TextField(
            label="CPF do Colaborador",
            width=350, 
            border_radius=10,
            on_change=self.functions.limpar_input_cpf,      
            keyboard_type=ft.KeyboardType.NUMBER 
        )
        
        self.start_button = ft.FloatingActionButton("Criar PDF",icon=ft.Icons.CREATE, 
                                                    bgcolor=ft.Colors.GREEN_ACCENT_200,
                                                    on_click=lambda: self.openCreateMenu() )
        self.registerKey_button = ft.FloatingActionButton("Registrar chave",
                                                          width=150,
                                                          icon=ft.Icons.EDIT,
                                                          bgcolor=ft.Colors.GREEN_ACCENT_200,
                                                          on_click=lambda: self.openRegisterMenu())
    
    def build_view(self):
        self.functions.load_companies(self.companies_list)
        self.functions.load_companies_registred(self.companies_registred)
        return ft.Column(
            [
                ft.Row(
                    [
                        self.companies_registred,self.registerKey_button
                    ],alignment=ft.MainAxisAlignment.CENTER,vertical_alignment=ft.CrossAxisAlignment.CENTER
                ),
                ft.Row(
                    [
                        self.colaborator_name,self.colaborator_cpf,self.start_button
                    ],alignment=ft.MainAxisAlignment.CENTER,vertical_alignment=ft.CrossAxisAlignment.CENTER
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            
            horizontal_alignment=ft.CrossAxisAlignment.CENTER 
        )

    def buildRegisterMenu(self):
        self.registerKey_key = ft.TextField(label="Insira uma senha",width=250, border_radius=10)
        self.registerKey_contentMenu = ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Selecione a Empresa e digite a chave para registrar:")
                    ]
                ),
                ft.Row(
                    [
                        self.companies_list,
                    ]
                ),
                ft.Row(
                    [
                        self.registerKey_key
                    ]
                ),
            ],height=200
        )
        self.registerKey_Menu = ft.AlertDialog(
            modal=True,title=ft.Text("Registrar chave de criptografia"),
            content=self.registerKey_contentMenu,
            actions=[ft.Button("Registrar",on_click=lambda: self.register()),ft.Button("Sair",on_click=lambda: self.page.pop_dialog() ),])
        
    def openRegisterMenu(self):
        try:
            self.buildRegisterMenu()
            Functions(self.page).open_menu(self.registerKey_Menu)
        except Exception as e:
            error_msg = f"ERRO: {str(e)}\n\nDETALHES:\n{traceback.format_exc()}"
            
            self.page.snack_bar = ft.SnackBar(ft.Text(f"ERRO: {e}"), bgcolor=ft.Colors.RED)
            self.page.snack_bar.open = True
            self.page.update()
            
            with open("ERRO_DO_SISTEMA.txt", "w", encoding="utf-8") as f:
                f.write(error_msg)

    def buildCreatePdf(self):
        self.exams_list = []
        self.ListForPDF = []
        self.exams_text = ft.Text(f"")
        self.count = 0
        self.loading_ = ft.Container(expand=True,alignment=ft.Alignment.CENTER)
        self.digitalizeButton = ft.FloatingActionButton("Digitalizar Exame",
                                                          width=150,
                                                          icon=ft.Icons.SCANNER,
                                                          bgcolor=ft.Colors.GREEN_ACCENT_200,
                                                          on_click=lambda: self.digitalize())
        self.content_menu = ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Insira o Documento no scanner e clique em Digitalizar")
                    ]
                ),
                ft.Row([ ft.Text(f"Exame de: "),ft.Text(f"{self.colaborator_name.value}",size=15, weight=ft.FontWeight.BOLD)]),
                ft.Row(
                    [
                        self.exams_text
                    ]
                ),
                ft.Row(
                    [
                        self.digitalizeButton
                    ]
                ),
            ],height=600,width=1000
        )
        self.create_Menu = ft.AlertDialog(
            modal=True,title=ft.Text(f"Criar exame para {self.companies_registred.value}"),
            content=self.content_menu,
            actions=[ft.Button("Lacrar e Exportar",on_click=lambda: self.createPdf()),ft.Button("Encerrar",on_click=lambda: self.page.pop_dialog())]
        )
        self.content_menu.controls.append(self.loading_)
    def openCreateMenu(self):
        if self.companies_registred.value != None:
            self.buildCreatePdf()
            self.page.show_dialog(self.create_Menu)
            return
        self.functions.snack_bar("Selecione a empresa do colaborador, se não existir nenhuma crie uma chave para ela",ft.Colors.RED)
        
    def register(self):
        try:
            if self.companies_list.value != None and self.registerKey_key.value != "":
                result = self.functions.singUp_companies(self.companies_list.value,self.registerKey_key.value)
                if result is True:
                    self.companies_registred.options.clear()
                    self.functions.load_companies_registred(self.companies_registred)
                    self.page.pop_dialog()
                    self.companies_list.value = "Empresa"
                    self.registerKey_key.value = ""
                    self.functions.snack_bar("Cadastrado com Sucesso",ft.Colors.GREEN)
                    return
                self.page.pop_dialog()
                self.functions.snack_bar(f"{result}",ft.Colors.RED)
            self.page.pop_dialog()
            self.functions.snack_bar("Preencha todos os campos!",ft.Colors.RED)
        except Exception as e:
            self.functions.snack_bar(f"ERRO: {str(e)}", ft.Colors.RED)
            print(e)

    def digitalize(self):
       self.exame_name = self.functions.digitalize(self.colaborator_name.value,f"{self.count}.jpg")
       self.count += 1
       self.exams_list.append(self.exame_name.name)
       self.exams_text.value = f"{self.exams_list}".replace("[","").replace("]","")
       self.ListForPDF.append(self.exame_name)
       self.loading_.content = ft.Text(f"Digitalizado documento: {self.exame_name.name}",size=20,weight=ft.FontWeight.BOLD)
       self.page.update(self.loading_)

    def createPdf(self):
        sleep(1.5)
        msg = ""
        
        try:
            msg = self.functions.createPDF(self.companies_registred.value,self.ListForPDF,self.colaborator_name.value,self.colaborator_cpf.value)
        except Exception as e:
            self.functions.snack_bar(f"Não foi possivel lacrar o pdf, ERRO: {str(e)} and {msg}",ft.Colors.GREEN)
            print(e)
        self.loading_.content = ft.Text("PDF LACRADO!",size=20,weight=ft.FontWeight.BOLD)
        self.page.update(self.loading_)