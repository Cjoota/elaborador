from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os

# Importa o cérebro do Excel que acabamos de arrumar
from relations.functions import ExcelProtocol

class Locker:
    def __init__(self):
        self.exames_path = "exames"

    def create_exam(self, enterprise: str, colaborator: str, exams: list, passw: str, cpf: str = ""):
        print(f"[LOCKER] Iniciando processo para: {colaborator}")
        try:
            # --- 1. GERAÇÃO DO PDF (Normal) ---
            os.makedirs(fr"{self.exames_path}\{enterprise}", exist_ok=True)
            # os.makedirs(fr"temp", exist_ok=True) 
            
            caminho_pdf = fr"{self.exames_path}\{enterprise}\{colaborator}.pdf"
            c = canvas.Canvas(caminho_pdf, pagesize=A4, encrypt=passw)
            
            for i, imagem in enumerate(exams):
                if os.path.exists(imagem):
                    c.drawImage(imagem, 0, 0, width=A4[0], height=A4[1], preserveAspectRatio=True)
                    if i < len(exams) - 1:
                        c.showPage()
            c.save()
            print("[LOCKER] PDF Gerado com sucesso.")
            # ----------------------------------

            # --- 2. GATILHO DO EXCEL (INTELIGENTE) ---
            # Se o CPF foi preenchido, o sistema entende que precisa gerar protocolo
            if cpf and cpf.strip() != "":
                print(f"[LOCKER] CPF detectado ({cpf}). Acionando ExcelProtocol...")
                
                # Instancia a classe e manda registrar
                protocolo = ExcelProtocol()
                resultado = protocolo.registrar_colaborador(colaborator, cpf)
                
                if not resultado:
                     print("[AVISO] PDF ok, mas falha ao gerar Excel.")
            else:
                print("[LOCKER] CPF vazio. Gerando apenas PDF.")
            
            return "Sucesso"

        except Exception as e:
            print(f"[ERRO LOCKER] {e}")
            return f"Erro: {e}"