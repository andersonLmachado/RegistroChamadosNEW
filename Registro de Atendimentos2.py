import customtkinter as Ctk 
from customtkinter import *
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook

class RegistroAtendimentos(Ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Registro de Atendimentos')
        self.geometry('800x600')
        
        set_default_color_theme('blue')
        
        self.frame_titulo = Ctk.CTkFrame(self, border_width=2)
        
        self.label_titulo = Ctk.CTkLabel(master=self.frame_titulo, text="Registro de Atendimentos", font=("Helvetica", 16, "bold"))
        
        # Contadores 
        self.contador_suzano1 = 0
        self.contador_suzano2 = 0
        self.contador_suzano3 = 0
        self.contador_suzano4 = 0
        self.contador_suzano5 = 0
        self.contador_suzano6 = 0
        self.contador_suzano7 = 0
        self.contador_suzano8 = 0
        
        ## Declarar empresa ##

        self.frame_suzano = Ctk.CTkFrame(self, border_width=2)
        
        self.label_suzano = Ctk.CTkLabel(master=self.frame_suzano, text=f"Suzano", font=("Helvetica", 16, "bold"))

        ## Declarar botões ##
        self.btn_suzano1 = Ctk.CTkButton(master=self.frame_suzano, text="Cadastro Omnilink", command=self.SuzanoContador1)
        self.label_suzano1 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano2 = Ctk.CTkButton(master=self.frame_suzano, text="Verificar Espelhamento", command=self.SuzanoContador2)
        self.label_suzano2 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano3 = Ctk.CTkButton(master=self.frame_suzano, text="Trocar Rastreador", command=self.SuzanoContador3)
        self.label_suzano3 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano4 = Ctk.CTkButton(master=self.frame_suzano, text="Voltar para PNA", command=self.SuzanoContador4)
        self.label_suzano4 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano5 = Ctk.CTkButton(master=self.frame_suzano, text="Treinamento", command=self.SuzanoContador5)
        self.label_suzano5 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano6 = Ctk.CTkButton(master=self.frame_suzano, text="Cadastro", command=self.SuzanoContador6)
        self.label_suzano6 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano7 = Ctk.CTkButton(master=self.frame_suzano, text="Dúvidas", command=self.SuzanoContador7)
        self.label_suzano7 = Ctk.CTkLabel(master=self.frame_suzano, text="0")

        self.btn_suzano8 = Ctk.CTkButton(master=self.frame_suzano, text="Outros", command=self.SuzanoContador8)
        self.label_suzano8 = Ctk.CTkLabel(master=self.frame_suzano, text="0")
        
        #Exportar para Excel
        self.btn_exportar_excel = Ctk.CTkButton(master=self, text="Exportar para Excel", fg_color="#52BE80",hover=True, hover_color="#229954", command=self.ExportarExcel)
        
        # #Limpeza de contadores
        self.btn_limpar = Ctk.CTkButton(master=self, text="Limpar Contadores", fg_color="#E74C3C",hover=True, hover_color="#943126", command=self.confirmar_limpeza_contadores)
        
         ## Posicionamento dos botões ##
        
        #Título
        self.frame_titulo.grid(row=0, column=6, pady=10, padx=10)
        self.label_titulo.grid(row=0, column=6, pady=10, padx=10)
        
        #Suzano
        self.label_suzano.grid(row=0, column=0, columnspan=2, pady=10, padx=10)
        self.frame_suzano.grid(row=4, column=0, pady=10, padx=10)
        
        self.btn_suzano1.grid(row=1, column=1, pady=10, padx=10)
        self.label_suzano1.grid(row=1, column=2, pady=10, padx=10)
        
        self.btn_suzano2.grid(row=2, column=1, pady=10, padx=10)
        self.label_suzano2.grid(row=2, column=2, pady=10, padx=10)
        
        self.btn_suzano3.grid(row=3, column=1, pady=10, padx=10)
        self.label_suzano3.grid(row=3, column=2, pady=10, padx=10)
        
        self.btn_suzano4.grid(row=4, column=1, pady=10, padx=10)
        self.label_suzano4.grid(row=4, column=2, pady=10, padx=10)
        
        self.btn_suzano5.grid(row=5, column=1, pady=10, padx=10)
        self.label_suzano5.grid(row=5, column=2, pady=10, padx=10)
        
        self.btn_suzano6.grid(row=6, column=1, pady=10, padx=10)
        self.label_suzano6.grid(row=6, column=2, pady=10, padx=10)
        
        self.btn_suzano7.grid(row=7, column=1, pady=10, padx=10)
        self.label_suzano7.grid(row=7, column=2, pady=10, padx=10)
        
        self.btn_suzano8.grid(row=8, column=1, pady=10, padx=10)
        self.label_suzano8.grid(row=8, column=2, pady=10, padx=10)
        
        #Exportar para Excel
        self.btn_exportar_excel.grid(row=10, column=6, pady=10, padx=10)
        
        #Limpeza de contadores
        self.btn_limpar.grid(row=10, column=7, pady=10, padx=10)
        
        #Funções
    def SuzanoContador1(self):
        self.contador_suzano1 += 1
        self.atualizar_rotulo(self.label_suzano1, self.contador_suzano1)
    
    def SuzanoContador2(self):
        self.contador_suzano2 += 1
        self.atualizar_rotulo(self.label_suzano2, self.contador_suzano2)
    
    def SuzanoContador3(self):
        self.contador_suzano3 += 1
        self.atualizar_rotulo(self.label_suzano3, self.contador_suzano3)
    
    def SuzanoContador4(self):
        self.contador_suzano4 += 1
        self.atualizar_rotulo(self.label_suzano4, self.contador_suzano4)
    
    def SuzanoContador5(self):
        self.contador_suzano5 += 1
        self.atualizar_rotulo(self.label_suzano5, self.contador_suzano5)
    
    def SuzanoContador6(self):
        self.contador_suzano6 += 1
        self.atualizar_rotulo(self.label_suzano6, self.contador_suzano6)
    
    def SuzanoContador7(self):
        self.contador_suzano7 += 1
        self.atualizar_rotulo(self.label_suzano7, self.contador_suzano7)
    
    def SuzanoContador8(self):
        self.contador_suzano8 += 1
        self.atualizar_rotulo(self.label_suzano8, self.contador_suzano8)
    
    def ExportarExcel(self):
        #Cria planilha
        workbook = Workbook()
        
        #Adicionar folha
        sheet = workbook.active
        
        #Cabeçalhos Empresas
        sheet['A1'] = "Empresa"
        sheet['A2'] = "Suzano"
        
        #Cabeçalhos Atividades
        sheet['B1'] = "Cadastro Omnilink"
        sheet['C1'] = "Verificar Espelhamento"
        sheet['D1'] = "Trocar Rastreador"
        sheet['E1'] = "Voltar para PNA"
        sheet['F1'] = "Treinamento"
        sheet['G1'] = "Cadastro"
        sheet['H1'] = "Dúvidas"
        sheet['I1'] = "Outros"
        
        #Contadores
        sheet['B2'] = self.contador_suzano1
        sheet['C2'] = self.contador_suzano2
        sheet['D2'] = self.contador_suzano3
        sheet['E2'] = self.contador_suzano4
        sheet['F2'] = self.contador_suzano5
        sheet['G2'] = self.contador_suzano6
        sheet['H2'] = self.contador_suzano7
        sheet['I2'] = self.contador_suzano8
        
        #Salvar planilha
        arquivo_excel = "contagem_atendimentos.xlsx"
        workbook.save(arquivo_excel)
        messagebox.showinfo("Exportar para Excel", f"Registro de Atendimentos salvo em {arquivo_excel} com sucesso!")
        
    def confirmar_limpeza_contadores(self):
        resposta = messagebox.askquestion("Limpar Contadores", "Deseja realmente limpar os contadores?")
        if resposta == 'yes':
            self.LimparContadores()
            
    def atualizar_rotulo(self, rotulo, contador):
        rotulo.configure(text=f"{contador}")
    
    def LimparContadores(self):
        self.contador_suzano1 = 0
        self.contador_suzano2 = 0
        self.contador_suzano3 = 0
        self.contador_suzano4 = 0
        self.contador_suzano5 = 0
        self.contador_suzano6 = 0
        self.contador_suzano7 = 0
        self.contador_suzano8 = 0
        
        self.atualizar_rotulo(self.label_suzano1, self.contador_suzano1)
        
        self.atualizar_rotulo(self.label_suzano2, self.contador_suzano2)
        
        self.atualizar_rotulo(self.label_suzano3, self.contador_suzano3)
        
        self.atualizar_rotulo(self.label_suzano4, self.contador_suzano4)
        
        self.atualizar_rotulo(self.label_suzano5, self.contador_suzano5)
        
        self.atualizar_rotulo(self.label_suzano6, self.contador_suzano6)
        
        self.atualizar_rotulo(self.label_suzano7, self.contador_suzano7)
        
        self.atualizar_rotulo(self.label_suzano8, self.contador_suzano8)
                  
if __name__ == '__main__':
    app = RegistroAtendimentos()
    app.mainloop()

        
        
        
