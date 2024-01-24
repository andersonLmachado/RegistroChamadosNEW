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
        
        set_default_color_theme('dark-blue')
        
        # self.frame_titulo = Ctk.CTkFrame(self, border_width=2)
        # self.label_titulo = Ctk.CTkLabel(master=self.frame_titulo, text="Registro de Atendimentos", font=("Helvetica", 16, "bold"))
        
        # Contadores 
        self.contador_suzano1 = None
        self.contador_suzano2 = None
        self.contador_suzano3 = None
        self.contador_suzano4 = None
        self.contador_suzano5 = None
        self.contador_suzano6 = None
        self.contador_suzano7 = None
        self.contador_suzano8 = None
        
        self.contador_adami1 = None
        self.contador_adami2 = None
        self.contador_adami3 = None
        self.contador_adami4 = None
        self.contador_adami5 = None
        self.contador_adami6 = None
        self.contador_adami7 = None
        self.contador_adami8 = None

        ## Declarar empresa ##

        self.frame_suzano = Ctk.CTkFrame(self, border_width=2)
        
        self.label_suzano = Ctk.CTkLabel(master=self.frame_suzano, text=f"Suzano", font=("Helvetica", 16, "bold"))

        self.frame_adami = Ctk.CTkFrame(self, border_width=2)

        self.label_adami = Ctk.CTkLabel(master=self.frame_adami, text=f"Adami", font=("Helvetica", 16, "bold"))

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

        
        #Adami
        self.btn_adami1 = Ctk.CTkButton(master=self.frame_adami, text="Cadastro Omnilink", command=self.AdamiContador1)
        self.label_adami1 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami2 = Ctk.CTkButton(master=self.frame_adami, text="Verificar Espelhamento", command=self.AdamiContador2)
        self.label_adami2 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami3 = Ctk.CTkButton(master=self.frame_adami, text="Trocar Rastreador", command=self.AdamiContador3)
        self.label_adami3 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami4 = Ctk.CTkButton(master=self.frame_adami, text="Voltar para PNA", command=self.AdamiContador4)
        self.label_adami4 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami5 = Ctk.CTkButton(master=self.frame_adami, text="Treinamento", command=self.AdamiContador5)
        self.label_adami5 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami6 = Ctk.CTkButton(master=self.frame_adami, text="Cadastro", command=self.AdamiContador6)
        self.label_adami6 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami7 = Ctk.CTkButton(master=self.frame_adami, text="Dúvidas", command=self.AdamiContador7)
        self.label_adami7 = Ctk.CTkLabel(master=self.frame_adami, text="0")

        self.btn_adami8 = Ctk.CTkButton(master=self.frame_adami, text="Outros", command=self.AdamiContador8)
        self.label_adami8 = Ctk.CTkLabel(master=self.frame_adami, text="0")
        
        # Exportar para Excel
        self.btn_exportar_excel = Ctk.CTkButton(master=self, text="Exportar para Excel", fg_color="#52BE80", hover=True, hover_color="#229954", command=self.ExportarExcel)
        
        # Limpeza de contadores
        self.btn_limpar = Ctk.CTkButton(master=self, text="Limpar Contadores", fg_color="#E74C3C", hover=True, hover_color="#943126", command=self.confirmar_limpeza_contadores)
        
        ## Posicionamento dos botões ##
        
        # Título
        # self.frame_titulo.grid(row=0, column=6, pady=2, padx=2)
        # self.label_titulo.grid(row=0, column=6, pady=2, padx=2)

        self.label_adami.grid(row=0, column=6, columnspan=2, pady=10, padx=10)
        self.frame_adami.grid(row=4, column=6, pady=2, padx=2)

        
        # Suzano
        self.label_suzano.grid(row=0, column=0, columnspan=2, pady=10, padx=10)
        self.frame_suzano.grid(row=4, column=0, pady=5, padx=5)
        
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

        # Adami
        self.btn_adami1.grid(row=1, column=7, pady=10, padx=10)
        self.label_adami1.grid(row=1, column=8, pady=10, padx=10)

        self.btn_adami2.grid(row=2, column=7, pady=10, padx=10)
        self.label_adami2.grid(row=2, column=8, pady=10, padx=10)

        self.btn_adami3.grid(row=3, column=7, pady=10, padx=10)
        self.label_adami3.grid(row=3, column=8, pady=10, padx=10)

        self.btn_adami4.grid(row=4, column=7, pady=10, padx=10)
        self.label_adami4.grid(row=4, column=8, pady=10, padx=10)

        self.btn_adami5.grid(row=5, column=7, pady=10, padx=10)
        self.label_adami5.grid(row=5, column=8, pady=10, padx=10)

        self.btn_adami6.grid(row=6, column=7, pady=10, padx=10)
        self.label_adami6.grid(row=6, column=8, pady=10, padx=10)

        self.btn_adami7.grid(row=7, column=7, pady=10, padx=10)
        self.label_adami7.grid(row=7, column=8, pady=10, padx=10)

        self.btn_adami8.grid(row=8, column=7, pady=10, padx=10)
        self.label_adami8.grid(row=8, column=8, pady=10, padx=10)
        
        # Exportar para Excel
        self.btn_exportar_excel.grid(row=10, column=6, pady=10, padx=10)
        
        # Limpeza de contadores
        self.btn_limpar.grid(row=10, column=7, pady=10, padx=10)
        
        # Carregar contagens da última sessão (se existirem)
        self.carregar_ultima_sessao()
        
        # Configurar evento de fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.fechar_janela)
        
    # Funções
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
    
    
    #Adami
    def AdamiContador1(self):
        self.contador_adami1 += 1
        self.atualizar_rotulo(self.label_adami1, self.contador_adami1)
    
    def AdamiContador2(self):
        self.contador_adami2 += 1
        self.atualizar_rotulo(self.label_adami2, self.contador_adami2)
    
    def AdamiContador3(self):
        self.contador_adami3 += 1
        self.atualizar_rotulo(self.label_adami3, self.contador_adami3)
    
    def AdamiContador4(self):
        self.contador_adami4 += 1
        self.atualizar_rotulo(self.label_adami4, self.contador_adami4)
    
    def AdamiContador5(self):
        self.contador_adami5 += 1
        self.atualizar_rotulo(self.label_adami5, self.contador_adami5)
    
    def AdamiContador6(self):
        self.contador_adami6 += 1
        self.atualizar_rotulo(self.label_adami6, self.contador_adami6)
    
    def AdamiContador7(self):
        self.contador_adami7 += 1
        self.atualizar_rotulo(self.label_adami7, self.contador_adami7)
    
    def AdamiContador8(self):
        self.contador_adami8 += 1
        self.atualizar_rotulo(self.label_adami8, self.contador_adami8)

    def atualizar_rotulo(self, rotulo, contador):
        rotulo.configure(text=f"{contador}")
    
    def ExportarExcel(self):
        # Cria planilha
        workbook = Workbook()
        
        # Adicionar folha
        sheet = workbook.active
        
        # Cabeçalhos Empresas
        sheet['A1'] = "Empresa"
        sheet['A2'] = "Suzano"
        
        # Cabeçalhos Atividades
        sheet['B1'] = "Cadastro Omnilink"
        sheet['C1'] = "Verificar Espelhamento"
        sheet['D1'] = "Trocar Rastreador"
        sheet['E1'] = "Voltar para PNA"
        sheet['F1'] = "Treinamento"
        sheet['G1'] = "Cadastro"
        sheet['H1'] = "Dúvidas"
        sheet['I1'] = "Outros"
        
        # Contadores
        sheet['B2'] = self.contador_suzano1
        sheet['C2'] = self.contador_suzano2
        sheet['D2'] = self.contador_suzano3
        sheet['E2'] = self.contador_suzano4
        sheet['F2'] = self.contador_suzano5
        sheet['G2'] = self.contador_suzano6
        sheet['H2'] = self.contador_suzano7
        sheet['I2'] = self.contador_suzano8
        
        # Salvar planilha
        arquivo_excel = "contagem_atendimentos.xlsx"
        workbook.save(arquivo_excel)
        messagebox.showinfo("Exportar para Excel", f"Registro de Atendimentos salvo em {arquivo_excel} com sucesso!")
        
    def confirmar_limpeza_contadores(self):
        resposta = messagebox.askquestion("Limpar Contadores", "Deseja realmente limpar os contadores?")
        if resposta == 'yes':
            self.LimparContadores()
            
    def LimparContadores(self):
        self.contador_suzano1 = 0
        self.contador_suzano2 = 0
        self.contador_suzano3 = 0
        self.contador_suzano4 = 0
        self.contador_suzano5 = 0
        self.contador_suzano6 = 0
        self.contador_suzano7 = 0
        self.contador_suzano8 = 0

        self.contador_adami1 = 0
        self.contador_adami2 = 0
        self.contador_adami3 = 0
        self.contador_adami4 = 0
        self.contador_adami5 = 0
        self.contador_adami6 = 0
        self.contador_adami7 = 0
        self.contador_adami8 = 0
        
        self.atualizar_rotulo(self.label_suzano1, self.contador_suzano1)
        self.atualizar_rotulo(self.label_suzano2, self.contador_suzano2)
        self.atualizar_rotulo(self.label_suzano3, self.contador_suzano3)
        self.atualizar_rotulo(self.label_suzano4, self.contador_suzano4)
        self.atualizar_rotulo(self.label_suzano5, self.contador_suzano5)
        self.atualizar_rotulo(self.label_suzano6, self.contador_suzano6)
        self.atualizar_rotulo(self.label_suzano7, self.contador_suzano7)
        self.atualizar_rotulo(self.label_suzano8, self.contador_suzano8)

        self.atualizar_rotulo(self.label_adami1, self.contador_adami1)
        self.atualizar_rotulo(self.label_adami2, self.contador_adami2)
        self.atualizar_rotulo(self.label_adami3, self.contador_adami3)
        self.atualizar_rotulo(self.label_adami4, self.contador_adami4)
        self.atualizar_rotulo(self.label_adami5, self.contador_adami5)
        self.atualizar_rotulo(self.label_adami6, self.contador_adami6)
        self.atualizar_rotulo(self.label_adami7, self.contador_adami7)
        self.atualizar_rotulo(self.label_adami8, self.contador_adami8)
        
    def salvar_ultima_sessao(self):
        with open("ultima_sessao.txt", "w") as arquivo:
            arquivo.write(f"Suzano Cadastro_Omnilink: {self.contador_suzano1}\n")
            arquivo.write(f"Suzano Verificar_Espelhamento: {self.contador_suzano2}\n")
            arquivo.write(f"Suzano Trocar_Rastreador: {self.contador_suzano3}\n")
            arquivo.write(f"Suzano Voltar_PNA: {self.contador_suzano4}\n")
            arquivo.write(f"Suzano Treinamento: {self.contador_suzano5}\n")
            arquivo.write(f"Suzano Cadastro: {self.contador_suzano6}\n")
            arquivo.write(f"Suzano Dúvidas: {self.contador_suzano7}\n")
            arquivo.write(f"Suzano Outros: {self.contador_suzano8}\n")

            arquivo.write(f"Adami Cadastro_Omnilink: {self.contador_adami1}\n")
            arquivo.write(f"Adami Verificar_Espelhamento: {self.contador_adami2}\n")
            arquivo.write(f"Adami Trocar_Rastreador: {self.contador_adami3}\n")
            arquivo.write(f"Adami Voltar_PNA: {self.contador_adami4}\n")
            arquivo.write(f"Adami Treinamento: {self.contador_adami5}\n")
            arquivo.write(f"Adami Cadastro: {self.contador_adami6}\n")
            arquivo.write(f"Adami Dúvidas: {self.contador_adami7}\n")
            arquivo.write(f"Adami Outros: {self.contador_adami8}\n")

    def carregar_ultima_sessao(self):
        try:
            # Carregar contagens de um arquivo de texto
            with open('ultima_sessao.txt', 'r') as file:
                for linha in file:
                    valores = linha.strip().split()
                    if len(valores) == 3:
                        empresa, atividade, contador = linha.strip().split()
                        if empresa == "Suzano":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_suzano1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_suzano2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_suzano3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_suzano4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_suzano5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_suzano6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_suzano7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_suzano8 = int(contador) if contador != "None" else None
                        elif empresa == "Adami":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_adami1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_adami2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_adami3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_adami4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_adami5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_adami6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_adami7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_adami8 = int(contador) if contador != "None" else None

            # Update labels with loaded values
            self.atualizar_rotulo(self.label_suzano1, self.contador_suzano1)
            self.atualizar_rotulo(self.label_suzano2, self.contador_suzano2)
            self.atualizar_rotulo(self.label_suzano3, self.contador_suzano3)
            self.atualizar_rotulo(self.label_suzano4, self.contador_suzano4)
            self.atualizar_rotulo(self.label_suzano5, self.contador_suzano5)
            self.atualizar_rotulo(self.label_suzano6, self.contador_suzano6)
            self.atualizar_rotulo(self.label_suzano7, self.contador_suzano7)
            self.atualizar_rotulo(self.label_suzano8, self.contador_suzano8)

            self.atualizar_rotulo(self.label_adami1, self.contador_adami1)
            self.atualizar_rotulo(self.label_adami2, self.contador_adami2)
            self.atualizar_rotulo(self.label_adami3, self.contador_adami3)
            self.atualizar_rotulo(self.label_adami4, self.contador_adami4)
            self.atualizar_rotulo(self.label_adami5, self.contador_adami5)
            self.atualizar_rotulo(self.label_adami6, self.contador_adami6)
            self.atualizar_rotulo(self.label_adami7, self.contador_adami7)
            self.atualizar_rotulo(self.label_adami8, self.contador_adami8)

        except FileNotFoundError:
            # If the file doesn't exist, initialize counters to 0
            self.contador_suzano1 = None
            self.contador_suzano2 = None
            self.contador_suzano3 = None
            self.contador_suzano4 = None
            self.contador_suzano5 = None
            self.contador_suzano6 = None
            self.contador_suzano7 = None
            self.contador_suzano8 = None

            self.contador_adami1 = None
            self.contador_adami2 = None
            self.contador_adami3 = None
            self.contador_adami4 = None   
            self.contador_adami5 = None
            self.contador_adami6 = None
            self.contador_adami7 = None
            self.contador_adami8 = None
    
    def fechar_janela(self):
        self.salvar_ultima_sessao()
        self.destroy()
    
if __name__ == '__main__':
    app = RegistroAtendimentos()
    app.mainloop()