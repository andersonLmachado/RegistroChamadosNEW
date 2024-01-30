import os
import customtkinter as Ctk 
import json
from customtkinter import *
from tkinter import messagebox
from openpyxl import Workbook
from datetime import datetime

class RegistroAtendimentos(Ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Registro de Atendimentos')
        self.geometry('900x900')
        
        set_default_color_theme('blue')
        
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
        
        self.contador_klabin1 = None
        self.contador_klabin2 = None
        self.contador_klabin3 = None
        self.contador_klabin4 = None
        self.contador_klabin5 = None
        self.contador_klabin6 = None
        self.contador_klabin7 = None
        self.contador_klabin8 = None

        self.contador_irani1 = None
        self.contador_irani2 = None
        self.contador_irani3 = None
        self.contador_irani4 = None
        self.contador_irani5 = None
        self.contador_irani6 = None
        self.contador_irani7 = None
        self.contador_irani8 = None

        self.contador_rfr1 = None
        self.contador_rfr2 = None
        self.contador_rfr3 = None
        self.contador_rfr4 = None
        self.contador_rfr5 = None
        self.contador_rfr6 = None
        self.contador_rfr7 = None
        self.contador_rfr8 = None

        self.contador_ind1 = None
        self.contador_ind2 = None
        self.contador_ind3 = None
        self.contador_ind4 = None
        self.contador_ind5 = None
        self.contador_ind6 = None
        self.contador_ind7 = None
        self.contador_ind8 = None

        self.contador_gkn1 = None
        self.contador_gkn2 = None
        self.contador_gkn3 = None
        self.contador_gkn4 = None
        self.contador_gkn5 = None
        self.contador_gkn6 = None
        self.contador_gkn7 = None
        self.contador_gkn8 = None

        self.contador_coopercarga1 = None
        self.contador_coopercarga2 = None
        self.contador_coopercarga3 = None
        self.contador_coopercarga4 = None
        self.contador_coopercarga5 = None
        self.contador_coopercarga6 = None
        self.contador_coopercarga7 = None
        self.contador_coopercarga8 = None

        ## Declarar empresa ##

        self.frame_suzano = Ctk.CTkFrame(self, border_width=2)
        self.label_suzano = Ctk.CTkLabel(master=self.frame_suzano, text=f"Suzano", font=("Helvetica", 16, "bold"))

        self.frame_adami = Ctk.CTkFrame(self, border_width=2)
        self.label_adami = Ctk.CTkLabel(master=self.frame_adami, text=f"Adami", font=("Helvetica", 16, "bold"))
        
        self.frame_klabin = Ctk.CTkFrame(self, border_width=2)
        self.label_klabin = Ctk.CTkLabel(master=self.frame_klabin, text=f"Klabin", font=("Helvetica", 16, "bold"))
        
        self.frame_irani = Ctk.CTkFrame(self, border_width=2)
        self.label_irani = Ctk.CTkLabel(master=self.frame_irani, text=f"Irani", font=("Helvetica", 16, "bold"))

        self.frame_rfr = Ctk.CTkFrame(self, border_width=2)
        self.label_rfr = Ctk.CTkLabel(master=self.frame_rfr, text=f"RFR - Guarulhos", font=("Helvetica", 16, "bold"))

        self.frame_ind = Ctk.CTkFrame(self, border_width=2)
        self.label_ind = Ctk.CTkLabel(master=self.frame_ind, text=f"RFR - Indaiatuba", font=("Helvetica", 16, "bold"))

        self.frame_gkn = Ctk.CTkFrame(self, border_width=2)
        self.label_gkn = Ctk.CTkLabel(master=self.frame_gkn, text=f"GKN", font=("Helvetica", 16, "bold"))

        self.frame_coopercarga = Ctk.CTkFrame(self, border_width=2)
        self.label_coopercarga = Ctk.CTkLabel(master=self.frame_coopercarga, text=f"Coopercarga", font=("Helvetica", 16, "bold"))

        self.frame_acoes = Ctk.CTkFrame(self, border_width=2)


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
        
        #Klabin
        self.btn_klabin1 = Ctk.CTkButton(master=self.frame_klabin, text="Cadastro Omnilink", command=self.KlabinContador1)
        self.label_klabin1 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin2 = Ctk.CTkButton(master=self.frame_klabin, text="Verificar Espelhamento", command=self.KlabinContador2)
        self.label_klabin2 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin3 = Ctk.CTkButton(master=self.frame_klabin, text="Trocar Rastreador", command=self.KlabinContador3)
        self.label_klabin3 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin4 = Ctk.CTkButton(master=self.frame_klabin, text="Voltar para PNA", command=self.KlabinContador4)
        self.label_klabin4 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin5 = Ctk.CTkButton(master=self.frame_klabin, text="Treinamento", command=self.KlabinContador5)
        self.label_klabin5 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin6 = Ctk.CTkButton(master=self.frame_klabin, text="Cadastro", command=self.KlabinContador6)
        self.label_klabin6 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin7 = Ctk.CTkButton(master=self.frame_klabin, text="Dúvidas", command=self.KlabinContador7)
        self.label_klabin7 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        self.btn_klabin8 = Ctk.CTkButton(master=self.frame_klabin, text="Outros", command=self.KlabinContador8)
        self.label_klabin8 = Ctk.CTkLabel(master=self.frame_klabin, text="0")
        
        #Irani
        self.btn_irani1 = Ctk.CTkButton(master=self.frame_irani, text="Cadastro Omnilink", command=self.IraniContador1)
        self.label_irani1 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani2 = Ctk.CTkButton(master=self.frame_irani, text="Verificar Espelhamento", command=self.IraniContador2)
        self.label_irani2 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani3 = Ctk.CTkButton(master=self.frame_irani, text="Trocar Rastreador", command=self.IraniContador3)
        self.label_irani3 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani4 = Ctk.CTkButton(master=self.frame_irani, text="Voltar para PNA", command=self.IraniContador4)
        self.label_irani4 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani5 = Ctk.CTkButton(master=self.frame_irani, text="Treinamento", command=self.IraniContador5)
        self.label_irani5 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani6 = Ctk.CTkButton(master=self.frame_irani, text="Cadastro", command=self.IraniContador6)
        self.label_irani6 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani7 = Ctk.CTkButton(master=self.frame_irani, text="Dúvidas", command=self.IraniContador7)
        self.label_irani7 = Ctk.CTkLabel(master=self.frame_irani, text="0")
        
        self.btn_irani8 = Ctk.CTkButton(master=self.frame_irani, text="Outros", command=self.IraniContador8)
        self.label_irani8 = Ctk.CTkLabel(master=self.frame_irani, text="0")

        #RFR

        self.btn_rfr1 = Ctk.CTkButton(master=self.frame_rfr, text="Cadastro Omnilink", command=self.RFRContador1)
        self.label_rfr1 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr2 = Ctk.CTkButton(master=self.frame_rfr, text="Verificar Espelhamento", command=self.RFRContador2)
        self.label_rfr2 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr3 = Ctk.CTkButton(master=self.frame_rfr, text="Trocar Rastreador", command=self.RFRContador3)
        self.label_rfr3 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr4 = Ctk.CTkButton(master=self.frame_rfr, text="Voltar para PNA", command=self.RFRContador4)
        self.label_rfr4 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr5 = Ctk.CTkButton(master=self.frame_rfr, text="Treinamento", command=self.RFRContador5)
        self.label_rfr5 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr6 = Ctk.CTkButton(master=self.frame_rfr, text="Cadastro", command=self.RFRContador6)
        self.label_rfr6 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr7 = Ctk.CTkButton(master=self.frame_rfr, text="Dúvidas", command=self.RFRContador7)
        self.label_rfr7 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        self.btn_rfr8 = Ctk.CTkButton(master=self.frame_rfr, text="Outros", command=self.RFRContador8)
        self.label_rfr8 = Ctk.CTkLabel(master=self.frame_rfr, text="0")

        #Indaiatuba
        self.btn_ind1 = Ctk.CTkButton(master=self.frame_ind, text="Cadastro Omnilink", command=self.IndContador1)
        self.label_ind1 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind2 = Ctk.CTkButton(master=self.frame_ind, text="Verificar Espelhamento", command=self.IndContador2)
        self.label_ind2 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind3 = Ctk.CTkButton(master=self.frame_ind, text="Trocar Rastreador", command=self.IndContador3)
        self.label_ind3 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind4 = Ctk.CTkButton(master=self.frame_ind, text="Voltar para PNA", command=self.IndContador4)
        self.label_ind4 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind5 = Ctk.CTkButton(master=self.frame_ind, text="Treinamento", command=self.IndContador5)
        self.label_ind5 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind6 = Ctk.CTkButton(master=self.frame_ind, text="Cadastro", command=self.IndContador6)
        self.label_ind6 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind7 = Ctk.CTkButton(master=self.frame_ind, text="Dúvidas", command=self.IndContador7)
        self.label_ind7 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        self.btn_ind8 = Ctk.CTkButton(master=self.frame_ind, text="Outros", command=self.IndContador8)
        self.label_ind8 = Ctk.CTkLabel(master=self.frame_ind, text="0")

        #GKN
        self.btn_gkn1 = Ctk.CTkButton(master=self.frame_gkn, text="Cadastro Omnilink", command=self.GKNContador1)
        self.label_gkn1 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn2 = Ctk.CTkButton(master=self.frame_gkn, text="Verificar Espelhamento", command=self.GKNContador2)
        self.label_gkn2 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn3 = Ctk.CTkButton(master=self.frame_gkn, text="Trocar Rastreador", command=self.GKNContador3)
        self.label_gkn3 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn4 = Ctk.CTkButton(master=self.frame_gkn, text="Voltar para PNA", command=self.GKNContador4)
        self.label_gkn4 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn5 = Ctk.CTkButton(master=self.frame_gkn, text="Treinamento", command=self.GKNContador5)
        self.label_gkn5 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn6 = Ctk.CTkButton(master=self.frame_gkn, text="Cadastro", command=self.GKNContador6)
        self.label_gkn6 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn7 = Ctk.CTkButton(master=self.frame_gkn, text="Dúvidas", command=self.GKNContador7)
        self.label_gkn7 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        self.btn_gkn8 = Ctk.CTkButton(master=self.frame_gkn, text="Outros", command=self.GKNContador8)
        self.label_gkn8 = Ctk.CTkLabel(master=self.frame_gkn, text="0")

        #Coopercarga
        self.btn_coopercarga1 = Ctk.CTkButton(master=self.frame_coopercarga, text="Cadastro Omnilink", command=self.CoopercargaContador1)
        self.label_coopercarga1 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga2 = Ctk.CTkButton(master=self.frame_coopercarga, text="Verificar Espelhamento", command=self.CoopercargaContador2)
        self.label_coopercarga2 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga3 = Ctk.CTkButton(master=self.frame_coopercarga, text="Trocar Rastreador", command=self.CoopercargaContador3)
        self.label_coopercarga3 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga4 = Ctk.CTkButton(master=self.frame_coopercarga, text="Voltar para PNA", command=self.CoopercargaContador4)
        self.label_coopercarga4 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga5 = Ctk.CTkButton(master=self.frame_coopercarga, text="Treinamento", command=self.CoopercargaContador5)
        self.label_coopercarga5 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga6 = Ctk.CTkButton(master=self.frame_coopercarga, text="Cadastro", command=self.CoopercargaContador6)
        self.label_coopercarga6 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga7 = Ctk.CTkButton(master=self.frame_coopercarga, text="Dúvidas", command=self.CoopercargaContador7)
        self.label_coopercarga7 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        self.btn_coopercarga8 = Ctk.CTkButton(master=self.frame_coopercarga, text="Outros", command=self.CoopercargaContador8)
        self.label_coopercarga8 = Ctk.CTkLabel(master=self.frame_coopercarga, text="0")

        # Exportar para Excel
        self.btn_exportar_excel = Ctk.CTkButton(master=self.frame_acoes, text="Exportar para Excel", fg_color="#52BE80", hover=True, hover_color="#229954", command=self.ExportarExcel)
        
        # Limpeza de contadores
        self.btn_limpar = Ctk.CTkButton(master=self.frame_acoes, text="Limpar Contadores", fg_color="#E74C3C", hover=True, hover_color="#943126", command=self.confirmar_limpeza_contadores)
        
        ## Posicionamento dos botões ##
        #Frames e Labels titulos

        # Título
        # self.frame_titulo.grid(row=0, column=6, pady=2, padx=2)
        # self.label_titulo.grid(row=0, column=6, pady=2, padx=2)
        
        self.label_suzano.grid(row=0, column=0, columnspan=2, pady=2, padx=2)
        self.frame_suzano.grid(row=4, column=0, pady=2, padx=2)
        
        self.label_adami.grid(row=0, column=2, columnspan=2, pady=2, padx=2)
        self.frame_adami.grid(row=4, column=2, pady=2, padx=2)

        self.label_klabin.grid(row=0, column=4, columnspan=2, pady=2, padx=2)
        self.frame_klabin.grid(row=4, column=4, pady=2, padx=2)
        
        self.label_irani.grid(row=0, column=6, columnspan=2, pady=2, padx=2)
        self.frame_irani.grid(row=4, column=6, pady=2, padx=2)

        self.label_rfr.grid(row=10, column=0, columnspan=2, pady=2, padx=2)
        self.frame_rfr.grid(row=14, column=0, pady=2, padx=2)

        self.label_ind.grid(row=10, column=2, columnspan=2, pady=2, padx=2)
        self.frame_ind.grid(row=14, column=2, pady=2, padx=2)

        self.label_gkn.grid(row=10, column=4, columnspan=2, pady=2, padx=2)
        self.frame_gkn.grid(row=14, column=4, pady=2, padx=2)

        self.label_coopercarga.grid(row=10, column=6, columnspan=2, pady=2, padx=2)
        self.frame_coopercarga.grid(row=14, column=6, pady=2, padx=2)

        self.frame_acoes.grid(row=20, column=0, columnspan=8, pady=2, padx=2)
        
        # Suzano 
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
        self.btn_adami1.grid(row=1, column=3, pady=10, padx=10)
        self.label_adami1.grid(row=1, column=4, pady=10, padx=10)

        self.btn_adami2.grid(row=2, column=3, pady=10, padx=10)
        self.label_adami2.grid(row=2, column=4, pady=10, padx=10)

        self.btn_adami3.grid(row=3, column=3, pady=10, padx=10)
        self.label_adami3.grid(row=3, column=4, pady=10, padx=10)

        self.btn_adami4.grid(row=4, column=3, pady=10, padx=10)
        self.label_adami4.grid(row=4, column=4, pady=10, padx=10)

        self.btn_adami5.grid(row=5, column=3, pady=10, padx=10)
        self.label_adami5.grid(row=5, column=4, pady=10, padx=10)

        self.btn_adami6.grid(row=6, column=3, pady=10, padx=10)
        self.label_adami6.grid(row=6, column=4, pady=10, padx=10)

        self.btn_adami7.grid(row=7, column=3, pady=10, padx=10)
        self.label_adami7.grid(row=7, column=4, pady=10, padx=10)

        self.btn_adami8.grid(row=8, column=3, pady=10, padx=10)
        self.label_adami8.grid(row=8, column=4, pady=10, padx=10)
        
        # Klabin
        self.btn_klabin1.grid(row=1, column=5, pady=10, padx=10)
        self.label_klabin1.grid(row=1, column=6, pady=10, padx=10)
        
        self.btn_klabin2.grid(row=2, column=5, pady=10, padx=10)
        self.label_klabin2.grid(row=2, column=6, pady=10, padx=10)
        
        self.btn_klabin3.grid(row=3, column=5, pady=10, padx=10)
        self.label_klabin3.grid(row=3, column=6, pady=10, padx=10)
        
        self.btn_klabin4.grid(row=4, column=5, pady=10, padx=10)
        self.label_klabin4.grid(row=4, column=6, pady=10, padx=10)
        
        self.btn_klabin5.grid(row=5, column=5, pady=10, padx=10)
        self.label_klabin5.grid(row=5, column=6, pady=10, padx=10)
        
        self.btn_klabin6.grid(row=6, column=5, pady=10, padx=10)
        self.label_klabin6.grid(row=6, column=6, pady=10, padx=10)
        
        self.btn_klabin7.grid(row=7, column=5, pady=10, padx=10)
        self.label_klabin7.grid(row=7, column=6, pady=10, padx=10)
        
        self.btn_klabin8.grid(row=8, column=5, pady=10, padx=10)
        self.label_klabin8.grid(row=8, column=6, pady=10, padx=10)
        
        #Irani
        self.btn_irani1.grid(row=1, column=7, pady=10, padx=10)
        self.label_irani1.grid(row=1, column=8, pady=10, padx=10)
        
        self.btn_irani2.grid(row=2, column=7, pady=10, padx=10)
        self.label_irani2.grid(row=2, column=8, pady=10, padx=10)
        
        self.btn_irani3.grid(row=3, column=7, pady=10, padx=10)
        self.label_irani3.grid(row=3, column=8, pady=10, padx=10)
        
        self.btn_irani4.grid(row=4, column=7, pady=10, padx=10)
        self.label_irani4.grid(row=4, column=8, pady=10, padx=10)
        
        self.btn_irani5.grid(row=5, column=7, pady=10, padx=10)
        self.label_irani5.grid(row=5, column=8, pady=10, padx=10)
        
        self.btn_irani6.grid(row=6, column=7, pady=10, padx=10)
        self.label_irani6.grid(row=6, column=8, pady=10, padx=10)
        
        self.btn_irani7.grid(row=7, column=7, pady=10, padx=10)
        self.label_irani7.grid(row=7, column=8, pady=10, padx=10)
        
        self.btn_irani8.grid(row=8, column=7, pady=10, padx=10)
        self.label_irani8.grid(row=8, column=8, pady=10, padx=10)

        #RFR
        self.btn_rfr1.grid(row=11, column=1, pady=10, padx=10)
        self.label_rfr1.grid(row=11, column=2, pady=10, padx=10)

        self.btn_rfr2.grid(row=12, column=1, pady=10, padx=10)
        self.label_rfr2.grid(row=12, column=2, pady=10, padx=10)

        self.btn_rfr3.grid(row=13, column=1, pady=10, padx=10)
        self.label_rfr3.grid(row=13, column=2, pady=10, padx=10)

        self.btn_rfr4.grid(row=14, column=1, pady=10, padx=10)
        self.label_rfr4.grid(row=14, column=2, pady=10, padx=10)

        self.btn_rfr5.grid(row=15, column=1, pady=10, padx=10)
        self.label_rfr5.grid(row=15, column=2, pady=10, padx=10)

        self.btn_rfr6.grid(row=16, column=1, pady=10, padx=10)
        self.label_rfr6.grid(row=16, column=2, pady=10, padx=10)

        self.btn_rfr7.grid(row=17, column=1, pady=10, padx=10)
        self.label_rfr7.grid(row=17, column=2, pady=10, padx=10)

        self.btn_rfr8.grid(row=18, column=1, pady=10, padx=10)
        self.label_rfr8.grid(row=18, column=2, pady=10, padx=10)

        #Indaiatuba
        self.btn_ind1.grid(row=11, column=3, pady=10, padx=10)
        self.label_ind1.grid(row=11, column=4, pady=10, padx=10)

        self.btn_ind2.grid(row=12, column=3, pady=10, padx=10)
        self.label_ind2.grid(row=12, column=4, pady=10, padx=10)

        self.btn_ind3.grid(row=13, column=3, pady=10, padx=10)
        self.label_ind3.grid(row=13, column=4, pady=10, padx=10)

        self.btn_ind4.grid(row=14, column=3, pady=10, padx=10)
        self.label_ind4.grid(row=14, column=4, pady=10, padx=10)

        self.btn_ind5.grid(row=15, column=3, pady=10, padx=10)
        self.label_ind5.grid(row=15, column=4, pady=10, padx=10)

        self.btn_ind6.grid(row=16, column=3, pady=10, padx=10)
        self.label_ind6.grid(row=16, column=4, pady=10, padx=10)

        self.btn_ind7.grid(row=17, column=3, pady=10, padx=10)
        self.label_ind7.grid(row=17, column=4, pady=10, padx=10)

        self.btn_ind8.grid(row=18, column=3, pady=10, padx=10)
        self.label_ind8.grid(row=18, column=4, pady=10, padx=10)

        #GKN
        self.btn_gkn1.grid(row=11, column=5, pady=10, padx=10)
        self.label_gkn1.grid(row=11, column=6, pady=10, padx=10)

        self.btn_gkn2.grid(row=12, column=5, pady=10, padx=10)
        self.label_gkn2.grid(row=12, column=6, pady=10, padx=10)

        self.btn_gkn3.grid(row=13, column=5, pady=10, padx=10)
        self.label_gkn3.grid(row=13, column=6, pady=10, padx=10)

        self.btn_gkn4.grid(row=14, column=5, pady=10, padx=10)
        self.label_gkn4.grid(row=14, column=6, pady=10, padx=10)

        self.btn_gkn5.grid(row=15, column=5, pady=10, padx=10)
        self.label_gkn5.grid(row=15, column=6, pady=10, padx=10)

        self.btn_gkn6.grid(row=16, column=5, pady=10, padx=10)
        self.label_gkn6.grid(row=16, column=6, pady=10, padx=10)

        self.btn_gkn7.grid(row=17, column=5, pady=10, padx=10)
        self.label_gkn7.grid(row=17, column=6, pady=10, padx=10)

        self.btn_gkn8.grid(row=18, column=5, pady=10, padx=10)
        self.label_gkn8.grid(row=18, column=6, pady=10, padx=10)

        #Coopercarga
        self.btn_coopercarga1.grid(row=11, column=7, pady=10, padx=10)
        self.label_coopercarga1.grid(row=11, column=8, pady=10, padx=10)

        self.btn_coopercarga2.grid(row=12, column=7, pady=10, padx=10)
        self.label_coopercarga2.grid(row=12, column=8, pady=10, padx=10)

        self.btn_coopercarga3.grid(row=13, column=7, pady=10, padx=10)
        self.label_coopercarga3.grid(row=13, column=8, pady=10, padx=10)

        self.btn_coopercarga4.grid(row=14, column=7, pady=10, padx=10)
        self.label_coopercarga4.grid(row=14, column=8, pady=10, padx=10)

        self.btn_coopercarga5.grid(row=15, column=7, pady=10, padx=10)
        self.label_coopercarga5.grid(row=15, column=8, pady=10, padx=10)

        self.btn_coopercarga6.grid(row=16, column=7, pady=10, padx=10)
        self.label_coopercarga6.grid(row=16, column=8, pady=10, padx=10)

        self.btn_coopercarga7.grid(row=17, column=7, pady=10, padx=10)
        self.label_coopercarga7.grid(row=17, column=8, pady=10, padx=10)

        self.btn_coopercarga8.grid(row=18, column=7, pady=10, padx=10)
        self.label_coopercarga8.grid(row=18, column=8, pady=10, padx=10)


        # Exportar para Excel
        self.btn_exportar_excel.grid(row=20, column=2, pady=5, padx=10)
        
        # Limpeza de contadores
        self.btn_limpar.grid(row=20, column=4, pady=5, padx=10)
        
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
        
    #Klabin
    def KlabinContador1(self):
        self.contador_klabin1 += 1
        self.atualizar_rotulo(self.label_klabin1, self.contador_klabin1)
    
    def KlabinContador2(self):
        self.contador_klabin2 += 1
        self.atualizar_rotulo(self.label_klabin2, self.contador_klabin2)
    
    def KlabinContador3(self):
        self.contador_klabin3 += 1
        self.atualizar_rotulo(self.label_klabin3, self.contador_klabin3)
    
    def KlabinContador4(self):
        self.contador_klabin4 += 1
        self.atualizar_rotulo(self.label_klabin4, self.contador_klabin4)
    
    def KlabinContador5(self):
        self.contador_klabin5 += 1
        self.atualizar_rotulo(self.label_klabin5, self.contador_klabin5)
    
    def KlabinContador6(self):
        self.contador_klabin6 += 1
        self.atualizar_rotulo(self.label_klabin6, self.contador_klabin6)
    
    def KlabinContador7(self):
        self.contador_klabin7 += 1
        self.atualizar_rotulo(self.label_klabin7, self.contador_klabin7)
    
    def KlabinContador8(self):
        self.contador_klabin8 += 1
        self.atualizar_rotulo(self.label_klabin8, self.contador_klabin8)
        
    #Irani
    
    def IraniContador1(self):
        self.contador_irani1 += 1
        self.atualizar_rotulo(self.label_irani1, self.contador_irani1)
    
    def IraniContador2(self):
        self.contador_irani2 += 1
        self.atualizar_rotulo(self.label_irani2, self.contador_irani2)
    
    def IraniContador3(self):
        self.contador_irani3 += 1
        self.atualizar_rotulo(self.label_irani3, self.contador_irani3)
    
    def IraniContador4(self):
        self.contador_irani4 += 1
        self.atualizar_rotulo(self.label_irani4, self.contador_irani4)
    
    def IraniContador5(self):
        self.contador_irani5 += 1
        self.atualizar_rotulo(self.label_irani5, self.contador_irani5)
    
    def IraniContador6(self):
        self.contador_irani6 += 1
        self.atualizar_rotulo(self.label_irani6, self.contador_irani6)
    
    def IraniContador7(self):
        self.contador_irani7 += 1
        self.atualizar_rotulo(self.label_irani7, self.contador_irani7)
    
    def IraniContador8(self):
        self.contador_irani8 += 1
        self.atualizar_rotulo(self.label_irani8, self.contador_irani8)

    #RFR
    def RFRContador1(self):
        self.contador_rfr1 += 1
        self.atualizar_rotulo(self.label_rfr1, self.contador_rfr1)
    
    def RFRContador2(self):
        self.contador_rfr2 += 1
        self.atualizar_rotulo(self.label_rfr2, self.contador_rfr2)

    def RFRContador3(self):
        self.contador_rfr3 += 1
        self.atualizar_rotulo(self.label_rfr3, self.contador_rfr3)
    
    def RFRContador4(self):
        self.contador_rfr4 += 1
        self.atualizar_rotulo(self.label_rfr4, self.contador_rfr4)
    
    def RFRContador5(self):
        self.contador_rfr5 += 1
        self.atualizar_rotulo(self.label_rfr5, self.contador_rfr5)
    
    def RFRContador6(self):
        self.contador_rfr6 += 1
        self.atualizar_rotulo(self.label_rfr6, self.contador_rfr6)
    
    def RFRContador7(self):
        self.contador_rfr7 += 1
        self.atualizar_rotulo(self.label_rfr7, self.contador_rfr7)
    
    def RFRContador8(self):
        self.contador_rfr8 += 1
        self.atualizar_rotulo(self.label_rfr8, self.contador_rfr8)
    
    #Indaiatuba
    def IndContador1(self):
        self.contador_ind1 += 1
        self.atualizar_rotulo(self.label_ind1, self.contador_ind1)
    
    def IndContador2(self):
        self.contador_ind2 += 1
        self.atualizar_rotulo(self.label_ind2, self.contador_ind2)
    
    def IndContador3(self):
        self.contador_ind3 += 1
        self.atualizar_rotulo(self.label_ind3, self.contador_ind3)
    
    def IndContador4(self):
        self.contador_ind4 += 1
        self.atualizar_rotulo(self.label_ind4, self.contador_ind4)
    
    def IndContador5(self):
        self.contador_ind5 += 1
        self.atualizar_rotulo(self.label_ind5, self.contador_ind5)
    
    def IndContador6(self):
        self.contador_ind6 += 1
        self.atualizar_rotulo(self.label_ind6, self.contador_ind6)
    
    def IndContador7(self):
        self.contador_ind7 += 1
        self.atualizar_rotulo(self.label_ind7, self.contador_ind7)
    
    def IndContador8(self):
        self.contador_ind8 += 1
        self.atualizar_rotulo(self.label_ind8, self.contador_ind8)

    #GKN
    def GKNContador1(self):
        self.contador_gkn1 += 1
        self.atualizar_rotulo(self.label_gkn1, self.contador_gkn1)
    
    def GKNContador2(self):
        self.contador_gkn2 += 1
        self.atualizar_rotulo(self.label_gkn2, self.contador_gkn2)
    
    def GKNContador3(self):
        self.contador_gkn3 += 1
        self.atualizar_rotulo(self.label_gkn3, self.contador_gkn3)
    
    def GKNContador4(self):
        self.contador_gkn4 += 1
        self.atualizar_rotulo(self.label_gkn4, self.contador_gkn4)
    
    def GKNContador5(self):
        self.contador_gkn5 += 1
        self.atualizar_rotulo(self.label_gkn5, self.contador_gkn5)
    
    def GKNContador6(self):
        self.contador_gkn6 += 1
        self.atualizar_rotulo(self.label_gkn6, self.contador_gkn6)
    
    def GKNContador7(self):
        self.contador_gkn7 += 1
        self.atualizar_rotulo(self.label_gkn7, self.contador_gkn7)
    
    def GKNContador8(self):
        self.contador_gkn8 += 1
        self.atualizar_rotulo(self.label_gkn8, self.contador_gkn8)
        
    #Coopercarga
    def CoopercargaContador1(self):
        self.contador_coopercarga1 += 1
        self.atualizar_rotulo(self.label_coopercarga1, self.contador_coopercarga1)

    def CoopercargaContador2(self):
        self.contador_coopercarga2 += 1
        self.atualizar_rotulo(self.label_coopercarga2, self.contador_coopercarga2)

    def CoopercargaContador3(self):
        self.contador_coopercarga3 += 1
        self.atualizar_rotulo(self.label_coopercarga3, self.contador_coopercarga3)

    def CoopercargaContador4(self):
        self.contador_coopercarga4 += 1
        self.atualizar_rotulo(self.label_coopercarga4, self.contador_coopercarga4)

    def CoopercargaContador5(self):
        self.contador_coopercarga5 += 1
        self.atualizar_rotulo(self.label_coopercarga5, self.contador_coopercarga5)

    def CoopercargaContador6(self):
        self.contador_coopercarga6 += 1
        self.atualizar_rotulo(self.label_coopercarga6, self.contador_coopercarga6)

    def CoopercargaContador7(self):
        self.contador_coopercarga7 += 1
        self.atualizar_rotulo(self.label_coopercarga7, self.contador_coopercarga7)

    def CoopercargaContador8(self):
        self.contador_coopercarga8 += 1
        self.atualizar_rotulo(self.label_coopercarga8, self.contador_coopercarga8)    
    
    #Atualizar Rotulo
    def atualizar_rotulo(self, rotulo, contador):
        rotulo.configure(text=f"{contador}")
    
    def ExportarExcel(self):
         
        diretorio_script = os.path.dirname(os.path.realpath(__file__))
        
        arquivo_excel = os.path.join(diretorio_script, "contagem_atendimentos.xlsx")
        
        # Cria planilha
        workbook = Workbook()
        
        # Adicionar folha
        sheet = workbook.active
        
        # Cabeçalhos Empresas
        sheet['A1'] = "Empresa"
        sheet['A2'] = "Suzano"
        sheet['A3'] = "Adami"
        sheet['A4'] = "Klabin"
        sheet['A5'] = "Irani"
        sheet['A6'] = "RFR Guarulhos"
        sheet['A7'] = "Indaiatuba"
        sheet['A8'] = "GKN"
        sheet['A9'] = "Coopercarga"
        
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
        
        sheet['B3'] = self.contador_adami1
        sheet['C3'] = self.contador_adami2
        sheet['D3'] = self.contador_adami3
        sheet['E3'] = self.contador_adami4
        sheet['F3'] = self.contador_adami5
        sheet['G3'] = self.contador_adami6
        sheet['H3'] = self.contador_adami7
        sheet['I3'] = self.contador_adami8
        
        sheet['B4'] = self.contador_klabin1
        sheet['C4'] = self.contador_klabin2
        sheet['D4'] = self.contador_klabin3
        sheet['E4'] = self.contador_klabin4
        sheet['F4'] = self.contador_klabin5
        sheet['G4'] = self.contador_klabin6
        sheet['H4'] = self.contador_klabin7
        sheet['I4'] = self.contador_klabin8
        
        sheet['B5'] = self.contador_irani1
        sheet['C5'] = self.contador_irani2
        sheet['D5'] = self.contador_irani3
        sheet['E5'] = self.contador_irani4
        sheet['F5'] = self.contador_irani5
        sheet['G5'] = self.contador_irani6
        sheet['H5'] = self.contador_irani7
        sheet['I5'] = self.contador_irani8

        sheet['B6'] = self.contador_rfr1
        sheet['C6'] = self.contador_rfr2
        sheet['D6'] = self.contador_rfr3
        sheet['E6'] = self.contador_rfr4
        sheet['F6'] = self.contador_rfr5
        sheet['G6'] = self.contador_rfr6
        sheet['H6'] = self.contador_rfr7
        sheet['I6'] = self.contador_rfr8

        sheet['B7'] = self.contador_ind1
        sheet['C7'] = self.contador_ind2
        sheet['D7'] = self.contador_ind3
        sheet['E7'] = self.contador_ind4
        sheet['F7'] = self.contador_ind5
        sheet['G7'] = self.contador_ind6
        sheet['H7'] = self.contador_ind7
        sheet['I7'] = self.contador_ind8

        sheet['B8'] = self.contador_gkn1
        sheet['C8'] = self.contador_gkn2
        sheet['D8'] = self.contador_gkn3
        sheet['E8'] = self.contador_gkn4
        sheet['F8'] = self.contador_gkn5
        sheet['G8'] = self.contador_gkn6
        sheet['H8'] = self.contador_gkn7
        sheet['I8'] = self.contador_gkn8

        sheet['B9'] = self.contador_coopercarga1
        sheet['C9'] = self.contador_coopercarga2
        sheet['D9'] = self.contador_coopercarga3
        sheet['E9'] = self.contador_coopercarga4
        sheet['F9'] = self.contador_coopercarga5
        sheet['G9'] = self.contador_coopercarga6
        sheet['H9'] = self.contador_coopercarga7
        sheet['I9'] = self.contador_coopercarga8
         
        # Salvar planilha
        # arquivo_excel = "contagem_atendimentos.xlsx"
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
        
        self.contador_klabin1 = 0
        self.contador_klabin2 = 0
        self.contador_klabin3 = 0
        self.contador_klabin4 = 0
        self.contador_klabin5 = 0
        self.contador_klabin6 = 0
        self.contador_klabin7 = 0
        self.contador_klabin8 = 0
        
        self.contador_irani1 = 0
        self.contador_irani2 = 0
        self.contador_irani3 = 0
        self.contador_irani4 = 0
        self.contador_irani5 = 0
        self.contador_irani6 = 0
        self.contador_irani7 = 0
        self.contador_irani8 = 0

        self.contador_rfr1 = 0
        self.contador_rfr2 = 0
        self.contador_rfr3 = 0
        self.contador_rfr4 = 0
        self.contador_rfr5 = 0
        self.contador_rfr6 = 0
        self.contador_rfr7 = 0
        self.contador_rfr8 = 0

        self.contador_ind1 = 0
        self.contador_ind2 = 0
        self.contador_ind3 = 0
        self.contador_ind4 = 0
        self.contador_ind5 = 0
        self.contador_ind6 = 0
        self.contador_ind7 = 0
        self.contador_ind8 = 0

        self.contador_gkn1 = 0
        self.contador_gkn2 = 0
        self.contador_gkn3 = 0
        self.contador_gkn4 = 0
        self.contador_gkn5 = 0
        self.contador_gkn6 = 0
        self.contador_gkn7 = 0
        self.contador_gkn8 = 0

        self.contador_coopercarga1 = 0
        self.contador_coopercarga2 = 0
        self.contador_coopercarga3 = 0
        self.contador_coopercarga4 = 0
        self.contador_coopercarga5 = 0
        self.contador_coopercarga6 = 0
        self.contador_coopercarga7 = 0
        self.contador_coopercarga8 = 0
        
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
        
        self.atualizar_rotulo(self.label_klabin1, self.contador_klabin1)
        self.atualizar_rotulo(self.label_klabin2, self.contador_klabin2)
        self.atualizar_rotulo(self.label_klabin3, self.contador_klabin3)
        self.atualizar_rotulo(self.label_klabin4, self.contador_klabin4)
        self.atualizar_rotulo(self.label_klabin5, self.contador_klabin5)
        self.atualizar_rotulo(self.label_klabin6, self.contador_klabin6)
        self.atualizar_rotulo(self.label_klabin7, self.contador_klabin7)
        self.atualizar_rotulo(self.label_klabin8, self.contador_klabin8)
        
        self.atualizar_rotulo(self.label_irani1, self.contador_irani1)
        self.atualizar_rotulo(self.label_irani2, self.contador_irani2)
        self.atualizar_rotulo(self.label_irani3, self.contador_irani3)
        self.atualizar_rotulo(self.label_irani4, self.contador_irani4)
        self.atualizar_rotulo(self.label_irani5, self.contador_irani5)
        self.atualizar_rotulo(self.label_irani6, self.contador_irani6)
        self.atualizar_rotulo(self.label_irani7, self.contador_irani7)
        self.atualizar_rotulo(self.label_irani8, self.contador_irani8)

        self.atualizar_rotulo(self.label_rfr1, self.contador_rfr1)
        self.atualizar_rotulo(self.label_rfr2, self.contador_rfr2)
        self.atualizar_rotulo(self.label_rfr3, self.contador_rfr3)
        self.atualizar_rotulo(self.label_rfr4, self.contador_rfr4)
        self.atualizar_rotulo(self.label_rfr5, self.contador_rfr5)
        self.atualizar_rotulo(self.label_rfr6, self.contador_rfr6)
        self.atualizar_rotulo(self.label_rfr7, self.contador_rfr7)
        self.atualizar_rotulo(self.label_rfr8, self.contador_rfr8)

        self.atualizar_rotulo(self.label_ind1, self.contador_ind1)
        self.atualizar_rotulo(self.label_ind2, self.contador_ind2)
        self.atualizar_rotulo(self.label_ind3, self.contador_ind3)
        self.atualizar_rotulo(self.label_ind4, self.contador_ind4)
        self.atualizar_rotulo(self.label_ind5, self.contador_ind5)
        self.atualizar_rotulo(self.label_ind6, self.contador_ind6)
        self.atualizar_rotulo(self.label_ind7, self.contador_ind7)
        self.atualizar_rotulo(self.label_ind8, self.contador_ind8)

        self.atualizar_rotulo(self.label_gkn1, self.contador_gkn1)
        self.atualizar_rotulo(self.label_gkn2, self.contador_gkn2)
        self.atualizar_rotulo(self.label_gkn3, self.contador_gkn3)
        self.atualizar_rotulo(self.label_gkn4, self.contador_gkn4)
        self.atualizar_rotulo(self.label_gkn5, self.contador_gkn5)
        self.atualizar_rotulo(self.label_gkn6, self.contador_gkn6)
        self.atualizar_rotulo(self.label_gkn7, self.contador_gkn7)
        self.atualizar_rotulo(self.label_gkn8, self.contador_gkn8)

        self.atualizar_rotulo(self.label_coopercarga1, self.contador_coopercarga1)
        self.atualizar_rotulo(self.label_coopercarga2, self.contador_coopercarga2)
        self.atualizar_rotulo(self.label_coopercarga3, self.contador_coopercarga3)
        self.atualizar_rotulo(self.label_coopercarga4, self.contador_coopercarga4)
        self.atualizar_rotulo(self.label_coopercarga5, self.contador_coopercarga5)
        self.atualizar_rotulo(self.label_coopercarga6, self.contador_coopercarga6)
        self.atualizar_rotulo(self.label_coopercarga7, self.contador_coopercarga7)
        self.atualizar_rotulo(self.label_coopercarga8, self.contador_coopercarga8)
        
    def salvar_ultima_sessao(self):
        diretorio_script = os.path.dirname(os.path.realpath(__file__))

        arquivo_ultima_sessao = os.path.join(diretorio_script, "ultima_sessao.txt")

        #Obter o diretório do script
        caminho_arquivo_json = os.path.join(diretorio_script, "contagem_atendimentos.json")

        # Carregar histórico existente ou criar uma lista vazia
        try:
            with open(caminho_arquivo_json, "r") as arquivo:
                historico_existente = json.load(arquivo)
        except (json.JSONDecodeError, FileNotFoundError):
            historico_existente = []

        # Adicionar novo item ao histórico
        nova_entrada = {
            "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "contadores": {
                "suzano":{
                    "Cadastro_Omnilink": self.contador_suzano1,
                    "Verificar_Espelhamento": self.contador_suzano2,
                    "Trocar_Rastreador": self.contador_suzano3,
                    "Voltar_PNA": self.contador_suzano4,
                    "Treinamento": self.contador_suzano5,
                    "Cadastro": self.contador_suzano6,
                    "Duvidas": self.contador_suzano7,
                    "Outros": self.contador_suzano8
                },
                "adami":{
                    "Cadastro_Omnilink": self.contador_adami1,
                    "Verificar_Espelhamento": self.contador_adami2,
                    "Trocar_Rastreador": self.contador_adami3,
                    "Voltar_PNA": self.contador_adami4,
                    "Treinamento": self.contador_adami5,
                    "Cadastro": self.contador_adami6,
                    "Duvidas": self.contador_adami7,
                    "Outros": self.contador_adami8
                },
                "klabin":{
                    "Cadastro_Omnilink": self.contador_klabin1,
                    "Verificar_Espelhamento": self.contador_klabin2,
                    "Trocar_Rastreador": self.contador_klabin3,
                    "Voltar_PNA": self.contador_klabin4,
                    "Treinamento": self.contador_klabin5,
                    "Cadastro": self.contador_klabin6,
                    "Duvidas": self.contador_klabin7,
                    "Outros": self.contador_klabin8
                },
                "irani":{
                    "Cadastro_Omnilink": self.contador_irani1,
                    "Verificar_Espelhamento": self.contador_irani2,
                    "Trocar_Rastreador": self.contador_irani3,
                    "Voltar_PNA": self.contador_irani4,
                    "Treinamento": self.contador_irani5,
                    "Cadastro": self.contador_irani6,
                    "Duvidas": self.contador_irani7,
                    "Outros": self.contador_irani8
                },
                "rfr":{
                    "Cadastro_Omnilink": self.contador_rfr1,
                    "Verificar_Espelhamento": self.contador_rfr2,
                    "Trocar_Rastreador": self.contador_rfr3,
                    "Voltar_PNA": self.contador_rfr4,
                    "Treinamento": self.contador_rfr5,
                    "Cadastro": self.contador_rfr6,
                    "Duvidas": self.contador_rfr7,
                    "Outros": self.contador_rfr8
                },
                "indaiatuba":{
                    "Cadastro_Omnilink": self.contador_ind1,
                    "Verificar_Espelhamento": self.contador_ind2,
                    "Trocar_Rastreador": self.contador_ind3,
                    "Voltar_PNA": self.contador_ind4,
                    "Treinamento": self.contador_ind5,
                    "Cadastro": self.contador_ind6,
                    "Duvidas": self.contador_ind7,
                    "Outros": self.contador_ind8
                },
                "gkn":{
                    "Cadastro_Omnilink": self.contador_gkn1,
                    "Verificar_Espelhamento": self.contador_gkn2,
                    "Trocar_Rastreador": self.contador_gkn3,
                    "Voltar_PNA": self.contador_gkn4,
                    "Treinamento": self.contador_gkn5,
                    "Cadastro": self.contador_gkn6,
                    "Duvidas": self.contador_gkn7,
                    "Outros": self.contador_gkn8
                },
                "coopercarga":{
                    "Cadastro_Omnilink": self.contador_coopercarga1,
                    "Verificar_Espelhamento": self.contador_coopercarga2,
                    "Trocar_Rastreador": self.contador_coopercarga3,
                    "Voltar_PNA": self.contador_coopercarga4,
                    "Treinamento": self.contador_coopercarga5,
                    "Cadastro": self.contador_coopercarga6,
                    "Duvidas": self.contador_coopercarga7,
                    "Outros": self.contador_coopercarga8
                }
            }
        }

        historico_existente.append(nova_entrada)

        #Salvar o historico atualizado no arquivo JSON
        with open(caminho_arquivo_json, "w") as arquivo:
            json.dump(historico_existente, arquivo, indent=4)
        
        with open(arquivo_ultima_sessao, "w") as arquivo:
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
            
            arquivo.write(f"Klabin Cadastro_Omnilink: {self.contador_klabin1}\n")
            arquivo.write(f"Klabin Verificar_Espelhamento: {self.contador_klabin2}\n")
            arquivo.write(f"Klabin Trocar_Rastreador: {self.contador_klabin3}\n")
            arquivo.write(f"Klabin Voltar_PNA: {self.contador_klabin4}\n")
            arquivo.write(f"Klabin Treinamento: {self.contador_klabin5}\n")
            arquivo.write(f"Klabin Cadastro: {self.contador_klabin6}\n")
            arquivo.write(f"Klabin Dúvidas: {self.contador_klabin7}\n")
            arquivo.write(f"Klabin Outros: {self.contador_klabin8}\n")
            
            arquivo.write(f"Irani Cadastro_Omnilink: {self.contador_irani1}\n")
            arquivo.write(f"Irani Verificar_Espelhamento: {self.contador_irani2}\n")
            arquivo.write(f"Irani Trocar_Rastreador: {self.contador_irani3}\n")
            arquivo.write(f"Irani Voltar_PNA: {self.contador_irani4}\n")
            arquivo.write(f"Irani Treinamento: {self.contador_irani5}\n")
            arquivo.write(f"Irani Cadastro: {self.contador_irani6}\n")
            arquivo.write(f"Irani Dúvidas: {self.contador_irani7}\n")
            arquivo.write(f"Irani Outros: {self.contador_irani8}\n")

            arquivo.write(f"RFR Cadastro_Omnilink: {self.contador_rfr1}\n")
            arquivo.write(f"RFR Verificar_Espelhamento: {self.contador_rfr2}\n")
            arquivo.write(f"RFR Trocar_Rastreador: {self.contador_rfr3}\n")
            arquivo.write(f"RFR Voltar_PNA: {self.contador_rfr4}\n")
            arquivo.write(f"RFR Treinamento: {self.contador_rfr5}\n")
            arquivo.write(f"RFR Cadastro: {self.contador_rfr6}\n")
            arquivo.write(f"RFR Dúvidas: {self.contador_rfr7}\n")
            arquivo.write(f"RFR Outros: {self.contador_rfr8}\n")

            arquivo.write(f"Indaiatuba Cadastro_Omnilink: {self.contador_ind1}\n")
            arquivo.write(f"Indaiatuba Verificar_Espelhamento: {self.contador_ind2}\n")
            arquivo.write(f"Indaiatuba Trocar_Rastreador: {self.contador_ind3}\n")
            arquivo.write(f"Indaiatuba Voltar_PNA: {self.contador_ind4}\n")
            arquivo.write(f"Indaiatuba Treinamento: {self.contador_ind5}\n")
            arquivo.write(f"Indaiatuba Cadastro: {self.contador_ind6}\n")
            arquivo.write(f"Indaiatuba Dúvidas: {self.contador_ind7}\n")
            arquivo.write(f"Indaiatuba Outros: {self.contador_ind8}\n")

            arquivo.write(f"GKN Cadastro_Omnilink: {self.contador_gkn1}\n")
            arquivo.write(f"GKN Verificar_Espelhamento: {self.contador_gkn2}\n")
            arquivo.write(f"GKN Trocar_Rastreador: {self.contador_gkn3}\n")
            arquivo.write(f"GKN Voltar_PNA: {self.contador_gkn4}\n")
            arquivo.write(f"GKN Treinamento: {self.contador_gkn5}\n")
            arquivo.write(f"GKN Cadastro: {self.contador_gkn6}\n")
            arquivo.write(f"GKN Dúvidas: {self.contador_gkn7}\n")
            arquivo.write(f"GKN Outros: {self.contador_gkn8}\n")

            arquivo.write(f"Coopercarga Cadastro_Omnilink: {self.contador_coopercarga1}\n")
            arquivo.write(f"Coopercarga Verificar_Espelhamento: {self.contador_coopercarga2}\n")
            arquivo.write(f"Coopercarga Trocar_Rastreador: {self.contador_coopercarga3}\n")
            arquivo.write(f"Coopercarga Voltar_PNA: {self.contador_coopercarga4}\n")
            arquivo.write(f"Coopercarga Treinamento: {self.contador_coopercarga5}\n")
            arquivo.write(f"Coopercarga Cadastro: {self.contador_coopercarga6}\n")
            arquivo.write(f"Coopercarga Dúvidas: {self.contador_coopercarga7}\n")
            arquivo.write(f"Coopercarga Outros: {self.contador_coopercarga8}\n")


    def carregar_ultima_sessao(self):
        
        diretorio_script = os.path.dirname(os.path.realpath(__file__))
        
        arquivo_ultima_sessao = os.path.join(diretorio_script, "ultima_sessao.txt")
        
        try:
            # Carregar contagens de um arquivo de texto
            with open(arquivo_ultima_sessao, 'r') as file:
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
                        elif empresa == "Klabin":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_klabin1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_klabin2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_klabin3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_klabin4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_klabin5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_klabin6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_klabin7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_klabin8 = int(contador) if contador != "None" else None
                        elif empresa == "Irani":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_irani1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_irani2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_irani3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_irani4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_irani5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_irani6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_irani7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_irani8 = int(contador) if contador != "None" else None
                        elif empresa == "RFR":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_rfr1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_rfr2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_rfr3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_rfr4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_rfr5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_rfr6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_rfr7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_rfr8 = int(contador) if contador != "None" else None
                        elif empresa == "Indaiatuba":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_ind1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_ind2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_ind3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_ind4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_ind5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_ind6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_ind7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_ind8 = int(contador) if contador != "None" else None
                        elif empresa == "GKN":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_gkn1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_gkn2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_gkn3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_gkn4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_gkn5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_gkn6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_gkn7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_gkn8 = int(contador) if contador != "None" else None
                        elif empresa == "Coopercarga":
                            if atividade == "Cadastro_Omnilink:":
                                self.contador_coopercarga1 = int(contador) if contador != "None" else None
                            elif atividade == "Verificar_Espelhamento:":
                                self.contador_coopercarga2 = int(contador) if contador != "None" else None
                            elif atividade == "Trocar_Rastreador:":
                                self.contador_coopercarga3 = int(contador) if contador != "None" else None
                            elif atividade == "Voltar_PNA:":
                                self.contador_coopercarga4 = int(contador) if contador != "None" else None
                            elif atividade == "Treinamento:":
                                self.contador_coopercarga5 = int(contador) if contador != "None" else None
                            elif atividade == "Cadastro:":
                                self.contador_coopercarga6 = int(contador) if contador != "None" else None
                            elif atividade == "Dúvidas:":
                                self.contador_coopercarga7 = int(contador) if contador != "None" else None
                            elif atividade == "Outros:":
                                self.contador_coopercarga8 = int(contador) if contador != "None" else None
                        
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
            
            self.atualizar_rotulo(self.label_klabin1, self.contador_klabin1)
            self.atualizar_rotulo(self.label_klabin2, self.contador_klabin2)
            self.atualizar_rotulo(self.label_klabin3, self.contador_klabin3)
            self.atualizar_rotulo(self.label_klabin4, self.contador_klabin4)
            self.atualizar_rotulo(self.label_klabin5, self.contador_klabin5)
            self.atualizar_rotulo(self.label_klabin6, self.contador_klabin6)
            self.atualizar_rotulo(self.label_klabin7, self.contador_klabin7)
            self.atualizar_rotulo(self.label_klabin8, self.contador_klabin8)
            
            self.atualizar_rotulo(self.label_irani1, self.contador_irani1)
            self.atualizar_rotulo(self.label_irani2, self.contador_irani2)
            self.atualizar_rotulo(self.label_irani3, self.contador_irani3)
            self.atualizar_rotulo(self.label_irani4, self.contador_irani4)
            self.atualizar_rotulo(self.label_irani5, self.contador_irani5)
            self.atualizar_rotulo(self.label_irani6, self.contador_irani6)
            self.atualizar_rotulo(self.label_irani7, self.contador_irani7)
            self.atualizar_rotulo(self.label_irani8, self.contador_irani8)

            self.atualizar_rotulo(self.label_rfr1, self.contador_rfr1)
            self.atualizar_rotulo(self.label_rfr2, self.contador_rfr2)
            self.atualizar_rotulo(self.label_rfr3, self.contador_rfr3)
            self.atualizar_rotulo(self.label_rfr4, self.contador_rfr4)
            self.atualizar_rotulo(self.label_rfr5, self.contador_rfr5)
            self.atualizar_rotulo(self.label_rfr6, self.contador_rfr6)
            self.atualizar_rotulo(self.label_rfr7, self.contador_rfr7)
            self.atualizar_rotulo(self.label_rfr8, self.contador_rfr8)

            self.atualizar_rotulo(self.label_ind1, self.contador_ind1)
            self.atualizar_rotulo(self.label_ind2, self.contador_ind2)
            self.atualizar_rotulo(self.label_ind3, self.contador_ind3)
            self.atualizar_rotulo(self.label_ind4, self.contador_ind4)
            self.atualizar_rotulo(self.label_ind5, self.contador_ind5)
            self.atualizar_rotulo(self.label_ind6, self.contador_ind6)
            self.atualizar_rotulo(self.label_ind7, self.contador_ind7)
            self.atualizar_rotulo(self.label_ind8, self.contador_ind8)

            self.atualizar_rotulo(self.label_gkn1, self.contador_gkn1)
            self.atualizar_rotulo(self.label_gkn2, self.contador_gkn2)
            self.atualizar_rotulo(self.label_gkn3, self.contador_gkn3)
            self.atualizar_rotulo(self.label_gkn4, self.contador_gkn4)
            self.atualizar_rotulo(self.label_gkn5, self.contador_gkn5)
            self.atualizar_rotulo(self.label_gkn6, self.contador_gkn6)  
            self.atualizar_rotulo(self.label_gkn7, self.contador_gkn7)
            self.atualizar_rotulo(self.label_gkn8, self.contador_gkn8)

            self.atualizar_rotulo(self.label_coopercarga1, self.contador_coopercarga1)
            self.atualizar_rotulo(self.label_coopercarga2, self.contador_coopercarga2)
            self.atualizar_rotulo(self.label_coopercarga3, self.contador_coopercarga3)
            self.atualizar_rotulo(self.label_coopercarga4, self.contador_coopercarga4)
            self.atualizar_rotulo(self.label_coopercarga5, self.contador_coopercarga5)
            self.atualizar_rotulo(self.label_coopercarga6, self.contador_coopercarga6)
            self.atualizar_rotulo(self.label_coopercarga7, self.contador_coopercarga7)
            self.atualizar_rotulo(self.label_coopercarga8, self.contador_coopercarga8)

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
            
            self.contador_klabin1 = None
            self.contador_klabin2 = None
            self.contador_klabin3 = None
            self.contador_klabin4 = None
            self.contador_klabin5 = None
            self.contador_klabin6 = None
            self.contador_klabin7 = None
            self.contador_klabin8 = None
            
            self.contador_irani1 = None
            self.contador_irani2 = None
            self.contador_irani3 = None
            self.contador_irani4 = None
            self.contador_irani5 = None
            self.contador_irani6 = None
            self.contador_irani7 = None
            self.contador_irani8 = None

            self.contador_rfr1 = None
            self.contador_rfr2 = None
            self.contador_rfr3 = None
            self.contador_rfr4 = None
            self.contador_rfr5 = None
            self.contador_rfr6 = None
            self.contador_rfr7 = None
            self.contador_rfr8 = None

            self.contador_ind1 = None
            self.contador_ind2 = None
            self.contador_ind3 = None
            self.contador_ind4 = None
            self.contador_ind5 = None
            self.contador_ind6 = None
            self.contador_ind7 = None
            self.contador_ind8 = None

            self.contador_gkn1 = None
            self.contador_gkn2 = None
            self.contador_gkn3 = None
            self.contador_gkn4 = None
            self.contador_gkn5 = None
            self.contador_gkn6 = None
            self.contador_gkn7 = None
            self.contador_gkn8 = None

            self.contador_coopercarga1 = None
            self.contador_coopercarga2 = None
            self.contador_coopercarga3 = None
            self.contador_coopercarga4 = None
            self.contador_coopercarga5 = None
            self.contador_coopercarga6 = None
            self.contador_coopercarga7 = None
            self.contador_coopercarga8 = None
    
    def fechar_janela(self):
        self.salvar_ultima_sessao()
        self.destroy()
    
if __name__ == '__main__':
    app = RegistroAtendimentos()
    app.mainloop()
    