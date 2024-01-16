import customtkinter as Ctk 
from customtkinter import *
import tkinter as tk

class RegistroAtendimentos(Ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Registro de Atendimentos')
        self.geometry('800x600')
        
        set_default_color_theme('green')
        
        self.frame_titulo = Ctk.CTkFrame(self)
        
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

        self.frame_suzano = Ctk.CTkFrame(self)
        
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

        self.frame_suzano.grid(row=4, column=0, pady=10, padx=10)
        
        self.btn_suzano1.grid(row=0, column=0, pady=10, padx=10)
        self.label_suzano1.grid(row=0, column=1, pady=10, padx=10)
        
        self.btn_suzano2.grid(row=1, column=0, pady=10, padx=10)
        self.label_suzano2.grid(row=1, column=1, pady=10, padx=10)
        
        self.btn_suzano3.grid(row=2, column=0, pady=10, padx=10)
        self.label_suzano3.grid(row=2, column=1, pady=10, padx=10)
        
        self.btn_suzano4.grid(row=3, column=0, pady=10, padx=10)
        self.label_suzano4.grid(row=3, column=1, pady=10, padx=10)
        
        self.btn_suzano5.grid(row=4, column=0, pady=10, padx=10)
        self.label_suzano5.grid(row=4, column=1, pady=10, padx=10)
        
        self.btn_suzano6.grid(row=5, column=0, pady=10, padx=10)
        self.label_suzano6.grid(row=5, column=1, pady=10, padx=10)
        
        self.btn_suzano7.grid(row=6, column=0, pady=10, padx=10)
        self.label_suzano7.grid(row=6, column=1, pady=10, padx=10)
        
        self.btn_suzano8.grid(row=7, column=0, pady=10, padx=10)
        self.label_suzano8.grid(row=7, column=1, pady=10, padx=10)
        
        ## Posicionamento dos botões ##
        
        #Título
        self.frame_titulo.grid(row=0, column=6, pady=10, padx=10)
        self.label_titulo.grid(row=0, column=6, pady=10, padx=10)
        
        #Suzano
        self.label_suzano.grid(row=4, column=0, pady=10, padx=10)
        
        self.btn_suzano1.grid(row=5,column=0, pady=10, padx=10)
        self.label_suzano1.grid(row=5,column=1, pady=10, padx=10)
        
        self.btn_suzano2.grid(row=6,column=0, pady=10, padx=10)
        self.label_suzano2.grid(row=6,column=1, pady=10, padx=10)
        
        self.btn_suzano3.grid(row=7,column=0, pady=10, padx=10)
        self.label_suzano3.grid(row=7,column=1, pady=10, padx=10)
        
        self.btn_suzano4.grid(row=8,column=0, pady=10, padx=10)
        self.label_suzano4.grid(row=8,column=1, pady=10, padx=10)
        
        self.btn_suzano5.grid(row=9,column=0, pady=10, padx=10)
        self.label_suzano5.grid(row=9,column=1, pady=10, padx=10)
        
        self.btn_suzano6.grid(row=10,column=0, pady=10, padx=10)
        self.label_suzano6.grid(row=10,column=1, pady=10, padx=10)
        
        self.btn_suzano7.grid(row=11,column=0, pady=10, padx=10)
        self.label_suzano7.grid(row=11,column=1, pady=10, padx=10)
        
        self.btn_suzano8.grid(row=12,column=0, pady=10, padx=10)
        self.label_suzano8.grid(row=12,column=1, pady=10, padx=10)
    
        #Funções
    def SuzanoContador1(self):
        self.contador_suzano1 += 1
        self.label_suzano1.configure(text=f"{self.contador_suzano1}")
    
    def SuzanoContador2(self):
        self.contador_suzano2 += 1
        self.label_suzano2.configure(text=f"{self.contador_suzano2}")
    
    def SuzanoContador3(self):
        self.contador_suzano3 += 1
        self.label_suzano3.configure(text=f"{self.contador_suzano3}")
    
    def SuzanoContador4(self):
        self.contador_suzano4 += 1
        self.label_suzano4.configure(text=f"{self.contador_suzano4}")
    
    def SuzanoContador5(self):
        self.contador_suzano5 += 1
        self.label_suzano5.configure(text=f"{self.contador_suzano5}")
    
    def SuzanoContador6(self):
        self.contador_suzano6 += 1
        self.label_suzano6.configure(text=f"{self.contador_suzano6}")
    
    def SuzanoContador7(self):
        self.contador_suzano7 += 1
        self.label_suzano7.configure(text=f"{self.contador_suzano7}")
    
    def SuzanoContador8(self):
        self.contador_suzano8 += 1
        self.label_suzano8.configure(text=f"{self.contador_suzano8}")

                  
if __name__ == '__main__':
    app = RegistroAtendimentos()
    app.mainloop()

        
        
        
