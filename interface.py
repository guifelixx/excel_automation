import tkinter as tk
from main import rodar

interface = tk.Tk()
interface.title("Automação de planilha")
interface.geometry("300x200")

#Configuração de redimensionamento proporcional
interface.grid_columnconfigure(0, weight=1)
interface.grid_rowconfigure(0, weight=1)
interface.grid_rowconfigure(3, weight=1)  # Adiciona uma linha extra para empurrar o botão para cima

texto_orientacao = tk.Label(interface, text="Aperte o botão para começar a automação:")
texto_orientacao.grid(column=0, row=0, padx=10, pady=10)

botao = tk.Button(interface, text="Começar", command=rodar)
botao.grid(column=0, row=1, padx=10, pady=10)

interface.mainloop()
