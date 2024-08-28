import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread
from main import Automatizacao

def submeter():
    try:
        file_path = entry_file.get()
        sheet_name = entry_sheet.get()
        
        thread = Thread(target=Automatizacao, args=(file_path, sheet_name))
        thread.start()
        messagebox.showinfo("O programa ocorreu bem, vc pode buscar a planilha nova em\nFundos Organizados")

        
    except Exception as e:
        print(f"Erro ao submeter: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro ao submeter: {e}")
        
def selecionar_file():
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*xlsx")])
    entry_file.delete(0,tk.END)
    entry_file.insert(0,file_path)
    
    
root = tk.Tk()
root.title("Tela de automacao")
root.geometry("800x600")
root.configure(bg="white")

tk.Label(root, text="Nome da página do Excel", bg="white").pack(pady=(10,0))
entry_sheet = tk.Entry(root, width=50)
entry_sheet.pack(pady=(10,0))

tk.Label(root, text="Submeta o arquivo Excel", bg="white").pack(pady=(10,0))
entry_file= tk.Entry(root, width=50)
entry_file.pack(pady=(0, 5))

btn_send= tk.Button(root, text="Selecionar arquivo Excel", command=selecionar_file)
btn_send.pack(pady=(0,10))

tk.Button(root, text="Submeter", command=submeter).pack(pady=(10,10))

root.mainloop()