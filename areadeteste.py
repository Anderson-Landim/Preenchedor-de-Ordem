import tkinter as tk

root = tk.Tk()
root.title("Tabela de Teste para Automação")
root.geometry("1200x900")

# Cabeçalho
tk.Label(root, text="Simulador de Tabela (para testes de automação)", font=("Arial", 12, "bold")).pack(pady=5)

# Frame principal
frame = tk.Frame(root)
frame.pack(pady=10)

linhas = 25
colunas = 8
entradas = []

# Função para mover o foco com Enter e Setas
def mover_foco(event):
    widget = event.widget
    for i in range(linhas):
        for j in range(colunas):
            if entradas[i][j] == widget:
                if event.keysym == "Return":  # Enter
                    # Move para próxima célula à direita
                    if j + 1 < colunas:
                        entradas[i][j + 1].focus()
                    else:
                        # Vai pra primeira célula da próxima linha
                        if i + 1 < linhas:
                            entradas[i + 1][0].focus()
                elif event.keysym == "Down":
                    if i + 1 < linhas:
                        entradas[i + 1][j].focus()
                elif event.keysym == "Up":
                    if i - 1 >= 0:
                        entradas[i - 1][j].focus()
                elif event.keysym == "Left":
                    if j - 1 >= 0:
                        entradas[i][j - 1].focus()
                elif event.keysym == "Right":
                    if j + 1 < colunas:
                        entradas[i][j + 1].focus()
                return "break"

# Cabeçalho das colunas
headers = ["Código", "Quantidade", "Lote", "Data", "valor", "un", "PS", "Data/Lote"]
for j, nome in enumerate(headers):
    tk.Label(frame, text=nome, font=("Arial", 10, "bold"), width=15, borderwidth=1, relief="solid").grid(row=0, column=j)

# Cria a tabela de entradas
for i in range(1, linhas + 1):
    linha = []
    for j in range(colunas):
        e = tk.Entry(frame, width=18, justify="center", font=("Arial", 10))
        e.grid(row=i, column=j, padx=3, pady=3)
        e.bind("<Return>", mover_foco)
        e.bind("<Up>", mover_foco)
        e.bind("<Down>", mover_foco)
        e.bind("<Left>", mover_foco)
        e.bind("<Right>", mover_foco)
        linha.append(e)
    entradas.append(linha)

# Botão para limpar tabela
def limpar():
    for linha in entradas:
        for e in linha:
            e.delete(0, tk.END)
    entradas[0][0].focus()

tk.Button(root, text="Limpar Tudo", command=limpar, bg="#d9534f", fg="white", width=15).pack(pady=5)

root.mainloop()
