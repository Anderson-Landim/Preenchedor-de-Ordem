📘 DIGITADOR DE ORDEM – README
🖥️ Sobre o projeto

O Digitador de Ordem é uma ferramenta em Python + Tkinter + ttkbootstrap criada para automatizar o preenchimento de códigos e quantidades em sistemas industriais.
O software lê listas de itens a partir de arquivos JSON ou Excel e simula o processo manual de digitação usando pyautogui.

O objetivo é reduzir tempo, erros manuais e repetição, permitindo que o operador apenas selecione a aba desejada e clique em Iniciar.

🚀 Funcionalidades
✔ Três abas independentes

CRUZILIA

BÚFALA

SORO

Cada aba possui seu próprio arquivo JSON:

cruzilia.json
bufala.json
soro.json

✔ Automação completa via PyAutoGUI

Para cada item:

Digita o código

Pressiona ENTER

Move 4× para a direita

Pressiona ENTER

Digita a quantidade

Move para baixo

Vai para o próximo item automaticamente

✔ Importação de arquivos Excel

Aceita .xlsx e .xls

Deve conter 3 colunas (sem cabeçalho)

Código

Nome / descrição

Quantidade

✔ Controle visual

Cards para cada item

Atualização dinâmica

Destaque automático do item sendo digitado

Barra inferior mostrando o status atual

✔ Botão global “Fixar” (Acrylic / Vidro)

Aplica efeito acrílico (blur transparente)

Mantém a janela sempre no topo (topmost)

Funciona em qualquer aba

ON/OFF sincronizado
