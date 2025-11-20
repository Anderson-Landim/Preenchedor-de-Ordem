"""
digitador.py
Tkinter + ttkbootstrap app com 3 abas (CRUZILIA, BÚFALA, SORO)
cada aba usa um JSON separado:
 - cruzilia.json
 - bufala.json
 - soro.json

Fluxo de automação (por item):
1- digitar o codigo
2- apertar ENTER
3- pressionar RIGHT 4x
4- apertar ENTER
5- digitar quantidade (do JSON, ou override por 'Quantidade fixa')
6- pressionar DOWN
7- repetir

Botão Iniciar dispara apenas os itens da aba selecionada.
"""
from ctypes import (
    Structure, sizeof, c_int, c_uint, c_void_p,
    windll, byref, addressof
)

import json
import threading
import time
from pathlib import Path
import tkinter as tk
import openpyxl
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
try:
    import pyautogui
except Exception:
    pyautogui = None


# --- configurações de arquivos JSON ---
BASE_DIR = Path(".")
JSON_FILES = {
    "CRUZILIA": BASE_DIR / "cruzilia.json",
    "BÚFALA": BASE_DIR / "bufala.json",
    "SORO": BASE_DIR / "soro.json",
}

# cria arquivos JSON vazios se não existirem (com [] por padrão)
for p in JSON_FILES.values():
    if not p.exists():
        p.write_text("[]", encoding="utf-8")

# ======= EFEITO VIDRO / ACRYLIC =======

class ACCENT_POLICY(Structure):
    _fields_ = [
        ("AccentState", c_int),
        ("AccentFlags", c_int),
        ("GradientColor", c_int),
        ("AnimationId", c_int)
    ]


class WINCOMPATTRDATA(Structure):
    _fields_ = [
        ("Attribute", c_int),
        ("Data", c_void_p),
        ("SizeOfData", c_uint)
    ]


def enable_acrylic(hwnd):
    accent = ACCENT_POLICY()
    accent.AccentState = 4  # Acrylic Blur
    accent.GradientColor = 0x99FFFFFF  # Transparência + cor (ARGB)

    data = WINCOMPATTRDATA()
    data.Attribute = 19  # WCA_ACCENT_POLICY
    data.Data = c_void_p(addressof(accent))
    data.SizeOfData = sizeof(accent)

    windll.user32.SetWindowCompositionAttribute(hwnd, byref(data))


def disable_acrylic(hwnd):
    accent = ACCENT_POLICY()
    accent.AccentState = 0  # Desativa

    data = WINCOMPATTRDATA()
    data.Attribute = 19
    data.Data = c_void_p(addressof(accent))
    data.SizeOfData = sizeof(accent)

    windll.user32.SetWindowCompositionAttribute(hwnd, byref(data))

class CodeStore:
    """Armazena e manipula lista de (codigo, nome, quantidade) a partir de um JSON específico."""
    def __init__(self, path: Path):
        self.path = Path(path)
        self.data = []
        self.load()

    def load(self):
        try:
            with self.path.open("r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception:
            raw = []
        cleaned = []
        for item in raw:
            if isinstance(item, dict):
                cod = str(item.get("codigo", "")).strip()
                nome = str(item.get("nome", "")).strip()
                qtd = str(item.get("quantidade", "100000")).strip() or "100000"
            elif isinstance(item, (list, tuple)):
                cod = str(item[0]).strip()
                nome = str(item[1]).strip() if len(item) > 1 else ""
                qtd = str(item[2]).strip() if len(item) > 2 else "100000"
            else:
                continue
            if cod:
                cleaned.append((cod, nome, qtd))
        self.data = cleaned

    def save(self):
        out = [{"codigo": c, "nome": n, "quantidade": q} for c, n, q in self.data]
        with self.path.open("w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    def get_all(self):
        return list(self.data)

    def add(self, codigo, nome, qtd="100000"):
        self.data.append((codigo, nome, qtd))
        self.save()

    def edit(self, idx, codigo, nome, qtd):
        if 0 <= idx < len(self.data):
            self.data[idx] = (codigo, nome, qtd)
            self.save()

    def delete(self, idx):
        if 0 <= idx < len(self.data):
            del self.data[idx]
            self.save()


class TabFrame(tb.Frame):
    """Frame que contém a lista e botões para cada aba."""
    def __init__(self, master, name, json_path):
        super().__init__(master)
        self.name = name
        self.store = CodeStore(json_path)
        self.cards = []
        self._build_ui()
    
    def _import_excel(self):
        messagebox.showinfo(
            "Instruções",
            "O arquivo Excel vai preencher a aba atual.\n"
            "O arquivo Excel deve ter exatamente 3 colunas:\n"
            "1) Código\n"
            "2) Item (descrição)\n"
            "3) Quantidade\n\n"
            "❗ Não precisa ter cabeçalho."
        )

        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active

            novos_itens = []
            for row in sheet.iter_rows(values_only=True):
                if not row or len(row) < 3:
                    continue

                codigo = str(row[0]).strip()
                nome = str(row[1]).strip()
                qtd = str(row[2]).strip()

                if codigo:
                    novos_itens.append((codigo, nome, qtd))

            if not novos_itens:
                messagebox.showwarning("Sem dados", "Nenhuma linha válida encontrada no Excel.")
                return

            # Adiciona os itens ao JSON atual
            for cod, nome, qtd in novos_itens:
                self.store.add(cod, nome, qtd)

            self._update_cards()

            messagebox.showinfo(
                "Importação concluída",
                f"{len(novos_itens)} itens foram importados com sucesso!"
            )

        except Exception as e:
            messagebox.showerror("Erro ao importar", f"Ocorreu um erro ao importar o Excel:\n{e}")


    def _build_ui(self):
        ctrl = tb.Frame(self)
        ctrl.pack(fill=X, pady=(8,12), padx=12)
     



        tb.Label(self, text=f"🗂️ {self.name}", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(6,4))

        bar = tb.Frame(self)
        bar.pack(fill=X, pady=6)
        tb.Button(bar, text="➕ Adicionar", bootstyle="success", command=self._add_item).pack(side="left", padx=6)
        tb.Button(bar, text="🔄 Atualizar", bootstyle="secondary", command=self._reload).pack(side="left", padx=6)
        tb.Button(bar, text="📥 Importar Excel", bootstyle="info", command=self._import_excel).pack(side="left", padx=6)



        container = tb.Frame(self)
        container.pack(fill=BOTH, expand=True, padx=4, pady=(6,8))

        self.canvas = tk.Canvas(container, highlightthickness=0, height=380)
        self.scroll_frame = tb.Frame(self.canvas)

        vsb = tb.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # 🔹 Salva o ID da janela criada dentro do Canvas
        # Salva o ID da janela
        self.window_id = self.canvas.create_window(
            (0, 0),
            window=self.scroll_frame,
            anchor="nw"
        )


        # 🔹 Atualiza o scrollregion normalmente
        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # 🔹 Ajusta a largura automaticamente para ocupar 100%
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.window_id, width=e.width)
        )


        self._update_cards()

    def _reload(self):
        self.store.load()
        self._update_cards()

    def _update_cards(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.cards.clear()
        for idx, (codigo, nome, qtd) in enumerate(self.store.get_all()):
            card = tb.Frame(self.scroll_frame, padding=8, bootstyle="dark")
            card.pack(fill=X, pady=4, padx=4)
            self.cards.append(card)
            left = tb.Frame(card)
            left.pack(side="left", fill="both", expand=True)
            tb.Label(left, text=f"Código: {codigo}", font=("Consolas", 11, "bold")).pack(anchor="w")
            tb.Label(left, text=f"{nome}", font=("Segoe UI", 10)).pack(anchor="w")
            tb.Label(left, text=f"Quantidade: {qtd}", font=("Segoe UI", 9, "italic")).pack(anchor="w")

            right = tb.Frame(card)
            right.pack(side="right")
            tb.Button(right, text="Editar", bootstyle="info-outline", width=9, command=lambda i=idx: self._edit_item(i)).pack(side="top", pady=2)
            tb.Button(right, text="Excluir", bootstyle="danger-outline", width=9, command=lambda i=idx: self._delete_item(i)).pack(side="top", pady=2)

    def _add_item(self):
        codigo = tb.dialogs.Querybox.get_string("Digite o código:", "Adicionar novo código")
        if not codigo:
            return
        nome = tb.dialogs.Querybox.get_string("Digite o nome/descritivo:", "Adicionar novo código") or ""
        qtd = tb.dialogs.Querybox.get_string("Digite a quantidade (padrão 100000):", "Adicionar novo código") or "100000"
        self.store.add(codigo.strip(), nome.strip(), qtd.strip())
        self._update_cards()

    def _edit_item(self, idx):
        codigo, nome, qtd = self.store.get_all()[idx]
        novo_codigo = tb.dialogs.Querybox.get_string("Editar código:", initialvalue=codigo)
        if not novo_codigo:
            return
        novo_nome = tb.dialogs.Querybox.get_string("Editar descrição:", initialvalue=nome) or ""
        nova_qtd = tb.dialogs.Querybox.get_string("Editar quantidade:", initialvalue=qtd) or "100000"
        self.store.edit(idx, novo_codigo.strip(), novo_nome.strip(), nova_qtd.strip())
        self._update_cards()

    def _delete_item(self, idx):
        codigo, nome, _ = self.store.get_all()[idx]
        if messagebox.askyesno("Confirmar exclusão", f"Deseja remover {codigo} - {nome}?"):
            self.store.delete(idx)
            self._update_cards()

    def get_items(self):
        """Retorna lista de (codigo, nome, quantidade) atual, aplicando override se preenchido."""
        items = self.store.get_all()
      
        return items

    def highlight_card(self, idx, style="success"):
        """destaca um card visualmente (usa classes do ttkbootstrap)"""
        # simplistic visual feedback: change background of the card
        try:
            if 0 <= idx < len(self.cards):
                widget = self.cards[idx]
                # configure background manually
                bg = {"active": "#1e88e5", "done": "#2e7d32", "err": "#b00020"}.get(style, "#1e88e5")
                widget.configure(style="secondary.TFrame")
                for child in widget.winfo_children():
                    try:
                        child.configure(background=bg)
                    except Exception:
                        pass
        except Exception:
            pass

    def scroll_to(self, idx):
        if 0 <= idx < len(self.cards):
            self.canvas.yview_moveto(idx / max(1, len(self.cards)))


class AutoTyperApp(tb.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("DIGITADOR DE ORDEM")
        self.geometry("512x720")
        self.stop_event = threading.Event()
        self._build_ui()

    def _build_ui(self):

        # notebook com 3 abasself.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        self.notebook = tb.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True, padx=12, pady=6)

        self.tabs = {}
        for name, path in JSON_FILES.items():
            frame = TabFrame(self.notebook, name, path)
            self.tabs[name] = frame
            self.notebook.add(frame, text=name)

        ctrl = tb.Frame(self)
        ctrl.pack(fill=X, pady=(8,12), padx=12)

        tb.Button(ctrl, text="▶️ Iniciar", bootstyle="primary", command=self._start, width=18).pack(side="left", padx=6)
        tb.Button(ctrl, text="⏹ Parar", bootstyle="danger", command=self._stop, width=12).pack(side="left", padx=6)

        # Botão global de vidro
        self.glass_on = False

        def toggle_glass():
            self.glass_on = not self.glass_on
            hwnd = self.winfo_id()

            if self.glass_on:
                enable_acrylic(hwnd)
                self.attributes("-topmost", True)
                btn_glass.configure(text="Fixar: ON", bootstyle="info")
            else:
                disable_acrylic(hwnd)
                self.attributes("-topmost", False)
                btn_glass.configure(text="Fixar: OFF", bootstyle="secondary")

        btn_glass = tb.Button(
            ctrl,
            text="Fixar: OFF",
            bootstyle="secondary",
            command=toggle_glass,
            width=14
        )
        btn_glass.pack(side="left", padx=6)



        self.status = tk.StringVar(value="Pronto")
        tb.Label(self, textvariable=self.status, anchor="w", bootstyle="secondary").pack(fill="x", padx=12, pady=(6,10))
    

    def _current_tab(self):
        cur = self.notebook.select()
        for name, frame in self.tabs.items():
            if str(frame) == str(cur) or self.notebook.index("current") == list(self.tabs.keys()).index(name):
                return frame
        # fallback
        idx = self.notebook.index("current")
        name = list(self.tabs.keys())[idx]
        return self.tabs[name]

    def _start(self):
        if pyautogui is None:
            messagebox.showerror("Erro", "pyautogui não está instalado. Rode: pip install pyautogui")
            return
        tab = self._current_tab()
        items = tab.get_items()
        if not items:
            messagebox.showwarning("Aviso", "Lista vazia para a aba selecionada.")
            return

        self.stop_event.clear()
        self.status.set("Iniciando em 3 segundos... Posicione o cursor no campo alvo.")
        threading.Thread(target=self._worker, args=(tab, items), daemon=True).start()

    def _stop(self):
        self.stop_event.set()
        self.status.set("Parando...")

    def _worker(self, tab: TabFrame, items):
        # Delay pra dar tempo de foco
        time.sleep(3)
        pyautogui.FAILSAFE = True
        try:
            for idx, (codigo, nome, qtd) in enumerate(items):
                if self.stop_event.is_set():
                    self.status.set("Parado pelo usuário.")
                    break
                self.status.set(f"[{idx+1}/{len(items)}] Digitando: {codigo} (Qtd: {qtd})")
                
                # UI feedback
                try:
                    self.after(0, lambda i=idx: tab.scroll_to(i))
                except Exception:
                    pass

                # 1) digitar o código
                pyautogui.typewrite(str(codigo))
                time.sleep(0.06)

                # 2) apertar ENTER
                pyautogui.press("enter")
                time.sleep(0.08)

                # 3) seta para a direita 4x
                for _ in range(4):
                    pyautogui.press("right")
                    time.sleep(0.04)

                # 4) apertar ENTER
                pyautogui.press("enter")
                time.sleep(0.08)

                # 5) digitar quantidade
                pyautogui.typewrite(str(qtd))
                time.sleep(0.06)

                # 6) seta para baixo
                pyautogui.press("down")
                time.sleep(0.5)

                # visual: destacar card (não remove cor depois)
                self.after(0, lambda i=idx: tab.highlight_card(i, style="done"))

                # pequeno intervalo entre itens (ajustável)
                time.sleep(0.18)

            else:
                # loop completo sem break
                self.status.set("✅ Concluído com sucesso.")
        except pyautogui.FailSafeException:
            self.status.set("Abortado: Fail-safe acionado (mova o mouse para um canto).")
        except Exception as e:
            self.status.set(f"Erro durante execução: {e}")
        finally:
            self.stop_event.clear()


if __name__ == "__main__":
    app = AutoTyperApp()
    app.mainloop()
