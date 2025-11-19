"""
App Tkinter + Ttkbootstrap com lista persistida em JSON.
- Itens têm código, nome e quantidade
- Quantidade padrão = "10000" se não definida
- Permite editar código, nome e quantidade
- Usa pyautogui para digitar os dados automaticamente
"""

import threading
import time
import json
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *

try:
    import pyautogui
except Exception:
    pyautogui = None


JSON_PATH = Path("dados.json")


class CodeStore:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.data = []
        self.load()

    def load(self):
        """Carrega dados do JSON ou usa padrão."""
        if self.path.exists():
            try:
                with self.path.open("r", encoding="utf-8") as f:
                    raw = json.load(f)
                cleaned = []
                for item in raw:
                    if isinstance(item, dict):
                        cod = str(item.get("codigo", "")).strip()
                        nome = str(item.get("nome", "")).strip()
                        qtd = str(item.get("quantidade", "10000")).strip() or "10000"
                    elif isinstance(item, (list, tuple)):
                        cod = str(item[0]).strip()
                        nome = str(item[1]).strip() if len(item) > 1 else ""
                        qtd = str(item[2]).strip() if len(item) > 2 else "10000"
                    else:
                        continue
                    if cod:
                        cleaned.append((cod, nome, qtd))
                self.data = cleaned
            except Exception as e:
                print(f"[CodeStore] Erro lendo JSON: {e}")
                self.data = []
        else:
            self.data = []
            self.save()

    def save(self):
        """Salva os dados no JSON."""
        try:
            out = [{"codigo": c, "nome": n, "quantidade": q} for c, n, q in self.data]
            with self.path.open("w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[CodeStore] Erro salvando JSON: {e}")

    def get_all(self):
        return self.data

    def add(self, codigo, nome, qtd="10000"):
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


class AutoTyperApp(tb.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("Automação de Pedido de Resíduo")
        self.geometry("600x950")

        self.store = CodeStore(JSON_PATH)
        self.stop_event = threading.Event()
        self.cards = []

        self._build_ui()

    def _build_ui(self):
        title = tb.Label(self, text="📋 Lista de Produtos", font=("Segoe UI", 16, "bold"))
        title.pack(pady=(15, 10))

        container = tb.Frame(self)
        container.pack(fill=BOTH, expand=True, padx=20)

        self.canvas = tk.Canvas(container, highlightthickness=0)
        self.scroll_frame = tb.Frame(self.canvas)
        vsb = tb.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        bar = tb.Frame(self)
        bar.pack(fill=X, pady=10)

        tb.Button(bar, text="➕ Adicionar", bootstyle="success", command=self._add_item, width=15).pack(side="left", padx=10)
        tb.Button(bar, text="▶️ Iniciar", bootstyle="primary", command=self._start, width=15).pack(side="left", padx=10)
        tb.Button(bar, text="⏹ Parar", bootstyle="danger", command=self._stop, width=15).pack(side="left", padx=10)

        self.status = tk.StringVar(value="Pronto")
        tb.Label(self, textvariable=self.status, anchor="w", bootstyle="secondary").pack(fill="x", padx=20, pady=(6, 8))

        self._update_cards()

    def _update_cards(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.cards.clear()

        for idx, (codigo, nome, qtd) in enumerate(self.store.get_all()):
            card = tk.Frame(self.scroll_frame, bg="#212529", padx=10, pady=10)
            card.pack(fill=X, pady=4, padx=5)
            self.cards.append(card)

            left = tb.Frame(card)
            left.pack(side="left", fill="both", expand=True)
            tb.Label(left, text=f"Código: {codigo}", font=("Consolas", 11, "bold")).pack(anchor="w")
            tb.Label(left, text=f"{nome}", font=("Segoe UI", 10)).pack(anchor="w")
            tb.Label(left, text=f"Quantidade: {qtd}", font=("Segoe UI", 9, "italic"), bootstyle="info").pack(anchor="w")

            right = tb.Frame(card)
            right.pack(side="right")
            tb.Button(right, text="Editar", bootstyle="info-outline", width=9,
                      command=lambda i=idx: self._edit_item(i)).pack(side="top", pady=2)
            tb.Button(right, text="Excluir", bootstyle="danger-outline", width=9,
                      command=lambda i=idx: self._delete_item(i)).pack(side="top", pady=2)

    def _add_item(self):
        codigo = tb.dialogs.Querybox.get_string("Digite o código:", "Adicionar novo código")
        if not codigo:
            return
        nome = tb.dialogs.Querybox.get_string("Digite o nome/descritivo:", "Adicionar novo código") or ""
        qtd = tb.dialogs.Querybox.get_string("Digite a quantidade (padrão 10000):", "Adicionar novo código") or "10000"
        self.store.add(codigo.strip(), nome.strip(), qtd.strip())
        self._update_cards()

    def _edit_item(self, idx):
        codigo, nome, qtd = self.store.get_all()[idx]
        novo_codigo = tb.dialogs.Querybox.get_string("Editar código:", initialvalue=codigo)
        if not novo_codigo:
            return
        novo_nome = tb.dialogs.Querybox.get_string("Editar descrição:", initialvalue=nome) or ""
        nova_qtd = tb.dialogs.Querybox.get_string("Editar quantidade:", initialvalue=qtd) or "10000"
        self.store.edit(idx, novo_codigo.strip(), novo_nome.strip(), nova_qtd.strip())
        self._update_cards()

    def _delete_item(self, idx):
        codigo, nome, _ = self.store.get_all()[idx]
        if messagebox.askyesno("Confirmar exclusão", f"Deseja remover {codigo} - {nome}?"):
            self.store.delete(idx)
            self._update_cards()

    def _start(self):
        if pyautogui is None:
            messagebox.showerror("Erro", "pyautogui não está instalado. pip install pyautogui")
            return
        if not self.store.get_all():
            messagebox.showwarning("Aviso", "Nenhum item na lista.")
            return

        self.stop_event.clear()
        self.status.set("Iniciando em 3 segundos...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _stop(self):
        self.stop_event.set()
        self.status.set("Parando...")

    def _worker(self):
        time.sleep(3)
        pyautogui.FAILSAFE = True
        try:
            for idx, (codigo, _, qtd) in enumerate(self.store.get_all()):
                if self.stop_event.is_set():
                    break
                self.status.set(f"[{idx+1}/{len(self.store.get_all())}] Digitando: {codigo} (Qtd {qtd})")

                self.after(0, lambda i=idx: self._highlight_card(i, "#1e88e5"))  # azul ativo
                self.after(100, lambda i=idx: self._scroll_to_card(i))

                pyautogui.typewrite(codigo)
                for _ in range(5):
                    pyautogui.press("enter")
                pyautogui.typewrite(qtd)
                pyautogui.press("down")
                for _ in range(8):
                    pyautogui.press("left")
                pyautogui.press("enter")

                self.after(0, lambda i=idx: self._highlight_card(i, "#2e7d32"))  # verde concluído
                time.sleep(0.6)

            self.status.set("✅ Concluído com sucesso.")
        except pyautogui.FailSafeException:
            self.status.set("Abortado: Fail-safe acionado.")
        except Exception as e:
            self.status.set(f"Erro durante execução: {e}")
        finally:
            self.stop_event.clear()

    def _highlight_card(self, idx, color):
        try:
            if 0 <= idx < len(self.cards):
                self.cards[idx].configure(background=color)
                for child in self.cards[idx].winfo_children():
                    child.configure(background=color)
        except Exception as e:
            print(f"Erro ao destacar card: {e}")

    def _scroll_to_card(self, idx):
        if 0 <= idx < len(self.cards):
            self.canvas.yview_moveto(idx / max(1, len(self.cards)))


if __name__ == "__main__":
    app = AutoTyperApp()
    app.mainloop()
