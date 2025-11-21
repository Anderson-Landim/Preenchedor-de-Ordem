# digitador.py
"""
digitador.py
Tkinter + ttkbootstrap app com categorias e sub-abas dinâmicas.
- Cada Categoria tem seu próprio notebook interno com sub-abas (cada sub-aba = TabFrame com JSON)
- Configuração salva em config_abas.json
- Dados das sub-abas em data/<categoria>/<subaba>.json
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
from tkinter import messagebox, filedialog, simpledialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
try:
    import pyautogui
except Exception:
    pyautogui = None

# --- configurações de arquivos JSON ---
BASE_DIR = Path(".")
CONFIG_FILE = BASE_DIR / "config_abas.json"
DATA_DIR = BASE_DIR / "data"

# garante pasta data
DATA_DIR.mkdir(exist_ok=True)

# ==== helpers para path / nomes ====

def safe_filename(name: str) -> str:
    # cria um nome de arquivo simples a partir do nome da aba
    s = name.strip().lower().replace(" ", "_")
    # remove caracteres não-alfanum/underscore
    s = "".join(ch for ch in s if ch.isalnum() or ch == "_")
    if not s:
        s = "tab"
    return s + ".json"

# cria config default se não existir (mantive as 3 abas iniciais como exemplo)
if not CONFIG_FILE.exists():
    default = {
        "Produtos em Processo": [
            {"name": "CRUZILIA", "file": str((DATA_DIR / "produtos_em_processo_cruzilia.json").resolve())},
            {"name": "BÚFALA", "file": str((DATA_DIR / "produtos_em_processo_bufala.json").resolve())},
            {"name": "SORO", "file": str((DATA_DIR / "produtos_em_processo_soro.json").resolve())}
        ]
    }
    CONFIG_FILE.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
    # create empty json files
    for cat, tabs in default.items():
        for t in tabs:
            p = Path(t["file"])
            p.parent.mkdir(parents=True, exist_ok=True)
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
    try:
        accent = ACCENT_POLICY()
        accent.AccentState = 4  # Acrylic Blur
        accent.GradientColor = 0x99FFFFFF  # Transparência + cor (ARGB)

        data = WINCOMPATTRDATA()
        data.Attribute = 19  # WCA_ACCENT_POLICY
        data.Data = c_void_p(addressof(accent))
        data.SizeOfData = sizeof(accent)

        windll.user32.SetWindowCompositionAttribute(hwnd, byref(data))
    except Exception:
        pass


def disable_acrylic(hwnd):
    try:
        accent = ACCENT_POLICY()
        accent.AccentState = 0  # Desativa

        data = WINCOMPATTRDATA()
        data.Attribute = 19
        data.Data = c_void_p(addressof(accent))
        data.SizeOfData = sizeof(accent)

        windll.user32.SetWindowCompositionAttribute(hwnd, byref(data))
    except Exception:
        pass

# ------------------ Classe CodeStore (mantida) ------------------

class CodeStore:
    """Armazena e manipula lista de (codigo, nome, quantidade, timer) a partir de um JSON específico."""
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
                timer = str(item.get("timer", "1")).strip()
            elif isinstance(item, (list, tuple)):
                cod = str(item[0]).strip()
                nome = str(item[1]).strip() if len(item) > 1 else ""
                qtd = str(item[2]).strip() if len(item) > 2 else "100000"
                timer = str(item[3]).strip() if len(item) > 3 else "1"
            else:
                continue
            if cod:
                cleaned.append((cod, nome, qtd, timer))
        self.data = cleaned

    def save(self):
        out = [{"codigo": c, "nome": n, "quantidade": q, "timer": t} for c, n, q, t in self.data]
        with self.path.open("w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    def get_all(self):
        return list(self.data)

    def add(self, codigo, nome, qtd="100000", timer="1"):
        self.data.append((codigo, nome, qtd, timer))
        self.save()

    def edit(self, idx, codigo, nome, qtd, timer):
        if 0 <= idx < len(self.data):
            self.data[idx] = (codigo, nome, qtd, timer)
            self.save()

    def delete(self, idx):
        if 0 <= idx < len(self.data):
            del self.data[idx]
            self.save()

# ------------------ TabFrame (mantida) ------------------

class TabFrame(tb.Frame):
    """Frame que contém a lista e botões para cada sub-aba / arquivo JSON."""
    def __init__(self, master, name, json_path):
        super().__init__(master)
        self.name = name
        self.json_path = Path(json_path)
        # garante pasta do arquivo
        self.json_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.json_path.exists():
            self.json_path.write_text("[]", encoding="utf-8")
        self.store = CodeStore(self.json_path)
        self.cards = []
        self._build_ui()

    def _import_excel(self):
        messagebox.showinfo(
            "Instruções",
            "O arquivo Excel vai preencher a aba atual.\n"
            "O arquivo Excel deve ter exatamente 3 colunas:\n"
            "1) Código\n"
            "2) Item (descrição)\n"
            "3) Quantidade\n"
            "4) Tempo padrao 1s por linha\n\n"
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
                timer = str(row[3]).strip() if len(row) > 3 else "1"

                if codigo:
                    novos_itens.append((codigo, nome, qtd, timer))

            if not novos_itens:
                messagebox.showwarning("Sem dados", "Nenhuma linha válida encontrada no Excel.")
                return

            # Adiciona os itens ao JSON atual
            for cod, nome, qtd, timer in novos_itens:
                self.store.add(cod, nome, qtd, timer)

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

        self.window_id = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.window_id, width=e.width))

        self._update_cards()

    def _reload(self):
        self.store.load()
        self._update_cards()

    def _update_cards(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.cards.clear()
        for idx, (codigo, nome, qtd, timer) in enumerate(self.store.get_all()):
            card = tb.Frame(self.scroll_frame, padding=8, bootstyle="dark")
            card.pack(fill=X, pady=4, padx=4)
            self.cards.append(card)
            left = tb.Frame(card)
            left.pack(side="left", fill="both", expand=True)
            tb.Label(left, text=f"Código: {codigo}", font=("Consolas", 11, "bold")).pack(anchor="w")
            tb.Label(left, text=f"{nome}", font=("Segoe UI", 10)).pack(anchor="w")
            tb.Label(left, text=f"Quantidade: {qtd}", font=("Segoe UI", 9, "italic")).pack(anchor="w")
            tb.Label(left, text=f"Tempo: {timer}", font=("Segoe UI", 9, "italic")).pack(anchor="w")

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
        timer = tb.dialogs.Querybox.get_string("Digite o tempo (s):", "Adicionar tempo") or "1"
        self.store.add(codigo.strip(), nome.strip(), qtd.strip(), timer.strip())
        self._update_cards()

    def _edit_item(self, idx):
        codigo, nome, qtd, timer = self.store.get_all()[idx]
        novo_codigo = tb.dialogs.Querybox.get_string("Editar código:", initialvalue=codigo)
        if not novo_codigo:
            return
        novo_nome = tb.dialogs.Querybox.get_string("Editar descrição:", initialvalue=nome) or ""
        nova_qtd = tb.dialogs.Querybox.get_string("Editar quantidade:", initialvalue=qtd) or "100000"
        nova_timer = tb.dialogs.Querybox.get_string("Editar tempo (s):", initialvalue=timer) or "1"

        self.store.edit(idx, novo_codigo.strip(), novo_nome.strip(), nova_qtd.strip(), nova_timer.strip())
        self._update_cards()

    def _delete_item(self, idx):
        codigo, nome, _ = self.store.get_all()[idx]
        if messagebox.askyesno("Confirmar exclusão", f"Deseja remover {codigo} - {nome}?"):
            self.store.delete(idx)
            self._update_cards()

    def get_items(self):
        """Retorna lista de (codigo, nome, quantidade, timer) atual, aplicando override se preenchido."""
        return self.store.get_all()

    def highlight_card(self, idx, style="success"):
        try:
            if 0 <= idx < len(self.cards):
                widget = self.cards[idx]
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

# ------------------ AutoTyperApp com categorias ------------------

class AutoTyperApp(tb.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("DIGITADOR DE ORDEM")
        self.geometry("640x780")
        self.stop_event = threading.Event()

        # estrutura em memória para categorias
        self.categories = {}  # { "Categoria": [ {"name": subname, "file": "/abs/path.json"}, ... ] }
        self.category_notebook_tabs = {}  # mapping de nome->tab widget
        self.sub_notebooks = {}  # mapping category -> sub-notebook widget

        self._load_config()
        self._build_ui()

    # -------- Config load/save ----------
    def _load_config(self):
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as f:
                self.categories = json.load(f)
        except Exception:
            self.categories = {}

    def _save_config(self):
        with CONFIG_FILE.open("w", encoding="utf-8") as f:
            json.dump(self.categories, f, ensure_ascii=False, indent=2)

    # -------- UI ----------
    def _build_ui(self):
        top_ctrl = tb.Frame(self)
        top_ctrl.pack(fill=X, padx=12, pady=(8,4))

        # Botões de controle de categoria
        tb.Button(top_ctrl, text="➕ Categoria", bootstyle="success", command=self._create_category).pack(side="left", padx=6)
        tb.Button(top_ctrl, text="✏️ Renomear", bootstyle="info", command=self._rename_category).pack(side="left", padx=6)
        tb.Button(top_ctrl, text="🗑️ Excluir", bootstyle="danger", command=self._delete_category).pack(side="left", padx=6)

        # Notebook principal de categorias
        self.cat_notebook = tb.Notebook(self)
        self.cat_notebook.pack(fill=BOTH, expand=False, padx=12, pady=(6,8))

        # para cada categoria carregada, cria uma aba
        for cat_name, tabs in self.categories.items():
            self._create_category_tab(cat_name, tabs, select=False)

        # se não houver nenhuma categoria, cria uma vazia para inicial
        if not self.categories:
            self._create_category("Produtos em Processo")

        # área de controle global (Iniciar / Parar / Fixar vidro) e status
        ctrl = tb.Frame(self)
        ctrl.pack(fill=X, pady=(8,4), padx=12)

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
        btn_glass = tb.Button(ctrl, text="Fixar: OFF", bootstyle="secondary", command=toggle_glass, width=14)
        btn_glass.pack(side="left", padx=6)

        # label de status grande
        self.status = tk.StringVar(value="Pronto")
        # ------ CAIXA DE INFORMAÇÃO DESTACADA ------
        self.status = tk.StringVar(value="Pronto")

        self.status_frame = tb.Frame(self, padding=15, bootstyle="info")
        self.status_frame.pack(fill="x", padx=10, pady=(5,10))

        self.status_label = tb.Label(
            self.status_frame,
            textvariable=self.status,
            anchor="center",
            font=("Segoe UI", 15, "bold")
        )
        self.status_label.pack(fill="x")


    # -------- Categoria: criação / renomear / excluir ----------
    def _create_category(self, name=None):
        # cria com dialog se não fornecido name
        if name is None:
            name = simpledialog.askstring("Nova categoria", "Nome da nova categoria:")
            if not name:
                return
            name = name.strip()
        if name in self.categories:
            messagebox.showwarning("Duplicado", "Já existe uma categoria com esse nome.")
            return

        # inicializa com uma sub-aba vazia
        default_file = DATA_DIR / safe_filename(name.replace(" ", "_") + "_default")
        default_file.parent.mkdir(parents=True, exist_ok=True)
        if not default_file.exists():
            default_file.write_text("[]", encoding="utf-8")

        self.categories[name] = [
            {"name": "Nova Aba", "file": str(default_file.resolve())}
        ]
        self._save_config()
        self._create_category_tab(name, self.categories[name], select=True)

    def _create_category_tab(self, cat_name, tabs, select=True):
        # cria uma tab no cat_notebook com um notebook interno
        frame = tb.Frame(self.cat_notebook)
        # cabeçalho com título e botões de gerência das sub-abas
        hdr = tb.Frame(frame)
        hdr.pack(fill=X, pady=(6,4), padx=8)
        tb.Label(hdr, text=f"Categoria: {cat_name}", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(0,8))
        tb.Button(hdr, text="➕ Sub", bootstyle="success", command=lambda c=cat_name: self._create_subtab(c)).pack(side="left", padx=6)
        tb.Button(hdr, text="✏️ Renomear Sub", bootstyle="info", command=lambda c=cat_name: self._rename_subtab(c)).pack(side="left", padx=6)
        tb.Button(hdr, text="🗑️ Excluir Sub", bootstyle="danger", command=lambda c=cat_name: self._delete_subtab(c)).pack(side="left", padx=6)

        # notebook interno de sub-abas
        sub_nb = tb.Notebook(frame)
        sub_nb.pack(fill=BOTH, expand=True, padx=8, pady=(6,8))
        self.sub_notebooks[cat_name] = sub_nb

        # cria cada sub-aba
        for t in tabs:
            name = t.get("name")
            file = t.get("file")
            self._add_subtab_to_notebook(cat_name, sub_nb, name, file)

        # adiciona ao notebook de categorias
        self.cat_notebook.add(frame, text=cat_name)
        self.category_notebook_tabs[cat_name] = frame

        if select:
            # seleciona essa aba
            idx = self.cat_notebook.index("end") - 1
            self.cat_notebook.select(idx)

    def _rename_category(self):
        # renomeia a categoria selecionada
        cur = self.cat_notebook.select()
        if not cur:
            return
        # acha qual categoria é
        for name, frame in self.category_notebook_tabs.items():
            if str(frame) == str(cur):
                old = name
                break
        else:
            return
        novo = simpledialog.askstring("Renomear categoria", "Novo nome:", initialvalue=old)
        if not novo:
            return
        novo = novo.strip()
        if novo == old:
            return
        if novo in self.categories:
            messagebox.showwarning("Duplicado", "Já existe outra categoria com esse nome.")
            return
        # move mapping: copy and delete old key
        self.categories[novo] = self.categories.pop(old)
        # também atualiza os internal mappings
        frame = self.category_notebook_tabs.pop(old)
        self.category_notebook_tabs[novo] = frame
        nb = self.sub_notebooks.pop(old)
        self.sub_notebooks[novo] = nb
        # atualiza tab text
        idx = self.cat_notebook.index(frame)
        self.cat_notebook.tab(idx, text=novo)
        self._save_config()

    def _delete_category(self):
        cur = self.cat_notebook.select()
        if not cur:
            return
        for name, frame in self.category_notebook_tabs.items():
            if str(frame) == str(cur):
                cat = name
                break
        else:
            return
        if not messagebox.askyesno("Confirmar", f"Excluir a categoria '{cat}' e suas sub-abas?"):
            return
        # remove da UI e da config
        frame = self.category_notebook_tabs.pop(cat)
        self.cat_notebook.forget(frame)
        self.sub_notebooks.pop(cat, None)
        self.categories.pop(cat, None)
        self._save_config()

    # -------- Sub-abas dentro de uma categoria ----------
    def _create_subtab(self, category):
        # pede o nome da nova sub-aba
        name = simpledialog.askstring("Nova sub-aba", "Nome da sub-aba:")
        if not name:
            return
        name = name.strip()
        # gera arquivo
        fname = safe_filename(f"{category}_{name}")
        path = DATA_DIR / category
        path.mkdir(parents=True, exist_ok=True)
        file_path = path / fname
        if not file_path.exists():
            file_path.write_text("[]", encoding="utf-8")
        # registra na config
        self.categories.setdefault(category, []).append({"name": name, "file": str(file_path.resolve())})
        self._save_config()
        # adiciona visualmente
        sub_nb = self.sub_notebooks.get(category)
        if sub_nb:
            self._add_subtab_to_notebook(category, sub_nb, name, str(file_path.resolve()))
            # seleciona a nova aba
            sub_nb.select(sub_nb.index("end") - 1)

    def _add_subtab_to_notebook(self, category, sub_nb, name, file):
        tab_frame = TabFrame(sub_nb, name, Path(file))
        sub_nb.add(tab_frame, text=name)

    def _rename_subtab(self, category):
        sub_nb = self.sub_notebooks.get(category)
        if not sub_nb:
            return
        cur = sub_nb.select()
        if not cur:
            return
        # locate index
        idx = sub_nb.index(cur)
        old_name = sub_nb.tab(idx, "text")
        novo = simpledialog.askstring("Renomear sub-aba", "Novo nome:", initialvalue=old_name)
        if not novo:
            return
        novo = novo.strip()
        # update config entry
        entries = self.categories.get(category, [])
        for e in entries:
            if e.get("name") == old_name:
                e["name"] = novo
                break
        # update UI
        sub_nb.tab(idx, text=novo)
        self._save_config()

    def _delete_subtab(self, category):
        sub_nb = self.sub_notebooks.get(category)
        if not sub_nb:
            return

        cur = sub_nb.select()
        if not cur:
            return

        idx = sub_nb.index(cur)
        name = sub_nb.tab(idx, "text")

        if not messagebox.askyesno("Confirmar", f"Excluir a sub-aba '{name}' da categoria '{category}'?\n\n⚠ Isso também irá apagar o arquivo JSON correspondente."):
            return

        # ---- REMOVE DA CONFIG ----
        entries = self.categories.get(category, [])
        removed_file = None

        new_entries = []
        for e in entries:
            if e.get("name") == name:
                removed_file = e.get("file")
            else:
                new_entries.append(e)

        self.categories[category] = new_entries
        self._save_config()

        # ---- EXCLUI O ARQUIVO JSON ----
        if removed_file:
            try:
                p = Path(removed_file)
                if p.exists():
                    p.unlink()  # REMOVE O ARQUIVO
            except Exception as e:
                messagebox.showwarning("Erro", f"Não foi possível remover o arquivo JSON:\n{e}")

        # ---- REMOVE A ABA VISUALMENTE ----
        sub_nb.forget(idx)


    # -------- utilitário para pegar TabFrame atual ----------
    def _current_tabframe(self):
        # retorna o TabFrame atualmente visível (sub-aba selecionada na categoria selecionada)
        cat_sel = self.cat_notebook.select()
        if not cat_sel:
            return None
        # encontra categoria
        for cat_name, frame in self.category_notebook_tabs.items():
            if str(frame) == str(cat_sel):
                sub_nb = self.sub_notebooks.get(cat_name)
                if not sub_nb:
                    return None
                cur = sub_nb.select()
                if not cur:
                    return None
                # cur é o widget path; procuramos seu objeto TabFrame
                for child in sub_nb.winfo_children():
                    if str(child) == str(cur):
                        return child  # esse child é um TabFrame
                # fallback: tab index mapping
                return None
        return None

    # -------- Start / Stop / Worker (reaproveitado, agora usa _current_tabframe) ----------
    def _start(self):
        if pyautogui is None:
            messagebox.showerror("Erro", "pyautogui não está instalado. Rode: pip install pyautogui")
            return
        tabframe = self._current_tabframe()
        if tabframe is None:
            messagebox.showwarning("Aviso", "Selecione uma sub-aba para iniciar.")
            return
        items = tabframe.get_items()
        if not items:
            messagebox.showwarning("Aviso", "Lista vazia para a aba selecionada.")
            return

        self.stop_event.clear()
        
        self.status.set("Iniciando em 4 segundos... Posicione o cursor no campo alvo.")
        threading.Thread(target=self._worker, args=(tabframe, items), daemon=True).start()

    def _stop(self):
        self.stop_event.set()
        self.status.set("Parando...")

    def _worker(self, tab: TabFrame, items):
        # Delay pra dar tempo de foco
        time.sleep(4)
        pyautogui.FAILSAFE = True
        try:
            for idx, item in enumerate(items):
                if len(item) == 4:
                    codigo, nome, qtd, timer = item
                else:
                    codigo, nome, qtd = item
                    timer = 1  # padrão caso não exista no JSON

                if self.stop_event.is_set():
                    self.status.set("Parado pelo usuário.")
                    break
                self.status.set(f"[{idx+1}/{len(items)}] Digitando: {codigo} (Qtd: {qtd})")

                # UI feedback: scroll e destaque
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

                # 🔄 timer visual em tempo real
                t = float(timer)
                # mudar visual do status_label para warning enquanto aguarda (thread-safe via after)
                self.after(0, lambda: self.status_label.configure(bootstyle="warning"))

                if t > 0:
                    elapsed = 0
                    step = 0.1  # atualização a cada 100ms
                    while elapsed < t:
                        if self.stop_event.is_set():
                            return
                        restante = round(t - elapsed, 1)
                        self.status.set(f"[{idx+1}/{len(items)}] Digitando: {codigo} (Qtd: {qtd}) | Aguardando: {restante}s")
                        time.sleep(step)
                        elapsed += step

                # volta a aparência normal
                self.after(0, lambda: self.status_label.configure(bootstyle="info"))

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
