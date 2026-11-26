"""Identidade visual: tema escuro azul-ardósia com um único destaque azul.

Menu, conteúdo e registro em tons de azul-marinho; cartões um pouco mais
claros que o fundo. Todas as seções usam o mesmo azul (sem cores por seção).
Bahnschrift nos títulos, Segoe UI no texto e Consolas no registro. Tema único.
"""
import tkinter as tk

import customtkinter as ctk

FUNDO = "#111827"
PAINEL = "#1E293B"
LINHA = "#334155"
TINTA = "#E5E7EB"
CHUMBO = "#94A3B8"
REALCE = "#273449"  # hover neutro
ERRO = "#DC2626"
ERRO_HOVER = "#B91C1C"
AVISO = "#F59E0B"
OK = "#22C55E"

# Menu lateral e coluna do registro
MENU_FUNDO = "#0F172A"
MENU_ATIVO = "#1E3A5F"
MENU_TEXTO = "#CBD5E1"
MENU_APOIO = "#64748B"
REG_FUNDO = "#0F172A"
REG_TEXTO = "#E2E8F0"
REG_APOIO = "#94A3B8"
REG_LINK = "#60A5FA"

CORES_LOG = {"erro": "#F87171", "sucesso": "#4ADE80", "aviso": "#FBBF24"}

# chave: (cor, hover, texto sobre a cor, cor de texto legível no fundo, fundo suave)
# Todas as seções compartilham o mesmo azul; as chaves ficam para não mexer em app.py.
_AZUL = ("#2E6DB4", "#24578F", "#FFFFFF", "#60A5FA", "#1E3A5F")
PALETA = {k: _AZUL for k in ("azul", "roxo", "amarelo", "verde_agua", "magenta", "laranja", "verde")}
DESTAQUE = PALETA["azul"][0]

TITULO_APP = ("Bahnschrift SemiBold", 19)
TITULO_SECAO = ("Bahnschrift SemiBold", 26)
TITULO_GRUPO = ("Bahnschrift SemiBold", 15)
MENU = ("Bahnschrift", 13)
MENU_ATIVO_FONTE = ("Bahnschrift SemiBold", 13)
TEXTO = ("Segoe UI", 11)
TEXTO_FORTE = ("Segoe UI Semibold", 11)
APOIO = ("Segoe UI", 10)
MONO = ("Consolas", 10)

# Cor da seção sendo montada; os componentes abaixo a usam ao serem criados.
_atual = PALETA["azul"]


def usar(chave):
    global _atual
    _atual = PALETA[chave]


def cor():
    return _atual[0]


def aplicar():
    """Chamar antes de criar a janela: os widgets leem o tema na criação."""
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    ctk.set_widget_scaling(1.1)
    t = ctk.ThemeManager.theme

    def c(valor):
        return [valor, valor]

    t["CTk"]["fg_color"] = c(FUNDO)
    t["CTkToplevel"]["fg_color"] = c(FUNDO)
    t["CTkFrame"].update(fg_color=c(PAINEL), top_fg_color=c(PAINEL), border_color=c(LINHA))
    t["CTkLabel"]["text_color"] = c(TINTA)
    t["CTkButton"].update(fg_color=c(DESTAQUE), hover_color=c(PALETA["azul"][1]), text_color=c("#FFFFFF"),
                          border_color=c(LINHA), corner_radius=4)
    t["CTkEntry"].update(fg_color=c(PAINEL), border_color=c(LINHA), text_color=c(TINTA),
                         placeholder_text_color=c("#64748B"), corner_radius=4, border_width=1)
    t["CTkComboBox"].update(fg_color=c(PAINEL), border_color=c(LINHA), button_color=c(LINHA),
                            button_hover_color=c(CHUMBO), text_color=c(TINTA), corner_radius=4, border_width=1)
    t["CTkCheckBox"].update(fg_color=c(DESTAQUE), hover_color=c(PALETA["azul"][1]), border_color=c(CHUMBO),
                            text_color=c(TINTA), corner_radius=3)
    t["CTkRadioButton"].update(fg_color=c(DESTAQUE), hover_color=c(PALETA["azul"][1]), border_color=c(CHUMBO),
                               text_color=c(TINTA))
    t["CTkProgressBar"].update(fg_color=c(LINHA), progress_color=c(DESTAQUE), corner_radius=0)
    t["CTkTextbox"].update(fg_color=c(PAINEL), border_color=c(LINHA), text_color=c(TINTA), corner_radius=4)
    t["CTkScrollbar"].update(button_color=c(LINHA), button_hover_color=c(CHUMBO))
    t["DropdownMenu"].update(fg_color=c(PAINEL), hover_color=c(REALCE), text_color=c(TINTA))


def fio(pai, vertical=False, cor_linha=LINHA):
    """Linha de 1px (um CTkFrame de 1px não é desenhado)."""
    return tk.Frame(pai, bd=0, highlightthickness=0, bg=cor_linha, **({"width": 1} if vertical else {"height": 1}))


def titulo_grupo(pai, texto, primeiro=False):
    linha = ctk.CTkFrame(pai, fg_color="transparent")
    linha.pack(anchor="w", fill="x", padx=2, pady=(0 if primeiro else 22, 4))
    ctk.CTkFrame(linha, width=4, height=18, corner_radius=0, fg_color=_atual[0]).pack(side="left", padx=(0, 8))
    rotulo = ctk.CTkLabel(linha, text=texto, font=TITULO_GRUPO, text_color=_atual[3], anchor="w")
    rotulo.pack(side="left")
    return rotulo


def apoio(pai, texto, largura=560):
    rotulo = ctk.CTkLabel(
        pai, text=texto, font=APOIO, text_color=CHUMBO, wraplength=largura, justify="left", anchor="w"
    )
    rotulo.pack(anchor="w", fill="x", padx=2, pady=(0, 8))
    return rotulo


def entrada(pai, dica, **opcoes):
    return ctk.CTkEntry(pai, placeholder_text=dica, height=34, font=TEXTO, **opcoes)


def botao_primario(pai, texto, comando, **opcoes):
    return ctk.CTkButton(
        pai, text=texto, command=comando, height=34, corner_radius=4, font=TEXTO_FORTE,
        fg_color=_atual[0], hover_color=_atual[1], text_color=_atual[2], **opcoes
    )


def botao_secundario(pai, texto, comando, **opcoes):
    return ctk.CTkButton(
        pai, text=texto, command=comando, height=34, corner_radius=4, font=TEXTO_FORTE,
        fg_color=PAINEL, hover_color=_atual[4], border_width=2, border_color=_atual[0],
        text_color=_atual[3], **opcoes
    )


def botao_perigo(pai, texto, comando, **opcoes):
    return ctk.CTkButton(
        pai, text=texto, command=comando, height=34, corner_radius=4, font=TEXTO_FORTE,
        fg_color=ERRO, hover_color=ERRO_HOVER, text_color="#FFFFFF", **opcoes
    )


def botao_link(pai, texto, comando, **opcoes):
    opcoes.setdefault("text_color", _atual[3])
    opcoes.setdefault("hover_color", _atual[4])
    return ctk.CTkButton(
        pai, text=texto, command=comando, height=26, corner_radius=4, font=TEXTO_FORTE,
        fg_color="transparent", **opcoes
    )


def linha_comando(pai, texto, comando):
    """Item de lista de ferramentas: texto à esquerda, fundo na cor suave da seção ao passar o mouse."""
    if pai.winfo_children():
        fio(pai).pack(fill="x", padx=12)
    botao = ctk.CTkButton(
        pai, text=texto, command=comando, anchor="w", height=32, corner_radius=4, font=TEXTO,
        fg_color="transparent", hover_color=REALCE, text_color=TINTA,
    )
    botao.pack(fill="x", padx=4, pady=2)
    return botao


def lista(pai):
    """Moldura branca com uma faixa da cor da seção à esquerda. Devolve a área interna."""
    moldura = ctk.CTkFrame(pai, fg_color=PAINEL, corner_radius=4, border_width=1, border_color=LINHA)
    moldura.pack(fill="x", padx=2)
    # height=1: sem ele o CTkFrame nasce com 200px e estica a lista inteira; fill="y" acompanha o conteúdo.
    ctk.CTkFrame(moldura, width=3, height=1, corner_radius=0, fg_color=_atual[0]).pack(side="left", fill="y", padx=(1, 0), pady=1)
    interna = ctk.CTkFrame(moldura, fg_color="transparent", corner_radius=0)
    interna.pack(side="left", fill="x", expand=True, padx=3, pady=3)
    return interna


def campo(pai, rotulo, dica="", largura=None, **opcoes):
    """Rótulo visível acima do campo; o placeholder fica só como exemplo."""
    ctk.CTkLabel(pai, text=rotulo, font=APOIO, text_color=CHUMBO, anchor="w").pack(anchor="w", padx=2, pady=(6, 1))
    caixa = entrada(pai, dica, **({"width": largura} if largura else {}), **opcoes)
    caixa.pack(anchor="w", fill=None if largura else "x", padx=2)
    return caixa


def pagina(pai):
    """Área rolável de uma seção, sem moldura."""
    frame = ctk.CTkScrollableFrame(pai, fg_color="transparent", corner_radius=0)
    frame.pack(fill="both", expand=True)
    return frame


def lista_marcavel(pai, titulo, altura=180):
    return ctk.CTkScrollableFrame(
        pai, label_text=titulo, label_font=TEXTO_FORTE, label_fg_color=_atual[0], label_text_color=_atual[2],
        height=altura, corner_radius=4, border_width=1, border_color=LINHA, fg_color=PAINEL
    )


def marcavel(pai, texto, variavel):
    return ctk.CTkCheckBox(
        pai, text=texto, variable=variavel, font=TEXTO, text_color=TINTA, corner_radius=3,
        border_width=2, fg_color=_atual[0], hover_color=_atual[1], border_color=CHUMBO
    )


def vazio(pai, texto):
    rotulo = ctk.CTkLabel(pai, text=texto, font=APOIO, text_color=CHUMBO, wraplength=600, justify="left")
    rotulo.pack(anchor="w", padx=8, pady=14)
    return rotulo
