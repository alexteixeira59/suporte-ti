"""Janela principal: navegação, visão geral e registro."""

import ctypes
import os
import platform
import shutil
import socket
import subprocess
import threading
import winreg
from datetime import datetime
from tkinter import messagebox

import customtkinter as ctk

from . import tema
from .comum import CREATE_NO_WINDOW, resource_path
from .execucao import ExecucaoMixin
from .manutencao import ManutencaoMixin
from .office import OfficeMixin
from .outlook import OutlookMixin
from .rede import RedeMixin

# (nome da seção, descrição, método que monta a página, grupo no menu, cor do cabo em tema.PALETA)
SECOES = (
    ("Visão geral", "Dados desta máquina e as correções mais usadas.", "_montar_visao_geral", None, "azul"),
    ("Manutenção", "Atualizações, integridade e limpeza do Windows.", "_montar_aba_manutencao", None, "amarelo"),
    ("Rede e domínio", "Conectividade, DNS, domínio e conta local de suporte.", "_montar_aba_rede", None, "verde_agua"),
    ("Impressoras", "Spooler e instalação de impressora por IP.", "_montar_aba_impressoras", None, "magenta"),
    ("Office", "Reinstalação automatizada do Office 2024.", "_montar_aba_office", None, "laranja"),
    ("Conta e PST", "Conta do Outlook, ano e pasta usados no arquivamento.", "_montar_aba_configuracao", "Outlook", "verde"),
    ("Arquivar e-mails", "Conta quantas mensagens há e move as antigas para um PST.", "_montar_aba_execucao", "Outlook", "verde"),
    ("Limpar caixa", "Encontra newsletters e mensagens grandes para enviar à Lixeira.", "_montar_aba_limpeza_email", "Outlook", "verde"),
)
COR_SECAO = {secao[0]: secao[4] for secao in SECOES}

ACOES_FREQUENTES = (
    ("winget", "Atualizar aplicativos", "Atualiza tudo que o Winget gerencia nesta máquina.",
     "acao_winget_upgrade_all", "Atualizar todos os aplicativos gerenciados pelo Winget?"),
    ("sfc", "Verificar arquivos do Windows", "SFC /scannow. Pede permissão de administrador.",
     "acao_sfc", "Executar o SFC? A operação pede permissão de administrador."),
    ("rede", "Reparar rede", "Renova o IP, limpa o DNS e reinicia Winsock e TCP/IP.",
     "acao_reparar_rede", "Reparar a rede? A conexão cai por alguns segundos."),
    ("limpeza", "Limpar temporários", "Apaga os arquivos temporários do usuário e do Windows.",
     "acao_limpar_temp", "Apagar os arquivos temporários do usuário e do Windows?"),
)


class App(OutlookMixin, OfficeMixin, RedeMixin, ManutencaoMixin, ExecucaoMixin, ctk.CTk):
    def __init__(self):
        super().__init__()

        self.deve_parar = False
        self.pasta_destino_pst = os.path.join(os.path.expanduser("~"), "Documents")
        self.ultima_pasta_pst = None
        self._tags_configuradas = set()
        self._tags_configuradas_ferramentas = set()
        self._linhas_log = []
        self._itens_menu = {}
        self._secoes = {}
        self._botoes_acoes_rapidas = {}
        self._progresso_acoes_rapidas = {}
        self._acao_rapida_em_execucao = None
        self._cancelar_ferramenta = False
        self._processo_ferramenta_ativo = None

        self.title("Suporte TI")
        self.minsize(1180, 720)
        self._centralizar_janela(1320, 840)
        try:
            self.iconbitmap(resource_path("logo.ico"))
        except Exception:
            pass

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._montar_menu()
        self._montar_conteudo()
        self._montar_registro()

        self._mostrar_secao("Visão geral")
        self.after(250, self._atualizar_resumo_sistema_async)

    # ------------------------------------------------------------------
    # Estrutura da janela
    # ------------------------------------------------------------------

    def _montar_menu(self):
        menu = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=tema.MENU_FUNDO)
        menu.grid(row=0, column=0, sticky="nsew")
        menu.grid_propagate(False)
        menu.pack_propagate(False)

        ctk.CTkLabel(menu, text="Suporte TI", font=tema.TITULO_APP, text_color="#FFFFFF", anchor="w").pack(
            fill="x", padx=22, pady=(26, 0)
        )
        self.lbl_maquina_menu = ctk.CTkLabel(
            menu, text=platform.node() or "", font=tema.APOIO, text_color=tema.MENU_APOIO, anchor="w"
        )
        self.lbl_maquina_menu.pack(fill="x", padx=22, pady=(0, 22))

        grupo_atual = None
        for nome, _descricao, _metodo, grupo, chave_cor in SECOES:
            if grupo != grupo_atual:
                grupo_atual = grupo
                ctk.CTkLabel(menu, text=grupo, font=tema.APOIO, text_color=tema.MENU_APOIO, anchor="w").pack(
                    fill="x", padx=22, pady=(18, 2)
                )
            item = ctk.CTkFrame(menu, fg_color="transparent", corner_radius=6, height=36)
            item.pack(fill="x", padx=10, pady=1)
            item.pack_propagate(False)
            # O "cabo": quadradinho com a cor da seção (itens do grupo Outlook ficam recuados).
            ctk.CTkFrame(item, width=10, height=10, corner_radius=2, fg_color=tema.PALETA[chave_cor][0]).pack(
                side="left", padx=(24 if grupo else 12, 0)
            )
            botao = ctk.CTkButton(
                item, text=nome, anchor="w", height=34, corner_radius=6, font=tema.MENU,
                fg_color="transparent", hover_color=tema.MENU_ATIVO, text_color=tema.MENU_TEXTO,
                command=lambda n=nome: self._mostrar_secao(n)
            )
            botao.pack(side="left", fill="x", expand=True, padx=(4, 2))
            self._itens_menu[nome] = (item, botao)

        rodape = ctk.CTkFrame(menu, fg_color="transparent")
        rodape.pack(side="bottom", fill="x", padx=14, pady=16)
        ctk.CTkLabel(
            rodape, text="Web Developer Alex Teixeira", font=tema.APOIO, text_color=tema.MENU_APOIO, anchor="w"
        ).pack(fill="x", padx=4)

    def _montar_conteudo(self):
        conteudo = ctk.CTkFrame(self, corner_radius=0, fg_color=tema.FUNDO)
        conteudo.grid(row=0, column=1, sticky="nsew")
        conteudo.grid_columnconfigure(0, weight=1)
        conteudo.grid_rowconfigure(1, weight=1)

        # Faixa do título: pinta com a cor da seção atual.
        self.faixa_secao = ctk.CTkFrame(conteudo, corner_radius=0, height=84)
        self.faixa_secao.grid(row=0, column=0, sticky="ew")
        self.faixa_secao.grid_propagate(False)
        self.faixa_secao.grid_columnconfigure(0, weight=1)
        self.lbl_titulo_secao = ctk.CTkLabel(self.faixa_secao, text="", font=tema.TITULO_SECAO, anchor="w")
        self.lbl_titulo_secao.grid(row=0, column=0, sticky="ew", padx=32, pady=(22, 0))
        self.lbl_subtitulo_secao = ctk.CTkLabel(self.faixa_secao, text="", font=tema.TEXTO, anchor="w")
        self.lbl_subtitulo_secao.grid(row=1, column=0, sticky="ew", padx=32)

        for nome, _descricao, metodo, _grupo, chave_cor in SECOES:
            pagina = ctk.CTkFrame(conteudo, fg_color="transparent", corner_radius=0)
            pagina.grid(row=1, column=0, sticky="nsew", padx=(26, 18), pady=(20, 18))
            pagina.grid_remove()
            self._secoes[nome] = pagina
            tema.usar(chave_cor)
            getattr(self, metodo)(pagina)

    def _montar_registro(self):
        """Coluna escura fixa com tudo que foi executado: a "fita" da sessão de atendimento."""
        coluna = ctk.CTkFrame(self, width=350, corner_radius=0, fg_color=tema.REG_FUNDO)
        coluna.grid(row=0, column=2, sticky="nsew")
        coluna.grid_propagate(False)
        coluna.grid_columnconfigure(0, weight=1)
        coluna.grid_rowconfigure(3, weight=1)

        topo = ctk.CTkFrame(coluna, fg_color="transparent")
        topo.grid(row=0, column=0, sticky="ew", padx=20, pady=(28, 0))
        ctk.CTkLabel(topo, text="Registro", font=tema.TITULO_GRUPO, text_color="#FFFFFF").pack(side="left")
        self.lbl_estado_atividade = ctk.CTkLabel(topo, text="Pronto", font=tema.APOIO, text_color=tema.REG_APOIO)
        self.lbl_estado_atividade.pack(side="right")

        self.lbl_status_ferramentas = ctk.CTkLabel(
            coluna, text="Nenhuma ação executada nesta sessão.", font=tema.APOIO, text_color=tema.REG_APOIO,
            anchor="w", justify="left", wraplength=300
        )
        self.lbl_status_ferramentas.grid(row=1, column=0, sticky="ew", padx=20, pady=(6, 8))

        self.progress_ferramentas = ctk.CTkProgressBar(
            coluna, height=3, fg_color="#1E293B", progress_color=tema.REG_LINK
        )
        self.progress_ferramentas.set(0)
        self.progress_ferramentas.grid(row=2, column=0, sticky="ew", padx=20)

        self.txt_log_ferramentas = ctk.CTkTextbox(
            coluna, font=tema.MONO, corner_radius=0, border_width=0, fg_color=tema.REG_FUNDO,
            text_color=tema.REG_TEXTO, wrap="word", activate_scrollbars=True,
            scrollbar_button_color="#334155", scrollbar_button_hover_color="#475569"
        )
        self.txt_log_ferramentas.grid(row=3, column=0, sticky="nsew", padx=(14, 6), pady=(10, 0))

        self.btn_cancelar_acao = tema.botao_perigo(coluna, "Cancelar ação", self._cancelar_acao_em_execucao)
        self.btn_cancelar_acao.grid(row=4, column=0, sticky="ew", padx=20, pady=(10, 0))
        self.btn_cancelar_acao.grid_remove()

        controles = ctk.CTkFrame(coluna, fg_color="transparent")
        controles.grid(row=5, column=0, sticky="ew", padx=14, pady=(8, 16))
        for texto, comando in (("Copiar", self._copiar_log), ("Limpar", self._limpar_log_geral)):
            tema.botao_link(controles, texto, comando, width=0, text_color=tema.REG_LINK,
                            hover_color="#1E293B").pack(side="left", padx=(0, 4))

    def _mostrar_secao(self, nome):
        for pagina in self._secoes.values():
            pagina.grid_remove()
        self._secoes[nome].grid()
        # Widgets criados depois (listas de resultado) usam a cor da seção visível.
        tema.usar(COR_SECAO[nome])
        descricao = next(s[1] for s in SECOES if s[0] == nome)
        titulo = nome
        if nome == "Visão geral" and getattr(self, "_ultimo_resumo_sistema", None):
            titulo = self._ultimo_resumo_sistema["computador"]
        # Cabeçalho escuro e discreto (padrão da captura); o azul fica só nos destaques.
        self.faixa_secao.configure(fg_color=tema.FUNDO)
        self.lbl_titulo_secao.configure(text=titulo, text_color="#FFFFFF")
        self.lbl_subtitulo_secao.configure(text=descricao, text_color=tema.CHUMBO)
        self._secao_atual = nome
        for item, (moldura, botao) in self._itens_menu.items():
            ativo = item == nome
            moldura.configure(fg_color=tema.MENU_ATIVO if ativo else "transparent")
            botao.configure(text_color="#FFFFFF" if ativo else tema.MENU_TEXTO,
                            font=tema.MENU_ATIVO_FONTE if ativo else tema.MENU)

    def _centralizar_janela(self, largura, altura):
        self.update_idletasks()
        x = max((self.winfo_screenwidth() - largura) // 2, 0)
        y = max((self.winfo_screenheight() - altura) // 2 - 20, 0)
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    # ------------------------------------------------------------------
    # Visão geral
    # ------------------------------------------------------------------

    def _montar_visao_geral(self, aba):
        pagina = tema.pagina(aba)

        tema.titulo_grupo(pagina, "Esta máquina", primeiro=True)
        ficha = ctk.CTkFrame(pagina, fg_color=tema.PAINEL, corner_radius=4, border_width=1, border_color=tema.LINHA)
        ficha.pack(fill="x", padx=2)
        ficha.grid_columnconfigure((1, 3), weight=1)
        self._labels_resumo = {}
        campos = (
            ("windows", "Windows"), ("internet", "Internet"),
            ("ip", "Endereço IP"), ("winget", "Winget"),
            ("disco", "Disco do sistema"), ("outlook", "Outlook"),
            (None, None), ("admin", "Permissão"),
        )
        for indice, (chave, rotulo) in enumerate(campos):
            linha, coluna = divmod(indice, 2)
            if chave is None:
                continue
            ctk.CTkLabel(ficha, text=rotulo, font=tema.APOIO, text_color=tema.CHUMBO, anchor="w").grid(
                row=linha, column=coluna * 2, sticky="nw", padx=(18, 12), pady=(14 if linha == 0 else 4, 4)
            )
            valor = ctk.CTkLabel(ficha, text="Lendo...", font=tema.TEXTO, anchor="w", justify="left")
            valor.grid(row=linha, column=coluna * 2 + 1, sticky="nw", pady=(14 if linha == 0 else 4, 4))
            self._labels_resumo[chave] = valor
        self.progress_disco = ctk.CTkProgressBar(ficha, height=4, width=180)
        self.progress_disco.set(0)
        self.progress_disco.grid(row=3, column=1, sticky="w", pady=(0, 4))  # vaga (None) abaixo do disco
        rodape = ctk.CTkFrame(ficha, fg_color="transparent")
        rodape.grid(row=4, column=0, columnspan=4, sticky="ew", padx=12, pady=(4, 10))
        self.lbl_ultima_verificacao = ctk.CTkLabel(rodape, text="", font=tema.APOIO, text_color=tema.CHUMBO)
        self.lbl_ultima_verificacao.pack(side="left", padx=6)
        tema.botao_link(rodape, "Ler novamente", self._atualizar_resumo_sistema_async, width=0).pack(side="right")

        tema.titulo_grupo(pagina, "Correções frequentes")
        lista = ctk.CTkFrame(pagina, fg_color=tema.PAINEL, corner_radius=4, border_width=1, border_color=tema.LINHA)
        lista.pack(fill="x", padx=2)
        lista.grid_columnconfigure(0, weight=1)
        for indice, (chave, titulo, descricao, metodo, confirmacao) in enumerate(ACOES_FREQUENTES):
            if indice:
                tema.fio(lista).grid(
                    row=indice * 3, column=0, columnspan=2, sticky="ew", padx=18
                )
            ctk.CTkLabel(lista, text=titulo, font=tema.TEXTO_FORTE, anchor="w").grid(
                row=indice * 3 + 1, column=0, sticky="ew", padx=18, pady=(12, 0)
            )
            ctk.CTkLabel(lista, text=descricao, font=tema.APOIO, text_color=tema.CHUMBO, anchor="w").grid(
                row=indice * 3 + 2, column=0, sticky="ew", padx=18, pady=(0, 12)
            )
            botao = tema.botao_primario(
                lista, "Executar",
                lambda c=chave, m=metodo, t=confirmacao: self._executar_acao_rapida(c, getattr(self, m), t),
                width=110
            )
            botao.grid(row=indice * 3 + 1, column=1, rowspan=2, padx=18)
            progresso = ctk.CTkProgressBar(lista, height=3, width=110, mode="indeterminate")
            progresso.grid(row=indice * 3 + 2, column=1, sticky="s", padx=18, pady=(0, 6))
            progresso.grid_remove()
            self._botoes_acoes_rapidas[chave] = botao
            self._progresso_acoes_rapidas[chave] = progresso

    def _executar_acao_rapida(self, chave, acao, confirmacao):
        if self._acao_rapida_em_execucao is not None:
            self.exibir_aviso("Ação em andamento", "Espere a ação atual terminar para iniciar outra.")
            return
        if confirmacao and not messagebox.askyesno("Confirmar", confirmacao):
            return
        self._acao_rapida_em_execucao = chave
        self._cancelar_ferramenta = False
        for botao in self._botoes_acoes_rapidas.values():
            botao.configure(state="disabled")
        self._botoes_acoes_rapidas[chave].configure(text="Executando")
        self._progresso_acoes_rapidas[chave].grid()
        self._progresso_acoes_rapidas[chave].start()
        self.btn_cancelar_acao.grid()
        self.atualizar_status_ferramentas("Preparando a ação...")

        def executar():
            try:
                acao()
            finally:
                self.after(0, self._finalizar_acao_rapida)
        threading.Thread(target=executar, daemon=True).start()

    def _finalizar_acao_rapida(self):
        self._acao_rapida_em_execucao = None
        # Sem este reset, depois de um "Cancelar" todas as próximas ações
        # (winget, gpupdate, reparo de rede...) abortariam imediatamente.
        self._cancelar_ferramenta = False
        for progresso in self._progresso_acoes_rapidas.values():
            progresso.stop()
            progresso.grid_remove()
        for botao in self._botoes_acoes_rapidas.values():
            botao.configure(state="normal", text="Executar")
        self.btn_cancelar_acao.grid_remove()
        self.btn_cancelar_acao.configure(state="normal", text="Cancelar ação")

    def _cancelar_acao_em_execucao(self):
        self._cancelar_ferramenta = True
        processo = self._processo_ferramenta_ativo
        if processo is not None and processo.poll() is None:
            try:
                processo.terminate()
            except Exception:
                pass
        self.log_geral("Cancelamento pedido pelo técnico.", "aviso")
        self.atualizar_status_ferramentas("Cancelando...")
        self.btn_cancelar_acao.configure(state="disabled", text="Cancelando")

    def _obter_resumo_sistema(self):
        resumo = {"computador": platform.node() or "Computador sem nome"}
        try:
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
            ) as chave_windows:
                build = str(winreg.QueryValueEx(chave_windows, "CurrentBuildNumber")[0])
                try:
                    ubr = str(winreg.QueryValueEx(chave_windows, "UBR")[0])
                except OSError:
                    ubr = "0"
                try:
                    versao_exibicao = str(winreg.QueryValueEx(chave_windows, "DisplayVersion")[0])
                except OSError:
                    versao_exibicao = ""
                nome_windows = "Windows 11" if int(build) >= 22000 else "Windows 10"
                resumo["windows"] = nome_windows
                resumo["windows_detalhe"] = f"{versao_exibicao} ({build}.{ubr})" if versao_exibicao else f"Build {build}.{ubr}"
        except (OSError, ValueError):
            resumo["windows"] = f"Windows {platform.release()}"
            resumo["windows_detalhe"] = platform.version()
        try:
            conexao = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            conexao.connect(("8.8.8.8", 80))
            resumo["ip"] = conexao.getsockname()[0]
            conexao.close()
        except OSError:
            resumo["ip"] = "Sem conexão"
        try:
            uso_disco = shutil.disk_usage(os.environ.get("SystemDrive", "C:") + "\\")
            livre_gb = uso_disco.free / (1024 ** 3)
            total_gb = uso_disco.total / (1024 ** 3)
            resumo["disco"] = f"{livre_gb:.0f} GB livres de {total_gb:.0f} GB"
            resumo["disco_livre_fracao"] = uso_disco.free / uso_disco.total
        except OSError:
            resumo["disco"] = "Indisponível"
            resumo["disco_livre_fracao"] = 0
        resumo["internet"] = resumo["ip"] != "Sem conexão"
        resumo["winget"] = shutil.which("winget") is not None
        try:
            resultado = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE"], capture_output=True, text=True, creationflags=CREATE_NO_WINDOW, timeout=5)
            resumo["outlook"] = "OUTLOOK.EXE" in (resultado.stdout or "").upper()
        except Exception:
            resumo["outlook"] = False
        try:
            resumo["admin"] = bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            resumo["admin"] = False
        return resumo

    def _atualizar_resumo_sistema_async(self):
        def executar():
            resumo = self._obter_resumo_sistema()
            self._ultimo_resumo_sistema = resumo
            self.after(0, lambda: self._aplicar_resumo_sistema(resumo))
        threading.Thread(target=executar, daemon=True).start()

    def _aplicar_resumo_sistema(self, resumo):
        # (texto se verdadeiro, texto se falso, falso é problema?)
        estados = {
            "internet": ("Conectada", "Sem conexão", True),
            "winget": ("Disponível", "Não instalado", True),
            "outlook": ("Aberto", "Fechado", False),
            "admin": ("Administrador", "Usuário comum", False),
        }
        self._labels_resumo["windows"].configure(text=f"{resumo['windows']} {resumo.get('windows_detalhe', '')}".strip())
        self._labels_resumo["ip"].configure(text=resumo["ip"], text_color=tema.TINTA if resumo["internet"] else tema.AVISO)
        for chave, (sim, nao, problema) in estados.items():
            ok = resumo[chave]
            cor_estado = tema.OK if ok and problema else tema.AVISO if problema else tema.TINTA
            self._labels_resumo[chave].configure(text=sim if ok else nao, text_color=cor_estado)
        fracao_livre = resumo.get("disco_livre_fracao", 0)
        cor = tema.ERRO if fracao_livre < .1 else tema.AVISO if fracao_livre < .2 else tema.OK
        self._labels_resumo["disco"].configure(text=f"{resumo['disco']} ({fracao_livre:.0%})", text_color=cor)
        self.progress_disco.set(1 - fracao_livre)  # a barra mostra o espaço ocupado
        self.progress_disco.configure(progress_color=cor)
        self.lbl_ultima_verificacao.configure(text=f"Lido às {datetime.now():%H:%M}")
        self.lbl_maquina_menu.configure(text=resumo["computador"])
        if getattr(self, "_secao_atual", None) == "Visão geral":
            self.lbl_titulo_secao.configure(text=resumo["computador"])


def main():
    tema.aplicar()
    App().mainloop()
