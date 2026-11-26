"""Abas do Outlook: conta e PST, arquivamento e limpeza da caixa."""

import ctypes
import os
import shutil
import threading
import time
from datetime import datetime
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pythoncom
import win32com.client

from . import tema

ROTULO_CATEGORIA = {"newsletter": "Newsletter", "palavra": "Palavra-chave", "grande": "Mensagem grande"}
ORDEM_CATEGORIA = {"newsletter": 0, "palavra": 1, "grande": 2}

# O Outlook trava o PST Unicode em 50 GB (MaxLargeFileSize padrão); 45 GB deixa folga.
LIMITE_PST_BYTES = 45 * 1024 ** 3
DRIVE_REMOTE = 4


def problema_destino_pst(pasta):
    """Motivo para não gravar PST nesta pasta, ou None.

    A Microsoft não dá suporte a PST em rede nem em pasta sincronizada: o
    arquivo corrompe, e os e-mails movidos para ele já saíram do servidor.
    """
    # normpath antes: o askdirectory devolve "//srv/pasta", e o realpath de um UNC
    # inacessível vira um caminho local ("D:\srv\pasta").
    if os.path.normpath(pasta).startswith("\\\\"):
        return "fica na rede"
    caminho = os.path.realpath(pasta)
    unidade = os.path.splitdrive(caminho)[0]
    if unidade and ctypes.windll.kernel32.GetDriveTypeW(unidade + "\\") == DRIVE_REMOTE:
        return "fica numa unidade de rede"
    # shortcut: só detecta OneDrive (pasta sempre tem "OneDrive" no nome); Dropbox/Google Drive passam.
    if "onedrive" in caminho.lower():
        return "é sincronizada pelo OneDrive"
    return None

class OutlookMixin:
    # Termos comuns em remetente/assunto de e-mails de newsletter, marketing e avisos automáticos.
    PALAVRAS_CHAVE_INFORMATIVO = [
        "noreply", "no-reply", "não-responda", "nao-responda", "newsletter", "informativo",
        "boletim", "marketing", "mailer", "notification", "notificacao", "notificação",
        "alerta", "comunicado", "divulgacao", "divulgação", "campanha", "promo", "promoção",
        "unsubscribe", "descadastr",
    ]

    # ==================================================================
    # Conta e PST, arquivamento e limpeza da caixa
    # ==================================================================

    def _montar_aba_configuracao(self, aba):
        pagina = tema.pagina(aba)

        ctk.CTkLabel(pagina, text="Conta do Outlook", font=tema.APOIO, text_color=tema.CHUMBO, anchor="w").pack(
            anchor="w", padx=2, pady=(0, 1)
        )
        self.txt_email = ctk.CTkComboBox(pagina, values=["Lendo contas do Outlook..."], height=34, width=420, font=tema.TEXTO)
        self.txt_email.set("Selecione ou digite o e-mail")
        self.txt_email.pack(anchor="w", padx=2)

        self._validar_ano_cmd = (self.register(self._validar_entrada_ano), "%P")
        self.txt_ano = tema.campo(
            pagina, "Ano a arquivar", "2024", largura=120, validate="key", validatecommand=self._validar_ano_cmd
        )

        ctk.CTkLabel(pagina, text="Pasta de origem", font=tema.APOIO, text_color=tema.CHUMBO, anchor="w").pack(
            anchor="w", padx=2, pady=(10, 2)
        )
        self.origem_var = ctk.StringVar(value="Inbox")
        self.frame_radios = ctk.CTkFrame(pagina, fg_color="transparent")
        self.frame_radios.pack(anchor="w", padx=2)
        self.rb_inbox = ctk.CTkRadioButton(self.frame_radios, text="Caixa de Entrada", variable=self.origem_var, value="Inbox", font=tema.TEXTO, fg_color=tema.cor(), hover_color=tema.cor())
        self.rb_inbox.pack(side="left", padx=(0, 24))
        self.rb_sent = ctk.CTkRadioButton(self.frame_radios, text="Itens Enviados", variable=self.origem_var, value="Sent", font=tema.TEXTO, fg_color=tema.cor(), hover_color=tema.cor())
        self.rb_sent.pack(side="left")

        ctk.CTkLabel(pagina, text="Salvar o PST em", font=tema.APOIO, text_color=tema.CHUMBO, anchor="w").pack(
            anchor="w", padx=2, pady=(12, 1)
        )
        self.frame_destino = ctk.CTkFrame(pagina, fg_color="transparent")
        self.frame_destino.pack(fill="x", padx=2)
        self.txt_destino = tema.entrada(self.frame_destino, "Documentos (padrão)")
        self.txt_destino.configure(state="readonly")
        self.txt_destino.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.btn_procurar = tema.botao_secundario(self.frame_destino, "Escolher pasta", self.selecionar_pasta_destino, width=130)
        self.btn_procurar.pack(side="left")

        self.manter_pst_aberto = ctk.BooleanVar(value=False)
        self.chk_remover_pst = ctk.CTkCheckBox(
            pagina, text="Deixar o PST aberto no Outlook ao terminar", variable=self.manter_pst_aberto, font=tema.TEXTO, fg_color=tema.cor()
        )
        self.chk_remover_pst.pack(anchor="w", padx=2, pady=(14, 18))

        self.btn_ir_para_execucao = tema.botao_primario(
            pagina, "Continuar para o arquivamento", lambda: self._mostrar_secao("Arquivar e-mails"), width=240
        )
        self.btn_ir_para_execucao.pack(anchor="w", padx=2)

    def _montar_aba_execucao(self, aba):
        frame = ctk.CTkFrame(aba, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=2)
        tema.apoio(frame, "Usa a conta, o ano e a pasta definidos em Conta e PST. Conte antes para saber quanto será movido.")

        self.frame_botoes = ctk.CTkFrame(frame, fg_color="transparent")
        self.frame_botoes.pack(fill="x", pady=(0, 12))
        self.btn_verificar = tema.botao_secundario(self.frame_botoes, "Contar e-mails", self.disparar_verificacao, width=150)
        self.btn_verificar.pack(side="left")
        self.btn_iniciar = tema.botao_primario(self.frame_botoes, "Arquivar", self.disparar_arquivamento, width=150)
        self.btn_iniciar.pack(side="left", padx=(8, 0))
        self.btn_cancelar = tema.botao_secundario(
            self.frame_botoes, "Cancelar", self.solicitar_cancelamento, width=110, state="disabled"
        )
        self.btn_cancelar.pack(side="left", padx=(8, 0))

        self.lbl_status = ctk.CTkLabel(frame, text="Nada em andamento.", font=tema.APOIO, text_color=tema.CHUMBO, anchor="w")
        self.lbl_status.pack(fill="x", pady=(0, 4))
        self.progress = ctk.CTkProgressBar(frame, height=4, progress_color=tema.cor())
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(0, 10))

        self.txt_log = ctk.CTkTextbox(frame, font=tema.MONO, border_width=1)
        self.txt_log.pack(fill="both", expand=True)

        self.btn_abrir_pasta = tema.botao_link(frame, "Abrir a pasta do PST", self.abrir_pasta_destino, state="disabled", width=0)
        self.btn_abrir_pasta.pack(anchor="w", pady=(6, 0))

        self.after(300, self.carregar_contas_outlook_async)

    def _validar_entrada_ano(self, valor_proposto):
        if valor_proposto == "":
            return True
        return valor_proposto.isdigit() and len(valor_proposto) <= 4

    def selecionar_pasta_destino(self):
        pasta_escolhida = filedialog.askdirectory(title="Selecione onde salvar o arquivo PST", initialdir=self.pasta_destino_pst)
        if pasta_escolhida:
            self.pasta_destino_pst = pasta_escolhida
            self.txt_destino.configure(state="normal")
            self.txt_destino.delete(0, "end")
            self.txt_destino.insert(0, pasta_escolhida)
            self.txt_destino.configure(state="readonly")

    def abrir_pasta_destino(self):
        if self.ultima_pasta_pst and os.path.isdir(self.ultima_pasta_pst):
            os.startfile(self.ultima_pasta_pst)
        else:
            messagebox.showinfo("Aviso", "Nenhuma pasta de arquivamento disponível ainda.")

    def carregar_contas_outlook_async(self):
        threading.Thread(target=self._carregar_contas_thread, daemon=True).start()

    def _carregar_contas_thread(self):
        contas = []
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            for account in namespace.Session.Accounts:
                try:
                    email = account.SmtpAddress
                    if email and email not in contas:
                        contas.append(email)
                except Exception:
                    continue
        except Exception as e:
            self.log(f"Não foi possível carregar as contas do Outlook automaticamente: {e}", "aviso")
            self.log("Verifique se o Outlook está instalado e pode ser aberto normalmente.", "aviso")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        self.after(0, lambda: self._atualizar_combo_contas(contas))

    def _atualizar_combo_contas(self, contas):
        if contas:
            self.txt_email.configure(values=contas)
            self.txt_email.set(contas[0])
            self.log(f"{len(contas)} conta(s) do Outlook detectada(s) e carregada(s) na lista.", "sucesso")
        else:
            self.txt_email.configure(values=["Nenhuma conta detectada - digite manualmente"])
            self.txt_email.set("")

    def log(self, mensagem, tipo="info"):
        self.after(0, self._log_seguro, mensagem, tipo)

    def _log_seguro(self, mensagem, tipo="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        linha = f"[{timestamp}] {mensagem}\n"
        self._linhas_log.append(linha)
        cores = tema.CORES_LOG
        try:
            if tipo in cores:
                if tipo not in self._tags_configuradas:
                    self.txt_log.tag_config(tipo, foreground=cores[tipo])
                    self._tags_configuradas.add(tipo)
                self.txt_log.insert("end", linha, tipo)
            else:
                self.txt_log.insert("end", linha)
        except Exception:
            self.txt_log.insert("end", linha)
        self.txt_log.see("end")

    def atualizar_status(self, message):
        self.after(0, lambda: self.lbl_status.configure(text=message))

    def atualizar_progresso(self, valor):
        self.after(0, lambda: self.progress.set(valor))

    def limpar_log(self):
        self.txt_log.delete("1.0", "end")
        self._linhas_log = []

    def salvar_log_em_arquivo(self, pasta_destino, nome_base):
        try:
            if not self._linhas_log:
                return None
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_log = f"log_{nome_base}_{timestamp}.txt"
            caminho_log = os.path.join(pasta_destino, nome_log)
            with open(caminho_log, "w", encoding="utf-8") as f:
                f.writelines(self._linhas_log)
            return caminho_log
        except Exception:
            return None

    def solicitar_cancelamento(self):
        self.deve_parar = True
        self.atualizar_status("Interrompendo processo... Aguarde.")
        self.log("\n[!] Solicitação de cancelamento enviada pelo usuário.")

    def _localizar_conta_e_pasta(self, namespace, email_alvo, pasta_escolhida):
        pasta_raiz_conta = None
        alvo_lower = email_alvo.strip().lower()
        candidatos_parciais = []

        # Correspondência exata tem prioridade. A versão anterior aceitava
        # substring, então "ana@empresa.com" podia selecionar a caixa de
        # "joana@empresa.com" (e mover/arquivar e-mails da conta errada).
        for folder in namespace.Folders:
            try:
                nome_store = (folder.Name or "").strip().lower()
                email_real = ""
                try:
                    store = folder.Store
                    if store is not None:
                        email_real = str(store.GetRootFolder().PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                        )).strip().lower()
                except Exception:
                    pass
                if alvo_lower in (nome_store, email_real):
                    pasta_raiz_conta = folder
                    break
                if alvo_lower in nome_store or (email_real and alvo_lower in email_real):
                    candidatos_parciais.append(folder)
            except Exception:
                continue

        if not pasta_raiz_conta and len(candidatos_parciais) == 1:
            pasta_raiz_conta = candidatos_parciais[0]
        if not pasta_raiz_conta:
            if len(candidatos_parciais) > 1:
                return None, None, (
                    f"Mais de uma conta do Outlook corresponde a '{email_alvo}'. "
                    "Digite o endereço de e-mail completo e exato."
                )
            return None, None, f"A conta '{email_alvo}' não foi encontrada no Outlook."

        # GetDefaultFolder funciona em qualquer idioma do Outlook; os nomes
        # ficam apenas como alternativa (ex.: arquivos PST).
        pasta_origem = None
        try:
            pasta_origem = pasta_raiz_conta.Store.GetDefaultFolder(6 if pasta_escolhida == "Inbox" else 5)
        except Exception:
            pasta_origem = None

        nomes_inbox = ["caixa de entrada", "inbox"] if pasta_escolhida == "Inbox" else ["itens enviados", "sent items", "sent", "enviados"]
        for subfolder in ([] if pasta_origem is not None else pasta_raiz_conta.Folders):
            try:
                if subfolder.Name.lower() in nomes_inbox:
                    pasta_origem = subfolder
                    break
            except Exception:
                continue

        if not pasta_origem:
            return pasta_raiz_conta, None, "Não foi possível localizar a pasta de origem selecionada nesta conta."

        return pasta_raiz_conta, pasta_origem, None

    def _filtrar_emails_por_ano(self, pasta_origem, pasta_escolhida, ano_alvo):
        if pasta_escolhida == "Inbox":
            propriedade_dasl = "urn:schemas:httpmail:datereceived"
        else:
            propriedade_dasl = "http://schemas.microsoft.com/mapi/proptag/0x00390040"

        data_inicio = f"{ano_alvo}/01/01 00:00:00"
        data_fim = f"{ano_alvo}/12/31 23:59:59"

        filtro_dasl = (
            f"@SQL=\"{propriedade_dasl}\" >= '{data_inicio}' AND "
            f"\"{propriedade_dasl}\" <= '{data_fim}'"
        )

        try:
            itens_filtrados = pasta_origem.Items.Restrict(filtro_dasl)
        except Exception as e_filtro:
            self.log(f"Aviso: filtro DASL falhou ({e_filtro}), usando filtro alternativo.", "aviso")
            propriedade_data = "[ReceivedTime]" if pasta_escolhida == "Inbox" else "[SentOn]"
            filtro_jet = (
                f"{propriedade_data} >= '01/01/{ano_alvo} 00:00 AM' AND "
                f"{propriedade_data} <= '31/12/{ano_alvo} 11:59 PM'"
            )
            itens_filtrados = pasta_origem.Items.Restrict(filtro_jet)

        propriedade_sort = "[ReceivedTime]" if pasta_escolhida == "Inbox" else "[SentOn]"
        try:
            itens_filtrados.Sort(propriedade_sort, True)
        except Exception:
            pass

        emails_encontrados = []
        for item in itens_filtrados:
            if self.deve_parar:
                break
            try:
                if item and hasattr(item, "Class") and item.Class == 43:
                    emails_encontrados.append(item)
            except Exception:
                continue

        return emails_encontrados

    def disparar_verificacao(self):
        email_alvo = self.txt_email.get().strip()
        ano_alvo = self.txt_ano.get().strip()

        if not email_alvo or email_alvo == "Selecione ou digite o e-mail" or "Nenhuma conta detectada" in email_alvo:
            messagebox.showwarning("Aviso", "Por favor, selecione ou digite o e-mail da conta.")
            return
        if not ano_alvo or not ano_alvo.isdigit() or len(ano_alvo) != 4:
            messagebox.showwarning("Aviso", "Por favor, digite um ano válido com 4 dígitos (Ex: 2024).")
            return

        self.btn_verificar.configure(state="disabled", text="Verificando...")
        self.atualizar_status("Verificando quantidade de e-mails...")
        threading.Thread(
            target=self._processo_verificacao, args=(email_alvo, ano_alvo, self.origem_var.get()), daemon=True
        ).start()

    def _processo_verificacao(self, email_alvo, ano_alvo, pasta_escolhida):
        pythoncom.CoInitialize()
        pasta_label = "Caixa de Entrada" if pasta_escolhida == "Inbox" else "Itens Enviados"
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                self.exibir_erro("Erro", f"Não foi possível conectar ao Outlook.\nVerifique se ele está instalado e aberto.\n\n{e}")
                return

            pasta_raiz_conta, pasta_origem, erro = self._localizar_conta_e_pasta(namespace, email_alvo, pasta_escolhida)
            if erro:
                self.exibir_erro("Erro", erro)
                return

            emails = self._filtrar_emails_por_ano(pasta_origem, pasta_escolhida, ano_alvo)
            total = len(emails)

            self.atualizar_status(f"Verificação concluída: {total} e-mail(s) encontrado(s).")
            self.exibir_info(
                "Verificação concluída",
                f"Conta: {email_alvo}\nPasta: {pasta_label}\nAno: {ano_alvo}\n\n"
                f"{total} e-mail(s) seriam movidos para o PST nesta operação.\n\n"
                "Nenhum e-mail foi movido — esta é apenas uma contagem."
            )
        except Exception as e:
            self.exibir_erro("Erro", f"Ocorreu um erro ao verificar: {e}")
        finally:
            self.after(0, lambda: self.btn_verificar.configure(state="normal", text="Contar e-mails"))
            pythoncom.CoUninitialize()

    def disparar_arquivamento(self):
        email_alvo = self.txt_email.get().strip()
        ano_alvo = self.txt_ano.get().strip()
        pasta_label = "Caixa de Entrada" if self.origem_var.get() == "Inbox" else "Itens Enviados"

        if not email_alvo or email_alvo == "Selecione ou digite o e-mail" or "Nenhuma conta detectada" in email_alvo:
            messagebox.showwarning("Aviso", "Por favor, selecione ou digite o e-mail da conta.")
            return
        if "@" not in email_alvo or "." not in email_alvo.split("@")[-1]:
            messagebox.showwarning("Aviso", "O e-mail informado não parece válido. Verifique e tente novamente.")
            return
        if not ano_alvo:
            messagebox.showwarning("Aviso", "Por favor, digite o ano para arquivar.")
            return
        motivo = problema_destino_pst(self.pasta_destino_pst)
        if motivo:
            messagebox.showerror(
                "Pasta não permitida",
                f"Não salve o PST em {self.pasta_destino_pst}: a pasta {motivo}.\n\n"
                "PST em rede ou no OneDrive corrompe, e a Microsoft não dá suporte. "
                "Escolha uma pasta local, por exemplo C:\\Arquivos_PST."
            )
            return

        confirmar = messagebox.askyesno(
            "Confirmar Arquivamento",
            "Você está prestes a iniciar o arquivamento com os seguintes dados:\n\n"
            f"Conta: {email_alvo}\nAno: {ano_alvo}\nPasta de origem: {pasta_label}\n"
            f"Salvar PST em: {self.pasta_destino_pst}\n\nDeseja continuar?"
        )
        if not confirmar:
            return

        self.deve_parar = False
        self.limpar_log()
        self.progress.set(0)
        self.btn_abrir_pasta.configure(state="disabled")
        self.ultima_pasta_pst = None

        self.txt_email.configure(state="disabled")
        self.txt_ano.configure(state="disabled")
        self.rb_inbox.configure(state="disabled")
        self.rb_sent.configure(state="disabled")
        self.btn_procurar.configure(state="disabled")
        self.chk_remover_pst.configure(state="disabled")
        self.btn_verificar.configure(state="disabled")

        self.btn_iniciar.configure(state="disabled")
        self.btn_cancelar.configure(state="normal")

        # Os valores dos campos são lidos aqui, na thread da interface, e
        # repassados à thread de trabalho (o Tkinter não é thread-safe).
        threading.Thread(
            target=self.processo_arquivamento,
            args=(email_alvo, ano_alvo, self.origem_var.get(), self.manter_pst_aberto.get()),
            daemon=True,
        ).start()

    def processo_arquivamento(self, email_alvo, ano_alvo, pasta_escolhida, manter_no_outlook):
        pythoncom.CoInitialize()

        if not email_alvo:
            self.exibir_aviso("Aviso", "Por favor, preencha o campo 'E-mail da conta'.")
            self.finalizar_interface()
            pythoncom.CoUninitialize()
            return
        if not ano_alvo or not ano_alvo.isdigit() or len(ano_alvo) != 4:
            self.exibir_aviso("Aviso", "Por favor, digite um ano válido com 4 dígitos (Ex: 2024).")
            self.finalizar_interface()
            pythoncom.CoUninitialize()
            return

        ano_int = int(ano_alvo)
        ano_atual = datetime.now().year
        if ano_int < 1990 or ano_int > ano_atual:
            self.exibir_aviso("Aviso", f"Digite um ano entre 1990 e {ano_atual}.")
            self.finalizar_interface()
            pythoncom.CoUninitialize()
            return
        if not os.path.isdir(self.pasta_destino_pst):
            self.exibir_aviso("Aviso", "A pasta de destino selecionada não existe mais. Escolha outra pasta.")
            self.finalizar_interface()
            pythoncom.CoUninitialize()
            return

        documentos_dir = self.pasta_destino_pst
        sufixo_nome = email_alvo.replace('@', '_').replace('.', '_').replace(' ', '_')
        nome_arquivo_pst = f"Arquivados_{ano_alvo}_{sufixo_nome}.pst"
        caminho_pst = os.path.join(documentos_dir, nome_arquivo_pst)
        nome_pasta_pst = f"Arquivados_{ano_alvo}"

        pasta_raiz_pst = None

        try:
            self.atualizar_status("Conectando ao Outlook...")
            self.log("Conectando ao Outlook...")

            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                self.log(f"Erro: não foi possível conectar ao Outlook ({e}).", "erro")
                self.log("Verifique se o Outlook está instalado e aberto.", "erro")
                self.exibir_erro("Erro", "Não foi possível conectar ao Outlook.\nVerifique se ele está instalado e aberto.")
                return

            pasta_raiz_conta, pasta_origem, erro = self._localizar_conta_e_pasta(namespace, email_alvo, pasta_escolhida)
            if erro:
                self.log(f"Erro: {erro}", "erro")
                self.atualizar_status("Não foi possível continuar.")
                self.exibir_erro("Erro", erro)
                return

            if self.deve_parar:
                return

            self.log(f"Conta conectada: '{pasta_raiz_conta.Name}'")
            self.log(f"Pasta localizada: '{pasta_origem.Name}' ({pasta_origem.Items.Count} itens)")

            # Filtra antes de criar o PST: se faltar espaço, nada é criado nem movido.
            self.log(f"\nFiltrando e-mails do ano {ano_alvo}...")
            self.atualizar_status("Buscando mensagens...")

            emails_para_mover = self._filtrar_emails_por_ano(pasta_origem, pasta_escolhida, ano_alvo)
            total_emails = len(emails_para_mover)
            self.log(f"Total de mensagens identificadas para transferência: {total_emails}")

            if total_emails == 0:
                self.log(f"Nenhum e-mail de {ano_alvo} foi encontrado nesta pasta.", "aviso")
                self.atualizar_status("Processo Concluído. 0 e-mails movidos.")
                self.exibir_info("Concluído", f"Nenhum e-mail do ano {ano_alvo} foi encontrado para mover.")
                return
            if self.deve_parar:
                return

            caminho_absoluto_pst = os.path.abspath(caminho_pst)
            tamanho_total = 0
            for item in emails_para_mover:
                try:
                    tamanho_total += item.Size or 0
                except Exception:
                    pass
            pst_existente = os.path.getsize(caminho_absoluto_pst) if os.path.exists(caminho_absoluto_pst) else 0
            livre = shutil.disk_usage(documentos_dir).free
            gb = 1024 ** 3
            self.log(f"Tamanho a mover: {tamanho_total / gb:.2f} GB. Livre no destino: {livre / gb:.1f} GB.")
            # PST ocupa mais que a soma das mensagens; 1,5x cobre o overhead.
            if livre < tamanho_total * 1.5:
                self.log("Espaço insuficiente no destino. Nada foi movido.", "erro")
                self.atualizar_status("Espaço insuficiente.")
                self.exibir_erro(
                    "Espaço insuficiente",
                    f"Os e-mails somam {tamanho_total / gb:.2f} GB e o destino tem {livre / gb:.1f} GB livres "
                    f"(é preciso ao menos {tamanho_total * 1.5 / gb:.1f} GB).\n\nNada foi movido. Libere espaço ou escolha outra pasta."
                )
                return
            if pst_existente + tamanho_total > LIMITE_PST_BYTES:
                self.log("O PST passaria do limite seguro de 45 GB. Nada foi movido.", "erro")
                self.atualizar_status("PST ficaria grande demais.")
                self.exibir_erro(
                    "PST grande demais",
                    f"O PST ficaria com cerca de {(pst_existente + tamanho_total) / gb:.1f} GB. Acima de 50 GB o "
                    "Outlook para de gravar e o arquivo pode corromper.\n\nNada foi movido. Arquive a Caixa de "
                    "Entrada e os Itens Enviados separadamente ou use outra pasta de destino."
                )
                return

            if pst_existente:
                self.log(f"Aviso: já existe um arquivo PST para esta conta/ano em '{caminho_absoluto_pst}'. Os e-mails serão adicionados a ele.", "aviso")

            self.atualizar_status("Criando/Abrindo arquivo PST de destino...")
            try:
                namespace.AddStore(caminho_absoluto_pst)
            except Exception as e_addstore:
                self.log(f"Erro ao chamar AddStore: {e_addstore}", "erro")
                self.exibir_erro(
                    "Erro",
                    "Não foi possível anexar o arquivo PST ao Outlook.\n\n"
                    "Isso pode acontecer se o arquivo estiver aberto/em uso por outro "
                    "programa, corrompido, ou se o caminho não for acessível.\n\n"
                    f"Caminho: {caminho_absoluto_pst}\n\nDetalhe: {e_addstore}"
                )
                return

            alvo_normalizado = os.path.normcase(os.path.normpath(caminho_absoluto_pst))
            for tentativa in range(10):
                for folder in namespace.Folders:
                    try:
                        if folder.Store:
                            caminho_folder = os.path.normcase(os.path.normpath(folder.Store.FilePath))
                            if caminho_folder == alvo_normalizado:
                                pasta_raiz_pst = folder
                                break
                    except Exception:
                        continue
                if pasta_raiz_pst or self.deve_parar:
                    break
                time.sleep(0.5)

            if not pasta_raiz_pst:
                self.log("Erro: Não foi possível localizar o arquivo PST após criá-lo/abri-lo.", "erro")
                self.exibir_erro(
                    "Erro",
                    "O Outlook aceitou o arquivo PST, mas ele não apareceu a tempo na lista de pastas.\n\n"
                    "Tente novamente. Se o erro persistir, feche o Outlook completamente, "
                    "abra-o de novo e repita o arquivamento."
                )
                return

            try:
                pasta_destino = pasta_raiz_pst.Folders(nome_pasta_pst)
            except Exception:
                pasta_destino = pasta_raiz_pst.Folders.Add(nome_pasta_pst)

            if self.deve_parar:
                return

            self.atualizar_status("Movendo mensagens...")
            contador = 0
            erros_no_item = 0
            itens_antes = pasta_destino.Items.Count

            for item in emails_para_mover:
                if self.deve_parar:
                    self.log("\n[!] Processo interrompido pelo usuário no meio da execução.", "aviso")
                    break
                try:
                    contador += 1
                    assunto = item.Subject if item.Subject else "(sem assunto)"
                    assunto_limpo = assunto[:40]
                    self.log(f"[{contador}/{total_emails}] Transferindo: {assunto_limpo}")
                    self.atualizar_progresso(contador / total_emails)
                    item.Move(pasta_destino)
                except Exception as e_item:
                    erros_no_item += 1
                    self.log(f"Erro no e-mail {contador}: {e_item}", "erro")
                    continue

            # Conferência: o que chegou ao PST tem de bater com o que saiu da caixa.
            movidos = contador - erros_no_item
            try:
                chegaram = pasta_destino.Items.Count - itens_antes
            except Exception:
                chegaram = None
            conferido = chegaram == movidos
            conferencia = (
                f"Conferência: {movidos} movido(s), {chegaram} encontrado(s) no PST."
                if chegaram is not None else "Conferência: não foi possível contar os itens do PST."
            )
            self.log(conferencia, "sucesso" if conferido else "erro")
            alerta_conferencia = "" if conferido else (
                f"\n\n{conferencia}\nOs números não batem: abra o PST no Outlook e confira antes de liberar o usuário."
            )

            if self.deve_parar:
                self.atualizar_status("Cancelado.")
                self.exibir_aviso(
                    "Interrompido",
                    f"O processo foi cancelado. {movidos} e-mail(s) foram movidos antes de parar.{alerta_conferencia}"
                )
                if contador > 0:
                    self.ultima_pasta_pst = documentos_dir
                    self.after(0, lambda: self.btn_abrir_pasta.configure(state="normal"))
            else:
                if erros_no_item == 0 and conferido:
                    self.log("\nSucesso! Todo o histórico foi transferido com êxito.", "sucesso")
                    self.exibir_info("Sucesso", f"Arquivamento concluído com sucesso!\n{movidos} mensagens foram movidas para o PST.\nLocal: {caminho_pst}")
                else:
                    self.log(f"\nConcluído com {erros_no_item} erro(s) em itens individuais.", "aviso")
                    self.exibir_aviso(
                        "Concluído com avisos",
                        f"{movidos} mensagens movidas.\n{erros_no_item} tiveram erro (veja o log).\n"
                        f"Local: {caminho_pst}{alerta_conferencia}"
                    )
                self.atualizar_status("Concluído.")
                self.ultima_pasta_pst = documentos_dir
                self.after(0, lambda: self.btn_abrir_pasta.configure(state="normal"))

        except pythoncom.com_error as e:
            self.log(f"Erro de comunicação com o Outlook: {e}", "erro")
            self.atualizar_status("Perda de conexão com o Outlook.")
            self.exibir_erro(
                "Erro de conexão com o Outlook",
                "A comunicação com o Outlook foi interrompida.\n\n"
                "Isso geralmente acontece se o Outlook foi fechado, travou, ou "
                "está exibindo uma caixa de diálogo de confirmação de segurança.\n\n"
                "Verifique o Outlook e tente novamente."
            )
        except Exception as e:
            self.log(f"Erro geral no sistema: {e}", "erro")
            self.atualizar_status("Ocorreu um erro.")
            self.exibir_erro("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
        finally:
            try:
                if pasta_raiz_pst is not None and not manter_no_outlook:
                    namespace.RemoveStore(pasta_raiz_pst)
            except Exception:
                pass
            try:
                caminho_log_salvo = self.salvar_log_em_arquivo(documentos_dir, f"{ano_alvo}_{sufixo_nome}")
                if caminho_log_salvo:
                    self.log(f"Log saved em: {caminho_log_salvo}")
            except Exception:
                pass
            self.finalizar_interface()
            pythoncom.CoUninitialize()

    def finalizar_interface(self):
        self.after(0, self._finalizar_interface_seguro)

    def _finalizar_interface_seguro(self):
        self.btn_iniciar.configure(state="normal")
        self.btn_cancelar.configure(state="disabled")
        self.txt_email.configure(state="normal")
        self.txt_ano.configure(state="normal")
        self.rb_inbox.configure(state="normal")
        self.rb_sent.configure(state="normal")
        self.btn_procurar.configure(state="normal")
        self.chk_remover_pst.configure(state="normal")
        self.btn_verificar.configure(state="normal")

    # ==================================================================
    # ABA: LIMPEZA DE E-MAIL (informativos / newsletters)
    # ==================================================================

    def _montar_aba_limpeza_email(self, aba):
        self._grupos_informativos = {}
        frame = ctk.CTkFrame(aba, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=2)

        tema.apoio(
            frame,
            "Procura na Caixa de Entrada da conta escolhida em Conta e PST: newsletters, avisos automáticos e "
            "qualquer mensagem acima do tamanho abaixo. Os resultados vêm agrupados por remetente e nada é "
            "apagado sem a sua confirmação."
        )
        linha = ctk.CTkFrame(frame, fg_color="transparent")
        linha.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(linha, text="Mensagem grande a partir de", font=tema.TEXTO).pack(side="left")
        self.txt_tamanho_minimo_mb = tema.entrada(linha, "5", width=60)
        self.txt_tamanho_minimo_mb.pack(side="left", padx=6)
        ctk.CTkLabel(linha, text="MB", font=tema.TEXTO).pack(side="left")
        self.btn_buscar_informativos = tema.botao_primario(linha, "Procurar", self.acao_buscar_informativos, width=130)
        self.btn_buscar_informativos.pack(side="left", padx=(16, 0))

        self.frame_resultados_email = tema.lista_marcavel(frame, "Remetentes encontrados", altura=260)
        self.frame_resultados_email.pack(fill="both", expand=True)
        tema.vazio(self.frame_resultados_email, "Clique em Procurar para listar os remetentes.")

        frame_acoes = ctk.CTkFrame(frame, fg_color="transparent")
        frame_acoes.pack(fill="x", pady=(10, 0))
        self.btn_marcar_todos = tema.botao_link(frame_acoes, "Marcar todos", lambda: self._marcar_todos_informativos(True), width=0)
        self.btn_marcar_todos.pack(side="left")
        self.btn_desmarcar_todos = tema.botao_link(frame_acoes, "Desmarcar todos", lambda: self._marcar_todos_informativos(False), width=0)
        self.btn_desmarcar_todos.pack(side="left", padx=(4, 0))
        self.btn_mover_lixeira_informativos = tema.botao_perigo(
            frame_acoes, "Mover para a Lixeira", self.acao_mover_informativos_lixeira, width=190, state="disabled"
        )
        self.btn_mover_lixeira_informativos.pack(side="right")

    def acao_buscar_informativos(self):
        email_alvo = self.txt_email.get().strip()
        if not email_alvo or email_alvo == "Selecione ou digite o e-mail" or "Nenhuma conta detectada" in email_alvo:
            messagebox.showwarning("Aviso", "Escolha a conta em Conta e PST primeiro.")
            return

        tamanho_texto = self.txt_tamanho_minimo_mb.get().strip().replace(",", ".")
        try:
            tamanho_minimo_mb = float(tamanho_texto) if tamanho_texto else 5.0
        except ValueError:
            messagebox.showwarning("Aviso", "Digite um número válido para o tamanho mínimo (ex: 5 ou 2.5).")
            return

        self.btn_buscar_informativos.configure(state="disabled", text="Buscando...")
        self.btn_mover_lixeira_informativos.configure(state="disabled")
        threading.Thread(
            target=self._processo_buscar_informativos, args=(email_alvo, tamanho_minimo_mb), daemon=True
        ).start()

    def _formatar_tamanho(self, tamanho_bytes):
        if tamanho_bytes >= 1024 * 1024:
            return f"{tamanho_bytes / (1024 * 1024):.1f} MB"
        return f"{tamanho_bytes / 1024:.0f} KB"

    def _classificar_informativo(self, item, remetente_nome, remetente_email, assunto, tamanho_minimo_bytes):
        """Retorna 'newsletter', 'palavra', 'grande' ou None, do critério mais confiável ao menos.

        Só 'newsletter' (cabeçalho List-Unsubscribe) é sinal forte. Palavra-chave
        ("alerta", "comunicado"...) e tamanho pegam muito e-mail de trabalho, como
        contrato em PDF de cliente, por isso esses grupos vêm desmarcados.
        """
        try:
            cabecalhos = item.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
            )
            if cabecalhos and "list-unsubscribe" in cabecalhos.lower():
                return "newsletter"
        except Exception:
            pass
        texto_busca = f"{remetente_nome} {remetente_email} {assunto}".lower()
        if any(palavra in texto_busca for palavra in self.PALAVRAS_CHAVE_INFORMATIVO):
            return "palavra"
        try:
            if item.Size and item.Size >= tamanho_minimo_bytes:
                return "grande"
        except Exception:
            pass
        return None

    def _processo_buscar_informativos(self, email_alvo, tamanho_minimo_mb):
        pythoncom.CoInitialize()
        tamanho_minimo_bytes = int(tamanho_minimo_mb * 1024 * 1024)
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                self.exibir_erro("Erro", f"Não foi possível conectar ao Outlook.\nVerifique se ele está instalado e aberto.\n\n{e}")
                return

            _, pasta_origem, erro = self._localizar_conta_e_pasta(namespace, email_alvo, "Inbox")
            if erro:
                self.exibir_erro("Erro", erro)
                return

            total_na_pasta = "?"
            try:
                total_na_pasta = pasta_origem.Items.Count
                self.log_geral(f"A Caixa de Entrada tem {total_na_pasta} item(ns) no total (e-mails + outros tipos como convites, avisos de leitura etc.).")
            except Exception:
                pass

            grupos = {}
            total_analisados = 0
            store_id = pasta_origem.StoreID
            self.atualizar_status_ferramentas("Analisando mensagens da Caixa de Entrada...")

            for item in pasta_origem.Items:
                try:
                    if not (item and hasattr(item, "Class") and item.Class == 43):
                        continue
                    total_analisados += 1
                    if total_analisados == 1 or total_analisados % 100 == 0:
                        progresso = f"Analisando e-mails: {total_analisados} de {total_na_pasta} item(ns)..."
                        self.atualizar_status_ferramentas(progresso)
                        self.log_geral(progresso)

                    remetente_nome = item.SenderName or "(desconhecido)"
                    try:
                        remetente_email = item.SenderEmailAddress or ""
                    except Exception:
                        remetente_email = ""
                    assunto = item.Subject or "(sem assunto)"
                    try:
                        tamanho_item = item.Size or 0
                    except Exception:
                        tamanho_item = 0

                    categoria = self._classificar_informativo(
                        item, remetente_nome, remetente_email, assunto, tamanho_minimo_bytes
                    )
                    if not categoria:
                        continue

                    # Um remetente pode ter newsletter e anexos grandes: cada critério vira um grupo.
                    chave = (categoria, remetente_email.lower() if remetente_email else remetente_nome.lower())
                    if chave not in grupos:
                        grupos[chave] = {
                            "nome": remetente_nome, "email": remetente_email, "itens": [],
                            "assuntos": [], "tamanho_total": 0, "store_id": store_id,
                            "categoria": categoria,
                        }
                    # Guardamos o EntryID (texto), não o objeto COM: o objeto
                    # pertence a esta thread e deixa de valer após o
                    # CoUninitialize, o que fazia todos os "Mover" falharem.
                    grupos[chave]["itens"].append(item.EntryID)
                    grupos[chave]["tamanho_total"] += tamanho_item
                    if len(grupos[chave]["assuntos"]) < 3:
                        rotulo_assunto = assunto[:50]
                        if tamanho_item >= tamanho_minimo_bytes:
                            rotulo_assunto += f" ({self._formatar_tamanho(tamanho_item)})"
                        grupos[chave]["assuntos"].append(rotulo_assunto)
                except Exception:
                    continue

            encontrados = sum(len(grupo["itens"]) for grupo in grupos.values())
            self.log_geral(
                f"Análise concluída: {total_analisados} e-mail(s) verificado(s), "
                f"{encontrados} informativo(s) encontrado(s).",
                "sucesso",
            )
            self.atualizar_status_ferramentas("Análise de e-mails concluída.")
            self.after(0, lambda: self._exibir_resultados_informativos(grupos, total_analisados))
        except Exception as e:
            self.exibir_erro("Erro", f"Erro ao buscar e-mails: {e}")
        finally:
            self.after(0, lambda: self.btn_buscar_informativos.configure(state="normal", text="Procurar"))
            pythoncom.CoUninitialize()

    def _exibir_resultados_informativos(self, grupos, total_analisados):
        for widget in self.frame_resultados_email.winfo_children():
            widget.destroy()

        self._grupos_informativos = {}

        if not grupos:
            ctk.CTkLabel(
                self.frame_resultados_email,
                text=f"Nenhum e-mail informativo/newsletter encontrado ({total_analisados} e-mail(s) analisado(s)).",
                font=("Segoe UI", 11, "italic"), text_color=("gray40", "gray70")
            ).pack(pady=20)
            self.btn_mover_lixeira_informativos.configure(state="disabled")
            return

        total_encontrados = sum(len(g["itens"]) for g in grupos.values())
        tamanho_total_geral = sum(g.get("tamanho_total", 0) for g in grupos.values())
        ctk.CTkLabel(
            self.frame_resultados_email,
            text=f"{len(grupos)} grupo(s) — {total_encontrados} e-mail(s) de {total_analisados} analisado(s) "
                 f"— {self._formatar_tamanho(tamanho_total_geral)} no total\n"
                 "Só as newsletters (com link de descadastro) vêm marcadas. Revise as demais antes de marcar.",
            font=("Segoe UI", 11, "bold"), justify="left"
        ).pack(anchor="w", padx=5, pady=(0, 8))

        for chave, dados in sorted(
            grupos.items(),
            key=lambda kv: (ORDEM_CATEGORIA[kv[1]["categoria"]], -kv[1].get("tamanho_total", 0)),
        ):
            var = ctk.BooleanVar(value=dados["categoria"] == "newsletter")
            texto_remetente = f'{dados["nome"]} <{dados["email"]}>' if dados["email"] else dados["nome"]
            texto_assuntos = " | ".join(dados["assuntos"])
            tamanho_formatado = self._formatar_tamanho(dados.get("tamanho_total", 0))
            tema.marcavel(
                self.frame_resultados_email,
                f'[{ROTULO_CATEGORIA[dados["categoria"]]}] {texto_remetente}\n'
                f'{len(dados["itens"])} e-mail(s), {tamanho_formatado}. {texto_assuntos}',
                var
            ).pack(anchor="w", padx=5, pady=4, fill="x")
            self._grupos_informativos[chave] = {
                "var": var, "itens": dados["itens"], "store_id": dados["store_id"], "categoria": dados["categoria"],
            }

        self.btn_mover_lixeira_informativos.configure(state="normal")

    def _marcar_todos_informativos(self, valor):
        for dados in self._grupos_informativos.values():
            dados["var"].set(valor)

    def acao_mover_informativos_lixeira(self):
        selecionados = {k: v for k, v in self._grupos_informativos.items() if v["var"].get()}
        if not selecionados:
            messagebox.showwarning("Aviso", "Nenhum remetente selecionado.")
            return

        total = sum(len(v["itens"]) for v in selecionados.values())
        incertos = sum(len(v["itens"]) for v in selecionados.values() if v["categoria"] != "newsletter")
        aviso_incertos = (
            f"\n\nATENÇÃO: {incertos} desses e-mail(s) foram encontrados só por palavra-chave ou tamanho "
            "e podem ser mensagens de trabalho."
        ) if incertos else ""
        if not messagebox.askyesno(
            "Confirmar",
            f"Isso vai mover {total} e-mail(s) de {len(selecionados)} grupo(s) para a Lixeira do "
            "Outlook (Itens Excluídos). Dá para recuperá-los de lá, a menos que uma política de "
            f"retenção da empresa esvazie essa pasta automaticamente.{aviso_incertos}\n\nContinuar?",
            icon="warning" if incertos else "question",
        ):
            return

        self.btn_mover_lixeira_informativos.configure(state="disabled", text="Movendo...")
        threading.Thread(
            target=self._processo_mover_informativos, args=(selecionados, self.txt_email.get().strip()), daemon=True
        ).start()

    def _processo_mover_informativos(self, selecionados, email_alvo):
        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                self.exibir_erro("Erro", f"Não foi possível conectar ao Outlook.\n\n{e}")
                return

            pasta_raiz_conta, _, erro = self._localizar_conta_e_pasta(namespace, email_alvo, "Inbox")
            if erro:
                self.exibir_erro("Erro", erro)
                return

            # 3 = olFolderDeletedItems. A busca anterior por nome incluía
            # "lixo eletrônico" e podia mandar os e-mails para a pasta de Spam.
            pasta_lixeira = None
            try:
                pasta_lixeira = pasta_raiz_conta.Store.GetDefaultFolder(3)
            except Exception:
                nomes_lixeira = ("itens excluídos", "deleted items", "lixeira")
                for subfolder in pasta_raiz_conta.Folders:
                    try:
                        if subfolder.Name.lower() in nomes_lixeira:
                            pasta_lixeira = subfolder
                            break
                    except Exception:
                        continue

            if not pasta_lixeira:
                self.exibir_erro("Erro", "Não foi possível localizar a pasta 'Itens Excluídos' desta conta.")
                return

            movidos = 0
            erros = 0
            total = sum(len(dados["itens"]) for dados in selecionados.values())
            self.log_geral(f"Movendo {total} e-mail(s) selecionado(s) para a Lixeira...")
            self.atualizar_status_ferramentas(f"Movendo e-mails: 0 de {total}...")
            for dados in selecionados.values():
                for entry_id in dados["itens"]:
                    try:
                        namespace.GetItemFromID(entry_id, dados["store_id"]).Move(pasta_lixeira)
                        movidos += 1
                    except Exception:
                        erros += 1
                    processados = movidos + erros
                    if processados == total or processados % 25 == 0:
                        progresso = f"Movendo e-mails: {processados} de {total}..."
                        self.atualizar_status_ferramentas(progresso)
                        self.log_geral(progresso)

            mensagem = f"{movidos} e-mail(s) movido(s) para a Lixeira."
            if erros:
                mensagem += f"\n{erros} e-mail(s) tiveram erro ao mover."
            self.log_geral(mensagem.replace("\n", " "), "aviso" if erros else "sucesso")
            self.atualizar_status_ferramentas("Movimentação de e-mails concluída.")
            self.exibir_info("Concluído", mensagem)
            self.after(0, lambda: self._exibir_resultados_informativos({}, 0))
        except Exception as e:
            self.exibir_erro("Erro", f"Erro ao mover e-mails: {e}")
        finally:
            self.after(
                0,
                lambda: self.btn_mover_lixeira_informativos.configure(
                    state="disabled", text="Mover para a Lixeira"
                )
            )
            pythoncom.CoUninitialize()
