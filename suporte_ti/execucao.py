"""Log da sessão, execução de processos e elevação (UAC)."""

import os
import queue
import re
import subprocess
import tempfile
import threading
import time
import uuid
from datetime import datetime
from tkinter import messagebox

from . import tema
from .comum import (
    CODIGO_UAC_CANCELADO, CREATE_NO_WINDOW, TIMEOUT_PADRAO_ELEVADO,
    caminho_log_ferramentas_do_dia, ps_aspas,
)


class ExecucaoMixin:
    # ==================================================================
    # LOG / STATUS GENÉRICOS DAS FERRAMENTAS (thread-safe)
    # ==================================================================

    def log_geral(self, mensagem, tipo="info"):
        self.after(0, self._log_geral_seguro, mensagem, tipo)

    def _log_geral_seguro(self, mensagem, tipo="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        linha = f"{timestamp}  {mensagem}\n"
        cores = tema.CORES_LOG
        try:
            if tipo in cores:
                if tipo not in self._tags_configuradas_ferramentas:
                    self.txt_log_ferramentas.tag_config(tipo, foreground=cores[tipo])
                    self._tags_configuradas_ferramentas.add(tipo)
                self.txt_log_ferramentas.insert("end", linha, tipo)
            else:
                self.txt_log_ferramentas.insert("end", linha)
        except Exception:
            self.txt_log_ferramentas.insert("end", linha)
        self.txt_log_ferramentas.see("end")

        # Persiste cada linha em disco (um arquivo por dia), para servir de
        # trilha de auditoria do que foi executado nesta máquina.
        try:
            with open(caminho_log_ferramentas_do_dia(), "a", encoding="utf-8") as f:
                prefixo = f"[{tipo.upper()}] " if tipo in cores else ""
                f.write(f"[{timestamp}] {prefixo}{mensagem}\n")
        except Exception:
            pass

    def _copiar_log(self):
        conteudo = self.txt_log_ferramentas.get("1.0", "end").strip()
        self.clipboard_clear()
        self.clipboard_append(conteudo)
        self.atualizar_status_ferramentas("Log copiado para a área de transferência.")

    def _limpar_log_geral(self):
        self.txt_log_ferramentas.delete("1.0", "end")
        self.atualizar_status_ferramentas("Registro limpo.")

    def atualizar_status_ferramentas(self, texto):
        def atualizar():
            self.lbl_status_ferramentas.configure(text=texto)
            if any(termo in texto.lower() for termo in ("erro", "falha", "expirou")):
                self.lbl_estado_atividade.configure(text="Atenção", text_color=tema.CORES_LOG["erro"])
        self.after(0, atualizar)

    def _executar_com_saida_em_tempo_real(self, comando, timeout, rotulo="Comando", encoding="oem"):
        """Executa um comando e envia cada linha de saída ao log da interface.

        Ferramentas de console do Windows (ipconfig, gpupdate, PowerShell 5.1)
        escrevem na página de código OEM (cp850 no Brasil); o winget escreve em
        UTF-8. Decodificar tudo como UTF-8 deixava os acentos corrompidos.
        """
        processo = subprocess.Popen(
            comando,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding=encoding,
            errors="replace",
            bufsize=1,
            creationflags=CREATE_NO_WINDOW,
        )
        self._processo_ferramenta_ativo = processo
        try:
            return self._acompanhar_saida_processo(processo, comando, timeout, rotulo)
        finally:
            self._processo_ferramenta_ativo = None
            if processo.poll() is None:
                try:
                    processo.kill()
                except OSError:
                    pass

    def _acompanhar_saida_processo(self, processo, comando, timeout, rotulo):
        linhas = queue.Queue()

        def ler_saida():
            try:
                for linha in iter(processo.stdout.readline, ""):
                    linhas.put(linha)
            finally:
                linhas.put(None)

        threading.Thread(target=ler_saida, daemon=True).start()
        inicio = time.monotonic()
        ultima_linha = None
        leitor_finalizado = False

        while not leitor_finalizado or processo.poll() is None:
            if self.__dict__.get("_cancelar_ferramenta", False):
                processo.terminate()
                try:
                    processo.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    processo.kill()
                raise RuntimeError(f"{rotulo} cancelado pelo usuário.")
            if time.monotonic() - inicio > timeout:
                processo.kill()
                processo.wait()
                raise subprocess.TimeoutExpired(comando, timeout)

            try:
                linha = linhas.get(timeout=0.2)
            except queue.Empty:
                continue

            if linha is None:
                leitor_finalizado = True
                continue

            # O winget usa sequências ANSI e atualizações com \r. Retiramos esses
            # controles para manter o log legível e evitamos linhas repetidas.
            linha = re.sub(r"\x1b\[[0-?]*[ -/]*[@-~]", "", linha).strip()
            if not linha or linha == ultima_linha:
                continue
            ultima_linha = linha
            self.log_geral(f"{rotulo}: {linha}")
            resumo = linha if len(linha) <= 110 else f"{linha[:107]}..."
            self.atualizar_status_ferramentas(f"{rotulo}: {resumo}")

        return processo.wait()

    def _iniciar_progresso_ferramentas(self):
        """Liga a barra de progresso em modo indeterminado (uso: operações
        longas como SFC, winget, reparo de rede, etc)."""
        self.after(0, lambda: self.progress_ferramentas.configure(mode="indeterminate"))
        self.after(0, self.progress_ferramentas.start)
        self.after(0, lambda: self.lbl_estado_atividade.configure(text="Em execução", text_color=tema.REG_LINK))

    def _parar_progresso_ferramentas(self):
        self.after(0, self.progress_ferramentas.stop)
        self.after(0, lambda: self.progress_ferramentas.configure(mode="determinate"))
        self.after(0, lambda: self.progress_ferramentas.set(0))
        self.after(0, lambda: self.lbl_estado_atividade.configure(text="Pronto", text_color=tema.REG_APOIO))

    def exibir_aviso(self, titulo, mensagem):
        self.after(0, lambda: messagebox.showwarning(titulo, message=mensagem))

    def exibir_erro(self, titulo, mensagem):
        self.after(0, lambda: messagebox.showerror(titulo, message=mensagem))

    def exibir_info(self, titulo, mensagem):
        self.after(0, lambda: messagebox.showinfo(titulo, message=mensagem))

    def _disparar_acao(self, funcao_execucao, mensagem_confirmacao=None, titulo_confirmacao="Confirmar ação"):
        """Confirma (se necessário) na thread principal e roda a ação numa thread separada."""
        if mensagem_confirmacao:
            if not messagebox.askyesno(titulo_confirmacao, mensagem_confirmacao):
                return
        self._cancelar_ferramenta = False
        threading.Thread(target=funcao_execucao, daemon=True).start()

    def _executar_elevado(self, comando_cmd, timeout=TIMEOUT_PADRAO_ELEVADO):
        """Executa um comando do cmd como administrador (UAC) e mostra a saída no log.

        Levanta PermissionError se o UAC for negado/cancelado e TimeoutError se
        o comando não terminar dentro de `timeout` segundos.
        """
        caminho_saida = os.path.join(tempfile.gettempdir(), f"saida_ti_{uuid.uuid4().hex}.txt")
        argumento_cmd = f'/c {comando_cmd} > "{caminho_saida}" 2>&1'
        # Sem o try/catch, cancelar o UAC terminava com código 0 e a ação era
        # registrada como "concluída com sucesso".
        ps_args = (
            f"try {{ $p = Start-Process cmd -ArgumentList {ps_aspas(argumento_cmd)} "
            "-Verb RunAs -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop; exit $p.ExitCode } "
            f"catch {{ exit {CODIGO_UAC_CANCELADO} }}"
        )
        self.log_geral("Aguardando autorização de administrador (UAC)...", "aviso")
        self.atualizar_status_ferramentas("Aguardando confirmação do UAC...")
        try:
            try:
                codigo = self._acompanhar_processo_com_arquivo(
                    ["powershell", "-NoProfile", "-Command", ps_args], caminho_saida, timeout,
                    "Administrador", aviso_inicial=False
                )
            except subprocess.TimeoutExpired:
                raise TimeoutError(
                    f"O comando não respondeu em {timeout // 60} minuto(s). "
                    "Verifique se não ficou uma janela de confirmação (UAC) pendente na tela."
                )
            if codigo == CODIGO_UAC_CANCELADO and not os.path.exists(caminho_saida):
                raise PermissionError("A permissão de administrador (UAC) foi negada ou cancelada.")
            if codigo != 0:
                self.log_geral(f"O comando terminou com código {codigo}.", "aviso")
            return subprocess.CompletedProcess(comando_cmd, codigo)
        finally:
            try:
                os.remove(caminho_saida)
            except OSError:
                pass

    def _ler_novas_linhas_arquivo(self, caminho, posicao, resto, rotulo, finalizar=False, encoding="cp850"):
        """Publica no log apenas o trecho novo de um arquivo de saída."""
        try:
            with open(caminho, "r", encoding=encoding, errors="replace") as arquivo:
                arquivo.seek(posicao)
                trecho = arquivo.read()
                posicao = arquivo.tell()
        except OSError:
            trecho = ""

        partes = re.split(r"[\r\n]+", resto + trecho)
        resto = "" if finalizar else partes.pop()
        for linha in partes + ([resto] if finalizar and resto else []):
            linha = re.sub(r"\x1b\[[0-?]*[ -/]*[@-~]", "", linha).replace("﻿", "").strip()
            if linha:
                self.log_geral(f"{rotulo}: {linha}")
                resumo = linha if len(linha) <= 110 else f"{linha[:107]}..."
                self.atualizar_status_ferramentas(f"{rotulo}: {resumo}")
        return posicao, resto

    def _acompanhar_processo_com_arquivo(self, comando, caminho_saida, timeout, rotulo,
                                         encoding="cp850", aviso_inicial=True):
        """Acompanha um processo que grava a saída em arquivo (inclusive via UAC)."""
        if aviso_inicial:
            self.log_geral(f"{rotulo}: aguardando confirmação do UAC...", "aviso")
        processo = subprocess.Popen(comando, creationflags=CREATE_NO_WINDOW)
        inicio = time.monotonic()
        proximo_aviso = inicio + 10
        posicao = 0
        resto = ""
        while processo.poll() is None:
            agora = time.monotonic()
            if agora - inicio > timeout:
                processo.kill()
                processo.wait()
                raise subprocess.TimeoutExpired(comando, timeout)
            posicao, resto = self._ler_novas_linhas_arquivo(
                caminho_saida, posicao, resto, rotulo, encoding=encoding
            )
            if agora >= proximo_aviso:
                self.log_geral(f"{rotulo}: operação ainda em execução ({int(agora - inicio)}s)...")
                proximo_aviso = agora + 10
            time.sleep(0.25)
        self._ler_novas_linhas_arquivo(
            caminho_saida, posicao, resto, rotulo, finalizar=True, encoding=encoding
        )
        return processo.returncode

    def _executar_script_ps_elevado(self, corpo_script, rotulo, timeout=TIMEOUT_PADRAO_ELEVADO):
        """Executa um script PowerShell como administrador e devolve (codigo, saida).

        O próprio script grava a saída no arquivo. A versão anterior usava
        `-File script.ps1 *> saida.txt`, mas com -File tudo depois do caminho é
        repassado ao script como argumento: o arquivo de saída nunca era criado
        e toda operação (impressora, DNS, domínio, senha) era exibida como falha.

        Levanta PermissionError se o UAC for negado e subprocess.TimeoutExpired
        se o tempo esgotar.
        """
        pasta_temp = tempfile.gettempdir()
        identificador = uuid.uuid4().hex
        caminho_script = os.path.join(pasta_temp, f"ti_script_{identificador}.ps1")
        caminho_saida = os.path.join(pasta_temp, f"ti_saida_{identificador}.txt")

        # Erros não tratados viram uma linha "ERRO: mensagem" (a formatação
        # padrão do PowerShell repetiria o código do script inteiro no log).
        script = (
            "& {\n" + corpo_script.strip("\n") + "\n} *>&1 | ForEach-Object {\n"
            "    if ($_ -is [System.Management.Automation.ErrorRecord]) { 'ERRO: ' + $_.Exception.Message } else { $_ }\n"
            "} | Out-File -LiteralPath " + ps_aspas(caminho_saida) + " -Encoding utf8 -Width 4096\n"
        )
        argumentos = f'-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{caminho_script}"'
        ps_args = (
            f"try {{ $p = Start-Process powershell -ArgumentList {ps_aspas(argumentos)} "
            "-Verb RunAs -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop; exit $p.ExitCode } "
            f"catch {{ exit {CODIGO_UAC_CANCELADO} }}"
        )
        try:
            # utf-8-sig: o PowerShell 5.1 lê .ps1 sem BOM como ANSI.
            with open(caminho_script, "w", encoding="utf-8-sig") as arquivo:
                arquivo.write(script)
            codigo = self._acompanhar_processo_com_arquivo(
                ["powershell", "-NoProfile", "-Command", ps_args], caminho_saida,
                timeout, rotulo, encoding="utf-8"
            )
            try:
                with open(caminho_saida, "r", encoding="utf-8-sig", errors="replace") as arquivo:
                    saida = arquivo.read().strip()
            except OSError:
                saida = ""
            if codigo == CODIGO_UAC_CANCELADO and not saida:
                raise PermissionError("A permissão de administrador (UAC) foi negada ou cancelada.")
            return codigo, saida
        finally:
            for caminho in (caminho_script, caminho_saida):
                try:
                    os.remove(caminho)
                except OSError:
                    pass

    def _botao_ferramenta(self, parent, texto, comando_func, confirmar=None, em_segundo_plano=True):
        """Cria um item de ferramenta.

        em_segundo_plano=True: linha da lista de comandos, roda numa thread.
        em_segundo_plano=False: botão que envia um formulário (lê campos e abre
        diálogos antes de iniciar a própria thread; o Tkinter só pode ser usado
        pela thread principal, e chamá-lo de outra thread pode travar o app).
        """
        if not em_segundo_plano:
            btn = tema.botao_primario(parent, texto, comando_func, width=220)
            btn.pack(anchor="w", pady=(4, 0), padx=2)
            return btn
        return tema.linha_comando(parent, texto, lambda: self._disparar_acao(comando_func, confirmar))
