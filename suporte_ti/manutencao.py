"""Abas Manutenção e Impressoras."""

import ipaddress
import json
import os
import subprocess
import tempfile
import threading
import time
import uuid
from tkinter import messagebox

import customtkinter as ctk

from . import tema
from .comum import (
    CREATE_NEW_CONSOLE, CREATE_NO_WINDOW, TIMEOUT_PADRAO_ELEVADO,
    limpar_conteudo_diretorio, ps_aspas,
)


class ManutencaoMixin:
    # ==================================================================
    # ABA: MANUTENÇÃO
    # ==================================================================

    def _montar_aba_manutencao(self, aba):
        pagina = tema.pagina(aba)

        tema.titulo_grupo(pagina, "Atualizações", primeiro=True)
        lista = tema.lista(pagina)
        self._botao_ferramenta(
            lista, "Atualizar todos os aplicativos (winget upgrade --all)", self.acao_winget_upgrade_all,
            confirmar="Atualizar todos os aplicativos que o Winget gerencia nesta máquina "
                      "(Chrome, leitores de PDF, players...)? Pode levar alguns minutos."
        )

        tema.titulo_grupo(pagina, "Integridade")
        lista = tema.lista(pagina)
        self._botao_ferramenta(lista, "Verificar arquivos do Windows (SFC /scannow)", self.acao_sfc)
        self._botao_ferramenta(lista, "Verificar o disco (CHKDSK, só leitura)", self.acao_chkdsk)
        self._botao_ferramenta(
            lista, "Reinstalar o Google Chrome", self.acao_reinstalar_chrome,
            confirmar="Desinstalar e reinstalar o Google Chrome?"
        )

        tema.titulo_grupo(pagina, "Limpeza")
        lista = tema.lista(pagina)
        self._botao_ferramenta(
            lista, "Apagar arquivos temporários (TEMP)", self.acao_limpar_temp,
            confirmar="Apagar os arquivos temporários do usuário e do Windows?"
        )
        self._botao_ferramenta(
            lista, "Esvaziar a pasta Prefetch", self.acao_limpar_prefetch,
            confirmar="Apagar todo o conteúdo de C:\\Windows\\Prefetch?"
        )
        self._botao_ferramenta(
            lista, "Limpar o cache do Windows Update", self.acao_limpar_cache_wu,
            confirmar="Parar o Windows Update, apagar o cache de downloads e reiniciar o serviço?"
        )
        self._botao_ferramenta(
            lista, "Esvaziar a Lixeira", self.acao_limpar_lixeira,
            confirmar="Esvaziar a Lixeira? Os arquivos não poderão ser recuperados."
        )

        tema.titulo_grupo(pagina, "Windows")
        lista = tema.lista(pagina)
        self._botao_ferramenta(
            lista, "Reiniciar a área de trabalho (explorer.exe)", self.acao_reiniciar_explorer,
            confirmar="Reiniciar a área de trabalho? As janelas do Explorer serão fechadas."
        )

    def _montar_aba_impressoras(self, aba):
        pagina = tema.pagina(aba)
        lista = tema.lista(pagina)
        self._botao_ferramenta(
            lista, "Reiniciar o spooler de impressão", self.acao_reiniciar_spooler,
            confirmar="Parar e reiniciar o serviço de impressão? Trabalhos na fila podem ser perdidos."
        )

        tema.titulo_grupo(pagina, "Instalar impressora por IP")
        tema.apoio(pagina, "O driver precisa já estar instalado no Windows. Use o nome exatamente como aparece na lista de drivers.")
        self.txt_imp_ip = tema.campo(pagina, "Endereço IP", "192.168.1.50", largura=260)
        self.txt_imp_nome = tema.campo(pagina, "Nome da impressora", "HP_Recepcao", largura=360)
        self.txt_imp_driver = tema.campo(pagina, "Driver", "HP LaserJet P2055 PCL6")
        ctk.CTkFrame(pagina, height=8, fg_color="transparent").pack()
        self._botao_ferramenta(pagina, "Instalar impressora", self.acao_instalar_impressora_ip, em_segundo_plano=False)

    def acao_winget_upgrade_all(self):
        self.atualizar_status_ferramentas("Atualizando aplicativos via winget (pode levar vários minutos)...")
        self.log_geral("Executando winget upgrade --all...")
        self._iniciar_progresso_ferramentas()
        try:
            codigo_retorno = self._executar_com_saida_em_tempo_real(
                ["winget", "upgrade", "--all", "--silent",
                 "--accept-source-agreements", "--accept-package-agreements"],
                timeout=TIMEOUT_PADRAO_ELEVADO,
                rotulo="winget", encoding="utf-8",
            )

            if codigo_retorno == 0:
                self.log_geral("Todos os aplicativos foram atualizados com sucesso.", "sucesso")
                self.atualizar_status_ferramentas("Aplicativos atualizados.")
                self.exibir_info("Concluído", "Atualização de aplicativos concluída com sucesso via winget.")
            else:
                self.log_geral(f"winget terminou com código {codigo_retorno} (pode ser normal se não havia nada para atualizar).", "aviso")
                self.atualizar_status_ferramentas("Winget concluído com avisos.")
                self.exibir_info(
                    "Concluído com avisos",
                    "O winget terminou a execução, mas retornou um código diferente de sucesso.\n"
                    "Isso costuma acontecer quando não havia nada para atualizar, ou quando algum "
                    "aplicativo específico não pôde ser atualizado silenciosamente.\n\n"
                    "Veja o log para detalhes."
                )
        except FileNotFoundError:
            self.log_geral("O comando 'winget' não foi encontrado nesta máquina.", "erro")
            self.exibir_erro(
                "Erro",
                "O winget não está instalado ou não está disponível nesta máquina.\n\n"
                "Ele já vem por padrão no Windows 10/11 atualizados (App Installer), mas em "
                "instalações mais antigas pode ser necessário instalar pela Microsoft Store."
            )
        except subprocess.TimeoutExpired:
            self.log_geral(
                f"O winget não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).", "erro"
            )
            self.atualizar_status_ferramentas("Winget expirou (timeout).")
            self.exibir_erro(
                "Tempo esgotado",
                f"A atualização via winget não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).\n\n"
                "Isso pode acontecer se algum instalador individual ficou aguardando confirmação "
                "manual. Verifique a tela e tente novamente."
            )
        except Exception as e:
            self.log_geral(f"Erro ao executar winget upgrade: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            self._parar_progresso_ferramentas()

    def acao_sfc(self):
        self.atualizar_status_ferramentas("Executando SFC /scannow (pode levar vários minutos)...")
        self.log_geral("Iniciando verificação SFC /scannow...")
        self._iniciar_progresso_ferramentas()
        try:
            self._executar_elevado("sfc /scannow")
            self.log_geral("Verificação SFC concluída.", "sucesso")
            self.atualizar_status_ferramentas("SFC concluído.")
            self.exibir_info("Concluído", "Verificação SFC concluída. O resultado completo está no registro, à direita.")
        except Exception as e:
            self.log_geral(f"Erro ao executar SFC: {e}", "erro")
            self.exibir_erro("Erro", f"Erro ao executar SFC: {e}")
        finally:
            self._parar_progresso_ferramentas()

    def acao_limpar_temp(self):
        self.atualizar_status_ferramentas("Limpando pastas temporárias...")
        self.log_geral("Limpando pasta TEMP do usuário...")
        try:
            temp_dir = os.environ.get("TEMP", "")
            if temp_dir:
                limpar_conteudo_diretorio(temp_dir)
            self.log_geral("Limpando C:\\Windows\\Temp (requer administrador)...")
            self._executar_elevado("del /f /s /q C:\\Windows\\Temp\\*")
            self.log_geral("Pastas temporárias limpas.", "sucesso")
            self.atualizar_status_ferramentas("Limpeza de TEMP concluída.")
            self.exibir_info("Concluído", "Pastas temporárias limpas com sucesso.")
        except Exception as e:
            self.log_geral(f"Erro ao limpar TEMP: {e}", "erro")
            self.exibir_erro("Erro", f"Erro ao limpar TEMP: {e}")

    def acao_limpar_prefetch(self):
        self.atualizar_status_ferramentas("Limpando Prefetch...")
        self.log_geral("Limpando C:\\Windows\\Prefetch (requer administrador)...")
        try:
            self._executar_elevado("del /f /s /q C:\\Windows\\Prefetch\\*")
            self.log_geral("Pasta Prefetch limpa.", "sucesso")
            self.atualizar_status_ferramentas("Prefetch limpo.")
            self.exibir_info("Concluído", "Pasta Prefetch limpa com sucesso.")
        except Exception as e:
            self.log_geral(f"Erro ao limpar Prefetch: {e}", "erro")
            self.exibir_erro("Erro", f"Erro ao limpar Prefetch: {e}")

    def acao_limpar_cache_wu(self):
        self.atualizar_status_ferramentas("Limpando cache do Windows Update...")
        self.log_geral("Parando serviço wuauserv, limpando cache e reiniciando o serviço...")
        self._iniciar_progresso_ferramentas()
        try:
            self._executar_elevado("net stop wuauserv & del /f /s /q C:\\Windows\\SoftwareDistribution\\Download\\* & net start wuauserv")
            self.log_geral("Cache do Windows Update limpo.", "sucesso")
            self.atualizar_status_ferramentas("Cache do Windows Update limpo.")
            self.exibir_info("Concluído", "Cache de downloads do Windows Update limpo com sucesso.")
        except Exception as e:
            self.log_geral(f"Erro ao limpar cache do Windows Update: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            self._parar_progresso_ferramentas()

    def acao_limpar_lixeira(self):
        self.atualizar_status_ferramentas("Esvaziando a Lixeira...")
        self.log_geral("Esvaziando a Lixeira do sistema...")
        try:
            codigo = self._executar_com_saida_em_tempo_real(
                ["powershell", "-Command", "Clear-RecycleBin -Force -ErrorAction SilentlyContinue"],
                timeout=120, rotulo="Lixeira"
            )
            if codigo != 0:
                raise RuntimeError(f"O PowerShell terminou com código {codigo}.")
            self.log_geral("Lixeira esvaziada.", "sucesso")
            self.atualizar_status_ferramentas("Lixeira esvaziada.")
            self.exibir_info("Concluído", "Lixeira esvaziada com sucesso.")
        except Exception as e:
            self.log_geral(f"Erro ao esvaziar a Lixeira: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")

    def acao_chkdsk(self):
        self.atualizar_status_ferramentas("Abrindo verificação de disco (CHKDSK)...")
        self.log_geral("Abrindo CHKDSK em modo leitura para C: numa janela separada...")
        try:
            subprocess.Popen(["cmd", "/k", "chkdsk", "C:"], creationflags=CREATE_NEW_CONSOLE)
            self.log_geral("Janela do CHKDSK aberta. Acompanhe o resultado nela.", "sucesso")
            self.atualizar_status_ferramentas("CHKDSK em execução em janela separada.")
        except Exception as e:
            self.log_geral(f"Erro ao abrir CHKDSK: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")

    def acao_reinstalar_chrome(self):
        self.atualizar_status_ferramentas("Instalando/atualizando o Google Chrome...")
        self.log_geral("Executando winget install Google.Chrome...")
        self._iniciar_progresso_ferramentas()
        try:
            codigo = self._executar_com_saida_em_tempo_real(
                ["winget", "install", "--id", "Google.Chrome", "--silent",
                 "--accept-source-agreements", "--accept-package-agreements"],
                timeout=TIMEOUT_PADRAO_ELEVADO, rotulo="Chrome", encoding="utf-8"
            )
            if codigo != 0:
                self.log_geral("Instalação padrão falhou, tentando reinstalação limpa...", "aviso")
                self._executar_com_saida_em_tempo_real(
                    ["winget", "uninstall", "--id", "Google.Chrome", "--silent"],
                    timeout=TIMEOUT_PADRAO_ELEVADO, rotulo="Remoção do Chrome", encoding="utf-8"
                )
                codigo = self._executar_com_saida_em_tempo_real(
                    ["winget", "install", "--id", "Google.Chrome", "--silent",
                     "--accept-source-agreements", "--accept-package-agreements"],
                    timeout=TIMEOUT_PADRAO_ELEVADO, rotulo="Reinstalação do Chrome", encoding="utf-8"
                )
                if codigo != 0:
                    raise RuntimeError(f"A reinstalação do Chrome terminou com código {codigo}.")
            self.log_geral("Processo do Google Chrome concluído.", "sucesso")
            self.atualizar_status_ferramentas("Chrome instalado/atualizado.")
            self.exibir_info("Concluído", "Processo do Google Chrome concluído.")
        except subprocess.TimeoutExpired:
            self.log_geral(
                f"A instalação do Chrome não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).", "erro"
            )
            self.atualizar_status_ferramentas("Instalação do Chrome expirou (timeout).")
            self.exibir_erro(
                "Tempo esgotado",
                f"A instalação do Chrome não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s). "
                "Tente novamente."
            )
        except Exception as e:
            self.log_geral(f"Erro ao instalar o Chrome: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            self._parar_progresso_ferramentas()

    # --- MÉTODOS OPERACIONAIS DA IMPRESSORA (CORRIGIDO: script .ps1 temporário) ---
    def acao_instalar_impressora_ip(self):
        ip = self.txt_imp_ip.get().strip()
        nome = self.txt_imp_nome.get().strip()
        driver = self.txt_imp_driver.get().strip()

        if not ip or not nome or not driver:
            self.exibir_aviso("Aviso", "Por favor, preencha todos os campos (IP, Nome e Driver).")
            return

        try:
            ip = str(ipaddress.IPv4Address(ip))
        except ipaddress.AddressValueError:
            self.exibir_aviso("Aviso", "Digite um endereço IPv4 válido para a impressora.")
            return

        if len(nome) > 128 or len(driver) > 256 or any(c in nome + driver for c in "\r\n\0"):
            self.exibir_aviso("Aviso", "Nome ou driver da impressora contém caracteres inválidos.")
            return

        confirmar = messagebox.askyesno(
            "Confirmar Instalação",
            f"Deseja instalar a impressora com as seguintes configurações?\n\n"
            f"IP: {ip}\n"
            f"Nome: {nome}\n"
            f"Driver: {driver}\n\n"
            f"Nota: Se o driver não for o correto ou não estiver instalado na máquina, o processo falhará."
        )
        if not confirmar:
            return

        threading.Thread(target=self._processo_instalar_impressora, args=(ip, nome, driver), daemon=True).start()

    def _processo_instalar_impressora(self, ip, nome, driver):
        self.atualizar_status_ferramentas("Criando porta TCP/IP e instalando impressora...")
        self.log_geral(f"Iniciando instalação da impressora '{nome}' no IP {ip}...")
        self._iniciar_progresso_ferramentas()

        caminho_params = os.path.join(tempfile.gettempdir(), f"impressora_params_{uuid.uuid4().hex}.json")

        # Os valores fornecidos pelo usuário ficam em JSON e são lidos como dados.
        # Eles nunca são interpolados no código PowerShell elevado.
        # Add-Printer substitui o rundll32 printui.dll: o printui retornava 0
        # mesmo em caso de falha e podia abrir uma caixa de erro invisível,
        # deixando a operação presa até o timeout.
        corpo_script = f"""
$ErrorActionPreference = 'Stop'
$parametros = Get-Content -Raw -LiteralPath {ps_aspas(caminho_params)} | ConvertFrom-Json
$ip = [string]$parametros.ip
$nomeImpressora = [string]$parametros.nome
$driver = [string]$parametros.driver

try {{
    if (Get-PrinterPort -Name $ip -ErrorAction SilentlyContinue) {{
        Write-Output "Porta TCP/IP '$ip' ja existia e sera reutilizada."
    }} else {{
        Add-PrinterPort -Name $ip -PrinterHostAddress $ip
        Write-Output "Porta TCP/IP '$ip' criada com sucesso."
    }}
    if (-not (Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue)) {{
        Add-PrinterDriver -Name $driver
        Write-Output "Driver '$driver' adicionado a partir do repositorio do Windows."
    }}
    Add-Printer -Name $nomeImpressora -DriverName $driver -PortName $ip
    Write-Output "Impressora '$nomeImpressora' instalada com sucesso."
}} catch {{
    Write-Output "ERRO: $($_.Exception.Message)"
}}
"""
        try:
            with open(caminho_params, "w", encoding="utf-8") as f:
                json.dump({"ip": ip, "nome": nome, "driver": driver}, f, ensure_ascii=False)

            _, saida = self._executar_script_ps_elevado(corpo_script, "Impressora")

            if not saida or "ERRO:" in saida.upper():
                self.log_geral("Falha ao instalar a impressora (veja a mensagem acima).", "erro")
                self.atualizar_status_ferramentas("Falha na instalação.")
                self.exibir_erro(
                    "Erro na Instalação",
                    "O Windows falhou ao mapear a impressora.\n\n"
                    "Causas comuns:\n"
                    "- Nome do driver não é exatamente igual ao instalado no repositório do sistema\n"
                    "- Driver não está instalado/nativo nesta máquina\n"
                    "- Já existe uma impressora com este nome\n\n"
                    f"Mensagem do Windows:\n{saida or '(nenhuma saída capturada)'}"
                )
            else:
                self.log_geral(f"Impressora '{nome}' instalada com sucesso!", "sucesso")
                self.atualizar_status_ferramentas("Impressora instalada.")
                self.exibir_info("Sucesso", f"A impressora '{nome}' foi instalada com sucesso no IP {ip}.")

                self.after(0, lambda: self.txt_imp_ip.delete(0, "end"))
                self.after(0, lambda: self.txt_imp_nome.delete(0, "end"))
                self.after(0, lambda: self.txt_imp_driver.delete(0, "end"))
        except subprocess.TimeoutExpired:
            self.log_geral(
                f"A instalação não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s). "
                "Verifique se não ficou uma janela de confirmação (UAC) pendente na tela.", "erro"
            )
            self.atualizar_status_ferramentas("Instalação expirou (timeout).")
            self.exibir_erro(
                "Tempo esgotado",
                f"A instalação da impressora não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).\n\n"
                "Verifique se não ficou uma janela de confirmação de administrador (UAC) "
                "pendente na tela e tente novamente."
            )
        except Exception as e:
            self.log_geral(f"Erro ao instalar a impressora: {e}", "erro")
            self.atualizar_status_ferramentas("Erro operacional.")
            self.exibir_erro("Erro", f"Falha operacional: {e}")
        finally:
            try:
                os.remove(caminho_params)
            except OSError:
                pass
            self._parar_progresso_ferramentas()

    def acao_reiniciar_explorer(self):
        self.atualizar_status_ferramentas("Reiniciando o Explorer...")
        self.log_geral("Encerrando explorer.exe...")
        try:
            subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], creationflags=CREATE_NO_WINDOW)
            time.sleep(1)
            subprocess.Popen("explorer.exe")
            self.log_geral("Interface reiniciada.", "sucesso")
            self.atualizar_status_ferramentas("Explorer reiniciado.")
        except Exception as e:
            self.log_geral(f"Erro ao reiniciar o Explorer: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")

    def acao_reiniciar_spooler(self):
        self.atualizar_status_ferramentas("Reiniciando o Spooler de Impressão...")
        self.log_geral("Parando e reiniciando o serviço Spooler...")
        self._iniciar_progresso_ferramentas()
        try:
            self._executar_elevado("net stop spooler & timeout /t 2 & net start spooler")
            self.log_geral("Spooler reiniciado.", "sucesso")
            self.atualizar_status_ferramentas("Spooler reiniciado.")
            self.exibir_info("Concluído", "Spooler de impressão reiniciado com sucesso.")
        except Exception as e:
            self.log_geral(f"Erro ao reiniciar o Spooler: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            self._parar_progresso_ferramentas()
