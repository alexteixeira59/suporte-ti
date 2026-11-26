"""Aba Rede e domínio (inclui a senha da conta local de suporte)."""

import ipaddress
import json
import os
import subprocess
import tempfile
import threading
import uuid
from tkinter import messagebox

import customtkinter as ctk

from . import tema
from .comum import TIMEOUT_PADRAO_ELEVADO, ps_aspas, restringir_arquivo_temporario


class RedeMixin:
    # ==================================================================
    # ABA: REDE E INFRA
    # ==================================================================

    def _montar_aba_rede(self, aba):
        pagina = tema.pagina(aba)
        lista = tema.lista(pagina)
        self._botao_ferramenta(
            lista, "Reparar rede (IP, DNS, Winsock e TCP/IP)", self.acao_reparar_rede,
            confirmar="Renovar o IP, limpar o cache DNS e reiniciar Winsock e TCP/IP? A rede cai por alguns segundos."
        )
        self._botao_ferramenta(lista, "Atualizar políticas de grupo (gpupdate /force)", self.acao_gpupdate)

        tema.titulo_grupo(pagina, "DNS")
        self.txt_dns_pref = tema.campo(pagina, "Servidor preferencial", "8.8.8.8", largura=260)
        self.txt_dns_alt = tema.campo(pagina, "Servidor alternativo (opcional)", "8.8.4.4", largura=260)
        ctk.CTkFrame(pagina, height=8, fg_color="transparent").pack()
        self._botao_ferramenta(pagina, "Aplicar DNS", self.acao_alterar_dns, em_segundo_plano=False)

        tema.titulo_grupo(pagina, "Ingressar no domínio")
        tema.apoio(pagina, "Antes, confira se o DNS desta máquina aponta para o controlador de domínio.")
        self.txt_dominio_nome = tema.campo(pagina, "Domínio", "empresa.local", largura=360)
        self.txt_dominio_novo_nome = tema.campo(pagina, "Novo nome da máquina (opcional)", "PC-RECEPCAO01", largura=360)
        self.txt_dominio_usuario = tema.campo(pagina, "Usuário com permissão", "empresa\\admin ou admin@empresa.local")
        self.txt_dominio_senha = tema.campo(pagina, "Senha desse usuário", largura=360, show="*")
        self.reiniciar_apos_dominio = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            pagina, text="Reiniciar ao terminar", variable=self.reiniciar_apos_dominio, font=tema.TEXTO
        ).pack(anchor="w", padx=2, pady=(10, 4))
        self._botao_ferramenta(pagina, "Ingressar no domínio", self.acao_ingressar_dominio, em_segundo_plano=False)

        tema.titulo_grupo(pagina, "Conta local de suporte")
        tema.apoio(pagina, "Troca a senha da conta local \"suporte\". A senha não vai para o registro.")
        self.txt_nova_senha = tema.campo(pagina, "Nova senha", largura=360, show="*")
        ctk.CTkFrame(pagina, height=8, fg_color="transparent").pack()
        self._botao_ferramenta(pagina, "Trocar senha", self.acao_alterar_senha_suporte, em_segundo_plano=False)

    def acao_reparar_rede(self):
        self.atualizar_status_ferramentas("Reparando rede...")
        self.log_geral("Renovando IP e limpando cache DNS...")
        self._iniciar_progresso_ferramentas()
        try:
            etapas = [
                (["ipconfig", "/release"], "Liberando endereço IP"),
                (["ipconfig", "/renew"], "Renovando endereço IP"),
                (["ipconfig", "/flushdns"], "Limpando cache DNS"),
            ]
            for comando, rotulo in etapas:
                self.log_geral(f"{rotulo}...")
                codigo = self._executar_com_saida_em_tempo_real(comando, timeout=60, rotulo=rotulo)
                if codigo != 0:
                    raise RuntimeError(f"{rotulo} terminou com código {codigo}.")
            self.log_geral("Resetando pilha TCP/IP e Winsock (requer administrador)...")
            self._executar_elevado("netsh int ip reset")
            self._executar_elevado("netsh winsock reset")
            self.log_geral("Reparo de rede concluído. Pode ser necessário reiniciar o computador.", "sucesso")
            self.atualizar_status_ferramentas("Reparo de rede concluído.")
            self.exibir_info("Concluído", "Reparo de rede executado. Recomenda-se reiniciar o computador.")
        except subprocess.TimeoutExpired:
            self.log_geral("Uma etapa do reparo de rede (ipconfig) não respondeu a tempo.", "erro")
            self.atualizar_status_ferramentas("Reparo de rede expirou (timeout).")
            self.exibir_erro("Tempo esgotado", "Uma das etapas de rede não respondeu a tempo. Tente novamente.")
        except Exception as e:
            self.log_geral(f"Erro ao reparar a rede: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            self._parar_progresso_ferramentas()

    def acao_gpupdate(self):
        self.atualizar_status_ferramentas("Atualizando políticas de grupo...")
        self.log_geral("Executando gpupdate /force...")
        self._iniciar_progresso_ferramentas()
        try:
            codigo = self._executar_com_saida_em_tempo_real(
                ["gpupdate", "/force"], timeout=TIMEOUT_PADRAO_ELEVADO, rotulo="GPUPDATE"
            )
            if codigo != 0:
                raise RuntimeError(f"O GPUPDATE terminou com código {codigo}.")
            self.log_geral("Políticas de grupo atualizadas.", "sucesso")
            self.atualizar_status_ferramentas("GPUPDATE concluído.")
            self.exibir_info("Concluído", "Políticas de grupo atualizadas com sucesso.")
        except subprocess.TimeoutExpired:
            self.log_geral(f"O GPUPDATE não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).", "erro")
            self.atualizar_status_ferramentas("GPUPDATE expirou (timeout).")
            self.exibir_erro("Tempo esgotado", "O GPUPDATE não respondeu a tempo. Tente novamente.")
        except Exception as e:
            self.log_geral(f"Erro no GPUPDATE: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            self._parar_progresso_ferramentas()

    def acao_alterar_dns(self):
        dns_pref = self.txt_dns_pref.get().strip()
        dns_alt = self.txt_dns_alt.get().strip()
        if not dns_pref:
            self.exibir_aviso("Aviso", "Digite ao menos o DNS Preferencial.")
            return
        try:
            dns_pref = str(ipaddress.IPv4Address(dns_pref))
            if dns_alt:
                dns_alt = str(ipaddress.IPv4Address(dns_alt))
        except ipaddress.AddressValueError:
            self.exibir_aviso("Aviso", "Digite somente endereços DNS IPv4 válidos.")
            return
        if not messagebox.askyesno("Confirmar", f"Aplicar DNS Preferencial={dns_pref} / Alternativo={dns_alt or '(nenhum)'}?"):
            return
        threading.Thread(target=self._processo_alterar_dns, args=(dns_pref, dns_alt), daemon=True).start()

    def _processo_alterar_dns(self, dns_pref, dns_alt):
        self.atualizar_status_ferramentas("Aplicando novo DNS...")
        self.log_geral(f"Aplicando DNS: preferencial={dns_pref}, alternativo={dns_alt or '(nenhum)'}...")
        self._iniciar_progresso_ferramentas()
        try:
            # Os endereços já foram validados com ipaddress.IPv4Address.
            enderecos = ps_aspas(dns_pref) if not dns_alt else f"{ps_aspas(dns_pref)},{ps_aspas(dns_alt)}"
            corpo_script = f"""
try {{
    $rota = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
        Sort-Object -Property RouteMetric | Select-Object -First 1
    $indiceInterface = $rota.InterfaceIndex

    Set-DnsClientServerAddress -InterfaceIndex $indiceInterface -ServerAddresses {enderecos} -ErrorAction Stop

    $confirmacao = Get-DnsClientServerAddress -InterfaceIndex $indiceInterface -AddressFamily IPv4
    Write-Output "DNS aplicado na interface $indiceInterface ($($rota.InterfaceAlias))"
    Write-Output "DNS atual: $($confirmacao.ServerAddresses -join ', ')"
}} catch {{
    Write-Output "ERRO: $($_.Exception.Message)"
}}
"""
            _, saida = self._executar_script_ps_elevado(corpo_script, "DNS")

            if not saida or "ERRO:" in saida.upper():
                self.log_geral("Falha ao aplicar o novo DNS (veja a mensagem acima).", "erro")
                self.atualizar_status_ferramentas("Falha ao atualizar o DNS.")
                self.exibir_erro(
                    "Erro",
                    "Não foi possível confirmar a alteração de DNS. "
                    "O Windows recusou a operação.\n\n"
                    f"{saida or '(nenhuma saída capturada)'}"
                )
            else:
                self.log_geral("DNS atualizado com sucesso.", "sucesso")
                self.atualizar_status_ferramentas("DNS atualizado.")
                self.exibir_info("Concluído", f"Novo DNS aplicado com sucesso.\n\n{saida}")
        except subprocess.TimeoutExpired:
            self.log_geral(
                f"A alteração de DNS não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s). "
                "Verifique se não ficou uma janela de confirmação (UAC) pendente na tela.", "erro"
            )
            self.atualizar_status_ferramentas("Alteração de DNS expirou (timeout).")
            self.exibir_erro(
                "Tempo esgotado",
                f"A alteração de DNS não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).\n\n"
                "Verifique se não ficou uma janela de confirmação de administrador (UAC) "
                "pendente na tela e tente novamente."
            )
        except Exception as e:
            self.log_geral(f"Erro ao alterar o DNS: {e}", "erro")
            self.exibir_erro("Erro", f"Erro ao alterar o DNS: {e}")
        finally:
            self._parar_progresso_ferramentas()

    def acao_ingressar_dominio(self):
        dominio = self.txt_dominio_nome.get().strip()
        novo_nome = self.txt_dominio_novo_nome.get().strip()
        usuario = self.txt_dominio_usuario.get().strip()
        senha = self.txt_dominio_senha.get()
        reiniciar = self.reiniciar_apos_dominio.get()

        if not dominio:
            self.exibir_aviso("Aviso", "Digite o nome do domínio (ex: empresa.local).")
            return
        if not usuario:
            self.exibir_aviso("Aviso", "Digite o usuário com permissão de ingressar máquinas no domínio.")
            return
        if not senha:
            self.exibir_aviso("Aviso", "Digite a senha do usuário de domínio.")
            return

        if novo_nome:
            caracteres_invalidos = set(novo_nome) - set(
                "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"
            )
            if caracteres_invalidos or len(novo_nome) > 15:
                self.exibir_aviso(
                    "Aviso",
                    "O nome da máquina deve ter até 15 caracteres, sem espaços, acentos ou "
                    "caracteres especiais (apenas letras, números e hífen)."
                )
                return

        confirmar = messagebox.askyesno(
            "Confirmar Ingresso no Domínio",
            f"Isso vai ingressar esta máquina no domínio '{dominio}'"
            + (f", renomeando-a para '{novo_nome}'" if novo_nome else "")
            + f", usando o usuário '{usuario}'.\n\n"
            + ("A máquina será reiniciada automaticamente ao concluir.\n\n" if reiniciar else "")
            + "Continuar?"
        )
        if not confirmar:
            return

        threading.Thread(
            target=self._processo_ingressar_dominio, args=(dominio, novo_nome, usuario, senha, reiniciar), daemon=True
        ).start()
        self.after(0, lambda: self.txt_dominio_senha.delete(0, "end"))

    def _processo_ingressar_dominio(self, dominio, novo_nome, usuario, senha, reiniciar):
        self.atualizar_status_ferramentas(f"Ingressando no domínio {dominio}...")
        self.log_geral(
            f"Iniciando ingresso no domínio '{dominio}'..."
            + (f" (renomeando a máquina para '{novo_nome}')" if novo_nome else "")
        )
        self._iniciar_progresso_ferramentas()

        caminho_params = os.path.join(tempfile.gettempdir(), f"dominio_params_{uuid.uuid4().hex}.json")

        # IMPORTANTE: caminho_params contém a senha em texto puro temporariamente.
        # A senha não vai na linha de comando (ficaria visível na lista de
        # processos e nos logs de auditoria). O arquivo tem ACL restrita, é
        # apagado pelo script elevado logo após a leitura e de novo no `finally`.
        try:
            with open(caminho_params, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "dominio": dominio,
                        "novo_nome": novo_nome,
                        "usuario": usuario,
                        "senha": senha,
                        "reiniciar": reiniciar,
                    },
                    f
                )
            restringir_arquivo_temporario(caminho_params)

            corpo_script = f"""
$ErrorActionPreference = 'Stop'
try {{
    $parametrosJson = Get-Content -Raw -LiteralPath {ps_aspas(caminho_params)} | ConvertFrom-Json
    Remove-Item -LiteralPath {ps_aspas(caminho_params)} -Force -ErrorAction SilentlyContinue
    $senhaSegura = ConvertTo-SecureString ([string]$parametrosJson.senha) -AsPlainText -Force
    $credencial = New-Object System.Management.Automation.PSCredential([string]$parametrosJson.usuario, $senhaSegura)

    $parametrosJoin = @{{
        DomainName = [string]$parametrosJson.dominio
        Credential = $credencial
        Force = $true
    }}
    if ($parametrosJson.novo_nome) {{ $parametrosJoin['NewName'] = [string]$parametrosJson.novo_nome }}

    Add-Computer @parametrosJoin

    if ($parametrosJson.novo_nome) {{
        Write-Output "Maquina renomeada para $($parametrosJson.novo_nome) e ingressada no dominio $($parametrosJson.dominio) com sucesso."
    }} else {{
        Write-Output "Maquina ingressada no dominio $($parametrosJson.dominio) com sucesso."
    }}
    if ($parametrosJson.reiniciar) {{
        # Atraso de 15 s para a saída ser gravada e lida pelo aplicativo antes do reinício.
        shutdown.exe /r /t 15 /c "Reiniciando para concluir o ingresso no dominio."
    }}
}} catch {{
    Write-Output "ERRO: $($_.Exception.Message)"
}} finally {{
    Remove-Item -LiteralPath {ps_aspas(caminho_params)} -Force -ErrorAction SilentlyContinue
}}
"""
            try:
                _, saida = self._executar_script_ps_elevado(corpo_script, "Domínio")
            except subprocess.TimeoutExpired:
                self.log_geral(
                    f"O ingresso no domínio não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s). "
                    "Verifique se não ficou uma janela de confirmação (UAC) pendente na tela.", "erro"
                )
                self.atualizar_status_ferramentas("Ingresso no domínio expirou (timeout).")
                self.exibir_erro(
                    "Tempo esgotado",
                    f"O ingresso no domínio não respondeu em {TIMEOUT_PADRAO_ELEVADO // 60} minuto(s).\n\n"
                    "Verifique se não ficou uma janela de confirmação de administrador (UAC) "
                    "pendente na tela e tente novamente."
                )
                return

            if not saida or "ERRO:" in saida.upper():
                self.log_geral("Falha ao ingressar no domínio (veja a mensagem acima).", "erro")
                self.atualizar_status_ferramentas("Falha ao ingressar no domínio.")
                self.exibir_erro(
                    "Erro",
                    "Não foi possível ingressar no domínio. Causas comuns:\n\n"
                    "- O DNS da máquina não aponta para o Domain Controller\n"
                    "- Sem conectividade de rede até o DC (porta 445/389/88 bloqueada)\n"
                    "- Usuário ou senha incorretos, ou sem permissão de ingressar máquinas\n"
                    "- Já existe um objeto de computador com este nome no domínio\n"
                    "- Nome do domínio digitado errado (precisa ser o FQDN, ex: empresa.local)\n"
                    "- Nome de máquina inválido ou já em uso no domínio\n\n"
                    f"Mensagem do Windows:\n{saida or '(nenhuma saída capturada)'}"
                )
            else:
                self.log_geral("Máquina ingressada no domínio com sucesso.", "sucesso")
                self.atualizar_status_ferramentas("Ingressado no domínio.")
                if reiniciar:
                    self.exibir_info("Concluído", f"{saida}\n\nA máquina será reiniciada em 15 segundos.")
                else:
                    self.exibir_info("Concluído", f"{saida}\n\nReinicie a máquina para concluir o ingresso no domínio.")
        except Exception as e:
            self.log_geral(f"Erro ao ingressar no domínio: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            senha = None
            self._parar_progresso_ferramentas()
            # Remoção garantida mesmo em caso de timeout: o arquivo contém a senha.
            try:
                os.remove(caminho_params)
            except OSError:
                pass

    def acao_alterar_senha_suporte(self):
        # Sem .strip(): espaços no início/fim fazem parte da senha digitada.
        nova_senha = self.txt_nova_senha.get()
        if not nova_senha.strip():
            self.exibir_aviso("Aviso", "Digite a nova senha antes de aplicar.")
            return
        if not messagebox.askyesno("Confirmar", "Isso vai alterar a senha da conta local 'suporte'. Continuar?"):
            return
        threading.Thread(target=self._processo_alterar_senha, args=(nova_senha,), daemon=True).start()

    def _processo_alterar_senha(self, nova_senha):
        self.atualizar_status_ferramentas("Alterando senha da conta suporte...")
        self.log_geral("Aplicando nova senha na conta 'suporte'...")
        self._iniciar_progresso_ferramentas()
        caminho_params = os.path.join(tempfile.gettempdir(), f"senha_params_{uuid.uuid4().hex}.json")
        try:
            # A senha é transportada como dado, não como parte de uma linha de
            # comando. Isso evita injeção de cmd e exposição na lista de processos.
            with open(caminho_params, "w", encoding="utf-8") as f:
                json.dump({"senha": nova_senha}, f)
            restringir_arquivo_temporario(caminho_params)

            corpo_script = f"""
$ErrorActionPreference = 'Stop'
try {{
    $parametros = Get-Content -Raw -LiteralPath {ps_aspas(caminho_params)} | ConvertFrom-Json
    Remove-Item -LiteralPath {ps_aspas(caminho_params)} -Force -ErrorAction SilentlyContinue
    $senhaSegura = ConvertTo-SecureString ([string]$parametros.senha) -AsPlainText -Force
    Set-LocalUser -Name 'suporte' -Password $senhaSegura -ErrorAction Stop
    Write-Output 'Senha alterada com sucesso.'
}} catch {{
    Write-Output "ERRO: $($_.Exception.Message)"
}} finally {{
    Remove-Item -LiteralPath {ps_aspas(caminho_params)} -Force -ErrorAction SilentlyContinue
}}
"""
            _, saida = self._executar_script_ps_elevado(corpo_script, "Senha da conta suporte")
            sucesso = bool(saida) and "ERRO:" not in saida.upper()
            if sucesso:
                self.log_geral("Senha da conta suporte alterada.", "sucesso")
                self.atualizar_status_ferramentas("Senha alterada.")
                self.exibir_info("Concluído", "Senha da conta suporte alterada com sucesso.")
                self.after(0, lambda: self.txt_nova_senha.delete(0, "end"))
            else:
                self.log_geral("O Windows recusou a nova senha (veja a mensagem acima).", "erro")
                self.atualizar_status_ferramentas("Falha ao alterar a senha.")
                self.exibir_erro(
                    "Erro",
                    "O Windows recusou a nova senha. Causas comuns:\n\n"
                    "- Não atende à política de complexidade (maiúscula, número, símbolo, "
                    "tamanho mínimo)\n"
                    "- É uma senha já usada recentemente (histórico de senhas)\n"
                    "- A conta 'suporte' não existe localmente nesta máquina\n\n"
                    "Veja o log para o texto exato retornado pelo Windows."
                )
        except subprocess.TimeoutExpired:
            self.log_geral("A alteração de senha não respondeu a tempo (UAC pendente?).", "erro")
            self.exibir_erro("Tempo esgotado", "A alteração de senha não respondeu a tempo. Verifique se há um UAC pendente.")
        except Exception as e:
            self.log_geral(f"Erro ao alterar a senha: {e}", "erro")
            self.exibir_erro("Erro", f"Erro: {e}")
        finally:
            nova_senha = None
            try:
                os.remove(caminho_params)
            except OSError:
                pass
            self._parar_progresso_ferramentas()
