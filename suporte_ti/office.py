"""Aba Office: reinstalação automatizada do Office 2024."""

import os
import subprocess
import threading
from tkinter import messagebox

from . import tema
from .comum import ps_aspas, resource_path

# Localiza o OSPP.VBS para limpar chaves antigas antes da reinstalação. A pasta
# Program Files vem do registro (HKLM) e não de variáveis de ambiente, que o
# usuário comum pode alterar para apontar para um OSPP.VBS falso.
SCRIPT_PS_LOCALIZAR_OSPP = r"""
$chaveReg = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
$bases = @($chaveReg.ProgramFilesDir, $chaveReg.'ProgramFilesDir (x86)', $chaveReg.ProgramW6432Dir) |
    Where-Object { $_ } | Select-Object -Unique
$subpastas = @('Microsoft Office\Office16', 'Microsoft Office\root\Office16',
               'Microsoft Office\Office15', 'Microsoft Office\root\Office15')
$ospp = $null
foreach ($base in $bases) {
    foreach ($sub in $subpastas) {
        $candidato = Join-Path (Join-Path $base $sub) 'OSPP.VBS'
        if (Test-Path -LiteralPath $candidato) { $ospp = $candidato; break }
    }
    if ($ospp) { break }
}
if (-not $ospp) { throw 'OSPP.VBS nao encontrado.' }
$cscript = Join-Path ([Environment]::SystemDirectory) 'cscript.exe'
"""

# Pasta OFFICE (setup.exe + Configuracao.xml do Office 2024), empacotada no .exe pelo build.ps1.
PASTA_INSTALADOR_OFFICE = resource_path("OFFICE")
MARCADOR_OFFICE_INSTALADO = "=== OFFICE 2024 INSTALADO ==="
TIMEOUT_INSTALACAO_OFFICE = 2 * 60 * 60  # download + instalação pela internet

# Remove as chaves antigas, desinstala o Office Click-to-Run atual (o <RemoveMSI />
# da Configuracao.xml cuida das versões MSI) e instala o 2024. Os arquivos são
# copiados para Program Files (só administradores alteram) e a assinatura do
# setup.exe é conferida antes de rodá-lo como administrador.
SCRIPT_PS_REINSTALAR_OFFICE = r"""
$ErrorActionPreference = 'Stop'
$chaveReg = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
$destino = Join-Path $chaveReg.ProgramFilesDir 'SuporteTI\Office2024'
try {
    if (Test-Path -LiteralPath $destino) { Remove-Item -LiteralPath $destino -Recurse -Force }
    New-Item -ItemType Directory -Path $destino -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $origem 'setup.exe'), (Join-Path $origem 'Configuracao.xml') -Destination $destino
    $setup = Join-Path $destino 'setup.exe'
    $assinatura = Get-AuthenticodeSignature -LiteralPath $setup
    if ($assinatura.Status -ne 'Valid' -or $assinatura.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        throw 'O setup.exe nao tem assinatura digital valida da Microsoft.'
    }

    $ospp = $null
    try {
""" + SCRIPT_PS_LOCALIZAR_OSPP + r"""
    } catch { $ospp = $null }
    if ($ospp) {
        Write-Output 'Removendo chaves antigas do Office...'
        $status = & $cscript //nologo $ospp /dstatus | Out-String
        foreach ($m in [regex]::Matches($status, 'Last 5 characters of installed product key:\s*([A-Z0-9]{5})')) {
            & $cscript //nologo $ospp "/unpkey:$($m.Groups[1].Value)"
        }
    }

    Write-Output 'Importando configuracao de proxy (netsh winhttp)...'
    netsh winhttp import proxy source=ie

    Write-Output 'Desinstalando o Office atual (pode levar alguns minutos)...'
    $xmlRemover = Join-Path $destino 'Remover.xml'
    Set-Content -LiteralPath $xmlRemover -Encoding UTF8 -Value '<Configuration><Remove All="TRUE" /><Display Level="None" AcceptEULA="TRUE" /><Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /></Configuration>'
    $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$xmlRemover`"" -Wait -PassThru
    Write-Output "Desinstalacao terminou (codigo $($p.ExitCode))."

    Write-Output 'Instalando o Office 2024 (download pela internet, pode levar bastante tempo)...'
    $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$(Join-Path $destino 'Configuracao.xml')`"" -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw "A instalacao do Office terminou com codigo $($p.ExitCode)." }
    Write-Output '""" + MARCADOR_OFFICE_INSTALADO + r"""'
} catch {
    Write-Output "ERRO: $($_.Exception.Message)"
} finally {
    Remove-Item -LiteralPath $destino -Recurse -Force -ErrorAction SilentlyContinue
}
"""


class OfficeMixin:
    # ==================================================================
    # ABA: OFFICE (reinstalação do Office 2024)
    # ==================================================================

    def _montar_aba_office(self, aba):
        frame = tema.pagina(aba)
        tema.titulo_grupo(frame, "Reinstalar o Office 2024", primeiro=True)
        tema.apoio(
            frame,
            "Desinstala o Office atual e instala o Office LTSC 2024 de 64 bits com a "
            "Configuracao.xml da pasta OFFICE. Precisa de internet e de administrador; leva de 15 a 60 minutos."
        )
        self.btn_reinstalar_office = tema.botao_secundario(
            frame, "Desinstalar e instalar o 2024", self.acao_reinstalar_office, width=240
        )
        self.btn_reinstalar_office.pack(anchor="w", padx=2)

    def _definir_botoes_office(self, ocupado):
        self.btn_reinstalar_office.configure(state="disabled" if ocupado else "normal")

    def acao_reinstalar_office(self):
        faltando = [n for n in ("setup.exe", "Configuracao.xml")
                    if not os.path.isfile(os.path.join(PASTA_INSTALADOR_OFFICE, n))]
        if faltando:
            messagebox.showerror(
                "Instalador não encontrado",
                f"Arquivos ausentes em {PASTA_INSTALADOR_OFFICE}: {', '.join(faltando)}.\n"
                "Recompile com o build.ps1 com a pasta OFFICE ao lado do script."
            )
            return
        if not messagebox.askyesno(
            "Reinstalar Office 2024",
            "Isso vai DESINSTALAR todo o Office desta máquina e depois instalar o Office 2024.\n\n"
            "• Salve o trabalho e feche Word, Excel, Outlook etc. (serão fechados à força).\n"
            "• É necessário internet; pode levar de 15 a 60 minutos.\n\nContinuar?",
            icon="warning",
        ):
            return
        self._definir_botoes_office(True)
        threading.Thread(target=self._processo_reinstalar_office, daemon=True).start()

    def _processo_reinstalar_office(self):
        self._iniciar_progresso_ferramentas()
        self.atualizar_status_ferramentas("Reinstalando o Office 2024...")
        self.log_geral("Iniciando desinstalação do Office atual e instalação do Office 2024...")
        try:
            corpo_script = f"$origem = {ps_aspas(PASTA_INSTALADOR_OFFICE)}\n" + SCRIPT_PS_REINSTALAR_OFFICE
            _, saida = self._executar_script_ps_elevado(corpo_script, "Office 2024", timeout=TIMEOUT_INSTALACAO_OFFICE)
            if MARCADOR_OFFICE_INSTALADO in saida:
                self.log_geral("Office 2024 instalado com sucesso.", "sucesso")
                self.atualizar_status_ferramentas("Office 2024 instalado.")
                self.exibir_info("Concluído", "Office 2024 instalado.")
            else:
                erro = next((l for l in saida.splitlines() if l.upper().startswith("ERRO:")), "sem detalhes")
                self.log_geral(f"Falha ao reinstalar o Office: {erro}", "erro")
                self.atualizar_status_ferramentas("Falha ao reinstalar o Office.")
                self.exibir_erro("Falha", f"A reinstalação do Office não foi concluída.\n{erro}")
        except PermissionError as e:
            self.log_geral(str(e), "erro")
            self.atualizar_status_ferramentas("Operação cancelada no UAC.")
        except subprocess.TimeoutExpired:
            self.log_geral("A instalação do Office passou do tempo limite.", "erro")
            self.exibir_erro("Tempo esgotado", "A instalação do Office não terminou dentro de 2 horas.")
        except Exception as e:
            self.log_geral(f"Erro ao reinstalar o Office: {e}", "erro")
            self.exibir_erro("Erro", f"Erro ao reinstalar o Office:\n{e}")
        finally:
            self._parar_progresso_ferramentas()
            self.after(0, lambda: self._definir_botoes_office(False))
