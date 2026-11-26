"""Utilitários compartilhados: caminhos, logs em disco, aspas do PowerShell."""

import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

def resource_path(nome_arquivo):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # Raiz do projeto (pasta acima de suporte_ti/), onde ficam OFFICE/ e logo.ico.
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, nome_arquivo)


# Flag usada pelo subprocess para não abrir janelas pretas de console
# quando não precisamos que o usuário veja o console rodando.
CREATE_NO_WINDOW = 0x08000000
CREATE_NEW_CONSOLE = 0x00000010

# Tempo máximo (segundos) que aguardamos por um comando elevado antes de
# desistir. Evita que a thread fique presa para sempre se houver um prompt
# de UAC esquecido em segundo plano ou um comando travado.
TIMEOUT_PADRAO_ELEVADO = 600  # 10 minutos


def restringir_arquivo_temporario(caminho):
    """Restringe um arquivo sensível ao usuário atual, SYSTEM e administradores.

    A operação falha de forma fechada: se a ACL não puder ser aplicada, o
    arquivo é removido e a ação que utilizaria a credencial deve ser abortada.
    """
    dominio = os.environ.get("USERDOMAIN", "").strip()
    usuario = os.environ.get("USERNAME", "").strip()
    identidade = f"{dominio}\\{usuario}" if dominio and usuario else usuario
    if not identidade:
        try:
            os.remove(caminho)
        except OSError:
            pass
        raise PermissionError("Não foi possível identificar o usuário atual para proteger a credencial.")

    resultado = subprocess.run(
        [
            "icacls", caminho,
            "/inheritance:r",
            "/grant:r", f"{identidade}:(F)",
            "*S-1-5-18:(F)",
            "*S-1-5-32-544:(F)",
        ],
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
        check=False,
    )
    if resultado.returncode != 0:
        try:
            os.remove(caminho)
        except OSError:
            pass
        raise PermissionError("O Windows não permitiu restringir o arquivo temporário de credencial.")


def caminho_pasta_logs():
    """Pasta onde os logs de sessão das ferramentas são gravados em disco,
    para auditoria (o que foi feito em cada máquina e quando)."""
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    pasta = os.path.join(base, "SuporteTI", "logs")
    try:
        os.makedirs(pasta, exist_ok=True)
    except Exception:
        pasta = tempfile.gettempdir()
    return pasta


def caminho_log_ferramentas_do_dia():
    nome_arquivo = f"log_ferramentas_{datetime.now().strftime('%Y%m%d')}.txt"
    return os.path.join(caminho_pasta_logs(), nome_arquivo)


def limpar_conteudo_diretorio(caminho):
    """Remove o conteúdo de uma pasta sem delegar a entrada ao shell.

    Arquivos em uso são ignorados, como é esperado numa limpeza de TEMP.
    A própria pasta informada nunca é removida.
    """
    caminho_real = os.path.realpath(caminho)
    unidade, restante = os.path.splitdrive(caminho_real)
    if not caminho_real or not unidade or restante in ("", os.sep):
        raise ValueError("Caminho de limpeza inválido ou excessivamente amplo.")
    if not os.path.isdir(caminho_real):
        return

    with os.scandir(caminho_real) as itens:
        for item in itens:
            try:
                if item.is_dir(follow_symlinks=False) and not item.is_symlink():
                    shutil.rmtree(item.path)
                else:
                    os.remove(item.path)
            except OSError:
                # Temporários em uso são comuns e não devem abortar a limpeza.
                continue


# Código devolvido quando o usuário cancela/nega o UAC (ERROR_CANCELLED).
CODIGO_UAC_CANCELADO = 1223


def ps_aspas(texto):
    """Coloca um texto entre aspas simples do PowerShell, escapando apóstrofos.

    Sem isso, um caminho como C:\\Users\\D'Avila\\AppData\\Local\\Temp quebra o
    comando (ou permite injetar código nele).
    """
    return "'" + str(texto).replace("'", "''") + "'"
