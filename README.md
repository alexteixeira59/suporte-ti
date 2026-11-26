# Suporte TI — Automação de suporte técnico para Windows

Aplicativo desktop (Python + CustomTkinter) que reúne em uma única janela as tarefas repetitivas do dia a dia de um técnico de suporte. Cada ferramenta nasceu de um atendimento real e foi pensada para **trocar um procedimento manual de vários passos por um clique**, com confirmação, registro e validação de segurança.

> Desenvolvido por **Alex Teixeira**, com apoio de IA no desenvolvimento.

---

## ⭐ Destaque: arquivamento de e-mails do Outlook em PST

O módulo que mais economizou tempo. Ele automatiza a criação do arquivo PST e a movimentação dos e-mails antigos de uma caixa do Outlook.

**Resultado em produção: cerca de 300 mil e-mails movidos para PST em aproximadamente 1 hora.**

Como funciona:

1. **Detecta as contas** configuradas no Outlook (via COM/MAPI com `pywin32`).
2. **Filtra por ano** com uma consulta DASL (`Items.Restrict`) direto no Outlook, em vez de percorrer a caixa inteira. Há um filtro alternativo (Jet) se o DASL falhar.
3. **Conta antes de mover**: o botão "Contar e-mails" mostra quantas mensagens seriam arquivadas, sem alterar nada.
4. **Verifica antes de criar o PST**:
   - espaço livre no destino (1,5× o tamanho das mensagens);
   - limite seguro de 45 GB por PST (o Outlook trava em 50 GB);
   - bloqueia destino em rede ou OneDrive, que corrompem PST.
5. **Cria ou reaproveita o PST** (`Arquivados_<ano>_<conta>.pst`) com `AddStore` e move as mensagens com progresso em tempo real e cancelamento a qualquer momento.
6. **Confere o resultado**: compara quantos e-mails saíram da caixa com quantos chegaram no PST e alerta se os números não batem.
7. Desanexa o PST do Outlook ao terminar (opcional) e salva o log da operação.

Também tem a **limpeza de caixa**: encontra newsletters, avisos automáticos e mensagens grandes e envia as selecionadas para a Lixeira.

---

## Outras ferramentas

| Área | O que automatiza |
|------|------------------|
| **Visão geral** | Ficha da máquina (Windows, IP, disco, internet, Winget, Outlook, permissão) e as correções mais usadas a um clique. |
| **Manutenção** | `winget upgrade --all`, SFC, CHKDSK, limpeza de TEMP/Prefetch/cache do Windows Update/Lixeira, reinício do Explorer, reinstalação do Chrome. |
| **Rede e domínio** | Reparo de rede (IP, DNS, Winsock, TCP/IP), troca de DNS, `gpupdate`, ingresso no domínio com renomeação e troca de senha da conta local de suporte. |
| **Impressoras** | Reinício do spooler e instalação de impressora por IP. |
| **Office** | Desinstala o Office atual e instala o Office 2024 pela Office Deployment Tool, conferindo a assinatura digital da Microsoft no `setup.exe`. |

Tudo o que é executado aparece na coluna **Registro**, que pode ser copiada para o chamado.

## Segurança

- Pede administrador (UAC) **só nas ações que precisam**.
- Valida toda entrada que vai para um script elevado (IP, nome, ano etc.).
- Senhas e parâmetros sensíveis passam por arquivo temporário com ACL restrita, nunca pela linha de comando.
- Ações destrutivas sempre pedem confirmação.

## Tecnologias

Python 3.12+ · CustomTkinter · pywin32 (COM do Outlook) · PowerShell · PyInstaller

## Como rodar

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python ferramentas_suporte_ti.py
```

Para gerar o `.exe`, veja [COMO_COMPILAR.md](COMO_COMPILAR.md).

O instalador do Office não está no repositório. Baixe na fonte oficial:
- `OFFICE/setup.exe`: [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117) (a `Configuracao.xml` já está na pasta).

## Requisitos

Windows 10/11 e Microsoft Outlook (desktop) instalado para as abas de e-mail.
