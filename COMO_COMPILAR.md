# Como compilar a Ferramenta de Suporte de TI

Resultado final: um único arquivo **`dist\SuporteTI.exe`**, que roda em qualquer
Windows 10/11 sem precisar instalar Python.

---

## 1. Pré-requisitos (uma vez só)

1. Instale o **Python 3.12 ou mais novo** em <https://www.python.org/downloads/windows/>.
   - Na primeira tela do instalador, marque **"Add python.exe to PATH"**.
2. Abra um PowerShell **comum** (não precisa ser administrador) e confira:
   ```powershell
   python --version
   ```
3. (Opcional) Coloque um ícone chamado **`logo.ico`** na pasta `D:\Projeto`.
   Sem ele o app funciona normalmente, só que com o ícone padrão.

## 2. Compilar (jeito rápido: script pronto)

```powershell
cd D:\Projeto
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

O script faz tudo sozinho: cria o ambiente virtual `.venv`, instala as
dependências do `requirements.txt` e roda o PyInstaller. No final ele mostra o
caminho do `.exe` e o **hash SHA-256** (guarde-o para conferir a integridade do
arquivo que for distribuído).

## 3. Compilar (passo a passo manual, se preferir)

```powershell
cd D:\Projeto

# 3.1 Ambiente virtual isolado (não mistura com outros projetos)
python -m venv .venv
.\.venv\Scripts\Activate.ps1
#   Se der erro de "execução de scripts desabilitada":
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 3.2 Dependências (versões fixas e testadas)
pip install -r requirements.txt

# 3.3 Teste antes de compilar
python ferramentas_suporte_ti.py

# 3.4 Compilação
pyinstaller --noconfirm --clean --onefile --windowed --name SuporteTI `
    --collect-data customtkinter `
    --icon logo.ico --add-data "logo.ico;." `
    --add-data "OFFICE\setup.exe;OFFICE" --add-data "OFFICE\Configuracao.xml;OFFICE" `
    ferramentas_suporte_ti.py
#   (sem logo.ico, remova a linha "--icon ... --add-data ...")
#   O ferramentas_suporte_ti.py só inicia o app; o código fica no pacote suporte_ti\,
#   que o PyInstaller inclui sozinho por causa do import.
```

| Opção | Para que serve |
|---|---|
| `--onefile` | Gera um único `.exe` (bom para pendrive). |
| `--windowed` | Não abre a janela preta do console junto com o app. |
| `--collect-data customtkinter` | Inclui os temas/cores do CustomTkinter. Sem isso o `.exe` fecha ao abrir. |
| `--icon` / `--add-data logo.ico` | Ícone do arquivo `.exe` / ícone da janela do app. |
| `--clean` | Apaga o cache da compilação anterior. |

## 4. Testar o executável

Rode `dist\SuporteTI.exe` em uma máquina **sem Python** e confira:

- [ ] O app abre e a ficha da Visão geral é preenchida (Windows, IP, disco...).
- [ ] Seção **Office** abre e mostra o botão de reinstalação (não execute em máquina de produção só para testar).
- [ ] Uma ação de administrador (ex.: Reiniciar Spooler) abre o UAC; ao clicar em **Não**, o app mostra erro (e não "concluído").
- [ ] Aba Outlook carrega as contas (com o Outlook aberto).
- [ ] "Abrir pasta de logs" (coluna Registro) abre `%LOCALAPPDATA%\SuporteTI\logs`.

## 5. Problemas comuns

| Sintoma | Solução |
|---|---|
| Antivírus/SmartScreen bloqueia o `.exe` | Normal em executáveis do PyInstaller sem assinatura. Compile com `.\build.ps1 -Pasta` (gera `dist\SuporteTI\` com vários arquivos, que é menos suspeito e abre mais rápido) e/ou **assine** o executável (veja abaixo). Peça à equipe de segurança para liberar pelo hash SHA-256. |
| O `.exe` fecha logo ao abrir | Compile sem `--windowed` para ver o erro no console. Na maioria das vezes é falta do `--collect-data customtkinter`. |
| `ModuleNotFoundError: win32com` | O `pip install` foi feito fora do `.venv`. Ative o ambiente (`.\.venv\Scripts\Activate.ps1`) e reinstale. |
| Abertura lenta (5 a 10 s) | Característica do `--onefile`, que extrai os arquivos a cada execução. Use `-Pasta`/`--onedir`. |

## 6. Recomendações de segurança para distribuir

1. **Não** compile com `--uac-admin`. Um app que roda inteiro como administrador
   não consegue falar com o Outlook do usuário (COM entre níveis de integridade
   diferentes), e qualquer falha nele vira uma falha com privilégio total. O app
   já pede o UAC apenas nas ações que precisam.
2. **Assine o executável** com um certificado de assinatura de código da empresa:
   ```powershell
   signtool sign /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 /a dist\SuporteTI.exe
   ```
3. Distribua por um local **somente leitura para usuários** (compartilhamento de
   rede com permissão de escrita só para a TI). Assim ninguém troca o `.exe` por
   uma versão adulterada.
4. Publique o **SHA-256** junto com o executável para quem baixar poder conferir.
5. Não versione `.venv\`, `build\` e `dist\`. Para recompilar, basta o
   `ferramentas_suporte_ti.py`, a pasta `suporte_ti\`, a pasta `OFFICE\`, o
   `requirements.txt`, o `build.ps1` e o `logo.ico`.
