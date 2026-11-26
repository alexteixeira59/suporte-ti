<#
    Compila o ferramentas_suporte_ti.py em um executável (dist\SuporteTI.exe).

    Uso (na pasta do projeto, em um PowerShell comum, SEM administrador):
        powershell -ExecutionPolicy Bypass -File .\build.ps1
        powershell -ExecutionPolicy Bypass -File .\build.ps1 -Pasta     # gera pasta em vez de .exe único
#>
param(
    [switch]$Pasta   # --onedir: abre mais rápido e gera menos falso positivo de antivírus
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { throw 'Python não encontrado no PATH. Instale pelo python.org marcando "Add python.exe to PATH".' }

if (-not (Test-Path '.\.venv\Scripts\python.exe')) {
    Write-Host '[1/3] Criando ambiente virtual (.venv)...' -ForegroundColor Cyan
    python -m venv .venv
}
$py = Resolve-Path '.\.venv\Scripts\python.exe'

Write-Host '[2/3] Instalando dependências...' -ForegroundColor Cyan
& $py -m pip install --disable-pip-version-check -q --upgrade pip
& $py -m pip install --disable-pip-version-check -q -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw 'Falha ao instalar as dependências.' }

$argumentos = @(
    '--noconfirm', '--clean', '--windowed',
    '--name', 'SuporteTI',
    '--collect-data', 'customtkinter'
)
$argumentos += if ($Pasta) { '--onedir' } else { '--onefile' }
if (Test-Path '.\logo.ico') {
    # --icon: ícone do .exe   |   --add-data: ícone da janela (resource_path)
    $argumentos += @('--icon', 'logo.ico', '--add-data', 'logo.ico;.')
} else {
    Write-Host 'Aviso: logo.ico não encontrado; o executável usará o ícone padrão.' -ForegroundColor Yellow
}
if (Test-Path '.\OFFICE\setup.exe') {
    # Instalador do Office 2024 usado pela aba Office (PASTA_INSTALADOR_OFFICE)
    $argumentos += @('--add-data', 'OFFICE\setup.exe;OFFICE', '--add-data', 'OFFICE\Configuracao.xml;OFFICE')
} else {
    Write-Host 'Aviso: OFFICE\setup.exe não encontrado; o botão de reinstalar o Office não funcionará.' -ForegroundColor Yellow
}
$argumentos += 'ferramentas_suporte_ti.py'

Write-Host '[3/3] Compilando com PyInstaller...' -ForegroundColor Cyan
& $py -m PyInstaller @argumentos
if ($LASTEXITCODE -ne 0) { throw 'O PyInstaller falhou. Veja as mensagens acima.' }

$saida = if ($Pasta) { '.\dist\SuporteTI\SuporteTI.exe' } else { '.\dist\SuporteTI.exe' }

$hash = (Get-FileHash -LiteralPath $saida -Algorithm SHA256).Hash
Write-Host ''
Write-Host "Pronto: $((Resolve-Path $saida).Path)" -ForegroundColor Green
Write-Host "SHA-256: $hash"
