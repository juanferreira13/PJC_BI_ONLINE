# Baixa a planilha v2 publicada no OneDrive, sem navegador e sem login.
#
# Como funciona: o link compartilhado do OneDrive nao entrega o arquivo de
# cara. Primeiro ele precisa ser "resgatado" - a visita ao link cria uma
# sessao anonima com permissao de leitura. So depois o endereco de download
# autoriza. Por isso sao dois passos, reaproveitando a mesma sessao.
#
# Se qualquer coisa falhar, o BI segue com a copia que ja esta na pasta:
# este script nunca derruba o processo.

$ErrorActionPreference = "Stop"
$AQUI = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---------------------------------------------------------------------------
# CONFIGURACAO - trocar aqui se a planilha mudar de lugar.
#   $LINK   : link do botao Compartilhar (precisa ser "Qualquer pessoa com o link")
#   $ID_ARQ : o identificador que aparece no link, entre chaves
$LINK   = "https://1drv.ms/x/c/8B3E728BD4125820/IQBv_jqRcFDkQrJAekilAC2DAdScEKFTq-V2bI5TGdrJ4qY?e=gatnWp"
$ID_ARQ = "913afe6f-5070-42e4-b240-7a48a5002d83"
$CONTA  = "8B3E728BD4125820"
# ---------------------------------------------------------------------------

$destino = Join-Path $AQUI "RELACAO DE OPERACAO - PJC v3.xlsx"
$UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"

function Usar-Copia-Local($motivo) {
    Write-Host "        $motivo" -ForegroundColor Yellow
    if (Test-Path $destino) {
        $h = [math]::Round(((Get-Date) - (Get-Item $destino).LastWriteTime).TotalHours, 1)
        Write-Host "        Sigo com a copia da pasta (de $h h atras)."
        exit 0
    }
    Write-Host "        E nao ha copia na pasta." -ForegroundColor Red
    exit 1
}

if (-not $LINK) { Usar-Copia-Local "Sem link configurado." }

try {
    # --- passo 1: resgatar o link, criando a sessao anonima ---
    $ProgressPreference = 'SilentlyContinue'
    $sessao = $null
    Invoke-WebRequest -Uri $LINK -UseBasicParsing -UserAgent $UA `
        -SessionVariable sessao -TimeoutSec 60 | Out-Null

    # --- passo 2: baixar o arquivo com a mesma sessao ---
    $urlDownload = "https://onedrive.live.com/personal/$CONTA/_layouts/15/download.aspx?UniqueId=$ID_ARQ"
    $tmp = Join-Path $env:TEMP ("plan_" + [Guid]::NewGuid().ToString("N") + ".xlsx")
    Invoke-WebRequest -Uri $urlDownload -OutFile $tmp -UseBasicParsing -UserAgent $UA `
        -WebSession $sessao -TimeoutSec 120

    # --- confere que veio mesmo uma planilha (e nao uma pagina de login) ---
    $fs = [IO.File]::OpenRead($tmp)
    $cab = New-Object byte[] 2
    [void]$fs.Read($cab, 0, 2)
    $fs.Close()
    $tam = (Get-Item $tmp).Length

    if ($tam -lt 10000 -or $cab[0] -ne 0x50 -or $cab[1] -ne 0x4B) {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Usar-Copia-Local "O link nao devolveu uma planilha (permissao mudou?)."
    }

    Move-Item $tmp $destino -Force
    $mb = [math]::Round($tam / 1MB, 2)
    Write-Host "        Planilha atualizada do OneDrive ($mb MB)." -ForegroundColor Green
    exit 0
}
catch {
    Usar-Copia-Local ("Nao consegui baixar: " + $_.Exception.Message)
}
