# Gerador do BI de Operacoes PJC — em PowerShell puro.
#
# Le a planilha v2 (.xlsx) e escreve dados_bi.js, que o painel HTML consome.
# NAO precisa de Python nem de Excel instalado: o .xlsx e um ZIP com XMLs
# dentro, e o Windows ja sabe abrir os dois.
#
# Uso:  powershell -ExecutionPolicy Bypass -File Gerar-BI.ps1 [-Planilha caminho]

param(
    [string]$Planilha = "",
    [string]$Saida    = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$AQUI = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Saida) { $Saida = Join-Path $AQUI "dados_bi.js" }

# ---------------------------------------------------------------- planilha
function Achar-Planilha {
    # Procura AQUI primeiro (pacote autocontido dos diretores) e depois na
    # pasta de cima (instalacao completa, onde os scripts ficam em _SISTEMA).
    foreach ($dir in @($AQUI, (Split-Path -Parent $AQUI))) {
        if (-not (Test-Path $dir)) { continue }
        $cands = @(Get-ChildItem $dir -Filter *.xlsx -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -notlike '~$*' })
        # Prefere a versao de numero mais alto (v3 > v2), para nao ficar preso
        # a um nome fixo quando a planilha virar de versao.
        $comVer = @($cands | Where-Object { $_.Name -match 'v(\d+)' } |
                    Sort-Object { [int]([regex]::Match($_.Name, 'v(\d+)').Groups[1].Value) } -Descending)
        if ($comVer.Count -gt 0) { return $comVer[0].FullName }
        $rel = @($cands | Where-Object { $_.Name -match 'RELA' })
        if ($rel.Count -gt 0) { return $rel[0].FullName }
    }
    return $null
}

if (-not $Planilha) { $Planilha = Achar-Planilha }
if (-not $Planilha -or -not (Test-Path $Planilha)) {
    Write-Host "ERRO: planilha nao encontrada." -ForegroundColor Red
    exit 1
}
Write-Host ("Lendo: " + (Split-Path $Planilha -Leaf))

# ------------------------------------------------------------ leitura xlsx
# Copia para o temp: se a planilha estiver aberta no Excel, ler direto falha.
$tmp = Join-Path $env:TEMP ("bi_" + [Guid]::NewGuid().ToString("N") + ".xlsx")
Copy-Item -LiteralPath $Planilha -Destination $tmp -Force
$zip = [IO.Compression.ZipFile]::OpenRead($tmp)

function Ler-Xml($zipArq, $nome) {
    $e = $zipArq.Entries | Where-Object { $_.FullName -eq $nome }
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open())
    try { [xml]$sr.ReadToEnd() } finally { $sr.Close() }
}

# textos ficam num dicionario compartilhado, referenciado por indice
$textos = @()
$ss = Ler-Xml $zip 'xl/sharedStrings.xml'
if ($ss) {
    foreach ($si in $ss.sst.si) {
        if ($si.t -is [string])        { $textos += $si.t }
        elseif ($si.t.'#text')         { $textos += $si.t.'#text' }
        elseif ($si.r)                 { $textos += (($si.r | ForEach-Object {
                                              if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '') }
        else                           { $textos += "" }
    }
}

# nome da aba -> arquivo xml da aba
$wbx  = Ler-Xml $zip 'xl/workbook.xml'
$rels = Ler-Xml $zip 'xl/_rels/workbook.xml.rels'
$abaArquivo = @{}
$ordem = @()
foreach ($s in $wbx.workbook.sheets.sheet) {
    $rid = $s.id; if (-not $rid) { $rid = $s.'r:id' }
    $t = $rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }
    if ($t) {
        $alvo = $t.Target -replace '^/xl/', '' -replace '^xl/', ''
        $abaArquivo[$s.name] = "xl/$alvo"
        $ordem += $s.name
    }
}

function Col-Letra($ref) { ($ref -replace '\d','') }
function Col-Num($letra) {
    $n = 0
    foreach ($ch in $letra.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    $n
}

function Ler-Aba($nomeAba) {
    <#  Devolve uma lista de hashtables: cabecalho -> valor  #>
    $arq = $abaArquivo[$nomeAba]
    if (-not $arq) { return @() }
    $sh = Ler-Xml $zip $arq
    if (-not $sh) { return @() }

    # 1) monta grade linha -> (coluna -> valor)
    $grade = @{}
    foreach ($row in $sh.worksheet.sheetData.row) {
        $r = [int]$row.r
        $cels = @{}
        foreach ($c in $row.c) {
            $v = $c.v
            if ($c.t -eq 's' -and $v -ne $null -and $v -ne '') {
                $i = [int]$v
                $v = if ($i -lt $textos.Count) { $textos[$i] } else { "" }
            }
            elseif ($c.t -eq 'inlineStr') { $v = $c.is.t }
            if ($v -ne $null -and $v -ne '') { $cels[(Col-Num (Col-Letra $c.r))] = $v }
        }
        if ($cels.Count) { $grade[$r] = $cels }
    }
    if ($grade.Count -eq 0) { return @() }

    # 2) acha a linha do cabecalho: e a que tem CLIENTE (ou GERENTE, no
    #    historico). Nao fixar na linha 4 — inserir linhas no topo da planilha
    #    e comum e deslocaria tudo, deixando o painel cego.
    $LCAB = 0
    $cab = @{}
    foreach ($r in ($grade.Keys | Sort-Object)) {
        if ($r -gt 30) { break }
        $tmp = @{}
        foreach ($k in $grade[$r].Keys) { $tmp[(Sem-Acento $grade[$r][$k])] = $k }
        if ($tmp.ContainsKey('CLIENTE')) { $LCAB = $r; $cab = $tmp; break }
    }
    if ($LCAB -eq 0) { return @() }

    # 3) linhas de dados
    $saida = @()
    foreach ($r in ($grade.Keys | Where-Object { $_ -gt $LCAB } | Sort-Object)) {
        $cels = $grade[$r]
        $obj = @{}
        foreach ($nome in $cab.Keys) {
            $ci = $cab[$nome]
            $obj[$nome] = if ($cels.ContainsKey($ci)) { $cels[$ci] } else { $null }
        }
        if ($obj['CLIENTE'] -and ([string]$obj['CLIENTE']).Trim()) { $saida += $obj }
    }
    return $saida
}

# ------------------------------------------------------------- utilitarios
function Como-Numero($v) {
    if ($v -eq $null -or $v -eq '') { return $null }
    $d = 0.0
    if ([double]::TryParse([string]$v, [Globalization.NumberStyles]::Any,
                           [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return $null
}

function Como-Data($v) {
    # No Excel a data e um numero de dias desde 30/12/1899.
    $n = Como-Numero $v
    if ($n -eq $null -or $n -le 0) { return "" }
    try { ([datetime]"1899-12-30").AddDays([int]$n).ToString("yyyy-MM-dd") } catch { "" }
}

function Sem-Acento($t) {
    # Normaliza via .NET: evita escrever letras acentuadas no proprio script,
    # que o Windows PowerShell 5.1 leria errado se o arquivo nao tiver BOM.
    if ($null -eq $t) { return "" }
    $norm = ([string]$t).Trim().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $norm.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    (($sb.ToString().ToUpper()) -replace '\s+',' ').Trim()
}

$ABAS_SISTEMA = @('PARAMETROS','DASHBOARD','METAS','HISTORICO','COMO USAR')
# Os cinco primeiros sao etapas do funil; APROVADO, REPROVADO e FINALIZADO sao
# desfechos (a operacao saiu do funil).
$STATUS_CANON = @('CADASTRO','ANALISE DE CREDITO','COMITE','APROVACAO',
                  'ASSINATURA DE CONTRATO','APROVADO','REPROVADO','FINALIZADO')
$ETAPAS_FUNIL = $STATUS_CANON[0..4]
$MESES = @('JANEIRO','FEVEREIRO','MARCO','ABRIL','MAIO','JUNHO','JULHO',
           'AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO')

# ------------------------------------------------------------- montar dados
$ops = @()

function Montar-Op($linha, $gerente, $aba, $ehHistorico) {
    $valor   = Como-Numero $linha['VALOR']
    $fee     = Como-Numero $linha['FEE']
    $receita = Como-Numero $linha['RECEITA']
    # RECEITA e formula: se o arquivo nao trouxer o valor calculado, refaz aqui
    if ($receita -eq $null -and $valor -ne $null -and $fee -ne $null) { $receita = $valor * $fee }

    [ordered]@{
        gerente        = $gerente
        aba            = $aba
        historico      = $ehHistorico
        origem         = [string]$linha['ORIGEM']
        cliente        = ([string]$linha['CLIENTE']).Trim()
        operacao       = ([string]$linha['OPERACAO']).Trim()
        valor          = $valor
        taxa           = $fee
        receita        = $receita
        mandato        = ([string]$linha['MANDATO']).Trim()
        mandatoTexto   = ([string]$linha['MANDATO']).Trim()
        instituicao    = ([string]$linha['INSTITUICAO']).Trim()
        indicador      = ([string]$linha['PARCEIRO INDICADOR']).Trim()
        data           = Como-Data $linha['ATUALIZACAO']
        dataFechamento = Como-Data $linha['DATA FECHAMENTO']
        fluxo          = ([string]$linha['FLUXO']).Trim()
        status         = Sem-Acento $linha['STATUS']
        motivoRecusa   = ([string]$linha['MOTIVO DE RECUSA']).Trim()
        obs            = ([string]$linha['OBSERVACAO']).Trim()
        unico          = $true      # na v2 cada linha e uma operacao
        grupoTam       = 1
        dupExata       = $false
    }
}

foreach ($aba in $ordem) {
    $chave = Sem-Acento $aba
    if ($ABAS_SISTEMA -contains $chave) { continue }
    $linhas = @(Ler-Aba $aba)
    foreach ($l in $linhas) { $ops += (Montar-Op $l $aba $aba $false) }
    Write-Host ("  {0,-24} {1,4} operacoes" -f $aba, $linhas.Count)
}

# HISTORICO: reprovados do ciclo atual + linhas das abas antigas.
# So as antigas sao historico de fato; os reprovados contam na taxa de aprovacao.
$abaHist = $ordem | Where-Object { (Sem-Acento $_) -eq 'HISTORICO' } | Select-Object -First 1
if ($abaHist) {
    $linhas = @(Ler-Aba $abaHist)
    foreach ($l in $linhas) {
        # Reprovado e Finalizado sao do ciclo atual e continuam contando nos
        # indicadores; so as linhas das abas antigas sao historico de fato.
        $origemNorm = Sem-Acento $l['ORIGEM']
        $ehHist = ($origemNorm -ne 'REPROVADO' -and $origemNorm -ne 'FINALIZADO')
        $ops += (Montar-Op $l ([string]$l['GERENTE']).Trim() $abaHist $ehHist)
    }
    Write-Host ("  {0,-24} {1,4} operacoes [historico]" -f $abaHist, $linhas.Count)
}

# ------------------------------------------------------------------- metas
$metas = [ordered]@{}
if ($abaArquivo.ContainsKey('METAS')) {
    $sh = Ler-Xml $zip $abaArquivo['METAS']
    $grade = @{}
    foreach ($row in $sh.worksheet.sheetData.row) {
        $r = [int]$row.r; $cels = @{}
        foreach ($c in $row.c) {
            $v = $c.v
            if ($c.t -eq 's' -and $v -ne $null -and $v -ne '') {
                $i = [int]$v; $v = if ($i -lt $textos.Count) { $textos[$i] } else { "" }
            }
            if ($v -ne $null -and $v -ne '') { $cels[(Col-Num (Col-Letra $c.r))] = $v }
        }
        if ($cels.Count) { $grade[$r] = $cels }
    }
    $atual = $null
    foreach ($r in ($grade.Keys | Sort-Object)) {
        $rot = Sem-Acento $grade[$r][1]
        if (-not $rot) { continue }
        if ($rot -eq 'META' -and $atual) {
            $vals = @()
            for ($c = 2; $c -le 13; $c++) {
                $n = Como-Numero $grade[$r][$c]
                $vals += $(if ($n -eq $null) { 0 } else { $n })
            }
            if (($vals | Measure-Object -Sum).Sum -gt 0) {
                $metas[$atual] = [ordered]@{ meta = $vals; realizadoPlanilha = $null }
            }
            $atual = $null
        }
        elseif ($rot -notin @('META','REALIZADO','% ATINGIDO','GERENTE / LINHA','ANO BASE:','COMO FUNCIONA') `
                -and -not $rot.StartsWith('TOTAL PJC') -and -not $rot.StartsWith('•')) {
            $atual = $rot
        }
    }
}

$zip.Dispose()
Remove-Item $tmp -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------- cobertura
$cobertura = [ordered]@{}
foreach ($g in ($ops | Where-Object { -not $_.historico } | ForEach-Object { $_.gerente } | Select-Object -Unique)) {
    $doGerente = @($ops | Where-Object { -not $_.historico -and $_.gerente -eq $g })
    $cobertura[$g] = [ordered]@{
        total  = $doGerente.Count
        comFee = @($doGerente | Where-Object { $_.taxa }).Count
    }
}

$avisos = @()
$semValor = @($ops | Where-Object { -not $_.historico -and $_.valor -eq $null }).Count
if ($semValor) { $avisos += "$semValor operacoes sem valor preenchido - ficam fora dos totais." }
$semFee = @($ops | Where-Object { -not $_.historico -and -not $_.taxa }).Count
if ($semFee) { $avisos += "$semFee operacoes sem fee - entram no volume, nao na receita." }

$instituicoes = @($ops | ForEach-Object { $_.instituicao } |
                  Where-Object { $_ } | Select-Object -Unique | Sort-Object)

$dados = [ordered]@{
    geradoEm             = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    arquivoFonte         = (Split-Path $Planilha -Leaf)
    versaoPlanilha       = "v2"
    geradoPor            = "PowerShell"
    operacoes            = $ops
    metas                = $metas
    metaVolume           = @{}
    meses                = $MESES
    statusCanon          = $STATUS_CANON
    etapasFunil          = $ETAPAS_FUNIL
    instituicoesOficiais = $instituicoes
    cobertura            = $cobertura
    avisos               = $avisos
}

$json = $dados | ConvertTo-Json -Depth 6 -Compress
[IO.File]::WriteAllText($Saida, "window.DADOS_BI = $json;`r`n", (New-Object Text.UTF8Encoding $false))

$kb = [math]::Round((Get-Item $Saida).Length / 1KB)
$ativas = @($ops | Where-Object { -not $_.historico }).Count
Write-Host ""
Write-Host ("Gerado: dados_bi.js ({0} KB) - {1} operacoes ({2} ativas, {3} no historico)" -f `
            $kb, $ops.Count, $ativas, ($ops.Count - $ativas)) -ForegroundColor Green
foreach ($a in $avisos) { Write-Host "  aviso: $a" -ForegroundColor DarkYellow }
