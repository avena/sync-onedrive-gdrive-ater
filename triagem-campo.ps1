<#
.SYNOPSIS
    Triagem de Fotos de Campo (Inbox Drive / Syncthing -> OneDrive ATER) + GPS EXIF + runs\
.DESCRIPTION
    Fonte primaria: inbox espelhada do Google Drive pessoal
    G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<Atividade>\<AAAA-MM-DD-NOME - RESUMO>\
    com pares .txt/.json (schema_version 1.0). Fallback: visitas.json manual ou colagem (FIM).
    Etapa 1: revisao interativa (TUI) -> cria pastas no OneDrive e grava APENAS o .txt limpo
    (sem YAML); o .json fica so na inbox (originais INTACTOS). Na ATV 1 copia anexos
    Mapa/Calendario Agricola (.png/.jpg/.jpeg). Registra append em _processados.json por ano.
    Etapa 2: busca cirurgica pelas datas do relatorio (so <AAAA>\<MM> alvo; nome/hora/GPS
    lidos direto do _visita.json, zero EXIF nas visitas; avulsas/brutas pre-filtradas pelo
    nome *AAAAMMDD*). Visitas ja viram lotes pre-agrupados; avulsas agrupam por 40 min.
    Pre-sugere o produtor por horario/periodo (Enter ou S aceita), distribui por TAG e gera
    coordenadas_gps.txt. Registra manifesto de auditoria em runs\ (nunca apagado).
.PARAMETER Colar
    Forca colagem do JSON no terminal (ignora inbox e visitas.json).
.PARAMETER Conferir
    Varre FISICAMENTE a inbox do Drive (todas as pastas, com status NOVA /
    TRIADA - PENDENTE SYNC / SINCRONIZADA cruzado do _processados.json), confere
    TXT e fotos no destino [OD] e permite reprocessar qualquer pasta: recopiar
    TXT/anexos da inbox [GD], triar pastas NOVA e/ou copiar apenas as fotos
    faltantes (S/N; nunca apaga nem força fotos a mais).
.PARAMETER VisitasJson
    Caminho do arquivo de visitas manual (padrao: visitas.json ao lado do script).
.NOTES
    Requer PowerShell 7. System.Drawing le EXIF de JPG/JPEG; .heic/.mp4 caem no LastWriteTime.
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [switch]$Colar,
    [switch]$Conferir,
    [string]$VisitasJson = "$PSScriptRoot\visitas.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ================= CONFIGURACOES =================
$OneDriveBase           = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
# Raízes de fotos em ordem de prioridade (dedup por nome: a 1ª raiz vence):
# 1. organizada por visita (organizar-camera.ps1)  2. disco externo  3. bruta Syncthing (recém-chegadas)
$SyncthingFotos         = @(
    "C:\Users\randomacesso\Sync\Camera_Celular_Organizada",
    "E:\Sync\Camera_Celular_Organizada",
    "C:\Users\randomacesso\Sync\Camera_Celular\Camera"
)
$InboxBase              = "G:\Meu Drive"                                   # <-- Google Drive pessoal (inbox espelhada)
$IntervaloVisitaMinutos = 40
$ExtensoesFoto          = @(".jpg", ".jpeg", ".png", ".heic", ".mp4")
$LogFile                = Join-Path $PSScriptRoot "triagem-log.txt"
$RunsDir                = Join-Path $PSScriptRoot "runs"                   # permanente: NUNCA apagar

# Nomes EXATOS das pastas locais [OD] (menu de edicao). Padrao do usuario: "Juquia" SEM acento.
$Cidades = @(
    "Barra do Turvo", "Cananéia", "Jacupiranga", "Juquia", "Pariquera-Açu", "Registro"
)
# A inbox [GD] e o OneDrive usam "Juquia" (sem acento). O mapa e rede de seguranca p/
# relatorios/metadados antigos que ainda tragam "Juquiá" com acento. Demais cidades: mesmo nome.
$CidadesOneDrive = @{ 'Juquiá' = 'Juquia' }
$Atividades = @(
    "ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3",
    "ATIVIDADE INDIVIDUAL 4", "ATIVIDADE INDIVIDUAL 5", "ATIVIDADE INDIVIDUAL 6",
    "PROJETO PRODUTIVO"
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ================= HELPERS BASICOS =================
function Test-CampoSeguro {
    param([string]$Valor)
    if ([string]::IsNullOrWhiteSpace($Valor)) { return $false }
    if ($Valor -match '\.\.' -or $Valor -match '[/\\:]') { return $false }
    if ($Valor -match '[\. ]$') { return $false }   # Windows apaga ponto/espaco final -> caminho diverge
    return $true
}

function Write-TriagemLog {
    param([string]$Mensagem)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $LogFile -Value "$timestamp | $Mensagem" -Encoding UTF8
}

function Clear-KeyBuffer {
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch { }
}

function Get-PropValue {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Add-SufixoResumo {
    param([string]$Nome)
    if ([string]::IsNullOrWhiteSpace($Nome)) { return $Nome }
    if ($Nome -match '\s*-\s*RESUMO$') { return $Nome }
    return "$Nome - RESUMO"
}

# Nome da cidade no destino [OD]: normaliza "Juquiá" (inbox/relatorios) -> "Juquia" (pasta local padrao)
function Get-NomeCidadeOneDrive {
    param([string]$Cidade)
    if ($CidadesOneDrive.ContainsKey($Cidade)) { return $CidadesOneDrive[$Cidade] }
    return $Cidade
}

# Nome da pasta = prefixo DD-AAAA-MM- (dia primeiro, p/ ordenar por dia dentro da atividade)
# + NOME + - RESUMO. Remove prefixo antigo (AAAA-MM-DD-/AAAA-MM-) ou novo (DD-AAAA-MM-) antes.
function Format-NomePastaProdutor {
    param([string]$Nome, [string]$Data)
    if ([string]::IsNullOrWhiteSpace($Nome)) { return $Nome }
    $Nome = $Nome.Trim()
    $Nome = $Nome -replace '^(\d{4}-\d{2}(-\d{2})?-|\d{2}-\d{4}-\d{2}-)', ''
    if ($Data -match '^\d{4}-\d{2}-\d{2}$') { $Nome = "$($Data.Substring(8, 2))-$($Data.Substring(0, 7))-$Nome" }
    return Add-SufixoResumo $Nome
}

function Remove-YamlHeader {
    param([string]$Texto)
    if ($null -eq $Texto) { return $null }
    if ($Texto -match '(?s)^\s*---\r?\n.*?\r?\n---\r?\n?(.*)$') { return $Matches[1] }
    return $Texto
}

function Read-MetadadosJson {
    param([string]$Path)
    try { $obj = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return $null }
    $meta = Get-PropValue $obj 'metadados'
    if (-not $meta) { $meta = $obj }
    return @{
        data           = [string](Get-PropValue $meta 'data')
        periodo        = [string](Get-PropValue $meta 'periodo')
        hora_inicio    = [string](Get-PropValue $meta 'hora_inicio')
        hora_fim       = [string](Get-PropValue $meta 'hora_fim')
        cidade         = [string](Get-PropValue $meta 'cidade')
        atividade      = [string](Get-PropValue $meta 'atividade')
        produtor_pasta = [string](Get-PropValue $meta 'produtor_pasta')
    }
}

function Read-MetadadosTxt {
    param([string]$Path)
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $campos = @{}
    $body = $raw
    if ($raw -match '(?s)^\s*---\r?\n(.*?)\r?\n---\r?\n?(.*)$') {
        $yaml = $Matches[1]; $body = $Matches[2]
        foreach ($linha in ($yaml -split "\r?\n")) {
            if ($linha -match '^\s*([A-Za-z_]+)\s*:\s*(.*)$') {
                $campos[$Matches[1]] = $Matches[2].Trim().Trim('"').Trim("'")
            }
        }
    }
    return @{ Campos = $campos; Body = $body }
}

# Texto puro p/ o .txt final: sem Markdown (**, *, #, crases); bullets '- '/'* ' viram '• '.
function Remove-MarkdownBasico {
    param([string]$Texto)
    if ($null -eq $Texto) { return $null }
    $t = $Texto -replace '(?m)^\s{0,3}#{1,6}\s*', ''
    $t = $t -replace '(?m)^\s*[-*]\s+', '• '
    return ($t -replace '\*|__|`', '')
}

# Remove a secao 10 (encaminhamentos/compromissos) do corpo: do titulo "10. ..." ate a
# proxima secao numerada ou fim do texto. Sem secao 10 -> devolve o texto intacto.
function Remove-Secao10 {
    param([string]$Texto)
    if ($null -eq $Texto) { return $null }
    $linhas = @($Texto -split "\r?\n")
    $ini = -1; $fim = $linhas.Count
    for ($i = 0; $i -lt $linhas.Count; $i++) {
        if ($ini -lt 0 -and $linhas[$i] -match '^\s*10\.\s') { $ini = $i; continue }
        if ($ini -ge 0 -and $linhas[$i] -match '^\s*\d+\.\s') { $fim = $i; break }
    }
    if ($ini -lt 0) { return $Texto }
    $antes = if ($ini -gt 0) { $linhas[0..($ini - 1)] } else { @() }
    $depois = if ($fim -lt $linhas.Count) { $linhas[$fim..($linhas.Count - 1)] } else { @() }
    return (@($antes) + @($depois) -join "`n").TrimEnd()
}

# Monta o corpo do relatorio (texto puro, padrao ANATER/Bem Viver v2) do JSON schema 1.0:
# cabecalho institucional + secoes 7/8/9 com orientacoes em '• ', quebra dupla entre blocos.
# Sem YAML, sem Markdown e SEM secao 10 (encaminhamentos ficam apenas no JSON).
function ConvertFrom-RelatorioJson {
    param($Json)   # caminho do .json ou objeto ja parseado
    $obj = if ($Json -is [string]) { try { Get-Content -LiteralPath $Json -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return $null } } else { $Json }
    if ($null -eq $obj) { return $null }
    $inst = Get-PropValue $obj 'institucional'
    $rel = Get-PropValue $obj 'relatorio'
    if (-not $inst -and -not $rel) { return $null }
    $blocos = [System.Collections.Generic.List[string]]::new()
    $cab = [System.Collections.Generic.List[string]]::new()
    $ent = [string](Get-PropValue $inst 'entidade_executora')
    if ($ent) { $cab.Add($ent) }
    foreach ($par in @(@('Agente ATER', 'agente_ater'), @('Instrumento', 'instrumento'), @('Beneficiário', 'beneficiario'), @('Localização', 'localizacao'), @('Data da etapa', 'data_etapa'))) {
        $v = [string](Get-PropValue $inst $par[1])
        if ($v) { $cab.Add("$($par[0]): $v") }
    }
    if ($cab.Count -gt 0) { $blocos.Add(($cab -join "`n")) }
    foreach ($eixo in @(@('7. EIXO PRODUTIVO', 'eixo_produtivo', 'ORIENTAÇÕES TÉCNICAS:', 'orientacoes_tecnicas'),
                        @('8. EIXO SOCIAL', 'eixo_social', 'ORIENTAÇÕES SOCIAIS:', 'orientacoes_sociais'),
                        @('9. EIXO AMBIENTAL', 'eixo_ambiental', 'ORIENTAÇÕES AMBIENTAIS:', 'orientacoes_ambientais'))) {
        $texto = [string](Get-PropValue $rel $eixo[1])
        $itens = @(Get-PropValue $rel $eixo[3])
        if (-not $texto -and $itens.Count -eq 0) { continue }
        $blocos.Add($eixo[0])
        if ($texto) { $blocos.Add($texto) }
        if ($itens.Count -gt 0) {
            $blocos.Add($eixo[2])
            foreach ($it in $itens) {
                $s = [string]$it
                if (-not $s.Trim()) { continue }
                $s = Remove-MarkdownBasico ($s -replace '^\s*[-*]\s+', '')   # remove bullet Markdown antes do '• ' (sem duplicar)
                $blocos.Add("• $s")
            }
        }
    }
    if ($blocos.Count -eq 0) { return $null }
    return (Remove-MarkdownBasico ($blocos -join "`n`n"))
}

# Grava o .txt do par na inbox [GD] com bloco YAML no topo, montado do JSON (schema 1.0).
# ADITIVO: so cria se o .txt nao existir; nunca sobrescreve/apaga originais. Atomico (.tmp + Move).
function Write-TxtInboxDoJson {
    param([string]$JsonPath)
    $txtPath = Join-Path (Split-Path $JsonPath -Parent) ([System.IO.Path]::GetFileNameWithoutExtension($JsonPath) + '.txt')
    if (Test-Path -LiteralPath $txtPath) { return $txtPath }
    try { $obj = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return $null }
    $body = ConvertFrom-RelatorioJson -Json $obj
    if (-not $body) { return $null }
    $meta = Get-PropValue $obj 'metadados'
    if (-not $meta) { $meta = $obj }
    $linhas = [System.Collections.Generic.List[string]]::new()
    $linhas.Add('---')
    foreach ($k in @('data', 'periodo', 'hora_inicio', 'hora_fim', 'cidade', 'atividade', 'produtor_pasta')) {
        $val = [string](Get-PropValue $meta $k)
        if (-not $val) { $val = 'null' }
        $linhas.Add("${k}: $val")
    }
    $linhas.Add('---')
    $tmp = "$txtPath.tmp"
    Set-Content -LiteralPath $tmp -Value (($linhas -join "`n") + "`n" + $body) -Encoding UTF8
    Move-Item -LiteralPath $tmp -Destination $txtPath -Force
    Write-TriagemLog "MONTADO TXT [GD] do JSON: $txtPath"
    return $txtPath
}

function New-Visita {
    param($Data, $Periodo, $HoraInicio, $HoraFim, $Cidade, $Atividade, $ProdutorPasta,
          $SourceTxt, $SourceJson, $TxtBody, $OrigemInbox, $Ano, $MM, $LeafName, $PastaInbox)
    return [PSCustomObject]@{
        data = $Data; periodo = $Periodo; hora_inicio = $HoraInicio; hora_fim = $HoraFim
        cidade = $Cidade; atividade = $Atividade; produtor_pasta = $ProdutorPasta; tag = $null
        SourceTxt = $SourceTxt; SourceJson = $SourceJson; TxtBody = $TxtBody
        OrigemInbox = $OrigemInbox; Ano = $Ano; MM = $MM; LeafName = $LeafName
        PastaInbox = $PastaInbox; CaminhoDestino = $null
    }
}

# ================= FONTE PRIMARIA: INBOX ESPELHADA (VARREDURA FISICA DO DRIVE) =================
# Lista FISICAMENTE todas as pastas <AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<Atividade>\<Produtor>,
# independentemente de ja constarem em _processados.json. O JSON serve APENAS para cruzar o status:
#   NOVA | TRIADA - PENDENTE SYNC | SINCRONIZADA. O registro (se houver) vai junto, com registro_path.
function Get-InboxPastasFisicas {
    $pastas = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $InboxBase)) { return , @() }
    foreach ($ano in @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' } | Sort-Object Name)) {
        $inbox = Join-Path $ano.FullName '_inbox_relatorio_ater'
        if (-not (Test-Path -LiteralPath $inbox)) { continue }
        $regs = @{}
        $regPath = Join-Path $inbox '_processados.json'
        if (Test-Path -LiteralPath $regPath) {
            try {
                foreach ($r in @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json)) {
                    $pi = [string](Get-PropValue $r 'pasta_inbox')
                    if ($pi) {
                        Add-Member -InputObject $r -NotePropertyName registro_path -NotePropertyValue $regPath -Force
                        $regs[$pi.TrimEnd('\')] = $r   # ultimo registro vence (reprocessamento faz append)
                    }
                }
            } catch { }
        }
        # (A pasta legada _processados\ ja fica de fora: o filtro de meses so aceita ^\d{2}$.)
        foreach ($mm in @(Get-ChildItem -LiteralPath $inbox -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{2}$' } | Sort-Object Name)) {
            foreach ($cid in @(Get-ChildItem -LiteralPath $mm.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
                foreach ($ativ in @(Get-ChildItem -LiteralPath $cid.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
                    foreach ($prod in @(Get-ChildItem -LiteralPath $ativ.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
                        $r = $regs[$prod.FullName.TrimEnd('\')]
                        $status = if (-not $r) { 'NOVA' }
                                  elseif (Get-PropValue $r 'sincronizada') { 'SINCRONIZADA' }
                                  else { 'TRIADA - PENDENTE SYNC' }
                        $pastas.Add([PSCustomObject]@{
                            Ano = $ano.Name; MM = $mm.Name; Cidade = $cid.Name; Atividade = $ativ.Name
                            Produtor = $prod.Name; PastaInbox = $prod.FullName; Status = $status; Registro = $r
                        })
                    }
                }
            }
        }
    }
    return , @($pastas)
}

# Exibicao da varredura fisica: Mes | Cidade | Atividade | Produtor + status + Origem [GD] completa.
function Show-InboxFisico {
    param([array]$Pastas)
    $ultGrupo = ''
    foreach ($p in @($Pastas | Sort-Object Ano, MM, Cidade, Atividade, Produtor)) {
        $grupo = "$($p.Ano)\$($p.MM) | $($p.Cidade) | $($p.Atividade)"
        if ($grupo -ne $ultGrupo) {
            Write-Host "`n  $grupo" -ForegroundColor Cyan
            $ultGrupo = $grupo
        }
        $cor = switch ($p.Status) { 'NOVA' { 'Green' } 'SINCRONIZADA' { 'DarkGreen' } default { 'Yellow' } }
        Write-Host "   [$($p.Status)] $($p.MM) | $($p.Cidade) | $($p.Atividade) | $($p.Produtor)" -ForegroundColor $cor
        Write-Host "      Origem [GD]: $($p.PastaInbox)" -ForegroundColor DarkGray
    }
}

# Visitas NOVAS (idempotencia): so pastas fisicas ainda nao registradas em _processados.json.
function Get-InboxVisitas {
    $visitas = [System.Collections.Generic.List[object]]::new()
    foreach ($p in (Get-InboxPastasFisicas)) {
        if ($p.Status -ne 'NOVA') { continue }
        $leaf = $p.Produtor
        $arqs = @(Get-ChildItem -LiteralPath $p.PastaInbox -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @('.txt', '.json') })
        if ($arqs.Count -eq 0) { continue }
        $grupos = @($arqs | Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })
        foreach ($g in $grupos) {
            $txt  = $g.Group | Where-Object { $_.Extension -eq '.txt' }  | Select-Object -First 1
            $json = $g.Group | Where-Object { $_.Extension -eq '.json' } | Select-Object -First 1

            $data = $null; $periodo = $null; $hi = $null; $hf = $null
            $cidade = $p.Cidade; $atividade = $p.Atividade; $pasta = $leaf; $body = $null

            if ($json) {
                $m = Read-MetadadosJson -Path $json.FullName
                if ($m) {
                    if ($m.data)           { $data = $m.data }
                    if ($m.periodo)        { $periodo = $m.periodo }
                    if ($m.hora_inicio)    { $hi = $m.hora_inicio }
                    if ($m.hora_fim)       { $hf = $m.hora_fim }
                    if ($m.cidade)         { $cidade = $m.cidade }
                    if ($m.atividade)      { $atividade = $m.atividade }
                    if ($m.produtor_pasta) { $pasta = $m.produtor_pasta }
                }
            }
            if ($txt) {
                $t = Read-MetadadosTxt -Path $txt.FullName
                $body = $t.Body; $c = $t.Campos
                if (-not $data -and $c['data'])       { $data = $c['data'] }
                if (-not $periodo -and $c['periodo']) { $periodo = $c['periodo'] }
                if (-not $hi -and $c['hora_inicio'])  { $hi = $c['hora_inicio'] }
                if (-not $hf -and $c['hora_fim'])     { $hf = $c['hora_fim'] }
                if ($c['cidade'])         { $cidade = $c['cidade'] }
                if ($c['atividade'])      { $atividade = $c['atividade'] }
                if ($c['produtor_pasta']) { $pasta = $c['produtor_pasta'] }
            }
            if (-not $body -and $json) { $body = ConvertFrom-RelatorioJson -Json $json.FullName }   # inbox sem .txt: corpo montado do JSON
            if (-not $data -and $leaf -match '^(\d{4}-\d{2}-\d{2})') { $data = $Matches[1] }
            if (-not $data -and $leaf -match '^(\d{2})-(\d{4}-\d{2})') { $data = "$($Matches[2])-$($Matches[1])" }   # formato novo DD-AAAA-MM-
            if (-not $data) { continue }

            # par pode ser so .txt ou so .json: $null.FullName quebraria no StrictMode
            $visitas.Add((New-Visita -Data $data -Periodo $periodo -HoraInicio $hi -HoraFim $hf `
                -Cidade $cidade -Atividade $atividade -ProdutorPasta $pasta `
                -SourceTxt $(if ($txt) { $txt.FullName } else { $null }) `
                -SourceJson $(if ($json) { $json.FullName } else { $null }) -TxtBody $body `
                -OrigemInbox $true -Ano $p.Ano -MM $p.MM -LeafName $leaf -PastaInbox $p.PastaInbox))
        }
    }
    return , @($visitas)
}

# Estrutura de dias (visitas.json / colado) -> lista plana de visitas
function ConvertFrom-DiasParsed {
    param($Parsed)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($dia in @($Parsed)) {
        $data = [string](Get-PropValue $dia 'data')
        foreach ($v in @(Get-PropValue $dia 'visitas')) {
            $list.Add((New-Visita -Data $data `
                -Periodo ([string](Get-PropValue $v 'periodo')) `
                -HoraInicio ([string](Get-PropValue $v 'hora_inicio')) `
                -HoraFim ([string](Get-PropValue $v 'hora_fim')) `
                -Cidade ([string](Get-PropValue $v 'cidade')) `
                -Atividade ([string](Get-PropValue $v 'atividade')) `
                -ProdutorPasta ([string](Get-PropValue $v 'produtor_pasta')) `
                -SourceTxt $null -SourceJson $null -TxtBody $null `
                -OrigemInbox $false -Ano $null -MM $null -LeafName $null))
        }
    }
    return , @($list)
}

# Agrupa visitas por dia, atribui tags A/B/C e normaliza o nome da pasta
function ConvertTo-RelatorioDias {
    param($Visitas)
    $dias = [System.Collections.Generic.List[object]]::new()
    $grupos = @($Visitas | Group-Object -Property data | Sort-Object Name -Descending)
    foreach ($g in $grupos) {
        $i = 0
        foreach ($v in $g.Group) {
            $v.tag = [string][char](65 + $i)
            $v.produtor_pasta = Format-NomePastaProdutor -Nome $v.produtor_pasta -Data $v.data
            $i++
        }
        $dias.Add([PSCustomObject]@{ data = $g.Name; visitas = @($g.Group) })
    }
    return , @($dias)
}

# Le JSON colado no terminal (FIM encerra) ou de arquivo
function Read-JsonManual {
    param([switch]$Colar, [string]$Arquivo)
    if ((-not $Colar) -and (Test-Path $Arquivo)) {
        Write-Host " [OK] Lendo arquivo: $Arquivo" -ForegroundColor Green
        return (Get-Content -Path $Arquivo -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    Write-Host " Cole o JSON do relatorio abaixo. Digite FIM numa nova linha para encerrar:" -ForegroundColor Yellow
    $linhas = [System.Collections.Generic.List[string]]::new()
    while ($true) {
        $linha = Read-Host
        if ($linha.Trim().ToUpper() -eq "FIM") { break }
        $linhas.Add($linha)
    }
    return (($linhas -join "`n") | ConvertFrom-Json)
}

# ================= METADADOS DE FOTO (DATA/HORA + GPS EXIF) =================
function Get-FotoMetadados {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.IO.FileInfo]$Arquivo)

    $DataHora = $null; $Lat = $null; $Lon = $null

    if ($Arquivo.BaseName -match '(\d{4})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})') {
        try {
            $DataHora = [datetime]::ParseExact(
                "$($Matches[1])-$($Matches[2])-$($Matches[3]) $($Matches[4]):$($Matches[5]):$($Matches[6])",
                "yyyy-MM-dd HH:mm:ss", $null)
        } catch { }
    }

    if ($Arquivo.Extension -match '\.jpe?g$') {
        $img = $null
        try {
            $img = [System.Drawing.Image]::FromFile($Arquivo.FullName)
            if (-not $DataHora) {
                $propDt = $img.PropertyItems | Where-Object { $_.Id -eq 36867 } | Select-Object -First 1
                if ($propDt) {
                    $dataStr = [System.Text.Encoding]::ASCII.GetString($propDt.Value).Trim("`0", " ")
                    $DataHora = [datetime]::ParseExact($dataStr, "yyyy:MM:dd HH:mm:ss", $null)
                }
            }
            $ConverterGps = {
                param([byte[]]$Bytes, [string]$Ref)
                if ($Bytes.Length -lt 24) { return $null }
                $gNum = [BitConverter]::ToUInt32($Bytes, 0);  $gDen = [BitConverter]::ToUInt32($Bytes, 4)
                $mNum = [BitConverter]::ToUInt32($Bytes, 8);  $mDen = [BitConverter]::ToUInt32($Bytes, 12)
                $sNum = [BitConverter]::ToUInt32($Bytes, 16); $sDen = [BitConverter]::ToUInt32($Bytes, 20)
                if ($gDen -eq 0 -or $mDen -eq 0 -or $sDen -eq 0) { return $null }
                $dec = ($gNum / $gDen) + (($mNum / $mDen) / 60.0) + (($sNum / $sDen) / 3600.0)
                if ($Ref -in @("S", "W")) { $dec = -$dec }
                return [Math]::Round($dec, 6)
            }
            $propLatRef = $img.PropertyItems | Where-Object { $_.Id -eq 1 } | Select-Object -First 1
            $propLat    = $img.PropertyItems | Where-Object { $_.Id -eq 2 } | Select-Object -First 1
            $propLonRef = $img.PropertyItems | Where-Object { $_.Id -eq 3 } | Select-Object -First 1
            $propLon    = $img.PropertyItems | Where-Object { $_.Id -eq 4 } | Select-Object -First 1
            if ($propLatRef -and $propLat -and $propLonRef -and $propLon) {
                $latRefStr = [System.Text.Encoding]::ASCII.GetString($propLatRef.Value).Trim("`0", " ")
                $lonRefStr = [System.Text.Encoding]::ASCII.GetString($propLonRef.Value).Trim("`0", " ")
                $Lat = & $ConverterGps $propLat.Value $latRefStr
                $Lon = & $ConverterGps $propLon.Value $lonRefStr
            }
        } catch { } finally { if ($img) { $img.Dispose() } }
    }

    if (-not $DataHora) { $DataHora = $Arquivo.LastWriteTime }

    return [PSCustomObject]@{
        Arquivo = $Arquivo; DataHora = $DataHora
        DataStr = $DataHora.ToString("yyyy-MM-dd"); HoraStr = $DataHora.ToString("HH:mm:ss")
        Lat = $Lat; Lon = $Lon; TemGps = ($null -ne $Lat -and $null -ne $Lon)
    }
}

# Formata hora_inicio/hora_fim sem nunca exibir "14:46-null" (vazio ou string "null").
function Format-HorarioVisita {
    param($HoraInicio, $HoraFim)
    $ini = ([string]$HoraInicio).Trim(); $fim = ([string]$HoraFim).Trim()
    if ($ini -eq 'null') { $ini = '' }
    if ($fim -eq 'null') { $fim = '' }
    if ($ini -and $fim) { return "$ini-$fim" }
    if ($ini) { return $ini }
    return $fim
}

# Item de foto no mesmo formato de Get-FotoMetadados + LoteVisita (pasta de visita ja agrupada).
function New-FotoItem {
    param([System.IO.FileInfo]$Arquivo, [datetime]$DataHora, $Lat, $Lon, [string]$LoteVisita)
    return [PSCustomObject]@{
        Arquivo = $Arquivo; DataHora = $DataHora
        DataStr = $DataHora.ToString('yyyy-MM-dd'); HoraStr = $DataHora.ToString('HH:mm:ss')
        Lat = $Lat; Lon = $Lon; TemGps = ($null -ne $Lat -and $null -ne $Lon)
        LoteVisita = $LoteVisita
    }
}

# Rotulo CURTO da origem da foto, no padrao do organizar-camera (AAAA\MM\DD_visita_N),
# sem o prefixo C:\Users\... — so o ano/mes/pasta p/ situar de onde vem o lote.
function Get-OrigemCurtaFoto {
    param($Foto)
    if ($Foto.LoteVisita) { return [string]$Foto.LoteVisita }
    $full = $Foto.Arquivo.FullName
    if ($full -match '(\d{4}\\\d{2})\\[^\\]+$') { return "$($Matches[1]) (avulsa)" }
    if ($full -match '\\Camera\\[^\\]+$') { return 'Camera (bruta)' }
    return (Split-Path (Split-Path $full -Parent) -Leaf)
}

# ================= BUSCA CIRURGICA POR DATA (zero EXIF nas visitas organizadas) =================
function Get-FotosCirurgico {
    # Em vez de varrer o acervo inteiro: acessa so <AAAA>\<MM> dos meses alvo,
    # le nome/hora/lat/lon direto do _visita.json (ZERO System.Drawing.Image),
    # e pre-filtra avulsas/brutas pelo padrao *AAAAMMDD* no nome antes do EXIF.
    param([string[]]$Raizes, [string[]]$DatasAlvo, [string[]]$Extensoes)
    $fotos  = [System.Collections.Generic.List[object]]::new()
    $vistos = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $meses   = @($DatasAlvo | ForEach-Object { $_.Substring(0, 7) } | Sort-Object -Unique)
    $padroes = @($DatasAlvo | ForEach-Object { '*' + ($_ -replace '-', '') + '*' })

    foreach ($raiz in $Raizes) {
        $ehOrganizada = @(Get-ChildItem -LiteralPath $raiz -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d{4}$' }).Count -gt 0
        if ($ehOrganizada) {
            foreach ($mes in $meses) {
                $pastaMes = Join-Path $raiz ($mes -replace '-', '\')
                if (-not (Test-Path -LiteralPath $pastaMes)) { continue }
                $diasDoMes = @($DatasAlvo | Where-Object { $_.Substring(0, 7) -eq $mes })
                # Pastas <DD>_visita* dos dias alvo: lote pre-agrupado via _visita.json
                foreach ($dir in @(Get-ChildItem -LiteralPath $pastaMes -Directory | Where-Object { $_.Name -match '^\d{2}_visita' })) {
                    $dd = $dir.Name.Substring(0, 2)
                    $dataDir = @($diasDoMes | Where-Object { $_.Substring(8, 2) -eq $dd })
                    if ($dataDir.Count -eq 0) { continue }
                    $loteKey = $dir.FullName.Substring($raiz.Length + 1)
                    $vj = $null
                    $vjPath = Join-Path $dir.FullName '_visita.json'
                    if (Test-Path -LiteralPath $vjPath) {
                        try { $vj = Get-Content -LiteralPath $vjPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
                    }
                    $arqsVisita = Get-PropValue $vj 'arquivos'
                    if ($vj -and $arqsVisita) {
                        $dataVisita = [string](Get-PropValue $vj 'data')
                        foreach ($a in @($arqsVisita)) {
                            $nome = [string](Get-PropValue $a 'nome')
                            if (-not $nome) { continue }
                            $arqPath = Join-Path $dir.FullName $nome
                            if (-not (Test-Path -LiteralPath $arqPath)) { continue }   # movida p/ E: -> outra raiz resolve
                            if (-not $vistos.Add($nome)) { continue }
                            $fi = Get-Item -LiteralPath $arqPath
                            $dt = $null
                            try { $dt = [datetime]::ParseExact("$dataVisita $([string](Get-PropValue $a 'hora'))", 'yyyy-MM-dd HH:mm:ss', $null) } catch { }
                            if (-not $dt) { $dt = $fi.LastWriteTime }
                            $fotos.Add((New-FotoItem -Arquivo $fi -DataHora $dt -Lat (Get-PropValue $a 'lat') -Lon (Get-PropValue $a 'lon') -LoteVisita $loteKey))
                        }
                    } else {
                        # Sem _visita.json legivel: fallback EXIF so nos arquivos do dia (pre-filtro no nome)
                        foreach ($f in @(Get-ChildItem -LiteralPath $dir.FullName -File | Where-Object {
                            $Extensoes -contains $_.Extension.ToLower() -and $(foreach ($p in $padroes) { if ($_.Name -like $p) { $true; break } })
                        })) {
                            if ($vistos.Add($f.Name)) {
                                $m = Get-FotoMetadados -Arquivo $f
                                Add-Member -InputObject $m -NotePropertyName LoteVisita -NotePropertyValue $loteKey -Force
                                $fotos.Add($m)
                            }
                        }
                    }
                }
                # Avulsas soltas na raiz do mes: pre-filtro *AAAAMMDD* antes de ler EXIF
                foreach ($f in @(Get-ChildItem -LiteralPath $pastaMes -File | Where-Object {
                    $Extensoes -contains $_.Extension.ToLower() -and $_.Name -notlike '.*' -and $_.Name -notlike '~syncthing~*' -and
                    $(foreach ($p in $padroes) { if ($_.Name -like $p) { $true; break } })
                })) {
                    if ($vistos.Add($f.Name)) {
                        $m = Get-FotoMetadados -Arquivo $f
                        Add-Member -InputObject $m -NotePropertyName LoteVisita -NotePropertyValue $null -Force
                        $fotos.Add($m)
                    }
                }
            }
        } else {
            # Raiz bruta (Camera\): pre-filtro pelo nome *AAAAMMDD* ou LastWriteTime do dia
            foreach ($f in @(Get-ChildItem -LiteralPath $raiz -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
                $Extensoes -contains $_.Extension.ToLower() -and $_.Name -notlike '.*' -and $_.Name -notlike '~syncthing~*'
            })) {
                $casa = $false
                foreach ($p in $padroes) { if ($f.Name -like $p) { $casa = $true; break } }
                if (-not $casa) { $casa = $DatasAlvo -contains $f.LastWriteTime.ToString('yyyy-MM-dd') }
                if (-not $casa) { continue }
                if ($vistos.Add($f.Name)) {
                    $m = Get-FotoMetadados -Arquivo $f
                    Add-Member -InputObject $m -NotePropertyName LoteVisita -NotePropertyValue $null -Force
                    $fotos.Add($m)
                }
            }
        }
    }
    return , @($fotos | Sort-Object DataHora)
}

# ================= GPS TXT (retorna objeto p/ manifesto runs\) =================
function Export-CoordenadasTxt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][array]$FotosLote,
        [Parameter(Mandatory = $true)][string]$PastaDestino,
        [Parameter(Mandatory = $true)][string]$NomeProdutor
    )
    $TxtPath = Join-Path -Path $PastaDestino -ChildPath "coordenadas_gps.txt"
    $FotosComGps = @($FotosLote | Where-Object { $_.TemGps })
    $MediaLat = $null; $MediaLon = $null

    $Linhas = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $TxtPath)) {
        $Linhas.Add("==================================================================")
        $Linhas.Add(" GEOLOCALIZACAO DAS FOTOS DE CAMPO - $NomeProdutor")
        $Linhas.Add(" Gerado em: $(Get-Date -Format 'dd/MM/yyyy HH:mm')")
        $Linhas.Add("==================================================================")
    }
    if ($FotosComGps.Count -gt 0) {
        $MediaLat = [Math]::Round(($FotosComGps | Measure-Object -Property Lat -Average).Average, 6)
        $MediaLon = [Math]::Round(($FotosComGps | Measure-Object -Property Lon -Average).Average, 6)
        $LatStr = $MediaLat.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $LonStr = $MediaLon.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $Linhas.Add("`n[PONTO MEDIO DA PROPRIEDADE ($($FotosComGps.Count) fotos com GPS)]")
        $Linhas.Add("Coordenadas (Decimal): $LatStr, $LonStr")
        $Linhas.Add("Google Maps          : https://www.google.com/maps?q=$LatStr,$LonStr")
        $Linhas.Add("------------------------------------------------------------------")
    } else {
        $Linhas.Add("`n[AVISO] Nenhuma foto deste lote continha metadados GPS EXIF.")
    }
    foreach ($f in $FotosLote) {
        if ($f.TemGps) {
            $fLat = $f.Lat.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            $fLon = $f.Lon.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            $Linhas.Add("$($f.DataStr) $($f.HoraStr) | $($f.Arquivo.Name) -> $fLat, $fLon | https://www.google.com/maps?q=$fLat,$fLon")
        } else {
            $Linhas.Add("$($f.DataStr) $($f.HoraStr) | $($f.Arquivo.Name) -> Sem GPS no EXIF")
        }
    }
    Add-Content -LiteralPath $TxtPath -Value $Linhas -Encoding UTF8
    return [PSCustomObject]@{ Count = $FotosComGps.Count; MediaLat = $MediaLat; MediaLon = $MediaLon; TxtPath = $TxtPath }
}

# ================= PRE-SUGESTAO POR HORARIO/PERIODO =================
function ConvertTo-Minutos {
    param([string]$Hora)
    if ($Hora -match '^(\d{1,2}):(\d{2})$') { return ([int]$Matches[1] * 60 + [int]$Matches[2]) }
    return $null
}

# Horario medio do lote (minutos desde 00:00 + HH:mm) e faixa inicial..final das fotos.
function Get-HorarioMedioLote {
    param([array]$Lote)
    $l = @($Lote | Sort-Object DataHora)
    if ($l.Count -eq 0) { return $null }
    $media = ($l | ForEach-Object { $_.DataHora.Hour * 60 + $_.DataHora.Minute } | Measure-Object -Average).Average
    $h = [int][Math]::Floor($media / 60); $m = [int][Math]::Round($media - $h * 60)
    if ($m -eq 60) { $h++; $m = 0 }
    return @{
        MediaMinutos = [double]$media
        MediaHora    = '{0:D2}:{1:D2}' -f $h, $m
        Faixa        = "$($l[0].DataHora.ToString('HH:mm')) as $($l[-1].DataHora.ToString('HH:mm'))"
    }
}

# Sugestao por proximidade: |media - hora_inicio| <= JanelaMin. Empate na menor distancia
# (2 visitas equidistantes) = ambiguo -> $null. Usado quando hora_fim e null/vazio (audio).
function Get-SugestaoTagProximidade {
    param([double]$MediaMinutos, $Visitas, [int]$JanelaMin = 90)
    $cand = [System.Collections.Generic.List[object]]::new()
    foreach ($v in $Visitas) {
        $ini = ConvertTo-Minutos ([string]$v.hora_inicio)
        if ($null -eq $ini) { continue }
        $diff = [Math]::Abs($MediaMinutos - $ini)
        if ($diff -le $JanelaMin) { $cand.Add([PSCustomObject]@{ Visita = $v; Diff = $diff }) }
    }
    if ($cand.Count -eq 0) { return $null }
    $menor = ($cand | Measure-Object -Property Diff -Minimum).Minimum
    $melhores = @($cand | Where-Object { $_.Diff -eq $menor })
    if ($melhores.Count -ne 1) { return $null }
    return $melhores[0].Visita.tag
}

# Retorna a TAG se um unico produtor do dia casa com o horario medio do lote; senao $null.
# inicio+fim: media dentro de [ini-30, fim+30]. So inicio (hora_fim null): visita unica mais
# proxima em <= 90 min. Periodo generico (manha/tarde) so como fallback sem hora_inicio.
function Get-SugestaoTag {
    param($Lote, $Visitas)
    $hm = Get-HorarioMedioLote -Lote $Lote
    if ($null -eq $hm) { return $null }
    $media = $hm.MediaMinutos
    $casaram = [System.Collections.Generic.List[object]]::new()
    $soInicio = [System.Collections.Generic.List[object]]::new()
    foreach ($v in $Visitas) {
        $ini = ConvertTo-Minutos ([string]$v.hora_inicio)
        $fim = ConvertTo-Minutos ([string]$v.hora_fim)
        if ($null -ne $ini -and $null -ne $fim) {
            if ($media -ge ($ini - 30) -and $media -le ($fim + 30)) { $casaram.Add($v) }
        } elseif ($null -ne $ini) {
            $soInicio.Add($v)
        } elseif ($v.periodo) {
            if ($v.periodo -eq 'manha' -and $media -ge 360 -and $media -le 720) { $casaram.Add($v) }
            elseif ($v.periodo -eq 'tarde' -and $media -ge 721 -and $media -le 1110) { $casaram.Add($v) }
        }
    }
    if ($casaram.Count -eq 1) { return $casaram[0].tag }
    if ($casaram.Count -gt 1) { return $null }
    if ($soInicio.Count -gt 0) { return Get-SugestaoTagProximidade -MediaMinutos $media -Visitas $soInicio }
    return $null
}

# Enter (vazio) ou S confirmam a sugestao ativa; sem sugestao, S segue como tag invalida.
function Resolve-TagSugestao {
    param([string]$Escolha, $Sugestao)
    if (($Escolha -eq '' -or $Escolha -eq 'S') -and $Sugestao) { return $Sugestao }
    return $Escolha
}

# ================= SELETOR DE PASTAS/LOTES DO DIA (setas + Espaco + Enter) =================
# Lista os lotes do dia (pastas DD_visita_N do organizar-camera + avulsas agrupadas) p/ escolha
# explicita da origem: setas navegam, Espaco ou numero [1-9] marca, A todos/nenhum, O abre a
# pasta no Explorer, Enter confirma (retorna indices marcados), Q/Esc cancela (retorna $null).
# Cada linha: [n] checkbox + pasta + qtd fotos + faixa + media HH:mm + GPS + sugestao do relatorio.
function Show-SeletorLotesDia {
    param([array]$Lotes, [string]$Data, [array]$Visitas = @())
    $Itens = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Lotes.Count; $i++) {
        $l = @($Lotes[$i])
        $origens = @($l | ForEach-Object { Get-OrigemCurtaFoto -Foto $_ } | Sort-Object -Unique)
        $hm = Get-HorarioMedioLote -Lote $l
        $sugTag = $null; $sugTxt = ''
        if ($Visitas.Count -gt 0) {
            $sugTag = Get-SugestaoTag -Lote $l -Visitas $Visitas
            if ($sugTag) {
                $vS = $Visitas | Where-Object { $_.tag -eq $sugTag } | Select-Object -First 1
                $hr = Format-HorarioVisita $vS.hora_inicio $vS.hora_fim
                if (-not $hr) { $hr = [string]$vS.periodo }
                $nomeS = [string]$vS.produtor_pasta
                if ($nomeS.Length -gt 34) { $nomeS = $nomeS.Substring(0, 34) + '...' }
                $sugTxt = " <- SUGESTAO [$sugTag] $nomeS (relatorio $hr)"
            }
        }
        $Itens.Add([PSCustomObject]@{
            Indice = $i; Selecionado = $false
            Rotulo = ($origens -join ' | ')
            Qtd = $l.Count
            Horario = $hm.Faixa
            MediaHora = $hm.MediaHora
            Gps = @($l | Where-Object { $_.TemGps }).Count
            Pasta = (Split-Path $l[0].Arquivo.FullName -Parent)
            Sugestao = $sugTxt
        })
    }
    if ($Itens.Count -eq 0) { return $null }
    $Atual = 0; $Offset = 0; $Altura = 15
    try { $Altura = [Math]::Max(6, $Host.UI.RawUI.WindowSize.Height - 12) } catch { }
    try { [Console]::CursorVisible = $false } catch { }
    try {
        while ($true) {
            if ($Atual -gt ($Itens.Count - 1)) { $Atual = $Itens.Count - 1 }
            if ($Atual -lt 0) { $Atual = 0 }
            Clear-Host
            $Marcados = @($Itens | Where-Object { $_.Selecionado }).Count
            Write-Host ("=" * 78) -ForegroundColor Cyan
            Write-Host " PASTAS DE FOTOS DO DIA $Data - escolha a origem correta    Marcadas: $Marcados / $($Itens.Count)" -ForegroundColor White
            Write-Host ("=" * 78) -ForegroundColor Cyan
            Write-Host " [Setas] Navegar | [Espaco/1-9] Marcar | [A] Todos/Nenhum | [O] Abrir pasta | [Enter] Confirmar | [Q/Esc] Cancelar" -ForegroundColor Yellow
            Write-Host ("-" * 78) -ForegroundColor Cyan
            if ($Atual -lt $Offset) { $Offset = $Atual }
            if ($Atual -ge ($Offset + $Altura)) { $Offset = $Atual - $Altura + 1 }
            $Fim = [Math]::Min($Itens.Count - 1, $Offset + $Altura - 1)
            for ($i = $Offset; $i -le $Fim; $i++) {
                $It = $Itens[$i]
                $Check = if ($It.Selecionado) { "[X]" } else { "[ ]" }
                $Seta  = if ($i -eq $Atual) { " > " } else { "   " }
                $Num = if ($i -lt 9) { "[$($i + 1)] " } else { "    " }
                $Linha = "$Seta$Num$Check $($It.Rotulo)  ($($It.Qtd) fotos, $($It.Horario), media $($It.MediaHora), GPS: $($It.Gps))$($It.Sugestao)"
                if ($i -eq $Atual) { Write-Host $Linha -ForegroundColor Black -BackgroundColor Gray }
                elseif ($It.Selecionado) { Write-Host $Linha -ForegroundColor Green }
                elseif ($It.Sugestao) { Write-Host $Linha -ForegroundColor Yellow }
                else { Write-Host $Linha -ForegroundColor White }
            }
            Write-Host ("-" * 78) -ForegroundColor Cyan
            $Tecla = [Console]::ReadKey($true)
            switch ($Tecla.Key) {
                "UpArrow"   { if ($Atual -gt 0) { $Atual-- } }
                "DownArrow" { if ($Atual -lt ($Itens.Count - 1)) { $Atual++ } }
                "Spacebar"  { $Itens[$Atual].Selecionado = -not $Itens[$Atual].Selecionado }
                { $_ -match '^(D|NumPad)[1-9]$' } {
                    $n = [int]::Parse($_.Substring($_.Length - 1)) - 1
                    if ($n -lt $Itens.Count) { $Itens[$n].Selecionado = -not $Itens[$n].Selecionado; $Atual = $n }
                }
                "A" { $novo = ($Marcados -lt $Itens.Count); foreach ($It in $Itens) { $It.Selecionado = $novo } }
                "O" { try { Invoke-Item -LiteralPath $Itens[$Atual].Pasta } catch { } }
                "Enter" { return , @($Itens | Where-Object { $_.Selecionado } | ForEach-Object { $_.Indice }) }
                { $_ -in "Q", "Escape" } { return $null }
            }
        }
    } finally {
        try { [Console]::CursorVisible = $true } catch { }
        Clear-KeyBuffer
    }
}

# Copia um lote inteiro p/ a pasta do produtor [OD] + GPS + manifesto + status (scriptblock marcar).
# Reutilizado pelo fluxo lote-a-lote e pelo seletor de pastas do dia [E].
function Copy-LoteParaVisita {
    param($Lote, $Visita, [string]$Modo, $RunFotos, $RunGps, $Marcar)
    $copiadas = 0
    foreach ($FotoItem in $Lote) {
        try {
            Copy-Item -LiteralPath $FotoItem.Arquivo.FullName -Destination $Visita.CaminhoDestino -Force -ErrorAction Stop
            $copiadas++
            $RunFotos.Add([PSCustomObject]@{ origem = $FotoItem.Arquivo.FullName; destino = (Join-Path $Visita.CaminhoDestino $FotoItem.Arquivo.Name); produtor = $Visita.produtor_pasta })
            Write-TriagemLog "COPIADA [Syncthing]: $($FotoItem.Arquivo.FullName) -> [OD]: $(Join-Path $Visita.CaminhoDestino $FotoItem.Arquivo.Name)"
        } catch { Write-Host "   [!] Falha ao copiar $($FotoItem.Arquivo.Name): $($_.Exception.Message)" -ForegroundColor Red }
    }
    $gps = Export-CoordenadasTxt -FotosLote $Lote -PastaDestino $Visita.CaminhoDestino -NomeProdutor $Visita.produtor_pasta
    $RunGps.Add([PSCustomObject]@{ produtor = $Visita.produtor_pasta; count = $gps.Count; lat = $gps.MediaLat; lon = $gps.MediaLon; txt = $gps.TxtPath })
    Write-Host "   [OK] $($Lote.Count) foto(s) + GPS ($($gps.Count)) copiadas para [OD]" -ForegroundColor Green
    Write-Host "        Destino [OD]: $($Visita.CaminhoDestino)" -ForegroundColor DarkGray
    & $Marcar $Visita.tag $Modo $copiadas (Split-Path $Lote[0].Arquivo.FullName -Parent)
}

# ================= EDITAR VISITA (data / periodo / hora / cidade / atividade / nome) =================
function Edit-VisitaItem {
    param($Item)
    $Dia = $Item.Dia; $V = $Item.Visita

    # 1. Data AAAA-MM-DD (usada exclusivamente para buscar as fotos do dia)
    Write-Host "`n  Data atual (busca de fotos): $($Dia.data)" -ForegroundColor DarkGray
    $R = Read-Host "  Nova data AAAA-MM-DD (Enter mantem)"
    if (-not [string]::IsNullOrWhiteSpace($R)) {
        if ($R.Trim() -match '^\d{4}-\d{2}-\d{2}$') { $Dia.data = $R.Trim(); $V.data = $R.Trim() }
        else { Write-Host "  [!] Data invalida (AAAA-MM-DD). Mantida." -ForegroundColor Red; Start-Sleep -Milliseconds 1200 }
    }

    # 2. Periodo
    Write-Host "`n  Periodo atual: $($V.periodo)" -ForegroundColor DarkGray
    Write-Host "   [1] manha (06:00-12:00)  [2] tarde (12:01-18:30)  [0] limpar" -ForegroundColor White
    $R = Read-Host "  Periodo (Enter mantem)"
    if ($R -eq '1') { $V.periodo = 'manha' } elseif ($R -eq '2') { $V.periodo = 'tarde' } elseif ($R -eq '0') { $V.periodo = $null }

    # 3. Horario inicio/fim (HH:mm; Enter mantem; '-' limpa)
    Write-Host "`n  Hora inicio atual: $($V.hora_inicio) | Hora fim atual: $($V.hora_fim)" -ForegroundColor DarkGray
    $R = Read-Host "  Hora inicio HH:mm (Enter mantem, '-' limpa)"
    if ($R -eq '-') { $V.hora_inicio = $null }
    elseif (-not [string]::IsNullOrWhiteSpace($R)) {
        if ($R.Trim() -match '^\d{1,2}:\d{2}$') { $V.hora_inicio = $R.Trim() }
        else { Write-Host "  [!] Hora invalida. Mantida." -ForegroundColor Red; Start-Sleep -Milliseconds 1000 }
    }
    $R = Read-Host "  Hora fim HH:mm (Enter mantem, '-' limpa)"
    if ($R -eq '-') { $V.hora_fim = $null }
    elseif (-not [string]::IsNullOrWhiteSpace($R)) {
        if ($R.Trim() -match '^\d{1,2}:\d{2}$') { $V.hora_fim = $R.Trim() }
        else { Write-Host "  [!] Hora invalida. Mantida." -ForegroundColor Red; Start-Sleep -Milliseconds 1000 }
    }

    # 4. Cidade
    Write-Host "`n  Cidades:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Cidades.Count; $i++) { Write-Host "   [$($i+1)] $($Cidades[$i])" -ForegroundColor White }
    Write-Host "   Atual: $($V.cidade)" -ForegroundColor DarkGray
    $R = Read-Host "  Numero da cidade (Enter mantem)"
    if ($R -match '^\d+$' -and [int]$R -ge 1 -and [int]$R -le $Cidades.Count) { $V.cidade = $Cidades[[int]$R - 1] }

    # 5. Atividade
    Write-Host "`n  Atividades:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Atividades.Count; $i++) { Write-Host "   [$($i+1)] $($Atividades[$i])" -ForegroundColor White }
    Write-Host "   Atual: $($V.atividade)" -ForegroundColor DarkGray
    $R = Read-Host "  Numero da atividade (Enter mantem)"
    if ($R -match '^\d+$' -and [int]$R -ge 1 -and [int]$R -le $Atividades.Count) { $V.atividade = $Atividades[[int]$R - 1] }

    # 6. Nome da pasta (prefixo AAAA-MM-DD- + sufixo - RESUMO)
    Write-Host "`n  Nome atual da pasta: $($V.produtor_pasta)" -ForegroundColor DarkGray
    $R = Read-Host "  Novo nome (Enter mantem)"
    if (-not [string]::IsNullOrWhiteSpace($R)) {
        $Novo = $R.Trim()
        if (Test-CampoSeguro $Novo) { $V.produtor_pasta = Format-NomePastaProdutor -Nome $Novo -Data $Dia.data }
        else { Write-Host "  [!] Nome invalido (caracteres de caminho). Mantido." -ForegroundColor Red; Start-Sleep -Milliseconds 1200 }
    }
}

# ================= REVISAO INTERATIVA (TUI com setas) =================
function Show-RevisaoVisitas {
    param([array]$RelatorioDias)

    $Itens = [System.Collections.Generic.List[object]]::new()
    foreach ($Dia in $RelatorioDias) {
        foreach ($V in $Dia.visitas) { $Itens.Add([PSCustomObject]@{ Dia = $Dia; Visita = $V }) }
    }
    if ($Itens.Count -eq 0) { Write-Host "`n[AVISO] Nenhuma visita para revisar." -ForegroundColor Yellow; return $false }

    $Atual = 0; $Offset = 0; $Altura = 15
    try { $Altura = [Math]::Max(6, $Host.UI.RawUI.WindowSize.Height - 12) } catch { }
    try { [Console]::CursorVisible = $false } catch { }

    try {
        while ($true) {
            if ($Itens.Count -eq 0) { Write-Host "`n[AVISO] Lista vazia. Nada a criar." -ForegroundColor Yellow; return $false }
            if ($Atual -gt ($Itens.Count - 1)) { $Atual = $Itens.Count - 1 }
            if ($Atual -lt 0) { $Atual = 0 }

            Clear-Host
            Write-Host ("=" * 78) -ForegroundColor Cyan
            Write-Host " REVISAO DAS VISITAS - antes de criar pastas / gravar relatorios    Total: $($Itens.Count)" -ForegroundColor White
            Write-Host ("=" * 78) -ForegroundColor Cyan
            Write-Host " [Setas] Navegar | [E] Editar | [D/Del] Remover | [C/Enter] Confirmar | [Q/Esc] Cancelar" -ForegroundColor Yellow
            Write-Host ("-" * 78) -ForegroundColor Cyan

            if ($Atual -lt $Offset) { $Offset = $Atual }
            if ($Atual -ge ($Offset + $Altura)) { $Offset = $Atual - $Altura + 1 }
            $Fim = [Math]::Min($Itens.Count - 1, $Offset + $Altura - 1)

            for ($i = $Offset; $i -le $Fim; $i++) {
                $It = $Itens[$i]; $V = $It.Visita
                $per = [string]$V.periodo
                $hor = Format-HorarioVisita $V.hora_inicio $V.hora_fim
                if ($hor) { $per = ("$per $hor").Trim() }
                $extra = if ($per) { " ($per)" } else { "" }
                $Linha = " [$($It.Dia.data)]$extra $($V.cidade) > $($V.atividade) > $($V.produtor_pasta)"
                if ($i -eq $Atual) { Write-Host $Linha -ForegroundColor Black -BackgroundColor Gray }
                else { Write-Host $Linha -ForegroundColor White }
            }

            Write-Host ("-" * 78) -ForegroundColor Cyan
            $Tecla = [Console]::ReadKey($true)

            switch ($Tecla.Key) {
                "UpArrow"   { if ($Atual -gt 0) { $Atual-- } }
                "DownArrow" { if ($Atual -lt ($Itens.Count - 1)) { $Atual++ } }
                "E" {
                    try { [Console]::CursorVisible = $true } catch { }
                    Clear-Host; Write-Host "=== EDITAR VISITA ===" -ForegroundColor Cyan
                    Edit-VisitaItem -Item $Itens[$Atual]
                    Clear-KeyBuffer
                    try { [Console]::CursorVisible = $false } catch { }
                }
                { $_ -in "D", "Delete" } {
                    $Rem = $Itens[$Atual]; $Itens.RemoveAt($Atual)
                    Write-TriagemLog "REMOVIDO na revisao: $($Rem.Dia.data) / $($Rem.Visita.produtor_pasta)"
                    if ($Atual -ge $Itens.Count -and $Atual -gt 0) { $Atual-- }
                }
                { $_ -in "C", "Enter" } {
                    try { [Console]::CursorVisible = $true } catch { }
                    Write-Host ""
                    $Resp = Read-Host "Deseja criar as pastas, gravar os relatorios no OneDrive e registrar em _processados.json? (S/N)"
                    Clear-KeyBuffer
                    if ($Resp -in @('s', 'S', 'y', 'Y')) {
                        foreach ($Dia in $RelatorioDias) {
                            $Dia.visitas = @($Itens | Where-Object { [object]::ReferenceEquals($_.Dia, $Dia) } | ForEach-Object { $_.Visita })
                        }
                        return $true
                    }
                    try { [Console]::CursorVisible = $false } catch { }
                }
                { $_ -in "Q", "Escape" } { return $false }
            }
        }
    } finally {
        try { [Console]::CursorVisible = $true } catch { }
        Clear-KeyBuffer
    }
}

# ================= REGISTRO UNICO _processados.json (append, por ano) =================
function Add-ProcessadoRegistro {
    param($Visita, [string]$Destino, [string[]]$Arquivos)
    $ano = $Visita.Ano
    if (-not $ano -and $Visita.data -match '^(\d{4})') { $ano = $Matches[1] }
    if (-not $ano) { throw 'Visita sem ano (inbox/data) para registrar.' }
    $inboxAno = Join-Path $InboxBase "$ano\_inbox_relatorio_ater"
    if (-not (Test-Path -LiteralPath $inboxAno)) { New-Item -ItemType Directory -Force -Path $inboxAno | Out-Null }
    $regPath = Join-Path $inboxAno '_processados.json'
    $lista = @()
    if (Test-Path -LiteralPath $regPath) {
        try { $lista = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { $lista = @() }
    }
    $lista += [PSCustomObject]@{
        processado_em    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        pasta_inbox      = $Visita.PastaInbox
        data             = $Visita.data
        cidade           = $Visita.cidade
        atividade        = $Visita.atividade
        produtor_pasta   = $Visita.produtor_pasta
        destino_onedrive = $Destino
        arquivos_levados = @($Arquivos)
        fotos_modo       = 'pulado'    # atualizado na Etapa 2: automatico | individual | manual | pulado
        fotos_copiadas   = 0
        fotos_origem_sugerida = $null
        sincronizada     = $false      # start-sync.ps1 marca true apos sync real (chave: destino_onedrive)
        sincronizada_em  = $null
        drive_url        = $null
    }
    # Gravacao atomica: .tmp + Move-Item -Force (nunca deixa JSON pela metade); append preserva os demais registros/campos
    $tmp = "$regPath.tmp"
    ConvertTo-Json -InputObject $lista -Depth 6 | Set-Content -LiteralPath $tmp -Encoding UTF8
    Move-Item -LiteralPath $tmp -Destination $regPath -Force
    return $regPath
}

# Resumo do historico do ano (usado no inicio da triagem; $null se ausente/ilegivel)
function Get-ResumoProcessadosAno {
    param([string]$Ano = (Get-Date -Format 'yyyy'))
    $regPath = Join-Path $InboxBase "$Ano\_inbox_relatorio_ater\_processados.json"
    if (-not (Test-Path -LiteralPath $regPath)) { return $null }
    try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
    $nSync = @($regs | Where-Object { Get-PropValue $_ 'sincronizada' }).Count
    return [PSCustomObject]@{
        ano = $Ano; triadas = $regs.Count; sincronizadas = $nSync; pendentes = ($regs.Count - $nSync)
        caminho = $regPath
    }
}

# Grava o resultado da triagem de fotos (Etapa 2) nos registros do dia, por destino_onedrive.
# $Status: hashtable tag -> { Modo, Copiadas, Origem }; visita sem status = 'pulado' (0 fotos).
# Append atomico (.tmp + Move-Item -Force); preserva todos os demais campos/registros.
function Update-FotosProcessados {
    param($Dia, $Status)
    if (-not $Dia.visitas -or @($Dia.visitas).Count -eq 0) { return }
    if ($Dia.data -notmatch '^(\d{4})') { return }
    $regPath = Join-Path $InboxBase "$($Matches[1])\_inbox_relatorio_ater\_processados.json"
    if (-not (Test-Path -LiteralPath $regPath)) { return }
    try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return }
    $mudou = $false
    foreach ($v in $Dia.visitas) {
        if (-not $v.CaminhoDestino) { continue }
        $st = $Status[$v.tag]
        $modo = if ($st) { $st.Modo } else { 'pulado' }
        $n    = if ($st) { [int]$st.Copiadas } else { 0 }
        $orig = if ($st) { $st.Origem } else { $null }
        foreach ($r in $regs) {
            $pd = $r.PSObject.Properties['destino_onedrive']
            if (-not $pd -or -not $pd.Value) { continue }
            if (([string]$pd.Value).TrimEnd('\') -ne ([string]$v.CaminhoDestino).TrimEnd('\')) { continue }
            Add-Member -InputObject $r -NotePropertyName fotos_modo -NotePropertyValue $modo -Force
            Add-Member -InputObject $r -NotePropertyName fotos_copiadas -NotePropertyValue $n -Force
            Add-Member -InputObject $r -NotePropertyName fotos_origem_sugerida -NotePropertyValue $orig -Force
            $mudou = $true
        }
    }
    if ($mudou) {
        $tmp = "$regPath.tmp"
        ConvertTo-Json -InputObject @($regs) -Depth 10 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $regPath -Force
    }
}

# ================= CONFERIR / REPROCESSAR VISITAS JA PROCESSADAS (-Conferir) =================
# Todos os registros de todos os anos, com registro_path (nao serializado) p/ salvar de volta.
function Get-ProcessadosTodos {
    $lista = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $InboxBase)) { return , @() }
    foreach ($ano in @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' } | Sort-Object Name)) {
        $p = Join-Path $ano.FullName '_inbox_relatorio_ater\_processados.json'
        if (-not (Test-Path -LiteralPath $p)) { continue }
        try { $regs = @(Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { continue }
        foreach ($r in $regs) {
            Add-Member -InputObject $r -NotePropertyName registro_path -NotePropertyValue $p -Force
            $lista.Add($r)
        }
    }
    return , @($lista)
}

# Auditoria do destino [OD]: TXT presente? quantas fotos vs registradas? (nunca altera nada)
function Test-VisitaDestino {
    param($Registro)
    $dest = [string]$Registro.destino_onedrive
    $existe = Test-Path -LiteralPath $dest
    $txts = @(Get-PropValue $Registro 'arquivos_levados' | Where-Object { $_ -like '*.txt' })
    $txtFaltando = @(); $fotosDest = 0
    if ($existe) {
        $fotosDest = @(Get-ChildItem -LiteralPath $dest -File | Where-Object { $ExtensoesFoto -contains $_.Extension.ToLower() }).Count
        foreach ($t in $txts) { if (-not (Test-Path -LiteralPath (Join-Path $dest $t))) { $txtFaltando += $t } }
    } else { $txtFaltando = $txts }
    $fotosReg = [int](Get-PropValue $Registro 'fotos_copiadas')
    return [PSCustomObject]@{
        PastaExiste = $existe; TxtFaltando = $txtFaltando; FotosNoDestino = $fotosDest
        FotosRegistradas = $fotosReg; FotosFaltando = [Math]::Max(0, $fotosReg - $fotosDest)
    }
}

# Recopia o .txt limpo (sem YAML) e anexos ATV 1 da inbox [GD] p/ o destino [OD]; inbox intacta.
function Repair-VisitaTxtAnexos {
    param($Registro)
    $dest = [string]$Registro.destino_onedrive
    $inbox = [string]$Registro.pasta_inbox
    $levados = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $inbox)) {
        Write-Host "  [!] Inbox [GD] ausente/offline: $inbox" -ForegroundColor Yellow
        return , @()
    }
    if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Force -Path $dest | Out-Null }
    $txt = Get-ChildItem -LiteralPath $inbox -File -Filter '*.txt' | Select-Object -First 1
    if ($txt) {
        $body = Remove-Secao10 (Remove-MarkdownBasico (Remove-YamlHeader (Get-Content -LiteralPath $txt.FullName -Raw -Encoding UTF8)))
        Set-Content -LiteralPath (Join-Path $dest $txt.Name) -Value $body -Encoding UTF8
        $levados.Add($txt.Name)
        Write-Host "  [OK] TXT limpo [GD] -> [OD]: $(Join-Path $dest $txt.Name)" -ForegroundColor Green
        Write-TriagemLog "REPROCESSO TXT [GD]: $($txt.FullName) -> [OD]: $(Join-Path $dest $txt.Name)"
    } else {
        # Sem .txt na inbox: monta do JSON (schema 1.0), grava o par com YAML na inbox e o limpo no [OD]
        $json = Get-ChildItem -LiteralPath $inbox -File -Filter '*.json' | Select-Object -First 1
        if ($json) {
            $null = Write-TxtInboxDoJson -JsonPath $json.FullName
            $body = ConvertFrom-RelatorioJson -Json $json.FullName
            if ($body) {
                $nomeTxt = [System.IO.Path]::GetFileNameWithoutExtension($json.Name) + '.txt'
                Set-Content -LiteralPath (Join-Path $dest $nomeTxt) -Value $body -Encoding UTF8
                $levados.Add($nomeTxt)
                Write-Host "  [OK] TXT montado do JSON [GD] -> [OD]: $(Join-Path $dest $nomeTxt)" -ForegroundColor Green
                Write-TriagemLog "REPROCESSO TXT DO JSON [GD]: $($json.FullName) -> [OD]: $(Join-Path $dest $nomeTxt)"
            } else { Write-Host "  [!] JSON sem relatorio legivel: $($json.FullName)" -ForegroundColor Yellow }
        } else { Write-Host "  [!] Nenhum .txt/.json na inbox [GD]: $inbox" -ForegroundColor Yellow }
    }
    if ([string]$Registro.atividade -ieq 'ATIVIDADE INDIVIDUAL 1') {
        foreach ($a in @(Get-ChildItem -LiteralPath $inbox -File | Where-Object {
            $_.Extension -in @('.png', '.jpg', '.jpeg') -and $_.BaseName -match 'mapa|c?alend[aá]rio'
        })) {
            Copy-Item -LiteralPath $a.FullName -Destination $dest -Force
            $levados.Add($a.Name)
            Write-Host "  [OK] Anexo ATV1 [GD] -> [OD]: $($a.Name)" -ForegroundColor Green
            Write-TriagemLog "REPROCESSO ANEXO [GD]: $($a.FullName) -> [OD]: $(Join-Path $dest $a.Name)"
        }
    }
    $uniao = @(@(Get-PropValue $Registro 'arquivos_levados') + @($levados) | Sort-Object -Unique)
    Add-Member -InputObject $Registro -NotePropertyName arquivos_levados -NotePropertyValue $uniao -Force
    return , @($levados)
}

# Fotos do dia da visita que AINDA nao estao no destino (busca cirurgica; respeita origem sugerida).
function Get-FotosFaltantesVisita {
    param($Registro)
    $dest = [string]$Registro.destino_onedrive
    $raizes = @($SyncthingFotos | Where-Object { Test-Path $_ })
    if ($raizes.Count -eq 0) { return , @() }
    $fotosDia = Get-FotosCirurgico -Raizes $raizes -DatasAlvo @([string]$Registro.data) -Extensoes $ExtensoesFoto
    $origem = [string](Get-PropValue $Registro 'fotos_origem_sugerida')
    if ($origem) {
        $origemNorm = $origem.TrimEnd('\')
        $fotosDia = @($fotosDia | Where-Object { (Split-Path $_.Arquivo.FullName -Parent) -ieq $origemNorm })
    }
    $faltantes = @($fotosDia | Where-Object { -not (Test-Path -LiteralPath (Join-Path $dest $_.Arquivo.Name)) })
    return , $faltantes
}

# Salva os registros de volta nos seus arquivos (atomico; remove registro_path antes de serializar).
function Save-ProcessadosPorArquivo {
    param($Registros)
    foreach ($g in @($Registros | Group-Object -Property registro_path)) {
        foreach ($r in $g.Group) { $r.PSObject.Properties.Remove('registro_path') }
        $tmp = "$($g.Name).tmp"
        ConvertTo-Json -InputObject @($g.Group) -Depth 10 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $g.Name -Force
        foreach ($r in $g.Group) { Add-Member -InputObject $r -NotePropertyName registro_path -NotePropertyValue $g.Name -Force }
    }
}

# -Conferir sem fotos_origem_sugerida: 2+ pastas/lotes distintos no dia -> NAO mistura nem copia
# tudo. Abre o seletor p/ o usuario marcar a pasta desta visita e grava a escolha em
# fotos_origem_sugerida no _processados.json. Retorna $true quando pode seguir com a copia.
function Resolve-OrigemFotosConferir {
    param($Registro)
    if ([string](Get-PropValue $Registro 'fotos_origem_sugerida')) { return $true }
    $raizes = @($SyncthingFotos | Where-Object { Test-Path $_ })
    if ($raizes.Count -eq 0) { return $true }
    $fotosDia = Get-FotosCirurgico -Raizes $raizes -DatasAlvo @([string]$Registro.data) -Extensoes $ExtensoesFoto
    $grupos = @($fotosDia | Group-Object { Split-Path $_.Arquivo.FullName -Parent })
    if ($grupos.Count -lt 2) { return $true }
    Write-Host "   [!] $($grupos.Count) pastas/lotes distintos em $($Registro.data) sem origem registrada — escolha a pasta desta visita." -ForegroundColor Yellow
    $lotes = @($grupos | ForEach-Object { , @($_.Group | Sort-Object DataHora) })
    $sel = Show-SeletorLotesDia -Lotes $lotes -Data ([string]$Registro.data)
    if ($null -eq $sel) { Write-Host "   -> Selecao cancelada. Fotos nao alteradas." -ForegroundColor DarkGray; return $false }
    $idx = @($sel)
    if ($idx.Count -ne 1) { Write-Host "   [!] Marque exatamente UMA pasta para esta visita. Nada copiado." -ForegroundColor Red; return $false }
    $pastaEscolhida = Split-Path $lotes[$idx[0]][0].Arquivo.FullName -Parent
    Add-Member -InputObject $Registro -NotePropertyName fotos_origem_sugerida -NotePropertyValue $pastaEscolhida -Force
    Save-ProcessadosPorArquivo -Registros @($Registro)
    Write-TriagemLog "CONFERIR origem escolhida: $($Registro.produtor_pasta) <- $pastaEscolhida"
    Write-Host "   [OK] Origem registrada em _processados.json: $pastaEscolhida" -ForegroundColor Green
    return $true
}

# Oferece e copia as fotos faltantes do dia (Camera_Celular_Organizada etc.) p/ o destino [OD].
# Retorna $true se copiou algo (registro atualizado em memoria; salvar com Save-ProcessadosPorArquivo).
function Copy-FotosFaltantesVisita {
    param($Registro)
    if (-not (Resolve-OrigemFotosConferir -Registro $Registro)) { return $false }
    $falt = Get-FotosFaltantesVisita -Registro $Registro
    if (@($falt).Count -eq 0) {
        Write-Host "   [OK] Nenhuma foto faltante encontrada para esta data/origem." -ForegroundColor Green
        return $false
    }
    Write-Host "   Fotos faltantes ($(@($falt).Count)):" -ForegroundColor Yellow
    $falt | Select-Object -First 20 | ForEach-Object { Write-Host "     $($_.Arquivo.Name)  [Syncthing]: $($_.Arquivo.DirectoryName)" -ForegroundColor DarkGray }
    $resp = (Read-Host "   Copiar $(@($falt).Count) foto(s) faltante(s) para [OD]? (S/N — N mantem como esta)").Trim()
    if ($resp -notmatch '^[Ss]') {
        Write-Host "   -> Fotos nao alteradas (ok manter menos fotos que o original)." -ForegroundColor DarkGray
        return $false
    }
    $copiadasConf = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $falt) {
        try {
            Copy-Item -LiteralPath $f.Arquivo.FullName -Destination $Registro.destino_onedrive -Force -ErrorAction Stop
            $copiadasConf.Add($f)
            Write-TriagemLog "REPROCESSO COPIADA [Syncthing]: $($f.Arquivo.FullName) -> [OD]: $(Join-Path $Registro.destino_onedrive $f.Arquivo.Name)"
        } catch { Write-Host "   [!] Falha ao copiar $($f.Arquivo.Name): $($_.Exception.Message)" -ForegroundColor Red }
    }
    if ($copiadasConf.Count -eq 0) { return $false }
    $null = Export-CoordenadasTxt -FotosLote $copiadasConf.ToArray() -PastaDestino $Registro.destino_onedrive -NomeProdutor $Registro.produtor_pasta
    $novoTotal = @(Get-ChildItem -LiteralPath $Registro.destino_onedrive -File | Where-Object { $ExtensoesFoto -contains $_.Extension.ToLower() }).Count
    Add-Member -InputObject $Registro -NotePropertyName fotos_copiadas -NotePropertyValue $novoTotal -Force
    if ((Get-PropValue $Registro 'fotos_modo') -eq 'pulado') { Add-Member -InputObject $Registro -NotePropertyName fotos_modo -NotePropertyValue 'automatico' -Force }
    Write-Host "   [OK] $($copiadasConf.Count) foto(s) copiadas para [OD]: $($Registro.destino_onedrive)" -ForegroundColor Green
    return $true
}

# Fluxo interativo do -Conferir sobre a VARREDURA FISICA do Drive (nao apenas o JSON):
# qualquer pasta pode ser reprocessada; pastas NOVA ganham triagem completa [R].
# Nunca apaga nem forca fotos a mais.
function Invoke-ConferirReprocessar {
    Write-Host "`n=== CONFERIR / REPROCESSAR VISITAS (VARREDURA FISICA [GD] -> [OD]) ===" -ForegroundColor Cyan
    $pastasConf = Get-InboxPastasFisicas
    if (@($pastasConf).Count -eq 0) {
        Write-Host "[i] Nenhuma pasta encontrada fisicamente na inbox ($InboxBase)." -ForegroundColor Yellow
        return
    }
    Write-Host " $(@($pastasConf).Count) pasta(s) encontrada(s) no Drive. Opcoes por visita:" -ForegroundColor Gray
    Write-Host "   [T] Recopiar TXT/anexos [F] fotos faltantes [A] ambos | [R] Triar/reprocessar pasta NOVA | [P] Pular | [Q] Sair" -ForegroundColor Gray
    $regsConf = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $pastasConf) {
        $corStatus = switch ($p.Status) { 'NOVA' { 'Green' } 'SINCRONIZADA' { 'DarkGreen' } default { 'Yellow' } }
        Write-Host "`n [$($p.Status)] $($p.MM) | $($p.Cidade) | $($p.Atividade) | $($p.Produtor)" -ForegroundColor $corStatus
        Write-Host "   Origem [GD]: $($p.PastaInbox)" -ForegroundColor DarkGray
        $r = $p.Registro
        if (-not $r) {
            # Pasta NOVA: triagem completa a partir do Drive (pasta padronizada DD-AAAA-MM-NOME - RESUMO
            # no OneDrive + .txt limpo sem YAML + anexos Mapa/Calendario na ATV 1 + registro)
            $op = (Read-Host "   [R] Triar/reprocessar esta visita [P] pular [Q] sair").Trim().ToUpper()
            if ($op -eq 'Q') { break }
            if ($op -ne 'R') { continue }
            $todasNovas = Get-InboxVisitas   # atribuicao primeiro: return , @() preserva o array como objeto unico
            $visitasNova = @($todasNovas | Where-Object { $_.PastaInbox.TrimEnd('\') -ieq $p.PastaInbox.TrimEnd('\') })
            if ($visitasNova.Count -eq 0) {
                Write-Host "   [!] Sem par .txt/.json com data legivel nesta pasta; nada a triar." -ForegroundColor Yellow
                continue
            }
            $null = Invoke-CriarPastasEArquivar -RelatorioDias (ConvertTo-RelatorioDias -Visitas $visitasNova)
            $r = @(Get-ProcessadosTodos | Where-Object { ([string]$_.pasta_inbox).TrimEnd('\') -ieq $p.PastaInbox.TrimEnd('\') }) | Select-Object -Last 1
            if (-not $r) { continue }
            $regsConf.Add($r)
            if (Copy-FotosFaltantesVisita -Registro $r) {
                Save-ProcessadosPorArquivo -Registros @($r)
                Write-Host "   [OK] Registro [GD] atualizado: $($r.registro_path)" -ForegroundColor Cyan
            }
            continue
        }
        $regsConf.Add($r)
        $aud = Test-VisitaDestino -Registro $r
        $temProblema = (-not $aud.PastaExiste) -or (@($aud.TxtFaltando).Count -gt 0) -or ($aud.FotosFaltando -gt 0)
        if ($temProblema) { Write-Host "   [!] Pendencia no destino [OD] (pasta/TXT/fotos)." -ForegroundColor Yellow }
        Write-Host "   Destino [OD]: $($r.destino_onedrive)" -ForegroundColor DarkGray
        # TXT: caminho completo no destino, ou o que falta + origem na inbox
        $txtsReg = @(Get-PropValue $r 'arquivos_levados' | Where-Object { $_ -like '*.txt' })
        if (@($aud.TxtFaltando).Count -gt 0) {
            Write-Host "   TXT: FALTA no destino ($(@($aud.TxtFaltando) -join ', '))" -ForegroundColor Yellow
            Write-Host "        Origem [GD]: $($r.pasta_inbox)" -ForegroundColor DarkGray
        } elseif ($txtsReg.Count -gt 0 -and $aud.PastaExiste) {
            Write-Host "   TXT [OD]: $(Join-Path $r.destino_onedrive $txtsReg[0])" -ForegroundColor DarkGray
        } else {
            Write-Host "   TXT: nenhum .txt registrado nesta visita" -ForegroundColor DarkGray
        }
        # Fotos: o que a triagem copiou, de onde, e o que ha hoje no destino
        Write-Host "   Fotos: $($aud.FotosNoDestino) no destino / $($aud.FotosRegistradas) copiadas na triagem | modo: $(Get-PropValue $r 'fotos_modo')"
        $origemFotos = [string](Get-PropValue $r 'fotos_origem_sugerida')
        if ($origemFotos) { Write-Host "        Origem [Syncthing]: $origemFotos" -ForegroundColor DarkGray }
        if (Get-PropValue $r 'sincronizada') {
            Write-Host "   [GDSA]: enviada em $(Get-PropValue $r 'sincronizada_em')" -ForegroundColor Green
        } else {
            Write-Host "   [GDSA]: AINDA NAO ENVIADA (rode .\start-sync.ps1)" -ForegroundColor Yellow
        }
        if (-not $aud.PastaExiste) { Write-Host "   [!] Pasta de destino nao existe no OneDrive." -ForegroundColor Red }
        $op = (Read-Host "   [T] TXT/anexos [F] fotos faltantes [A] ambos [P] pular [Q] sair").Trim().ToUpper()
        if ($op -eq 'Q') { break }
        if ($op -eq 'P' -or $op -eq '') { continue }
        $mudou = $false
        if ($op -in @('T', 'A')) {
            $null = Repair-VisitaTxtAnexos -Registro $r
            $mudou = $true
        }
        if ($op -in @('F', 'A')) {
            if (Copy-FotosFaltantesVisita -Registro $r) { $mudou = $true }
        }
        if ($mudou) {
            Save-ProcessadosPorArquivo -Registros $regsConf
            Write-Host "   [OK] Registro [GD] atualizado: $($r.registro_path)" -ForegroundColor Cyan
        }
    }
    Write-TriagemLog "CONFERENCIA FINALIZADA"
    Write-Host "`nConferencia finalizada. Rode .\start-sync.ps1 para enviar o que falta de [OD] para [GDSA]." -ForegroundColor Cyan
}

# ================= CRIAR PASTAS + GRAVAR RELATORIOS (copia nao-destrutiva da inbox) =================
function Invoke-CriarPastasEArquivar {
    param([array]$RelatorioDias)
    $PastasCriadas = [System.Collections.Generic.List[string]]::new()

    foreach ($Dia in $RelatorioDias) {
        foreach ($V in $Dia.visitas) {
            $V.cidade = Get-NomeCidadeOneDrive $V.cidade   # legado "Juquiá" -> pasta local "Juquia" (sem acento)
            if (-not (Test-CampoSeguro $V.cidade) -or -not (Test-CampoSeguro $V.atividade) -or -not (Test-CampoSeguro $V.produtor_pasta)) {
                Write-Host "  [ERRO] Campos invalidos (path traversal?) em $($Dia.data): $($V.produtor_pasta)" -ForegroundColor Red
                Write-TriagemLog "REJEITADO path-traversal: $($Dia.data) / $($V.produtor_pasta)"
                continue
            }
            $dest = Join-Path (Join-Path (Join-Path $OneDriveBase $V.cidade) $V.atividade) $V.produtor_pasta
            $V.CaminhoDestino = $dest

            if (-not (Test-Path -LiteralPath $dest)) {
                try {
                    New-Item -ItemType Directory -Force -Path $dest | Out-Null
                    Write-Host "  [CRIADA]    [OD] $($V.cidade) > $($V.atividade) > $($V.produtor_pasta)" -ForegroundColor Green
                    Write-Host "              Caminho [OD]: $dest" -ForegroundColor DarkGray
                    Write-TriagemLog "CRIADA [OD]: $dest"
                    $PastasCriadas.Add($dest)
                } catch {
                    Write-Host "  [ERRO] Falha ao criar $dest : $($_.Exception.Message)" -ForegroundColor Red
                    Write-TriagemLog "ERRO criar: $dest - $($_.Exception.Message)"
                    continue
                }
            } else {
                Write-Host "  [JA EXISTE] [OD] $($V.cidade) > $($V.atividade) > $($V.produtor_pasta)" -ForegroundColor DarkGray
                Write-Host "              Caminho [OD]: $dest" -ForegroundColor DarkGray
            }

            # Copia NAO-DESTRUTIVA: originais da inbox ficam intactos no Drive.
            # No OneDrive vai APENAS o .txt limpo (sem YAML) + fotos/GPS + anexos ATV 1.
            # O .json original permanece somente na inbox do Drive pessoal (nao e copiado).
            $levados = [System.Collections.Generic.List[string]]::new()
            $nomeBase = if ($V.SourceTxt) { [System.IO.Path]::GetFileNameWithoutExtension($V.SourceTxt) }
                        elseif ($V.SourceJson) { [System.IO.Path]::GetFileNameWithoutExtension($V.SourceJson) }
                        else { "relatorio" }
            $body = $null
            if ($V.SourceTxt -and (Test-Path -LiteralPath $V.SourceTxt)) {
                $body = if ($V.TxtBody) { $V.TxtBody } else { Remove-YamlHeader (Get-Content -LiteralPath $V.SourceTxt -Raw -Encoding UTF8) }
            } elseif ($V.TxtBody) {
                $body = $V.TxtBody   # corpo montado do JSON (inbox sem .txt)
                if ($V.SourceJson) { $null = Write-TxtInboxDoJson -JsonPath $V.SourceJson }   # par na inbox com YAML no topo (aditivo)
            }
            if ($body) {
                # [OD] recebe APENAS o .txt limpo: sem YAML, sem Markdown, sem secao 10 (.json nunca sobe)
                $body = Remove-Secao10 (Remove-MarkdownBasico $body)
                Set-Content -LiteralPath (Join-Path $dest "$nomeBase.txt") -Value $body -Encoding UTF8
                $levados.Add("$nomeBase.txt")
            }

            # Anexos automaticos so na ATV 1: Mapa / Calendario Agricola (imagens na pasta da inbox)
            if (($V.atividade -ieq 'ATIVIDADE INDIVIDUAL 1') -and $V.PastaInbox -and (Test-Path -LiteralPath $V.PastaInbox)) {
                foreach ($a in @(Get-ChildItem -LiteralPath $V.PastaInbox -File -ErrorAction SilentlyContinue | Where-Object {
                    $_.Extension -in @('.png', '.jpg', '.jpeg') -and $_.BaseName -match 'mapa|c?alend[aá]rio'
                })) {
                    try {
                        Copy-Item -LiteralPath $a.FullName -Destination $dest -Force -ErrorAction Stop
                        $levados.Add($a.Name)
                        Write-Host "  [ANEXO ATV1] $($a.Name)" -ForegroundColor Green
                        Write-Host "              Origem [GD] : $($a.FullName)" -ForegroundColor DarkGray
                        Write-Host "              Destino [OD]: $(Join-Path $dest $a.Name)" -ForegroundColor DarkGray
                        Write-TriagemLog "ANEXO ATV1 [GD]: $($a.FullName) -> [OD]: $(Join-Path $dest $a.Name)"
                    } catch { Write-Host "  [!] Falha ao copiar anexo $($a.Name): $($_.Exception.Message)" -ForegroundColor Yellow }
                }
            }

            # Registro unico append em _processados.json (sem arvore _processados\; nada e movido/apagado)
            if ($V.OrigemInbox -and $V.PastaInbox) {
                try {
                    $regPathGD = Add-ProcessadoRegistro -Visita $V -Destino $dest -Arquivos @($levados)
                    Write-Host "  [REGISTRADO] $($V.produtor_pasta)" -ForegroundColor Cyan
                    Write-Host "              Registro [GD]: $regPathGD" -ForegroundColor DarkGray
                    Write-TriagemLog "REGISTRADO [GD]: $regPathGD | inbox: $($V.PastaInbox) -> [OD]: $dest"
                } catch {
                    Write-Host "  [!] Falha ao registrar em _processados.json: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-TriagemLog "ERRO registrar: $($_.Exception.Message)"
                }
            }
        }
    }
    return , @($PastasCriadas)
}

# ================= MANIFESTO DE AUDITORIA EM runs\ (NUNCA apagar) =================
function Write-TriagemRun {
    param($Registro)
    if (-not (Test-Path $RunsDir)) { New-Item -ItemType Directory -Force -Path $RunsDir | Out-Null }
    $ts = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $base = Join-Path $RunsDir "${ts}__TRIAGEM__triagem-campo"
    $Registro | ConvertTo-Json -Depth 10 | Set-Content -Path "$base.json" -Encoding UTF8

    $linhas = [System.Collections.Generic.List[string]]::new()
    $linhas.Add("TRIAGEM DE CAMPO - $ts")
    $linhas.Add("Visitas: $($Registro.contagem.visitas) | Pastas criadas: $($Registro.contagem.pastas) | Fotos copiadas: $($Registro.contagem.fotos)")
    $linhas.Add("Inbox [GD]        : $($Registro.inbox_gd)")
    foreach ($rg in @($Registro.registros_gd)) { $linhas.Add("Registro [GD]     : $rg") }
    $linhas.Add("Raizes [Syncthing]: $(@($Registro.raizes_syncthing) -join ' | ')")
    $linhas.Add("Base [OD]         : $($Registro.base_od)")
    $linhas.Add("")
    $linhas.Add("== PASTAS CRIADAS [OD] ==")
    foreach ($p in $Registro.pastas_criadas) { $linhas.Add("  $p") }
    $linhas.Add("")
    $linhas.Add("== ORIGENS [GD] (inbox Drive pessoal) ==")
    foreach ($o in $Registro.origens) { $linhas.Add("  [$($o.data)] $($o.produtor) | txt=$($o.txt) | json=$($o.json)") }
    $linhas.Add("")
    $linhas.Add("== FOTOS COPIADAS ([Syncthing] -> [OD]) ==")
    foreach ($f in $Registro.fotos_copiadas) { $linhas.Add("  $($f.origem) -> $($f.destino)") }
    $linhas.Add("")
    $linhas.Add("== GPS POR PRODUTOR [OD] ==")
    foreach ($g in $Registro.gps) { $linhas.Add("  $($g.produtor): $($g.count) foto(s) com GPS | medio=$($g.lat),$($g.lon) | $($g.txt)") }
    Set-Content -Path "$base.txt" -Value $linhas -Encoding UTF8
    Write-Host "`n[OK] Manifesto de auditoria salvo em: $base.json" -ForegroundColor Cyan
}

# ================= 1. CARREGAR VISITAS (INBOX -> visitas.json -> COLAR) =================
Clear-Host
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "   TRIAGEM DE CAMPO - INBOX DRIVE / SYNCTHING -> ONEDRIVE (PASTAS+GPS)  " -ForegroundColor White
Write-Host "========================================================================" -ForegroundColor Cyan

# Resumo rapido do historico do ano atual (G: offline/ausente -> silencioso)
$resumoProc = Get-ResumoProcessadosAno
if ($resumoProc) {
    Write-Host " Historico $($resumoProc.ano): $($resumoProc.triadas) visita(s) triadas | $($resumoProc.sincronizadas) sincronizada(s) | $($resumoProc.pendentes) pendente(s) de envio" -ForegroundColor DarkGray
    Write-Host "      Registro [GD]: $($resumoProc.caminho)" -ForegroundColor DarkGray
}

# Varredura FISICA do Drive (independente do _processados.json): situacao real da inbox no inicio
$ScanFisico = Get-InboxPastasFisicas
if (@($ScanFisico).Count -gt 0) {
    $nNova = @($ScanFisico | Where-Object { $_.Status -eq 'NOVA' }).Count
    $nPend = @($ScanFisico | Where-Object { $_.Status -eq 'TRIADA - PENDENTE SYNC' }).Count
    $nSync = @($ScanFisico | Where-Object { $_.Status -eq 'SINCRONIZADA' }).Count
    Write-Host " Drive [GD]: $(@($ScanFisico).Count) pasta(s) fisica(s) | $nNova nova(s) | $nPend pendente(s) de sync | $nSync sincronizada(s)" -ForegroundColor DarkGray
}

# ================= MODO -Conferir: LISTAR / CONFERIR / REPROCESSAR VISITAS JA PROCESSADAS =================
if ($Conferir) {
    Invoke-ConferirReprocessar
    exit 0
}

$VisitasBrutas = $null
try {
    if ($Colar) {
        $VisitasBrutas = ConvertFrom-DiasParsed (Read-JsonManual -Colar -Arquivo $VisitasJson)
    } else {
        $inbox = Get-InboxVisitas
        if ($inbox.Count -gt 0) {
            Write-Host " [OK] Inbox espelhada [GD]: $($inbox.Count) visita(s) detectada(s)" -ForegroundColor Green
            Write-Host "      Caminho [GD]: $InboxBase" -ForegroundColor DarkGray
            $VisitasBrutas = $inbox
        } else {
            # Sem visitas novas: lista o que existe FISICAMENTE no Drive (status cruzado do JSON),
            # em vez de apenas dizer "Nenhuma visita NOVA". Inbox vazia/offline -> fallback manual.
            if (@($ScanFisico).Count -gt 0) {
                Write-Host " [i] Nenhuma visita NOVA na inbox [GD] — pastas encontradas fisicamente no Drive:" -ForegroundColor Yellow
                Show-InboxFisico -Pastas $ScanFisico
                $opInbox = (Read-Host "`n [C] Conferir/reprocessar qualquer pasta | [M] Colar JSON manual | [Q] Sair").Trim().ToUpper()
                if ($opInbox -eq 'C') { Invoke-ConferirReprocessar; exit 0 }
                if ($opInbox -ne 'M') { Write-Host "[OK] Nada a fazer." -ForegroundColor Green; exit 0 }
            } else {
                Write-Host " [i] Inbox vazia/ausente. Usando visitas.json ou colagem." -ForegroundColor DarkGray
            }
            $VisitasBrutas = ConvertFrom-DiasParsed (Read-JsonManual -Arquivo $VisitasJson)
        }
    }
} catch {
    Write-Host "`n[ERRO] Falha ao ler visitas: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

if (-not $VisitasBrutas -or $VisitasBrutas.Count -eq 0) {
    Write-Host "`n[AVISO] Nenhuma visita encontrada (inbox, visitas.json ou colagem)." -ForegroundColor Yellow
    exit 0
}

$RelatorioDias = ConvertTo-RelatorioDias -Visitas $VisitasBrutas

# ================= ETAPA 1a: REVISAO INTERATIVA =================
$Confirmado = Show-RevisaoVisitas -RelatorioDias $RelatorioDias
if (-not $Confirmado) {
    Write-Host "`n[CANCELADO] Nenhuma pasta foi criada." -ForegroundColor Yellow
    Write-TriagemLog "REVISAO CANCELADA - nenhuma pasta criada"
    exit 0
}

# Origens para o manifesto (antes de copiar/registrar)
$RunOrigens = [System.Collections.Generic.List[object]]::new()
$RunFotos   = [System.Collections.Generic.List[object]]::new()
$RunGps     = [System.Collections.Generic.List[object]]::new()
$TotalVisitas = 0
foreach ($Dia in $RelatorioDias) {
    foreach ($V in $Dia.visitas) {
        $TotalVisitas++
        $RunOrigens.Add([PSCustomObject]@{ data = $Dia.data; produtor = $V.produtor_pasta; txt = $V.SourceTxt; json = $V.SourceJson })
    }
}

# ================= ETAPA 1b: CRIAR PASTAS + GRAVAR + ARQUIVAR =================
Write-Host "`n=== ETAPA 1: CRIANDO PASTAS / GRAVANDO RELATORIOS NO ONEDRIVE ===" -ForegroundColor Cyan
$PastasCriadas = Invoke-CriarPastasEArquivar -RelatorioDias $RelatorioDias

# Caminhos [GD] dos registros anuais (_processados.json) para o manifesto
$AnosRel = @($RelatorioDias | ForEach-Object { $_.data.Substring(0, 4) } | Sort-Object -Unique)
$RegistrosGD = @($AnosRel | ForEach-Object { Join-Path $InboxBase "$_\_inbox_relatorio_ater\_processados.json" } | Where-Object { Test-Path -LiteralPath $_ })

# ================= ETAPA 2: INDEXAR FOTOS DO SYNCTHING =================
$RaizesExistentes = @($SyncthingFotos | Where-Object { Test-Path $_ })
if ($RaizesExistentes.Count -eq 0) {
    Write-Host "`n[AVISO] Nenhuma raiz de fotos encontrada: $($SyncthingFotos -join ' | ')" -ForegroundColor Yellow
    Write-Host "Pastas/relatorios ja gravados. Ajuste `$SyncthingFotos para triar as fotos." -ForegroundColor Yellow
    Write-TriagemRun -Registro ([PSCustomObject]@{
        timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); origens = $RunOrigens
        pastas_criadas = $PastasCriadas; fotos_copiadas = $RunFotos; gps = $RunGps
        inbox_gd = $InboxBase; registros_gd = $RegistrosGD; raizes_syncthing = $SyncthingFotos; base_od = $OneDriveBase
        contagem = [PSCustomObject]@{ visitas = $TotalVisitas; pastas = $PastasCriadas.Count; fotos = 0 }
    })
    exit 0
}

Write-Host "`n=== ETAPA 2: BUSCA CIRURGICA POR DATA (ZERO EXIF NAS VISITAS ORGANIZADAS) ===" -ForegroundColor Cyan
# Datas unicas do relatorio -> so os meses/dias alvo sao acessados (nada de varrer o acervo).
$DatasAlvo = @($RelatorioDias | ForEach-Object { $_.data } | Sort-Object -Unique)
Write-Host " Datas alvo: $($DatasAlvo -join ', ') | Raizes [Syncthing]: $($RaizesExistentes -join ' | ')" -ForegroundColor DarkGray
# Organizada C: -> E: -> Camera bruta, dedup por nome; visitas vem prontas do _visita.json.
$TodasFotos = Get-FotosCirurgico -Raizes $RaizesExistentes -DatasAlvo $DatasAlvo -Extensoes $ExtensoesFoto
$QtdVisita = @($TodasFotos | Where-Object { $_.LoteVisita }).Count
Write-Host " [OK] $($TodasFotos.Count) foto(s) do(s) dia(s) alvo ($QtdVisita via _visita.json, sem abrir EXIF)" -ForegroundColor Green

# ================= ETAPA 2: DISTRIBUIR FOTOS (PRE-SUGESTAO + ENTER) =================
$Sair = $false
foreach ($Dia in $RelatorioDias) {
    if ($Sair) { break }
    # Rastreio de fotos por produtor do dia (tag -> Modo/Copiadas/Origem), gravado no _processados.json ao fim do dia
    $StatusFotos = @{}
    $marcar = {
        param($tag, $modo, $n, $origem)
        $atual = $StatusFotos[$tag]
        if (-not $atual) {
            $StatusFotos[$tag] = [PSCustomObject]@{ Modo = $modo; Copiadas = [int]$n; Origem = $origem }
            return
        }
        $atual.Copiadas += [int]$n
        if ($modo -eq 'manual') { $atual.Modo = 'manual'; if ($origem) { $atual.Origem = $origem } }
        elseif ($modo -ne 'pulado' -and $atual.Modo -ne 'manual') { $atual.Modo = $modo }
        if ($origem -and -not $atual.Origem) { $atual.Origem = $origem }
    }
    # Filtra pelo campo "data" (AAAA-MM-DD) do dia — nao pelo prefixo da pasta.
    $FotosDoDia = @($TodasFotos | Where-Object { $_.DataStr -eq $Dia.data })

    Write-Host "`n========================================================================" -ForegroundColor Cyan
    Write-Host " DATA: $($Dia.data) | Fotos encontradas: $($FotosDoDia.Count)" -ForegroundColor White
    Write-Host "========================================================================" -ForegroundColor Cyan
    if ($FotosDoDia.Count -eq 0) { Write-Host "  Nenhuma foto encontrada para esta data." -ForegroundColor DarkGray; continue }

    Write-Host " Produtores visitados neste dia:" -ForegroundColor Yellow
    foreach ($v in $Dia.visitas) {
        $per = [string]$v.periodo
        $hor = Format-HorarioVisita $v.hora_inicio $v.hora_fim
        if ($hor) { $per = ("$per $hor").Trim() }
        $extra = if ($per) { " ($per)" } else { "" }
        Write-Host "   [$($v.tag)]$extra $($v.cidade) | $($v.atividade) | $($v.produtor_pasta)" -ForegroundColor Green
    }
    Write-Host "   [Enter/S] Aceitar sugestao | [A/B/C] TAG | [E] Escolher pastas do dia | [P] Pular | [I] Foto a foto | [M] Manual | [Q] Sair" -ForegroundColor Gray
    Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan

    # 1) Lotes pre-agrupados por visita (tempo+GPS do organizar-camera) — preserva a separacao original
    $Lotes = [System.Collections.Generic.List[object]]::new()
    foreach ($g in @($FotosDoDia | Where-Object { $_.LoteVisita } | Group-Object -Property LoteVisita)) {
        $Lotes.Add(@($g.Group | Sort-Object DataHora))
    }
    # 2) Avulsas/brutas (sem visita): agrupa por proximidade no tempo (janela de deslocamento)
    $Restantes = @($FotosDoDia | Where-Object { -not $_.LoteVisita } | Sort-Object DataHora)
    $LoteAtual = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $Restantes) {
        if ($LoteAtual.Count -eq 0) { $LoteAtual.Add($f) }
        else {
            $DiffMin = ($f.DataHora - $LoteAtual[-1].DataHora).TotalMinutes
            if ($DiffMin -le $IntervaloVisitaMinutos) { $LoteAtual.Add($f) }
            else { $Lotes.Add($LoteAtual.ToArray()); $LoteAtual = [System.Collections.Generic.List[object]]::new(); $LoteAtual.Add($f) }
        }
    }
    if ($LoteAtual.Count -gt 0) { $Lotes.Add($LoteAtual.ToArray()) }
    $Lotes = @($Lotes | Sort-Object { $_[0].DataHora })

    $Consumidos = [System.Collections.Generic.HashSet[int]]::new()   # lotes ja copiados via [E] nao voltam ao prompt
    for ($iLote = 0; $iLote -lt $Lotes.Count; $iLote++) {
        if ($Consumidos.Contains($iLote)) { continue }
        $Lote = $Lotes[$iLote]
        $NumLote = $iLote + 1
        $QtdGps = @($Lote | Where-Object { $_.TemGps }).Count
        $Sugestao = Get-SugestaoTag -Lote $Lote -Visitas $Dia.visitas
        $hmLote = Get-HorarioMedioLote -Lote $Lote
        Write-Host "`n Lote #$NumLote : $($Lote.Count) foto(s) | $($hmLote.Faixa) | media $($hmLote.MediaHora) | GPS: $QtdGps" -ForegroundColor Cyan
        $origensLote = @($Lote | ForEach-Object { Get-OrigemCurtaFoto -Foto $_ } | Sort-Object -Unique)
        Write-Host "   Origem [Syncthing]: $($origensLote -join ' | ')  (AAAA\MM\pasta do organizar-camera)" -ForegroundColor DarkGray
        # Horario medio das fotos vs horario/periodo de cada visita do relatorio (conferencia visual)
        foreach ($v in $Dia.visitas) {
            $hr = Format-HorarioVisita $v.hora_inicio $v.hora_fim
            if (-not $hr) { $hr = [string]$v.periodo }
            if (-not $hr) { $hr = 'sem horario' }
            $dif = ''
            $iniV = ConvertTo-Minutos ([string]$v.hora_inicio)
            if ($null -ne $iniV) { $dif = " | dif $([int][Math]::Abs($hmLote.MediaMinutos - $iniV)) min" }
            Write-Host "   Relatorio [$($v.tag)]: $hr  vs  fotos media $($hmLote.MediaHora)$dif" -ForegroundColor DarkGray
        }
        if ($Lotes.Count -ge 2) {
            Write-Host "   [!] $($Lotes.Count) lotes/pastas neste dia — use [E] p/ escolher exatamente quais pastas copiar (nao misture visitas)." -ForegroundColor Yellow
        }
        if ($Sugestao) {
            $vSug = $Dia.visitas | Where-Object { $_.tag -eq $Sugestao } | Select-Object -First 1
            Write-Host "   Sugestao: [$Sugestao] $($vSug.produtor_pasta) — pressione ENTER ou S para aceitar" -ForegroundColor Yellow
        }

        $Escolha = Resolve-TagSugestao -Escolha ((Read-Host " TAG (A/B/C), Enter/S aceita sugestao, E escolhe pastas do dia, P pula, I individual, M manual, Q sair").Trim().ToUpper()) -Sugestao $Sugestao

        if ($Escolha -eq 'Q') { Write-Host "   -> Triagem encerrada." -ForegroundColor Yellow; Write-TriagemLog "ENCERRADO no lote #$NumLote de $($Dia.data)"; $Sair = $true; break }
        if ($Escolha -eq 'P' -or [string]::IsNullOrWhiteSpace($Escolha)) { Write-Host "   -> Lote pulado." -ForegroundColor DarkGray; continue }

        if ($Escolha -eq 'M') {
            # Manual: fotos de produtores vizinhos misturadas — abre origem + destinos no Explorer p/ copiar a mao
            $origemManual = Split-Path $Lote[0].Arquivo.FullName -Parent
            try { Invoke-Item -LiteralPath $origemManual } catch { }
            foreach ($v in $Dia.visitas) {
                if ($v.CaminhoDestino -and (Test-Path -LiteralPath $v.CaminhoDestino)) {
                    try { Invoke-Item -LiteralPath $v.CaminhoDestino } catch { }
                }
                & $marcar $v.tag 'manual' 0 $origemManual
            }
            Write-TriagemLog "MANUAL lote #$NumLote de $($Dia.data): origem [Syncthing] $origemManual"
            Write-Host "   -> Modo MANUAL: pastas abertas no Explorer. Copie as fotos manualmente; registro marcado como 'manual'." -ForegroundColor Yellow
            Write-Host "      Origem [Syncthing]: $origemManual" -ForegroundColor DarkGray
            foreach ($v in $Dia.visitas) {
                if ($v.CaminhoDestino) {
                    Write-Host "      Destino [OD]      : $($v.CaminhoDestino)" -ForegroundColor DarkGray
                    Write-TriagemLog "MANUAL destino [OD]: $($v.CaminhoDestino)"
                }
            }
            continue
        }

        if ($Escolha -eq 'I') {
            $PorTag = @{}
            foreach ($FotoItem in $Lote) {
                $TagFoto = (Read-Host "   [$($FotoItem.HoraStr)] $($FotoItem.Arquivo.Name) -> Tag (A/B/P)").Trim().ToUpper()
                if ($TagFoto -eq 'P' -or [string]::IsNullOrWhiteSpace($TagFoto)) { continue }
                $VisitaAlvo = $Dia.visitas | Where-Object { $_.tag -eq $TagFoto -and $_.CaminhoDestino } | Select-Object -First 1
                if ($VisitaAlvo) {
                    try {
                        Copy-Item -LiteralPath $FotoItem.Arquivo.FullName -Destination $VisitaAlvo.CaminhoDestino -Force -ErrorAction Stop
                        if (-not $PorTag.ContainsKey($TagFoto)) { $PorTag[$TagFoto] = [System.Collections.Generic.List[object]]::new() }
                        $PorTag[$TagFoto].Add($FotoItem)
                        $RunFotos.Add([PSCustomObject]@{ origem = $FotoItem.Arquivo.FullName; destino = (Join-Path $VisitaAlvo.CaminhoDestino $FotoItem.Arquivo.Name); produtor = $VisitaAlvo.produtor_pasta })
                        Write-TriagemLog "COPIADA [Syncthing]: $($FotoItem.Arquivo.FullName) -> [OD]: $(Join-Path $VisitaAlvo.CaminhoDestino $FotoItem.Arquivo.Name)"
                        & $marcar $TagFoto 'individual' 1 (Split-Path $FotoItem.Arquivo.FullName -Parent)
                    } catch { Write-Host "     [ERRO] $($_.Exception.Message)" -ForegroundColor Red }
                }
            }
            foreach ($k in $PorTag.Keys) {
                $vAlvo = $Dia.visitas | Where-Object { $_.tag -eq $k } | Select-Object -First 1
                $gps = Export-CoordenadasTxt -FotosLote $PorTag[$k].ToArray() -PastaDestino $vAlvo.CaminhoDestino -NomeProdutor $vAlvo.produtor_pasta
                $RunGps.Add([PSCustomObject]@{ produtor = $vAlvo.produtor_pasta; count = $gps.Count; lat = $gps.MediaLat; lon = $gps.MediaLon; txt = $gps.TxtPath })
                Write-Host "   [OK] $($PorTag[$k].Count) foto(s) + GPS ($($gps.Count)) copiadas para [OD]" -ForegroundColor Green
                Write-Host "        Destino [OD]: $($vAlvo.CaminhoDestino)" -ForegroundColor DarkGray
            }
            continue
        }

        if ($Escolha -eq 'E') {
            # Escolha explicita das pastas de origem do dia (setas + Espaco + Enter): copia SO o marcado
            $selPastas = Show-SeletorLotesDia -Lotes $Lotes -Data $Dia.data -Visitas $Dia.visitas
            if ($null -eq $selPastas) { continue }
            $indicesSel = @($selPastas)
            if ($indicesSel.Count -eq 0) { Write-Host "   -> Nenhuma pasta marcada." -ForegroundColor DarkGray; continue }
            Write-Host "   Produtores do dia:" -ForegroundColor Yellow
            foreach ($v in $Dia.visitas) { Write-Host "     [$($v.tag)] $($v.produtor_pasta)" -ForegroundColor Green }
            $tagE = (Read-Host "   Copiar as pastas marcadas para qual TAG? (Enter cancela)").Trim().ToUpper()
            $vE = $Dia.visitas | Where-Object { $_.tag -eq $tagE -and $_.CaminhoDestino } | Select-Object -First 1
            if (-not $vE) { Write-Host "   -> Tag invalida. Nada copiado." -ForegroundColor Red; continue }
            foreach ($idx in $indicesSel) {
                Write-Host "   Pasta: $(Get-OrigemCurtaFoto -Foto $Lotes[$idx][0]) -> [$tagE] $($vE.produtor_pasta)" -ForegroundColor Cyan
                Copy-LoteParaVisita -Lote $Lotes[$idx] -Visita $vE -Modo 'automatico' -RunFotos $RunFotos -RunGps $RunGps -Marcar $marcar
                [void]$Consumidos.Add($idx)
            }
            continue
        }

        $VisitaEscolhida = $Dia.visitas | Where-Object { $_.tag -eq $Escolha -and $_.CaminhoDestino } | Select-Object -First 1
        if ($VisitaEscolhida) {
            Copy-LoteParaVisita -Lote $Lote -Visita $VisitaEscolhida -Modo 'automatico' -RunFotos $RunFotos -RunGps $RunGps -Marcar $marcar
        } else {
            Write-Host "   [!] Tag '$Escolha' invalida. Lote pulado." -ForegroundColor Red
        }
    }
    # Grava o resultado da triagem de fotos do dia no _processados.json (atomico; demais campos intactos)
    Update-FotosProcessados -Dia $Dia -Status $StatusFotos
}

# ================= MANIFESTO FINAL =================
Write-TriagemRun -Registro ([PSCustomObject]@{
    timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); origens = $RunOrigens
    pastas_criadas = $PastasCriadas; fotos_copiadas = $RunFotos; gps = $RunGps
    inbox_gd = $InboxBase; registros_gd = $RegistrosGD; raizes_syncthing = $SyncthingFotos; base_od = $OneDriveBase
    contagem = [PSCustomObject]@{ visitas = $TotalVisitas; pastas = $PastasCriadas.Count; fotos = $RunFotos.Count }
})
Write-TriagemLog "TRIAGEM FINALIZADA | inbox [GD]: $InboxBase | base [OD]: $OneDriveBase"
Write-Host "`nTriagem finalizada! Rode .\start-sync.ps1 para enviar pastas, fotos e coordenadas_gps.txt de [OD] para [GDSA]." -ForegroundColor Cyan
Write-Host "  [OD]  : $OneDriveBase" -ForegroundColor DarkGray
Write-Host "  [GDSA]: gdrive_saconsult:VALE DO RIBEIRA/FERNANDO" -ForegroundColor DarkGray