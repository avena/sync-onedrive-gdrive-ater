<#
.SYNOPSIS
    Triagem de Fotos de Campo (Inbox Drive / Syncthing -> OneDrive ATER) + GPS EXIF + runs\
.DESCRIPTION
    Fonte primaria: inbox espelhada do Google Drive pessoal
    G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<Atividade>\<AAAA-MM-DD-NOME - RESUMO>\
    com pares .txt/.json (schema_version 1.0). Fallback: visitas.json manual ou colagem (FIM).
    Etapa 1: revisao interativa (TUI) -> cria pastas no OneDrive, grava .txt limpo (sem YAML)
    e .json, e arquiva os originais em _processados.
    Etapa 2: busca cirurgica pelas datas do relatorio (so <AAAA>\<MM> alvo; nome/hora/GPS
    lidos direto do _visita.json, zero EXIF nas visitas; avulsas/brutas pre-filtradas pelo
    nome *AAAAMMDD*). Visitas ja viram lotes pre-agrupados; avulsas agrupam por 40 min.
    Pre-sugere o produtor por horario/periodo (Enter ou S aceita), distribui por TAG e gera
    coordenadas_gps.txt. Registra manifesto de auditoria em runs\ (nunca apagado).
.PARAMETER Colar
    Forca colagem do JSON no terminal (ignora inbox e visitas.json).
.PARAMETER VisitasJson
    Caminho do arquivo de visitas manual (padrao: visitas.json ao lado do script).
.NOTES
    Requer PowerShell 7. System.Drawing le EXIF de JPG/JPEG; .heic/.mp4 caem no LastWriteTime.
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [switch]$Colar,
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

$Cidades = @(
    "Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro"
)
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

# Nome da pasta = prefixo AAAA-MM-DD- (dia completo) + NOME + - RESUMO.
function Format-NomePastaProdutor {
    param([string]$Nome, [string]$Data)
    if ([string]::IsNullOrWhiteSpace($Nome)) { return $Nome }
    $Nome = $Nome.Trim()
    $Nome = $Nome -replace '^\d{4}-\d{2}(-\d{2})?-', ''
    if ($Data -match '^\d{4}-\d{2}-\d{2}$') { $Nome = "$Data-$Nome" }
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

function New-Visita {
    param($Data, $Periodo, $HoraInicio, $HoraFim, $Cidade, $Atividade, $ProdutorPasta,
          $SourceTxt, $SourceJson, $TxtBody, $OrigemInbox, $Ano, $MM, $LeafName)
    return [PSCustomObject]@{
        data = $Data; periodo = $Periodo; hora_inicio = $HoraInicio; hora_fim = $HoraFim
        cidade = $Cidade; atividade = $Atividade; produtor_pasta = $ProdutorPasta; tag = $null
        SourceTxt = $SourceTxt; SourceJson = $SourceJson; TxtBody = $TxtBody
        OrigemInbox = $OrigemInbox; Ano = $Ano; MM = $MM; LeafName = $LeafName; CaminhoDestino = $null
    }
}

# ================= FONTE PRIMARIA: INBOX ESPELHADA =================
function Get-InboxVisitas {
    $visitas = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $InboxBase)) { return , @() }

    $anos = @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' })
    foreach ($ano in $anos) {
        $inbox = Join-Path $ano.FullName "_inbox_relatorio_ater"
        if (-not (Test-Path -LiteralPath $inbox)) { continue }
        $meses = @(Get-ChildItem -LiteralPath $inbox -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{2}$' })
        foreach ($mm in $meses) {
            foreach ($cid in @(Get-ChildItem -LiteralPath $mm.FullName -Directory -ErrorAction SilentlyContinue)) {
                foreach ($ativ in @(Get-ChildItem -LiteralPath $cid.FullName -Directory -ErrorAction SilentlyContinue)) {
                    foreach ($prod in @(Get-ChildItem -LiteralPath $ativ.FullName -Directory -ErrorAction SilentlyContinue)) {
                        $leaf = $prod.Name
                        $arqs = @(Get-ChildItem -LiteralPath $prod.FullName -File -ErrorAction SilentlyContinue |
                            Where-Object { $_.Extension -in @('.txt', '.json') })
                        if ($arqs.Count -eq 0) { continue }
                        $grupos = @($arqs | Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })
                        foreach ($g in $grupos) {
                            $txt  = $g.Group | Where-Object { $_.Extension -eq '.txt' }  | Select-Object -First 1
                            $json = $g.Group | Where-Object { $_.Extension -eq '.json' } | Select-Object -First 1

                            $data = $null; $periodo = $null; $hi = $null; $hf = $null
                            $cidade = $cid.Name; $atividade = $ativ.Name; $pasta = $leaf; $body = $null

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
                            if (-not $data -and $leaf -match '^(\d{4}-\d{2}-\d{2})') { $data = $Matches[1] }
                            if (-not $data) { continue }

                            # par pode ser so .txt ou so .json: $null.FullName quebraria no StrictMode
                            $visitas.Add((New-Visita -Data $data -Periodo $periodo -HoraInicio $hi -HoraFim $hf `
                                -Cidade $cidade -Atividade $atividade -ProdutorPasta $pasta `
                                -SourceTxt $(if ($txt) { $txt.FullName } else { $null }) `
                                -SourceJson $(if ($json) { $json.FullName } else { $null }) -TxtBody $body `
                                -OrigemInbox $true -Ano $ano.Name -MM $mm.Name -LeafName $leaf))
                        }
                    }
                }
            }
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

# Retorna a TAG se um unico produtor do dia casa com o horario medio do lote; senao $null
function Get-SugestaoTag {
    param($Lote, $Visitas)
    $media = ($Lote | ForEach-Object { $_.DataHora.Hour * 60 + $_.DataHora.Minute } | Measure-Object -Average).Average
    $casaram = [System.Collections.Generic.List[object]]::new()
    foreach ($v in $Visitas) {
        $ok = $false
        $ini = ConvertTo-Minutos $v.hora_inicio
        $fim = ConvertTo-Minutos $v.hora_fim
        if ($null -ne $ini -and $null -ne $fim) {
            if ($media -ge $ini -and $media -le $fim) { $ok = $true }
        } elseif ($v.periodo) {
            if ($v.periodo -eq 'manha' -and $media -ge 360 -and $media -le 720) { $ok = $true }
            elseif ($v.periodo -eq 'tarde' -and $media -ge 721 -and $media -le 1110) { $ok = $true }
        }
        if ($ok) { $casaram.Add($v) }
    }
    if ($casaram.Count -eq 1) { return $casaram[0].tag }
    return $null
}

# Enter (vazio) ou S confirmam a sugestao ativa; sem sugestao, S segue como tag invalida.
function Resolve-TagSugestao {
    param([string]$Escolha, $Sugestao)
    if (($Escolha -eq '' -or $Escolha -eq 'S') -and $Sugestao) { return $Sugestao }
    return $Escolha
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
                    $Resp = Read-Host "Deseja criar as pastas, gravar os relatorios no OneDrive e arquivar na inbox? (S/N)"
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

# ================= CRIAR PASTAS + GRAVAR RELATORIOS + ARQUIVAR INBOX =================
function Invoke-CriarPastasEArquivar {
    param([array]$RelatorioDias)
    $PastasCriadas = [System.Collections.Generic.List[string]]::new()

    foreach ($Dia in $RelatorioDias) {
        foreach ($V in $Dia.visitas) {
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
                    Write-Host "  [CRIADA]    $($V.cidade) > $($V.atividade) > $($V.produtor_pasta)" -ForegroundColor Green
                    Write-TriagemLog "CRIADA: $dest"
                    $PastasCriadas.Add($dest)
                } catch {
                    Write-Host "  [ERRO] Falha ao criar $dest : $($_.Exception.Message)" -ForegroundColor Red
                    Write-TriagemLog "ERRO criar: $dest - $($_.Exception.Message)"
                    continue
                }
            } else {
                Write-Host "  [JA EXISTE] $($V.cidade) > $($V.atividade) > $($V.produtor_pasta)" -ForegroundColor DarkGray
            }

            # Grava .txt limpo (sem YAML) e copia o .json irmao (ambos ficam locais/OneDrive)
            $nomeBase = if ($V.SourceTxt) { [System.IO.Path]::GetFileNameWithoutExtension($V.SourceTxt) } else { "relatorio" }
            if ($V.SourceTxt -and (Test-Path -LiteralPath $V.SourceTxt)) {
                $body = if ($V.TxtBody) { $V.TxtBody } else { Remove-YamlHeader (Get-Content -LiteralPath $V.SourceTxt -Raw -Encoding UTF8) }
                Set-Content -LiteralPath (Join-Path $dest "$nomeBase.txt") -Value $body -Encoding UTF8
            }
            if ($V.SourceJson -and (Test-Path -LiteralPath $V.SourceJson)) {
                Copy-Item -LiteralPath $V.SourceJson -Destination (Join-Path $dest ([System.IO.Path]::GetFileName($V.SourceJson))) -Force
            }

            # Arquiva originais da inbox em _processados\<MM>\<Cidade>\<Atividade>\<leaf>
            if ($V.OrigemInbox -and $V.Ano -and $V.MM) {
                $proc = Join-Path $InboxBase "$($V.Ano)\_inbox_relatorio_ater\_processados\$($V.MM)\$($V.cidade)\$($V.atividade)\$($V.LeafName)"
                try {
                    New-Item -ItemType Directory -Force -Path $proc | Out-Null
                    if ($V.SourceTxt -and (Test-Path -LiteralPath $V.SourceTxt))   { Move-Item -LiteralPath $V.SourceTxt -Destination $proc -Force }
                    if ($V.SourceJson -and (Test-Path -LiteralPath $V.SourceJson)) { Move-Item -LiteralPath $V.SourceJson -Destination $proc -Force }
                    Write-TriagemLog "ARQUIVADO inbox: $($V.SourceTxt) -> $proc"
                } catch {
                    Write-Host "  [!] Falha ao arquivar inbox: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-TriagemLog "ERRO arquivar: $($_.Exception.Message)"
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
    $linhas.Add("")
    $linhas.Add("== PASTAS CRIADAS NO ONEDRIVE ==")
    foreach ($p in $Registro.pastas_criadas) { $linhas.Add("  $p") }
    $linhas.Add("")
    $linhas.Add("== ORIGENS (inbox Drive) ==")
    foreach ($o in $Registro.origens) { $linhas.Add("  [$($o.data)] $($o.produtor) | txt=$($o.txt) | json=$($o.json)") }
    $linhas.Add("")
    $linhas.Add("== FOTOS COPIADAS (Syncthing -> OneDrive) ==")
    foreach ($f in $Registro.fotos_copiadas) { $linhas.Add("  $($f.origem) -> $($f.destino)") }
    $linhas.Add("")
    $linhas.Add("== GPS POR PRODUTOR ==")
    foreach ($g in $Registro.gps) { $linhas.Add("  $($g.produtor): $($g.count) foto(s) com GPS | medio=$($g.lat),$($g.lon) | $($g.txt)") }
    Set-Content -Path "$base.txt" -Value $linhas -Encoding UTF8
    Write-Host "`n[OK] Manifesto de auditoria salvo em: $base.json" -ForegroundColor Cyan
}

# ================= 1. CARREGAR VISITAS (INBOX -> visitas.json -> COLAR) =================
Clear-Host
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "   TRIAGEM DE CAMPO - INBOX DRIVE / SYNCTHING -> ONEDRIVE (PASTAS+GPS)  " -ForegroundColor White
Write-Host "========================================================================" -ForegroundColor Cyan

$VisitasBrutas = $null
try {
    if ($Colar) {
        $VisitasBrutas = ConvertFrom-DiasParsed (Read-JsonManual -Colar -Arquivo $VisitasJson)
    } else {
        $inbox = Get-InboxVisitas
        if ($inbox.Count -gt 0) {
            Write-Host " [OK] Inbox espelhada: $($inbox.Count) visita(s) detectada(s) em $InboxBase" -ForegroundColor Green
            $VisitasBrutas = $inbox
        } else {
            Write-Host " [i] Inbox vazia/ausente. Usando visitas.json ou colagem." -ForegroundColor DarkGray
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

# Origens para o manifesto (antes de arquivar/mover)
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

# ================= ETAPA 2: INDEXAR FOTOS DO SYNCTHING =================
$RaizesExistentes = @($SyncthingFotos | Where-Object { Test-Path $_ })
if ($RaizesExistentes.Count -eq 0) {
    Write-Host "`n[AVISO] Nenhuma raiz de fotos encontrada: $($SyncthingFotos -join ' | ')" -ForegroundColor Yellow
    Write-Host "Pastas/relatorios ja gravados. Ajuste `$SyncthingFotos para triar as fotos." -ForegroundColor Yellow
    Write-TriagemRun -Registro ([PSCustomObject]@{
        timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); origens = $RunOrigens
        pastas_criadas = $PastasCriadas; fotos_copiadas = $RunFotos; gps = $RunGps
        contagem = [PSCustomObject]@{ visitas = $TotalVisitas; pastas = $PastasCriadas.Count; fotos = 0 }
    })
    exit 0
}

Write-Host "`n=== ETAPA 2: BUSCA CIRURGICA POR DATA (ZERO EXIF NAS VISITAS ORGANIZADAS) ===" -ForegroundColor Cyan
# Datas unicas do relatorio -> so os meses/dias alvo sao acessados (nada de varrer o acervo).
$DatasAlvo = @($RelatorioDias | ForEach-Object { $_.data } | Sort-Object -Unique)
Write-Host " Datas alvo: $($DatasAlvo -join ', ') | Raizes: $($RaizesExistentes -join ' | ')" -ForegroundColor DarkGray
# Organizada C: -> E: -> Camera bruta, dedup por nome; visitas vem prontas do _visita.json.
$TodasFotos = Get-FotosCirurgico -Raizes $RaizesExistentes -DatasAlvo $DatasAlvo -Extensoes $ExtensoesFoto
$QtdVisita = @($TodasFotos | Where-Object { $_.LoteVisita }).Count
Write-Host " [OK] $($TodasFotos.Count) foto(s) do(s) dia(s) alvo ($QtdVisita via _visita.json, sem abrir EXIF)" -ForegroundColor Green

# ================= ETAPA 2: DISTRIBUIR FOTOS (PRE-SUGESTAO + ENTER) =================
$Sair = $false
foreach ($Dia in $RelatorioDias) {
    if ($Sair) { break }
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
    Write-Host "   [Enter/S] Aceitar sugestao | [A/B/C] TAG | [P] Pular | [I] Foto a foto | [Q] Sair" -ForegroundColor Gray
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

    $NumLote = 0
    foreach ($Lote in $Lotes) {
        $NumLote++
        $QtdGps = @($Lote | Where-Object { $_.TemGps }).Count
        $Sugestao = Get-SugestaoTag -Lote $Lote -Visitas $Dia.visitas
        $OrigemLote = if ($Lote[0].LoteVisita) { " | visita: $(Split-Path $Lote[0].LoteVisita -Leaf)" } else { '' }
        Write-Host "`n Lote #$NumLote : $($Lote.Count) foto(s) | $($Lote[0].HoraStr) as $($Lote[-1].HoraStr) | GPS: $QtdGps$OrigemLote" -ForegroundColor Cyan
        if ($Sugestao) {
            $vSug = $Dia.visitas | Where-Object { $_.tag -eq $Sugestao } | Select-Object -First 1
            Write-Host "   Sugestao: [$Sugestao] $($vSug.produtor_pasta) — pressione ENTER ou S para aceitar" -ForegroundColor Yellow
        }

        $Escolha = Resolve-TagSugestao -Escolha ((Read-Host " TAG (A/B/C), Enter/S aceita sugestao, P pula, I individual, Q sair").Trim().ToUpper()) -Sugestao $Sugestao

        if ($Escolha -eq 'Q') { Write-Host "   -> Triagem encerrada." -ForegroundColor Yellow; Write-TriagemLog "ENCERRADO no lote #$NumLote de $($Dia.data)"; $Sair = $true; break }
        if ($Escolha -eq 'P' -or [string]::IsNullOrWhiteSpace($Escolha)) { Write-Host "   -> Lote pulado." -ForegroundColor DarkGray; continue }

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
                        Write-TriagemLog "COPIADA: $($FotoItem.Arquivo.Name) -> $($VisitaAlvo.produtor_pasta)"
                    } catch { Write-Host "     [ERRO] $($_.Exception.Message)" -ForegroundColor Red }
                }
            }
            foreach ($k in $PorTag.Keys) {
                $vAlvo = $Dia.visitas | Where-Object { $_.tag -eq $k } | Select-Object -First 1
                $gps = Export-CoordenadasTxt -FotosLote $PorTag[$k].ToArray() -PastaDestino $vAlvo.CaminhoDestino -NomeProdutor $vAlvo.produtor_pasta
                $RunGps.Add([PSCustomObject]@{ produtor = $vAlvo.produtor_pasta; count = $gps.Count; lat = $gps.MediaLat; lon = $gps.MediaLon; txt = $gps.TxtPath })
                Write-Host "   [OK] $($PorTag[$k].Count) foto(s) + GPS ($($gps.Count)) em $($vAlvo.produtor_pasta)" -ForegroundColor Green
            }
            continue
        }

        $VisitaEscolhida = $Dia.visitas | Where-Object { $_.tag -eq $Escolha -and $_.CaminhoDestino } | Select-Object -First 1
        if ($VisitaEscolhida) {
            foreach ($FotoItem in $Lote) {
                try {
                    Copy-Item -LiteralPath $FotoItem.Arquivo.FullName -Destination $VisitaEscolhida.CaminhoDestino -Force -ErrorAction Stop
                    $RunFotos.Add([PSCustomObject]@{ origem = $FotoItem.Arquivo.FullName; destino = (Join-Path $VisitaEscolhida.CaminhoDestino $FotoItem.Arquivo.Name); produtor = $VisitaEscolhida.produtor_pasta })
                    Write-TriagemLog "COPIADA: $($FotoItem.Arquivo.Name) -> $($VisitaEscolhida.produtor_pasta)"
                } catch { Write-Host "   [!] Falha ao copiar $($FotoItem.Arquivo.Name): $($_.Exception.Message)" -ForegroundColor Red }
            }
            $gps = Export-CoordenadasTxt -FotosLote $Lote -PastaDestino $VisitaEscolhida.CaminhoDestino -NomeProdutor $VisitaEscolhida.produtor_pasta
            $RunGps.Add([PSCustomObject]@{ produtor = $VisitaEscolhida.produtor_pasta; count = $gps.Count; lat = $gps.MediaLat; lon = $gps.MediaLon; txt = $gps.TxtPath })
            Write-Host "   [OK] $($Lote.Count) foto(s) + GPS ($($gps.Count)) em: $($VisitaEscolhida.produtor_pasta)" -ForegroundColor Green
        } else {
            Write-Host "   [!] Tag '$Escolha' invalida. Lote pulado." -ForegroundColor Red
        }
    }
}

# ================= MANIFESTO FINAL =================
Write-TriagemRun -Registro ([PSCustomObject]@{
    timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); origens = $RunOrigens
    pastas_criadas = $PastasCriadas; fotos_copiadas = $RunFotos; gps = $RunGps
    contagem = [PSCustomObject]@{ visitas = $TotalVisitas; pastas = $PastasCriadas.Count; fotos = $RunFotos.Count }
})
Write-TriagemLog "TRIAGEM FINALIZADA"
Write-Host "`nTriagem finalizada! Rode .\start-sync.ps1 para enviar pastas, fotos e coordenadas_gps.txt ao Google Drive." -ForegroundColor Cyan