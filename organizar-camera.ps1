<#
.SYNOPSIS
    Organiza fotos da camera do Android por VISITA (tempo + raio GPS) em
    Camera_Celular_Organizada, com _visita.json, estado persistente e
    sincronizacao campo (C:) vs. casa (disco externo E:).
.DESCRIPTION
    Origem: C:\Users\randomacesso\Sync\Camera_Celular\Camera (Syncthing).
    Destino local: C:\Users\randomacesso\Sync\Camera_Celular_Organizada\
      Visitas (clusters >=3 fotos): <AAAA>\<MM>\<DD>_visita\ (1 visita no dia)
      ou <DD>_visita_1, <DD>_visita_2... (2+ visitas, ordem cronologica).
      Avulsas (1-2 fotos): raiz <AAAA>\<MM>\.
      Videos (VID_*/.mp4/.mov/.3gp/.mkv): Videos\<AAAA>\<MM>\.
    Cluster: (dt<=40min E d<=500m) OU (mesmo dia E dt<=120min E d<=300m)
    OU (sem GPS E dt<=40min). Haversine em metros.
    Cada pasta de visita ganha _visita.json (data, horarios, GPS medio,
    raio, URL Maps, arquivos com hora/lat/lon/modelo/tamanho).
    estado-camera.json (versionado) registra totais e pendencias C:->E:.
    Disco externo ausente (campo): organiza so no C: (pergunta, ou -SemMenu).
    Presente (casa): menu TUI setas — organizar+sincronizar / so organizar /
    so sincronizar / cancelar; sincronizar pode MOVER do C: p/ E: (libera SSD).
    Manifesto: runs\AAAA-MM-DD_HH-mm-ss__CAMERA__organizar-camera.json.
.PARAMETER SemMenu
    Nao interativo: organiza no C:; com -SincronizarExterno tambem sincroniza.
.PARAMETER SincronizarExterno
    Sincroniza tudo (novas + pendentes) do C: para o disco externo.
.PARAMETER MoverParaExterno
    Ao sincronizar, MOVE do C: para o E: (libera espaco no SSD C:).
.PARAMETER DriveExterno
    Letra da unidade externa (ex.: "E"). Padrao: prioriza E:, senao outra
    unidade secundaria conectada alem de C:.
.PARAMETER Copiar
    Copia em vez de mover da Camera\ (mantem o bruto no Syncthing).
.EXAMPLE
    .\organizar-camera.ps1 -SemMenu -WhatIf   # simulacao (254 clusters esperados)
    .\organizar-camera.ps1                    # interativo (menu campo/casa)
    .\organizar-camera.ps1 -SemMenu -SincronizarExterno -MoverParaExterno
.NOTES
    Requer PowerShell 7. Verbos aprovados (Get-GpsCoord etc. — o alias gps
    do PowerShell aponta para Get-Process; nunca nomear funcao como Gps).
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$SemMenu,
    [switch]$SincronizarExterno,
    [switch]$MoverParaExterno,
    [string]$DriveExterno,
    [switch]$Copiar
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ================= CONFIGURACOES =================
$OrigemCamera  = "C:\Users\randomacesso\Sync\Camera_Celular\Camera"
$DestinoLocal  = "C:\Users\randomacesso\Sync\Camera_Celular_Organizada"
$EstadoFile    = Join-Path $PSScriptRoot "estado-camera.json"   # versionado no Git
$RunsDir       = Join-Path $PSScriptRoot "runs"                 # permanente
$ExtensoesFoto  = @(".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp", ".dng")
$ExtensoesVideo = @(".mp4", ".mov", ".3gp", ".mkv", ".avi")

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ================= HELPERS DE METADADOS =================
function Get-RacionalExif {
    param([byte[]]$Bytes, [int]$Offset)
    $n = [BitConverter]::ToUInt32($Bytes, $Offset)
    $d = [BitConverter]::ToUInt32($Bytes, $Offset + 4)
    if ($d -eq 0) { return 0.0 }
    return $n / $d
}

function Get-GpsCoord {
    # Coordenada decimal a partir dos tags EXIF (2/1 lat, 4/3 lon). NUNCA "Gps" (alias).
    param($Img, [int]$IdCoord, [int]$IdRef)
    $p = $Img.PropertyItems | Where-Object { $_.Id -eq $IdCoord } | Select-Object -First 1
    $r = $Img.PropertyItems | Where-Object { $_.Id -eq $IdRef } | Select-Object -First 1
    if (-not $p -or -not $r -or $p.Value.Length -lt 24) { return $null }
    $v = (Get-RacionalExif $p.Value 0) + (Get-RacionalExif $p.Value 8) / 60 + (Get-RacionalExif $p.Value 16) / 3600
    $c = [char]$r.Value[0]
    if ($c -eq 'S' -or $c -eq 'W') { $v = -$v }
    if ($v -eq 0) { return $null }
    return $v
}

function Get-AsciiExif {
    param($Img, [int]$Id)
    $p = $Img.PropertyItems | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
    if ($p -and $p.Value) { return ([System.Text.Encoding]::ASCII.GetString($p.Value)).Trim([char]0).Trim() }
    return ""
}

function Get-MetadadosCamera {
    # DataHora: EXIF DateTimeOriginal -> nome IMG_/VID_AAAAMMDD_HHMMSS -> LastWriteTime.
    param([System.IO.FileInfo]$Arquivo)
    $dt = $null; $lat = $null; $lon = $null; $modelo = ""
    $ext = $Arquivo.Extension.ToLower()
    if ($ext -in $ExtensoesFoto) {
        $img = $null
        try {
            $img = [System.Drawing.Image]::FromFile($Arquivo.FullName)
            $s = Get-AsciiExif $img 36867
            if ($s) { try { $dt = [datetime]::ParseExact($s, "yyyy:MM:dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture) } catch { } }
            $lat = Get-GpsCoord $img 2 1
            $lon = Get-GpsCoord $img 4 3
            $mk = Get-AsciiExif $img 271; $md = Get-AsciiExif $img 272
            $modelo = ("$mk $md").Trim()
        } catch { } finally { if ($img) { $img.Dispose() } }
    }
    $m = [regex]::Match($Arquivo.Name, '(?<!\d)(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})(?!\d)')
    if (-not $dt -and $m.Success) {
        $dt = [datetime]::new([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value, [int]$m.Groups[4].Value, [int]$m.Groups[5].Value, [int]$m.Groups[6].Value)
    }
    if (-not $dt) { $dt = $Arquivo.LastWriteTime }
    return [PSCustomObject]@{
        Arquivo = $Arquivo; DataHora = $dt; Lat = $lat; Lon = $lon; Modelo = $modelo
        TemGps = ($null -ne $lat -and $null -ne $lon)
    }
}

function Get-DistanciaMetros {
    # Haversine.
    param([double]$Lat1, [double]$Lon1, [double]$Lat2, [double]$Lon2)
    $R = 6371000.0; $p = [Math]::PI / 180
    $dLa = ($Lat2 - $Lat1) * $p; $dLo = ($Lon2 - $Lon1) * $p
    $a = [Math]::Sin($dLa / 2) * [Math]::Sin($dLa / 2) + [Math]::Cos($Lat1 * $p) * [Math]::Cos($Lat2 * $p) * [Math]::Sin($dLo / 2) * [Math]::Sin($dLo / 2)
    return 2 * $R * [Math]::Asin([Math]::Sqrt($a))
}

# ================= CLUSTERIZACAO DE VISITAS =================
function Get-ClustersVisita {
    # Regras: (dt<=40min E d<=500m) OU (mesmo dia E dt<=120min E d<=300m) OU (sem GPS E dt<=40min).
    param($Fotos)
    $ordenadas = @($Fotos | Sort-Object DataHora)
    $clusters = [System.Collections.Generic.List[object]]::new()
    $cur = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $ordenadas) {
        if ($cur.Count -eq 0) { $cur.Add($r); continue }
        $prev = $cur[$cur.Count - 1]
        $min = ($r.DataHora - $prev.DataHora).TotalMinutes
        $mesmoDia = $r.DataHora.Date -eq $prev.DataHora.Date
        $ambosGps = $prev.TemGps -and $r.TemGps
        $ok = $false
        if ($ambosGps) {
            $d = Get-DistanciaMetros $prev.Lat $prev.Lon $r.Lat $r.Lon
            if (($min -le 40 -and $d -le 500) -or ($mesmoDia -and $min -le 120 -and $d -le 300)) { $ok = $true }
        } elseif ($min -le 40) { $ok = $true }
        if ($ok) { $cur.Add($r) }
        else { $clusters.Add($cur); $cur = [System.Collections.Generic.List[object]]::new(); $cur.Add($r) }
    }
    if ($cur.Count -gt 0) { $clusters.Add($cur) }
    return , $clusters
}

function Set-NomesVisita {
    # DD_visita (1 visita no dia) ou DD_visita_1..N (2+, cronologico). Avulsos: raiz do mes.
    param($Clusters)
    $visitas = @($Clusters | Where-Object { $_.Count -ge 3 })
    $porDia = @($visitas | Group-Object { $_[0].DataHora.Date })
    foreach ($g in $porDia) {
        $ordenadas = @($g.Group | Sort-Object { $_[0].DataHora })
        $i = 1
        foreach ($c in $ordenadas) {
            $dd = $c[0].DataHora.ToString('dd')
            $nome = if ($ordenadas.Count -eq 1) { "${dd}_visita" } else { "${dd}_visita_$i" }
            Add-Member -InputObject $c -NotePropertyName PastaNome -NotePropertyValue $nome -Force
            Add-Member -InputObject $c -NotePropertyName IndiceVisita -NotePropertyValue $i -Force
            $i++
        }
    }
    return , $visitas
}

function New-VisitaJson {
    param($Cluster, [string]$PastaNome, [int]$Indice)
    $ini = $Cluster[0].DataHora; $fim = $Cluster[-1].DataHora
    $comGps = @($Cluster | Where-Object { $_.TemGps })
    $lat = $null; $lon = $null; $raio = 0.0; $url = $null
    if ($comGps.Count -gt 0) {
        $lat = [math]::Round(($comGps | ForEach-Object Lat | Measure-Object -Average).Average, 6)
        $lon = [math]::Round(($comGps | ForEach-Object Lon | Measure-Object -Average).Average, 6)
        foreach ($x in $comGps) {
            $d = Get-DistanciaMetros $lat $lon $x.Lat $x.Lon
            if ($d -gt $raio) { $raio = $d }
        }
        $raio = [math]::Round($raio, 1)
        $latStr = $lat.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $lonStr = $lon.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $url = "https://www.google.com/maps?q=$latStr,$lonStr"
    }
    return [PSCustomObject]@{
        data = $ini.ToString('yyyy-MM-dd'); indice_visita = $Indice; pasta_nome = $PastaNome
        hora_inicio = $ini.ToString('HH:mm:ss'); hora_fim = $fim.ToString('HH:mm:ss')
        duracao_minutos = [math]::Round(($fim - $ini).TotalMinutes, 1)
        qtd_fotos = $Cluster.Count; fotos_com_gps = $comGps.Count
        lat_media = $lat; lon_media = $lon; raio_metros = $raio; google_maps_url = $url
        arquivos = @($Cluster | ForEach-Object {
            [PSCustomObject]@{
                nome = $_.Arquivo.Name; hora = $_.DataHora.ToString('HH:mm:ss')
                lat = $_.Lat; lon = $_.Lon; modelo_camera = $_.Modelo; tamanho_bytes = $_.Arquivo.Length
            }
        })
    }
}

# ================= ESTADO PERSISTENTE (estado-camera.json, versionado) =================
function Get-EstadoCamera {
    if (Test-Path -LiteralPath $EstadoFile) {
        try { return (Get-Content -LiteralPath $EstadoFile -Raw | ConvertFrom-Json) } catch { }
    }
    return [PSCustomObject]@{
        ultima_execucao = $null; ultima_foto_processada = $null
        total_fotos_organizadas_c = 0; total_videos_organizados_c = 0
        ultima_sincronizacao_e = $null
        pendentes_sync_e = [PSCustomObject]@{ arquivos = 0; bytes = 0 }
    }
}

function Save-EstadoCamera {
    param($Estado)
    $Estado | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $EstadoFile -Encoding UTF8
}

# ================= DISCO EXTERNO =================
function Get-DriveExterno {
    # Prioriza E:; senao qualquer unidade secundaria (Fixa/Removivel) alem de C:.
    if ($DriveExterno) {
        $l = $DriveExterno.TrimEnd(':', '\').ToUpper()
        if (Test-Path "${l}:\") { return $l }
        return $null
    }
    # ponytail: exclui rede e o disco virtual do Google Drive (G:, label "Google Drive",
    # que reporta Fixed) pelo rotulo; upgrade = checar bus type do dispositivo se surgirem
    # outros discos virtuais. Sem isso, no campo (E: ausente) o acervo iria para o Drive pessoal.
    $unidades = @([System.IO.DriveInfo]::GetDrives() | Where-Object {
        $_.IsReady -and $_.DriveType -in @('Fixed', 'Removable') -and
        $_.Name -ne 'C:\' -and $_.VolumeLabel -ne 'Google Drive'
    })
    $e = $unidades | Where-Object { $_.Name -eq 'E:\' } | Select-Object -First 1
    if ($e) { return 'E' }
    $outra = $unidades | Sort-Object Name | Select-Object -First 1
    if ($outra) { return $outra.Name.Substring(0, 1) }   # DriveInfo.Name = "E:\" -> letra E
    return $null
}

function Get-PendentesSync {
    # Arquivos no C: ainda nao presentes no E: (por caminho relativo).
    param([string]$DestinoE)
    $pend = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $DestinoLocal)) { return , $pend }
    foreach ($f in @(Get-ChildItem -LiteralPath $DestinoLocal -Recurse -File)) {
        $rel = $f.FullName.Substring($DestinoLocal.Length + 1)
        $noE = if ($DestinoE) { Join-Path $DestinoE $rel } else { $null }
        if (-not $noE -or -not (Test-Path -LiteralPath $noE)) {
            $pend.Add([PSCustomObject]@{ Relativo = $rel; Bytes = $f.Length })
        }
    }
    return , $pend
}

# ================= VARREDURA DA CAMERA =================
function Get-ArquivosCamera {
    param([string]$Raiz)
    return @(Get-ChildItem -LiteralPath $Raiz -Recurse -File | Where-Object {
        $_.Name -notlike '.*' -and $_.Name -notlike '~syncthing~*' -and $_.Extension -ne '.tmp'
    })
}

function Get-NomeSemColisaoCamera {
    param([string]$Pasta, [System.IO.FileInfo]$Arquivo)
    $dest = Join-Path $Pasta $Arquivo.Name
    if (-not (Test-Path -LiteralPath $dest)) { return $dest }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Arquivo.Name)
    $i = 2
    do { $dest = Join-Path $Pasta ("{0}__{1}{2}" -f $base, $i, $Arquivo.Extension); $i++ } while (Test-Path -LiteralPath $dest)
    return $dest
}

# ================= ORGANIZACAO (C:) =================
function Invoke-OrganizarNovas {
    # Move (ou copia com -Copiar) fotos/videos novos da Camera\ para a estrutura organizada.
    param($Metadados, $Videos, [string]$Destino)
    $clusters = Get-ClustersVisita -Fotos $Metadados
    $visitas = Set-NomesVisita -Clusters $clusters
    $avulsos = @($clusters | Where-Object { $_.Count -lt 3 })
    $movidos = [System.Collections.Generic.List[object]]::new()

    foreach ($c in $visitas) {
        $pasta = Join-Path $Destino (Join-Path $c[0].DataHora.ToString('yyyy\\MM') $c.PastaNome)
        if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Force -Path $pasta | Out-Null }
        foreach ($m in $c) {
            $dest = Get-NomeSemColisaoCamera -Pasta $pasta -Arquivo $m.Arquivo
            if ($Copiar) { Copy-Item -LiteralPath $m.Arquivo.FullName -Destination $dest }
            else { Move-Item -LiteralPath $m.Arquivo.FullName -Destination $dest }
            $movidos.Add([PSCustomObject]@{ origem = $m.Arquivo.FullName; destino = $dest; tipo = 'visita'; bytes = $m.Arquivo.Length })
        }
        $json = New-VisitaJson -Cluster $c -PastaNome $c.PastaNome -Indice $c.IndiceVisita
        $json | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $pasta '_visita.json') -Encoding UTF8
    }
    foreach ($c in $avulsos) {
        foreach ($m in $c) {
            $pasta = Join-Path $Destino $m.DataHora.ToString('yyyy\\MM')
            if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Force -Path $pasta | Out-Null }
            $dest = Get-NomeSemColisaoCamera -Pasta $pasta -Arquivo $m.Arquivo
            if ($Copiar) { Copy-Item -LiteralPath $m.Arquivo.FullName -Destination $dest }
            else { Move-Item -LiteralPath $m.Arquivo.FullName -Destination $dest }
            $movidos.Add([PSCustomObject]@{ origem = $m.Arquivo.FullName; destino = $dest; tipo = 'avulsa'; bytes = $m.Arquivo.Length })
        }
    }
    $vidMovidos = [System.Collections.Generic.List[object]]::new()
    foreach ($v in $Videos) {
        $pasta = Join-Path $Destino (Join-Path 'Videos' $v.DataHora.ToString('yyyy\\MM'))
        if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Force -Path $pasta | Out-Null }
        $dest = Get-NomeSemColisaoCamera -Pasta $pasta -Arquivo $v.Arquivo
        if ($Copiar) { Copy-Item -LiteralPath $v.Arquivo.FullName -Destination $dest }
        else { Move-Item -LiteralPath $v.Arquivo.FullName -Destination $dest }
        $vidMovidos.Add([PSCustomObject]@{ origem = $v.Arquivo.FullName; destino = $dest; tipo = 'video'; bytes = $v.Arquivo.Length })
    }
    $fv = 0; foreach ($c in $visitas) { $fv += $c.Count }
    $fa = 0; foreach ($c in $avulsos) { $fa += $c.Count }
    return [PSCustomObject]@{
        Clusters = $clusters.Count; Visitas = $visitas.Count; Avulsos = $avulsos.Count
        FotosVisitas = $fv; FotosAvulsas = $fa
        Movidos = $movidos; VideosMovidos = $vidMovidos
    }
}

# ================= SINCRONIZACAO C: -> E: =================
function Invoke-SincronizarExterno {
    # Copia (ou move com -MoverParaExterno) tudo que esta no C: para o E:, preservando a arvore.
    param([string]$DestinoE, [bool]$Mover)
    $feitos = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $DestinoLocal)) { return , $feitos }
    foreach ($f in @(Get-ChildItem -LiteralPath $DestinoLocal -Recurse -File)) {
        $rel = $f.FullName.Substring($DestinoLocal.Length + 1)
        $dest = Join-Path $DestinoE $rel
        if (Test-Path -LiteralPath $dest) { continue }   # ja sincronizado
        $pastaE = Split-Path $dest -Parent
        if (-not (Test-Path -LiteralPath $pastaE)) { New-Item -ItemType Directory -Force -Path $pastaE | Out-Null }
        if ($Mover) { Move-Item -LiteralPath $f.FullName -Destination $dest }
        else { Copy-Item -LiteralPath $f.FullName -Destination $dest }
        $feitos.Add([PSCustomObject]@{ origem = $f.FullName; destino = $dest; bytes = $f.Length })
    }
    if ($Mover) {
        # remove pastas vazias do C: (bottom-up), preservando a raiz
        $dirs = @(Get-ChildItem -LiteralPath $DestinoLocal -Recurse -Directory |
            Sort-Object { $_.FullName.Split('\').Count } -Descending)
        foreach ($d in $dirs) {
            if (@(Get-ChildItem -LiteralPath $d.FullName -Force).Count -eq 0) {
                Remove-Item -LiteralPath $d.FullName -Force
            }
        }
    }
    return , $feitos
}

# ================= MENU TUI (CENARIO B — disco externo presente) =================
function Show-MenuCasa {
    param([string]$Drive)
    $opcoes = @(
        "Organizar novas no C: E sincronizar tudo (novas + pendentes) para ${Drive}:\Sync\Camera_Celular_Organizada",
        "Apenas organizar novas no C: (sem sincronizar agora)",
        "Apenas sincronizar/mover o que ja esta no C: para ${Drive}:",
        "Cancelar"
    )
    $sel = 0
    [Console]::CursorVisible = $false
    try {
        while ($true) {
            Write-Host "`n  DISCO EXTERNO ${Drive}: CONECTADO (em casa/escritorio)" -ForegroundColor Green
            for ($i = 0; $i -lt $opcoes.Count; $i++) {
                $marca = if ($i -eq $sel) { '>' } else { ' ' }
                $cor = if ($i -eq $sel) { 'Cyan' } else { 'Gray' }
                Write-Host ("  {0} [{1}] {2}" -f $marca, ($i + 1), $opcoes[$i]) -ForegroundColor $cor
            }
            Write-Host "  (setas cima/baixo + Enter; Q cancela)" -ForegroundColor DarkGray
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq 'UpArrow') { $sel = ($sel - 1 + $opcoes.Count) % $opcoes.Count }
            elseif ($k.Key -eq 'DownArrow') { $sel = ($sel + 1) % $opcoes.Count }
            elseif ($k.Key -eq 'Enter') { return $sel }
            elseif ($k.Key -eq 'Escape' -or $k.Key -eq 'Q') { return 3 }
            for ($i = 0; $i -lt ($opcoes.Count + 2); $i++) { Write-Host "`e[1A`e[2K" -NoNewline }
        }
    } finally { [Console]::CursorVisible = $true }
}

function Read-Confirmacao {
    param([string]$Pergunta)
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch { }
    $r = Read-Host $Pergunta
    return ($r -match '^[Ss]')
}

# ================= FLUXO PRINCIPAL =================
$estado = Get-EstadoCamera
Write-Host "ORGANIZAR CAMERA: $OrigemCamera -> $DestinoLocal" -ForegroundColor White
if (-not (Test-Path -LiteralPath $OrigemCamera)) { throw "Pasta de origem nao encontrada: $OrigemCamera" }

$todos = Get-ArquivosCamera -Raiz $OrigemCamera
$fotosBrutas = @($todos | Where-Object { $ExtensoesFoto -contains $_.Extension.ToLower() })
$videosBrutos = @($todos | Where-Object { $ExtensoesVideo -contains $_.Extension.ToLower() })

# Modo -Copiar: pula o que ja existe no destino (dedup por nome)
$jaOrganizados = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if ($Copiar -and (Test-Path -LiteralPath $DestinoLocal)) {
    foreach ($f in @(Get-ChildItem -LiteralPath $DestinoLocal -Recurse -File)) { [void]$jaOrganizados.Add($f.Name) }
}
$fotosBrutas = @($fotosBrutas | Where-Object { -not $jaOrganizados.Contains($_.Name) })
$videosBrutos = @($videosBrutos | Where-Object { -not $jaOrganizados.Contains($_.Name) })

Write-Host "Lendo metadados de $($fotosBrutas.Count) foto(s) e $($videosBrutos.Count) video(s)..."
$metadados = [System.Collections.Generic.List[object]]::new()
$i = 0
foreach ($f in $fotosBrutas) {
    $i++
    Write-Progress -Activity 'Lendo EXIF/GPS' -Status $f.Name -PercentComplete (($i / [math]::Max(1, $fotosBrutas.Count)) * 100)
    $metadados.Add((Get-MetadadosCamera -Arquivo $f))
}
Write-Progress -Activity 'Lendo EXIF/GPS' -Completed
$videos = @($videosBrutos | ForEach-Object { Get-MetadadosCamera -Arquivo $_ })

$clusters = Get-ClustersVisita -Fotos $metadados
$visitas = Set-NomesVisita -Clusters $clusters
$avulsos = @($clusters | Where-Object { $_.Count -lt 3 })
$fv = 0; foreach ($c in $visitas) { $fv += $c.Count }
$fa = 0; foreach ($c in $avulsos) { $fa += $c.Count }

$drive = Get-DriveExterno
$destinoE = if ($drive) { "${drive}:\Sync\Camera_Celular_Organizada" } else { $null }
$pendentes = Get-PendentesSync -DestinoE $destinoE
$pendBytes = 0; foreach ($p in $pendentes) { $pendBytes += $p.Bytes }

$desde = if ($estado.ultima_execucao) { $estado.ultima_execucao } else { '(primeira execucao)' }
Write-Host "`n== ESTADO DESDE $desde ==" -ForegroundColor White
Write-Host ("  Fotos novas: {0} | Videos novos: {1}" -f $metadados.Count, $videos.Count)
Write-Host ("  Clusters: {0} | Visitas (>=3 fotos): {1} ({2} fotos) | Avulsos (1-2): {3} ({4} fotos)" -f $clusters.Count, $visitas.Count, $fv, $avulsos.Count, $fa)
Write-Host ("  Pendentes C: -> E: {0} arquivo(s), {1:N2} GB" -f $pendentes.Count, ($pendBytes / 1GB))
if ($drive) { Write-Host "  Disco externo: ${drive}: CONECTADO" -ForegroundColor Green }
else { Write-Host "  Disco externo: AUSENTE (em campo)" -ForegroundColor Yellow }

if ($WhatIfPreference) {
    Write-Host "`n== AMOSTRA DE PASTAS DE VISITA (simulacao) ==" -ForegroundColor White
    $visitas | Select-Object -First 10 | ForEach-Object {
        $c = $_
        Write-Host ("  {0}\{1}  ({2} fotos, {3}-{4})" -f $c[0].DataHora.ToString('yyyy\\MM'), $c.PastaNome, $c.Count, $c[0].DataHora.ToString('HH:mm'), $c[-1].DataHora.ToString('HH:mm'))
    }
    Write-Host ("  Videos -> Videos\{0}" -f ($videos | Select-Object -First 1 | ForEach-Object { $_.DataHora.ToString('yyyy\\MM') }))
    Write-Host "`n[SIMULACAO] $($metadados.Count) foto(s) e $($videos.Count) video(s) seriam organizados; $($pendentes.Count) pendente(s) de sync. Nada foi alterado (-WhatIf)." -ForegroundColor Cyan
    return
}

# ---- Decisao de acao ----
$organizar = $false; $sincronizar = $false; $moverE = [bool]$MoverParaExterno
if ($SemMenu -or $SincronizarExterno) {
    $organizar = $true
    $sincronizar = [bool]$SincronizarExterno
    if ($sincronizar -and -not $drive) { Write-Host "[AVISO] -SincronizarExterno ignorado: nenhum disco externo conectado." -ForegroundColor Yellow; $sincronizar = $false }
} elseif (-not $drive) {
    Write-Host "`n[CENARIO A] Disco E: nao conectado (em campo)." -ForegroundColor Yellow
    $organizar = Read-Confirmacao "Organizar localmente em $DestinoLocal? (S/N)"
} else {
    $op = Show-MenuCasa -Drive $drive
    if ($op -eq 0) {
        $organizar = $true; $sincronizar = $true
        if (-not $MoverParaExterno) { $moverE = -not (Read-Confirmacao "Manter copia no C:? (S = manter / N = MOVER para ${drive}: e liberar o SSD)") }
    } elseif ($op -eq 1) { $organizar = $true }
    elseif ($op -eq 2) {
        $sincronizar = $true
        if (-not $MoverParaExterno) { $moverE = -not (Read-Confirmacao "Manter copia no C:? (S = manter / N = MOVER para ${drive}: e liberar o SSD)") }
    } else { Write-Host "[CANCELADO] Nada foi feito." -ForegroundColor Green; return }
}
if (-not $organizar -and -not $sincronizar) { Write-Host "[OK] Nada a fazer." -ForegroundColor Green; return }

# ---- Organizacao no C: ----
$res = $null
if ($organizar) {
    $res = Invoke-OrganizarNovas -Metadados $metadados -Videos $videos -Destino $DestinoLocal
    Write-Host ("[OK] {0} foto(s) organizadas ({1} visitas, {2} avulsas) + {3} video(s) em $DestinoLocal" -f ($res.Movidos.Count), $res.Visitas, $res.Avulsos, $res.VideosMovidos.Count) -ForegroundColor Green
}

# ---- Sincronizacao C: -> E: ----
$syncFeitos = $null
if ($sincronizar -and $drive) {
    $syncFeitos = Invoke-SincronizarExterno -DestinoE $destinoE -Mover $moverE
    $sb = 0; foreach ($x in $syncFeitos) { $sb += $x.bytes }
    $acao = if ($moverE) { 'movidos' } else { 'copiados' }
    Write-Host ("[OK] {0} arquivo(s) {1} para $destinoE ({2:N2} GB)" -f $syncFeitos.Count, $acao, ($sb / 1GB)) -ForegroundColor Green
}

# ---- Estado persistente ----
$processados = @($metadados | ForEach-Object { $_.Arquivo.Name }) + @($videos | ForEach-Object { $_.Arquivo.Name })
$ultima = if ($processados.Count -gt 0) { ($processados | Sort-Object | Select-Object -Last 1) } else { $estado.ultima_foto_processada }
$fotosC = [int]$estado.total_fotos_organizadas_c; $videosC = [int]$estado.total_videos_organizados_c
if ($res) { $fotosC += $res.Movidos.Count; $videosC += $res.VideosMovidos.Count }
$pendApos = Get-PendentesSync -DestinoE $destinoE
$pendBytesApos = 0; foreach ($p in $pendApos) { $pendBytesApos += $p.Bytes }
$ts = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$novoEstado = [PSCustomObject]@{
    ultima_execucao = $ts; ultima_foto_processada = $ultima
    total_fotos_organizadas_c = $fotosC; total_videos_organizados_c = $videosC
    ultima_sincronizacao_e = $(if ($syncFeitos) { $ts } else { $estado.ultima_sincronizacao_e })
    pendentes_sync_e = [PSCustomObject]@{ arquivos = $pendApos.Count; bytes = $pendBytesApos }
}
Save-EstadoCamera -Estado $novoEstado

# ---- Manifesto de auditoria em runs\ ----
if (-not (Test-Path -LiteralPath $RunsDir)) { New-Item -ItemType Directory -Force -Path $RunsDir | Out-Null }
$runJson = Join-Path $RunsDir "${ts}__CAMERA__organizar-camera.json"
$registro = [PSCustomObject]@{
    script = 'organizar-camera'; timestamp = $ts
    origem = $OrigemCamera; destino_local = $DestinoLocal; destino_externo = $destinoE
    parametros = [PSCustomObject]@{ sem_menu = [bool]$SemMenu; sincronizar_externo = [bool]$SincronizarExterno; mover_para_externo = $moverE; copiar = [bool]$Copiar }
    novas = [PSCustomObject]@{ fotos = $metadados.Count; videos = $videos.Count }
    clusters = [PSCustomObject]@{ total = $clusters.Count; visitas = $visitas.Count; avulsos = $avulsos.Count; fotos_visitas = $fv; fotos_avulsas = $fa }
    organizacao = $(if ($res) { $res.Movidos + $res.VideosMovidos } else { @() })
    sincronizacao = [PSCustomObject]@{ executada = [bool]$syncFeitos; mover = $moverE; arquivos = $(if ($syncFeitos) { $syncFeitos.Count } else { 0 }); feitos = $(if ($syncFeitos) { $syncFeitos } else { @() }) }
    estado = $novoEstado
}
$registro | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $runJson -Encoding UTF8
Write-Host "[OK] Estado atualizado: $EstadoFile" -ForegroundColor Cyan
Write-Host "[OK] Manifesto salvo em: $runJson" -ForegroundColor Cyan
