<#
.SYNOPSIS
    Sincronizador unidirecional (OneDrive Local -> Google Drive SA Consult)
.DESCRIPTION
    Copia e atualiza pastas de produtores por Município e Atividade sem apagar
    arquivos adicionados por terceiros no Google Drive.

.EXAMPLE
    .\start-sync.ps1           # menu interativo
    .\start-sync.ps1 -Auto     # execução silenciosa (Agendador de Tarefas)
#>
param(
    [switch]$Auto # Permite rodar sem menu pelo Agendador de Tarefas
)

# ================= CONFIGURAÇÕES DO PROJETO =================
$OrigemBase  = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
$NomeRemoto  = "gdrive_saconsult"
$RemotoDrive = "${NomeRemoto}:VALE DO RIBEIRA/FERNANDO"
$LogDir      = Join-Path $PSScriptRoot 'logs'
$RunsDir     = Join-Path $PSScriptRoot 'runs'   # manifestos de auditoria: NUNCA apagar
$NomeTarefa  = "Sync_ATER_OneDrive_GDrive"
$ScriptPath  = $PSCommandPath
$InboxBase   = "G:\Meu Drive"   # inbox espelhada (Drive pessoal): _processados.json compartilhado com triagem-campo.ps1

# Nomes EXATOS das pastas locais [OD] (varredura). Padrao do usuario: "Juquia" SEM acento
# (a antiga pasta "Juquiá" foi consolidada em "Juquia"). A trava - RESUMO + .txt bloqueia
# as pastas manuais antigas sem relatorio.
$Cidades = @(
    "Barra do Turvo",
    "Cananéia",
    "Jacupiranga",
    "Juquia",
    "Pariquera-Açu",
    "Registro"
)

# Nome da cidade no destino [GDSA] quando difere da pasta local [OD] (chave = nome local).
# Cidades fora deste mapa sobem com o mesmo nome local.
$CidadesRemoto = @{
    'Pariquera-Açu' = 'PARIQUERA AÇU'
    'Juquia'        = 'Juquiá'
}

$Atividades = @(
    "ATIVIDADE INDIVIDUAL 1",
    "ATIVIDADE INDIVIDUAL 2",
    "ATIVIDADE INDIVIDUAL 3",
    "ATIVIDADE INDIVIDUAL 4",
    "ATIVIDADE INDIVIDUAL 5",
    "ATIVIDADE INDIVIDUAL 6",
    "PROJETO PRODUTIVO"
)

# copy + --update: copia novos e atualiza mais recentes; NUNCA apaga nada no destino.
$RcloneFlags = @(
    "--update",
    "--exclude", "*.json",   # .json estruturado fica APENAS local/OneDrive (não sobe ao Drive da SA)
    "--transfers=4",
    "--checkers=8",
    "--drive-shared-with-me",
    "--stats-one-line",
    "--stats=2s",
    "--retries=3",          # retry nativo em falhas transitórias (internet de campo)
    "--retries-sleep=10s",
    "--contimeout=30s",
    "-v"
)

# ================= VERIFICAÇÃO E INSTALAÇÃO DO RCLONE =================
function Test-RcloneAvailable {
    if (Get-Command rclone -ErrorAction SilentlyContinue) { return $true }
    # shims do Scoop e links do winget podem não estar no PATH desta sessão
    foreach ($Dir in @("$env:USERPROFILE\scoop\shims", "$env:LOCALAPPDATA\Microsoft\WinGet\Links")) {
        if (Test-Path "$Dir\rclone.exe") {
            if ($env:PATH -notlike "*$Dir*") { $env:PATH += ";$Dir" }
            return $true
        }
    }
    return $false
}

function Test-AndInstallRclone {
    param([switch]$Silent) # -Silent: sem prompt (modo -Auto / Agendador não tem console)

    if (Test-RcloneAvailable) {
        Write-Host "[OK] rclone já está instalado: $(rclone version | Select-Object -First 1)" -ForegroundColor Green
        return $true
    }

    Write-Host "[AVISO] rclone não encontrado no sistema." -ForegroundColor Yellow

    # Preferência: scoop > winget
    $Gerenciador = if (Get-Command scoop -ErrorAction SilentlyContinue) { "scoop" } elseif (Get-Command winget -ErrorAction SilentlyContinue) { "winget" } else { $null }

    if (-not $Gerenciador) {
        Write-Host "[ERRO] Nem scoop nem winget encontrados. Instale um deles primeiro." -ForegroundColor Red
        Write-Host "  scoop:  iwr -useb get.scoop.sh | iex" -ForegroundColor Gray
        Write-Host "  winget: já vem no Windows 10/11 (App Installer)" -ForegroundColor Gray
        return $false
    }

    if (-not $Silent) {
        $Resposta = Read-Host "Deseja instalar o rclone via $Gerenciador? (S/N)"
        if ($Resposta -notin @('s','S','y','Y')) {
            Write-Host "Instalação cancelada. O script não funcionará sem rclone." -ForegroundColor Red
            return $false
        }
    } else {
        Write-Host "Instalando rclone automaticamente via $Gerenciador..." -ForegroundColor Cyan
    }

    try {
        if ($Gerenciador -eq "scoop") {
            scoop install rclone
        } else {
            winget install --id Rclone.Rclone --source winget --accept-source-agreements --accept-package-agreements
        }

        # Recarrega PATH da máquina + usuário (o teste abaixo cobre shims/links)
        $env:PATH = [Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [Environment]::GetEnvironmentVariable("Path","User")

        if (Test-RcloneAvailable) {
            Write-Host "[OK] rclone instalado com sucesso: $(rclone version | Select-Object -First 1)" -ForegroundColor Green
            return $true
        }
        Write-Host "[ERRO] Falha ao instalar rclone. Reinicie o terminal e tente novamente." -ForegroundColor Red
        return $false
    } catch {
        Write-Host "[ERRO] Exceção durante instalação: $_" -ForegroundColor Red
        return $false
    }
}

function Test-RemoteConfigurado {
    # --quiet evita NOTICE de "config file not found" poluindo o console/stdout
    return ((rclone listremotes --quiet) -contains "${NomeRemoto}:")
}

# ================= NOMES NO DESTINO E IGNORADOS =================
$IgnoradosFile = Join-Path $PSScriptRoot 'ignorados.json'

# "2026-11-07-MARIA - RESUMO" -> "07-2026-11-MARIA - RESUMO" (DD-AAAA-MM-, dia primeiro).
# O destino [GDSA] recebe o MESMO nome da pasta local [OD]. Idempotente: nome ja no formato
# novo (ou sem data) e preservado; sufixo - RESUMO garantido sem duplicar (case-insensitive).
function Get-NomeResumo {
    param([string]$Nome)
    if ($Nome -match '^(\d{4})-(\d{2})-(\d{2})-(.+)$') {
        $Nome = "$($Matches[3])-$($Matches[1])-$($Matches[2])-$($Matches[4])"
    }
    if ($Nome -match '\s*-\s*RESUMO$') { return $Nome }
    return "$Nome - RESUMO"
}

# Nome da cidade no destino [GDSA]: usa o mapa $CidadesRemoto quando existir; senao o proprio nome local
function Get-NomeCidadeRemoto {
    param([string]$Cidade)
    if ($CidadesRemoto.ContainsKey($Cidade)) { return $CidadesRemoto[$Cidade] }
    return $Cidade
}

# ================= URLs DO DRIVE + MANIFESTO runs\ =================
# Cache de listagens lsjson por pasta pai (so consulta pastas atualizadas; economia de API)
$Script:DriveDirCache = @{}

function Get-DriveFolderUrl {
    param([string]$Atividade, [string]$Cidade, [string]$NomeDrive)
    $parent = "$RemotoDrive/$Atividade/$(Get-NomeCidadeRemoto $Cidade)"
    if (-not $Script:DriveDirCache.ContainsKey($parent)) {
        $entries = @()
        try {
            $raw = & rclone lsjson $parent --dirs-only --drive-shared-with-me --quiet 2>$null
            if ($LASTEXITCODE -eq 0 -and $raw) { $entries = @(($raw -join "`n") | ConvertFrom-Json) }
        } catch { $entries = @() }
        $Script:DriveDirCache[$parent] = $entries
    }
    $entry = $Script:DriveDirCache[$parent] | Where-Object { $_.Name -eq $NomeDrive } | Select-Object -First 1
    if ($entry -and $entry.ID) { return "https://drive.google.com/drive/folders/$($entry.ID)" }
    return $null
}

# Grava manifesto em runs\ mapeando produtor -> caminho local + URL no Drive (so pastas modificadas)
function Write-SyncRun {
    param([string]$Escopo, [array]$Modificados, [bool]$Simulacao)
    if ($Simulacao) { return }
    if (-not $Modificados -or $Modificados.Count -eq 0) { return }
    if (-not (Test-Path $RunsDir)) { New-Item -ItemType Directory -Force -Path $RunsDir | Out-Null }
    $ts = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $base = Join-Path $RunsDir "${ts}__${Escopo}__start-sync"

    $itens = [System.Collections.Generic.List[object]]::new()
    foreach ($m in $Modificados) {
        $url = Get-DriveFolderUrl -Atividade $m.Atividade -Cidade $m.Cidade -NomeDrive $m.NomeDrive
        $itens.Add([PSCustomObject]@{
            produtor = $m.Produtor; cidade = $m.Cidade; atividade = $m.Atividade
            nome_drive = $m.NomeDrive; local = $m.Local
            nuvem = "$RemotoDrive/$($m.Atividade)/$($m.Cidade)/$($m.NomeDrive)"; url = $url
            copiados = $m.Copiados; atualizados = $m.Atualizados
        })
    }
    [PSCustomObject]@{ timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); escopo = $Escopo; produtores = $itens } |
        ConvertTo-Json -Depth 10 | Set-Content -Path "$base.json" -Encoding UTF8

    $linhas = [System.Collections.Generic.List[string]]::new()
    $linhas.Add("SYNC ATER - $ts - escopo $Escopo")
    $linhas.Add("Produtores com arquivos copiados/atualizados: $($itens.Count)")
    $linhas.Add("")
    foreach ($i in $itens) {
        $linhas.Add("[$($i.cidade) | $($i.atividade)] $($i.produtor) -> $($i.nome_drive)")
        $linhas.Add("  Local [OD]  : $($i.local)")
        $linhas.Add("  Drive [GDSA]: $($i.nuvem)")
        $linhas.Add("  URL [GDSA]  : $(if ($i.url) { $i.url } else { '(URL nao obtida)' })")
        $linhas.Add("  $($i.copiados) novo(s), $($i.atualizados) atualizado(s)")
        $linhas.Add("")
    }
    Set-Content -Path "$base.txt" -Value $linhas -Encoding UTF8
    Write-Host "[OK] Manifesto de URLs salvo em: $base.txt" -ForegroundColor Cyan
}

# ================= CICLO DE VIDA DAS VISITAS (_processados.json, compartilhado com triagem-campo) =================
# Marca sincronizada=true nos registros cujo destino_onedrive bate com as pastas sincronizadas
# com sucesso (Codigo 0, mesmo com 0 cópias). Falha silenciosa se G: offline; gravação atômica;
# preserva TODOS os demais campos/registros. Retorna quantos registros foram marcados.
function Update-ProcessadosSyncStatus {
    param([array]$Sincronizadas)   # objetos com Local (caminho OneDrive), Cidade, Atividade, NomeDrive
    if (-not $Sincronizadas -or @($Sincronizadas).Count -eq 0) { return 0 }
    try {
        if (-not (Test-Path -LiteralPath $InboxBase)) { return 0 }
        $alvos = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($s in $Sincronizadas) { [void]$alvos.Add(([string]$s.Local).TrimEnd('\')) }
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $total = 0
        $pathsAtualizados = [System.Collections.Generic.List[string]]::new()
        foreach ($ano in @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' })) {
            $regPath = Join-Path $ano.FullName '_inbox_relatorio_ater\_processados.json'
            if (-not (Test-Path -LiteralPath $regPath)) { continue }
            try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { continue }
            $mudou = $false
            foreach ($r in $regs) {
                $pd = $r.PSObject.Properties['destino_onedrive']
                if (-not $pd -or -not $pd.Value) { continue }
                if (-not $alvos.Contains(([string]$pd.Value).TrimEnd('\'))) { continue }
                Add-Member -InputObject $r -NotePropertyName sincronizada -NotePropertyValue $true -Force
                Add-Member -InputObject $r -NotePropertyName sincronizada_em -NotePropertyValue $ts -Force
                $pu = $r.PSObject.Properties['drive_url']
                $url = if ($pu -and $pu.Value) { $pu.Value } else {
                    try { Get-DriveFolderUrl -Atividade ([string]$r.atividade) -Cidade ([string]$r.cidade) -NomeDrive (Get-NomeResumo ([string]$r.produtor_pasta)) } catch { $null }
                }
                Add-Member -InputObject $r -NotePropertyName drive_url -NotePropertyValue $url -Force
                $mudou = $true; $total++
            }
            if ($mudou) {
                $tmp = "$regPath.tmp"
                ConvertTo-Json -InputObject @($regs) -Depth 10 | Set-Content -LiteralPath $tmp -Encoding UTF8
                Move-Item -LiteralPath $tmp -Destination $regPath -Force
                $pathsAtualizados.Add($regPath)
            }
        }
        if ($total -gt 0) {
            Write-Host "[OK] $total visita(s) marcadas como sincronizadas ([GDSA])" -ForegroundColor Cyan
            foreach ($p in $pathsAtualizados) { Write-Host "     Registro [GD]: $p" -ForegroundColor DarkGray }
        }
        return $total
    } catch { return 0 }   # G: offline/erro inesperado: nunca derruba o sync
}

# Extrai a data do prefixo da pasta ("DD-AAAA-MM-" ou "AAAA-MM-DD-") como [datetime].
# Sem prefixo de data completa -> [datetime]::MinValue (vai para o fim na ordenação decrescente).
function Get-DataPrefixoPasta {
    param([string]$Nome)
    if ($Nome -match '^(\d{2})-(\d{4})-(\d{2})-') {
        try { return [datetime]::new([int]$Matches[2], [int]$Matches[3], [int]$Matches[1]) } catch { return [datetime]::MinValue }
    }
    if ($Nome -match '^(\d{4})-(\d{2})-(\d{2})-') {
        try { return [datetime]::new([int]$Matches[1], [int]$Matches[2], [int]$Matches[3]) } catch { return [datetime]::MinValue }
    }
    return [datetime]::MinValue
}

# Conjunto (case-insensitive, sem barra final) dos destino_onedrive com sincronizada=true
# em TODOS os anos de $InboxBase. É o identificador único compartilhado com o seletor.
function Get-ProcessadosTransferidos {
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path -LiteralPath $InboxBase)) { return $set }
    foreach ($ano in @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' })) {
        $regPath = Join-Path $ano.FullName '_inbox_relatorio_ater\_processados.json'
        if (-not (Test-Path -LiteralPath $regPath)) { continue }
        try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { continue }
        foreach ($r in $regs) {
            $pd = $r.PSObject.Properties['destino_onedrive']
            $ps = $r.PSObject.Properties['sincronizada']
            if ($pd -and $pd.Value -and $ps -and $ps.Value) { [void]$set.Add(([string]$pd.Value).TrimEnd('\')) }
        }
    }
    return $set
}

# Marca/desmarca manualmente um destino_onedrive como JÁ TRANSFERIDO no _processados.json.
# - Existindo registro: atualiza sincronizada/sincronizada_em (null ao desmarcar), preservando os demais campos.
# - Não existindo e Transferido=$true: cria entrada atômica com origem_registro="manual_sync" no ano da pasta.
# - Não existindo e Transferido=$false: nada a fazer. Gravação atômica (.tmp + Move), nunca apaga outros registros.
function Set-ProcessadosTransferidoManual {
    param([string]$DestinoOD, [bool]$Transferido, [string]$Cidade = '', [string]$Atividade = '', [string]$ProdutorPasta = '')
    if (-not $DestinoOD -or -not (Test-Path -LiteralPath $InboxBase)) { return $false }
    $alvo = $DestinoOD.TrimEnd('\')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    try {
        foreach ($ano in @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' })) {
            $regPath = Join-Path $ano.FullName '_inbox_relatorio_ater\_processados.json'
            if (-not (Test-Path -LiteralPath $regPath)) { continue }
            try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { continue }
            $mudou = $false
            foreach ($r in $regs) {
                $pd = $r.PSObject.Properties['destino_onedrive']
                if (-not $pd -or -not $pd.Value) { continue }
                if (([string]$pd.Value).TrimEnd('\') -ne $alvo) { continue }
                Add-Member -InputObject $r -NotePropertyName sincronizada -NotePropertyValue $Transferido -Force
                Add-Member -InputObject $r -NotePropertyName sincronizada_em -NotePropertyValue $(if ($Transferido) { $ts } else { $null }) -Force
                $mudou = $true
            }
            if ($mudou) {
                $tmp = "$regPath.tmp"
                ConvertTo-Json -InputObject @($regs) -Depth 10 | Set-Content -LiteralPath $tmp -Encoding UTF8
                Move-Item -LiteralPath $tmp -Destination $regPath -Force
                return $true
            }
        }
        # Sem registro prévio: só cria ao MARCAR como transferido (desmarcar algo inexistente = no-op)
        if (-not $Transferido) { return $false }
        $dataPasta = Get-DataPrefixoPasta $ProdutorPasta
        $anoNovo = if ($dataPasta -gt [datetime]::MinValue) { $dataPasta.ToString('yyyy') } else { (Get-Date).ToString('yyyy') }
        $dirAno = Join-Path $InboxBase "$anoNovo\_inbox_relatorio_ater"
        if (-not (Test-Path -LiteralPath $dirAno)) { New-Item -ItemType Directory -Force -Path $dirAno | Out-Null }
        $regPath = Join-Path $dirAno '_processados.json'
        $regs = @()
        if (Test-Path -LiteralPath $regPath) { try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json | Where-Object { $_ }) } catch { $regs = @() } }
        $novo = [PSCustomObject]@{
            destino_onedrive = $DestinoOD
            cidade           = $Cidade
            atividade        = $Atividade
            produtor_pasta   = $ProdutorPasta
            sincronizada     = $true
            sincronizada_em  = $ts
            origem_registro  = 'manual_sync'
        }
        $tmp = "$regPath.tmp"
        ConvertTo-Json -InputObject @(@($regs) + $novo) -Depth 10 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $regPath -Force
        return $true
    } catch { return $false }   # G: offline/erro inesperado: nunca derruba o seletor
}

# Opção [11]: lista visitas triadas separadas entre pendentes e sincronizadas
function Show-ProcessadosStatus {
    Write-Host "`n=== STATUS DAS VISITAS TRIADAS ([GD] _processados.json -> [OD] -> [GDSA]) ===" -ForegroundColor Cyan
    if (-not (Test-Path -LiteralPath $InboxBase)) {
        Write-Host "[i] Inbox ausente/offline: $InboxBase" -ForegroundColor Yellow
        return
    }
    Write-Host " Registro [GD]: $InboxBase\<AAAA>\_inbox_relatorio_ater\_processados.json" -ForegroundColor DarkGray
    $pend = [System.Collections.Generic.List[object]]::new()
    $sync = [System.Collections.Generic.List[object]]::new()
    foreach ($ano in @(Get-ChildItem -LiteralPath $InboxBase -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' } | Sort-Object Name)) {
        $regPath = Join-Path $ano.FullName '_inbox_relatorio_ater\_processados.json'
        if (-not (Test-Path -LiteralPath $regPath)) { continue }
        try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { continue }
        foreach ($r in $regs) {
            $ps = $r.PSObject.Properties['sincronizada']
            if ($ps -and $ps.Value) { $sync.Add($r) } else { $pend.Add($r) }
        }
    }
    Write-Host "`n PENDENTES DE SINCRONIZACAO ($($pend.Count)):" -ForegroundColor Yellow
    if ($pend.Count -eq 0) { Write-Host "   (nenhuma)" -ForegroundColor DarkGray }
    foreach ($r in $pend) {
        $pm = $r.PSObject.Properties['fotos_modo']
        $modo = if ($pm -and $pm.Value) { $pm.Value } else { '-' }
        $cor = if ($modo -eq 'manual') { 'Yellow' } else { 'White' }
        $aviso = if ($modo -eq 'manual') { '  <-- fotos copiadas MANUALMENTE, conferir antes de enviar' } else { '' }
        Write-Host "   [$($r.data)] $($r.cidade) | $($r.produtor_pasta) | fotos: $modo | triada em $($r.processado_em)$aviso" -ForegroundColor $cor
        if ($r.destino_onedrive) { Write-Host "      Destino [OD]: $($r.destino_onedrive)" -ForegroundColor DarkGray }
    }
    Write-Host "`n SINCRONIZADAS ($($sync.Count)):" -ForegroundColor Green
    if ($sync.Count -eq 0) { Write-Host "   (nenhuma)" -ForegroundColor DarkGray }
    foreach ($r in $sync) {
        $pm = $r.PSObject.Properties['fotos_modo']
        $modo = if ($pm -and $pm.Value) { $pm.Value } else { '-' }
        $cor = if ($modo -eq 'manual') { 'Yellow' } else { 'White' }
        Write-Host "   [$($r.data)] $($r.cidade) | $($r.produtor_pasta) | fotos: $modo | enviada em $($r.sincronizada_em)" -ForegroundColor $cor
        if ($r.destino_onedrive) { Write-Host "      Destino [OD] : $($r.destino_onedrive)" -ForegroundColor DarkGray }
        if ($r.drive_url) { Write-Host "      URL [GDSA]   : $($r.drive_url)" -ForegroundColor DarkCyan }
    }
}

# Lista de caminhos relativos ignorados: "Cidade\Atividade\NomePastaProdutor"
function Get-Ignorados {
    if (-not (Test-Path $IgnoradosFile)) { return @() }
    try {
        return @(@(Get-Content $IgnoradosFile -Raw -Encoding UTF8 | ConvertFrom-Json) | Where-Object { $_ })
    } catch { return @() }
}

function Save-Ignorados {
    param([string[]]$Lista)
    ConvertTo-Json -InputObject @($Lista | Sort-Object -Unique) -Depth 2 |
        Set-Content -Path $IgnoradosFile -Encoding UTF8
}

# Trava anti-retroativo: só sobe pasta cujo nome termina em "- RESUMO" E que contenha
# pelo menos um .txt de relatório (coordenadas_gps.txt não conta como relatório).
# Pastas manuais antigas sem relatório nunca sobem retroativamente ao gdrive_saconsult.
function Test-TravaResumoTxt {
    param([System.IO.DirectoryInfo]$Pasta)
    if ($Pasta.Name -notmatch '\s*-\s*RESUMO$') { return $false }
    return (@(Get-ChildItem -LiteralPath $Pasta.FullName -File -Filter '*.txt' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne 'coordenadas_gps.txt' }).Count -gt 0)
}

# Uma entrada por pasta de produtor, já com o nome padronizado usado no Google Drive.
# Por padrão retorna APENAS pastas que passam na trava (sync + seletor TUI).
# -IncluirBloqueadas: devolve também as bloqueadas (flag Sincronizavel=$false) só p/ contagem informativa.
function Get-Produtores {
    param([string[]]$ListaCidades = $Cidades, [switch]$IncluirBloqueadas)
    foreach ($Cidade in $ListaCidades) {
        foreach ($Atividade in $Atividades) {
            $CaminhoAtiv = Join-Path $OrigemBase "$Cidade\$Atividade"
            if (-not (Test-Path $CaminhoAtiv)) { continue }
            foreach ($Prod in (Get-ChildItem -Path $CaminhoAtiv -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
                $Sincronizavel = Test-TravaResumoTxt $Prod
                if (-not $Sincronizavel -and -not $IncluirBloqueadas) { continue }
                $NomeDrive = Get-NomeResumo $Prod.Name
                [PSCustomObject]@{
                    Cidade    = $Cidade
                    Atividade = $Atividade
                    Produtor  = $Prod.Name
                    NomeDrive = $NomeDrive
                    Relativo  = "$Cidade\$Atividade\$($Prod.Name)"
                    Local     = $Prod.FullName
                    Nuvem     = "$RemotoDrive/$Atividade/$(Get-NomeCidadeRemoto $Cidade)/$NomeDrive"
                    Sincronizavel = $Sincronizavel
                }
            }
            # ponytail: arquivos soltos na raiz da atividade não são copiados; criar tarefa extra se existirem
        }
    }
}

# ================= AÇÕES AUXILIARES =================
function Clear-KeyBuffer {
    # Evita que teclas pendentes "vazem" para o próximo Read-Host
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch { }
}

# Executa o rclone para um par Local->Nuvem e devolve código de saída + contadores.
# Sempre usa --log-file: o -v do rclone escreve em stderr, que no PS 5.1 aparece
# como erro vermelho (NativeCommandError) se for capturado pelo pipeline.
function Invoke-RcloneCopy {
    param(
        [string]$Local,
        [string]$Nuvem,
        [bool]$Simulacao,
        [string]$LogFile
    )

    # $Args e variavel automatica do PowerShell — nunca atribuir a ela
    $RcloneArgs = @("copy", $Local, $Nuvem) + $RcloneFlags + @("--log-file=$LogFile")
    if ($Simulacao) { $RcloneArgs += "--dry-run" }

    # Marca onde o log estava para ler depois só as linhas desta execução
    $LinhasAntes = 0
    if (Test-Path $LogFile) { $LinhasAntes = @(Get-Content $LogFile -Encoding UTF8).Count }

    & rclone @RcloneArgs
    $Codigo = $LASTEXITCODE

    $Novas = @()
    if (Test-Path $LogFile) {
        $Novas = @(Get-Content $LogFile -Encoding UTF8 -ErrorAction SilentlyContinue | Select-Object -Skip $LinhasAntes)
    }

    [PSCustomObject]@{
        Codigo      = $Codigo
        Pendentes   = @($Novas | Select-String 'Skipped (copy|update)').Count
        Copiados    = @($Novas | Select-String 'Copied \(new\)').Count
        Atualizados = @($Novas | Select-String 'Copied \(replaced|Updated').Count
        Erros       = @($Novas | Select-String 'ERROR').Count
    }
}

# Seletor TUI: Cidade > Atividade > Produtor (setas + Espaço + A/D/S/T/H/Q)
# "Já transferido" é unificado no _processados.json (destino_onedrive = chave única), substituindo o [I] Ignorar.
function Show-ProdutoresSelector {
    $Transferidos = Get-ProcessadosTransferidos
    $Itens = New-Object System.Collections.ArrayList

    foreach ($P in (Get-Produtores)) {
        [void]$Itens.Add([PSCustomObject]@{
            Selecionado = $false
            Transferido = $Transferidos.Contains($P.Local.TrimEnd('\'))
            Data        = (Get-DataPrefixoPasta $P.Produtor)
            Cidade      = $P.Cidade
            Atividade   = $P.Atividade
            Produtor    = $P.Produtor
            NomeDrive   = $P.NomeDrive
            Relativo    = $P.Relativo
            QtdArquivos = @(Get-ChildItem -LiteralPath $P.Local -File -Recurse -ErrorAction SilentlyContinue).Count
            Local       = $P.Local
            Nuvem       = $P.Nuvem
        })
    }

    # Mais recentes primeiro: data do prefixo (DD-AAAA-MM-/AAAA-MM-DD-) decrescente, depois Cidade/Atividade/Nome.
    # Pastas sem data completa (Data = MinValue) vão para o fim.
    $Itens = @($Itens | Sort-Object -Property @{ Expression = 'Data'; Descending = $true }, Cidade, Atividade, Produtor)

    if ($Itens.Count -eq 0) {
        Write-Host "`n[AVISO] Nenhuma pasta de produtor encontrada em $OrigemBase" -ForegroundColor Yellow
        return
    }

    $Atual  = 0
    $Offset = 0
    $Altura = 20
    $MostrarTransferidos = $false
    try { $Altura = [Math]::Max(8, $Host.UI.RawUI.WindowSize.Height - 15) } catch { }
    try { [Console]::CursorVisible = $false } catch { }

    try {
        while ($true) {
            # Já transferidos ficam ocultos por padrão e só aparecem no fim da lista com [H]
            $Visiveis = @($Itens | Where-Object { -not $_.Transferido })
            if ($MostrarTransferidos) { $Visiveis += @($Itens | Where-Object { $_.Transferido }) }
            if ($Atual -gt ($Visiveis.Count - 1)) { $Atual = $Visiveis.Count - 1 }
            if ($Atual -lt 0) { $Atual = 0 }

            Clear-Host
            $Marcados = @($Visiveis | Where-Object { $_.Selecionado }).Count
            $QtdTransferidos = @($Itens | Where-Object { $_.Transferido }).Count

            Write-Host ("=" * 74) -ForegroundColor Cyan
            Write-Host " SELETOR DE PRODUTORES - OneDrive -> Google Drive    Marcados: $Marcados / $($Visiveis.Count)" -ForegroundColor White
            if (-not $MostrarTransferidos -and $QtdTransferidos -gt 0) {
                Write-Host " $QtdTransferidos pasta(s) oculta(s) por já transferida(s) - [H] para exibir" -ForegroundColor DarkGray
            }
            Write-Host ("=" * 74) -ForegroundColor Cyan
            Write-Host " [Setas] Navegar | [Espaco] Marcar | [A] Todos/Nenhum | [D] Dry-Run | [S] Sincronizar" -ForegroundColor Yellow
            Write-Host " [T] Marcar/desmarcar JÁ TRANSFERIDO | [H] Mostrar/Ocultar já transferidos ($QtdTransferidos) | [Q/Esc] Voltar" -ForegroundColor Yellow
            Write-Host ("-" * 74) -ForegroundColor Cyan

            if ($Atual -lt $Offset) { $Offset = $Atual }
            if ($Atual -ge ($Offset + $Altura)) { $Offset = $Atual - $Altura + 1 }
            $Fim = [Math]::Min($Visiveis.Count - 1, $Offset + $Altura - 1)
            $UltCidade = ""
            $UltAtiv   = ""

            for ($i = $Offset; $i -le $Fim; $i++) {
                $Item = $Visiveis[$i]
                if ($Item.Cidade -ne $UltCidade -or $Item.Atividade -ne $UltAtiv) {
                    Write-Host "`n  $($Item.Cidade)  >  $($Item.Atividade)" -ForegroundColor Cyan
                    $UltCidade = $Item.Cidade
                    $UltAtiv   = $Item.Atividade
                }
                $Check = if ($Item.Selecionado) { "[X]" } else { "[ ]" }
                $Seta  = if ($i -eq $Atual) { " > " } else { "   " }
                $Marca = if ($Item.Transferido) { "[TRANSFERIDO] " } else { "" }
                $Linha = "$Seta$Check $Marca$($Item.Produtor) -> $($Item.NomeDrive)  ($($Item.QtdArquivos) arq.)"

                if ($i -eq $Atual) { Write-Host $Linha -ForegroundColor Black -BackgroundColor Gray }
                elseif ($Item.Transferido) { Write-Host $Linha -ForegroundColor DarkGray }
                elseif ($Item.Selecionado) { Write-Host $Linha -ForegroundColor Green }
                else { Write-Host $Linha -ForegroundColor White }
            }

            Write-Host "`n$("-" * 74)" -ForegroundColor Cyan
            $Tecla = [Console]::ReadKey($true)

            switch ($Tecla.Key) {
                "UpArrow"   { if ($Atual -gt 0) { $Atual-- } }
                "DownArrow" { if ($Atual -lt ($Visiveis.Count - 1)) { $Atual++ } }
                "Spacebar"  { if ($Visiveis.Count -gt 0) { $Visiveis[$Atual].Selecionado = -not $Visiveis[$Atual].Selecionado } }
                "A" {
                    $PendVis = @($Visiveis | Where-Object { -not $_.Transferido })
                    $NovoEstado = (@($PendVis | Where-Object { $_.Selecionado }).Count -lt $PendVis.Count)
                    foreach ($It in $PendVis) { $It.Selecionado = $NovoEstado }
                }
                "H" { $MostrarTransferidos = -not $MostrarTransferidos; $Atual = 0; $Offset = 0 }
                "T" {
                    if ($Visiveis.Count -eq 0) { continue }
                    $Item = $Visiveis[$Atual]
                    $novo = -not $Item.Transferido
                    $ok = Set-ProcessadosTransferidoManual -DestinoOD $Item.Local -Transferido $novo -Cidade $Item.Cidade -Atividade $Item.Atividade -ProdutorPasta $Item.Produtor
                    if (-not $ok) {
                        Write-Host "`n[!] Não foi possível gravar no _processados.json (G: offline?)." -ForegroundColor Red
                        Start-Sleep -Milliseconds 1200
                        continue
                    }
                    $Item.Transferido = $novo
                    if ($novo) {
                        $Item.Selecionado = $false
                        Write-Host "`n[OK] '$($Item.Produtor)' marcado como JÁ TRANSFERIDO." -ForegroundColor DarkGray
                    } else {
                        Write-Host "`n[OK] '$($Item.Produtor)' voltou para a lista de pendentes." -ForegroundColor Green
                    }
                    Start-Sleep -Milliseconds 900
                }
                { $_ -in "D", "S" } {
                    $Alvos = @($Visiveis | Where-Object { $_.Selecionado })
                    if ($Alvos.Count -eq 0) {
                        Write-Host "`n[!] Marque pelo menos um produtor com a barra de Espaco." -ForegroundColor Red
                        Start-Sleep -Milliseconds 1200
                        continue
                    }

                    $Simular = ($Tecla.Key -eq "D")
                    $TituloAcao = if ($Simular) { "DRY-RUN (SIMULACAO)" } else { "SINCRONIZACAO REAL (COPY)" }
                    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }
                    $LogSel = Join-Path $LogDir ("sync_seletivo_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))

                    try { [Console]::CursorVisible = $true } catch { }
                    Clear-Host
                    Write-Host "=== $TituloAcao - $($Alvos.Count) produtor(es) ===`n" -ForegroundColor Cyan

                    $Pend = 0
                    $FalhasSel = 0
                    $ModSel = [System.Collections.Generic.List[object]]::new()
                    $SyncOkSel = [System.Collections.Generic.List[object]]::new()
                    foreach ($Sel in $Alvos) {
                        Write-Host "-> $($Sel.Cidade) | $($Sel.Atividade) | $($Sel.Produtor)" -ForegroundColor Green
                        Write-Host "   Origem [OD]  : $($Sel.Local)" -ForegroundColor DarkGray
                        Write-Host "   Destino [GDSA]: $($Sel.Nuvem)" -ForegroundColor DarkGray
                        $R = Invoke-RcloneCopy -Local $Sel.Local -Nuvem $Sel.Nuvem -Simulacao $Simular -LogFile $LogSel
                        $Pend += $R.Pendentes
                        if ($R.Codigo -ne 0) { $FalhasSel++ }
                        elseif (-not $Simular) {
                            $SyncOkSel.Add([PSCustomObject]@{ Local = $Sel.Local; Cidade = $Sel.Cidade; Atividade = $Sel.Atividade; NomeDrive = $Sel.NomeDrive })
                        }
                        if (-not $Simular -and ($R.Copiados + $R.Atualizados) -gt 0) {
                            $ModSel.Add([PSCustomObject]@{
                                Produtor = $Sel.Produtor; Cidade = $Sel.Cidade; Atividade = $Sel.Atividade
                                NomeDrive = $Sel.NomeDrive; Local = $Sel.Local
                                Copiados = $R.Copiados; Atualizados = $R.Atualizados
                            })
                        }
                        $Resumo = if ($Simular) {
                            "$($R.Pendentes) arquivo(s) pendente(s)"
                        } else {
                            "$($R.Copiados) novo(s), $($R.Atualizados) atualizado(s)"
                        }
                        if ($R.Erros -gt 0) { $Resumo += " | $($R.Erros) erro(s)" }
                        Write-Host "   -> $Resumo`n" -ForegroundColor DarkCyan
                    }
                    Write-SyncRun -Escopo 'SELETOR' -Modificados $ModSel -Simulacao $Simular
                    if (-not $Simular) {
                        $null = Update-ProcessadosSyncStatus -Sincronizadas $SyncOkSel   # exibe [OK] + Registro [GD] internamente
                        # Reflete imediatamente na tela ao voltar: itens sincronizados ficam Transferido=true
                        foreach ($okItem in $SyncOkSel) {
                            $match = $Itens | Where-Object { $_.Local.TrimEnd('\') -eq $okItem.Local.TrimEnd('\') } | Select-Object -First 1
                            if ($match) { $match.Transferido = $true; $match.Selecionado = $false }
                        }
                    }

                    if ($Simular) {
                        if ($Pend -eq 0 -and $FalhasSel -eq 0) {
                            Write-Host "[OK] Nenhuma alteracao pendente nos produtores marcados." -ForegroundColor Green
                        } else {
                            Write-Host "$Pend arquivo(s) seriam copiados/atualizados." -ForegroundColor Yellow
                        }
                    }
                    Write-Host "Concluido ($($Alvos.Count) produtor(es), $FalhasSel com falha). Log: $LogSel" -ForegroundColor Cyan

                    Read-Host "Pressione Enter para voltar a selecao"
                    Clear-KeyBuffer
                    try { [Console]::CursorVisible = $false } catch { }
                }
                { $_ -in "Q", "Escape" } { return }
            }
        }
    } finally {
        try { [Console]::CursorVisible = $true } catch { }
        Clear-KeyBuffer
        Clear-Host
    }
}

function Show-UltimoLog {
    $Ultimo = Get-ChildItem $LogDir -Filter *.log -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $Ultimo) { Write-Host "`nNenhum log gerado ainda." -ForegroundColor Yellow; return }
    Write-Host "`n=== $($Ultimo.Name) (últimas 30 linhas) ===" -ForegroundColor Cyan
    Get-Content $Ultimo.FullName -Tail 30
}

function Register-SyncTask {
    # pwsh.exe se existir; senão usa o Windows PowerShell (não falha por ausência do PS7)
    $Exe = if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { "pwsh.exe" } else { "powershell.exe" }

    try {
        $Acao = New-ScheduledTaskAction -Execute $Exe `
            -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Auto"
        $Gatilho = New-ScheduledTaskTrigger -Once -At (Get-Date) `
            -RepetitionInterval (New-TimeSpan -Minutes 5) `
            -RepetitionDuration (New-TimeSpan -Days 3650)
        $Config = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $NomeTarefa -Action $Acao -Trigger $Gatilho -Settings $Config -Force `
            -Description "Copia pastas de produtores do OneDrive para o Google Drive SA Consult a cada 5 min" | Out-Null

        Write-Host "[OK] Tarefa '$NomeTarefa' agendada a cada 5 min usando $Exe." -ForegroundColor Green
    } catch {
        Write-Host "[ERRO] Falha ao agendar: $_" -ForegroundColor Red
    }
}

function Unregister-SyncTask {
    if (-not (Get-ScheduledTask -TaskName $NomeTarefa -ErrorAction SilentlyContinue)) {
        Write-Host "Nenhum agendamento com o nome '$NomeTarefa'." -ForegroundColor Yellow
        return
    }
    try {
        Unregister-ScheduledTask -TaskName $NomeTarefa -Confirm:$false -ErrorAction Stop
        Write-Host "[OK] Agendamento removido." -ForegroundColor Green
    } catch {
        Write-Host "[ERRO] Falha ao remover: $_" -ForegroundColor Red
    }
}

function Update-RcloneBinario {
    if (Get-Command scoop -ErrorAction SilentlyContinue) {
        scoop update rclone
    } elseif (Get-Command winget -ErrorAction SilentlyContinue) {
        winget upgrade --id Rclone.Rclone --source winget --accept-source-agreements --accept-package-agreements
    } else {
        Write-Host "Nenhum gerenciador de pacotes disponível." -ForegroundColor Red
    }
    Write-Host "Versão atual: $(rclone version | Select-Object -First 1)" -ForegroundColor Cyan
}

# ================= FUNÇÃO DE SINCRONIZAÇÃO =================
# Envio em LOTE (todas as cidades ou uma cidade via -ApenasCidade). Foi REMOVIDO do menu interativo
# (antigas opções [2]/[3]) para priorizar o Seletor de Produtores, mas a função e o parâmetro -Auto
# permanecem INTACTOS: o lote fica reservado para a execução automática (Agendador de Tarefas) e
# para uso futuro. O dry-run do menu [1] ainda a utiliza com -Simulacao $true.
function Invoke-AterSync {
    param(
        [bool]$Simulacao = $false,
        [string]$ApenasCidade = "",   # vazio = todas as cidades
        [string]$Escopo = "TODAS"     # rotulo do manifesto em runs\
    )

    if (-not (Test-Path $OrigemBase)) {
        Write-Host "`n[ERRO] Pasta base do OneDrive não encontrada:" -ForegroundColor Red
        Write-Host "       $OrigemBase" -ForegroundColor Yellow
        return $false
    }

    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
    }

    # Retenção: apaga logs com mais de 7 dias
    Get-ChildItem $LogDir -Filter *.log -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt (Get-Date).AddDays(-7) | Remove-Item -Force

    # Pré-checagem: remote configurado? Evita 15 erros CRITICAL idênticos
    if (-not (Test-RemoteConfigurado)) {
        Write-Host "`n[ERRO] Remote '$NomeRemoto' não configurado no rclone." -ForegroundColor Red
        Write-Host "       Use a opção [7] do menu (Configurar rclone)." -ForegroundColor Yellow
        return $false
    }

    # Log por execução (não por sessão): cada sync no menu gera seu próprio arquivo
    $LogFile = Join-Path $LogDir ("sync_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))

    $ModoTexto = if ($Simulacao) { "SIMULAÇÃO (DRY-RUN - NADA SERÁ ALTERADO)" } else { "SINCRONIZAÇÃO REAL (COPY)" }
    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host " INICIANDO: $ModoTexto" -ForegroundColor Cyan
    Write-Host " Horário: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host " Origem [OD]  : $OrigemBase" -ForegroundColor Gray
    Write-Host " Destino [GDSA]: $RemotoDrive" -ForegroundColor Gray
    Write-Host "============================================================`n" -ForegroundColor Cyan

    # Uma tarefa por pasta de produtor (nome padronizado no destino), já sem as ignoradas.
    # A trava - RESUMO + .txt bloqueia pastas antigas/sem relatório (nunca sobem retroativamente).
    $ListaCidades = if ($ApenasCidade) { @($ApenasCidade) } else { $Cidades }
    $Ignorados = Get-Ignorados
    $Todos = @(Get-Produtores -ListaCidades $ListaCidades -IncluirBloqueadas)
    $BloqueadasTrava = @($Todos | Where-Object { -not $_.Sincronizavel })
    $Elegiveis = @($Todos | Where-Object { $_.Sincronizavel })
    $Tarefas = @($Elegiveis | Where-Object { $Ignorados -notcontains $_.Relativo })
    $QtdIgnorados = $Elegiveis.Count - $Tarefas.Count

    if ($BloqueadasTrava.Count -gt 0) {
        Write-Host "[INFO] $($BloqueadasTrava.Count) pasta(s) antiga(s)/sem relatório ignorada(s) pela trava - RESUMO + .txt." -ForegroundColor DarkGray
    }
    if ($QtdIgnorados -gt 0) {
        Write-Host "[INFO] $QtdIgnorados pasta(s) ignorada(s) via ignorados.json.`n" -ForegroundColor DarkYellow
    }

    if ($Tarefas.Count -eq 0) {
        Write-Host "[AVISO] Nenhuma pasta de produtor encontrada no OneDrive local." -ForegroundColor Yellow
        return $true
    }

    $Indice    = 0
    $Falhas    = 0
    $Pendentes = 0
    $Erros     = 0
    $Modificados = [System.Collections.Generic.List[object]]::new()
    $SyncOk = [System.Collections.Generic.List[object]]::new()   # pastas com Codigo 0 (p/ _processados.json)

    foreach ($T in $Tarefas) {
        $Indice++
        Write-Host "[$Indice/$($Tarefas.Count)] $($T.Cidade) | $($T.Atividade) | $($T.Produtor)" -ForegroundColor Green
        Write-Host "   Origem [OD]  : $($T.Local)" -ForegroundColor DarkGray
        Write-Host "   Destino [GDSA]: $($T.Nuvem)" -ForegroundColor DarkGray

        $Resultado = Invoke-RcloneCopy -Local $T.Local -Nuvem $T.Nuvem -Simulacao $Simulacao -LogFile $LogFile
        $Pendentes += $Resultado.Pendentes
        $Erros     += $Resultado.Erros
        if (-not $Simulacao -and $Resultado.Codigo -eq 0) {
            $SyncOk.Add([PSCustomObject]@{ Local = $T.Local; Cidade = $T.Cidade; Atividade = $T.Atividade; NomeDrive = $T.NomeDrive })
        }

        # So pastas com arquivos efetivamente copiados/atualizados entram no manifesto de URLs
        if (-not $Simulacao -and ($Resultado.Copiados + $Resultado.Atualizados) -gt 0) {
            $Modificados.Add([PSCustomObject]@{
                Produtor = $T.Produtor; Cidade = $T.Cidade; Atividade = $T.Atividade
                NomeDrive = $T.NomeDrive; Local = $T.Local
                Copiados = $Resultado.Copiados; Atualizados = $Resultado.Atualizados
            })
        }

        $Resumo = if ($Simulacao) {
            "$($Resultado.Pendentes) arquivo(s) pendente(s)"
        } else {
            "$($Resultado.Copiados) novo(s), $($Resultado.Atualizados) atualizado(s)"
        }
        if ($Resultado.Erros -gt 0) { $Resumo += " | $($Resultado.Erros) erro(s)" }
        Write-Host "   -> $Resumo" -ForegroundColor DarkCyan

        if ($Resultado.Codigo -eq 0) {
            Write-Host "   Status: Concluído.`n" -ForegroundColor Cyan
        } else {
            $Falhas++
            Write-Host "   Status: Atenção - rclone retornou $($Resultado.Codigo).`n" -ForegroundColor Red
        }
    }

    if ($Simulacao) {
        Write-Host "--- RESUMO DO DRY-RUN ---" -ForegroundColor Cyan
        if ($Pendentes -eq 0 -and $Erros -eq 0) {
            Write-Host "[OK] Tudo sincronizado! Nenhuma alteração pendente encontrada." -ForegroundColor Green
        } else {
            Write-Host "$Pendentes arquivo(s) seriam copiados/atualizados | $Erros erro(s) no log." -ForegroundColor Yellow
        }
    }

    Write-Host "Processo finalizado ($($Tarefas.Count) pastas verificadas, $QtdIgnorados ignoradas, $Falhas com falha)." -ForegroundColor Cyan
    Write-Host "Log salvo em: $LogFile`n" -ForegroundColor Gray

    Write-SyncRun -Escopo $Escopo -Modificados $Modificados -Simulacao $Simulacao

    # Ciclo de vida: marca sincronizada=true no _processados.json da triagem (nunca em dry-run)
    if (-not $Simulacao) {
        $null = Update-ProcessadosSyncStatus -Sincronizadas $SyncOk   # exibe [OK] + Registro [GD] internamente
    }

    return ($Falhas -eq 0)
}

# ================= EXECUÇÃO AUTOMÁTICA (AGENDADOR) =================
if ($Auto) {
    # Checagem sempre roda, mas sem prompt (Agendador não tem console)
    if (-not (Test-AndInstallRclone -Silent)) { exit 1 }
    # Exit code importa: Agendador marca a tarefa como falha se retornar != 0
    if (Invoke-AterSync -Simulacao $false -Escopo 'AUTO') { exit 0 } else { exit 1 }
}

# ================= CHECAGEM INICIAL (MODO MANUAL) =================
if (-not (Test-AndInstallRclone)) {
    Read-Host "Pressione Enter para sair"
    exit 1
}

# ================= MENU INTERATIVO =================
do {
    $Agendado = Get-ScheduledTask -TaskName $NomeTarefa -ErrorAction SilentlyContinue
    $StatusAgenda = if ($Agendado) { "ATIVO (5 min)" } else { "desativado" }

    Clear-Host
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "   SYNC ATER - ONEDRIVE -> GOOGLE DRIVE (SA CONSULTORIA)    " -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " Origem [OD]  : $OrigemBase" -ForegroundColor Gray
    Write-Host " Destino [GDSA]: $RemotoDrive" -ForegroundColor Gray
    Write-Host " Modo   : COPY (--update) | Não apaga arquivos do Drive" -ForegroundColor DarkGreen
    Write-Host " Agendar: $StatusAgenda" -ForegroundColor DarkGreen
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [1] Verificar alterações (Dry-Run / simulação segura)" -ForegroundColor Yellow
    Write-Host " [2] Selecionar produtores (setas/Espaco/D/S/T/H)" -ForegroundColor White
    Write-Host " [3] Ver últimas linhas do log" -ForegroundColor White
    Write-Host " [4] Abrir pasta de Logs" -ForegroundColor White
    Write-Host " [5] Testar conexão com [GDSA] (Google Drive SA Consultoria)" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [6] Testar rclone (versão e remotes)" -ForegroundColor Magenta
    Write-Host " [7] Configurar rclone (rclone config)" -ForegroundColor Magenta
    Write-Host " [8] Agendar tarefa (a cada 5 min)" -ForegroundColor Magenta
    Write-Host " [9] Remover agendamento" -ForegroundColor Magenta
    Write-Host "[10] Atualizar rclone" -ForegroundColor Magenta
    Write-Host "[11] Status de visitas ([GD] _processados.json -> [OD] -> [GDSA])" -ForegroundColor Magenta
    Write-Host " [0] Sair" -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Cyan

    Clear-KeyBuffer   # teclas digitadas durante o sync nao vazam para a escolha do menu
    $Opcao = Read-Host "Escolha uma opção"

    switch ($Opcao) {
        "1" {
            Invoke-AterSync -Simulacao $true
            Clear-KeyBuffer
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "2" {
            if (-not (Test-Path $OrigemBase)) {
                Write-Host "`n[ERRO] Pasta base do OneDrive não encontrada: $OrigemBase" -ForegroundColor Red
                Read-Host "Pressione Enter para voltar ao menu"
            } else {
                Show-ProdutoresSelector   # volta ao menu ao apertar Q/Esc
            }
        }
        "3" {
            Show-UltimoLog
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "4" {
            if (Test-Path $LogDir) { Invoke-Item $LogDir }
            else { Write-Host "Pasta de logs ainda vazia." -ForegroundColor Yellow; Start-Sleep 2 }
        }
        "5" {
            Write-Host "`nListando pastas dentro de VALE DO RIBEIRA/FERNANDO..." -ForegroundColor Cyan
            rclone lsd "$RemotoDrive" --drive-shared-with-me
            if ($LASTEXITCODE -ne 0) {
                Write-Host "`n[ERRO] Falha ao acessar o Drive. Verifique a configuração (opção [6]/[7])." -ForegroundColor Red
            }
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "6" {
            Write-Host "`n=== STATUS DO RCLONE ===" -ForegroundColor Cyan
            rclone version
            Write-Host "`n--- Remotes configurados ---" -ForegroundColor Cyan
            rclone listremotes
            Write-Host "`n--- Config (tokens ocultos) ---" -ForegroundColor Cyan
            rclone config redacted
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "7" {
            Write-Host "`nAbrindo configurador do rclone..." -ForegroundColor Cyan
            Write-Host "  n | name: $NomeRemoto | drive | client_id: (Enter) | client_secret: (Enter)" -ForegroundColor Gray
            Write-Host "  scope: 1 | service_account: (Enter) | Edit advanced: n | shared client_id: y" -ForegroundColor Gray
            Write-Host "  browser: y (logar com fernandoavena.saconsult@gmail.com) | Shared Drive: n | Keep: y | q" -ForegroundColor Gray
            rclone config
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "8" {
            Register-SyncTask
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "9" {
            Unregister-SyncTask
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "10" {
            Update-RcloneBinario
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "11" {
            Show-ProcessadosStatus
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "0" { Write-Host "Saindo..." -ForegroundColor Gray }
        default { Write-Host "Opção inválida." -ForegroundColor Red; Start-Sleep 1 }
    }
} while ($Opcao -ne "0")