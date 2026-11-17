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
$NomeTarefa  = "Sync_ATER_OneDrive_GDrive"
$ScriptPath  = $PSCommandPath

$Cidades = @(
    "Barra do Turvo",
    "Cananéia",
    "Jacupiranga",
    "Juquiá",
    "Pariquera-Açu",
    "Registro"
)

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

    $Args = @("copy", $Local, $Nuvem) + $RcloneFlags + @("--log-file=$LogFile")
    if ($Simulacao) { $Args += "--dry-run" }

    # Marca onde o log estava para ler depois só as linhas desta execução
    $LinhasAntes = 0
    if (Test-Path $LogFile) { $LinhasAntes = @(Get-Content $LogFile -Encoding UTF8).Count }

    & rclone @Args
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

# Seletor TUI: Cidade > Atividade > Produtor (setas + Espaço + A/D/S/Q)
function Show-ProdutoresSelector {
    $Itens = New-Object System.Collections.ArrayList

    foreach ($Cidade in $Cidades) {
        foreach ($Atividade in $Atividades) {
            $CaminhoAtiv = Join-Path $OrigemBase "$Cidade\$Atividade"
            if (-not (Test-Path $CaminhoAtiv)) { continue }
            foreach ($Prod in (Get-ChildItem -Path $CaminhoAtiv -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
                $Qtd = @(Get-ChildItem -Path $Prod.FullName -File -Recurse -ErrorAction SilentlyContinue).Count
                [void]$Itens.Add([PSCustomObject]@{
                    Selecionado = $false
                    Cidade      = $Cidade
                    Atividade   = $Atividade
                    Produtor    = $Prod.Name
                    QtdArquivos = $Qtd
                    Local       = $Prod.FullName
                    Nuvem       = "$RemotoDrive/$Atividade/$Cidade/$($Prod.Name)"
                })
            }
        }
    }

    if ($Itens.Count -eq 0) {
        Write-Host "`n[AVISO] Nenhuma pasta de produtor encontrada em $OrigemBase" -ForegroundColor Yellow
        return
    }

    $Atual  = 0
    $Offset = 0
    $Altura = 20
    try { $Altura = [Math]::Max(8, $Host.UI.RawUI.WindowSize.Height - 14) } catch { }
    try { [Console]::CursorVisible = $false } catch { }

    try {
        while ($true) {
            Clear-Host
            $Marcados = @($Itens | Where-Object { $_.Selecionado }).Count

            Write-Host ("=" * 74) -ForegroundColor Cyan
            Write-Host " SELETOR DE PRODUTORES - OneDrive -> Google Drive    Marcados: $Marcados / $($Itens.Count)" -ForegroundColor White
            Write-Host ("=" * 74) -ForegroundColor Cyan
            Write-Host " [Setas] Navegar | [Espaco] Marcar | [A] Todos/Nenhum | [D] Dry-Run | [S] Sincronizar | [Q/Esc] Voltar" -ForegroundColor Yellow
            Write-Host ("-" * 74) -ForegroundColor Cyan

            if ($Atual -lt $Offset) { $Offset = $Atual }
            if ($Atual -ge ($Offset + $Altura)) { $Offset = $Atual - $Altura + 1 }
            $Fim = [Math]::Min($Itens.Count - 1, $Offset + $Altura - 1)
            $UltCidade = ""
            $UltAtiv   = ""

            for ($i = $Offset; $i -le $Fim; $i++) {
                $Item = $Itens[$i]
                if ($Item.Cidade -ne $UltCidade -or $Item.Atividade -ne $UltAtiv) {
                    Write-Host "`n  $($Item.Cidade)  >  $($Item.Atividade)" -ForegroundColor Cyan
                    $UltCidade = $Item.Cidade
                    $UltAtiv   = $Item.Atividade
                }
                $Check = if ($Item.Selecionado) { "[X]" } else { "[ ]" }
                $Seta  = if ($i -eq $Atual) { " > " } else { "   " }
                $Linha = "$Seta$Check $($Item.Produtor)  ($($Item.QtdArquivos) arq.)"

                if ($i -eq $Atual) { Write-Host $Linha -ForegroundColor Black -BackgroundColor Gray }
                elseif ($Item.Selecionado) { Write-Host $Linha -ForegroundColor Green }
                else { Write-Host $Linha -ForegroundColor White }
            }

            Write-Host "`n$("-" * 74)" -ForegroundColor Cyan
            $Tecla = [Console]::ReadKey($true)

            switch ($Tecla.Key) {
                "UpArrow"   { if ($Atual -gt 0) { $Atual-- } }
                "DownArrow" { if ($Atual -lt ($Itens.Count - 1)) { $Atual++ } }
                "Spacebar"  { $Itens[$Atual].Selecionado = -not $Itens[$Atual].Selecionado }
                "A" {
                    $NovoEstado = ($Marcados -lt $Itens.Count)
                    foreach ($It in $Itens) { $It.Selecionado = $NovoEstado }
                }
                { $_ -in "D", "S" } {
                    $Alvos = @($Itens | Where-Object { $_.Selecionado })
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
                    foreach ($Sel in $Alvos) {
                        Write-Host "-> $($Sel.Cidade) | $($Sel.Atividade) | $($Sel.Produtor)" -ForegroundColor Green
                        $R = Invoke-RcloneCopy -Local $Sel.Local -Nuvem $Sel.Nuvem -Simulacao $Simular -LogFile $LogSel
                        $Pend += $R.Pendentes
                        if ($R.Codigo -ne 0) { $FalhasSel++ }
                        $Resumo = if ($Simular) {
                            "$($R.Pendentes) arquivo(s) pendente(s)"
                        } else {
                            "$($R.Copiados) novo(s), $($R.Atualizados) atualizado(s)"
                        }
                        if ($R.Erros -gt 0) { $Resumo += " | $($R.Erros) erro(s)" }
                        Write-Host "   -> $Resumo`n" -ForegroundColor DarkCyan
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
function Invoke-AterSync {
    param(
        [bool]$Simulacao = $false,
        [string]$ApenasCidade = ""   # vazio = todas as cidades
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
        Write-Host "       Use a opção [9] do menu (Configurar rclone)." -ForegroundColor Yellow
        return $false
    }

    # Log por execução (não por sessão): cada sync no menu gera seu próprio arquivo
    $LogFile = Join-Path $LogDir ("sync_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))

    $ModoTexto = if ($Simulacao) { "SIMULAÇÃO (DRY-RUN - NADA SERÁ ALTERADO)" } else { "SINCRONIZAÇÃO REAL (COPY)" }
    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host " INICIANDO: $ModoTexto" -ForegroundColor Cyan
    Write-Host " Horário: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "============================================================`n" -ForegroundColor Cyan

    # Monta a lista uma vez: permite contador [i/N] e evita varrer as pastas duas vezes
    $ListaCidades = if ($ApenasCidade) { @($ApenasCidade) } else { $Cidades }
    $Tarefas = @()
    foreach ($Cidade in $ListaCidades) {
        foreach ($Atividade in $Atividades) {
            # Caminho no OneDrive: ...\Fernando\<Cidade>\<Atividade>
            $CaminhoLocal = Join-Path -Path $OrigemBase -ChildPath "$Cidade\$Atividade"
            if (Test-Path $CaminhoLocal) {
                $Tarefas += [PSCustomObject]@{
                    Cidade    = $Cidade
                    Atividade = $Atividade
                    Local     = $CaminhoLocal
                    # Caminho invertido no Google Drive: ...\FERNANDO/<Atividade>/<Cidade>
                    Nuvem     = "$RemotoDrive/$Atividade/$Cidade"
                }
            }
        }
    }

    if ($Tarefas.Count -eq 0) {
        Write-Host "[AVISO] Nenhuma pasta de Cidade/Atividade encontrada no OneDrive local." -ForegroundColor Yellow
        return $true
    }

    $Indice    = 0
    $Falhas    = 0
    $Pendentes = 0
    $Erros     = 0

    foreach ($T in $Tarefas) {
        $Indice++
        Write-Host "[$Indice/$($Tarefas.Count)] $($T.Cidade) | $($T.Atividade)" -ForegroundColor Green
        Write-Host "   Local : $($T.Local)" -ForegroundColor DarkGray
        Write-Host "   Nuvem : $($T.Nuvem)" -ForegroundColor DarkGray

        $Resultado = Invoke-RcloneCopy -Local $T.Local -Nuvem $T.Nuvem -Simulacao $Simulacao -LogFile $LogFile
        $Pendentes += $Resultado.Pendentes
        $Erros     += $Resultado.Erros

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

    Write-Host "Processo finalizado ($($Tarefas.Count) pastas verificadas, $Falhas com falha)." -ForegroundColor Cyan
    Write-Host "Log salvo em: $LogFile`n" -ForegroundColor Gray

    return ($Falhas -eq 0)
}

# ================= EXECUÇÃO AUTOMÁTICA (AGENDADOR) =================
if ($Auto) {
    # Checagem sempre roda, mas sem prompt (Agendador não tem console)
    if (-not (Test-AndInstallRclone -Silent)) { exit 1 }
    # Exit code importa: Agendador marca a tarefa como falha se retornar != 0
    if (Invoke-AterSync -Simulacao $false) { exit 0 } else { exit 1 }
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
    Write-Host " Origem : $OrigemBase" -ForegroundColor Gray
    Write-Host " Destino: $RemotoDrive" -ForegroundColor Gray
    Write-Host " Modo   : COPY (--update) | Não apaga arquivos do Drive" -ForegroundColor DarkGreen
    Write-Host " Agendar: $StatusAgenda" -ForegroundColor DarkGreen
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [1] Verificar alterações (Dry-Run / simulação segura)" -ForegroundColor Yellow
    Write-Host " [2] Sincronizar AGORA (todas as cidades)" -ForegroundColor Green
    Write-Host " [3] Sincronizar UMA cidade específica" -ForegroundColor Green
    Write-Host " [4] Selecionar produtores (setas/Espaco/D/S)" -ForegroundColor White
    Write-Host " [5] Ver últimas linhas do log" -ForegroundColor White
    Write-Host " [6] Abrir pasta de Logs" -ForegroundColor White
    Write-Host " [7] Testar conexão com Google Drive" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [8] Testar rclone (versão e remotes)" -ForegroundColor Magenta
    Write-Host " [9] Configurar rclone (rclone config)" -ForegroundColor Magenta
    Write-Host "[10] Agendar tarefa (a cada 5 min)" -ForegroundColor Magenta
    Write-Host "[11] Remover agendamento" -ForegroundColor Magenta
    Write-Host "[12] Atualizar rclone" -ForegroundColor Magenta
    Write-Host " [0] Sair" -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Cyan

    $Opcao = Read-Host "Escolha uma opção"

    switch ($Opcao) {
        "1" {
            Invoke-AterSync -Simulacao $true
            Clear-KeyBuffer
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "2" {
            Invoke-AterSync -Simulacao $false
            Clear-KeyBuffer
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "3" {
            Write-Host "`nCidades disponíveis:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $Cidades.Count; $i++) { Write-Host " [$($i+1)] $($Cidades[$i])" -ForegroundColor White }
            $Escolha = Read-Host "Número da cidade (Enter cancela)"
            if ($Escolha -match '^\d+$' -and [int]$Escolha -ge 1 -and [int]$Escolha -le $Cidades.Count) {
                Invoke-AterSync -Simulacao $false -ApenasCidade $Cidades[[int]$Escolha - 1]
            } else {
                Write-Host "Cancelado." -ForegroundColor Yellow
            }
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "4" {
            if (-not (Test-Path $OrigemBase)) {
                Write-Host "`n[ERRO] Pasta base do OneDrive não encontrada: $OrigemBase" -ForegroundColor Red
                Read-Host "Pressione Enter para voltar ao menu"
            } else {
                Show-ProdutoresSelector   # volta ao menu ao apertar Q/Esc
            }
        }
        "5" {
            Show-UltimoLog
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "6" {
            if (Test-Path $LogDir) { Invoke-Item $LogDir }
            else { Write-Host "Pasta de logs ainda vazia." -ForegroundColor Yellow; Start-Sleep 2 }
        }
        "7" {
            Write-Host "`nListando pastas dentro de VALE DO RIBEIRA/FERNANDO..." -ForegroundColor Cyan
            rclone lsd "$RemotoDrive" --drive-shared-with-me
            if ($LASTEXITCODE -ne 0) {
                Write-Host "`n[ERRO] Falha ao acessar o Drive. Verifique a configuração (opção [8]/[9])." -ForegroundColor Red
            }
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "8" {
            Write-Host "`n=== STATUS DO RCLONE ===" -ForegroundColor Cyan
            rclone version
            Write-Host "`n--- Remotes configurados ---" -ForegroundColor Cyan
            rclone listremotes
            Write-Host "`n--- Config (tokens ocultos) ---" -ForegroundColor Cyan
            rclone config redacted
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "9" {
            Write-Host "`nAbrindo configurador do rclone..." -ForegroundColor Cyan
            Write-Host "  n | name: $NomeRemoto | drive | client_id: (Enter) | client_secret: (Enter)" -ForegroundColor Gray
            Write-Host "  scope: 1 | service_account: (Enter) | Edit advanced: n | shared client_id: y" -ForegroundColor Gray
            Write-Host "  browser: y (logar com fernandoavena.saconsult@gmail.com) | Shared Drive: n | Keep: y | q" -ForegroundColor Gray
            rclone config
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "10" {
            Register-SyncTask
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "11" {
            Unregister-SyncTask
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "12" {
            Update-RcloneBinario
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "0" { Write-Host "Saindo..." -ForegroundColor Gray }
        default { Write-Host "Opção inválida." -ForegroundColor Red; Start-Sleep 1 }
    }
} while ($Opcao -ne "0")