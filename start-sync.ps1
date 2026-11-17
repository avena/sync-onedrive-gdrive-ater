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
    "ATIVIDADE INDIVIDUAL 3"
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
function Show-PastasLocais {
    Write-Host "`nPastas encontradas em $OrigemBase" -ForegroundColor Cyan
    $Total = 0
    foreach ($Cidade in $Cidades) {
        foreach ($Atividade in $Atividades) {
            if (Test-Path (Join-Path $OrigemBase "$Cidade\$Atividade")) {
                $Total++
                Write-Host "  [OK] $Cidade | $Atividade" -ForegroundColor Green
            }
        }
    }
    if ($Total -eq 0) { Write-Host "  Nenhuma pasta Cidade\Atividade encontrada." -ForegroundColor Yellow }
    else { Write-Host "`n  $Total pasta(s) prontas para sincronizar." -ForegroundColor Cyan }
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

    $ParesEncontrados = 0
    $ParesComFalha    = 0

    $ListaCidades = if ($ApenasCidade) { @($ApenasCidade) } else { $Cidades }
    foreach ($Cidade in $ListaCidades) {
        foreach ($Atividade in $Atividades) {
            # Caminho no OneDrive: ...\Fernando\<Cidade>\<Atividade>
            $CaminhoLocal = Join-Path -Path $OrigemBase -ChildPath "$Cidade\$Atividade"

            if (Test-Path $CaminhoLocal) {
                $ParesEncontrados++
                # Caminho invertido no Google Drive: ...\FERNANDO/<Atividade>/<Cidade>
                $CaminhoNuvem = "$RemotoDrive/$Atividade/$Cidade"

                Write-Host "-> [$Cidade] | [$Atividade]" -ForegroundColor Green
                Write-Host "   Local : $CaminhoLocal" -ForegroundColor DarkGray
                Write-Host "   Nuvem : $CaminhoNuvem" -ForegroundColor DarkGray

                $ArgsRclone = @("copy", $CaminhoLocal, $CaminhoNuvem) + $RcloneFlags + @("--log-file=$LogFile")
                if ($Simulacao) {
                    $ArgsRclone += "--dry-run"
                }

                & rclone @ArgsRclone

                if ($LASTEXITCODE -eq 0) {
                    Write-Host "   Status: Concluído com sucesso.`n" -ForegroundColor Cyan
                } else {
                    $ParesComFalha++
                    Write-Host "   Status: Atenção - Rclone retornou código $LASTEXITCODE. Verifique o log.`n" -ForegroundColor Red
                }
            }
        }
    }

    if ($ParesEncontrados -eq 0) {
        Write-Host "[AVISO] Nenhuma pasta de Cidade/Atividade encontrada no OneDrive local." -ForegroundColor Yellow
    } else {
        Write-Host "Processo finalizado ($ParesEncontrados pastas verificadas, $ParesComFalha com falha)." -ForegroundColor Cyan
        Write-Host "Log salvo em: $LogFile`n" -ForegroundColor Gray
    }

    return ($ParesComFalha -eq 0)
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
    Write-Host " [4] Mostrar pastas locais encontradas" -ForegroundColor White
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
            Read-Host "Pressione Enter para voltar ao menu"
        }
        "2" {
            Invoke-AterSync -Simulacao $false
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
            Show-PastasLocais
            Read-Host "`nPressione Enter para voltar ao menu"
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
            Write-Host "Passos: n | name: $NomeRemoto | drive | scope: 1 | y (navegador) | conta da consultoria | n | y | q" -ForegroundColor Gray
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