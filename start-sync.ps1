<#
.SYNOPSIS
    Sincronizador unidirecional (OneDrive Local -> Google Drive SA Consult)
.DESCRIPTION
    Copia e atualiza pastas de produtores por Município e Atividade sem apagar
    arquivos adicionados por terceiros no Google Drive.
#>
param(
    [switch]$Auto # Permite rodar sem menu pelo Agendador de Tarefas
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

        # Recarrega PATH da máquina + usuário (teste abaixo cobre shims/links)
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

# Checagem sempre roda (inclusive -Auto), mas sem prompt no Agendador
if (-not (Test-AndInstallRclone -Silent:$Auto)) {
    if (-not $Auto) { Read-Host "Pressione Enter para sair" }
    exit 1
}

# ================= CONFIGURAÇÕES DO PROJETO =================
$OrigemBase  = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
$RemotoDrive = "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO"
$LogDir      = "$PSScriptRoot\logs"
$Timestamp   = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogFile     = "$LogDir\sync_$Timestamp.log"

# Mapeamento: Cidades e Atividades ativas
# O script cruza cada Cidade com cada Atividade automaticamente se a pasta existir no seu OneDrive
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

# Flags do Rclone:
# copy + --update: copia novos e atualiza mais recentes; NUNCA apaga nada no destino.
# --drive-shared-with-me: acessa a pasta compartilhada com fernandoavena.saconsult.
$RcloneFlags = @(
    "--update",
    "--transfers=4",
    "--checkers=8",
    "--drive-shared-with-me",
    "--stats-one-line",
    "--stats=2s",
    "-v"
)

# ================= FUNÇÃO DE EXECUÇÃO =================
function Invoke-AterSync {
    param([bool]$Simulacao = $false)

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
    $NomeRemoto = ($RemotoDrive -split ':')[0]
    if ((rclone listremotes) -notcontains "$NomeRemoto`:") {
        Write-Host "`n[ERRO] Remote '$NomeRemoto' não configurado no rclone." -ForegroundColor Red
        Write-Host "       Use a opção [5] do menu (rclone config) para configurar." -ForegroundColor Yellow
        return $false
    }

    $ModoTexto = if ($Simulacao) { "SIMULAÇÃO (DRY-RUN - NADA SERÁ ALTERADO)" } else { "SINCRONIZAÇÃO REAL (COPY)" }
    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host " INICIANDO: $ModoTexto" -ForegroundColor Cyan
    Write-Host " Horário: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "============================================================`n" -ForegroundColor Cyan

    $ParesEncontrados = 0
    $ParesComFalha    = 0

    foreach ($Cidade in $Cidades) {
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
    # Exit code importa: Agendador marca a tarefa como falha se retornar != 0
    if (Invoke-AterSync -Simulacao $false) { exit 0 } else { exit 1 }
}

# ================= MENU INTERATIVO (EXECUÇÃO MANUAL) =================
do {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "   SYNC ATER - ONEDRIVE -> GOOGLE DRIVE (SA CONSULTORIA)    " -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " Origem : $OrigemBase" -ForegroundColor Gray
    Write-Host " Destino: $RemotoDrive" -ForegroundColor Gray
    Write-Host " Modo   : COPY (--update) | Não apaga arquivos do Drive" -ForegroundColor DarkGreen
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [1] Verificar alterações (Dry-Run / Simulação segura)" -ForegroundColor Yellow
    Write-Host " [2] Sincronizar AGORA (Enviar novos e atualizados)" -ForegroundColor Green
    Write-Host " [3] Abrir pasta de Logs" -ForegroundColor White
    Write-Host " [4] Testar conexão com Google Drive (SA Consult)" -ForegroundColor White
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " [5] Configurar rclone (rclone config)" -ForegroundColor Magenta
    Write-Host " [6] Ver status do rclone (versão e remotes)" -ForegroundColor Magenta
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
            if (Test-Path $LogDir) { Invoke-Item $LogDir }
            else { Write-Host "Pasta de logs ainda vazia." -ForegroundColor Yellow; Start-Sleep 2 }
        }
        "4" {
            Write-Host "`nListando pastas dentro de VALE DO RIBEIRA/FERNANDO..." -ForegroundColor Cyan
            rclone lsd "$RemotoDrive" --drive-shared-with-me
            Read-Host "`nPressione Enter para voltar ao menu"
        }
        "5" {
            Write-Host "`nAbrindo configurador do rclone..." -ForegroundColor Cyan
            Write-Host "Siga os passos: n → name: gdrive_saconsult → drive → scope: 1 → y (browser) → conta da consultoria → n → y → q" -ForegroundColor Gray
            rclone config
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
        "0" { Write-Host "Saindo..." -ForegroundColor Gray }
        default { Write-Host "Opção inválida." -ForegroundColor Red; Start-Sleep 1 }
    }
} while ($Opcao -ne "0")