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
        return
    }

    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
    }

    $ModoTexto = if ($Simulacao) { "SIMULAÇÃO (DRY-RUN - NADA SERÁ ALTERADO)" } else { "SINCRONIZAÇÃO REAL (COPY)" }
    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host " INICIANDO: $ModoTexto" -ForegroundColor Cyan
    Write-Host " Horário: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "============================================================`n" -ForegroundColor Cyan

    $ParesEncontrados = 0

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
                    Write-Host "   Status: Atenção - Rclone retornou código $LASTEXITCODE. Verifique o log.`n" -ForegroundColor Red
                }
            }
        }
    }

    if ($ParesEncontrados -eq 0) {
        Write-Host "[AVISO] Nenhuma pasta de Cidade/Atividade encontrada no OneDrive local." -ForegroundColor Yellow
    } else {
        Write-Host "Processo finalizado ($ParesEncontrados pastas verificadas)." -ForegroundColor Cyan
        Write-Host "Log salvo em: $LogFile`n" -ForegroundColor Gray
    }
}

# ================= EXECUÇÃO AUTOMÁTICA (AGENDADOR) =================
if ($Auto) {
    Invoke-AterSync -Simulacao $false
    exit 0
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
        "0" { Write-Host "Saindo..." -ForegroundColor Gray }
        default { Write-Host "Opção inválida." -ForegroundColor Red; Start-Sleep 1 }
    }
} while ($Opcao -ne "0")