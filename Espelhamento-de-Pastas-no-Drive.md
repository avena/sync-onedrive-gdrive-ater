# Espelhamento de Pastas no Drive

**Date**: November 17, 2026 at 01:00 AM
**Turns**: 1
**Source**: [Gemini Chat](https://gemini.google.com/app/f88c6bcdb63459de)

---

## Turn 1

### 🤖 Assistente

Perfeito! As imagens confirmam que a conta **`fernandoavena.saconsult@gmail.com`** ("Fernando sa consult") tem permissão de **pode editar** na pasta compartilhada. Como essa conta não está instalada no Google Drive Desktop do seu Windows, o **`rclone` via API (OAuth)** é exatamente a ferramenta ideal: ele autentica uma vez pelo navegador na conta da consultoria, guarda o token no seu PC e faz as transferências direto para a nuvem, sem montar disco virtual (`Z:`) e sem interferir na sua conta pessoal.
Abaixo está o guia completo em **PowerShell 7 (`pwsh`)** com a estrutura de pastas do projeto, o script interativo com menu (Simulação `--dry-run`, Sincronização Manual e Logs), a configuração opcional do Agendador de Tarefas e um prompt técnico pronto caso queira expandir ou automatizar com algum assistente de código.

### 1. Pré-requisitos e Instalação (PowerShell 7 + Rclone)
Abra o terminal do Windows (Windows Terminal ou PowerShell) e instale o **PowerShell 7** e o **Rclone** via `winget`:

```powershell
winget install --id Microsoft.PowerShell --source winget
winget install --id Rclone.Rclone --source winget

```
Após instalar, feche o terminal e abra especificamente o **PowerShell 7** (busque por `PowerShell 7` ou digite `pwsh` no terminal).
*Verificação:* Digite `$PSVersionTable.PSVersion` (deve mostrar versão `7.x`) e `rclone version` (deve exibir a versão instalada do rclone).

### 2. Autenticar a Conta `fernandoavena.saconsult@gmail.com` no Rclone
Como a pasta `VALE DO RIBEIRA` está na aba *Compartilhados comigo* da conta da consultoria, vamos criar um remoto chamado `gdrive_saconsult` usando autenticação OAuth via API:

1. No PowerShell 7, rode:
   ```powershell
   rclone config
   
   ```
2. Siga exatamente estas respostas:
  - `n` (New remote)
  - **name:** `gdrive_saconsult`
  - **Storage:** digite `drive` (Google Drive)
  - **client_id:** *(Enter em branco)*
  - **client_secret:** *(Enter em branco)*
  - **scope:** `1` (Acesso total a todos os arquivos)
  - **service_account_file:** *(Enter em branco)*
  - **Edit advanced config:** `n`
  - **Use web browser to automatically authenticate:** `y`
3. **Atenção no navegador:** Quando abrir a tela do Google, **não clique na sua conta pessoal**. Clique em **"Usar outra conta"** e entre com **`fernandoavena.saconsult@gmail.com`**, depois clique em **Permitir**.
4. De volta ao terminal:
  - **Configure this as a Shared Drive (Team Drive)?** `n`
  - **Keep this "gdrive_saconsult" remote?** `y`
  - Digite `q` para sair.
*Verificação:* Rode o comando abaixo para confirmar que o rclone enxerga a pasta compartilhada da consultoria:

```powershell
rclone lsd gdrive_saconsult: --drive-shared-with-me

```
Ele deve listar `VALE DO RIBEIRA` na tela.

### 3. Estrutura da Pasta do Projeto no Windows
Vamos criar a pasta do projeto dentro do seu diretório de usuário (`$HOME`, que equivale a `C:\Users\randomacesso`):

```powershell
$ProjetoDir = "$HOME\sync-onedrive-gdrive-ater"
New-Item -ItemType Directory -Force -Path "$ProjetoDir\logs"
Set-Location $ProjetoDir

```
A estrutura ficará limpa e organizada assim:

```plaintext
C:\Users\randomacesso\sync-onedrive-gdrive-ater\
├── start-sync.ps1    <-- Script principal com menu interativo e modo automático
└── logs\             <-- Histórico detalhado de tudo que foi copiado/atualizado

```

### 4. Script Completo em PowerShell 7 (`start-sync.ps1`)
Este script já resolve a **inversão de hierarquia** entre o seu OneDrive (`Cidade \ Atividade \ Produtor`) e o Google Drive (`Atividade \ Cidade \ Produtor`), opera exclusivamente em modo **`copy --update`** (adiciona novos e atualiza editados, **sem nunca apagar** nada que a equipe colocar no Drive) e aceita tanto o **menu manual** (para quando você voltar do campo offline) quanto execução silenciosa (`-Auto`).
Crie o arquivo `start-sync.ps1` dentro de `C:\Users\randomacesso\sync-onedrive-gdrive-ater\` com o código abaixo:

```powershell
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

```

### 5. Como Usar no Dia a Dia (Modo Manual Offline/Online)
Sempre que você voltar do campo onde trabalhou offline nos dados dos produtores e quiser subir tudo de uma vez sob seu controle:

1. Abra o **PowerShell 7** (`pwsh`).
2. Entre na pasta do projeto e inicie o script:
   ```powershell
   cd ~\sync-onedrive-gdrive-ater
   .\start-sync.ps1
   
   ```
3. Escolha a **opção `[1]` (Dry-Run)** se quiser conferir antes quais fotos/txts serão enviados sem alterar nada, ou vá direto na **opção `[2]` (Sincronizar AGORA)**.
*(Nota: Caso o Windows bloqueie a execução de scripts locais na primeira vez, libere apenas para o seu usuário rodando uma única vez: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`).*

### 6. Automatizar no Agendador de Tarefas do Windows (A cada 5 minutos)
Se além do menu manual você quiser deixar o script rodando silenciosamente a cada 5 minutos usando o parâmetro `-Auto`, execute este bloco no **PowerShell 7** (não precisa rodar como Administrador, pois cria a tarefa apenas no seu usuário):

```powershell
$Acao = New-ScheduledTaskAction `
    -Execute "pwsh.exe" `
    -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$HOME\sync-onedrive-gdrive-ater\start-sync.ps1`" -Auto"

$Gatilho = New-ScheduledTaskTrigger `
    -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$Config = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName "Sync_ATER_OneDrive_GDrive" `
    -Action $Acao `
    -Trigger $Gatilho `
    -Settings $Config `
    -Description "Copia pastas de produtores do OneDrive para o Google Drive SA Consult a cada 5 min"

```

- **Por que `-MultipleInstances IgnoreNew` é importante?** Se você tirar 300 fotos pesadas em campo e a sincronização demorar mais de 5 minutos, o Windows **não** abre uma segunda cópia do script por cima da primeira — ele espera terminar com segurança.
- **Para pausar ou remover o agendamento** (caso queira operar 100% manual): abra o *Agendador de Tarefas* (`taskschd.msc`) e desabilite `Sync_ATER_OneDrive_GDrive`, ou rode `Disable-ScheduledTask -TaskName "Sync_ATER_OneDrive_GDrive"` no PowerShell.

### 7. Dicas de Segurança e Boas Práticas

- **Proteção do Token OAuth (`rclone.conf`):** O arquivo que guarda a autorização da conta `fernandoavena.saconsult@gmail.com` fica em `$env:APPDATA\rclone\rclone.conf`. Nunca envie esse arquivo para o GitHub ou para pastas sincronizadas na nuvem.
- **Nomes exatos das pastas:** O Google Drive diferencia acentos e maiúsculas (`Cananéia` vs `Cananeia`). Antes de rodar a opção `[2]` pela primeira vez, use a **opção `[4]`** do menu para confirmar que os nomes das atividades (`ATIVIDADE INDIVIDUAL 1`, etc.) e dos municípios no Drive batem exatamente com os do seu OneDrive.
- **Limpeza automática de logs:** Como cada execução gera um `.log`, se ativar o modo de 5 minutos, lembre-se de limpar a pasta `logs\` periodicamente (ou adicionar `Get-ChildItem $LogDir -Filter *.log | Where-Object LastWriteTime -lt (Get-Date).AddDays(-7) | Remove-Item` no início do script para apagar logs com mais de 7 dias).

### 8. Prompt Pronto para Implementação / Manutenção Futura
Se você quiser colar este contexto em um assistente no seu terminal (como OpenCode ou outro agente local) para ajustar caminhos ou criar novos recursos:

```plaintext
Atue como engenheiro de automação Windows com PowerShell 7 e Rclone.
Tenho um projeto em "$HOME\sync-onedrive-gdrive-ater\start-sync.ps1" que sincroniza pastas de produtores rurais do OneDrive local para o Google Drive compartilhado da conta "fernandoavena.saconsult@gmail.com" (remoto rclone: "gdrive_saconsult").

Regras obrigatórias da arquitetura:
1. Origem local: "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\<Cidade>\<Atividade>\<Produtor>"
2. Destino nuvem: "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor>" (note a inversão entre Cidade e Atividade).
3. Usar exclusivamente `rclone copy --update --drive-shared-with-me` (nunca `sync`), pois arquivos adicionados por terceiros no Google Drive jamais podem ser apagados.
4. Manter suporte a menu interativo (opções de Dry-Run, Sync Manual, Logs e Teste de Conexão) e flag `-Auto` para execução silenciosa via Agendador de Tarefas do Windows com bloqueio de instâncias simultâneas (`IgnoreNew`).
5. Incluir rotina automática de limpeza de arquivos de log com mais de 7 dias.

```

---

*Exported from [Voyager](https://github.com/voyager-crew/voyager)*  
*Generated on November 17, 2026 at 01:00 AM*