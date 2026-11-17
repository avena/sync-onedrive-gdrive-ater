# Sync ATER - OneDrive → Google Drive (SA Consultoria)

Sincronizador unidirecional para copiar pastas de produtores rurais do **OneDrive local** para o **Google Drive compartilhado** da conta `fernandoavena.saconsult@gmail.com`.

## 📋 Visão Geral

- **Origem**: `C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\<Cidade>\<Atividade>`
- **Destino**: `gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>` *(hierarquia invertida)*
- **Modo**: `rclone copy --update` — **nunca apaga** arquivos no destino (protege alterações da equipe)
- **Autenticação**: OAuth via `rclone config` (remote `gdrive_saconsult`)

## 🚀 Como Executar

### Pré-requisitos (uma vez)
```powershell
# PowerShell 7 (pwsh) - já incluso no Windows 10/11 via winget
winget install --id Microsoft.PowerShell --source winget

# Rclone - o script instala automaticamente se não tiver
# (preferência: Scoop; fallback: winget)
scoop install rclone
# ou: winget install --id Rclone.Rclone --source winget

# Configurar remote do Google Drive (conta da consultoria)
rclone config
# name: gdrive_saconsult
# storage: drive
# scope: 1 (full access)
# use web browser: y
# login com fernandoavena.saconsult@gmail.com
```

### Execução Manual (Menu Interativo)
```powershell
# Opção 1: PowerShell 7 (recomendado)
pwsh -ExecutionPolicy Bypass -File "C:\Users\randomacesso\sync-onedrive-gdrive-ater\start-sync.ps1"

# Opção 2: Windows PowerShell 5.1
powershell -ExecutionPolicy Bypass -File "C:\Users\randomacesso\sync-onedrive-gdrive-ater\start-sync.ps1"

# Opção 3: Git Bash / Terminal Linux-like
pwsh.exe -ExecutionPolicy Bypass -File "C:/Users/randomacesso/sync-onedrive-gdrive-ater/start-sync.ps1"
```

### Execução Automática (Agendador de Tarefas - a cada 5 min)
```powershell
# Registra tarefa no usuário atual (não precisa Admin)
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

**Para desativar/remover:**
```powershell
Disable-ScheduledTask -TaskName "Sync_ATER_OneDrive_GDrive"
# ou: Unregister-ScheduledTask -TaskName "Sync_ATER_OneDrive_GDrive" -Confirm:$false
```

## 📁 Estrutura do Projeto

```
C:\Users\randomacesso\sync-onedrive-gdrive-ater\
├── start-sync.ps1              # Script principal (menu + modo -Auto)
├── logs\                       # Logs de cada execução (sync_YYYY-MM-DD_HH-mm-ss.log)
├── Espelhamento-de-Pastas-no-Drive.md  # Documentação original
└── README.md                   # Este arquivo
```

## ⚙️ Menu Interativo

| Opção | Ação |
|-------|------|
| `[1]` | **Dry-Run** — Simula sincronização sem alterar nada (seguro para conferir) |
| `[2]` | **Sincronizar AGORA** — Executa `rclone copy --update` real |
| `[3]` | **Abrir pasta de Logs** — Abre `logs\` no Explorer |
| `[4]` | **Testar conexão** — Lista pastas no `gdrive_saconsult:VALE DO RIBEIRA/FERNANDO` |
| `[5]` | **Configurar rclone** — Abre `rclone config` com o passo a passo |
| `[6]` | **Status do rclone** — Versão, remotes e config redigida (tokens ocultos) |
| `[0]` | Sair |

### Comportamento automático do script

- **rclone ausente** → pergunta e instala via **Scoop** (preferência) ou **winget**
- **Modo `-Auto`** → instala sem perguntar (o Agendador não tem console)
- **Remote não configurado** → aborta na hora com mensagem clara (não gera 15 erros no log)
- **Exit code** → `0` se tudo ok, `1` se alguma pasta falhar (o Agendador mostra "falhou")
- **Retenção de logs** → apaga automaticamente `.log` com mais de 7 dias

## 🔧 Configurações Importantes (edite no `start-sync.ps1`)

```powershell
# Linha 78: Caminho base do seu OneDrive local
$OrigemBase = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"

# Linha 79: Remote + pasta no Google Drive
$RemotoDrive = "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO"

# Linhas 101-114: Cidades e Atividades monitoradas
$Cidades = @("Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro")
$Atividades = @("ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3")
```

> **Atenção**: Nomes devem bater **exatamente** (acentos, maiúsculas) com as pastas no OneDrive e no Google Drive. Use a opção `[4]` para confirmar.

## 🛡️ Segurança e Boas Práticas

- **Token OAuth** (`%APPDATA%\rclone\rclone.conf`): **NUNCA** commite no Git nem sincronize na nuvem
- **Modo `--update`**: Só adiciona/atualiza — **nunca deleta** arquivos que a equipe colocou no Drive
- **`-MultipleInstances IgnoreNew`**: Se sync demorar >5 min, não abre segunda instância por cima
- **Limpeza de logs**: já embutida — arquivos `.log` com mais de 7 dias são apagados a cada execução

## 🧪 Testes Rápidos

```powershell
# 1. Verifica se rclone está no PATH
rclone version

# 2. Testa autenticação e lista pastas compartilhadas
rclone lsd gdrive_saconsult: --drive-shared-with-me

# 3. Dry-run de uma pasta específica
rclone copy "C:\...\Fernando\Registro\ATIVIDADE INDIVIDUAL 1" "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/ATIVIDADE INDIVIDUAL 1/Registro" --update --drive-shared-with-me --dry-run -v

# 4. Roda script em modo simulação (opção 1 do menu)
pwsh -ExecutionPolicy Bypass -File ".\start-sync.ps1"
# Escolha [1] no menu
```

## 🐛 Troubleshooting

| Problema | Solução |
|----------|---------|
| `rclone: command not found` | Reinicie o terminal após instalar; o script recarrega PATH automaticamente |
| `Access denied` no OneDrive | Verifique se `$OrigemBase` existe e OneDrive está sincronizado |
| `Failed to get token` | Rode `rclone config` novamente e refaça login com a conta da consultoria |
| Nomes com acento não batem | Use opção `[4]` para ver nomes exatos no Drive; ajuste arrays `$Cidades`/`$Atividades` |
| Script não executa (política) | Use `-ExecutionPolicy Bypass` ou `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

## 📝 Licença

Uso interno — SA Consultoria.