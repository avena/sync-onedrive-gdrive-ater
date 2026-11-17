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
```

> Depois de instalar, configure o remote do Google Drive — veja
> **[Configurar o Remote do Google Drive](#configurar-o-remote-do-google-drive-rclone-config)**.

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

## 🔐 Configurar o Remote do Google Drive (`rclone config`)

Execute a opção **[9]** no menu do script (ou rode `rclone config` no terminal) e responda exatamente nesta ordem:

| Prompt do Rclone | Resposta | Observação |
| :--- | :--- | :--- |
| `n/s/q>` | `n` | Criar novo remote |
| `name>` | `gdrive_saconsult` | Deve ser exatamente este nome |
| `Storage>` | `24` (ou `drive`) | Google Drive |
| `client_id>` | *(Enter vazio)* | Usa o ID padrão (ou cole o seu do Google Cloud) |
| `client_secret>` | *(Enter vazio)* | Deixe vazio se não usou `client_id` próprio |
| `scope>` | `1` | Acesso total (`drive`) |
| `service_account_file>` | *(Enter vazio)* | Não utilizado |
| `Edit advanced config?` | `n` | Pula configurações avançadas |
| `Continue using the shared client_id anyway?` | `y` | Confirma uso do ID padrão do rclone |
| `Use web browser to automatically authenticate?` | `y` | **Logar com `fernandoavena.saconsult@gmail.com` no navegador** |
| `Configure this as a Shared Drive (Team Drive)?` | `n` | **Sempre `n`** (a pasta está em *Compartilhados comigo*, não é Team Drive corporativo) |
| `Keep this "gdrive_saconsult" remote?` | `y` | Salva em `%APPDATA%\rclone\rclone.conf` |
| `e/n/d/r/c/s/q>` | `q` | Fecha o configurador |

### Por que cada resposta

- **`client_id`/`client_secret` vazios** → usa a chave compartilhada padrão do rclone. O Google vai aposentar esse ID compartilhado em 2026; funciona hoje, mas no futuro vale criar o seu próprio no Google Cloud.
- **`scope 1`** → acesso completo, necessário para ler e gravar nas pastas compartilhadas com você.
- **`service_account_file`** → só para robôs de servidor sem login interativo. Não é o nosso caso.
- **`Shared Drive (Team Drive)?` = `n`** → Shared Drive é recurso exclusivo de contas Google Workspace pagas. A pasta `VALE DO RIBEIRA` é uma pasta comum da aba *Compartilhados comigo* de uma conta `@gmail.com`.

> **Errou o Shared Drive?** Entre em `[9]` → `e` (Edit) → `1` (`gdrive_saconsult`) → vá dando Enter até `Configure this as a Shared Drive?` e mude para `n`.

### Checklist Pós-Configuração (Validação)

1. Rode a opção **[7]** (`Testar conexão com Google Drive`) para confirmar acesso a `VALE DO RIBEIRA/FERNANDO`.
2. Rode a opção **[4]** (`Mostrar pastas locais encontradas`) para conferir os pares `Cidade\Atividade` no OneDrive.
3. Rode a opção **[1]** (`Dry-Run`) para simular a cópia antes de transferir arquivos reais.
4. Rode a opção **[2]** (`Sincronizar AGORA`) para executar a primeira sincronização.

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
| `[2]` | **Sincronizar AGORA** — Executa `rclone copy --update` em todas as cidades |
| `[3]` | **Sincronizar UMA cidade** — Escolhe uma cidade da lista e sincroniza só ela |
| `[4]` | **Mostrar pastas locais** — Lista quais pares Cidade\Atividade existem no OneDrive |
| `[5]` | **Ver últimas linhas do log** — Mostra as 30 últimas linhas do log mais recente |
| `[6]` | **Abrir pasta de Logs** — Abre `logs\` no Explorer |
| `[7]` | **Testar conexão** — Lista pastas no `gdrive_saconsult:VALE DO RIBEIRA/FERNANDO` |
| `[8]` | **Testar rclone** — Versão, remotes e config redigida (tokens ocultos) |
| `[9]` | **Configurar rclone** — Abre `rclone config` com o passo a passo |
| `[10]` | **Agendar tarefa** — Registra/atualiza a tarefa do Agendador (5 min, `-Auto`) |
| `[11]` | **Remover agendamento** — Desregistra a tarefa do Agendador |
| `[12]` | **Atualizar rclone** — `scoop update rclone` (ou `winget upgrade`) |
| `[0]` | Sair |

> O cabeçalho do menu mostra se o agendamento está **ATIVO (5 min)** ou desativado.

### Comportamento automático do script

- **rclone ausente** → pergunta e instala via **Scoop** (preferência) ou **winget**
- **Modo `-Auto`** → instala sem perguntar (o Agendador não tem console)
- **Remote não configurado** → aborta na hora com mensagem clara (não gera 15 erros no log)
- **Exit code** → `0` se tudo ok, `1` se alguma pasta falhar (o Agendador mostra "falhou")
- **Retenção de logs** → apaga automaticamente `.log` com mais de 7 dias
- **Um log por execução** → `sync_AAAA-MM-DD_HH-mm-ss.log` (não por sessão do menu)
- **Rede instável** → `--retries=3 --retries-sleep=10s --contimeout=30s` (retry nativo do rclone)
- **Agendamento** → usa `pwsh.exe` se existir, senão cai para `powershell.exe`

## 🔧 Configurações Importantes (edite no `start-sync.ps1`)

```powershell
# Linha 17: Caminho base do seu OneDrive local
$OrigemBase = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"

# Linhas 18-19: nome do remote e pasta no Google Drive
$NomeRemoto  = "gdrive_saconsult"
$RemotoDrive = "${NomeRemoto}:VALE DO RIBEIRA/FERNANDO"

# Linhas 24-38: Cidades e Atividades monitoradas
$Cidades = @("Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro")
$Atividades = @("ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3")
```

> **Atenção**: Nomes devem bater **exatamente** (acentos, maiúsculas) com as pastas no OneDrive e no Google Drive. Use `[4]` (pastas locais) e `[7]` (pastas no Drive) para confirmar.

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
| Nomes com acento não batem | Use `[7]` para ver os nomes exatos no Drive; ajuste os arrays `$Cidades`/`$Atividades` |
| Script não executa (política) | Use `-ExecutionPolicy Bypass` ou `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

## 🤖 Prompt de Contexto para Manutenção (IA)

```text
Você está trabalhando no projeto "Sync ATER" localizado em "C:\Users\randomacesso\sync-onedrive-gdrive-ater\start-sync.ps1" (PowerShell 7 + Rclone).

Contexto técnico obrigatório:
- Origem local (OneDrive): "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\<Cidade>\<Atividade>\<Produtor>"
- Destino nuvem (Google Drive): "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor>" (hierarquia invertida entre Cidade e Atividade).
- Autenticação: Remote OAuth chamado "gdrive_saconsult" autenticado na conta "fernandoavena.saconsult@gmail.com". A pasta destino fica em "Compartilhados comigo" (NÃO é Team Drive/Shared Drive), exigindo a flag `--drive-shared-with-me`.
- Regra de ouro: Usar exclusivamente `rclone copy --update` (NUNCA `rclone sync`), pois arquivos adicionados por outros técnicos no Google Drive jamais podem ser apagados.
- O script deve manter o menu interativo (opções 0 a 12), suporte ao parâmetro `-Auto` para o Agendador de Tarefas do Windows (`-MultipleInstances IgnoreNew`), retry para rede instável (`--retries=3`) e limpeza automática de logs com mais de 7 dias.
```

## 📝 Licença

Uso interno — SA Consultoria.