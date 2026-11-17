# Espelhamento de Pastas no Drive — Arquitetura Final (`start-sync.ps1`)

**Status**: implementado em `start-sync.ps1` (raiz do projeto). Este documento reflete a arquitetura final.
**Script vivo**: [`../start-sync.ps1`](../start-sync.ps1) — fonte da verdade do código.
**Triagem que alimenta o OneDrive**: [`Triagem-de-Fotos-no-Campo.md`](Triagem-de-Fotos-no-Campo.md).
**Histórico de prompts / guia original**: [`historico-prompts.md`](historico-prompts.md).

---

## 1. Visão geral

Sincronizador **unidirecional e aditivo** que copia pastas de produtores rurais do **OneDrive local** para o **Google Drive compartilhado** da conta `fernandoavena.saconsult@gmail.com`, invertendo a hierarquia e padronizando o nome com o sufixo ` - RESUMO`.

```
OneDrive (local):
  ...\SA Consultoria\Vale do Ribeira\Fernando\<Cidade>\<Atividade>\<AAAA-MM-Produtor>

Google Drive (nuvem):
  gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor> - RESUMO
```

**Hierarquia invertida**: no OneDrive é `Cidade > Atividade > Produtor`; no Drive é `Atividade > Cidade > Produtor - RESUMO`.

---

## 2. Autenticação (remote OAuth)

- Remote rclone: **`gdrive_saconsult`**, autenticado via OAuth na conta **`fernandoavena.saconsult@gmail.com`**.
- A pasta `VALE DO RIBEIRA` está em **"Compartilhados comigo"** — **não é Team Drive / Shared Drive**. Por isso:
  - `rclone config` → `Configure this as a Shared Drive (Team Drive)?` = **`n`**.
  - Todas as operações usam a flag **`--drive-shared-with-me`**.
- Token OAuth fica em `%APPDATA%\rclone\rclone.conf` — **nunca** commitar nem sincronizar na nuvem (já no `.gitignore`).
- Verificação: `rclone lsd gdrive_saconsult: --drive-shared-with-me` deve listar `VALE DO RIBEIRA`.

---

## 3. Modo exclusivamente aditivo

- Regra de ouro: **`rclone copy --update`** — copia novos e atualiza mais recentes; **NUNCA apaga** nada no destino.
- **Nunca usar `rclone sync`** (apagaria arquivos que outros técnicos colocaram no Drive).
- Flags principais: `--update --exclude *.json --transfers=4 --checkers=8 --drive-shared-with-me --stats-one-line --stats=2s -v`.
- **`--exclude "*.json"`**: os `.json` estruturados ficam **apenas** no OneDrive/PC local; sobem ao Drive só `.txt`, fotos e `coordenadas_gps.txt`.
- Rede instável (campo): `--retries=3 --retries-sleep=10s --contimeout=30s`.
- Cada pasta de produtor é uma tarefa `rclone copy` independente, sempre com `--log-file` (o `-v` escreve em stderr; no PS 5.1 isso viraria erro vermelho se capturado pelo pipeline).

---

## 4. Padronização do nome no destino (idempotente para `- RESUMO`)

Função `Get-NomeResumo`:

```powershell
$Limpo = $Nome -replace '^(\d{4}-\d{2}(-\d{2})?-|\d+\s*-\s*)', ''  # remove dia completo, ano/mês ou número
if ($Limpo -match '\s*-\s*RESUMO$') { return $Limpo }              # já tem - RESUMO (case-insensitive): preserva
return "$Limpo - RESUMO"                                           # pasta antiga: acrescenta
```

- Remove o prefixo de data/número do início (`2026-11-…` ou `03 - …`).
- **Regra idempotente**: se o nome limpo já termina em `\s*-\s*RESUMO$` (case-insensitive), preserva como está — **sem duplicar** para `- RESUMO - RESUMO`. Pastas criadas pelo `triagem-campo.ps1` já vêm com ` - RESUMO`; pastas antigas recebem o sufixo.
- Exemplos:
  - `2026-11-JOAO DA SILVA` → `JOAO DA SILVA - RESUMO`
  - `03 - Pedro Oliveira` → `Pedro Oliveira - RESUMO`
  - `2026-11-CARLOS - ITAPITANGUI - RESUMO` → `CARLOS - ITAPITANGUI - RESUMO` (sem duplicar)
- Vale para **todas** as rotinas: `[1]`, `[2]`, `[3]`, `[4]` e `-Auto`.

---

## 5. Seletor TUI interativo (opção `[4]`)

Função `Show-ProdutoresSelector` — TUI nativa com `[Console]::ReadKey($true)`, árvore `Cidade > Atividade > Produtor`:

| Tecla | Ação |
| :--- | :--- |
| `↑` / `↓` | Navegar |
| `Espaço` | Marcar/desmarcar checkbox (`[X]` / `[ ]`) |
| `A` | Marcar/desmarcar todos |
| `I` | Ignorar a pasta atual (grava em `ignorados.json`) / desfazer se já `[IGNORADO]` |
| `H` | Mostrar/ocultar ignorados (exibidos no fim, em cinza) |
| `D` | Dry-run só nos marcados |
| `S` | Sincronizar só os marcados |
| `Q` / `Esc` | Voltar ao menu |

- Cada produtor mostra a contagem de arquivos e o mapeamento `local -> NOME - RESUMO`.
- `try/finally` restaura `[Console]::CursorVisible = $true`; `Clear-KeyBuffer` drena teclas pendentes antes de `Read-Host`.

### Ignorados persistentes (`ignorados.json`)

- Guarda **só o caminho relativo completo** `Cidade\Atividade\NomePastaProdutor` (nunca o nome isolado do produtor — a mesma pessoa pode existir em outra atividade).
- Consultado por `[1]`, `[2]`, `[3]`, `[4]` e `-Auto`, que pulam as entradas e informam quantas foram ignoradas.
- Exemplo: `["Jacupiranga\\ATIVIDADE INDIVIDUAL 1\\2026-11-JOAO DA SILVA"]`.

---

## 5b. URLs do Drive em `runs\` (só pastas atualizadas)

- Ao final de `[2]`, `[3]`, `[4]` ou `-Auto`, `Write-SyncRun` identifica **apenas** as pastas com arquivos efetivamente copiados/atualizados (`Copiados + Atualizados > 0`).
- Só para essas, `Get-DriveFolderUrl` consulta o ID via `rclone lsjson <pai> --dirs-only --drive-shared-with-me --quiet`, com **cache por pasta pai** (`$Script:DriveDirCache`) para economizar chamadas de API.
- Monta a URL clicável `https://drive.google.com/drive/folders/<ID>`.
- Grava `runs\AAAA-MM-DD_HH-mm-ss__<ESCOPO>__start-sync.json` (+ `.txt` legível com links) mapeando cada produtor ao **caminho local** (`C:\Users\...`) e à **URL no Drive** da SA Consultoria.
- Escopos: `TODAS` (`[2]`), `CIDADE_<nome>` (`[3]`), `SELETOR` (`[4]`), `AUTO` (`-Auto`). **Dry-run não gera manifesto.**
- `runs\` é **permanente e versionado** — nenhuma rotina de limpeza o apaga (diferente de `logs\`, que gira 7 dias).

---

## 6. Menu interativo e modo automático

Menu (opções `0`–`12`): `[1]` Dry-run, `[2]` Sincronizar todas, `[3]` Uma cidade, `[4]` Seletor TUI, `[5]` Último log, `[6]` Abrir logs, `[7]` Testar conexão, `[8]` Testar rclone, `[9]` Configurar rclone, `[10]` Agendar, `[11]` Remover agendamento, `[12]` Atualizar rclone, `[0]` Sair.

### Agendamento silencioso (`-Auto`)

- Tarefa `Sync_ATER_OneDrive_GDrive` a cada **5 min**, `-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File … -Auto`.
- **`-MultipleInstances IgnoreNew`**: se a sync demorar >5 min, não abre segunda instância por cima.
- Usa `pwsh.exe` se existir, senão `powershell.exe`.
- `-Auto` instala rclone sem prompt (Agendador não tem console) e aborta cedo se o remote não estiver configurado.
- **Exit code**: `0` se tudo ok, `1` se alguma pasta falhar (o Agendador marca "falhou").

### Manutenção automática

- **Retenção de logs**: apaga `.log` com mais de **7 dias** a cada execução.
- **Um log por execução**: `logs\sync_AAAA-MM-DD_HH-mm-ss.log`.
- **rclone ausente**: instala via Scoop (preferência) ou winget; recarrega PATH (shims/links).

---

## 7. Configurações principais (editar no `start-sync.ps1`)

```powershell
$OrigemBase  = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
$NomeRemoto  = "gdrive_saconsult"
$RemotoDrive = "${NomeRemoto}:VALE DO RIBEIRA/FERNANDO"
$Cidades     = @("Barra do Turvo","Cananéia","Jacupiranga","Juquiá","Pariquera-Açu","Registro")
$Atividades  = @("ATIVIDADE INDIVIDUAL 1".."ATIVIDADE INDIVIDUAL 6","PROJETO PRODUTIVO")
```

> Nomes devem bater **exatamente** (acentos, maiúsculas) com as pastas no OneDrive e no Drive. `$Cidades × $Atividades` = 42 combinações; só as existentes no OneDrive entram na fila.

---

## 8. Origem histórica

A primeira versão deste sincronizador nasceu de um guia em PowerShell 7 + rclone (chat Gemini, 17/11/2026). O guia original completo, o prompt de implementação e os prompts de refatoração subsequentes estão preservados em [`historico-prompts.md`](historico-prompts.md).