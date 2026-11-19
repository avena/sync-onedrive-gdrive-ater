> 📋 **[DIA-A-DIA.md](DIA-A-DIA.md)** — Guia rápido e prático da rotina diária (campo → nuvem), pronto para imprimir em uma folha.

# Sync ATER - OneDrive → Google Drive (SA Consultoria)

Sincronizador unidirecional para copiar pastas de produtores rurais do **OneDrive local** para o **Google Drive compartilhado** da conta `fernandoavena.saconsult@gmail.com`.

## 📋 Visão Geral

- **Origem**: `C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\<Cidade>\<Atividade>\<Pasta do produtor>`
- **Destino**: `gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor - RESUMO>` *(hierarquia invertida + nome padronizado)*
- **Modo**: `rclone copy --update` — **nunca apaga** arquivos no destino (protege alterações da equipe)
- **Autenticação**: OAuth via `rclone config` (remote `gdrive_saconsult`)
- **Exceções**: pastas listadas em `ignorados.json` são puladas em todas as rotinas
- **`.json` fica só local**: o `start-sync.ps1` usa `--exclude "*.json"` — sobem ao Drive apenas `.txt`, fotos e `coordenadas_gps.txt`

## 📍 Onde salvar as fotos e os relatórios (IMPORTANTE)

O fluxo tem **duas entradas** que você precisa alimentar antes de rodar os scripts:

### 1. Fotos do celular → pasta do Syncthing
Salve/sincronize as fotos do campo na pasta do **Syncthing** no PC:

```
C:\Users\randomacesso\Sync\Camera_Celular\Camera
```

- É a variável **`$SyncthingFotos`** no topo do `triagem-campo.ps1` — hoje um **array de raízes em ordem de prioridade** (com dedup por nome): `Camera_Celular_Organizada` (C:) → `E:\Sync\Camera_Celular_Organizada` (disco externo, se conectado) → `Camera_Celular\Camera` (bruta, recém-chegadas). Ajuste se suas pastas tiverem outros nomes/caminhos.
- No celular, o Syncthing sincroniza `DCIM/Camera` em modo **Enviar apenas** para essa pasta.
- Estrutura real do Syncthing: `Sync\Camera_Celular\` contém `Camera\` (fotos da câmera, `IMG_AAAAMMDD_HHMMSS.jpg`), `Screenshots\`, `CamScanner\`, `PlantNet\`, `Restored\`, `ScreenRecorder\`. O script usa **só `Camera\`** (as fotos de campo); os `~syncthing~*.tmp` (sync em andamento) são ignorados.
- **Ative as tags de localização/GPS na câmera** do celular — o script lê o GPS do EXIF para gerar `coordenadas_gps.txt`.
- O script **filtra as fotos pelo dia exato** (`data` `AAAA-MM-DD` do relatório) com **busca cirúrgica**: acessa só os meses/dias alvo, lê nome/hora/GPS direto do `_visita.json` das visitas organizadas (**zero EXIF**) e pré-filtra avulsas/brutas pelo padrão `*AAAAMMDD*` no nome (ou `LastWriteTime`) antes de abrir qualquer foto.
- As fotos são **copiadas** (não movidas) — a pasta do Syncthing continua sendo o "negativo" da câmera.

### 2. Relatórios (.txt/.json) → inbox espelhada do Google Drive pessoal
A skill `relatorio-ater-anater` gera a estrutura de relatórios na **inbox** do seu Google Drive pessoal (unidade `G:`):

```
G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<ATIVIDADE INDIVIDUAL X>\<AAAA-MM-DD-NOME DO PRODUTOR - RESUMO>\
    ├── relatorio.txt    (com cabeçalho YAML --- ... ---)
    └── relatorio.json   (schema_version 1.0, metadados: data, periodo, hora_inicio, hora_fim, cidade, atividade, produtor_pasta)
```

- É a variável **`$InboxBase`** no topo do `triagem-campo.ps1` (padrão `G:\Meu Drive`).
- O `triagem-campo.ps1` varre os anos disponíveis, agrupa o `.txt` + `.json` de mesmo nome base como **uma única visita** e **pula as pastas já registradas** em `_processados.json` (a pasta legada `_processados\`, se existir, também fica de fora).
- Após confirmar, os originais **permanecem intactos na inbox** (cópia não-destrutiva): o `.txt` limpo e o `.json` são gravados/copiados no OneDrive e um registro é acrescentado (**append**) a `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\_processados.json` — é esse registro que garante a idempotência (não reprocessa na próxima execução).
- **Fallback**: se a inbox estiver vazia/ausente, o script usa `visitas.json` na raiz ou a colagem manual no terminal (`FIM` encerra).

### Fluxo ponta a ponta

```
Relato / Áudio
   └─ Skill relatorio-ater-anater  →  G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\  (.txt + .json)
 Fotos do celular  →  Syncthing  →  C:\Users\randomacesso\Sync\Camera_Celular\Camera
   └─ triagem-campo.ps1
        ├─ Etapa 1: revisão TUI → cria pastas no OneDrive + grava .txt limpo/.json (+ anexos ATV 1) + append _processados.json (inbox intacta)
        └─ Etapa 2: busca cirúrgica por data (_visita.json, zero EXIF), lotes por visita + 40 min, pré-sugere produtor (Enter/S), copia fotos + coordenadas_gps.txt
             └─ OneDrive: ...\Fernando\<Cidade>\<Atividade>\<AAAA-MM-DD-Produtor - RESUMO>
                  └─ start-sync.ps1 (rclone copy --update --exclude *.json)
                       └─ Google Drive SA Consultoria: VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor - RESUMO>
                            └─ runs\ : manifesto com caminho local + URL clicável de cada pasta atualizada
```

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
2. Rode a opção **[4]** (`Selecionar produtores`) para conferir a árvore `Cidade > Atividade > Produtor` do OneDrive.
3. Rode a opção **[1]** (`Dry-Run`) para simular a cópia antes de transferir arquivos reais.
4. Rode a opção **[2]** (`Sincronizar AGORA`) para executar a primeira sincronização.

## 📁 Estrutura do Projeto

```
C:\Users\randomacesso\sync-onedrive-gdrive-ater\
├── start-sync.ps1              # Script principal (menu + modo -Auto)
├── triagem-campo.ps1           # Triagem de fotos Syncthing -> OneDrive (revisão interativa + GPS)
├── limpar-screenshots.ps1      # Organiza prints por app em Screenshots\<App> + remove lockscreen + limpeza por período
├── organizar-camera.ps1        # Organiza fotos da câmera por visita (tempo+GPS) em Camera_Celular_Organizada + sync C:/E:
├── organizar-audios.ps1        # Move áudios soltos do GravadorDeConversa (Syncthing) para pastas AAAA-MM
├── estado-camera.json          # Estado persistente do organizar-camera (totais, pendências C:->E:; versionado)
├── visitas.json                # Visitas do dia geradas do relatório WhatsApp (opcional; ou colar no terminal)
├── ignorados.json              # Pastas que nunca devem ser sincronizadas (criado ao usar [I])
├── self-check.ps1              # Verificação rápida das regras de nome, idempotência e ignorados
├── triagem-log.txt             # Log de auditoria da triagem (gerado em runtime)
├── logs\                       # Logs técnicos do rclone (giram a cada 7 dias)
├── runs\                       # Manifestos de auditoria PERMANENTES (nunca apagar; locais, fora do Git)
│   ├── ...__TRIAGEM__triagem-campo.json / .txt   # origens, fotos copiadas, pastas, GPS
│   ├── ...__SCREENSHOTS__limpar-screenshots.json # prints isolados (origem -> destino, critério, bytes)
│   └── ...__<ESCOPO>__start-sync.json / .txt     # caminho local + URL Drive das pastas atualizadas
├── .plan\                      # Documentação de planejamento e histórico (oculta/dedicada)
│   ├── Espelhamento-de-Pastas-no-Drive.md  # Arquitetura final do start-sync.ps1
│   ├── Triagem-de-Fotos-no-Campo.md        # Arquitetura completa do triagem-campo.ps1
│   └── historico-prompts.md                # Histórico de prompts (implementação/refatoração/WhatsApp→JSON)
└── README.md                   # Este arquivo
```

> 📚 Toda a documentação de planejamento e o histórico de prompts ficam em [`.plan\`](.plan/):
> [Espelhamento-de-Pastas-no-Drive.md](.plan/Espelhamento-de-Pastas-no-Drive.md) ·
> [Triagem-de-Fotos-no-Campo.md](.plan/Triagem-de-Fotos-no-Campo.md) ·
> [historico-prompts.md](.plan/historico-prompts.md).

## 🏷️ Padronização do nome no Google Drive

As pastas no OneDrive costumam vir com prefixo de data ou número (`2026-11-JOAO DA SILVA`, `03 - Pedro Oliveira`). No destino o prefixo é removido por regex e o sufixo ` - RESUMO` é acrescentado **de forma idempotente**:

```powershell
$Limpo = $NomeLocal -replace '^(\d{4}-\d{2}(-\d{2})?-|\d+\s*-\s*)', ''
# Se já termina em "- RESUMO" (case-insensitive), preserva; senão acrescenta.
if ($Limpo -match '\s*-\s*RESUMO$') { $NomeDrive = $Limpo } else { $NomeDrive = "$Limpo - RESUMO" }

# 2026-11-JOAO DA SILVA                 ->  JOAO DA SILVA - RESUMO
# 2026-11-12-CARLOS - ITAPITANGUI - RESUMO ->  CARLOS - ITAPITANGUI - RESUMO  (dia completo, sem duplicar)
# 03 - Pedro Oliveira                      ->  Pedro Oliveira - RESUMO
```

**Regra idempotente para `- RESUMO`**: pastas antigas (sem o sufixo) recebem ` - RESUMO`; pastas já criadas pelo `triagem-campo.ps1` (que já vêm com ` - RESUMO`) são preservadas como estão — **nunca** vira `- RESUMO - RESUMO`. A verificação é case-insensitive (`- RESUMO`, `-RESUMO`, `- resumo` contam como já presentes).

Isso vale para **todas** as rotinas: `[1]`, `[2]`, `[3]`, `[4]` e `-Auto`. Cada pasta de produtor é uma tarefa `rclone copy` independente (`.../FERNANDO/<Atividade>/<Cidade>/<Produtor> - RESUMO`), então o seletor `[4]` mostra o mapeamento ao lado do nome local:

### URLs do Drive em `runs\` (só pastas atualizadas)

Ao final de `[2]`, `[3]`, `[4]` ou `-Auto`, o script identifica **apenas** as pastas que tiveram arquivos copiados/atualizados naquela execução e, só para essas, consulta o ID via `rclone lsjson --dirs-only --drive-shared-with-me` (com cache por pasta pai, economizando API). Grava `runs\AAAA-MM-DD_HH-mm-ss__<ESCOPO>__start-sync.json` (+ `.txt` com links clicáveis) mapeando cada produtor ao **caminho local** (`C:\Users\...`) e à **URL direta** `https://drive.google.com/drive/folders/<ID>`. Escopos: `TODAS`, `CIDADE_<nome>`, `SELETOR`, `AUTO`. Dry-run não gera manifesto.

```
  [4] Jacupiranga  >  ATIVIDADE INDIVIDUAL 1
   [ ] 2026-11-JOAO DA SILVA -> JOAO DA SILVA - RESUMO  (14 arq.)
```

> Pastas sem prefixo viram `<nome> - RESUMO` do mesmo jeito.

## 📷 Triagem de Fotos de Campo (`triagem-campo.ps1`)

Script separado que alimenta o OneDrive a partir das fotos do celular (via **Syncthing**), antes do `start-sync.ps1` subir tudo ao Drive. Arquitetura completa em [`.plan/Triagem-de-Fotos-no-Campo.md`](.plan/Triagem-de-Fotos-no-Campo.md).

```powershell
pwsh -ExecutionPolicy Bypass -File ".\triagem-campo.ps1"          # lê visitas.json se existir
pwsh -ExecutionPolicy Bypass -File ".\triagem-campo.ps1" -Colar    # força colar o JSON no terminal (FIM encerra)
```

### Fonte primária: inbox espelhada (`.txt` + `.json`)

O script varre `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\`, agrupa o `.txt` + `.json` de mesmo nome base como uma visita e usa o `.json` (`schema_version 1.0`) como fonte de metadados (`data`, `periodo`, `hora_inicio`, `hora_fim`, `cidade`, `atividade`, `produtor_pasta`). **Idempotência**: pastas cujo caminho já consta em `_processados.json` (registro único por ano, com `pasta_inbox`) são puladas automaticamente; a pasta legada `_processados\` também é ignorada. Fallback: `.txt` com YAML, `visitas.json` manual ou colagem. Só trata `.txt`/`.json` como relatório (`.gdoc`/`.gsheet` ficam intactos no Drive).

### Nome da pasta: prefixo `AAAA-MM-DD-` + sufixo `- RESUMO`

Cada `produtor_pasta` é normalizado para o prefixo de **dia completo** e sufixo ` - RESUMO` (case-insensitive, idempotente):

```
CARLOS - ITAPITANGUI (data 2026-11-12)  ->  2026-11-12-CARLOS - ITAPITANGUI - RESUMO
2026-11-ANA BEATRIZ               ->  2026-11-12-ANA BEATRIZ - RESUMO
```

O `start-sync.ps1` remove esse prefixo (dia completo ou ano-mês) ao montar o nome no Drive, então a pasta local e o destino ficam consistentes sem duplicar o sufixo.

### Separação explícita das datas

- **`data` (`AAAA-MM-DD`)**: dia exato da visita, usado **exclusivamente para buscar as fotos** daquele dia no Syncthing (EXIF `DateTimeOriginal`, padrão `AAAAMMDD` do nome ou `LastWriteTime`).
- **`produtor_pasta` (`AAAA-MM-NOME - RESUMO`)**: usa **apenas ano/mês**, nunca o dia.

### Etapa 1 — Revisão interativa ANTES de criar pastas

Antes de tocar no disco, o script abre uma lista navegável com `[Console]::ReadKey($true)` mostrando cada visita como `[AAAA-MM-DD] Cidade > Atividade > Nome_Pasta_Final` (a data exibida é a da busca de fotos):

| Tecla | Ação |
| :--- | :--- |
| `↑` / `↓` | Navegar na lista |
| `E` | **Editar** o item: **data** `AAAA-MM-DD`, **período/horário** (`manha`/`tarde`, `hora_inicio`/`hora_fim`), **cidade** (`1–6`), **atividade** (`1–6` + `7 PROJETO PRODUTIVO`), **nome da pasta** (Enter mantém em cada passo; sempre volta a ter prefixo `AAAA-MM-DD-` e sufixo ` - RESUMO`) |
| `D` / `Delete` | **Remover** o item selecionado da lista (não cria a pasta) |
| `C` / `Enter` | **Confirmar** — pergunta `"Deseja criar as pastas, gravar os relatórios no OneDrive e registrar em _processados.json? (S/N)"`. `S` cria as pastas, grava os relatórios e registra; `N` volta à lista |
| `Q` / `Esc` | **Cancelar e sair** sem criar nenhuma pasta |

`try/finally` garante que `[Console]::CursorVisible` volte a `$true` ao sair. Só depois da confirmação (`S`):
- cria `Cidade\Atividade\AAAA-MM-DD-PRODUTOR - RESUMO` no `$OneDriveBase`;
- grava o **`.txt` limpo** (remove o bloco YAML `---…---` do topo, pronto para colar no Word) e copia o **`.json`** irmão para a pasta do produtor (ambos ficam locais/OneDrive);
- **anexos automáticos só na `ATIVIDADE INDIVIDUAL 1`** (case-insensitive): imagens `.png`/`.jpg`/`.jpeg` da pasta da visita cujo nome contenha `mapa` ou `calendario` (regex `mapa|c?alend[aá]rio`, tolera `alendario-agricola.png`) são **copiadas** junto — ex.: `... - Mapa.png`, `... - Mapa.jpeg`, `calendario-agricola.png`. Da ATV 2 em diante, não busca anexos;
- **cópia não-destrutiva**: nada é movido/apagado da inbox (`.txt`, `.json`, `.gdoc`, imagens ficam intactos no `G:`); em vez da antiga árvore `_processados\`, cada visita processada ganha um registro **append** no arquivo único `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\_processados.json` com `processado_em`, `pasta_inbox`, `data`, `cidade`, `atividade`, `produtor_pasta`, `destino_onedrive`, `arquivos_levados`, o rastreio de fotos **`fotos_modo` (`automatico`/`individual`/`manual`/`pulado`), `fotos_copiadas`, `fotos_origem_sugerida`** (atualizados ao fim da Etapa 2) e o rastreio de envio **`sincronizada: false`, `sincronizada_em: null`, `drive_url: null`**. Gravação **atômica** (`.tmp` + `Move-Item -Force`), preservando todos os registros/campos existentes. O `Get-InboxVisitas` consulta esse registro e pula as pastas já processadas (idempotência). No início da triagem, um resumo do ano atual é exibido: `Historico 2026: X visita(s) triadas | Y sincronizada(s) | Z pendente(s) de envio`.

### Ciclo de vida `sincronizada` (triagem ↔ start-sync)

O `_processados.json` é o registro compartilhado do ciclo completo da visita:

1. **`triagem-campo.ps1`** registra a visita com **`sincronizada: false`** (chave de ligação: `destino_onedrive`, o caminho completo da pasta no OneDrive).
2. **`start-sync.ps1`** (`Update-ProcessadosSyncStatus`), ao final de **toda sincronização real** (`[2]`, `[3]`, `[4]` ou `-Auto` — **nunca** em Dry-Run), marca **`sincronizada: true`**, `sincronizada_em` (timestamp) e `drive_url` (URL da pasta no Drive, quando obtida via `rclone lsjson`) para cada pasta com `Codigo -eq 0` — **mesmo com 0 cópias** (os arquivos podem já ter sido enviados antes). O match com `destino_onedrive` é case-insensitive e ignora barra final; todos os demais campos/registros ficam intactos (gravação atômica `.tmp` + `Move-Item -Force`). Se o `G:` estiver offline, a atualização falha em silêncio sem derrubar o sync.
3. **Opção `[13]`** do menu lista as visitas **Pendentes de sincronização** vs. **Sincronizadas** (data, cidade, produtor, `fotos_modo`, timestamp de envio e link do Drive). Visitas com **`fotos_modo: manual`** aparecem destacadas em amarelo com o aviso *"fotos copiadas MANUALMENTE, conferir antes de enviar"* — útil quando o lote tinha fotos de produtores vizinhos misturadas.

### Etapa 2 — Triagem inteligente de fotos + GPS

Depois de criar as pastas, o script extrai as **datas únicas** do relatório (`$DatasAlvo`) e faz **busca cirúrgica** nas raízes (`$SyncthingFotos`: Organizada C: → E: → `Camera` bruta, dedup por nome) — **sem varrer o acervo inteiro**:

- **Organizada (C:/E:)**: acessa diretamente só `<AAAA>\<MM>` dos meses alvo. Cada pasta `<DD>_visita*` do dia com `_visita.json` vira um **lote pré-agrupado** (preserva a separação por tempo+GPS do `organizar-camera.ps1`), com `nome`/`hora`/`lat`/`lon` lidos **direto do JSON — zero abertura de `.jpg` via EXIF**. Avulsas soltas na raiz do mês são pré-filtradas pelo padrão `*AAAAMMDD*` no nome antes de ler EXIF.
- **Bruta (`Camera\`)**: pré-filtro pelo padrão `*AAAAMMDD*` no nome **ou** `LastWriteTime` do dia antes de chamar `Get-FotoMetadados` (fotos de outros dias nunca são abertas).
- Fotos sem visita (avulsas/brutas) agrupam por bloco de horário (intervalo de 40 min).

O script **cruza o horário médio do lote** com `hora_inicio`/`hora_fim` ou `periodo` (`manha` 06:00–12:00 / `tarde` 12:01–18:30). Havendo correspondência única, destaca a **pré-sugestão** e basta apertar **`Enter` ou `S`** para aceitar (ou digitar outra TAG / `P` pular / `I` foto a foto / `M` manual / `Q` sair). Horários incompletos nunca exibem `-null` (ex.: `14:46` em vez de `14:46-null`). Para cada lote copiado, gera `coordenadas_gps.txt` com o GPS (das visitas, vindo do `_visita.json`).

**`[M] Manual` (produtores vizinhos misturados no lote):** abre no Explorer (`Invoke-Item`) a **pasta de origem** das fotos do lote **e** as pastas de destino dos produtores daquele dia no OneDrive, para você copiar as fotos manualmente. Os registros do dia no `_processados.json` são marcados com **`fotos_modo: "manual"`** e `fotos_origem_sugerida` (caminho da pasta de origem). Ao final de cada dia, o resultado da triagem é gravado por produtor: `fotos_modo` (`automatico` = lote inteiro numa TAG, `individual` = foto a foto, `manual`, `pulado` = nada copiado), `fotos_copiadas` (contagem real de cópias com sucesso) e `fotos_origem_sugerida`.

### Auditoria permanente (`runs\`)

Ao finalizar, grava `runs\AAAA-MM-DD_HH-mm-ss__TRIAGEM__triagem-campo.json` (+ `.txt` legível) com: origens `.txt`/`.json` no Drive, caminhos exatos das fotos copiadas do Syncthing, pastas criadas no OneDrive, coordenadas GPS e contagens. **`runs\` nunca é apagado** (diferente de `logs\`). Auditoria textual também em `triagem-log.txt`.

## 📸 Organização da Câmera por Visita (`organizar-camera.ps1`)

Move as fotos da câmera bruta para **fora da pasta sincronizada**, agrupadas por **visita técnica** (tempo + raio GPS), com fluxo **campo (C:) vs. casa (disco externo E:)**. Arquitetura completa em [`.plan/Triagem-de-Fotos-no-Campo.md`](.plan/Triagem-de-Fotos-no-Campo.md) (seção 12).

```powershell
pwsh -ExecutionPolicy Bypass -File ".\organizar-camera.ps1" -SemMenu -WhatIf  # simulação (não move nada)
pwsh -ExecutionPolicy Bypass -File ".\organizar-camera.ps1"                   # interativo (menu campo/casa)
pwsh -ExecutionPolicy Bypass -File ".\organizar-camera.ps1" -SemMenu -SincronizarExterno -MoverParaExterno
pwsh -ExecutionPolicy Bypass -File ".\organizar-camera.ps1" -ReconstruirIndice  # só (re)gera o índice, não move nada
```

- **Origem**: `C:\Users\randomacesso\Sync\Camera_Celular\Camera` · **Destino local**: `C:\Users\randomacesso\Sync\Camera_Celular_Organizada\` (irmã da sincronizada, não volta ao celular) · **Externo**: `<E|outra>:\Sync\Camera_Celular_Organizada\` (prioriza `E:`; `-DriveExterno "E"` força a letra).
- **Hierarquia**: visitas (clusters **≥ 3 fotos**) em `<AAAA>\<MM>\<DD>_visita\` — ou `<DD>_visita_1`, `<DD>_visita_2`… quando há 2+ visitas no mesmo dia (ordem cronológica); avulsas (1–2 fotos) na raiz `<AAAA>\<MM>\`; vídeos (`VID_*`/`.mp4/.mov/.3gp/.mkv`) em `Videos\<AAAA>\<MM>\`.
- **Regra de cluster** (Haversine em metros): `(Δt ≤ 40 min E d ≤ 500 m)` OU `(mesmo dia E Δt ≤ 120 min E d ≤ 300 m)` OU `(sem GPS E Δt ≤ 40 min)`. Data/hora: EXIF `DateTimeOriginal` → nome `AAAAMMDD_HHMMSS` → `LastWriteTime`.
- **`_visita.json`** em cada pasta de visita: `data`, `indice_visita`, `pasta_nome`, `hora_inicio`/`hora_fim`, `duracao_minutos`, `qtd_fotos`, `fotos_com_gps`, `lat_media`/`lon_media`, `raio_metros`, `google_maps_url` e `arquivos` (nome, hora, lat, lon, modelo_camera, tamanho_bytes).
- **`_indice_visitas.json`** (raiz da Organizada, leve): gerado/atualizado ao final de **toda execução real** (e copiado ao E: na sincronização) lendo **apenas os `_visita.json` — zero EXIF**. Mapeia cada `data` (`AAAA-MM-DD`) → lista de visitas do dia com `pasta_relativa` (ex.: `2026\11\06_visita_1`), `hora_inicio`/`hora_fim`, `qtd_fotos`, `lat_media`/`lon_media`, `raio_metros` e o caminho do `_visita.json`. **`-ReconstruirIndice`** (re)gera o índice sobre o acervo já organizado **sem mover nada** (útil uma única vez sobre as 137 visitas existentes). Com `-MoverParaExterno`, o índice é gerado sobre a árvore do E: e copiado de volta à raiz do C:.
- **`estado-camera.json`** (versionado): `ultima_execucao`, `ultima_foto_processada`, totais organizados no C:, `ultima_sincronizacao_e` e pendências C:→E: (arquivos/bytes). A cada execução o script mostra fotos/vídeos novos, visitas agrupadas e GB pendentes de sync.
- **Cenário A — disco externo ausente (campo)**: avisa e pergunta se organiza localmente no C: (direto com `-SemMenu`), movendo da `Camera\` (libera o celular via Syncthing).
- **Cenário B — disco presente (casa)**: menu TUI (↑/↓ + Enter): [1] organizar novas + sincronizar tudo ao E: (pergunta manter cópia no C: ou **mover** para liberar o SSD), [2] só organizar no C:, [3] só sincronizar/mover o que já está no C:, [4] cancelar.
- **Parâmetros**: `-SemMenu`, `-SincronizarExterno`, `-MoverParaExterno`, `-DriveExterno "E"`, `-Copiar` (copia em vez de mover da `Camera\`), `-ReconstruirIndice` (só regenera o `_indice_visitas.json`, zero EXIF, não move nada), `-WhatIf`. Manifesto por execução: `runs\AAAA-MM-DD_HH-mm-ss__CAMERA__organizar-camera.json`.
- **Integração**: o `triagem-campo.ps1` lê a pasta Organizada (C: e E:) antes da `Camera\` bruta, com dedup por nome — e aproveita os `_visita.json`/`_indice_visitas.json` para busca cirúrgica por data **sem reler EXIF do acervo**. Fotos recém-chegadas ainda não organizadas continuam sendo encontradas na bruta (pré-filtro por nome/data).
- **Baseline real (17/11/2026)**: 1.368 fotos (96,2% com GPS, Xiaomi POCO F4 GT) + 1 vídeo → **254 clusters / 137 visitas / 117 avulsos**; o `-WhatIf` reproduz esses números.

## 🎙️ Organização dos Áudios do Gravador (`organizar-audios.ps1`)

Move os áudios soltos na raiz de `C:\Users\randomacesso\Sync\GravadorDeConversa` (Syncthing bidirecional com o celular HyperOS) para subpastas **`AAAA-MM`** dentro da própria pasta — a reorganização reflete automaticamente na pasta `GRAVADORDECONVERSA` do celular.

```powershell
pwsh -ExecutionPolicy Bypass -File ".\organizar-audios.ps1" -WhatIf   # simulação (não move nada)
pwsh -ExecutionPolicy Bypass -File ".\organizar-audios.ps1"           # organiza de verdade
```

- **Escopo**: só `.m4a`/`.mp3`/`.aac`/`.wav` **soltos na raiz**; subpastas existentes, ocultos do Syncthing (`.stfolder`, `.stignore`) e temporários (`~syncthing~*`, `*.tmp`, `.syncthing.*`) são ignorados.
- **Mês**: extraído do nome do Gravador Xiaomi (`3 de nov., 11.39 ...m4a` → `11`); **ano**: `LastWriteTime`. Nome fora do padrão (renomeado manualmente) → fallback completo pelo `LastWriteTime`.
- **Sem sobrescrita**: se já existir arquivo de mesmo nome em `AAAA-MM`, o áudio é mantido na raiz e um aviso é exibido.
- Compatível com Windows PowerShell 5.1 e PowerShell 7; pode ser agendado no Agendador de Tarefas como o `start-sync.ps1`.

## 🧹 Isolamento e Limpeza de Prints (`limpar-screenshots.ps1`)

Script separado que **organiza** as capturas de tela (prints) do Android **por aplicativo** dentro da própria pasta nativa `Screenshots` — nenhuma pasta é criada fora de `Camera_Celular` —, **remove prints de lockscreen automaticamente** e permite apagar prints antigos por período. Arquitetura em [`.plan/Triagem-de-Fotos-no-Campo.md`](.plan/Triagem-de-Fotos-no-Campo.md) (seção 11).

```powershell
pwsh -ExecutionPolicy Bypass -File ".\limpar-screenshots.ps1" -WhatIf              # simula: ranking por app + lockscreens
pwsh -ExecutionPolicy Bypass -File ".\limpar-screenshots.ps1"                      # organiza por app + remove lockscreen
pwsh -ExecutionPolicy Bypass -File ".\limpar-screenshots.ps1" -ManterLockscreen    # mantém lockscreen em Lockscreen\
pwsh -ExecutionPolicy Bypass -File ".\limpar-screenshots.ps1" -LimparAte "2026-09" # apaga prints até set/2026 (pede APAGAR)
```

- **Origem**: `C:\Users\randomacesso\Sync\Camera_Celular` · **Destino**: `...\Screenshots\<NomeDoAplicativo>\` (os nomes `Screenshot_AAAA-MM-DD...` já ordenam por data dentro de cada pasta). Inclui prints do layout antigo `Screenshots\AAAA\MM`. **Ao final de toda execução real** (organização ou `-LimparAte`), uma varredura *bottom-up* remove **qualquer subpasta vazia** de `Screenshots` — apps cujos prints foram apagados manualmente e anos/mês legados (`2026\01..11`, `2026`); pastas ocultas (`.stfolder` etc.) são preservadas.
- **Detecção de print** (qualquer critério basta): dentro de `*\Screenshots\*`; nome começa com `Screenshot_`; `.png` na raiz/`Camera\`/`Restored\`; resolução exata de tela (`1080x2400`/`2400x1080`) **sem** EXIF de câmera (`Make`/`Model`). Fotos reais `IMG_*` com EXIF de câmera ou proporção 4:3 **nunca** são movidas.
- **Classificação por app**: pacote extraído após o carimbo de data/hora (`Screenshot_AAAA-MM-DD-HH-mm-ss-fff_<APP>` ou `Screenshot_AAAAMMDD_HHMMSS_<APP>`), ignorando colisão `__2`/`__3` e o marcador `-edit` do Samsung. Conhecidos → `Facebook`, `LinkedIn`, `Mercado Livre`, `WhatsApp`, `Chrome`, `Instagram`, `YouTube`, `Tela Inicial`; desconhecidos → remove prefixos genéricos (`com.`, `br.com.`, `org.`, `lt.`, `eu.`, `air.`, `jus.`…) e o token `.android`, capitaliza o 1º segmento (ex.: `com.dimep.kairosemployee.br` → `Dimep`), sanitizado para Windows. Sem sufixo de app → `Outros`.
- **Lockscreen**: app `lockscreen` ou `com.android.systemui` é **removido automaticamente** (contagem + MB no console e no manifesto), salvo com `-WhatIf` ou `-ManterLockscreen` (nestes casos vai para `Lockscreen\`).
- **Ignorados na varredura**: ocultos de sistema/Syncthing (`.stfolder`, `.android`, `.deleteRecord`, `.globalTrash`, `.tmfs`, `.trashed-*` — qualquer nome iniciado por ponto), `~syncthing~*.tmp` e prints **já organizados** em `Screenshots\<App>\` (re-execução idempotente; prints soltos na raiz e no layout antigo `AAAA\MM` continuam sendo organizados).
- **Auditoria**: ranking de pastas por app (quantidade + MB) e manifesto `runs\AAAA-MM-DD_HH-mm-ss__SCREENSHOTS__limpar-screenshots.json` (`por_aplicativo`, `lockscreen_removidos`, cada `origem → destino` e **`pastas_vazias_removidas`**). **Ambos os modos geram manifesto** (o `-LimparAte` registra `corte`, arquivos removidos e pastas vazias). `-WhatIf` não altera nada e não gera manifesto.
- **`-LimparAte "AAAA-MM"`**: varre **todas as pastas de aplicativo** e seleciona prints cuja data extraída do nome (`AAAAMMDD`/`AAAA-MM-DD`, fallback `LastWriteTime`) tenha chave `AAAA-MM` **≤** ao corte; lista por pasta (arquivos + MB), avisa que **o Syncthing espelhará a remoção no celular** e só remove após digitar **`APAGAR`**. Ao final, a varredura *bottom-up* remove qualquer pasta vazia e grava manifesto próprio em `runs\`.

## 🚫 Nunca sincronizar (`ignorados.json`)

No seletor `[4]`, `I` grava a pasta atual em `ignorados.json` e `H` alterna a exibição das entradas ignoradas (aparecem no fim, em cinza, como `[IGNORADO]`). `I` sobre uma entrada `[IGNORADO]` remove o registro e a pasta volta a sincronizar.

O arquivo guarda **só o caminho relativo completo** até a pasta do produtor:

```json
[
  "Jacupiranga\\ATIVIDADE INDIVIDUAL 1\\2026-11-JOAO DA SILVA"
]
```

**Regra crítica**: o nome do produtor isolado nunca é usado. Ignorar `JOAO DA SILVA` na `ATIVIDADE INDIVIDUAL 1` não afeta a `ATIVIDADE INDIVIDUAL 2` — cada atividade tem sua própria entrada. O arquivo é consultado por `[1]`, `[2]`, `[3]`, `[4]` e `-Auto`, que pulam as entradas listadas e informam quantas foram ignoradas no resumo.

## ⚙️ Menu Interativo

| Opção | Ação |
|-------|------|
| `[1]` | **Dry-Run** — Simula sincronização sem alterar nada (seguro para conferir) |
| `[2]` | **Sincronizar AGORA** — Executa `rclone copy --update` em todas as cidades |
| `[3]` | **Sincronizar UMA cidade** — Escolhe uma cidade da lista e sincroniza só ela |
| `[4]` | **Selecionar produtores** — TUI: árvore Cidade > Atividade > Produtor com navegação por setas |
| `[5]` | **Ver últimas linhas do log** — Mostra as 30 últimas linhas do log mais recente |
| `[6]` | **Abrir pasta de Logs** — Abre `logs\` no Explorer |
| `[7]` | **Testar conexão** — Lista pastas no `gdrive_saconsult:VALE DO RIBEIRA/FERNANDO` |
| `[8]` | **Testar rclone** — Versão, remotes e config redigida (tokens ocultos) |
| `[9]` | **Configurar rclone** — Abre `rclone config` com o passo a passo |
| `[10]` | **Agendar tarefa** — Registra/atualiza a tarefa do Agendador (5 min, `-Auto`) |
| `[11]` | **Remover agendamento** — Desregistra a tarefa do Agendador |
| `[12]` | **Atualizar rclone** — `scoop update rclone` (ou `winget upgrade`) |
| `[13]` | **Status de visitas (`_processados.json`)** — lista as visitas triadas separadas entre **Pendentes de sincronização** (`sincronizada: false`) e **Sincronizadas** (`sincronizada: true`), com data, cidade, produtor, `fotos_modo` (destaque amarelo quando `manual`), timestamp de envio e URL do Drive |
| `[0]` | Sair |

> O cabeçalho do menu mostra se o agendamento está **ATIVO (5 min)** ou desativado.

### Seletor de produtores (opção `[4]`)

Interface TUI nativa (sem módulos externos), usando `[Console]::ReadKey($true)`:

| Tecla | Ação |
| :--- | :--- |
| `↑` / `↓` | Navegar na lista |
| `Espaço` | Marcar/desmarcar o produtor (`[X]` / `[ ]`) |
| `A` | Marcar/desmarcar **todos** |
| `I` | Ignorar esta atividade (grava em `ignorados.json`) / desfazer se já estiver `[IGNORADO]` |
| `H` | Mostrar/ocultar as entradas ignoradas (exibidas no fim, em cinza) |
| `D` | Dry-Run **só nos marcados** |
| `S` | Sincronizar **só os marcados** |
| `Q` / `Esc` | Voltar ao menu |

Cada produtor mostra a contagem de arquivos e como o nome ficará no Drive (`local -> NOME - RESUMO`). `try/finally` garante que o cursor volte ao normal ao sair.

### Feedback durante a execução

O `rclone` roda sempre com `--log-file`, porque o `-v` escreve em **stderr** — no Windows PowerShell 5.1 isso aparece como erro vermelho (`NativeCommandError`) se capturado pelo pipeline. Em vez disso:

- Cada pasta de produtor imprime `[3/15] Registro | ATIVIDADE INDIVIDUAL 1 | 2026-11-JOAO DA SILVA` antes de rodar.
- Depois de cada pasta: `-> 12 novo(s), 3 atualizado(s)` (ou `-> N arquivo(s) pendente(s)` no dry-run), lido das linhas novas do log.
- No fim do dry-run: **`--- RESUMO DO DRY-RUN ---`** com o total ou `[OK] Tudo sincronizado! Nenhuma alteração pendente encontrada.`
- Teclas pendentes são drenadas (`Clear-KeyBuffer`) antes do `Read-Host`, então o Enter responde na hora.

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
$Atividades = @(
    "ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3",
    "ATIVIDADE INDIVIDUAL 4", "ATIVIDADE INDIVIDUAL 5", "ATIVIDADE INDIVIDUAL 6",
    "PROJETO PRODUTIVO"
)
```

> `$Cidades` × `$Atividades` = 42 combinações possíveis; só as que existem no OneDrive entram na fila (o contador `[i/N]` mostra o total real).

> **Atenção**: Nomes devem bater **exatamente** (acentos, maiúsculas) com as pastas no OneDrive e no Google Drive. Use `[4]` (árvore local) e `[7]` (pastas no Drive) para confirmar.

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

# 3. Dry-run de uma pasta específica (nome já padronizado no destino)
rclone copy "C:\...\Fernando\Registro\ATIVIDADE INDIVIDUAL 1\2026-11-JOAO DA SILVA" "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/ATIVIDADE INDIVIDUAL 1/Registro/JOAO DA SILVA - RESUMO" --update --drive-shared-with-me --dry-run -v

# 4. Roda script em modo simulação (opção 1 do menu)
pwsh -ExecutionPolicy Bypass -File ".\start-sync.ps1"
# Escolha [1] no menu

# 5. Verifica as regras de renomeação e de ignorados (usar antes de commitar)
pwsh -NoProfile -File ".\self-check.ps1"

# 6. Simula a detecção de prints do Android (não move nada)
pwsh -NoProfile -File ".\limpar-screenshots.ps1" -WhatIf

# 7. Piloto ponta a ponta com visita "TESTE PILOTO" real (inbox -> triagem -> OneDrive -> Drive)
pwsh -NoProfile -File ".\piloto-ponta-a-ponta.ps1"          # executa e valida com asserts
pwsh -NoProfile -File ".\piloto-ponta-a-ponta.ps1" -Limpar  # remove os rastros do teste
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
- Destino nuvem (Google Drive): "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor - RESUMO>" (hierarquia invertida entre Cidade e Atividade).
- Padronização do nome: remover o prefixo com `-replace '^(\d{4}-\d{2}-|\d+\s*-\s*)', ''` e acrescentar " - RESUMO" de forma IDEMPOTENTE (se o nome limpo já termina em `\s*-\s*RESUMO$`, case-insensitive, preservar sem duplicar). Aplicar em [1], [2], [3], [4] e -Auto, uma tarefa de produtor por vez.
- Triagem de campo (triagem-campo.ps1): normaliza produtor_pasta para terminar em " - RESUMO" e abre revisão interativa TUI (setas, E editar, D/Delete remover, C/Enter confirmar S/N, Q/Esc cancelar) ANTES de criar qualquer pasta no OneDrive; try/finally restaura CursorVisible.
- Nunca sincronizar: "ignorados.json" guarda exclusivamente o caminho relativo "Cidade\Atividade\NomePastaProdutor" (nunca o nome isolado do produtor, pois a mesma pessoa pode existir em outra atividade) e todas as rotinas pulam essas entradas.
- Autenticação: Remote OAuth chamado "gdrive_saconsult" autenticado na conta "fernandoavena.saconsult@gmail.com". A pasta destino fica em "Compartilhados comigo" (NÃO é Team Drive/Shared Drive), exigindo a flag `--drive-shared-with-me`.
- Regra de ouro: Usar exclusivamente `rclone copy --update` (NUNCA `rclone sync`), pois arquivos adicionados por outros técnicos no Google Drive jamais podem ser apagados.
- O script deve manter o menu interativo (opções 0 a 13, incluindo [13] Status de visitas do _processados.json), suporte ao parâmetro `-Auto` para o Agendador de Tarefas do Windows (`-MultipleInstances IgnoreNew`), retry para rede instável (`--retries=3`) e limpeza automática de logs com mais de 7 dias.
- Ciclo de vida compartilhado: triagem-campo.ps1 registra visitas em G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\_processados.json com sincronizada=false (append atômico .tmp+Move-Item, inbox intacta) e campos fotos_modo (automatico/individual/manual/pulado), fotos_copiadas, fotos_origem_sugerida atualizados ao fim da Etapa 2 (Update-FotosProcessados); a tecla M na Etapa 2 abre origem+destinos no Explorer e marca fotos_modo=manual (lotes com produtores vizinhos misturados). start-sync.ps1 (Update-ProcessadosSyncStatus) marca sincronizada=true/sincronizada_em/drive_url após sync real (Codigo 0, nunca dry-run), casando destino_onedrive case-insensitive sem barra final e preservando os demais campos; a opção [13] lista pendentes vs sincronizadas destacando fotos_modo=manual.
```

## 📝 Licença

**Código fechado — Todos os Direitos Reservados.** Propriedade exclusiva de
Fernando Avena / SA Consultoria; uso restrito e interno. É proibida a cópia,
redistribuição, modificação ou uso comercial sem autorização expressa por
escrito. Veja [LICENSE](LICENSE).