# Histórico de Prompts — Projeto Sync ATER / Triagem de Campo

Registro dos prompts de implementação, refatoração e conversão WhatsApp → JSON usados na construção do projeto, para referência futura. Os scripts vivos na raiz (`start-sync.ps1`, `triagem-campo.ps1`) são a fonte da verdade do código; este arquivo guarda a intenção e as regras pedidas em cada etapa.

---

## A. Prompt original de implementação (guia Gemini, 17/11/2026)

Contexto: a conta `fernandoavena.saconsult@gmail.com` tem permissão de edição na pasta compartilhada `VALE DO RIBEIRA`, mas não está no Google Drive Desktop do Windows. Solução: `rclone` via API (OAuth), sem montar disco virtual.

Prompt de manutenção extraído do guia original:

```text
Atue como engenheiro de automação Windows com PowerShell 7 e Rclone.
Tenho um projeto em "$HOME\sync-onedrive-gdrive-ater\start-sync.ps1" que sincroniza pastas de produtores rurais do OneDrive local para o Google Drive compartilhado da conta "fernandoavena.saconsult@gmail.com" (remoto rclone: "gdrive_saconsult").

Regras obrigatórias da arquitetura:
1. Origem local: "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\<Cidade>\<Atividade>\<Produtor>"
2. Destino nuvem: "gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/<Atividade>/<Cidade>/<Produtor>" (note a inversão entre Cidade e Atividade).
3. Usar exclusivamente `rclone copy --update --drive-shared-with-me` (nunca `sync`), pois arquivos adicionados por terceiros no Google Drive jamais podem ser apagados.
4. Manter suporte a menu interativo (opções de Dry-Run, Sync Manual, Logs e Teste de Conexão) e flag `-Auto` para execução silenciosa via Agendador de Tarefas do Windows com bloqueio de instâncias simultâneas (`IgnoreNew`).
5. Incluir rotina automática de limpeza de arquivos de log com mais de 7 dias.
```

Respostas do `rclone config` (remote `gdrive_saconsult`): `n` → name `gdrive_saconsult` → storage `drive` → client_id/client_secret em branco → scope `1` → service_account em branco → Edit advanced `n` → browser `y` (logar com a conta da consultoria) → **Shared Drive (Team Drive)? `n`** → Keep `y` → `q`.

---

## B. Prompt reutilizável: Relatório WhatsApp → JSON

Usado para converter o relatório diário de visitas no JSON consumido por `triagem-campo.ps1`:

```text
Converta o meu relatório diário de visitas de ATER abaixo em um array JSON válido para o script `triagem-campo.ps1`.

Regras obrigatórias de normalização:
1. Cada bloco "Atividade Realizada no dia DD/MM" vira um objeto com "data": "AAAA-MM-DD" e "mes_prefixo": "AAAA-MM".
2. Atribua sequencialmente "tag": "A", "B", "C", etc. para cada produtor dentro daquele dia.
3. Normalize o nome da cidade com acentuação oficial do Vale do Ribeira:
   "Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro".
4. Normalize a atividade em caixa alta: "ATIVIDADE INDIVIDUAL 1".."ATIVIDADE INDIVIDUAL 6" ou "PROJETO PRODUTIVO".
5. Formate "produtor_pasta" em CAIXA ALTA no padrão "AAAA-MM-NOME DO PRODUTOR"
   (substituindo "/" por " - " se houver localidade, ex: "2026-11-CARLOS - ITAPITANGUI").
   O sufixo " - RESUMO" é acrescentado automaticamente pelo script.
6. Retorne APENAS o bloco JSON puro, sem texto adicional, pronto para copiar e colar.
```

---

## C. Linha do tempo dos prompts de implementação/refatoração

### C1. Documentar a ideia da triagem de campo
> "Cria agora um novo markdown… a gente vai criar um outro projeto, mas só vamos documentar toda essa ideia… documenta toda ela registra pelo menos continuar."

Ideia registrada: fluxo **Celular (Câmera) → Syncthing → Relatório do WhatsApp → JSON → Criação automática de pastas + distribuição das fotos por data/hora → Sync com Google Drive**. Syncthing elimina limites de API do Google Fotos e preserva EXIF. Incluiu o contrato JSON (3 dias de exemplo), o prompt WhatsApp→JSON e um rascunho de `triagem-campo.ps1` (agrupamento por 40 min, tags A/B/C, `Copy-Item`, extração GPS EXIF → `coordenadas_gps.txt`). Resultado: documento de planejamento (hoje `.plan/Triagem-de-Fotos-no-Campo.md`).

### C2. Criar os códigos (sem executar)
> "E a todos os códigos, mas não executa nada só cria o código seguindo às boas normas do planejamento… amanhã eu testo."

Criação de `triagem-campo.ps1` com params `-Colar`/`-VisitasJson`, leitura de `visitas.json` ou colagem (`FIM`), criação de pastas, indexação EXIF/nome/LastWriteTime, agrupamento por 40 min, classificação por tag, GPS e `coordenadas_gps.txt`.

### C3. Análise pelos skills de PowerShell
> "A sua análise é usando todos os skills que foram já passadas como que está os scripts em Powershop."

Revisão com 5 skills (powershell, powershell-windows, reliable-execution, security-hardening, ui-architect). Achados principais: path-traversal nos campos do JSON, `Join-Path` multi-segmento, null-checks (`@()`), `+=` em loop, `Set-StrictMode`, `#Requires -Version 7.0`, `Write-Progress`, opção de sair, log em arquivo.

### C4. Commit + push
> "Quando terminar, registra adiciona todos os arquivos para comitar, faça uma mensagem de comita explicando tudo e depois faça o push."

Aplicadas as correções (alta+média) e commit/push de todos os arquivos. `.gitignore` passou a excluir `triagem-log.txt` e `visitas.json`.

### C5. Revisão interativa + sufixo `- RESUMO` (1ª versão)
> "Atualize os scripts `triagem-campo.ps1` e `start-sync.ps1` (e a documentação no `README.md`)… Revisão Interativa Antes de Criar Pastas + Sufixo `- RESUMO`… Idempotência para `- RESUMO`."

- `triagem-campo.ps1`: normalização ` - RESUMO` (case-insensitive) + TUI de revisão com setas (`E` editar cidade/atividade/nome, `D`/`Delete` remover, `C`/`Enter` confirmar S/N, `Q`/`Esc` cancelar), `try/finally` restaurando cursor.
- `start-sync.ps1`: `Get-NomeResumo` idempotente (remove prefixo, não duplica `- RESUMO`).

### C6. Separação explícita de datas + edição de data + pasta `.plan` (versão atual)
> "Atualize os scripts… Separação Explícita: `data` AAAA-MM-DD para Fotos vs. `AAAA-MM` na Pasta… Organize e consolide toda a documentação de planejamento e histórico de prompts… em uma pasta oculta/dedicada chamada `.plan`."

- `triagem-campo.ps1`:
  - `data` (`AAAA-MM-DD`) usado **exclusivamente** para filtrar fotos do dia (EXIF `DateTimeOriginal`, nome `AAAAMMDD` ou `LastWriteTime`); `produtor_pasta` usa **só ano/mês** (`AAAA-MM-NOME - RESUMO`).
  - Revisão TUI ganha **edição da data** como passo 1 do `[E]` (depois cidade, atividade `[1..7]`, nome com prefixo `AAAA-MM-` + sufixo ` - RESUMO`).
  - Helpers `Add-SufixoResumo` e `Format-NomePastaProdutor`.
- `start-sync.ps1`: regra idempotente `- RESUMO` confirmada.
- Documentação consolidada em `.plan/` (`Espelhamento-de-Pastas-no-Drive.md`, `Triagem-de-Fotos-no-Campo.md`, `historico-prompts.md`); `README.md` atualizado para apontar para `.plan/`.

### C7. Prompt Mestre consolidado (`.plan/Relato_Drive_Triagem.md`) — inbox espelhada + runs\ (versão atual)
> "Vamos ver esse plano e aplicar ele… a skill vai gerar uma nova estrutura e a gente vai copiar essa estrutura no OneDrive e depois escanear as fotos… me deixa no README descrito onde eu tenho que salvar as fotos que eu vou sincronizar." → escopo escolhido: **Plano completo (Prompt Mestre)**.

Aplicado integralmente (após auditoria prévia exigida pelo plano):
- `triagem-campo.ps1`:
  - **Fonte primária inbox espelhada** `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<Atividade>\<AAAA-MM-DD-NOME - RESUMO>\` (`Get-InboxVisitas`); varre anos, ignora `_processados`; **dedup** `.txt`+`.json` por nome base; `.json` `schema_version 1.0` como fonte de metadados; fallback `.txt` YAML / `visitas.json` / colagem.
  - Prefixo da pasta passa a **dia completo** `AAAA-MM-DD-` (`Format-NomePastaProdutor -Data`); `data` continua só para buscar fotos.
  - Revisão TUI: `[E]` edita data/período/horário/cidade/atividade/nome; confirmação grava `.txt` **sem YAML** (`Remove-YamlHeader`) + `.json` e **move originais para `_processados`** (`Invoke-CriarPastasEArquivar`).
  - Etapa 2: **pré-sugestão** por `hora_inicio`/`hora_fim` ou `periodo` (`manha`/`tarde`), **`Enter` aceita** (`Get-SugestaoTag`).
  - **Manifesto `runs\`** (`Write-TriagemRun`): origens, fotos copiadas, pastas, GPS, contagens.
- `start-sync.ps1`:
  - **`--exclude "*.json"`** (`.json` fica só local/OneDrive).
  - Regex remove **dia completo** também: `^(\d{4}-\d{2}(-\d{2})?-|\d+\s*-\s*)`.
  - **URLs do Drive** só das pastas atualizadas via `rclone lsjson --dirs-only` com cache (`Get-DriveFolderUrl`) + **manifesto `runs\`** (`Write-SyncRun`, escopos `TODAS`/`CIDADE_*`/`SELETOR`/`AUTO`).
- `runs/` criada (permanente, versionada, `.gitkeep`); `.gitignore` ignora `logs/`, `*.log`, `*.tmp`.
- `README.md`: seção **"Onde salvar as fotos e os relatórios"** (Syncthing `C:\Users\randomacesso\Sync\Camera_Celular\Camera` + inbox `G:\Meu Drive`), fluxo ponta a ponta e estrutura com `runs\`.
- `self-check.ps1`: testes de `Remove-YamlHeader`, `Read-MetadadosTxt`, `ConvertTo-Minutos`, `Get-SugestaoTag`, `ConvertTo-RelatorioDias`, `Format-NomePastaProdutor -Data` e regex de dia completo.

---

## D. Regras de ouro acumuladas (referência rápida)

1. **`rclone copy --update`**, nunca `sync` (não apagar arquivos da equipe no Drive).
2. **`--drive-shared-with-me`** e Shared Drive = `n` (pasta em "Compartilhados comigo").
3. **Hierarquia invertida**: OneDrive `Cidade>Atividade>Produtor` → Drive `Atividade>Cidade>Produtor - RESUMO`.
4. **`- RESUMO` idempotente** nos dois scripts (case-insensitive, sem duplicar).
5. **`data` (dia) só para buscar fotos; pasta usa prefixo de dia completo `AAAA-MM-DD-` + ` - RESUMO`.**
6. **`ignorados.json` por caminho relativo completo** (nunca nome isolado do produtor).
7. **Copiar, não mover** as fotos do Syncthing (negativo da câmera).
8. **Token OAuth nunca commitado** (`rclone.conf` no `.gitignore`).
9. **`.json` fica só local/OneDrive** (`--exclude "*.json"` no rclone); sobem `.txt`, fotos e `coordenadas_gps.txt`.
10. **Inbox espelhada é a fonte primária**; originais processados vão para `_processados` (não reprocessa).
11. **`runs\` é permanente e versionado** (manifestos de auditoria); só `logs\` gira 7 dias.
12. **URLs do Drive só das pastas atualizadas** (lsjson com cache, economia de API).