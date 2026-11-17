Esse prompt das 4h da manhã já foi superado pelas melhorias que definimos depois, mas há **três detalhes operacionais excelentes** nele que valem muito a pena resgatar e incorporar no nosso Prompt Mestre final (descartando a parte do `.gdoc`, já que agora trabalhamos apenas com `.txt` + `.json`):

1. **Pasta `_processados` dentro da Inbox do Drive:** Em vez de apenas copiar e deixar os arquivos soltos na inbox (ou apagá-los sem rastro), mover os pares `.txt` e `.json` já processados para `<AAAA>\_inbox_relatorio_ater\_processados\` evita que o script leia os mesmos relatórios na próxima execução e mantém um histórico limpo na nuvem.
2. **Remoção do bloco YAML ao salvar o `.txt` final no OneDrive:** Quando o `triagem-campo.ps1` copiar o `.txt` para a pasta do produtor no OneDrive, ele pode remover as linhas do cabeçalho `--- ... ---` (já que esses metadados já estarão salvos no `.json` irmão), deixando o `.txt` ainda mais limpo para você abrir e colar no Word.
3. **Atalho `Enter` para aceitar pré-sugestão de período/horário na triagem de fotos:** Quando o lote de fotos do Syncthing bater com o `hora_inicio`/`hora_fim` ou `periodo` (`manha`/`tarde`) de um único produtor daquele dia, o script já destaca a sugestão e permite apenas apertar **`Enter`** para confirmar e copiar.

---

### Prompt Mestre Final (Consolidado + Melhores Ideias das 4h da Manhã, Apenas `.txt` + `.json`)

Este prompt substitui todos os anteriores. Cole-o no seu assistente de código na pasta `C:\Users\randomacesso\sync-onedrive-gdrive-ater\`:

```text
Atue como engenheiro de software e automação Windows (PowerShell 7 + Rclone + Git).
Você vai trabalhar no projeto `C:\Users\randomacesso\sync-onedrive-gdrive-ater\`.

### REGRA OBRIGATÓRIA DE EXECUÇÃO (AUDITORIA ANTES DE ALTERAR)
ANTES de criar, mover ou editar qualquer arquivo, faça uma varredura completa nos arquivos atuais da pasta `C:\Users\randomacesso\sync-onedrive-gdrive-ater\` (`start-sync.ps1`, `triagem-campo.ps1`, `README.md`, `.plan\*`, `ignorados.json`, etc.).
Compare o estado atual de cada arquivo com todos os requisitos listados abaixo e apresente primeiro um **Diagnóstico de Auditoria** separando claramente:
1. O que JÁ está implementado e funcional;
2. O que precisa ser AJUSTADO ou REFATORADO;
3. O que ainda NÃO existe e precisa ser CRIADO.
Somente após concluir e exibir essa auditoria, aplique as alterações abaixo.

---

### ESPECIFICAÇÃO COMPLETA DA ARQUITETURA (APENAS .TXT + .JSON)

#### 1. Organização de Pastas, Planos (`.plan\`), Histórico (`runs\`) e Versionamento Git
- **Pasta `.plan\`:** Mova `Espelhamento-de-Pastas-no-Drive.md` e qualquer plano da triagem para `.plan\`. Mantenha atualizados:
  - `.plan\Espelhamento-de-Pastas-no-Drive.md`
  - `.plan\Triagem-de-Fotos-no-Campo.md`
  - `.plan\historico-prompts.md` (consolidando o histórico das decisões e prompts desta arquitetura).
- **Pasta `runs\` (Permanente — NUNCA apagar):** Crie `C:\Users\randomacesso\sync-onedrive-gdrive-ater\runs\`. Diferente da pasta `logs\` (que gira logs técnicos do rclone a cada 7 dias), os arquivos em `runs\` NUNCA devem ser apagados por nenhuma rotina de limpeza.
- **Versionamento Git:** Inicialize um repositório Git (`git init`) na raiz do projeto caso ainda não exista. Crie/atualize o `.gitignore` para ignorar arquivos temporários/sensíveis (`logs/*.log`, `*.tmp`), mas **versione os scripts (`.ps1`), planos (`.plan\`), `README.md`, `ignorados.json` e os manifestos de auditoria em `runs\`**.

#### 2. Atualizações no `triagem-campo.ps1` (Inbox Espelhada `.txt`/`.json` + Revisão TUI + Fotos/GPS + `runs\`)
- **Fonte Primária — Inbox Espelhada do Google Drive (`_inbox_relatorio_ater`):**
  - Varra dinamicamente os anos disponíveis na pasta sincronizada do Google Drive pessoal (`G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\`), ignorando a subpasta `_processados`, buscando a árvore espelhada:
    `<MM>\<Cidade>\<ATIVIDADE INDIVIDUAL X>\<AAAA-MM-DD-NOME DO PRODUTOR - RESUMO>\`
  - Trate **apenas arquivos `.txt` e `.json`** (não implementar suporte a `.gdoc`).
  - **Deduplicação:** Agrupe o `.txt` e o `.json` de mesmo nome base como uma única visita, usando o `.json` (`schema_version: "1.0"`) como fonte estruturada de metadados (`metadados.data`, `periodo`, `hora_inicio`, `hora_fim`, `cidade`, `atividade`, `produtor_pasta`). Mantenha fallback para `.txt` com YAML ou `visitas.json` manual caso necessário.
  - Garanta que `produtor_pasta` termine com `" - RESUMO"` (verificação case-insensitive sem duplicar).
- **Etapa 1 — Revisão Interativa (TUI com Setas ↑/↓) ANTES de criar pastas no OneDrive:**
  - Exiba todas as visitas detectadas mostrando `[AAAA-MM-DD] (periodo / hora_inicio-hora_fim) Cidade > Atividade > produtor_pasta`.
  - Atalhos:
    - `[E]` **Editar item selecionado:** permite alterar Data (`AAAA-MM-DD` — usada exclusivamente para buscar fotos do dia), Período/Horário, Cidade (`[1..6]`: Barra do Turvo, Cananéia, Jacupiranga, Juquiá, Pariquera-Açu, Registro), Atividade (`[1..7]`: ATIVIDADE INDIVIDUAL 1 a 6 e PROJETO PRODUTIVO) e Nome da pasta do produtor (mantendo prefixo `AAAA-MM-DD-` e sufixo ` - RESUMO`).
    - `[D]` ou `[Delete]` **Remover item** da lista antes de criar.
    - `[C]` ou `[Enter]` **Confirmar e Criar Pastas:** pergunta `"Deseja criar as pastas, gravar os relatórios no OneDrive e arquivar na inbox? (S/N)"`. Ao confirmar `S`:
      - Cria `$OneDriveBase\<Cidade>\<Atividade>\<produtor_pasta>`.
      - Salva o `.txt` limpo (removendo o bloco YAML `--- ... ---` do topo para facilitar a colagem direta no Word) e copia o `.json` estruturado para dentro da pasta do produtor no OneDrive (mantendo ambos locais no OneDrive/PC).
      - Move a pasta/arquivos originais processados da inbox para `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\_processados\<MM>\<Cidade>\<Atividade>\` para não reprocessar nas próximas execuções.
    - `[Q]` / `[Esc]` **Cancelar e Sair**.
  - Sempre restaure `[Console]::CursorVisible = $true` em bloco `try/finally`.
- **Etapa 2 — Triagem Inteligente de Fotos do Syncthing + GPS EXIF:**
  - Filtra as fotos do `$SyncthingFotos` pelo dia exato (`data`: `AAAA-MM-DD`).
  - Agrupa fotos próximas no tempo (`40 min` de intervalo de deslocamento) e cruza o horário médio do lote com `hora_inicio`/`hora_fim` (quando não forem `null`) ou `periodo` (`manha` 06:00–12:00 / `tarde` 12:01–18:30).
  - Se houver correspondência única de horário/período, destaque a **pré-sugestão automática** permitindo pressionar **`Enter` para aceitar** ou digitar outra TAG / `P` para pular.
  - Copia as fotos para a pasta do produtor no OneDrive e gera/atualiza `coordenadas_gps.txt` (Lat/Lon decimal, ponto médio e links Google Maps).
- **Registro de Execução em `runs\`:**
  - Ao finalizar, salve em `runs\AAAA-MM-DD_HH-mm-ss__TRIAGEM__triagem-campo.json` (e um resumo `.txt` legível) contendo: data/hora, caminhos de origem do `.txt`/`.json` no Google Drive, lista exata dos caminhos das fotos originais copiadas do Syncthing, caminho completo criado no OneDrive (`C:\Users\...`), coordenadas GPS extraídas e contagem de arquivos.

#### 3. Atualizações no `start-sync.ps1` (Sync OneDrive -> Google Drive SA Consult + URLs em `runs\`)
- **Proteção de Arquivos Locais (`.json` fica apenas local/OneDrive):**
  - Adicione `--exclude "*.json"` nas flags do `rclone copy` para que os arquivos `.json` estruturados permaneçam **apenas** no seu OneDrive/PC local e **não** sejam enviados ao Google Drive da SA Consultoria (enviando apenas `.txt`, fotos e `coordenadas_gps.txt`).
- **Normalização Idempotente do Nome de Destino:**
  - Remova o prefixo de data completa ou ano-mês ou número do início do nome da pasta local: `-replace '^(\d{4}-\d{2}(-\d{2})?-|\d+\s*-\s*)', ''`.
  - Se o nome resultante já terminar com `\s*-\s*RESUMO$` (case-insensitive), preserve-o sem duplicar; caso contrário, anexe `" - RESUMO"`.
- **Persistência de Ignorados por Atividade (`ignorados.json`):**
  - No seletor TUI (opção `[4]`), mantenha `[I]` para ignorar/restaurar exclusivamente a chave completa `Cidade\Atividade\NomePastaProdutor` no `ignorados.json` (nunca afetando outras atividades do mesmo produtor) e `[H]` para exibir/ocultar ignorados.
- **Extração de URLs do Google Drive Apenas das Pastas Atualizadas (`runs\`):**
  - Ao final da sincronização (opções `[2]`, `[3]`, `[4]` ou `-Auto`), identifique **apenas** as pastas de produtores que tiveram arquivos efetivamente copiados ou atualizados naquela execução.
  - Apenas para essas pastas modificadas (economizando chamadas de API), consulte o `ID` da pasta no Google Drive via `rclone lsjson` (`--dirs-only --drive-shared-with-me`) e monte a URL direta clicável: `https://drive.google.com/drive/folders/<ID>`.
  - Salve em `runs\AAAA-MM-DD_HH-mm-ss__<ESCOPO>__start-sync.json` (e um `.txt` legível com links clicáveis) mapeando cada produtor sincronizado ao seu **caminho local no Windows (`C:\Users\...`)** e à sua **URL direta no Google Drive da SA Consultoria**.

#### 4. Atualização do `README.md`
- Atualize o `README.md` refletindo a nova estrutura de diretórios (`.plan\`, `runs\`, `logs\`, `ignorados.json`), o fluxo completo ponta a ponta (`Relato/Áudio -> Skill relatorio-ater-anater -> _inbox_relatorio_ater -> triagem-campo.ps1 -> OneDrive -> start-sync.ps1 -> Google Drive SA Consultoria`) e as instruções de uso.

```