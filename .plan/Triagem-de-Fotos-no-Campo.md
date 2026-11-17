# Triagem de Fotos de Campo — Syncthing → OneDrive → Google Drive

**Status**: implementado em `triagem-campo.ps1` (raiz do projeto). Este documento é a arquitetura/planejamento de referência.
**Script vivo**: [`../triagem-campo.ps1`](../triagem-campo.ps1) — fonte da verdade do código.
**Sync final**: [`../start-sync.ps1`](../start-sync.ps1) sobe as pastas ao Google Drive (ver [`Espelhamento-de-Pastas-no-Drive.md`](Espelhamento-de-Pastas-no-Drive.md)).
**Histórico de prompts**: [`historico-prompts.md`](historico-prompts.md).

---

## 1. Objetivo e fluxo completo

Eliminar a classificação manual de fotos de campo. O ciclo:

```
Relato / Áudio → Skill relatorio-ater-anater
   └─ G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<Atividade>\<AAAA-MM-DD-NOME - RESUMO>\ (.txt + .json)
Celular (câmera, GPS ativado) → Syncthing (DCIM/Camera, "enviar apenas")
   └─ PC: C:\Users\randomacesso\Sync\Camera_Celular\Camera
        └─ triagem-campo.ps1
             ├─ Etapa 1: revisão TUI → cria pastas + grava .txt limpo/.json + arquiva _processados
             └─ Etapa 2: filtra fotos por data, agrupa 40 min, pré-sugere (Enter), distribui TAG, gera GPS
                  └─ OneDrive: ...\Fernando\<Cidade>\<Atividade>\<AAAA-MM-DD-PRODUTOR - RESUMO>
                       └─ start-sync.ps1 (rclone copy --update --exclude *.json) → Google Drive da consultoria
                            └─ runs\ : manifesto de auditoria (origens, fotos, pastas, GPS, URLs)
```

Por que Syncthing e não API do Google Fotos: **zero limite de API**, arquivo original preservado, **EXIF intacto** (data + GPS). O Syncthing só copia bytes, não mexe em metadado.

**Fonte primária** dos relatórios: a **inbox espelhada** do Google Drive pessoal (`$InboxBase = G:\Meu Drive`), gerada pela skill `relatorio-ater-anater` como pares `.txt`/`.json`. **Fallback**: `.txt` com YAML, `visitas.json` manual ou colagem no terminal (`FIM`). O relatório original nasce do relato/áudio do WhatsApp.

---

## 2. Separação explícita das datas do JSON (regra central)

O JSON tem **dois** campos de data com papéis distintos e que **não se misturam**:

| Campo | Formato | Papel |
|---|---|---|
| `data` | `AAAA-MM-DD` (ex. `2026-11-12`) | **Dia exato** da visita. Usado **exclusivamente para filtrar/localizar as fotos** daquele dia na pasta do Syncthing. |
| `produtor_pasta` | `AAAA-MM-DD-NOME DO PRODUTOR - RESUMO` | Nome da pasta de destino, com prefixo de **dia completo** (`AAAA-MM-DD-`) + sufixo ` - RESUMO`. |

- A busca de fotos compara `data` com a data real de cada arquivo, resolvida nesta ordem:
  1. EXIF `DateTimeOriginal` (tag `0x9003` = 36867), via `System.Drawing` (JPG/JPEG);
  2. padrão `AAAAMMDD_HHMMSS` no nome do arquivo (Android/Samsung, ex. `20261112_143022.jpg`);
  3. fallback `LastWriteTime` (preservado pelo Syncthing).
- No código: `$FotosDoDia = @($TodasFotos | Where-Object { $_.DataStr -eq $Dia.data })`. O prefixo da pasta **não** entra nesse filtro — o dia da busca vem sempre do campo `data`.
- `Format-NomePastaProdutor -Nome … -Data …` remove qualquer prefixo de data existente (parcial `AAAA-MM-` ou completo `AAAA-MM-DD-`) e repõe o prefixo de **dia completo** a partir de `data`, mais o sufixo ` - RESUMO` (idempotente, case-insensitive).
- O `start-sync.ps1` remove esse prefixo (dia completo ou ano-mês) ao montar o nome no Drive — ver regex em [`Espelhamento-de-Pastas-no-Drive.md`](Espelhamento-de-Pastas-no-Drive.md).

---

## 3. Estrutura de pastas (OneDrive)

```
C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\
└─ <Cidade>\
   └─ <ATIVIDADE INDIVIDUAL N | PROJETO PRODUTIVO>\
      └─ <AAAA-MM-DD-NOME DO PRODUTOR - RESUMO>\
         ├─ relatorio.txt        ← .txt limpo (sem YAML), pronto p/ Word
         ├─ relatorio.json       ← metadados estruturados (fica só local/OneDrive)
         ├─ foto.jpg ...
         └─ coordenadas_gps.txt  ← gerado pelo script
```

### 3b. Inbox espelhada (fonte primária) e `_processados`

```
G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\
└─ <MM>\<Cidade>\<ATIVIDADE INDIVIDUAL X>\<AAAA-MM-DD-NOME DO PRODUTOR - RESUMO>\
   ├─ relatorio.txt   (cabeçalho YAML --- ... ---)
   └─ relatorio.json  (schema_version 1.0; metadados: data, periodo, hora_inicio, hora_fim, cidade, atividade, produtor_pasta)
```

- `Get-InboxVisitas` varre os anos (`^\d{4}$`), meses numéricos (`^\d{2}$`, **ignora `_processados`**), e desce `Cidade > Atividade > Produtor`.
- **Dedup**: agrupa `.txt` + `.json` de mesmo nome base como uma visita; `.json` é a fonte estruturada, `.txt` (YAML) é fallback. Só `.txt`/`.json` (sem `.gdoc`).
- A `data` vem do `.json`/YAML ou, na falta, do prefixo `AAAA-MM-DD` do nome da pasta.
- Após confirmar (`S`), `Invoke-CriarPastasEArquivar` grava o `.txt` **sem o bloco YAML** (`Remove-YamlHeader`) e copia o `.json` para a pasta do produtor, e **move** os originais para `_processados\<MM>\<Cidade>\<Atividade>\<leaf>\` (não reprocessa).

### 3c. Auditoria permanente (`runs\`)

`Write-TriagemRun` grava `runs\AAAA-MM-DD_HH-mm-ss__TRIAGEM__triagem-campo.json` (+ `.txt`) com: data/hora, origens `.txt`/`.json` no Drive, caminhos exatos das fotos copiadas do Syncthing, pastas criadas no OneDrive, GPS por produtor (média + contagem) e totais. **`runs\` nunca é apagado** por nenhuma rotina (diferente de `logs\`, que gira 7 dias).

Cidades oficiais (acentuação precisa bater com o OneDrive/Drive): `Barra do Turvo`, `Cananéia`, `Jacupiranga`, `Juquiá`, `Pariquera-Açu`, `Registro`. O WhatsApp chega sem acento (`Cananeia`) — o parser corrige para evitar pasta duplicada.

---

## 4. Contrato do JSON

Um objeto por dia, com as visitas e uma **tag** para digitar no terminal.

```json
[
  {
    "data": "2026-11-12",
    "mes_prefixo": "2026-11",
    "visitas": [
      { "tag": "A", "cidade": "Cananéia", "atividade": "ATIVIDADE INDIVIDUAL 2", "produtor_pasta": "2026-11-OSMAR - ITAPITANGUI" },
      { "tag": "B", "cidade": "Juquiá", "atividade": "ATIVIDADE INDIVIDUAL 2", "produtor_pasta": "2026-11-MARIA DAS GRAÇAS" }
    ]
  },
  {
    "data": "2026-11-11",
    "mes_prefixo": "2026-11",
    "visitas": [
      { "tag": "A", "cidade": "Jacupiranga", "atividade": "ATIVIDADE INDIVIDUAL 1", "produtor_pasta": "2026-11-MATILDES E OSVALDO - MORANGABA" },
      { "tag": "B", "cidade": "Juquiá", "atividade": "ATIVIDADE INDIVIDUAL 1", "produtor_pasta": "2026-11-MARIA DAS GRAÇAS" }
    ]
  },
  {
    "data": "2026-11-06",
    "mes_prefixo": "2026-11",
    "visitas": [
      { "tag": "A", "cidade": "Juquiá", "atividade": "ATIVIDADE INDIVIDUAL 1", "produtor_pasta": "2026-11-FABIA E AFONSO" },
      { "tag": "B", "cidade": "Juquiá", "atividade": "ATIVIDADE INDIVIDUAL 1", "produtor_pasta": "2026-11-IVANETE E ADEMIR" },
      { "tag": "C", "cidade": "Juquiá", "atividade": "ATIVIDADE INDIVIDUAL 1", "produtor_pasta": "2026-11-MARIA DA GRAÇA - MODESTO" }
    ]
  }
]
```

Regras:
- `data`: `AAAA-MM-DD` — dia exato para buscar fotos.
- `mes_prefixo`: `AAAA-MM` — prefixo da pasta do produtor.
- `tag`: `A`, `B`, `C`… sequencial **dentro do dia**.
- `cidade`: nome oficial acentuado.
- `atividade`: caixa alta, `ATIVIDADE INDIVIDUAL 1..6` ou `PROJETO PRODUTIVO`.
- `produtor_pasta`: caixa alta, `AAAA-MM-NOME` (o script acrescenta ` - RESUMO`); ` - ` no lugar de `/` quando há localidade.

---

## 5. Prompt reutilizável (WhatsApp → JSON)

```text
Converta o meu relatório diário de visitas de ATER abaixo em um array JSON válido para o script `triagem-campo.ps1`.

Regras obrigatórias de normalização:
1. Cada bloco "Atividade Realizada no dia DD/MM" vira um objeto com "data": "AAAA-MM-DD" e "mes_prefixo": "AAAA-MM".
2. Atribua sequencialmente "tag": "A", "B", "C", etc. para cada produtor dentro daquele dia.
3. Normalize o nome da cidade com acentuação oficial do Vale do Ribeira:
   "Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro".
4. Normalize a atividade em caixa alta: "ATIVIDADE INDIVIDUAL 1".."ATIVIDADE INDIVIDUAL 6" ou "PROJETO PRODUTIVO".
5. Formate "produtor_pasta" em CAIXA ALTA no padrão "AAAA-MM-NOME DO PRODUTOR"
   (substituindo "/" por " - " se houver localidade, ex: "2026-11-OSMAR - ITAPITANGUI").
   O sufixo " - RESUMO" é acrescentado automaticamente pelo script.
6. Retorne APENAS o bloco JSON puro, sem texto adicional, pronto para copiar e colar.
```

---

## 6. Etapa 1 — Revisão interativa (TUI) ANTES de criar pastas

Função `Show-RevisaoVisitas`. Depois de ler o JSON e normalizar o sufixo, o script abre uma lista navegável com `[Console]::ReadKey($true)` (setas ↑/↓), mostrando cada visita como:

```
[AAAA-MM-DD] (periodo / hora_inicio-hora_fim) Cidade > Atividade > Nome_Pasta_Final
```

A data exibida é a **data de busca das fotos** (`data`); o nome final já vem com `AAAA-MM-DD-… - RESUMO`.

| Tecla | Ação |
| :--- | :--- |
| `↑` / `↓` | Navegar na lista |
| `E` | **Editar** o item (`Edit-VisitaItem`): |
| | 1. **Data** `AAAA-MM-DD` (a que busca as fotos) — Enter mantém; valida formato. |
| | 2. **Período** — `[1] manha`, `[2] tarde`, `[0] limpar`; Enter mantém. |
| | 3. **Horário** — `hora_inicio`/`hora_fim` `HH:mm` (Enter mantém, `-` limpa). |
| | 4. **Cidade** — `[1] Barra do Turvo … [6] Registro`; Enter mantém. |
| | 5. **Atividade** — `[1..6] ATIVIDADE INDIVIDUAL 1..6`, `[7] PROJETO PRODUTIVO`; Enter mantém. |
| | 6. **Nome da pasta** — Enter mantém; garante prefixo `AAAA-MM-DD-` e sufixo ` - RESUMO`. |
| `D` / `Delete` | **Remover** o item da lista (não cria a pasta) |
| `C` / `Enter` | **Confirmar** — pergunta `"Deseja criar as pastas, gravar os relatórios no OneDrive e arquivar na inbox? (S/N)"`. `S` cria/grava/arquiva e segue à Etapa 2; `N` volta à lista. |
| `Q` / `Esc` | **Cancelar e sair** sem criar nenhuma pasta |

- `try/finally` garante `[Console]::CursorVisible = $true` ao sair.
- A revisão edita/remove itens **no próprio `$RelatorioDias`** (referências): editar a data muda `$Dia.data` (usado depois no filtro de fotos); remover reconstrói `$Dia.visitas` só com os sobreviventes.
- Validação anti path-traversal (`Test-CampoSeguro`) rejeita campos com `..`, `/`, `\`, `:` antes de criar pastas.
- Só após `S`, `Invoke-CriarPastasEArquivar` cria `Cidade\Atividade\AAAA-MM-DD-PRODUTOR - RESUMO` no `$OneDriveBase`, grava o `.txt` limpo + `.json` e move os originais da inbox para `_processados` (`[CRIADA]` / `[JÁ EXISTE]`, log em `triagem-log.txt`).

---

## 7. Etapa 2 — Triagem de fotos do Syncthing + GPS

1. **Indexação** (`Get-FotoMetadados`): varre `$SyncthingFotos`, resolve data/hora (EXIF → nome → `LastWriteTime`) e GPS de cada arquivo, com `Write-Progress`. Extensões: `.jpg .jpeg .png .heic .mp4`.
2. **Filtro por dia**: para cada dia do JSON, seleciona as fotos cujo `DataStr` == `data` (`AAAA-MM-DD`).
3. **Agrupamento por janela de deslocamento**: fotos do mesmo dia são cortadas em lotes quando o intervalo entre fotos consecutivas passa de **40 min** (`$IntervaloVisitaMinutos`). O deslocamento entre sítios abre essa janela naturalmente → cada visita vira um lote único.
4. **Pré-sugestão + distribuição por TAG** (`Get-SugestaoTag`): cruza o **horário médio do lote** com `hora_inicio`/`hora_fim` (quando não nulos) ou `periodo` (`manha` 06:00–12:00 / `tarde` 12:01–18:30). Se **um único** produtor casa, destaca a sugestão e **`Enter` aceita**; senão digita-se `A`/`B`/`C` (lote inteiro), `P` pula, `I` foto a foto, `Q` encerra.
5. **GPS** (`Export-CoordenadasTxt`): das fotos copiadas, lê `GPSLatitude/Longitude` (tags EXIF 1–4, racionais → graus decimais) e grava/append `coordenadas_gps.txt` na pasta do produtor com:
   - **ponto médio da propriedade** (média das fotos com GPS);
   - coordenadas decimais (ex. `-24.684120, -47.882340`);
   - **link Google Maps** por foto e por lote (`https://www.google.com/maps?q=lat,lon`).
6. **Cópia**: `Copy-Item` (não `Move-Item`) — a pasta do Syncthing é o "negativo" da câmera e não deve ser esvaziada.

> `.heic`/`.mp4` não têm EXIF lido por `System.Drawing`: caem no fallback de `LastWriteTime` e ficam "sem GPS".

---

## 8. Uso no dia a dia

```powershell
pwsh -ExecutionPolicy Bypass -File ".\triagem-campo.ps1"          # lê visitas.json se existir
pwsh -ExecutionPolicy Bypass -File ".\triagem-campo.ps1" -Colar    # força colar JSON no terminal (FIM encerra)
```

1. **Uma vez só**: ajustar `$SyncthingFotos` no topo do script; conferir que a câmera do celular salva localização.
2. Colar o relatório do WhatsApp no chat → receber o JSON (prompt da seção 5).
3. Rodar o script → revisar na TUI (editar/remover/confirmar) → `S` cria as pastas.
4. Para cada lote de horário, digitar `A`/`B`/`C` (ou `I`). No fim, `.\start-sync.ps1` sobe tudo ao Drive.

---

## 9. Decisões congeladas

- **Copiar, não mover** — a pasta do Syncthing é o negativo da câmera.
- **Lote por intervalo de tempo (40 min)** — 1 tecla classifica ~12 fotos; o deslocamento separa as visitas.
- **Sem API do Google Fotos** — EXIF/qualidade original preservados, sem quota.
- **JSON vindo da IA, não do script** — interpretar texto livre de WhatsApp é trabalho de LLM, não de regex frágil.
- **`data` (dia) só para fotos; pasta só ano/mês** — separação explícita evita confundir o dia da visita com o nome da pasta.
- **`- RESUMO` idempotente** — a pasta local já nasce com o nome que o `start-sync.ps1` espera no Drive, sem duplicar.
- **`coordenadas_gps.txt` por produtor** — vai junto com as fotos para a pasta `- RESUMO` no Drive (QGIS/CAR/laudo).

---

## 10. Pendências / próximos passos

1. Rodar com um dia real do relatório, validando: revisão TUI, agrupamento de 40 min, `Copy-Item` e `coordenadas_gps.txt`.
2. **Validar o parser GPS em fotos reais** — a conversão assume racionais little-endian; EXIF big-endian e fotos sem GPS precisam de teste.
3. Teste de estresse: dia com 3 produtores e 60+ fotos; dia em que faltou foto de um produtor.
4. ~~Avaliar limpeza da pasta do Syncthing só se o disco apertar (com confirmação).~~ → implementado em `limpar-screenshots.ps1` (seção 11): prints isolados em `Capturas_de_Tela` e limpeza por período com confirmação `APAGAR`.

---

## 11. Isolamento e limpeza de prints (`limpar-screenshots.ps1`)

**Status**: implementado em `limpar-screenshots.ps1` (raiz do projeto). Complementa a triagem: a pasta do Syncthing continua sendo o "negativo" da **câmera** (`Camera\`, intocada); os **prints** são organizados **por aplicativo** dentro da própria pasta nativa `Screenshots` — nenhuma pasta é criada fora de `Camera_Celular` — e prints de **lockscreen** são removidos automaticamente.

### 11.1 Caminhos e modos

| Item | Valor |
|---|---|
| `$OrigemSync` | `C:\Users\randomacesso\Sync\Camera_Celular` |
| `$DestinoBase` | `C:\Users\randomacesso\Sync\Camera_Celular\Screenshots` (pasta nativa; só cria `<App>\ dentro dela) |
| Modo padrão | **Migra** prints para `$DestinoBase\<NomeDoAplicativo>\` + remove lockscreen + ranking por app + manifesto `runs\` |
| `-WhatIf` | Simula: ranking por app + lockscreens que seriam removidos, nada alterado, sem manifesto |
| `-ManterLockscreen` | Não remove lockscreen; organiza-os em `Lockscreen\` |
| `-LimparAte "AAAA-MM"` | **Apaga** prints de todas as pastas de app cuja data do nome seja ≤ corte (confirmação `APAGAR`) |

**Destino dentro da origem → idempotência obrigatória**: a varredura pula arquivos já organizados em `Screenshots\<App>\` (regex `^Screenshots\\(?!\d{4}\\)[^\\]+\\` sobre o caminho relativo — qualquer subpasta cujo 1º segmento **não** seja ano numérico); sem isso, o critério 1 (`*\Screenshots\*`) re-moveria os mesmos prints a cada execução. Prints **soltos na raiz** de `Screenshots` e os do **layout antigo `AAAA\MM`** (execução anterior) continuam sendo detectados e reorganizados por app; as pastas de ano/mês que ficarem vazias são removidas ao final.

### 11.2 Critérios combinados de detecção (qualquer um basta)

1. Dentro de subpasta `*\Screenshots\*` (qualquer profundidade);
2. Nome começa com `Screenshot_` (case-insensitive);
3. `.png` na raiz de `$OrigemSync` ou em `Camera\`/`Restored\`;
4. Resolução **exata** de tela (`1080x2400` ou `2400x1080`, via `System.Drawing.Image`) **e** sem EXIF de câmera (`Make` tag 271 / `Model` tag 272 vazios).

**Proteções e exclusões da varredura**:
- O critério 4 nunca se aplica a arquivos `IMG_*` (fotos reais); EXIF de câmera preenchido ou proporção 4:3 jamais casa com a resolução exata de tela.
- **Ocultos de sistema/Syncthing ignorados**: qualquer segmento do caminho relativo iniciado por ponto — pastas `.stfolder`, `.android`, `.deleteRecord`, `.globalTrash`, `.tmfs` e arquivos `.trashed-*` (lixeira do Syncthing) — via regex `(^|\\)\.`.
- `~syncthing~*.tmp` (sync em andamento) ignorados; prints já em `Screenshots\<App>\` pulados (idempotência, ver 11.1).

Na validação real (17/11/2026, com o layout antigo `2026\01..11` já populado pela execução anterior): **473 prints** = 471 do layout antigo + 2 novos soltos na raiz (sync ao vivo do celular); **18 lockscreen** (15,3 MB) marcados para remoção automática; **zero** arquivos de `Camera\` (`IMG_*`), `CamScanner\`, `PlantNet\` ou `ScreenRecorder\` detectados.

### 11.3 Classificação por aplicativo e organização

**Extração do app** (`Get-AppDoPrint`): remove a extensão e o sufixo de colisão `__N`, então casa `^Screenshot_[\d_-]+_([A-Za-z].+)$` — o pacote vem após o carimbo de data/hora nos dois formatos reais: `Screenshot_AAAA-MM-DD-HH-mm-ss-fff_<APP>` (Samsung) e `Screenshot_AAAAMMDD_HHMMSS_<APP>`. O capture exige inicial alfabética, então `Screenshot_20261117_213706` (sem app) não confunde o horário com pacote → `Outros`.

**Nome da pasta** (`Get-NomePastaApp`):
1. Minúsculas + remove o marcador `-edit` final (print editado no Samsung; ex.: `com.android.chrome-edit` → Chrome).
2. Dicionário `$AppsAmigaveis`: `com.facebook.katana`/`com.facebook.lite` → **Facebook**; `com.linkedin.android` → **LinkedIn**; `com.mercadolibre` → **Mercado Livre**; `com.whatsapp`/`com.whatsapp.w4b` → **WhatsApp**; `com.android.chrome` → **Chrome**; `com.instagram.android` → **Instagram**; `com.google.android.youtube` → **YouTube**; `com.sec.android.app.launcher` → **Tela Inicial**.
3. Desconhecidos: divide por `.`, descarta o token `android` e os segmentos iniciais genéricos (`com`, `org`, `net`, `br`, `lt`, `eu`, `air`, `jus`, `gov`), capitaliza o 1º segmento restante e sanitiza para Windows (`\/:*?"<>|`, pontos/espaços finais). Ex.: `com.dimep.kairosemployee.br` → **Dimep**; `br.com.maxatacadista` → **Maxatacadista**; `lt.noframe.fieldsareameasure` → **Noframe**; `org.plantnet` → **Plantnet**; `com.android.settings` → **Settings**.
4. Sem sufixo de app (ou nome vazio após sanitização) → **Outros** (ex.: `image.png` do `Restored\`).

**Lockscreen** (`$AppsLockscreen = lockscreen, com.android.systemui`): removido automaticamente com `Remove-Item -Force` (contagem + MB no console e no manifesto), **salvo** `-WhatIf` ou `-ManterLockscreen` — nestes casos é organizado em `Lockscreen\`.

**Organização**: destino `$DestinoBase\<App>\<nome original>` — como os nomes começam com `Screenshot_AAAA-MM-DD...`, ficam naturalmente ordenados por data dentro de cada pasta. Colisão de nomes ganha sufixo `__2`, `__3`… antes da extensão. Ao final, `Remove-PastasVazias` apaga meses/anos vazios do layout antigo (`2026\01..11`) e pastas de app vazias.

**Data do print** (`Get-DataPrint`, usada no `-LimparAte`): `AAAAMMDD` → `AAAA-MM-DD` no nome (validação de mês/dia + fronteiras de dígito) → fallback `LastWriteTime`.

**Auditoria**: ranking por app (quantidade + MB) no console e manifesto permanente `runs\AAAA-MM-DD_HH-mm-ss__SCREENSHOTS__limpar-screenshots.json` com `totais` (movidos + lockscreen removidos), `por_aplicativo`, `lockscreen_removidos` e cada `origem → destino` + app + critério + bytes.

### 11.4 Limpeza por corte temporal (`-LimparAte`)

1. Valida o formato `AAAA-MM` (`ValidatePattern`).
2. Varre **todas as pastas de aplicativo** em `$DestinoBase` (recursivo, ignorando ocultos/`.trashed-*`/`~syncthing~*.tmp`; alcança também o layout antigo `AAAA\MM` e a raiz) e seleciona os prints cuja **data extraída do nome** (fallback `LastWriteTime`) tenha chave `AAAA-MM` **≤** ao corte (comparação de string).
3. Lista os alvos agrupados por pasta (arquivos + MB) e o total; avisa em vermelho que **o Syncthing espelhará a remoção no celular**.
4. Exige `Read-Host` com resposta **exata** `APAGAR` (case-sensitive); qualquer outra coisa cancela.
5. `Remove-Item -Force` em cada arquivo e, ao final, `Remove-PastasVazias` (pastas de app que ficaram vazias + meses/anos do layout antigo). `-WhatIf` lista sem pedir confirmação e sem remover.

### 11.5 Decisões congeladas

- **Destino na pasta nativa `Screenshots`, por app** — nenhuma pasta nova fora de `Camera_Celular` (o Syncthing já espelha essa pasta com o celular); só as subpastas `<App>\` são criadas. `Camera\` permanece intacta.
- **Ordenação por data vem de graça** — os nomes `Screenshot_AAAA-MM-DD...` já ordenam cronologicamente dentro de cada pasta de app; não precisa de hierarquia `AAAA\MM` (o layout antigo é absorvido e as pastas vazias removidas).
- **Lockscreen é lixo** — `lockscreen`/`com.android.systemui` some automaticamente (18 prints/15,3 MB na validação); `-ManterLockscreen` é a válvula de escape para auditoria.
- **Re-execução idempotente** — prints já em `Screenshots\<App>\` são pulados na varredura; rodar de novo só organiza os novos/soltos (o sync ao vivo do celular entrega prints novos na raiz entre execuções).
- **`-LimparAte` pela data do nome, não pela pasta** — com a organização por app, o corte temporal varre todas as pastas de app e compara a data extraída do arquivo; continua funcionando mesmo antes de migrar o layout antigo.
- **Critério 4 conservador** — só resolução exata + EXIF vazio + nome não `IMG_*`; na dúvida, o arquivo fica na origem.
- **Ocultos e `.trashed-*` ignorados** — pastas de sistema/Syncthing (`.stfolder`, `.android`, `.deleteRecord`, `.globalTrash`, `.tmfs`) e a lixeira `.trashed-*` não entram na migração nem na contagem.
- **Confirmação `APAGAR`** — remoção é destrutiva e espelhada no celular; nunca automática.