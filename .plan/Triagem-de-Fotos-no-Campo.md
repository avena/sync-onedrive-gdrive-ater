# Triagem de Fotos de Campo — Syncthing → OneDrive → Google Drive

**Status**: implementado em `triagem-campo.ps1` (raiz do projeto). Este documento é a arquitetura/planejamento de referência.
**Script vivo**: [`../triagem-campo.ps1`](../triagem-campo.ps1) — fonte da verdade do código.
**Sync final**: [`../start-sync.ps1`](../start-sync.ps1) sobe as pastas ao Google Drive (ver [`Espelhamento-de-Pastas-no-Drive.md`](Espelhamento-de-Pastas-no-Drive.md)).
**Histórico de prompts**: [`historico-prompts.md`](historico-prompts.md).

---

## 1. Objetivo e fluxo completo

Eliminar a classificação manual de fotos de campo. O ciclo:

```
Celular (câmera, GPS ativado)
   └─ Syncthing (DCIM/Camera, "enviar apenas")
        └─ PC: C:\Users\randomacesso\Syncthing\Camera_Celular
             └─ triagem-campo.ps1
                  ├─ Etapa 1: revisão interativa (TUI) → cria pastas no OneDrive
                  └─ Etapa 2: filtra fotos por data, agrupa por horário, distribui por TAG, gera GPS
                       └─ OneDrive: ...\Fernando\<Cidade>\<Atividade>\<AAAA-MM-PRODUTOR - RESUMO>
                            └─ start-sync.ps1 (rclone copy --update) → Google Drive da consultoria
```

Por que Syncthing e não API do Google Fotos: **zero limite de API**, arquivo original preservado, **EXIF intacto** (data + GPS). O Syncthing só copia bytes, não mexe em metadado.

A entrada é o **relatório diário do WhatsApp** (texto livre). A IA converte em **JSON padronizado**; o script consome o JSON.

---

## 2. Separação explícita das datas do JSON (regra central)

O JSON tem **dois** campos de data com papéis distintos e que **não se misturam**:

| Campo | Formato | Papel |
|---|---|---|
| `data` | `AAAA-MM-DD` (ex. `2026-11-12`) | **Dia exato** da visita. Usado **exclusivamente para filtrar/localizar as fotos** daquele dia na pasta do Syncthing. |
| `produtor_pasta` | `AAAA-MM-NOME DO PRODUTOR - RESUMO` | Nome da pasta de destino. Usa **apenas ano e mês** (`AAAA-MM-`), **nunca o dia**. |

- A busca de fotos compara `data` com a data real de cada arquivo, resolvida nesta ordem:
  1. EXIF `DateTimeOriginal` (tag `0x9003` = 36867), via `System.Drawing` (JPG/JPEG);
  2. padrão `AAAAMMDD_HHMMSS` no nome do arquivo (Android/Samsung, ex. `20261112_143022.jpg`);
  3. fallback `LastWriteTime` (preservado pelo Syncthing).
- No código: `$FotosDoDia = @($TodasFotos | Where-Object { $_.DataStr -eq $Dia.data })`. O prefixo da pasta **não** entra nesse filtro.
- Ao carregar o JSON, `produtor_pasta` é normalizado para terminar em ` - RESUMO` (case-insensitive, idempotente — função `Add-SufixoResumo`).
- Na edição do nome (revisão), `Format-NomePastaProdutor` garante **prefixo `AAAA-MM-`** (derivado de `mes_prefixo` ou dos 7 primeiros caracteres de `data`) **e sufixo ` - RESUMO`**.

---

## 3. Estrutura de pastas (OneDrive)

```
C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\
└─ <Cidade>\
   └─ <ATIVIDADE INDIVIDUAL N | PROJETO PRODUTIVO>\
      └─ <AAAA-MM-NOME DO PRODUTOR - RESUMO>\
         ├─ foto.jpg ...
         └─ coordenadas_gps.txt   ← gerado pelo script
```

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
[AAAA-MM-DD] Cidade > Atividade > Nome_Pasta_Final
```

A data exibida é a **data de busca das fotos** (`data`); o nome final já vem com `AAAA-MM-… - RESUMO`.

| Tecla | Ação |
| :--- | :--- |
| `↑` / `↓` | Navegar na lista |
| `E` | **Editar** o item selecionado (função `Edit-VisitaItem`), em 4 passos: |
| | 1. **Data da visita** `AAAA-MM-DD` (a que busca as fotos) — Enter mantém; valida formato. |
| | 2. **Cidade** — lista numerada `[1] Barra do Turvo … [6] Registro`; Enter mantém. |
| | 3. **Atividade** — `[1..6] ATIVIDADE INDIVIDUAL 1..6`, `[7] PROJETO PRODUTIVO`; Enter mantém. |
| | 4. **Nome da pasta** — Enter mantém; garante prefixo `AAAA-MM-` e sufixo ` - RESUMO`. |
| `D` / `Delete` | **Remover** o item da lista (não cria a pasta) |
| `C` / `Enter` | **Confirmar** — pergunta `"Deseja criar todas essas pastas no OneDrive? (S/N)"`. `S` cria as pastas que faltam e segue à Etapa 2; `N` volta à lista. |
| `Q` / `Esc` | **Cancelar e sair** sem criar nenhuma pasta |

- `try/finally` garante `[Console]::CursorVisible = $true` ao sair.
- A revisão edita/remove itens **no próprio `$RelatorioDias`** (referências): editar a data muda `$Dia.data` (usado depois no filtro de fotos); remover reconstrói `$Dia.visitas` só com os sobreviventes.
- Validação anti path-traversal (`Test-CampoSeguro`) rejeita campos com `..`, `/`, `\`, `:` antes de criar pastas.
- Só após `S` as pastas `Cidade\Atividade\AAAA-MM-PRODUTOR - RESUMO` são criadas no `$OneDriveBase` (`[CRIADA]` / `[JÁ EXISTE]`, log em `triagem-log.txt`).

---

## 7. Etapa 2 — Triagem de fotos do Syncthing + GPS

1. **Indexação** (`Get-FotoMetadados`): varre `$SyncthingFotos`, resolve data/hora (EXIF → nome → `LastWriteTime`) e GPS de cada arquivo, com `Write-Progress`. Extensões: `.jpg .jpeg .png .heic .mp4`.
2. **Filtro por dia**: para cada dia do JSON, seleciona as fotos cujo `DataStr` == `data` (`AAAA-MM-DD`).
3. **Agrupamento por janela de deslocamento**: fotos do mesmo dia são cortadas em lotes quando o intervalo entre fotos consecutivas passa de **40 min** (`$IntervaloVisitaMinutos`). O deslocamento entre sítios abre essa janela naturalmente → cada visita vira um lote único.
4. **Distribuição por TAG**: para cada lote mostra horário e produtores do dia; `A`/`B`/`C` manda o lote inteiro à pasta do produtor; `P` pula; `I` classifica foto a foto; `Q` encerra a triagem.
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
4. Avaliar limpeza da pasta do Syncthing só se o disco apertar (com confirmação).