# Correção do fluxo de triagem e sincronização ATER — 2026-11-19

## Diagnóstico

### Problema 1 — `start-sync.ps1`: envio retroativo de pastas manuais sem relatório
- `Get-Produtores` (linha ~305) retorna **toda** pasta de produtor encontrada em
  `<OneDrive>\<Cidade>\<Atividade>\`, sem exigir sufixo `- RESUMO` nem `.txt` de relatório.
- Consequência: pastas manuais antigas (ex.: `04-2026-11-Juari`,
  `05-2026-11-Coicieçao E Hermes - coleta de solo`, `06-2026-11-Maria Da Graça - Modesto`)
  sobem retroativamente para `gdrive_saconsult` mesmo sem relatório, tanto no sync
  automático/menu quanto no seletor TUI (`Show-ProdutoresSelector`).

### Problema 2 — `triagem-campo.ps1`: `_processados.json` oculta pastas do Drive
- `Get-InboxVisitas` (linha ~161) pula qualquer pasta já registrada em `_processados.json`
  (`$jaProcessados`), e o modo `[C] Conferir/reprocessar` (`Invoke-ConferirReprocessar`)
  lista **apenas registros do JSON** (`Get-ProcessadosTodos`).
- Consequência: a tela nunca mostra o que existe fisicamente no Drive; pastas registradas
  ficam invisíveis para nova triagem/reprocessamento, e quando não há visitas novas o
  script só diz "Nenhuma visita NOVA" sem listar o conteúdo real da inbox.

## Checklist de alterações (arquivo por arquivo)

### `start-sync.ps1`
- [x] Nova função `Test-TravaResumoTxt`: pasta só é sincronizável se o nome termina com
      `- RESUMO` (`-match '\s*-\s*RESUMO$'`) **e** contém ≥1 `.txt` de relatório
      (ignorando `coordenadas_gps.txt`).
- [x] `Get-Produtores`: aplica a trava por padrão (sync + seletor TUI só veem pastas válidas);
      switch `-IncluirBloqueadas` devolve também as bloqueadas com flag `Sincronizavel`
      (para contagem informativa, nunca para envio).
- [x] `Invoke-AterSync`: separa bloqueadas das ignoradas via `ignorados.json` e exibe
      `[INFO] X pasta(s) antiga(s)/sem relatório ignorada(s) pela trava - RESUMO + .txt`.

### `triagem-campo.ps1`
- [x] Nova função `Get-InboxPastasFisicas`: varredura física de
      `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\<MM>\<Cidade>\<Atividade>\<Produtor>\`
      listando **todas** as pastas (mesmo vazias), cruzando `_processados.json` apenas
      para o status: `[NOVA]`, `[TRIADA - PENDENTE SYNC]`, `[SINCRONIZADA]`.
- [x] `Get-InboxVisitas` passa a consumir a varredura física (idempotência preservada:
      só pastas `NOVA` viram visitas novas; parsing txt/json inalterado).
- [x] Nova função `Show-InboxFisico`: lista agrupada por Mês | Cidade | Atividade | Produtor
      com status e, abaixo de cada item, `Origem [GD]: <caminho completo>`.
- [x] Início do script: resumo da varredura física; quando não há visitas novas, lista
      todas as pastas do Drive com status em vez de só "Nenhuma visita NOVA".
- [x] `Invoke-ConferirReprocessar` reescrito sobre a varredura física: qualquer pasta pode
      ser reprocessada — pastas já triadas mantêm o fluxo [T]/[F]/[A]; pastas `NOVA` ganham
      opção [R] que recria/atualiza a pasta padronizada `AAAA-MM-DD-NOME - RESUMO` no
      OneDrive (`.txt` limpo sem YAML + anexos Mapa/Calendário na ATV 1 + registro) e
      oferece a cópia das fotos faltantes do `Camera_Celular_Organizada`.
- [x] Extração do bloco de fotos faltantes para `Copy-FotosFaltantesVisita` (reuso entre
      pastas triadas e reprocessamento de pastas NOVA; sem duplicar lógica).

### `self-check.ps1`
- [x] Pastas de teste válidas do `Get-Produtores` passam a ter `- RESUMO` + `.txt` de relatório.
- [x] Teste: `Get-Produtores` ignora pasta **sem** `- RESUMO` (mesmo com `.txt`).
- [x] Teste: `Get-Produtores` ignora pasta `- RESUMO` só com fotos ou só com `coordenadas_gps.txt`.
- [x] Teste: varredura física lista todas as pastas do Drive (inclusive vazia) com status
      correto (NOVA vs. TRIADA - PENDENTE SYNC vs. SINCRONIZADA) e `Get-InboxVisitas`
      continua idempotente.
- [x] Rodar `pwsh -NoProfile -File .\self-check.ps1` até passar 100%.

## Resultado da execução — 2026-11-19

**Status: CONCLUÍDO. `pwsh -NoProfile -File .\self-check.ps1` → exit code 0,
"[SELF-CHECK] Ambiente temporario removido. OK - self-check passed." (100%, sem erros).**

### O que foi alterado e validado
1. `start-sync.ps1` — trava anti-retroativo ativa: `Get-Produtores` só retorna pastas
   `- RESUMO` com `.txt` de relatório (≠ `coordenadas_gps.txt`); pastas como
   `04-2026-11-Juari`, `05-...-Coicieçao E Hermes - coleta de solo` e
   `06-...-Maria Da Graça - Modesto` nunca sobem ao `gdrive_saconsult`.
   `Invoke-AterSync` exibe a linha `[INFO] X pasta(s) antiga(s)/sem relatório ignorada(s)
   pela trava - RESUMO + .txt` (contagem via `-IncluirBloqueadas`, sem afetar o envio).
   Validado pelos testes 1t/2/3/4 do self-check (4 válidas, 3 bloqueadas, ignorados.json intacto).
2. `triagem-campo.ps1` — Drive primeiro: `Get-InboxPastasFisicas` lista fisicamente todas
   as pastas da inbox (inclusive vazias) e cruza o status do `_processados.json`
   (NOVA / TRIADA - PENDENTE SYNC / SINCRONIZADA). Início do script mostra o resumo e,
   sem visitas novas, a listagem completa com `Origem [GD]: <caminho>` por item.
   Modo `[C]` reprocessa qualquer pasta: [T]/[F]/[A] para já triadas; [R] para NOVA
   (recria pasta padronizada + .txt limpo + anexos ATV 1 + registro + oferta de fotos
   faltantes via `Copy-FotosFaltantesVisita`). Validado pelo teste 1u (4 pastas, 3 status,
   triagem da NOVA → PENDENTE, idempotência preservada) e pelos testes 1c3/1p/1q/1r/1s existentes.
3. `self-check.ps1` — testes existentes do `Get-Produtores` atualizados (pastas válidas
   com `- RESUMO` + `.txt`) e novos testes explícitos da trava (1t) e da varredura física (1u).

### Nota técnica (armadilha corrigida durante a validação)
`Get-InboxVisitas`/`Get-InboxPastasFisicas` usam `return , @(...)` (array preservado como
objeto único). Chamadas diretas `@(Get-InboxVisitas)` ou pipe direto contam/filtram o array
inteiro em vez dos itens — por isso os pontos de chamada atribuem primeiro
(`$x = Get-InboxVisitas`) e só então usam `@($x)`/pipe. O filtro de pasta NOVA no modo [C]
foi corrigido para esse padrão.