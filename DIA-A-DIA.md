# 📋 Rotina Dia a Dia — Fluxo ATER (Campo → Nuvem)

> **Legenda das nuvens** (aparece no terminal, logs e manifestos, sempre com o caminho completo na linha de baixo):
> **`[GD]`** = Google Drive Pessoal (`G:\Meu Drive\<AAAA>\_inbox_relatorio_ater` + `_processados.json`) ·
> **`[OD]`** = OneDrive Local (`C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\...`) ·
> **`[GDSA]`** = Google Drive Remoto da SA Consultoria (`gdrive_saconsult:VALE DO RIBEIRA/FERNANDO`) ·
> `[Syncthing]` = fotos do celular (`Sync\Camera_Celular*`).

Abrir o terminal (**PowerShell 7**) na pasta dos scripts:
```powershell
cd C:\Users\randomacesso\sync-onedrive-gdrive-ater
```

---

## 1. No Campo (Durante e Logo Após as Visitas)

* **Fotos e Vídeos:** Tire normalmente no celular (o Syncthing envia para `Sync\Camera_Celular\Camera`).
* **Relatórios:** Envie o áudio ou relato das visitas aqui no chat para a skill **`relatorio-ater-anater`**.
* Ela salva automaticamente os pares `.txt` + `.json` em `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\`.

---

## 2. No Computador (Processar o Dia de Trabalho)

### Passo 1 — Organizar Fotos e Vídeos do Celular (`organizar-camera.ps1`)

Tira os arquivos brutos da pasta do Syncthing (liberando memória no celular) e separa:

* **Fotos de visita (≥ 3 fotos):** agrupadas por horário + GPS em `Camera_Celular_Organizada\<AAAA>\<MM>\<DD>_visita\` (com `_visita.json`).
* **Fotos avulsas (1–2 fotos):** ficam na raiz `<AAAA>\<MM>\`.
* **Vídeos (`VID_*.mp4`):** separados automaticamente em `Camera_Celular_Organizada\Videos\<AAAA>\<MM>\`.

```powershell
.\organizar-camera.ps1
```

* **Sem HD/SSD externo (`E:`):** organiza no `C:\Users\randomacesso\Sync\Camera_Celular_Organizada`.
* **Com HD/SSD externo (`E:`) plugado:** escolha no menu se quer apenas organizar no `C:` ou já espelhar/mover para `E:\Sync\Camera_Celular_Organizada`.

### Passo 2 — Triagem de Campo (`triagem-campo.ps1`)

Cruza os relatórios da inbox (`G:`) com as fotos do dia, cria as pastas no OneDrive e gera `coordenadas_gps.txt`:

```powershell
.\triagem-campo.ps1                # menu inicial de modos
.\triagem-campo.ps1 -SomenteTexto  # direto: só textos + anexos ATV 1 (sem fotos)
```

* **Menu inicial** (após o resumo do ano + varredura da inbox): **`[1]`** Somente TEXTOS e anexos ATV 1 (cria pastas, gera TXT do JSON com diff e **pula fotos**) · **`[2]`** Triagem COMPLETA (Textos + Fotos/GPS) · **`[3]`** Conferir/reprocessar · **`[4]`** Colar JSON manual · **`[0]`** Sair.
* **Modo `[1]` / `-SomenteTexto`:** encerra após os textos (manifesto em `runs\` com `fotos = 0`) e orienta: rode `.\start-sync.ps1` para enviar ao `[GDSA]` agora, ou `.\triagem-campo.ps1 -Conferir` depois para anexar as fotos.
* **TXT com diff no OneDrive:** pasta existente é avisada (`[JA EXISTE] [OD]`); se o `.txt` já existe, é comparado antes de sobrescrever — **idêntico** → `[TXT IDENTICO]` (mantido sem alteração); **diferente** → `[TXT ATUALIZADO (diff)]` com resumo `-`/`+` das linhas, registrado no `triagem-log.txt`.
* **Tela 1 (Relatórios):** Confira as visitas (`E` edita, `D` remove) → aperte **`C` / `Enter`** e confirme **`S`**.
* Rodou de novo e **tudo já estava triado**? Ele avisa (nada de tela estranha de colar JSON) e oferece **`C`** = Conferir/reprocessar as visitas anteriores, **`M`** = colar JSON manual, **`Q`** = sair.
* **Tela 2 (Fotos):** O script mostra a **média de horário das fotos vs o horário do relatório** e sugere o produtor (mesmo sem `hora_fim`, caso dos relatórios de áudio) → aperte **`Enter` ou `S`** para aceitar (ou digite a letra do produtor / `P` para pular / `I` foto a foto).
* **Dia com várias pastas de fotos (`07_visita_1`, `07_visita_2`, ...)?** Aperte **`E`** para abrir o seletor: setas navegam, **`Espaço` ou `1–9`** marca, **`A`** todos/nenhum, **`O`** abre a pasta no Explorer, **`Enter`** copia **só as pastas marcadas** para a TAG escolhida. Cada linha mostra qtd de fotos, faixa, média `HH:mm`, GPS e a sugestão do relatório.
* **Lote com fotos de vizinhos misturadas?** Aperte **`M`** (Manual): o Explorer abre a pasta de origem das fotos e as pastas dos produtores do dia no OneDrive para você copiar à mão; o registro fica marcado `fotos_modo: manual` (destacado na opção `13` do `start-sync`).
* Cada visita confirmada é registrada em `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\_processados.json` com **`sincronizada: false`** (os originais da inbox ficam intactos). No início, o script mostra o resumo do ano: `X triadas | Y sincronizadas | Z pendentes de envio`.

### Passo 2b — Conferir / Reprocessar um dia já triado (`-Conferir`)

Faltou o TXT na pasta do produtor? Apagou fotos sem querer (ou quer manter só algumas)? Liste tudo que já foi processado e confira:

```powershell
.\triagem-campo.ps1 -Conferir
```

* Mostra cada visita com **TXT: OK/FALTA** e **fotos: X no destino / Y registradas**.
* **`T`** = recopia o TXT limpo (com diff: idêntico é mantido) + anexos da ATV 1 do Drive pessoal `[GD]` · **`F`** = lista as fotos que faltam e pergunta se copia (**S/N** — `N` mantém como está, sem forçar fotos a mais) · **`A`** = ambos · **`P`** = pular · **`Q`** = sair.
* **Dia com 2+ pastas de fotos e sem origem registrada?** Ao pedir fotos (`F`/`A`), o seletor de lotes abre para você marcar **qual pasta pertence àquela visita** — a escolha é gravada em `fotos_origem_sugerida` no `_processados.json` e as pastas do dia **nunca são misturadas** automaticamente.

### Passo 3 — Enviar para o Google Drive da Consultoria (`start-sync.ps1`)

Copia as pastas prontas do OneDrive para o Drive compartilhado da SA Consultoria (ignorando os `.json`) e salva os links diretos em `runs\`:

```powershell
.\start-sync.ps1
```

* Digite **`2`** (**Sincronizar AGORA**) para enviar todas as cidades, ou **`4`** para escolher produtores específicos.
* Ao final de cada sync real, as visitas correspondentes são marcadas **`sincronizada: true`** (com data/hora e link do Drive) no `_processados.json` — o Dry-Run nunca marca.
* Digite **`13`** (**Status de visitas**) para ver o que está **pendente de envio** e o que já foi **sincronizado**, incluindo o modo das fotos (`automatico`/`individual`/`manual`/`pulado`) — as **manuais** aparecem em amarelo para conferência.
* *(Se o agendamento de 5 min estiver **ATIVO**, este passo roda sozinho).*

---

## 3. Manutenção Periódica (Quando Quiser Limpar o Celular)

```powershell
.\organizar-audios.ps1     # Move áudios soltos do Gravador Xiaomi para pastas AAAA-MM
.\limpar-screenshots.ps1   # Organiza prints por aplicativo, apaga lockscreen e remove pastas vazias
.\self-check.ps1           # Valida se todos os scripts estão funcionando sem erros
```

### Teste Ponta a Ponta Real (piloto) e Limpeza dos Rastros

```powershell
.\piloto-ponta-a-ponta.ps1          # Cria visita "TESTE PILOTO" e valida inbox -> triagem -> OneDrive -> Drive
.\piloto-ponta-a-ponta.ps1 -Limpar  # Remove a visita de teste do Drive, do OneDrive e da inbox
.\triagem-campo.ps1 -Conferir       # Confere/reprocessa visitas reais sem duplicar registros
```

* O piloto exige o remote `gdrive_saconsult` configurado e que **não exista** visita TESTE PILOTO anterior (rode `-Limpar` antes de repetir).
* Manifestos em `runs\` e `triagem-log.txt` são auditoria permanente — o `-Limpar` **não** os apaga.

### Atalhos Úteis de Simulação (`-WhatIf` não altera nada)

```powershell
.\organizar-camera.ps1 -SemMenu -WhatIf    # Simula quantas visitas/vídeos novos existem
.\limpar-screenshots.ps1 -WhatIf           # Mostra o ranking de prints por app antes de mover
.\limpar-screenshots.ps1 -LimparAte "2026-09"  # Apaga prints antigos até o mês informado (pede APAGAR)
```
