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
.\triagem-campo.ps1
```

* **Tela 1 (Relatórios):** Confira as visitas (`E` edita, `D` remove) → aperte **`C` / `Enter`** e confirme **`S`**.
* **Tela 2 (Fotos):** O script sugere o produtor pelo horário da visita → aperte **`Enter` ou `S`** para aceitar (ou digite a letra do produtor / `P` para pular / `I` foto a foto).
* **Lote com fotos de vizinhos misturadas?** Aperte **`M`** (Manual): o Explorer abre a pasta de origem das fotos e as pastas dos produtores do dia no OneDrive para você copiar à mão; o registro fica marcado `fotos_modo: manual` (destacado na opção `13` do `start-sync`).
* Cada visita confirmada é registrada em `G:\Meu Drive\<AAAA>\_inbox_relatorio_ater\_processados.json` com **`sincronizada: false`** (os originais da inbox ficam intactos). No início, o script mostra o resumo do ano: `X triadas | Y sincronizadas | Z pendentes de envio`.

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

### Atalhos Úteis de Simulação (`-WhatIf` não altera nada)

```powershell
.\organizar-camera.ps1 -SemMenu -WhatIf    # Simula quantas visitas/vídeos novos existem
.\limpar-screenshots.ps1 -WhatIf           # Mostra o ranking de prints por app antes de mover
.\limpar-screenshots.ps1 -LimparAte "2026-09"  # Apaga prints antigos até o mês informado (pede APAGAR)
```
