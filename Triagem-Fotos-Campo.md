# Triagem de Fotos de Campo — Syncthing → OneDrive → Google Drive

**Status**: ideia registrada, projeto **não iniciado** (nenhum script criado ainda)
**Data**: 17/11/2026
**Projeto futuro**: pasta própria, separada deste `sync-onedrive-gdrive-ater` (que só faz o último trecho: OneDrive → Drive)

---

## 1. Objetivo

Eliminar a classificação manual de fotos de campo. O ciclo:

```
Celular (câmera)
   └─ Syncthing (pasta DCIM/Camera, "enviar apenas")
        └─ PC: C:\Users\randomacesso\Syncthing\Camera_Celular
             └─ triagem-campo.ps1  ← classifica por produtor, lendo o relatório do dia
                  └─ OneDrive: ...\Vale do Ribeira\Fernando\<Cidade>\<Atividade>\<AAAA-MM-PRODUTOR>
                       └─ start-sync.ps1 (rclone) → Google Drive da consultoria
```

Por que Syncthing e não API do Google Fotos: **zero limite de API**, arquivo original preservado, **EXIF intacto** (data + GPS). O Syncthing só copia bytes, não mexe em metadado.

A entrada é o **relatório diário do WhatsApp** (texto livre: datas, cidades, produtores, atividade). A IA converte em **JSON padronizado**; o script consome o JSON.

Ganho principal: as fotos chegam ao OneDrive já dentro da pasta certa do produtor, sem ninguém arrastar arquivo.

---

## 2. Papéis de cada etapa

| Etapa | Onde | O que faz |
|---|---|---|
| Captura | Celular | Fotos com tags de localização **ativadas** na câmera |
| Transporte | Syncthing | `DCIM/Camera` → `C:\Users\randomacesso\Syncthing\Camera_Celular` (envio apenas) |
| Parsing | Chat / IA | Texto do WhatsApp → JSON (normaliza cidade, atividade, nome da pasta) |
| Triagem | `triagem-campo.ps1` | Cria pastas, agrupa fotos por data + bloco de horário, distribui por tag (A/B/C), gera `coordenadas_gps.txt` |
| Envio | `start-sync.ps1` | rclone copy para `gdrive_saconsult:VALE DO RIBEIRA/FERNANDO/...` |

---

## 3. Estrutura de pastas

Destino (OneDrive, já usado pelo projeto atual):

```
C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando\
└─ <Cidade>\
   └─ <ATIVIDADE INDIVIDUAL N>\
      └─ <AAAA-MM-NOME DO PRODUTOR>\
         ├─ foto.jpg ...
         └─ coordenadas_gps.txt   ← gerado pelo script
```

Cidades com acentuação oficial do Vale do Ribeira (nome de pasta precisa bater com o OneDrive): `Barra do Turvo`, `Cananéia`, `Jacupiranga`, `Juquiá`, `Pariquera-Açu`, `Registro`.

O relatório do WhatsApp chega sem acento (`Cananeia`) — o parser corrige. É a razão da normalização: nome diferente cria pasta duplicada.

---

## 4. Contrato do JSON

Um objeto por dia, com as visitas e uma **tag** para digitar no terminal.

```json
[
  {
    "data": "2026-11-12",
    "mes_prefixo": "2026-11",
    "visitas": [
      {
        "tag": "A",
        "cidade": "Cananéia",
        "atividade": "ATIVIDADE INDIVIDUAL 2",
        "produtor_pasta": "2026-11-OSMAR - ITAPITANGUI"
      },
      {
        "tag": "B",
        "cidade": "Juquiá",
        "atividade": "ATIVIDADE INDIVIDUAL 2",
        "produtor_pasta": "2026-11-MARIA DAS GRAÇAS"
      }
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
- `data`: `AAAA-MM-DD`
- `mes_prefixo`: `AAAA-MM` (prefixo da pasta do produtor)
- `tag`: `A`, `B`, `C`… sequencial **dentro do dia** (só vale para aquele dia)
- `cidade`: nome oficial acentuado
- `atividade`: caixa alta, `ATIVIDADE INDIVIDUAL 1|2|3…`
- `produtor_pasta`: caixa alta, `AAAA-MM-NOME`, ` - ` no lugar de `/` quando há localidade (`2026-11-OSMAR - ITAPITANGUI`)

---

## 5. Prompt reutilizável (WhatsApp → JSON)

```text
Converta o meu relatório diário de visitas de ATER abaixo em um array JSON válido para o script `triagem-campo.ps1`.

Regras obrigatórias de normalização:
1. Cada bloco "Atividade Realizada no dia DD/MM" vira um objeto com "data": "AAAA-MM-DD" e "mes_prefixo": "AAAA-MM".
2. Atribua sequencialmente "tag": "A", "B", "C", etc. para cada produtor dentro daquele dia.
3. Normalize o nome da cidade com acentuação oficial do Vale do Ribeira:
   "Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro".
4. Normalize a atividade em caixa alta: "ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3", etc.
5. Formate "produtor_pasta" em CAIXA ALTA no padrão "AAAA-MM-NOME DO PRODUTOR"
   (substituindo "/" por " - " se houver localidade, ex: "2026-11-OSMAR - ITAPITANGUI").
6. Retorne APENAS o bloco JSON puro, sem texto adicional, pronto para copiar e colar.
```

---

## 6. Como o script funciona

1. **Entrada**: lê `visitas.json` na pasta do script se existir; senão (ou com `-Colar`) abre colagem no terminal — `Ctrl+V`, digita `FIM` na última linha, Enter.
2. **Pastas**: cria no OneDrive o caminho `Cidade\Atividade\Produtor` que faltar; loga `[CRIADA]` / `[JÁ EXISTE]`.
3. **Indexação**: varre a pasta do Syncthing; data/hora de cada arquivo por:
   1. nome padrão Android/Samsung (`20261112_143022.jpg`);
   2. EXIF `DateTimeOriginal` (0x9003 = 36867) via `System.Drawing`;
   3. fallback `LastWriteTime` (preservado pelo Syncthing).
4. **Agrupamento**: fotos do mesmo dia são cortadas em lotes quando o intervalo entre fotos consecutivas passa de **40 min** (`$IntervaloVisitaMinutos`). O deslocamento entre sítios abre essa janela naturalmente, então cada visita vira um lote único.
5. **Classificação**: para cada lote, mostra o intervalo de horário e os produtores do dia com tag; digitar `A`/`B`/`C` manda o lote inteiro para a pasta. `P` pula, `I` entra em modo foto-a-foto.
6. **GPS**: das fotos do lote copiado, lê `GPSLatitude/Longitude` (tags EXIF 1–4, racionais), converte para graus decimais e grava/append `coordenadas_gps.txt` na pasta do produtor: ponto médio da propriedade, link do Google Maps por foto e por lote.
7. **Cópia**: `Copy-Item`, não `Move-Item` — a pasta do Syncthing é o "negativo" da câmera e não deve ser esvaziada pelo script.

Extensões tratadas: `.jpg .jpeg .png .heic .mp4`.

---

## 7. Script de referência (`triagem-campo.ps1`)

Rascunho validado na conversa original, ainda não salvo no disco. Ao criar o projeto: salvar como `triagem-campo.ps1` e ajustar `$SyncthingFotos`.

```powershell
<#
.SYNOPSIS
    Triagem de Fotos de Campo (Syncthing -> OneDrive ATER) + Extrator GPS EXIF
.DESCRIPTION
    Lê visitas de `visitas.json` (se existir) ou coladas no terminal, cria as pastas
    no OneDrive, agrupa fotos por horário de visita, distribui por TAG (A, B, C...)
    e gera `coordenadas_gps.txt` com Lat/Long decimal e links do Google Maps.
#>
param(
    [switch]$Colar # Força colar no terminal mesmo se existir visitas.json
)

# ================= CONFIGURAÇÕES =================
$OneDriveBase           = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
$SyncthingFotos         = "C:\Users\randomacesso\Syncthing\Camera_Celular" # <-- Ajustar
$ArquivoVisitasJson     = "$PSScriptRoot\visitas.json"
$IntervaloVisitaMinutos = 40
$ExtensoesFoto          = @(".jpg", ".jpeg", ".png", ".heic", ".mp4")

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ================= FUNÇÃO: LER DATA/HORA + GPS EXIF =================
function Get-FotoMetadados {
    param([System.IO.FileInfo]$Arquivo)

    $DataHora = $null
    $Lat = $null
    $Lon = $null

    # 1. Data pelo nome Android/Samsung (ex: 20261112_123456.jpg)
    if ($Arquivo.BaseName -match '(\d{4})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})') {
        try {
            $DataHora = [datetime]::ParseExact("$($Matches[1])-$($Matches[2])-$($Matches[3]) $($Matches[4]):$($Matches[5]):$($Matches[6])", "yyyy-MM-dd HH:mm:ss", $null)
        } catch {}
    }

    # 2. EXIF (Data Original + GPS) para JPG/JPEG
    if ($Arquivo.Extension -match '\.jpe?g$') {
        try {
            $img = [System.Drawing.Image]::FromFile($Arquivo.FullName)

            if (-not $DataHora) {
                $propDt = $img.PropertyItems | Where-Object { $_.Id -eq 36867 } | Select-Object -First 1
                if ($propDt) {
                    $dataStr = [System.Text.Encoding]::ASCII.GetString($propDt.Value).Trim("`0", " ")
                    $DataHora = [datetime]::ParseExact($dataStr, "yyyy:MM:dd HH:mm:ss", $null)
                }
            }

            # 3 racionais EXIF (graus, minutos, segundos) -> decimal
            $ConverterGps = {
                param([byte[]]$Bytes, [string]$Ref)
                if ($Bytes.Length -lt 24) { return $null }
                $gNum = [BitConverter]::ToUInt32($Bytes, 0);  $gDen = [BitConverter]::ToUInt32($Bytes, 4)
                $mNum = [BitConverter]::ToUInt32($Bytes, 8);  $mDen = [BitConverter]::ToUInt32($Bytes, 12)
                $sNum = [BitConverter]::ToUInt32($Bytes, 16); $sDen = [BitConverter]::ToUInt32($Bytes, 20)
                if ($gDen -eq 0 -or $mDen -eq 0 -or $sDen -eq 0) { return $null }

                $dec = ($gNum / $gDen) + (($mNum / $mDen) / 60.0) + (($sNum / $sDen) / 3600.0)
                if ($Ref -in @("S", "W")) { $dec = -$dec }
                return [Math]::Round($dec, 6)
            }

            $propLatRef = $img.PropertyItems | Where-Object { $_.Id -eq 1 } | Select-Object -First 1
            $propLat    = $img.PropertyItems | Where-Object { $_.Id -eq 2 } | Select-Object -First 1
            $propLonRef = $img.PropertyItems | Where-Object { $_.Id -eq 3 } | Select-Object -First 1
            $propLon    = $img.PropertyItems | Where-Object { $_.Id -eq 4 } | Select-Object -First 1

            if ($propLatRef -and $propLat -and $propLonRef -and $propLon) {
                $latRefStr = [System.Text.Encoding]::ASCII.GetString($propLatRef.Value).Trim("`0", " ")
                $lonRefStr = [System.Text.Encoding]::ASCII.GetString($propLonRef.Value).Trim("`0", " ")
                $Lat = & $ConverterGps $propLat.Value $latRefStr
                $Lon = & $ConverterGps $propLon.Value $lonRefStr
            }

            $img.Dispose()
        } catch {}
    }

    if (-not $DataHora) { $DataHora = $Arquivo.LastWriteTime }

    return [PSCustomObject]@{
        Arquivo  = $Arquivo
        DataHora = $DataHora
        DataStr  = $DataHora.ToString("yyyy-MM-dd")
        HoraStr  = $DataHora.ToString("HH:mm:ss")
        Lat      = $Lat
        Lon      = $Lon
        TemGps   = ($null -ne $Lat -and $null -ne $Lon)
    }
}

# ================= FUNÇÃO: SALVAR TXT DE GEOLOCALIZAÇÃO =================
function Export-CoordenadasTxt {
    param(
        [array]$FotosLote,
        [string]$PastaDestino,
        [string]$NomeProdutor
    )

    $TxtPath = Join-Path -Path $PastaDestino -ChildPath "coordenadas_gps.txt"
    $FotosComGps = $FotosLote | Where-Object { $_.TemGps }

    $Linhas = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $TxtPath)) {
        $Linhas.Add("==================================================================")
        $Linhas.Add(" GEOLOCALIZAÇÃO DAS FOTOS DE CAMPO - $NomeProdutor")
        $Linhas.Add(" Gerado em: $(Get-Date -Format 'dd/MM/yyyy HH:mm')")
        $Linhas.Add("==================================================================")
    }

    if ($FotosComGps.Count -gt 0) {
        $MediaLat = [Math]::Round(($FotosComGps | Measure-Object -Property Lat -Average).Average, 6)
        $MediaLon = [Math]::Round(($FotosComGps | Measure-Object -Property Lon -Average).Average, 6)
        $LatStr = $MediaLat.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $LonStr = $MediaLon.ToString([System.Globalization.CultureInfo]::InvariantCulture)

        $Linhas.Add("`n[PONTO MÉDIO DA PROPRIEDADE ($($FotosComGps.Count) fotos com GPS)]")
        $Linhas.Add("Coordenadas (Decimal): $LatStr, $LonStr")
        $Linhas.Add("Google Maps          : https://www.google.com/maps?q=$LatStr,$LonStr")
        $Linhas.Add("------------------------------------------------------------------")
    } else {
        $Linhas.Add("`n[AVISO] Nenhuma foto deste lote continha metadados GPS EXIF.")
    }

    foreach ($f in $FotosLote) {
        if ($f.TemGps) {
            $fLat = $f.Lat.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            $fLon = $f.Lon.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            $Linhas.Add("$($f.DataStr) $($f.HoraStr) | $($f.Arquivo.Name) -> $fLat, $fLon | https://www.google.com/maps?q=$fLat,$fLon")
        } else {
            $Linhas.Add("$($f.DataStr) $($f.HoraStr) | $($f.Arquivo.Name) -> Sem GPS no EXIF")
        }
    }

    Add-Content -Path $TxtPath -Value $Linhas -Encoding UTF8
    return $FotosComGps.Count
}

# ================= 1. CARREGAR JSON (ARQUIVO OU COLADO) =================
Clear-Host
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "   TRIAGEM DE CAMPO - SYNCTHING -> ONEDRIVE (PASTAS + FOTOS + GPS)      " -ForegroundColor White
Write-Host "========================================================================" -ForegroundColor Cyan

$JsonBruto = ""
if ((-not $Colar) -and (Test-Path $ArquivoVisitasJson)) {
    Write-Host " [OK] Lendo arquivo encontrado: $ArquivoVisitasJson" -ForegroundColor Green
    $JsonBruto = Get-Content -Path $ArquivoVisitasJson -Raw -Encoding UTF8
} else {
    Write-Host " Cole o JSON gerado do seu relatório abaixo." -ForegroundColor Yellow
    Write-Host " Ao terminar de colar, digite FIM em uma nova linha e pressione Enter:" -ForegroundColor Gray
    Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan

    $LinhasJson = [System.Collections.Generic.List[string]]::new()
    while ($true) {
        $Linha = Read-Host
        if ($Linha.Trim().ToUpper() -eq "FIM") { break }
        $LinhasJson.Add($Linha)
    }
    $JsonBruto = $LinhasJson -join "`n"
}

try {
    $RelatorioDias = $JsonBruto | ConvertFrom-Json
} catch {
    Write-Host "`n[ERRO] JSON inválido: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ================= 2. CRIAR PASTAS NO ONEDRIVE =================
Write-Host "`n=== ETAPA 1: VERIFICANDO E CRIANDO PASTAS NO ONEDRIVE ===" -ForegroundColor Cyan

foreach ($Dia in $RelatorioDias) {
    foreach ($Visita in $Dia.visitas) {
        $CaminhoDestino = Join-Path -Path $OneDriveBase -ChildPath "$($Visita.cidade)\$($Visita.atividade)\$($Visita.produtor_pasta)"
        $Visita | Add-Member -NotePropertyName "CaminhoDestino" -NotePropertyValue $CaminhoDestino -Force

        if (-not (Test-Path $CaminhoDestino)) {
            New-Item -ItemType Directory -Force -Path $CaminhoDestino | Out-Null
            Write-Host "  [CRIADA]    $($Visita.cidade) > $($Visita.atividade) > $($Visita.produtor_pasta)" -ForegroundColor Green
        } else {
            Write-Host "  [JÁ EXISTE] $($Visita.cidade) > $($Visita.atividade) > $($Visita.produtor_pasta)" -ForegroundColor DarkGray
        }
    }
}

# ================= 3. INDEXAR FOTOS DO SYNCTHING =================
if (-not (Test-Path $SyncthingFotos)) {
    Write-Host "`n[AVISO] Pasta do Syncthing não encontrada em: $SyncthingFotos" -ForegroundColor Yellow
    Write-Host "Pastas criadas no OneDrive. Ajuste `$SyncthingFotos no início do script." -ForegroundColor Yellow
    exit 0
}

Write-Host "`n=== ETAPA 2: LENDO DATA/HORA E GPS DAS FOTOS DO SYNCTHING ===" -ForegroundColor Cyan
$TodasFotos = Get-ChildItem -Path $SyncthingFotos -File | Where-Object { $ExtensoesFoto -contains $_.Extension.ToLower() } | ForEach-Object {
    Get-FotoMetadados -Arquivo $_
} | Sort-Object DataHora

# ================= 4. DISTRIBUIR FOTOS E GERAR GPS TXT =================
foreach ($Dia in $RelatorioDias) {
    $FotosDoDia = $TodasFotos | Where-Object { $_.DataStr -eq $Dia.data }

    Write-Host "`n========================================================================" -ForegroundColor Cyan
    Write-Host " DATA: $($Dia.data) | Fotos encontradas: $($FotosDoDia.Count)" -ForegroundColor White
    Write-Host "========================================================================" -ForegroundColor Cyan

    if ($FotosDoDia.Count -eq 0) {
        Write-Host "  Nenhuma foto encontrada para esta data." -ForegroundColor DarkGray
        continue
    }

    Write-Host " Produtores visitados neste dia:" -ForegroundColor Yellow
    foreach ($v in $Dia.visitas) {
        Write-Host "   [$($v.tag)] $($v.cidade) | $($v.atividade) | $($v.produtor_pasta)" -ForegroundColor Green
    }
    Write-Host "   [P] Pular lote | [I] Classificar foto por foto dentro do lote" -ForegroundColor Gray
    Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan

    # Agrupa fotos próximas no tempo
    $Lotes = [System.Collections.Generic.List[object]]::new()
    $LoteAtual = [System.Collections.Generic.List[object]]::new()

    foreach ($f in $FotosDoDia) {
        if ($LoteAtual.Count -eq 0) {
            $LoteAtual.Add($f)
        } else {
            $DiffMin = ($f.DataHora - $LoteAtual[-1].DataHora).TotalMinutes
            if ($DiffMin -le $IntervaloVisitaMinutos) {
                $LoteAtual.Add($f)
            } else {
                $Lotes.Add($LoteAtual.ToArray())
                $LoteAtual = [System.Collections.Generic.List[object]]::new()
                $LoteAtual.Add($f)
            }
        }
    }
    if ($LoteAtual.Count -gt 0) { $Lotes.Add($LoteAtual.ToArray()) }

    $NumLote = 0
    foreach ($Lote in $Lotes) {
        $NumLote++
        $QtdGps = ($Lote | Where-Object { $_.TemGps }).Count
        Write-Host "`n Lote #$NumLote : $($Lote.Count) foto(s) | Horário: $($Lote[0].HoraStr) às $($Lote[-1].HoraStr) | Com GPS: $QtdGps" -ForegroundColor Cyan

        $Escolha = (Read-Host " Digite a TAG do produtor (A, B, C...) ou P para pular").Trim().ToUpper()

        if ($Escolha -eq "P" -or [string]::IsNullOrWhiteSpace($Escolha)) {
            Write-Host "   -> Lote pulado." -ForegroundColor DarkGray
            continue
        }

        if ($Escolha -eq "I") {
            $PorTag = @{}
            foreach ($FotoItem in $Lote) {
                $TagFoto = (Read-Host "   [$($FotoItem.HoraStr)] $($FotoItem.Arquivo.Name) -> Tag (A/B/P)").Trim().ToUpper()
                $VisitaAlvo = $Dia.visitas | Where-Object { $_.tag -eq $TagFoto } | Select-Object -First 1
                if ($VisitaAlvo) {
                    Copy-Item -Path $FotoItem.Arquivo.FullName -Destination $VisitaAlvo.CaminhoDestino -Force
                    if (-not $PorTag.ContainsKey($TagFoto)) { $PorTag[$TagFoto] = @() }
                    $PorTag[$TagFoto] += $FotoItem
                }
            }
            foreach ($k in $PorTag.Keys) {
                $vAlvo = $Dia.visitas | Where-Object { $_.tag -eq $k } | Select-Object -First 1
                $gpsSalvos = Export-CoordenadasTxt -FotosLote $PorTag[$k] -PastaDestino $vAlvo.CaminhoDestino -NomeProdutor $vAlvo.produtor_pasta
                Write-Host "   [OK] $($PorTag[$k].Count) foto(s) + coordenadas_gps.txt ($gpsSalvos GPS) salvos em $($vAlvo.produtor_pasta)" -ForegroundColor Green
            }
            continue
        }

        $VisitaEscolhida = $Dia.visitas | Where-Object { $_.tag -eq $Escolha } | Select-Object -First 1
        if ($VisitaEscolhida) {
            foreach ($FotoItem in $Lote) {
                Copy-Item -Path $FotoItem.Arquivo.FullName -Destination $VisitaEscolhida.CaminhoDestino -Force
            }
            $gpsSalvos = Export-CoordenadasTxt -FotosLote $Lote -PastaDestino $VisitaEscolhida.CaminhoDestino -NomeProdutor $VisitaEscolhida.produtor_pasta
            Write-Host "   [OK] $($Lote.Count) foto(s) + coordenadas_gps.txt ($gpsSalvos com GPS) salvos em: $($VisitaEscolhida.produtor_pasta)" -ForegroundColor Green
        } else {
            Write-Host "   [!] Tag '$Escolha' inválida. Lote pulado." -ForegroundColor Red
        }
    }
}

Write-Host "`nTriagem finalizada! Rode .\start-sync.ps1 para enviar pastas, fotos e coordenadas_gps.txt ao Google Drive." -ForegroundColor Cyan
```

---

## 8. Uso no dia a dia (4 passos)

1. **Uma vez só**: ajustar `$SyncthingFotos` no topo do script; conferir no celular que a câmera salva localização nas fotos.
2. Colar o relatório do WhatsApp aqui no chat → receber o JSON.
3. Rodar `.\triagem-campo.ps1` → colar JSON (`Ctrl+V`) → `FIM` → Enter. (Ou salvar como `visitas.json` na pasta do script e só rodar; `-Colar` força a colagem.)
4. Para cada lote de horário, digitar `A`/`B`/`C` (ou `I` para foto-a-foto). No fim, `.\start-sync.ps1` sobe tudo para o Drive.

---

## 9. Decisões congeladas

- **Copiar, não mover** — a pasta do Syncthing é o negativo da câmera. Se o disco apertar, aí é hora de pensar em limpeza (com confirmação).
- **Lote por intervalo de tempo (40 min)** em vez de classificar foto por foto — o deslocamento entre propriedades cria a separação naturalmente; 1 tecla classifica ~12 fotos.
- **Sem API do Google Fotos** — EXIF/qualidade original preservados, sem quota.
- **JSON vindo da IA, não do script** — interpretar texto livre de WhatsApp (nomes errados, sem acento, formatos variados) é trabalho de LLM, não de regex frágil em PowerShell.
- **`coordenadas_gps.txt` por produtor** — vai junto com as fotos para a pasta `- RESUMO` no Drive, pronto para colar em QGIS/CAR/laudo.

---

## 10. Pendências / próximos passos (ao retomar)

1. Decidir nome e pasta do novo projeto (ex.: `triagem-campo-ater`), separado deste repo.
2. Salvar `triagem-campo.ps1` e rodar com um dia real do relatório, validando: agrupamento de 40 min, `Copy-Item` e `coordenadas_gps.txt`.
3. **Validar o parser GPS em fotos reais** — a conversão assume racionais little-endian; EXIF big-endian e fotos sem GPS precisam de teste. `System.Drawing` não lê EXIF de `.heic`/`.mp4`: hoje esses caem no fallback de `LastWriteTime` e ficam "sem GPS".
4. Avaliar se vale ir direto com `rclone` (pasta do produtor já pronta) em vez de passar pelo OneDrive — hoje o OneDrive é o intermediário confiável, então fica assim.
5. Teste de estresse: dia com 3 produtores e 60+ fotos; dia em que faltou foto de um produtor.