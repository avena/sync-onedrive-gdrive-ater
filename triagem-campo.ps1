<#
.SYNOPSIS
    Triagem de Fotos de Campo (Syncthing -> OneDrive ATER) + Extrator GPS EXIF
.DESCRIPTION
    Lê visitas de `visitas.json` (se existir) ou coladas no terminal, cria as pastas
    Cidade\Atividade\Produtor no OneDrive, agrupa fotos do Syncthing por data +
    bloco de horário de visita, distribui por TAG (A, B, C...) e gera
    `coordenadas_gps.txt` com Lat/Long decimal e links do Google Maps.
.PARAMETER Colar
    Força a colagem do JSON no terminal mesmo que `visitas.json` exista.
.PARAMETER VisitasJson
    Caminho alternativo para o arquivo de visitas (padrão: visitas.json ao lado do script).
.NOTES
    Requer PowerShell 7. System.Drawing é usado para ler EXIF (JPG/JPEG apenas).
    Fotos .heic/.mp4 caem no fallback de LastWriteTime e ficam "sem GPS".
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [switch]$Colar,
    [string]$VisitasJson = "$PSScriptRoot\visitas.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ================= CONFIGURAÇÕES =================
$OneDriveBase           = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
$SyncthingFotos         = "C:\Users\randomacesso\Syncthing\Camera_Celular" # <-- AJUSTE para o caminho exato da pasta do Syncthing
$IntervaloVisitaMinutos = 40
$ExtensoesFoto          = @(".jpg", ".jpeg", ".png", ".heic", ".mp4")
$LogFile                = Join-Path $PSScriptRoot "triagem-log.txt"

# Garante saída em UTF8 no console (Windows)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ================= FUNÇÃO: VALIDAR CAMPOS DO JSON (anti path-traversal) =================
function Test-CampoSeguro {
    param([string]$Valor)
    if ([string]::IsNullOrWhiteSpace($Valor)) { return $false }
    if ($Valor -match '\.\.' -or $Valor -match '[/\\:]') { return $false }
    return $true
}

# ================= FUNÇÃO: LOG SIMPLES =================
function Write-TriagemLog {
    param([string]$Mensagem)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $LogFile -Value "$timestamp | $Mensagem" -Encoding UTF8
}

# ================= FUNÇÃO: LER DATA/HORA + GPS EXIF =================
function Get-FotoMetadados {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$Arquivo
    )

    $DataHora = $null
    $Lat = $null
    $Lon = $null

    # 1. Data pelo nome do arquivo Android/Samsung (ex: 20261112_143022.jpg ou IMG_20261112_143022.jpg)
    if ($Arquivo.BaseName -match '(\d{4})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})') {
        try {
            $DataHora = [datetime]::ParseExact(
                "$($Matches[1])-$($Matches[2])-$($Matches[3]) $($Matches[4]):$($Matches[5]):$($Matches[6])",
                "yyyy-MM-dd HH:mm:ss",
                $null
            )
        } catch {
            # Ignora falha de parse do nome; tenta EXIF abaixo
        }
    }

    # 2. EXIF (Data Original + GPS) para JPG/JPEG
    if ($Arquivo.Extension -match '\.jpe?g$') {
        $img = $null
        try {
            $img = [System.Drawing.Image]::FromFile($Arquivo.FullName)

            if (-not $DataHora) {
                $propDt = $img.PropertyItems | Where-Object { $_.Id -eq 36867 } | Select-Object -First 1
                if ($propDt) {
                    $dataStr = [System.Text.Encoding]::ASCII.GetString($propDt.Value).Trim("`0", " ")
                    $DataHora = [datetime]::ParseExact($dataStr, "yyyy:MM:dd HH:mm:ss", $null)
                }
            }

            # Converte 3 racionais EXIF (graus, minutos, segundos) em decimal
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
        } catch {
            # Falha na leitura EXIF; mantém valores nulos
        } finally {
            if ($img) { $img.Dispose() }
        }
    }

    # 3. Fallback: LastWriteTime (preservado pelo Syncthing)
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$FotosLote,
        [Parameter(Mandatory = $true)]
        [string]$PastaDestino,
        [Parameter(Mandatory = $true)]
        [string]$NomeProdutor
    )

    $TxtPath = Join-Path -Path $PastaDestino -ChildPath "coordenadas_gps.txt"
    $FotosComGps = @($FotosLote | Where-Object { $_.TemGps })

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
if ((-not $Colar) -and (Test-Path $VisitasJson)) {
    Write-Host " [OK] Lendo arquivo encontrado: $VisitasJson" -ForegroundColor Green
    $JsonBruto = Get-Content -Path $VisitasJson -Raw -Encoding UTF8
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
        # Validação anti path-traversal: rejeita campos com .. / \ :
        if (-not (Test-CampoSeguro $Visita.cidade) -or
            -not (Test-CampoSeguro $Visita.atividade) -or
            -not (Test-CampoSeguro $Visita.produtor_pasta)) {
            Write-Host "  [ERRO] Campos invalidos (path traversal?) em $($Dia.data): $($Visita.produtor_pasta)" -ForegroundColor Red
            Write-TriagemLog "REJEITADO path-traversal: $($Dia.data) / $($Visita.produtor_pasta)"
            continue
        }

        $CaminhoDestino = Join-Path (Join-Path (Join-Path $OneDriveBase $Visita.cidade) $Visita.atividade) $Visita.produtor_pasta
        $Visita | Add-Member -NotePropertyName "CaminhoDestino" -NotePropertyValue $CaminhoDestino -Force

        if (-not (Test-Path $CaminhoDestino)) {
            try {
                New-Item -ItemType Directory -Force -Path $CaminhoDestino | Out-Null
                Write-Host "  [CRIADA]    $($Visita.cidade) > $($Visita.atividade) > $($Visita.produtor_pasta)" -ForegroundColor Green
                Write-TriagemLog "CRIADA: $CaminhoDestino"
            } catch {
                Write-Host "  [ERRO]      Falha ao criar $($Visita.cidade) > $($Visita.atividade) > $($Visita.produtor_pasta) : $($_.Exception.Message)" -ForegroundColor Red
                Write-TriagemLog "ERRO criar: $CaminhoDestino - $($_.Exception.Message)"
            }
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
$ArquivosFoto = @(Get-ChildItem -Path $SyncthingFotos -File |
    Where-Object { $ExtensoesFoto -contains $_.Extension.ToLower() })
$TotalArquivos = $ArquivosFoto.Count
$Idx = 0
$TodasFotos = @($ArquivosFoto | ForEach-Object {
    $Idx++
    Write-Progress -Activity "Indexando fotos" -Status "$Idx / $TotalArquivos" -PercentComplete (($Idx / [Math]::Max($TotalArquivos, 1)) * 100)
    Get-FotoMetadados -Arquivo $_
} | Sort-Object DataHora)
Write-Progress -Activity "Indexando fotos" -Completed

# ================= 4. DISTRIBUIR FOTOS E GERAR GPS TXT =================
foreach ($Dia in $RelatorioDias) {
    $FotosDoDia = @($TodasFotos | Where-Object { $_.DataStr -eq $Dia.data })

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
    Write-Host "   [P] Pular lote | [I] Foto por foto | [Q] Sair da triagem" -ForegroundColor Gray
    Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan

    # Agrupa fotos próximas no tempo (mesma propriedade)
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
    $Escolha = ''
    foreach ($Lote in $Lotes) {
        $NumLote++
        $QtdGps = ($Lote | Where-Object { $_.TemGps }).Count
        Write-Host "`n Lote #$NumLote : $($Lote.Count) foto(s) | Horário: $($Lote[0].HoraStr) às $($Lote[-1].HoraStr) | Com GPS: $QtdGps" -ForegroundColor Cyan

        $Escolha = (Read-Host " Digite a TAG do produtor (A, B, C...), P pular, Q sair").Trim().ToUpper()

        if ($Escolha -eq "Q") {
            Write-Host "   -> Triagem encerrada pelo usuario." -ForegroundColor Yellow
            Write-TriagemLog "ENCERRADO pelo usuario no lote #$NumLote do dia $($Dia.data)"
            break
        }

        if ($Escolha -eq "P" -or [string]::IsNullOrWhiteSpace($Escolha)) {
            Write-Host "   -> Lote pulado." -ForegroundColor DarkGray
            continue
        }

        if ($Escolha -eq "I") {
            # Modo individual: classifica foto por foto dentro do lote
            $PorTag = @{}
            foreach ($FotoItem in $Lote) {
                $TagFoto = (Read-Host "   [$($FotoItem.HoraStr)] $($FotoItem.Arquivo.Name) -> Tag (A/B/P)").Trim().ToUpper()
                if ($TagFoto -eq "P" -or [string]::IsNullOrWhiteSpace($TagFoto)) {
                    continue
                }
                $VisitaAlvo = $Dia.visitas | Where-Object { $_.tag -eq $TagFoto } | Select-Object -First 1
                if ($VisitaAlvo) {
                    try {
                        Copy-Item -Path $FotoItem.Arquivo.FullName -Destination $VisitaAlvo.CaminhoDestino -Force -ErrorAction Stop
                        if (-not $PorTag.ContainsKey($TagFoto)) {
                            $PorTag[$TagFoto] = [System.Collections.Generic.List[object]]::new()
                        }
                        $PorTag[$TagFoto].Add($FotoItem)
                        Write-Host "     Copiada para $($VisitaAlvo.produtor_pasta)" -ForegroundColor Green
                        Write-TriagemLog "COPIADA: $($FotoItem.Arquivo.Name) -> $($VisitaAlvo.produtor_pasta)"
                    } catch {
                        Write-Host "     [ERRO] Falha ao copiar $($FotoItem.Arquivo.Name): $($_.Exception.Message)" -ForegroundColor Red
                        Write-TriagemLog "ERRO copia: $($FotoItem.Arquivo.Name) - $($_.Exception.Message)"
                    }
                }
            }
            foreach ($k in $PorTag.Keys) {
                $vAlvo = $Dia.visitas | Where-Object { $_.tag -eq $k } | Select-Object -First 1
                $gpsSalvos = Export-CoordenadasTxt -FotosLote $PorTag[$k].ToArray() -PastaDestino $vAlvo.CaminhoDestino -NomeProdutor $vAlvo.produtor_pasta
                Write-Host "   [OK] $($PorTag[$k].Count) foto(s) + coordenadas_gps.txt ($gpsSalvos GPS) salvos em $($vAlvo.produtor_pasta)" -ForegroundColor Green
            }
            continue
        }

        $VisitaEscolhida = $Dia.visitas | Where-Object { $_.tag -eq $Escolha } | Select-Object -First 1
        if ($VisitaEscolhida) {
            foreach ($FotoItem in $Lote) {
                try {
                    Copy-Item -Path $FotoItem.Arquivo.FullName -Destination $VisitaEscolhida.CaminhoDestino -Force -ErrorAction Stop
                    Write-TriagemLog "COPIADA: $($FotoItem.Arquivo.Name) -> $($VisitaEscolhida.produtor_pasta)"
                } catch {
                    Write-Host "   [!] Falha ao copiar $($FotoItem.Arquivo.Name): $($_.Exception.Message)" -ForegroundColor Red
                    Write-TriagemLog "ERRO copia: $($FotoItem.Arquivo.Name) - $($_.Exception.Message)"
                }
            }
            $gpsSalvos = Export-CoordenadasTxt -FotosLote $Lote -PastaDestino $VisitaEscolhida.CaminhoDestino -NomeProdutor $VisitaEscolhida.produtor_pasta
            Write-Host "   [OK] $($Lote.Count) foto(s) + coordenadas_gps.txt ($gpsSalvos com GPS) salvos em: $($VisitaEscolhida.produtor_pasta)" -ForegroundColor Green
        } else {
            Write-Host "   [!] Tag '$Escolha' inválida. Lote pulado." -ForegroundColor Red
        }
    }
    if ($Escolha -eq "Q") { break }
}

Write-TriagemLog "TRIAGEM FINALIZADA"
Write-Host "`nTriagem finalizada! Rode .\start-sync.ps1 para enviar pastas, fotos e coordenadas_gps.txt ao Google Drive." -ForegroundColor Cyan