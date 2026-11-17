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

# Listas oficiais para a revisao interativa (batem com start-sync.ps1)
$Cidades = @(
    "Barra do Turvo", "Cananéia", "Jacupiranga", "Juquiá", "Pariquera-Açu", "Registro"
)
$Atividades = @(
    "ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3",
    "ATIVIDADE INDIVIDUAL 4", "ATIVIDADE INDIVIDUAL 5", "ATIVIDADE INDIVIDUAL 6",
    "PROJETO PRODUTIVO"
)

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

# ================= FUNÇÃO: DRENAR TECLAS PENDENTES =================
function Clear-KeyBuffer {
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch { }
}

# ================= FUNÇÃO: GARANTIR SUFIXO - RESUMO (idempotente) =================
function Add-SufixoResumo {
    param([string]$Nome)
    if ([string]::IsNullOrWhiteSpace($Nome)) { return $Nome }
    # -match e case-insensitive por padrao: nao duplica se ja terminar em - RESUMO
    if ($Nome -match '\s*-\s*RESUMO$') { return $Nome }
    return "$Nome - RESUMO"
}

# ================= FUNÇÃO: NOME DA PASTA = PREFIXO AAAA-MM- + NOME + - RESUMO =================
# A pasta usa APENAS ano/mes (nunca o dia). O dia exato fica no campo "data" do JSON,
# que e usado somente para buscar as fotos no Syncthing.
function Format-NomePastaProdutor {
    param([string]$Nome, [string]$MesPrefixo)
    if ([string]::IsNullOrWhiteSpace($Nome)) { return $Nome }
    $Nome = $Nome.Trim()
    # Garante o prefixo AAAA-MM- (se o usuario digitou so o nome, prepende o mes)
    if ($Nome -notmatch '^\d{4}-\d{2}-') {
        $Nome = "$MesPrefixo-$Nome"
    }
    return Add-SufixoResumo $Nome
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

# ================= FUNÇÃO: EDITAR UMA VISITA (data / cidade / atividade / nome) =================
# Recebe o item da revisao (Dia + Visita). A data editada e a do DIA (campo "data"),
# usada exclusivamente para buscar as fotos; o nome da pasta usa so ano/mes.
function Edit-VisitaItem {
    param($Item)
    $Dia    = $Item.Dia
    $Visita = $Item.Visita

    # 1. Data da visita AAAA-MM-DD (usada para localizar as fotos do dia no Syncthing)
    Write-Host "`n  Data atual da visita (busca de fotos): $($Dia.data)" -ForegroundColor DarkGray
    $R = Read-Host "  Nova data AAAA-MM-DD (Enter mantem)"
    if (-not [string]::IsNullOrWhiteSpace($R)) {
        $NovaData = $R.Trim()
        if ($NovaData -match '^\d{4}-\d{2}-\d{2}$') {
            $Dia.data = $NovaData
        } else {
            Write-Host "  [!] Data invalida (use AAAA-MM-DD). Mantida a anterior." -ForegroundColor Red
            Start-Sleep -Milliseconds 1200
        }
    }

    # 2. Cidade (lista numerada; Enter mantem a atual)
    Write-Host "`n  Cidades:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Cidades.Count; $i++) { Write-Host "   [$($i+1)] $($Cidades[$i])" -ForegroundColor White }
    Write-Host "   Atual: $($Visita.cidade)" -ForegroundColor DarkGray
    $R = Read-Host "  Numero da cidade (Enter mantem)"
    if ($R -match '^\d+$' -and [int]$R -ge 1 -and [int]$R -le $Cidades.Count) {
        $Visita.cidade = $Cidades[[int]$R - 1]
    }

    # 3. Atividade (lista numerada; Enter mantem a atual)
    Write-Host "`n  Atividades:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Atividades.Count; $i++) { Write-Host "   [$($i+1)] $($Atividades[$i])" -ForegroundColor White }
    Write-Host "   Atual: $($Visita.atividade)" -ForegroundColor DarkGray
    $R = Read-Host "  Numero da atividade (Enter mantem)"
    if ($R -match '^\d+$' -and [int]$R -ge 1 -and [int]$R -le $Atividades.Count) {
        $Visita.atividade = $Atividades[[int]$R - 1]
    }

    # 4. Nome da pasta do produtor (Enter mantem; garante prefixo AAAA-MM- e sufixo - RESUMO)
    # MesPrefixo vem de mes_prefixo do JSON; senao deriva dos 7 primeiros chars da data.
    $MesPrefixo = $Dia.data.Substring(0, 7)
    $PropMes = $Dia.PSObject.Properties['mes_prefixo']
    if ($PropMes -and $PropMes.Value) { $MesPrefixo = [string]$PropMes.Value }

    Write-Host "`n  Nome atual da pasta: $($Visita.produtor_pasta)" -ForegroundColor DarkGray
    $R = Read-Host "  Novo nome (Enter mantem)"
    if (-not [string]::IsNullOrWhiteSpace($R)) {
        $Novo = $R.Trim()
        if (Test-CampoSeguro $Novo) {
            $Visita.produtor_pasta = Format-NomePastaProdutor -Nome $Novo -MesPrefixo $MesPrefixo
        } else {
            Write-Host "  [!] Nome invalido (caracteres de caminho). Mantido o anterior." -ForegroundColor Red
            Start-Sleep -Milliseconds 1200
        }
    }
}

# ================= FUNÇÃO: REVISÃO INTERATIVA (TUI com setas) =================
# Retorna $true se o usuario confirmar a criacao; $false se cancelar/sair.
# Edita/remove itens no proprio $RelatorioDias (referencias), antes de criar pastas.
function Show-RevisaoVisitas {
    param([array]$RelatorioDias)

    $Itens = [System.Collections.Generic.List[object]]::new()
    foreach ($Dia in $RelatorioDias) {
        foreach ($Visita in $Dia.visitas) {
            $Itens.Add([PSCustomObject]@{ Dia = $Dia; Visita = $Visita })
        }
    }

    if ($Itens.Count -eq 0) {
        Write-Host "`n[AVISO] Nenhuma visita no JSON para revisar." -ForegroundColor Yellow
        return $false
    }

    $Atual  = 0
    $Offset = 0
    $Altura = 15
    try { $Altura = [Math]::Max(6, $Host.UI.RawUI.WindowSize.Height - 12) } catch { }
    try { [Console]::CursorVisible = $false } catch { }

    try {
        while ($true) {
            if ($Itens.Count -eq 0) {
                Write-Host "`n[AVISO] Lista vazia (todos os itens foram removidos). Nada a criar." -ForegroundColor Yellow
                return $false
            }
            if ($Atual -gt ($Itens.Count - 1)) { $Atual = $Itens.Count - 1 }
            if ($Atual -lt 0) { $Atual = 0 }

            Clear-Host
            Write-Host ("=" * 74) -ForegroundColor Cyan
            Write-Host " REVISAO DAS VISITAS - antes de criar pastas no OneDrive    Total: $($Itens.Count)" -ForegroundColor White
            Write-Host ("=" * 74) -ForegroundColor Cyan
            Write-Host " [Setas] Navegar | [E] Editar | [D/Del] Remover | [C/Enter] Confirmar | [Q/Esc] Cancelar" -ForegroundColor Yellow
            Write-Host ("-" * 74) -ForegroundColor Cyan

            if ($Atual -lt $Offset) { $Offset = $Atual }
            if ($Atual -ge ($Offset + $Altura)) { $Offset = $Atual - $Altura + 1 }
            $Fim = [Math]::Min($Itens.Count - 1, $Offset + $Altura - 1)

            for ($i = $Offset; $i -le $Fim; $i++) {
                $It = $Itens[$i]
                $V  = $It.Visita
                $Linha = " [$($It.Dia.data)] $($V.cidade) > $($V.atividade) > $($V.produtor_pasta)"
                if ($i -eq $Atual) { Write-Host $Linha -ForegroundColor Black -BackgroundColor Gray }
                else { Write-Host $Linha -ForegroundColor White }
            }

            Write-Host ("-" * 74) -ForegroundColor Cyan
            $Tecla = [Console]::ReadKey($true)

            switch ($Tecla.Key) {
                "UpArrow"   { if ($Atual -gt 0) { $Atual-- } }
                "DownArrow" { if ($Atual -lt ($Itens.Count - 1)) { $Atual++ } }
                "E" {
                    try { [Console]::CursorVisible = $true } catch { }
                    Clear-Host
                    Write-Host "=== EDITAR VISITA ===" -ForegroundColor Cyan
                    Edit-VisitaItem -Item $Itens[$Atual]
                    Clear-KeyBuffer
                    try { [Console]::CursorVisible = $false } catch { }
                }
                { $_ -in "D", "Delete" } {
                    $Removido = $Itens[$Atual]
                    $Itens.RemoveAt($Atual)
                    Write-TriagemLog "REMOVIDO na revisao: $($Removido.Dia.data) / $($Removido.Visita.produtor_pasta)"
                    if ($Atual -ge $Itens.Count -and $Atual -gt 0) { $Atual-- }
                }
                { $_ -in "C", "Enter" } {
                    try { [Console]::CursorVisible = $true } catch { }
                    Write-Host ""
                    $Resp = Read-Host "Deseja criar todas essas pastas no OneDrive? (S/N)"
                    Clear-KeyBuffer
                    if ($Resp -in @('s', 'S', 'y', 'Y')) {
                        # Reconstroi as visitas de cada dia apenas com os itens sobreviventes
                        foreach ($Dia in $RelatorioDias) {
                            $Dia.visitas = @($Itens |
                                Where-Object { [object]::ReferenceEquals($_.Dia, $Dia) } |
                                ForEach-Object { $_.Visita })
                        }
                        return $true
                    }
                    try { [Console]::CursorVisible = $false } catch { }
                }
                { $_ -in "Q", "Escape" } { return $false }
            }
        }
    } finally {
        try { [Console]::CursorVisible = $true } catch { }
        Clear-KeyBuffer
    }
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

# SEPARACAO EXPLICITA DAS DATAS DO JSON:
#   - "data" (AAAA-MM-DD) = dia exato da visita -> usado SO para filtrar/localizar as
#     fotos daquele dia no Syncthing (EXIF DateTimeOriginal, nome AAAAMMDD ou LastWriteTime).
#   - "produtor_pasta" (AAAA-MM-NOME - RESUMO) = usa APENAS ano/mes, nunca o dia.
# Normaliza o sufixo - RESUMO em todas as visitas (idempotente, case-insensitive).
foreach ($Dia in $RelatorioDias) {
    foreach ($Visita in $Dia.visitas) {
        $Visita.produtor_pasta = Add-SufixoResumo $Visita.produtor_pasta
    }
}

# ================= ETAPA 1a: REVISÃO INTERATIVA (antes de criar qualquer pasta) =================
$Confirmado = Show-RevisaoVisitas -RelatorioDias $RelatorioDias
if (-not $Confirmado) {
    Write-Host "`n[CANCELADO] Nenhuma pasta foi criada no OneDrive." -ForegroundColor Yellow
    Write-TriagemLog "REVISAO CANCELADA - nenhuma pasta criada"
    exit 0
}

# ================= ETAPA 1b: CRIAR PASTAS NO ONEDRIVE (confirmado) =================
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
    # Filtra as fotos pelo campo "data" (AAAA-MM-DD) do dia — nao pelo prefixo da pasta.
    # DataStr vem do EXIF DateTimeOriginal, do padrao AAAAMMDD do nome ou do LastWriteTime.
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