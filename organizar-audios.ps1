<#
.SYNOPSIS
    Move audios soltos na raiz de GravadorDeConversa (Syncthing) para pastas AAAA-MM.
.DESCRIPTION
    Origem: C:\Users\randomacesso\Sync\GravadorDeConversa (bidirecional com o
    celular HyperOS — a reorganizacao reflete na pasta GRAVADORDECONVERSA).
    Regras:
      1. So arquivos de audio (.m4a/.mp3/.aac/.wav) SOLTOS NA RAIZ; subpastas,
         ocultos (.stfolder/.stignore) e temporarios do Syncthing sao ignorados.
      2. Mes pelo nome do Gravador Xiaomi ("3 de nov., 11.39 ...m4a") + ano do
         LastWriteTime -> pasta AAAA-MM (ex.: 2026-11).
      3. Nome fora do padrao Xiaomi -> fallback: LastWriteTime (ano e mes).
      4. Cria AAAA-MM se nao existir; nunca sobrescreve (conflito de nome = aviso).
.EXAMPLE
    .\organizar-audios.ps1 -WhatIf   # simulacao (nao move nada)
    .\organizar-audios.ps1           # organiza de verdade
.NOTES
    Compativel com Windows PowerShell 5.1 e PowerShell 7.
    -LiteralPath em tudo: nomes do Xiaomi podem conter caracteres especiais.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param()

$ErrorActionPreference = 'Stop'

$PastaRaiz      = Join-Path $env:USERPROFILE 'Sync\GravadorDeConversa'
$ExtensoesAudio = @('.m4a', '.mp3', '.aac', '.wav')

function Get-PastaDestinoAudio {
    # Mes pelo padrao Xiaomi "DD de mmm., ..."; ano e fallback pelo LastWriteTime.
    param([string]$Nome, [datetime]$LastWriteTime)
    $MapaMeses = @{
        'jan' = '01'; 'fev' = '02'; 'mar' = '03'; 'abr' = '04'
        'mai' = '05'; 'jun' = '06'; 'jul' = '07'; 'ago' = '08'
        'set' = '09'; 'out' = '10'; 'nov' = '11'; 'dez' = '12'
    }
    $mes = $null
    if ($Nome -match '^\d{1,2}\s+de\s+([a-zA-ZçÇ]{3})\.') {
        $abrev = $Matches[1].ToLower()
        if ($MapaMeses.ContainsKey($abrev)) { $mes = $MapaMeses[$abrev] }
    }
    if (-not $mes) { $mes = $LastWriteTime.ToString('MM') }
    return '{0}-{1}' -f $LastWriteTime.ToString('yyyy'), $mes
}

if (-not (Test-Path -LiteralPath $PastaRaiz)) { throw "Pasta nao encontrada: $PastaRaiz" }

# So a raiz (sem -Recurse); sem -Force ja exclui ocultos (.stfolder etc.)
$arquivos = @(Get-ChildItem -LiteralPath $PastaRaiz -File | Where-Object {
    $_.Extension.ToLower() -in $ExtensoesAudio -and
    $_.Name -notlike '.*' -and
    $_.Name -notlike '~syncthing~*' -and
    $_.Name -notlike '*.tmp'
})

if ($arquivos.Count -eq 0) {
    Write-Host 'Nenhum audio solto na raiz para organizar.' -ForegroundColor Yellow
    return
}

$movidos = 0; $ignorados = 0
foreach ($arq in $arquivos) {
    $pastaNome = Get-PastaDestinoAudio -Nome $arq.Name -LastWriteTime $arq.LastWriteTime
    $pasta     = Join-Path $PastaRaiz $pastaNome
    $destino   = Join-Path $pasta $arq.Name

    if (Test-Path -LiteralPath $destino) {
        Write-Warning "Ignorado (ja existe em $pastaNome): $($arq.Name)"
        $ignorados++
        continue
    }
    if ($PSCmdlet.ShouldProcess($arq.Name, "Mover para $pastaNome")) {
        if (-not (Test-Path -LiteralPath $pasta)) {
            New-Item -ItemType Directory -Path $pasta | Out-Null
            Write-Host "Pasta criada: $pastaNome" -ForegroundColor Cyan
        }
        Move-Item -LiteralPath $arq.FullName -Destination $destino
        Write-Host "Movido: $($arq.Name) -> $pastaNome\" -ForegroundColor Green
        $movidos++
    }
}

Write-Host "`nConcluido: $movidos movido(s), $ignorados ignorado(s) por conflito." -ForegroundColor Cyan
