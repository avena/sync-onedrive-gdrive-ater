<#
.SYNOPSIS
    Organiza prints do Android por APLICATIVO na pasta nativa Screenshots e
    remove prints de lockscreen automaticamente.
.DESCRIPTION
    Varre C:\Users\randomacesso\Sync\Camera_Celular (incluindo o layout antigo
    Screenshots\AAAA\MM) e MOVE cada print para ...\Screenshots\<App>\.
    Nenhuma pasta e criada fora de Camera_Celular. Fotos reais (IMG_* com
    EXIF de camera ou 4:3) nunca sao movidas.

    Deteccao de print (qualquer criterio): 1) dentro de *\Screenshots\*;
    2) nome comeca com Screenshot_; 3) .png na raiz/Camera/Restored;
    4) resolucao exata 1080x2400/2400x1080 sem EXIF Make/Model (nunca IMG_*).

    App = pacote apos o carimbo de data/hora (Screenshot_AAAA-MM-DD-HH-mm-ss-fff_<APP>
    ou Screenshot_AAAAMMDD_HHMMSS_<APP>), ignorando colisao __N e o marcador
    -edit (Samsung). Conhecidos -> nomes amigaveis; desconhecidos -> remove
    prefixos genericos (com., br.com., org., lt., eu., air., jus., ...) e o
    token .android, capitaliza o 1o segmento (ex: Dimep). Sem app -> Outros.
    lockscreen/com.android.systemui -> REMOVIDOS automaticamente, salvo com
    -WhatIf ou -ManterLockscreen (vao para Lockscreen\).

    Ignorados: ocultos (.stfolder, .android, .deleteRecord, .globalTrash,
    .tmfs, .trashed-*), ~syncthing~*.tmp e prints ja em Screenshots\<App>\
    (idempotente). Ao final, remove pastas vazias do layout antigo AAAA\MM.
    Manifesto: runs\AAAA-MM-DD_HH-mm-ss__SCREENSHOTS__limpar-screenshots.json.
.PARAMETER LimparAte
    Corte AAAA-MM. Lista prints de TODAS as pastas de app cuja data do nome
    (fallback LastWriteTime) seja <= ao corte; exige APAGAR e remove
    (Syncthing espelha no celular). Pastas vazias sao removidas ao final.
.PARAMETER ManterLockscreen
    Mantem prints de lockscreen, organizando-os em Lockscreen\.
.EXAMPLE
    .\limpar-screenshots.ps1 -WhatIf               # ranking por app (simulacao)
    .\limpar-screenshots.ps1                       # organiza + remove lockscreen
    .\limpar-screenshots.ps1 -ManterLockscreen     # mantem lockscreen
    .\limpar-screenshots.ps1 -LimparAte "2026-09"  # apaga prints ate set/2026
.NOTES
    Requer PowerShell 7. System.Drawing le resolucao/EXIF de JPG/PNG.
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')]
    [string]$LimparAte,
    [switch]$ManterLockscreen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ================= CONFIGURACOES =================
$OrigemSync  = "C:\Users\randomacesso\Sync\Camera_Celular"
$DestinoBase = "C:\Users\randomacesso\Sync\Camera_Celular\Screenshots"   # pasta nativa; so cria <App>\ dentro dela
$RunsDir     = Join-Path $PSScriptRoot "runs"   # permanente: NUNCA apagar

$AppsAmigaveis = @{
    'com.facebook.katana'          = 'Facebook'
    'com.facebook.lite'            = 'Facebook'
    'com.linkedin.android'         = 'LinkedIn'
    'com.mercadolibre'             = 'Mercado Livre'
    'com.whatsapp'                 = 'WhatsApp'
    'com.whatsapp.w4b'             = 'WhatsApp'
    'com.android.chrome'           = 'Chrome'
    'com.instagram.android'        = 'Instagram'
    'com.google.android.youtube'   = 'YouTube'
    'com.sec.android.app.launcher' = 'Tela Inicial'
}
$AppsLockscreen    = @('lockscreen', 'com.android.systemui')
$PrefixosGenericos = @('com', 'org', 'net', 'br', 'lt', 'eu', 'air', 'jus', 'gov')

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ================= HELPERS =================
function Get-DataPrint {
    # Data real do print: AAAAMMDD ou AAAA-MM-DD no nome; fallback LastWriteTime.
    param([System.IO.FileInfo]$Arquivo)
    foreach ($padrao in @(
        '(?<!\d)(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)',
        '(?<!\d)(\d{4})-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])(?!\d)'
    )) {
        $m = [regex]::Match($Arquivo.Name, $padrao)
        if ($m.Success) { return @{ Ano = $m.Groups[1].Value; Mes = $m.Groups[2].Value } }
    }
    return @{ Ano = $Arquivo.LastWriteTime.ToString('yyyy'); Mes = $Arquivo.LastWriteTime.ToString('MM') }
}

function Get-AppDoPrint {
    # Pacote apos o carimbo de data/hora; ignora colisao __N. $null se sem sufixo.
    param([System.IO.FileInfo]$Arquivo)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Arquivo.Name) -replace '__\d+$', ''
    $m = [regex]::Match($base, '^Screenshot_[\d_-]+_([A-Za-z].+)$')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Get-NomePastaApp {
    # Nome amigavel sanitizado p/ Windows; desconhecido -> 1o segmento capitalizado.
    param([string]$AppId)
    if (-not $AppId) { return 'Outros' }
    $id = $AppId.ToLower() -replace '-edit$', ''   # marcador de print editado (Samsung)
    if ($AppsAmigaveis.ContainsKey($id)) { return $AppsAmigaveis[$id] }
    $seg = @($id.Split('.') | Where-Object { $_ -and $_ -ne 'android' })
    $i = 0
    while ($i -lt $seg.Count -and $PrefixosGenericos -contains $seg[$i]) { $i++ }
    if ($i -ge $seg.Count) { return 'Outros' }
    $nome = $seg[$i]
    $nome = $nome.Substring(0, 1).ToUpper() + $nome.Substring(1)
    $nome = ($nome -replace '[\\/:*?"<>|]', '').TrimEnd('.', ' ')
    if (-not $nome) { return 'Outros' }
    return $nome
}

function Test-ExifCamera {
    param($Img)
    foreach ($id in 271, 272) {
        $p = $Img.PropertyItems | Where-Object { $_.Id -eq $id } | Select-Object -First 1
        if ($p -and $p.Value -and $p.Value.Length -gt 0) {
            $s = [System.Text.Encoding]::ASCII.GetString($p.Value).Trim([char]0).Trim()
            if ($s) { return $true }
        }
    }
    return $false
}

function Test-EhScreenshot {
    # Retorna o criterio atendido (string) ou $null se nao for print.
    param([System.IO.FileInfo]$Arquivo, [string]$DirRelativo)
    if ($DirRelativo -match '(^|\\)Screenshots(\\|$)') { return 'pasta Screenshots' }
    if ($Arquivo.Name -like 'Screenshot_*') { return 'nome Screenshot_' }
    if ($Arquivo.Extension -ieq '.png' -and ($DirRelativo -eq '' -or $DirRelativo -eq 'Camera' -or $DirRelativo -eq 'Restored')) {
        return 'PNG na raiz/Camera/Restored'
    }
    $ext = $Arquivo.Extension.ToLower()
    if ($ext -in '.jpg', '.jpeg', '.png' -and $Arquivo.Name -notlike 'IMG_*') {
        $img = $null
        try {
            $img = [System.Drawing.Image]::FromFile($Arquivo.FullName)
            $tela = ($img.Width -eq 1080 -and $img.Height -eq 2400) -or ($img.Width -eq 2400 -and $img.Height -eq 1080)
            if ($tela -and -not (Test-ExifCamera $img)) { return 'resolucao de tela sem EXIF de camera' }
        } catch { } finally { if ($img) { $img.Dispose() } }
    }
    return $null
}

function Get-NomeSemColisao {
    param([string]$Pasta, [System.IO.FileInfo]$Arquivo)
    $dest = Join-Path $Pasta $Arquivo.Name
    if (-not (Test-Path -LiteralPath $dest)) { return $dest }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Arquivo.Name)
    $i = 2
    do {
        $dest = Join-Path $Pasta ("{0}__{1}{2}" -f $base, $i, $Arquivo.Extension)
        $i++
    } while (Test-Path -LiteralPath $dest)
    return $dest
}

function Remove-PastasVazias {
    # Meses/anos vazios do layout antigo + pastas de app que ficarem vazias.
    foreach ($anoDir in @(Get-ChildItem -LiteralPath $DestinoBase -Directory | Where-Object { $_.Name -match '^\d{4}$' })) {
        foreach ($mesDir in @(Get-ChildItem -LiteralPath $anoDir.FullName -Directory | Where-Object { $_.Name -match '^(0[1-9]|1[0-2])$' })) {
            if (@(Get-ChildItem -LiteralPath $mesDir.FullName -Force).Count -eq 0) {
                Remove-Item -LiteralPath $mesDir.FullName -Recurse -Force
                Write-Host "  [REMOVIDA] pasta vazia: $($anoDir.Name)\$($mesDir.Name)"
            }
        }
        if (@(Get-ChildItem -LiteralPath $anoDir.FullName -Force).Count -eq 0) {
            Remove-Item -LiteralPath $anoDir.FullName -Recurse -Force
            Write-Host "  [REMOVIDA] pasta de ano vazia: $($anoDir.Name)"
        }
    }
    foreach ($d in @(Get-ChildItem -LiteralPath $DestinoBase -Directory)) {
        if (@(Get-ChildItem -LiteralPath $d.FullName -Force).Count -eq 0) {
            Remove-Item -LiteralPath $d.FullName -Recurse -Force
            Write-Host "  [REMOVIDA] pasta vazia: $($d.Name)"
        }
    }
}

function Get-SomaBytes {
    # Measure-Object devolve $null com entrada vazia (StrictMode quebraria o .Sum).
    param($Itens)
    $m = $Itens | Measure-Object -Property Bytes -Sum
    if ($m) { return [double]$m.Sum }
    return 0
}

function Show-RankingPorApp {
    param($Itens, [string]$Titulo)
    Write-Host "`n$Titulo" -ForegroundColor White
    $grupos = @($Itens | Group-Object Pasta | Sort-Object Count -Descending)
    if ($grupos.Count -eq 0) { Write-Host "  (nenhum)" -ForegroundColor DarkGray; return }
    foreach ($g in $grupos) {
        $bytes = Get-SomaBytes $g.Group
        Write-Host ("  {0,-20} {1,4} print(s)  {2,8:N1} MB" -f $g.Name, $g.Count, ($bytes / 1MB))
    }
    $totalBytes = Get-SomaBytes $Itens
    Write-Host ("  TOTAL: {0} print(s), {1:N1} MB" -f @($Itens).Count, ($totalBytes / 1MB)) -ForegroundColor Cyan
}

# ================= MODO LIMPEZA (-LimparAte AAAA-MM, por data do nome) =================
function Invoke-Limpeza {
    param([string]$Corte)
    Write-Host "LIMPEZA DE PRINTS ATE $Corte (inclusive) em $DestinoBase" -ForegroundColor White
    if (-not (Test-Path -LiteralPath $DestinoBase)) {
        Write-Host "[OK] Nada a limpar: $DestinoBase nao existe." -ForegroundColor Green
        return
    }
    $arqs = @(Get-ChildItem -LiteralPath $DestinoBase -Recurse -File | Where-Object {
        $rel = $_.FullName.Substring($DestinoBase.Length + 1)
        $_.Name -notlike '~syncthing~*' -and $_.Extension -ne '.tmp' -and $rel -notmatch '(^|\\)\.'
    })
    $alvos = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $arqs) {
        $data = Get-DataPrint -Arquivo $f
        $chave = "$($data.Ano)-$($data.Mes)"
        if ($chave -le $Corte) {
            $pastaRel = Split-Path $f.FullName.Substring($DestinoBase.Length + 1) -Parent
            $alvos.Add([PSCustomObject]@{ Chave = $chave; Pasta = $pastaRel; Arquivo = $f; Bytes = $f.Length })
        }
    }
    if ($alvos.Count -eq 0) {
        Write-Host "[OK] Nenhum print com data <= $Corte." -ForegroundColor Green
        return
    }
    Write-Host "`nPrints que serao REMOVIDOS (por pasta de aplicativo):" -ForegroundColor Yellow
    foreach ($g in @($alvos | Group-Object Pasta | Sort-Object Name)) {
        $bytes = Get-SomaBytes $g.Group
        Write-Host ("  {0,-20} {1,4} arquivo(s)  {2,8:N1} MB" -f $g.Name, $g.Count, ($bytes / 1MB))
    }
    $totMB = (Get-SomaBytes $alvos) / 1MB
    Write-Host ("  TOTAL: {0} arquivo(s), {1:N1} MB" -f $alvos.Count, $totMB) -ForegroundColor Yellow
    Write-Host "`n[ATENCAO] O Syncthing espelhará esta remoção no celular (os prints somem de lá também)." -ForegroundColor Red

    if ($WhatIfPreference) {
        Write-Host "[SIMULAÇÃO] Nada foi removido (-WhatIf)." -ForegroundColor Cyan
        return
    }
    $resp = Read-Host "Digite APAGAR para confirmar"
    if ($resp -cne 'APAGAR') {
        Write-Host "[CANCELADO] Confirmacao incorreta. Nada foi removido." -ForegroundColor Green
        return
    }
    foreach ($a in $alvos) { Remove-Item -LiteralPath $a.Arquivo.FullName -Force }
    Remove-PastasVazias
    Write-Host "`n[OK] Limpeza concluida: $($alvos.Count) arquivo(s), $([math]::Round($totMB,1)) MB liberados." -ForegroundColor Green
}

# ================= MODO PADRAO (ORGANIZACAO POR APLICATIVO) =================
function Invoke-Migracao {
    Write-Host "ORGANIZAR PRINTS POR APP: $OrigemSync -> $DestinoBase\<App>" -ForegroundColor White
    if (-not (Test-Path -LiteralPath $OrigemSync)) { throw "Pasta de origem nao encontrada: $OrigemSync" }

    # Destino DENTRO da origem: pular o que ja esta em <App>\ (idempotente),
    # mas incluir o layout antigo <AAAA>\<MM> e prints soltos na raiz.
    $destinoRel = $DestinoBase.Substring($OrigemSync.Length).TrimStart('\')
    $padraoOrganizados = '^' + [regex]::Escape($destinoRel) + '\\(?!\d{4}\\)[^\\]+\\'
    $todos = @(Get-ChildItem -LiteralPath $OrigemSync -Recurse -File | Where-Object {
        $rel = $_.FullName.Substring($OrigemSync.Length + 1)
        $_.Name -notlike '~syncthing~*' -and $_.Extension -ne '.tmp' -and
        $rel -notmatch '(^|\\)\.' -and          # ocultos: .stfolder, .android, .deleteRecord, .globalTrash, .tmfs, .trashed-*
        $rel -notmatch $padraoOrganizados
    })
    Write-Host "Varrendo $($todos.Count) arquivo(s) (ignorando ocultos/.trashed-*, ~syncthing~*.tmp e ja organizados por app)..."

    $detectados = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($f in $todos) {
        $i++
        Write-Progress -Activity 'Detectando capturas de tela' -Status $f.Name -PercentComplete (($i / $todos.Count) * 100)
        $dirRel = $f.DirectoryName.Substring($OrigemSync.Length).TrimStart('\')
        $criterio = Test-EhScreenshot -Arquivo $f -DirRelativo $dirRel
        if ($criterio) {
            $app = Get-AppDoPrint -Arquivo $f
            $ehLock = $app -and ($AppsLockscreen -contains $app.ToLower())
            $detectados.Add([PSCustomObject]@{
                Arquivo = $f; Criterio = $criterio; App = $app; EhLock = [bool]$ehLock
                Pasta = $(if ($ehLock) { 'Lockscreen' } else { Get-NomePastaApp -AppId $app })
                Bytes = $f.Length
            })
        }
    }
    Write-Progress -Activity 'Detectando capturas de tela' -Completed

    $locks = @($detectados | Where-Object { $_.EhLock })
    $paraMover = @(if ($ManterLockscreen) { $detectados } else { $detectados | Where-Object { -not $_.EhLock } })
    $lockBytes = Get-SomaBytes $locks

    Show-RankingPorApp -Itens $paraMover -Titulo 'RANKING DE PASTAS POR APLICATIVO:'
    if ($locks.Count -gt 0) {
        $acao = if ($ManterLockscreen) { 'serão MANTIDOS em Lockscreen\' } else { 'serão REMOVIDOS automaticamente' }
        Write-Host ("`n  LOCKSCREEN: {0} print(s), {1:N1} MB {2}." -f $locks.Count, ($lockBytes / 1MB), $acao) -ForegroundColor $(if ($ManterLockscreen) { 'Yellow' } else { 'Red' })
    }
    if ($detectados.Count -eq 0) { Write-Host "[OK] Nada a fazer." -ForegroundColor Green; return }

    if ($WhatIfPreference) {
        Write-Host "`n[SIMULAÇÃO] $($paraMover.Count) print(s) seriam movidos, $($locks.Count) lockscreen(s) $(if ($ManterLockscreen) { 'mantidos' } else { 'removidos' }). Nada foi alterado (-WhatIf)." -ForegroundColor Cyan
        return
    }

    # 1) Remocao automatica de lockscreen (salvo -ManterLockscreen)
    $removidos = [System.Collections.Generic.List[object]]::new()
    if (-not $ManterLockscreen) {
        foreach ($d in $locks) {
            Remove-Item -LiteralPath $d.Arquivo.FullName -Force
            $removidos.Add([PSCustomObject]@{ origem = $d.Arquivo.FullName; bytes = $d.Bytes })
        }
        if ($locks.Count -gt 0) {
            Write-Host ("[LIMPO] {0} lockscreen(s) removidos, {1:N1} MB liberados." -f $locks.Count, ($lockBytes / 1MB)) -ForegroundColor Red
        }
    }

    # 2) Migracao para Screenshots\<App>\
    $movidos = [System.Collections.Generic.List[object]]::new()
    foreach ($d in $paraMover) {
        $pasta = Join-Path $DestinoBase $d.Pasta
        if (-not (Test-Path -LiteralPath $pasta)) { New-Item -ItemType Directory -Force -Path $pasta | Out-Null }
        $dest = Get-NomeSemColisao -Pasta $pasta -Arquivo $d.Arquivo
        Move-Item -LiteralPath $d.Arquivo.FullName -Destination $dest
        $movidos.Add([PSCustomObject]@{
            origem = $d.Arquivo.FullName; destino = $dest; aplicativo = $d.App; pasta = $d.Pasta
            criterio = $d.Criterio; bytes = $d.Bytes
        })
    }

    # 3) Pastas vazias do layout antigo (AAAA\MM) e apps vazios
    Remove-PastasVazias

    # 4) Manifesto de auditoria permanente em runs\
    if (-not (Test-Path -LiteralPath $RunsDir)) { New-Item -ItemType Directory -Force -Path $RunsDir | Out-Null }
    $ts = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $runJson = Join-Path $RunsDir "${ts}__SCREENSHOTS__limpar-screenshots.json"
    $porApp = @($movidos | Group-Object pasta | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{
            pasta = $_.Name; arquivos = $_.Count
            bytes = Get-SomaBytes $_.Group
        }
    })
    $registro = [PSCustomObject]@{
        script = 'limpar-screenshots'; modo = 'organizacao_por_aplicativo'; timestamp = $ts
        origem = $OrigemSync; destino = $DestinoBase; manter_lockscreen = [bool]$ManterLockscreen
        totais = [PSCustomObject]@{
            movidos = $movidos.Count
            bytes_movidos = Get-SomaBytes $movidos
            lockscreen_removidos = $removidos.Count
            bytes_lockscreen_removidos = Get-SomaBytes $removidos
        }
        por_aplicativo = $porApp; lockscreen_removidos = $removidos; arquivos = $movidos
    }
    $registro | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $runJson -Encoding UTF8

    Show-RankingPorApp -Itens $paraMover -Titulo 'PRINTS ORGANIZADOS POR APLICATIVO:'
    Write-Host "`n[OK] $($movidos.Count) print(s) movidos para $DestinoBase\<App>" -ForegroundColor Green
    Write-Host "[OK] Manifesto de auditoria salvo em: $runJson" -ForegroundColor Cyan
}

# ================= ENTRADA =================
if ($LimparAte) { Invoke-Limpeza -Corte $LimparAte } else { Invoke-Migracao }
