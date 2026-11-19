<#
.SYNOPSIS
    Teste piloto controlado do fluxo completo: inbox Drive -> triagem -> OneDrive -> Google Drive.
.DESCRIPTION
    Cria uma visita claramente marcada "TESTE PILOTO" na inbox real (G:\Meu Drive),
    executa as MESMAS funcoes do triagem-campo.ps1 e do start-sync.ps1 (carregadas
    via AST, como no self-check) e valida cada etapa com asserts. A TUI (ReadKey)
    e o unico trecho pulado — este terminal nao e interativo; as chamadas sao as
    mesmas que o fluxo interativo faz apos confirmar com S e aceitar a sugestao.
    Ao final, .\piloto-ponta-a-ponta.ps1 -Limpar remove a visita de teste do
    Google Drive, do OneDrive e da inbox (_processados). Manifestos em runs\ e
    triagem-log.txt sao auditoria permanente e NAO sao apagados.
.PARAMETER Limpar
    Somente limpeza: remove a pasta TESTE PILOTO do Drive, do OneDrive e da inbox.
.EXAMPLE
    pwsh -NoProfile -File .\piloto-ponta-a-ponta.ps1          # executa o piloto
    pwsh -NoProfile -File .\piloto-ponta-a-ponta.ps1 -Limpar  # remove os rastros
#>
#Requires -Version 7.0
[CmdletBinding()]
param([switch]$Limpar)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ===== CONFIG (espelha o topo de triagem-campo.ps1 / start-sync.ps1) =====
$OneDriveBase           = "C:\Users\randomacesso\OneDrive_contaedgesoronote\OneDrive\SA Consultoria\Vale do Ribeira\Fernando"
$SyncthingFotos         = @(
    "C:\Users\randomacesso\Sync\Camera_Celular_Organizada",
    "E:\Sync\Camera_Celular_Organizada",
    "C:\Users\randomacesso\Sync\Camera_Celular\Camera"
)
$InboxBase              = "G:\Meu Drive"
$IntervaloVisitaMinutos = 40
$ExtensoesFoto          = @(".jpg", ".jpeg", ".png", ".heic", ".mp4")
$LogFile                = Join-Path $PSScriptRoot "triagem-log.txt"
$RunsDir                = Join-Path $PSScriptRoot "runs"
$LogDir                 = Join-Path $PSScriptRoot "logs"
$NomeRemoto             = "gdrive_saconsult"
$RemotoDrive            = "${NomeRemoto}:VALE DO RIBEIRA/FERNANDO"
$IgnoradosFile          = Join-Path $PSScriptRoot "ignorados.json"
$Cidades                = @("Barra do Turvo", "Cananéia", "Jacupiranga", "Juquia", "Pariquera-Açu", "Registro")
$CidadesRemoto          = @{ 'Pariquera-Açu' = 'PARIQUERA AÇU'; 'Juquia' = 'Juquiá' }   # nome local -> nome no destino [GDSA]
$Atividades             = @("ATIVIDADE INDIVIDUAL 1", "ATIVIDADE INDIVIDUAL 2", "ATIVIDADE INDIVIDUAL 3",
                            "ATIVIDADE INDIVIDUAL 4", "ATIVIDADE INDIVIDUAL 5", "ATIVIDADE INDIVIDUAL 6", "PROJETO PRODUTIVO")
$RcloneFlags            = @("--update", "--exclude", "*.json", "--transfers=4", "--checkers=8",
                            "--drive-shared-with-me", "--stats-one-line", "--stats=2s",
                            "--retries=3", "--retries-sleep=10s", "--contimeout=30s", "-v")
$Script:DriveDirCache   = @{}

# ===== CARREGA SOMENTE AS FUNCOES DOS SCRIPTS REAIS (tecnica do self-check) =====
foreach ($arq in @('triagem-campo.ps1', 'start-sync.ps1')) {
    $erros = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $arq), [ref]$null, [ref]$erros)
    if ($erros.Count) { throw "Sintaxe ${arq}: $($erros[0].Message)" }
    foreach ($f in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        . ([scriptblock]::Create($f.Extent.Text))
    }
}

# ===== ALVO DO PILOTO =====
$DataPiloto   = (Get-Date -Format 'yyyy-MM-dd')
$CidadePiloto = 'Jacupiranga'
$AtivPiloto   = 'ATIVIDADE INDIVIDUAL 1'
$LeafPiloto   = "$DataPiloto-TESTE PILOTO - RESUMO"   # inbox: formato do app (AAAA-MM-DD-)
$LeafOD       = "$($DataPiloto.Substring(8, 2))-$($DataPiloto.Substring(0, 7))-TESTE PILOTO - RESUMO"   # [OD]/[GDSA]: DD-AAAA-MM-
$NomeDrive    = $LeafOD
$InboxLeaf    = Join-Path $InboxBase "$(Get-Date -Format 'yyyy')\_inbox_relatorio_ater\$(Get-Date -Format 'MM')\$CidadePiloto\$AtivPiloto\$LeafPiloto"
$OneDriveLeaf = Join-Path $OneDriveBase "$CidadePiloto\$AtivPiloto\$LeafOD"
$NuvemPiloto  = "$RemotoDrive/$AtivPiloto/$(Get-NomeCidadeRemoto $CidadePiloto)/$NomeDrive"

function Assert([bool]$Cond, [string]$Msg) {
    if (-not $Cond) { throw "[FALHOU] $Msg" }
    Write-Host "  [OK] $Msg" -ForegroundColor Green
}

# Remove do _processados.json do ano os registros da visita TESTE PILOTO (casa pasta_inbox OU
# destino_onedrive, sem barra final, case-insensitive). Preserva os demais registros; gravacao
# atomica (.tmp + Move). Retorna quantos removeu. Nunca toca em runs\ nem triagem-log.txt.
function Remove-ProcessadosPiloto {
    param([string]$PastaInboxAno, [string]$PastaInbox, [string]$DestinoOD)
    $regPath = Join-Path $PastaInboxAno '_processados.json'
    if (-not (Test-Path -LiteralPath $regPath)) { return 0 }
    try { $regs = @(Get-Content -LiteralPath $regPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return 0 }
    $ficam = @($regs | Where-Object {
        (([string](Get-PropValue $_ 'pasta_inbox')).TrimEnd('\') -ine $PastaInbox.TrimEnd('\')) -and
        (([string](Get-PropValue $_ 'destino_onedrive')).TrimEnd('\') -ine $DestinoOD.TrimEnd('\'))
    })
    $removidos = $regs.Count - $ficam.Count
    if ($removidos -gt 0) {
        $json = if ($ficam.Count -eq 0) { '[]' } else { ConvertTo-Json -InputObject @($ficam) -Depth 10 }
        $tmp = "$regPath.tmp"
        Set-Content -LiteralPath $tmp -Value $json -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $regPath -Force
    }
    return $removidos
}

# ================= LIMPEZA (-Limpar) =================
if ($Limpar) {
    Write-Host "=== LIMPEZA DO PILOTO (TESTE PILOTO) ===" -ForegroundColor Cyan
    if (Test-Path -LiteralPath $OneDriveLeaf) {
        Remove-Item -LiteralPath $OneDriveLeaf -Recurse -Force
        Write-Host "  [REMOVIDO] OneDrive: $OneDriveLeaf" -ForegroundColor Yellow
    }
    $procLeaf = Join-Path $InboxBase "$(Get-Date -Format 'yyyy')\_inbox_relatorio_ater\_processados\$(Get-Date -Format 'MM')\$CidadePiloto\$AtivPiloto\$LeafPiloto"
    foreach ($p in @($InboxLeaf, $procLeaf)) {
        if (Test-Path -LiteralPath $p) {
            Remove-Item -LiteralPath $p -Recurse -Force
            Write-Host "  [REMOVIDO] inbox: $p" -ForegroundColor Yellow
        }
    }
    # registro da visita de teste no _processados.json do ano (a Etapa 1 registra; sem isso
    # sobraria entrada orfa TESTE PILOTO). Demais registros do ano ficam intactos.
    $nReg = Remove-ProcessadosPiloto -PastaInboxAno (Join-Path $InboxBase "$(Get-Date -Format 'yyyy')\_inbox_relatorio_ater") -PastaInbox $InboxLeaf -DestinoOD $OneDriveLeaf
    if ($nReg -gt 0) { Write-Host "  [REMOVIDO] $nReg registro(s) TESTE PILOTO do _processados.json" -ForegroundColor Yellow }
    # remove pastas vazias criadas pelo piloto (sobe ate G:\Meu Drive, nunca remove a raiz)
    $dir = Split-Path $InboxLeaf -Parent
    while ($dir.StartsWith($InboxBase) -and $dir -ne $InboxBase -and (Test-Path -LiteralPath $dir) -and
           @(Get-ChildItem -LiteralPath $dir -Force).Count -eq 0) {
        Remove-Item -LiteralPath $dir -Force
        Write-Host "  [REMOVIDA] pasta vazia: $dir" -ForegroundColor DarkGray
        $dir = Split-Path $dir -Parent
    }
    Write-Host "  Limpando no Google Drive: $NuvemPiloto" -ForegroundColor Cyan
    & rclone purge $NuvemPiloto --drive-shared-with-me --quiet
    Assert ($LASTEXITCODE -eq 0) 'pasta TESTE PILOTO removida do Google Drive'
    Write-Host "[OK] Limpeza concluida. Manifestos runs\ e triagem-log.txt preservados (auditoria)." -ForegroundColor Green
    return
}

# ================= GUARDAS =================
Write-Host "=== PILOTO PONTA A PONTA - $DataPiloto ===" -ForegroundColor Cyan
if ((Test-Path -LiteralPath $InboxLeaf) -or (Test-Path -LiteralPath $OneDriveLeaf)) {
    throw "Ja existe visita TESTE PILOTO (inbox ou OneDrive). Rode -Limpar antes de repetir."
}
Assert ((rclone listremotes --quiet) -contains "${NomeRemoto}:") "remote $NomeRemoto configurado"

# ================= ETAPA 0: visita de teste na inbox real =================
Write-Host "`n[0] Criando visita de teste na inbox do Drive pessoal..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $InboxLeaf | Out-Null
$yaml = @"
---
data: $DataPiloto
periodo: manha
hora_inicio: "09:00"
hora_fim: "11:30"
cidade: $CidadePiloto
atividade: $AtivPiloto
produtor_pasta: $(Get-Date -Format 'yyyy-MM')-TESTE PILOTO
---
RELATORIO DE TESTE PILOTO - visita simulada para validar o fluxo ponta a ponta.
Este arquivo nao e um relatorio real de atendimento.
"@
Set-Content -LiteralPath (Join-Path $InboxLeaf 'relatorio.txt') -Value $yaml -Encoding UTF8
$json = [PSCustomObject]@{
    schema_version = '1.0'
    metadados = [PSCustomObject]@{
        data = $DataPiloto; periodo = 'manha'; hora_inicio = '09:00'; hora_fim = '11:30'
        cidade = $CidadePiloto; atividade = $AtivPiloto
        produtor_pasta = "$(Get-Date -Format 'yyyy-MM')-TESTE PILOTO"
    }
}
$json | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $InboxLeaf 'relatorio.json') -Encoding UTF8
Assert (Test-Path -LiteralPath (Join-Path $InboxLeaf 'relatorio.json')) 'inbox criada (.txt + .json)'

# ================= ETAPA 1: triagem (funcoes reais, sem a TUI) =================
Write-Host "`n[1] Get-InboxVisitas + normalizacao (Etapa 1 do triagem)..." -ForegroundColor Cyan
$inbox = Get-InboxVisitas
$VisitasPiloto = @($inbox | Where-Object { $_.produtor_pasta -like '*TESTE PILOTO*' })
Assert ($VisitasPiloto.Count -eq 1) "inbox detectou a visita de teste (total na inbox: $($inbox.Count))"
# piloto controlado: segue somente com a visita de teste (equivale a remover as demais na TUI)
$RelatorioDias = ConvertTo-RelatorioDias -Visitas $VisitasPiloto
Assert ($RelatorioDias[0].visitas[0].tag -eq 'A') 'tag A atribuida'
Assert ($RelatorioDias[0].visitas[0].produtor_pasta -eq $LeafOD) "nome normalizado: $LeafOD"

$PastasCriadas = Invoke-CriarPastasEArquivar -RelatorioDias $RelatorioDias
Assert (Test-Path -LiteralPath $OneDriveLeaf) 'pasta criada no OneDrive'
$txtDest = Join-Path $OneDriveLeaf 'relatorio.txt'
Assert ((Test-Path -LiteralPath $txtDest) -and ((Get-Content -LiteralPath $txtDest -Raw) -notmatch '(?s)^\s*---')) '.txt limpo gravado (sem YAML)'
Assert (-not (Test-Path -LiteralPath (Join-Path $OneDriveLeaf 'relatorio.json'))) '.json NAO vai para o OneDrive (fica so na inbox)'
Assert ((Test-Path -LiteralPath (Join-Path $InboxLeaf 'relatorio.txt')) -and (Test-Path -LiteralPath (Join-Path $InboxLeaf 'relatorio.json'))) 'originais intactos na inbox (.txt + .json)'

# ================= ETAPA 2: fotos do dia (indexacao real das raizes) =================
Write-Host "`n[2] Indexando fotos do Syncthing (raizes reais, dedup por nome)..." -ForegroundColor Cyan
$RaizesExistentes = @($SyncthingFotos | Where-Object { Test-Path $_ })
Assert ($RaizesExistentes.Count -gt 0) "raizes existentes: $($RaizesExistentes -join ' | ')"
$ArquivosFoto = [System.Collections.Generic.List[object]]::new()
$NomesVistos = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($raiz in $RaizesExistentes) {
    foreach ($f in @(Get-ChildItem -LiteralPath $raiz -Recurse -File | Where-Object { $ExtensoesFoto -contains $_.Extension.ToLower() -and $_.Name -notlike '.*' -and $_.Name -notlike '~syncthing~*' })) {
        if ($NomesVistos.Add($f.Name)) { $ArquivosFoto.Add($f) }
    }
}
$TodasFotos = @($ArquivosFoto | ForEach-Object { Get-FotoMetadados -Arquivo $_ } | Sort-Object DataHora)
Write-Host "  $($TodasFotos.Count) foto(s) indexadas." -ForegroundColor Gray
$FotosDoDia = @($TodasFotos | Where-Object { $_.DataStr -eq $DataPiloto })
Assert ($FotosDoDia.Count -ge 1) "$($FotosDoDia.Count) foto(s) encontradas para $DataPiloto"

# lotes de 40 min (mesma logica da Etapa 2)
$Lotes = [System.Collections.Generic.List[object]]::new()
$LoteAtual = [System.Collections.Generic.List[object]]::new()
foreach ($f in $FotosDoDia) {
    if ($LoteAtual.Count -eq 0) { $LoteAtual.Add($f) }
    elseif (($f.DataHora - $LoteAtual[-1].DataHora).TotalMinutes -le $IntervaloVisitaMinutos) { $LoteAtual.Add($f) }
    else { $Lotes.Add($LoteAtual.ToArray()); $LoteAtual = [System.Collections.Generic.List[object]]::new(); $LoteAtual.Add($f) }
}
if ($LoteAtual.Count -gt 0) { $Lotes.Add($LoteAtual.ToArray()) }
$Sugestao = Get-SugestaoTag -Lote $Lotes[0] -Visitas $RelatorioDias[0].visitas
Assert ($Sugestao -eq 'A') "pre-sugestao por horario = A (lote $($Lotes[0][0].HoraStr)-$($Lotes[0][-1].HoraStr))"

# copia + GPS (equivale a Enter na sugestao)
$V = $RelatorioDias[0].visitas[0]
$RunFotos = [System.Collections.Generic.List[object]]::new()
$RunGps = [System.Collections.Generic.List[object]]::new()
foreach ($FotoItem in $Lotes[0]) {
    Copy-Item -LiteralPath $FotoItem.Arquivo.FullName -Destination $V.CaminhoDestino -Force
    $RunFotos.Add([PSCustomObject]@{ origem = $FotoItem.Arquivo.FullName; destino = (Join-Path $V.CaminhoDestino $FotoItem.Arquivo.Name); produtor = $V.produtor_pasta })
    Write-TriagemLog "PILOTO COPIADA: $($FotoItem.Arquivo.Name) -> $($V.produtor_pasta)"
}
$gps = Export-CoordenadasTxt -FotosLote $Lotes[0] -PastaDestino $V.CaminhoDestino -NomeProdutor $V.produtor_pasta
$RunGps.Add([PSCustomObject]@{ produtor = $V.produtor_pasta; count = $gps.Count; lat = $gps.MediaLat; lon = $gps.MediaLon; txt = $gps.TxtPath })
Assert (Test-Path -LiteralPath (Join-Path $OneDriveLeaf $FotosDoDia[0].Arquivo.Name)) 'foto copiada ao OneDrive'
Assert (Test-Path -LiteralPath $gps.TxtPath) "coordenadas_gps.txt gerado ($($gps.Count) foto(s) com GPS)"

$RunOrigens = @([PSCustomObject]@{ data = $DataPiloto; produtor = $V.produtor_pasta; txt = $V.SourceTxt; json = $V.SourceJson })
Write-TriagemRun -Registro ([PSCustomObject]@{
    timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); origens = $RunOrigens
    pastas_criadas = $PastasCriadas; fotos_copiadas = $RunFotos; gps = $RunGps
    contagem = [PSCustomObject]@{ visitas = 1; pastas = $PastasCriadas.Count; fotos = $RunFotos.Count }
})

# ================= ETAPA 3: start-sync (dry-run -> real -> manifesto URL) =================
Write-Host "`n[3] start-sync: dry-run da pasta de teste..." -ForegroundColor Cyan
Assert ((Get-NomeResumo $LeafPiloto) -eq $NomeDrive) "nome no Drive: $NomeDrive (data DD-AAAA-MM-, mesmo nome local, sem duplicar RESUMO)"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }
$LogFilePiloto = Join-Path $LogDir ("sync_piloto_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
$Dry = Invoke-RcloneCopy -Local $OneDriveLeaf -Nuvem $NuvemPiloto -Simulacao $true -LogFile $LogFilePiloto
Assert ($Dry.Codigo -eq 0 -and $Dry.Pendentes -ge 2) "dry-run: $($Dry.Pendentes) arquivo(s) pendente(s), 0 erros"

Write-Host "`n[4] start-sync: copia REAL (rclone copy --update --exclude *.json)..." -ForegroundColor Cyan
$Real = Invoke-RcloneCopy -Local $OneDriveLeaf -Nuvem $NuvemPiloto -Simulacao $false -LogFile $LogFilePiloto
Assert ($Real.Codigo -eq 0 -and $Real.Copiados -ge 2) "copia real: $($Real.Copiados) novo(s), $($Real.Atualizados) atualizado(s), $($Real.Erros) erro(s)"
Write-SyncRun -Escopo 'PILOTO' -Modificados @([PSCustomObject]@{
    Produtor = $LeafPiloto; Cidade = $CidadePiloto; Atividade = $AtivPiloto
    NomeDrive = $NomeDrive; Local = $OneDriveLeaf
    Copiados = $Real.Copiados; Atualizados = $Real.Atualizados
}) -Simulacao $false

# ================= ETAPA 4: verificacao remota =================
Write-Host "`n[5] Verificando o destino no Google Drive..." -ForegroundColor Cyan
$remotos = @(rclone lsf $NuvemPiloto --drive-shared-with-me --quiet)
Assert ($remotos -contains 'relatorio.txt') 'relatorio.txt presente no Drive'
Assert ($remotos -contains 'coordenadas_gps.txt') 'coordenadas_gps.txt presente no Drive'
Assert ($remotos -notcontains 'relatorio.json') 'relatorio.json NAO subiu (--exclude *.json)'
Assert (@($remotos | Where-Object { $_ -like '*.jpg' }).Count -ge 1) 'foto presente no Drive'
$Url = Get-DriveFolderUrl -Atividade $AtivPiloto -Cidade $CidadePiloto -NomeDrive $NomeDrive
Assert ($null -ne $Url) "URL clicavel obtida: $Url"

Write-Host "`n=== PILOTO PASSOU - fluxo inbox -> triagem -> OneDrive -> Drive validado ===" -ForegroundColor Green
Write-Host "Para remover os rastros do teste: pwsh -NoProfile -File .\piloto-ponta-a-ponta.ps1 -Limpar" -ForegroundColor Yellow
