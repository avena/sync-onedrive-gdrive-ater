# Self-check: roda em pasta temporária, sem tocar no OneDrive real nem no ignorados.json
$ErrorActionPreference = 'Stop'
$Script = Join-Path $PSScriptRoot 'start-sync.ps1'

$Erros = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($Script, [ref]$null, [ref]$Erros)
if ($Erros.Count) { throw "Sintaxe: $($Erros[0].Message)" }

# Carrega somente as funções (sem executar menu/modo -Auto)
foreach ($F in $Ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($F.Extent.Text))
}

$Tmp = Join-Path ([IO.Path]::GetTempPath()) ("ater_check_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $Tmp | Out-Null
try {
    $OrigemBase    = Join-Path $Tmp 'Fernando'
    $Cidades       = @('Jacupiranga')
    $Atividades    = @('ATIVIDADE INDIVIDUAL 1', 'ATIVIDADE INDIVIDUAL 2')
    $RemotoDrive   = 'gdrive:TESTE'
    $IgnoradosFile = Join-Path $Tmp 'ignorados.json'

    foreach ($Ativ in $Atividades) {
        foreach ($Nome in @('2026-11-RICARDO ISHIDA', '03 - Rafael Souza')) {
            New-Item -ItemType Directory -Path (Join-Path $OrigemBase "Jacupiranga\$Ativ\$Nome") -Force | Out-Null
        }
    }

    # 1) renomeação no destino
    if ((Get-NomeResumo '2026-11-RICARDO ISHIDA') -ne 'RICARDO ISHIDA - RESUMO') { throw 'renomeação com data falhou' }
    if ((Get-NomeResumo '03 - Rafael Souza') -ne 'Rafael Souza - RESUMO') { throw 'renomeação com número falhou' }
    if ((Get-NomeResumo 'SEM PREFIXO') -ne 'SEM PREFIXO - RESUMO') { throw 'renomeação sem prefixo falhou' }

    # 2) caminho de destino montado
    $P = @(Get-Produtores | Where-Object { $_.Atividade -eq 'ATIVIDADE INDIVIDUAL 1' -and $_.Produtor -eq '2026-11-RICARDO ISHIDA' })[0]
    if ($P.Nuvem -ne 'gdrive:TESTE/ATIVIDADE INDIVIDUAL 1/Jacupiranga/RICARDO ISHIDA - RESUMO') { throw "nuvem errada: $($P.Nuvem)" }

    # 3) ignorar vale só para aquela atividade, não para o nome do produtor
    $Alvo = 'Jacupiranga\ATIVIDADE INDIVIDUAL 1\2026-11-RICARDO ISHIDA'
    Save-Ignorados @($Alvo)
    if (@(Get-Ignorados).Count -ne 1) { throw 'ignorados.json não persistiu' }
    $Ignorados = Get-Ignorados
    $Restantes = @(Get-Produtores | Where-Object { $Ignorados -notcontains $_.Relativo })
    if ($Restantes.Count -ne 3) { throw "esperado 3 pastas sincronizáveis, obtido $($Restantes.Count)" }
    if (-not ($Restantes.Relativo -contains 'Jacupiranga\ATIVIDADE INDIVIDUAL 2\2026-11-RICARDO ISHIDA')) { throw 'ATIVIDADE INDIVIDUAL 2 foi afetada' }

    # 4) ignorar de novo restaura a pasta
    Save-Ignorados (@(Get-Ignorados) | Where-Object { $_ -ne $Alvo })
    if (@(Get-Ignorados).Count -ne 0) { throw 'remoção do ignorado falhou' }

    Write-Host 'OK - self-check passed' -ForegroundColor Green
} finally {
    Remove-Item $Tmp -Recurse -Force -ErrorAction SilentlyContinue
}