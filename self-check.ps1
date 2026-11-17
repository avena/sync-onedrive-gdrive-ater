# Self-check: roda em pasta temporária, sem tocar no OneDrive real nem no ignorados.json
$ErrorActionPreference = 'Stop'
$Script        = Join-Path $PSScriptRoot 'start-sync.ps1'
$ScriptTriagem = Join-Path $PSScriptRoot 'triagem-campo.ps1'

$Erros = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($Script, [ref]$null, [ref]$Erros)
if ($Erros.Count) { throw "Sintaxe: $($Erros[0].Message)" }

# Carrega somente as funções (sem executar menu/modo -Auto)
foreach ($F in $Ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($F.Extent.Text))
}

# Triagem: valida sintaxe e carrega só as funções (sem executar o fluxo interativo)
$ErrosT = $null
$AstT = [System.Management.Automation.Language.Parser]::ParseFile($ScriptTriagem, [ref]$null, [ref]$ErrosT)
if ($ErrosT.Count) { throw "Sintaxe triagem: $($ErrosT[0].Message)" }
foreach ($F in $AstT.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($F.Extent.Text))
}

# Screenshots: valida sintaxe e carrega só as funções (sem executar a migração)
$ScriptShots = Join-Path $PSScriptRoot 'limpar-screenshots.ps1'
$ErrosS = $null
$AstS = [System.Management.Automation.Language.Parser]::ParseFile($ScriptShots, [ref]$null, [ref]$ErrosS)
if ($ErrosS.Count) { throw "Sintaxe screenshots: $($ErrosS[0].Message)" }
foreach ($F in $AstS.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
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

    # 1b) idempotência do - RESUMO no start-sync (não duplicar)
    if ((Get-NomeResumo '2026-11-OSMAR - ITAPITANGUI - RESUMO') -ne 'OSMAR - ITAPITANGUI - RESUMO') { throw 'idempotência: duplicou RESUMO com prefixo de data' }
    if ((Get-NomeResumo 'OSMAR - ITAPITANGUI - RESUMO') -ne 'OSMAR - ITAPITANGUI - RESUMO') { throw 'idempotência: duplicou RESUMO sem prefixo' }
    if ((Get-NomeResumo 'JOAO - resumo') -ne 'JOAO - resumo') { throw 'idempotência case-insensitive falhou' }

    # 1c) Add-SufixoResumo da triagem (idempotente, case-insensitive)
    if ((Add-SufixoResumo '2026-11-OSMAR - ITAPITANGUI') -ne '2026-11-OSMAR - ITAPITANGUI - RESUMO') { throw 'triagem: sufixo não adicionado' }
    if ((Add-SufixoResumo '2026-11-OSMAR - ITAPITANGUI - RESUMO') -ne '2026-11-OSMAR - ITAPITANGUI - RESUMO') { throw 'triagem: duplicou RESUMO' }
    if ((Add-SufixoResumo 'NOME - RESUMO') -ne 'NOME - RESUMO') { throw 'triagem: idempotência falhou' }

    # 1d) Format-NomePastaProdutor: prefixo AAAA-MM-DD- (dia completo) + sufixo - RESUMO
    if ((Format-NomePastaProdutor -Nome 'OSMAR - ITAPITANGUI' -Data '2026-11-12') -ne '2026-11-12-OSMAR - ITAPITANGUI - RESUMO') { throw 'triagem: prefixo+sufixo não garantidos' }
    if ((Format-NomePastaProdutor -Nome '2026-11-12-OSMAR - ITAPITANGUI - RESUMO' -Data '2026-11-12') -ne '2026-11-12-OSMAR - ITAPITANGUI - RESUMO') { throw 'triagem: duplicou prefixo/sufixo' }
    if ((Format-NomePastaProdutor -Nome '2026-11-OSMAR' -Data '2026-11-12') -ne '2026-11-12-OSMAR - RESUMO') { throw 'triagem: prefixo parcial não normalizado' }

    # 1e) start-sync remove prefixo de data COMPLETA (AAAA-MM-DD-) tambem
    if ((Get-NomeResumo '2026-11-12-OSMAR - ITAPITANGUI - RESUMO') -ne 'OSMAR - ITAPITANGUI - RESUMO') { throw 'start-sync: não removeu dia completo' }

    # 1f) Remove-YamlHeader / Read-MetadadosTxt
    $yamlTxt = "---`ndata: 2026-11-12`nperiodo: manha`n---`nCorpo do relatorio"
    if ((Remove-YamlHeader $yamlTxt) -ne 'Corpo do relatorio') { throw 'triagem: YAML não removido' }
    $txtFile = Join-Path $Tmp 'rel.txt'
    Set-Content -Path $txtFile -Value $yamlTxt -Encoding UTF8
    $mt = Read-MetadadosTxt -Path $txtFile
    if ($mt.Campos['data'] -ne '2026-11-12' -or $mt.Campos['periodo'] -ne 'manha') { throw 'triagem: YAML parse falhou' }
    if ($mt.Body.Trim() -ne 'Corpo do relatorio') { throw 'triagem: corpo YAML falhou' }

    # 1g) ConvertTo-Minutos / Get-SugestaoTag (periodo e horario)
    if ((ConvertTo-Minutos '09:30') -ne 570) { throw 'triagem: ConvertTo-Minutos falhou' }
    $vA = [PSCustomObject]@{ tag = 'A'; periodo = 'manha'; hora_inicio = $null; hora_fim = $null }
    $vB = [PSCustomObject]@{ tag = 'B'; periodo = 'tarde'; hora_inicio = $null; hora_fim = $null }
    $loteManha = @([PSCustomObject]@{ DataHora = [datetime]'2026-11-12 09:00:00' }, [PSCustomObject]@{ DataHora = [datetime]'2026-11-12 09:30:00' })
    if ((Get-SugestaoTag -Lote $loteManha -Visitas @($vA, $vB)) -ne 'A') { throw 'triagem: sugestão por período falhou' }
    $vC = [PSCustomObject]@{ tag = 'C'; periodo = $null; hora_inicio = '09:00'; hora_fim = '10:00' }
    if ((Get-SugestaoTag -Lote $loteManha -Visitas @($vC)) -ne 'C') { throw 'triagem: sugestão por horário falhou' }
    if ($null -ne (Get-SugestaoTag -Lote $loteManha -Visitas @($vA, $vC))) { throw 'triagem: sugestão ambígua deveria ser nula' }

    # 1h) ConvertTo-RelatorioDias: tags A/B e prefixo de dia completo
    $va = New-Visita -Data '2026-11-12' -Cidade 'Juquiá' -Atividade 'ATIVIDADE INDIVIDUAL 1' -ProdutorPasta 'OSMAR'
    $vb = New-Visita -Data '2026-11-12' -Cidade 'Juquiá' -Atividade 'ATIVIDADE INDIVIDUAL 1' -ProdutorPasta 'MARIA'
    $dias = ConvertTo-RelatorioDias -Visitas @($va, $vb)
    if ($dias[0].visitas[0].tag -ne 'A') { throw 'triagem: tag A não atribuída' }
    if ($dias[0].visitas[1].produtor_pasta -ne '2026-11-12-MARIA - RESUMO') { throw 'triagem: nome de pasta não normalizado' }

    # 1i) limpar-screenshots: classificação por app, data e soma segura
    $AppsAmigaveis = @{ 'com.whatsapp' = 'WhatsApp'; 'com.android.chrome' = 'Chrome'; 'com.facebook.katana' = 'Facebook' }
    $PrefixosGenericos = @('com', 'org', 'net', 'br', 'lt', 'eu', 'air', 'jus', 'gov')
    if ((Get-NomePastaApp 'com.whatsapp') -ne 'WhatsApp') { throw 'screenshots: dicionário falhou' }
    if ((Get-NomePastaApp 'com.android.chrome-edit') -ne 'Chrome') { throw 'screenshots: marcador -edit não removido' }
    if ((Get-NomePastaApp 'com.dimep.kairosemployee.br') -ne 'Dimep') { throw 'screenshots: fallback de prefixos falhou' }
    if ((Get-NomePastaApp $null) -ne 'Outros') { throw 'screenshots: sem app deveria ser Outros' }
    if ((Get-SomaBytes @()) -ne 0) { throw 'screenshots: soma de lista vazia deveria ser 0' }
    $fakePrint = New-Item -ItemType File -Path (Join-Path $Tmp 'Screenshot_2026-11-17-13-16-32-698_com.whatsapp__2.jpg')
    if ((Get-AppDoPrint $fakePrint) -ne 'com.whatsapp') { throw 'screenshots: extração de app (colisão __2) falhou' }
    $dataPrint = Get-DataPrint $fakePrint
    if ($dataPrint.Ano -ne '2026' -or $dataPrint.Mes -ne '11') { throw 'screenshots: data do nome falhou' }

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