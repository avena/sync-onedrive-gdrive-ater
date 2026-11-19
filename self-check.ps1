# Self-check: roda em pasta temporária, sem tocar no OneDrive real nem no ignorados.json
$ErrorActionPreference = 'Stop'
Write-Host '[SELF-CHECK] Ambiente temporario isolado — nenhuma pasta real ([GD], [OD], [GDSA] ou Syncthing) sera alterada.' -ForegroundColor Cyan
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

# Camera: valida sintaxe e carrega só as funções (sem executar o fluxo campo/casa)
$ScriptCamera = Join-Path $PSScriptRoot 'organizar-camera.ps1'
$ErrosC = $null
$AstC = [System.Management.Automation.Language.Parser]::ParseFile($ScriptCamera, [ref]$null, [ref]$ErrosC)
if ($ErrosC.Count) { throw "Sintaxe camera: $($ErrosC[0].Message)" }
foreach ($F in $AstC.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($F.Extent.Text))
}

# Audios: valida sintaxe e carrega só as funções (sem executar a organização)
$ScriptAudios = Join-Path $PSScriptRoot 'organizar-audios.ps1'
$ErrosA = $null
$AstA = [System.Management.Automation.Language.Parser]::ParseFile($ScriptAudios, [ref]$null, [ref]$ErrosA)
if ($ErrosA.Count) { throw "Sintaxe audios: $($ErrosA[0].Message)" }
foreach ($F in $AstA.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($F.Extent.Text))
}

# Piloto ponta a ponta: valida sintaxe (6º script; não executa o fluxo real). Carrega SOMENTE
# a função de limpeza do _processados.json (nome único, sem colisão com os demais scripts).
$ScriptPiloto = Join-Path $PSScriptRoot 'piloto-ponta-a-ponta.ps1'
$ErrosP = $null
$AstP = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPiloto, [ref]$null, [ref]$ErrosP)
if ($ErrosP.Count) { throw "Sintaxe piloto: $($ErrosP[0].Message)" }
foreach ($F in $AstP.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Remove-ProcessadosPiloto' }, $true)) {
    . ([scriptblock]::Create($F.Extent.Text))
}

$Tmp = Join-Path ([IO.Path]::GetTempPath()) ("ater_check_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $Tmp | Out-Null
$Passou = $false
try {
    $OrigemBase    = Join-Path $Tmp 'Fernando'
    $Cidades       = @('Jacupiranga')
    $Atividades    = @('ATIVIDADE INDIVIDUAL 1', 'ATIVIDADE INDIVIDUAL 2')
    $RemotoDrive   = 'gdrive:TESTE'
    $CidadesRemoto = @{ 'Pariquera-Açu' = 'PARIQUERA AÇU'; 'Juquia' = 'Juquiá' }
    $CidadesOneDrive = @{ 'Juquiá' = 'Juquia' }
    $IgnoradosFile = Join-Path $Tmp 'ignorados.json'

    # Pastas válidas p/ a trava anti-retroativo: nome com "- RESUMO" + .txt de relatório
    foreach ($Ativ in $Atividades) {
        foreach ($Nome in @('07-2026-11-JOAO DA SILVA - RESUMO', '03 - Pedro Oliveira - RESUMO')) {
            $pastaValida = Join-Path $OrigemBase "Jacupiranga\$Ativ\$Nome"
            New-Item -ItemType Directory -Path $pastaValida -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $pastaValida 'relatorio.txt') -Encoding UTF8 -Value 'Corpo do relatorio'
        }
    }
    # Pastas que a trava DEVE bloquear (nunca sobem retroativamente ao GDSA)
    $baseAtv1 = Join-Path $OrigemBase 'Jacupiranga\ATIVIDADE INDIVIDUAL 1'
    $pSemResumo = Join-Path $baseAtv1 '04-2026-11-Juari'
    New-Item -ItemType Directory -Path $pSemResumo -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $pSemResumo 'relatorio.txt') -Encoding UTF8 -Value 'Corpo'
    $pSoFotos = Join-Path $baseAtv1 '05-2026-11-Coicieçao E Hermes - RESUMO'
    New-Item -ItemType Directory -Path $pSoFotos -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $pSoFotos 'IMG_20261105_090000.jpg') | Out-Null
    $pSoGps = Join-Path $baseAtv1 '06-2026-11-Maria Da Graça - RESUMO'
    New-Item -ItemType Directory -Path $pSoGps -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $pSoGps 'coordenadas_gps.txt') -Encoding UTF8 -Value 'GPS'

    # 1) nome no destino = mesmo nome local, data completa vira DD-AAAA-MM- (dia primeiro)
    if ((Get-NomeResumo '2026-11-07-MARIA DA GRACA - MODESTO - RESUMO') -ne '07-2026-11-MARIA DA GRACA - MODESTO - RESUMO') { throw 'renomeação: data completa deveria virar DD-AAAA-MM-' }
    if ((Get-NomeResumo '2026-11-07-MARIA DA GRACA - MODESTO') -ne '07-2026-11-MARIA DA GRACA - MODESTO - RESUMO') { throw 'renomeação: sufixo RESUMO não adicionado' }
    if ((Get-NomeResumo 'SEM PREFIXO') -ne 'SEM PREFIXO - RESUMO') { throw 'renomeação sem prefixo falhou' }

    # 1b) idempotência do formato novo e do - RESUMO no start-sync (não duplicar)
    if ((Get-NomeResumo '07-2026-11-CARLOS - ITAPITANGUI - RESUMO') -ne '07-2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'idempotência: formato novo foi alterado' }
    if ((Get-NomeResumo 'CARLOS - ITAPITANGUI - RESUMO') -ne 'CARLOS - ITAPITANGUI - RESUMO') { throw 'idempotência: duplicou RESUMO sem prefixo' }
    if ((Get-NomeResumo 'JOAO - resumo') -ne 'JOAO - resumo') { throw 'idempotência case-insensitive falhou' }

    # 1c) Add-SufixoResumo da triagem (idempotente, case-insensitive)
    if ((Add-SufixoResumo '2026-11-CARLOS - ITAPITANGUI') -ne '2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: sufixo não adicionado' }
    if ((Add-SufixoResumo '2026-11-CARLOS - ITAPITANGUI - RESUMO') -ne '2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: duplicou RESUMO' }
    if ((Add-SufixoResumo 'NOME - RESUMO') -ne 'NOME - RESUMO') { throw 'triagem: idempotência falhou' }

    # 1c2) Test-CampoSeguro: traversal, caracteres de caminho e ponto/espaco final
    if (Test-CampoSeguro '..\evil') { throw 'triagem: path traversal deveria ser rejeitado' }
    if (Test-CampoSeguro 'NOME.') { throw 'triagem: ponto final deveria ser rejeitado' }
    if (Test-CampoSeguro 'NOME ') { throw 'triagem: espaco final deveria ser rejeitado' }
    if (-not (Test-CampoSeguro 'CARLOS - ITAPITANGUI')) { throw 'triagem: nome valido rejeitado' }

    # 1c3) inbox com par SO .json (sem .txt): nao pode quebrar no StrictMode
    $InboxBase = Join-Path $Tmp 'inbox'
    $leafJson = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Jacupiranga\ATIVIDADE INDIVIDUAL 1\2026-11-17-SO JSON'
    New-Item -ItemType Directory -Path $leafJson -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $leafJson 'relatorio.json') -Encoding UTF8 -Value '{"metadados":{"data":"2026-11-17","cidade":"Jacupiranga","atividade":"ATIVIDADE INDIVIDUAL 1","produtor_pasta":"SO JSON"}}'
    $inboxTest = @(& { Set-StrictMode -Version Latest; Get-InboxVisitas })
    if ($inboxTest.Count -ne 1) { throw 'triagem: par so .json deveria ser detectado' }
    if ($null -ne $inboxTest[0].SourceTxt) { throw 'triagem: SourceTxt deveria ser nulo sem .txt' }

    # 1d) Format-NomePastaProdutor: prefixo DD-AAAA-MM- (dia primeiro) + sufixo - RESUMO
    if ((Format-NomePastaProdutor -Nome 'CARLOS - ITAPITANGUI' -Data '2026-11-12') -ne '12-2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: prefixo DD-AAAA-MM- + sufixo não garantidos' }
    if ((Format-NomePastaProdutor -Nome '12-2026-11-CARLOS - ITAPITANGUI - RESUMO' -Data '2026-11-12') -ne '12-2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: duplicou prefixo/sufixo' }
    if ((Format-NomePastaProdutor -Nome '2026-11-12-CARLOS - ITAPITANGUI - RESUMO' -Data '2026-11-12') -ne '12-2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: prefixo antigo não normalizado' }
    if ((Format-NomePastaProdutor -Nome '2026-11-CARLOS' -Data '2026-11-12') -ne '12-2026-11-CARLOS - RESUMO') { throw 'triagem: prefixo parcial não normalizado' }

    # 1e) start-sync converte prefixo de data COMPLETA (AAAA-MM-DD-) para DD-AAAA-MM-
    if ((Get-NomeResumo '2026-11-12-CARLOS - ITAPITANGUI - RESUMO') -ne '12-2026-11-CARLOS - ITAPITANGUI - RESUMO') { throw 'start-sync: não converteu dia completo para DD-AAAA-MM-' }

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
    $va = New-Visita -Data '2026-11-12' -Cidade 'Juquiá' -Atividade 'ATIVIDADE INDIVIDUAL 1' -ProdutorPasta 'CARLOS'
    $vb = New-Visita -Data '2026-11-12' -Cidade 'Juquiá' -Atividade 'ATIVIDADE INDIVIDUAL 1' -ProdutorPasta 'ANA'
    $dias = ConvertTo-RelatorioDias -Visitas @($va, $vb)
    if ($dias[0].visitas[0].tag -ne 'A') { throw 'triagem: tag A não atribuída' }
    if ($dias[0].visitas[1].produtor_pasta -ne '12-2026-11-ANA - RESUMO') { throw 'triagem: nome de pasta não normalizado' }

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

    # 1j) limpar-screenshots: varredura bottom-up de pastas vazias (apps + legado)
    $DestinoBase = Join-Path $Tmp 'Screenshots'
    foreach ($d in @('AppVazio', '2026\01', 'ComPrint')) { New-Item -ItemType Directory -Path (Join-Path $DestinoBase $d) -Force | Out-Null }
    New-Item -ItemType File -Path (Join-Path $DestinoBase 'ComPrint\a.jpg') | Out-Null
    $vazias = Remove-PastasVazias
    if ($vazias -notcontains 'AppVazio') { throw 'screenshots: pasta de app vazia não removida' }
    if ($vazias -notcontains '2026\01' -or $vazias -notcontains '2026') { throw 'screenshots: pastas legadas vazias não removidas (bottom-up)' }
    if (-not (Test-Path (Join-Path $DestinoBase 'ComPrint\a.jpg'))) { throw 'screenshots: removeu pasta que tinha arquivo' }

    # 1k) organizar-camera: GPS, Haversine, clusters, nomenclatura, _visita.json, videos, estado, colisao
    $ExtensoesFoto = @('.jpg', '.jpeg', '.png')
    $mockImg = [PSCustomObject]@{ PropertyItems = @(
        [PSCustomObject]@{ Id = 2; Value = [byte[]]@(23,0,0,0,1,0,0,0, 30,0,0,0,1,0,0,0, 0,0,0,0,1,0,0,0) }
        [PSCustomObject]@{ Id = 1; Value = [byte[]]@(83,0) }
        [PSCustomObject]@{ Id = 4; Value = [byte[]]@(47,0,0,0,1,0,0,0, 15,0,0,0,1,0,0,0, 0,0,0,0,1,0,0,0) }
        [PSCustomObject]@{ Id = 3; Value = [byte[]]@(87,0) }
    ) }
    if ([math]::Abs((Get-GpsCoord $mockImg 2 1) - (-23.5)) -gt 1e-9) { throw 'camera: GPS latitude falhou' }
    if ([math]::Abs((Get-GpsCoord $mockImg 4 3) - (-47.25)) -gt 1e-9) { throw 'camera: GPS longitude falhou' }
    if ([math]::Abs((Get-DistanciaMetros 0 0 0 1) - 111195) -gt 100) { throw 'camera: Haversine falhou' }
    if ((Get-DistanciaMetros -23.5 -47.25 -23.5 -47.25) -ne 0) { throw 'camera: Haversine ponto igual falhou' }
    function New-FotoTeste([string]$nome, [datetime]$dt, $lat, $lon) {
        [PSCustomObject]@{ Arquivo = [PSCustomObject]@{ Name = $nome; Length = 1000 }; DataHora = $dt; Lat = $lat; Lon = $lon; TemGps = ($null -ne $lat -and $null -ne $lon); Modelo = 'Xiaomi 21121210G' }
    }
    $ft = @(
        (New-FotoTeste 'IMG_20260402_090000.jpg' ([datetime]'2026-04-02 09:00:00') -23.5 -47.25),
        (New-FotoTeste 'IMG_20260402_091000.jpg' ([datetime]'2026-04-02 09:10:00') -23.5 -47.25),
        (New-FotoTeste 'IMG_20260402_092000.jpg' ([datetime]'2026-04-02 09:20:00') -23.5 -47.25),
        (New-FotoTeste 'IMG_20260402_093000.jpg' ([datetime]'2026-04-02 09:30:00') -23.5 -47.25),
        (New-FotoTeste 'IMG_20260402_113000.jpg' ([datetime]'2026-04-02 11:30:00') -23.509 -47.25),
        (New-FotoTeste 'IMG_20260402_113500.jpg' ([datetime]'2026-04-02 11:35:00') -23.509 -47.25),
        (New-FotoTeste 'IMG_20260403_090000.jpg' ([datetime]'2026-04-03 09:00:00') $null $null),
        (New-FotoTeste 'IMG_20260403_093000.jpg' ([datetime]'2026-04-03 09:30:00') $null $null),
        (New-FotoTeste 'IMG_20260403_094000.jpg' ([datetime]'2026-04-03 09:40:00') $null $null),
        (New-FotoTeste 'IMG_20260402_130000.jpg' ([datetime]'2026-04-02 13:00:00') -23.52 -47.25),
        (New-FotoTeste 'IMG_20260402_130500.jpg' ([datetime]'2026-04-02 13:05:00') -23.52 -47.25),
        (New-FotoTeste 'IMG_20260402_131000.jpg' ([datetime]'2026-04-02 13:10:00') -23.52 -47.25)
    )
    $clT = Get-ClustersVisita -Fotos $ft
    if ($clT.Count -ne 4) { throw "camera: clusters esperados 4, obtidos $($clT.Count)" }
    $visT = Set-NomesVisita -Clusters $clT
    if ($visT.Count -ne 3) { throw "camera: visitas esperadas 3, obtidas $($visT.Count)" }
    if ($visT[0].PastaNome -ne '02_visita_1' -or $visT[1].PastaNome -ne '02_visita_2') { throw 'camera: nomenclatura DD_visita_1/_2 falhou' }
    if ($visT[2].PastaNome -ne '03_visita') { throw 'camera: nomenclatura DD_visita (unica) falhou' }
    $vj = New-VisitaJson -Cluster $visT[0] -PastaNome $visT[0].PastaNome -Indice $visT[0].IndiceVisita
    if ($vj.data -ne '2026-04-02' -or $vj.qtd_fotos -ne 4 -or $vj.fotos_com_gps -ne 4) { throw 'camera: _visita.json campos falhou' }
    if ($vj.hora_inicio -ne '09:00:00' -or $vj.hora_fim -ne '09:30:00' -or $vj.arquivos.Count -ne 4) { throw 'camera: _visita.json horarios/arquivos falhou' }
    if ($vj.google_maps_url -notmatch 'maps\?q=-23\.5,-47\.25') { throw 'camera: _visita.json URL Maps (invariante) falhou' }
    $imgT = New-Item -ItemType File -Path (Join-Path $Tmp 'IMG_20260402_090000.jpg')
    $miT = Get-MetadadosCamera -Arquivo $imgT
    if ($miT.DataHora.ToString('yyyy-MM-dd HH:mm') -ne '2026-04-02 09:00') { throw 'camera: fallback data pelo nome falhou' }
    $vidT = New-Item -ItemType File -Path (Join-Path $Tmp 'VID_20261105_184822.mp4')
    $mvT = Get-MetadadosCamera -Arquivo $vidT
    if ($mvT.DataHora.ToString('yyyy-MM-dd HH:mm:ss') -ne '2026-11-05 18:48:22') { throw 'camera: data de video pelo nome falhou' }
    if ((Join-Path 'Videos' $mvT.DataHora.ToString('yyyy\\MM')) -ne 'Videos\2026\11') { throw 'camera: destino Videos\AAAA\MM falhou' }
    $EstadoFile = Join-Path $Tmp 'estado-camera.json'
    $est0 = Get-EstadoCamera
    if ($est0.total_fotos_organizadas_c -ne 0) { throw 'camera: estado inicial falhou' }
    $est0.ultima_execucao = '2026-11-17_14-00-00'; $est0.total_fotos_organizadas_c = 42
    Save-EstadoCamera -Estado $est0
    $est1 = Get-EstadoCamera
    if ($est1.total_fotos_organizadas_c -ne 42 -or $est1.ultima_execucao -ne '2026-11-17_14-00-00') { throw 'camera: estado roundtrip falhou' }
    $colP = Join-Path $Tmp 'colisao'; New-Item -ItemType Directory -Path $colP -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $colP 'IMG_20260402_090000.jpg') | Out-Null
    if ((Split-Path (Get-NomeSemColisaoCamera -Pasta $colP -Arquivo $imgT) -Leaf) -ne 'IMG_20260402_090000__2.jpg') { throw 'camera: colisao __2 falhou' }

    # 1m) triagem: busca cirurgica por data + leitura zero-EXIF via _visita.json + dedup
    $orgC = Join-Path $Tmp 'orgC'; $orgE = Join-Path $Tmp 'orgE'; $bruta = Join-Path $Tmp 'bruta'
    $visDir = Join-Path $orgC '2026\11\06_visita_1'
    New-Item -ItemType Directory -Path $visDir -Force | Out-Null
    New-Item -ItemType Directory -Path $bruta -Force | Out-Null
    # arquivos vazios: se o script abrisse EXIF nao haveria GPS/hora -> prova de leitura via JSON
    New-Item -ItemType File -Path (Join-Path $visDir 'IMG_20261106_144600.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $visDir 'foto_sem_padrao.jpg') | Out-Null
    $vjTeste = [PSCustomObject]@{
        data = '2026-11-06'; hora_inicio = '14:46:00'; hora_fim = '15:30:00'; qtd_fotos = 2
        lat_media = -23.5; lon_media = -47.25; raio_metros = 10.0
        arquivos = @(
            [PSCustomObject]@{ nome = 'IMG_20261106_144600.jpg'; hora = '14:46:00'; lat = -23.5; lon = -47.25 },
            [PSCustomObject]@{ nome = 'foto_sem_padrao.jpg'; hora = '15:30:00'; lat = -23.501; lon = -47.251 }
        )
    }
    $vjTeste | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $visDir '_visita.json') -Encoding UTF8
    # avulsa do dia alvo na raiz do mes + foto de OUTRO dia (nao pode entrar)
    New-Item -ItemType File -Path (Join-Path $orgC '2026\11\IMG_20261106_160000.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $orgC '2026\11\IMG_20261107_100000.jpg') | Out-Null
    # bruta: casa pelo nome, casa pelo LastWriteTime, e nao casa (outro dia)
    New-Item -ItemType File -Path (Join-Path $bruta 'IMG_20261106_170000.jpg') | Out-Null
    $semPadrao = New-Item -ItemType File -Path (Join-Path $bruta 'SEM_PADRAO.jpg')
    $semPadrao.LastWriteTime = [datetime]'2026-11-06 18:00:00'
    $outroDia = New-Item -ItemType File -Path (Join-Path $bruta 'IMG_20261105_090000.jpg')
    $outroDia.LastWriteTime = [datetime]'2026-11-05 09:00:00'
    # dedup: mesma visita no "E:" nao pode duplicar fotos
    $visDirE = Join-Path $orgE '2026\11\06_visita_1'
    New-Item -ItemType Directory -Path $visDirE -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $visDir '_visita.json') -Destination $visDirE
    New-Item -ItemType File -Path (Join-Path $visDirE 'IMG_20261106_144600.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $visDirE 'foto_sem_padrao.jpg') | Out-Null
    $fc = Get-FotosCirurgico -Raizes @($orgC, $orgE, $bruta) -DatasAlvo @('2026-11-06') -Extensoes $ExtensoesFoto
    if ($fc.Count -ne 5) { throw "triagem: busca cirurgica esperava 5 fotos, obtidas $($fc.Count)" }
    $daVisita = @($fc | Where-Object { $_.LoteVisita })
    if ($daVisita.Count -ne 2) { throw 'triagem: fotos da visita deveriam vir pre-agrupadas (LoteVisita)' }
    $semPadraoJson = $daVisita | Where-Object { $_.Arquivo.Name -eq 'foto_sem_padrao.jpg' }
    if ($semPadraoJson.HoraStr -ne '15:30:00' -or -not $semPadraoJson.TemGps) { throw 'triagem: hora/GPS deveriam vir do _visita.json (zero EXIF)' }
    if ($semPadraoJson.Lat -ne -23.501) { throw 'triagem: lat do _visita.json falhou' }
    if (@($fc | Where-Object { $_.DataStr -ne '2026-11-06' }).Count -ne 0) { throw 'triagem: fotos de outro dia nao podem entrar' }
    if (@($fc | Where-Object { $_.Arquivo.Name -eq 'SEM_PADRAO.jpg' }).Count -ne 1) { throw 'triagem: fallback LastWriteTime do dia falhou' }
    if (@($fc | Where-Object { $_.Arquivo.Name -eq 'IMG_20261106_144600.jpg' }).Count -ne 1) { throw 'triagem: dedup por nome entre raizes falhou' }
    # origem curta dos lotes (AAAA\MM\DD_visita / avulsa / Camera bruta), sem C:\Users\...
    $fVis = $fc | Where-Object { $_.Arquivo.Name -eq 'foto_sem_padrao.jpg' }
    if ((Get-OrigemCurtaFoto -Foto $fVis) -ne '2026\11\06_visita_1') { throw 'triagem: origem curta da visita falhou' }
    $fAv = $fc | Where-Object { $_.Arquivo.Name -eq 'IMG_20261106_160000.jpg' }
    if ((Get-OrigemCurtaFoto -Foto $fAv) -ne '2026\11 (avulsa)') { throw 'triagem: origem curta da avulsa falhou' }
    $fCam = [PSCustomObject]@{ LoteVisita = $null; Arquivo = [PSCustomObject]@{ FullName = 'C:\Sync\Camera_Celular\Camera\IMG_20261106_170000.jpg' } }
    if ((Get-OrigemCurtaFoto -Foto $fCam) -ne 'Camera (bruta)') { throw 'triagem: origem curta da Camera bruta falhou' }

    # 1n) triagem: horario sem -null e atalho S
    if ((Format-HorarioVisita '14:46' 'null') -ne '14:46') { throw 'triagem: hora_fim "null" vazou na exibicao' }
    if ((Format-HorarioVisita '14:46' '15:30') -ne '14:46-15:30') { throw 'triagem: intervalo completo falhou' }
    if ((Format-HorarioVisita $null $null) -ne '') { throw 'triagem: ambos vazios deveria ser vazio' }
    if ((Format-HorarioVisita 'null' '15:30') -ne '15:30') { throw 'triagem: hora_inicio "null" falhou' }
    if ((Resolve-TagSugestao -Escolha 'S' -Sugestao 'B') -ne 'B') { throw 'triagem: atalho S nao aceitou a sugestao' }
    if ((Resolve-TagSugestao -Escolha '' -Sugestao 'B') -ne 'B') { throw 'triagem: Enter nao aceitou a sugestao' }
    if ((Resolve-TagSugestao -Escolha 'A' -Sugestao 'B') -ne 'A') { throw 'triagem: tag explicita nao preservada' }
    if ((Resolve-TagSugestao -Escolha 'S' -Sugestao $null) -ne 'S') { throw 'triagem: S sem sugestao deveria seguir como tag' }

    # 1o) camera: indice central _indice_visitas.json (zero EXIF, nada movido)
    $visDir2 = Join-Path $orgC '2026\11\07_visita'
    New-Item -ItemType Directory -Path $visDir2 -Force | Out-Null
    ([PSCustomObject]@{ data = '2026-11-07'; hora_inicio = '09:00:00'; hora_fim = '10:00:00'; qtd_fotos = 3; lat_media = -24.0; lon_media = -48.0; raio_metros = 5.0; arquivos = @() }) |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $visDir2 '_visita.json') -Encoding UTF8
    $idxPath = New-IndiceVisitas -Raiz $orgC
    if (-not (Test-Path -LiteralPath $idxPath)) { throw 'camera: _indice_visitas.json nao gerado' }
    $idxObj = Get-Content -LiteralPath $idxPath -Raw | ConvertFrom-Json
    $e06 = @($idxObj.'2026-11-06')
    if ($e06.Count -ne 1 -or $e06[0].pasta_relativa -ne '2026\11\06_visita_1') { throw 'camera: indice pasta_relativa falhou' }
    if ($e06[0].qtd_fotos -ne 2 -or $e06[0].hora_inicio -ne '14:46:00' -or $e06[0].visita_json -ne '2026\11\06_visita_1\_visita.json') { throw 'camera: indice campos falhou' }
    if (@($idxObj.'2026-11-07').Count -ne 1) { throw 'camera: indice segunda data falhou' }
    if (-not (Test-Path -LiteralPath (Join-Path $visDir 'IMG_20261106_144600.jpg'))) { throw 'camera: indice nao pode mover/alterar fotos' }

    # 1p) triagem: copia nao-destrutiva da inbox + anexos ATV1 + _processados.json append/idempotencia
    $OneDriveBase = Join-Path $Tmp 'onedrive'
    $LogFile      = Join-Path $Tmp 'triagem-log.txt'
    $InboxBase    = Join-Path $Tmp 'inbox2'
    $leaf1 = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Jacupiranga\ATIVIDADE INDIVIDUAL 1\2026-11-18-JOAO - RESUMO'
    New-Item -ItemType Directory -Path $leaf1 -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $leaf1 'relatorio.txt') -Encoding UTF8 -Value "---`ndata: 2026-11-18`n---`nCorpo do relatorio JOAO"
    Set-Content -LiteralPath (Join-Path $leaf1 'relatorio.json') -Encoding UTF8 -Value '{"metadados":{"data":"2026-11-18","cidade":"Jacupiranga","atividade":"ATIVIDADE INDIVIDUAL 1","produtor_pasta":"JOAO - RESUMO"}}'
    foreach ($n in @('2026-11-18-JOAO - Mapa.png', '2026-11-18-JOAO - Mapa.jpeg', 'calendario-agricola.png', 'alendario-agricola.png', 'foto_qualquer.png', 'doc.gdoc')) {
        New-Item -ItemType File -Path (Join-Path $leaf1 $n) | Out-Null
    }
    $inbox1 = Get-InboxVisitas
    if (@($inbox1).Count -ne 1) { throw "triagem: inbox deveria detectar 1 visita, detectou $(@($inbox1).Count)" }
    $dias1 = ConvertTo-RelatorioDias -Visitas $inbox1
    $null = Invoke-CriarPastasEArquivar -RelatorioDias $dias1
    $dest1 = $inbox1[0].CaminhoDestino
    # originais INTACTOS na inbox (nada movido/apagado, sem arvore _processados\)
    foreach ($n in @('relatorio.txt', 'relatorio.json', 'doc.gdoc', '2026-11-18-JOAO - Mapa.png')) {
        if (-not (Test-Path -LiteralPath (Join-Path $leaf1 $n))) { throw "triagem: original '$n' nao pode sair da inbox" }
    }
    if (Test-Path -LiteralPath (Join-Path $InboxBase '2026\_inbox_relatorio_ater\_processados')) { throw 'triagem: arvore legada _processados\ nao pode ser criada' }
    # .txt limpo + anexos mapa/calendario no OneDrive; .json fica APENAS na inbox
    $txtDest = Get-Content -LiteralPath (Join-Path $dest1 'relatorio.txt') -Raw
    if ($txtDest -match '---' -or $txtDest.Trim() -ne 'Corpo do relatorio JOAO') { throw 'triagem: .txt destino deveria estar limpo (sem YAML)' }
    if (Test-Path -LiteralPath (Join-Path $dest1 'relatorio.json')) { throw 'triagem: .json nao deve ser copiado ao OneDrive (fica so na inbox)' }
    if (-not (Test-Path -LiteralPath (Join-Path $leaf1 'relatorio.json'))) { throw 'triagem: .json original deveria permanecer na inbox' }
    foreach ($n in @('2026-11-18-JOAO - Mapa.png', '2026-11-18-JOAO - Mapa.jpeg', 'calendario-agricola.png', 'alendario-agricola.png')) {
        if (-not (Test-Path -LiteralPath (Join-Path $dest1 $n))) { throw "triagem: anexo ATV1 '$n' nao copiado" }
    }
    foreach ($n in @('foto_qualquer.png', 'doc.gdoc')) {
        if (Test-Path -LiteralPath (Join-Path $dest1 $n)) { throw "triagem: '$n' nao deveria ser copiado" }
    }
    # registro unico com campos corretos
    $regPath = Join-Path $InboxBase '2026\_inbox_relatorio_ater\_processados.json'
    if (-not (Test-Path -LiteralPath $regPath)) { throw 'triagem: _processados.json nao criado' }
    $reg1 = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if ($reg1.Count -ne 1) { throw 'triagem: registro deveria ter 1 entrada' }
    if ($reg1[0].pasta_inbox -ne $leaf1 -or $reg1[0].destino_onedrive -ne $dest1) { throw 'triagem: pasta_inbox/destino_onedrive incorretos' }
    if ($reg1[0].data -ne '2026-11-18' -or $reg1[0].processado_em -notmatch '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$') { throw 'triagem: data/processado_em incorretos' }
    if ($reg1[0].arquivos_levados -notcontains 'relatorio.txt' -or $reg1[0].arquivos_levados -notcontains 'calendario-agricola.png') { throw 'triagem: arquivos_levados incompleto' }
    if ($reg1[0].arquivos_levados -contains 'relatorio.json') { throw 'triagem: arquivos_levados nao deve listar o .json (nao vai ao OneDrive)' }
    # idempotencia: visita registrada nao volta; ATV2 nova aparece e NAO leva anexos
    $leaf2 = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Jacupiranga\ATIVIDADE INDIVIDUAL 2\2026-11-18-MARIA - RESUMO'
    New-Item -ItemType Directory -Path $leaf2 -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $leaf2 'relatorio.txt') -Encoding UTF8 -Value "---`ndata: 2026-11-18`n---`nCorpo MARIA"
    Set-Content -LiteralPath (Join-Path $leaf2 'relatorio.json') -Encoding UTF8 -Value '{"metadados":{"data":"2026-11-18","cidade":"Jacupiranga","atividade":"ATIVIDADE INDIVIDUAL 2","produtor_pasta":"MARIA - RESUMO"}}'
    New-Item -ItemType File -Path (Join-Path $leaf2 'MARIA - Mapa.png') | Out-Null
    $inbox2 = Get-InboxVisitas
    if (@($inbox2).Count -ne 1 -or $inbox2[0].PastaInbox -ne $leaf2) { throw 'triagem: idempotencia falhou (visita registrada voltou ou ATV2 nao detectada)' }
    $null = Invoke-CriarPastasEArquivar -RelatorioDias (ConvertTo-RelatorioDias -Visitas $inbox2)
    $dest2 = $inbox2[0].CaminhoDestino
    if (Test-Path -LiteralPath (Join-Path $dest2 'MARIA - Mapa.png')) { throw 'triagem: anexo de mapa nao deve ser copiado fora da ATV 1' }
    if (-not (Test-Path -LiteralPath (Join-Path $leaf2 'MARIA - Mapa.png'))) { throw 'triagem: original da ATV2 nao pode sair da inbox' }
    $reg2 = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if ($reg2.Count -ne 2) { throw "triagem: append cumulativo falhou (esperado 2, obtido $($reg2.Count))" }
    $inbox3 = Get-InboxVisitas
    if (@($inbox3).Count -ne 0) { throw 'triagem: segunda leitura deveria pular todas as visitas registradas' }

    # 1q) ciclo de vida sincronizada false->true (triagem registra, start-sync marca) + atomicidade + idempotencia
    $regCiclo = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    foreach ($r in $regCiclo) {
        if ($r.sincronizada -ne $false -or $null -ne $r.sincronizada_em -or $null -ne $r.drive_url) { throw 'triagem: registro novo deveria nascer sincronizada=false, sincronizada_em/drive_url nulos' }
    }
    # cache fake do Drive p/ Get-DriveFolderUrl resolver a URL sem rclone
    $Script:DriveDirCache = @{ "$RemotoDrive/ATIVIDADE INDIVIDUAL 1/Jacupiranga" = @([PSCustomObject]@{ Name = '18-2026-11-JOAO - RESUMO'; ID = 'IDJOAO123' }) }
    $procAntes = $regCiclo[0].processado_em
    $nMarc = Update-ProcessadosSyncStatus -Sincronizadas @([PSCustomObject]@{ Local = $dest1; Cidade = 'Jacupiranga'; Atividade = 'ATIVIDADE INDIVIDUAL 1'; NomeDrive = 'JOAO - RESUMO' })
    if ($nMarc -ne 1) { throw "start-sync: deveria marcar 1 visita, marcou $nMarc" }
    $regApos = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if ($regApos.Count -ne 2) { throw 'start-sync: append/registros nao podem ser perdidos na atualizacao' }
    $joao = $regApos | Where-Object { $_.pasta_inbox -eq $leaf1 }
    $maria = $regApos | Where-Object { $_.pasta_inbox -eq $leaf2 }
    if ($joao.sincronizada -ne $true -or $joao.sincronizada_em -notmatch '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$') { throw 'start-sync: sincronizada/sincronizada_em nao atualizados' }
    if ($joao.drive_url -ne 'https://drive.google.com/drive/folders/IDJOAO123') { throw 'start-sync: drive_url nao preenchida' }
    if ($joao.processado_em -ne $procAntes -or $joao.destino_onedrive -ne $dest1 -or $joao.arquivos_levados -notcontains 'calendario-agricola.png') { throw 'start-sync: campos existentes foram alterados/apagados' }
    if ($maria.sincronizada -ne $false -or $null -ne $maria.drive_url) { throw 'start-sync: registro nao sincronizado foi alterado indevidamente' }
    if (Test-Path -LiteralPath "$regPath.tmp") { throw 'start-sync: gravacao atomica deixou .tmp para tras' }
    # idempotencia: segunda execucao nao duplica nem corrompe
    $null = Update-ProcessadosSyncStatus -Sincronizadas @([PSCustomObject]@{ Local = $dest1; Cidade = 'Jacupiranga'; Atividade = 'ATIVIDADE INDIVIDUAL 1'; NomeDrive = 'JOAO - RESUMO' })
    $regIdem = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if ($regIdem.Count -ne 2) { throw 'start-sync: segunda execucao duplicou registros' }
    if (($regIdem | Where-Object { $_.pasta_inbox -eq $leaf1 }).drive_url -ne 'https://drive.google.com/drive/folders/IDJOAO123') { throw 'start-sync: idempotencia perdeu drive_url' }
    # comparacao case-insensitive e sem barra final: marca MARIA com caminho "sujo"
    $null = Update-ProcessadosSyncStatus -Sincronizadas @([PSCustomObject]@{ Local = ($dest2.ToLower() + '\'); Cidade = 'Jacupiranga'; Atividade = 'ATIVIDADE INDIVIDUAL 2'; NomeDrive = 'MARIA - RESUMO' })
    $regFinal = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if (($regFinal | Where-Object { $_.pasta_inbox -eq $leaf2 }).sincronizada -ne $true) { throw 'start-sync: match case-insensitive/sem barra final falhou' }
    $resumo = Get-ResumoProcessadosAno -Ano '2026'
    if ($resumo.triadas -ne 2 -or $resumo.sincronizadas -ne 2 -or $resumo.pendentes -ne 0) { throw 'triagem: resumo do ano incorreto' }

    # 1r) modo de fotos no _processados.json: campos iniciais + Update-FotosProcessados (manual/automatico/pulado)
    $regF0 = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    foreach ($r in $regF0) {
        if ($r.fotos_modo -ne 'pulado' -or $r.fotos_copiadas -ne 0 -or $null -ne $r.fotos_origem_sugerida) { throw 'triagem: fotos_modo/fotos_copiadas/fotos_origem_sugerida iniciais incorretos' }
    }
    $inbox1[0].tag = 'A'; $inbox2[0].tag = 'B'
    $diaFotos = [PSCustomObject]@{ data = '2026-11-18'; visitas = @($inbox1[0], $inbox2[0]) }
    $statusFotos = @{
        'A' = [PSCustomObject]@{ Modo = 'manual'; Copiadas = 0; Origem = 'C:\origem\lote_vizinhos' }
        'B' = [PSCustomObject]@{ Modo = 'automatico'; Copiadas = 5; Origem = 'C:\origem\visita_b' }
    }
    Update-FotosProcessados -Dia $diaFotos -Status $statusFotos
    $regF1 = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    $jf = $regF1 | Where-Object { $_.pasta_inbox -eq $leaf1 }
    $mf = $regF1 | Where-Object { $_.pasta_inbox -eq $leaf2 }
    if ($jf.fotos_modo -ne 'manual' -or $jf.fotos_origem_sugerida -ne 'C:\origem\lote_vizinhos') { throw 'triagem: modo manual nao registrado' }
    if ($mf.fotos_modo -ne 'automatico' -or $mf.fotos_copiadas -ne 5) { throw 'triagem: modo automatico/contagem nao registrado' }
    if ($jf.sincronizada -ne $true -or $jf.drive_url -ne 'https://drive.google.com/drive/folders/IDJOAO123') { throw 'triagem: Update-FotosProcessados sobrescreveu campos de sync' }
    if ($regF1.Count -ne 2 -or (Test-Path -LiteralPath "$regPath.tmp")) { throw 'triagem: append atomico de fotos falhou' }
    # visita sem status -> pulado (0 fotos); segunda chamada nao duplica
    Update-FotosProcessados -Dia $diaFotos -Status @{}
    $regF2 = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if (($regF2 | Where-Object { $_.pasta_inbox -eq $leaf1 }).fotos_modo -ne 'pulado') { throw 'triagem: visita sem status deveria ficar pulado' }
    if ($regF2.Count -ne 2) { throw 'triagem: segunda chamada duplicou registros' }

    # 1s) -Conferir: listar processados, auditar TXT/fotos, reparar TXT e fotos faltantes (nao-destrutivo)
    $SyncthingFotos = @((Join-Path $Tmp 'syncthing'))
    $todosConf = Get-ProcessadosTodos
    if (@($todosConf).Count -ne 2) { throw 'conferir: deveria listar 2 visitas processadas' }
    if (-not $todosConf[0].registro_path) { throw 'conferir: registro_path ausente' }
    $rJ = $todosConf | Where-Object { $_.pasta_inbox -eq $leaf1 }
    $audJ = Test-VisitaDestino -Registro $rJ
    if (@($audJ.TxtFaltando).Count -ne 0 -or $audJ.FotosNoDestino -ne 4) { throw 'conferir: auditoria inicial incorreta' }
    # apaga o txt do destino -> auditoria detecta; reparo recopia limpo da inbox (inbox intacta)
    Remove-Item -LiteralPath (Join-Path $dest1 'relatorio.txt') -Force
    $audJ2 = Test-VisitaDestino -Registro $rJ
    if (@($audJ2.TxtFaltando).Count -ne 1) { throw 'conferir: TXT faltante nao detectado' }
    $null = Repair-VisitaTxtAnexos -Registro $rJ
    $txtRep = Get-Content -LiteralPath (Join-Path $dest1 'relatorio.txt') -Raw
    if ($txtRep -match '---' -or $txtRep.Trim() -ne 'Corpo do relatorio JOAO') { throw 'conferir: TXT reparado deveria estar limpo (sem YAML)' }
    if (-not (Test-Path -LiteralPath (Join-Path $leaf1 'relatorio.txt'))) { throw 'conferir: inbox nao pode ser alterada' }
    # fotos faltantes: origem cirurgica com 2 fotos, destino com 0 -> 2 faltantes; copia 1 -> 1 faltante
    $visConf = Join-Path $Tmp 'syncthing\2026\11\18_visita'
    New-Item -ItemType Directory -Path $visConf -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $visConf 'IMG_20261118_090000.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $visConf 'IMG_20261118_091000.jpg') | Out-Null
    ([PSCustomObject]@{ data = '2026-11-18'; arquivos = @(
        [PSCustomObject]@{ nome = 'IMG_20261118_090000.jpg'; hora = '09:00:00'; lat = -23.5; lon = -47.25 },
        [PSCustomObject]@{ nome = 'IMG_20261118_091000.jpg'; hora = '09:10:00'; lat = -23.5; lon = -47.25 }) }) |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $visConf '_visita.json') -Encoding UTF8
    Add-Member -InputObject $rJ -NotePropertyName fotos_origem_sugerida -NotePropertyValue $visConf -Force
    $falt1 = Get-FotosFaltantesVisita -Registro $rJ
    if (@($falt1).Count -ne 2) { throw "conferir: esperava 2 fotos faltantes, obtidas $(@($falt1).Count)" }
    Copy-Item -LiteralPath (Join-Path $visConf 'IMG_20261118_090000.jpg') -Destination $dest1 -Force
    $falt2 = Get-FotosFaltantesVisita -Registro $rJ
    if (@($falt2).Count -ne 1) { throw 'conferir: apos copiar 1, deveria faltar 1' }
    # save atomico: sem registro_path no JSON, campos preservados, sem .tmp residual
    Save-ProcessadosPorArquivo -Registros $todosConf
    if ((Get-Content -LiteralPath $regPath -Raw) -match 'registro_path') { throw 'conferir: registro_path nao pode ir para o JSON' }
    $regConf = @(Get-Content -LiteralPath $regPath -Raw | ConvertFrom-Json)
    if ($regConf.Count -ne 2) { throw 'conferir: save perdeu registros' }
    if (($regConf | Where-Object { $_.pasta_inbox -eq $leaf1 }).sincronizada -ne $true) { throw 'conferir: save sobrescreveu sincronizada' }
    if (Test-Path -LiteralPath "$regPath.tmp") { throw 'conferir: .tmp residual' }

    # 1u) varredura física da inbox: lista TUDO do Drive; _processados.json só cruza o status
    $InboxBase = Join-Path $Tmp 'inbox3'
    $leafNova = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Registro\ATIVIDADE INDIVIDUAL 1\2026-11-19-NOVA - RESUMO'
    $leafPend = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Registro\ATIVIDADE INDIVIDUAL 1\2026-11-19-PENDENTE - RESUMO'
    $leafSync = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Registro\ATIVIDADE INDIVIDUAL 2\2026-11-19-SYNC - RESUMO'
    $leafVazia = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Registro\ATIVIDADE INDIVIDUAL 2\2026-11-19-SEM ARQUIVOS'
    foreach ($d in @($leafNova, $leafPend, $leafSync, $leafVazia)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    Set-Content -LiteralPath (Join-Path $leafNova 'relatorio.txt') -Encoding UTF8 -Value "---`ndata: 2026-11-19`n---`nCorpo NOVA"
    $regScan = Join-Path $InboxBase '2026\_inbox_relatorio_ater\_processados.json'
    @(
        [PSCustomObject]@{ pasta_inbox = $leafPend; sincronizada = $false },
        [PSCustomObject]@{ pasta_inbox = $leafSync; sincronizada = $true }
    ) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $regScan -Encoding UTF8
    $scan = Get-InboxPastasFisicas
    if (@($scan).Count -ne 4) { throw "varredura física: esperava 4 pastas (inclusive vazia), obtidas $(@($scan).Count)" }
    $stNova = $scan | Where-Object { $_.PastaInbox -eq $leafNova }
    $stPend = $scan | Where-Object { $_.PastaInbox -eq $leafPend }
    $stSync = $scan | Where-Object { $_.PastaInbox -eq $leafSync }
    $stVazia = $scan | Where-Object { $_.PastaInbox -eq $leafVazia }
    if ($stNova.Status -ne 'NOVA' -or $null -ne $stNova.Registro) { throw 'varredura física: pasta não registrada deveria ser NOVA' }
    if ($stPend.Status -ne 'TRIADA - PENDENTE SYNC') { throw 'varredura física: registrada sem sync deveria ser TRIADA - PENDENTE SYNC' }
    if ($stSync.Status -ne 'SINCRONIZADA') { throw 'varredura física: registrada com sync deveria ser SINCRONIZADA' }
    if ($stVazia.Status -ne 'NOVA') { throw 'varredura física: pasta vazia também deve ser listada' }
    if ($stPend.Registro.registro_path -ne $regScan) { throw 'varredura física: registro_path ausente no registro cruzado' }
    # Get-InboxVisitas continua idempotente: só a NOVA com .txt vira visita
    $visScan = Get-InboxVisitas
    if (@($visScan).Count -ne 1 -or $visScan[0].PastaInbox -ne $leafNova) { throw 'varredura física: Get-InboxVisitas deveria retornar só a visita NOVA' }
    # reprocessamento a partir do Drive: triagem da NOVA cria pasta padronizada + registro -> status vira PENDENTE
    $null = Invoke-CriarPastasEArquivar -RelatorioDias (ConvertTo-RelatorioDias -Visitas $visScan)
    $scan2 = Get-InboxPastasFisicas
    if (($scan2 | Where-Object { $_.PastaInbox -eq $leafNova }).Status -ne 'TRIADA - PENDENTE SYNC') { throw 'varredura física: após triagem, NOVA deveria virar TRIADA - PENDENTE SYNC' }
    $restScan = Get-InboxVisitas
    if (@($restScan).Count -ne 0) { throw 'varredura física: após triagem, nenhuma visita NOVA deveria restar' }

    # 1l) organizar-audios: mes pelo nome Xiaomi, ano/fallback pelo LastWriteTime
    if ((Get-PastaDestinoAudio '3 de nov., 11.39 Jose e Marta.m4a' ([datetime]'2026-11-03 11:39:00')) -ne '2026-11') { throw 'audios: padrao Xiaomi nov falhou' }
    if ((Get-PastaDestinoAudio '10 de jan., 08.32 Antonio .m4a' ([datetime]'2026-01-10 08:32:00')) -ne '2026-01') { throw 'audios: padrao Xiaomi jan falhou' }
    if ((Get-PastaDestinoAudio '31 de dez., 23.59.m4a' ([datetime]'2026-12-31 23:59:00')) -ne '2026-12') { throw 'audios: padrao Xiaomi dez falhou' }
    if ((Get-PastaDestinoAudio 'gravacao renomeada.m4a' ([datetime]'2026-05-10 09:00:00')) -ne '2026-05') { throw 'audios: fallback LastWriteTime falhou' }
    if ((Get-PastaDestinoAudio '3 de xyz., 11.39 teste.m4a' ([datetime]'2026-07-03 11:39:00')) -ne '2026-07') { throw 'audios: mes invalido deveria cair no fallback' }

    # 1t) trava anti-retroativo do Get-Produtores: só "- RESUMO" + .txt de relatório
    $nomesProd = @(Get-Produtores).Produtor
    if ($nomesProd -contains '04-2026-11-Juari') { throw 'trava: pasta sem - RESUMO não pode sincronizar (mesmo com .txt)' }
    if ($nomesProd -contains '05-2026-11-Coicieçao E Hermes - RESUMO') { throw 'trava: pasta - RESUMO só com fotos não pode sincronizar' }
    if ($nomesProd -contains '06-2026-11-Maria Da Graça - RESUMO') { throw 'trava: pasta - RESUMO só com coordenadas_gps.txt não pode sincronizar' }
    if (@(Get-Produtores).Count -ne 4) { throw "trava: esperava 4 pastas válidas, obtidas $(@(Get-Produtores).Count)" }
    $bloqueadas = @(Get-Produtores -IncluirBloqueadas | Where-Object { -not $_.Sincronizavel })
    if ($bloqueadas.Count -ne 3) { throw "trava: contagem informativa esperava 3 bloqueadas, obtidas $($bloqueadas.Count)" }

    # 2) caminho de destino montado (mesmo nome local no Drive)
    $P = @(Get-Produtores | Where-Object { $_.Atividade -eq 'ATIVIDADE INDIVIDUAL 1' -and $_.Produtor -eq '07-2026-11-JOAO DA SILVA - RESUMO' })[0]
    if ($P.Nuvem -ne 'gdrive:TESTE/ATIVIDADE INDIVIDUAL 1/Jacupiranga/07-2026-11-JOAO DA SILVA - RESUMO') { throw "nuvem errada: $($P.Nuvem)" }

    # 3) ignorar vale só para aquela atividade, não para o nome do produtor
    $Alvo = 'Jacupiranga\ATIVIDADE INDIVIDUAL 1\07-2026-11-JOAO DA SILVA - RESUMO'
    Save-Ignorados @($Alvo)
    if (@(Get-Ignorados).Count -ne 1) { throw 'ignorados.json não persistiu' }
    $Ignorados = Get-Ignorados
    $Restantes = @(Get-Produtores | Where-Object { $Ignorados -notcontains $_.Relativo })
    if ($Restantes.Count -ne 3) { throw "esperado 3 pastas sincronizáveis, obtido $($Restantes.Count)" }
    if (-not ($Restantes.Relativo -contains 'Jacupiranga\ATIVIDADE INDIVIDUAL 2\07-2026-11-JOAO DA SILVA - RESUMO')) { throw 'ATIVIDADE INDIVIDUAL 2 foi afetada' }

    # 4) ignorar de novo restaura a pasta
    Save-Ignorados (@(Get-Ignorados) | Where-Object { $_ -ne $Alvo })
    if (@(Get-Ignorados).Count -ne 0) { throw 'remoção do ignorado falhou' }

    # 1v) mapeamento cidade local [OD] -> nome no destino [GDSA] (Pariquera-Açu -> PARIQUERA AÇU)
    if ((Get-NomeCidadeRemoto 'Pariquera-Açu') -ne 'PARIQUERA AÇU') { throw 'mapa: Pariquera-Açu deveria virar PARIQUERA AÇU' }
    if ((Get-NomeCidadeRemoto 'Juquia') -ne 'Juquiá') { throw 'mapa: Juquia deveria cair em Juquiá' }
    if ((Get-NomeCidadeRemoto 'Cananéia') -ne 'Cananéia') { throw 'mapa: cidade sem mapeamento deveria manter o nome local' }
    $pPar = Join-Path $OrigemBase 'Pariquera-Açu\ATIVIDADE INDIVIDUAL 1\2026-11-19-TESTE MAPA - RESUMO'
    New-Item -ItemType Directory -Path $pPar -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $pPar 'relatorio.txt') -Encoding UTF8 -Value 'Corpo'
    $pm = @(Get-Produtores -ListaCidades @('Pariquera-Açu') | Where-Object { $_.Produtor -eq '2026-11-19-TESTE MAPA - RESUMO' })[0]
    if ($pm.Nuvem -ne 'gdrive:TESTE/ATIVIDADE INDIVIDUAL 1/PARIQUERA AÇU/19-2026-11-TESTE MAPA - RESUMO') { throw "mapa: Nuvem deveria usar PARIQUERA AÇU + data DD-AAAA-MM-: $($pm.Nuvem)" }
    if ($pm.Relativo -ne 'Pariquera-Açu\ATIVIDADE INDIVIDUAL 1\2026-11-19-TESTE MAPA - RESUMO') { throw 'mapa: Relativo (ignorados.json) deveria manter o nome local' }
    $Script:DriveDirCache = @{ 'gdrive:TESTE/ATIVIDADE INDIVIDUAL 1/PARIQUERA AÇU' = @([PSCustomObject]@{ Name = '19-2026-11-TESTE MAPA - RESUMO'; ID = 'IDMAPA123' }) }
    if ((Get-DriveFolderUrl -Atividade 'ATIVIDADE INDIVIDUAL 1' -Cidade 'Pariquera-Açu' -NomeDrive '19-2026-11-TESTE MAPA - RESUMO') -ne 'https://drive.google.com/drive/folders/IDMAPA123') { throw 'mapa: Get-DriveFolderUrl deveria usar o nome remoto da cidade' }

    # 1w) Copy-LoteParaVisita: copia o lote p/ [OD] + GPS + manifesto + status (usado pelo fluxo e pelo [E])
    $srcLote = Join-Path $Tmp 'srcLote'; $destLote = Join-Path $Tmp 'destLote'
    New-Item -ItemType Directory -Path $srcLote -Force | Out-Null
    New-Item -ItemType Directory -Path $destLote -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $srcLote 'IMG_20261107_090000.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $srcLote 'IMG_20261107_091000.jpg') | Out-Null
    $fotosLote = @(Get-ChildItem -LiteralPath $srcLote -File | ForEach-Object {
        [PSCustomObject]@{ Arquivo = $_; DataHora = [datetime]'2026-11-07 09:00:00'; DataStr = '2026-11-07'; HoraStr = '09:00:00'; Lat = $null; Lon = $null; TemGps = $false; LoteVisita = '2026\11\07_visita_1' }
    })
    $visLote = [PSCustomObject]@{ tag = 'A'; produtor_pasta = 'TESTE LOTE - RESUMO'; CaminhoDestino = $destLote }
    $rfLote = [System.Collections.Generic.List[object]]::new()
    $rgLote = [System.Collections.Generic.List[object]]::new()
    $stLote = @{}
    $mkLote = { param($tag, $modo, $n, $origem) $stLote[$tag] = [PSCustomObject]@{ Modo = $modo; Copiadas = $n; Origem = $origem } }
    Copy-LoteParaVisita -Lote $fotosLote -Visita $visLote -Modo 'automatico' -RunFotos $rfLote -RunGps $rgLote -Marcar $mkLote
    if (@(Get-ChildItem -LiteralPath $destLote -File -Filter '*.jpg').Count -ne 2) { throw 'lote: deveria copiar as 2 fotos do lote' }
    if (-not (Test-Path -LiteralPath (Join-Path $destLote 'coordenadas_gps.txt'))) { throw 'lote: coordenadas_gps.txt nao gerado' }
    if ($rfLote.Count -ne 2 -or $rgLote.Count -ne 1) { throw 'lote: manifesto de fotos/GPS incompleto' }
    if ($stLote['A'].Copiadas -ne 2 -or $stLote['A'].Modo -ne 'automatico' -or $stLote['A'].Origem -ne $srcLote) { throw 'lote: status (marcar) nao registrado' }

    # 1x) normalizacao Juquiá -> Juquia ao gravar no OneDrive (pasta com acento nao pode ser criada)
    if ((Get-NomeCidadeOneDrive 'Juquiá') -ne 'Juquia') { throw 'normalizacao: Juquiá deveria virar Juquia' }
    if ((Get-NomeCidadeOneDrive 'Registro') -ne 'Registro') { throw 'normalizacao: cidade sem mapeamento deveria manter o nome' }
    $vNorm = New-Visita -Data '2026-11-19' -Cidade 'Juquiá' -Atividade 'ATIVIDADE INDIVIDUAL 1' -ProdutorPasta 'TESTE NORM - RESUMO' -OrigemInbox $false
    $null = Invoke-CriarPastasEArquivar -RelatorioDias (ConvertTo-RelatorioDias -Visitas @($vNorm))
    if ($vNorm.CaminhoDestino -notmatch '\\Juquia\\') { throw "normalizacao: destino deveria ser Juquia: $($vNorm.CaminhoDestino)" }
    if (Test-Path -LiteralPath (Join-Path $OneDriveBase 'Juquiá')) { throw 'normalizacao: pasta com acento nao pode ser criada no OneDrive' }

    # 1y) inbox no formato novo DD-AAAA-MM- (app atualizado): data extraida do nome da pasta, sem duplicar prefixo
    $InboxBase = Join-Path $Tmp 'inbox4'
    $leafNovoFmt = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Juquia\ATIVIDADE INDIVIDUAL 1\07-2026-11-TESTE INBOX NOVO - RESUMO'
    New-Item -ItemType Directory -Path $leafNovoFmt -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $leafNovoFmt 'relatorio.txt') -Encoding UTF8 -Value 'Corpo sem YAML'
    $visNovoFmt = Get-InboxVisitas
    if (@($visNovoFmt).Count -ne 1) { throw 'inbox novo formato: deveria detectar 1 visita' }
    if ($visNovoFmt[0].data -ne '2026-11-07') { throw "inbox novo formato: data deveria vir do nome da pasta (2026-11-07), veio $($visNovoFmt[0].data)" }
    $null = Invoke-CriarPastasEArquivar -RelatorioDias (ConvertTo-RelatorioDias -Visitas $visNovoFmt)
    if ($visNovoFmt[0].CaminhoDestino -notmatch '\\Juquia\\ATIVIDADE INDIVIDUAL 1\\07-2026-11-TESTE INBOX NOVO - RESUMO$') { throw "inbox novo formato: destino incorreto: $($visNovoFmt[0].CaminhoDestino)" }

    # 1z) JSON schema 1.0 -> txt puro ANATER/Bem Viver v2 (sem YAML/Markdown/secao 10) + par YAML na inbox
    $objRico = [PSCustomObject]@{
        schema_version = '1.0'
        metadados = [PSCustomObject]@{ data = '2026-11-07'; periodo = 'manha'; hora_inicio = '09:00'; hora_fim = '11:00'; cidade = 'Juquia'; atividade = 'ATIVIDADE INDIVIDUAL 1'; produtor_pasta = '07-2026-11-SO JSON RICO - RESUMO' }
        institucional = [PSCustomObject]@{ entidade_executora = 'SA CONSULTORIA'; agente_ater = 'FERNANDO'; instrumento = 'ATER'; beneficiario = 'MARIA'; localizacao = 'Juquia'; data_etapa = '2026-11-07' }
        relatorio = [PSCustomObject]@{ eixo_produtivo = '**Solo** corrigido.'; orientacoes_tecnicas = @('Adubar cobertura', '- Plantar direto'); eixo_social = 'Familia participante.'; orientacoes_sociais = @('Curso dia 10'); eixo_ambiental = 'APP preservada.'; orientacoes_ambientais = @('Cercar nascente') }
        encaminhamentos = [PSCustomObject]@{ compromissos = 'NAO DEVE APARECER' }
    }
    $corpoJson = ConvertFrom-RelatorioJson -Json $objRico
    foreach ($esperado in @('7. EIXO PRODUTIVO', '8. EIXO SOCIAL', '9. EIXO AMBIENTAL', 'ORIENTAÇÕES TÉCNICAS:', 'ORIENTAÇÕES SOCIAIS:', 'ORIENTAÇÕES AMBIENTAIS:', '• Adubar cobertura', '• Plantar direto', 'SA CONSULTORIA', 'Agente ATER: FERNANDO')) {
        if ($corpoJson -notmatch [regex]::Escape($esperado)) { throw "json->txt: faltou '$esperado'" }
    }
    if ($corpoJson -match '\*\*|10\.|NAO DEVE APARECER') { throw 'json->txt: Markdown/secao 10 vazou no corpo' }
    if (-not $corpoJson.Contains("Solo corrigido.`n`nORIENTAÇÕES TÉCNICAS:")) { throw 'json->txt: quebra dupla entre blocos falhou' }
    $txt10 = "9. EIXO AMBIENTAL`n`nTexto ambiental.`n`n10. ENCAMINHAMENTOS`n`nCompromisso X.`n`n11. ANEXOS`n`nMapa."
    $sem10 = Remove-Secao10 $txt10
    if ($sem10 -match '10\.|Compromisso X' -or $sem10 -notmatch '11\. ANEXOS' -or $sem10 -notmatch 'Texto ambiental') { throw 'Remove-Secao10 falhou' }
    if ((Remove-Secao10 'Sem secao dez.') -ne 'Sem secao dez.') { throw 'Remove-Secao10 alterou texto sem secao 10' }
    # inbox so com .json rico: visita detectada com corpo montado; [OD] recebe txt limpo + anexo ATV1; inbox ganha par com YAML
    $InboxBase = Join-Path $Tmp 'inbox5'
    $leafJ = Join-Path $InboxBase '2026\_inbox_relatorio_ater\11\Juquia\ATIVIDADE INDIVIDUAL 1\07-2026-11-SO JSON RICO - RESUMO'
    New-Item -ItemType Directory -Path $leafJ -Force | Out-Null
    $objRico | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $leafJ 'relatorio.json') -Encoding UTF8
    New-Item -ItemType File -Path (Join-Path $leafJ '07-2026-11-SO JSON RICO - Mapa.png') | Out-Null
    $visJ = Get-InboxVisitas
    if (@($visJ).Count -ne 1 -or -not $visJ[0].TxtBody) { throw 'json rico: visita deveria ser detectada com corpo montado do JSON' }
    $null = Invoke-CriarPastasEArquivar -RelatorioDias (ConvertTo-RelatorioDias -Visitas $visJ)
    $destJ = $visJ[0].CaminhoDestino
    $txtJ = Get-Content -LiteralPath (Join-Path $destJ 'relatorio.txt') -Raw
    if ($txtJ -match '---|10\.|\*\*' -or $txtJ -notmatch '• ') { throw 'json rico: txt [OD] deveria ser limpo (sem YAML/secao 10) e com bullets' }
    if (Test-Path -LiteralPath (Join-Path $destJ 'relatorio.json')) { throw 'json rico: .json nao pode ir ao OneDrive' }
    if (-not (Test-Path -LiteralPath (Join-Path $destJ '07-2026-11-SO JSON RICO - Mapa.png'))) { throw 'json rico: anexo ATV1 nao copiado' }
    $txtInboxJ = Get-Content -LiteralPath (Join-Path $leafJ 'relatorio.txt') -Raw
    if ($txtInboxJ -notmatch '(?s)^---\r?\n.*?data: 2026-11-07.*?\r?\n---') { throw 'json rico: .txt da inbox deveria ter bloco YAML no topo' }
    # reparo sem .txt (destino e inbox): Repair reconstroi do JSON, sem duplicar registro
    Remove-Item -LiteralPath (Join-Path $destJ 'relatorio.txt') -Force
    Remove-Item -LiteralPath (Join-Path $leafJ 'relatorio.txt') -Force
    $regJ = @(Get-ProcessadosTodos | Where-Object { $_.pasta_inbox -eq $leafJ })
    if ($regJ.Count -ne 1) { throw 'json rico: registro unico esperado' }
    $null = Repair-VisitaTxtAnexos -Registro $regJ[0]
    if (-not (Test-Path -LiteralPath (Join-Path $destJ 'relatorio.txt'))) { throw 'json rico: reparo deveria reconstruir o txt a partir do JSON' }
    if (@(Get-ProcessadosTodos | Where-Object { $_.pasta_inbox -eq $leafJ }).Count -ne 1) { throw 'json rico: reparo duplicou registro' }

    # 1aa) Get-HorarioMedioLote: media em minutos + HH:mm + faixa inicial..final do lote
    $loteHM = @(
        [PSCustomObject]@{ DataHora = [datetime]'2026-11-07 11:40:00' },
        [PSCustomObject]@{ DataHora = [datetime]'2026-11-07 12:00:00' }
    )
    $hm = Get-HorarioMedioLote -Lote $loteHM
    if ($hm.MediaMinutos -ne 710 -or $hm.MediaHora -ne '11:50' -or $hm.Faixa -ne '11:40 as 12:00') { throw 'horario medio: media/minutos/faixa falhou' }
    if ($null -ne (Get-HorarioMedioLote -Lote @())) { throw 'horario medio: lote vazio deveria ser nulo' }

    # 1ab) Get-SugestaoTag com hora_fim null (relatorio de audio): visita unica mais proxima em <= 90 min
    $vAudioA = [PSCustomObject]@{ tag = 'A'; periodo = $null; hora_inicio = '11:54'; hora_fim = $null }
    $vAudioB = [PSCustomObject]@{ tag = 'B'; periodo = $null; hora_inicio = '15:00'; hora_fim = $null }
    if ((Get-SugestaoTag -Lote $loteHM -Visitas @($vAudioA, $vAudioB)) -ne 'A') { throw 'sugestao so hora_inicio: visita mais proxima falhou' }
    $vEmp1 = [PSCustomObject]@{ tag = 'A'; periodo = $null; hora_inicio = '11:20'; hora_fim = $null }
    $vEmp2 = [PSCustomObject]@{ tag = 'B'; periodo = $null; hora_inicio = '12:20'; hora_fim = $null }
    if ($null -ne (Get-SugestaoTag -Lote $loteHM -Visitas @($vEmp1, $vEmp2))) { throw 'sugestao so hora_inicio: empate equidistante deveria ser ambiguo (nulo)' }
    $vLonge = [PSCustomObject]@{ tag = 'A'; periodo = 'manha'; hora_inicio = '11:54'; hora_fim = $null }
    $loteCedo = @([PSCustomObject]@{ DataHora = [datetime]'2026-11-07 08:00:00' })
    if ($null -ne (Get-SugestaoTag -Lote $loteCedo -Visitas @($vLonge))) { throw 'sugestao: fora da janela de 90 min nao pode cair no periodo generico' }
    $vJanela = [PSCustomObject]@{ tag = 'C'; periodo = $null; hora_inicio = '12:00'; hora_fim = '13:00' }
    if ((Get-SugestaoTag -Lote $loteHM -Visitas @($vJanela)) -ne 'C') { throw 'sugestao intervalo: margem de -30 min antes do inicio falhou' }
    $loteFora = @([PSCustomObject]@{ DataHora = [datetime]'2026-11-07 11:20:00' })
    if ($null -ne (Get-SugestaoTag -Lote $loteFora -Visitas @($vJanela))) { throw 'sugestao intervalo: media antes de ini-30 nao pode casar' }

    # 1ac) multiplos lotes no mesmo dia (07_visita_1/07_visita_2): pastas NAO se misturam;
    #      sugestao por horario distingue o lote certo; -Conferir restringe a pasta escolhida.
    #      (A TUI do seletor em si exige ReadKey e nao roda neste ambiente nao-interativo.)
    $rootMulti = Join-Path $Tmp 'syncthingMulti'
    $d1 = Join-Path $rootMulti '2026\11\07_visita_1'; $d2 = Join-Path $rootMulti '2026\11\07_visita_2'
    New-Item -ItemType Directory -Path $d1 -Force | Out-Null
    New-Item -ItemType Directory -Path $d2 -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $d1 'IMG_20261107_090000.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $d1 'IMG_20261107_091000.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $d2 'IMG_20261107_114000.jpg') | Out-Null
    New-Item -ItemType File -Path (Join-Path $d2 'IMG_20261107_120000.jpg') | Out-Null
    ([PSCustomObject]@{ data = '2026-11-07'; arquivos = @(
        [PSCustomObject]@{ nome = 'IMG_20261107_090000.jpg'; hora = '09:00:00'; lat = -23.5; lon = -47.25 },
        [PSCustomObject]@{ nome = 'IMG_20261107_091000.jpg'; hora = '09:10:00'; lat = -23.5; lon = -47.25 }) }) |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $d1 '_visita.json') -Encoding UTF8
    ([PSCustomObject]@{ data = '2026-11-07'; arquivos = @(
        [PSCustomObject]@{ nome = 'IMG_20261107_114000.jpg'; hora = '11:40:00'; lat = -23.6; lon = -47.3 },
        [PSCustomObject]@{ nome = 'IMG_20261107_120000.jpg'; hora = '12:00:00'; lat = -23.6; lon = -47.3 }) }) |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $d2 '_visita.json') -Encoding UTF8
    $fm = Get-FotosCirurgico -Raizes @($rootMulti) -DatasAlvo @('2026-11-07') -Extensoes $ExtensoesFoto
    if ($fm.Count -ne 4) { throw "multi-lote: esperava 4 fotos, obtidas $($fm.Count)" }
    $gruposDia = @($fm | Group-Object -Property LoteVisita)
    if ($gruposDia.Count -ne 2) { throw 'multi-lote: pastas do mesmo dia devem permanecer separadas (2 lotes)' }
    $loteM1 = @($gruposDia | Where-Object { $_.Name -eq '2026\11\07_visita_1' } | ForEach-Object { $_.Group })
    $loteM2 = @($gruposDia | Where-Object { $_.Name -eq '2026\11\07_visita_2' } | ForEach-Object { $_.Group })
    $vDia = [PSCustomObject]@{ tag = 'A'; periodo = $null; hora_inicio = '11:54'; hora_fim = $null }
    if ((Get-SugestaoTag -Lote $loteM2 -Visitas @($vDia)) -ne 'A') { throw 'multi-lote: lote de 11:40-12:00 deveria sugerir A (relatorio 11:54)' }
    if ($null -ne (Get-SugestaoTag -Lote $loteM1 -Visitas @($vDia))) { throw 'multi-lote: lote da manha nao pode sugerir a visita de 11:54' }
    # -Conferir: sem origem o dia inteiro e o material do seletor; com a pasta escolhida gravada
    # em fotos_origem_sugerida, somente aquela pasta entra (nada de copiar o dia todo).
    $destMulti = Join-Path $Tmp 'destMulti'; New-Item -ItemType Directory -Path $destMulti -Force | Out-Null
    $regMulti = [PSCustomObject]@{ data = '2026-11-07'; destino_onedrive = $destMulti; produtor_pasta = 'TESTE MULTI - RESUMO'; fotos_origem_sugerida = $null }
    $synctAntes = $SyncthingFotos
    $SyncthingFotos = @($rootMulti)
    $faltMulti = Get-FotosFaltantesVisita -Registro $regMulti
    if (@($faltMulti).Count -ne 4) { throw 'conferir multi-lote: sem origem, o dia inteiro deveria ser listado p/ o seletor' }
    Add-Member -InputObject $regMulti -NotePropertyName fotos_origem_sugerida -NotePropertyValue $d2 -Force
    $faltSo2 = Get-FotosFaltantesVisita -Registro $regMulti
    if (@($faltSo2).Count -ne 2) { throw 'conferir multi-lote: origem escolhida deveria restringir a pasta' }
    if (@($faltSo2 | Where-Object { (Split-Path $_.Arquivo.FullName -Parent) -ne $d2 }).Count -ne 0) { throw 'conferir multi-lote: fotos de outra pasta do dia vazaram' }
    $SyncthingFotos = $synctAntes

    # 1ad) piloto -Limpar: Remove-ProcessadosPiloto tira SO os registros TESTE PILOTO do
    #      _processados.json (atomico, preserva os demais; [] valido quando nao sobra nada)
    $anoPiloto = Join-Path $Tmp 'inboxPiloto\2026\_inbox_relatorio_ater'
    New-Item -ItemType Directory -Path $anoPiloto -Force | Out-Null
    $regPiloto = Join-Path $anoPiloto '_processados.json'
    $leafPilotoT = Join-Path $anoPiloto '11\Jacupiranga\ATIVIDADE INDIVIDUAL 1\2026-11-19-TESTE PILOTO - RESUMO'
    @(
        [PSCustomObject]@{ pasta_inbox = $leafPilotoT; destino_onedrive = 'C:\od\TESTE PILOTO - RESUMO'; sincronizada = $false },
        [PSCustomObject]@{ pasta_inbox = 'G:\real\JOAO - RESUMO'; destino_onedrive = 'C:\od\JOAO - RESUMO'; sincronizada = $true }
    ) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $regPiloto -Encoding UTF8
    $nRem = Remove-ProcessadosPiloto -PastaInboxAno $anoPiloto -PastaInbox $leafPilotoT -DestinoOD 'C:\od\TESTE PILOTO - RESUMO'
    if ($nRem -ne 1) { throw "piloto limpar: deveria remover 1 registro, removeu $nRem" }
    $regSobrou = @(Get-Content -LiteralPath $regPiloto -Raw | ConvertFrom-Json)
    if ($regSobrou.Count -ne 1 -or $regSobrou[0].destino_onedrive -ne 'C:\od\JOAO - RESUMO' -or $regSobrou[0].sincronizada -ne $true) { throw 'piloto limpar: registro real nao preservado' }
    if (Test-Path -LiteralPath "$regPiloto.tmp") { throw 'piloto limpar: .tmp residual' }
    if ((Remove-ProcessadosPiloto -PastaInboxAno $anoPiloto -PastaInbox $leafPilotoT -DestinoOD 'C:\od\TESTE PILOTO - RESUMO') -ne 0) { throw 'piloto limpar: segunda chamada deveria remover 0 (idempotente)' }
    # arquivo com SO o registro do piloto -> vira [] valido (leituras futuras nao quebram)
    @([PSCustomObject]@{ pasta_inbox = $leafPilotoT; destino_onedrive = 'C:\od\TESTE PILOTO - RESUMO' }) | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $regPiloto -Encoding UTF8
    $null = Remove-ProcessadosPiloto -PastaInboxAno $anoPiloto -PastaInbox $leafPilotoT -DestinoOD 'C:\od\TESTE PILOTO - RESUMO'
    if (@(Get-Content -LiteralPath $regPiloto -Raw | ConvertFrom-Json).Count -ne 0) { throw 'piloto limpar: arquivo sem registros deveria ser lista vazia valida' }
    # caminho inexistente -> 0, sem criar arquivo
    if ((Remove-ProcessadosPiloto -PastaInboxAno (Join-Path $Tmp 'naoExiste') -PastaInbox $leafPilotoT -DestinoOD 'C:\od\X') -ne 0) { throw 'piloto limpar: sem _processados.json deveria retornar 0' }

    $Passou = $true
} finally {
    Remove-Item $Tmp -Recurse -Force -ErrorAction SilentlyContinue
    if ($Passou) { Write-Host '[SELF-CHECK] Ambiente temporario removido. OK - self-check passed.' -ForegroundColor Green }
}