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

# Piloto ponta a ponta: valida sintaxe (6º script; não executa o fluxo real nem carrega funções p/ evitar colisões)
$ScriptPiloto = Join-Path $PSScriptRoot 'piloto-ponta-a-ponta.ps1'
$ErrosP = $null
$AstP = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPiloto, [ref]$null, [ref]$ErrosP)
if ($ErrosP.Count) { throw "Sintaxe piloto: $($ErrosP[0].Message)" }

$Tmp = Join-Path ([IO.Path]::GetTempPath()) ("ater_check_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $Tmp | Out-Null
$Passou = $false
try {
    $OrigemBase    = Join-Path $Tmp 'Fernando'
    $Cidades       = @('Jacupiranga')
    $Atividades    = @('ATIVIDADE INDIVIDUAL 1', 'ATIVIDADE INDIVIDUAL 2')
    $RemotoDrive   = 'gdrive:TESTE'
    $IgnoradosFile = Join-Path $Tmp 'ignorados.json'

    foreach ($Ativ in $Atividades) {
        foreach ($Nome in @('2026-11-JOAO DA SILVA', '03 - Pedro Oliveira')) {
            New-Item -ItemType Directory -Path (Join-Path $OrigemBase "Jacupiranga\$Ativ\$Nome") -Force | Out-Null
        }
    }

    # 1) renomeação no destino
    if ((Get-NomeResumo '2026-11-JOAO DA SILVA') -ne 'JOAO DA SILVA - RESUMO') { throw 'renomeação com data falhou' }
    if ((Get-NomeResumo '03 - Pedro Oliveira') -ne 'Pedro Oliveira - RESUMO') { throw 'renomeação com número falhou' }
    if ((Get-NomeResumo 'SEM PREFIXO') -ne 'SEM PREFIXO - RESUMO') { throw 'renomeação sem prefixo falhou' }

    # 1b) idempotência do - RESUMO no start-sync (não duplicar)
    if ((Get-NomeResumo '2026-11-CARLOS - ITAPITANGUI - RESUMO') -ne 'CARLOS - ITAPITANGUI - RESUMO') { throw 'idempotência: duplicou RESUMO com prefixo de data' }
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

    # 1d) Format-NomePastaProdutor: prefixo AAAA-MM-DD- (dia completo) + sufixo - RESUMO
    if ((Format-NomePastaProdutor -Nome 'CARLOS - ITAPITANGUI' -Data '2026-11-12') -ne '2026-11-12-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: prefixo+sufixo não garantidos' }
    if ((Format-NomePastaProdutor -Nome '2026-11-12-CARLOS - ITAPITANGUI - RESUMO' -Data '2026-11-12') -ne '2026-11-12-CARLOS - ITAPITANGUI - RESUMO') { throw 'triagem: duplicou prefixo/sufixo' }
    if ((Format-NomePastaProdutor -Nome '2026-11-CARLOS' -Data '2026-11-12') -ne '2026-11-12-CARLOS - RESUMO') { throw 'triagem: prefixo parcial não normalizado' }

    # 1e) start-sync remove prefixo de data COMPLETA (AAAA-MM-DD-) tambem
    if ((Get-NomeResumo '2026-11-12-CARLOS - ITAPITANGUI - RESUMO') -ne 'CARLOS - ITAPITANGUI - RESUMO') { throw 'start-sync: não removeu dia completo' }

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
    if ($dias[0].visitas[1].produtor_pasta -ne '2026-11-12-ANA - RESUMO') { throw 'triagem: nome de pasta não normalizado' }

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
    $Script:DriveDirCache = @{ "$RemotoDrive/ATIVIDADE INDIVIDUAL 1/Jacupiranga" = @([PSCustomObject]@{ Name = 'JOAO - RESUMO'; ID = 'IDJOAO123' }) }
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

    # 1l) organizar-audios: mes pelo nome Xiaomi, ano/fallback pelo LastWriteTime
    if ((Get-PastaDestinoAudio '3 de nov., 11.39 Jose e Marta.m4a' ([datetime]'2026-11-03 11:39:00')) -ne '2026-11') { throw 'audios: padrao Xiaomi nov falhou' }
    if ((Get-PastaDestinoAudio '10 de jan., 08.32 Antonio .m4a' ([datetime]'2026-01-10 08:32:00')) -ne '2026-01') { throw 'audios: padrao Xiaomi jan falhou' }
    if ((Get-PastaDestinoAudio '31 de dez., 23.59.m4a' ([datetime]'2026-12-31 23:59:00')) -ne '2026-12') { throw 'audios: padrao Xiaomi dez falhou' }
    if ((Get-PastaDestinoAudio 'gravacao renomeada.m4a' ([datetime]'2026-05-10 09:00:00')) -ne '2026-05') { throw 'audios: fallback LastWriteTime falhou' }
    if ((Get-PastaDestinoAudio '3 de xyz., 11.39 teste.m4a' ([datetime]'2026-07-03 11:39:00')) -ne '2026-07') { throw 'audios: mes invalido deveria cair no fallback' }

    # 2) caminho de destino montado
    $P = @(Get-Produtores | Where-Object { $_.Atividade -eq 'ATIVIDADE INDIVIDUAL 1' -and $_.Produtor -eq '2026-11-JOAO DA SILVA' })[0]
    if ($P.Nuvem -ne 'gdrive:TESTE/ATIVIDADE INDIVIDUAL 1/Jacupiranga/JOAO DA SILVA - RESUMO') { throw "nuvem errada: $($P.Nuvem)" }

    # 3) ignorar vale só para aquela atividade, não para o nome do produtor
    $Alvo = 'Jacupiranga\ATIVIDADE INDIVIDUAL 1\2026-11-JOAO DA SILVA'
    Save-Ignorados @($Alvo)
    if (@(Get-Ignorados).Count -ne 1) { throw 'ignorados.json não persistiu' }
    $Ignorados = Get-Ignorados
    $Restantes = @(Get-Produtores | Where-Object { $Ignorados -notcontains $_.Relativo })
    if ($Restantes.Count -ne 3) { throw "esperado 3 pastas sincronizáveis, obtido $($Restantes.Count)" }
    if (-not ($Restantes.Relativo -contains 'Jacupiranga\ATIVIDADE INDIVIDUAL 2\2026-11-JOAO DA SILVA')) { throw 'ATIVIDADE INDIVIDUAL 2 foi afetada' }

    # 4) ignorar de novo restaura a pasta
    Save-Ignorados (@(Get-Ignorados) | Where-Object { $_ -ne $Alvo })
    if (@(Get-Ignorados).Count -ne 0) { throw 'remoção do ignorado falhou' }

    $Passou = $true
} finally {
    Remove-Item $Tmp -Recurse -Force -ErrorAction SilentlyContinue
    if ($Passou) { Write-Host '[SELF-CHECK] Ambiente temporario removido. OK - self-check passed.' -ForegroundColor Green }
}