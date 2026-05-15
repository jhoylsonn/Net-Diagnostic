<#
Invoke-NetDiagnostic - Single file build
Build: 05/14/2026 18:12:26
#>

# ============================================================
# FILE: 00-Globals.ps1
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:AppName = 'Invoke-NetDiagnostic'
$Script:Version = '1.0-preview5'
$Script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ([string]::IsNullOrWhiteSpace($Script:BaseDir)) {
    $Script:BaseDir = (Get-Location).Path
}

$Script:ReportsDir = Join-Path $Script:BaseDir 'Reports'
$Script:LastResult = $null
$Script:CurrentRun = $null

if (-not (Test-Path $Script:ReportsDir)) {
    New-Item -Path $Script:ReportsDir -ItemType Directory -Force | Out-Null
}

function Write-Section {
    param([string]$Title)

    Write-Host ''
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
}

function Write-Status {
    param(
        [ValidateSet('INFO','OK','ALERTA','CRITICO','FALHA','INSTABILIDADE')]
        [string]$Level,
        [string]$Message
    )

    $color = switch ($Level) {
        'INFO'           { 'Gray' }
        'OK'             { 'Green' }
        'ALERTA'         { 'Yellow' }
        'INSTABILIDADE'  { 'Yellow' }
        'CRITICO'        { 'Magenta' }
        'FALHA'          { 'Red' }
        default          { 'White' }
    }

    $timestamp = Get-Date -Format 'HH:mm:ss'
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-TimeStampFileName {
    param([string]$Prefix = 'NetDiagnostic')

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    "$Prefix-$stamp"
}

# ============================================================
# FILE: Events.ps1
# ============================================================

function New-NetDiagEvent {
    param(
        [ValidateSet('INFO','ALERTA','CRITICO','FALHA')]
        [string]$Level,

        [string]$Type,

        [string[]]$Symptoms,

        [string]$Details = '',

        [datetime]$Time = (Get-Date)
    )

    [pscustomobject]@{
        Time     = $Time
        Level    = $Level
        Type     = $Type
        Symptoms = $Symptoms
        Details  = $Details
    }
}

function Test-NetDiagBadPingStatus {
    param([object]$Result)

    if (-not $Result) { return $false }

    return (
        $Result.Status -in @('ALERTA','CRITICO','FALHA') -or
        ($Result.LossPercent -ne $null -and [double]$Result.LossPercent -gt 0) -or
        ($Result.AvgMs -ne $null -and [double]$Result.AvgMs -ge 300)
    )
}

function Get-EventsFromSample {
    param(
        [object[]]$PingResults = @(),
        [object[]]$DnsResults = @()
    )

    $events = New-Object System.Collections.Generic.List[object]

    foreach ($p in @($PingResults)) {
        $eventTime = if ($p.Time) { [datetime]$p.Time } else { Get-Date }

        if ($p.Status -eq 'SEM_ICMP') {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'INFO' `
                -Type 'DESTINO SEM RESPOSTA ICMP' `
                -Symptoms @("$($p.Target) nao respondeu ping/ICMP", 'DNS/TCP devem ser usados como validacao principal para este destino') `
                -Details "Target=$($p.Target)")) | Out-Null
        }
        elseif ($p.Status -eq 'FALHA') {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'FALHA' `
                -Type "FALHA EM $($p.Layer)" `
                -Symptoms @(
                    "Falha temporaria detectada em $($p.Target)",
                    "Perda de pacotes observada: $($p.LossPercent)%",
                    'Evento ocorreu em uma amostra especifica do monitoramento'
                ) `
                -Details "Target=$($p.Target)")) | Out-Null
        }
        elseif ($p.AvgMs -ne $null -and [double]$p.AvgMs -ge 300) {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'CRITICO' `
                -Type "LATENCIA CRITICA EM $($p.Layer)" `
                -Symptoms @("Latencia media acima de 300 ms", "Media: $($p.AvgMs) ms", "Maxima: $($p.MaxMs) ms") `
                -Details "Target=$($p.Target)")) | Out-Null
        }
        elseif ($p.AvgMs -ne $null -and [double]$p.AvgMs -ge 150) {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'ALERTA' `
                -Type "LATENCIA ALTA EM $($p.Layer)" `
                -Symptoms @("Latencia media acima de 150 ms", "Media: $($p.AvgMs) ms") `
                -Details "Target=$($p.Target)")) | Out-Null
        }
        elseif ($p.LossPercent -ne $null -and [double]$p.LossPercent -gt 0) {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'ALERTA' `
                -Type "PERDA DE PACOTES EM $($p.Layer)" `
                -Symptoms @("Perda de pacotes: $($p.LossPercent)%") `
                -Details "Target=$($p.Target)")) | Out-Null
        }
    }

    foreach ($d in @($DnsResults)) {
        $eventTime = if ($d.Time) { [datetime]$d.Time } else { Get-Date }

        if (-not $d.Success) {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'FALHA' `
                -Type 'FALHA DNS' `
                -Symptoms @("Falha ao resolver $($d.Name)", "Servidor DNS: $($d.Server)") `
                -Details 'Resolve-DnsName falhou')) | Out-Null
        }
        elseif ($d.LatencyMs -ne $null -and [double]$d.LatencyMs -ge 1000) {
            $events.Add((New-NetDiagEvent `
                -Time $eventTime `
                -Level 'ALERTA' `
                -Type 'DNS LENTO' `
                -Symptoms @("Resolucao DNS acima de 1000 ms", "Tempo: $($d.LatencyMs) ms") `
                -Details "Name=$($d.Name)")) | Out-Null
        }
    }

    @($events.ToArray())
}

function Get-PersistentTimelineEvents {
    param(
        [object[]]$PingResults = @(),
        [object[]]$ExistingEvents = @()
    )

    $events = New-Object System.Collections.Generic.List[object]
    $pings = @($PingResults | Where-Object { $_ -and $_.Time } | Sort-Object Time)

    if ($pings.Count -eq 0) {
        return @()
    }

    $existingKeys = @{}
    foreach ($e in @($ExistingEvents)) {
        $existingKeys["$($e.Type)|$($e.Details)|$($e.Time)"] = $true
    }

    # Recuperacao automatica por alvo: FALHA seguida de OK no mesmo Target.
    foreach ($target in @($pings | Select-Object -ExpandProperty Target -Unique)) {
        $items = @($pings | Where-Object { $_.Target -eq $target } | Sort-Object Time)
        $openFailure = $null

        foreach ($item in $items) {
            if ($item.Status -eq 'FALHA') {
                if (-not $openFailure) { $openFailure = $item }
                continue
            }

            if ($openFailure -and $item.Status -eq 'OK') {
                $key = "RECUPERACAO|Target=$target|$($item.Time)"
                if (-not $existingKeys.ContainsKey($key)) {
                    $events.Add((New-NetDiagEvent `
                        -Time ([datetime]$item.Time) `
                        -Level 'INFO' `
                        -Type 'RECUPERACAO' `
                        -Symptoms @(
                            "$($openFailure.Layer) $target voltou a responder",
                            "Falha anterior detectada em $($openFailure.Time)"
                        ) `
                        -Details "Target=$target")) | Out-Null
                }
                $openFailure = $null
            }
        }
    }

    $bad = @($pings | Where-Object { Test-NetDiagBadPingStatus $_ })
    if ($bad.Count -eq 0) {
        return @($events.ToArray())
    }

    # Janela curta de 60s para detectar padrao forte.
    $windowSeconds = 60
    $strongLoopDetected = $false
    $loopTime = $null

    foreach ($point in $bad) {
        $start = ([datetime]$point.Time).AddSeconds(-1 * $windowSeconds)
        $end   = ([datetime]$point.Time).AddSeconds($windowSeconds)
        $window = @($bad | Where-Object { ([datetime]$_.Time) -ge $start -and ([datetime]$_.Time) -le $end })

        $gatewayFailures = @($window | Where-Object { $_.Layer -eq 'Gateway' -and $_.Status -eq 'FALHA' }).Count
        $internetFailures = @($window | Where-Object { $_.Layer -eq 'InternetIP' -and $_.Status -eq 'FALHA' }).Count
        $affectedTargets = @($window | Select-Object -ExpandProperty Target -Unique).Count
        $highLatencyCount = @($window | Where-Object { $_.AvgMs -ne $null -and [double]$_.AvgMs -ge 300 }).Count
        $lossCount = @($window | Where-Object { $_.LossPercent -ne $null -and [double]$_.LossPercent -gt 0 }).Count

        # POSSIVEL LOOP somente com padrao forte:
        # multiplas ocorrencias, varios destinos, gateway recorrente, perda generalizada e reincidencia em janela curta.
        if (
            $gatewayFailures -ge 3 -and
            $internetFailures -ge 3 -and
            $affectedTargets -ge 3 -and
            ($lossCount -ge 6 -or $highLatencyCount -ge 3)
        ) {
            $strongLoopDetected = $true
            $loopTime = [datetime]$point.Time
            break
        }
    }

    if ($strongLoopDetected) {
        $events.Add((New-NetDiagEvent `
            -Time $loopTime `
            -Level 'CRITICO' `
            -Type 'POSSIVEL LOOP' `
            -Symptoms @(
                'multiplas ocorrencias consecutivas',
                'falha em varios destinos simultaneamente',
                'gateway oscilando repetidamente',
                'perda generalizada',
                'latencia explosiva continua ou recorrente',
                'reincidencia em janela curta'
            ) `
            -Details 'Padrao forte de instabilidade local recorrente detectado')) | Out-Null
    }
    else {
        # Caso comum: microqueda ou oscilacao breve. Nao chamar de loop.
        foreach ($point in $bad) {
            $start = ([datetime]$point.Time).AddSeconds(-5)
            $end   = ([datetime]$point.Time).AddSeconds(5)
            $near = @($bad | Where-Object { ([datetime]$_.Time) -ge $start -and ([datetime]$_.Time) -le $end })
            $hasGateway = @($near | Where-Object { $_.Layer -eq 'Gateway' }).Count -gt 0
            $hasInternet = @($near | Where-Object { $_.Layer -eq 'InternetIP' }).Count -gt 0

            if ($hasGateway -or $hasInternet) {
                $firstTime = [datetime]($near | Sort-Object Time | Select-Object -First 1).Time
                $key = "INSTABILIDADE MOMENTANEA/INTERMITENTE|microqueda|$firstTime"
                if (-not $existingKeys.ContainsKey($key)) {
                    $events.Add((New-NetDiagEvent `
                        -Time $firstTime `
                        -Level 'ALERTA' `
                        -Type 'INSTABILIDADE MOMENTANEA/INTERMITENTE' `
                        -Symptoms @(
                            'Instabilidade momentanea/intermitente detectada',
                            'Possivel microqueda local ou perda transitoria de conectividade',
                            'Nao ha evidencia suficiente para classificar como loop de rede'
                        ) `
                        -Details 'microqueda')) | Out-Null
                }
                break
            }
        }
    }

    @($events.ToArray())
}

# ============================================================
# FILE: NetworkInfo.ps1
# ============================================================

function Get-ActiveNetworkInfo {
    $ipConfigs = Get-NetIPConfiguration | Where-Object {
        $_.IPv4Address -and $_.NetAdapter.Status -eq 'Up'
    }

    $selected = $ipConfigs | Where-Object { $_.IPv4DefaultGateway } | Select-Object -First 1

    if (-not $selected) {
        $selected = $ipConfigs | Select-Object -First 1
    }

    if (-not $selected) {
        return [pscustomobject]@{
            Found          = $false
            InterfaceAlias = $null
            InterfaceIndex = $null
            IPv4Address    = $null
            PrefixLength   = $null
            Gateway        = $null
            DnsServers     = @()
            MacAddress     = $null
            LinkSpeed      = $null
        }
    }

    $adapter = Get-NetAdapter -InterfaceIndex $selected.InterfaceIndex

    [pscustomobject]@{
        Found          = $true
        InterfaceAlias = $selected.InterfaceAlias
        InterfaceIndex = $selected.InterfaceIndex
        IPv4Address    = $selected.IPv4Address.IPAddress
        PrefixLength   = $selected.IPv4Address.PrefixLength
        Gateway        = $selected.IPv4DefaultGateway.NextHop
        DnsServers     = @($selected.DNSServer.ServerAddresses | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' })
        MacAddress     = $adapter.MacAddress
        LinkSpeed      = $adapter.LinkSpeed
    }
}

# ============================================================
# FILE: PingTests.ps1
# ============================================================

function Test-PingTarget {
    param(
        [Parameter(Mandatory)]
        [string]$Target,

        [int]$Count = 4,

        [int]$TimeoutSeconds = 2,

        [string]$Layer = 'Custom'
    )

    $latencies = New-Object System.Collections.Generic.List[int]
    $success = 0
    $fail = 0
    $timeoutMs = [math]::Max(1, $TimeoutSeconds) * 1000

    $ping = [System.Net.NetworkInformation.Ping]::new()

    for ($i = 1; $i -le $Count; $i++) {
        try {
            $reply = $ping.Send($Target, $timeoutMs)

            if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $success++
                $latencies.Add([int]$reply.RoundtripTime) | Out-Null
            }
            else {
                $fail++
            }
        }
        catch {
            $fail++
        }
    }

    $lossPercent = if ($Count -gt 0) { [math]::Round(($fail / $Count) * 100, 2) } else { 100 }
    $avg = if ($latencies.Count -gt 0) { [math]::Round(($latencies | Measure-Object -Average).Average, 2) } else { $null }
    $min = if ($latencies.Count -gt 0) { ($latencies | Measure-Object -Minimum).Minimum } else { $null }
    $max = if ($latencies.Count -gt 0) { ($latencies | Measure-Object -Maximum).Maximum } else { $null }

    $status = 'OK'

    if ($lossPercent -ge 100) {
        $status = 'FALHA'
    }
    elseif ($avg -ne $null -and $avg -ge 300) {
        $status = 'CRITICO'
    }
    elseif ($lossPercent -gt 0 -or ($avg -ne $null -and $avg -ge 150)) {
        $status = 'ALERTA'
    }

    [pscustomobject]@{
        Time        = Get-Date
        Layer       = $Layer
        Target      = $Target
        Sent        = $Count
        Received    = $success
        Lost        = $fail
        LossPercent = $lossPercent
        MinMs       = $min
        AvgMs       = $avg
        MaxMs       = $max
        Status      = $status
    }
}

# ============================================================
# FILE: DnsTests.ps1
# ============================================================

function Test-DnsResolution {
    param(
        [string]$Name = 'google.com',
        [string]$Server = $null
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        if ($Server) {
            $result = Resolve-DnsName -Name $Name -Server $Server -ErrorAction Stop
        }
        else {
            $result = Resolve-DnsName -Name $Name -ErrorAction Stop
        }

        $sw.Stop()

        [pscustomobject]@{
            Time      = Get-Date
            Name      = $Name
            Server    = $Server
            Success   = $true
            LatencyMs = $sw.ElapsedMilliseconds
            Answer    = ($result | Where-Object { $_.IPAddress } | Select-Object -First 3 -ExpandProperty IPAddress) -join ', '
            Status    = if ($sw.ElapsedMilliseconds -ge 1000) { 'ALERTA' } else { 'OK' }
        }
    }
    catch {
        $sw.Stop()

        [pscustomobject]@{
            Time      = Get-Date
            Name      = $Name
            Server    = $Server
            Success   = $false
            LatencyMs = $sw.ElapsedMilliseconds
            Answer    = ''
            Status    = 'FALHA'
        }
    }
}

# ============================================================
# FILE: TcpTests.ps1
# ============================================================

function Test-TcpPort {
    param(
        [string]$ComputerName,
        [int]$Port,
        [string]$Label = ''
    )

    try {
        $result = Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet

        [pscustomobject]@{
            Time         = Get-Date
            ComputerName = $ComputerName
            Port         = $Port
            Label        = $Label
            Success      = [bool]$result
            Status       = if ($result) { 'OK' } else { 'FALHA' }
        }
    }
    catch {
        [pscustomobject]@{
            Time         = Get-Date
            ComputerName = $ComputerName
            Port         = $Port
            Label        = $Label
            Success      = $false
            Status       = 'FALHA'
        }
    }
}

# ============================================================
# FILE: TraceRoute.ps1
# ============================================================

function Invoke-TraceRouteSimple {
    param([string]$Target)

    try {
        $output = tracert -d -h 15 -w 1000 $Target 2>&1

        [pscustomobject]@{
            Time   = Get-Date
            Target = $Target
            Output = @($output)
            Status = 'OK'
        }
    }
    catch {
        [pscustomobject]@{
            Time   = Get-Date
            Target = $Target
            Output = @("Erro ao executar tracert para $Target")
            Status = 'FALHA'
        }
    }
}

# ============================================================
# FILE: TraceOverview.ps1
# ============================================================

function Get-TraceRouteOverview {
    param(
        [object[]]$TraceResults = @()
    )

    $items = New-Object System.Collections.Generic.List[object]

    foreach ($tr in @($TraceResults)) {
        $hopLines = @($tr.Output) | Where-Object { $_ -match '^\s*\d+\s+' }
        $hopCount = $hopLines.Count
        $timeoutHops = New-Object System.Collections.Generic.List[int]
        $latencyByHop = New-Object System.Collections.Generic.List[object]

        foreach ($line in $hopLines) {
            $hopNumber = $null
            if ($line -match '^\s*(\d+)\s+') {
                $hopNumber = [int]$Matches[1]
            }

            if ($line -match '\*') {
                if ($hopNumber -ne $null) { $timeoutHops.Add($hopNumber) | Out-Null }
            }

            $matches = [regex]::Matches($line, '(?<v><1|\d+)\s*ms')
            $vals = New-Object System.Collections.Generic.List[int]
            foreach ($m in $matches) {
                $raw = $m.Groups['v'].Value
                if ($raw -eq '<1') { $vals.Add(1) | Out-Null }
                else { $vals.Add([int]$raw) | Out-Null }
            }

            if ($hopNumber -ne $null -and $vals.Count -gt 0) {
                $avg = [math]::Round(($vals | Measure-Object -Average).Average, 2)
                $latencyByHop.Add([pscustomobject]@{ Hop=$hopNumber; AvgMs=$avg }) | Out-Null
            }
        }

        $firstLatency = $latencyByHop | Select-Object -First 1
        $lastLatency = $latencyByHop | Select-Object -Last 1
        $firstDegradation = $null

        for ($i = 1; $i -lt $latencyByHop.Count; $i++) {
            $prev = $latencyByHop[$i - 1]
            $cur = $latencyByHop[$i]
            if (($cur.AvgMs -ge 80 -and ($cur.AvgMs - $prev.AvgMs) -ge 40) -or ($cur.AvgMs -ge 150)) {
                $firstDegradation = $cur.Hop
                break
            }
        }

        $lossText = if ($timeoutHops.Count -gt 0) { 'SIM (hop ' + (($timeoutHops.ToArray() | Select-Object -Unique) -join ', ') + ')' } else { 'NAO' }
        $degradationText = if ($firstDegradation) { "Hop $firstDegradation" } else { 'NAO DETECTADA' }
        $initialText = if ($firstLatency) { "$($firstLatency.AvgMs)ms" } else { 'N/A' }
        $finalText = if ($lastLatency) { "$($lastLatency.AvgMs)ms" } else { 'N/A' }

        $conclusion = 'Rota sem degradacao relevante aparente.'
        if ($timeoutHops.Count -gt 0 -and $firstDegradation) {
            $conclusion = 'Instabilidade detectada na rota: ha timeout intermediario e aumento de latencia.'
        }
        elseif ($timeoutHops.Count -gt 0) {
            $conclusion = 'Ha timeout em saltos intermediarios. Pode ser perda real ou roteador bloqueando resposta ICMP.'
        }
        elseif ($firstDegradation) {
            $conclusion = 'Aumento abrupto de latencia apos a rede local. Possivel rota externa, backbone, peering ou destino remoto.'
        }

        $items.Add([pscustomobject]@{
            Target            = $tr.Target
            HopCount          = $hopCount
            FirstDegradation  = $degradationText
            InitialLatency    = $initialText
            FinalLatency      = $finalText
            ApparentLoss      = $lossText
            Conclusion        = $conclusion
        }) | Out-Null
    }

    @($items.ToArray())
}

# ============================================================
# FILE: Score.ps1
# ============================================================

function Get-NetworkScore {
    param(
        [object[]]$PingResults = @(),
        [object[]]$DnsResults = @(),
        [object[]]$TcpResults = @(),
        [object[]]$Events = @()
    )

    $score = 100
    $details = New-Object System.Collections.Generic.List[string]

    $pingGroups = @($PingResults) | Where-Object { $_ } | Group-Object Layer, Target

    foreach ($group in $pingGroups) {
        $items = @($group.Group)
        if ($items.Count -eq 0) { continue }

        $first = $items | Select-Object -First 1
        $layer = $first.Layer
        $target = $first.Target

        if ($layer -eq 'InternetDNSName' -and @($items | Where-Object { $_.Status -eq 'SEM_ICMP' }).Count -gt 0) {
            $details.Add("InternetDNSName $target sem resposta ICMP: nao penalizado quando DNS/TCP validam o destino") | Out-Null
            continue
        }

        $total = [double]$items.Count
        $failCount = [double]@($items | Where-Object { $_.Status -eq 'FALHA' }).Count
        $critCount = [double]@($items | Where-Object { $_.Status -eq 'CRITICO' }).Count
        $alertCount = [double]@($items | Where-Object { $_.Status -eq 'ALERTA' }).Count
        $badCount = $failCount + $critCount + $alertCount
        $badRate = if ($total -gt 0) { $badCount / $total } else { 0 }
        $failRate = if ($total -gt 0) { $failCount / $total } else { 0 }
        $maxLoss = @($items | ForEach-Object { if ($_.LossPercent -ne $null) { [double]$_.LossPercent } else { 0 } } | Measure-Object -Maximum).Maximum
        $avgLatency = @($items | Where-Object { $_.AvgMs -ne $null } | ForEach-Object { [double]$_.AvgMs } | Measure-Object -Average).Average

        if ($failRate -ge 0.50) {
            $penalty = if ($layer -eq 'Gateway') { 30 } elseif ($layer -eq 'InternetIP') { 25 } else { 20 }
            $score -= $penalty
            $details.Add("$layer $target falhou de forma recorrente: $([math]::Round($failRate * 100,2))% das amostras") | Out-Null
        }
        elseif ($failRate -gt 0) {
            $penalty = if ($layer -eq 'Gateway') { 8 } elseif ($layer -eq 'InternetIP') { 6 } else { 5 }
            $score -= $penalty
            $details.Add("$layer $target teve falha momentanea: $([int]$failCount)/$([int]$total) amostras") | Out-Null
        }
        elseif ($badRate -ge 0.50) {
            $score -= 15
            $details.Add("$layer $target apresentou instabilidade recorrente: $([math]::Round($badRate * 100,2))% das amostras") | Out-Null
        }
        elseif ($badRate -gt 0) {
            $score -= 4
            $details.Add("$layer $target apresentou instabilidade pontual: $([int]$badCount)/$([int]$total) amostras") | Out-Null
        }

        if ($maxLoss -gt 0 -and $failRate -ge 0.50) {
            $score -= [math]::Min(20, [int]$maxLoss)
            $details.Add("Perda recorrente detectada em $layer ${target}: ate $maxLoss%") | Out-Null
        }
        elseif ($maxLoss -gt 0 -and $failRate -gt 0) {
            $score -= 3
            $details.Add("Perda pontual detectada em $layer ${target}: ate $maxLoss%") | Out-Null
        }

        if ($avgLatency -ne $null -and $avgLatency -gt 150) {
            $score -= 10
            $details.Add("Latencia elevada em $layer ${target}: media $([math]::Round($avgLatency,2))ms") | Out-Null
        }
    }

    foreach ($d in @($DnsResults)) {
        if (-not $d.Success) {
            $score -= 12
            $details.Add("DNS falhou para $($d.Name)") | Out-Null
        }
        elseif ($d.LatencyMs -ne $null -and [double]$d.LatencyMs -gt 1000) {
            $score -= 8
            $details.Add("DNS lento para $($d.Name): $($d.LatencyMs)ms") | Out-Null
        }
    }

    foreach ($t in @($TcpResults)) {
        if (-not $t.Success) {
            $score -= 5
            $details.Add("TCP falhou em $($t.ComputerName):$($t.Port) ($($t.Label))") | Out-Null
        }
    }

    foreach ($e in @($Events)) {
        if ($e.Type -eq 'RECUPERACAO') { continue }
        if ($e.Type -eq 'INSTABILIDADE MOMENTANEA/INTERMITENTE') {
            $score -= 3
            $details.Add("Evento ALERTA: $($e.Type)") | Out-Null
            continue
        }

        if ($e.Level -eq 'FALHA') {
            # Falhas ja sao ponderadas pelos resultados de ping. Evento fica registrado, mas com penalidade menor.
            $score -= 3
            $details.Add("Evento FALHA: $($e.Type)") | Out-Null
        }
        elseif ($e.Level -eq 'CRITICO') {
            $score -= 10
            $details.Add("Evento CRITICO: $($e.Type)") | Out-Null
        }
        elseif ($e.Level -eq 'ALERTA') {
            $score -= 4
            $details.Add("Evento ALERTA: $($e.Type)") | Out-Null
        }
    }

    if ($score -lt 0) { $score = 0 }
    if ($score -gt 100) { $score = 100 }

    $status = if ($score -ge 85) {
        'OK'
    }
    elseif ($score -ge 60) {
        'INSTABILIDADE PARCIAL'
    }
    elseif ($score -ge 35) {
        'CRITICO'
    }
    else {
        'FALHA'
    }

    [pscustomobject]@{
        Score  = [int]$score
        Status = $status
        Details = @($details.ToArray() | Select-Object -Unique)
    }
}

# ============================================================
# FILE: Diagnosis.ps1
# ============================================================

function Get-SmartDiagnosis {
    param(
        [object[]]$PingResults = @(),
        [object[]]$DnsResults = @(),
        [object[]]$Events = @(),
        [object]$NetworkInfo,
        [string]$Mode = 'Automatic',
        [object]$Score = $null
    )

    $lines = New-Object System.Collections.Generic.List[string]

    if (-not $NetworkInfo -or -not $NetworkInfo.Found) {
        return 'Nenhum adaptador de rede ativo com IPv4 foi identificado.'
    }

    $isPersistent = ($Mode -eq 'Persistent')
    $title = if ($isPersistent) { 'Monitoramento persistente:' } else { 'Analise inteligente:' }

    if ($Score) {
        $lines.Add('Status geral:') | Out-Null
        $lines.Add("$($Score.Status) - Saude da rede em $($Score.Score)/100.") | Out-Null
        $lines.Add('') | Out-Null
    }

    $lines.Add($title) | Out-Null

    if ($isPersistent) {
        $totalPing = @($PingResults).Count
        $badPing = @($PingResults | Where-Object { $_.Status -in @('ALERTA','CRITICO','FALHA') }).Count
        $recoveries = @($Events | Where-Object { $_.Type -eq 'RECUPERACAO' }).Count
        $lines.Add("Foram registradas $totalPing amostras de ping durante a janela de monitoramento.") | Out-Null
        $lines.Add("Foram observadas $badPing amostras com alerta, falha ou criticidade.") | Out-Null
        if ($recoveries -gt 0) {
            $lines.Add("Foram detectadas $recoveries recuperacoes automaticas apos falhas momentaneas.") | Out-Null
        }
    }
    else {
        $lines.Add('Foram correlacionados os resultados de ping, DNS, TCP, eventos e rota disponiveis.') | Out-Null
    }
    $lines.Add('') | Out-Null

    $possibleLoop = @($Events | Where-Object { $_.Type -eq 'POSSIVEL LOOP' })
    $momentary = @($Events | Where-Object { $_.Type -eq 'INSTABILIDADE MOMENTANEA/INTERMITENTE' })

    if ($possibleLoop.Count -gt 0) {
        $lines.Add('Instabilidade critica:') | Out-Null
        $lines.Add('Foi detectado padrao forte compativel com possivel loop ou instabilidade local severa recorrente.') | Out-Null
        $lines.Add('A classificacao foi gerada porque houve multiplas ocorrencias em janela curta, com gateway recorrente, perda generalizada e varios destinos afetados.') | Out-Null
        $lines.Add('') | Out-Null
    }
    elseif ($momentary.Count -gt 0) {
        $lines.Add('Instabilidade momentanea:') | Out-Null
        $lines.Add('Instabilidade momentanea/intermitente detectada.') | Out-Null
        $lines.Add('Possivel microqueda local ou perda transitoria de conectividade.') | Out-Null
        $lines.Add('Nao ha evidencia suficiente para classificar o comportamento como loop de rede.') | Out-Null
        $lines.Add('') | Out-Null
    }

    $gatewayItems = @($PingResults | Where-Object { $_.Layer -eq 'Gateway' })
    if ($gatewayItems.Count -gt 0) {
        $gwFail = @($gatewayItems | Where-Object { $_.Status -eq 'FALHA' }).Count
        $gwLast = $gatewayItems | Select-Object -Last 1
        $lines.Add('Rede local:') | Out-Null
        if ($gwFail -gt 0 -and $gwLast.Status -eq 'OK') {
            $lines.Add("Gateway $($gwLast.Target) apresentou falha momentanea, mas voltou a responder normalmente.") | Out-Null
        }
        elseif ($gwLast.Status -eq 'OK') {
            $lines.Add("Gateway $($gwLast.Target) respondeu normalmente com latencia de $($gwLast.AvgMs)ms e perda de $($gwLast.LossPercent)%.") | Out-Null
        }
        else {
            $lines.Add("Gateway $($gwLast.Target) apresenta status $($gwLast.Status). Verifique maquina, cabo, porta de switch, switch local, VLAN ou uplink.") | Out-Null
        }
        $lines.Add('') | Out-Null
    }

    $dnsPing = @($PingResults | Where-Object { $_.Layer -eq 'DNS' })
    $dnsResolveFail = @($DnsResults | Where-Object { -not $_.Success })
    $dnsSlow = @($DnsResults | Where-Object { $_.LatencyMs -ne $null -and [double]$_.LatencyMs -ge 1000 })
    $dnsBadPing = @($dnsPing | Where-Object { $_.Status -in @('ALERTA','CRITICO','FALHA') })

    $lines.Add('DNS:') | Out-Null
    if ($dnsResolveFail.Count -eq 0 -and $dnsSlow.Count -eq 0 -and $dnsBadPing.Count -eq 0) {
        $lines.Add('Os testes de DNS nao indicaram falha relevante no periodo analisado.') | Out-Null
    }
    else {
        foreach ($d in $dnsBadPing | Select-Object -First 3) {
            $lines.Add("DNS $($d.Target) apresentou status $($d.Status), media $($d.AvgMs)ms e perda $($d.LossPercent)%.") | Out-Null
        }
        foreach ($d in $dnsResolveFail | Select-Object -First 3) {
            $lines.Add("Falha de resolucao para $($d.Name) no servidor $($d.Server).") | Out-Null
        }
        foreach ($d in $dnsSlow | Select-Object -First 3) {
            $lines.Add("Resolucao DNS lenta para $($d.Name): $($d.LatencyMs)ms.") | Out-Null
        }
    }
    $lines.Add('') | Out-Null

    $internet = @($PingResults | Where-Object { $_.Layer -eq 'InternetIP' })
    if ($internet.Count -gt 0) {
        $internetFail = @($internet | Where-Object { $_.Status -eq 'FALHA' }).Count
        $internetLast = $internet | Select-Object -Last 1
        $lines.Add('Internet:') | Out-Null
        if ($internetFail -gt 0 -and $internetLast.Status -eq 'OK') {
            $lines.Add('Foi observada falha momentanea em pelo menos um destino externo, com recuperacao nos ciclos seguintes.') | Out-Null
        }
        elseif ($internetLast.Status -eq 'OK') {
            $lines.Add("Internet por IP respondeu com latencia de $($internetLast.AvgMs)ms e perda de $($internetLast.LossPercent)%.") | Out-Null
        }
        else {
            $lines.Add("Internet por IP apresentou status $($internetLast.Status). Possivel problema de rota, firewall, core ou link externo.") | Out-Null
        }
        $lines.Add('') | Out-Null
    }

    $eventList = @($Events | Where-Object { $_.Type -ne 'RECUPERACAO' })
    if ($eventList.Count -gt 0) {
        $lines.Add('Eventos:') | Out-Null
        foreach ($e in $eventList | Select-Object -First 6) {
            $lines.Add("$($e.Level): $($e.Type) em $($e.Time).") | Out-Null
        }
        $lines.Add('') | Out-Null
    }

    $lines.Add('Conclusao:') | Out-Null
    if ($possibleLoop.Count -gt 0) {
        $lines.Add('Existe evidencia de instabilidade local recorrente forte. Recomenda-se verificar loop fisico/logico, switch, VLAN, uplink, STP e broadcast storm.') | Out-Null
    }
    elseif ($momentary.Count -gt 0) {
        $lines.Add('O comportamento observado se parece mais com microqueda/intermitencia pontual do que com falha permanente ou loop de rede.') | Out-Null
    }
    elseif (@($PingResults | Where-Object { $_.Status -in @('ALERTA','CRITICO','FALHA') }).Count -gt 0) {
        $lines.Add('Foi detectada instabilidade parcial. Verifique os detalhes por camada no relatorio.') | Out-Null
    }
    else {
        $lines.Add('Nenhuma falha critica detectada no periodo analisado.') | Out-Null
    }

    @($lines.ToArray()) -join [Environment]::NewLine
}

# ============================================================
# FILE: Export.ps1
# ============================================================

function Add-NetDiagLine {
    param(
        [AllowEmptyCollection()][System.Collections.Generic.List[string]]$Target,
        [AllowNull()][AllowEmptyString()][object]$Value = ''
    )

    if ($null -eq $Value) {
        $Target.Add('') | Out-Null
    }
    else {
        $Target.Add([string]$Value) | Out-Null
    }
}

function Add-NetDiagSection {
    param(
        [AllowEmptyCollection()][System.Collections.Generic.List[string]]$Target,
        [AllowNull()][AllowEmptyString()][string]$Title = ''
    )

    Add-NetDiagLine -Target $Target -Value ''
    Add-NetDiagLine -Target $Target -Value '========================================='
    Add-NetDiagLine -Target $Target -Value "[$Title]"
    Add-NetDiagLine -Target $Target -Value '========================================='
}

function Add-NetDiagSubSection {
    param(
        [AllowEmptyCollection()][System.Collections.Generic.List[string]]$Target,
        [AllowNull()][AllowEmptyString()][string]$Title = ''
    )

    Add-NetDiagLine -Target $Target -Value ''
    Add-NetDiagLine -Target $Target -Value '*****************************************'
    Add-NetDiagLine -Target $Target -Value "[$Title]"
    Add-NetDiagLine -Target $Target -Value '*****************************************'
}

function Add-NetDiagItemSeparator {
    param([AllowEmptyCollection()][System.Collections.Generic.List[string]]$Target)
    Add-NetDiagLine -Target $Target -Value '_________________________________________'
}

function Export-NetDiagnosticReport {
    param(
        [Parameter(Mandatory)]
        [object]$Run,

        [switch]$ExportTxt,
        [switch]$ExportCsv,
        [switch]$ExportJson
    )

    if (-not (Test-Path $Script:ReportsDir)) {
        New-Item -Path $Script:ReportsDir -ItemType Directory -Force | Out-Null
    }

    $baseName = Get-TimeStampFileName -Prefix $Run.ReportPrefix
    $txtPath = Join-Path $Script:ReportsDir "$baseName.txt"
    $csvPath = Join-Path $Script:ReportsDir "$baseName.csv"
    $jsonPath = Join-Path $Script:ReportsDir "$baseName.json"

    $pingResults  = @($Run.PingResults  | ForEach-Object { $_ })
    $dnsResults   = @($Run.DnsResults   | ForEach-Object { $_ })
    $tcpResults   = @($Run.TcpResults   | ForEach-Object { $_ })
    $events       = @($Run.Events       | ForEach-Object { $_ })
    $traceResults = @($Run.TraceResults | ForEach-Object { $_ })

    if ($ExportTxt) {
        $lines = New-Object System.Collections.Generic.List[string]

        Add-NetDiagLine -Target $lines -Value '========================================'
        Add-NetDiagLine -Target $lines -Value "      $($Script:AppName) - Relatorio"
        Add-NetDiagLine -Target $lines -Value '========================================'
        Add-NetDiagLine -Target $lines -Value "Versao: $Script:Version"
        Add-NetDiagLine -Target $lines -Value "Inicio: $($Run.StartTime)"
        Add-NetDiagLine -Target $lines -Value "Fim: $($Run.EndTime)"
        Add-NetDiagLine -Target $lines -Value "Duracao: $($Run.Duration)"
        Add-NetDiagLine -Target $lines -Value "Tipo de execucao: $($Run.Mode)"
        Add-NetDiagLine -Target $lines -Value "Tipo de encerramento: $($Run.EndReason)"

        $net = $Run.NetworkInfo
        Add-NetDiagSection -Target $lines -Title 'REDE DETECTADA'
        Add-NetDiagLine -Target $lines -Value "Adaptador: $($net.InterfaceAlias)"
        Add-NetDiagLine -Target $lines -Value "InterfaceIndex: $($net.InterfaceIndex)"
        Add-NetDiagLine -Target $lines -Value "IPv4: $($net.IPv4Address)/$($net.PrefixLength)"
        Add-NetDiagLine -Target $lines -Value "Gateway: $($net.Gateway)"
        Add-NetDiagLine -Target $lines -Value "DNS: $(@($net.DnsServers) -join ', ')"
        Add-NetDiagLine -Target $lines -Value "MAC: $($net.MacAddress)"
        Add-NetDiagLine -Target $lines -Value "LinkSpeed: $($net.LinkSpeed)"

        Add-NetDiagSection -Target $lines -Title 'SCORE DA REDE'
        Add-NetDiagLine -Target $lines -Value "Saude: $($Run.Score.Score)/100"
        Add-NetDiagLine -Target $lines -Value "Status: $($Run.Score.Status)"

        if ($Run.Score.Details -and @($Run.Score.Details).Count -gt 0) {
            Add-NetDiagSubSection -Target $lines -Title 'DETALHES DO SCORE'
            foreach ($detail in @($Run.Score.Details)) {
                Add-NetDiagLine -Target $lines -Value "- $detail"
            }
        }

        $diagTitle = if ($Run.Mode -eq 'Persistent') { 'DIAGNOSTICO PERSISTENTE' } else { 'DIAGNOSTICO INTELIGENTE' }
        Add-NetDiagSection -Target $lines -Title $diagTitle
        foreach ($diagLine in ([string]$Run.Diagnosis -split "`r?`n")) {
            Add-NetDiagLine -Target $lines -Value $diagLine
        }

        Add-NetDiagSection -Target $lines -Title 'PING / TESTE POR CAMADAS'
        foreach ($p in $pingResults) {
            Add-NetDiagLine -Target $lines -Value "$($p.Time) | $($p.Layer) | $($p.Target) | Status=$($p.Status) | Avg=$($p.AvgMs)ms | Min=$($p.MinMs)ms | Max=$($p.MaxMs)ms | Perda=$($p.LossPercent)%"
        }

        Add-NetDiagSection -Target $lines -Title 'DNS TEST'
        if ($dnsResults.Count -eq 0) {
            Add-NetDiagLine -Target $lines -Value 'Nenhum teste DNS executado nesta etapa.'
        }
        foreach ($d in $dnsResults) {
            Add-NetDiagLine -Target $lines -Value "$($d.Time) | Name=$($d.Name) | Server=$($d.Server) | Success=$($d.Success) | Latency=$($d.LatencyMs)ms | Answer=$($d.Answer) | Status=$($d.Status)"
        }

        Add-NetDiagSection -Target $lines -Title 'TCP TEST'
        if ($tcpResults.Count -eq 0) {
            Add-NetDiagLine -Target $lines -Value 'Nenhum teste TCP executado nesta etapa.'
        }
        foreach ($t in $tcpResults) {
            Add-NetDiagLine -Target $lines -Value "$($t.Time) | $($t.ComputerName):$($t.Port) | $($t.Label) | Success=$($t.Success) | Status=$($t.Status)"
        }

        Add-NetDiagSection -Target $lines -Title 'EVENTOS'
        if ($events.Count -eq 0) {
            Add-NetDiagLine -Target $lines -Value 'Nenhum evento registrado.'
        }
        foreach ($e in $events) {
            Add-NetDiagItemSeparator -Target $lines
            Add-NetDiagLine -Target $lines -Value '[EVENTO]'
            Add-NetDiagLine -Target $lines -Value "Horario: $($e.Time)"
            Add-NetDiagLine -Target $lines -Value "Nivel: $($e.Level)"
            Add-NetDiagLine -Target $lines -Value "Tipo: $($e.Type)"
            Add-NetDiagLine -Target $lines -Value 'Sintomas:'
            foreach ($s in @($e.Symptoms)) {
                Add-NetDiagLine -Target $lines -Value "- $s"
            }
            if ($e.Details) {
                Add-NetDiagLine -Target $lines -Value "Detalhes: $($e.Details)"
            }
            Add-NetDiagLine -Target $lines -Value ''
        }

        Add-NetDiagSection -Target $lines -Title 'TRACERT'
        if ($traceResults.Count -eq 0) {
            Add-NetDiagLine -Target $lines -Value 'Nenhum tracert executado nesta etapa.'
        }
        foreach ($tr in $traceResults) {
            Add-NetDiagLine -Target $lines -Value "--- Tracert para $($tr.Target) ---"
            foreach ($line in @($tr.Output)) {
                Add-NetDiagLine -Target $lines -Value $line
            }
            Add-NetDiagLine -Target $lines -Value ''
        }

        $lines | Set-Content -Path $txtPath -Encoding UTF8
        Write-Status -Level 'OK' -Message "Relatorio TXT salvo em: $txtPath"
    }

    if ($ExportCsv) {
        $pingResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Status -Level 'OK' -Message "Relatorio CSV salvo em: $csvPath"
    }

    if ($ExportJson) {
        $Run | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonPath -Encoding UTF8
        Write-Status -Level 'OK' -Message "Relatorio JSON salvo em: $jsonPath"
    }

    [pscustomobject]@{
        Txt  = if ($ExportTxt) { $txtPath } else { $null }
        Csv  = if ($ExportCsv) { $csvPath } else { $null }
        Json = if ($ExportJson) { $jsonPath } else { $null }
    }
}

# ============================================================
# FILE: PersistentMonitor.ps1
# ============================================================

# Reservado para evolucao futura.
# A logica persistente principal esta implementada no Core.ps1 nesta versao preview.

# ============================================================
# FILE: Core.ps1
# ============================================================

function Invoke-NetDiagnostic {
    [CmdletBinding()]
    param(
        [switch]$Auto,

        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile = 'Laboratorio',

        [int]$Count = 20,

        [switch]$TraceRoute,
        [switch]$DnsTest,
        [switch]$TcpTest,
        [switch]$TestInternalRdp,

        [switch]$Persistent,
        [int]$IntervalSeconds = 10,
        [int]$DurationMinutes = 0,

        [switch]$ExportTxt,
        [switch]$ExportCsv,
        [switch]$ExportJson,

        [string[]]$Targets = @(),

        [string[]]$DnsNames = @('google.com','microsoft.com.br'),

        [string[]]$InternetIpTargets = @('8.8.8.8','1.1.1.1'),

        [string[]]$InternalTargets = @('10.250.250.2')
    )

    $run = [pscustomobject]@{
        ReportPrefix = if ($Persistent) { 'Persistent-NetDiagnostic' } else { 'NetDiagnostic' }
        Mode         = if ($Persistent) { 'Persistent' } else { 'Automatic' }
        StartTime    = Get-Date
        EndTime      = $null
        Duration     = $null
        EndReason    = 'Concluido'
        NetworkInfo  = $null
        PingResults  = New-Object System.Collections.Generic.List[object]
        DnsResults   = New-Object System.Collections.Generic.List[object]
        TcpResults   = New-Object System.Collections.Generic.List[object]
        TraceResults = New-Object System.Collections.Generic.List[object]
        Events       = New-Object System.Collections.Generic.List[object]
        Score        = $null
        Diagnosis    = ''
    }

    $Script:CurrentRun = $run

    switch ($Profile) {
        'Empresa' {
            if (-not $PSBoundParameters.ContainsKey('InternalTargets')) {
                $InternalTargets = @('10.250.250.2')
            }
            if (-not $PSBoundParameters.ContainsKey('DnsNames')) {
                $DnsNames = @('google.com','microsoft.com.br')
            }
        }
        'Laboratorio' {
            if (-not $PSBoundParameters.ContainsKey('InternalTargets')) {
                $InternalTargets = @()
            }
            if (-not $PSBoundParameters.ContainsKey('DnsNames')) {
                $DnsNames = @('google.com','microsoft.com.br')
            }
        }
    }

    try {
        $netInfo = Get-ActiveNetworkInfo
        $run.NetworkInfo = $netInfo

        if (-not $netInfo.Found) {
            $run.Events.Add((New-NetDiagEvent -Level 'FALHA' -Type 'SEM ADAPTADOR ATIVO' -Symptoms @('Nenhum adaptador ativo com IPv4 foi detectado') -Details 'Get-NetIPConfiguration nao encontrou interface valida')) | Out-Null
            Write-Status -Level 'FALHA' -Message 'Nenhum adaptador ativo com IPv4 foi detectado.'
            return
        }

        Write-Status -Level 'INFO' -Message "Adaptador: $($netInfo.InterfaceAlias) | IP: $($netInfo.IPv4Address) | Gateway: $($netInfo.Gateway)"

        $autoTargets = New-Object System.Collections.Generic.List[object]

        if ($netInfo.Gateway) {
            $autoTargets.Add([pscustomobject]@{ Layer='Gateway'; Target=$netInfo.Gateway }) | Out-Null
        }

        foreach ($dns in @($netInfo.DnsServers)) {
            if (-not [string]::IsNullOrWhiteSpace($dns)) {
                $autoTargets.Add([pscustomobject]@{ Layer='DNS'; Target=$dns }) | Out-Null
            }
        }

        foreach ($internalTarget in @($InternalTargets)) {
            if (-not [string]::IsNullOrWhiteSpace($internalTarget)) {
                $autoTargets.Add([pscustomobject]@{ Layer='ServidorInterno'; Target=$internalTarget }) | Out-Null
            }
        }

        foreach ($ip in @($InternetIpTargets)) {
            if (-not [string]::IsNullOrWhiteSpace($ip)) {
                $autoTargets.Add([pscustomobject]@{ Layer='InternetIP'; Target=$ip }) | Out-Null
            }
        }

        foreach ($target in @($Targets)) {
            if (-not [string]::IsNullOrWhiteSpace($target)) {
                $autoTargets.Add([pscustomobject]@{ Layer='Custom'; Target=$target }) | Out-Null
            }
        }

        if (-not $Persistent) {
            foreach ($name in @($DnsNames)) {
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $autoTargets.Add([pscustomobject]@{ Layer='InternetDNSName'; Target=$name }) | Out-Null
                }
            }
        }

        if ($Persistent) {
            $endAt = if ($DurationMinutes -gt 0) { (Get-Date).AddMinutes($DurationMinutes) } else { $null }

            Write-Status -Level 'INFO' -Message 'Diagnostico persistente iniciado. Pressione Ctrl+C para interromper e gerar relatorio.'
            Write-Status -Level 'INFO' -Message "Intervalo: $IntervalSeconds segundos | Duracao: $(if ($DurationMinutes -gt 0) { "$DurationMinutes minutos" } else { 'indefinida' })"

            $cycle = 0

            while ($true) {
                if ($endAt -and (Get-Date) -ge $endAt) {
                    $run.EndReason = 'Duracao finalizada'
                    break
                }

                $cycle++
                $samplePing = New-Object System.Collections.Generic.List[object]
                $sampleDns  = New-Object System.Collections.Generic.List[object]

                foreach ($entry in @($autoTargets.ToArray())) {
                    $result = Test-PingTarget -Target $entry.Target -Count 1 -TimeoutSeconds 2 -Layer $entry.Layer
                    $run.PingResults.Add($result) | Out-Null
                    $samplePing.Add($result) | Out-Null
                }

                if ($DnsTest) {
                    foreach ($name in @($DnsNames)) {
                        if (-not [string]::IsNullOrWhiteSpace($name)) {
                            $dnsResult = Test-DnsResolution -Name $name
                            $run.DnsResults.Add($dnsResult) | Out-Null
                            $sampleDns.Add($dnsResult) | Out-Null
                        }
                    }
                }

                $sampleEvents = Get-EventsFromSample -PingResults @($samplePing.ToArray()) -DnsResults @($sampleDns.ToArray())

                foreach ($ev in @($sampleEvents)) {
                    $run.Events.Add($ev) | Out-Null
                    Write-Status -Level $ev.Level -Message $ev.Type
                }

                $gw = @($samplePing.ToArray()) | Where-Object { $_.Layer -eq 'Gateway' } | Select-Object -First 1
                $internal = @($samplePing.ToArray()) | Where-Object { $_.Layer -eq 'ServidorInterno' } | Select-Object -First 1
                $internet = @($samplePing.ToArray()) | Where-Object { $_.Layer -eq 'InternetIP' } | Select-Object -First 1

                $elapsed = New-TimeSpan -Start $run.StartTime -End (Get-Date)
                $remainingText = 'Indefinido'

                if ($endAt) {
                    $remaining = New-TimeSpan -Start (Get-Date) -End $endAt
                    if ($remaining.TotalSeconds -lt 0) {
                        $remainingText = '00:00:00'
                    }
                    else {
                        $remainingText = $remaining.ToString('hh\:mm\:ss')
                    }
                }

                $gwText = if ($gw) { "Gateway=$($gw.AvgMs)ms/$($gw.Status)" } else { 'Gateway=N/A' }
                $internalText = if ($internal) { "Interno=$($internal.Target):$($internal.AvgMs)ms/$($internal.Status)" } else { 'Interno=N/A' }
                $ipText = if ($internet) { "InternetIP=$($internet.AvgMs)ms/$($internet.Status)" } else { 'InternetIP=N/A' }

                Write-Status -Level 'INFO' -Message ("[CICLO {0:D3}] Decorrido={1} | Restante={2} | Amostras={3}" -f $cycle, $elapsed.ToString('hh\:mm\:ss'), $remainingText, $run.PingResults.Count)
                Write-Status -Level 'INFO' -Message "$gwText | $internalText | $ipText"

                Start-Sleep -Seconds $IntervalSeconds
            }
        }
        else {
            Write-Status -Level 'INFO' -Message 'Executando diagnostico automatico completo...'

            foreach ($entry in @($autoTargets.ToArray())) {
                $result = Test-PingTarget -Target $entry.Target -Count $Count -TimeoutSeconds 2 -Layer $entry.Layer
                $run.PingResults.Add($result) | Out-Null
                Write-Status -Level $result.Status -Message "$($entry.Layer) $($entry.Target) | Avg=$($result.AvgMs)ms | Perda=$($result.LossPercent)%"
            }

            if ($DnsTest) {
                foreach ($name in @($DnsNames)) {
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        $dnsResult = Test-DnsResolution -Name $name
                        $run.DnsResults.Add($dnsResult) | Out-Null
                        Write-Status -Level $dnsResult.Status -Message "DNS $name | Sucesso=$($dnsResult.Success) | Latencia=$($dnsResult.LatencyMs)ms"
                    }
                }

                foreach ($dnsServer in @($netInfo.DnsServers)) {
                    foreach ($name in @($DnsNames | Select-Object -First 1)) {
                        if (-not [string]::IsNullOrWhiteSpace($dnsServer) -and -not [string]::IsNullOrWhiteSpace($name)) {
                            $dnsResult = Test-DnsResolution -Name $name -Server $dnsServer
                            $run.DnsResults.Add($dnsResult) | Out-Null
                            Write-Status -Level $dnsResult.Status -Message "DNS $name via $dnsServer | Sucesso=$($dnsResult.Success) | Latencia=$($dnsResult.LatencyMs)ms"
                        }
                    }
                }
            }

            if ($TcpTest) {
                $tcpTargets = @(
                    [pscustomobject]@{ ComputerName='8.8.8.8'; Port=53; Label='DNS Google' },
                    [pscustomobject]@{ ComputerName='1.1.1.1'; Port=53; Label='DNS Cloudflare' },
                    [pscustomobject]@{ ComputerName='google.com'; Port=443; Label='HTTPS externo' },
                    [pscustomobject]@{ ComputerName='microsoft.com'; Port=443; Label='HTTPS externo' }
                )

                foreach ($internalTarget in @($InternalTargets)) {
                    if (-not [string]::IsNullOrWhiteSpace($internalTarget)) {
                        $tcpTargets += [pscustomobject]@{ ComputerName=$internalTarget; Port=445; Label='SMB interno' }

                        if ($TestInternalRdp) {
                            $tcpTargets += [pscustomobject]@{ ComputerName=$internalTarget; Port=3389; Label='RDP interno' }
                        }
                    }
                }

                foreach ($tcp in @($tcpTargets)) {
                    $tcpResult = Test-TcpPort -ComputerName $tcp.ComputerName -Port $tcp.Port -Label $tcp.Label
                    $run.TcpResults.Add($tcpResult) | Out-Null
                    Write-Status -Level $tcpResult.Status -Message "TCP $($tcp.ComputerName):$($tcp.Port) | $($tcp.Label)"
                }
            }

            $events = Get-EventsFromSample -PingResults @($run.PingResults.ToArray()) -DnsResults @($run.DnsResults.ToArray())

            foreach ($ev in @($events)) {
                $run.Events.Add($ev) | Out-Null
            }

            if ($TraceRoute) {
                $traceTargets = @()

                if ($netInfo.Gateway) {
                    $traceTargets += $netInfo.Gateway
                }

                foreach ($internalTarget in @($InternalTargets)) {
                    if (-not [string]::IsNullOrWhiteSpace($internalTarget)) {
                        $traceTargets += $internalTarget
                    }
                }

                $traceTargets += '8.8.8.8'
                $traceTargets += 'google.com'

                foreach ($trg in @($traceTargets | Select-Object -Unique)) {
                    Write-Status -Level 'INFO' -Message "Executando tracert para $trg..."
                    $trace = Invoke-TraceRouteSimple -Target $trg
                    $run.TraceResults.Add($trace) | Out-Null
                }
            }
        }
    }
    catch {
        $run.EndReason = "Erro interno: $($_.Exception.Message)"
        Write-Status -Level 'FALHA' -Message "Erro interno no diagnostico: $($_.Exception.Message)"
    }
    finally {
        $run.EndTime = Get-Date
        $run.Duration = New-TimeSpan -Start $run.StartTime -End $run.EndTime

        if ($Persistent -and $run.EndReason -eq 'Concluido') {
            $run.EndReason = 'Interrompido pelo usuario ou finalizado'
        }

        if (Get-Command Get-PersistentTimelineEvents -ErrorAction SilentlyContinue) {
            $timelineEvents = Get-PersistentTimelineEvents -PingResults @($run.PingResults.ToArray()) -ExistingEvents @($run.Events.ToArray())
            foreach ($ev in @($timelineEvents)) {
                $alreadyExists = @($run.Events.ToArray()) | Where-Object {
                    $_.Type -eq $ev.Type -and $_.Details -eq $ev.Details -and $_.Time -eq $ev.Time
                } | Select-Object -First 1

                if (-not $alreadyExists) {
                    $run.Events.Add($ev) | Out-Null
                }
            }
        }

        $run.Score = Get-NetworkScore -PingResults @($run.PingResults.ToArray()) -DnsResults @($run.DnsResults.ToArray()) -TcpResults @($run.TcpResults.ToArray()) -Events @($run.Events.ToArray())

        $run.Diagnosis = Get-SmartDiagnosis -PingResults @($run.PingResults.ToArray()) -DnsResults @($run.DnsResults.ToArray()) -Events @($run.Events.ToArray()) -NetworkInfo $run.NetworkInfo -Mode $run.Mode -Score $run.Score

        $Script:LastResult = $run

        Write-Section 'Resumo do diagnostico'
        Write-Host "Score: $($run.Score.Score)/100" -ForegroundColor Cyan
        Write-Host "Status: $($run.Score.Status)" -ForegroundColor Cyan
        Write-Host "Diagnostico: $($run.Diagnosis)" -ForegroundColor Yellow
        Write-Host "Eventos: $($run.Events.Count)" -ForegroundColor Cyan
        Write-Host "Duracao: $($run.Duration)" -ForegroundColor Cyan
        Write-Host "Encerramento: $($run.EndReason)" -ForegroundColor Cyan

        if ($ExportTxt -or $ExportCsv -or $ExportJson) {
            Export-NetDiagnosticReport -Run $run -ExportTxt:$ExportTxt -ExportCsv:$ExportCsv -ExportJson:$ExportJson | Out-Null
        }
        else {
            Export-NetDiagnosticReport -Run $run -ExportTxt | Out-Null
        }
    }
}

# ============================================================
# FILE: Main.ps1
# ============================================================

function Show-MainMenu {
    Clear-Host
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host '      Invoke-NetDiagnostic' -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '[1] Perfil Empresa'
    Write-Host '[2] Perfil Laboratorio/Casa'
    Write-Host '[0] Sair'
    Write-Host ''
}

function Show-ProfileMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    Clear-Host
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host "      Invoke-NetDiagnostic - Perfil $Profile" -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '[1] Diagnostico automatico completo'
    Write-Host '[2] Diagnostico Persistente'
    Write-Host '[3] Opcoes Individuais'
    Write-Host '[0] Voltar'
    Write-Host ''
}

function Show-PersistentMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    Clear-Host
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host "      Diagnostico Persistente - Perfil $Profile" -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '[1] Monitoramento rapido - 5 minutos'
    Write-Host '[2] Monitoramento medio - 30 minutos'
    Write-Host '[3] Monitoramento longo - 2 horas'
    Write-Host '[4] Monitoramento personalizado'
    Write-Host '[0] Voltar'
    Write-Host ''
}

function Show-IndividualMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    Clear-Host
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host "      Opcoes Individuais - Perfil $Profile" -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '[1] AutoDetect - Ver informacoes da rede atual'
    Write-Host '[2] Teste por camadas'
    Write-Host '[3] Teste DNS'
    Write-Host '[4] Teste TCP portas basicas'
    Write-Host '[5] Tracert'
    Write-Host '[6] Exportar ultimo resultado TXT/CSV/JSON'
    Write-Host '[0] Voltar'
    Write-Host ''
}

function Start-AutoDiagnosticMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    Clear-Host
    Write-Section "Diagnostico automatico completo - Perfil $Profile"
    Invoke-NetDiagnostic -Profile $Profile -Auto -Count 20 -TraceRoute -DnsTest -TcpTest -ExportTxt -ExportCsv -ExportJson
    Pause
}

function Start-PersistentDiagnosticMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    while ($true) {
        Show-PersistentMenu -Profile $Profile
        $choice = Read-Host 'Escolha uma opcao'

        switch ($choice) {
            '1' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Persistent -IntervalSeconds 10 -DurationMinutes 5 -DnsTest -ExportTxt -ExportCsv -ExportJson
                Pause
            }
            '2' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Persistent -IntervalSeconds 10 -DurationMinutes 30 -DnsTest -ExportTxt -ExportCsv -ExportJson
                Pause
            }
            '3' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Persistent -IntervalSeconds 15 -DurationMinutes 120 -DnsTest -ExportTxt -ExportCsv -ExportJson
                Pause
            }
            '4' {
                $interval = Read-Host 'Intervalo em segundos'
                $duration = Read-Host 'Duracao em minutos (0 para indefinido)'

                if (-not ($interval -as [int])) { $interval = 10 }
                if (-not ($duration -as [int])) { $duration = 0 }

                Invoke-NetDiagnostic -Profile $Profile -Auto -Persistent -IntervalSeconds ([int]$interval) -DurationMinutes ([int]$duration) -DnsTest -ExportTxt -ExportCsv -ExportJson
                Pause
            }
            '0' { return }
            default {
                Write-Status -Level 'ALERTA' -Message 'Opcao invalida.'
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Start-IndividualMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    while ($true) {
        Show-IndividualMenu -Profile $Profile
        $choice = Read-Host 'Escolha uma opcao'

        switch ($choice) {
            '1' {
                Write-Section "AutoDetect - Perfil $Profile"
                Get-ActiveNetworkInfo | Format-List
                Pause
            }
            '2' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Count 10 -ExportTxt
                Pause
            }
            '3' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Count 4 -DnsTest -ExportTxt
                Pause
            }
            '4' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Count 4 -TcpTest -ExportTxt
                Pause
            }
            '5' {
                Invoke-NetDiagnostic -Profile $Profile -Auto -Count 4 -TraceRoute -ExportTxt
                Pause
            }
            '6' {
                if ($Script:LastResult) {
                    Export-NetDiagnosticReport -Run $Script:LastResult -ExportTxt -ExportCsv -ExportJson | Out-Null
                }
                else {
                    Write-Status -Level 'ALERTA' -Message 'Nenhum resultado anterior encontrado.'
                }
                Pause
            }
            '0' { return }
            default {
                Write-Status -Level 'ALERTA' -Message 'Opcao invalida.'
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Start-ProfileMenu {
    param(
        [ValidateSet('Empresa','Laboratorio')]
        [string]$Profile
    )

    while ($true) {
        Show-ProfileMenu -Profile $Profile
        $choice = Read-Host 'Escolha uma opcao'

        switch ($choice) {
            '1' { Start-AutoDiagnosticMenu -Profile $Profile }
            '2' { Start-PersistentDiagnosticMenu -Profile $Profile }
            '3' { Start-IndividualMenu -Profile $Profile }
            '0' { return }
            default {
                Write-Status -Level 'ALERTA' -Message 'Opcao invalida.'
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Start-NetDiagnosticMenu {
    while ($true) {
        Show-MainMenu
        $choice = Read-Host 'Escolha uma opcao'

        switch ($choice) {
            '1' { Start-ProfileMenu -Profile 'Empresa' }
            '2' { Start-ProfileMenu -Profile 'Laboratorio' }
            '0' { return }
            default {
                Write-Status -Level 'ALERTA' -Message 'Opcao invalida.'
                Start-Sleep -Seconds 1
            }
        }
    }
}

Start-NetDiagnosticMenu
