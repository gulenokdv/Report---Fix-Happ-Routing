param(
    [string]$ConfigPath = "$env:LOCALAPPDATA\Happ\config.json",
    [string]$RuleSetDirectory = (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "rulesets")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-RuleSet {
    param(
        [string]$Tag,
        [string]$Path
    )

    [pscustomobject]@{
        tag    = $Tag
        type   = "local"
        format = "binary"
        path   = $Path
    }
}

function New-RouteRule {
    param(
        [object]$RuleSet,
        [string]$Outbound
    )

    [pscustomobject]@{
        rule_set = $RuleSet
        outbound = $Outbound
    }
}

function Set-JsonProperty {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        $Value
    )

    if (Test-JsonProperty -Object $Object -Name $Name) {
        $Object.$Name = $Value
    }
    else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Test-JsonProperty {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Object) {
        return $false
    }

    $property = $Object.PSObject.Properties[$Name]
    return $null -ne $property
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Config not found: $ConfigPath"
}

$resolvedRuleSetDirectory = [System.IO.Path]::GetFullPath($RuleSetDirectory)
$requiredRuleSets = @(
    @{ Tag = "geosite-vk"; Path = (Join-Path -Path $resolvedRuleSetDirectory -ChildPath "geosite-vk.srs") },
    @{ Tag = "geosite-category-ru"; Path = (Join-Path -Path $resolvedRuleSetDirectory -ChildPath "geosite-category-ru.srs") },
    @{ Tag = "geoip-ru"; Path = (Join-Path -Path $resolvedRuleSetDirectory -ChildPath "geoip-ru.srs") }
)

foreach ($ruleSet in $requiredRuleSets) {
    if (-not (Test-Path -LiteralPath $ruleSet.Path)) {
        throw "Rule-set file not found: $($ruleSet.Path)"
    }
}

$raw = Get-Content -LiteralPath $ConfigPath -Raw
$config = $raw | ConvertFrom-Json

if (-not (Test-JsonProperty -Object $config -Name "route")) {
    $config | Add-Member -NotePropertyName route -NotePropertyValue ([pscustomobject]@{})
}
elseif ($null -eq $config.route) {
    $config.route = [pscustomobject]@{}
}

$processRule = $null
$sniffRule = $null
$dnsHijackRule = $null

if ($config.route.rules) {
    foreach ($rule in @($config.route.rules)) {
        if (-not $processRule -and (Test-JsonProperty -Object $rule -Name "process_name")) {
            $processRule = $rule
            continue
        }
        if (-not $sniffRule -and $rule.action -eq "sniff") {
            $sniffRule = $rule
            continue
        }
        if (-not $dnsHijackRule -and $rule.action -eq "hijack-dns") {
            $dnsHijackRule = $rule
            continue
        }
    }
}

if (-not $processRule) {
    $processRule = [pscustomobject]@{
        outbound     = "direct"
        process_name = @("xray.exe", "sing-box.exe", "antifilter.exe", "xray", "sing-box", "antifilter")
    }
}

if (-not $sniffRule) {
    $sniffRule = [pscustomobject]@{
        action = "sniff"
    }
}

if (-not $dnsHijackRule) {
    $dnsHijackRule = [pscustomobject]@{
        action   = "hijack-dns"
        protocol = "dns"
    }
}

Set-JsonProperty -Object $config.route -Name "auto_detect_interface" -Value $true
Set-JsonProperty -Object $config.route -Name "final" -Value "proxy"
Set-JsonProperty -Object $config.route -Name "rule_set" -Value @(
    (New-RuleSet -Tag "geosite-vk" -Path $requiredRuleSets[0].Path),
    (New-RuleSet -Tag "geosite-category-ru" -Path $requiredRuleSets[1].Path),
    (New-RuleSet -Tag "geoip-ru" -Path $requiredRuleSets[2].Path)
)
Set-JsonProperty -Object $config.route -Name "rules" -Value @(
    $processRule,
    $sniffRule,
    $dnsHijackRule,
    (New-RouteRule -RuleSet @("geosite-vk", "geosite-category-ru") -Outbound "direct"),
    (New-RouteRule -RuleSet @("geoip-ru") -Outbound "direct")
)

if (-not (Test-JsonProperty -Object $config -Name "experimental")) {
    $config | Add-Member -NotePropertyName experimental -NotePropertyValue ([pscustomobject]@{})
}
elseif ($null -eq $config.experimental) {
    $config.experimental = [pscustomobject]@{}
}
if (-not (Test-JsonProperty -Object $config.experimental -Name "cache_file")) {
    $config.experimental | Add-Member -NotePropertyName cache_file -NotePropertyValue ([pscustomobject]@{})
}
elseif ($null -eq $config.experimental.cache_file) {
    $config.experimental.cache_file = [pscustomobject]@{}
}
Set-JsonProperty -Object $config.experimental.cache_file -Name "enabled" -Value $true

$backupPath = "$ConfigPath.codex-backup"
if (-not (Test-Path -LiteralPath $backupPath)) {
    Copy-Item -LiteralPath $ConfigPath -Destination $backupPath
}

$json = $config | ConvertTo-Json -Depth 100
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($ConfigPath, $json, $utf8NoBom)
Write-Host "Patched $ConfigPath"
