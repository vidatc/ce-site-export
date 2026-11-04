#Requires -Version 5.1
<#
.SYNOPSIS
    Per-site count of files with SIT / trainable classifier hits, grouped by document type,
    from Microsoft Purview Content Explorer (Export-ContentExplorerData).

.DESCRIPTION
    Stage 1 : Site Name, Site URL, Tag Type, SIT / Trainable Classifier, Doc Type, Count of Files
    Stage 2 : Count of Hits and Highest Confidence - SITs only. Comes from the
              SensitiveInfoTypesData field in the same records, so it costs no extra calls.

    Flow
      1. Probe   - one call per tag at workload level (PageSize 1) to read TotalCount. Tags with
                   zero hits tenant-wide are skipped for every site. Cached in probe_<Workload>.csv.
      2. Extract - for each site x each tag with hits: item-level export filtered by -SiteUrl.
                   Each finished (site, tag) pair is cached under .\cache, so re-running the same
                   command resumes where it stopped (token expiry, RDP drop, etc.).
      3. Report  - rebuilds SiteSummary.csv and SiteDetail.csv from the cache for the sites given.

    Notes
      - Run in Windows PowerShell 5.1 (powershell.exe), not PS7, for Connect-IPPSSession.
      - Needs the Content Explorer List Viewer role group (item-level export).
      - Trainable classifiers have no hit count or confidence in Content Explorer: file count only.
      - Bundled SITs (e.g. All Credential Types, All Full Names): file counts are reliable; if the
        bundle's own GUID isn't in a file's SensitiveInfoTypesData, hits/confidence are left blank
        and the Note column says so. The member SITs' own rows carry their hits.

.PARAMETER SitesCsv
    CSV with columns SiteName, SiteUrl. Content Explorer returns URLs only, so SiteName comes from here.

.PARAMETER SiteUrl
    Alternative to -SitesCsv for a quick single-site run. SiteName = last URL segment.

.PARAMETER TagsCsv
    Optional CSV with columns TagType, TagName.
      TagType = SensitiveInformationType | TrainableClassifier | Sensitivity | Retention
    - No SensitiveInformationType rows  -> every SIT from Get-DlpSensitiveInformationType is swept.
    - Trainable classifiers are swept only if listed here. Copy the names from
      Content Explorer > Trainable classifiers (that list only shows TCs with hits).

.PARAMETER GroupByExtension
    Group by raw file extension instead of the Doc Type family (Word / Excel / PDF ...).

.PARAMETER RefreshProbe
    Ignore the cached probe results and re-probe every tag.

.PARAMETER SkipProbe
    Don't probe; query every tag against every site (slow - only for small tag lists).

.PARAMETER Force
    Ignore cached (site, tag) results and re-extract.

.EXAMPLE
    # In Windows PowerShell 5.1
    Connect-IPPSSession -UserPrincipalName vida@contoso.com
    .\Export-CESiteSummary.ps1 -SitesCsv .\Sites.csv -TagsCsv .\Tags.csv -UserPrincipalName vida@contoso.com

.EXAMPLE
    # One site, all SITs, no TCs
    .\Export-CESiteSummary.ps1 -SiteUrl https://contoso.sharepoint.com/sites/it-team
#>
[CmdletBinding(DefaultParameterSetName = 'Csv')]
param(
    [Parameter(ParameterSetName = 'Csv', Mandatory = $true)]
    [string]$SitesCsv,

    [Parameter(ParameterSetName = 'Url', Mandatory = $true)]
    [string[]]$SiteUrl,

    [string]$TagsCsv,

    [string]$OutDir = (Join-Path $PSScriptRoot 'CE_SiteExport'),

    [string]$UserPrincipalName,

    [ValidateRange(1, 10000)]
    [int]$PageSize = 5000,

    [int]$ReconnectMinutes = 45,

    [switch]$GroupByExtension,
    [switch]$RefreshProbe,
    [switch]$SkipProbe,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ValidTagTypes = @('SensitiveInformationType', 'TrainableClassifier', 'Sensitivity', 'Retention')
$GuidPattern   = '[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}'

# ---------------------------------------------------------------------------------------------
# Setup and logging
# ---------------------------------------------------------------------------------------------
$CacheDir = Join-Path $OutDir 'cache'
foreach ($d in @($OutDir, $CacheDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
}
$LogFile = Join-Path $OutDir 'run.log'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0}  {1,-5} {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Log "Running in PowerShell $($PSVersionTable.PSVersion). Connect-IPPSSession can break in PS7 (WAM broker) - use Windows PowerShell 5.1 if the connection fails." 'WARN'
}

# ---------------------------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------------------------
function ConvertTo-Bool($Value) {
    # Remoting can hand back "False" as a string, and [bool]"False" is $true.
    if ($Value -is [bool]) { return $Value }
    return ([string]$Value -eq 'True')
}

function Get-Prop($Object, [string[]]$Names) {
    foreach ($n in $Names) {
        $p = $Object.PSObject.Properties[$n]
        if ($p -and $null -ne $p.Value -and [string]$p.Value -ne '') { return $p.Value }
    }
    return $null
}

function Get-SafeName([string]$Text, [int]$Max = 60) {
    $s = [regex]::Replace($Text, '[^A-Za-z0-9._-]', '_')
    if ($s.Length -gt $Max) { $s = $s.Substring(0, $Max) }
    return $s
}

function Get-ShortHash([string]$Text) {
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $bytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant()))
    return (($bytes[0..3] | ForEach-Object { $_.ToString('x2') }) -join '')
}

$DocTypeMap = @{
    'doc' = 'Word'; 'docx' = 'Word'; 'docm' = 'Word'; 'dot' = 'Word'; 'dotx' = 'Word'; 'dotm' = 'Word'
    'xls' = 'Excel'; 'xlsx' = 'Excel'; 'xlsm' = 'Excel'; 'xlsb' = 'Excel'; 'xlt' = 'Excel'; 'xltx' = 'Excel'; 'xltm' = 'Excel'
    'ppt' = 'PowerPoint'; 'pptx' = 'PowerPoint'; 'pptm' = 'PowerPoint'; 'pps' = 'PowerPoint'; 'ppsx' = 'PowerPoint'; 'pot' = 'PowerPoint'; 'potx' = 'PowerPoint'
    'pdf' = 'PDF'
    'csv' = 'CSV'
    'msg' = 'Email'; 'eml' = 'Email'
    'txt' = 'Text'
    'one' = 'OneNote'
    'vsd' = 'Visio'; 'vsdx' = 'Visio'
    'jpg' = 'Image'; 'jpeg' = 'Image'; 'png' = 'Image'; 'gif' = 'Image'; 'bmp' = 'Image'; 'tif' = 'Image'; 'tiff' = 'Image'
    'zip' = 'Archive'; '7z' = 'Archive'; 'rar' = 'Archive'
}

function Get-Extension([string]$FileName) {
    if ($FileName -match '\.([A-Za-z0-9]{1,10})$') { return $Matches[1].ToLowerInvariant() }
    return ''
}

function Get-DocType([string]$Extension) {
    if (-not $Extension) { return '(no extension)' }
    if ($GroupByExtension) { return $Extension.ToUpperInvariant() }
    if ($DocTypeMap.ContainsKey($Extension)) { return $DocTypeMap[$Extension] }
    return $Extension.ToUpperInvariant()
}

$LevelRank = @{ '' = 0; 'Low' = 1; 'Medium' = 2; 'High' = 3 }
$TagTypeLabel = @{ 'SensitiveInformationType' = 'SIT'; 'TrainableClassifier' = 'TC'; 'Sensitivity' = 'Sensitivity label'; 'Retention' = 'Retention label' }

function Get-SitConfidence($Data, [string]$TargetGuid) {
    # Reads SensitiveInfoTypesData: [{"Id":"<guid>","LowConfidenceMatch":n,"MediumConfidenceMatch":n,"HighConfidenceMatch":n}, ...]
    $res = [ordered]@{ High = ''; Medium = ''; Low = ''; Level = ''; Note = '' }
    if ($null -eq $Data -or [string]$Data -eq '') { $res.Note = 'No confidence data on record'; return [pscustomobject]$res }
    if (-not $TargetGuid) { $res.Note = 'SIT GUID not resolved'; return [pscustomobject]$res }

    $entries = $Data
    if ($Data -is [string]) {
        try {
            # Assign before enumerating: PS 5.1 emits a JSON array as one object when piped.
            $entries = ConvertFrom-Json -InputObject $Data
        } catch {
            $res.Note = 'Confidence data not parsable'; return [pscustomobject]$res
        }
    }

    foreach ($e in @($entries)) {
        if ($null -eq $e) { continue }
        if ([string](Get-Prop $e @('Id', 'SensitiveInfoTypeId', 'Guid')) -ne $TargetGuid) { continue }
        $h = [int](Get-Prop $e @('HighConfidenceMatch', 'HighConfidenceCount', 'High'))
        $m = [int](Get-Prop $e @('MediumConfidenceMatch', 'MediumConfidenceCount', 'Medium'))
        $l = [int](Get-Prop $e @('LowConfidenceMatch', 'LowConfidenceCount', 'Low'))
        $res.High = $h; $res.Medium = $m; $res.Low = $l
        if ($h -gt 0) { $res.Level = 'High' } elseif ($m -gt 0) { $res.Level = 'Medium' } elseif ($l -gt 0) { $res.Level = 'Low' }
        return [pscustomobject]$res
    }
    $res.Note = 'SIT GUID not in item data - likely a bundle SIT'
    return [pscustomobject]$res
}

# ---------------------------------------------------------------------------------------------
# Connection handling
# ---------------------------------------------------------------------------------------------
$script:ConnectedAt = Get-Date

function Connect-CE {
    $p = @{}
    if ($UserPrincipalName) { $p.UserPrincipalName = $UserPrincipalName }
    Write-Log 'Connecting to Security & Compliance PowerShell...'
    Connect-IPPSSession @p | Out-Null
    $script:ConnectedAt = Get-Date
}

if (-not (Get-Command Export-ContentExplorerData -ErrorAction SilentlyContinue)) {
    if ($UserPrincipalName) { Connect-CE }
    else { throw 'Export-ContentExplorerData not available. Run Connect-IPPSSession first, or pass -UserPrincipalName.' }
}
if (-not $UserPrincipalName) {
    Write-Log "No -UserPrincipalName given: the ~1h token refresh (every $ReconnectMinutes min) will prompt for sign-in." 'WARN'
}

function Invoke-CE([hashtable]$Params, [int]$MaxAttempts = 3) {
    $attempt = 0
    while ($true) {
        $attempt++
        if (((Get-Date) - $script:ConnectedAt).TotalMinutes -ge $ReconnectMinutes) { Connect-CE }
        try {
            $r = Export-ContentExplorerData @Params -WarningAction SilentlyContinue -ErrorAction Stop
            if ($null -eq $r -or @($r).Count -eq 0) {
                throw 'Export-ContentExplorerData returned nothing (server-side error not raised by the cmdlet).'
            }
            return , @($r)
        } catch {
            $msg = $_.Exception.Message
            if ($attempt -ge $MaxAttempts) { throw }
            # Only a lost/expired session needs Connect-CE, which can open a sign-in prompt that waits
            # forever on an unattended run. Transient server errors just get a pause and a retry.
            if ($msg -match 'role definition|token|unauthori[sz]ed|\b40[13]\b|session') {
                Write-Log "    attempt $attempt failed: $msg - reconnecting and retrying" 'WARN'
                Start-Sleep -Seconds (10 * $attempt)
                try { Connect-CE } catch { Write-Log "    reconnect failed: $($_.Exception.Message)" 'WARN' }
            } else {
                Write-Log "    attempt $attempt failed: $msg - retrying in $(30 * $attempt)s without reconnecting" 'WARN'
                Start-Sleep -Seconds (30 * $attempt)
            }
        }
    }
}

# ---------------------------------------------------------------------------------------------
# Sites
# ---------------------------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    $rawSites = @(Import-Csv -Path $SitesCsv)
    foreach ($col in @('SiteName', 'SiteUrl')) {
        if ($rawSites.Count -gt 0 -and -not $rawSites[0].PSObject.Properties[$col]) { throw "SitesCsv is missing the '$col' column." }
    }
    $Sites = foreach ($s in $rawSites) {
        [pscustomobject]@{ SiteName = ([string]$s.SiteName).Trim(); SiteUrl = ([string]$s.SiteUrl).Trim().TrimEnd('/') }
    }
} else {
    $Sites = foreach ($u in $SiteUrl) {
        $clean = $u.Trim().TrimEnd('/')
        [pscustomobject]@{ SiteName = ($clean -split '/')[-1]; SiteUrl = $clean }
    }
}
$Sites = @($Sites | Where-Object { $_.SiteUrl })
foreach ($s in $Sites) {
    if ($s.SiteUrl -notmatch '^https://[^\s<>]+\.sharepoint\.com/') { throw "Not a usable SharePoint URL: '$($s.SiteUrl)'" }
    if (-not $s.SiteName) { $s.SiteName = ($s.SiteUrl -split '/')[-1] }
    $wl = 'SPO'
    if ($s.SiteUrl -match '-my\.sharepoint\.com') { $wl = 'ODB' }
    $s | Add-Member -NotePropertyName Workload -NotePropertyValue $wl
    $s | Add-Member -NotePropertyName CacheKey -NotePropertyValue ('{0}_{1}' -f (Get-SafeName ($s.SiteUrl -replace '^https://', '') 40), (Get-ShortHash $s.SiteUrl))
}
if ($Sites.Count -eq 0) { throw 'No sites to process.' }
Write-Log ("Sites: {0}" -f $Sites.Count)

# ---------------------------------------------------------------------------------------------
# Tags and GUID maps
# ---------------------------------------------------------------------------------------------
$SitGuidByName = @{}
$sitObjects = @()
try {
    $sitObjects = @(Get-DlpSensitiveInformationType)
    foreach ($sit in $sitObjects) {
        foreach ($prop in @('Id', 'Guid', 'Identity')) {
            $v = $sit.PSObject.Properties[$prop]
            if ($v -and [string]$v.Value -match $GuidPattern) { $SitGuidByName[[string]$sit.Name] = $Matches[0].ToLowerInvariant(); break }
        }
    }
    Write-Log ("SIT definitions loaded: {0}" -f $sitObjects.Count)
} catch {
    Write-Log "Get-DlpSensitiveInformationType failed ($($_.Exception.Message)) - hit counts/confidence will be blank." 'WARN'
}

$LabelNameByGuid = @{}
try {
    foreach ($lbl in @(Get-Label)) {
        $name = [string](Get-Prop $lbl @('DisplayName', 'Name'))
        foreach ($prop in @('Guid', 'ImmutableId', 'Identity')) {
            $v = $lbl.PSObject.Properties[$prop]
            if ($v -and [string]$v.Value -match $GuidPattern) { $LabelNameByGuid[$Matches[0].ToLowerInvariant()] = $name }
        }
    }
} catch {
    Write-Log "Get-Label failed ($($_.Exception.Message)) - Sensitivity Label column will show GUIDs." 'WARN'
}

$Tags = New-Object System.Collections.Generic.List[object]
$seen = @{}
function Add-Tag([string]$Type, [string]$Name) {
    $Name = $Name.Trim()
    if (-not $Name) { return }
    if ($Name -match '^<.*>$') { Write-Log "TagsCsv: placeholder '$Name' skipped - replace it with a real name." 'WARN'; return }
    $key = "$Type|$Name".ToLowerInvariant()
    if ($seen.ContainsKey($key)) { return }
    $seen[$key] = $true
    $Tags.Add([pscustomobject]@{ TagType = $Type; TagName = $Name })
}

if ($TagsCsv) {
    foreach ($row in @(Import-Csv -Path $TagsCsv)) {
        $t = ([string]$row.TagType).Trim()
        $match = $ValidTagTypes | Where-Object { $_ -eq $t }
        if (-not $match) { Write-Log "TagsCsv: unknown TagType '$t' for '$($row.TagName)' - skipped." 'WARN'; continue }
        Add-Tag $match ([string]$row.TagName)
    }
}
if (-not ($Tags | Where-Object { $_.TagType -eq 'SensitiveInformationType' })) {
    foreach ($sit in $sitObjects) { Add-Tag 'SensitiveInformationType' ([string]$sit.Name) }
}
if (-not ($Tags | Where-Object { $_.TagType -eq 'TrainableClassifier' })) {
    Write-Log 'No TrainableClassifier rows in -TagsCsv: trainable classifiers will not be swept.' 'WARN'
}
foreach ($t in $Tags | Where-Object { $_.TagType -eq 'SensitiveInformationType' }) {
    if (-not $SitGuidByName.ContainsKey($t.TagName)) { Write-Log "SIT '$($t.TagName)' not found in tenant SIT list - check spelling (it will error in the probe)." 'WARN' }
}
Write-Log ("Tags to consider: {0} ({1})" -f $Tags.Count, (($Tags | Group-Object TagType | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '))

# ---------------------------------------------------------------------------------------------
# 1. Probe - which tags have any hits in each workload
# ---------------------------------------------------------------------------------------------
$TagsByWorkload = @{}
foreach ($wl in @($Sites | Select-Object -ExpandProperty Workload -Unique)) {
    if ($SkipProbe) { $TagsByWorkload[$wl] = @($Tags); continue }

    $probeFile = Join-Path $OutDir "probe_$wl.csv"
    $probe = @{}
    if ((Test-Path $probeFile) -and -not $RefreshProbe) {
        foreach ($p in @(Import-Csv $probeFile)) {
            if ($p.Status -eq 'OK') { $probe["$($p.TagType)|$($p.TagName)".ToLowerInvariant()] = $p }
        }
        Write-Log ("Probe cache loaded for {0}: {1} tags" -f $wl, $probe.Count)
    }

    $i = 0
    foreach ($t in $Tags) {
        $i++
        $key = "$($t.TagType)|$($t.TagName)".ToLowerInvariant()
        if ($probe.ContainsKey($key)) { continue }
        Write-Progress -Activity "Probing $wl" -Status $t.TagName -PercentComplete (100 * $i / $Tags.Count)
        $rec = [ordered]@{ TagType = $t.TagType; TagName = $t.TagName; Workload = $wl; TotalCount = ''; Status = ''; Checked = (Get-Date -Format s) }
        try {
            $r = Invoke-CE @{ TagType = $t.TagType; TagName = $t.TagName; Workload = $wl; PageSize = 1 } 2
            $rec.TotalCount = [long]$r[0].TotalCount
            $rec.Status = 'OK'
        } catch {
            $rec.Status = 'Error: ' + $_.Exception.Message
            Write-Log "  probe failed: $($t.TagType) '$($t.TagName)' - $($_.Exception.Message)" 'WARN'
        }
        $obj = [pscustomobject]$rec
        $obj | Export-Csv -Path $probeFile -Append -NoTypeInformation -Encoding UTF8
        if ($rec.Status -eq 'OK') { $probe[$key] = $obj }
    }
    Write-Progress -Activity "Probing $wl" -Completed

    $hit = @($Tags | Where-Object {
        $k = "$($_.TagType)|$($_.TagName)".ToLowerInvariant()
        $probe.ContainsKey($k) -and [long]$probe[$k].TotalCount -gt 0
    })
    $TagsByWorkload[$wl] = $hit
    Write-Log ("{0}: {1} of {2} tags have hits tenant-wide" -f $wl, $hit.Count, $Tags.Count)
}

# ---------------------------------------------------------------------------------------------
# 2. Extract - item level, per site per tag
# ---------------------------------------------------------------------------------------------
foreach ($site in $Sites) {
    $siteCache = Join-Path $CacheDir $site.CacheKey
    if (-not (Test-Path $siteCache)) { New-Item -ItemType Directory -Path $siteCache | Out-Null }
    $siteTags = @($TagsByWorkload[$site.Workload])
    Write-Log ("== {0} ({1}) - {2} tags to check" -f $site.SiteName, $site.SiteUrl, $siteTags.Count)

    $j = 0
    foreach ($t in $siteTags) {
        $j++
        $base = Join-Path $siteCache ('{0}__{1}' -f $t.TagType, (Get-SafeName $t.TagName 80))
        $doneFile = "$base.done"
        $csvFile  = "$base.csv"
        if ((Test-Path $doneFile) -and -not $Force) { continue }
        Write-Progress -Activity "Extracting $($site.SiteName)" -Status "$($t.TagType): $($t.TagName)" -PercentComplete (100 * $j / [Math]::Max(1, $siteTags.Count))

        $targetGuid = ''
        if ($t.TagType -eq 'SensitiveInformationType' -and $SitGuidByName.ContainsKey($t.TagName)) { $targetGuid = $SitGuidByName[$t.TagName] }

        $rows   = New-Object System.Collections.Generic.List[object]
        $seenFiles = @{}
        $dupes  = 0
        $raw    = 0
        $cookie = $null
        $lastCookie = $null
        $total  = ''
        $pages  = 0
        $started = Get-Date
        try {
            do {
                $params = @{ TagType = $t.TagType; TagName = $t.TagName; Workload = $site.Workload; SiteUrl = $site.SiteUrl; PageSize = $PageSize }
                if ($cookie) { $params.PageCookie = $cookie }
                $r = Invoke-CE $params
                $pages++
                $meta = $r[0]
                if ($pages -eq 1) { $total = [string]$meta.TotalCount }
                $n = 0
                if ($meta.PSObject.Properties['RecordsReturned']) { $n = [int]$meta.RecordsReturned }
                $n = [Math]::Min($n, $r.Count - 1)

                if ($n -gt 0) {
                    foreach ($rec in $r[1..$n]) {
                        $fileUrl  = [string](Get-Prop $rec @('FileUrl', 'FileURL', 'Url'))
                        $fileName = [string](Get-Prop $rec @('FileName', 'DocumentName', 'Name'))
                        if (-not $fileName -and $fileUrl) { $fileName = ($fileUrl -split '/')[-1] }
                        $raw++
                        # De-duplicate on full URL only - never on file name, which repeats across folders.
                        if ($fileUrl) {
                            if ($seenFiles.ContainsKey($fileUrl)) { $dupes++; continue }
                            $seenFiles[$fileUrl] = $true
                        }

                        $ext = Get-Extension $fileName
                        $conf = $null
                        if ($t.TagType -eq 'SensitiveInformationType') { $conf = Get-SitConfidence (Get-Prop $rec @('SensitiveInfoTypesData')) $targetGuid }

                        $labelGuid = [string](Get-Prop $rec @('SensitivityLabel', 'SensitivityLabelId'))
                        $labelName = $labelGuid
                        if ($labelGuid -match $GuidPattern -and $LabelNameByGuid.ContainsKey($Matches[0].ToLowerInvariant())) { $labelName = $LabelNameByGuid[$Matches[0].ToLowerInvariant()] }

                        $row = [ordered]@{
                            SiteName         = $site.SiteName
                            SiteUrl          = $site.SiteUrl
                            Workload         = $site.Workload
                            TagType          = $t.TagType
                            TagName          = $t.TagName
                            DocType          = Get-DocType $ext
                            Extension        = $ext
                            FileName         = $fileName
                            FileUrl          = $fileUrl
                            HitsHigh         = ''
                            HitsMedium       = ''
                            HitsLow          = ''
                            Confidence       = ''
                            Note             = ''
                            SensitivityLabel = $labelName
                            LastModifiedTime = [string](Get-Prop $rec @('LastModifiedTime', 'DocLastModifiedTime'))
                        }
                        if ($conf) {
                            $row.HitsHigh = $conf.High; $row.HitsMedium = $conf.Medium; $row.HitsLow = $conf.Low
                            $row.Confidence = $conf.Level; $row.Note = $conf.Note
                        } else {
                            $row.Confidence = 'N/A'
                        }
                        $rows.Add([pscustomobject]$row)
                    }
                }

                $more = ConvertTo-Bool $meta.MorePagesAvailable
                $lastCookie = $cookie
                $cookie = $null
                if ($more) { $cookie = [string]$meta.PageCookie }
                if ($more -and ($cookie -eq $lastCookie -or -not $cookie)) {
                    Write-Log "    page cookie did not advance - stopping this tag early (results may be partial)" 'WARN'
                    $more = $false
                }
            } while ($more)

            if ($rows.Count -gt 0) { $rows | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 }
            elseif (Test-Path $csvFile) { Remove-Item $csvFile }
            ('TotalCount={0}; RecordsReturned={1}; DuplicatesRemoved={2}; Files={3}; Pages={4}; Seconds={5:N0}; Finished={6}' -f $total, $raw, $dupes, $rows.Count, $pages, ((Get-Date) - $started).TotalSeconds, (Get-Date -Format s)) |
                Set-Content -Path $doneFile -Encoding UTF8
            if ($rows.Count -gt 0) {
                Write-Log ("  {0}: {1} - {2} files (TotalCount {3}; records returned {4}; duplicates removed {5})" -f $t.TagType, $t.TagName, $rows.Count, $total, $raw, $dupes)
            }
        } catch {
            Write-Log "  FAILED $($t.TagType) '$($t.TagName)' on $($site.SiteName): $($_.Exception.Message) - re-run to retry" 'ERROR'
        }
    }
    Write-Progress -Activity "Extracting $($site.SiteName)" -Completed
}

# ---------------------------------------------------------------------------------------------
# 3. Report - rebuild from cache for the sites in this run
# ---------------------------------------------------------------------------------------------
# One streaming pass with hashtables: Group-Object and Select-Object -Unique are quadratic in
# Windows PowerShell 5.1 and take hours (and GBs of memory) once a site has ~1M detail rows.
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'
$summaryFile = Join-Path $OutDir "SiteSummary_$stamp.csv"
$detailFile  = Join-Path $OutDir "SiteDetail_$stamp.csv"
if (Test-Path $detailFile) { Remove-Item $detailFile }   # rows are appended below
$groups = @{}
$pending = 0
$detailCount = 0
Write-Log 'Building report from cache...'
foreach ($site in $Sites) {
    $siteCache = Join-Path $CacheDir $site.CacheKey
    if (-not (Test-Path $siteCache)) { continue }
    foreach ($t in @($TagsByWorkload[$site.Workload])) {
        $base = Join-Path $siteCache ('{0}__{1}' -f $t.TagType, (Get-SafeName $t.TagName 80))
        if (-not (Test-Path "$base.done")) { $pending++; continue }
        if (-not (Test-Path "$base.csv")) { continue }
        $rows = @(Import-Csv "$base.csv")
        foreach ($row in $rows) {
            $row.SiteName = $site.SiteName   # Sites.csv wins if a name was corrected
            $key = '{0}|{1}|{2}|{3}' -f $row.SiteUrl, $row.TagType, $row.TagName, $row.DocType
            $g = $groups[$key]
            if (-not $g) {
                $g = [pscustomobject]@{
                    First    = $row
                    Files    = New-Object 'System.Collections.Generic.HashSet[string]'
                    Exts     = New-Object 'System.Collections.Generic.HashSet[string]'
                    Notes    = New-Object 'System.Collections.Generic.List[string]'
                    Rows     = 0
                    WithData = 0
                    Hits     = 0
                    Best     = ''
                }
                $groups[$key] = $g
            }
            $g.Rows++
            $fileKey = $row.FileName
            if ($row.FileUrl) { $fileKey = $row.FileUrl }
            [void]$g.Files.Add([string]$fileKey)
            if ($row.Extension) { [void]$g.Exts.Add([string]$row.Extension) }
            if ($row.TagType -eq 'SensitiveInformationType') {
                if ($row.HitsHigh -ne '' -or $row.HitsMedium -ne '' -or $row.HitsLow -ne '') {
                    $g.WithData++
                    # Content Explorer's buckets are cumulative (a High match is also counted in Medium and Low),
                    # so a file's total hits is its largest bucket, not the sum of the three.
                    $g.Hits += [Math]::Max([int]$row.HitsLow, [Math]::Max([int]$row.HitsMedium, [int]$row.HitsHigh))
                    if ($LevelRank[[string]$row.Confidence] -gt $LevelRank[$g.Best]) { $g.Best = [string]$row.Confidence }
                }
                if ($row.Note -and -not $g.Notes.Contains([string]$row.Note)) { $g.Notes.Add([string]$row.Note) }
            }
        }
        if ($rows.Count -gt 0) {
            $rows | Export-Csv -Path $detailFile -Append -NoTypeInformation -Encoding UTF8
            $detailCount += $rows.Count
        }
    }
}

$summary = foreach ($g in $groups.Values) {
    $first = $g.First
    $hits = ''; $best = ''; $note = ''
    if ($first.TagType -eq 'SensitiveInformationType') {
        if ($g.WithData -gt 0) { $hits = $g.Hits; $best = $g.Best }
        $missing = $g.Rows - $g.WithData
        if ($missing -gt 0) { $note = "$missing of $($g.Rows) files without hit data: $($g.Notes -join '; ')" }
    } else {
        $best = 'N/A'
    }

    [pscustomobject][ordered]@{
        'Site Name'                  = $first.SiteName
        'Site URL'                   = $first.SiteUrl
        'Tag Type'                   = $TagTypeLabel[[string]$first.TagType]
        'SIT / Trainable Classifier' = $first.TagName
        'Doc Type'                   = $first.DocType
        'Extensions'                 = (@($g.Exts) | Sort-Object) -join ', '
        'Count of Files'             = $g.Files.Count
        'Count of Hits'              = $hits
        'Highest Confidence'         = $best
        'Note'                       = $note
    }
}

if (@($summary).Count -eq 0) {
    Write-Log 'No files with hits found for these sites/tags - no report written.' 'WARN'
    if ($pending -gt 0) { Write-Log "$pending (site, tag) pairs not finished - re-run the same command to resume." 'WARN' }
    return
}

@($summary) | Sort-Object 'Site Name', 'Tag Type', 'SIT / Trainable Classifier', @{ Expression = 'Count of Files'; Descending = $true } |
    Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8

Write-Log ("Summary: {0} rows -> {1}" -f @($summary).Count, $summaryFile)
Write-Log ("Detail : {0} rows -> {1}" -f $detailCount, $detailFile)
if ($pending -gt 0) { Write-Log "$pending (site, tag) pairs not finished - re-run the same command to resume." 'WARN' }
