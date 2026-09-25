<#
.SYNOPSIS
    Read-only SQL Server performance health check that writes a Word (.docx) report.

.DESCRIPTION
    Connects to one or more SQL Server instances, runs read-only diagnostic queries
    (DMVs and system catalogs), and produces one Word report per instance listing the
    problems worth looking at, ranked by severity, with a recommendation for each.

    Nothing to install: uses only .NET classes that ship with Windows PowerShell 5.1
    (also works in PowerShell 7). No Office, Python or modules required.
    Nothing is changed on the audited server.

    Permissions needed: VIEW SERVER STATE (sysadmin not required) and read access to
    msdb backup history.

.PARAMETER SqlInstance
    One or more instances: SERVER, SERVER\INSTANCE or SERVER,PORT.

.PARAMETER Credential
    SQL login (use Get-Credential). Omit to use Windows authentication.

.PARAMETER OutputFolder
    Where to write the reports. Default: current folder.

.PARAMETER Top
    Number of rows in "top" tables (queries, missing indexes, files). Default 10.

.EXAMPLE
    .\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01

.EXAMPLE
    .\Invoke-SqlPerfAudit.ps1 -SqlInstance 'SQLPROD01\INST1','SQLPROD02' -OutputFolder C:\Audits

.EXAMPLE
    .\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01 -Credential (Get-Credential) -TrustServerCertificate

.NOTES
    If the script is blocked by execution policy, run it with:
    powershell.exe -ExecutionPolicy Bypass -File .\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string[]]$SqlInstance,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$OutputFolder = (Get-Location).Path,
    [ValidateRange(5, 50)][int]$Top = 10,
    [switch]$Encrypt,
    [switch]$TrustServerCertificate
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.IO.Compression
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:ScriptVersion = '1.0'

# =============================================================================
# General helpers
# =============================================================================

function Get-ShortError($Err) {
    $e = $Err
    if ($e -is [System.Management.Automation.ErrorRecord]) { $e = $e.Exception }
    while ($null -ne $e.InnerException) { $e = $e.InnerException }
    $m = (([string]$e.Message) -replace '\s+', ' ').Trim()
    if ($m.Length -gt 300) { $m = $m.Substring(0, 297) + '...' }
    return $m
}

function Format-N0($v) { if ($null -eq $v) { return '' }; return ([double]$v).ToString('N0', $script:Inv) }
function Format-N1($v) { if ($null -eq $v) { return '' }; return ([double]$v).ToString('N1', $script:Inv) }

function Format-Value($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [bool]) { if ($v) { return 'Yes' } else { return 'No' } }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd HH:mm', $script:Inv) }
    if ($v -is [int] -or $v -is [long] -or $v -is [int16] -or $v -is [byte]) { return ([long]$v).ToString('N0', $script:Inv) }
    if ($v -is [decimal] -or $v -is [double] -or $v -is [single]) {
        if ([math]::Abs([double]$v) -ge 1000) { return ([double]$v).ToString('N0', $script:Inv) }
        return ([double]$v).ToString('N1', $script:Inv)
    }
    $s = [string]$v
    if ($s.Length -gt 400) { $s = $s.Substring(0, 397) + '...' }
    return $s
}

function Test-Numeric($v) {
    return ($v -is [int] -or $v -is [long] -or $v -is [int16] -or $v -is [byte] -or $v -is [decimal] -or $v -is [double] -or $v -is [single])
}

function Join-Names($Names, [int]$Max = 6) {
    $n = @($Names)
    if ($n.Count -le $Max) { return ($n -join ', ') }
    return (($n[0..($Max - 1)] -join ', ') + ' and ' + ($n.Count - $Max) + ' more')
}

function Format-QueryText([string]$Text) {
    if (-not $Text) { return '' }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt 300) { $t = $t.Substring(0, 297) + '...' }
    return $t
}

# =============================================================================
# SQL access
# =============================================================================

function Connect-Instance([string]$Instance) {
    $csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $csb['Data Source'] = $Instance
    $csb['Initial Catalog'] = 'master'
    $csb['Application Name'] = 'SQL Perf Audit (read-only)'
    $csb['Connect Timeout'] = 30
    if ($Encrypt) { $csb['Encrypt'] = $true }
    if ($TrustServerCertificate) { $csb['TrustServerCertificate'] = $true }
    if ($Credential) {
        $csb['Integrated Security'] = $false
        $csb['User ID'] = $Credential.UserName
        $csb['Password'] = $Credential.GetNetworkCredential().Password
    } else {
        $csb['Integrated Security'] = $true
    }
    $c = New-Object System.Data.SqlClient.SqlConnection($csb.ConnectionString)
    $c.Open()
    return $c
}

# Runs a query and returns rows as objects (DBNull converted to $null).
# Always call as @(Invoke-Q ...) so a single row is still an array.
function Invoke-Q([string]$Sql) {
    $cmd = $script:Conn.CreateCommand()
    $cmd.CommandText = "SET NOCOUNT ON;`r`n" + $Sql
    $cmd.CommandTimeout = 300
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $ds = New-Object System.Data.DataSet
    [void]$da.Fill($ds)
    if ($ds.Tables.Count -eq 0) { return }
    $t = $ds.Tables[$ds.Tables.Count - 1]
    $out = New-Object System.Collections.ArrayList
    foreach ($row in $t.Rows) {
        $h = [ordered]@{}
        foreach ($col in $t.Columns) {
            $v = $row[$col]
            if ($v -is [DBNull]) { $v = $null }
            $h[$col.ColumnName] = $v
        }
        [void]$out.Add([pscustomobject]$h)
    }
    return $out.ToArray()
}

function Get-One([string]$Sql) {
    $r = @(Invoke-Q $Sql)
    if ($r.Count -gt 0) { return $r[0] }
    return $null
}

# =============================================================================
# Findings, sections, tables
# =============================================================================

function Add-Finding([string]$Severity, [string]$Title, [string]$Detail, [string]$Fix) {
    [void]$script:Findings.Add([pscustomobject]@{
        Severity = $Severity; Area = $script:CurrentSection.Title
        Title = $Title; Detail = $Detail; Fix = $Fix
    })
}

function Add-Table([string]$Caption, $Rows, [double[]]$Weights, [string]$Note) {
    if ($null -eq $Rows) { $Rows = @() }
    [void]$script:CurrentSection.Tables.Add([pscustomobject]@{
        Caption = $Caption; Rows = @($Rows); Weights = $Weights; Note = $Note
    })
}

function Add-Note([string]$Text) {
    [void]$script:CurrentSection.Notes.Add($Text)
}

function Invoke-Check([string]$Title, [string]$Intro, [scriptblock]$Body) {
    Write-Host ('    ' + $Title) -ForegroundColor Gray
    $sec = [pscustomobject]@{
        Title = $Title; Intro = $Intro; Error = $null
        Tables = (New-Object System.Collections.ArrayList)
        Notes  = (New-Object System.Collections.ArrayList)
    }
    $script:CurrentSection = $sec
    try {
        & $Body
    } catch {
        $sec.Error = Get-ShortError $_
        Write-Warning ('    ' + $Title + ' could not run: ' + $sec.Error)
    }
    [void]$script:Sections.Add($sec)
}

# =============================================================================
# Reference data
# =============================================================================

# Waits that are normal background activity and are ignored.
$script:BenignWaits = @(
    'BROKER_EVENTHANDLER','BROKER_RECEIVE_WAITFOR','BROKER_TASK_STOP','BROKER_TO_FLUSH','BROKER_TRANSMITTER',
    'CHECKPOINT_QUEUE','CHKPT','CLR_AUTO_EVENT','CLR_MANUAL_EVENT','CLR_SEMAPHORE','CXCONSUMER',
    'DBMIRROR_DBM_EVENT','DBMIRROR_EVENTS_QUEUE','DBMIRROR_WORKER_QUEUE','DBMIRRORING_CMD','DIRTY_PAGE_POLL',
    'DISPATCHER_QUEUE_SEMAPHORE','EXECSYNC','FSAGENT','FT_IFTS_SCHEDULER_IDLE_WAIT','FT_IFTSHC_MUTEX',
    'HADR_CLUSAPI_CALL','HADR_FILESTREAM_IOMGR_IOCOMPLETION','HADR_LOGCAPTURE_WAIT','HADR_NOTIFICATION_DEQUEUE',
    'HADR_TIMER_TASK','HADR_WORK_QUEUE','KSOURCE_WAKEUP','LAZYWRITER_SLEEP','LOGMGR_QUEUE','MEMORY_ALLOCATION_EXT',
    'ONDEMAND_TASK_QUEUE','PARALLEL_REDO_DRAIN_WORKER','PARALLEL_REDO_LOG_CACHE','PARALLEL_REDO_TRAN_LIST',
    'PARALLEL_REDO_WORKER_SYNC','PARALLEL_REDO_WORKER_WAIT_WORK','PREEMPTIVE_OS_FLUSHFILEBUFFERS',
    'PVS_PREALLOCATE','PWAIT_ALL_COMPONENTS_INITIALIZED','PWAIT_DIRECTLOGCONSUMER_GETNEXT',
    'PWAIT_EXTENSIBILITY_CLEANUP_TASK','REDO_THREAD_PENDING_WORK','REQUEST_FOR_DEADLOCK_SEARCH','RESOURCE_QUEUE',
    'SERVER_IDLE_CHECK','SNI_HTTP_ACCEPT','SOS_WORK_DISPATCHER','SP_SERVER_DIAGNOSTICS_SLEEP',
    'SQLTRACE_BUFFER_FLUSH','SQLTRACE_INCREMENTAL_FLUSH_SLEEP','SQLTRACE_WAIT_ENTRIES','VDI_CLIENT_OTHER',
    'WAIT_FOR_RESULTS','WAITFOR','WAITFOR_TASKSHUTDOWN','WAIT_XTP_RECOVERY','WAIT_XTP_HOST_WAIT',
    'WAIT_XTP_OFFLINE_CKPT_NEW_LOG','WAIT_XTP_CKPT_CLOSE','XE_DISPATCHER_JOIN','XE_DISPATCHER_WAIT',
    'XE_TIMER_EVENT','XE_LIVE_TARGET_TVF'
)
$script:BenignPrefixes = @('SLEEP_', 'QDS_', 'PREEMPTIVE_XE_', 'XE_', 'BROKER_')

# Plain-language meaning of common waits (matched by prefix, first match wins).
$script:WaitInfo = @(
    @{ P = 'THREADPOOL';        M = 'Worker thread starvation: requests waited for a thread. Can cause timeouts and apparent freezes.'; F = 'Find the cause (usually heavy blocking or excessive parallelism). Do not simply raise max worker threads.' },
    @{ P = 'RESOURCE_SEMAPHORE';M = 'Queries waited for memory (sorts, hash joins) before they could start.'; F = 'Tune queries with large memory grants (missing indexes, bad estimates) and check memory settings.' },
    @{ P = 'CX';                M = 'Parallelism: threads of parallel queries waiting for each other.'; F = 'Check MAXDOP and cost threshold for parallelism, then tune the largest parallel queries.' },
    @{ P = 'SOS_SCHEDULER_YIELD'; M = 'CPU pressure: threads used their full CPU time slice, often from scanning data in memory.'; F = 'Tune the top CPU queries (see Top queries) and look for scans that indexes could avoid.' },
    @{ P = 'PAGEIOLATCH_';      M = 'Reading data pages from disk into memory.'; F = 'Check disk read latency and memory pressure; reduce data read by indexing and query tuning.' },
    @{ P = 'WRITELOG';          M = 'Commits waiting for the transaction log to be written to disk.'; F = 'Check log file write latency; put logs on fast storage; avoid many tiny transactions.' },
    @{ P = 'LCK_M_';            M = 'Blocking: sessions waiting for locks held by others.'; F = 'Find blocking chains and long transactions; consider READ_COMMITTED_SNAPSHOT; index to shorten locks.' },
    @{ P = 'PAGELATCH_';        M = 'Contention on hot pages in memory, often tempdb allocation pages.'; F = 'Check tempdb file count; consider OPTIMIZE_FOR_SEQUENTIAL_KEY for hot insert tables.' },
    @{ P = 'ASYNC_NETWORK_IO';  M = 'SQL Server waited for the client application to consume results.'; F = 'Usually an application issue (large result sets read row by row, slow app server).' },
    @{ P = 'HADR_SYNC_COMMIT';  M = 'Commits waiting for synchronous Availability Group replicas.'; F = 'Check network and disk latency on secondary replicas.' },
    @{ P = 'IO_COMPLETION';     M = 'Non-data I/O such as sort/hash spills to tempdb.'; F = 'Look for tempdb spills and storage latency.' },
    @{ P = 'ASYNC_IO_COMPLETION'; M = 'Backups, file growth and bulk I/O.'; F = 'Check backup schedule; enable instant file initialization; pre-size files.' },
    @{ P = 'BACKUP';            M = 'Backup activity.'; F = 'Normal during backups; schedule outside peak hours.' },
    @{ P = 'OLEDB';             M = 'Linked server calls or some DBCC/monitoring activity.'; F = 'Review linked server queries.' },
    @{ P = 'PREEMPTIVE_';       M = 'Calls out to the operating system (authentication, file operations, CLR).'; F = 'Identify the specific call; often Active Directory or file system latency.' }
)

function Get-WaitInfo([string]$WaitType) {
    foreach ($w in $script:WaitInfo) { if ($WaitType.StartsWith($w.P)) { return $w } }
    return @{ P = ''; M = 'Less common wait type; see Microsoft documentation.'; F = 'Investigate if it stays near the top.' }
}

function Test-BenignWait([string]$WaitType) {
    if ($script:BenignWaits -contains $WaitType) { return $true }
    foreach ($p in $script:BenignPrefixes) { if ($WaitType.StartsWith($p)) { return $true } }
    return $false
}

function Get-RecommendedMaxMemoryMB([double]$PhysMB) {
    # Common starting point: keep 1 GB for the OS, +1 GB per 4 GB from 4-16 GB, +1 GB per 8 GB above 16 GB.
    $gb = $PhysMB / 1024
    $reserve = 1 + [math]::Min([math]::Max($gb - 4, 0), 12) / 4 + [math]::Max($gb - 16, 0) / 8
    return [math]::Max([math]::Floor(($gb - $reserve) * 1024), [math]::Floor($PhysMB / 2))
}

function Get-RecommendedMaxDop([int]$PerNode, [int]$Nodes) {
    if ($PerNode -le 0) { return 8 }
    if ($Nodes -le 1) { if ($PerNode -le 8) { return $PerNode } else { return 8 } }
    if ($PerNode -le 16) { return $PerNode }
    return [int][math]::Min(16, [math]::Floor($PerNode / 2))
}

# =============================================================================
# Checks
# =============================================================================

function Test-Instance {
    $p = Get-One @'
SELECT CAST(SERVERPROPERTY('ServerName') AS nvarchar(256)) AS ServerName,
       CAST(SERVERPROPERTY('Edition') AS nvarchar(256)) AS Edition,
       CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64)) AS ProductVersion,
       CAST(SERVERPROPERTY('ProductLevel') AS nvarchar(64)) AS ProductLevel,
       CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(64)) AS UpdateLevel,
       SUSER_SNAME() AS AuditLogin,
       GETDATE() AS ServerTime
'@
    $si = Get-One 'SELECT * FROM sys.dm_os_sys_info'
    $sch = Get-One @'
SELECT SUM(CASE WHEN status = 'VISIBLE ONLINE' THEN 1 ELSE 0 END) AS online_cnt,
       SUM(CASE WHEN status = 'VISIBLE OFFLINE' THEN 1 ELSE 0 END) AS offline_cnt,
       COUNT(DISTINCT CASE WHEN status = 'VISIBLE ONLINE' THEN parent_node_id END) AS nodes
FROM sys.dm_os_schedulers
'@
    $major = [int]($p.ProductVersion.Split('.')[0])
    $names = @{ 10 = '2008 / 2008 R2'; 11 = '2012'; 12 = '2014'; 13 = '2016'; 14 = '2017'; 15 = '2019'; 16 = '2022'; 17 = '2025' }
    $verName = 'SQL Server ' + $(if ($names.ContainsKey($major)) { $names[$major] } else { 'version ' + $major })

    $physMB = 0.0
    if ($null -ne $si.physical_memory_kb) { $physMB = [double]$si.physical_memory_kb / 1024 } elseif ($null -ne $si.physical_memory_in_bytes) { $physMB = [double]$si.physical_memory_in_bytes / 1MB }

    $online = [int]$sch.online_cnt; $offline = [int]$sch.offline_cnt
    $nodes = [math]::Max(1, [int]$sch.nodes)
    $uptime = ([datetime]$p.ServerTime) - ([datetime]$si.sqlserver_start_time)

    $ifi = $null
    try {
        $svc = Get-One "SELECT instant_file_initialization_enabled FROM sys.dm_server_services WHERE servicename LIKE N'SQL Server (%'"
        if ($svc) { $ifi = [string]$svc.instant_file_initialization_enabled }
    } catch { }

    $script:Ctx.Server     = [string]$p.ServerName
    $script:Ctx.Version    = $verName + ' ' + $p.ProductLevel + $(if ($p.UpdateLevel) { ' ' + $p.UpdateLevel } else { '' }) + ' (' + $p.ProductVersion + ')'
    $script:Ctx.Edition    = [string]$p.Edition
    $script:Ctx.Major      = $major
    $script:Ctx.CpuCount   = $online
    $script:Ctx.Nodes      = $nodes
    $script:Ctx.PerNode    = [int][math]::Ceiling($online / $nodes)
    $script:Ctx.PhysMB     = $physMB
    $script:Ctx.ServerTime = [datetime]$p.ServerTime
    $script:Ctx.UptimeDays = $uptime.TotalDays
    $script:Ctx.StartTime  = [datetime]$si.sqlserver_start_time
    $script:Ctx.AuditLogin = [string]$p.AuditLogin
    $script:Ctx.MaxWorkers = $si.max_workers_count

    $uptimeText = ([int][math]::Floor($uptime.TotalDays)).ToString() + ' days ' + $uptime.Hours + ' h'
    $rows = @(
        [pscustomobject]@{ Property = 'Server';           Value = $p.ServerName },
        [pscustomobject]@{ Property = 'Version';          Value = $script:Ctx.Version },
        [pscustomobject]@{ Property = 'Edition';          Value = $p.Edition },
        [pscustomobject]@{ Property = 'Logical CPUs used by SQL'; Value = ([string]$online + ' (' + $nodes + ' NUMA node(s))') },
        [pscustomobject]@{ Property = 'Physical memory';  Value = ((Format-N0 $physMB) + ' MB') },
        [pscustomobject]@{ Property = 'Virtual machine';  Value = [string]$si.virtual_machine_type_desc },
        [pscustomobject]@{ Property = 'Started';          Value = $si.sqlserver_start_time },
        [pscustomobject]@{ Property = 'Uptime';           Value = $uptimeText },
        [pscustomobject]@{ Property = 'Instant file initialization'; Value = $(if ($ifi -eq 'Y') { 'Enabled' } elseif ($ifi -eq 'N') { 'Disabled' } else { 'Unknown' }) }
    )
    Add-Table 'Instance summary' $rows @(1, 2)

    if ($major -ge 10 -and $major -le 13) {
        Add-Finding 'Medium' ($verName + ' is out of Microsoft extended support') 'No more security or performance fixes are released for this version (SQL Server 2016 support ended in July 2026).' 'Plan an upgrade to a supported version.'
    } elseif ($major -eq 14) {
        Add-Finding 'Info' 'SQL Server 2017 extended support ends in October 2027' 'After that date no more fixes are released.' 'Start planning the upgrade.'
    }
    if ($offline -gt 0) {
        Add-Finding 'High' ([string]$offline + ' CPU(s) cannot be used by SQL Server') 'Some schedulers are VISIBLE OFFLINE, usually because of edition licensing limits or the VM socket/core layout.' 'Check the edition CPU limit and reconfigure the VM with fewer sockets and more cores per socket, or change edition.'
    }
    if ($ifi -eq 'N') {
        Add-Finding 'Medium' 'Instant file initialization is disabled' 'Data file growth and restores must zero out new space, which stalls activity during autogrowth.' 'Grant the SQL Server service account the "Perform volume maintenance tasks" right and restart the service.'
    }
    if ($uptime.TotalDays -lt 7) {
        Add-Finding 'Info' ('SQL Server was restarted ' + (Format-N1 $uptime.TotalDays) + ' days ago') 'Most statistics in this report accumulate since the last restart, so they may not reflect a typical week.' 'Re-run the audit after a full business cycle for a more reliable picture.'
    }
}

function Test-Configuration {
    $rows = @(Invoke-Q @'
SELECT name, CAST(value AS bigint) AS value, CAST(value_in_use AS bigint) AS value_in_use
FROM sys.configurations
WHERE name IN (N'max server memory (MB)', N'min server memory (MB)', N'max degree of parallelism',
               N'cost threshold for parallelism', N'optimize for ad hoc workloads', N'priority boost',
               N'lightweight pooling', N'max worker threads')
'@)
    $cfg = @{}
    foreach ($r in $rows) { $cfg[$r.name] = $r }
    $assess = @{}

    $phys = [double]$script:Ctx.PhysMB
    $maxMem = [long]$cfg['max server memory (MB)'].value_in_use
    $script:Ctx.MaxMemDefault = ($maxMem -ge 2147483647)
    if ($phys -gt 0) {
        $rec = Get-RecommendedMaxMemoryMB $phys
        if ($maxMem -ge 2147483647) {
            $assess['max server memory (MB)'] = 'Not set (unlimited)'
            Add-Finding 'High' 'Max server memory is not configured' ('It is unlimited on a server with ' + (Format-N0 $phys) + ' MB of RAM, so SQL Server can starve the operating system, causing paging and instability.') ('Set max server memory to about ' + (Format-N0 $rec) + ' MB as a starting point, then keep an eye on free OS memory.')
        } elseif ($maxMem -gt ($phys - ($phys - $rec) / 2)) {
            $assess['max server memory (MB)'] = 'Leaves little for the OS'
            Add-Finding 'Medium' 'Max server memory leaves little room for the operating system' ('It is ' + (Format-N0 $maxMem) + ' MB out of ' + (Format-N0 $phys) + ' MB of RAM.') ('Lower it to about ' + (Format-N0 $rec) + ' MB unless you have verified the OS keeps enough free memory.')
        } else {
            $assess['max server memory (MB)'] = 'OK'
        }
    }

    $maxdop = [long]$cfg['max degree of parallelism'].value_in_use
    $cpu = [int]$script:Ctx.CpuCount
    $recDop = Get-RecommendedMaxDop $script:Ctx.PerNode $script:Ctx.Nodes
    if ($maxdop -eq 0 -and $cpu -gt 8) {
        $assess['max degree of parallelism'] = ('Unlimited on ' + $cpu + ' CPUs; suggest ' + $recDop)
        Add-Finding 'Medium' 'MAXDOP is unlimited' ('With ' + $cpu + ' logical CPUs, a single query can use all of them, hurting concurrency.') ('Set max degree of parallelism to ' + $recDop + ' (Microsoft guidance for this CPU/NUMA layout).')
    } elseif ($maxdop -gt $recDop) {
        $assess['max degree of parallelism'] = ('Higher than suggested ' + $recDop)
        Add-Finding 'Low' 'MAXDOP is higher than recommended' ('Current ' + $maxdop + ', suggested ' + $recDop + ' for this CPU/NUMA layout.') ('Consider lowering MAXDOP to ' + $recDop + '.')
    } elseif ($maxdop -eq 1 -and $cpu -gt 1) {
        $assess['max degree of parallelism'] = 'Parallelism disabled'
        Add-Finding 'Info' 'Parallelism is disabled (MAXDOP 1)' 'Large queries cannot use more than one CPU. This is required by some applications (e.g. SharePoint) but slows reporting-type queries.' ('Keep it if the application vendor requires it; otherwise consider MAXDOP ' + $recDop + ' with a higher cost threshold.')
    } else {
        $assess['max degree of parallelism'] = 'OK'
    }

    $ctfp = [long]$cfg['cost threshold for parallelism'].value_in_use
    if ($ctfp -le 5) {
        $assess['cost threshold for parallelism'] = 'Default (5) is too low'
        Add-Finding 'Medium' 'Cost threshold for parallelism is at the default of 5' 'Even cheap queries go parallel, which wastes CPU and increases parallelism waits.' 'Raise it to 50 as a starting point and adjust based on your workload.'
    } elseif ($ctfp -lt 25) {
        $assess['cost threshold for parallelism'] = 'Low'
    } else {
        $assess['cost threshold for parallelism'] = 'OK'
    }

    if ([long]$cfg['optimize for ad hoc workloads'].value_in_use -eq 0) {
        $assess['optimize for ad hoc workloads'] = 'Off - usually worth enabling'
        Add-Finding 'Low' 'Optimize for ad hoc workloads is off' 'Single-use query plans are fully cached and can waste plan cache memory.' 'Enable it; it is safe for almost all workloads.'
    } else { $assess['optimize for ad hoc workloads'] = 'OK' }

    if ([long]$cfg['priority boost'].value_in_use -eq 1) {
        $assess['priority boost'] = 'ON - not supported'
        Add-Finding 'High' 'Priority boost is enabled' 'This deprecated setting can starve the operating system and cluster services.' 'Disable it and restart SQL Server.'
    }
    if ([long]$cfg['lightweight pooling'].value_in_use -eq 1) {
        $assess['lightweight pooling'] = 'ON - not recommended'
        Add-Finding 'Medium' 'Lightweight pooling (fiber mode) is enabled' 'Fiber mode is rarely beneficial and breaks several features.' 'Disable it unless a specific test proved a benefit.'
    }
    if ([long]$cfg['max worker threads'].value_in_use -ne 0) {
        $assess['max worker threads'] = 'Changed from default'
        Add-Finding 'Low' 'Max worker threads has been changed from the default' 'Raising it usually hides blocking or parallelism problems rather than fixing them.' 'Reset to 0 (automatic) unless there is a documented reason.'
    }

    $pending = @($rows | Where-Object { $_.value -ne $_.value_in_use } | ForEach-Object { $_.name })
    if ($pending.Count -gt 0) {
        Add-Finding 'Low' 'Configuration changes are pending' ('Configured but not active: ' + (Join-Names $pending) + '.') 'Run RECONFIGURE or restart SQL Server during a maintenance window.'
    }

    $out = foreach ($r in ($rows | Sort-Object name)) {
        [pscustomobject]@{
            'Setting' = $r.name
            'Value in use' = $r.value_in_use
            'Assessment' = $(if ($assess.ContainsKey($r.name)) { $assess[$r.name] } else { 'OK' })
        }
    }
    Add-Table 'Key server settings' $out @(3, 1.4, 3)
}

function Get-PerfCounters {
    $rows = @(Invoke-Q @'
SELECT RTRIM(counter_name) AS counter_name, cntr_value
FROM sys.dm_os_performance_counters
WHERE (object_name LIKE N'%Buffer Manager%' AND counter_name = N'Page life expectancy')
   OR (object_name LIKE N'%Memory Manager%' AND counter_name IN (N'Memory Grants Pending', N'Target Server Memory (KB)', N'Total Server Memory (KB)'))
   OR (object_name LIKE N'%:Locks%' AND counter_name = N'Number of Deadlocks/sec' AND instance_name = N'_Total')
   OR (object_name LIKE N'%SQL Statistics%' AND counter_name IN (N'Batch Requests/sec', N'SQL Compilations/sec'))
'@)
    $h = @{}
    foreach ($r in $rows) { if (-not $h.ContainsKey($r.counter_name)) { $h[$r.counter_name] = [double]$r.cntr_value } }
    return $h
}

function Test-Waits {
    $all = @(Invoke-Q 'SELECT wait_type, waiting_tasks_count, wait_time_ms, signal_wait_time_ms FROM sys.dm_os_wait_stats WHERE waiting_tasks_count > 0 AND wait_time_ms > 0')
    $w = @($all | Where-Object { -not (Test-BenignWait $_.wait_type) } | Sort-Object -Property { [double]$_.wait_time_ms } -Descending)
    $total = 0.0; $signal = 0.0
    foreach ($r in $w) { $total += [double]$r.wait_time_ms; $signal += [double]$r.signal_wait_time_ms }
    if ($total -le 0) { Add-Note 'No significant waits recorded since startup.'; return }

    $out = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($r in $w) {
        if ($i -ge $Top) { break }
        $pct = 100.0 * [double]$r.wait_time_ms / $total
        $avg = [double]$r.wait_time_ms / [math]::Max(1, [double]$r.waiting_tasks_count)
        $info = Get-WaitInfo $r.wait_type
        [void]$out.Add([pscustomobject]@{
            'Wait type' = $r.wait_type
            '% of waits' = [math]::Round($pct, 1)
            'Avg wait (ms)' = [math]::Round($avg, 1)
            'What it usually means' = $info.M
        })
        if ($i -lt 5 -and $pct -ge 10) {
            $sev = 'Low'
            if ($pct -ge 25) { $sev = 'Medium' }
            if ($r.wait_type -like 'RESOURCE_SEMAPHORE*' -or $pct -ge 50) { $sev = 'High' }
            if ($r.wait_type -eq 'ASYNC_NETWORK_IO' -or $r.wait_type -like 'BACKUP*') { $sev = 'Low' }
            Add-Finding $sev ($r.wait_type + ' is ' + (Format-N1 $pct) + '% of all waits') ($info.M + ' Average wait ' + (Format-N1 $avg) + ' ms.') $info.F
        }
        $i++
    }
    Add-Table ('Top waits since ' + (Format-Value $script:Ctx.StartTime)) $out @(2.2, 1, 1, 5)

    $tp = @($all | Where-Object { $_.wait_type -eq 'THREADPOOL' })
    if ($tp.Count -gt 0 -and [double]$tp[0].wait_time_ms -ge 10000) {
        Add-Finding 'High' 'Worker thread starvation has occurred (THREADPOOL waits)' ('Requests waited ' + (Format-N0 ([double]$tp[0].wait_time_ms / 1000)) + ' seconds in total for a worker thread. Users would have seen timeouts or a frozen server.') 'Investigate blocking and parallelism at the times it happens; do not simply raise max worker threads.'
    }
    $sigPct = 100.0 * $signal / $total
    if ($sigPct -ge 20) {
        Add-Finding 'Medium' ('Signal waits are ' + (Format-N1 $sigPct) + '% of wait time') 'Tasks spend a large share of time waiting for a CPU after their resource became available, which points to CPU pressure.' 'Tune the top CPU queries and review parallelism settings; confirm CPU capacity.'
    }
}

function Test-Cpu {
    $rb = @(Invoke-Q @'
DECLARE @ts_now bigint = (SELECT cpu_ticks / (cpu_ticks / ms_ticks) FROM sys.dm_os_sys_info);
SELECT TOP (256)
       DATEADD(ms, -1 * (@ts_now - y.[timestamp]), GETDATE()) AS event_time,
       y.sql_cpu,
       100 - y.system_idle - y.sql_cpu AS other_cpu
FROM (SELECT x.[timestamp],
             x.record.value('(./Record/@id)[1]', 'int') AS record_id,
             x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS system_idle,
             x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS sql_cpu
      FROM (SELECT [timestamp], CONVERT(xml, record) AS record
            FROM sys.dm_os_ring_buffers
            WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
              AND record LIKE N'%<SystemHealth>%') AS x) AS y
ORDER BY y.record_id DESC;
'@)
    if ($rb.Count -eq 0) { Add-Note 'No CPU history available.'; return }
    $sum = 0.0; $max = 0.0; $other = 0.0; $busy = 0
    foreach ($r in $rb) {
        $v = [double]$r.sql_cpu
        $sum += $v; $other += [math]::Max(0, [double]$r.other_cpu)
        if ($v -gt $max) { $max = $v }
        if ($v -ge 80) { $busy++ }
    }
    $n = $rb.Count
    $avg = $sum / $n; $avgOther = $other / $n; $busyPct = 100.0 * $busy / $n
    $oldest = ($rb | Sort-Object -Property { [datetime]$_.event_time } | Select-Object -First 1).event_time
    $script:Ctx.CpuAvg = $avg

    Add-Table 'CPU usage (one sample per minute)' @(
        [pscustomobject]@{ Metric = 'Period covered';                     Value = ('Since ' + (Format-Value $oldest) + ' (' + $n + ' minutes)') },
        [pscustomobject]@{ Metric = 'Average SQL Server CPU %';           Value = [math]::Round($avg, 1) },
        [pscustomobject]@{ Metric = 'Peak SQL Server CPU %';              Value = [math]::Round($max, 1) },
        [pscustomobject]@{ Metric = 'Minutes with SQL CPU at 80% or more'; Value = ([string]$busy + ' (' + (Format-N1 $busyPct) + '%)') },
        [pscustomobject]@{ Metric = 'Average CPU % used by other processes'; Value = [math]::Round($avgOther, 1) }
    ) @(2, 2)

    if ($avg -ge 80) {
        Add-Finding 'High' ('SQL Server CPU averages ' + (Format-N1 $avg) + '%') 'The server is CPU-bound; queries queue for CPU and response times suffer.' 'Tune the top CPU queries (see Top queries), check parallelism settings, then consider more CPU.'
    } elseif ($avg -ge 60 -or $busyPct -ge 20) {
        Add-Finding 'Medium' ('SQL Server CPU is high (average ' + (Format-N1 $avg) + '%, peak ' + (Format-N1 $max) + '%)') ('CPU was at 80% or more for ' + (Format-N1 $busyPct) + '% of the period.') 'Tune the top CPU queries (see Top queries) before adding hardware.'
    }
    if ($avgOther -ge 20) {
        Add-Finding 'Medium' ('Other processes use ' + (Format-N1 $avgOther) + '% CPU on this server') 'Something other than SQL Server (antivirus, other services, another instance) competes for CPU.' 'Identify the process on the server; move it off or exclude SQL files from antivirus scanning.'
    }

    $c = $script:Ctx.Counters
    $up = [math]::Max(1, [double]$script:Ctx.UptimeDays * 86400)
    if ($c.ContainsKey('Batch Requests/sec') -and $c.ContainsKey('SQL Compilations/sec')) {
        $bps = $c['Batch Requests/sec'] / $up; $cps = $c['SQL Compilations/sec'] / $up
        if ($bps -ge 10 -and $cps / $bps -ge 0.15) {
            Add-Finding 'Medium' ('Compilations are ' + (Format-N0 (100 * $cps / $bps)) + '% of batch requests') 'Most queries are compiled instead of reusing cached plans, which costs CPU. Typical cause: non-parameterized ad hoc SQL.' 'Parameterize queries in the application; enable optimize for ad hoc workloads; consider forced parameterization for the worst database.'
        }
    }
}

function Test-Memory {
    $sm = Get-One 'SELECT total_physical_memory_kb, available_physical_memory_kb, system_memory_state_desc FROM sys.dm_os_sys_memory'
    $pm = Get-One 'SELECT physical_memory_in_use_kb, locked_page_allocations_kb, process_physical_memory_low FROM sys.dm_os_process_memory'
    $c = $script:Ctx.Counters
    $ple = $null; if ($c.ContainsKey('Page life expectancy')) { $ple = $c['Page life expectancy'] }
    $pending = 0; if ($c.ContainsKey('Memory Grants Pending')) { $pending = $c['Memory Grants Pending'] }
    $totalKB = 0; if ($c.ContainsKey('Total Server Memory (KB)')) { $totalKB = $c['Total Server Memory (KB)'] }
    $targetKB = 0; if ($c.ContainsKey('Target Server Memory (KB)')) { $targetKB = $c['Target Server Memory (KB)'] }

    $availMB = [double]$sm.available_physical_memory_kb / 1024
    $totalOsMB = [double]$sm.total_physical_memory_kb / 1024
    $availPct = 0; if ($totalOsMB -gt 0) { $availPct = 100.0 * $availMB / $totalOsMB }
    # PLE threshold scaled to memory size: 300 s per 4 GB of SQL memory.
    $pleThreshold = [math]::Max(300, [math]::Round(($totalKB / 1024 / 1024) / 4 * 300))

    Add-Table 'Memory' @(
        [pscustomobject]@{ Metric = 'OS memory available';         Value = ((Format-N0 $availMB) + ' MB of ' + (Format-N0 $totalOsMB) + ' MB (' + (Format-N1 $availPct) + '%)') },
        [pscustomobject]@{ Metric = 'OS memory state';             Value = $sm.system_memory_state_desc },
        [pscustomobject]@{ Metric = 'SQL Server memory in use';    Value = ((Format-N0 ([double]$pm.physical_memory_in_use_kb / 1024)) + ' MB') },
        [pscustomobject]@{ Metric = 'SQL Server target / total';   Value = ((Format-N0 ($targetKB / 1024)) + ' MB / ' + (Format-N0 ($totalKB / 1024)) + ' MB') },
        [pscustomobject]@{ Metric = 'Locked pages in memory';      Value = $(if ([double]$pm.locked_page_allocations_kb -gt 0) { 'In use' } else { 'Not in use' }) },
        [pscustomobject]@{ Metric = 'Page life expectancy (now)';  Value = ((Format-N0 $ple) + ' s (guideline for this size: ' + (Format-N0 $pleThreshold) + ' s)') },
        [pscustomobject]@{ Metric = 'Memory grants pending (now)'; Value = $pending }
    ) @(2, 3)

    if ($pm.process_physical_memory_low -or $sm.system_memory_state_desc -like '*low*' -or ($availMB -lt 512 -and $totalOsMB -gt 0)) {
        Add-Finding 'High' 'The server is low on memory' ('Only ' + (Format-N0 $availMB) + ' MB is available to the OS (state: ' + $sm.system_memory_state_desc + ').') 'Check max server memory and other processes on the server; the OS may be paging.'
    }
    if ($null -ne $ple -and $ple -lt $pleThreshold -and [double]$script:Ctx.UptimeDays -ge 0.05) {
        Add-Finding 'Medium' ('Page life expectancy is low (' + (Format-N0 $ple) + ' s)') ('Data pages stay in memory for a short time (guideline for this memory size: ' + (Format-N0 $pleThreshold) + ' s), so SQL Server re-reads data from disk. This is a snapshot; check it at several times of day.') 'Reduce large scans (top reads queries, missing indexes) or add memory.'
    }
    if ($pending -gt 0) {
        Add-Finding 'Medium' ([string]$pending + ' queries are waiting for a memory grant right now') 'Queries cannot start until memory is available for their sorts and hash joins.' 'Find queries with large memory grants and tune them; check max server memory.'
    }
}

function Test-Io {
    $rows = @(Invoke-Q @'
SELECT DB_NAME(vfs.database_id) AS db, mf.name AS file_name, mf.type_desc,
       vfs.num_of_reads, vfs.num_of_writes, vfs.io_stall_read_ms, vfs.io_stall_write_ms, vfs.io_stall
FROM sys.dm_io_virtual_file_stats(NULL, NULL) AS vfs
JOIN sys.master_files AS mf ON mf.database_id = vfs.database_id AND mf.file_id = vfs.file_id
ORDER BY vfs.io_stall DESC
'@)
    $files = foreach ($r in $rows) {
        $reads = [double]$r.num_of_reads; $writes = [double]$r.num_of_writes
        [pscustomobject]@{
            Db = $r.db; File = $r.file_name; Type = $r.type_desc; Reads = $reads; Writes = $writes
            ReadMs  = $(if ($reads -gt 0) { [double]$r.io_stall_read_ms / $reads } else { $null })
            WriteMs = $(if ($writes -gt 0) { [double]$r.io_stall_write_ms / $writes } else { $null })
        }
    }
    $files = @($files)

    $table = foreach ($f in ($files | Select-Object -First $Top)) {
        [pscustomobject]@{
            'Database' = $f.Db; 'File' = $f.File; 'Type' = $f.Type
            'Reads' = [long]$f.Reads
            'Avg read (ms)' = $(if ($null -ne $f.ReadMs) { [math]::Round($f.ReadMs, 1) } else { $null })
            'Writes' = [long]$f.Writes
            'Avg write (ms)' = $(if ($null -ne $f.WriteMs) { [math]::Round($f.WriteMs, 1) } else { $null })
        }
    }
    Add-Table 'Files with the most I/O wait time (since startup)' $table @(2, 2, 1, 1.3, 1.2, 1.3, 1.2) 'Guidelines: data file reads under 20 ms, log file writes under 5 ms (under 2 ms on SSD).'

    $slowData = @($files | Where-Object { $_.Type -eq 'ROWS' -and $_.Reads -ge 1000 -and $_.ReadMs -ge 20 } | Sort-Object ReadMs -Descending)
    if ($slowData.Count -gt 0) {
        $worst = $slowData[0]
        $sev = 'Medium'; if ($worst.ReadMs -ge 50) { $sev = 'High' }
        $list = @($slowData | ForEach-Object { $_.Db + '/' + $_.File + ' ' + (Format-N0 $_.ReadMs) + ' ms' })
        Add-Finding $sev ('Slow data file reads on ' + $slowData.Count + ' file(s)') ('Average read latency: ' + (Join-Names $list 5) + '.') 'Check the storage (SAN/VM datastore latency, disk queue). Reducing scans through indexing also lowers read volume.'
    }
    $slowLog = @($files | Where-Object { $_.Type -eq 'LOG' -and $_.Writes -ge 1000 -and $_.WriteMs -ge 5 } | Sort-Object WriteMs -Descending)
    if ($slowLog.Count -gt 0) {
        $worst = $slowLog[0]
        $sev = 'Low'; if ($worst.WriteMs -ge 10) { $sev = 'Medium' }; if ($worst.WriteMs -ge 20) { $sev = 'High' }
        $list = @($slowLog | ForEach-Object { $_.Db + ' ' + (Format-N1 $_.WriteMs) + ' ms' })
        Add-Finding $sev ('Slow transaction log writes on ' + $slowLog.Count + ' database(s)') ('Every commit waits for the log write. Average write latency: ' + (Join-Names $list 5) + '.') 'Move busy log files to low-latency storage and keep them separate from data files.'
    }
}

function Test-TempDb {
    $files = @(Invoke-Q 'SELECT name, type_desc, CAST(size / 128.0 AS decimal(18,1)) AS size_mb, growth, is_percent_growth, physical_name FROM tempdb.sys.database_files')
    $data = @($files | Where-Object { $_.type_desc -eq 'ROWS' })
    $cpu = [math]::Max(1, [int]$script:Ctx.CpuCount)
    $rec = [math]::Min(8, $cpu)

    $table = foreach ($f in $files) {
        [pscustomobject]@{
            'File' = $f.name; 'Type' = $f.type_desc; 'Size (MB)' = $f.size_mb
            'Growth' = $(if ($f.is_percent_growth) { [string]$f.growth + ' %' } else { (Format-N0 ([double]$f.growth / 128)) + ' MB' })
            'Path' = $f.physical_name
        }
    }
    Add-Table 'TempDB files' $table @(1.5, 1, 1, 1, 4)

    if ($data.Count -lt $rec) {
        Add-Finding 'Medium' ('TempDB has ' + $data.Count + ' data file(s) for ' + $cpu + ' CPUs') 'Too few files cause allocation contention (PAGELATCH waits) when many sessions use temporary tables.' ('Use ' + $rec + ' equally sized data files (one per CPU up to 8).')
    }
    if (@($data | ForEach-Object { [double]$_.size_mb } | Select-Object -Unique).Count -gt 1) {
        Add-Finding 'Low' 'TempDB data files have different sizes' 'SQL Server favours the largest file, which defeats the purpose of multiple files.' 'Make all tempdb data files the same size and growth.'
    }
    if (@($files | Where-Object { $_.is_percent_growth }).Count -gt 0) {
        Add-Finding 'Low' 'TempDB uses percentage autogrowth' 'Growth becomes large and unpredictable as files grow.' 'Use a fixed growth increment (for example 256-1024 MB).'
    }
}

function Test-Databases {
    $dbs = @(Invoke-Q @'
SELECT d.name, d.recovery_model_desc, d.compatibility_level, d.page_verify_option_desc,
       d.is_auto_close_on, d.is_auto_shrink_on, d.is_auto_create_stats_on, d.is_auto_update_stats_on
FROM sys.databases AS d
WHERE d.database_id > 4 AND d.state_desc = 'ONLINE' AND d.source_database_id IS NULL
ORDER BY d.name
'@)
    if ($dbs.Count -eq 0) { Add-Note 'No online user databases.'; return }
    $issues = @{}
    foreach ($d in $dbs) { $issues[$d.name] = New-Object System.Collections.ArrayList }
    $shrink = @(); $close = @(); $verify = @(); $stats = @(); $compat = @()
    $maxCompat = 0; if ([int]$script:Ctx.Major -ge 11) { $maxCompat = [int]$script:Ctx.Major * 10 }

    foreach ($d in $dbs) {
        if ($d.is_auto_shrink_on) { $shrink += $d.name; [void]$issues[$d.name].Add('Auto-shrink on') }
        if ($d.is_auto_close_on) { $close += $d.name; [void]$issues[$d.name].Add('Auto-close on') }
        if ($d.page_verify_option_desc -ne 'CHECKSUM') { $verify += $d.name; [void]$issues[$d.name].Add('Page verify ' + $d.page_verify_option_desc) }
        if (-not $d.is_auto_create_stats_on -or -not $d.is_auto_update_stats_on) { $stats += $d.name; [void]$issues[$d.name].Add('Auto statistics off') }
        if ($maxCompat -gt 0 -and [int]$d.compatibility_level -lt ($maxCompat - 20)) { $compat += ($d.name + ' (' + $d.compatibility_level + ')'); [void]$issues[$d.name].Add('Old compatibility level ' + $d.compatibility_level) }
    }

    $files = @(Invoke-Q 'SELECT DB_NAME(database_id) AS db, name, growth, is_percent_growth, CAST(size / 128.0 AS decimal(18,1)) AS size_mb FROM sys.master_files WHERE database_id > 4')
    $pct = @(); $tiny = @()
    foreach ($f in $files) {
        if (-not $f.db -or -not $issues.ContainsKey($f.db)) { continue }
        if ($f.is_percent_growth -and [int]$f.growth -gt 0) {
            $pct += ($f.db + '/' + $f.name); [void]$issues[$f.db].Add('Percent growth: ' + $f.name)
        } elseif (-not $f.is_percent_growth -and [int]$f.growth -gt 0 -and [int]$f.growth -le 128) {
            $tiny += ($f.db + '/' + $f.name); [void]$issues[$f.db].Add('Growth 1 MB or less: ' + $f.name)
        }
    }

    $noLogBackup = @()
    try {
        $bk = @(Invoke-Q @'
SELECT d.name,
       (SELECT MAX(b.backup_finish_date) FROM msdb.dbo.backupset AS b WHERE b.database_name = d.name AND b.type = 'L') AS last_log
FROM sys.databases AS d
WHERE d.database_id > 4 AND d.state_desc = 'ONLINE' AND d.recovery_model_desc <> 'SIMPLE' AND d.source_database_id IS NULL
'@)
        foreach ($b in $bk) {
            if ($null -eq $b.last_log -or ([datetime]$b.last_log) -lt $script:Ctx.ServerTime.AddHours(-24)) {
                $noLogBackup += $b.name
                if ($issues.ContainsKey($b.name)) { [void]$issues[$b.name].Add('No log backup in 24 h') }
            }
        }
    } catch { Add-Note ('Backup history not checked: ' + (Get-ShortError $_)) }

    $manyVlf = @(); $worstVlf = 0
    try {
        $vl = @(Invoke-Q "SELECT d.name, COUNT(*) AS vlfs FROM sys.databases AS d CROSS APPLY sys.dm_db_log_info(d.database_id) AS li WHERE d.database_id > 4 AND d.state_desc = 'ONLINE' GROUP BY d.name HAVING COUNT(*) > 300")
        foreach ($v in $vl) {
            $manyVlf += ($v.name + ' (' + $v.vlfs + ')')
            if ([int]$v.vlfs -gt $worstVlf) { $worstVlf = [int]$v.vlfs }
            if ($issues.ContainsKey($v.name)) { [void]$issues[$v.name].Add('' + $v.vlfs + ' VLFs') }
        }
    } catch { }

    $table = foreach ($d in $dbs) {
        if ($issues[$d.name].Count -gt 0) {
            [pscustomobject]@{
                'Database' = $d.name; 'Recovery' = $d.recovery_model_desc; 'Compat' = $d.compatibility_level
                'Issues found' = ($issues[$d.name] -join '; ')
            }
        }
    }
    $table = @($table)
    $clean = $dbs.Count - $table.Count
    Add-Table 'Databases with configuration issues' $table @(2, 1.3, 0.8, 5) ([string]$clean + ' of ' + $dbs.Count + ' user databases have no issues and are not listed.')

    if ($shrink.Count) { Add-Finding 'High' ('Auto-shrink is on for ' + $shrink.Count + ' database(s)') ((Join-Names $shrink) + '. Shrink/grow cycles fragment indexes and burn CPU and I/O.') 'Turn AUTO_SHRINK off.' }
    if ($close.Count) { Add-Finding 'Medium' ('Auto-close is on for ' + $close.Count + ' database(s)') ((Join-Names $close) + '. The database is closed and reopened repeatedly, flushing its cache.') 'Turn AUTO_CLOSE off.' }
    if ($stats.Count) { Add-Finding 'Medium' ('Automatic statistics are off for ' + $stats.Count + ' database(s)') ((Join-Names $stats) + '. Outdated statistics lead to bad query plans.') 'Turn AUTO_CREATE_STATISTICS and AUTO_UPDATE_STATISTICS on, unless a vendor requires otherwise.' }
    if ($verify.Count) { Add-Finding 'Medium' ('Page verify is not CHECKSUM on ' + $verify.Count + ' database(s)') ((Join-Names $verify) + '. Corruption may go undetected.') 'Set PAGE_VERIFY CHECKSUM.' }
    if ($pct.Count) { Add-Finding 'Low' ('Percentage autogrowth on ' + $pct.Count + ' file(s)') ((Join-Names $pct) + '. Growth events get larger and slower as files grow.') 'Use fixed growth increments (e.g. 256-1024 MB for data, 256-512 MB for logs).' }
    if ($tiny.Count) { Add-Finding 'Medium' ('Autogrowth of 1 MB or less on ' + $tiny.Count + ' file(s)') ((Join-Names $tiny) + '. Files grow in thousands of tiny steps, each one stalling activity.') 'Pre-size files and set fixed growth increments of at least 64 MB.' }
    if ($noLogBackup.Count) { Add-Finding 'Medium' ([string]$noLogBackup.Count + ' database(s) in FULL recovery without a log backup in 24 hours') ((Join-Names $noLogBackup) + '. The log keeps growing, which slows backups, restores and startup. (If log backups run on another AG replica, ignore this.)') 'Schedule regular log backups, or switch to SIMPLE recovery if point-in-time restore is not needed.' }
    if ($manyVlf.Count) {
        $sev = 'Low'; if ($worstVlf -gt 1000) { $sev = 'Medium' }
        Add-Finding $sev ('High number of virtual log files in ' + $manyVlf.Count + ' database(s)') ((Join-Names $manyVlf) + '. Many VLFs slow down recovery, restores and log backups.') 'Shrink the log once during a quiet period and regrow it in large fixed steps.'
    }
    if ($compat.Count) { Add-Finding 'Low' ('Old compatibility level on ' + $compat.Count + ' database(s)') ((Join-Names $compat) + '. These databases do not benefit from newer optimizer features.') 'Test and raise the compatibility level (Query Store helps catch regressions).' }
}

function Test-TopQueries {
    $tot = Get-One 'SELECT SUM(total_worker_time) AS cpu_us FROM sys.dm_exec_query_stats'
    $totCpu = [math]::Max(1, [double]$tot.cpu_us)
    $sql = @'
WITH q AS (
    SELECT qs.sql_handle, qs.plan_handle, qs.statement_start_offset, qs.statement_end_offset,
           ROW_NUMBER() OVER (PARTITION BY qs.query_hash ORDER BY qs.total_worker_time DESC) AS rn,
           SUM(qs.execution_count) OVER (PARTITION BY qs.query_hash) AS execs,
           SUM(qs.total_worker_time) OVER (PARTITION BY qs.query_hash) AS cpu_us,
           SUM(qs.total_logical_reads) OVER (PARTITION BY qs.query_hash) AS reads,
           SUM(qs.total_elapsed_time) OVER (PARTITION BY qs.query_hash) AS dur_us
    FROM sys.dm_exec_query_stats AS qs
)
SELECT TOP ({TOP})
       DB_NAME(CAST(pa.value AS int)) AS db,
       q.execs, q.cpu_us, q.reads, q.dur_us,
       SUBSTRING(st.text, (q.statement_start_offset / 2) + 1,
                 ((CASE q.statement_end_offset WHEN -1 THEN DATALENGTH(st.text) ELSE q.statement_end_offset END
                   - q.statement_start_offset) / 2) + 1) AS query_text
FROM q
CROSS APPLY sys.dm_exec_sql_text(q.sql_handle) AS st
OUTER APPLY (SELECT TOP (1) value FROM sys.dm_exec_plan_attributes(q.plan_handle) WHERE attribute = N'dbid') AS pa
WHERE q.rn = 1
ORDER BY q.cpu_us DESC;
'@
    $rows = @(Invoke-Q ($sql.Replace('{TOP}', [string]$Top)))
    $table = foreach ($r in $rows) {
        $ex = [math]::Max(1, [double]$r.execs)
        [pscustomobject]@{
            'Database' = $r.db
            'Executions' = [long]$r.execs
            'Total CPU (s)' = [math]::Round([double]$r.cpu_us / 1000000, 1)
            'Avg CPU (ms)' = [math]::Round([double]$r.cpu_us / 1000 / $ex, 1)
            'Avg reads' = [long]([double]$r.reads / $ex)
            'Avg duration (ms)' = [math]::Round([double]$r.dur_us / 1000 / $ex, 1)
            '% of CPU' = [math]::Round(100 * [double]$r.cpu_us / $totCpu, 1)
            'Query text' = (Format-QueryText $r.query_text)
        }
    }
    $table = @($table)
    Add-Table 'Top statements by total CPU (similar statements grouped)' $table @(1.4, 1.1, 1, 1, 1.1, 1.1, 0.8, 6) 'Based on plans currently in cache. Queries recompiled or evicted from cache are not included.'

    if ($table.Count -gt 0) {
        $top1 = $table[0]
        if ($top1.'% of CPU' -ge 25) {
            Add-Finding 'Medium' ('One statement uses ' + (Format-N1 $top1.'% of CPU') + '% of all cached CPU') ('Database ' + $top1.Database + ', ' + (Format-N0 $top1.Executions) + ' executions, ' + (Format-N1 $top1.'Avg CPU (ms)') + ' ms CPU each. Tuning it would have a large effect.') 'Review its execution plan (look for scans, key lookups, implicit conversions) and indexing. It is the first row of the Top queries table.'
        }
        $top5 = 0.0; foreach ($t in ($table | Select-Object -First 5)) { $top5 += [double]$t.'% of CPU' }
        if ($top5 -ge 60 -and $top1.'% of CPU' -lt 25) {
            Add-Finding 'Info' ('The top 5 statements use ' + (Format-N0 $top5) + '% of cached CPU') 'The workload is concentrated: tuning a handful of queries would make a noticeable difference.' 'Start with the first rows of the Top queries table.'
        }
    }
}

function Test-MissingIndexes {
    $sql = @'
SELECT TOP ({TOP})
       DB_NAME(mid.database_id) AS [Database],
       mid.[statement] AS [Table],
       mid.equality_columns AS [Equality columns],
       mid.inequality_columns AS [Inequality columns],
       mid.included_columns AS [Included columns],
       migs.user_seeks + migs.user_scans AS [Uses],
       CAST(migs.avg_user_impact AS decimal(5,1)) AS [Est. gain %],
       CAST(migs.avg_total_user_cost * migs.avg_user_impact * (migs.user_seeks + migs.user_scans) AS decimal(18,0)) AS [Score]
FROM sys.dm_db_missing_index_group_stats AS migs
JOIN sys.dm_db_missing_index_groups AS mig ON mig.index_group_handle = migs.group_handle
JOIN sys.dm_db_missing_index_details AS mid ON mid.index_handle = mig.index_handle
WHERE mid.database_id > 4
ORDER BY [Score] DESC;
'@
    $rows = @(Invoke-Q ($sql.Replace('{TOP}', [string]$Top)))
    Add-Table 'Missing index suggestions (highest benefit first)' $rows @(1.3, 2.5, 2, 1.5, 2, 0.9, 0.9, 1.1) 'Suggestions come from the optimizer and are often overlapping or too wide. Review and combine them; never create them blindly.'
    $high = @($rows | Where-Object { [double]$_.Score -ge 100000 })
    if ($high.Count -gt 0) {
        $tables = @($high | ForEach-Object { $_.Table } | Select-Object -Unique)
        Add-Finding 'Medium' ([string]$high.Count + ' high-value missing index suggestion(s)') ('The optimizer repeatedly wanted indexes on: ' + (Join-Names $tables 4) + '.') 'Review the Missing indexes table, check existing indexes on those tables, and add consolidated indexes after testing.'
    }
}

function Test-Blocking {
    $blocked = @(Invoke-Q @'
SELECT r.session_id AS [Session], r.blocking_session_id AS [Blocked by], DB_NAME(r.database_id) AS [Database],
       r.wait_type AS [Wait type], CAST(r.wait_time / 1000.0 AS decimal(18,1)) AS [Waiting (s)],
       s.login_name AS [Login], LEFT(st.text, 300) AS [Statement]
FROM sys.dm_exec_requests AS r
JOIN sys.dm_exec_sessions AS s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) AS st
WHERE r.blocking_session_id <> 0
ORDER BY r.wait_time DESC
'@)
    $longTx = @(Invoke-Q @'
SELECT TOP (10) st.session_id AS [Session], s.login_name AS [Login], s.program_name AS [Program],
       s.status AS [Status], at.transaction_begin_time AS [Started],
       DATEDIFF(MINUTE, at.transaction_begin_time, GETDATE()) AS [Open (min)]
FROM sys.dm_tran_session_transactions AS st
JOIN sys.dm_tran_active_transactions AS at ON at.transaction_id = st.transaction_id
JOIN sys.dm_exec_sessions AS s ON s.session_id = st.session_id
WHERE s.is_user_process = 1 AND st.session_id <> @@SPID
  AND at.transaction_begin_time < DATEADD(MINUTE, -10, GETDATE())
ORDER BY at.transaction_begin_time
'@)
    foreach ($b in $blocked) { $b.Statement = Format-QueryText $b.Statement }
    if ($blocked.Count -gt 0) {
        Add-Table 'Requests blocked at the time of the audit' $blocked @(0.8, 0.8, 1.3, 1.3, 0.9, 1.5, 5)
        Add-Finding 'Medium' ([string]$blocked.Count + ' request(s) were blocked during the audit') ('Longest wait: ' + (Format-N1 $blocked[0].'Waiting (s)') + ' s, blocked by session ' + $blocked[0].'Blocked by' + '.') 'Identify the head blocker and what it is doing; look for long transactions and missing indexes on the tables involved.'
    } else {
        Add-Note 'No blocking at the time of the audit.'
    }
    if ($longTx.Count -gt 0) {
        Add-Table 'Transactions open for more than 10 minutes' $longTx @(0.8, 1.6, 2.2, 1, 1.5, 0.9)
        $sleeping = @($longTx | Where-Object { $_.Status -eq 'sleeping' })
        $detail = 'Oldest open for ' + (Format-N0 $longTx[0].'Open (min)') + ' minutes (session ' + $longTx[0].Session + ', ' + $longTx[0].Program + ').'
        if ($sleeping.Count -gt 0) { $detail += ' ' + $sleeping.Count + ' of them are idle (sleeping) with an open transaction, which usually means the application forgot to commit.' }
        Add-Finding 'Medium' ([string]$longTx.Count + ' long-running open transaction(s)') ($detail + ' Long transactions hold locks and prevent log reuse.') 'Check with the application owner; fix missing COMMIT/ROLLBACK handling.'
    }
    $c = $script:Ctx.Counters
    if ($c.ContainsKey('Number of Deadlocks/sec') -and [double]$script:Ctx.UptimeDays -gt 0) {
        $perDay = $c['Number of Deadlocks/sec'] / [math]::Max(1, [double]$script:Ctx.UptimeDays)
        if ($perDay -ge 10) {
            Add-Finding 'Medium' ('About ' + (Format-N0 $perDay) + ' deadlocks per day') ([string](Format-N0 $c['Number of Deadlocks/sec']) + ' deadlocks since startup. Each one kills a transaction that the application must retry.') 'Capture deadlock graphs from the system_health Extended Events session and fix the access order or indexing.'
        } elseif ($perDay -ge 1) {
            Add-Finding 'Low' ('About ' + (Format-N1 $perDay) + ' deadlocks per day') ([string](Format-N0 $c['Number of Deadlocks/sec']) + ' deadlocks since startup.') 'Review deadlock graphs in the system_health Extended Events session.'
        }
    }
}

# =============================================================================
# Minimal .docx writer (Office Open XML, no dependencies)
# =============================================================================

$script:SevStyle = @{
    High   = @{ Fg = 'B42318'; Bg = 'FDE2E1' }
    Medium = @{ Fg = '9A5B00'; Bg = 'FDEBC8' }
    Low    = @{ Fg = '1F5F99'; Bg = 'DCEBFA' }
    Info   = @{ Fg = '4B5563'; Bg = 'ECEEF1' }
}
$script:SevRank = @{ High = 0; Medium = 1; Low = 2; Info = 3 }
$script:ContentWidth = 9866   # A4 width minus 1.8 cm margins, in twips

function ConvertTo-XmlText([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', ' ')
    return [System.Security.SecurityElement]::Escape($s)
}

function New-WRun([string]$Text, [switch]$Bold, [switch]$Italic, [string]$Color, [int]$Size = 0) {
    $rpr = ''
    if ($Bold) { $rpr += '<w:b/>' }
    if ($Italic) { $rpr += '<w:i/>' }
    if ($Color) { $rpr += '<w:color w:val="' + $Color + '"/>' }
    if ($Size -gt 0) { $rpr += '<w:sz w:val="' + $Size + '"/><w:szCs w:val="' + $Size + '"/>' }
    if ($rpr) { $rpr = '<w:rPr>' + $rpr + '</w:rPr>' }
    return '<w:r>' + $rpr + '<w:t xml:space="preserve">' + (ConvertTo-XmlText $Text) + '</w:t></w:r>'
}

function New-WPara([string]$Runs, [string]$Style, [int]$After = -1, [switch]$PageBreak, [int]$Indent = 0, [string]$Shade) {
    $ppr = ''
    if ($Style) { $ppr += '<w:pStyle w:val="' + $Style + '"/>' }
    if ($PageBreak) { $ppr += '<w:pageBreakBefore/>' }
    if ($Shade) { $ppr += '<w:shd w:val="clear" w:color="auto" w:fill="' + $Shade + '"/>' }
    if ($After -ge 0) { $ppr += '<w:spacing w:after="' + $After + '"/>' }
    if ($Indent -gt 0) { $ppr += '<w:ind w:left="' + $Indent + '"/>' }
    if ($ppr) { $ppr = '<w:pPr>' + $ppr + '</w:pPr>' }
    return '<w:p>' + $ppr + $Runs + '</w:p>'
}

function Add-WXml([string]$Xml) { [void]$script:Body.Append($Xml) }

function New-WCell([int]$Width, $Content, [string]$Fill, [switch]$Bold, [string]$Color, [string]$Align) {
    $x = '<w:tc><w:tcPr><w:tcW w:w="' + $Width + '" w:type="dxa"/>'
    if ($Fill) { $x += '<w:shd w:val="clear" w:color="auto" w:fill="' + $Fill + '"/>' }
    $x += '</w:tcPr>'
    if ($Content -is [hashtable] -and $Content.ContainsKey('Xml')) {
        $x += $Content.Xml
    } else {
        $ppr = '<w:pPr><w:spacing w:before="0" w:after="0"/>'
        if ($Align) { $ppr += '<w:jc w:val="' + $Align + '"/>' }
        $ppr += '</w:pPr>'
        $x += '<w:p>' + $ppr + (New-WRun (Format-Value $Content) -Bold:$Bold -Color $Color -Size 16) + '</w:p>'
    }
    return $x + '</w:tc>'
}

function Add-WTable($Rows, [double[]]$Weights, [string]$SeverityColumn) {
    $rows = @($Rows)
    if ($rows.Count -eq 0) { return }
    $cols = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    $n = $cols.Count
    if ($null -eq $Weights -or $Weights.Count -ne $n) { $Weights = @(1..$n | ForEach-Object { 1.0 }) }
    $sum = 0.0; foreach ($w in $Weights) { $sum += $w }
    $widths = @($Weights | ForEach-Object { [int][math]::Floor($script:ContentWidth * $_ / $sum) })
    $total = 0; foreach ($w in $widths) { $total += $w }

    $x = New-Object System.Text.StringBuilder
    [void]$x.Append('<w:tbl><w:tblPr><w:tblW w:w="' + $total + '" w:type="dxa"/><w:tblBorders>')
    foreach ($e in @('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) {
        [void]$x.Append('<w:' + $e + ' w:val="single" w:sz="4" w:space="0" w:color="C9D1DB"/>')
    }
    [void]$x.Append('</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>')
    foreach ($w in $widths) { [void]$x.Append('<w:gridCol w:w="' + $w + '"/>') }
    [void]$x.Append('</w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>')
    for ($i = 0; $i -lt $n; $i++) { [void]$x.Append((New-WCell $widths[$i] ([string]$cols[$i]) -Fill '1F3A5F' -Bold -Color 'FFFFFF')) }
    [void]$x.Append('</w:tr>')
    $ri = 0
    foreach ($r in $rows) {
        [void]$x.Append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        for ($i = 0; $i -lt $n; $i++) {
            $v = $r.($cols[$i])
            $fill = $null; if ($ri % 2 -eq 1) { $fill = 'F5F7FA' }
            if ($SeverityColumn -and $cols[$i] -eq $SeverityColumn -and $script:SevStyle.ContainsKey([string]$v)) {
                $st = $script:SevStyle[[string]$v]
                [void]$x.Append((New-WCell $widths[$i] $v -Fill $st.Bg -Bold -Color $st.Fg))
            } elseif (Test-Numeric $v) {
                [void]$x.Append((New-WCell $widths[$i] $v -Fill $fill -Align 'right'))
            } else {
                [void]$x.Append((New-WCell $widths[$i] $v -Fill $fill))
            }
        }
        [void]$x.Append('</w:tr>')
        $ri++
    }
    [void]$x.Append('</w:tbl>')
    Add-WXml $x.ToString()
    Add-WXml (New-WPara '' -After 120)
}

function Add-WFinding($F) {
    $st = $script:SevStyle[$F.Severity]
    $runs = (New-WRun ('[' + $F.Severity + ']  ') -Bold -Color $st.Fg) + (New-WRun $F.Title -Bold)
    Add-WXml (New-WPara $runs -After 20 -Indent 120)
    if ($F.Detail) { Add-WXml (New-WPara (New-WRun $F.Detail) -After 20 -Indent 360) }
    if ($F.Fix) { Add-WXml (New-WPara ((New-WRun 'What to do: ' -Bold -Color '1F3A5F') + (New-WRun $F.Fix)) -After 140 -Indent 360) }
}

function Save-Docx([string]$Path, [string]$Title) {
    $ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

    $contentTypes = $hdr + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '</Types>'

    $rootRels = $hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '</Relationships>'

    $docRels = $hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
        '</Relationships>'

    $styles = $hdr + '<w:styles ' + $ns + '>' +
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="60"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="1F3A5F"/></w:pBdr><w:spacing w:before="360" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2E5C8A"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="6B7280"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>' +
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
        '</w:styles>'

    $footer = $hdr + '<w:ftr ' + $ns + '><w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="right"/></w:pPr>' +
        (New-WRun ($Title + '   |   Page ')) +
        '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
        (New-WRun ' of ') +
        '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>'

    $created = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', $script:Inv)
    $core = $hdr + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>' + (ConvertTo-XmlText $Title) + '</dc:title><dc:creator>Invoke-SqlPerfAudit.ps1</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + $created + '</dcterms:created></cp:coreProperties>'

    $document = $hdr + '<w:document ' + $ns + '><w:body>' + $script:Body.ToString() +
        '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1020" w:right="1020" w:bottom="1020" w:left="1020" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr>' +
        '</w:body></w:document>'

    $parts = [ordered]@{
        '[Content_Types].xml'          = $contentTypes
        '_rels/.rels'                  = $rootRels
        'word/document.xml'            = $document
        'word/styles.xml'              = $styles
        'word/footer1.xml'             = $footer
        'word/_rels/document.xml.rels' = $docRels
        'docProps/core.xml'            = $core
    }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $parts.Keys) {
                $entry = $zip.CreateEntry($name)
                $st = $entry.Open()
                try { $bytes = $utf8.GetBytes([string]$parts[$name]); $st.Write($bytes, 0, $bytes.Length) } finally { $st.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

# =============================================================================
# Report layout
# =============================================================================

function Write-Report([string]$Path, [datetime]$Started) {
    $script:Body = New-Object System.Text.StringBuilder
    $ctx = $script:Ctx
    $server = $ctx.Server; if (-not $server) { $server = $ctx.Instance }
    $findings = @($script:Findings | Sort-Object -Property @{ Expression = { $script:SevRank[$_.Severity] } }, @{ Expression = { $script:SectionOrder[$_.Area] } })
    $cnt = @{ High = 0; Medium = 0; Low = 0; Info = 0 }
    foreach ($f in $findings) { $cnt[$f.Severity]++ }

    # --- Title block
    Add-WXml (New-WPara (New-WRun 'SQL Server Performance Audit') -Style 'Title')
    Add-WXml (New-WPara (New-WRun $server -Bold -Color '2E5C8A' -Size 28) -After 200)
    $facts = @(
        [pscustomobject]@{ Item = 'Version';       Value = $ctx.Version },
        [pscustomobject]@{ Item = 'Edition';       Value = $ctx.Edition },
        [pscustomobject]@{ Item = 'Audit date';    Value = $Started.ToString('yyyy-MM-dd HH:mm', $script:Inv) },
        [pscustomobject]@{ Item = 'Running since'; Value = $(if ($ctx.StartTime) { (Format-Value $ctx.StartTime) + ' (' + (Format-N1 $ctx.UptimeDays) + ' days)' } else { '' }) },
        [pscustomobject]@{ Item = 'Audited as';    Value = $ctx.AuditLogin }
    )
    Add-WTable $facts @(1, 3)

    # --- Summary
    Add-WXml (New-WPara (New-WRun 'Summary') -Style 'Heading1')
    if ($cnt.High -gt 0) {
        $verdict = 'There are ' + $cnt.High + ' high-severity problem(s) that deserve attention soon.'
    } elseif ($cnt.Medium -gt 0) {
        $verdict = 'No critical problems were found, but ' + $cnt.Medium + ' medium-severity issue(s) are worth looking at.'
    } elseif ($cnt.Low -gt 0) {
        $verdict = 'The instance looks healthy; only minor improvements were found.'
    } else {
        $verdict = 'No problems were found above the thresholds used by this audit.'
    }
    Add-WXml (New-WPara (New-WRun $verdict -Bold) -After 120)
    $summary = @(
        [pscustomobject]@{ 'Severity' = 'High';   'Count' = $cnt.High;   'Meaning' = 'Likely hurting performance or stability now - act soon.' },
        [pscustomobject]@{ 'Severity' = 'Medium'; 'Count' = $cnt.Medium; 'Meaning' = 'A real problem or risk worth fixing.' },
        [pscustomobject]@{ 'Severity' = 'Low';    'Count' = $cnt.Low;    'Meaning' = 'Best-practice improvement.' },
        [pscustomobject]@{ 'Severity' = 'Info';   'Count' = $cnt.Info;   'Meaning' = 'Context, no action required.' }
    )
    Add-WTable $summary @(1, 0.7, 6) -SeverityColumn 'Severity'

    if ($findings.Count -gt 0) {
        Add-WXml (New-WPara (New-WRun 'All findings, most important first') -Style 'Heading2')
        $i = 0
        $rows = foreach ($f in $findings) {
            $i++
            $cell = @{ Xml = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>' + (New-WRun $f.Title -Bold -Size 16) + '</w:p>' +
                              '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>' + (New-WRun $f.Detail -Size 16) + '</w:p>') }
            [pscustomobject]@{ '#' = $i; 'Severity' = $f.Severity; 'Area' = $f.Area; 'Finding' = $cell; 'What to do' = $f.Fix }
        }
        Add-WTable $rows @(0.4, 0.9, 1.3, 4.5, 3.4) -SeverityColumn 'Severity'
    }

    # --- Detail sections
    $n = 0
    foreach ($s in $script:Sections) {
        $n++
        Add-WXml (New-WPara (New-WRun ([string]$n + '. ' + $s.Title)) -Style 'Heading1')
        if ($s.Intro) { Add-WXml (New-WPara (New-WRun $s.Intro -Italic -Color '4B5563') -After 120) }
        if ($s.Error) {
            Add-WXml (New-WPara (New-WRun ('This check could not run: ' + $s.Error) -Color '9A5B00') -Shade 'FDEBC8' -After 120)
        }
        $secFindings = @($findings | Where-Object { $_.Area -eq $s.Title })
        foreach ($f in $secFindings) { Add-WFinding $f }
        if (-not $s.Error -and $secFindings.Count -eq 0) {
            Add-WXml (New-WPara (New-WRun 'No problems found in this area.' -Color '2E7D32') -After 120)
        }
        foreach ($note in $s.Notes) { Add-WXml (New-WPara (New-WRun $note -Italic -Color '6B7280' -Size 18) -After 100) }
        foreach ($t in $s.Tables) {
            Add-WXml (New-WPara (New-WRun $t.Caption) -Style 'Heading2')
            if ($t.Rows.Count -eq 0) {
                Add-WXml (New-WPara (New-WRun 'None.' -Italic -Color '6B7280') -After 120)
            } else {
                Add-WTable $t.Rows $t.Weights
            }
            if ($t.Note) { Add-WXml (New-WPara (New-WRun $t.Note -Italic -Color '6B7280' -Size 16) -After 160) }
        }
    }

    # --- About
    Add-WXml (New-WPara (New-WRun 'About this report') -Style 'Heading1')
    $about = @(
        'This audit is read-only: it only queries system views and changes nothing on the server.',
        'Most figures (waits, I/O, query statistics, deadlocks) accumulate since the last restart. Values such as page life expectancy, blocking and open transactions are a snapshot taken at the time of the audit.',
        'Thresholds are common industry guidelines, not hard rules. A finding means "worth looking at", and the context (application, time of day, maintenance jobs) decides whether action is needed. Test changes before applying them in production.',
        ('Generated by Invoke-SqlPerfAudit.ps1 version ' + $script:ScriptVersion + ' in ' + (Format-N0 ((Get-Date) - $Started).TotalSeconds) + ' seconds.')
    )
    foreach ($a in $about) { Add-WXml (New-WPara (New-WRun $a -Size 18) -After 80) }

    Save-Docx $Path ('SQL Server Performance Audit - ' + $server)
}

# =============================================================================
# Main
# =============================================================================

function Invoke-InstanceAudit([string]$Instance) {
    $started = Get-Date
    $script:Ctx = @{ Instance = $Instance; Counters = @{} }
    $script:Findings = New-Object System.Collections.ArrayList
    $script:Sections = New-Object System.Collections.ArrayList

    Write-Host ('Auditing ' + $Instance + ' ...') -ForegroundColor Cyan
    $script:Conn = Connect-Instance $Instance
    try {
        [void](Invoke-Q 'SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SET LOCK_TIMEOUT 10000; SET DEADLOCK_PRIORITY LOW;')
        try { $script:Ctx.Counters = Get-PerfCounters } catch { Write-Warning ('Performance counters unavailable: ' + (Get-ShortError $_)) }

        Invoke-Check 'Instance' 'Version, hardware seen by SQL Server, uptime and service settings.' { Test-Instance }
        Invoke-Check 'Server configuration' 'Server-wide settings that most often cause performance problems when left at defaults.' { Test-Configuration }
        Invoke-Check 'Wait statistics' 'What SQL Server spends its time waiting on since the last restart. The top waits point to the main bottleneck.' { Test-Waits }
        Invoke-Check 'CPU' 'CPU usage over roughly the last four hours, from the SQL Server ring buffer.' { Test-Cpu }
        Invoke-Check 'Memory' 'Memory available to the OS and to SQL Server, and signs of memory pressure.' { Test-Memory }
        Invoke-Check 'Storage latency' 'Average read and write latency per database file since the last restart.' { Test-Io }
        Invoke-Check 'TempDB' 'TempDB file layout, a frequent source of contention.' { Test-TempDb }
        Invoke-Check 'Databases' 'Database options, file growth settings, log backups and virtual log files.' { Test-Databases }
        Invoke-Check 'Top queries' 'The statements that consumed the most CPU, from the plan cache.' { Test-TopQueries }
        Invoke-Check 'Missing indexes' 'Indexes the query optimizer reported it would have used.' { Test-MissingIndexes }
        Invoke-Check 'Blocking and transactions' 'Blocking, long open transactions and deadlocks.' { Test-Blocking }
    } finally {
        $script:Conn.Close()
        $script:Conn.Dispose()
    }

    $script:SectionOrder = @{}
    $k = 0; foreach ($s in $script:Sections) { $script:SectionOrder[$s.Title] = $k; $k++ }

    if (-not (Test-Path -LiteralPath $OutputFolder)) { [void](New-Item -ItemType Directory -Path $OutputFolder) }
    $folder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath
    $safe = $Instance -replace '[\\/:*?"<>|,]', '_'
    $file = Join-Path $folder ('SQLPerfAudit_' + $safe + '_' + $started.ToString('yyyyMMdd_HHmm') + '.docx')
    Write-Report $file $started

    $cnt = @{ High = 0; Medium = 0; Low = 0; Info = 0 }
    foreach ($f in $script:Findings) { $cnt[$f.Severity]++ }
    Write-Host ('  Report: ' + $file) -ForegroundColor Green
    Write-Host ('  Findings: ' + $cnt.High + ' high, ' + $cnt.Medium + ' medium, ' + $cnt.Low + ' low, ' + $cnt.Info + ' info') -ForegroundColor Green
    [pscustomobject]@{ Instance = $Instance; Report = $file; High = $cnt.High; Medium = $cnt.Medium; Low = $cnt.Low; Info = $cnt.Info }
}

foreach ($inst in $SqlInstance) {
    try {
        Invoke-InstanceAudit $inst
    } catch {
        Write-Warning ('Audit of ' + $inst + ' failed: ' + (Get-ShortError $_))
    }
}
