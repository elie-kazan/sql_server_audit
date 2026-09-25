<#
.SYNOPSIS
    Audit de performance SQL Server en lecture seule, avec rapport Word (.docx) en français.

.DESCRIPTION
    Se connecte à une ou plusieurs instances SQL Server, exécute des requêtes de diagnostic
    en lecture seule (DMV et catalogues système) et produit un rapport Word par instance,
    listant les problèmes qui méritent attention, classés par gravité, avec une
    recommandation pour chacun.

    Aucune installation : utilise uniquement des classes .NET fournies avec Windows
    PowerShell 5.1 (fonctionne aussi avec PowerShell 7). Ni Office, ni Python, ni module.
    Rien n'est modifié sur le serveur audité.

    Droits nécessaires : VIEW SERVER STATE et VIEW ANY DEFINITION (sysadmin inutile),
    plus la lecture de l'historique de sauvegarde (msdb.dbo.backupset).

    IMPORTANT : ce fichier doit rester enregistré en UTF-8 avec BOM, sinon Windows
    PowerShell 5.1 corrompt les accents du rapport.

.PARAMETER SqlInstance
    Une ou plusieurs instances : SERVEUR, SERVEUR\INSTANCE ou SERVEUR,PORT.

.PARAMETER Credential
    Login SQL (via Get-Credential). Sans ce paramètre, l'authentification Windows est utilisée.

.PARAMETER OutputFolder
    Dossier des rapports. Par défaut : le dossier courant.

.PARAMETER Top
    Nombre de lignes des tableaux « top » (attentes, requêtes, fichiers, index). Défaut : 10.

.EXAMPLE
    .\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01

.EXAMPLE
    .\Invoke-SqlPerfAudit.ps1 -SqlInstance 'SQLPROD01\INST1','SQLPROD02' -OutputFolder C:\Audits

.EXAMPLE
    .\Invoke-SqlPerfAudit.ps1 -SqlInstance SQLPROD01 -Credential (Get-Credential) -TrustServerCertificate

.NOTES
    Si la stratégie d'exécution bloque le script :
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
$script:Fr = [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')
$script:ScriptVersion = '1.1'

# =============================================================================
# Fonctions utilitaires
# =============================================================================

function Get-ShortError($Err) {
    $e = $Err
    if ($e -is [System.Management.Automation.ErrorRecord]) { $e = $e.Exception }
    while ($null -ne $e.InnerException) { $e = $e.InnerException }
    $m = (([string]$e.Message) -replace '\s+', ' ').Trim()
    if ($m.Length -gt 300) { $m = $m.Substring(0, 297) + '...' }
    return $m
}

function Format-N0($v) { if ($null -eq $v) { return '' }; return ([double]$v).ToString('N0', $script:Fr) }
function Format-N1($v) { if ($null -eq $v) { return '' }; return ([double]$v).ToString('N1', $script:Fr) }
function Format-Date($v) { if ($null -eq $v) { return '' }; return ([datetime]$v).ToString('dd/MM/yyyy HH:mm', $script:Fr) }

function Format-Value($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [bool]) { if ($v) { return 'Oui' } else { return 'Non' } }
    if ($v -is [datetime]) { return (Format-Date $v) }
    if ($v -is [int] -or $v -is [long] -or $v -is [int16] -or $v -is [byte]) { return ([long]$v).ToString('N0', $script:Fr) }
    if ($v -is [decimal] -or $v -is [double] -or $v -is [single]) {
        if ([math]::Abs([double]$v) -ge 1000) { return ([double]$v).ToString('N0', $script:Fr) }
        return ([double]$v).ToString('N1', $script:Fr)
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
    return (($n[0..($Max - 1)] -join ', ') + ' et ' + ($n.Count - $Max) + ' autre(s)')
}

function Format-QueryText([string]$Text) {
    if (-not $Text) { return '' }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt 300) { $t = $t.Substring(0, 297) + '...' }
    return $t
}

# =============================================================================
# Accès SQL Server
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

# Exécute une requête et renvoie les lignes sous forme d'objets (DBNull converti en $null).
# Toujours appeler sous la forme @(Invoke-Q ...) pour obtenir un tableau même avec une seule ligne.
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
# Constats, sections, tableaux
# =============================================================================

# Gravité interne (High/Medium/Low/Info) et libellé affiché en français.
$script:SevLabel = @{ High = 'Élevée'; Medium = 'Moyenne'; Low = 'Faible'; Info = 'Info' }

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
        Write-Warning ('    ' + $Title + " n'a pas pu s'exécuter : " + $sec.Error)
    }
    [void]$script:Sections.Add($sec)
}

# =============================================================================
# Données de référence
# =============================================================================

# Attentes correspondant à une activité de fond normale : ignorées.
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

# Signification des attentes courantes (recherche par préfixe, la première correspondance l'emporte).
$script:WaitInfo = @(
    @{ P = 'THREADPOOL';          M = "Pénurie de threads de travail : des requêtes ont attendu un thread. Peut provoquer des délais d'attente et un serveur qui semble figé."; F = "Trouver la cause (le plus souvent des blocages importants ou un parallélisme excessif). Ne pas se contenter d'augmenter max worker threads." },
    @{ P = 'RESOURCE_SEMAPHORE';  M = "Des requêtes ont attendu de la mémoire (tris, jointures de hachage) avant de pouvoir démarrer."; F = "Optimiser les requêtes qui demandent beaucoup de mémoire (index manquants, mauvaises estimations) et vérifier la configuration mémoire." },
    @{ P = 'CX';                  M = "Parallélisme : les threads d'une requête parallèle s'attendent les uns les autres."; F = "Vérifier MAXDOP et cost threshold for parallelism, puis optimiser les plus grosses requêtes parallèles." },
    @{ P = 'SOS_SCHEDULER_YIELD'; M = "Pression CPU : des threads ont épuisé leur quantum de temps processeur, souvent en parcourant des données en mémoire."; F = "Optimiser les requêtes les plus consommatrices de CPU (voir Requêtes les plus coûteuses) et chercher les parcours qu'un index pourrait éviter." },
    @{ P = 'PAGEIOLATCH_';        M = "Lecture de pages de données depuis le disque vers la mémoire."; F = "Vérifier la latence de lecture des disques et la pression mémoire ; réduire le volume lu grâce à l'indexation et à l'optimisation des requêtes." },
    @{ P = 'WRITELOG';            M = "Validations de transactions en attente de l'écriture du journal sur disque."; F = "Vérifier la latence d'écriture des fichiers journaux ; les placer sur un stockage rapide ; éviter les très nombreuses petites transactions." },
    @{ P = 'LCK_M_';              M = "Blocages : des sessions attendent des verrous détenus par d'autres sessions."; F = "Identifier les chaînes de blocage et les transactions longues ; envisager READ_COMMITTED_SNAPSHOT ; indexer pour raccourcir la durée des verrous." },
    @{ P = 'PAGELATCH_';          M = "Contention sur des pages très sollicitées en mémoire, souvent les pages d'allocation de tempdb."; F = "Vérifier le nombre de fichiers de tempdb ; envisager OPTIMIZE_FOR_SEQUENTIAL_KEY pour les tables à insertions intensives." },
    @{ P = 'ASYNC_NETWORK_IO';    M = "SQL Server a attendu que l'application cliente lise les résultats."; F = "Généralement un problème applicatif (gros jeux de résultats lus ligne à ligne, serveur d'application lent)." },
    @{ P = 'HADR_SYNC_COMMIT';    M = "Validations en attente des réplicas synchrones du groupe de disponibilité."; F = "Vérifier la latence réseau et disque des réplicas secondaires." },
    @{ P = 'IO_COMPLETION';       M = "E/S hors pages de données, par exemple des débordements de tris ou de hachages dans tempdb."; F = "Rechercher les débordements vers tempdb et vérifier la latence du stockage." },
    @{ P = 'ASYNC_IO_COMPLETION'; M = "Sauvegardes, croissance de fichiers et E/S en masse."; F = "Vérifier le planning des sauvegardes ; activer l'initialisation instantanée des fichiers ; pré-dimensionner les fichiers." },
    @{ P = 'BACKUP';              M = "Activité de sauvegarde."; F = "Normal pendant les sauvegardes ; les planifier en dehors des heures de pointe." },
    @{ P = 'OLEDB';               M = "Appels à des serveurs liés, ou certaines activités DBCC ou de supervision."; F = "Examiner les requêtes qui passent par des serveurs liés." },
    @{ P = 'PREEMPTIVE_';         M = "Appels au système d'exploitation (authentification, opérations sur fichiers, CLR)."; F = "Identifier l'appel concerné ; souvent une latence Active Directory ou du système de fichiers." }
)

function Get-WaitInfo([string]$WaitType) {
    foreach ($w in $script:WaitInfo) { if ($WaitType.StartsWith($w.P)) { return $w } }
    return @{ P = ''; M = "Type d'attente moins courant ; voir la documentation Microsoft."; F = "À examiner s'il reste en tête du classement." }
}

function Test-BenignWait([string]$WaitType) {
    if ($script:BenignWaits -contains $WaitType) { return $true }
    foreach ($p in $script:BenignPrefixes) { if ($WaitType.StartsWith($p)) { return $true } }
    return $false
}

function Get-RecommendedMaxMemoryMB([double]$PhysMB) {
    # Point de départ courant : 1 Go pour l'OS, +1 Go par 4 Go entre 4 et 16 Go, +1 Go par 8 Go au-delà de 16 Go.
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
# Vérifications
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

    $uptimeText = ([int][math]::Floor($uptime.TotalDays)).ToString() + ' jours ' + $uptime.Hours + ' h'
    $rows = @(
        [pscustomobject]@{ 'Propriété' = 'Serveur';                       'Valeur' = $p.ServerName },
        [pscustomobject]@{ 'Propriété' = 'Version';                       'Valeur' = $script:Ctx.Version },
        [pscustomobject]@{ 'Propriété' = 'Édition';                       'Valeur' = $p.Edition },
        [pscustomobject]@{ 'Propriété' = 'Processeurs utilisés par SQL Server'; 'Valeur' = ([string]$online + ' (' + $nodes + ' nœud(s) NUMA)') },
        [pscustomobject]@{ 'Propriété' = 'Mémoire physique';              'Valeur' = ((Format-N0 $physMB) + ' Mo') },
        [pscustomobject]@{ 'Propriété' = 'Machine virtuelle';             'Valeur' = [string]$si.virtual_machine_type_desc },
        [pscustomobject]@{ 'Propriété' = 'Démarré le';                    'Valeur' = $si.sqlserver_start_time },
        [pscustomobject]@{ 'Propriété' = 'Durée de fonctionnement';       'Valeur' = $uptimeText },
        [pscustomobject]@{ 'Propriété' = 'Initialisation instantanée des fichiers'; 'Valeur' = $(if ($ifi -eq 'Y') { 'Activée' } elseif ($ifi -eq 'N') { 'Désactivée' } else { 'Inconnue' }) }
    )
    Add-Table "Résumé de l'instance" $rows @(1, 2)

    if ($major -ge 10 -and $major -le 13) {
        Add-Finding 'Medium' ($verName + " n'est plus couvert par le support étendu de Microsoft") "Aucun correctif de sécurité ou de performance n'est plus publié pour cette version (le support de SQL Server 2016 a pris fin en juillet 2026)." "Planifier une migration vers une version supportée."
    } elseif ($major -eq 14) {
        Add-Finding 'Info' "Le support étendu de SQL Server 2017 prend fin en octobre 2027" "Après cette date, plus aucun correctif ne sera publié." "Commencer à planifier la migration."
    }
    if ($offline -gt 0) {
        Add-Finding 'High' ([string]$offline + ' processeur(s) inutilisable(s) par SQL Server') "Certains ordonnanceurs sont VISIBLE OFFLINE, généralement à cause d'une limite de licence de l'édition ou de la topologie sockets/cœurs de la machine virtuelle." "Vérifier la limite de processeurs de l'édition et reconfigurer la VM avec moins de sockets et plus de cœurs par socket, ou changer d'édition."
    }
    if ($ifi -eq 'N') {
        Add-Finding 'Medium' "L'initialisation instantanée des fichiers est désactivée" "Chaque croissance de fichier de données et chaque restauration doivent remettre l'espace à zéro, ce qui bloque l'activité pendant les croissances automatiques." "Accorder au compte de service SQL Server le droit « Effectuer les tâches de maintenance de volume », puis redémarrer le service."
    }
    if ($uptime.TotalDays -lt 7) {
        Add-Finding 'Info' ('SQL Server a redémarré il y a ' + (Format-N1 $uptime.TotalDays) + ' jours') "La plupart des statistiques de ce rapport sont cumulées depuis le dernier redémarrage : elles peuvent ne pas refléter une semaine type." "Relancer l'audit après un cycle d'activité complet pour une image plus fiable."
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
    if ($phys -gt 0) {
        $rec = Get-RecommendedMaxMemoryMB $phys
        if ($maxMem -ge 2147483647) {
            $assess['max server memory (MB)'] = 'Non défini (illimité)'
            Add-Finding 'High' "La mémoire maximale du serveur (max server memory) n'est pas configurée" ('Elle est illimitée sur un serveur de ' + (Format-N0 $phys) + " Mo de RAM : SQL Server peut priver le système d'exploitation de mémoire, ce qui provoque de la pagination et de l'instabilité.") ('Définir max server memory à environ ' + (Format-N0 $rec) + " Mo comme point de départ, puis surveiller la mémoire libre du système.")
        } elseif ($maxMem -gt ($phys - ($phys - $rec) / 2)) {
            $assess['max server memory (MB)'] = 'Laisse peu de mémoire au système'
            Add-Finding 'Medium' "La mémoire maximale du serveur laisse peu de place au système d'exploitation" ('Elle est de ' + (Format-N0 $maxMem) + ' Mo sur ' + (Format-N0 $phys) + ' Mo de RAM.') ('La réduire à environ ' + (Format-N0 $rec) + " Mo, sauf si vous avez vérifié que le système garde assez de mémoire libre.")
        } else {
            $assess['max server memory (MB)'] = 'OK'
        }
    }

    $maxdop = [long]$cfg['max degree of parallelism'].value_in_use
    $cpu = [int]$script:Ctx.CpuCount
    $recDop = Get-RecommendedMaxDop $script:Ctx.PerNode $script:Ctx.Nodes
    if ($maxdop -eq 0 -and $cpu -gt 8) {
        $assess['max degree of parallelism'] = ('Illimité sur ' + $cpu + ' processeurs ; suggéré : ' + $recDop)
        Add-Finding 'Medium' 'MAXDOP est illimité' ('Avec ' + $cpu + ' processeurs logiques, une seule requête peut tous les utiliser, au détriment des autres utilisateurs.') ('Définir max degree of parallelism à ' + $recDop + ' (recommandation Microsoft pour cette topologie CPU/NUMA).')
    } elseif ($maxdop -gt $recDop) {
        $assess['max degree of parallelism'] = ('Supérieur à la valeur suggérée (' + $recDop + ')')
        Add-Finding 'Low' 'MAXDOP est supérieur à la recommandation' ('Valeur actuelle : ' + $maxdop + ' ; valeur suggérée pour cette topologie CPU/NUMA : ' + $recDop + '.') ('Envisager de réduire MAXDOP à ' + $recDop + '.')
    } elseif ($maxdop -eq 1 -and $cpu -gt 1) {
        $assess['max degree of parallelism'] = 'Parallélisme désactivé'
        Add-Finding 'Info' 'Le parallélisme est désactivé (MAXDOP 1)' "Les requêtes lourdes ne peuvent utiliser qu'un seul processeur. Certaines applications l'exigent (SharePoint par exemple), mais cela ralentit les requêtes de type reporting." ("Le conserver si l'éditeur de l'application l'exige ; sinon, envisager MAXDOP " + $recDop + ' avec un seuil de coût plus élevé.')
    } else {
        $assess['max degree of parallelism'] = 'OK'
    }

    $ctfp = [long]$cfg['cost threshold for parallelism'].value_in_use
    if ($ctfp -le 5) {
        $assess['cost threshold for parallelism'] = 'Valeur par défaut (5) trop basse'
        Add-Finding 'Medium' 'Le seuil de coût du parallélisme est à sa valeur par défaut (5)' "Même des requêtes légères s'exécutent en parallèle, ce qui gaspille du CPU et augmente les attentes de parallélisme." "Le porter à 50 comme point de départ, puis ajuster selon la charge."
    } elseif ($ctfp -lt 25) {
        $assess['cost threshold for parallelism'] = 'Bas'
    } else {
        $assess['cost threshold for parallelism'] = 'OK'
    }

    if ([long]$cfg['optimize for ad hoc workloads'].value_in_use -eq 0) {
        $assess['optimize for ad hoc workloads'] = 'Désactivé - à activer en général'
        Add-Finding 'Low' "L'option optimize for ad hoc workloads est désactivée" "Les plans de requêtes à usage unique sont mis en cache en entier et gaspillent la mémoire du cache de plans." "L'activer ; c'est sans risque pour la quasi-totalité des charges."
    } else { $assess['optimize for ad hoc workloads'] = 'OK' }

    if ([long]$cfg['priority boost'].value_in_use -eq 1) {
        $assess['priority boost'] = 'ACTIVÉ - non supporté'
        Add-Finding 'High' "L'option priority boost est activée" "Ce paramètre obsolète peut affamer le système d'exploitation et les services de cluster." "La désactiver et redémarrer SQL Server."
    }
    if ([long]$cfg['lightweight pooling'].value_in_use -eq 1) {
        $assess['lightweight pooling'] = 'ACTIVÉ - déconseillé'
        Add-Finding 'Medium' "Le mode fibre (lightweight pooling) est activé" "Le mode fibre est rarement bénéfique et empêche plusieurs fonctionnalités de fonctionner." "Le désactiver, sauf si un test précis a démontré un gain."
    }
    if ([long]$cfg['max worker threads'].value_in_use -ne 0) {
        $assess['max worker threads'] = 'Modifié par rapport à la valeur par défaut'
        Add-Finding 'Low' "Le paramètre max worker threads a été modifié" "L'augmenter masque généralement un problème de blocage ou de parallélisme au lieu de le résoudre." "Le remettre à 0 (automatique), sauf raison documentée."
    }

    $pending = @($rows | Where-Object { $_.value -ne $_.value_in_use } | ForEach-Object { $_.name })
    if ($pending.Count -gt 0) {
        Add-Finding 'Low' 'Des modifications de configuration sont en attente' ('Configurées mais pas encore actives : ' + (Join-Names $pending) + '.') "Exécuter RECONFIGURE ou redémarrer SQL Server pendant une fenêtre de maintenance."
    }

    $out = foreach ($r in ($rows | Sort-Object name)) {
        [pscustomobject]@{
            'Paramètre' = $r.name
            'Valeur active' = $r.value_in_use
            'Évaluation' = $(if ($assess.ContainsKey($r.name)) { $assess[$r.name] } else { 'OK' })
        }
    }
    Add-Table 'Principaux paramètres du serveur' $out @(3, 1.4, 3)
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
    if ($total -le 0) { Add-Note "Aucune attente significative enregistrée depuis le démarrage."; return }

    $out = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($r in $w) {
        if ($i -ge $Top) { break }
        $pct = 100.0 * [double]$r.wait_time_ms / $total
        $avg = [double]$r.wait_time_ms / [math]::Max(1, [double]$r.waiting_tasks_count)
        $info = Get-WaitInfo $r.wait_type
        [void]$out.Add([pscustomobject]@{
            "Type d'attente" = $r.wait_type
            '% des attentes' = [math]::Round($pct, 1)
            'Attente moy. (ms)' = [math]::Round($avg, 1)
            'Signification habituelle' = $info.M
        })
        if ($i -lt 5 -and $pct -ge 10) {
            $sev = 'Low'
            if ($pct -ge 25) { $sev = 'Medium' }
            if ($r.wait_type -like 'RESOURCE_SEMAPHORE*' -or $pct -ge 50) { $sev = 'High' }
            if ($r.wait_type -eq 'ASYNC_NETWORK_IO' -or $r.wait_type -like 'BACKUP*') { $sev = 'Low' }
            Add-Finding $sev ($r.wait_type + ' représente ' + (Format-N1 $pct) + ' % des attentes') ($info.M + ' Attente moyenne : ' + (Format-N1 $avg) + ' ms.') $info.F
        }
        $i++
    }
    Add-Table ('Principales attentes depuis le ' + (Format-Date $script:Ctx.StartTime)) $out @(2.2, 1, 1, 5)

    $tp = @($all | Where-Object { $_.wait_type -eq 'THREADPOOL' })
    if ($tp.Count -gt 0 -and [double]$tp[0].wait_time_ms -ge 10000) {
        Add-Finding 'High' "Pénurie de threads de travail constatée (attentes THREADPOOL)" ('Des requêtes ont attendu ' + (Format-N0 ([double]$tp[0].wait_time_ms / 1000)) + " secondes au total qu'un thread se libère. Les utilisateurs ont probablement subi des délais d'attente ou un serveur figé.") "Analyser les blocages et le parallélisme aux moments concernés ; ne pas se contenter d'augmenter max worker threads."
    }
    $sigPct = 100.0 * $signal / $total
    if ($sigPct -ge 20) {
        Add-Finding 'Medium' ('Les attentes de signal représentent ' + (Format-N1 $sigPct) + " % du temps d'attente") "Les tâches passent une grande partie de leur temps à attendre un processeur une fois leur ressource disponible : c'est un signe de pression CPU." "Optimiser les requêtes les plus consommatrices de CPU et revoir les paramètres de parallélisme ; vérifier la capacité CPU."
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
    if ($rb.Count -eq 0) { Add-Note "Aucun historique CPU disponible."; return }
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

    Add-Table 'Utilisation du processeur (une mesure par minute)' @(
        [pscustomobject]@{ 'Mesure' = 'Période couverte';                          'Valeur' = ('Depuis le ' + (Format-Date $oldest) + ' (' + $n + ' minutes)') },
        [pscustomobject]@{ 'Mesure' = 'CPU moyen de SQL Server (%)';               'Valeur' = [math]::Round($avg, 1) },
        [pscustomobject]@{ 'Mesure' = 'Pic de CPU de SQL Server (%)';              'Valeur' = [math]::Round($max, 1) },
        [pscustomobject]@{ 'Mesure' = 'Minutes avec un CPU SQL à 80 % ou plus';    'Valeur' = ([string]$busy + ' (' + (Format-N1 $busyPct) + ' %)') },
        [pscustomobject]@{ 'Mesure' = 'CPU moyen des autres processus (%)';        'Valeur' = [math]::Round($avgOther, 1) }
    ) @(2, 2)

    if ($avg -ge 80) {
        Add-Finding 'High' ('Le CPU de SQL Server est en moyenne à ' + (Format-N1 $avg) + ' %') "Le serveur est limité par le processeur : les requêtes font la queue pour obtenir du CPU et les temps de réponse en pâtissent." "Optimiser les requêtes les plus consommatrices (voir Requêtes les plus coûteuses), vérifier les paramètres de parallélisme, puis envisager plus de CPU."
    } elseif ($avg -ge 60 -or $busyPct -ge 20) {
        Add-Finding 'Medium' ('Le CPU de SQL Server est élevé (moyenne ' + (Format-N1 $avg) + ' %, pic ' + (Format-N1 $max) + ' %)') ('Le CPU a été à 80 % ou plus pendant ' + (Format-N1 $busyPct) + ' % de la période.') "Optimiser les requêtes les plus consommatrices (voir Requêtes les plus coûteuses) avant d'ajouter du matériel."
    }
    if ($avgOther -ge 20) {
        Add-Finding 'Medium' ("D'autres processus utilisent " + (Format-N1 $avgOther) + ' % du CPU de ce serveur') "Un autre programme que SQL Server (antivirus, autres services, autre instance) lui dispute le processeur." "Identifier le processus sur le serveur ; le déplacer, ou exclure les fichiers SQL Server de l'analyse antivirus."
    }

    $c = $script:Ctx.Counters
    $up = [math]::Max(1, [double]$script:Ctx.UptimeDays * 86400)
    if ($c.ContainsKey('Batch Requests/sec') -and $c.ContainsKey('SQL Compilations/sec')) {
        $bps = $c['Batch Requests/sec'] / $up; $cps = $c['SQL Compilations/sec'] / $up
        if ($bps -ge 10 -and $cps / $bps -ge 0.15) {
            Add-Finding 'Medium' ('Les compilations représentent ' + (Format-N0 (100 * $cps / $bps)) + ' % des requêtes') "La plupart des requêtes sont compilées au lieu de réutiliser un plan en cache, ce qui coûte du CPU. Cause habituelle : du SQL dynamique non paramétré." "Paramétrer les requêtes dans l'application ; activer optimize for ad hoc workloads ; envisager le paramétrage forcé (forced parameterization) sur la base la plus concernée."
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
    # Seuil de page life expectancy proportionnel à la mémoire : 300 s par tranche de 4 Go.
    $pleThreshold = [math]::Max(300, [math]::Round(($totalKB / 1024 / 1024) / 4 * 300))

    Add-Table 'Mémoire' @(
        [pscustomobject]@{ 'Mesure' = 'Mémoire disponible pour le système';          'Valeur' = ((Format-N0 $availMB) + ' Mo sur ' + (Format-N0 $totalOsMB) + ' Mo (' + (Format-N1 $availPct) + ' %)') },
        [pscustomobject]@{ 'Mesure' = 'État mémoire du système';                     'Valeur' = $sm.system_memory_state_desc },
        [pscustomobject]@{ 'Mesure' = 'Mémoire utilisée par SQL Server';             'Valeur' = ((Format-N0 ([double]$pm.physical_memory_in_use_kb / 1024)) + ' Mo') },
        [pscustomobject]@{ 'Mesure' = 'Mémoire cible / totale de SQL Server';        'Valeur' = ((Format-N0 ($targetKB / 1024)) + ' Mo / ' + (Format-N0 ($totalKB / 1024)) + ' Mo') },
        [pscustomobject]@{ 'Mesure' = 'Verrouillage des pages en mémoire';           'Valeur' = $(if ([double]$pm.locked_page_allocations_kb -gt 0) { 'Utilisé' } else { 'Non utilisé' }) },
        [pscustomobject]@{ 'Mesure' = 'Page life expectancy (instantané)';           'Valeur' = ((Format-N0 $ple) + ' s (repère pour cette taille : ' + (Format-N0 $pleThreshold) + ' s)') },
        [pscustomobject]@{ 'Mesure' = 'Allocations mémoire en attente (instantané)'; 'Valeur' = $pending }
    ) @(2, 3)

    if ($pm.process_physical_memory_low -or $sm.system_memory_state_desc -like '*low*' -or ($availMB -lt 512 -and $totalOsMB -gt 0)) {
        Add-Finding 'High' 'Le serveur manque de mémoire' ('Seulement ' + (Format-N0 $availMB) + ' Mo sont disponibles pour le système (état : ' + $sm.system_memory_state_desc + ').') "Vérifier max server memory et les autres processus du serveur ; le système est peut-être en train de paginer."
    }
    if ($null -ne $ple -and $ple -lt $pleThreshold -and [double]$script:Ctx.UptimeDays -ge 0.05) {
        Add-Finding 'Medium' ('La page life expectancy est basse (' + (Format-N0 $ple) + ' s)') ('Les pages de données restent peu de temps en mémoire (repère pour cette taille de mémoire : ' + (Format-N0 $pleThreshold) + " s) : SQL Server relit souvent les données sur disque. C'est une mesure instantanée, à vérifier à plusieurs moments de la journée.") "Réduire les gros parcours de tables (requêtes à fortes lectures, index manquants) ou ajouter de la mémoire."
    }
    if ($pending -gt 0) {
        Add-Finding 'Medium' ([string]$pending + " requête(s) en attente d'allocation mémoire au moment de l'audit") "Ces requêtes ne peuvent pas démarrer tant que la mémoire nécessaire à leurs tris et jointures n'est pas disponible." "Identifier les requêtes qui demandent beaucoup de mémoire et les optimiser ; vérifier max server memory."
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
            'Base' = $f.Db; 'Fichier' = $f.File; 'Type' = $f.Type
            'Lectures' = [long]$f.Reads
            'Lecture moy. (ms)' = $(if ($null -ne $f.ReadMs) { [math]::Round($f.ReadMs, 1) } else { $null })
            'Écritures' = [long]$f.Writes
            'Écriture moy. (ms)' = $(if ($null -ne $f.WriteMs) { [math]::Round($f.WriteMs, 1) } else { $null })
        }
    }
    Add-Table "Fichiers cumulant le plus d'attente d'E/S (depuis le démarrage)" $table @(2, 2, 1, 1.3, 1.2, 1.3, 1.2) "Repères : lectures des fichiers de données sous 20 ms, écritures du journal sous 5 ms (sous 2 ms sur SSD)."

    $slowData = @($files | Where-Object { $_.Type -eq 'ROWS' -and $_.Reads -ge 1000 -and $_.ReadMs -ge 20 } | Sort-Object ReadMs -Descending)
    if ($slowData.Count -gt 0) {
        $worst = $slowData[0]
        $sev = 'Medium'; if ($worst.ReadMs -ge 50) { $sev = 'High' }
        $list = @($slowData | ForEach-Object { $_.Db + '/' + $_.File + ' ' + (Format-N0 $_.ReadMs) + ' ms' })
        Add-Finding $sev ('Lectures lentes sur ' + $slowData.Count + ' fichier(s) de données') ('Latence moyenne de lecture : ' + (Join-Names $list 5) + '.') "Vérifier le stockage (latence du SAN ou du datastore de la VM, file d'attente disque). Réduire les parcours de tables grâce à l'indexation diminue aussi le volume lu."
    }
    $slowLog = @($files | Where-Object { $_.Type -eq 'LOG' -and $_.Writes -ge 1000 -and $_.WriteMs -ge 5 } | Sort-Object WriteMs -Descending)
    if ($slowLog.Count -gt 0) {
        $worst = $slowLog[0]
        $sev = 'Low'; if ($worst.WriteMs -ge 10) { $sev = 'Medium' }; if ($worst.WriteMs -ge 20) { $sev = 'High' }
        $list = @($slowLog | ForEach-Object { $_.Db + ' ' + (Format-N1 $_.WriteMs) + ' ms' })
        Add-Finding $sev ('Écritures lentes du journal de transactions sur ' + $slowLog.Count + ' base(s)') ("Chaque validation de transaction attend l'écriture du journal. Latence moyenne d'écriture : " + (Join-Names $list 5) + '.') "Placer les journaux les plus sollicités sur un stockage à faible latence, séparé des fichiers de données."
    }
}

function Test-TempDb {
    $files = @(Invoke-Q 'SELECT name, type_desc, CAST(size / 128.0 AS decimal(18,1)) AS size_mb, growth, is_percent_growth, physical_name FROM tempdb.sys.database_files')
    $data = @($files | Where-Object { $_.type_desc -eq 'ROWS' })
    $cpu = [math]::Max(1, [int]$script:Ctx.CpuCount)
    $rec = [math]::Min(8, $cpu)

    $table = foreach ($f in $files) {
        [pscustomobject]@{
            'Fichier' = $f.name; 'Type' = $f.type_desc; 'Taille (Mo)' = $f.size_mb
            'Croissance' = $(if ($f.is_percent_growth) { [string]$f.growth + ' %' } else { (Format-N0 ([double]$f.growth / 128)) + ' Mo' })
            'Chemin' = $f.physical_name
        }
    }
    Add-Table 'Fichiers de TempDB' $table @(1.5, 1, 1, 1, 4)

    if ($data.Count -lt $rec) {
        Add-Finding 'Medium' ('TempDB a ' + $data.Count + ' fichier(s) de données pour ' + $cpu + ' processeurs') "Un nombre insuffisant de fichiers provoque de la contention sur les pages d'allocation (attentes PAGELATCH) quand de nombreuses sessions utilisent des tables temporaires." ('Utiliser ' + $rec + ' fichiers de données de même taille (un par processeur, jusqu''à 8).')
    }
    if (@($data | ForEach-Object { [double]$_.size_mb } | Select-Object -Unique).Count -gt 1) {
        Add-Finding 'Low' 'Les fichiers de données de TempDB ont des tailles différentes' "SQL Server privilégie le plus gros fichier, ce qui annule l'intérêt d'avoir plusieurs fichiers." "Donner la même taille et la même croissance à tous les fichiers de données de tempdb."
    }
    if (@($files | Where-Object { $_.is_percent_growth }).Count -gt 0) {
        Add-Finding 'Low' 'TempDB utilise une croissance automatique en pourcentage' "Les croissances deviennent de plus en plus grosses et imprévisibles à mesure que les fichiers grossissent." "Utiliser une croissance fixe (par exemple de 256 à 1 024 Mo)."
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
    if ($dbs.Count -eq 0) { Add-Note "Aucune base utilisateur en ligne."; return }
    $issues = @{}
    foreach ($d in $dbs) { $issues[$d.name] = New-Object System.Collections.ArrayList }
    $shrink = @(); $close = @(); $verify = @(); $stats = @(); $compat = @()
    $maxCompat = 0; if ([int]$script:Ctx.Major -ge 11) { $maxCompat = [int]$script:Ctx.Major * 10 }

    foreach ($d in $dbs) {
        if ($d.is_auto_shrink_on) { $shrink += $d.name; [void]$issues[$d.name].Add('Auto-shrink activé') }
        if ($d.is_auto_close_on) { $close += $d.name; [void]$issues[$d.name].Add('Auto-close activé') }
        if ($d.page_verify_option_desc -ne 'CHECKSUM') { $verify += $d.name; [void]$issues[$d.name].Add('Page verify ' + $d.page_verify_option_desc) }
        if (-not $d.is_auto_create_stats_on -or -not $d.is_auto_update_stats_on) { $stats += $d.name; [void]$issues[$d.name].Add('Statistiques automatiques désactivées') }
        if ($maxCompat -gt 0 -and [int]$d.compatibility_level -lt ($maxCompat - 20)) { $compat += ($d.name + ' (' + $d.compatibility_level + ')'); [void]$issues[$d.name].Add('Niveau de compatibilité ancien : ' + $d.compatibility_level) }
    }

    $files = @(Invoke-Q 'SELECT DB_NAME(database_id) AS db, name, growth, is_percent_growth, CAST(size / 128.0 AS decimal(18,1)) AS size_mb FROM sys.master_files WHERE database_id > 4')
    $pct = @(); $tiny = @()
    foreach ($f in $files) {
        if (-not $f.db -or -not $issues.ContainsKey($f.db)) { continue }
        if ($f.is_percent_growth -and [int]$f.growth -gt 0) {
            $pct += ($f.db + '/' + $f.name); [void]$issues[$f.db].Add('Croissance en pourcentage : ' + $f.name)
        } elseif (-not $f.is_percent_growth -and [int]$f.growth -gt 0 -and [int]$f.growth -le 128) {
            $tiny += ($f.db + '/' + $f.name); [void]$issues[$f.db].Add('Croissance de 1 Mo ou moins : ' + $f.name)
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
                if ($issues.ContainsKey($b.name)) { [void]$issues[$b.name].Add('Pas de sauvegarde du journal depuis 24 h') }
            }
        }
    } catch { Add-Note ("Historique des sauvegardes non vérifié : " + (Get-ShortError $_)) }

    $manyVlf = @(); $worstVlf = 0
    try {
        $vl = @(Invoke-Q "SELECT d.name, COUNT(*) AS vlfs FROM sys.databases AS d CROSS APPLY sys.dm_db_log_info(d.database_id) AS li WHERE d.database_id > 4 AND d.state_desc = 'ONLINE' GROUP BY d.name HAVING COUNT(*) > 300")
        foreach ($v in $vl) {
            $manyVlf += ($v.name + ' (' + $v.vlfs + ')')
            if ([int]$v.vlfs -gt $worstVlf) { $worstVlf = [int]$v.vlfs }
            if ($issues.ContainsKey($v.name)) { [void]$issues[$v.name].Add('' + $v.vlfs + ' VLF') }
        }
    } catch { }

    $table = foreach ($d in $dbs) {
        if ($issues[$d.name].Count -gt 0) {
            [pscustomobject]@{
                'Base' = $d.name; 'Récupération' = $d.recovery_model_desc; 'Compat.' = $d.compatibility_level
                'Problèmes détectés' = ($issues[$d.name] -join ' ; ')
            }
        }
    }
    $table = @($table)
    $clean = $dbs.Count - $table.Count
    Add-Table 'Bases présentant des problèmes de configuration' $table @(2, 1.3, 0.8, 5) ([string]$clean + ' base(s) utilisateur sur ' + $dbs.Count + ' ne présentent aucun problème et ne sont pas listées.')

    if ($shrink.Count) { Add-Finding 'High' ('Auto-shrink est activé sur ' + $shrink.Count + ' base(s)') ((Join-Names $shrink) + '. Les cycles de réduction et de croissance fragmentent les index et consomment CPU et E/S.') "Désactiver AUTO_SHRINK." }
    if ($close.Count) { Add-Finding 'Medium' ('Auto-close est activé sur ' + $close.Count + ' base(s)') ((Join-Names $close) + '. La base est fermée et rouverte en permanence et perd son cache.') "Désactiver AUTO_CLOSE." }
    if ($stats.Count) { Add-Finding 'Medium' ('Les statistiques automatiques sont désactivées sur ' + $stats.Count + ' base(s)') ((Join-Names $stats) + ". Des statistiques obsolètes produisent de mauvais plans d'exécution.") "Activer AUTO_CREATE_STATISTICS et AUTO_UPDATE_STATISTICS, sauf exigence contraire de l'éditeur." }
    if ($verify.Count) { Add-Finding 'Medium' ("PAGE_VERIFY n'est pas à CHECKSUM sur " + $verify.Count + ' base(s)') ((Join-Names $verify) + '. Une corruption peut passer inaperçue.') "Passer PAGE_VERIFY à CHECKSUM." }
    if ($pct.Count) { Add-Finding 'Low' ('Croissance automatique en pourcentage sur ' + $pct.Count + ' fichier(s)') ((Join-Names $pct) + '. Les croissances deviennent de plus en plus grosses et lentes à mesure que les fichiers grossissent.') "Utiliser une croissance fixe (par exemple 256 à 1 024 Mo pour les données, 256 à 512 Mo pour les journaux)." }
    if ($tiny.Count) { Add-Finding 'Medium' ('Croissance automatique de 1 Mo ou moins sur ' + $tiny.Count + ' fichier(s)') ((Join-Names $tiny) + ". Les fichiers grossissent par milliers de petits pas, chacun bloquant l'activité.") "Pré-dimensionner les fichiers et définir une croissance fixe d'au moins 64 Mo." }
    if ($noLogBackup.Count) { Add-Finding 'Medium' ([string]$noLogBackup.Count + ' base(s) en récupération FULL sans sauvegarde du journal depuis 24 heures') ((Join-Names $noLogBackup) + ". Le journal grossit sans limite, ce qui ralentit les sauvegardes, les restaurations et le démarrage. (À ignorer si les sauvegardes du journal sont faites sur un autre réplica d'un groupe de disponibilité.)") "Planifier des sauvegardes régulières du journal, ou passer en récupération SIMPLE si la restauration à un instant précis n'est pas nécessaire." }
    if ($manyVlf.Count) {
        $sev = 'Low'; if ($worstVlf -gt 1000) { $sev = 'Medium' }
        Add-Finding $sev ('Nombre élevé de fichiers journaux virtuels (VLF) dans ' + $manyVlf.Count + ' base(s)') ((Join-Names $manyVlf) + '. Un grand nombre de VLF ralentit la récupération, les restaurations et les sauvegardes du journal.') "Réduire le journal une fois pendant une période calme, puis le faire regrossir par grands paliers fixes."
    }
    if ($compat.Count) { Add-Finding 'Low' ('Niveau de compatibilité ancien sur ' + $compat.Count + ' base(s)') ((Join-Names $compat) + ". Ces bases ne profitent pas des améliorations récentes de l'optimiseur.") "Tester puis relever le niveau de compatibilité (le Query Store aide à repérer les régressions)." }
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
            'Base' = $r.db
            'Exécutions' = [long]$r.execs
            'CPU total (s)' = [math]::Round([double]$r.cpu_us / 1000000, 1)
            'CPU moy. (ms)' = [math]::Round([double]$r.cpu_us / 1000 / $ex, 1)
            'Lectures moy.' = [long]([double]$r.reads / $ex)
            'Durée moy. (ms)' = [math]::Round([double]$r.dur_us / 1000 / $ex, 1)
            '% du CPU' = [math]::Round(100 * [double]$r.cpu_us / $totCpu, 1)
            'Texte de la requête' = (Format-QueryText $r.query_text)
        }
    }
    $table = @($table)
    Add-Table 'Instructions les plus consommatrices de CPU (instructions similaires regroupées)' $table @(1.4, 1.1, 1, 1, 1.1, 1.1, 0.8, 6) "Basé sur les plans actuellement en cache. Les requêtes recompilées ou évincées du cache ne sont pas incluses."

    if ($table.Count -gt 0) {
        $top1 = $table[0]
        if ($top1.'% du CPU' -ge 25) {
            Add-Finding 'Medium' ('Une seule instruction consomme ' + (Format-N1 $top1.'% du CPU') + ' % du CPU des requêtes en cache') ('Base ' + $top1.Base + ', ' + (Format-N0 $top1.'Exécutions') + ' exécutions, ' + (Format-N1 $top1.'CPU moy. (ms)') + " ms de CPU chacune. L'optimiser aurait un effet important.") "Examiner son plan d'exécution (parcours de tables, key lookups, conversions implicites) et son indexation. C'est la première ligne du tableau des requêtes les plus coûteuses."
        }
        $top5 = 0.0; foreach ($t in ($table | Select-Object -First 5)) { $top5 += [double]$t.'% du CPU' }
        if ($top5 -ge 60 -and $top1.'% du CPU' -lt 25) {
            Add-Finding 'Info' ('Les 5 premières instructions consomment ' + (Format-N0 $top5) + ' % du CPU des requêtes en cache') "La charge est concentrée : optimiser quelques requêtes ferait une différence visible." "Commencer par les premières lignes du tableau des requêtes les plus coûteuses."
        }
    }
}

function Test-MissingIndexes {
    $sql = @'
SELECT TOP ({TOP})
       DB_NAME(mid.database_id) AS [Base],
       mid.[statement] AS [Table],
       mid.equality_columns AS [Colonnes d'égalité],
       mid.inequality_columns AS [Colonnes d'inégalité],
       mid.included_columns AS [Colonnes incluses],
       migs.user_seeks + migs.user_scans AS [Utilisations],
       CAST(migs.avg_user_impact AS decimal(5,1)) AS [Gain estimé %],
       CAST(migs.avg_total_user_cost * migs.avg_user_impact * (migs.user_seeks + migs.user_scans) AS decimal(18,0)) AS [Score]
FROM sys.dm_db_missing_index_group_stats AS migs
JOIN sys.dm_db_missing_index_groups AS mig ON mig.index_group_handle = migs.group_handle
JOIN sys.dm_db_missing_index_details AS mid ON mid.index_handle = mig.index_handle
WHERE mid.database_id > 4
ORDER BY [Score] DESC;
'@
    $rows = @(Invoke-Q ($sql.Replace('{TOP}', [string]$Top)))
    Add-Table 'Suggestions d''index manquants (plus fort bénéfice en premier)' $rows @(1.3, 2.5, 2, 1.5, 2, 1, 0.9, 1.1) "Ces suggestions viennent de l'optimiseur : elles se recoupent souvent et sont parfois trop larges. Les examiner et les regrouper ; ne jamais les créer telles quelles."
    $high = @($rows | Where-Object { [double]$_.Score -ge 100000 })
    if ($high.Count -gt 0) {
        $tables = @($high | ForEach-Object { $_.Table } | Select-Object -Unique)
        Add-Finding 'Medium' ([string]$high.Count + " suggestion(s) d'index manquant à fort bénéfice") ("L'optimiseur a demandé à plusieurs reprises des index sur : " + (Join-Names $tables 4) + '.') "Examiner le tableau des index manquants, vérifier les index existants sur ces tables, puis ajouter des index regroupés après test."
    }
}

function Test-Blocking {
    $blocked = @(Invoke-Q @'
SELECT r.session_id AS [Session], r.blocking_session_id AS [Bloquée par], DB_NAME(r.database_id) AS [Base],
       r.wait_type AS [Type d'attente], CAST(r.wait_time / 1000.0 AS decimal(18,1)) AS [Attente (s)],
       s.login_name AS [Login], LEFT(st.text, 300) AS [Instruction]
FROM sys.dm_exec_requests AS r
JOIN sys.dm_exec_sessions AS s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) AS st
WHERE r.blocking_session_id <> 0
ORDER BY r.wait_time DESC
'@)
    $longTx = @(Invoke-Q @'
SELECT TOP (10) st.session_id AS [Session], s.login_name AS [Login], s.program_name AS [Programme],
       s.status AS [Statut], at.transaction_begin_time AS [Début],
       DATEDIFF(MINUTE, at.transaction_begin_time, GETDATE()) AS [Ouverte depuis (min)]
FROM sys.dm_tran_session_transactions AS st
JOIN sys.dm_tran_active_transactions AS at ON at.transaction_id = st.transaction_id
JOIN sys.dm_exec_sessions AS s ON s.session_id = st.session_id
WHERE s.is_user_process = 1 AND st.session_id <> @@SPID
  AND at.transaction_begin_time < DATEADD(MINUTE, -10, GETDATE())
ORDER BY at.transaction_begin_time
'@)
    foreach ($b in $blocked) { $b.Instruction = Format-QueryText $b.Instruction }
    if ($blocked.Count -gt 0) {
        Add-Table "Requêtes bloquées au moment de l'audit" $blocked @(0.8, 0.9, 1.3, 1.3, 0.9, 1.5, 5)
        Add-Finding 'Medium' ([string]$blocked.Count + " requête(s) bloquée(s) pendant l'audit") ('Attente la plus longue : ' + (Format-N1 $blocked[0].'Attente (s)') + ' s, bloquée par la session ' + $blocked[0].'Bloquée par' + '.') "Identifier la session en tête de chaîne et ce qu'elle fait ; rechercher les transactions longues et les index manquants sur les tables concernées."
    } else {
        Add-Note "Aucun blocage au moment de l'audit."
    }
    if ($longTx.Count -gt 0) {
        Add-Table 'Transactions ouvertes depuis plus de 10 minutes' $longTx @(0.8, 1.6, 2.2, 1, 1.5, 1.1)
        $sleeping = @($longTx | Where-Object { $_.Statut -eq 'sleeping' })
        $detail = 'La plus ancienne est ouverte depuis ' + (Format-N0 $longTx[0].'Ouverte depuis (min)') + ' minutes (session ' + $longTx[0].Session + ', ' + $longTx[0].Programme + ').'
        if ($sleeping.Count -gt 0) { $detail += ' ' + $sleeping.Count + " d'entre elles sont inactives (sleeping) avec une transaction ouverte, ce qui indique généralement une application qui a oublié de valider." }
        Add-Finding 'Medium' ([string]$longTx.Count + ' transaction(s) ouverte(s) depuis longtemps') ($detail + " Les transactions longues conservent leurs verrous et empêchent la réutilisation du journal.") "Voir avec le responsable de l'application ; corriger la gestion des COMMIT/ROLLBACK manquants."
    }
    $c = $script:Ctx.Counters
    if ($c.ContainsKey('Number of Deadlocks/sec') -and [double]$script:Ctx.UptimeDays -gt 0) {
        $perDay = $c['Number of Deadlocks/sec'] / [math]::Max(1, [double]$script:Ctx.UptimeDays)
        if ($perDay -ge 10) {
            Add-Finding 'Medium' ('Environ ' + (Format-N0 $perDay) + ' deadlocks par jour') ([string](Format-N0 $c['Number of Deadlocks/sec']) + " deadlocks depuis le démarrage. Chacun annule une transaction que l'application doit rejouer.") "Récupérer les graphes de deadlock dans la session Extended Events system_health et corriger l'ordre d'accès aux objets ou l'indexation."
        } elseif ($perDay -ge 1) {
            Add-Finding 'Low' ('Environ ' + (Format-N1 $perDay) + ' deadlocks par jour') ([string](Format-N0 $c['Number of Deadlocks/sec']) + ' deadlocks depuis le démarrage.') "Examiner les graphes de deadlock dans la session Extended Events system_health."
        }
    }
}

# =============================================================================
# Générateur .docx minimal (Office Open XML, sans dépendance)
# =============================================================================

$script:SevStyle = @{
    High   = @{ Fg = 'B42318'; Bg = 'FDE2E1' }
    Medium = @{ Fg = '9A5B00'; Bg = 'FDEBC8' }
    Low    = @{ Fg = '1F5F99'; Bg = 'DCEBFA' }
    Info   = @{ Fg = '4B5563'; Bg = 'ECEEF1' }
}
# Mêmes couleurs accessibles par libellé français (pour les cellules de tableau).
foreach ($k in @('High', 'Medium', 'Low', 'Info')) { $script:SevStyle[$script:SevLabel[$k]] = $script:SevStyle[$k] }
$script:SevRank = @{ High = 0; Medium = 1; Low = 2; Info = 3 }
$script:ContentWidth = 9866   # largeur A4 moins 2 x 1,8 cm de marges, en twips

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
    $runs = (New-WRun ('[' + $script:SevLabel[$F.Severity] + ']  ') -Bold -Color $st.Fg) + (New-WRun $F.Title -Bold)
    Add-WXml (New-WPara $runs -After 20 -Indent 120)
    if ($F.Detail) { Add-WXml (New-WPara (New-WRun $F.Detail) -After 20 -Indent 360) }
    if ($F.Fix) { Add-WXml (New-WPara ((New-WRun 'Que faire : ' -Bold -Color '1F3A5F') + (New-WRun $F.Fix)) -After 140 -Indent 360) }
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
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault>' +
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
        (New-WRun ' sur ') +
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
# Mise en page du rapport
# =============================================================================

function Write-Report([string]$Path, [datetime]$Started) {
    $script:Body = New-Object System.Text.StringBuilder
    $ctx = $script:Ctx
    $server = $ctx.Server; if (-not $server) { $server = $ctx.Instance }
    $findings = @($script:Findings | Sort-Object -Property @{ Expression = { $script:SevRank[$_.Severity] } }, @{ Expression = { $script:SectionOrder[$_.Area] } })
    $cnt = @{ High = 0; Medium = 0; Low = 0; Info = 0 }
    foreach ($f in $findings) { $cnt[$f.Severity]++ }

    # --- En-tête
    Add-WXml (New-WPara (New-WRun 'Audit de performance SQL Server') -Style 'Title')
    Add-WXml (New-WPara (New-WRun $server -Bold -Color '2E5C8A' -Size 28) -After 200)
    $facts = @(
        [pscustomobject]@{ 'Élément' = 'Version';           'Valeur' = $ctx.Version },
        [pscustomobject]@{ 'Élément' = 'Édition';           'Valeur' = $ctx.Edition },
        [pscustomobject]@{ 'Élément' = "Date de l'audit";   'Valeur' = (Format-Date $Started) },
        [pscustomobject]@{ 'Élément' = 'En service depuis'; 'Valeur' = $(if ($ctx.StartTime) { (Format-Date $ctx.StartTime) + ' (' + (Format-N1 $ctx.UptimeDays) + ' jours)' } else { '' }) },
        [pscustomobject]@{ 'Élément' = 'Compte utilisé';    'Valeur' = $ctx.AuditLogin }
    )
    Add-WTable $facts @(1, 3)

    # --- Synthèse
    Add-WXml (New-WPara (New-WRun 'Synthèse') -Style 'Heading1')
    if ($cnt.High -gt 0) {
        $verdict = 'Il y a ' + $cnt.High + ' problème(s) de gravité élevée qui méritent une attention rapide.'
    } elseif ($cnt.Medium -gt 0) {
        $verdict = 'Aucun problème critique, mais ' + $cnt.Medium + " point(s) de gravité moyenne méritent d'être examinés."
    } elseif ($cnt.Low -gt 0) {
        $verdict = "L'instance semble en bonne santé ; seules des améliorations mineures ont été relevées."
    } else {
        $verdict = "Aucun problème n'a été détecté au-delà des seuils utilisés par cet audit."
    }
    Add-WXml (New-WPara (New-WRun $verdict -Bold) -After 120)
    $summary = @(
        [pscustomobject]@{ 'Gravité' = $script:SevLabel['High'];   'Nombre' = $cnt.High;   'Signification' = 'Pénalise probablement déjà les performances ou la stabilité : agir rapidement.' },
        [pscustomobject]@{ 'Gravité' = $script:SevLabel['Medium']; 'Nombre' = $cnt.Medium; 'Signification' = 'Vrai problème ou risque à corriger.' },
        [pscustomobject]@{ 'Gravité' = $script:SevLabel['Low'];    'Nombre' = $cnt.Low;    'Signification' = 'Amélioration de bonne pratique.' },
        [pscustomobject]@{ 'Gravité' = $script:SevLabel['Info'];   'Nombre' = $cnt.Info;   'Signification' = 'Contexte, aucune action nécessaire.' }
    )
    Add-WTable $summary @(1, 0.7, 6) -SeverityColumn 'Gravité'

    if ($findings.Count -gt 0) {
        Add-WXml (New-WPara (New-WRun 'Tous les constats, du plus important au moins important') -Style 'Heading2')
        $i = 0
        $rows = foreach ($f in $findings) {
            $i++
            $cell = @{ Xml = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>' + (New-WRun $f.Title -Bold -Size 16) + '</w:p>' +
                              '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>' + (New-WRun $f.Detail -Size 16) + '</w:p>') }
            [pscustomobject]@{ '#' = $i; 'Gravité' = $script:SevLabel[$f.Severity]; 'Domaine' = $f.Area; 'Constat' = $cell; 'Que faire' = $f.Fix }
        }
        Add-WTable $rows @(0.4, 0.9, 1.3, 4.5, 3.4) -SeverityColumn 'Gravité'
    }

    # --- Sections détaillées
    $n = 0
    foreach ($s in $script:Sections) {
        $n++
        Add-WXml (New-WPara (New-WRun ([string]$n + '. ' + $s.Title)) -Style 'Heading1')
        if ($s.Intro) { Add-WXml (New-WPara (New-WRun $s.Intro -Italic -Color '4B5563') -After 120) }
        if ($s.Error) {
            Add-WXml (New-WPara (New-WRun ("Cette vérification n'a pas pu s'exécuter : " + $s.Error) -Color '9A5B00') -Shade 'FDEBC8' -After 120)
        }
        $secFindings = @($findings | Where-Object { $_.Area -eq $s.Title })
        foreach ($f in $secFindings) { Add-WFinding $f }
        if (-not $s.Error -and $secFindings.Count -eq 0) {
            Add-WXml (New-WPara (New-WRun 'Aucun problème détecté dans ce domaine.' -Color '2E7D32') -After 120)
        }
        foreach ($note in $s.Notes) { Add-WXml (New-WPara (New-WRun $note -Italic -Color '6B7280' -Size 18) -After 100) }
        foreach ($t in $s.Tables) {
            Add-WXml (New-WPara (New-WRun $t.Caption) -Style 'Heading2')
            if ($t.Rows.Count -eq 0) {
                Add-WXml (New-WPara (New-WRun 'Aucun.' -Italic -Color '6B7280') -After 120)
            } else {
                Add-WTable $t.Rows $t.Weights
            }
            if ($t.Note) { Add-WXml (New-WPara (New-WRun $t.Note -Italic -Color '6B7280' -Size 16) -After 160) }
        }
    }

    # --- À propos
    Add-WXml (New-WPara (New-WRun 'À propos de ce rapport') -Style 'Heading1')
    $about = @(
        "Cet audit est en lecture seule : il interroge uniquement des vues système et ne modifie rien sur le serveur.",
        "La plupart des chiffres (attentes, E/S, statistiques de requêtes, deadlocks) sont cumulés depuis le dernier redémarrage. D'autres, comme la page life expectancy, les blocages et les transactions ouvertes, sont des instantanés pris au moment de l'audit.",
        "Les seuils utilisés sont des repères courants, pas des règles absolues. Un constat signifie « à examiner » : le contexte (application, heure de la journée, travaux de maintenance) décide si une action est nécessaire. Testez toute modification avant de l'appliquer en production.",
        ('Généré par Invoke-SqlPerfAudit.ps1 version ' + $script:ScriptVersion + ' en ' + (Format-N0 ((Get-Date) - $Started).TotalSeconds) + ' secondes.')
    )
    foreach ($a in $about) { Add-WXml (New-WPara (New-WRun $a -Size 18) -After 80) }

    Save-Docx $Path ('Audit de performance SQL Server - ' + $server)
}

# =============================================================================
# Programme principal
# =============================================================================

function Invoke-InstanceAudit([string]$Instance) {
    $started = Get-Date
    $script:Ctx = @{ Instance = $Instance; Counters = @{} }
    $script:Findings = New-Object System.Collections.ArrayList
    $script:Sections = New-Object System.Collections.ArrayList

    Write-Host ('Audit de ' + $Instance + ' ...') -ForegroundColor Cyan
    $script:Conn = Connect-Instance $Instance
    try {
        [void](Invoke-Q 'SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SET LOCK_TIMEOUT 10000; SET DEADLOCK_PRIORITY LOW;')
        try { $script:Ctx.Counters = Get-PerfCounters } catch { Write-Warning ('Compteurs de performance indisponibles : ' + (Get-ShortError $_)) }

        Invoke-Check 'Instance' "Version, matériel vu par SQL Server, durée de fonctionnement et paramètres du service." { Test-Instance }
        Invoke-Check 'Configuration du serveur' "Paramètres globaux qui causent le plus souvent des problèmes de performance lorsqu'ils sont laissés par défaut." { Test-Configuration }
        Invoke-Check "Statistiques d'attente" "Ce que SQL Server passe son temps à attendre depuis le dernier redémarrage. Les principales attentes désignent le goulet d'étranglement." { Test-Waits }
        Invoke-Check 'Processeur' "Utilisation du processeur sur environ les quatre dernières heures, d'après le ring buffer de SQL Server." { Test-Cpu }
        Invoke-Check 'Mémoire' "Mémoire disponible pour le système et pour SQL Server, et signes de pression mémoire." { Test-Memory }
        Invoke-Check 'Latence du stockage' "Latence moyenne de lecture et d'écriture par fichier de base depuis le dernier redémarrage." { Test-Io }
        Invoke-Check 'TempDB' "Organisation des fichiers de TempDB, source fréquente de contention." { Test-TempDb }
        Invoke-Check 'Bases de données' "Options des bases, paramètres de croissance des fichiers, sauvegardes du journal et fichiers journaux virtuels." { Test-Databases }
        Invoke-Check 'Requêtes les plus coûteuses' "Les instructions qui ont consommé le plus de CPU, d'après le cache de plans." { Test-TopQueries }
        Invoke-Check 'Index manquants' "Index que l'optimiseur de requêtes aurait voulu utiliser." { Test-MissingIndexes }
        Invoke-Check 'Blocages et transactions' "Blocages, transactions ouvertes depuis longtemps et deadlocks." { Test-Blocking }
    } finally {
        $script:Conn.Close()
        $script:Conn.Dispose()
    }

    $script:SectionOrder = @{}
    $k = 0; foreach ($s in $script:Sections) { $script:SectionOrder[$s.Title] = $k; $k++ }

    if (-not (Test-Path -LiteralPath $OutputFolder)) { [void](New-Item -ItemType Directory -Path $OutputFolder) }
    $folder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath
    $safe = $Instance -replace '[\\/:*?"<>|,]', '_'
    $file = Join-Path $folder ('SQLPerfAudit_' + $safe + '_' + $started.ToString('yyyyMMdd_HHmm', $script:Inv) + '.docx')
    Write-Report $file $started

    $cnt = @{ High = 0; Medium = 0; Low = 0; Info = 0 }
    foreach ($f in $script:Findings) { $cnt[$f.Severity]++ }
    Write-Host ('  Rapport : ' + $file) -ForegroundColor Green
    Write-Host ('  Constats : ' + $cnt.High + ' élevée(s), ' + $cnt.Medium + ' moyenne(s), ' + $cnt.Low + ' faible(s), ' + $cnt.Info + ' info') -ForegroundColor Green
    [pscustomobject]@{ 'Instance' = $Instance; 'Rapport' = $file; 'Élevée' = $cnt.High; 'Moyenne' = $cnt.Medium; 'Faible' = $cnt.Low; 'Info' = $cnt.Info }
}

foreach ($inst in $SqlInstance) {
    try {
        Invoke-InstanceAudit $inst
    } catch {
        Write-Warning ("L'audit de " + $inst + ' a échoué : ' + (Get-ShortError $_))
    }
}
