
[CmdletBinding()]
param(
    # Servidor de sitio (usado para logs y para el fallback de conexion SQL si -SqlInstance no se indica)
    [string]$SiteServer = "SERVERSCCM01.arag.local",

    # Codigo de sitio de Configuration Manager
    [string]$SiteCode = "ESP",

    # Instancia de SQL Server que aloja la base de datos del sitio (host o host\instancia)
    [string]$SqlInstance = "SERVERSCCM01",

    # Nombre de la base de datos del sitio
    [string]$Database = "CM_ESP",

    # Carpeta donde se generara el informe HTML y los CSV
    [string]$OutputPath = ".\SCCM_Assessment_Data",

    # Carpeta de instalacion de Configuration Manager en el servidor de sitio (para leer logs).
    # Ajusta si tu instalacion no esta en la ruta por defecto.
    [string]$CMInstallPath = "\\$SiteServer\SMS_$SiteCode\Logs",

    # Secciones a ejecutar. Por defecto, todas.
    [ValidateSet('Sql','Disco','Servicios','Logs','Certificados','Todo')]
    [string[]]$Sections = @('Todo'),

    # Numero de dias hacia atras para el resumen de mensajes de estado (status messages)
    [int]$DiasStatusMessages = 7,

    # Numero de lineas a leer del final de cada log revisado
    [int]$LineasLog = 400,

    # Muestra la lista de secciones disponibles y termina
    [switch]$ListSections,

    # Credenciales SQL explicitas (si no se indica, se usa autenticacion de Windows integrada)
    [System.Management.Automation.PSCredential]$SqlCredential
)

# ------------------------------------------------------------------------------
# Utilidades
# ------------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$script:Resultados = [ordered]@{}
$script:Errores     = New-Object System.Collections.Generic.List[string]

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = 'INFO')
    $color = switch ($Nivel) {
        'OK'   { 'Green' }
        'WARN' { 'Yellow' }
        'ERR'  { 'Red' }
        default { 'Cyan' }
    }
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [$Nivel] $Mensaje" -ForegroundColor $color
}

if ($ListSections) {
    Write-Host "Secciones disponibles: Sql, Disco, Servicios, Logs, Certificados, Todo"
    return
}

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$csvPath = Join-Path $OutputPath 'csv'
if (-not (Test-Path $csvPath)) {
    New-Item -ItemType Directory -Path $csvPath -Force | Out-Null
}

function Test-SectionEnabled {
    param([string]$Nombre)
    return ($Sections -contains 'Todo') -or ($Sections -contains $Nombre)
}

# Guardarraíl: solo se permiten sentencias de lectura (SELECT / WITH ... SELECT).
# Cualquier otra sentencia se rechaza explicitamente, incluso si el propio SQL
# la aceptaria (esto es una proteccion adicional, no solo una guia de estilo).
function Invoke-ReadOnlyQuery {
    param(
        [Parameter(Mandatory)][string]$Query,
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$DatabaseName,
        [string]$Descripcion = ''
    )

    $trimmed = $Query.Trim()
    if ($trimmed -notmatch '^(?is)\s*(SELECT|WITH)\b') {
        throw "Consulta rechazada por el guardarrail de solo lectura (debe empezar por SELECT o WITH): $Descripcion"
    }
    # Bloqueo adicional de palabras clave de escritura, por si aparecieran en un
    # segundo statement (defensa en profundidad; no deberian usarse ';' en estas consultas).
    $prohibidas = 'INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|TRUNCATE|EXEC(UTE)?|CREATE|GRANT|DENY|REVOKE'
    if ($trimmed -match "(?is)\b($prohibidas)\b") {
        throw "Consulta rechazada por el guardarrail de solo lectura (contiene una palabra clave de escritura): $Descripcion"
    }

    try {
        if (Get-Module -ListAvailable -Name SqlServer -ErrorAction SilentlyContinue) {
            Import-Module SqlServer -ErrorAction SilentlyContinue
        }
        $cmdInvokeSqlcmd = Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue
        if ($cmdInvokeSqlcmd) {
            # Se construyen todos los parametros "deseados" y luego se filtran a los
            # que realmente existen en la version de Invoke-Sqlcmd instalada en este
            # servidor (versiones antiguas del modulo SqlServer/SQLPS no tienen
            # -TrustServerCertificate, por ejemplo). Esto evita errores de
            # "a parameter cannot be found" en servidores con modulos mas antiguos.
            $paramsDeseados = @{
                ServerInstance          = $Instance
                Database                = $DatabaseName
                Query                   = $trimmed
                QueryTimeout            = 120
                TrustServerCertificate  = $true
                ErrorAction             = 'Stop'
            }
            if ($SqlCredential) {
                $paramsDeseados['Username'] = $SqlCredential.UserName
                $paramsDeseados['Password'] = $SqlCredential.GetNetworkCredential().Password
            }
            $params = @{}
            foreach ($clave in $paramsDeseados.Keys) {
                if ($cmdInvokeSqlcmd.Parameters.ContainsKey($clave)) {
                    $params[$clave] = $paramsDeseados[$clave]
                }
            }
            return Invoke-Sqlcmd @params
        }
        else {
            # Fallback nativo de solo lectura via ADO.NET, sin depender del modulo SqlServer.
            $connString = "Server=$Instance;Database=$DatabaseName;Integrated Security=True;Connection Timeout=30;"
            if ($SqlCredential) {
                $connString = "Server=$Instance;Database=$DatabaseName;User Id=$($SqlCredential.UserName);Password=$($SqlCredential.GetNetworkCredential().Password);Connection Timeout=30;"
            }
            $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
            $cmd  = $conn.CreateCommand()
            $cmd.CommandText = $trimmed
            $cmd.CommandTimeout = 120
            $conn.Open()
            $reader  = $cmd.ExecuteReader()
            $table   = New-Object System.Data.DataTable
            $table.Load($reader)
            $conn.Close()
            return $table
        }
    }
    catch {
        $script:Errores.Add("[$Descripcion] $($_.Exception.Message)")
        Write-Log "Fallo en '$Descripcion': $($_.Exception.Message)" 'ERR'
        return $null
    }
}

function Save-Resultado {
    param([string]$Clave, $Datos, [string]$Titulo)
    $script:Resultados[$Clave] = [pscustomobject]@{ Titulo = $Titulo; Datos = $Datos }
    if ($Datos) {
        try {
            $csvFile = Join-Path $csvPath "$Clave.csv"
            $Datos | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -Force
        } catch { }
    }
}

Write-Log "Iniciando recopilacion de datos de apoyo al assessment (SOLO LECTURA)"
Write-Log "Sitio: $SiteCode | Servidor: $SiteServer | SQL: $SqlInstance | BD: $Database"

# Deteccion local vs remoto: varias comprobaciones (WMI, ficheros locales de log,
# certificados) requieren permisos distintos segun si el script se ejecuta EN el
# propio servidor de sitio o EN REMOTO desde otro equipo (p.ej. tu portatil).
# Esto no evita los problemas de permisos, pero permite dar un mensaje de error
# mucho mas util y, cuando es posible, evitar por completo la necesidad de
# WinRM/DCOM remoto.
$nombreCortoServidor = ($SiteServer -split '\.')[0]
$script:EsLocal = ($nombreCortoServidor -ieq $env:COMPUTERNAME)
if ($script:EsLocal) {
    Write-Log "El script se esta ejecutando localmente en $SiteServer: se usaran llamadas locales (sin WinRM/DCOM remoto)." 'OK'

    # Si el usuario NO ha indicado -CMInstallPath explicitamente, intentamos resolver
    # la carpeta de logs LOCAL real via registro, en vez de usar la ruta de red por
    # defecto (\\servidor\SMS_<site>\Logs), que sigue pasando por el stack de red
    # (SMB) incluso ejecutando el script en el propio servidor y puede dar problemas
    # de permisos evitables. Lectura de registro: solo lectura, no se modifica nada.
    if (-not $PSBoundParameters.ContainsKey('CMInstallPath')) {
        try {
            $regSetup = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\SMS\Setup' -ErrorAction Stop
            $instalacionLocal = $regSetup.'Installation Directory'
            if ($instalacionLocal) {
                $rutaLogsLocal = Join-Path $instalacionLocal 'Logs'
                if (Test-Path $rutaLogsLocal -ErrorAction SilentlyContinue) {
                    $CMInstallPath = $rutaLogsLocal
                    Write-Log "Carpeta de logs local detectada via registro: $CMInstallPath" 'OK'
                }
            }
        }
        catch {
            Write-Log "No se pudo leer la carpeta de instalacion desde el registro (HKLM:\SOFTWARE\Microsoft\SMS\Setup); se mantiene la ruta de red por defecto para Logs. Puedes forzar la ruta local con -CMInstallPath." 'WARN'
        }
    }
}
else {
    Write-Log "El script se esta ejecutando en REMOTO contra $SiteServer desde $env:COMPUTERNAME." 'WARN'
    Write-Log "Las secciones Disco/Servicios (WMI remoto) y Certificados (WinRM) necesitan permisos remotos explicitos sobre $SiteServer (grupo local 'Remote Management Users' y/o 'Distributed COM Users', mas permisos WMI en el namespace root\cimv2). Si fallan, la alternativa mas simple es ejecutar el script por RDP directamente en el servidor de sitio." 'WARN'
}

# ------------------------------------------------------------------------------
# SECCION 1 - BASE DE DATOS SQL SERVER (area sin evidencia en el assessment)
# ------------------------------------------------------------------------------
if (Test-SectionEnabled 'Sql') {
    Write-Log "Seccion SQL Server..." 
    try {
        # 1.1 Version y edicion de SQL Server
        $q = "SELECT @@VERSION AS Version, SERVERPROPERTY('Edition') AS Edicion, SERVERPROPERTY('ProductLevel') AS ServicePack"
        Save-Resultado 'sql_version' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName 'master' -Descripcion 'Version SQL') 'Version y edicion de SQL Server'

        # 1.2 Tamano y crecimiento de ficheros de datos/log de la BD del sitio
        $q = @"
SELECT DB_NAME(mf.database_id) AS BaseDatos,
       mf.name AS NombreLogico,
       mf.type_desc AS Tipo,
       CAST(mf.size/128.0 AS DECIMAL(18,2)) AS TamanoActualMB,
       CASE WHEN mf.max_size = -1 THEN 'Sin limite'
            ELSE CAST(mf.max_size/128.0 AS VARCHAR(20)) END AS TamanoMaximoMB,
       CASE WHEN mf.is_percent_growth = 1
            THEN CAST(mf.growth AS VARCHAR(10)) + ' %'
            ELSE CAST(mf.growth/128.0 AS VARCHAR(20)) + ' MB' END AS Crecimiento,
       mf.physical_name AS RutaFisica
FROM sys.master_files mf
WHERE DB_NAME(mf.database_id) = '$Database'
"@
        Save-Resultado 'sql_ficheros' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName 'master' -Descripcion 'Tamano ficheros BD') 'Tamano y crecimiento de los ficheros de la base de datos'

        # 1.3 Espacio libre real en disco de los ficheros de datos/log (via xp_fixeddrives es EXEC -> no permitido por el guardarrail;
        #     se calcula en la seccion Disco con Get-PSDrive/Get-Volume en su lugar).

        # 1.4 Estado de Service Broker (usado internamente por Configuration Manager)
        $q = "SELECT name AS BaseDatos, is_broker_enabled AS BrokerHabilitado, state_desc AS Estado, recovery_model_desc AS ModeloRecuperacion FROM sys.databases WHERE name = '$Database'"
        Save-Resultado 'sql_broker' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName 'master' -Descripcion 'Service Broker') 'Estado de Service Broker y modelo de recuperacion'

        # 1.5 Historial de backups (ultimos 10) - requiere lectura sobre msdb
        $q = @"
SELECT TOP 10
       bs.database_name AS BaseDatos,
       CASE bs.type WHEN 'D' THEN 'Completo' WHEN 'I' THEN 'Diferencial' WHEN 'L' THEN 'Log' ELSE bs.type END AS TipoBackup,
       bs.backup_start_date AS Inicio,
       bs.backup_finish_date AS Fin,
       CAST(bs.backup_size/1024.0/1024.0 AS DECIMAL(18,2)) AS TamanoMB,
       bmf.physical_device_name AS Destino
FROM msdb.dbo.backupset bs
JOIN msdb.dbo.backupmediafamily bmf ON bs.media_set_id = bmf.media_set_id
WHERE bs.database_name = '$Database'
ORDER BY bs.backup_start_date DESC
"@
        Save-Resultado 'sql_backups' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName 'msdb' -Descripcion 'Historial de backups') 'Ultimos 10 backups de la base de datos del sitio'

        # 1.6 Historial de las tareas de mantenimiento de sitio (SQL Agent, si se ejecutan asi)
        $q = @"
SELECT TOP 20
       j.name AS Tarea,
       h.run_date AS FechaEjecucion,
       h.run_time AS HoraEjecucion,
       CASE h.run_status WHEN 0 THEN 'Fallo' WHEN 1 THEN 'Correcto' WHEN 2 THEN 'Reintentando' WHEN 3 THEN 'Cancelado' ELSE 'Desconocido' END AS Resultado,
       h.message AS Mensaje
FROM msdb.dbo.sysjobhistory h
JOIN msdb.dbo.sysjobs j ON h.job_id = j.job_id
WHERE j.name LIKE '%SMS%' OR j.name LIKE '%ConfigMgr%' OR j.name LIKE '%SCCM%' OR j.name LIKE '%$Database%'
ORDER BY h.run_date DESC, h.run_time DESC
"@
        Save-Resultado 'sql_agent_jobs' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName 'msdb' -Descripcion 'SQL Agent jobs de CM') 'Historial de tareas de SQL Agent relacionadas con Configuration Manager'

        # 1.7 Fragmentacion de indices (modo LIMITED: barato en I/O), top 25 mas fragmentados
        $q = @"
SELECT TOP 25
       OBJECT_NAME(ips.object_id) AS Tabla,
       i.name AS Indice,
       ips.index_type_desc AS Tipo,
       CAST(ips.avg_fragmentation_in_percent AS DECIMAL(5,2)) AS FragmentacionPct,
       ips.page_count AS Paginas
FROM sys.dm_db_index_physical_stats(DB_ID('$Database'), NULL, NULL, NULL, 'LIMITED') ips
JOIN sys.indexes i ON ips.object_id = i.object_id AND ips.index_id = i.index_id
WHERE ips.page_count > 100
ORDER BY ips.avg_fragmentation_in_percent DESC
"@
        Save-Resultado 'sql_fragmentacion' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName $Database -Descripcion 'Fragmentacion de indices') 'Top 25 indices mas fragmentados (modo LIMITED)'

        # 1.8 Resumen agregado de cumplimiento de actualizaciones de software (contraste con lo visto en consola)
        $q = @"
SELECT
    CASE Status
        WHEN 0 THEN 'Unknown'
        WHEN 1 THEN 'NotRequired'
        WHEN 2 THEN 'Required'
        WHEN 3 THEN 'Installed'
    END AS Estado,
    COUNT(*) AS NumRegistros
FROM v_UpdateComplianceStatus
GROUP BY Status
ORDER BY Status
"@
        Save-Resultado 'sql_update_compliance' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName $Database -Descripcion 'Compliance de actualizaciones') 'Resumen agregado de v_UpdateComplianceStatus (contraste con el Software Updates Dashboard)'

        # 1.9 Recuento real de Software Update Groups y cuantos estan realmente vacios (0 updates)
        #     Ajusta el nombre de la vista si tu version usa un esquema distinto.
        $q = @"
SELECT
    COUNT(*) AS TotalGrupos,
    SUM(CASE WHEN NumberOfUpdates = 0 THEN 1 ELSE 0 END) AS GruposVacios,
    SUM(CASE WHEN CI_ID IS NOT NULL AND Title LIKE 'Despliegue Antivirus Defender%' THEN 1 ELSE 0 END) AS GruposDefenderADR
FROM v_AuthListInfo
"@
        Save-Resultado 'sql_update_groups' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName $Database -Descripcion 'Recuento Software Update Groups') 'Recuento real de Software Update Groups (total / vacios / generados por la ADR de Defender)'

        # 1.10 Resumen de estado de clientes (v_CH_ClientSummary), si el rol Client Health esta activo
        $q = @"
SELECT
    SUM(CASE WHEN ClientActiveStatus = 1 THEN 1 ELSE 0 END) AS ClientesActivos,
    SUM(CASE WHEN ClientActiveStatus = 0 THEN 1 ELSE 0 END) AS ClientesInactivos,
    SUM(CASE WHEN ClientRemediationSuccess = 1 THEN 1 ELSE 0 END) AS RemediacionOK,
    COUNT(*) AS TotalEvaluados
FROM v_CH_ClientSummary
"@
        Save-Resultado 'sql_client_health' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName $Database -Descripcion 'Client Health summary') 'Resumen agregado de salud de cliente (v_CH_ClientSummary)'

        # 1.11 Mensajes de estado (status messages) de severidad Error/Warning en los ultimos N dias, por componente
        $q = @"
SELECT TOP 30
       ModuleName AS Componente,
       Severity   AS Severidad,
       COUNT(*)   AS NumMensajes,
       MAX(Time)  AS UltimaOcurrencia
FROM v_StatusMessage
WHERE Time >= DATEADD(day, -$DiasStatusMessages, GETDATE())
  AND Severity IN (-1073741824, 1073741824)   -- -1073741824=Error, 1073741824=Warning en el esquema SMS
GROUP BY ModuleName, Severity
ORDER BY NumMensajes DESC
"@
        Save-Resultado 'sql_status_messages' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName $Database -Descripcion 'Status messages') "Mensajes de Error/Warning por componente (ultimos $DiasStatusMessages dias)"

        # 1.12 Tamano del Content Library / paquetes distribuidos (vista de contenido)
        $q = @"
SELECT TOP 20
       PkgID AS IdPaquete,
       Name  AS Nombre,
       CAST(SourceSize AS DECIMAL(18,2)) AS TamanoOrigenKB,
       CAST(CompressedSize AS DECIMAL(18,2)) AS TamanoComprimidoKB
FROM v_Package
ORDER BY SourceSize DESC
"@
        Save-Resultado 'sql_content_grande' (Invoke-ReadOnlyQuery -Query $q -Instance $SqlInstance -DatabaseName $Database -Descripcion 'Paquetes mas grandes') 'Top 20 paquetes/contenido por tamano (candidatos a limpieza si estan obsoletos)'

        Write-Log "Seccion SQL Server completada" 'OK'
    }
    catch {
        Write-Log "Error inesperado en la seccion SQL: $($_.Exception.Message)" 'ERR'
        $script:Errores.Add("Seccion SQL: $($_.Exception.Message)")
    }
}

# ------------------------------------------------------------------------------
# SECCION 2 - ESPACIO EN DISCO (servidor de sitio / SQL / Distribution Point)
# ------------------------------------------------------------------------------
if (Test-SectionEnabled 'Disco') {
    Write-Log "Seccion Espacio en disco..."
    try {
        if ($script:EsLocal) {
            $discosRaw = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
        }
        else {
            $discosRaw = Get-CimInstance -ComputerName $SiteServer -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
        }
        $discos = $discosRaw |
            Select-Object DeviceID,
                @{N='VolumeName';E={$_.VolumeName}},
                @{N='TamanoTotalGB';E={[math]::Round($_.Size/1GB,2)}},
                @{N='EspacioLibreGB';E={[math]::Round($_.FreeSpace/1GB,2)}},
                @{N='PctLibre';E={ if ($_.Size -gt 0) { [math]::Round(($_.FreeSpace/$_.Size)*100,1) } else { 0 } }}
        Save-Resultado 'disco_servidor' $discos 'Espacio en disco del servidor de sitio (todas las unidades)'

        foreach ($d in $discos) {
            if ($d.PctLibre -lt 15) {
                Write-Log "AVISO: unidad $($d.DeviceID) con solo $($d.PctLibre)% libre ($($d.EspacioLibreGB) GB)" 'WARN'
            }
        }
        Write-Log "Seccion Espacio en disco completada" 'OK'
    }
    catch {
        $pista = if ($script:EsLocal) {
            "Comprueba que tu usuario tiene permisos para consultar WMI localmente (normalmente no requiere ser admin, pero politicas locales pueden restringirlo)."
        } else {
            "Ejecucion en remoto: comprueba que tu cuenta esta en el grupo local 'Distributed COM Users' (o es administradora local) en $SiteServer, y que el firewall permite WMI/DCOM. Alternativa mas simple: ejecuta el script directamente en el servidor."
        }
        Write-Log "Error en la seccion Disco: $($_.Exception.Message) -- $pista" 'ERR'
        $script:Errores.Add("Seccion Disco: $($_.Exception.Message) -- $pista")
    }
}

# ------------------------------------------------------------------------------
# SECCION 3 - SERVICIOS CLAVE (site component manager, SQL, IIS, WSUS)
# ------------------------------------------------------------------------------
if (Test-SectionEnabled 'Servicios') {
    Write-Log "Seccion Servicios..."
    try {
        $nombresServicio = @(
            'SMS_EXECUTIVE', 'SMS_SITE_COMPONENT_MANAGER', 'SMS_NOTIFICATION_SERVER',
            'MSSQLSERVER', 'SQLSERVERAGENT', 'W3SVC', 'WsusService', 'BITS'
        )
        $servicios = foreach ($n in $nombresServicio) {
            if ($script:EsLocal) {
                $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$n'" -ErrorAction SilentlyContinue
            }
            else {
                $svc = Get-CimInstance -ComputerName $SiteServer -ClassName Win32_Service -Filter "Name='$n'" -ErrorAction SilentlyContinue
            }
            if ($svc) {
                [pscustomobject]@{
                    Servicio     = $svc.DisplayName
                    NombreCorto  = $svc.Name
                    Estado       = $svc.State
                    TipoInicio   = $svc.StartMode
                    CuentaServicio = $svc.StartName
                }
            }
            else {
                [pscustomobject]@{ Servicio = $n; NombreCorto = $n; Estado = 'No encontrado en este servidor'; TipoInicio = ''; CuentaServicio = '' }
            }
        }
        Save-Resultado 'servicios' $servicios 'Estado de servicios clave (CM, SQL, IIS, WSUS, BITS)'
        $paradosCriticos = $servicios | Where-Object { $_.Estado -notin @('Running') -and $_.Estado -ne 'No encontrado en este servidor' }
        foreach ($s in $paradosCriticos) {
            Write-Log "AVISO: el servicio '$($s.Servicio)' esta en estado '$($s.Estado)'" 'WARN'
        }
        Write-Log "Seccion Servicios completada" 'OK'
    }
    catch {
        Write-Log "Error en la seccion Servicios: $($_.Exception.Message)" 'ERR'
        $script:Errores.Add("Seccion Servicios: $($_.Exception.Message)")
    }
}

# ------------------------------------------------------------------------------
# SECCION 4 - LOGS (lectura de las ultimas lineas, busqueda de errores)
# ------------------------------------------------------------------------------
if (Test-SectionEnabled 'Logs') {
    Write-Log "Seccion Logs..."
    try {
        $logsARevisar = @(
            'sitecomp.log', 'hman.log', 'sitestat.log', 'rcmctrl.log',
            'wsyncmgr.log', 'wcm.log', 'distmgr.log', 'smsdbmon.log',
            'srsrp.log', 'ruleengine.log'
        )
        $resumenLogs = New-Object System.Collections.Generic.List[object]
        foreach ($logName in $logsARevisar) {
            $logFile = Join-Path $CMInstallPath $logName
            $existe = $false
            try { $existe = Test-Path $logFile -ErrorAction Stop } catch { $existe = $false }

            if ($existe) {
                try {
                    $lineas = Get-Content -Path $logFile -Tail $LineasLog -ErrorAction Stop
                    $errores = $lineas | Select-String -Pattern 'failed|error|severity="3"' -SimpleMatch:$false

                    # Se calcula aparte (en vez de en linea dentro del hashtable) para evitar
                    # ambiguedades del parser de PowerShell 5.1 con if/else + -replace encadenados.
                    $ultimoError = ''
                    if ($errores.Count -gt 0) {
                        $textoBruto  = $errores[-1].Line
                        $textoLimpio = $textoBruto  -replace '<!\[LOG\[', ''
                        $textoLimpio = $textoLimpio -replace '\]LOG\]!>.*', ''
                        $ultimoError = $textoLimpio.Trim()
                    }

                    $resumenLogs.Add([pscustomobject]@{
                        Log                 = $logName
                        LineasRevisadas     = $lineas.Count
                        CoincidenciasError  = $errores.Count
                        UltimaModificacion  = (Get-Item $logFile).LastWriteTime
                        EjemploUltimoError  = $ultimoError
                    })
                }
                catch {
                    $resumenLogs.Add([pscustomobject]@{ Log = $logName; LineasRevisadas = 0; CoincidenciasError = 'No se pudo leer'; UltimaModificacion=''; EjemploUltimoError = $_.Exception.Message })
                }
            }
            else {
                $resumenLogs.Add([pscustomobject]@{ Log = $logName; LineasRevisadas = 0; CoincidenciasError = 'No encontrado'; UltimaModificacion=''; EjemploUltimoError = "Ruta comprobada: $logFile" })
            }
        }
        Save-Resultado 'logs' $resumenLogs "Resumen de errores en las ultimas $LineasLog lineas de logs clave del sitio"
        foreach ($l in $resumenLogs) {
            if ($l.CoincidenciasError -is [int] -and $l.CoincidenciasError -gt 0) {
                Write-Log "AVISO: $($l.CoincidenciasError) coincidencias de error en $($l.Log)" 'WARN'
            }
        }
        Write-Log "Seccion Logs completada" 'OK'
    }
    catch {
        $pista = if ($script:EsLocal) {
            "Comprueba que tu usuario tiene permisos de lectura sobre '$CMInstallPath'."
        } else {
            "Ejecucion en remoto: la ruta '$CMInstallPath' es un recurso compartido de red (\\$SiteServer\SMS_$SiteCode\Logs). Necesitas permisos de lectura sobre ese recurso compartido (grupo 'SMS Admins' o permiso NTFS/compartido explicito). Alternativas: (a) ejecuta el script directamente en el servidor de sitio con -CMInstallPath apuntando a la carpeta local de logs (p.ej. 'C:\Program Files\Microsoft Configuration Manager\Logs'), o (b) pide que te den permiso de lectura sobre el recurso compartido SMS_$SiteCode."
        }
        Write-Log "Error en la seccion Logs: $($_.Exception.Message) -- $pista" 'ERR'
        $script:Errores.Add("Seccion Logs: $($_.Exception.Message) -- $pista")
    }
}

# ------------------------------------------------------------------------------
# SECCION 5 - CERTIFICADOS PROXIMOS A CADUCAR (informativo, revision manual despues)
# ------------------------------------------------------------------------------
if (Test-SectionEnabled 'Certificados') {
    Write-Log "Seccion Certificados..."
    try {
        if ($script:EsLocal) {
            $certs = Get-ChildItem -Path 'Cert:\LocalMachine\My' | Select-Object Subject, Issuer, NotAfter, Thumbprint
        }
        else {
            $certs = Invoke-Command -ComputerName $SiteServer -ScriptBlock {
                Get-ChildItem -Path 'Cert:\LocalMachine\My' | Select-Object Subject, Issuer, NotAfter, Thumbprint
            } -ErrorAction Stop
        }

        $proximos = $certs | Where-Object { $_.NotAfter -lt (Get-Date).AddDays(90) } |
            Select-Object Subject, Issuer, NotAfter, Thumbprint,
                @{N='DiasParaExpirar';E={ [math]::Round((New-TimeSpan -Start (Get-Date) -End $_.NotAfter).TotalDays,0) }}

        Save-Resultado 'certificados' $proximos 'Certificados en LocalMachine\My que caducan en menos de 90 dias (revisar manualmente cuales son de Configuration Manager: CMG, IIS, firma de contenido, etc.)'
        foreach ($c in $proximos) {
            Write-Log "AVISO: certificado '$($c.Subject)' caduca en $($c.DiasParaExpirar) dias" 'WARN'
        }
        Write-Log "Seccion Certificados completada" 'OK'
    }
    catch {
        $pista = if ($script:EsLocal) {
            "Ejecucion local: comprueba que tu usuario puede leer Cert:\LocalMachine\My (normalmente requiere estar en el grupo de administradores locales)."
        } else {
            "Ejecucion en remoto: requiere WinRM habilitado en $SiteServer y tu cuenta en el grupo local 'Remote Management Users' (o ser administradora local). Alternativa mas simple: ejecuta el script directamente en el servidor."
        }
        Write-Log "No se pudo leer el almacen de certificados: $($_.Exception.Message) -- $pista" 'WARN'
        $script:Errores.Add("Seccion Certificados: $($_.Exception.Message) -- $pista")
    }
}

# ------------------------------------------------------------------------------
# INFORME HTML
# ------------------------------------------------------------------------------
Write-Log "Generando informe HTML..."

function ConvertTo-HtmlTabla {
    param($Datos)
    if (-not $Datos) { return "<p class='vacio'>Sin datos (revisa el log de ejecucion; puede requerir permisos adicionales o el nombre de vista/tabla difiere en tu version).</p>" }
    return ($Datos | ConvertTo-Html -Fragment) -join "`n"
}

$fecha = Get-Date -Format 'dd/MM/yyyy HH:mm'
$secciones = foreach ($clave in $script:Resultados.Keys) {
    $r = $script:Resultados[$clave]
    "<h2>$($r.Titulo)</h2>`n" + (ConvertTo-HtmlTabla $r.Datos)
}

$erroresHtml = "<p>Sin errores durante la ejecucion.</p>"
if ($script:Errores.Count -gt 0) {
    $listaErrores = ($script:Errores | ForEach-Object { "<li>$_</li>" }) -join "`n"
    $erroresHtml = "<h2>Avisos durante la ejecucion</h2><ul>" + $listaErrores + "</ul>"
}

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Datos de apoyo — Assessment SCCM ARAG</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; margin: 30px; color: #222; }
  h1 { color: #1F4E78; border-bottom: 3px solid #2E75B6; padding-bottom: 8px; }
  h2 { color: #2E75B6; margin-top: 30px; border-bottom: 1px solid #BFBFBF; padding-bottom: 4px; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 10px; font-size: 13px; }
  th { background: #2E75B6; color: white; text-align: left; padding: 6px 8px; }
  td { padding: 5px 8px; border-bottom: 1px solid #eee; }
  tr:nth-child(even) { background: #f7f9fb; }
  .vacio { color: #999; font-style: italic; }
  .meta { color: #595959; font-size: 13px; }
</style>
</head>
<body>
<h1>Datos de apoyo al Assessment SCCM — ARAG</h1>
<p class="meta">Generado el $fecha | Sitio: $SiteCode | Servidor: $SiteServer | SQL: $SqlInstance\$Database</p>
<p class="meta"><b>Nota:</b> este informe es un complemento tecnico de solo lectura al assessment realizado desde la consola. Los datos aqui mostrados deben cruzarse con los hallazgos ya documentados en el Word del assessment.</p>
$($secciones -join "`n")
$erroresHtml
</body>
</html>
"@

$htmlFile = Join-Path $OutputPath "Datos_Apoyo_Assessment_SCCM_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
$html | Out-File -FilePath $htmlFile -Encoding UTF8

Write-Log "Informe generado en: $htmlFile" 'OK'
Write-Log "CSV individuales en: $csvPath" 'OK'
if ($script:Errores.Count -gt 0) {
    Write-Log "$($script:Errores.Count) seccion(es)/consulta(s) tuvieron avisos. Revisa el informe HTML para el detalle." 'WARN'
}
Write-Log "Fin de la ejecucion." 'OK'
