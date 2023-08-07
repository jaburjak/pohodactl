<#
Copyright 2023 Jakub Jabůrek

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
#>

<#
.SYNOPSIS
    Automates STORMWARE POHODA accounting software.

.DESCRIPTION
    The pohodactl command can list connected clients, start and stop mServers and run automatic tasks.
    
    A “pohodactl.conf” configuration file should be located in the script directory, or you can provide your own path.
    The file must contain “CLIENT” and “SQLSERVER” options, each on a separate line, followed by equals and option
    value. No extra whitespace or comments are allowed. A backslash can be used as an escape character, and it must be
    escaped itself if meant verbatim. The “CLIENT” option must contain path to “Pohoda.exe” and “SQLSERVER” hostname of
    the SQL Server where POHODA database is located. The script will use credentials of the current user to
    authenticate.

.PARAMETER Command
    Main command, one of “client”, “mserver” and “task”.

.PARAMETER SubCommand
    Secondary command.
    
    “client list-active” displays connected POHODA clients.
    
    “mserver start” can be followed by the name of the mServer you want to start, otherwise all stopped mServer will be
    started.
    
    “mserver stop” can be followed by the name of the mServer you want to stop, otherwise all running mServers will be
    stopped.
    
    “mserver status” outputs all configured mServers and their status.

    “mserver health” sends a health-check request to verify mServers are really working. This command requires two
    additional options to be present in the “pohodactl.conf” file – “PHUSER” and “PHPASSWORD”, which must contain valid
    POHODA user credentials. The command can be followed by the name of a single mServer you want to check, otherwise
    it will query every configured mServer.
    
    “task run” must be followed by the number of the automatic task you want to run.

.PARAMETER Config
    Path to the “pohodactl.conf” configuration file.

.EXAMPLE
    PS> .\pohodactl.ps1 client list-active
    
     Id PohodaUser LastActive          Computer   RemoteComputer WindowsUser Database
     -- ---------- ----------          --------   -------------- ----------- --------
    123 @          01.01.2022 12:00:00 UCTOPC                    winusr      StwPh_12345678_2022
    124 @          01.01.2022 12:00:00 WINSERVER  SEF-PC         sef         StwPh_12345678_2022

.EXAMPLE
    PS> .\pohodactl.ps1 mserver status
    
    Year IsRunning Name     Ico      Url
    ---- --------- ----     ---      ---
    2022      True mserver  12345678 http://WINSERVER:8001

.EXAMPLE
    PS> .\pohodactl.ps1 mserver health

    IsRunning Name    IsResponding
    --------- ----    ------------
         True mserver        True

.EXAMPLE
    PS> .\pohodactl.ps1 mserver health mserver

    IsRunning Name    IsResponding
    --------- ----    ------------
         True mserver        True

.EXAMPLE
    PS> .\pohodactl.ps1 mserver start

.EXAMPLE
    PS> .\pohodactl.ps1 mserver start mserver_name

.EXAMPLE
    PS> .\pohodactl.ps1 mserver stop

.EXAMPLE
    PS> .\pohodactl.ps1 mserver stop mserver_name

.EXAMPLE
    PS> .\pohodactl.ps1 task run 42
#>

param(
    [Parameter(Mandatory = $true, Position = 0)] [string] $Command,
    [Parameter(Mandatory = $true, Position = 1)] [string] $SubCommand,
    [Parameter(Mandatory = $false, Position = 2)] [string] $Argument = "",
    [Parameter(Mandatory = $false)] [string] $Config = ""
)

if ($Config -eq "") {
    $Config = "$PSScriptRoot/pohodactl.conf"
}

function Get-PohodactlConfiguration {
    <#
    .SYNOPSIS
        Returns pohodactl configuration.
    
    .OUTPUTS
        System.Collections.Hashtable
        
        pohodactl configuration.
    #>
    
    param(
        # Path to pohodactl.conf configuration file.
        [Parameter(Mandatory = $true)] [string] $File
    )
    
    $config = @{}
    
    Get-Content $File | ConvertFrom-StringData | ForEach-Object {
        $name = $PSItem.Keys[0] | Select -First 1
        $config.Add($name, $PSItem[$name])
    }
    
    $options = @("SQLSERVER", "CLIENT")
    
    foreach ($option in $options) {
        if (!$config.ContainsKey($option)) {
            throw "Configuration is missing option $option."
        }
    }
    
    $config
}


function Invoke-Sql {
    <#
    .SYNOPSIS
        Executes SQL query.
    
    .OUTPUTS
        System.Data.DataTableCollection
        
        Query result.
    #>
    
    param(
        # SQL Server hostname.
        [Parameter(Mandatory = $true)] [string] $Server,
        # Database name.
        [Parameter(Mandatory = $true)] [string] $Database,
        # SQL query to execute.
        [Parameter(Mandatory = $true)] [string] $Query
    )
    
    $connectionString = "Data Source=$Server; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$Database"
    
    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($Query, $connection)
    $connection.Open()
    
    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null
    
    $connection.Close()
    $dataSet.Tables
}


function Get-PohodaActiveClients {
    <#
    .SYNOPSIS
        Returns a list of currently connected clients to POHODA.
    
    .OUTPUTS
        System.Collections.ArrayList
        
        List of connected clients. Each item will be a HashTable with keys “ID”, “PohodaUser”, “Database”,
        “WindowsUser”, “Computer”, “RemoteComputer” and “LastActive”.
    #>
    
    param(
        # SQL Server hostname.
        [Parameter(Mandatory = $true)] [string] $SqlServer
    )
    
    $query = Invoke-Sql -Query "SELECT * FROM dbo.ConnectedUsr;" -Database "StwPh_sys" -Server $SqlServer
    
    $clients = @()
    
    foreach ($row in $query) {
        $clients += @{
            Id = $row["ID"];
            PohodaUser = $row["PhUsr"];
            Database = $row["Db"];
            WindowsUser = $row["WinUsr"];
            Computer = $row["HostName"];
            RemoteComputer = $row["TsHostName"];
            LastActive = $row["LastTime"]
        }
    }
    
    $clients
}


function Get-PohodaMservers {
    <#
    .SYNOPSIS
        Returns a list of POHODA mServers and their status.
    
    .OUTPUTS
        System.Collections.ArrayList
        
        List of mServers. Each item will be a HashTable with keys “Name”, “IsRunning”, “Ico”, “Year” and “Url”.
    #>
    
    param(
        # Path to Pohoda.exe.
        [Parameter(Mandatory = $true)] [string] $Client
    )
    
    if (Test-Path "${env:temp}/mserver.xml" -PathType Leaf) {
        Remove-Item -Force "${env:temp}/mserver.xml"
    }
    
    Start-Process -NoNewWindow -FilePath $Client -ArgumentList @("/http", "list:xml", "${env:temp}/mserver.xml") -Wait
    
    [xml] $xml = Get-Content "${env:temp}/mserver.xml"
    
    $response = $xml.mServer.ChildNodes
    
    $mservers = @()
    
    foreach ($instance in $response) {
        $mservers += @{
            Name = $($instance.name);
            IsRunning = $($instance.running) -ieq "true";
            Ico = $($instance.company.ico);
            Year = $($instance.company.year);
            Url = $($instance.URI)
        }
    }
    
    $mservers
}


function Check-PohodaMserverHealth {
    <#
    .SYNOPSIS
        Checks whether the POHODA mServer is responding to requests.

    .DESCRIPTION
        Queries the mServer over HTTP and verifies whether it can return a success response. The mServer must respond
        within 15 seconds. Authentication is required in order to force the mServer to retrieve accounting unit
        information from the database, which verifies the database connection is working.

    .OUTPUTS
        bool
    #>

    param(
        # URL of the mServer.
        [Parameter(Mandatory = $true)] [string] $Url,
        # POHODA user name.
        [Parameter(Mandatory = $true)] [string] $User,
        # POHODA user password.
        [Parameter(Mandatory = $true)] [string] $Password
    )

    $authorization = "${User}:$Password"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($authorization)
    $encoded =[Convert]::ToBase64String($bytes)

    $headers = @{
        "STW-Authorization" = "Basic $encoded"
    }

    try {
        $response = Invoke-WebRequest -Uri "$Url/status?companyDetail" -Headers $headers -TimeoutSec 15
        $status = $response.StatusCode
    } catch {
        $status = $_.Exception.Response.StatusCode.value__
    }

    if ($status -eq 200) {
        return $true
    } else {
        return $false
    }
}


function Start-PohodaMserver {
    <#
    .SYNOPSIS
        Starts POHODA mServer.

    .DESCRIPTION
        The function will wait the specified amount of seconds after the start request is sent to POHODA. Only a single
        instance of a mServer will be started.
    #>
    
    param(
        # Path to Pohoda.exe.
        [Parameter(Mandatory = $true)] [string] $Client,
        # Name of the mServer to start.
        [Parameter(Mandatory = $true)] [string] $Name,
        # Wait time after starting the mServer process.
        [Parameter(Mandatory = $false)] [int] $Wait = 10
    )
    
    $status = Get-PohodaMservers -Client $Client
    $running = $false

    foreach ($instance in $status) {
        if ($instance.Name -eq $Name) {
            $running = $instance.IsRunning
        }
    }

    if (!$running) {
        Start-Process -NoNewWindow -FilePath $Client -ArgumentList @("/http", "start", $Name)
    
        # Cannot use -Wait in Start-Process, as it would block while the mServer is running.
    
        Sleep -Seconds $Wait
    }
}


function Stop-PohodaMserver {
    <#
    .SYNOPSIS
        Stops POHODA mServer.
    #>
    
    param(
        # Path to Pohoda.exe.
        [Parameter(Mandatory = $true)] [string] $Client,
        # Name of the mServer to stop. Omit this parameter to stop all mServers.
        [Parameter(Mandatory = $false)] [string] $Name = "*",
        # Wait time after stopping the mServer(s).
        [Parameter(Mandatory = $false)] [int] $Wait = 10
    )
    
    Start-Process -NoNewWindow -FilePath $Client -ArgumentList @("/http", "stop", $Name, "/f")

    Sleep -Seconds $Wait
}


function Invoke-PohodaTask {
    <#
    .SYNOPSIS
        Runs a POHODA automatic task.
    
    .DESCRIPTION
        For more information about POHODA automatic tasks, see
        https://www.stormware.cz/podpora/faq/pohoda/185/Jake-jsou-moznosti-automatickych-uloh-programu-POHODA/?id=3245.
    #>
    
    param(
        # Path to Pohoda.exe.
        [Parameter(Mandatory = $true)] [string] $Client,
        # Number of the task to run.
        [Parameter(Mandatory = $true)] [string] $Task
    )
    
    Start-Process -NoNewWindow -FilePath $Client -ArgumentList @("/TASK", $Task) -Wait
}


$cfg = Get-PohodactlConfiguration $Config

if ($Command -eq "client") {
    if ($SubCommand -eq "list-active") {
        Get-PohodaActiveClients -SqlServer $cfg.SQLSERVER | ForEach { [PSCustomObject] $_ } | Format-Table -AutoSize
        
        exit 0
    } else {
        throw "Unknown subcommand: $SubCommand."
    }
} elseif ($Command -eq "mserver") {
    if ($SubCommand -eq "start") {
        if ($Argument -eq "") {
            $Argument = "*"
        }

        if ($Argument -eq "*") {
            Get-PohodaMservers -Client $cfg.CLIENT | ForEach {
                if (-not $_.IsRunning) {
                    Start-PohodaMserver -Client $cfg.CLIENT -Name $_.Name
                }
            }
        } else {
            Start-PohodaMserver -Client $cfg.CLIENT -Name $Argument
        }

        exit 0
    } elseif ($SubCommand -eq "stop") {
        if ($Argument -eq "") {
            $Argument = "*"
        }
        
        Stop-PohodaMserver -Client $cfg.CLIENT -Name $Argument
        
        exit 0
    } elseif ($SubCommand -eq "status") {
        Get-PohodaMservers -Client $cfg.CLIENT | ForEach { [PSCustomObject] $_ } | Format-Table -AutoSize
        
        exit 0
    } elseif ($SubCommand -eq "health") {
        $requiredCfgOptions = @("PHUSER", "PHPASSWORD")

        foreach ($option in $requiredCfgOptions) {
            if (!$cfg.ContainsKey($option)) {
                throw "Configuration is missing option $option, which is required for this command."
            }
        }

        if ($Argument -eq "") {
            $Argument = "*"
        }

        $result = @()
        $code = 0

        Get-PohodaMservers -Client $cfg.CLIENT | ForEach {
            if ($Argument -eq "*" -or $_.Name -eq $Argument) {
                $running = $false
                $responding = $false

                if ($_.IsRunning) {
                    $running = $true

                    $responding = Check-PohodaMserverHealth -Url $_.Url -User $cfg.PHUSER -Password $cfg.PHPASSWORD

                    if (-not $responding) {
                        $code = 1
                    }
                } else {
                    $code = 1
                }

                $result += @{
                    Name = $_.Name;
                    IsRunning = $running;
                    IsResponding = $responding;
                }
            }
        }

        $result | ForEach { [PSCustomObject] $_ } | Format-Table -AutoSize

        exit $code
    } else {
        throw "Unknown subcommand: $SubCommand."
    }
} elseif ($Command -eq "task") {
    if ($SubCommand -eq "run") {
        if ($Argument -eq "") {
            throw "Missing task name."
        }
        
        Invoke-PohodaTask -Client $cfg.CLIENT -Task $Argument
        
        exit 0
    } else {
        throw "Unknown subcommand: $SubCommand."
    }
} else {
    throw "Unknown command: $Command."
}
