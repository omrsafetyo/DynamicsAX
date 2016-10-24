#region Load Assemblies
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoExtended")
#endregion Load Assemblies

#region Get AX SQL Instance
Function Get-AXSQLInstance {
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [string[]] $Computername,
        
        [parameter(Mandatory=$false)]
        [string] $Instance
    )
    
    BEGIN {
        Write-Verbose "Starting Get-AXSQLInstance"
        if ( $null -eq $Computername ) {
            $Computername = $ENV:COMPUTERNAME
        }
    }
    
    PROCESS {
        ForEach ( $Computer in $Computername ) {
            Write-Verbose "Checking $Computer"
            if ($PSBoundParameters.ContainsKey("Instance")) {
                if ( @(".","(local)") -contains $Instance ) {
                    $InstanceName = $Computer
                }
                else {
                    $InstanceName = "{0}\{1}" -f $Computer, $InstanceName
                }
            }
            if ( [string]::IsNullOrEmpty($InstanceName) ) {
                $SQLInstances = Get-WMIObject win32_service -computername $Computer -ErrorAction SilentlyContinue | Where-Object { ($_.Name -eq "MSSQLSERVER" -or $_.Name -like "MSSQL$*") -and $_.State -eq "Running" }
                
                ForEach ( $SQLInstance in $SQLInstances ) {
                    if ( $SQLInstance.Name -eq "MSSQLSERVER" ) {
                        Write-Verbose "Setting instance name to $Computer"
                        $InstanceName = $Computer
                    }
                    else {
                        $msg = "Non default instance, connecting to {0}" -f $SQLInstance.Name 
                        Write-Verbose $msg
                        $InstanceName = $SQLInstance.Name -replace "MSSQL\$", "${Computer}\"
                    }
                    Write-Verbose "Connecting to $InstanceName"
                    New-Object Microsoft.SqlServer.Management.Smo.Server $InstanceName
                }
            }
            else {
                Write-Verbose "Instance $InstanceName was passed"
                New-Object Microsoft.SqlServer.Management.Smo.Server $InstanceName
            }
        }
    }
}
#endregion Get AX SQL Instance

#region Get AX Database
Function Get-AXDatabase {
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Microsoft.SqlServer.Management.Smo.Server[]] $Instance,
        
        [parameter(Mandatory=$false)]
        [string[]] $Database
    )

    PROCESS {
        ForEach ( $SQLServer in $Instance ) {
            if ( $PSBoundParameters.ContainsKey("Database") ) {
                $SQLServer.Databases | Where-Object { -not($_.IsSystemObject) -and ($Database -contains $_.Name) -and ($null -ne $_.Tables["SYSSERVERSESSIONS"])}
            }
            else {
                $SQLServer.Databases | Where-Object { -not($_.IsSystemObject) -and $null -ne $_.Tables["SYSSERVERSESSIONS"] }
            }
        }
    }
}
#endregion Get AX Database

#region Get AOS Connections
Function Get-AXConnectedAOS {
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Microsoft.SqlServer.Management.Smo.Database[]] $Database
    )
    
    BEGIN {
        # Alive = 1
        # Drain = 2
        $TSQL = "SELECT AOSID, INSTANCE_NAME FROM [SYSSERVERSESSIONS] WHERE STATUS IN (1,2)"
    }
    
    PROCESS {
        ForEach ( $SQLDb in $Database ) {
            $AOSDataTable = $SQLDb.ExecuteWithResults($TSQL).Tables[0].Rows
            
            ForEach ($Record in $AOSDataTable) {
                $AOSServer = $Record.AOSID.Split("@")[0]
                $AOSInstance = $Record.INSTANCE_NAME
                
                $InstanceSearchFilter = '%Ax32Serv.exe" {0}' -f $AOSInstance
                $Service = Get-WmiObject Win32_Service -Computer $AOSServer -Filter "PathName like '$InstanceSearchFilter'" -ErrorAction SilentlyContinue
                $Service
            }
        }
    }
}
#endregion Get AOS Connections

#region Restart AX Instance
Function Restart-AXInstancesOnServer {
    [CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact="Medium")]
    PARAM(
        [parameter(Mandatory=$false,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Alias("Computer","Server")]
        [string[]] $Computername,
        
        [parameter(Mandatory=$false)]
        [string] $Instance,
        
        [parameter(Mandatory=$false)]
        [string[]] $Database,
        
        [switch]$Force
    )
    
    BEGIN {
        #region initialize
        if ( $null -eq $Computername ) {
            Write-Verbose "Using local computer"
            $Computername = $ENV:COMPUTERNAME
        }
        #endregion initialize
    }
    
    PROCESS {
        #region process
        ForEach ( $Computer in $Computername ) {
            #region find SQL Instances
            if ($PSBoundParameters.ContainsKey("Instance")) {
                Write-Verbose "Checking $instance on $computername"
                $SQLInstances = Get-AXSQLInstance -Computername $Computername -Instance $Instance
            }
            else {
                Write-Verbose "Scanning for SQL Services on $Computername"
                $SQLInstances = Get-AXSQLInstance -Computername $Computername
            }
            #endregion find SQL Instances
            
            ForEach ( $SQLInstance in $SQLInstances ) {
                #region find SQL Services
                $ConnectedAOSs = @()
                $InstanceName = $SQLInstance.InstanceName
                $NetBIOSName = $SQLInstance.ComputerNamePhysicalNetBIOS
                if ( [string]::IsNullOrEmpty($InstanceName) ) {
                    $InstanceName = $NetBIOSName
                    $ServiceName = "MSSQLSERVER"
                }
                else {
                    $ServiceName = 'MSSQL${0}' -f $InstanceName
                }
                $msg = "Checking {0} instance for Dynamics AX Databases" -f $InstanceName
                Write-Verbose $msg
                $SQLService = Get-Service -Computername $NetBIOSName $ServiceName
                #endregion find SQL Services
                
                #region find SQL Databases
                if ( $PSBoundParameters.ContainsKey("Database") ) {
                    $AXDBs = Get-AXDatabase -Instance $SQLInstance -Database $Database
                }
                else {
                    $AXDBs = Get-AXDatabase -Instance $SQLInstance
                }
                #endregion find SQL Databases
                
                #region find AOS Instances
                ForEach ( $DB in $AXDBs ) {
                    $msg = "Found database {0}, checking for AOS instances." -f $DB.Name
                    Write-Verbose $msg
                    $ConnectedAOSs += Get-AXConnectedAOS $DB
                }
                if ( $ConnectedAOSs.Count -eq 0 ) {
                    Write-Verbose "Didn't identify any AOS instances for $InstanceName"
                    continue
                }
                else {
                    $msg = "AOSInstances Found: {0}" -f $ConnectedAOSs.Count
                    Write-Verbose $msg
                }
                #endregion find AOS Instances
                
                #region Stop AOS Instances
                ForEach ( $AOS in $ConnectedAOSs ) {
                    # Stop the services
                    if ( $Force -or $PSCmdlet.ShouldProcess($AOS.Name,"Stop service")  ) {
                        $msg = "Stopping {0}" -f $AOS.Name
                        Write-Verbose $msg
                        $AOS.StopService()
                    }
                }
                #endregion Stop AOS Instances
                
                #region Restart SQL
                if ( $Force -or $PSCmdlet.ShouldProcess($SQLService.Name,"Restart service")  ) {
                    $msg = "Stopping {0}" -f $ServiceName
                    Write-Verbose $msg
                    Restart-Service -InputObject $SQLService
                }
                #endregion Restart SQL
                
                #region Start AOS Instances
                ForEach ( $AOS in $ConnectedAOSs ) {
                    # Stop the services
                    if ( $Force -or $PSCmdlet.ShouldProcess($AOS.Name,"Start service")  ) {
                        $msg = "Starting {0}" -f $AOS.Name
                        Write-Verbose $msg
                        $AOS.StartService()
                    }
                }
                #endregion Start AOS Instances
            }
        }
    #endregion process
    }
}
#region Restart AX Instance

Function Get-AXAOSInstanceReport {
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [Alias("Computer","Server")]
        [string[]] $Computername,
        
        [parameter(Mandatory=$false)]
        [string] $Instance,
        
        [parameter(Mandatory=$false)]
        [string[]] $Database
    )
    
    BEGIN {
        #region initialize
        if ( $null -eq $Computername ) {
            Write-Verbose "Using local computer"
            $Computername = $ENV:COMPUTERNAME
        }
        #endregion initialize
    }
    
    PROCESS {
        #region process
        if ( $PSBoundParameters.ContainsKey("Instance") ) {
            $AXSQLInstances = Get-AXSQLInstance -Computername $Computername -Instance $Instance
        }
        else {
            $AXSQLInstances = Get-AXSQLInstance -Computername $Computername
        }
        if ( $PSBoundParameters.ContainsKey("Database") ) {
            $AXDBs = Get-AXDatabase -Instance $AXSQLInstances -Database $Database
        }
        else {
            $AXDBs = Get-AXDatabase -Instance $AXSQLInstances
        }
        $ConnectedAOSs = Get-AXConnectedAOS $AXDBs
        $ConnectedAOSs | select PSComputerName, Name, ProcessId, State, PathName | format-table -autosize
        #endregion process
    }
}
