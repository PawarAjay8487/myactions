<#
    .SYNOPSIS 
        This script returns information about given servers.

    .DESCRIPTION
        This script accepts csv file as a input and provide information such as server ou, server os and other information regarding each server.
        output will be saved in csv format at given location by uer.

    .PARAMETER inputFile
        provide list of server in the form of text file.

    .PARAMETER outputFIle
        Provide path where you want output of this script to be saved.
    
    .Notes
        Created ON: 03/30/2021
        Modified ON: 03:20:2021
        Author: Ajay Pawar

    .Example
        Get-ServerInventory.ps1 -inputFile c:\temp\servers.txt -outputfile c:\temp\inventory.csv
        Save your server name in a text file, one server at each line and pass this file as a input to this script.
#>

[cmdletbinding()]
param (
    # this parameter accepts input file for server list
    [Parameter(Mandatory = $true)]
    [validatescript( { Test-Path -path $_ })]
    [string]
    $inputFile,

    # This parameter accepts path where output will be stored.
    [Parameter(Mandatory = $false)]
    [validatescript( { Test-Path -path $_ })]
    [string]
    $outputPath = "c:\temp\",

    # This parameter accepts path where output will be stored.
    [Parameter(Mandatory = $false)]
    [string]
    $To,
    
    # What action need to be taken
    [Parameter(Mandatory = $false)]
    [validateset("Display", "Email", "Save")]
    [string]
    $Action = "Display",

    # Add IP Address
    [Parameter(Mandatory = $false)]
    [switch]
    $retrieveIP
)

function get-rdpStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName
    )

    try {
        $ErrorActionPreference = "stop"
        $RDPStatus = (New-Object System.Net.Sockets.TcpClient($ComputerName, 3389)).connected    
    }
    catch {
        $RDPStatus = $false
    }
    finally {
        $ErrorActionPreference = "continue"
    }

    return $RDPStatus
}


function get-serviceCompliance {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,

        [Parameter(Mandatory = $true)]
        [string]
        $ServiceName,

        [Parameter(Mandatory = $true)]
        [string]
        $expectedStatus
    )

    $State = (Get-Service -Name $ServiceName -ComputerName $ComputerName).Status
    if ($State -eq $expectedStatus) {
        $isCompliant = $true
    }
    else {
        $isCompliant = $false
    }

    return $isCompliant

}

function get-cdriveSpaceInof {
    param ($computerName)    
    
    $CDrive = Get-CimInstance -ClassName win32_logicaldisk -ComputerName $ServerName -ErrorAction SilentlyContinue | Where-Object { $_.deviceID -eq "c:" }
    if ($CDrive) {
        $DriveCapacity = [math]::round($cdrive.Size / 1gb, 2)
        $freeSpace = [math]::round($cdrive.FreeSpace / 1gb, 2)
        $PercentFree = [math]::round($freespace / $DriveCapacity * 100, 2)
    }

    $cDriveSpaceInfo = [PSCustomObject]@{
        DriveCapacity = $DriveCapacity
        FreesSpace = $freeSpace
        PercentFree = $PercentFree
    }

    return $cDriveSpaceInfo
}

$inventoryResult = @()
$ServerList = Import-Csv -Path $inputFile -Header "ServerName", "ServiceName", "Status"

if ($Action -eq "Email") {
    if (-not $To) {
        Write-Output "Email Address is Mandatory"
        Exit 0
    }
}

foreach ($server in $ServerList) {
    try {
        $ServerName = $server.ServerName
        $ServiceName = $server.ServiceName
        $ServiceStatus = $server.Status

        IF ($retrieveIP.IsPresent) {
            [regex]$pattern = '^((?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
            $ipAddress = Resolve-DnsName -Name $ServerName | Where-Object { $_ -match $pattern } | Out-String        
        }

        $pingStatus = Test-Connection -ComputerName $ServerName -Quiet

        $RDPStatus = get-rdpStatus -ComputerName $ServerName        

        $winops = Get-CimInstance -Class win32_operatingsystem -ComputerName $ServerName -ErrorAction SilentlyContinue        

        $RAM = (Get-CimInstance -ClassName win32_physicalmemory -ComputerName $ServerName -ErrorAction SilentlyContinue | Measure-Object capacity -Sum).Sum / 1GB

        $cdrive = get-cdriveSpaceInof -computerName $ServerName
  
        $isCompliant = get-serviceCompliance -ComputerName $ServerName -ServiceName $ServiceName -expectedStatus $ServiceStatus

        $ServerOU = (get-adcomputer -identity $ServerName -properties canonicalName).canonicalName

        $inventoryResult += [PSCustomObject]@{
            ServerName     = $ServerName
            OU             = $ServerOU.substring(0, $ServerOU.Lastindexof("/"))
            PingState      = $pingStatus
            RDPStatus      = $RDPStatus
            OS             = $winops.caption
            LastBootupTime = $winops.lastbootuptime
            RAM            = $RAM
            CDriveCapacity = $cdrive.DriveCapacity
            FreeSpace      = $cdrive.freeSpace
            PercentFree    = $cdrive.PercentFree
            ServiceName    = $ServiceName
            isCompliant    = $isCompliant
        }
    }
    catch {
        Write-Output $_.Tostring()   
    }   

}

switch ($action) {
    "display" { $inventoryResult }
    "save" {
        $date = (Get-Date).Tostring("MMddyyyyhhmmss")

        If ((Get-ChildItem -Path $outputPath).PSIsContainer) {
            if ($outputPath.EndsWith("\")) {
                $outputFile = $outputPath + "serverinventoryReport-" + $date + ".csv"
            }
            else {
                $outputFile = $outputPath + "\serverinventoryReport-" + $date + ".csv"
            }
        }
        else {
            $outputFile = $outputPath
        }

        $inventoryResult | Export-Csv $outputPath -NoTypeInformation
    }
    "Email" {
        
        $mailbody = $inventoryResult | ConvertTo-Html
        $from = "IT@psautomate.com"
        $Subject = "inventory Report"
        $smtp = "smtp.psautomate.com"
        $port = 25

        Send-MailMessage -to $to -From $from -Subject $Subject -SmtpServer $smtp -Port $port -Body $mailbody -BodyAsHtml -Priority Normal
    }

    Default {}
}


