[CmdletBinding()]
param (
    [Parameter(Position=0,mandatory=$true)]
    [ValidateScript({Test-Path -Path "$_" -PathType Leaf})]
    [string]$CSVFilePath,
    [Parameter(Position=1,mandatory=$false)]
    [ValidatePattern("[^\s]+")]
    [string]$ComputerName = $env:COMPUTERNAME
)

# If want compatibility with older versions, insert this shim (top-level or inside a function, where function name is returned): 
if ($PSCommandPath -eq $null) { function Get-PSCommandPath() { return $MyInvocation.PSCommandPath; } $PSCommandPath = Get-PSCommandPath; }

$Private:PrintManagementScript = ".\Add-PrintQueue.ps1"

if (-not (Test-Path -Path $Private:PrintManagementScript -PathType Leaf)) {
    throw [System.IO.FileNotFoundException] "Requires '$Private:PrintManagementScript' to be in same directory as $PSCommandPath"
    return
}

$csv_obj = Import-Csv -Path "$CSVFilePath" | foreach {
    New-Object PSObject -Property @{
        Name = $_.'Printer Name';
        Location = $_.Location;
        Comments = $_.Comments;
        IsShared = ([Bool]"$($_.'Is Shared')")
        PortHostAddress = $_.PortHostAddress
    }
}


$csv_obj | foreach {
    $QueueName = "$($_.Name)"
    $Comments = "$($_.Comments)"
    [string]$FQDN = "$($_.PortHostAddress)"

    if (-not (Test-Connection -ComputerName $FQDN -Count 3 -Quiet)) {
        Write-Warning -Message "$FQDN is unreachable!"
        $Comments += " | Config=BASIC PRINTING MODE, DUE TO UNREACHABLE AT TIME OF QUEUE CREATION"
    }

    if ($QueueName -like '*-SECURE*') { $QueueName = $FQDN.Split('.')[0] }
    [string]$command = "$Private:PrintManagementScript -ComputerName $ComputerName -HostName `'$QueueName`' -PortHostAddress `'$FQDN`' -MachineLocation `'$($_.Location)`' -Comment `'$Comments`'"
    if ("$($_.Name)" -like '*-SECURE*') { $command += ' -SecurePrint' }
    if ($_.IsShared -eq $true) { $command += ' -Shared' }
    
    Invoke-Expression -Command "$command"
}