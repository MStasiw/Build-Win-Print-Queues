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

$csv_obj = Import-Csv -Path "$CSVFilePath"  | foreach {
    New-Object PSObject -Property @{
        Name = $_.'Printer Name';
        Location = $_.Location;
        Comments = $_.Comments;
        #[Bool]'IsShared' = if ($_.'Is Shared' -eq 'TRUE') { $true } else { $false }
        IsShared = $_.'Is Shared'
    }
}

#$csv_obj; Write-Output ''

$csv_obj | foreach {
    #[string]$command = "BEFORE $($_.Comments) AFTER"
    [string]$FQDN = "$($_.Name).homeoffice.ca.wal-mart.com"

    if (-not (Test-Connection -ComputerName $FQDN -Count 3 -Quiet)) { Write-Warning -Message "$FQDN is unreachable!" }

    [string]$command = "$Private:PrintManagementScript -ComputerName $ComputerName -HostName `'$($_.'Name')`' -MachineLocation `'$($_.Location)`' -Comment `'$($_.Comments)`'"
    if ($_.IsShared -eq 'TRUE') { $command += ' -Shared' }
        
    Invoke-Expression -Command $command
}