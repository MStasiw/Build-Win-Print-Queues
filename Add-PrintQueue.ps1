#Requires -Modules PrintManagement
[CmdletBinding()]
param (
    [Parameter(Position=0,mandatory=$false)]
    [ValidatePattern("[^\s]+")]
    [string]$ComputerName = $env:COMPUTERNAME,
    [Parameter(Position=1,mandatory=$true)]
    [ValidatePattern("[A-Za-z0-9-_]{2,15}")]
    [string]$HostName,
    [Parameter(mandatory=$false)]
    [ValidatePattern("[A-Za-z0-9-_]{2,15}")]
    [string]$QueueName,

    [Parameter(mandatory=$false,ParameterSetName="ByCustomPortHostAddress")]
    [ValidatePattern("[^\s]+")]
    [string]$PortHostAddress, # In case needs to be a different/unique DNS address or a IP address instead; instead of by Default = Fully Qualified Domain Name (FQDN)

    [Parameter(Position=2,mandatory=$false)]
    [ValidateNotNull()]
    [string]$MachineLocation,
    [Parameter(Position=3,mandatory=$false)]
    [ValidateNotNull()]
    [string]$Comment,
    [Parameter(Position=4,mandatory=$false)]
    [switch]$SecurePrint,
    [Parameter(mandatory=$false)]
    [switch]$Shared # In case want it shared; instead of by default not shared
)

[string]$DriverName='Xerox Global Print Driver PS'
<#
[string]$DCNum = Read-Host -Prompt "Please input DC number (4 digits)"
[string]$PtrNum = Read-Host -Prompt "Please input PTR number (3 digits)"
#>
[string]$HostName = $HostName.Trim().ToUpper()
[string]$PortName = $HostName

if ($QueueName) { [string]$QueueName = $QueueName.Trim().ToUpper() }
else { [string]$QueueName = $HostName }

if ($PortHostAddress) { [string]$PortHostAddress = $PortHostAddress.Trim().ToLower() }
#elseif ($PortHostAddress -eq $null) { [string]$PortHostAddress = "$($HostName.ToLower()).homeoffice.ca.wal-mart.com" }
else { [string]$PortHostAddress = "$($HostName.ToLower()).homeoffice.ca.wal-mart.com" }

if ($SecurePrint) { $QueueName = "$QueueName-SECURED" }

[string]$LocationField = $MachineLocation.Trim()
[string]$CommentField = $Comment.Trim()

[string[]]$RenderingModes = @('CSR', 'SSR')
if ($SecurePrint) { [string]$RenderingMode = $RenderingModes[1] } else { [string]$RenderingMode = $RenderingModes[0] }
<#
# Add-Printer -RenderingMode
# param value: SSR or CSR
#
# CSR (Client Side Rendering)  Use for all printers except secured
# SSR (Service Side Rendering)  USE THIS FOR -SECURED PRINT QUEUES Only
#>

Write-Output "Creating Print Queue `"$QueueName`" for `"$PortHostAddress`" ..."

# If Port already exists it will just throw exception on continue
Add-PrinterPort -ComputerName $ComputerName -Name $PortName -PrinterHostAddress $PortHostAddress -PortNumber 9100 -SNMP 1 -SNMPCommunity "public" -ErrorAction SilentlyContinue

# If printer already exists it will just throw exception and continue
Add-Printer -ComputerName $ComputerName -Name $QueueName -DriverName $DriverName -PortName $PortName -Datatype "RAW" -PrintProcessor "winprint" -RenderingMode $RenderingMode -ShareName $QueueName -Location $LocationField -Comment $CommentField

# Share the printer if $Shared parameter is included at command line #
if ($Shared) { Set-Printer -ComputerName $ComputerName -Name $QueueName -Shared $true }

# Enable spooling (buffer and queue print jobs): Advanced > "Spool print dpcuments sp program finishes printing faster" > "Start printing after last page is spooled"
Get-CimInstance -ComputerName $ComputerName -ClassName Win32_Printer -Filter "Name='$QueueName'" | Set-CimInstance -Property @{Queued="true"}