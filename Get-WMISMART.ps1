#requires -version 3  # ��������� ������ PowerShell

<#
.SYNOPSIS
    �������� ������ �� WMI S.M.A.R.T. ������ ������ ������ ���������� ����������(-��) � ��������� ����� � csv-�������

.DESCRIPTION
    ��� ���������� ������ �������� ���������� ��������� ��-��� �� �������������� �������� ����������

    ��������:
        ��������� ��������� ���� ���� ������ �����
        �������� ����� WMI ��������� S.M.A.R.T.-������ ������ ������
        ��������� ����� � .\output\yyyy-MM-dd_HH-mm-ss.csv

.INPUTS
    ��� ���������� � ������� "mylaptop"
        ���
    csv-���� ������ �����������
    ��������� ������ �����, ������ ������ �������� ���������
        "HostName","Status"
        "MyHomePC",""
        "laptop",""

    � ������������ ���� "HostName" ����������� ����� �����������
    ���� "Status" ����� ���� ������, � ���� �� ��������� ������ ������� ����� �������� on-line/off-line ������ ����������

.OUTPUTS
    csv-���� � "������" S.M.A.R.T.-�������

.PARAMETER Inp
    ��� ���������� ��� ���� � csv-����� ������ �����������

.PARAMETER Out
    ���� � ����� ������

.EXAMPLE
    .\Get-WMISMART.ps1 $env:COMPUTERNAME
        �������� S.M.A.R.T. �������� ������ ���������� ����������

.EXAMPLE
    .\Get-WMISMART.ps1 HOST_NAME
        �������� S.M.A.R.T. �������� ���������� HOST_NAME

.EXAMPLE
    .\Get-WMISMART.ps1 .\input\hostname_list.csv
        �������� S.M.A.R.T. �������� ������ ����������� �� ������ hostname_list.csv

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

    �� ����� ���������� �������� ��������� 512-���� ������� S.M.A.R.T.: ��� ��������, ��� 1� ���� ���������� ����� � 0�� �����, � �� ����� ���� ������ ��� ����� �������� ������ ��������� S.M.A.R.T.
        https://social.msdn.microsoft.com/Forums/en-US/af01ce5d-b2a6-4442-b229-6bb32033c755/using-wmi-to-get-smart-status-of-a-hard-disk?forum=vbgeneral

.NOTES
    Author: Dmitry Mikhaylov
#>

[CmdletBinding(DefaultParameterSetName="all")]

param
(
#     [string] $Inp = "$env:COMPUTERNAME",  # ��� ����� ���� ���� � ����� ������ ������
    [string] $Inp = ".\input\hostname_list.csv",  # ��� ����� ���� ���� � ����� ������ ������
#     [string] $Out = ".\output\$($Inp.ToString().Split('\')[-1].Replace('.csv', '')) $('{0:yyyy-MM-dd_HH}' -f $(Get-Date))_drives.csv"
    [string] $Out = ".\output\$($Inp.ToString().Split('\')[-1].Replace('.csv', '')) $('{0:yyyy-MM-dd_HH-mm-ss}' -f $(Get-Date))_drives.csv"
)
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = Get-Date # ����� ������� ���������� �������
Set-Location "$($MyInvocation.MyCommand.Definition | Split-Path -Parent)"  # ��������� �������� ����� "./" = ������� ���������� �������

Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;Status = "";})}  # ��������, ��� � ���������� - ��� ����� ��� ����-������ ������

$clones = ($Computers | Group-Object -Property HostName | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group)  # �������� �� ���������
if ($clones -ne $null) {Write-Host "'$Inp' �������� ������������� ����� �����������, ��� �������� ����� ��������� ����������", $clones.HostName -ForegroundColor Red -Separator "`n"}

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # ����� �� ������

$ComputersOnLine = @()
$DiskInfo = @()

# Multy-Threading: ������������� �������� ����������� ����� �� ����
#region: Ping, WMI: ������ ��� �������
$PingPool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * 1 + 1)
$PingPool.ApartmentState = "MTA"
$PingPool.Open()
$PingRunSpaces = @()

$WMIPool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * 1 + 1)
$WMIPool.ApartmentState = "MTA"
$WMIPool.Open()
$WMIRunSpaces = @()
#endregion

#region: Ping, WMI: ������-����� ��� ��������� ����������
$PingPayload = {Param ([string] $name = '127.0.0.1')
    for ($i = 0; $i -lt 3; $i++) {  # ������� ��� 'Test-Connection -Count 3'
        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)
        if ($ping) {break}
    }

    Return (New-Object psobject -Property @{HostName = $name;Status = $ping;})
}

$WMIPayLoad = {Param ([string] $name = '127.0.0.1')
    try {$Win32_DiskDrive = Get-WmiObject -ComputerName $name -class Win32_DiskDrive -ErrorAction Stop}
    catch {$Win32_DiskDrive = @()}

    foreach ($Disk in $Win32_DiskDrive) {
        $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # � wql-������� ��������� '\', ������� ������� �� �� '_' (��� �������� "���� ����� ������"), ��. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

        # �����-��������, �����
        try {$WMIData = (Get-WmiObject -ComputerName $name -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific}
        catch {$WMIData = @()}
        if ($WMIData.Length -ne 512) {
            Write-Host "`t", $Disk.Model, "- � WMI ��� ������ S.M.A.R.T." -ForegroundColor DarkYellow
            continue
        }  # ���� ������ �� ��������, �� ����� ������ WMI ��� ������ ���������, ������� � ���������� ����� �����

        # ��������� ��������
        try {$WMIThresholds = (Get-WmiObject -ComputerName $name -namespace root\wmi -Class MSStorageDriver_FailurePredictThresholds -Filter $wql -ErrorAction Stop).VendorSpecific}
        catch {$WMIThresholds = @()}

        # ������ ����� (Windows OS IMHO)
        try {$WMIStatus = (Get-WmiObject -ComputerName $name -namespace root\wmi �class MSStorageDriver_FailurePredictStatus -Filter $wql -ErrorAction Stop).PredictFailure}  # ������ (TRUE), ���� �������������� ���� �����. � ���� ������ ����� ���������� ��������� ��������� ����������� �����
        catch {$WMIStatus = $null}

        # ��������� ����� ���� � ������ ������ �� ������
        Import-Module -Name ".\helper.psm1" -verbose
        $DiskInfo = New-Object psobject -Property @{
            ScanDate      =          Get-Date
            HostName      = [string] $name
            SerialNumber  = [string] $Disk.SerialNumber.Trim()  # Convert-hex2txt -wmisn ([string] $Disk.SerialNumber)  # �� �������� ������ ������ � ������-�����
            Model         = [string] $Disk.Model
            Size          = [System.Math]::Round($Disk.Size / (1000 * 1000 * 1000),0)
            InterfaceType = [string] $Disk.InterfaceType
            MediaType     = [string] $Disk.MediaType
            DeviceID      = [string] $Disk.DeviceID
            PNPDeviceID   = [string] $Disk.PNPDeviceID
            WMIData       = [string] $WMIData
            WMIThresholds = [string] $WMIThresholds
            WMIStatus     = [boolean]$WMIStatus
        }
    }

    Return $DiskInfo
}
#endregion

#region: Ping:      ��������� � ��������� ������ � ���
foreach ($C in $Computers) {
    $PingNewShell = [PowerShell]::Create()

    $null = $PingNewShell.AddScript($PingPayload)
    $null = $PingNewShell.AddArgument($C.HostName)

    $PingNewShell.RunspacePool = $PingPool

    $PingRunSpaces += [PSCustomObject]@{ Pipe = $PingNewShell; Status = $PingNewShell.BeginInvoke() }
}
#endregion

#region: Ping:      ����� ���������� ������ �������� ������ � ���������, ��� ��������� ����� ���������� ���� �������
if ($PingRunSpaces.Count -gt 0) {  # fixed bug: � ���� �� �������� ������? - � ������ ������� input-����� ������ ������� �� ����� while, �.�. ���� ���������� ���� �� ������ ������, ���� �� ������ �� ����...
    while ($PingRunSpaces.Status.IsCompleted -notcontains $true) {}  # ����� ���������� ���� �� ������ ������ �������� ��������� ������

    $t = ($PingRunSpaces.Status).Count  # ����� ���-�� �������
    foreach ($RS in $PingRunSpaces ) {
        $p = ($PingRunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "�������� ������� ����������� ����������" -Status "�����: $t" -CurrentOperation "��������: $p"

        $ComputersOnLine += $RS.Pipe.EndInvoke($RS.Status)
        $RS.Pipe.Dispose()
    }
    $PingPool.Close()
    $PingPool.Dispose()
}
#endregion

#region: WMI:       ��������� � ��������� ������ � ���
foreach ($C in $ComputersOnLine | Where-Object {$_.Status -eq 'online'}) {
    $WMINewShell = [PowerShell]::Create()

    $null = $WMINewShell.AddScript($WMIPayload)
    $null = $WMINewShell.AddArgument($C.HostName)

    $WMINewShell.RunspacePool = $WMIPool

    $WMIRunSpaces += [PSCustomObject]@{ Pipe = $WMINewShell; Status = $WMINewShell.BeginInvoke() }
}
#endregion

#region: WMI:       ����� ���������� ������ �������� ������ � ���������, ��� ��������� ����� ���������� ���� �������
if ($WMIRunSpaces.Count -gt 0) {  # fixed bug: � ���� �� �������� ������? - � ������ ������� input-����� ������ ������� �� ����� while, �.�. ���� ���������� ���� �� ������ ������, ���� �� ������ �� ����...
    while ($WMIRunSpaces.Status.IsCompleted -notcontains $true) {}  # ����� ���������� ���� �� ������ ������ �������� ��������� ������

    $t = ($WMIRunSpaces.Status).Count  # ����� ���-�� �������
    foreach ($RS in $WMIRunSpaces ) {
        $p = ($WMIRunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "��������� �� WMI ������ �� ������ ������" -Status "�����: $t" -CurrentOperation "��������: $p"

        $DiskInfo += $RS.Pipe.EndInvoke($RS.Status)
        $RS.Pipe.Dispose()
    }
    $WMIPool.Close()
    $WMIPool.Dispose()
}
#endregion

#region: ��������:  ����� �� ������, ���������� �������� ����� (��� �������������)
$DiskInfo | Select-Object `
    'ScanDate',`
    'HostName',`
    'SerialNumber',`
    'Model',`
    'Size',`
    'InterfaceType',`
    'MediaType',`
    'DeviceID',`
    'PNPDeviceID',`
    'WMIData',`
    'WMIThresholds',`
    'WMIStatus'`
    | Export-Csv -Path $Out -NoTypeInformation -Encoding UTF8 #-Delimiter ';' -Append

# ���� �� ���� ��� ����� ���� �� ������� ������, �� ������������ ���� ������ �� ��������� ��-����\���-����
if (Test-Path -Path $Inp) {$ComputersOnLine | Select-Object "HostName", "Status" | Export-Csv -Path $Inp -NoTypeInformation -Encoding UTF8}
#endregion

# ����� ������� ���������� �������
$ExecTime = [System.Math]::Round($( $(Get-Date) - $TimeStart ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
