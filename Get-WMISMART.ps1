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
    .\Get-WMISMART.ps1 .\input\debug.csv
        �������� S.M.A.R.T. �������� ������ ����������� �� ������ debug.csv

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
    [string] $Inp = ".\input\debug.csv",  # ��� ����� ���� ���� � ����� ������ ������
    [string] $Out = ".\output\$($Inp.ToString().Split('\')[-1].Replace('.csv', '')) $('{0:yyyy-MM-dd_HH}' -f $(Get-Date))_drives.csv"
#     [string] $Out = ".\output\$($Inp.ToString().Split('\')[-1].Replace('.csv', '')) $('{0:yyyy-MM-dd_HH-mm-ss}' -f $(Get-Date))_drives.csv"
)
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = Get-Date # ����� ������� ���������� �������
set-location "$($MyInvocation.MyCommand.Definition | split-path -parent)"  # ��������� �������� ����� "./" = ������� ���������� �������

Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;Status = "";})}  # ��������, ��� � ���������� - ��� ����� ��� ����-������ ������

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # ����� �� ������

# Multy-Threading: ������������� �������� ����������� ����� �� ����
#region ������ ��� runspace`��
$PingPool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * 1 + 1)
$PingPool.ApartmentState = "MTA"
$PingPool.Open()
$PingRunSpaces = @()
$ComputersOnLine = @()
#endregion

#region ��� ����� ��������������
$PingPayload = {Param ([string] $name = '127.0.0.1')
    $ping = $false
    for ($i = 0; $i -lt 3; $i++) {
        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)
        if ($ping) {break}
    }

    Return (New-Object psobject -Property @{HostName = $name;Status = $ping;})
}
#endregion

#region ������ runspace`�, ��������� � ��������� �� � ���
foreach ($C in $Computers) {
    $PingNewShell = [PowerShell]::Create()

    $null = $PingNewShell.AddScript($PingPayload)
    $null = $PingNewShell.AddArgument($C.HostName)

    $PingNewShell.RunspacePool = $PingPool

    $PingRunSpaces += [PSCustomObject]@{ Pipe = $PingNewShell; Status = $PingNewShell.BeginInvoke() }
}
#endregion

#region ����� ���������� ���� ���������� ������� �������� ������, ��������� ������ � ���
if ($PingRunSpaces.Count -gt 0) {  # � ���� �� �������� ������? - ��� ���: ���� �� ������� ��� ����, �� ������ �����, � ������ ������� �� ����� while, ������-��� ���� ���������� ���� �� ������ ������, � ������� �� � �� ����...
    while ($PingRunSpaces.Status.IsCompleted -notcontains $true) {}  # ����� ���������� ���� �� ������ ������ ��������� ��� ������

    foreach ($RS in $PingRunSpaces ) {
        $t = ($PingRunSpaces.Status).Count  # ����� ���-�� �������
        $p = ($PingRunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "�������� ������� ����������� ����������" -Status "�����: $t" -CurrentOperation "��������: $p"

        $ComputersOnLine += $RS.Pipe.EndInvoke($RS.Status)

        $RS.Pipe.Dispose()
    }
#     $PingRunSpaces.Status
    $PingPool.Close()
    $PingPool.Dispose()
}
#endregion

$DiskInfo = @()
if ($ComputersOnLine.HostName -is [array]) {$t = $ComputersOnLine.HostName.Length} else {$t = 1}
$p = 0  # ��������-���: $p - ������� (������� ����), $t - ����� ���������� (����� ������ � ���������)
foreach ($C in $ComputersOnLine | Where-Object {$_.Status -eq 'online'}) {
    $p += 1
    Write-Progress -id 2 -PercentComplete (100 * $p / $t) -Activity $Inp -Status "����������� $p-� ��������� �� $t" -CurrentOperation $C.HostName

    Write-Host "`n$($C.HostName) $($C.Status)" -ForegroundColor Green

    try {$Win32_DiskDrive = Get-WmiObject -ComputerName $C.HostName -class Win32_DiskDrive -ErrorAction Stop}
    catch {$Win32_DiskDrive = @()}

    foreach ($Disk in $Win32_DiskDrive) {
        $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # � wql-������� ��������� '\', ������� ������� �� �� '_' (��� �������� "���� ����� ������"), ��. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

        # �����-��������, �����
        try {$WMIData = (Get-WmiObject -ComputerName $C.HostName -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific}
        catch {$WMIData = @()}
        if ($WMIData.Length -ne 512) {
            Write-Host "`t", $Disk.Model, "- � WMI ��� ������ S.M.A.R.T." -ForegroundColor DarkYellow
            continue
        }  # ���� ������ �� ��������, �� ����� ������ WMI ��� ������ ���������, ������� � ���������� ����� �����

        # ��������� ��������
        try {$WMIThresholds = (Get-WmiObject -ComputerName $C.HostName -namespace root\wmi -Class MSStorageDriver_FailurePredictThresholds -Filter $wql -ErrorAction Stop).VendorSpecific}
        catch {$WMIThresholds = @()}

        # ������ ����� (Windows OS IMHO)
        try {$WMIStatus = (Get-WmiObject -ComputerName $C.HostName -namespace root\wmi �class MSStorageDriver_FailurePredictStatus -Filter $wql -ErrorAction Stop).PredictFailure}  # ������ (TRUE), ���� �������������� ���� �����. � ���� ������ ����� ���������� ��������� ��������� ����������� �����
        catch {$WMIStatus = $null}

        Write-Host "`t", $Disk.Model, $WMIStatus, $WMIData.Length, $WMIThresholds.Length -ForegroundColor Cyan

        # ��������� ����� ���� � ������ ������ �� ������
        $DiskInfo += (New-Object psobject -Property @{
            ScanDate      =          $TimeStart
            HostName      = [string] $C.HostName
            SerialNumber  = Convert-hex2txt -wmisn ([string] $Disk.SerialNumber)  #.Trim()
            Model         = [string] $Disk.Model
            Size          = [System.Math]::Round($Disk.Size / (1000 * 1000 * 1000),0)
            InterfaceType = [string] $Disk.InterfaceType
            MediaType     = [string] $Disk.MediaType
            DeviceID      = [string] $Disk.DeviceID
            PNPDeviceID   = [string] $Disk.PNPDeviceID
            WMIData       = [string] $WMIData
            WMIThresholds = [string] $WMIThresholds
            WMIStatus     = [boolean]$WMIStatus
        })
    }
}

# ������� ������ �� ������
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

# ����� ������� ���������� �������
$ExecTime = [System.Math]::Round($( $(Get-Date) - $TimeStart ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
