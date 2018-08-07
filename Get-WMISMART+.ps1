#requires -version 3  # ��������� ������ PowerShell

<#
.SYNOPSIS
    �������� � ������������� ������ ������ �� WMI S.M.A.R.T. ������ ������ ������ ���������� ����������(-��) � ��������� ����� � csv-�������

.DESCRIPTION
    ��� ���������� ������ �������� ���������� ��������� ��-��� �� �������������� �������� ����������

    ��������:
        ��������� ��������� ���� ���� ������ �����
        �������� ����� WMI ��������� S.M.A.R.T.-������ ������ ������
        ��������� ����� � '.\output\yyyy-MM-dd_HH-mm $inp drives.csv'

.INPUTS
    ��� ���������� � ������� "mylaptop"
        ���
    csv-���� ������ �����������
    ��������� ������ �����, ������ ������ �������� ���������
        "HostName","LastScan"
        "MyHomePC",""
        "laptop",""
    ����������� - ���-�� �������, ������� ����� �������� �� ������ ���������� ����

    � ������������ ���� "HostName" ����������� ����� �����������
    ���� "LastScan" ����� ���� ������, � ���� �� ��������� ������ ������� ����� �������� on-line/off-line ������ ����������

.OUTPUTS
    csv-���� � "������" S.M.A.R.T.-�������

.PARAMETER Inp
    ��� ���������� ��� ���� � csv-����� ������ �����������

.PARAMETER Out
    ���� � ����� ������

.PARAMETER k
    ���-�� ������� �� ���� ���������� ����, ������� ���� �� i5 ���������� k=35: ����������� ����� ��� ������ �������� ������

.EXAMPLE
    .\Get-WMISMART.ps1 $env:COMPUTERNAME
        �������� S.M.A.R.T. �������� ������ ���������� ����������

.EXAMPLE
    .\Get-WMISMART.ps1 HOST_NAME
        �������� S.M.A.R.T. �������� ���������� HOST_NAME

.EXAMPLE
    .\Get-WMISMART.ps1 .\input\example.csv
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
    [string] $Inp = ".\input\example.csv",  # ��� ����� ���� ���� � ����� ������ ������
    [string] $Out = ".\output\$('{0:yyyy-MM-dd_HH-mm}' -f $(Get-Date)) $($Inp.ToString().Split('\')[-1].Replace('.csv', '')) drives.csv",
    [int]    $k   = 35
)

#region  # ������
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = @(Get-Date) # ����� ������� ���������� �������
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # ��������� �������� ����� "./" = ������� ���������� �������
#endregion

Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;LastScan = "";})}  # ��������, ��� � ���������� - ��� ����� ��� ����-������ ������

$clones = ($Computers | Group-Object -Property HostName | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group)  # �������� �� ���������
if ($clones -ne $null) {Write-Host "'$Inp' �������� ������������� ����� �����������, ��� �������� ����� ��������� ����������", $clones.HostName -ForegroundColor Red -Separator "`n"}

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # ����� �� ������

$ComputersOnLine = @()
$DiskInfo = @()

#region Multi-Threading: ������������� �������� ����������� ����� �� ����
#region: ������������� ����
$Pool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * $k + 0)
$Pool.ApartmentState = "MTA"
$Pool.Open()
$RunSpaces = @()
#endregion

#region: ������-���� �������, ������� ����� ����������� � ������
$Payload = {Param ([string] $name = '127.0.0.1')

    for ($i = 0; $i -lt 2; $i++) {  # ������� ��� 'Test-Connection -Count 3'
        $WMIInfo = @()
        $ping = $false

        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)
        if ($ping) {

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
                $WMIInfo += New-Object psobject -Property @{
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
            break
        }
    }

    Return @((New-Object psobject -Property @{HostName = $name;LastScan = [string] (Get-Date);}), $WMIInfo)
}
#endregion

#region: ��������� ������� � ��������� ������ � ���
foreach ($C in $Computers) {
    $NewShell = [PowerShell]::Create()

    $null = $NewShell.AddScript($Payload)
    $null = $NewShell.AddArgument($C.HostName)

    $NewShell.RunspacePool = $Pool

    $RunSpaces += [PSCustomObject]@{ Pipe = $NewShell; Status = $NewShell.BeginInvoke() }
}
#endregion

#region: ����� ���������� ������� ������ �������� ��� ������ � ���������, � ����� ���������� ���� ������� ��������� ���
if ($RunSpaces.Count -gt 0) {  # fixed bug: � ���� �� �������� ������? - � ������ ������� input-����� ������ ������� �� ����� while, �.�. ���� ���������� ���� �� ������ ������, ���� �� ������ �� ����...
    while ($RunSpaces.Status.IsCompleted -notcontains $true) {}  # ��� ���������� ���� �� ������ ������ �������� ��������� ������

    $t = ($RunSpaces.Status).Count  # ����� ���-�� �������
    foreach ($RS in $RunSpaces ) {
        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "�������� ������� ����������� ���������� � ���� S.M.A.R.T. ������" -Status "�����: $t" -CurrentOperation "��������: $p"

        $Result = $RS.Pipe.EndInvoke($RS.Status)
        $ComputersOnLine += $Result[0]
        if ($Result[1].Count -gt 0) {$DiskInfo += $Result[1]}
        $RS.Pipe.Dispose()
    }
    $Pool.Close()
    $Pool.Dispose()
}
#endregion
#endregion Multi-Threading

#region: ���������� ��
foreach($d in $DiskInfo){$d.SerialNumber = (Convert-hex2txt -wmisn $d.SerialNumber)}  # ��� ������������� �������� �������� ������ � �������� ������

$DiskID = Get-DBHashTable -table 'Disk'
$HostID = Get-DBHashTable -table 'Host'
foreach ($Scan in $DiskInfo)
{  # 'ScanDate' 'HostName' 'SerialNumber' 'Model' 'Size' 'InterfaceType' 'MediaType' 'DeviceID' 'PNPDeviceID' 'WMIData' 'WMIThresholds' 'WMIStatus'

    # Disk
    if (!$DiskID.ContainsKey($Scan.SerialNumber))
    {
        $dID = Add-DBDisk -obj $Scan
        if($dID -gt 0) {$DiskID[$Scan.SerialNumber] = $dID}
    }

    # Host
    if ($HostID.ContainsKey($Scan.HostName))  # ���� � ������� � ������������ ��� ���� ������
    {  # update record
        $hID = Add-DBHost -obj $Scan -id $HostID[$Scan.HostName]
    }
    else
    {  # new record
        $hID = Add-DBHost -obj $Scan
        if ($hID -gt 0) {$HostID[$Scan.HostName] = $hID}
    }

    # Scan
    $sID = Add-DBScan -obj $Scan -dID $DiskID[$Scan.SerialNumber] -hID $HostID[$Scan.HostName]

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
if (Test-Path -Path $Inp) {$ComputersOnLine | Select-Object "HostName", "LastScan" | Export-Csv -Path $Inp -NoTypeInformation -Encoding UTF8}
#endregion

#region  # �����
# ����� ������� ���������� �������
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion
