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
        "HostName"
        "MyHomePC"
        "laptop"
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
    ���-�� ������� �� ���������� ����, ������� ���� �� i5 ���������� k=35: ����������� ����� ��� ������ �������� ������

.PARAMETER t
    ����� (� ��������) �������� ���������� ������� ����� �� ��������� ���������. ���������� ��� ������ �������, "��������" �� Get-WmiObject
    ������� (� �������) ��� ������ �� while-����� ����� ����������� �������

.EXAMPLE
    .\Get-WMISMART.ps1 $env:COMPUTERNAME
        �������� S.M.A.R.T. �������� ������ ���������� ����������

.EXAMPLE
    .\Get-WMISMART.ps1 HOST_NAME
        �������� S.M.A.R.T. �������� ���������� HOST_NAME

.EXAMPLE
    .\Get-WMISMART.ps1 .\input\example.csv -k 37 -t 13
        �������� S.M.A.R.T. �������� ������ ����������� ������ .\input\example.csv, ��������� maximum 37 ������� �� ����, ����� � 13 ������ � ������� 13 �����

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
    [string] $Out = ".\output\$('{0:yyyy-MM-dd}' -f $(Get-Date)) $($Inp.ToString().Split('\')[-1].Replace('.csv', '')) drives.csv",
    [int]    $k   = 37,
    [int]    $t   = 17  # once again timer
)


#region  # ������

$psCmdlet.ParameterSetName | Out-Null
# Clear-Host
$WatchDogTimer = [system.diagnostics.stopwatch]::startNew()
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # ��������� �������� ����� "./" = ������� ���������� �������
# [System.Security.Principal.WindowsIdentity]::GetCurrent().Name  # ���������� ��� set-credentials

#endregion


Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;LastScan = "";})}  # ��������, ��� � ���������� - ��� ����� ��� ����-������ ������

$clones = ($Computers | Group-Object -Property HostName | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group)  # �������� �� ���������
if ($clones -ne $null) {Write-Host "'$Inp' �������� ���������, ��� �������� ����� ��������� ����������", $clones.HostName -ForegroundColor Red -Separator "`n"}

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # ����� �� ������

$ComputersOnLine = @()
$DiskInfo = @()


#region Multi-Threading: ������������� �������� ����������� ����� �� ����

#region: ������������� ����

$x = $Computers.Count / [int] $env:NUMBER_OF_PROCESSORS
$max = if (($x + 1) -lt $k) {[int] $env:NUMBER_OF_PROCESSORS * ($x + 1) + 1} else {[int] $env:NUMBER_OF_PROCESSORS * $k}
$Pool = [RunspaceFactory]::CreateRunspacePool(1, $max)
$Pool.ApartmentState = "MTA"
$Pool.Open()
$RunSpaces = @()

#endregion


#region: ������-���� �������, ������� ����� ����������� � ������

$Payload = {Param ([string] $name = $env:COMPUTERNAME)

    Write-Debug $name -Debug

    for ($i = 0; $i -lt 2; $i++)
    {

        $WMIInfo = @()

        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)

        if ($ping)  # ��� ������� ping �������� ������
        {
            try {$Win32_DiskDrive = Get-WmiObject -ComputerName $name -class Win32_DiskDrive -ErrorAction Stop}
            catch {break}

            foreach ($Disk in $Win32_DiskDrive)
            {
                $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # � wql-������� ��������� '\', ������� ������� �� �� '_' (��� �������� "���� ����� ������"), ��. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

                # �����-��������, �����
                try {$WMIData = (Get-WmiObject -ComputerName $name -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific}
                catch {$WMIData = @()}

                if ($WMIData.Length -ne 512)
                {
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
                    ScanDate      =          $('{0:yyyy.MM.dd}' -f $(Get-Date))
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

            break  # � ��������� ���� ��������
        }
    }

    Return @(
        (New-Object psobject -Property @{
            HostName = $name
            LastScan = $('{0:yyyy.MM.dd}' -f $(Get-Date))
            Ping = if ($ping) {1} else {0}
            WMIInfo = $WMIInfo.Count
        }),
        $WMIInfo
    )
}

#endregion


#region: ��������� ������� � ��������� ������ � ���

foreach ($C in $Computers)
{
    $NewShell = [PowerShell]::Create()

    $null = $NewShell.AddScript($Payload)
    $null = $NewShell.AddArgument($C.HostName)

    $NewShell.RunspacePool = $Pool

    $RunSpace = [PSCustomObject]@{ Pipe = $NewShell; Status = $NewShell.BeginInvoke() }
    $RunSpaces += $RunSpace
}

#endregion


#region: ����� ���������� ������� ������ �������� ��� ������ � ���������, � ����� ���������� ���� ������� ��������� ���

$dctCompleted = @{}  # ������� ����������� �������
$dctHang = @{}  # ������������� �������
$total = $RunSpaces.Count  # ����� ���-�� �������

# � ������ ���������� ���������� ����� ����� ����������� �� ������ �����, � ����� �������� ������ ����� ��������� ��� ���� �� ���� ���
While ($RunSpaces.Status.IsCompleted -contains $false -or ($RunSpaces.Count -eq ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $true}).Count) )
{
    $wpCompl   = "�������� ���������� ping'�� � ���� S.M.A.R.T. ������,   �����ب���� ������"
    $wpNoCompl = "�������� ���������� ping'�� � ���� S.M.A.R.T. ������, �������ب���� ������"

    $c_true = $RunSpaces | Where-Object -FilterScript {$_.Status.IsCompleted -eq $true}  # ������� ����������� ������ ����������� �������
    $c_true_filtred = $c_true | Where-Object -FilterScript {!$dctCompleted.ContainsKey($_.Pipe.InstanceId.Guid)}  # ����� ������ ��, ������� ��� �� ������������, �.�. ������� ���� �� ������ � ������� $dctCompleted

    foreach ($RS in $c_true_filtred)  # ���� �� �����������_����, ������� ��� ��� � ������� (�����������_����)
    {
        $Result = @()

        if($RS.Status.IsCompleted -and !$dctCompleted.ContainsKey($RS.Pipe.InstanceId.Guid))
        {
            $Result = $RS.Pipe.EndInvoke($RS.Status)
            $RS.Pipe.Dispose()

            $ComputersOnLine += $Result[0]
            if ($Result[1].Count -gt 0) {$DiskInfo += $Result[1]}

            $dctCompleted[$RS.Pipe.InstanceId.Guid] = $WatchDogTimer.Elapsed.TotalSeconds

            $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������
            Write-Progress -id 1 -PercentComplete (100 * ($dctCompleted.Count) / $total) -Activity $wpCompl -Status "�����: $total" -CurrentOperation "������: $($dctCompleted.Count)"
            Write-Progress -id 2 -PercentComplete (100 * $p / $total) -Activity $wpNoCompl -Status "�����: $total" -CurrentOperation "��������: $p"

            if($dctHang.ContainsKey($RS.Pipe.InstanceId.Guid))
            {
                $dctHang.Remove($RS.Pipe.InstanceId.Guid)
            }
        }
    }


    # Start-Sleep -Milliseconds 100
    $c_false = $RunSpaces | Where-Object -FilterScript {$_.Status.IsCompleted -eq $false}

    foreach($RS in $c_false)  # ���� �� �������������
    {
        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������

        Write-Progress -id 1 -PercentComplete (100 * ($dctCompleted.Count) / $total) -Activity $wpCompl -Status "�����: $total" -CurrentOperation "������: $($dctCompleted.Count)"
        Write-Progress -id 2 -PercentComplete (100 * $p / $total) -Activity $wpNoCompl -Status "�����: $total" -CurrentOperation "��������: $p"

        $dctHang[$RS.Pipe.InstanceId.Guid] = $WatchDogTimer.Elapsed.TotalSeconds

        if ($c_false.Count -ne $p) {break}
    }

    # Start-Sleep -Milliseconds 100
    $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������

Write-Progress -id 2 -PercentComplete (100 * $p / $total) -Activity $wpNoCompl -Status "�����: $total" -CurrentOperation "��������: $p"

    Write-Host "timer:" $WatchDogTimer.Elapsed.TotalSeconds, "`tdctCompleted:", $dctCompleted.Count,  "`tdctHang:",$dctHang.Count,  "`t��������, `$p:",$p -ForegroundColor Yellow

Write-Progress -id 1 -PercentComplete (100 * ($dctCompleted.Count) / $total) -Activity $wpCompl -Status "�����: $total" -CurrentOperation "������: $($dctCompleted.Count)"

    #         ���-�� �������� �� ����������                                          �  (�������� ���� �� ���� � ����� ���� �� 2)
    $escape = ($p -eq $dctHang.Count) -and ($total -eq ($dctCompleted.Count + $p)) -and ( $(if ($total -gt $p) {$dctCompleted.Count -gt 0} else {$true}) )
    #                                   �  ����� = ����������_���� + ��_����������_����         ^���^ ������� �������� ��� ��� ���� ������� ���� ����� � �� �� ��� ���������

    if ($escape)
    {
        Start-Sleep -Seconds $t  # ����� ����� ��������� ���������� ������ � �� ����� ������ ��� "��������"

        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # ���-�� ������������� �������

        $escape = ($p -eq $dctHang.Count) -and ($total -eq ($dctCompleted.Count + $p)) -and ( $(if ($total -gt $p) {$dctCompleted.Count -gt 0} else {$true}) )

        if ($escape)
        {
            Write-Host '����� ', $t, '��� ���-�� ������������� ������� �������� �������. ����� �� Multi-Threading-�����...' -ForegroundColor Red
            Write-Host "timer:" $WatchDogTimer.Elapsed.TotalSeconds, "`tdctCompleted:", $dctCompleted.Count,  "`tdctHang:",$dctHang.Count,  "`t��������, `$p:",$p -ForegroundColor Magenta

            $RunSpaces | Where-Object -FilterScript {$_.Pipe.InstanceId.Guid -in ($dctHang.Keys)} | foreach {Write-Host $_.Pipe.Streams.Debug, $_.Pipe.Streams.Information, "`t", $_.Pipe.InstanceId.Guid -ForegroundColor Red}

            break  # ���������� while-�����: ���� ���-�� ������������� ������� ����� ����� (� ��������) �� ����������, ������� ��� �������)
        }

        else

        {
            Write-Host -ForegroundColor Green '����� �', $t, '��� �����������. �� ���, ��� ���... :-)'
        }
    }

    else

    {
        continue
    }


    if ($WatchDogTimer.Elapsed.TotalMinutes -gt $t)
    {
        break  # ���������� ���������� while-����� �� ��������, � �������
    }
}

<# ����� ��������� "��������" ������ (�������� ����� Job � ���������), ������ ��� ��������� ���
foreach($RS in  $RunSpaces | Where-Object -FilterScript {$dctHang.ContainsKey($_.Pipe.InstanceId.Guid)})
{
    $RS.Pipe.InstanceId.Guid
    $RS.Pipe.Streams
    # $RS.Pipe.Stop()  # hang
}

# $Pool.Close()
# $Pool.Dispose()
#>

#endregion


Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): Multi-Threading passed' -ForegroundColor Cyan

#endregion Multi-Threading


foreach($d in $DiskInfo){$d.SerialNumber = (Convert-hex2txt -wmisn $d.SerialNumber)}  # ��� ������������� �������� �������� ������ � �������� ������
Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): SerialNumber`s converted' -ForegroundColor Cyan


#region: ��������:  ����� �� ������, ���������� �������� ����� (��� �������������)
if ($DiskInfo.Count -gt 0)
{
    $DiskInfo | Select-Object `
        'HostName',`
        'ScanDate',`
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
        | Sort-Object -Property 'HostName' | Export-Csv -Path $Out -NoTypeInformation -Encoding UTF8 #-Delimiter ';' -Append


    if (Test-Path -Path $Inp)  # ���� �� ���� ��� ����� ���� �� ������� ������, �� ������������ ���� ������ �� ��������� ��-����\���-����
    {
        $ComputersOnLine += ($Computers | Where-Object -FilterScript {$_.HostName -notin $ComputersOnLine.HostName})  # | Select-Object -Property 'HostName')
        $ComputersOnLine | Select-Object 'HostName' | Sort-Object -Property 'HostName' | Export-Csv -Path $Inp -NoTypeInformation -Encoding UTF8
    }
}

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): Export-Csv completed' -ForegroundColor Cyan

#endregion


#region: ���������� ��

$DiskID = Get-DBHashTable -table 'Disk'
$HostID = Get-DBHashTable -table 'Host'

foreach ($Scan in $DiskInfo)  # 'HostName' 'ScanDate' 'SerialNumber' 'Model' 'Size' 'InterfaceType' 'MediaType' 'DeviceID' 'PNPDeviceID' 'WMIData' 'WMIThresholds' 'WMIStatus'
{
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
Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): DataBase updated' -ForegroundColor Cyan

#endregion


#region  # �����

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): executed' -ForegroundColor  Green
$WatchDogTimer.Stop()  # $WatchDogTimer.Elapsed.TotalSeconds

#endregion
