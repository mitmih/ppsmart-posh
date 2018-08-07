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

# ��������� �������� ����� ���������� ��� ����������� ���������� ����������
if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $RootDir -ChildPath 'x64\System.Data.SQLite.dll'}  # 64-bit
elseif ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $RootDir -ChildPath 'x32\System.Data.SQLite.dll'}  # 32-bit
else {Write-Host 'Hmmm... not 32 or 64 bit...'}

# ���������� ���������� ��� ������ � sqlite
try {Add-Type -Path $sqlite -ErrorAction Stop}
catch {Write-Host "Importing the SQLite assemblies, '$sqlite', failed..."}

$db = Join-Path -Path $RootDir -ChildPath 'ppsmart-posh.db'  # ���� � ��

# ��������� ���������� � ��
$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$db"
$con.Open()

# �� �� �������� ID � ��������� ������
$sqlDisk = $con.CreateCommand()
$sqlDisk.CommandText = "SELECT Disk.ID, Disk.SerialNumber FROM Disk;"
$adapterDisk = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sqlDisk
$dataDisk = New-Object System.Data.DataSet
[void]$adapterDisk.Fill($dataDisk)
$sqlDisk.Dispose()

# �� �� �������� ID � ����� ������
$sqlHost = $con.CreateCommand()
$sqlHost.CommandText = "SELECT Host.ID, Host.HostName FROM Host;"
$adapterHost = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sqlHost
$dataHost = New-Object System.Data.DataSet
[void]$adapterHost.Fill($dataHost)
$sqlHost.Dispose()

# ������� � ������� ����� � �����, ����� ������ ��������� �� ��
$dctDisk = @{}
$dctHost = @{}
foreach ($r in $dataDisk.Tables.Rows) {$dctDisk[$r.SerialNumber] = $r.ID}
foreach ($r in $dataHost.Tables.Rows) {$dctHost[$r.HostName] = $r.ID}

foreach ($d in $DiskInfo) {  # 'ScanDate' 'HostName' 'SerialNumber' 'Model' 'Size' 'InterfaceType' 'MediaType' 'DeviceID' 'PNPDeviceID' 'WMIData' 'WMIThresholds' 'WMIStatus'

    #region Disk # ���� ����� ��� � �� - ������ � ���������� ID
    if (!($dctDisk.ContainsKey($d.SerialNumber))) {
        $sqlDisk = $con.CreateCommand()
        $sqlDisk.CommandText = @'
            INSERT OR IGNORE INTO `Disk` (
                `SerialNumber`,
                `Model`,
                `Size`,
                `InterfaceType`,
                `MediaType`,
                `DeviceID`,
                `PNPDeviceID`
            ) VALUES (
                @SerialNumber,
                @Model,
                @Size,
                @InterfaceType,
                @MediaType,
                @DeviceID,
                @PNPDeviceID
            );
'@
        $null = $sqlDisk.Parameters.AddWithValue("@SerialNumber", $d.SerialNumber)
        $null = $sqlDisk.Parameters.AddWithValue("@Model", $d.Model)
        $null = $sqlDisk.Parameters.AddWithValue("@Size", [int] $d.Size)
        $null = $sqlDisk.Parameters.AddWithValue("@InterfaceType", $d.InterfaceType)
        $null = $sqlDisk.Parameters.AddWithValue("@MediaType", $d.MediaType)
        $null = $sqlDisk.Parameters.AddWithValue("@DeviceID", $d.DeviceID)
        $null = $sqlDisk.Parameters.AddWithValue("@PNPDeviceID", $d.PNPDeviceID)

        if ($sqlDisk.ExecuteNonQuery()) {$dctDisk[$d.SerialNumber] = $sqlDisk.Connection.LastInsertRowId}
        $sqlDisk.Dispose()
    }  #endregion

    #region Host # ���� ����� ��� � �� - ������ � ���������� ID
    if (!($dctHost.ContainsKey($d.HostName)))
    {
        $sqlHost = $con.CreateCommand()
        $sqlHost.CommandText = "INSERT OR IGNORE INTO `Host` (`HostName`,`LastScan`) VALUES (@HostName, @LastScan);"
        $null = $sqlHost.Parameters.AddWithValue("@HostName", $d.HostName)
        $null = $sqlHost.Parameters.AddWithValue("@LastScan", (Get-Date))
        if ($sqlHost.ExecuteNonQuery()) {$dctHost[$d.HostName] = $sqlHost.Connection.LastInsertRowId}  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
        $sqlHost.Dispose()
    }
    else
    {
        # ���� ���� � ��, ����� �������� ��� ������
        $sqlHost = $con.CreateCommand()
        $sqlHost.CommandText = 'UPDATE `Host` SET LastScan = @LastScan WHERE ID = @ID'
        $null = $sqlHost.Parameters.AddWithValue("@LastScan", (Get-Date))
        $null = $sqlHost.Parameters.AddWithValue("@ID", $dctHost[$d.HostName])
        $null = $sqlHost.ExecuteNonQuery() #) {$dctHost[$d.HostName] = $sqlHost.Connection.LastInsertRowId}  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
        $sqlHost.Dispose()
    }
    #endregion

    #region Scan
    $sqlScan = $con.CreateCommand()
    $sqlScan.CommandText = @'
        INSERT OR IGNORE INTO `Scan` (
            `DiskID`,
            `HostID`,
            `ScanDate`,
            `WMIData`,
            `WMIThresholds`,
            `WMIStatus`
        ) VALUES (
            @DiskID,
            @HostID,
            @ScanDate,
            @WMIData,
            @WMIThresholds,
            @WMIStatus
        );
'@
    $null = $sqlScan.Parameters.AddWithValue("@DiskID", $dctDisk[$d.SerialNumber])
    $null = $sqlScan.Parameters.AddWithValue("@HostID", $dctHost[$d.HostName])
    $null = $sqlScan.Parameters.AddWithValue("@ScanDate", $d.ScanDate)
    $null = $sqlScan.Parameters.AddWithValue("@WMIData", $d.WMIData)
    $null = $sqlScan.Parameters.AddWithValue("@WMIThresholds", $d.WMIThresholds)
    $null = $sqlScan.Parameters.AddWithValue("@WMIStatus", $d.WMIStatus)
    $null = $sqlScan.ExecuteNonQuery()  #) {}  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
    $sqlScan.Dispose()  #endregion
}
$con.Close()
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

<#
UPDATE `Host`
SET LastScan = "1"
WHERE
    Host.HostName = "qwe"
#>
