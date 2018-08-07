#requires -version 3  # ��������� ������ PowerShell

<#
.SYNOPSIS

.DESCRIPTION

.INPUTS

.OUTPUTS

.PARAMETER Inp

.EXAMPLE

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

.NOTES
    Author: Dmitry Mikhaylov
#>

[CmdletBinding(DefaultParameterSetName="all")]

param
(
    [string] $dbname = 'ppsmart-posh.db',
    [string] $csvDir = 'output'
)

#region  # ������
$psCmdlet.ParameterSetName | Out-Null
#Clear-Host
$TimeStart = @(Get-Date) # ����� ������� ���������� �������
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # ��������� �������� ����� "./" = ������� ���������� �������
#endregion

Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

    if ([IntPtr]::Size -eq 8)
    {  # 64-bit
        $sqlite = Join-Path -Path $RootDir -ChildPath 'x64\System.Data.SQLite.dll'
    }
    elseif ([IntPtr]::Size -eq 4)
    {  # 32-bit
        $sqlite = Join-Path -Path $RootDir -ChildPath 'x32\System.Data.SQLite.dll'
    }
    else {Write-Host 'can not choose between 32 or 64 bit dll`s'}

try   {Add-Type -Path $sqlite -ErrorAction Stop}
catch {Write-Host "Importing the SQLite assemblies, '$sqlite', failed..."}

$db = Join-Path -Path $RootDir -ChildPath $dbname

$WMIFiles = Get-ChildItem -Path (Join-Path -Path $RootDir -ChildPath $csvDir) -Filter '*drives.csv' -Recurse

$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$db"
$con.Open()

$p1 = Measure-Command {
#region �������� ������� �� ����
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
$DiskID = @{}  # ��� �������� ID ������
$HostID = @{}  # ��� �������� ID ������
foreach ($r in $dataDisk.Tables.Rows) {$DiskID[$r.SerialNumber] = $r.ID}
foreach ($r in $dataHost.Tables.Rows) {$HostID[$r.HostName] = $r.ID}
#endregion
} | select -ExpandProperty TotalSeconds

$p2 = Measure-Command {
#region ������� ������ ������ �� CSV-������ � SQLite ��
$sqlHost = $con.CreateCommand()
foreach ($f in Get-ChildItem -Path (Join-Path -Path $RootDir -ChildPath 'input') -Filter '*.csv' -Recurse)
{
    foreach ($l in Import-Csv $f.FullName)
    {
        if (!$HostID.ContainsKey($l.HostName))
        {
            $sqlHost.CommandText = "INSERT OR IGNORE INTO `Host` (`HostName`,`LastScan`) VALUES (@HostName, @LastScan);"
            $null = $sqlHost.Parameters.AddWithValue("@HostName", $l.HostName)
            $null = $sqlHost.Parameters.AddWithValue("@LastScan", 0)
            if ($sqlHost.ExecuteNonQuery()) {$HostID[$l.HostName] = $sqlHost.Connection.LastInsertRowId}  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
        }
    }
}
$sqlHost.Dispose()
#endregion
} | select -ExpandProperty TotalSeconds

$p3 = Measure-Command {
#region ������� ����������� ������������ �� CSV-������ � SQLite ��
$sqlDisk = $con.CreateCommand()
$sqlHost = $con.CreateCommand()
$sqlScan = $con.CreateCommand()
foreach ($f in $WMIFiles) {
    foreach ($HardDrive in Import-Csv $f.FullName) {  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
        #region Disk
        $SerialNumber = (Convert-hex2txt -wmisn $HardDrive.SerialNumber)

        if (!$DiskID.ContainsKey($SerialNumber))
        {  # add new record
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
            $null = $sqlDisk.Parameters.AddWithValue("@SerialNumber", $SerialNumber)
            $null = $sqlDisk.Parameters.AddWithValue("@Model", $HardDrive.Model)
            $null = $sqlDisk.Parameters.AddWithValue("@Size", [int] $HardDrive.Size)
            $null = $sqlDisk.Parameters.AddWithValue("@InterfaceType", $HardDrive.InterfaceType)
            $null = $sqlDisk.Parameters.AddWithValue("@MediaType", $HardDrive.MediaType)
            $null = $sqlDisk.Parameters.AddWithValue("@DeviceID", $HardDrive.DeviceID)
            $null = $sqlDisk.Parameters.AddWithValue("@PNPDeviceID", $HardDrive.PNPDeviceID)

            if ($sqlDisk.ExecuteNonQuery()) {$DiskID[$SerialNumber] = $sqlDisk.Connection.LastInsertRowId}
        }
        #endregion

        #region Host
        if ($HostID.ContainsKey($HardDrive.HostName))  # ���� � ������� � ������������ ��� ���� ������
        {  # update record
            $sqlHost.CommandText = 'UPDATE `Host` SET LastScan = @LastScan WHERE ID = @ID'
            $null = $sqlHost.Parameters.AddWithValue("@LastScan", $HardDrive.ScanDate)
            $null = $sqlHost.Parameters.AddWithValue("@ID", $HostID[$HardDrive.HostName])
            $null = $sqlHost.ExecuteNonQuery()
        }
        else
        {  # add new record
            $sqlHost.CommandText = "INSERT OR IGNORE INTO `Host` (`HostName`, `LastScan`) VALUES (@HostName,@LastScan);"
            $null = $sqlHost.Parameters.AddWithValue("@HostName", $HardDrive.HostName)
            $null = $sqlHost.Parameters.AddWithValue("@LastScan", $HardDrive.ScanDate)
            if ($sqlHost.ExecuteNonQuery())  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
            {
                $HostID[$HardDrive.HostName] = $sqlHost.Connection.LastInsertRowId
            }
        }

        #endregion

        #region Scan
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
        $null = $sqlScan.Parameters.AddWithValue("@DiskID", $DiskID[$SerialNumber])
        $null = $sqlScan.Parameters.AddWithValue("@HostID", $HostID[$HardDrive.HostName])
        $null = $sqlScan.Parameters.AddWithValue("@ScanDate", $HardDrive.ScanDate)
        $null = $sqlScan.Parameters.AddWithValue("@WMIData", $HardDrive.WMIData)
        $null = $sqlScan.Parameters.AddWithValue("@WMIThresholds", $HardDrive.WMIThresholds)
        $null = $sqlScan.Parameters.AddWithValue("@WMIStatus", $(if($HardDrive.WMIStatus.ToLower() -eq 'true') {1} else {0}))
        $null = $sqlScan.ExecuteNonQuery()  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
        #endregion
    }
}  #endregion
$sqlDisk.Dispose()
$sqlHost.Dispose()
$sqlScan.Dispose()
$con.Close()
} | select -ExpandProperty TotalSeconds

#region  # �����
# ����� ������� ���������� �������
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion

Write-Host -ForegroundColor Green $p1, 'region �������� ������� �� ����'
Write-Host -ForegroundColor Green $p2, 'region ������� ������ ������ �� CSV-������ � SQLite ��'
Write-Host -ForegroundColor Green $p3, 'region ������� ����������� ������������ �� CSV-������ � SQLite ��'
