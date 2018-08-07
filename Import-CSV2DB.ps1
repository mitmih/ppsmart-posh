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

$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$db"
$con.Open()

#region ������� ������ �� CSV-������ � SQLite ��
$WMIFiles = Get-ChildItem -Path (Join-Path -Path $RootDir -ChildPath $csvDir) -Filter '*drives.csv' -Recurse
$DiskID = @{}  # ��� �������� ID ������
$HostID = @{}  # ��� �������� ID ������
foreach ($f in $WMIFiles) {
    foreach ($HardDrive in Import-Csv $f.FullName) {  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
        #region Disk
        $SerialNumber = (Convert-hex2txt -wmisn $HardDrive.SerialNumber)
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
        $null = $sqlDisk.Parameters.AddWithValue("@SerialNumber", $SerialNumber)
        $null = $sqlDisk.Parameters.AddWithValue("@Model", $HardDrive.Model)
        $null = $sqlDisk.Parameters.AddWithValue("@Size", [int] $HardDrive.Size)
        $null = $sqlDisk.Parameters.AddWithValue("@InterfaceType", $HardDrive.InterfaceType)
        $null = $sqlDisk.Parameters.AddWithValue("@MediaType", $HardDrive.MediaType)
        $null = $sqlDisk.Parameters.AddWithValue("@DeviceID", $HardDrive.DeviceID)
        $null = $sqlDisk.Parameters.AddWithValue("@PNPDeviceID", $HardDrive.PNPDeviceID)

        if ($sqlDisk.ExecuteNonQuery()) {$DiskID[$SerialNumber] = $sqlDisk.Connection.LastInsertRowId}
        $sqlDisk.Dispose()  #endregion

        #region Host
        $sqlHost = $con.CreateCommand()
        $sqlHost.CommandText = "INSERT OR IGNORE INTO `Host` (`HostName`, `LastScan`) VALUES (@HostName,@LastScan);"
        $null = $sqlHost.Parameters.AddWithValue("@HostName", $HardDrive.HostName)
        $null = $sqlHost.Parameters.AddWithValue("@LastScan", $HardDrive.ScanDate)
        if ($sqlHost.ExecuteNonQuery())  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
        {
            $HostID[$HardDrive.HostName] = $sqlHost.Connection.LastInsertRowId
        }
        $sqlHost.Dispose()  #endregion

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
        $null = $sqlScan.Parameters.AddWithValue("@DiskID", $DiskID[$SerialNumber])
        $null = $sqlScan.Parameters.AddWithValue("@HostID", $HostID[$HardDrive.HostName])
        $null = $sqlScan.Parameters.AddWithValue("@ScanDate", $HardDrive.ScanDate)
        $null = $sqlScan.Parameters.AddWithValue("@WMIData", $HardDrive.WMIData)
        $null = $sqlScan.Parameters.AddWithValue("@WMIThresholds", $HardDrive.WMIThresholds)
        $null = $sqlScan.Parameters.AddWithValue("@WMIStatus", $(if($HardDrive.WMIStatus.ToLower() -eq 'true') {1} else {0}))
        if ($sqlScan.ExecuteNonQuery()) {}  # 1 ���� ������ ������� ���������, 0 ���� ���� ������
        $sqlScan.Dispose()  #endregion
    }
}  #endregion
$con.Close()

#region  # �����
# ����� ������� ���������� �������
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion
