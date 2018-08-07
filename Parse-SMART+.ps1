#requires -version 3  # ��������� ������ PowerShell

<#
.SYNOPSIS
    �������������� � �������� ���������� .\Get-WMISMART.ps1

.DESCRIPTION
    �������� �������������� "�����" S.M.A.R.T. ������ �� ������� .\Get-WMISMART.ps1
    ��������� ��� ������� ������:
        _SMART_STABLE.csv - ���������� �����, ��������� �� �������� �� ������ � ������
        _SMART_DEGRADATION.csv - �����, � �������� ��������� ��������������� ��������

.INPUTS
    ����� � �������� �� ��������������� ������ ������

.OUTPUTS
    ����� � ���������� ���������� S.M.A.R.T.

.PARAMETER ReportDir
    ����� � ������������ ������ .\Get-WMISMART.ps1

.EXAMPLE
    .\Parse-SMART.ps1 .\output
        �������� S.M.A.R.T. �������� ������ ���������� ����������

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

    ������������� ���������
        https://3dnews.ru/618813
        http://www.ixbt.com/storage/hdd-smart-testing.shtml
        https://www.opennet.ru/base/sys/smart_hdd_mon.txt.html
        https://ru.wikipedia.org/wiki/S.M.A.R.T.

.NOTES
    Author: Dmitry Mikhaylov
#>

[CmdletBinding(DefaultParameterSetName="all")]
param
(
     [string] $ReportDir = '.\output'  #
)

#region  # ������
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = @(Get-Date) # ����� ������� ���������� �������
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # ��������� �������� ����� "./" = ������� ���������� �������
#endregion

Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

$WMIFiles = Get-ChildItem -Path $ReportDir -Filter '*drives.csv' -Recurse  # ������ �� ������
$AllInfo = @()  # ������ ���� �� ���� ������ �� ���� �������
$Degradation = @()  # ���������� �� 5-�� �������� (remap)
$Stable = @()  # ����������, ��� ���������� �� 5-�� �������� (remap)

#region  # ������ WMI-������ �� ��
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

# ����������� ���������� ������
$sql = $con.CreateCommand()
$sql.CommandText = @'
SELECT
	Host.HostName,
	Disk.Model,
	Disk.SerialNumber,
	Disk.Size,
	Scan.ScanDate,
	Scan.WMIData,
	Scan.WMIThresholds,
	Scan.WMIStatus
FROM `Scan`
INNER JOIN `Host` ON Scan.HostID = Host.ID
INNER JOIN `Disk` ON Scan.DiskID = Disk.ID
ORDER BY Scan.ID;
'@
$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
$data = New-Object System.Data.DataSet
[void]$adapter.Fill($data)
$sql.Dispose()
$con.Close()

Measure-Command {###########################################
# ������������ ���������� ������ $data.Tables.Rows.Count
foreach ($hd in $data.Tables.Rows) {
    $Disk = (New-Object PSObject -Property @{
        ScanDate = $hd.ScanDate
        PC       = $hd.HostName
        Model    = $hd.Model.Replace(' ATA Device', '').Replace(' SCSI Disk Device', '')
        SerNo    = $hd.SerialNumber  # Convert-hex2txt -wmisn ([string] $hd.SerialNumber.Trim())
        })

    foreach ($atr in ( Convert-WMIArrays -data $hd.WMIData -thresh $hd.WMIThresholds | where {$_.saIDDec -in (9,5,184,187,197,198,200)})) {  # �������������� �������� � ��-������ ��������� � ������� $Disk
        try {$Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop}  # �� ������ ���������� � ������
        catch {Write-Host "DUPLICATE: '$($hd.HostName)'" -ForegroundColor Red}
    }
    $AllInfo += $Disk
}
} | select -ExpandProperty TotalSeconds ###########################################
#endregion

#region  # ������ WMI-������ �� ������
<#foreach ($f in $WMIFiles) {
    foreach ($HardDrive in Import-Csv $f.FullName) {  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
        $Disk = (New-Object PSObject -Property @{
            ScanDate = $HardDrive.ScanDate
            PC       = $HardDrive.HostName
            Model    = $HardDrive.Model
            SerNo    = Convert-hex2txt -wmisn ([string] $HardDrive.SerialNumber.Trim())
            })
        foreach ($atr in Convert-WMIArrays -data $HardDrive.WMIData -thresh $HardDrive.WMIThresholds) {  # �������������� �������� � ��-������ ��������� � ������� $Disk
            try {$Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop}  # �� ������ ���������� � ������
            catch {Write-Host "DUPLICATE: '$($HardDrive.HostName)' in '$($f.Name)'" -ForegroundColor Red}
        }
        $AllInfo += $Disk
    }
}#>
#endregion

#region  # ����� ���������� �� 5-�� ��������
foreach ($g in $AllInfo | Sort-Object -Property PC,ScanDate | Group-Object -Property SerNo) {
    $g_ex = $g | Select-Object -ExpandProperty Group
    $g_5 = $g_ex | Group-Object -Property '5'

    # ���� ������ 'SerNo' > ������ �� '5', ��� ���������� �� 'remap'
    # if ($g.Count -gt 1) {"`v", $g, $g_5} # ��� ����������� ����� �������� ����� � ������� ����� �����
    if ($g.Count -eq $g_5.Count) {
        $Stable += $g_ex | Select-Object -Last 1
        # Write-Host ($g_ex | Select-Object -Property 'PC' -Unique).PC -ForegroundColor Green
        Continue
    }

    Write-Host ($g_ex | Select-Object -Property 'PC' -Unique).PC -ForegroundColor Magenta
#     $Degradation += $g_ex
    foreach ($r in $g_5) {
        $Degradation += ($r | Select-Object -ExpandProperty Group | Select-Object -First 1)
    }
}
#endregion#>

#region  # ����� Degradation
$ReportDegradation = Join-Path -Path $ReportDir -ChildPath '_SMART_DEGRADATION.csv'  # degradation, from all reports
if (Test-Path -Path $ReportDegradation) {Remove-Item -Path $ReportDegradation}
$Degradation | Select-Object -Property `
    'PC',`
    'Model',`
    'SerNo',`
    'ScanDate',`
    @{Expression = {$_.'9'};Name='9 Power-On Hours'},`
    @{Expression = {$_.'5'};Name='5 Reallocated Sectors Count'},`
    @{Expression = {$_.'184'};Name='184 End-to-End error / IOEDC'},`
    @{Expression = {$_.'187'};Name='187 Reported Uncorrectable Errors'},`
    @{Expression = {$_.'197'};Name='197 Current Pending Sector Count'},`
    @{Expression = {$_.'198'};Name='198 (Offline) Uncorrectable Sector Count'},`
    @{Expression = {$_.'200'};Name='200 Multi-Zone Error Rate / Write Error Rate (Fujitsu)'}`
| Sort-Object -Property `
    @{Expression = "PC"; Descending = $false}, `
    @{Expression = "5 Reallocated Sectors Count"; Descending = $false}, `
    @{Expression = "ScanDate"; Descending = $false} `
| Export-Csv -NoTypeInformation -Path $ReportDegradation
#endregion

#region  # ����� Stable
$ReportStable = Join-Path -Path $ReportDir -ChildPath '_SMART_STABLE.csv'  # full smart values, from all reports
if (Test-Path -Path $ReportStable) {Remove-Item -Path $ReportStable}
$Stable | Select-Object -Property `
    'PC',`
    'Model',`
    'SerNo',`
    'ScanDate',`
    @{Expression = {$_.'9'};Name='9 Power-On Hours'},`
    @{Expression = {$_.'5'};Name='5 Reallocated Sectors Count'},`
    @{Expression = {$_.'184'};Name='184 End-to-End error / IOEDC'},`
    @{Expression = {$_.'187'};Name='187 Reported Uncorrectable Errors'},`
    @{Expression = {$_.'197'};Name='197 Current Pending Sector Count'},`
    @{Expression = {$_.'198'};Name='198 (Offline) Uncorrectable Sector Count'},`
    @{Expression = {$_.'200'};Name='200 Multi-Zone Error Rate / Write Error Rate (Fujitsu)'}`
| Sort-Object -Property `
    @{Expression = '5 Reallocated Sectors Count'; Descending = $True},  `
    @{Expression = '184 End-to-End error / IOEDC'; Descending = $True},`
    @{Expression = '187 Reported Uncorrectable Errors'; Descending = $True},`
    @{Expression = '197 Current Pending Sector Count'; Descending = $True},`
    @{Expression = '198 (Offline) Uncorrectable Sector Count'; Descending = $True},`
    @{Expression = '200 Multi-Zone Error Rate / Write Error Rate (Fujitsu)'; Descending = $True} `
| Export-Csv -NoTypeInformation -Path $ReportStable
#endregion

#region  # �����
# ����� ������� ���������� �������
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion
