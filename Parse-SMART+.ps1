#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    расшифровывает и обобщает результаты .\Get-WMISMART.ps1

.DESCRIPTION
    сценарий расшифровывает "сырые" S.M.A.R.T. данные из отчётов .\Get-WMISMART.ps1
    формирует два сводных отчёта:
        _SMART_STABLE.csv - стабильные диски, состояние не меняется от отчёта к отчёту
        _SMART_DEGRADATION.csv - диски, с растущим значением переназначенных секторов

.INPUTS
    папка с отчётами по отсканированным жёстким дискам

.OUTPUTS
    отчёт с подробными атрибутами S.M.A.R.T.

.PARAMETER ReportDir
    папка с результатами работы .\Get-WMISMART.ps1

.EXAMPLE
    .\Parse-SMART.ps1 .\output
        получить S.M.A.R.T. атрибуты дисков локального компьютера

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

    интерпретация атрибутов
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

#region  # НАЧАЛО
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = @(Get-Date) # замер времени выполнения скрипта
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта
#endregion

Import-Module -Name ".\helper.psm1" -verbose  # вспомогательный модуль с функциями

$WMIFiles = Get-ChildItem -Path $ReportDir -Filter '*drives.csv' -Recurse  # отчёты по дискам
$AllInfo = @()  # полная инфа по всем дискам из всех отчётов
$Degradation = @()  # деградация по 5-му атрибуту (remap)
$Stable = @()  # стабильные, без деградации по 5-му атрибуту (remap)

#region  # читаем WMI-отчёты из БД
# проверяем битность среды выполнения для подключения подходящей библиотеки
if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $RootDir -ChildPath 'x64\System.Data.SQLite.dll'}  # 64-bit
elseif ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $RootDir -ChildPath 'x32\System.Data.SQLite.dll'}  # 32-bit
else {Write-Host 'Hmmm... not 32 or 64 bit...'}

# подключаем библиотеку для работы с sqlite
try {Add-Type -Path $sqlite -ErrorAction Stop}
catch {Write-Host "Importing the SQLite assemblies, '$sqlite', failed..."}

$db = Join-Path -Path $RootDir -ChildPath 'ppsmart-posh.db'  # путь к БД

# открываем соединение с БД
$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$db"
$con.Open()

# запрашиваем результаты сканов
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
# обрабатываем результаты сканов $data.Tables.Rows.Count
foreach ($hd in $data.Tables.Rows) {
    $Disk = (New-Object PSObject -Property @{
        ScanDate = $hd.ScanDate
        PC       = $hd.HostName
        Model    = $hd.Model.Replace(' ATA Device', '').Replace(' SCSI Disk Device', '')
        SerNo    = $hd.SerialNumber  # Convert-hex2txt -wmisn ([string] $hd.SerialNumber.Trim())
        })

    foreach ($atr in ( Convert-WMIArrays -data $hd.WMIData -thresh $hd.WMIThresholds | where {$_.saIDDec -in (9,5,184,187,197,198,200)})) {  # расшифровываем атрибуты и по-одному добавляем к объекту $Disk
        try {$Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop}  # на случай дубликатов в отчёте
        catch {Write-Host "DUPLICATE: '$($hd.HostName)'" -ForegroundColor Red}
    }
    $AllInfo += $Disk
}
} | select -ExpandProperty TotalSeconds ###########################################
#endregion

#region  # читаем WMI-отчёты из файлов
<#foreach ($f in $WMIFiles) {
    foreach ($HardDrive in Import-Csv $f.FullName) {  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
        $Disk = (New-Object PSObject -Property @{
            ScanDate = $HardDrive.ScanDate
            PC       = $HardDrive.HostName
            Model    = $HardDrive.Model
            SerNo    = Convert-hex2txt -wmisn ([string] $HardDrive.SerialNumber.Trim())
            })
        foreach ($atr in Convert-WMIArrays -data $HardDrive.WMIData -thresh $HardDrive.WMIThresholds) {  # расшифровываем атрибуты и по-одному добавляем к объекту $Disk
            try {$Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop}  # на случай дубликатов в отчёте
            catch {Write-Host "DUPLICATE: '$($HardDrive.HostName)' in '$($f.Name)'" -ForegroundColor Red}
        }
        $AllInfo += $Disk
    }
}#>
#endregion

#region  # поиск деградаций по 5-му атрибуту
foreach ($g in $AllInfo | Sort-Object -Property PC,ScanDate | Group-Object -Property SerNo) {
    $g_ex = $g | Select-Object -ExpandProperty Group
    $g_5 = $g_ex | Group-Object -Property '5'

    # если группа 'SerNo' > группы по '5', это деградация по 'remap'
    # if ($g.Count -gt 1) {"`v", $g, $g_5} # для наглядности можно включить вывод в консоль обеих групп
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

#region  # отчёт Degradation
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

#region  # отчёт Stable
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

#region  # КОНЕЦ
# замер времени выполнения скрипта
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion
