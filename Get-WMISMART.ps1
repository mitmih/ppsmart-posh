#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    сценарий читает из WMI S.M.A.R.T. данные жёстких дисков указанного компьютера(-ов) и сохраняет отчёт в csv-формате

.DESCRIPTION
    для корректной работы сценарий необходимо запускать из-под УЗ администратора целевого компьютера

    сценарий:
        сканирует одиночный хост либо список машин
        получает через WMI доступные S.M.A.R.T.-данные жёстких дисков
        сохраняет отчёт в .\output\yyyy-MM-dd_HH-mm-ss.csv

.INPUTS
    имя компьютера в формате "mylaptop"
        или
    csv-файл списка компьютеров
    ожидаемый формат файла, первая строка содержит заголовки
        "HostName","Status"
        "MyHomePC",""
        "laptop",""

    в обязательном поле "HostName" указываются имена компьютеров
    поле "Status" может быть пустым, в него по окончании работы скрипта будет записать on-line/off-line статус компьютера

.OUTPUTS
    csv-файл с "сырыми" S.M.A.R.T.-данными

.PARAMETER Inp
    имя компьютера или путь к csv-файлу списка компьютеров

.PARAMETER Out
    путь к файлу отчёта

.EXAMPLE
    .\Get-WMISMART.ps1 $env:COMPUTERNAME
        получить S.M.A.R.T. атрибуты дисков локального компьютера

.EXAMPLE
    .\Get-WMISMART.ps1 HOST_NAME
        получить S.M.A.R.T. атрибуты компьютера HOST_NAME

.EXAMPLE
    .\Get-WMISMART.ps1 .\input\debug.csv
        получить S.M.A.R.T. атрибуты дисков компьютеров из списка debug.csv

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

    не очень корректное описание структуры 512-байт массива S.M.A.R.T.: тут написано, что 1й блок начинается сразу с 0го байта, а на самом деле первые два байта означают версию структуры S.M.A.R.T.
        https://social.msdn.microsoft.com/Forums/en-US/af01ce5d-b2a6-4442-b229-6bb32033c755/using-wmi-to-get-smart-status-of-a-hard-disk?forum=vbgeneral

.NOTES
    Author: Dmitry Mikhaylov
#>

[CmdletBinding(DefaultParameterSetName="all")]

param
(
#     [string] $Inp = "$env:COMPUTERNAME",  # имя хоста либо путь к файлу списка хостов
    [string] $Inp = ".\input\debug.csv",  # имя хоста либо путь к файлу списка хостов
    [string] $Out = ".\output\$($Inp.ToString().Split('\')[-1].Replace('.csv', '')) $('{0:yyyy-MM-dd_HH}' -f $(Get-Date))_drives.csv"
#     [string] $Out = ".\output\$($Inp.ToString().Split('\')[-1].Replace('.csv', '')) $('{0:yyyy-MM-dd_HH-mm-ss}' -f $(Get-Date))_drives.csv"
)
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = Get-Date # замер времени выполнения скрипта
set-location "$($MyInvocation.MyCommand.Definition | split-path -parent)"  # локальная корневая папка "./" = текущая директория скрипта

Import-Module -Name ".\helper.psm1" -verbose  # вспомогательный модуль с функциями

if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;Status = "";})}  # проверям, что в параметрах - имя хоста или файл-список хостов

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # отчёт по дискам

# Multy-Threading: распараллелим проверку доступности компа по сети
#region создаём пул runspace`ов
$PingPool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * 1 + 1)
$PingPool.ApartmentState = "MTA"
$PingPool.Open()
$PingRunSpaces = @()
$ComputersOnLine = @()
#endregion

#region что будет распараллелено
$PingPayload = {Param ([string] $name = '127.0.0.1')
    $ping = $false
    for ($i = 0; $i -lt 3; $i++) {
        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)
        if ($ping) {break}
    }

    Return (New-Object psobject -Property @{HostName = $name;Status = $ping;})
}
#endregion

#region создаём runspace`ы, запускаем и добавляем их в пул
foreach ($C in $Computers) {
    $PingNewShell = [PowerShell]::Create()

    $null = $PingNewShell.AddScript($PingPayload)
    $null = $PingNewShell.AddArgument($C.HostName)

    $PingNewShell.RunspacePool = $PingPool

    $PingRunSpaces += [PSCustomObject]@{ Pipe = $PingNewShell; Status = $PingNewShell.BeginInvoke() }
}
#endregion

#region после завершения всех запущенных потоков собираем данные, закрываем потоки и пул
if ($PingRunSpaces.Count -gt 0) {  # а были ли запущены потоки? - был баг: файл со списком был пуст, ни одного компа, и скрипт зависал на цилке while, потому-что ждал завершения хотя бы одного потока, а потоков то и не было...
    while ($PingRunSpaces.Status.IsCompleted -notcontains $true) {}  # после завершения хотя бы одного потока сохраняем его данные

    foreach ($RS in $PingRunSpaces ) {
        $t = ($PingRunSpaces.Status).Count  # общее кол-во потоков
        $p = ($PingRunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "Проверка сетевой доступности компьютера" -Status "всего: $t" -CurrentOperation "осталось: $p"

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
$p = 0  # прогресс-бар: $p - счётчик (текущий комп), $t - общее количество (всего компов к обработке)
foreach ($C in $ComputersOnLine | Where-Object {$_.Status -eq 'online'}) {
    $p += 1
    Write-Progress -id 2 -PercentComplete (100 * $p / $t) -Activity $Inp -Status "обрабатываю $p-й компьютер из $t" -CurrentOperation $C.HostName

    Write-Host "`n$($C.HostName) $($C.Status)" -ForegroundColor Green

    try {$Win32_DiskDrive = Get-WmiObject -ComputerName $C.HostName -class Win32_DiskDrive -ErrorAction Stop}
    catch {$Win32_DiskDrive = @()}

    foreach ($Disk in $Win32_DiskDrive) {
        $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # в wql-запросе запрещены '\', поэтому заменим их на '_' (что означает "один любой символ"), см. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

        # смарт-атрибуты, флаги
        try {$WMIData = (Get-WmiObject -ComputerName $C.HostName -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific}
        catch {$WMIData = @()}
        if ($WMIData.Length -ne 512) {
            Write-Host "`t", $Disk.Model, "- в WMI нет данных S.M.A.R.T." -ForegroundColor DarkYellow
            continue
        }  # если данные не получены, не будем дёргать WMI ещё дважды вхолостую, переход к следующему диску хоста

        # пороговые значения
        try {$WMIThresholds = (Get-WmiObject -ComputerName $C.HostName -namespace root\wmi -Class MSStorageDriver_FailurePredictThresholds -Filter $wql -ErrorAction Stop).VendorSpecific}
        catch {$WMIThresholds = @()}

        # статус диска (Windows OS IMHO)
        try {$WMIStatus = (Get-WmiObject -ComputerName $C.HostName -namespace root\wmi –class MSStorageDriver_FailurePredictStatus -Filter $wql -ErrorAction Stop).PredictFailure}  # ИСТИНА (TRUE), если прогнозируется сбой диска. В этом случае нужно немедленно выполнить резервное копирование диска
        catch {$WMIStatus = $null}

        Write-Host "`t", $Disk.Model, $WMIStatus, $WMIData.Length, $WMIThresholds.Length -ForegroundColor Cyan

        # добавляем новый диск в массив отчёта по дискам
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

# экспорт отчёта по дискам
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

# если на вход был подан файл со списком хостов, то экспортируем этот список со статусами он-лайн\офф-лайн
if (Test-Path -Path $Inp) {$ComputersOnLine | Select-Object "HostName", "Status" | Export-Csv -Path $Inp -NoTypeInformation -Encoding UTF8}

# замер времени выполнения скрипта
$ExecTime = [System.Math]::Round($( $(Get-Date) - $TimeStart ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
