#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    сценарий в многопоточном режиме читает из WMI S.M.A.R.T. данные жёстких дисков указанного компьютера(-ов) и сохраняет отчёт в csv-формате

.DESCRIPTION
    для корректной работы сценарий необходимо запускать из-под УЗ администратора целевого компьютера

    сценарий:
        сканирует одиночный хост либо список машин
        получает через WMI доступные S.M.A.R.T.-данные жёстких дисков
        сохраняет отчёт в '.\output\yyyy-MM-dd_HH-mm $inp drives.csv'

.INPUTS
    имя компьютера в формате "mylaptop"
        или
    csv-файл списка компьютеров
    ожидаемый формат файла, первая строка содержит заголовки
        "HostName","LastScan"
        "MyHomePC",""
        "laptop",""
    коэффициент - кол-во потоков, которые будут запущены на каждом логическом ядре

    в обязательном поле "HostName" указываются имена компьютеров
    поле "LastScan" может быть пустым, в него по окончании работы скрипта будет записать on-line/off-line статус компьютера

.OUTPUTS
    csv-файл с "сырыми" S.M.A.R.T.-данными

.PARAMETER Inp
    имя компьютера или путь к csv-файлу списка компьютеров

.PARAMETER Out
    путь к файлу отчёта

.PARAMETER k
    кол-во потоков на одно логическое ядро, опытным путём на i5 оптимально k=35: минимальное время без ошибок нехватки памяти

.EXAMPLE
    .\Get-WMISMART.ps1 $env:COMPUTERNAME
        получить S.M.A.R.T. атрибуты дисков локального компьютера

.EXAMPLE
    .\Get-WMISMART.ps1 HOST_NAME
        получить S.M.A.R.T. атрибуты компьютера HOST_NAME

.EXAMPLE
    .\Get-WMISMART.ps1 .\input\example.csv
        получить S.M.A.R.T. атрибуты дисков компьютеров из списка hostname_list.csv

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
    [string] $Inp = ".\input\example.csv",  # имя хоста либо путь к файлу списка хостов
    [string] $Out = ".\output\$('{0:yyyy-MM-dd_HH-mm}' -f $(Get-Date)) $($Inp.ToString().Split('\')[-1].Replace('.csv', '')) drives.csv",
    [int]    $k   = 35
)

#region  # НАЧАЛО
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = @(Get-Date) # замер времени выполнения скрипта
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта
#endregion

Import-Module -Name ".\helper.psm1" -verbose  # вспомогательный модуль с функциями

if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;LastScan = "";})}  # проверям, что в параметрах - имя хоста или файл-список хостов

$clones = ($Computers | Group-Object -Property HostName | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group)  # проверка на дубликаты
if ($clones -ne $null) {Write-Host "'$Inp' содержит повторяющиеся имена компьютеров, это увеличит время получения результата", $clones.HostName -ForegroundColor Red -Separator "`n"}

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # отчёт по дискам

$ComputersOnLine = @()
$DiskInfo = @()

#region Multi-Threading: распараллелим проверку доступности компа по сети
#region: инициализация пула
$Pool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * $k + 0)
$Pool.ApartmentState = "MTA"
$Pool.Open()
$RunSpaces = @()
#endregion

#region: скрипт-блок задания, которое будет выполняться в потоке
$Payload = {Param ([string] $name = '127.0.0.1')

    for ($i = 0; $i -lt 2; $i++) {  # быстрее чем 'Test-Connection -Count 3'
        $WMIInfo = @()
        $ping = $false

        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)
        if ($ping) {

            try {$Win32_DiskDrive = Get-WmiObject -ComputerName $name -class Win32_DiskDrive -ErrorAction Stop}
            catch {$Win32_DiskDrive = @()}

            foreach ($Disk in $Win32_DiskDrive) {
                $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # в wql-запросе запрещены '\', поэтому заменим их на '_' (что означает "один любой символ"), см. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

                # смарт-атрибуты, флаги
                try {$WMIData = (Get-WmiObject -ComputerName $name -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific}
                catch {$WMIData = @()}
                if ($WMIData.Length -ne 512) {
                    Write-Host "`t", $Disk.Model, "- в WMI нет данных S.M.A.R.T." -ForegroundColor DarkYellow
                    continue
                }  # если данные не получены, не будем дёргать WMI ещё дважды вхолостую, переход к следующему диску хоста

                # пороговые значения
                try {$WMIThresholds = (Get-WmiObject -ComputerName $name -namespace root\wmi -Class MSStorageDriver_FailurePredictThresholds -Filter $wql -ErrorAction Stop).VendorSpecific}
                catch {$WMIThresholds = @()}

                # статус диска (Windows OS IMHO)
                try {$WMIStatus = (Get-WmiObject -ComputerName $name -namespace root\wmi –class MSStorageDriver_FailurePredictStatus -Filter $wql -ErrorAction Stop).PredictFailure}  # ИСТИНА (TRUE), если прогнозируется сбой диска. В этом случае нужно немедленно выполнить резервное копирование диска
                catch {$WMIStatus = $null}

                # добавляем новый диск в массив отчёта по дискам
                Import-Module -Name ".\helper.psm1" -verbose
                $WMIInfo += New-Object psobject -Property @{
                    ScanDate      =          Get-Date
                    HostName      = [string] $name
                    SerialNumber  = [string] $Disk.SerialNumber.Trim()  # Convert-hex2txt -wmisn ([string] $Disk.SerialNumber)  # не работает импорт модуля в скрипт-блоке
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

#region: запускаем задание и добавляем потоки в пул
foreach ($C in $Computers) {
    $NewShell = [PowerShell]::Create()

    $null = $NewShell.AddScript($Payload)
    $null = $NewShell.AddArgument($C.HostName)

    $NewShell.RunspacePool = $Pool

    $RunSpaces += [PSCustomObject]@{ Pipe = $NewShell; Status = $NewShell.BeginInvoke() }
}
#endregion

#region: после завершения каждого потока собираем его данные и закрываем, а после завершения всех потоков закрываем пул
if ($RunSpaces.Count -gt 0) {  # fixed bug: а были ли запущены потоки? - в случае пустого input-файла скрипт зависал на цикле while, т.к. ждал завершения хотя бы одного потока, хотя их вообще не было...
    while ($RunSpaces.Status.IsCompleted -notcontains $true) {}  # ждём завершения хотя бы одного потока начинаем принимать данные

    $t = ($RunSpaces.Status).Count  # общее кол-во потоков
    foreach ($RS in $RunSpaces ) {
        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "Проверка сетевой доступности компьютера и сбор S.M.A.R.T. данных" -Status "всего: $t" -CurrentOperation "осталось: $p"

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

#region: обновление БД
foreach($d in $DiskInfo){$d.SerialNumber = (Convert-hex2txt -wmisn $d.SerialNumber)}  # при необходимости приводим серийные номера в читаемый формат

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
    if ($HostID.ContainsKey($Scan.HostName))  # если в таблице с компьютерами уже есть запись
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

#region: экспорты:  отчёт по дискам, обновление входного файла (при необходимости)
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
if (Test-Path -Path $Inp) {$ComputersOnLine | Select-Object "HostName", "LastScan" | Export-Csv -Path $Inp -NoTypeInformation -Encoding UTF8}
#endregion

#region  # КОНЕЦ
# замер времени выполнения скрипта
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion
