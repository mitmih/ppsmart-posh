#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    сценарий в многопоточном режиме читает из WMI S.M.A.R.T. данные жёстких дисков указанного компьютера(-ов) и сохраняет отчёт в csv-формате

.DESCRIPTION
    для корректной работы сценарий необходимо запускать из-под УЗ администратора целевого компьютера

    сценарий:
        сканирует одиночный хост либо список машин
        получает через WMI доступные S.M.A.R.T.-данные жёстких дисков
        сохраняет отчёт в '.\output\yyyy-MM-dd $inp drives.csv'

.INPUTS
    имя компьютера в формате "mylaptop"
        или
    csv-файл списка компьютеров
    ожидаемый формат файла, первая строка содержит заголовки
        "HostName"
        "MyHomePC"
        "laptop"
    коэффициент - кол-во потоков, которые будут запущены на каждом логическом ядре

    в обязательном поле "HostName" указываются имена компьютеров
    поле "ScanDate" может быть пустым, в него по окончании работы скрипта будет записать on-line/off-line статус компьютера

.OUTPUTS
    csv-файл с "сырыми" S.M.A.R.T.-данными

.PARAMETER Inp
    имя компьютера или путь к csv-файлу списка компьютеров

.PARAMETER Out
    путь к файлу отчёта

.PARAMETER k
    кол-во потоков на логическое ядро, опытным путём на i5 оптимально k=35: минимальное время без ошибок нехватки памяти

.PARAMETER t
    пауза (в секундах) ожидания завершения потоков перед их повторной проверкой. Необходимо для отсева потоков, "зависших" на Get-WmiObject
    таймаут (в минутах) для выхода из while-цикла сбора результатов потоков
    1...3  для локальной машины должно хватить для корректного завершения сканирования
    7...9+ для списка удалённых компьютеров

.EXAMPLE
    .\Get-WMISMART.ps1 $env:COMPUTERNAME
        получить S.M.A.R.T. атрибуты дисков локального компьютера

.EXAMPLE
    .\Get-WMISMART.ps1 HOST_NAME
        получить S.M.A.R.T. атрибуты компьютера HOST_NAME

.EXAMPLE
    .\Get-WMISMART.ps1 .\input\example.csv -k 37 -t 13
        получить S.M.A.R.T. атрибуты дисков компьютеров списка .\input\example.csv, запустить maximum 37 потоков на ядро, пауза в 13 секунд и таймаут 13 минут

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
    [string] $Inp = "$env:COMPUTERNAME",  # имя хоста либо путь к файлу списка хостов
    # [string] $Inp = ".\input\example.csv",  # имя хоста либо путь к файлу списка хостов
    [string] $Out = ".\output\$('{0:yyyy-MM-dd}' -f $(Get-Date)) $($Inp.ToString().Split('\')[-1].Replace('.csv', '')) drives.csv",
    [int]    $k   = 37,
    [int]    $t   = 1  # once again timer
)


#region  # НАЧАЛО

$psCmdlet.ParameterSetName | Out-Null

Clear-Host

$WatchDogTimer = [system.diagnostics.stopwatch]::startNew()

$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта

Import-Module -Name ".\helper.psm1" -verbose -Force  # вспомогательный модуль с функциями

#endregion


if (Test-Path -Path $Inp) {$Computers = Import-Csv $Inp} else {$Computers = (New-Object psobject -Property @{HostName = $Inp;ScanDate = "";})}  # проверям, что в параметрах - имя хоста или файл-список хостов

$clones = ($Computers | Group-Object -Property HostName | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group)  # проверка на дубликаты
if ($clones -ne $null) {Write-Host "'$Inp' содержит дубликаты, это увеличит время получения результата", $clones.HostName -ForegroundColor Red -Separator "`n"}

if (Test-Path $Out) {Remove-Item -Path $Out -Force}  # отчёт по дискам

$ComputersOnLine = @()
$DiskInfo = @()


#region Multi-Threading: распараллелим проверку доступности компа по сети

#region: инициализация пула

$x = $Computers.Count / [int] $env:NUMBER_OF_PROCESSORS
$max = if (($x + 1) -lt $k) {[int] $env:NUMBER_OF_PROCESSORS * ($x + 1) + 1} else {[int] $env:NUMBER_OF_PROCESSORS * $k}
$Pool = [RunspaceFactory]::CreateRunspacePool(1, $max)
$Pool.ApartmentState = "MTA"
$Pool.Open()
$RunSpaces = @()

#endregion


#region: скрипт-блок задания, которое будет выполняться в потоке

$Payload = {Param ([string] $name = $env:COMPUTERNAME)

    Write-Debug $name -Debug

    for ($i = 0; $i -lt 2; $i++)
    {

        $WMIInfo = @()

        $ping = (Test-Connection -Count 1 -ComputerName $name -Quiet)

        if ($ping)  # при удачном ping получаем данные
        {
            try {$Win32_DiskDrive = Get-WmiObject -ComputerName $name -class Win32_DiskDrive -ErrorAction Stop}
            catch {break}

            foreach ($Disk in $Win32_DiskDrive)
            {
                $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # в wql-запросе запрещены '\', поэтому заменим их на '_' (что означает "один любой символ"), см. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

                # смарт-атрибуты, флаги
                try {$WMIData = (Get-WmiObject -ComputerName $name -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific}
                catch {$WMIData = @()}

                if ($WMIData.Length -ne 512)
                {
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
                    ScanDate      =          $('{0:yyyy.MM.dd}' -f $(Get-Date))
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

            break  # и прерываем цикл проверки
        }
    }

    Return @(
        (New-Object psobject -Property @{
            HostName = $name
            ScanDate = $('{0:yyyy.MM.dd}' -f $(Get-Date))
            Ping = if ($ping) {1} else {0}
            WMIInfo = $WMIInfo.Count
        }),
        $WMIInfo
    )
}

#endregion


#region: запускаем задание и добавляем потоки в пул

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


#region: после завершения каждого потока собираем его данные и закрываем, а после завершения всех потоков закрываем пул

$dctCompleted = @{}  # словарь завершенных заданий
$dctHang = @{}  # незавершённые задания
$total = $RunSpaces.Count  # общее кол-во потоков

# в случае локального компьютера поток может завершиться до начала цикла, и чтобы получить данные нужно выполнить его хотя бы один раз
While ($RunSpaces.Status.IsCompleted -contains $false -or ($total -eq ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $true}).Count) )
{
    $wpCompl   = "Проверка компьютера ping'ом и сбор S.M.A.R.T. данных,   ЗАВЕРШЁННЫЕ потоки"
    $wpNoCompl = "Проверка компьютера ping'ом и сбор S.M.A.R.T. данных, НЕЗАВЕРШЁННЫЕ потоки"

    $c_true = $RunSpaces | Where-Object -FilterScript {$_.Status.IsCompleted -eq $true}  # сначала отфильтруем только завершённые задания
    $c_true_filtred = $c_true | Where-Object -FilterScript {!$dctCompleted.ContainsKey($_.Pipe.InstanceId.Guid)}  # затем только те, которые ещё не обрабатывали, т.е. которые пока не попали в словарь $dctCompleted

    foreach ($RS in $c_true_filtred)  # цикл по завершённым_факт, который ещё нет в словаре (завершённых_учёт)
    {
        $Result = @()

        if($RS.Status.IsCompleted -and !$dctCompleted.ContainsKey($RS.Pipe.InstanceId.Guid))
        {
            $Result = $RS.Pipe.EndInvoke($RS.Status)
            $RS.Pipe.Dispose()

            $ComputersOnLine += $Result[0]
            if ($Result[1].Count -gt 0) {$DiskInfo += $Result[1]}

            $dctCompleted[$RS.Pipe.InstanceId.Guid] = $WatchDogTimer.Elapsed.TotalSeconds

            $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков
            Write-Progress -id 1 -PercentComplete (100 * ($dctCompleted.Count) / $total) -Activity $wpCompl -Status "всего: $total" -CurrentOperation "готово: $($dctCompleted.Count)"
            Write-Progress -id 2 -PercentComplete (100 * $p / $total) -Activity $wpNoCompl -Status "всего: $total" -CurrentOperation "осталось: $p"

            if($dctHang.ContainsKey($RS.Pipe.InstanceId.Guid))
            {
                $dctHang.Remove($RS.Pipe.InstanceId.Guid)
            }
        }
    }


    # Start-Sleep -Milliseconds 100
    $c_false = $RunSpaces | Where-Object -FilterScript {$_.Status.IsCompleted -eq $false}

    foreach($RS in $c_false)  # цикл по незавершённым
    {
        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков

        Write-Progress -id 1 -PercentComplete (100 * ($dctCompleted.Count) / $total) -Activity $wpCompl -Status "всего: $total" -CurrentOperation "готово: $($dctCompleted.Count)"
        Write-Progress -id 2 -PercentComplete (100 * $p / $total) -Activity $wpNoCompl -Status "всего: $total" -CurrentOperation "осталось: $p"

        $dctHang[$RS.Pipe.InstanceId.Guid] = $WatchDogTimer.Elapsed.TotalSeconds

        if ($c_false.Count -ne $p) {break}
    }

    # Start-Sleep -Milliseconds 100
    $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков

    Write-Progress -id 2 -PercentComplete (100 * $p / $total) -Activity $wpNoCompl -Status "всего: $total" -CurrentOperation "осталось: $p"
    Write-Host "timer:" $WatchDogTimer.Elapsed.TotalSeconds, "`tdctCompleted:", $dctCompleted.Count,  "`tdctHang:",$dctHang.Count,  "`tосталось, `$p:",$p -ForegroundColor Yellow
    Write-Progress -id 1 -PercentComplete (100 * ($dctCompleted.Count) / $total) -Activity $wpCompl -Status "всего: $total" -CurrentOperation "готово: $($dctCompleted.Count)"

    #         кол-во зависших не изменилось                                          И  (завершён хотя бы один И всего хотя бы 2)
    $escape = ($p -eq $dctHang.Count) -and ($total -eq ($dctCompleted.Count + $p)) -and ( $(if ($total -gt $p) {$dctCompleted.Count -gt 0} else {$true}) )
    #                                   И  всего = завершёнка_учёт + не_завершёнка_факт         ^это^ условие означает что мог быть запущен один поток и он же мог зависнуть

    if ($escape)
    {
        Start-Sleep -Seconds $t  # пауза чтобы завершить опоздавшие потоки и не ждать лишний раз "зависшие"

        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков

        $escape = ($p -eq $dctHang.Count) -and ($total -eq ($dctCompleted.Count + $p)) -and ( $(if ($total -gt $p) {$dctCompleted.Count -gt 0} else {$true}) )

        if ($escape)
        {
            Write-Host 'пауза в', $t, 'сек закончилась. Все незавершённые потоки считаются "зависшими" .Выход из Multi-Threading-цикла...' -ForegroundColor Red
            Write-Host "timer:" $WatchDogTimer.Elapsed.TotalSeconds, "`tdctCompleted:", $dctCompleted.Count,  "`tdctHang:",$dctHang.Count,  "`tосталось, `$p:",$p -ForegroundColor Magenta

            $RunSpaces | Where-Object -FilterScript {$_.Pipe.InstanceId.Guid -in ($dctHang.Keys)} | foreach {Write-Host $_.Pipe.Streams.Debug, $_.Pipe.Streams.Information, "`t", $_.Pipe.InstanceId.Guid -ForegroundColor Red}

            break  # завершение while-цикла: если кол-во незавершённых потоков после паузы (в секундах) не изменилось, значить они зависли)
        }

        else

        {
            Write-Host -ForegroundColor Green 'пауза в', $t, 'сек закончилась. Эх раз, ещё раз... :-)'
        }
    }

    else

    {
        continue
    }


    if ($WatchDogTimer.Elapsed.TotalMinutes -gt $t)
    {
        break  # внештатное завершение while-цикла по таймауту, в минутах
    }
}

<# нужно завершить "зависшие" потоки (возможно через Job с таймаутом), прежде чем закрывать пул
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


foreach($d in $DiskInfo){$d.SerialNumber = (Convert-hex2txt -wmisn $d.SerialNumber)}  # при необходимости приводим серийные номера в читаемый формат
Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): SerialNumber`s converted' -ForegroundColor Cyan


#region: экспорты:  отчёт по дискам, обновление входного файла (при необходимости)
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


    if (Test-Path -Path $Inp)  # если на вход был подан файл со списком хостов, то экспортируем этот список со статусами он-лайн\офф-лайн
    {
        $ComputersOnLine += ($Computers | Where-Object -FilterScript {$_.HostName -notin $ComputersOnLine.HostName})  # | Select-Object -Property 'HostName')
        $ComputersOnLine | Select-Object 'HostName' | Sort-Object -Property 'HostName' | Export-Csv -Path $Inp -NoTypeInformation -Encoding UTF8
    }
}

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): Export-Csv completed' -ForegroundColor Cyan

#endregion


#region: обновление БД

$DiskID = Get-DBHashTable -table 'Disk'
$HostID = Get-DBHashTable -table 'Host'

foreach ($Scan in $DiskInfo)  # 'HostName' 'ScanDate' 'SerialNumber' 'Model' 'Size' 'InterfaceType' 'MediaType' 'DeviceID' 'PNPDeviceID' 'WMIData' 'WMIThresholds' 'WMIStatus'
{

    # Host
    if (!$HostID.ContainsKey($Scan.HostName))  # new record

    {
        $hID = Update-DB -tact NewHost -obj ($Scan | Select-Object -Property 'HostName')

        if ($hID -gt 0) {$HostID[$Scan.HostName] = $hID}
    }


    # Disk
    if (!$DiskID.ContainsKey($Scan.SerialNumber))  # new record

    {
        $dID = Update-DB -tact NewDisk -obj ($Scan | Select-Object -Property `
                'SerialNumber', 'Model', 'Size', 'InterfaceType', 'MediaType', 'DeviceID', 'PNPDeviceID')

        if($dID -gt 0) {$DiskID[$Scan.SerialNumber] = $dID}
    }


        # Scan  # таблица имеет уникальный индекс 	`DiskID`, `HostID`, `ScanDate`
        $sID = Update-DB -tact NewScan -obj ($Scan | Select-Object -Property `
                @{Name="DiskID"; Expression = {$DiskID[$Scan.SerialNumber]}},
                @{Name="HostID"; Expression = {$HostID[$Scan.HostName]}},
                #@{Name="Archived"; Expression = {0}},
                'ScanDate',
                'WMIData',
                'WMIThresholds',
                'WMIStatus')

}

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): DataBase updated' -ForegroundColor Cyan

#endregion


#region  # КОНЕЦ

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): executed' -ForegroundColor  Green
$WatchDogTimer.Stop()  # $WatchDogTimer.Elapsed.TotalSeconds

#endregion
