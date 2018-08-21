#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    сценарий в многопоточном режиме собирает из WMI S.M.A.R.T. данные жёстких дисков указанного компьютера(-ов) и сохраняет отчёт в базу данных и в csv-файл в папке .\output


.DESCRIPTION
    для корректной работы сценарий необходимо запускать из-под УЗ администратора целевого компьютера

    сценарий:
        сканирует одиночный хост либо список машин
        получает через WMI доступные S.M.A.R.T.-данные жёстких дисков
        сохраняет отчёт в sqlite БД
        сохраняет отчёт в '.\output\yyyy-MM-dd $inp drives.csv'


.INPUTS
    имя компьютера в формате "mylaptop"
        или
    csv-файл списка компьютеров
    
    ожидаемый формат файла
        "HostName"
        "MyHomePC"
        "laptop"
    где первая строка - заголовок, а остальные - имена компьютеров


.OUTPUTS
    csv-файл с "сырыми" S.M.A.R.T.-данными


.PARAMETER Inp
    имя компьютера / путь к списку компьютеров в формате csv


.PARAMETER Out
    путь к файлу отчёта


.PARAMETER k
    кол-во потоков на логическое ядро
    опытным путём на i5 + 8 GB RAM оптимально k=35..37: минимальное время без ошибок нехватки памяти


.PARAMETER t
    пауза (в секундах) ожидания завершения потоков перед их повторной проверкой.
    Необходима для корректного завершения "нормальных" и отсева "зависших" потоков (напр. при снятии S.M.A.R.T. на этапе Get-WmiObject)
    
    таймаут (в минутах) для выхода из while-цикла сбора результатов потоков
    
    1...3  должно хватить для корректного завершения сканирования локального компьютера
    
    7...9+ для списка удалённых компьютеров


.PARAMETER aa
    переключатель 0/1, нужен для выбора режима того, как формировать сводные отчёты
    
    0 - пометить все записи в таблице `Scan` как активные
        В сводные отчёты попадут все диски за весь период сбора информации
        Т.е. даже не смотря на фактическую замену ЖД на компьютере, данные по старому диску все равно попадут в отчёты
        Этот режим может быть полезен, если необходимо увидеть всю картину целиком

    1 - автоархивация скан-записей перед пополнением БД свежими S.M.A.R.T.-данными
        Сводные отчёты будут строиться на основании записей только тех дисков, с чьих компьютеров удастся получить S.M.A.R.T. в этот раз
        Этот режим выбран по-умолчанию, т.к. отчёты покажут необходимый минимум информации, а если диск уже заменили, скан-записи старого диска будут помечены как архивные и не попадут в сводные отчёты


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
    # [alias('i')][string] $inp = ".\input\example.csv",
    
    [alias('i')][string] $inp = "$env:COMPUTERNAME",
    
    [alias('o')][string] $out = ".\output\$('{0:yyyy-MM-dd}' -f $(Get-Date)) $($inp.ToString().Split('\')[-1].Replace('.csv', '')) drives.csv",
    
                   [int] $k   = 1,

                   [int] $t   = 1,

    [alias('a')]   [int] $aa  = 1
)


#region  # НАЧАЛО

$psCmdlet.ParameterSetName | Out-Null

Clear-Host

$WatchDogTimer = [system.diagnostics.stopwatch]::startNew()

$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent

Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта

Import-Module -Name ".\helper.psm1" -verbose -Force  # вспомогательный модуль с функциями

#endregion


# проверям параметр inp - либо это файл-список хостов, либо считаем его именем хоста
if (Test-Path -Path $inp)

{
    $Computers = Import-Csv $inp
}

else

{
    $Computers = (New-Object psobject -Property @{HostName = $inp;ScanDate = "";})
}


$clones = ($Computers | Group-Object -Property HostName | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group)  # проверка на дубликаты

if ($clones -ne $null)

{
    Write-Host "'$inp' содержит дубликаты, это увеличит время получения результата", $clones.HostName -ForegroundColor Red -Separator "`n"
}


if (Test-Path $out)

{
    Remove-Item -Path $out -Force  # кому нужен старый отчёт по дискам ?!
}


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
            try
            
            {
                $Win32_DiskDrive = Get-WmiObject -ComputerName $name -class Win32_DiskDrive -ErrorAction Stop
            }
            
            catch {break}

            
            foreach ($Disk in $Win32_DiskDrive)
            
            {
                $wql = "InstanceName LIKE '%$($Disk.PNPDeviceID.Replace('\', '_'))%'"  # в wql-запросе запрещены '\', поэтому заменим их на '_' (что означает "один любой символ"), см. https://msdn.microsoft.com/en-us/library/aa392263(v=vs.85).aspx

                
                # смарт-атрибуты, флаги
                try
                
                {
                    $WMIData = (Get-WmiObject -ComputerName $name -namespace root\wmi -class MSStorageDriver_FailurePredictData -Filter $wql -ErrorAction Stop).VendorSpecific
                }
                
                catch {$WMIData = @()}

                
                if ($WMIData.Length -ne 512)  # если данные не получены, не будем дёргать WMI ещё дважды вхолостую, переход к следующему диску хоста
                
                {
                    Write-Host "`t", $Disk.Model, "- в WMI нет данных S.M.A.R.T." -ForegroundColor DarkYellow
                    continue
                }

                
                # пороговые значения
                try
                
                {
                    $WMIThresholds = (Get-WmiObject -ComputerName $name -namespace root\wmi -Class MSStorageDriver_FailurePredictThresholds -Filter $wql -ErrorAction Stop).VendorSpecific
                }
                
                catch {$WMIThresholds = @()}

                
                # статус диска (Windows OS IMHO) = true, если ОС прогнозирует сбой диска (и, как правило, при этом советует выполнить резервное копирование)
                try
                
                {
                    $WMIStatus = (Get-WmiObject -ComputerName $name -namespace root\wmi –class MSStorageDriver_FailurePredictStatus -Filter $wql -ErrorAction Stop).PredictFailure
                }
                
                catch {$WMIStatus = $null}

                
                # добавляем новый диск в массив отчёта по дискам
                # Import-Module -Name ".\helper.psm1" -verbose
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

        if ($c_false.Count -ne $p)
        
        {
            break
        }
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
            if ($p -gt 0)
            
            {
                Write-Host "пауза в $t сек закончилась. Всего зависло потоков: $p `n Multi-Threading режим сбора S.M.A.R.T. завершён...`n" -ForegroundColor Red
            }
            
            Write-Host "timer:" $WatchDogTimer.Elapsed.TotalSeconds, "`tdctCompleted:", $dctCompleted.Count,  "`tdctHang:",$dctHang.Count,  "`tосталось, `$p:",$p -ForegroundColor Magenta

            $RunSpaces | Where-Object -FilterScript {$_.Pipe.InstanceId.Guid -in ($dctHang.Keys)} | foreach {Write-Host $_.Pipe.Streams.Debug, $_.Pipe.Streams.Information, "`t", $_.Pipe.InstanceId.Guid -ForegroundColor Red}

            break  # завершение while-цикла: если кол-во незавершённых потоков после паузы (в секундах) не изменилось, значить они зависли)
        }

        else

        {
            Write-Host -ForegroundColor Green 'пауза в', $t, 'сек закончилась. Проверим незавершённые потоки ещё раз... :-)'
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
        'SerialNumber',`
        'ScanDate',`
        'Model',`
        'Size',`
        'InterfaceType',`
        'MediaType',`
        'DeviceID',`
        'PNPDeviceID',`
        'WMIStatus',`
        'WMIData',`
        'WMIThresholds'`
        | Sort-Object -Property 'HostName' | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8 #-Delimiter ';' -Append


    if (Test-Path -Path $inp)  # если на вход был подан файл со списком хостов, то экспортируем этот список со статусами он-лайн\офф-лайн
    
    {
        $ComputersOnLine += ($Computers | Where-Object -FilterScript {$_.HostName -notin $ComputersOnLine.HostName})  # | Select-Object -Property 'HostName')
        
        $ComputersOnLine | Select-Object 'HostName' | Sort-Object -Property 'HostName' | Export-Csv -Path $inp -NoTypeInformation -Encoding UTF8
    }
}

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): Export-Csv completed' -ForegroundColor Cyan

#endregion


#region: обновление БД

# автоархивация скан-записей перед пополнением БД свежими S.M.A.R.T.-данными
if ($aa)

{
    $null = Update-DB -tact ScanArch+
}

else

{
    $null = Update-DB -tact ScanArch-
}


$DiskID = Get-DBHashTable -query 'Disk'  # хэштаблица вида 'серийный номер диска' = ID диска

$HostID = Get-DBHashTable -query 'Host'  # хэштаблица вида 'имя хоста' = ID хоста

$ScanID = Get-DBHashTable -query 'Scan'  # хэштаблица вида 'дискИД хостИД ДатаСкана' = ID скан-записи

$ArchID = Get-DBHashTable -query 'Arch'  # хэштаблица вида 'дискИД' = значение Scan.Archived для "разархивирования" скан-записей, если эти записи были помечены как архивные, а диск только что отсканировался

foreach ($Scan in $DiskInfo)  # 'HostName' 'ScanDate' 'SerialNumber' 'Model' 'Size' 'InterfaceType' 'MediaType' 'DeviceID' 'PNPDeviceID' 'WMIData' 'WMIThresholds' 'WMIStatus'
{

        # Host
        if (!$HostID.ContainsKey($Scan.HostName))  # если в таблице с компьютерами уже есть запись

        {  # new record
            $hID = Update-DB -tact NewHost -obj ($Scan | Select-Object -Property 'HostName')

            if ($hID -gt 0) {$HostID[$Scan.HostName] = $hID}
        }

        else

        {
            $hID = $HostID[$Scan.HostName]
        }


        # Disk
        if (!$DiskID.ContainsKey($Scan.SerialNumber))

        {  # new record
            $dID = Update-DB -tact NewDisk -obj ($Scan | Select-Object -Property `
                    'SerialNumber', 'Model', 'Size', 'InterfaceType', 'MediaType', 'DeviceID', 'PNPDeviceID')

            if($dID -gt 0) {$DiskID[$Scan.SerialNumber] = $dID}
        }

        else

        {
            $dID = $DiskID[$Scan.SerialNumber]
        }

        # Scan
        $skey = "$dID $hID $($Scan.ScanDate.ToString())"

        if (!$ScanID.ContainsKey($skey))
        
        {  # new record
            $sID = Update-DB -tact NewScan -obj ($Scan | Select-Object -Property `
                    @{Name="DiskID"; Expression = {$DiskID[$Scan.SerialNumber]}},
                    @{Name="HostID"; Expression = {$HostID[$Scan.HostName]}},
                    'ScanDate',
                    'WMIData',
                    'WMIThresholds',
                    @{Name="WMIStatus"; Expression = {[int][System.Convert]::ToBoolean($Scan.WMIStatus)}})  # convert string 'false' to 0, 'true' to 1

            if($sID -gt 0) {$ScanID[$skey] = $sID}
        }

        # else { $sID = $ScanID[$skey] }  # если запись существует, то пропускаем её без обновления БД

        
        # диск "всплыл" при сканировании - ЕСЛИ он был архивным ($ArchID[$dID] -eq 1)
        
        # нужно "разархивировать" (Archived = 0) все его скан-записи - UPDATE `Scan` SET `Archived` = 0 WHERE `DiskID` = @DiskID;
        
        if ( <# $ArchID.ContainsKey($dID) -and #> $ArchID[$dID] -eq 1)  # 5-я версия powershell отрабатывает несуществующий ключ без исключений

        {
            $null = Update-DB -tact UpdScan -obj (New-Object psobject -Property @{DiskID = $dID})
        }
}

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): DataBase updated' -ForegroundColor Cyan

#endregion


#region  # КОНЕЦ

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): executed' -ForegroundColor  Green
$WatchDogTimer.Stop()  # $WatchDogTimer.Elapsed.TotalSeconds

#endregion
