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
     [string] $ReportDir = '.\output',  #
     [int] $k = 10
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

#region читаем WMI-отчёты из БД
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
$sql.CommandText = @"
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
"@
$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
$data = New-Object System.Data.DataSet
[void]$adapter.Fill($data)
$sql.Dispose()
$con.Close()

Measure-Command {###########################################

# 625,539048 seconds in Multi-Thread mode VS 41,4786834 in single mode...

#region: инициализация пула
$Pool = [RunspaceFactory]::CreateRunspacePool(1, [int] $env:NUMBER_OF_PROCESSORS * $k + 0)
$Pool.ApartmentState = "MTA"
$Pool.Open()
$RunSpaces = @()
#endregion

#region: скрипт-блок задания, которое будет выполняться в потоке
$Payload = {Param ($hd = $null) # HostName Model ScanDate SerialNumber Size WMIData WMIStatus WMIThresholds
function Convert-WMIArrays ([string] $data, [string] $thresh) {
<#
.SYNOPSIS
    функция переводит сырые массивы S.M.A.R.T.-данных диска в более понятный формат

.DESCRIPTION
    функция парсит массивы, конвертирует данные в зависимости от атрибута и возвращает весь набор атрибутов текущего диска

.INPUTS
    512-байт массив с данными из класса MSStorageDriver_FailurePredictData
    512-байт массив с данными из класса MSStorageDriver_FailurePredictThresholds

.OUTPUTS
    массив объектов-атрибутов S.M.A.R.T.

.PARAMETER data
    512-байт массив с данными из класса MSStorageDriver_FailurePredictData

.PARAMETER thresh
    512-байт массив с данными из класса MSStorageDriver_FailurePredictThresholds

.EXAMPLE
    $AtrInfo = Convert-WMIArrays -data $WMIData -thresh $WMIThresholds

.LINK
    структура 512-байт массива данных S.M.A.R.T. (корректное описание)
        http://www.t13.org/Documents/UploadedDocuments/docs2005/e05171r0-ACS-SMARTAttributes_Overview.pdf

.NOTES
    Author: Dmitry Mikhaylov

    чтение 512-байт массива 12ти байтовыми блоками начинается со смещения +2
    т.к. первые два байта являются ВЕРСИЕЙ ИДЕНТИФИКАТОРА структуры S.M.A.R.T.-данных (vendor-specific)

    структура 512-байт S.M.A.R.T.-массива данных
        Offset      Length (bytes)  Description
        0           2               SMART structure version (this is vendor-specific)
        2           12              Attribute entry 1
        2+(12)      12              Attribute entry 2
        ...         ...             ...
        2+(12*29)   12              Attribute entry 30

    структура 12-байтового блока
        0     Attribute ID (dec)
        1     Status flag (dec)
        2     Status flag (Bits 6–15 – Reserved)
        3     Value
        4     Worst
        5     Raw Value 1st byte
        6     Raw Value 2nd byte
        7     Raw Value 3rd byte
        8     Raw Value 4th byte
        9     Raw Value 5th byte
        10    Raw Value 6th byte
        11    Raw Value 7th byte / Reserved
#>

    $Result = @()
    [int[]] $data = $data.Split(' ')
    [int[]] $thresh = $thresh.Split(' ')
    if ($data.Length -eq 512 ) {  # работаем по массивам, каждая 12-байтовая группа отвечает за свой атрибут
        for ($i = 2; $i -le 350; $i = $i + 12) {  # всего может быть 29 атрибутов, длина значащей части массива 2+12*29 = 350 байт, оставшиеся байты будут мешать и могут вызвать исключение при добавлении в PropertyNote объекта диска
            if ($data[$i] -eq 0) {continue}  # пропускаем атрибуты с нулевым кодом. Thresholds при этом в расчёт не берём, т.к. у каких-то новых моделей ЖД в этом массиве почти одни нули

            $h0  = '{0:x2}' -f $data[$i+0]  # Attribute ID
            $h1  = '{0:x2}' -f $data[$i+1]  # Status flag
            $h2  = '{0:x2}' -f $data[$i+2]  # Status flag (Reserved)
            $h3  = '{0:x2}' -f $data[$i+3]  # Value
            $h4  = '{0:x2}' -f $data[$i+4]  # Worst
            $h5  = '{0:x2}' -f $data[$i+5]  # Raw Value 1st byte
            $h6  = '{0:x2}' -f $data[$i+6]  # Raw Value 2nd byte
            $h7  = '{0:x2}' -f $data[$i+7]  # Raw Value 3rd byte
            $h8  = '{0:x2}' -f $data[$i+8]  # Raw Value 4th byte
            $h9  = '{0:x2}' -f $data[$i+9]  # Raw Value 5th byte
            $h10 = '{0:x2}' -f $data[$i+10] # Raw Value 6th byte
            $h11 = '{0:x2}' -f $data[$i+11] # Raw Value 7th byte

            $raw = [Convert]::ToInt64("$h11 $h10 $h9 $h8 $h7 $h6 $h5".Replace(' ',''),16)  # переводим из 16-ричной в 10-тичную систему счисления

            if ($data[$i] -in @(190, 194)) {$raw = $data[$i+5]}  # температуру оставляем как есть
            if ($data[$i] -in @(9)) {$raw = [Convert]::ToInt64("$h6 $h5".Replace(' ',''),16)}  # наработка в часах

            $Result += (New-Object PSObject -Property @{
                saIDDec     = [int]    $data[$i]
                saIDHex     = '{0:x2}' -f $data[$i]
                saName      = [string] $(if ($dctSMART.ContainsKey([int] $data[$i])) {$dctSMART[[int] $data[$i]]} else {$null})  # вытаскиваем из словаря $dctSMART
                saThreshold = [int]    $thresh[$i+1]
                saValue     = [int]    $data[$i+3]
                saWorst     = [int]    $data[$i+4]
                saRaw       = [long]   $raw
                saRawHex    = [string] "$h11 $h10 $h9 $h8 $h7 $h6 $h5"
                saRawDec    = [string] "$($data[$i+11]) $($data[$i+10]) $($data[$i+9]) $($data[$i+8]) $($data[$i+7]) $($data[$i+6]) $($data[$i+5])"
                saFlagDec   = [int]    $data[$i+1]
                saFlagBin   = [string] [convert]::ToString($data[$i+1],2)
                saFullHex   = [string] "$h0 $h1 $h2 $h3 $h4 $h5 $h6 $h7 $h8 $h9 $h10 $h11"
                saFullDec   = [string] "$($data[$i+0]) $($data[$i+1]) $($data[$i+2]) $($data[$i+3]) $($data[$i+4]) $($data[$i+5]) $($data[$i+6]) $($data[$i+7]) $($data[$i+8]) $($data[$i+9]) $($data[$i+10]) $($data[$i+11])"
            })
        }
    }
    return $Result
}
    $Disk = (New-Object PSObject -Property @{
        ScanDate = $hd.ScanDate
        PC       = $hd.HostName
        Model    = $hd.Model
        SerNo    = $hd.SerialNumber  # Convert-hex2txt -wmisn ([string] $hd.SerialNumber.Trim())
        })
    foreach ($atr in Convert-WMIArrays -data $hd.WMIData -thresh $hd.WMIThresholds) {  # расшифровываем атрибуты и по-одному добавляем к объекту $Disk
        try {$Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop}  # на случай дубликатов в отчёте
        catch {Write-Host "DUPLICATE: '$($hd.HostName)'" -ForegroundColor Red}
    }

    Return $Disk
}
#endregion

#region: запускаем задание и добавляем потоки в пул
foreach ($hd in $data.Tables.Rows) {  # обрабатываем результаты сканов $data.Tables.Rows.Count
    $NewShell = [PowerShell]::Create()

    $null = $NewShell.AddScript($Payload)
    $null = $NewShell.AddArgument($hd)

    $NewShell.RunspacePool = $Pool

    $RunSpaces += [PSCustomObject]@{ Pipe = $NewShell; Status = $NewShell.BeginInvoke() }
}
#endregion

#region: после завершения каждого потока собираем его данные и закрываем, а после завершения всех потоков закрываем пул
if ($RunSpaces.Count -gt 0) {  # fixed bug: а были ли запущены потоки? - чтобы не зависал на цикле while, т.к. ждал завершения хотя бы одного потока, хотя их вообще не было...
    while ($RunSpaces.Status.IsCompleted -notcontains $true) {}  # ждём завершения хотя бы одного потока начинаем принимать данные

    $t = ($RunSpaces.Status).Count  # общее кол-во потоков
    foreach ($RS in $RunSpaces ) {
        $p = ($RunSpaces.Status | Where-Object -FilterScript {$_.IsCompleted -eq $false}).Count  # кол-во незавершённых потоков
        Write-Progress -id 1 -PercentComplete (100 * $p / $t) -Activity "ПРЕОБРАЗОВАНИЕ СКАНОВ В СВОДНЫЕ ТАБЛИЦЫ" -Status "всего: $t" -CurrentOperation "осталось: $p"

        $AllInfo += $RS.Pipe.EndInvoke($RS.Status)  # сбор информации с заданий
        
        $RS.Pipe.Dispose()
    }
    $Pool.Close()
    $Pool.Dispose()
}
#endregion

} | select -ExpandProperty TotalSeconds ###########################################
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


<#


Replace(' ATA Device', '')
Replace(' SCSI Disk Device', '')
#>