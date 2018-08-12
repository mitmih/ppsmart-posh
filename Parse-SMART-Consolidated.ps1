﻿#requires -version 3  # требуемая версия PowerShell

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
     [string] $ReportDir = 'output',  # директория для вывода отчётов
     [int]    $5edge     = 0,        # грань по 5-му атрибуту, после которой диск попадёт в отчёт
     [int]    $csv       = 0,         # формировать csv
     [int]    $html      = 1          # формировать html
)

#region  # НАЧАЛО

$psCmdlet.ParameterSetName | Out-Null

Clear-Host

$WatchDogTimer = [system.diagnostics.stopwatch]::startNew()

$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта

Import-Module -Name ".\helper.psm1" -verbose -Force  # вспомогательный модуль с функциями

#endregion


$WMIFiles = Get-ChildItem -Path $ReportDir -Filter '*drives.csv' -Recurse  # массив отчётов по дискам

$AllInfo = @()  # полная инфа по всем дискам из всех отчётов

$Degradation = @()  # деградация по 5-му атрибуту (remap)

$Stable = @()  # стабильные, без деградации по remap`у


#region  # читаем WMI-отчёты из БД

#region  # refactor to helper module functions!

# проверяем битность среды выполнения для подключения подходящей библиотеки
if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $RootDir -ChildPath 'x64\System.Data.SQLite.dll'}  # 64-bit
elseif ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $RootDir -ChildPath 'x32\System.Data.SQLite.dll'}  # 32-bit
else {Write-Host 'Hmmm... not 32 or 64 bit...'}

# подключаем библиотеку для работы с sqlite
try {Add-Type -Path $sqlite -ErrorAction Stop}
catch {Write-Host "Importing the SQLite assemblies, '$sqlite', failed..."}

$db = Join-Path -Path $RootDir -ChildPath 'ppsmart-posh.db'  # путь к БД

if (Test-Path $db)

{
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
    WHERE Scan.Archived = 0
    ORDER BY Scan.ID;
'@

    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql

    $data = New-Object System.Data.DataSet

    [void]$adapter.Fill($data)

    $sql.Dispose()
    $con.Close()
}

#endregion


#region "разворот" данных

foreach ($hd in $data.Tables.Rows)  # обрабатываем результаты сканов $data.Tables.Rows.Count
# $data.Tables.Rows | Select-Object -Property SerialNumber, HostName, ScanDate, @{Name="Model"; Expression = {$_.Model.Replace(' ATA Device', '').Replace(' SCSI Disk Device', '')}}
{
    $Disk = (New-Object PSObject -Property @{
        ScanDate     = $hd.ScanDate
        HostName     = $hd.HostName
        Model        = $hd.Model.Replace(' ATA Device', '').Replace(' SCSI Disk Device', '')
        SerialNumber = $hd.SerialNumber
        Size         = $hd.Size  # Convert-hex2txt -wmisn ([string] $hd.SerialNumber.Trim())
        })


    foreach ($atr in ( Convert-WMIArrays -data $hd.WMIData -thresh $hd.WMIThresholds | where {$_.saIDDec -in (9,5,184,187,197,198,200)}))  # расшифровываем атрибуты и по-одному добавляем к объекту $Disk

    {
        try    # на случай дубликатов в отчёте и ошибки добавления атрибутов

        {
            $Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop
        }

        catch

        {
            Write-Host "DUPLICATE: '$($hd.HostName)'" -ForegroundColor Red
        }
    }


    $AllInfo += $Disk
} #endregion

#endregion


#region  # поиск деградаций по 5-му атрибуту

$dcteo = @{}  # словарь для хранения серийника и чёт/нечёт стиля

$eoStbl = 0  # ($eoStbl % 2 -eq 0) для применения разных css-стилей к чёт/нечёт строкам таблицы в html-отчёте

$eoDegr = 0  #

foreach ($g in $AllInfo | Sort-Object -Property SerialNumber,ScanDate | Group-Object -Property SerialNumber)

{
    $5val = ($g | Select-Object -ExpandProperty Group | Sort-Object -Property '5' | Select-Object -Last 1)
    if ($5val.'5' -le $5edge)
    # вырезаем из отчёта диски, у которых на последнем скане remap <= $5edge
    {
        Write-Host "excluded:`t" $5val.HostName "`t (fact) $($5val.'5') <= $5edge (edge)" -ForegroundColor Yellow
        continue
    }

    $g_ex = $g | Select-Object -ExpandProperty Group

    $g_5 = $g_ex | Group-Object -Property '5'


    # if ($g.Count -gt 1) {"`v", $g, $g_5} # для наглядности можно включить вывод в консоль обеих групп


    if ($g.Count -eq $g_5.Count)  # если группа 'SerialNumber' = группе по '5', это НЕ деградация по 'remap', добавим диск в стабильные и продолжим поиск

    {
        $Stable += $g_ex | Select-Object -Last 1

        # stable в консоль
        # Write-Host "stable:`t`t" ($g_ex | Select-Object -Last 1).HostName -ForegroundColor Green  # Stable

        $dcteo.Add(($g_ex.SerialNumber | Group-Object).Name, [int]($eoStbl % 2 -eq 0))
        $eoStbl += 1

        Continue
    }


    # Degradation в консоль
    # Write-Host "degradation:" ($g_ex | Select-Object -Property 'HostName' -Unique).HostName -ForegroundColor Magenta  # Degradation

    foreach ($r in $g_5)

    {
        # в отчёт попадёт только изменившееся значение атрибута remap
        $Degradation += ($r | Select-Object -ExpandProperty Group | Select-Object -First 1 )
    }

    $dcteo.Add(($g_ex.SerialNumber | Group-Object).Name, [int]($eoDegr % 2 -eq 0))
    $eoDegr += 1
}

#endregion


#region  # Reports

$ReportSelectProperties = @(
    'HostName',
    'Model',
    'SerialNumber',
    @{Expression = { $_.Size.ToString() + ' Gb' };   Name='Size'},
    'ScanDate',
    @{Expression = { $_.'9'.ToString() + ' hours' };   Name='9 Power-On Hours'},
    @{Expression = { $_.'5' };   Name='5 Reallocated Sectors Count'},
    @{Expression = { $_.'184' }; Name='184 End-to-End error / IOEDC'},
    @{Expression = { $_.'187' }; Name='187 Reported Uncorrectable Errors'},
    @{Expression = { $_.'197' }; Name='197 Current Pending Sector Count'},
    @{Expression = { $_.'198' }; Name='198 (Offline) Uncorrectable Sector Count'},
    @{Expression = { $_.'200' }; Name='200 Multi-Zone Error Rate / Write Error Rate (Fujitsu)'}
)

$ReportSortProperties = @(
    @{Expression = '5 Reallocated Sectors Count'; Descending = $True},
    @{Expression = '184 End-to-End error / IOEDC'; Descending = $True},
    @{Expression = '187 Reported Uncorrectable Errors'; Descending = $True},
    @{Expression = '197 Current Pending Sector Count'; Descending = $True},
    @{Expression = '198 (Offline) Uncorrectable Sector Count'; Descending = $True},
    @{Expression = '200 Multi-Zone Error Rate / Write Error Rate (Fujitsu)'; Descending = $True}
)


#region  # CSV отчёт Degradation

if ($csv)

{
    $csvDegradation = Join-Path -Path $ReportDir -ChildPath '_SMART_DEGRADATION.csv'  # degradation, from all reports

    if (Test-Path -Path $csvDegradation) {Remove-Item -Path $csvDegradation}

    $Degradation | Select-Object -Property $ReportSelectProperties `
    <#| Sort-Object -Property `
        @{Expression = "HostName"; Descending = $false}, `
        @{Expression = "5 Reallocated Sectors Count"; Descending = $false}, `
        @{Expression = "ScanDate"; Descending = $false} #> `
    | Export-Csv -NoTypeInformation -Path $csvDegradation


    $csvStable = Join-Path -Path $ReportDir -ChildPath '_SMART_STABLE.csv'  # full smart values, from all reports

    if (Test-Path -Path $csvStable) {Remove-Item -Path $csvStable}

    $Stable | Select-Object -Property $ReportSelectProperties `
    | Sort-Object -Property $ReportSortProperties `
    | Export-Csv -NoTypeInformation -Path $csvStable
}

#endregion


#region  # HTML

if ($html)

{
    $htmlReport = Join-Path -Path $RootDir -ChildPath (Join-Path -Path $ReportDir -ChildPath 'Consolidated Report.html')

    [xml]$htmlStableFrag = $Stable | Select-Object -Property $ReportSelectProperties `
        | Sort-Object -Property $ReportSortProperties `
        |  ConvertTo-HTML -Fragment #-PreContent '<h2><p>Stable Disk Info</p></h2>' -PostContent '<p></p>'

    [xml]$htmlDegradFrag = $Degradation `
        | Select-Object -Property $ReportSelectProperties `
        | ConvertTo-HTML -Fragment #-PreContent '<h2><p>Degradation Disk Info</p></h2>' -PostContent '<p></p>'


    for ($i = 1; $i -le $htmlDegradFrag.table.tr.count - 1; $i++)

                                {
    $key = $htmlDegradFrag.table.tr[$i].td[2]

    if ($dcteo.ContainsKey($key))

    {
        $class = $htmlDegradFrag.CreateAttribute('class')

        $class.Value = $( if ($dcteo[$key]) {'odd'} else {'even'} )

        $null = $htmlDegradFrag.table.tr[$i].attributes.Append($class)
    }
    }

                                #region  # классы таблиц

$class = $htmlDegradFrag.CreateAttribute('class')

$class.Value = 'degradation'

$null = $htmlDegradFrag.table.attributes.Append($class)


$class = $htmlStableFrag.CreateAttribute('class')

$class.Value = 'stable'

$null = $htmlStableFrag.table.attributes.Append($class)

#endregion

                                                                                                                                                                                $ConvertHtmlParams = @{
    # 'Title' = 'Python & PowerShell S.M.A.R.T. monitoring ToolKit'  # не срабатывает, поэтому напрямую вставлен в head
    'CssUri' = 'style.css'  # используется встроенная таблица стилей
    'Head' = @"
<h2>report generated: $((Get-Date).ToString())</h2>

<title>ppsmart-posh</title>

<style>
    body {
        font-family:Calibri;
        font-size:11pt;
        color:#1d6195;
    }

    table {
        width:96%;
        margin-left:2%;
        border-collapse:collapse;
        text-align: center;
    }

    td, th {
        border:0px solid black;
        border-collapse:collapse;
    }

    th {
        background-color:#226faa /*40%, dark*/;
        color:white;
    }

    td {
        padding: 4px;
        margin: 0px ;
        white-space: pre;
        color: black;
    }

    .odd  {background-color: #bfdcf2 /*85%, light*/;}
    .even {background-color: #95c5e9 /*75%, a bit more dark*/;}

    .stable tr:nth-child(odd) {background-color: #bfdcf2 /*85%, light*/;}
    .stable tr:nth-child(even) {background-color: #95c5e9 /*75%, a bit more dark*/;}
</style>
"@
    'Body' = '<h2><p>Degradation Disk Info</p></h2>' + $htmlDegradFrag.InnerXml + " `n" + '<h2><p>Stable Disk Info</p></h2>'      + $htmlStableFrag.InnerXml
    'PreContent'  = '<h3>Python & PowerShell S.M.A.R.T. monitoring ToolKit</h3>'
    'PostContent' = @'
<p>Author: Dmitry Mikhaylov</p>
<p><a href="https://github.com/mitmih/ppsmart-posh">View Project on GitHub</a></p>
'@
    }

    ConvertTo-Html @ConvertHtmlParams | Out-File $htmlReport

    Invoke-Item $htmlReport

    # $IE=new-object -com internetexplorer.application
    # $IE.navigate2($htmlReport)
    # $IE.visible=$true  #>

}

#endregion

#endregion


#region  # КОНЕЦ

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): executed' -ForegroundColor  Green
$WatchDogTimer.Stop()  # $WatchDogTimer.Elapsed.TotalSeconds

#endregion


#region  # читаем WMI-отчёты из файлов
<#foreach ($f in $WMIFiles) {
    foreach ($HardDrive in Import-Csv $f.FullName) {  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
        $Disk = (New-Object PSObject -Property @{
            ScanDate = $HardDrive.ScanDate
            HostName       = $HardDrive.HostName
            Model    = $HardDrive.Model
            SerialNumber    = Convert-hex2txt -wmisn ([string] $HardDrive.SerialNumber.Trim())
            })
        foreach ($atr in Convert-WMIArrays -data $HardDrive.WMIData -thresh $HardDrive.WMIThresholds) {  # расшифровываем атрибуты и по-одному добавляем к объекту $Disk
            try {$Disk | Add-Member -MemberType NoteProperty -Name $atr.saIDDec -Value $atr.saRaw -ErrorAction Stop}  # на случай дубликатов в отчёте
            catch {Write-Host "DUPLICATE: '$($HardDrive.HostName)' in '$($f.Name)'" -ForegroundColor Red}
        }
        $AllInfo += $Disk
    }
}#>
#endregion
