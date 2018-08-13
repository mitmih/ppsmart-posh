#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS

.DESCRIPTION

.INPUTS

.OUTPUTS

.PARAMETER

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


#region  # НАЧАЛО

$psCmdlet.ParameterSetName | Out-Null

Clear-Host

$WatchDogTimer = [system.diagnostics.stopwatch]::startNew()

$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта

Import-Module -Name ".\helper.psm1" -verbose -Force  # вспомогательный модуль с функциями

#endregion

$WMIFiles = Get-ChildItem -Path (Join-Path -Path $RootDir -ChildPath $csvDir) -Filter '*drives.csv' -Recurse


#region перенос результатов сканирований из CSV-файлов в SQLite БД
$dateError = @{}

$DiskID = Get-DBHashTable -table 'Disk'
$HostID = Get-DBHashTable -table 'Host'
$ScanID = Get-DBHashTable -table 'Scan'

foreach ($f in $WMIFiles)

{
    $Scans = Import-Csv $f.FullName

    foreach ($Scan in $Scans)  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"

    {
        #region преобразование SerialNumber и ScanDate к виду yyyy.MM.dd

        try

        {
            if ($Scan.ScanDate -match '\d\d_\d\d')  # 2018-08-08_15-59
            { $Scan.ScanDate = [datetime]::ParseExact($Scan.ScanDate, 'yyyy-MM-dd_HH-mm', $null) }

            elseif ($Scan.ScanDate -match '\d* \d*:\d\d:\d\d')  # 16.07.2018 9:43:49
            { $Scan.ScanDate = [datetime]::ParseExact($Scan.ScanDate, 'dd.MM.yyyy H:mm:ss', $null) }

            elseif ($Scan.ScanDate -match '\d* \d*:\d\d')  # 2018.08.09 19:19
            { $Scan.ScanDate = [datetime]::ParseExact($Scan.ScanDate, 'yyyy.MM.dd H:mm', $null) }

            elseif ($Scan.ScanDate.Length -le 11)  # 2018.08.10
            { $Scan.ScanDate = [datetime]::ParseExact($Scan.ScanDate, 'yyyy.MM.dd', $null) }

            else
            { $Scan.ScanDate = [datetime] $Scan.ScanDate }

            $Scan.SerialNumber = (Convert-hex2txt -wmisn $Scan.SerialNumber)

        }

        catch

        {
            $dateError[$f] = $Scan.ScanDate
        }

        $Scan.ScanDate = $('{0:yyyy.MM.dd}' -f $Scan.ScanDate)

        $Scan.SerialNumber = (Convert-hex2txt -wmisn $Scan.SerialNumber)

        #endregion

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

        {
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

        {
            $sID = Update-DB -tact NewScan -obj ($Scan | Select-Object -Property `
                    @{Name="DiskID"; Expression = {$DiskID[$Scan.SerialNumber]}},
                    @{Name="HostID"; Expression = {$HostID[$Scan.HostName]}},
                    'ScanDate',
                    'WMIData',
                    'WMIThresholds',
                    @{Name="WMIStatus"; Expression = {[int][System.Convert]::ToBoolean($Scan.WMIStatus)}})  # convert string 'false' to 0, 'true' to 1

            if($sID -gt 0) {$ScanID[$skey] = $sID}
        }
    }

    Write-Host $WatchDogTimer.Elapsed.TotalSeconds "seconds`t file" $f.name -ForegroundColor  Green
}

$dateError

#endregion


#region  # КОНЕЦ

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): executed' -ForegroundColor  Green
$WatchDogTimer.Stop()  # $WatchDogTimer.Elapsed.TotalSeconds

#endregion
