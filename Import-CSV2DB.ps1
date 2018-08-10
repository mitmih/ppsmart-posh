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
$TimeStart = @(Get-Date) # замер времени выполнения скрипта
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта
#endregion

Import-Module -Name ".\helper.psm1" -verbose  # вспомогательный модуль с функциями

$WMIFiles = Get-ChildItem -Path (Join-Path -Path $RootDir -ChildPath $csvDir) -Filter '*drives.csv' -Recurse

#region заполним словари из базы
$DiskID = Get-DBHashTable -table 'Disk'
$HostID = Get-DBHashTable -table 'Host'
#endregion

$p2 = Measure-Command {
#region перенос списка хостов из CSV-файлов в SQLite БД
foreach ($f in Get-ChildItem -Path (Join-Path -Path $RootDir -ChildPath 'input') -Filter '*.csv' -Recurse)
{
    foreach ($line in Import-Csv $f.FullName)
    {
        if (!$HostID.ContainsKey($line.HostName))
        {  # new record
            $hID = Add-DBHost -obj (New-Object psobject -Property @{HostName = $line.HostName;ScanDate = 0;})
            if ($hID -gt 0) {$HostID[$line.HostName] = $hID}
        }
    }
}
#endregion
} | select -ExpandProperty TotalSeconds

$p3 = Measure-Command {
#region перенос результатов сканирований из CSV-файлов в SQLite БД
foreach ($f in $WMIFiles) {
    foreach ($Scan in Import-Csv $f.FullName) {  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
        # Disk
        $Scan.SerialNumber = (Convert-hex2txt -wmisn $Scan.SerialNumber)
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
}  #endregion
} | select -ExpandProperty TotalSeconds

#region  # КОНЕЦ
# замер времени выполнения скрипта
$TimeStart += Get-Date
$ExecTime = [System.Math]::Round($( $TimeStart[-1] - $TimeStart[0] ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
#endregion

# Write-Host -ForegroundColor Green $p1, 'region заполним словари из базы'
Write-Host -ForegroundColor Green $p2, 'region перенос списка хостов из CSV-файлов в SQLite БД'
Write-Host -ForegroundColor Green $p3, 'region перенос результатов сканирований из CSV-файлов в SQLite БД'
