#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    расшифровывает отчёты .\Get-WMISMART.ps1

.DESCRIPTION
    сценарий расшифровывает "сырые" S.M.A.R.T. данные из отчётов .\Get-WMISMART.ps1
    по каждому отчёту формирует подробную расшифровку, сколько отчётов - столько расшифровок

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
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = Get-Date # замер времени выполнения скрипта
set-location "$($MyInvocation.MyCommand.Definition | split-path -parent)"  # локальная корневая папка "./" = текущая директория скрипта

Import-Module -Name ".\helper.psm1" -verbose  # вспомогательный модуль с функциями

$WMIFiles = Get-ChildItem -Path $ReportDir -Filter '*drives.csv'  # отчёты по дискам

# удалим старые расшифровки, если есть новые отчёты с "сырыми" данными
if ($WMIFiles -ne $null) {foreach ($SmartReportFile in Get-ChildItem -Path $ReportDir -Filter '*smart.csv') {Remove-Item -Path "$ReportDir\$SmartReportFile"}}

foreach ($DrivesReportFile in $WMIFiles) {
    $DrivesReportFile = "$ReportDir\$DrivesReportFile"
    $Drives = Import-Csv "$DrivesReportFile"  # "ScanDate","HostName","SerialNumber","Model","Size","InterfaceType","MediaType","DeviceID","PNPDeviceID","WMIData","WMIThresholds","WMIStatus"
    $AtrInfo = $null
    foreach ($disk in $Drives) {
        $AtrInfo = Convert-WMIArrays -data $disk.WMIData -thresh $disk.WMIThresholds
        $AtrInfo | Add-Member -MemberType NoteProperty -Name ScanDate  -Value $disk.ScanDate
        $AtrInfo | Add-Member -MemberType NoteProperty -Name PC        -Value $disk.HostName
        $AtrInfo | Add-Member -MemberType NoteProperty -Name Model     -Value $disk.Model
        $AtrInfo | Add-Member -MemberType NoteProperty -Name SerNo     -Value (Convert-hex2txt -wmisn ([string] $disk.SerialNumber.Trim()))  # $disk.SerialNumber
        $AtrInfo | Add-Member -MemberType NoteProperty -Name WMIStatus -Value $disk.WMIStatus

        foreach ($atr in $AtrInfo) {$atr | Add-Member -MemberType NoteProperty -Name saFlagString -Value $(Convert-Flags -flagDec ([System.Convert]::ToInt32($atr.saFlagBin,2)))}  # флаги bin-to-char

        # экспорт расшифровки S.M.A.R.T. диска
        $AtrInfo | Select-Object `
        'ScanDate',`
        'PC',`
        'Model',`
        'SerNo',`
        'WMIStatus',`
        'saIDHex',`
        'saIDDec',`
        'saName',`
        'saRaw',`
        'saValue',`
        'saWorst',`
        'saThreshold',`
        'saRawDec',`
        'saRawHex',`
        'saFullHex',`
        'saFullDec',`
        'saFlagDec',`
        'saFlagBin',`
        'saFlagString'`
        | Export-Csv -Append -NoTypeInformation -Path $DrivesReportFile.Replace('drives', 'smart')
    }
    # Remove-Item -Path $DrivesReportFile  # удаляем прочитанный отчёт с "сырыми" данными
}

# замер времени выполнения скрипта
$ExecTime = [System.Math]::Round($( $(Get-Date) - $TimeStart ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
