#requires -version 3  # ��������� ������ PowerShell

<#
.SYNOPSIS
    �������������� ������ .\Get-WMISMART.ps1

.DESCRIPTION
    �������� �������������� "�����" S.M.A.R.T. ������ �� ������� .\Get-WMISMART.ps1
    �� ������� ������ ��������� ��������� �����������, ������� ������� - ������� �����������

.INPUTS
    ����� � �������� �� ��������������� ������ ������

.OUTPUTS
    ����� � ���������� ���������� S.M.A.R.T.

.PARAMETER ReportDir
    ����� � ������������ ������ .\Get-WMISMART.ps1

.EXAMPLE
    .\Parse-SMART.ps1 .\output
        �������� S.M.A.R.T. �������� ������ ���������� ����������

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

    ������������� ���������
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
$TimeStart = Get-Date # ����� ������� ���������� �������
set-location "$($MyInvocation.MyCommand.Definition | split-path -parent)"  # ��������� �������� ����� "./" = ������� ���������� �������

Import-Module -Name ".\helper.psm1" -verbose  # ��������������� ������ � ���������

$WMIFiles = Get-ChildItem -Path $ReportDir -Filter '*drives.csv'  # ������ �� ������

# ������ ������ �����������, ���� ���� ����� ������ � "������" �������
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

        foreach ($atr in $AtrInfo) {$atr | Add-Member -MemberType NoteProperty -Name saFlagString -Value $(Convert-Flags -flagDec ([System.Convert]::ToInt32($atr.saFlagBin,2)))}  # ����� bin-to-char

        # ������� ����������� S.M.A.R.T. �����
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
    # Remove-Item -Path $DrivesReportFile  # ������� ����������� ����� � "������" �������
}

# ����� ������� ���������� �������
$ExecTime = [System.Math]::Round($( $(Get-Date) - $TimeStart ).TotalSeconds,1)
Write-Host "execution time is" $ExecTime "second(s)"
