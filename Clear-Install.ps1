#requires -version 3  # ��������� ������ PowerShell

<#
.SYNOPSIS
    ��������� ���������

.DESCRIPTION
    ����� ������������� ������:
    ������ ��� ������ ������
    ������� ����������� ������ ��������� System.Data.SQLite
    ��������� �� � ����� x32 � x64
    �������������� ��
    ���������� ����� input � output
    ���������� ������ �������� ����� ./input/example.csv

.EXAMPLE
    ./Clear-Install.ps1

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

.LINK
    System.Data.SQLite downloads page
        https://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki

.LINK
    some examples
        https://social.technet.microsoft.com/wiki/contents/articles/30562.powershell-accessing-sqlite-databases.aspx

.NOTES
    Author: Dmitry Mikhaylov
#>

# first unblock script by right mouse menu and execute
#   Set-ExecutionPolicy RemoteSigned

#region  # ������
$psCmdlet.ParameterSetName | Out-Null
Clear-Host
$TimeStart = @(Get-Date) # ����� ������� ���������� �������
$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # ��������� �������� ����� "./" = ������� ���������� �������

$db = Join-Path -Path $RootDir -ChildPath 'ppsmart-posh.db'
$input = Join-Path -Path $RootDir -ChildPath 'input'
$output = Join-Path -Path $RootDir -ChildPath 'output'
$x32 = Join-Path -Path $RootDir -ChildPath 'x32'
$x64 = Join-Path -Path $RootDir -ChildPath 'x64'

if ([IntPtr]::Size -eq 8)  # 64-bit
{
    Write-Host "64-bit" -ForeGround Green
    $sqlite = Join-Path -Path $RootDir -ChildPath 'x64\System.Data.SQLite.dll'
}
elseif ([IntPtr]::Size -eq 4)  # 32-bit
{
    Write-Host "32-bit" -ForeGround Green
    $sqlite = Join-Path -Path $RootDir -ChildPath 'x32\System.Data.SQLite.dll'
}
else {Write-Host 'can not choose between 32 or 64 bit dll`s'}
#endregion

#region ����������� ��������
$confirmation = Read-Host "���� �� ����������, ��� ������ ����� �������`n������� y(es) ��� �����������..."
if ($confirmation.ToLower() -notin @('y','yes')) {exit}
if (Test-Path -Path $db)     {Remove-Item -Path $db              -Force}
if (Test-Path -Path $input)  {Remove-Item -Path $input  -Recurse -Force}
if (Test-Path -Path $output) {Remove-Item -Path $output -Recurse -Force}
if (Test-Path -Path $x32)    {Remove-Item -Path $x32    -Recurse -Force}
if (Test-Path -Path $x64)    {Remove-Item -Path $x64    -Recurse -Force}
Get-ChildItem -Filter "*.zip" -Path $RootDir | Remove-Item -Force
#endregion

#region setup SQLite

    #region ����������, ����� ������ ����� �������, � ����������� �� $PSVersionTable.CLRVersion
    $links = @{}
    $url = "https://system.data.sqlite.org"  # root-page
    $page = "https://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki"  # downloads are listed here

    foreach ($l in $(Invoke-WebRequest -Uri $page).Links | ?{$_.href -match ".*static-binary-Win32.*" -or $_.href -match ".*static-binary-x64.*"} | select -ExpandProperty href)
    {
        if ("$($PSVersionTable.CLRVersion.Major)$($PSVersionTable.CLRVersion.Minor)" -like $l.Split('/')[3].split('-')[1].ToLower().replace('netfx',''))  # 20 35 40 451 45 46
        {
            $links[$l.Split('/')[3].split('-')[4].ToLower().replace('win32','x32')] = "$url$l".Replace('downloads','blobs')  # https://system.data.sqlite.org/blobs/1.0.108.0/sqlite-netFx40-static-binary-x64-2010-1.0.108.0.zip  # real download link
        }
    }
    #endregion

    foreach ($k in $links.Keys)
    {
        #region ������ ������
        $url = $links[$k]
        $file = Join-Path -Path $RootDir -ChildPath $links[$k].Split('/')[-1]

        if(!(Test-Path $file))
        {  #download
            (New-Object System.Net.WebClient).DownloadFile($links[$k], $file)
            #Start-BitsTransfer -Source $links[$k] -Destination $links[$k].Split('/')[-1] #-Asynchronous
        }
        #endregion

        # ����� �����, ������������� �������� ���������
        if(!(Test-Path -Path (Join-Path -Path $RootDir -ChildPath $k))) {$null = New-Item -ItemType Directory -Path (Join-Path -Path $RootDir -ChildPath $k)}

        #region ������������� SQLite.Interop.dll � System.Data.SQLite.dll
        try
        {
            Add-Type -Assembly System.IO.Compression.FileSystem
            $zip = [IO.Compression.ZipFile]::OpenRead($file)
            foreach ($f in $zip.Entries | where {$_.Name -like "SQLite.Interop.dll" -or $_.Name -like "System.Data.SQLite.dll"})
            {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($f, (Join-Path -Path $RootDir -ChildPath (Join-Path -Path $k -ChildPath $f)), $true)
            }
            $zip.Dispose()
        }
        catch
        {
            Write-Host "Ops... try to download and unzip files manually" -ForeGround Yellow
            Write-Host "page:`t`t$page`nfile:`t`t$($links[$k].Split('/')[-1])`nunzip to:`t$k`n`n" -ForeGround Magenta
        }
        #endregion

        #region �������� ���������� ���������
        $subdir = Join-Path -Path $RootDir -ChildPath $k
        if ((Test-Path (Join-Path -Path $RootDir -ChildPath (Join-Path -Path $k -ChildPath "SQLite.Interop.dll"))) -and
            (Test-Path (Join-Path -Path $RootDir -ChildPath (Join-Path -Path $k -ChildPath "System.Data.SQLite.dll"))))
        {
            Get-ChildItem -Path $subdir | Select-Object -Property FullName, LastAccessTime
            Get-ChildItem -Filter "*.zip" -Path $RootDir | Remove-Item -Force
        }
        #endregion
    }

#endregion

#region init ppsmart-posh.db
try   {Add-Type -Path $sqlite -ErrorAction Stop}
catch {Write-Host "Importing the SQLite assemblies, '$sqlite', failed..."}

$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$db"
$con.Open()

$sql = $con.CreateCommand()
$sql.CommandText = @'
CREATE TABLE `Disk` (
    `ID`			INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE,
    `SerialNumber`	TEXT NOT NULL UNIQUE,
    `Model`			TEXT NOT NULL,
    `Size`			INTEGER NOT NULL,
    `InterfaceType`	TEXT,
    `MediaType`		TEXT,
    `DeviceID`		TEXT,
    `PNPDeviceID`	TEXT
);

CREATE TABLE `Host` (
	`ID`			INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE,
	`HostName`		TEXT NOT NULL UNIQUE,
	`LastScan`		TEXT
);

CREATE TABLE `Scan` (
    `ID`			INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE,
    `DiskID`		INTEGER NOT NULL DEFAULT 0,
    `HostID`		INTEGER NOT NULL DEFAULT 0,
    `ScanDate`		TEXT NOT NULL,
    `WMIData`		TEXT NOT NULL,
    `WMIThresholds`	TEXT,
    `WMIStatus`		INTEGER
);

CREATE UNIQUE INDEX `SN` ON `Disk` (`SerialNumber`);
CREATE UNIQUE INDEX `HN` ON `Host` (`HostName`);
'@
# $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
# $data = New-Object System.Data.DataSet
[void] (New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql).Fill((New-Object System.Data.DataSet))
$sql.Dispose()
$con.Update
$con.Close()
#endregion

#region input, output
# ����� �����, ������������� �������� ���������
if(!(Test-Path $input))  {$null = New-Item -ItemType Directory -Path $input }
(New-Object psobject -Property @{HostName = $env:COMPUTERNAME;LastScan = "never";}) | Select-Object "HostName", "LastScan" | Export-Csv -Path (Join-Path -Path $input -ChildPath "example.csv") -NoTypeInformation -Encoding UTF8

if(!(Test-Path $output)) {$null = New-Item -ItemType Directory -Path $output}
#endregion

#region
#endregion

#region
#endregion
