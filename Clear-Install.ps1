﻿#requires -version 3  # требуемая версия PowerShell

<#
.SYNOPSIS
    начальная установка

.DESCRIPTION
    ПОСЛЕ ПОДТВЕРЖДЕНИЯ скрипт:
        УДАЛИТ ВСЕ прежние результаты работы - файл базы данных ppsmart-posh.db, папки input со списком компьютеров и output с csv-результатами сканирований, файлы библиотек System.Data.SQLite

        скачает необходимые версии библиотек System.Data.SQLite и установит их в папки x32 и x64

        инициализирует новую базу данных

        подготовит папки input и output и пример входного файла ./input/example.csv

.EXAMPLE
    powershell -File .\Clear-Install.ps1

.LINK
    github-page
        https://github.com/mitmih/ppsmart-posh

.LINK
    страница загрузки System.Data.SQLite
        https://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki

.LINK
    примеры работы с System.Data.SQLite
        https://social.technet.microsoft.com/wiki/contents/articles/30562.powershell-accessing-sqlite-databases.aspx

.NOTES
    Author: Dmitry Mikhaylov
#>

# first unblock script by right mouse menu and execute
#   Set-ExecutionPolicy RemoteSigned

#region  # НАЧАЛО

$psCmdlet.ParameterSetName | Out-Null

Clear-Host

$WatchDogTimer = [system.diagnostics.stopwatch]::startNew()

$RootDir = $MyInvocation.MyCommand.Definition | Split-Path -Parent
Set-Location $RootDir  # локальная корневая папка "./" = текущая директория скрипта

$db = Join-Path -Path $RootDir -ChildPath 'ppsmart-posh.db'

$input  = Join-Path -Path $RootDir -ChildPath 'input'
$output = Join-Path -Path $RootDir -ChildPath 'output'

$x32 = Join-Path -Path $RootDir -ChildPath 'x32'
$x64 = Join-Path -Path $RootDir -ChildPath 'x64'

if ([IntPtr]::Size -eq 8)  # 64-bit

{
    $sqlite = Join-Path -Path $RootDir -ChildPath 'x64\System.Data.SQLite.dll'
    $bit = 64
}

elseif ([IntPtr]::Size -eq 4)  # 32-bit

{
    $sqlite = Join-Path -Path $RootDir -ChildPath 'x32\System.Data.SQLite.dll'
    $bit = 32
}

else {$bit = '??'}

Write-Host -ForeGround Green "environment`t$bit-bit`n"
Write-Host -ForeGround Green "PowerShell`t$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)`n"
Write-Host -ForeGround Green ".Net`t`t$($PSVersionTable.CLRVersion.Major).$($PSVersionTable.CLRVersion.Minor)`n"



#endregion


#region  # подтверждение и очистка

$confirmation = Read-Host "Если вы продолжите, все данные будут очищены`nНажмите y(es) для продолжения..."

try

{
    if ($confirmation.ToLower() -notin @('y','yes')) {exit}

    if (Test-Path -Path $db)     {Remove-Item -Path $db              -Force}

    if (Test-Path -Path $input)  {Remove-Item -Path $input  -Recurse -Force}
    if (Test-Path -Path $output) {Remove-Item -Path $output -Recurse -Force}

    if (Test-Path -Path $x32)    {Remove-Item -Path $x32    -Recurse -Force}
    if (Test-Path -Path $x64)    {Remove-Item -Path $x64    -Recurse -Force}

    Get-ChildItem -Filter "*.zip" -Path $RootDir | Remove-Item -Force
}

catch

{
    Write-Host "Remove-Item failed... Please, reopen your console/ISE and run script again"
}

#endregion


#region  # new folders: input, output

# новая папка, соответствует битности библиотек
if(!(Test-Path $input))

{
    $null = New-Item -ItemType Directory -Path $input
}

(New-Object psobject -Property @{HostName = $env:COMPUTERNAME;}) | Export-Csv -Path (Join-Path -Path $input -ChildPath "example.csv") -NoTypeInformation -Encoding UTF8

if(!(Test-Path $output))

{
    $null = New-Item -ItemType Directory -Path $output
}

#endregion


#region  # setup System.Data.SQLite dll`s

    #region определяем, какие архивы нужно скачать, в зависимости от $PSVersionTable.CLRVersion

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
        #region download
        $url = $links[$k]
        $file = Join-Path -Path $RootDir -ChildPath $links[$k].Split('/')[-1]

        if(!(Test-Path $file))

        {
            (New-Object System.Net.WebClient).DownloadFile($links[$k], $file)
            # Start-BitsTransfer -Source $links[$k] -Destination $links[$k].Split('/')[-1] #-Asynchronous
        }

        #endregion


        # новая папка, соответствует битности библиотек
        if(!(Test-Path -Path (Join-Path -Path $RootDir -ChildPath $k))) {$null = New-Item -ItemType Directory -Path (Join-Path -Path $RootDir -ChildPath $k)}


        #region распаковываем SQLite.Interop.dll и System.Data.SQLite.dll

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
            Write-Host "unziping failed... try to download and unzip files manually" -ForeGround Yellow
            Write-Host "page:`t`t$page`nfile:`t`t$($links[$k].Split('/')[-1])`nunzip to:`t$k`n`n" -ForeGround Magenta
        }

        #endregion


        #region проверка распаковки библиотек

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


#region  # init ppsmart-posh.db

try

{
    Add-Type -Path $sqlite -ErrorAction Stop
}

catch

{
    Write-Host "Importing SQLite assemblies by path '$sqlite' failed..."
}

try
{
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
	    `ID`		INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE,
	    `HostName`	TEXT NOT NULL UNIQUE
    );

CREATE TABLE `Scan` (
	`ID`			INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE,
	`Archived`		INTEGER DEFAULT 0,
	`DiskID`		INTEGER NOT NULL DEFAULT 0,
	`HostID`		INTEGER NOT NULL DEFAULT 0,
	`ScanDate`		TEXT NOT NULL,
	`WMIStatus`		INTEGER DEFAULT -1,
	`WMIData`		TEXT NOT NULL,
	`WMIThresholds`	TEXT
);

    --CREATE UNIQUE INDEX `IndexDisk` ON `Disk` (`SerialNumber`);
    --CREATE UNIQUE INDEX `IndexHost` ON `Host` (`HostName`);
    CREATE UNIQUE INDEX `IndexScan` ON `Scan` (`DiskID`,`HostID`,`ScanDate`);
'@
    
    if ($sql.ExecuteNonQuery() -eq 0)

    {
        Write-Output 'New DataBase is ok'
    }

    else
    
    {
        Write-Output 'New DataBase: "$sql.ExecuteNonQuery()" failed'
    }
}

catch {}

finally
{
    if ($sql -ne $null)        { $sql.Dispose() }
    if ($con.State -eq 'Open') { $con.Close()   }
}

#endregion


#region  # КОНЕЦ

Write-Host $WatchDogTimer.Elapsed.TotalSeconds 'second(s): executed' -ForegroundColor  Green
$WatchDogTimer.Stop()  # $WatchDogTimer.Elapsed.TotalSeconds

#endregion
