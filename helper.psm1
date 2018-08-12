$dctSMART = @{
    1='Read Error Rate'
    2='Throughput Performance'
    3='Spin-Up Time'
    4='Start/Stop Count'
    5='Reallocated Sectors Count'
    6='Read Channel Margin'
    7='Seek Error Rate'
    8='Seek Time Performance'
    9='Power-On Hours'
    10='Spin Retry Count'
    11='Recalibration Retries or Calibration Retry Count'
    12='Power Cycle Count'
    13='Soft Read Error Rate'
    22='Current Helium Level'
    170='Available Reserved Space'
    171='SSD Program Fail Count'
    172='SSD Erase Fail Count'
    173='SSD Wear Leveling Count'
    174='Unexpected power loss count'
    175='Power Loss Protection Failure'
    176='Erase Fail Count'
    177='Wear Range Delta'
    179='Used Reserved Block Count Total'
    180='Unused Reserved Block Count Total'
    181='Program Fail Count Total or Non-4K Aligned Access Count'
    182='Erase Fail Count'
    183='SATA Downshift Error Count or Runtime Bad Block'
    184='End-to-End error / IOEDC'
    185='Head Stability'
    186='Induced Op-Vibration Detection'
    187='Reported Uncorrectable Errors'
    188='Command Timeout'
    189='High Fly Writes'
    190='Temperature Difference or Airflow Temperature'
    191='G-sense Error Rate'
    192='Power-off Retract Count, Emergency Retract Cycle Count (Fujitsu), or Unsafe Shutdown Count'
    193='Load Cycle Count or Load/Unload Cycle Count (Fujitsu)'
    194='Temperature or Temperature Celsius'
    195='Hardware ECC Recovered'
    196='Reallocation Event Count'
    197='Current Pending Sector Count'
    198='(Offline) Uncorrectable Sector Count'
    199='UltraDMA CRC Error Count'
    200='Multi-Zone Error Rate / Write Error Rate (Fujitsu)'
    201='Soft Read Error Rate or TA Counter Detected'
    202='Data Address Mark errors or TA Counter Increased'
    203='Run Out Cancel'
    204='Soft ECC Correction'
    205='Thermal Asperity Rate'
    206='Flying Height'
    207='Spin High Current'
    208='Spin Buzz'
    209='Offline Seek Performance'
    210='Vibration During Write'
    211='Vibration During Write'
    212='Shock During Write'
    220='Disk Shift'
    221='G-Sense Error Rate'
    222='Loaded Hours'
    223='Load/Unload Retry Count'
    224='Load Friction'
    225='Load/Unload Cycle Count'
    226='Load In-time'
    227='Torque Amplification Count'
    228='Power-Off Retract Cycle'
    230='GMR Head Amplitude (magnetic HDDs), Drive Life Protection Status (SSDs)'
    231='Life Left (SSDs) or Temperature'
    232='Endurance Remaining or Available Reserved Space'
    233='Media Wearout Indicator (SSDs) or Power-On Hours'
    234='Average erase count AND Maximum Erase Count'
    235='Good Block Count AND System(Free) Block Count'
    240='Head Flying Hours or Transfer Error Rate (Fujitsu)'
    241='Total LBAs Written'
    242='Total LBAs Read'
    243='Total LBAs Written Expanded'
    244='Total LBAs Read Expanded'
    249='NAND Writes (1GiB)'
    250='Read Error Retry Rate'
    251='Minimum Spares Remaining'
    252='Newly Added Bad Flash Block'
    253='Free Fall Protection'
    999='testName'
}  # ������� ���������


function Convert-WMIArrays ([string] $data, [string] $thresh) {
<#
.SYNOPSIS
    ������� ��������� ����� ������� S.M.A.R.T.-������ ����� � ����� �������� ������

.DESCRIPTION
    ������� ������ �������, ������������ ������ � ����������� �� �������� � ���������� ���� ����� ��������� �������� �����

.INPUTS
    512-���� ������ � ������� �� ������ MSStorageDriver_FailurePredictData
    512-���� ������ � ������� �� ������ MSStorageDriver_FailurePredictThresholds

.OUTPUTS
    ������ ��������-��������� S.M.A.R.T.

.PARAMETER data
    512-���� ������ � ������� �� ������ MSStorageDriver_FailurePredictData

.PARAMETER thresh
    512-���� ������ � ������� �� ������ MSStorageDriver_FailurePredictThresholds

.EXAMPLE
    $AtrInfo = Convert-WMIArrays -data $WMIData -thresh $WMIThresholds

.LINK
    ��������� 512-���� ������� ������ S.M.A.R.T. (���������� ��������)
        http://www.t13.org/Documents/UploadedDocuments/docs2005/e05171r0-ACS-SMARTAttributes_Overview.pdf

.NOTES
    Author: Dmitry Mikhaylov

    ������ 512-���� ������� 12�� ��������� ������� ���������� �� �������� +2
    �.�. ������ ��� ����� �������� ������� �������������� ��������� S.M.A.R.T.-������ (vendor-specific)

    ��������� 512-���� S.M.A.R.T.-������� ������
        Offset      Length (bytes)  Description
        0           2               SMART structure version (this is vendor-specific)
        2           12              Attribute entry 1
        2+(12)      12              Attribute entry 2
        ...         ...             ...
        2+(12*29)   12              Attribute entry 30

    ��������� 12-��������� �����
        0     Attribute ID (dec)
        1     Status flag (dec)
        2     Status flag (Bits 6�15 � Reserved)
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
    if ($data.Length -eq 512 ) {  # �������� �� ��������, ������ 12-�������� ������ �������� �� ���� �������
        for ($i = 2; $i -le 350; $i = $i + 12) {  # ����� ����� ���� 29 ���������, ����� �������� ����� ������� 2+12*29 = 350 ����, ���������� ����� ����� ������ � ����� ������� ���������� ��� ���������� � PropertyNote ������� �����
            if ($data[$i] -eq 0) {continue}  # ���������� �������� � ������� �����. Thresholds ��� ���� � ������ �� ����, �.�. � �����-�� ����� ������� �� � ���� ������� ����� ���� ����

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

            $raw = [Convert]::ToInt64("$h11 $h10 $h9 $h8 $h7 $h6 $h5".Replace(' ',''),16)  # ��������� �� 16-������ � 10-������ ������� ���������

            if ($data[$i] -in @(190, 194)) {$raw = $data[$i+5]}  # ����������� ��������� ��� ����
            if ($data[$i] -in @(9)) {$raw = [Convert]::ToInt64("$h6 $h5".Replace(' ',''),16)}  # ��������� � �����

            $Result += (New-Object PSObject -Property @{
                saIDDec     = [int]    $data[$i]
                saIDHex     = '{0:x2}' -f $data[$i]
                saName      = [string] $(if ($dctSMART.ContainsKey([int] $data[$i])) {$dctSMART[[int] $data[$i]]} else {$null})  # ����������� �� ������� $dctSMART
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


function Convert-Flags ([int] $flagDec) {
<#
.SYNOPSIS
    ������� ������������ ���������� ������ ����� � ������, ������������ ������������� ���� ������ ��������

.DESCRIPTION
    ����������� ���������� ��������� ���������� ����� � �����, ������ ��� ����� � ���������� �������
    �.�. �� ������ ������ ������-���� �������� ����� 6 ��� (���� � 6�� �� 15� ���������������), ������� ��������� ������ ��� ����

.INPUTS
    ���������� �������� ������-����� (2� � 3� ����� �� 12��-��������� ����� S.M.A.R.T.-��������)

.OUTPUTS
    ������ �� ����� ��������, ���������� ���������, ����. "- C - P O W", ��� ����� �������� ��� ���� ����������, � ������� - ���������� �����
    S - Self-preserving - ������� ����� �������� ������, ���� ���� S.M.A.R.T. ��������
    C - Event count - ������� �����������
    R - Rate Error - ���������� ������� ������
    P - Performance - ���������� ������������������ �����
    O - Online / Offline - ������� ����������� � �� ����� ������ (��-����) � � �������� (���-����)
    W - Warranty / Prefailure - ����� �������� �������� ��������� ����������, ���� ��������� � ��������� 24 ���� � ���� ����� ���� ������� � ������� ������������ �����

.PARAMETER flagDec
    ���������� �������� ������-�����

.EXAMPLE
    $MyStringFlag = Convert-Flags -flagDec 255

.LINK
    ��������� 512-���� ������� ������ S.M.A.R.T. (���������� ��������)
        http://www.t13.org/Documents/UploadedDocuments/docs2005/e05171r0-ACS-SMARTAttributes_Overview.pdf
    ��������� � �����
        https://www.micron.com/~/media/documents/products/technical-note/solid-state-storage/tnfd21_m500-mu02_smart_attributes.pdf
        https://www.micron.com/~/media/documents/products/technical-note/solid-state-storage/tnfd22_client_ssd_smart_attributes.pdf
        https://www.micron.com/~/media/documents/products/technical-note/solid-state-storage/tnfd34_5100_ssd_smart_implementation.pdf

    ��� ������� ��� �����
        https://white55.ru/smart.html
        https://www.hdsentinel.com/smart/index.php
        http://sysadm.pp.ua/linux/monitoring-systems/smart-attributes.html

.NOTES
    Author: Dmitry Mikhaylov

    Status flag ����:
        0 Bit   Warranty - Prefailure/advisory bit. Applicable only when the current value is less than or equal to its threshold.
            0 = Advisory: the device has exceeded its intended design life; the failure is not covered under the drive warranty.
            1 = Prefailure: warrantable, failure is expected in 24 hours and is covered in the drive warranty.

        1 Bit   Offline - Online collection bit
            0 = Attribute is updated only during off-line activities
            1 = Attribute is updated during both online and off-line activities.

        2 Bit   Performance bit
            0 = Not a performance attribute.
            1 = Performance attribute.

        3 Bit   Error Rate bit - Expected, non-fatal errors that are inherent in the device.
            0 = Not an error rate attribute.
            1 = Error rate attribute.

        4 Bit   Event count bit
            0 = Not an event count attribute.
            1 = Event count attribute.

        5 Bit   Self-preserving bit - The attribute is collected and saved by the drive without host intervention.
            0 = Not a self-preserving attribute.
            1 = Self-preserving attribute.

        Bits 6�15 � Reserved
#>
    #foreach ($m in @(1,2,4,8,16,32)) {Write-Host ([System.Convert]::ToString($flagDec,2)), ([convert]::ToString($m,2)), (($flagDec -band $m) -gt 0)}  # for debug
    $b0 = if (($flagDec -band   1) -gt 0) {'W'} else {'-'}  # bitmask 000001
    $b1 = if (($flagDec -band   2) -gt 0) {'O'} else {'-'}  # bitmask 000010
    $b2 = if (($flagDec -band   4) -gt 0) {'P'} else {'-'}  # bitmask 000100
    $b3 = if (($flagDec -band   8) -gt 0) {'R'} else {'-'}  # bitmask 001000
    $b4 = if (($flagDec -band  16) -gt 0) {'C'} else {'-'}  # bitmask 010000
    $b5 = if (($flagDec -band  32) -gt 0) {'S'} else {'-'}  # bitmask 100000
    return "$b5 $b4 $b3 $b2 $b1 $b0"
}


function Convert-hex2txt ([string] $wmisn) {
<#
.SYNOPSIS
    ������� ������������ �������� ����� ������� ����� �� ������������������ � ��������� ������

.DESCRIPTION
    �� Windows 7 ������ � WMI �������� ����� � ����������������� �������, ������ ����� ��-�����
    �������� �������, � �� Windows 10 - � ���������, ����������� � ������� �� �������� ��

    ������� ��������� ������ ��������� ������
    ��� ������������� ������������ ����������������� �������� ����� � �����, ��������� ������� ��������

    ��������: win7 format                        -> win10 format
        56394b4d37345130202020202020202020202020 -> 9VMK470Q
        32535841394a5a46354138333932202020202020 -> S2AXJ9FZA53829
        5639514d42363439202020202020202020202020 -> 9VMQ6B94


.INPUTS
    ������ � �������� ������� �� �� WMI

.OUTPUTS
    ��������������� ������ ������ � �������� �������

.PARAMETER wmisn
    ������ � �������� ������� �� WMI

.EXAMPLE
    $TXTSerialNumber = Convert-hex2txt -wmisn $WMISerialNumber

.LINK
    https://blogs.technet.microsoft.com/heyscriptingguy/2011/09/09/convert-hexadecimal-to-ascii-using-powershell/

.NOTES
    Author: Dmitry Mikhaylov

#>


    if ($wmisn.Length -eq 40) {  # �������� �� �����
        $txt = ""
        for ($i = 0; $i -lt 40; $i = $i + 4) {
            $txt = $txt + [CHAR][CONVERT]::toint16("$($wmisn[$i+2])$($wmisn[$i+3])",16) + [CHAR][CONVERT]::toint16("$($wmisn[$i+0])$($wmisn[$i+1])",16)
        }

    } else {$txt = $wmisn}

    return $txt.Trim()
}


function Get-DBHashTable ([string] $table) {
<#
.SYNOPSIS
    ���������� ������ �� ������� ���� ������ � ���� �������

.DESCRIPTION
    ������������ SQL-������� � �������� Disk, Host
    ��������� ������� � ����
        'HostName' = ID
        'SerialNumber' = ID

.INPUTS
    ��� �������

.OUTPUTS
    ������ ���� HashTable

.PARAMETER table
    ��� ������� ��� ������������ �������

.EXAMPLE
    $dctDisk = Get-DBHashTable -table 'Disk'

.LINK

.NOTES
    Author: Dmitry Mikhaylov
#>

try
{

    # �������� �������� ����� ���������� ��� ����������� ���������� ����������
    if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x64\System.Data.SQLite.dll' -ErrorAction Stop}  # 64-bit
    if ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x32\System.Data.SQLite.dll' -ErrorAction Stop}  # 32-bit
    
    # ����������� ���������� ��� ������ � sqlite
    Add-Type -Path $sqlite -ErrorAction Stop

    # �������� ���������� � ��
    $db = Join-Path -Path $PSScriptRoot -ChildPath 'ppsmart-posh.db' -ErrorAction Stop
    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source=$db"
    $con.Open()

    # ���������� SQL-�������, ��������� ����������� � ����������
    $sql = $con.CreateCommand()
    if ($table -eq 'Disk') {$sql.CommandText = "SELECT ID, SerialNumber AS TXT FROM Disk;"}
    if ($table -eq 'Host') {$sql.CommandText = "SELECT ID, HostName AS TXT FROM Host;"}
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data = New-Object System.Data.DataSet
    [void]$adapter.Fill($data)
    $sql.Dispose()
}
catch
{
    return $null
}


$dct = @{}
foreach ($r in $data.Tables.Rows) {$dct[$r.TXT] = $r.ID}

return $dct
}


function Update-DB (
    [parameter(Mandatory=$true, ValueFromPipeline=$false)] [string]$tact = $null,
    [parameter(Mandatory=$true, ValueFromPipeline=$false)]         $obj  = $null,
    [parameter(Mandatory=$false, ValueFromPipeline=$false)]   [int]$id   = $null)
{
<#
.SYNOPSIS
    ���������/��������� ������ � ������� ��

.DESCRIPTION
    ��� ���������� ���������� primary key ����� ������
    ��� ���������� ������������ ������ ���������� $null

.INPUTS
    ������ � ������������ ������������ ������� �����

.OUTPUTS
    ID ����� ������ ��� $null

.PARAMETER obj
    ������, ���������� Properties ��� SQL-������� � ��

.PARAMETER tact
    table action - �������� SQL-�������, ��������� �������:
        NewHost / UpdHost
        NewDisk / UpdDisk
        NewScan / UpdScan

.PARAMETER id
    primary key ������, ������� ����� ��������

.EXAMPLE
    $hID = Update-DB -tact NewHost -obj ($Scan | Select-Object -Property 'HostName')
        �������� ����� ������

.EXAMPLE
    $null= Update-DB -tact UpdHost -obj ($Scan | Select-Object -Property 'HostName') -id $HostID[$Scan.HostName]
        �������� ������ �� � primary key

.LINK

.NOTES
    Author: Dmitry Mikhaylov
#>

    $newrecID = $null

    $dctQuery = @{
        'NewHost' = 'INSERT OR IGNORE INTO `Host` (`HostName`) VALUES (@HostName);'

        'UpdHost' = 'UPDATE `Host` SET `ScanDate` = @ScanDate, `Ping` = @Ping, `WMIInfo` = @WMIInfo WHERE ID = @ID;'

        'NewDisk' = @'
            INSERT OR IGNORE INTO `Disk` (
                `SerialNumber`,
                `Model`,
                `Size`,
                `InterfaceType`,
                `MediaType`,
                `DeviceID`,
                `PNPDeviceID`
            ) VALUES (
                @SerialNumber,
                @Model,
                @Size,
                @InterfaceType,
                @MediaType,
                @DeviceID,
                @PNPDeviceID
            );
'@

        'UpdDisk' = ''

        'NewScan' = @'
            INSERT OR IGNORE INTO `Scan` (
                `DiskID`,
                `HostID`,
                `ScanDate`,
                `WMIData`,
                `WMIThresholds`,
                `WMIStatus`
            ) VALUES (
                @DiskID,
                @HostID,
                @ScanDate,
                @WMIData,
                @WMIThresholds,
                @WMIStatus
            );
'@

        'UpdScan' = ''
    }


    try

    {
        # �������� �������� ����� ���������� ��� ����������� ���������� ����������
        if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x64\System.Data.SQLite.dll' -ErrorAction Stop}  # 64-bit
        if ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x32\System.Data.SQLite.dll' -ErrorAction Stop}  # 32-bit
    
        # ����������� ���������� ��� ������ � sqlite
        Add-Type -Path $sqlite -ErrorAction Stop

        # �������� ���������� � ��
        $db = Join-Path -Path $PSScriptRoot -ChildPath 'ppsmart-posh.db' -ErrorAction Stop
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=$db"
        $con.Open()

        $sql = $con.CreateCommand()
        $sql.CommandText = $dctQuery[$tact]
        
        $objValidProperties = $obj | Get-Member -MemberType Properties -ErrorAction Stop
    }
    
    catch
    
    {
        if($sql -ne $null) { $sql.Dispose() }
        if ($con.State -eq 'Open') { $con.Close() }

        return $newrecID
    }


    foreach ($p in $objValidProperties)
        
    {
        $null = $sql.Parameters.AddWithValue("@$($p.Name)", ($obj | select -ExpandProperty $p.Name) )
    }


    if ($id)  # update record (������� primary key)
    {
        $null = $sql.Parameters.AddWithValue("@ID", $id)
        $null = $sql.ExecuteNonQuery()
    }

    else  # add new record

    {
        if ($sql.ExecuteNonQuery()) { $newrecID = $sql.Connection.LastInsertRowId }
    }

    return $newrecID
}


function Get-DBData (
    [parameter(Mandatory=$true, ValueFromPipeline=$false)] [string]$Query = $null,
    [parameter(Mandatory=$true, ValueFromPipeline=$false)] [string]$base = $null
)
{
<#
.SYNOPSIS
    ���������� ��������� SQL-�������

.DESCRIPTION
    ���������� ��������� SQL-�������

.INPUTS
    ������ � SQL-��������, sqlite ���� ������

.OUTPUTS
    ����� �������

.PARAMETER Query
    SQL-������

.PARAMETER base
    ���� � ����� ��

.EXAMPLE
    $data = Get-DBData -Query 'SELECT * FROM Scan' -base 'ppsmart-posh.db'

.LINK

.NOTES
    Author: Dmitry Mikhaylov
#>

    $DataSet = $null

    try

    {
        # �������� �������� ����� ���������� ��� ����������� ���������� ����������
        if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x64\System.Data.SQLite.dll' -ErrorAction Stop}  # 64-bit
        if ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x32\System.Data.SQLite.dll' -ErrorAction Stop}  # 32-bit
    
        # ����������� ���������� ��� ������ � sqlite
        Add-Type -Path $sqlite -ErrorAction Stop

        # �������� ���������� � ��
        $db = Join-Path -Path $PSScriptRoot -ChildPath 'ppsmart-posh.db' -ErrorAction Stop
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=$db"
        $con.Open()

        $sql = $con.CreateCommand()
        $sql.CommandText = $Query

        $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    
        $DataSet = New-Object System.Data.DataSet

        [void]$adapter.Fill($DataSet)
    }
    
    catch
    
    {
        if($sql -ne $null) { $sql.Dispose() }
        if ($con.State -eq 'Open') { $con.Close() }

        return $DataSet
    }


    return $DataSet
}
