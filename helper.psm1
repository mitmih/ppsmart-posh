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
}  # словарь атрибутов


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


function Get-RawValues ([string] $wmi) {
<#
.SYNOPSIS
    функция переводит сырой массив S.M.A.R.T.-данных диска в формат 'код атрибута' = 'raw-значение'

.DESCRIPTION
    функция парсит массивы, конвертирует данные в зависимости от атрибута и возвращает набор атрибутов и их значений в виде хэш-таблицы

.INPUTS
    512-байт массив с данными из класса MSStorageDriver_FailurePredictData

.OUTPUTS
    хэш-таблица @{'код атрибута' = 'raw-значение'}

.PARAMETER data
    512-байт массив с данными из класса MSStorageDriver_FailurePredictData

.EXAMPLE
    $raws = Get-RawValues -data $WMIData

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

    $dctRes = @{}

    [int[]] $data = $wmi.Split(' ')

    for ($i = 2; $i -le 350; $i += 12)

    {
        if ($data[$i] -eq 0) {continue}  # пропускаем атрибуты с нулевым кодом. Thresholds при этом в расчёт не берём, т.к. у каких-то новых моделей ЖД в этом массиве почти одни нули

        $h0  = '{0:x2}' -f $data[$i+0]  # Attribute ID

        $h5  = '{0:x2}' -f $data[$i+5]  # Raw Value 1st byte
        $h6  = '{0:x2}' -f $data[$i+6]  # Raw Value 2nd byte
        $h7  = '{0:x2}' -f $data[$i+7]  # Raw Value 3rd byte
        $h8  = '{0:x2}' -f $data[$i+8]  # Raw Value 4th byte
        $h9  = '{0:x2}' -f $data[$i+9]  # Raw Value 5th byte
        $h10 = '{0:x2}' -f $data[$i+10] # Raw Value 6th byte
        $h11 = '{0:x2}' -f $data[$i+11] # Raw Value 7th byte

    
        if ($data[$i] -in @(190, 194))  # температура оставляем как есть

        {
            $raw = $data[$i+5]
        }

        elseif ($data[$i] -in @(9))  # наработка в часах

        {
            $raw = [Convert]::ToInt64("$h6 $h5".Replace(' ',''),16)
        }

        else  # переводим raw из 16-ричной в 10-тичную систему счисления
    
        {
            $raw = [Convert]::ToInt64("$h11 $h10 $h9 $h8 $h7 $h6 $h5".Replace(' ',''),16)
        }


        $dctRes.Add($data[$i+0], $raw)
    }

return $dctRes

}


function Convert-Flags ([int] $flagDec) {
<#
.SYNOPSIS
    функция конвертирует десятичный статус флага в строку, показывающую установленные биты флагов атрибута

.DESCRIPTION
    конвертация происходит побитовым умножением флага И маски, причём оба числа в десятичной системе
    т.к. на данный момент статус-флаг занимает всего 6 бит (биты с 6го по 15й зарезервированы), функция маскирует только эти биты

.INPUTS
    десятичное значение статус-флага (2й и 3й байты из 12ти-байтового блока S.M.A.R.T.-атрибута)

.OUTPUTS
    строка из шести символов, разделённых пробелами, напр. "- C - P O W", где буква означает что флаг установлен, а прочерк - отсутствие флага
    S - Self-preserving - атрибут может собирать данные, даже если S.M.A.R.T. выключен
    C - Event count - счётчик проишествий
    R - Rate Error - показатель частоты ошибок
    P - Performance - показатель производительности диска
    O - Online / Offline - атрибут обновляется и во время работы (он-лайн) и в простоях (офф-лайн)
    W - Warranty / Prefailure - когда значение атрибута достигает порогового, сбой ожидается в ближайшие 24 часа и диск может быть заменён в течении гарантийного срока

.PARAMETER flagDec
    десятичное значение статус-флага

.EXAMPLE
    $MyStringFlag = Convert-Flags -flagDec 255

.LINK
    структура 512-байт массива данных S.M.A.R.T. (корректное описание)
        http://www.t13.org/Documents/UploadedDocuments/docs2005/e05171r0-ACS-SMARTAttributes_Overview.pdf
    структура и флаги
        https://www.micron.com/~/media/documents/products/technical-note/solid-state-storage/tnfd21_m500-mu02_smart_attributes.pdf
        https://www.micron.com/~/media/documents/products/technical-note/solid-state-storage/tnfd22_client_ssd_smart_attributes.pdf
        https://www.micron.com/~/media/documents/products/technical-note/solid-state-storage/tnfd34_5100_ssd_smart_implementation.pdf

    ещё немного про флаги
        https://white55.ru/smart.html
        https://www.hdsentinel.com/smart/index.php
        http://sysadm.pp.ua/linux/monitoring-systems/smart-attributes.html

.NOTES
    Author: Dmitry Mikhaylov

    Status flag биты:
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

        Bits 6–15 – Reserved
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
    функция конвертирует серийный номер жёсткого диска из шестнадцатиричного в текстовый формат

.DESCRIPTION
    ОС Windows 7 хранит в WMI серийный номер в шестнадцатиричном формате, причем байты по-парно
    поменяны местами, а ОС Windows 10 - в текстовом, совпадающем с номером на наклейке ЖД

    функция проверяет формат серийного номера
    при необходимости конвертирует шестнадцатиричный серийный номер в текст, исправляя порядок символов

    например: win7 format                        -> win10 format
        56394b4d37345130202020202020202020202020 -> 9VMK470Q
        32535841394a5a46354138333932202020202020 -> S2AXJ9FZA53829
        5639514d42363439202020202020202020202020 -> 9VMQ6B94


.INPUTS
    строка с серийным номером ЖД из WMI

.OUTPUTS
    нормализованная строка строка с серийным номером

.PARAMETER wmisn
    строка с серийным номером из WMI

.EXAMPLE
    $TXTSerialNumber = Convert-hex2txt -wmisn $WMISerialNumber

.LINK
    https://blogs.technet.microsoft.com/heyscriptingguy/2011/09/09/convert-hexadecimal-to-ascii-using-powershell/

.NOTES
    Author: Dmitry Mikhaylov

#>


    if ($wmisn.Length -eq 40) {  # проверка на длину
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
    возвращает данные из таблицы базы данных в виде словаря

.DESCRIPTION
    поддерживает SQL-запросы к таблицам Disk, Host
    формирует словарь в виде
        'HostName' = ID
        'SerialNumber' = ID

.INPUTS
    имя таблицы

.OUTPUTS
    объект типа HashTable

.PARAMETER table
    имя таблицы для формирования словаря

.EXAMPLE
    $dctDisk = Get-DBHashTable -table 'Disk'

.LINK

.NOTES
    Author: Dmitry Mikhaylov
#>

$dctQuery = @{  # использовать "... AS key FROM ..." потому что затем так будет заполняться словарь {$dct[$r.key] = $r.ID}
    'Disk' = 'SELECT ID, SerialNumber AS key FROM Disk;'
    'Host' = 'SELECT ID, HostName AS key FROM Host;'
    'Scan' = "SELECT Scan.DiskID || ' ' || Scan.HostID || ' ' || Scan.ScanDate As key, Scan.ID FROM Scan;"
}

try
{

    # проверка битности среды выполнения для подключения подходящей библиотеки
    if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x64\System.Data.SQLite.dll' -ErrorAction Stop}  # 64-bit
    if ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x32\System.Data.SQLite.dll' -ErrorAction Stop}  # 32-bit
    
    # подключение библиотеки для работы с sqlite
    Add-Type -Path $sqlite -ErrorAction Stop

    # открытие соединения с БД
    $db = Join-Path -Path $PSScriptRoot -ChildPath 'ppsmart-posh.db' -ErrorAction Stop
    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = "Data Source=$db"
    $con.Open()

    # исполнение SQL-запроса, результат сохраняется в переменную
    $sql = $con.CreateCommand()

    $sql.CommandText = $dctQuery[$table]


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
foreach ($r in $data.Tables.Rows) {$dct[$r.key] = $r.ID}

return $dct
}


function Update-DB (
    [parameter(Mandatory=$true, ValueFromPipeline=$false)] [string]$tact = $null,
    [parameter(Mandatory=$true, ValueFromPipeline=$false)]         $obj  = $null,
    [parameter(Mandatory=$false, ValueFromPipeline=$false)]   [int]$id   = $null)
{
<#
.SYNOPSIS
    добавляет/обновляет запись в таблицу БД

.DESCRIPTION
    при добавлении возвращает primary key новой записи
    при обновлении существующей записи возвращает $null

.INPUTS
    объект с результатами сканирования жёсткого диска

.OUTPUTS
    ID новой записи или $null

.PARAMETER obj
    объект, содержащий Properties для SQL-запроса в БД

.PARAMETER tact
    table action - название SQL-запроса, возможные запросы:
        NewHost / UpdHost
        NewDisk / UpdDisk
        NewScan / UpdScan

.PARAMETER id
    primary key записи, которую нужно обновить

.EXAMPLE
    $hID = Update-DB -tact NewHost -obj ($Scan | Select-Object -Property 'HostName')
        добавить новую запись

.EXAMPLE
    $null= Update-DB -tact UpdHost -obj ($Scan | Select-Object -Property 'HostName') -id $HostID[$Scan.HostName]
        обновить запись по её primary key

.LINK

.NOTES
    Author: Dmitry Mikhaylov
#>

    $newrecID = $null

    $dctQuery = @{
        'NewHost' = 'INSERT OR IGNORE INTO `Host` (`HostName`) VALUES (@HostName);'

        'UpdHost' = ''  # not in use 'UPDATE `Host` SET `ScanDate` = @ScanDate, `Ping` = @Ping, `WMIInfo` = @WMIInfo WHERE ID = @ID;'

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
        # проверка битности среды выполнения для подключения подходящей библиотеки
        if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x64\System.Data.SQLite.dll' -ErrorAction Stop}  # 64-bit
        if ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x32\System.Data.SQLite.dll' -ErrorAction Stop}  # 32-bit
    
        # подключение библиотеки для работы с sqlite
        Add-Type -Path $sqlite -ErrorAction Stop

        # открытие соединения с БД
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


    if ($id)  # update record (передан primary key)
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
    возвращает результат SQL-запроса

.DESCRIPTION
    возвращает результат SQL-запроса

.INPUTS
    строка с SQL-запросом, sqlite база данных

.OUTPUTS
    набор записей

.PARAMETER Query
    SQL-запрос

.PARAMETER base
    путь к файлу БД

.EXAMPLE
    $data = Get-DBData -Query 'SELECT * FROM Scan' -base 'ppsmart-posh.db'

.LINK

.NOTES
    Author: Dmitry Mikhaylov
#>

    $DataSet = $null

    try

    {
        # проверка битности среды выполнения для подключения подходящей библиотеки
        if ([IntPtr]::Size -eq 8) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x64\System.Data.SQLite.dll' -ErrorAction Stop}  # 64-bit
        if ([IntPtr]::Size -eq 4) {$sqlite = Join-Path -Path $PSScriptRoot -ChildPath 'x32\System.Data.SQLite.dll' -ErrorAction Stop}  # 32-bit
    
        # подключение библиотеки для работы с sqlite
        Add-Type -Path $sqlite -ErrorAction Stop

        # открытие соединения с БД
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
