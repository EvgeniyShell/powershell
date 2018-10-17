Function SizeInBytes ($itemSizeString)
{
    $posOpenParen = $itemSizeString.IndexOf("(") + 1
    $numCharsInSize = $itemSizeString.IndexOf(" bytes") - $posOpenParen 
    $SizeInBytes = $itemSizeString.SubString($posOpenParen,$numCharsInSize).Replace(",","")
	return $SizeInBytes 
}


function Show_Users_ADEX($DOU="OU=Users,OU=SOCI,OU=OU_GBU_GKU_OBU,DC=ASO,DC=RT,DC=LOCAL") {
if ($DOU -eq $null) {write-host "Введите путь до OU"}
else
{
$spisok=@()
            $spisok += [PSCustomObject]@{
                'DisplayName' = "";
                'SamAccountname' = "";
                'Database' = "";
                'LastLogonAD' = "";
                'LastLogonEX' = "";
                'QuotaMBX' = "";
                'QuataMail' = "";
                'TotalItemSize' = "";
                'Description' = "";
                'Enabled' = "";
                'WhenCreated' = "";
                'WhenMailboxCreated' = "";
                'CanonicalName' = "";
                'CanonicalName2' = "";
            }
$test = Get-ADUser -SearchBase $DOU -Filter * -Properties mail,lastLogonTimestamp,LastLogonDate,DisplayName,Description,whenCreated,CanonicalName -ErrorAction SilentlyContinue
$i=0

if ($test) {
$test1 = ($test | Measure-Object).Count

Write-Host -BackgroundColor Yellow -ForegroundColor Black "Путь до OU: $DOU"
Write-Host -BackgroundColor Yellow -ForegroundColor Black "Количество пользователей в OU: $test1"
foreach ($tests in $test) 
    {
    $lastlogon = [DateTime]::FromFileTime($tests.lastLogonTimestamp)
    $err2 = get-mailbox $tests.SamAccountName -ErrorAction SilentlyContinue | select ProhibitSendQuota,WhenMailboxCreated,DisplayName,Database
    $cannonical = $tests.CanonicalName -replace "ASO.RT.LOCAL/" -replace "OU_GBU_GKU_OBU/" -replace "Users/" -replace "Inactive/" -replace "OU_RT/" -replace "/"+$tests.DisplayName
    $cannonical2 = $cannonical.Split("/")
        if ($err2)
        {
            $err = Get-MailboxStatistics $tests.SamAccountName -ErrorAction SilentlyContinue | select LastLogonTime,DatabaseProhibitSendQuota,TotalItemSize
            if ($err){
            $GB = SizeInBytes([string]$err.DatabaseProhibitSendQuota.Value)
            $GB2 = SizeInBytes([string]$err.TotalItemSize.value)


            if ($err2.ProhibitSendQuota -eq "Unlimited") {$out = "Unlimited"} 
            else{$GB3 = SizeInBytes([string]$err2.ProhibitSendQuota)
            $out = "{0:N2}" -f $($GB3/1GB) +" GB"}     
            }

            $i++
            $spisok += [PSCustomObject]@{
                'DisplayName' = $err2.DisplayName;
                'SamAccountname' = $tests.SamAccountName;
                'Database' = $err2.Database;
                'LastLogonAD' = $lastlogon
                'LastLogonEX' = $err.LastLogonTime;
                'QuotaMBX' = "{0:N2}" -f $($GB/1GB) +" GB";
                'QuataMail' = $out;
                'TotalItemSize' = "{0:N2}" -f $($GB2/1GB) +" GB";
                'Description' = $tests.Description;
                'Enabled' = $tests.enabled;
                'WhenCreated' = $tests.whenCreated;
                'WhenMailboxCreated' = $err2.WhenMailboxCreated;
                'CanonicalName' = $tests.CanonicalName;
                'CanonicalName2' = $cannonical;
                
            }
        }
        else
        {
            $i++
            $spisok += [PSCustomObject]@{
                'DisplayName' = $tests.DisplayName;
                'SamAccountname' = $tests.SamAccountName;
                'Database' = "";
                'LastLogonAD' = $lastlogon;
                'LastLogonEX' = "";
                'QuotaMBX' = "";
                'QuataMail' = "";
                'TotalItemSize' = "";
                'Description' = $tests.Description;
                'Enabled' = $tests.enabled;
                'WhenCreated' = $tests.whenCreated;
                'WhenMailboxCreated' = "";
                'CanonicalName' = $tests.CanonicalName;
                'CanonicalName2' = $cannonical;

            }
        }
        Write-Progress -Activity "Создание списка из $test1 пользователей" -Status "Обработка пользователя: $i" -CurrentOperation $tests.Name -PercentComplete ($i / $test1*100)
             
    }
    return $spisok
}
else
{write "Путь до OU введен неверно"}
}


}
