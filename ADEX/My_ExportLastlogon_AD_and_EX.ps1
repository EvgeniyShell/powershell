Function SizeInBytes ($itemSizeString)
{
    $posOpenParen = $itemSizeString.IndexOf("(") + 1
    $numCharsInSize = $itemSizeString.IndexOf(" bytes") - $posOpenParen 
    $SizeInBytes = $itemSizeString.SubString($posOpenParen,$numCharsInSize).Replace(",","")
	return $SizeInBytes 
}


function Show_Users_ADEX($DOU="OU=Users,OU=SOCI,OU=OU_GBU_GKU_OBU,DC=ASO,DC=RT,DC=LOCAL",$lastlogon=0) {
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
                'CanonicalName' = "";
                
            }
$date = (Get-Date).AddMonths($lastlogon)
$test = Get-ADUser -SearchBase $DOU -Filter {LastLogonDate -le $date} -Properties mail,lastLogonTimestamp,LastLogonDate,DisplayName,Description,whenCreated,CanonicalName -ErrorAction SilentlyContinue
$i=0

if ($test) {
$test1 = ($test | Measure-Object).Count

Write-Host -BackgroundColor Yellow -ForegroundColor Black "Путь до OU: $DOU"
Write-Host -BackgroundColor Yellow -ForegroundColor Black "Количество пользователей в OU: $test1"
foreach ($tests in $test) 
    {
    $lastlogon = [DateTime]::FromFileTime($tests.lastLogonTimestamp)
    $err = Get-MailboxStatistics $tests.SamAccountName -ErrorAction SilentlyContinue | select DisplayName,Database,LastLogonTime,DatabaseProhibitSendQuota,TotalItemSize
        if ($err)
        {
            $err2 = get-mailbox $tests.SamAccountName -ErrorAction SilentlyContinue | select ProhibitSendQuota
            $GB = SizeInBytes([string]$err.DatabaseProhibitSendQuota.Value)
            $GB2 = SizeInBytes([string]$err.TotalItemSize.value)

            if ($err2.ProhibitSendQuota -eq "Unlimited") {$out = "Unlimited"} 
            else{$GB3 = SizeInBytes([string]$err2.ProhibitSendQuota)
            $out = "{0:N2}" -f $($GB3/1GB) +" GB"}
            

            $i++
            $spisok += [PSCustomObject]@{
                'DisplayName' = $err.DisplayName;
                'SamAccountname' = $tests.SamAccountName;
                'Database' = $err.Database;
                'LastLogonAD' = $lastlogon
                'LastLogonEX' = $err.LastLogonTime;
                'QuotaMBX' = "{0:N2}" -f $($GB/1GB) +" GB";
                'QuataMail' = $out;
                'TotalItemSize' = "{0:N2}" -f $($GB2/1GB) +" GB";
                'Description' = $tests.Description;
                'Enabled' = $tests.enabled;
                'WhenCreated' = $tests.whenCreated;
                'CanonicalName' = $tests.CanonicalName;
                
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
                'Quota' = "";
                'TotalItemSize' = "";
                'Description' = $tests.Description;
                'Enabled' = $tests.enabled;
                'WhenCreated' = $tests.whenCreated;
                'CanonicalName' = $tests.CanonicalName;

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
