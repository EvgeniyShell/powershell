$ids = 4740
#$ids = 4740,4726,4720
#$ids = 4726
$newest = 3

[int64]$time = [int64](New-TimeSpan (Get-Date).AddHours(-3)).TotalMilliseconds
$Addomain = Get-ADDomainController -Filter {IsReadOnly -eq $false} | select HostName
$out = @()

foreach ($id in $ids)
{

foreach ($AD in $Addomain)
{
Invoke-Command -ComputerName $AD.HostName {get-eventlog -logname security -newest $using:newest -InstanceId $using:id -ErrorAction SilentlyContinue} -AsJob -JobName EVENT | out-null
}

write-Host Waiting.... $id
Wait-Job -Name EVENT | Out-Null
write-Host Completed.... $id
$events = Get-Job | Receive-Job


if (Get-Job -Name EVENT -HasMoreData $False)
{
Get-Job -Name EVENT | Remove-Job
}



foreach ($event in $events)
{
        switch($Event.EventID)
        {
            "4740" {$action = "УЗ Заблокирована";$comp = $Event.ReplacementStrings[1];$whdel = ""}
            "4720" {$action = "УЗ создана"; $comp = ""; $whdel = $Event.ReplacementStrings[4]}
            "4726" {$action = "УЗ удалена"; $comp = ""; $whdel = $Event.ReplacementStrings[4]}
            "Default" {$action = "Unknown"; $comp = ""; $whdel = ""}
        }

        [array]$out += [PSCustomObject]@{
        'Action' = $action
        'Server' = $Event.PSComputerName -replace ".aso.rt.local",""
        'Time' = $Event.TimeGenerated
        'ID' = [string]$Event.EventID
        'Accountname' = $Event.ReplacementStrings[0]
        'Computername' = $comp
        'Who change/Delete' = $whdel

        }
}
}