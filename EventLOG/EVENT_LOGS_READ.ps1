Get-Eventlog Security -InstanceId 4624 -ErrorAction Ignore |
   Select TimeGenerated,ReplacementStrings |
   % {
     New-Object PSObject -Property @{
     UserName = $_.ReplacementStrings[5]
     IPAddress = $_.ReplacementStrings[18]
     Workstation = $_.ReplacementStrings[11]
     AuthenticationName = $_.ReplacementStrings[10]
     ID = $_.ReplacementStrings[3]
     Date = $_.TimeGenerated
    }
   } | ?{$_.UserName -notlike "*PSO-EX*"} | Out-GridView