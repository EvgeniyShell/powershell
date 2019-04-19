function remove-groupsinactive ([bool]$show = $true,[bool]$clearNULL = $false,$dest = "OU=Inactive,DC=ASO,DC=RT,DC=LOCAL"){

#
# $show - показать сформированый список
#

$users = Get-ADUser -SearchBase $dest -Properties memberof -filter * |?{!($_.memberOf[0] -eq $null)} | select name,memberof,SamAccountName #| ?{$_.memberof -like "*mdaemon*"}


if ($show -eq $true)
{
    $users
}
elseif ($show -eq $false)
{
 

if (!($users.count -eq 0))
{


foreach ($user in $users)
{

    if ($user.memberOf.count -gt 0)
    {
        
        write-host $user.name -> $user.memberOf.count
        
        for ($i=0; $i -lt $user.memberOf.count;$i++)
        {
            try
            {
                $tmp = $user.memberOf[$i]
                if ($clearNULL -eq $false)
                {
                    if (($tmp -like "*mdaemon*") -or ($tmp -like "*NULL-DELIVERY*"))
                    {
                        write-host -BackgroundColor Yellow -ForegroundColor Black Группа - $user.memberOf[$i] -> $user.SamAccountName -> пропущена
                    }
                    else
                    { 
                        Remove-ADGroupMember -Identity $user.memberOf[$i] -Members $user.SamAccountName -Confirm:$false
                        write-host -BackgroundColor Yellow -ForegroundColor Black $user.memberOf[$i] -> $user.SamAccountName -> удален
                    }
                } # ($clearALL -eq $false)


                if ($clearNULL -eq $true)
                {
                    if (($tmp -like "*mdaemon*"))
                    {
                        write-host -BackgroundColor Yellow -ForegroundColor Black Группа - $user.memberOf[$i] -> $user.SamAccountName -> пропущена
                    }
                    else
                    { 
                        Remove-ADGroupMember -Identity $user.memberOf[$i] -Members $user.SamAccountName -Confirm:$false
                        write-host -BackgroundColor Yellow -ForegroundColor Black $user.memberOf[$i] -> $user.SamAccountName -> удален
                    }
                } # ($clearALL -eq $true)

            }
            catch
            {
                $Error[0].Exception.Message
            }
        }

        write-host "---------------------------------------------------------------------------------"

    }

}


}
else
{
    Write-Host Список пустой
}

} # ($show -eq $true)

}