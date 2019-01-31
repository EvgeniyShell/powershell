function remove-groupsinactive ([bool]$mdaemon = $false,[bool]$show = $true){

if ($mdaemon -eq $True)
{
    $users = Get-ADUser -SearchBase "OU=Inactive,DC=ASO,DC=RT,DC=LOCAL" -Properties memberof -filter * |?{!($_.memberOf[0] -eq $null)} | select name,memberof,SamAccountName | ?{$_.memberof -like "*mdaemon*"}
}
if ($mdaemon -eq $false)
{
    $users = Get-ADUser -SearchBase "OU=Inactive,DC=ASO,DC=RT,DC=LOCAL" -Properties memberof -filter * |?{!($_.memberOf[0] -eq $null)} | select name,memberof,SamAccountName | ?{!($_.memberof -like "*mdaemon*")}
}


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
                Remove-ADGroupMember -Identity $user.memberOf[$i] -Members $user.SamAccountName -Confirm:$false
                write-host -BackgroundColor Yellow -ForegroundColor Black $user.memberOf[$i] -> $user.SamAccountName -> удален
            }
            catch
            {
                $error[0].Message
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