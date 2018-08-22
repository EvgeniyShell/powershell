Function getgroups
{
    try
    {
        $groups_load_user1 = Get-ADUser $xMF_textbox_Group_newuser.Text -Properties memberof
        $groups_load_user2 = Get-ADUser $xMF_textbox_Group_existuser.Text -Properties memberof
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show($error[0].Exception.Message,"Ошибка","OK","Warning")
        Write-Host -BackgroundColor Yellow "ERROR -> " $error[0].Exception.Message
        return "ERROR"
    }
            if ($groups_load_user2.MemberOf[0] -eq $null)
            {
                Write-Host "У пользователя "$xMF_textbox_Group_existuser.Text" нет групп"
                return "0"
            }
            else
            {
                $groups_load_compare = Compare-Object -ReferenceObject $groups_load_user1.MemberOf -DifferenceObject $groups_load_user2.MemberOf -IncludeEqual
                $global:er = 0
                foreach ($strings in $groups_load_compare)
                {
                    try
                    {
                        $strings2 = Get-ADGroup -Identity $strings.InputObject
                        if ($strings.SideIndicator -eq "=>")
                        {
                            write-host -BackgroundColor red -ForegroundColor white У пользователя $groups_load_user1.Name нет групп $strings2.SamAccountName
                            $xMF_groups_lstv_user1.Items.Add($strings2.SamAccountName)
                        }
                    }
                    catch
                    {
                        $global:er++
                        Write-Host "ERROR -> " $error[0].Exception.Message
                    }
                }
            }
}