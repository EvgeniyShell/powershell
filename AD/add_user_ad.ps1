function add_user_ad([bool]$multi = $true)
{
#Проверяем входное значение, если True, то отрабаывает скрипт добавления всех пользователей в списке.
if ($multi -eq $true)
{

        if (!($xmf_lstv_SingleUser.Items.Count -eq 0))
        {
            $temp = ""
            $t=0
            $y=0
            $err_count=0
            $xMF_prBar.Value = 0
            $xMF_prBar.Maximum = $xmf_lstv_SingleUser.Items.Count - 1
            $xMF_label_prBar.Content = ""

            for ($i=0 ; $i -ne $xmf_lstv_SingleUser.Items.Count ; $i++)
            {

                if (($xmf_lstv_SingleUser.Items[$i].adcheck -eq "checked") -or ($xmf_lstv_SingleUser.Items[$i].obshie -eq "yes"))
                {

                    $DisplayName = $xmf_lstv_SingleUser.Items[$i].Displayname
                    $UserFirstname = $xmf_lstv_SingleUser.Items[$i].firstname
                    $UserLastname = $xmf_lstv_SingleUser.Items[$i].lastname
                    $ini = $xmf_lstv_SingleUser.Items[$i].ini
                    $physicalDeliveryOfficeName = $xmf_lstv_SingleUser.Items[$i].office
                    $phone = $xmf_lstv_SingleUser.Items[$i].officePhone
                    $jobtitle = $xmf_lstv_SingleUser.Items[$i].Jobtitle
                    $department = $xmf_lstv_SingleUser.Items[$i].department
                    $ncompany = $xmf_lstv_SingleUser.Items[$i].company
                    $street = $xmf_lstv_SingleUser.Items[$i].streetAddress
                    $city = $xmf_lstv_SingleUser.Items[$i].city
                    $code = $xmf_lstv_SingleUser.Items[$i].postalCode
                    $SAM = $xmf_lstv_SingleUser.Items[$i].samaccountname
                    $LogonName = $SAM + "@SAKHALIN.GOV.RU"
                    $Password = $xmf_lstv_SingleUser.Items[$i].pass
                    $ncountry = "RU"
                    $mail = $xmf_lstv_SingleUser.Items[$i].mail

                    try
                    {
                        New-ADUser -Name $DisplayName -SamAccountName $SAM -Initials $Ini -Office $physicalDeliveryOfficeName -UserPrincipalName $LogonName -DisplayName $DisplayName -GivenName $UserFirstname -surName $UserLastname -OfficePhone $phone -EmailAddress $mail -department $Department -Title $jobtitle -City $city -StreetAddress $street -PostalCode $code -Country $ncountry -Company $ncompany -Enabled $true -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Path $xMF_txtbox_OU_path.Text -PasswordNeverExpires $true -Passthru -ErrorVariable err
                        Write-Host -BackgroundColor Yellow -ForegroundColor black "+ Пользователь -> $DisplayName добавлен в АД"
                        $xMF_label_prBar.Content = "+ Пользователь -> $DisplayName добавлен в АД"
                        $y=$y+1
                        $xMF_prBar.Value += 1
                        #$xMF_prBar.Value = $y + $err_count
                        $Form_Main.Dispatcher.Invoke([action]{},"Render")
                        $xmf_lstv_SingleUser.Items[$i].adcheck = "Added"
                        $xMF_lstv_SingleUser.Items.Refresh()
                        if ($xmf_lstv_SingleUser.Items.Count -eq 1)
                        {
                        $xMF_Statustext.Background = "green"
                        $xMF_Statustext.Foreground = "white"
                        $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$i].adcheck
                        }
                    }
                    catch
                    {
                        if ($err)
                        {
                            $err_count=$err_count+1
                            $xMF_prBar.Value += 1
                            #$xMF_prBar.Value = $err_count + $y
                            $Form_Main.Dispatcher.Invoke([action]{},"Render")
                            $Bufer = $error[0].CategoryInfo.Category
                            $bufer2 = $error[0].FullyQualifiedErrorId
                            $bufer3 = $error[0].Exception.Message
                            Write-Host -BackgroundColor Red -ForegroundColor white "!+ Пользователь -> $DisplayName не добавлен в АД ($Bufer - $bufer2) ($bufer3)"
                            $xMF_label_prBar.Content = "!+ Пользователь -> $DisplayName не добавлен в АД ($Bufer - $bufer2) ($bufer3)"
                            $xMF_Statustext.Background = "red"
                            $xMF_Statustext.Foreground = "black"
                            $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$i].adcheck
                            $xmf_lstv_SingleUser.Items[$i].adcheck = "ERROR"
                            $xMF_lstv_SingleUser.Items.Refresh()
                        }
                    }
                }
                elseif (!($xmf_lstv_SingleUser.Items[$i].adcheck -eq "EXIST"))
                {
                    $t += 1
                    $xMF_prBar.Value += 1
                    $temp += $xmf_lstv_SingleUser.Items[$i].Displayname+" STATUS: "+$xmf_lstv_SingleUser.Items[$i].adcheck+"`n"
                }
                else
                {
                    $t += 1
                    $xMF_prBar.Value += 1
                    $temp += $xmf_lstv_SingleUser.Items[$i].Displayname+" STATUS: "+$xmf_lstv_SingleUser.Items[$i].adcheck+"`n"
                }

            }


            if (!($temp -eq ""))
            {
                if ($t -cgt 10)
                {
                    [System.Windows.Forms.MessageBox]::Show("Всего не добавленых пользователей - $t`nВсего добавлено пользователей - $y","Уведомление","OK","Warning")
                }
                else
                {
                    [System.Windows.Forms.MessageBox]::Show("Не добавленные пользователи *$t*:`n$temp","Уведомление","OK","Warning")
                }
                
            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("Добавлено пользователей: $y`nПользователей с ошибками: $err_count ","Уведомление","OK","Information")
            }

            $temp = ""
        }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
        }
}
else
{

        if (!($xmf_lstv_SingleUser.items.Count -eq 0))
        {

            if ((!($xmf_lstv_SingleUser.SelectedIndex -eq -1) -and ($xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck -eq "checked")) -or (!($xmf_lstv_SingleUser.SelectedIndex -eq -1) -and ($xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].obshie -eq "yes")))
            {
                $DisplayName = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].Displayname
                $UserFirstname = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].firstname
                $UserLastname = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].lastname
                $ini = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].ini
                $physicalDeliveryOfficeName = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].office
                $phone = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].officePhone
                $jobtitle = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].Jobtitle
                $department = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].department
                $ncompany = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].company
                $street = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].streetAddress
                $city = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].city
                $code = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].postalCode
                $SAM = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].samaccountname
                $LogonName = $SAM + "@SAKHALIN.GOV.RU"
                $Password = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].pass
                $ncountry = "RU"
                $mail = $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].mail

                try
                {
                    New-ADUser -Name $DisplayName -SamAccountName $SAM -Initials $Ini -Office $physicalDeliveryOfficeName -UserPrincipalName $LogonName -DisplayName $DisplayName -GivenName $UserFirstname -surName $UserLastname -OfficePhone $phone -EmailAddress $mail -department $Department -Title $jobtitle -City $city -StreetAddress $street -PostalCode $code -Country $ncountry -Company $ncompany -Enabled $true -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Path $xMF_txtbox_OU_path.Text -PasswordNeverExpires $true -Passthru -ErrorVariable err
                    Write-Host -BackgroundColor Yellow -ForegroundColor black "+ Пользователь -> $DisplayName добавлен в АД"
                    $xMF_label_prBar.Content = "+ Пользователь -> $DisplayName добавлен в АД"
                    $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck = "Added"
                    $xMF_lstv_SingleUser.Items.Refresh()
                    $xMF_Statustext.Background = "green"
                    $xMF_Statustext.Foreground = "white"
                    $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
                }
                catch
                {
                    if ($err)
                    {
                        $Bufer = $error[0].CategoryInfo.Category
                        $bufer2 = $error[0].FullyQualifiedErrorId
                        $bufer3 = $error[0].Exception.Message
                        Write-Host -BackgroundColor Red -ForegroundColor white "!+ Пользователь -> $DisplayName не добавлен в АД ($Bufer - $bufer2) ($bufer3)"
                        $xMF_label_prBar.Content = "!+ Пользователь -> $DisplayName не добавлен в АД ($Bufer - $bufer2) ($bufer3)"
                        $xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck = "ERROR"
                        $xMF_lstv_SingleUser.Items.Refresh()
                        $xMF_Statustext.Background = "red"
                        $xMF_Statustext.Foreground = "black"
                        $xMF_Statustext.Text = "Статус в АД: "+$xmf_lstv_SingleUser.Items[$xmf_lstv_SingleUser.SelectedIndex].adcheck
                    }
                }
            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("Данный пользователь не проверен в АД","Уведомление","OK","Warning")
            }
        }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Список пустой, загрузите файл.","Уведомление","OK","Warning")
        }
}
}