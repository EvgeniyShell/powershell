
function MakeCopy {

if ((Test-Path -Path "$ScriptPath\users_add.csv") -or (Test-Path -Path "$ScriptPath\users_exist.csv"))
        {
            if (!(Test-Path -Path "$ScriptPath\$Global:ArchiveDer\users_add.csv") -or !(Test-Path -Path "$ScriptPath\$Global:ArchiveDer\users_exist.csv"))
            {
                try
                {
                    Move-Item -Path "$ScriptPath\users_add.csv" -Destination $ScriptPath\$Global:ArchiveDer -ErrorAction stop
                    write-host -BackgroundColor Yellow -ForegroundColor Black "Архив -> Файл $ScriptPath\users_add.csv перенесен"
                }
                catch
                {
                    Switch -Wildcard ($error[0].CategoryInfo.Category)
                    {
                        "*WriteError*" {write-host -BackgroundColor Red -ForegroundColor Yellow "Архив -> Файл $ScriptPath\users_add.csv открыт, необходимо его закрыть"
                        [System.Windows.Forms.MessageBox]::Show("Файл $ScriptPath\users_add.csv открыт, необходимо его закрыть","Уведомление","OK","Warning")}
                    }
                }


                try
                {
                    Move-Item -Path "$ScriptPath\users_exist.csv" -Destination $ScriptPath\$Global:ArchiveDer -ErrorAction stop
                    write-host -BackgroundColor Yellow -ForegroundColor Black "Архив -> Файл $ScriptPath\users_exist.csv перенесен"
                }
                catch
                {
                    Switch -Wildcard ($error[0].CategoryInfo.Category)
                    {
                        "*WriteError*" {write-host -BackgroundColor Red -ForegroundColor Yellow "Архив -> Файл $ScriptPath\users_exist.csv открыт, необходимо его закрыть"
                        [System.Windows.Forms.MessageBox]::Show("Файл $ScriptPath\users_exist.csv открыт, необходимо его закрыть","Уведомление","OK","Warning")}
                    }
                }




            }
            else
            {
                write-host -BackgroundColor Red -ForegroundColor Yellow "Архив -> В папке $ScriptPath\$Global:ArchiveDer уже имеется файлы, скопировать копию?"
                $message = [System.Windows.Forms.MessageBox]::Show("В папке $ScriptPath\$Global:ArchiveDer уже имеется файлы, скопировать копию?","Уведомление","OKCancel","Warning")
                if ($message -eq "OK")
                {
                    try
                    {
                        $i = 2
                        while (Test-Path -Path "$ScriptPath\$Global:ArchiveDer\users_add_$i.csv")
                        {
                        $i = $i + 1
                        }
                        Move-Item -Path "$ScriptPath\users_add.csv" -Destination "$ScriptPath\$Global:ArchiveDer\users_add_$i.csv" -ErrorAction stop
                        write-host -BackgroundColor Yellow -ForegroundColor Black "Архив -> Файл $ScriptPath\users_add.csv перенесен и назван users_add_$i.csv"
                    }
                    catch
                    {
                    Switch -Wildcard ($error[0].CategoryInfo.Category)
                        {
                        "*WriteError*" {write-host -BackgroundColor Red -ForegroundColor Yellow "Архив -> Файл $ScriptPath\users_add.csv открыт, необходимо его закрыть"
                        [System.Windows.Forms.MessageBox]::Show("Файл $ScriptPath\users_add.csv открыт, необходимо его закрыть","Уведомление","OK","Warning")}
                        }
                    }


                    try
                    {
                        $i = 2
                        while (Test-Path -Path "$ScriptPath\$Global:ArchiveDer\users_exist_$i.csv")
                        {
                        $i = $i + 1
                        }
                        Move-Item -Path "$ScriptPath\users_exist.csv" -Destination "$ScriptPath\$Global:ArchiveDer\users_exist_$i.csv" -ErrorAction stop
                        write-host -BackgroundColor Yellow -ForegroundColor Black "Архив -> Файл $ScriptPath\users_exist.csv перенесен и назван users_exist_$i.csv"
                    }
                    catch
                    {
                    Switch -Wildcard ($error[0].CategoryInfo.Category)
                        {
                        "*WriteError*" {write-host -BackgroundColor Red -ForegroundColor Yellow "Архив -> Файл $ScriptPath\users_exist.csv открыт, необходимо его закрыть"
                        [System.Windows.Forms.MessageBox]::Show("Файл $ScriptPath\users_exist.csv открыт, необходимо его закрыть","Уведомление","OK","Warning")}
                        }
                    }

                }
            }
        }
}




