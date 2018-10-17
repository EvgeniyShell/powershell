[xml]$xaml = @"
<Window x:Name="Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        Title="Добавление пользователей v.1.0" Height="800" Width="1500" ShowInTaskbar="True" WindowStartupLocation="CenterScreen" MinHeight="800" MinWidth="1400">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <TabControl Margin="10,32.304,10,29.499">
            <TabItem x:Name="Tab_singleUser" Header="Добавление пользователя">
                <Grid Margin="0,0,2.001,0.237">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button Content="Button" HorizontalAlignment="Left" Height="100" Margin="0,0,0,-100" VerticalAlignment="Top" Width="0"/>
                    <Button x:Name="btn_clearsingle" Content="Очистить списки" HorizontalAlignment="Left" Margin="11.333,359.46,0,0" Width="232.5" Height="30.46" VerticalAlignment="Top"/>
                    <ListView x:Name="lstv_SingleUser" Margin="10,10,8.5,0" Height="295" VerticalAlignment="Top" AlternationCount="2">
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <ContextMenu.ContextMenu>
                                    <ContextMenu/>
                                </ContextMenu.ContextMenu>
                                <MenuItem x:Name="lstv_Menu_paste" Header="Вставить"/>
                                <MenuItem x:Name="lstv_Menu_check_ad" Header="Проверить в АД"/>
                                <MenuItem x:Name="lstv_Menu_add_ad" Header="Добавить в АД"/>
                                <MenuItem x:Name="lstv_menu_changeAD" Header="Изменить данные в АД" Background="#FF8ABF39"/>
                                <MenuItem x:Name="lstv_menu_moveOU" Header="Перенести пользователя в OU"/>
                                <MenuItem x:Name="lstv_menu_change" Header="Изменить"/>
                                <MenuItem x:Name="lstv_menu_changePASS" Header="Изменить ПАРОЛЬ"/>
                                <MenuItem x:Name="lstv_Menu_Pass" Header="Новый пароль"/>
                                <MenuItem x:Name="lstv_Menu_status" Header="Сменить статус на ОБЩАЯ УЗ"/>
                                <MenuItem x:Name="lstv_Menu_copy2bufer" Header="Копировать в буфер"/>
                                <MenuItem x:Name="lstv_Menu_Clear" Header="Очистить"/>
                                <MenuItem x:Name="lstv_Menu_delete" Header="Удалить"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="firstname" DisplayMemberBinding="{Binding firstname}"/>
                                <GridViewColumn Header="lastname" DisplayMemberBinding="{Binding lastname}"/>
                                <GridViewColumn Header="displayname" DisplayMemberBinding="{Binding displayname}"/>
                                <GridViewColumn Header="office" DisplayMemberBinding="{Binding office}"/>
                                <GridViewColumn Header="OfficePhone" DisplayMemberBinding="{Binding OfficePhone}"/>
                                <GridViewColumn Header="jobtitle" DisplayMemberBinding="{Binding jobtitle}"/>
                                <GridViewColumn Header="department" DisplayMemberBinding="{Binding department}"/>
                                <GridViewColumn Header="company" DisplayMemberBinding="{Binding company}"/>
                                <GridViewColumn Header="streetaddress" DisplayMemberBinding="{Binding streetaddress}"/>
                                <GridViewColumn Header="city" DisplayMemberBinding="{Binding city}"/>
                                <GridViewColumn Header="postalcode" DisplayMemberBinding="{Binding postalcode}"/>
                                <GridViewColumn Header="samaccountname" DisplayMemberBinding="{Binding samaccountname}"/>
                                <GridViewColumn Header="pass" DisplayMemberBinding="{Binding pass}"/>
                                <GridViewColumn Header="mail" DisplayMemberBinding="{Binding mail}"/>
                                <GridViewColumn Header="ini" DisplayMemberBinding="{Binding ini}"/>
                                <GridViewColumn Header="adcheck" DisplayMemberBinding="{Binding adcheck}"/>
                                <GridViewColumn Header="obshie" DisplayMemberBinding="{Binding obshie}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Button x:Name="btn_AddSingle" Content="Добавить из буфера в ListView (_V)" Margin="11.333,324,0,0" HorizontalAlignment="Left" Width="232.5" Height="30.46" VerticalAlignment="Top"/>
                    <ListView x:Name="lstv_SingleUser_Exist" Margin="11.333,409.5,7.167,10.5">
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <ContextMenu.ContextMenu>
                                    <ContextMenu/>
                                </ContextMenu.ContextMenu>
                                <MenuItem x:Name="lstvExist_Menu_copy2bufer" Header="Копировать в буфер"/>
                                <MenuItem x:Name="lstvExist_Menu_Disable" Header="Отключить пользователя"/>
                                <MenuItem x:Name="lstvExist_Menu_Clear" Header="Очистить"/>
                                <MenuItem x:Name="lstvExist_Menu_delete" Header="Удалить"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="firstname" DisplayMemberBinding="{Binding firstname}"/>
                                <GridViewColumn Header="lastname" DisplayMemberBinding="{Binding lastname}"/>
                                <GridViewColumn Header="displayname" DisplayMemberBinding="{Binding displayname}"/>
                                <GridViewColumn Header="office" DisplayMemberBinding="{Binding office}"/>
                                <GridViewColumn Header="OfficePhone" DisplayMemberBinding="{Binding OfficePhone}"/>
                                <GridViewColumn Header="jobtitle" DisplayMemberBinding="{Binding jobtitle}"/>
                                <GridViewColumn Header="department" DisplayMemberBinding="{Binding department}"/>
                                <GridViewColumn Header="company" DisplayMemberBinding="{Binding company}"/>
                                <GridViewColumn Header="streetaddress" DisplayMemberBinding="{Binding streetaddress}"/>
                                <GridViewColumn Header="city" DisplayMemberBinding="{Binding city}"/>
                                <GridViewColumn Header="postalcode" DisplayMemberBinding="{Binding postalcode}"/>
                                <GridViewColumn Header="samaccountname" DisplayMemberBinding="{Binding samaccountname}"/>
                                <GridViewColumn Header="mail" DisplayMemberBinding="{Binding mail}"/>
                                <GridViewColumn Header="ini" DisplayMemberBinding="{Binding ini}"/>
                                <GridViewColumn Header="Enabled" DisplayMemberBinding="{Binding enabled}"/>
                                <GridViewColumn Header="OU" DisplayMemberBinding="{Binding OU}"/>
                                <GridViewColumn Header="Description" DisplayMemberBinding="{Binding Description}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <TextBox x:Name="txtbox_mail_single" HorizontalAlignment="Left" Height="30" Margin="860.833,324,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="271" TextAlignment="Center" FontSize="16"/>
                    <Button x:Name="btn_single_add_ad" Content="Добавить всех в АД" HorizontalAlignment="Left" Margin="248.833,359.46,0,0" Width="220" Height="30.46" VerticalAlignment="Top"/>
                    <Button x:Name="Btn_single_CheckAD" Content="Проверить всех пользователей в АД" HorizontalAlignment="Left" Margin="248.833,324,0,0" VerticalAlignment="Top" Width="220" Height="30.46"/>
                    <Label Content="Домен для почты:&#xA;[ОТКЛ]Description:" HorizontalAlignment="Left" Margin="750.373,317.08,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.523,0.28"/>
                    <TextBlock x:Name="Statustext" HorizontalAlignment="Left" Margin="1150.5,359.46,0,0" TextWrapping="Wrap" Text="Статус в АД:" VerticalAlignment="Top" RenderTransformOrigin="0.24,0.438" Height="31" Width="271" FontSize="20"/>
                    <Button x:Name="status_canc" Content="Сбросить статусы" Margin="1150.5,324,0,0" Height="30" VerticalAlignment="Top" HorizontalAlignment="Left" Width="124.5"/>
                    <Button x:Name="OU" Content="OU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="37.703" Margin="818.13,359.46,0,0" Height="29.54"/>
                    <TextBox x:Name="txtbox_OU_path" HorizontalAlignment="Left" Height="30" Margin="860.833,359,0,0" VerticalAlignment="Top" Width="271" MaxLines="2" IsUndoEnabled="False"/>
                    <Label Content="Путь к OU:" HorizontalAlignment="Left" Margin="753.833,361.71,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_change_OU" Content="Перенести всех пользователей в OU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="213" Margin="473.833,324,0,0" Height="30"/>
                    <Button x:Name="btn_genpass" Content="gen.pass" HorizontalAlignment="Left" VerticalAlignment="Top" Width="70.998" Margin="473.833,359,0,0" Height="30.92"/>
                    <TextBox x:Name="textbox_pass" HorizontalAlignment="Left" Height="28.21" TextWrapping="Wrap" VerticalAlignment="Top" Width="137.002" Margin="549.831,359.46,0,0" FontSize="16" TextAlignment="Center"/>
                    <Button x:Name="btn_clearOU" Content="Сбросить" HorizontalAlignment="Left" VerticalAlignment="Top" Width="37.703" Margin="818.13,389,0,0" Height="15.5" FontSize="7"/>
                    <RadioButton x:Name="pwd_old" Content="old" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="691.833,374.04,0,0" IsChecked="True"/>
                    <RadioButton x:Name="pwd_new" Content="new" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="691.833,358.55,0,0"/>


                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_Groups" Header="Группы">
                <Grid Margin="0,0,-6.5,-1.263">
                    <ListView x:Name="groups_lstv_user1" Margin="25.5,168,0,0" HorizontalAlignment="Left" Width="253.5" Height="383" VerticalAlignment="Top">
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <ContextMenu.ContextMenu>
                                    <ContextMenu/>
                                </ContextMenu.ContextMenu>
                                <MenuItem x:Name="group_right" Header="Вправо"/>
                                <MenuItem x:Name="group_delete_lstv1" Header="Удалить"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="samAccountname"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <ListView x:Name="groups_lstv_user2" Margin="427.5,168,0,0" HorizontalAlignment="Left" Width="251.5" Height="383" VerticalAlignment="Top">
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <ContextMenu.ContextMenu>
                                    <ContextMenu/>
                                </ContextMenu.ContextMenu>
                                <MenuItem x:Name="group_left" Header="Влево"/>
                                <MenuItem x:Name="group_delete_lstv2" Header="Удалить"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="samAccountname"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Button x:Name="btn_groups_load" Content="Загрузка данных" Margin="25.5,556,0,0" HorizontalAlignment="Left" Width="253.5" Height="100" VerticalAlignment="Top"/>
                    <Button x:Name="btn_groups_save" Content="Добавление пользователя в группы" Margin="427.5,556,0,0" HorizontalAlignment="Left" Width="253.5" Height="100" VerticalAlignment="Top"/>
                    <TextBox x:Name="textbox_Group_newuser" TextWrapping="Wrap" Margin="25.5,42,0,0" HorizontalAlignment="Left" Width="253.5" Height="23" VerticalAlignment="Top"/>
                    <TextBox x:Name="textbox_Group_existuser" TextWrapping="Wrap" Margin="25.5,105,0,0" HorizontalAlignment="Left" Width="253.5" Height="23" VerticalAlignment="Top"/>
                    <Label x:Name="group_label_user1" Content="Новый пользователь" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25.5,11.04,0,0"/>
                    <Label x:Name="group_label_user2" Content="Существующий пользователь" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25.5,74.04,0,0"/>
                    <Label x:Name="group_label_lstv1" Content="Группы существующего пользователя" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25.5,142.04,0,0" Width="253.5"/>
                    <Label x:Name="group_label_lstv2" Content="Группы для нового пользователя" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="427.5,142.04,0,0" Width="251.5"/>
                    <Button x:Name="group_btn_right" Content="&gt;" HorizontalAlignment="Left" Margin="284,297.02,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_left" Content="&lt;" HorizontalAlignment="Left" Margin="284,343.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_deleteall" Content="Очистить" HorizontalAlignment="Left" Margin="284,388.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_moveall" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="284,250.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_moveall_left" Content="&lt;&lt;&lt;" HorizontalAlignment="Left" Margin="284,205.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <GroupBox Header="Копирование групп" Margin="340.5,15.04,0,0" HorizontalAlignment="Left" Width="211" Height="112.96" VerticalAlignment="Top">
                        <Grid HorizontalAlignment="Left" Height="90.667" Margin="0,0,-2,-0.667" VerticalAlignment="Top" Width="201">
                            <RadioButton x:Name="group_radio_one" Content="С источника на одного" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,10,0,0" IsChecked="True"/>
                            <RadioButton x:Name="group_radio_many" Content="С источника для всех в OU" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,29.96,0,0"/>
                            <RadioButton x:Name="group_radio_many2" Content="С источника для нескольких" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,49.92,0,0"/>
                        </Grid>
                    </GroupBox>
                    <Button x:Name="btn_group_OU" Content="OU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="40.37" Height="23" Margin="284,105,0,0"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Menu Height="27.304" VerticalAlignment="Top">
            <MenuItem Header="Файл" Height="27.304" Width="43.135" Background="{x:Null}" RenderTransformOrigin="0.5,0.5">
                <MenuItem x:Name="Btn_LoadFiles" Header="Загрузить данные из файла Users.csv" Height="23" Margin="0,0,0.998,0"/>
                <MenuItem x:Name="btn_export" Header="Выгрузить данные из списка в файл"/>
                <MenuItem x:Name="Btn_Exit" Header="Выход" Height="23" Margin="0,0,0.998,0"/>
            </MenuItem>
            <MenuItem Header="Обработка">
                <MenuItem x:Name="btn_GenPasswords" Header="Генерировать всем новые пароли"/>
                <MenuItem x:Name="btn_allobshie" Header="Пометить все записи как ОБЩИЕ УЗ"/>
                <MenuItem x:Name="btn_exporttoconsole" Header="Выгрузить в Console список"/>
            </MenuItem>
            <MenuItem Header="Настройки">
                <MenuItem Header="Изменение данных">
                    <CheckBox x:Name="chk_all" Content="Убрать всем" IsChecked="True"/>
                    <Label Content="-------------------" IsEnabled="False"/>
                    <CheckBox x:Name="chk_office" Content="office" IsChecked="True"/>
                    <CheckBox x:Name="chk_OfficePhone" Content="OfficePhone" IsChecked="True"/>
                    <CheckBox x:Name="chk_jobtitle" Content="jobtitle" IsChecked="True"/>
                    <CheckBox x:Name="chk_department" Content="department" IsChecked="True"/>
                    <CheckBox x:Name="chk_company" Content="company" IsChecked="True"/>
                    <CheckBox x:Name="chk_streetaddress" Content="streetaddress" IsChecked="True"/>
                    <CheckBox x:Name="chk_city" Content="city" IsChecked="True"/>
                    <CheckBox x:Name="chk_postalcode" Content="postalcode" IsChecked="True"/>
                </MenuItem>
                <CheckBox x:Name="chk_hideshow" Content="Спрятать консоль" IsChecked="True"/>
            </MenuItem>
        </Menu>
        <ProgressBar x:Name="prBar" Margin="10,0,10,10" Background="#FFE6E6E6" RenderTransformOrigin="0,0" Height="14.499" VerticalAlignment="Bottom"/>
        <Label x:Name="label_prBar" Content="" Margin="10,0" FontSize="10" Height="29.499" VerticalAlignment="Bottom"/>

    </Grid>
</Window>


"@

$XReader=(New-Object System.Xml.XmlNodeReader $xaml)
$Form_Main=[Windows.Markup.XamlReader]::Load( $XReader )
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{Set-Variable -Name "xMF_$($_.Name)" -Value $Form_Main.FindName($_.Name)}



[xml]$xaml2 = @"
<Window x:Name="Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        Title="Изменение данных пользователя" Height="458.98" Width="709.339" WindowStartupLocation="CenterScreen" MinWidth="496.339" MinHeight="270.48" ResizeMode="NoResize" BorderThickness="0" WindowStyle="None">
    <Grid Background="#FFD6D6D6" Margin="0">
        <GroupBox Header="Изменения пользователя" HorizontalAlignment="Left" Height="448.98" Margin="10,0,0,0" VerticalAlignment="Top" Width="689.339" BorderBrush="#FF110B91">
            
        </GroupBox>

        <TextBox x:Name="textbox_firstname" HorizontalAlignment="Left" Height="23" Margin="111,23.244,0,0" TextWrapping="Wrap" Text="Firstname" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_lastname" HorizontalAlignment="Left" Height="23" Margin="111,51.244,0,0" TextWrapping="Wrap" Text="Lastname" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_Displayname" HorizontalAlignment="Left" Height="23" Margin="111,79.244,0,0" TextWrapping="Wrap" Text="Displayname" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_Office" HorizontalAlignment="Left" Height="23" Margin="111,107.244,0,0" TextWrapping="Wrap" Text="Office" VerticalAlignment="Top" Width="75.499"/>
        <TextBox x:Name="textbox_OfficePhone" HorizontalAlignment="Left" Height="23" Margin="111,135.244,0,0" TextWrapping="Wrap" Text="OfficePhone" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_Jobtitle" HorizontalAlignment="Left" Height="23" Margin="111,163.244,0,0" TextWrapping="Wrap" Text="Jobtitle" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_Department" HorizontalAlignment="Left" Height="23" Margin="111,191.244,0,0" TextWrapping="Wrap" Text="Department" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_Company" HorizontalAlignment="Left" Margin="111,219.244,0,0" TextWrapping="Wrap" Text="Company" Width="350.499" Height="21.736" VerticalAlignment="Top"/>
        <TextBox x:Name="textbox_StreetAddress" HorizontalAlignment="Left" Height="23" Margin="111,245.98,0,0" TextWrapping="Wrap" Text="StreetAddress" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_City" HorizontalAlignment="Left" Height="23" Margin="111,273.98,0,0" TextWrapping="Wrap" Text="City" VerticalAlignment="Top" Width="350.499"/>
        <TextBox x:Name="textbox_Postalcode" HorizontalAlignment="Left" Height="23" Margin="111,301.98,0,0" TextWrapping="Wrap" Text="PostalCode" VerticalAlignment="Top" Width="350.499" RenderTransformOrigin="0.501,-0.425"/>
        <TextBox x:Name="textbox_Samaccountname" HorizontalAlignment="Left" Height="23" Margin="111,329.98,0,0" TextWrapping="Wrap" Text="Samaccountname" VerticalAlignment="Top" Width="350.499" RenderTransformOrigin="0.501,-0.425"/>
        <TextBox x:Name="textbox_Pass" HorizontalAlignment="Left" Height="23" Margin="111,357.98,0,0" TextWrapping="Wrap" Text="Pass" VerticalAlignment="Top" Width="174.499" RenderTransformOrigin="0.501,-0.425"/>
        <TextBox x:Name="textbox_Mail" HorizontalAlignment="Left" Height="23" Margin="111,385.98,0,0" TextWrapping="Wrap" Text="Mail" VerticalAlignment="Top" Width="350.499" RenderTransformOrigin="0.501,-0.425"/>
        <TextBox x:Name="textbox_ini" HorizontalAlignment="Left" Height="23" Margin="111,413.98,0,0" TextWrapping="Wrap" Text="ini" VerticalAlignment="Top" Width="350.499" RenderTransformOrigin="0.501,-0.425"/>
        <Label Content="Имя" HorizontalAlignment="Left" Margin="29,17.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Фамилия" HorizontalAlignment="Left" Margin="29,45.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="ФИО" HorizontalAlignment="Left" Margin="29,73.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Офис" HorizontalAlignment="Left" Margin="29,101.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Телефон" HorizontalAlignment="Left" Margin="29,129.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Должность" HorizontalAlignment="Left" Margin="29,157.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Отдел" HorizontalAlignment="Left" Margin="29,185.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Компания" HorizontalAlignment="Left" Margin="29,213.087,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Улица" HorizontalAlignment="Left" Margin="29,239.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Город" HorizontalAlignment="Left" Margin="29,267.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Индекс" HorizontalAlignment="Left" Margin="29.08,295.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Логин" HorizontalAlignment="Left" Margin="29.08,323.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Пароль" HorizontalAlignment="Left" Margin="29.08,351.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Почта" HorizontalAlignment="Left" Margin="29.08,379.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Label Content="Инициалы" HorizontalAlignment="Left" Margin="29.08,407.823,0,0" VerticalAlignment="Top" Height="34.244" RenderTransformOrigin="0.589,0.468" FontSize="14"/>
        <Button x:Name="Saveandclose" Content="Сохранить и закрыть" Margin="493,333.432,0,0" Height="30.157" VerticalAlignment="Top" HorizontalAlignment="Left" Width="188"/>
        <Button x:Name="Cancel" Content="Отмена" Margin="493,367.589,0,0" Height="28" VerticalAlignment="Top" HorizontalAlignment="Left" Width="188"/>
        <Button x:Name="Save" Content="Применить" Margin="493,400.589,0,0" Height="27.96" VerticalAlignment="Top" HorizontalAlignment="Left" Width="188" Visibility="Hidden"/>
        <Button x:Name="btn_PassGen" Content="Сгенерировать пароль" HorizontalAlignment="Left" Margin="290.499,357.98,0,0" VerticalAlignment="Top" Width="171" Height="23"/>
    </Grid>
</Window>

"@


$XReader2=(New-Object System.Xml.XmlNodeReader $xaml2)
$Form_2=[Windows.Markup.XamlReader]::Load( $XReader2 )
$xaml2.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{Set-Variable -Name "xMF2_$($_.Name)" -Value $Form_2.FindName($_.Name)}

