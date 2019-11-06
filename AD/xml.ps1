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
                    <Button Content="Button" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="0"/>
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
                                <GridViewColumn Header="obshie" DisplayMemberBinding="{Binding obshie}">

                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <TextBlock x:Name="Txt" Text="{Binding obshie}" />
                                            <DataTemplate.Triggers>
                                                <DataTrigger Binding="{Binding obshie}" Value="no">
                                                    <Setter TargetName="Txt" Property="Foreground" Value="Red" />
                                                </DataTrigger>
                                            </DataTemplate.Triggers>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
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
                                <MenuItem x:Name="lstvExist_Menu_deletemanager" Header="Удалить менеджера у пользователя"/>
                                <MenuItem x:Name="lstvExist_Menu_deletegroup" Header="Удалить группы у пользователя"/>
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
                    <TextBox x:Name="txtbox_mail_single" HorizontalAlignment="Left" Height="30" Margin="861.833,324,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="271" TextAlignment="Center" FontSize="16"/>
                    <Button x:Name="btn_single_add_ad" Content="Добавить всех в АД" HorizontalAlignment="Left" Margin="249.833,359.46,0,0" Width="220" Height="30.46" VerticalAlignment="Top"/>
                    <Button x:Name="Btn_single_CheckAD" Content="Проверить всех пользователей в АД" HorizontalAlignment="Left" Margin="249.833,324,0,0" VerticalAlignment="Top" Width="220" Height="30.46"/>
                    <Label Content="Домен для почты:&#xA;[ОТКЛ]Description:" HorizontalAlignment="Left" Margin="751.373,317.08,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.523,0.28"/>
                    <TextBlock x:Name="Statustext" HorizontalAlignment="Left" Margin="1151.5,359.46,0,0" TextWrapping="Wrap" Text="Статус в АД:" VerticalAlignment="Top" RenderTransformOrigin="0.24,0.438" Height="31" Width="271" FontSize="20"/>
                    <Button x:Name="status_canc" Content="Сбросить статусы" Margin="1151.5,324,0,0" Height="30" VerticalAlignment="Top" HorizontalAlignment="Left" Width="124.5"/>
                    <Button x:Name="OU" Content="OU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="37.703" Margin="819.13,359.46,0,0" Height="29.54"/>
                    <TextBox x:Name="txtbox_OU_path" HorizontalAlignment="Left" Height="30" Margin="861.833,359,0,0" VerticalAlignment="Top" Width="271" MaxLines="2" IsUndoEnabled="False"/>
                    <Label Content="Путь к OU:" HorizontalAlignment="Left" Margin="754.833,361.71,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_change_OU" Content="Перенести всех пользователей в OU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="213" Margin="474.833,324,0,0" Height="30"/>
                    <Button x:Name="btn_genpass" Content="gen.pass" HorizontalAlignment="Left" VerticalAlignment="Top" Width="70.998" Margin="474.833,359,0,0" Height="30.92"/>
                    <TextBox x:Name="textbox_pass" HorizontalAlignment="Left" Height="28.21" TextWrapping="Wrap" VerticalAlignment="Top" Width="137.002" Margin="550.831,359.46,0,0" FontSize="16" TextAlignment="Center"/>
                    <Button x:Name="btn_clearOU" Content="Сбросить" HorizontalAlignment="Left" VerticalAlignment="Top" Width="37.703" Margin="819.13,389,0,0" Height="15.5" FontSize="7"/>
                    <RadioButton x:Name="pwd_old" Content="old" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="692.833,374.04,0,0" IsChecked="True"/>
                    <RadioButton x:Name="pwd_new" Content="new" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="692.833,358.55,0,0"/>


                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_Groups" Header="Группы">
                <Grid Margin="0,0,-6.5,-1.263">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
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
                    <ListView x:Name="groups_lstv_user2" Margin="427.42,168,0,0" HorizontalAlignment="Left" Width="251.5" Height="383" VerticalAlignment="Top">
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
                    <Button x:Name="btn_groups_save" Content="Добавление пользователя в группы" Margin="427.42,556,0,0" HorizontalAlignment="Left" Width="253.5" Height="100" VerticalAlignment="Top"/>
                    <TextBox x:Name="textbox_Group_newuser" TextWrapping="Wrap" Margin="25.5,42,0,0" HorizontalAlignment="Left" Width="253.5" Height="23" VerticalAlignment="Top"/>
                    <TextBox x:Name="textbox_Group_existuser" TextWrapping="Wrap" Margin="25.5,105,0,0" HorizontalAlignment="Left" Width="253.5" Height="23" VerticalAlignment="Top"/>
                    <Label x:Name="group_label_user1" Content="Новый пользователь" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25.5,11.04,0,0"/>
                    <Label x:Name="group_label_user2" Content="Существующий пользователь" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25.5,74.04,0,0"/>
                    <Label x:Name="group_label_lstv1" Content="Группы существующего пользователя" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="25.5,142.04,0,0" Width="253.5"/>
                    <Label x:Name="group_label_lstv2" Content="Группы для нового пользователя" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="427.42,142.04,0,0" Width="251.5"/>
                    <Button x:Name="group_btn_right" Content="&gt;" HorizontalAlignment="Left" Margin="283.92,297.02,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_left" Content="&lt;" HorizontalAlignment="Left" Margin="283.92,343.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_deleteall" Content="Очистить" HorizontalAlignment="Left" Margin="283.92,388.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_moveall" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="283.92,250.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <Button x:Name="group_btn_moveall_left" Content="&lt;&lt;&lt;" HorizontalAlignment="Left" Margin="283.92,205.52,0,0" VerticalAlignment="Top" Width="138.5" Height="27.46"/>
                    <GroupBox Header="Копирование групп" Margin="340.42,15.04,0,0" HorizontalAlignment="Left" Width="211" Height="112.96" VerticalAlignment="Top">
                        <Grid HorizontalAlignment="Left" Height="90.667" Margin="0,0,-2,-0.667" VerticalAlignment="Top" Width="201">
                            <RadioButton x:Name="group_radio_one" Content="С источника на одного" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,10,0,0" IsChecked="True"/>
                            <RadioButton x:Name="group_radio_many" Content="С источника для всех в OU" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,29.96,0,0"/>
                            <RadioButton x:Name="group_radio_many2" Content="С источника для нескольких" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,49.92,0,0"/>
                        </Grid>
                    </GroupBox>
                    <Button x:Name="btn_group_OU" Content="OU" HorizontalAlignment="Left" VerticalAlignment="Top" Width="40.37" Height="23" Margin="283.92,105,0,0"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_Inactive" Header="Inactive">
                <Grid Margin="0,0,-5.13,0.277">
                    <ListView x:Name="lstv_inactive" Margin="10,262.5,10,10">
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <ContextMenu.ContextMenu>
                                    <ContextMenu/>
                                </ContextMenu.ContextMenu>
                                <MenuItem x:Name="lstv_inactive_del" Header="Удалить из списка"/>
                                <MenuItem x:Name="lstv_inactive_move" Header="Перенести в Inactive"/>
                                <MenuItem x:Name="lstv_inactive_disable" Header="Отключить" Background="Red" Foreground="White"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Enabled" DisplayMemberBinding="{Binding Enabled}" Width="60"/>
                                <GridViewColumn Header="DisplayName" DisplayMemberBinding="{Binding DisplayName}" Width="250"/>
                                <GridViewColumn Header="SamAccountname" DisplayMemberBinding="{Binding SamAccountname}" Width="120"/>
                                <GridViewColumn Header="LastLogonAD" DisplayMemberBinding="{Binding LastLogonAD}" Width="100"/>
                                <GridViewColumn Header="LastLogonEX" DisplayMemberBinding="{Binding LastLogonEX}" Width="100"/>
                                <GridViewColumn Header="Whencreated" DisplayMemberBinding="{Binding Whencreated}" Width="100"/>
                                <GridViewColumn Header="Description" DisplayMemberBinding="{Binding Description}" Width="150"/>
                                <GridViewColumn Header="mail" DisplayMemberBinding="{Binding mail}" Width="150"/>
                                <GridViewColumn Header="DistinguishedName" DisplayMemberBinding="{Binding DistinguishedName}" Width="300"/>
                                <GridViewColumn Header="status" DisplayMemberBinding="{Binding status}" Width="100"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <GroupBox Header="Функции" Height="252.5" VerticalAlignment="Top" Margin="10,5,10,0" BorderBrush="#FFA09F9F">
                        <Grid Margin="0,0,-2,-4.043">
                            <ComboBox x:Name="inactive_cmbx_move" HorizontalAlignment="Left" VerticalAlignment="Top" Width="186" Margin="302,71.663,0,0">
                                <ComboBoxItem Content="OU_RT"/>
                                <ComboBoxItem Content="GBU"/>
                                <ComboBoxItem Content="MO"/>
                                <ComboBoxItem Content="ALL"/>
                            </ComboBox>
                            <GroupBox HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="350" Margin="10,98.623,0,0" Header="Основное меню">
                                <Grid HorizontalAlignment="Left" Height="79" Margin="0,0,-2,-1.96" VerticalAlignment="Top" Width="340">
                                    <Button x:Name="Inactive_btn_cleardannie" Content="очистить Inactive" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Margin="10,50.344,0,0"/>
                                    <CheckBox x:Name="inactive_cbox_delgr" Content="Удалить все группы" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="139,9.999,0,0"/>
                                    <CheckBox x:Name="inactive_cbox_dannie" Content="Очистить данные" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="139,50.196,0,0"/>
                                    <CheckBox x:Name="inactive_cbox_delman" Content="Удалить менеджеров" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="139,30.098,0,0"/>
                                    <RadioButton x:Name="inactive_radio_disable" Content="Заблокировать" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,9.999,0,0" IsChecked="True"/>
                                    <RadioButton x:Name="inactive_radio_move" Content="Перенести" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,30.236,0,0"/>
                                </Grid>
                            </GroupBox>
                            <Button x:Name="inactive_go" Content="В путь" HorizontalAlignment="Left" VerticalAlignment="Top" Width="152" Margin="515,113.583,0,0" Height="44.96"/>
                            <Button x:Name="inactive_OU" Content="OU" Margin="10,10,0,0" Height="24.46" VerticalAlignment="Top" HorizontalAlignment="Left" Width="42"/>
                            <TextBox x:Name="inactive_txtbox_ou" HorizontalAlignment="Left" Height="24.46" Margin="57,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="532"/>
                            <Button x:Name="inactive_load_disabled" Content="Загрузить отключенных" HorizontalAlignment="Left" VerticalAlignment="Top" Width="166" Margin="10,39.46,0,0" Height="24.96"/>
                            <Button x:Name="inactive_load_6months" Content="Загрузить не активных более" HorizontalAlignment="Left" VerticalAlignment="Top" Width="166" Margin="10,69.42,0,0" Height="24.96"/>
                            <TextBox x:Name="inactive_months" Height="24.96" TextWrapping="Wrap" VerticalAlignment="Top" Margin="181,69.42,0,0" MaxLines="1" MaxLength="2" HorizontalAlignment="Left" Width="45" TextAlignment="Center"/>
                            <Button x:Name="Inactive_button_clear" Content="Всего записей: 0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="166" Margin="10,203.623,0,0"/>
                            <Button x:Name="inactive_exch_btn" Content="Статус: не подключено" HorizontalAlignment="Left" VerticalAlignment="Top" Width="179" Margin="181,203.623,0,0" Background="Red" Foreground="White"/>
                            <GroupBox Header="Сортировка" HorizontalAlignment="Left" Height="120" VerticalAlignment="Top" Width="145" Margin="365,103.583,0,0">
                                <Grid HorizontalAlignment="Left" Height="113" VerticalAlignment="Top" Width="135" Margin="0,0,-2,-0.96">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="3*"/>
                                        <ColumnDefinition Width="2*"/>
                                    </Grid.ColumnDefinitions>
                                    <RadioButton x:Name="inactive_rb_lastloggonad" Content="LastLogonAD" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="21,24.92,0,0" Grid.ColumnSpan="2"/>
                                    <RadioButton x:Name="inactive_rb_displayname" Content="DisplayName" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="21,9.96,0,0" Grid.ColumnSpan="2"/>
                                    <RadioButton x:Name="inactive_rb_lastlogonex" Content="LastLogonEX" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="21,39.88,0,0" Grid.ColumnSpan="2"/>
                                    <RadioButton x:Name="inactive_rb_whencreated" Content="WhenCreated" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="21,54.84,0,0" Grid.ColumnSpan="2" IsChecked="True"/>
                                    <Button x:Name="inactive_btn_sort" Content="Сортировка" HorizontalAlignment="Left" VerticalAlignment="Top" Width="112" Margin="10,74.8,0,0" Grid.ColumnSpan="2" />
                                </Grid>
                            </GroupBox>
                            <Button x:Name="inactive_btn_delstatus" Content="Удалить польз. со статусом" HorizontalAlignment="Left" VerticalAlignment="Top" Width="180" Margin="515,199.38,0,0" Height="24.203"/>
                            <TextBox x:Name="inactive_txtbox_filter" HorizontalAlignment="Left" Height="24.46" Margin="302,42.203,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="186"/>
                            <Label Content="Фильтр DN" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="231.283,42.203,0,0" RenderTransformOrigin="0.001,0.378"/>
                            <Label Content="Перенос" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="244.34,67.663,0,0"/>
                            <Label Content="Фильтр Description" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="493,42.203,0,0" RenderTransformOrigin="0.001,0.378"/>
                            <TextBox x:Name="inactive_txtbox_filter_Descr" HorizontalAlignment="Left" Height="24.46" Margin="611.58,42.203,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="186"/>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>
        </TabControl>
        <Menu Height="27.304" VerticalAlignment="Top">
            <MenuItem Header="Файл" Height="27.304" Width="43.135" Background="{x:Null}" RenderTransformOrigin="0.5,0.5">
                <MenuItem x:Name="Btn_LoadFiles" Header="Загрузить данные из файла Users.csv" Height="23" Margin="0,0,0.998,0"/>
                <MenuItem x:Name="btn_export" Header="Выгрузить данные из списка в файл"/>
                <MenuItem x:Name="Btn_Exit" Header="Выход" Height="23" Margin="0,0,0.998,0"/>
            </MenuItem>
            <MenuItem x:Name="menu_obrab" Header="Обработка">
                <MenuItem x:Name="btn_GenPasswords" Header="Генерировать всем новые пароли"/>
                <MenuItem x:Name="btn_allobshie" Header="Пометить все записи как ОБЩИЕ УЗ"/>
                <MenuItem x:Name="btn_blockall" Header="Заблокировать всех в нижнем списке"/>
                <MenuItem x:Name="btn_changeall" Header="Изменить данные у всех"/>
                <MenuItem x:Name="btn_exporttoconsole" Header="Выгрузить в Console список" Background="#FF8ABF39"/>
            </MenuItem>
            <MenuItem x:Name="menu_settings" Header="Настройки">
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
                <CheckBox x:Name="chk_deletmanagermove" Content="Удалять менеджеров при переносе" IsChecked="False"/>
                <CheckBox x:Name="chk_deletgroupsmove" Content="Удалять группы при переносе" IsChecked="False"/>
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




[xml]$xamlMain = @'
<Window x:Name="windowSelectOU"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select OU" Height="623.912" Width="600" WindowStyle="None">
    <Grid>
        <TreeView x:Name="treeviewOUs" Margin="10,10,10.4,61.8"/>
        <Button x:Name="btnCancel" Content="Cancel" Margin="0,0,10.4,5.8" ToolTip="Filter" Height="23" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="71" IsCancel="True"/>
        <Button x:Name="btnSelect" Content="Select" Margin="0,0,86.4,5.8" ToolTip="Filter" HorizontalAlignment="Right" Width="71" Height="23" VerticalAlignment="Bottom" IsDefault="True"/>
        <TextBlock x:Name="txtSelectedOU" Margin="10,0,162.4,5.8" TextWrapping="Wrap" VerticalAlignment="Bottom" Height="23" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" IsEnabled="False"/>
        <TextBlock x:Name="txtSelectedOU_2" Margin="10,0,10.4,33.8" TextWrapping="Wrap" VerticalAlignment="Bottom" Height="23" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" IsEnabled="False"/>
    </Grid>
</Window>
'@


$reader=(New-Object System.Xml.XmlNodeReader $xamlMain) 
    $window=[Windows.Markup.XamlReader]::Load( $reader )

    $namespace = @{ x = 'http://schemas.microsoft.com/winfx/2006/xaml' }
    $xpath_formobjects = "//*[@*[contains(translate(name(.),'n','N'),'Name')]]" 

    # Create a variable for every named xaml element
    Select-Xml $xamlMain -Namespace $namespace -xpath $xpath_formobjects | Foreach {
        $_.Node | Foreach {
            Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name)
        }
    }