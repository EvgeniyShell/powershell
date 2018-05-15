function Get-OUDialog {
    <#
    .SYNOPSIS
    A self contained WPF/XAML treeview organizational unit selection dialog box.
    .DESCRIPTION
    A self contained WPF/XAML treeview organizational unit selection dialog box. No AD modules required, just need to be joined to the domain.
    .EXAMPLE
    $OU = Get-OUDialog
    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 4.0
    Version History
    1.0.0 - 03/21/2015
        - Initial release (the function is a bit overbloated because I'm simply embedding some of my prior functions directly
          in the thing instead of customizing the code for the function. Meh, it gets the job done...
    .LINK
    https://github.com/zloeber/Powershell/blob/master/ActiveDirectory/Select-OU/Get-OUDialog.ps1
    .LINK
    http://www.the-little-things.net
    #>
    [CmdletBinding()]
    param()
    
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$xamlMain = @'
<Window x:Name="windowSelectOU"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select OU" Height="623.912" Width="600">
    <Grid>
        <TreeView x:Name="treeviewOUs" Margin="10,10,10.4,61.8"/>
        <Button x:Name="btnCancel" Content="Cancel" Margin="0,0,10.4,5.8" ToolTip="Filter" Height="23" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="71" IsCancel="True"/>
        <Button x:Name="btnSelect" Content="Select" Margin="0,0,86.4,5.8" ToolTip="Filter" HorizontalAlignment="Right" Width="71" Height="23" VerticalAlignment="Bottom" IsDefault="True"/>
        <TextBlock x:Name="txtSelectedOU" Margin="10,0,162.4,5.8" TextWrapping="Wrap" VerticalAlignment="Bottom" Height="23" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" IsEnabled="False"/>
        <TextBlock x:Name="txtSelectedOU_2" Margin="10,0,10.4,33.8" TextWrapping="Wrap" VerticalAlignment="Bottom" Height="23" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" IsEnabled="False"/>
    </Grid>
</Window>
'@

    # Read XAML
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

    Add-TreeItem -TreeObj $domstruct -Name $domstruct.Name -Parent $treeviewOUs -Tag $domstruct.path

    $treeviewOUs.add_SelectedItemChanged({
        $convert = Convert-CNToDN $this.SelectedItem.Tag
        $txtSelectedOU.Text = $convert
        if (!($convert -eq "DC=ASO,DC=RT,DC=LOCAL"))
        {
        $txtSelectedOU_2.text = (Get-ADOrganizationalUnit $convert -Properties Description).description
        }
    })

    $btnSelect.add_Click({
        $script:DialogResult = $txtSelectedOU.Text
        $windowSelectOU.Close()
    })
    $btnCancel.add_Click({
        $script:DialogResult = $null
    })

    # Due to some bizarre bug with showdialog and xaml we need to invoke this asynchronously 
    #  to prevent a segfault
    $async = $windowSelectOU.Dispatcher.InvokeAsync({
        $retval = $windowSelectOU.ShowDialog()
    })
    $async.Wait() | Out-Null

    #Clear out previously created variables for every named xaml element to be nice...
    Select-Xml $xamlMain -Namespace $namespace -xpath $xpath_formobjects | Foreach {
        $_.Node | Foreach {
            Remove-Variable -Name ($_.Name)
        }
    }
    return $DialogResult
}

function Get-OUDialog-start {
if (($global:domstruct -eq $null) -or ($global:domstruct -eq ""))
{
$global:domstruct = @(Search-AD -DirectoryEntry $conn -Filter '(ObjectClass=organizationalUnit)' -Properties CanonicalName).CanonicalName | sort | Get-ChildOUStructure
}
return Get-OUDialog

}



function Get-ChildOUStructure {
        <#
        .SYNOPSIS
        Create JSON exportable tree view of AD OU (or other) structures.
        .DESCRIPTION
        Create JSON exportable tree view of AD OU (or other) structures in Canonical Name format.
        .PARAMETER ouarray
        Array of OUs in CanonicalName format (ie. domain/ou1/ou2)
        .PARAMETER oubase
        Base of OU
        .EXAMPLE
        $OUs = @(Get-ADObject -Filter {(ObjectClass -eq "OrganizationalUnit")} -Properties CanonicalName).CanonicalName
        $test = $OUs | Get-ChildOUStructure | ConvertTo-Json -Depth 20
        .NOTES
        Author: Zachary Loeber
        Requires: Powershell 3.0, Lync
        Version History
        1.0.0 - 12/24/2014
            - Initial release
        .LINK
        https://github.com/zloeber/Powershell/blob/master/ActiveDirectory/Get-ChildOUStructure.ps1
        .LINK
        http://www.the-little-things.net
        #>
        [CmdletBinding()]
        param(
            [Parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true, HelpMessage='Array of OUs in CanonicalName formate (ie. domain/ou1/ou2)')]
            [string[]]$ouarray,
            [Parameter(Position=1, HelpMessage='Base of OU.')]
            [string]$oubase = ''
        )
        begin {
            $newarray = @()
            $base = ''
            $firstset = $false
            $ouarraylist = @()
        }
        process {
            $ouarraylist += $ouarray
        }
        end {
            $ouarraylist = $ouarraylist | Where {($_ -ne $null) -and ($_ -ne '')} | Select -Unique | Sort-Object
            if ($ouarraylist.count -gt 0) {
                $ouarraylist | Foreach {
                   # $prioroupath = if ($oubase -ne '') {$oubase + '/' + $_} else {''}
                    $firstelement = @($_ -split '/')[0]
                    $regex = "`^`($firstelement`?`)"
                    $tmp = $_ -replace $regex,'' -replace "^(\/?)",''

                    if (-not $firstset) {
                        $base = $firstelement
                        $firstset = $true
                    }
                    else {
                        if (($base -ne $firstelement) -or ($tmp -eq '')) {
                            Write-Verbose "Processing Subtree for: $base"
                            $fulloupath = if ($oubase -ne '') {$oubase + '/' + $base} else {$base}
                            New-Object psobject -Property @{
                                'name' = $base
                                'path' = $fulloupath
                                'children' = if ($newarray.Count -gt 0) {,@(Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath)} else {$null}
                            }
                            $base = $firstelement
                            $newarray = @()
                            $firstset = $True
                        }
                    }
                    if ($tmp -ne '') {
                        $newarray += $tmp
                    }
                }
                Write-Verbose "Processing Subtree for: $base"
                $fulloupath = if ($oubase -ne '') {$oubase + '/' + $base} else {$base}
                New-Object psobject -Property @{
                    'name' = $base
                    'path' = $fulloupath
                    'children' = if ($newarray.Count -gt 0) {,@(Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath)} else {$null}
                }
            }
        }
    }



function Search-AD {
        # Original Author (largely unmodified btw): 
        #  http://becomelotr.wordpress.com/2012/11/02/quick-active-directory-search-with-pure-powershell/
        [CmdletBinding()]
        param (
            [string[]]$Filter,
            [string[]]$Properties = @('Name','ADSPath'),
            [string]$SearchRoot='',
            [switch]$DontJoinAttributeValues,
            [System.DirectoryServices.DirectoryEntry]$DirectoryEntry = $null
        )

        if ($DirectoryEntry -ne $null) {
            if ($SearchRoot -ne '') {
                $DirectoryEntry.set_Path($SearchRoot)
            }
        }
        else {
            $DirectoryEntry = [System.DirectoryServices.DirectoryEntry]$SearchRoot
        }

        if ($Filter) {
            $LDAP = "(&({0}))" -f ($Filter -join ')(')
        }
        else {
            $LDAP = "(name=*)"
        }
        try {
            (New-Object System.DirectoryServices.DirectorySearcher -ArgumentList @(
                $DirectoryEntry,
                $LDAP,
                $Properties
            ) -Property @{
                PageSize = 1000
            }).FindAll() | ForEach-Object {
                $ObjectProps = @{}
                $_.Properties.GetEnumerator() |
                    Foreach-Object {
                        $Val = @($_.Value)
                        if ($_.Name -ne $null) {
                            if ($DontJoinAttributeValues -and ($Val.Count -gt 1)) {
                                $ObjectProps.Add($_.Name,$_.Value)
                            }
                            else {
                                $ObjectProps.Add($_.Name,(-join $_.Value))
                            }
                        }
                    }
                if ($ObjectProps.psbase.keys.count -ge 1) {
                    New-Object PSObject -Property $ObjectProps | Select $Properties
                }
            }
        }
        catch {
            Write-Warning -Message ('Search-AD: Filter - {0}: Root - {1}: Error - {2}' -f $LDAP,$Root.Path,$_.Exception.Message)
        }
    }

function Convert-CNToDN {
        param([string]$CN)
        $SplitCN = $CN -split '/'
        if ($SplitCN.Count -eq 1) {
            return 'DC=' + (($SplitCN)[0] -replace '\.',',DC=')
        }
        else {
            $basedn = '.'+($SplitCN)[0] -replace '\.',',DC='
            [array]::Reverse($SplitCN)
            $ous = ''
            for ($index = 0; $index -lt ($SplitCN.count - 1); $index++) {
                $ous += 'OU=' + $SplitCN[$index] + ','
            }
            $result = ($ous + $basedn) -replace ',,',','
            return $result
        }
    }

function Add-TreeItem {
        param(
              $TreeObj,
              $Name,
              $Parent,
              $Tag
              )

            $ChildItem = New-Object System.Windows.Controls.TreeViewItem
            $ChildItem.Header = $Name
            $ChildItem.Tag = $Tag
            $Parent.Items.Add($ChildItem) | Out-Null

        if (($TreeObj.children).Count -gt 0) {
            foreach ($ou in $TreeObj.children) {
                $treeparent = Add-TreeItem -TreeObj $ou -Name $ou.Name -Parent $ChildItem -Tag $ou.path
            }
        }
    }

    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {               
        Write-Warning 'Run PowerShell.exe with -Sta switch, then run this script.'
        Write-Warning 'Example:'
        Write-Warning '    PowerShell.exe -noprofile -Sta'
        break
    }



