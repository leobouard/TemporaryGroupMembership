#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'ActiveDirectory' ; ModuleVersion = '1.0.1.0' }

$pam = Get-ADOptionalFeature -Filter * | Where-Object { $_.Name -eq 'Privileged Access Management Feature' }
if (!$pam.EnabledScopes) {
    Write-Host 'Privileged Access Management Feature is not enabled on the current domain!' -ForegroundColor Red
    break
}

function ConvertTo-TTLString {
    param([System.TimeSpan]$TimeSpan)

    $string = 'days', 'hours', 'minutes' | ForEach-Object { "$($TimeSpan.$_) $_" }
    $string = $string | Where-Object { $_ -notlike '0*' }
    $string = $string | ForEach-Object { if ($_ -like '1*') { $_ -replace 's', '' } else { $_ } }
    $string -join ', '
}

function Get-ADGroupMemberWithTTL {
    param([string]$Identity)

    (Get-ADGroup -Identity $Identity -Properties Members -ShowMemberTimeToLive).Members | ForEach-Object {
        if ($_ -match "<TTL=(\d+)>") { 
            $sec = $matches[1]
            $ttl = New-TimeSpan -Seconds $sec
            $ttl = ConvertTo-TTLString -TimeSpan $ttl
            $date = (Get-Date).AddSeconds($sec)
            $dn = ($_ -split ',' | Select-Object -Skip 1) -join ','
        }
        else {
            $ttl = $null
            $date = $null
            $dn = $_
        }

        Get-ADObject -Identity $dn | Select-Object *,
        @{ Name = 'TimeToLive' ; Expression = { $ttl } },
        @{ Name = 'RemoveDate' ; Expression = { Get-Date $date -Format 'yyyy-MM-dd HH:mm' } } |
        Select-Object Name, ObjectClass, TimeToLive, RemoveDate, DistinguishedName
    }
}

function Update-UIGroupMember {
    $group = Get-ADGroup $textboxSearchGroup.ToolTip -Properties Members
    $labelDirectMembers.Content = ($group.Members | Measure-Object).Count
    $datagridMembers.ItemsSource = @(Get-ADGroupMemberWithTTL $textboxSearchGroup.ToolTip)
    if ($checkboxTTL.IsChecked -eq $true) {
        $datagridMembers.ItemsSource = @($datagridMembers.ItemsSource | Where-Object { $_.TimeToLive })
    }
}

[xml]$xml1 = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="window" Title="Temporary group membership" MinHeight="500" MinWidth="400" Height="630" Width="400">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="140"/>
            <RowDefinition/>
            <RowDefinition Height="35"/>
        </Grid.RowDefinitions>
        <StackPanel Orientation="Vertical" Grid.Row="0">
            <Label Content="Search group"/>
            <Grid>
                <ComboBox Name="comboboxSearchGroup" IsEnabled="False" Height="25"/>
                <TextBox Name="textboxSearchGroup" Height="25" VerticalContentAlignment="Center"/>
            </Grid>
        </StackPanel>
        <GroupBox Name="groupboxGroupInfo" Header="Group information" IsEnabled="False" Grid.Row="1">
            <Grid Margin="5">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="105"/>
                    <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition/>
                    <RowDefinition/>
                    <RowDefinition/>
                    <RowDefinition/>
                </Grid.RowDefinitions>
                <Label Content="Description" Grid.Row="0" Grid.Column="0"/>
                <Label Name="labelDescription" Grid.Row="0" Grid.Column="1"/>
                <Label Content="Canonical name" Grid.Row="1" Grid.Column="0"/>
                <Label Name="labelCanonicalName" Grid.Row="1" Grid.Column="1"/>
                <Label Content="Scope &amp; category" Grid.Row="2" Grid.Column="0"/>
                <Label Name="labelGroupScopeCategory" Grid.Row="2" Grid.Column="1"/>
                <Label Content="Direct member(s)" Grid.Row="3" Grid.Column="0"/>
                <Label Name="labelDirectMembers" Grid.Row="3" Grid.Column="1"/>
            </Grid>
        </GroupBox>
        <GroupBox Name="groupboxMembers" Header="Members" IsEnabled="False" Grid.Row="2">
            <Grid Margin="5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="30"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <DataGrid Name="datagridMembers" Grid.Row="0" Grid.ColumnSpan="2" GridLinesVisibility="None" FontSize="12" IsReadOnly="True" Background="Transparent"/>
                <CheckBox Name="checkboxTTL" Content="Member(s) with TTL only" VerticalAlignment="Center" Grid.Row="1"/>
                <Button Name="buttonRefresh" Content="Refresh" Padding="10,0" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Grid.Row="1" Grid.Column="1"/>
            </Grid>
        </GroupBox>
        <StackPanel Name="stackpanelButtons1" Orientation="Horizontal" HorizontalAlignment="Right" IsEnabled="False" Grid.Row="3">
            <Button Name="buttonRemoveMember" Content="Remove selected member(s)" Padding="10,0" Margin="5"/>
            <Button Name="buttonAddMember" Content="Add temporary member" Padding="10,0" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
'@

[xml]$xml2 = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="window" Title="Add temporary member" Height="330" Width="450">
    <Grid Margin="15">
        <StackPanel Orientation="Vertical">
            <StackPanel Height="55" Margin="0,0,0,10" Orientation="Vertical">
                <Label Content="Search object"/>
                <Grid>
                    <ComboBox Name="comboboxSearchMember" IsEnabled="False" Height="25"/>
                    <TextBox Name="textboxSearchMember" Height="25" VerticalContentAlignment="Center"/>
                </Grid>
            </StackPanel>
            <GroupBox Name="groupboxMembershipDuration" Header="Membership duration" Margin="0,0,0,10" IsEnabled="False">
                <Grid Margin="0,10,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="30"/>
                        <RowDefinition Height="30"/>
                        <RowDefinition Height="30"/>
                        <RowDefinition Height="30"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="155"/>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Label Content="Choose when to remove the selected member from the group:" Grid.Row="0" Grid.ColumnSpan="2" Grid.RowSpan="2"/>
                    <RadioButton Name="radiobuttonNever" Content="Never" GroupName="TTL" VerticalAlignment="Center" Margin="20,0,0,0" Grid.Row="1" Height="15"/>
                    <RadioButton Name="radiobuttonHours" Content="After a few hours" GroupName="TTL" IsChecked="True" VerticalAlignment="Center" Margin="20,0,0,0"  Grid.Row="2" Height="15"/>
                    <Slider Name="sliderHours" Minimum="1" Maximum="12" TickFrequency="1" TickPlacement="BottomRight" IsSnapToTickEnabled="True" AutoToolTipPlacement="TopLeft" Grid.Row="2" Grid.Column="1" Margin="0" Grid.RowSpan="2"/>
                    <RadioButton Name="radiobuttonDate" Content="After a specific date" GroupName="TTL" VerticalAlignment="Center" Margin="20,0,0,0" Grid.Row="3" Height="15"/>
                    <DatePicker Name="datepickerDate" IsEnabled="False" Grid.Row="3" Grid.Column="1" Height="25"/>
                </Grid>
            </GroupBox>
            <StackPanel Name="stackpanelButtons2" Orientation="Horizontal" HorizontalAlignment="Right" Height="35" IsEnabled="False">
                <Button Name="buttonCancel" Content="Cancel" Padding="10,0" Margin="5"/>
                <Button Name="buttonConfirm" Content="Add to group" Padding="10,0" Margin="5"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
'@

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
$window1 = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml1))
$xml1.SelectNodes("//*[@Name]") | ForEach-Object { 
    Set-Variable -Name ($_.Name) -Value $window1.FindName($_.Name) -Scope Global
}

$textboxSearchGroup.Add_KeyDown{
    if ($_.Key -eq 'Return') {

        if ($textboxSearchGroup.Text) {
            $groupList = Get-ADGroup -Filter { ANR -eq $textboxSearchGroup.Text } -Properties Members, CanonicalName, Description
            switch (($groupList | Measure-Object).Count) {
                0 { $group = $null }
                1 { $group = $groupList }
                default { 
                    $comboboxSearchGroup.Items.Clear()
                    $groupList | Sort-Object Name | Select-Object -First 10 | ForEach-Object { $comboboxSearchGroup.AddChild($_.Name) }
                    $comboboxSearchGroup.IsEnabled = $true
                    $comboboxSearchGroup.IsDropDownOpen = $true
                }
            }
        }

        if ($group) {
            $groupboxGroupInfo.IsEnabled = $true
            $groupboxMembers.IsEnabled = $true
            $stackpanelButtons1.IsEnabled = $true
            $checkboxTTL.IsChecked = $false
            $textboxSearchGroup.Text = $group.Name
            $textboxSearchGroup.ToolTip = $group.DistinguishedName
            $labelDescription.Content = $group.Description
            $labelDescription.ToolTip = $group.Description
            $labelCanonicalName.Content = $group.CanonicalName
            $labelCanonicalName.ToolTip = $group.CanonicalName
            $labelGroupScopeCategory.Content = "$($group.GroupScope) - $($group.GroupCategory)"
            $labelDirectMembers.Content = ($group.Members | Measure-Object).Count
            $datagridMembers.ItemsSource = @(Get-ADGroupMemberWithTTL -Identity $group.DistinguishedName)
        }
    }
}

$textboxSearchGroup.Add_KeyUp{
    if (!$textboxSearchGroup.Text) {
        $groupboxGroupInfo.IsEnabled = $false
        $groupboxMembers.IsEnabled = $false
        $stackpanelButtons1.IsEnabled = $false
        $textboxSearchGroup.Text = $null
        $textboxSearchGroup.ToolTip = $null
        $labelDescription.Content = $null
        $labelDescription.ToolTip = $null
        $labelCanonicalName.Content = $null
        $labelCanonicalName.ToolTip = $null
        $labelGroupScopeCategory.Content = $null
        $labelDirectMembers.Content = $null
        $datagridMembers.ItemsSource = @()
    }
}

$comboboxSearchGroup.Add_DropDownClosed{
    $comboboxSearchGroup.IsEnabled = $false
    $group = Get-ADGroup -Filter { Name -eq $comboboxSearchGroup.Text } -Properties Members, CanonicalName, Description

    if ($group) {
        $groupboxGroupInfo.IsEnabled = $true
        $groupboxMembers.IsEnabled = $true
        $stackpanelButtons1.IsEnabled = $true
        $checkboxTTL.IsChecked = $false
        $textboxSearchGroup.Text = $group.Name
        $textboxSearchGroup.ToolTip = $group.DistinguishedName
        $labelDescription.Content = $group.Description
        $labelDescription.ToolTip = $group.Description
        $labelCanonicalName.Content = $group.CanonicalName
        $labelCanonicalName.ToolTip = $group.CanonicalName
        $labelGroupScopeCategory.Content = "$($group.GroupScope) - $($group.GroupCategory)"
        $labelDirectMembers.Content = ($group.Members | Measure-Object).Count
        $datagridMembers.ItemsSource = @(Get-ADGroupMemberWithTTL -Identity $group.DistinguishedName)
    }
}

$checkboxTTL.Add_Checked{
    $datagridMembers.ItemsSource = @(Get-ADGroupMemberWithTTL $textboxSearchGroup.ToolTip | Where-Object { $_.TimeToLive })
}

$checkboxTTL.Add_Unchecked{
    $datagridMembers.ItemsSource = @(Get-ADGroupMemberWithTTL $textboxSearchGroup.ToolTip)
}

$buttonRefresh.Add_Click{
    Update-UIGroupMember
}

$buttonRemoveMember.Add_Click{

    $splat = @{
        Identity    = $textboxSearchGroup.ToolTip
        Members     = $datagridMembers.SelectedItems.DistinguishedName
        Confirm     = $false
        Verbose     = $true
        ErrorAction = 'Stop'
    }

    Add-Type -AssemblyName PresentationFramework
    $title = 'Remove selected member(s)'
    try {
        Remove-ADGroupMember @splat
        $null = [System.Windows.MessageBox]::Show('The member(s) have been removed from the group', $title, 0, 64)
    }
    catch {
        $errorMessage = $_.Exception.Message
        $null = [System.Windows.MessageBox]::Show($errorMessage, $title, 0, 16)
    }
    
    Update-UIGroupMember
}

$buttonAddMember.Add_Click{
    $window2 = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml2))
    $xml2.SelectNodes("//*[@Name]") | ForEach-Object { 
        Set-Variable -Name ($_.Name) -Value $window2.FindName($_.Name) -Scope Global
    }
    $datepickerDate.SelectedDate = (Get-Date).AddDays(3)

    $radiobuttonHours.Add_Checked{
        $sliderHours.IsEnabled = $true
    }

    $radiobuttonHours.Add_Unchecked{
        $sliderHours.IsEnabled = $false
    }

    $radiobuttonDate.Add_Checked{
        $datepickerDate.IsEnabled = $true
    }

    $radiobuttonDate.Add_Unchecked{
        $datepickerDate.IsEnabled = $false
    }

    $textboxSearchMember.Add_KeyDown{
        if ($_.Key -eq 'Return') {

            if ($textboxSearchMember.Text) {
                $memberList = Get-ADObject -Filter { ANR -eq $textboxSearchMember.Text }
                switch (($memberList | Measure-Object).Count) {
                    0 { $member = $null }
                    1 { $member = $memberList }
                    default { 
                        $comboboxSearchMember.Items.Clear()
                        $memberList | Sort-Object Name | Select-Object -First 10 | ForEach-Object { $comboboxSearchMember.AddChild($_.DistinguishedName) }
                        $comboboxSearchMember.IsEnabled = $true
                        $comboboxSearchMember.IsDropDownOpen = $true
                    }
                }
            }

            if ($member) {
                $groupboxMembershipDuration.IsEnabled = $true
                $stackpanelButtons2.IsEnabled = $true
                $textboxSearchMember.Text = $member.Name
                $textboxSearchMember.ToolTip = $member.DistinguishedName
            }
        }
    }

    $textboxSearchMember.Add_KeyUp{
        if (!$textboxSearchMember.Text) {
            $groupboxMembershipDuration.IsEnabled = $false
            $stackpanelButtons2.IsEnabled = $false
            $textboxSearchMember.Text = $null
            $textboxSearchMember.ToolTip = $null
        }
    }

    $comboboxSearchMember.Add_DropDownClosed{
        $comboboxSearchMember.IsEnabled = $false
        if ($comboboxSearchMember.Text) {
            $member = Get-ADObject $comboboxSearchMember.Text
        }

        if ($member) {
            $groupboxMembershipDuration.IsEnabled = $true
            $stackpanelButtons2.IsEnabled = $true
            $textboxSearchMember.Text = $member.Name
            $textboxSearchMember.ToolTip = $member.DistinguishedName
        }
    }

    $buttonConfirm.Add_Click{

        $splat = @{
            Identity    = $textboxSearchGroup.ToolTip
            Members     = $textboxSearchMember.ToolTip
            Verbose     = $true
            ErrorAction = 'Stop'
        }

        if ($radiobuttonHours.IsChecked) {
            $ttl = New-TimeSpan -Hours $sliderHours.Value
            $splat.Add('MemberTimeToLive', $ttl)
        }
        if ($radiobuttonDate.IsChecked) {
            $ttl = New-TimeSpan -Start (Get-Date -S 0 -Mil 0) -End ($datepickerDate.SelectedDate)
            $splat.Add('MemberTimeToLive', $ttl) 
        }

        Add-Type -AssemblyName PresentationFramework
        $title = 'Add temporary member'
        try {
            Add-ADGroupMember @splat
            $null = [System.Windows.MessageBox]::Show('The object was added to the group', $title, 0, 64)
        }
        catch {
            $errorMessage = $_.Exception.Message
            $null = [System.Windows.MessageBox]::Show($errorMessage, $title, 0, 16)
        }
        $window2.Close()

        Update-UIGroupMember
    }

    $buttonCancel.Add_Click{
        $window2.Close()
    }

    $null = ($window2.Dispatcher.InvokeAsync{ $window2.ShowDialog() }).Wait()
}

$null = ($window1.Dispatcher.InvokeAsync{ $window1.ShowDialog() }).Wait()