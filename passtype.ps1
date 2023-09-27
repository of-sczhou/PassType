### SciptStart
$appVersion = "1.0.0.0"
$appName = "PassType"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Hide powershell console (if not running in Powershell ISE)
if ((Get-Process -PID $pid).ProcessName -ne "powershell_ise") { $null = $(Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -name Win32ShowWindowAsync -namespace Win32Functions -PassThru)::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0) }

$ExecDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\')

Class DBInstance {
    [string]$DBPath
    [string]$DBKeyPath
    [bool]$Include
    [SecureString]$DBMasterKey
}

Class EntryBrief {
    [string]$Uuid
    [string]$Name
    [string]$DBPath
    [Int32]$OrderNum
    [bool]$IsVisible
}

[DBInstance[]]$Global:DBInstances = @()
[EntryBrief[]]$Global:AttributedEntries = @([EntryBrief]::new())
[bool[]]$Global:CheckBoxes = @($false,$false)

Import-Module -Name $($ExecDir + "\poshkeepass")

[xml]$XAMLMainWindow = @"
<Window x:Name="Window_Main"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="PassType" Height="67" Width="130" ResizeMode="CanResize" WindowStyle="None" BorderThickness="0" AllowsTransparency="True" Background="Transparent" Topmost="{Binding ElementName=CheckBox_AlwaysOnTop, Path=IsChecked}">
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="0" ResizeBorderThickness="5"/>
    </WindowChrome.WindowChrome>
    <Border x:Name="WindowMain_Border" CornerRadius="7" BorderBrush="#FF263238" BorderThickness="1" Background="#FF9FC4D6">
        <Grid x:Name="WindowMain_Grid">
            <Button x:Name="Button_Hide" Background="Transparent" HorizontalAlignment="Right" Height="20" Width="20" VerticalAlignment="Top" BorderThickness="0,0,0,2" BorderBrush="Black" Margin="0,4,4,0"/>
            <Button x:Name="Button_Filter" Content="..." Background="Transparent" HorizontalAlignment="Left" Height="20" Width="20" VerticalAlignment="Top" BorderThickness="0" BorderBrush="Black" Margin="7,4,0,0" FontWeight="Bold" FontSize="16">
                <Button.ToolTip>
                    <ToolTip>Filter, Order, Refresh</ToolTip>
                </Button.ToolTip>
            </Button>
            <CheckBox x:Name="CheckBox_AlwaysOnTop" HorizontalAlignment="Left" Margin="34,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="Black">
                <CheckBox.ToolTip>
                    <ToolTip>Always on Top</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <CheckBox x:Name="CheckBox_AutoRun" HorizontalAlignment="Center" Margin="0,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="Black">
                <CheckBox.ToolTip>
                    <ToolTip>Autorun</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <CheckBox x:Name="CheckBox_AutoComplete" HorizontalAlignment="Center" Margin="0,6,-43,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="Black">
                <CheckBox.ToolTip>
                    <ToolTip>Auto Complete</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <Grid x:Name="WindowMain_KPButtons_Grid" Margin="0,29,0,0"/>
            <Button x:Name="Button_Clipboard" Content="Clipboard" Margin="6,0,6,5" VerticalAlignment="Bottom" Background="Transparent" BorderBrush="Black" Height="20"/>
        </Grid>
    </Border>
</Window>
"@
$Reader=(New-Object System.Xml.XmlNodeReader $XAMLMainWindow)
try { $Window_main = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
$XAMLMainWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_main.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

[xml]$XAMLSelectorWindow = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        x:Name="Window_Selector" Title="Filter, Order, Refresh" Height="250" Width="650" ResizeMode="CanResize" WindowStyle="None" SnapsToDevicePixels="True" BorderThickness="1" AllowsTransparency="True" Background="White" BorderBrush="{DynamicResource {x:Static SystemColors.ControlDarkBrushKey}}" WindowStartupLocation="CenterScreen">
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="0" ResizeBorderThickness="5"/>
    </WindowChrome.WindowChrome>
    <Grid>
        <ListView BorderThickness="0" x:Name="ListView_Selector" SelectionMode="Single" Margin="0,22,0,0">
            <ListView.View>
                <GridView x:Name="GridView_Selector">
                    <GridViewColumn Header="IsVisible" Width="NaN">
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <Grid HorizontalAlignment="Stretch">
                                    <CheckBox  IsChecked="{Binding IsVisible}" />
                                </Grid>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Header="Name" Width="NaN" DisplayMemberBinding="{Binding Name}"/>
                    <GridViewColumn Header="uuid" Width="NaN" DisplayMemberBinding ="{Binding Uuid}"/>
                    <GridViewColumn Header="Database" Width="NaN"  DisplayMemberBinding ="{Binding DBPath}"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="Selector_Button_Apply" Content=" Apply " HorizontalAlignment="Right" Margin="0,1,53,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0"/>
        <Button x:Name="Selector_Button_Cancel" Content=" Cancel " HorizontalAlignment="Right" Margin="0,1,2,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0"/>
        <Button x:Name="Selector_Button_Up" Content="▲" Background="White" HorizontalAlignment="Left" Height="18" VerticalAlignment="Top" BorderThickness="0" Width="18" Margin="6,1,0,0" Padding="1,-4,1,1"/>
        <Button x:Name="Selector_Button_Down" Content="▼" Background="White" HorizontalAlignment="Left" Height="18" VerticalAlignment="Top" BorderThickness="0" Width="18" Margin="25,0,0,0" Padding="1,4,1,1"/>
        <Label Content="change order" HorizontalAlignment="Left" Margin="44,-3,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>
"@
$Reader=(New-Object System.Xml.XmlNodeReader $XAMLSelectorWindow)
try { $Window_Selector = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
$XAMLSelectorWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_Selector.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

Try {
    $AppSettings = Get-Content -Path $($ExecDir + "\" + $appName + ".ini") -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
    $Global:DBInstances = $AppSettings[0].value
    $Global:AttributedEntries = $AppSettings[1].value
    $Global:CheckBoxes = $AppSettings[2].value
} catch {
    if (-Not (Get-Item -Path $($ExecDir + "\" + $appName + ".ini") -ErrorAction SilentlyContinue)) {[System.Windows.MessageBox]::Show("Exception occured while reading content of config file. Please check it's content or delete config file and relaunch application.")}
}

function PassType_Entrance {
[xml]$XAMLWindow_PassType_Entrance = @"
<Window x:Name="Window_PassType_Entrance"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="PassType Entrance" Height="110" Width="250" ResizeMode="NoResize" ShowInTaskbar="False" Topmost="True" WindowStartupLocation="CenterScreen" WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid >
        <TabControl x:Name="TabControl" Height="NaN" Margin="0,0,0,0" SelectedIndex="0">
            <TabItem x:Name="TabItem1" Header="1">
                <TabItem.ToolTip>
                    <ToolTip>DataBase Instance 1</ToolTip>
                </TabItem.ToolTip>
                <Grid Background="{Binding ElementName=Window_PassTypeAuthentication, Path=BackGround}">
                    <TextBox x:Name="TextBox_DBPath_1" Margin="3,3,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem1, Path=Background}">
                        <TextBox.ToolTip>
                            <ToolTip>Path to Keepass Database File</ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>
                    <TextBox x:Name="TextBox_KeyPath_1" Margin="3,23,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem1, Path=Background}">
                        <TextBox.ToolTip>
                            <ToolTip>Path to Keepass Key File if applicable</ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>
                    <PasswordBox x:Name="PasswordBox_MasterKey_1" Margin="3,44,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem3, Path=Background}" IsEnabled="{Binding ElementName=CheckBox_Include_1, Path=IsChecked }" PasswordChar="*">
                        <PasswordBox.ToolTip>
                            <ToolTip>MasterKey</ToolTip>
                        </PasswordBox.ToolTip>
                    </PasswordBox>
                    <CheckBox x:Name="CheckBox_Include_1" Content="Include" HorizontalAlignment="Left" Margin="3,0,0,2" VerticalAlignment="Bottom" IsChecked="False" BorderBrush="#FFABADB3"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="TabItem2" Header="2">
                <TabItem.ToolTip>
                    <ToolTip>DataBase Instance 2</ToolTip>
                </TabItem.ToolTip>
                <Grid Background="{Binding ElementName=Window_PassTypeAuthentication, Path=BackGround}">
                    <TextBox x:Name="TextBox_DBPath_2" Margin="3,3,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem2, Path=Background}">
                        <TextBox.ToolTip>
                            <ToolTip>Path to Keepass Database File</ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>
                    <TextBox x:Name="TextBox_KeyPath_2" Margin="3,23,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem2, Path=Background}">
                        <TextBox.ToolTip>
                            <ToolTip>Path to Keepass Key File if applicable</ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>
                    <PasswordBox x:Name="PasswordBox_MasterKey_2" Margin="3,43,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem3, Path=Background}" IsEnabled="{Binding ElementName=CheckBox_Include_2, Path=IsChecked }" PasswordChar="*">
                        <PasswordBox.ToolTip>
                            <ToolTip>MasterKey</ToolTip>
                        </PasswordBox.ToolTip>
                    </PasswordBox>
                    <CheckBox x:Name="CheckBox_Include_2" Content="Include" HorizontalAlignment="Left" Margin="3,0,0,2" VerticalAlignment="Bottom" IsChecked="False" BorderBrush="#FFABADB3"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="TabItem3" Header="3">
                <TabItem.ToolTip>
                    <ToolTip>DataBase Instance 3</ToolTip>
                </TabItem.ToolTip>
                <Grid Background="{Binding ElementName=Window_PassTypeAuthentication, Path=BackGround}">
                    <TextBox x:Name="TextBox_DBPath_3" Margin="3,3,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem3, Path=Background}">
                        <TextBox.ToolTip>
                            <ToolTip>Path to Keepass Database File</ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>
                    <TextBox x:Name="TextBox_KeyPath_3" Margin="3,23,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem3, Path=Background}">
                        <TextBox.ToolTip>
                            <ToolTip>Path to Keepass Key File if applicable</ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>
                    <PasswordBox x:Name="PasswordBox_MasterKey_3" Margin="3,43,3,0" VerticalAlignment="Top" Background="{Binding ElementName=TabItem3, Path=Background}" IsEnabled="{Binding ElementName=CheckBox_Include_3, Path=IsChecked }" PasswordChar="*">
                        <PasswordBox.ToolTip>
                            <ToolTip>MasterKey</ToolTip>
                        </PasswordBox.ToolTip>
                    </PasswordBox>
                    <CheckBox x:Name="CheckBox_Include_3" Content="Include" HorizontalAlignment="Left" Margin="3,0,0,2" VerticalAlignment="Bottom" IsChecked="False" BorderBrush="#FFABADB3"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Button x:Name="Button_OK" Content="OK" HorizontalAlignment="Right" Margin="0,0,42,2" VerticalAlignment="Bottom" Height="20" Width="32" BorderBrush="#FFABADB3"/>
        <Button x:Name="Button_Quit" Content="Quit" HorizontalAlignment="Right" Margin="0,0,6,2" VerticalAlignment="Bottom" Width="32" BorderBrush="#FFABADB3"/>
    </Grid>
</Window>
"@
    $Reader=(New-Object System.Xml.XmlNodeReader $XAMLWindow_PassType_Entrance)
    try { $Window_PassType_Entrance = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
    $XAMLWindow_PassType_Entrance.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_PassType_Entrance.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

    $Button_OK.add_Click.Invoke({
        [DBInstance[]]$Global:DBInstances = @()

        1..3 | % {
            $DBNew = New-Object -TypeName DBInstance
            $DBNew.DBPath = (Get-Variable -Name $("TextBox_DBPath_" + $_)).Value.Text
            $DBNew.DBKeyPath = (Get-Variable -Name $("TextBox_KeyPath_" + $_)).Value.Text
            $DBNew.Include = (Get-Variable -Name $("CheckBox_Include_" + $_)).Value.IsChecked
            $Global:DBInstances += $DBNew
        }

        if (($PasswordBox_MasterKey_1.Password -ne "") -and $Global:DBInstances[0].Include) {$Global:DBInstances[0].DBMasterKey = $(ConvertTo-SecureString $($PasswordBox_MasterKey_1.Password) -AsPlainText -Force)}
        if (($PasswordBox_MasterKey_2.Password -ne "") -and $Global:DBInstances[1].Include) {$Global:DBInstances[1].DBMasterKey = $(ConvertTo-SecureString $($PasswordBox_MasterKey_2.Password) -AsPlainText -Force)}
        if (($PasswordBox_MasterKey_3.Password -ne "") -and $Global:DBInstances[2].Include) {$Global:DBInstances[2].DBMasterKey = $(ConvertTo-SecureString $($PasswordBox_MasterKey_3.Password) -AsPlainText -Force)}
        
        #Check Connection
        $AuthOK = $true
        $Global:DBInstances | ? {$_.Include} | % {
            Try {
                Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
                $DBStr = $_.DBPath
                If ($_.DBKeyPath) {
                    New-KeePassDatabaseConfiguration -Default -DatabaseProfileName 'TempDatabase' -DatabasePath $_.DBPath -UseMasterKey -KeyPath $_.DBKeyPath
                } else {
                    New-KeePassDatabaseConfiguration -Default -DatabaseProfileName 'TempDatabase' -DatabasePath $_.DBPath -UseMasterKey
                }
                Get-KeePassEntry -DatabaseProfileName 'TempDatabase' -MasterKey $_.DBMasterKey
            } catch {$AuthOK = $false ; [System.Windows.MessageBox]::Show($("Error connecting to DataBase " + $DBStr))}
        }

        If ($AuthOK) {
            Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
            $Window_PassType_Entrance.Close()
        }
    })

    $Button_Quit.add_Click.Invoke({ $Window_PassType_Entrance.Close() ; Exit })

    $TabItem1.Add_GotFocus({ if ($CheckBox_Include_1.IsChecked) {$PasswordBox_MasterKey_1.Focus()} else {$TextBox_DBPath_1.Focus()} })
    $TabItem2.Add_GotFocus({ if ($CheckBox_Include_2.IsChecked) {$PasswordBox_MasterKey_2.Focus()} else {$TextBox_DBPath_2.Focus()} })
    $TabItem3.Add_GotFocus({ if ($CheckBox_Include_3.IsChecked) {$PasswordBox_MasterKey_3.Focus()} else {$TextBox_DBPath_3.Focus()} })

    $Window_PassType_Entrance.Add_Loaded({
        if ($Global:DBInstances.Count -ne 0) {
            1..3 | % {
               if ($Global:DBInstances[$_ - 1].DBPath -ne "") {
                    (Get-Variable -Name $("TextBox_DBPath_" + $_)).Value.Text = $Global:DBInstances[$_ - 1].DBPath
                    (Get-Variable -Name $("TextBox_KeyPath_" + $_)).Value.Text = $Global:DBInstances[$_ - 1].DBKeyPath
                    (Get-Variable -Name $("CheckBox_Include_" + $_)).Value.IsChecked = $Global:DBInstances[$_ - 1].Include
                    if (((Get-Variable -Name $("CheckBox_Include_" + $_)).Value).IsChecked) {(Get-Variable -Name $("PasswordBox_MasterKey_" + $_)).Value.Focus()}
                } 
            }
        }
    })
    $Window_PassType_Entrance.Activate() | Out-Null
    $Window_PassType_Entrance.Focus() | Out-Null
    $Window_PassType_Entrance.ShowDialog()
}

PassType_Entrance

Add-Type -Path $($ExecDir + "\InputManager.dll")

# Common Variables, types
$Global:Delay = 20
$Global:FadeDelay = 15
$InitialWindowHeight = $Window_main.Height
$Global:FadeAllowed = $true
[Diagnostics.Stopwatch]$Global:timer = New-Object Diagnostics.Stopwatch

Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WinAp {
      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool SetForegroundWindow(IntPtr hWnd);

      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@

Function SHIFT_KEY {
    param(
        [string]$KEY
    )

    [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::ShiftKey) ; Start-Sleep -Milliseconds $Global:Delay
    [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; Start-Sleep -Milliseconds $Global:Delay
    [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY) ; Start-Sleep -Milliseconds $Global:Delay
    [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::ShiftKey) ; Start-Sleep -Milliseconds $Global:Delay
}

Function SINGLE_KEY {
    param(
        [string]$KEY
    )

    [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; Start-Sleep -Milliseconds $Global:Delay
    [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY)
}

Function SendKey {
    param (
            [string]$KEY
    )
    
    Switch -regex -CaseSensitive ($KEY) {
        '^[A-Z]' { SHIFT_KEY $KEY }
        '^[a-z]' { [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY) }
        '^[0-9]' { [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::("D"+$KEY)) ; [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::("D"+$KEY)) }
        DEFAULT {
            Switch ($KEY) {
                "~" { SHIFT_KEY "Oem3" }
                "!" { SHIFT_KEY "D1"}
                "@" { SHIFT_KEY "D2"}
                "#" { SHIFT_KEY "D3"}
                "$" { SHIFT_KEY "D4"}
                "%" { SHIFT_KEY "D5"}
                "^" { SHIFT_KEY "D6"}
                "&" { SHIFT_KEY "D7"}
                "*" { SHIFT_KEY "D8"}
                "(" { SHIFT_KEY "D9"}
                ")" { SHIFT_KEY "D0"}
                "_" { SHIFT_KEY "OemMinus"}
                "+" { SHIFT_KEY "Oemplus" }
                "<" { SHIFT_KEY "Oemcomma"}
                ">" { SHIFT_KEY "OemPeriod"}
                "?" { SHIFT_KEY "Oem2"}
                ":" { SHIFT_KEY "Oem1"}
                """" { SHIFT_KEY "Oem7"}
                "|" { SHIFT_KEY "Oem5"}
                "{" { SHIFT_KEY "Oem4"}
                "}" { SHIFT_KEY "Oem6"}
                "``" {SINGLE_KEY "Oem3"}
                "-" {SINGLE_KEY "OemMinus"}
                "=" {SINGLE_KEY "Oemplus"}
                "," {SINGLE_KEY "Oemcomma"}
                "." {SINGLE_KEY "OemPeriod"}
                "/" {SINGLE_KEY "Oem2"}
                ";" {SINGLE_KEY "Oem1"}
                "'" {SINGLE_KEY "Oem7"}
                "\" {SINGLE_KEY "Oem5"}
                "[" {SINGLE_KEY "Oem4"}
                "]" {SINGLE_KEY "Oem6"}
                " " {SINGLE_KEY "Space"}
                "`n" {SINGLE_KEY "Enter"}
                "`t" {SINGLE_KEY "Tab"}
                DEFAULT {}
            }
        }
    }
    Start-Sleep -Milliseconds $Global:Delay
}

Function Send_Credentials {
    param(
        [string]$uuid,
        [bool]$PasswordOnly
    )

    $Global:DBInstances | ? {$_.Include} | % {
        $DatabasePath = $_.DBPath
        $TryGetEntry = Get-KeePassEntry -MasterKey $_.DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $DatabasePath}).Name)  | ? {$($_.uuid.Tostring()) -eq $uuid}
        If ($TryGetEntry) {$Entry = $TryGetEntry}
    }

    Start-sleep -Milliseconds 100

    if (-Not $PasswordOnly) {
        $Entry.UserName.ToCharArray() | % { SendKey $_ }
        Start-sleep -Milliseconds 100
        [InputManager.Keyboard]::KeyPress([System.Windows.Forms.Keys]::Tab)
    }
    # Ожидание, что пользователь отпустит кнопку Ctrl
    Start-sleep -Milliseconds 500

    $(([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Entry.Password)))).ToCharArray() | % { SendKey $_ }
    if ($CheckBox_AutoComplete.IsChecked) {[InputManager.Keyboard]::KeyPress([System.Windows.Forms.Keys]::Enter)}
    Return
}

Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
$Global:DBInstances | ? {$_.Include} | % {
    $RandomIndex = (Get-Random).ToString()
    if ($_.DBKeyPath) {
        if (Get-KeePassDatabaseConfiguration) {
            New-KeePassDatabaseConfiguration -DatabaseProfileName $RandomIndex -DatabasePath $_.DBPath -UseMasterKey -KeyPath $_.DBKeyPath
        } else {
            New-KeePassDatabaseConfiguration -Default -DatabaseProfileName $RandomIndex -DatabasePath $_.DBPath -UseMasterKey -KeyPath $_.DBKeyPath
        }
    } else {
        if (Get-KeePassDatabaseConfiguration) {
            New-KeePassDatabaseConfiguration -DatabaseProfileName $RandomIndex -DatabasePath $_.DBPath -UseMasterKey
        } else {
            New-KeePassDatabaseConfiguration -Default -DatabaseProfileName $RandomIndex -DatabasePath $_.DBPath -UseMasterKey
        }
    }
}

function DrawButtons {
    [EntryBrief[]]$EntriesUnsorted = @()
    $Global:DBInstances | ? {$_.Include} | % {
        $DatabasePath = $_.DBPath
        Get-KeePassEntry -MasterKey $_.DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $DatabasePath}).Name) | ? {$_.FullPath -notlike "*/Recycle Bin"} | % {
            $Uuid = $_.uuid.ToString()
            $AttributedEntry = $Global:AttributedEntries.Where({$_.uuid -eq $Uuid})
            If (($AttributedEntry).Count -eq 0) {
                [EntryBrief]$NewEntry = New-Object -TypeName EntryBrief
                $NewEntry.uuid = $Uuid
                $NewEntry.Name = $_.Title
                $NewEntry.DBPath = $_.DBPath
                $NewEntry.OrderNum = 0
                $NewEntry.IsVisible = $true
                $EntriesUnsorted += $NewEntry
            } else {
                if ($AttributedEntry.IsVisible) {
                    [EntryBrief]$NewEntry = New-Object -TypeName EntryBrief
                    $NewEntry.uuid = $Uuid
                    $NewEntry.Name = $_.Title
                    $NewEntry.DBPath = $_.DBPath
                    $NewEntry.OrderNum = $AttributedEntry.OrderNum
                    $NewEntry.IsVisible = $true
                    $EntriesUnsorted += $NewEntry
                }
            }
        }    
    }

    [EntryBrief[]]$EntriesSorted = @()
    $EntriesSorted += $EntriesUnsorted.Where({$_.OrderNum -ne 0}) | Sort-Object OrderNum
    $EntriesSorted += $EntriesUnsorted.Where({$_.OrderNum -eq 0}) | Sort-Object Name

    If ($WindowMain_KPButtons_Grid.Children.Count -ne 0) {$WindowMain_KPButtons_Grid.Children.RemoveRange(0,$($WindowMain_KPButtons_Grid.Children.Count))}
    $Window_Main.Height = $InitialWindowHeight
    $i = 0
    $EntriesSorted | % {
        $Window_main.Height += 20
        $Button = [System.Windows.Controls.Button]::new()
        $Button.Name = "Button_" + $_.uuid
        $Button.Content = $_.Name
        $Button.HorizontalAlignment = "Center" ; $Button.VerticalAlignment = "Top"
        $Button.Width = 117 ; $Button.Height = 20 
        $Button.Margin = "0,$([string]($i*($Button.Height - 1))),0,0"
        $Button.Background = "Transparent"
        $Button.Add_Click({
            Send_Credentials $($This.Name.Substring(7)) $(([System.Windows.Input.Keyboard]::IsKeyDown("LeftCtrl")) -or ([System.Windows.Input.Keyboard]::IsKeyDown("RightCtrl")))
        })
        $WindowMain_KPButtons_Grid.Children.Add($Button) | Out-Null
        #[System.Windows.Data.Binding]$binding = [System.Windows.Data.Binding]::new("Background")
        #$binding.Source = $WindowMain_Border
        #($WindowMain_Grid.Children.Where({$_.Name -eq $Button.Name})).SetBinding([System.Windows.Controls.Border]::BackgroundProperty,$binding) | Out-Null
        $i += 1
    }
}

DrawButtons

function ArrangeEntries {
    [EntryBrief[]]$EntriesAll = @()

    $Global:DBInstances | ? {$_.Include} | % {
        $DatabasePath = $_.DBPath
        Get-KeePassEntry -MasterKey $_.DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $DatabasePath}).Name) | ? {$_.FullPath -notlike "*/Recycle Bin"} | % {
            $Uuid = $_.uuid.Tostring()
            $AttributedEntry = $Global:AttributedEntries.Where({$_.uuid -eq $Uuid})
            [EntryBrief]$NewEntry = New-Object -TypeName EntryBrief
            $NewEntry.uuid = $Uuid
            $NewEntry.Name = $_.Title
            $NewEntry.DBPath = $DatabasePath
            If (($AttributedEntry).Count -eq 0) {
                $NewEntry.OrderNum = 0
                $NewEntry.IsVisible = $true
            } else {
                if ($AttributedEntry.IsVisible) {
                    $NewEntry.OrderNum = $AttributedEntry.OrderNum
                    $NewEntry.IsVisible = $true
                } else {
                    $NewEntry.OrderNum = 0
                    $NewEntry.IsVisible = $false
                }
            }
            $EntriesAll += $NewEntry
        }    
    }

    [EntryBrief[]]$EntriesArranged = @()
    $EntriesArranged += $EntriesAll.Where({($_.OrderNum -ne 0 -and $_.IsVisible)}) | Sort-Object OrderNum
    $EntriesArranged += $EntriesAll.Where({($_.OrderNum -eq 0 -and $_.IsVisible)}) | Sort-Object Name
    $EntriesArranged += $EntriesAll.Where({-Not $_.IsVisible}) | Sort-Object Name
    Return $EntriesArranged
}

$Global:CurrentEntries = ArrangeEntries

Try { if (Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName) {$CheckBox_AutoRun.IsChecked = $true} } catch {}

## from https://gist.github.com/selvalogesh/37b99e43b932d42b5a9901a33284b4fa
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')       | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')      | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')          | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null

$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($($ExecDir + "\app_icon.ico"))

$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = $appName + " v." + $appVersion
$Main_Tool_Icon.Icon = $icon
$Main_Tool_Icon.Visible = $true

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"

$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)

$Main_Tool_Icon.Add_Click({                    
    If ($_.Button -eq [Windows.Forms.MouseButtons]::Right) {
        $Main_Tool_Icon.GetType().GetMethod("ShowContextMenu",[System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($Main_Tool_Icon,$null)
    } else {
        Try {
            $Global:psCmd.Stop() ; $Global:psCmd.Dispose()
        } catch {}
        $Global:timer.Stop()
        if ($Window_main.Opacity -ne 1) {$Window_main.Opacity = 1}
        if ($Global:FadeAllowed) { HideWithFadeDelay }
        $Window_main.Activate()
        If (-Not $Window_main.IsVisible) {
            $Window_main.ShowDialog()
        }
    }
})

$Menu_Exit.add_Click({
    $Global:DBInstances[0].DBMasterKey = $null
    $Global:DBInstances[1].DBMasterKey = $null
    $Global:DBInstances[2].DBMasterKey = $null
    $Global:CheckBoxes[0] = $CheckBox_AlwaysOnTop.IsChecked
    $Global:CheckBoxes[1] = $CheckBox_AutoComplete.IsChecked
    $Global:DBInstances,$Global:CurrentEntries,$Global:CheckBoxes | ConvertTo-Json | Out-File $($ExecDir + "\" + $appName + ".ini")
    $Window_main.OwnedWindows | % {$_.Close()}
    $Window_main.Close()
    [Environment]::Exit(1)
})
## from https://gist.github.com/selvalogesh/37b99e43b932d42b5a9901a33284b4fa

function HideWithFadeDelay {
    $Global:timer.Restart()
    $SyncHash = [hashtable]::Synchronized(@{Window_Main = $Window_Main; Timer = $Global:timer; FadeDelay = $Global:FadeDelay ; ConsoleHost = (Get-Host)})
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "Default"         
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash)
    $Global:psCmd = [PowerShell]::Create().AddScript({
        Start-Sleep -Seconds $($SyncHash.FadeDelay)
        #$SyncHash.ConsoleHost.Ui.WriteLine("***")
        If (-Not $SyncHash.Window_Main.IsMouseOver) {
            $SyncHash.Window_Main.Dispatcher.Invoke([action]{ $SyncHash.Window_Main.Opacity = 0.25 }, "Normal")
        }
        $SyncHash.Window_Main.Dispatcher.Invoke([action]{ $SyncHash.Timer.Stop() }, "Normal")
    })
    $Global:psCmd.Runspace = $newRunspace
    $Global:psCmd.BeginInvoke()
}

$Button_Filter.Add_Click({
    $Global:FadeAllowed = $false
    $Global:CurrentEntries = ArrangeEntries
    $Global:CurrentEntriesCopy = $Global:CurrentEntries
    $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
    $ListView_Selector.SelectedIndex = 0
    $Window_Selector.Owner = $Window_main
    $Window_Selector.Activate() | Out-Null
    $Window_Selector.Focus()
    $Window_Selector.ShowDialog()
})

$Selector_Button_Up.Add_Click({
    if ($ListView_Selector.SelectedIndex -ne 0) {
        $Index = $ListView_Selector.SelectedIndex
        $EntryBuffer = $Global:CurrentEntries[$ListView_Selector.SelectedIndex -1]
        $Global:CurrentEntries[$ListView_Selector.SelectedIndex - 1] = $Global:CurrentEntries[$ListView_Selector.SelectedIndex]
        $Global:CurrentEntries[$ListView_Selector.SelectedIndex] = $EntryBuffer
        $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
        $ListView_Selector.SelectedIndex = $Index - 1
    }
})

$Selector_Button_Down.Add_Click({
    if ($ListView_Selector.SelectedIndex -lt ($ListView_Selector.Items.Count - 1)) {
        $Index = $ListView_Selector.SelectedIndex
        $EntryBuffer = $Global:CurrentEntries[$ListView_Selector.SelectedIndex + 1]
        $Global:CurrentEntries[$ListView_Selector.SelectedIndex + 1] = $Global:CurrentEntries[$ListView_Selector.SelectedIndex]
        $Global:CurrentEntries[$ListView_Selector.SelectedIndex] = $EntryBuffer
        $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
        $ListView_Selector.SelectedIndex = $Index + 1
    }
})

$Selector_Button_Apply.Add_Click({
    $i = 0 ; $OrderNum = 1
    $Global:CurrentEntries | % {
        If ($_.IsVisible) {
            $Global:CurrentEntries[$i].OrderNum = $OrderNum
            $OrderNum += 1
        } else { $Global:CurrentEntries[$i].OrderNum = 0 }
        $i += 1
    }
    $Global:AttributedEntries = $Global:CurrentEntries

    DrawButtons
    
    $Window_Selector.Hide() | Out-Null
    $Global:FadeAllowed = $true
})

$Selector_Button_Cancel.Add_Click({
    $Global:CurrentEntries = $Global:CurrentEntriesCopy
    $Window_Selector.Hide()
    $Global:FadeAllowed = $true
})

$Window_main.Top = ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height) - $Window_main.Height
$Window_main.Left = ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width) - $Window_main.Width

$CheckBox_AutoRun.Add_Checked({ New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName -Value $("cmd /c " + $([char]'"') + "Start /D $ExecDir powershell -WindowStyle hidden -file $ExecDir\" + $appName + ".ps1" + $([char]'"')) })
$CheckBox_AutoRun.Add_UnChecked({ Remove-ItemProperty  -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName })

$Window_main.add_MouseLeftButtonDown({$Window_main.DragMove()})

$Window_main.Add_MouseEnter({
    Try {$Global:psCmd.Stop() ; $Global:psCmd.Dispose()} catch {}
    $Global:timer.Stop()
    if ($Window_main.Opacity -ne 1) {$Window_main.Opacity = 1}
})

$Window_main.Add_MouseLeave({
    if ($Global:FadeAllowed -and ($Window_Main.WindowState -eq "Normal")) {
        HideWithFadeDelay
    }
})

$Window_main.Add_Loaded({
    $Window_main.Title = $appName + " v." + $appVersion

    $CheckBox_AlwaysOnTop.IsChecked = $Global:CheckBoxes[0]
    $CheckBox_AutoComplete.IsChecked = $Global:CheckBoxes[1]
    Try { if (Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName) {$CheckBox_AutoRun.IsChecked = $true} } catch {}

    # Установка режима без фокусировки окна
    Add-Type @"
                using System;
                using System.Runtime.InteropServices;
                public class Window {
                    [DllImport("user32.dll")]
                    public static extern int GetWindowLong(IntPtr hwnd, int index);

                    [DllImport("user32.dll")]
                    public static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);
                }
"@
    $WindowHandle = (Get-Process | ? {(($_.Name -eq "powershell")  -or ($_.Name -eq "powershell_ise") -or ($_.Name -eq "pwsh")) -and ($_.MainWindowTitle -eq $Window_main.Title)}).MainWindowHandle
    [int]$extendedStyle = [Window]::GetWindowLong($WindowHandle, (-20))
    [Window]::SetWindowLong($WindowHandle,-20,0x08000000)
    #https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
})

$Button_Clipboard.add_Click.Invoke({
    (Get-Clipboard  -Raw).ToCharArray() | % { SendKey $_ }
    if ($CheckBox_AutoComplete.IsChecked) {[InputManager.Keyboard]::KeyPress([System.Windows.Forms.Keys]::Enter)}
})

$Button_Hide.add_Click.Invoke({$Window_main.Hide()})

## from https://gist.github.com/selvalogesh/37b99e43b932d42b5a9901a33284b4fa
# Force garbage collection just to start slightly lower RAM usage.
[System.GC]::Collect()

# Create an application context for it to all run within.
# This helps with responsiveness, especially when clicking Exit.
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)
## from https://gist.github.com/selvalogesh/37b99e43b932d42b5a9901a33284b4fa

### SciptEnd