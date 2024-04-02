$appVersion = "1.2.0.0"
$appName = "PassType"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
$ExecDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\')
Add-Type -Path $($ExecDir + "\InputManager.dll")

# Hide powershell console (if not run Powershell ISE)
if ((Get-Process -PID $pid).ProcessName -ne "powershell_ise") { $null = $(Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -name Win32ShowWindowAsync -namespace Win32Functions -PassThru)::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0) }

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

Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class SystemWindowsFunctions {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd, int index);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);
}
"@

[DBInstance[]]$Global:DBInstances = @()
[EntryBrief[]]$Global:AttributedEntries = @([EntryBrief]::new())
[bool[]]$Global:CheckBoxes = @($false,$false)

# Check KeePass Installation presence
[string]$Global:KeePass_Path = "Portable"
[System.Collections.ArrayList]$AllSoftware = @()
$Registry = [microsoft.win32.registrykey]::OpenRemoteBaseKey(‘LocalMachine’,$env:COMPUTERNAME)
$Registry.OpenSubKey(”SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall").GetSubKeyNames() | % {
    $SoftwareKey = $Registry.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\" + $_)
    $AllSoftware += (New-Object -Type PSObject -Prop @{ ‘DisplayName’ = $SoftwareKey.GetValue("DisplayName") ; ‘DisplayVersion’ = $SoftwareKey.GetValue("DisplayVersion") ; ‘Publisher’ = $SoftwareKey.GetValue("Publisher") ; 'InstallLocation' = $SoftwareKey.GetValue("InstallLocation"); ‘InstallDate’ = $SoftwareKey.GetValue("InstallDate")})
}
$Registry.OpenSubKey(”SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall").GetSubKeyNames() | % {
    $SoftwareKey = $Registry.OpenSubKey(”SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\" + $_)
    $AllSoftware += (New-Object -Type PSObject -Prop @{ ‘DisplayName’ = $SoftwareKey.GetValue("DisplayName") ; ‘DisplayVersion’ = $SoftwareKey.GetValue("DisplayVersion") ; ‘Publisher’ = $SoftwareKey.GetValue("Publisher") ; 'InstallLocation' = $SoftwareKey.GetValue("InstallLocation"); ‘InstallDate’ = $SoftwareKey.GetValue("InstallDate")})
}

$KeePassRecord = $AllSoftware | ? {($_.DisplayName -like "KeePass Password Safe*") -and ($_.Publisher -like "Dominik Reichl")}
if ($KeePassRecord) {
    if ($KeePassRecord.Count -gt 1) { # More then one instances of KeePass presents, point to most fresh version
        $KeePassRecord = $KeePassRecord | ? {$_.DisplayVersion -eq @(($KeePassRecord | measure DisplayVersion -Maximum).Maximum)}
    }
    if ((Get-AuthenticodeSignature -FilePath "$(($KeePassRecord).InstallLocation)KeePass.exe").SignerCertificate.Subject -like "E=cert@dominik-reichl.de,*") {
        $Global:KeePass_Path = "$(($KeePassRecord).InstallLocation)KeePass.exe"
    }
}
# Check KeePass Installation presence

Import-Module -Name $($ExecDir + "\poshkeepass")

[xml]$XAMLMainWindow = @"
<Window x:Name="Window_Main"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    Title="PassType" Height="67" Width="130" ResizeMode="CanResize" WindowStyle="None" BorderThickness="0" AllowsTransparency="True" Background="Transparent" Topmost="{Binding ElementName=CheckBox_AlwaysOnTop, Path=IsChecked}" WindowStartupLocation="CenterScreen" Opacity="0">
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
            <CheckBox x:Name="CheckBox_AutoRun" HorizontalAlignment="Left" Margin="55,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="Black">
                <CheckBox.ToolTip>
                    <ToolTip>Autorun</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <CheckBox x:Name="CheckBox_AutoComplete" HorizontalAlignment="Left" Margin="76,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="Black">
                <CheckBox.ToolTip>
                    <ToolTip>Auto Complete</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <Grid x:Name="WindowMain_KPButtons_Grid" Margin="0,29,0,0"/>
            <Button x:Name="Button_Clipboard" Content="Clipboard" Margin="6,0,6,5" VerticalAlignment="Bottom" Background="Transparent" BorderBrush="Black" Height="20" HorizontalAlignment="Stretch"/>
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
        x:Name="Window_Selector" Title="Filter, Order, Refresh" Height="250" Width="650" ResizeMode="CanResize" WindowStyle="None" SnapsToDevicePixels="True" BorderThickness="1" AllowsTransparency="True" Background="White" BorderBrush="{DynamicResource {x:Static SystemColors.ControlDarkBrushKey}}" WindowStartupLocation="CenterScreen" ShowInTaskbar="False">
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
                                    <CheckBox IsChecked="{Binding IsVisible}"/>
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
        <Button x:Name="Selector_Button_KeePass" Content="KeePass" HorizontalAlignment="Right" Margin="0,1,98,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" />
        <Button x:Name="Selector_Button_Apply" Content=" Apply " HorizontalAlignment="Right" Margin="0,1,53,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0"/>
        <Button x:Name="Selector_Button_Cancel" Content=" Cancel " HorizontalAlignment="Right" Margin="0,1,2,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0"/>
        <Button x:Name="Selector_Button_Up" Content="▲" Background="White" HorizontalAlignment="Left" Height="18" VerticalAlignment="Top" BorderThickness="0" Width="18" Margin="6,1,0,0" Padding="1,-4,1,1"/>
        <Button x:Name="Selector_Button_Down" Content="▼" Background="White" HorizontalAlignment="Left" Height="18" VerticalAlignment="Top" BorderThickness="0" Width="18" Margin="25,0,0,0" Padding="1,4,1,1"/>
        <Label Content="order" HorizontalAlignment="Left" Margin="44,-3,0,0" VerticalAlignment="Top"/>
        <Label Content="search" HorizontalAlignment="Left" Margin="208,-3,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="Selector_Textbox_Search" Text="" HorizontalAlignment="Center" Margin="0,2,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="144" TabIndex="0"/>
        <Button x:Name="Selector_Button_ClearSearchString" HorizontalAlignment="Right" Margin="0,4,252,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Height="16" Width="16" Visibility="Hidden">
            <TextBlock  Text="X" Margin="-1,-3,0,0" RenderTransformOrigin="0.5,0.5" Width="8">
                <TextBlock.RenderTransform>
                    <TransformGroup>
                        <ScaleTransform ScaleY="1" ScaleX="1.65"/>
                    </TransformGroup>
                </TextBlock.RenderTransform>
            </TextBlock>
        </Button>
    </Grid>
</Window>
"@

Function WindowMain_FadeAnimation {
    Param ($From,$To,$DurationSec)
    $Window_main.BeginAnimation([System.Windows.Window]::OpacityProperty, $([System.Windows.Media.Animation.DoubleAnimation]::new($From,$To,$(New-TimeSpan -Seconds $DurationSec))))
}

# Read config file if it exists
Try {
    $AppSettings = Get-Content -Path $($ExecDir + "\" + $appName + ".ini") -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
    $Global:DBInstances = $AppSettings[0].value
    $Global:AttributedEntries = $AppSettings[1].value
    $Global:CheckBoxes = $AppSettings[2].value
    If ($Global:KeePass_Path -eq "Portable") {$Global:KeePass_Path = $AppSettings[3]}
} catch {
    if ((Get-Item -Path $($ExecDir + "\" + $appName + ".ini") -ErrorAction SilentlyContinue)) {[System.Windows.MessageBox]::Show("Exception occured while reading content of config file. Please check it's content or delete config file and relaunch application.")}
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

    $Window_PassType_Entrance.add_MouseLeftButtonDown({$Window_PassType_Entrance.DragMove()})

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
    $Window_PassType_Entrance.ShowDialog() | Out-Null
}

PassType_Entrance

# Common variables, objects
$Global:Delay = 20
$InitialWindowHeight = $Window_main.Height
$Global:FadeAllowed = $true

Function SaveConfiguration {
    $Global:CheckBoxes[0] = $CheckBox_AlwaysOnTop.IsChecked
    $Global:CheckBoxes[1] = $CheckBox_AutoComplete.IsChecked
    [DBInstance[]]$DBInstancesOut = @()
    $Global:DBInstances | % {
        [DBInstance]$TempItem = New-Object -TypeName DBInstance
        $TempItem.DBPath = $_.DBPath
        $TempItem.DBKeyPath = $_.DBKeyPath
        $TempItem.Include = $_.Include
        
        $DBInstancesOut += $TempItem
    }
    $DBInstancesOut | % {$_.DBMasterKey = $null}
    If (-Not $Global:CurrentEntries) { [EntryBrief[]]$Global:CurrentEntries = @() }
    $DBInstancesOut,$Global:CurrentEntries,$Global:CheckBoxes,$Global:KeePass_Path | ConvertTo-Json | Out-File $($ExecDir + "\" + $appName + ".ini")    
}

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
        '[A-Z]$' { SHIFT_KEY $KEY }
        '[a-z]$' { [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY) }
        '^[0-9]' { [InputManager.Keyboard]::KeyDown([System.Windows.Forms.Keys]::("D"+$KEY)) ; [InputManager.Keyboard]::KeyUp([System.Windows.Forms.Keys]::("D" + $KEY)) }
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
                "F1" {SINGLE_KEY "F1"}
                "F2" {SINGLE_KEY "F2"}
                "F3" {SINGLE_KEY "F3"}
                "F4" {SINGLE_KEY "F4"}
                "F5" {SINGLE_KEY "F5"}
                "F6" {SINGLE_KEY "F6"}
                "F7" {SINGLE_KEY "F7"}
                "F8" {SINGLE_KEY "F8"}
                "F9" {SINGLE_KEY "F9"}
                "F10" {SINGLE_KEY "F10"}
                DEFAULT {}
            }
        }
    }
    Start-Sleep -Milliseconds $Global:Delay
}

Function Send_Credentials {
    param(
        [string]$uuid,
        [bool]$Ctrl,
        [bool]$Shift
    )

    $Global:DBInstances | ? {$_.Include} | % {
        $DatabasePath = $_.DBPath
        $TryGetEntry = Get-KeePassEntry -MasterKey $_.DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $DatabasePath}).Name)  | ? {$($_.uuid.Tostring()) -eq $uuid}
        If ($TryGetEntry) {$Entry = $TryGetEntry}
    }

    ## Start-sleep -Milliseconds 100

    # Type entry name, TAB and password
    If (-Not $Shift) {
        if (-Not $Ctrl) { # type entry if not Ctrl pressed
            $Entry.UserName.ToCharArray() | % { SendKey $_ }
            Start-sleep -Milliseconds 100
            [InputManager.Keyboard]::KeyPress([System.Windows.Forms.Keys]::Tab)
        }
        # Waiting for the user to release the Ctrl button
        Start-sleep -Milliseconds 400

        $(([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Entry.Password)))).ToCharArray() | % { SendKey $_ }
    }

    if ($Shift -and (-Not $Ctrl)) { $Entry.UserName.ToCharArray() | % { SendKey $_ } } # Type only entry Name

    if ($Shift -and $Ctrl) { Start $Entry.URL } # entry URL open

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
    $ToolTipText = "with Ctrl - send Password$([System.Environment]::NewLine)with Shift - send Username$([System.Environment]::NewLine)with Ctrl+Shift - open URL"
    $EntriesSorted | % {
        $Window_main.Height += 20
        $Button = [System.Windows.Controls.Button]::new()
        $Button.Name = "Button_" + $_.uuid
        $Button.Content = $_.Name
        $Button.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
        $Button.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
        $Button.Height = 20 
        $Button.Margin = "6,$( [string]($i * ($Button.Height - 1)) ),6,5"
        $Button.Background = "Transparent"
        $Button.ToolTip = $ToolTipText
        $Button.Add_Click({
            #[System.Windows.Forms.InputLanguage]::CurrentInputLanguage = [System.Windows.Forms.InputLanguage]::InstalledInputLanguages | ? { $_.Culture -eq 'en-US' }
            Send_Credentials $($This.Name.Substring(7)) $(([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl)) -or ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightCtrl))) $(([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftShift)))
        })
        $WindowMain_KPButtons_Grid.Children.Add($Button) | Out-Null
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
    $ActiveWindowHandle = [SystemWindowsFunctions]::GetForegroundWindow()
    # Set NotifyIcon never getting focus - ref: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga, https://learn.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
    [int]$extendedStyle = [SystemWindowsFunctions]::GetWindowLong($ActiveWindowHandle, (-20))
    [SystemWindowsFunctions]::SetWindowLong($ActiveWindowHandle,-20,0x08000000)
    If ($_.Button -eq [Windows.Forms.MouseButtons]::Right) {
        $Main_Tool_Icon.GetType().GetMethod("ShowContextMenu",[System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($Main_Tool_Icon,$null)
    } else {
        WindowMain_FadeAnimation -From -1 -To 2 -DurationSec 0.6
        $Global:FadeAllowed = $true
    }
})

$Menu_Exit.add_Click({
    SaveConfiguration
    $Global:DBInstances | % {$_.DBMasterKey = $null}
    $Window_main.OwnedWindows | % {$_.Close()}
    $Window_main.Close()
    [Environment]::Exit(1)
})

$Button_Filter.Add_Click({
    $Global:FadeAllowed = $false
    $Global:CurrentEntries = ArrangeEntries
    $Global:CurrentEntriesCopy = $Global:CurrentEntries

    $Reader=(New-Object System.Xml.XmlNodeReader $XAMLSelectorWindow)
    try { $Window_Selector = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
    $XAMLSelectorWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_Selector.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

    $Window_Selector.add_MouseLeftButtonDown({$Window_Selector.DragMove()})

    $Selector_Textbox_Search.Add_TextChanged({
        if ($this.Text.Length -ge 2) {
            [EntryBrief[]]$EntriesMatched = $Global:CurrentEntries.Where({$_.Name -like "*$($this.Text)*"})
            $ListView_Selector.ItemsSource = @($EntriesMatched)
        } else {
            If ($this.Text.Length -eq 0) { $Selector_Button_ClearSearchString.Visibility = [System.Windows.Visibility]::Hidden } else { $Selector_Button_ClearSearchString.Visibility = [System.Windows.Visibility]::Visible }
            $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
        }
    })

    $Selector_Button_ClearSearchString.Add_Click({
        $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
        $Selector_Textbox_Search.Text = ""
        $Selector_Textbox_Search.Focus()
    })

    $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
    $ListView_Selector.SelectedIndex = 0

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
        SaveConfiguration
        $Window_Selector.Close() | Out-Null
        $Global:FadeAllowed = $true
    })

    $Selector_Button_Cancel.Add_Click({
        $Global:CurrentEntries = $Global:CurrentEntriesCopy
        $Window_Selector.Close()
        $Global:FadeAllowed = $true
    })

    $Window_Selector.Add_Loaded({
        $Window_Selector.Activate() | Out-Null
        $Selector_Textbox_Search.Focus() | Out-Null
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

    $Selector_Button_KeePass.Add_Click({
        if ($Global:KeePass_Path -eq "Portable") {
            $objForm = New-Object System.Windows.Forms.OpenFileDialog
            $objForm.Title = "Select location of keepass.exe"
            $objForm.InitialDirectory = $env:SystemDrive
            $objForm.Filter = "keepass.exe|keepass.exe"
            $objForm.Multiselect = $false
            $objForm.ShowDialog()
            if ($objForm.FileName) {$Global:KeePass_Path = $objForm.FileName}
        }

        If ($Global:KeePass_Path -ne "Portable") {
            If (Get-Process -Name KeePass -ea 0) {
                Start-Process -FilePath $Global:KeePass_Path -ArgumentList @("-exit-all")
                While (Get-Process -Name KeePass -ea 0) {Start-Sleep -Milliseconds 200}
            }
        
            # Select User XML config for launch Keepass
            if (Test-Path @($env:APPDATA + "\KeePass\KeePass.config.xml") -ea 0) { # use user's config and modify some nodes
                [XML]$XML_config = Get-Content -Path @($env:APPDATA + "\KeePass\KeePass.config.xml")
                Try {
                    If (-Not $XML_config.Configuration.Application.Start.MinimizedAndLocked) {
                        $XML_config.Configuration.Application.Start.AppendChild($XML_config.CreateElement("MinimizedAndLocked")) | Out-Null
                    }$XML_config.Configuration.Application.Start.MinimizedAndLocked = "false"
                } catch {}

                Try {
                    If (-Not $XML_config.Configuration.Application.Start.OpenLastFile) {
                        $XML_config.Configuration.Application.Start.AppendChild($XML_config.CreateElement("OpenLastFile")) | Out-Null
                    }$XML_config.Configuration.Application.Start.OpenLastFile = "false"
                } catch {}
                $XML_config.Save(@($env:TEMP + "\KeePass.config.xml"))
                $XMLPath = @($env:TEMP + "\KeePass.config.xml")
            } else {$XMLPath = "$ExecDir\KeePass.config.xml"} # Use config file from script folder
            & $Global:KeePass_Path "-cfg-local:$XMLPath"
            While (-Not (Get-Process -Name KeePass -ea 0)) {Start-Sleep -Milliseconds 200}

            $Global:DBInstances | ? {$_.Include} | % {
                Try {
                  Start-Sleep -Seconds 1
                  $KeePassProcess = New-Object System.Diagnostics.Process
                  $KeePassProcess.StartInfo.FileName = $Global:KeePass_Path
                  if ($_.DBKeyPath) {
                    $KeePassProcess.StartInfo.Arguments = $($_.DBPath),"-keyfile:$($_.DBKeyPath)","-pw-stdin"
                  } else {
                    $KeePassProcess.StartInfo.Arguments = $($_.DBPath),"-pw-stdin"
                  }
                  $KeePassProcess.StartInfo.UseShellExecute = $false
                  $KeePassProcess.StartInfo.RedirectStandardInput = $true
                  $KeePassProcess.Start()

                  $StdIn = $KeePassProcess.StandardInput
                  $StdIn.WriteLine($([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($_.DBMasterKey))))
                  While (-Not $KeePassProcess.Responding) {Start-Sleep -Milliseconds 100}
                } Finally {
                      if($StdIn) { $StdIn.Close() }
                }
            }
        }
    })

    $Window_Selector.Owner = $Window_main
    $Window_Selector.ShowDialog()
})

$Window_main.Top = ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height) - $Window_main.Height
$Window_main.Left = ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width) - $Window_main.Width

$CheckBox_AutoRun.Add_Checked({ New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName -Value $("cmd /c " + $([char]'"') + "Start /D $ExecDir powershell -WindowStyle hidden -file $ExecDir\" + $appName + ".ps1" + $([char]'"')) })
$CheckBox_AutoRun.Add_UnChecked({ Remove-ItemProperty  -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName })

$CheckBox_AlwaysOnTop.Add_Checked({ SaveConfiguration }) ; $CheckBox_AlwaysOnTop.Add_UnChecked({ SaveConfiguration })
$CheckBox_AutoComplete.Add_Checked({ SaveConfiguration }) ; $CheckBox_AlwaysOnTop.Add_UnChecked({ SaveConfiguration })

$Window_main.add_MouseLeftButtonDown({$Window_main.DragMove()})

$Window_main.Add_MouseEnter({
    If ($Window_main.Opacity -ne 1) { WindowMain_FadeAnimation -From $Window_main.Opacity -to 2 -DurationSec 0.6 }
})

$Window_main.Add_MouseLeave({
    If ($Global:FadeAllowed) {
        $Animation = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $Animation.From = $Window_main.Opacity
        $Animation.To = 0.25
        $Animation.BeginTime = New-TimeSpan -Seconds 15
        $Animation.Duration = New-TimeSpan -Seconds 0.6
        $Window_main.BeginAnimation([System.Windows.Window]::OpacityProperty, $Animation)
    }
})

$Button_Clipboard.add_Click.Invoke({
    If ((Get-Clipboard -Raw) -match '[fF]\d{1,2};') { # Typing functional keys from clipboard, format Fn; - examples F2;F8;F10;
        [regex]::Matches($(Get-Clipboard -Raw),"[fF]\d{1,2};") | % {
            SendKey $(($_.Value).TrimEnd(";"))
        }
    } else {# Typing non Fn keys
        (Get-Clipboard -Raw).ToCharArray() | % { SendKey $_ }
    }
    if ($CheckBox_AutoComplete.IsChecked) {[InputManager.Keyboard]::KeyPress([System.Windows.Forms.Keys]::Enter)}
})

$Button_Hide.add_Click.Invoke({$Global:FadeAllowed = $False ; WindowMain_FadeAnimation -From 2 -to -1 -DurationSec 0.6})

$Window_main.Add_Loaded({
    $Window_main.Title = $appName + " v." + $appVersion
    $CheckBox_AlwaysOnTop.IsChecked = $Global:CheckBoxes[0]
    $CheckBox_AutoComplete.IsChecked = $Global:CheckBoxes[1]
    Try { if (Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName) {$CheckBox_AutoRun.IsChecked = $true} } catch {}

    $WindowHandle = (Get-Process | ? {(($_.Name -eq "powershell")  -or ($_.Name -eq "powershell_ise") -or ($_.Name -eq "pwsh")) -and ($_.MainWindowTitle -eq $Window_main.Title)}).MainWindowHandle
    
    # Set Window never getting focus after activation ref: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga, https://learn.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
    [int]$extendedStyle = [SystemWindowsFunctions]::GetWindowLong($WindowHandle, (-20))
    [SystemWindowsFunctions]::SetWindowLong($WindowHandle,-20,0x08000000)

    WindowMain_FadeAnimation -From 0 -to 1 -DurationSec 0.6
})

$Window_main.Activate()
$Window_main.ShowDialog()

[System.GC]::Collect()
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)