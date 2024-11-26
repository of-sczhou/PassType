$appVersion = "1.2.1.0"
$appName = "PassType"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

if (-Not (Get-Variable psISE -ea 0)) {
    $null = $(Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -name Win32ShowWindowAsync -namespace Win32Functions -PassThru)::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0) # Hide powershell console
    $ExecDir = $MyInvocation.MyCommand.Path.Substring(0,$($MyInvocation.MyCommand.Path.LastIndexOf("\")))
} else {
    $ExecDir = $psISE.CurrentFile.FullPath.Substring(0,$($psISE.CurrentFile.FullPath.LastIndexOf("\")))
}

Class DBInstance {
    [string]$DBName
    [string]$DBPath
    [string]$DBKeyPath
    [SecureString]$DBMasterKey
}

Class EntryBrief {
    [string]$Uuid
    [string]$Name
    [string]$DBPath
    [string]$DBName
    [Int32]$OrderNum
    [bool]$IsVisible
}

Add-Type @"
  using System;
  using System.Drawing;
  using System.Runtime.InteropServices;

  public class SystemWindowsFunctions {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd, int index);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    public static void BringToFront(IntPtr handle)
    {
        SetForegroundWindow(handle);
    }
}
"@

# from https://www.codeproject.com/Articles/117657/InputManager-library-Track-user-input-and-simulate
Add-Type @"
    Imports System
    Imports System.Windows.Forms
    Imports System.Runtime
    Imports System.Runtime.InteropServices
    Imports System.Threading
    ''' <summary>
    ''' Provide methods to send keyboard input that also works in DirectX games.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Keyboard
    #Region "API Declaring"
    #Region "SendInput"
        Private Declare Function SendInput Lib "user32.dll" (ByVal cInputs As Integer, ByRef pInputs As INPUT, ByVal cbSize As Integer) As Integer
        Private Structure INPUT
            Dim dwType As Integer
            Dim mkhi As MOUSEKEYBDHARDWAREINPUT
        End Structure

        Private Structure KEYBDINPUT
            Public wVk As Short
            Public wScan As Short
            Public dwFlags As Integer
            Public time As Integer
            Public dwExtraInfo As IntPtr
        End Structure

        Private Structure HARDWAREINPUT
            Public uMsg As Integer
            Public wParamL As Short
            Public wParamH As Short
        End Structure

        <StructLayout(LayoutKind.Explicit)>
        Private Structure MOUSEKEYBDHARDWAREINPUT
            <FieldOffset(0)> Public mi As MOUSEINPUT
            <FieldOffset(0)> Public ki As KEYBDINPUT
            <FieldOffset(0)> Public hi As HARDWAREINPUT
        End Structure

        Private Structure MOUSEINPUT
            Public dx As Integer
            Public dy As Integer
            Public mouseData As Integer
            Public dwFlags As Integer
            Public time As Integer
            Public dwExtraInfo As IntPtr
        End Structure

        Const INPUT_MOUSE As UInt32 = 0
        Const INPUT_KEYBOARD As Integer = 1
        Const INPUT_HARDWARE As Integer = 2
        Const KEYEVENTF_EXTENDEDKEY As UInt32 = &H1
        Const KEYEVENTF_KEYUP As UInt32 = &H2
        Const KEYEVENTF_UNICODE As UInt32 = &H4
        Const KEYEVENTF_SCANCODE As UInt32 = &H8
        Const XBUTTON1 As UInt32 = &H1
        Const XBUTTON2 As UInt32 = &H2
        Const MOUSEEVENTF_MOVE As UInt32 = &H1
        Const MOUSEEVENTF_LEFTDOWN As UInt32 = &H2
        Const MOUSEEVENTF_LEFTUP As UInt32 = &H4
        Const MOUSEEVENTF_RIGHTDOWN As UInt32 = &H8
        Const MOUSEEVENTF_RIGHTUP As UInt32 = &H10
        Const MOUSEEVENTF_MIDDLEDOWN As UInt32 = &H20
        Const MOUSEEVENTF_MIDDLEUP As UInt32 = &H40
        Const MOUSEEVENTF_XDOWN As UInt32 = &H80
        Const MOUSEEVENTF_XUP As UInt32 = &H100
        Const MOUSEEVENTF_WHEEL As UInt32 = &H800
        Const MOUSEEVENTF_VIRTUALDESK As UInt32 = &H4000
        Const MOUSEEVENTF_ABSOLUTE As UInt32 = &H8000
    #End Region
        Private Declare Auto Function MapVirtualKey Lib "user32.dll" (ByVal uCode As UInt32, ByVal uMapType As MapVirtualKeyMapTypes) As UInt32
        Private Declare Auto Function MapVirtualKeyEx Lib "user32.dll" (ByVal uCode As UInt32, ByVal uMapType As MapVirtualKeyMapTypes, ByVal dwhkl As IntPtr) As UInt32
        Private Declare Auto Function GetKeyboardLayout Lib "user32.dll" (ByVal idThread As UInteger) As IntPtr
        ''' <summary>The set of valid MapTypes used in MapVirtualKey
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum MapVirtualKeyMapTypes As UInt32
            ''' <summary>uCode is a virtual-key code and is translated into a scan code.
            ''' If it is a virtual-key code that does not distinguish between left- and
            ''' right-hand keys, the left-hand scan code is returned.
            ''' If there is no translation, the function returns 0.
            ''' </summary>
            ''' <remarks></remarks>
            MAPVK_VK_TO_VSC = &H0

            ''' <summary>uCode is a scan code and is translated into a virtual-key code that
            ''' does not distinguish between left- and right-hand keys. If there is no
            ''' translation, the function returns 0.
            ''' </summary>
            ''' <remarks></remarks>
            MAPVK_VSC_TO_VK = &H1

            ''' <summary>uCode is a virtual-key code and is translated into an unshifted
            ''' character value in the low-order word of the return value. Dead keys (diacritics)
            ''' are indicated by setting the top bit of the return value. If there is no
            ''' translation, the function returns 0.
            ''' </summary>
            ''' <remarks></remarks>
            MAPVK_VK_TO_CHAR = &H2

            ''' <summary>Windows NT/2000/XP: uCode is a scan code and is translated into a
            ''' virtual-key code that distinguishes between left- and right-hand keys. If
            ''' there is no translation, the function returns 0.
            ''' </summary>
            ''' <remarks></remarks>
            MAPVK_VSC_TO_VK_EX = &H3

            ''' <summary>Not currently documented
            ''' </summary>
            ''' <remarks></remarks>
            MAPVK_VK_TO_VSC_EX = &H4
        End Enum
    #End Region
        Private Shared Function GetScanKey(ByVal VKey As UInteger) As ScanKey
            Dim ScanCode As UInteger = MapVirtualKey(VKey, MapVirtualKeyMapTypes.MAPVK_VK_TO_VSC)
            Dim Extended As Boolean = (VKey = Keys.RMenu Or VKey = Keys.RControlKey Or VKey = Keys.Left Or VKey = Keys.Right Or VKey = Keys.Up Or VKey = Keys.Down Or VKey = Keys.Home Or VKey = Keys.Delete Or VKey = Keys.PageUp Or VKey = Keys.PageDown Or VKey = Keys.End Or VKey = Keys.Insert Or VKey = Keys.NumLock Or VKey = Keys.PrintScreen Or VKey = Keys.Divide)
            Return New ScanKey(ScanCode, Extended)
        End Function
        Private Structure ScanKey
            Dim ScanCode As UInteger
            Dim Extended As Boolean
            Public Sub New(ByVal sCode As UInteger, Optional ByVal ex As Boolean = False)
                ScanCode = sCode
                Extended = ex
            End Sub
        End Structure
        ''' <summary>
        ''' Sends shortcut keys (key down and up) signals.
        ''' </summary>
        ''' <param name="kCode">The array of keys to send as a shortcut.</param>
        ''' <param name="Delay">The delay in milliseconds between the key down and up events.</param>
        ''' <remarks></remarks>
        Public Shared Sub ShortcutKeys(ByVal kCode() As Keys, Optional ByVal Delay As Integer = 0)
            Dim KeysPress As New KeyPressStruct(kCode, Delay)
            Dim t As New Thread(New ParameterizedThreadStart(AddressOf KeyPressThread))
            t.Start(KeysPress)
        End Sub
        ''' <summary>
        ''' Sends a key down signal.
        ''' </summary>
        ''' <param name="kCode">The virtual keycode to send.</param>
        ''' <remarks></remarks>
        Public Shared Sub KeyDown(ByVal kCode As Keys)
            Dim sKey As ScanKey = GetScanKey(kCode)
            Dim input As New INPUT()
            input.dwType = INPUT_KEYBOARD
            input.mkhi.ki = New KEYBDINPUT()
            input.mkhi.ki.wScan = sKey.ScanCode
            input.mkhi.ki.dwExtraInfo = IntPtr.Zero
            input.mkhi.ki.dwFlags = KEYEVENTF_SCANCODE Or Microsoft.VisualBasic.Interaction.IIf(sKey.Extended, KEYEVENTF_EXTENDEDKEY, Nothing)
            Dim cbSize As Integer = Marshal.SizeOf(GetType(INPUT))
            SendInput(1, input, cbSize)
        End Sub
        ''' <summary>
        ''' Sends a key up signal.
        ''' </summary>
        ''' <param name="kCode">The virtual keycode to send.</param>
        ''' <remarks></remarks>
        Public Shared Sub KeyUp(ByVal kCode As Keys)
            Dim sKey As ScanKey = GetScanKey(kCode)
            Dim input As New INPUT()
            input.dwType = INPUT_KEYBOARD
            input.mkhi.ki = New KEYBDINPUT()
            input.mkhi.ki.wScan = sKey.ScanCode
            input.mkhi.ki.dwExtraInfo = IntPtr.Zero
            input.mkhi.ki.dwFlags = KEYEVENTF_SCANCODE Or KEYEVENTF_KEYUP Or Microsoft.VisualBasic.Interaction.IIf(sKey.Extended, KEYEVENTF_EXTENDEDKEY, Nothing)
            Dim cbSize As Integer = Marshal.SizeOf(GetType(INPUT))
            SendInput(1, input, cbSize)
        End Sub
        ''' <summary>
        ''' Sends a key press signal (key down and up).
        ''' </summary>
        ''' <param name="kCode">The virtual keycode to send.</param>
        ''' <param name="Delay">The delay to set between the key down and up commands.</param>
        ''' <remarks></remarks>
        Public Shared Sub KeyPress(ByVal kCode As Keys, Optional ByVal Delay As Integer = 0)
            Dim SendKeys() As Keys = {kCode}
            Dim KeysPress As New KeyPressStruct(SendKeys, Delay)
            Dim t As New Thread(New ParameterizedThreadStart(AddressOf KeyPressThread))
            t.Start(KeysPress)
        End Sub
        Private Shared Sub KeyPressThread(ByVal KeysP As KeyPressStruct)
            For Each k As Keys In KeysP.Keys
                KeyDown(k)
            Next
            If KeysP.Delay > 0 Then Thread.Sleep(KeysP.Delay)
            For Each k As Keys In KeysP.Keys
                KeyUp(k)
            Next
        End Sub
        Private Structure KeyPressStruct
            Dim Keys() As Keys
            Dim Delay As Integer
            Public Sub New(ByVal KeysToPress() As Keys, Optional ByVal DelayTime As Integer = 0)
                Keys = KeysToPress
                Delay = DelayTime
            End Sub
        End Structure
    End Class
    ''' <summary>
    ''' Provides methods to send keyboard input. The keys are being sent virtually and cannot be used with DirectX.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class VirtualKeyboard
    #Region "API Declaring"
        <DllImport("user32.dll", CallingConvention:=CallingConvention.StdCall,
               CharSet:=CharSet.Unicode, EntryPoint:="keybd_event",
               ExactSpelling:=True, SetLastError:=True)>
        Public Shared Function keybd_event(ByVal bVk As Int32, ByVal bScan As Int32,
                                  ByVal dwFlags As Int32, ByVal dwExtraInfo As Int32) As Boolean
        End Function
        Const KEYEVENTF_EXTENDEDKEY = &H1
        Const KEYEVENTF_KEYUP = &H2
    #End Region
        ''' <summary>
        ''' Sends shortcut keys (key down and up) signals.
        ''' </summary>
        ''' <param name="kCode">The array of keys to send as a shortcut.</param>
        ''' <param name="Delay">The delay in milliseconds between the key down and up events.</param>
        ''' <remarks></remarks>
        Public Shared Sub ShortcutKeys(ByVal kCode() As Keys, Optional ByVal Delay As Integer = 0)
            Dim KeyPress As New KeyPressStruct(kCode, Delay)
            Dim t As New Thread(New ParameterizedThreadStart(AddressOf KeyPressThread))
            t.Start(KeyPress)
        End Sub
        ''' <summary>
        ''' Sends a key down signal.
        ''' </summary>
        ''' <param name="kCode">The virtual keycode to send.</param>
        ''' <remarks></remarks>
        Public Shared Sub KeyDown(ByVal kCode As Keys)
            keybd_event(kCode, 0, 0, 0)
        End Sub
        ''' <summary>
        ''' Sends a key up signal.
        ''' </summary>
        ''' <param name="kCode">The virtual keycode to send.</param>
        ''' <remarks></remarks>
        Public Shared Sub KeyUp(ByVal kCode As Keys)
            keybd_event(kCode, 0, KEYEVENTF_KEYUP, 0)
        End Sub
        ''' <summary>
        ''' Sends a key press signal (key down and up).
        ''' </summary>
        ''' <param name="kCode">The virtual key code to send.</param>
        ''' <param name="Delay">The delay to set between the key down and up commands.</param>
        ''' <remarks></remarks>
        Public Shared Sub KeyPress(ByVal kCode As Keys, Optional ByVal Delay As Integer = 0)
            Dim SendKeys() As Keys = {kCode}
            Dim KeyPress As New KeyPressStruct(SendKeys, Delay)
            Dim t As New Thread(New ParameterizedThreadStart(AddressOf KeyPressThread))
            t.Start(KeyPress)
        End Sub
        Private Shared Sub KeyPressThread(ByVal KeysP As KeyPressStruct)
            For Each k As Keys In KeysP.Keys
                KeyDown(k)
            Next
            If KeysP.Delay > 0 Then Thread.Sleep(KeysP.Delay)
            For Each k As Keys In KeysP.Keys
                KeyUp(k)
            Next
        End Sub
        Private Structure KeyPressStruct
            Dim Keys() As Keys
            Dim Delay As Integer
            Public Sub New(ByVal KeysToPress() As Keys, Optional ByVal DelayTime As Integer = 0)
                Keys = KeysToPress
                Delay = DelayTime
            End Sub
        End Structure
    End Class
"@ -Language VisualBasic -ReferencedAssemblies System.Windows.Forms

[DBInstance[]]$Global:DBInstances = @()
[EntryBrief[]]$Global:AttributedEntries = @([EntryBrief]::new())
[bool[]]$Global:CheckBoxes = @($false,$false)

$Global:PreviousWindowHandle = $null # Handle of window wich is active when we click application NotifyIcon in system tray
$Global:NotifyIconMouseOverOnce = $false
$Global:NotifyIconMouseC1ickonce = $false
$Global:WindowMainHandle = $null

0..10 | % {
    New-Variable -Name "PasswordBox$_" -Value $(new-object System.Windows.Controls.PasswordBox) -Scope Global -Force
}

# Check KeePass Installation presence
[string]$Global:KeePass_Path = "Undefined"
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
    $Global:KeePass_Path = "$(($KeePassRecord).InstallLocation)KeePass.exe"
}
# Check KeePass Installation presence

Import-Module -Name $($ExecDir + "\poshkeepass")

[xml]$XAMLMainWindow = @"
<Window x:Name="Window_Main"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    Title="PassType" Height="67" Width="130" ResizeMode="CanResize" WindowStyle="None" BorderThickness="0" AllowsTransparency="True" Background="Transparent" WindowStartupLocation="CenterScreen" Opacity="0" Topmost="True">
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="0" ResizeBorderThickness="5"/>
    </WindowChrome.WindowChrome>
    <Border x:Name="WindowMain_Border" CornerRadius="7" BorderBrush="#FF263238" BorderThickness="1" Background="#FF9FC4D6">
        <Grid x:Name="WindowMain_Grid">
            <Button x:Name="Button_Hide" Background="Transparent" HorizontalAlignment="Right" Height="20" Width="20" VerticalAlignment="Top" BorderThickness="0,0,0,2" BorderBrush="#FF263238" Margin="0,4,4,0"/>
            <Button x:Name="Button_Filter" Content="..." Background="Transparent" HorizontalAlignment="Left" Height="20" Width="20" VerticalAlignment="Top" BorderThickness="0" BorderBrush="Black" Margin="7,4,0,0" FontWeight="Bold" FontSize="16" Foreground="#FF263238">
                <Button.ToolTip>
                    <ToolTip>Filter, Order, Refresh</ToolTip>
                </Button.ToolTip>
            </Button>
            <CheckBox x:Name="CheckBox_TypeOrClip" HorizontalAlignment="Left" Margin="34,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="#FF263238">
                <CheckBox.ToolTip>
                    <ToolTip>
                        <TextBlock>  
                            Type characters like keyboard if checkbox is checked
                        <LineBreak/>
                            otherwise paste via clipboard (does not work in environments
                        <LineBreak/>
                            where clipboard sharing is not possible or prohibited)
                        </TextBlock>
                    </ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <CheckBox x:Name="CheckBox_AutoRun" HorizontalAlignment="Left" Margin="55,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="#FF263238">
                <CheckBox.ToolTip>
                    <ToolTip>Autorun</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <CheckBox x:Name="CheckBox_AutoComplete" HorizontalAlignment="Left" Margin="76,6,0,0" VerticalAlignment="Top" Background="Transparent" BorderBrush="#FF263238">
                <CheckBox.ToolTip>
                    <ToolTip>Auto Complete</ToolTip>
                </CheckBox.ToolTip>
            </CheckBox>
            <Grid x:Name="WindowMain_KPButtons_Grid" Margin="0,29,0,0"/>
            <Rectangle Fill="#FFBECDD4" Height="12" Margin="6,0,6,28" VerticalAlignment="Bottom" Opacity="0.3" Width="117"/>
            <Button x:Name="Button_Clipboard" Content="Clipboard" Margin="6,0,6,5" VerticalAlignment="Bottom" Background="Transparent" BorderBrush="#FF263238" BorderThickness="1" Height="20" HorizontalAlignment="Stretch" Foreground="#FF263238"/>
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
        x:Name="Window_Selector" Title="Options" Height="250" Width="508" ResizeMode="NoResize" WindowStyle="None" SnapsToDevicePixels="True" BorderThickness="1" AllowsTransparency="True" Background="White" BorderBrush="{DynamicResource {x:Static SystemColors.ControlDarkBrushKey}}" WindowStartupLocation="CenterOwner" ShowInTaskbar="False">
<WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="0" ResizeBorderThickness="5"/>
</WindowChrome.WindowChrome>
<Grid>
    <ListView BorderThickness="0" x:Name="ListView_Selector" SelectionMode="Single" Margin="0,22,0,0">
        <ListView.View>
            <GridView x:Name="GridView_Selector">
                <GridViewColumn Header="Visible" Width="NaN">
                    <GridViewColumn.CellTemplate>
                        <DataTemplate>
                            <Grid HorizontalAlignment="Stretch">
                                <CheckBox IsChecked="{Binding IsVisible}"/>
                            </Grid>
                        </DataTemplate>
                    </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Name" Width="NaN" DisplayMemberBinding="{Binding Name}"/>
                <GridViewColumn Header="Database" Width="NaN"  DisplayMemberBinding ="{Binding DBName}"/>
            </GridView>
        </ListView.View>
    </ListView>
    <Button x:Name="Selector_Button_Sources" Content=" Sources " HorizontalAlignment="Right" Margin="0,2,155,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3" />
    <Button x:Name="Selector_Button_KeePass" Content=" KeePass " HorizontalAlignment="Right" Margin="0,2,98,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3" />
    <Button x:Name="Selector_Button_Apply" Content=" Apply " HorizontalAlignment="Right" Margin="0,2,53,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3"/>
    <Button x:Name="Selector_Button_Cancel" Content=" Cancel " HorizontalAlignment="Right" Margin="0,2,4,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3"/>
    <Button x:Name="Selector_Button_Up" Content="▲" Background="White" HorizontalAlignment="Left" Height="18" VerticalAlignment="Top" BorderThickness="0" Width="18" Margin="6,1,0,0" Padding="1,-4,1,1"/>
    <Button x:Name="Selector_Button_Down" Content="▼" Background="White" HorizontalAlignment="Left" Height="18" VerticalAlignment="Top" BorderThickness="0" Width="18" Margin="25,0,0,0" Padding="1,4,1,1"/>
    <Label Content="order" HorizontalAlignment="Left" Margin="44,-3,0,0" VerticalAlignment="Top"/>
    <Label Content="search" HorizontalAlignment="Center" Margin="-250,-3,0,0" VerticalAlignment="Top"/>
    <TextBox x:Name="Selector_Textbox_Search" Text="" HorizontalAlignment="Center" Margin="-60,2,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="144" TabIndex="0"/>
    <Button x:Name="Selector_Button_ClearSearchString" HorizontalAlignment="Center" Margin="0,4,-70,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Height="16" Width="16" Visibility="Hidden">
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

[xml]$XAMLWindow_Sources = @"
        <Window x:Name="Window_Sources"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="Sources" Height="171" Width="650" ResizeMode="CanResize" WindowStyle="None" SnapsToDevicePixels="True" BorderThickness="1" AllowsTransparency="True" Background="White" BorderBrush="{DynamicResource {x:Static SystemColors.ControlDarkBrushKey}}" ShowInTaskbar="False" WindowStartupLocation="CenterOwner">
<WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="0" ResizeBorderThickness="5"/>
</WindowChrome.WindowChrome>
<Grid>
    <Button x:Name="Sources_Button_Save" Content=" Save and restart app " HorizontalAlignment="Right" Margin="0,2,53,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3"/>
    <Button x:Name="Sources_Button_Cancel" Content=" Cancel " HorizontalAlignment="Right" Margin="0,2,4,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3"/>
    <Button x:Name="Sources_Button_Insert" Content=" + " HorizontalAlignment="Center" Margin="-8,-3,4,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3" FontWeight="Bold" FontSize="18"/>
    <Button x:Name="Sources_Button_Remove" Content=" - " HorizontalAlignment="Center" Margin="24,-3,4,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3" FontWeight="Bold" FontSize="18"/>
    <Button x:Name="Sources_Button_Verify" Content=" Verify access " HorizontalAlignment="Left" Margin="2,2,0,0" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Padding="1,0,1,3"/>
    <DataGrid Margin="0,21,0,0" x:Name="Sources_DataGrid" ColumnWidth="*" AutoGenerateColumns="False" CanUserAddRows="False">
        <DataGrid.Columns>
            <DataGridTextColumn Header="Database" Binding="{Binding DBPath}"/>
            <DataGridTextColumn Header="Database Key" Binding="{Binding DBKeyPath}"/>
            <DataGridTextColumn Header="Verified" Binding="{Binding DBAccessVerified}" IsReadOnly="True" Width="50">
                <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock">
                        <Setter Property="HorizontalAlignment" Value="Center" />
                    </Style>
                </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
        </DataGrid.Columns>
    </DataGrid>
</Grid>
</Window>
"@

[xml]$XAMLMaster_Password_Window = @"
    <Window x:Name="Window_Password"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Master Password" Height="72" Width="320" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" WindowStyle="ToolWindow" ScrollViewer.VerticalScrollBarVisibility="Disabled" HorizontalContentAlignment="Center" SnapsToDevicePixels="True">
        <Grid>
            <PasswordBox x:Name="PasswordBox_Window_Password" HorizontalAlignment="Left" Margin="4,0,0,0" VerticalAlignment="Center" Width="240" Height="21" Padding="2,0,2,0"/>
            <Button x:Name="BottonFO_Window_Password" Content=" GO " HorizontalAlignment="Right" Margin="0,0,5,0" VerticalAlignment="Center" Background="{x:Null}" BorderBrush="#FFABADB3"/>
            <Label x:Name="Label_Window_Password" Content="" Height="20" Margin="0,0,38,0" VerticalAlignment="Center" Width="18" HorizontalAlignment="Right" Padding="0,0,0,0" HorizontalContentAlignment="Center" VerticalContentAlignment="Center"/>
        </Grid>
    </Window>
"@

Function WindowMain_FadeAnimation {
    Param ($From,$To,$DurationSec)
    $Window_main.BeginAnimation([System.Windows.Window]::OpacityProperty, $([System.Windows.Media.Animation.DoubleAnimation]::new($From,$To,$(New-TimeSpan -Seconds $DurationSec))))
}

Function Button_Filter_BlinkAnimation {
    Param ([bool]$Animation)
    $DoubleAnimation = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $DoubleAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $DoubleAnimation.AutoReverse = $true
    $DoubleAnimation.From = 0
    $DoubleAnimation.To = 1
    $DoubleAnimation.Duration = New-TimeSpan -Seconds 0.6
    if ($Animation) {
        $Button_Filter.BeginAnimation([System.Windows.Controls.Button]::OpacityProperty, $DoubleAnimation)
    } else {
        $DoubleAnimation.BeginTime = $null
    }
}

Function Button_Filter_BlinkAnimation {
    Param ([bool]$Animation)
    $DoubleAnimation = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $DoubleAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $DoubleAnimation.AutoReverse = $true
    $DoubleAnimation.From = 0
    $DoubleAnimation.To = 1
    $DoubleAnimation.Duration = New-TimeSpan -Seconds 0.6
    if ($Animation) {
        $Button_Filter.BeginAnimation([System.Windows.Controls.Button]::OpacityProperty, $DoubleAnimation)
    } else {
        $DoubleAnimation.BeginTime = $null
    }
}

# Read config file if it exists
Try {
    $AppSettings = Get-Content -Path $($ExecDir + "\" + $appName + ".ini") -Raw -ea 0 | ConvertFrom-Json
    If ($Global:KeePass_Path -eq "Undefined") {$Global:KeePass_Path = $AppSettings[0]}
    $Global:CheckBoxes = $AppSettings[1].value
    $Global:DBInstances = $AppSettings[2].value
    $Global:AttributedEntries = $AppSettings[3].value
} catch {
    #if ((Get-Item -Path $($ExecDir + "\" + $appName + ".ini") -ErrorAction SilentlyContinue)) {[System.Windows.MessageBox]::Show("Exception occured while reading content of config file. Please check it's content or delete config file and relaunch application.")}
}

function PassType_Entrance {
[xml]$XAMLWindow_PassType_Entrance = @"
<Window x:Name="Window_PassType_Entrance"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="PassType Entrance" Height="74" Width="250" ResizeMode="NoResize" ShowInTaskbar="False" Topmost="True" WindowStartupLocation="CenterScreen" WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid >
        <TabControl x:Name="TabControl" Height="NaN" Margin="0,0,0,0" SelectedIndex="0"/>
        <Button x:Name="Button_OK" Content="OK" HorizontalAlignment="Right" Margin="0,0,42,4" VerticalAlignment="Bottom" Height="20" Width="32" BorderBrush="#FFABADB3" IsTabStop="False"/>
        <Button x:Name="Button_Quit" Content="Quit" HorizontalAlignment="Right" Margin="0,0,6,4" VerticalAlignment="Bottom" Width="32" BorderBrush="#FFABADB3" IsTabStop="False"/>
    </Grid>
</Window>
"@

    $Reader=(New-Object System.Xml.XmlNodeReader $XAMLWindow_PassType_Entrance)
    try { $Window_PassType_Entrance = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
    $XAMLWindow_PassType_Entrance.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_PassType_Entrance.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

    $Window_PassType_Entrance.add_MouseLeftButtonDown({$Window_PassType_Entrance.DragMove()})

function VerifyMasterKeys {
    [DBInstance[]]$Global:DBInstances | % {
        $_.DBMasterKey = $(ConvertTo-SecureString $(((Get-Variable -Name "PasswordBox$($Global:DBInstances.IndexOf($_))").Value).Password) -AsPlainText -Force)
    }

    $AuthOK = $true
    $Global:DBInstances | % {
        $DBStr = $_.DBName
        Try {
            Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
            If ($_.DBKeyPath) {
                New-KeePassDatabaseConfiguration -Default -DatabaseProfileName 'TempDatabase' -DatabasePath $_.DBPath -UseMasterKey -KeyPath $_.DBKeyPath
            } else {
                New-KeePassDatabaseConfiguration -Default -DatabaseProfileName 'TempDatabase' -DatabasePath $_.DBPath -UseMasterKey
            }
            $Global:DBInstances[$Global:DBInstances.Indexof($_)].DBName = (Get-KeePassGroup -DatabaseProfileName 'TempDatabase' -MasterKey $_.DBMasterKey | ? {$_.FullPath -notlike "*/*"})[0].FullPath
        } catch {$AuthOK = $false ; [System.Windows.MessageBox]::Show($("Error connecting to DataBase " + $DBStr))}
    }
    Return $AuthOK
}

    $Button_OK.add_Click.Invoke({
        If ((VerifyMasterKeys)[-1]) {
            Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
            $Window_PassType_Entrance.Close()
        }
        ((Get-Variable -Name "PasswordBox$($TabControl.SelectedIndex)").Value).Focus() | Out-Null
    })

    $Button_OK.add_Keydown({
        Param (
            [Object]$Sender,
            [System.Windows.Input.KeyEventArgs]$e
        )

        If (($e.Key -eq [System.Windows.Input.Key]::Enter) -and ($This.SecurePassword.Length -ne 0)) {
            If ((VerifyMasterKeys)[-1]) {
                Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
                $Window_PassType_Entrance.Close()
            }
        }
    })

    $Button_Quit.add_Click.Invoke({ $Window_PassType_Entrance.Close() ; Exit })

    $Index = 0
    $Global:DBInstances | % {
        $TabItem = New-Object System.Windows.Controls.TabItem
        $TabItem.Name = "TabItem$Index"
        $TabItem.Header = $_.DBName
        $TabItem.IsTabStop = $false
        #$TabItem.Focusable = $false
        $TabItem.TabIndex = $null
        $TabItem.Focus() | Out-Null

        $Grid = New-Object System.Windows.Controls.Grid
        $Grid.RowDefinitions.Add($(new-object system.windows.controls.rowdefinition))
        $Grid.ColumnDefinitions.Add($(new-object system.windows.controls.columndefinition -Property @{width = "Auto"}))
        $Grid.ColumnDefinitions.Add($(new-object system.windows.controls.columndefinition -Property @{width = "*"}))
        $Grid.Focus() | Out-Null

        $Label = New-Object System.Windows.Controls.Label
        $Label.Content = "Password:"
        [System.Windows.Controls.Grid]::SetRow($Label,0)
        [System.Windows.Controls.Grid]::SetColumn($Label,0)

        ((Get-Variable -Name "PasswordBox$Index").Value).PasswordChar = "*"
        ((Get-Variable -Name "PasswordBox$Index").Value).Margin = "3,3,3,3"
        ((Get-Variable -Name "PasswordBox$Index").Value).VerticalAlignment = "Top"
        ((Get-Variable -Name "PasswordBox$Index").Value).TabIndex = $Index
        ((Get-Variable -Name "PasswordBox$Index").Value).Background = [System.Windows.SystemColors]::ControlBrush
        ((Get-Variable -Name "PasswordBox$Index").Value).Focus() | Out-Null
        $PasswordBoxToolTip = New-Object System.Windows.Controls.ToolTip
        $PasswordBoxToolTip.Content = "MasterKey"
        ((Get-Variable -Name "PasswordBox$Index").Value).ToolTip = $PasswordBoxToolTip
        [System.Windows.Controls.Grid]::SetRow(((Get-Variable -Name "PasswordBox$Index").Value),0)
        [System.Windows.Controls.Grid]::SetColumn(((Get-Variable -Name "PasswordBox$Index").Value),1)
        ((Get-Variable -Name "PasswordBox$Index").Value).Add_KeyDown({
            Param (
                [Object]$Sender,
                [System.Windows.Input.KeyEventArgs]$e
            )

            If (($e.Key -eq [System.Windows.Input.Key]::Enter) -and ($This.SecurePassword.Length -ne 0)) {
                If ((VerifyMasterKeys)[-1]) {
                    Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
                    $Window_PassType_Entrance.Close()
                }
            }
            If (($e.Key -eq [System.Windows.Input.Key]::Tab) -and (-not [System.Windows.Input.Keyboard]::IsKeyDown("Control"))) {
                [Keyboard]::KeyDown([System.Windows.Forms.Keys]::ControlKey)
                [Keyboard]::KeyDown([System.Windows.Forms.Keys]::Tab)
                [Keyboard]::KeyUp([System.Windows.Forms.Keys]::Tab)
                [Keyboard]::KeyUp([System.Windows.Forms.Keys]::ControlKey)
            }
         })

        $Grid.AddChild(((Get-Variable -Name "PasswordBox$Index").Value))
        $Grid.AddChild($Label)

        $TabItem.AddChild($Grid)

        $TabControl.AddChild($TabItem)

        $Index += 1
    }

    $PasswordBox0.Focus() | Out-Null

    $Window_PassType_Entrance.Activate() | Out-Null
    $Window_PassType_Entrance.Focus() | Out-Null
    $Window_PassType_Entrance.ShowDialog() | Out-Null
}

If ($Global:DBInstances.Count -gt 0) { PassType_Entrance }

# Common variables, objects
$Global:Delay = 20
$InitialWindowHeight = $Window_main.Height
$Global:FadeAllowed = $true

Function SaveConfiguration {
    $Global:CheckBoxes[0] = $CheckBox_TypeOrClip.IsChecked
    $Global:CheckBoxes[1] = $CheckBox_AutoComplete.IsChecked
    [DBInstance[]]$DBInstancesOut = @()
    $Global:DBInstances | % {
        [DBInstance]$TempItem = New-Object -TypeName DBInstance
        $TempItem.DBName = $_.DBName
        $TempItem.DBPath = $_.DBPath
        $TempItem.DBKeyPath = $_.DBKeyPath
        
        $DBInstancesOut += $TempItem
    }
    $DBInstancesOut | % {$_.DBMasterKey = $null}
    If (-Not $Global:CurrentEntries) { [EntryBrief[]]$Global:CurrentEntries = @() }
    $Global:KeePass_Path,$Global:CheckBoxes,$DBInstancesOut,$Global:CurrentEntries | ConvertTo-Json | Out-File $($ExecDir + "\" + $appName + ".ini")    
}

Function SHIFT_KEY {
    param(
        [string]$KEY
    )

    [Keyboard]::KeyDown([System.Windows.Forms.Keys]::ShiftKey)# ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY)# ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyUp([System.Windows.Forms.Keys]::ShiftKey)# ; Start-Sleep -Milliseconds $Global:Delay
}

Function СTRL_KEY {
    param(
        [string]$KEY
    )

    [Keyboard]::KeyDown([System.Windows.Forms.Keys]::ControlKey)# ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY)# ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyUp([System.Windows.Forms.Keys]::ControlKey)# ; Start-Sleep -Milliseconds $Global:Delay
}

Function SINGLE_KEY {
    param(
        [string]$KEY
    )

    [Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; Start-Sleep -Milliseconds $Global:Delay
    [Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY)
}

Class KeysClass {
    [string]$KeyEntered
    [string]$FunctionName
    [string]$TypeThis
}

[KeysClass[]]$KeysArray = @()
$KeysArray += [KeysClass]@{KeyEntered = "!"; TypeThis = "D1"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "@"; TypeThis = "D2"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "#"; TypeThis = "D3"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "$"; TypeThis = "D4"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "%"; TypeThis = "D5"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "^"; TypeThis = "D6"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "&"; TypeThis = "D7"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "*"; TypeThis = "D8"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "("; TypeThis = "D9"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = ")"; TypeThis = "D0"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "_"; TypeThis = "OemMinus"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "+"; TypeThis = "Oemplus"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "<"; TypeThis = "Oemcomma"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = ">"; TypeThis = "OemPeriod"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "?"; TypeThis = "Oem2"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = ":"; TypeThis = "Oem1"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = """"; TypeThis = "Oem7"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "|"; TypeThis = "Oem5"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "{"; TypeThis = "Oem4"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "}"; TypeThis = "Oem6"; FunctionName = "SHIFT_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "``"; TypeThis = "Oem3"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "-"; TypeThis = "OemMinus"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "="; TypeThis = "Oemplus"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = ","; TypeThis = "Oemcomma"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "."; TypeThis = "OemPeriod"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "/"; TypeThis = "Oem2"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = ";"; TypeThis = "Oem1"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "'"; TypeThis = "Oem7"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "\"; TypeThis = "Oem5"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "["; TypeThis = "Oem4"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "]"; TypeThis = "Oem6"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = " "; TypeThis = "Space"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "`n"; TypeThis = "Enter"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "`t"; TypeThis = "Tab"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F1"; TypeThis = "F1"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F2"; TypeThis = "F2"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F3"; TypeThis = "F3"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F4"; TypeThis = "F4"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F5"; TypeThis = "F5"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F6"; TypeThis = "F6"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F7"; TypeThis = "F7"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F8"; TypeThis = "F8"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F9"; TypeThis = "F9"; FunctionName = "SINGLE_KEY"}
$KeysArray += [KeysClass]@{KeyEntered = "F10"; TypeThis = "F10"; FunctionName = "SINGLE_KEY"}

Function SendKey {
    param (
            [string]$KEY
    )
    
    Switch -regex -CaseSensitive ($KEY) {
        '[A-Z]$' { SHIFT_KEY $KEY }
        '[a-z]$' { [Keyboard]::KeyDown([System.Windows.Forms.Keys]::$KEY) ; [Keyboard]::KeyUp([System.Windows.Forms.Keys]::$KEY) }
        '^[0-9]' { [Keyboard]::KeyDown([System.Windows.Forms.Keys]::("D"+$KEY)) ; [Keyboard]::KeyUp([System.Windows.Forms.Keys]::("D" + $KEY)) }
        DEFAULT {
            $Index = $KeysArray.KeyEntered.IndexOf($KEY)
            If ($Index -ne -1) {&$KeysArray[$Index].FunctionName $KeysArray[$Index].TypeThis}
        }
    }
    Start-Sleep -Milliseconds $Global:Delay
}

Function Send_Credentials {
    param(
        [bool]$TypeKeys,
        [string]$uuid,
        [string]$DatabasePath_Title,
        [bool]$Ctrl,
        [bool]$Shift,
        [bool]$WinKey
    )

    $TryGetEntry = Get-KeePassEntry -MasterKey $($Global:DBInstances | ? {$_.DBPath -eq $($DatabasePath_Title.Split("`t")[0])}).DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $($DatabasePath_Title.Split("`t")[0])}).Name)  | ? {$($_.Uuid.ToHexString()) -eq $uuid}
    If ($TryGetEntry) {$Entry = $TryGetEntry}

    if ($WinKey) {  # entry URL open
        Start $Entry.URL
    } else { # Type content
                

        if (-Not $TypeKeys) {
            $PerviousClipBoard = Get-Clipboard -Raw
        }

        # Type entry name, TAB and password
        If (-Not $Shift) {
            if (-Not $Ctrl) { # type entry if not Ctrl pressed
                if ($TypeKeys) {
                    $Entry.UserName.ToCharArray() | % { SendKey $_ }
                    Start-sleep -Milliseconds 100
                } else {
                    $Entry.UserName | Set-Clipboard
                    СTRL_KEY "v"
                    $null | Set-Clipboard
                }
                [Keyboard]::KeyPress([System.Windows.Forms.Keys]::Tab)
            }
            # Waiting for the user to release the Ctrl button after click
            Start-sleep -Milliseconds 300

            if ($TypeKeys) {
                $(([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Entry.Password)))).ToCharArray() | % { SendKey $_ }
            } else {
                [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Entry.Password)) | Set-Clipboard
                СTRL_KEY "v"
                $null | Set-Clipboard
            }
        }

        if ($Shift -and (-Not $Ctrl)) {  # Type only entry Name
            if ($TypeKeys) {
                $Entry.UserName.ToCharArray() | % { SendKey $_ }
            } else {
                $Entry.UserName | Set-Clipboard
                    СTRL_KEY "v"
                    $null | Set-Clipboard
            }
        }

        if (-Not $TypeKeys) {
            $PerviousClipBoard | Set-Clipboard
        }
    }

    if ($CheckBox_AutoComplete.IsChecked) {[Keyboard]::KeyPress([System.Windows.Forms.Keys]::Enter)}
    Return
}

Get-KeePassDatabaseConfiguration | Remove-KeePassDatabaseConfiguration -Confirm:$false
$Global:DBInstances | % {
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
    $Global:DBInstances | % {
        $DatabasePath = $_.DBPath
        Get-KeePassEntry -MasterKey $_.DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $DatabasePath}).Name) | ? {$_.FullPath -notlike "*/Recycle Bin"} | % {
            $Uuid = $_.Uuid.ToHexString()
            $AttributedEntry = $Global:AttributedEntries[$Global:AttributedEntries.uuid.IndexOf($Uuid)]
            If (-Not $AttributedEntry) {
                [EntryBrief]$NewEntry = New-Object -TypeName EntryBrief
                $NewEntry.uuid = $Uuid
                $NewEntry.Name = $_.Title
                $NewEntry.DBPath = $AttributedEntry.DBPath
                $NewEntry.DBName = $AttributedEntry.DBName
                $NewEntry.OrderNum = 0
                $NewEntry.IsVisible = $false
                $EntriesUnsorted += $NewEntry
            } else {
                if ($AttributedEntry.IsVisible) {
                    [EntryBrief]$NewEntry = New-Object -TypeName EntryBrief
                    $NewEntry.uuid = $Uuid
                    $NewEntry.Name = $_.Title
                    $NewEntry.DBPath = $AttributedEntry.DBPath
                    $NewEntry.DBName = $AttributedEntry.DBName
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
    $ToolTipText = "with Ctrl - send Password$([System.Environment]::NewLine)with Shift - send Username$([System.Environment]::NewLine)with Win+Shift - open URL"
    $EntriesSorted | % {
        $Window_main.Height += 20
        $Button = [System.Windows.Controls.Button]::new()
        $Button.Name = "Button_" + $_.uuid
        $Button.Tag = $_.DBPath + "`t" + $_.Name
        $Button.Content = $_.Name
        $Button.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
        $Button.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Stretch
        $Button.Height = 20 
        $Button.Margin = "6,$( [string]($i * ($Button.Height - 1)) ),6,5"
        $Button.Foreground = "#FF263238"
        $Button.Background = "Transparent"
        $Button.BorderBrush = "#FF263238"
        If($i -eq ($EntriesSorted.Count - 1)) {$Button.BorderThickness = 1} else {$Button.BorderThickness = "1,1,1,0"}
        $Button.ToolTip = $ToolTipText
        $Button.Add_Click({
            Send_Credentials $($CheckBox_TypeOrClip.IsChecked) $($This.Name.Substring(7)) $($This.Tag) $(([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl)) -or ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightCtrl))) $([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftShift)) $(([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LWin)) -or ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RWin)))
        })
        $WindowMain_KPButtons_Grid.Children.Add($Button) | Out-Null
        $i += 1
    }
}

Button_Filter_BlinkAnimation $($Global:DBInstances.Count -eq 0)

DrawButtons

function ArrangeEntries {
    [EntryBrief[]]$EntriesAll = @()

    $Global:DBInstances | % {
        $DatabasePath = $_.DBPath
        $DatabaseName = $_.DBName
        Get-KeePassEntry -MasterKey $_.DBMasterKey -DatabaseProfileName $((Get-KeePassDatabaseConfiguration | ? {$_.DatabasePath -eq $DatabasePath}).Name) | ? {$_.FullPath -notlike "*/Recycle Bin"} | % {
            $Uuid = $_.Uuid.ToHexString()
            $AttributedEntry = $Global:AttributedEntries.Where({$_.uuid -eq $Uuid})
            [EntryBrief]$NewEntry = New-Object -TypeName EntryBrief
            $NewEntry.uuid = $Uuid
            $NewEntry.Name = $_.Title
            $NewEntry.DBPath = $DatabasePath
            $NewEntry.DBName = $DatabaseName
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

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')    | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')   | out-null
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

$Main_Tool_Icon.Add_MouseMove({
    if (-Not $Global:NotifyIconMouseOverOnce) {
        $Global:NotifyIconMouseOverOnce = $true
        $Global:PreviousWindowHandle = [SystemWindowsFunctions]::GetForegroundWindow()
    }
})

$Main_Tool_Icon.Add_Click({
    if (-Not $Global:NotifyIconMouseC1ickonce) {
        $Global:NotifyIconMouseC1ickonce = $true

        $ForegroundWindoHandle = [SystemWindowsFunctions]::GetForegroundWindow()
        [SystemWindowsFunctions]::SetWindowLong($ForegroundWindoHandle,-20,0x08000000)

        If ($Global:PreviousWindowHand1e -ne $Global:WindowMainHandle) {
            [SystemWindowsFunctions]::BringToFront($Global:PreviousWindowHandle)
            [SystemWindowsFunctions]::SetWindowLong($Global:PreviousWindowHandle,-20,0x00000000)
        }
    }
    If ($_.Button -eq [Windows.Forms.MouseButtons]::Right) {
        $Main_Tool_Icon.GetType().GetMethod("ShowContextMenu",[System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($Main_Tool_Icon,$null)
    } else {
        WindowMain_FadeAnimation -From -1 -To 2 -DurationSec 0.6
        $Global:FadeAllowed = $true
    }
})

$Button_Filter.Add_Click({
    $Global:FadeAllowed = $false
    $Global:CurrentEntries = ArrangeEntries
    $Global:CurrentEntriesCopy = $Global:CurrentEntries

    $Reader=(New-Object System.Xml.XmlNodeReader $XAMLSelectorWindow)
    try { $Window_Selector = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
    $XAMLSelectorWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_Selector.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

Function Selector_Button_Sources_BlinkAnimation {
    Param ([bool]$Animation)
    $DoubleAnimation = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $DoubleAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $DoubleAnimation.AutoReverse = $true
    $DoubleAnimation.From = 0
    $DoubleAnimation.To = 1
    $DoubleAnimation.Duration = New-TimeSpan -Seconds 0.6
    if ($Animation) {
        $Selector_Button_Sources.BeginAnimation([System.Windows.Controls.Button]::OpacityProperty, $DoubleAnimation)
    } else {
        $DoubleAnimation.BeginTime = $null
    }
}

    Selector_Button_Sources_BlinkAnimation $($Global:DBInstances.Count -eq 0)

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
        $Selector_Textbox_Search.Focus() | Out-Null
    })

    $ListView_Selector.ItemsSource = @($Global:CurrentEntries)
    $ListView_Selector.SelectedIndex = 0

    $Selector_Button_Sources.Add_Click({
        $Reader=(New-Object System.Xml.XmlNodeReader $XAMLWindow_Sources)
        try { $Window_Sources = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
        $XAMLWindow_Sources.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_Sources.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

        $Window_Sources.add_MouseLeftButtonDown({$Window_Sources.DragMove()})

        Class DBRecord {
            [string]$DBPath
            [string]$DBKeyPath
            [string]$DBName
            [bool]$DBAccessVerified
        }

        $Global:SelectedColumnIndex = 0

        $DBRecords = New-Object System.Collections.ObjectModel.ObservableCollection[DBRecord]

        $Global:DBInstances | % {
            $DBRecord = [DBRecord]::new()
            $DBRecord.DBPath = $_.DBPath
            $DBRecord.DBName = $_.DBName
            $DBRecord.DBKeyPath = $_.DBKeyPath
            $DBRecord.DBAccessVerified = $true
            $DBRecords.Add($DBRecord)
        }
        
        $Sources_DataGrid.ItemsSource = $DBRecords

        $Sources_Button_Insert.Add_Click({
            $DBRecord = [DBRecord]::new()
            $DBRecord.DBPath = ""
            $DBRecord.DBKeyPath = ""
            $DBRecord.DBName = ""
            $DBRecord.DBAccessVerified = $false
            $DBRecords.Add($DBRecord)
        })

        $Sources_Button_Verify.Add_Click({
            if ($Sources_DataGrid.SelectedIndex -ne -1) {
                $Reader=(New-Object System.Xml.XmlNodeReader $XAMLMaster_Password_Window)
                try { $Window_Password = [Windows.Markup.XamlReader]::Load($Reader) } catch { Write-Warning $_.Exception ; throw }
                $XAMLMaster_Password_Window.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % { New-Variable  -Name $_.Name -Value $Window_Password.FindName($_.Name) -Force -ErrorAction SilentlyContinue}

                $DBPath = $DBRecords[$Sources_DataGrid.SelectedIndex].DBPath
                $DBKeyPath = $DBRecords[$Sources_DataGrid.SelectedIndex].DBKeyPath

                $BottonFO_Window_Password.Add_Click({
                    If ($DBKeyPath) {
                        New-KeePassDatabaseConfiguration -DatabaseProfileName 'TestConnection' -DatabasePath $DBPath -UseMasterKey -KeyPath $DBKeyPath
                    } else {
                        New-KeePassDatabaseConfiguration -DatabaseProfileName 'TestConnection' -DatabasePath $DBPath -UseMasterKey
                    }
                    
                    Try {
                        $DBName = (Get-KeePassGroup -DatabaseProfileName 'TestConnection' -MasterKey $PasswordBox_Window_Password.SecurePassword)[0].Name
                        $DBRecords[$Sources_DataGrid.SelectedIndex].DBAccessVerified = $true
                        $DBRecords[$Sources_DataGrid.SelectedIndex].DBName = $DBName
                        $Sources_DataGrid.Items.Refresh()
                        Remove-KeePassDatabaseConfiguration -DatabaseProfileName 'TestConnection' -Confirm:$false
                        $Window_Password.Close()
                    } catch {
                        Remove-KeePassDatabaseConfiguration -DatabaseProfileName 'TestConnection' -Confirm:$false
                        [System.Windows.MessageBox]::Show("Error connecting to selected database",$([System.Windows.MessageBoxButton]::YesNo))
                    }
                })
                
                $CurrentLanguageManager = [System.Windows.Input.InputLanguageManager]::Current
                $Label_Window_Password.Content = ($CurrentLanguageManager).CurrentInputLanguage.TwoLetterISOLanguageName.ToUpper()
                $CurrentLanguageManager.add_InputLanguageChanged({
                    Param(
                        [System.Object]$Sender,
                        [System.Windows.Input.InputLanguageChangedEventArgs]$e
                    )

                    Try { $Label_NewTerm_Language.Content = $e.NewLanguage.TwoLetterISOLanguageName.ToUpper() } catch {}
                    $Label_Window_Password.Content = $e.NewLanguage.TwoLetterISOLanguageName.ToUpper()
                })

                $PasswordBox_Window_Password.Focus() | Out-Null
                $Window_Password.Owner = $Window_Sources
                $Window_Password.Activate() | Out-Null
                $Window_Password.Focus() | Out-Null
                $Window_Password.ShowDialog() | Out-Null
            }
        })

        $Sources_DataGrid.Add_MouseDoubleClick({
            $OpenFileDialog = New-Object Microsoft.Win32.OpenFileDialog
            $OpenFileDialog.Title = "Select File"
            $OpenFileDialog.InitialDirectory = ""
            $OpenFileDialog.Filter = "All Supported Files|*"
            $OpenFileDialog.Multiselect = $False
    
            if ($OpenFileDialog.ShowDialog()) {
                Switch ($Global:SelectedColumnIndex) {
                    0 {$DBRecords[$Sources_DataGrid.SelectedIndex].DBPath = $OpenFileDialog.FileName}
                    1 {$DBRecords[$Sources_DataGrid.SelectedIndex].DBKeyPath = $OpenFileDialog.FileName}
                }
                $DBRecords[$Sources_DataGrid.SelectedIndex].DBAccessVerified = $false
                $Sources_DataGrid.Items.Refresh()
            }
            $Sources_DataGrid.CancelEdit()
        })

        $Sources_Button_Remove.Add_Click({
            if ($Sources_DataGrid.SelectedIndex -ne -1) {
                $DBRecords.RemoveAt($Sources_DataGrid.SelectedIndex)
            }
        })
        
        $Sources_Button_Save.Add_Click({
            if ($DBRecords | ? {-Not $_.DBAccessVerified}) {
                [System.Windows.MessageBox]::Show("All Entries must be verified before save")
            } else {
                [DBInstance[]]$Global:DBInstances = @()
                $Global:CurrentEntries = $Global:CurrentEntries | ? {$DBRecords.DBName.IndexOf($_.DBName) -ne -1}
                $DBRecords | % {
                    [DBInstance]$NewDBInstance = [DBInstance]::new()
                    $NewDBInstance.DBName = $_.DBName
                    $NewDBInstance.DBKeyPath = $_.DBKeyPath
                    $NewDBInstance.DBPath = $_.DBPath
                    $Global:DBInstances += $NewDBInstance
                }
                SaveConfiguration
                $Window_main.OwnedWindows | % {$_.Close()}
                $Window_main.Close()
                Start powershell $PSCommandPath
                [Environment]::Exit(1)
            }
        })

        $Sources_Button_Cancel.Add_Click({
            $Window_Sources.Close()
        })
        
        $Sources_DataGrid.Add_BeginningEdit({
            $Sources_Button_Browse.IsEnabled = $true
            $Global:SelectedColumnIndex = $_.Column.DisplayIndex
        })

        $Window_Sources.Owner = $Window_Selector
        $Window_Sources.Activate() | Out-Null
        $Window_Sources.Focus() | Out-Null
        $Window_Sources.ShowDialog() | Out-Null
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
        if (($Global:KeePass_Path -eq "Undefined") -or [String]::IsNullOrEmpty($Global:KeePass_Path)) {
            $Global:KeePass_Path = "Undefined"
            $objForm = New-Object System.Windows.Forms.OpenFileDialog
            $objForm.Title = "Select location of keepass.exe"
            $objForm.InitialDirectory = $env:SystemDrive
            $objForm.Filter = "keepass.exe|keepass.exe"
            $objForm.Multiselect = $false
            $objForm.ShowDialog()
            if ($objForm.FileName) {$Global:KeePass_Path = $objForm.FileName}
        }

        If ($Global:KeePass_Path -ne "Undefined") {
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

            $Global:DBInstances | % {
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

#$CheckBox_AlwaysOnTop.Add_Checked({ SaveConfiguration }) ; $CheckBox_AlwaysOnTop.Add_UnChecked({ SaveConfiguration })
#$CheckBox_AutoComplete.Add_Checked({ SaveConfiguration }) ; $CheckBox_AlwaysOnTop.Add_UnChecked({ SaveConfiguration })

$Window_main.add_MouseLeftButtonDown({$Window_main.DragMove()})

$Window_main.Add_MouseEnter({ If ($Window_main.Opacity -ne 1) { WindowMain_FadeAnimation -From $Window_main.Opacity -to 2 -DurationSec 0.6 } })

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
    Start-Sleep -Milliseconds 400
    If ((Get-Clipboard -Raw) -match '[fF]\d{1,2};') { # Typing functional keys from clipboard, format Fn; - examples F2;F8;F10;
        [regex]::Matches($(Get-Clipboard -Raw),"[fF]\d{1,2};") | % {
            SendKey $(($_.Value).TrimEnd(";"))
        }
    } else {# Typing non Fn keys
        (Get-Clipboard -Raw).ToCharArray() | % { SendKey $_ }
    }
    if ($CheckBox_AutoComplete.IsChecked) {[Keyboard]::KeyPress([System.Windows.Forms.Keys]::Enter)}
})

$Button_Hide.add_Click.Invoke({$Global:FadeAllowed = $False ; WindowMain_FadeAnimation -From 2 -to -1 -DurationSec 0.6})

$Menu_Exit.add_Click({
    SaveConfiguration
    $Global:DBInstances | % {$_.DBMasterKey = $null}
    $Window_main.OwnedWindows | % {$_.Close()}
    $Window_main.Close()
    [Environment]::Exit(1)
})

$Window_main.Add_Loaded({
    $Window_main.Title = $appName + " v." + $appVersion
    $CheckBox_TypeOrClip.IsChecked = $Global:CheckBoxes[0]
    $CheckBox_AutoComplete.IsChecked = $Global:CheckBoxes[1]
    Try { if (Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $appName) {$CheckBox_AutoRun.IsChecked = $true} } catch {}

    #$WindowHandle = (Get-Process | ? {(($_.Name -eq "powershell")  -or ($_.Name -eq "powershell_ise") -or ($_.Name -eq "pwsh")) -and ($_.MainWindowTitle -eq $Window_main.Title)}).MainWindowHandle
    
    # Set Window never getting focus after activation - ref: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga, https://learn.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
    $Global:WindowMainHandle = [SystemWindowsFunctions]::GetForegroundWindow()
    [SystemWindowsFunctions]::SetWindowLong($Global:WindowMainHandle,-20,0x08000000)

    WindowMain_FadeAnimation -From 0 -to 1 -DurationSec 0.6
})

$Window_main.Activate() | Out-Null
$Window_main.ShowDialog()

[System.GC]::Collect()
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)