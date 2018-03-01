# chocolatey-gpo - Linking the chocolatey package manager with Active Directory Group Policies
# Copyright (c) 2018 Dorian Stoll
# Licensed under the Terms of the MIT License

# Hide the powershell window
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 0)

function Get-RegValue([String] $KeyPath, [String] $ValueName) {
    (Get-ItemProperty -LiteralPath $KeyPath -Name $ValueName).$ValueName
}
function Get-RegValues([String] $KeyPath) {
    $RegKey = (Get-ItemProperty $KeyPath)
    $RegKey.PSObject.Properties | 
        Where-Object { $_.Name -ne "PSPath" -and $_.Name -ne "PSParentPath" -and $_.Name -ne "PSChildName" -and $_.Name -ne "PSDrive" -and $_.Name -ne "PSProvider" } | 
        ForEach-Object {
        $_.Name
    }
}
function Call-CMD([String] $command) {
    return cmd.exe /C $command
}

#region Window Overlay

# Creates a new simple message box. Stripped down version of this
# https://smsagent.wordpress.com/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
function New-WPFMessageBox {

    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$False,Position=1)]
        [string]$Title,

        # Content font size
        [Parameter(Mandatory=$False,Position=2)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$False,Position=3)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$False,Position=4)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$False,Position=5)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$False,Position=6)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$False,Position=7)]
        [int]$BlurRadius = 20

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segui"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {
        
# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="Manual" Top="10" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid Cursor="No">
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Name="ContentText" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

        # Load the window from XAML
        $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

        # Remove the title bar if no title is provided
        if ($Title -eq "")
        {
            $TitleBar = $Window.FindName('TitleBar')
            $Window.FindName('StackPanel').Children.Remove($TitleBar)
        }

        # Replace double quotes with single to avoid quote issues in strings
        if ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)


        # Display the window
        $window.Show();
        $Window.Left = [System.Windows.SystemParameters]::PrimaryScreenWidth - $Window.Width - 10;
        return $Window
    }
}

function Display-StatusMessage([string] $message) {
    $Params = @{
        FontFamily = 'Verdana'
        Title = "          CHOCOLATEY          "
        TitleFontSize = 20
        TitleTextForeground = 'WhiteSmoke'
        TitleBackground = 'SteelBlue'
        TitleFontWeight = "UltraBold"
        ContentFontSize = 16
        ContentBackground = 'WhiteSmoke'
        BorderThickness = 0
    }
    $newWindow = New-WPFMessageBox @Params -Content "$message"
    if ($global:_window -ne $null) {
        $global:_window.Close()
    }
    $global:_window = $newWindow
}

#endregion

#region Install Boxstarter & Chocolatey

function Check-Chocolatey {
    if(-not $env:ChocolateyInstall -or -not (Test-Path "$env:ChocolateyInstall\bin\choco.exe")) {
        return $false;
    } else {
        return $true
    }
}


Display-StatusMessage "Installing Chocolatey..."
if (!(Check-Chocolatey)) {
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString("http://boxstarter.org/bootstrapper.ps1")); Get-Boxstarter -Force
}
if (Test-Path "C:\Users\Public\Desktop\Boxstarter Shell.lnk") { 
    Remove-Item -Path "C:\Users\Public\Desktop\Boxstarter Shell.lnk" -Force
}

#endregion

#region Install Updates

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey" | ForEach-Object {
    if ($_ -eq "installUpdates") {
        $mode = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey" $_
        if ($mode -eq 1) {
            Display-StatusMessage "Installing Windows Updates..."
            Install-WindowsUpdate -Full -AcceptEula -SupressReboots
        }
    }
}
Update-ExecutionPolicy Unrestricted

#endregion

#region Apply Config

Display-StatusMessage "Configuring Chocolatey..."
Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Config" | ForEach-Object {
    if (!($_ -match ".+?Mode$")) {
        $value = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Config" $_
        $mode = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Config" ($_ + "Mode")
        
        if ($mode -eq 1) {
            # set
            Call-CMD "choco config set --name $_ --value $value"
        }
        else {
            # unset
            Call-CMD "choco config unset --name $_"
        }
    }
}

#endregion

#region Apply Features

Display-StatusMessage "Configuring Chocolatey Features..."
Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Features" | ForEach-Object { 
    $mode = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Features" $_
    
    if ($mode -eq 1) {
        # enable
        Call-CMD "choco feature enable --name $_"
    }
    elseif ($mode -eq 0) {
        # disable
        Call-CMD "choco feature disable --name $_"
    }
}

#endregion

#region Package Sources

Display-StatusMessage "Configuring Chocolatey Sources..."
Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Sources" | ForEach-Object {
    $value = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Sources" $_
    if ($value -eq "remove") {
        Call-CMD "choco sources remove -n $_"
    }
    else {
        Call-CMD "choco sources add -n $_ -s $value"
    }
}

#endregion

#region Install Packages

Get-RegValues "HKLM:\SOFTWARE\Policies\Chocolatey\Packages" | ForEach-Object {
    $params = Get-RegValue "HKLM:\SOFTWARE\Policies\Chocolatey\Packages" $_
    $packageList = Call-CMD "choco list --local-only"
    if ($params -eq "remove") {
        if ($packageList -match $_ + " ") {
            Display-StatusMessage "Removing $_..."
            Call-CMD "cuninst $_ -y"
        }
    }
    else {
        if (!($packageList -match $_ + " ")) {
            Display-StatusMessage "Installing $_..."
            Call-CMD "cinst $_ -y $params"
        } else {
            Display-StatusMessage "Updating $_..."
            Call-CMD "cup $_ -y $params"
        }
    }
}

#endregion

#region Finalize

$_window.Close();

#endregion
