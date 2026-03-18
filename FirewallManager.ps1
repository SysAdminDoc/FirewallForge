<#
.SYNOPSIS
    FirewallManager - Windows Firewall Rule Backup, Edit, and Restore Tool
.DESCRIPTION
    A professional GUI application for backing up, editing, and restoring Windows Firewall rules.
.NOTES
    Author: Matt
    Version: 1.1.0
    Requires: Administrator privileges
#>

# ============================================================
# Error Handling Wrapper
# ============================================================
$ErrorActionPreference = "Stop"

try {
    # Check for admin rights
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        throw "This script requires Administrator privileges. Please right-click and 'Run as Administrator'."
    }

    Write-Host "Loading assemblies..." -ForegroundColor Cyan
    Add-Type -AssemblyName PresentationFramework
    Write-Host "  - PresentationFramework OK" -ForegroundColor Gray
    Add-Type -AssemblyName PresentationCore
    Write-Host "  - PresentationCore OK" -ForegroundColor Gray
    Add-Type -AssemblyName WindowsBase
    Write-Host "  - WindowsBase OK" -ForegroundColor Gray
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "  - System.Windows.Forms OK" -ForegroundColor Gray
    Write-Host "All assemblies loaded successfully.`n" -ForegroundColor Green

# ============================================================
# XAML GUI Definition
# ============================================================
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FirewallManager v1.1.0"
        Height="750" Width="1050"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E1E"
        ResizeMode="CanResizeWithGrip">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#1084D9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#006CBE"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#555555"/>
                                <Setter Property="Foreground" Value="#888888"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DangerButton" TargetType="Button">
            <Setter Property="Background" Value="#D32F2F"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E53935"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#B71C1C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SuccessButton" TargetType="Button">
            <Setter Property="Background" Value="#388E3C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#43A047"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#2E7D32"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="WarningButton" TargetType="Button">
            <Setter Property="Background" Value="#F57C00"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#FF9800"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#E65100"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#252526"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#3C3C3C"/>
            <Setter Property="RowBackground" Value="#2D2D30"/>
            <Setter Property="AlternatingRowBackground" Value="#252526"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#3C3C3C"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderBrush" Value="#4C4C4C"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#0078D4"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <SolidColorBrush x:Key="ComboBoxBackground" Color="#3C3C3C"/>
        <SolidColorBrush x:Key="ComboBoxBorder" Color="#555555"/>
        <SolidColorBrush x:Key="ComboBoxForeground" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="ComboBoxHoverBackground" Color="#4A4A4A"/>
        <SolidColorBrush x:Key="ComboBoxDropdownBackground" Color="#2D2D30"/>
        <SolidColorBrush x:Key="ComboBoxItemHover" Color="#3E3E42"/>
        <SolidColorBrush x:Key="ComboBoxItemSelected" Color="#0078D4"/>

        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" Background="{StaticResource ComboBoxBackground}"
                        BorderBrush="{StaticResource ComboBoxBorder}" BorderThickness="1" CornerRadius="3"/>
                <Border Grid.Column="0" Background="Transparent" Margin="1"/>
                <Path x:Name="Arrow" Grid.Column="1" Fill="#E0E0E0" HorizontalAlignment="Center"
                      VerticalAlignment="Center" Data="M 0 0 L 6 6 L 12 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxHoverBackground}"/>
                </Trigger>
                <Trigger Property="IsChecked" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxHoverBackground}"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="TextBox">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="Transparent"/>
        </ControlTemplate>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Border" Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}" BorderThickness="0">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxItemHover}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxItemSelected}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{StaticResource ComboBoxBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource ComboBoxForeground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ComboBoxBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton" Template="{StaticResource ComboBoxToggleButton}"
                                          Grid.Column="2" Focusable="False"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="10,3,30,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid Name="DropDown" SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="{StaticResource ComboBoxDropdownBackground}"
                                            BorderThickness="1" BorderBrush="{StaticResource ComboBoxBorder}" CornerRadius="3">
                                        <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#B0B0B0"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="FirewallManager" FontSize="28" FontWeight="Bold" Foreground="#0078D4"/>
            <TextBlock Text="Backup, Edit, and Restore Windows Firewall Rules" FontSize="14" Foreground="#808080"/>
        </StackPanel>

        <!-- Action Buttons Row 1 -->
        <WrapPanel Grid.Row="1" Margin="0,0,0,5">
            <Button x:Name="btnBackup" Content="Backup Rules" Width="140"/>
            <Button x:Name="btnRestore" Content="Restore Rules" Style="{StaticResource SuccessButton}" Width="140"/>
            <Button x:Name="btnRefresh" Content="Refresh List" Width="140"/>
            <Button x:Name="btnExportCSV" Content="Export to CSV" Width="140"/>
            <Button x:Name="btnFindDuplicates" Content="Find Duplicates" Style="{StaticResource WarningButton}" Width="140"/>
            <Button x:Name="btnQuickBlock" Content="Quick Block" Style="{StaticResource DangerButton}" Width="120"/>
        </WrapPanel>

        <!-- Search and Filter Row -->
        <WrapPanel Grid.Row="2" Margin="0,0,0,10" VerticalAlignment="Center">
            <TextBox x:Name="txtSearch" Width="250" Margin="5,5,5,5"
                     ToolTip="Search rules by name, program, or port"/>
            <Button x:Name="btnSearch" Content="Search" Width="80"/>
            <Button x:Name="btnClearSearch" Content="Clear" Width="80"/>
            <Label Content="Group:" VerticalAlignment="Center" Margin="15,0,0,0"/>
            <ComboBox x:Name="cmbGroupFilter" Width="200" Margin="5" ToolTip="Filter by rule group"/>
            <Button x:Name="btnToggleStats" Content="Stats" Width="80" Style="{StaticResource WarningButton}"/>
        </WrapPanel>

        <!-- Stats Panel (collapsible) -->
        <Border x:Name="statsPanel" Grid.Row="3" Background="#252526" Margin="0,0,0,10" Padding="15,10" CornerRadius="4"
                BorderBrush="#3C3C3C" BorderThickness="1" Visibility="Collapsed">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="txtStatTotal" Text="Total: 0" Foreground="#E0E0E0" FontSize="14" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtStatInbound" Text="  Inbound: 0" Foreground="#B0B0B0" FontSize="12" Margin="0,2,0,0"/>
                    <TextBlock x:Name="txtStatOutbound" Text="  Outbound: 0" Foreground="#B0B0B0" FontSize="12" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                    <TextBlock x:Name="txtStatEnabled" Text="Enabled: 0" Foreground="#00FF00" FontSize="14" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtStatDisabled" Text="Disabled: 0" Foreground="#FF6666" FontSize="12" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                    <TextBlock x:Name="txtStatAllow" Text="Allow: 0" Foreground="#4CAF50" FontSize="14" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtStatBlock" Text="Block: 0" Foreground="#F44336" FontSize="12" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="3">
                    <TextBlock x:Name="txtStatDomain" Text="Domain: 0" Foreground="#B0B0B0" FontSize="12"/>
                    <TextBlock x:Name="txtStatPrivate" Text="Private: 0" Foreground="#B0B0B0" FontSize="12" Margin="0,2,0,0"/>
                    <TextBlock x:Name="txtStatPublic" Text="Public: 0" Foreground="#B0B0B0" FontSize="12" Margin="0,2,0,0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Rules DataGrid -->
        <DataGrid x:Name="dgRules" Grid.Row="4"
                  AutoGenerateColumns="False"
                  IsReadOnly="True"
                  SelectionMode="Extended"
                  CanUserAddRows="False"
                  CanUserDeleteRows="False"
                  CanUserReorderColumns="True"
                  CanUserSortColumns="True"
                  VirtualizingPanel.IsVirtualizing="True"
                  VirtualizingPanel.VirtualizationMode="Recycling">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="200"/>
                <DataGridTextColumn Header="Direction" Binding="{Binding Direction}" Width="80"/>
                <DataGridTextColumn Header="Action" Binding="{Binding Action}" Width="70"/>
                <DataGridTextColumn Header="Enabled" Binding="{Binding Enabled}" Width="70"/>
                <DataGridTextColumn Header="Profile" Binding="{Binding Profile}" Width="100"/>
                <DataGridTextColumn Header="Protocol" Binding="{Binding Protocol}" Width="80"/>
                <DataGridTextColumn Header="Local Port" Binding="{Binding LocalPort}" Width="100"/>
                <DataGridTextColumn Header="Remote Port" Binding="{Binding RemotePort}" Width="100"/>
                <DataGridTextColumn Header="Program" Binding="{Binding Program}" Width="250"/>
                <DataGridTextColumn Header="Group" Binding="{Binding Group}" Width="150"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Edit Panel -->
        <GroupBox Grid.Row="5" Header="Edit Selected Rule" Margin="0,15,0,0"
                  Foreground="#B0B0B0" BorderBrush="#3C3C3C">
            <Grid Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Column="0" Grid.Row="0" Margin="5">
                    <Label Content="Enabled"/>
                    <ComboBox x:Name="cmbEnabled">
                        <ComboBoxItem Content="True" IsSelected="True"/>
                        <ComboBoxItem Content="False"/>
                    </ComboBox>
                </StackPanel>

                <StackPanel Grid.Column="1" Grid.Row="0" Margin="5">
                    <Label Content="Action"/>
                    <ComboBox x:Name="cmbAction">
                        <ComboBoxItem Content="Allow"/>
                        <ComboBoxItem Content="Block"/>
                    </ComboBox>
                </StackPanel>

                <StackPanel Grid.Column="2" Grid.Row="0" Margin="5">
                    <Label Content="Direction"/>
                    <ComboBox x:Name="cmbDirection">
                        <ComboBoxItem Content="Inbound"/>
                        <ComboBoxItem Content="Outbound"/>
                    </ComboBox>
                </StackPanel>

                <StackPanel Grid.Column="3" Grid.Row="0" Margin="5">
                    <Label Content="Profile"/>
                    <ComboBox x:Name="cmbProfile">
                        <ComboBoxItem Content="Any"/>
                        <ComboBoxItem Content="Domain"/>
                        <ComboBoxItem Content="Private"/>
                        <ComboBoxItem Content="Public"/>
                        <ComboBoxItem Content="Domain, Private"/>
                        <ComboBoxItem Content="Domain, Public"/>
                        <ComboBoxItem Content="Private, Public"/>
                    </ComboBox>
                </StackPanel>

                <StackPanel Grid.Column="0" Grid.Row="1" Margin="5">
                    <Label Content="Protocol"/>
                    <ComboBox x:Name="cmbProtocol">
                        <ComboBoxItem Content="Any"/>
                        <ComboBoxItem Content="TCP"/>
                        <ComboBoxItem Content="UDP"/>
                        <ComboBoxItem Content="ICMPv4"/>
                        <ComboBoxItem Content="ICMPv6"/>
                    </ComboBox>
                </StackPanel>

                <StackPanel Grid.Column="1" Grid.Row="1" Margin="5">
                    <Label Content="Local Port"/>
                    <TextBox x:Name="txtLocalPort"/>
                </StackPanel>

                <StackPanel Grid.Column="2" Grid.Row="1" Margin="5">
                    <Label Content="Remote Port"/>
                    <TextBox x:Name="txtRemotePort"/>
                </StackPanel>

                <StackPanel Grid.Column="3" Grid.Row="1" Grid.ColumnSpan="2" Margin="5" Orientation="Horizontal"
                            VerticalAlignment="Bottom">
                    <Button x:Name="btnApplyEdit" Content="Apply Changes" Style="{StaticResource SuccessButton}"/>
                    <Button x:Name="btnDeleteRule" Content="Delete Rule" Style="{StaticResource DangerButton}"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- Status Bar -->
        <Border Grid.Row="6" Background="#252526" Margin="0,15,0,0" Padding="10,8" CornerRadius="4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Text="Loading firewall rules, please wait..." Foreground="#FFA500" VerticalAlignment="Center"/>
                <TextBlock x:Name="txtRuleCount" Grid.Column="1" Text="Rules: 0" Foreground="#808080" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ============================================================
# Load XAML
# ============================================================
Write-Host "Parsing XAML interface..." -ForegroundColor Cyan
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)
Write-Host "  - XAML parsed successfully" -ForegroundColor Gray

Write-Host "Binding controls..." -ForegroundColor Cyan
# Get controls
$btnBackup = $Window.FindName("btnBackup")
$btnRestore = $Window.FindName("btnRestore")
$btnRefresh = $Window.FindName("btnRefresh")
$btnExportCSV = $Window.FindName("btnExportCSV")
$btnFindDuplicates = $Window.FindName("btnFindDuplicates")
$btnQuickBlock = $Window.FindName("btnQuickBlock")
$btnSearch = $Window.FindName("btnSearch")
$btnClearSearch = $Window.FindName("btnClearSearch")
$btnToggleStats = $Window.FindName("btnToggleStats")
$btnApplyEdit = $Window.FindName("btnApplyEdit")
$btnDeleteRule = $Window.FindName("btnDeleteRule")
$txtSearch = $Window.FindName("txtSearch")
$txtStatus = $Window.FindName("txtStatus")
$txtRuleCount = $Window.FindName("txtRuleCount")
$dgRules = $Window.FindName("dgRules")
$cmbEnabled = $Window.FindName("cmbEnabled")
$cmbAction = $Window.FindName("cmbAction")
$cmbDirection = $Window.FindName("cmbDirection")
$cmbProfile = $Window.FindName("cmbProfile")
$cmbProtocol = $Window.FindName("cmbProtocol")
$cmbGroupFilter = $Window.FindName("cmbGroupFilter")
$txtLocalPort = $Window.FindName("txtLocalPort")
$txtRemotePort = $Window.FindName("txtRemotePort")
$statsPanel = $Window.FindName("statsPanel")
$txtStatTotal = $Window.FindName("txtStatTotal")
$txtStatInbound = $Window.FindName("txtStatInbound")
$txtStatOutbound = $Window.FindName("txtStatOutbound")
$txtStatEnabled = $Window.FindName("txtStatEnabled")
$txtStatDisabled = $Window.FindName("txtStatDisabled")
$txtStatAllow = $Window.FindName("txtStatAllow")
$txtStatBlock = $Window.FindName("txtStatBlock")
$txtStatDomain = $Window.FindName("txtStatDomain")
$txtStatPrivate = $Window.FindName("txtStatPrivate")
$txtStatPublic = $Window.FindName("txtStatPublic")

# Global variables
$Script:AllRules = @()
$Script:BackupFolder = Join-Path $env:USERPROFILE "FirewallBackups"
$Script:SearchActive = $false
Write-Host "  - Controls bound successfully`n" -ForegroundColor Gray

# ============================================================
# Functions
# ============================================================
function Update-Status {
    param([string]$Message, [string]$Color = "#808080")
    $txtStatus.Text = $Message
    $txtStatus.Foreground = $Color
    $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Update-Statistics {
    $rules = $Script:AllRules
    if (-not $rules -or $rules.Count -eq 0) {
        $txtStatTotal.Text = "Total: 0"
        $txtStatInbound.Text = "  Inbound: 0"
        $txtStatOutbound.Text = "  Outbound: 0"
        $txtStatEnabled.Text = "Enabled: 0"
        $txtStatDisabled.Text = "Disabled: 0"
        $txtStatAllow.Text = "Allow: 0"
        $txtStatBlock.Text = "Block: 0"
        $txtStatDomain.Text = "Domain: 0"
        $txtStatPrivate.Text = "Private: 0"
        $txtStatPublic.Text = "Public: 0"
        return
    }

    $total = $rules.Count
    $inbound = @($rules | Where-Object { $_.Direction -eq "Inbound" }).Count
    $outbound = @($rules | Where-Object { $_.Direction -eq "Outbound" }).Count
    $enabled = @($rules | Where-Object { $_.Enabled -eq "True" }).Count
    $disabled = @($rules | Where-Object { $_.Enabled -eq "False" }).Count
    $allow = @($rules | Where-Object { $_.Action -eq "Allow" }).Count
    $block = @($rules | Where-Object { $_.Action -eq "Block" }).Count
    $domain = @($rules | Where-Object { $_.Profile -match "Domain" }).Count
    $private = @($rules | Where-Object { $_.Profile -match "Private" }).Count
    $public = @($rules | Where-Object { $_.Profile -match "Public" }).Count

    $txtStatTotal.Text = "Total: $total"
    $txtStatInbound.Text = "  Inbound: $inbound"
    $txtStatOutbound.Text = "  Outbound: $outbound"
    $txtStatEnabled.Text = "Enabled: $enabled"
    $txtStatDisabled.Text = "Disabled: $disabled"
    $txtStatAllow.Text = "Allow: $allow"
    $txtStatBlock.Text = "Block: $block"
    $txtStatDomain.Text = "Domain: $domain"
    $txtStatPrivate.Text = "Private: $private"
    $txtStatPublic.Text = "Public: $public"
}

function Update-GroupFilter {
    $cmbGroupFilter.Items.Clear()
    $cmbGroupFilter.Items.Add("(All Groups)") | Out-Null

    $groups = @($Script:AllRules | Where-Object { -not [string]::IsNullOrEmpty($_.Group) } |
                Select-Object -ExpandProperty Group -Unique | Sort-Object)

    foreach ($g in $groups) {
        $cmbGroupFilter.Items.Add($g) | Out-Null
    }

    $cmbGroupFilter.SelectedIndex = 0
}

function Get-FirewallRules {
    Update-Status "Loading firewall rules... (this may take a moment)" "#FFA500"
    $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        Write-Host "  Fetching firewall rules..." -ForegroundColor Gray
        $rawRules = @(Get-NetFirewallRule -ErrorAction Stop)
        $ruleCount = $rawRules.Count
        Write-Host "  Found $ruleCount rules" -ForegroundColor Gray

        Update-Status "Fetching port filters..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Fetching port filters..." -ForegroundColor Gray
        $portFilters = @{}
        Get-NetFirewallPortFilter -ErrorAction SilentlyContinue | ForEach-Object {
            $portFilters[$_.InstanceID] = $_
        }

        Update-Status "Fetching application filters..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Fetching application filters..." -ForegroundColor Gray
        $appFilters = @{}
        Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue | ForEach-Object {
            $appFilters[$_.InstanceID] = $_
        }

        Update-Status "Processing $ruleCount rules..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Processing rules..." -ForegroundColor Gray

        $rules = $rawRules | ForEach-Object {
            $ruleId = $_.Name
            $portFilter = $portFilters[$_.InstanceID]
            $appFilter = $appFilters[$_.InstanceID]

            [PSCustomObject]@{
                Name = $_.Name
                DisplayName = $_.DisplayName
                Description = $_.Description
                Direction = $_.Direction.ToString()
                Action = $_.Action.ToString()
                Enabled = $_.Enabled.ToString()
                Profile = $_.Profile.ToString()
                Protocol = if ($portFilter) { $portFilter.Protocol } else { "Any" }
                LocalPort = if ($portFilter -and $portFilter.LocalPort) { $portFilter.LocalPort } else { "Any" }
                RemotePort = if ($portFilter -and $portFilter.RemotePort) { $portFilter.RemotePort } else { "Any" }
                Program = if ($appFilter -and $appFilter.Program) { $appFilter.Program } else { "Any" }
                Group = if ($_.Group) { $_.Group } else { "" }
            }
        }

        $stopwatch.Stop()
        $elapsed = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)

        $Script:AllRules = $rules
        $dgRules.ItemsSource = $rules
        $txtRuleCount.Text = "Rules: $($rules.Count)"
        Update-Status "Loaded $($rules.Count) firewall rules in $elapsed seconds" "#00FF00"
        Update-Statistics
        Update-GroupFilter
        Write-Host "  Completed in $elapsed seconds" -ForegroundColor Green
    }
    catch {
        $stopwatch.Stop()
        Update-Status "Error loading rules: $($_.Exception.Message)" "#FF0000"
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.MessageBox]::Show(
            "Failed to load firewall rules.`n`nError: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Backup-FirewallRules {
    Update-Status "Preparing backup..." "#FFA500"

    # Ensure backup folder exists
    if (-not (Test-Path $Script:BackupFolder)) {
        New-Item -ItemType Directory -Path $Script:BackupFolder -Force | Out-Null
    }

    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.InitialDirectory = $Script:BackupFolder
    $saveDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = ".fwbackup"
    $saveDialog.FileName = "FirewallBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    if ($saveDialog.ShowDialog()) {
        try {
            Update-Status "Exporting firewall rules..." "#FFA500"

            # Export using netsh for complete backup
            $tempFile = [System.IO.Path]::GetTempFileName()
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue  # GetTempFileName creates the file, netsh needs it gone
            $netshResult = netsh advfirewall export $tempFile 2>&1

            if ($LASTEXITCODE -eq 0 -and (Test-Path $tempFile)) {
                # Create our backup package with metadata
                $backupData = @{
                    BackupDate = (Get-Date).ToString("o")
                    ComputerName = $env:COMPUTERNAME
                    RuleCount = $Script:AllRules.Count
                    NetshBackup = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($tempFile))
                    RuleDetails = $Script:AllRules
                }

                $backupData | ConvertTo-Json -Depth 10 -Compress | Out-File $saveDialog.FileName -Encoding UTF8
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

                Update-Status "Backup saved successfully: $($saveDialog.FileName)" "#00FF00"
                [System.Windows.MessageBox]::Show(
                    "Firewall rules backed up successfully!`n`nFile: $($saveDialog.FileName)`nRules: $($Script:AllRules.Count)",
                    "Backup Complete",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
            }
            else {
                throw "netsh export failed: $netshResult"
            }
        }
        catch {
            Update-Status "Backup failed: $($_.Exception.Message)" "#FF0000"
            [System.Windows.MessageBox]::Show(
                "Failed to backup firewall rules.`n`nError: $($_.Exception.Message)",
                "Backup Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
    else {
        Update-Status "Backup cancelled" "#808080"
    }
}

function Restore-FirewallRules {
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.InitialDirectory = $Script:BackupFolder
    $openDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|All Files (*.*)|*.*"
    $openDialog.DefaultExt = ".fwbackup"

    if ($openDialog.ShowDialog()) {
        try {
            Update-Status "Reading backup file..." "#FFA500"

            $backupContent = Get-Content $openDialog.FileName -Raw -Encoding UTF8
            $backupData = $backupContent | ConvertFrom-Json

            $confirmResult = [System.Windows.MessageBox]::Show(
                "Restore firewall rules from backup?`n`nBackup Date: $($backupData.BackupDate)`nOriginal Computer: $($backupData.ComputerName)`nRules: $($backupData.RuleCount)`n`nWARNING: This will replace your current firewall configuration!",
                "Confirm Restore",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )

            if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
                Update-Status "Restoring firewall rules..." "#FFA500"

                # Extract and restore netsh backup
                $tempFile = [System.IO.Path]::GetTempFileName()
                $netshBytes = [Convert]::FromBase64String($backupData.NetshBackup)
                [System.IO.File]::WriteAllBytes($tempFile, $netshBytes)

                $netshResult = netsh advfirewall import $tempFile 2>&1
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

                if ($LASTEXITCODE -eq 0) {
                    Update-Status "Restore completed successfully" "#00FF00"
                    [System.Windows.MessageBox]::Show(
                        "Firewall rules restored successfully!",
                        "Restore Complete",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )

                    # Refresh the display
                    Get-FirewallRules
                }
                else {
                    throw "netsh import failed: $netshResult"
                }
            }
            else {
                Update-Status "Restore cancelled" "#808080"
            }
        }
        catch {
            Update-Status "Restore failed: $($_.Exception.Message)" "#FF0000"
            [System.Windows.MessageBox]::Show(
                "Failed to restore firewall rules.`n`nError: $($_.Exception.Message)",
                "Restore Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
}

function Search-Rules {
    $searchText = $txtSearch.Text.Trim()
    $groupFilter = $cmbGroupFilter.SelectedItem

    $filtered = $Script:AllRules

    # Apply group filter
    if ($groupFilter -and $groupFilter -ne "(All Groups)") {
        $filtered = @($filtered | Where-Object { $_.Group -eq $groupFilter })
    }

    # Apply text search
    if (-not [string]::IsNullOrEmpty($searchText)) {
        $Script:SearchActive = $true
        $filtered = @($filtered | Where-Object {
            $_.DisplayName -like "*$searchText*" -or
            $_.Program -like "*$searchText*" -or
            $_.LocalPort -like "*$searchText*" -or
            $_.RemotePort -like "*$searchText*" -or
            $_.Description -like "*$searchText*"
        })
    }
    else {
        $Script:SearchActive = $false
    }

    if ($filtered.Count -eq $Script:AllRules.Count -and -not $Script:SearchActive) {
        $dgRules.ItemsSource = $Script:AllRules
        $txtRuleCount.Text = "Rules: $($Script:AllRules.Count)"
        Update-Status "Showing all rules" "#808080"
    }
    else {
        $dgRules.ItemsSource = $filtered
        $txtRuleCount.Text = "Rules: $($filtered.Count) / $($Script:AllRules.Count)"
        Update-Status "Found $($filtered.Count) matching rules" "#00FF00"
    }

    # Apply search highlighting via row style
    Apply-SearchHighlighting
}

function Apply-SearchHighlighting {
    $searchText = $txtSearch.Text.Trim()

    if ([string]::IsNullOrEmpty($searchText)) {
        # Reset row styles - clear any bold formatting
        $dgRules.RowStyle = $null
        return
    }

    # Create a row style that makes matching rows bold with accent foreground
    $rowStyle = New-Object System.Windows.Style ([System.Type]::GetType("System.Windows.Controls.DataGridRow"))
    $rowStyle.Setters.Add((New-Object System.Windows.Setter ([System.Windows.Controls.DataGridRow]::FontWeightProperty, [System.Windows.FontWeights]::Bold)))
    $rowStyle.Setters.Add((New-Object System.Windows.Setter ([System.Windows.Controls.DataGridRow]::ForegroundProperty, (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x4F, 0xC3, 0xF7))))))
    $dgRules.RowStyle = $rowStyle
}

function Export-RulesToCSV {
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = ".csv"
    $saveDialog.FileName = "FirewallRules_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    if ($saveDialog.ShowDialog()) {
        try {
            $dgRules.ItemsSource | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            Update-Status "Exported to: $($saveDialog.FileName)" "#00FF00"
            [System.Windows.MessageBox]::Show(
                "Rules exported successfully!`n`nFile: $($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        catch {
            Update-Status "Export failed: $($_.Exception.Message)" "#FF0000"
        }
    }
}

function Update-EditPanel {
    $selectedRule = $dgRules.SelectedItem

    if ($null -ne $selectedRule) {
        # Set Enabled
        $cmbEnabled.SelectedIndex = if ($selectedRule.Enabled -eq "True") { 0 } else { 1 }

        # Set Action
        $cmbAction.SelectedIndex = if ($selectedRule.Action -eq "Allow") { 0 } else { 1 }

        # Set Direction
        $cmbDirection.SelectedIndex = if ($selectedRule.Direction -eq "Inbound") { 0 } else { 1 }

        # Set Profile
        $profileMap = @{
            "Any" = 0; "Domain" = 1; "Private" = 2; "Public" = 3;
            "Domain, Private" = 4; "Domain, Public" = 5; "Private, Public" = 6
        }
        $cmbProfile.SelectedIndex = if ($profileMap.ContainsKey($selectedRule.Profile)) { $profileMap[$selectedRule.Profile] } else { 0 }

        # Set Protocol
        $protocolMap = @{ "Any" = 0; "TCP" = 1; "UDP" = 2; "ICMPv4" = 3; "ICMPv6" = 4 }
        $cmbProtocol.SelectedIndex = if ($protocolMap.ContainsKey($selectedRule.Protocol)) { $protocolMap[$selectedRule.Protocol] } else { 0 }

        # Set Ports
        $txtLocalPort.Text = $selectedRule.LocalPort
        $txtRemotePort.Text = $selectedRule.RemotePort

        Update-Status "Selected: $($selectedRule.DisplayName)" "#0078D4"
    }
}

function Apply-RuleChanges {
    $selectedRule = $dgRules.SelectedItem

    if ($null -eq $selectedRule) {
        [System.Windows.MessageBox]::Show(
            "Please select a rule to edit.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    try {
        Update-Status "Applying changes to: $($selectedRule.DisplayName)" "#FFA500"

        $params = @{
            Name = $selectedRule.Name
        }

        # Enabled
        $params.Enabled = if ($cmbEnabled.SelectedIndex -eq 0) { "True" } else { "False" }

        # Action
        $params.Action = if ($cmbAction.SelectedIndex -eq 0) { "Allow" } else { "Block" }

        # Direction cannot be changed on existing rules, so we skip it

        # Profile
        $profileValues = @("Any", "Domain", "Private", "Public", "Domain, Private", "Domain, Public", "Private, Public")
        $params.Profile = $profileValues[$cmbProfile.SelectedIndex]

        Set-NetFirewallRule @params -ErrorAction Stop

        # Update port filter if protocol is TCP or UDP
        $protocol = @("Any", "TCP", "UDP", "ICMPv4", "ICMPv6")[$cmbProtocol.SelectedIndex]
        if ($protocol -in @("TCP", "UDP")) {
            $portParams = @{}
            if (-not [string]::IsNullOrWhiteSpace($txtLocalPort.Text) -and $txtLocalPort.Text -ne "Any") {
                $portParams.LocalPort = $txtLocalPort.Text
            }
            if (-not [string]::IsNullOrWhiteSpace($txtRemotePort.Text) -and $txtRemotePort.Text -ne "Any") {
                $portParams.RemotePort = $txtRemotePort.Text
            }

            if ($portParams.Count -gt 0) {
                Get-NetFirewallRule -Name $selectedRule.Name | Set-NetFirewallPortFilter @portParams -ErrorAction Stop
            }
        }

        Update-Status "Changes applied successfully" "#00FF00"
        [System.Windows.MessageBox]::Show(
            "Rule updated successfully!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )

        # Refresh
        Get-FirewallRules
    }
    catch {
        Update-Status "Failed to apply changes: $($_.Exception.Message)" "#FF0000"
        [System.Windows.MessageBox]::Show(
            "Failed to update rule.`n`nError: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Delete-SelectedRule {
    $selectedRules = @($dgRules.SelectedItems)

    if ($selectedRules.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select one or more rules to delete.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    $confirmResult = [System.Windows.MessageBox]::Show(
        "Are you sure you want to delete $($selectedRules.Count) rule(s)?`n`nThis action cannot be undone!",
        "Confirm Delete",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )

    if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
        $deleted = 0
        $failed = 0

        foreach ($rule in $selectedRules) {
            try {
                Remove-NetFirewallRule -Name $rule.Name -ErrorAction Stop
                $deleted++
            }
            catch {
                $failed++
            }
        }

        if ($failed -eq 0) {
            Update-Status "Deleted $deleted rule(s) successfully" "#00FF00"
        }
        else {
            Update-Status "Deleted $deleted rule(s), $failed failed" "#FFA500"
        }

        Get-FirewallRules
    }
}

function Find-DuplicateRules {
    Update-Status "Scanning for duplicate rules..." "#FFA500"
    $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    if (-not $Script:AllRules -or $Script:AllRules.Count -eq 0) {
        Update-Status "No rules loaded" "#FF0000"
        return
    }

    # Group by Program + Direction + Action + LocalPort
    $groups = @{}
    foreach ($rule in $Script:AllRules) {
        $prog = if ($rule.Program) { $rule.Program.ToLower() } else { "any" }
        $key = "$prog|$($rule.Direction)|$($rule.Action)|$($rule.LocalPort)"
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [System.Collections.Generic.List[PSObject]]::new()
        }
        $groups[$key].Add($rule)
    }

    # Find groups with more than one rule
    $duplicateGroups = @($groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })

    if ($duplicateGroups.Count -eq 0) {
        Update-Status "No duplicate rules found" "#00FF00"
        [System.Windows.MessageBox]::Show(
            "No duplicate rules found.`n`nAll rules have unique combinations of program, direction, action, and ports.",
            "No Duplicates",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        return
    }

    # Build report
    $totalDupes = 0
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("DUPLICATE FIREWALL RULES FOUND")
    [void]$sb.AppendLine("=" * 60)
    [void]$sb.AppendLine("")

    foreach ($group in $duplicateGroups) {
        $parts = $group.Key -split '\|'
        $totalDupes += ($group.Value.Count - 1)
        [void]$sb.AppendLine("--- Group ($($group.Value.Count) rules) ---")
        [void]$sb.AppendLine("  Program: $($parts[0])")
        [void]$sb.AppendLine("  Direction: $($parts[1])  Action: $($parts[2])  Port: $($parts[3])")
        foreach ($r in $group.Value) {
            [void]$sb.AppendLine("    - $($r.DisplayName) [Enabled: $($r.Enabled)]")
        }
        [void]$sb.AppendLine("")
    }

    [void]$sb.AppendLine("Summary: $($duplicateGroups.Count) duplicate group(s), $totalDupes extra rule(s)")

    # Collect all duplicate rules and show them in the grid
    $dupeRules = @()
    foreach ($group in $duplicateGroups) {
        $dupeRules += $group.Value
    }

    $dgRules.ItemsSource = $dupeRules
    $txtRuleCount.Text = "Duplicates: $($dupeRules.Count) / $($Script:AllRules.Count)"
    Update-Status "Found $($duplicateGroups.Count) duplicate groups ($totalDupes extra rules)" "#FFA500"

    # Show report dialog
    $reportWindow = New-Object System.Windows.Window
    $reportWindow.Title = "Duplicate Rules Report"
    $reportWindow.Width = 700
    $reportWindow.Height = 500
    $reportWindow.WindowStartupLocation = "CenterOwner"
    $reportWindow.Owner = $Window
    $reportWindow.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E)))

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)

    $rowDef1 = New-Object System.Windows.Controls.RowDefinition
    $rowDef1.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $grid.RowDefinitions.Add($rowDef1)
    $rowDef2 = New-Object System.Windows.Controls.RowDefinition
    $rowDef2.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($rowDef2)

    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Text = $sb.ToString()
    $textBox.IsReadOnly = $true
    $textBox.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26)))
    $textBox.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0)))
    $textBox.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $textBox.FontSize = 12
    $textBox.AcceptsReturn = $true
    $textBox.VerticalScrollBarVisibility = "Auto"
    $textBox.HorizontalScrollBarVisibility = "Auto"
    [System.Windows.Controls.Grid]::SetRow($textBox, 0)
    $grid.Children.Add($textBox)

    $closeBtn = New-Object System.Windows.Controls.Button
    $closeBtn.Content = "Close"
    $closeBtn.Width = 100
    $closeBtn.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    $closeBtn.HorizontalAlignment = "Right"
    $closeBtn.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4)))
    $closeBtn.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White))
    $closeBtn.Padding = New-Object System.Windows.Thickness(15, 8, 15, 8)
    $closeBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $closeBtn.Add_Click({ $reportWindow.Close() })
    [System.Windows.Controls.Grid]::SetRow($closeBtn, 1)
    $grid.Children.Add($closeBtn)

    $reportWindow.Content = $grid
    $reportWindow.ShowDialog() | Out-Null
}

function Show-QuickBlockMenu {
    $contextMenu = New-Object System.Windows.Controls.ContextMenu
    $contextMenu.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x2D, 0x2D, 0x30))
    $contextMenu.BorderBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x55, 0x55, 0x55))
    $contextMenu.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))

    $menuBlockProgram = New-Object System.Windows.Controls.MenuItem
    $menuBlockProgram.Header = "Block Program..."
    $menuBlockProgram.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $menuBlockProgram.Add_Click({
        $fileBrowser = New-Object Microsoft.Win32.OpenFileDialog
        $fileBrowser.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
        $fileBrowser.Title = "Select Program to Block"

        if ($fileBrowser.ShowDialog()) {
            $programPath = $fileBrowser.FileName
            $programName = [System.IO.Path]::GetFileNameWithoutExtension($programPath)

            try {
                # Block both inbound and outbound
                New-NetFirewallRule -DisplayName "Block $programName (Outbound)" `
                    -Direction Outbound -Action Block -Program $programPath `
                    -Profile Any -Enabled True `
                    -Description "Quick-blocked by FirewallManager" | Out-Null

                New-NetFirewallRule -DisplayName "Block $programName (Inbound)" `
                    -Direction Inbound -Action Block -Program $programPath `
                    -Profile Any -Enabled True `
                    -Description "Quick-blocked by FirewallManager" | Out-Null

                Update-Status "Blocked program: $programName (in+out)" "#00FF00"
                Get-FirewallRules
            }
            catch {
                Update-Status "Failed to block program: $($_.Exception.Message)" "#FF0000"
                [System.Windows.MessageBox]::Show(
                    "Failed to create firewall rule.`n`nError: $($_.Exception.Message)",
                    "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })
    $contextMenu.Items.Add($menuBlockProgram) | Out-Null

    $menuBlockPort = New-Object System.Windows.Controls.MenuItem
    $menuBlockPort.Header = "Block Port..."
    $menuBlockPort.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $menuBlockPort.Add_Click({
        # Input dialog for port
        $inputWindow = New-Object System.Windows.Window
        $inputWindow.Title = "Block Port"
        $inputWindow.Width = 350
        $inputWindow.Height = 200
        $inputWindow.WindowStartupLocation = "CenterOwner"
        $inputWindow.Owner = $Window
        $inputWindow.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E)))
        $inputWindow.ResizeMode = "NoResize"

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Margin = New-Object System.Windows.Thickness(20)

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = "Enter port number (e.g. 8080, 1234-1240):"
        $lbl.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0)))
        $lbl.FontSize = 13
        $lbl.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $sp.Children.Add($lbl) | Out-Null

        $txtPort = New-Object System.Windows.Controls.TextBox
        $txtPort.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x3C, 0x3C, 0x3C)))
        $txtPort.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0)))
        $txtPort.FontSize = 14
        $txtPort.Padding = New-Object System.Windows.Thickness(8, 6, 8, 6)
        $txtPort.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
        $sp.Children.Add($txtPort) | Out-Null

        $btnOK = New-Object System.Windows.Controls.Button
        $btnOK.Content = "Block Port"
        $btnOK.Width = 120
        $btnOK.HorizontalAlignment = "Right"
        $btnOK.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xD3, 0x2F, 0x2F)))
        $btnOK.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White))
        $btnOK.Padding = New-Object System.Windows.Thickness(15, 8, 15, 8)
        $btnOK.Cursor = [System.Windows.Input.Cursors]::Hand
        $btnOK.Add_Click({
            $port = $txtPort.Text.Trim()
            if ([string]::IsNullOrEmpty($port)) { return }

            try {
                New-NetFirewallRule -DisplayName "Block Port $port (TCP In)" `
                    -Direction Inbound -Action Block -Protocol TCP -LocalPort $port `
                    -Profile Any -Enabled True -Description "Quick-blocked by FirewallManager" | Out-Null
                New-NetFirewallRule -DisplayName "Block Port $port (TCP Out)" `
                    -Direction Outbound -Action Block -Protocol TCP -LocalPort $port `
                    -Profile Any -Enabled True -Description "Quick-blocked by FirewallManager" | Out-Null
                New-NetFirewallRule -DisplayName "Block Port $port (UDP In)" `
                    -Direction Inbound -Action Block -Protocol UDP -LocalPort $port `
                    -Profile Any -Enabled True -Description "Quick-blocked by FirewallManager" | Out-Null
                New-NetFirewallRule -DisplayName "Block Port $port (UDP Out)" `
                    -Direction Outbound -Action Block -Protocol UDP -LocalPort $port `
                    -Profile Any -Enabled True -Description "Quick-blocked by FirewallManager" | Out-Null

                $inputWindow.DialogResult = $true
                $inputWindow.Close()
                Update-Status "Blocked port $port (TCP+UDP, in+out)" "#00FF00"
                Get-FirewallRules
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Failed to block port.`n`nError: $($_.Exception.Message)",
                    "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
        $sp.Children.Add($btnOK) | Out-Null

        $inputWindow.Content = $sp
        $inputWindow.ShowDialog() | Out-Null
    })
    $contextMenu.Items.Add($menuBlockPort) | Out-Null

    $menuBlockIP = New-Object System.Windows.Controls.MenuItem
    $menuBlockIP.Header = "Block IP Address..."
    $menuBlockIP.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $menuBlockIP.Add_Click({
        # Input dialog for IP
        $inputWindow = New-Object System.Windows.Window
        $inputWindow.Title = "Block IP Address"
        $inputWindow.Width = 350
        $inputWindow.Height = 200
        $inputWindow.WindowStartupLocation = "CenterOwner"
        $inputWindow.Owner = $Window
        $inputWindow.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E)))
        $inputWindow.ResizeMode = "NoResize"

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Margin = New-Object System.Windows.Thickness(20)

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = "Enter IP address or CIDR range (e.g. 10.0.0.1, 192.168.0.0/24):"
        $lbl.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0)))
        $lbl.FontSize = 13
        $lbl.TextWrapping = "Wrap"
        $lbl.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $sp.Children.Add($lbl) | Out-Null

        $txtIP = New-Object System.Windows.Controls.TextBox
        $txtIP.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x3C, 0x3C, 0x3C)))
        $txtIP.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0)))
        $txtIP.FontSize = 14
        $txtIP.Padding = New-Object System.Windows.Thickness(8, 6, 8, 6)
        $txtIP.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
        $sp.Children.Add($txtIP) | Out-Null

        $btnOK = New-Object System.Windows.Controls.Button
        $btnOK.Content = "Block IP"
        $btnOK.Width = 120
        $btnOK.HorizontalAlignment = "Right"
        $btnOK.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xD3, 0x2F, 0x2F)))
        $btnOK.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White))
        $btnOK.Padding = New-Object System.Windows.Thickness(15, 8, 15, 8)
        $btnOK.Cursor = [System.Windows.Input.Cursors]::Hand
        $btnOK.Add_Click({
            $ip = $txtIP.Text.Trim()
            if ([string]::IsNullOrEmpty($ip)) { return }

            try {
                New-NetFirewallRule -DisplayName "Block IP $ip (Outbound)" `
                    -Direction Outbound -Action Block -RemoteAddress $ip `
                    -Profile Any -Enabled True -Description "Quick-blocked by FirewallManager" | Out-Null
                New-NetFirewallRule -DisplayName "Block IP $ip (Inbound)" `
                    -Direction Inbound -Action Block -RemoteAddress $ip `
                    -Profile Any -Enabled True -Description "Quick-blocked by FirewallManager" | Out-Null

                $inputWindow.DialogResult = $true
                $inputWindow.Close()
                Update-Status "Blocked IP: $ip (in+out)" "#00FF00"
                Get-FirewallRules
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Failed to block IP.`n`nError: $($_.Exception.Message)",
                    "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
        $sp.Children.Add($btnOK) | Out-Null

        $inputWindow.Content = $sp
        $inputWindow.ShowDialog() | Out-Null
    })
    $contextMenu.Items.Add($menuBlockIP) | Out-Null

    $contextMenu.IsOpen = $true
    $contextMenu.PlacementTarget = $btnQuickBlock
    $contextMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
}

# ============================================================
# Event Handlers
# ============================================================
$btnBackup.Add_Click({ Backup-FirewallRules })
$btnRestore.Add_Click({ Restore-FirewallRules })
$btnRefresh.Add_Click({ Get-FirewallRules })
$btnExportCSV.Add_Click({ Export-RulesToCSV })
$btnFindDuplicates.Add_Click({ Find-DuplicateRules })
$btnQuickBlock.Add_Click({ Show-QuickBlockMenu })
$btnSearch.Add_Click({ Search-Rules })
$btnClearSearch.Add_Click({
    $txtSearch.Text = ""
    $cmbGroupFilter.SelectedIndex = 0
    $dgRules.RowStyle = $null
    Search-Rules
})
$btnApplyEdit.Add_Click({ Apply-RuleChanges })
$btnDeleteRule.Add_Click({ Delete-SelectedRule })

$btnToggleStats.Add_Click({
    if ($statsPanel.Visibility -eq [System.Windows.Visibility]::Collapsed) {
        $statsPanel.Visibility = [System.Windows.Visibility]::Visible
        Update-Statistics
    }
    else {
        $statsPanel.Visibility = [System.Windows.Visibility]::Collapsed
    }
})

$cmbGroupFilter.Add_SelectionChanged({
    if ($cmbGroupFilter.SelectedItem -ne $null) {
        Search-Rules
    }
})

$dgRules.Add_SelectionChanged({ Update-EditPanel })

$txtSearch.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        Search-Rules
    }
})

# ============================================================
# Initialize
# ============================================================
    Write-Host "Launching GUI..." -ForegroundColor Green
    Write-Host "(Firewall rules will load after window appears)" -ForegroundColor Yellow

    # Load rules after window is shown
    $Window.Add_ContentRendered({
        Write-Host "Window rendered, now loading firewall rules..." -ForegroundColor Cyan
        Get-FirewallRules
        Write-Host "Firewall rules loaded." -ForegroundColor Green
    })

    # Show window
    $Window.ShowDialog() | Out-Null

    Write-Host "Application closed normally." -ForegroundColor Green
}
catch {
    Write-Host "`n============================================" -ForegroundColor Red
    Write-Host "ERROR OCCURRED" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "`nFull Error:" -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor Gray
    Write-Host "`nStack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host "============================================`n" -ForegroundColor Red
}
finally {
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
