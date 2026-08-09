<#
.SYNOPSIS
    FirewallRulesEditor - Offline Firewall Rules Editor
.DESCRIPTION
    A standalone tool for importing, editing, and exporting firewall rules from backups.
    Does NOT modify your system's actual firewall - purely an offline editor.
.NOTES
    Author: Matt
    Version: 1.3.0
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# XAML GUI Definition
# ============================================================
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Firewall Rules Editor v1.3.0 (Offline) - 0 rules"
        Height="800" Width="1200"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E1E"
        ResizeMode="CanResizeWithGrip">
    <Window.Resources>
        <SolidColorBrush x:Key="ComboBoxBackground" Color="#3C3C3C"/>
        <SolidColorBrush x:Key="ComboBoxBorder" Color="#555555"/>
        <SolidColorBrush x:Key="ComboBoxForeground" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="ComboBoxHoverBackground" Color="#4A4A4A"/>
        <SolidColorBrush x:Key="ComboBoxDropdownBackground" Color="#2D2D30"/>
        <SolidColorBrush x:Key="ComboBoxItemHover" Color="#3E3E42"/>
        <SolidColorBrush x:Key="ComboBoxItemSelected" Color="#0078D4"/>

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
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <DataTrigger Binding="{Binding Selected}" Value="True">
                    <Setter Property="Background" Value="#1B3E1B"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#B0B0B0"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>

        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" Background="{StaticResource ComboBoxBackground}"
                        BorderBrush="{StaticResource ComboBoxBorder}" BorderThickness="1" CornerRadius="3"/>
                <Path x:Name="Arrow" Grid.Column="1" Fill="#E0E0E0" HorizontalAlignment="Center"
                      VerticalAlignment="Center" Data="M 0 0 L 6 6 L 12 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxHoverBackground}"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxItemHover}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource ComboBoxItemSelected}"/>
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
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton" Template="{StaticResource ComboBoxToggleButton}"
                                          Focusable="False"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              Margin="10,3,30,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid Name="DropDown" SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border Background="{StaticResource ComboBoxDropdownBackground}"
                                            BorderThickness="1" BorderBrush="{StaticResource ComboBoxBorder}" CornerRadius="3">
                                        <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="Firewall Rules Editor" FontSize="28" FontWeight="Bold" Foreground="#0078D4"/>
            <TextBlock Text="Import, Edit, and Export Firewall Rules (Offline - Does NOT affect your system firewall)"
                       FontSize="13" Foreground="#808080"/>
        </StackPanel>

        <!-- Import/Export Buttons -->
        <WrapPanel Grid.Row="1" Margin="0,0,0,10">
            <Button x:Name="btnImportBackup" Content="Import .fwbackup" Width="140"/>
            <Button x:Name="btnImportCSV" Content="Import CSV" Width="120"/>
            <Button x:Name="btnMergeBackup" Content="Merge .fwbackup" Style="{StaticResource WarningButton}" Width="140"/>
            <Button x:Name="btnExportBackup" Content="Export Selected to .fwbackup" Style="{StaticResource SuccessButton}" Width="200"/>
            <Button x:Name="btnExportCSV" Content="Export Selected to CSV" Style="{StaticResource SuccessButton}" Width="180"/>
            <Button x:Name="btnShowChanges" Content="Show Changes" Width="120"/>
            <Button x:Name="btnCompareBackups" Content="Compare Backups" Width="140"/>
            <Button x:Name="btnCompareFleet" Content="Fleet Compare" Width="125"/>
            <Button x:Name="btnTemplates" Content="Templates" Width="105"/>
            <Button x:Name="btnExportPolicy" Content="Export Policy" Width="120"/>
            <Button x:Name="btnClearAll" Content="Clear All" Style="{StaticResource DangerButton}" Width="100"/>
        </WrapPanel>

        <!-- Selection and Filter Tools -->
        <WrapPanel Grid.Row="2" Margin="0,0,0,10" VerticalAlignment="Center">
            <Button x:Name="btnSelectAll" Content="Select All" Style="{StaticResource WarningButton}" Width="100"/>
            <Button x:Name="btnSelectNone" Content="Select None" Style="{StaticResource WarningButton}" Width="110"/>
            <Button x:Name="btnInvertSelection" Content="Invert Selection" Style="{StaticResource WarningButton}" Width="130"/>
            <Button x:Name="btnSelectFiltered" Content="Select Filtered" Style="{StaticResource WarningButton}" Width="120"/>
            <TextBox x:Name="txtSearch" Width="250" Margin="20,0,5,0" VerticalAlignment="Center"
                     ToolTip="Search by name, program, port, etc."/>
            <CheckBox x:Name="chkRegexSearch" Content="Regex" Margin="5,0,0,0" VerticalAlignment="Center"/>
            <Button x:Name="btnSearch" Content="Filter" Width="80"/>
            <Button x:Name="btnClearSearch" Content="Clear Filter" Width="100"/>
            <Button x:Name="btnDeleteSelected" Content="Delete Selected" Style="{StaticResource DangerButton}" Width="130" Margin="20,0,0,0"/>
            <Button x:Name="btnRefreshCounts" Content="Refresh Count" Width="110" Margin="10,0,0,0"/>
        </WrapPanel>

        <!-- Rules DataGrid with Checkbox -->
        <DataGrid x:Name="dgRules" Grid.Row="3"
                  AutoGenerateColumns="False"
                  IsReadOnly="False"
                  SelectionMode="Extended"
                  CanUserAddRows="False"
                  CanUserDeleteRows="False"
                  CanUserReorderColumns="True"
                  CanUserSortColumns="True">
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="Export" Width="60">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="200" IsReadOnly="True"/>
                <DataGridComboBoxColumn Header="Direction" SelectedItemBinding="{Binding Direction}" Width="90">
                    <DataGridComboBoxColumn.ElementStyle>
                        <Style TargetType="ComboBox">
                            <Setter Property="Background" Value="#3C3C3C"/>
                            <Setter Property="Foreground" Value="#E0E0E0"/>
                        </Style>
                    </DataGridComboBoxColumn.ElementStyle>
                    <DataGridComboBoxColumn.ItemsSource>
                        <x:Array Type="sys:String" xmlns:sys="clr-namespace:System;assembly=mscorlib">
                            <sys:String>Inbound</sys:String>
                            <sys:String>Outbound</sys:String>
                        </x:Array>
                    </DataGridComboBoxColumn.ItemsSource>
                </DataGridComboBoxColumn>
                <DataGridComboBoxColumn Header="Action" SelectedItemBinding="{Binding Action}" Width="80">
                    <DataGridComboBoxColumn.ItemsSource>
                        <x:Array Type="sys:String" xmlns:sys="clr-namespace:System;assembly=mscorlib">
                            <sys:String>Allow</sys:String>
                            <sys:String>Block</sys:String>
                        </x:Array>
                    </DataGridComboBoxColumn.ItemsSource>
                </DataGridComboBoxColumn>
                <DataGridComboBoxColumn Header="Enabled" SelectedItemBinding="{Binding Enabled}" Width="80">
                    <DataGridComboBoxColumn.ItemsSource>
                        <x:Array Type="sys:String" xmlns:sys="clr-namespace:System;assembly=mscorlib">
                            <sys:String>True</sys:String>
                            <sys:String>False</sys:String>
                        </x:Array>
                    </DataGridComboBoxColumn.ItemsSource>
                </DataGridComboBoxColumn>
                <DataGridComboBoxColumn Header="Profile" SelectedItemBinding="{Binding Profile}" Width="120">
                    <DataGridComboBoxColumn.ItemsSource>
                        <x:Array Type="sys:String" xmlns:sys="clr-namespace:System;assembly=mscorlib">
                            <sys:String>Any</sys:String>
                            <sys:String>Domain</sys:String>
                            <sys:String>Private</sys:String>
                            <sys:String>Public</sys:String>
                            <sys:String>Domain, Private</sys:String>
                            <sys:String>Domain, Public</sys:String>
                            <sys:String>Private, Public</sys:String>
                            <sys:String>Domain, Private, Public</sys:String>
                        </x:Array>
                    </DataGridComboBoxColumn.ItemsSource>
                </DataGridComboBoxColumn>
                <DataGridTextColumn Header="Protocol" Binding="{Binding Protocol}" Width="80"/>
                <DataGridTextColumn Header="Local Port" Binding="{Binding LocalPort}" Width="100"/>
                <DataGridTextColumn Header="Remote Port" Binding="{Binding RemotePort}" Width="100"/>
                <DataGridTextColumn Header="Program" Binding="{Binding Program}" Width="300"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Add New Rule Panel -->
        <GroupBox Grid.Row="4" Header="Add New Rule" Margin="0,10,0,0" Foreground="#B0B0B0" BorderBrush="#3C3C3C">
            <Grid Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="5">
                    <Label Content="Display Name"/>
                    <TextBox x:Name="txtNewName"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Margin="5">
                    <Label Content="Direction"/>
                    <ComboBox x:Name="cmbNewDirection">
                        <ComboBoxItem Content="Inbound" IsSelected="True"/>
                        <ComboBoxItem Content="Outbound"/>
                    </ComboBox>
                </StackPanel>
                <StackPanel Grid.Column="2" Margin="5">
                    <Label Content="Action"/>
                    <ComboBox x:Name="cmbNewAction">
                        <ComboBoxItem Content="Allow" IsSelected="True"/>
                        <ComboBoxItem Content="Block"/>
                    </ComboBox>
                </StackPanel>
                <StackPanel Grid.Column="3" Margin="5">
                    <Label Content="Protocol"/>
                    <ComboBox x:Name="cmbNewProtocol">
                        <ComboBoxItem Content="Any" IsSelected="True"/>
                        <ComboBoxItem Content="TCP"/>
                        <ComboBoxItem Content="UDP"/>
                    </ComboBox>
                </StackPanel>
                <StackPanel Grid.Column="4" Margin="5">
                    <Label Content="Local Port"/>
                    <TextBox x:Name="txtNewPort" Text="Any"/>
                </StackPanel>
                <StackPanel Grid.Column="5" Margin="5" VerticalAlignment="Bottom">
                    <Button x:Name="btnAddRule" Content="Add Rule" Style="{StaticResource SuccessButton}" Width="120"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- Status Bar -->
        <Border Grid.Row="5" Background="#252526" Margin="0,10,0,0" Padding="10,8" CornerRadius="4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Text="Ready - Import a backup or CSV to begin" Foreground="#808080" VerticalAlignment="Center"/>
                <TextBlock x:Name="txtSelectedCount" Grid.Column="1" Text="Selected: 0" Foreground="#00FF00" VerticalAlignment="Center" Margin="0,0,20,0"/>
                <TextBlock x:Name="txtRuleCount" Grid.Column="2" Text="Total: 0" Foreground="#808080" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ============================================================
# Load XAML
# ============================================================
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# codex-branding:start
                try {
                    $brandingIconPath = Join-Path $PSScriptRoot 'icon.ico'
                    if (Test-Path $brandingIconPath) {
                        $Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.Uri($brandingIconPath)))
                    }
                } catch {
                }
                # codex-branding:end
# Get controls
$btnImportBackup = $Window.FindName("btnImportBackup")
$btnImportCSV = $Window.FindName("btnImportCSV")
$btnMergeBackup = $Window.FindName("btnMergeBackup")
$btnExportBackup = $Window.FindName("btnExportBackup")
$btnExportCSV = $Window.FindName("btnExportCSV")
$btnShowChanges = $Window.FindName("btnShowChanges")
$btnCompareBackups = $Window.FindName("btnCompareBackups")
$btnCompareFleet = $Window.FindName("btnCompareFleet")
$btnTemplates = $Window.FindName("btnTemplates")
$btnExportPolicy = $Window.FindName("btnExportPolicy")
$btnClearAll = $Window.FindName("btnClearAll")
$btnSelectAll = $Window.FindName("btnSelectAll")
$btnSelectNone = $Window.FindName("btnSelectNone")
$btnInvertSelection = $Window.FindName("btnInvertSelection")
$btnSelectFiltered = $Window.FindName("btnSelectFiltered")
$btnDeleteSelected = $Window.FindName("btnDeleteSelected")
$btnRefreshCounts = $Window.FindName("btnRefreshCounts")
$btnSearch = $Window.FindName("btnSearch")
$btnClearSearch = $Window.FindName("btnClearSearch")
$chkRegexSearch = $Window.FindName("chkRegexSearch")
$btnAddRule = $Window.FindName("btnAddRule")
$txtSearch = $Window.FindName("txtSearch")
$txtStatus = $Window.FindName("txtStatus")
$txtSelectedCount = $Window.FindName("txtSelectedCount")
$txtRuleCount = $Window.FindName("txtRuleCount")
$dgRules = $Window.FindName("dgRules")
$txtNewName = $Window.FindName("txtNewName")
$cmbNewDirection = $Window.FindName("cmbNewDirection")
$cmbNewAction = $Window.FindName("cmbNewAction")
$cmbNewProtocol = $Window.FindName("cmbNewProtocol")
$txtNewPort = $Window.FindName("txtNewPort")

# Global variables
$Script:AllRules = [System.Collections.Generic.List[PSObject]]::new()
$Script:FilteredView = $null
$Script:OriginalRules = $null  # Snapshot at import time for diff
$Script:CurrentBackupDate = $null

# ============================================================
# Rule Class using PowerShell class (avoids C# assembly issues)
# ============================================================
class FirewallRuleItem {
    [bool]$Selected
    [string]$Name
    [string]$DisplayName
    [string]$Description
    [string]$Direction
    [string]$Action
    [string]$Enabled
    [string]$Profile
    [string]$Protocol
    [string]$LocalPort
    [string]$RemotePort
    [string]$Program
}

# ============================================================
# Functions
# ============================================================
function Update-Status {
    param([string]$Message)
    $txtStatus.Text = $Message
}

function Update-TitleBar {
    $count = $Script:AllRules.Count
    $Window.Title = "Firewall Rules Editor v1.3.0 (Offline) - $count rules"
}

function Update-Counts {
    $total = $Script:AllRules.Count
    $selected = ($Script:AllRules | Where-Object { $_.Selected }).Count
    $txtRuleCount.Text = "Total: $total"
    $txtSelectedCount.Text = "Selected: $selected"
    Update-TitleBar
}

function New-RuleObject {
    param(
        [string]$Name = "",
        [string]$DisplayName = "",
        [string]$Description = "",
        [string]$Direction = "Inbound",
        [string]$Action = "Allow",
        [string]$Enabled = "True",
        [string]$Profile = "Any",
        [string]$Protocol = "Any",
        [string]$LocalPort = "Any",
        [string]$RemotePort = "Any",
        [string]$Program = "Any",
        [bool]$Selected = $false
    )

    $rule = [FirewallRuleItem]::new()
    $rule.Name = $Name
    $rule.DisplayName = $DisplayName
    $rule.Description = $Description
    $rule.Direction = $Direction
    $rule.Action = $Action
    $rule.Enabled = $Enabled
    $rule.Profile = $Profile
    $rule.Protocol = $Protocol
    $rule.LocalPort = $LocalPort
    $rule.RemotePort = $RemotePort
    $rule.Program = $Program
    $rule.Selected = $Selected

    return $rule
}

function Take-Snapshot {
    # Deep-copy current rules for later diff comparison
    $Script:OriginalRules = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($r in $Script:AllRules) {
        $copy = New-RuleObject `
            -Name $r.Name -DisplayName $r.DisplayName -Description $r.Description `
            -Direction $r.Direction -Action $r.Action -Enabled $r.Enabled `
            -Profile $r.Profile -Protocol $r.Protocol -LocalPort $r.LocalPort `
            -RemotePort $r.RemotePort -Program $r.Program -Selected $r.Selected
        $Script:OriginalRules.Add($copy)
    }
}

function Import-FWBackup {
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $openDialog.Title = "Import Firewall Backup"

    if ($openDialog.ShowDialog()) {
        try {
            Update-Status "Importing backup..."
            $content = Get-Content $openDialog.FileName -Raw -Encoding UTF8
            $backup = $content | ConvertFrom-Json

            $Script:AllRules = [System.Collections.Generic.List[PSObject]]::new()
            $Script:FilteredView = $null
            $Script:CurrentBackupDate = if ($backup.BackupDate) { [string]$backup.BackupDate } else { $null }

            $ruleData = if ($backup.RuleDetails) { $backup.RuleDetails } else { $backup }

            foreach ($r in $ruleData) {
                $rule = New-RuleObject `
                    -Name $r.Name `
                    -DisplayName $r.DisplayName `
                    -Description $r.Description `
                    -Direction $r.Direction `
                    -Action $r.Action `
                    -Enabled $r.Enabled `
                    -Profile $r.Profile `
                    -Protocol $r.Protocol `
                    -LocalPort $r.LocalPort `
                    -RemotePort $r.RemotePort `
                    -Program $r.Program `
                    -Selected $false

                $Script:AllRules.Add($rule)
            }

            $dgRules.ItemsSource = $Script:AllRules
            Update-Counts
            Update-Status "Imported $($Script:AllRules.Count) rules from backup"

            # Take snapshot for diff tracking
            Take-Snapshot

            if ($backup.BackupDate) {
                [System.Windows.MessageBox]::Show(
                    "Backup imported successfully!`n`nBackup Date: $($backup.BackupDate)`nComputer: $($backup.ComputerName)`nRules: $($Script:AllRules.Count)",
                    "Import Complete",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
            }
        }
        catch {
            Update-Status "Import failed: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Failed to import backup.`n`nError: $($_.Exception.Message)",
                "Import Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
}

function Import-CSV {
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openDialog.Title = "Import CSV"

    if ($openDialog.ShowDialog()) {
        try {
            Update-Status "Importing CSV..."
            $csvData = Import-Csv $openDialog.FileName -Encoding UTF8

            $Script:AllRules = [System.Collections.Generic.List[PSObject]]::new()
            $Script:FilteredView = $null
            $Script:CurrentBackupDate = $null

            foreach ($r in $csvData) {
                $rule = New-RuleObject `
                    -Name $(if ($r.Name) { $r.Name } else { [guid]::NewGuid().ToString() }) `
                    -DisplayName $(if ($r.DisplayName) { $r.DisplayName } else { $r.Name }) `
                    -Description $(if ($r.Description) { $r.Description } else { "" }) `
                    -Direction $(if ($r.Direction) { $r.Direction } else { "Inbound" }) `
                    -Action $(if ($r.Action) { $r.Action } else { "Allow" }) `
                    -Enabled $(if ($r.Enabled) { $r.Enabled } else { "True" }) `
                    -Profile $(if ($r.Profile) { $r.Profile } else { "Any" }) `
                    -Protocol $(if ($r.Protocol) { $r.Protocol } else { "Any" }) `
                    -LocalPort $(if ($r.LocalPort) { $r.LocalPort } else { "Any" }) `
                    -RemotePort $(if ($r.RemotePort) { $r.RemotePort } else { "Any" }) `
                    -Program $(if ($r.Program) { $r.Program } else { "Any" }) `
                    -Selected $false

                $Script:AllRules.Add($rule)
            }

            $dgRules.ItemsSource = $Script:AllRules
            Update-Counts
            Update-Status "Imported $($Script:AllRules.Count) rules from CSV"

            # Take snapshot for diff tracking
            Take-Snapshot
        }
        catch {
            Update-Status "Import failed: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Failed to import CSV.`n`nError: $($_.Exception.Message)",
                "Import Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
}

function Get-RuleMergeDifferences {
    param(
        [object]$Current,
        [object]$Imported
    )

    $properties = @("DisplayName", "Description", "Direction", "Action", "Enabled", "Profile", "Protocol", "LocalPort", "RemotePort", "Program")
    foreach ($property in $properties) {
        $currentValue = [string]$Current.$property
        $importedValue = [string]$Imported.$property
        if ($currentValue -ne $importedValue) {
            "  $property : $currentValue -> $importedValue"
        }
    }
}

function Show-MergeStrategyPicker {
    param(
        [int]$ConflictCount,
        [string]$ImportedDate,
        [string]$CurrentDate
    )

    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Choose Merge Conflict Strategy"
    $dialog.Width = 560
    $dialog.Height = 300
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $dialog.Tag = $null

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = New-Object System.Windows.Thickness(20)
    $heading = New-Object System.Windows.Controls.TextBlock
    $heading.Text = "$ConflictCount conflicting rule(s) found"
    $heading.FontSize = 18
    $heading.FontWeight = "Bold"
    $heading.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $heading.Margin = New-Object System.Windows.Thickness(0, 0, 0, 12)
    $panel.Children.Add($heading) | Out-Null

    $details = New-Object System.Windows.Controls.TextBlock
    $details.Text = "Imported backup: $ImportedDate`nCurrent backup: $CurrentDate`n`nChoose how conflicting rules should be applied. Manual conflict asks for a decision on each rule."
    $details.TextWrapping = "Wrap"
    $details.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $details.Margin = New-Object System.Windows.Thickness(0, 0, 0, 12)
    $panel.Children.Add($details) | Out-Null

    $strategyCombo = New-Object System.Windows.Controls.ComboBox
    foreach ($strategy in @("Prefer newer", "Prefer imported", "Manual conflict")) {
        $strategyCombo.Items.Add($strategy) | Out-Null
    }
    $strategyCombo.SelectedIndex = 0
    $strategyCombo.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($strategyCombo) | Out-Null

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Width = 100
    $cancel.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($cancel) | Out-Null
    $apply = New-Object System.Windows.Controls.Button
    $apply.Content = "Continue Merge"
    $apply.Width = 130
    $apply.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $apply.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $apply.Add_Click({ $dialog.Tag = [string]$strategyCombo.SelectedItem; $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($apply) | Out-Null
    $panel.Children.Add($buttons) | Out-Null

    $dialog.Content = $panel
    $dialog.ShowDialog() | Out-Null
    return [string]$dialog.Tag
}

function Resolve-MergeConflict {
    param(
        [object]$Current,
        [object]$Imported,
        [string[]]$Changes
    )

    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Resolve Rule Conflict"
    $dialog.Width = 760
    $dialog.Height = 460
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $dialog.Tag = "Skip"

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    $contentRow = New-Object System.Windows.Controls.RowDefinition
    $contentRow.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.RowDefinitions.Add($contentRow)
    $buttonRow = New-Object System.Windows.Controls.RowDefinition
    $buttonRow.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($buttonRow)

    $text = New-Object System.Windows.Controls.TextBox
    $text.IsReadOnly = $true
    $text.AcceptsReturn = $true
    $text.VerticalScrollBarVisibility = "Auto"
    $text.HorizontalScrollBarVisibility = "Auto"
    $text.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26))
    $text.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $text.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $text.Text = @(
        "CONFLICT: $($Current.DisplayName)"
        "Rule name: $($Current.Name)"
        ""
        "CURRENT VALUES"
        "  Action=$($Current.Action), Enabled=$($Current.Enabled), Profile=$($Current.Profile), Protocol=$($Current.Protocol)"
        "  Ports=$($Current.LocalPort)->$($Current.RemotePort), Program=$($Current.Program)"
        ""
        "IMPORTED VALUES"
        "  Action=$($Imported.Action), Enabled=$($Imported.Enabled), Profile=$($Imported.Profile), Protocol=$($Imported.Protocol)"
        "  Ports=$($Imported.LocalPort)->$($Imported.RemotePort), Program=$($Imported.Program)"
        ""
        "CHANGES"
    ) -join [Environment]::NewLine
    $text.Text += [Environment]::NewLine + ($Changes -join [Environment]::NewLine)
    [System.Windows.Controls.Grid]::SetRow($text, 0)
    $grid.Children.Add($text) | Out-Null

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    foreach ($choice in @(
        @{ Label = "Keep Current"; Value = "Keep current"; Color = @(0x55, 0x55, 0x55); Width = 120 }
        @{ Label = "Use Imported"; Value = "Use imported"; Color = @(0x38, 0x8E, 0x3C); Width = 120 }
        @{ Label = "Skip"; Value = "Skip"; Color = @(0xD3, 0x2F, 0x2F); Width = 90 }
    )) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = $choice.Label
        $button.Width = $choice.Width
        $button.Margin = New-Object System.Windows.Thickness(5, 10, 0, 0)
        $rgb = $choice.Color
        $button.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb($rgb[0], $rgb[1], $rgb[2]))
        $button.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
        $value = $choice.Value
        $button.Add_Click({ $dialog.Tag = $value; $dialog.Close() }.GetNewClosure())
        $buttons.Children.Add($button) | Out-Null
    }
    [System.Windows.Controls.Grid]::SetRow($buttons, 1)
    $grid.Children.Add($buttons) | Out-Null

    $dialog.Content = $grid
    $dialog.ShowDialog() | Out-Null
    return [string]$dialog.Tag
}

function Merge-FWBackup {
    if ($Script:AllRules.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No rules loaded. Import a backup first, then merge a second one.",
            "No Rules Loaded",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $openDialog.Title = "Merge Firewall Backup"

    if (-not $openDialog.ShowDialog()) {
        return
    }

    try {
        Update-Status "Reading backup for merge..."
        $incomingData = Read-BackupRuleSet -Path $openDialog.FileName
        $incomingRules = @($incomingData.Rules)
        $existingByName = @{}
        foreach ($rule in $Script:AllRules) {
            $existingByName[$rule.Name] = $rule
        }

        $conflicts = @()
        foreach ($incomingRule in $incomingRules) {
            if ($existingByName.ContainsKey($incomingRule.Name)) {
                $currentRule = $existingByName[$incomingRule.Name]
                if ((Get-RuleMergeDifferences -Current $currentRule -Imported $incomingRule).Count -gt 0) {
                    $conflicts += $incomingRule
                }
            }
        }

        $strategy = "Prefer imported"
        if ($conflicts.Count -gt 0) {
            $strategy = Show-MergeStrategyPicker -ConflictCount $conflicts.Count `
                -ImportedDate $incomingData.Backup.BackupDate -CurrentDate $Script:CurrentBackupDate
            if ([string]::IsNullOrWhiteSpace($strategy)) {
                Update-Status "Merge cancelled"
                return
            }
        }

        $importedDate = [datetime]::MinValue
        $currentDate = [datetime]::MinValue
        $parsedDate = $null
        if ($incomingData.Backup.BackupDate -and [datetime]::TryParse([string]$incomingData.Backup.BackupDate, [ref]$parsedDate)) {
            $importedDate = $parsedDate
        }
        $parsedDate = $null
        if ($Script:CurrentBackupDate -and [datetime]::TryParse([string]$Script:CurrentBackupDate, [ref]$parsedDate)) {
            $currentDate = $parsedDate
        }

        $added = 0
        $replaced = 0
        $skipped = 0
        $manualSkipped = 0
        foreach ($incomingRule in $incomingRules) {
            if (-not $existingByName.ContainsKey($incomingRule.Name)) {
                $incomingRule.Selected = $true
                $Script:AllRules.Add($incomingRule)
                $existingByName[$incomingRule.Name] = $incomingRule
                $added++
                continue
            }

            $currentRule = $existingByName[$incomingRule.Name]
            $differences = @(Get-RuleMergeDifferences -Current $currentRule -Imported $incomingRule)
            if ($differences.Count -eq 0) {
                $skipped++
                continue
            }

            $decision = $strategy
            if ($strategy -eq "Prefer newer") {
                $decision = if ($importedDate -gt $currentDate) { "Use imported" } else { "Keep current" }
            }
            elseif ($strategy -eq "Prefer imported") {
                $decision = "Use imported"
            }
            else {
                $decision = Resolve-MergeConflict -Current $currentRule -Imported $incomingRule -Changes $differences
            }

            if ($decision -eq "Use imported") {
                $incomingRule.Selected = $currentRule.Selected
                $index = $Script:AllRules.IndexOf($currentRule)
                if ($index -ge 0) {
                    $Script:AllRules[$index] = $incomingRule
                    $existingByName[$incomingRule.Name] = $incomingRule
                    $replaced++
                }
            }
            elseif ($decision -eq "Keep current") {
                $skipped++
            }
            else {
                $manualSkipped++
            }
        }

        if ($importedDate -gt $currentDate -and $importedDate -ne [datetime]::MinValue) {
            $Script:CurrentBackupDate = $incomingData.Backup.BackupDate
        }
        $Script:FilteredView = $null
        $dgRules.ItemsSource = $null
        $dgRules.ItemsSource = $Script:AllRules
        Update-Counts
        Update-Status "Merge complete: $added added, $replaced replaced, $skipped kept, $manualSkipped manually skipped"

        [System.Windows.MessageBox]::Show(
            "Merge complete!`n`nNew rules added: $added`nConflicts replaced: $replaced`nExisting rules kept: $skipped`nManual conflicts skipped: $manualSkipped`nTotal rules: $($Script:AllRules.Count)",
            "Merge Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
    catch {
        Update-Status "Merge failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to merge backup.`n`nError: $($_.Exception.Message)",
            "Merge Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Show-Changes {
    if ($null -eq $Script:OriginalRules) {
        [System.Windows.MessageBox]::Show(
            "No original snapshot available.`nImport a backup first to enable change tracking.",
            "No Snapshot",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    # Build lookup of original rules by Name
    $origByName = @{}
    foreach ($r in $Script:OriginalRules) {
        $origByName[$r.Name] = $r
    }

    # Build lookup of current rules by Name
    $currByName = @{}
    foreach ($r in $Script:AllRules) {
        $currByName[$r.Name] = $r
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("CHANGE REPORT")
    [void]$sb.AppendLine("=" * 60)
    [void]$sb.AppendLine("")

    # Added rules (in current but not in original)
    $addedRules = @()
    foreach ($r in $Script:AllRules) {
        if (-not $origByName.ContainsKey($r.Name)) {
            $addedRules += $r
        }
    }

    # Deleted rules (in original but not in current)
    $deletedRules = @()
    foreach ($r in $Script:OriginalRules) {
        if (-not $currByName.ContainsKey($r.Name)) {
            $deletedRules += $r
        }
    }

    # Modified rules (same name but different properties)
    $modifiedRules = @()
    $propNames = @("DisplayName", "Direction", "Action", "Enabled", "Profile", "Protocol", "LocalPort", "RemotePort", "Program")
    foreach ($r in $Script:AllRules) {
        if ($origByName.ContainsKey($r.Name)) {
            $orig = $origByName[$r.Name]
            $changes = @()
            foreach ($prop in $propNames) {
                $oldVal = $orig.$prop
                $newVal = $r.$prop
                if ($oldVal -ne $newVal) {
                    $changes += "  $prop : $oldVal -> $newVal"
                }
            }
            if ($changes.Count -gt 0) {
                $modifiedRules += @{ Rule = $r; Changes = $changes }
            }
        }
    }

    # Build output
    [void]$sb.AppendLine("ADDED RULES ($($addedRules.Count)):")
    [void]$sb.AppendLine("-" * 40)
    if ($addedRules.Count -eq 0) {
        [void]$sb.AppendLine("  (none)")
    }
    else {
        foreach ($r in $addedRules) {
            [void]$sb.AppendLine("  + $($r.DisplayName) [$($r.Direction) $($r.Action)]")
        }
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("DELETED RULES ($($deletedRules.Count)):")
    [void]$sb.AppendLine("-" * 40)
    if ($deletedRules.Count -eq 0) {
        [void]$sb.AppendLine("  (none)")
    }
    else {
        foreach ($r in $deletedRules) {
            [void]$sb.AppendLine("  - $($r.DisplayName) [$($r.Direction) $($r.Action)]")
        }
    }
    [void]$sb.AppendLine("")

    [void]$sb.AppendLine("MODIFIED RULES ($($modifiedRules.Count)):")
    [void]$sb.AppendLine("-" * 40)
    if ($modifiedRules.Count -eq 0) {
        [void]$sb.AppendLine("  (none)")
    }
    else {
        foreach ($m in $modifiedRules) {
            [void]$sb.AppendLine("  * $($m.Rule.DisplayName)")
            foreach ($c in $m.Changes) {
                [void]$sb.AppendLine("    $c")
            }
        }
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Summary: +$($addedRules.Count) added, -$($deletedRules.Count) deleted, ~$($modifiedRules.Count) modified")

    # Show in dialog
    $reportWindow = New-Object System.Windows.Window
    $reportWindow.Title = "Changes Since Import"
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

function Show-EditorReportWindow {
    param(
        [string]$Title,
        [string]$Text,
        [int]$Width = 820,
        [int]$Height = 600
    )

    $reportWindow = New-Object System.Windows.Window
    $reportWindow.Title = $Title
    $reportWindow.Width = $Width
    $reportWindow.Height = $Height
    $reportWindow.WindowStartupLocation = "CenterOwner"
    $reportWindow.Owner = $Window
    $reportWindow.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    $contentRow = New-Object System.Windows.Controls.RowDefinition
    $contentRow.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.RowDefinitions.Add($contentRow)
    $buttonRow = New-Object System.Windows.Controls.RowDefinition
    $buttonRow.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($buttonRow)

    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Text = $Text
    $textBox.IsReadOnly = $true
    $textBox.AcceptsReturn = $true
    $textBox.TextWrapping = "NoWrap"
    $textBox.VerticalScrollBarVisibility = "Auto"
    $textBox.HorizontalScrollBarVisibility = "Auto"
    $textBox.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26))
    $textBox.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $textBox.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $textBox.FontSize = 12
    [System.Windows.Controls.Grid]::SetRow($textBox, 0)
    $grid.Children.Add($textBox) | Out-Null

    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Width = 100
    $closeButton.HorizontalAlignment = "Right"
    $closeButton.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    $closeButton.Add_Click({ $reportWindow.Close() }.GetNewClosure())
    [System.Windows.Controls.Grid]::SetRow($closeButton, 1)
    $grid.Children.Add($closeButton) | Out-Null

    $reportWindow.Content = $grid
    $reportWindow.ShowDialog() | Out-Null
}

function Read-BackupRuleSet {
    param([string]$Path)

    $backup = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $ruleData = if ($backup.RuleDetails) { @($backup.RuleDetails) } else { @($backup) }
    $rules = New-Object System.Collections.Generic.List[PSObject]
    foreach ($r in $ruleData) {
        $rules.Add((New-RuleObject `
            -Name $r.Name `
            -DisplayName $r.DisplayName `
            -Description $r.Description `
            -Direction $r.Direction `
            -Action $r.Action `
            -Enabled $r.Enabled `
            -Profile $(if ($r.Profile) { $r.Profile } else { "Any" }) `
            -Protocol $(if ($r.Protocol) { $r.Protocol } else { "Any" }) `
            -LocalPort $(if ($r.LocalPort) { $r.LocalPort } else { "Any" }) `
            -RemotePort $(if ($r.RemotePort) { $r.RemotePort } else { "Any" }) `
            -Program $(if ($r.Program) { $r.Program } else { "Any" }) `
            -Selected $false))
    }
    return [PSCustomObject]@{
        Backup = $backup
        Rules = $rules
    }
}

function Compare-BackupRuleSets {
    param(
        [string]$LeftPath,
        [string]$RightPath
    )

    $leftData = Read-BackupRuleSet -Path $LeftPath
    $rightData = Read-BackupRuleSet -Path $RightPath
    $leftByName = @{}
    $rightByName = @{}
    foreach ($rule in $leftData.Rules) { $leftByName[$rule.Name] = $rule }
    foreach ($rule in $rightData.Rules) { $rightByName[$rule.Name] = $rule }

    $added = @($rightData.Rules | Where-Object { -not $leftByName.ContainsKey($_.Name) })
    $removed = @($leftData.Rules | Where-Object { -not $rightByName.ContainsKey($_.Name) })
    $modified = New-Object System.Collections.Generic.List[PSObject]
    $properties = @("DisplayName", "Description", "Direction", "Action", "Enabled", "Profile", "Protocol", "LocalPort", "RemotePort", "Program")
    foreach ($name in $leftByName.Keys) {
        if (-not $rightByName.ContainsKey($name)) { continue }
        $leftRule = $leftByName[$name]
        $rightRule = $rightByName[$name]
        $changes = New-Object System.Collections.Generic.List[string]
        foreach ($property in $properties) {
            $oldValue = [string]$leftRule.$property
            $newValue = [string]$rightRule.$property
            if ($oldValue -ne $newValue) {
                $changes.Add("  $property : $oldValue -> $newValue")
            }
        }
        if ($changes.Count -gt 0) {
            $modified.Add([PSCustomObject]@{ Rule = $rightRule; Changes = @($changes) })
        }
    }

    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("FIREWALL BACKUP COMPARISON")
    [void]$report.AppendLine("=" * 72)
    [void]$report.AppendLine("Left (baseline):  $LeftPath")
    [void]$report.AppendLine("Right (compared): $RightPath")
    [void]$report.AppendLine("")

    [void]$report.AppendLine("ADDED IN RIGHT ($($added.Count))")
    [void]$report.AppendLine("-" * 40)
    if ($added.Count -eq 0) { [void]$report.AppendLine("  (none)") }
    foreach ($rule in $added) {
        [void]$report.AppendLine("  + $($rule.DisplayName) [$($rule.Direction) $($rule.Action)]")
    }
    [void]$report.AppendLine("")

    [void]$report.AppendLine("REMOVED FROM RIGHT ($($removed.Count))")
    [void]$report.AppendLine("-" * 40)
    if ($removed.Count -eq 0) { [void]$report.AppendLine("  (none)") }
    foreach ($rule in $removed) {
        [void]$report.AppendLine("  - $($rule.DisplayName) [$($rule.Direction) $($rule.Action)]")
    }
    [void]$report.AppendLine("")

    [void]$report.AppendLine("MODIFIED IN RIGHT ($($modified.Count))")
    [void]$report.AppendLine("-" * 40)
    if ($modified.Count -eq 0) { [void]$report.AppendLine("  (none)") }
    foreach ($item in $modified) {
        [void]$report.AppendLine("  * $($item.Rule.DisplayName)")
        foreach ($change in $item.Changes) { [void]$report.AppendLine("    $change") }
    }
    [void]$report.AppendLine("")
    [void]$report.AppendLine("Summary: +$($added.Count) added, -$($removed.Count) removed, ~$($modified.Count) modified")

    Show-EditorReportWindow -Title "Compare Firewall Backups" -Text $report.ToString()
}

function Get-MultiMachineProperty {
    param(
        [object]$Object,
        [string]$Name,
        [object]$DefaultValue = $null
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }
    return $property.Value
}

function ConvertTo-MultiMachineRule {
    param([object]$Rule)

    [PSCustomObject][ordered]@{
        Name = [string](Get-MultiMachineProperty -Object $Rule -Name "Name" -DefaultValue "")
        DisplayName = [string](Get-MultiMachineProperty -Object $Rule -Name "DisplayName" -DefaultValue (Get-MultiMachineProperty -Object $Rule -Name "Name" -DefaultValue ""))
        Description = [string](Get-MultiMachineProperty -Object $Rule -Name "Description" -DefaultValue "")
        Direction = [string](Get-MultiMachineProperty -Object $Rule -Name "Direction" -DefaultValue "Unknown")
        Action = [string](Get-MultiMachineProperty -Object $Rule -Name "Action" -DefaultValue "Unknown")
        Enabled = [string](Get-MultiMachineProperty -Object $Rule -Name "Enabled" -DefaultValue "Unknown")
        Profile = [string](Get-MultiMachineProperty -Object $Rule -Name "Profile" -DefaultValue "Any")
        Protocol = [string](Get-MultiMachineProperty -Object $Rule -Name "Protocol" -DefaultValue "Any")
        LocalPort = [string](Get-MultiMachineProperty -Object $Rule -Name "LocalPort" -DefaultValue "Any")
        RemotePort = [string](Get-MultiMachineProperty -Object $Rule -Name "RemotePort" -DefaultValue "Any")
        LocalAddress = [string](Get-MultiMachineProperty -Object $Rule -Name "LocalAddress" -DefaultValue "Any")
        RemoteAddress = [string](Get-MultiMachineProperty -Object $Rule -Name "RemoteAddress" -DefaultValue "Any")
        Program = [string](Get-MultiMachineProperty -Object $Rule -Name "Program" -DefaultValue "Any")
        Service = [string](Get-MultiMachineProperty -Object $Rule -Name "Service" -DefaultValue "Any")
        Group = [string](Get-MultiMachineProperty -Object $Rule -Name "Group" -DefaultValue "")
    }
}

function Read-MultiMachineBackup {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolvedPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path)
    $backup = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $ruleDetailsProperty = $backup.PSObject.Properties["RuleDetails"]
    $rawRules = if ($null -ne $ruleDetailsProperty) { @($ruleDetailsProperty.Value) } else { @($backup) }
    $rules = New-Object System.Collections.Generic.List[PSObject]
    foreach ($rawRule in $rawRules) {
        $normalizedRule = ConvertTo-MultiMachineRule -Rule $rawRule
        if ([string]::IsNullOrWhiteSpace($normalizedRule.Name)) {
            throw "Backup '$resolvedPath' contains a rule without a Name."
        }
        $rules.Add($normalizedRule)
    }

    $computerName = [string](Get-MultiMachineProperty -Object $backup -Name "ComputerName" -DefaultValue ([System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)))
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        $computerName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)
    }

    [PSCustomObject]@{
        Path = $resolvedPath
        MachineName = $computerName
        BackupDate = [string](Get-MultiMachineProperty -Object $backup -Name "BackupDate" -DefaultValue "")
        Rules = $rules.ToArray()
    }
}

function Get-MultiMachineRuleSignature {
    param([Parameter(Mandatory = $true)][object]$Rule)

    $properties = @(
        "DisplayName", "Description", "Direction", "Action", "Enabled", "Profile",
        "Protocol", "LocalPort", "RemotePort", "LocalAddress", "RemoteAddress", "Program",
        "Service", "Group"
    )
    return (($properties | ForEach-Object { [string]$Rule.$_ }) -join [char]31)
}

function Compare-MultiMachineRuleSets {
    param([Parameter(Mandatory = $true)][object[]]$Backups)

    if ($Backups.Count -lt 2) {
        throw "Select at least two firewall backups for a multi-machine comparison."
    }

    $ruleNames = @{}
    foreach ($backup in $Backups) {
        foreach ($rule in @($backup.Rules)) {
            if (-not $ruleNames.ContainsKey($rule.Name)) {
                $ruleNames[$rule.Name] = $true
            }
        }
    }

    $records = New-Object System.Collections.Generic.List[PSObject]
    foreach ($ruleName in @($ruleNames.Keys | Sort-Object)) {
        $machineEntries = New-Object System.Collections.Generic.List[PSObject]
        $signatures = New-Object System.Collections.Generic.List[string]
        foreach ($backup in $Backups) {
            $matchingRule = @($backup.Rules | Where-Object { $_.Name -eq $ruleName } | Select-Object -First 1)
            if ($matchingRule.Count -eq 0) {
                $machineEntries.Add([PSCustomObject]@{
                        Machine = $backup.MachineName
                        Path = $backup.Path
                        Present = $false
                        Rule = $null
                        Signature = ""
                    })
            }
            else {
                $signature = Get-MultiMachineRuleSignature -Rule $matchingRule[0]
                $signatures.Add($signature)
                $machineEntries.Add([PSCustomObject]@{
                        Machine = $backup.MachineName
                        Path = $backup.Path
                        Present = $true
                        Rule = $matchingRule[0]
                        Signature = $signature
                    })
            }
        }

        $distinctSignatures = @($signatures | Select-Object -Unique)
        $missingCount = @($machineEntries | Where-Object { -not $_.Present }).Count
        $state = if ($missingCount -gt 0) { "Missing" } elseif ($distinctSignatures.Count -gt 1) { "Drift" } else { "Consistent" }
        $displayRule = @($machineEntries | Where-Object { $_.Present } | Select-Object -First 1).Rule
        $details = foreach ($entry in $machineEntries) {
            if ($entry.Present) {
                $rule = $entry.Rule
                "[$($entry.Machine)] $($rule.Direction) $($rule.Action), $($rule.Protocol), local $($rule.LocalPort), remote $($rule.RemotePort), program $($rule.Program)"
            }
            else {
                "[$($entry.Machine)] MISSING"
            }
        }

        $records.Add([PSCustomObject]@{
                Name = $ruleName
                DisplayName = if ($displayRule) { $displayRule.DisplayName } else { $ruleName }
                State = $state
                PresentCount = $Backups.Count - $missingCount
                MachineCount = $Backups.Count
                VariantCount = $distinctSignatures.Count
                Details = @($details)
            })
    }
    return $records.ToArray()
}

function Show-MultiMachineComparisonReport {
    param(
        [object[]]$Backups,
        [object[]]$Records
    )

    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("MULTI-MACHINE FIREWALL BACKUP COMPARISON")
    [void]$report.AppendLine("=" * 100)
    [void]$report.AppendLine("Endpoints: $($Backups.Count)")
    foreach ($backup in $Backups) {
        $backupDateText = if ($backup.BackupDate) { " - $($backup.BackupDate)" } else { "" }
        [void]$report.AppendLine("  $($backup.MachineName) - $([System.IO.Path]::GetFileName($backup.Path)) - $($backup.Rules.Count) rules$backupDateText")
    }
    [void]$report.AppendLine("")

    $consistent = @($Records | Where-Object { $_.State -eq "Consistent" })
    $drifted = @($Records | Where-Object { $_.State -eq "Drift" })
    $missing = @($Records | Where-Object { $_.State -eq "Missing" })
    [void]$report.AppendLine("Rule names: $($Records.Count) | Consistent: $($consistent.Count) | Drifted: $($drifted.Count) | Missing: $($missing.Count)")
    [void]$report.AppendLine("")

    foreach ($record in @($Records | Sort-Object @{ Expression = { switch ($_.State) { "Drift" { 0; break } "Missing" { 1; break } default { 2 } } } }, DisplayName)) {
        [void]$report.AppendLine("[$($record.State)] $($record.DisplayName)  <$($record.Name)>")
        foreach ($detail in $record.Details) {
            [void]$report.AppendLine("  $detail")
        }
    }

    if ($Records.Count -eq 0) {
        [void]$report.AppendLine("No rules were present in the selected backups.")
    }
    Show-EditorReportWindow -Title "Multi-Machine Firewall Comparison" -Text $report.ToString() -Width 1200 -Height 760
}

function Compare-FWBackupFleet {
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $openDialog.Title = "Select two or more endpoint backups"
    $openDialog.Multiselect = $true
    if (-not $openDialog.ShowDialog()) { return }
    if ($openDialog.FileNames.Count -lt 2) {
        [System.Windows.MessageBox]::Show(
            "Select at least two backup files for a multi-machine comparison.",
            "Not Enough Backups",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    try {
        Update-Status "Reading $($openDialog.FileNames.Count) endpoint backups..."
        $backups = foreach ($path in $openDialog.FileNames) {
            Read-MultiMachineBackup -Path $path
        }
        $records = @(Compare-MultiMachineRuleSets -Backups @($backups))
        Show-MultiMachineComparisonReport -Backups @($backups) -Records $records
        Update-Status "Multi-machine comparison complete: $($records.Count) rule names"
    }
    catch {
        Update-Status "Multi-machine comparison failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to compare endpoint backups.`n`nError: $($_.Exception.Message)",
            "Multi-Machine Compare Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Compare-FWBackups {
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $openDialog.Title = "Select baseline firewall backup"
    if (-not $openDialog.ShowDialog()) { return }
    $leftPath = $openDialog.FileName

    $openDialog.Title = "Select backup to compare"
    if (-not $openDialog.ShowDialog()) { return }
    $rightPath = $openDialog.FileName

    try {
        Update-Status "Comparing backup files..."
        Compare-BackupRuleSets -LeftPath $leftPath -RightPath $rightPath
        Update-Status "Backup comparison complete"
    }
    catch {
        Update-Status "Backup comparison failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to compare backups.`n`nError: $($_.Exception.Message)",
            "Compare Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Get-TemplateFolder {
    $folder = Join-Path $env:USERPROFILE "FirewallForge_Templates"
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    return $folder
}

function Get-RuleTemplateRecords {
    $folder = Get-TemplateFolder
    $records = New-Object System.Collections.Generic.List[PSObject]
    foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter "*.json" -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try {
            $data = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $templateName = if ($data.TemplateName) { [string]$data.TemplateName } else { $file.BaseName }
            $ruleData = if ($data.Rule) { $data.Rule } else { $data }
            $records.Add([PSCustomObject]@{
                Name = $templateName
                Path = $file.FullName
                Rule = $ruleData
                CreatedDate = $data.CreatedDate
            })
        }
        catch {
            # Ignore malformed templates; the library remains usable and the file can be removed.
        }
    }
    return $records
}

function Show-TemplateNameInput {
    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Save Rule Template"
    $dialog.Width = 420
    $dialog.Height = 190
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.ResizeMode = "NoResize"
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $dialog.Tag = $null

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = New-Object System.Windows.Thickness(20)
    $label = New-Object System.Windows.Controls.Label
    $label.Content = "Template name"
    $panel.Children.Add($label) | Out-Null
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($nameBox) | Out-Null
    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Width = 90
    $cancel.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($cancel) | Out-Null
    $save = New-Object System.Windows.Controls.Button
    $save.Content = "Save"
    $save.Width = 90
    $save.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x38, 0x8E, 0x3C))
    $save.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $save.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($nameBox.Text)) {
            $dialog.Tag = $nameBox.Text.Trim()
            $dialog.Close()
        }
    }.GetNewClosure())
    $buttons.Children.Add($save) | Out-Null
    $panel.Children.Add($buttons) | Out-Null
    $dialog.Content = $panel
    $dialog.ShowDialog() | Out-Null
    return [string]$dialog.Tag
}

function Save-RuleTemplate {
    param(
        [object]$Rule,
        [string]$TemplateName
    )

    $safeName = ($TemplateName -replace '[\\/:*?"<>|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw "Template name cannot be empty."
    }
    $folder = Get-TemplateFolder
    $path = Join-Path $folder "$safeName.json"
    if (Test-Path -LiteralPath $path) {
        $confirm = [System.Windows.MessageBox]::Show(
            "A template named '$safeName' already exists. Replace it?",
            "Replace Template",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            return $false
        }
    }

    $template = [ordered]@{
        TemplateName = $safeName
        CreatedDate = (Get-Date).ToString("o")
        Rule = [ordered]@{
            DisplayName = $Rule.DisplayName
            Description = $Rule.Description
            Direction = $Rule.Direction
            Action = $Rule.Action
            Enabled = $Rule.Enabled
            Profile = $Rule.Profile
            Protocol = $Rule.Protocol
            LocalPort = $Rule.LocalPort
            RemotePort = $Rule.RemotePort
            Program = $Rule.Program
        }
    }
    $template | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding UTF8
    return $true
}

function Insert-RuleTemplate {
    param([object]$TemplateRecord)

    $data = $TemplateRecord.Rule
    $rule = New-RuleObject `
        -Name ("Template_" + [guid]::NewGuid().ToString().Substring(0, 8)) `
        -DisplayName $data.DisplayName `
        -Description $data.Description `
        -Direction $(if ($data.Direction) { $data.Direction } else { "Inbound" }) `
        -Action $(if ($data.Action) { $data.Action } else { "Allow" }) `
        -Enabled $(if ($data.Enabled) { $data.Enabled } else { "True" }) `
        -Profile $(if ($data.Profile) { $data.Profile } else { "Any" }) `
        -Protocol $(if ($data.Protocol) { $data.Protocol } else { "Any" }) `
        -LocalPort $(if ($data.LocalPort) { $data.LocalPort } else { "Any" }) `
        -RemotePort $(if ($data.RemotePort) { $data.RemotePort } else { "Any" }) `
        -Program $(if ($data.Program) { $data.Program } else { "Any" }) `
        -Selected $true
    $Script:AllRules.Add($rule)
    $Script:FilteredView = $null
    $txtSearch.Text = ""
    $dgRules.ItemsSource = $null
    $dgRules.ItemsSource = $Script:AllRules
    Update-Counts
    Update-Status "Inserted template '$($TemplateRecord.Name)'"
}

function Show-TemplateLibrary {
    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Firewall Rule Templates"
    $dialog.Width = 720
    $dialog.Height = 520
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    $contentRow = New-Object System.Windows.Controls.RowDefinition
    $contentRow.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.RowDefinitions.Add($contentRow)
    $buttonRow = New-Object System.Windows.Controls.RowDefinition
    $buttonRow.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($buttonRow)
    $listColumn = New-Object System.Windows.Controls.ColumnDefinition
    $listColumn.Width = New-Object System.Windows.GridLength(220)
    $grid.ColumnDefinitions.Add($listColumn)
    $previewColumn = New-Object System.Windows.Controls.ColumnDefinition
    $previewColumn.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.ColumnDefinitions.Add($previewColumn)

    $list = New-Object System.Windows.Controls.ListBox
    [System.Windows.Controls.Grid]::SetRow($list, 0)
    [System.Windows.Controls.Grid]::SetColumn($list, 0)
    $grid.Children.Add($list) | Out-Null
    $preview = New-Object System.Windows.Controls.TextBox
    $preview.IsReadOnly = $true
    $preview.AcceptsReturn = $true
    $preview.VerticalScrollBarVisibility = "Auto"
    $preview.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26))
    $preview.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $preview.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $preview.Margin = New-Object System.Windows.Thickness(12, 0, 0, 0)
    [System.Windows.Controls.Grid]::SetRow($preview, 0)
    [System.Windows.Controls.Grid]::SetColumn($preview, 1)
    $grid.Children.Add($preview) | Out-Null

    $templateState = @{ Records = @() }
    $refresh = {
        $templateState.Records = @(Get-RuleTemplateRecords)
        $list.Items.Clear()
        foreach ($record in $templateState.Records) { $list.Items.Add($record.Name) | Out-Null }
        if ($templateState.Records.Count -gt 0) { $list.SelectedIndex = 0 }
        else { $preview.Text = "No templates saved yet." }
    }.GetNewClosure()
    $getSelected = {
        if ($list.SelectedIndex -lt 0 -or $list.SelectedIndex -ge $templateState.Records.Count) { return $null }
        return $templateState.Records[$list.SelectedIndex]
    }.GetNewClosure()
    $list.Add_SelectionChanged({
        $record = & $getSelected
        if ($record) {
            $preview.Text = @(
                "Template: $($record.Name)"
                "Created: $($record.CreatedDate)"
                ""
                "DisplayName: $($record.Rule.DisplayName)"
                "Direction: $($record.Rule.Direction)"
                "Action: $($record.Rule.Action)"
                "Enabled: $($record.Rule.Enabled)"
                "Profile: $($record.Rule.Profile)"
                "Protocol: $($record.Rule.Protocol)"
                "LocalPort: $($record.Rule.LocalPort)"
                "RemotePort: $($record.Rule.RemotePort)"
                "Program: $($record.Rule.Program)"
            ) -join [Environment]::NewLine
        }
    }.GetNewClosure())

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    [System.Windows.Controls.Grid]::SetRow($buttons, 1)
    [System.Windows.Controls.Grid]::SetColumnSpan($buttons, 2)
    $saveButton = New-Object System.Windows.Controls.Button
    $saveButton.Content = "Save Selected"
    $saveButton.Width = 125
    $saveButton.Add_Click({
        $rule = $dgRules.SelectedItem
        if ($null -eq $rule) {
            [System.Windows.MessageBox]::Show("Select a rule in the editor first.", "No Rule Selected") | Out-Null
            return
        }
        $name = Show-TemplateNameInput
        if ($name) {
            try {
                if (Save-RuleTemplate -Rule $rule -TemplateName $name) {
                    & $refresh
                    Update-Status "Saved template '$name'"
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Could not save template.`n`n$($_.Exception.Message)", "Template Error") | Out-Null
            }
        }
    }.GetNewClosure())
    $buttons.Children.Add($saveButton) | Out-Null

    $insertButton = New-Object System.Windows.Controls.Button
    $insertButton.Content = "Insert Selected"
    $insertButton.Width = 125
    $insertButton.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x38, 0x8E, 0x3C))
    $insertButton.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $insertButton.Add_Click({
        $record = & $getSelected
        if ($record) {
            Insert-RuleTemplate -TemplateRecord $record
            $dialog.Close()
        }
    }.GetNewClosure())
    $buttons.Children.Add($insertButton) | Out-Null

    $deleteButton = New-Object System.Windows.Controls.Button
    $deleteButton.Content = "Delete Template"
    $deleteButton.Width = 125
    $deleteButton.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xD3, 0x2F, 0x2F))
    $deleteButton.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $deleteButton.Add_Click({
        $record = & $getSelected
        if ($record) {
            $confirm = [System.Windows.MessageBox]::Show("Delete template '$($record.Name)'?", "Delete Template", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
                Remove-Item -LiteralPath $record.Path -Force
                & $refresh
            }
        }
    }.GetNewClosure())
    $buttons.Children.Add($deleteButton) | Out-Null
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Width = 90
    $closeButton.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($closeButton) | Out-Null
    $grid.Children.Add($buttons) | Out-Null

    $dialog.Content = $grid
    & $refresh
    $dialog.ShowDialog() | Out-Null
}

function ConvertTo-NetshFirewallScript {
    param([object[]]$Rules)

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("@echo off")
    $lines.Add("rem FirewallForge generated netsh policy script")
    $lines.Add("rem Run from an elevated Command Prompt.")
    $lines.Add("")
    foreach ($rule in $Rules) {
        $name = ([string]$rule.DisplayName) -replace '"', '\"'
        $direction = if ($rule.Direction -eq "Inbound") { "in" } else { "out" }
        $action = if ($rule.Action -eq "Block") { "block" } else { "allow" }
        $enabled = if ($rule.Enabled -eq "False") { "no" } else { "yes" }
        $line = 'netsh advfirewall firewall add rule name="{0}" dir={1} action={2} enable={3}' -f $name, $direction, $action, $enabled
        if ($rule.Profile -and $rule.Profile -ne "Any") { $line += " profile=$(([string]$rule.Profile).Replace(', ', ','))" }
        if ($rule.Protocol -and $rule.Protocol -ne "Any") { $line += " protocol=$($rule.Protocol.ToLowerInvariant())" }
        if ($rule.LocalPort -and $rule.LocalPort -ne "Any") { $line += (' localport="{0}"' -f $rule.LocalPort) }
        if ($rule.RemotePort -and $rule.RemotePort -ne "Any") { $line += (' remoteport="{0}"' -f $rule.RemotePort) }
        if ($rule.Program -and $rule.Program -ne "Any") { $line += (' program="{0}"' -f $rule.Program) }
        if ($rule.PSObject.Properties.Name -contains "RemoteAddress" -and $rule.RemoteAddress -and $rule.RemoteAddress -ne "Any") { $line += (' remoteip="{0}"' -f $rule.RemoteAddress) }
        if ($rule.PSObject.Properties.Name -contains "Service" -and $rule.Service -and $rule.Service -ne "Any") { $line += (' service="{0}"' -f $rule.Service) }
        $lines.Add($line)
    }
    return ($lines -join [Environment]::NewLine) + [Environment]::NewLine
}

function ConvertTo-PowerShellFirewallScript {
    param([object[]]$Rules)

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# FirewallForge generated New-NetFirewallRule script")
    $lines.Add("# Run from an elevated PowerShell session.")
    $lines.Add("")
    foreach ($rule in $Rules) {
        $parts = New-Object System.Collections.Generic.List[string]
        $escape = { param([object]$Value) "'$(if ($null -eq $Value) { '' } else { ([string]$Value -replace "'", "''") })'" }
        $parts.Add("-DisplayName $(& $escape $rule.DisplayName)")
        $parts.Add("-Direction $(& $escape $rule.Direction)")
        $parts.Add("-Action $(& $escape $rule.Action)")
        $enabledToken = if ($rule.Enabled -eq "False") { '$false' } else { '$true' }
        $parts.Add("-Enabled $enabledToken")
        if ($rule.Profile -and $rule.Profile -ne "Any") { $parts.Add("-Profile $(& $escape $rule.Profile)") } else { $parts.Add("-Profile 'Any'") }
        if ($rule.Protocol -and $rule.Protocol -ne "Any") { $parts.Add("-Protocol $(& $escape $rule.Protocol)") }
        if ($rule.LocalPort -and $rule.LocalPort -ne "Any") { $parts.Add("-LocalPort $(& $escape $rule.LocalPort)") }
        if ($rule.RemotePort -and $rule.RemotePort -ne "Any") { $parts.Add("-RemotePort $(& $escape $rule.RemotePort)") }
        if ($rule.Program -and $rule.Program -ne "Any") { $parts.Add("-Program $(& $escape $rule.Program)") }
        if ($rule.PSObject.Properties.Name -contains "RemoteAddress" -and $rule.RemoteAddress -and $rule.RemoteAddress -ne "Any") { $parts.Add("-RemoteAddress $(& $escape $rule.RemoteAddress)") }
        if ($rule.PSObject.Properties.Name -contains "Service" -and $rule.Service -and $rule.Service -ne "Any") { $parts.Add("-Service $(& $escape $rule.Service)") }
        $lines.Add("New-NetFirewallRule $($parts -join ' ') -ErrorAction Stop")
    }
    return ($lines -join [Environment]::NewLine) + [Environment]::NewLine
}

function New-GpoFirewallRuleData {
    param([object]$Rule)

    $tokens = New-Object System.Collections.Generic.List[string]
    $tokens.Add("v2.10")
    $tokens.Add("Action=$(if ($Rule.Action -eq 'Block') { 'Block' } else { 'Allow' })")
    $tokens.Add("Active=$(if ($Rule.Enabled -eq 'False') { 'FALSE' } else { 'TRUE' })")
    $tokens.Add("Dir=$(if ($Rule.Direction -eq 'Inbound') { 'In' } else { 'Out' })")
    if ($Rule.Profile -and $Rule.Profile -ne "Any") {
        foreach ($profile in ([string]$Rule.Profile -split ',')) {
            $tokens.Add("Profile=$($profile.Trim())")
        }
    }
    if ($Rule.Protocol -and $Rule.Protocol -ne "Any") {
        $protocol = switch ($Rule.Protocol.ToUpperInvariant()) {
            "TCP" { "6"; break }
            "UDP" { "17"; break }
            "ICMPV4" { "1"; break }
            "ICMPV6" { "58"; break }
            default { [string]$Rule.Protocol }
        }
        $tokens.Add("Protocol=$protocol")
    }
    if ($Rule.LocalPort -and $Rule.LocalPort -ne "Any") { $tokens.Add("LPort=$($Rule.LocalPort)") }
    if ($Rule.RemotePort -and $Rule.RemotePort -ne "Any") { $tokens.Add("RPort=$($Rule.RemotePort)") }
    if ($Rule.Program -and $Rule.Program -ne "Any") { $tokens.Add("App=$($Rule.Program)") }
    if ($Rule.PSObject.Properties.Name -contains "RemoteAddress" -and $Rule.RemoteAddress -and $Rule.RemoteAddress -ne "Any") {
        foreach ($address in ([string]$Rule.RemoteAddress -split ',')) {
            $trimmed = $address.Trim()
            if ($trimmed -match ':') { $tokens.Add("RA6=$trimmed") } else { $tokens.Add("RA4=$trimmed") }
        }
    }
    if ($Rule.PSObject.Properties.Name -contains "Service" -and $Rule.Service -and $Rule.Service -ne "Any") { $tokens.Add("Svc=$($Rule.Service)") }
    $tokens.Add("Name=$($Rule.DisplayName)")
    $tokens.Add("Desc=$([string]$Rule.Description)")
    return ($tokens -join '|') + '|'
}

function Write-RegistryPolString {
    param(
        [System.IO.BinaryWriter]$Writer,
        [string]$Value
    )

    if ($null -eq $Value) { $Value = "" }
    $Writer.Write([System.Text.Encoding]::Unicode.GetBytes($Value))
    $Writer.Write([byte]0)
    $Writer.Write([byte]0)
}

function Write-RegistryPolRecord {
    param(
        [System.IO.BinaryWriter]$Writer,
        [string]$Key,
        [string]$ValueName,
        [int]$Type,
        [byte[]]$Data
    )

    $Writer.Write([byte]0x5B); $Writer.Write([byte]0)
    Write-RegistryPolString -Writer $Writer -Value $Key
    $Writer.Write([byte]0x3B); $Writer.Write([byte]0)
    Write-RegistryPolString -Writer $Writer -Value $ValueName
    $Writer.Write([byte]0x3B); $Writer.Write([byte]0)
    $Writer.Write([int32]$Type)
    $Writer.Write([byte]0x3B); $Writer.Write([byte]0)
    $Writer.Write([int32]$Data.Length)
    $Writer.Write([byte]0x3B); $Writer.Write([byte]0)
    $Writer.Write($Data)
    $Writer.Write([byte]0x5D); $Writer.Write([byte]0)
}

function Write-GpoRegistryPol {
    param(
        [object[]]$Rules,
        [string]$Path
    )

    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $writer = New-Object System.IO.BinaryWriter($stream)
    try {
        # REGFILE_SIGNATURE (PReg) and registry policy file version 1.
        $writer.Write([int32]0x67655250)
        $writer.Write([int32]1)
        $key = "Software\Policies\Microsoft\WindowsFirewall\FirewallRules"
        foreach ($rule in $Rules) {
            $valueName = if ($rule.Name -match '^\{[0-9A-Fa-f-]+\}$') { $rule.Name } else { "{$([guid]::NewGuid().ToString().ToUpperInvariant())}" }
            $data = [System.Text.Encoding]::Unicode.GetBytes((New-GpoFirewallRuleData -Rule $rule) + [char]0)
            Write-RegistryPolRecord -Writer $writer -Key $key -ValueName $valueName -Type 1 -Data $data
        }
    }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Show-PolicyExportPicker {
    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Export Firewall Policy"
    $dialog.Width = 560
    $dialog.Height = 260
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $dialog.Tag = $null
    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = New-Object System.Windows.Thickness(20)
    $label = New-Object System.Windows.Controls.Label
    $label.Content = "Output format"
    $panel.Children.Add($label) | Out-Null
    $formatCombo = New-Object System.Windows.Controls.ComboBox
    foreach ($format in @("netsh firewall script", "PowerShell New-NetFirewallRule script", "GPO Registry.pol fragment")) { $formatCombo.Items.Add($format) | Out-Null }
    $formatCombo.SelectedIndex = 0
    $formatCombo.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $panel.Children.Add($formatCombo) | Out-Null
    $details = New-Object System.Windows.Controls.TextBlock
    $details.Text = "The GPO option writes a computer-scoped Registry.pol fragment using the Windows registry policy file format."
    $details.TextWrapping = "Wrap"
    $details.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xB0, 0xB0, 0xB0))
    $details.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($details) | Out-Null
    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Width = 90
    $cancel.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($cancel) | Out-Null
    $export = New-Object System.Windows.Controls.Button
    $export.Content = "Continue"
    $export.Width = 100
    $export.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $export.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $export.Add_Click({ $dialog.Tag = [string]$formatCombo.SelectedItem; $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($export) | Out-Null
    $panel.Children.Add($buttons) | Out-Null
    $dialog.Content = $panel
    $dialog.ShowDialog() | Out-Null
    return [string]$dialog.Tag
}

function Export-SelectedPolicy {
    $selectedRules = @($Script:AllRules | Where-Object { $_.Selected })
    if ($selectedRules.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No rules selected for policy export.`nUse the checkboxes to select rules.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $format = Show-PolicyExportPicker
    if ([string]::IsNullOrWhiteSpace($format)) { return }
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    if ($format -eq "netsh firewall script") {
        $saveDialog.Filter = "Command script (*.cmd)|*.cmd|All files (*.*)|*.*"
        $saveDialog.DefaultExt = ".cmd"
    }
    elseif ($format -eq "PowerShell New-NetFirewallRule script") {
        $saveDialog.Filter = "PowerShell script (*.ps1)|*.ps1|All files (*.*)|*.*"
        $saveDialog.DefaultExt = ".ps1"
    }
    else {
        $saveDialog.Filter = "Registry policy file (*.pol)|*.pol|All files (*.*)|*.*"
        $saveDialog.DefaultExt = ".pol"
    }
    $saveDialog.FileName = "FirewallPolicy_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    if (-not $saveDialog.ShowDialog()) { return }

    try {
        if ($format -eq "netsh firewall script") {
            ConvertTo-NetshFirewallScript -Rules $selectedRules | Set-Content -LiteralPath $saveDialog.FileName -Encoding UTF8
        }
        elseif ($format -eq "PowerShell New-NetFirewallRule script") {
            ConvertTo-PowerShellFirewallScript -Rules $selectedRules | Set-Content -LiteralPath $saveDialog.FileName -Encoding UTF8
        }
        else {
            Write-GpoRegistryPol -Rules $selectedRules -Path $saveDialog.FileName
        }
        Update-Status "Exported $($selectedRules.Count) rules as $format"
        [System.Windows.MessageBox]::Show(
            "Policy export complete.`n`nRules: $($selectedRules.Count)`nFile: $($saveDialog.FileName)",
            "Policy Export Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
    catch {
        Update-Status "Policy export failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to export policy.`n`nError: $($_.Exception.Message)",
            "Policy Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Export-SelectedToBackup {
    $selectedRules = @($Script:AllRules | Where-Object { $_.Selected })

    if ($selectedRules.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No rules selected for export.`nUse the checkboxes to select rules.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "Firewall Backup (*.fwbackup)|*.fwbackup|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = ".fwbackup"
    $saveDialog.FileName = "FirewallRules_Custom_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    if ($saveDialog.ShowDialog()) {
        try {
            $ruleDetails = $selectedRules | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    Direction = $_.Direction
                    Action = $_.Action
                    Enabled = $_.Enabled
                    Profile = $_.Profile
                    Protocol = $_.Protocol
                    LocalPort = $_.LocalPort
                    RemotePort = $_.RemotePort
                    Program = $_.Program
                }
            }

            $backup = @{
                BackupDate = (Get-Date).ToString("o")
                ComputerName = "Custom Export"
                RuleCount = $selectedRules.Count
                RuleDetails = $ruleDetails
            }

            $backup | ConvertTo-Json -Depth 10 | Out-File $saveDialog.FileName -Encoding UTF8

            Update-Status "Exported $($selectedRules.Count) rules to backup"
            [System.Windows.MessageBox]::Show(
                "Export successful!`n`nRules exported: $($selectedRules.Count)`nFile: $($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        catch {
            Update-Status "Export failed: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Failed to export.`n`nError: $($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    }
}

function Export-SelectedToCSV {
    $selectedRules = @($Script:AllRules | Where-Object { $_.Selected })

    if ($selectedRules.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No rules selected for export.`nUse the checkboxes to select rules.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = ".csv"
    $saveDialog.FileName = "FirewallRules_Custom_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    if ($saveDialog.ShowDialog()) {
        try {
            $selectedRules | Select-Object Name, DisplayName, Description, Direction, Action, Enabled, Profile, Protocol, LocalPort, RemotePort, Program |
                Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8

            Update-Status "Exported $($selectedRules.Count) rules to CSV"
            [System.Windows.MessageBox]::Show(
                "Export successful!`n`nRules exported: $($selectedRules.Count)`nFile: $($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        catch {
            Update-Status "Export failed: $($_.Exception.Message)"
        }
    }
}

function Test-EditorRuleSearchMatch {
    param(
        [object]$Rule,
        [string]$SearchText,
        [bool]$UseRegex
    )

    $fields = @(
        $Rule.Name, $Rule.DisplayName, $Rule.Description, $Rule.Direction, $Rule.Action,
        $Rule.Enabled, $Rule.Profile, $Rule.Protocol, $Rule.LocalPort, $Rule.RemotePort, $Rule.Program
    )
    foreach ($field in $fields) {
        $value = if ($null -eq $field) { "" } else { [string]$field }
        if ($UseRegex) {
            if ([regex]::IsMatch($value, $SearchText, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                return $true
            }
        }
        elseif ($value -like "*$SearchText*") {
            return $true
        }
    }
    return $false
}

function Filter-Rules {
    $searchText = $txtSearch.Text.Trim()
    $useRegex = [bool]$chkRegexSearch.IsChecked

    if ([string]::IsNullOrEmpty($searchText)) {
        $Script:FilteredView = $null
        $dgRules.ItemsSource = $Script:AllRules
        Update-Status "Showing all $($Script:AllRules.Count) rules"
    }
    else {
        if ($useRegex) {
            try {
                $null = [regex]::new($searchText)
            }
            catch {
                Update-Status "Invalid regex: $($_.Exception.Message)"
                return
            }
        }
        $Script:FilteredView = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($rule in $Script:AllRules) {
            if (Test-EditorRuleSearchMatch -Rule $rule -SearchText $searchText -UseRegex $useRegex) {
                $Script:FilteredView.Add($rule)
            }
        }
        $dgRules.ItemsSource = $Script:FilteredView
        Update-Status "Showing $($Script:FilteredView.Count) of $($Script:AllRules.Count) rules"
    }
    Update-Counts
}

# ============================================================
# Event Handlers
# ============================================================
$btnImportBackup.Add_Click({ Import-FWBackup })
$btnImportCSV.Add_Click({ Import-CSV })
$btnMergeBackup.Add_Click({ Merge-FWBackup })
$btnExportBackup.Add_Click({ Export-SelectedToBackup })
$btnExportCSV.Add_Click({ Export-SelectedToCSV })
$btnShowChanges.Add_Click({ Show-Changes })
$btnCompareBackups.Add_Click({ Compare-FWBackups })
$btnCompareFleet.Add_Click({ Compare-FWBackupFleet })
$btnTemplates.Add_Click({ Show-TemplateLibrary })
$btnExportPolicy.Add_Click({ Export-SelectedPolicy })

$btnClearAll.Add_Click({
    $Script:AllRules = [System.Collections.Generic.List[PSObject]]::new()
    $Script:FilteredView = $null
    $Script:OriginalRules = $null
    $dgRules.ItemsSource = $Script:AllRules
    Update-Counts
    Update-Status "All rules cleared"
})

$btnSelectAll.Add_Click({
    foreach ($rule in $Script:AllRules) {
        $rule.Selected = $true
    }
    # Force UI refresh
    $dgRules.ItemsSource = $null
    $dgRules.ItemsSource = if ($Script:FilteredView) { $Script:FilteredView } else { $Script:AllRules }
    Update-Counts
})

$btnSelectNone.Add_Click({
    foreach ($rule in $Script:AllRules) {
        $rule.Selected = $false
    }
    # Force UI refresh
    $dgRules.ItemsSource = $null
    $dgRules.ItemsSource = if ($Script:FilteredView) { $Script:FilteredView } else { $Script:AllRules }
    Update-Counts
})

$btnInvertSelection.Add_Click({
    foreach ($rule in $Script:AllRules) {
        $rule.Selected = -not $rule.Selected
    }
    # Force UI refresh
    $dgRules.ItemsSource = $null
    $dgRules.ItemsSource = if ($Script:FilteredView) { $Script:FilteredView } else { $Script:AllRules }
    Update-Counts
})

$btnSelectFiltered.Add_Click({
    $currentSource = if ($Script:FilteredView) { $Script:FilteredView } else { $Script:AllRules }
    foreach ($rule in $currentSource) {
        $rule.Selected = $true
    }
    # Force UI refresh
    $dgRules.ItemsSource = $null
    $dgRules.ItemsSource = $currentSource
    Update-Counts
})

$btnDeleteSelected.Add_Click({
    # Get rules that are either checkbox-selected OR row-selected in the DataGrid
    $checkboxSelected = @($Script:AllRules | Where-Object { $_.Selected })
    $rowSelected = @($dgRules.SelectedItems)

    # Combine both selections (unique)
    $toDelete = @{}
    foreach ($rule in $checkboxSelected) {
        $toDelete[$rule.Name] = $rule
    }
    foreach ($rule in $rowSelected) {
        if ($rule -and $rule.Name) {
            $toDelete[$rule.Name] = $rule
        }
    }

    $selectedRules = @($toDelete.Values)

    if ($selectedRules.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No rules selected.`n`nYou can select rules by:`n- Checking the 'Export' checkbox`n- Clicking on rows (Ctrl+Click for multiple)",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Delete $($selectedRules.Count) selected rule(s)?`n`nThis cannot be undone.",
        "Confirm Delete",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )

    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        foreach ($rule in $selectedRules) {
            $Script:AllRules.Remove($rule) | Out-Null
            if ($Script:FilteredView) {
                $Script:FilteredView.Remove($rule) | Out-Null
            }
        }

        # Refresh the view
        $dgRules.ItemsSource = $null
        $dgRules.ItemsSource = if ($Script:FilteredView) { $Script:FilteredView } else { $Script:AllRules }
        Update-Counts
        Update-Status "Deleted $($selectedRules.Count) rule(s)"
    }
})

$btnSearch.Add_Click({ Filter-Rules })
$chkRegexSearch.Add_Click({ if (-not [string]::IsNullOrWhiteSpace($txtSearch.Text)) { Filter-Rules } })
$btnRefreshCounts.Add_Click({ Update-Counts })
$btnClearSearch.Add_Click({
    $txtSearch.Text = ""
    $Script:FilteredView = $null
    $dgRules.ItemsSource = $Script:AllRules
    Update-Status "Showing all $($Script:AllRules.Count) rules"
    Update-Counts
})

$txtSearch.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        Filter-Rules
    }
})

$btnAddRule.Add_Click({
    $displayName = $txtNewName.Text.Trim()

    if ([string]::IsNullOrEmpty($displayName)) {
        [System.Windows.MessageBox]::Show("Please enter a display name for the rule.", "Missing Name",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $rule = New-RuleObject `
        -Name ("Custom_" + [guid]::NewGuid().ToString().Substring(0, 8)) `
        -DisplayName $displayName `
        -Direction $cmbNewDirection.SelectedItem.Content `
        -Action $cmbNewAction.SelectedItem.Content `
        -Enabled "True" `
        -Profile "Any" `
        -Protocol $cmbNewProtocol.SelectedItem.Content `
        -LocalPort $txtNewPort.Text `
        -RemotePort "Any" `
        -Program "Any" `
        -Selected $true

    $Script:AllRules.Add($rule)

    # Clear filter and refresh view
    $Script:FilteredView = $null
    $txtSearch.Text = ""
    $dgRules.ItemsSource = $null
    $dgRules.ItemsSource = $Script:AllRules

    Update-Counts
    Update-Status "Added rule: $displayName"

    # Clear input fields
    $txtNewName.Text = ""
    $txtNewPort.Text = "Any"

    # Scroll to the new rule
    $dgRules.ScrollIntoView($rule)
})

# Initialize
$Script:AllRules = [System.Collections.Generic.List[PSObject]]::new()
$dgRules.ItemsSource = $Script:AllRules

# Update counts when cells are edited (including checkboxes)
$dgRules.Add_CellEditEnding({
    $Window.Dispatcher.BeginInvoke([Action]{ Update-Counts }, [System.Windows.Threading.DispatcherPriority]::Background)
})

# Handle checkbox clicks specifically
$dgRules.Add_PreviewMouseLeftButtonUp({
    $Window.Dispatcher.BeginInvoke([Action]{
        Start-Sleep -Milliseconds 100
        Update-Counts
    }, [System.Windows.Threading.DispatcherPriority]::Background)
})

# Show window
$Window.ShowDialog() | Out-Null
