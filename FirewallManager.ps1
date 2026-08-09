<#
.SYNOPSIS
    FirewallManager - Windows Firewall Rule Backup, Edit, and Restore Tool
.DESCRIPTION
    A professional GUI application for backing up, editing, and restoring Windows Firewall rules.
.NOTES
    Author: Matt
    Version: 1.2.0
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
        Title="FirewallManager v1.2.0"
        Height="860" Width="1360"
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
            <Button x:Name="btnProgramWizard" Content="Program Wizard" Style="{StaticResource SuccessButton}" Width="140"/>
            <Button x:Name="btnConnectionMonitor" Content="Monitor Off" Width="120"/>
            <Button x:Name="btnOutboundLockdown" Content="Lockdown" Style="{StaticResource DangerButton}" Width="110"/>
            <Button x:Name="btnRollbackLockdown" Content="Rollback" Style="{StaticResource WarningButton}" Width="110"/>
            <Button x:Name="btnGroupOps" Content="Group Ops" Width="110"/>
            <Button x:Name="btnRulePriority" Content="Rule Test" Width="110"/>
            <Button x:Name="btnLogViewer" Content="Log Viewer" Width="110"/>
            <Button x:Name="btnAuditRules" Content="Audits" Style="{StaticResource WarningButton}" Width="95"/>
            <Button x:Name="btnSavedViews" Content="Views" Width="85"/>
            <Button x:Name="btnScheduleBackups" Content="Schedule" Width="100"/>
            <Button x:Name="btnBulkTags" Content="Tags" Width="80"/>
        </WrapPanel>

        <!-- Search and Filter Row -->
        <WrapPanel Grid.Row="2" Margin="0,0,0,10" VerticalAlignment="Center">
            <TextBox x:Name="txtSearch" Width="250" Margin="5,5,5,5"
                     ToolTip="Search rules by name, program, or port"/>
            <CheckBox x:Name="chkRegexSearch" Content="Regex" Margin="5,0,0,0" VerticalAlignment="Center"/>
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
                <DataGridTextColumn Header="Remote Address" Binding="{Binding RemoteAddress}" Width="150"/>
                <DataGridTextColumn Header="Program" Binding="{Binding Program}" Width="250"/>
                <DataGridTextColumn Header="Service" Binding="{Binding Service}" Width="100"/>
                <DataGridTextColumn Header="Group" Binding="{Binding Group}" Width="150"/>
                <DataGridTextColumn Header="Policy" Binding="{Binding PolicySource}" Width="150"/>
                <DataGridTextColumn Header="Interface" Binding="{Binding VirtualizationScope}" Width="130"/>
                <DataGridTextColumn Header="Tags" Binding="{Binding Tags}" Width="150"/>
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
                    <Label Content="Profiles"/>
                    <WrapPanel>
                        <CheckBox x:Name="chkProfileDomain" Content="Domain" Margin="0,4,10,0"/>
                        <CheckBox x:Name="chkProfilePrivate" Content="Private" Margin="0,4,10,0"/>
                        <CheckBox x:Name="chkProfilePublic" Content="Public" Margin="0,4,0,0"/>
                    </WrapPanel>
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

                <StackPanel Grid.Column="3" Grid.Row="1" Margin="5">
                    <Label Content="Remote Address"/>
                    <TextBox x:Name="txtRemoteAddress"/>
                </StackPanel>

                <StackPanel Grid.Column="0" Grid.Row="2" Margin="5">
                    <Label Content="Service"/>
                    <TextBox x:Name="txtService"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Grid.Row="2" Grid.ColumnSpan="2" Margin="5">
                    <Label Content="Tags"/>
                    <TextBox x:Name="txtTags"/>
                </StackPanel>

                <StackPanel Grid.Column="3" Grid.Row="2" Grid.ColumnSpan="2" Margin="5" Orientation="Horizontal"
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
try {
    $brandingIconPath = Join-Path $PSScriptRoot 'icon.ico'
    if (Test-Path $brandingIconPath) {
        $Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.Uri($brandingIconPath)))
    }
} catch {
}
Write-Host "  - XAML parsed successfully" -ForegroundColor Gray

Write-Host "Binding controls..." -ForegroundColor Cyan
# Get controls
$btnBackup = $Window.FindName("btnBackup")
$btnRestore = $Window.FindName("btnRestore")
$btnRefresh = $Window.FindName("btnRefresh")
$btnExportCSV = $Window.FindName("btnExportCSV")
$btnFindDuplicates = $Window.FindName("btnFindDuplicates")
$btnQuickBlock = $Window.FindName("btnQuickBlock")
$btnProgramWizard = $Window.FindName("btnProgramWizard")
$btnConnectionMonitor = $Window.FindName("btnConnectionMonitor")
$btnOutboundLockdown = $Window.FindName("btnOutboundLockdown")
$btnRollbackLockdown = $Window.FindName("btnRollbackLockdown")
$btnGroupOps = $Window.FindName("btnGroupOps")
$btnRulePriority = $Window.FindName("btnRulePriority")
$btnLogViewer = $Window.FindName("btnLogViewer")
$btnAuditRules = $Window.FindName("btnAuditRules")
$btnSavedViews = $Window.FindName("btnSavedViews")
$btnScheduleBackups = $Window.FindName("btnScheduleBackups")
$btnBulkTags = $Window.FindName("btnBulkTags")
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
$chkProfileDomain = $Window.FindName("chkProfileDomain")
$chkProfilePrivate = $Window.FindName("chkProfilePrivate")
$chkProfilePublic = $Window.FindName("chkProfilePublic")
$cmbProtocol = $Window.FindName("cmbProtocol")
$cmbGroupFilter = $Window.FindName("cmbGroupFilter")
$chkRegexSearch = $Window.FindName("chkRegexSearch")
$txtLocalPort = $Window.FindName("txtLocalPort")
$txtRemotePort = $Window.FindName("txtRemotePort")
$txtRemoteAddress = $Window.FindName("txtRemoteAddress")
$txtService = $Window.FindName("txtService")
$txtTags = $Window.FindName("txtTags")
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
$Script:ConnectionMonitorTimer = $null
$Script:SeenConnectionEvents = @{}
$Script:ConnectionMonitorLastTime = (Get-Date).AddSeconds(-5)
$Script:ViewsFolder = Join-Path $env:APPDATA "FirewallForge\Views"
$Script:LastLockdownSnapshotPath = $null
$Script:LockdownRulePrefix = "FirewallForge Lockdown - "
$Script:ScheduledBackupTaskName = "FirewallForge Scheduled Backup"
$Script:ScheduledBackupWorker = Join-Path $PSScriptRoot "FirewallForgeScheduledBackup.ps1"
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

function Show-ReportWindow {
    param(
        [string]$Title,
        [string]$Text,
        [int]$Width = 760,
        [int]$Height = 560
    )

    $reportWindow = New-Object System.Windows.Window
    $reportWindow.Title = $Title
    $reportWindow.Width = $Width
    $reportWindow.Height = $Height
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
    $textBox.Text = $Text
    $textBox.IsReadOnly = $true
    $textBox.Background = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26)))
    $textBox.Foreground = (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0)))
    $textBox.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $textBox.FontSize = 12
    $textBox.AcceptsReturn = $true
    $textBox.VerticalScrollBarVisibility = "Auto"
    $textBox.HorizontalScrollBarVisibility = "Auto"
    [System.Windows.Controls.Grid]::SetRow($textBox, 0)
    $grid.Children.Add($textBox) | Out-Null

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
    $grid.Children.Add($closeBtn) | Out-Null

    $reportWindow.Content = $grid
    $reportWindow.ShowDialog() | Out-Null
}

function Get-ConnectionEventValue {
    param(
        [hashtable]$Data,
        [string[]]$Names,
        [string]$Default = ""
    )

    foreach ($name in $Names) {
        if ($Data.ContainsKey($name) -and -not [string]::IsNullOrWhiteSpace([string]$Data[$name])) {
            return [string]$Data[$name]
        }
    }

    return $Default
}

function Convert-ConnectionEvent {
    param([System.Diagnostics.Eventing.Reader.EventRecord]$EventRecord)

    $data = @{}
    try {
        $eventXml = [xml]$EventRecord.ToXml()
        foreach ($item in @($eventXml.Event.EventData.Data)) {
            if ($item.Name) {
                $data[[string]$item.Name] = [string]$item.'#text'
            }
        }
    }
    catch {
        return $null
    }

    $directionValue = Get-ConnectionEventValue -Data $data -Names @("Direction") -Default "Unknown"
    $direction = switch -Regex ($directionValue) {
        "^(%%)?14592$|Inbound" { "Inbound"; break }
        "^(%%)?14593$|Outbound" { "Outbound"; break }
        default { "Unknown" }
    }

    $protocolValue = Get-ConnectionEventValue -Data $data -Names @("Protocol") -Default ""
    $protocol = switch ($protocolValue) {
        "6" { "TCP"; break }
        "17" { "UDP"; break }
        "1" { "ICMPv4"; break }
        "58" { "ICMPv6"; break }
        default { if ($protocolValue) { $protocolValue } else { "Any" } }
    }

    $application = Get-ConnectionEventValue -Data $data -Names @("ApplicationName", "ApplicationInformation") -Default "Unknown application"
    $processId = Get-ConnectionEventValue -Data $data -Names @("ProcessID", "ProcessId") -Default ""
    $sourceAddress = Get-ConnectionEventValue -Data $data -Names @("SourceAddress") -Default "Any"
    $sourcePort = Get-ConnectionEventValue -Data $data -Names @("SourcePort") -Default "Any"
    $destinationAddress = Get-ConnectionEventValue -Data $data -Names @("DestAddress", "DestinationAddress") -Default "Any"
    $destinationPort = Get-ConnectionEventValue -Data $data -Names @("DestPort", "DestinationPort") -Default "Any"

    [PSCustomObject]@{
        RecordId = $EventRecord.RecordId
        Timestamp = $EventRecord.TimeCreated
        EventId = [int]$EventRecord.Id
        Action = if ([int]$EventRecord.Id -eq 5157) { "Blocked" } else { "Allowed" }
        Program = $application
        ProcessId = $processId
        Direction = $direction
        SourceAddress = $sourceAddress
        SourcePort = $sourcePort
        DestinationAddress = $destinationAddress
        DestinationPort = $destinationPort
        Protocol = $protocol
    }
}

function Get-ConnectionEvents {
    param([datetime]$StartTime)

    $events = @(Get-WinEvent -FilterHashtable @{
        LogName = "Security"
        Id = @(5156, 5157)
        StartTime = $StartTime
    } -MaxEvents 50 -ErrorAction Stop | Sort-Object TimeCreated)

    foreach ($eventRecord in $events) {
        $connectionEvent = Convert-ConnectionEvent -EventRecord $eventRecord
        if ($connectionEvent) {
            $connectionEvent
        }
    }
}

function Show-ConnectionDecision {
    param([object]$ConnectionEvent)

    $programName = [System.IO.Path]::GetFileName($ConnectionEvent.Program)
    if ([string]::IsNullOrWhiteSpace($programName)) {
        $programName = $ConnectionEvent.Program
    }

    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Firewall connection detected"
    $dialog.Width = 560
    $dialog.Height = 330
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.ResizeMode = "NoResize"
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $dialog.Tag = "Ignore"

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = New-Object System.Windows.Thickness(20)

    $heading = New-Object System.Windows.Controls.TextBlock
    $heading.Text = "New $($ConnectionEvent.Action.ToLowerInvariant()) connection"
    $heading.FontSize = 20
    $heading.FontWeight = "Bold"
    $heading.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $heading.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($heading) | Out-Null

    $details = New-Object System.Windows.Controls.TextBlock
    $details.Text = @(
        "Program: $programName"
        "Path: $($ConnectionEvent.Program)"
        "Direction: $($ConnectionEvent.Direction)"
        "Remote: $($ConnectionEvent.DestinationAddress):$($ConnectionEvent.DestinationPort)"
        "Protocol: $($ConnectionEvent.Protocol)    Process ID: $($ConnectionEvent.ProcessId)"
        ""
        "Choose an action to create a persistent firewall rule."
    ) -join [Environment]::NewLine
    $details.TextWrapping = "Wrap"
    $details.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $details.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($details) | Out-Null

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    foreach ($choice in @(
        @{ Label = "Allow"; Value = "Allow"; Color = @(0x38, 0x8E, 0x3C) }
        @{ Label = "Block"; Value = "Block"; Color = @(0xD3, 0x2F, 0x2F) }
        @{ Label = "Ignore"; Value = "Ignore"; Color = @(0x55, 0x55, 0x55) }
    )) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = $choice.Label
        $button.Width = 100
        $button.Margin = New-Object System.Windows.Thickness(5, 0, 0, 0)
        $button.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
        $rgb = $choice.Color
        $button.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb($rgb[0], $rgb[1], $rgb[2]))
        $value = $choice.Value
        $button.Add_Click({ $dialog.Tag = $value; $dialog.Close() }.GetNewClosure())
        $buttons.Children.Add($button) | Out-Null
    }
    $panel.Children.Add($buttons) | Out-Null

    $dialog.Content = $panel
    $dialog.ShowDialog() | Out-Null
    return [string]$dialog.Tag
}

function Add-ConnectionDecisionRule {
    param(
        [object]$ConnectionEvent,
        [ValidateSet("Allow", "Block")]
        [string]$Action
    )

    if ([string]::IsNullOrWhiteSpace($ConnectionEvent.Program) -or
        $ConnectionEvent.Program -eq "Unknown application" -or
        $ConnectionEvent.Program -notmatch "^[A-Za-z]:\\|^\\\\") {
        throw "Windows did not provide a usable application path for this event."
    }
    if ($ConnectionEvent.Direction -notin @("Inbound", "Outbound")) {
        throw "Windows did not provide a usable connection direction for this event."
    }

    $programName = [System.IO.Path]::GetFileNameWithoutExtension($ConnectionEvent.Program)
    if ([string]::IsNullOrWhiteSpace($programName)) {
        $programName = "connection"
    }
    $displayName = "FirewallForge Prompt - $Action $programName ($($ConnectionEvent.Direction)) - $([guid]::NewGuid().ToString().Substring(0, 8))"
    $params = @{
        DisplayName = $displayName
        Direction = $ConnectionEvent.Direction
        Action = $Action
        Program = $ConnectionEvent.Program
        Profile = "Any"
        Enabled = "True"
        Description = "Created by FirewallForge connection monitor"
        ErrorAction = "Stop"
    }

    if ($ConnectionEvent.Protocol -in @("TCP", "UDP")) {
        $params.Protocol = $ConnectionEvent.Protocol
        if ($ConnectionEvent.Direction -eq "Outbound") {
            if ($ConnectionEvent.DestinationPort -match "^\d+(?:-\d+)?$") {
                $params.RemotePort = $ConnectionEvent.DestinationPort
            }
            if ($ConnectionEvent.DestinationAddress -and $ConnectionEvent.DestinationAddress -notin @("Any", "-")) {
                $params.RemoteAddress = $ConnectionEvent.DestinationAddress
            }
        }
        else {
            if ($ConnectionEvent.DestinationPort -match "^\d+(?:-\d+)?$") {
                $params.LocalPort = $ConnectionEvent.DestinationPort
            }
            if ($ConnectionEvent.SourceAddress -and $ConnectionEvent.SourceAddress -notin @("Any", "-")) {
                $params.RemoteAddress = $ConnectionEvent.SourceAddress
            }
            if ($ConnectionEvent.SourcePort -match "^\d+(?:-\d+)?$") {
                $params.RemotePort = $ConnectionEvent.SourcePort
            }
        }
    }

    New-NetFirewallRule @params | Out-Null
}

function Process-ConnectionMonitorTick {
    try {
        $events = @(Get-ConnectionEvents -StartTime $Script:ConnectionMonitorLastTime)
        if ($events.Count -eq 0) {
            return
        }

        $Script:ConnectionMonitorLastTime = ($events | Sort-Object Timestamp | Select-Object -Last 1).Timestamp
        foreach ($connectionEvent in $events) {
            $signature = @(
                $connectionEvent.Program, $connectionEvent.Direction, $connectionEvent.Protocol,
                $connectionEvent.DestinationAddress, $connectionEvent.DestinationPort
            ) -join "|"
            if ($Script:SeenConnectionEvents.ContainsKey($signature) -and
                ((Get-Date) - $Script:SeenConnectionEvents[$signature]).TotalMinutes -lt 5) {
                continue
            }
            $Script:SeenConnectionEvents[$signature] = Get-Date

            $decision = Show-ConnectionDecision -ConnectionEvent $connectionEvent
            if ($decision -in @("Allow", "Block")) {
                try {
                    Add-ConnectionDecisionRule -ConnectionEvent $connectionEvent -Action $decision
                    Update-Status "Created $decision rule from connection monitor" "#00FF00"
                    Get-FirewallRules
                }
                catch {
                    Update-Status "Connection rule failed: $($_.Exception.Message)" "#FF0000"
                    [System.Windows.MessageBox]::Show(
                        "Could not create the connection rule.`n`nError: $($_.Exception.Message)",
                        "Connection Monitor Error",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error) | Out-Null
                }
            }
            else {
                Update-Status "Ignored connection from $($connectionEvent.Program)" "#808080"
            }

            # Avoid opening a burst of modal dialogs in one timer tick.
            break
        }

        $expired = @($Script:SeenConnectionEvents.GetEnumerator() | Where-Object {
            ((Get-Date) - $_.Value).TotalMinutes -ge 5
        })
        foreach ($entry in $expired) {
            $Script:SeenConnectionEvents.Remove($entry.Key)
        }
    }
    catch {
        Stop-ConnectionMonitor -ErrorMessage "Connection monitor stopped: $($_.Exception.Message)"
    }
}

function Start-ConnectionMonitor {
    if ($Script:ConnectionMonitorTimer) {
        return
    }

    $Script:ConnectionMonitorLastTime = (Get-Date).AddSeconds(-5)
    $Script:SeenConnectionEvents = @{}
    $Script:ConnectionMonitorTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:ConnectionMonitorTimer.Interval = New-Object System.TimeSpan(0, 0, 3)
    $Script:ConnectionMonitorTimer.Add_Tick({ Process-ConnectionMonitorTick })
    $Script:ConnectionMonitorTimer.Start()
    $btnConnectionMonitor.Content = "Monitor On"
    Update-Status "Connection monitor enabled (Security events 5156/5157)" "#00FF00"
}

function Stop-ConnectionMonitor {
    param([string]$ErrorMessage = "")

    if ($Script:ConnectionMonitorTimer) {
        $Script:ConnectionMonitorTimer.Stop()
        $Script:ConnectionMonitorTimer = $null
    }
    $btnConnectionMonitor.Content = "Monitor Off"
    if ($ErrorMessage) {
        Update-Status $ErrorMessage "#FFA500"
    }
    else {
        Update-Status "Connection monitor disabled" "#808080"
    }
}

function Toggle-ConnectionMonitor {
    if ($Script:ConnectionMonitorTimer) {
        Stop-ConnectionMonitor
    }
    else {
        Start-ConnectionMonitor
    }
}

function Get-RuleTags {
    param([string]$Description)

    if ([string]::IsNullOrWhiteSpace($Description)) {
        return ""
    }

    $match = [regex]::Match($Description, "\[Tags:\s*(?<tags>[^\]]+)\]")
    if ($match.Success) {
        return $match.Groups["tags"].Value.Trim()
    }

    return ""
}

function Set-RuleTagsInDescription {
    param(
        [string]$Description,
        [string]$Tags
    )

    $base = if ($Description) { $Description } else { "" }
    $base = [regex]::Replace($base, "\s*\[Tags:\s*[^\]]+\]\s*", " ").Trim()

    if ([string]::IsNullOrWhiteSpace($Tags)) {
        return $base
    }

    if ([string]::IsNullOrWhiteSpace($base)) {
        return "[Tags: $Tags]"
    }

    return "$base [Tags: $Tags]"
}

function Get-VirtualizationScope {
    param(
        [object]$Rule,
        [object]$InterfaceFilter
    )

    $source = @($Rule.DisplayName, $Rule.Group, $Rule.Description, $(if ($InterfaceFilter) { $InterfaceFilter.InterfaceAlias } else { "" })) -join " "
    if ($source -match "WSL|Windows Subsystem|vEthernet|Hyper-V|HNS") {
        return "Hyper-V/WSL"
    }

    if ($InterfaceFilter -and $InterfaceFilter.InterfaceType -and $InterfaceFilter.InterfaceType.ToString() -ne "Any") {
        return $InterfaceFilter.InterfaceType.ToString()
    }

    return ""
}

function Test-RuleTextMatch {
    param(
        [object]$Rule,
        [string]$SearchText,
        [bool]$UseRegex
    )

    $fields = @(
        $Rule.DisplayName, $Rule.Program, $Rule.LocalPort, $Rule.RemotePort, $Rule.RemoteAddress,
        $Rule.Description, $Rule.Group, $Rule.Service, $Rule.Tags, $Rule.PolicySource
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

function Get-SelectedProfiles {
    $profiles = @()
    if ($chkProfileDomain.IsChecked) { $profiles += "Domain" }
    if ($chkProfilePrivate.IsChecked) { $profiles += "Private" }
    if ($chkProfilePublic.IsChecked) { $profiles += "Public" }

    if ($profiles.Count -eq 0 -or $profiles.Count -eq 3) {
        return "Any"
    }

    return ($profiles -join ", ")
}

function Set-ProfileCheckboxes {
    param([string]$Profile)

    $profileText = if ($Profile) { $Profile } else { "Any" }
    $isAny = $profileText -eq "Any" -or $profileText -eq "Domain, Private, Public"
    $chkProfileDomain.IsChecked = $isAny -or $profileText -match "Domain"
    $chkProfilePrivate.IsChecked = $isAny -or $profileText -match "Private"
    $chkProfilePublic.IsChecked = $isAny -or $profileText -match "Public"
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

        Update-Status "Fetching address filters..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Fetching address filters..." -ForegroundColor Gray
        $addressFilters = @{}
        Get-NetFirewallAddressFilter -ErrorAction SilentlyContinue | ForEach-Object {
            $addressFilters[$_.InstanceID] = $_
        }

        Update-Status "Fetching service filters..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Fetching service filters..." -ForegroundColor Gray
        $serviceFilters = @{}
        Get-NetFirewallServiceFilter -ErrorAction SilentlyContinue | ForEach-Object {
            $serviceFilters[$_.InstanceID] = $_
        }

        Update-Status "Fetching interface filters..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Fetching interface filters..." -ForegroundColor Gray
        $interfaceFilters = @{}
        Get-NetFirewallInterfaceFilter -ErrorAction SilentlyContinue | ForEach-Object {
            $interfaceFilters[$_.InstanceID] = $_
        }

        Update-Status "Processing $ruleCount rules..." "#FFA500"
        $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        Write-Host "  Processing rules..." -ForegroundColor Gray

        $rules = $rawRules | ForEach-Object {
            $ruleId = $_.Name
            $portFilter = $portFilters[$_.InstanceID]
            $appFilter = $appFilters[$_.InstanceID]
            $addressFilter = $addressFilters[$_.InstanceID]
            $serviceFilter = $serviceFilters[$_.InstanceID]
            $interfaceFilter = $interfaceFilters[$_.InstanceID]
            $policySource = if ($_.PolicyStoreSource) { $_.PolicyStoreSource } else { "Local" }
            $policySourceType = if ($_.PolicyStoreSourceType) { $_.PolicyStoreSourceType.ToString() } else { "Local" }
            $isGpo = $policySourceType -match "GroupPolicy|GPO" -or ($policySource -and $policySource -ne "PersistentStore" -and $policySource -ne "Local")

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
                RemoteAddress = if ($addressFilter -and $addressFilter.RemoteAddress) { $addressFilter.RemoteAddress } else { "Any" }
                LocalAddress = if ($addressFilter -and $addressFilter.LocalAddress) { $addressFilter.LocalAddress } else { "Any" }
                Program = if ($appFilter -and $appFilter.Program) { $appFilter.Program } else { "Any" }
                Service = if ($serviceFilter -and $serviceFilter.Service) { $serviceFilter.Service } else { "Any" }
                Group = if ($_.Group) { $_.Group } else { "" }
                PolicySource = $policySource
                PolicySourceType = $policySourceType
                IsGpo = $isGpo
                VirtualizationScope = Get-VirtualizationScope -Rule $_ -InterfaceFilter $interfaceFilter
                Tags = Get-RuleTags -Description $_.Description
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

function New-LockdownSnapshot {
    if (-not (Test-Path -LiteralPath $Script:BackupFolder)) {
        New-Item -ItemType Directory -Path $Script:BackupFolder -Force | Out-Null
    }

    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        $netshResult = netsh advfirewall export $tempFile 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tempFile)) {
            throw "netsh export failed: $netshResult"
        }

        $snapshot = [ordered]@{
            BackupDate = (Get-Date).ToString("o")
            ComputerName = $env:COMPUTERNAME
            RuleCount = @($Script:AllRules).Count
            NetshBackup = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($tempFile))
            RuleDetails = @($Script:AllRules)
            LockdownSnapshot = $true
            ProfileDefaults = @(Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name, DefaultInboundAction, DefaultOutboundAction)
        }
        $path = Join-Path $Script:BackupFolder "FirewallLockdown_$(Get-Date -Format 'yyyyMMdd_HHmmss').fwbackup"
        $snapshot | ConvertTo-Json -Depth 20 -Compress | Set-Content -LiteralPath $path -Encoding UTF8
        return $path
    }
    finally {
        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-LockdownAllowRule {
    param(
        [string]$Name,
        [ValidateSet("TCP", "UDP")]
        [string]$Protocol,
        [string]$RemotePort,
        [string]$LocalPort = "",
        [string]$Service
    )

    $params = @{
        DisplayName = "$Script:LockdownRulePrefix$Name"
        Direction = "Outbound"
        Action = "Allow"
        Profile = "Any"
        Enabled = "True"
        Protocol = $Protocol
        RemotePort = $RemotePort
        Description = "Essential service exception created by FirewallForge outbound lockdown"
        ErrorAction = "Stop"
    }
    if ($LocalPort) {
        $params.LocalPort = $LocalPort
    }
    if ($Service) {
        $params.Service = $Service
    }

    New-NetFirewallRule @params | Out-Null
}

function Restore-LockdownSnapshot {
    param(
        [string]$SnapshotPath = "",
        [switch]$SkipConfirmation
    )

    if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
        if ($Script:LastLockdownSnapshotPath -and (Test-Path -LiteralPath $Script:LastLockdownSnapshotPath)) {
            $SnapshotPath = $Script:LastLockdownSnapshotPath
        }
        else {
            $SnapshotPath = Get-ChildItem -LiteralPath $Script:BackupFolder -Filter "FirewallLockdown_*.fwbackup" -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
        }
    }

    if ([string]::IsNullOrWhiteSpace($SnapshotPath) -or -not (Test-Path -LiteralPath $SnapshotPath)) {
        throw "No outbound lockdown snapshot was found in $Script:BackupFolder."
    }

    $snapshot = Get-Content -LiteralPath $SnapshotPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace($snapshot.NetshBackup)) {
        throw "The selected file is not a valid FirewallForge lockdown snapshot."
    }

    if (-not $SkipConfirmation) {
        $confirm = [System.Windows.MessageBox]::Show(
            "Restore the firewall state captured on $($snapshot.BackupDate)?`n`nSnapshot: $SnapshotPath`n`nThis replaces the current firewall configuration.",
            "Confirm Lockdown Rollback",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            Update-Status "Lockdown rollback cancelled" "#808080"
            return
        }
    }

    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        Update-Status "Restoring outbound lockdown snapshot..." "#FFA500"
        $netshBytes = [Convert]::FromBase64String($snapshot.NetshBackup)
        [System.IO.File]::WriteAllBytes($tempFile, $netshBytes)
        $netshResult = netsh advfirewall import $tempFile 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "netsh import failed: $netshResult"
        }

        $Script:LastLockdownSnapshotPath = $SnapshotPath
        Update-Status "Firewall state restored from lockdown snapshot" "#00FF00"
        if ($Window.IsVisible) {
            Get-FirewallRules
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Enable-OutboundLockdown {
    $confirm = [System.Windows.MessageBox]::Show(
        "Set the Domain, Private, and Public firewall profiles to block outbound traffic?`n`nFirewallForge will first save a complete rollback snapshot and add essential DNS, DHCP, time-sync, and update exceptions.",
        "Confirm Outbound Lockdown",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
        Update-Status "Outbound lockdown cancelled" "#808080"
        return
    }

    $snapshotPath = $null
    try {
        Update-Status "Saving rollback snapshot..." "#FFA500"
        $snapshotPath = New-LockdownSnapshot
        $Script:LastLockdownSnapshotPath = $snapshotPath

        # Replace only rules previously created by this workflow.
        @(Get-NetFirewallRule -DisplayName "$Script:LockdownRulePrefix*" -ErrorAction SilentlyContinue) |
            Remove-NetFirewallRule -ErrorAction SilentlyContinue

        Set-NetFirewallProfile -Profile Domain, Private, Public -DefaultOutboundAction Block -ErrorAction Stop

        New-LockdownAllowRule -Name "DNS UDP" -Protocol UDP -RemotePort "53" -Service "Dnscache"
        New-LockdownAllowRule -Name "DNS TCP" -Protocol TCP -RemotePort "53" -Service "Dnscache"
        New-LockdownAllowRule -Name "DHCP" -Protocol UDP -LocalPort "68" -RemotePort "67" -Service "Dhcp"
        New-LockdownAllowRule -Name "Time Sync" -Protocol UDP -RemotePort "123" -Service "W32Time"
        New-LockdownAllowRule -Name "Windows Update" -Protocol TCP -RemotePort "80,443" -Service "wuauserv"
        New-LockdownAllowRule -Name "BITS Updates" -Protocol TCP -RemotePort "80,443" -Service "BITS"

        Update-Status "Outbound lockdown enabled; rollback snapshot saved to $snapshotPath" "#00FF00"
        [System.Windows.MessageBox]::Show(
            "Outbound blocking is enabled for all firewall profiles.`n`nRollback snapshot:`n$snapshotPath",
            "Outbound Lockdown Enabled",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information) | Out-Null
        Get-FirewallRules
    }
    catch {
        if ($snapshotPath -and (Test-Path -LiteralPath $snapshotPath)) {
            try {
                Restore-LockdownSnapshot -SnapshotPath $snapshotPath -SkipConfirmation
            }
            catch {
                Update-Status "Lockdown failed and automatic rollback also failed: $($_.Exception.Message)" "#FF0000"
            }
        }
        Update-Status "Outbound lockdown failed: $($_.Exception.Message)" "#FF0000"
        [System.Windows.MessageBox]::Show(
            "Outbound lockdown could not be completed.`n`nError: $($_.Exception.Message)",
            "Outbound Lockdown Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Search-Rules {
    $searchText = $txtSearch.Text.Trim()
    $groupFilter = $cmbGroupFilter.SelectedItem
    $useRegex = $chkRegexSearch.IsChecked

    $filtered = $Script:AllRules

    # Apply group filter
    if ($groupFilter -and $groupFilter -ne "(All Groups)") {
        $filtered = @($filtered | Where-Object { $_.Group -eq $groupFilter })
    }

    # Apply text search
    if (-not [string]::IsNullOrEmpty($searchText)) {
        $Script:SearchActive = $true
        try {
            if ($useRegex) {
                $null = [regex]::new($searchText)
            }
            $filtered = @($filtered | Where-Object { Test-RuleTextMatch -Rule $_ -SearchText $searchText -UseRegex $useRegex })
        }
        catch {
            Update-Status "Invalid regex: $($_.Exception.Message)" "#FF0000"
            return
        }
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

        Set-ProfileCheckboxes -Profile $selectedRule.Profile

        # Set Protocol
        $protocolMap = @{ "Any" = 0; "TCP" = 1; "UDP" = 2; "ICMPv4" = 3; "ICMPv6" = 4 }
        $cmbProtocol.SelectedIndex = if ($protocolMap.ContainsKey($selectedRule.Protocol)) { $protocolMap[$selectedRule.Protocol] } else { 0 }

        # Set Ports
        $txtLocalPort.Text = $selectedRule.LocalPort
        $txtRemotePort.Text = $selectedRule.RemotePort
        $txtRemoteAddress.Text = $selectedRule.RemoteAddress
        $txtService.Text = $selectedRule.Service
        $txtTags.Text = $selectedRule.Tags

        $btnApplyEdit.IsEnabled = -not $selectedRule.IsGpo
        $btnDeleteRule.IsEnabled = -not $selectedRule.IsGpo

        if ($selectedRule.IsGpo) {
            Update-Status "Selected read-only GPO rule: $($selectedRule.DisplayName)" "#FFA500"
        }
        else {
            Update-Status "Selected: $($selectedRule.DisplayName)" "#0078D4"
        }
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
        if ($selectedRule.IsGpo) {
            Update-Status "GPO-delivered rules are read-only in FirewallForge" "#FFA500"
            return
        }

        Update-Status "Applying changes to: $($selectedRule.DisplayName)" "#FFA500"

        $params = @{
            Name = $selectedRule.Name
        }

        # Enabled
        $params.Enabled = if ($cmbEnabled.SelectedIndex -eq 0) { "True" } else { "False" }

        # Action
        $params.Action = if ($cmbAction.SelectedIndex -eq 0) { "Allow" } else { "Block" }

        # Direction cannot be changed on existing rules, so we skip it

        $params.Profile = Get-SelectedProfiles

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

        $remoteAddress = $txtRemoteAddress.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($remoteAddress)) {
            $remoteAddress = "Any"
        }
        Get-NetFirewallRule -Name $selectedRule.Name | Set-NetFirewallAddressFilter -RemoteAddress $remoteAddress -ErrorAction Stop

        $serviceName = $txtService.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($serviceName)) {
            $serviceName = "Any"
        }
        Get-NetFirewallRule -Name $selectedRule.Name | Set-NetFirewallServiceFilter -Service $serviceName -ErrorAction Stop

        $newDescription = Set-RuleTagsInDescription -Description $selectedRule.Description -Tags $txtTags.Text.Trim()
        Set-NetFirewallRule -Name $selectedRule.Name -Description $newDescription -ErrorAction Stop

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

    $deleted = 0
    $failed = 0
    $skippedGpo = 0

    foreach ($rule in $selectedRules) {
        if ($rule.IsGpo) {
            $skippedGpo++
            continue
        }

        try {
            Remove-NetFirewallRule -Name $rule.Name -ErrorAction Stop
            $deleted++
        }
        catch {
            $failed++
        }
    }

    if ($failed -eq 0 -and $skippedGpo -eq 0) {
        Update-Status "Deleted $deleted rule(s) successfully" "#00FF00"
    }
    else {
        Update-Status "Deleted $deleted rule(s), $failed failed, $skippedGpo GPO skipped" "#FFA500"
    }

    Get-FirewallRules
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

function Show-ProgramRuleWizard {
    $fileBrowser = New-Object Microsoft.Win32.OpenFileDialog
    $fileBrowser.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
    $fileBrowser.Title = "Select a program"

    if (-not $fileBrowser.ShowDialog()) {
        return
    }

    $programPath = $fileBrowser.FileName
    $programName = [System.IO.Path]::GetFileNameWithoutExtension($programPath)
    $matchingRules = @($Script:AllRules | Where-Object {
        $_.Program -and $_.Program -ne "Any" -and $_.Program -ieq $programPath
    })

    $summary = New-Object System.Text.StringBuilder
    [void]$summary.AppendLine("Program: $programPath")
    [void]$summary.AppendLine("Existing rules: $($matchingRules.Count)")
    [void]$summary.AppendLine("")
    if ($matchingRules.Count -eq 0) {
        [void]$summary.AppendLine("No existing rules match this program path.")
    }
    else {
        foreach ($rule in $matchingRules) {
            $state = if ($rule.Enabled -eq "True") { "Enabled" } else { "Disabled" }
            [void]$summary.AppendLine("- $($rule.DisplayName) [$($rule.Direction), $($rule.Action), $state, Profile: $($rule.Profile)]")
        }
    }

    $wizard = New-Object System.Windows.Window
    $wizard.Title = "Program Rule Wizard - $programName"
    $wizard.Width = 720
    $wizard.Height = 520
    $wizard.WindowStartupLocation = "CenterOwner"
    $wizard.Owner = $Window
    $wizard.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))

    $details = New-Object System.Windows.Controls.TextBox
    $details.Text = $summary.ToString()
    $details.IsReadOnly = $true
    $details.AcceptsReturn = $true
    $details.TextWrapping = "Wrap"
    $details.VerticalScrollBarVisibility = "Auto"
    $details.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26))
    $details.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $details.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $details.FontSize = 12
    [System.Windows.Controls.Grid]::SetRow($details, 0)
    $grid.Children.Add($details) | Out-Null

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    $buttons.Margin = New-Object System.Windows.Thickness(0, 12, 0, 0)

    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Width = 100
    $closeButton.Add_Click({ $wizard.Close() }.GetNewClosure())
    $buttons.Children.Add($closeButton) | Out-Null

    $createPreset = {
        param([string]$Preset)

        $confirm = [System.Windows.MessageBox]::Show(
            "Create the '$Preset' firewall preset for:`n$programPath`n`nExisting matching rules: $($matchingRules.Count)",
            "Create Program Rules",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }

        try {
            $directions = switch ($Preset) {
                "Block" { @("Inbound", "Outbound") }
                "Allow In" { @("Inbound") }
                "Allow Out" { @("Outbound") }
            }
            $action = if ($Preset -eq "Block") { "Block" } else { "Allow" }
            foreach ($direction in $directions) {
                $displayName = "FirewallForge - $action $programName ($direction)"
                New-NetFirewallRule -DisplayName $displayName -Direction $direction -Action $action `
                    -Program $programPath -Profile Any -Enabled True `
                    -Description "Created by FirewallForge Program Wizard" -ErrorAction Stop | Out-Null
            }

            Update-Status "Created $Preset rule preset for $programName" "#00FF00"
            $wizard.Close()
            Get-FirewallRules
        }
        catch {
            Update-Status "Program wizard failed: $($_.Exception.Message)" "#FF0000"
            [System.Windows.MessageBox]::Show(
                "Failed to create program rules.`n`nError: $($_.Exception.Message)",
                "Program Wizard Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    }.GetNewClosure()

    foreach ($preset in @(
        @{ Content = "Block In + Out"; Value = "Block"; Style = "DangerButton"; Width = 135 },
        @{ Content = "Allow In"; Value = "Allow In"; Style = "SuccessButton"; Width = 100 },
        @{ Content = "Allow Out"; Value = "Allow Out"; Style = "SuccessButton"; Width = 105 }
    )) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = $preset.Content
        $button.Width = $preset.Width
        if ($preset.Style -eq "DangerButton") {
            $button.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xD3, 0x2F, 0x2F))
        }
        else {
            $button.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x38, 0x8E, 0x3C))
        }
        $button.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
        $value = $preset.Value
        $button.Add_Click({ & $createPreset $value }.GetNewClosure())
        $buttons.Children.Insert(0, $button)
    }

    [System.Windows.Controls.Grid]::SetRow($buttons, 1)
    $grid.Children.Add($buttons) | Out-Null
    $wizard.Content = $grid
    $wizard.ShowDialog() | Out-Null
}

function Test-RulePortMatch {
    param(
        [string]$RulePort,
        [string]$FlowPort
    )

    if ([string]::IsNullOrWhiteSpace($RulePort) -or $RulePort -eq "Any" -or $RulePort -eq "*") {
        return $true
    }
    if ([string]::IsNullOrWhiteSpace($FlowPort) -or $FlowPort -eq "Any" -or $FlowPort -eq "*") {
        return $true
    }
    if ($FlowPort -notmatch '^\d+$') {
        return $false
    }

    $port = [int]$FlowPort
    foreach ($part in ([string]$RulePort -split ',')) {
        $candidate = $part.Trim()
        if ($candidate -match '^(\d+)\s*-\s*(\d+)$') {
            if ($port -ge [int]$Matches[1] -and $port -le [int]$Matches[2]) {
                return $true
            }
        }
        elseif ($candidate -match '^\d+$' -and $port -eq [int]$candidate) {
            return $true
        }
    }

    return $false
}

function Test-RuleAddressMatch {
    param(
        [string]$RuleAddress,
        [string]$FlowAddress
    )

    if ([string]::IsNullOrWhiteSpace($RuleAddress) -or $RuleAddress -eq "Any" -or $RuleAddress -eq "*") {
        return $true
    }
    if ([string]::IsNullOrWhiteSpace($FlowAddress) -or $FlowAddress -eq "Any" -or $FlowAddress -eq "*") {
        return $true
    }

    try {
        $flowIp = [System.Net.IPAddress]::Parse($FlowAddress)
    }
    catch {
        return @($RuleAddress -split ',') | Where-Object { $_.Trim() -ieq $FlowAddress } | Select-Object -First 1
    }

    foreach ($candidate in ([string]$RuleAddress -split ',')) {
        $address = $candidate.Trim()
        if ($address -ieq $FlowAddress) {
            return $true
        }
        if ($address -notmatch '^(.+)/(\d{1,3})$') {
            continue
        }

        try {
            $network = [System.Net.IPAddress]::Parse($Matches[1])
            $prefix = [int]$Matches[2]
            $networkBytes = $network.GetAddressBytes()
            $flowBytes = $flowIp.GetAddressBytes()
            if ($networkBytes.Length -ne $flowBytes.Length -or $prefix -lt 0 -or $prefix -gt ($networkBytes.Length * 8)) {
                continue
            }

            $fullBytes = [math]::Floor($prefix / 8)
            $remainingBits = $prefix % 8
            $matchesNetwork = $true
            for ($i = 0; $i -lt $fullBytes; $i++) {
                if ($networkBytes[$i] -ne $flowBytes[$i]) {
                    $matchesNetwork = $false
                    break
                }
            }
            if ($matchesNetwork -and $remainingBits -gt 0) {
                $mask = [byte](0xFF -shl (8 - $remainingBits))
                if (($networkBytes[$fullBytes] -band $mask) -ne ($flowBytes[$fullBytes] -band $mask)) {
                    $matchesNetwork = $false
                }
            }
            if ($matchesNetwork) {
                return $true
            }
        }
        catch {
            continue
        }
    }

    return $false
}

function Test-FirewallRuleFlowMatch {
    param(
        [object]$Rule,
        [object]$Flow
    )

    if ($Rule.Enabled -eq "False" -or $Rule.Direction -ne $Flow.Direction) {
        return $false
    }
    if ($Rule.Profile -and $Rule.Profile -ne "Any" -and $Rule.Profile -notmatch [regex]::Escape($Flow.Profile)) {
        return $false
    }
    if ($Rule.Program -and $Rule.Program -ne "Any" -and $Flow.Program -ne "Any" -and $Rule.Program -ine $Flow.Program) {
        return $false
    }
    if ($Rule.Protocol -and $Rule.Protocol -ne "Any" -and $Flow.Protocol -ne "Any" -and $Rule.Protocol -ine $Flow.Protocol) {
        return $false
    }
    if (-not (Test-RulePortMatch -RulePort $Rule.LocalPort -FlowPort $Flow.LocalPort)) {
        return $false
    }
    if (-not (Test-RulePortMatch -RulePort $Rule.RemotePort -FlowPort $Flow.RemotePort)) {
        return $false
    }
    if (-not (Test-RuleAddressMatch -RuleAddress $Rule.RemoteAddress -FlowAddress $Flow.RemoteAddress)) {
        return $false
    }

    return $true
}

function Get-FirewallRuleSpecificity {
    param([object]$Rule)

    $score = 10 # Direction is an exact predicate for every candidate.
    if ($Rule.Profile -and $Rule.Profile -ne "Any") { $score += 4 }
    if ($Rule.Program -and $Rule.Program -ne "Any") { $score += 8 }
    if ($Rule.Protocol -and $Rule.Protocol -ne "Any") { $score += 3 }
    if ($Rule.LocalPort -and $Rule.LocalPort -ne "Any") { $score += 3 }
    if ($Rule.RemotePort -and $Rule.RemotePort -ne "Any") { $score += 3 }
    if ($Rule.RemoteAddress -and $Rule.RemoteAddress -ne "Any") { $score += 5 }
    if ($Rule.Service -and $Rule.Service -ne "Any") { $score += 3 }
    if ($Rule.IsGpo) { $score += 2 }
    if ($Rule.Action -eq "Block") { $score += 1 }
    return $score
}

function Show-RulePriorityInput {
    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Test Firewall Rule Priority"
    $dialog.Width = 620
    $dialog.Height = 520
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $dialog.Tag = $null

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    for ($rowIndex = 0; $rowIndex -lt 9; $rowIndex++) {
        $rowDefinition = New-Object System.Windows.Controls.RowDefinition
        $rowDefinition.Height = [System.Windows.GridLength]::Auto
        $grid.RowDefinitions.Add($rowDefinition)
    }
    $labelColumn = New-Object System.Windows.Controls.ColumnDefinition
    $labelColumn.Width = New-Object System.Windows.GridLength(140)
    $grid.ColumnDefinitions.Add($labelColumn)
    $inputColumn = New-Object System.Windows.Controls.ColumnDefinition
    $inputColumn.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.ColumnDefinitions.Add($inputColumn)

    $controls = @{}
    $fieldDefinitions = @(
        @{ Label = "Direction"; Name = "Direction"; Type = "Combo"; Values = @("Inbound", "Outbound"); Default = "Outbound" }
        @{ Label = "Profile"; Name = "Profile"; Type = "Combo"; Values = @("Domain", "Private", "Public"); Default = "Private" }
        @{ Label = "Program"; Name = "Program"; Type = "Text"; Default = "Any" }
        @{ Label = "Protocol"; Name = "Protocol"; Type = "Combo"; Values = @("Any", "TCP", "UDP", "ICMPv4", "ICMPv6"); Default = "TCP" }
        @{ Label = "Local port"; Name = "LocalPort"; Type = "Text"; Default = "Any" }
        @{ Label = "Remote port"; Name = "RemotePort"; Type = "Text"; Default = "443" }
        @{ Label = "Remote address"; Name = "RemoteAddress"; Type = "Text"; Default = "Any" }
    )

    for ($i = 0; $i -lt $fieldDefinitions.Count; $i++) {
        $definition = $fieldDefinitions[$i]
        $label = New-Object System.Windows.Controls.Label
        $label.Content = $definition.Label
        [System.Windows.Controls.Grid]::SetRow($label, $i)
        [System.Windows.Controls.Grid]::SetColumn($label, 0)
        $grid.Children.Add($label) | Out-Null

        if ($definition.Type -eq "Combo") {
            $control = New-Object System.Windows.Controls.ComboBox
            foreach ($value in $definition.Values) {
                $control.Items.Add($value) | Out-Null
            }
            $control.SelectedItem = $definition.Default
        }
        else {
            $control = New-Object System.Windows.Controls.TextBox
            $control.Text = $definition.Default
        }
        $control.Margin = New-Object System.Windows.Thickness(5, 3, 5, 3)
        $controls[$definition.Name] = $control
        [System.Windows.Controls.Grid]::SetRow($control, $i)
        [System.Windows.Controls.Grid]::SetColumn($control, 1)
        $grid.Children.Add($control) | Out-Null
    }

    $reachability = New-Object System.Windows.Controls.CheckBox
    $reachability.Content = "Run Test-NetConnection for a concrete remote address and port"
    $reachability.Margin = New-Object System.Windows.Thickness(5, 8, 5, 8)
    [System.Windows.Controls.Grid]::SetRow($reachability, 7)
    [System.Windows.Controls.Grid]::SetColumn($reachability, 1)
    $grid.Children.Add($reachability) | Out-Null

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    [System.Windows.Controls.Grid]::SetRow($buttons, 8)
    [System.Windows.Controls.Grid]::SetColumn($buttons, 1)
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Width = 100
    $cancel.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($cancel) | Out-Null
    $evaluate = New-Object System.Windows.Controls.Button
    $evaluate.Content = "Evaluate Rules"
    $evaluate.Width = 140
    $evaluate.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $evaluate.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $evaluate.Add_Click({
        $dialog.Tag = [PSCustomObject]@{
            Direction = [string]$controls.Direction.SelectedItem
            Profile = [string]$controls.Profile.SelectedItem
            Program = [string]$controls.Program.Text.Trim()
            Protocol = [string]$controls.Protocol.SelectedItem
            LocalPort = [string]$controls.LocalPort.Text.Trim()
            RemotePort = [string]$controls.RemotePort.Text.Trim()
            RemoteAddress = [string]$controls.RemoteAddress.Text.Trim()
            RunReachability = [bool]$reachability.IsChecked
        }
        $dialog.Close()
    }.GetNewClosure())
    $buttons.Children.Add($evaluate) | Out-Null
    $grid.Children.Add($buttons) | Out-Null

    $dialog.Content = $grid
    $dialog.ShowDialog() | Out-Null
    return $dialog.Tag
}

function Test-FirewallRulePriority {
    $flow = Show-RulePriorityInput
    if ($null -eq $flow) {
        return
    }

    if ($flow.Program -eq "") { $flow.Program = "Any" }
    if ($flow.LocalPort -eq "") { $flow.LocalPort = "Any" }
    if ($flow.RemotePort -eq "") { $flow.RemotePort = "Any" }
    if ($flow.RemoteAddress -eq "") { $flow.RemoteAddress = "Any" }

    $matchingRules = @($Script:AllRules | Where-Object {
        Test-FirewallRuleFlowMatch -Rule $_ -Flow $flow
    } | Sort-Object @{ Expression = { Get-FirewallRuleSpecificity -Rule $_ }; Descending = $true }, @{ Expression = { $_.Action -eq "Block" }; Descending = $true })

    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("FIREWALL RULE PRIORITY VIEW (APPROXIMATE)")
    [void]$report.AppendLine("=" * 72)
    [void]$report.AppendLine("Flow: $($flow.Direction) / $($flow.Protocol) / $($flow.Program)")
    [void]$report.AppendLine("Profile: $($flow.Profile), Local port: $($flow.LocalPort), Remote: $($flow.RemoteAddress):$($flow.RemotePort)")
    [void]$report.AppendLine("")

    if ($matchingRules.Count -eq 0) {
        [void]$report.AppendLine("No enabled rules in the loaded rule set match this flow.")
    }
    else {
        [void]$report.AppendLine("Matching rules (highest estimated specificity first):")
        $rank = 0
        foreach ($rule in $matchingRules) {
            $rank++
            $gpo = if ($rule.IsGpo) { ", GPO" } else { "" }
            [void]$report.AppendLine(("{0,2}. [{1}] {2} (score {3}{4})" -f $rank, $rule.Action.ToUpperInvariant(), $rule.DisplayName, (Get-FirewallRuleSpecificity -Rule $rule), $gpo))
            [void]$report.AppendLine("    $($rule.Direction), Profile=$($rule.Profile), Protocol=$($rule.Protocol), Program=$($rule.Program), Ports=$($rule.LocalPort)->$($rule.RemotePort), Remote=$($rule.RemoteAddress)")
        }
        $winner = $matchingRules[0]
        [void]$report.AppendLine("")
        [void]$report.AppendLine("Estimated winning action: $($winner.Action) via '$($winner.DisplayName)'")
    }

    if ($flow.RunReachability -and $flow.RemoteAddress -notmatch "^(Any|\*|$)" -and $flow.RemoteAddress -notmatch "[,/]" -and $flow.RemotePort -match '^\d+$') {
        try {
            $reachable = Test-NetConnection -ComputerName $flow.RemoteAddress -Port ([int]$flow.RemotePort) -InformationLevel Quiet -WarningAction SilentlyContinue
            [void]$report.AppendLine("")
            [void]$report.AppendLine("Test-NetConnection: $(if ($reachable) { 'reachable' } else { 'not reachable' })")
        }
        catch {
            [void]$report.AppendLine("")
            [void]$report.AppendLine("Test-NetConnection failed: $($_.Exception.Message)")
        }
    }
    else {
        [void]$report.AppendLine("")
        [void]$report.AppendLine("Test-NetConnection: not run (use a concrete address/port and enable the option to run it).")
    }

    [void]$report.AppendLine("")
    [void]$report.AppendLine("Windows can apply additional precedence from policy stores, GPO, interface/service filters, and block-rule semantics; use this list as an explainable predicate walk, not a kernel verdict.")
    Show-ReportWindow -Title "Firewall Rule Priority" -Text $report.ToString() -Width 900 -Height 650
}

function Show-GroupOperations {
    $groups = @($Script:AllRules | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Group) } |
        Select-Object -ExpandProperty Group -Unique | Sort-Object)
    if (@($Script:AllRules | Where-Object { [string]::IsNullOrWhiteSpace($_.Group) }).Count -gt 0) {
        $groups += "(Ungrouped)"
    }

    if ($groups.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No rule groups are available in the loaded firewall rules.",
            "Group Operations",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Firewall Group Operations"
    $dialog.Width = 560
    $dialog.Height = 300
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = New-Object System.Windows.Thickness(20)

    $heading = New-Object System.Windows.Controls.TextBlock
    $heading.Text = "Enable, disable, or delete every rule in a group"
    $heading.FontSize = 18
    $heading.FontWeight = "Bold"
    $heading.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $heading.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($heading) | Out-Null

    $groupLabel = New-Object System.Windows.Controls.Label
    $groupLabel.Content = "Group"
    $panel.Children.Add($groupLabel) | Out-Null
    $groupCombo = New-Object System.Windows.Controls.ComboBox
    foreach ($group in $groups) {
        $groupCombo.Items.Add($group) | Out-Null
    }
    $groupCombo.SelectedIndex = 0
    $groupCombo.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $panel.Children.Add($groupCombo) | Out-Null

    $operationLabel = New-Object System.Windows.Controls.Label
    $operationLabel.Content = "Operation"
    $panel.Children.Add($operationLabel) | Out-Null
    $operationCombo = New-Object System.Windows.Controls.ComboBox
    foreach ($operation in @("Enable", "Disable", "Delete")) {
        $operationCombo.Items.Add($operation) | Out-Null
    }
    $operationCombo.SelectedIndex = 0
    $operationCombo.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $panel.Children.Add($operationCombo) | Out-Null

    $countText = New-Object System.Windows.Controls.TextBlock
    $countText.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xB0, 0xB0, 0xB0))
    $countText.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $panel.Children.Add($countText) | Out-Null

    $getGroupRules = {
        $selectedGroup = [string]$groupCombo.SelectedItem
        if ($selectedGroup -eq "(Ungrouped)") {
            return @($Script:AllRules | Where-Object { [string]::IsNullOrWhiteSpace($_.Group) })
        }
        return @($Script:AllRules | Where-Object { $_.Group -eq $selectedGroup })
    }.GetNewClosure()

    $updateCount = {
        $targetRules = @(& $getGroupRules)
        $gpoCount = @($targetRules | Where-Object { $_.IsGpo }).Count
        $countText.Text = "$($targetRules.Count) rule(s) in this group; $gpoCount GPO rule(s) will be skipped."
    }.GetNewClosure()
    $groupCombo.Add_SelectionChanged({ & $updateCount }.GetNewClosure())
    & $updateCount

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Width = 100
    $cancel.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($cancel) | Out-Null
    $apply = New-Object System.Windows.Controls.Button
    $apply.Content = "Apply Operation"
    $apply.Width = 140
    $apply.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xD3, 0x2F, 0x2F))
    $apply.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $apply.Add_Click({
        $targetRules = @(& $getGroupRules)
        $operation = [string]$operationCombo.SelectedItem
        if ($targetRules.Count -eq 0) {
            return
        }

        $confirm = [System.Windows.MessageBox]::Show(
            "$operation $($targetRules.Count) rule(s) in '$($groupCombo.SelectedItem)'?`n`nGPO-delivered rules are read-only and will be skipped.",
            "Confirm Group Operation",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }

        $changed = 0
        $failed = 0
        $skipped = 0
        foreach ($rule in $targetRules) {
            if ($rule.IsGpo) {
                $skipped++
                continue
            }
            try {
                if ($operation -eq "Delete") {
                    Remove-NetFirewallRule -Name $rule.Name -ErrorAction Stop
                }
                else {
                    Set-NetFirewallRule -Name $rule.Name -Enabled ($operation -eq "Enable") -ErrorAction Stop
                }
                $changed++
            }
            catch {
                $failed++
            }
        }

        $dialog.Close()
        $status = "$operation complete: $changed changed"
        if ($skipped -gt 0) { $status += ", $skipped GPO skipped" }
        if ($failed -gt 0) { $status += ", $failed failed" }
        Update-Status $status $(if ($failed -gt 0) { "#FFA500" } else { "#00FF00" })
        Get-FirewallRules
    }.GetNewClosure())
    $buttons.Children.Add($apply) | Out-Null
    $panel.Children.Add($buttons) | Out-Null

    $dialog.Content = $panel
    $dialog.ShowDialog() | Out-Null
}

function Read-FirewallLogEntries {
    param(
        [string]$Path,
        [int]$Tail = 1000
    )

    $entries = New-Object System.Collections.Generic.List[PSObject]
    foreach ($line in @(Get-Content -LiteralPath $Path -Tail $Tail -ErrorAction Stop)) {
        if ([string]::IsNullOrWhiteSpace($line) -or $line.TrimStart().StartsWith("#")) {
            continue
        }
        $parts = @($line.Trim() -split '\s+')
        if ($parts.Count -lt 8) {
            continue
        }

        $pathField = if ($parts.Count -gt 16) { $parts[16].ToUpperInvariant() } else { "" }
        $direction = switch ($pathField) {
            "RECEIVE" { "Inbound"; break }
            "SEND" { "Outbound"; break }
            default { "Unknown" }
        }
        $entries.Add([PSCustomObject]@{
            Raw = $line.Trim()
            Date = $parts[0]
            Time = $parts[1]
            Action = $parts[2].ToUpperInvariant()
            Protocol = $parts[3].ToUpperInvariant()
            SourceAddress = $parts[4]
            DestinationAddress = $parts[5]
            SourcePort = $parts[6]
            DestinationPort = $parts[7]
            Direction = $direction
        })
    }
    return $entries
}

function Show-FirewallLogViewer {
    $logPath = Join-Path $env:windir "System32\LogFiles\Firewall\pfirewall.log"
    if (-not (Test-Path -LiteralPath $logPath)) {
        [System.Windows.MessageBox]::Show(
            "The Windows firewall log was not found at:`n$logPath`n`nEnable firewall logging in Windows Defender Firewall settings, then try again.",
            "Firewall Log Unavailable",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Firewall Log Viewer"
    $dialog.Width = 1120
    $dialog.Height = 700
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))
    $logState = @{ Path = $logPath }

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(12)
    $headerRow = New-Object System.Windows.Controls.RowDefinition
    $headerRow.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($headerRow)
    $logRow = New-Object System.Windows.Controls.RowDefinition
    $logRow.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.RowDefinitions.Add($logRow)
    $statusRow = New-Object System.Windows.Controls.RowDefinition
    $statusRow.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($statusRow)

    $toolbar = New-Object System.Windows.Controls.WrapPanel
    $toolbar.VerticalAlignment = "Center"
    $pathText = New-Object System.Windows.Controls.TextBlock
    $pathText.Text = $logPath
    $pathText.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xB0, 0xB0, 0xB0))
    $pathText.Margin = New-Object System.Windows.Thickness(0, 0, 15, 5)
    $toolbar.Children.Add($pathText) | Out-Null

    $chooseButton = New-Object System.Windows.Controls.Button
    $chooseButton.Content = "Choose Log"
    $chooseButton.Width = 100
    $toolbar.Children.Add($chooseButton) | Out-Null

    $actionCombo = New-Object System.Windows.Controls.ComboBox
    $actionCombo.Width = 110
    foreach ($value in @("All actions", "ALLOW", "DROP")) { $actionCombo.Items.Add($value) | Out-Null }
    $actionCombo.SelectedIndex = 0
    $actionCombo.Margin = New-Object System.Windows.Thickness(10, 0, 5, 5)
    $toolbar.Children.Add($actionCombo) | Out-Null

    $directionCombo = New-Object System.Windows.Controls.ComboBox
    $directionCombo.Width = 115
    foreach ($value in @("All directions", "Inbound", "Outbound", "Unknown")) { $directionCombo.Items.Add($value) | Out-Null }
    $directionCombo.SelectedIndex = 0
    $directionCombo.Margin = New-Object System.Windows.Thickness(5, 0, 5, 5)
    $toolbar.Children.Add($directionCombo) | Out-Null

    $portBox = New-Object System.Windows.Controls.TextBox
    $portBox.Width = 100
    $portBox.ToolTip = "Source or destination port"
    $portBox.Margin = New-Object System.Windows.Thickness(5, 0, 5, 5)
    $toolbar.Children.Add($portBox) | Out-Null

    $refreshButton = New-Object System.Windows.Controls.Button
    $refreshButton.Content = "Refresh"
    $refreshButton.Width = 90
    $toolbar.Children.Add($refreshButton) | Out-Null

    $autoRefresh = New-Object System.Windows.Controls.CheckBox
    $autoRefresh.Content = "Auto-refresh (3s)"
    $autoRefresh.Margin = New-Object System.Windows.Thickness(10, 0, 5, 5)
    $toolbar.Children.Add($autoRefresh) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($toolbar, 0)
    $grid.Children.Add($toolbar) | Out-Null

    $richText = New-Object System.Windows.Controls.RichTextBox
    $richText.IsReadOnly = $true
    $richText.VerticalScrollBarVisibility = "Auto"
    $richText.HorizontalScrollBarVisibility = "Auto"
    $richText.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x25, 0x25, 0x26))
    $richText.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
    $richText.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $richText.FontSize = 11
    [System.Windows.Controls.Grid]::SetRow($richText, 1)
    $grid.Children.Add($richText) | Out-Null

    $status = New-Object System.Windows.Controls.TextBlock
    $status.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xB0, 0xB0, 0xB0))
    $status.Margin = New-Object System.Windows.Thickness(0, 8, 0, 0)
    [System.Windows.Controls.Grid]::SetRow($status, 2)
    $grid.Children.Add($status) | Out-Null

    $render = {
        try {
            $entries = @(Read-FirewallLogEntries -Path $logState.Path -Tail 1000)
            $action = [string]$actionCombo.SelectedItem
            $direction = [string]$directionCombo.SelectedItem
            $port = $portBox.Text.Trim()
            $filtered = @($entries | Where-Object {
                ($action -eq "All actions" -or $_.Action -eq $action) -and
                ($direction -eq "All directions" -or $_.Direction -eq $direction) -and
                ([string]::IsNullOrWhiteSpace($port) -or $_.SourcePort -eq $port -or $_.DestinationPort -eq $port)
            })

            $document = New-Object System.Windows.Documents.FlowDocument
            $document.PagePadding = New-Object System.Windows.Thickness(6)
            foreach ($entry in $filtered) {
                $paragraph = New-Object System.Windows.Documents.Paragraph
                $paragraph.Margin = New-Object System.Windows.Thickness(0)
                $run = New-Object System.Windows.Documents.Run($entry.Raw)
                if ($entry.Action -eq "DROP") {
                    $run.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xFF, 0x66, 0x66))
                }
                elseif ($entry.Action -eq "ALLOW") {
                    $run.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x66, 0xDD, 0x88))
                }
                else {
                    $run.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xE0, 0xE0, 0xE0))
                }
                $paragraph.Inlines.Add($run) | Out-Null
                $document.Blocks.Add($paragraph) | Out-Null
            }
            $richText.Document = $document
            $status.Text = "Showing $($filtered.Count) of $($entries.Count) log entries (tail 1000)"
        }
        catch {
            $status.Text = "Could not read firewall log: $($_.Exception.Message)"
        }
    }.GetNewClosure()

    $chooseButton.Add_Click({
        $openDialog = New-Object Microsoft.Win32.OpenFileDialog
        $openDialog.Filter = "Firewall logs (*.log)|*.log|All files (*.*)|*.*"
        $openDialog.FileName = [System.IO.Path]::GetFileName($logState.Path)
        if ($openDialog.ShowDialog()) {
            $logState.Path = $openDialog.FileName
            $pathText.Text = $logState.Path
            & $render
        }
    }.GetNewClosure())
    $refreshButton.Add_Click({ & $render }.GetNewClosure())
    $actionCombo.Add_SelectionChanged({ & $render }.GetNewClosure())
    $directionCombo.Add_SelectionChanged({ & $render }.GetNewClosure())
    $portBox.Add_KeyDown({ if ($_.Key -eq [System.Windows.Input.Key]::Enter) { & $render } }.GetNewClosure())

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = New-Object System.TimeSpan(0, 0, 3)
    $timer.Add_Tick({ if ($autoRefresh.IsChecked) { & $render } }.GetNewClosure())
    $timer.Start()
    $dialog.Add_Closed({ $timer.Stop() }.GetNewClosure())

    $dialog.Content = $grid
    & $render
    $dialog.ShowDialog() | Out-Null
}

function Show-ScheduledBackupDialog {
    if (-not (Test-Path -LiteralPath $Script:ScheduledBackupWorker)) {
        [System.Windows.MessageBox]::Show(
            "The scheduled-backup worker is missing:`n$Script:ScheduledBackupWorker",
            "Scheduled Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Schedule Firewall Backups"
    $dialog.Width = 560
    $dialog.Height = 390
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Window
    $dialog.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x1E, 0x1E, 0x1E))

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    for ($rowIndex = 0; $rowIndex -lt 7; $rowIndex++) {
        $row = New-Object System.Windows.Controls.RowDefinition
        $row.Height = [System.Windows.GridLength]::Auto
        $grid.RowDefinitions.Add($row)
    }
    $labelColumn = New-Object System.Windows.Controls.ColumnDefinition
    $labelColumn.Width = New-Object System.Windows.GridLength(145)
    $grid.ColumnDefinitions.Add($labelColumn)
    $inputColumn = New-Object System.Windows.Controls.ColumnDefinition
    $inputColumn.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.ColumnDefinitions.Add($inputColumn)

    $addField = {
        param([int]$Row, [string]$LabelText, [System.Windows.Controls.Control]$Control)
        $label = New-Object System.Windows.Controls.Label
        $label.Content = $LabelText
        [System.Windows.Controls.Grid]::SetRow($label, $Row)
        [System.Windows.Controls.Grid]::SetColumn($label, 0)
        $grid.Children.Add($label) | Out-Null
        $Control.Margin = New-Object System.Windows.Thickness(5, 3, 5, 3)
        [System.Windows.Controls.Grid]::SetRow($Control, $Row)
        [System.Windows.Controls.Grid]::SetColumn($Control, 1)
        $grid.Children.Add($Control) | Out-Null
    }.GetNewClosure()

    $frequency = New-Object System.Windows.Controls.ComboBox
    foreach ($value in @("Daily", "Weekly")) { $frequency.Items.Add($value) | Out-Null }
    $frequency.SelectedIndex = 0
    & $addField 0 "Frequency" $frequency

    $day = New-Object System.Windows.Controls.ComboBox
    foreach ($value in [System.Enum]::GetNames([System.DayOfWeek])) { $day.Items.Add($value) | Out-Null }
    $day.SelectedItem = "Sunday"
    & $addField 1 "Weekly day" $day

    $time = New-Object System.Windows.Controls.TextBox
    $time.Text = "02:00"
    $time.ToolTip = "24-hour local time, for example 02:00 or 22:30"
    & $addField 2 "Time" $time

    $retention = New-Object System.Windows.Controls.TextBox
    $retention.Text = "14"
    $retention.ToolTip = "Number of scheduled backups to retain"
    & $addField 3 "Retention count" $retention

    $folder = New-Object System.Windows.Controls.TextBox
    $folder.Text = $Script:BackupFolder
    & $addField 4 "Backup folder" $folder

    $info = New-Object System.Windows.Controls.TextBlock
    $info.Text = "The task runs the non-interactive worker with highest available privileges and rotates only FirewallScheduledBackup_*.fwbackup files."
    $info.TextWrapping = "Wrap"
    $info.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xB0, 0xB0, 0xB0))
    $info.Margin = New-Object System.Windows.Thickness(5, 8, 5, 8)
    [System.Windows.Controls.Grid]::SetRow($info, 5)
    [System.Windows.Controls.Grid]::SetColumnSpan($info, 2)
    $grid.Children.Add($info) | Out-Null

    $buttons = New-Object System.Windows.Controls.WrapPanel
    $buttons.HorizontalAlignment = "Right"
    [System.Windows.Controls.Grid]::SetRow($buttons, 6)
    [System.Windows.Controls.Grid]::SetColumnSpan($buttons, 2)
    $remove = New-Object System.Windows.Controls.Button
    $remove.Content = "Remove Schedule"
    $remove.Width = 125
    $remove.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0xD3, 0x2F, 0x2F))
    $remove.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $remove.Add_Click({
        try {
            Unregister-ScheduledTask -TaskName $Script:ScheduledBackupTaskName -Confirm:$false -ErrorAction Stop
            Update-Status "Scheduled backup task removed" "#00FF00"
            $dialog.Close()
        }
        catch {
            [System.Windows.MessageBox]::Show("Could not remove the scheduled task.`n`n$($_.Exception.Message)", "Task Scheduler Error") | Out-Null
        }
    }.GetNewClosure())
    $buttons.Children.Add($remove) | Out-Null

    $runNow = New-Object System.Windows.Controls.Button
    $runNow.Content = "Run Now"
    $runNow.Width = 90
    $runNow.Add_Click({
        try {
            Start-ScheduledTask -TaskName $Script:ScheduledBackupTaskName -ErrorAction Stop
            Update-Status "Scheduled backup task started" "#00FF00"
            $dialog.Close()
        }
        catch {
            [System.Windows.MessageBox]::Show("Could not start the scheduled task. Register it first.`n`n$($_.Exception.Message)", "Task Scheduler Error") | Out-Null
        }
    }.GetNewClosure())
    $buttons.Children.Add($runNow) | Out-Null

    $register = New-Object System.Windows.Controls.Button
    $register.Content = "Register Schedule"
    $register.Width = 135
    $register.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x38, 0x8E, 0x3C))
    $register.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Colors]::White)
    $register.Add_Click({
        try {
            [datetime]$parsedTime = [datetime]::MinValue
            if (-not [datetime]::TryParse($time.Text.Trim(), [ref]$parsedTime)) {
                throw "Enter a valid time such as 02:00 or 22:30."
            }
            [int]$retentionCount = 0
            if (-not [int]::TryParse($retention.Text.Trim(), [ref]$retentionCount) -or $retentionCount -lt 1 -or $retentionCount -gt 3650) {
                throw "Retention count must be between 1 and 3650."
            }
            if ([string]::IsNullOrWhiteSpace($folder.Text)) {
                throw "Backup folder cannot be empty."
            }
            $taskAction = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `
                -Argument ('-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -BackupPath "{1}" -RetentionCount {2}' -f $Script:ScheduledBackupWorker, $folder.Text.Trim(), $retentionCount)
            $atTime = [datetime]::Today.Add($parsedTime.TimeOfDay)
            if ([string]$frequency.SelectedItem -eq "Weekly") {
                $dayOfWeek = [System.DayOfWeek]::Parse([string]$day.SelectedItem)
                $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dayOfWeek -At $atTime
            }
            else {
                $trigger = New-ScheduledTaskTrigger -Daily -At $atTime
            }
            $principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable
            Register-ScheduledTask -TaskName $Script:ScheduledBackupTaskName -Action $taskAction -Trigger $trigger -Principal $principal -Settings $settings -Description "FirewallForge scheduled firewall backup" -Force -ErrorAction Stop | Out-Null
            Update-Status "Scheduled backup registered ($($frequency.SelectedItem) at $($time.Text))" "#00FF00"
            $dialog.Close()
        }
        catch {
            [System.Windows.MessageBox]::Show("Could not register the scheduled task.`n`n$($_.Exception.Message)", "Task Scheduler Error") | Out-Null
        }
    }.GetNewClosure())
    $buttons.Children.Add($register) | Out-Null
    $close = New-Object System.Windows.Controls.Button
    $close.Content = "Cancel"
    $close.Width = 90
    $close.Add_Click({ $dialog.Close() }.GetNewClosure())
    $buttons.Children.Add($close) | Out-Null
    $grid.Children.Add($buttons) | Out-Null

    $dialog.Content = $grid
    $dialog.ShowDialog() | Out-Null
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
$btnProgramWizard.Add_Click({ Show-ProgramRuleWizard })
$btnConnectionMonitor.Add_Click({ Toggle-ConnectionMonitor })
$btnOutboundLockdown.Add_Click({ Enable-OutboundLockdown })
$btnRollbackLockdown.Add_Click({
    try {
        Restore-LockdownSnapshot
    }
    catch {
        Update-Status "Lockdown rollback failed: $($_.Exception.Message)" "#FF0000"
        [System.Windows.MessageBox]::Show(
            "Could not roll back the outbound lockdown.`n`nError: $($_.Exception.Message)",
            "Lockdown Rollback Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})
$btnRulePriority.Add_Click({ Test-FirewallRulePriority })
$btnGroupOps.Add_Click({ Show-GroupOperations })
$btnLogViewer.Add_Click({ Show-FirewallLogViewer })
$btnScheduleBackups.Add_Click({ Show-ScheduledBackupDialog })
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

    $Window.Add_Closed({
        if ($Script:ConnectionMonitorTimer) {
            $Script:ConnectionMonitorTimer.Stop()
            $Script:ConnectionMonitorTimer = $null
        }
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
