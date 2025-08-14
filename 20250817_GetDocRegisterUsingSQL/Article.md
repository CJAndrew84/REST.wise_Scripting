# From REST to Raw Speed: Building a ProjectWise Doc Register with SQL (and a Friendly UI)

TL;DR: 
Last time, I pulled a ProjectWise document list using the WSG REST API — reliable, supported, and portable. This time, I’m going straight to the database with SQL via PowerShell for speed, version control, and a better UI. Bonus: a WPF interface with icons, buttons, and expandable rows.

In this post I show you the code to build your own UI (customise the one I have included below), and use a SQL SELECT statement to grab data directly from the database tables and join tables and display in your nice UI.  

Fair warning, this article contains a lot of code.

Minimize image
Edit image
Delete image

Add a caption (optional)
The full code should:

Lets you select a datasource from a TreeView

Lets you browse rich projects inside that datasource

Runs a read-only SQL query to get all docs + versions under the selected project

Groups results by document and shows version history inline

Exports the register to Excel with one click

Use SQL if you have DB access and need results fast. Use WSG if you want official, portable, no-DB-required queries.

Minimize image
Edit image
Delete image

Add a caption (optional)

The Story (Why)
Last issue, I pulled a document list with the WSG REST API. It was tidy, but not very fast.

And time is money… we want raw speed. No paging. No property decoding. No waiting for the API and drinking coffee.

Minimize image
Edit image
Delete image

Add a caption (optional)
This time I’m going straight to the source — the ProjectWise database — with a single SQL statement. And yes, we’re doing it in PowerShell. And yes, we’re giving it buttons. And icons. And expandable rows.

The result: a one-click document register to Excel, with version history baked in. You choose a datasource, expand a project, click a button, and out pops a register. No paging. No loops. No “let me decode that WSG payload for you” step.

Minimize image
Edit image
Delete image

Add a caption (optional)

What we’re building
A PowerShell script with a WPF UI

A direct SQL query (read-only) that targets your chosen rich project + subfolders

A grouped register with expandable version rows

An Excel export with project number + timestamp in the filename


When to use this vs WSG (the 10-second version)
Use SQL when you have DB access and need speed + full control over joins/columns.

Use WSG when you want officially supported, portable, and no DBA required.


Prerequisites
Windows PowerShell 5.1 (or PowerShell 7 started with -STA for WinForms)

PWPS_DAB module (ProjectWise cmdlets)

ImportExcel module for Export-Excel

Read-only access to the ProjectWise DB via Select-PWSQL (through PWPS_DAB)


Step-by-step (with “why” sprinkled in)
Pick your datasource - TreeView lists your PW datasources. No typing. Just click.

Browse rich projects - Expand a datasource to lazy-load its rich projects. Each project gets a button.

Click to run SQL - The button runs a grouped SQL query that pulls all versions under the selected project.

Group + display - Documents are grouped by original GUID. Latest version shown in the grid. Older versions shown inline via RowDetailsTemplate.

Export to Excel - One click. Instant register. Timestamped filename. Worksheet named after the project.

Minimize image
Edit image
Delete image

Add a caption (optional)

The code


Maximize image
Edit image
Delete image

Add a caption (optional)

Add-Type -AssemblyName PresentationFramework

# Load XAML
$xaml = Get-Content "path-to-xaml.xaml" -Raw
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find the Image control
$footerImage = $window.FindName("FooterImage")

# Set the PNG image source
$imagePath = "path-to-footer logo.png"  # Replace with your actual path
$uri = New-Object System.Uri($imagePath, [System.UriKind]::Absolute)
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.UriSource = $uri
$bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
$bitmap.EndInit()

$footerImage.Source = $bitmap
$title = "{REST:wise} Scripting - Example UI for Doc Register"
$window.Title = $title
$appName = $window.FindName("AppName")

$appName.FontFamily = "Arial"
$appName.Text = $title
$appName.FontSize = "32"
$appName.Foreground = "#4ec9b0"

$leftPanel = $window.FindName("LeftPanelContent")

# Create TreeView
$treeView = New-Object System.Windows.Controls.TreeView
$treeView.Name = "ProjectTree"
$treeView.Margin = "10"

$mainGrid = $window.FindName("MainPanelContent")


# Find the style resources
$mainStyle = $window.FindResource("MainDataGridStyle")
$childStyle = $window.FindResource("ChildDataGridStyle")
$dsTemplate = $window.FindResource("DatabaseTemplate")
$projTemplate = $window.FindResource("ProjectTemplate")


$buttonRibbon = $window.FindName("ButtonRibbon")

# Create Export Button
$exportButton = New-Object System.Windows.Controls.Button
$exportButton.Content = "Export to Excel"
$exportButton.Margin = "5"
$exportButton.ToolTip = "Export the register to Excel"

# Add Click event handler
$exportButton.Add_Click({
        try {
            if ($script:data -and $script:data.Rows.Count -gt 0) {
                $script:data | Export-Excel -AutoSize -WorksheetName "Register" -Show
            }
            else {
                [System.Windows.MessageBox]::Show("No data available to export.", "Export Failed", "OK", "Error")
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

# Add the button to the ribbon
$buttonRibbon.Children.Add($exportButton)

$datasource = Get-PWDSConfigEntry

$projectName = ""

$data = @()
$grpData = @()

$datasource.ForEach({
        $item = $_
        $treeItem = New-Object System.Windows.Controls.TreeViewItem
        # Create StackPanel for header
        $headerPanel = New-Object System.Windows.Controls.StackPanel
        $headerPanel.Orientation = "Horizontal"
    
        # Add icon
        $icon = New-Object System.Windows.Controls.Image
        $icon.Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]::new("Database-icon.png"))
        $icon.Width = 16
        $icon.Height = 16
        $icon.Margin = "0,0,5,0"
        $headerPanel.Children.Add($icon)

        # Add label
        $label = New-Object System.Windows.Controls.Label
        $label.Content = $item.Name
        $label.Foreground = $window.Resources["PrimaryBrush"]
        $headerPanel.Children.Add($label)
    
        # Assign header
        $treeItem.Header = $headerPanel
        $treeItem.Tag = $item
    
        # Add dummy child to show expand arrow
        $treeItem.Items.Add("Loading...")

        # Attach lazy loading handler
        $treeItem.Add_Expanded({
                param($sender, $args)

                if ($sender.Items.Count -eq 1 -and $sender.Items[0] -eq "Loading...") {
                    $sender.Items.Clear()

                    $fullpath = $sender.Tag.Name
                    $token = Get-PWConnectionClientToken -UsePWRelyingParty
                    $successLogin = New-PWLogin -DatasourceName $fullpath -BentleyIMS -Token $token

                    if ($successLogin.ToString() -eq "True") {
                        $rps = Get-PWRichProjects -PopulateProjectProperties

                        $rps.ForEach({
                                $project = $_
                                $projectID = $project.ProjectID
                                $projectName = $project.Name
                    
                                # Create Grid layout
                                $grid = New-Object System.Windows.Controls.Grid
                                $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
                                $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))

                                # Create inner StackPanel for icon + label
                                $leftPanel = New-Object System.Windows.Controls.StackPanel
                                $leftPanel.Orientation = "Horizontal"

                                # Icon
                                $icon = New-Object System.Windows.Controls.Image
                                $icon.Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]::new("Project-icon.png"))
                                $icon.Width = 16
                                $icon.Height = 16
                                $icon.Margin = "0,0,5,0"
                                $leftPanel.Children.Add($icon)

                                # Label
                                $label = New-Object System.Windows.Controls.TextBlock
                                $label.Text = $project.Name
                                $label.VerticalAlignment = "Center"
                                $label.Margin = "0,0,10,0"
                                $leftPanel.Children.Add($label)

                                # Place leftPanel in column 0
                                [System.Windows.Controls.Grid]::SetColumn($leftPanel, 0)
                                $grid.Children.Add($leftPanel)

                                # Button
                                $button = New-Object System.Windows.Controls.Button
                                $button.Content = [char]0x25B6
                                $button.Tag = $projectID
                                $button.Width = 30
                                $button.Height = 20
                                $button.ToolTip = "Run SQL for $($project.Name)"
                                $button.HorizontalAlignment = "Right"

                                # Place button in column 1
                                [System.Windows.Controls.Grid]::SetColumn($button, 1)
                                $grid.Children.Add($button)
                    
                                $button.Add_Click({
                                        param($s, $e)
                                        try {
                                            $sqlStatementGetAllVersions = "
    SELECT 
        D.o_docguid AS DocGuid, 
        D.o_itemname AS Document_Name, 
        D.o_dmsdate AS CheckedOut_Date, 
        D.o_version AS Version, 
        S.o_statename AS State,
        D.o_origguid AS OrigGUID,
        D.o_version_seq AS VersionSeq,
        ROW_NUMBER() OVER (PARTITION BY D.o_origguid ORDER BY D.o_version_seq DESC) AS rn
    FROM 
        dms_doc D 
    JOIN 
        dms_stat S ON D.o_stateno = S.o_stateno 
    JOIN 
        (SELECT o_projectno FROM dbo.dsqlGetSubFolders (1, $($s.Tag), 0)) AS SubProjects 
        ON D.o_projectno = SubProjects.o_projectno 
    WHERE 
        D.o_size != 0
"

                                            $SQLResults = Select-PWSQL -SQLSelectStatement $sqlStatementGetAllVersions
                        
                                            $groupedData = @()

                                            # Convert DataTable rows to array for easier processing
                                            $rows = @()
                                            foreach ($row in $SQLResults.Rows) {
                                                $rows += $row
                                            }

                                            # Find parent rows 
                                            $parents = ($rows | ConvertTo-DataTable).Select( "OrigGUID = ''")

                                            foreach ($parent in $parents) {
                                                $parentGuid = $parent.DocGuid

                                                # Find child rows where o_origguid matches parent's o_docguid
                                                $children = $rows | Where-Object { $_.OrigGUID -eq $parentGuid }

                                                # Combine parent and children, then sort by VersionSeq descending
                                                $allVersions = @($parent) + $children
                                                $sortedVersions = $allVersions | Sort-Object { $_.VersionSeq } -Descending
    
                                                # Select only the desired columns
                                                $cleanedVersions = @($sortedVersions | ForEach-Object {
                                                        [PSCustomObject]@{
                                                            Document_Name   = $_.Document_Name
                                                            Version         = $_.Version
                                                            State           = $_.State
                                                            CheckedOut_Date = $_.CheckedOut_Date
                                                            DocGuid         = $_.DocGuid
                                                        }
                                                    })

                                                if ($cleanedVersions.Count -gt 0) {
                                                    $groupedData += [PSCustomObject]@{
                                                        Versions         = $cleanedVersions
                                                        DocumentName     = $cleanedVersions[0].Document_Name
                                                        VersionCount     = $cleanedVersions.Count
                                                        CurrentVersion   = $cleanedVersions[0].Version
                                                        CurrentState     = $cleanedVersions[0].State
                                                        LastCheckOutDate = $cleanedVersions[0].CheckedOut_Date
                                                    }
                                                }

                                            }

                                            foreach ($group in $groupedData) {
                                                if ($group.Versions.Count -gt 1) {
                                                    $group | Add-Member -MemberType NoteProperty -Name FilteredVersions -Value $group.Versions[1..($group.Versions.Count - 1)]
                                                }
                                                else {
                                                    $group | Add-Member -MemberType NoteProperty -Name FilteredVersions -Value @()
                                                }
                                            }

                                            $script:grpData = $groupedData

                                            # Create DataGrid
                                            $dataGrid = New-Object System.Windows.Controls.DataGrid
                                            $dataGrid.Style = $mainStyle

                                            $dataGrid.AutoGenerateColumns = $false
                                            $dataGrid.Margin = "10"
                                            $dataGrid.ItemsSource = $groupedData
                                            $dataGrid.RowDetailsVisibilityMode = "VisibleWhenSelected"
                                            $dataGrid.IsReadOnly = $true

                                            # Optional: Add summary column for latest version name
                                            $colLatestName = New-Object System.Windows.Controls.DataGridTextColumn
                                            $colLatestName.Header = "Document Name"
                                            $colLatestName.Binding = New-Object System.Windows.Data.Binding("Versions[0].Document_Name")
                                            $dataGrid.Columns.Add($colLatestName)

                                            # Optional: Add summary column for version count
                                            $colCount = New-Object System.Windows.Controls.DataGridTextColumn
                                            $colCount.Header = "Version Count"
                                            $colCount.Binding = New-Object System.Windows.Data.Binding("Versions.Count")
                                            $dataGrid.Columns.Add($colCount)

                                            # Optional: Add summary column for version count
                                            $colVer = New-Object System.Windows.Controls.DataGridTextColumn
                                            $colVer.Header = "Current Version"
                                            $colVer.Binding = New-Object System.Windows.Data.Binding("Versions[0].Version")
                                            $dataGrid.Columns.Add($colVer)

                                            # Optional: Add summary column for version count
                                            $colState = New-Object System.Windows.Controls.DataGridTextColumn
                                            $colState.Header = "Current State"
                                            $colState.Binding = New-Object System.Windows.Data.Binding("Versions[0].State")
                                            $dataGrid.Columns.Add($colState)

                                            # Optional: Add summary column for version count
                                            $colDate = New-Object System.Windows.Controls.DataGridTextColumn
                                            $colDate.Header = "Last Check Out Date"
                                            $binding = New-Object System.Windows.Data.Binding("Versions[0].CheckedOut_Date")
                                            $binding.StringFormat = "dd/MM/yyyy HH:mm:ss"  # Or "yyyy-MM-dd", etc.
                                            $colDate.Binding = $binding
                                            $dataGrid.Columns.Add($colDate)

                                            # Create RowDetailsTemplate using FrameworkElementFactory
                                            $rowDetailsTemplate = New-Object System.Windows.DataTemplate

                                            # ItemsControl to hold the list of versions
                                            $itemsControlFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.ItemsControl])
                                            $itemsControlFactory.SetBinding(
                                                [System.Windows.Controls.ItemsControl]::ItemsSourceProperty,
                                                (New-Object System.Windows.Data.Binding("FilteredVersions"))
                                            )

                                            $xamlTemplate = @"
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <DataGrid ItemsSource="{Binding FilteredVersions}"
              AutoGenerateColumns="False"
              IsReadOnly="True"
              HeadersVisibility="Column"
              Background="#F8FAEB"
              Foreground="#333333"
              BorderBrush="#B3CB37"
              FontSize="13"
              RowBackground="White"
              AlternatingRowBackground="#F0F0F0"
              GridLinesVisibility="Horizontal">
        <DataGrid.Columns>
            <DataGridTextColumn Header="Document Name" Binding="{Binding Document_Name}" />
            <DataGridTextColumn Header="Version" Binding="{Binding Version}" />
            <DataGridTextColumn Header="State" Binding="{Binding State}" />
            <DataGridTextColumn Header="CheckedOut Date" Binding="{Binding CheckedOut_Date}" />
        </DataGrid.Columns>
    </DataGrid>
</DataTemplate>
"@

                                            $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlTemplate)
                                            $rowDetailsTemplate = [Windows.Markup.XamlReader]::Load($reader)

                                            $dataGrid.RowDetailsTemplate = $rowDetailsTemplate

                                            $mainGrid.Children.Clear()
                                            $mainGrid.Children.Add($dataGrid)
                        
                                            $rows = @()

                                            foreach ($group in $groupedData) {
                                                # Add a header row for the group
                                                $rows += [PSCustomObject]@{
                                                    DocGUID         = $group.Versions[0].DocGuid
                                                    Document_Name   = $group.DocumentName
                                                    Version         = $group.CurrentVersion
                                                    State           = $group.CurrentState
                                                    CheckedOut_Date = $group.LastCheckOutDate
                                                }

                                                # Add each version row
                                                # Skip the first version (index 0)
                                                $versions = $group.Versions
                                                for ($i = 1; $i -lt $versions.Count; $i++) {
                                                    $version = $versions[$i]

                                                    $rows += [PSCustomObject]@{
                                                        VersionGuid     = $version.DocGuid
                                                        Document_Name   = $version.Document_Name
                                                        Version         = $version.Version
                                                        State           = $version.State
                                                        CheckedOut_Date = $version.CheckedOut_Date
                                                    }
                                                }
                                            }

                                            $script:data = $rows
                                        }
                                        catch {
                                            Write-Host "An error occurred: $($_.Exception.Message)"
                                            $errorText = New-Object System.Windows.Controls.TextBlock
                                            $errorText.Text = "Failed to retrieve Register"
                                            $errorText.VerticalAlignment = "Center"
                                            $errorText.HorizontalAlignment = "Center"
                            
                                            $mainGrid.Children.Clear()
                                            $mainGrid.Children.Add($errorText)
                                        }
                        
                                    })
                    
                                $projectItem = New-Object System.Windows.Controls.TreeViewItem
                                $projectItem.Header = $grid
                                #$projectItem.HeaderTemplate = $projTemplate
                                $sender.Items.Add($projectItem)
                            })
                    }
                    else {
                        $errorItem = New-Object System.Windows.Controls.TreeViewItem
                        $errorItem.Header = "Login failed"
                        $sender.Items.Add($errorItem)
                    }
                }
            })



        $treeView.Items.Add($treeItem)
    })

# Add TreeView to LeftPanel
$leftPanel.Children.Add($treeView)

# Show the window
$window.ShowDialog()

The UI (WPF XAML)
WPF gives us:

A proper layout: header, ribbon, panels

Custom styles and brushes

TreeView with icons and buttons

DataGrid with expandable rows

<!-- x:Class="PowerShellWPF.MainWindow" 
    xmlns:local="clr-namespace:PowerShellWPF
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="d""
    -->
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        
        x:Name="MainApp"
        Title="App Name" Height="800" Width="1600" FontFamily="Yaro St">
    <Window.Resources>
        <SolidColorBrush x:Key="PrimaryBrush" Color="#192d38"/>
        <SolidColorBrush x:Key="SecondaryBrush" Color="#BEDAE5"/>
        <SolidColorBrush x:Key="Green" Color="#05E560"/>
        <SolidColorBrush x:Key="Blue" Color="#3F32F1"/>
        <SolidColorBrush x:Key="Yellow" Color="#B9FF00"/>
        <SolidColorBrush x:Key="Purple" Color="#BE02F8"/>
        <SolidColorBrush x:Key="White" Color="#ffffff"/>
        <SolidColorBrush x:Key="Black" Color="#000000"/>
        <FontFamily x:Key="HeaderFont">Gamechanger</FontFamily>

            <!-- Icons from root of solution folder -->
            <BitmapImage x:Key="DatabaseIcon" UriSource="Database.png"/>
            <BitmapImage x:Key="ProjectIcon" UriSource="Project.png"/>

            <!-- TreeView Style -->
            <Style TargetType="TreeView">
                <Setter Property="Background" Value="{StaticResource SecondaryBrush}"/>
                <Setter Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                <Setter Property="BorderThickness" Value="1.5"/>
                <Setter Property="FontSize" Value="14"/>
                <Setter Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
            </Style>

        <!-- Template for Database items -->
        <DataTemplate x:Key="DatabaseTemplate">
            <StackPanel Orientation="Horizontal">
                <Image Source="{StaticResource DatabaseIcon}" Width="16" Height="16" Margin="0,0,5,0"/>
                <TextBlock Text="{Binding}" Foreground="{StaticResource PrimaryBrush}" FontWeight="Bold"/>
            </StackPanel>
        </DataTemplate>

        <!-- Template for Project items -->
        <DataTemplate x:Key="ProjectTemplate">
            <StackPanel Orientation="Horizontal">
                <Image Source="{StaticResource ProjectIcon}" Width="16" Height="16" Margin="0,0,5,0"/>
                <TextBlock Text="{Binding}" Foreground="{StaticResource PrimaryBrush}"/>
            </StackPanel>
        </DataTemplate>

        <Style x:Key="MainDataGridStyle" TargetType="DataGrid">
            <Setter Property="Background" Value="{StaticResource SecondaryBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F0F0F0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
        </Style>

        <Style x:Key="ChildDataGridStyle" TargetType="DataGrid">
            <Setter Property="Background" Value="#F8FAEB"/>
            <Setter Property="Foreground" Value="#333"/>
            <Setter Property="BorderBrush" Value="#B3CB37"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F0F0F0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
        </Style>

        <Style TargetType="{x:Type Button}">
            <Setter Property="Foreground" Value="{StaticResource SecondaryBrush}" />
            <Setter Property="Background" Value="{StaticResource PrimaryBrush}" />
            <Setter Property="BorderBrush" Value="{StaticResource PrimaryBrush}" />
            <Setter Property="BorderThickness" Value="1.5" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border Background="{StaticResource PrimaryBrush}"
                                BorderBrush="{StaticResource PrimaryBrush}"
                                BorderThickness="1.5"
                                CornerRadius="5"
                                >
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                              VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                                              TextElement.Foreground="{TemplateBinding Foreground}" />
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource Green}" />
                    <Setter Property="Foreground" Value="{StaticResource White}" />
                    <Setter Property="BorderBrush" Value="{StaticResource White}" />
                    <Setter Property="BorderThickness" Value="2" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="ComboBoxToggleButton" TargetType="{x:Type ToggleButton}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ToggleButton}">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition />
                                <ColumnDefinition Width="32" />
                            </Grid.ColumnDefinitions>
                            <Border
                      x:Name="Border"
                      Grid.ColumnSpan="2"
                      CornerRadius="8"
                      Background="{StaticResource SecondaryBrush}"
                      BorderBrush="{StaticResource PrimaryBrush}"
                      BorderThickness="1.5" 
                    />

                            <Path
                        x:Name="Arrow"
                        Grid.Column="1"    
                        Fill="{TemplateBinding Foreground}"
                        Stroke="{TemplateBinding Foreground}"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Center"
                        Data="M 0 0 L 4 4 L 8 0 Z"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="{x:Type TextBox}">
            <Border x:Name="PART_ContentHost" Focusable="True" />
        </ControlTemplate>

        <Style x:Key="theComboBox" TargetType="{x:Type ComboBox}">
            <Setter Property="Foreground" Value="#333" />
            <Setter Property="BorderBrush" Value="Gray" />
            <Setter Property="Background" Value="White" />
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="FontSize" Value="13" />
            <Setter Property="MinWidth" Value="150"/>
            <Setter Property="MinHeight" Value="35"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBox}">
                        <Grid>
                            <ToggleButton
                        Cursor="Hand"
                        x:Name="ToggleButton"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        Background="{TemplateBinding Background}"
                        Foreground="{TemplateBinding Foreground}"
                        Style="{StaticResource ComboBoxToggleButton}"
                        Grid.Column="2"
                        Focusable="false"
                        IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource Mode=TemplatedParent}}"
                        ClickMode="Press"/>

                            <ContentPresenter
                        x:Name="ContentSite"
                        IsHitTestVisible="False"
                        Content="{TemplateBinding SelectionBoxItem}"
                        ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                        ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                        Margin="10,3,30,3"
                        VerticalAlignment="Center"
                        HorizontalAlignment="Left" />
                            <TextBox x:Name="PART_EditableTextBox"
                        Style="{x:Null}"
                        Template="{StaticResource ComboBoxTextBox}"
                        HorizontalAlignment="Left"
                        VerticalAlignment="Center"
                        Margin="3,3,23,3"
                        Focusable="True"                               
                        Visibility="Hidden"
                        IsReadOnly="{TemplateBinding IsReadOnly}"/>
                            <Popup
                        x:Name="Popup"
                        Placement="Bottom"
                        IsOpen="{TemplateBinding IsDropDownOpen}"
                        AllowsTransparency="True"
                        Focusable="False"
                        PopupAnimation="Slide">
                                <Grid
                          x:Name="DropDown"
                          SnapsToDevicePixels="True"               
                          MinWidth="{TemplateBinding ActualWidth}"
                          MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border
                                CornerRadius="8"
                                x:Name="DropDownBorder"
                                Background="White"
                                BorderThickness="1"
                                BorderBrush="#F6F6F6"
                                />
                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Grid>
                            </Popup>

                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="true">
                                <Setter Property="IsTabStop" Value="false"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="theComboBoxItem" TargetType="{x:Type ComboBoxItem}">
            <Setter Property="SnapsToDevicePixels" Value="true" />
            <Setter Property="HorizontalAlignment" Value="Stretch" />
            <Setter Property="VerticalAlignment" Value="Stretch" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBoxItem}">
                        <Border
                    x:Name="Border"
                    Padding="5"
                    Margin="2"
                    BorderThickness="2,0,0,0"
                    CornerRadius="0"
                    Background="Transparent"
                    BorderBrush="Transparent">
                            <TextBlock TextAlignment="Left"><InlineUIContainer>
                        <ContentPresenter />
                            </InlineUIContainer></TextBlock>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#B3CB37"/>
                                <Setter TargetName="Border" Property="Background" Value="#F8FAEB"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition x:Name="Header" Height="60"/>
            <RowDefinition x:Name="Ribbon" Height="60" MinHeight="10" MaxHeight="60"/>
            <RowDefinition x:Name="RowSplitter" Height="Auto"/>
            <RowDefinition x:Name="MainPanelRow" Height="*"/>
            <RowDefinition x:Name="Footer" Height="60"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition x:Name="LeftColumnPanel" Width="Auto" MinWidth="100" MaxWidth="400"/>
            <ColumnDefinition x:Name="LeftSplitter" Width="Auto"/>
            <ColumnDefinition x:Name="MainPanelColumn"  Width="*" MinWidth="300" MaxWidth="1400"/>
            <ColumnDefinition x:Name="RightSplitter" Width="Auto"/>
            <ColumnDefinition x:Name="RightColumnPanel" Width="Auto" MinWidth="100" MaxWidth="400"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Row="0" Grid.ColumnSpan="5" Background="{StaticResource PrimaryBrush}">
            <TextBlock x:Name="AppName" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,10,10,10" Text="App Name" FontSize="40" Foreground="{StaticResource SecondaryBrush}" FontFamily="{StaticResource HeaderFont}" FontWeight="Bold"/>
        </Border>

        <StackPanel x:Name="ButtonRibbon" Grid.Row="1" Grid.ColumnSpan="5" Background="{StaticResource SecondaryBrush}" Orientation="Horizontal">
            
        </StackPanel>

        <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" VerticalAlignment="Center" Height="5" Background="{StaticResource PrimaryBrush}" Grid.ColumnSpan="5"/>

        <Border x:Name="LeftPanel" Grid.Row="3" Grid.Column="0" Background="{StaticResource SecondaryBrush}">
            <Grid x:Name="LeftPanelContent"/>
        </Border>

        <GridSplitter Grid.Row="3" Grid.Column="1" Width="5" Background="{StaticResource PrimaryBrush}" HorizontalAlignment="Center" />

        <Border x:Name="MainPanel" Grid.Row="3" Grid.Column="2" Background="{StaticResource SecondaryBrush}">
            <Grid x:Name="MainPanelContent"/>
        </Border>

        <GridSplitter Grid.Row="3" Grid.Column="3" Width="5" Background="{StaticResource PrimaryBrush}" HorizontalAlignment="Center"/>

        <Border x:Name="RightPanel" Grid.Row="3" Grid.Column="4" Background="{StaticResource SecondaryBrush}">

        </Border>

        <Border Grid.Row="4" Grid.ColumnSpan="5" Background="{StaticResource PrimaryBrush}">
            <Image x:Name="FooterImage" VerticalAlignment="Center" HorizontalAlignment="Right" Margin="10"/>
        </Border>
    </Grid>
</Window>
WSG vs SQL — the “how it feels” comparison
WSG (REST) Predictable, supported, works everywhere WSG works. You’ll do paging, decode properties, and loop results. Perfect for service accounts and cloud-friendly scripts.

SQL (read-only) Fast and flexible—especially for register reporting. You can join to anything. But you need DB access, and schema changes can break queries.


Side-by-side snippet
WSG

$top = 100; $skip = 0; $all = @()
do {
  $url = "$base/repositories/PW_WSG/Document!poly?`$top=$top&`$skip=$skip&$filter=FolderId eq '$FolderId'"
  $page = Invoke-WSG -Method GET -Uri $url
  $all += $page.instances
  $skip += $top
} while ($page.instances.Count -eq $top)

$all | Select-Object Name, Version, State | Export-Excel C:\Temp\WSG-DocReg.xlsx

SQL

$SQLResults = Select-PWSQL -SQLSelectStatement $sqlLatestPerOrigGuid
$SQLResults | Export-Excel C:\Temp\SQL-DocReg.xlsx


Lessons learned
WPF : Flexible layout, better styling, and richer controls.

Manual header composition: TreeViewItem headers built with Grid for icon + label + button.

SQL grouping: Using ROW_NUMBER() and OrigGUID to group versions.

RowDetailsTemplate: Shows version history inline.

ExportExcel: Still the MVP for quick exports.


Try this next
Add a state filter dropdown that updates the WHERE clause.

Add file type filters (.dgn, .dwg, .pdf).

Include environment attribute joins in your register.

Show a progress label + quick preview grid for large projects.