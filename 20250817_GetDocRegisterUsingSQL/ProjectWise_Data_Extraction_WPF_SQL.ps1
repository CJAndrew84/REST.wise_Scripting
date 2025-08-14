Add-Type -AssemblyName PresentationFramework

Load XAML
$xaml = Get-Content "path-to-xaml.xaml" -Raw $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml) $window = [Windows.Markup.XamlReader]::Load($reader)

Find the Image control
$footerImage = $window.FindName("FooterImage")

Set the PNG image source
$imagePath = "path-to-footer logo.png" # Replace with your actual path $uri = New-Object System.Uri($imagePath, [System.UriKind]::Absolute) $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage $bitmap.BeginInit() $bitmap.UriSource = $uri $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad $bitmap.EndInit()

$footerImage.Source = $bitmap $title = "{REST:wise} Scripting - Example UI for Doc Register" $window.Title = $title $appName = $window.FindName("AppName")

$appName.FontFamily = "Arial" $appName.Text = $title $appName.FontSize = "32" $appName.Foreground = "#4ec9b0"

$leftPanel = $window.FindName("LeftPanelContent")

Create TreeView
$treeView = New-Object System.Windows.Controls.TreeView $treeView.Name = "ProjectTree" $treeView.Margin = "10"

$mainGrid = $window.FindName("MainPanelContent")

Find the style resources
$mainStyle = $window.FindResource("MainDataGridStyle") $childStyle = $window.FindResource("ChildDataGridStyle") $dsTemplate = $window.FindResource("DatabaseTemplate") $projTemplate = $window.FindResource("ProjectTemplate")

$buttonRibbon = $window.FindName("ButtonRibbon")

Create Export Button
$exportButton = New-Object System.Windows.Controls.Button $exportButton.Content = "Export to Excel" $exportButton.Margin = "5" $exportButton.ToolTip = "Export the register to Excel"

Add Click event handler
$exportButton.Add_Click({ try { if ($script:data -and $script:data.Rows.Count -gt 0) { $script:data | Export-Excel -AutoSize -WorksheetName "Register" -Show } else { [System.Windows.MessageBox]::Show("No data available to export.", "Export Failed", "OK", "Error") } } catch { [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error", "OK", "Error") } })

Add the button to the ribbon
$buttonRibbon.Children.Add($exportButton)

$datasource = Get-PWDSConfigEntry

$projectName = ""

$data = @() $grpData = @()

$datasource.ForEach({ $item = $_ $treeItem = New-Object System.Windows.Controls.TreeViewItem # Create StackPanel for header $headerPanel = New-Object System.Windows.Controls.StackPanel $headerPanel.Orientation = "Horizontal"

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
Add TreeView to LeftPanel
$leftPanel.Children.Add($treeView)

Show the window
$window.ShowDialog()