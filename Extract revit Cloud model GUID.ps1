# BIM360Scanner.ps1

# Load required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Constants
$script:CLIENT_ID = "CLIENT_ID"
$script:CLIENT_SECRET = "CLIENT_SECRET"
$script:ACCESS_TOKEN = ""
$script:EXPORT_PATH = Join-Path ([Environment]::GetFolderPath("Desktop")) "BIM360_Models.csv"

# XAML for main window
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="BIM360 Scanner" Height="700" Width="1000"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="BIM360 Project:" FontWeight="Bold" Margin="0,0,0,5"/>
            <ComboBox x:Name="ProjectComboBox" Height="30"/>
        </StackPanel>

        <TextBlock Grid.Row="1" Text="Select Folders to Scan:" FontWeight="Bold" Margin="0,10,0,5"/>
        
        <TreeView x:Name="FolderTreeView" Grid.Row="2" Margin="0,0,0,10"/>
        
        <ProgressBar x:Name="ScanProgress" Grid.Row="3" Height="20" Margin="0,0,0,10"/>
        
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <TextBlock x:Name="StatusText" Grid.Column="0" VerticalAlignment="Center"/>
            <Button x:Name="ScanButton" Grid.Column="1" Content="Start Scan" Width="100" Height="30" Margin="0,0,10,0"/>
            <Button x:Name="CancelButton" Grid.Column="2" Content="Close" Width="100" Height="30"/>
        </Grid>
    </Grid>
</Window>
"@

# Create Window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$projectComboBox = $window.FindName("ProjectComboBox")
$folderTreeView = $window.FindName("FolderTreeView")
$scanProgress = $window.FindName("ScanProgress")
$statusText = $window.FindName("StatusText")
$scanButton = $window.FindName("ScanButton")
$cancelButton = $window.FindName("CancelButton")

# Authentication function
function Get-ForgeAccessToken {
    try {
        # Create form data
        $form = [System.Collections.Specialized.NameValueCollection]::new()
        $form.Add("grant_type", "client_credentials")
        $form.Add("scope", "data:read account:read")
        
        # Create WebClient for the request
        $client = New-Object System.Net.WebClient
        $auth = "$($script:CLIENT_ID):$($script:CLIENT_SECRET)"
        $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($auth))
        $client.Headers.Add("Authorization", "Basic $encodedAuth")
        $client.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
        
        # Make the request
        $response = $client.UploadValues(
            "https://developer.api.autodesk.com/authentication/v2/token", 
            "POST", 
            $form
        )
        $responseString = [System.Text.Encoding]::UTF8.GetString($response)
        $responseJson = $responseString | ConvertFrom-Json
        
        if ($responseJson.access_token) {
            $script:ACCESS_TOKEN = $responseJson.access_token
            return $true
        }
        
        return $false
    }
    catch {
        return $false
    }
}

# Get BIM360 Projects
function Get-BIM360Projects {
    try {
        $headers = @{
            "Authorization" = "Bearer $($script:ACCESS_TOKEN)"
        }
        
        # First get hubs
        $hubsResponse = Invoke-RestMethod `
            -Uri "https://developer.api.autodesk.com/project/v1/hubs" `
            -Method Get `
            -Headers $headers
            
        # Find BIM 360 hub
        $bimHub = $hubsResponse.data | Where-Object { 
            $_.attributes.extension.type -eq "hubs:autodesk.bim360:Account"
        }
        
        if ($bimHub) {
            # Get projects for this hub
            $projectsResponse = Invoke-RestMethod `
                -Uri "https://developer.api.autodesk.com/project/v1/hubs/$($bimHub.id)/projects" `
                -Method Get `
                -Headers $headers
                
            return $projectsResponse.data | ForEach-Object {
                [PSCustomObject]@{
                    HubId = $bimHub.id
                    ProjectId = $_.id
                    ProjectName = $_.attributes.name
                }
            }
        }
        
        return $null
    }
    catch {
        return $null
    }
}

function Is-HiddenFolder {
    param($FolderName)
    
    # Simple pattern to match GUID-like strings and system folders
    if ($FolderName -match '[0-9a-f]{8}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{12}') {
        return $true
    }
    
    return $false
}

# Modified Load-Folders function
function Load-Folders {
    param($Project)
    try {
        $headers = @{
            "Authorization" = "Bearer $($script:ACCESS_TOKEN)"
        }
        
        $response = Invoke-RestMethod `
            -Uri "https://developer.api.autodesk.com/project/v1/hubs/$($Project.HubId)/projects/$($Project.ProjectId)/topFolders" `
            -Method Get `
            -Headers $headers
            
        $folderTreeView.Dispatcher.Invoke([Action]{
            $folderTreeView.Items.Clear()
            
            foreach ($folder in $response.data) {
                $folderName = $folder.attributes.name
                
                # Skip hidden folders
                if (Is-HiddenFolder -FolderName $folderName) {
                    continue
                }
                
                $item = New-Object System.Windows.Controls.TreeViewItem
                
                # Create checkbox
                $checkBox = New-Object System.Windows.Controls.CheckBox
                $checkBox.Content = $folderName
                $checkBox.Margin = New-Object System.Windows.Thickness(2)
                $item.Header = $checkBox
                
                $item.Tag = @{
                    Id = $folder.id
                    Path = $folderName
                    IsExpanded = $false
                }
                
                # Add dummy item to enable expansion
                $dummy = New-Object System.Windows.Controls.TreeViewItem
                $dummy.Header = "Loading..."
                $item.Items.Add($dummy)
                
                $folderTreeView.Items.Add($item)
            }
        })
    }
    catch {
    }
}

# Modified Load-SubFolders function
function Load-SubFolders {
    param($ParentItem)
    
    try {
        $project = $projectComboBox.SelectedItem
        $parentInfo = $ParentItem.Tag
        
        $headers = @{
            "Authorization" = "Bearer $($script:ACCESS_TOKEN)"
        }
        
        $response = Invoke-RestMethod `
            -Uri "https://developer.api.autodesk.com/data/v1/projects/$($project.ProjectId)/folders/$($parentInfo.Id)/contents" `
            -Method Get `
            -Headers $headers
            
        $ParentItem.Dispatcher.Invoke([Action]{
            $ParentItem.Items.Clear()
        
            foreach ($item in $response.data) {
                if ($item.type -eq "folders") {
                    $folderName = $item.attributes.displayName
                    
                    # Skip hidden folders
                    if (Is-HiddenFolder -FolderName $folderName) {
                        continue
                    }
                    
                    $treeItem = New-Object System.Windows.Controls.TreeViewItem
                    
                    # Create checkbox
                    $checkBox = New-Object System.Windows.Controls.CheckBox
                    $checkBox.Content = $folderName
                    $checkBox.Margin = New-Object System.Windows.Thickness(2)
                    $treeItem.Header = $checkBox
                    
                    $newPath = "$($parentInfo.Path)\$folderName"
                    
                    $treeItem.Tag = @{
                        Id = $item.id
                        Path = $newPath
                        IsExpanded = $false
                    }
                    
                    # Add dummy item for expansion
                    $dummy = New-Object System.Windows.Controls.TreeViewItem
                    $dummy.Header = "Loading..."
                    $treeItem.Items.Add($dummy)
                    
                    $ParentItem.Items.Add($treeItem)
                }
            }
        })
    }
    catch {
    }
}

# Get selected folders
function Get-SelectedFolders {
    $selected = New-Object System.Collections.Generic.List[PSObject]
    
    function Process-TreeItem {
        param($Item)
        
        $checkbox = $Item.Header
        if ($checkbox -is [System.Windows.Controls.CheckBox] -and $checkbox.IsChecked) {
            $selected.Add($Item.Tag)
        }
        
        foreach ($childItem in $Item.Items) {
            if ($childItem.Header -ne "Loading...") {
                Process-TreeItem $childItem
            }
        }
    }
    
    foreach ($item in $folderTreeView.Items) {
        Process-TreeItem $item
    }
    
    return $selected
}

# Initialize CSV file
function Initialize-CsvFile {
    try {
        if (Test-Path $script:EXPORT_PATH) {
            Remove-Item $script:EXPORT_PATH -Force
        }
        
        "Model GUID,Project Name,Project GUID,Folder Path,Revit Version,Source File Name" | 
            Out-File -FilePath $script:EXPORT_PATH -Encoding UTF8
    }
    catch {
        throw
    }
}

# Add model to CSV
function Add-ModelToCsv {
    param($Model)
    
    try {
        "$($Model.ModelGuid),$($Model.ProjectName),$($Model.ProjectGuid),$($Model.FolderPath),$($Model.RevitVersion),$($Model.SourceFileName)" |
            Add-Content -Path $script:EXPORT_PATH
    }
    catch {
    }
}

# Get folder contents
function Get-FolderContents {
    param($ProjectId, $FolderId)
    
    try {
        $headers = @{
            "Authorization" = "Bearer $($script:ACCESS_TOKEN)"
        }
        
        $response = Invoke-RestMethod `
            -Uri "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/folders/$FolderId/contents" `
            -Method Get `
            -Headers $headers
            
        return $response.data
    }
    catch {
        return $null
    }
}

# Get item details
function Get-ItemDetails {
    param($ProjectId, $ItemId)
    
    try {
        $headers = @{
            "Authorization" = "Bearer $($script:ACCESS_TOKEN)"
        }
        
        $response = Invoke-RestMethod `
            -Uri "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/items/$ItemId" `
            -Method Get `
            -Headers $headers
            
        return $response
    }
    catch {
        return $null
    }
}

# Start the scan process
# Update the Start-BIM360Scan function to include better progress tracking
# Update the Start-BIM360Scan function for real-time GUI updates
# Add these at the beginning of the script
$script:MAX_JOBS = 1  # Maximum number of parallel jobs
$script:BATCH_SIZE = 1 # Number of items to process in each batch

# Add this function for parallel processing
function Process-FolderBatch {
    param (
        [string]$ProjectId,
        [string]$FolderId,
        [string]$FolderPath,
        [string]$ProjectName,
        [string]$AccessToken
    )
    
    try {
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
        }
        
        # Get all contents in one API call
        $contents = Invoke-RestMethod `
            -Uri "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/folders/$FolderId/contents" `
            -Method Get `
            -Headers $headers
            
        $results = @()
        
        $revitFiles = $contents.data | Where-Object { 
            $_.type -eq "items" -and $_.attributes.displayName -like "*.rvt" 
        }
        
        foreach ($item in $revitFiles) {
            try {
                $itemDetails = Invoke-RestMethod `
                    -Uri "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/items/$($item.id)" `
                    -Method Get `
                    -Headers $headers
                
                $versionData = $itemDetails.included | 
                    Where-Object { $_.type -eq "versions" } |
                    Select-Object -First 1 -ExpandProperty attributes |
                    Select-Object -ExpandProperty extension |
                    Select-Object -ExpandProperty data
                
                if ($versionData) {
                    $results += [PSCustomObject]@{
                        ModelGuid = $versionData.modelGuid
                        ProjectName = $ProjectName
                        ProjectGuid = $versionData.projectGuid
                        FolderPath = $FolderPath
                        RevitVersion = $versionData.revitProjectVersion
                        SourceFileName = $item.attributes.displayName
                    }
                }
            }
            catch {
            }
        }
        
        return $results
    }
    catch {
        return $null
    }
}

# Modified Start-BIM360Scan function with parallel processing
function Start-BIM360Scan {
    try {
        $project = $projectComboBox.SelectedItem
        $selectedFolders = Get-SelectedFolders
        
        if ($selectedFolders.Count -eq 0) {
            return
        }
        
        # Disable UI during scan
        $scanButton.IsEnabled = $false
        $projectComboBox.IsEnabled = $false
        $folderTreeView.IsEnabled = $false
        
        # Initialize CSV file
        Initialize-CsvFile
        
        # Initialize scan variables
        $processedFolders = 0
        $totalFolders = $selectedFolders.Count
        $totalModels = 0
        
        # Update initial status
        $window.Dispatcher.Invoke([Action]{
            $scanProgress.Value = 0
            $statusText.Text = "Starting scan... Total folders to process: $totalFolders"
        })
        
        # Create scriptblock for the job
        $jobScriptBlock = {
            param($ProjectId, $FolderId, $FolderPath, $ProjectName, $AccessToken)
            
            # Define the function within the scriptblock
            function Process-FolderContents {
                param (
                    [string]$ProjectId,
                    [string]$FolderId,
                    [string]$FolderPath,
                    [string]$ProjectName,
                    [string]$AccessToken
                )
                
                try {
                    $headers = @{
                        "Authorization" = "Bearer $AccessToken"
                    }
                    
                    $contents = Invoke-RestMethod `
                        -Uri "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/folders/$FolderId/contents" `
                        -Method Get `
                        -Headers $headers
                        
                    $results = @()
                    
                    $revitFiles = $contents.data | Where-Object { 
                        $_.type -eq "items" -and $_.attributes.displayName -like "*.rvt" 
                    }
                    
                    foreach ($item in $revitFiles) {
                        try {
                            $itemDetails = Invoke-RestMethod `
                                -Uri "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/items/$($item.id)" `
                                -Method Get `
                                -Headers $headers
                            
                            $versionData = $itemDetails.included | 
                                Where-Object { $_.type -eq "versions" } |
                                Select-Object -First 1 -ExpandProperty attributes |
                                Select-Object -ExpandProperty extension |
                                Select-Object -ExpandProperty data
                            
                            if ($versionData) {
                                $results += [PSCustomObject]@{
                                    ModelGuid = $versionData.modelGuid
                                    ProjectName = $ProjectName
                                    ProjectGuid = $versionData.projectGuid
                                    FolderPath = $FolderPath
                                    RevitVersion = $versionData.revitProjectVersion
                                    SourceFileName = $item.attributes.displayName
                                }
                            }
                        }
                        catch {
                        }
                    }
                    
                    return $results
                }
                catch {
                    return $null
                }
            }
            
            # Call the function with the provided parameters
            Process-FolderContents -ProjectId $ProjectId -FolderId $FolderId -FolderPath $FolderPath -ProjectName $ProjectName -AccessToken $AccessToken
        }
        
        # Process folders in batches
        for ($i = 0; $i -lt $selectedFolders.Count; $i += $script:BATCH_SIZE) {
            $batch = $selectedFolders | Select-Object -Skip $i -First $script:BATCH_SIZE
            
            # Create jobs for parallel processing
            $jobs = @()
            foreach ($folder in $batch) {
                $jobs += Start-Job -ScriptBlock $jobScriptBlock -ArgumentList @(
                    $project.ProjectId,
                    $folder.Id,
                    $folder.Path,
                    $project.ProjectName,
                    $script:ACCESS_TOKEN
                )
            }
            
            # Wait for all jobs and process results
            foreach ($job in $jobs) {
                $results = Receive-Job -Job $job -Wait
                Remove-Job -Job $job
                
                if ($results) {
                    foreach ($model in $results) {
                        Add-ModelToCsv -Model $model
                        $totalModels++
                    }
                }
                
                $processedFolders++
                
                # Update progress
                $window.Dispatcher.Invoke([Action]{
                    $progressPercentage = [math]::Round(($processedFolders / $totalFolders) * 100)
                    $scanProgress.Value = $progressPercentage
                    $statusText.Text = "Processed $processedFolders of $totalFolders folders | Models found: $totalModels"
                })
            }
            
            # Force UI update
            [System.Windows.Forms.Application]::DoEvents()
            
            # Add a small delay to prevent overwhelming the API
            Start-Sleep -Milliseconds 1
        }

        # Update final progress
        $window.Dispatcher.Invoke([Action]{
            $scanProgress.Value = 100
            $statusText.Text = "Scan complete! Processed $totalFolders folders | Total models found: $totalModels"
        })
        
        # Re-enable UI
        $scanButton.IsEnabled = $true
        $projectComboBox.IsEnabled = $true
        $folderTreeView.IsEnabled = $true
    }
    catch {
        # Re-enable UI in case of error
        $scanButton.IsEnabled = $true
        $projectComboBox.IsEnabled = $true
        $folderTreeView.IsEnabled = $true
    }
}

# Add this at the beginning of the script with other Add-Type statements
Add-Type -AssemblyName System.Windows.Forms

# Initialize UI
function Initialize-UI {
    # Setup progress bar
    $scanProgress.Minimum = 0
    $scanProgress.Maximum = 100
    $scanProgress.Value = 0
    
    # Set initial status
    $statusText.Text = "Ready"
    
    # Get and load projects
    if (Get-ForgeAccessToken) {
        $projects = Get-BIM360Projects
        
        if ($projects) {
            $projectComboBox.Dispatcher.Invoke([Action]{
                $projectComboBox.ItemsSource = $projects
                $projectComboBox.DisplayMemberPath = "ProjectName"
            })
        }
    }
    
    # Event handlers
    $projectComboBox.Add_SelectionChanged({
        $project = $projectComboBox.SelectedItem
        if ($project) {
            Load-Folders -Project $project
        }
    })
    
    $scanButton.Add_Click({
        Start-BIM360Scan
    })
    
    $cancelButton.Add_Click({
        $window.Close()
    })
    
    # Add expansion handler for tree items
    $folderTreeView.AddHandler([System.Windows.Controls.TreeViewItem]::ExpandedEvent, [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        $selectedItem = $e.Source
        if ($selectedItem -and $selectedItem.Items.Count -eq 1 -and $selectedItem.Items[0].Header -eq "Loading...") {
            Load-SubFolders -ParentItem $selectedItem
        }
    })
}

# Start application
Initialize-UI
$window.ShowDialog()