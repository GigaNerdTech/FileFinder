# File finder
# Find old files on Linux-based systems
# Joshua Woleben
# 10/1/2019

# Load SSH

Import-Module -Name "\\funzone\team\POSH\Powershell\Modules\Posh-SSH.psm1"
Import-Module -Name "\\funzone\team\POSH\Powershell\Modules\Posh-SSH.psd1"

$script:file_hash = @{}

# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="File Finder" Height="1000" Width="800" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="HostLabel" Content="Host to connect to:"/>
        <TextBox x:Name="HostTextBox" Height="20"/>
        <Label x:Name="UserNameLabel" Content="Host Username:"/>
        <TextBox x:Name="UserNameTextBox" Height="20"/>
        <Label x:Name="PasswordNameLabel" Content="Host Password:"/>
        <PasswordBox x:Name="PasswordTextBox" Height="20"/>
        <Label x:Name="StartDateLabel" Content="Start Date"/>
        <DatePicker x:Name="StartDatePicker"/>
        <Label x:Name="EndDateLabel" Content="End Date"/>
        <DatePicker x:Name="EndDatePicker"/>
        <Label x:Name="FilesystemLabel" Content="Filesystem to search:"/>
        <TextBox x:Name="FilesystemTextBox" Height="20"/>
        <Button x:Name="SearchButton" Content="Search for Files" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/> 
        <Button x:Name="ClearFormButton" Content="Clear Form" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Label x:Name="ResultsLabel" Content="Search Results"/>
        <DataGrid x:Name="Results" AutoGenerateColumns="True" Height="400">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Path" Binding="{Binding Path}" Width="400"/>
                <DataGridTextColumn Header="Size" Binding="{Binding Size}" Width="120"/>
                <DataGridTextColumn Header="Date" Binding="{Binding Date}" Width="80"/>
            </DataGrid.Columns>
        </DataGrid>
        <Button x:Name="ExcelButton" Content="Export to Excel" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Button x:Name="ScriptButton" Content="Generate Deletion Script" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
 # <ListBox x:Name="ResultsSelect" Height = "300" SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Visible"/>
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls
$HostTextBox = $global:Form.FindName('HostTextBox')
$UserNameTextBox = $global:Form.FindName('UserNameTextBox')
$PasswordTextBox = $global:Form.FindName('PasswordTextBox')
$StartDatePicker = $global:Form.FindName('StartDatePicker')
$EndDatePicker = $global:Form.FindName('EndDatePicker')
$FilesystemTextBox = $global:Form.FindName('FilesystemTextBox')
$SearchButton = $global:Form.FindName('SearchButton')
$ClearFormButton = $global:Form.FindName('ClearFormButton')
#$ResultsSelect = $global:Form.FindName('ResultsSelect')
$Results = $global:Form.FindName('Results')
$ScriptButton = $global:Form.FindName('ScriptButton')
$ExcelButton = $global:Form.FindName('ExcelButton')

$SearchButton.Add_Click({
    # Get Variables
    $hostname = $HostTextBox.Text
    $username = $UserNameTextBox.Text
    $password = $PasswordTextBox.SecurePassword
    $start_date = $StartDatePicker.SelectedDate
    $end_date = $EndDatePicker.SelectedDate
    $filesystem = $FilesystemTextBox.Text

    # Calculate days from current date
    $start_days = (New-TimeSpan -Start $start_date -End (Get-Date)).Days
    $end_days = (New-TimeSpan -Start $end_date -End (Get-Date)).Days
    Write-Host "Start days: $start_days End days: $end_days"
    # Build find command
    $find_command = "find $filesystem -ctime +$end_days -ctime -$start_days -exec stat -c `"%n,%s,%y`" {} \;"

    # Build credentials
    $creds = New-Object -TypeName System.Management.Automation.PSCredential ($username,$password)

    # Create SSH session
    $ssh_session = New-SSHSession -ComputerName $hostname -Credential $creds

    # Issue find command
    $output = Invoke-SSHCommand -Command $find_command -Session $ssh_session.SessionId -TimeOut 900

    # Disconnect SSH
    Remove-SSHSession -Session $ssh_session.SessionId

    # Check for results
    if ($output.Output -eq $null) {
         [System.Windows.MessageBox]::Show("No results returned!")
         return
    }
    # Add output to ListBox
    $output.Output | ForEach-Object {
        $path = ($_  |Select-String -AllMatches -Pattern "(.*?),(.*?),(.*)").Matches.Groups[1].Value
        $size = ($_  |Select-String -AllMatches -Pattern "(.*?),(.*?),(.*)").Matches.Groups[2].Value
        $date = ($_  |Select-String -AllMatches -Pattern "(.*?),(.*?),(.*)").Matches.Groups[3].Value

        $Results.AddChild([PSCustomObject]@{Path=$path; Size=$size; Date = $date})
    }
})

$ClearFormButton.Add_Click({
    $HostTextBox.Text = ""
    $UserNameTextBox.Text = ""
    $PasswordTextBox.Text = ""
    $StartDatePicker.SelectedDate = ""
    $EndDatePicker.SelectedDate = ""
    $FilesystemTextBox.Text = ""
    $Results.Items.Clear()
    $global:Form.InvalidateVisual()

})

$ScriptButton.Add_Click({
    foreach ($item in $Results.SelectedItems) {
        ("rm -f " + $item.Path) | Add-Content -Path "$env:USERPROFILE\Documents\remove_files.sh"
    }
    [System.Windows.MessageBox]::Show("Script written to $env:USERPROFILE\Documents\remove_files.sh!")
})
$ExcelButton.Add_Click({
    $excel_file = "$env:USERPROFILE\Documents\FileFinderLog_$(get-date -f MMddyyyyHHmmss).xlsx"
    # Open Excel

    # Create new Excel object
    $excel_object = New-Object -comobject Excel.Application
    $excel_object.visible = $True 

    # Create new Excel workbook
    $excel_workbook = $excel_object.Workbooks.Add()

    # Select the first worksheet in the new workbook
    $excel_worksheet = $excel_workbook.Worksheets.Item(1)

    # Create headers
    $excel_worksheet.Cells.Item(1,1) = "Path"
    $excel_worksheet.Cells.Item(1,2) = "Size"
    $excel_worksheet.Cells.Item(1,3) = "Date"

    # Format headers
    $d = $excel_worksheet.UsedRange

    # Set headers to backrgound pale yellow color, bold font, blue font color
    $d.Interior.ColorIndex = 19
    $d.Font.ColorIndex = 11
    $d.Font.Bold = $True

    # Set first data row
    $row_counter = 2
    Foreach ($item in $Results.Items) {
        $excel_worksheet.Cells.Item($row_counter,1) = $item.Path
        $excel_worksheet.Cells.Item($row_counter,2) = $item.Size
        $excel_worksheet.Cells.Item($row_counter,3) = $item.Date
        $row_counter++
    }
    # Create borders around the cells in the spreadsheet. The below code creates all borders
    $e = $excel_worksheet.Range("A1:C$row_counter")
    $e.Borders.Item(12).Weight = 2
    $e.Borders.Item(12).LineStyle = 1
    $e.Borders.Item(12).ColorIndex = 1

    $e.Borders.Item(11).Weight = 2
    $e.Borders.Item(11).LineStyle = 1
    $e.Borders.Item(11).ColorIndex = 1

    # Set thicker border around outside
    $e.BorderAround(1,4,1)

    # Fit all columns
    $e.Columns("A:F").AutoFit()

    # Save Excel
    $excel_workbook.SaveAs($excel_file) | out-null

    # Quit Excel
    $excel_workbook.Close | out-null
    $excel_object.Quit() | out-null

    [System.Windows.MessageBox]::Show("File written to $excel_file")
})

# Show GUI
$global:Form.ShowDialog() | out-null