<# 
This program takes in csv of progress notes and csv list of resident names. 
It searches through progress notes with the event 'Doctor' and gets the date of entry.
If no progress note is found, it will return 'No Progress Note Found'.
Results are appended to a text file that opens on closing the GUI.
#>

Add-Type -AssemblyName PresentationFramework


# Function processes the data inputs and writes output to file.
Function Get-Doctor()
{
    # arrays
    $lastseen =@()
    $notseen =@()
    $seendetail =@()
    $notinlist =@()

    # Create new fullname column on Res list.
    $res_list = Import-Csv $Global:resi_list_path | Select-Object *, @{Name="FullName";Expression={$_.LastName + "," + $_.FirstName}} 

    # Create new fullname column, filter by event and sort by date in P/Notes
    $doc = Import-Csv $Global:file_path | Select-Object *, @{Name="FullName";Expression={$_.LastName + "," + $_.FirstName}} |
    Where-Object {($_.Event -eq "Doctor")} | Sort-Object -Property Date -Descending 

    # Loop through resident list against progress notes. 
    foreach($name in $res_list){
        if ($name.FullName -in $doc.FullName){
            $lastseen += $name.FullName
            }
        else{
            $nil_notes = $name.FullName + ": No GP progress notes found"
            $notseen += $nil_notes
            }   
        }

    # Get progress note date for matching residents
    foreach($d in $doc){
        if ($d.FullName -in $lastseen){
            $dbcheck = $d.FullName + ": Last GP progress note entry was on the " + $d.Date 
            $seendetail += $dbcheck
            }
        else{
            $lost = $d.FullName + "No GP progress notes found"
            $notinlist += $lost # quality check.
            }   
        }

    # Write results to file
    $seendetail | Out-File -FilePath $Global:save_path -Append 
    $notseen | Out-File -FilePath $Global:save_path -Append
    $lost | Out-File -FilePath $Global:save_path -Append

# Open results on closing GUI
. $Global:save_path
    
} #end


# Function to input progress note file.
Function Get-FileName()
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “CSV files (*.csv)| *.csv”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.Filename

} #end 


# Function to input file of resident names.
Function Get-ResiList()
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog2 = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog2.initialDirectory = $initialDirectory
    $OpenFileDialog2.filter = “CSV files (*.csv)| *.csv”
    $OpenFileDialog2.ShowDialog() | Out-Null
    $OpenFileDialog2.Filename

} #end


# Function gets location to save results.
function Save-File()
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "Text files (*.txt)| *.txt"
    $SaveFileDialog.ShowDialog() |  Out-Null
    $SaveFileDialog.FileName

} #end


# GUI XAML file
$xamlFile = "MainWindow.xaml" 
$inputXAML = Get-Content -Path $xamlFile -Raw

# Tidy up XAMl to work with PowerShell
$inputXAML = $inputXAML -replace 'mc:Ignorable="d"', '' -replace "x:N", "N" -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXAML

# Read XML in
$reader = New-Object System.Xml.XmlNodeReader $XAML

# Load & catch the error if it doesn't load.
try {
    $psform = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Output $_.Exception
    throw
}

# Create controls already named in XAML
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $psform.FindName($_.Name) -ErrorAction Stop
    }
    catch {
        throw
    }
}

# Select_file button action - get filename and update label.
$var_select_file_button.Add_Click({
    $Global:file_path = Get-FileName
    $var_file_label.Content = $Global:file_path  
})

# Button action for resident list file of names and update label.
 $var_resi_list_button.Add_Click({
    $Global:resi_list_path = Get-ResiList
    $var_resi_list_label.Content = $Global:resi_list_path  
})

# Save_file button action and update label.
$var_save_file_button.Add_Click({
    $Global:save_path = Save-File
    $var_save_file_label.Content = $Global:save_path
})

# start button action - Calls main Get-Doctor function
$var_start_button.Add_Click({Get-Doctor})

# show the form
$psform.ShowDialog()
