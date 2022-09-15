Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Select file to audit
Function Get-FileName()
{  
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “CSV files (*.csv)| *.csv”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.Filename

}


# Save output to a file
function Save-File()
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "CSV files (*.csv)| *.csv"
    $SaveFileDialog.ShowDialog() |  Out-Null
    $SaveFileDialog.FileName

}


#<------------------GUI---------------------->#

# new instance
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Mobility Sweeper' 
$form.Size = New-Object System.Drawing.Size(390, 300) #  w x h
$form.StartPosition = 'CenterScreen' 

# audit label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(50,45)
$label.Size = New-Object System.Drawing.Size(180,20)
$label.Text = 'Please select csv file to audit: '
$form.Controls.Add($label)

# audit file button 
$auditButton = New-Object System.Windows.Forms.Button
$auditButton.Location = New-Object System.Drawing.Point(240,45) #
$auditButton.Size = New-Object System.Drawing.Size(75,23)
$auditButton.Text = 'Browse'
$auditButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.AcceptButton = $auditButton
$form.Controls.Add($auditButton)

# action when clicked
$auditButton.Add_Click({$Global:OpenFile = Get-FileName})

# export label
$exportLabel = New-Object System.Windows.Forms.Label
$exportLabel.Location = New-Object System.Drawing.Point(50,110)
$exportLabel.Size = New-Object System.Drawing.Size(180,20)
$exportLabel.Text = 'Select where to save the results to: '
$form.Controls.Add($exportLabel)

# export file button 
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(240,110) 
$exportButton.Size = New-Object System.Drawing.Size(75,23)
$exportButton.Text = 'Browse'
$exportButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.AcceptButton = $exportButton
$form.Controls.Add($exportButton)

# action when clicked
$exportButton.Add_Click({$Global:SaveFile = Save-File})

#  OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(100,195) 
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

# cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(175,195)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# open ontop of other windows
$form.Topmost = $true
$form.ShowDialog()

Write-Output "The file to audit has been selected to open $Global:OpenFile "
Write-Output "The file to save results has been selected $Global:SaveFile"




#< ------------ Logic ----------- >#

# subject keywords
$mob_words = @('fall', 'falls', 'walking', 'pain', 'legs', 'weak', 'unsteady', 'dizzy', 'unstable', 'fell', 'wobble', 'wobbly' ) 

# write header to results csv
Get-Content $Global:OpenFile -First 1 | Out-File $Global:SaveFile

# filter entries against keywords and write results to file.
Get-Content $Global:OpenFile | Select-String -Pattern $mob_words -Raw | Out-File -FilePath $Global:SaveFile -Append

# read in the results csv into out-gridview
Import-Csv $Global:SaveFile | Out-GridView -Title 'Mobility Audit Results'
