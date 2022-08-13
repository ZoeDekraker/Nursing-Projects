<# ACFI Auditing Athlete - does the leg work for you. 
I am not responsible for anything you do using this. Use at your own risk. 
Never run a program if you do not know what it is does.
#>


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


#------------------GUI----------------------#

# new instance
$form = New-Object System.Windows.Forms.Form
$form.Text = 'ACFI Auditing Athlete' 
$form.Size = New-Object System.Drawing.Size(390, 300) #  w x h
$form.StartPosition = 'CenterScreen' 

#  OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(100,215) 
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

# cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(175,215)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# audit label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(50,25)
$label.Size = New-Object System.Drawing.Size(180,20)
$label.Text = 'Please select csv file to audit: '
$form.Controls.Add($label)

# audit file button 
$auditButton = New-Object System.Windows.Forms.Button
$auditButton.Location = New-Object System.Drawing.Point(240,25) #
$auditButton.Size = New-Object System.Drawing.Size(75,23)
$auditButton.Text = 'Select'
$auditButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.AcceptButton = $auditButton
$form.Controls.Add($auditButton)

# action when clicked
$auditButton.Add_Click({$Global:OpenFile = Get-FileName})


# export label
$exportLabel = New-Object System.Windows.Forms.Label
$exportLabel.Location = New-Object System.Drawing.Point(50,70)
$exportLabel.Size = New-Object System.Drawing.Size(180,20)
$exportLabel.Text = 'Where to save the results: '
$form.Controls.Add($exportLabel)

# export file button 
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(240,70) 
$exportButton.Size = New-Object System.Drawing.Size(75,23)
$exportButton.Text = 'Select'
$exportButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.AcceptButton = $exportButton
$form.Controls.Add($exportButton)

# action when clicked
$exportButton.Add_Click({$Global:SaveFile = Save-File})


# domain select label
$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Location = New-Object System.Drawing.Point(100,120)
$domainLabel.Size = New-Object System.Drawing.Size(180,20)
$domainLabel.Text = 'Select one ACFI domain to audit '
$form.Controls.Add($domainLabel)

# adl domain
$adlButton = New-Object System.Windows.Forms.Button
$adlButton.Location = New-Object System.Drawing.Point(58,150)
$adlButton.Size = New-Object System.Drawing.Size(50,20)
$adlButton.Text = 'ADL'
$adlButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.Controls.Add($adlButton)
$adlButton.Add_Click({$Global:domainAudit = 'adl'})

# beh domain
$behButton = New-Object System.Windows.Forms.Button
$behButton.Location = New-Object System.Drawing.Point(158,150)
$behButton.Size = New-Object System.Drawing.Size(50,20)
$behButton.Text = 'BEH'
$behButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.Controls.Add($behButton)
$behButton.Add_Click({$Global:domainAudit = 'beh'})

#chc domain
$chcButton = New-Object System.Windows.Forms.Button
$chcButton.Location = New-Object System.Drawing.Point(258,150)
$chcButton.Size = New-Object System.Drawing.Size(50,20)
$chcButton.Text = 'CHC'
$chcButton.DialogResult = [System.Windows.Forms.DialogResult]::None
$form.Controls.Add($chcButton)
$chcButton.Add_Click({$Global:domainAudit = 'chc'})

# open ontop of others
$form.Topmost = $true
$result = $form.ShowDialog()

Write-Output "The file to audit has been selected to open $Global:OpenFile "
Write-Output "The file to save results has been selected $Global:SaveFile"
Write-Output "The chosen domain is $Global:domainAudit"


#------------- Main Script--------------#

# domain keywords 
$adl_words = @('urine', 'wet', 'wee', 'urinary', 'incontinence', 'continence', 'faecal', 'fecal', 
'soil', 'poo', 'smell', 'pad', 'aid', 'fall', 'falls', 'walking', 'pain', 'assisted', 'feed', 'fed',
'food', 'eating') 
$beh_words = @('Dementia', 'Depression', 'Anxiety', 'Anxious', 'Worried', 'confused', 'behaviour', 'yelling', 'yell',
'hit', 'hitting', 'kick', 'kicking', 'loud', 'screaming', 'scream', 'upset', 'agitated', 'angry', 'grabbed', 
'swearing', 'swore', 'push', 'pushing', 'disturbing', 'wandering', 'intruding', 'lost', 'wander', 'abscond')
$chc_words = @('pain', 'pressure', 'sore', 'swelling', 'swollen', 'legs', 'oedema', 'odema', 'ankle', 'falls', 'wound',
'injury', 'diabetes', 'catheter', 'compression', 'stockings', 'PAC', 'reposition', 'bsl', 'bgl', 'oxygen',
'breathlessness', 'SOB')

# update variable with selection
if ($Global:domainAudit -eq 'adl')
{
    $word_list = $adl_words
}
elseif ($Global:domainAudit -eq 'beh')
{
    $word_list = $beh_words
}
elseif ($Global:domainAudit -eq 'chc')
{
    $word_list = $chc_words
}
else
{
    Write-Output "No domain selected"
}


# audit file and write out results
Get-Content -Path $Global:OpenFile -First 1 | Out-File -FilePath $Global:SaveFile
Select-String -Path $Global:OpenFile -Pattern $word_list -Raw | Out-File -FilePath $Global:SaveFile -Append

#opens result spreadsheet with Excel
. $Global:SaveFile

